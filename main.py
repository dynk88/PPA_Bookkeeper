import tkinter as tk
from tkinter import ttk, messagebox, Menu
from tkcalendar import DateEntry 
import os 
from datetime import datetime
from backend import BookkeepingSystem
from num2words import num2words 

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Department Bookkeeper") 
        self.center_window(1000, 650) 
        
        self.system = BookkeepingSystem()
        self.is_session_saved = False
        self.editing_item_iid = None 

        # STYLING
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'))
        style.configure("Treeview", font=('Segoe UI', 10), rowheight=28)
        
        # --- FOOTER (Developed By) ---
        # We pack this first with side="bottom" so it sticks to the bottom
        footer_frame = tk.Frame(self.root, bg="#f0f0f0", height=20)
        footer_frame.pack(side="bottom", fill="x")
        
        lbl_dev = tk.Label(footer_frame, text="Developed by D.N", 
                           font=("Segoe UI", 8), fg="gray", bg="#f0f0f0")
        lbl_dev.pack(side="right", padx=10, pady=2)
        # -----------------------------

        # MAIN CONTAINER
        self.container = tk.Frame(self.root)
        self.container.pack(fill="both", expand=True)

        self.frame_entry = tk.Frame(self.container)
        self.frame_summary = tk.Frame(self.container)
        self.frame_history = tk.Frame(self.container)

        self.frame_entry.pack(fill="both", expand=True)
        self.setup_entry_view()
        self.setup_summary_view()
        self.setup_history_view()

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def format_indian_currency(self, value):
        try:
            value = int(value)
        except: return value
        s = str(value)
        if len(s) <= 3: return f"₹ {s}"
        last_three = s[-3:]
        remaining = s[:-3]
        formatted_remaining = ""
        for i, digit in enumerate(reversed(remaining)):
            if i > 0 and i % 2 == 0: formatted_remaining = "," + formatted_remaining
            formatted_remaining = digit + formatted_remaining
        return f"₹ {formatted_remaining},{last_three}"

    def parse_currency(self, val_str):
        try:
            return int(str(val_str).replace("₹", "").replace(",", "").strip())
        except:
            return 0

    # ==================== ENTRY VIEW ====================
    def setup_entry_view(self):
        left_panel = tk.Frame(self.frame_entry, bg="#f4f4f4", padx=25, pady=25)
        left_panel.place(relx=0, rely=0, relwidth=0.45, relheight=1)

        right_panel = tk.Frame(self.frame_entry, bg="white", padx=20, pady=20)
        right_panel.place(relx=0.45, rely=0, relwidth=0.55, relheight=1)

        # --- LEFT PANEL ---
        header_frame = tk.Frame(left_panel, bg="#f4f4f4")
        header_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(header_frame, text="Transaction Form", font=("Segoe UI", 16, "bold"), bg="#f4f4f4").pack(side="left")
        
        btn_restart = tk.Button(header_frame, text="⟳ Restart Session", command=self.restart_session,
                                bg="#e0a800", fg="white", font=("Segoe UI", 9, "bold"), cursor="hand2")
        btn_restart.pack(side="right")

        # Department
        self.lbl_sub = tk.Label(left_panel, text="Department:", bg="#f4f4f4", font=("Segoe UI", 10))
        self.lbl_sub.pack(anchor="w")
        self.sub_var = tk.StringVar()
        self.sub_combo = ttk.Combobox(left_panel, textvariable=self.sub_var, state="readonly", font=("Segoe UI", 11))
        self.sub_combo['values'] = self.system.get_subsidiaries()
        self.sub_combo.pack(fill="x", pady=(5, 15))

        # PPA SECTION
        ppa_frame = tk.Frame(left_panel, bg="#f4f4f4")
        ppa_frame.pack(fill="x", pady=(5, 15))
        self.lbl_ppa = tk.Label(ppa_frame, text="PPA Number (0/13):", bg="#f4f4f4", font=("Segoe UI", 10))
        self.lbl_ppa.pack(anchor="w")

        self.ppa_var = tk.StringVar()
        self.ppa_var.trace_add('write', self.on_ppa_change) 
        
        self.ppa_entry = tk.Entry(ppa_frame, textvariable=self.ppa_var, font=("Consolas", 14))
        self.ppa_entry.pack(fill="x", pady=(2, 0))

        self.lbl_ppa_preview = tk.Label(ppa_frame, text="", bg="#f4f4f4", fg="#0078D7", 
                                        font=("Segoe UI", 16, "bold"))
        self.lbl_ppa_preview.pack(anchor="w")

        # Amount
        tk.Label(left_panel, text="Amount (₹):", bg="#f4f4f4", font=("Segoe UI", 10)).pack(anchor="w")
        vcmd = (self.root.register(self.validate_amount), '%P')
        self.amount_entry = tk.Entry(left_panel, font=("Segoe UI", 11), validate="key", validatecommand=vcmd)
        self.amount_entry.pack(fill="x", pady=(5, 2))

        self.lbl_amt_words = tk.Label(left_panel, text="", bg="#f4f4f4", fg="#666", 
                                      font=("Segoe UI", 9, "italic"), wraplength=350, justify="left")
        self.lbl_amt_words.pack(anchor="w", pady=(0, 15))
        self.amount_entry.bind("<KeyRelease>", self.update_amount_words)

        # Date
        tk.Label(left_panel, text="Date:", bg="#f4f4f4", font=("Segoe UI", 10)).pack(anchor="w")
        self.date_entry = DateEntry(left_panel, width=12, background='darkblue',
                                    foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy', font=("Segoe UI", 11))
        self.date_entry.pack(fill="x", pady=(5, 25))

        # Submit
        self.btn_submit = tk.Button(left_panel, text="Submit PPA (To Preview)", command=self.submit_data, 
                               bg="#0078D7", fg="white", font=("Segoe UI", 11, "bold"), height=2, cursor="hand2")
        self.btn_submit.pack(fill="x")

        # Cancel Edit
        self.btn_cancel_edit = tk.Button(left_panel, text="Cancel Editing", command=self.cancel_edit,
                                         bg="gray", fg="black", font=("Segoe UI", 9))

        # Nav
        nav_frame = tk.Frame(left_panel, bg="#f4f4f4")
        nav_frame.pack(side="bottom", anchor="w", pady=10, fill="x")
        
        btn_records = tk.Button(nav_frame, text="See Records (Dashboard) →", command=self.show_summary_screen,
                                bg="#555555", fg="white", font=("Segoe UI", 10), cursor="hand2")
        btn_records.pack(side="left", padx=(0, 5))

        btn_hist = tk.Button(nav_frame, text="Search & History 🔍", command=self.show_history_screen,
                                bg="#555555", fg="white", font=("Segoe UI", 10), cursor="hand2")
        btn_hist.pack(side="left")

        # --- RIGHT PANEL ---
        right_header = tk.Frame(right_panel, bg="white")
        right_header.pack(fill="x", pady=(0, 10))
        
        tk.Label(right_header, text="Session Preview", font=("Segoe UI", 12, "bold"), bg="white").pack(side="left")
        tk.Label(right_panel, text="Tip: Right-click row to Edit or Delete", 
                 font=("Segoe UI", 9, "italic"), fg="gray", bg="white").pack(anchor="w")

        action_frame = tk.Frame(right_header, bg="white")
        action_frame.pack(side="right")

        self.btn_validate = tk.Button(action_frame, text="✓ Validate & Save", command=self.validate_data,
                                bg="#d63384", fg="white", font=("Segoe UI", 9, "bold"), cursor="hand2")
        self.btn_validate.pack(side="left", padx=5)

        self.btn_export = tk.Button(action_frame, text="Export noting", command=self.export_word,
                               bg="gray", fg="white", font=("Segoe UI", 9, "bold"), cursor="hand2", state="disabled")
        self.btn_export.pack(side="left")

        columns = ("sub", "ppa", "amt", "date")
        self.tree_session = ttk.Treeview(right_panel, columns=columns, show="headings")
        self.tree_session.heading("sub", text="Department")
        self.tree_session.heading("ppa", text="PPA")
        self.tree_session.heading("amt", text="Amount")
        self.tree_session.heading("date", text="Date")
        
        self.tree_session.column("sub", width=120)
        self.tree_session.column("ppa", width=120)
        self.tree_session.column("amt", width=100, anchor="w") 
        self.tree_session.column("date", width=90)

        self.tree_session.pack(fill="both", expand=True)
        self.tree_session.bind("<Button-3>", self.show_context_menu) 
        
        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Modify Entry", command=self.on_modify_context)
        self.context_menu.add_command(label="Delete Entry", command=self.delete_selected_row)

    # ==================== HISTORY VIEW ====================
    def setup_history_view(self):
        top_frame = tk.Frame(self.frame_history, bg="#2C3E50", height=60)
        top_frame.pack(fill="x")
        tk.Label(top_frame, text="Transaction History & Search", bg="#2C3E50", fg="white", font=("Segoe UI", 16, "bold")).pack(side="left", padx=20, pady=15)
        tk.Button(top_frame, text="← Back to Entry", command=self.show_entry_screen, bg="white", font=("Segoe UI", 10)).pack(side="right", padx=20, pady=15)

        filter_frame = tk.Frame(self.frame_history, bg="#ECF0F1", padx=20, pady=10)
        filter_frame.pack(fill="x")
        tk.Label(filter_frame, text="Filter Department:", bg="#ECF0F1").pack(side="left", padx=5) 
        self.hist_sub_var = tk.StringVar()
        self.hist_sub_combo = ttk.Combobox(filter_frame, textvariable=self.hist_sub_var, state="readonly", width=30)
        subs = ["All Departments"] + self.system.get_subsidiaries()
        self.hist_sub_combo['values'] = subs
        self.hist_sub_combo.current(0)
        self.hist_sub_combo.pack(side="left", padx=5)
        tk.Label(filter_frame, text="Search PPA:", bg="#ECF0F1").pack(side="left", padx=(20, 5))
        self.hist_ppa_var = tk.StringVar()
        tk.Entry(filter_frame, textvariable=self.hist_ppa_var, width=20).pack(side="left", padx=5)
        tk.Button(filter_frame, text="Search / Refresh", command=self.run_history_search,
                  bg="#0078D7", fg="white").pack(side="left", padx=20)

        content_frame = tk.Frame(self.frame_history, padx=20, pady=20)
        content_frame.pack(fill="both", expand=True)
        columns = ("sub", "ppa", "date", "amt")
        self.tree_history = ttk.Treeview(content_frame, columns=columns, show="headings")
        self.tree_history.heading("sub", text="Department")
        self.tree_history.heading("ppa", text="PPA Number")
        self.tree_history.heading("date", text="Date")
        self.tree_history.heading("amt", text="Amount")
        self.tree_history.column("sub", width=200)
        self.tree_history.column("ppa", width=150)
        self.tree_history.column("date", width=100)
        self.tree_history.column("amt", width=120, anchor="w")
        scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=self.tree_history.yview)
        self.tree_history.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree_history.pack(fill="both", expand=True)

    def run_history_search(self):
        for item in self.tree_history.get_children(): self.tree_history.delete(item)
        sub = self.hist_sub_var.get()
        ppa = self.hist_ppa_var.get().strip()
        data = self.system.search_transactions(subsidiary=sub, ppa_text=ppa)
        for row in data:
            s_sub, s_ppa, s_date, s_amt = row
            fmt_date = s_date.strftime("%d-%m-%Y") if isinstance(s_date, datetime) else str(s_date)
            fmt_amt = self.format_indian_currency(s_amt)
            self.tree_history.insert("", "end", values=(s_sub, s_ppa, fmt_date, fmt_amt))

    # ==================== SUMMARY VIEW ====================
    def setup_summary_view(self):
        top_frame = tk.Frame(self.frame_summary, bg="#59A4E2", height=60)
        top_frame.pack(fill="x")
        tk.Label(top_frame, text="Financial Dashboard", bg="#59A4E2", fg="white", font=("Segoe UI", 16, "bold")).pack(side="left", padx=20, pady=15)
        
        tk.Button(top_frame, text="← Back to Entry", command=self.show_entry_screen, bg="white", font=("Segoe UI", 10)).pack(side="right", padx=20, pady=15)

        self.btn_pdf = tk.Button(top_frame, text="Export Report to PDF", command=self.export_pdf_report,
                                 bg="#E78F3C", fg="white", font=("Segoe UI", 10, "bold"), cursor="hand2")
        self.btn_pdf.pack(side="right", padx=10, pady=15)

        content_frame = tk.Frame(self.frame_summary, padx=30, pady=30)
        content_frame.pack(fill="both", expand=True)
        columns = ("sub", "limit", "spent", "bal")
        self.tree_summary = ttk.Treeview(content_frame, columns=columns, show="headings")
        self.tree_summary.heading("sub", text="Department")
        self.tree_summary.heading("limit", text="Approved Limit")
        self.tree_summary.heading("spent", text="Total Disbursed")
        self.tree_summary.heading("bal", text="Remaining Balance")
        self.tree_summary.column("limit", anchor="w")
        self.tree_summary.column("spent", anchor="w")
        self.tree_summary.column("bal", anchor="w")
        self.tree_summary.pack(fill="both", expand=True)

    def export_pdf_report(self):
        summary_data = self.system.get_summary_report()
        if not summary_data:
             messagebox.showinfo("Export", "No data available to export.")
             return
        success, result = self.system.create_dashboard_pdf(summary_data)
        if success:
            try: os.startfile(result)
            except Exception as e: messagebox.showinfo("Saved", f"PDF saved at:\n{result}")
        else:
            messagebox.showerror("Error", result)

    # ==================== COMMON LOGIC ====================
    def update_amount_words(self, event):
        val = self.amount_entry.get()
        if not val:
            self.lbl_amt_words.config(text="")
            return
        try:
            amt = int(val)
            words = num2words(amt, lang='en_IN')
            clean_text = f"{words} Rupees Only".title().replace("-", " ")
            self.lbl_amt_words.config(text=clean_text)
        except:
            self.lbl_amt_words.config(text="")

    def show_context_menu(self, event):
        if self.is_session_saved: return 
        item = self.tree_session.identify_row(event.y)
        if item:
            self.tree_session.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_row(self):
        selected = self.tree_session.selection()
        if not selected: return
        self.tree_session.delete(selected[0])
        if not self.tree_session.get_children():
            self.sub_combo.config(state="readonly")
            self.lbl_sub.config(text="Department:", fg="black") 
            self.cancel_edit()

    def on_modify_context(self):
        self.load_item_for_editing()

    def load_item_for_editing(self):
        selected = self.tree_session.selection()
        if not selected: return
        iid = selected[0]
        values = self.tree_session.item(iid)['values']
        self.sub_var.set(values[0])
        self.ppa_var.set(values[1])
        raw_amt = self.parse_currency(values[2])
        self.amount_entry.delete(0, tk.END)
        self.amount_entry.insert(0, str(raw_amt))
        self.update_amount_words(None)
        self.date_entry.set_date(values[3])
        self.editing_item_iid = iid
        self.btn_submit.config(text="Update Entry", bg="#E67E22") 
        self.btn_cancel_edit.pack(pady=5) 

    def cancel_edit(self):
        self.editing_item_iid = None
        self.btn_submit.config(text="Submit PPA (To Preview)", bg="#0078D7")
        self.btn_cancel_edit.pack_forget()
        self.ppa_var.set("")
        self.amount_entry.delete(0, tk.END)
        self.lbl_amt_words.config(text="")

    def submit_data(self):
        sub = self.sub_var.get()
        ppa = self.ppa_entry.get() 
        amt_str = self.amount_entry.get()
        date_obj = self.date_entry.get_date() 

        if not sub:
            messagebox.showwarning("Error", "Select a Department")
            return
        if not ppa:
            messagebox.showwarning("Error", "PPA Number is empty.")
            return
        if not ppa.isalnum():
            messagebox.showwarning("Error", "PPA Number must contain only Letters and Numbers.")
            return
        if len(ppa) != 13:
            messagebox.showwarning("Error", "PPA must be exactly 13 characters.")
            return
        if not amt_str:
            messagebox.showwarning("Error", "Enter Amount")
            return
        try:
            amt_int = int(amt_str)
        except ValueError:
            messagebox.showerror("Error", "Amount must be a whole number.")
            return

        display_amt = self.format_indian_currency(amt_int)
        display_date = date_obj.strftime("%d-%m-%Y")
        
        if self.editing_item_iid:
            self.tree_session.item(self.editing_item_iid, values=(sub, ppa, display_amt, display_date))
            self.cancel_edit()
            messagebox.showinfo("Updated", "Entry updated successfully.")
        else:
            self.tree_session.insert("", 0, values=(sub, ppa, display_amt, display_date))
            self.sub_combo.config(state="disabled")
            self.lbl_sub.config(text="Department (Locked for Session):", fg="gray")
            self.ppa_var.set("") 
            self.amount_entry.delete(0, tk.END)
            self.lbl_amt_words.config(text="")
            # Preview label clears automatically via variable trace

    def restart_session(self):
        for item in self.tree_session.get_children():
            self.tree_session.delete(item)
        self.sub_combo.config(state="readonly")
        self.lbl_sub.config(text="Department:", fg="black") 
        self.cancel_edit() 
        self.sub_var.set("")
        self.ppa_var.set("")
        self.amount_entry.delete(0, tk.END)
        self.lbl_amt_words.config(text="")
        self.is_session_saved = False
        self.btn_submit.config(state="normal", bg="#0078D7")
        self.btn_validate.config(state="normal", bg="#d63384")
        self.btn_export.config(state="disabled", bg="gray")

    def validate_data(self):
        children = self.tree_session.get_children()
        if not children:
            messagebox.showwarning("Validation", "No transactions to validate.")
            return
        batch_list = []
        target_sub = None
        for child in children:
            row = self.tree_session.item(child)['values']
            sub = row[0]
            if target_sub is None: target_sub = sub
            ppa = row[1]
            amt = self.parse_currency(row[2])
            dt_obj = datetime.strptime(row[3], "%d-%m-%Y").date()
            batch_list.append((ppa, dt_obj, amt))
        success, msg = self.system.save_batch(target_sub, batch_list)
        if success:
            messagebox.showinfo("Success", "Data validated and saved to Excel successfully.")
            self.is_session_saved = True
            self.btn_submit.config(state="disabled", bg="gray")
            self.btn_validate.config(state="disabled", bg="gray")
            if self.editing_item_iid: self.cancel_edit() 
            self.btn_export.config(state="normal", bg="#28a745")
        else:
            messagebox.showerror("Validation Failed", msg)

    def export_word(self):
        children = self.tree_session.get_children()
        if not children: return
        first_row = self.tree_session.item(children[0])['values']
        target_sub = first_row[0]
        ref_date = first_row[3]
        batch_transactions = []
        for child in children:
            row = self.tree_session.item(child)['values']
            batch_transactions.append((row[1], row[2]))
        success, result = self.system.create_word_advice(target_sub, ref_date, batch_transactions)
        if success:
            try: os.startfile(result)
            except Exception as e: messagebox.showinfo("Saved", f"File saved at:\n{result}")
        else:
            messagebox.showerror("Error", result)

    def on_ppa_change(self, *args):
        val = self.ppa_var.get()
        upper_val = val.upper()
        if val != upper_val:
            self.ppa_var.set(upper_val)
            return
        
        count = len(upper_val)
        is_alnum = upper_val.isalnum()
        
        if count > 0 and not is_alnum:
            self.lbl_ppa.config(text="Error: Letters & Numbers only!", fg="red")
            self.lbl_ppa_preview.config(fg="red")
        elif count > 13:
            self.lbl_ppa.config(text=f"PPA Number (Too Long! {count}/13):", fg="red")
            self.lbl_ppa_preview.config(fg="red")
        elif count == 13:
            self.lbl_ppa.config(text=f"PPA Number (Perfect: 13):", fg="green")
            self.lbl_ppa_preview.config(fg="green")
        else:
            self.lbl_ppa.config(text=f"PPA Number ({count}/13 chars):", fg="black")
            self.lbl_ppa_preview.config(fg="#0078D7")

        spaced_text = " ".join(list(upper_val))
        self.lbl_ppa_preview.config(text=spaced_text)

    def validate_amount(self, P):
        if P == "" or P.isdigit(): return True
        return False

    def show_summary_screen(self):
        self.frame_entry.pack_forget()
        self.frame_history.pack_forget()
        self.frame_summary.pack(fill="both", expand=True)
        for item in self.tree_summary.get_children(): self.tree_summary.delete(item)
        data = self.system.get_summary_report()
        for row in data:
            name, limit, spent, rem = row
            self.tree_summary.insert("", "end", values=(
                name, self.format_indian_currency(limit),
                self.format_indian_currency(spent), self.format_indian_currency(rem)
            ))

    def show_entry_screen(self):
        self.frame_summary.pack_forget()
        self.frame_history.pack_forget()
        self.frame_entry.pack(fill="both", expand=True)

    def show_history_screen(self):
        self.frame_entry.pack_forget()
        self.frame_summary.pack_forget()
        self.frame_history.pack(fill="both", expand=True)
        self.run_history_search() 

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()