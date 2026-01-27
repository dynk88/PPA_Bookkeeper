import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar # CHANGED: Import Calendar instead of DateEntry
from datetime import datetime, date
from num2words import num2words
from config import Config
import os 

class EntryView(tk.Frame):
    def __init__(self, parent, app_controller):
        super().__init__(parent)
        self.controller = app_controller
        self.editing_item_iid = None
        self.setup_ui()

    def setup_ui(self):
        left_panel = tk.Frame(self, bg=Config.COLOR_BG_MAIN, padx=25, pady=25)
        left_panel.place(relx=0, rely=0, relwidth=0.45, relheight=1)
        right_panel = tk.Frame(self, bg=Config.COLOR_BG_WHITE, padx=20, pady=20)
        right_panel.place(relx=0.45, rely=0, relwidth=0.55, relheight=1)

        # --- LEFT: INPUTS ---
        header_frame = tk.Frame(left_panel, bg=Config.COLOR_BG_MAIN)
        header_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(header_frame, text="Transaction Form", font=Config.FONT_HEADER, bg=Config.COLOR_BG_MAIN).pack(side="left")
        
        btn_restart = tk.Button(header_frame, text="⟳ Restart Session", command=self.restart_session,
                                bg=Config.COLOR_WARNING, fg=Config.COLOR_BG_WHITE, font=Config.FONT_SMALL, cursor="hand2")
        btn_restart.pack(side="right")
        
        # Dept
        tk.Label(left_panel, text="Department:", bg=Config.COLOR_BG_MAIN, font=Config.FONT_BODY).pack(anchor="w")
        self.sub_var = tk.StringVar()
        self.sub_combo = ttk.Combobox(left_panel, textvariable=self.sub_var, state="readonly", font=Config.FONT_ENTRY)
        self.sub_combo['values'] = self.controller.system.get_subsidiaries()
        self.sub_combo.pack(fill="x", pady=(5, 15))

        # PPA
        tk.Label(left_panel, text="PPA Number (0/13):", bg=Config.COLOR_BG_MAIN, font=Config.FONT_BODY).pack(anchor="w")
        self.ppa_var = tk.StringVar()
        self.ppa_var.trace_add('write', self.on_ppa_change)
        self.ppa_entry = tk.Entry(left_panel, textvariable=self.ppa_var, font=Config.FONT_MONO_LARGE)
        self.ppa_entry.pack(fill="x", pady=(2, 0))
        self.ppa_entry.bind("<Key>", self.restore_ppa_style)
        self.ppa_entry.bind("<Button-1>", self.restore_ppa_style)
        
        self.lbl_ppa_preview = tk.Label(left_panel, text="", bg=Config.COLOR_BG_MAIN, fg=Config.COLOR_PRIMARY, font=Config.FONT_PREVIEW_LARGE)
        self.lbl_ppa_preview.pack(anchor="w")

        # Amount
        tk.Label(left_panel, text="Amount (₹):", bg=Config.COLOR_BG_MAIN, font=Config.FONT_BODY).pack(anchor="w")
        vcmd = (self.register(self.validate_amount), '%P')
        self.amount_entry = tk.Entry(left_panel, font=Config.FONT_ENTRY, validate="key", validatecommand=vcmd)
        self.amount_entry.pack(fill="x", pady=(5, 2))
        self.lbl_amt_words = tk.Label(left_panel, text="", bg=Config.COLOR_BG_MAIN, fg=Config.COLOR_AMOUNT_PREVIEW, font=Config.FONT_BODY_BOLD, wraplength=350, justify="left")
        self.lbl_amt_words.pack(anchor="w", pady=(0, 15))
        self.amount_entry.bind("<KeyRelease>", self.update_amount_words)

        # Date (FIXED: Using Entry + Button instead of DateEntry)
        tk.Label(left_panel, text="Date:", bg=Config.COLOR_BG_MAIN, font=Config.FONT_BODY).pack(anchor="w")
        
        date_frame = tk.Frame(left_panel, bg=Config.COLOR_BG_MAIN)
        date_frame.pack(fill="x", pady=(5, 25))
        
        self.date_entry = tk.Entry(date_frame, font=Config.FONT_ENTRY)
        self.date_entry.pack(side="left", fill="x", expand=True)
        self.date_entry.insert(0, date.today().strftime("%d-%m-%Y")) # Default Today
        
        btn_cal = tk.Button(date_frame, text="📅", command=self.open_calendar, 
                            bg=Config.COLOR_PRIMARY, fg="white", font=Config.FONT_BODY)
        btn_cal.pack(side="left", padx=5)

        # Buttons
        self.btn_submit = tk.Button(left_panel, text="Submit PPA (To Preview)", command=self.submit_data, bg=Config.COLOR_PRIMARY, fg=Config.COLOR_BG_WHITE, font=Config.FONT_SUBHEADER, height=2, cursor="hand2")
        self.btn_submit.pack(fill="x")
        self.btn_cancel = tk.Button(left_panel, text="Cancel Editing", command=self.cancel_edit, bg=Config.COLOR_TEXT_LIGHT, fg=Config.COLOR_TEXT, font=Config.FONT_SMALL)

        # Nav Buttons
        nav_frame = tk.Frame(left_panel, bg=Config.COLOR_BG_MAIN)
        nav_frame.pack(side="bottom", anchor="w", pady=10, fill="x")
        tk.Button(nav_frame, text="See Records (Dashboard) →", command=lambda: self.controller.show_view("DashboardView"), bg=Config.COLOR_SECONDARY, fg="white", font=Config.FONT_BODY).pack(side="left", padx=5)
        tk.Button(nav_frame, text="History 🔍", command=lambda: self.controller.show_view("HistoryView"), bg=Config.COLOR_SECONDARY, fg="white", font=Config.FONT_BODY).pack(side="left")

        # --- RIGHT: LIST ---
        tk.Label(right_panel, text="Session Preview", font=Config.FONT_SUBHEADER, bg="white").pack(anchor="w")
        
        # Actions
        act_frame = tk.Frame(right_panel, bg="white")
        act_frame.pack(anchor="e", pady=5)
        self.btn_val = tk.Button(act_frame, text="✓ Validate & Save", command=self.validate_data, bg=Config.COLOR_ACCENT, fg="white", font=Config.FONT_SMALL)
        self.btn_val.pack(side="left", padx=5)
        self.btn_exp = tk.Button(act_frame, text="Export noting", command=self.export_word, bg=Config.COLOR_TEXT_LIGHT, fg="white", state="disabled", font=Config.FONT_SMALL)
        self.btn_exp.pack(side="left")

        cols = ("sub", "ppa", "amt", "date")
        self.tree = ttk.Treeview(right_panel, columns=cols, show="headings")
        self.tree.heading("sub", text="Department")
        self.tree.heading("ppa", text="PPA")
        self.tree.heading("amt", text="Amount")
        self.tree.heading("date", text="Date")
        self.tree.column("sub", width=120)
        self.tree.column("ppa", width=120)
        self.tree.column("amt", width=100, anchor="w")
        self.tree.column("date", width=90)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Button-3>", self.show_context)

        self.lbl_total = tk.Label(right_panel, text="Session Total: ₹ 0", font=Config.FONT_SUBHEADER, bg="white", fg=Config.COLOR_BG_HEADER)
        self.lbl_total.pack(side="bottom", anchor="e", pady=10)

        # Context Menu
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Modify", command=self.load_edit)
        self.menu.add_command(label="Delete", command=self.delete_row)

    # --- LOGIC ---
    def open_calendar(self):
        """Opens a stable popup calendar"""
        top = tk.Toplevel(self)
        top.title("Select Date")
        
        # Center the popup
        x = self.winfo_rootx() + 150
        y = self.winfo_rooty() + 150
        top.geometry(f"+{x}+{y}")
        
        cal = Calendar(top, selectmode='day', date_pattern='dd-mm-yyyy')
        cal.pack(padx=10, pady=10)
        
        def set_date():
            selected = cal.get_date()
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, selected)
            top.destroy()
            
        tk.Button(top, text="Confirm", command=set_date, bg=Config.COLOR_PRIMARY, fg="white").pack(pady=5)

    def on_ppa_change(self, *args):
        val = self.ppa_var.get().upper()
        if self.ppa_var.get() != val: self.ppa_var.set(val)
        
        # Color Logic
        color = Config.COLOR_DANGER if not val.isalnum() or len(val) > 13 else Config.COLOR_SUCCESS if len(val)==13 else Config.COLOR_TEXT
        self.lbl_ppa_preview.config(text=" ".join(val), fg=color)

    def update_amount_words(self, e):
        try:
            amt = int(self.amount_entry.get())
            self.lbl_amt_words.config(text=f"{num2words(amt, lang='en_IN').title().replace('-', ' ')} Rupees Only")
        except: self.lbl_amt_words.config(text="")

    def submit_data(self):
        # Gather
        sub, ppa, amt_str = self.sub_var.get(), self.ppa_entry.get(), self.amount_entry.get()
        date_str = self.date_entry.get() # Get string directly
        
        if not sub or not ppa or not amt_str or not ppa.isalnum() or len(ppa)!=13:
            messagebox.showerror("Error", "Check Inputs")
            return
        
        try: 
            amt = int(amt_str)
            # Validate Date format
            datetime.strptime(date_str, "%d-%m-%Y")
        except ValueError: 
            messagebox.showerror("Error", "Invalid Amount or Date (DD-MM-YYYY)")
            return

        row_data = (sub, ppa, self.controller.format_currency(amt), date_str)
        
        if self.editing_item_iid:
            self.tree.item(self.editing_item_iid, values=row_data)
            self.cancel_edit()
        else:
            self.tree.insert("", 0, values=row_data)
            self.sub_combo.config(state="disabled")
            self.amount_entry.delete(0, tk.END)
            self.lbl_amt_words.config(text="")
            # Auto-fill styling
            self.ppa_entry.config(fg=Config.COLOR_DIM_TEXT)
            self.ppa_entry.focus_set()
            self.ppa_entry.select_range(0, tk.END)
        
        self.update_total()

    def update_total(self):
        total = 0
        for item in self.tree.get_children():
            total += self.controller.parse_currency(self.tree.item(item)['values'][2])
        self.lbl_total.config(text=f"Session Total: {self.controller.format_currency(total)}")

    def validate_data(self):
        if not self.tree.get_children(): return
        batch = []
        target_sub = None
        for item in self.tree.get_children():
            v = self.tree.item(item)['values']
            if not target_sub: target_sub = v[0]
            amt = self.controller.parse_currency(v[2])
            dt = datetime.strptime(v[3], "%d-%m-%Y").date()
            batch.append((v[1], dt, amt))
        
        ok, msg = self.controller.system.save_batch(target_sub, batch)
        if ok:
            messagebox.showinfo("Success", "Saved!")
            self.controller.is_session_saved = True
            self.btn_submit.config(state="disabled")
            self.btn_val.config(state="disabled")
            self.btn_exp.config(state="normal", bg=Config.COLOR_SUCCESS)
        else:
            messagebox.showerror("Error", msg)

    def export_word(self):
        if not self.tree.get_children(): return
        first = self.tree.item(self.tree.get_children()[0])['values']
        batch = [(self.tree.item(i)['values'][1], self.tree.item(i)['values'][2]) for i in self.tree.get_children()]
        ok, res = self.controller.system.create_word_advice(first[0], first[3], batch)
        if ok: os.startfile(res)
        else: messagebox.showerror("Error", res)

    def restart_session(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        self.sub_combo.config(state="readonly")
        self.sub_var.set("")
        self.ppa_var.set("")
        self.amount_entry.delete(0, tk.END)
        
        # Reset Date to Today
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, date.today().strftime("%d-%m-%Y"))
        
        self.controller.is_session_saved = False
        self.btn_submit.config(state="normal")
        self.btn_val.config(state="normal")
        self.btn_exp.config(state="disabled", bg=Config.COLOR_TEXT_LIGHT)
        self.restore_ppa_style(None)
        self.update_total()

    def show_context(self, e):
        if not self.controller.is_session_saved:
            i = self.tree.identify_row(e.y)
            if i:
                self.tree.selection_set(i)
                self.menu.post(e.x_root, e.y_root)

    def delete_row(self):
        sel = self.tree.selection()
        if sel:
            self.tree.delete(sel[0])
            self.update_total()
            if not self.tree.get_children():
                self.sub_combo.config(state="readonly")

    def load_edit(self):
        sel = self.tree.selection()
        if sel:
            self.editing_item_iid = sel[0]
            v = self.tree.item(sel[0])['values']
            self.sub_var.set(v[0])
            self.ppa_var.set(v[1])
            self.amount_entry.delete(0, tk.END)
            self.amount_entry.insert(0, str(self.controller.parse_currency(v[2])))
            
            # Set Date in Entry
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, v[3])
            
            self.btn_submit.config(text="Update Entry", bg=Config.COLOR_WARNING)
            self.btn_cancel.pack(pady=5)
            self.update_amount_words(None)
            self.restore_ppa_style(None)

    def cancel_edit(self):
        self.editing_item_iid = None
        self.btn_submit.config(text="Submit PPA", bg=Config.COLOR_PRIMARY)
        self.btn_cancel.pack_forget()
        self.ppa_var.set("")
        self.amount_entry.delete(0, tk.END)

    def restore_ppa_style(self, e):
        self.ppa_entry.config(fg="black")

    def validate_amount(self, P):
        return P == "" or P.isdigit()