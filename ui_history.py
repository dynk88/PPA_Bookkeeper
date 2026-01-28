import tkinter as tk
from tkinter import ttk
from datetime import datetime
from config import Config

class HistoryView(tk.Frame):
    def __init__(self, parent, app_controller):
        super().__init__(parent)
        self.controller = app_controller
        
        top = tk.Frame(self, bg=Config.COLOR_BG_HEADER, height=60)
        top.pack(fill="x")
        tk.Label(top, text="History & Search", bg=Config.COLOR_BG_HEADER, fg="white", font=Config.FONT_HEADER).pack(side="left", padx=20, pady=15)
        tk.Button(top, text="← Back", command=lambda: self.controller.show_view("EntryView"), bg="white").pack(side="right", padx=20)

        # Filters
        f_frame = tk.Frame(self, bg=Config.COLOR_BG_MAIN, padx=20, pady=10)
        f_frame.pack(fill="x")
        
        # Dept Filter
        tk.Label(f_frame, text="Dept:", bg=Config.COLOR_BG_MAIN).pack(side="left")
        self.dept_var = tk.StringVar()
        self.combo = ttk.Combobox(f_frame, textvariable=self.dept_var, state="readonly", width=25)
        self.combo.pack(side="left", padx=5)
        
        # Quarter Filter
        tk.Label(f_frame, text="Quarter:", bg=Config.COLOR_BG_MAIN).pack(side="left", padx=(10, 0))
        self.q_var = tk.StringVar()
        self.q_combo = ttk.Combobox(f_frame, textvariable=self.q_var, state="readonly", width=10)
        self.q_combo['values'] = ["All", "Q1", "Q2", "Q3", "Q4"]
        self.q_combo.current(0)
        self.q_combo.pack(side="left", padx=5)

        # PPA Filter
        tk.Label(f_frame, text="PPA:", bg=Config.COLOR_BG_MAIN).pack(side="left", padx=(10, 5))
        self.ppa_var = tk.StringVar()
        tk.Entry(f_frame, textvariable=self.ppa_var, width=15).pack(side="left")
        
        tk.Button(f_frame, text="Search", command=self.run_search, bg=Config.COLOR_PRIMARY, fg="white").pack(side="left", padx=20)

        # Table
        cols = ("sub", "ppa", "date", "amt")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        self.tree.heading("sub", text="Department")
        self.tree.heading("ppa", text="Reference / Type") # UPDATED HEADER
        self.tree.heading("date", text="Date")
        self.tree.heading("amt", text="Amount")
        self.tree.column("sub", width=200)
        self.tree.column("ppa", width=150)
        self.tree.column("date", width=100)
        self.tree.column("amt", width=120, anchor="w")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, padx=20, pady=20)

    def refresh(self):
        # Update dropdown
        subs = ["All Departments"] + self.controller.system.get_subsidiaries()
        self.combo['values'] = subs
        self.combo.current(0)
        self.run_search()

    def run_search(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        d_val = self.dept_var.get()
        q_val = self.q_var.get()
        p_val = self.ppa_var.get().strip()
        
        data = self.controller.system.search_transactions(subsidiary=d_val, ppa_text=p_val, quarter=q_val)
        
        for row in data:
            dt = row[2].strftime("%d-%m-%Y") if isinstance(row[2], datetime) else str(row[2])
            amt = self.controller.format_currency(row[3])
            self.tree.insert("", "end", values=(row[0], row[1], dt, amt))