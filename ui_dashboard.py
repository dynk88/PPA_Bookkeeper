import tkinter as tk
from tkinter import ttk, messagebox
import os
from config import Config

class DashboardView(tk.Frame):
    def __init__(self, parent, app_controller):
        super().__init__(parent)
        self.controller = app_controller
        
        # Header
        top = tk.Frame(self, bg=Config.COLOR_PRIMARY, height=60)
        top.pack(fill="x")
        tk.Label(top, text="Financial Dashboard (Quarterly)", bg=Config.COLOR_PRIMARY, fg="white", font=Config.FONT_HEADER).pack(side="left", padx=20, pady=15)
        tk.Button(top, text="← Back to Entry", command=lambda: self.controller.show_view("EntryView"), bg="white").pack(side="right", padx=10)
        tk.Button(top, text="Export PDF", command=self.export_pdf, bg=Config.COLOR_DANGER, fg="white", font=Config.FONT_BODY_BOLD).pack(side="right", padx=10)

        # Table
        # COLS: Sub, Limit, Q1, Q2, Q3, Q4, Total, Bal
        cols = ("sub", "limit", "q1", "q2", "q3", "q4", "spent", "bal")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        
        self.tree.heading("sub", text="Department")
        self.tree.heading("limit", text="Limit")
        self.tree.heading("q1", text="Q1 (Apr-Jun)")
        self.tree.heading("q2", text="Q2 (Jul-Sep)")
        self.tree.heading("q3", text="Q3 (Oct-Dec)")
        self.tree.heading("q4", text="Q4 (Jan-Mar)")
        self.tree.heading("spent", text="Total")
        self.tree.heading("bal", text="Balance")
        
        # Widths
        self.tree.column("sub", width=180)
        for c in cols[1:]:
            self.tree.column(c, width=90, anchor="w")
            
        self.tree.pack(fill="both", expand=True, padx=20, pady=20)

    def refresh(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        data = self.controller.system.get_summary_report()
        for row in data:
            # row is tuple: (Name, Limit, q1, q2, q3, q4, tot, bal)
            fmt_row = [row[0]]
            for val in row[1:]:
                fmt_row.append(self.controller.format_currency(val))
            self.tree.insert("", "end", values=fmt_row)

    def export_pdf(self):
        # CHANGED: Now fetching detailed data specifically for the PDF
        data = self.controller.system.get_detailed_report_data()
        if not data: 
            messagebox.showinfo("Info", "No data available to export.")
            return
            
        ok, res = self.controller.system.create_dashboard_pdf(data)
        if ok: os.startfile(res)
        else: messagebox.showerror("Error", res)