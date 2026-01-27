import tkinter as tk
from tkinter import ttk
from config import Config
from backend import BookkeepingSystem
from ui_entry import EntryView
from ui_dashboard import DashboardView
from ui_history import HistoryView

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(Config.APP_TITLE)
        self.center_window(1100, 650) # Wider for dashboard
        
        self.system = BookkeepingSystem()
        self.is_session_saved = False # Shared state

        # Styling
        style = ttk.Style()
        style.configure("Treeview.Heading", font=Config.FONT_BODY_BOLD)
        style.configure("Treeview", font=Config.FONT_BODY, rowheight=28)

        # Container
        self.container = tk.Frame(self.root)
        self.container.pack(fill="both", expand=True, pady=(0, 20)) # Space for footer

        # Footer
        footer = tk.Frame(self.root, bg="#f0f0f0", height=20)
        footer.place(relx=0, rely=1, anchor="sw", relwidth=1, y=-1)
        tk.Label(footer, text=Config.DEV_NAME, font=Config.FONT_FOOTER, fg="gray", bg="#f0f0f0").pack(side="right", padx=10)

        # Initialize Views
        self.views = {}
        
        # We pass 'self' (App) as the controller to all views
        self.views["EntryView"] = EntryView(self.container, self)
        self.views["DashboardView"] = DashboardView(self.container, self)
        self.views["HistoryView"] = HistoryView(self.container, self)

        self.show_view("EntryView")

    def show_view(self, view_name):
        # Hide all
        for view in self.views.values():
            view.pack_forget()
        
        # Show selected
        view = self.views[view_name]
        view.pack(fill="both", expand=True)
        
        # Trigger refresh if applicable
        if hasattr(view, "refresh"):
            view.refresh()

    def center_window(self, width, height):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw // 2) - (width // 2)
        y = (sh // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    # Shared Helpers
    def format_currency(self, val):
        try: val = int(val)
        except: return val
        s = str(val)
        if len(s) <= 3: return f"₹ {s}"
        return f"₹ {self._comma_fmt(s[:-3])},{s[-3:]}"
        
    def _comma_fmt(self, s):
        res = ""
        for i, d in enumerate(reversed(s)):
            if i > 0 and i % 2 == 0: res = "," + res
            res = d + res
        return res

    def parse_currency(self, val_str):
        try: return int(str(val_str).replace("₹", "").replace(",", "").strip())
        except: return 0

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()