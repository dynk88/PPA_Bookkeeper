class Config:
    # --- APP SETTINGS ---
    APP_TITLE = "Department Bookkeeper"
    DEV_NAME = "Developed by Your Name"
    WINDOW_SIZE = "1000x650"
    
    # --- FILE PATHS ---
    DB_FILENAME = "data.xlsx"
    SHEET_LIMITS = "Limits"
    SHEET_TXN = "Transactions"

    # --- COLORS (THEME) ---
    # Change these codes to switch themes (e.g., Blue to Green)
    COLOR_PRIMARY = "#0078D7"       # Main Blue
    COLOR_SECONDARY = "#555555"     # Dark Gray (Secondary buttons)
    COLOR_ACCENT = "#d63384"        # Pink/Purple (Validate button)
    COLOR_SUCCESS = "#28a745"       # Green (Export/Success)
    COLOR_WARNING = "#e0a800"       # Orange/Yellow (Restart/Update)
    COLOR_DANGER = "#E74C3C"        # Red (PDF Export/Errors)
    
    COLOR_BG_MAIN = "#f4f4f4"       # Light Gray Background
    COLOR_BG_WHITE = "white"        # White Background
    COLOR_BG_HEADER = "#2C3E50"     # Dark Blue/Grey (History Header)
    COLOR_TEXT = "black"
    COLOR_TEXT_LIGHT = "gray"
    
    # --- FONTS ---
    FONT_FAMILY = "Segoe UI"
    FONT_MONO_FAMILY = "Consolas"
    
    # Pre-defined font tuples for Tkinter
    FONT_HEADER = (FONT_FAMILY, 16, "bold")
    FONT_SUBHEADER = (FONT_FAMILY, 12, "bold")
    FONT_BODY = (FONT_FAMILY, 10)
    FONT_BODY_BOLD = (FONT_FAMILY, 10, "bold")
    FONT_SMALL = (FONT_FAMILY, 9)
    FONT_SMALL_ITALIC = (FONT_FAMILY, 9, "italic")
    FONT_FOOTER = (FONT_FAMILY, 8)
    
    # Input specific fonts
    FONT_ENTRY = (FONT_FAMILY, 11)
    FONT_MONO_LARGE = (FONT_MONO_FAMILY, 14) # For PPA Input
    FONT_PREVIEW_LARGE = (FONT_FAMILY, 16, "bold") # For PPA Preview