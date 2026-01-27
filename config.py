class Config:
    # --- APP SETTINGS ---
    APP_TITLE = "Department Bookkeeper"
    DEV_NAME = "Developed by D. N"
    WINDOW_SIZE = "1000x650"
    
    # --- FILE PATHS ---
    DB_FILENAME = "data.xlsx"
    SHEET_LIMITS = "Limits"
    SHEET_TXN = "Transactions"

    # --- COLORS (THEME) ---
    COLOR_PRIMARY = "#0078D7"       # Main Blue
    COLOR_SECONDARY = "#555555"     # Dark Gray
    COLOR_ACCENT = "#d63384"        # Pink/Purple
    COLOR_SUCCESS = "#28a745"       # Green
    COLOR_WARNING = "#e0a800"       # Orange/Yellow
    COLOR_DANGER = "#E74C3C"        # Red
    
    # NEW COLORS
    COLOR_DIM_TEXT = "#A0A0A0"      # Lighter grey for auto-filled PPA
    COLOR_AMOUNT_PREVIEW = "#008080" # Teal/SeaGreen for Amount Words
    
    COLOR_BG_MAIN = "#f4f4f4"       
    COLOR_BG_WHITE = "white"        
    COLOR_BG_HEADER = "#2C3E50"     
    COLOR_TEXT = "black"
    COLOR_TEXT_LIGHT = "gray"
    
    # --- FONTS ---
    FONT_FAMILY = "Segoe UI"
    FONT_MONO_FAMILY = "Consolas"
    
    FONT_HEADER = (FONT_FAMILY, 16, "bold")
    FONT_SUBHEADER = (FONT_FAMILY, 12, "bold")
    FONT_BODY = (FONT_FAMILY, 10)
    FONT_BODY_BOLD = (FONT_FAMILY, 10, "bold")
    FONT_SMALL = (FONT_FAMILY, 9)
    FONT_SMALL_ITALIC = (FONT_FAMILY, 9, "italic")
    FONT_FOOTER = (FONT_FAMILY, 8)
    
    FONT_ENTRY = (FONT_FAMILY, 11)
    FONT_MONO_LARGE = (FONT_MONO_FAMILY, 14) 
    FONT_PREVIEW_LARGE = (FONT_FAMILY, 16, "bold")