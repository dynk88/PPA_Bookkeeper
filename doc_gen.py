import os
from docx import Document
from docx.shared import Pt

def _fmt_rupee(value):
    """Formats integer to Indian Currency String"""
    try:
        value = int(value)
    except:
        return str(value)
    
    s = str(value)
    if len(s) <= 3:
        return f"₹ {s}"
        
    last_three = s[-3:]
    remaining = s[:-3]
    formatted_remaining = ""
    for i, digit in enumerate(reversed(remaining)):
        if i > 0 and i % 2 == 0:
            formatted_remaining = "," + formatted_remaining
        formatted_remaining = digit + formatted_remaining
        
    return f"₹ {formatted_remaining},{last_three}"

def _parse_rupee(value_str):
    """Converts '₹ 1,000' back to integer 1000 for math"""
    try:
        clean = str(value_str).replace("₹", "").replace(",", "").strip()
        return int(clean)
    except:
        return 0

def generate_payment_advice(subsidiary, date_str, transaction_list):
    """
    subsidiary: str
    date_str: str (Used for the content header)
    transaction_list: list of tuples -> [(ppa_number, amount_string), ...]
    """
    # 1. Setup Folder
    valid_chars = "-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    safe_sub_name = "".join(c for c in subsidiary if c in valid_chars).strip()
    
    if not os.path.exists(safe_sub_name):
        try: os.makedirs(safe_sub_name)
        except OSError as e: return False, f"Error creating folder: {e}"
    
    # 2. Determine Filename (noting1, noting2, etc.)
    i = 1
    while True:
        filename = f"{safe_sub_name}/noting{i}.docx"
        if not os.path.exists(filename):
            break
        i += 1
    
    # 3. Init Document
    try:
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
    except:
        return False, "Error initializing Word."

    # 4. Header
    p_head = doc.add_paragraph()
    run_reg = p_head.add_run("Reg :-  ")
    run_reg.bold = True
    run_reg.underline = True
    
    run_title = p_head.add_run("Signature of Print Payment Advice (PPA) under SDRF Parent Account")
    run_title.bold = True
    
    doc.add_paragraph() 

    # 5. Body
    p_body = doc.add_paragraph()
    p_body.add_run("Print payment Advice (PPA) in respect of ")
    p_body.add_run(subsidiary).underline = True 
    p_body.add_run(" under SDRF parent account has been placed below for JD (DM) signature please.")
    
    doc.add_paragraph()

    # 6. Table
    num_rows = 1 + len(transaction_list) + 1
    table = doc.add_table(rows=num_rows, cols=4)
    table.style = 'Table Grid'
    
    # --- Headers ---
    headers = ["Sl. No.", "In favour of", "Print Payment Advice No.", "Amount"]
    for i, text in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = text
        for p in cell.paragraphs:
            for run in p.runs: run.font.bold = True
            
    # --- Data Rows ---
    grand_total = 0
    
    for idx, (ppa, amount_str) in enumerate(transaction_list):
        row_idx = idx + 1 
        row_cells = table.rows[row_idx].cells
        
        # Calculate total
        val = _parse_rupee(amount_str)
        grand_total += val
        
        # Fill cells
        row_cells[0].text = str(idx + 1) + "."
        row_cells[1].text = subsidiary
        row_cells[2].text = str(ppa)
        row_cells[3].text = amount_str

    # --- Total Row ---
    last_row = table.rows[-1]
    
    cell_label = last_row.cells[2]
    cell_label.text = "G/TOTAL"
    for p in cell_label.paragraphs:
        for r in p.runs: r.font.bold = True
        
    cell_total = last_row.cells[3]
    cell_total.text = _fmt_rupee(grand_total)
    for p in cell_total.paragraphs:
        for r in p.runs: r.font.bold = True

    # 7. Save
    try:
        doc.save(filename)
        return True, os.path.abspath(filename)
    except PermissionError:
        return False, "Error: Word file is open. Close it and try again."
    except Exception as e:
        return False, str(e)