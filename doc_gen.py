import os
from datetime import datetime
from docx import Document
from docx.shared import Pt

# REPORTLAB IMPORTS
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

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
    try:
        clean = str(value_str).replace("₹", "").replace(",", "").strip()
        return int(clean)
    except:
        return 0

# --- WORD GENERATION ---
def generate_payment_advice(subsidiary, date_str, transaction_list):
    valid_chars = "-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    safe_sub_name = "".join(c for c in subsidiary if c in valid_chars).strip()
    
    if not os.path.exists(safe_sub_name):
        try: os.makedirs(safe_sub_name)
        except OSError as e: return False, f"Error creating folder: {e}"
    
    i = 1
    while True:
        filename = f"{safe_sub_name}/noting{i}.docx"
        if not os.path.exists(filename): break
        i += 1
    
    try:
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
        
        p_head = doc.add_paragraph()
        run_reg = p_head.add_run("Reg :-  ")
        run_reg.bold = True
        run_reg.underline = True
        run_title = p_head.add_run("Signature of Print Payment Advice (PPA) under SDRF Parent Account")
        run_title.bold = True
        doc.add_paragraph() 

        p_body = doc.add_paragraph()
        p_body.add_run("Print payment Advice (PPA) in respect of ")
        p_body.add_run(subsidiary).underline = True 
        p_body.add_run(" under SDRF parent account has been placed below for JD (DM) signature please.")
        doc.add_paragraph()

        num_rows = 1 + len(transaction_list) + 1
        table = doc.add_table(rows=num_rows, cols=4)
        table.style = 'Table Grid'
        
        headers = ["Sl. No.", "In favour of", "Print Payment Advice No.", "Amount"]
        for i, text in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = text
            for p in cell.paragraphs:
                for run in p.runs: run.font.bold = True
        
        grand_total = 0
        for idx, (ppa, amount_str) in enumerate(transaction_list):
            row_idx = idx + 1 
            row_cells = table.rows[row_idx].cells
            val = _parse_rupee(amount_str)
            grand_total += val
            row_cells[0].text = str(idx + 1) + "."
            row_cells[1].text = subsidiary
            row_cells[2].text = str(ppa)
            row_cells[3].text = amount_str

        last_row = table.rows[-1]
        last_row.cells[2].text = "G/TOTAL"
        for p in last_row.cells[2].paragraphs:
            for r in p.runs: r.font.bold = True
            
        last_row.cells[3].text = _fmt_rupee(grand_total)
        for p in last_row.cells[3].paragraphs:
            for r in p.runs: r.font.bold = True

        doc.save(filename)
        return True, os.path.abspath(filename)
    except PermissionError: return False, "Error: Word file is open. Close it."
    except Exception as e: return False, str(e)

# --- PDF GENERATION ---
def generate_summary_pdf(data_list):
    filename = f"Financial_Report_{datetime.now().strftime('%d-%m-%Y')}.pdf"
    
    try:
        doc = SimpleDocTemplate(filename, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()
        
        # 1. Title
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            alignment=TA_CENTER,
            textColor=colors.black
        )
        elements.append(Paragraph("Financial Status Report", title_style))
        
        # 2. Date
        date_style = ParagraphStyle(
            'DateStyle',
            parent=styles['Normal'],
            fontSize=10,
            alignment=TA_LEFT,
            spaceAfter=10,
            textColor=colors.black
        )
        date_str = datetime.now().strftime("%d-%m-%Y %H:%M %p")
        elements.append(Paragraph(f"Generated on: {date_str}", date_style))
        elements.append(Spacer(1, 10))

        # 3. Table Data Prep (LABELS CHANGED HERE)
        table_data = [['Department', 'Approved Limit', 'Total Sanctioned', 'Remaining Balance']]
        
        for row in data_list:
            sub, limit, spent, bal = row
            table_data.append([
                sub,
                _fmt_rupee(limit).replace("₹", "Rs. "),
                _fmt_rupee(spent).replace("₹", "Rs. "),
                _fmt_rupee(bal).replace("₹", "Rs. ")
            ])

        # 4. Table Styling
        col_widths = [180, 110, 110, 110]
        
        t = Table(table_data, colWidths=col_widths)
        
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ]))
        
        elements.append(t)
        
        doc.build(elements)
        return True, os.path.abspath(filename)
        
    except PermissionError:
        return False, "Error: PDF file is open. Close it and try again."
    except Exception as e:
        return False, str(e)