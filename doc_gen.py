import os
from datetime import datetime
from docx import Document
from docx.shared import Pt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape # CHANGED: Added landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

def _fmt_rupee(value):
    try: value = int(value)
    except: return str(value)
    s = str(value)
    if len(s) <= 3: return f"₹ {s}"
    last_three = s[-3:]
    remaining = s[:-3]
    formatted_remaining = ""
    for i, digit in enumerate(reversed(remaining)):
        if i > 0 and i % 2 == 0: formatted_remaining = "," + formatted_remaining
        formatted_remaining = digit + formatted_remaining
    return f"₹ {formatted_remaining},{last_three}"

def _parse_rupee(value_str):
    try: return int(str(value_str).replace("₹", "").replace(",", "").strip())
    except: return 0

# --- WORD GENERATION (Unchanged) ---
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
        p = doc.add_paragraph()
        p.add_run("Reg :-  ").bold = True
        p.runs[0].underline = True
        p.add_run("Signature of Print Payment Advice (PPA) under SDRF Parent Account").bold = True
        doc.add_paragraph()
        p2 = doc.add_paragraph()
        p2.add_run("Print payment Advice (PPA) in respect of ")
        p2.add_run(subsidiary).underline = True 
        p2.add_run(" under SDRF parent account has been placed below for JD (DM) signature please.")
        doc.add_paragraph()
        table = doc.add_table(rows=1+len(transaction_list)+1, cols=4)
        table.style = 'Table Grid'
        headers = ["Sl. No.", "In favour of", "Print Payment Advice No.", "Amount"]
        for i, t in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = t
            for p in cell.paragraphs: 
                for r in p.runs: r.font.bold = True
        total = 0
        for idx, (ppa, amt) in enumerate(transaction_list):
            row = table.rows[idx+1].cells
            total += _parse_rupee(amt)
            row[0].text = f"{idx+1}."
            row[1].text = subsidiary
            row[2].text = str(ppa)
            row[3].text = amt
        last = table.rows[-1].cells
        last[2].text = "G/TOTAL"
        last[3].text = _fmt_rupee(total)
        for c in [last[2], last[3]]:
            for p in c.paragraphs:
                for r in p.runs: r.font.bold = True
        doc.save(filename)
        return True, os.path.abspath(filename)
    except Exception as e: return False, str(e)

# --- PDF GENERATION (UPDATED FOR LANDSCAPE) ---
def generate_summary_pdf(data_list):
    filename = f"Financial_Report_{datetime.now().strftime('%d-%m-%Y')}.pdf"
    
    try:
        # CHANGED: Use landscape(A4)
        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('CT', parent=styles['Heading1'], fontSize=18, alignment=TA_CENTER, textColor=colors.black, spaceAfter=20)
        elements.append(Paragraph("Financial Status Report (Quarterly)", title_style))
        
        date_str = datetime.now().strftime("%d-%m-%Y %H:%M %p")
        elements.append(Paragraph(f"Generated on: {date_str}", styles['Normal']))
        elements.append(Spacer(1, 10))

        # UPDATED COLUMNS
        headers = ['Department', 'Approved Limit', 'Q1 (Apr-Jun)', 'Q2 (Jul-Sep)', 'Q3 (Oct-Dec)', 'Q4 (Jan-Mar)', 'Total Sanctioned', 'Balance']
        table_data = [headers]
        
        for row in data_list:
            # row: (Sub, Limit, Q1, Q2, Q3, Q4, Total, Bal)
            clean_row = [row[0]] # Name
            # Format numbers
            for val in row[1:]:
                clean_row.append(_fmt_rupee(val).replace("₹", "Rs. "))
            table_data.append(clean_row)

        # ADJUST WIDTHS FOR LANDSCAPE (Total width ~800)
        # Dept gets 180, Numbers get ~85 each
        col_widths = [180, 85, 85, 85, 85, 85, 95, 95]
        
        t = Table(table_data, colWidths=col_widths)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTSIZE', (0, 0), (-1, -1), 9), # Slightly smaller font
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('PADDING', (0, 0), (-1, -1), 4),
        ]))
        
        elements.append(t)
        doc.build(elements)
        return True, os.path.abspath(filename)
    except Exception as e: return False, str(e)