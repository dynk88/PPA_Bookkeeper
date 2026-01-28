import pandas as pd
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from doc_gen import generate_payment_advice, generate_summary_pdf
from config import Config

class BookkeepingSystem:
    def __init__(self):
        # Default active sheet is based on TODAY, but saving is dynamic
        self.ensure_file_exists()

    def get_sheet_name_for_date(self, date_obj):
        """
        Determines the correct sheet name based on the specific transaction date.
        Format: Transactions_YYYY_YY (e.g., Transactions_2026_27)
        """
        if date_obj.month >= 4: # April to Dec
            start_year = date_obj.year
            end_year = date_obj.year + 1
        else: # Jan to March
            start_year = date_obj.year - 1
            end_year = date_obj.year
            
        short_end = str(end_year)[-2:]
        return f"{Config.TXN_PREFIX}{start_year}_{short_end}"

    def ensure_file_exists(self):
        if not os.path.exists(Config.DB_FILENAME):
            wb = openpyxl.Workbook()
            # Setup Limits
            ws_limits = wb.active
            ws_limits.title = Config.SHEET_LIMITS
            ws_limits.append(["Department", "Approved_Limit"]) 
            
            # Setup Current FY Sheet (Default)
            current_sheet = self.get_sheet_name_for_date(datetime.now())
            wb.create_sheet(current_sheet)
            wb.save(Config.DB_FILENAME)

    def _ensure_fy_sheet_exists(self, wb, sheet_name):
        """Creates the transaction sheet if it doesn't exist (Year Rollover)"""
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            # No headers needed because columns are dynamically created per department
            return ws
        return wb[sheet_name]

    def get_subsidiaries(self):
        try:
            wb = openpyxl.load_workbook(Config.DB_FILENAME, data_only=True)
            ws = wb[Config.SHEET_LIMITS]
            subs = []
            for row in range(2, ws.max_row + 1):
                val = ws.cell(row=row, column=1).value
                if val: subs.append(val)
            return subs
        except: return []

    def get_limit_info(self, subsidiary):
        try:
            wb = openpyxl.load_workbook(Config.DB_FILENAME, data_only=True)
            ws = wb[Config.SHEET_LIMITS]
            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=1).value == subsidiary:
                    val = ws.cell(row=row, column=2).value
                    return int(val) if val else 0
            return 0
        except: return 0

    def _get_or_create_subsidiary_columns(self, wb, ws, subsidiary):
        bold_font = Font(bold=True)
        center_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))

        for col in range(1, ws.max_column + 2, 3):
            cell_value = ws.cell(row=1, column=col).value
            if cell_value == subsidiary:
                return col
            if cell_value is None:
                start_col = col
                ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=start_col+2)
                title_cell = ws.cell(row=1, column=start_col, value=subsidiary)
                title_cell.font = bold_font
                title_cell.alignment = center_align
                for c in range(start_col, start_col+3):
                    ws.cell(row=1, column=c).border = thin_border

                headers = ["PPA_Number", "Date", "Amount"]
                for i, header in enumerate(headers):
                    cell = ws.cell(row=2, column=start_col + i, value=header)
                    cell.font = bold_font
                    cell.alignment = center_align
                    cell.border = thin_border
                return start_col
        return 1

    def _fmt_money(self, value):
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

    # --- PPA SAVE (FIXED: Uses Batch Date for Sheet) ---
    def save_batch(self, subsidiary, batch_list):
        try:
            wb = openpyxl.load_workbook(Config.DB_FILENAME)
            
            # Determine Sheet Name based on the FIRST entry's date
            # (Assumes all entries in a batch belong to same session/date roughly)
            first_date = batch_list[0][1] # (ppa, date_obj, amt)
            target_sheet_name = self.get_sheet_name_for_date(first_date)
            
            ws = self._ensure_fy_sheet_exists(wb, target_sheet_name)
            
        except Exception as e:
            return False, f"Error: {e}"

        limit = self.get_limit_info(subsidiary)
        start_col = self._get_or_create_subsidiary_columns(wb, ws, subsidiary)
        col_ppa = start_col
        col_amt = start_col + 2
        
        current_spent = 0
        existing_ppas = []
        
        # Calculate Spent in THIS Financial Year
        for row in range(3, ws.max_row + 1):
            val_ppa = ws.cell(row=row, column=col_ppa).value
            val_amt = ws.cell(row=row, column=col_amt).value
            if val_ppa: existing_ppas.append(str(val_ppa))
            if isinstance(val_amt, (int, float)): current_spent += val_amt

        batch_total = 0
        new_ppas = []
        for (ppa, _, amt) in batch_list:
            ppa = str(ppa)
            if ppa in existing_ppas: return False, f"Error: PPA {ppa} already exists in {target_sheet_name}."
            if ppa in new_ppas: return False, f"Error: PPA {ppa} duplicated."
            new_ppas.append(ppa)
            batch_total += amt
            
        final_total = current_spent + batch_total
        if final_total > limit:
            remaining = limit - current_spent
            msg = (f"Limit Exceeded for {target_sheet_name}\nApproved: {self._fmt_money(limit)}\n"
                   f"Spent (This FY): {self._fmt_money(current_spent)}\n"
                   f"Batch: {self._fmt_money(batch_total)}\nAvailable: {self._fmt_money(remaining)}")
            return False, msg

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        current_row = 3
        while ws.cell(row=current_row, column=col_ppa).value is not None:
            current_row += 1
            
        for (ppa, date_obj, amt) in batch_list:
            ws.cell(row=current_row, column=col_ppa, value=ppa).border = thin_border
            d_cell = ws.cell(row=current_row, column=col_ppa+1, value=date_obj)
            d_cell.number_format = 'DD-MM-YYYY'
            d_cell.border = thin_border
            a_cell = ws.cell(row=current_row, column=col_ppa+2, value=amt)
            a_cell.number_format = '"₹" #,##0'
            a_cell.border = thin_border
            current_row += 1

        try:
            wb.save(Config.DB_FILENAME)
        except PermissionError:
            return False, "Error: File open."
        return True, f"Saved to {target_sheet_name}."

    # --- ALLOCATION SAVE (FIXED: Headers) ---
    def save_allocation_batch(self, subsidiary, batch_list):
        try:
            wb = openpyxl.load_workbook(Config.DB_FILENAME)
            ws = wb[Config.SHEET_LIMITS]
        except Exception as e: return False, str(e)

        target_row = None
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == subsidiary:
                target_row = r
                break
        
        if not target_row:
            target_row = ws.max_row + 1
            ws.cell(row=target_row, column=1, value=subsidiary)
            ws.cell(row=target_row, column=2, value=0)

        curr_limit_cell = ws.cell(row=target_row, column=2)
        current_limit = curr_limit_cell.value if curr_limit_cell.value else 0
        
        # Start searching from Col 3 for empty slot
        current_col = 3
        while ws.cell(row=target_row, column=current_col).value is not None:
            current_col += 2 # Skip Amount & Date columns
            
        total_added = 0
        for (_, date_obj, amt) in batch_list:
            total_added += amt
            
            # --- HEADER GENERATION LOGIC ---
            # Calculate Allocation Number: (Col_Index - 1) / 2
            # e.g., Col 3 -> 1st, Col 5 -> 2nd
            alloc_num = int((current_col - 1) / 2)
            
            # Helper for suffix (st, nd, rd, th)
            if 11 <= (alloc_num % 100) <= 13: suffix = "th"
            else:
                rem = alloc_num % 10
                if rem == 1: suffix = "st"
                elif rem == 2: suffix = "nd"
                elif rem == 3: suffix = "rd"
                else: suffix = "th"
            
            header_title = f"{alloc_num}{suffix} allocation"
            header_date = f"Date_{alloc_num}"
            
            # Write Headers if missing (Row 1)
            if ws.cell(row=1, column=current_col).value is None:
                ws.cell(row=1, column=current_col, value=header_title).font = Font(bold=True)
                ws.cell(row=1, column=current_col+1, value=header_date).font = Font(bold=True)

            # Write Data
            ws.cell(row=target_row, column=current_col, value=amt)
            d_cell = ws.cell(row=target_row, column=current_col+1, value=date_obj)
            d_cell.number_format = 'DD-MM-YYYY'
            
            current_col += 2

        curr_limit_cell.value = current_limit + total_added

        try:
            wb.save(Config.DB_FILENAME)
        except PermissionError: return False, "File open."
        return True, f"Allocated {self._fmt_money(total_added)}."

    def get_summary_report(self):
        try:
            df_limits = pd.read_excel(Config.DB_FILENAME, sheet_name=Config.SHEET_LIMITS)
            # Default to TODAY's FY for dashboard
            active_sheet = self.get_sheet_name_for_date(datetime.now())
            wb = openpyxl.load_workbook(Config.DB_FILENAME, data_only=True) 
            
            if active_sheet in wb.sheetnames:
                ws = wb[active_sheet]
            else:
                return [] # No data for this year yet
        except: return []

        sub_col_map = {}
        for col in range(1, ws.max_column + 2):
            val = ws.cell(row=1, column=col).value
            if val: sub_col_map[val] = col

        summary_data = []
        for index, row in df_limits.iterrows():
            sub_name = row["Department"] if "Department" in row else row["Subsidiary"]
            limit = int(row["Approved_Limit"]) if pd.notna(row["Approved_Limit"]) else 0
            
            q1, q2, q3, q4, total_spent = 0, 0, 0, 0, 0
            
            if sub_name in sub_col_map:
                start_col = sub_col_map[sub_name]
                col_date = start_col + 1
                col_amt = start_col + 2
                for r in range(3, ws.max_row + 1):
                    amt = ws.cell(row=r, column=col_amt).value
                    dt = ws.cell(row=r, column=col_date).value
                    if isinstance(amt, (int, float)) and isinstance(dt, datetime):
                        total_spent += amt
                        m = dt.month
                        if 4 <= m <= 6: q1 += amt
                        elif 7 <= m <= 9: q2 += amt
                        elif 10 <= m <= 12: q3 += amt
                        elif 1 <= m <= 3: q4 += amt
            
            remaining = limit - total_spent
            summary_data.append((sub_name, limit, q1, q2, q3, q4, total_spent, remaining))
        return summary_data

    def create_word_advice(self, subsidiary, date_str, transaction_list):
        return generate_payment_advice(subsidiary, date_str, transaction_list)

    def search_transactions(self, subsidiary=None, ppa_text=None, quarter=None):
        try:
            # Search CURRENT FY by default
            active_sheet = self.get_sheet_name_for_date(datetime.now())
            wb = openpyxl.load_workbook(Config.DB_FILENAME, data_only=True)
            if active_sheet not in wb.sheetnames: return []
            ws = wb[active_sheet]
        except: return []

        results = []
        for col in range(1, ws.max_column + 1, 3):
            sub_name = ws.cell(row=1, column=col).value
            if not sub_name: continue
            if subsidiary and subsidiary != "All Departments" and sub_name != subsidiary: continue
            for row in range(3, ws.max_row + 1):
                ppa = ws.cell(row=row, column=col).value
                date_val = ws.cell(row=row, column=col+1).value
                amt = ws.cell(row=row, column=col+2).value
                if not ppa: continue 
                if ppa_text and str(ppa_text).upper() not in str(ppa).upper(): continue
                if quarter and quarter != "All":
                    if not isinstance(date_val, datetime): continue
                    m = date_val.month
                    row_q = ""
                    if 4 <= m <= 6: row_q = "Q1"
                    elif 7 <= m <= 9: row_q = "Q2"
                    elif 10 <= m <= 12: row_q = "Q3"
                    elif 1 <= m <= 3: row_q = "Q4"
                    if row_q != quarter: continue
                results.append((sub_name, str(ppa), date_val, amt))
        return results
    
    def create_dashboard_pdf(self, summary_data):
        return generate_summary_pdf(summary_data)