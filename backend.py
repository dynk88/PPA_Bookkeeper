import pandas as pd
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from doc_gen import generate_payment_advice, generate_summary_pdf
from config import Config  # IMPORT CONFIG

class BookkeepingSystem:
    def __init__(self):
        self.ensure_file_exists()

    def ensure_file_exists(self):
        if not os.path.exists(Config.DB_FILENAME):
            wb = openpyxl.Workbook()
            ws_limits = wb.active
            ws_limits.title = Config.SHEET_LIMITS
            # Note: We keep "Subsidiary" internal key for database compatibility
            ws_limits.append(["Subsidiary", "Approved_Limit"])
            wb.create_sheet(Config.SHEET_TXN)
            wb.save(Config.DB_FILENAME)
            print("Created new database file.")

    def get_subsidiaries(self):
        try:
            df = pd.read_excel(Config.DB_FILENAME, sheet_name=Config.SHEET_LIMITS)
            if df.empty: return []
            return df["Subsidiary"].dropna().unique().tolist()
        except: return []

    def get_limit_info(self, subsidiary):
        try:
            df = pd.read_excel(Config.DB_FILENAME, sheet_name=Config.SHEET_LIMITS)
            row = df[df["Subsidiary"] == subsidiary]
            if row.empty: return 0
            return int(row.iloc[0]["Approved_Limit"])
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
        try:
            value = int(value)
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

    def save_batch(self, subsidiary, batch_list):
        try:
            wb = openpyxl.load_workbook(Config.DB_FILENAME)
            ws = wb[Config.SHEET_TXN]
        except Exception as e:
            return False, f"Error: Close Excel file! {e}"

        limit = self.get_limit_info(subsidiary)
        start_col = self._get_or_create_subsidiary_columns(wb, ws, subsidiary)

        col_ppa = start_col
        col_amt = start_col + 2
        
        current_spent = 0
        existing_ppas = []
        
        for row in range(3, ws.max_row + 1):
            val_ppa = ws.cell(row=row, column=col_ppa).value
            val_amt = ws.cell(row=row, column=col_amt).value
            if val_ppa: existing_ppas.append(str(val_ppa))
            if isinstance(val_amt, (int, float)): current_spent += val_amt

        batch_total = 0
        new_ppas = []
        
        for (ppa, _, amt) in batch_list:
            ppa = str(ppa)
            if ppa in existing_ppas:
                return False, f"Error: PPA {ppa} already exists in database."
            if ppa in new_ppas:
                return False, f"Error: PPA {ppa} is entered twice in this session."
            new_ppas.append(ppa)
            batch_total += amt
            
        final_total = current_spent + batch_total
        if final_total > limit:
            remaining = limit - current_spent
            msg = (
                "Transaction Rejected: Limit Exceeded\n\n"
                f"Approved Limit:      {self._fmt_money(limit)}\n"
                f"Previously Spent:    {self._fmt_money(current_spent)}\n"
                f"Current Batch:       {self._fmt_money(batch_total)}\n"
                "----------------------------------------\n"
                f"Total Required:      {self._fmt_money(final_total)}\n"
                f"Available Balance:   {self._fmt_money(remaining)}"
            )
            return False, msg

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        
        current_row = 3
        while ws.cell(row=current_row, column=col_ppa).value is not None:
            current_row += 1
            
        for (ppa, date_obj, amt) in batch_list:
            cell_ppa = ws.cell(row=current_row, column=col_ppa, value=ppa)
            cell_ppa.border = thin_border
            
            cell_date = ws.cell(row=current_row, column=col_ppa+1, value=date_obj)
            cell_date.number_format = 'DD-MM-YYYY' 
            cell_date.border = thin_border
            
            cell_amt = ws.cell(row=current_row, column=col_ppa+2, value=amt)
            cell_amt.number_format = '"₹" #,##0'
            cell_amt.border = thin_border
            
            current_row += 1

        try:
            wb.save(Config.DB_FILENAME)
        except PermissionError:
            return False, "Error: File is open. Close Excel and try again."

        return True, "Batch Saved Successfully."

    def get_summary_report(self):
        try:
            df_limits = pd.read_excel(Config.DB_FILENAME, sheet_name=Config.SHEET_LIMITS)
        except: return []
        if df_limits.empty: return []

        summary_data = []
        try:
            wb = openpyxl.load_workbook(Config.DB_FILENAME, data_only=True) 
            ws = wb[Config.SHEET_TXN]
        except: return []

        sub_col_map = {}
        for col in range(1, ws.max_column + 2):
            val = ws.cell(row=1, column=col).value
            if val: sub_col_map[val] = col

        for index, row in df_limits.iterrows():
            sub_name = row["Subsidiary"]
            limit = int(row["Approved_Limit"])
            spent = 0
            
            if sub_name in sub_col_map:
                start_col = sub_col_map[sub_name]
                amount_col = start_col + 2
                for r in range(3, ws.max_row + 1):
                    val = ws.cell(row=r, column=amount_col).value
                    if isinstance(val, (int, float)):
                        spent += int(val)
            
            remaining = limit - spent
            summary_data.append((sub_name, limit, spent, remaining))
            
        return summary_data

    def create_word_advice(self, subsidiary, date_str, transaction_list):
        return generate_payment_advice(subsidiary, date_str, transaction_list)

    def search_transactions(self, subsidiary=None, ppa_text=None):
        try:
            wb = openpyxl.load_workbook(Config.DB_FILENAME, data_only=True)
            ws = wb[Config.SHEET_TXN]
        except:
            return []

        results = []
        for col in range(1, ws.max_column + 1, 3):
            sub_name = ws.cell(row=1, column=col).value
            if not sub_name: continue
            
            if subsidiary and subsidiary != "All Departments" and sub_name != subsidiary:
                continue
                
            for row in range(3, ws.max_row + 1):
                ppa = ws.cell(row=row, column=col).value
                date_val = ws.cell(row=row, column=col+1).value
                amt = ws.cell(row=row, column=col+2).value
                
                if not ppa: continue 
                
                if ppa_text:
                    if str(ppa_text).upper() not in str(ppa).upper():
                        continue
                
                results.append((sub_name, str(ppa), date_val, amt))
        
        return results
    
    def create_dashboard_pdf(self, summary_data):
        return generate_summary_pdf(summary_data)