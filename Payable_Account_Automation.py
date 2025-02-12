# Packages
import numpy as np
import pandas as pd
from datetime import datetime
import os

import xlsxwriter
from typing import Dict, List, Union, Any
import xlwings as xw

import warnings
warnings.filterwarnings('ignore')

import winreg

def enable_vba_access():
    """Enable programmatic access to the VBA project object model in Excel via registry modification."""
    try:
        office_versions = ["16.0", "15.0", "14.0", "12.0"]  # Office 2016, 2013, 2010, 2007
        for version in office_versions:
            reg_path = f"SOFTWARE\\Microsoft\\Office\\{version}\\Excel\\Security"
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_ALL_ACCESS)
                winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                winreg.CloseKey(key)
                print(f"Enabled VBA access for Office {version}.")
                break
            except FileNotFoundError:
                continue  # If the registry path doesn't exist, try the next version
        else:
            print("Could not enable VBA access. Registry key not found.")
    except Exception as e:
        print(f"Failed to enable VBA access: {e}")

import sys
print("Python interpreter being used:", sys.executable)


# Get the base directory dynamically
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
PROCESSED_FOLDER = os.path.join(BASE_DIR, "processed")

# Ensure directories exist
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# File paths from the uploads folder
bank_balance_path = os.path.join(UPLOAD_FOLDER, "bank_balance.xlsx")
account_payables_path = os.path.join(UPLOAD_FOLDER, "account_payables.xlsx")
cash_management_path = os.path.join(UPLOAD_FOLDER, "cash_management.xlsx")

# Ensure all required files exist before reading
if not all(os.path.exists(f) for f in [bank_balance_path, account_payables_path, cash_management_path]):
    raise FileNotFoundError("One or more required input files are missing.")

# Read Bank Balance (ensure sheet name is dynamically determined)
bb_xl = pd.ExcelFile(bank_balance_path)
bb_sheet_name = "Balance" if "Balance" in bb_xl.sheet_names else bb_xl.sheet_names[0]
bb_raw = pd.read_excel(bank_balance_path, sheet_name=bb_sheet_name)
print(f"Shape of bank_balance_raw: {bb_raw.shape}\n")

# Read Account Payables
ap_raw = pd.read_excel(account_payables_path)
print(f"Shape of ap_raw: {ap_raw.shape}\n")

# Read Cash Management (always take the last sheet dynamically)
cm_xl = pd.ExcelFile(cash_management_path)
last_sheet = cm_xl.sheet_names[-1]  # Automatically selects the last sheet
cm_raw = pd.read_excel(cash_management_path, sheet_name=last_sheet)
print(f"Shape of cm_raw: {cm_raw.shape}\n")

# Get Account Payable DataFrame
ap = ap_raw.copy()
translation_dict = {
    'Code de fournisseur': 'Supplier code',
    'Immeuble': 'Building',
    'Nom du fournisseur': 'Supplier name',
    'Compagnie': 'Company',
    'Commentaire': 'Comment',
    'Montant payé': 'Paid amount',
    'No facture': 'Invoice no'
}
ap.columns = [translation_dict.get(col, col) for col in ap.columns]

ap.drop_duplicates(inplace=True)
ap['Paid amount'] = pd.to_numeric(ap['Paid amount'], errors='coerce').fillna(0)

ap['Date'] = pd.to_datetime(ap['Date'], errors='coerce')
ap = ap[ap['Date'] >= '2023-10-01']

ap = ap[ap['Total'] != ap['Paid amount']]

# CT stands for reverse payment
ap.loc[ap['Comment'].str[:2] == 'CT', 'Total'] = pd.to_numeric(ap['Total'], errors='coerce') * (-1)

# Get Bank Balance and Cash Management DataFrames
bb = bb_raw.copy()
cm = cm_raw.copy()
cm = cm.rename(columns={'Co. no.': 'Company'})
cm['Available'] = pd.to_numeric(cm['Available'], errors='coerce')
cm['Available'] = cm['Available'].replace(0, np.nan)
# Flip the sign of 'Available'
cm['Available'] = cm['Available'] * (-1)

# Check if the Account is a valid four-digit number (including strings with leading zeros)
def valid_account(val):
    if isinstance(val, str) and val.isdigit():
        if len(val) == 4:
            return val
        else:
            return val.zfill(4)
    else:
        return 'nan'

# Ensure both dataframes are valid for iteration and modification
def clean_column(df, column):
    if column in df.columns:
        df[column] = df[column].astype(str).str.strip().replace(r'\.0$', '', regex=True).str.lower()
        if column == 'Bank':
            df[column] = df[column].apply(valid_account)
    return df

merge_keys = ['Company', 'Building', 'Bank']
dataframes = [ap, bb, cm]
for df in dataframes:
    for key in merge_keys:
        if key in df.columns:
            df = clean_column(df, key)

# Perform the merge
df = pd.merge(
    ap,
    bb[['Company', 'Company Name']],
    on='Company',
    how='left'
)

df = pd.merge(
    df,
    bb[['Company', 'Building', 'Bank', 'Bank Account', 'Status']],
    on=['Company', 'Building'],
    how='left'
)

df = pd.merge(
    df,
    cm[['Company', 'Bank', 'Available']],
    on=['Company', 'Bank'],
    how='left'
)

#df.loc[df['Building'] == 'nan', ['Bank Account','Available']] = np.nan

# Calculate Balance of Account Payable
df['Payable Balance'] = df['Total'] - df['Paid amount']


# Exclude PPA suppliers from the balance sheet
PPA_supplier_codes = [
    'ALT003', 'BEL001', 'BRA001', 'CONR001', 'ENE001', 'ENVIROCONN', 'GAZIFERE',
    'HYDROSOL', 'HYDRO', 'HYDRO WEST', 'INTELECOM', 'MILLER WAS', 'NOVA SCOTI',
    'PRIMACO', 'SUPERIEUR', 'VIDEOTRON']
df_clean = df[~df['Supplier code'].isin(PPA_supplier_codes)]

# Remove all records whose 'Status' == 'REMOVE'
df_clean = df_clean[df_clean['Status'] != 'REMOVE']

cols_to_keep = ['Company Name', 'Bank Account', 'Available', 'Supplier name', 'Date',
                'Invoice no', 'Comment', 'Total', 'Paid amount', 'Payable Balance', 'Status']
df_keep = df_clean[cols_to_keep]
df_keep.columns = [col.title() for col in df_keep.columns]
df_keep.rename(columns={'Bank Account': 'Bank'}, inplace=True)
df_AR = df_keep[df_keep['Supplier Name'] == 'Gestion Hazout Inc']

df_keep = df_keep.set_index(['Company Name', 'Bank', 'Available', 'Supplier Name']).sort_index()
df_keep.sort_values(by=['Date', 'Invoice No'], ascending=True, inplace=True)
df_keep.drop_duplicates(inplace=True)
df_AR = df_AR.set_index(['Company Name', 'Bank', 'Available', 'Supplier Name']).sort_index()
df_AR.sort_values(by=['Date', 'Invoice No'], ascending=True, inplace=True)
df_AR.drop_duplicates(inplace=True)

# Filter rows based on Status
df_active = df_keep[df_keep['Status'] == 'ACTIVE']
df_active.drop(columns=['Status'], inplace=True)
df_zagora = df_keep[df_keep['Status'] == 'ZAGORA']
df_zagora.drop(columns=['Status'], inplace=True)
df_others = df_keep[(df_keep['Status'] != 'ACTIVE') & (df_keep['Status'] != 'ZAGORA')]


class ExcelReportGenerator:
    def __init__(self, output_file: str):
        """Initialize the Excel Report Generator.

        Args:
            output_file (str): The path where the Excel file will be saved
        """
        self.output_file = output_file
        self.writer = None
        self.workbook = None
        self.header_format = None
        self.date_format = None

    def __enter__(self):
        self.writer = pd.ExcelWriter(
            self.output_file,
            engine='xlsxwriter',
            engine_kwargs={'options': {'nan_inf_to_errors': True}}
        )
        self.workbook = self.writer.book
        self.header_format = self.workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'bottom',
            'font_name': 'Calibri',
            # Add top and bottom thick border lines
            'top': 2,
            'bottom': 2,
        })
        self.date_format = self.workbook.add_format({
            'num_format': 'yyyy-mm-dd',
            'align': 'center'
        })
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.writer is not None:
            self.writer.close()

    def _clean_value(self, value: Any) -> Union[float, str, None]:
        """Clean and prepare values for Excel writing."""
        if pd.isna(value):
            return None
        if isinstance(value, (float, np.floating)):
            if np.isinf(value) or np.isnan(value):
                return None
            return float(value)
        if isinstance(value, (int, np.integer)):
            return float(value)
        if isinstance(value, str) and value.strip() == '':
            return None
        return value

    def _safe_numeric(self, value: Any) -> float:
        """Safely convert value to numeric, preserving NaN values."""
        if pd.isna(value) or value == '':
            return None
        try:
            return float(value)
        except (ValueError, TypeError):
            return None

    def _safe_date(self, value: Any) -> Union[datetime, str]:
        """Safely convert value to datetime, returning empty string if not possible."""
        if pd.isna(value) or value == '':
            return ''
        try:
            if isinstance(value, datetime):
                return value
            elif isinstance(value, str):
                return pd.to_datetime(value).to_pydatetime()
            else:
                return ''
        except (ValueError, TypeError):
            return ''

    def _adjust_column_widths(self, worksheet, *dfs: pd.DataFrame):
        """Adjust column widths based on content."""
        max_cols = max((len(df.columns) for df in dfs if not df.empty), default=0)
        for col_idx in range(max_cols):
            max_width = 0
            for df in dfs:
                if not df.empty and col_idx < len(df.columns):
                    col_values = df.iloc[:, col_idx]
                    # Handle datetime columns
                    if pd.api.types.is_datetime64_any_dtype(col_values):
                        col_values = col_values.dt.strftime('%Y-%m-%d')

                    col_values = col_values.astype(str)
                    max_width = max(
                        max_width,
                        col_values.map(len).max(),
                        len(str(df.columns[col_idx]))
                    )
            worksheet.set_column(col_idx, col_idx, max_width + 2)

    def _apply_conditional_formatting(self, worksheet, net_col_idx: int, start_row: int, end_row: int):
        """Apply conditional formatting for negative values in 'Net of Balance'."""
        worksheet.conditional_format(
            f'{xlsxwriter.utility.xl_rowcol_to_cell(start_row, net_col_idx)}:' +
            f'{xlsxwriter.utility.xl_rowcol_to_cell(end_row, net_col_idx)}',
            {
                'type': 'cell',
                'criteria': '<',
                'value': 0,
                'format': self.workbook.add_format({
                    'font_color': 'red',
                    'num_format': '#,##0.00_);(#,##0.00)'
                })
            }
        )

    def create_sheet(self, sheet_name: str, df: pd.DataFrame, has_status: bool = False, include_grand_total: bool = True, hidden: bool = False):
        """Creates an individual sheet for each stutas and every active company."""
        worksheet = self.workbook.add_worksheet(sheet_name)
        self.writer.sheets[sheet_name] = worksheet

        if hidden:
            worksheet.hide()

        # 1) Write headers
        headers = ['Company Name', 'Bank', 'Available', 'Supplier Name', 'Date',
                'Invoice No', 'Comment', 'Total', 'Paid Amount', 'Sum of Balance', 'Net of Balance']
        if has_status:
            headers.append('Status')
        worksheet.write_row(0, 0, headers, self.header_format)

        # Freeze the first row
        worksheet.freeze_panes(1, 0)

        df.reset_index(inplace=True)
        current_row = 1
        total_available = []
        total_sum = []

        # 2) Iterate over each Company
        for comp_key, comp_group in df.groupby('Company Name'):
            company_start_row = current_row
            company_availables = []
            company_totals = []

            company = comp_key
            first_bank = True

            # 3) Iterate over each Bank
            for bank_key, bank_group in comp_group.groupby(['Bank', 'Available'], dropna=False):
                bank, available = bank_key
                available = self._safe_numeric(available)
                available = '' if pd.isna(available) else available

                if first_bank:
                    # Write row with Company + Bank
                    company_header_data = [
                        self._clean_value(company),
                        self._clean_value(bank),
                        available,
                        '', '', '', '', '', '', '', ''
                    ]
                    worksheet.write_row(current_row, 0, company_header_data)
                    first_bank = False
                else:
                    # Write row with Bank only
                    bank_header_data = [
                        '',
                        self._clean_value(bank),
                        available,
                        '', '', '', '', '', '', '', ''
                    ]
                    worksheet.write_row(current_row, 0, bank_header_data)

                current_row += 1
                bank_start_row = current_row
                bank_totals = []

                # 4) Iterate over each Supplier
                for supplier_key, supplier_group in bank_group.groupby('Supplier Name'):
                    # ↓↓↓ Assign supplier_start_row here before the supplier header row
                    supplier_start_row = current_row

                    supplier = supplier_key
                    supplier_header_data = [
                        '',
                        '',
                        '',
                        self._clean_value(supplier),
                        '', '', '', '', '', '', ''
                    ]
                    worksheet.write_row(current_row, 0, supplier_header_data)
                    worksheet.set_row(current_row, None, None, {'level': 1, 'hidden': False})
                    current_row += 1

                    # !!! IMPORTANT: Do not reassign supplier_start_row here !!!
                    # so that the sum range for the supplier total remains correct.

                    # 5) Detail rows
                    for _, detail_row in supplier_group.iterrows():
                        detail_data = [
                            '',
                            '',
                            '',
                            '',
                            self._safe_date(detail_row['Date']),
                            self._clean_value(detail_row['Invoice No']),
                            self._clean_value(detail_row['Comment']),
                            self._clean_value(self._safe_numeric(detail_row['Total'])),
                            self._clean_value(self._safe_numeric(detail_row['Paid Amount'])),
                            self._clean_value(self._safe_numeric(detail_row['Payable Balance'])),
                            ''  # Net of Balance placeholder
                        ]
                        if has_status:
                            detail_data.append(self._clean_value(detail_row['Status']))

                        worksheet.write_row(current_row, 0, detail_data)
                        worksheet.set_row(current_row, None, None, {'level': 2, 'hidden': False})
                        current_row += 1

                    # 6) Supplier total
                    supplier_total_row = current_row
                    # Sum from J( supplier_start_row+2 ) to J( current_row )
                    supplier_total_formula = f"=SUM(J{supplier_start_row + 2}:J{current_row})"
                    supplier_total_data = [
                        '',
                        '',
                        '',
                        f'{supplier} Total',
                        '', '', '', '', '',
                        supplier_total_formula,
                        ''
                    ]
                    worksheet.write_row(
                        current_row, 0, supplier_total_data,
                        self.workbook.add_format({'bold': True, 'num_format': '#,##0.00_);(#,##0.00)'})
                    )
                    worksheet.set_row(current_row, None, None, {'level': 1, 'hidden': False})
                    current_row += 1

                    bank_totals.append(f"J{supplier_total_row + 1}")

                # 7) Bank total
                bank_total_row = current_row
                bank_available_formula = f"=C{bank_start_row}" if available else ''
                bank_total_formula = f"=SUM({','.join(map(str, bank_totals))})"
                bank_net_formula = f"=C{bank_total_row + 1} - J{bank_total_row + 1}"
                bank_total_data = [
                    '',
                    f'{bank} Total',
                    bank_available_formula,
                    '',
                    '', '', '', '', '',
                    bank_total_formula,
                    bank_net_formula
                ]
                worksheet.write_row(
                    current_row, 0, bank_total_data,
                    self.workbook.add_format({
                        'bold': True,
                        'bg_color': '#E8E8E8',
                        'num_format': '#,##0.00_);(#,##0.00)'
                    })
                )
                current_row += 1

                if available != '':
                    company_availables.append(f"C{bank_total_row + 1}")
                company_totals.append(f"J{bank_total_row + 1}")

            # 8) Company total
            company_total_row = current_row
            company_total_formula = f"=SUM({','.join(map(str, company_totals))})"
            company_available_formula = f"=SUM({','.join(map(str, company_availables))})" if company_availables else ""
            company_net_formula = f"=C{company_total_row + 1} - J{company_total_row + 1}"
            company_total_data = [
                f'{company} Total',
                '',
                company_available_formula,
                '',
                '', '', '', '', '',
                company_total_formula,
                company_net_formula
            ]
            worksheet.write_row(
                current_row, 0, company_total_data,
                self.workbook.add_format({
                    'bold': True,
                    'bg_color': '#D3D3D3',
                    'num_format': '#,##0.00_);(#,##0.00)',
                    'bottom': 2
                })
            )
            current_row += 1

            if company_available_formula != '':
                total_available.append(f"C{company_total_row + 1}")
            total_sum.append(f"J{company_total_row + 1}")

        # 9) Grand total
        if include_grand_total:
            grand_total_row = current_row
            grand_available_formula = f"=SUM({','.join(map(str, total_available))})" if total_available else ""
            grand_total_formula = f"=SUM({','.join(map(str, total_sum))})"
            grand_net_formula = f"=C{grand_total_row + 1} - J{grand_total_row + 1}"
            grand_total_data = [
                'Grand Total',
                '',
                grand_available_formula,
                '',
                '', '', '', '', '',
                grand_total_formula,
                grand_net_formula
            ]
            worksheet.write_row(
                current_row, 0, grand_total_data,
                self.workbook.add_format({
                    'bold': True,
                    'bg_color': '#B0B0B0',
                    'num_format': '#,##0.00_);(#,##0.00)',
                    'bottom': 2
                })
            )

        # Hide/unhide columns as you had before
        worksheet.set_column(3, 3, None, None, {'hidden': False, 'level': 0})
        worksheet.set_column(4, 5, None, None, {'hidden': False, 'level': 1})

        # Auto-adjust column widths
        self._adjust_column_widths(worksheet, df)
        worksheet.set_column(4, 4, 16, self.date_format)

        # Conditional formatting for Net of Balance
        self._apply_conditional_formatting(worksheet, 10, 1, current_row)

        # Numeric formatting
        number_format = self.workbook.add_format({
            'num_format': '#,##0.00_);(#,##0.00)',
            'align': 'right'
        })
        numeric_col_names = ['Available', 'Total', 'Paid Amount', 'Sum of Balance', 'Net of Balance']
        for idx, col_name in enumerate(headers):
            if col_name in numeric_col_names:
                worksheet.set_column(idx, idx, 16, number_format)

        # Optionally hide Comment, Total, Paid Amount, and/or Status columns
        worksheet.set_column(6, 8, None, number_format, {'hidden': True})
        if has_status:
            status_col_idx = headers.index('Status')
            worksheet.set_column(11, 11, None, number_format, {'hidden': True})


    def generate_report(self, df_active: pd.DataFrame, df_others: pd.DataFrame,
                        df_zagora: pd.DataFrame, df_AR: pd.DataFrame):
        """Generate Excel report from four dfs."""
        self.create_sheet('Active', df_active)
        self.create_sheet('Others', df_others, has_status=True)
        self.create_sheet('Zagora_AP', df_zagora)
        self.create_sheet('Zagora_AR', df_AR, has_status=True)

        # Add a sheet for each company in df_active after the tab 'Zagora_AR'
        companies = df_active['Company Name'].unique()
        for company in companies:
            df_company = df_active[df_active['Company Name'] == company].copy()
            if df_company.empty:
                continue # Skip if no data
            sheet_name = str(company)[:31]  # Max sheet name length is 31 characters
            self.create_sheet(
                sheet_name, 
                df_company, 
                include_grand_total=False,
                hidden=True
            )

def add_vba_buttons(output_xlsx):
    """Adds a VBA button to each sheet in the workbook to fix #REF! errors."""
    output_xlsm = output_xlsx.replace('.xlsx', '.xlsm')
    vba_code = """
Sub FixRefErrors()
    Dim ws As Worksheet
    Dim cell As Range    
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                If InStr(cell.Formula, "#REF!") > 0 Then
                    cell.Formula = Replace(cell.Formula, "#REF!", "0")
                End If
            End If
        Next cell
    Next ws
End Sub
"""

    try:
        with xw.App(visible=False) as app:
            wb = app.books.open(output_xlsx)
            # Add VBA module
            try:
                vb_mod = wb.api.VBProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
                vb_mod.CodeModule.AddFromString(vba_code.strip())
            except Exception as e:
                print(f"Error adding VBA module: {e}")
                raise
            # Add button to each sheet
            for sheet in wb.sheets:
                button_exists = False
                for shape in sheet.shapes:
                    if shape.name == "Fix #REF!":
                        button_exists = True
                        break
                if not button_exists:
                    cell = sheet.range("L1")
                    left, top = cell.left, cell.top
                    button = sheet.api.Buttons().Add(left, top, 100, 14.4)
                    button.OnAction = "FixRefErrors"
                    button.Name = "Fix #REF!"
                    button.Text = "Fix #REF!"
            
            # Save as .xlsm file
            wb.save(output_xlsm)
            wb.close()
        
        # Remove the original .xlsx file
        os.remove(output_xlsx)
    except Exception as e:
        print(f"An error occurred during VBA buttons addition: {e}")
        raise
    return output_xlsm

# Get the user's Downloads folder path
downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')

today_date = datetime.today().strftime('%Y-%m-%d')
output_xlsx = os.path.join(downloads_dir, f"Payables Summary_{today_date}.xlsx")

try:
    with ExcelReportGenerator(output_xlsx) as report_generator:
        report_generator.generate_report(df_active, df_others, df_zagora, df_AR)
    
    # Convert to xlsm and add VBA buttons
    output_xlsm = os.path.join(downloads_dir, add_vba_buttons(output_xlsx))
    print(f"Report generated successfully: {output_xlsm}")
except Exception as e:
    print(f"An error occurred during report generation: {e}")
    raise

