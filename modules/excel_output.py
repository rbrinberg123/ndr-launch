import io
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL  = PatternFill('solid', start_color='1F3864')
HEADER_FONT  = Font(name='Arial', bold=True, color='FFFFFF', size=10)
BODY_FONT    = Font(name='Arial', size=10)
ALT_FILL     = PatternFill('solid', start_color='EEF2F7')
NO_FILL      = PatternFill(fill_type=None)
CENTER       = Alignment(horizontal='center', vertical='center', wrap_text=False)
LEFT         = Alignment(horizontal='left',   vertical='center', wrap_text=False)
THIN         = Border(bottom=Side(style='thin', color='D9D9D9'))

NUMERIC_COLS = {
    'Shares', 'Fund Shares', 'Passive or Index Shares',
    'Total Funds', 'Passive or Index Funds', 'Match Count',
    'Contact Equity Assets (USD, mm)', 'EAUM ($mm)',
    'Account Equity % Portfolio Turnover', 'AUM ($mm)',
}


def _format_sheet(ws):
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = CENTER
        cell.border    = Border(bottom=Side(style='medium', color='FFFFFF'))
    ws.row_dimensions[1].height = 20

    for row_idx in range(2, ws.max_row + 1):
        fill = ALT_FILL if row_idx % 2 == 0 else NO_FILL
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font   = BODY_FONT
            cell.fill   = fill
            cell.border = THIN
            if header in NUMERIC_COLS:
                cell.alignment = CENTER
                if header == 'Account Equity % Portfolio Turnover':
                    cell.number_format = '#,##0.0'
                elif header in ('EAUM ($mm)', 'AUM ($mm)', 'Contact Equity Assets (USD, mm)'):
                    cell.number_format = '#,##0'
                else:
                    cell.number_format = '#,##0'
            else:
                cell.alignment = LEFT

    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(header)) if header else 8
        for row_idx in range(2, min(ws.max_row + 1, 300)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 42)

    ws.freeze_panes = 'A2'
    if ws.max_column > 0 and ws.max_row > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}1"


def generate_excel(results, company_name):
    output = io.BytesIO()

    sheets = {'Contacts': results['main']}
    if results.get('hf') is not None and len(results['hf']) > 0:
        sheets['HFs'] = results['hf']
    if results.get('dnc') is not None and len(results['dnc']) > 0:
        sheets['DNC'] = results['dnc']
    if results.get('check') is not None and len(results['check']) > 0:
        sheets['Check'] = results['check']
    if results.get('quant') is not None and len(results['quant']) > 0:
        sheets['Quant'] = results['quant']

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            if df is not None and len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    wb = load_workbook(output)
    for sheet_name in wb.sheetnames:
        _format_sheet(wb[sheet_name])

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output.read()
