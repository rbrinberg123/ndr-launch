import io
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill('solid', start_color='1F3864')
HEADER_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=10)
BODY_FONT   = Font(name='Arial', size=10)
ALT_FILL    = PatternFill('solid', start_color='EEF2F7')
NO_FILL     = PatternFill(fill_type=None)
CENTER      = Alignment(horizontal='center', vertical='center')
LEFT        = Alignment(horizontal='left',   vertical='center')
THIN        = Border(bottom=Side(style='thin', color='D9D9D9'))

NUMERIC_COLS = {
    'Shares', 'Fund Shares', 'Passive or Index Shares',
    'Total Funds', 'Passive or Index Funds', 'Match Count',
    'EAUM ($mm)', 'AUM ($mm)', 'T/O %',
    'L12M', 'Total', '3rd Party', 'Rose & Co',
}
DATE_COLS  = {
    'Specifically with Co.', 'Anyone at Inst. with Co', 'Last Meeting', 'As of',
    'Last Mtg btwn Contact & Co', 'Last Mtg btwn firm & Co', 'Last Mtg. w/ Any Co',
}
SHRINK_COLS = {'Industry', 'Geo', 'Style', 'Mkt. Cap'}
SHRINK_ALIGN = Alignment(horizontal='left', vertical='center', shrink_to_fit=True)


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
                if header == 'T/O %':
                    cell.number_format = '#,##0.0'
                elif header in ('EAUM ($mm)', 'AUM ($mm)'):
                    cell.number_format = '#,##0'
                else:
                    cell.number_format = '#,##0'
            elif header in DATE_COLS:
                cell.alignment    = CENTER
                cell.number_format = 'mm/dd/yyyy'
            elif header in SHRINK_COLS:
                cell.alignment = SHRINK_ALIGN
            else:
                cell.alignment = LEFT

    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        if header in SHRINK_COLS:
            ws.column_dimensions[col_letter].width = 27
        else:
            max_len = len(str(header)) if header else 8
            for row_idx in range(2, min(ws.max_row + 1, 300)):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 42)

    ws.freeze_panes = 'A2'
    if ws.max_column > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}1"


def generate_excel(results, company_name):
    frames         = results['frames']
    city_selections = results.get('city_selections') or []
    has_city       = results.get('has_city_routing', False)

    output = io.BytesIO()

    # Determine sheet write order
    sheet_order = []
    if has_city and city_selections:
        for tab_name, _ in city_selections:
            if tab_name in frames and frames[tab_name] is not None and len(frames[tab_name]) > 0:
                sheet_order.append(tab_name)
        # Virtual sub-tabs (city routing mode)
        for vname in ['Virtual - NAM', 'Virtual - EUR', 'Virtual - Other', 'Virtual']:
            if vname in frames and frames[vname] is not None and len(frames[vname]) > 0:
                sheet_order.append(vname)
    else:
        if 'Contacts' in frames and frames['Contacts'] is not None and len(frames['Contacts']) > 0:
            sheet_order.append('Contacts')

    for extra in ['Too Small', 'Fixed Income', 'HFs', 'DNC', 'Check', 'Quant', 'Activist', 'Excluded']:
        if extra in frames and frames[extra] is not None and len(frames[extra]) > 0:
            sheet_order.append(extra)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name in sheet_order:
            df = frames[sheet_name]
            if df is not None and len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Ensure at least one sheet exists
        if not sheet_order:
            pd.DataFrame(columns=['No results']).to_excel(writer, sheet_name='Contacts', index=False)

    output.seek(0)
    wb = load_workbook(output)
    for sheet_name in wb.sheetnames:
        _format_sheet(wb[sheet_name])

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.read()
