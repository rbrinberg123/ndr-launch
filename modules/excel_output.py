import io
import re
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

def sanitize_sheet_name(name):
    """Make a string safe for use as an Excel sheet name."""
    name = name.replace('/', '-')
    name = re.sub(r'[\\?*\[\]:]', ' ', name)
    name = re.sub(r' {2,}', ' ', name)
    name = name.strip()
    return name[:31]


NUMERIC_COLS = {
    'Shares', 'Fund Shares', 'Passive or Index Shares',
    'Total Funds', 'Passive or Index Funds', 'Match Count',
    'EAUM ($mm)', 'AUM ($mm)', 'T/O %',
    'L12M', 'Total', '3rd Party', 'Rose & Co',
}
DATE_COLS  = {'Specifically with Co.', 'Anyone at Inst. with Co', 'Last Meeting', 'As of'}
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
                    cell.number_format = '0.0%'
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


MCAP_DISPLAY = {
    '****Micro': 'Micro', '***Small': 'Small', '**Mid': 'Mid',
    '*Large': 'Large', 'Mega': 'Mega',
}

HF_LABELS = {
    'separate': 'Separate into HFs tab',
    'low_turnover': 'Include low-turnover HFs only (T/O ≤ 100%)',
    'include': 'Include all',
}

MEETING_LABELS = {
    'include_all': 'Include all contacts',
    'exclude_l12m': 'Exclude contacts met in last 12 months',
    'exclude_l24m': 'Exclude contacts met in last 24 months',
    'exclude_all': 'Exclude all contacts with prior meetings',
}

SECTION_FILL = PatternFill('solid', start_color='2E75B6')
SECTION_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=11)
LABEL_FONT   = Font(name='Arial', bold=True, size=10, color='1F3864')
VALUE_FONT   = Font(name='Arial', size=10)


def _build_criteria_sheet(wb, results, safe_names=None):
    ws = wb.create_sheet('Summary', 0)
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 60

    if safe_names is None:
        safe_names = {}

    criteria          = results.get('criteria') or {}
    city_selections   = results.get('city_selections') or []
    has_city          = results.get('has_city_routing', False)
    hf_treatment      = results.get('hf_treatment', 'separate')
    meeting_exclusion = results.get('meeting_exclusion', 'include_all')
    eaum_min          = results.get('eaum_min')
    subject_symbols   = results.get('subject_symbols') or []
    company_name      = results.get('company_name', 'Company')
    frames            = results.get('frames', {})

    row = 1

    def write_section(title):
        nonlocal row
        ws.cell(row=row, column=1, value=title).font = SECTION_FONT
        ws.cell(row=row, column=1).fill = SECTION_FILL
        ws.cell(row=row, column=2).fill = SECTION_FILL
        ws.row_dimensions[row].height = 22
        row += 1

    def write_row(label, value):
        nonlocal row
        c1 = ws.cell(row=row, column=1, value=label)
        c1.font = LABEL_FONT
        c1.alignment = LEFT
        c2 = ws.cell(row=row, column=2, value=value)
        c2.font = VALUE_FONT
        c2.alignment = LEFT
        row += 1

    def write_blank():
        nonlocal row
        row += 1

    # ── Company / Tickers
    write_section('Company / Tickers')
    write_row('Company Name', company_name)
    if subject_symbols:
        write_row('Tickers', ', '.join(subject_symbols))

    write_blank()

    # ── CDF Criteria
    write_section('CDF Criteria')
    industry = criteria.get('industry')
    style    = criteria.get('style')
    mcap     = criteria.get('mcap')
    geo      = criteria.get('geo')

    write_row('Industry Focus', ', '.join(sorted(industry)) if industry else '(all)')
    write_row('Investment Style', ', '.join(sorted(style)) if style else '(all)')
    if mcap:
        display_mcap = sorted(MCAP_DISPLAY.get(v, v) for v in mcap)
        write_row('Market Cap', ', '.join(display_mcap))
    else:
        write_row('Market Cap', '(all)')
    write_row('Geography', ', '.join(sorted(geo)) if geo else '(all)')

    write_blank()

    # ── Routing
    write_section('Routing')
    if has_city and city_selections:
        write_row('Mode', 'City routing')
        city_names = [name for name, _ in city_selections]
        write_row('Cities', ', '.join(city_names))
    else:
        write_row('Mode', 'Virtual (single list)')

    write_blank()

    # ── Exclusion Settings
    write_section('Exclusion Settings')
    write_row('Hedge Fund Treatment', HF_LABELS.get(hf_treatment, hf_treatment))
    write_row('Meeting Exclusion', MEETING_LABELS.get(meeting_exclusion, meeting_exclusion))
    if eaum_min is not None:
        write_row('EAUM Minimum', f'${eaum_min:,.0f}M')
    else:
        write_row('EAUM Minimum', '(none)')

    write_blank()

    # ── Sheet Summary
    write_section('Sheet Summary')
    excluded_keys = {'Too Small', 'HFs', 'DNC', 'Check', 'Quant', 'Activist', 'Excluded'}
    for k, v in frames.items():
        if v is not None and len(v) > 0:
            display = safe_names.get(k, k)
            label = f'{display} (excluded)' if k in excluded_keys else display
            write_row(label, f'{len(v):,} contacts')

    ws.sheet_properties.tabColor = '2E75B6'


def generate_excel(results, company_name):
    frames         = results['frames']
    city_selections = results.get('city_selections') or []
    has_city       = results.get('has_city_routing', False)

    output = io.BytesIO()

    # Determine sheet write order (original frame keys)
    sheet_order = []
    if has_city and city_selections:
        for tab_name, _ in city_selections:
            if tab_name in frames and frames[tab_name] is not None and len(frames[tab_name]) > 0:
                sheet_order.append(tab_name)
        if 'Virtual' in frames and frames['Virtual'] is not None and len(frames['Virtual']) > 0:
            sheet_order.append('Virtual')
    else:
        if 'Contacts' in frames and frames['Contacts'] is not None and len(frames['Contacts']) > 0:
            sheet_order.append('Contacts')

    for extra in ['HFs', 'DNC', 'Check', 'Quant', 'Activist', 'Excluded', 'Too Small']:
        if extra in frames and frames[extra] is not None and len(frames[extra]) > 0:
            sheet_order.append(extra)

    # Build mapping: original frame key → sanitized Excel sheet name (computed once)
    safe_names = {}
    used = set()
    for key in sheet_order:
        safe = sanitize_sheet_name(key)
        # Handle collisions after truncation
        base = safe
        n = 2
        while safe in used:
            suffix = f' {n}'
            safe = base[:31 - len(suffix)] + suffix
            n += 1
        used.add(safe)
        safe_names[key] = safe

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for key in sheet_order:
            df = frames[key]
            if df is not None and len(df) > 0:
                df.to_excel(writer, sheet_name=safe_names[key], index=False)

        # Ensure at least one sheet exists
        if not sheet_order:
            pd.DataFrame(columns=['No results']).to_excel(writer, sheet_name='Contacts', index=False)

    output.seek(0)
    wb = load_workbook(output)
    for sheet_name in wb.sheetnames:
        _format_sheet(wb[sheet_name])

    # Add Summary tab as the first sheet
    _build_criteria_sheet(wb, results, safe_names)

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.read()
