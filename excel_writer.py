import tempfile
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HEADER = ['Date', 'الصنف للسلطات\nItem for Salads', 'Qty by\nBox/Piece',
          'Qty Needed\n(GM)', 'Unit', 'Current\nInventory (GM)', 'Qty Received',
          'Rec. Unit', 'Invoice\nQty (كمية الفاتورة)', 'Invoice\nPrice SAR (سعر الفاتورة)']

TITLE_FILL = PatternFill('solid', start_color='EC1510')
TITLE_FONT = Font(bold=True, color='FFFFFF', size=14)
HEADER_FILL = PatternFill('solid', start_color='1F4E78')
HEADER_FONT = Font(bold=True, color='FFFFFF')
SECTION_FILL = PatternFill('solid', start_color='DDE7F0')
SECTION_FONT = Font(bold=True, color='1F4E78')
COL_WIDTHS = [12, 32, 10, 12, 8, 14, 12, 10, 14, 16]


def _sheet_title_for_date(dt):
    """Sheet/tab name for a single day, e.g. '01-Apr-2026' (Excel forbids / \\ ? * [ ] : in names)."""
    return dt.strftime('%d-%b-%Y')


def _write_day_sheet(wb, dt, day_rows):
    sheet_name = _sheet_title_for_date(dt)
    ws = wb.create_sheet(sheet_name)

    n_cols = len(HEADER)
    last_col_letter = get_column_letter(n_cols)

    # Big, clearly visible title bar across the top of the sheet
    title_text = f"OCTA FOOD — Daily Vegetables Order  |  طلبات الخضار اليومية  |  {dt.strftime('%d %B %Y')}"
    ws.merge_cells(f'A1:{last_col_letter}1')
    title_cell = ws['A1']
    title_cell.value = title_text
    title_cell.fill = TITLE_FILL
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.append([])  # spacer row
    ws.append(HEADER)
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=3, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    ws.row_dimensions[3].height = 32

    for col, width in zip('ABCDEFGHIJ', COL_WIDTHS):
        ws.column_dimensions[col].width = width

    ws.freeze_panes = 'A4'  # keep title + header visible while scrolling

    current_section = None
    for r in sorted(day_rows, key=lambda x: (x.get('section') or '', x.get('name_en') or '')):
        if r.get('section') != current_section:
            current_section = r.get('section')
            label = 'السلطات' if current_section == 'salads' else ('الصوص' if current_section == 'dressing' else '')
            if label:
                ws.append([label] + [''] * (n_cols - 1))
                for c in range(1, n_cols + 1):
                    cell = ws.cell(row=ws.max_row, column=c)
                    cell.fill = SECTION_FILL
                    cell.font = SECTION_FONT

        display_name = f"{r.get('name_ar') or ''} — {r.get('name_en') or ''}".strip(' —')
        ws.append([
            dt, display_name, r.get('qty_box'), r.get('qty_needed'), r.get('unit'),
            r.get('current_inventory'), r.get('qty_received'), r.get('rec_unit'),
            r.get('invoice_qty'), r.get('invoice_price'),
        ])
        ws.cell(row=ws.max_row, column=1).number_format = 'DD-MMM-YYYY'

    return ws


def build_workbook(rows):
    """rows: list of dicts from the daily_items table (Supabase).
    Creates ONE separate sheet/tab per calendar day (not stacked under a single
    monthly tab), sorted chronologically left-to-right. Returns path to .xlsx file.
    """
    by_day = defaultdict(list)
    for r in rows:
        raw_date = r.get('item_date')
        if not raw_date:
            continue
        try:
            # Supabase ممكن ترجع التاريخ كـ 'YYYY-MM-DD' (date) أو كـ
            # 'YYYY-MM-DDTHH:MM:SS+00:00' (timestamp) — نتعامل مع الاتنين
            dt = datetime.strptime(str(raw_date)[:10], '%Y-%m-%d')
        except ValueError:
            continue
        by_day[dt].append(r)

    wb = Workbook()
    wb.remove(wb.active)

    for dt in sorted(by_day.keys()):
        _write_day_sheet(wb, dt, by_day[dt])

    if not wb.sheetnames:
        wb.create_sheet('No Data')

    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(tmp.name)
    return tmp.name
