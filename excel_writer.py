import tempfile
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

HEADER = ['Date', 'الصنف للسلطات\nItem for Salads', 'Qty by\nBox/Piece',
          'Qty Needed\n(GM)', 'Unit', 'Current\nInventory (GM)', 'Qty Received',
          'Rec. Unit', 'Invoice\nQty (كمية الفاتورة)', 'Invoice\nPrice SAR (سعر الفاتورة)']

HEADER_FILL = PatternFill('solid', start_color='1F4E78')
HEADER_FONT = Font(bold=True, color='FFFFFF')
SECTION_FILL = PatternFill('solid', start_color='DDE7F0')


def _sheet_name(date_obj):
    return date_obj.strftime('%B %Y')


def build_workbook(rows):
    """rows: list of dicts from the daily_items table (Supabase). Returns path to .xlsx file."""
    by_month = defaultdict(list)
    for r in rows:
        dt = datetime.strptime(r['item_date'], '%Y-%m-%d')
        by_month[_sheet_name(dt)].append((dt, r))

    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name in sorted(by_month, key=lambda s: datetime.strptime(s, '%B %Y')):
        ws = wb.create_sheet(sheet_name)
        ws.append(['OCTA FOOD — Daily Vegetables Order  |  طلبات الخضار اليومية  |  ' + sheet_name])
        ws.append([])
        ws.append(HEADER)
        for c in range(1, len(HEADER) + 1):
            cell = ws.cell(row=3, column=c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
        for col, width in zip('ABCDEFGHIJ', [12, 32, 10, 12, 8, 14, 12, 10, 14, 16]):
            ws.column_dimensions[col].width = width

        day_rows = sorted(by_month[sheet_name], key=lambda t: (t[0], t[1].get('section', ''), t[1].get('name_en', '')))
        current_section = None
        for dt, r in day_rows:
            if r.get('section') != current_section:
                current_section = r.get('section')
                label = 'السلطات' if current_section == 'salads' else ('الصوص' if current_section == 'dressing' else '')
                if label:
                    ws.append([label] + [''] * (len(HEADER) - 1))
                    for c in range(1, len(HEADER) + 1):
                        ws.cell(row=ws.max_row, column=c).fill = SECTION_FILL

            display_name = f"{r.get('name_ar') or ''} — {r.get('name_en') or ''}".strip(' —')
            ws.append([
                dt, display_name, r.get('qty_box'), r.get('qty_needed'), r.get('unit'),
                r.get('current_inventory'), r.get('qty_received'), r.get('rec_unit'),
                r.get('invoice_qty'), r.get('invoice_price'),
            ])
            ws.cell(row=ws.max_row, column=1).number_format = 'DD-MMM-YYYY'

    if not wb.sheetnames:
        wb.create_sheet('No Data')

    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(tmp.name)
    return tmp.name
