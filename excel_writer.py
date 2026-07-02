import tempfile
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# الأعمدة بالترتيب:
# A  التاريخ
# B  الصنف (عربي — إنجليزي)
# C  الوحدة
# D  كمية الأوردر  ← من الأوردر
# E  كمية الفاتورة ← من الفاتورة
# F  كمية المستلم  ← فاضي للإدخال اليدوي (مميّز بلون أخضر فاتح)
# G  الفرق (فاتورة - مستلم)  ← معادلة تلقائية
# H  السعر (SAR)
# I  المخزون الحالي
# ============================================================

HEADER = [
    'التاريخ\nDate',
    'الصنف\nItem',
    'الوحدة\nUnit',
    'كمية الأوردر\nOrder Qty',
    'كمية الفاتورة\nInvoice Qty',
    'كمية المستلم\nReceived Qty',
    'الفرق\nDiff',
    'سعر الفاتورة\nInvoice Price',
    'المخزون الحالي\nCurrent Inv.',
]

COL_WIDTHS = [13, 34, 9, 14, 15, 15, 10, 16, 15]

# ألوان
TITLE_FILL   = PatternFill('solid', start_color='EC1510')       # أحمر — عنوان
TITLE_FONT   = Font(bold=True, color='FFFFFF', size=13)

HEADER_FILL  = PatternFill('solid', start_color='1F4E78')       # أزرق داكن — هيدر
HEADER_FONT  = Font(bold=True, color='FFFFFF', size=10)

SECTION_FILL = PatternFill('solid', start_color='D6E4F0')       # أزرق فاتح — قسم
SECTION_FONT = Font(bold=True, color='1F4E78', size=10)

RECV_FILL    = PatternFill('solid', start_color='E8F5E9')       # أخضر فاتح — كمية المستلم
RECV_FONT    = Font(bold=True, color='1B5E20', size=10)
RECV_HEADER_FILL = PatternFill('solid', start_color='2E7D32')   # أخضر داكن — هيدر المستلم
RECV_HEADER_FONT = Font(bold=True, color='FFFFFF', size=10)

DIFF_POS_FILL = PatternFill('solid', start_color='FFF9C4')      # أصفر — فرق موجب
DIFF_NEG_FILL = PatternFill('solid', start_color='FFEBEE')      # أحمر فاتح — فرق سالب

EVEN_FILL = PatternFill('solid', start_color='F5F9FF')          # صفوف زوجية
ODD_FILL  = PatternFill('solid', start_color='FFFFFF')          # صفوف فردية

THIN  = Side(style='thin',   color='BDBDBD')
THICK = Side(style='medium', color='9E9E9E')
CELL_BORDER  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
RECV_BORDER  = Border(left=THICK, right=THICK, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
RIGHT  = Alignment(horizontal='right',  vertical='center')
LEFT   = Alignment(horizontal='left',   vertical='center')

# أعمدة الأرقام (C=3 فاكثر عدا B=2)
NUM_COLS = {4, 5, 6, 7, 8, 9}   # D,E,F,G,H,I


def _sheet_title_for_date(dt):
    return dt.strftime('%d-%b-%Y')


def _write_day_sheet(wb, dt, day_rows):
    ws = wb.create_sheet(_sheet_title_for_date(dt))
    n_cols = len(HEADER)
    last_col = get_column_letter(n_cols)

    # ===== صف العنوان =====
    ws.merge_cells(f'A1:{last_col}1')
    c = ws['A1']
    c.value = f"OCTA FOOD  —  طلبات الخضار اليومية  |  {dt.strftime('%d %B %Y')}"
    c.fill = TITLE_FILL; c.font = TITLE_FONT
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # ===== صف الهيدر =====
    ws.append(HEADER)
    hr = ws.max_row
    ws.row_dimensions[hr].height = 36
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=hr, column=col_idx)
        # عمود المستلم (F=6) بلون أخضر مميّز
        if col_idx == 6:
            cell.fill = RECV_HEADER_FILL
            cell.font = RECV_HEADER_FONT
        else:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = CELL_BORDER

    # عرض الأعمدة
    for i, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = 'A3'

    # ===== الصفوف =====
    row_counter = 0
    current_section = None

    for r in sorted(day_rows, key=lambda x: (x.get('section') or '', x.get('name_en') or '')):
        # صف القسم (سلطات / صوص)
        if r.get('section') != current_section:
            current_section = r.get('section')
            label = 'السلطات — Salads' if current_section == 'salads' else \
                    ('الصوص — Dressing' if current_section == 'dressing' else '')
            if label:
                ws.append(['', label] + [''] * (n_cols - 2))
                sr = ws.max_row
                ws.merge_cells(f'B{sr}:{last_col}{sr}')
                for col_idx in range(1, n_cols + 1):
                    sc = ws.cell(row=sr, column=col_idx)
                    sc.fill = SECTION_FILL
                    sc.font = SECTION_FONT
                    sc.border = CELL_BORDER
                ws.cell(row=sr, column=2).alignment = LEFT
                ws.row_dimensions[sr].height = 20

        # صف البيانات
        row_counter += 1
        base_fill = EVEN_FILL if row_counter % 2 == 0 else ODD_FILL

        name = f"{r.get('name_ar') or ''} — {r.get('name_en') or ''}".strip(' —')
        order_qty   = r.get('qty_needed')
        invoice_qty = r.get('invoice_qty')
        recv_qty    = r.get('qty_received')   # قد يكون None لو مفيش مستلم
        inv_price   = r.get('invoice_price')
        curr_inv    = r.get('current_inventory')

        ws.append([
            dt,           # A — التاريخ
            name,         # B — الصنف
            r.get('unit') or '',  # C — الوحدة
            order_qty,    # D — كمية الأوردر
            invoice_qty,  # E — كمية الفاتورة
            recv_qty,     # F — كمية المستلم (فاضي لو مفيش)
            None,         # G — الفرق (معادلة)
            inv_price,    # H — سعر الفاتورة
            curr_inv,     # I — المخزون الحالي
        ])
        dr = ws.max_row
        ws.row_dimensions[dr].height = 18
        ws.cell(row=dr, column=1).number_format = 'DD-MMM-YYYY'

        # تنسيق كل خلية في الصف
        for col_idx in range(1, n_cols + 1):
            cell = ws.cell(row=dr, column=col_idx)

            # عمود المستلم (F=6) — لون أخضر + border سميك
            if col_idx == 6:
                cell.fill = RECV_FILL
                cell.font = RECV_FONT
                cell.border = RECV_BORDER
                cell.alignment = CENTER
            else:
                cell.fill = base_fill
                cell.font = Font(size=10)
                cell.border = CELL_BORDER
                cell.alignment = CENTER if col_idx in NUM_COLS else LEFT

        # معادلة الفرق في عمود G (invoice - received)
        e_letter = get_column_letter(5)   # E
        f_letter = get_column_letter(6)   # F
        g_cell = ws.cell(row=dr, column=7)
        g_cell.value = f'=IF({f_letter}{dr}="","",{e_letter}{dr}-{f_letter}{dr})'
        g_cell.number_format = '#,##0.00'
        g_cell.alignment = CENTER

        # تنسيق الأرقام
        for col_idx in [4, 5, 6, 8, 9]:
            cell = ws.cell(row=dr, column=col_idx)
            if cell.value is not None:
                cell.number_format = '#,##0.00'

        # عمود A — التاريخ
        ws.cell(row=dr, column=2).alignment = LEFT

    # ===== صف التوتال =====
    ws.append([])
    tr = ws.max_row + 1
    ws.append([''] * n_cols)
    tr = ws.max_row

    total_fill = PatternFill('solid', start_color='1F4E78')
    total_font = Font(bold=True, color='FFFFFF', size=10)
    ws.cell(row=tr, column=2).value = 'الإجمالي — Total'
    ws.cell(row=tr, column=2).font = total_font
    ws.cell(row=tr, column=2).fill = total_fill
    ws.cell(row=tr, column=2).alignment = LEFT

    # SUM لكل عمود رقمي
    data_start = 3  # صف 3 = بعد العنوان والهيدر
    for col_idx in [4, 5, 6, 8]:
        col_l = get_column_letter(col_idx)
        cell = ws.cell(row=tr, column=col_idx)
        cell.value = f'=SUM({col_l}{data_start}:{col_l}{tr-1})'
        cell.font = total_font
        cell.fill = total_fill if col_idx != 6 else PatternFill('solid', start_color='2E7D32')
        cell.number_format = '#,##0.00'
        cell.alignment = CENTER
        cell.border = CELL_BORDER

    for col_idx in [1, 3, 7, 9]:
        cell = ws.cell(row=tr, column=col_idx)
        cell.fill = total_fill
        cell.border = CELL_BORDER

    ws.row_dimensions[tr].height = 22
    return ws


def build_workbook(rows):
    by_day = defaultdict(list)
    for r in rows:
        raw_date = r.get('item_date')
        if not raw_date:
            continue
        try:
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
