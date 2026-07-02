"""
Octa Food - Daily Excel Builder
================================
يقرأ ملفات PDF لـ:
  1) طلبات الخضار اليومية (Order sheets)  -> المصنف: الكمية المطلوبة + المخزون الحالي
  2) الفواتير الضريبية (Invoice PDFs)      -> المصنف: كمية وسعر الفاتورة الفعلية

ويجمعهم في ملف إكسيل واحد، تاب منفصل لكل يوم، بالتنسيق المتفق عليه.

الاستخدام:
    python octa_excel_builder.py --orders order1.pdf order2.pdf --invoices inv1.pdf inv2.pdf --out final.xlsx
"""
import re
import subprocess
import unicodedata
import difflib
from collections import defaultdict

import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# 1) قراءة ملف الأوردر (طلبات الخضار اليومية)
# ---------------------------------------------------------------------------
def parse_order_pdf(path):
    """Returns (date_str, [items]) - items include en/ar name, unit, needed qty (g/ml),
    box_qty (if tracked in boxes), current_inventory."""
    txt = subprocess.run(['pdftotext', '-layout', path, '-'],
                          capture_output=True, encoding='utf-8').stdout

    date_match = re.search(r'\d{1,2}-[A-Za-z]{3}-\d{4}', txt)
    date_str = date_match.group() if date_match else None

    items = []
    for line in txt.split('\n'):
        if '—' not in line:
            continue
        stripped_check = line.replace('\u202b', '').replace('\u202c', '')
        if 'Gram' not in stripped_check and 'ML-' not in stripped_check:
            continue
        cols = re.split(r'\s{2,}', line.strip())
        if len(cols) < 3:
            continue
        m = re.match(r'^(.*?)\s*—\s*\u202b?(.*?)\u202c?$', cols[0])
        if not m:
            continue
        en_name, ar_name = m.group(1).strip(), m.group(2).strip()

        unit_idx = next((i for i, c in enumerate(cols)
                          if c.startswith('Gram') or c.startswith('ML-')), None)
        if unit_idx is None:
            continue

        nums_before = cols[1:unit_idx]
        nums_after = cols[unit_idx + 1:]
        box_qty = needed_qty = None
        if len(nums_before) == 2:
            box_qty, needed_qty = nums_before
        elif len(nums_before) == 1:
            needed_qty = nums_before[0]
        current_inv = nums_after[0] if nums_after else None
        unit = 'Gram' if cols[unit_idx].startswith('Gram') else 'ML'

        items.append(dict(
            en=en_name, ar=ar_name, unit=unit,
            box_qty=float(box_qty) if box_qty not in (None, '') else None,
            needed=float(needed_qty) if needed_qty not in (None, '') else 0.0,
            current_inv=float(current_inv) if current_inv not in (None, '') else None,
        ))
    return date_str, items


# ---------------------------------------------------------------------------
# 2) قراءة الفاتورة الضريبية
# ---------------------------------------------------------------------------
INVOICE_LINE_RE = re.compile(
    r'(?P<row>\d{1,2})(?P<item>[^\d]+?)(?P<price>\d+\.\d+)(?P<unit>كج|علبه|قطعة)\s*(?P<qty>\d+(?:\.\d+)?)\s*'
    r'\n\s*(?P<discpct>[\d.]+)%\s*\n\s*(?P<discamt>[\d.]+)\s*\n\s*(?P<subtotal>[\d.]+)\s*'
    r'\n\(\s*(?P<taxamt>[\d.]+)\s*\n\s*%\s*\)\s*\n15\s*\n\s*(?P<total>[\d.]+)'
)


def parse_invoice_pdf(path):
    doc = fitz.open(path)
    full_text = "".join(p.get_text() + "\n" for p in doc)
    inv_no = re.search(r'INV\d+', full_text)
    inv_date = re.search(r'\d{4}-\d{2}-\d{2}', full_text)
    text = unicodedata.normalize('NFKC', full_text.replace('\u200a', ' ').replace('\ue900', ''))

    rows = []
    for m in INVOICE_LINE_RE.finditer(text):
        d = m.groupdict()
        item = re.sub(r'\s+', ' ', d['item']).strip()
        rows.append(dict(item_raw=item, price=float(d['price']), unit=d['unit'],
                          qty=float(d['qty']), subtotal=float(d['subtotal']), total=float(d['total'])))
    return dict(invoice_no=inv_no.group() if inv_no else None,
                date=inv_date.group() if inv_date else None,
                rows=rows)


# ---------------------------------------------------------------------------
# 3) مطابقة أصناف الفاتورة بأصناف الأوردر (Arabic fuzzy matching)
# ---------------------------------------------------------------------------
def _norm_ar(s):
    s = re.sub(r'[\u064B-\u0652]', '', s)
    s = (s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
           .replace('ى', 'ي').replace('ة', 'ه'))
    s = re.sub(r'[^\u0600-\u06FF\s]', '', s)
    return s.strip()


def _token_set(s):
    return set(_norm_ar(s).split())


def match_item(invoice_ar, catalog):
    inv_tokens = _token_set(invoice_ar)
    inv_joined = _norm_ar(invoice_ar).replace(' ', '')
    best, best_key = None, (-1, -1.0)
    for c in catalog:
        cat_tokens = _token_set(c['ar'])
        if not cat_tokens or not inv_tokens:
            continue
        overlap = len(inv_tokens & cat_tokens)
        if overlap == 0:
            continue
        jaccard = overlap / len(inv_tokens | cat_tokens)
        cat_joined = _norm_ar(c['ar']).replace(' ', '')
        ratio = difflib.SequenceMatcher(None, inv_joined, cat_joined).ratio()
        key = (overlap, jaccard + ratio * 0.01)
        if key > best_key:
            best_key = key
            best = c
    return best


def invoice_qty_to_grams(inv_row, order_item):
    """Convert invoice quantity to grams using the order sheet's own box->gram ratio
    when the invoice is billed by carton/box; kg is a straight x1000."""
    if inv_row['unit'] == 'كج':
        return inv_row['qty'] * 1000.0
    if inv_row['unit'] in ('علبه', 'قطعة'):
        box_qty = order_item.get('box_qty')
        needed = order_item.get('needed')
        if box_qty and needed and box_qty > 0:
            grams_per_box = needed / box_qty
            return inv_row['qty'] * grams_per_box
        return None  # can't convert without a ratio
    return None


# ---------------------------------------------------------------------------
# 4) بناء الإكسيل
# ---------------------------------------------------------------------------
HEADER_BLUE = "1F4E78"
HEADER_GREEN = "2E7D32"
FILL_GREEN_LIGHT = "E2F0D9"
FILL_ROW_ALT = "DCE6F1"
FILL_ROW_WHITE = "FFFFFF"

HEADERS = ["التاريخ", "الصنف", "الوحدة", "كمية الأوردر", "كمية الفاتورة",
           "كمية المستلم", "الفرق", "سعر الفاتورة", "المخزون الحالي"]


def build_workbook(day_records):
    """day_records: dict[date_str] -> list of row dicts with keys:
       date, item, unit, order_qty, invoice_qty, price, current_inv"""
    wb = Workbook()
    wb.remove(wb.active)

    thin = Side(style='thin', color='B7B7B7')
    thick = Side(style='medium', color='1B5E20')

    for date_str, rows in day_records.items():
        sheet_name = date_str[:31] if date_str else "Sheet"
        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.rightToLeft = True

        for col, htext in enumerate(HEADERS, start=1):
            cell = ws.cell(row=1, column=col, value=htext)
            cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if col == 6:  # كمية المستلم
                cell.fill = PatternFill('solid', start_color=HEADER_GREEN)
                cell.border = Border(top=thick, bottom=thick, left=thick, right=thick)
            else:
                cell.fill = PatternFill('solid', start_color=HEADER_BLUE)
        ws.row_dimensions[1].height = 30

        r = 2
        for row in rows:
            band = FILL_ROW_WHITE if (r % 2 == 0) else FILL_ROW_ALT
            values = [row['date'], row['item'], row['unit'], row['order_qty'],
                      row['invoice_qty'], None, None, row['price'], row['current_inv']]
            for col, val in enumerate(values, start=1):
                cell = ws.cell(row=r, column=col, value=val)
                cell.font = Font(name="Arial", size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if col == 6:
                    cell.fill = PatternFill('solid', start_color=FILL_GREEN_LIGHT)
                    cell.border = Border(top=thick, bottom=thick, left=thick, right=thick)
                else:
                    cell.fill = PatternFill('solid', start_color=band)
            # الفرق = كمية الفاتورة - كمية المستلم
            ws.cell(row=r, column=7, value=f"=E{r}-F{r}")
            ws.cell(row=r, column=7).font = Font(name="Arial", size=10)
            ws.cell(row=r, column=7).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=7).fill = PatternFill('solid', start_color=band)
            r += 1

        last_data_row = r - 1
        total_row = r
        ws.cell(row=total_row, column=2, value="الإجمالي").font = Font(bold=True, name="Arial")
        for col in (4, 5, 6, 7):
            letter = get_column_letter(col)
            c = ws.cell(row=total_row, column=col,
                        value=f"=SUM({letter}2:{letter}{last_data_row})")
            c.font = Font(bold=True, name="Arial")
            c.fill = PatternFill('solid', start_color="FFE699")
        for col in range(1, 10):
            ws.cell(row=total_row, column=col).border = Border(top=Side(style='double'))

        widths = [12, 26, 9, 13, 13, 13, 10, 13, 14]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    return wb


# ---------------------------------------------------------------------------
# 5) تجميع الملفات (main pipeline)
# ---------------------------------------------------------------------------
def build_from_files(order_paths, invoice_paths, out_path):
    day_catalog = {}   # date_str -> order items list
    for p in order_paths:
        date_str, items = parse_order_pdf(p)
        day_catalog[date_str] = items

    invoices_by_date = defaultdict(list)
    for p in invoice_paths:
        inv = parse_invoice_pdf(p)
        invoices_by_date[inv['date']].append(inv)

    day_records = {}
    for date_str, order_items in day_catalog.items():
        # find matching invoice(s) by date (order date format DD-Mon-YYYY vs invoice YYYY-MM-DD)
        norm_order_date = None
        try:
            import datetime
            norm_order_date = datetime.datetime.strptime(date_str, "%d-%b-%Y").strftime("%Y-%m-%d")
        except Exception:
            pass

        matched_invoice_rows = []
        for inv_date, invs in invoices_by_date.items():
            if inv_date == norm_order_date:
                for inv in invs:
                    matched_invoice_rows.extend(inv['rows'])

        # Match each invoice line ONCE against the full catalog (best global match),
        # then key the result by the matched order item's identity.
        invoice_by_order_key = {}
        for ir in matched_invoice_rows:
            m = match_item(ir['item_raw'], order_items)
            if m is None:
                continue
            key = (m['en'], m['ar'])
            grams = invoice_qty_to_grams(ir, m)
            # if two invoice lines match the same catalog item, sum them
            if key in invoice_by_order_key:
                prev = invoice_by_order_key[key]
                prev_grams = prev[0] or 0
                invoice_by_order_key[key] = ((grams or 0) + prev_grams, ir['price'])
            else:
                invoice_by_order_key[key] = (grams, ir['price'])

        rows_out = []
        for oi in order_items:
            key = (oi['en'], oi['ar'])
            best_grams, best_price = invoice_by_order_key.get(key, (None, None))
            rows_out.append(dict(
                date=date_str,
                item=f"{oi['en']} — {oi['ar']}",
                unit="GM" if oi['unit'] == 'Gram' else "ML",
                order_qty=oi['needed'],
                invoice_qty=best_grams,
                price=best_price,
                current_inv=oi['current_inv'],
            ))
        day_records[date_str] = rows_out

    wb = build_workbook(day_records)
    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('--orders', nargs='+', required=True)
    ap.add_argument('--invoices', nargs='+', required=True)
    ap.add_argument('--out', default='octa_food_output.xlsx')
    args = ap.parse_args()
    build_from_files(args.orders, args.invoices, args.out)
    print("Saved:", args.out)
