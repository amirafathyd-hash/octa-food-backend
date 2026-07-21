"""
محطة التنقية من التقارير إلى اتخاذ القرار
==========================================

الفكرة: بترفع "فاتورة الكمية للمشتركين" الخام (نفس شيت Export اللي فيه
عمود لكل صنف/باقة/كمية) وبيرجع لك ملف بنفس شكل وترتيب ملف اليوم الجاهز
(زي "Octa Food Tue ... الثلاثاء.xlsx"): شيت Export فيه بياناتك المرفوعة
زي ما هي، شيت "Don't Use just refresh" فيه تفصيل كل صنف لـ Protein/Side
وربطه بالباقة النهائية، وشيت Update فيه الجدول المُجمّع (زي الـ Pivot
Table بالظبط): كل صنف (بروتين أو طبق جانبي) في صف، وكل باقة نهائية في
عمود بعدد ووزن، + إجمالي لكل صف وصف إجمالي كلي في الآخر.

مهم جدًا: الحساب هنا مبني على قاعدة عمل مؤكدة مع صاحب المشروع بعد فحص
ملف حقيقي (يوم الثلاثاء 21-7-2026) صف بصف:
  - كل صنف إنجليزي بيتقسم لـ Protein (والـ Side لو موجود) - قاموس ثابت.
  - كل باقة أصلية (زي "لايت دايت") بتتحول لتصنيف نهائي (زي "تكميم لايت")
    عبر شيت Packages.
  - 3 أصناف بس (ساندوتش البيض المسلوق / كلوب ساندوتش بالدجاج / ساندوتش
    الدجاج المشوي) بتتضاعف ×2 حصريًا لما تكون الباقة النهائية "تضخيم".
  - أي صنف أو باقة مش موجودين في القاموس (data/decision_station_lookup.json)
    بيوقف التشغيل برسالة واضحة، بدل ما يتجاهلهم ويطلع رقم إنتاج غلط.

القاموس بيتوسع بمرور الوقت: أول ما يظهر صنف جديد (يوم تاني غير الثلاثاء)
هيوقف التشغيل ويقولك بالظبط الاسم الناقص عشان تضيفه لملف الـ JSON.
"""
import json
import os
import re
import tempfile
from collections import OrderedDict
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side as BorderSide
from openpyxl.utils import get_column_letter

LOOKUP_PATH = os.path.join(os.path.dirname(__file__), 'data', 'decision_station_lookup.json')

DAY_NAMES_BY_WEEKDAY = {
    5: 'السبت',    # Saturday
    6: 'الأحد',    # Sunday
    0: 'الاثنين',  # Monday
    1: 'الثلاثاء', # Tuesday
    2: 'الأربعاء', # Wednesday
    3: 'الخميس',   # Thursday
    4: 'الجمعة',   # Friday (نادرًا ما يستخدم فعليًا)
}

PREFERRED_PACKAGE_ORDER = ['تضخيم', 'تكميم لايت', 'جيم', 'سمارت دايت', 'غذاء العمل']

RAW_REQUIRED_HEADERS = {
    '#', 'الاسم الإنجليزي', 'الاسم العربي', 'الباقة', 'التصنيف',
    'مجموع الكارب', 'مجموع البروتين', 'الكمية',
}

EXPORT_COLUMNS = [
    '#', 'الاسم الإنجليزي', 'الاسم العربي', 'الباقة', 'المرجع',
    'التصنيف', 'الحجم', 'مجموع الكارب', 'مجموع البروتين', 'الكمية',
]

DONT_USE_COLUMNS = [
    'الاسم الإنجليزي', 'Protein', 'Side', 'الباقة', 'التصنيف',
    'مجموع الكارب', 'مجموع البروتين', 'الكمية', 'Total_GM',
    'Final_Package', 'Multiplier', 'Final_Count', 'Final_GM', 'Grams',
]


def load_lookup():
    with open(LOOKUP_PATH, encoding='utf-8') as fh:
        return json.load(fh)


def _category_rank(category):
    cat = category or ''
    if 'فطور' in cat:
        return 0
    if 'الوجبات الرئيسية' in cat or 'لو كارب' in cat:
        return 1
    if 'سلطات' in cat or 'فواكه' in cat or 'الإضافات' in cat:
        return 2
    if 'حلى' in cat or 'حلويات' in cat:
        return 3
    return 4


def _number(value):
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _find_export_sheet(wb):
    """بيدور على أي شيت فيه أعمدة فاتورة المشتركين الخام (زي شيت Sheet1
    أو Export)، بغض النظر عن اسم الشيت."""
    for ws in wb.worksheets:
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        headers = {
            str(val or '').strip(): idx
            for idx, val in enumerate(first_row, start=1)
        }
        if RAW_REQUIRED_HEADERS.issubset(headers):
            return ws, headers
    return None, None


def read_subscribers_invoice(file_storage):
    """بيقرا فاتورة المشتركين الخام ويرجع (rows, day_label).
    rows: list[dict] بنفس أعمدة EXPORT_COLUMNS.
    day_label: اسم اليوم بالعربي، مستنتج من تاريخ في اسم الملف لو موجود."""
    file_storage.seek(0)
    # read_only=False لأن القراءة العشوائية بالخلية (ws.cell) في وضع read_only
    # بطيئة/غير مضمونة على ملفات الشيتات الكبيرة زي دي.
    wb = load_workbook(file_storage, data_only=True, read_only=False)
    ws, headers = _find_export_sheet(wb)
    if ws is None:
        raise ValueError(
            "الملف المرفوع لازم يحتوي على أعمدة فاتورة المشتركين "
            "(الاسم الإنجليزي / الباقة / التصنيف / الكمية ... إلخ)"
        )

    col_index = {col: headers[col] for col in EXPORT_COLUMNS if col in headers}
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        english_idx = col_index.get('الاسم الإنجليزي')
        if english_idx is None or english_idx - 1 >= len(values):
            continue
        english = values[english_idx - 1]
        if not english or not str(english).strip():
            continue
        row = {}
        for col, idx in col_index.items():
            row[col] = values[idx - 1] if idx - 1 < len(values) else None
        for col in EXPORT_COLUMNS:
            row.setdefault(col, None)
        rows.append(row)
    wb.close()
    file_storage.seek(0)

    upload_name = (
        getattr(file_storage, 'filename', '') or
        os.path.basename(getattr(file_storage, 'name', '') or '')
    )
    day_label = _day_label_from_filename(upload_name)
    return rows, day_label


def _day_label_from_filename(filename):
    match = re.search(r'(20\d{2})[-_/](\d{1,2})[-_/](\d{1,2})', filename or '')
    if not match:
        return None
    date_value = datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    return DAY_NAMES_BY_WEEKDAY.get(date_value.weekday())


def compute_decision_tables(rows, lookup):
    """بياخد صفوف فاتورة المشتركين الخام ويرجع:
    dont_use_rows, pivot_rows, package_order, report

    شكل جدول Update الحقيقي مش عبارة عن صف واحد لكل اسم صنف بس - هو تجميع
    على مستويين زي أي Pivot Table فيها Row field مركّب:
      1) صف "Protein" بيجمع كل الكميات لنفس البروتين (زي "دجاج بالفطر")
         مهما اختلفت الأطباق الجانبية أو الباقات.
      2) تحته مباشرة، صف منفصل لكل طبق جانبي حقيقي (Side != '-') ظهر مع
         البروتين ده بالذات - حتى لو نفس اسم الطبق الجانبي ("خضار سوتيه"
         مثلاً) ظهر تحت بروتينات تانية، هيبقى له صف منفصل لكل بروتين
         (بالظبط زي ما بيحصل في ملف الإكسل الحقيقي - راجعنا ده صف بصف).
      3) لو الصنف مالوش طبق جانبي حقيقي (Side == '-')، مفيش صف تاني
         بيتضاف تحته - صف البروتين نفسه كافي.

    pivot_rows: list[(display_name, {final_package: [count, grams]})]
    بنفس ترتيب الظهور الحقيقي في ملف اليوم الجاهز."""
    items_map = lookup['items']
    package_map = lookup['package_map']
    double_items = set(lookup.get('double_in_bulking', []))
    bulking_pkg = lookup.get('bulking_final_package', 'تضخيم')

    unmapped_items = set()
    unmapped_packages = set()
    dont_use_rows = []
    package_order = []

    protein_order = []
    protein_category = {}  # protein -> أول تصنيف اتشاف بيه (لترتيب البلوكات زي الملف الأصلي)
    protein_data = {}  # protein -> {'totals': OrderedDict, 'side_order': [], 'sides': {side: OrderedDict}}

    for row in rows:
        english = str(row.get('الاسم الإنجليزي') or '').strip()
        qty = _number(row.get('الكمية'))
        if not english or qty <= 0:
            continue
        item_info = items_map.get(english)
        orig_pkg = str(row.get('الباقة') or '').strip()
        final_pkg = package_map.get(orig_pkg)

        if item_info is None:
            unmapped_items.add(english)
        if not final_pkg:
            unmapped_packages.add(orig_pkg or '(فارغ)')
        if item_info is None or not final_pkg:
            continue

        protein = item_info.get('protein')
        side = item_info.get('side')  # None لو مفيش طبق جانبي حقيقي
        category = row.get('التصنيف') or item_info.get('category') or ''
        carb = _number(row.get('مجموع الكارب'))
        protein_g = _number(row.get('مجموع البروتين'))
        base_gm = carb or protein_g
        total_gm = base_gm if base_gm else qty

        multiplier = 2 if (english in double_items and final_pkg == bulking_pkg) else 1
        final_count = qty * multiplier
        final_gm = total_gm * multiplier

        dont_use_rows.append({
            'الاسم الإنجليزي': english,
            'Protein': protein,
            'Side': side or '-',
            'الباقة': orig_pkg,
            'التصنيف': category,
            'مجموع الكارب': carb,
            'مجموع البروتين': protein_g,
            'الكمية': qty,
            'Total_GM': total_gm,
            'Final_Package': final_pkg,
            'Multiplier': multiplier,
            'Final_Count': final_count,
            'Final_GM': final_gm,
            'Grams': final_gm,
        })

        if final_pkg not in package_order:
            package_order.append(final_pkg)

        if protein:
            if protein not in protein_data:
                protein_data[protein] = {'totals': OrderedDict(), 'side_order': [], 'sides': {}}
                protein_order.append(protein)
                protein_category[protein] = category
            pd = protein_data[protein]
            bucket = pd['totals'].setdefault(final_pkg, [0.0, 0.0])
            bucket[0] += final_count
            bucket[1] += final_gm

            if side:
                if side not in pd['sides']:
                    pd['sides'][side] = OrderedDict()
                    pd['side_order'].append(side)
                sbucket = pd['sides'][side].setdefault(final_pkg, [0.0, 0.0])
                sbucket[0] += final_count
                sbucket[1] += final_gm

    if unmapped_items or unmapped_packages:
        parts = []
        if unmapped_items:
            parts.append('أصناف غير معروفة: ' + '، '.join(sorted(unmapped_items)))
        if unmapped_packages:
            parts.append('باقات غير معروفة: ' + '، '.join(sorted(unmapped_packages)))
        raise ValueError(
            'محطة التنقية محتاجة تحديث القاموس قبل ما تكمل (عشان الأرقام '
            'تطلع صح دايمًا): ' + ' | '.join(parts)
        )

    # pivot_rows: (display_name, totals_by_package, is_protein_level)
    # is_protein_level=True بس للصف الأب (البروتين) - ده اللي بيدخل في
    # حساب Grand Total، أما صفوف الطبق الجانبي فهي تفصيل تحت صف البروتين
    # (نفس قيمه جزئيًا) ومينفعش تتجمع تاني في الإجمالي الكلي وإلا
    # هيتضاعف الرقم.
    # نفس ترتيب "البلوكات" في ملف اليوم الجاهز: فطور، بعدين الوجبات
    # الرئيسية (وبنودها اللي بلو كارب)، بعدين سلطات/فواكه/إضافات، بعدين
    # حلى، وأي تصنيف تاني مش معروف يتحط في الآخر. الترتيب جوه كل بلوك
    # زي ترتيب أول ظهور في الملف المرفوع نفسه.
    ordered_proteins = sorted(
        protein_order,
        key=lambda p: (_category_rank(protein_category.get(p, '')), protein_order.index(p)),
    )

    pivot_rows = []
    for protein in ordered_proteins:
        pd = protein_data[protein]
        pivot_rows.append((protein, pd['totals'], True))
        for side in pd['side_order']:
            pivot_rows.append((side, pd['sides'][side], False))

    ordered_packages = [p for p in PREFERRED_PACKAGE_ORDER if p in package_order]
    ordered_packages += [p for p in package_order if p not in ordered_packages]

    report = {
        'source_rows': len(rows),
        'computed_rows': len(dont_use_rows),
        'row_labels': len(pivot_rows),
        'package_columns': ordered_packages,
    }
    return dont_use_rows, pivot_rows, ordered_packages, report


# ---------------------------------------------------------------------------
# بناء ملف الإخراج بنفس شكل ملف اليوم الجاهز (Export / Don't Use / Update / Packages)
# ---------------------------------------------------------------------------

THIN = BorderSide(style='thin', color='DDDDDD')
BORDER_THIN = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def _plain_header_row(ws, columns, bold=False):
    for idx, col in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=idx, value=col)
        cell.font = Font(bold=bold, size=11)


def _category_fill(category, lookup):
    palette = lookup.get('category_colors') or {}
    cat = category or ''
    key = None
    if 'فطور' in cat:
        key = 'فطور'
    elif 'حلى' in cat or 'حلويات' in cat:
        key = 'حلى'
    elif 'سلطات' in cat or 'فواكه' in cat or 'الإضافات' in cat:
        key = 'سلطات'
    elif 'الوجبات الرئيسية' in cat or 'لو كارب' in cat:
        key = 'الوجبات الرئيسية'
    spec = palette.get(key) or palette.get('__default__') or {'theme': 4, 'tint': 0.8}
    if 'rgb' in spec:
        from openpyxl.styles.colors import Color
        return PatternFill(patternType='solid', fgColor=Color(rgb=spec['rgb']))
    from openpyxl.styles.colors import Color
    return PatternFill(patternType='solid', fgColor=Color(theme=spec.get('theme', 4), tint=spec.get('tint', 0.8)))


def build_output_workbook(day_label, export_rows, dont_use_rows, pivot_rows,
                           package_order, lookup, out_path=None):
    wb = Workbook()

    # ---------------- Export ----------------
    ws_export = wb.active
    ws_export.title = 'Export'
    _plain_header_row(ws_export, EXPORT_COLUMNS)
    for r, row in enumerate(export_rows, start=2):
        for c, col in enumerate(EXPORT_COLUMNS, start=1):
            ws_export.cell(row=r, column=c, value=row.get(col))
    widths = [6, 32, 30, 26, 10, 16, 34, 12, 12, 9]
    for i, w in enumerate(widths, start=1):
        ws_export.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Don't Use just refresh ----------------
    ws_du = wb.create_sheet("Don't Use just refresh")
    _plain_header_row(ws_du, DONT_USE_COLUMNS)
    for r, row in enumerate(dont_use_rows, start=2):
        for c, col in enumerate(DONT_USE_COLUMNS, start=1):
            ws_du.cell(row=r, column=c, value=row.get(col))
    du_widths = [40, 24, 16, 18, 24, 12, 12, 9, 10, 14, 10, 12, 10, 9]
    for i, w in enumerate(du_widths, start=1):
        ws_du.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Packages ----------------
    ws_pkg = wb.create_sheet('Packages')
    header_fill = PatternFill(patternType='solid', fgColor='FFB7C6E8')
    for i, col in enumerate(['Original_Package', 'Final_Package'], start=1):
        cell = ws_pkg.cell(row=1, column=i, value=col)
        cell.font = Font(bold=True, size=11)
    for r, (orig, final) in enumerate(sorted(lookup['package_map'].items()), start=2):
        ws_pkg.cell(row=r, column=1, value=orig)
        ws_pkg.cell(row=r, column=2, value=final)
    ws_pkg.column_dimensions['A'].width = 30
    ws_pkg.column_dimensions['B'].width = 26

    # ---------------- Update (نفس شكل الـ Pivot تمامًا، بس قيم ثابتة) ----------------
    ws_up = wb.create_sheet('Update')
    header_font = Font(bold=True, size=18)
    plain_font = Font(bold=False, size=18)
    center = Alignment(horizontal='center', vertical='center')

    a6 = ws_up.cell(row=6, column=1, value=day_label or '')
    a6.font = Font(bold=True, size=18)
    a6.fill = PatternFill(patternType='solid', fgColor='FFFFFF00')
    a6.alignment = center

    ws_up.cell(row=7, column=2, value='Column Labels').font = plain_font

    n_pkg = len(package_order)
    total_count_col = 2 + n_pkg * 2
    total_gm_col = total_count_col + 1

    for i, pkg in enumerate(package_order):
        col = 2 + i * 2
        c1 = ws_up.cell(row=8, column=col, value=pkg)
        c1.font = header_font
        c1.alignment = center
        ws_up.cell(row=8, column=col + 1).font = header_font
        ws_up.cell(row=9, column=col, value='Count').font = plain_font
        ws_up.cell(row=9, column=col, value='Count').alignment = center
        ws_up.cell(row=9, column=col + 1, value='Grams').font = plain_font
        ws_up.cell(row=9, column=col + 1).alignment = center

    ws_up.cell(row=8, column=total_count_col, value='Total Count').font = header_font
    ws_up.cell(row=8, column=total_gm_col, value='Total Grams').font = header_font
    ws_up.cell(row=9, column=1, value='Row Labels').font = header_font
    ws_up.cell(row=9, column=1).alignment = center

    r = 10
    grand = [0.0] * (n_pkg * 2)
    for name, bucket, is_protein_level in pivot_rows:
        category = lookup.get('name_category', {}).get(name, '')
        fill = _category_fill(category, lookup)
        name_cell = ws_up.cell(row=r, column=1, value=name)
        name_cell.font = header_font
        name_cell.alignment = center
        name_cell.fill = fill
        row_total_count = 0.0
        row_total_gm = 0.0
        for i, pkg in enumerate(package_order):
            count, gm = bucket.get(pkg, [0.0, 0.0])
            col = 2 + i * 2
            cc = ws_up.cell(row=r, column=col, value=(count or None))
            gc = ws_up.cell(row=r, column=col + 1, value=(gm or None))
            cc.font = header_font
            gc.font = header_font
            cc.alignment = center
            gc.alignment = center
            cc.fill = fill
            gc.fill = fill
            if is_protein_level:
                grand[i * 2] += count
                grand[i * 2 + 1] += gm
            row_total_count += count
            row_total_gm += gm
        tc = ws_up.cell(row=r, column=total_count_col, value=row_total_count or None)
        tg = ws_up.cell(row=r, column=total_gm_col, value=row_total_gm or None)
        tc.font = header_font
        tg.font = header_font
        tc.alignment = center
        tg.alignment = center
        tc.fill = fill
        tg.fill = fill
        r += 1

    total_row = r
    gt_cell = ws_up.cell(row=total_row, column=1, value='Grand Total')
    gt_cell.font = header_font
    gt_cell.alignment = center
    gt_total_count, gt_total_gm = 0.0, 0.0
    for i in range(n_pkg):
        col = 2 + i * 2
        cc = ws_up.cell(row=total_row, column=col, value=grand[i * 2] or None)
        gc = ws_up.cell(row=total_row, column=col + 1, value=grand[i * 2 + 1] or None)
        cc.font = header_font
        gc.font = header_font
        cc.alignment = center
        gc.alignment = center
        gt_total_count += grand[i * 2]
        gt_total_gm += grand[i * 2 + 1]
    tc = ws_up.cell(row=total_row, column=total_count_col, value=gt_total_count or None)
    tg = ws_up.cell(row=total_row, column=total_gm_col, value=gt_total_gm or None)
    tc.font = header_font
    tg.font = header_font
    tc.alignment = center
    tg.alignment = center

    ws_up.column_dimensions['A'].width = 44
    for i in range(2, total_gm_col + 1):
        ws_up.column_dimensions[get_column_letter(i)].width = 11

    if out_path is None:
        out_path = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name
    wb.save(out_path)
    return out_path


def process_subscribers_invoice(file_storage, day_label_override=None, out_path=None):
    """الدالة الرئيسية اللي بيستدعيها الـ endpoint: من ملف مرفوع لملف جاهز."""
    lookup = load_lookup()
    rows, detected_day = read_subscribers_invoice(file_storage)
    if not rows:
        raise ValueError('الملف المرفوع فاضي أو مالوش صفوف بكمية أكبر من صفر')

    day_label = day_label_override or detected_day
    if not day_label:
        raise ValueError(
            'مش قادر أحدد اسم اليوم. سمّي الملف وفيه تاريخ بصيغة YYYY-MM-DD '
            '(زي 2026-07-21) أو ابعت اسم اليوم صراحة.'
        )

    dont_use_rows, pivot_rows, package_order, report = compute_decision_tables(rows, lookup)
    out_path = build_output_workbook(
        day_label, rows, dont_use_rows, pivot_rows, package_order, lookup, out_path=out_path
    )
    report['day_label'] = day_label
    return out_path, report
