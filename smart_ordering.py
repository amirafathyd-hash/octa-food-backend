"""
محرك الطلب الذكي (Smart Ordering Engine)
==========================================
بياخد اسم الوجبة + عدد الأوردرات، ويحسب جدول التجهيز (Batches) بتاعها عن طريق
تشغيل صيغ الإكسل الحقيقية لنفس الوجبة (مش إعادة كتابة المعادلة بإيد) — عشان
نضمن إن كل وجبة من الـ88 بتتحسب بمعادلتها الخاصة بالظبط زي ما هي في الملف
الأصلي، حتى لو الصيغة مختلفة عن باقي الوجبات.

الفكرة: 
  1. Z1 ("Update") هي الخلية الوحيدة اللي بتتغيّر يدويًا في الإكسل الأصلي.
  2. بنحسبها هنا = عدد الأوردرات × حصة الوجبة بالجرام (مستخرجة من نفس الملف
     المرجعي، عمود Z1/Count الأصلي لكل وجبة).
  3. بنستخدم pycel (بيشغّل صيغ إكسل حقيقية، مش إعادة تنفيذ يدوي) عشان نحسب
     باقي الجدول من القيمة الجديدة دي.
"""
import os
import json
from pycel import ExcelCompiler

TOKYO_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'tokyo_ordering_template.xlsm')
PORTIONS_PATH = os.path.join(os.path.dirname(__file__), 'meal_portions_data.json')
PACKAGES_PATH = os.path.join(os.path.dirname(__file__), 'menu_packages.json')

with open(PORTIONS_PATH, encoding='utf-8') as f:
    # { meal_name: [portion_grams_per_order, arabic_name] }
    MEAL_PORTIONS = json.load(f)

with open(PACKAGES_PATH, encoding='utf-8') as f:
    # { "Saturday": [["اسم المنيو", "اسم الشيت"], ...], ... }
    MENU_PACKAGES = json.load(f)

DAY_LABELS_AR = {
    'Saturday': 'السبت', 'Sunday': 'الأحد', 'Monday': 'الإثنين',
    'Tuesday': 'الثلاثاء', 'Wednesday': 'الأربعاء', 'Thursday': 'الخميس',
    'Friday': 'الجمعة',
}

def _get_excel():
    """بنعمل نسخة جديدة من المحرك في كل مرة (مش بنعيد استخدام نسخة قديمة) —
    لأن pycel مابيعملش invalidation كامل لكل الخلايا المعتمدة على خلية اتغيّرت
    لو استخدمنا نفس الكائن أكتر من مرة، وده كان بيسبب أرقام قديمة/خاطئة في
    بعض الحالات. تحميل الملف بياخد وقت أطول شوية، بس ده ضمان الدقة المطلوب."""
    return ExcelCompiler(filename=TOKYO_TEMPLATE_PATH)


def list_available_meals():
    """بترجع قائمة بكل الوجبات المتاحة للطلب (اسم إنجليزي + عربي لو موجود)،
    مرتبة أبجديًا، عشان تتعرض في الـ dropdown بتاع الداش بورد."""
    out = []
    for name, (portion, arabic) in MEAL_PORTIONS.items():
        out.append({'name': name, 'arabic_name': arabic, 'portion_g': portion})
    out.sort(key=lambda m: m['name'].lower())
    return out


def list_menu_packages():
    """بترجع الوجبات منظّمة حسب باقات الأيام (زي المنيو اللي العملاء بيطلبوا
    منه فعليًا)، بالاسم اللي العميل شايفه + اسم الشيت الحقيقي اللي هيتحسب
    عليه. أي وجبة موجودة في الشيتات بس مش في أي باقة لسه، بتتحط في باقة
    "أخرى" عشان تفضل متاحة للإضافة من غير ما تضيع."""
    mapped_sheets = set()
    packages = []
    for day_key, items in MENU_PACKAGES.items():
        day_items = []
        for menu_name, sheet_name in items:
            info = MEAL_PORTIONS.get(sheet_name)
            if not info:
                continue  # لو الشيت اتشال من الملف الأصلي مستقبلًا
            mapped_sheets.add(sheet_name)
            day_items.append({
                'menu_name': menu_name,
                'sheet_name': sheet_name,
                'arabic_name': info[1],
                'portion_g': info[0],
            })
        packages.append({
            'day_key': day_key,
            'day_label_ar': DAY_LABELS_AR.get(day_key, day_key),
            'items': day_items,
        })

    extra_items = []
    for sheet_name, (portion, arabic) in MEAL_PORTIONS.items():
        if sheet_name not in mapped_sheets:
            extra_items.append({
                'menu_name': sheet_name,
                'sheet_name': sheet_name,
                'arabic_name': arabic,
                'portion_g': portion,
            })
    extra_items.sort(key=lambda m: m['sheet_name'].lower())
    if extra_items:
        packages.append({'day_key': 'Extra', 'day_label_ar': 'أصناف أخرى', 'items': extra_items})

    return packages


def _read_cell(excel, sheet, cell):
    try:
        return excel.evaluate(f"'{sheet}'!{cell}")
    except Exception:
        return None


def _read_block(excel, sheet):
    """بيرجّع جدول الباتشات الكامل (يمين وشمال، زي ما بنستخرجه في محطات
    التجهيز) بعد ما يكون Z1 اتغيّر للقيمة الجديدة."""
    rows = []
    labels_left = {51: 'For 10KG Sauce output', 52: 'Total', 53: 'Batch 1', 54: 'Batch 2', 55: 'Batch 3'}
    for r in range(51, 56):
        left_label = _read_cell(excel, sheet, f'B{r}') or labels_left.get(r, '')
        conv = _read_cell(excel, sheet, f'C{r}')
        final_kg = _read_cell(excel, sheet, f'D{r}')
        right_label = _read_cell(excel, sheet, f'G{r}')
        conv_r = _read_cell(excel, sheet, f'H{r}')
        uncooked = _read_cell(excel, sheet, f'I{r}')
        cooked = _read_cell(excel, sheet, f'J{r}')
        rows.append({
            'label': str(left_label) if left_label else '',
            'conversion_factor': conv,
            'final_kg': final_kg,
            'protein_label': str(right_label) if right_label else '',
            'protein_conversion_factor': conv_r,
            'uncooked_protein': uncooked,
            'cooked_protein': cooked,
        })
    return rows


def _has_batch_table(excel, sheet):
    """بعض الوجبات (زي السندوتشات) معندهاش جدول الباتشات (B51:J55) خالص —
    بتتحسب بطريقة أبسط (وزن كل مكوّن = عدد الأوردرات × كمية المكوّن في
    الوحدة الواحدة، من غير batches خالص). بنتأكد الأول قبل ما نحاول نقرا
    جدول مش موجود، عشان منرميش خطأ غلط."""
    val = _read_cell(excel, sheet, 'B51')
    return val is not None and str(val).strip() != ''


def _read_ingredient_list(excel, sheet, max_row=40):
    """بديل لجدول الباتشات للوجبات اللي مالهاش الجدول ده (زي السندوتشات) —
    بيرجّع كل صف مكوّن (B..H) بعد ما يكون Z1 اتغيّر، عشان نشوف وزن كل
    مكوّن للعدد الجديد من الأوردرات. بنوقف عند أول صف فاضي عشان منلقطش
    أقسام تانية تحت في نفس الشيت (زي مكونات التتبيلة المنفصلة)."""
    rows = []
    for r in range(5, max_row + 1):
        name = _read_cell(excel, sheet, f'B{r}')
        unit = _read_cell(excel, sheet, f'C{r}')
        if not name or not str(name).strip():
            if rows:
                break  # خلصنا أول قسم حقيقي، نوقف هنا
            continue
        if not unit:
            # صف عنوان قسم جديد (زي "Base Recipe (10kg Yield)") مش مكوّن حقيقي
            if rows:
                break
            continue
        amount = _read_cell(excel, sheet, f'H{r}') or _read_cell(excel, sheet, f'G{r}')
        if amount is None:
            continue
        rows.append({'label': str(name).strip(), 'unit': unit, 'amount': amount})
    return rows


def calculate_meal(meal_name, order_count):
    """بياخد اسم وجبة وعدد أوردرات، ويرجّع جدول الباتشات المحسوب بناءً على
    صيغ الإكسل الحقيقية لنفس الوجبة دي."""
    if meal_name not in MEAL_PORTIONS:
        raise ValueError(f'الوجبة "{meal_name}" مش موجودة في قائمة الوجبات المتاحة')
    if not isinstance(order_count, (int, float)) or order_count <= 0:
        raise ValueError('عدد الأوردرات لازم يكون رقم أكبر من صفر')

    portion_g, arabic_name = MEAL_PORTIONS[meal_name]
    z1_value = round(order_count * portion_g, 2)

    excel = _get_excel()
    z1_addr = f"'{meal_name}'!Z1"

    if _has_batch_table(excel, meal_name):
        # مهم جدًا: لازم نقرا كل الخلايا النهائية اللي محتاجينها الأول (قبل ما
        # نغيّر Z1)، عشان pycel يبني شجرة الاعتماديات الكاملة من Z1 لحدهم.
        target_cells = ['B', 'C', 'D', 'G', 'H', 'I', 'J']
        for r in range(51, 56):
            for col in target_cells:
                excel.evaluate(f"'{meal_name}'!{col}{r}")
        excel.set_value(z1_addr, z1_value)
        rows = _read_block(excel, meal_name)
        mode = 'batches'
    else:
        # مفيش جدول باتشات للوجبة دي (زي السندوتشات) — بنحسب قايمة المكونات
        # مباشرة بدل منه.
        for r in range(5, 40):
            for col in ['B', 'C', 'G', 'H']:
                excel.evaluate(f"'{meal_name}'!{col}{r}")
        excel.evaluate(z1_addr)
        excel.set_value(z1_addr, z1_value)
        ingredients = _read_ingredient_list(excel, meal_name)
        rows = [{
            'label': ing['label'],
            'conversion_factor': None,
            'final_kg': ing['amount'],
            'protein_label': '',
            'protein_conversion_factor': None,
            'uncooked_protein': None,
            'cooked_protein': None,
            'unit': ing['unit'],
        } for ing in ingredients]
        mode = 'ingredients'

    return {
        'meal_name': meal_name,
        'arabic_name': arabic_name,
        'order_count': order_count,
        'portion_g': portion_g,
        'z1_calculated': z1_value,
        'mode': mode,
        'rows': rows,
    }


def calculate_multiple(meal_orders):
    """meal_orders: [{'meal_name': ..., 'order_count': ...}, ...]
    بيحسب كل وجبة على حدة ويرجّع كل النتائج مع بعض."""
    results = []
    errors = []
    for item in meal_orders:
        name = item.get('meal_name')
        count = item.get('order_count')
        try:
            results.append(calculate_meal(name, count))
        except Exception as e:
            errors.append({'meal_name': name, 'error': str(e)})
    return {'results': results, 'errors': errors}
