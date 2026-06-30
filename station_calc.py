"""
محرك حساب موحّد لكل المحطات الستة (Rice / Breakfast / Dessert / Sauce / Salads /
Hot+Marination) — كل محطة بطريقتها الحقيقية المتأكَّد منها بالاختبار على بيانات
حقيقية (مش تخمين):

  - rice:      خلية الإدخال دايمًا Z1، جدول الباتشات يبدأ من صف 13 (مش 51).
  - breakfast/
    dessert:   خلية الإدخال بتتغيّر مكانها لكل صنف (اتلقت بالـ sensitivity
               resolver)، مفيش جدول باتشات، بس قايمة مكونات (B..H من صف 5).
  - salads:    خلية الإدخال = عمود H في شيت 'User' (جدول Table3)، مباشر بدون
               أي صيغة وسيطة.
  - sauce:     مفيش ملف إكسل خالص — معادلة بسيطة: الجرامات = عدد الأوردرات ×
               portion الوجبة (من meal_portions_data.json) × النسبة المئوية
               (من جدول "كمية الصوص بالوجبات" اللي حدد المستخدم).
"""
import os
import json
from pycel import ExcelCompiler

BASE_DIR = os.path.dirname(__file__)
STATION_ITEMS = json.load(open(os.path.join(BASE_DIR, 'data', 'station_items.json'), encoding='utf-8'))


def _full_path(rel_path):
    return os.path.join(BASE_DIR, rel_path)


# ===================== Breakfast / Dessert (قايمة مكونات) =====================

def calc_ingredient_item(station, sheet_name, order_count):
    info = STATION_ITEMS[station][sheet_name]
    path = _full_path(info['file'])
    input_cell = info['input_cell']
    portion_g = info.get('portion_g', 1.0)

    excel = ExcelCompiler(filename=path)
    addr = f"'{sheet_name}'!{input_cell}"

    for r in range(5, 40):
        for col in ['B', 'C', 'G', 'H']:
            excel.evaluate(f"'{sheet_name}'!{col}{r}")
    excel.evaluate(addr)

    new_value = round(order_count * portion_g, 2)
    excel.set_value(addr, new_value)

    rows = []
    for r in range(5, 40):
        name = excel.evaluate(f"'{sheet_name}'!B{r}")
        unit = excel.evaluate(f"'{sheet_name}'!C{r}")
        if not name or not str(name).strip():
            if rows:
                break
            continue
        if not unit:
            if rows:
                break
            continue
        amount = excel.evaluate(f"'{sheet_name}'!H{r}")
        if amount is None:
            amount = excel.evaluate(f"'{sheet_name}'!G{r}")
        if amount is None:
            continue
        rows.append({'label': str(name).strip(), 'unit': unit, 'amount': amount})

    return {
        'station': station, 'sheet_name': sheet_name, 'order_count': order_count,
        'arabic_name': info.get('arabic_name', ''),
        'input_value_set': new_value, 'mode': 'ingredients', 'rows': rows,
    }


# ===================== Rice (جدول باتشات يبدأ من صف 13) =====================

def calc_rice_item(sheet_name, order_count):
    """عدد الأوردرات هنا = جرامات الأرز المطلوبة مباشرة (مفيش portion ضرب)،
    لأن مفيش ربط تلقائي بين وجبة معينة وكمية الرز — الشيف بيحدد الكمية بنفسه."""
    info = STATION_ITEMS['rice'][sheet_name]
    path = _full_path(info['file'])
    excel = ExcelCompiler(filename=path)
    addr = f"'{sheet_name}'!Z1"

    for r in range(12, 17):
        for col in ['B', 'C', 'D']:
            excel.evaluate(f"'{sheet_name}'!{col}{r}")
    excel.evaluate(addr)
    excel.set_value(addr, order_count)

    rows = []
    for r in range(13, 16):
        label = excel.evaluate(f"'{sheet_name}'!B{r}")
        conv = excel.evaluate(f"'{sheet_name}'!C{r}")
        final_kg = excel.evaluate(f"'{sheet_name}'!D{r}")
        rows.append({'label': str(label) if label else f'Batch {r-12}',
                      'conversion_factor': conv, 'final_kg': final_kg})

    return {
        'station': 'rice', 'sheet_name': sheet_name, 'order_count': order_count,
        'arabic_name': info.get('arabic_name', ''),
        'mode': 'batches', 'rows': rows,
    }


# ===================== Salads (جدول Table3 مباشر) =====================

def calc_salad_item(salad_name, order_count):
    info = STATION_ITEMS['salads'][salad_name]
    path = _full_path(info['file'])
    excel = ExcelCompiler(filename=path)

    excel.evaluate(f"'Usage'!E2")  # نخلي pycel يبني شجرة الاعتماديات للـ Usage الأول
    for r in range(2, 60):
        excel.evaluate(f"'Usage'!B{r}")
        excel.evaluate(f"'Usage'!C{r}")
        excel.evaluate(f"'Usage'!H{r}")

    addr = f"'{info['input_sheet']}'!{info['input_cell']}"
    excel.evaluate(addr)
    excel.set_value(addr, order_count)

    rows = []
    for r in range(2, 60):
        salad = excel.evaluate(f"'Usage'!B{r}")
        if not salad or salad_name.split(' — ')[0].strip().lower() not in str(salad).lower():
            continue
        ing = excel.evaluate(f"'Usage'!C{r}")
        amount = excel.evaluate(f"'Usage'!H{r}")
        if ing and amount is not None:
            rows.append({'label': str(ing).strip(), 'unit': 'gm', 'amount': amount})

    return {
        'station': 'salads', 'sheet_name': salad_name, 'order_count': order_count,
        'mode': 'ingredients', 'rows': rows,
    }


# ===================== Sauce (معادلة بسيطة، بدون إكسل) =====================

def calc_sauce_for_meal(meal_name, order_count, portion_g):
    entries = STATION_ITEMS['sauce'].get(meal_name, [])
    rows = []
    for e in entries:
        grams = round(order_count * portion_g * e['pct'], 2)
        rows.append({'label': e['label'], 'unit': 'gm', 'amount': grams})
    return {
        'station': 'sauce', 'sheet_name': meal_name, 'order_count': order_count,
        'mode': 'ingredients', 'rows': rows,
    }


def list_station_items(station):
    """بترجع قايمة أسماء الأصناف المتاحة في محطة معينة (للداش بورد)."""
    if station not in STATION_ITEMS:
        return []
    return [{'sheet_name': k, 'arabic_name': v.get('arabic_name', '') if isinstance(v, dict) else ''}
            for k, v in STATION_ITEMS[station].items()]
