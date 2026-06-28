"""
محرك "Tokyo Ordering" — واجهة إدخال سهلة فوق ملف الإكسل الأصلي بتاعك.

مهم: الكود ده ميعملش أي حساب بديل عن الإكسل. هو بس بيقرا ويكتب في خليتين فاضيتين
لكل وجبة لكل يوم (عدد الوجبات والوزن بالجرام) في شيت All_Ingredients، ويرجّع نسخة
محدّثة من نفس الملف بالماكرو زي ما هو، عشان تفتحها في إكسل وتدوس على زرار
"Update" بتاعك زي العادة — الحساب الحقيقي يفضل في الماكرو نفسه، إحنا غيّرنا بس
طريقة إدخال الأرقام.
"""
import shutil
import tempfile
from openpyxl import load_workbook

DAY_NAMES = {1: 'السبت', 2: 'الأحد', 3: 'الاثنين', 4: 'الثلاثاء', 5: 'الأربعاء', 6: 'الخميس'}

DAY_NO_COL = 36   # AJ
SHEET_NAME_COL = 37  # AK
COUNT_COL = 44    # AR
GRAMS_COL = 45    # AS


def read_current_inputs(template_path):
    wb = load_workbook(template_path, data_only=False, keep_vba=True, read_only=False)
    ws = wb['All_Ingredients']
    days = {}
    for r in range(2, ws.max_row + 1):
        day_no = ws.cell(row=r, column=DAY_NO_COL).value
        sheet_name = ws.cell(row=r, column=SHEET_NAME_COL).value
        if not day_no or not sheet_name or sheet_name == 'Butchery':
            continue
        count = ws.cell(row=r, column=COUNT_COL).value
        grams = ws.cell(row=r, column=GRAMS_COL).value
        day_key = str(int(day_no))
        days.setdefault(day_key, {'dayNo': int(day_no), 'dayName': DAY_NAMES.get(int(day_no), f'يوم {day_no}'), 'meals': []})
        days[day_key]['meals'].append({
            'row': r,
            'sheetName': sheet_name,
            'count': count if isinstance(count, (int, float)) else 0,
            'grams': grams if isinstance(grams, (int, float)) else 0,
        })
    wb.close()
    return [days[k] for k in sorted(days.keys(), key=int)]


def write_updated_workbook(template_path, days_payload, out_path=None):
    """days_payload: نفس شكل read_current_inputs الناتج، بعد تعديل count/grams من المستخدم."""
    wb = load_workbook(template_path, data_only=False, keep_vba=True, read_only=False)
    ws = wb['All_Ingredients']

    for day in days_payload or []:
        for meal in day.get('meals') or []:
            row = meal.get('row')
            if not row:
                continue
            count = meal.get('count')
            grams = meal.get('grams')
            if count is not None:
                try:
                    ws.cell(row=row, column=COUNT_COL, value=float(count))
                except (TypeError, ValueError):
                    pass
            if grams is not None:
                try:
                    ws.cell(row=row, column=GRAMS_COL, value=float(grams))
                except (TypeError, ValueError):
                    pass

    if out_path is None:
        out_path = tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False).name
    wb.save(out_path)
    wb.close()
    return out_path
