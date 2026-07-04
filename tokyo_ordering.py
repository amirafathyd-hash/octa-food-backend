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


DAY_NAME_TO_NO = {v: k for k, v in DAY_NAMES.items()}


def read_day_file_meals(file_storage):
    """بتقرا شيت 'Update' من ملف يوم واحد (زي Octa_Food_Sat_...xlsx) وترجع:
    (day_no, {اسم الصنف: (Total Count, Total Grams)})
    - اسم اليوم بيتقرا من الخلية A6 (زي 'السبت') وبيتحول لرقم اليوم.
    - الأعمدة الثابتة: A=اسم الصنف، L=Total Count، M=Total Grams (الصفوف بتبدأ من 10)."""
    file_storage.seek(0)
    wb = load_workbook(file_storage, data_only=True, read_only=True)
    if 'Update' not in wb.sheetnames:
        raise ValueError("شيت 'Update' مش موجود في الملف ده")
    ws = wb['Update']

    day_label = str(ws['A6'].value or '').strip()
    day_no = DAY_NAME_TO_NO.get(day_label)
    if not day_no:
        raise ValueError(f"مش عارف أحدد اليوم من الخلية A6 (لقيت: '{day_label}')")

    meals = {}
    for r in range(10, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name:
            continue
        name = str(name).strip()
        count = ws.cell(row=r, column=12).value   # L = Total Count
        grams = ws.cell(row=r, column=13).value   # M = Total Grams
        count = count if isinstance(count, (int, float)) else None
        grams = grams if isinstance(grams, (int, float)) else None
        if count is None and grams is None:
            continue
        meals[name] = (count, grams)
    wb.close()
    file_storage.seek(0)
    return day_no, meals


def _normalize_meal_name(s):
    """بتشيل الإيموجي والمسافات الزيادة عشان المطابقة تنجح حتى لو فيه رمز
    حار 🌶️ أو مسافات مختلفة بين النسختين."""
    import re
    s = str(s or '')
    s = re.sub(r'[\U0001F300-\U0001FAFF\u2600-\u27BF]', '', s)  # إيموجي
    return re.sub(r'\s+', ' ', s).strip()


def merge_day_into_template(template_path, day_no, meals_by_name, out_path=None):
    """بتاخد قاموس {اسم الصنف: (count, grams)} من ملف يوم واحد، وتحدّث بيه
    صفوف نفس اليوم (AJ=day_no) بس في شيت All_Ingredients، بالمطابقة على
    عمود AQ (Meal name). بترجع (out_path, report) - الـreport بيوضح كل صنف
    اتطابق واتحدّث، وكل صنف في اليوم ده متلقاش ليه مطابقة (يفضل زي ما هو،
    من غير ما يتصفّر غلط)."""
    wb = load_workbook(template_path, data_only=False, keep_vba=True, read_only=False)
    ws = wb['All_Ingredients']

    norm_lookup = {_normalize_meal_name(k): v for k, v in meals_by_name.items()}

    matched, unmatched = [], []
    for r in range(2, ws.max_row + 1):
        row_day = ws.cell(row=r, column=DAY_NO_COL).value
        if not row_day or int(row_day) != day_no:
            continue
        meal_name = ws.cell(row=r, column=43).value  # AQ
        if not meal_name:
            continue
        key = str(meal_name).strip()
        norm_key = _normalize_meal_name(key)
        found = meals_by_name.get(key) or norm_lookup.get(norm_key)
        if found is None:
            unmatched.append(key)
            continue
        count, grams = found
        if count is not None:
            ws.cell(row=r, column=COUNT_COL, value=float(count))
        if grams is not None:
            ws.cell(row=r, column=GRAMS_COL, value=float(grams))
        matched.append({'row': r, 'name': key, 'count': count, 'grams': grams})

    if out_path is None:
        out_path = tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False).name
    wb.save(out_path)
    wb.close()

    report = {
        'day_no': day_no,
        'day_name': DAY_NAMES.get(day_no, str(day_no)),
        'matched_count': len(matched),
        'unmatched_count': len(unmatched),
        'matched': matched,
        'unmatched': unmatched,  # أصناف في اليوم ده مالقتش ليها رقم في الملف المرفوع - اتسابت زي ما هي
    }
    return out_path, report


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
