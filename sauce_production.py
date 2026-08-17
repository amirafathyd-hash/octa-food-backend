"""Permanent count-driven production engine for the approved sauce workbook."""
import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime

from openpyxl import load_workbook

from sauce_storage import SAUCE_TEMPLATE_PATH
from tokyo_ordering import read_day_file_payload


DAY_NAMES = {1: 'السبت', 2: 'الأحد', 3: 'الاثنين', 4: 'الثلاثاء', 5: 'الأربعاء', 6: 'الخميس'}
# The approved master is arranged in these six production-day blocks.
DAY_BLOCK_LENGTHS = (9, 7, 6, 7, 8, 4)


def _number(value):
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value or '').replace(',', '').strip())
    except (TypeError, ValueError):
        return 0.0


def _name_key(value):
    text = str(value or '').strip().casefold()
    text = re.sub(r'[\u064b-\u065f\u0670]', '', text).replace('ـ', '')
    text = text.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    text = text.replace('ى', 'ي').replace('ة', 'ه')
    return re.sub(r'[^a-z0-9\u0600-\u06ff]+', ' ', text).strip()


def _recipe_sheet_names(wb):
    return [name for name in wb.sheetnames if name != 'Ordering']


def _day_sheets(wb, day_no):
    day_no = int(_number(day_no))
    if day_no not in DAY_NAMES:
        raise ValueError('رقم يوم الصوص غير صحيح')
    recipes = _recipe_sheet_names(wb)
    start = sum(DAY_BLOCK_LENGTHS[:day_no - 1])
    end = start + DAY_BLOCK_LENGTHS[day_no - 1]
    selected = recipes[start:end]
    if len(selected) != DAY_BLOCK_LENGTHS[day_no - 1]:
        raise ValueError('ترتيب شيتات الصوص لا يطابق الـMaster المعتمد')
    return selected


# Each tuple is one operational meal. Aliases in a tuple are the same row;
# different tuples are deliberately added together (for shared recipes).
SOURCE_GROUPS = {
    'Tahina Small': [('Foul', 'فول'), ('Fish', 'سمك')],
    'Red sauce': [('Foul', 'فول'), ('Fish', 'سمك')],
    'Cucumber yogurt sauce': [('Kabsa', 'كبسة')],
    'Herbal Sauce': [('Herbal Chicken', 'دجاج الأعشاب')],
    'Red sauce Big': [('Kabsa', 'كبسة'), ('Mandi', 'مندي')],
    'Tahina Big': [('Mandi', 'مندي')],
    'Coctail Sauce Weight': [
        ('Asian chicken sandwich with oat bread', 'ساندويتش الدجاج الآسيوي'),
        ('Chicken Awsal sandwich', 'ساندويتش دجاج اوصل'),
    ],
    'Lemon Herb Sauce Packed': [('Arabic Chicken Burger', 'برجر دجاج العربي')],
    'Lemon Herb Sauce unpacked': [('Arabic Chicken Burger', 'برجر دجاج العربي')],
    'Cucumber yogurt sauce (2)': [('Zurbian', 'زربيان')],
    'Red sauce Big (2)': [('Zurbian', 'زربيان')],
    'House Sauce Packed': [('Beef Burger BBQ', 'برجر لحم باربكيو')],
    'House Sauce unpacked': [('Beef Burger BBQ', 'برجر لحم باربكيو')],
    'Honey Mustard Sauce Packed': [('Chicken Burger Honey Mustard', 'برجر دجاج بالعسل والخردل')],
    'Honey Mustard Sauce unpacked': [('Chicken Burger Honey Mustard', 'برجر دجاج بالعسل والخردل')],
    'Tahina Big (2)': [('Mandi', 'مندي')],
    'Red sauce Big (3)': [('Mandi', 'مندي'), ('Chicken Saleeq', 'سليق دجاج', 'Saleeq')],
    'Coctail Sauce Dip In  (3)': [('Almond chicken in the oven with potato wedges', 'دجاج باللوز في الفرن', 'Almond Chicken', 'دجاج باللوز')],
    'Coctail Sauce Weight (2)': [
        ('Chicken Fajita Sandwich on Oat Bread', 'ساندوتش فاهيتا الدجاج بخبز الشوفان'),
        ('Chicken Bell Pepper Sandwich', 'ساندوتش الدجاج بالفلفل الرومي'),
    ],
    'Mixed Pickle': [
        ('Chicken Fajita Sandwich on Oat Bread', 'ساندوتش فاهيتا الدجاج بخبز الشوفان'),
        ('Chicken Bell Pepper Sandwich', 'ساندوتش الدجاج بالفلفل الرومي'),
    ],
    'Smoked Sauce Packed': [('Smoky Beef Burger', 'برجر سموكي لحم')],
    'Smoked Sauce unpacked': [('Smoky Beef Burger', 'برجر سموكي لحم')],
    'Cucumber yogurt sauce (3)': [('Chicken Maqluba', 'مقلوبة دجاج', 'Maklouba', 'مقلوبة')],
    'Shabat Sauce': [
        ('BBQ Chicken Sandwich in Ciabatta Bread', 'ساندوتش الدجاج بالباربيكيو بخبز الشيباتا'),
        ('Grilled chicken sandwich', 'ساندوتش الدجاج المشوي'),
    ],
    'Biryani Yoghrt sauce': [('Tikka chicken with Buryani rice', 'دجاج تكا', 'Chicken Tikka')],
    'Coctail Sauce Dip In  (4)': [('Octa Chicken Poke bowl (spicy)', 'أوكتا بوكي بول الدجاج', 'Octa PokiPowl')],
    'Coctail Sauce Weight (5)': [('Octa Chicken Poke bowl (spicy)', 'أوكتا بوكي بول الدجاج', 'Octa PokiPowl')],
    'Coleslaw Salad': [('Classic Chicken Burger', 'برجر الدجاج الكلاسيكي')],
    'Garlic Aioli (2)': [('Classic Chicken Burger', 'برجر الدجاج الكلاسيكي')],
    'Red sauce Big (4)': [('Boukhary Beef', 'بخارى لحم', 'Boukhary')],
    'Cucumber yogurt sauce (4)': [('Boukhary Beef', 'بخارى لحم', 'Boukhary')],
    'Coctail Sauce Weight (4)': [
        ('Philadelphia beef sandwich with oat bread', 'ساندويتش فلادلفيا لحم'),
        ('Classic beef sandwich', 'Classic meat sandwich', 'ساندوتش اللحم الكلاسيكي'),
    ],
    'BBQ Sauce Packed': [('Chicken Burger BBQ', 'برجر دجاج باربكيو')],
    'BBQ Sauce unpacked': [('Chicken Burger BBQ', 'برجر دجاج باربكيو')],
    'Sumak Onion': [('Kebab Sandwich', 'ساندويتش كباب لحم')],
    'Tahina Small (2)': [('Fish with lemon', 'سمك بالليمون'), ('Kebab Sandwich', 'ساندويتش كباب لحم')],
    'Red sauce (2)': [('Fish with lemon', 'سمك بالليمون')],
    'Red sauce Big (5)': [('Chicken Saleeq', 'سليق دجاج', 'Saleeq')],
    'Cucumber yogurt sauce (5)': [('Kabli', 'كابلى', 'كابلي')],
    'Garlic Aioli': [('Beef Burger Arabic', 'برجر لحم عربي')],
    'Garlic Aioli unpacked': [('Beef Burger Arabic', 'برجر لحم عربي')],
}


def _meal_count_lookup(meals):
    result = {}
    for name, pair in (meals or {}).items():
        count = _number(pair[0] if isinstance(pair, (list, tuple)) else pair)
        key = _name_key(name)
        if key:
            result[key] = count
    return result


def _match_group(aliases, lookup):
    for alias in aliases:
        key = _name_key(alias)
        if key in lookup:
            return lookup[key], key
    candidates = []
    for alias in aliases:
        key = _name_key(alias)
        if len(key) < 5:
            continue
        for source_key, count in lookup.items():
            if key in source_key or source_key in key:
                candidates.append((source_key, count))
    unique = {key: value for key, value in candidates}
    if len(unique) == 1:
        key, value = next(iter(unique.items()))
        return value, key
    return 0.0, ''


def _compute_day_counts(wb, day_no, meals):
    lookup = _meal_count_lookup(meals)
    results, missing = [], []
    for sheet_name in _day_sheets(wb, day_no):
        ws = wb[sheet_name]
        groups = SOURCE_GROUPS.get(sheet_name)
        if not groups:
            meal_title = str(ws['B3'].value or '').strip()
            parts = [part.strip(' ()') for part in re.split(r'\s[-–—]\s', meal_title) if part.strip(' ()')]
            groups = [tuple(parts or [meal_title])]
        total, sources = 0.0, []
        for aliases in groups:
            count, matched_key = _match_group(aliases, lookup)
            total += count
            sources.append({'aliases': list(aliases), 'count': count, 'matched': bool(matched_key)})
            if not matched_key:
                missing.append(str(aliases[0]))
        results.append({'sheet': sheet_name, 'input_count': round(total, 3), 'safety_count': 0.0, 'sources': sources})
    return results, sorted(set(missing), key=_name_key)


def _soffice_bin():
    return os.environ.get('SOFFICE_BIN') or shutil.which('soffice') or 'soffice'


def _convert(source_path, extension, prefix):
    out_dir = tempfile.mkdtemp(prefix=prefix)
    profile_dir = tempfile.mkdtemp(prefix=f'{prefix}profile_')
    proc = subprocess.run([
        _soffice_bin(), f'-env:UserInstallation=file://{profile_dir}', '--headless',
        '--convert-to', extension, '--outdir', out_dir, source_path,
    ], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=150)
    if proc.returncode != 0:
        raise RuntimeError((proc.stderr or proc.stdout or 'LibreOffice failed').strip())
    result = os.path.join(out_dir, os.path.splitext(os.path.basename(source_path))[0] + f'.{extension}')
    if not os.path.exists(result):
        raise RuntimeError(f'تعذر إنشاء ملف {extension.upper()} للصوص')
    return result


def _write_inputs(day_no, inputs, template_path=SAUCE_TEMPLATE_PATH):
    wb = load_workbook(template_path, data_only=False, keep_vba=True)
    try:
        selected = set(_day_sheets(wb, day_no))
        values = {str(item.get('sheet') or ''): max(0.0, _number(item.get('input_count'))) for item in inputs}
        safety = {str(item.get('sheet') or ''): max(0.0, _number(item.get('safety_count'))) for item in inputs}
        wb['Ordering']['R1'] = int(day_no)
        for sheet_name in selected:
            wb[sheet_name]['V1'] = values.get(sheet_name, 0.0)
            wb[sheet_name]['Q19'] = safety.get(sheet_name, 0.0)
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = 'auto'
        output = tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False).name
        wb.save(output)
    finally:
        wb.close()
    return output


def _pdf_source(calculated_path, day_no):
    wb = load_workbook(calculated_path, data_only=False)
    try:
        selected = set(_day_sheets(wb, day_no))
        for sheet_name in list(wb.sheetnames):
            if sheet_name not in selected:
                wb.remove(wb[sheet_name])
        if not wb.sheetnames:
            raise ValueError('لا توجد وصفات صوص مرتبطة بهذا اليوم')
        wb.active = 0
        output = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name
        wb.save(output)
        return output, len(wb.sheetnames)
    finally:
        wb.close()


def _state_from_workbook(path, day_no):
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        items = []
        for sheet_name in _day_sheets(wb, day_no):
            ws = wb[sheet_name]
            items.append({
                'sheet': sheet_name,
                'name': str(ws['B2'].value or sheet_name).strip(),
                'meal': str(ws['B3'].value or '').strip(),
                'input_count': round(_number(ws['V1'].value), 3),
                'safety_count': 0,
                'final_count': round(_number(ws['Q20'].value), 3),
            })
        return {'day_no': int(day_no), 'day_name': DAY_NAMES[int(day_no)], 'items': items}
    finally:
        wb.close()


def get_sauce_production_state(template_path=SAUCE_TEMPLATE_PATH, day_no=1):
    if not os.path.exists(template_path):
        raise FileNotFoundError('ملف الصوص الأساسي غير موجود')
    state = _state_from_workbook(template_path, int(_number(day_no) or 1))
    state['template_updated_at'] = datetime.fromtimestamp(os.path.getmtime(template_path)).isoformat(timespec='seconds')
    return state


def _files_from_inputs(day_no, inputs, template_path=SAUCE_TEMPLATE_PATH):
    macro_path = _write_inputs(day_no, inputs, template_path)
    calculated = _convert(macro_path, 'xlsx', 'sauce_recalc_')
    pdf_source, page_count = _pdf_source(calculated, day_no)
    pdf_path = _convert(pdf_source, 'pdf', 'sauce_pdf_')
    report = {
        'day_no': int(day_no), 'page_count': page_count,
        'matched_count': sum(1 for item in inputs if _number(item.get('input_count')) > 0),
        'state': _state_from_workbook(calculated, day_no),
    }
    return calculated, pdf_path, report


def build_sauce_day_files(file_storage, template_path=SAUCE_TEMPLATE_PATH, safety_items=None, expected_day_no=None):
    day_no, meals, input_report = read_day_file_payload(file_storage)
    if expected_day_no is not None and int(_number(expected_day_no)) != int(day_no):
        raise ValueError(f'ملف اليوم يخص يوم {day_no} بينما اليوم المختار في لوحة الصوص هو {int(_number(expected_day_no))}')
    wb = load_workbook(template_path, data_only=False, keep_vba=True)
    try:
        inputs, missing = _compute_day_counts(wb, day_no, meals)
    finally:
        wb.close()
    safety_lookup = {
        str(item.get('sheet') or ''): max(0.0, _number(item.get('safety_count')))
        for item in (safety_items or [])
    }
    for item in inputs:
        item['safety_count'] = safety_lookup.get(item['sheet'], 0.0)
    if missing:
        raise ValueError('لم تتم مطابقة وجبات الصوص التالية: ' + '، '.join(missing))
    excel_path, pdf_path, report = _files_from_inputs(day_no, inputs, template_path)
    report.update({'input_report': input_report, 'missing_sources': missing, 'inputs': inputs})
    return excel_path, pdf_path, report


def build_sauce_manual_files(day_no, items, template_path=SAUCE_TEMPLATE_PATH):
    day_no = int(_number(day_no))
    wb = load_workbook(template_path, data_only=False, read_only=True, keep_vba=True)
    try:
        allowed = set(_day_sheets(wb, day_no))
    finally:
        wb.close()
    inputs = []
    for item in items or []:
        sheet = str(item.get('sheet') or '')
        if sheet in allowed:
            inputs.append({
                'sheet': sheet,
                'input_count': max(0.0, _number(item.get('input_count'))),
                'safety_count': max(0.0, _number(item.get('safety_count'))),
            })
    return _files_from_inputs(day_no, inputs, template_path)


def package_sauce_files(excel_path, pdf_path, day_no):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(excel_path, f'Day{int(day_no)}_Sauce_Updated.xlsx')
        archive.write(pdf_path, f'Day{int(day_no)}_Sauce.pdf')
    buffer.seek(0)
    return buffer


def replace_sauce_production_template(file_storage, template_path=SAUCE_TEMPLATE_PATH):
    if os.path.splitext(file_storage.filename or '')[1].lower() != '.xlsm':
        raise ValueError('الشيت الأساسي للصوص لازم يكون XLSM')
    upload = tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False).name
    file_storage.seek(0)
    file_storage.save(upload)
    wb = load_workbook(upload, data_only=False, read_only=True, keep_vba=True)
    try:
        if 'Ordering' not in wb.sheetnames or len(_recipe_sheet_names(wb)) != sum(DAY_BLOCK_LENGTHS):
            raise ValueError('الشيت الجديد لا يطابق ترتيب Master الصوص المعتمد')
        for day_no in DAY_NAMES:
            for sheet_name in _day_sheets(wb, day_no):
                wb[sheet_name]['V1']
                wb[sheet_name]['Q19']
                wb[sheet_name]['Q20']
    finally:
        wb.close()
    shutil.copy2(upload, template_path)
    return get_sauce_production_state(template_path, 1), {'template_file': os.path.basename(template_path)}
