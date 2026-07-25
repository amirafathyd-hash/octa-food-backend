import re
from copy import copy
from datetime import datetime
from io import BytesIO

from openpyxl import load_workbook


DAY_OPTIONS = {
    'السبت': (1, 'السبت'),
    'الأحد': (2, 'الأحد'),
    'الاحد': (2, 'الأحد'),
    'الإثنين': (3, 'الاثنين'),
    'الاثنين': (3, 'الاثنين'),
    'الثلاثاء': (4, 'الثلاثاء'),
    'الأربعاء': (5, 'الاربعاء'),
    'الاربعاء': (5, 'الاربعاء'),
    'الخميس': (6, 'خميس'),
    'خميس': (6, 'خميس'),
}

ARABIC_RE = re.compile(r'[\u0600-\u06FF]+(?:[\s\u0600-\u06FF]+)*')
PART_RE = re.compile(r'(\d+)\s*/\s*(\d+)')
STANDALONE_TOTAL_RE = re.compile(r'^total\b', re.I)


def _cell(ws, row, col):
    value = ws.cell(row, col).value
    if isinstance(value, float):
        return round(value, 6)
    return value


def _number(value):
    try:
        n = float(value or 0)
    except (TypeError, ValueError):
        return None
    if abs(n) < 0.000001:
        return None
    return round(n, 3)


def _norm_text(value):
    text = str(value or '').strip()
    text = text.replace('إ', 'ا').replace('أ', 'ا').replace('آ', 'ا')
    text = text.replace('ة', 'ه').replace('ى', 'ي')
    text = re.sub(r'[ـ،,]+', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.lower().strip()


def _batch_from_any_text(value, rtl_template=False):
    match = PART_RE.search(str(value or ''))
    if not match:
        return ''
    left, right = match.group(1), match.group(2)
    # In Arabic RTL labels, "2/1" is visually "1 of 2".
    return f'{right}/{left}' if rtl_template else f'{left}/{right}'


def _strip_batch(value):
    text = str(value or '')
    text = re.sub(r'دفعة\s*\d+\s*/\s*\d+', ' ', text)
    text = re.sub(r'\d+\s*/\s*\d+', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()


def _arabic_name(ws):
    value = _cell(ws, 2, 37)  # AK2
    if isinstance(value, str):
        match = ARABIC_RE.search(value)
        if match:
            return match.group(0).strip()
    return None


def _find_day_meals(wb, day_no):
    if 'All_Ingredients' not in wb.sheetnames:
        return []
    ws = wb['All_Ingredients']
    day_col = name_col = header_row = None
    for row in range(1, min(ws.max_row, 200) + 1):
        for col in range(1, min(ws.max_column, 80) + 1):
            value = str(ws.cell(row, col).value or '').strip()
            if value == 'Day No.':
                day_col = col
                header_row = row
            elif value == 'Sheet Name':
                name_col = col
        if day_col and name_col:
            break
    if not day_col or not name_col or not header_row:
        return []

    names = []
    for row in range(header_row + 1, min(ws.max_row, 500) + 1):
        try:
            same_day = int(ws.cell(row, day_col).value or 0) == int(day_no)
        except (TypeError, ValueError):
            same_day = False
        name = str(ws.cell(row, name_col).value or '').strip()
        if same_day and name and name in wb.sheetnames and name not in names:
            names.append(name)
    return names


def _find_nearest_banner(ws, before_row):
    for row in range(before_row - 1, max(1, before_row - 30), -1):
        value = ws.cell(row, 2).value
        if isinstance(value, str) and PART_RE.search(value):
            arabic = ws.cell(row - 1, 2).value
            return {
                'arabic': str(arabic or '').strip(),
                'english': value.strip(),
                'batch': _batch_from_any_text(value),
            }
    return None


def _total_rows(ws):
    rows = []
    part = 0
    for row in range(1, ws.max_row + 1):
        b = ws.cell(row, 2).value
        c = ws.cell(row, 3).value
        c_text = str(c or '').strip().lower()
        if c_text == 'total sauce':
            part += 1
            value_row = row + 1
            rows.append({
                'part': part,
                'simple': False,
                'banner': _find_nearest_banner(ws, row),
                'label': b or 'Protein',
                'protein': _number(ws.cell(value_row, 2).value),
                'total_sauce': _number(ws.cell(value_row, 3).value),
                'mix': _number(ws.cell(value_row, 5).value),
                'topping': _number(ws.cell(value_row, 7).value),
                'protein_mix': _number(ws.cell(value_row, 8).value),
            })
        elif (
            isinstance(b, str)
            and STANDALONE_TOTAL_RE.search(b.strip())
            and b.strip().lower() != 'total'
            and (c is None or str(c).strip() == '')
        ):
            part += 1
            value_row = row + 1
            rows.append({
                'part': part,
                'simple': True,
                'banner': _find_nearest_banner(ws, row),
                'label': b.strip(),
                'value': _number(ws.cell(value_row, 2).value),
            })
    return [r for r in rows if any(_number(v) for v in r.values() if isinstance(v, (int, float)))]


def _component_for_template_name(name):
    compact = _norm_text(_strip_batch(name))
    if 'صوص الاضافي' in compact or 'صوص اضافي' in compact:
        return 'topping'
    if 'بدون صوص' in compact:
        return 'protein'
    if 'مع الصوص' in compact:
        return 'protein_mix'
    if 'صوص' in compact:
        return 'total_sauce'
    return 'simple'


def _base_for_template_name(name):
    text = _strip_batch(name)
    for phrase in ('صوص الاضافي', 'صوص اضافي', 'بدون صوص', 'مع الصوص', 'صوص'):
        text = text.replace(phrase, ' ')
    return _norm_text(text)


def _candidate_key(base, batch, component):
    return f'{_norm_text(base)}|{batch or ""}|{component}'


def _build_tokyo_value_index(wb, day_no):
    index = {}
    diagnostics = []
    for sheet_name in _find_day_meals(wb, day_no):
        ws = wb[sheet_name]
        sheet_ar = _arabic_name(ws) or sheet_name
        for item in _total_rows(ws):
            banner = item.get('banner') or {}
            base_ar = _strip_batch(banner.get('arabic') or sheet_ar)
            batch = banner.get('batch') or (f"{item['part']}/1" if item.get('part') else '')
            if item.get('simple'):
                value = item.get('value')
                if value is not None:
                    index[_candidate_key(base_ar, batch, 'simple')] = value
                    index[_candidate_key(base_ar, '', 'simple')] = value
                    diagnostics.append({'name': base_ar, 'batch': batch, 'component': 'simple', 'value': value})
                continue
            for component in ('protein', 'total_sauce', 'topping', 'protein_mix'):
                value = item.get(component)
                if value is not None:
                    index[_candidate_key(base_ar, batch, component)] = value
                    diagnostics.append({'name': base_ar, 'batch': batch, 'component': component, 'value': value})
    return index, diagnostics


def _build_weight_index(entries):
    index = {}
    rows = {}
    for row in entries or []:
        if row.get('deleted'):
            continue
        item_name = row.get('item_name') or ''
        batch = str(row.get('batch_no') or '').strip()
        if not batch:
            batch = _batch_from_any_text(item_name, rtl_template=True)
        base = _base_for_template_name(item_name)
        component = _component_for_template_name(item_name)
        try:
            weight = float(row.get('weight') or 0)
        except (TypeError, ValueError):
            continue
        if weight <= 0:
            continue
        key = _candidate_key(base, batch, component)
        index[key] = round(index.get(key, 0) + weight, 3)
        rows[key] = {
            'item_name': str(item_name).strip(),
            'base': base,
            'batch': batch,
            'component': component,
            'weight': index[key],
        }
    return index, rows


def _batch_sort_value(batch):
    match = PART_RE.search(str(batch or ''))
    if not match:
        return (999, 999)
    return (int(match.group(2)), int(match.group(1)))


def _component_sort_value(component):
    return {
        'protein': 1,
        'total_sauce': 2,
        'protein_mix': 3,
        'topping': 4,
        'simple': 5,
    }.get(component, 9)


def _copy_row_style(ws, source_row, target_row):
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)


def _last_template_item_row(ws):
    for row in range(ws.max_row, 3, -1):
        if ws.cell(row, 1).value:
            return row
    return max(4, ws.max_row)


def _copy_sheet_to_single_workbook(template_path, template_sheet):
    wb = load_workbook(template_path)
    if template_sheet not in wb.sheetnames:
        raise ValueError(f'قالب صدى لا يحتوي على شيت اليوم: {template_sheet}')
    keep = wb[template_sheet]
    for ws in list(wb.worksheets):
        if ws.title != keep.title:
            wb.remove(ws)
    return wb, keep


def _set_cell_value(cell, value):
    cell.value = value
    if value is not None:
        cell.number_format = '#,##0.00'


def build_sada_scales_workbook(tokyo_path, template_path, day_name, output_date, weight_entries):
    day_key = str(day_name or '').strip()
    if day_key not in DAY_OPTIONS:
        raise ValueError('اختار يوم صحيح لموازين صدى')
    day_no, template_sheet = DAY_OPTIONS[day_key]

    tokyo_wb = load_workbook(tokyo_path, data_only=True, keep_vba=False)
    value_index, _diagnostics = _build_tokyo_value_index(tokyo_wb, day_no)
    tokyo_wb.close()

    wb, ws = _copy_sheet_to_single_workbook(template_path, template_sheet)
    if output_date:
        try:
            dt = datetime.strptime(output_date, '%Y-%m-%d')
            ws.title = output_date
            ws['A2'] = dt
            ws['A2'].number_format = 'yyyy-mm-dd'
        except ValueError:
            ws.title = str(output_date)[:31]
            ws['A2'] = output_date

    weight_index, weight_rows = _build_weight_index(weight_entries)
    matched_tokyo = matched_actual = 0
    missing_tokyo = []
    missing_actual = []
    consumed_weight_keys = set()

    for row in range(4, ws.max_row + 1):
        name = ws.cell(row, 1).value
        if not name:
            continue
        batch = _batch_from_any_text(name, rtl_template=True)
        component = _component_for_template_name(name)
        base = _base_for_template_name(name)
        key = _candidate_key(base, batch, component)
        fallback_key = _candidate_key(base, '', component)

        planned = value_index.get(key)
        if planned is None:
            planned = value_index.get(fallback_key)
        if planned is not None:
            _set_cell_value(ws.cell(row, 2), planned)
            matched_tokyo += 1
        else:
            missing_tokyo.append(str(name))

        actual = weight_index.get(key)
        if actual is not None:
            _set_cell_value(ws.cell(row, 3), actual)
            matched_actual += 1
            consumed_weight_keys.add(key)
        else:
            missing_actual.append(str(name))

    extra_rows = [row for key, row in weight_rows.items() if key not in consumed_weight_keys]
    extra_rows.sort(key=lambda r: (_norm_text(r['base']), _batch_sort_value(r['batch']), _component_sort_value(r['component']), _norm_text(r['item_name'])))
    if extra_rows:
        style_row = _last_template_item_row(ws)
        insert_at = style_row + 1
        ws.insert_rows(insert_at, amount=len(extra_rows))
        for offset, item in enumerate(extra_rows):
            target_row = insert_at + offset
            _copy_row_style(ws, style_row, target_row)
            ws.cell(target_row, 1).value = item['item_name']
            ws.cell(target_row, 2).value = None
            _set_cell_value(ws.cell(target_row, 3), item['weight'])
            matched_actual += 1
            missing_tokyo.append(item['item_name'])

    out = BytesIO()
    wb.save(out)
    wb.close()
    out.seek(0)
    return out, {
        'matched_tokyo': matched_tokyo,
        'matched_actual': matched_actual,
        'missing_tokyo_count': len(missing_tokyo),
        'missing_actual_count': len(missing_actual),
        'added_actual_rows': len(extra_rows),
    }
