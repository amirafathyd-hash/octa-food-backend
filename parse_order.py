import re
import pdfplumber

UNIT_RE = re.compile(r'(Gram-\S+|ML-\S+)')
NUM_RE = re.compile(r'\d+\.?\d*')


def _parse_line(line, has_box_col):
    """Parse one item line from the order PDF into a dict."""
    m = UNIT_RE.search(line)
    if not m:
        return None
    unit = m.group(1)
    before = line[:m.start()]
    after = line[m.end():]

    # English name is everything before the em-dash '—'
    if '—' not in before:
        return None
    en_name, rest = before.split('—', 1)
    en_name = en_name.strip()

    nums_before = NUM_RE.findall(rest)
    nums_after = NUM_RE.findall(after)

    # Arabic name = rest with numbers stripped out, then char-reversed to fix RTL glyph order
    ar_raw = NUM_RE.sub('', rest).strip()
    name_ar = ar_raw[::-1].strip() if ar_raw else ''

    qty_box = None
    qty_needed = None
    if has_box_col:
        if len(nums_before) >= 2:
            qty_box = float(nums_before[0])
            qty_needed = float(nums_before[1])
        elif len(nums_before) == 1:
            qty_needed = float(nums_before[0])
    else:
        if len(nums_before) >= 1:
            qty_needed = float(nums_before[0])

    current_inventory = float(nums_after[0]) if nums_after else None

    return {
        'name_en': en_name,
        'name_ar': name_ar,
        'qty_box': qty_box,
        'qty_needed': qty_needed,
        'unit': unit,
        'current_inventory': current_inventory,
    }


def parse_order_pdf(path):
    """Returns dict with 'date', 'salads': [...], 'dressing': [...]"""
    result = {'date': None, 'salads': [], 'dressing': []}
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            lines = text.split('\n')
            section = None
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if line.startswith('Date'):
                    dm = re.search(r'Date\s+(.*)', line)
                    if dm:
                        result['date'] = dm.group(1).strip()
                    continue
                if line.startswith('Item for Salads'):
                    section = 'salads'
                    continue
                if line.startswith('Item for Dressing'):
                    section = 'dressing'
                    continue
                if section is None:
                    continue
                item = _parse_line(line, has_box_col=(section == 'salads'))
                if item:
                    item['section'] = section
                    result[section].append(item)
    return result


if __name__ == '__main__':
    import json
    import sys
    data = parse_order_pdf(sys.argv[1])
    print(json.dumps(data, ensure_ascii=False, indent=2))
