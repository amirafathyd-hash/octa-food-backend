"""
Parses a photographed/scanned "received" sheet (the printed order form with
handwritten received quantities added next to each item) using the OCR.space
cloud API (free tier, no credit card, Engine 2 works on free keys).

Strategy:
1. Browser calls OCR.space with isOverlayRequired=true to get word-level positions.
2. Identify "item rows" by fuzzy-matching detected text against known English item names.
3. For each item row, find handwritten text to the RIGHT at roughly the SAME vertical position.
4. Parse that handwritten text into (quantity, unit).
5. Only flag as needs_review when the quantity number itself couldn't be read at all.
"""
import os
import re
import requests
from rapidfuzz import fuzz, process

OCR_SPACE_API_KEY = os.environ.get('OCR_SPACE_API_KEY')
OCR_SPACE_URL = 'https://api.ocr.space/parse/image'

QTY_RE = re.compile(r'(\d+[.,]?\d*)\s*[-=–]?\s*([A-Za-z]{0,4})')
DASH_DECIMAL_RE = re.compile(r'^(\d+)-(\d+)-?([A-Za-z.]*)$')
# Also match pure numbers: "15", "7.5", "2,5"
NUM_ONLY_RE = re.compile(r'^(\d+[.,]?\d*)$')

UNIT_NORMALIZE = {
    # KG and common OCR misreads
    'kg': 'KG', 'k': 'KG', 'kc': 'KG', 'kco': 'KG', 'kcm': 'KG', 'kcn': 'KG',
    'ko': 'KG', 'kq': 'KG', 'kd': 'KG', 'kp': 'KG', 'kn': 'KG', 'kr': 'KG',
    'ky': 'KG', 'km': 'KG', 'kv': 'KG', 'kz': 'KG', 'kl': 'KG',
    'kkg': 'KG', 'kgs': 'KG', 'kge': 'KG', 'kg': 'KG',
    # BOX variants
    'box': 'BOX', 'bo': 'BOX', 'boz': 'BOX', 'boa': 'BOX', 'boo': 'BOX',
    'bx': 'BOX', 'bax': 'BOX', 'bix': 'BOX', 'b': 'BOX', 'boes': 'BOX',
    'boxs': 'BOX', 'boox': 'BOX', 'boks': 'BOX',
    # PKT variants
    'pkt': 'PKT', 'pk': 'PKT', 'pct': 'PKT', 'pkr': 'PKT', 'pt': 'PKT',
    'pkt': 'PKT', 'pkts': 'PKT', 'pks': 'PKT', 'pkt': 'PKT',
    'pck': 'PKT', 'pckt': 'PKT', 'pakt': 'PKT',
    # PC variants
    'pc': 'PC', 'pcs': 'PC', 'p': 'PC', 'pce': 'PC', 'piec': 'PC',
    # TRY / TRAY
    'tray': 'TRAY', 'try': 'TRAY', 'tr': 'TRAY', 'tra': 'TRAY',
    # BTL / Bottle
    'btl': 'BTL', 'bt': 'BTL', 'bot': 'BTL', 'bottle': 'BTL',
    # Liter
    'ltr': 'LTR', 'lt': 'LTR', 'l': 'LTR', 'liter': 'LTR', 'litr': 'LTR',
}

VALID_UNITS = {'KG', 'BOX', 'PKT', 'PC', 'TRAY', 'BTL', 'LTR'}

# --- Date extraction: supports BOTH text-month and numeric formats ---
# الفواصل بتشمل الشرطة العادية والشرطات الطويلة اللي ممكن الـOCR يرجّعها بدلها (– —)
SEP = r'[\s\-–—/]+'
# ملحوظة مهمة: OCR.space أحيانًا بيقرا حروف "Apr" كحروف سيريلية (روسية) شبه
# مطابقة بصريًا (زي "Арт" بدل "Apr") خصوصًا مع خطوط معينة أو صور خط يد. لو
# قصرنا الـregex على [A-Za-z] بس، الكلمة دي مش هتتطابق خالص وهيضيع التاريخ كله.
# فبنوسّع الـregex يقبل أي حروف (مش أرقام ولا فواصل)، وبعدين نحوّل الحروف
# السيريلية الشبيهة بصريًا للاتيني قبل ما نقارن باسم الشهر.
MONTH_TOKEN_RE = r'([^\d\s\-–—/]{3,9})'
DATE_RE_TEXT = re.compile(r'(\d{1,2})' + SEP + MONTH_TOKEN_RE + SEP + r'(\d{4})')
DATE_RE_NUMERIC = re.compile(r'(\d{1,2})[\s\-–—/.](\d{1,2})[\s\-–—/.](\d{2,4})')
MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}
MONTH_ABBRS = list(MONTH_MAP.keys())

# خريطة الحروف السيريلية (كبيرة وصغيرة) اللي شكلها مطابق بصريًا لحروف لاتينية،
# وبتظهر غلط في نتائج OCR لأسماء الشهور الإنجليزية أحيانًا.
CYRILLIC_TO_LATIN = str.maketrans({
    'А': 'A', 'В': 'B', 'Е': 'E', 'К': 'K', 'М': 'M', 'Н': 'H', 'О': 'O',
    'Р': 'P', 'С': 'C', 'Т': 'T', 'У': 'Y', 'Х': 'X',
    'а': 'a', 'в': 'b', 'е': 'e', 'к': 'k', 'м': 'm', 'н': 'h', 'о': 'o',
    'р': 'p', 'с': 'c', 'т': 't', 'у': 'y', 'х': 'x',
})


def _match_month(token):
    """بيدّور على اسم الشهر مع تسامح غلطة حرف واحد من الـOCR (مثلاً Mor بدل Mar)،
    وكمان بيحوّل أي حروف سيريلية شبيهة بصريًا للاتيني الأول."""
    normalized = token.translate(CYRILLIC_TO_LATIN)
    key = normalized.lower()[:3]
    if key in MONTH_MAP:
        return MONTH_MAP[key]
    match = process.extractOne(key, MONTH_ABBRS, scorer=fuzz.ratio)
    if match and match[1] >= 60:  # يسمح بفرق حرف واحد في اسم الشهر المختصر (3 حروف)
        return MONTH_MAP[match[0]]
    return None


def _extract_date(all_text):
    # Try text-month format first: "4 Apr 2026", "4-Apr-2026"
    m = DATE_RE_TEXT.search(all_text.replace(' ', '-'))
    if m:
        day, mon, year = m.groups()
        mon_num = _match_month(mon)
        if mon_num:
            return f'{int(year):04d}-{mon_num:02d}-{int(day):02d}'

    # Try numeric format: "4/4/2026", "04-04-2026", "4.4.2026"
    for m in DATE_RE_NUMERIC.finditer(all_text):
        d, mo, y = m.groups()
        d, mo, y = int(d), int(mo), int(y)
        if y < 100:
            y += 2000
        # Validate: day 1-31, month 1-12, year reasonable
        if 1 <= d <= 31 and 1 <= mo <= 12 and 2020 <= y <= 2030:
            return f'{y:04d}-{mo:02d}-{d:02d}'

    return None


def _normalize_unit(raw):
    key = re.sub(r'[^a-zA-Z]', '', raw).lower()
    return UNIT_NORMALIZE.get(key, raw.upper() if raw else '')


def _parse_quantity_text(raw_text):
    """
    Parse handwritten quantity text.
    Returns (qty, unit, ok) where:
    - ok=True  → confident parse, no review needed
    - ok=False → couldn't extract a meaningful number at all
    Note: unknown/ambiguous units still return ok=True if the number was found —
    the human reviewer only needs to see rows where the NUMBER itself is missing.
    """
    cleaned = raw_text.strip().replace(',', '.').replace('—', '-').replace('،', '.')

    # Pattern: "1-5-KG" meaning 1.5 KG
    dm = DASH_DECIMAL_RE.match(cleaned)
    if dm:
        whole, frac, unit_str = dm.groups()
        try:
            qty = float(f'{whole}.{frac}')
            unit = _normalize_unit(unit_str) if unit_str else ''
            return qty, unit, True  # number found → ok
        except ValueError:
            pass

    # Pattern: "15-KG", "7 Box", "2.5KG"
    m = QTY_RE.search(cleaned)
    if m:
        qty_str, unit_str = m.groups()
        try:
            qty = float(qty_str.replace(',', '.'))
            unit = _normalize_unit(unit_str) if unit_str else ''
            return qty, unit, True  # number found → ok
        except ValueError:
            pass

    # Pattern: pure number "15", "7.5" (no unit — unit inferred later)
    nm = NUM_ONLY_RE.match(cleaned.split()[0]) if cleaned.split() else None
    if nm:
        try:
            qty = float(nm.group(1).replace(',', '.'))
            return qty, '', True  # number found → ok
        except ValueError:
            pass

    # Nothing useful found
    return None, None, False


def _call_ocr_space(image_path):
    if not OCR_SPACE_API_KEY:
        raise RuntimeError('لازم تحدد OCR_SPACE_API_KEY في متغيرات البيئة (Environment Variables)')

    last_error = None
    for attempt in range(2):
        try:
            with open(image_path, 'rb') as f:
                resp = requests.post(
                    OCR_SPACE_URL,
                    files={'file': f},
                    data={
                        'apikey': OCR_SPACE_API_KEY,
                        'OCREngine': 2,
                        'isOverlayRequired': True,
                        'language': 'eng',
                        'scale': True,
                        'isTable': True,
                    },
                    timeout=(10, 45),
                )
            resp.raise_for_status()
            data = resp.json()
            if data.get('IsErroredOnProcessing'):
                raise RuntimeError(data.get('ErrorMessage', ['OCR.space error'])[0])
            return data
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            last_error = e
            continue
    raise RuntimeError(f'فشل الاتصال بخدمة OCR.space بعد محاولتين: {last_error}')


def process_ocr_data(data, master_items, name_match_threshold=72):
    """
    Takes an already-fetched OCR.space response (dict) and does row-matching/parsing.
    Threshold lowered from 80→72 to catch more item names despite OCR noise.
    needs_review is now only True when the quantity NUMBER couldn't be read at all.
    """
    parsed_results = data.get('ParsedResults', [])
    if not parsed_results:
        return {'date': None, 'rows': []}

    full_text = ' '.join(pr.get('ParsedText', '') for pr in parsed_results)
    date_iso = _extract_date(full_text)

    words = []
    for pr in parsed_results:
        overlay = pr.get('TextOverlay', {})
        for line in overlay.get('Lines', []):
            for w in line.get('Words', []):
                text = w.get('WordText', '').strip()
                if not text:
                    continue
                left, top = w.get('Left', 0), w.get('Top', 0)
                width, height = w.get('Width', 0), w.get('Height', 0)
                words.append((top + height / 2, left, left + width, text))

    name_lookup = {info['display_name']: key for key, info in master_items.items() if info.get('display_name')}
    name_list = list(name_lookup.keys())
    name_list_upper = [n.upper() for n in name_list]

    received_col_x = None
    for y, xl, xr, text in sorted(words, key=lambda w: w[0]):
        if text.strip().lower() in ('received', 'recieved', 'rec', 'recv'):
            received_col_x = xl
            break
    MIN_GAP_FROM_NAME = 5

    item_candidates = []
    other_words = []

    for i, (y, xl, xr, text) in enumerate(words):
        candidate_texts = [text]
        if i + 1 < len(words) and abs(words[i + 1][0] - y) < 8:
            candidate_texts.append(text + ' ' + words[i + 1][3])

        best_match, best_score = None, 0
        for ct in candidate_texts:
            match = process.extractOne(ct.upper(), name_list_upper, scorer=fuzz.partial_ratio)
            if match and match[1] > best_score:
                best_match, best_score = match[0], match[1]

        if best_match and best_score >= name_match_threshold and len(text) >= 3:
            matched_name = name_list[name_list_upper.index(best_match)]
            item_candidates.append((y, xr, name_lookup[matched_name], matched_name))
        else:
            other_words.append((y, xl, text))

    rows = []
    used_idx = set()
    sorted_items = sorted(set(item_candidates), key=lambda r: r[0])

    n = len(sorted_items)
    bands = []
    for idx in range(n):
        y_center = sorted_items[idx][0]
        upper_bound = (-1e9 if idx == 0 else (sorted_items[idx - 1][0] + y_center) / 2)
        lower_bound = (1e9 if idx == n - 1 else (y_center + sorted_items[idx + 1][0]) / 2)
        bands.append((upper_bound, lower_bound))

    Y_BIAS_CORRECTION = -8

    for idx, (y_center, x_right, item_key, matched_name) in enumerate(sorted_items):
        upper_bound, lower_bound = bands[idx]
        min_x = max(x_right + MIN_GAP_FROM_NAME, received_col_x - 20) if received_col_x else x_right + MIN_GAP_FROM_NAME
        candidates = []
        for i, (oy, ox, text) in enumerate(other_words):
            if i in used_idx or ox < min_x:
                continue
            corrected_y = oy + Y_BIAS_CORRECTION
            if upper_bound <= corrected_y < lower_bound:
                candidates.append((i, ox, text))

        if not candidates:
            continue
        candidates.sort(key=lambda c: c[1])
        combined_text = ' '.join(c[2] for c in candidates)
        for i, _, _ in candidates:
            used_idx.add(i)

        qty, unit, ok = _parse_quantity_text(combined_text)
        rows.append({
            'item_key': item_key,
            'name_en': matched_name,
            'raw_text': combined_text,
            'qty': qty,
            'unit': unit,
            'confidence': 90.0 if ok else 30.0,
            'needs_review': not ok,  # only flag when NUMBER itself not found
        })

    return {'date': date_iso, 'rows': rows}


def parse_received_image(image_path, master_items, name_match_threshold=72):
    """Server-side convenience wrapper for local testing/CLI use."""
    data = _call_ocr_space(image_path)
    return process_ocr_data(data, master_items, name_match_threshold)


if __name__ == '__main__':
    import sys
    import json
    from item_db import load_db
    db = load_db()
    result = parse_received_image(sys.argv[1], db)
    print(json.dumps(result, ensure_ascii=False, indent=2))
