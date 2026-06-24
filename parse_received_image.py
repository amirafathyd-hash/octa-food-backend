"""
Parses a photographed/scanned "received" sheet (the printed order form with
handwritten received quantities added next to each item) using the OCR.space
cloud API (free tier, no credit card, Engine 3 supports handwriting).

Why a cloud API instead of a local model (EasyOCR/Tesseract):
- Local deep-learning OCR models (EasyOCR + PyTorch) need ~1-2GB RAM just to
  load, which exceeds the 1GB cap on free hosting tiers (Railway free trial)
  and crashes the server (OOM/SIGKILL).
- OCR.space runs the heavy model on their servers; our server just sends the
  image and gets text + coordinates back, so the lightweight free tier is fine.
- Tesseract (the free *local* alternative) was tested directly against real
  samples and was not accurate enough for this cursive handwriting.

Strategy:
1. Send the image to OCR.space with OCREngine=3 (best for handwriting) and
   isOverlayRequired=true to get word-level text + (x, y) positions.
2. Identify "item rows" by fuzzy-matching detected text against the known
   English item names from the master item dictionary (these are
   machine-printed, so OCR confidence on them is reliable).
3. For each item row, find handwritten text detected to the RIGHT of the item
   name and at roughly the SAME vertical position (same row) — the quantity.
4. Parse that handwritten text into (quantity, unit) with a regex tuned to the
   observed format ("15-K.G", "7-Box", "2-PKT", "6.5-K.G", "1-5-K.G" meaning
   1.5, etc.).
5. Anything that doesn't parse cleanly is flagged as 'needs_review' rather
   than being silently accepted.
"""
import os
import re
import requests
from rapidfuzz import fuzz, process

OCR_SPACE_API_KEY = os.environ.get('OCR_SPACE_API_KEY')
OCR_SPACE_URL = 'https://api.ocr.space/parse/image'

QTY_RE = re.compile(r'(\d+[.,]?\d*)\s*[-=–]?\s*([A-Za-z]{1,4})')
DASH_DECIMAL_RE = re.compile(r'^(\d+)-(\d+)-([A-Za-z.]+)$')

UNIT_NORMALIZE = {
    'kg': 'KG', 'k': 'KG', 'kc': 'KG', 'kco': 'KG', 'kcm': 'KG', 'kcn': 'KG',
    'box': 'BOX', 'bo': 'BOX', 'boz': 'BOX', 'boa': 'BOX',
    'pkt': 'PKT', 'pk': 'PKT', 'pct': 'PKT', 'pkr': 'PKT',
    'pc': 'PC', 'pcs': 'PC',
}

DATE_RE = re.compile(r'(\d{1,2})[\s\-]+([A-Za-z]{3,9})[\s\-]+(\d{4})')
MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def _extract_date(all_text):
    m = DATE_RE.search(all_text.replace(' ', '-'))
    if m:
        day, mon, year = m.groups()
        mon_num = MONTH_MAP.get(mon.lower()[:3])
        if mon_num:
            return f'{int(year):04d}-{mon_num:02d}-{int(day):02d}'
    return None


def _normalize_unit(raw):
    key = re.sub(r'[^a-zA-Z]', '', raw).lower()
    return UNIT_NORMALIZE.get(key, raw.upper() if raw else '')


VALID_UNITS = {'KG', 'BOX', 'PKT', 'PC'}


def _parse_quantity_text(raw_text):
    cleaned = raw_text.strip().replace(',', '.').replace('—', '-')
    dm = DASH_DECIMAL_RE.match(cleaned)
    if dm:
        whole, frac, unit_str = dm.groups()
        try:
            qty = float(f'{whole}.{frac}')
        except ValueError:
            return None, None, False
        unit = _normalize_unit(unit_str)
        return qty, unit, unit in VALID_UNITS

    m = QTY_RE.search(cleaned)
    if not m:
        return None, None, False
    qty_str, unit_str = m.groups()
    try:
        qty = float(qty_str.replace(',', '.'))
    except ValueError:
        return None, None, False
    unit = _normalize_unit(unit_str)
    return qty, unit, unit in VALID_UNITS


def _call_ocr_space(image_path):
    if not OCR_SPACE_API_KEY:
        raise RuntimeError('لازم تحدد OCR_SPACE_API_KEY في متغيرات البيئة (Environment Variables)')

    last_error = None
    for attempt in range(2):  # try once, retry once on a transient network failure
        try:
            with open(image_path, 'rb') as f:
                resp = requests.post(
                    OCR_SPACE_URL,
                    files={'file': f},
                    data={
                        'apikey': OCR_SPACE_API_KEY,
                        'OCREngine': 3,            # best engine for handwriting
                        'isOverlayRequired': True,  # get word-level (x, y) positions
                        'language': 'eng',
                    },
                    timeout=(10, 45),  # (connect timeout, read timeout) — fail fast instead of hanging
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


def parse_received_image(image_path, master_items, name_match_threshold=80):
    """
    master_items: dict from item_db.load_db() -> {key: {display_name, ...}}
    Returns: {
        'date': 'YYYY-MM-DD' or None,
        'rows': [ {item_key, name_en, raw_text, qty, unit, confidence, needs_review}, ... ],
    }
    """
    data = _call_ocr_space(image_path)
    parsed_results = data.get('ParsedResults', [])
    if not parsed_results:
        return {'date': None, 'rows': []}

    full_text = ' '.join(pr.get('ParsedText', '') for pr in parsed_results)
    date_iso = _extract_date(full_text)

    words = []  # (y_center, x_left, x_right, text)
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

    # Find the printed "Quantity Received" column header so we know where the
    # handwriting column actually starts. Without this, the nearest text to the
    # right of an item name is often the PRINTED "Qty Needed" number (which sits
    # closer to the item name than the handwritten value further right) and we'd
    # wrongly copy that instead of the real handwritten quantity.
    received_col_x = None
    for y, xl, xr, text in sorted(words, key=lambda w: w[0]):
        if text.strip().lower() in ('received', 'recieved'):
            received_col_x = xl
            break
    # fallback: if we couldn't find the header at all, don't filter by it
    MIN_GAP_FROM_NAME = 5

    item_candidates = []   # (y_center, x_right_edge, item_key, matched_name)
    other_words = []       # (y_center, x_left, text)

    # group words into short n-grams (1-3 words) to match multi-word item names too
    for i, (y, xl, xr, text) in enumerate(words):
        candidate_texts = [text]
        if i + 1 < len(words) and abs(words[i + 1][0] - y) < 8:
            candidate_texts.append(text + ' ' + words[i + 1][3])

        best_match, best_score, best_span = None, 0, None
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

    # Assign each item a vertical "band" bounded by the midpoints to its
    # neighbors above/below, instead of just picking the nearest text by raw
    # distance. This avoids a handwritten value bleeding into the wrong row
    # when two printed rows sit close together (tall handwriting strokes can
    # otherwise end up nearer to the row above or below than to their own).
    n = len(sorted_items)
    bands = []
    for idx in range(n):
        y_center = sorted_items[idx][0]
        upper_bound = (-1e9 if idx == 0
                       else (sorted_items[idx - 1][0] + y_center) / 2)
        lower_bound = (1e9 if idx == n - 1
                        else (y_center + sorted_items[idx + 1][0]) / 2)
        bands.append((upper_bound, lower_bound))

    Y_BIAS_CORRECTION = -8  # handwriting tends to sit a few px lower than its printed row

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
            'confidence': 90.0 if ok else 40.0,  # OCR.space doesn't give per-word confidence on free tier
            'needs_review': not ok,
        })

    return {'date': date_iso, 'rows': rows}


if __name__ == '__main__':
    import sys
    import json
    from item_db import load_db
    db = load_db()
    result = parse_received_image(sys.argv[1], db)
    print(json.dumps(result, ensure_ascii=False, indent=2))
