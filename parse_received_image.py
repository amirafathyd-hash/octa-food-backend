"""
Parses a photographed/scanned "received" sheet (the printed order form with
handwritten received quantities added next to each item) using EasyOCR.

Strategy:
1. Run EasyOCR once on the full image -> list of (bbox, text, confidence).
2. Identify "item rows" by fuzzy-matching detected text against the known
   English item names from the master item dictionary (these are machine-printed,
   so OCR confidence on them is high and reliable).
3. For each item row, look for handwritten text detected to the RIGHT of the
   item name and at roughly the SAME vertical position (same row) — that's the
   handwritten quantity.
4. Parse the handwritten text into (quantity, unit) using a regex tuned to the
   observed format ("15-K.G", "7-Box", "2-PKT", "6.5-K.G", etc.).
5. Anything that doesn't parse cleanly, or whose OCR confidence is low, is
   flagged as 'needs_review' instead of being silently accepted.
"""
import re
from rapidfuzz import fuzz, process

QTY_RE = re.compile(
    r'(\d+[.,]?\d*)\s*[-=–]?\s*([A-Za-z]{1,4})',
)
DATE_RE = re.compile(r'(\d{1,2})[\s\-]+([A-Za-z]{3,9})[\s\-]+(\d{4})')
MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def _extract_date(detections):
    """Finds a printed date like '3-Apr-2026' among the OCR detections."""
    for _, text, _ in detections:
        m = DATE_RE.search(text.replace(' ', '-'))
        if m:
            day, mon, year = m.groups()
            mon_num = MONTH_MAP.get(mon.lower()[:3])
            if mon_num:
                return f'{int(year):04d}-{mon_num:02d}-{int(day):02d}'
    return None


DASH_DECIMAL_RE = re.compile(r'^(\d+)-(\d+)-([A-Za-z.]+)$')

UNIT_NORMALIZE = {
    'kg': 'KG', 'k': 'KG', 'kc': 'KG', 'kco': 'KG', 'kcm': 'KG', 'kcn': 'KG',
    'box': 'BOX', 'bo': 'BOX', 'boz': 'BOX', 'boa': 'BOX',
    'pkt': 'PKT', 'pk': 'PKT', 'pct': 'PKT', 'pkr': 'PKT',
    'pc': 'PC', 'pcs': 'PC',
}


def _normalize_unit(raw):
    key = re.sub(r'[^a-zA-Z]', '', raw).lower()
    return UNIT_NORMALIZE.get(key, raw.upper() if raw else '')


def _parse_quantity_text(raw_text):
    """Returns (qty, unit, ok) — ok=False if the text doesn't look like a clean
    '<number>-<unit>' pattern."""
    cleaned = raw_text.strip().replace(',', '.').replace('—', '-')

    # handwriting sometimes uses a dash instead of a decimal point, e.g.
    # "1-5-K.G" meaning "1.5 KG" rather than literally "5 KG"
    dm = DASH_DECIMAL_RE.match(cleaned)
    if dm:
        whole, frac, unit_str = dm.groups()
        try:
            qty = float(f'{whole}.{frac}')
        except ValueError:
            return None, None, False
        return qty, _normalize_unit(unit_str), True

    m = QTY_RE.search(cleaned)
    if not m:
        return None, None, False
    qty_str, unit_str = m.groups()
    try:
        qty = float(qty_str.replace(',', '.'))
    except ValueError:
        return None, None, False
    unit = _normalize_unit(unit_str)
    return qty, unit, True


def _get_reader():
    import easyocr
    global _READER
    try:
        return _READER
    except NameError:
        _READER = easyocr.Reader(['en'], gpu=False)
        return _READER


def parse_received_image(image_path, master_items, name_match_threshold=80):
    """
    master_items: dict from item_db.load_db() -> {key: {display_name, ...}}
    Returns: {
        'rows': [ {item_key, name_en, raw_text, qty, unit, confidence, needs_review}, ... ],
        'unmatched_text': [ raw detected strings that looked like quantities but
                             couldn't be linked to any item row ]
    }
    """
    reader = _get_reader()
    detections = reader.readtext(image_path)  # [(bbox, text, conf), ...]

    date_iso = _extract_date(detections)

    item_candidates = []   # (y_center, x_right_edge, item_key, matched_name, conf)
    other_detections = []  # (y_center, x_left_edge, text, conf)

    name_lookup = {info['display_name']: key for key, info in master_items.items() if info.get('display_name')}
    name_list = list(name_lookup.keys())
    name_list_upper = [n.upper() for n in name_list]

    for bbox, text, conf in detections:
        xs = [p[0] for p in bbox]
        ys = [p[1] for p in bbox]
        y_center = sum(ys) / len(ys)
        x_left, x_right = min(xs), max(xs)

        if len(text.strip()) < 3:
            other_detections.append((y_center, x_left, text, conf))
            continue

        match = process.extractOne(text.strip().upper(), name_list_upper, scorer=fuzz.partial_ratio)
        if match and match[1] >= name_match_threshold:
            matched_name = name_list[name_list_upper.index(match[0])]
            item_candidates.append((y_center, x_right, name_lookup[matched_name], matched_name, conf))
        else:
            other_detections.append((y_center, x_left, text, conf))

    rows = []
    used_other_idx = set()
    ROW_TOLERANCE_RATIO = 0.6

    sorted_items = sorted(item_candidates, key=lambda r: r[0])
    if len(sorted_items) >= 2:
        diffs = [sorted_items[i + 1][0] - sorted_items[i][0] for i in range(len(sorted_items) - 1)]
        diffs = [d for d in diffs if d > 2]
        row_height = sorted(diffs)[len(diffs) // 2] if diffs else 20
    else:
        row_height = 20

    for y_center, x_right, item_key, matched_name, conf in sorted_items:
        best = None
        best_dist = None
        for i, (oy, ox, text, oconf) in enumerate(other_detections):
            if i in used_other_idx:
                continue
            if ox < x_right - 5:
                continue
            dy = abs(oy - y_center)
            if dy > row_height * ROW_TOLERANCE_RATIO:
                continue
            if best is None or dy < best_dist:
                best, best_dist = (i, oy, ox, text, oconf), dy

        if best:
            idx, oy, ox, text, oconf = best
            used_other_idx.add(idx)
            qty, unit, ok = _parse_quantity_text(text)
            rows.append({
                'item_key': item_key,
                'name_en': matched_name,
                'raw_text': text,
                'qty': qty,
                'unit': unit,
                'confidence': round(float(oconf) * 100, 1),
                'needs_review': (not ok) or oconf < 0.5,
            })

    unmatched_text = [text for i, (oy, ox, text, oconf) in enumerate(other_detections)
                       if i not in used_other_idx and _parse_quantity_text(text)[2]]

    return {'date': date_iso, 'rows': rows, 'unmatched_text': unmatched_text}


if __name__ == '__main__':
    import sys
    import json
    from item_db import load_db
    db = load_db()
    result = parse_received_image(sys.argv[1], db)
    print(json.dumps(result, ensure_ascii=False, indent=2))
