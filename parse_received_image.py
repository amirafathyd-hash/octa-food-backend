"""
Parses a photographed/scanned "received" sheet (the printed order form with
handwritten received quantities added next to each item) using Claude Vision
(Anthropic API) instead of OCR.space + regex heuristics.

Why the switch: OCR.space + positional word-matching struggled badly with
handwritten Arabic/English quantities and unit abbreviations (misreads like
"kco"/"boz" needing huge normalization tables, Cyrillic-lookalike glyphs in
month names, etc). Claude Vision reads the handwriting + printed item list
together in one pass and returns structured matches directly, the same
approach already proven reliable in the Telegram scale-reading bot.

Public contract (kept identical to the old OCR.space version so app.py needs
NO changes):

    parse_received_image(path, db) -> {
        'date': 'YYYY-MM-DD' | None,
        'rows': [
            {
                'item_key': str,       # matches a key in `db`
                'qty': float | None,
                'unit': str,           # one of VALID_UNITS, or '' if unknown
                'needs_review': bool,
                'raw_text': str,       # what Claude actually read, for debugging
                'name_en': str,
                'confidence': int,     # 0-100
            }, ...
        ]
    }
"""
import os
import re
import json
import base64
import unicodedata
import difflib
from datetime import datetime

import requests

ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY', '')
ANTHROPIC_MODEL = os.environ.get('ANTHROPIC_VISION_MODEL', 'claude-sonnet-5')
ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages'

VALID_UNITS = {'KG', 'GM', 'BOX', 'PKT', 'PC', 'TRAY', 'BTL', 'LTR', 'ML'}
CONFIDENCE_MIN = int(os.environ.get('RECEIVED_CONFIDENCE_MIN', '70'))

VISION_PROMPT = """\
دي صورة "ورقة أوردر يومي" مطبوعة، وجنب كل صنف اتكتبت كمية "مستلمة" بخط اليد
(غالبًا في عمود على اليمين اسمه Qty Received أو جنب الصنف مباشرة).

اقرأ التاريخ المطبوع فوق الورقة (لو موجود)، وبعدين لكل صف فيه رقم مكتوب بخط
اليد بجانب الصنف، ارجع:
- الاسم العربي للصنف زي ما هو مطبوع بالظبط
- الاسم الإنجليزي للصنف زي ما هو مطبوع بالظبط
- الكمية المكتوبة بخط اليد (رقم بس)
- الوحدة لو مكتوبة (KG / BOX / PKT / PC / TRAY / BTL / LTR)، أو فاضي لو مش واضحة
- confidence من 0 لـ100 لمدى وضوح الخط

تجاهل الصفوف اللي معندهاش أي رقم مكتوب بخط اليد بجانبها خالص.
لو مش متأكد من رقم معين، سيبه في النتيجة لكن حط confidence أقل من 60.

ارجع JSON فقط بالشكل ده، من غير أي شرح أو نص تاني:
{
  "date": "YYYY-MM-DD or empty string if not visible",
  "rows": [
    {"name_ar": "...", "name_en": "...", "qty": 12.5, "unit": "KG", "confidence": 90}
  ]
}
"""


# ---------------------------------------------------------------------------
# Arabic normalization + matching (same proven approach as the invoice matcher)
# ---------------------------------------------------------------------------
def _normalize_ar(s):
    if not s:
        return ''
    s = unicodedata.normalize('NFKC', s)
    s = re.sub(r'[\u064B-\u065F\u0670]', '', s)
    s = s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    s = s.replace('ى', 'ي').replace('ة', 'ه')
    s = re.sub(r'[^\u0600-\u06FF\s]', '', s)
    return re.sub(r'\s+', ' ', s).strip()


def _token_set(s):
    return set(_normalize_ar(s).split())


def _match_db_item(name_ar, name_en, db):
    """Match an item read off the photo against the catalog `db`
    (dict[item_key] -> {section, display_name, ar_main, ar_aliases})."""
    best_key, best_score = None, 0.0

    # 1) try exact/near English key match first (cheap + very reliable when present)
    if name_en:
        key_guess = name_en.strip().upper()
        if key_guess in db:
            return key_guess

    # 2) Arabic token-overlap matching (mirrors the invoice matcher)
    ar_tokens = _token_set(name_ar)
    if ar_tokens:
        for key, item in db.items():
            candidates = [item.get('ar_main', '')] + list(item.get('ar_aliases') or [])
            for cand in candidates:
                cand_tokens = _token_set(cand)
                if not cand_tokens:
                    continue
                overlap = len(ar_tokens & cand_tokens)
                if overlap == 0:
                    continue
                jaccard = overlap / len(ar_tokens | cand_tokens)
                if jaccard > best_score:
                    best_score = jaccard
                    best_key = key

    # 3) fallback: fuzzy English name match
    if not best_key and name_en:
        display_names = {k: v.get('display_name', k) for k, v in db.items()}
        matches = difflib.get_close_matches(name_en.upper(), display_names.values(), n=1, cutoff=0.7)
        if matches:
            for k, v in display_names.items():
                if v == matches[0]:
                    best_key = k
                    break

    return best_key


def _normalize_unit(u):
    u = (u or '').strip().upper()
    return u if u in VALID_UNITS else ''


def _normalize_date(d):
    if not d:
        return None
    d = d.strip()
    for fmt in ('%Y-%m-%d', '%d-%b-%Y', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(d, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Anthropic Vision call
# ---------------------------------------------------------------------------
def _call_claude_vision(image_path):
    if not ANTHROPIC_API_KEY:
        raise RuntimeError('ANTHROPIC_API_KEY غير موجود في متغيرات البيئة (Railway env vars)')

    with open(image_path, 'rb') as f:
        img_bytes = f.read()

    ext = os.path.splitext(image_path)[1].lower()
    media_type = 'image/png' if ext == '.png' else 'image/jpeg'

    payload = {
        'model': ANTHROPIC_MODEL,
        'max_tokens': 2000,
        'messages': [{
            'role': 'user',
            'content': [
                {'type': 'image', 'source': {'type': 'base64', 'media_type': media_type,
                                              'data': base64.b64encode(img_bytes).decode('ascii')}},
                {'type': 'text', 'text': VISION_PROMPT},
            ],
        }],
    }
    headers = {
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
    }
    resp = requests.post(ANTHROPIC_URL, headers=headers, json=payload, timeout=60)
    if resp.status_code != 200:
        raise RuntimeError(f'Anthropic HTTP {resp.status_code}: {resp.text[:500]}')

    data = resp.json()
    text = ''.join(block.get('text', '') for block in data.get('content', []) if block.get('type') == 'text')

    match = re.search(r'\{.*\}', text, re.DOTALL)
    if not match:
        raise RuntimeError('Claude لم يرجّع JSON صالح: ' + text[:300])
    return json.loads(match.group(0))


# ---------------------------------------------------------------------------
# Public entry point (drop-in replacement)
# ---------------------------------------------------------------------------
def parse_received_image(path, db):
    ai_result = _call_claude_vision(path)

    date_iso = _normalize_date(ai_result.get('date'))
    rows_out = []

    for r in ai_result.get('rows', []):
        name_ar = (r.get('name_ar') or '').strip()
        name_en = (r.get('name_en') or '').strip()
        qty = r.get('qty')
        try:
            qty = float(qty) if qty not in (None, '') else None
        except (TypeError, ValueError):
            qty = None
        unit = _normalize_unit(r.get('unit'))
        confidence = int(r.get('confidence') or 0)

        item_key = _match_db_item(name_ar, name_en, db)
        needs_review = (
            item_key is None
            or qty is None
            or confidence < CONFIDENCE_MIN
        )

        rows_out.append({
            'item_key': item_key or (name_en.strip().upper() if name_en else 'UNKNOWN'),
            'qty': qty,
            'unit': unit,
            'needs_review': needs_review,
            'raw_text': f"{name_ar} / {name_en} -> {qty} {unit}".strip(),
            'name_en': name_en or (db.get(item_key, {}).get('display_name', '') if item_key else ''),
            'confidence': confidence,
        })

    return {'date': date_iso, 'rows': rows_out}
