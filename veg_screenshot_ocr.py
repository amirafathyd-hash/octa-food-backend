"""
استخراج جدول أصناف الخضار من صور سكرين شوت (زي اللي بتوصل على واتساب) باستخدام
Claude Vision - نفس النمط المستخدم بالظبط في parse_received_image.py.

الصورة الواحدة ممكن تحتوي على أكتر من "بلوك" ملوّن (كل بلوك رسالة واتساب
منفصلة)، كل بلوك فيه كذا صف: اسم الصنف EN - AR، الفئة (خضروات)، الكمية،
والوحدة (Kg/KG/kg/Pack/Box/gm). بنتجاهل وقت الرسالة وعلامة "✓✓" بتاعة
واتساب لأنهم مش جزء من البيانات.
"""
import os
import re
import json
import base64

import requests

ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY', '')
ANTHROPIC_MODEL = (
    os.environ.get('ANTHROPIC_VISION_MODEL')
    or os.environ.get('ANTHROPIC_MODEL')
    or 'claude-sonnet-4-20250514'
)
ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages'

VALID_UNITS = {'KG', 'GM', 'PACK', 'BOX', 'PC', 'TRAY', 'BTL', 'LTR'}

VISION_PROMPT = """\
دي صورة فيها سكرين شوت لجدول أصناف خضار (ممكن يكون فيها أكتر من رسالة واتساب
مجمّعة فوق بعض بألوان خلفية مختلفة). كل صف فيه:
- اسم الصنف بالإنجليزي والعربي مع بعض، غالبًا بالشكل "English Name - الاسم العربي"
- الفئة (زي "خضروات") - تجاهلها، مش محتاجها
- كمية (رقم عشري)
- وحدة القياس (KG أو Kg أو kg أو Pack أو Box أو gm أو PC)

تجاهل تمامًا: وقت الرسالة (زي "7:28PM")، وعلامات صح الواتساب (✓✓)، وأي نص
مش صف بيانات فعلي.

اقرأ كل صف موجود في الصورة (من كل البلوكات، مش بس بلوك واحد) وارجعهم في
مصفوفة واحدة. لو نفس الصنف اتكرر في الصورة (حتى لو في بلوكات مختلفة)، رجّعه
كصفين منفصلين برضه - الدمج هنعمله إحنا بعدين، انت بس اقرأ كل صف زي ما هو.

ارجع JSON فقط بالشكل ده، من غير أي شرح أو نص تاني:
{
  "rows": [
    {"name_en": "Tomato", "name_ar": "طماطم", "qty": 20.576, "unit": "KG"}
  ]
}
"""


def _normalize_unit(u):
    u = (u or '').strip().upper()
    if u in ('KG', 'KGS', 'KILO', 'KILOGRAM'):
        return 'KG'
    if u in ('GM', 'GRAM', 'GRAMS', 'G'):
        return 'GM'
    if u in ('PACK', 'PKT', 'PCK'):
        return 'PACK'
    if u in ('BOX', 'BOXES'):
        return 'BOX'
    if u in ('PC', 'PCS', 'PIECE', 'PIECES'):
        return 'PC'
    if u in ('TRAY', 'TRAYS'):
        return 'TRAY'
    if u in ('BTL', 'BOTTLE'):
        return 'BTL'
    if u in ('LTR', 'LITER', 'L'):
        return 'LTR'
    return u if u in VALID_UNITS else (u or 'UNKNOWN')


def _call_claude_vision(image_path):
    if not ANTHROPIC_API_KEY:
        raise RuntimeError('ANTHROPIC_API_KEY غير موجود في متغيرات البيئة (Railway env vars)')

    with open(image_path, 'rb') as f:
        img_bytes = f.read()

    ext = os.path.splitext(image_path)[1].lower()
    media_type = 'image/png' if ext == '.png' else 'image/jpeg'

    payload = {
        'model': ANTHROPIC_MODEL,
        'max_tokens': 4000,
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
    resp = requests.post(ANTHROPIC_URL, headers=headers, json=payload, timeout=90)
    if resp.status_code != 200:
        try:
            err = resp.json().get('error', {}).get('message') or resp.text
        except Exception:
            err = resp.text
        raise RuntimeError(f'Anthropic HTTP {resp.status_code}: {err[:500]}')

    data = resp.json()
    text = ''.join(block.get('text', '') for block in data.get('content', []) if block.get('type') == 'text')

    match = re.search(r'\{.*\}', text, re.DOTALL)
    if not match:
        raise RuntimeError('Claude لم يرجّع JSON صالح: ' + text[:300])
    return json.loads(match.group(0))


def extract_vegetable_rows(image_path):
    """بترجع list من {'name_en','name_ar','qty','unit'} من صورة واحدة.
    الأرقام والوحدات بتتنضّف، بس الأسماء بتفضل زي ما قراها Claude بالظبط
    (التطبيع/الدمج بيحصل بعدين في مرحلة التجميع مش هنا)."""
    ai_result = _call_claude_vision(image_path)
    rows_out = []
    for r in ai_result.get('rows', []):
        name_en = (r.get('name_en') or '').strip()
        name_ar = (r.get('name_ar') or '').strip()
        if not name_en and not name_ar:
            continue
        qty = r.get('qty')
        try:
            qty = float(qty) if qty not in (None, '') else 0.0
        except (TypeError, ValueError):
            qty = 0.0
        unit = _normalize_unit(r.get('unit'))
        rows_out.append({
            'name_en': name_en, 'name_ar': name_ar,
            'qty': qty, 'unit': unit,
        })
    return rows_out
