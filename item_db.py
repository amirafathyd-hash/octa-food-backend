"""
Master item dictionary, now persisted in Supabase (table: master_items) instead of a
local JSON file, so it survives across deployments/sessions and is shared centrally.
"""
import re
import unicodedata
from db import get_client


def normalize_ar(s):
    """Normalize Arabic text for comparison (presentation forms, diacritics, alef/ya variants)."""
    if not s:
        return ''
    s = unicodedata.normalize('NFKC', s)
    s = re.sub(r'[\u064B-\u065F\u0670]', '', s)
    s = s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    s = s.replace('ى', 'ي').replace('ة', 'ه')
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def load_db():
    """Returns dict keyed by item_key -> {section, display_name, ar_main, ar_aliases}"""
    sb = get_client()
    res = sb.table('master_items').select('*').execute()
    db = {}
    for row in res.data:
        db[row['item_key']] = {
            'section': row.get('section'),
            'display_name': row.get('display_name'),
            'ar_main': row.get('ar_main') or '',
            'ar_aliases': row.get('ar_aliases') or [],
        }
    return db


def upsert_item(key, section, display_name, ar_main):
    sb = get_client()
    sb.table('master_items').upsert({
        'item_key': key,
        'section': section,
        'display_name': display_name,
        'ar_main': ar_main,
    }, on_conflict='item_key').execute()


def add_alias(key, alias):
    sb = get_client()
    res = sb.table('master_items').select('ar_aliases').eq('item_key', key).execute()
    aliases = (res.data[0]['ar_aliases'] if res.data else []) or []
    if alias not in aliases:
        aliases.append(alias)
        sb.table('master_items').update({'ar_aliases': aliases}).eq('item_key', key).execute()


def seed_from_order(order_data):
    """Populate/update master_items using items parsed from an order PDF."""
    db = load_db()
    for section in ('salads', 'dressing'):
        for item in order_data.get(section, []):
            key = item['name_en'].strip().upper()
            if key not in db:
                upsert_item(key, section, item['name_en'].strip(), item.get('name_ar', ''))
                db[key] = {
                    'section': section,
                    'display_name': item['name_en'].strip(),
                    'ar_main': item.get('name_ar', ''),
                    'ar_aliases': [],
                }
    return db
