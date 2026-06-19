from rapidfuzz import fuzz, process
from item_db import normalize_ar, add_alias

FUZZY_THRESHOLD = 75


def match_invoice_item(name_ar, db):
    """
    db: dict from item_db.load_db()
    Returns (matched_key, score, method) where method in ('exact','alias','fuzzy','none').
    On a fuzzy match, the new alias is persisted to Supabase automatically.
    """
    norm_target = normalize_ar(name_ar)

    for key, entry in db.items():
        if normalize_ar(entry.get('ar_main', '')) == norm_target:
            return key, 100, 'exact'
        for alias in entry.get('ar_aliases', []):
            if normalize_ar(alias) == norm_target:
                return key, 100, 'alias'

    candidates = {}
    for key, entry in db.items():
        names = [entry.get('ar_main', '')] + entry.get('ar_aliases', [])
        for n in names:
            if n:
                candidates[normalize_ar(n)] = key

    if not candidates:
        return None, 0, 'none'

    best = process.extractOne(norm_target, list(candidates.keys()), scorer=fuzz.token_sort_ratio)
    if best and best[1] >= FUZZY_THRESHOLD:
        matched_norm_name, score, _ = best
        key = candidates[matched_norm_name]
        if name_ar != db[key].get('ar_main') and name_ar not in db[key].get('ar_aliases', []):
            add_alias(key, name_ar)
            db[key].setdefault('ar_aliases', []).append(name_ar)
        return key, score, 'fuzzy'

    return None, best[1] if best else 0, 'none'
