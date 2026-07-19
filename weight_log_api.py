# ============================================================
# Octa Food — Backend: تسجيل وأرشفة أوزان الميزان (weight-log)
# ============================================================
#
# الفيتشر ده كان ناقص بالكامل من الباك إند (الفرونت إند كان بينادي على
# /api/weight-log/* وهي مش موجودة خالص، عشان كده كل صفحات الوزن كانت
# بترجع "Failed to fetch"). الملف ده بيضيفه من غير ما يغيّر أي حاجة
# تانية شغالة في المشروع.
#
# التركيب:
# 1) في app.py الأساسي، ضيف:
#       from weight_log_api import weight_log_bp
#       app.register_blueprint(weight_log_bp)
#
# 2) الجدولين اللي الفيتشر ده بيستخدمهم (weight_log_entries و weight_log_items)
#    لازم يكونوا موجودين في Supabase (هم موجودين بالفعل).
#
# 3) في متغيرات البيئة (اختياري، مش لازم تغيّرهم عشان يشتغل فورًا):
#       WEIGHT_LOG_ENTRY_TOKEN  — التوكين اللي رابط العامل (/w) بيستخدمه (افتراضي:
#                                 نفس التوكين المكتوب حاليًا في .htaccess عشان الرابط
#                                 القديم يفضل شغال من غير أي تغيير).
#       WEIGHT_LOG_VIEW_TOKEN   — التوكين اللي رابط الأرشيف للعرض بيستخدمه (افتراضي:
#                                 نفس التوكين المكتوب حاليًا في weight-log-dashboard.html).
#
# ============================================================

import os
import base64
import secrets
import binascii
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

from flask import Blueprint, request, jsonify, Response

from db import get_client, execute_with_retry

weight_log_bp = Blueprint('weight_log_api', __name__)

# نفس التوكينز المكتوبة حاليًا في .htaccess (رابط /w) وفي weight-log-dashboard.html
# (رابط الأرشيف) — سايبينهم كـ default عشان الروابط الشغالة دلوقتي متتكسرش.
WEIGHT_LOG_ENTRY_TOKEN = os.environ.get('WEIGHT_LOG_ENTRY_TOKEN', 'pNrAYo0cIwXhdsgVdXKSJYCGAS8')
WEIGHT_LOG_VIEW_TOKEN = os.environ.get('WEIGHT_LOG_VIEW_TOKEN', 'vXq3mZpLd8RwTfKhY0eB2nCsUj7A')

RIYADH_TZ = ZoneInfo('Asia/Riyadh')  # نفس التوقيت المستخدم في كل صفحات الوزن بالفرونت إند


def _check_admin_session(token):
    """نفس منطق _check_session في app.py بالظبط، مكرر هنا عشان الملف يفضل مستقل
    (من غير استيراد دائري من app.py)."""
    if not token:
        return None
    sb = get_client()
    res = execute_with_retry(sb.table('app_sessions').select('*').eq('token', token))
    rows = res.data or []
    if not rows:
        return None
    row = rows[0]
    expires_at = datetime.fromisoformat(row['expires_at'].replace('Z', '+00:00'))
    if expires_at < datetime.now(timezone.utc):
        return None
    return row['username']


def _is_admin_request():
    token = request.headers.get('X-Auth-Token')
    return bool(_check_admin_session(token))


def _has_valid_entry_token():
    token = request.values.get('token') or (request.get_json(silent=True) or {}).get('token')
    return bool(token) and secrets.compare_digest(str(token), WEIGHT_LOG_ENTRY_TOKEN)


def _has_valid_view_token():
    token = request.args.get('view_token')
    return bool(token) and secrets.compare_digest(str(token), WEIGHT_LOG_VIEW_TOKEN)


def _require_admin():
    if not _is_admin_request():
        return jsonify({'error': 'جلسة غير صالحة، سجّل دخول تاني'}), 401
    return None


def _require_worker_or_admin():
    """للـ routes اللي العامل (برابط /w) والأدمن (بجلسة تسجيل الدخول) الاتنين
    مسموحلهم يستخدموها (زي تعديل/حذف صنف اتسجل النهاردة)."""
    if _has_valid_entry_token() or _is_admin_request():
        return None
    return jsonify({'error': 'الرابط غير صالح'}), 401


def _today_range_riyadh():
    """بداية ونهاية يوم النهاردة (بتوقيت آسيا/الرياض، نفس التوقيت المستخدم في
    كل حسابات الأيام بالفرونت إند) كـ UTC ISO strings، عشان نفلتر بيهم logged_at."""
    now_riyadh = datetime.now(RIYADH_TZ)
    start_riyadh = now_riyadh.replace(hour=0, minute=0, second=0, microsecond=0)
    end_riyadh = start_riyadh.replace(hour=23, minute=59, second=59, microsecond=999999)
    return start_riyadh.astimezone(timezone.utc).isoformat(), end_riyadh.astimezone(timezone.utc).isoformat()


def _strip_photo_payload(rows):
    """بتشيل الصورة الكاملة (base64) من ردود القوايم/الأرشيف وتستبدلها بـ has_photo
    (True/False) بس. ده بيقلل حجم رد الشبكة بشكل كبير جدًا لأن القايمة مبتحتاجش
    تعرض الصور كلها مرة واحدة - الصورة الفعلية بتتحمّل بس لما المستخدم يفتحها،
    من الرابط الموجود بالفعل /api/weight-log/<id>/photo. الشكل والسلوك في
    الفرونت إند بيفضلوا زي ما هم بالظبط، بس أسرع بكتير مع زيادة عدد السجلات."""
    stripped = []
    for row in rows:
        row = dict(row)
        row['has_photo'] = bool(row.get('photo_base64'))
        row.pop('photo_base64', None)
        stripped.append(row)
    return stripped


def _decode_photo_data_url(data_url):
    """يرجّع (mimetype, bytes) من data URL زي 'data:image/jpeg;base64,...'، أو None لو مش صالح."""
    if not data_url or ',' not in data_url:
        return None
    header, b64data = data_url.split(',', 1)
    mimetype = 'image/jpeg'
    if header.startswith('data:') and ';' in header:
        mimetype = header[5:header.index(';')] or mimetype
    try:
        raw = base64.b64decode(b64data)
    except (binascii.Error, ValueError):
        return None
    return mimetype, raw


# ===================== تسجيل الأصناف (رابط العامل /w) =====================

@weight_log_bp.route('/api/weight-log', methods=['GET', 'POST'])
def weight_log_root():
    if request.method == 'GET':
        # قايمة الأدمن الكاملة (بدون المحذوف) - صفحة weight-log-dashboard
        err = _require_admin()
        if err:
            return err
        sb = get_client()
        res = execute_with_retry(
            sb.table('weight_log_entries').select('*').eq('deleted', False).order('logged_at', desc=True)
        )
        return jsonify({'entries': _strip_photo_payload(res.data or [])})

    # POST - العامل بيضيف صنف جديد من صفحة weight-log-entry.html
    if not _has_valid_entry_token():
        return jsonify({'error': 'الرابط غير صالح'}), 401

    item_name = (request.form.get('item_name') or '').strip()
    weight = (request.form.get('weight') or '').strip()
    day_names = (request.form.get('day_names') or '').strip()
    batch_no = (request.form.get('batch_no') or '').strip()

    if not item_name:
        return jsonify({'error': 'اسم الصنف مطلوب'}), 400
    try:
        weight_val = float(weight)
    except (TypeError, ValueError):
        return jsonify({'error': 'الوزن غير صحيح'}), 400

    photo_base64 = None
    photo_file = request.files.get('photo')
    if photo_file and photo_file.filename:
        raw = photo_file.read()
        mimetype = photo_file.mimetype or 'image/jpeg'
        photo_base64 = f'data:{mimetype};base64,{base64.b64encode(raw).decode("ascii")}'

    sb = get_client()
    res = execute_with_retry(sb.table('weight_log_entries').insert({
        'item_name': item_name,
        'weight': weight_val,
        'photo_base64': photo_base64,
        'logged_at': datetime.now(timezone.utc).isoformat(),
        'deleted': False,
        'day_names': day_names,
        'batch_no': batch_no,
    }))
    row_id = res.data[0]['id']
    return jsonify({'id': row_id})


@weight_log_bp.route('/api/weight-log/mine', methods=['GET'])
def weight_log_mine():
    """أصناف النهاردة بس (اللي العامل نفسه سجلها) - قايمة "اليوم" في صفحة الإدخال."""
    if not _has_valid_entry_token():
        return jsonify({'error': 'الرابط غير صالح'}), 401
    start_iso, end_iso = _today_range_riyadh()
    sb = get_client()
    res = execute_with_retry(
        sb.table('weight_log_entries').select('*')
        .eq('deleted', False)
        .gte('logged_at', start_iso).lte('logged_at', end_iso)
        .order('logged_at', desc=True)
    )
    return jsonify({'entries': _strip_photo_payload(res.data or [])})


@weight_log_bp.route('/api/weight-log/<int:entry_id>', methods=['PUT', 'DELETE'])
def weight_log_entry_detail(entry_id):
    err = _require_worker_or_admin()
    if err:
        return err
    sb = get_client()

    if request.method == 'DELETE':
        # حذف "ناعم" (soft delete) - يفضل ظاهر في الأرشيف بعلامة "محذوف" لكن
        # يختفي من القايمة العادية. الحذف النهائي في /permanent (أدمن بس).
        execute_with_retry(sb.table('weight_log_entries').update({'deleted': True}).eq('id', entry_id))
        return jsonify({'ok': True})

    payload = request.get_json(silent=True) or {}
    item_name = (payload.get('item_name') or '').strip()
    weight = payload.get('weight')
    batch_no = (payload.get('batch_no') or '').strip()
    if not item_name:
        return jsonify({'error': 'اسم الصنف مطلوب'}), 400
    try:
        weight_val = float(weight)
    except (TypeError, ValueError):
        return jsonify({'error': 'الوزن غير صحيح'}), 400

    execute_with_retry(sb.table('weight_log_entries').update({
        'item_name': item_name, 'weight': weight_val, 'batch_no': batch_no,
    }).eq('id', entry_id))
    return jsonify({'ok': True})


@weight_log_bp.route('/api/weight-log/<int:entry_id>/permanent', methods=['DELETE'])
def weight_log_entry_delete_permanent(entry_id):
    """حذف نهائي - أدمن بس (من صفحة الأرشيف)."""
    err = _require_admin()
    if err:
        return err
    sb = get_client()
    execute_with_retry(sb.table('weight_log_entries').delete().eq('id', entry_id))
    return jsonify({'ok': True})


@weight_log_bp.route('/api/weight-log/<int:entry_id>/photo', methods=['GET'])
def weight_log_entry_photo(entry_id):
    """بترجّع صورة الميزان كـ image بايتس مباشرة (مش JSON) - مستخدمة في روابط
    Excel المصدّر وفي صفحة weight-photo.html."""
    sb = get_client()
    res = execute_with_retry(sb.table('weight_log_entries').select('photo_base64').eq('id', entry_id))
    rows = res.data or []
    if not rows or not rows[0].get('photo_base64'):
        return jsonify({'error': 'الصورة غير متاحة'}), 404
    decoded = _decode_photo_data_url(rows[0]['photo_base64'])
    if not decoded:
        return jsonify({'error': 'الصورة تالفة'}), 404
    mimetype, raw = decoded
    return Response(raw, mimetype=mimetype)


@weight_log_bp.route('/api/weight-log/archive', methods=['GET'])
def weight_log_archive():
    """أرشيف كامل (شامل المحذوف) - رابط عرض فقط بتوكين ثابت، صفحة weight-log-archive.html."""
    if not _has_valid_view_token():
        return jsonify({'error': 'الرابط غير صالح'}), 401
    sb = get_client()
    res = execute_with_retry(sb.table('weight_log_entries').select('*').order('logged_at', desc=True))
    return jsonify({'entries': _strip_photo_payload(res.data or [])})


# ===================== إدارة قايمة الأصناف (weight_log_items) =====================

@weight_log_bp.route('/api/weight-log/items', methods=['GET', 'POST'])
def weight_log_items_root():
    sb = get_client()

    if request.method == 'GET':
        # العامل بيطلب أصناف يوم معين عشان يملى الـ dropdown
        if not _has_valid_entry_token():
            return jsonify({'error': 'الرابط غير صالح'}), 401
        day = request.args.get('day') or ''
        res = execute_with_retry(
            sb.table('weight_log_items').select('*').eq('day_name', day).order('sort_order')
        )
        return jsonify({'items': res.data or []})

    # POST - تسجيل صنف جديد (من العامل بتوكين، أو من الأدمن بجلسة تسجيل دخول)
    if not (_has_valid_entry_token() or _is_admin_request()):
        return jsonify({'error': 'الرابط غير صالح'}), 401

    payload = request.get_json(silent=True) or {}
    day = (payload.get('day') or '').strip()
    item_name = (payload.get('item_name') or '').strip()
    if not day or not item_name:
        return jsonify({'error': 'اليوم واسم الصنف مطلوبين'}), 400

    # لو الصنف موجود أصلاً في نفس اليوم (بغض النظر عن حالة الأحرف)، رجّعه من
    # غير ما نكرره
    existing_res = execute_with_retry(
        sb.table('weight_log_items').select('*').eq('day_name', day)
    )
    for row in (existing_res.data or []):
        if str(row.get('item_name', '')).strip().lower() == item_name.lower():
            return jsonify({'item': row})

    max_order_res = execute_with_retry(
        sb.table('weight_log_items').select('sort_order').eq('day_name', day).order('sort_order', desc=True).limit(1)
    )
    max_rows = max_order_res.data or []
    next_order = (max_rows[0]['sort_order'] + 1) if max_rows else 0

    insert_res = execute_with_retry(sb.table('weight_log_items').insert({
        'day_name': day, 'item_name': item_name, 'sort_order': next_order,
    }))
    return jsonify({'item': insert_res.data[0]})


@weight_log_bp.route('/api/weight-log/items/all', methods=['GET'])
def weight_log_items_all():
    """كل الأصناف في كل الأيام - أدمن بس، صفحة weight-log-items-dashboard.html."""
    err = _require_admin()
    if err:
        return err
    sb = get_client()
    res = execute_with_retry(sb.table('weight_log_items').select('*').order('day_name').order('sort_order'))
    return jsonify({'items': res.data or []})


@weight_log_bp.route('/api/weight-log/items/reorder', methods=['PUT'])
def weight_log_items_reorder():
    err = _require_admin()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    order = payload.get('order') or []
    if not order:
        return jsonify({'error': 'الترتيب فارغ'}), 400
    sb = get_client()
    for index, item_id in enumerate(order):
        execute_with_retry(sb.table('weight_log_items').update({'sort_order': index}).eq('id', item_id))
    return jsonify({'ok': True})


@weight_log_bp.route('/api/weight-log/items/<int:item_id>', methods=['PUT', 'DELETE'])
def weight_log_item_detail(item_id):
    err = _require_admin()
    if err:
        return err
    sb = get_client()
    if request.method == 'DELETE':
        execute_with_retry(sb.table('weight_log_items').delete().eq('id', item_id))
        return jsonify({'ok': True})

    payload = request.get_json(silent=True) or {}
    item_name = (payload.get('item_name') or '').strip()
    if not item_name:
        return jsonify({'error': 'اسم الصنف مطلوب'}), 400
    execute_with_retry(sb.table('weight_log_items').update({'item_name': item_name}).eq('id', item_id))
    return jsonify({'ok': True})
