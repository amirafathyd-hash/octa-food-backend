import os
import io
import re
import secrets
import tempfile
import zipfile
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from copy import copy
from datetime import datetime, timedelta, timezone
from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from parse_order import parse_order_pdf
from parse_invoice import parse_invoice_pdf
from parse_received import parse_received_xlsx
from parse_received_image import parse_received_image, process_ocr_data
from item_db import load_db, seed_from_order
from matcher import match_invoice_item
from db import get_client, execute_with_retry
from invoice_export import parse_invoice_full, build_invoices_workbook
from tokyo_ordering import read_current_inputs, write_updated_workbook

TOKYO_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'tokyo_ordering_template.xlsm')

app = Flask(__name__)
CORS(app)  # allow calls from the Netlify frontend domain


def _to_iso(date_str):
    """'1-Apr-2026' -> '2026-04-01'"""
    return datetime.strptime(date_str, '%d-%b-%Y').strftime('%Y-%m-%d')


def _bulk_upsert_daily(rows):
    """Upserts many daily_items rows in ONE request instead of one request per row
    (a single order PDF can have 70+ items; doing them one by one was slow enough
    to hit request timeouts on the free hosting tier).
    Also de-duplicates by (item_date, item_key) within the same batch, since Postgres
    rejects an upsert where the same conflict target appears twice in one statement
    (this can happen if the same item name shows up in both the salads and dressing
    sections of an order PDF)."""
    if not rows:
        return
    deduped = {}
    for row in rows:
        deduped[(row['item_date'], row['item_key'])] = row
    sb = get_client()
    execute_with_retry(sb.table('daily_items').upsert(list(deduped.values()), on_conflict='item_date,item_key'))


def _log(file_type, file_name, item_date, message, level='info'):
    sb = get_client()
    try:
        execute_with_retry(sb.table('upload_log').insert({
            'file_type': file_type, 'file_name': file_name,
            'item_date': item_date, 'message': message, 'level': level,
        }), max_attempts=2)
    except Exception:
        pass  # logging is best-effort; never let a logging failure break the actual request


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


@app.route('/api/upload-order', methods=['POST'])
def upload_order():
    """Accepts one or more order PDFs (multipart 'files'). Upserts qty_needed/box/inventory."""
    files = request.files.getlist('files')
    results = []
    for f in files:
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            f.save(tmp.name)
            path = tmp.name
        try:
            order = parse_order_pdf(path)
            seed_from_order(order)
            date_iso = _to_iso(order['date'])
            rows = []
            for section in ('salads', 'dressing'):
                for item in order[section]:
                    key = item['name_en'].strip().upper()
                    rows.append({
                        'item_date': date_iso,
                        'item_key': key,
                        'name_en': item['name_en'],
                        'name_ar': item['name_ar'],
                        'section': section,
                        'qty_box': item['qty_box'],
                        'qty_needed': item['qty_needed'],
                        'unit': item['unit'].split('-')[0],
                        'current_inventory': item['current_inventory'],
                    })
            _bulk_upsert_daily(rows)
            count = len(rows)
            _log('order', f.filename, date_iso, f'تم استيراد {count} صنف بنجاح')
            results.append({'file': f.filename, 'date': order['date'], 'items': count, 'status': 'ok'})
        except Exception as e:
            _log('order', f.filename, None, str(e), level='warning')
            results.append({'file': f.filename, 'status': 'error', 'error': str(e)})
        finally:
            os.unlink(path)
    return jsonify({'results': results})


@app.route('/api/upload-invoice', methods=['POST'])
def upload_invoice():
    """Accepts one or more invoice PDFs. Matches items and upserts invoice_qty/price."""
    files = request.files.getlist('files')
    results = []
    db = load_db()
    for f in files:
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            f.save(tmp.name)
            path = tmp.name
        try:
            invoice = parse_invoice_pdf(path)
            if not invoice['date']:
                raise ValueError('لم يتم العثور على تاريخ في الفاتورة')
            date_iso = invoice['date']  # already YYYY-MM-DD in parse_invoice
            rows = []
            log_messages = []
            matched, unmatched = 0, 0
            for it in invoice['items']:
                key, score, method = match_invoice_item(it['name_ar'], db)
                if key:
                    rows.append({
                        'item_date': date_iso,
                        'item_key': key,
                        'invoice_qty': it['qty'],
                        'invoice_price': it['total'],
                        'invoice_unit_label': it['unit_label'],
                    })
                    matched += 1
                    if method == 'fuzzy':
                        log_messages.append((
                            'info',
                            f"مطابقة ذكية: \"{it['name_ar']}\" -> {key} (تشابه {score:.0f}%)"
                        ))
                else:
                    unmatched += 1
                    log_messages.append((
                        'warning',
                        f"لم يتم العثور على تطابق لـ \"{it['name_ar']}\" (أعلى تشابه {score:.0f}%)"
                    ))
            _bulk_upsert_daily(rows)
            for level, msg in log_messages:
                _log('invoice', f.filename, date_iso, msg, level=level)
            results.append({'file': f.filename, 'date': date_iso, 'matched': matched,
                             'unmatched': unmatched, 'status': 'ok'})
        except Exception as e:
            _log('invoice', f.filename, None, str(e), level='warning')
            results.append({'file': f.filename, 'status': 'error', 'error': str(e)})
        finally:
            os.unlink(path)
    return jsonify({'results': results})


@app.route('/api/upload-received', methods=['POST'])
def upload_received():
    """Accepts 'received' files — either Excel sheets (legacy) or photos/scans of
    the printed sheet with handwritten received quantities (OCR via EasyOCR)."""
    files = request.files.getlist('files')
    results = []
    db = None  # lazily loaded only if we hit an image file

    for f in files:
        is_image = f.filename.lower().endswith(('.jpg', '.jpeg', '.png'))
        suffix = '.jpg' if is_image else '.xlsx'
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            f.save(tmp.name)
            path = tmp.name
        try:
            if is_image:
                if db is None:
                    db = load_db()
                parsed = parse_received_image(path, db)
                if not parsed['date']:
                    raise ValueError('لم يتم العثور على تاريخ مطبوع في الصورة')
                date_iso = parsed['date']
                rows = []
                review_count = 0
                for r in parsed['rows']:
                    rows.append({
                        'item_date': date_iso,
                        'item_key': r['item_key'],
                        'qty_received': r['qty'],
                        'rec_unit': r['unit'],
                    })
                    if r['needs_review']:
                        review_count += 1
                        _log('received', f.filename, date_iso,
                             f"يحتاج مراجعة: \"{r['raw_text']}\" بجانب {r['name_en']} "
                             f"(ثقة {r['confidence']}%)", level='warning')
                _bulk_upsert_daily(rows)
                _log('received', f.filename, date_iso,
                     f"تم استيراد {len(rows)} قيمة من الصورة بالـ OCR "
                     f"({review_count} منهم يحتاجون مراجعة)")
                results.append({'file': f.filename, 'date': date_iso, 'rows': len(rows),
                                 'needs_review': review_count, 'status': 'ok'})
            else:
                records = parse_received_xlsx(path)
                rows = []
                dates_seen = set()
                for rec in records:
                    if rec['qty_received'] is None and rec['qty_needed'] is None:
                        continue
                    date_iso = _to_iso(rec['date'])
                    dates_seen.add(date_iso)
                    rows.append({
                        'item_date': date_iso,
                        'item_key': rec['key'],
                        'name_en': rec['name_en'],
                        'name_ar': rec['name_ar'],
                        'qty_box': rec['qty_box'],
                        'qty_needed': rec['qty_needed'],
                        'unit': rec['unit'],
                        'current_inventory': rec['current_inventory'],
                        'qty_received': rec['qty_received'],
                        'rec_unit': rec['rec_unit'],
                    })
                for i in range(0, len(rows), 500):
                    _bulk_upsert_daily(rows[i:i + 500])
                count = len(rows)
                _log('received', f.filename, None, f'تم استيراد {count} صف ({len(dates_seen)} يوم)')
                results.append({'file': f.filename, 'rows': count, 'days': len(dates_seen), 'status': 'ok'})
        except Exception as e:
            _log('received', f.filename, None, str(e), level='warning')
            results.append({'file': f.filename, 'status': 'error', 'error': str(e)})
        finally:
            os.unlink(path)
    return jsonify({'results': results})


@app.route('/api/process-received-ocr', methods=['POST'])
def process_received_ocr():
    """Accepts OCR.space's JSON result (already fetched by the BROWSER, which has
    no network restrictions) for one 'received' image, plus the original
    filename. Does the row-matching/parsing here on the server — no outbound
    network call needed for this part, so it works fine even on a network-
    restricted free hosting tier."""
    payload = request.get_json(silent=True) or {}
    ocr_data = payload.get('ocr_data')
    filename = payload.get('filename', 'image')
    if not ocr_data:
        return jsonify({'results': [{'file': filename, 'status': 'error', 'error': 'لا توجد بيانات OCR'}]})

    db = load_db()
    try:
        parsed = process_ocr_data(ocr_data, db, filename=filename)
        if not parsed['date']:
            raise ValueError('لم يتم العثور على تاريخ مطبوع في الصورة')
        date_iso = parsed['date']
        rows = []
        review_count = 0
        for r in parsed['rows']:
            rows.append({
                'item_date': date_iso,
                'item_key': r['item_key'],
                'qty_received': r['qty'],
                'rec_unit': r['unit'],
            })
            if r['needs_review']:
                review_count += 1
                _log('received', filename, date_iso,
                     f"يحتاج مراجعة: \"{r['raw_text']}\" بجانب {r['name_en']} (ثقة {r['confidence']}%)",
                     level='warning')
        _bulk_upsert_daily(rows)
        _log('received', filename, date_iso,
             f"تم استيراد {len(rows)} قيمة من الصورة بالـ OCR ({review_count} منهم يحتاجون مراجعة)")
        return jsonify({'results': [{'file': filename, 'date': date_iso, 'rows': len(rows),
                                      'needs_review': review_count, 'status': 'ok'}]})
    except Exception as e:
        _log('received', filename, None, str(e), level='warning')
        return jsonify({'results': [{'file': filename, 'status': 'error', 'error': str(e)}]})


@app.route('/api/log', methods=['GET'])
def get_log():
    sb = get_client()
    res = execute_with_retry(sb.table('upload_log').select('*').order('created_at', desc=True).limit(200))
    return jsonify(res.data)


@app.route('/api/report', methods=['GET'])
def report():
    """Returns JSON rows + summary stats for a given month, used to render the in-page table."""
    month = request.args.get('month')
    sb = get_client()
    q = sb.table('daily_items').select('*').order('item_date')
    if month:
        q = q.gte('item_date', f'{month}-01').lt('item_date', _next_month(month))
    res = execute_with_retry(q)
    rows = res.data

    stats = {'matched': 0, 'fuzzy': 0, 'needs_review': 0, 'no_invoice': 0, 'total': len(rows)}
    for r in rows:
        if r.get('invoice_qty') is None:
            stats['no_invoice'] += 1
        else:
            stats['matched'] += 1
        qn, qr = r.get('qty_needed'), r.get('qty_received')
        if qn and qr and abs(qn - qr) / qn > 0.10:
            stats['needs_review'] += 1

    return jsonify({'rows': rows, 'stats': stats})


@app.route('/api/finalize', methods=['GET'])
def finalize():
    """Builds the final comparison Excel for a given month (e.g. ?month=2026-04)."""
    month = request.args.get('month')  # 'YYYY-MM'
    try:
        sb = get_client()
        q = sb.table('daily_items').select('*').order('item_date')
        if month:
            q = q.gte('item_date', f'{month}-01').lt('item_date', _next_month(month))
        res = execute_with_retry(q)
        rows = res.data

        from excel_writer import build_workbook
        wb_path = build_workbook(rows)
        return send_file(wb_path, as_attachment=True,
                          download_name=f"octa_food_report_{month or 'all'}.xlsx")
    except Exception as e:
        app.logger.exception('finalize failed for month=%s', month)
        return jsonify({'error': f'تعذر إنشاء التقرير: {e}'}), 500


@app.route('/api/invoices-export', methods=['POST'])
def invoices_export():
    """يستقبل عدة ملفات PDF فواتير، يستخرج منها كل البيانات (تاريخ، رقم فاتورة،
    مورد، عميل، بنود، إجماليات) مع تصحيح ترتيب الحروف العربي، ويرجعها JSON
    عشان الفرونت إند يبني منها شيت إكسل قابل للتعديل قبل التحميل.
    مستقل تمامًا عن /api/upload-invoice ومفيش أي تأثير على قاعدة البيانات."""
    files = request.files.getlist('files')
    results = []
    for f in files:
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            f.save(tmp.name)
            path = tmp.name
        try:
            data = parse_invoice_full(path, f.filename)
            results.append(data)
        except Exception as e:
            results.append({
                'fileName': f.filename,
                'date': '', 'number': '', 'party': '',
                'subtotal': 0, 'vat': 0, 'total': 0,
                'items': [],
                'notes': f'تعذر قراءة الملف: {e}',
            })
        finally:
            os.unlink(path)
    return jsonify({'invoices': results})


@app.route('/api/invoices-export-xlsx', methods=['POST'])
def invoices_export_xlsx():
    """يستقبل بيانات الفواتير (بعد ما المستخدم يراجعها ويعدّلها في الواجهة) كـ JSON
    ويرجّع ملف إكسل منسّق بالكامل (ألوان، حدود، خط عريض للإجماليات) جاهز للتحميل."""
    payload = request.get_json(silent=True) or {}
    invoices = payload.get('invoices') or []
    wb_path = build_invoices_workbook(invoices)
    return send_file(wb_path, as_attachment=True,
                      download_name=f"octa-invoices-{datetime.now().strftime('%Y-%m-%d')}.xlsx")


@app.route('/api/tokyo-ordering/current', methods=['GET'])
def tokyo_ordering_current():
    """يرجّع الأيام والوجبات الحالية (العدد والوزن) من ملف القالب المخزّن على السيرفر."""
    if not os.path.exists(TOKYO_TEMPLATE_PATH):
        return jsonify({'error': 'ملف القالب tokyo_ordering_template.xlsm غير موجود على السيرفر'}), 404
    days = read_current_inputs(TOKYO_TEMPLATE_PATH)
    return jsonify({'days': days})


@app.route('/api/tokyo-ordering/export', methods=['POST'])
def tokyo_ordering_export():
    """يستقبل الأيام بعد التعديل، ويرجّع نسخة محدّثة من نفس ملف القالب (بالماكرو والمعادلات
    زي ما هي) فيها القيم الجديدة بس، جاهزة تتحمّل وتفتحها في إكسل وتدوس Update زي العادة."""
    if not os.path.exists(TOKYO_TEMPLATE_PATH):
        return jsonify({'error': 'ملف القالب tokyo_ordering_template.xlsm غير موجود على السيرفر'}), 404
    payload = request.get_json(silent=True) or {}
    days = payload.get('days') or []
    out_path = write_updated_workbook(TOKYO_TEMPLATE_PATH, days)
    return send_file(out_path, as_attachment=True, download_name='Tokyo_Ordering_Updated.xlsm')


def _next_month(month):
    y, m = map(int, month.split('-'))
    return f'{y+1}-01-01' if m == 12 else f'{y}-{m+1:02d}-01'


SESSION_DAYS = 7


def _new_session(username):
    token = secrets.token_hex(32)
    expires_at = (datetime.now(timezone.utc) + timedelta(days=SESSION_DAYS)).isoformat()
    sb = get_client()
    execute_with_retry(sb.table('app_sessions').insert({
        'token': token, 'username': username, 'expires_at': expires_at,
    }))
    return token


def _check_session(token):
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


def _require_auth():
    """يرجّع username لو التوكين صحيح، أو يرجّع None ومعاه الـ response المناسب لو لأ."""
    token = request.headers.get('X-Auth-Token') or (request.get_json(silent=True) or {}).get('token')
    username = _check_session(token)
    if not username:
        return None, (jsonify({'error': 'جلسة غير صالحة، سجّل دخول تاني'}), 401)
    return username, None


@app.route('/api/setup-first-user', methods=['POST'])
def setup_first_user():
    """شغّال بس أول مرة (لما جدول app_users يكون فاضي)، عشان تعمل أول حساب admin.
    بعد ما يتضاف أول يوزر، الراوت ده بيقفل نفسه أوتوماتيك ومايقبلش طلبات تانية."""
    sb = get_client()
    existing = execute_with_retry(sb.table('app_users').select('id').limit(1))
    if existing.data:
        return jsonify({'error': 'فيه يوزرز موجودين بالفعل، استخدم صفحة تسجيل الدخول.'}), 403

    payload = request.get_json(silent=True) or {}
    username = (payload.get('username') or '').strip()
    password = payload.get('password') or ''
    if not username or len(password) < 4:
        return jsonify({'error': 'اليوزر نيم مطلوب والباسورد لازم يكون 4 حروف على الأقل'}), 400

    execute_with_retry(sb.table('app_users').insert({
        'username': username, 'password_hash': generate_password_hash(password),
    }))
    return jsonify({'ok': True})


@app.route('/api/login', methods=['POST'])
def login():
    payload = request.get_json(silent=True) or {}
    username = (payload.get('username') or '').strip()
    password = payload.get('password') or ''
    if not username or not password:
        return jsonify({'error': 'اليوزر نيم والباسورد مطلوبين'}), 400

    sb = get_client()
    res = execute_with_retry(sb.table('app_users').select('*').eq('username', username))
    rows = res.data or []
    if not rows or not check_password_hash(rows[0]['password_hash'], password):
        return jsonify({'error': 'اليوزر نيم أو الباسورد غلط'}), 401

    token = _new_session(username)
    return jsonify({'token': token, 'username': username})


@app.route('/api/verify-session', methods=['GET'])
def verify_session():
    token = request.args.get('token')
    username = _check_session(token)
    if not username:
        return jsonify({'valid': False}), 401
    return jsonify({'valid': True, 'username': username})


@app.route('/api/logout', methods=['POST'])
def logout():
    payload = request.get_json(silent=True) or {}
    token = payload.get('token')
    if token:
        sb = get_client()
        execute_with_retry(sb.table('app_sessions').delete().eq('token', token))
    return jsonify({'ok': True})


@app.route('/api/users', methods=['GET'])
def list_users():
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    res = execute_with_retry(sb.table('app_users').select('id, username, created_at').order('created_at'))
    return jsonify({'users': res.data or []})


@app.route('/api/users', methods=['POST'])
def create_user():
    _, err = _require_auth()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    username = (payload.get('username') or '').strip()
    password = payload.get('password') or ''
    if not username or not password:
        return jsonify({'error': 'اليوزر نيم والباسورد مطلوبين'}), 400
    if len(password) < 4:
        return jsonify({'error': 'الباسورد لازم يكون 4 حروف/أرقام على الأقل'}), 400

    sb = get_client()
    try:
        execute_with_retry(sb.table('app_users').insert({
            'username': username, 'password_hash': generate_password_hash(password),
        }))
    except Exception as e:
        return jsonify({'error': f'تعذر إضافة اليوزر (ممكن يكون موجود قبل كده): {e}'}), 400
    return jsonify({'ok': True})


@app.route('/api/users/<int:user_id>', methods=['DELETE'])
def delete_user(user_id):
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    execute_with_retry(sb.table('app_users').delete().eq('id', user_id))
    return jsonify({'ok': True})


@app.route('/api/users/<int:user_id>/password', methods=['PUT'])
def change_user_password(user_id):
    _, err = _require_auth()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    password = payload.get('password') or ''
    if len(password) < 4:
        return jsonify({'error': 'الباسورد لازم يكون 4 حروف/أرقام على الأقل'}), 400

    sb = get_client()
    execute_with_retry(sb.table('app_users').update({
        'password_hash': generate_password_hash(password),
    }).eq('id', user_id))
    return jsonify({'ok': True})


STATION_SHEET_MAP = {
    'breakfast': 'Ordering',
    'desserts': 'Ordering',
    'hot': 'All_Ingredients',
    'marination': 'Marination_Ordering',
    'rice': 'Ordering',
    'salads': 'Ordering',
    'sauce': 'Ordering',
}
# الترتيب الأبجدي اللي طلبه العميل (Breakfast, Desserts, Hot Section, Marination, Rice, Salads, Sauces)
STATION_ORDER = ['breakfast', 'desserts', 'hot', 'marination', 'rice', 'salads', 'sauce']
STATION_LABELS = {
    'breakfast': 'Breakfast', 'desserts': 'Desserts', 'hot': 'Hot Section',
    'marination': 'Marination', 'rice': 'Rice', 'salads': 'Salads', 'sauce': 'Sauces',
}
STATION_TAB_NAMES = {
    'breakfast': 'Breakfast', 'desserts': 'Desserts', 'hot': 'Hot Section',
    'marination': 'Marination', 'rice': 'Rice', 'salads': 'Salads', 'sauce': 'Sauce',
}
PURPLE_FILL = PatternFill(fill_type='solid', fgColor='6600FF')


def _read_station_rows(file_storage, sheet_name):
    """بيرجّع dict: name -> {'unit':..,'category':..,'weekly':..} من شيت المحطة
    المحدّد بالاسم (عشان ملف توكيو فيه أكتر من شيت محتمل، ولازم نحدد الصحيح
    لكل محطة بالاسم مش بالتخمين).
    أعمدة المصدر: A=الاسم، B=الفئة، C=الوحدة، D=الوزن اليومي، E=الوزن الأسبوعي."""
    wb = openpyxl.load_workbook(file_storage, data_only=True)
    if sheet_name not in wb.sheetnames:
        return None, {}
    ws = wb[sheet_name]
    out = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=5, values_only=True):
        name, category, unit, _daily, weekly = (list(row) + [None] * 5)[:5]
        if not name or not str(name).strip():
            continue
        if str(name).strip().lower() == 'items':
            continue
        out[str(name).strip()] = {
            'unit': unit, 'category': category,
            'weekly': weekly if isinstance(weekly, (int, float)) else 0,
        }
    return sheet_name, out


def _style_header_cell(cell, size=11, bold=True):
    cell.font = Font(name='Calibri', size=size, bold=bold)
    cell.fill = PURPLE_FILL


def _build_purchasing_workbook(station_data):
    """station_data: {station_key: {ingredient: {'unit','category','weekly'}}}
    بيرجّع openpyxl.Workbook فيه شيت Purchasing + شيت لكل محطة (A:E، نفس التنسيق)."""
    all_names = set()
    for data in station_data.values():
        all_names.update(data.keys())
    sorted_names = sorted(all_names, key=lambda s: s.lower())

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Purchasing'

    # ===== الهيدر =====
    ws['A3'] = 'Sum of Weekly Weight'
    for col in range(1, 19):  # A..R
        _style_header_cell(ws.cell(row=3, column=col), size=18)
    for idx, key in enumerate(STATION_ORDER):
        col = 4 + idx  # D=4
        ws.cell(row=4, column=col, value=STATION_LABELS[key])
        _style_header_cell(ws.cell(row=4, column=col), size=18)
    sum_col = 4 + len(STATION_ORDER)         # K
    ws.cell(row=4, column=sum_col, value='Sum of Weekly Consumption')
    _style_header_cell(ws.cell(row=4, column=sum_col), size=18)

    dup_col = sum_col + 2                     # M (سايبين عمود فاضي زي الأصل)
    maq_col = dup_col + 1                     # N
    exp_col = maq_col + 1                     # O
    avail_col = exp_col + 1                   # P
    order_col = avail_col + 1                 # Q
    next_col = order_col + 1                  # R

    headers_en = {
        dup_col: 'Weekly Consumtion', maq_col: 'Min. Available Quantity (MAQ )',
        exp_col: 'Expected available Stock', avail_col: 'Available Stock',
        order_col: 'Weekly Order', next_col: 'Expected available Stock Next week',
    }
    headers_ar = {
        dup_col: 'الاستهلاك الأسبوعى', maq_col: 'الحد الأدنى للكمية المتاحة (MAQ)',
        exp_col: 'المخزون المتوقع المتاح', avail_col: 'المخزون المتاح',
        order_col: 'الطلب الأسبوعي', next_col: 'المخزون المتوقع المتاح للأسبوع القادم',
    }
    for col, text in headers_en.items():
        ws.cell(row=3, column=col, value=text)
        _style_header_cell(ws.cell(row=3, column=col), size=11)
    for col, text in headers_ar.items():
        ws.cell(row=4, column=col, value=text)
        _style_header_cell(ws.cell(row=4, column=col), size=11)

    # ===== صفوف البيانات (من صف 5) =====
    sum_letter = get_column_letter(sum_col)
    dup_letter = get_column_letter(dup_col)
    maq_letter = get_column_letter(maq_col)
    exp_letter = get_column_letter(exp_col)
    avail_letter = get_column_letter(avail_col)
    order_letter = get_column_letter(order_col)
    d_letter = get_column_letter(4)
    j_letter = get_column_letter(4 + len(STATION_ORDER) - 1)

    for i, name in enumerate(sorted_names):
        r = 5 + i
        unit, category = '', ''
        for key in STATION_ORDER:
            info = station_data.get(key, {}).get(name)
            if info:
                unit = unit or info.get('unit') or ''
                category = category or info.get('category') or ''
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=unit)
        ws.cell(row=r, column=3, value=category)
        for idx, key in enumerate(STATION_ORDER):
            col = 4 + idx
            info = station_data.get(key, {}).get(name)
            if info and info.get('weekly'):
                ws.cell(row=r, column=col, value=info['weekly'])
        ws.cell(row=r, column=sum_col, value=f'=SUM({d_letter}{r}:{j_letter}{r})')
        ws.cell(row=r, column=dup_col, value=f'={sum_letter}{r}')
        # MAQ / Expected / Available تتكتب يدوي كل أسبوع — تفضل فاضية عمدًا
        ws.cell(row=r, column=order_col,
                value=f'=({dup_letter}{r})-({avail_letter}{r}-{maq_letter}{r})')
        ws.cell(row=r, column=next_col,
                value=f'={order_letter}{r}+{avail_letter}{r}-{dup_letter}{r}')

    ws.column_dimensions['A'].width = 44
    ws.column_dimensions['B'].width = 8.5
    ws.column_dimensions['C'].width = 13
    for col in range(4, 4 + len(STATION_ORDER)):
        ws.column_dimensions[get_column_letter(col)].width = 17
    ws.column_dimensions[get_column_letter(sum_col)].width = 32
    ws.column_dimensions[get_column_letter(dup_col)].width = 27
    ws.column_dimensions[get_column_letter(maq_col)].width = 27
    ws.column_dimensions[get_column_letter(exp_col)].width = 31
    ws.column_dimensions[get_column_letter(avail_col)].width = 21
    ws.column_dimensions[get_column_letter(order_col)].width = 21
    ws.column_dimensions[get_column_letter(next_col)].width = 32
    ws.freeze_panes = 'A5'

    return wb, {
        'sum_col': sum_col, 'dup_col': dup_col, 'maq_col': maq_col, 'exp_col': exp_col,
        'avail_col': avail_col, 'order_col': order_col, 'next_col': next_col,
    }


def _add_station_tab(wb, station_key, file_storage):
    """بيضيف تاب لمحطة بنفس التنسيق الكامل (A:E)، باستخدام نفس منطق extract-sheet-range."""
    file_storage.seek(0)
    src_wb = openpyxl.load_workbook(file_storage, data_only=True)
    sheet_name = STATION_SHEET_MAP[station_key]
    if sheet_name not in src_wb.sheetnames:
        return None
    src_ws = src_wb[sheet_name]
    out_ws = wb.create_sheet(title=STATION_TAB_NAMES[station_key])

    COLS = 5
    for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row, min_col=1, max_col=COLS):
        for cell in row:
            new_cell = out_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format
    for col_letter in ['A', 'B', 'C', 'D', 'E']:
        if col_letter in src_ws.column_dimensions:
            out_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
    for merged_range in src_ws.merged_cells.ranges:
        if merged_range.max_col <= COLS:
            out_ws.merge_cells(str(merged_range))
    return out_ws


@app.route('/api/mega-purchasing', methods=['POST'])
def mega_purchasing():
    missing = [k for k in STATION_ORDER if k not in request.files]
    if missing:
        return jsonify({'error': f'محطات ناقصة: {", ".join(missing)}'}), 400

    try:
        station_data = {}
        for key in STATION_ORDER:
            _, rows = _read_station_rows(request.files[key], STATION_SHEET_MAP[key])
            station_data[key] = rows

        # ===== النسخة الكاملة (كل حاجة ظاهرة) =====
        wb_full, cols = _build_purchasing_workbook(station_data)
        for key in STATION_ORDER:
            request.files[key].seek(0)
            _add_station_tab(wb_full, key, request.files[key])

        # ===== نسخة المطبخ (التابات والأعمدة التفصيلية مخفية) =====
        wb_kitchen, _ = _build_purchasing_workbook(station_data)
        for key in STATION_ORDER:
            request.files[key].seek(0)
            ws_station = _add_station_tab(wb_kitchen, key, request.files[key])
            if ws_station is not None:
                ws_station.sheet_state = 'hidden'
        kitchen_ws = wb_kitchen['Purchasing']
        hide_from = 4  # D
        hide_to = cols['sum_col']  # K (آخر عمود تفصيلي قبل الفاصل)
        for col in range(hide_from, hide_to + 1):
            kitchen_ws.column_dimensions[get_column_letter(col)].hidden = True

        buf_full = io.BytesIO()
        wb_full.save(buf_full)
        buf_full.seek(0)
        buf_kitchen = io.BytesIO()
        wb_kitchen.save(buf_kitchen)
        buf_kitchen.seek(0)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            today = datetime.now().strftime('%Y-%m-%d')
            zf.writestr(f'Mega_Purchasing_Full_{today}.xlsx', buf_full.getvalue())
            zf.writestr(f'Mega_Purchasing_Kitchen_{today}.xlsx', buf_kitchen.getvalue())
        zip_buf.seek(0)
        return send_file(zip_buf, as_attachment=True, download_name=f'Mega_Purchasing_{datetime.now().strftime("%Y-%m-%d")}.zip',
                          mimetype='application/zip')
    except Exception as e:
        return jsonify({'error': f'حصل خطأ في التجميع: {e}'}), 500


@app.route('/api/extract-sheet-range', methods=['POST'])
def extract_sheet_range():
    """بتاخد ملف Excel مرفوع + اسم تاب، وترجّع نفس التاب (أعمدة A:E بس) في ملف
    جديد، بنفس الخط والألوان والحدود والمحاذاة وعرض الأعمدة والدمج تمامًا — لأن
    openpyxl بيقرا ويكتب كل تفاصيل التنسيق دي بدقة كاملة (بعكس مكتبات الجافاسكريبت
    المجانية اللي بس بتقرا لون الخلفية)."""
    file = request.files.get('file')
    sheet_name = request.form.get('sheet_name')
    if not file or not sheet_name:
        return jsonify({'error': 'محتاج الملف واسم التاب'}), 400

    try:
        src_wb = openpyxl.load_workbook(file, data_only=True)
    except Exception as e:
        return jsonify({'error': f'تعذر فتح الملف: {e}'}), 400

    if sheet_name not in src_wb.sheetnames:
        return jsonify({'error': f'مش لاقي تاب "{sheet_name}" في الملف ده'}), 404

    src_ws = src_wb[sheet_name]
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = sheet_name

    COLS = 5  # A..E
    for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row, min_col=1, max_col=COLS):
        for cell in row:
            new_cell = out_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format

    # عرض الأعمدة
    for col_letter in ['A', 'B', 'C', 'D', 'E']:
        if col_letter in src_ws.column_dimensions:
            out_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width

    # ارتفاع الصفوف
    for row_num, dim in src_ws.row_dimensions.items():
        if dim.height:
            out_ws.row_dimensions[row_num].height = dim.height

    # الخلايا المدموجة (بس اللي جوه A:E)
    for merged_range in src_ws.merged_cells.ranges:
        if merged_range.max_col <= COLS:
            out_ws.merge_cells(str(merged_range))

    out_ws.sheet_view.rightToLeft = src_ws.sheet_view.rightToLeft

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    safe_name = re.sub(r'[^\w\-]+', '-', sheet_name)
    return send_file(buf, as_attachment=True, download_name=f'{safe_name}-extract.xlsx',
                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/smart-order/meals', methods=['GET'])
def smart_order_meals():
    """بترجع قائمة كل الوجبات المتاحة للطلب الذكي (للـ dropdown)."""
    from smart_ordering import list_available_meals
    return jsonify({'meals': list_available_meals()})


@app.route('/api/smart-order/packages', methods=['GET'])
def smart_order_packages():
    """بترجع الوجبات منظّمة حسب باقات أيام المنيو، بنفس الأسماء اللي العميل
    شايفها على الموقع، مربوطة بأسماء الشيتات الحقيقية اللي بتتحسب عليها."""
    from smart_ordering import list_menu_packages
    return jsonify({'packages': list_menu_packages()})


@app.route('/api/smart-order/calculate', methods=['POST'])
def smart_order_calculate():
    """بتاخد {"orders": [{"meal_name": "...", "order_count": N}, ...]} وترجّع
    جدول التجهيز المحسوب لكل وجبة، عن طريق تشغيل صيغ الإكسل الحقيقية لكل
    وجبة على حدة (مش إعادة كتابة المعادلة يدويًا)."""
    from smart_ordering import calculate_multiple
    payload = request.get_json(silent=True) or {}
    orders = payload.get('orders', [])
    if not orders:
        return jsonify({'error': 'محتاج تبعت قائمة فيها وجبة واحدة على الأقل'}), 400
    try:
        result = calculate_multiple(orders)
        return jsonify(result)
    except Exception as e:
        app.logger.exception('smart_order_calculate failed')
        return jsonify({'error': f'حصل خطأ في الحساب: {e}'}), 500



# ============================================================================
# محطات الإنتاج الجديدة (Rice / Breakfast / Dessert / Salads / Sauce)
# بُنيت عشان زرار "احسب الإنتاج" في غرفة التحكم يطلّع نتايج المحطات دي كمان،
# مش بس الـ49 وجبة الأصلية.
# ============================================================================

@app.route('/api/smart-order/station-items', methods=['GET'])
def smart_order_station_items():
    """بترجع كل أصناف المحطات الخمسة الجديدة (الـ49 وجبة الأصلية شغالة أصلًا
    عن طريق /api/smart-order/packages)."""
    from station_calc import list_station_items, STATION_ITEMS
    return jsonify({
        'rice': list_station_items('rice'),
        'breakfast': list_station_items('breakfast'),
        'dessert': list_station_items('dessert'),
        'salads': list_station_items('salads'),
        'sauce_linked_meals': list(STATION_ITEMS['sauce'].keys()),
    })


@app.route('/api/smart-order/calculate-stations', methods=['POST'])
def smart_order_calculate_stations():
    """بتاخد:
    {
      "rice": [{"sheet_name": "...", "order_count": N}, ...],
      "breakfast": [...], "dessert": [...], "salads": [...],
      "main_meals": [{"meal_name": "...", "order_count": N}, ...]
    }
    وترجّع نتيجة كل صنف + الصوص المرتبط بأي وجبة من main_meals."""
    from station_calc import (
        calc_rice_item, calc_ingredient_item, calc_salad_item,
        calc_sauce_for_meal, STATION_ITEMS,
    )
    from smart_ordering import MEAL_PORTIONS

    payload = request.get_json(silent=True) or {}
    results = {'rice': [], 'breakfast': [], 'dessert': [], 'salads': [], 'sauce': []}
    errors = []

    for item in payload.get('rice', []):
        try:
            results['rice'].append(calc_rice_item(item['sheet_name'], item['order_count']))
        except Exception as e:
            errors.append({'station': 'rice', 'item': item.get('sheet_name'), 'error': str(e)})

    for item in payload.get('breakfast', []):
        try:
            results['breakfast'].append(calc_ingredient_item('breakfast', item['sheet_name'], item['order_count']))
        except Exception as e:
            errors.append({'station': 'breakfast', 'item': item.get('sheet_name'), 'error': str(e)})

    for item in payload.get('dessert', []):
        try:
            results['dessert'].append(calc_ingredient_item('dessert', item['sheet_name'], item['order_count']))
        except Exception as e:
            errors.append({'station': 'dessert', 'item': item.get('sheet_name'), 'error': str(e)})

    for item in payload.get('salads', []):
        try:
            results['salads'].append(calc_salad_item(item['sheet_name'], item['order_count']))
        except Exception as e:
            errors.append({'station': 'salads', 'item': item.get('sheet_name'), 'error': str(e)})

    for item in payload.get('main_meals', []):
        meal_name = item.get('meal_name')
        order_count = item.get('order_count')
        if meal_name in STATION_ITEMS['sauce'] and meal_name in MEAL_PORTIONS:
            try:
                portion_g = MEAL_PORTIONS[meal_name][0]
                results['sauce'].append(calc_sauce_for_meal(meal_name, order_count, portion_g))
            except Exception as e:
                errors.append({'station': 'sauce', 'item': meal_name, 'error': str(e)})

    return jsonify({'results': results, 'errors': errors})


_STATION_TITLE_FILL = PatternFill('solid', start_color='8C1810')
_STATION_TITLE_FONT = Font(name='Calibri', bold=True, color='FFFFFF', size=14)
_STATION_HEADER_FILL = PatternFill('solid', start_color='6600FF')
_STATION_HEADER_FONT = Font(name='Calibri', bold=True, color='FFFFFF')
_STATION_DATA_FONT = Font(name='Calibri', size=11)


def _build_station_workbook(items):
    """items: list of {'sheet_name','arabic_name','order_count','mode','rows':[...]}
    بيرجّع workbook فيه تاب لكل صنف، بانر بني + جدول بنفسجي، زي الستايل المعتمد
    في باقي النظام (نفس ألوان PURPLE_FILL المستخدمة في mega-purchasing)."""
    wb = Workbook()
    wb.remove(wb.active)
    for entry in items:
        title = (entry.get('sheet_name') or entry.get('meal_name') or 'صنف')[:31]
        safe_title = title
        n = 1
        existing = set(wb.sheetnames)
        while safe_title in existing:
            n += 1
            safe_title = f'{title[:28]} ({n})'
        ws = wb.create_sheet(title=safe_title)

        ar = entry.get('arabic_name', '')
        order_count = entry.get('order_count', '')
        ws.merge_cells('A1:D1')
        ws['A1'] = f"{entry.get('sheet_name') or entry.get('meal_name', '')}  {('— ' + ar) if ar else ''}  ({order_count} أوردر)"
        ws['A1'].fill = _STATION_TITLE_FILL
        ws['A1'].font = _STATION_TITLE_FONT
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 26

        is_batch = entry.get('mode') == 'batches'
        headers = ['البيان', 'Conversion Factor', 'Final KG'] if is_batch else ['المكوّن', 'الوحدة', 'الكمية']
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.fill = _STATION_HEADER_FILL
            cell.font = _STATION_HEADER_FONT
            cell.alignment = Alignment(horizontal='center')

        r = 4
        for row in entry.get('rows', []):
            if is_batch:
                ws.cell(row=r, column=1, value=row.get('label')).font = _STATION_DATA_FONT
                ws.cell(row=r, column=2, value=row.get('conversion_factor')).font = _STATION_DATA_FONT
                ws.cell(row=r, column=3, value=row.get('final_kg')).font = _STATION_DATA_FONT
            else:
                ws.cell(row=r, column=1, value=row.get('label')).font = _STATION_DATA_FONT
                ws.cell(row=r, column=2, value=row.get('unit')).font = _STATION_DATA_FONT
                ws.cell(row=r, column=3, value=row.get('amount')).font = _STATION_DATA_FONT
            r += 1

        ws.column_dimensions['A'].width = 36
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 16
        ws.sheet_view.rightToLeft = True

    if not wb.sheetnames:
        wb.create_sheet('فاضي')
    return wb


@app.route('/api/smart-order/export-stations', methods=['POST'])
def smart_order_export_stations():
    """بتاخد النتايج اللي اتحسبت بالفعل وظاهرة في الداش بورد (مش بتعيد حساب
    من الإكسل تاني — ده اللي كان بيسبب Worker Timeout لأنه بيفتح ملف توكيو
    الرئيسي (90 شيت) من الصفر لكل وجبة مرتين، مرة وقت الحساب ومرة وقت التصدير):
    {
      "main_results": [ {نفس شكل نتيجة calculate_meal}, ... ],   // الوجبات الأصلية
      "station_results": {
        "rice": [...], "breakfast": [...], "dessert": [...],
        "salads": [...], "sauce": [...]
      }
    }
    وترجّع zip فيه ملف إكسل لكل محطة، نفس الستايل المعتمد."""
    payload = request.get_json(silent=True) or {}

    results = {}
    main_results = payload.get('main_results') or []
    if main_results:
        results['hot_marination'] = main_results

    station_results = payload.get('station_results') or {}
    for key in ('rice', 'breakfast', 'dessert', 'salads', 'sauce'):
        items = station_results.get(key) or []
        if items:
            results[key] = items

    if not results:
        return jsonify({'error': 'مفيش نتايج اتحسبت عشان نصدّرها — احسب الإنتاج الأول'}), 400

    try:
        labels = {'rice': 'Rice', 'breakfast': 'Breakfast', 'dessert': 'Desserts', 'salads': 'Salads',
                  'sauce': 'Sauce', 'hot_marination': 'Hot_Marination'}
        zip_buf = io.BytesIO()
        today = datetime.now().strftime('%Y-%m-%d')
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for station, items in results.items():
                wb = _build_station_workbook(items)
                buf = io.BytesIO()
                wb.save(buf)
                zf.writestr(f'{labels[station]}_{today}.xlsx', buf.getvalue())

        zip_buf.seek(0)
        return send_file(zip_buf, as_attachment=True,
                          download_name=f'Production_Stations_{today}.zip',
                          mimetype='application/zip')
    except Exception as e:
        app.logger.exception('smart_order_export_stations failed')
        return jsonify({'error': f'حصل خطأ في التصدير: {e}'}), 500


if __name__ == '__main__':
    app.run(debug=True, port=8000)
