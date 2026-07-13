import os
import requests
import io
import re
import secrets
import base64
import shutil
import json
import tempfile
import zipfile
import smtplib
from email.message import EmailMessage
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
from parse_received_image import parse_received_image
from item_db import load_db, seed_from_order
from matcher import match_invoice_item
from db import get_client, execute_with_retry
from invoice_export import parse_invoice_full, build_invoices_workbook
from tokyo_ordering import read_day_file_meals, merge_day_into_template
from xlsx_to_images import add_workbook_images_to_zip

TOKYO_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'tokyo_ordering_template.xlsm')

# إعدادات إرسال الإيميل (لزرار "إرسال نسخة بالإيميل" في صفحة استلام الصوص)
SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.office365.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '587'))
SMTP_USER = os.environ.get('SMTP_USER')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD')
NOTIFY_EMAIL_TO = os.environ.get('NOTIFY_EMAIL_TO')

app = Flask(__name__)
CORS(app)  # allow calls from the Netlify frontend domain

from appointments_api import appointments_bp
app.register_blueprint(appointments_bp)


@app.after_request
def _ensure_cors_headers(response):
    """CORS(app) بيضيف الهيدرات دي للردود العادية بس - في حالات معينة (زي رد
    خطأ 400/500 راجع من جوه دالة، أو استثناء قبل ما الطلب يوصل للـ view function)
    الهيدرات ممكن متتحطش، فالمتصفح بيرفض حتى يعرض رسالة الخطأ الحقيقية ويطلع
    "Failed to fetch" بدل كده. الكود ده بيضمن إن كل رد (نجح أو فشل) شايل الهيدر."""
    origin = request.headers.get('Origin')
    if origin:
        response.headers.setdefault('Access-Control-Allow-Origin', origin)
        response.headers.setdefault('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Auth-Token')
        response.headers.setdefault('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS')
        response.headers.setdefault('Access-Control-Expose-Headers', 'X-Match-Report')
    return response


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


@app.route('/api/upload-order-preview', methods=['POST'])
def upload_order_preview():
    """بيقرأ ملفات الأوردر ويرجع الأصناف بدون ما يحفظ الكميات في الداتابيز —
    للاستخدام في أداة المطابقة السريعة (quick-match.html).
    ملحوظة: بيغذي (seed) قاعدة الأصناف الرئيسية master_items عشان مطابقة
    الفواتير في نفس الأداة تلاقي أصناف تتقارن بيها - من غيرها match_invoice_item
    بيرجع 'غير موجود' للكل لأن الـ db بيكون فاضي."""
    files = request.files.getlist('files')
    all_items = []
    errors = []
    for f in files:
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            f.save(tmp.name)
            path = tmp.name
        try:
            order = parse_order_pdf(path)
            seed_from_order(order)
            date_iso = _to_iso(order['date'])
            for section in ('salads', 'dressing'):
                for item in order[section]:
                    all_items.append({
                        'item_date': date_iso,
                        'item_key': item['name_en'].strip().upper(),
                        'name_en': item['name_en'],
                        'name_ar': item['name_ar'],
                        'section': section,
                        'qty_box': item['qty_box'],
                        'qty_needed': item['qty_needed'],
                        'unit': item['unit'].split('-')[0],
                        'current_inventory': item['current_inventory'],
                        'source_file': f.filename,
                    })
        except Exception as e:
            errors.append({'file': f.filename, 'error': str(e)})
        finally:
            os.unlink(path)
    return jsonify({'items': all_items, 'errors': errors})


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


@app.route('/api/tokyo-ordering/update-from-day-file', methods=['POST'])
def tokyo_ordering_update_from_day_file():
    """كارت محطة التجهيز: بترفع ملف يوم واحد بس (زي Octa_Food_Sat_....xlsx)،
    والسيستم بيعرف اليوم تلقائي من شيت Update بتاعه، ويطابق الأصناف مع
    عمود Meal name في All_Ingredients، ويحدّث Total Count/Total Grams بتاعت
    نفس اليوم بس في ملف توكيو الأساسي (بيتحفظ التحديث على السيرفر عشان
    الأيام اللي بترفعها بعد كده تتراكم على بعضها)، وبيرجّعلك الملف كامل
    بالماكرو والمعادلات زي ما هي، + تقرير بالأصناف اللي اتطابقت واللي لأ."""
    if not os.path.exists(TOKYO_TEMPLATE_PATH):
        return jsonify({'error': 'ملف القالب tokyo_ordering_template.xlsm غير موجود على السيرفر'}), 404
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'ارفع ملف يوم واحد (اسمه file في الطلب)'}), 400

    try:
        day_no, meals = read_day_file_meals(f)
    except Exception as e:
        return jsonify({'error': f'تعذّر قراءة ملف اليوم: {e}'}), 400

    try:
        out_path, report = merge_day_into_template(TOKYO_TEMPLATE_PATH, day_no, meals)
        shutil.copyfile(out_path, TOKYO_TEMPLATE_PATH)  # حفظ التحديث على القالب نفسه عشان يتراكم
    except Exception as e:
        app.logger.exception('tokyo_ordering_update_from_day_file failed')
        return jsonify({'error': f'حصل خطأ أثناء الدمج: {e}'}), 500

    response = send_file(out_path, as_attachment=True,
                          download_name=f"Tokyo_Ordering_Updated_{report['day_name']}.xlsm")
    # مهم: هيدرات HTTP لازم تكون ASCII بس - النص العربي في التقرير لازم يتحوّل
    # لصيغة \uXXXX (ensure_ascii=True) وإلا السيرفر (gunicorn) بيرفض يبعت الرد
    # كله بخطأ "Invalid HTTP Header" والمتصفح بيشوفه فشل اتصال تام (CORS مضلِّل).
    response.headers['X-Match-Report'] = json.dumps(report, ensure_ascii=True)
    return response


@app.route('/api/sauce-receipt/list', methods=['GET'])
def sauce_receipt_list():
    """قايمة كل روابط الاستلام (الأحدث الأول) - محتاجة تسجيل دخول، تستخدمها
    صفحة sauce-notifications.html عشان تعرض لك أول ما حد يبعت البيانات."""
    username, err = _require_auth()
    if err:
        return err
    sb = get_client()
    res = execute_with_retry(
        sb.table('sauce_receipts').select('*').order('created_at', desc=True).limit(100)
    )
    return jsonify({'receipts': res.data or []})


@app.route('/api/sauce-receipt/<receipt_id>', methods=['DELETE'])
def sauce_receipt_delete(receipt_id):
    """حذف رابط استلام بالكامل - محتاج تسجيل دخول."""
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    execute_with_retry(sb.table('sauce_receipts').delete().eq('id', receipt_id))
    return jsonify({'ok': True})


@app.route('/api/sauce-receipt/<receipt_id>/reopen', methods=['POST'])
def sauce_receipt_reopen(receipt_id):
    """يرجّع رابط اتبعت خلاص لحالة 'pending' تاني عشان يتملى/يتعدّل من الأول -
    محتاج تسجيل دخول. مش بتمسح آخر بيانات مُرسلة (submitted_days) خالص، بس
    بتفتح الرابط يقبل إرسال جديد يستبدلها."""
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    execute_with_retry(sb.table('sauce_receipts').update({
        'status': 'pending',
    }).eq('id', receipt_id))
    return jsonify({'ok': True})


@app.route('/api/sauce-receipt/create', methods=['POST'])
def sauce_receipt_create():
    """بتاخد قايمة أيام/صفوف الصوص (اللي طلعت من زرار استخراج الصوص) وتعمل
    سجل جديد في sauce_receipts، وترجّع id تستخدمه في بناء رابط تبعته للمسؤول
    على واتساب يدويًا (زي: pixivo.org/sauce-receipt.html?id=...)."""
    payload = request.get_json(silent=True) or {}
    days = payload.get('days') or []
    if not days:
        return jsonify({'error': 'مفيش بيانات صوص مبعوتة'}), 400
    sb = get_client()
    res = execute_with_retry(sb.table('sauce_receipts').insert({
        'days': days, 'status': 'pending',
    }))
    row_id = res.data[0]['id']
    return jsonify({'id': row_id})


@app.route('/api/sauce-receipt/<receipt_id>', methods=['GET'])
def sauce_receipt_get(receipt_id):
    """اندبوينت عام (من غير تسجيل دخول) - صفحة الاستلام sauce-receipt.html
    بتناديه عشان تعرض للمسؤول الأصناف اللي محتاجة يملأ كمياتها.
    بيرجّع submitted_days كمان (لو موجودة) عشان الصفحة تقدر تعرض آخر قيم
    اتبعتت لكل يوم ووقت آخر تعديل، وتسيب المستخدم يعدّل تاني براحته."""
    sb = get_client()
    res = execute_with_retry(sb.table('sauce_receipts').select('*').eq('id', receipt_id))
    rows = res.data or []
    if not rows:
        return jsonify({'error': 'الرابط ده مش موجود أو انتهى'}), 404
    receipt = rows[0]
    submitted_days = receipt.get('submitted_days') or {}
    if isinstance(submitted_days, list):
        # صيغة قديمة (List) من قبل التحديث - نتجاهلها ونبدأ فاضية بالصيغة الجديدة (dict لكل يوم)
        submitted_days = {}
    return jsonify({
        'id': receipt['id'], 'days': receipt['days'], 'status': receipt['status'],
        'submitted_days': submitted_days,
        'created_at': receipt['created_at'],
    })


@app.route('/api/sauce-receipt/<receipt_id>/submit-day', methods=['POST'])
def sauce_receipt_submit_day(receipt_id):
    """بتسجّل استلام يوم واحد بس من الرابط، من غير ما تقفل باقي الأيام أو
    تمنع تعديل اليوم ده تاني بعدين - الرابط يفضل قابل للفتح والتعديل أي وقت.
    كل مرة يتبعت فيها نفس اليوم، بيتسجّل وقت وتاريخ التعديل، وبيوصل إشعار
    جديد للوحة التحكم (يظهر في sauce-notifications.html) يوضّح لو ده أول
    إرسال أو تعديل على إرسال سابق."""
    payload = request.get_json(silent=True) or {}
    day_name = payload.get('day')
    submitted_rows = payload.get('rows') or []
    if not day_name or not submitted_rows:
        return jsonify({'error': 'محتاج اسم اليوم وبيانات الصفوف'}), 400

    sb = get_client()
    res = execute_with_retry(sb.table('sauce_receipts').select('*').eq('id', receipt_id))
    found = res.data or []
    if not found:
        return jsonify({'error': 'الرابط ده مش موجود أو انتهى'}), 404
    receipt = found[0]

    # اربط الصفوف المُستلَمة بالبيانات الأصلية المتوقعة لنفس اليوم واحسب الزيادة/النقص
    expected_by_key = {}
    for day in receipt['days']:
        if day.get('day') == day_name:
            for r in day.get('rows', []):
                expected_by_key[r['key']] = r
            break

    result_rows = []
    summary_lines = []
    for r in submitted_rows:
        exp = expected_by_key.get(r.get('key'), {})
        pm_expected = exp.get('expectedProteinMix') or 0
        tp_expected = exp.get('expectedTopping')
        pm_received = r.get('proteinMixReceived') or 0
        tp_received = r.get('toppingReceived') or 0
        total_expected = pm_expected + (tp_expected or 0)
        total_received = pm_received + (tp_received or 0)
        excess = max(0, total_received - total_expected)
        shortage = max(0, total_expected - total_received)
        result_rows.append({
            'key': r.get('key'), 'title': exp.get('title', ''),
            'proteinMixReceived': pm_received, 'toppingReceived': tp_received,
            'excess': excess, 'shortage': shortage,
        })
        if excess or shortage:
            summary_lines.append(
                f"{exp.get('title','')}: "
                f"{'زيادة ' + str(round(excess,2)) if excess else ''}"
                f"{'نقص ' + str(round(shortage,2)) if shortage else ''}"
            )

    now_iso = datetime.now(timezone.utc).isoformat()
    submitted_days = receipt.get('submitted_days') or {}
    if isinstance(submitted_days, list):
        submitted_days = {}
    is_edit = day_name in submitted_days
    prev_edit_count = (submitted_days.get(day_name) or {}).get('edit_count', 0)
    submitted_days[day_name] = {
        'rows': result_rows,
        'submitted_at': now_iso,
        'edit_count': prev_edit_count + 1,
    }

    execute_with_retry(sb.table('sauce_receipts').update({
        'submitted_days': submitted_days,
        'submitted_at': now_iso,  # وقت آخر تعديل عمومًا على الرابط كله
    }).eq('id', receipt_id))

    day_time_cairo = (datetime.now(timezone.utc) + timedelta(hours=3)).strftime('%Y-%m-%d %I:%M %p')
    action_word = 'تعديل' if is_edit else 'استلام'
    notice = f'تم {action_word} صوص يوم {day_name} — الساعة {day_time_cairo} (توقيت القاهرة)'
    if is_edit:
        notice += f' — ده تعديل رقم {prev_edit_count + 1} على نفس اليوم'
    if summary_lines:
        notice += '\nفيه فروقات محتاجة مراجعة:\n' + '\n'.join(summary_lines)
    _log('sauce_receipt', f'رابط استلام {receipt_id[:8]} - يوم {day_name}', None, notice,
         level='warning' if summary_lines else 'info')

    return jsonify({'ok': True, 'day': day_name, 'submitted_at': now_iso, 'is_edit': is_edit,
                     'has_discrepancy': bool(summary_lines)})


@app.route('/api/sauce-receipt/<receipt_id>/email-day', methods=['POST'])
def sauce_receipt_email_day(receipt_id):
    """بتاخد ملف الإكسيل بتاع يوم واحد (نفس الملف اللي بينزل عند المستخدم بالظبط،
    مبعوت من الفرونت إند كـ multipart file) وتبعته بالإيميل مرفق مباشرة للإيميل
    اللي كتبه العامل نفسه في خانة "إيميل المستلم" — من غير ما تفتح أي برنامج
    ميل، الإرسال بيتم من السيرفر على طول. لو NOTIFY_EMAIL_TO متظبطة، بتتبعتلها
    نسخة BCC كمان (سجلّ عندك) بدون ما تظهر للمستلم الأساسي."""
    if not (SMTP_USER and SMTP_PASSWORD):
        return jsonify({'error': 'إعدادات الإيميل لسه مش متظبطة على السيرفر (SMTP_USER / SMTP_PASSWORD)'}), 503

    day_name = request.form.get('day', '')
    to_email = (request.form.get('to') or '').strip()
    if not to_email or '@' not in to_email:
        return jsonify({'error': 'إيميل المستلم ناقص أو غير صحيح'}), 400

    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'مفيش ملف مرفوع'}), 400

    file_bytes = file.read()

    msg = EmailMessage()
    msg['Subject'] = f'استلام الصوص — يوم {day_name} — {datetime.now().strftime("%Y-%m-%d")}'
    msg['From'] = SMTP_USER
    msg['To'] = to_email
    if NOTIFY_EMAIL_TO:
        msg['Bcc'] = NOTIFY_EMAIL_TO
    msg.set_content(
        f'تم استلام صوص يوم {day_name}.\n'
        f'الملف المرفق فيه كل التفاصيل (المتوقع، المستلم فعليًا، الزيادة والنقص).\n'
        f'رقم رابط الاستلام: {receipt_id}'
    )
    msg.add_attachment(
        file_bytes,
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=file.filename or f'sauce-receipt-{day_name}.xlsx',
    )

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(SMTP_USER, SMTP_PASSWORD)
            smtp.send_message(msg)
    except Exception as e:
        app.logger.exception('sauce_receipt_email_day failed to send email')
        return jsonify({'error': f'تعذر إرسال الإيميل: {e}'}), 500

    _log('sauce_receipt', f'رابط استلام {receipt_id[:8]} - يوم {day_name}', None,
         f'اتبعتت نسخة بالإيميل لملف يوم {day_name} إلى {to_email}', level='info')

    return jsonify({'ok': True})


@app.route('/api/sauce-receipt/<receipt_id>/submit', methods=['POST'])
def sauce_receipt_submit(receipt_id):
    """[أقدم إصدار - بيستقبل كل الأيام مرة واحدة] سايبها شغالة للتوافق مع
    أي نسخة قديمة مفتوحة عند حد، بس sauce-receipt.html الجديدة بتستخدم
    /submit-day بدل منها (إرسال كل يوم لوحده، وتسمح بالتعديل بعد كده)."""
    payload = request.get_json(silent=True) or {}
    submitted_days = payload.get('days') or []
    if not submitted_days:
        return jsonify({'error': 'مفيش بيانات استلام مبعوتة'}), 400

    sb = get_client()
    res = execute_with_retry(sb.table('sauce_receipts').select('*').eq('id', receipt_id))
    rows = res.data or []
    if not rows:
        return jsonify({'error': 'الرابط ده مش موجود أو انتهى'}), 404
    receipt = rows[0]

    # اربط كل صف مُستلَم بالصف الأصلي (بالـ key) واحسب الزيادة/النقص
    expected_by_key = {}
    for day in receipt['days']:
        for r in day.get('rows', []):
            expected_by_key[r['key']] = r

    result_days = []
    summary_lines = []
    for day in submitted_days:
        day_name = day.get('day', '')
        result_rows = []
        for r in day.get('rows', []):
            exp = expected_by_key.get(r.get('key'), {})
            pm_expected = exp.get('expectedProteinMix') or 0
            tp_expected = exp.get('expectedTopping')
            pm_received = r.get('proteinMixReceived') or 0
            tp_received = r.get('toppingReceived') or 0
            total_expected = pm_expected + (tp_expected or 0)
            total_received = pm_received + (tp_received or 0)
            excess = max(0, total_received - total_expected)
            shortage = max(0, total_expected - total_received)
            result_rows.append({
                'key': r.get('key'), 'title': exp.get('title', ''),
                'proteinMixReceived': pm_received, 'toppingReceived': tp_received,
                'excess': excess, 'shortage': shortage,
            })
            if excess or shortage:
                summary_lines.append(
                    f"{day_name} - {exp.get('title','')}: "
                    f"{'زيادة ' + str(round(excess,2)) if excess else ''}"
                    f"{'نقص ' + str(round(shortage,2)) if shortage else ''}"
                )
        result_days.append({'day': day_name, 'rows': result_rows})

    execute_with_retry(sb.table('sauce_receipts').update({
        'status': 'submitted',
        'submitted_days': result_days,
        'submitted_at': datetime.now(timezone.utc).isoformat(),
    }).eq('id', receipt_id))

    notice = 'تم استلام الصوص بالكامل، كل الكميات مطابقة ✅' if not summary_lines else (
        'تم استلام الصوص - فيه فروقات محتاجة مراجعة:\n' + '\n'.join(summary_lines))
    _log('sauce_receipt', f'رابط استلام {receipt_id[:8]}', None, notice,
         level='warning' if summary_lines else 'info')

    return jsonify({'ok': True, 'has_discrepancy': bool(summary_lines)})


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


# ============================================================
# نظام إدارة نصوص السيستم (Texts Dashboard) — Supabase table: system_texts
# key TEXT PRIMARY KEY, value TEXT, page TEXT, updated_at, updated_by
# ============================================================
@app.route('/api/texts', methods=['GET'])
def get_texts():
    """بيرجّع كل النصوص كـ map { key: value }. مفيش auth هنا عشان أي صفحة
    (حتى صفحة اللوجين) تقدر تجيب النصوص بتاعتها من غير ما تتوقف على تسجيل الدخول."""
    sb = get_client()
    res = execute_with_retry(sb.table('system_texts').select('key, value'))
    texts = {row['key']: row['value'] for row in (res.data or [])}
    return jsonify({'texts': texts})


@app.route('/api/texts', methods=['PUT'])
def update_texts():
    """بيحفظ تعديلات على نص أو أكتر دفعة واحدة. Body: { "texts": { "key": "value", ... } }"""
    username, err = _require_auth()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    updates = payload.get('texts') or {}
    if not isinstance(updates, dict) or not updates:
        return jsonify({'error': 'مفيش نصوص للحفظ'}), 400

    sb = get_client()
    now = datetime.now(timezone.utc).isoformat()
    rows = [
        {'key': k, 'value': v, 'page': k.split('.')[0], 'updated_at': now, 'updated_by': username}
        for k, v in updates.items()
    ]
    try:
        execute_with_retry(sb.table('system_texts').upsert(rows, on_conflict='key'))
    except Exception as e:
        return jsonify({'error': f'تعذر حفظ النصوص: {e}'}), 400
    return jsonify({'ok': True, 'count': len(rows)})


# ============================================================
# إعدادات ثيم السيستم (ألوان، خط، أنيميشن، وضع ليلي) — صف واحد بس، بيتطبق
# على كل صفحات النظام اللي فيها texts-runtime.js تلقائي
# ============================================================
DEFAULT_THEME = {
    'primary_color': '#EC1510', 'primary_dark_color': '#C01210', 'ink_color': '#1A1A1A',
    'muted_color': '#8A8A8A', 'soft_color': '#FFF6F5', 'line_color': '#F1D8D6', 'ok_color': '#1F8A4C',
    'font_family': 'Tahoma, Arial, sans-serif', 'font_label': 'Tahoma (افتراضي)',
    'animations_enabled': True, 'dark_mode_enabled': False,
    'dark_bg': '#17120F', 'dark_surface': '#241C17', 'dark_text': '#F2EAE4',
    'dark_muted': '#B3A79C', 'dark_border': '#3B2E25',
}


@app.route('/api/theme', methods=['GET'])
def get_theme():
    """عام من غير لوجين - كل صفحة بتجيب إعدادات الثيم منه."""
    sb = get_client()
    res = execute_with_retry(sb.table('system_theme').select('*').eq('id', 1))
    rows = res.data or []
    theme = {**DEFAULT_THEME, **(rows[0] if rows else {})}
    return jsonify({'theme': theme})


@app.route('/api/theme', methods=['PUT'])
def update_theme():
    """محمي بتسجيل الدخول - تحديث إعدادات الثيم من داش بورد التصميم."""
    username, err = _require_auth()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    allowed_keys = set(DEFAULT_THEME.keys())
    updates = {k: v for k, v in payload.items() if k in allowed_keys}
    if not updates:
        return jsonify({'error': 'مفيش إعدادات للحفظ'}), 400
    updates['updated_at'] = datetime.now(timezone.utc).isoformat()
    updates['updated_by'] = username

    sb = get_client()
    try:
        execute_with_retry(sb.table('system_theme').update(updates).eq('id', 1))
    except Exception as e:
        return jsonify({'error': f'تعذر الحفظ: {e}'}), 400
    return jsonify({'ok': True})


# ============================================================
# سجل الأصناف والأوزان اليدوي (Weight Log) — لينك ثابت للعامل من غير لوجين
# ============================================================
# التوكين ده جزء من اللينك اللي بيتبعت للعامل مرة واحدة ويفضل يستخدمه يوميًا.
# لو حبيت تغيّره في أي وقت (مثلاً لو حد غريب وصله)، غيّر القيمة دي وابعت
# للعامل لينك جديد بالتوكين الجديد.
WEIGHT_LOG_TOKEN = 'pNrAYo0cIwXhdsgVdXKSJYCGAS8'

# توكين منفصل تمامًا لمركز التخزين (عرض بس، من غير لوجين) - ابعته لأي حد
# عايزه يطّلع على الأرشيف كامل من غير ما يدخل السيستم خالص. لو حبيت تسحب
# صلاحية حد وصله بالغلط، غيّر القيمة دي وهيبقى معاه لينك قديم مبيشتغلش.
WEIGHT_LOG_VIEW_TOKEN = 'vXq3mZpLd8RwTfKhY0eB2nCsUj7A'


def _weight_log_worker_ok():
    token = request.values.get('token') or (request.form.get('token') if request.method != 'GET' else None)
    return bool(token) and token == WEIGHT_LOG_TOKEN


def _weight_log_edit_authorized():
    """التعديل/الحذف مسموح إما بتوكين العامل نفسه، أو بتسجيل دخول الأدمن."""
    if _weight_log_worker_ok():
        return True
    auth_token = request.headers.get('X-Auth-Token')
    return bool(auth_token and _check_session(auth_token))


def _weight_log_day_bounds_utc(offset_days=0):
    """بترجع (start_utc_iso, end_utc_iso) لبداية ونهاية يوم بتوقيت السعودية
    (UTC+3 ثابت - مفيش توقيت صيفي)، عشان نفلتر 'إنهاردة' صح في نظام الأصناف
    والأوزان تحديدًا."""
    from datetime import timedelta
    ksa_now = datetime.now(timezone.utc) + timedelta(hours=3) + timedelta(days=offset_days)
    day_start_ksa = ksa_now.replace(hour=0, minute=0, second=0, microsecond=0)
    day_end_ksa = day_start_ksa + timedelta(days=1)
    start_utc = day_start_ksa - timedelta(hours=3)
    end_utc = day_end_ksa - timedelta(hours=3)
    return start_utc.isoformat(), end_utc.isoformat()


@app.route('/api/weight-log', methods=['POST'])
def weight_log_add():
    """بيستقبل صنف واحد (اسم + وزن + صورة اختيارية) من صفحة العامل. من غير
    لوجين، بس محمي بتوكين ثابت في اللينك نفسه."""
    if not _weight_log_worker_ok():
        return jsonify({'error': 'الرابط ده مش صحيح أو قديم'}), 403

    item_name = (request.form.get('item_name') or '').strip()
    weight_raw = (request.form.get('weight') or '').strip()
    if not item_name:
        return jsonify({'error': 'اكتب اسم الصنف'}), 400
    try:
        weight_val = float(weight_raw)
    except (TypeError, ValueError):
        return jsonify({'error': 'الوزن لازم يكون رقم'}), 400

    photo_b64 = None
    photo_file = request.files.get('photo')
    if photo_file and photo_file.filename:
        photo_bytes = photo_file.read()
        if len(photo_bytes) > 6 * 1024 * 1024:
            return jsonify({'error': 'الصورة كبيرة جدًا (أكبر من 6 ميجا)'}), 400
        mime = photo_file.mimetype or 'image/jpeg'
        photo_b64 = f'data:{mime};base64,' + base64.b64encode(photo_bytes).decode('ascii')

    sb = get_client()
    row = {
        'item_name': item_name,
        'weight': weight_val,
        'photo_base64': photo_b64,
        'logged_at': datetime.now(timezone.utc).isoformat(),
        'deleted': False,
    }
    try:
        res = execute_with_retry(sb.table('weight_log_entries').insert(row))
        new_row = (res.data or [{}])[0]
    except Exception as e:
        return jsonify({'error': f'تعذر الحفظ: {e}'}), 400
    return jsonify({'ok': True, 'id': new_row.get('id')})


@app.route('/api/weight-log/mine', methods=['GET'])
def weight_log_mine():
    """للعامل بس - بترجّع أصناف إنهاردة اللي هو سجّلها (بتوقيت القاهرة)،
    عشان يقدر يراجع اللي بعته وهو لسه فاتح نفس اللينك، حتى لو قفل الصفحة
    وفتحها تاني. من غير لوجين، بتوكين العامل بس."""
    if not _weight_log_worker_ok():
        return jsonify({'error': 'الرابط ده مش صحيح أو قديم'}), 403
    start_iso, end_iso = _weight_log_day_bounds_utc()
    sb = get_client()
    res = execute_with_retry(
        sb.table('weight_log_entries').select('id, item_name, weight, photo_base64, logged_at')
        .gte('logged_at', start_iso).lt('logged_at', end_iso)
        .eq('deleted', False)
        .order('logged_at', desc=True)
    )
    return jsonify({'entries': res.data or []})


@app.route('/api/weight-log', methods=['GET'])
def weight_log_list():
    """للداش بورد الشغّالة بتاعتك - محمي بتسجيل الدخول. بترجّع الأصناف
    الغير محذوفة بس (اللي اتمسح هيفضل موجود في مركز التخزين بس مش هنا)."""
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    res = execute_with_retry(
        sb.table('weight_log_entries').select('id, item_name, weight, photo_base64, logged_at')
        .eq('deleted', False)
        .order('logged_at', desc=True)
    )
    return jsonify({'entries': res.data or []})


@app.route('/api/weight-log/archive', methods=['GET'])
def weight_log_archive():
    """مركز التخزين - بيرجّع كل حاجة اتسجلت على الإطلاق (حتى اللي اتمسح من
    الداش بورد الشغّالة)، عشان الأرشيف يفضل كامل زي ما هو دايمًا. محمي
    بتوكين عرض منفصل تمامًا عن توكين العامل، من غير لوجين."""
    token = request.args.get('view_token')
    if not token or token != WEIGHT_LOG_VIEW_TOKEN:
        return jsonify({'error': 'الرابط ده مش صحيح'}), 403
    sb = get_client()
    res = execute_with_retry(
        sb.table('weight_log_entries').select('id, item_name, weight, photo_base64, logged_at, deleted')
        .order('logged_at', desc=True)
    )
    return jsonify({'entries': res.data or []})


@app.route('/api/weight-log/<int:entry_id>', methods=['PUT'])
def weight_log_update(entry_id):
    """تعديل اسم الصنف أو الوزن - مسموح للعامل (بتوكينه) أو للأدمن (بلوجينه)،
    قبل الإرسال أو بعده، وبيتحدّث في كل الأماكن (الداش بورد ومركز التخزين)."""
    if not _weight_log_edit_authorized():
        return jsonify({'error': 'مش مسموح'}), 403
    payload = request.get_json(silent=True) or {}
    updates = {}
    if 'item_name' in payload:
        name = (payload.get('item_name') or '').strip()
        if not name:
            return jsonify({'error': 'اسم الصنف مينفعش يبقى فاضي'}), 400
        updates['item_name'] = name
    if 'weight' in payload:
        try:
            updates['weight'] = float(payload.get('weight'))
        except (TypeError, ValueError):
            return jsonify({'error': 'الوزن لازم يكون رقم'}), 400
    if not updates:
        return jsonify({'error': 'مفيش حاجة للتعديل'}), 400

    sb = get_client()
    try:
        execute_with_retry(sb.table('weight_log_entries').update(updates).eq('id', entry_id))
    except Exception as e:
        return jsonify({'error': f'تعذر التعديل: {e}'}), 400
    return jsonify({'ok': True})


@app.route('/api/weight-log/<int:entry_id>', methods=['DELETE'])
def weight_log_delete(entry_id):
    """حذف ناعم بس - الصنف بيختفي من الداش بورد الشغّالة ومن صفحة العامل،
    بس بيفضل موجود في مركز التخزين للأبد. مسموح للعامل (بتوكينه) أو للأدمن
    (بلوجينه)."""
    if not _weight_log_edit_authorized():
        return jsonify({'error': 'مش مسموح'}), 403
    sb = get_client()
    try:
        execute_with_retry(sb.table('weight_log_entries').update({'deleted': True}).eq('id', entry_id))
    except Exception as e:
        return jsonify({'error': f'تعذر الحذف: {e}'}), 400
    return jsonify({'ok': True})


@app.route('/api/weight-log/<int:entry_id>/photo', methods=['GET'])
def weight_log_photo(entry_id):
    """بترجّع الصورة كملف مباشر (مش base64) - مستخدمة كلينك جوه ملف الإكسيل
    المُصدَّر، عشان الفايل يفضل خفيف بدل ما يشيل الصور جواه."""
    sb = get_client()
    res = execute_with_retry(
        sb.table('weight_log_entries').select('photo_base64').eq('id', entry_id)
    )
    rows = res.data or []
    if not rows or not rows[0].get('photo_base64'):
        return jsonify({'error': 'مفيش صورة للصنف ده'}), 404
    data_url = rows[0]['photo_base64']
    try:
        header, b64data = data_url.split(',', 1)
        mime = header.split(':')[1].split(';')[0]
        img_bytes = base64.b64decode(b64data)
    except Exception:
        return jsonify({'error': 'الصورة تالفة'}), 400
    return send_file(io.BytesIO(img_bytes), mimetype=mime)


# ============================================================
# أصناف كل يوم (Weight Log Items) — القايمة اللي العامل يختار منها بدل
# ما يكتب اسم الصنف حر، مأخوذة من ملف "مشروع صدى" بالترتيب بالظبط
# ============================================================
WEIGHT_LOG_DAYS = ['السبت', 'الأحد', 'الاثنين', 'الثلاثاء', 'الاربعاء', 'خميس', 'الجمعة']


@app.route('/api/weight-log/items', methods=['GET'])
def weight_log_items_list():
    """بترجّع أصناف يوم معيّن بالترتيب بالظبط زي ما هو متسجل. للعامل بس
    (بتوكينه)."""
    if not _weight_log_worker_ok():
        return jsonify({'error': 'الرابط ده مش صحيح أو قديم'}), 403
    day = (request.args.get('day') or '').strip()
    if not day:
        return jsonify({'error': 'حدد اليوم'}), 400
    sb = get_client()
    res = execute_with_retry(
        sb.table('weight_log_items').select('id, item_name')
        .eq('day_name', day).order('sort_order')
    )
    return jsonify({'items': res.data or []})


@app.route('/api/weight-log/items', methods=['POST'])
def weight_log_items_add():
    """العامل بيضيف صنف جديد لليوم ده - بيتسجل في القائمة عشان يفضل موجود
    لاستخدامه تاني في أي يوم زي ده جاي. لو الاسم موجود بالفعل (حتى بحروف
    مختلفة كبيرة/صغيرة) بيرجّع نفس الصنف الموجود من غير ما يكرره."""
    if not _weight_log_worker_ok():
        return jsonify({'error': 'الرابط ده مش صحيح أو قديم'}), 403
    payload = request.get_json(silent=True) or {}
    day = (payload.get('day') or '').strip()
    item_name = (payload.get('item_name') or '').strip()
    if not day or not item_name:
        return jsonify({'error': 'محتاج اليوم واسم الصنف'}), 400

    sb = get_client()
    existing = execute_with_retry(
        sb.table('weight_log_items').select('id, item_name')
        .eq('day_name', day).ilike('item_name', item_name)
    )
    if existing.data:
        return jsonify({'ok': True, 'item': existing.data[0]})

    max_order_res = execute_with_retry(
        sb.table('weight_log_items').select('sort_order')
        .eq('day_name', day).order('sort_order', desc=True).limit(1)
    )
    next_order = (max_order_res.data[0]['sort_order'] + 1) if max_order_res.data else 0
    try:
        res = execute_with_retry(
            sb.table('weight_log_items').insert(
                {'day_name': day, 'item_name': item_name, 'sort_order': next_order}
            )
        )
    except Exception as e:
        return jsonify({'error': f'تعذر إضافة الصنف: {e}'}), 400
    return jsonify({'ok': True, 'item': (res.data or [{}])[0]})


# ============================================================
# مخزون الخضار اليومي للأقسام (Veg Inventory) — لينك ثابت للعامل من غير
# لوجين، نموذج واحد لكل يوم قابل للتحديث طول اليوم (زي شيت الصوص بالظبط)
# ============================================================
VEG_INVENTORY_TOKEN = 'hR7wKqLm2XdNpTs9BvYcGz4eAf6J'


def _veg_inventory_worker_ok():
    token = request.values.get('token') or (request.get_json(silent=True) or {}).get('token')
    return bool(token) and token == VEG_INVENTORY_TOKEN


def _veg_inventory_edit_authorized():
    if _veg_inventory_worker_ok():
        return True
    auth_token = request.headers.get('X-Auth-Token')
    return bool(auth_token and _check_session(auth_token))


def _riyadh_today_date():
    from datetime import timedelta
    return (datetime.now(timezone.utc) + timedelta(hours=3)).date().isoformat()


@app.route('/api/veg-inventory/items', methods=['GET'])
def veg_inventory_items_list():
    """قايمة الأصناف الثابتة (خضروات/أعشاب/فواكه) بالترتيب - للعامل بتوكينه."""
    if not _veg_inventory_worker_ok():
        return jsonify({'error': 'الرابط ده مش صحيح أو قديم'}), 403
    sb = get_client()
    res = execute_with_retry(
        sb.table('veg_inventory_items').select('id, item_name, category, unit').order('sort_order')
    )
    return jsonify({'items': res.data or []})


@app.route('/api/veg-inventory/today', methods=['GET'])
def veg_inventory_today_get():
    """بترجّع قيم إنهاردة المحفوظة لحد دلوقتي (لو العامل رجع يعدّل) - للعامل
    بتوكينه."""
    if not _veg_inventory_worker_ok():
        return jsonify({'error': 'الرابط ده مش صحيح أو قديم'}), 403
    today = _riyadh_today_date()
    sb = get_client()
    res = execute_with_retry(
        sb.table('veg_inventory_entries').select('item_name, remaining_stock, updated_at')
        .eq('entry_date', today)
    )
    rows = res.data or []
    last_updated = max((r['updated_at'] for r in rows), default=None)
    return jsonify({
        'date': today,
        'entries': {r['item_name']: r['remaining_stock'] for r in rows},
        'last_updated': last_updated,
    })


@app.route('/api/veg-inventory/today', methods=['POST'])
def veg_inventory_today_save():
    """العامل بيحفظ/يحدّث قيم إنهاردة - upsert لكل صنف مبعوت. Body:
    { "token": "...", "entries": { "اسم الصنف": 1200, ... } }"""
    if not _veg_inventory_worker_ok():
        return jsonify({'error': 'الرابط ده مش صحيح أو قديم'}), 403
    payload = request.get_json(silent=True) or {}
    entries = payload.get('entries') or {}
    if not isinstance(entries, dict) or not entries:
        return jsonify({'error': 'مفيش قيم للحفظ'}), 400

    today = _riyadh_today_date()
    now = datetime.now(timezone.utc).isoformat()
    rows = []
    for item_name, value in entries.items():
        if value is None or str(value).strip() == '':
            continue
        try:
            val = float(value)
        except (TypeError, ValueError):
            return jsonify({'error': f'قيمة غير صحيحة للصنف "{item_name}"'}), 400
        rows.append({'entry_date': today, 'item_name': item_name, 'remaining_stock': val, 'updated_at': now})

    if not rows:
        return jsonify({'error': 'مفيش قيم صحيحة للحفظ'}), 400

    sb = get_client()
    try:
        execute_with_retry(
            sb.table('veg_inventory_entries').upsert(rows, on_conflict='entry_date,item_name')
        )
    except Exception as e:
        return jsonify({'error': f'تعذر الحفظ: {e}'}), 400
    return jsonify({'ok': True, 'date': today, 'updated_at': now, 'count': len(rows)})


@app.route('/api/veg-inventory', methods=['GET'])
def veg_inventory_list_all():
    """للداش بورد بتاعتك - محمي بتسجيل الدخول. بترجّع كل الأيام المسجلة."""
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    entries_res = execute_with_retry(
        sb.table('veg_inventory_entries').select('id, entry_date, item_name, remaining_stock, updated_at')
        .order('entry_date', desc=True)
    )
    items_res = execute_with_retry(
        sb.table('veg_inventory_items').select('item_name, category, unit').order('sort_order')
    )
    return jsonify({'entries': entries_res.data or [], 'items': items_res.data or []})


@app.route('/api/veg-inventory/entry/<int:entry_id>', methods=['PUT'])
def veg_inventory_entry_update(entry_id):
    """تعديل قيمة صنف في يوم معيّن - محمي بتسجيل الدخول (الداش بورد بتاعتك)."""
    _, err = _require_auth()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    if 'remaining_stock' not in payload:
        return jsonify({'error': 'مفيش قيمة للتعديل'}), 400
    try:
        val = float(payload.get('remaining_stock'))
    except (TypeError, ValueError):
        return jsonify({'error': 'القيمة لازم تكون رقم'}), 400

    sb = get_client()
    try:
        execute_with_retry(
            sb.table('veg_inventory_entries')
            .update({'remaining_stock': val, 'updated_at': datetime.now(timezone.utc).isoformat()})
            .eq('id', entry_id)
        )
    except Exception as e:
        return jsonify({'error': f'تعذر التعديل: {e}'}), 400
    return jsonify({'ok': True})


@app.route('/api/veg-inventory/entry/<int:entry_id>', methods=['DELETE'])
def veg_inventory_entry_delete(entry_id):
    """حذف قيمة صنف اتسجلت غلط ليوم معيّن - محمي بتسجيل الدخول."""
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    try:
        execute_with_retry(sb.table('veg_inventory_entries').delete().eq('id', entry_id))
    except Exception as e:
        return jsonify({'error': f'تعذر الحذف: {e}'}), 400
    return jsonify({'ok': True})


@app.route('/api/veg-inventory/day/<entry_date>', methods=['PUT'])
def veg_inventory_day_update(entry_date):
    """تعديل يوم كامل دفعة واحدة من الداش بورد - محمي بتسجيل الدخول.
    Body: { "entries": { "اسم الصنف": 1200, ... } } - بيعمل upsert لكل صنف
    مبعوت، ومش بيلمس الأصناف اللي مبعتتش."""
    _, err = _require_auth()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    entries = payload.get('entries') or {}
    if not isinstance(entries, dict) or not entries:
        return jsonify({'error': 'مفيش قيم للحفظ'}), 400

    now = datetime.now(timezone.utc).isoformat()
    rows = []
    for item_name, value in entries.items():
        if value is None or str(value).strip() == '':
            continue
        try:
            val = float(value)
        except (TypeError, ValueError):
            return jsonify({'error': f'قيمة غير صحيحة للصنف "{item_name}"'}), 400
        rows.append({'entry_date': entry_date, 'item_name': item_name, 'remaining_stock': val, 'updated_at': now})

    if not rows:
        return jsonify({'error': 'مفيش قيم صحيحة للحفظ'}), 400

    sb = get_client()
    try:
        execute_with_retry(sb.table('veg_inventory_entries').upsert(rows, on_conflict='entry_date,item_name'))
    except Exception as e:
        return jsonify({'error': f'تعذر الحفظ: {e}'}), 400
    return jsonify({'ok': True, 'updated_at': now, 'count': len(rows)})


@app.route('/api/veg-inventory/day/<entry_date>', methods=['DELETE'])
def veg_inventory_day_delete(entry_date):
    """حذف يوم كامل دفعة واحدة (كل الأصناف المسجلة للتاريخ ده) - محمي
    بتسجيل الدخول."""
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    try:
        execute_with_retry(sb.table('veg_inventory_entries').delete().eq('entry_date', entry_date))
    except Exception as e:
        return jsonify({'error': f'تعذر حذف اليوم: {e}'}), 400
    return jsonify({'ok': True})


# ============================================================
# أكواد البروموشن (Promo Codes) — لينك فريد لكل عميل، عداد تنازلي بيبدأ من
# أول لحظة فتح، استخدام مرة واحدة بس
# ============================================================
def _promo_row_to_public(row):
    """بتجهّز الصف عشان يترجع للعميل (من غير id داخلي زيادة عن اللازم)."""
    return {
        'code': row['code'],
        'title': row['title'],
        'discount_text': row['discount_text'],
        'duration_hours': row['duration_hours'],
        'first_opened_at': row['first_opened_at'],
        'used': row['used'],
    }


@app.route('/api/promo', methods=['POST'])
def promo_create():
    """إنشاء كود بروموشن جديد بلينك فريد - محمي بتسجيل الدخول."""
    _, err = _require_auth()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    code = (payload.get('code') or '').strip()
    title = (payload.get('title') or '').strip()
    discount_text = (payload.get('discount_text') or '').strip()
    try:
        duration_hours = int(payload.get('duration_hours') or 24)
    except (TypeError, ValueError):
        duration_hours = 24
    if not code or not title or not discount_text:
        return jsonify({'error': 'اكتب الكود والعنوان ووصف الخصم'}), 400
    if duration_hours < 1:
        return jsonify({'error': 'مدة الصلاحية لازم تكون ساعة على الأقل'}), 400

    sb = get_client()
    existing = execute_with_retry(sb.table('promo_codes').select('id').ilike('code', code))
    if existing.data:
        return jsonify({'error': f'الكود "{code}" مستخدم قبل كده — اختار كود تاني'}), 400

    token = secrets.token_urlsafe(16)
    row = {
        'token': token, 'code': code, 'title': title,
        'discount_text': discount_text, 'duration_hours': duration_hours,
    }
    try:
        res = execute_with_retry(sb.table('promo_codes').insert(row))
    except Exception as e:
        return jsonify({'error': f'تعذر الإنشاء: {e}'}), 400
    return jsonify({'ok': True, 'promo': (res.data or [{}])[0]})


@app.route('/api/promo', methods=['GET'])
def promo_list():
    """للداش بورد بتاعتك - محمي بتسجيل الدخول."""
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    res = execute_with_retry(sb.table('promo_codes').select('*').order('created_at', desc=True))
    return jsonify({'promos': res.data or []})


@app.route('/api/promo/<int:promo_id>/mark-used', methods=['PUT'])
def promo_mark_used(promo_id):
    """تعليم الكود كمستخدم - بيحصل لما خدمة العملاء تطبّق الخصم فعليًا."""
    username, err = _require_auth()
    if err:
        return err
    sb = get_client()
    try:
        execute_with_retry(
            sb.table('promo_codes').update({
                'used': True,
                'used_at': datetime.now(timezone.utc).isoformat(),
                'used_by_note': username,
            }).eq('id', promo_id)
        )
    except Exception as e:
        return jsonify({'error': f'تعذر التحديث: {e}'}), 400
    return jsonify({'ok': True})


@app.route('/api/promo/<int:promo_id>', methods=['DELETE'])
def promo_delete(promo_id):
    """إلغاء/حذف كود بروموشن - محمي بتسجيل الدخول."""
    _, err = _require_auth()
    if err:
        return err
    sb = get_client()
    try:
        execute_with_retry(sb.table('promo_codes').delete().eq('id', promo_id))
    except Exception as e:
        return jsonify({'error': f'تعذر الحذف: {e}'}), 400
    return jsonify({'ok': True})


@app.route('/api/promo/public/<token>', methods=['GET'])
def promo_public_get(token):
    """الصفحة اللي العميل بيفتحها - من غير لوجين. أول فتح بيسجّل توقيت
    بداية العداد التنازلي، وأي فتح بعد كده بيرجّع نفس التوقيت (العداد
    مايرجعش يبدأ من الأول)."""
    sb = get_client()
    res = execute_with_retry(sb.table('promo_codes').select('*').eq('token', token))
    rows = res.data or []
    if not rows:
        return jsonify({'error': 'الرابط ده مش صحيح'}), 404
    row = rows[0]

    if not row.get('first_opened_at'):
        now_iso = datetime.now(timezone.utc).isoformat()
        try:
            execute_with_retry(
                sb.table('promo_codes').update({'first_opened_at': now_iso}).eq('id', row['id'])
            )
            row['first_opened_at'] = now_iso
        except Exception:
            pass

    return jsonify({'promo': _promo_row_to_public(row)})


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
    بيرجّع openpyxl.Workbook فيه شيت Purchasing منسّق بشكل احترافي."""
    from openpyxl.styles import Border, Side, GradientFill
    THIN = Side(style='thin', color='D0D0D0')
    BOX  = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

    DARK_FILL  = PatternFill('solid', start_color='1A1A2E')   # هيدر داكن
    STAT_FILL  = PatternFill('solid', start_color='6600FF')   # محطات بنفسجي
    SUM_FILL   = PatternFill('solid', start_color='C04000')   # مجموع احمر
    EXTRA_FILL = PatternFill('solid', start_color='2E4057')   # أعمدة إضافية
    EVEN_FILL  = PatternFill('solid', start_color='F5F3FF')
    ODD_FILL   = PatternFill('solid', start_color='FFFFFF')

    WHITE_BOLD = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    WHITE_SM   = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    DATA_FONT  = Font(name='Calibri', size=10)
    NUM_FONT   = Font(name='Calibri', size=10, bold=True)
    CENTER     = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LEFT       = Alignment(horizontal='left',   vertical='center')

    all_names = set()
    for data in station_data.values():
        all_names.update(data.keys())
    sorted_names = sorted(all_names, key=lambda s: s.lower())

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Purchasing'
    ws.sheet_view.rightToLeft = False

    # ===== الصف الأول: عنوان رئيسي =====
    ws.merge_cells('A1:R1')
    ws['A1'] = 'Weekly Purchasing Report'
    ws['A1'].font = Font(name='Calibri', bold=True, color='FFFFFF', size=14)
    ws['A1'].fill = DARK_FILL
    ws['A1'].alignment = CENTER
    ws.row_dimensions[1].height = 30

    # ===== الصف الثاني: هيدر الأعمدة =====
    ws.row_dimensions[2].height = 32
    # A-C: بيانات الصنف
    for col, txt in [(1,'ITEMS'), (2,'Unit'), (3,'Category')]:
        c = ws.cell(row=2, column=col, value=txt)
        c.fill = DARK_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BOX

    # D-J: أعمدة المحطات
    for idx, key in enumerate(STATION_ORDER):
        col = 4 + idx
        c = ws.cell(row=2, column=col, value=STATION_LABELS[key])
        c.fill = STAT_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BOX

    sum_col  = 4 + len(STATION_ORDER)  # K
    dup_col  = sum_col + 1             # L
    maq_col  = dup_col + 1             # M
    exp_col  = maq_col + 1             # N
    avail_col= exp_col + 1             # O
    order_col= avail_col + 1           # P
    next_col = order_col + 1           # Q

    # K: مجموع
    c = ws.cell(row=2, column=sum_col, value='Weekly\nConsumption')
    c.fill = SUM_FILL; c.font = WHITE_BOLD; c.alignment = CENTER; c.border = BOX

    # L-Q: أعمدة المخزون
    extra_headers = {
        dup_col:  'Weekly\nConsumption',
        maq_col:  'Min. Available\nQty (MAQ)',
        exp_col:  'Expected\nStock',
        avail_col:'Available\nStock',
        order_col:'Weekly\nOrder',
        next_col: 'Next Week\nExpected Stock',
    }
    for col, txt in extra_headers.items():
        c = ws.cell(row=2, column=col, value=txt)
        c.fill = EXTRA_FILL; c.font = WHITE_SM; c.alignment = CENTER; c.border = BOX

    # ===== صفوف البيانات =====
    sum_letter   = get_column_letter(sum_col)
    dup_letter   = get_column_letter(dup_col)
    maq_letter   = get_column_letter(maq_col)
    avail_letter = get_column_letter(avail_col)
    order_letter = get_column_letter(order_col)
    d_letter     = get_column_letter(4)
    j_letter     = get_column_letter(4 + len(STATION_ORDER) - 1)

    for i, name in enumerate(sorted_names):
        r = 3 + i
        fill = EVEN_FILL if i % 2 == 0 else ODD_FILL

        unit, category = '', ''
        for key in STATION_ORDER:
            info = station_data.get(key, {}).get(name)
            if info:
                unit     = unit     or info.get('unit')     or ''
                category = category or info.get('category') or ''

        for col, val, fnt, aln in [
            (1, name,     DATA_FONT, LEFT),
            (2, unit,     DATA_FONT, CENTER),
            (3, category, DATA_FONT, CENTER),
        ]:
            cell = ws.cell(row=r, column=col, value=val)
            cell.fill = fill; cell.font = fnt; cell.alignment = aln; cell.border = BOX

        for idx, key in enumerate(STATION_ORDER):
            col = 4 + idx
            info = station_data.get(key, {}).get(name)
            val  = info['weekly'] if (info and info.get('weekly')) else None
            cell = ws.cell(row=r, column=col, value=val)
            cell.fill = fill; cell.font = NUM_FONT if val else DATA_FONT
            cell.alignment = CENTER; cell.border = BOX
            if val: cell.number_format = '#,##0.00'

        # K: SUM
        cell = ws.cell(row=r, column=sum_col, value=f'=SUM({d_letter}{r}:{j_letter}{r})')
        cell.fill = fill; cell.font = NUM_FONT; cell.alignment = CENTER
        cell.border = BOX; cell.number_format = '#,##0.00'

        # L: duplicate of sum
        cell = ws.cell(row=r, column=dup_col, value=f'={sum_letter}{r}')
        cell.fill = fill; cell.font = NUM_FONT; cell.alignment = CENTER
        cell.border = BOX; cell.number_format = '#,##0.00'

        # M, N, O: يدوي — فاضية
        for col in [maq_col, exp_col, avail_col]:
            cell = ws.cell(row=r, column=col)
            cell.fill = fill; cell.border = BOX

        # P: Weekly Order formula
        cell = ws.cell(row=r, column=order_col,
                       value=f'=({dup_letter}{r})-({avail_letter}{r}-{maq_letter}{r})')
        cell.fill = fill; cell.font = NUM_FONT; cell.alignment = CENTER
        cell.border = BOX; cell.number_format = '#,##0.00'

        # Q: Next week
        cell = ws.cell(row=r, column=next_col,
                       value=f'={order_letter}{r}+{avail_letter}{r}-{dup_letter}{r}')
        cell.fill = fill; cell.font = NUM_FONT; cell.alignment = CENTER
        cell.border = BOX; cell.number_format = '#,##0.00'

    # ===== عرض الأعمدة =====
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 16
    for col in range(4, 4 + len(STATION_ORDER)):
        ws.column_dimensions[get_column_letter(col)].width = 16
    for col, w in [(sum_col, 20), (dup_col, 20), (maq_col, 18),
                   (exp_col, 18), (avail_col, 16), (order_col, 16), (next_col, 22)]:
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = 'A3'

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


VEGETABLE_CATEGORY_LABELS = {'خضروات', 'خضراوات'}
FRUIT_CATEGORY_LABELS = {'فاكهة', 'فاكهه', 'فواكه'}
PRODUCE_CATEGORY_LABELS = VEGETABLE_CATEGORY_LABELS | FRUIT_CATEGORY_LABELS


def _read_vegetable_rows(file_storage, sheet_name):
    """بترجع صفوف الأصناف المصنّفة 'خضروات'/'خضراوات' بس من شيت المحطة،
    بنفس أعمدة A (الاسم) + B (الفئة) + D (الوزن اليومي) + L (طلب اليوم) +
    M (وحدة الطلب)، وبتشيل أي صف وزنه اليومي صفر بالظبط (زي باقي Daily Ordering)."""
    file_storage.seek(0)
    wb = openpyxl.load_workbook(file_storage, data_only=True)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    out = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=13, values_only=True):
        name, category = row[0], row[1]
        if not name or not str(name).strip():
            continue
        if str(name).strip().lower() == 'items':
            continue
        if not category or str(category).strip() not in VEGETABLE_CATEGORY_LABELS:
            continue
        daily_weight = row[3] if len(row) > 3 else None
        if isinstance(daily_weight, (int, float)) and not isinstance(daily_weight, bool) and daily_weight == 0:
            continue
        daily_order = row[11] if len(row) > 11 else None
        order_unit = row[12] if len(row) > 12 else None
        out.append({
            'name': str(name).strip(), 'category': str(category).strip(),
            'daily_weight': daily_weight, 'daily_order': daily_order, 'order_unit': order_unit,
        })
    return out


def _build_vegetables_workbook(station_vegetable_data):
    """station_vegetable_data: {station_key: [rows]}
    - تاب لكل محطة فيها خضروات فعلاً (المحطات الفاضية بتتشال تلقائي)
    - تاب أخير 'All Vegetables' مجمّع
    - ستايل: هيدر بنفسجي، ألوان متبادلة على الصفوف، أعمدة واسعة، RTL"""
    HEADER_FILL = PatternFill('solid', start_color='6600FF')
    HEADER_FONT = Font(name='Tahoma', bold=True, color='FFFFFF', size=11)
    HEADER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

    EVEN_FILL = PatternFill('solid', start_color='F2EEFF')
    ODD_FILL  = PatternFill('solid', start_color='FFFFFF')
    DATA_FONT = Font(name='Tahoma', size=11)
    NUM_FONT  = Font(name='Tahoma', size=11, bold=True)
    CENTER    = Alignment(horizontal='center', vertical='center')
    RIGHT     = Alignment(horizontal='right',  vertical='center')

    from openpyxl.styles import Border, Side
    THIN = Side(style='thin', color='D0C8F0')
    BOX  = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

    COL_WIDTHS_NORMAL = [48, 14, 16, 14]  # A:D (بدون عمود المحطة ولا Daily Weight)
    COL_WIDTHS_ALL    = [18, 48, 14, 16, 14]  # A:E (مع عمود المحطة، بدون Daily Weight)

    HEADERS_NORMAL = ['ITEMS', 'Category', 'Daily Order', 'Order Unit']
    HEADERS_ALL    = ['Station', 'ITEMS', 'Category', 'Daily Order', 'Order Unit']

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def _write_sheet(title, rows, with_station_col=False):
        if not rows and not with_station_col:
            return  # لا تعمل تاب للمحطات الفاضية
        ws = wb.create_sheet(title=title[:31])
        ws.sheet_view.rightToLeft = False  # LTR
        ws.row_dimensions[1].height = 24

        headers = HEADERS_ALL if with_station_col else HEADERS_NORMAL
        widths  = COL_WIDTHS_ALL if with_station_col else COL_WIDTHS_NORMAL

        for c, (h, w) in enumerate(zip(headers, widths), start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill      = HEADER_FILL
            cell.font      = HEADER_FONT
            cell.alignment = HEADER_ALIGN
            cell.border    = BOX
            ws.column_dimensions[get_column_letter(c)].width = w

        for i, row in enumerate(rows):
            r = i + 2
            fill = EVEN_FILL if i % 2 == 0 else ODD_FILL
            c = 1
            if with_station_col:
                cell = ws.cell(row=r, column=c, value=row.get('_station_label', ''))
                cell.fill = fill; cell.font = DATA_FONT; cell.border = BOX
                cell.alignment = CENTER
                c += 1
            # ITEMS
            cell = ws.cell(row=r, column=c, value=row['name'])
            cell.fill = fill; cell.font = DATA_FONT; cell.border = BOX
            cell.alignment = RIGHT; c += 1
            # Category
            cell = ws.cell(row=r, column=c, value=row['category'])
            cell.fill = fill; cell.font = DATA_FONT; cell.border = BOX
            cell.alignment = CENTER; c += 1
            # Daily Order
            cell = ws.cell(row=r, column=c, value=row['daily_order'])
            cell.fill = fill; cell.font = NUM_FONT; cell.border = BOX
            cell.alignment = CENTER
            cell.number_format = '#,##0.000'; c += 1
            # Order Unit
            cell = ws.cell(row=r, column=c, value=row['order_unit'])
            cell.fill = fill; cell.font = DATA_FONT; cell.border = BOX
            cell.alignment = CENTER

        ws.freeze_panes = 'A2'

    # تابات المحطات (بس اللي فيها خضروات)
    all_rows = []
    for key in STATION_ORDER:
        rows = station_vegetable_data.get(key, [])
        if rows:  # تخطي المحطات الفاضية
            _write_sheet(STATION_TAB_NAMES[key], rows)
            for row in rows:
                all_rows.append({**row, '_station_label': STATION_LABELS.get(key, key)})

    # تاب All Vegetables (كل الصفوف raw)
    _write_sheet('All Vegetables', all_rows, with_station_col=True)

    # تاب Summary — كل صنف مرة واحدة، الوزن اليومي مجمّع من كل المحطات
    from collections import defaultdict
    summary = {}  # name -> {category, daily_weight_total, daily_order_total, order_unit}
    for row in all_rows:
        name = row['name']
        if name not in summary:
            summary[name] = {
                'name': name,
                'category': row['category'],
                'daily_weight': row['daily_weight'] or 0,
                'daily_order': row['daily_order'] or 0,
                'order_unit': row['order_unit'],
            }
        else:
            summary[name]['daily_weight'] = (summary[name]['daily_weight'] or 0) + (row['daily_weight'] or 0)
            summary[name]['daily_order']  = (summary[name]['daily_order']  or 0) + (row['daily_order']  or 0)

    summary_rows = sorted(summary.values(), key=lambda r: r['name'].lower())
    _write_sheet('Summary', summary_rows)

    if not wb.sheetnames:
        wb.create_sheet('فاضي')
    return wb


def _add_station_tab_daily(wb, station_key, file_storage):
    """زي _add_station_tab بالظبط، بس بترجع أعمدة A:D بس (من غير الوزن
    الأسبوعي في E)، وبتشيل: (1) أي صف يكون الوزن اليومي بتاعه (عمود D) صفر
    رقمي بالظبط، و(2) أي صف فئته 'خضروات/خضراوات/فاكهة' لأنها بتنزل في
    Vegetables.xlsx لوحدها فمفيش داعي تتكرر هنا. صفوف العناوين والفئات
    (اللي عمود D فيها فاضي) بتفضل زي ما هي."""
    file_storage.seek(0)
    src_wb = openpyxl.load_workbook(file_storage, data_only=True)
    sheet_name = STATION_SHEET_MAP[station_key]
    if sheet_name not in src_wb.sheetnames:
        return None
    src_ws = src_wb[sheet_name]
    out_ws = wb.create_sheet(title=STATION_TAB_NAMES[station_key])

    COLS = 4
    out_row = 1
    for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row, min_col=1, max_col=COLS):
        d_value = row[3].value if len(row) > 3 else None
        if isinstance(d_value, (int, float)) and not isinstance(d_value, bool) and d_value == 0:
            continue  # الصف ده وزنه اليومي صفر بالظبط — نتخطاه بالكامل
        category = row[1].value if len(row) > 1 else None
        if category and str(category).strip() in PRODUCE_CATEGORY_LABELS:
            continue  # خضروات/فاكهة - موجودة في Vegetables.xlsx لوحدها
        for cell in row:
            new_cell = out_ws.cell(row=out_row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format
        out_row += 1

    for col_letter in ['A', 'B', 'C', 'D']:
        if col_letter in src_ws.column_dimensions:
            out_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
    return out_ws


def _build_daily_ordering_zip(wb_daily, wb_veg, today, with_images=True, day_num_override=None):
    """بتبني zip فيه Daily_Ordering + Vegetables (إكسيل) + صورة PNG لكل تاب
    فيهم لو with_images=True (لو توليد الصور فشل لأي سبب - مثلاً LibreOffice
    مش متظبط على السيرفر - بيرجع الإكسيل عادي بدون ما يكسر الطلب كله).
    day_num_override: رقم اليوم (١=السبت...٧=الجمعة) واحد للكل، أو dict
    {tab_name: day_num} عشان كل تاب ياخد رقم اليوم بتاع ملف المحطة بتاعه هو
    (متقري من خلية R1 في كل ملف لوحده)، بدل ما يتحسب من تاريخ السيرفر."""
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        buf1 = io.BytesIO(); wb_daily.save(buf1)
        zf.writestr(f'Daily_Ordering_{today}.xlsx', buf1.getvalue())
        buf2 = io.BytesIO(); wb_veg.save(buf2)
        zf.writestr(f'Vegetables_{today}.xlsx', buf2.getvalue())

        if with_images:
            try:
                add_workbook_images_to_zip(zf, wb_daily, today, prefix='DailyOrdering_',
                                            day_num_override=day_num_override)
                add_workbook_images_to_zip(zf, wb_veg, today, prefix='Vegetables_',
                                            day_num_override=day_num_override)
            except Exception as e:
                app.logger.exception('تعذر توليد صور التابات (الإكسيل نزل عادي بدونها)')
                zf.writestr('images/تعذر_توليد_الصور.txt',
                             f'حصل خطأ أثناء توليد الصور: {e}')
    zip_buf.seek(0)
    return zip_buf


def _build_single_workbook_zip(wb, today, file_label, image_prefix, day_num_override=None):
    """زي _build_daily_ordering_zip بالظبط بس لملف واحد بس (مش اتنين) — مستخدمة
    في زرار "Daily Ordering" أو "Vegetables" لوحدهم، عشان صور التابات PNG
    تفضل متضافة زي ما كانت أول ما الزرارين كانوا مدموجين في واحد."""
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        buf = io.BytesIO(); wb.save(buf)
        zf.writestr(f'{file_label}_{today}.xlsx', buf.getvalue())
        try:
            add_workbook_images_to_zip(zf, wb, today, prefix=image_prefix,
                                        day_num_override=day_num_override)
        except Exception as e:
            app.logger.exception('تعذر توليد صور التابات (الإكسيل نزل عادي بدونها)')
            zf.writestr('images/تعذر_توليد_الصور.txt',
                         f'حصل خطأ أثناء توليد الصور: {e}')
    zip_buf.seek(0)
    return zip_buf


def _read_report_day_numbers_per_station(files_by_key):
    """بترجع dict {station_key: day_num} — كل ملف من ملفات المحطات بيتقرا لوحده
    (مش بتوقف عند أول ملف لاقيه)، ورقم اليوم بتاعه (١=السبت...٧=الجمعة) بيتاخد
    من خلية R1. لكل محطة، بندوّر الأول على شيتها بالاسم المعروف من
    STATION_SHEET_MAP (مهم جدًا لو نفس الملف بالظبط مستخدم لمحطتين مختلفين
    في نفس الوقت - زي ملف توكيو اللي بيتحط لـ hot و marination مع بعض،
    عشان hot تاخد R1 بتاعة All_Ingredients ومارينيشن تاخد R1 بتاعة
    Marination_Ordering، مش نفس الرقم للاتنين). لو الشيت بالاسم ده مش
    موجود أو R1 بتاعه مش رقم صالح، بترجع تدوّر في كل تابات نفس الملف
    كـ fallback. الملفات اللي مفيش في أي تاب فيها رقم صالح بتتسيب برة الـ dict."""
    result = {}
    for key in STATION_ORDER:
        f = files_by_key.get(key)
        if not f:
            continue
        try:
            f.seek(0)
            wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
            found = None

            # 1) جرّب الأول شيت المحطة المعروف بالاسم (يفرّق بين hot وmarination
            #    حتى لو الاتنين بيتقروا من نفس الملف بالظبط)
            named_sheet = STATION_SHEET_MAP.get(key)
            if named_sheet and named_sheet in wb.sheetnames:
                raw = wb[named_sheet]['R1'].value
                try:
                    n = int(str(raw).strip())
                    if 1 <= n <= 7:
                        found = n
                except (TypeError, ValueError):
                    pass

            # 2) لو مالقتش حاجة بالاسم المعروف، دوّر في كل تابات الملف
            if found is None:
                for ws in wb.worksheets:
                    try:
                        raw = ws['R1'].value
                    except Exception:
                        continue
                    try:
                        n = int(str(raw).strip())
                    except (TypeError, ValueError):
                        continue
                    if 1 <= n <= 7:
                        found = n
                        break

            wb.close()
            f.seek(0)
            if found is not None:
                result[key] = found
        except Exception:
            f.seek(0)
            continue
    if not result:
        app.logger.warning('تعذّر قراءة رقم اليوم من R1 في أي ملف من ملفات المحطات - هيتحسب من تاريخ السيرفر بدل منه')
    return result


def _day_numbers_by_tab(day_numbers_by_station):
    """بتحوّل {station_key: day_num} لـ {tab_name: day_num} عشان xlsx_to_images
    تقدر تطابق كل تاب في الإكسيل الناتج (Daily_Ordering أو Vegetables) برقم
    اليوم بتاع ملف المحطة اللي طلع منها التاب ده بالظبط."""
    return {
        STATION_TAB_NAMES[key]: day_num
        for key, day_num in day_numbers_by_station.items()
        if key in STATION_TAB_NAMES
    }


@app.route('/api/daily-ordering', methods=['POST'])
def daily_ordering():
    """بتاخد نفس ملفات الـ7 محطات بتاعة Weekly Purchasing، وبترجع zip فيه
    ملفين: Daily_Ordering.xlsx (تاب لكل محطة بأعمدة A:D، أي صف وزنه اليومي
    صفر بيتشال)، و Vegetables.xlsx (تاب لكل محطة فيه أصناف 'خضروات' بس +
    تاب أخير 'All Vegetables' مجمّع فيه كل الخضروات من كل المحطات)."""
    missing = [k for k in STATION_ORDER if k not in request.files]
    if missing:
        return jsonify({'error': f'محطات ناقصة: {", ".join(missing)}'}), 400

    try:
        wb_daily = openpyxl.Workbook()
        wb_daily.remove(wb_daily.active)
        vegetable_data = {}

        # قراءة بيانات كل المحطات
        all_daily_rows = {}  # name -> {unit, category, daily_weight, daily_order, order_unit}
        for key in STATION_ORDER:
            request.files[key].seek(0)
            _add_station_tab_daily(wb_daily, key, request.files[key])
            request.files[key].seek(0)
            vegetable_data[key] = _read_vegetable_rows(request.files[key], STATION_SHEET_MAP[key])
            # جمع كل الأصناف من الـ Ordering sheet لعمل Summary
            request.files[key].seek(0)
            src_wb = openpyxl.load_workbook(request.files[key], data_only=True)
            sheet_name = STATION_SHEET_MAP[key]
            if sheet_name in src_wb.sheetnames:
                src_ws = src_wb[sheet_name]
                for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row, min_col=1, max_col=13, values_only=True):
                    name = row[0]
                    if not name or str(name).strip().lower() in ('', 'items'):
                        continue
                    daily_w = row[3] if len(row) > 3 else None
                    if not isinstance(daily_w, (int, float)) or daily_w == 0:
                        continue
                    cat  = row[1] if len(row) > 1 else None
                    if cat and str(cat).strip() in PRODUCE_CATEGORY_LABELS:
                        continue  # خضروات/فاكهة - موجودة في Vegetables.xlsx لوحدها
                    n = str(name).strip()
                    unit = row[2] if len(row) > 2 else None
                    d_order = row[11] if len(row) > 11 else None
                    o_unit  = row[12] if len(row) > 12 else None
                    if n not in all_daily_rows:
                        all_daily_rows[n] = {'name': n, 'category': cat or '', 'unit': unit or '',
                                              'daily_weight': daily_w, 'daily_order': d_order or 0,
                                              'order_unit': o_unit or ''}
                    else:
                        all_daily_rows[n]['daily_weight'] = (all_daily_rows[n]['daily_weight'] or 0) + daily_w
                        all_daily_rows[n]['daily_order']  = (all_daily_rows[n]['daily_order']  or 0) + (d_order or 0)

        # تاب Summary في Daily_Ordering.xlsx
        summary_rows = sorted(all_daily_rows.values(), key=lambda r: r['name'].lower())
        ws_sum = wb_daily.create_sheet(title='Summary')
        ws_sum.sheet_view.rightToLeft = False  # LTR
        ws_sum.row_dimensions[1].height = 24
        HEADER_FILL2 = PatternFill('solid', start_color='6600FF')
        HEADER_FONT2 = Font(name='Tahoma', bold=True, color='FFFFFF', size=11)
        from openpyxl.styles import Border, Side
        THIN2 = Side(style='thin', color='D0C8F0')
        BOX2  = Border(top=THIN2, bottom=THIN2, left=THIN2, right=THIN2)
        EVEN2 = PatternFill('solid', start_color='F2EEFF')

        headers_s = ['ITEMS', 'Category', 'Unit', 'Daily Weight', 'Daily Order', 'Order Unit']
        widths_s   = [48, 16, 8, 16, 14, 14]
        for c, (h, w) in enumerate(zip(headers_s, widths_s), start=1):
            cell = ws_sum.cell(row=1, column=c, value=h)
            cell.fill = HEADER_FILL2; cell.font = HEADER_FONT2
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = BOX2
            ws_sum.column_dimensions[get_column_letter(c)].width = w

        for i, row in enumerate(summary_rows):
            r = i + 2
            fill = EVEN2 if i % 2 == 0 else PatternFill('solid', start_color='FFFFFF')
            vals = [row['name'], row['category'], row['unit'],
                    row['daily_weight'], row['daily_order'], row['order_unit']]
            aligns = ['right', 'center', 'center', 'center', 'center', 'center']
            for c, (v, al) in enumerate(zip(vals, aligns), start=1):
                cell = ws_sum.cell(row=r, column=c, value=v)
                cell.fill = fill
                cell.font = Font(name='Tahoma', size=11, bold=(c in (4, 5)))
                cell.alignment = Alignment(horizontal=al, vertical='center')
                cell.border = BOX2
                if c in (4, 5) and isinstance(v, float):
                    cell.number_format = '#,##0.00'
        ws_sum.freeze_panes = 'A2'

        if not wb_daily.sheetnames:
            wb_daily.create_sheet('فاضي')

        wb_veg = _build_vegetables_workbook(vegetable_data)

        today = datetime.now().strftime('%Y-%m-%d')
        day_numbers_by_station = _read_report_day_numbers_per_station({k: request.files[k] for k in STATION_ORDER})
        day_num_by_tab = _day_numbers_by_tab(day_numbers_by_station)
        zip_buf = _build_daily_ordering_zip(wb_daily, wb_veg, today, day_num_override=day_num_by_tab)
        return send_file(zip_buf, as_attachment=True,
                          download_name=f'Daily_Ordering_{today}.zip',
                          mimetype='application/zip')
    except Exception as e:
        app.logger.exception('daily_ordering failed')
        return jsonify({'error': f'حصل خطأ في التجميع: {e}'}), 500


@app.route('/api/whatsapp-send', methods=['POST'])
def whatsapp_send():
    """بيستقبل صورة من المتصفح ويمررها لسيرفر الواتساب المنفصل (Node.js) عشان
    يبعتها تلقائي للرقم المتظبط. لو الـenv vars مش متظبطة، بيرجع خطأ واضح."""
    bot_url = os.environ.get('WHATSAPP_BOT_URL')
    api_key = os.environ.get('WHATSAPP_BOT_API_KEY')
    if not bot_url:
        return jsonify({'error': 'سيرفر الواتساب لسه مش متظبط (WHATSAPP_BOT_URL ناقصة)'}), 503

    if 'image' not in request.files:
        return jsonify({'error': 'مفيش صورة مبعوتة'}), 400

    try:
        files = {'image': (request.files['image'].filename or 'card.png',
                            request.files['image'].stream, 'image/png')}
        data = {}
        if request.form.get('number'):
            data['number'] = request.form['number']
        if request.form.get('caption'):
            data['caption'] = request.form['caption']
        headers = {'x-api-key': api_key} if api_key else {}

        resp = requests.post(f'{bot_url}/send-image', files=files, data=data, headers=headers, timeout=30)
        if resp.status_code != 200:
            return jsonify({'error': f'فشل سيرفر الواتساب: {resp.text[:200]}'}), 502
        return jsonify(resp.json())
    except Exception as e:
        app.logger.exception('whatsapp_send failed')
        return jsonify({'error': f'حصل خطأ في الاتصال بسيرفر الواتساب: {e}'}), 500


def _build_summary_sheet(ws, rows, with_unit_col=False):
    """بيكتب شيت Summary منسّق (هيدر بنفسجي، ألوان متبادلة، اتجاه شمال لأيمن LTR)."""
    from openpyxl.styles import Border, Side
    THIN = Side(style='thin', color='D0C8F0')
    BOX  = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    H_FILL = PatternFill('solid', start_color='6600FF')
    H_FONT = Font(name='Tahoma', bold=True, color='FFFFFF', size=11)
    EVEN   = PatternFill('solid', start_color='F2EEFF')
    ODD    = PatternFill('solid', start_color='FFFFFF')
    D_FONT = Font(name='Tahoma', size=11)
    N_FONT = Font(name='Tahoma', size=11, bold=True)
    CENTER = Alignment(horizontal='center', vertical='center')
    LEFT   = Alignment(horizontal='left', vertical='center')

    ws.sheet_view.rightToLeft = False  # LTR — شمال لأيمن
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = 'A2'

    if with_unit_col:
        headers = ['ITEMS', 'Category', 'Unit', 'Daily Weight']
        widths  = [48, 16, 8, 16]
    else:
        headers = ['ITEMS', 'Category', 'Daily Weight']
        widths  = [48, 16, 16]

    for c, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = H_FILL; cell.font = H_FONT
        cell.alignment = CENTER; cell.border = BOX
        ws.column_dimensions[get_column_letter(c)].width = w

    for i, row in enumerate(rows):
        r = i + 2
        fill = EVEN if i % 2 == 0 else ODD
        vals = ([row['name'], row.get('category', ''), row.get('unit', ''), row['daily_weight']]
                if with_unit_col else
                [row['name'], row.get('category', ''), row['daily_weight']])
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = fill; cell.border = BOX
            is_num = isinstance(v, (int, float)) and not isinstance(v, bool)
            cell.font = N_FONT if is_num else D_FONT
            cell.alignment = CENTER if (c > 1) else LEFT
            if is_num:
                cell.number_format = '#,##0.00'


def _detect_station_from_workbook(wb):
    """بتحدد نوع المحطة من الشيتات الموجودة في الملف — بدون الاعتماد على اسم الملف.
    الأولوية بالترتيب عشان الفحص يكون دقيق ومحدد."""
    sheets = set(wb.sheetnames)
    if 'All_Ingredients' in sheets and 'Marination_Ordering' in sheets:
        return 'tokyo'  # ملف توكيو الرئيسي (فيه الاتنين مع بعض)
    if 'Marination_Ordering' in sheets:
        return 'marination'
    if 'All_Ingredients' in sheets:
        return 'hot'
    if 'User' in sheets and 'Usage' in sheets:
        return 'salads'  # ملف السلطات عنده شيت User + Usage مميزين
    # الملفات اللي عندها شيت Ordering + شيتات وجبات عربية
    if 'Ordering' in sheets:
        ar_count = sum(1 for s in sheets if any('\u0600' <= c <= '\u06FF' for c in s))
        if ar_count >= 3:
            return 'rice'  # شيت الأرز فيه أسماء شيتات عربية كتير
        if 'List of Meals' in sheets:
            return 'sauce'
        # فطار أو حلويات — نفرق بينهم من اسم أول شيت بعد Ordering
        others = [s for s in wb.sheetnames if s != 'Ordering']
        if others:
            first = others[0].lower()
            if any(w in first for w in ('foul', 'egg', 'croissant', 'sandwich', 'omelette', 'fool')):
                return 'breakfast'
            if any(w in first for w in ('pie', 'cake', 'cookie', 'brownie', 'dessert', 'zatar')):
                return 'desserts'
    return None


@app.route('/api/auto-detect-stations', methods=['POST'])
def auto_detect_stations():
    """بتاخد ملفات متعددة مرة واحدة (multipart 'files')، بتحدد محطة كل ملف
    تلقائياً من محتواه (مش اسمه)، وبترجع نفس zip بتاع daily-ordering بس من
    ملف واحد بس بدل 7 ملفات منفصلين.
    ملف توكيو الرئيسي (فيه All_Ingredients + Marination_Ordering) بيتعامل معاه
    تلقائي على إنه hot + marination في نفس الوقت."""
    uploaded = request.files.getlist('files')
    if not uploaded:
        return jsonify({'error': 'مفيش ملفات مبعوتة'}), 400

    # خطوة 1: اكتشف محطة كل ملف
    station_files = {}
    undetected = []
    for f in uploaded:
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            kind = _detect_station_from_workbook(wb)
            wb.close()
            f.seek(0)
            if kind == 'tokyo':
                station_files['hot'] = f
                station_files['marination'] = f
            elif kind:
                station_files[kind] = f
            else:
                undetected.append(f.filename)
        except Exception as e:
            undetected.append(f'{f.filename} (خطأ: {e})')

    if not station_files:
        return jsonify({
            'error': f'مش قادر أحدد محطة أي ملف من اللي رفعتهم: {undetected}'
        }), 400
    detected_keys = list(station_files.keys())

    # خطوة 2: نفس منطق daily_ordering بالضبط
    try:
        wb_daily = openpyxl.Workbook()
        wb_daily.remove(wb_daily.active)
        vegetable_data = {}
        all_daily_rows = {}

        for key in detected_keys:
            f = station_files[key]
            f.seek(0)
            _add_station_tab_daily(wb_daily, key, f)
            f.seek(0)
            vegetable_data[key] = _read_vegetable_rows(f, STATION_SHEET_MAP[key])
            f.seek(0)
            src_wb = openpyxl.load_workbook(f, data_only=True)
            sheet_name = STATION_SHEET_MAP[key]
            if sheet_name in src_wb.sheetnames:
                src_ws = src_wb[sheet_name]
                for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row,
                                             min_col=1, max_col=13, values_only=True):
                    name = row[0]
                    if not name or str(name).strip().lower() in ('', 'items'):
                        continue
                    daily_w = row[3] if len(row) > 3 else None
                    if not isinstance(daily_w, (int, float)) or daily_w == 0:
                        continue
                    category0 = row[1] if len(row) > 1 else None
                    if category0 and str(category0).strip() in PRODUCE_CATEGORY_LABELS:
                        continue  # خضروات/فاكهة - موجودة في Vegetables.xlsx لوحدها
                    n = str(name).strip()
                    d_order = row[11] if len(row) > 11 else None
                    o_unit  = row[12] if len(row) > 12 else None
                    if n not in all_daily_rows:
                        all_daily_rows[n] = {
                            'name': n, 'category': row[1] or '', 'unit': row[2] or '',
                            'daily_weight': daily_w, 'daily_order': d_order or 0,
                            'order_unit': o_unit or '',
                        }
                    else:
                        all_daily_rows[n]['daily_weight'] += daily_w
                        all_daily_rows[n]['daily_order']  += (d_order or 0)

        # Summary tab في Daily_Ordering
        summary_rows = sorted(all_daily_rows.values(), key=lambda r: r['name'].lower())
        ws_sum = wb_daily.create_sheet(title='Summary')
        _build_summary_sheet(ws_sum, summary_rows, with_unit_col=True)

        wb_veg = _build_vegetables_workbook(vegetable_data)

        today = datetime.now().strftime('%Y-%m-%d')
        day_numbers_by_station = _read_report_day_numbers_per_station(station_files)
        day_num_override = _day_numbers_by_tab(day_numbers_by_station)

        # ?only=daily أو ?only=vegetables — بيرجّع zip فيه ملف واحد بس + صوره،
        # عشان الواجهة تقدر تفصل زرار "Daily Ordering" عن زرار "Vegetables" لوحدهم
        # (لسه بيرجع zip مش xlsx خام، عشان صور التابات متضاعش زي الأول).
        only = request.args.get('only')
        if only == 'daily':
            zip_buf = _build_single_workbook_zip(wb_daily, today, 'Daily_Ordering', 'DailyOrdering_', day_num_override)
            return send_file(zip_buf, as_attachment=True,
                              download_name=f'Daily_Ordering_{today}.zip',
                              mimetype='application/zip')
        if only == 'vegetables':
            zip_buf = _build_single_workbook_zip(wb_veg, today, 'Vegetables', 'Vegetables_', day_num_override)
            return send_file(zip_buf, as_attachment=True,
                              download_name=f'Vegetables_{today}.zip',
                              mimetype='application/zip')

        zip_buf = _build_daily_ordering_zip(wb_daily, wb_veg, today, day_num_override=day_num_override)
        return send_file(zip_buf, as_attachment=True,
                          download_name=f'Daily_Ordering_{today}.zip',
                          mimetype='application/zip')
    except Exception as e:
        app.logger.exception('auto_detect_stations failed')
        return jsonify({'error': f'حصل خطأ: {e}'}), 500


@app.route('/api/auto-weekly-purchasing', methods=['POST'])
def auto_weekly_purchasing():
    """نفس فكرة auto-detect-stations بس بيطلع Weekly Purchasing (نسختين كاملة + مطبخ)."""
    uploaded = request.files.getlist('files')
    if not uploaded:
        return jsonify({'error': 'مفيش ملفات مبعوتة'}), 400

    station_files = {}
    for f in uploaded:
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            kind = _detect_station_from_workbook(wb)
            wb.close(); f.seek(0)
            if kind == 'tokyo':
                station_files['hot'] = f
                station_files['marination'] = f
            elif kind:
                station_files[kind] = f
        except Exception:
            pass

    if not station_files:
        return jsonify({'error': 'مش قادر أحدد محطة أي ملف من اللي رفعتهم'}), 400
    detected_keys = list(station_files.keys())

    try:
        station_data = {}
        for key in detected_keys:
            station_files[key].seek(0)
            _, rows = _read_station_rows(station_files[key], STATION_SHEET_MAP[key])
            station_data[key] = rows

        wb_full, cols = _build_purchasing_workbook(station_data)
        for key in detected_keys:
            station_files[key].seek(0)
            _add_station_tab(wb_full, key, station_files[key])

        wb_kitchen, _ = _build_purchasing_workbook(station_data)
        for key in detected_keys:
            station_files[key].seek(0)
            ws_station = _add_station_tab(wb_kitchen, key, station_files[key])
            if ws_station:
                ws_station.sheet_state = 'hidden'
        kitchen_ws = wb_kitchen['Purchasing']
        for col in range(4, cols['sum_col'] + 1):
            kitchen_ws.column_dimensions[get_column_letter(col)].hidden = True

        today = datetime.now().strftime('%Y-%m-%d')
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            buf1 = io.BytesIO(); wb_full.save(buf1)
            zf.writestr(f'Weekly_Purchasing_Full_{today}.xlsx', buf1.getvalue())
            buf2 = io.BytesIO(); wb_kitchen.save(buf2)
            zf.writestr(f'Weekly_Purchasing_Kitchen_{today}.xlsx', buf2.getvalue())

        zip_buf.seek(0)
        return send_file(zip_buf, as_attachment=True,
                          download_name=f'Weekly_Purchasing_{today}.zip',
                          mimetype='application/zip')
    except Exception as e:
        app.logger.exception('auto_weekly_purchasing failed')
        return jsonify({'error': f'حصل خطأ: {e}'}), 500



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
