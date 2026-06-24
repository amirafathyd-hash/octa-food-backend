import os
import io
import tempfile
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from parse_order import parse_order_pdf
from parse_invoice import parse_invoice_pdf
from parse_received import parse_received_xlsx
from parse_received_image import parse_received_image, process_ocr_data
from item_db import load_db, seed_from_order
from matcher import match_invoice_item
from db import get_client, execute_with_retry

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
        parsed = process_ocr_data(ocr_data, db)
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


def _next_month(month):
    y, m = map(int, month.split('-'))
    return f'{y+1}-01-01' if m == 12 else f'{y}-{m+1:02d}-01'


if __name__ == '__main__':
    app.run(debug=True, port=8000)
