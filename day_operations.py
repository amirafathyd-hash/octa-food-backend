import io
import json
import os
import re
import tempfile
import uuid
import zipfile
from collections import OrderedDict, defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from decision_station import (
    build_output_workbook,
    compute_decision_tables,
    load_lookup,
    read_subscribers_invoice,
)
from tokyo_ordering import merge_day_into_template, read_day_file_payload


TOKYO_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'tokyo_ordering_template.xlsm')
ARCHIVE_DIR = os.path.join(os.path.dirname(__file__), 'data', 'day_operations_archive')


PACKAGE_ORDER = ['تضخيم', 'تكميم لايت', 'جيم', 'سمارت دايت', 'غذاء العمل']

WORKER_LINKS = [
    ('مركز تشغيل اليوم', 'day-operations.html', 'رفع واحد وتشغيل مخرجات اليوم'),
    ('محطة القرار', 'decision-station.html', 'ملف Update و Packages و Dont Use'),
    ('محطات التجهيز', 'ordering-stations.html', 'الأرز والفطار والسلطات والحلى والصوص'),
    ('Tokyo Production Master', 'tokyo-production-dashboard.html', 'إدارة إنتاج توكيو وتقارير المحطات'),
    ('روابط العاملين', 'worker-links-dashboard.html', 'إنشاء وتحديث روابط المهام للعاملين'),
    ('العمليات اللايف', 'live-operations.html', 'متابعة التشغيل الحي'),
    ('سجل الموازين', 'weight-log-dashboard.html', 'أداة ثابتة منفصلة عن تشغيل اليوم'),
]


def _static_worker_links():
    return [
        {
            'title': title,
            'url': href,
            'description': desc,
            'worker_name': 'تشغيل النظام',
            'username': '',
            'source': 'system',
        }
        for title, href, desc in WORKER_LINKS
    ]


def _load_worker_links():
    links = _static_worker_links()
    try:
        from db import execute_with_retry, get_client
        sb = get_client()
        res = execute_with_retry(
            sb.table('worker_link_assignments')
            .select('worker_name, username, task_title, target_url, active, created_at')
            .eq('active', True)
            .order('created_at', desc=True)
        )
        for row in res.data or []:
            links.append({
                'title': row.get('task_title') or 'مهمة عامل',
                'url': row.get('target_url') or '',
                'description': 'رابط عامل نشط',
                'worker_name': row.get('worker_name') or '',
                'username': row.get('username') or '',
                'source': 'worker',
            })
    except Exception:
        pass
    return links


def _num(value):
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _clean_sheet_title(value, fallback='Sheet'):
    safe = ''.join(ch for ch in str(value or fallback) if ch not in r'[]:*?/\\').strip()
    return (safe or fallback)[:31]


def _station_key(category):
    text = category or ''
    if 'فطور' in text:
        return 'breakfast'
    if 'حلى' in text or 'حلويات' in text:
        return 'dessert'
    if 'سلطات' in text:
        return 'salads'
    if 'فواكه' in text or 'الإضافات' in text:
        return 'addons'
    if 'الوجبات الرئيسية' in text or 'لو كارب' in text:
        return 'main_kitchen'
    return 'unclassified'


def _station_title(key):
    return {
        'breakfast': 'محطة الفطار',
        'main_kitchen': 'المطبخ الرئيسي',
        'salads': 'محطة السلطات',
        'dessert': 'محطة الحلى',
        'addons': 'الفواكه والإضافات',
        'unclassified': 'غير مصنف',
    }.get(key, key)


def _style_sheet(ws, title=None):
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A2'
    thin = Side(style='thin', color='E5D6C8')
    border = Border(top=thin, right=thin, bottom=thin, left=thin)
    header_fill = PatternFill('solid', fgColor='7A2118')
    header_font = Font(name='Arial', bold=True, color='FFFFFF', size=12)
    data_font = Font(name='Arial', size=11, color='211713')
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
            cell.font = data_font
    if ws.max_row >= 1:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 12
        for cell in ws[letter]:
            max_len = max(max_len, len(str(cell.value or '')) + 2)
        ws.column_dimensions[letter].width = min(max_len, 38)
    if title:
        ws.title = _clean_sheet_title(title)


def _summaries(dont_use_rows):
    package_totals = OrderedDict()
    protein_totals = OrderedDict()
    station_totals = OrderedDict()
    for row in dont_use_rows:
        pkg = row.get('Final_Package') or 'غير محدد'
        protein = row.get('Protein') or row.get('الاسم الإنجليزي') or 'غير محدد'
        category = row.get('التصنيف') or ''
        station = _station_key(category)
        count = _num(row.get('Final_Count'))
        grams = _num(row.get('Grams'))

        package_totals.setdefault(pkg, {'count': 0.0, 'grams': 0.0, 'rows': 0})
        package_totals[pkg]['count'] += count
        package_totals[pkg]['grams'] += grams
        package_totals[pkg]['rows'] += 1

        protein_totals.setdefault(protein, {'category': category, 'count': 0.0, 'grams': 0.0, 'packages': defaultdict(float)})
        protein_totals[protein]['count'] += count
        protein_totals[protein]['grams'] += grams
        protein_totals[protein]['packages'][pkg] += count

        station_totals.setdefault(station, {'count': 0.0, 'grams': 0.0, 'items': set(), 'rows': 0})
        station_totals[station]['count'] += count
        station_totals[station]['grams'] += grams
        station_totals[station]['items'].add(protein)
        station_totals[station]['rows'] += 1
    return package_totals, protein_totals, station_totals


def _append_rows(ws, headers, rows):
    ws.append(headers)
    for row in rows:
        ws.append(row)


def _build_operations_workbook(day_label, dont_use_rows, pivot_rows, package_order, report, worker_links):
    package_totals, protein_totals, station_totals = _summaries(dont_use_rows)
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet('Dashboard')
    _append_rows(ws, ['المؤشر', 'القيمة'], [
        ['اليوم', day_label],
        ['صفوف فاتورة المشتركين', report.get('source_rows')],
        ['صفوف التشغيل المحسوبة', report.get('computed_rows')],
        ['صفوف Pivot', report.get('row_labels')],
        ['الباقات النهائية', '، '.join(package_order)],
        ['أصناف جديدة تم التعامل معها تلقائيا', len(report.get('inferred_items') or [])],
        ['باقات جديدة تم التعامل معها تلقائيا', len(report.get('inferred_packages') or [])],
    ])
    _style_sheet(ws)

    ws = wb.create_sheet('Packages')
    pkg_rows = []
    for pkg in package_order:
        total = package_totals.get(pkg, {'count': 0, 'grams': 0, 'rows': 0})
        pkg_rows.append([pkg, total['rows'], total['count'], total['grams']])
    for pkg, total in package_totals.items():
        if pkg not in package_order:
            pkg_rows.append([pkg, total['rows'], total['count'], total['grams']])
    _append_rows(ws, ['الباقة النهائية', 'عدد الصفوف', 'إجمالي العدد', 'إجمالي الجرام'], pkg_rows)
    _style_sheet(ws)

    ws = wb.create_sheet('Stations')
    station_rows = []
    preferred_stations = ['breakfast', 'main_kitchen', 'salads', 'dessert', 'addons', 'unclassified']
    for key in preferred_stations:
        total = station_totals.get(key)
        if not total:
            continue
        station_rows.append([
            _station_title(key), total['rows'], len(total['items']),
            total['count'], total['grams'],
        ])
    _append_rows(ws, ['المحطة', 'صفوف التشغيل', 'عدد الأصناف', 'إجمالي العدد', 'إجمالي الجرام'], station_rows)
    _style_sheet(ws)

    ws = wb.create_sheet('By Protein')
    protein_rows = []
    for protein, total in protein_totals.items():
        package_breakdown = ' | '.join(
            f'{pkg}: {round(qty, 2)}'
            for pkg, qty in total['packages'].items()
            if qty
        )
        protein_rows.append([
            protein, total['category'], _station_title(_station_key(total['category'])),
            total['count'], total['grams'], package_breakdown,
        ])
    _append_rows(ws, ['الصنف / البروتين', 'التصنيف', 'المحطة', 'إجمالي العدد', 'إجمالي الجرام', 'تفصيل الباقات'], protein_rows)
    _style_sheet(ws)

    ws = wb.create_sheet('Worker Links')
    _append_rows(ws, ['العامل / القسم', 'الحساب', 'المهمة / الوجهة', 'الرابط', 'الوصف', 'المصدر'], [
        [
            link.get('worker_name'),
            link.get('username'),
            link.get('title'),
            link.get('url'),
            link.get('description'),
            'رابط عامل' if link.get('source') == 'worker' else 'رابط نظام',
        ]
        for link in worker_links
    ])
    _style_sheet(ws)

    inferred = (report.get('inferred_items') or [])
    if inferred:
        ws = wb.create_sheet('New Items')
        _append_rows(ws, ['صنف جديد تم تمريره تلقائيا'], [[item] for item in inferred])
        _style_sheet(ws)

    return wb, {
        'packages': [
            {'name': row[0], 'rows': row[1], 'count': row[2], 'grams': row[3]}
            for row in pkg_rows
        ],
        'stations': [
            {'name': row[0], 'rows': row[1], 'items': row[2], 'count': row[3], 'grams': row[4]}
            for row in station_rows
        ],
        'top_items': [
            {'name': row[0], 'category': row[1], 'station': row[2], 'count': row[3], 'grams': row[4]}
            for row in sorted(protein_rows, key=lambda r: r[3] or 0, reverse=True)[:12]
        ],
    }


def _build_worker_links_workbook(day_label, worker_links):
    wb = Workbook()
    ws = wb.active
    ws.title = 'روابط العاملين'
    _append_rows(ws, ['العامل / القسم', 'الحساب', 'المهمة / الوجهة', 'الرابط', 'الوصف', 'المصدر'], [
        [
            link.get('worker_name'),
            link.get('username'),
            link.get('title'),
            link.get('url'),
            link.get('description'),
            'رابط عامل' if link.get('source') == 'worker' else 'رابط نظام',
        ]
        for link in worker_links
    ])
    _style_sheet(ws)

    info = wb.create_sheet('معلومات')
    _append_rows(info, ['البند', 'القيمة'], [
        ['اليوم', day_label],
        ['عدد الروابط', len(worker_links)],
        ['ملاحظة', 'يتم إضافة روابط العاملين النشطة من قاعدة البيانات إذا كان الاتصال متاحا، مع روابط التشغيل الأساسية.'],
    ])
    _style_sheet(info)
    return wb


def _build_station_outputs_workbook(day_label, dont_use_rows):
    grouped = OrderedDict()
    station_order = ['breakfast', 'main_kitchen', 'salads', 'dessert', 'addons', 'unclassified']
    for key in station_order:
        grouped[key] = OrderedDict()

    for row in dont_use_rows:
        station = _station_key(row.get('التصنيف') or '')
        protein = row.get('Protein') or row.get('الاسم الإنجليزي') or 'غير محدد'
        side = row.get('Side') or '-'
        pkg = row.get('Final_Package') or 'غير محدد'
        key = (protein, side, pkg)
        grouped.setdefault(station, OrderedDict())
        bucket = grouped[station].setdefault(key, {
            'protein': protein,
            'side': side,
            'package': pkg,
            'category': row.get('التصنيف') or '',
            'count': 0.0,
            'grams': 0.0,
            'source_rows': 0,
        })
        bucket['count'] += _num(row.get('Final_Count'))
        bucket['grams'] += _num(row.get('Grams'))
        bucket['source_rows'] += 1

    wb = Workbook()
    wb.remove(wb.active)
    for station in station_order:
        rows = list(grouped.get(station, {}).values())
        ws = wb.create_sheet(_station_title(station))
        _append_rows(ws, ['الصنف / البروتين', 'الصوص / الجانب', 'الباقة النهائية', 'التصنيف', 'العدد', 'الجرام', 'عدد صفوف المصدر'], [
            [
                item['protein'],
                item['side'],
                item['package'],
                item['category'],
                item['count'],
                item['grams'],
                item['source_rows'],
            ]
            for item in rows
        ])
        _style_sheet(ws)

    ws = wb.create_sheet('الصوصات')
    main_rows = [
        item
        for item in grouped.get('main_kitchen', {}).values()
        if item['count'] or item['grams']
    ]
    _append_rows(ws, ['وجبة مرتبطة بالصوص', 'الجانب', 'الباقة النهائية', 'عدد الوجبات', 'ملاحظة'], [
        [item['protein'], item['side'], item['package'], item['count'], 'يحتاج ربط وصفة الصوص من صفحة الصوصات لو مطلوب جرامات صوص دقيقة']
        for item in main_rows
    ])
    _style_sheet(ws)

    ws = wb.create_sheet('ملاحظات')
    _append_rows(ws, ['البند', 'التفاصيل'], [
        ['اليوم', day_label],
        ['مصدر الملف', 'فاتورة المشتركين فقط'],
        ['Weekly Purchasing', 'يحتاج ملفات/مصادر المشتريات الأصلية، لذلك لم يتم توليد ملف وهمي من فاتورة المشتركين'],
        ['استخراج البروتين والصوصات', 'تم تجهيز قائمة تشغيل من الفاتورة، أما الجرامات الدقيقة للصوص تعتمد على وصفات/قوالب الصوص الموجودة في صفحة المحطة'],
        ['Tokyo Production', 'لوحة توكيو تعتمد على ملف توكيو الرئيسي وفحص الماكرو، لذلك تظل كرابط تشغيل منفصل'],
    ])
    _style_sheet(ws)
    return wb


def _worker_links_html(day_label, worker_links):
    rows = '\n'.join(
        f'<a class="link" href="{link.get("url") or "#"}"><b>{link.get("title") or ""}</b><span>{link.get("worker_name") or "تشغيل النظام"} · {link.get("description") or ""}</span></a>'
        for link in worker_links
    )
    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<meta charset="utf-8">
<title>روابط تشغيل {day_label}</title>
<style>
body{{margin:0;background:#17100b;color:#fff3df;font-family:Tahoma,Arial,sans-serif;padding:28px}}
.wrap{{max-width:900px;margin:auto}}h1{{margin:0 0 8px;font-size:28px}}p{{color:#cdbca9}}
.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-top:20px}}
.link{{display:block;text-decoration:none;color:#fff;background:#2a1b13;border:1px solid #6f3a22;border-radius:14px;padding:16px}}
.link b{{display:block;font-size:16px;margin-bottom:7px}}.link span{{color:#d8c2a9;font-size:13px;line-height:1.6}}
</style>
<div class="wrap"><h1>روابط تشغيل اليوم - {day_label}</h1><p>ملف سريع للمدير أو المشرف يفتح منه صفحات التشغيل المطلوبة.</p><div class="grid">{rows}</div></div>
"""


def process_day_operations(file_storage, day_label_override=None):
    lookup = load_lookup()
    worker_links = _load_worker_links()
    rows, detected_day = read_subscribers_invoice(file_storage)
    if not rows:
        raise ValueError('الملف المرفوع فاضي أو مالوش صفوف تشغيل')

    day_label = day_label_override or detected_day
    if not day_label:
        raise ValueError('مش قادر أحدد اليوم. اكتب اسم اليوم أو سمّي الملف بتاريخ YYYY-MM-DD.')

    dont_use_rows, pivot_rows, package_order, report = compute_decision_tables(rows, lookup)
    report['day_label'] = day_label

    decision_path = build_output_workbook(
        day_label, rows, dont_use_rows, pivot_rows, package_order, lookup,
        out_path=tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name,
    )
    ops_wb, ops_summary = _build_operations_workbook(day_label, dont_use_rows, pivot_rows, package_order, report, worker_links)

    ops_buf = io.BytesIO()
    ops_wb.save(ops_buf)
    ops_buf.seek(0)

    full_report = {
        **report,
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'files': [
            f'ملف اتخاذ القرار - {day_label}.xlsx',
            f'ملخص تشغيل اليوم - {day_label}.xlsx',
            f'مخرجات محطات التجهيز - {day_label}.xlsx',
            f'قائمة روابط العاملين - {day_label}.xlsx',
            f'روابط التشغيل - {day_label}.html',
        ],
        'operations': ops_summary,
        'worker_links_count': len(worker_links),
    }

    station_wb = _build_station_outputs_workbook(day_label, dont_use_rows)
    station_buf = io.BytesIO()
    station_wb.save(station_buf)
    station_buf.seek(0)

    links_wb = _build_worker_links_workbook(day_label, worker_links)
    links_buf = io.BytesIO()
    links_wb.save(links_buf)
    links_buf.seek(0)

    tokyo_path = None
    tokyo_report = None
    tokyo_error = None
    if os.path.exists(TOKYO_TEMPLATE_PATH):
        try:
            file_storage.seek(0)
            tokyo_day_no, tokyo_meals, tokyo_input_report = read_day_file_payload(file_storage)
            tokyo_path, tokyo_report = merge_day_into_template(
                TOKYO_TEMPLATE_PATH,
                tokyo_day_no,
                tokyo_meals,
                out_path=tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False).name,
            )
            tokyo_report['input_report'] = tokyo_input_report
        except Exception as exc:
            tokyo_error = str(exc)
        finally:
            file_storage.seek(0)
    else:
        tokyo_error = 'ملف قالب توكيو غير موجود على السيرفر'

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        with open(decision_path, 'rb') as fh:
            zf.writestr(f'ملف اتخاذ القرار - {day_label}.xlsx', fh.read())
        zf.writestr(f'ملخص تشغيل اليوم - {day_label}.xlsx', ops_buf.getvalue())
        zf.writestr(f'مخرجات محطات التجهيز - {day_label}.xlsx', station_buf.getvalue())
        if tokyo_path:
            with open(tokyo_path, 'rb') as fh:
                zf.writestr(f'شيت توكيو المحدث - {day_label}.xlsm', fh.read())
        zf.writestr(f'قائمة روابط العاملين - {day_label}.xlsx', links_buf.getvalue())
        zf.writestr(f'روابط التشغيل - {day_label}.html', _worker_links_html(day_label, worker_links))
    if tokyo_path:
        full_report['files'].insert(3, f'شيت توكيو المحدث - {day_label}.xlsm')
    full_report['tokyo'] = tokyo_report or {'error': tokyo_error}
    zip_buf.seek(0)
    return zip_buf, full_report


def _archive_id(day_label):
    stamp = datetime.now().strftime('%Y%m%d-%H%M%S')
    safe_day = re.sub(r'[^\w\u0600-\u06FF-]+', '-', str(day_label or 'day')).strip('-')[:40]
    return f'{stamp}-{safe_day}-{uuid.uuid4().hex[:6]}'


def _ensure_archive_dir():
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    return ARCHIVE_DIR


def save_day_operations_archive(file_storage, day_label_override=None):
    zip_buf, report = process_day_operations(file_storage, day_label_override=day_label_override)
    archive_id = _archive_id(report.get('day_label'))
    archive_dir = _ensure_archive_dir()
    zip_name = f'{archive_id}.zip'
    zip_path = os.path.join(archive_dir, zip_name)
    report = {
        **report,
        'archive_id': archive_id,
        'saved_at': datetime.now().isoformat(timespec='seconds'),
        'zip_name': zip_name,
    }
    with open(zip_path, 'wb') as fh:
        fh.write(zip_buf.getvalue())
    with open(os.path.join(archive_dir, f'{archive_id}.json'), 'w', encoding='utf-8') as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)
    return report


def list_day_operations_archives():
    archive_dir = _ensure_archive_dir()
    items = []
    for name in os.listdir(archive_dir):
        if not name.endswith('.json'):
            continue
        try:
            with open(os.path.join(archive_dir, name), encoding='utf-8') as fh:
                report = json.load(fh)
        except Exception:
            continue
        archive_id = report.get('archive_id') or name[:-5]
        zip_path = os.path.join(archive_dir, report.get('zip_name') or f'{archive_id}.zip')
        items.append({
            'archive_id': archive_id,
            'day_label': report.get('day_label'),
            'saved_at': report.get('saved_at') or report.get('generated_at'),
            'source_rows': report.get('source_rows'),
            'computed_rows': report.get('computed_rows'),
            'row_labels': report.get('row_labels'),
            'files': report.get('files') or [],
            'worker_links_count': report.get('worker_links_count') or 0,
            'tokyo': report.get('tokyo') or {},
            'size': os.path.getsize(zip_path) if os.path.exists(zip_path) else 0,
        })
    items.sort(key=lambda item: item.get('saved_at') or '', reverse=True)
    return items


def get_day_operations_archive_path(archive_id):
    safe = re.sub(r'[^A-Za-z0-9_\-\u0600-\u06FF]+', '', str(archive_id or ''))
    path = os.path.join(_ensure_archive_dir(), f'{safe}.zip')
    if not safe or not os.path.exists(path):
        raise FileNotFoundError('الأرشيف غير موجود')
    return path
