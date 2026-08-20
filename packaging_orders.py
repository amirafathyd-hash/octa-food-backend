"""Weekly packaging-order extraction and inventory endpoints."""

from collections import OrderedDict
from datetime import datetime, timezone
import base64
import csv
import hashlib
import io
import json
import os
import re
import uuid

from flask import Blueprint, jsonify, request, send_file
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
import requests
try:
    from bidi.algorithm import get_display
except ImportError:
    from bidi import get_display

from db import execute_with_retry, get_client


packaging_orders_bp = Blueprint("packaging_orders", __name__)

DAY_ORDER = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday")
DAY_AR = {
    "Saturday": "السبت", "Sunday": "الأحد", "Monday": "الإثنين",
    "Tuesday": "الثلاثاء", "Wednesday": "الأربعاء", "Thursday": "الخميس",
}
PACKING_ALIASES = {
    "morning": ("Morning Packing", "Morn Packing", "Packing Morn"),
    "evening": ("Evening Packing", "Even Packing", "Packing Even"),
}
COUNT_ITEM_COLUMNS = ((3, 4), (5, 6), (7, 8), (9, 10), (11, 12))
FONT_DIR = os.path.join(os.path.dirname(__file__), "fonts")


class PackagingWorkbookError(ValueError):
    pass


def _clean(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def _key(value):
    text = _clean(value).casefold()
    text = text.translate(str.maketrans("أإآىةؤئ", "ااايهوي"))
    text = re.sub(r"[\sـ_\-]+", "", text)
    return text


def _number(value):
    if isinstance(value, (int, float)):
        return float(value)
    text = _clean(value).replace(",", "")
    try:
        return float(text)
    except (TypeError, ValueError):
        return 0.0


def _canonical_item(value):
    raw = _clean(value)
    key = _key(raw)
    digits = "".join(re.findall(r"\d+", key))
    if "مستطيل" in key:
        size = "16" if "16" in digits else ("28" if "28" in digits else "")
        return f"صحن مستطيل {size} أونز".strip()
    if "ساندوتش" in key or "ساندوتش" in key:
        return "صحن ساندوتش 32 أونز"
    if "مقسم" in key:
        return "صحن مقسم 32 أونز"
    if "دائري" in key or "دايري" in key:
        size = "20" if "20" in digits else ("32" if "32" in digits else "")
        white = " أبيض" if "ابيض" in key else ""
        return f"صحن دائري{white} {size} أونز".strip()
    if "مكرون" in key:
        return "علبة مكرونة"
    if "سلط" in key and "250" in digits:
        return "علبة سلطة 250 مل"
    if "فواكه" in key and "200" in digits:
        return "علبة فواكه 200 مل"
    if "كيس" in key and "5" in digits and "شفاف" in key:
        return "كيس شفاف 5×5"
    if "ذهبي" in key and "2" in digits:
        return "علبة ذهبية 2 أونز"
    if ("علب" in key or "علبه" in key) and "2" in digits and "اونز" in key:
        return "علبة 2 أونز"
    return raw


def _find_sheet(wb, aliases):
    exact = {_key(name): name for name in wb.sheetnames}
    for alias in aliases:
        if _key(alias) in exact:
            return exact[_key(alias)]
    # Structural fallback: only accept sheets whose first cell identifies a packing schedule.
    for name in wb.sheetnames:
        title = _key(wb[name].cell(1, 1).value)
        if "pack" in _key(name) and ("schedule" in title or "جدول" in title):
            wanted_morning = any("morn" in _key(alias) for alias in aliases)
            is_morning = "morn" in _key(name) or "صباح" in title
            if wanted_morning == is_morning:
                return name
    return None


def _read_day_name(ws):
    haystack = " ".join(_clean(ws.cell(row, col).value) for row in range(1, 4) for col in range(1, 4))
    lowered = haystack.casefold()
    for day in DAY_ORDER:
        if day.casefold() in lowered:
            return day
    for day, arabic in DAY_AR.items():
        if _key(arabic) in _key(haystack):
            return day
    return ""


def _sheet_items(ws):
    items = OrderedDict()
    for row in range(4, ws.max_row + 1):
        first = _key(ws.cell(row, 1).value)
        if first in {"total", "totals", "الاجمالي", "الاجماليات"}:
            continue
        for qty_col, item_col in COUNT_ITEM_COLUMNS:
            qty = _number(ws.cell(row, qty_col).value)
            raw_item = _clean(ws.cell(row, item_col).value)
            if qty <= 0 or not raw_item:
                continue
            item = _canonical_item(raw_item)
            items[item] = items.get(item, 0.0) + qty
    return items


def extract_packaging_files(files):
    days = OrderedDict()
    sources = []
    seen_files = set()
    for uploaded in files:
        raw = uploaded.read()
        uploaded.seek(0)
        signature = (len(raw), raw[:4096], raw[-4096:])
        if signature in seen_files:
            continue
        seen_files.add(signature)
        try:
            wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        except Exception as exc:
            raise PackagingWorkbookError(f"تعذر قراءة الملف {uploaded.filename}: {exc}") from exc
        matched = []
        day_name = ""
        file_items = OrderedDict()
        for shift, aliases in PACKING_ALIASES.items():
            sheet_name = _find_sheet(wb, aliases)
            if not sheet_name:
                continue
            ws = wb[sheet_name]
            day_name = day_name or _read_day_name(ws)
            shift_items = _sheet_items(ws)
            for item, qty in shift_items.items():
                file_items[item] = file_items.get(item, 0.0) + qty
            matched.append({"shift": shift, "sheet": sheet_name, "units": round(sum(shift_items.values()), 3)})
        wb.close()
        if len(matched) != 2:
            raise PackagingWorkbookError(
                f"الملف {uploaded.filename} لازم يحتوي شيت صباح وشيت مساء للتغليف"
            )
        if not day_name:
            raise PackagingWorkbookError(f"تعذر تحديد يوم التشغيل داخل {uploaded.filename}")
        if day_name in days:
            raise PackagingWorkbookError(f"تم رفع أكثر من ملف ليوم {DAY_AR[day_name]}")
        days[day_name] = file_items
        sources.append({"filename": uploaded.filename, "day": day_name, "day_ar": DAY_AR[day_name], "sheets": matched})
    if not days:
        raise PackagingWorkbookError("لم يتم العثور على ملفات تغليف صالحة")
    ordered_days = OrderedDict((day, days[day]) for day in DAY_ORDER if day in days)
    all_items = sorted({item for rows in ordered_days.values() for item in rows}, key=_key)
    rows = []
    for item in all_items:
        daily = {day: round(ordered_days[day].get(item, 0.0), 3) for day in ordered_days}
        rows.append({"item": item, "daily": daily, "total": round(sum(daily.values()), 3), "unit": "قطعة"})
    return ordered_days, rows, sources


def _read_inventory_state():
    try:
        rows = execute_with_retry(
            get_client().table("upload_log").select("message,created_at")
            .eq("file_type", "packaging_inventory").order("created_at", desc=True).limit(1)
        ).data or []
        if rows:
            data = json.loads(rows[0].get("message") or "{}")
            return data if isinstance(data, dict) else {}
    except Exception:
        pass
    return {}


def _save_inventory_state(items):
    state = {
        "items": {str(name): max(0, _number(value)) for name, value in items.items() if _clean(name)},
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    execute_with_retry(get_client().table("upload_log").insert({
        "file_type": "packaging_inventory", "file_name": "packaging_inventory",
        "item_date": None, "message": json.dumps(state, ensure_ascii=False), "level": "info",
    }))
    return state


def _inventory_rows_from_matrix(matrix):
    """Find an item/quantity pair without depending on exact sheet names or columns."""
    if not matrix:
        return {}
    item_words = ("الصنف", "اسم الصنف", "item", "items", "product", "name")
    qty_words = ("المخزون", "الكمية", "الرصيد", "stock", "inventory", "quantity", "qty")
    header_index = item_col = qty_col = None
    for row_index, row in enumerate(matrix[:25]):
        keys = [_key(value) for value in row]
        found_item = next((i for i, key in enumerate(keys) if any(_key(word) in key for word in item_words)), None)
        found_qty = next((i for i, key in enumerate(keys) if any(_key(word) in key for word in qty_words)), None)
        if found_item is not None and found_qty is not None and found_item != found_qty:
            header_index, item_col, qty_col = row_index, found_item, found_qty
            break
    if header_index is None:
        # Fallback: score every column as text/item or numeric/quantity.
        width = max((len(row) for row in matrix), default=0)
        best = None
        for text_col in range(width):
            for number_col in range(width):
                if text_col == number_col:
                    continue
                score = 0
                for row in matrix[:100]:
                    name = _clean(row[text_col] if text_col < len(row) else "")
                    raw_qty = row[number_col] if number_col < len(row) else None
                    qty = _number(raw_qty)
                    if name and not re.fullmatch(r"[\d\s.,]+", name) and qty > 0:
                        score += 1
                if best is None or score > best[0]:
                    best = (score, text_col, number_col)
        if not best or best[0] < 1:
            raise PackagingWorkbookError("لم أجد عمود اسم الصنف وعمود المخزون في الملف")
        header_index, item_col, qty_col = -1, best[1], best[2]
    items = OrderedDict()
    for row in matrix[header_index + 1:]:
        name = _clean(row[item_col] if item_col < len(row) else "")
        qty = _number(row[qty_col] if qty_col < len(row) else None)
        if not name or qty < 0:
            continue
        key = _key(name)
        if any(_key(word) == key for word in item_words) or any(word in key for word in ("total", "اجمالي")):
            continue
        canonical = _canonical_item(name)
        items[canonical] = items.get(canonical, 0.0) + qty
    if not items:
        raise PackagingWorkbookError("الملف لا يحتوي أصناف مخزون صالحة")
    return dict(items)


def _read_inventory_spreadsheet(upload):
    filename = _clean(upload.filename)
    ext = os.path.splitext(filename)[1].lower()
    raw = upload.read()
    if ext == ".csv":
        text = raw.decode("utf-8-sig", errors="replace")
        matrix = list(csv.reader(io.StringIO(text)))
    else:
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
        matrix = []
        for ws in wb.worksheets:
            candidate = [list(row) for row in ws.iter_rows(values_only=True)]
            try:
                return _inventory_rows_from_matrix(candidate)
            except PackagingWorkbookError:
                matrix.extend(candidate)
    return _inventory_rows_from_matrix(matrix)


def _read_inventory_image(upload):
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise PackagingWorkbookError("قراءة الصور تحتاج ANTHROPIC_API_KEY على السيرفر")
    raw = upload.read()
    ext = os.path.splitext(_clean(upload.filename))[1].lower()
    media_type = {".png": "image/png", ".webp": "image/webp"}.get(ext, "image/jpeg")
    prompt = """اقرأ هذه الصورة كقائمة مخزون تغليف. استخرج اسم كل صنف والكمية المتاحة فقط.
أعد JSON فقط بهذا الشكل: {\"items\":[{\"name\":\"اسم الصنف\",\"quantity\":10}]}.
لا تخمن رقما غير ظاهر، وتجاهل العناوين والإجماليات والتواريخ."""
    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
        json={
            "model": os.environ.get("ANTHROPIC_VISION_MODEL") or os.environ.get("ANTHROPIC_MODEL") or "claude-sonnet-4-20250514",
            "max_tokens": 3000,
            "messages": [{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": base64.b64encode(raw).decode("ascii")}},
                {"type": "text", "text": prompt},
            ]}],
        }, timeout=90,
    )
    if response.status_code != 200:
        raise PackagingWorkbookError(f"تعذر قراءة الصورة: HTTP {response.status_code}")
    text = "".join(block.get("text", "") for block in response.json().get("content", []) if block.get("type") == "text")
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise PackagingWorkbookError("لم أستطع استخراج قائمة مخزون من الصورة")
    data = json.loads(match.group(0))
    items = {}
    for row in data.get("items", []):
        name = _canonical_item(row.get("name"))
        if _clean(name):
            items[name] = items.get(name, 0.0) + max(0, _number(row.get("quantity")))
    if not items:
        raise PackagingWorkbookError("الصورة لا تحتوي أصناف مخزون واضحة")
    return items


def _save_week_event(record):
    execute_with_retry(get_client().table("upload_log").insert({
        "file_type": "packaging_week_record", "file_name": record["id"],
        "item_date": record.get("date") or None,
        "message": json.dumps(record, ensure_ascii=False), "level": "info",
    }))


def _week_records():
    rows = (execute_with_retry(
        get_client().table("upload_log").select("file_name,message,created_at")
        .eq("file_type", "packaging_week_record").order("created_at", desc=True).limit(500)
    ).data or [])
    latest = OrderedDict()
    for row in rows:
        try:
            record = json.loads(row.get("message") or "{}")
        except Exception:
            continue
        record_id = _clean(record.get("id") or row.get("file_name"))
        if record_id and record_id not in latest:
            record["created_at"] = row.get("created_at")
            latest[record_id] = record
    return list(latest.values())


def _consume_inventory(current, rows):
    updated = {str(key): max(0, _number(value)) for key, value in (current or {}).items()}
    saved_rows = []
    for row in rows or []:
        item = _canonical_item(row.get("item"))
        total = max(0, _number(row.get("total")))
        before = max(0, _number(updated.get(item, 0)))
        used = min(before, total)
        remaining = max(0, total - used)
        updated[item] = round(before - used, 3)
        saved_rows.append({
            "item": item, "total": round(total, 3), "inventory_before": round(before, 3),
            "used_inventory": round(used, 3), "remaining": round(remaining, 3),
        })
    return updated, saved_rows


def _rtl(value):
    return get_display(arabic_reshaper.reshape(_clean(value)))


def _fonts():
    regular = os.path.join(FONT_DIR, "IBMPlexSansArabic-Regular.ttf")
    bold = os.path.join(FONT_DIR, "IBMPlexSansArabic-Bold.ttf")
    return regular, bold


def build_packaging_png(payload):
    rows = payload.get("rows") or []
    days = [day for day in DAY_ORDER if day in (payload.get("days") or [])]
    if not rows or not days:
        raise PackagingWorkbookError("لا توجد بيانات لإنشاء الصورة")
    width = 1800
    header_h, table_head_h, row_h, footer_h = 190, 72, 62, 58
    height = header_h + table_head_h + row_h * len(rows) + footer_h + 70
    image = Image.new("RGB", (width, height), "#F7F4EE")
    draw = ImageDraw.Draw(image)

    def rtl_text(xy, value, font, fill="#183B42", anchor="ra"):
        text = _clean(value)
        try:
            # On the deployed image Pillow uses libraqm, so pass the original
            # Unicode text and let it shape Arabic exactly once.
            draw.text(
                xy, text, font=font, fill=fill, anchor=anchor,
                direction="rtl", language="ar",
            )
        except (KeyError, TypeError, ValueError):
            # Fallback for Pillow builds without libraqm.
            draw.text(xy, _rtl(text), font=font, fill=fill, anchor=anchor)

    regular_path, bold_path = _fonts()
    f_title = ImageFont.truetype(bold_path, 45)
    f_sub = ImageFont.truetype(regular_path, 21)
    f_head = ImageFont.truetype(bold_path, 22)
    f_cell = ImageFont.truetype(regular_path, 20)
    f_num = ImageFont.truetype(bold_path, 21)
    margin = 45
    draw.rounded_rectangle((margin, 35, width - margin, header_h - 12), 28, fill="#163B47")
    rtl_text((width - margin - 34, 68), "طلبات التغليف الأسبوعية", f_title, "white")
    rtl_text((width - margin - 34, 133), "تجميع ذكي من شيتات الصباح والمساء", f_sub, "#BCE3DD")
    draw.text((margin + 34, 98), f"{len(rows)}", font=f_title, fill="#FFB84D", anchor="lm")
    rtl_text((margin + 115, 103), "صنف تغليف", f_sub, "white", anchor="lm")

    table_x0, table_x1 = margin, width - margin
    item_w, total_w, stock_w, needed_w = 420, 155, 155, 180
    day_w = (table_x1 - table_x0 - item_w - total_w - stock_w - needed_w) / len(days)
    columns = [("item", item_w, "الصنف")]
    columns += [(day, day_w, DAY_AR[day]) for day in days]
    columns += [("total", total_w, "الإجمالي"), ("used_inventory", stock_w, "استخدام المخزون"), ("remaining", needed_w, "المطلوب")]
    y = header_h
    x = table_x0
    for _, col_w, label in columns:
        draw.rectangle((x, y, x + col_w, y + table_head_h), fill="#FFB84D", outline="#E6C58D", width=2)
        rtl_text((x + col_w / 2, y + table_head_h / 2), label, f_head, "#15323B", anchor="mm")
        x += col_w
    for index, row in enumerate(rows):
        y0 = y + table_head_h + index * row_h
        fill = "#FFFFFF" if index % 2 == 0 else "#EDF5F3"
        x = table_x0
        for key, col_w, _ in columns:
            draw.rectangle((x, y0, x + col_w, y0 + row_h), fill=fill, outline="#D7E0DE", width=1)
            if key == "item":
                rtl_text((x + col_w - 18, y0 + row_h / 2), row.get("item"), f_cell, "#183B42", anchor="rm")
            else:
                if key in days:
                    value = (row.get("daily") or {}).get(key, 0)
                else:
                    value = row.get(key, 0)
                draw.text((x + col_w / 2, y0 + row_h / 2), f"{float(value or 0):,.0f}", font=f_num, fill="#183B42", anchor="mm")
            x += col_w
    footer_y = y + table_head_h + len(rows) * row_h
    draw.rectangle((margin, footer_y, width - margin, footer_y + footer_h), fill="#163B47")
    draw.text((margin + 20, footer_y + footer_h / 2), "OCTA FOOD · PACKAGING ORDERS", font=f_sub, fill="white", anchor="lm")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    output.seek(0)
    return output


def build_packaging_xlsx(payload):
    rows = payload.get("rows") or []
    days = [day for day in DAY_ORDER if day in (payload.get("days") or [])]
    if not rows or not days:
        raise PackagingWorkbookError("لا توجد بيانات لإنشاء Excel")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Packaging Orders"
    columns = ["الصنف"] + [DAY_AR[day] for day in days] + ["الإجمالي", "المستخدم من المخزون", "المطلوب بعد المخزون", "الوحدة"]
    ws.append(columns)
    for row in rows:
        ws.append([row.get("item")] + [(row.get("daily") or {}).get(day, 0) for day in days] + [
            row.get("total", 0), row.get("used_inventory", 0), row.get("remaining", 0), row.get("unit", "قطعة")
        ])
    dark, gold, pale = "163B47", "FFB84D", "EDF5F3"
    thin = Side(style="thin", color="CBD8D5")
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=gold)
        cell.font = Font(bold=True, color=dark, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=pale if cell.row % 2 == 0 else "FFFFFF")
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 34
    for index in range(2, len(columns) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(index)].width = 16
    ws.freeze_panes = "B2"
    ws.sheet_view.rightToLeft = True
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@packaging_orders_bp.route("/api/packaging-orders/extract", methods=["POST"])
def packaging_orders_extract():
    files = [file for file in request.files.getlist("files") if file and file.filename]
    if not 1 <= len(files) <= 6:
        return jsonify({"error": "ارفع من ملف واحد إلى 6 ملفات تشغيل"}), 400
    if any(not file.filename.lower().endswith((".xlsx", ".xlsm")) for file in files):
        return jsonify({"error": "مسموح فقط بملفات XLSX وXLSM"}), 400
    try:
        days, rows, sources = extract_packaging_files(files)
        inventory = (_read_inventory_state().get("items") or {})
        for row in rows:
            stock = max(0, _number(inventory.get(row["item"], 0)))
            row["inventory"] = round(stock, 3)
            row["used_inventory"] = round(min(stock, row["total"]), 3)
            row["remaining"] = round(max(0, row["total"] - stock), 3)
        fingerprint = json.dumps({"days": list(days.keys()), "rows": rows}, ensure_ascii=False, sort_keys=True)
        return jsonify({
            "ok": True, "days": list(days.keys()), "rows": rows, "sources": sources,
            "total_units": round(sum(row["total"] for row in rows), 3),
            "run_id": hashlib.sha256(fingerprint.encode("utf-8")).hexdigest()[:24],
        })
    except PackagingWorkbookError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"تعذر تجهيز طلبات التغليف: {str(exc)[:180]}"}), 500


@packaging_orders_bp.route("/api/packaging-inventory", methods=["GET", "POST"])
def packaging_inventory():
    if request.method == "GET":
        return jsonify({"ok": True, **_read_inventory_state()})
    payload = request.get_json(silent=True) or {}
    items = payload.get("items") or {}
    if not isinstance(items, dict):
        return jsonify({"error": "بيانات المخزون غير صحيحة"}), 400
    try:
        return jsonify({"ok": True, **_save_inventory_state(items)})
    except Exception as exc:
        return jsonify({"error": f"تعذر حفظ المخزون: {str(exc)[:180]}"}), 500


@packaging_orders_bp.route("/api/packaging-inventory/import", methods=["POST"])
def packaging_inventory_import():
    upload = request.files.get("file")
    if not upload or not upload.filename:
        return jsonify({"error": "اختر صورة أو ملف مخزون"}), 400
    ext = os.path.splitext(upload.filename)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm", ".csv"):
            items = _read_inventory_spreadsheet(upload)
        elif ext in (".png", ".jpg", ".jpeg", ".webp"):
            items = _read_inventory_image(upload)
        else:
            return jsonify({"error": "مسموح بصورة PNG/JPG/WEBP أو ملف XLSX/XLSM/CSV"}), 400
        state = _save_inventory_state(items)
        return jsonify({"ok": True, **state, "count": len(items), "filename": upload.filename})
    except PackagingWorkbookError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"تعذر استيراد المخزون: {str(exc)[:180]}"}), 500


@packaging_orders_bp.route("/api/packaging-orders/weeks", methods=["GET", "POST"])
def packaging_weeks():
    if request.method == "GET":
        try:
            records = [record for record in _week_records() if not record.get("deleted")]
            return jsonify({"ok": True, "records": records})
        except Exception as exc:
            return jsonify({"error": f"تعذر تحميل سجل الأسابيع: {str(exc)[:180]}"}), 500
    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows") or []
    name = _clean(payload.get("name"))
    week_date = _clean(payload.get("date"))
    run_id = _clean(payload.get("run_id"))
    if not rows or not name or not week_date or not run_id:
        return jsonify({"error": "اكتب اسم الأسبوع وتاريخه أولًا"}), 400
    try:
        for existing in _week_records():
            if existing.get("run_id") == run_id and not existing.get("deleted"):
                return jsonify({"error": "تم اعتماد هذه الملفات وخصم مخزونها من قبل"}), 409
        current = (_read_inventory_state().get("items") or {})
        updated, saved_rows = _consume_inventory(current, rows)
        _save_inventory_state(updated)
        now = datetime.now(timezone.utc).isoformat()
        record = {
            "id": uuid.uuid4().hex, "run_id": run_id, "name": name, "date": week_date,
            "rows": saved_rows, "days": payload.get("days") or [], "created_at": now,
            "updated_at": now, "deleted": False,
        }
        _save_week_event(record)
        return jsonify({"ok": True, "record": record, "inventory": updated})
    except Exception as exc:
        return jsonify({"error": f"تعذر اعتماد الأسبوع: {str(exc)[:180]}"}), 500


@packaging_orders_bp.route("/api/packaging-orders/weeks/<record_id>", methods=["PUT", "DELETE"])
def packaging_week_record(record_id):
    try:
        record = next((row for row in _week_records() if row.get("id") == record_id), None)
        if not record or record.get("deleted"):
            return jsonify({"error": "السجل غير موجود"}), 404
        now = datetime.now(timezone.utc).isoformat()
        if request.method == "PUT":
            payload = request.get_json(silent=True) or {}
            record["name"] = _clean(payload.get("name")) or record.get("name")
            record["date"] = _clean(payload.get("date")) or record.get("date")
            record["updated_at"] = now
            _save_week_event(record)
            return jsonify({"ok": True, "record": record})
        # Deleting a weekly transaction reverses its exact inventory usage.
        current = (_read_inventory_state().get("items") or {})
        restored = {str(key): max(0, _number(value)) for key, value in current.items()}
        for row in record.get("rows") or []:
            item = _canonical_item(row.get("item"))
            restored[item] = round(restored.get(item, 0) + max(0, _number(row.get("used_inventory"))), 3)
        _save_inventory_state(restored)
        record["deleted"] = True
        record["updated_at"] = now
        _save_week_event(record)
        return jsonify({"ok": True, "inventory": restored})
    except Exception as exc:
        return jsonify({"error": f"تعذر تعديل السجل: {str(exc)[:180]}"}), 500


@packaging_orders_bp.route("/api/packaging-orders/export-png", methods=["POST"])
def packaging_orders_export_png():
    try:
        output = build_packaging_png(request.get_json(silent=True) or {})
        return send_file(output, as_attachment=True, download_name="Packaging_Orders.png", mimetype="image/png")
    except PackagingWorkbookError as exc:
        return jsonify({"error": str(exc)}), 400


@packaging_orders_bp.route("/api/packaging-orders/export-xlsx", methods=["POST"])
def packaging_orders_export_xlsx():
    try:
        output = build_packaging_xlsx(request.get_json(silent=True) or {})
        return send_file(output, as_attachment=True, download_name="Packaging_Orders.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except PackagingWorkbookError as exc:
        return jsonify({"error": str(exc)}), 400
