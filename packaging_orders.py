"""Weekly packaging-order extraction and inventory endpoints."""

from collections import OrderedDict
from datetime import datetime, timezone
import io
import json
import os
import re

from flask import Blueprint, jsonify, request, send_file
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
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
    regular_path, bold_path = _fonts()
    f_title = ImageFont.truetype(bold_path, 45)
    f_sub = ImageFont.truetype(regular_path, 21)
    f_head = ImageFont.truetype(bold_path, 22)
    f_cell = ImageFont.truetype(regular_path, 20)
    f_num = ImageFont.truetype(bold_path, 21)
    margin = 45
    draw.rounded_rectangle((margin, 35, width - margin, header_h - 12), 28, fill="#163B47")
    draw.text((width - margin - 34, 68), _rtl("طلبات التغليف الأسبوعية"), font=f_title, fill="white", anchor="ra")
    draw.text((width - margin - 34, 133), _rtl("تجميع ذكي من شيتات الصباح والمساء"), font=f_sub, fill="#BCE3DD", anchor="ra")
    draw.text((margin + 34, 98), f"{len(rows)}", font=f_title, fill="#FFB84D", anchor="lm")
    draw.text((margin + 115, 103), _rtl("صنف تغليف"), font=f_sub, fill="white", anchor="lm")

    table_x0, table_x1 = margin, width - margin
    item_w, total_w, stock_w, needed_w = 420, 155, 155, 180
    day_w = (table_x1 - table_x0 - item_w - total_w - stock_w - needed_w) / len(days)
    columns = [("item", item_w, "الصنف")]
    columns += [(day, day_w, DAY_AR[day]) for day in days]
    columns += [("total", total_w, "الإجمالي"), ("inventory", stock_w, "المخزون"), ("remaining", needed_w, "المطلوب")]
    y = header_h
    x = table_x0
    for _, col_w, label in columns:
        draw.rectangle((x, y, x + col_w, y + table_head_h), fill="#FFB84D", outline="#E6C58D", width=2)
        draw.text((x + col_w / 2, y + table_head_h / 2), _rtl(label), font=f_head, fill="#15323B", anchor="mm")
        x += col_w
    for index, row in enumerate(rows):
        y0 = y + table_head_h + index * row_h
        fill = "#FFFFFF" if index % 2 == 0 else "#EDF5F3"
        x = table_x0
        for key, col_w, _ in columns:
            draw.rectangle((x, y0, x + col_w, y0 + row_h), fill=fill, outline="#D7E0DE", width=1)
            if key == "item":
                draw.text((x + col_w - 18, y0 + row_h / 2), _rtl(row.get("item")), font=f_cell, fill="#183B42", anchor="rm")
            else:
                if key in days:
                    value = (row.get("daily") or {}).get(key, 0)
                else:
                    value = row.get(key, 0)
                draw.text((x + col_w / 2, y0 + row_h / 2), f"{float(value or 0):,.0f}", font=f_num, fill="#183B42", anchor="mm")
            x += col_w
    footer_y = y + table_head_h + len(rows) * row_h
    draw.rectangle((margin, footer_y, width - margin, footer_y + footer_h), fill="#163B47")
    draw.text((width - margin - 20, footer_y + footer_h / 2), _rtl("OCTA FOOD · PACKAGING ORDERS"), font=f_sub, fill="white", anchor="rm")
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
    columns = ["الصنف"] + [DAY_AR[day] for day in days] + ["الإجمالي", "المخزون", "المطلوب بعد المخزون", "الوحدة"]
    ws.append(columns)
    for row in rows:
        ws.append([row.get("item")] + [(row.get("daily") or {}).get(day, 0) for day in days] + [
            row.get("total", 0), row.get("inventory", 0), row.get("remaining", 0), row.get("unit", "قطعة")
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
            row["remaining"] = round(max(0, row["total"] - stock), 3)
        return jsonify({
            "ok": True, "days": list(days.keys()), "rows": rows, "sources": sources,
            "total_units": round(sum(row["total"] for row in rows), 3),
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
