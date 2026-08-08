"""Extract and combine vegetable cutting instructions from Tokyo workbooks."""

from collections import OrderedDict
import io
import math
import os
import re

import openpyxl
import arabic_reshaper
try:
    from bidi.algorithm import get_display
except ImportError:  # python-bidi 0.6.11+ exposes the Rust implementation here
    from bidi import get_display
from flask import Blueprint, jsonify, request, send_file
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


vegetable_cutting_bp = Blueprint("vegetable_cutting", __name__)

DAY_NAMES_AR = {
    1: "السبت",
    2: "الأحد",
    3: "الإثنين",
    4: "الثلاثاء",
    5: "الأربعاء",
    6: "الخميس",
    7: "الجمعة",
}

DAY_NAMES_EN = {
    1: "Saturday",
    2: "Sunday",
    3: "Monday",
    4: "Tuesday",
    5: "Wednesday",
    6: "Thursday",
    7: "Friday",
}

_METHOD_WORDS = (
    "cut", "slice", "julienne", "wedge", "dice", "cube", "chop",
    "mince", "mash", "puree", "ring", "crescent", "shred", "clean",
)
_NON_METHODS = {
    "category", "cutting method", "protein", "dairy", "pate",
    "vegetables", "spices & seasonings", "bread & bakery", "filling",
}

_FONT_DIR = os.path.join(os.path.dirname(__file__), "assets", "fonts")
_PDF_FONT_REGULAR = "OctaArabic"
_PDF_FONT_BOLD = "OctaArabicBold"


def _register_pdf_fonts():
    if _PDF_FONT_REGULAR not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(
            _PDF_FONT_REGULAR,
            os.path.join(_FONT_DIR, "IBMPlexSansArabic-Regular.ttf"),
        ))
    if _PDF_FONT_BOLD not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(
            _PDF_FONT_BOLD,
            os.path.join(_FONT_DIR, "IBMPlexSansArabic-Bold.ttf"),
        ))


def _rtl(value):
    return get_display(arabic_reshaper.reshape(_clean_text(value)))


def _pdf_icon(pdf, kind, x, y, width=68, height=38):
    """Draw a compact vector cutting-method icon without external images."""
    green = HexColor("#176047")
    brass = HexColor("#B38636")
    pdf.setStrokeColor(green)
    pdf.setFillColor(HexColor("#EEF5E8"))
    pdf.roundRect(x, y, width, height, 10, stroke=0, fill=1)
    pdf.setLineWidth(1.8)

    left, right = x + 12, x + width - 12
    bottom, top = y + 9, y + height - 9

    if kind in ("julienne", "slices", "grated"):
        count = 4 if kind != "grated" else 5
        for idx in range(count):
            yy = bottom + idx * ((top - bottom) / max(count - 1, 1))
            offset = 4 if idx % 2 else 0
            pdf.line(left + offset, yy, right - 5, yy)
        pdf.setStrokeColor(brass)
        pdf.line(right - 3, bottom - 1, right + 2, top + 1)
    elif kind == "dice":
        size = 9
        for row in range(2):
            for col in range(3):
                pdf.rect(left + col * 15, bottom + row * 13, size, size, stroke=1, fill=0)
    elif kind == "rectangles":
        pdf.rect(left, top - 8, 18, 8, stroke=1, fill=0)
        pdf.rect(left + 23, top - 8, 23, 8, stroke=1, fill=0)
        pdf.rect(left + 5, bottom, 25, 8, stroke=1, fill=0)
        pdf.rect(left + 35, bottom, 14, 8, stroke=1, fill=0)
    elif kind == "rings":
        for cx in (left + 11, left + 34):
            pdf.circle(cx, y + height / 2, 9, stroke=1, fill=0)
            pdf.circle(cx, y + height / 2, 4, stroke=1, fill=0)
    elif kind == "wedges":
        for start in (left, left + 25):
            path = pdf.beginPath()
            path.moveTo(start, bottom)
            path.lineTo(start + 11, top)
            path.lineTo(start + 21, bottom)
            path.close()
            pdf.drawPath(path, stroke=1, fill=0)
    elif kind == "crescent":
        pdf.arc(left, bottom - 1, left + 23, top + 1, 70, 220)
        pdf.arc(left + 8, bottom + 1, left + 27, top - 1, 75, 210)
        pdf.arc(left + 27, bottom, right + 2, top, 70, 220)
    elif kind in ("minced", "chunks"):
        points = ((16, 24), (28, 27), (40, 22), (21, 12), (35, 10), (49, 15))
        for px, py in points:
            pdf.circle(x + px, y + py, 2.2 if kind == "minced" else 3.2, stroke=1, fill=0)
    elif kind == "puree":
        path = pdf.beginPath()
        path.moveTo(left, bottom + 2)
        path.curveTo(left + 5, top + 4, left + 19, top + 2, left + 24, top - 3)
        path.curveTo(left + 34, top + 5, right - 2, top - 5, right, bottom + 2)
        path.close()
        pdf.drawPath(path, stroke=1, fill=0)
        pdf.setStrokeColor(brass)
        pdf.line(right - 2, bottom, right + 3, top + 3)
    elif kind == "clean":
        path = pdf.beginPath()
        path.moveTo(left, bottom)
        path.curveTo(left + 5, top + 5, right - 6, top + 4, right, top)
        path.curveTo(right - 6, bottom - 4, left + 14, bottom - 2, left, bottom)
        pdf.drawPath(path, stroke=1, fill=0)
        pdf.line(left + 5, bottom + 2, right - 4, top - 2)
    else:
        pdf.line(left, bottom, right - 5, top)
        pdf.setStrokeColor(brass)
        pdf.line(right - 7, top - 2, right + 2, top + 5)


def build_cutting_pdf(payload):
    _register_pdf_fonts()
    rows = payload.get("rows") or []
    if not rows:
        raise CuttingWorkbookError("لا توجد بيانات لإنشاء ملف PDF")
    if len(rows) > 500:
        raise CuttingWorkbookError("عدد الصفوف أكبر من الحد المسموح")

    page_width, page_height = landscape(A4)
    margin = 30
    header_height = 104
    columns_height = 32
    footer_height = 24
    row_height = 48
    rows_per_page = int((page_height - (margin * 2) - header_height - columns_height - footer_height) // row_height)
    page_count = math.ceil(len(rows) / rows_per_page)
    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=(page_width, page_height), pageCompression=1)
    pdf.setTitle(f"Vegetable Cutting - Day {payload.get('day_number', '')}")
    pdf.setAuthor("Octa Food")

    table_left = margin
    table_width = page_width - (margin * 2)
    icon_width = 92
    method_width = 170
    weight_width = 125
    ingredient_width = table_width - icon_width - method_width - weight_width

    for page_index in range(page_count):
        # Header band
        header_y = page_height - margin - header_height
        pdf.setFillColor(HexColor("#164F3C"))
        pdf.roundRect(margin, header_y, table_width, header_height, 10, stroke=0, fill=1)
        pdf.setFillColor(HexColor("#B9D94A"))
        pdf.setFont(_PDF_FONT_BOLD, 10)
        pdf.drawRightString(page_width - margin - 22, header_y + 78, _rtl("كشف تجهيز المطبخ"))
        pdf.setFillColor(HexColor("#FFFFFF"))
        pdf.setFont(_PDF_FONT_BOLD, 24)
        day_ar = payload.get("day_name_ar") or f"يوم {payload.get('day_number', '')}"
        pdf.drawRightString(page_width - margin - 22, header_y + 46, _rtl(day_ar))
        pdf.setFont(_PDF_FONT_BOLD, 11)
        pdf.setFillColor(HexColor("#CFDED7"))
        pdf.drawRightString(page_width - margin - 22, header_y + 27, _clean_text(payload.get("day_name_en")))

        pdf.setFillColor(HexColor("#FFFFFF"))
        pdf.setFont(_PDF_FONT_BOLD, 17)
        pdf.drawString(margin + 22, header_y + 56, f"{len(rows):,}")
        pdf.setFont(_PDF_FONT_REGULAR, 8)
        pdf.setFillColor(HexColor("#CFDED7"))
        pdf.drawString(margin + 22, header_y + 42, _rtl("صنف وطريقة"))
        pdf.setFillColor(HexColor("#FFFFFF"))
        pdf.setFont(_PDF_FONT_BOLD, 17)
        total_weight = sum(float(row.get("weight_grams") or 0) for row in rows)
        pdf.drawString(margin + 112, header_y + 56, f"{total_weight:,.0f}")
        pdf.setFont(_PDF_FONT_REGULAR, 8)
        pdf.setFillColor(HexColor("#CFDED7"))
        pdf.drawString(margin + 112, header_y + 42, _rtl("إجمالي جرام"))

        # Column header - order is ingredient, weight, method, visual from left to right.
        columns_y = header_y - columns_height
        pdf.setFillColor(HexColor("#E4EDCC"))
        pdf.rect(table_left, columns_y, table_width, columns_height, stroke=0, fill=1)
        pdf.setFillColor(HexColor("#234B39"))
        pdf.setFont(_PDF_FONT_BOLD, 10)
        pdf.drawRightString(table_left + ingredient_width - 12, columns_y + 11, _rtl("الصنف"))
        pdf.drawCentredString(table_left + ingredient_width + weight_width / 2, columns_y + 11, _rtl("الكمية (جرام)"))
        pdf.drawRightString(table_left + ingredient_width + weight_width + method_width - 12, columns_y + 11, _rtl("طريقة التقطيع"))
        pdf.drawCentredString(table_left + table_width - icon_width / 2, columns_y + 11, _rtl("شكل التقطيع"))

        start = page_index * rows_per_page
        page_rows = rows[start:start + rows_per_page]
        row_top = columns_y
        for row_index, row in enumerate(page_rows):
            row_y = row_top - ((row_index + 1) * row_height)
            pdf.setFillColor(HexColor("#FFFFFF") if row_index % 2 == 0 else HexColor("#FBFCFA"))
            pdf.rect(table_left, row_y, table_width, row_height, stroke=0, fill=1)
            pdf.setStrokeColor(HexColor("#E7ECE8"))
            pdf.setLineWidth(0.5)
            pdf.line(table_left, row_y, table_left + table_width, row_y)

            pdf.setFillColor(HexColor("#17211D"))
            pdf.setFont(_PDF_FONT_BOLD, 10.5)
            ingredient = _rtl(row.get("ingredient"))
            pdf.drawRightString(table_left + ingredient_width - 12, row_y + 18, ingredient[:95])

            pdf.setFillColor(HexColor("#164F3C"))
            pdf.setFont(_PDF_FONT_BOLD, 13)
            weight = float(row.get("weight_grams") or 0)
            pdf.drawCentredString(table_left + ingredient_width + weight_width / 2, row_y + 17, f"{weight:,.0f}")

            pdf.setFont(_PDF_FONT_BOLD, 10)
            method = _rtl(row.get("method"))
            pdf.drawRightString(table_left + ingredient_width + weight_width + method_width - 12, row_y + 18, method[:55])

            icon_x = table_left + ingredient_width + weight_width + method_width + (icon_width - 68) / 2
            _pdf_icon(pdf, row.get("icon") or "knife", icon_x, row_y + 5)

        # Footer and page number
        pdf.setFillColor(HexColor("#87938C"))
        pdf.setFont(_PDF_FONT_REGULAR, 8)
        pdf.drawString(margin, 14, "OCTA FOOD - VEGETABLE PREP")
        pdf.drawRightString(page_width - margin, 14, _rtl(f"صفحة {page_index + 1} من {page_count}"))
        pdf.showPage()

    pdf.save()
    output.seek(0)
    return output


class CuttingWorkbookError(ValueError):
    pass


def _clean_text(value):
    return re.sub(r"\s+", " ", str(value or "").replace("ـ", " ")).strip()


def _as_day(value):
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return None
    if not numeric.is_integer():
        return None
    day = int(numeric)
    return day if day in DAY_NAMES_AR else None


def _as_weight(value):
    if isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number) or number <= 0:
        return None
    return number


def _is_cutting_method(value):
    method = _clean_text(value)
    if not method or method.casefold() in _NON_METHODS:
        return False
    if re.search(r"[\u0600-\u06ff]", method):
        return True
    lowered = method.casefold()
    return any(word in lowered for word in _METHOD_WORDS)


def _icon_kind(method):
    text = _clean_text(method).casefold()
    if any(word in text for word in ("مكعب", "dice", "cube")):
        return "dice"
    if any(word in text for word in ("جوليان", "julienne", "شعيرات", "shred")):
        return "julienne"
    if any(word in text for word in ("ويدجز", "أرباع", "wedge")):
        return "wedges"
    if any(word in text for word in ("نص قمر", "crescent")):
        return "crescent"
    if any(word in text for word in ("دائر", "دوائر", "ring")):
        return "rings"
    if any(word in text for word in ("مبشور", "grated", "grate")):
        return "grated"
    if any(word in text for word in ("مستطيل", "rectangle")):
        return "rectangles"
    if any(word in text for word in ("عشوائي", "random")):
        return "chunks"
    if any(word in text for word in ("مهروس", "بورية", "معجون", "mash", "puree")):
        return "puree"
    if any(word in text for word in ("مفروم", "مقطع ناعم", "mince", "chop")):
        return "minced"
    if any(word in text for word in ("تنضيف", "تنظيف", "clean")):
        return "clean"
    if any(word in text for word in ("شرائح", "slice")):
        return "slices"
    return "knife"


def _read_day_number(workbook):
    preferred = ("All_Ingredients", "Ordering", "Marination_Ordering")
    for sheet_name in preferred:
        if sheet_name in workbook.sheetnames:
            day = _as_day(workbook[sheet_name]["R1"].value)
            if day:
                return day, sheet_name
    for worksheet in workbook.worksheets:
        day = _as_day(worksheet["R1"].value)
        if day:
            return day, worksheet.title
    raise CuttingWorkbookError("تعذر قراءة رقم اليوم من الخلية R1")


def _summary_rows(workbook, day_number):
    if "Cutting_Shapes_Ordering" not in workbook.sheetnames:
        return None
    worksheet = workbook["Cutting_Shapes_Ordering"]
    rows = []
    # The prepared output block is F:I: day, ingredient, weight, cutting method.
    for day_value, ingredient, weight, method in worksheet.iter_rows(
        min_row=3, min_col=6, max_col=9, values_only=True
    ):
        if _as_day(day_value) != day_number:
            continue
        numeric_weight = _as_weight(weight)
        ingredient_text = _clean_text(ingredient)
        method_text = _clean_text(method)
        if ingredient_text and numeric_weight and _is_cutting_method(method_text):
            rows.append((ingredient_text, numeric_weight, method_text))
    return rows


def _recipe_rows(workbook):
    rows = []
    ignored_sheets = {
        "Ordering", "All_Ingredients", "Marination_Ordering",
        "Cutting_Shapes_Ordering", "Butchery",
    }
    for worksheet in workbook.worksheets:
        if worksheet.title in ignored_sheets:
            continue
        for row in worksheet.iter_rows(
            min_row=1, min_col=1, max_col=8, values_only=True
        ):
            method, ingredient, weight = row[0], row[1], row[7]
            numeric_weight = _as_weight(weight)
            ingredient_text = _clean_text(ingredient)
            method_text = _clean_text(method)
            if ingredient_text and numeric_weight and _is_cutting_method(method_text):
                rows.append((ingredient_text, numeric_weight, method_text))
    return rows


def extract_workbook(file_storage):
    try:
        workbook = openpyxl.load_workbook(
            file_storage.stream,
            read_only=True,
            data_only=True,
            keep_links=False,
        )
    except Exception as exc:
        raise CuttingWorkbookError(f"تعذر فتح الملف: {exc}") from exc

    try:
        day_number, day_sheet = _read_day_number(workbook)
        rows = _summary_rows(workbook, day_number)
        extraction_mode = "summary" if rows is not None else "recipes"
        if rows is None:
            rows = _recipe_rows(workbook)
        return {
            "filename": file_storage.filename or "workbook.xlsm",
            "day_number": day_number,
            "day_sheet": day_sheet,
            "mode": extraction_mode,
            "rows": rows,
        }
    finally:
        workbook.close()


def combine_workbooks(extracted):
    days = {item["day_number"] for item in extracted}
    if len(days) != 1:
        details = "، ".join(
            f'{item["filename"]}: يوم {item["day_number"]}' for item in extracted
        )
        raise CuttingWorkbookError(f"الملفان ليسا لنفس يوم التشغيل ({details})")

    combined = OrderedDict()
    for source in extracted:
        for ingredient, weight, method in source["rows"]:
            key = (_clean_text(ingredient).casefold(), _clean_text(method).casefold())
            if key not in combined:
                combined[key] = {
                    "ingredient": ingredient,
                    "method": method,
                    "weight_grams": 0.0,
                    "sources": [],
                }
            combined[key]["weight_grams"] += weight
            if source["filename"] not in combined[key]["sources"]:
                combined[key]["sources"].append(source["filename"])

    rows = []
    for index, row in enumerate(combined.values(), start=1):
        row["id"] = index
        row["weight_grams"] = round(row["weight_grams"], 2)
        row["icon"] = _icon_kind(row["method"])
        rows.append(row)
    return next(iter(days)), rows


@vegetable_cutting_bp.route("/api/vegetable-cutting/extract", methods=["POST"])
def vegetable_cutting_extract():
    files = [item for item in request.files.getlist("files") if item and item.filename]
    if len(files) != 2:
        return jsonify({"error": "ارفع ملف توكيو الرئيسي وملف الإفطار معًا"}), 400

    for item in files:
        if not item.filename.lower().endswith((".xlsx", ".xlsm")):
            return jsonify({"error": f"صيغة الملف غير مدعومة: {item.filename}"}), 400

    try:
        extracted = [extract_workbook(item) for item in files]
        day_number, rows = combine_workbooks(extracted)
    except CuttingWorkbookError as exc:
        return jsonify({"error": str(exc)}), 400

    if not rows:
        return jsonify({"error": "لم يتم العثور على أصناف لها طريقة تقطيع وكمية"}), 400

    return jsonify({
        "ok": True,
        "day_number": day_number,
        "day_name_ar": DAY_NAMES_AR[day_number],
        "day_name_en": DAY_NAMES_EN[day_number],
        "rows": rows,
        "total_weight_grams": round(sum(row["weight_grams"] for row in rows), 2),
        "sources": [
            {
                "filename": item["filename"],
                "day_number": item["day_number"],
                "day_cell": f'{item["day_sheet"]}!R1',
                "extraction_mode": item["mode"],
                "matched_rows": len(item["rows"]),
            }
            for item in extracted
        ],
    })


@vegetable_cutting_bp.route("/api/vegetable-cutting/export-pdf", methods=["POST"])
def vegetable_cutting_export_pdf():
    payload = request.get_json(silent=True) or {}
    try:
        output = build_cutting_pdf(payload)
    except (CuttingWorkbookError, TypeError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400
    day_number = _as_day(payload.get("day_number")) or 1
    return send_file(
        output,
        as_attachment=True,
        download_name=f"Vegetable_Cutting_Day_{day_number}.pdf",
        mimetype="application/pdf",
    )
