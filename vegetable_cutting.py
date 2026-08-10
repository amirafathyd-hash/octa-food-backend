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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A1, A2, landscape
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

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_FONT_DIRS = (
    os.path.join(_APP_DIR, "fonts"),
    os.path.join(_APP_DIR, "assets", "fonts"),
)
_PDF_FONT_REGULAR = "OctaArabic"
_PDF_FONT_BOLD = "OctaArabicBold"


def _find_pdf_font(*names):
    search_dirs = _FONT_DIRS + (
        "/usr/share/fonts/truetype/freefont",
        "/usr/share/fonts/truetype/dejavu",
        "/usr/local/share/fonts",
        "/Library/Fonts",
        "/System/Library/Fonts/Supplemental",
    )
    for directory in search_dirs:
        for name in names:
            path = os.path.join(directory, name)
            if os.path.isfile(path):
                return path
    return None


def _register_pdf_fonts():
    regular_path = _find_pdf_font(
        "IBMPlexSansArabic-Regular.ttf", "FreeSans.ttf", "DejaVuSans.ttf", "Arial Unicode.ttf", "Arial.ttf",
    )
    bold_path = _find_pdf_font(
        "IBMPlexSansArabic-Bold.ttf", "FreeSansBold.ttf", "DejaVuSans-Bold.ttf", "Arial Bold.ttf",
    ) or regular_path
    if not regular_path or not bold_path:
        raise CuttingWorkbookError(
            "خط PDF العربي غير متاح على الخادم. أعد نشر الـ backend باستخدام Dockerfile المرفق."
        )
    if _PDF_FONT_REGULAR not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(_PDF_FONT_REGULAR, regular_path))
    if _PDF_FONT_BOLD not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(_PDF_FONT_BOLD, bold_path))


def _rtl(value):
    return get_display(arabic_reshaper.reshape(_clean_text(value)))


def _pdf_icon(pdf, kind, x, y, width=68, height=38):
    """Draw a compact vector cutting-method icon without external images."""
    pdf.saveState()
    pdf.translate(x, y)
    pdf.scale(width / 68.0, height / 38.0)
    x, y, width, height = 0, 0, 68, 38
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
    pdf.restoreState()


def _row_methods(row):
    methods = row.get("methods") or []
    if methods:
        return methods
    return [{
        "method": row.get("method") or "غير محدد",
        "weight_grams": float(row.get("weight_grams") or 0),
        "icon": row.get("icon") or "knife",
    }]


def _pdf_row_height(row):
    method_count = len(_row_methods(row))
    return max(34, 12 + method_count * 17)


def build_cutting_pdf(payload):
    _register_pdf_fonts()
    rows = payload.get("rows") or []
    if not rows:
        raise CuttingWorkbookError("لا توجد بيانات لإنشاء ملف PDF")
    if len(rows) > 500:
        raise CuttingWorkbookError("عدد الصفوف أكبر من الحد المسموح")

    margin = 40
    header_height = 126
    columns_height = 40
    footer_height = 28
    row_heights = [_pdf_row_height(row) for row in rows]
    required_height = (margin * 2) + header_height + columns_height + footer_height + sum(row_heights)
    a2_width, a2_height = landscape(A2)
    if required_height <= a2_height:
        page_width, page_height = a2_width, a2_height
    else:
        a1_width, a1_height = landscape(A1)
        if required_height <= a1_height:
            page_width, page_height = a1_width, a1_height
        else:
            # Keep the promise of one page even for unusually large uploads.
            page_width, page_height = a1_width, required_height + 20
    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=(page_width, page_height), pageCompression=1)
    pdf.setTitle(f"Vegetable Cutting - Day {payload.get('day_number', '')}")
    pdf.setAuthor("Octa Food")

    table_left = margin
    table_width = page_width - (margin * 2)
    cutting_width = min(510, table_width * 0.43)
    weight_width = min(170, table_width * 0.14)
    ingredient_width = table_width - cutting_width - weight_width

    # Header band
    header_y = page_height - margin - header_height
    pdf.setFillColor(HexColor("#164F3C"))
    pdf.roundRect(margin, header_y, table_width, header_height, 14, stroke=0, fill=1)
    pdf.setFillColor(HexColor("#B9D94A"))
    pdf.setFont(_PDF_FONT_BOLD, 13)
    pdf.drawRightString(page_width - margin - 28, header_y + 94, _rtl("كشف تجهيز المطبخ"))
    pdf.setFillColor(HexColor("#FFFFFF"))
    pdf.setFont(_PDF_FONT_BOLD, 32)
    day_ar = payload.get("day_name_ar") or f"يوم {payload.get('day_number', '')}"
    pdf.drawRightString(page_width - margin - 28, header_y + 53, _rtl(day_ar))
    pdf.setFont(_PDF_FONT_BOLD, 14)
    pdf.setFillColor(HexColor("#CFDED7"))
    pdf.drawRightString(page_width - margin - 28, header_y + 29, _clean_text(payload.get("day_name_en")))

    pdf.setFillColor(HexColor("#FFFFFF"))
    pdf.setFont(_PDF_FONT_BOLD, 24)
    pdf.drawString(margin + 28, header_y + 68, f"{len(rows):,}")
    pdf.setFont(_PDF_FONT_REGULAR, 11)
    pdf.setFillColor(HexColor("#CFDED7"))
    pdf.drawString(margin + 28, header_y + 48, _rtl("صنف مجمّع"))
    pdf.setFillColor(HexColor("#FFFFFF"))
    pdf.setFont(_PDF_FONT_BOLD, 24)
    total_weight = sum(float(row.get("weight_grams") or 0) for row in rows)
    pdf.drawString(margin + 155, header_y + 68, f"{total_weight:,.0f}")
    pdf.setFont(_PDF_FONT_REGULAR, 11)
    pdf.setFillColor(HexColor("#CFDED7"))
    pdf.drawString(margin + 155, header_y + 48, _rtl("إجمالي جرام"))

    columns_y = header_y - columns_height
    pdf.setFillColor(HexColor("#E4EDCC"))
    pdf.rect(table_left, columns_y, table_width, columns_height, stroke=0, fill=1)
    pdf.setFillColor(HexColor("#234B39"))
    pdf.setFont(_PDF_FONT_BOLD, 13)
    pdf.drawRightString(table_left + ingredient_width - 16, columns_y + 14, _rtl("الصنف"))
    pdf.drawCentredString(table_left + ingredient_width + weight_width / 2, columns_y + 14, _rtl("الإجمالي (جرام)"))
    pdf.drawRightString(table_left + table_width - 18, columns_y + 14, _rtl("طريقة وشكل التقطيع"))

    row_top = columns_y
    cutting_left = table_left + ingredient_width + weight_width
    cutting_right = table_left + table_width
    for row_index, (row, row_height) in enumerate(zip(rows, row_heights)):
        row_y = row_top - row_height
        pdf.setFillColor(HexColor("#FFFFFF") if row_index % 2 == 0 else HexColor("#FBFCFA"))
        pdf.rect(table_left, row_y, table_width, row_height, stroke=0, fill=1)
        pdf.setStrokeColor(HexColor("#E7ECE8"))
        pdf.setLineWidth(0.55)
        pdf.line(table_left, row_y, table_left + table_width, row_y)

        text_y = row_y + (row_height / 2) - 4
        pdf.setFillColor(HexColor("#17211D"))
        pdf.setFont(_PDF_FONT_BOLD, 12.5)
        pdf.drawRightString(table_left + ingredient_width - 16, text_y, _rtl(row.get("ingredient"))[:100])

        pdf.setFillColor(HexColor("#164F3C"))
        pdf.setFont(_PDF_FONT_BOLD, 16)
        weight = float(row.get("weight_grams") or 0)
        pdf.drawCentredString(table_left + ingredient_width + weight_width / 2, text_y, f"{weight:,.0f}")

        methods = _row_methods(row)
        line_gap = 17
        methods_height = len(methods) * line_gap
        method_y = row_y + (row_height + methods_height) / 2 - 13
        icon_w, icon_h = 42, 24
        icon_x = cutting_right - icon_w - 14
        method_right = icon_x - 12
        for method_row in methods:
            _pdf_icon(pdf, method_row.get("icon") or "knife", icon_x, method_y - 6, icon_w, icon_h)
            pdf.setFillColor(HexColor("#164F3C"))
            pdf.setFont(_PDF_FONT_BOLD, 11.5)
            pdf.drawRightString(method_right, method_y, _rtl(method_row.get("method"))[:52])
            pdf.setFillColor(HexColor("#718078"))
            pdf.setFont(_PDF_FONT_REGULAR, 9)
            method_weight = float(method_row.get("weight_grams") or 0)
            pdf.drawString(cutting_left + 14, method_y + 1, f"{method_weight:,.0f} g")
            method_y -= line_gap
        row_top = row_y

    pdf.setFillColor(HexColor("#87938C"))
    pdf.setFont(_PDF_FONT_REGULAR, 10)
    pdf.drawString(margin, 18, "OCTA FOOD - VEGETABLE PREP")
    pdf.drawRightString(page_width - margin, 18, _rtl("تقرير صفحة واحدة"))
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


_MAPPING_LAYOUTS = (
    {
        "name": "breakfast",
        "day_column": "AB",
        "sheet_column": "AA",
        "min_column": 27,
        "day_index": 1,
        "sheet_index": 0,
    },
    {
        "name": "hot_section",
        "day_column": "AJ",
        "sheet_column": "AK",
        "min_column": 36,
        "day_index": 0,
        "sheet_index": 1,
    },
)
_MAPPING_MAX_ROW = 40
_RECIPE_MAX_ROW = 40


def _mapped_recipe_sheets(workbook, day_number):
    """Find the workbook mapping table and return tabs assigned to its active day."""
    sheet_lookup = {
        _clean_text(sheet_name).casefold(): sheet_name
        for sheet_name in workbook.sheetnames
    }
    preferred = ("All_Ingredients", "Ordering", "Marination_Ordering")
    candidates = [workbook[name] for name in preferred if name in workbook.sheetnames]
    if not candidates:
        candidates = list(workbook.worksheets)

    best = None
    for worksheet in candidates:
        for layout in _MAPPING_LAYOUTS:
            valid_pairs = []
            selected_sheets = []
            seen_selected = set()
            for values in worksheet.iter_rows(
                min_row=1,
                max_row=_MAPPING_MAX_ROW,
                min_col=layout["min_column"],
                max_col=layout["min_column"] + 1,
                values_only=True,
            ):
                mapped_day = _as_day(values[layout["day_index"]])
                mapped_name = _clean_text(values[layout["sheet_index"]])
                actual_name = sheet_lookup.get(mapped_name.casefold())
                if not mapped_day or not actual_name:
                    continue
                valid_pairs.append((mapped_day, actual_name))
                selected_key = actual_name.casefold()
                if mapped_day == day_number and selected_key not in seen_selected:
                    selected_sheets.append(actual_name)
                    seen_selected.add(selected_key)

            candidate = {
                "score": len(valid_pairs),
                "control_sheet": worksheet.title,
                "layout": layout,
                "selected_sheets": selected_sheets,
            }
            if best is None or candidate["score"] > best["score"]:
                best = candidate

    if not best or best["score"] == 0:
        raise CuttingWorkbookError(
            "تعذر العثور على جدول ربط الأيام بالتابات في AA/AB أو AJ/AK حتى الصف 40"
        )
    if not best["selected_sheets"]:
        layout = best["layout"]
        raise CuttingWorkbookError(
            f"لا توجد تابات مرتبطة باليوم {day_number} في العمودين "
            f'{layout["day_column"]}/{layout["sheet_column"]} حتى الصف 40'
        )
    return best["selected_sheets"], {
        "type": best["layout"]["name"],
        "control_sheet": best["control_sheet"],
        "day_column": best["layout"]["day_column"],
        "sheet_column": best["layout"]["sheet_column"],
    }


def _recipe_rows(workbook, recipe_sheet_names):
    """Read A/B/H through row 40 from the recipe tabs assigned to the active day."""
    rows = []
    for sheet_name in recipe_sheet_names:
        worksheet = workbook[sheet_name]
        for row in worksheet.iter_rows(
            min_row=1,
            max_row=_RECIPE_MAX_ROW,
            min_col=1,
            max_col=8,
            values_only=True,
        ):
            method, ingredient, weight = row[0], row[1], row[7]
            numeric_weight = _as_weight(weight)
            ingredient_text = _clean_text(ingredient)
            method_text = _clean_text(method)
            if ingredient_text and numeric_weight and _is_cutting_method(method_text):
                rows.append((ingredient_text, numeric_weight, method_text, worksheet.title))
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
        recipe_sheets, mapping = _mapped_recipe_sheets(workbook, day_number)
        rows = _recipe_rows(workbook, recipe_sheets)
        return {
            "filename": file_storage.filename or "workbook.xlsm",
            "day_number": day_number,
            "day_sheet": day_sheet,
            "mode": "mapped_recipe_tabs_a_b_h",
            "mapping": mapping,
            "selected_sheets": recipe_sheets,
            "rows": rows,
        }
    finally:
        workbook.close()


def combine_workbooks(extracted):
    combined = OrderedDict()
    for source in extracted:
        for ingredient, weight, method, source_sheet in source["rows"]:
            ingredient_key = _clean_text(ingredient).casefold()
            method_key = _clean_text(method).casefold()
            key = (ingredient_key, method_key)
            if key not in combined:
                combined[key] = {
                    "ingredient": ingredient,
                    "method": method,
                    "icon": _icon_kind(method),
                    "weight_grams": 0.0,
                    "sources": [],
                    "source_sheets": [],
                }
            combined[key]["weight_grams"] += weight
            if source["filename"] not in combined[key]["sources"]:
                combined[key]["sources"].append(source["filename"])
            if source_sheet not in combined[key]["source_sheets"]:
                combined[key]["source_sheets"].append(source_sheet)

    rows = []
    for index, row in enumerate(combined.values(), start=1):
        row["id"] = index
        row["weight_grams"] = round(row["weight_grams"], 2)
        row["methods"] = [{
            "method": row["method"],
            "weight_grams": row["weight_grams"],
            "icon": row["icon"],
        }]
        rows.append(row)
    return extracted[0]["day_number"], rows


def build_cutting_xlsx(payload):
    rows = payload.get("rows") or []
    if not rows:
        raise CuttingWorkbookError("لا توجد بيانات لإنشاء ملف إكسيل")

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "كشف التقطيع"
    worksheet.sheet_view.rightToLeft = True
    worksheet.sheet_view.showGridLines = False
    worksheet.freeze_panes = "A6"

    dark_green = "164F3C"
    medium_green = "26745A"
    light_green = "E4EDCC"
    lime = "B9D94A"
    white = "FFFFFF"
    muted = "718078"
    border_color = "DDE7E1"
    thin_border = Border(bottom=Side(style="thin", color=border_color))

    worksheet.merge_cells("A1:D1")
    title_cell = worksheet["A1"]
    title_cell.value = "كشف تجهيز وتقطيع الخضار"
    title_cell.font = Font(name="Arial", size=20, bold=True, color=white)
    title_cell.fill = PatternFill("solid", fgColor=dark_green)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[1].height = 38

    day_ar = _clean_text(payload.get("day_name_ar")) or f"اليوم {payload.get('day_number', '')}"
    day_en = _clean_text(payload.get("day_name_en"))
    worksheet.merge_cells("A2:B2")
    worksheet["A2"] = f"اليوم: {day_ar} - {day_en}".strip(" -")
    worksheet["A2"].font = Font(name="Arial", size=12, bold=True, color=dark_green)
    worksheet["A2"].alignment = Alignment(horizontal="right", vertical="center")

    worksheet["C2"] = "عدد الأصناف"
    worksheet["D2"] = len(rows)
    worksheet["C3"] = "إجمالي الكمية (جرام)"
    worksheet["D3"] = sum(float(row.get("weight_grams") or 0) for row in rows)
    for cell in (worksheet["C2"], worksheet["C3"]):
        cell.font = Font(name="Arial", size=10, bold=True, color=muted)
        cell.alignment = Alignment(horizontal="right")
    for cell in (worksheet["D2"], worksheet["D3"]):
        cell.font = Font(name="Arial", size=12, bold=True, color=dark_green)
        cell.fill = PatternFill("solid", fgColor=light_green)
        cell.alignment = Alignment(horizontal="center")
    worksheet["D3"].number_format = "#,##0"
    worksheet.row_dimensions[2].height = 24
    worksheet.row_dimensions[3].height = 24

    headers = ["الصنف", "الكمية (جرام)", "طريقة وشكل التقطيع", "تاب مصدر الأرقام"]
    for column, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=5, column=column, value=header)
        cell.font = Font(name="Arial", size=11, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=medium_green)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[5].height = 30

    first_data_row = 6
    for row_index, row in enumerate(rows, start=first_data_row):
        methods = row.get("methods") or []
        method_text = " | ".join(
            f'{_clean_text(item.get("method"))} ({float(item.get("weight_grams") or 0):,.0f} g)'
            for item in methods
        ) or _clean_text(row.get("method"))
        values = [
            _clean_text(row.get("ingredient")),
            float(row.get("weight_grams") or 0),
            method_text,
            "، ".join(_clean_text(name) for name in (row.get("source_sheets") or []) if _clean_text(name)),
        ]
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row=row_index, column=column, value=value)
            cell.font = Font(name="Arial", size=10, color="17211D")
            cell.alignment = Alignment(
                horizontal="center" if column == 2 else "right",
                vertical="center",
                wrap_text=True,
            )
            cell.border = thin_border
        worksheet.cell(row=row_index, column=2).number_format = "#,##0"
        worksheet.row_dimensions[row_index].height = 36

    last_data_row = first_data_row + len(rows) - 1
    table = Table(displayName="VegetableCuttingTable", ref=f"A5:D{last_data_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium4",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)

    total_row = last_data_row + 2
    worksheet.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=1)
    worksheet.cell(row=total_row, column=1, value="الإجمالي")
    worksheet.cell(row=total_row, column=2, value=f"=SUM(B{first_data_row}:B{last_data_row})")
    for column in range(1, 5):
        cell = worksheet.cell(row=total_row, column=column)
        cell.fill = PatternFill("solid", fgColor=light_green)
        cell.font = Font(name="Arial", size=11, bold=True, color=dark_green)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.cell(row=total_row, column=2).number_format = "#,##0"
    worksheet.row_dimensions[total_row].height = 26

    worksheet.column_dimensions["A"].width = 34
    worksheet.column_dimensions["B"].width = 18
    worksheet.column_dimensions["C"].width = 58
    worksheet.column_dimensions["D"].width = 32
    worksheet.auto_filter.ref = f"A5:D{last_data_row}"
    worksheet.print_title_rows = "1:5"
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.oddFooter.center.text = "OCTA FOOD - VEGETABLE PREP"

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


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
    except Exception as exc:
        return jsonify({
            "error": f"تعذر إنشاء PDF على الخادم: {str(exc)[:180]}"
        }), 500
    day_number = _as_day(payload.get("day_number")) or 1
    return send_file(
        output,
        as_attachment=True,
        download_name=f"Vegetable_Cutting_Day_{day_number}.pdf",
        mimetype="application/pdf",
    )


@vegetable_cutting_bp.route("/api/vegetable-cutting/export-xlsx", methods=["POST"])
def vegetable_cutting_export_xlsx():
    payload = request.get_json(silent=True) or {}
    try:
        output = build_cutting_xlsx(payload)
    except (CuttingWorkbookError, TypeError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({
            "error": f"تعذر إنشاء ملف إكسيل على الخادم: {str(exc)[:180]}"
        }), 500
    day_number = _as_day(payload.get("day_number")) or 1
    return send_file(
        output,
        as_attachment=True,
        download_name=f"Vegetable_Cutting_Day_{day_number}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
