"""Extract and combine vegetable cutting instructions from Tokyo workbooks."""

from collections import OrderedDict
import math
import re

import openpyxl
from flask import Blueprint, jsonify, request


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
