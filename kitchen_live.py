import json
import os
import threading
from datetime import datetime, timezone

from flask import abort, jsonify, request, send_file
from pypdf import PdfReader


BASE_DIR = os.path.dirname(__file__)
LIVE_DIR = os.path.join(BASE_DIR, "data", "kitchen_live")
PDF_DIR = os.path.join(LIVE_DIR, "pdfs")
STATE_PATH = os.path.join(LIVE_DIR, "state.json")
_LOCK = threading.Lock()


STATIONS = {
    "marination": {
        "title_ar": "شاشة التتبيل",
        "title_en": "Marination",
        "keywords": ["marination", "marinade", "تتبيل", "التتبيل"],
        "accent": "#f47a2a",
    },
    "hot": {
        "title_ar": "شاشة القسم الساخن",
        "title_en": "Hot Section",
        "keywords": ["hot", "section", "ساخن", "القسم الساخن"],
        "accent": "#f14a2e",
    },
    "sauce": {
        "title_ar": "شاشة الصوص",
        "title_en": "Sauce",
        "keywords": ["sauce", "صوص", "الصوص"],
        "accent": "#25a7b7",
    },
    "breakfast": {
        "title_ar": "شاشة الإفطار",
        "title_en": "Breakfast",
        "keywords": ["breakfast", "فطور", "افطار", "الإفطار", "الفطار"],
        "accent": "#e9a52d",
    },
    "dessert": {
        "title_ar": "شاشة الحلويات",
        "title_en": "Dessert",
        "keywords": ["dessert", "desserts", "حلى", "حلويات", "الحلويات"],
        "accent": "#b75ddc",
    },
    "rice_actuals": {
        "title_ar": "شاشة الأرز - الفعلي",
        "title_en": "Rice Actuals",
        "keywords": ["rice_actuals", "actuals", "rice actual", "ارز فعلي", "الأرز فعلي"],
        "accent": "#4d87d6",
    },
    "rice_batches": {
        "title_ar": "شاشة الأرز - الدفعات",
        "title_en": "Rice Batches",
        "keywords": ["rice_batches", "batches", "rice batch", "ارز دفعات", "دفعات"],
        "accent": "#5aa56b",
    },
}


def _now_iso():
    return datetime.now(timezone.utc).isoformat()


def _ensure_dirs():
    os.makedirs(PDF_DIR, exist_ok=True)


def _empty_station_state(key):
    station = STATIONS[key]
    return {
        "key": key,
        "title_ar": station["title_ar"],
        "title_en": station["title_en"],
        "accent": station["accent"],
        "pdf_name": "",
        "pdf_url": "",
        "pages": 0,
        "page": 1,
        "zoom": 115,
        "message": "",
        "enabled": True,
        "fit_mode": "page",
        "auto_scroll": False,
        "uploaded_at": "",
        "updated_at": "",
    }


def _default_state():
    return {key: _empty_station_state(key) for key in STATIONS}


def _read_state():
    _ensure_dirs()
    if not os.path.exists(STATE_PATH):
        return _default_state()
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as handle:
            loaded = json.load(handle)
    except (OSError, json.JSONDecodeError):
        loaded = {}
    state = _default_state()
    for key, value in loaded.items():
        if key in state and isinstance(value, dict):
            state[key].update(value)
            state[key]["title_ar"] = STATIONS[key]["title_ar"]
            state[key]["title_en"] = STATIONS[key]["title_en"]
            state[key]["accent"] = STATIONS[key]["accent"]
    return state


def _write_state(state):
    _ensure_dirs()
    with open(STATE_PATH, "w", encoding="utf-8") as handle:
        json.dump(state, handle, ensure_ascii=False, indent=2)


def _page_count(path):
    try:
        return max(1, len(PdfReader(path).pages))
    except Exception:
        return 1


def _station_from_filename(filename):
    name = (filename or "").lower().replace("-", "_").replace(" ", "_")
    original = (filename or "").lower()
    for key, station in STATIONS.items():
        for keyword in station["keywords"]:
            needle = keyword.lower().replace("-", "_").replace(" ", "_")
            if needle in name or keyword.lower() in original:
                return key
    return None


def _public_station_state(key, value):
    result = dict(value)
    pdf_path = os.path.join(PDF_DIR, f"{key}.pdf")
    if os.path.exists(pdf_path):
        pages = _clamp_int(result.get("pages"), 1, 9999, 0)
        if pages <= 0:
            pages = _page_count(pdf_path)
        result["pages"] = pages
        result["page"] = _clamp_int(result.get("page"), 1, pages, 1)
        version = result.get("updated_at") or result.get("uploaded_at") or ""
        result["pdf_url"] = f"/api/kitchen-live/pdf/{key}?v={version}"
    else:
        result["pdf_url"] = ""
        result["pages"] = 0
        result["page"] = 1
    return result


def _clamp_int(value, minimum, maximum, fallback):
    try:
        number = int(value)
    except (TypeError, ValueError):
        return fallback
    return max(minimum, min(maximum, number))


def register_kitchen_live_routes(app):
    @app.get("/api/kitchen-live/state")
    def kitchen_live_state():
        with _LOCK:
            state = _read_state()
        return jsonify({
            "ok": True,
            "stations": [_public_station_state(key, state[key]) for key in STATIONS],
        })

    @app.get("/api/kitchen-live/state/<station_key>")
    def kitchen_live_station_state(station_key):
        if station_key not in STATIONS:
            abort(404)
        with _LOCK:
            state = _read_state()
            station = _public_station_state(station_key, state[station_key])
        return jsonify({"ok": True, "station": station})

    @app.post("/api/kitchen-live/upload")
    def kitchen_live_upload():
        uploaded_files = request.files.getlist("files")
        if not uploaded_files:
            uploaded_files = list(request.files.values())
        if not uploaded_files:
            return jsonify({"ok": False, "error": "لم يتم اختيار ملفات PDF"}), 400

        saved = []
        unmatched = []
        with _LOCK:
            state = _read_state()
            for file_storage in uploaded_files:
                filename = file_storage.filename or ""
                if not filename.lower().endswith(".pdf"):
                    unmatched.append({"file": filename, "reason": "الملف ليس PDF"})
                    continue
                station_key = request.form.get("station") or _station_from_filename(filename)
                if station_key not in STATIONS:
                    unmatched.append({"file": filename, "reason": "لم يتم التعرف على القسم من اسم الملف"})
                    continue
                path = os.path.join(PDF_DIR, f"{station_key}.pdf")
                file_storage.save(path)
                pages = _page_count(path)
                current = state.get(station_key, _empty_station_state(station_key))
                current.update({
                    "pdf_name": filename,
                    "pages": pages,
                    "page": 1,
                    "enabled": True,
                    "uploaded_at": _now_iso(),
                    "updated_at": _now_iso(),
                })
                state[station_key] = current
                saved.append(_public_station_state(station_key, current))
            _write_state(state)
        return jsonify({"ok": True, "saved": saved, "unmatched": unmatched})

    @app.post("/api/kitchen-live/screen/<station_key>")
    def kitchen_live_update_screen(station_key):
        if station_key not in STATIONS:
            abort(404)
        payload = request.get_json(silent=True) or request.form.to_dict()
        with _LOCK:
            state = _read_state()
            station = state[station_key]
            pages = max(1, int(station.get("pages") or 1))
            if "page" in payload:
                station["page"] = _clamp_int(payload.get("page"), 1, pages, station.get("page", 1))
            if "zoom" in payload:
                station["zoom"] = _clamp_int(payload.get("zoom"), 65, 220, station.get("zoom", 115))
            if "message" in payload:
                station["message"] = str(payload.get("message") or "")[:700]
            if "enabled" in payload:
                station["enabled"] = str(payload.get("enabled")).lower() not in {"0", "false", "no", "off"}
            if "fit_mode" in payload and payload.get("fit_mode") in {"page", "width"}:
                station["fit_mode"] = payload["fit_mode"]
            if "auto_scroll" in payload:
                station["auto_scroll"] = str(payload.get("auto_scroll")).lower() in {"1", "true", "yes", "on"}
            station["updated_at"] = _now_iso()
            state[station_key] = station
            _write_state(state)
            public = _public_station_state(station_key, station)
        return jsonify({"ok": True, "station": public})

    @app.get("/api/kitchen-live/pdf/<station_key>")
    def kitchen_live_pdf(station_key):
        if station_key not in STATIONS:
            abort(404)
        path = os.path.join(PDF_DIR, f"{station_key}.pdf")
        if not os.path.exists(path):
            abort(404)
        response = send_file(path, mimetype="application/pdf", as_attachment=False, download_name=f"{station_key}.pdf")
        response.headers["Cache-Control"] = "no-store, max-age=0"
        response.headers["Pragma"] = "no-cache"
        return response
