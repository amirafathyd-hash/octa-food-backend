"""Public kitchen-violation reporting and read-only monitoring links.

Records reuse ``upload_log`` so the feature deploys without a database migration.
Media lives in the already provisioned Supabase Storage bucket under an isolated
``kitchen-violations/`` prefix.  Submission and viewing use different bearer
tokens and neither public endpoint accepts an application session.
"""

import json
import os
import re
import secrets
import uuid
from datetime import date, datetime, time, timezone

from flask import Blueprint, Response, jsonify, request

from db import execute_with_retry, get_client


kitchen_violations_bp = Blueprint("kitchen_violations", __name__)

SUBMIT_TOKEN = os.environ.get(
    "KITCHEN_VIOLATION_SUBMIT_TOKEN",
    "kv-report-J8fY3rQv7Xn2Wp6Ks4Mh9Tc5",
)
VIEW_TOKEN = os.environ.get(
    "KITCHEN_VIOLATION_VIEW_TOKEN",
    "kv-board-R4mZ8aQ2Hu7Ns5Xc9Tp3Wk6V",
)
BUCKET_NAME = os.environ.get(
    "KITCHEN_VIOLATION_BUCKET",
    os.environ.get("INVOICE_RECEIPT_BUCKET", "invoice-receipts"),
)
# The existing invoice bucket can be restricted to PDF MIME metadata.  Evidence
# is always served through our API with its original MIME stored in upload_log,
# so use bucket-compatible metadata without changing the actual file bytes.
STORAGE_CONTENT_TYPE = os.environ.get(
    "KITCHEN_VIOLATION_STORAGE_CONTENT_TYPE",
    "application/pdf",
)
MAX_FILE_BYTES = int(os.environ.get("KITCHEN_VIOLATION_MAX_MB", "60")) * 1024 * 1024
ALLOWED_MIME = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/heic": ".heic",
    "image/heif": ".heif",
    "video/mp4": ".mp4",
    "video/webm": ".webm",
    "video/quicktime": ".mov",
    "video/x-m4v": ".m4v",
}

_auth_checker = None


def configure_kitchen_violations(auth_checker):
    global _auth_checker
    _auth_checker = auth_checker


def _require_admin():
    if _auth_checker is None:
        return None, (jsonify({"error": "إعداد التحقق من الجلسة غير مكتمل"}), 500)
    return _auth_checker()


def _token(kind):
    supplied = (
        request.args.get("token")
        or request.form.get("token")
        or request.headers.get("X-Kitchen-Violation-Token")
        or ""
    )
    expected = SUBMIT_TOKEN if kind == "submit" else VIEW_TOKEN
    return bool(supplied) and secrets.compare_digest(str(supplied), str(expected))


def _valid_date(value):
    try:
        return date.fromisoformat(str(value)).isoformat()
    except (TypeError, ValueError):
        return None


def _valid_time(value):
    raw = str(value or "").strip()
    try:
        parsed = time.fromisoformat(raw)
        return parsed.strftime("%H:%M")
    except (TypeError, ValueError):
        return None


def _normalize_type(value):
    text = str(value or "").strip().lower()
    text = re.sub(r"[\u064B-\u065F\u0670]", "", text)
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ة", "ه")
    return re.sub(r"[^\w\u0600-\u06FF]+", " ", text, flags=re.UNICODE).strip()


def _safe_original_name(value, extension):
    raw = os.path.basename(str(value or "evidence"))
    stem = os.path.splitext(raw)[0]
    stem = re.sub(r"[^\w\u0600-\u06FF.-]+", "-", stem, flags=re.UNICODE).strip(".-")
    return f"{stem or 'evidence'}{extension}"


def _payload(row):
    try:
        value = json.loads(row.get("message") or "{}")
    except (TypeError, ValueError):
        return None
    if not isinstance(value, dict) or value.get("event") != "kitchen_violation":
        return None
    value["id"] = row.get("id")
    value["created_at"] = row.get("created_at") or value.get("created_at")
    return value


def _all_records():
    rows = execute_with_retry(
        get_client().table("upload_log")
        .select("id,file_name,item_date,message,created_at")
        .eq("file_type", "kitchen_violation")
        .order("created_at", desc=False)
        .limit(5000)
    ).data or []
    records = [value for value in (_payload(row) for row in rows) if value]
    history = {}
    for record in records:
        key = record.get("type_key") or _normalize_type(record.get("violation_type"))
        previous = list(history.get(key, [])) if key else []
        record["repeated"] = bool(previous)
        record["repeat_count"] = len(previous) + 1
        record["previous_occurrences"] = previous[-12:]
        if key:
            history.setdefault(key, []).append({
                "id": record.get("id"),
                "date": record.get("violation_date"),
                "time": record.get("violation_time"),
            })
    return records


@kitchen_violations_bp.get("/api/kitchen-violations/links")
def kitchen_violation_links():
    _, err = _require_admin()
    if err:
        return err
    return jsonify({
        "submit_path": f"kitchen-violation-report?c={SUBMIT_TOKEN}",
        "dashboard_path": f"kitchen-violations-dashboard?v={VIEW_TOKEN}",
        "max_mb": MAX_FILE_BYTES // (1024 * 1024),
    })


@kitchen_violations_bp.get("/api/public/kitchen-violations/status")
def kitchen_violation_public_status():
    if not _token("submit"):
        return jsonify({"error": "رابط رصد المخالفة غير صالح"}), 403
    return jsonify({"ok": True, "max_mb": MAX_FILE_BYTES // (1024 * 1024)})


@kitchen_violations_bp.post("/api/public/kitchen-violations")
def kitchen_violation_submit():
    if not _token("submit"):
        return jsonify({"error": "رابط رصد المخالفة غير صالح"}), 403

    violation_type = str(request.form.get("violation_type") or "").strip()[:180]
    violation_date = _valid_date(request.form.get("violation_date"))
    violation_time = _valid_time(request.form.get("violation_time"))
    note = str(request.form.get("note") or "").strip()[:1500]
    evidence = request.files.get("evidence")
    if not violation_type:
        return jsonify({"error": "نوع المخالفة مطلوب"}), 400
    if not violation_date or not violation_time:
        return jsonify({"error": "تاريخ وساعة المخالفة مطلوبان"}), 400
    if not evidence or not evidence.filename:
        return jsonify({"error": "صورة أو فيديو المخالفة مطلوب"}), 400

    content_type = str(evidence.mimetype or "").lower().split(";", 1)[0].strip()
    extension = ALLOWED_MIME.get(content_type)
    if not extension:
        return jsonify({"error": "الملف لازم يكون صورة أو فيديو صالحًا"}), 400
    content = evidence.read(MAX_FILE_BYTES + 1)
    if not content:
        return jsonify({"error": "ملف المخالفة فارغ"}), 400
    if len(content) > MAX_FILE_BYTES:
        return jsonify({"error": f"الحد الأقصى لحجم الملف {MAX_FILE_BYTES // (1024 * 1024)}MB"}), 400

    type_key = _normalize_type(violation_type)
    previous = [
        item for item in _all_records()
        if item.get("type_key") == type_key
    ]
    storage_path = (
        f"kitchen-violations/{violation_date[:4]}/{violation_date[5:7]}/"
        f"{violation_date[8:10]}/{uuid.uuid4().hex}{extension}"
    )
    original_name = _safe_original_name(evidence.filename, extension)
    sb = get_client()
    try:
        sb.storage.from_(BUCKET_NAME).upload(
            storage_path,
            content,
            file_options={"content-type": STORAGE_CONTENT_TYPE, "upsert": "false"},
        )
        created_at = datetime.now(timezone.utc).isoformat()
        message = {
            "event": "kitchen_violation",
            "violation_type": violation_type,
            "type_key": type_key,
            "violation_date": violation_date,
            "violation_time": violation_time,
            "note": note,
            "media_type": "video" if content_type.startswith("video/") else "image",
            "mime_type": content_type,
            "storage_path": storage_path,
            "original_name": original_name,
            "file_size": len(content),
            "created_at": created_at,
        }
        inserted = execute_with_retry(sb.table("upload_log").insert({
            "file_type": "kitchen_violation",
            "file_name": original_name,
            "item_date": violation_date,
            "message": json.dumps(message, ensure_ascii=False),
            "level": "warning",
        }))
        row = (inserted.data or [{}])[0]
        return jsonify({
            "ok": True,
            "id": row.get("id"),
            "repeated": bool(previous),
            "repeat_count": len(previous) + 1,
        })
    except Exception as exc:
        try:
            sb.storage.from_(BUCKET_NAME).remove([storage_path])
        except Exception:
            pass
        return jsonify({"error": f"تعذر حفظ المخالفة: {exc}"}), 500


@kitchen_violations_bp.get("/api/public/kitchen-violations")
def kitchen_violation_list():
    if not _token("view"):
        return jsonify({"error": "رابط متابعة المخالفات غير صالح"}), 403
    year = str(request.args.get("year") or "").strip()
    month = str(request.args.get("month") or "").strip()
    day = str(request.args.get("day") or "").strip()
    records = _all_records()
    if re.fullmatch(r"\d{4}", year):
        records = [item for item in records if str(item.get("violation_date") or "").startswith(year)]
    if re.fullmatch(r"\d{2}", month):
        records = [item for item in records if str(item.get("violation_date") or "")[5:7] == month]
    if re.fullmatch(r"\d{2}", day):
        records = [item for item in records if str(item.get("violation_date") or "")[8:10] == day]
    records.sort(
        key=lambda item: (item.get("violation_date") or "", item.get("violation_time") or "", item.get("created_at") or ""),
        reverse=True,
    )
    for item in records:
        item.pop("storage_path", None)
        item["media_url"] = f"/api/public/kitchen-violations/{item['id']}/media?token={VIEW_TOKEN}"
    return jsonify({
        "records": records,
        "summary": {
            "total": len(records),
            "repeated": sum(1 for item in records if item.get("repeated")),
            "images": sum(1 for item in records if item.get("media_type") == "image"),
            "videos": sum(1 for item in records if item.get("media_type") == "video"),
        },
    })


@kitchen_violations_bp.get("/api/public/kitchen-violations/<int:record_id>/media")
def kitchen_violation_media(record_id):
    if not _token("view"):
        return jsonify({"error": "رابط متابعة المخالفات غير صالح"}), 403
    rows = execute_with_retry(
        get_client().table("upload_log")
        .select("id,message")
        .eq("id", record_id)
        .eq("file_type", "kitchen_violation")
        .limit(1)
    ).data or []
    record = _payload(rows[0]) if rows else None
    if not record or not record.get("storage_path"):
        return jsonify({"error": "ملف المخالفة غير موجود"}), 404
    try:
        content = get_client().storage.from_(BUCKET_NAME).download(record["storage_path"])
        response = Response(content, mimetype=record.get("mime_type") or "application/octet-stream")
        response.headers["Cache-Control"] = "private, max-age=300"
        response.headers["Content-Disposition"] = f'inline; filename="evidence-{record_id}"'
        return response
    except Exception as exc:
        return jsonify({"error": f"تعذر تحميل ملف المخالفة: {exc}"}), 500
