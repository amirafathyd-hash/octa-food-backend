# ============================================================
# Octa Food — Backend: appointments, notes, and Web Push reminders
# ============================================================
#
# التركيب (خطوات النشر):
# 1) ثبّت المكتبات دي على السيرفر (requirements.txt):
#       pywebpush==2.0.0
#       py-vapid==1.9.0
#
# 2) شغّل السكريبت appointments_schema.sql مرة واحدة في Supabase (SQL Editor).
#
# 3) ولّد مفاتيح VAPID (مرة واحدة بس) على جهازك أو على السيرفر:
#       pip install py-vapid
#       vapid --gen
#    ده هيطلعلك ملفين: private_key.pem و public_key.pem في نفس المجلد.
#    - حط مسار private_key.pem في متغير البيئة VAPID_PRIVATE_KEY_PATH (أو ارفعه كملف على Railway).
#    - علشان تجيب المفتاح العام بصيغة base64url (اللي محتاجه الفرونت إند)، شغّل:
#       vapid --applicationServerKey
#    - انسخ الناتج وحطه في appointments-dashboard.html بدل PASTE_YOUR_VAPID_PUBLIC_KEY_HERE.
#
# 4) في app.py الأساسي بتاعك، ضيف:
#       from appointments_api import appointments_bp
#       app.register_blueprint(appointments_bp)
#
# 5) في متغيرات البيئة على Railway، تأكد إن دول موجودين
#    (لو مختلفين عن الأسماء اللي بتستخدمها فعليًا لباقي المشروع، غيّر الأسماء في أول الملف):
#       SUPABASE_URL
#       SUPABASE_SERVICE_KEY   (Service Role key، مش الـ anon key، عشان الكتابة تشتغل من السيرفر)
#       VAPID_PRIVATE_KEY_PATH  (مسار private_key.pem)
#       VAPID_CLAIM_EMAIL       (إيميلك، مطلوب من مواصفة VAPID، مثال: mailto:you@example.com)
#
# 6) اعمل Cron Job يضرب الـ endpoint بتاع فحص التذكيرات كل دقيقة. على هوستنجر (زي كرون
#    قارئ الميزان بالظبط) تقدر تستخدم curl:
#       * * * * * curl -s https://api.pixivo.org/api/push/check-reminders > /dev/null 2>&1
#    (غيّر الدومين للدومين الحقيقي بتاع الباك إند عندك).
#
# ============================================================

import os
import json
from datetime import datetime, timedelta, timezone

from flask import Blueprint, request, jsonify
from supabase import create_client
from pywebpush import webpush, WebPushException

appointments_bp = Blueprint("appointments_api", __name__)

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY")
VAPID_PRIVATE_KEY_PATH = os.environ.get("VAPID_PRIVATE_KEY_PATH", "private_key.pem")
VAPID_CLAIM_EMAIL = os.environ.get("VAPID_CLAIM_EMAIL", "mailto:admin@example.com")

supabase = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY) if SUPABASE_URL and SUPABASE_SERVICE_KEY else None


def _now_iso():
    return datetime.now(timezone.utc).isoformat()


# ===================== المواعيد =====================

@appointments_bp.route("/api/appointments", methods=["GET"])
def list_appointments():
    res = supabase.table("appointments").select("*").order("starts_at").execute()
    return jsonify({"items": res.data})


@appointments_bp.route("/api/appointments", methods=["POST"])
def create_appointment():
    body = request.get_json(force=True)
    if not body.get("title") or not body.get("starts_at"):
        return jsonify({"error": "title and starts_at are required"}), 400

    row = {
        "title": body["title"],
        "starts_at": body["starts_at"],
        "reminder_minutes_before": body.get("reminder_minutes_before", 10),
        "meet_link": body.get("meet_link"),
        "notes": body.get("notes"),
        "reminded": False,
    }
    res = supabase.table("appointments").insert(row).execute()
    return jsonify({"item": res.data[0] if res.data else row}), 201


@appointments_bp.route("/api/appointments/<appointment_id>", methods=["DELETE"])
def delete_appointment(appointment_id):
    supabase.table("appointments").delete().eq("id", appointment_id).execute()
    return jsonify({"ok": True})


# ===================== النوتس =====================

@appointments_bp.route("/api/notes", methods=["GET"])
def list_notes():
    res = supabase.table("quick_notes").select("*").order("created_at", desc=True).execute()
    return jsonify({"items": res.data})


@appointments_bp.route("/api/notes", methods=["POST"])
def create_note():
    body = request.get_json(force=True)
    if not body.get("text"):
        return jsonify({"error": "text is required"}), 400

    row = {
        "text": body["text"],
        "remind_at": body.get("remind_at"),
        "reminded": False,
    }
    res = supabase.table("quick_notes").insert(row).execute()
    return jsonify({"item": res.data[0] if res.data else row}), 201


@appointments_bp.route("/api/notes/<note_id>", methods=["DELETE"])
def delete_note(note_id):
    supabase.table("quick_notes").delete().eq("id", note_id).execute()
    return jsonify({"ok": True})


# ===================== اشتراكات التنبيهات (Web Push) =====================

@appointments_bp.route("/api/push/subscribe", methods=["POST"])
def push_subscribe():
    body = request.get_json(force=True)
    sub = body.get("subscription")
    if not sub or not sub.get("endpoint"):
        return jsonify({"error": "invalid subscription"}), 400

    keys = sub.get("keys", {})
    row = {
        "endpoint": sub["endpoint"],
        "p256dh": keys.get("p256dh", ""),
        "auth": keys.get("auth", ""),
    }
    # upsert بالـ endpoint عشان ميتكررش نفس الاشتراك
    supabase.table("push_subscriptions").upsert(row, on_conflict="endpoint").execute()
    return jsonify({"ok": True})


def _send_push(subscription_row, title, body_text, url):
    try:
        webpush(
            subscription_info={
                "endpoint": subscription_row["endpoint"],
                "keys": {"p256dh": subscription_row["p256dh"], "auth": subscription_row["auth"]},
            },
            data=json.dumps({"title": title, "body": body_text, "url": url}),
            vapid_private_key=VAPID_PRIVATE_KEY_PATH,
            vapid_claims={"sub": VAPID_CLAIM_EMAIL},
        )
        return True
    except WebPushException as ex:
        # لو الاشتراك بقى غير صالح (المستخدم قفل الإذن أو مسح البيانات)، امسحه
        if ex.response is not None and ex.response.status_code in (404, 410):
            supabase.table("push_subscriptions").delete().eq("endpoint", subscription_row["endpoint"]).execute()
        return False


# ===================== فحص وإرسال التذكيرات (تستدعيها الـ cron job كل دقيقة) =====================

@appointments_bp.route("/api/push/check-reminders", methods=["GET", "POST"])
def check_reminders():
    if supabase is None:
        return jsonify({"error": "Supabase not configured"}), 500

    now = datetime.now(timezone.utc)
    subs = supabase.table("push_subscriptions").select("*").execute().data
    sent_count = 0

    # ---- المواعيد ----
    upcoming = supabase.table("appointments").select("*").eq("reminded", False).execute().data
    for appt in upcoming:
        starts_at = datetime.fromisoformat(appt["starts_at"].replace("Z", "+00:00"))
        remind_at = starts_at - timedelta(minutes=appt.get("reminder_minutes_before", 10))
        if remind_at <= now <= starts_at:
            title = "🔔 تذكير: " + appt["title"]
            body_text = "الميعاد بعد " + str(appt.get("reminder_minutes_before", 10)) + " دقيقة"
            for s in subs:
                if _send_push(s, title, body_text, "appointments-dashboard"):
                    sent_count += 1
            supabase.table("appointments").update({"reminded": True}).eq("id", appt["id"]).execute()

    # ---- النوتس ----
    pending_notes = supabase.table("quick_notes").select("*").eq("reminded", False).not_.is_("remind_at", "null").execute().data
    for note in pending_notes:
        remind_at = datetime.fromisoformat(note["remind_at"].replace("Z", "+00:00"))
        if remind_at <= now:
            title = "📝 تذكير بنوت"
            body_text = note["text"][:120]
            for s in subs:
                if _send_push(s, title, body_text, "appointments-dashboard"):
                    sent_count += 1
            supabase.table("quick_notes").update({"reminded": True}).eq("id", note["id"]).execute()

    return jsonify({"ok": True, "checked_at": _now_iso(), "notifications_sent": sent_count})
