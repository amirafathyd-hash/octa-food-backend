# تفعيل المواعيد والتنبيهات — خطوات النشر

الملفات دي مش بترفعها على هوستنجر مع باقي الموقع — دي كود Python لازم يتضاف
لمشروع الباك إند بتاعك على Railway (نفس مكان app.py).

## 1) المكتبات
ضيف اللي في `requirements-additions.txt` لملف `requirements.txt` بتاعك، وادفعهم لـ Railway.

## 2) قاعدة البيانات
افتح Supabase → SQL Editor → شغّل `appointments_schema.sql` مرة واحدة.

## 3) مفاتيح VAPID (مرة واحدة بس)
```
pip install py-vapid
vapid --gen
vapid --applicationServerKey
```
- `vapid --gen` بيطلعلك `private_key.pem` و `public_key.pem`.
- `vapid --applicationServerKey` بيطبعلك المفتاح العام بصيغة base64url.
- انسخ الناتج وحطه في `appointments-dashboard.html` بدل النص
  `PASTE_YOUR_VAPID_PUBLIC_KEY_HERE` (سطر `VAPID_PUBLIC_KEY` في أول السكريبت).
- ارفع `private_key.pem` مع مشروع الباك إند (أو خزّنه في متغير بيئة لو عايز أأمن).

## 4) متغيرات البيئة على Railway
| المتغير | القيمة |
|---|---|
| `SUPABASE_URL` | رابط مشروع Supabase بتاعك |
| `SUPABASE_SERVICE_KEY` | الـ Service Role key (مش anon) |
| `VAPID_PRIVATE_KEY_PATH` | مسار ملف `private_key.pem` |
| `VAPID_CLAIM_EMAIL` | مثال: `mailto:you@example.com` |

## 5) ربط الـ Blueprint
في `app.py` الأساسي:
```python
from appointments_api import appointments_bp
app.register_blueprint(appointments_bp)
```

## 6) Cron Job (فحص التذكيرات كل دقيقة)
زي بالظبط كرون قارئ الميزان بتاع تليجرام — من هوستنجر (Advanced → Cron Jobs):
```
* * * * * curl -s https://api.pixivo.org/api/push/check-reminders > /dev/null 2>&1
```
غيّر الدومين لدومين الباك إند الحقيقي عندك.

## 7) الاختبار
- افتح `appointments-dashboard.html`، دوس "تفعيل التنبيهات" ووافق على إذن المتصفح.
- ضيف ميعاد بعد دقيقتين من دلوقتي، والتذكير خليه 1 دقيقة (لو مش موجود في القايمة، اختار 5 وانتظر).
- استنى الكرون يشتغل (أو شغّل `/api/push/check-reminders` يدوي من المتصفح للتجربة السريعة).
- المفروض يوصلك إشعار حتى لو قفلت التاب.
