# ============================================================================
# إضافات مطلوبة لملف app.py الأساسي بتاع الباك إند (Railway)
# محطة التنقية من التقارير إلى اتخاذ القرار
# ============================================================================
#
# 1) ضيف الاستيراد ده في أول app.py مع باقي الاستيرادات (بجوار استيراد
#    tokyo_ordering مثلاً):
#
#    from decision_station import process_subscribers_invoice
#
# 2) ضيف الـ route ده في أي مكان بعد تعريف app = Flask(__name__):

@app.route('/api/decision-station/process', methods=['POST'])
def decision_station_process():
    """محطة التنقية من التقارير إلى اتخاذ القرار: بترفع فاتورة الكمية
    الخام للمشتركين (شيت فيه عمود لكل صنف/باقة/كمية)، والسيستم بيطلعلك
    ملف كامل بنفس شكل ملف اليوم الجاهز (Export / Don't Use just refresh /
    Update / Packages) - بنفس الحساب اللي كان بيتعمل يدوي في الإكسل، من
    غير ما تلمس حاجة."""
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'ارفع فاتورة الكمية للمشتركين باسم file'}), 400

    day_label = request.form.get('day_label') or None

    try:
        out_path, report = process_subscribers_invoice(f, day_label_override=day_label)
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        app.logger.exception('decision_station_process failed')
        return jsonify({'error': f'تعذّر إنشاء ملف محطة التنقية: {exc}'}), 500

    response = send_file(
        out_path, as_attachment=True,
        download_name=f"Octa_Food_Decision_{report['day_label']}.xlsx",
    )
    response.headers['X-Decision-Report'] = json.dumps(report, ensure_ascii=True)
    response.headers['Access-Control-Expose-Headers'] = 'X-Decision-Report, Content-Disposition'
    return response


# 3) الملفات اللي لازم تترفع مع الباك إند على Railway (نفس فولدر app.py):
#      - decision_station.py
#      - data/decision_station_lookup.json
#
# 4) القاموس (data/decision_station_lookup.json) فيه حاليًا الأصناف اللي
#    ظهرت في فاتورة يوم الثلاثاء بس (27 صنف). أول ما يظهر صنف أو باقة
#    جديدة في يوم تاني، المحطة هتوقف وتقولك بالظبط الاسم الناقص برسالة
#    "محطة التنقية محتاجة تحديث القاموس..." - وقتها تبعتلي اسم الصنف
#    الإنجليزي + اسمه العربي (Protein) + الطبق الجانبي لو موجود (Side)
#    + التصنيف، وأنا أضيفهم للقاموس وأرفعلك نسخة محدّثة.
