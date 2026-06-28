"""
استخراج بيانات الفاتورة الكاملة من ملف PDF، مع تصحيح ترتيب الحروف العربي.

كتير من مولّدات الفواتير العربية (زي اللي بتستخدمها الموردين هنا) بتخزّن النص العربي
داخل ملف PDF بترتيب حروف معكوس (right-to-left storage order بدون إعادة ترتيب صحيحة) —
فلو قرأت النص "زي ما هو" بيطلع مقلوب. الدالة fix_line بتاخد كل سطر وتصحح ترتيب الحروف
العربية فيه فقط، وتسيب الأرقام والحروف اللاتينية كما هي (عشان الأرقام لا تتقلب).
"""
import re
import unicodedata
import pdfplumber

ARABIC_RUN = re.compile(r'[\u0600-\u06FF]+')


def _fix_token(token):
    # يقسّم الكلمة لقطع: عربي / غير عربي، ويقلب ترتيب القطع (لأن التخزين كان معكوسًا)،
    # وبيقلب ترتيب الحروف داخل القطعة العربية فقط (الأرقام/اللاتيني تفضل كما هي)
    segments = re.findall(r'[\u0600-\u06FF]+|[^\u0600-\u06FF]+', token)
    fixed = []
    for seg in reversed(segments):
        if ARABIC_RUN.fullmatch(seg):
            fixed.append(seg[::-1])
        else:
            fixed.append(seg)
    return ''.join(fixed)


def fix_line(line):
    line = re.sub(r'[\ue000-\uf8ff]', '', line)  # شيل رموز الأيقونات (private-use glyphs)
    norm = unicodedata.normalize('NFKC', line)
    tokens = norm.split(' ')
    tokens = [_fix_token(t) for t in tokens if t]
    tokens.reverse()
    return ' '.join(tokens)


def extract_fixed_lines(pdf_path):
    raw_lines = []
    fixed_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            for line in text.split('\n'):
                raw_lines.append(line)
                fixed_lines.append(fix_line(line))
    return raw_lines, fixed_lines


NUM = r'[\d.,]+'


def _to_float(s):
    if s is None:
        return 0.0
    s = s.replace(',', '').strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_items(fixed_lines):
    """
    كل بند بياخد عادة 2-3 سطور بعد التصحيح:
      - سطر اسم الصنف (ممكن يكون فاضي لو الاسم قصير وداخل سطر الأرقام)
      - سطر الأرقام: # سعر_الوحدة كمية+وحدة  اجمالي_قبل_الضريبة %نسبة_الضريبة قيمة_الضريبة) ( اجمالي
      - سطر تكملة الاسم (لو الاسم طويل) + قيمة الخصم (عادة 0.00)
    """
    # نمط السطر الرقمي: من الآخر للأول (الجزء الثابت أوضح من اليمين)
    row_re = re.compile(
        r'^(?P<idx>\d+)\s+(?P<mid>.*?)\s+'
        r'(?P<before_tax>' + NUM + r')\s+%(?P<tax_pct>\d+)\s+'
        r'(?P<tax_val>' + NUM + r')\)\s*\(\s*(?P<total>' + NUM + r')$'
    )
    qty_unit_re = re.compile(r'^(?P<qty>' + NUM + r')(?P<unit>[\u0600-\u06FF]+)$')

    items = []
    pending_name_parts = []
    header_seen = False
    header_keywords = re.compile(r'(فاتورة|بيانات|الرقم|المنتج|الكمية|الخصم|الضريبة|الاجمالي|اﻻﺠﻤﺎﻟﻲ|الرضيبة|االجمالي|المجموع|قيمة|العميل|موافق|الرئيس)')

    for i, line in enumerate(fixed_lines):
        clean = line.strip()
        if not clean:
            continue
        m = row_re.match(clean)
        if not m:
            if header_keywords.search(clean):
                header_seen = True
                pending_name_parts = []  # أي كلام قبل/فوق رأس الجدول مالوش دعوة بأي صنف
                continue
            if header_seen and re.search(r'[\u0600-\u06FF]', clean):
                pending_name_parts.append(clean)
            continue

        mid_tokens = m.group('mid').split()
        qty_unit_token = None
        for t in reversed(mid_tokens):
            if qty_unit_re.match(t):
                qty_unit_token = t
                break
        qty, unit = '', ''
        if qty_unit_token:
            qm = qty_unit_re.match(qty_unit_token)
            qty, unit = qm.group('qty'), qm.group('unit')
            idx_in_mid = mid_tokens.index(qty_unit_token)
            before_qty = mid_tokens[:idx_in_mid]
        else:
            before_qty = mid_tokens

        unit_price = ''
        inline_name_tokens = []
        for t in reversed(before_qty):
            if re.match(r'^' + NUM + r'$', t) and not unit_price:
                unit_price = t
            else:
                inline_name_tokens.insert(0, t)

        inline_name = ' '.join(inline_name_tokens).strip()
        name = inline_name if inline_name else ' '.join(pending_name_parts).strip()
        name = re.sub(r'\d+\.?\d*%', '', name).strip()
        pending_name_parts = []

        # نلحق سطر تكملة الاسم اللي بعد سطر الأرقام (لو موجود وغير رقمي خالص ومش سطر بند جديد)
        if i + 1 < len(fixed_lines):
            nxt = fixed_lines[i + 1].strip()
            if nxt and re.search(r'[\u0600-\u06FF]', nxt) and not row_re.match(nxt):
                extra = re.sub(r'[\d.]+', '', nxt).strip()
                if extra:
                    name = f'{name} {extra}'.strip()

        items.append({
            'name': name or '(بدون اسم)',
            'unit_price': _to_float(unit_price),
            'qty': _to_float(qty),
            'unit': unit,
            'total_before_tax': _to_float(m.group('before_tax')),
            'tax_pct': _to_float(m.group('tax_pct')),
            'tax_value': _to_float(m.group('tax_val')),
            'total': _to_float(m.group('total')),
        })

    return items


def parse_invoice_full(pdf_path, file_name=''):
    raw_lines, fixed_lines = extract_fixed_lines(pdf_path)
    raw_text = '\n'.join(raw_lines)

    date_m = re.search(r'(20\d{2}-\d{2}-\d{2})', raw_text)
    number_m = re.search(r'\b([A-Za-z]{2,}\d{3,})\b', raw_text)
    tax_numbers = re.findall(r'\b(\d{15})\b', raw_text)

    non_empty_fixed = [l.strip() for l in fixed_lines if l.strip()]
    supplier = non_empty_fixed[0] if non_empty_fixed else ''
    customer = ''
    for l in non_empty_fixed[1:]:
        if re.search(r'(شركة|رشكة|مؤسسة|موسسة|مطعم|مصنع)', l) and l != supplier:
            customer = l
            break

    def find_amount(keyword_patterns, exclude=None):
        for l in fixed_lines:
            s = l.strip()
            if not s:
                continue
            if exclude and re.search(exclude, s):
                continue
            if re.search(keyword_patterns, s):
                m = re.search(r'(' + NUM + r')\s*$', s)
                if m:
                    return _to_float(m.group(1))
        return 0.0

    subtotal = find_amount(r'قبل')
    vat = find_amount(r'قيمة')
    total = find_amount(r'المجموع', exclude=r'المستحق')

    BOILERPLATE = ['موافق عليه', 'المركز الرئيسي', 'فاتورة مبيعات', 'بيانات العميل']

    def clean_party(s):
        for b in BOILERPLATE:
            s = s.replace(b, '')
        return re.sub(r'\s+', ' ', s).strip()

    items = parse_items(fixed_lines)

    return {
        'fileName': file_name,
        'date': date_m.group(1) if date_m else '',
        'number': number_m.group(1) if number_m else '',
        'party': clean_party(customer) or clean_party(supplier),
        'supplier': clean_party(supplier),
        'customer': clean_party(customer),
        'taxNumbers': tax_numbers,
        'subtotal': subtotal,
        'vat': vat,
        'total': total,
        'items': [
            {
                'item': it['name'],
                'qty': it['qty'],
                'unit': it['unit'],
                'unitPrice': it['unit_price'],
                'total': it['total'],
            } for it in items
        ],
        'notes': '' if items else 'لم يتم استخراج بنود — قد تحتاج مراجعة يدوية',
    }
