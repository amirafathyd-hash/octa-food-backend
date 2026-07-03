"""
Octa Food - Excel Sheet -> Images
===================================
بياخد ملف إكسيل (زي Daily_Ordering_YYYY-MM-DD.xlsx) وبيحوّل كل تاب فيه لصورة PNG
منفصلة، وحاطط عنوان فوق كل صورة = اسم التاب + اسم اليوم (مشتق من تاريخ اسم الملف
أو باراميتر منفصل)، مع الحفاظ على نفس تنسيق وألوان الإكسيل الأصلي 100% (بيستخدم
LibreOffice للتحويل، مش إعادة رسم الجدول من الصفر).

الاستخدام:
    python xlsx_to_images.py Daily_Ordering_2026-07-03.xlsx --date 2026-07-03 --outdir images/
"""
import sys
import os
import re
import shutil
import subprocess
import tempfile
import argparse

from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins
from PIL import Image, ImageDraw, ImageFont

def _find_font():
    candidates = [
        '/usr/share/fonts/truetype/freefont/FreeSerifBold.ttf',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    try:
        out = subprocess.run(['fc-match', '-f', '%{file}', 'FreeSerif:bold'],
                              capture_output=True, text=True, timeout=10)
        if out.stdout and os.path.exists(out.stdout.strip()):
            return out.stdout.strip()
    except Exception:
        pass
    raise RuntimeError(
        'مفيش خط عربي متاح على السيرفر (fonts-freefont-ttf). '
        'أضف الباكدج ده في nixpacks.toml/Dockerfile وأعد الـ deploy.'
    )


FONT_BOLD = None  # يتحدد أول مرة نحتاجه بس (lazy) عشان الاستيراد ميفشلش لو الخط لسه مش متظبط
TITLE_BG = (31, 78, 120)      # نفس أزرق الهيدر في الإكسيل
TITLE_FG = (255, 255, 255)
TITLE_FONT_SIZE = 34
TITLE_PAD_Y = 14

# ترقيم الأيام زي ما هو متبع عندك: ١=السبت ... ٧=الجمعة
DAY_NAMES_AR = {
    1: 'السبت', 2: 'الأحد', 3: 'الإثنين', 4: 'الثلاثاء',
    5: 'الأربعاء', 6: 'الخميس', 7: 'الجمعة',
}


def arabic_day_number(date_obj):
    """Python weekday(): Monday=0 ... Sunday=6. نحوّلها لترقيم السبت=1."""
    return (date_obj.weekday() - 5) % 7 + 1


def day_name_from_date_str(date_str):
    import datetime
    dt = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
    n = arabic_day_number(dt)
    return DAY_NAMES_AR[n], n


def add_title_bar(im, sheet_name, day_label, day_num):
    """يضيف شريط عنوان فوق الصورة مباشرة - باستخدام محرك raqm بتاع Pillow
    (نفس تقنية HarfBuzz+FriBidi اللي بيستخدمها المتصفح) عشان العربي يترسم
    صح تلقائيًا من غير أي هاكات يدوية بتفشل مع حالات مختلفة."""
    title = f"{sheet_name} — {day_label} ({day_num})"
    global FONT_BOLD
    if FONT_BOLD is None:
        FONT_BOLD = _find_font()
    font = ImageFont.truetype(FONT_BOLD, TITLE_FONT_SIZE, layout_engine=ImageFont.Layout.RAQM)

    dummy_im = Image.new('RGB', (10, 10))
    dummy_draw = ImageDraw.Draw(dummy_im)
    bbox = dummy_draw.textbbox((0, 0), title, font=font, direction='rtl')
    text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    bar_h = text_h + TITLE_PAD_Y * 2

    new_im = Image.new('RGB', (im.width, im.height + bar_h), TITLE_BG)
    draw = ImageDraw.Draw(new_im)
    draw.text(((im.width - text_w) / 2 - bbox[0], TITLE_PAD_Y - bbox[1]),
               title, font=font, fill=TITLE_FG, direction='rtl')
    new_im.paste(im, (0, bar_h))
    return new_im


def prepare_sheet_for_export(xlsx_path, out_path):
    """بس بيظبط إعدادات الطباعة (صفحة واحدة لكل تاب) من غير ما يلمس أي هيدر أو
    بيانات - العنوان هيتحط بعدين بـ PIL على الصورة الجاهزة."""
    wb = load_workbook(xlsx_path)
    for ws in wb.worksheets:
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = 'landscape'
        ws.page_margins = PageMargins(left=0.15, right=0.15, top=0.15, bottom=0.15,
                                       header=0, footer=0)
    wb.save(out_path)
    return [ws.title for ws in wb.worksheets]


def convert_to_pdf(xlsx_path, workdir):
    if not shutil.which('soffice'):
        raise RuntimeError(
            'LibreOffice (soffice) مش متظبط على السيرفر. '
            'أضف الباكدج libreoffice في nixpacks.toml/Dockerfile وأعد الـ deploy.'
        )
    subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', workdir, xlsx_path],
        check=True, capture_output=True, timeout=120,
    )
    pdf_path = os.path.join(workdir, os.path.splitext(os.path.basename(xlsx_path))[0] + '.pdf')
    if not os.path.exists(pdf_path):
        raise RuntimeError('فشل تحويل الإكسيل لـ PDF')
    return pdf_path


def autocrop_white(im, padding=15):
    """يرجّع نسخة مقصوصة من الصورة على قد المحتوى بس (بدون فراغات بيضا حواليها)."""
    import numpy as np
    arr = np.array(im.convert('RGB'))
    non_white = np.any(arr < 250, axis=2)
    rows = np.any(non_white, axis=1)
    cols = np.any(non_white, axis=0)
    if not rows.any() or not cols.any():
        return im
    top, bottom = np.where(rows)[0][[0, -1]]
    left, right = np.where(cols)[0][[0, -1]]
    top = max(0, top - padding)
    left = max(0, left - padding)
    bottom = min(im.height, bottom + padding)
    right = min(im.width, right + padding)
    return im.crop((left, top, right, bottom))


def pdf_pages_to_png(pdf_path, workdir, prefix='page'):
    subprocess.run(
        ['pdftoppm', '-png', '-r', '150', pdf_path, os.path.join(workdir, prefix)],
        check=True, capture_output=True, timeout=120,
    )
    pages = sorted(
        [f for f in os.listdir(workdir) if f.startswith(prefix) and f.endswith('.png')],
        key=lambda f: int(re.search(r'(\d+)\.png$', f).group(1))
    )
    return [os.path.join(workdir, f) for f in pages]


def xlsx_sheets_to_images(xlsx_path, date_str, outdir):
    os.makedirs(outdir, exist_ok=True)
    day_label, day_num = day_name_from_date_str(date_str)

    with tempfile.TemporaryDirectory() as tmp:
        prepared_xlsx = os.path.join(tmp, 'prepared.xlsx')
        sheet_names = prepare_sheet_for_export(xlsx_path, prepared_xlsx)

        pdf_path = convert_to_pdf(prepared_xlsx, tmp)
        pages = pdf_pages_to_png(pdf_path, tmp)

        # ملحوظة: لو تاب معين طلع أكتر من صفحة PDF (جدول طويل جدًا)، هيبقى فيه
        # صور أكتر من عدد التابات - بيصح مع fitToHeight=1، ولو مش مظبوط هيبان
        # فورًا من عدد الصور الناتج.
        saved = []
        for i, page_path in enumerate(pages):
            sheet_name = sheet_names[i] if i < len(sheet_names) else f"page{i+1}"

            im = Image.open(page_path)
            im = autocrop_white(im)
            im = add_title_bar(im, sheet_name, day_label, day_num)

            safe_name = re.sub(r'[^\w\u0600-\u06FF]+', '_', sheet_name).strip('_')
            dest = os.path.join(outdir, f"{safe_name}_{date_str}.png")
            im.save(dest)
            saved.append(dest)
        return saved


def add_workbook_images_to_zip(zf, wb_or_path, date_str, folder='images', prefix=''):
    """بتاخد Workbook (كائن openpyxl) أو مسار ملف إكسيل، وبتضيف صورة PNG لكل
    تاب فيه جوه zip مفتوح بالفعل (zf) تحت فولدر `folder`. للاستخدام مباشرة
    من app.py من غير ما تتعامل مع مسارات ملفات بنفسك."""
    with tempfile.TemporaryDirectory() as tmp:
        if isinstance(wb_or_path, str):
            xlsx_path = wb_or_path
        else:
            xlsx_path = os.path.join(tmp, 'wb.xlsx')
            wb_or_path.save(xlsx_path)

        img_dir = os.path.join(tmp, 'imgs')
        saved = xlsx_sheets_to_images(xlsx_path, date_str, img_dir)
        for p in saved:
            arcname = f"{folder}/{prefix}{os.path.basename(p)}"
            zf.write(p, arcname)
        return saved


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx_path')
    ap.add_argument('--date', required=True, help='YYYY-MM-DD')
    ap.add_argument('--outdir', default='images')
    args = ap.parse_args()

    saved = xlsx_sheets_to_images(args.xlsx_path, args.date, args.outdir)
    print(f"تم حفظ {len(saved)} صورة في {args.outdir}:")
    for p in saved:
        print(' -', p)
