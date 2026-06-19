import re
import pdfplumber

# Matches a data line like: "55.72 ( 7.27) %15 48.45  كج5.7 8.50 1"
ROW_RE = re.compile(
    r'([\d.]+)\s*\(\s*([\d.]+)\)\s*%15\s*([\d.]+)\s+(\S*?)([\d.]+)\s+([\d.]+)\s+(.*?)\s*(\d+)\s*$'
)
NAME_PREFIX_RE = re.compile(r'^\s*([\d.]+)%\s*(.*)$')  # "0.0% <name part>"
CONT_RE = re.compile(r'^\s*([\d.]+)\s*(.*)$')           # "0.00 <name part>"


def _fix_arabic(s):
    """pdfplumber extracts this PDF's Arabic glyphs in reversed order; reverse to fix."""
    return s[::-1].strip()


def parse_invoice_pdf(path):
    """Returns dict: date, invoice_no, items: [{name_ar, unit_price, qty, unit_label, total, ...}]"""
    full_text = ''
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            full_text += (page.extract_text() or '') + '\n'

    lines = [l for l in full_text.split('\n')]

    date_m = re.search(r'(\d{4}-\d{2}-\d{2})', full_text)
    inv_m = re.search(r'(INV\d+)', full_text)

    items = []
    pending_name_parts = []

    for i, raw_line in enumerate(lines):
        line = raw_line.strip()
        line = re.sub(r'[\ue000-\uf8ff]', ' ', line)  # strip private-use icon glyphs
        line = re.sub(r'\s+', ' ', line).strip()
        if not line:
            continue

        row_m = ROW_RE.search(line)
        if row_m:
            total, vat, subtotal, unit_label, qty, unit_price, embedded_name, idx = row_m.groups()
            lead = line[:row_m.start()]
            name_parts = list(pending_name_parts)
            if lead.strip():
                name_parts.append(_fix_arabic(lead.strip()))
            if embedded_name.strip():
                name_parts.append(_fix_arabic(embedded_name.strip()))
            # next line may continue the name, prefixed by "0.00"
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                nxt = re.sub(r'[\ue000-\uf8ff]', ' ', nxt)
                nxt = re.sub(r'\s+', ' ', nxt).strip()
                cm = CONT_RE.match(nxt)
                if cm and cm.group(2).strip() and not re.match(r'^[\d.\s%()]*$', nxt):
                    name_parts.append(_fix_arabic(cm.group(2).strip()))

            name_ar = ' '.join(p for p in name_parts if p).strip()
            items.append({
                'index': int(idx),
                'name_ar': name_ar,
                'unit_price': float(unit_price),
                'qty': float(qty),
                'unit_label': _fix_arabic(unit_label) if unit_label else '',
                'subtotal': float(subtotal),
                'vat': float(vat),
                'total': float(total),
            })
            pending_name_parts = []
            continue

        # name-only line before the data row (e.g. "  0.0% فلفل رومي ال")
        pm = NAME_PREFIX_RE.match(line)
        if pm and pm.group(2).strip():
            pending_name_parts = [_fix_arabic(pm.group(2).strip())]

    return {
        'date': date_m.group(1) if date_m else None,
        'invoice_no': inv_m.group(1) if inv_m else None,
        'items': items,
    }


if __name__ == '__main__':
    import json
    import sys
    data = parse_invoice_pdf(sys.argv[1])
    print(json.dumps(data, ensure_ascii=False, indent=2))
