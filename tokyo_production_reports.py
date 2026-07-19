"""Create the two daily Tokyo production PDFs from the original workbook.

The ranges and ordering below mirror the workbook's own VBA PDF modules:
BatchPDF, SpecialTablesPDF, ActualsPDF, GarnishPDF and MarinationPDF.
LibreOffice is used only as the spreadsheet renderer/calculation engine; the
recipe cells, formulas, formatting and print tables remain those of the source
Tokyo workbook.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from pypdf import PdfWriter

from tokyo_ordering import (
    DAY_NAMES,
    merge_day_into_template,
    read_day_file_payload,
    validate_raw_targets_for_day,
)


SPECIAL_SHEETS = [
    'Asian Chicken Sandwich', 'Jollof Sauce', 'Almond Chicken',
    'Chicken Caesar Salad', 'Octa Poki Bowl', 'Makloba Veggi',
    'Lasagne Bachamel', 'Beef Lasagne Sauce',
    'Beef philly cheese steak', 'Sigapore Vegetables',
    'Mexican Vegetables', 'Beef Kebab Sandwich',
]

ACTUALS_TYPE_B = {
    'Herbal Potato Wedges', 'Sautee Vegetables (1)', 'Mached Potato(1)',
    'Grilled Vegetables(2)', 'Mached Potato(3)', 'Potato Wedges',
    'Oven Vegetables (3)', 'Mached Potato(4)', 'Sautee Vegetables (4)',
    'Grilled Vegetables(5)', 'Sigapore Vegetables', 'Mexican Vegetables',
    'Chicken Steak Topping', 'Beef Kebab Sandwich', 'Spaghetti pasta (3)',
    'Oven Vegetables (6)', 'Chicken Mandi (1)', 'Beef Zurbian',
    'Chicken Saleeq (3)', 'Chicken Mandi (3)', 'Chicken Makloba',
    'Beef Bokhary', 'Chicken Saleeq (6)', 'Beef Kabli', 'Jollof Sauce',
    'Spaghetti pasta', 'Spaghetti pasta (2)',
}


def _soffice_path() -> str:
    candidates = [
        os.environ.get('LIBREOFFICE_PATH'),
        shutil.which('libreoffice'),
        shutil.which('soffice'),
        '/usr/bin/libreoffice',
        '/opt/libreoffice/program/soffice',
        '/Users/mostafaabdo/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice',
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    raise RuntimeError('LibreOffice غير موجود على السيرفر؛ لا يمكن إنشاء ملفات PDF')


def _run_soffice(source: Path, output_dir: Path, target: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    profile = tempfile.mkdtemp(prefix='tokyo-lo-profile-')
    try:
        cmd = [
            _soffice_path(), '--headless', '--nologo', '--nodefault',
            '--nolockcheck', '--nofirststartwizard',
            f'-env:UserInstallation=file://{profile}',
            '--convert-to', target, '--outdir', str(output_dir), str(source),
        ]
        completed = subprocess.run(cmd, capture_output=True, text=True, timeout=240)
        if completed.returncode != 0:
            raise RuntimeError((completed.stderr or completed.stdout or 'LibreOffice failed').strip())
        suffix = '.pdf' if target.startswith('pdf') else '.xlsx'
        result = output_dir / f'{source.stem}{suffix}'
        if not result.exists():
            raise RuntimeError(f'LibreOffice لم ينشئ الملف المتوقع: {result.name}')
        return result
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def _day_sheets(workbook, master_name: str, day_no: int) -> list[str]:
    ws = workbook[master_name]
    names = []
    for row in range(2, min(ws.max_row, 500) + 1):
        value = ws.cell(row, 36).value
        name = ws.cell(row, 37).value
        try:
            same_day = int(value) == int(day_no)
        except (TypeError, ValueError):
            same_day = False
        if same_day and name:
            clean = str(name).strip()
            if clean in workbook.sheetnames and clean not in names:
                names.append(clean)
    return names


def _number(value) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _batch_ranges(ws_values) -> list[str]:
    ranges = []
    for header in range(62, min(ws_values.max_row, 500) + 1):
        if str(ws_values.cell(header, 2).value or '').strip().lower() != 'ingredient':
            continue
        if _number(ws_values.cell(header + 1, 5).value) <= 0:
            continue
        end = None
        next_header = min(ws_values.max_row, header + 80)
        for row in range(header + 1, next_header + 1):
            label = str(ws_values.cell(row, 2).value or '').strip()
            low = label.lower()
            if row > header + 1 and low == 'ingredient':
                break
            if low.startswith('total ') or low in {'protein', 'pasta', 'potato'}:
                end = row + 1
                break
        if end:
            ranges.append(f'A{max(1, header - 2)}:H{end}')
    return ranges


def _special_range(ws_values) -> str | None:
    ceiling = min(30, ws_values.max_row)
    for row in range(5, ceiling + 1):
        if 'garnish' in str(ws_values.cell(row, 2).value or '').lower():
            ceiling = row - 1
            break
    last = None
    for row in range(ceiling, 4, -1):
        label = str(ws_values.cell(row, 2).value or '').strip()
        if label and label.lower() not in {'ingredient', 'category', 'total'} and not label.lower().startswith('base recipe'):
            last = row
            break
    return f'B2:H{last}' if last else None


def _garnish_range(ws_values) -> str | None:
    if 'garnish' not in str(ws_values.cell(35, 2).value or '').lower():
        return None
    last = 36
    for row in range(37, min(ws_values.max_row, 60) + 1):
        if str(ws_values.cell(row, 2).value or '').strip():
            last = row
        elif last >= 37:
            break
    return f'B35:F{last}'


def _marination_range(ws_values) -> str | None:
    last = 4
    for row in range(5, min(ws_values.max_row, 60) + 1):
        if str(ws_values.cell(row, 36).value or '').strip():
            last = row
        elif last >= 5:
            break
    return f'AJ1:AP{last}' if last >= 5 else None


def _page_setup(ws, day_no: int, section: str) -> None:
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.4 if section in {'batch', 'special', 'marination'} else 0.5
    ws.page_margins.right = ws.page_margins.left
    ws.page_margins.top = 1.0
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3
    ws.print_options.horizontalCentered = True
    ws.oddHeader.center.text = f'&B&16{ws.title}'
    if section == 'marination':
        ws.oddHeader.right.text = f'&B&16Day {day_no}  |  Marination'
    else:
        ws.oddHeader.right.text = f'&B&12Day {day_no}'
    ws.oddFooter.center.text = '&B&10Page &P of &N'


def _make_block_workbook(source: Path, values_path: Path, destination: Path,
                         day_no: int, section: str, day_sheets: list[str]) -> tuple[Path, int]:
    wb = load_workbook(source, data_only=False, keep_links=True)
    values = load_workbook(values_path, data_only=True, keep_links=True)
    included = []
    page_total = 0

    for name in day_sheets:
        if name not in wb.sheetnames or name not in values.sheetnames:
            continue
        ws = wb[name]
        wsv = values[name]
        ranges = []
        if section == 'batch':
            ranges = _batch_ranges(wsv)
        elif section == 'special' and name in SPECIAL_SHEETS:
            found = _special_range(wsv)
            ranges = [found] if found else []
        elif section == 'actuals':
            end_col, end_row = ('T', 30) if name in ACTUALS_TYPE_B else ('V', 32) if name == 'Chicken Mushroom' else ('U', 32)
            ranges = [f'R25:{end_col}{end_row}']
        elif section == 'garnish':
            found = _garnish_range(wsv)
            ranges = [found] if found else []
        elif section == 'marination':
            found = _marination_range(wsv)
            ranges = [found] if found else []

        if not ranges:
            continue
        ws.print_area = ranges
        _page_setup(ws, day_no, section)
        included.append(name)
        page_total += len(ranges)

    if not included:
        wb.close()
        values.close()
        raise RuntimeError(f'لا توجد جداول قابلة للطباعة في قسم {section} لليوم المختار')

    included_set = set(included)
    # LibreOffice includes hidden sheets in whole-workbook PDF export. Freeze
    # the already-calculated values in the selected recipe sheets, then remove
    # non-selected tabs from this disposable print copy. The source XLSM and
    # its formulas are untouched.
    for name in included:
        ws = wb[name]
        value_ws = values[name]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.value = value_ws[cell.coordinate].value
    for ws in list(wb.worksheets):
        if ws.title not in included_set:
            wb.remove(ws)
    wb.active = wb.sheetnames.index(included[0])
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = 'auto'
    wb.save(destination)
    wb.close()
    values.close()
    return destination, page_total


def _merge_pdfs(paths: list[Path], destination: Path) -> Path:
    writer = PdfWriter()
    for path in paths:
        writer.append(str(path))
    with destination.open('wb') as handle:
        writer.write(handle)
    writer.close()
    return destination


def _make_libreoffice_copy(source: Path, destination: Path) -> Path:
    """Make a calculation-only XLSX copy.

    Microsoft 365 stores one Moussaka formula with the internal
    ``_xlfn._TRO_LEADING`` compatibility wrapper. LibreOffice interprets that
    wrapper as an unknown function. Removing the no-op wrapper in this
    disposable renderer copy lets the original formula calculate normally;
    the delivered macro workbook is never changed this way.
    """
    wb = load_workbook(source, data_only=False, keep_vba=False, keep_links=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith('=') and '_xlfn._TRO_LEADING(' in value:
                    cell.value = value.replace('_xlfn._TRO_LEADING(', '(')
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = 'auto'
    wb.save(destination)
    wb.close()
    return destination


def build_tokyo_day_package(template_path: str, uploaded_file, output_dir: str | None = None,
                            safety_overrides=None):
    """Return ``(zip_path, updated_xlsm, report)`` for one uploaded day file."""
    root = Path(output_dir or tempfile.mkdtemp(prefix='tokyo-day-reports-'))
    root.mkdir(parents=True, exist_ok=True)

    day_no, meals, input_report = read_day_file_payload(uploaded_file)
    if input_report.get('kind') == 'repeat_update':
        validate_raw_targets_for_day(template_path, day_no, meals)
    updated_xlsm, match_report = merge_day_into_template(
        template_path, day_no, meals, safety_overrides=safety_overrides
    )
    updated_xlsm = Path(updated_xlsm)

    # A raw Repeat Update file is the authoritative input for the whole day.
    # Never continue with old values left in an unmatched Tokyo recipe: that
    # would make the exported production sheets look valid while incomplete.
    if input_report.get('kind') == 'repeat_update' and match_report['unmatched_count']:
        raise ValueError(
            'تم إيقاف التشغيل لحماية الأرقام؛ توجد وصفات توكيو بلا بيانات في ابديت تكرار: ' +
            ', '.join(match_report['unmatched'])
        )

    wb = load_workbook(updated_xlsm, keep_vba=True, data_only=False)
    wb['All_Ingredients']['R1'] = day_no
    wb['Marination_Ordering']['R1'] = day_no
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = 'auto'
    wb.save(updated_xlsm)
    wb.close()

    calculation_source = _make_libreoffice_copy(updated_xlsm, root / 'calculation-source.xlsx')
    recalculated_dir = root / 'recalculated'
    recalculated = _run_soffice(calculation_source, recalculated_dir, 'xlsx')
    values = load_workbook(recalculated, data_only=False, read_only=False)
    hot_sheets = _day_sheets(values, 'All_Ingredients', day_no)
    marination_sheets = _day_sheets(values, 'Marination_Ordering', day_no)
    values.close()

    blocks = []
    counts = {}
    for section in ('batch', 'special', 'actuals', 'garnish'):
        block_book = root / f'{section}.xlsx'
        block_book, page_count = _make_block_workbook(
            recalculated, recalculated, block_book, day_no, section, hot_sheets
        )
        block_pdf = _run_soffice(block_book, root / 'pdf', 'pdf')
        blocks.append(block_pdf)
        counts[section] = page_count

    mar_book = root / 'marination.xlsx'
    mar_book, mar_count = _make_block_workbook(
        recalculated, recalculated, mar_book, day_no, 'marination', marination_sheets
    )
    mar_pdf_part = _run_soffice(mar_book, root / 'pdf', 'pdf')

    day_name = DAY_NAMES.get(day_no, f'Day {day_no}')
    hot_pdf = _merge_pdfs(blocks, root / f'Tokyo_Hot_Section_Day{day_no}.pdf')
    mar_pdf = _merge_pdfs([mar_pdf_part], root / f'Tokyo_Marination_Day{day_no}.pdf')
    final_xlsm = root / f'Tokyo_Ordering_Updated_Day{day_no}.xlsm'
    shutil.copyfile(updated_xlsm, final_xlsm)

    zip_path = root / f'Tokyo_Production_{day_name}_Day{day_no}.zip'
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(hot_pdf, hot_pdf.name)
        archive.write(mar_pdf, mar_pdf.name)
        archive.write(final_xlsm, final_xlsm.name)

    report = {
        **match_report,
        'input': input_report,
        'hot_sheets': len(hot_sheets),
        'marination_sheets': len(marination_sheets),
        'pages': {**counts, 'hot_total': sum(counts.values()), 'marination': mar_count},
        'files': [hot_pdf.name, mar_pdf.name, final_xlsm.name],
    }
    return str(zip_path), str(final_xlsm), report
