"""Local Excel -> LibreOffice -> Poppler -> branded PDF, with no UI or network."""
from io import BytesIO
from pathlib import Path
import shutil
import subprocess
import tempfile
import time
from lxml import etree as ET
from PIL import Image
from pypdf import PdfReader, PdfWriter
from .xlsx_package import Package, S
from .branding import make_intro_page, make_sheet_page, crop_white_space, transparent_circle_logo

DEFAULT_LOGO = Path(__file__).parent / 'assets/default_logo.png'


def dependencies():
    executables = {'soffice': shutil.which('soffice') or shutil.which('libreoffice'),
                   'pdftoppm': shutil.which('pdftoppm')}
    missing = [name for name, path in executables.items() if not path]
    if missing:
        raise RuntimeError('Missing renderer executables: ' + ', '.join(missing))
    return executables


def prepare_for_print(source, destination, pages, print_area='B1:T39'):
    package = Package(source)
    if not pages or len(set(pages)) != len(pages) or any(p not in package.sheets for p in pages):
        raise ValueError('Invalid or duplicate PDF page selection')
    workbook = package.xml('xl/workbook.xml')
    sheets = workbook.find(f'{{{S}}}sheets')
    elements = {s.get('name'): s for s in sheets}
    order = list(pages) + [s for s in elements if s not in pages]
    for s in list(sheets):
        sheets.remove(s)
    names = workbook.find(f'{{{S}}}definedNames')
    if names is None:
        names = ET.Element(f'{{{S}}}definedNames')
        workbook.insert(workbook.index(sheets) + 1, names)
    for node in list(names):
        if node.get('name') == '_xlnm.Print_Area':
            names.remove(node)
    for i, name in enumerate(order):
        element = elements[name]
        element.set('state', 'visible' if name in pages else 'hidden')
        sheets.append(element)
        if name not in pages:
            continue
        ET.SubElement(names, f'{{{S}}}definedName', name='_xlnm.Print_Area', localSheetId=str(i)).text = (
            "'" + name.replace("'", "''") + "'!" + print_area)
        root = package.xml(package.sheets[name])
        prop = root.find(f'{{{S}}}sheetPr')
        if prop is None:
            prop = ET.Element(f'{{{S}}}sheetPr')
            root.insert(0, prop)
        fit = prop.find(f'{{{S}}}pageSetUpPr')
        if fit is None:
            fit = ET.SubElement(prop, f'{{{S}}}pageSetUpPr')
        fit.set('fitToPage', '1')
        setup = root.find(f'{{{S}}}pageSetup')
        if setup is None:
            setup = ET.Element(f'{{{S}}}pageSetup')
            before = next((x for x in root if ET.QName(x).localname in (
                'headerFooter', 'rowBreaks', 'colBreaks', 'drawing', 'legacyDrawing', 'tableParts', 'extLst')), None)
            root.insert(root.index(before) if before is not None else len(root), setup)
        setup.set('orientation', 'landscape')
        setup.set('paperSize', '9')
        setup.set('fitToWidth', '1')
        setup.set('fitToHeight', '1')
        setup.attrib.pop('scale', None)
    views = workbook.find(f'{{{S}}}bookViews')
    if views is not None:
        for view in views:
            view.set('activeTab', '0')
    package.save(destination)


def render_pdf(source, output, *, pages, year, report_date, cover=True, dpi=190,
               timeout=300, print_area='B1:T39', max_pdf_bytes=18000000):
    binaries = dependencies()
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    start = time.monotonic()
    with tempfile.TemporaryDirectory(prefix='bcc-render-') as folder:
        work = Path(folder)
        prepared = work / 'dashboard.xlsx'
        prepare_for_print(source, prepared, pages, print_area)
        converted = work / 'converted'
        converted.mkdir()
        profile = (work / 'office-profile').as_uri()
        result = subprocess.run([binaries['soffice'], f'-env:UserInstallation={profile}', '--headless',
                                 '--convert-to', 'pdf:calc_pdf_Export', '--outdir', str(converted), str(prepared)],
                                capture_output=True, text=True, timeout=timeout, check=False)
        raw_pdf = converted / 'dashboard.pdf'
        if result.returncode or not raw_pdf.is_file():
            raise RuntimeError(f'LibreOffice conversion failed: {(result.stderr + result.stdout)[-3000:]}')
        actual_pages = len(PdfReader(raw_pdf).pages)
        if actual_pages != len(pages):
            raise RuntimeError(f'Expected {len(pages)} worksheet pages, got {actual_pages}; refusing incorrect labels')
        writer = PdfWriter()
        def append_image(img):
            with BytesIO() as buffer:
                img.convert('RGB').save(buffer, 'PDF', resolution=180.0)
                buffer.seek(0)
                writer.add_page(PdfReader(buffer).pages[0])
            img.close()
        logo = work / 'logo.png'
        transparent_circle_logo(DEFAULT_LOGO, logo)
        if cover:
            append_image(make_intro_page(f'Class of {year} Full-Time Placement', logo,
                                        f'As of {report_date} | BYU Business Career Center'))
        for i, page in enumerate(pages, 1):
            prefix = work / 'page'
            subprocess.run([binaries['pdftoppm'], '-f', str(i), '-l', str(i), '-singlefile',
                            '-r', str(dpi), '-png', str(raw_pdf), str(prefix)],
                           capture_output=True, text=True, timeout=timeout, check=True)
            shot = prefix.with_suffix('.png')
            crop_white_space(shot, shot)
            append_image(make_sheet_page(shot, page, logo, True))
        temporary = output.with_suffix('.pdf.tmp')
        try:
            with temporary.open('wb') as stream:
                writer.write(stream)
            if len(PdfReader(temporary).pages) != len(pages) + int(cover):
                raise RuntimeError('Final PDF page count mismatch')
            if not 1000 < temporary.stat().st_size <= max_pdf_bytes:
                raise RuntimeError('PDF is empty or exceeds the configured size limit')
            temporary.replace(output)
        finally:
            temporary.unlink(missing_ok=True)
    return {'pages': len(pages) + int(cover), 'bytes': output.stat().st_size,
            'elapsed_seconds': round(time.monotonic() - start, 2), 'sheet_order': list(pages)}

