"""Dashboard sheets -> cropped PNG screenshots (LibreOffice + Poppler), one per page."""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from pypdf import PdfReader
from Visualization.render_excel import prepare_for_print
from Visualization.branding import crop_white_space


def _windows_install(pattern):
    # Windows installers (local VS Code testing) do not put these on PATH.
    for base in (Path('C:/Program Files'), Path.home() / 'AppData/Local/Microsoft/WinGet/Packages'):
        found = sorted(base.glob(pattern)) if base.is_dir() else []
        if found:
            return str(found[-1])
    return None


def dependencies():
    windows = os.name == 'nt'
    found = {'soffice': shutil.which('soffice') or shutil.which('libreoffice') or (
                 _windows_install('LibreOffice/program/soffice.exe') if windows else None),
             'pdftoppm': shutil.which('pdftoppm') or (
                 _windows_install('oschwartz10612.Poppler*/poppler-*/Library/bin/pdftoppm.exe') if windows else None)}
    missing = [name for name, path in found.items() if not path]
    if missing:
        raise RuntimeError('Missing renderer executables: ' + ', '.join(missing))
    return found


def render_sheet_images(source, pages, folder, *, dpi=190, timeout=300, print_area='B1:T39'):
    binaries = dependencies()
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='bcc-vt-render-') as scratch:
        work = Path(scratch)
        prepared = work / 'dashboard.xlsx'
        prepare_for_print(source, prepared, pages, print_area)
        converted = work / 'converted'
        converted.mkdir()
        result = subprocess.run([binaries['soffice'], f'-env:UserInstallation={(work / "profile").as_uri()}',
                                 '--headless', '--convert-to', 'pdf:calc_pdf_Export',
                                 '--outdir', str(converted), str(prepared)],
                                capture_output=True, text=True, timeout=timeout, check=False)
        raw_pdf = converted / 'dashboard.pdf'
        if result.returncode or not raw_pdf.is_file():
            raise RuntimeError(f'LibreOffice conversion failed: {(result.stderr + result.stdout)[-3000:]}')
        if len(PdfReader(raw_pdf).pages) != len(pages):
            raise RuntimeError('Rendered page count does not match the dashboard pages')
        shots = []
        for i, page in enumerate(pages, 1):
            prefix = work / 'page'
            subprocess.run([binaries['pdftoppm'], '-f', str(i), '-l', str(i), '-singlefile',
                            '-r', str(dpi), '-png', str(raw_pdf), str(prefix)],
                           capture_output=True, text=True, timeout=timeout, check=True)
            shot = folder / f'{i:02d}-{page}.png'
            crop_white_space(prefix.with_suffix('.png'), shot)
            shots.append(shot)
    return shots
