"""Populate the supplied chart template from a finished report; never query SQL."""
from datetime import date, datetime
from pathlib import Path
import calendar
import hashlib
import json
import math
import re
from lxml import etree as ET
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.datetime import from_excel
from .xlsx_package import Package, S, C, A, NS

PROGRAMS = ('BSAcc', 'BSEDM', 'BSEnt', 'BSFin', 'BSGSCM', 'BSHRM', 'BSIS',
            'BSBusM', 'BSMktg', 'BSStrat', 'MAcc', 'MBA', 'MISM', 'MPA')
OVERALL = 'MSB Full-Time Dashboard'
PAGES = (OVERALL,) + PROGRAMS
NITTY = 'NittyGrittySheet'
STATUSES = ('Accepted an offer', 'Actively seeking', 'No Recent Information Available',
            'Not Reported', 'Not seeking (continuing education)',
            'Not seeking (full-time homemaker/stay-at-home parent)',
            'Not seeking (other reasons)', 'Not seeking (postponing job search for a specific reason)',
            'Not seeking (returning to a previous company)',
            'Not seeking (starting a new business as owner)', 'Class Size')
TEMPLATE = Path(__file__).parent / 'templates/placement_dashboard_template.xlsx'
TEMPLATE_SHA256 = '5441e0b2b4511884168cc34ef82ec46bcf66c1cb419ce03c7f60e7b6b940ca1a'


class ContractError(ValueError):
    pass


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def parse_date(value, epoch=None):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 20000 < value < 100000:
        return (from_excel(value, epoch=epoch) if epoch else from_excel(value)).date()
    if isinstance(value, str):
        value = re.sub(r'_\d+$', '', value.strip())
        for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%Y-%m-%d %H:%M:%S'):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def monthly_row(cohort_year, observation):
    offset = (observation.year - (int(cohort_year) - 1)) * 12 + observation.month - 8
    return 25 + offset if 0 <= offset < 14 else None


def table_name(page, leadership):
    if page == OVERALL:
        return 'FT_total_wh' if leadership else 'Class2'
    return page + ('_2' if page in ('MBA', 'MPA') else '2')


def read_history(source, pages, as_of, start_date, diagnostics=None, require_current=True):
    wb = load_workbook(source, data_only=True)
    try:
        lookup = {name: (ws, ws.tables[name]) for ws in wb for name in ws.tables}
        leadership = 'FT_total_wh' in lookup
        history = {}
        for page in pages:
            name = table_name(page, leadership)
            if name not in lookup:
                raise ContractError(f'{source}: missing named table {name}')
            ws, table = lookup[name]
            left, top, right, bottom = range_boundaries(table.ref)
            labels = {}
            for row in range(top + 1, bottom + 1):
                label = str(ws.cell(row, left).value or '').strip()
                if label in labels:
                    raise ContractError(f'{name}: duplicate status {label}')
                labels[label] = row
            if 'Class Size' not in labels and 'Total' in labels:
                labels['Class Size'] = labels['Total']
            missing = set(STATUSES) - set(labels)
            pct_label = next((k for k in labels if k.lower() in ('% placed', 'placement %')), None)
            if missing or not pct_label:
                raise ContractError(f'{name}: missing statuses/rate: {sorted(missing)}')
            values = {}
            for col in range(left + 1, right + 1):
                header = ws.cell(top, col).value
                when = parse_date(header, wb.epoch)
                if when is None:
                    if re.fullmatch(r'Column\d+', str(header)) and all(
                            ws.cell(r, col).value in (None, '', '-', 0) for r in range(top + 1, bottom + 1)):
                        continue
                    if diagnostics is not None:
                        diagnostics.append({'source': Path(source).name, 'table': name, 'cell': ws.cell(top, col).coordinate,
                                            'header': str(header), 'action': 'ignored unrecognized historical date'})
                    continue
                if not start_date <= when <= as_of:
                    continue
                counts = []
                for status in STATUSES:
                    value = ws.cell(labels[status], col).value
                    if value == '-':
                        value = 0
                    if not isinstance(value, (int, float)) or not math.isfinite(value) or value < 0 or int(value) != value:
                        raise ContractError(f'{name}: invalid count for {status} on {when}: {value!r}')
                    counts.append(int(value))
                rate = ws.cell(labels[pct_label], col).value
                if not isinstance(rate, (int, float)) or not math.isfinite(rate) or not 0 <= rate <= 1:
                    raise ContractError(f'{name}: invalid rate on {when}: {rate!r}')
                if when == as_of:
                    denominator = counts[0] + counts[1] + counts[3]
                    calculated = counts[0] / denominator if denominator else 0
                    if abs(calculated - rate) > 0.00005001:
                        raise ContractError(f'{name}: current rate does not reconcile to accepted/seeking/not reported')
                observation = {'counts': counts, 'rate': rate}
                if when in values and values[when] != observation:
                    raise ContractError(f'{name}: conflicting duplicate source observations on {when}')
                values[when] = observation
            if require_current and as_of not in values:
                raise ContractError(f'{name}: report has no observation for {as_of}; refusing a stale visualization')
            history[page] = values
        return history
    finally:
        wb.close()


def refresh_caches(package, last_column, year):
    end = get_column_letter(last_column)
    for part in package.chart_parts():
        root = package.xml(part)
        if any(f.text and f.text.startswith(NITTY + '!') for f in root.findall('.//c:f', NS)):
            for axis in root.findall('.//c:catAx', NS):
                for tag in ('tickLblSkip', 'tickMarkSkip'):
                    tick = axis.find(f'{{{C}}}{tag}')
                    if tick is None:
                        tick = ET.Element(f'{{{C}}}{tag}')
                        before = next((n for n in axis if ET.QName(n).localname in (
                            ('tickMarkSkip', 'noMultiLvlLbl', 'extLst') if tag == 'tickLblSkip' else ('noMultiLvlLbl', 'extLst'))), None)
                        axis.insert(axis.index(before) if before is not None else len(axis), tick)
                    tick.set('val', str(max(1, math.ceil((last_column - 1) / 8))))
        for txt in root.findall('.//c:title//a:t', NS):
            if txt.text and 'Weekly' in txt.text:
                txt.text = re.sub(r'20\d{2}', str(year), txt.text)
            elif txt.text and '5 Year Summary' in txt.text:
                txt.text = 'Current class + five prior classes'
        for series in root.findall('.//c:ser', NS):
            values = series.find('c:val/c:numRef/c:f', NS)
            if values is not None and values.text.endswith('!$N$25:$N$38'):
                marker = series.find(f'{{{C}}}marker')
                if marker is None:
                    marker = ET.Element(f'{{{C}}}marker')
                    before = series.find(f'{{{C}}}cat')
                    series.insert(series.index(before) if before is not None else len(series), marker)
                for child in list(marker):
                    marker.remove(child)
                ET.SubElement(marker, f'{{{C}}}symbol', val='circle')
                ET.SubElement(marker, f'{{{C}}}size', val='5')
        # Explicitly include data on the helper sheet, which is excluded from printing.
        visible = root.find('.//c:plotVisOnly', NS)
        if visible is not None:
            visible.set('val', '0')
        for reference in root.findall('.//c:numRef', NS) + root.findall('.//c:strRef', NS):
            formula = reference.find(f'{{{C}}}f')
            if formula is None or not formula.text:
                continue
            if formula.text.startswith(NITTY + '!') and ':' in formula.text:
                formula.text = re.sub(r':\$[A-Z]+\$(\d+)', lambda m: f':${end}${m[1]}', formula.text)
            sheet, address = formula.text.rsplit('!', 1)
            sheet = sheet.strip("'").replace("''", "'")
            left, top, right, bottom = range_boundaries(address)
            values = [package.get(sheet, f'{get_column_letter(c)}{r}')
                      for r in range(top, bottom + 1) for c in range(left, right + 1)]
            numeric = reference.tag == f'{{{C}}}numRef'
            tag = 'numCache' if numeric else 'strCache'
            old = reference.find(f'{{{C}}}{tag}')
            fmt = old.find(f'{{{C}}}formatCode').text if old is not None and old.find(f'{{{C}}}formatCode') is not None else 'General'
            if old is not None:
                reference.remove(old)
            cache = ET.SubElement(reference, f'{{{C}}}{tag}')
            if numeric:
                ET.SubElement(cache, f'{{{C}}}formatCode').text = fmt
            ET.SubElement(cache, f'{{{C}}}ptCount', val=str(len(values)))
            for idx, value in enumerate(values):
                if value is None:
                    continue
                if numeric and value == '-':
                    value = 0
                if numeric and not isinstance(value, (int, float)):
                    raise ContractError(f'Non-numeric chart data: {formula.text}: {value!r}')
                if not numeric and isinstance(value, float) and value.is_integer():
                    value = int(value)
                point = ET.SubElement(cache, f'{{{C}}}pt', idx=str(idx))
                ET.SubElement(point, f'{{{C}}}v').text = str(value)


def update_dashboard(source, output, *, cohort_year, as_of, start_date, programs=None, base=None,
                     historical_sources=None):
    source, output = Path(source), Path(output)
    base = Path(base) if base else TEMPLATE
    if base.resolve() == TEMPLATE.resolve() and digest(base) != TEMPLATE_SHA256:
        raise ContractError('The versioned dashboard template changed; revalidate its mapping and fingerprint')
    if output.resolve() in (source.resolve(), base.resolve(), TEMPLATE.resolve()):
        raise ContractError('Dashboard output must be a separate staged file')
    pages = PAGES if programs is None else (OVERALL,) + tuple(programs)
    if len(set(pages)) != len(pages) or any(p not in PAGES for p in pages):
        raise ContractError('Unsupported or duplicate dashboard pages')
    diagnostics = []
    history = read_history(source, pages, as_of, start_date, diagnostics)
    package = Package(base)
    reference_styles = Package(TEMPLATE)
    if tuple(package.sheets) != (OVERALL, NITTY) + PROGRAMS or len(package.chart_parts()) != 45:
        raise ContractError('Dashboard template sheet/chart contract changed')
    old_year = int(package.get(OVERALL, 'N24'))
    year = int(cohort_year)
    if year not in (old_year, old_year + 1):
        raise ContractError('Rollover supports the same year or one consecutive cohort only')
    metadata_path = base.with_suffix('.json')
    previous = json.loads(metadata_path.read_text()) if metadata_path.exists() else {}
    if previous and previous.get('dashboard_sha256') != digest(base):
        raise ContractError('Dashboard state and metadata hashes disagree')
    if previous and previous.get('report_date', '') > as_of.isoformat():
        raise ContractError('Refusing to move persistent dashboard state backwards')
    metadata = {'schema_version': 1, 'cohort': str(year), 'report_date': as_of.isoformat(),
                'source_sha256': digest(source), 'template_sha256': digest(base),
                'pages': list(pages), 'source_warnings': diagnostics,
                'monthly': previous.get('monthly', {}) if year == old_year else {},
                'historical_note': 'Prior-year values are recorded snapshots; June template has incomplete 2026 history.'}
    existing = {}
    for c in range(2, 16385):
        when = parse_date(package.get(NITTY, f'{get_column_letter(c)}23'))
        if when is None:
            break
        if when in existing:
            raise ContractError('Duplicate dashboard weekly date')
        existing[when] = c
    dates = sorted(d for d in existing if d <= as_of) if year == old_year else []
    boundary = max(dates) if dates else start_date
    additions = {d for records in history.values() for d in records
                 if d.weekday() == 4 and (not dates or d > boundary)}
    dates = sorted(set(dates) | additions)
    if not dates:
        raise ContractError('No eligible weekly observations available for this dashboard')
    for pidx, page in enumerate(PAGES):
        rate_row, status_header = 24 + 17 * pidx, 26 + 17 * pidx
        destination = {package.get(NITTY, f'A{row}'): row for row in range(status_header + 1, status_header + 12)}
        if set(destination) != set(STATUSES):
            raise ContractError(f'Dashboard status labels changed: {page}')
        target_rows = [rate_row] + [destination[label] for label in STATUSES]
        # Capture old observations before clearing/rebuilding the affected ranges.
        retained = {d: [package.get(NITTY, f'{get_column_letter(c)}{r}')
                       for r in target_rows]
                    for d, c in existing.items()}
        for r in [rate_row - 1, rate_row, status_header] + list(range(status_header + 1, status_header + 12)):
            for c in range(2, max(len(existing), len(dates)) + 2):
                package.set(NITTY, f'{get_column_letter(c)}{r}', None, f'B{r}')
        for c, when in enumerate(dates, 2):
            letter = get_column_letter(c)
            for row in (rate_row - 1, status_header):
                package.set(NITTY, f'{letter}{row}', when.strftime('%m/%d/%Y'), f'B{row}')
            observation = history.get(page, {}).get(when)
            data = ([observation['rate']] + observation['counts']) if observation else (
                retained.get(when) if year == old_year else None)
            if data is None and page in pages:
                raise ContractError(f'{page}: no source observation for {when}')
            if data:
                for row, value in zip(target_rows, data):
                    package.set(NITTY, f'{letter}{row}', value, f'B{row}')
        if year != old_year:
            for col in range(19, 14, -1):
                for row in range(24, 39):
                    package.set(page, f'{get_column_letter(col)}{row}', package.get(page, f'{get_column_letter(col - 1)}{row}'))
            for row in range(25, 39):
                package.set(page, f'N{row}', None)
        package.set(page, 'N24', year)
        for table in package.tables(page):
            cols = table.find(f'{{{S}}}tableColumns')
            for col, node in enumerate(cols, 13):
                value = package.get(page, f'{get_column_letter(col)}24')
                node.set('name', str(int(value)) if isinstance(value, float) else str(value))
        package.set(page, 'B1', f'Weekly: {"MSB (All Programs)" if page == OVERALL else page} Class of {year} Placement')
        package.set(page, 'M23', f'Current month: {as_of:%m/%d/%Y}; history: snapshots.')
        for row in range(25, 39):
            style = reference_styles.cell(page, 'N35' if row == monthly_row(year, as_of) else 'N36').get('s')
            if style:
                package.cell(page, f'N{row}', True).set('s', style)
        if page not in history:
            continue
        monthly = {}
        for when in sorted(history[page]):
            row = monthly_row(year, when)
            if row is not None:
                monthly[row] = when
        page_meta = metadata['monthly'].setdefault(page, {})
        for row, when in monthly.items():
            current = when.year == as_of.year and when.month == as_of.month
            final = when.day == calendar.monthrange(when.year, when.month)[1]
            old_meta = page_meta.get(str(row), {})
            if current or package.get(page, f'N{row}') is None or (final and old_meta and old_meta.get('kind') != 'month_end'):
                package.set(page, f'N{row}', history[page][when]['rate'])
                page_meta[str(row)] = {'date': when.isoformat(), 'kind': 'month_end' if final else 'last_available',
                                       'rate': history[page][when]['rate']}
    # Refresh available prior-cohort comparisons from actual archived reports. Do not
    # infer September finals or use observations later than this run's as-of date.
    metadata['historical_sources'] = {}
    for prior_year, prior_source in (historical_sources or {}).items():
        prior_year = int(prior_year)
        if not year - 5 <= prior_year < year:
            continue
        prior_source = Path(prior_source)
        if not prior_source.is_file():
            metadata['source_warnings'].append({'source': str(prior_source),
                'action': 'prior-year report unavailable; retained supplied template observations'})
            continue
        histories = read_history(prior_source, pages, as_of, date(prior_year - 1, 1, 1),
                                 metadata['source_warnings'], require_current=False)
        provenance = {'sha256': digest(prior_source), 'monthly': {}}
        for page, records in histories.items():
            column = next((get_column_letter(col) for col in range(15, 20)
                           if int(package.get(page, f'{get_column_letter(col)}24')) == prior_year), None)
            if column is None:
                continue
            selected = {}
            for when in sorted(records):
                row = monthly_row(prior_year, when)
                if row is not None:
                    selected[row] = when
            provenance['monthly'][page] = {}
            for row, when in selected.items():
                rate = records[when]['rate']
                package.set(page, f'{column}{row}', rate)
                provenance['monthly'][page][str(row)] = {'date': when.isoformat(), 'rate': rate,
                    'kind': 'month_end' if when.day == calendar.monthrange(when.year, when.month)[1] else 'last_available'}
        metadata['historical_sources'][str(prior_year)] = provenance
    package.set(NITTY, 'A22', f'Class of {year} full-time placement')
    for table in package.tables(NITTY):
        _, top, _, _ = range_boundaries(table.get('ref'))
        if top >= 23:
            package.resize_table(NITTY, table, len(dates) + 1)
    package.xml(package.sheets[NITTY]).find(f'{{{S}}}dimension').set('ref', f'A1:{get_column_letter(len(dates)+1)}275')
    refresh_caches(package, len(dates) + 1, year)
    package.remove_calc_chain()
    package.save(output)
    metadata.update({'weekly_dates': [d.isoformat() for d in dates], 'dashboard_sha256': digest(output)})
    output.with_suffix('.json').write_text(json.dumps(metadata, indent=2) + '\n')
    return metadata
