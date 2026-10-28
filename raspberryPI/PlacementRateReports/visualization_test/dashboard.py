"""Academic-year dashboard on top of the unchanged Visualization.update_dashboard.

Table columns are always N..S = current class, then the five prior classes. The
current class is the report class (core.academic_year: October 2026-September 2027
-> 2027), so the table rolls automatically each October: the oldest class drops off
S and every other class moves one column right. The new class's August/September
rows come from its own history, which starts in the last week of June. Monthly
values persist per audience in a JSON state file, and the previous class's column
is refreshed from that class's final report, so a finished class keeps its full
row of months (through September) after rolling.
"""
from datetime import date
from pathlib import Path
import json
import tempfile
from openpyxl.utils import get_column_letter
from core.academic_year import report_class_year
from core.report_artifacts import sha256, write_json
from Visualization import update_dashboard as updater
from Visualization.update_dashboard import (update_dashboard, read_history, refresh_caches,
    monthly_row, Package, PAGES, OVERALL, NITTY, NS, S)

YEAR_COLUMNS = [get_column_letter(c) for c in range(14, 20)]  # N..S
MONTH_ROWS = range(25, 39)                                    # Aug Y-1 .. Sep Y


def academic_class_year(day):
    return report_class_year(day)


def table_years(class_year):
    return [int(class_year) - i for i in range(len(YEAR_COLUMNS))]


def template_seed(template):
    """Recorded history shipped in the template: {page: {year: {row: {rate, date}}}}."""
    package = Package(template)
    seed = {}
    for page in PAGES:
        for letter in YEAR_COLUMNS:
            header = package.get(page, f'{letter}24')
            if header is None:
                continue
            rows = seed.setdefault(page, {}).setdefault(str(int(float(header))), {})
            for row in MONTH_ROWS:
                value = package.get(page, f'{letter}{row}')
                if isinstance(value, float):
                    rows[str(row)] = {'rate': value, 'date': None}
    return seed


def monthly_values(history, class_year):
    """Latest observation in each month -> that month's row (same rule as update_dashboard)."""
    values = {}
    for page, records in history.items():
        for when in sorted(records):
            row = monthly_row(class_year, when)
            if row is not None:
                values.setdefault(page, {})[str(row)] = {'rate': records[when]['rate'],
                                                         'date': when.isoformat()}
    return values


def merge(state, live, year):
    for page, rows in live.items():
        stored = state.setdefault(page, {}).setdefault(str(year), {})
        for row, entry in rows.items():
            old = stored.get(row)
            if old is None or not old.get('date') or entry['date'] >= old['date']:
                stored[row] = entry


def table_value(year, page, row, live, state, seed, as_of):
    """Live report value, else saved state (never later than the report date), else template."""
    entry = live.get(year, {}).get(page, {}).get(str(row))
    if entry is None:
        entry = state.get(page, {}).get(str(year), {}).get(str(row))
        if entry is not None and entry.get('date') and entry['date'] > as_of.isoformat():
            entry = None
    if entry is None:
        entry = seed.get(page, {}).get(str(year), {}).get(str(row))
    return entry['rate'] if entry else None


def write_table(package, pages, class_year, live, state, seed, as_of):
    """N..S = class_year, class_year-1, ... (header row 24 on every tab, month rows 25-38 on pages)."""
    for page in PAGES:
        for letter, year in zip(YEAR_COLUMNS, table_years(class_year)):
            # Table headers must be text: the updater's numeric N24 makes Excel "repair" the table.
            package.set(page, f'{letter}24', str(year))
    for page in pages:
        for letter, year in zip(YEAR_COLUMNS, table_years(class_year)):
            for row in MONTH_ROWS:
                package.set(page, f'{letter}{row}', table_value(year, page, row, live, state, seed, as_of))


def set_weekly_titles(package, class_year):
    """'Weekly 2026 Placement Trend' / 'Weekly Placement % Trend' -> 'Weekly Class of 2027 Placement Trend'."""
    for part in package.chart_parts():
        for title in package.xml(part).findall('.//c:title', NS):
            texts = title.findall('.//a:t', NS)
            full = ''.join(t.text or '' for t in texts).strip()
            if not texts or not full.lower().startswith('weekly'):
                continue
            label = 'Search Status' if 'search' in full.lower() else 'Placement Trend'
            texts[0].text = f'Weekly Class of {class_year} {label}'
            for run in title.findall('.//a:r', NS)[1:]:
                run.getparent().remove(run)


def show_dashboards_only(package, pages):
    """Open on the MSB overall tab; show only this audience's dashboards.

    The chart-data helper and other programs' (not updated) tabs are hidden, not
    deleted, because the charts and template contract still reference them.
    """
    workbook = package.xml('xl/workbook.xml')
    for sheet in workbook.find(f'{{{S}}}sheets'):
        if sheet.get('name') in pages:
            sheet.attrib.pop('state', None)
        else:
            sheet.set('state', 'hidden')
    views = workbook.find(f'{{{S}}}bookViews')
    for view in (views if views is not None else []):
        view.set('activeTab', '0')
        view.set('firstSheet', '0')
    for name in package.sheets:
        for view in package.xml(package.sheets[name]).iter(f'{{{S}}}sheetView'):
            view.set('tabSelected', '1' if name == OVERALL else '0')


def build_dashboard(source, output, *, class_year, as_of, start_date, programs, template,
                    state_path, prior_source=None):
    """Returns update_dashboard metadata plus the table years and prior-class status."""
    class_year, output, state_path = int(class_year), Path(output), Path(state_path)
    pages = PAGES if programs is None else (OVERALL,) + tuple(programs)
    # Start from the template with N..R labelled as the five prior classes, so the
    # unchanged updater performs exactly one roll and puts the current class in N.
    with tempfile.TemporaryDirectory(prefix='bcc-vt-') as folder:
        base = Path(folder) / 'base.xlsx'
        package = Package(template)
        for page in PAGES:
            for letter, year in zip(YEAR_COLUMNS[:-1], table_years(class_year - 1)):
                package.set(page, f'{letter}24', year)
        package.save(base)
        # The updater copies current-month cell styles from its module TEMPLATE; point
        # that at the configured template so style indexes match this base.
        default_template = updater.TEMPLATE
        updater.TEMPLATE = Path(template)
        try:
            metadata = update_dashboard(source, output, cohort_year=class_year, as_of=as_of,
                                        start_date=start_date, programs=programs, base=base,
                                        historical_sources={})
        finally:
            updater.TEMPLATE = default_template

    live = {class_year: monthly_values(read_history(source, pages, as_of, start_date), class_year)}
    warnings = []
    if prior_source is not None:
        try:
            prior = read_history(prior_source, pages, as_of, date(class_year - 3, 1, 1),
                                 warnings, require_current=False)
            live[class_year - 1] = monthly_values(prior, class_year - 1)
        except Exception as exc:  # The prior class is supplementary; keep the current dashboard.
            warnings.append({'source': Path(prior_source).name, 'error': f'{type(exc).__name__}: {exc}'})
    state = json.loads(state_path.read_text()) if state_path.exists() else {}
    seed = template_seed(template)

    package = Package(output)
    write_table(package, pages, class_year, live, state, seed, as_of)
    overlap = bool(live.get(class_year - 1))
    refresh_caches(package, len(metadata['weekly_dates']) + 1, class_year)
    set_weekly_titles(package, class_year)
    show_dashboards_only(package, pages)
    package.save(output)

    for year, values in live.items():
        merge(state, values, year)
    write_json(state_path, state)
    metadata.update({'table_years': table_years(class_year), 'prior_class_updated': overlap,
                     'prior_source': str(prior_source) if prior_source else None,
                     'prior_warnings': warnings, 'dashboard_sha256': sha256(output)})
    write_json(output.with_suffix('.json'), metadata)
    return metadata
