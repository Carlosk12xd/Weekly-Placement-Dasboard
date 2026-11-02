"""Keep every Weekly History table inside its class's tracking window.

A class's history runs from the last week of June Y-1 to the last week of September Y
(core.academic_year). Columns dated outside that window, the later copies of a date
written twice on one day ("07/10/2026_2"), and empty ColumnN placeholders are removed
from the table; the remaining columns keep their values and styles in date order.
"""
from copy import copy
from datetime import date, datetime
import re
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.table import TableColumn
from .academic_year import in_window


def header_date(value):
    """Date of a history column header ('07/10/2026', '07/10/2026_2' or an Excel date); else None."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        match = re.fullmatch(r'(\d{1,2}/\d{1,2}/\d{4})(?:_\d+)?', value.strip())
        if match:
            try:
                return datetime.strptime(match[1], '%m/%d/%Y').date()
            except ValueError:
                return None
    return None


def table_columns(table):
    columns = getattr(table.tableColumns, 'tableColumn', None)
    return table.tableColumns if columns is None else columns


def prune_history(ws, table, class_year):
    """Remove out-of-window, duplicate and empty placeholder columns. Returns removed headers.

    Same-date copies keep the rightmost (last written) column. A non-empty column whose
    header is not a date is kept untouched; report_checks reports it.
    """
    left, top, right, bottom = range_boundaries(table.ref)
    keep, removed, by_date = [], [], {}
    for col in range(left + 1, right + 1):
        header = ws.cell(top, col).value
        when = header_date(header)
        if when is None:
            if all(ws.cell(r, col).value in (None, '', '-', 0) for r in range(top + 1, bottom + 1)):
                removed.append(str(header))
            else:
                keep.append(col)
        elif not in_window(class_year, when):
            removed.append(str(header))
        else:
            if when in by_date:
                removed.append(str(ws.cell(top, by_date[when]).value))
                keep.remove(by_date[when])
            by_date[when] = col
            keep.append(col)
    if not removed:
        return []
    saved = {(r, c): (ws.cell(r, c).value, copy(ws.cell(r, c)._style))
             for c in keep for r in range(top, bottom + 1)}
    for offset, col in enumerate(keep, 1):
        for r in range(top, bottom + 1):
            cell = ws.cell(r, left + offset)
            cell.value, cell._style = saved[(r, col)]
        when = header_date(ws.cell(top, left + offset).value)
        if when is not None:  # the surviving copy of a duplicated date drops its _2 suffix
            ws.cell(top, left + offset).value = when.strftime('%m/%d/%Y')
    for col in range(left + 1 + len(keep), right + 1):
        for r in range(top, bottom + 1):
            cell = ws.cell(r, col)
            cell.value, cell._style = None, StyleArray()
    resize_table(ws, table, left + len(keep))
    return removed


def resize_table(ws, table, last_col):
    """Set the table's right edge and rebuild its column metadata from the header cells."""
    left, top, _, bottom = range_boundaries(table.ref)
    old = list(table_columns(table))
    ref = f'{ws.cell(top, left).column_letter}{top}:{ws.cell(bottom, last_col).column_letter}{bottom}'
    table.ref = ref
    if getattr(table, 'autoFilter', None) is not None:
        table.autoFilter.ref = ref
    names, columns = set(), []
    for idx, col in enumerate(range(left, last_col + 1)):
        raw = ws.cell(top, col).value
        base = str(raw).strip() if raw not in (None, '') else f'Column{idx + 1}'
        name, k = base, 1
        while name in names:
            k += 1
            name = f'{base}_{k}'
        names.add(name)
        column = old[idx] if idx < len(old) else TableColumn(id=idx + 1, name=name)
        column.id, column.name = idx + 1, name
        columns.append(column)
    table.tableColumns = columns


def prune_zoho_rows(ws, class_year):
    """Drop ZOHO_DATA rows dated outside the window or copied from a duplicate column."""
    rows = [[cell.value for cell in row] for row in ws.iter_rows(min_row=2, max_col=5)]
    def kept(row):
        when = row[3]
        if isinstance(when, datetime):
            when = when.date()
        return isinstance(when, date) and in_window(class_year, when)
    remaining = [row for row in rows if any(v not in (None, '') for v in row) and kept(row)]
    dropped = sum(1 for row in rows if any(v not in (None, '') for v in row)) - len(remaining)
    if dropped:
        ws.delete_rows(2, ws.max_row)
        for row in remaining:
            ws.append(row)
            ws.cell(ws.max_row, 4).number_format = 'yyyy-mm-dd'
    return dropped


def prune_workbook(path, class_year):
    """Apply the window to every Weekly History table and ZOHO_DATA in a saved workbook
    (offline runs, which do not rebuild the report). Returns {table: removed}."""
    from openpyxl import load_workbook
    from .report_checks import is_history
    wb = load_workbook(path)
    try:
        pruned = {}
        for ws in wb.worksheets:
            if ws.title == 'ZOHO_DATA':
                dropped = prune_zoho_rows(ws, class_year)
                if dropped:
                    pruned[ws.title] = dropped
                continue
            for name in list(ws.tables):
                table = ws.tables[name]
                left, top, _, _ = range_boundaries(table.ref)
                if is_history(name) and 'search status' in str(ws.cell(top, left).value or '').lower():
                    removed = prune_history(ws, table, class_year)
                    if removed:
                        pruned[name] = removed
        if pruned:
            wb.save(path)
        return pruned
    finally:
        wb.close()
