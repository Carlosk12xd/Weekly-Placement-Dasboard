"""Create a new class's report workbook from the previous class's workbook.

Runs on the new class's first report (the last week of June). The layout, table names
and styles are kept. Every class year in sheet names and title text moves forward
("2027 MSB Overall" -> "2028 MSB Overall", "CLASS OF 2028 INTERNSHIP" -> "CLASS OF 2029
INTERNSHIP"). All counts, Weekly History columns, ZOHO_DATA rows and placement details
are cleared, so the new class's history starts with its own first Friday.
"""
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils import range_boundaries
from .cohort_manager import shift_years
from .history_window import resize_table

DETAIL_HEADER_ROWS = 3  # placement_details rewrites everything below its 3 header rows


def class_report_path(path, old_year, new_year):
    """'.../2027_Placement_Reports/2027-WeeklyPlacement-MBA.xlsx' -> the same path for 2028.
    Only the file and folder names change, never the directories above them."""
    path, years = Path(path), int(new_year) - int(old_year)
    return path.parent.parent / shift_years(path.parent.name, years) / shift_years(path.name, years)


def create_class_workbook(previous, output, *, old_year, new_year):
    years = int(new_year) - int(old_year)
    if years < 1:
        raise ValueError('A new class workbook must be for a later class')
    wb = load_workbook(previous)
    try:
        for ws in wb.worksheets:
            ws.title = shift_years(ws.title, years)
            if ws.title == 'ZOHO_DATA':
                if ws.max_row > 1:
                    ws.delete_rows(2, ws.max_row)
                continue
            areas = [range_boundaries(ws.tables[name].ref) for name in ws.tables]
            if not areas and ws.max_row > DETAIL_HEADER_ROWS:
                ws.delete_rows(DETAIL_HEADER_ROWS + 1, ws.max_row)
            for row in ws.iter_rows():
                for cell in row:
                    inside = any(l <= cell.column <= r and t <= cell.row <= b for l, t, r, b in areas)
                    if isinstance(cell.value, str) and not inside and not cell.value.startswith('='):
                        cell.value = shift_years(cell.value, years)
            for name in list(ws.tables):
                clear_status_table(ws, ws.tables[name])
        output = Path(output)
        output.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output)
    finally:
        wb.close()
    return output


def clear_status_table(ws, table):
    """Empty a Most Recent Friday / Weekly History table; other tables (Summary) are rebuilt by SQL."""
    left, top, right, bottom = range_boundaries(table.ref)
    if 'search status' not in str(ws.cell(top, left).value or '').lower():
        return
    if right - left == 1:  # Most Recent Friday: one data column, overwritten on the first run
        ws.cell(top, right).value = 'Column1'
        for r in range(top + 1, bottom + 1):
            ws.cell(r, right).value = None
        resize_table(ws, table, right)
        return
    for col in range(left + 1, right + 1):  # Weekly History: the first run appends the first date
        for r in range(top, bottom + 1):
            cell = ws.cell(r, col)
            cell.value, cell._style = None, StyleArray()
    resize_table(ws, table, left)
