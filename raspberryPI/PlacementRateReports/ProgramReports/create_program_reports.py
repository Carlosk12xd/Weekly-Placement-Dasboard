from core.database import connect_database, load_environment
#!/usr/bin/env python3

# --- minimal imports ---
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableColumn
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Alignment, Border, Side
from datetime import date
import re

from ProgramReports.cd_sql_queries import write_sql_bsfin_queries, write_sql_overall_queries, write_sql_program_queries

# =========================
# ENV + CONSTANTS
# =========================

BASE_DIR = Path(__file__).resolve().parent


# Fixed run date label (per your instruction)
RUN_DATE_LABEL = date.today().strftime("%m/%d/%Y")

# Percent inputs (must match SQL/Excel labels exactly)
STATUS_ACCEPTED = "Accepted an offer"     # from your working script
STATUS_SEEKING = "Actively seeking"       # from your working script
STATUS_NOT_REPORTED = "Not Reported"      # from your working script

RIGHT_ALIGN = Alignment(horizontal="right")
THIN_BORDER = Border(bottom=Side(style="thin", color="000000"))

IGNORE_LABELS = {"total", "class size", "% placed", "placement %"}  # convenience


# =========================
# SMALL HELPERS (simple + essential)
# =========================

def table_names(programs):
    """
    4 tables per sheet:
      1: MRF FT,  2: WH FT,  3: MRF INT,  4: WH INT
    Class sheet uses 'Class1'..'Class4'
    MBA/MPA use underscores; others do not.
    """
    tbl_nms = {"Class": ("Class1", "Class2", "Class3", "Class4")}
    for program in programs:
        if program in ("MPA", "MBA"):
            tbl_nms[program] = (f"{program}_1", f"{program}_2", f"{program}_3", f"{program}_4")
        elif program == "BSFin":
            tbl_nms[program] = (f"{program}1", f"{program}2", f"{program}3", f"{program}4", f"{program}5", f"{program}6")
        else:
            tbl_nms[program] = (f"{program}1", f"{program}2", f"{program}3", f"{program}4")
    return tbl_nms

def program_to_filename(programs):
    if not programs:
        raise RuntimeError("No programs provided.")
    return programs[0] if len(programs) == 1 else "-".join(programs)

def fetch_rows(cur, sql, params=()):
    cur.execute(sql, params)
    return list(cur.fetchall())

def get_table(ws: Worksheet, name: str) -> Table:
    # Using ws.tables dict is the stable pattern
    if name not in ws.tables:
        raise RuntimeError(f"Expected table '{name}' not found on sheet '{ws.title}'.")
    return ws.tables[name]

def table_bounds(ref: str):
    start, end = ref.split(":")
    cell_re = re.compile(r"([A-Z]+)(\d+)")
    c1 = cell_re.fullmatch(start).groups()
    c2 = cell_re.fullmatch(end).groups()
    min_col = column_index_from_string(c1[0]); min_row = int(c1[1])
    max_col = column_index_from_string(c2[0]); max_row = int(c2[1])
    return min_row, max_row, min_col, max_col

def expected_header_for_table(ws: Worksheet, tbl_name: str) -> str:
    """Choose FT vs INT header text based on sheet or table name."""
    title = (ws.title or "").lower()
    name  = (tbl_name or "").lower()
    if ("internship" in title) or ("_int" in name) or name.endswith("int1") or name.endswith("int2"):
        return "Internship Search Status"   # working script logic
    return "Job Search Status"              # working script logic

def detect_header_row(ws: Worksheet, min_row: int, max_row: int, min_col: int, expected_first_header: str) -> int:
    """Find header row inside the table by header text; fall back to min_row."""
    first_col_letter = get_column_letter(min_col)
    exp = expected_first_header.strip().lower()

    # exact match
    for r in range(min_row, max_row + 1):
        v = ws[f"{first_col_letter}{r}"].value
        if isinstance(v, str) and v.strip().lower() == exp:
            return r

    # loose match: "*search status*"
    pat = re.compile(r"search\s+status", re.IGNORECASE)
    for r in range(min_row, max_row + 1):
        v = ws[f"{first_col_letter}{r}"].value
        if isinstance(v, str) and pat.search(v):
            return r

    return min_row  # fallback

def ensure_header(ws: Worksheet, header_row: int, data_cols, header_label: str, force_append: bool=False):
    """
    For MRF: set the single data column header to run date.
    For WH: append a rightmost column with run date if force_append=True.
    """
    if force_append:
        new_col_idx = data_cols[-1] + 1
        ws.cell(row=header_row, column=new_col_idx, value=header_label)
        return data_cols + [new_col_idx]

    if len(data_cols) == 1:  # MRF
        ws.cell(row=header_row, column=data_cols[0], value=header_label)
        return data_cols

    # WH fallback: append
    new_col_idx = data_cols[-1] + 1
    ws.cell(row=header_row, column=new_col_idx, value=header_label)
    return data_cols + [new_col_idx]

def set_table_ref(ws: Worksheet, tbl: Table, min_row: int, max_row: int, min_col: int, max_col: int):
    """
    SAFE table-widening that preserves tableColumns + AutoFilter metadata.
    Prevents Excel "Removed Feature: Table/AutoFilter" messages.
    """
    new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    tbl.ref = new_ref
    if getattr(tbl, "autoFilter", None) is not None:
        tbl.autoFilter.ref = new_ref

    # Keep <tableColumn> list in sync with new width
    tc_container = getattr(tbl, "tableColumns", None)
    tc_list = getattr(tc_container, "tableColumn", None)
    if tc_list is None:
        tc_list = tc_container
    if tc_list is None:
        raise RuntimeError("Table has no tableColumns list; cannot adjust metadata.")

    width = max_col - min_col + 1
    meta_count = len(tc_list)

    # Trim extras (rare)
    if meta_count > width:
        del tc_list[width:]
        meta_count = width

    # Append missing entries, using header cell text
    existing_names = {tc.name for tc in tc_list}
    next_id = max((tc.id for tc in tc_list), default=0) + 1
    for offset in range(meta_count, width):
        c = min_col + offset
        raw = ws.cell(row=min_row, column=c).value
        base = (str(raw).strip() if raw not in (None, "") else f"Column{offset+1}")
        name, k = base, 1
        while name in existing_names:
            k += 1
            name = f"{base}_{k}"
        existing_names.add(name)
        tc_list.append(TableColumn(id=next_id, name=name))
        next_id += 1

    if len(tc_list) != width:
        raise RuntimeError(f"Table metadata columns={len(tc_list)} but width={width} for {tbl.ref}")

def relabel_total_row_to_class_size(ws: Worksheet, label_col, min_row, max_row):
    """Rename 'Total' → 'Class Size' if present."""
    for r in range(min_row, max_row + 1):
        v = ws.cell(row=r, column=label_col).value
        if isinstance(v, str) and v.strip().lower() == "total":
            ws.cell(row=r, column=label_col, value="Class Size")
            return r
    # if already 'Class Size', return that row
    for r in range(min_row, max_row + 1):
        v = ws.cell(row=r, column=label_col).value
        if isinstance(v, str) and v.strip() == "Class Size":
            return r
    # else, assume penultimate row is total
    return max_row - 1

def find_percent_row(ws: Worksheet, label_col, min_row, max_row):
    for r in range(min_row, max_row + 1):
        v = ws.cell(row=r, column=label_col).value
        if isinstance(v, str) and v.strip() == "% Placed":
            return r
    # else, assume last row
    return max_row

def write_dash(cell):
    cell.value = "-"
    cell.alignment = RIGHT_ALIGN

def to_int(v):
    try:
        if v in (None, "", "-"):
            return 0
        return int(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return 0

def placement_percent(accepted: int, seeking: int, not_reported: int) -> float:
    denom = (accepted or 0) + (seeking or 0) + (not_reported or 0)
    if denom <= 0:
        return 0.0
    return round((accepted or 0) * 100.0 / denom, 2)

def write_percent(cell, pct: float):
    cell.value = pct / 100.0
    cell.number_format = "0.00%"

def reset_workbook_view_to_a1(wb):
    """
    Reset each worksheet's visible selection so Excel opens at A1 instead of
    the last edited cell.
    """
    if wb.worksheets:
        wb.active = 0

    for ws in wb.worksheets:
        ws.sheet_view.topLeftCell = "A1"
        if ws.sheet_view.selection:
            ws.sheet_view.selection[0].activeCell = "A1"
            ws.sheet_view.selection[0].sqref = "A1"

# =========================
# TABLE UPDATERS (MRF/WH)
# =========================

def update_mrf_table(ws: Worksheet, tbl_name: str, sql_rows):
    """
    MRF: two columns total (Status | value). Overwrite the single data column with RUN_DATE_LABEL,
    fill values, then recompute Class Size and % Placed for that column.
    """
    tbl = get_table(ws, tbl_name)
    min_row, max_row, min_col, max_col = table_bounds(tbl.ref)

    header_expected = expected_header_for_table(ws, tbl_name)
    header_row = detect_header_row(ws, min_row, max_row, min_col, header_expected)

    label_col = min_col
    data_cols = list(range(min_col + 1, max_col + 1))
    if len(data_cols) != 1:
        raise RuntimeError(f"MRF table '{tbl_name}' should have exactly 1 data column; found {len(data_cols)}.")

    # header
    ensure_header(ws, header_row, data_cols, RUN_DATE_LABEL)
    # sync tableColumns name for that data column (keeps metadata tidy)
    tc_list = getattr(getattr(tbl, "tableColumns", None), "tableColumn", None) or getattr(tbl, "tableColumns", None)
    if tc_list:
        idx = data_cols[0] - min_col
        tc_list[idx].name = str(ws.cell(row=header_row, column=data_cols[0]).value or f"Column{idx+1}").strip()

    # map SQL to dict
    sql_map = {str(r[0]).strip(): int(r[1]) for r in sql_rows}

    # fill
    for r in range(header_row + 1, max_row + 1):
        label = ws.cell(row=r, column=label_col).value
        if not label:
            continue
        s = str(label).strip()
        if s.lower() in IGNORE_LABELS:
            continue
        if s in sql_map:
            ws.cell(row=r, column=data_cols[0], value=int(sql_map[s]))
        else:
            write_dash(ws.cell(row=r, column=data_cols[0]))

    # totals + % placed for this column
    compute_totals_and_percent(ws, min_row, max_row, min_col, data_cols[-1])

def update_wh_table(ws: Worksheet, tbl_name: str, sql_rows):
    """
    WH: append a new column at the right, label it RUN_DATE_LABEL, fill,
    draw a thin line above Class Size, and compute totals/% placed for that new column.
    """
    tbl = get_table(ws, tbl_name)
    min_row, max_row, min_col, max_col = table_bounds(tbl.ref)

    header_expected = expected_header_for_table(ws, tbl_name)
    header_row = detect_header_row(ws, min_row, max_row, min_col, header_expected)

    label_col = min_col
    existing_data_cols = list(range(min_col + 1, max_col + 1))

    # add header at right
    if not existing_data_cols:
        newest_col = min_col + 1
        ws.cell(row=header_row, column=newest_col, value=RUN_DATE_LABEL)
    elif str(ws.cell(row=header_row, column=max_col).value or "").strip() == RUN_DATE_LABEL:
        # Today's column already exists (report re-run on the same day): reuse it
        # instead of appending a duplicate. Appending creates duplicate header cells,
        # which Excel flags as corrupt ("we found a problem with some content").
        newest_col = max_col
    else:
        new_cols = ensure_header(ws, header_row, existing_data_cols, RUN_DATE_LABEL, force_append=True)
        newest_col = new_cols[-1]

    # widen the table safely (preserves metadata)
    set_table_ref(ws, tbl, min_row, max_row, min_col, newest_col)

    # map results
    sql_map = {str(r[0]).strip(): int(r[1]) for r in sql_rows}

    # fill
    for r in range(header_row + 1, max_row + 1):
        label = ws.cell(row=r, column=label_col).value
        if not label:
            continue
        s = str(label).strip()
        if s.lower() in IGNORE_LABELS:
            continue
        if s in sql_map:
            ws.cell(row=r, column=newest_col, value=int(sql_map[s]))
        else:
            write_dash(ws.cell(row=r, column=newest_col))

    # thin border above Class Size for visual separation
    total_row = relabel_total_row_to_class_size(ws, label_col, min_row, max_row)
    ws.cell(row=total_row - 1, column=newest_col).border = THIN_BORDER

    # totals + % placed for newest column
    compute_totals_and_percent(ws, min_row, max_row, min_col, newest_col)

def compute_totals_and_percent(ws: Worksheet, min_row: int, max_row: int, min_col: int, col: int):
    """
    Total = sum of numeric rows (exclude 'Class Size' and '% Placed')
    % Placed = Accepted an offer / (Accepted an offer + Actively seeking + Not Reported)
    """
    label_col = min_col

    # locate special rows
    total_row = relabel_total_row_to_class_size(ws, label_col, min_row, max_row)
    pct_row   = find_percent_row(ws, label_col, min_row, max_row)

    # collect status rows
    status_to_row = {}
    for r in range(min_row + 1, max_row + 1):
        label = ws.cell(row=r, column=label_col).value
        if not isinstance(label, str):
            continue
        s = label.strip()
        if s.lower() in IGNORE_LABELS:
            continue
        status_to_row[s] = r

    # total
    total = 0
    for s, rr in status_to_row.items():
        total += to_int(ws.cell(row=rr, column=col).value)
    ws.cell(row=total_row, column=col, value=total)

    # percent placed
    acc = to_int(ws.cell(row=status_to_row.get(STATUS_ACCEPTED, 0), column=col).value if STATUS_ACCEPTED in status_to_row else 0)
    seek = to_int(ws.cell(row=status_to_row.get(STATUS_SEEKING, 0), column=col).value if STATUS_SEEKING in status_to_row else 0)
    nr  = to_int(ws.cell(row=status_to_row.get(STATUS_NOT_REPORTED, 0), column=col).value if STATUS_NOT_REPORTED in status_to_row else 0)

    pct = placement_percent(acc, seek, nr)
    write_percent(ws.cell(row=pct_row, column=col), pct)

# =========================
# SHEET UPDATER
# =========================

def update_sheet_with_ft_int(ws: Worksheet, table_tuple, ft_rows, int_rows):
    """
    Update 4 tables on a sheet:
      0: MRF FT (ft_rows)
      1: WH  FT (ft_rows)
      2: MRF INT (int_rows)
      3: WH  INT (int_rows)
    """
    t1, t2, t3, t4 = table_tuple
    update_mrf_table(ws, t1, ft_rows)
    update_wh_table(ws, t2, ft_rows)
    update_mrf_table(ws, t3, int_rows)
    update_wh_table(ws, t4, int_rows)

def update_bsfin_with_ft_int(ws: Worksheet, table_tuple, ft_rows, int_2027_rows, int_2028_rows):
    """
    Update 4 tables on a sheet:
      0: MRF FT (ft_rows)
      1: WH  FT (ft_rows)
      2: MRF INT (int_rows)
      3: WH  INT (int_rows)
    """
    t1, t2, t3, t4, t5, t6 = table_tuple
    update_mrf_table(ws, t1, ft_rows)
    update_wh_table(ws, t2, ft_rows)
    update_mrf_table(ws, t3, int_2027_rows)
    update_wh_table(ws, t4, int_2027_rows)
    update_mrf_table(ws, t5, int_2028_rows)
    update_wh_table(ws, t6, int_2028_rows)

# =========================
# MAIN
# =========================

def main(programs, cohort, output_path=None, report_date=None):
    global RUN_DATE_LABEL
    RUN_DATE_LABEL = (report_date or date.today()).strftime("%m/%d/%Y")
    # DB
    conn = connect_database(BASE_DIR)
    cur = conn.cursor()
    try:

        # variables
        c_year = cohort.id
        p_year1 = int(c_year) - 1
        p_year2 = int(c_year) - 2
        semester_byu = cohort.semester_byu
        int_yrs = cohort.internship_years

        # totals
        total_queries = write_sql_overall_queries(c_year, p_year1, p_year2, semester_byu, int_yrs)

        total_ft = total_queries[0]
        total_int = total_queries[1]

        total_ft_rows = fetch_rows(cur, total_ft)
        total_int_rows = fetch_rows(cur, total_int)

        byprog_ft = {}
        byprog_int = {}
        byprog_bsfin_int = {}

        for prog in programs:
            if prog == 'BSFin':
                prog_sqls = write_sql_bsfin_queries(c_year, p_year1, p_year2, semester_byu, prog)
            
                full_sql = prog_sqls[0]
                int_sql1 = prog_sqls[1]
                int_sql2 = prog_sqls[2]

                nxt_yr1 = int(c_year) + 1
                nxt_yr2 = int(c_year) + 2

                byprog_ft[prog] = fetch_rows(cur, full_sql)
                byprog_bsfin_int[nxt_yr1] = fetch_rows(cur, int_sql1)
                byprog_bsfin_int[nxt_yr2] = fetch_rows(cur, int_sql2)

            else:
                prog_sqls = write_sql_program_queries(c_year, p_year1, p_year2, semester_byu, int_yrs, prog)

                full_sql = prog_sqls[0]
                int_sql = prog_sqls[1]

                byprog_ft[prog] = fetch_rows(cur, full_sql)
                byprog_int[prog] = fetch_rows(cur, int_sql)

    finally:
        cur.close()
        conn.close()

    # workbook
    FILEPATH_TEMPLATE = str(output_path) if output_path else load_environment(BASE_DIR).get("OUTPUT_PATH", str(BASE_DIR / f"{c_year}_Placement_Reports" / (str(c_year) + "-WeeklyPlacement-{file_label}.xlsx")))
    fileLbl = program_to_filename(programs)
    wb_path = FILEPATH_TEMPLATE.format(file_label=fileLbl)
    wb = load_workbook(wb_path, data_only=False)

    # tables
    tbls = table_names(programs)

    # totals sheet (exact name confirmed earlier)
    class_ws_name = f"{c_year} MSB Overall"
    if class_ws_name not in wb.sheetnames:
        raise RuntimeError(f"Expected sheet '{class_ws_name}' not found.")
    class_ws = wb[class_ws_name]
    update_sheet_with_ft_int(class_ws, tbls["Class"], total_ft_rows, total_int_rows)

    # program sheets
    for program in programs:
        if program not in wb.sheetnames:
            raise RuntimeError(f"Expected program sheet '{program}' not found.")
        elif program == "BSFin":
            ws = wb[program]
            update_bsfin_with_ft_int(ws, tbls[program], byprog_ft[program], byprog_bsfin_int[nxt_yr1], byprog_bsfin_int[nxt_yr2])
        else:
            ws = wb[program]
            update_sheet_with_ft_int(ws, tbls[program], byprog_ft[program], byprog_int[program])

    reset_workbook_view_to_a1(wb)

    wb.save(wb_path)
    print(f"Updated: {wb_path}")
    return str(wb_path)

if __name__ == "__main__":
    # example
    main(["MPA"])
