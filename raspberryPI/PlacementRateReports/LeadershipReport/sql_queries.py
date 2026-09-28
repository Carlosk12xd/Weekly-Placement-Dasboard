# Creates the summary sheet: compares each program to each other
def summary_template(programs, c_year, p_year1, p_year2, semester_byu):
    programs_sql = sql_list(programs)
    semester_sql = sql_list(semester_byu)
    return f"""
SELECT
    program,
    SUM(job_search_status = 'Accepted an offer') AS offer_accepted,
    SUM(job_search_status = 'Actively seeking') AS still_seeking,
    SUM(CASE WHEN COALESCE(job_search_status,'') IN ('Not Reported','No Recent Information Available','') THEN 1 ELSE 0 END) AS no_info,
    SUM(job_search_status LIKE 'Not seeking%') AS not_seeking,
    SUM(CASE WHEN is_international = 1 AND (work_authorization NOT IN ('U.S. Permanent Resident', 'U.S. Citizen') OR work_authorization IS NULL) THEN 1 ELSE 0 END) AS intl_all,
    COUNT(*) AS total
FROM msmdatabase.bcc_student_view
WHERE ((class_of = {c_year} and enroll_status IN ('Enrolled', 'Graduated')) or (class_of IN ({p_year1}, {p_year2}) and enroll_status = 'Enrolled'))
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND program IN {programs_sql}
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {semester_sql}
GROUP BY program
ORDER BY program;
"""

# Creates the second sheet that has full time MSB class totals split up by job_search_status
def total_full(c_year, p_year1, p_year2, semester_byu):
    semester_sql = sql_list(semester_byu)
    return f"""
SELECT
    COALESCE(job_search_status, 'Not Reported') AS job_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE ((class_of = {c_year} and enroll_status IN ('Enrolled', 'Graduated')) or (class_of IN ({p_year1}, {p_year2}) and enroll_status = 'Enrolled'))
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {semester_sql}
GROUP BY COALESCE(job_search_status, 'Not Reported')
ORDER BY job_search_status;
"""

# Creates the fourth sheet that has internship MSB class totals split up by job_search_status
def total_int(int_yrs, semester_byu):
    int_sql = sql_list(int_yrs)
    semester_sql = sql_list(semester_byu)
    return f"""
SELECT
    COALESCE(internship_search_status, 'Not Reported') AS internship_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE class_of IN {int_sql}
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {semester_sql}
GROUP BY COALESCE(internship_search_status, 'Not Reported')
ORDER BY internship_search_status;
"""

# gets each program's full time job search status
def by_program_full(prog, c_year, p_year1, p_year2, semester_byu):
    semester_sql = sql_list(semester_byu)
    # program is a single value -> treat as quoted SQL literal via sql_list and strip parentheses
    prog_sql = sql_list(prog)   # produces "('BSAcc')" or "(BSAcc)"
    # if sql_list returns "('BSAcc')" we want the literal 'BSAcc'
    prog_literal = prog_sql.strip("()")
    return f"""
SELECT
    COALESCE(job_search_status, 'Not Reported') AS job_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE ((class_of = {c_year} and enroll_status IN ('Enrolled', 'Graduated')) or (class_of IN ({p_year1}, {p_year2}) and enroll_status = 'Enrolled'))
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND program = {prog_literal}
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {semester_sql}
GROUP BY COALESCE(job_search_status, 'Not Reported')
ORDER BY job_search_status;
"""

# gets each program's internship job search status
def by_program_int(prog, int_yrs, semester_byu):
    int_sql = sql_list(int_yrs)
    semester_sql = sql_list(semester_byu)
    prog_sql = sql_list(prog).strip("()")
    return f"""
SELECT
    COALESCE(internship_search_status, 'Not Reported') AS internship_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE class_of IN {int_sql}
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND program = {prog_sql}
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {semester_sql}
GROUP BY COALESCE(internship_search_status, 'Not Reported')
ORDER BY internship_search_status;
"""

def sql_list(items):
    """
    Convert Python sequence of strings/ints into SQL (a,b,c) form with quotes for strings.
    Example: ['20265','20275'] -> "('20265','20275')"
    """
    # accept single string/number passed accidentally
    if isinstance(items, (str, int)):
        return f"('{items}')" if isinstance(items, str) else f"({items})"

    formatted = []
    for it in items:
        # leave numeric-looking values unquoted (optional)
        if isinstance(it, (int, float)) or (isinstance(it, str) and it.isdigit()):
            formatted.append(str(it))
        else:
            # escape single quotes by doubling them (very simple sanitization)
            s = str(it).replace("'", "''")
            formatted.append(f"'{s}'")
    return "(" + ",".join(formatted) + ")"

def write_overall_queries(programs, c_year, p_year1, p_year2, semester_byu, int_yrs):
    sqls = []

    sqls.append(summary_template(programs, c_year, p_year1, p_year2, semester_byu))
    sqls.append(total_full(c_year, p_year1, p_year2, semester_byu))
    sqls.append(total_int(int_yrs, semester_byu))

    return sqls


def write_byProg_queries(prog, c_year, p_year1, p_year2, semester_byu, int_yrs):
    sqls = []    
    
    sqls.append(by_program_full(prog, c_year, p_year1, p_year2, semester_byu))
    sqls.append(by_program_int(prog, int_yrs, semester_byu))

    return sqls
