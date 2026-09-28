# Below Queries are all for the create_program_report file
def sql_total_full(c_year, p_year1, p_year2, semester_byu):
    sem_byu = sql_list(semester_byu)
    return f"""
SELECT
    COALESCE(job_search_status, 'Not Reported') AS job_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE ((class_of = {c_year} and enroll_status IN ("Enrolled", "Graduated")) or (class_of IN ({p_year1}, {p_year2}) and enroll_status = "Enrolled"))
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {sem_byu}
GROUP BY COALESCE(job_search_status, 'Not Reported')
ORDER BY job_search_status;
"""

def sql_total_int(int_yrs, semester_byu):
    intern_yrs = sql_list(int_yrs)
    sem_byu = sql_list(semester_byu)
    return f"""
SELECT
    COALESCE(internship_search_status, 'Not Reported') AS internship_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE class_of IN {intern_yrs}
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {sem_byu}
GROUP BY COALESCE(internship_search_status, 'Not Reported')
ORDER BY internship_search_status;
"""

def sql_byprog_full(c_year, p_year1, p_year2, semester_byu, prog):
    sem_byu = sql_list(semester_byu)
    prog_sql = sql_list(prog)
    program = prog_sql.strip("()")
    return f"""
SELECT
    COALESCE(job_search_status, 'Not Reported') AS job_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE ((class_of = {c_year} and enroll_status IN ("Enrolled", "Graduated")) or (class_of IN ({p_year1}, {p_year2}) and enroll_status = "Enrolled"))
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND program = {program}
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {sem_byu}
GROUP BY COALESCE(job_search_status, 'Not Reported')
ORDER BY job_search_status;
"""

def sql_byprog_int(int_yrs, semester_byu, prog):
    sem_byu = sql_list(semester_byu)
    intern_yrs = sql_list(int_yrs)
    prog_sql = sql_list(prog)
    program = prog_sql.strip("()")
    return f"""
SELECT
    COALESCE(internship_search_status, 'Not Reported') AS internship_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE class_of IN {intern_yrs}
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND program = {program}
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {sem_byu}
GROUP BY COALESCE(internship_search_status, 'Not Reported')
ORDER BY internship_search_status;
"""

def sql_bsfin_int(class_of, semester_byu):
    sem_byu = sql_list(semester_byu)
    return f"""
SELECT
    COALESCE(internship_search_status, 'Not Reported') AS internship_search_status,
    COUNT(*) AS count
FROM msmdatabase.bcc_student_view
WHERE class_of = {class_of}
  AND program NOT IN ('EMBA','EMPA','StratMnr')
  AND program = 'BSFin'
  AND enroll_status IN ('Enrolled','Graduated')
  AND record_status = 'A'
  AND semester_byu NOT IN {sem_byu}
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

def write_sql_overall_queries(c_year, p_year1, p_year2, semester_byu, int_yrs):
    sqls = []

    sqls.append(sql_total_full(c_year, p_year1, p_year2, semester_byu))
    sqls.append(sql_total_int(int_yrs, semester_byu))

    return sqls

def write_sql_program_queries(c_year, p_year1, p_year2, semester_byu, int_yrs, prog):
    sqls = []

    sqls.append(sql_byprog_full(c_year, p_year1, p_year2, semester_byu, prog))
    sqls.append(sql_byprog_int(int_yrs, semester_byu, prog))

    return sqls

def write_sql_bsfin_queries(c_year, p_year1, p_year2, semester_byu, prog):
    sqls = []

    sqls.append(sql_byprog_full(c_year, p_year1, p_year2, semester_byu, prog))

    nxt_yr1 = int(c_year) + 1
    nxt_yr2 = int(c_year) + 2

    sqls.append(sql_bsfin_int(nxt_yr1, semester_byu))
    sqls.append(sql_bsfin_int(nxt_yr2, semester_byu))

    return sqls


# -------------------------------------------------------------------------------------------------------
# Below are queries for the placement_details (pd) file

def pd_ft_sql(prog, c_year, p_year1, p_year2):
    prog_sql = sql_list(prog)
    program = prog_sql.strip("()")
    return f"""
SELECT 
    a.last_name,
    a.first_name,
	a.byu_net_id,
    b.organization_name,
	b.placement_job_role,
    b.job_title,
    b.state,
    b.city,
	b.offer_date,
    b.start_date,
    a.job_search_status
FROM msmdatabase.bcc_student_view a
INNER JOIN msmdatabase.student_job_offer_view b ON a.student_id = b.student_id
WHERE b.created_date >= NOW() - INTERVAL 7 DAY
	AND a.program = {program}
	AND a.record_status = 'A'
    AND b.is_accepted = '1'
	AND enroll_status IN ('Enrolled','Graduated')
    AND ((class_of = {c_year} and enroll_status IN ("Enrolled", "Graduated")) or (class_of IN ({p_year1}, {p_year2}) and enroll_status = "Enrolled"))
ORDER BY a.last_name;
"""

def pd_ft_health_sql(c_year, p_year1, p_year2):
    return f"""
SELECT 
    a.last_name,
    a.first_name,
	a.byu_net_id,
    a.program,
    b.organization_name,
	b.placement_job_role,
    b.job_title,
    b.state,
    b.city,
	b.offer_date,
    b.start_date,
    a.job_search_status
FROM msmdatabase.bcc_student_view a
INNER JOIN msmdatabase.student_job_offer_view b ON a.student_id = b.student_id
WHERE b.created_date >= NOW() - INTERVAL 7 DAY
	AND b.healthcare_related = 'Y'
	AND a.record_status = 'A'
    AND b.is_accepted = '1'
	AND enroll_status IN ('Enrolled','Graduated')
    AND ((class_of = {c_year} and enroll_status IN ("Enrolled", "Graduated")) or (class_of IN ({p_year1}, {p_year2}) and enroll_status = "Enrolled"))
ORDER BY a.last_name;
"""

def pd_int_sql(prog, c_year):
    prog_sql = sql_list(prog)
    program = prog_sql.strip("()")
    return f"""
SELECT 
    a.last_name,
    a.first_name,
	a.byu_net_id,
    a.class_of,
    b.organization_name,
	b.placement_job_role,
    b.title,
    b.state,
    b.city,
	b.offer_date,
    b.start_date,
    a.internship_search_status
FROM msmdatabase.bcc_student_view a
INNER JOIN msmdatabase.student_internship_view b ON a.student_id = b.student_id
WHERE b.created_date >= NOW() - INTERVAL 7 DAY
	AND a.program = {program}
	AND a.record_status = 'A'
    AND b.is_accepted = '1'
	AND enroll_status IN ('Enrolled','Graduated')
    AND class_of IN ({c_year}, {int(c_year) + 1}, {int(c_year) + 2}, {int(c_year) + 3})
ORDER BY a.last_name;
"""

def pd_int_health_sql(c_year):
    return f"""
SELECT 
    a.last_name,
    a.first_name,
	a.byu_net_id,
    a.program,
    b.organization_name,
	b.placement_job_role,
    b.title,
    b.state,
    b.city,
	b.offer_date,
    b.start_date,
    a.internship_search_status
FROM msmdatabase.bcc_student_view a
INNER JOIN msmdatabase.student_internship_view b ON a.student_id = b.student_id
WHERE b.created_date >= NOW() - INTERVAL 7 DAY
	AND b.healthcare_related = 'Y'
	AND a.record_status = 'A'
    AND b.is_accepted = '1'
	AND enroll_status IN ('Enrolled','Graduated')
    AND class_of IN ({c_year}, {int(c_year) + 1}, {int(c_year) + 2}, {int(c_year) + 3})
ORDER BY a.last_name;
"""

def pd_int_bsfin_sql(class_of):
    return f"""
SELECT 
    a.last_name,
    a.first_name,
	a.byu_net_id,
    b.organization_name,
	b.placement_job_role,
    b.title,
    b.state,
    b.city,
	b.offer_date,
    b.start_date,
    a.internship_search_status
FROM msmdatabase.bcc_student_view a
INNER JOIN msmdatabase.student_internship_view b ON a.student_id = b.student_id
WHERE b.created_date >= NOW() - INTERVAL 7 DAY
	AND a.program = 'BSFin'
	AND a.record_status = 'A'
    AND b.is_accepted = '1'
	AND enroll_status IN ('Enrolled','Graduated')
    AND class_of = {class_of}
ORDER BY a.last_name;
"""

def pd_total_ft_health_sql(c_year):
    return f"""
select 
	preferred_full_name as "Student Name", CONCAT(byu_net_id, "@byu.edu") as "BYU Email", 
    CASE WHEN
		is_international = 1 THEN "Yes"
        ELSE "No"
	END as "International?", -- student details
    program as "Major", class_of as "Class of", -- program details
    exp_grad_semester as "Expected Graduation", act_grad_semester as "Actual Graduation", enroll_status as "Enrollment", -- enrollment details
    organization_name as "Company", city as "City", state as "State", -- company information
    offer_date as "Offer Date", accepted_date as "Accepted Date", start_date as "Start Date", date(created_date) as "Reported Date", -- timeline information
    job_title as "Job Title", placement_job_role as "Job Role", placement_source as "Source", -- job specific information
    annual_salary as "Salary" -- payment information
from msmdatabase.bcc_student_view bsv
inner join msmdatabase.student_job_offer_view sjo on bsv.student_id = sjo.student_id
where offer_date BETWEEN '{int(c_year) - 1}-08-01' AND '{c_year}-07-31'
    and healthcare_related = "Y"
    and is_accepted = 1
order by preferred_full_name ASC;
"""

def pd_total_int_health_sql(c_year, start_date, end_date):
    return f"""
select 
	preferred_full_name as "Student Name", CONCAT(byu_net_id, "@byu.edu") as "BYU Email", -- student details
    CASE WHEN
		is_international = 1 THEN "Yes"
        ELSE "No"
	END as "International?", 
    program as "Major", class_of as "Class of", -- program details
    exp_grad_semester as "Expected Graduation", -- enrollment details
    act_grad_semester as "Actual Graduation", 
    enroll_status as "Enrollment", 
    organization_name as "Company",  -- company information
    city as "City", 
    state as "State",
    offer_date as "Offer Date", -- timeline information
    accepted_date as "Accepted Date", 
    start_date as "Start Date", 
    date(created_date) as "Reported Date", 
    title as "Job Title", -- job specific information
    placement_job_role as "Job Role", 
    placement_source as "Source", 
    monthly_salary as "Salary" -- payment information
from msmdatabase.bcc_student_view bsv
inner join msmdatabase.student_internship_view sjo on bsv.student_id = sjo.student_id
where
	(
	(sjo.start_date IS NOT NULL AND sjo.start_date BETWEEN '{int(c_year) - 1}-08-01' AND '{c_year}-07-31')
    OR
    (sjo.start_date IS NULL AND sjo.offer_date BETWEEN '{int(c_year) - 1}-08-01' AND '{c_year}-07-31')
    )
	and healthcare_related = "Y"
    and is_accepted = 1
order by preferred_full_name ASC;
"""

def write_pd_bsfin(prog, cohort):
    sqls = []
    
    c_year = cohort.id
    p_year1 = int(c_year) - 1
    p_year2 = int(c_year) - 2

    sqls.append(pd_ft_sql(prog, c_year, p_year1, p_year2))

    int_class_of = int(c_year) + 1
    sqls.append(pd_int_bsfin_sql(int_class_of))

    int_class_of = int(c_year) + 2
    sqls.append(pd_int_bsfin_sql(int_class_of))

    return sqls

def write_pd_program(prog, cohort):
    sqls = []

    c_year = cohort.id
    p_year1 = int(c_year) - 1
    p_year2 = int(c_year) - 2

    sqls.append(pd_ft_sql(prog, c_year, p_year1, p_year2))
    sqls.append(pd_int_sql(prog, c_year))

    return sqls

def write_pd_healthcare(cohort):
    sqls = []

    c_year = cohort.id
    p_year1 = int(c_year) - 1
    p_year2 = int(c_year) - 2
    start_date = cohort.start_date
    end_date = cohort.grad_date

    sqls.append(pd_ft_health_sql(c_year, p_year1, p_year2))
    sqls.append(pd_int_health_sql(c_year))

    sqls.append(pd_total_ft_health_sql(c_year))
    sqls.append(pd_total_int_health_sql(c_year, start_date, end_date))

    return sqls