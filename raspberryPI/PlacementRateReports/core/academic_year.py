"""Academic-year rules shared by the reports, dashboards and the Friday runner.

Class of Y: students expected to graduate from August Y-1 through June Y.
Weekly history window: the last week of June Y-1 through the last week of September Y.
The report sent to people is class Y from October 1, Y-1 through September 30, Y;
between late June and September the next class is collected but only archived.
"""
from datetime import date, timedelta


def last_friday(year, month):
    day = date(year, month + 1, 1) - timedelta(days=1) if month < 12 else date(year, 12, 31)
    return day - timedelta(days=(day.weekday() - 4) % 7)


def window_start(class_year):
    """Monday of the last week of June before the class graduates (2027 -> 2026-06-22)."""
    friday = last_friday(int(class_year) - 1, 6)
    return friday - timedelta(days=4)


def window_end(class_year):
    """September 30 of the graduation year, so the last Friday of September is included."""
    return date(int(class_year), 9, 30)


def in_window(class_year, day):
    return window_start(class_year) <= day <= window_end(class_year)


def report_class_year(day):
    """The class named on the report people receive: October 2026 -> 2027, September 2026 -> 2026."""
    return day.year + 1 if day.month >= 10 else day.year


SEASONS = ('Winter', 'Spring', 'Summer', 'Fall')


def graduation_terms(class_year):
    """Terms graduating August Y-1 through June Y, named as the database stores them ('Fall 2026')."""
    y = int(class_year)
    return {f'Summer {y - 1}', f'Fall {y - 1}', f'Winter {y}', f'Spring {y}'}


def later_terms(class_year, years=6):
    """Terms after June Y (Summer Y onward); students graduating then belong to a later class."""
    y = int(class_year)
    return [f'Summer {y}', f'Fall {y}'] + [f'{s} {y + n}' for n in range(1, years) for s in SEASONS]


def earlier_terms(class_year, years=6):
    """Terms before August Y-1 (Spring Y-1 and earlier); graduates from then belong to an earlier class."""
    y = int(class_year)
    return [f'{s} {y - n}' for n in range(years, 1, -1) for s in SEASONS] + [f'Winter {y - 1}', f'Spring {y - 1}']


def graduation_term_sql(prefix=''):
    """The actual graduation term when there is one, else the expected term ('' when neither)."""
    return f"COALESCE(NULLIF({prefix}act_grad_semester, ''), {prefix}exp_grad_semester, '')"


def full_time_population_sql(class_year, prefix=''):
    """SQL condition for the Class of Y full-time population.

    Base rule: class_of Y enrolled or graduated, plus class_of Y-1 and Y-2 still enrolled.
    The graduation term then decides: anyone graduating August Y-1..June Y is added (a class_of
    Y-2 student who graduated August Y-1), anyone graduating after June Y is left for a later
    class (a class_of Y student graduating August Y), and anyone who already graduated before
    August Y-1 was counted in an earlier class. Students still enrolled past their term stay.
    Program and semester_byu exclusions stay in each query.
    """
    y, p = int(class_year), prefix
    term = graduation_term_sql(prefix)
    quoted = lambda terms: ', '.join(f"'{t}'" for t in terms)
    return (f"((({p}class_of = {y} AND {p}enroll_status IN ('Enrolled', 'Graduated'))"
            f" OR ({p}class_of IN ({y - 1}, {y - 2}) AND {p}enroll_status = 'Enrolled')"
            f" OR {term} IN ({quoted(sorted(graduation_terms(y)))}))"
            f" AND {term} NOT IN ({quoted(later_terms(y))})"
            f" AND NOT ({p}enroll_status = 'Graduated' AND {term} IN ({quoted(earlier_terms(y))})))")
