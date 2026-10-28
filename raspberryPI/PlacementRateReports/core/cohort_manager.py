from dataclasses import dataclass, field, replace
from datetime import date
import re
from typing import List, Optional
from .academic_year import window_start, window_end

@dataclass
class Cohort:
    id: str
    label: str
    start_date: date
    grad_date: date
    internship_years: List[str]
    semester_byu: List[str]
    overview_template: str
    weekly_box_leadership: str
    weekly_box_programs: str
    monthly_box_leadership: str
    monthly_box_programs: str
    enabled: bool = True
    human_modes: List[int] = field(default_factory=lambda: [0, 1, 2])
    derived: bool = False


def shift_years(text, years):
    """Move every standalone 4-digit year (not part of a date) forward by `years`."""
    return re.sub(r'(?<!\d)(?<!\d/)(20\d\d)(?!\d)(?!/\d)', lambda m: str(int(m[1]) + years), text)


def derive_cohort(previous: Cohort, class_year: int) -> Cohort:
    """The next class's settings from the latest configured one: same Box folders and
    human modes, every year field moved forward, dates from the academic-year rules."""
    years = int(class_year) - int(previous.id)
    return replace(previous, id=str(class_year), label=f'Class of {class_year}',
                   start_date=window_start(class_year), grad_date=window_start(int(class_year) + 1),
                   internship_years=[str(int(y) + years) for y in previous.internship_years],
                   semester_byu=[str(int(s) + 10 * years) for s in previous.semester_byu],
                   overview_template=shift_years(previous.overview_template, years),
                   enabled=True, derived=True)


class CohortManager:
    def __init__(self, cohorts: List[Cohort]):
        self.cohorts = cohorts

    def tracking_end(self, cohort: Cohort) -> date:
        return window_end(cohort.id)

    def cohort(self, class_year) -> Optional[Cohort]:
        """The configured class, else one derived from the latest earlier configured class."""
        class_year = str(class_year)
        found = next((c for c in self.cohorts if c.id == class_year), None)
        if found is not None:
            return found
        earlier = [c for c in self.cohorts if int(c.id) < int(class_year)]
        if not earlier:
            return None
        return derive_cohort(max(earlier, key=lambda c: int(c.id)), int(class_year))

    def get_active_cohorts(self, as_of: Optional[date] = None):
        if as_of is None:
            as_of = date.today()
        # At most two classes overlap: this calendar year's (through September) and
        # next year's (from the last week of June).
        candidates = [self.cohort(year) for year in (as_of.year, as_of.year + 1)]
        return [c for c in candidates
                if c is not None and c.enabled and c.start_date <= as_of <= self.tracking_end(c)]
