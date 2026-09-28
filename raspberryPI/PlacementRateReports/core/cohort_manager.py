from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import List, Optional

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

class CohortManager:
    def __init__(self, cohorts: List[Cohort], grace_days: int = 90):
        self.cohorts = cohorts
        self.grace_days = grace_days

    def tracking_end(self, cohort: Cohort) -> date:
        return cohort.grad_date + timedelta(days=self.grace_days)

    def get_active_cohorts(self, as_of: Optional[date] = None):
        if as_of is None:
            as_of = date.today()
        return [c for c in self.cohorts if c.enabled and c.start_date <= as_of <= self.tracking_end(c)]
