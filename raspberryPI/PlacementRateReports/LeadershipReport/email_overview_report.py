"""Leadership routing; no settings lookup, database access, or sending on import."""
from pathlib import Path
from core.report_artifacts import ROOT, RunOptions


def mainflow(cohort, a, today, options=None):
    from core.pipeline import execute_bundle
    options = options or RunOptions()
    path = Path(cohort.overview_template)
    if not path.is_absolute():
        path = options.source_root / path
    return [execute_bundle(cohort, a, today, 'leadership', path, None, None,
                           options)]
