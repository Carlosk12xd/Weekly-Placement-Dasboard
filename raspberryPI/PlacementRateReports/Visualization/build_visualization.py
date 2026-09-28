from pathlib import Path
import json
import yaml
from .update_dashboard import update_dashboard, PAGES, OVERALL, TEMPLATE
from .render_excel import render_pdf

ROOT = Path(__file__).resolve().parents[1]


def load_profile(root=ROOT):
    with (Path(root) / 'config/visualizations.yml').open() as stream:
        profile = yaml.safe_load(stream)
    if not isinstance(profile, dict) or type(profile.get('enabled', False)) is not bool:
        raise ValueError('Visualization enabled must be a YAML boolean')
    if not isinstance(profile.get('cohorts', []), list):
        raise ValueError('Visualization cohorts must be a list')
    if type(profile.get('cover', True)) is not bool:
        raise ValueError('Visualization cover must be a YAML boolean')
    for name in ('dpi', 'timeout', 'max_pdf_bytes', 'max_message_bytes'):
        if name in profile and (type(profile[name]) is not int or profile[name] <= 0):
            raise ValueError(f'Visualization {name} must be a positive integer')
    return profile


def build_visualization(source, directory, *, cohort, report_date, programs=None,
                        base=None, profile=None, render=True, source_root=None):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    profile = profile or load_profile()
    label = 'Leadership' if programs is None else '-'.join(programs)
    dashboard = directory / f'{cohort.id}-Dashboard-{label}-{report_date}.xlsx'
    pdf = directory / f'{cohort.id}-FullTime-{label}-{report_date}.pdf'
    template = ROOT / profile.get('template', str(TEMPLATE))
    historical_sources = {year: Path(source_root or ROOT) / path
                          for year, path in profile.get('historical_reports', {}).items()}
    metadata = update_dashboard(source, dashboard, cohort_year=cohort.id, as_of=report_date,
                                start_date=cohort.start_date, programs=programs,
                                base=base or template, historical_sources=historical_sources)
    if render:
        metadata['render'] = render_pdf(dashboard, pdf, pages=metadata['pages'], year=cohort.id,
            report_date=report_date, **{k: profile[k] for k in (
                'cover', 'dpi', 'timeout', 'print_area', 'max_pdf_bytes') if k in profile})
        dashboard.with_suffix('.json').write_text(json.dumps(metadata, indent=2) + '\n')
    return dashboard, pdf if render else None, metadata
