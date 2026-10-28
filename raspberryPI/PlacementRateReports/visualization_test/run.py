"""Visualization test pipeline: dashboards, slide decks and test-only emails.

Separate from core.daily_runner. It reads the report workbooks the production
pipeline already wrote (read-only), never queries SQL, never writes production
files, and only ever emails the test_recipients in visualization_test/config.yml.

    cd PlacementRateReports
    python -m visualization_test                 # today, previews only
    python -m visualization_test --send          # also email the test inboxes
    python -m visualization_test --date 2026-08-21 --source-root tests/fixtures
"""
import argparse
from datetime import date, datetime, timezone
import json
import logging
from pathlib import Path
import yaml
from core.cohort_manager import CohortManager
from core.daily_runner import load_cohorts
from core.delivery import message, save_route, deliver, validate_recipients
from core.report_artifacts import ROOT, write_json
from ProgramReports.email_program_reports import program_dict
from Visualization.update_dashboard import ContractError
from .checks import check_dashboard
from .dashboard import academic_class_year, build_dashboard
from .emails import subject_line, body_text
from .render import render_sheet_images
from .slides import build_deck, deck_filename

HERE = Path(__file__).resolve().parent
logger = logging.getLogger(__name__)


def load_config(path):
    with Path(path).open() as stream:
        config = yaml.safe_load(stream)
    recipients = config.get('test_recipients') or []
    box = {str(k): v for k, v in (config.get('box_upload_addresses') or {}).items()}
    directors = config.get('email_directors') or {}
    if not isinstance(config.get('template'), str):
        raise ValueError('config template must be a path')
    if not isinstance(recipients, list) or not all(isinstance(x, str) for x in recipients):
        raise ValueError('test_recipients must be a list of addresses')
    if recipients:
        validate_recipients(recipients)
    for year, addresses in box.items():
        if not isinstance(addresses, list) or not all(isinstance(x, str) for x in addresses):
            raise ValueError('box_upload_addresses must map class years to lists of Box upload addresses')
        if addresses:
            validate_recipients(addresses)
    if not isinstance(directors, dict) or not all(
            isinstance(v, list) and v and all(isinstance(n, str) for n in v) for v in directors.values()):
        raise ValueError('email_directors must map director keys to lists of names')
    return {**config, 'test_recipients': recipients, 'box_upload_addresses': box, 'email_directors': directors}


def report_path(cohort, programs, source_root):
    if programs is None:
        return Path(source_root) / cohort.overview_template
    name = f"{cohort.id}-WeeklyPlacement-{'-'.join(programs)}.xlsx"
    return Path(source_root) / 'ProgramReports' / f'{cohort.id}_Placement_Reports' / name


def targets(only=()):
    found = [('leadership', None, [])] + [(key, tuple(data['programs']), list(data['emails']))
                                          for key, data in program_dict.items()
                                          if tuple(data['programs']) != ('Healthcare',)]
    unknown = set(only) - {key for key, _, _ in found}
    if unknown:
        raise ValueError(f'Unknown audience: {sorted(unknown)}')
    return [t for t in found if not only or t[0] in only]


def prior_report(prior, programs, source_root):
    """The prior class's own report for this audience, else its leadership report."""
    if prior is None:
        return None
    for path in ([report_path(prior, programs, source_root)] if programs else []) + [
            report_path(prior, None, source_root)]:
        if path.is_file():
            return path
    return None


def run_audience(audience, programs, intended, *, cohort, prior, day, config, args):
    leadership = programs is None
    class_year = int(cohort.id)
    directory = args.run_root / str(day) / audience
    directory.mkdir(parents=True, exist_ok=True)
    label = 'Leadership' if leadership else '-'.join(programs)
    dashboard = directory / f'{class_year}-Dashboard-{label}-{day}.xlsx'
    # Directors normally read their own program workbook; --director-source leadership
    # reads the same programs from the leadership report instead (testing only).
    source = report_path(cohort, None if args.director_source == 'leadership' else programs,
                         args.source_root)
    metadata = build_dashboard(source, dashboard,
                               class_year=class_year, as_of=day, start_date=cohort.start_date,
                               programs=programs, template=args.template,
                               state_path=args.run_root / 'state' / f'{audience}.json',
                               prior_source=prior_report(prior, programs, args.source_root))
    # Never screenshot or send a dashboard that disagrees with its report workbook.
    checks = check_dashboard(dashboard, source, pages=metadata['pages'], class_year=class_year,
                             as_of=day, start_date=cohort.start_date)
    write_json(directory / 'checks.json', checks)
    if checks['errors']:
        raise ContractError(f"Dashboard cross-check failed ({len(checks['errors'])}): {checks['errors'][0]}")
    images = render_sheet_images(dashboard, metadata['pages'], directory / 'slide-images',
                                 **{k: config[k] for k in ('dpi', 'timeout', 'print_area') if k in config})
    deck = directory / deck_filename(day)
    build_deck(images, deck, report_date=day)

    years = [class_year]
    subject = subject_line(years, 'All Programs' if leadership else ', '.join(programs), day)
    body = body_text(config['email_directors'].get(audience, [audience]), day)
    attachments = [dashboard, deck]
    recipients = config['test_recipients']
    manifest_path = directory / 'manifest.json'
    manifest = json.loads(manifest_path.read_text()) if manifest_path.exists() else {'routes': {}}
    def prepare(channel, build, envelope):
        if manifest['routes'].get(channel, {}).get('state') != 'sent':
            manifest['routes'].pop(channel, None)  # Rebuilt attachments replace an unsent preview.
            save_route(directory, manifest, channel, build(), envelope)

    def email():
        msg = message(subject, body, attachments, recipients or ['undisclosed-recipients:;'],
                      message_key=f'visualization-test/{day}/{audience}')
        msg['X-Intended-To'] = ', '.join(intended) if intended else 'leadership distribution'
        return msg
    prepare('email', email, recipients)
    # Box: <year> Placement Reports/Placement Visualizations, by the current class year.
    box = config['box_upload_addresses'].get(str(class_year), [])
    if box:
        # Every audience's deck shares one name; label the Box copy so uploads don't collide.
        box_deck = directory / 'box' / f'{deck.stem}-{label}{deck.suffix}'
        box_deck.parent.mkdir(exist_ok=True)
        box_deck.write_bytes(deck.read_bytes())
        prepare('box', lambda: message(f'Visualization upload: {label} {day}', 'Automated weekly visualization archive.',
                                       [dashboard, box_deck], box, message_key=f'visualization-test/{day}/{audience}/box'), box)
    manifest.update({'audience': audience, 'class_years': years, 'report_date': str(day),
                     'dashboard': dashboard.name, 'deck': deck.name, 'source': source.name,
                     'checks': {'errors': 0, 'weeks': checks['weeks']},
                     'updated_at': datetime.now(timezone.utc).isoformat()})
    write_json(manifest_path, manifest)
    if args.send:
        from core.database import load_environment
        channels = (('email',) if recipients else ()) + (('box',) if box else ())
        deliver(directory, manifest, dry_run=False, allowed_channels=channels,
                settings=load_environment(ROOT / 'ProgramReports'))
    return {'audience': audience, 'status': 'sent' if args.send else 'preview',
            'subject': subject, 'deck': str(deck), 'dashboard': str(dashboard)}


def acquire_lock(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    handle = path.open('a')
    try:
        import fcntl
        fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except ImportError:  # Windows (local testing)
        import msvcrt
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
    return handle


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--date', type=date.fromisoformat, default=date.today())
    parser.add_argument('--audience', action='append', default=[], help='leadership or a director key')
    parser.add_argument('--send', action='store_true', help='Email the configured test_recipients only')
    parser.add_argument('--config', type=Path, default=HERE / 'config.yml')
    parser.add_argument('--template', type=Path, help='Override the dashboard template from config')
    parser.add_argument('--source-root', type=Path, default=ROOT, help='Where the report workbooks are read from')
    parser.add_argument('--run-root', type=Path, default=ROOT / 'visualization_test_runs')
    parser.add_argument('--director-source', choices=('program', 'leadership'), default='program',
                        help="Directors' data: their own program workbook (default) or the leadership report")
    args = parser.parse_args(argv)
    config = load_config(args.config)
    if args.send and not (config['test_recipients'] or any(config['box_upload_addresses'].values())):
        parser.error('--send requires test_recipients or box_upload_addresses in the config; '
                     'directors are never emailed from here')
    args.template = (args.template or ROOT / config['template']).resolve()
    args.source_root, args.run_root = args.source_root.resolve(), args.run_root.resolve()

    _, cohorts = load_cohorts(ROOT / 'config/cohorts.yml')
    # The report class switches each October (core.academic_year); a class missing from
    # cohorts.yml is derived from the latest earlier one.
    manager = CohortManager(cohorts)
    class_year = academic_class_year(args.date)
    cohort = manager.cohort(class_year)
    if cohort is None:
        parser.error(f'Class of {class_year} is not configured in config/cohorts.yml')
    # The previous class's final report refreshes its comparison column when it exists.
    prior = manager.cohort(class_year - 1)

    lock = acquire_lock(args.run_root / 'visualization_test.lock')
    try:
        results = []
        for audience, programs, intended in targets(args.audience):
            try:
                results.append(run_audience(audience, programs, intended, cohort=cohort, prior=prior,
                                            day=args.date, config=config, args=args))
            except Exception as exc:
                logger.exception('%s failed', audience)
                results.append({'audience': audience, 'status': 'failed', 'error': f'{type(exc).__name__}: {exc}'})
            logger.info('%s: %s', audience, results[-1]['status'])
        write_json(args.run_root / str(args.date) / 'summary.json',
                   {'report_date': str(args.date), 'class_year': class_year,
                    'prior_class': prior.id if prior else None, 'results': results})
        return int(any(r['status'] == 'failed' for r in results))
    finally:
        lock.close()


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    raise SystemExit(main())
