"""Friday step after the report bundles: dashboards and slide decks for the report class.

Uses this run's cross-checked workbook snapshots (not the rolling files, which a dry run
leaves unchanged), laid out as the source tree visualization_test expects, plus the
previous class's final reports for its comparison column. Delivery follows the run:
previews on a dry run, otherwise visualization_test's own test_recipients/Box routes
(it never emails directors). Also compares each director workbook with the leadership
report taken the same morning; those differences are warnings, never failures.
"""
import json
import logging
from pathlib import Path
import shutil
from .academic_year import report_class_year
from .cohort_manager import CohortManager
from .pipeline import namespace_for
from .report_checks import compare_reports

logger = logging.getLogger(__name__)


def snapshots(results, class_year):
    """{audience: (snapshot path, programs)} for bundles of this class whose workbook passed its checks."""
    found = {}
    for result in results:
        if result.get('cohort') != str(class_year) or not result.get('directory'):
            continue
        directory = Path(result['directory'])
        try:
            manifest = json.loads((directory / 'manifest.json').read_text())
        except (OSError, ValueError):
            continue
        if 'xlsx' in manifest.get('artifacts', {}):
            found[result['audience']] = (directory / manifest['artifacts']['xlsx']['file'],
                                         manifest['identity'].get('programs'))
    return found


def source_tree(folder, cohort, prior, current, source_root):
    """Lay out the snapshots (and the prior class's rolling reports) as <root>/LeadershipReport/...
    and <root>/ProgramReports/<year>_Placement_Reports/..."""
    if folder.exists():
        shutil.rmtree(folder)
    for audience, (path, programs) in current.items():
        target = folder / (cohort.overview_template if programs is None else
                           Path('ProgramReports') / f'{cohort.id}_Placement_Reports' / path.name)
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, target)
    if prior is not None:
        relative = [Path(prior.overview_template)]
        program_folder = Path(source_root) / 'ProgramReports' / f'{prior.id}_Placement_Reports'
        relative += [p.relative_to(source_root) for p in program_folder.glob('*.xlsx')] if program_folder.is_dir() else []
        for rel in relative:
            if (Path(source_root) / rel).is_file():
                (folder / rel).parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(Path(source_root) / rel, folder / rel)
    return folder


def run(report_date, cohorts, results, options):
    from visualization_test.run import HERE, load_config, main as visualize
    config = load_config(HERE / 'config.yml')
    if not config.get('friday_runner', False):
        return {'status': 'disabled'}
    manager = CohortManager(cohorts)
    class_year = report_class_year(report_date)
    cohort, prior = manager.cohort(class_year), manager.cohort(class_year - 1)
    current = snapshots(results, class_year)
    if cohort is None or not current:
        return {'status': 'skipped', 'reason': f'No checked Class of {class_year} workbooks in this run'}

    outcome = {'class_year': class_year, 'audiences': sorted(current), 'comparison_warnings': []}
    if 'leadership' in current:
        for audience, (path, programs) in current.items():
            if programs and tuple(programs) != ('Healthcare',):
                try:
                    outcome['comparison_warnings'] += [f'{audience}: {w}' for w in compare_reports(
                        current['leadership'][0], path, report_date=report_date, programs=programs)]
                except Exception as exc:
                    outcome['comparison_warnings'].append(f'{audience}: comparison failed: {type(exc).__name__}: {exc}')

    run_root = Path(options.run_root) / namespace_for(options) / 'visualizations'
    tree = source_tree(run_root / str(report_date) / 'source', cohort, prior, current, options.source_root)
    audiences = [a for a in current if a == 'leadership' or tuple(current[a][1]) != ('Healthcare',)]
    argv = ['--date', str(report_date), '--source-root', str(tree), '--run-root', str(run_root)]
    argv += [x for a in audiences for x in ('--audience', a)]
    send = not options.dry_run and bool(config['test_recipients'] or any(config['box_upload_addresses'].values()))
    if send:
        argv.append('--send')
    try:
        code = visualize(argv)
    except SystemExit as exc:  # argparse/config errors inside visualization_test
        code = exc.code or 1
    except Exception as exc:
        logger.exception('Friday visualizations failed')
        return {**outcome, 'status': 'failed', 'error': f'{type(exc).__name__}: {exc}'}
    summary_path = run_root / str(report_date) / 'summary.json'
    summary = json.loads(summary_path.read_text()) if summary_path.exists() else {}
    outcome.update(status='failed' if code else ('sent' if send else 'preview'),
                   run_root=str(run_root / str(report_date)),
                   results={r['audience']: r['status'] for r in summary.get('results', [])},
                   errors={r['audience']: r['error'] for r in summary.get('results', []) if r.get('error')})
    return outcome
