"""Check local prerequisites without queries, email, report generation or writes."""
import argparse
from datetime import date, datetime
from importlib import import_module, metadata
import json
import os
from pathlib import Path
import platform
import shutil
import sys
from zipfile import ZipFile

from .cohort_manager import CohortManager
from .daily_runner import load_cohorts, report_mode
from .report_artifacts import ROOT, sha256
from .report_targets import selected_reports


def inspect_setup(*, config_path=ROOT / 'config/cohorts.yml', source_root=ROOT,
                  run_root=ROOT / 'runs', report_date=None, cohort_ids=(), audiences=(),
                  offline=False, dashboard_only=False):
    """Return scoped evidence only; presence of credentials is not authentication."""
    report_date = report_date or date.today()
    checks = []
    def add(name, ok, detail, *, warning=False):
        checks.append({'check': name, 'status': 'ok' if ok else 'warning' if warning else 'error',
                       'detail': detail})

    add('python', sys.version_info >= (3, 10), platform.python_version())
    add('scheduler_platform', sys.platform.startswith('linux'), 'Linux is required for the runner lock')
    packages = {'yaml': 'PyYAML', 'openpyxl': 'openpyxl', 'lxml.etree': 'lxml',
                'PIL': 'Pillow', 'pypdf': 'pypdf'}
    if not offline:
        packages.update({'dotenv': 'python-dotenv', 'mysql.connector': 'mysql-connector-python'})
    for module, distribution in packages.items():
        try:
            import_module(module)
            version = metadata.version(distribution)
            add('package:' + distribution, distribution != 'openpyxl' or version == '3.1.5', version)
        except (ImportError, metadata.PackageNotFoundError):
            add('package:' + distribution, False, 'Missing or cannot import; install requirements.txt')

    ancestor = Path(run_root).resolve()
    while not ancestor.exists() and ancestor != ancestor.parent:
        ancestor = ancestor.parent
    add('run_storage_access', ancestor.is_dir() and os.access(ancestor, os.W_OK | os.X_OK),
        'Nearest existing run directory must be writable (permission check only)')
    if ancestor.is_dir():
        add('disk_free', True, f'{shutil.disk_usage(ancestor).free // (1024 * 1024)} MiB available; no Pi workload estimate')
    add('timezone', True, datetime.now().astimezone().strftime('%Z (%z)') + '; confirm intended cron time locally')

    result = {'scope': 'offline' if offline else 'live prerequisites', 'date': str(report_date),
              'checks': checks, 'network_checked': False, 'writes_performed': False}
    try:
        config, cohorts = load_cohorts(config_path)
        if set(cohort_ids) - {c.id for c in cohorts}:
            raise ValueError('Requested cohort is not configured')
        active = CohortManager(cohorts, config.get('tracking_grace_period_days', 90)).get_active_cohorts(report_date)
        selected = [c for c in cohorts if c.id in cohort_ids] if cohort_ids else active
        targets = selected_reports(selected, source_root, audiences)
        add('configuration', True, f'{len(selected)} selected cohorts; {len(targets)} report bundles')
        add('cohort_selection', bool(selected), 'Select an eligible cohort or pass --cohort to inspect an inactive one')
        add('calendar', report_mode(report_date) is not None, 'Selected date must be Friday or month end to execute', warning=True)
        for cohort in selected:
            add(f'eligibility:{cohort.id}', cohort in active, 'Tracking window and enabled setting', warning=True)
        add('delivery_policy', True, 'Enabled in configuration' if config.get('delivery_enabled') else 'Disabled in configuration')
    except Exception as exc:
        add('configuration', False, f'{type(exc).__name__}: {exc}')
        targets = []

    for target in targets:
        label = f"input:{target['cohort'].id}/{target['audience']}"
        try:
            with ZipFile(target['rolling']) as archive:
                valid = archive.testzip() is None and 'xl/workbook.xml' in archive.namelist()
            add(label, valid, 'Readable XLSX container; current source-date/metric validation happens during the run')
        except Exception as exc:
            add(label, False, f'Missing or unreadable XLSX ({type(exc).__name__})')

    try:
        from Visualization.build_visualization import load_profile
        from Visualization.update_dashboard import TEMPLATE_SHA256
        profile = load_profile()
        supported = [t for t in targets if profile.get('enabled') and
                     t['cohort'].id in [str(x) for x in profile.get('cohorts', [])] and
                     t['programs'] != ('Healthcare',)]
        if supported:
            template = ROOT / profile['template']
            add('dashboard_template', template.is_file() and sha256(template) == TEMPLATE_SHA256,
                'Must match the reviewed June dashboard fingerprint')
            for year, path in profile.get('historical_reports', {}).items():
                add(f'historical_report:{year}', (Path(source_root) / path).is_file(),
                    'Missing archived source retains the supplied template comparison values', warning=True)
            if not dashboard_only:
                for executable, fallback in (('soffice', 'libreoffice'), ('pdftoppm', 'pdftoppm')):
                    add('renderer:' + executable, bool(shutil.which(executable) or shutil.which(fallback)),
                        'Executable required for scheduled PDF output')
                fonts = Path('/usr/share/fonts/truetype/dejavu')
                add('branding_fonts', all((fonts / name).is_file() for name in ('DejaVuSans.ttf', 'DejaVuSans-Bold.ttf')),
                    'DejaVu Sans regular/bold recommended; missing fonts use fallbacks', warning=True)
        add('visualization_configuration', True, f'{len(supported)} bundles use the full-time dashboard')
    except Exception as exc:
        add('visualization_configuration', False, f'{type(exc).__name__}: {exc}')

    if not offline:
        from .database import load_environment
        from .delivery import env_list, validate_recipients
        for group in sorted({'LeadershipReport' if t['audience'] == 'leadership' else 'ProgramReports' for t in targets}):
            try:
                settings = load_environment(ROOT / group)
                for key in ('DB_HOST', 'DB_USER', 'DB_PASSWORD', 'DB_NAME', 'SMTP_PASS'):
                    add(f'setting:{group}/{key}', bool(settings.get(key)), 'Present' if settings.get(key) else 'Missing')
                if group == 'LeadershipReport':
                    for key in ('TO_ADDRS', 'Zoho_Box', 'CC_ADDRS', 'BCC_ADDRS'):
                        addresses = env_list(key, settings)
                        if addresses or key in ('TO_ADDRS', 'Zoho_Box'):
                            validate_recipients(addresses)
                        add('recipients:' + key, True, f'{len(addresses)} configured; addresses are not displayed')
            except Exception as exc:
                # Values and parser errors may contain secrets; expose only the exception type.
                add('settings:' + group, False, f'Cannot validate settings ({type(exc).__name__}); inspect locally')
        from .delivery import validate_recipients
        for target in targets:
            cohort = target['cohort']
            try:
                if target['recipients'] is not None:
                    validate_recipients(target['recipients'])
                for attr in ('weekly_box_leadership', 'monthly_box_leadership') if target['audience'] == 'leadership' else (
                        'weekly_box_programs', 'monthly_box_programs'):
                    validate_recipients([getattr(cohort, attr)])
            except ValueError:
                add(f"routing:{cohort.id}/{target['audience']}", False, 'Missing or malformed configured recipient')
    result['status'] = 'failed' if any(c['status'] == 'error' for c in checks) else 'passed'
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, default=ROOT / 'config/cohorts.yml')
    parser.add_argument('--source-root', type=Path, default=ROOT)
    parser.add_argument('--run-root', type=Path, default=ROOT / 'runs')
    parser.add_argument('--date', type=date.fromisoformat, default=date.today())
    parser.add_argument('--cohort', action='append', default=[])
    parser.add_argument('--audience', action='append', default=[])
    parser.add_argument('--offline', action='store_true')
    parser.add_argument('--dashboard-only', action='store_true')
    parser.add_argument('--json', action='store_true')
    args = parser.parse_args(argv)
    result = inspect_setup(config_path=args.config, source_root=args.source_root,
        run_root=args.run_root, report_date=args.date, cohort_ids=args.cohort,
        audiences=args.audience, offline=args.offline, dashboard_only=args.dashboard_only)
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"Local readiness: {result['status']} ({result['scope']})")
        for check in result['checks']:
            print(f"{check['status'].upper():7} {check['check']}: {check['detail']}")
        print('No queries, emails or files written. This does not verify credentials, connectivity or Pi performance.')
    return int(result['status'] == 'failed')


if __name__ == '__main__':
    raise SystemExit(main())
