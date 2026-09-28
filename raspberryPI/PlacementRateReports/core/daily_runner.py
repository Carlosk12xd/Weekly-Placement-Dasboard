"""Cron entry point; release defaults stage previews without sending."""
import argparse
import calendar
from datetime import date, datetime
import logging
from pathlib import Path
import re
import yaml
from .cohort_manager import Cohort, CohortManager
from .report_artifacts import ROOT, RunOptions
from .report_targets import selected_reports
from .run_records import RunRecord

logger = logging.getLogger(__name__)


def report_mode(today):
    friday = today.weekday() == 4
    monthend = today.day == calendar.monthrange(today.year, today.month)[1]
    return 2 if friday and monthend else 1 if monthend else 0 if friday else None


def load_cohorts(path=ROOT / 'config/cohorts.yml'):
    with Path(path).open() as stream:
        config = yaml.safe_load(stream)
    if not isinstance(config, dict) or not isinstance(config.get('cohorts'), list):
        raise ValueError('Configuration must contain a cohorts list')
    if type(config.get('delivery_enabled', False)) is not bool:
        raise ValueError('delivery_enabled must be a YAML boolean, not quoted text')
    grace = config.get('tracking_grace_period_days', 90)
    if type(grace) is not int or grace < 0:
        raise ValueError('tracking_grace_period_days must be a nonnegative integer')
    cohorts = []
    seen = set()
    for entry in config['cohorts']:
        if not isinstance(entry, dict):
            raise ValueError('Every cohort must be a mapping')
        data = dict(entry)
        data['id'] = str(data['id'])
        if not re.fullmatch(r'\d{4}', data['id']) or data['id'] in seen:
            raise ValueError('Cohort IDs must be unique four-digit years')
        seen.add(data['id'])
        if type(data.get('enabled', True)) is not bool:
            raise ValueError('Cohort enabled must be a YAML boolean')
        for name in ('start_date', 'grad_date'):
            if not isinstance(data[name], date):
                data[name] = date.fromisoformat(data[name])
            if isinstance(data[name], datetime):
                raise ValueError('Cohort dates must be calendar dates, without times')
        if data['start_date'] > data['grad_date']:
            raise ValueError('Cohort start_date must not follow grad_date')
        for name in ('internship_years', 'semester_byu'):
            if not isinstance(data.get(name, []), list):
                raise ValueError(f'{name} must be a list')
            data[name] = [str(v) for v in data.get(name, [])]
        modes = data.get('human_modes', [0, 1, 2])
        if not isinstance(modes, list) or any(type(m) is not int or m not in (0, 1, 2) for m in modes):
            raise ValueError('human_modes may contain only 0, 1, 2')
        cohorts.append(Cohort(**data))
    return config, cohorts


def preflight(cohorts, options):
    """Validate shared configuration; input/render failures stay bundle-local."""
    from Visualization.build_visualization import load_profile
    load_profile()
    return selected_reports(cohorts, options.source_root, options.audiences)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, default=ROOT / 'config/cohorts.yml')
    parser.add_argument('--date', type=date.fromisoformat, default=date.today())
    parser.add_argument('--cohort', action='append')
    parser.add_argument('--audience', action='append', help='leadership or a current director dictionary key')
    delivery = parser.add_mutually_exclusive_group()
    delivery.add_argument('--dry-run', action='store_true', help='Staged copies and .eml previews; no SMTP or live workbook writes')
    delivery.add_argument('--send', action='store_true', help='Enable configured SMTP routes for this live run')
    parser.add_argument('--offline', action='store_true', help='Use saved report files without SQL; always no-send')
    parser.add_argument('--retry', action='store_true', help='Use completed saved snapshots; never requery SQL')
    parser.add_argument('--test-recipient', default='', help='Redirect ALL routes, including archives, to one pilot inbox')
    parser.add_argument('--resend-uncertain', action='store_true', help='Resend uncertain routes after reviewing their SMTP outcome')
    parser.add_argument('--dashboard-only', action='store_true', help='Prepare dashboards without PDF conversion; no-send only')
    parser.add_argument('--run-root', type=Path, default=ROOT / 'runs')
    parser.add_argument('--source-root', type=Path, default=ROOT, help='Alternative report tree, permitted only with --offline')
    args = parser.parse_args(argv)
    if args.offline and args.send:
        parser.error('--offline cannot be combined with --send')
    if args.source_root.resolve() != ROOT and not args.offline:
        parser.error('--source-root requires --offline')
    if args.date != date.today() and not (args.offline or args.retry):
        parser.error('Historical dates require --offline or --retry; live SQL details use NOW()')
    record = RunRecord(args.run_root.resolve(), args.date, offline=args.offline,
                       retry=args.retry, pilot=bool(args.test_recipient),
                       requested_cohorts=args.cohort or [], requested_audiences=args.audience or [])
    lock_path = ROOT / 'runs/runner.lock'
    locked = False
    lock = None
    try:
        import fcntl
        lock_path.parent.mkdir(parents=True, exist_ok=True)
        lock = lock_path.open('a')
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            record.finish('blocked', publish=False, error='Another report run holds the global lock')
            logger.error('Another report run is already active')
            return 1
        locked = True
        record.save()
        config, cohorts = load_cohorts(args.config)
        if args.cohort and set(args.cohort) - {c.id for c in cohorts}:
            raise ValueError('Requested cohort is not configured')
        selected_reports([], args.source_root, args.audience or ())  # Validate selection even on skipped days.
        if args.test_recipient:
            from .delivery import validate_recipients
            validate_recipients([args.test_recipient])
        dry_run = args.dry_run or args.offline or not (args.send or config.get('delivery_enabled', False))
        if args.dashboard_only and not dry_run:
            raise ValueError('--dashboard-only requires a no-send run')
        options = RunOptions(dry_run=dry_run, offline=args.offline, run_root=args.run_root.resolve(),
            retry=args.retry, test_recipient=args.test_recipient, resend_uncertain=args.resend_uncertain,
            dashboard_only=args.dashboard_only, audiences=tuple(args.audience or ()), source_root=args.source_root.resolve())
        flag = report_mode(args.date)
        record.value.update(mode=flag, dry_run=dry_run, dashboard_only=args.dashboard_only)
        if flag is None:
            record.finish('skipped', reason='Not a Friday or month end')
            logger.info('Not a Friday or month end; no action.')
            return 0
        active = CohortManager(cohorts, config.get('tracking_grace_period_days', 90)).get_active_cohorts(args.date)
        active = [c for c in active if not args.cohort or c.id in args.cohort]
        if not active:
            record.finish('skipped', reason='No selected active cohorts for this date')
            logger.info('No active cohorts for this date.')
            return 0
        record.value.update(scheduled=True, stage='preflight', cohorts=[c.id for c in active])
        record.save()
        targets = preflight(active, options)
        record.value.update(stage='bundles', expected_bundles=len(targets))
        record.save()
        from .pipeline import execute_bundle
        for target in targets:
            record.value['current_bundle'] = f"{target['cohort'].id}/{target['audience']}"
            record.save()
            result = execute_bundle(mode=flag, report_date=args.date, options=options, **target)
            record.value['results'].append(result)
            record.save()
            logger.info('%s/%s: %s', target['cohort'].id, result['audience'], result['status'])
        failed = any(r['status'] == 'failed' for r in record.value['results'])
        record.value.pop('current_bundle', None)
        record.finish('failed' if failed else 'completed', stage='finished')
        return int(failed)
    except KeyboardInterrupt:
        record.finish('interrupted', publish=locked, error='Interrupted before completion')
        return 130
    except Exception as exc:
        logger.error('Report invocation failed: %s: %s', type(exc).__name__, exc)
        try:
            record.finish('failed', publish=locked, error=f'{type(exc).__name__}: {exc}')
        except OSError:
            logger.error('Cannot save invocation record; check run directory permissions and disk space')
        return 1
    finally:
        if lock is not None:
            lock.close()


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    raise SystemExit(main())
