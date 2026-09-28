"""Inspect saved execution and delivery outcomes; never send or repair anything."""
import argparse
import json
from pathlib import Path
from .report_artifacts import ROOT, verify_file


def local_path(root, value):
    path = (root / value).resolve()
    if not path.is_relative_to(root.resolve()):
        raise ValueError('Saved path leaves the selected run directory')
    return path


def inspect_run(run_root, *, last_report=False, verify=False):
    root = Path(run_root).resolve()
    path = root / ('last-report-run.json' if last_report else 'last-run.json')
    if not path.is_file():
        return {'status': 'missing', 'error': 'No saved invocation summary at the selected path', 'bundles': []}
    try:
        record = json.loads(path.read_text())
        if not isinstance(record.get('results'), list):
            raise ValueError('Invalid results list')
    except (OSError, ValueError, TypeError, AttributeError) as exc:
        return {'status': 'failed', 'error': f'Cannot read invocation summary ({type(exc).__name__})', 'bundles': []}
    state = record.get('status') or ('failed' if any(r.get('status') == 'failed' for r in record['results']) else 'completed')
    result = {key: record[key] for key in ('run_id', 'date', 'mode', 'started_at', 'updated_at',
        'finished_at', 'dry_run', 'offline', 'pilot', 'stage', 'current_bundle', 'error', 'reason') if key in record}
    result.update(status=state, hashes_checked=verify, bundles=[])
    if state == 'running':
        result['note'] = 'No completion was recorded; the process may still be running or may have stopped abruptly'
    for saved in record['results']:
        bundle = {key: saved[key] for key in ('cohort', 'audience', 'status', 'error') if key in saved}
        try:
            if not saved.get('directory'):
                raise ValueError('Missing bundle directory')
            directory = local_path(root, saved['directory'])
            manifest = json.loads((directory / 'manifest.json').read_text())
            bundle['cohort'] = manifest['identity']['cohort']
            bundle['artifacts'] = sorted(manifest['artifacts'])
            bundle['human_held'] = manifest.get('human_held')
            bundle['routes'] = {key: value['state'] for key, value in manifest['routes'].items()}
            if any(value in ('submitting', 'uncertain') for value in bundle['routes'].values()):
                bundle['status'] = 'failed'
                bundle['error'] = 'SMTP outcome needs review before an explicit resend'
            if manifest.get('status') == 'failed':
                bundle['status'] = 'failed'
                bundle['error'] = manifest.get('error') or manifest.get('visualization_error')
            if verify:
                for artifact in manifest['artifacts'].values():
                    artifact_path = local_path(directory, artifact['file'])
                    verify_file(artifact_path, artifact['sha256'])
                    if 'metadata_sha256' in artifact:
                        verify_file(artifact_path.with_suffix('.json'), artifact['metadata_sha256'])
                for route in manifest['routes'].values():
                    verify_file(local_path(directory, route['file']), route['sha256'])
                bundle['hashes'] = 'verified'
        except (OSError, ValueError, TypeError, KeyError, AttributeError, RuntimeError) as exc:
            bundle['status'] = 'failed'
            bundle['error'] = f'Cannot inspect saved bundle: {type(exc).__name__}: {exc}'
        if bundle.get('status') == 'failed':
            result['status'] = 'failed'
        result['bundles'].append(bundle)
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run-root', type=Path, default=ROOT / 'runs')
    parser.add_argument('--last-report', action='store_true', help='Keep the last eligible report visible after skipped days')
    parser.add_argument('--verify', action='store_true', help='Check saved XLSX, dashboard, metadata, PDF and email hashes')
    parser.add_argument('--json', action='store_true')
    args = parser.parse_args(argv)
    result = inspect_run(args.run_root, last_report=args.last_report, verify=args.verify)
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"Run: {result['status']} | report date: {result.get('date', 'unknown')}")
        if 'dry_run' in result:
            print('Mode: ' + ('offline preview' if result.get('offline') else 'no-send preview' if result['dry_run'] else 'delivery') +
                  (' (pilot)' if result.get('pilot') else ''))
        for key in ('started_at', 'finished_at', 'current_bundle', 'reason', 'error', 'note'):
            if result.get(key):
                print(f'{key}: {result[key]}')
        for bundle in result['bundles']:
            routes = ', '.join(f'{k}={v}' for k, v in bundle.get('routes', {}).items()) or 'none recorded'
            print(f"{bundle.get('cohort', '?')}/{bundle.get('audience', '?')}: {bundle.get('status', 'unknown')} | {routes}")
            for key in ('human_held', 'error'):
                if bundle.get(key):
                    print(f'  {key}: {bundle[key]}')
        print('sent means SMTP accepted the submission; receipt and archive ingestion are not verified.')
    return 0 if result['status'] in ('completed', 'skipped', 'preview') else 1


if __name__ == '__main__':
    raise SystemExit(main())
