#!/usr/bin/env python3
"""Run the included archived example without SQL or email delivery."""
import argparse
from pathlib import Path
import subprocess
import sys


def main():
    repository = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path,
                        default=repository / 'example-runs/2026-08-21')
    args = parser.parse_args()
    if sys.version_info < (3, 10) or not sys.platform.startswith('linux'):
        parser.error('Run this example on Linux/Raspberry Pi with Python 3.10 or newer')
    project = repository / 'raspberryPI/PlacementRateReports'
    output = args.output.resolve()
    common = ['--offline', '--cohort', '2027', '--date', '2026-08-21',
              '--source-root', str(project / 'tests/fixtures'), '--run-root', str(output)]
    check = subprocess.run([sys.executable, '-m', 'core.doctor', *common], cwd=project)
    if check.returncode:
        return check.returncode
    run = subprocess.run([sys.executable, '-m', 'core.daily_runner', *common], cwd=project)
    inspected = subprocess.run([sys.executable, '-m', 'core.status', '--run-root', str(output),
                               '--last-report', '--verify'], cwd=project)
    return run.returncode or inspected.returncode


if __name__ == '__main__':
    raise SystemExit(main())
