"""Offline rendering: this entry point imports no database or email sender."""
import argparse
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from .build_visualization import build_visualization


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--cohort', required=True)
    parser.add_argument('--date', type=date.fromisoformat, required=True)
    parser.add_argument('--start-date', type=date.fromisoformat, required=True)
    parser.add_argument('--programs', nargs='+')
    parser.add_argument('--base', type=Path)
    parser.add_argument('--dashboard-only', action='store_true')
    args = parser.parse_args()
    cohort = SimpleNamespace(id=args.cohort, start_date=args.start_date)
    dashboard, pdf, _ = build_visualization(args.source, args.output, cohort=cohort,
        report_date=args.date, programs=args.programs, base=args.base, render=not args.dashboard_only)
    print(dashboard)
    if pdf:
        print(pdf)


if __name__ == '__main__':
    main()
