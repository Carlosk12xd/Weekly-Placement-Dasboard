# Friday placement reports and dashboards

The Raspberry Pi can now prepare the full-time dashboard, update its weekly and
monthly observations, render a branded PDF locally, and attach it beside the XLSX.
The same process handles leadership and the ten supported director groups.
Healthcare remains XLSX-only; internship data remains in the report workbooks.

**Release status:** implemented and checked offline. Not installed or benchmarked
on the actual Pi. No real database queries, SMTP sends, Box uploads or Zoho writes
were performed during development. `delivery_enabled` defaults to `false`.

## Run the supplied fixture without credentials

From this directory, using Python 3.10+ and the dependencies in requirements.txt:

```bash
python -m unittest discover -s tests -v
python -m core.daily_runner --offline --cohort 2027 --date 2026-08-21 \
  --source-root tests/fixtures --run-root /tmp/bcc-fixture-preview
```

Expected: 12 report bundles, 11 PDFs, 25 saved email messages, zero SMTP calls.
The PDFs use the report's August 21, 2026 snapshot, not today's live data.
Review `last-run.json` and the per-audience `manifest.json`, dashboard, PDF and
`.eml` files in the run directory. Fixtures are separate from production paths.

For just dashboard data preparation, add `--dashboard-only`. For one group, add
`--audience leadership`, `--audience Tanya`, or another current dictionary key.

An independent offline renderer is also available:

```bash
python -m Visualization \
  --source tests/fixtures/LeadershipReport/Reports/2027_weekly_placement_report.xlsx \
  --cohort 2027 --date 2026-08-21 --start-date 2026-06-22 \
  --output /tmp/bcc-leadership-render
```

The scheduled runner remains `python -m core.daily_runner`. The release defaults
produce database-backed previews on eligible days. `--dry-run` always prevents
all SMTP routes and live workbook writes. `--offline` also prevents SQL access.
See DEPLOYMENT.md before installing or activating delivery.

## Check readiness and inspect a run

```bash
python -m core.doctor --cohort 2027
python -m core.status --last-report --verify
```

The readiness command checks local dependencies, selected inputs, settings and
storage access without SQL, SMTP or file writes. It does not test passwords or
connectivity. For fixtures, add `--offline --date 2026-08-21 --source-root tests/fixtures`.
The status command shows recorded outcomes and checks saved file hashes. Add
`--run-root /tmp/bcc-fixture-preview` to inspect the example above, or `--json`
to either command for machine-readable output.

Each runner invocation now records start/progress/completion under `invocations/`.
Configuration failures cannot silently leave the previous successful summary in
place. Missing/corrupt audience inputs are isolated; a missing converter holds
required-PDF human mail while valid workbook archives and Healthcare can continue.

## Documents and implementation

- DOCUMENTATION.md: current architecture, data rules and failure/retry behavior.
- DEPLOYMENT.md: dependencies, Pi checks, pilot, activation and rollback.
- IMPLEMENTATION_STATUS.md: completed work, evidence and remaining deployment checks.
- SOURCE_MAP.md: current function line anchors.
- IMPLEMENTATION_PLAN.md: original inspected-source design with status updates.
- AGENTS.md: concise invariants for future changes.
- THIRD_PARTY_NOTICES.md: the pinned Excel-to-PDF source attribution.
- READINESS_PLAN.md: pre-edit file/line anchors for the unattended-execution stage.
- CHANGES-2026-09-24.md: operational changes and the current validation evidence.

The overlay excludes current production report workbooks and environment files.
Do not copy the test fixtures over the Pi's live report history.
