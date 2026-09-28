# Friday 2027 report automation: repository and executed example

This is the report-automation project, its test fixtures, and the output from a
fresh complete offline run on September 28, 2026. The application code is the
September 24 release. The report data is the supplied **August 21, 2026 snapshot
for the Class of 2027**, not current live placement data.

## Open the example

- [Leadership PDF: 16 pages](examples/2026-08-21/run/preview/2026-08-21/2027/leadership/2027-FullTime-Leadership-2026-08-21.pdf)
- [Editable leadership dashboard](examples/2026-08-21/run/preview/2026-08-21/2027/leadership/2027-Dashboard-Leadership-2026-08-21.xlsx)
- [Leadership report workbook](examples/2026-08-21/run/preview/2026-08-21/2027/leadership/2027_weekly_placement_report.xlsx)
- [All audience outputs](examples/2026-08-21/run/preview/2026-08-21/2027/)
- [Detailed run results](examples/2026-08-21/RESULTS.json)
- [Test log](examples/2026-08-21/regression-tests.log)

Each audience folder contains its report workbook, manifest and saved `.eml`
messages. Supported groups also have an editable dashboard, provenance JSON and
PDF. The `.eml` previews contain the actual generated attachments and can be opened
in a compatible mail application. Leadership/Zoho recipients are empty in this
offline example because their live environment files were not loaded. Director
routing is the supplied mapping.

## Executed results

- 29 regression tests passed from a fresh extraction of the released code.
- 12 report workbooks and 11 editable dashboard workbooks prepared.
- 11 PDFs containing 50 pages, including the 16-page leadership PDF.
- 25 saved email previews: 11 human PDF+XLSX messages, one Healthcare XLSX message,
  and 13 workbook-only archive/import messages.
- All artifact hashes and attachment payloads verified. All 39 selected current
  rates/monthly cells and 429 status/count cells reconciled to the attached reports.
- All leadership pages and selected director pages visually reviewed.
- No live SQL, SMTP, Box/Zoho upload, Pi deployment or cron change occurred.

| Audience key | Programs | PDF pages | Email previews |
| --- | --- | ---: | ---: |
| leadership | Leadership: overall + all programs | 16 | 3 |
| Tracie | BSAcc, MAcc | 4 | 2 |
| Noelani | BSEDM | 3 | 2 |
| Soraya | BSHRM, BSStrat | 4 | 2 |
| Tanya | BSFin | 3 | 2 |
| Kurt | BSGSCM | 3 | 2 |
| Steven | BSEnt, BSBusM | 4 | 2 |
| Bob | BSIS, MISM | 4 | 2 |
| Mike | BSMktg | 3 | 2 |
| Perry | MBA | 3 | 2 |
| Staci | MPA | 3 | 2 |
| Erica | Healthcare | 0 | 2 |

Healthcare has no dashboard page in the supplied template and remains XLSX-only.
Internship details remain in the workbooks; the PDFs cover full-time placement.
Fixtures retain aggregate history and omit individual placement-detail records.

## Repository contents

`raspberryPI/PlacementRateReports/` contains the actual application modules, SQL
builders, configuration, template, branding, tests and deployment documentation.
`examples/2026-08-21/` contains this completed run. `run_example.py` is a small
launcher that always passes `--offline`. The package contains the reporting
project; unrelated Raspberry Pi applications and git commit history are not included.

Application settings remain `delivery_enabled: false`. Live rolling report history
and credential files belong on the Pi and are not provided as deployment inputs.
The fixture workbooks are used only through `--source-root tests/fixtures`.

## Run the example yourself

Viewing the existing PDFs and workbooks does not require Python. Running the
scheduler example requires **Linux/Raspberry Pi and Python 3.10+**, the Python
packages in `raspberryPI/PlacementRateReports/requirements.txt`, and LibreOffice,
Poppler (`pdftoppm`) and DejaVu Sans fonts. See the project's DEPLOYMENT.md for setup.

From the extracted repository root, using the configured Python environment:

```bash
python3 run_example.py
```

New output goes to `example-runs/2026-08-21/`. Choose a fresh destination for a
new execution rather than reusing the bundled evidence:

```bash
python3 run_example.py --output /tmp/bcc-fresh-example
```

The exact report command used for this delivered run, from PlacementRateReports:

```bash
python3 -m core.daily_runner --offline --cohort 2027 --date 2026-08-21 \
  --source-root tests/fixtures --run-root ../../examples/2026-08-21/run
```

Inspect the bundled run after extracting the ZIP:

```bash
cd raspberryPI/PlacementRateReports
python3 -m core.status --run-root ../../examples/2026-08-21/run --last-report --verify
```

The exported summaries use relative bundle paths so this inspection works after
moving the repository. Generated workbooks, PDFs, email bytes and per-route ledgers
are unchanged. EXPORT_NOTES.json describes the path adaptation and preserves the
original execution summary. Stored source paths in provenance describe where the
run actually executed. No delivered file claims that an email was sent.

## Implementation and deployment references

- [Current code documentation](raspberryPI/PlacementRateReports/DOCUMENTATION.md)
- [Exact file/function line map](raspberryPI/PlacementRateReports/SOURCE_MAP.md)
- [Deployment and controlled activation](raspberryPI/PlacementRateReports/DEPLOYMENT.md)
- [Original implementation plan and handoff](raspberryPI/PlacementRateReports/IMPLEMENTATION_PLAN.md)
- [September 24 changes](raspberryPI/PlacementRateReports/CHANGES-2026-09-24.md)

RELEASE_MANIFEST.json and the project's SOURCE_MANIFEST.json identify the September
24 code release. PACKAGE_MANIFEST.json covers this repository plus its example
outputs and new launcher. The original application files were not changed while
running this example. Actual Pi performance, live SQL and delivery still require
target-system validation.
