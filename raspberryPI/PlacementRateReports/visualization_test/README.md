# Visualization test pipeline

A separate pipeline for testing the weekly dashboard + slide deck + email automation.
It does **not** change or replace the production email pipeline (`core.daily_runner`):

- Reads the report workbooks production already wrote (read-only). No SQL.
- Writes only under `visualization_test_runs/` (or `--run-root`).
- Emails only the `test_recipients` in `config.yml`, and only with `--send`.
  Director addresses appear in an `X-Intended-To` header for checking, never as recipients.

## What it builds, per audience (leadership + each director except Healthcare)

| File | Contents |
| --- | --- |
| `<year>-Dashboard-<programs>-<date>.xlsx` | Combined workbook: MSB overall tab first, one tab per program, helper data hidden |
| `<Month>_<Nth>_Placement_Rate_Visualization.pptx` | Title slide (date), MSB overall dashboard, then each program's dashboard |
| `email.eml` | The weekly email exactly as it would be sent |

Email subject: `Class of <year>: <programs> Weekly Placement Report <YYYY-MM-DD>`, where the
class is the report class (`core/academic_year.py`): October 2026-September 2027 = Class
of 2027. Every email attaches the workbook and the deck. Director workbooks show only the MSB tab and that director's programs
(other tabs are hidden, not deleted). Greeting names are in `config.yml`.

## Box upload

Uploads go to `Reports > Placement Reports > <year> Placement Reports > Placement
Visualizations`, picked by the current class year. Box needs each folder's upload email
address (folder > ... > Settings > Email Options > "Allow uploads to this folder via
email"); a shared link cannot be used. Put it under that year in `config.yml`:

```yaml
box_upload_addresses:
  "2027": ["Placement_Visualizations.xxxxxx@u.box.com"]
```

With `--send`, each audience's workbook and deck are emailed there every run, the same
way production archives to Box. Add the next year's folder address before October. Decks are renamed per audience for Box
(`October_16th_Placement_Rate_Visualization-BSFin.pptx`) so uploads don't collide.
Without `--send`, the upload is only saved as `box.eml`.

## Comparison table (M22:S38) and yearly rollover

- Columns N..S are always the current class and the five prior classes.
- Each October the table rolls automatically: the oldest class drops off S, every
  class moves one column right, the new class appears in N. Its August/September rows
  come from its own weekly history, which starts the last week of June.
- Monthly values are saved per audience in `<run-root>/state/<audience>.json`, so a
  finished class keeps all of its months after it rolls into history. The template's
  recorded values fill older years. Values dated after the report date are never used.
- The previous class's column is refreshed from that class's final report (through its
  last Friday of September) whenever that report is available.
- Weekly chart titles read `Weekly Class of <year> Placement Trend / Search Status`.
- The template is `visualization_test/assets/placement_dashboard_template.xlsx`
  (`template:` in config.yml). Change its formatting freely; keep the table at M22:S38
  with years in row 24, and keep the sheets and charts. Close it in Excel before a run.

## Cross-check before anything is sent

Before screenshots, `checks.py` compares each dashboard with the workbook it came from:
every weekly chart point (% placed and all status counts) against the report's Fridays,
the class named in titles/headers, this month's comparison cell against the report-date
rate, and the chart caches' point counts. Any difference fails that audience (saved in
`checks.json`), so no deck or email is made from it. A workbook without the report date
is refused as stale.

## Run

From `PlacementRateReports`, after `pip install -r visualization_test/requirements.txt`
(adds python-pptx; LibreOffice + Poppler are already required by production):

```bash
python -m visualization_test                         # today, previews only
python -m visualization_test --send                  # also email test_recipients
python -m visualization_test --date 2026-10-09 --source-root /path/to/reports \
    --run-root /tmp/vt-check                          # a past Friday from saved workbooks
python -m unittest discover -s visualization_test/tests -t . -v
```

`--audience Tanya` (repeatable) limits the run. `--send` needs SMTP_PASS, read the
same way as production (`ProgramReports/.env` on the Pi).

## In the Friday runner

With `friday_runner: true` (config.yml), `core.daily_runner` runs this pipeline on
Fridays after the report bundles (`core/friday_visualization.py`). It uses that run's
cross-checked workbook snapshots for the report class, so a dry run (which leaves the
rolling workbooks unchanged) still gets this week's data, plus the previous class's
final reports for its comparison column. Results are in `last-run.json` under
`visualizations` and in `runs/<preview|delivery>/visualizations/<date>/`. A dry run
writes previews only; a sending run emails `test_recipients` (and Box), never directors.
It also compares each director workbook with the leadership report from the same
morning and lists differences as warnings. No extra cron line is needed.

## On the Raspberry Pi

1. Copy the updated `PlacementRateReports/` (including `visualization_test/`) to the Pi.
2. `pip install -r visualization_test/requirements.txt` in the Pi's Python environment.
3. Put your own address in `test_recipients`, then try it by hand:
   `python -m visualization_test --audience Tanya` and inspect the run folder;
   then `--send` to receive the test email.

## Known issue found in the production updater (not changed here)

`Visualization/update_dashboard.py` writes the current-year table header (N24) as a
number. Excel requires table headers to be text, so the production dashboard workbooks
(including the shipped Sept 24 example) open only after Excel "repairs" all 15
comparison tables. This pipeline rewrites the headers as text, so its workbooks open cleanly.
