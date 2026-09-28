# Current report automation

Updated September 24, 2026. This document describes implemented code. The original
planning baseline is retained in IMPLEMENTATION_PLAN.md. See IMPLEMENTATION_STATUS.md
for which checks were executed and which still require the actual Pi.

## Execution

`core.daily_runner.main()` is the only scheduled entry point. Imports do not load
credentials, query SQL, mutate workbooks or send emails. The existing daily cron
is retained: mode 0 is Friday, mode 1 is month end, mode 2 is both, and other days
exit. Cohort eligibility still uses inclusive configured start/graduation-plus-
grace dates. Class 2027 ends September 19, 2027 under the supplied configuration.

Configuration parses `human_modes` with the cohort fields. Class 2027 uses `[0,2]`;
class 2026 retains `[0,1,2]` while eligible. `delivery_enabled: false` stages previews.
It can be activated in the YAML after deployment verification or overridden for
one explicit live run with `--send`. `--dry-run` always wins over the YAML.
`--offline` prohibits `--send` and reads the saved source tree instead of SQL.

The runner holds one nonblocking Linux file lock across preparation and delivery.
It validates shared configuration, then processes selected leadership and director
bundles sequentially. Missing reports, corrupt ledgers, templates and converter
executables are handled in the affected bundle. One failed bundle is
recorded without presenting it as a success or accidentally attaching another
bundle's artifacts. The process returns nonzero if any selected bundle fails.

## Files and responsibilities

| Component | Implemented responsibility |
| --- | --- |
| core/daily_runner.py | Calendar, cohort configuration, CLI, overlap lock, preflight, outcome summary |
| core/database.py | Return isolated settings for the selected directory; connect only for a database build |
| core/report_targets.py | Shared enumeration of selected audiences and existing workbook paths |
| core/run_records.py | Invocation timestamps, progress, outcome and latest-report summaries |
| core/doctor.py | Read-only local prerequisites/settings check; no SQL/SMTP authentication |
| core/status.py | Read-only invocation/bundle inspection and optional artifact-hash verification |
| core/pipeline.py | Stage/build/validate each report, prepare presentation, construct routes and deliver |
| core/report_artifacts.py | Explicit options, atomic copies/JSON and hash verification |
| core/delivery.py | MIME attachments, saved previews, SMTP and per-route delivery journal |
| LeadershipReport/update_overall_report.py | Existing SQL/table updates with explicit date/path and connection cleanup |
| ProgramReports/create_program_reports.py | Existing program aggregate updates with explicit date/path |
| ProgramReports/placement_details.py | Detail write into that same staged report; corrected Finance second-year source |
| LeadershipReport/email_overview_report.py | Import-safe leadership routing facade |
| ProgramReports/email_program_reports.py | Preserved director dictionary and import-safe routing facade |
| ProgramReports/email_templates.py | Attachment-aware text and correct Healthcare dispatch |
| Visualization/update_dashboard.py | Source-table validation, weekly/monthly updates, rollover, historical comparison and cache refresh |
| Visualization/xlsx_package.py | Narrow XML changes preserving charts, styles, colors and other unmodified ZIP parts |
| Visualization/render_excel.py | Page profiles, isolated LibreOffice conversion, Poppler and PDF validation |
| Visualization/branding.py | Adapted cover/sidebar/cropping functions from the pinned Excel-to-PDF app |

SOURCE_MAP.md supplies current function line numbers. Names remain more stable
than lines after edits. The report SQL filters and recipient tuples were retained.
The data builders still use openpyxl; the chart template uses narrow OOXML edits
to avoid rewriting unsupported style parts through a generic workbook save.

Report-specific `.env` files are parsed into separate mappings. Inherited process
settings take precedence, but reading one directory never mutates os.environ or
populates settings for the next group. Explicit mappings reach the DB connector,
recipient parser and SMTP login. Offline fixture runs never read `.env`. Program
builders invoked directly still honor their directory's OUTPUT_PATH setting.
Configuration booleans are type-checked; strings such as "false" fail validation.

## Source and metric contracts

The attached XLSX snapshot supplies every current-cohort value in its PDF.
Leadership uses `FT_total_wh` plus the 14 program full-time history tables. A
director copy uses `Class2` plus that director's program history tables. Queries
can occur at different times across bundles; this does not provide one shared
SQL transaction across all recipients.

Map source and destination rows by status label. BSGSCM and BSHRM order the
postponing/returning/starting-business rows differently from most blocks. Current
rates must reconcile to accepted / (accepted + seeking + not reported), at the
source's stored percentage precision. No Recent Information Available is excluded;
zero denominator remains zero. Overall is read as pooled counts/rate, not an
average of program rates. Missing observations never become invented zeros.

Source headers support normal Excel dates, common date strings and historical
`_2`/`_3` suffixes. Identical observations on one date are collapsed. Conflicting
duplicates fail. Empty `Column1` placeholders are ignored. Other malformed headers
are listed in metadata; the supplied MISM history includes `10/31/20252`, which is
not guessed. The requested current date must exist in every selected source table.

The existing cohort/population filters, Finance internship years, Healthcare
queries and append-only Zoho helper behavior are unchanged. Visualization reads
fresh report tables, not `ZOHO_DATA`. The summary sheet's existing total formulas
remain formulas in the attached XLSX and recalculate when opened in Excel; the
PDF uses validated stored history values and does not depend on those formula
caches. SQL detail queries still use NOW() - INTERVAL 7 DAY: backdating a live build
is therefore rejected. Mocked database tests do not establish live SQL correctness.

## Dashboard preparation

The immutable June template has 16 sheets and 45 charts. The updater preserves
all 15 presentation pages and the helper sheet. It validates sheet identities,
chart count and destination status labels before writing. The default template's
SHA-256 is pinned; changing it requires revisiting the contract and fingerprint.

A new 2027 dashboard shifts the previous cohort year headers AND comparison
values together. New-cohort weekly history is populated only from observations
within the configured window. Weekly points use actual Fridays; historical
baseline points in an existing same-cohort dashboard are retained. Missing Fridays
are not manufactured. Each update synchronizes 30 weekly tables, category/value
references and chart caches. Same-date updates reuse their point.

Each page's current cohort occupies column N in `M24:S38`. The month cycle is
August of Y-1 through September of Y. September 2026 maps to row 26 for 2027;
September 2027 maps to row 38. Early June/July 2026 observations remain weekly
because this template's monthly axis starts in August 2026. The current-month
cell is highlighted automatically; the old June highlight is cleared.

The current month updates on each run and is labeled with the report date. An
available actual month-end observation finalizes a previously provisional month.
Otherwise the last recorded date remains explicit in JSON metadata. On initial
creation, available past months are seeded from the report history. Future and
unavailable months remain blank. The last eligible Friday in September 2027 is
September 17; this configuration cannot promise a September 30 observation.

`historical_reports` can supply archived prior-cohort reports. Available monthly
observations refresh the corresponding previous-year column with recorded source
hashes/dates. No observation later than the current report date is used. Template
values remain where an archived source does not provide that month; these are
preserved supplied observations, not independently certified measurements. The
2026 report has observations through August 31, while the supplied current 2027
report ends August 21. The August 21 preview therefore uses only 2026 observations
available by August 21 and keeps September 2026 blank.

A persistent leadership dashboard is selected by an atomic pointer under
`runs/state/`. It refers to an immutable dashboard/metadata pair in a saved run.
Director dashboards are isolated copies populated from their own source reports.
A new yearly cohort beyond the currently configured 2027 profile requires its
appropriate previous-cohort template/configuration; the code refuses multi-year
relabeling of the June 2026 template.

## Rendering

Only presentation pages print. The helper sheet stays in the workbook because
charts reference it. `plotVisOnly=0` preserves plotted data on hidden helper sheets.
The explicit print area is B1:T39, accounting for drawings beyond populated cells.
Weekly axes request sparse date labels while retaining all observations; LibreOffice may select its own readable spacing/rotation.
Page order is selected deliberately and the converter must produce exactly one
page per selected sheet before any labels are attached.

The path is dashboard copy -> LibreOffice PDF -> Poppler images -> app-style
cover/sidebar -> final PDF. Each conversion has a private writable LibreOffice
profile and temporary directory. Pages are composed sequentially to reduce memory.
A current-year marker makes a single monthly observation visible. All 45 chart
objects and chart style/color parts are preserved. The output remains an image-
based PDF, like the app's branded output; editable PowerPoint is outside this release.

Leadership prints 15 content pages plus optional cover. Four paired directors
print overall plus their two programs; six single-program directors print overall
plus their one program. Healthcare has no PDF profile and retains its workbook.
Internship charts were not present in the supplied template and are not invented.

Default limits are 18 MB per PDF and 25 MB per serialized email. MIME size is
checked after encoding. These are configuration limits, not claims about every
mail gateway. Pi-specific fonts, speed and memory still require target testing.

## Artifact and delivery lifecycle

Invocation summaries under `runs/invocations/` record start, each bundle result,
completion, failure or skip reason. `last-run.json` is refreshed before work and
after progress. `last-report-run.json` retains eligible report outcomes across
skipped days. A blocked overlapping invocation gets its own record but cannot
overwrite the active summary. Abrupt termination can leave `running` without a
finish timestamp; this records incomplete work, not process liveness. CLI usage
errors and an unwritable run directory cannot guarantee a saved invocation.

`core.doctor` checks local prerequisites only and prints setting presence without
values. `core.status` inspects recorded results and optional hashes; it cannot
authenticate, send, retry or repair a record. Neither command verifies actual
receipt, Box/Zoho ingestion, SQL connectivity or target-Pi runtime.

Each bundle records identity (cohort/date/mode/audience), hashes and routes in
`manifest.json`. Completed XLSX snapshots are reused on retries without SQL. A live
production build promotes validated report history atomically, with a recovery
copy; dry runs and test-recipient pilots do not overwrite rolling reports. Dashboard
history and email state are separate. A saved XLSX is not proof of a sent message.

Friday PDFs are attached beside the matching XLSX. Box and the separate leadership
Zoho import message retain XLSX-only payloads. The original director dictionary
and CC/BCC behavior are retained. Healthcare wording is selected with exclusive
branches; the old future-manual-visualization promise is removed.

SMTP opens only after a bundle's artifacts and messages are ready. A failed PDF
holds that supported human route while valid archive routes may continue. An
unsupported Healthcare PDF is not a failure. Non-Friday month ends update the
internal presentation history and archives without enabling extra 2027 human mail.

Routes record prepared, submitting, sent or uncertain states. Sent means SMTP
accepted the submission, not that a person received it. A lost acknowledgement or
partial refusal is recorded as uncertain and requires review before the explicit
resend option. Confirmed submissions are skipped on retry. Snapshot/profile/hash
mismatches fail instead of being silently bypassed. Saved route bytes and envelope
lists are retained, including BCC recipients only in the envelope record.

## Verification and remaining deployment work

See IMPLEMENTATION_STATUS.md and VALIDATION_RESULTS.json for the executed checks.
The package's fixtures are reproducible aggregate snapshots; placement-detail data
was cleared and unreferenced shared strings were removed from fixture copies.
The production workbooks and original attachments were not edited during offline
verification. Review DEPLOYMENT.md for target-Pi validation and controlled activation.
