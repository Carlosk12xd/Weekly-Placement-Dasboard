# Implementation status - September 24, 2026

**Implemented and verified offline; not yet installed or exercised on the Raspberry Pi.**
No real database query, SMTP email, Box upload, Zoho write or cron change occurred.
The shipped configuration defaults to staged previews with delivery disabled.

## Completed

- Import-safe daily runner with explicit cohort/date/mode, project-relative paths,
  audience selection, preflight, one-process lock, offline/no-send modes and retries.
- Existing report builders take a staged path and explicit report date; database
  connections close on query failures. Current SQL populations are retained.
- Dashboard data adapter preserves 45 native charts and their style/color parts,
  synchronizes weekly table/chart references/caches, reuses same-date observations,
  handles 2027 rollover and the two-year monthly axis, and highlights the current month.
- Status matching is by label, including the BSGSCM/BSHRM row-order exceptions.
  Identical duplicate dates collapse; conflicting duplicates fail; malformed old
  headers are recorded. A missing current observation cannot become a stale PDF.
- Archived 2026 report values refresh available historical comparisons without
  using observations after the current report date. Unknown months remain blank.
- Local LibreOffice/Poppler/Pillow PDF generation reuses the pinned app's branding,
  includes chart bounds in print areas, preserves helper-sheet data, checks page
  mapping/count/size, and uses isolated temporary profiles.
- Per-bundle snapshots, hashes, recovery copies, immutable dashboard state pointers,
  saved MIME previews and per-route SMTP journals. Required-PDF failures hold human
  delivery while valid XLSX archive routes remain available. Retry does not requery SQL.
- Leadership and ten director groups receive full-time PDF + XLSX in configured
  Friday modes. Healthcare retains XLSX only. Archive and Zoho routes remain XLSX only.
- The supplied recipient/program mapping is preserved. Finance's second internship
  detail section now uses year+2 records; Healthcare body dispatch no longer falls
  through to the standard program message.
- Current source map, deployment instructions, agent invariants, reproducible
  aggregate fixtures, regression tests and source/release fingerprints are included.
- Local readiness and run-status commands, invocation start/progress/failure
  records, skipped-day separation and independent handling of missing/corrupt
  audience inputs. Missing PDF executables no longer block unrelated bundles.
- Directory-scoped settings passed explicitly into SQL/mail without process
  environment mutation. Inherited settings retain precedence; direct program
  builders retain OUTPUT_PATH. Actual YAML boolean types are required.

## Executed evidence

| Check | Result |
| --- | --- |
| Automated regression tests | 29 passed on September 24; original dashboard/build checks plus 15 operational checks |
| June 26 reference reconciliation | All 165 status/count cells, 15 weekly rates and 15 monthly June rates agree |
| Full 2027 fixture preparation | 12 report bundles, including Healthcare |
| Full PDF generation | 11 PDFs, 50 pages total; 16-page leadership, 4-page paired and 3-page single-program outputs |
| Email preview validation | 25 MIME messages: 11 human PDF+XLSX, one Healthcare human XLSX, 13 XLSX-only archive/import messages |
| Visual inspection | All 16 leadership pages reviewed; chart/table content stays within page bounds |
| Failure/retry | Simulated converter failure holds human mail; retry reuses the same source hash |
| SMTP uncertainty | Mocked lost acknowledgement records uncertain state and blocks implicit resend |
| Finance detail routing | Distinct mocked 2028/2029 records appear in their separate sections |
| Input preservation | All 28 original production workbooks and the original June dashboard unchanged |
| Operational failure paths | Missing/corrupt inputs isolate one audience; configuration errors replace stale summaries; overlap and interruption recorded; skipped days preserve last report |
| Settings and mail | Mocked dotenv parsing, SQL connection and SMTP prove directory isolation, explicit settings and archive-only submission on PDF failure |
| Revised-runner real PDF check | Two bundles (Finance/Healthcare), one 3-page PDF and four MIME previews; all saved hashes verified by core.status |
| Local readiness check | Full 12-bundle offline prerequisites passed; no SQL/SMTP authentication attempted |

The full preview used the archived August 21, 2026 class-2027 report and the
available prior-cohort history. It is a fixture, not current live placement data.
The full 12-bundle/11-PDF rendering and visual review above occurred September 23.
September 24 re-ran all 29 tests and exercised the revised runner with Finance
and Healthcare; the unchanged renderer was not rerun for every audience.
The conversion workload measured 223.20 seconds in this environment, excluding
SQL and SMTP. This is not a Pi benchmark. Versions were Python 3.12.14,
LibreOfficeDev 26.8.0.0.alpha0, openpyxl 3.1.5, lxml 6.1.1, Pillow 12.3.0,
pypdf 6.10.0 and PyYAML 6.0.3. See VALIDATION_RESULTS.json for bundle-level output.

## Corrections and refinements to the original plan

1. BSGSCM and BSHRM have different ordering for three not-seeking rows. The default
   positional table in the plan must not be applied to them. Code maps labels.
2. The 2026 leadership report has history through August 31; the current 2027
   snapshot ends August 21. A source observation is bounded by the run's as-of date.
3. The source contains identical July 10 duplicate headers with suffixes and one
   malformed MISM header. The code handles them explicitly without guessing dates.
4. The chart template is edited through narrow OOXML operations. The data-report
   builders retain openpyxl. This preserves the supplied chart styling parts.
5. Persistent presentation state uses atomic pointer files to validated per-run
   workbooks, rather than overwriting one shared dashboard file in place.
6. Current-year single monthly points use markers, the highlighted month moves,
   and weekly axes request sparse labels while retaining
   all weekly data points; LibreOffice also adjusts spacing and rotation. This supports the complete 65-week tracking window.
7. Generated PDF filenames include the group/leadership label and report date.
   Presentation layout/filenames were refined after the full-bundle check; a fresh
   Finance rendering verifies the final filename/attachment path.

## Remaining target-Pi work

Install the overlay while preserving current workbooks and .env files; validate
Python/dependency availability, fonts, timezone, memory and runtime on the Pi.
Use `python -m core.doctor --cohort 2027` for local prerequisites and
`python -m core.status --last-report --verify` for saved run inspection.
Run the included offline fixture command, then a database-backed no-send run under
cron's account/environment. Verify the live recipients and archive settings, send
a controlled pilot to a chosen test inbox, and activate the existing cron through
`delivery_enabled: true` after review. DEPLOYMENT.md gives concrete commands.

Live SQL responses, credentials, SMTP acceptance, Box/Zoho ingestion and actual
cron execution remain unverified here. Internship and Healthcare chart layouts
were not supplied and are not implemented. PPTX export was not part of this PDF
integration. The existing summary formulas in report attachments remain formulas
for Excel to recalculate; chart generation uses validated stored history values.
