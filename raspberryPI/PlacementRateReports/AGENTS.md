# Placement report automation: agent operating record

This is the production project selected by the existing Raspberry Pi cron. Read
DOCUMENTATION.md and IMPLEMENTATION_STATUS.md before changes. IMPLEMENTATION_PLAN.md
is the original design baseline, not the current implementation status.

- Entry point: `python -m core.daily_runner`. Imports perform no SQL or SMTP work.
  Release default `delivery_enabled: false` means staged previews. `--offline`
  reads fixed reports, never SQL, and always disables every SMTP route.
- The no-send regression command is `python -m unittest discover -s tests -v`.
  The complete fixture command is in README.md. Do not use the archived TESTING
  project as a dry-run harness.
- Preserve the current `program_dict` in ProgramReports/email_program_reports.py.
  Historical conversational staff assignments are not authoritative routing.
- Placement = accepted / (accepted + actively seeking + not reported). No Recent
  Information Available and not-seeking categories are excluded. Zero denominator
  remains zero. Current rates must reconcile to the attached workbook.
- Map statuses by label: BSGSCM and BSHRM reorder three not-seeking categories.
  Never assume every NittyGritty block has identical row ordering.
- Normalize dates with optional `_2` suffixes. Collapse identical observations;
  conflicting same-date observations fail. Malformed historical headers are
  recorded in metadata, never guessed. The requested date must exist.
- The dashboard has 16 sheets, 45 native charts and 15 full-time pages. Keep all
  chart/style/color parts. Exclude helper sheets from printing without deleting
  their referenced data. Maintain chart formulas AND cached points/table metadata.
- The month axis runs August Y-1 through September Y. September 2026 is row 26
  for cohort 2027; September 2027 is row 38. Early June/July 2026 are weekly only.
  Roll headers and values together. Future/unknown months remain blank.
- Historical comparisons may use separately identified archived cohort reports;
  record their hashes and observation dates. Never use data after the run date.
- Current-year director values come from that director's saved XLSX; they are
  not assumed to equal leadership queries taken at another instant.
- Healthcare has no PDF profile. Finance's PDF is full-time only; the XLSX keeps
  the two distinct internship populations. Preserve Healthcare message/layout.
- The run ledger and snapshot hashes are authoritative for retry. Never rebuild
  SQL data for `--retry`, silently reuse a stale PDF, or automatically resend a
  route whose SMTP outcome is uncertain. SMTP acceptance is not proof of receipt.
- Keep one daily cron entry; month-end generation needs non-Friday runs. Human
  modes for 2027 are 0 and 2. Internal dashboard monthly updates also run in mode 1.
- After edits, update IMPLEMENTATION_STATUS.md, validation evidence and release
  hashes. Distinguish offline-tested, mocked, target-Pi-tested, and deployed work.
- Operational entry points: `python -m core.doctor` inspects prerequisites only;
  `python -m core.status --last-report --verify` reads journals and verifies hashes.
  Neither connects to SQL/SMTP or writes report state.
- Settings are returned by core.database.load_environment; it must never mutate
  os.environ. Pass the selected report group's mapping into database/mail calls.
  Explicit process settings override .env; offline runs never read .env files.
- Shared preflight validates configuration. Missing workbooks/ledgers/converters
  fail the affected bundle, not every audience. Required-PDF failures still hold
  human mail; valid archive routes can proceed. Never erase a corrupt ledger.
- Invocation records are separate from per-route SMTP ledgers. Persist progress
  before work; keep skipped-day summaries separate from last-report-run.json.
  A blocked attempt must not replace the active invocation's latest summary.
- Parse delivery_enabled, enabled and cover as actual YAML booleans. Quoted
  strings such as "false" are errors, never a way to activate delivery.
