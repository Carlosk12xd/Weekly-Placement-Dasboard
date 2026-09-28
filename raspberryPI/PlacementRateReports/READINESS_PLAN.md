# Next implementation stage: unattended execution

Recorded September 24, 2026 before this stage's code changes. The supplied
workbooks, SQL populations, director mapping and PDF layout remain the baseline.
This stage closes observed operational gaps in the September 23 package.

| Existing anchor (before this stage) | Change and intended behavior |
| --- | --- |
| core/database.py:5, load_environment | Return settings for the selected report directory without mutating os.environ; inherited process settings retain precedence. Database and mail receive the matching settings. |
| core/delivery.py:17,65, env_list/deliver | Accept explicit settings; verify addresses before live SMTP; continue to preserve saved envelopes, confirmed sends and uncertain outcomes. |
| core/pipeline.py:23, execute_bundle | Contain missing/corrupt per-bundle records so the next audience still runs; preserve corrupt evidence; include cohort and artifact/delivery status in outcomes. |
| core/pipeline.py:134, prepare_and_deliver | Pass the selected directory's settings explicitly into route preparation and SMTP instead of sharing environment mutations. |
| ProgramReports/create_program_reports.py:465 and placement_details.py:125 | Preserve the standalone OUTPUT_PATH setting through the returned settings mapping; scheduled builds still use their explicit staged path. |
| core/daily_runner.py:21, load_cohorts | Validate booleans, dates, IDs, grace days and human modes; a quoted false must never enable delivery. |
| core/daily_runner.py:39, preflight | Move workbook/template/converter failures to the affected bundle. Required PDF failures still hold human mail while valid workbook archives can proceed. |
| core/daily_runner.py:64, main | Record start, progress, failures and skip reasons atomically for each invocation; update last-report-run separately so ordinary skipped days retain the last report outcome. Keep one global lock. |
| New core/report_targets.py | Shared selected-audience/path enumeration for the runner and readiness command; retain the existing routing dictionary. |
| New core/run_records.py | Invocation IDs, timestamps and atomic execution records; overlapping invocations must not replace the active run's summary. |
| New core/doctor.py | Read-only setup check: configuration, selected inputs, Python/packages, converter paths, template fingerprint, storage access and setting names. No SQL connections, SMTP, reports, cron edits or activation. |
| New core/status.py | Inspect the latest invocation or last scheduled report, per-bundle artifacts and route outcomes; optional file-hash verification; never send or retry automatically. |
| New tests/test_operations.py | Exercise settings isolation, invalid configuration, early failures, missing/corrupt bundles, lock contention, status integrity and no-network readiness checks. |
| Deployment/status/source-map documents and release manifests | Replace stale descriptions, record exact current function anchors and distinguish offline/mocked evidence from target-Pi validation. |

Acceptance requires existing dashboard regressions to keep passing, deterministic
tests for the failure paths above, one real offline PDF/email-preview run through
the revised runner, and validation of the packaged file hashes. No live SQL,
messages, remote deployment, scheduling or delivery activation is part of this
workspace stage. The final Pi installation/pilot remains necessary.
