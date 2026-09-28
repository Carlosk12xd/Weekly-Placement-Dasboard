# Deploying the Friday dashboard automation

The delivered overlay contains code, the dashboard template, and isolated test
fixtures. It deliberately excludes production report workbooks, environment files,
run journals and the cron configuration. Keep the Pi's current report history and
credentials. No live deployment or email was performed during development.

## 1. Check the Pi baseline and install dependencies

Compare the live source with the original hashes in IMPLEMENTATION_PLAN.md before
copying the overlay. Preserve any changes made on the Pi after that snapshot.
Back up changed code/configuration and current workbooks before rollout. Pause the
one existing placement cron entry while replacing code; restore that same entry.

The code requires Python 3.10+ and Linux for the scheduler's `fcntl` overlap lock.
Use the virtual environment already activated by the existing cron entry:

```bash
cd /home/bccdatapi/raspberryPI/PlacementRateReports
source /home/bccdatapi/raspberryPI/raspberrypi/bin/activate
python --version
python -m pip install -r requirements.txt
```

On Raspberry Pi OS/Debian, install the converter, rasterizer and fonts:

```bash
sudo apt-get update
sudo apt-get install libreoffice-calc poppler-utils fonts-dejavu-core
soffice --version
pdftoppm -v
timedatectl
```

Confirm that cron's timezone gives the intended 10:00 start. This is a preparation
start time; sending follows generation and conversion. Pi performance has not
been measured in this workspace. The renderer uses a 300-second timeout per
conversion/rasterization subprocess and processes reports sequentially.

Existing `.env` files stay in LeadershipReport and ProgramReports. Database builds
require DB_HOST, DB_USER, DB_PASSWORD, DB_NAME. SMTP requires SMTP_PASS. Leadership
delivery reads TO_ADDRS, CC_ADDRS, BCC_ADDRS and Zoho_Box. Box routes remain in the
cohort YAML; director routing remains in the existing program dictionary. Verify
their live values locally without copying secrets into logs or support documents.
Settings now stay scoped to their report directory. Explicit process environment
variables override `.env` values, so check cron's inherited settings if the two
report groups intentionally use different credentials. Neither group loads values
into the other group's process environment. Standalone program builders retain
their `OUTPUT_PATH` setting; the scheduled runner supplies explicit staged paths.

Run the read-only setup check under the cron account and environment:

```bash
python -m core.doctor --cohort 2027
```

It reports package availability, input workbooks, template fingerprint, converter
paths, font availability, timezone, disk space, storage permissions and missing
setting names. It displays no passwords or recipient addresses and does not query
SQL, authenticate to SMTP, generate reports, write files or change cron. A pass
does not establish working credentials/connectivity or acceptable Pi performance.
`--cohort` can inspect a cohort outside its active dates; calendar/eligibility
warnings explain why the runner would skip it. Quoted boolean values in YAML are
rejected: use `delivery_enabled: false`, not `delivery_enabled: "false"`.

## 2. Verify a complete offline run

```bash
python -m unittest discover -s tests -v
python -m core.doctor --offline --cohort 2027 --date 2026-08-21 \
  --source-root tests/fixtures --run-root /tmp/bcc-fixture-preview
python -m core.daily_runner --offline --cohort 2027 --date 2026-08-21 \
  --source-root tests/fixtures --run-root /tmp/bcc-fixture-preview
```

This uses the archived August 21 fixture date. It does not claim to reconstruct
today's database. The expected result is 12 XLSX bundles, 11 full-time PDFs and 25
saved `.eml` messages. Leadership has 16 PDF pages including the cover; paired
directors have 4 and single-program directors have 3. Healthcare has XLSX only.
Fixtures contain aggregate history; individual-placement detail cells and unused
shared strings were removed. They are separate from production workbook paths.

Review `last-run.json`, PDF pages and email attachments. Confirm that the current
month is highlighted and the 2026 historical series uses available archived
observations; unavailable September data remains blank. The complete production
workload, memory use, font/layout fidelity and serialized email sizes must pass on
the actual Pi before enabling delivery.

## 3. Prepare a current database-backed preview

On a Friday or month end:

```bash
python -m core.daily_runner --cohort 2027 --dry-run
```

This queries the live database and writes staged copies and email previews. It
does not update rolling workbooks or call SMTP, including Box/Zoho. It uses today's
date because detail queries still use SQL NOW(). A historical label is accepted
only with `--offline` or `--retry`.

The shipped `delivery_enabled: false` is a release safeguard: the unchanged cron
command would also prepare previews until that setting is activated. It suppresses
ALL delivery routes, including archives, and prevents rolling-workbook writes.

## 4. Send a controlled pilot, then activate the existing job

Once the previews pass and a test inbox is selected, a live pilot command is:

```bash
python -m core.daily_runner --cohort 2027 --audience leadership \
  --send --test-recipient YOUR_VERIFIED_TEST_ADDRESS
```

Replace the placeholder before execution. This command sends real email. Every
route, including Box and Zoho, is redirected to that test inbox. Pilot sends use
their own journal/state and do not promote data into production rolling workbooks.
The code does not resolve or verify the identity of a supplied test address.

For production activation set `delivery_enabled: true` in config/cohorts.yml after
reviewing the live recipient/destination settings. Keep the existing single daily
cron command `python3 -m core.daily_runner`. Do not add a second sending cron job.
The 2027 human modes `[0, 2]` deliver Friday reports and Friday month-end reports;
non-Friday month ends refresh reports/dashboard history and preserve archive routes.

## 5. Retry and recovery

```bash
python -m core.status --last-report --verify
```

Add `--run-root /tmp/bcc-fixture-preview` for fixture output. Without `--last-report`,
status reads the most recent invocation, including skipped days. `--json` emits
machine-readable results. Hash checking covers report/dashboard/PDF/metadata and
saved emails without changing files or sending anything. Completed preview runs
are explicitly labeled previews; `sent` means SMTP accepted a route.

Every parsed runner invocation records start, per-bundle progress and completion
in `runs/invocations/<run-id>.json`. `last-run.json` tracks the latest invocation;
`last-report-run.json` retains the most recent eligible report invocation after
ordinary skipped days. Early configuration/import failures are recorded with a
nonzero exit. A busy lock records a separate blocked invocation without replacing
the active summary. If the process is killed or power fails, a record may remain
`running`: this means completion was not recorded, not proof the process is alive.
If the run directory itself is unwritable, the failure is logged to stderr and a
record cannot be guaranteed. Invalid CLI syntax is rejected before recording.

Missing workbooks and damaged bundle ledgers now fail only that audience. The
damaged ledger is retained for inspection. Missing converters/templates hold PDF
human mail while valid workbook archive routes and independent bundles proceed.
Invalid shared configuration/dependency imports still stop the invocation.

Run records live under `runs/<preview-or-delivery>/<date>/<cohort>/<audience>/`.
`manifest.json` records snapshot hashes, presentation metadata, PDF status and
per-route delivery state. Pointer files under `runs/state/` select the latest
validated leadership dashboard without overwriting older dashboard snapshots.
Retain run directories referenced by these pointers; no retention cleanup is automatic.

For a failed live run with a completed source snapshot:

```bash
python -m core.daily_runner --cohort 2027 --date YYYY-MM-DD \
  --audience leadership --retry --send
```

Retries use the saved report and existing successful artifacts. Confirmed SMTP
acceptances are skipped. Failed PDF generation holds the human email but allows
valid XLSX archive routes to proceed. A missing/corrupt saved artifact fails instead
of silently regenerating it with different values. If no completed report snapshot
exists, fix the build failure and run normally for today's date.

`submitting` or `uncertain` delivery states require checking what the SMTP server
accepted. `--resend-uncertain` explicitly permits another submission after that
review; it may duplicate a message already accepted. No exactly-once delivery
claim is made. A changed visualization profile requires a fresh preview directory;
never change the production run root merely to bypass a sent/uncertain journal.

To stop sending, set `delivery_enabled: false`. `--dry-run` also overrides an enabled
configuration for one invocation. To intentionally fall back to XLSX-only delivery,
set `enabled: false` in config/visualizations.yml and review the affected run state;
do not disguise conversion errors as a successful dashboard report.

Restore code/configuration when needed, preserving workbook history accumulated
since deployment. The snapshot files in this package are test fixtures and must
never replace current production report workbooks.
