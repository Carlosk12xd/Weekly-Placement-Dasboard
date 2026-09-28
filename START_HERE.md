# Start here

Updated September 24, 2026. This package implements the Friday full-time dashboard
automation and adds unattended-run diagnostics and failure isolation. It is a code
overlay; preserve the Pi's current report history and credentials.

1. Read raspberryPI/PlacementRateReports/CHANGES-2026-09-24.md and IMPLEMENTATION_STATUS.md.
2. Follow DEPLOYMENT.md for installation, local readiness, a no-send run and a pilot.
3. Use SOURCE_MAP.md for exact function anchors and DOCUMENTATION.md for current behavior.

The release defaults to delivery_enabled: false. No deployment or real email
occurred here. Production workbooks, .env files, run journals and cron configuration
are excluded. The original SQL populations and recipient dictionary are retained.

From PlacementRateReports after installing requirements:

```bash
python -m core.doctor --cohort 2027
python -m core.status --last-report --verify
```

For the credential-free fixture commands, see README.md. The archived example date
is August 21, 2026 for the class of 2027, not current live placement data.
