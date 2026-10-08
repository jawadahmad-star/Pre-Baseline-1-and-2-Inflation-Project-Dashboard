# Inflation Project — Pre-Baseline Dashboard

Live: https://prebaseline.inflationproject.rs.org.pk (password protected)

## Daily update
1. Drop the newest SurveyCTO WIDE export (`.csv` or `.dta`) into `..\Data 1` (Pre-Baseline 1) or `..\Data 2` (Pre-Baseline 2). Newest file wins.
2. Double-click `update_dashboard.bat` (rebuilds, commits only dashboard files, pushes).
   `update_dashboard.bat nopush` rebuilds locally only. `schedule_daily_update.ps1` automates it.
3. With no real data present the dashboard shows clearly-labelled **dummy data** (`python build_dashboard.py --dummy` forces it).

## Targets
Default target per survey = shops with `sample_role = 1` in the prefill file (50 each). Override in `dashboard_config.json`.

## Password
Stored only in `password.local.txt` (git-ignored). Change: `python build_dashboard.py --set-password`, rebuild, push.
The interview rows embedded in `index.html` are de-identified and **AES-256-GCM encrypted** with that password.

## Privacy
Raw exports, prefill files, `dummy_data/` and the password file are git-ignored. Never commit names, phones or GPS.
