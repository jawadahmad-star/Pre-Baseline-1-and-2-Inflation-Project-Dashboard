@echo off
REM ================================================================
REM  Inflation Project dashboard - daily refresh
REM  1) rebuilds index.html from the newest export in ..\Data 1 / ..\Data 2
REM  2) commits ONLY the dashboard files  3) pushes to GitHub Pages
REM  Double-click to run.  Add "nopush" to rebuild without pushing.
REM ================================================================
cd /d "%~dp0"
python build_dashboard.py || (echo. & echo BUILD FAILED & pause & exit /b 1)
if /i "%1"=="nopush" (echo Built only - not pushed. & pause & exit /b 0)

REM safety: refuse to continue if any raw data file is staged
git add index.html build_dashboard.py generate_dummy_data.py dashboard_template.html dashboard_config.json README.md update_dashboard.bat schedule_daily_update.ps1 requirements.txt .gitignore .nojekyll CNAME
git diff --cached --name-only | findstr /i /r "\.csv \.dta \.xls password" >nul && (echo. & echo STOP: a data/password file is staged. & git reset >nul & pause & exit /b 1)

git commit -m "Daily data refresh: %date% %time%" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
git push origin main || (echo. & echo PUSH FAILED & pause & exit /b 1)
echo.
echo Done. Live at https://prebaseline.inflationproject.rs.org.pk
pause
