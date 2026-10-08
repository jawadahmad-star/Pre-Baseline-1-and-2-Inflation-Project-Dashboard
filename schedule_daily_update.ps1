# Registers a Windows scheduled task that runs update_dashboard.bat every day (default 08:00 and 20:00).
# Run once from PowerShell:   .\schedule_daily_update.ps1            (or pass -Times "09:00","18:00")
param([string[]]$Times = @("08:00","20:00"))
$bat = Join-Path $PSScriptRoot "update_dashboard.bat"
$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$bat`" < nul" -WorkingDirectory $PSScriptRoot
$triggers = $Times | ForEach-Object { New-ScheduledTaskTrigger -Daily -At $_ }
Register-ScheduledTask -TaskName "InflationProject_Dashboard_Refresh" -Action $action -Trigger $triggers -Force | Out-Null
Write-Host "Scheduled daily refresh at: $($Times -join ', ').  Remove with: Unregister-ScheduledTask InflationProject_Dashboard_Refresh"
