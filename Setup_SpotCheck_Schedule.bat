@echo off
:: Recreates the scheduled task. Run as Administrator.

echo Fixing AeroTracker daily report task...

schtasks /delete /tn "AeroTracker - Daily Spot-Check Report" /f 2>nul

schtasks /create /tn "AeroTracker - Daily Spot-Check Report" /tr "py C:\AeroTracker\spotcheck_report.py" /sc WEEKLY /d MON,TUE,WED,THU,FRI /st 16:00 /f

if %ERRORLEVEL% == 0 (
    echo.
    echo Done! Report will run at 4pm every weekday.
) else (
    echo Something went wrong. Try running as Administrator.
)
pause
