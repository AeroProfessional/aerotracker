@echo off
:: Recreates the scheduled task. Run as Administrator.

echo Fixing AeroTracker daily report task...

schtasks /delete /tn "AeroTracker - Daily Spot-Check Report" /f 2>nul

schtasks /create /tn "AeroTracker - Daily Spot-Check Report" /tr "\"C:\AeroTracker\Run_Daily_Report.bat\"" /sc WEEKLY /d MON,TUE,WED,THU,FRI /st 16:00 /f /rl HIGHEST

if %ERRORLEVEL% == 0 (
    echo.
    echo Done! Report will run at 4pm every weekday.
) else (
    echo.
    echo Something went wrong. Try running as Administrator.
)
pause
