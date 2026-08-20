@echo off
:: Creates a Windows Scheduled Task to run the spot-check report every weekday at 4pm.
:: Run this once — the task then runs automatically forever.

echo Setting up AeroTracker daily spot-check task...

schtasks /create ^
  /tn "AeroTracker - Daily Spot-Check Report" ^
  /tr "python C:\AeroTracker\spotcheck_report.py" ^
  /sc WEEKLY ^
  /d MON,TUE,WED,THU,FRI ^
  /st 16:00 ^
  /f

if %ERRORLEVEL% == 0 (
    echo.
    echo Done! The spot-check report will now run automatically every weekday at 4pm.
    echo It will save to the shared team folder and open in your browser.
) else (
    echo.
    echo Something went wrong. Try running this file as Administrator.
)

pause
