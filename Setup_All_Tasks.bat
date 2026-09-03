@echo off
:: Sets up both AeroTracker scheduled tasks.
:: Run this as Administrator when installing on a new computer.

echo Setting up AeroTracker scheduled tasks...
echo.

:: Delete existing tasks if they exist
schtasks /delete /tn "AeroTracker - Hourly Processing" /f 2>nul
schtasks /delete /tn "AeroTracker - Daily Spot-Check Report" /f 2>nul

:: Hourly processing task — runs every hour, catches up on missed runs when laptop turns on
schtasks /create /tn "AeroTracker - Hourly Processing" /tr "\"C:\AeroTracker\Run_Hourly.bat\"" /sc HOURLY /mo 1 /f /rl HIGHEST /it
:: Enable "run as soon as possible after a missed start" via PowerShell (schtasks doesn't support it directly)
powershell -Command "& { $t = Get-ScheduledTask -TaskName 'AeroTracker - Hourly Processing'; $t.Settings.StartWhenAvailable = $true; Set-ScheduledTask -InputObject $t }"

if %ERRORLEVEL% neq 0 (
    echo ERROR: Could not create hourly task. Make sure you are running as Administrator.
    pause
    exit /b 1
)
echo [OK] Hourly processing task created.

:: Daily 4pm report task — weekdays only
schtasks /create /tn "AeroTracker - Daily Spot-Check Report" /tr "\"C:\AeroTracker\Run_Daily_Report.bat\"" /sc WEEKLY /d MON,TUE,WED,THU,FRI /st 16:00 /f /rl HIGHEST

if %ERRORLEVEL% neq 0 (
    echo ERROR: Could not create daily report task.
    pause
    exit /b 1
)
:: Allow task to run on battery, catch up if missed, and run whether user is logged in or not
powershell -Command "& { $t = Get-ScheduledTask -TaskName 'AeroTracker - Daily Spot-Check Report'; $t.Settings.DisallowStartIfOnBatteries = $false; $t.Settings.StopIfGoingOnBatteries = $false; $t.Settings.StartWhenAvailable = $true; Set-ScheduledTask -InputObject $t }"
echo [OK] Daily 4pm report task created.

echo.
echo ============================================
echo All tasks set up successfully!
echo.
echo AeroTracker will now:
echo   - Process candidates every hour automatically
echo   - Save a report to the shared folder at 4pm weekdays
echo.
echo Your computer must be ON for these to run.
echo ============================================
pause
