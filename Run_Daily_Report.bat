@echo off
:: Wrapper for the daily spot-check report.
:: Task Scheduler calls this bat, which sets the correct working directory
:: before running the Python script — so git pull and file paths work correctly.

cd /d "C:\AeroTracker"
py "C:\AeroTracker\spotcheck_report.py" >> "C:\AeroTracker\report_log.txt" 2>&1
