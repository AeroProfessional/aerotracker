@echo off
cd /d C:\AeroTracker
del update_tracker.lock 2>nul
py update_tracker.py
pause
