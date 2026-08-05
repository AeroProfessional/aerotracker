@echo off
cd /d C:\AeroTracker
git add -A
git commit -m "Update"
git push origin HEAD:main
pause
