@echo off
cd /d C:\AeroTracker
git add -A
git diff --staged --quiet && echo Nothing to commit || git commit -m "fix: resource ID dedup + cache reset"
git pull --rebase origin master
git push origin master
git push origin master:main
echo.
echo Done - pushed to master and main.
pause
