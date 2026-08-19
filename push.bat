@echo off
cd /d C:\AeroTracker
git add -A
git diff --staged --quiet && echo Nothing to commit || git commit -m "chore: local updates"
git pull --rebase origin master
git push origin master
echo Done.
pause
