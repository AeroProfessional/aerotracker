@echo off
cd /d C:\AeroTracker
del /f "C:\AeroTracker\.git\index.lock" 2>nul
del /f "C:\AeroTracker\.git\HEAD.lock" 2>nul
git rebase --abort 2>nul
python -c "import json; open('tracker_processed.json','w').write(json.dumps({'names':[],'ids':[],'resource_ids':{}}))"
git add -A
git diff --staged --quiet && echo Nothing to commit || git commit -m "fix: workflow force-push + dedup"
git push --force origin master
git push --force origin master:main
echo.
echo Done.
pause
