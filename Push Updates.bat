@echo off
:: One-click: save any changes Claude made → send to GitHub
:: Run this whenever Claude tells you to push an update.

echo Pushing updates to GitHub...

:: Clear stale git locks (safe to delete — they're just leftover files)
del /f /q "%~dp0.git\HEAD.lock"    2>nul
del /f /q "%~dp0.git\index.lock"   2>nul

:: Stage all changed files, commit with timestamp, push
cd /d "%~dp0"
git add -A
git diff --staged --quiet && (
    echo Nothing new to push.
    pause
    exit /b 0
)
git commit -m "update %date% %time%"
git push origin master

if %ERRORLEVEL% == 0 (
    echo.
    echo Done! Changes are live on GitHub.
) else (
    echo.
    echo Something went wrong. Try running as Administrator.
)
pause
