@echo off
:: One-click: get the latest version of everything from GitHub
:: Run this to make sure your local copy is up to date.

echo Pulling latest from GitHub...

del /f /q "%~dp0.git\HEAD.lock"   2>nul
del /f /q "%~dp0.git\index.lock"  2>nul

cd /d "%~dp0"
git pull origin master

if %ERRORLEVEL% == 0 (
    echo.
    echo Done! You have the latest version.
) else (
    echo.
    echo Something went wrong. Check your internet connection.
)
pause
