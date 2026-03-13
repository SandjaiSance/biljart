@echo off
setlocal
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0update-site.ps1"
echo.
echo Klaar. Ververs nu index.html in je browser.
pause