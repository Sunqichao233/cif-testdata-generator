@echo off
setlocal
cd /d "%~dp0"

echo.
echo  ====================================================
echo   Event log (ev_app_*.log) -^> processing time CSV
echo  ====================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0script\extract_time.ps1" %*

echo.
pause
endlocal
