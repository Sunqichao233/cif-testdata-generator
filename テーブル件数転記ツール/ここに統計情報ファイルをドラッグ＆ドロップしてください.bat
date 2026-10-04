@echo off
setlocal
cd /d "%~dp0"

echo.
echo  ====================================================
echo   Table statistics -^> CSV
echo  ====================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0script\extract.ps1" %1

echo.
pause
endlocal
