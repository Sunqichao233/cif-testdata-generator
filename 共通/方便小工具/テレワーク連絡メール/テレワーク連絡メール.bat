@echo off
setlocal
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0script\telework_mail.ps1"

if errorlevel 1 (
  echo.
  echo  Failed to start. See the message above.
  pause
)
endlocal
