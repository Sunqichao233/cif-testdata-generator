@echo off
setlocal
cd /d "%~dp0"

start "" powershell -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0script\keepalive.ps1"

endlocal
