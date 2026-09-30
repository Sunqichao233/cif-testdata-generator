@echo off
rem ============================================================
rem  出力の突き合わせ（スクリプトの出力 vs 手作業の正解）
rem  ダブルクリックで実行できます
rem ============================================================

setlocal
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS1=%~dp0compare.ps1"

if not exist "%PS1%" (
    echo [エラー] compare.ps1 が見つかりません
    goto :fin
)

if "%~2"=="" (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
) else (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Mine "%~1" -Ref "%~2"
)

:fin
echo.
pause
endlocal
