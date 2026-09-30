@echo off
rem ============================================================
rem  libInfo.txt 変換処理
rem  この bat に libInfo のファイルをドラッグ＆ドロップしてください
rem ============================================================

setlocal
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS1=%~dp0script\convert.ps1"
set "TARGET=%~1"

if not exist "%PS1%" (
    echo [エラー] script\convert.ps1 が見つかりません
    goto :fin
)

rem ---- 1) 調査（ファイルは作成しない）----
if "%TARGET%"=="" (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Mode check
) else (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Path "%TARGET%" -Mode check
)
if errorlevel 1 goto :fin

rem ---- 2) 確認してから変換 ----
echo.
set "YN="
set /p YN=変換を実行しますか? [Y/N] :
if /i not "%YN%"=="Y" (
    echo 中止しました。
    goto :fin
)

if "%TARGET%"=="" (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Mode convert
) else (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Path "%TARGET%" -Mode convert
)

:fin
echo.
pause
endlocal
