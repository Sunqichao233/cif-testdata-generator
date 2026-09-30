@echo off
rem ============================================================
rem  libInfo.txt 変換処理
rem
rem  使い方1 : libInfo.txt をこの run.bat にドラッグ＆ドロップする
rem  使い方2 : libInfo.txt を同じフォルダに置いて run.bat をダブルクリック
rem
rem  流れ : まず調査結果を表示 → Y を押すと変換を実行
rem  出力 : <ファイル名>_converted.txt
rem ============================================================

setlocal
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS1=%~dp0convert.ps1"
set "TARGET=%~1"

if not exist "%PS1%" (
    echo [エラー] convert.ps1 が見つかりません: %PS1%
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
