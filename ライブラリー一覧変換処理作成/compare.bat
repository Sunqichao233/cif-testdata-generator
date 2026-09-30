@echo off
rem ============================================================
rem  输出对账（脚本产物 vs 手工正解）
rem
rem  直接双击即可。会自动找同一文件夹里的：
rem      *_converted.txt   脚本生成的
rem      *_変換後.txt      手工做的正解
rem
rem  文件名不规则时，改成手动指定：
rem      compare.bat "脚本输出.txt" "手工正解.txt"
rem ============================================================

setlocal
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS1=%~dp0compare.ps1"

if not exist "%PS1%" (
    echo [ERROR] compare.ps1 not found: %PS1%
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
