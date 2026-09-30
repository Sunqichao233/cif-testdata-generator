@echo off
rem ============================================================
rem  libInfo.txt 变换处理
rem
rem  用法1 : 把 libInfo.txt 拖到这个 run.bat 上
rem  用法2 : 把 libInfo.txt 放在同一个文件夹，直接双击 run.bat
rem
rem  流程 : 先显示调查结果 -> 按 Y 才真正变换
rem  输出 : <文件名>_converted.txt
rem ============================================================

setlocal
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS1=%~dp0convert.ps1"
set "TARGET=%~1"

if not exist "%PS1%" (
    echo [ERROR] convert.ps1 not found: %PS1%
    goto :fin
)

rem ---- 第1步：调查（不生成文件）----
if "%TARGET%"=="" (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Mode check
) else (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Path "%TARGET%" -Mode check
)
if errorlevel 1 goto :fin

rem ---- 第2步：确认后再变换 ----
echo.
set "YN="
set /p YN=Convert now? [Y/N] :
if /i not "%YN%"=="Y" (
    echo Canceled.
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
