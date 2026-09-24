@echo off
rem ============================================================
rem  CIF 测试数据生成（Tab 分隔 27 列）
rem
rem  用法 : run.bat [条数] [姓名类型] [日期YYYYMMDD]
rem    run.bat                    生成 1 条（和「正确的数据」对照用）
rem    run.bat 40000              生成 4 万条
rem    run.bat 100 kana           姓名用片假名
rem    run.bat 100 kanji 20260929 再指定文件名里的日期
rem
rem  姓名类型 : kanji(默认,汉字) / kana(片假名)  ※二选一，不混用
rem  条数上限 : 40000
rem
rem  想改文件内容（店铺代码、日期、某一列） -> 改 layout.csv
rem  想改文件名                             -> 改下面的 FILENAME
rem  详细说明见 README.md
rem ============================================================

setlocal
set "RC=0"

rem ============================================================
rem  配置区（要改的地方都在这里）
rem ============================================================

rem ---- 输出文件名 ----
rem   留空 = 自动用 CIF_<日期>.txt
rem   想固定文件名就写死，例如：set "FILENAME=U2KJD001"
set "FILENAME="

rem ---- 递增序号的起始值（默认从 1 开始）----
set "STARTNO="

rem ---- 输出目录 ----
set "OUTDIR=%~dp0output"

rem ---- 定义文件（一般不用改）----
set "LAYOUT=%~dp0layout.csv"
set "NAMES=%~dp0names.csv"

rem ============================================================

set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS1=%~dp0make_cif.ps1"

rem ---- 必要文件检查 ----
if not exist "%PS1%" (
    echo [ERROR] make_cif.ps1 not found: %PS1%
    goto :err
)
if not exist "%LAYOUT%" (
    echo [ERROR] layout.csv not found: %LAYOUT%
    goto :err
)
if not exist "%NAMES%" (
    echo [ERROR] names.csv not found: %NAMES%
    goto :err
)

rem ---- 输出目录，没有就建 ----
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

rem ---- 参数整理（没给就用脚本侧的默认值）----
set "OPT="
if not "%~1"==""        set "OPT=%OPT% -Count %~1"
if not "%~2"==""        set "OPT=%OPT% -NameMode %~2"
if not "%~3"==""        set "OPT=%OPT% -YMD %~3"
if not "%FILENAME%"=="" set "OPT=%OPT% -FileName "%FILENAME%""
if not "%STARTNO%"==""  set "OPT=%OPT% -StartNo %STARTNO%"

echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Layout "%LAYOUT%" -Names "%NAMES%" -OutDir "%OUTDIR%"%OPT%
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo *** FAILED - see the messages above. ***
    goto :fin
)

echo.
echo Output : %OUTDIR%
goto :fin

:err
set "RC=1"

:fin
echo.
pause
endlocal & exit /b %RC%
