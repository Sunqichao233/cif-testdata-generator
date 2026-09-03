@echo off
rem ============================================================
rem  CIF Test Data Generator - launcher for Amazon WorkSpaces
rem  (Japanese messages are printed by the PowerShell script itself)
rem
rem  Usage:
rem    run_generate.bat                      - interactive menu
rem    run_generate.bat -Count 500           - pass options to PowerShell
rem    run_generate.bat -DataProfile full -IncludeBoundary
rem    run_generate.bat -DataProfile matching
rem    run_generate.bat -ValidateOnly
rem ============================================================
setlocal
set "RC=0"

set "BASE_DIR=%~dp0"
set "PS_SCRIPT=%BASE_DIR%scripts\Generate-CifTestData.ps1"
set "CONFIG=%BASE_DIR%config\config.json"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS_EXE%" (
    echo [ERROR] PowerShell not found: %PS_EXE%
    goto :the_end
)
if not exist "%PS_SCRIPT%" (
    echo [ERROR] Script not found: %PS_SCRIPT%
    goto :the_end
)
if not exist "%CONFIG%" (
    echo [ERROR] Config not found: %CONFIG%
    goto :the_end
)

rem --- Unblock files copied into the WorkSpace (Mark of the Web) ---
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command ^
    "Get-ChildItem -LiteralPath '%BASE_DIR%' -Recurse -Include *.ps1,*.bat,*.json -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1

rem --- Arguments given: pass straight through ---
if not "%~1"=="" (
    set "PS_ARGS=%*"
    goto :run
)

rem ------------------------------------------------------------
rem  Interactive menu (double-click use)
rem ------------------------------------------------------------
:menu
cls
echo ==============================================================
echo   CIF Test Data Generator   ( CIF test data generator )
echo ==============================================================
echo.
echo   Output : %BASE_DIR%output
echo   Config : %CONFIG%
echo.
echo   [1] Standard   - minimal profile  (No.1 / No.3 / No.13 only)
echo   [2] Full       - all 13 fields filled
echo   [3] Full + boundary records (max length rows appended)
echo   [4] Matching   - collation patterns 2x3x3 = 18 records
echo   [5] Custom record count
echo   [6] Validate existing files only (no generation)
echo   [7] Open output folder
echo   [0] Exit
echo.
set "CHOICE="
set /p CHOICE=Select [0-7] :

if "%CHOICE%"=="1" ( set "PS_ARGS=-DataProfile minimal" & goto :run )
if "%CHOICE%"=="2" ( set "PS_ARGS=-DataProfile full" & goto :run )
if "%CHOICE%"=="3" ( set "PS_ARGS=-DataProfile full -IncludeBoundary" & goto :run )
if "%CHOICE%"=="4" ( set "PS_ARGS=-DataProfile matching" & goto :run )
if "%CHOICE%"=="5" goto :custom
if "%CHOICE%"=="6" ( set "PS_ARGS=-ValidateOnly" & goto :run )
if "%CHOICE%"=="7" ( start "" "%BASE_DIR%output" & goto :menu )
if "%CHOICE%"=="0" goto :the_end
echo Invalid selection.
timeout /t 2 >nul
goto :menu

:custom
set "REC_COUNT="
set /p REC_COUNT=Records per file [1-999999] :
echo %REC_COUNT%| findstr /r "^[1-9][0-9]*$" >nul
if errorlevel 1 (
    echo [ERROR] Please enter a positive number.
    timeout /t 2 >nul
    goto :menu
)
set "DP="
set /p DP=Profile [1=minimal / 2=full] :
if "%DP%"=="2" (set "PS_ARGS=-Count %REC_COUNT% -DataProfile full") else (set "PS_ARGS=-Count %REC_COUNT% -DataProfile minimal")
goto :run

rem ------------------------------------------------------------
:run
echo.
echo [INFO] powershell -File Generate-CifTestData.ps1 %PS_ARGS%
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %PS_ARGS%
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
    echo [DONE] Generation finished successfully. exit code=%RC%
) else (
    echo [FAILED] Errors were detected. exit code=%RC%
    echo          See the log under: %BASE_DIR%logs
)

:the_end
echo.
if "%~1"=="" pause
endlocal & exit /b %RC%
