@echo off
REM ============================================================
REM WorkTimeDB -- initial setup (DESTRUCTIVE)
REM   Runs db\init\*.sql against an empty DB.
REM   WARNING: 01_schema.sql drops the existing WorkTimeDB.
REM            DO NOT run this against a production DB.
REM ============================================================

setlocal
cd /d "%~dp0"

set SERVER=192.168.1.8\SQLEXPRESS

echo ================================================================
echo  WorkTimeDB INITIAL SETUP (destructive)
echo   Server: %SERVER%
echo.
echo  The existing WorkTimeDB will be DROPPED. Continue?
echo ================================================================
set /p CONFIRM="Type YES to continue: "
if /i not "%CONFIRM%"=="YES" (
    echo Aborted.
    exit /b 0
)

echo.
echo ---- 01_schema.sql ----
sqlcmd -S %SERVER% -E -b -i "init\01_schema.sql"
if errorlevel 1 goto :fail

echo.
echo ---- 02_masterdata.sql ----
sqlcmd -S %SERVER% -E -b -i "init\02_masterdata.sql"
if errorlevel 1 goto :fail

echo.
set /p SAMPLE="Load sample data (37 users / 15 projects / 370 logs)? [Y/N]: "
if /i "%SAMPLE%"=="Y" (
    echo.
    echo ---- 03_sampledata.sql ----
    sqlcmd -S %SERVER% -E -b -i "init\03_sampledata.sql"
    if errorlevel 1 goto :fail
)

echo.
echo ================================================================
echo  Init complete. Next: run apply_migrations.bat
echo ================================================================
endlocal
exit /b 0

:fail
echo.
echo [ERROR] Initialization failed.
endlocal
exit /b 1
