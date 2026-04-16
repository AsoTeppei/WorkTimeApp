@echo off
REM ============================================================
REM WorkTimeDB -- apply migrations (SQL Authentication)
REM   Runs db\migrations\*.sql in order against the live DB.
REM
REM   - Uses SQL auth (a login with db_owner or sysadmin rights).
REM   - Prompts for the password at run time so it is NOT stored
REM     on disk.
REM   - SchemaMigrations table tracks applied migrations, so
REM     running this twice is safe (idempotent).
REM   - For a fresh DB, run run_init.bat first.
REM
REM Before running:
REM   * SQL Server must be up
REM   * Run from an elevated cmd
REM   * Stop server.js first
REM   * Take a backup first (backup_db.bat)
REM ============================================================

setlocal
cd /d "%~dp0"

set SERVER=192.168.1.8\SQLEXPRESS
set DB=WorkTimeDB
set SQLUSER=yonekura

echo ================================================================
echo  Applying migrations
echo   Server: %SERVER%
echo   DB    : %DB%
echo   User  : %SQLUSER% (SQL auth)
echo ================================================================
echo.

REM ---- Prompt for password (not echoed is not possible in plain cmd,
REM      but at least we do not store it on disk) ----
set /p SQLPASS=Enter password for %SQLUSER%:
if "%SQLPASS%"=="" (
    echo [ERROR] Password is empty. Aborting.
    exit /b 1
)

echo.
echo ---- Connection test ----
sqlcmd -S %SERVER% -U %SQLUSER% -P "%SQLPASS%" -d %DB% -b -I -Q "SELECT SUSER_NAME() AS LoginName, DB_NAME() AS CurrentDB;"
if errorlevel 1 (
    echo [ERROR] Cannot connect to %SERVER% / %DB% as %SQLUSER%.
    echo         Check the password and try again.
    exit /b 1
)
echo.

REM ---- Bootstrap SchemaMigrations table if missing ----
REM Existing production DBs built from the old setup_database.sql
REM don't have this table, so we create it here. Idempotent.
echo ---- SchemaMigrations bootstrap ----
sqlcmd -S %SERVER% -U %SQLUSER% -P "%SQLPASS%" -d %DB% -b -I -Q "IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='SchemaMigrations') BEGIN CREATE TABLE SchemaMigrations (MigrationID NVARCHAR(100) NOT NULL CONSTRAINT PK_SchemaMigrations PRIMARY KEY, AppliedAt DATETIME2 NOT NULL DEFAULT GETDATE(), Description NVARCHAR(500) NULL); PRINT 'SchemaMigrations created.'; END ELSE PRINT 'SchemaMigrations already exists.';"
if errorlevel 1 (
    echo [ERROR] SchemaMigrations bootstrap failed.
    exit /b 1
)
echo.

REM ---- Apply each migration file in order ----
REM NOTE: .sql files are written in ASCII only (no Japanese comments)
REM       so sqlcmd can parse them with the default codepage without
REM       any encoding gymnastics.
for %%f in (migrations\*.sql) do (
    echo ---- %%~nxf ----
    sqlcmd -S %SERVER% -U %SQLUSER% -P "%SQLPASS%" -d %DB% -b -I -i "%%f"
    if errorlevel 1 (
        echo.
        echo [ERROR] %%~nxf failed. Aborting.
        exit /b 1
    )
    echo.
)

echo ================================================================
echo  Applied migrations:
echo ================================================================
sqlcmd -S %SERVER% -U %SQLUSER% -P "%SQLPASS%" -d %DB% -I -Q "SELECT MigrationID, AppliedAt, Description FROM SchemaMigrations ORDER BY AppliedAt"

echo.
echo Done.
endlocal
exit /b 0
