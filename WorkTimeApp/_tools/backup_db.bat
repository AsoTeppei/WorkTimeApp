@echo off
REM ============================================================
REM YonekuraSystemDB daily backup (SQL Authentication)
REM   Intended to run from Windows Task Scheduler.
REM
REM   Output dir : %BACKUP_DIR%
REM   File name  : YonekuraSystemDB_YYYYMMDD_HHMMSS.bak
REM   Retention  : %RETAIN_DAYS% days (older files auto-deleted)
REM
REM IMPORTANT - SQL Login required:
REM   This script uses SQL Server authentication. A dedicated
REM   login named 'worktime_backup' must exist on the server
REM   with db_backupoperator role on YonekuraSystemDB.
REM
REM NOTE - Delayed expansion is intentionally DISABLED here so
REM   that '!' characters in the password survive the SET line.
REM
REM SECURITY NOTE:
REM   The password is stored in this file. Restrict NTFS
REM   permissions on this .bat to administrators only:
REM     icacls backup_db.bat /inheritance:r /grant:r Administrators:F
REM ============================================================

setlocal
cd /d "%~dp0"

REM ---- Settings ----
set SERVER=192.168.1.8\SQLEXPRESS
set DB=YonekuraSystemDB
set SQLUSER=worktime_backup
set SQLPASS=ChangeMe_Backup_2026!
set BACKUP_DIR=C:\WorkTimeDB_Backup
set RETAIN_DAYS=30

REM ---- Create backup folder if missing ----
if not exist "%BACKUP_DIR%" (
    mkdir "%BACKUP_DIR%"
    if errorlevel 1 (
        echo [ERROR] Cannot create backup folder: %BACKUP_DIR%
        exit /b 1
    )
)

REM ---- Build timestamp via PowerShell (locale independent) ----
for /f %%d in ('powershell -NoProfile -Command "(Get-Date).ToString('yyyyMMdd_HHmmss')"') do set STAMP=%%d

set BACKUP_FILE=%BACKUP_DIR%\%DB%_%STAMP%.bak
set LOG_FILE=%BACKUP_DIR%\backup.log

echo. >> "%LOG_FILE%"
echo ==== %DATE% %TIME% ==== >> "%LOG_FILE%"
echo Target : %SERVER% / %DB% (SQL auth as %SQLUSER%) >> "%LOG_FILE%"
echo Output : %BACKUP_FILE% >> "%LOG_FILE%"

echo.
echo Backing up %DB% to %BACKUP_FILE% ...

REM ---- BACKUP DATABASE ----
REM NOTE: CHECKSUM option validates every page's checksum as it is
REM       written to the backup, which is a stronger integrity check
REM       than the legacy RESTORE VERIFYONLY step. It also works
REM       with only db_backupoperator rights (no CREATE DATABASE
REM       permission on master is needed).
sqlcmd -S %SERVER% -U %SQLUSER% -P "%SQLPASS%" -b -Q "BACKUP DATABASE [%DB%] TO DISK = N'%BACKUP_FILE%' WITH INIT, CHECKSUM, STATS = 10, NAME = N'%DB% daily backup';" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] Backup failed. See %LOG_FILE% for details.
    echo [ERROR] backup failed >> "%LOG_FILE%"
    exit /b 1
)

echo [OK] %BACKUP_FILE%
echo [OK] backup + checksum success >> "%LOG_FILE%"

REM ---- Delete backups older than %RETAIN_DAYS% days ----
forfiles /P "%BACKUP_DIR%" /M "%DB%_*.bak" /D -%RETAIN_DAYS% /C "cmd /c del @path" >nul 2>&1

REM ---- Show latest 5 backups (uses :print_line subroutine
REM      to avoid needing delayed expansion) ----
echo.
echo ---- Latest backups ----
set CNT=0
for /f "delims=" %%f in ('dir /B /O-D "%BACKUP_DIR%\%DB%_*.bak" 2^>nul') do call :print_line "%%f"

endlocal
exit /b 0

:print_line
set /a CNT+=1
if %CNT% LEQ 5 echo   %~1
goto :eof
