@echo off
REM ============================================================
REM start_server.bat -- WorkTimeApp server launcher
REM
REM   Intended to run from Windows Task Scheduler at PC startup.
REM
REM Purpose:
REM   1. Avoid the "path with spaces" pitfall in Task Scheduler
REM      (C:\Program Files\nodejs\node.exe would otherwise get
REM      split at the first space).
REM   2. Redirect stdout + stderr to a log file so long-running
REM      diagnostics (24-hour auto-close, errors, crashes) are
REM      observable after the fact, not just while the cmd
REM      window happens to be open.
REM
REM Output log:
REM   D:\My Documents\...\WorkTimeApp\logs\server.log
REM
REM Notes:
REM   - The log file grows forever. Rotate it manually every few
REM     months if you care, or just delete it (it will be
REM     recreated next startup).
REM ============================================================

cd /d "D:\My Documents\各グループ\ソフト\WorkTimeApp"

if not exist logs mkdir logs

echo. >> logs\server.log
echo ============================================================ >> logs\server.log
echo  server.js starting at %DATE% %TIME% >> logs\server.log
echo ============================================================ >> logs\server.log

"C:\Program Files\nodejs\node.exe" server.js >> logs\server.log 2>&1

REM If node.exe exits (crash or manual stop), record it and return
REM the exit code so Task Scheduler's "restart on failure" option
REM can see the failure.
echo. >> logs\server.log
echo ============================================================ >> logs\server.log
echo  server.js exited at %DATE% %TIME% (exit code %ERRORLEVEL%) >> logs\server.log
echo ============================================================ >> logs\server.log

exit /b %ERRORLEVEL%
