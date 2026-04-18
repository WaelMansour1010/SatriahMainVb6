@echo off
REM =====================================================================
REM  Uninstall-AutoStart.bat
REM  Removes the Scheduled Task "RdpPdfAutoPrintWatcher" if it exists.
REM  Safe to run more than once: exits 0 when the task is already absent.
REM =====================================================================
setlocal ENABLEEXTENSIONS

set "TASK_NAME=RdpPdfAutoPrintWatcher"

REM --- Does the task exist? -------------------------------------------------
schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Task "%TASK_NAME%" is not installed. Nothing to do.
    endlocal & exit /b 0
)

echo [INFO] Deleting Scheduled Task: %TASK_NAME%
schtasks /Delete /TN "%TASK_NAME%" /F
if errorlevel 1 (
    echo [ERROR] schtasks /Delete failed with code %ERRORLEVEL%.
    echo         Re-run this .bat from an elevated command prompt if the
    echo         task was originally created elevated.
    endlocal & exit /b 30
)

REM --- Verify ---------------------------------------------------------------
schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
if not errorlevel 1 (
    echo [ERROR] Task still present after delete.
    endlocal & exit /b 31
)

echo [INFO] Task "%TASK_NAME%" removed successfully.
echo        NOTE: this does NOT stop an already-running watcher process.
echo              If one is active, close its console window or run:
echo                  taskkill /IM powershell.exe /FI "WINDOWTITLE eq *RdpPdfAutoPrint*"
endlocal & exit /b 0
