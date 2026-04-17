@echo off
REM =====================================================================
REM  Check-AutoStart.bat
REM  Reports the state of the "RdpPdfAutoPrintWatcher" Scheduled Task:
REM    - installed? enabled? next run? last run? last result?
REM    - also shows whether a watcher process is currently running.
REM  Exit codes:
REM    0  task installed and enabled
REM    40 task not installed
REM    41 task installed but disabled
REM    42 could not query task (access denied, etc.)
REM =====================================================================
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

set "TASK_NAME=RdpPdfAutoPrintWatcher"
set "QUEUE_ROOT=C:\RDPPrintQueue"

echo ================================================================
echo  Checking Scheduled Task: %TASK_NAME%
echo  User                   : %USERDOMAIN%\%USERNAME%
echo  Host                   : %COMPUTERNAME%
echo ================================================================
echo.

REM --- 1. Existence ---------------------------------------------------------
schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
if errorlevel 1 (
    echo [RESULT] Task NOT installed.
    echo          Run Install-AutoStart.bat to register it.
    endlocal & exit /b 40
)
echo [OK] Task is installed.
echo.

REM --- 2. Detailed status ---------------------------------------------------
echo --- schtasks /Query (key fields) ---
schtasks /Query /TN "%TASK_NAME%" /FO LIST /V > "%TEMP%\_rdppq_task.txt" 2>nul
if errorlevel 1 (
    echo [ERROR] Could not read task details.
    endlocal & exit /b 42
)

for %%F in (TaskName "Next Run Time" Status "Logon Mode" "Last Run Time" "Last Result" "Run As User" Author "Scheduled Task State" Comment) do (
    findstr /I /C:"%%~F" "%TEMP%\_rdppq_task.txt"
)
echo ------------------------------------
echo.

REM --- 3. Enabled / disabled? ----------------------------------------------
findstr /I /C:"Scheduled Task State: *Disabled" "%TEMP%\_rdppq_task.txt" >nul
if not errorlevel 1 (
    echo [WARN] Task exists but is DISABLED.
    echo        Enable it with:  schtasks /Change /TN "%TASK_NAME%" /ENABLE
    del "%TEMP%\_rdppq_task.txt" >nul 2>&1
    endlocal & exit /b 41
)

REM --- 4. Last run result decoded -------------------------------------------
for /f "tokens=2,* delims=:" %%A in ('findstr /I /C:"Last Result" "%TEMP%\_rdppq_task.txt"') do (
    set "LASTRESULT=%%B"
)
if defined LASTRESULT (
    set "LASTRESULT=!LASTRESULT: =!"
    echo [INFO] Last Result code : !LASTRESULT!
    if "!LASTRESULT!"=="0"         echo        -^> 0x0       : success / last run ended OK
    if "!LASTRESULT!"=="267011"    echo        -^> 0x41303   : task has not run yet
    if "!LASTRESULT!"=="267009"    echo        -^> 0x41301   : task is currently running
    if "!LASTRESULT!"=="267014"    echo        -^> 0x41306   : task was terminated by user
)
echo.

REM --- 5. Is a watcher process currently alive? -----------------------------
echo --- Running watcher processes ---
tasklist /FI "IMAGENAME eq powershell.exe" /V /FO LIST 2>nul | findstr /I /C:"RdpPdfAutoPrint" >nul
if errorlevel 1 (
    echo [INFO] No running powershell.exe process whose command line contains RdpPdfAutoPrint was found.
    echo        ^(This check is best-effort; it only inspects window titles.^)
) else (
    echo [OK] A watcher process appears to be running:
    tasklist /FI "IMAGENAME eq powershell.exe" /V /FO LIST | findstr /I /C:"PID" /C:"Image Name" /C:"Window Title"
)
echo.

REM --- 6. Recent log file ---------------------------------------------------
echo --- Newest log file under %QUEUE_ROOT%\logs ---
if exist "%QUEUE_ROOT%\logs\" (
    pushd "%QUEUE_ROOT%\logs" >nul
    for /f "delims=" %%L in ('dir /b /o-d /a-d 2^>nul') do (
        echo        %%L
        goto :_log_done
    )
    echo        ^<no log files yet^>
    :_log_done
    popd >nul
) else (
    echo        ^<logs folder missing^>
)

del "%TEMP%\_rdppq_task.txt" >nul 2>&1
echo.
echo [RESULT] Task installed and enabled.
endlocal & exit /b 0
