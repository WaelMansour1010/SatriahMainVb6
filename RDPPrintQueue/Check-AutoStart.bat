@echo off
REM =====================================================================
REM  Check-AutoStart.bat
REM  Health-check for the RdpPdfAutoPrintWatcher Scheduled Task.
REM
REM  Checks:
REM    1. Task installed?
REM    2. Task enabled?  Last run time / result?
REM    3. Watcher process alive? (via CIM command-line, not window title)
REM    4. Lock file present and PID still alive?
REM    5. Newest log file path
REM
REM  Exit codes:
REM    0  task installed and enabled
REM    40 task not installed
REM    41 task installed but disabled
REM    42 schtasks query failed
REM =====================================================================
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

set "TASK_NAME=RdpPdfAutoPrintWatcher"
set "QUEUE_ROOT=C:\RDPPrintQueue"
set "LOCK_FILE=%QUEUE_ROOT%\state\watcher.lock"

echo ================================================================
echo  RDP PDF Auto-Print Health Check
echo  Task    : %TASK_NAME%
echo  User    : %USERDOMAIN%\%USERNAME%
echo  Host    : %COMPUTERNAME%
echo  Time    : %DATE% %TIME%
echo ================================================================
echo.

REM -----------------------------------------------------------------------
REM 1. Scheduled Task existence
REM -----------------------------------------------------------------------
schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
if errorlevel 1 (
    echo [FAIL] Task NOT installed.
    echo        Run Install-AutoStart.bat to register it.
    endlocal & exit /b 40
)
echo [OK]   Task is installed.
echo.

REM -----------------------------------------------------------------------
REM 2. Task details via schtasks
REM -----------------------------------------------------------------------
echo --- Scheduled Task details ---
schtasks /Query /TN "%TASK_NAME%" /FO LIST /V > "%TEMP%\_rdppq_task.txt" 2>nul
if errorlevel 1 (
    echo [FAIL] Could not read task details (access denied?).
    endlocal & exit /b 42
)

for %%F in (TaskName "Next Run Time" Status "Logon Mode" "Last Run Time" "Last Result" "Run As User" "Scheduled Task State") do (
    findstr /I /C:"%%~F" "%TEMP%\_rdppq_task.txt"
)
echo.

REM Enabled check
findstr /I "Scheduled Task State" "%TEMP%\_rdppq_task.txt" | findstr /I "Disabled" >nul
if not errorlevel 1 (
    echo [WARN] Task is DISABLED.
    echo        Re-enable:  schtasks /Change /TN "%TASK_NAME%" /ENABLE
    del "%TEMP%\_rdppq_task.txt" >nul 2>&1
    endlocal & exit /b 41
)
echo [OK]   Task is enabled.
echo.

REM Last-result decode
for /f "tokens=2,* delims=:" %%A in ('findstr /I "Last Result" "%TEMP%\_rdppq_task.txt" 2^>nul') do set "LASTRESULT=%%B"
del "%TEMP%\_rdppq_task.txt" >nul 2>&1
if defined LASTRESULT (
    set "LASTRESULT=!LASTRESULT: =!"
    echo [INFO] Last Result code : !LASTRESULT!
    if "!LASTRESULT!"=="0"      echo        ^-^> success
    if "!LASTRESULT!"=="267011" echo        ^-^> task has not run yet ^(0x41303^)
    if "!LASTRESULT!"=="267009" echo        ^-^> task is currently running ^(0x41301^)
    if "!LASTRESULT!"=="267014" echo        ^-^> task was terminated ^(0x41306^)
    echo.
)

REM -----------------------------------------------------------------------
REM 3. Watcher process detection via CIM (command-line match, not window title)
REM    Works correctly even when the window is hidden.
REM -----------------------------------------------------------------------
echo --- Watcher process (CIM command-line search) ---
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "$procs = Get-CimInstance Win32_Process -Filter \"Name='powershell.exe' OR Name='pwsh.exe'\" | Where-Object { $_.CommandLine -match 'RdpPdfAutoPrint\\.ps1' }; if ($procs) { $procs | ForEach-Object { Write-Host ('[OK]   PID=' + $_.ProcessId + '  Start=' + $_.CreationDate + '  CMD=' + ($_.CommandLine -replace '\s+',' ')) } } else { Write-Host '[INFO] No PowerShell process found running RdpPdfAutoPrint.ps1.' }" 2>nul
echo.

REM -----------------------------------------------------------------------
REM 4. Lock file check — read PID from file, confirm process is alive
REM -----------------------------------------------------------------------
echo --- Lock file ---
if exist "%LOCK_FILE%" (
    set /p LOCK_CONTENT=<"%LOCK_FILE%"
    echo [INFO] Lock file content : !LOCK_CONTENT!
    REM Extract PID (first token before |)
    for /f "tokens=1 delims=|" %%P in ("!LOCK_CONTENT!") do set "LOCK_PID=%%P"
    if defined LOCK_PID (
        tasklist /FI "PID eq !LOCK_PID!" /NH >nul 2>&1
        if errorlevel 1 (
            echo [WARN] Lock file PID !LOCK_PID! is NOT running. Stale lock?
            echo        Delete manually: del "%LOCK_FILE%"
        ) else (
            echo [OK]   Lock PID !LOCK_PID! is alive.
        )
    )
) else (
    echo [INFO] No lock file present. Watcher is not running or exited cleanly.
)
echo.

REM -----------------------------------------------------------------------
REM 5. Newest log file
REM -----------------------------------------------------------------------
echo --- Newest log file ---
if exist "%QUEUE_ROOT%\logs\" (
    pushd "%QUEUE_ROOT%\logs" >nul
    set "FOUND_LOG="
    for /f "delims=" %%L in ('dir /b /o-d /a-d 2^>nul') do (
        if not defined FOUND_LOG (
            set "FOUND_LOG=%%L"
            echo [INFO] %QUEUE_ROOT%\logs\%%L
            echo        (last 5 lines:)
            powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
                "Get-Content '%QUEUE_ROOT%\logs\%%L' -Tail 5 | ForEach-Object { Write-Host ('        ' + $_) }" 2>nul
        )
    )
    if not defined FOUND_LOG echo [INFO] No log files yet.
    popd >nul
) else (
    echo [INFO] Logs folder missing: %QUEUE_ROOT%\logs
)
echo.

echo [RESULT] Check complete.
endlocal & exit /b 0
