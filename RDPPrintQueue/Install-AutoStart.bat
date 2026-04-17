@echo off
REM =====================================================================
REM  Install-AutoStart.bat
REM  Registers the RDP PDF auto-print watcher as a per-user Scheduled Task
REM  that runs at logon of the CURRENT interactive user.
REM
REM  Task name  : RdpPdfAutoPrintWatcher
REM  Trigger    : At logon of %USERDOMAIN%\%USERNAME%
REM  Action     : C:\RDPPrintQueue\Start-RdpPdfAutoPrint.bat
REM  Start in   : C:\RDPPrintQueue  (applied via "cmd /c cd /d ...")
REM  Run level  : Limited (NOT elevated; NOT SYSTEM) - printers are per-user.
REM  Idempotent : /F overwrites any existing task with the same name.
REM =====================================================================
setlocal ENABLEEXTENSIONS

set "TASK_NAME=RdpPdfAutoPrintWatcher"
set "QUEUE_ROOT=C:\RDPPrintQueue"
set "LAUNCHER=%QUEUE_ROOT%\Start-RdpPdfAutoPrint.bat"

REM --- Sanity checks --------------------------------------------------------
if not exist "%QUEUE_ROOT%\" (
    echo [ERROR] Queue root not found: %QUEUE_ROOT%
    echo         Deploy the RDPPrintQueue folder first.
    exit /b 20
)
if not exist "%LAUNCHER%" (
    echo [ERROR] Launcher not found: %LAUNCHER%
    exit /b 21
)

REM Refuse to install under SYSTEM / elevated service accounts: the whole
REM point of this task is the per-user interactive printer context.
if /I "%USERNAME%"=="SYSTEM"              goto :bad_user
if /I "%USERNAME%"=="LOCAL SERVICE"        goto :bad_user
if /I "%USERNAME%"=="NETWORK SERVICE"      goto :bad_user

REM --- Build the action -----------------------------------------------------
REM schtasks has no native "Start in" flag, so we chain via cmd.exe:
REM   cmd /c "cd /d C:\RDPPrintQueue && Start-RdpPdfAutoPrint.bat"
set "ACTION=cmd /c \"cd /d %QUEUE_ROOT% ^&^& \"%LAUNCHER%\"\""

echo [INFO] Installing Scheduled Task
echo        Name     : %TASK_NAME%
echo        User     : %USERDOMAIN%\%USERNAME%
echo        Trigger  : At logon
echo        Action   : %LAUNCHER%
echo        Start in : %QUEUE_ROOT%
echo.

REM --- Create / replace the task -------------------------------------------
schtasks /Create ^
    /TN "%TASK_NAME%" ^
    /SC ONLOGON ^
    /RU "%USERDOMAIN%\%USERNAME%" ^
    /RL LIMITED ^
    /DELAY 0000:30 ^
    /TR "%ACTION%" ^
    /F

if errorlevel 1 (
    echo.
    echo [ERROR] schtasks /Create failed with code %ERRORLEVEL%.
    echo         Try running this .bat from an elevated command prompt
    echo         ^(right-click -^> Run as administrator^) so it can write to
    echo         the task store; the task itself still runs as the current
    echo         user, NOT as admin/SYSTEM.
    exit /b 22
)

REM --- Verify ---------------------------------------------------------------
echo.
echo [INFO] Task registered. Current state:
schtasks /Query /TN "%TASK_NAME%" /FO LIST /V | findstr /I /C:"TaskName" /C:"Next Run Time" /C:"Status" /C:"Last Run" /C:"Run As User" /C:"Scheduled Task State"

echo.
echo [INFO] Done. The watcher will start automatically at the next logon.
echo        To start it right now without logging off, either:
echo          - run %LAUNCHER% manually, or
echo          - run:  schtasks /Run /TN "%TASK_NAME%"
endlocal & exit /b 0

:bad_user
echo [ERROR] Refusing to install under the service account "%USERNAME%".
echo         This task must run as the INTERACTIVE user whose printers
echo         are redirected through RDP. Log on as that user and re-run.
endlocal & exit /b 23
