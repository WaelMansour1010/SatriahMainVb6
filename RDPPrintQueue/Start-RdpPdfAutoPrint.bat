@echo off
REM =====================================================================
REM  Start-RdpPdfAutoPrint.bat
REM  Production launcher for the RDP PDF auto-print watcher.
REM  - Does NOT hardcode a user profile Sumatra path.
REM  - Prefers system-wide installs, then lets PowerShell resolve the rest
REM    (LOCALAPPDATA fallback) via Resolve-SumatraPath.
REM =====================================================================
setlocal ENABLEEXTENSIONS

set "QUEUE_ROOT=C:\RDPPrintQueue"
set "SCRIPT=%QUEUE_ROOT%\RdpPdfAutoPrint.ps1"

REM --- Preferred Sumatra locations (system-wide) ---------------------------
set "SUMATRA="
if exist "C:\Program Files\SumatraPDF\SumatraPDF.exe" (
    set "SUMATRA=C:\Program Files\SumatraPDF\SumatraPDF.exe"
) else if exist "C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe" (
    set "SUMATRA=C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe"
)

REM Any remaining fallback (per-user LOCALAPPDATA) is resolved inside the
REM PowerShell script by Resolve-SumatraPath, so we do NOT hardcode it here.

if not exist "%SCRIPT%" (
    echo [ERROR] Script not found: %SCRIPT%
    echo         Copy RdpPdfAutoPrint.ps1 into %QUEUE_ROOT% and try again.
    exit /b 10
)

echo [INFO] QueueRoot  = %QUEUE_ROOT%
echo [INFO] Script     = %SCRIPT%
if defined SUMATRA (
    echo [INFO] Sumatra    = %SUMATRA%
) else (
    echo [INFO] Sumatra    = ^<let PowerShell resolve (LOCALAPPDATA fallback)^>
)

REM --- Launch PowerShell ---------------------------------------------------
REM   -NoProfile         : faster, predictable environment
REM   -ExecutionPolicy   : bypass for this process only
REM   -WindowStyle       : hidden once you are confident; keep Normal during rollout
if defined SUMATRA (
    powershell.exe -NoProfile -ExecutionPolicy Bypass ^
        -File "%SCRIPT%" ^
        -QueueRoot "%QUEUE_ROOT%" ^
        -SumatraPath "%SUMATRA%"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass ^
        -File "%SCRIPT%" ^
        -QueueRoot "%QUEUE_ROOT%"
)

set "RC=%ERRORLEVEL%"
echo [INFO] PowerShell exited with code %RC%
endlocal & exit /b %RC%
