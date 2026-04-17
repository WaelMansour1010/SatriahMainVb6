@echo off
REM =====================================================================
REM  Test-RdpPdfAutoPrint.bat
REM  Prints a single known PDF and exits. Used to verify that:
REM    - the script is running under the correct user context
REM    - Sumatra can be resolved
REM    - a default (or explicit) printer is visible
REM    - the PDF itself is actually printable
REM  Logs go to C:\RDPPrintQueue\logs\auto-print-YYYY-MM-DD.log
REM =====================================================================
setlocal ENABLEEXTENSIONS

set "QUEUE_ROOT=C:\RDPPrintQueue"
set "SCRIPT=%QUEUE_ROOT%\RdpPdfAutoPrint.ps1"
set "TEST_PDF=C:\RDPPrintQueue\failed\SaleReport545.pdf"

REM --- Preferred Sumatra locations (system-wide) ---------------------------
set "SUMATRA="
if exist "C:\Program Files\SumatraPDF\SumatraPDF.exe" (
    set "SUMATRA=C:\Program Files\SumatraPDF\SumatraPDF.exe"
) else if exist "C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe" (
    set "SUMATRA=C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe"
)

if not exist "%SCRIPT%" (
    echo [ERROR] Script not found: %SCRIPT%
    exit /b 10
)
if not exist "%TEST_PDF%" (
    echo [ERROR] Test PDF not found: %TEST_PDF%
    echo         Put a sample PDF at that path or edit TEST_PDF above.
    exit /b 11
)

echo [INFO] TEST MODE
echo [INFO] QueueRoot  = %QUEUE_ROOT%
echo [INFO] Script     = %SCRIPT%
echo [INFO] TestPDF    = %TEST_PDF%
if defined SUMATRA (
    echo [INFO] Sumatra    = %SUMATRA%
) else (
    echo [INFO] Sumatra    = ^<let PowerShell resolve (LOCALAPPDATA fallback)^>
)

REM Keep window open so the tech running the test can read the diagnostics.
if defined SUMATRA (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit ^
        -File "%SCRIPT%" ^
        -QueueRoot "%QUEUE_ROOT%" ^
        -SumatraPath "%SUMATRA%" ^
        -TestFile "%TEST_PDF%"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit ^
        -File "%SCRIPT%" ^
        -QueueRoot "%QUEUE_ROOT%" ^
        -TestFile "%TEST_PDF%"
)

set "RC=%ERRORLEVEL%"
echo [INFO] PowerShell exited with code %RC%
endlocal & exit /b %RC%
