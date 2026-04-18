@echo off
setlocal ENABLEEXTENSIONS

REM =====================================================================
REM  Setup-Printer-Watcher.bat
REM  One-click setup for the RDP PDF auto-print watcher
REM
REM  What it does:
REM    1) Verifies required files exist
REM    2) Runs Install-AutoStart.bat
REM    3) Runs Check-AutoStart.bat
REM    4) Shows a simple final result for the user
REM
REM  Run this ONCE as the same Windows user who will use printing.
REM =====================================================================

set "APPDIR=%~dp0"
if "%APPDIR:~-1%"=="\" set "APPDIR=%APPDIR:~0,-1%"

set "INSTALL_BAT=%APPDIR%\Install-AutoStart.bat"
set "CHECK_BAT=%APPDIR%\Check-AutoStart.bat"
set "START_BAT=%APPDIR%\Start-RdpPdfAutoPrint.bat"
set "PS1=%APPDIR%\RdpPdfAutoPrint.ps1"

echo ================================================================
echo   RDP PDF Auto Print - One-Time Setup
echo ================================================================
echo.
echo User        : %USERDOMAIN%\%USERNAME%
echo Computer    : %COMPUTERNAME%
echo Folder      : %APPDIR%
echo.

REM --- Basic file checks ---------------------------------------------------
if not exist "%INSTALL_BAT%" (
    echo [ERROR] Missing file: Install-AutoStart.bat
    goto :fail
)

if not exist "%CHECK_BAT%" (
    echo [ERROR] Missing file: Check-AutoStart.bat
    goto :fail
)

if not exist "%START_BAT%" (
    echo [ERROR] Missing file: Start-RdpPdfAutoPrint.bat
    goto :fail
)

if not exist "%PS1%" (
    echo [ERROR] Missing file: RdpPdfAutoPrint.ps1
    goto :fail
)

echo [INFO] Required files found.
echo.

REM --- Install auto-start --------------------------------------------------
echo [STEP 1] Installing auto-start task...
call "%INSTALL_BAT%"
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo [ERROR] Install-AutoStart.bat failed with code %RC%.
    echo.
    echo Please make sure:
    echo   1) You are logged in with the same user who will print
    echo   2) You have permission to create Scheduled Tasks
    echo   3) If needed, run this file as Administrator
    goto :fail
)

echo.
echo [INFO] Auto-start installation completed.
echo.

REM --- Check final status --------------------------------------------------
echo [STEP 2] Checking task status...
call "%CHECK_BAT%"
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo ================================================================
    echo   SUCCESS
    echo ================================================================
    echo The printer watcher has been installed successfully.
    echo.
    echo From now on, it should start automatically when you sign in.
    echo.
    echo You do NOT need to run this setup again.
    echo ================================================================
    goto :done
)

if "%RC%"=="40" (
    echo [ERROR] Scheduled Task was not found after installation.
    goto :fail
)

if "%RC%"=="41" (
    echo [ERROR] Scheduled Task exists but is disabled.
    echo Please contact support.
    goto :fail
)

if "%RC%"=="42" (
    echo [ERROR] Could not query Scheduled Task status.
    echo Please contact support.
    goto :fail
)

echo [ERROR] Unexpected Check-AutoStart.bat return code: %RC%
goto :fail

:fail
echo.
echo ================================================================
echo   SETUP DID NOT COMPLETE
echo ================================================================
echo Please contact support and send them a screenshot of this window.
echo ================================================================
pause
endlocal
exit /b 1

:done
echo.
pause
endlocal
exit /b 0