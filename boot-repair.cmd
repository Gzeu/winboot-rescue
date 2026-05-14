@echo off
:: ============================================================
:: boot-repair.cmd - Minimal launcher for boot-repair.ps1
:: Drop this file alongside boot-repair.ps1 on the USB stick.
:: Run from WinRE/WinPE command prompt as Administrator.
:: ============================================================
setlocal

:: Check for Administrator privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] This script must be run as Administrator.
    echo         Right-click and select "Run as Administrator".
    pause
    exit /b 1
)

:: Determine the directory this CMD lives in
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%boot-repair.ps1"

:: Verify the PowerShell script exists
if not exist "%PS_SCRIPT%" (
    echo [ERROR] boot-repair.ps1 not found in: %SCRIPT_DIR%
    echo         Ensure both files are in the same folder on the USB stick.
    pause
    exit /b 1
)

echo ============================================================
echo   WINDOWS BOOT REPAIR UTILITY v2.0
echo   Launching PowerShell script...
echo ============================================================
echo.

:: Launch PowerShell with ExecutionPolicy bypass
:: -NoProfile avoids loading any profile that might interfere
:: -NonInteractive is NOT set - we need interactive prompts
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%PS_SCRIPT%"

:: Capture exit code
set PS_EXIT=%errorlevel%

if %PS_EXIT% neq 0 (
    echo.
    echo [WARN] PowerShell exited with code %PS_EXIT%
    echo        Review the log files on this USB stick for details.
)

pause
exit /b %PS_EXIT%