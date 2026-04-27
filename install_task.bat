@echo off
setlocal

set "MSIX_FILE=%~dp0python-manager-26.1.msix"

REM --- Check administrator privilege ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo FAILED: This script requires administrator privileges.
    echo Right-click install_task.bat and choose "Run as administrator".
    pause
    exit /b 1
)

REM --- Find pythonw.exe (excluding WindowsApps stub) ---
call :find_pythonw
if defined PYTHONW goto :have_python

REM --- Python not found: try bundled MSIX installer ---
echo Python not found.
if not exist "%MSIX_FILE%" (
    echo FAILED: Bundled installer not found at:
    echo   %MSIX_FILE%
    echo Install Python manually from https://www.python.org/ then retry.
    pause
    exit /b 1
)

echo Installing Python Manager from bundled MSIX...
echo   %MSIX_FILE%
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Add-AppxPackage -Path '%MSIX_FILE%' -ErrorAction Stop; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if %errorlevel% neq 0 (
    echo Silent install failed. Launching App Installer UI as fallback...
    start "" "%MSIX_FILE%"
    echo.
    echo Complete the install dialog, then re-run install_task.bat.
    pause
    exit /b 1
)

echo OK: Python Manager installed.

REM --- Re-detect pythonw.exe after install (PATH may need refresh) ---
call :find_pythonw
if not defined PYTHONW (
    echo.
    echo Python Manager is installed but pythonw.exe is not yet available.
    echo Open a NEW terminal window and re-run install_task.bat.
    pause
    exit /b 0
)

:have_python
echo Using: %PYTHONW%

REM --- python.exe should be in the same dir (used for import check) ---
set "PYTHON=%PYTHONW:pythonw.exe=python.exe%"
if not exist "%PYTHON%" (
    echo FAILED: python.exe not found at "%PYTHON%".
    pause
    exit /b 1
)

REM --- Verify required stdlib modules can be imported ---
echo Checking required modules...
"%PYTHON%" -c "import base64, ctypes, ctypes.wintypes, email, imaplib, json, logging, os, re, signal, subprocess, threading, winsound, http.server" 2>nul
if %errorlevel% neq 0 (
    echo FAILED: Required stdlib modules are missing. Python install may be broken.
    "%PYTHON%" -c "import base64, ctypes, ctypes.wintypes, email, imaplib, json, logging, os, re, signal, subprocess, threading, winsound, http.server"
    pause
    exit /b 1
)
echo OK: All required modules available.

REM --- Register task: start 7 seconds after logon ---
schtasks /create /tn "SagaOTP_Watcher" /tr "\"%PYTHONW%\" \"%~dp0otp_watcher.pyw\"" /sc onlogon /delay 0000:07 /rl limited /f
if %errorlevel% neq 0 (
    echo FAILED: Could not register scheduled task.
    pause
    exit /b 1
)
echo OK: Task registered.

REM --- Disable AC-power-only and stop-on-battery conditions (avoids "Queued" state on laptops) ---
echo Disabling battery-power restrictions...
powershell -NoProfile -Command "try { $t = Get-ScheduledTask -TaskName 'SagaOTP_Watcher' -ErrorAction Stop; $t.Settings.DisallowStartIfOnBatteries = $false; $t.Settings.StopIfGoingOnBatteries = $false; Set-ScheduledTask -InputObject $t | Out-Null; Write-Host 'OK: Battery restrictions disabled.'; exit 0 } catch { Write-Host ('WARNING: ' + $_.Exception.Message); exit 1 }"
if %errorlevel% neq 0 (
    echo WARNING: Battery restrictions could not be disabled automatically.
    echo If the task is "Queued" while on battery, fix it manually:
    echo   taskschd.msc -^> SagaOTP_Watcher -^> Properties -^> Conditions
    echo   uncheck "Start the task only if the computer is on AC power"
)

echo.
echo OK: OTP Watcher will start 7 seconds after logon.
pause
exit /b 0


REM ============================================================
REM Subroutine: find pythonw.exe (excluding WindowsApps stub)
REM ============================================================
:find_pythonw
set "PYTHONW="
for /f "delims=" %%i in ('where pythonw.exe 2^>nul ^| findstr /v /i "WindowsApps"') do (
    if not defined PYTHONW set "PYTHONW=%%i"
)
exit /b 0
