@echo off
setlocal EnableDelayedExpansion

rem ============================================================
rem  DeployAntiLockout.cmd - swap sethc.exe for AntiLockout.exe
rem  MUST be run as Administrator.
rem ============================================================

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Run this as Administrator. Right-click ^> "Run as administrator".
    pause
    exit /b 1
)

set "SYS32=%SystemRoot%\System32"
set "SRC=%~dp0AntiLockout.exe"
set "SETHC=%SYS32%\sethc.exe"
set "BACKUP=%SYS32%\sethc.original.exe"

if not exist "%SRC%" (
    echo [ERROR] AntiLockout.exe not found next to this script.
    pause
    exit /b 1
)

rem --- fail if a swap is already in place (unless restore needed) ---
if /i "%1"=="-restore" goto :restore

if exist "%BACKUP%" (
    echo [ERROR] Backup %BACKUP% already exists - looks like a swap is already active.
    echo         Run this with -restore first, or delete the backup manually.
    pause
    exit /b 1
)

echo [1/4] Taking ownership of sethc.exe...
takeown /f "%SETHC%" >nul 2>&1
icacls "%SETHC%" /grant "%USERNAME%:F" >nul 2>&1

echo [2/4] Backing up original to sethc.original.exe...
move /y "%SETHC%" "%BACKUP%" >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Could not move original sethc.exe.
    pause
    exit /b 1
)

echo [3/4] Copying AntiLockout.exe into place as sethc.exe...
copy /y "%SRC%" "%SETHC%" >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Could not copy AntiLockout.exe.
    echo         Restoring original...
    move /y "%BACKUP%" "%SETHC%" >nul 2>&1
    pause
    exit /b 1
)

echo [4/4] Swap complete.
echo.
echo Press Shift 5 times at the login screen to launch AntiLockout.
echo.
pause
exit /b 0

:restore
if not exist "%BACKUP%" (
    echo [INFO] No backup found - nothing to restore.
    pause
    exit /b 0
)

echo [1/2] Taking ownership of sethc.exe...
takeown /f "%SETHC%" >nul 2>&1
icacls "%SETHC%" /grant "%USERNAME%:F" >nul 2>&1

echo [2/2] Restoring original sethc.exe...
move /y "%SETHC%" "%SYS32%\sethc.antilockout.exe" >nul 2>&1
move /y "%BACKUP%" "%SETHC%" >nul 2>&1
echo [OK] Original sethc.exe restored.
pause
exit /b 0
