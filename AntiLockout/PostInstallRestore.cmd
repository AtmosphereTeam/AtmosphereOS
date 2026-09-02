@echo off
rem ============================================================
rem  PostInstallRestore.cmd - restore original sethc.exe
rem  Run as Administrator after installing.
rem ============================================================

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Run this as Administrator. Right-click ^> "Run as administrator".
    pause
    exit /b 1
)

set "SYS32=%SystemRoot%\System32"
set "SETHC=%SYS32%\sethc.exe"
set "BACKUP=%SYS32%\sethc.original.exe"

if not exist "%BACKUP%" (
    echo [INFO] No backup found - original sethc.exe not swapped. Nothing to do.
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
