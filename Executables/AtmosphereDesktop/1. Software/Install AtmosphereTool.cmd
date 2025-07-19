@echo off
setlocal EnableDelayedExpansion

set "script=%windir%\AtmosphereModules\Scripts\ScriptWrappers\InstallAtmosphereTool.ps1"
if not exist "%script%" (
	echo Script not found.
	echo "%script%"
	pause
	exit /b 1
)

fltmc >nul 2>&1 || (
    echo.
    echo This script requires Administrator privileges.
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\" %*' -Verb RunAs"
    exit /b
)

choice /c YN /n /m "Do you want to install AtmosphereTool? (Y/N)"
if errorlevel 2 (
    echo Cancelled.
    pause
    exit /b
)

powershell -EP Bypass -NoP -File "%script%" %*