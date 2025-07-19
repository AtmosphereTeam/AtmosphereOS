@echo off
setlocal EnableDelayedExpansion
echo This is the 2nd stage of AtmosphereOS uninstallation
fltmc >nul 2>&1 || (
    echo.
    echo This script requires Administrator privileges.
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\" %*' -Verb RunAs"
    exit /b
)
echo Removing AtmosphereTool...
rmdir /s /q "C:\Program Files\AtmosphereTool"
echo Removing AtmosphereOS Folders...
rmdir /s /q "C:\Windows\AtmosphereDesktop"
rmdir /s /q "C:\Windows\AtmosphereModules"
echo Resetting Themes...
del /f /q "C:\Windows\Resources\Themes\atmosphere-dark.theme"
echo Cleaning up...
rmdir /s /q "C:\cleanup"
@echo on
echo Running SFC.
echo This might take a while...
sfc /scannow
echo Finished.
echo Please restart your device.
echo The script will delete itself after continuing.
pause
start "" cmd /c "del /f /q \"%~f0\""