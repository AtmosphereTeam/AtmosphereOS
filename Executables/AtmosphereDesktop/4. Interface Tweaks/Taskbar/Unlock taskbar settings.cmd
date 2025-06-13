@echo off
echo Run this as administrator
echo WARNING!
echo Atmosphere disables the search bar.
echo Enabling search bar on taskbar just shows the icon and does nothing.
echo If you want to use search bar then run the Enable Search.cmd file as administrator 
choice /c YN /n /m "Do you want to unlock taskbar settings? (Y/N)"

if errorlevel 2 (
    echo No changes have been made.
    pause >nul
) else (
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "TaskbarNoSearchBox" /t REG_DWORD /d 0 /f >nul 2>&1
if errorlevel 1 (
    echo Failed to unlock taskbar settings.
    pause >nul
) else (
    echo Taskbar settings have been unlocked.
    pause >nul
)
)