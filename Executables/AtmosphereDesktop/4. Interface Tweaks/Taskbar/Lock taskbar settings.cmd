@echo off
echo Run this as administrator
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "TaskbarNoSearchBox" /t REG_DWORD /d 1 /f >nul 2>&1
if errorlevel 1 (
    echo Failed to lock taskbar settings.
    pause >nul
) else (
    echo Taskbar settings have been locked.
    pause >nul
)