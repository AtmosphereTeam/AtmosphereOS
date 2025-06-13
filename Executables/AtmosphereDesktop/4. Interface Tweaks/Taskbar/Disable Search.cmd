@echo off
echo You need to run this script as Administrator
echo If you haven't, then please run it as administrator again
choice /c YN /n /m "Do you want to disable search? (Y/N)"
if errorlevel 2 (
    echo No changes have been done.
    pause
    exit /b
)
REG ADD HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search /v SearchboxTaskbarMode /t REG_DWORD /d 0 /f
REG ADD HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search /v SearchboxTaskbarModeCache /t REG_DWORD /d 0 /f
sc stop WSearch
sc config WSearch start= disabled
echo Search has been disabled
pause