@echo off
echo Are you sure you want to enable Client CBS?
echo Enabling Client CBS will restore the Start Menu, Taskbar and Search functionality and possibly Windows Update.
echo Enabling will increase background processes.
choice /c YN /n /m "Do you want to proceed? (Y/N)"
if errorlevel 2 (
    echo No changes have been done.
    pause
    exit /b
)
takeown /f "C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy" /r /d y
icacls "C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy" /grant administrators:F /t
cd "C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy"
ren appxmanifest2.disabled appxmanifest.xml
echo Client CBS has been enabled
echo You need to restart your system for the changes to take effect
pause
exit /b