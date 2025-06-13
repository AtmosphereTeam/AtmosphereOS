@echo off
echo Are you sure you want to disable Client CBS?
echo Disabling Client CBS will break the Start Menu, Taskbar and Search functionality possibly even Windows Update.
echo You shouldn't disable Client CBS unless you know what you are doing.
echo If you want to disable Client CBS make sure that you have Open-Shell or a start menu replacement installed.
choice /c YN /n /m "Do you want to proceed? (Y/N)"
if errorlevel 2 (
    echo No changes have been done.
    pause
    exit /b
)
takeown /f "C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy" /r /d y
icacls "C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy" /grant administrators:F /t
cd "C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy"
ren appxmanifest.xml appxmanifest2.disabled
echo Client CBS has been disabled.
echo You can re-enable Client CBS by running the Enable.cmd script in the same directory.
echo You need to restart your computer for the changes to take effect.
pause
exit /b