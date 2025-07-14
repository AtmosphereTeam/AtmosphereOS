@echo off
setlocal
set "script=%windir%\AtmosphereModules\Scripts\ScriptWrappers\ToggleTF.ps1"

echo Press the number of which operation you wanna do
echo.
echo 1. Install / Enable Translucent Flyouts
echo 2. Disable / Uninstall Translucent Flyouts

choice /c:12 /n

if errorlevel 2 (
    choice /c:yn /n /m "Do you want to uninstall Translucent Flyouts or just disable? (Y/N)"
    if errorlevel 2 (
        powershell -ExecutionPolicy Bypass -NoProfile -File "%script%" -Operation 3 %*
        pause
        exit /b
    )
    powershell -ExecutionPolicy Bypass -NoProfile -File "%script%" -Operation 4 %*
    pause
    exit /b
)

if errorlevel 1 (
    choice /c:yn /n /m "Do you need to install Translucent Flyouts? (Y/N)"
    echo If you chose the Modify UI option during Atmosphere installation you don't need to install
    if errorlevel 2 (
        powershell -ExecutionPolicy Bypass -NoProfile -File "%script%" -Operation 2 %*
        pause
        exit /b
    )
    powershell -ExecutionPolicy Bypass -NoProfile -File "%script%" -Operation 1 %*
    pause
    exit /b
)
