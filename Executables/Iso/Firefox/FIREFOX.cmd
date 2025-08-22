@echo OFF

setlocal EnableDelayedExpansion

echo. & echo Comparing Firefox entries...
set /a "count2=0"
for /f "usebackq tokens=2 delims=-" %%D in (`reg query "HKLM\SOFTWARE\Clients\StartMenuInternet" /f "Firefox-" ^| findstr /c:"Firefox-"`) do (
	set /a "count2=!count2!+1"
	%arg%set "NewCode=%%D"%par%
)

if "%count1%"=="0" (if "%count2%"=="0" (set "NewCode=NULL"))
endlocal & set "NewCode=%NewCode%"

if not exist "%~dp0\AME-Firefox-Injection" (
    echo.
    echo No supplied AME-Firefox-Injection folder detected.
    exit /b 4
)

:: Generate random string for profile folder
:GenRND
setlocal EnableDelayedExpansion
set "RNDConsist=abcdefghijklmnopqrstuvwxyz0123456789"
set /a "RND=%RANDOM% %% 36"
set "RNDStr=!RNDStr!!RNDConsist:~%RND%,1!"
if "%RNDStr:~7%"=="" (goto GenRND)
endlocal & set "RNDStr=%RNDStr%"

:: Determine profileName
set /a count0=1
:PROFILENAME
if %count0% GTR 50 (
    echo.
    echo Default-release count exceeded 50
    exit /b 0
)
if exist "%AppData%\Mozilla\Firefox\profiles.ini" (
    findstr /c:"Name=default-release" "%AppData%\Mozilla\Firefox\profiles.ini" > NUL 2>&1
    if not errorlevel 1 (
        findstr /c:"Name=default-release-%count0%" "%AppData%\Mozilla\Firefox\profiles.ini"
        if not errorlevel 1 (
            set /a "count0=%count0%+1"
            goto PROFILENAME
        ) else (
            set "profileName=default-release-%count0%"
        )
    ) else (
        set "profileName=default-release"
    )
) else (
    set "profileName=default-release"
)

echo. & echo Injecting profile...

:: Create new profile folder and copy injection files
mkdir "%AppData%\Mozilla\Firefox\Profiles\%RNDStr%.%profileName%"
robocopy "%~dp0\AME-Firefox-Injection" "%AppData%\Mozilla\Firefox\Profiles\%RNDStr%.%profileName%" /E /xf "3647222921wleabcEoxlt-eengsairo.sqlite" > NUL
mkdir "%AppData%\Mozilla\Firefox\Profiles\%RNDStr%.%profileName%\storage\default\moz-extension+++41087662-660a-4251-8c0c-38aa4da5b325^userContextId=4294967295\idb"
copy /y "%~dp0\AME-Firefox-Injection\3647222921wleabcEoxlt-eengsairo.sqlite" "%AppData%\Mozilla\Firefox\Profiles\%RNDStr%.%profileName%\storage\default\moz-extension+++41087662-660a-4251-8c0c-38aa4da5b325^userContextId=4294967295\idb"
if exist "%AppData%\Mozilla\Extensions\" (
    copy /y "%~dp0\AME-Firefox-Injection\extensions\" "%AppData%\Mozilla\Extensions\"
) else (
    mkdir "%AppData%\Mozilla\Extensions"
    copy /y "%~dp0\AME-Firefox-Injection\extensions\" "%AppData%\Mozilla\Extensions\"
)

:: Update installs.ini file
echo [%NewCode%]>> "%AppData%\Mozilla\Firefox\installs.ini"
echo Default=Profiles/%RNDStr%.%profileName%>> "%AppData%\Mozilla\Firefox\installs.ini"
echo Locked=^1>> "%AppData%\Mozilla\Firefox\installs.ini"

:: Update profiles.ini with new profile as default
echo [Install%NewCode%]>> "%AppData%\Mozilla\Firefox\profiles.ini"
echo Default=Profiles/%RNDStr%.%profileName%>> "%AppData%\Mozilla\Firefox\profiles.ini"
echo Locked=^1>> "%AppData%\Mozilla\Firefox\profiles.ini"
echo.>> "%AppData%\Mozilla\Firefox\profiles.ini"
echo [Profile0]>> "%AppData%\Mozilla\Firefox\profiles.ini"
echo Name=%profileName%>> "%AppData%\Mozilla\Firefox\profiles.ini"
echo IsRelative=^1>> "%AppData%\Mozilla\Firefox\profiles.ini"
echo Path=Profiles/%RNDStr%.%profileName%>> "%AppData%\Mozilla\Firefox\profiles.ini"

:: Add prefs to other Firefox profiles in all users
for /f "usebackq delims=" %%B in (`dir /B /A:d "%AppData%\Mozilla\Firefox\Profiles" ^| findstr /v /x /c:"%RNDStr%.%profileName%"`) do (
    if exist "%AppData%\Mozilla\Firefox\Profiles\%%B\prefs.js" (
        findstr /V /C:"app.shield.optoutstudies.enabled" /C:"browser.aboutwelcome.enabled" /C:"browser.disableResetPrompt" /C:"browser.newtabpage.activity-stream.asrouter.userprefs.cfr.addons" /C:"browser.newtabpage.activity-stream.asrouter.userprefs.cfr.features" /C:"browser.newtabpage.activity-stream.feeds.section.topstories" /C:"browser.newtabpage.activity-stream.feeds.topsites" /C:"browser.newtabpage.activity-stream.section.highlights.includePocket" /C:"browser.newtabpage.activity-stream.section.highlights.includeVisited" /C:"browser.newtabpage.activity-stream.showSponsored" /C:"browser.newtabpage.activity-stream.showSponsoredTopSites" /C:"browser.urlbar.placeholderNam" /C:"browser.urlbar.suggest.quicksuggest.nonsponsored" /C:"browser.urlbar.suggest.quicksuggest.sponsored" /C:"browser.urlbar.suggest.topsites" /C:"datareporting.healthreport.uploadEnabled" /C:"dom.security.https_only_mode" /C:"dom.security.https_only_mode_ever_enabled" "%AppData%\Mozilla\Firefox\Profiles\%%B\prefs.js" > "%TEMP%\prefs.js.tmp"
        findstr /V /C:"browser.toolbars.bookmarks.visibility" /C:"extensions.webextensions.uuids" /C:"extensions.ui.extension.hidden" /C:"extensions.ui.lastCategory" "%~dp0\AME-Firefox-Injection\prefs.js" >> "%TEMP%\prefs.js.tmp"
        move /y "%TEMP%\prefs.js.tmp" "%AppData%\Mozilla\Firefox\Profiles\%%B\prefs.js"
        if exist "%AppData%\Mozilla\Firefox\Profiles\%%B\search.json.mozlz4" del /Q /F "%AppData%\Mozilla\Firefox\Profiles\%%B\search.json.mozlz4"
        robocopy "%~dp0\AME-Firefox-Injection" "%AppData%\Mozilla\Firefox\Profiles\%%B" search.json.mozlz4 /E > NUL
    )
)

exit /b 0