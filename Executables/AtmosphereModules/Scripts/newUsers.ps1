# Make sure
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$windir = [Environment]::GetFolderPath('Windows')
& "$windir\AtmosphereModules\initPowerShell.ps1"
$AtmosphereDesktop = "$windir\AtmosphereDesktop"
$AtmosphereModules = "$windir\AtmosphereModules"

$title = 'Preparing Atmosphere user settings...'

if (!(Test-Path $AtmosphereDesktop) -or !(Test-Path $AtmosphereModules)) {
    Write-Host "Atmosphere was about to configure user settings, but its files weren't found. :(" -ForegroundColor Red
    Read-Pause
    exit 1
}

function Get-InteractiveUserStartup {
    $explorer = Get-Process explorer -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $explorer) {
        Write-Warning "No explorer.exe process found. Cannot determine interactive user."
        return $null
    }
    
    # Use Get-WmiObject to get process with method GetOwnerSid()
    $wmiProcess = Get-WmiObject Win32_Process -Filter "ProcessId = $($explorer.Id)"
    $ownerSidResult = $wmiProcess.GetOwnerSid()
    if ($ownerSidResult.ReturnValue -ne 0) {
        Write-Warning "Failed to get owner SID of explorer.exe process."
        return $null
    }
    $ownerSid = $ownerSidResult.SID
    
    # Get user profile path from registry
    $profileListKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$ownerSid"
    $profilePath = (Get-ItemProperty -Path $profileListKey -Name ProfileImagePath -ErrorAction SilentlyContinue).ProfileImagePath
    
    if (-not $profilePath) {
        Write-Warning "Cannot find profile path for user SID $ownerSid"
        return $null
    }
    
    return @{
        ProfilePath = $profilePath
        StartupPath = Join-Path $profilePath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        SID = $ownerSid
    }
}
$userInfo = Get-InteractiveUserStartup
if (-not $userInfo) {
    Write-Host "Could not determine interactive user profile. Exiting..." -ForegroundColor Red
    exit 1
}
$profilePath = $userInfo.ProfilePath
$startupdir = $userInfo.StartupPath
$ownerSid = $userInfo.SID

$Host.UI.RawUI.WindowTitle = $title
Write-Host $title -ForegroundColor Yellow
Write-Host $('-' * ($title.length + 3)) -ForegroundColor Yellow
Write-Host "You'll be logged out in 10 to 20 seconds, and once you login again, your new account will be ready for use."

# Disable Windows 11 context menu & 'Gallery' in File Explorer
if ([System.Environment]::OSVersion.Version.Build -ge 22000) {
    reg import "$AtmosphereDesktop\4. Interface Tweaks\Context Menus\Windows 11\Old Context Menu (default).reg" *>$null
    reg import "$AtmosphereDesktop\4. Interface Tweaks\File Explorer Customization\Gallery\Disable Gallery (default).reg" *>$null

    # Set ThemeMRU (recent themes)
    Set-ThemeMRU | Out-Null
}

# Set lockscreen wallpaper
Set-LockscreenImage

# Disable 'Network' in navigation pane
reg import "$AtmosphereDesktop\3. General Configuration\File Sharing\Network Navigation Pane\Disable Network Navigation Pane (default).reg" *>$null

# Disable Automatic Folder Discovery
reg import "$AtmosphereDesktop\4. Interface Tweaks\File Explorer Customization\Automatic Folder Discovery\Disable Automatic Folder Discovery (default).reg" *>$null

# Set visual effects
Start-Process -FilePath "$AtmosphereDesktop\4. Interface Tweaks\Visual Effects (Animations)\Atmosphere Visual Effects (default).cmd" -ArgumentList "/silent" -Wait

# Pin 'Videos' and 'Music' folders to Home/Quick Acesss
$o = new-object -com shell.application
$currentPins = $o.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}').Items() | ForEach-Object { $_.Path }
foreach ($path in @(
    [Environment]::GetFolderPath('MyVideos'),
    [Environment]::GetFolderPath('MyMusic')
)) {
    if ($currentPins -notcontains $path) {
        $o.Namespace($path).Self.InvokeVerb('pintohome')
    }
}

# Disable taskbar search box
Set-ItemProperty -Path "Registry::HKEY_USERS\$ownerSid\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0

# Hide the Task View button
$taskViewPath = "Registry::HKEY_USERS\$ownerSid\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
New-Item -Path $taskViewPath -Force | Out-Null
Set-ItemProperty -Path $taskViewPath -Name "ShowTaskViewButton" -Type DWord -Value 0

# Delete AllUpView\Enabled if it exists (to disable timeline/multitasking view)
$allUpViewPath = "Registry::HKEY_USERS\$ownerSid\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MultiTaskingView\AllUpView"
if (Test-Path $allUpViewPath) {
    Remove-ItemProperty -Path $allUpViewPath -Name "Enabled" -ErrorAction SilentlyContinue
}

# Start TranslucentFlyouts on login
New-Shortcut -Source "$windir\AtmosphereModules\Tools\TranslucentFlyouts\launch_win32.cmd" -Destination "$startupdir\TranslucentFlyouts.lnk"

# Open-Shell
Start-Process -FilePath "$windir\AtmosphereModules\Scripts\SLNT.bat" -ArgumentList "nu"
 
# Set Atmosphere theme as default for current user
$themeKey = "Registry::HKEY_USERS\$ownerSid\Software\Policies\Microsoft\Windows\Personalization"
New-Item -Path $themeKey -Force | Out-Null
Set-ItemProperty -Path $themeKey -Name "ThemeFile" -Value "$windir\Resources\Themes\atmosphere-dark.theme"

# Apply the theme immediately
Start-Process -FilePath "$windir\Resources\Themes\atmosphere-dark.theme" -Wait
Stop-Process -Name "SystemSettings" -Force -ErrorAction SilentlyContinue

# Set custom wallpaper
$wallpaperKey = "Registry::HKEY_USERS\$ownerSid\Control Panel\Desktop"
Set-ItemProperty -Path $wallpaperKey -Name "WallPaper" -Value "$windir\AtmosphereModules\Wallpapers\atmosphere.png"

# Set dark mode
$personalizePath = "Registry::HKEY_USERS\$ownerSid\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
New-Item -Path $personalizePath -Force | Out-Null
Set-ItemProperty -Path $personalizePath -Name "AppsUseLightTheme" -Value 0
Set-ItemProperty -Path $personalizePath -Name "SystemUsesLightTheme" -Value 0

# Atmosphere Desktop
if (-not (Test-Path "$profilePath\Desktop\Atmosphere.lnk")) {
    New-Shortcut -Source "$windir\AtmosphereDesktop" -Destination "$profilePath\Desktop\Atmosphere.lnk" -Icon "$windir\AtmosphereModules\Other\atmosphere-folder.ico,0"
}

# Leave
Remove-Item "$startupdir\AtmosphereUser.lnk"
Start-Sleep 5
logoff