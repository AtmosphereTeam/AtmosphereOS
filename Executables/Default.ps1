$windir = [Environment]::GetFolderPath('Windows')
& "$windir\AtmosphereModules\initPowerShell.ps1"

$usersFolder = "$env:SystemDrive\Users"
$Default = Join-Path $usersFolder "Default"
$defaultuser0 = Join-Path $usersFolder "defaultuser0"
$startupDir = "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
$scriptPath = "$windir\AtmosphereModules\Scripts\newUsers.ps1"

# --------------- Configure New Users --------------- #

foreach ($userTemplate in @($Default, $defaultuser0)) {
    if (Test-Path $userTemplate) {
        $startup = Join-Path $userTemplate $startupDir
        New-Item $startup -ItemType Directory -Force | Out-Null
        New-Shortcut `
            -Source "cmd.exe" `
            -Destination "$startup\AtmosphereUser.lnk" `
            -Arguments "/c powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        New-Shortcut -Source "$windir\AtmosphereDesktop" -Destination "$userTemplate\Desktop" -Icon "$windir\AtmosphereModules\Other\atmosphere-folder.ico,0"
    }
}