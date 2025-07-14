# Initialize
param (
    [int]$Operation
)

$windir = [Environment]::GetFolderPath('Windows')
& "$windir\AtmosphereModules\initPowerShell.ps1"
$TranslucentFlyouts = Join-Path -Path $windir -ChildPath "AtmosphereModules\Tools\TranslucentFlyouts"

# Install / Enable
if ($Operation -eq 1 -or $Operation -eq 2) {
    New-Shortcut -Source "$windir\AtmosphereModules\Tools\TranslucentFlyouts\launch_win32.cmd" -Destination "$([Environment]::GetFolderPath('Startup'))\TranslucentFlyouts.lnk"
    if ($Operation = = 1) { Start-Process "Rundll32.exe" -ArgumentList "$TranslucentFlyouts\TFMain64.dll,Main /install" -Wait }
    Start-Process "Rundll32.exe" -ArgumentList "$TranslucentFlyouts\TFMain64.dll,Main /start" -Wait
}

# Uninstall / Remove
if ($Operation -eq 3 -or $Operation -eq 4) {
    Remove-Item "$([Environment]::GetFolderPath('Startup'))\TranslucentFlyouts.lnk"
    Start-Process "Rundll32.exe" -ArgumentList "$TranslucentFlyouts\TFMain64.dll,Main /stop" -Wait
    if ($Operation = = 4) { Start-Process "Rundll32.exe" -ArgumentList "$TranslucentFlyouts\TFMain64.dll,Main /uninstall" -Wait }
}