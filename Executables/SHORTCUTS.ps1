.\AtmosphereModules\initPowerShell.ps1
$windir = [Environment]::GetFolderPath('Windows')

Write-Title "Creating Desktop & Start Menu shortcuts..."

# 1. Make shortcut for primary user. Should even if the 2nd part fails.
New-Shortcut -Source "$windir\AtmosphereDesktop" -Destination "$([Environment]::GetFolderPath('Desktop'))\Atmosphere.lnk" -Icon "$windir\AtmosphereModules\Other\atmosphere-folder.ico,0"

# 2. Iterate through user profiles in C:\Users to copy to their Desktops
#    (Requires Administrator privileges)
$defaultShortcut = "$(Get-UserPath)\Atmosphere.lnk"
$usersRoot = "C:\Users"
$systemUserFolders = @("Public", "Default", "defaultuser0", "All Users", "System") # Folders to exclude

if (Test-Path $usersRoot -PathType Container) {
    Get-ChildItem -Path $usersRoot -Directory | ForEach-Object {
        $username = $_.Name

        # Skip system-related user folders
        if ($systemUserFolders -notcontains $username) {
            $userDesktopPath = Join-Path -Path $_.FullName -ChildPath "Desktop"

            # Check if the Desktop folder actually exists for this user profile
            if (Test-Path $userDesktopPath -PathType Container) {
                Write-Output "Copying Desktop shortcut for '$username' to '$userDesktopPath'..."
                try {
                    Copy-Item $defaultShortcut -Destination $userDesktopPath -Force -ErrorAction Stop
                    Write-Output "Successfully copied shortcut for '$username'." -ForegroundColor Green
                }
                catch {
                    Write-Warning "Failed to copy shortcut to '$userDesktopPath' for user '$username'. Error: $($_.Exception.Message)"
                }
            } else {
                Write-Warning "Desktop path '$userDesktopPath' not found for user '$username', skipping shortcut copy."
            }
        }
    }
} else {
    Write-Error "Could not find C:\Users directory. Cannot copy shortcuts to user desktops."
}


# 3. Create shortcut for the Common (Public) Start Menu
# This shortcut will be available to all users.
Copy-Item $defaultShortcut -Destination "$([Environment]::GetFolderPath('CommonStartMenu'))\Programs" -Force

Write-Title "Creating services restore shortcut..."
$desktop = "$windir\AtmosphereDesktop"
New-Shortcut -Source "$desktop\8. Troubleshooting\Set services to defaults.cmd" -Destination "$desktop\6. Advanced Configuration\Services\Set services to defaults.lnk"