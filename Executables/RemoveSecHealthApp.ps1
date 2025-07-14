# Remove Security Health UI (Windows Security app)

# App identifiers to remove
$remove_appx = @("SecHealthUI")
$eol = @()

# Collect provisioned and installed packages
$provisioned = Get-AppxProvisionedPackage -Online
$appxpackage = Get-AppxPackage -AllUsers

# Registry path for AppxAllUserStore
$store = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore'

# System and local user SIDs
$users = @('S-1-5-18')  # SYSTEM SID
if (Test-Path $store) {
    $users += Get-ChildItem $store -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -like 'S-1-5-21-*' } |
        Select-Object -ExpandProperty PSChildName
}

foreach ($choice in $remove_appx) {
    if ([string]::IsNullOrWhiteSpace($choice)) { continue }

    # --- Remove provisioned packages ---
    foreach ($appx in $provisioned | Where-Object { $_.PackageName -like "*$choice*" }) {
        $PackageName = $appx.PackageName
        $DisplayName = $appx.DisplayName
        $PackageFamilyName = ($appxpackage | Where-Object { $_.Name -eq $DisplayName }).PackageFamilyName
        if (-not $PackageFamilyName) { continue }

        # Registry cleanup
        New-Item -Path "$store\Deprovisioned\$PackageFamilyName" -Force -ErrorAction SilentlyContinue | Out-Null
        foreach ($sid in $users) {
            New-Item -Path "$store\EndOfLife\$sid\$PackageName" -Force -ErrorAction SilentlyContinue | Out-Null
        }

        # Make app removable and uninstall
        dism /online /set-nonremovableapppolicy /packagefamily:$PackageFamilyName /nonremovable:0 | Out-Null
        Remove-AppxProvisionedPackage -PackageName $PackageName -Online -AllUsers -ErrorAction SilentlyContinue

        $eol += $PackageName
    }

    # --- Remove installed packages for all users ---
    foreach ($appx in $appxpackage | Where-Object { $_.PackageFullName -like "*$choice*" }) {
        $PackageFullName = $appx.PackageFullName
        $PackageFamilyName = $appx.PackageFamilyName

        New-Item -Path "$store\Deprovisioned\$PackageFamilyName" -Force -ErrorAction SilentlyContinue | Out-Null
        foreach ($sid in $users) {
            New-Item -Path "$store\EndOfLife\$sid\$PackageFullName" -Force -ErrorAction SilentlyContinue | Out-Null
        }

        dism /online /set-nonremovableapppolicy /packagefamily:$PackageFamilyName /nonremovable:0 | Out-Null
        Remove-AppxPackage -Package $PackageFullName -AllUsers -ErrorAction SilentlyContinue

        $eol += $PackageFullName
    }
}

# Apply registry edits
$regFilePath = Join-Path $PSScriptRoot "RemoveDefender.reg"
Start-Process -FilePath "regedit.exe" -ArgumentList "/s `"$regFilePath`"" -Wait -NoNewWindow


# Output summary
if ($eol.Count -gt 0) {
    Write-Host "`n[✓] Successfully removed packages:" -ForegroundColor Green
    $eol | ForEach-Object { Write-Host " - $_" }
} else {
    Write-Host "`n[!] No matching packages were found or removed." -ForegroundColor Yellow
}
