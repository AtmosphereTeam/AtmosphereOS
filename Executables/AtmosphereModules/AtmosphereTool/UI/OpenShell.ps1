$windir = [Environment]::GetFolderPath('Windows')
& "$windir\AtmosphereModules\initPowerShell.ps1"

$timeouts = @("--connect-timeout", "10", "--retry", "5", "--retry-delay", "0", "--retry-all-errors")

# Create temporary directory
function Remove-TempDirectory { Pop-Location; Remove-Item -Path $tempDir -Force -Recurse -EA 0 }
$tempDir = Join-Path -Path $(Get-SystemDrive) -ChildPath $([System.Guid]::NewGuid())
New-Item $tempDir -ItemType Directory -Force | Out-Null
Push-Location $tempDir

# Open-Shell
$programfiles = [System.Environment]::GetFolderPath('ProgramFiles')
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
Write-Output "Downloading Open-Shell..."
& curl.exe -LSs "https://github.com/Open-Shell/Open-Shell-Menu/releases/download/v4.4.196/OpenShellSetup_4_4_196.exe" -o "$tempDir\OpenShellSetup.exe" $timeouts
Write-Output "Installing Open-Shell..."
Start-Process -FilePath "$tempDir\OpenShellSetup.exe" -WindowStyle Hidden -ArgumentList '/qn /quiet ADDLOCAL=StartMenu' -Wait
Write-Output "Open-Shell installed successfully."
Remove-TempDirectory
Write-Output "Configuring Open-Shell"
New-Item -Path "$programfiles\Open-Shell" -ItemType Directory -Force
Start-Process -FilePath "$windir\AtmosphereModules\Scripts\SLNT.bat" -ArgumentList "nu" -WindowStyle Hidden -Wait
Copy-Item -Path "$scriptDir\Skins\Fluent-Metro.skin" -Destination "$programfiles\Open-Shell" -Force
Copy-Item -Path "$scriptDir\Skins\Fluent-Metro.skin7" -Destination "$programfiles\Open-Shell" -Force
Write-Output "Open-Shell Configured successfully."
exit