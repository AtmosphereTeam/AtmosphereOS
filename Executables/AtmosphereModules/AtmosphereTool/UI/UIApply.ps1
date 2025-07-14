param (
	[string]$StartupDir
)
$windir = [Environment]::GetFolderPath('Windows')
& "$windir\AtmosphereModules\initPowerShell.ps1"

# Create temporary directory
function Remove-TempDirectory { Pop-Location; Remove-Item -Path $tempDir -Force -Recurse -EA 0 }
$tempDir = Join-Path -Path $(Get-SystemDrive) -ChildPath $([System.Guid]::NewGuid())
New-Item $tempDir -ItemType Directory -Force | Out-Null
Push-Location $tempDir

# Modify Windows Ui
# Nilesoft Shell https://github.com/moudey/Shell
# broken by microsft (i think)
#    Write-Output "Downloading Nilesoft Shell..."
#    $ShellArch = ('x64', 'arm64')[$arm]
#    $nilesoftInstallerPath = Join-Path $tempDir "NilesoftShell.msi"
#
#    & curl.exe -LSs "https://nilesoft.org/download/shell/1.9.18/setup-$ShellArch.msi" -o $nilesoftInstallerPath $timeouts
#    if (!$?) {
#        Write-Error "Failed to download Nilesoft Shell. Error: $($LASTEXITCODE). Exiting."
#        Remove-TempDirectory
#        exit 1
#    }
#
#    Write-Output "Installing Nilesoft Shell..."
#    try {
#        # Corrected: Run MSI using msiexec.exe and use global $msiArgs
#        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$nilesoftInstallerPath`" $msiArgs" -WindowStyle Hidden -Wait -PassThru
#        if ($process.ExitCode -ne 0) {
#            Write-Error "Nilesoft Shell installation failed with exit code $($process.ExitCode). Exiting."
#            Remove-TempDirectory
#            exit 1
#        }
#        Write-Output "Nilesoft Shell installed successfully."
#    } catch {
#        Write-Error "An error occurred during Nilesoft Shell installation: $_. Exiting."
#        Remove-TempDirectory
#        exit 1
#    }

# AccentColorizer11 + AccentColorizer https://github.com/krlvm/AccentColorizer + https://github.com/krlvm/AccentColorizer-E11
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
# Check if Win11 to install AccentColorizer-E11 
$build = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
if ($build -ge 22000) {
	$taskXmlFilePath = Join-Path $scriptDir "AccentColorizer11.xml"
}
else {
	$taskXmlFilePath = Join-Path $scriptDir "AccentColorizer.xml"
}

if (-not (Test-Path $taskXmlFilePath)) {
	Write-Error "AccentColorizer xml not found at '$taskXmlFilePath'. Skipping scheduled task registration."
	Remove-TempDirectory 
	Write-Output "Continuing now with ExplorerBlurMica installation"
	return
}
$taskXml = Get-Content -Path $taskXmlFilePath -Raw
$taskName = "AccentColorizer"
# Unregister existing task if it exists (for clean re-registration)
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
	Write-Output "Existing scheduled task '$taskName' found. Unregistering for update."
	Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
}
try {
	Register-ScheduledTask -TaskName $taskName -Xml $taskXml -Force -ErrorAction Stop
	Write-Output "Scheduled task '$taskName' registered successfully from XML."
}
catch {
	Write-Error "Failed to register scheduled task '$taskName': $($_.Exception.Message). Exiting."
	Remove-TempDirectory
	Write-Output "Continuing now with ExplorerBlurMica installation"
	return
}
Remove-TempDirectory

# ExplorerBlurMica https://github.com/Maplespe/ExplorerBlurMica
$regsvr32Path = Join-Path $env:windir "System32\regsvr32.exe"
$dllPath = "$env:windir\AtmosphereDesktop\4. Interface Tweaks\File Explorer Customization\Mica Explorer\ExplorerBlurMica.dll"
$timeoutSeconds = 2 
Write-Host "Attempting to register DLL: $dllPath with a $timeoutSeconds-second timeout..."
try {
	$process = Start-Process -FilePath $regsvr32Path -ArgumentList "`"$dllPath`" /s" -PassThru -ErrorAction Stop
	$didExit = $process.WaitForExit($timeoutSeconds * 1000)
	if ($didExit) {
		if ($process.ExitCode -eq 0) {
			Write-Host "DLL registered successfully."
		}
		else {
			Write-Error "regsvr32 command exited with non-zero code $($process.ExitCode). This may indicate an issue with the DLL or registration."
			return
		}
	}
 else {
		Write-Error "regsvr32 command timed out after $timeoutSeconds seconds. Attempting to terminate process."
		$process | Stop-Process -Force -ErrorAction SilentlyContinue
		return
	}
}
catch {
	Write-Error "Failed to start or register DLL: $($_.Exception.Message)"
	return 
}
# TranslucentFlyouts
Write-Output "Installing TranslucentFlyouts..."
New-Shortcut -Source "$windir\AtmosphereModules\Tools\TranslucentFlyouts\launch_win32.cmd" -Destination "$StartupDir\TranslucentFlyouts.lnk"
$p = Start-Process -FilePath "C:\Windows\AtmosphereModules\Tools\TranslucentFlyouts\launch_win32.cmd" -PassThru
$p.WaitForExit()
Write-Output "TranslucentFlyouts installed..."
exit
