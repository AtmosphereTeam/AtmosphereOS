param (
	[switch]$ModifyUi
)

Push-Location -Path $PSScriptRoot
& "..\AtmosphereModules\initPowerShell.ps1"
$msiArgs = "/qn /quiet /norestart ALLUSERS=1 REBOOT=ReallySuppress"
Copy-Item -Path ".\Installers" -Destination "C:\Iso\Installers" -Force -Recurse
$installers = Get-ChildItem -Path ".\Installers" -Depth 2

foreach ($installer in $installers.Name) { Write-Output $installer }
if (-not $installers) { Write-Error "No installers found." }

# AtmosphereTool
if ($installers.Name -like "AtmosphereTool.zip") {
	Write-Output "Installing AtmosphereTool..."
	$AtmosphereTool = ".\Installers\AtmosphereTool.zip"
	Expand-Archive -Path $AtmosphereTool -DestinationPath "C:\Program Files\AtmosphereTool" -Force
	Write-Output "Finished Installing AtmosphereTool."
}

# Files
if ($installers.Name -like "Files.msixbundle") {
	Write-Output "Installing Files..."
	$Files = ".\Installers\Files.msixbundle"
	Copy-Item $Files -Destination "C:\Iso\Installers" -Force
	Write-Output "Finished Installing Files."
}

# Notepads
if ($installers.Name -like "Notepads.msixbundle") {
	Write-Output "Installing Notepads..."
	$Notepads = ".\Installers\Notepads.msixbundle"
	Copy-Item $Notepads -Destination "C:\Iso\Installers" -Force
	Write-Output "Finished Installing Notepads."
}

# AppFetch
if ($installers.Name -like "AppFetch.exe") {
	Write-Output "Installing AppFetch..."
	$AppFetch = ".\Installers\AppFetch.exe"
	Move-Item -Path "$AppFetch" -Destination "C:\ProgramData\AppFetch.exe" -Force
	Write-Output "Finished Installing AppFetch."
}

# UniGetUI
if ($installers.Name -like "UniGetUI.exe") {
	Write-Output "Installing UniGetUI..."
	$UniGetUI = ".\Installers\UniGetUI.exe"
	Start-Process -FilePath $UniGetUI -ArgumentList  "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /NoAutoStart"
	Write-Output "Finished Installing UniGetUI."
}

# FluentTerminal
if ($installers.Name -like "FluentTerminal.zip") {
	Write-Output "Installing FlunetTerminal..."
	$FluentTerminal = ".\Installers\FluentTerminal.zip"
	Copy-Item $FluentTerminal -Destination "C:\Iso\Installers" -Force
	# Expand-Archive -Path $FluentTerminal -DestinationPath ".\Installers\FluentTerminal" -Force
	# $FluentTerminalBundle = Get-ChildItem -Path ".\Installers\FluentTerminal" | Where-Object { $_.FullName -like "*\FluentTerminal*.msixbundle"}
	# Add-AppPackage $FluentTerminalBundle
	Write-Output "Finished Installing FluentTerminal."
}

# Brave
if ($installers.Name -like "BraveSetup.exe") {
	Write-Output "Installing Brave..."
    $brave = ".\Installers\BraveSetup.exe"
	Start-Process -FilePath $brave -WindowStyle Hidden -ArgumentList '/silent /install'
	do {
		$processesFound = Get-Process | Where-Object { "BraveSetup" -contains $_.Name } | Select-Object -ExpandProperty Name
		if ($processesFound) {
			Write-Output "Still running BraveSetup."
			Start-Sleep -Seconds 2
		}
	} until (!$processesFound)
	Stop-Process -Name "brave" -Force -EA 0
	Write-Output "Brave installed successfully."
}

# Firefox
if ($installers.Name -like "FirefoxSetup.exe") {
	$Firefox = ".\Installers\FirefoxSetup.exe"
	Write-Output "Installing Firefox..."
	Start-Process -FilePath $Firefox -WindowStyle Hidden -ArgumentList '/S /ALLUSERS=1' -Wait
	Copy-Item -Path ".\Firefox" -Destination "C:\Iso\" -Force -Recurse
	Write-Output "Firefox installed successfully."
}

# Opera GX
if ($installers.Name -like "OperaGXSetup.exe") {
    $OperaGX = ".\Installers\OperaGXSetup.exe"
	Write-Output "Installing Opera GX..."
	Start-Process -FilePath $OperaGX -WindowStyle Hidden -ArgumentList ' /install /silent /norestart /launchopera=0 /setdefaultbrowser=1 /allusers=1' -Wait
	Write-Output "Opera GX installed successfully."
}

# Open-Shell
if ($installers.Name -like "OpenShellSetup.exe") {
	$windir = [Environment]::GetFolderPath('Windows')
    $OpenShell = ".\Installers\OpenShellSetup.exe"
	Write-Output "Installing Open-Shell..."
	Start-Process -FilePath $OpenShell -WindowStyle Hidden -ArgumentList '/qn /quiet ADDLOCAL=StartMenu' -Wait
	Copy-Item -Path "C:\Windows\AtmosphereModules\AtmosphereTool\UI\Fluent-Metro.skin" -Destination "C:\Program Files\Open-Shell\Skins" -Force
	Copy-Item -Path "C:\Windows\AtmosphereModules\AtmosphereTool\UI\Fluent-Metro.skin7" -Destination "C:\Program Files\Open-Shell\Skins" -Force
	Write-Output "Open-Shell installed."
}

if ($ModifyUi) {
	# AccentColorizer11 + AccentColorizer https://github.com/krlvm/AccentColorizer + https://github.com/krlvm/AccentColorizer-E11
	$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
	$build = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
	if ($build -ge 22000) {
		$taskXmlFilePath = Join-Path $scriptDir "..\AccentColorizer11.xml"
	}
    else {
		$taskXmlFilePath = Join-Path $scriptDir "..\AccentColorizer.xml"
	}
	if (-not (Test-Path $taskXmlFilePath)) {
		Write-Error "AccentColorizer xml not found at '$taskXmlFilePath'. Skipping scheduled task registration."
		Write-Output "Continuing now with ExplorerBlurMica installation"
		return
	}
	$taskXml = Get-Content -Path $taskXmlFilePath -Raw
	$taskName = "AccentColorizer"
	if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
		Write-Output "Existing scheduled task '$taskName' found. Unregistering for update."
		Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
	}
	try {
		Register-ScheduledTask -TaskName $taskName -Xml $taskXml -Force -ErrorAction Stop
		Write-Output "Scheduled task '$taskName' registered successfully from XML."
	}
    catch {
		Write-Error "Failed to register scheduled task '$taskName': $($_.Exception.Message). ing."
		Write-Output "Continuing now with ExplorerBlurMica installation"
		return
	}
	# ExplorerBlurMica https://github.com/Maplespe/ExplorerBlurMica
	$regsvr32Path = Join-Path $env:windir "System32\regsvr32.exe"
	$dllPath = "$env:windir\AtmosphereDesktop\4. Interface Tweaks\File Explorer Customization\Mica Explorer\ExplorerBlurMica.dll"
	$timeoutSeconds = 2 
	Write-Host "Attempting to register DLL: $dllPath with a $timeoutSeconds-second timeout..."
	try {
		$process = Start-Process -FilePath $regsvr32Path -ArgumentList "`"$dllPath`" /s" -PassThru -ErrorAction Stop
		$did = $process.WaitForExit($timeoutSeconds * 1000)
		if ($did) {
			if ($process.Code -eq 0) {
				Write-Host "DLL registered successfully."
			}
			else {
				Write-Error "regsvr32 command ed with non-zero code $($process.Code). This may indicate an issue with the DLL or registration."
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
	# ---------- TranslucentFlyouts
	Write-Output "Installing TranslucentFlyouts..."
	New-Shortcut -Source "C:\Windows\AtmosphereModules\Tools\TranslucentFlyouts\launch_win32.cmd" -Destination "$([Environment]::GetFolderPath('Startup'))\TranslucentFlyouts.lnk"
	# Copy to other users
	$startupShortcut = "$([Environment]::GetFolderPath('Startup'))\TranslucentFlyouts.lnk"
	$usersRoot = "C:\Users"
	$startupPath = "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
	if (Test-Path $usersRoot -PathType Container) {
		Get-ChildItem -Path $usersRoot -Directory | ForEach-Object {
			$username = Join-Path -Path $usersRoot -ChildPath $_.Name
			# Check if the Startup folder actually exists for this user profile
			$startupdir = Join-Path -Path $username -ChildPath $startupPath
			if (Test-Path $startupdir -PathType Container) {
				Write-Output "Copying Startup shortcut for '$username' to '$startupdir'..."
				try {
					Copy-Item $startupShortcut -Destination $startupdir -Force -ErrorAction Stop
					Write-Output "Successfully copied shortcut for '$username'." -ForegroundColor Green
				}
				catch {
					Write-Warning "Failed to copy shortcut to '$startupdir' for user '$username'. Error: $($_.Exception.Message)"
				}
			}
			else {
				Write-Warning "Desktop path '$startupdir' not found for user '$username', skipping shortcut copy."
			}			
		}
	}
	# ---------- Install
	# Start-Process -FilePath "rundll32.exe" -ArgumentList "C:\Windows\AtmosphereModules\Tools\TranslucentFlyouts\TFMain64.dll /start" -WindowStyle Hidden
	# Start-Sleep 180
	# Stop-Process -Name "rundll32.exe"
	Write-Output "TranslucentFlyouts Installed."
}
		
#####################
##    Utilities    ##
#####################

# Visual C++ Runtimes (referred to as vcredists for short)
# https://learn.microsoft.com/en-US/cpp/windows/latest-supported-vc-redist
$legacyArgs = '/q /norestart'
$modernArgs = "/install /quiet /norestart"
$vcredistDir = ".\Installers\vcredists"
$vcredists = [ordered] @{
	# 2005 - version 8.0.50727.6195 (MSI 8.0.61000/8.0.61001) SP1
	"vcredist_1_x64.exe" = @("2005-x64", "/c /q /t:")
	"vcredist_1_x86.exe" = @("2005-x86", "/c /q /t:")
	# 2008 - version 9.0.30729.6161 (EXE 9.0.30729.5677) SP1
	"vcredist_2_x64.exe" = @("2008-x64", "/q /extract:")
	"vcredist_2_x86.exe" = @("2008-x86", "/q /extract:")
	# 2010 - version 10.0.40219.325 SP1
	"vcredist_3_x64.exe" = @("2010-x64", $legacyArgs)
	"vcredist_3_x86.exe" = @("2010-x86", $legacyArgs)
	# 2012 - version 11.0.61030.0
	"vcredist_4_x64.exe" = @("2012-x64", $modernArgs)
	"vcredist_4_x86.exe" = @("2012-x86", $modernArgs)
	# 2013 - version 12.0.40664.0
	"vcredist_5_x64.exe" = @("2013-x64", $modernArgs)
	"vcredist_5_x86.exe" = @("2013-x86", $modernArgs)
	# 2015-2022 (2015+) - latest version
	"vcredist_6_x64.exe" = @("2015+-x64", $modernArgs)
	"vcredist_6_x86.exe" = @("2015+-x86", $modernArgs)
}
foreach ($a in $vcredists.GetEnumerator()) {
	$vcName = $a.Value[0]
	$vcArgs = $a.Value[1]
    $vcExe = $a.Name
	$vcExePath = "$vcredistDir\$vcExe"
	Write-Host "`nProcessing Visual C++ Runtime: $vcName"
	Write-Host "Executable path: $vcExePath"
	if ($vcName -eq "2010-x64") {
		Start-Process -FilePath $vcExePath -ArgumentList $vcArgs -WindowStyle Hidden
		continue
	}
	if ($vcArgs -match ":") {
		$msiDir = "$vcredistDir\vcredist-$vcName"
		Start-Process -FilePath $vcExePath -ArgumentList "$vcArgs`"$msiDir`"" -Wait -WindowStyle Hidden
		$msiPaths = (Get-ChildItem -Path $msiDir -Filter *.msi -EA 0).FullName
		if (!$msiPaths) {
			Write-Output "Failed to extract MSI for $vcName, not installing."
		}
		else {
			$msiPaths | ForEach-Object {
				Start-Process -FilePath "msiexec.exe" -ArgumentList "/log `"$msiDir\logfile.log`" /i `"$_`" $msiArgs" -WindowStyle Hidden
			}
		}
	}
    else {
		Start-Process -FilePath $vcExePath -ArgumentList $vcArgs -Wait -WindowStyle Hidden
	}
}

# 7zip
if ($installers.Name -like "7zip.exe") {
	Write-Output "Installing 7-Zip..."
    $7zip = ".\Installers\7zip.exe"
	Start-Process -FilePath $7zip -WindowStyle Hidden -ArgumentList '/S' -Wait
	Write-Output "Finished Installing 7-zip..."
}

# Legacy DirectX runtimes
if ($installers.Name -like "diretx.exe") {
    Write-Output "Extracting Legacy DirectX Runtimes..."
    $directx = ".\Installers\directx.exe"
    $directxPath = ".\Installers\directx"
    Start-Process -FilePath $directx -WindowStyle Hidden -ArgumentList "/q /c /t:`"$directxPath`"" -Wait
    Write-Output "Installing Legacy DirectX Runtimes..."
    Start-Process -FilePath "$directxPath\dxsetup.exe" -WindowStyle Hidden -ArgumentList '/silent' -Wait
	Write-Output "Finished Installing Legacy DirectX Runtimes"
}

exit