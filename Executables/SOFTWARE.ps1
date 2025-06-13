param (
	[switch]$Chrome,
	[switch]$Brave,
	[switch]$Firefox,
	[switch]$OperaGX,
	[switch]$OpenShell,
	[switch]$ModifyUi
)

.\AtmosphereModules\initPowerShell.ps1

# ----------------------------------------------------------------------------------------------------------- #
# Software is no longer installed with a package manager anymore to be as fast and as reliable as possible.   #
# ----------------------------------------------------------------------------------------------------------- #

# To the people who wrote line 12 - 14... I'm sorry. Needed to

$timeouts = @("--connect-timeout", "10", "--retry", "5", "--retry-delay", "0", "--retry-all-errors")
$msiArgs = "/qn /quiet /norestart ALLUSERS=1 REBOOT=ReallySuppress"
$arm = ((Get-CimInstance -Class Win32_ComputerSystem).SystemType -match 'ARM64') -or ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64')
$DesktopPath = [Environment]::GetFolderPath('Desktop')

# Create temporary directory
function Remove-TempDirectory { Pop-Location; Remove-Item -Path $tempDir -Force -Recurse -EA 0 }
$tempDir = Join-Path -Path $(Get-SystemDrive) -ChildPath $([System.Guid]::NewGuid())
New-Item $tempDir -ItemType Directory -Force | Out-Null
Push-Location $tempDir

# Brave
if ($Brave) {
	Write-Output "Downloading Brave..."
	& curl.exe -LSs "https://laptop-updates.brave.com/latest/winx64" -o "$tempDir\BraveSetup.exe" $timeouts
	if (!$?) {
		Write-Error "Downloading Brave failed."
		exit 1
	}

	Write-Output "Installing Brave..."
	Start-Process -FilePath "$tempDir\BraveSetup.exe" -WindowStyle Hidden -ArgumentList '/silent /install'

	do {
		$processesFound = Get-Process | Where-Object { "BraveSetup" -contains $_.Name } | Select-Object -ExpandProperty Name
		if ($processesFound) {
			Write-Output "Still running BraveSetup."
			Start-Sleep -Seconds 2
		} else {
			Remove-TempDirectory
		}
	} until (!$processesFound)

	Stop-Process -Name "brave" -Force -EA 0
	Write-Output "Brave installed successfully."
	exit
}

# Firefox
if ($Firefox) {
	$firefoxArch = ('win64', 'win64-aarch64')[$arm]

	Write-Output "Downloading Firefox..."
	& curl.exe -LSs "https://download.mozilla.org/?product=firefox-latest-ssl&os=$firefoxArch&lang=en-US" -o "$tempDir\firefox.exe" $timeouts
	Write-Output "Installing Firefox..."
	Start-Process -FilePath "$tempDir\firefox.exe" -WindowStyle Hidden -ArgumentList '/S /ALLUSERS=1' -Wait
	Write-Output "Firefox installed successfully."
	Remove-TempDirectory
	exit
}

# Chrome
if ($Chrome) {
	Write-Output "Downloading Google Chrome..."
	$chromeArch = ('64', '_Arm64')[$arm]
	& curl.exe -LSs "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise$chromeArch.msi" -o "$tempDir\chrome.msi" $timeouts
	Write-Output "Installing Google Chrome..."
	Start-Process -FilePath "$tempDir\chrome.msi" -WindowStyle Hidden -ArgumentList '/qn' -Wait
	Write-Output "Google Chrome installed successfully."
	Remove-TempDirectory
	exit
}

# Opera GX
if ($OperaGX) {
	Write-Output "Downloading Opera GX..."
#	$operaArch = ('x64', 'arm64')[$arm] # Opera GX does not have an ARM64 version, so we always use x64
	# if system is x64 we will default to firefox
	if ($arm) {
		Write-Output "Detected arm architecture"
		Write-Output "Downloading firefox instead"
		& curl.exe -LSs "https://download.mozilla.org/?product=firefox-latest-ssl&os=win64-aarch64&lang=en-US" -o "$tempDir\firefox.exe" $timeouts
		Write-Output "Installing Firefox..."
		Start-Process -FilePath "$tempDir\firefox.exe" -WindowStyle Hidden -ArgumentList '/S /ALLUSERS=1' -Wait
		Write-Output "Firefox installed successfully."
		Remove-TempDirectory
		exit
	}
	& curl.exe -LSs "https://download.opera.com/download/get/?id=71824&location=424&nothanks=yes&sub=marine&utm_tryagain=yes" -o "$tempDir\OperaGXSetup.exe" $timeouts
	Write-Output "Installing Opera GX..."
	Start-Process -FilePath "$tempDir\OperaGXSetup.exe" -WindowStyle Hidden -ArgumentList ' /install /silent /norestart /launchopera=0 /setdefaultbrowser=1 /allusers=1' -Wait
	Write-Output "Opera GX installed successfully."
	Remove-TempDirectory
	exit
}

# Open-Shell
if ($OpenShell) {
	Write-Output "Downloading Open-Shell..."
	& curl.exe -LSs "https://github.com/Open-Shell/Open-Shell-Menu/releases/download/v4.4.196/OpenShellSetup_4_4_196.exe" -o "$tempDir\OpenShellSetup.exe" $timeouts
	Write-Output "Installing Open-Shell..."
	Start-Process -FilePath "$tempDir\OpenShellSetup.exe" -WindowStyle Hidden -ArgumentList '/qn /quiet ADDLOCAL=StartMenu' -Wait
	Write-Output "Open-Shell installed successfully."
	Remove-TempDirectory
	exit
}

# Modify Windows Ui
if ($ModifyUi) {
	function Invoke-WingetInstallFallback {
    param (
        [Parameter(Mandatory = $true)][string]$PackageId,
        [string]$DisplayName = $PackageId
    )

    Write-Warning "Falling back to winget for $DisplayName..."

    try {
        winget install --id=$PackageId -e --silent --accept-source-agreements --accept-package-agreements
        if ($LASTEXITCODE -eq 0) {
            Write-Output "$DisplayName installed successfully via winget."
            return $true
        } else {
            Write-Error "winget failed to install $DisplayName (ExitCode: $LASTEXITCODE)."
            return $false
        }
    } catch {
        Write-Error "Exception while using winget to install ${DisplayName}: $_"
        return $false
    }
}
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
    $taskXmlFilePath = Join-Path $scriptDir "AccentColorizer.xml"

    if (-not (Test-Path $taskXmlFilePath)) {
        Write-Error "AccentColorizer.xml not found at '$taskXmlFilePath'. Skipping scheduled task registration."
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
    } catch {
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
            } else {
                Write-Error "regsvr32 command exited with non-zero code $($process.ExitCode). This may indicate an issue with the DLL or registration."
                return
            }
        } else {
            Write-Error "regsvr32 command timed out after $timeoutSeconds seconds. Attempting to terminate process."
            $process | Stop-Process -Force -ErrorAction SilentlyContinue
            return
        }
    } catch {
        Write-Error "Failed to start or register DLL: $($_.Exception.Message)"
        return 
    }
# TranslucentFlyouts
	Write-Output "Installing TranslucentFlyouts..."
	$p = Start-Process -FilePath "C:\Windows\AtmosphereModules\Tools\TranslucentFlyouts\launch_win32.cmd" -PassThru
	$p.WaitForExit()
	Write-Output "TranslucentFlyouts installed..."
}

#####################
##    Utilities    ##
#####################

# MS Store App Fetcher 
# https://github.com/Ameliorated-LLC/appfetch
# Experimental and has known compatibility issues with some apps.
Write-Output "Downloading MS Store App Fetcher..."
$githubApi = Invoke-RestMethod "https://api.github.com/repos/Ameliorated-LLC/appfetch/releases/latest" -EA 0
$exeUrl = $githubApi.assets.browser_download_url | Where-Object { $_ -like "*.exe" } | Select-Object -First 1
$exePath = Join-Path $tempDir "App Fetch.exe"
& curl.exe -L $exeUrl -o $exePath
Move-Item -Path "$exePath" -Destination "$DesktopPath\App Fetch.exe" -Force
Write-Output "MS Store App Fetcher downloaded to your Desktop."
Remove-TempDirectory

Start-Sleep -Seconds 1 # please dont touch this, it will break vcredists curl for some reason

# Visual C++ Runtimes (referred to as vcredists for short)
# https://learn.microsoft.com/en-US/cpp/windows/latest-supported-vc-redist
$legacyArgs = '/q /norestart'
$modernArgs = "/install /quiet /norestart"
$vcredistDir = Join-Path -Path $(Get-SystemDrive) -ChildPath "vcredistDir"
New-Item $vcredistDir -ItemType Directory -Force | Out-Null
Push-Location $vcredistDir

$vcredists = [ordered] @{
	# 2005 - version 8.0.50727.6195 (MSI 8.0.61000/8.0.61001) SP1
	"https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.exe" = @("2005-x64", "/c /q /t:")
	"https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x86.exe" = @("2005-x86", "/c /q /t:")
	# 2008 - version 9.0.30729.6161 (EXE 9.0.30729.5677) SP1
	"https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x64.exe" = @("2008-x64", "/q /extract:")
	"https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x86.exe" = @("2008-x86", "/q /extract:")
	# 2010 - version 10.0.40219.325 SP1
	"https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe" = @("2010-x64", $legacyArgs)
	"https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x86.exe" = @("2010-x86", $legacyArgs)
	# 2012 - version 11.0.61030.0
	"https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe" = @("2012-x64", $modernArgs)
	"https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe" = @("2012-x86", $modernArgs)
	# 2013 - version 12.0.40664.0
	"https://aka.ms/highdpimfc2013x64enu" = @("2013-x64", $modernArgs)
	"https://aka.ms/highdpimfc2013x86enu" = @("2013-x86", $modernArgs)
	# 2015-2022 (2015+) - latest version
	"https://aka.ms/vs/17/release/vc_redist.x64.exe" = @("2015+-x64", $modernArgs)
	"https://aka.ms/vs/17/release/vc_redist.x86.exe" = @("2015+-x86", $modernArgs)
}
foreach ($a in $vcredists.GetEnumerator()) {
	$vcName = $a.Value[0]
	$vcArgs = $a.Value[1]
	$vcUrl = $a.Name
	$vcExePath = "$vcredistDir\vcredist-$vcName.exe"

	Write-Host "`nProcessing Visual C++ Runtime: $vcName"
	Write-Host "Temp directory: $vcredistDir"
	Write-Host "Executable path: $vcExePath"
	Write-Host "Download URL: $vcUrl`n"

	
	# curl is faster than Invoke-WebRequest
	Write-Output "Downloading and installing Visual C++ Runtime $vcName..."
	& curl.exe -LSs "$vcUrl" -o "$vcExePath" $timeouts

	if ($vcArgs -match ":") {
		$msiDir = "$vcredistDir\vcredist-$vcName"
		Start-Process -FilePath $vcExePath -ArgumentList "$vcArgs`"$msiDir`"" -Wait -WindowStyle Hidden
		
		$msiPaths = (Get-ChildItem -Path $msiDir -Filter *.msi -EA 0).FullName
		if (!$msiPaths) {
			Write-Output "Failed to extract MSI for $vcName, not installing."
		} else {
			$msiPaths | ForEach-Object {
				Start-Process -FilePath "msiexec.exe" -ArgumentList "/log `"$msiDir\logfile.log`" /i `"$_`" $msiArgs" -WindowStyle Hidden
			}
		}
	} else {
		Start-Process -FilePath $vcExePath -ArgumentList $vcArgs -Wait -WindowStyle Hidden
	}
}

Pop-Location; Remove-Item -Path $vcredistDir -Force -Recurse -EA 0

# NanaZip
function Install7Zip {
	$website = 'https://7-zip.org/'
	$7zipArch = ('x64', 'arm64')[$arm]
	$download = $website + ((Invoke-WebRequest $website -UseBasicParsing).Links.href | Where-Object { $_ -like "a/7z*-$7zipArch.exe" })
	Write-Output "Downloading 7-Zip..."
	& curl.exe -LSs $download -o "$tempDir\7zip.exe" $timeouts
	Write-Output "Installing 7-Zip..."
	Start-Process -FilePath "$tempDir\7zip.exe" -WindowStyle Hidden -ArgumentList '/S' -Wait
}

$githubApi = Invoke-RestMethod "https://api.github.com/repos/M2Team/NanaZip/releases/latest" -EA 0
$assets = $githubApi.Assets.browser_download_url | Select-String ".xml", ".msixbundle" | Select-Object -Unique -First 2
function InstallNanaZip {
	# NanaZip is slow but im too lazy to default to 7zip
	Write-Output "Downloading NanaZip..."	
	$path = New-Item "$tempDir\nanazip" -ItemType Directory
	$assets | ForEach-Object {
		$filename = $_ -split '/' | Select-Object -Last 1
		Write-Output "Downloading '$filename'..."
		& curl.exe -LSs $_ -o "$path\$filename" $timeouts
	}

	Write-Output "Installing NanaZip..."	
	try {
		$appxArgs = @{
			"PackagePath" = (Get-ChildItem $path -Filter "*.msixbundle" | Select-Object -First 1).FullName
			"LicensePath" = (Get-ChildItem $path -Filter "*.xml" | Select-Object -First 1).FullName
		}
		Add-AppxProvisionedPackage -Online @appxArgs | Out-Null
		
		Write-Output "Installed NanaZip!"
	} catch {
		Write-Error "Failed to install NanaZip! Getting 7-Zip instead. $_"
		Install7Zip
	}
}

if ($assets.Count -eq 2) {
	$7zipRegistry = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\7-Zip"
	if (Test-Path $7zipRegistry) {
		$Message = @'
Would you like to uninstall 7-Zip and replace it with NanaZip?

NanaZip is a fork of 7-Zip with an updated user interface and extra features.
'@

		if ((Read-MessageBox -Title 'Installing NanaZip - Atmosphere' -Body $Message -Icon Question) -eq 'Yes') {
			$7zipUninstall = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\7-Zip" -Name "QuietUninstallString" -EA 0).QuietUninstallString
			Write-Output "Uninstalling 7-Zip..."
			Start-Process -FilePath "cmd" -WindowStyle Hidden -ArgumentList "/c $7zipUninstall" -Wait
			InstallNanaZip
		}
	} else {
		InstallNanaZip
	}
} else {
	Write-Error "Can't access GitHub API, downloading 7-Zip instead of NanaZip."
	Install7Zip
}

# Legacy DirectX runtimes
& curl.exe -LSs "https://download.microsoft.com/download/8/4/A/84A35BF1-DAFE-4AE8-82AF-AD2AE20B6B14/directx_Jun2010_redist.exe" -o "$tempDir\directx.exe" $timeouts
Write-Output "Extracting legacy DirectX runtimes..."
Start-Process -FilePath "$tempDir\directx.exe" -WindowStyle Hidden -ArgumentList "/q /c /t:`"$tempDir\directx`"" -Wait
Write-Output "Installing legacy DirectX runtimes..."
Start-Process -FilePath "$tempDir\directx\dxsetup.exe" -WindowStyle Hidden -ArgumentList '/silent' -Wait

# Remove temporary directory
Remove-TempDirectory