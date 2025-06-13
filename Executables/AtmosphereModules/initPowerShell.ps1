$windir = [Environment]::GetFolderPath('Windows')

# Add Atmosphere's PowerShell modules
$env:PSModulePath += ";$windir\AtmosphereModules\Scripts\Modules"