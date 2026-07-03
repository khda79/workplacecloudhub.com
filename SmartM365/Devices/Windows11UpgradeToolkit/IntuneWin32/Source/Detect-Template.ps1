<#
.SYNOPSIS
    Detects a generated SmartM365 Windows 11 Upgrade Toolkit Intune package.
.DESCRIPTION
    Template used by the package builder to generate a language/package-specific Intune detection script.
.VERSION
    1.0.1
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
$packageId = '__PACKAGE_ID__'
$packageVersion = '__PACKAGE_VERSION__'
$registrySubKey = "SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages\$packageId"

function Get-Registry64PackageState {
    param([string]$SubKey)

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key = $baseKey.OpenSubKey($SubKey)
        if (-not $key) { return $null }
        try {
            return [pscustomobject]@{
                PackageId = [string]$key.GetValue('PackageId', '')
                PackageVersion = [string]$key.GetValue('PackageVersion', '')
            }
        }
        finally { $key.Dispose() }
    }
    finally { $baseKey.Dispose() }
}

try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    if ([int]$os.BuildNumber -ge 22000 -or ([string]$os.Caption) -match 'Windows 11') { exit 0 }
}
catch { }

try {
    $item = Get-Registry64PackageState -SubKey $registrySubKey
    if ($null -eq $item) { exit 1 }
    if ([string]$item.PackageId -ne $packageId) { exit 1 }
    if ([string]$item.PackageVersion -ne $packageVersion) { exit 1 }
    exit 0
}
catch {
    exit 1
}
