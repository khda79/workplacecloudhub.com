<#
.SYNOPSIS
    Detects a generated SmartM365 Windows 11 Upgrade Toolkit Intune package.
.DESCRIPTION
    Template used by the package builder to generate a language/package-specific Intune detection script.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
$packageId = '__PACKAGE_ID__'
$packageVersion = '__PACKAGE_VERSION__'
$registryPath = "HKLM:\SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages\$packageId"

try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    if ([int]$os.BuildNumber -ge 22000 -or ([string]$os.Caption) -match 'Windows 11') { exit 0 }
}
catch { }

try {
    $item = Get-ItemProperty -LiteralPath $registryPath -ErrorAction Stop
    if ([string]$item.PackageId -ne $packageId) { exit 1 }
    if ([string]$item.PackageVersion -ne $packageVersion) { exit 1 }
    exit 0
}
catch {
    exit 1
}
