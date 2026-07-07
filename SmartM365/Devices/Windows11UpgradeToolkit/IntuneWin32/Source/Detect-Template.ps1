<#
.SYNOPSIS
    Detects a generated SmartM365 Windows 11 Upgrade Toolkit Intune package.
.DESCRIPTION
    Template used by the package builder to generate a language/package-specific Intune detection script.
.VERSION
    1.0.2
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
                InstallState = [string]$key.GetValue('InstallState', '')
            }
        }
        finally { $key.Dispose() }
    }
    finally { $baseKey.Dispose() }
}

function Test-PackageVersionAtLeast {
    param(
        [string]$Actual,
        [string]$Minimum
    )

    if ([string]::IsNullOrWhiteSpace($Minimum)) { return $true }
    if ([string]$Actual -eq [string]$Minimum) { return $true }

    $actualVersion = $null
    $minimumVersion = $null
    if ([version]::TryParse([string]$Actual, [ref]$actualVersion) -and [version]::TryParse([string]$Minimum, [ref]$minimumVersion)) {
        return ($actualVersion -ge $minimumVersion)
    }

    return $false
}

function Complete-Detected {
    param([string]$Reason)

    Write-Output ("OK: {0}" -f $Reason)
    exit 0
}

try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    if ([int]$os.BuildNumber -ge 22000 -or ([string]$os.Caption) -match 'Windows 11') { Complete-Detected -Reason 'Device is already Windows 11' }
}
catch {
    $null = $_
}

try {
    $item = Get-Registry64PackageState -SubKey $registrySubKey
    if ($null -eq $item) { exit 1 }
    if ([string]$item.PackageId -ne $packageId) { exit 1 }
    if (-not [string]::IsNullOrWhiteSpace($item.InstallState) -and @('Installed', 'AlreadyWindows11') -notcontains [string]$item.InstallState) { exit 1 }
    if (-not (Test-PackageVersionAtLeast -Actual ([string]$item.PackageVersion) -Minimum $packageVersion)) { exit 1 }
    Complete-Detected -Reason ("Package registry state found. InstalledVersion={0}; RequiredVersion={1}; InstallState={2}" -f $item.PackageVersion,$packageVersion,$item.InstallState)
}
catch {
    exit 1
}
