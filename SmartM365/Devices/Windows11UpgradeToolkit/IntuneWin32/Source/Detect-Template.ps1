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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBoFxNQEk48hoiE
# EtJzaKFx/SsDICrR/pqnlSr23PymXqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCDj3rnbnhN/xCVVdNXLz5hGPC+3u0bW+LJsnum2ntjWxDANBgkqhkiG9w0B
# AQEFAASCAYB+4cH/EEFVF2QWbnmFVmj8iH/LhRfcHEiMqxGbDg5+vEWOl4XqnTi3
# R6ScIPHfPPg3E9AiXgQFOhhxuh85snKzi6VBQ6Df1MEDUXAtdn3oVlYA5tPGywsw
# vV4NB1xogkBYdQiZd0NgO/RJnbtU2k0hHvAOXhhWz1rk/TOzq2pKh7EOshvJN+AD
# SDkS2fjy3nq52nIgMOo9m4KEanuv9jeTagQjDMgmLxuuaPki3gMVQzAaOnzKEtxV
# Wk+cB9HeedymcAj8cwGTI7hqXqGIw8swmfgT8ftD0IFFoNm+eFYjH7sNdFY9wAIg
# 3sPAmw4HdbLZVI4Jo56vHgXLJcvabjSva9hK8aw8+wnaQ1SgtRD3fCMyXh3ac4zb
# rJUKqZlbB2fz5xFrz2u5GDgxgPU2yOnzw/k214UahO83GHBDdWLsjEJzgQ3XdbKY
# I6KrQPz499zZDjMAJ/HmFnMS23Mx/mORE7GIFwVrS95VDMrkLBp4/2ldD1W1CyQk
# LHDZ8OAFvyI=
# SIG # End signature block
