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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBoFxNQEk48hoiE
# EtJzaKFx/SsDICrR/pqnlSr23PymXqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDj3rnbnhN/xCVVdNXL
# z5hGPC+3u0bW+LJsnum2ntjWxDANBgkqhkiG9w0BAQEFAASCAYB3RT9vnuknMPx+
# 7KyD8Ll6iCoozQAe3RW5HhDRijAQye2x14gxhsSPrKurLSmVz/oKzncUBKsDhtQ8
# uvMpTud9lakKNlr7++gElcQdVE/VsJKoOrmeO78LBDRTPymW9AX8d9XVGJLELXQI
# KDSoB3ZgSz8OXUNvPt3zRdyWhDVj3svlQwAOiwjt92RXE/dxwaYWyH4Ub1nKhc/Y
# thWQ7XfVJ3b/3n4U9FwK93AXH0iJ57tg7i2UPP2PWlQfOKPgRyD4wLq4TB7x4JCb
# I2HF+c7ekBgbV5EU/nAZKt0CUam4IkJ5h7MrAJBJbE6uv1B2S7ZZ8Z64LfsQBq4i
# gej5IbxovfnSEFXKObm9qJvNF9HZNviM5OVmziiwacJ9e63tYwBGm8vCyTnLAUNw
# sxWGRS3N5KJyA9fa9gaYR4EliJ0bVB8OuPfRHVUpfWQZbNTEHsdRo7JZG+xcXNZL
# zNs2C+3tvsidnAUcpTyGvODiyFEaqy08KG20sBGowAzvlgG4CoI=
# SIG # End signature block
