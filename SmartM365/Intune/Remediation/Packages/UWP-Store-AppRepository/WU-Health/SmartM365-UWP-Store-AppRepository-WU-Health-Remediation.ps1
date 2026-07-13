<#
    Name: SmartM365-UWP-Store-AppRepository-WU-Health-Remediation.ps1
    Version: 1.1
    Description: Safely remediates common Windows Update, Microsoft Store, UWP/AppX, AppRepository, and WinRM health issues that may block Windows upgrade scenarios.

    Intended use:
    - Microsoft Intune remediation script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell when possible
    - No forced reboot

    Notes:
    - This script avoids destructive AppRepository repair by default.
    - This script does not reset WindowsApps ACLs recursively because that can break Microsoft Store/UWP servicing.
#>

$ErrorActionPreference = "Stop"

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\WU-Health'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Remediate-UWP-Store-AppRepository-WU-Health.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
$ErrorFound = $false
$RemediationErrors = New-Object System.Collections.Generic.List[string]

function Format-CompactText {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [int]$MaxLength = 180
    )

    $compactText = ($Text -replace "\s+", " ").Trim()

    if ($compactText.Length -gt $MaxLength) {
        return ($compactText.Substring(0, $MaxLength) + "...")
    }

    return $compactText
}

function Write-IntuneResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Status,

        [hashtable]$Data = @{}
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("Status=$Status")

    foreach ($key in ($Data.Keys | Sort-Object)) {
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 240
        $parts.Add(("{0}={1}" -f $key, $value))
    }

    Write-Output ($parts -join "; ")
}

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
}

function Add-RemediationError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-SmartM365Log "ERROR: $Message"
    $script:ErrorFound = $true
    $script:RemediationErrors.Add($Message)
}

function Invoke-ServiceStartIfAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$StartupType = "Manual"
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Add-RemediationError "Service is missing: $Name"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType $StartupType -ErrorAction SilentlyContinue
            Write-SmartM365Log "Service startup type updated: $Name -> $StartupType"
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-SmartM365Log "Service start requested: $Name"
        }
    }
    catch {
        Add-RemediationError "Failed to repair service ${Name}: $($_.Exception.Message)"
    }
}

Write-SmartM365Log "===== Remediation started ====="

# ---------------------------------------------------------
# 1. WinRM
# ---------------------------------------------------------
Write-SmartM365Log "Checking WinRM service"
try {
    $winRmService = Get-Service -Name "WinRM" -ErrorAction SilentlyContinue

    if ($null -eq $winRmService) {
        Add-RemediationError "WinRM service is missing"
    }
    else {
        $winRmCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'" -ErrorAction SilentlyContinue

        if ($null -ne $winRmCim -and $winRmCim.StartMode -eq "Disabled") {
            Set-Service -Name "WinRM" -StartupType Manual -ErrorAction SilentlyContinue
            Write-SmartM365Log "WinRM startup type changed from Disabled to Manual"
        }

        Start-Service -Name "WinRM" -ErrorAction SilentlyContinue
        Write-SmartM365Log "WinRM start requested"
    }
}
catch {
    Add-RemediationError "WinRM remediation failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 2. Windows Update cache remediation
# ---------------------------------------------------------
Write-SmartM365Log "Cleaning Windows Update download cache"
try {
    $servicesToStop = @("wuauserv", "bits", "dosvc")

    foreach ($serviceName in $servicesToStop) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($null -ne $service -and $service.Status -ne "Stopped") {
            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "Service stop requested: $serviceName"
        }
    }

    $downloadCachePath = "C:\Windows\SoftwareDistribution\Download"

    if (Test-Path -Path $downloadCachePath) {
        Get-ChildItem -Path $downloadCachePath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Write-SmartM365Log "Windows Update download cache cleaned"
    }
    else {
        Write-SmartM365Log "Windows Update download cache folder not found; nothing to clean"
    }

    foreach ($serviceName in @("bits", "dosvc", "wuauserv")) {
        Invoke-ServiceStartIfAvailable -Name $serviceName -StartupType Manual
    }
}
catch {
    Add-RemediationError "Windows Update cache remediation failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 3. DISM and SFC
# ---------------------------------------------------------
Write-SmartM365Log "Running DISM RestoreHealth"
try {
    $dismProcess = Start-Process -FilePath "dism.exe" -ArgumentList "/Online", "/Cleanup-Image", "/RestoreHealth" -Wait -NoNewWindow -PassThru
    Write-SmartM365Log "DISM completed with exit code: $($dismProcess.ExitCode)"

    if ($dismProcess.ExitCode -notin @(0, 3010)) {
        Add-RemediationError "DISM returned unexpected exit code: $($dismProcess.ExitCode)"
    }
}
catch {
    Add-RemediationError "DISM execution failed: $($_.Exception.Message)"
}

Write-SmartM365Log "Running SFC scan"
try {
    $sfcProcess = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru
    Write-SmartM365Log "SFC completed with exit code: $($sfcProcess.ExitCode)"

    if ($sfcProcess.ExitCode -notin @(0, 1)) {
        Add-RemediationError "SFC returned unexpected exit code: $($sfcProcess.ExitCode)"
    }
}
catch {
    Add-RemediationError "SFC execution failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 4. UWP/AppX critical services
# ---------------------------------------------------------
Write-SmartM365Log "Repairing UWP/AppX service configuration"
foreach ($serviceName in @("AppXSvc", "ClipSVC", "StateRepository")) {
    Invoke-ServiceStartIfAvailable -Name $serviceName -StartupType Manual
}

# WpnService may be intentionally disabled by hardening baselines. Do not change its startup type automatically.
try {
    $wpnService = Get-Service -Name "WpnService" -ErrorAction SilentlyContinue

    if ($null -eq $wpnService) {
        Write-SmartM365Log "WpnService not found; skipped"
    }
    elseif ($wpnService.Status -in @("StopPending", "PausePending", "Paused")) {
        Start-Service -Name "WpnService" -ErrorAction SilentlyContinue
        Write-SmartM365Log "WpnService start requested because it was in state: $($wpnService.Status)"
    }
    else {
        Write-SmartM365Log "WpnService startup type was not changed"
    }
}
catch {
    Add-RemediationError "WpnService check failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 5. AppRepository sanity check
# ---------------------------------------------------------
Write-SmartM365Log "Checking AppRepository"
try {
    $appRepositoryPath = "C:\ProgramData\Microsoft\Windows\AppRepository"
    $stateRepositoryFile = Join-Path -Path $appRepositoryPath -ChildPath "StateRepository-Machine.srd"

    if (-not (Test-Path -Path $appRepositoryPath)) {
        Add-RemediationError "AppRepository folder is missing"
    }
    elseif (-not (Test-Path -Path $stateRepositoryFile)) {
        Add-RemediationError "StateRepository-Machine.srd is missing"
    }
    else {
        Write-SmartM365Log "AppRepository core database file is present"

        # Do not run esentutl /p automatically.
        # Hard repair may cause data loss and should be performed only after backup and explicit approval.
    }
}
catch {
    Add-RemediationError "AppRepository check failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 6. WindowsApps ACL sanity check
# ---------------------------------------------------------
Write-SmartM365Log "Checking WindowsApps accessibility"
try {
    $windowsAppsPath = "C:\Program Files\WindowsApps"

    if (-not (Test-Path -Path $windowsAppsPath)) {
        Add-RemediationError "WindowsApps folder is missing"
    }
    else {
        Get-ChildItem -Path $windowsAppsPath -Directory -ErrorAction Stop | Select-Object -First 1 | Out-Null
        Write-SmartM365Log "WindowsApps folder is readable"

        # Do not run icacls /reset /t on WindowsApps automatically.
        # Recursive ACL reset can break Microsoft Store/UWP servicing.
    }
}
catch {
    Add-RemediationError "WindowsApps folder is unreadable or ACLs may be broken: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 7. Re-register UWP packages with valid manifests
# ---------------------------------------------------------
Write-SmartM365Log "Re-registering UWP packages with valid manifests"
try {
    $packages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue

    foreach ($package in $packages) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($package.InstallLocation) -and (Test-Path -Path $package.InstallLocation)) {
                $manifestPath = Join-Path -Path $package.InstallLocation -ChildPath "AppxManifest.xml"

                if (Test-Path -Path $manifestPath) {
                    Add-AppxPackage -DisableDevelopmentMode -Register $manifestPath -ErrorAction SilentlyContinue
                    Write-SmartM365Log "Package re-registration requested: $($package.Name)"
                }
            }
        }
        catch {
            Add-RemediationError "Failed to re-register package $($package.Name): $($_.Exception.Message)"
        }
    }
}
catch {
    Add-RemediationError "Global UWP package re-registration failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 8. Microsoft Store remediation
# ---------------------------------------------------------
Write-SmartM365Log "Repairing Microsoft Store registration"
try {
    $storePackages = Get-AppxPackage -AllUsers -Name "Microsoft.WindowsStore" -ErrorAction SilentlyContinue

    if ($null -eq $storePackages) {
        Add-RemediationError "Microsoft Store package is missing; offline package or Microsoft Store source is required for reinstall"
    }
    else {
        foreach ($storePackage in $storePackages) {
            if (-not [string]::IsNullOrWhiteSpace($storePackage.InstallLocation)) {
                $storeManifest = Join-Path -Path $storePackage.InstallLocation -ChildPath "AppxManifest.xml"

                if (Test-Path -Path $storeManifest) {
                    Add-AppxPackage -DisableDevelopmentMode -Register $storeManifest -ErrorAction SilentlyContinue
                    Write-SmartM365Log "Microsoft Store re-registration requested"
                }
                else {
                    Add-RemediationError "Microsoft Store manifest is missing"
                }
            }
            else {
                Add-RemediationError "Microsoft Store package has no InstallLocation"
            }
        }
    }
}
catch {
    Add-RemediationError "Microsoft Store remediation failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 9. UWP frameworks remediation
# ---------------------------------------------------------
Write-SmartM365Log "Re-registering UWP frameworks"
$frameworkPatterns = @(
    "Microsoft.NET.Native.Runtime*",
    "Microsoft.NET.Native.Framework*",
    "Microsoft.VCLibs*",
    "Microsoft.UI.Xaml*"
)

foreach ($frameworkPattern in $frameworkPatterns) {
    try {
        $frameworkPackages = Get-AppxPackage -AllUsers -Name $frameworkPattern -ErrorAction SilentlyContinue

        if ($null -eq $frameworkPackages) {
            Add-RemediationError "Required UWP framework package was not found: $frameworkPattern"
            continue
        }

        foreach ($frameworkPackage in $frameworkPackages) {
            if (-not [string]::IsNullOrWhiteSpace($frameworkPackage.InstallLocation)) {
                $frameworkManifest = Join-Path -Path $frameworkPackage.InstallLocation -ChildPath "AppxManifest.xml"

                if (Test-Path -Path $frameworkManifest) {
                    Add-AppxPackage -DisableDevelopmentMode -Register $frameworkManifest -ErrorAction SilentlyContinue
                    Write-SmartM365Log "Framework re-registration requested: $($frameworkPackage.Name)"
                }
                else {
                    Add-RemediationError "Framework manifest is missing: $($frameworkPackage.Name)"
                }
            }
            else {
                Add-RemediationError "Framework package has no InstallLocation: $($frameworkPackage.Name)"
            }
        }
    }
    catch {
        Add-RemediationError "Framework remediation failed for ${frameworkPattern}: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------
# 10. Final result
# ---------------------------------------------------------
Write-SmartM365Log "===== Remediation finished; no reboot was forced ====="

if ($ErrorFound) {
    Write-SmartM365Log "Script exit code: 1"
    $sampleErrors = @($RemediationErrors | Select-Object -First 3)
    Write-IntuneResult -Status "CompletedWithErrors" -Data @{
        ErrorCount = $RemediationErrors.Count
        LogPath = $LogPath
        Samples = ($sampleErrors -join " | ")
    }
    exit 1
}

Write-SmartM365Log "Script exit code: 0"
Write-IntuneResult -Status "Completed" -Data @{
    LogPath = $LogPath
    Reboot = "NotForced"
}
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCm0umFF9wUqJ0I
# E/R14U4plpLYOCWGRVXdT6dFuTyiraCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEID/apQiyGV1TWq/heIj1heAnKFKkcipDBA5wU+vYT/P5MA0GCSqG
# SIb3DQEBAQUABIIBgCY1CP+3y0wVDAKm5AgdWASy3RC4UWX/e9rSa7K5scLQwYJr
# v8wo5BVr0weYfCeTbY+XLQ9GtbHhwQdThbUrwsVN87SolxvvUELANbcPrfiZS6k2
# PuEZxomZInJvwjfXgFO5VXe4CH98OSyx/OYo6r5tBfXkEq1x9JdqvIpHo6+SZTNB
# YTPS6gTvfPcFxXPExW/Qp+VvDwmH/KtFbEs36ehNO18yH4xySFs+TSfijLMplGCg
# 9IzzXz0zY2kBz+Q4CQ58eQaWwWuvhPCjvsJXJT+tEtMC/wjRYP9JJItB5Ecvih69
# h2JygqydtTT+rTaJQsMXLGL/6DggaAzhqvEGm+hpyuPob56OrYlOQQZdQTHzAN9k
# +yhHyE6XPRXtld6je7Bfa4Ai0CvSDOERPIkfVSBIiRxPdW1hOOz4Hk6vuVVcq5es
# /J3DOSpoCaBI5oHbYNrWopcppQbWlrWzfJ+jmz0Y8Ltb1MI9PWAArg+h3HM/mX9F
# B0BXlgXQJeMGZ6egkKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# NTlaMC8GCSqGSIb3DQEJBDEiBCBSXvyfrpoqEaPOUnF+jrRwXlPCG2Ed7T9iC2cy
# ICKkRzANBgkqhkiG9w0BAQEFAASCAgBikWCZdOn3vaRPXw145bgh2X6V/DtulGo/
# 3d4OUWheOxeL4KPc52cEiFjuEXfbTF95nfxVw4mAdj2kuDWWicAI9i4xQkwTydhq
# 5uguK6CD3vUFROApmk4bS1lIbI4eg4RJWKvf9eIzeCfcjUmWDJ8xM8EFu0NslfF0
# 2/4sf+/OlfIrtzH3MDyvdYpS2LlPPKn7cHdrliW211Q8J2tD1WEr6OCYl9zAyZt0
# GCt/u3ewVj88sGlOC9he3zHmiInMrZ+hFu4+O7mfpkYPdMlsAZkihEV4ppUVyrR0
# 2Hm90C/P6il3fgGyeRp6w3Q5T3Zzu/5J+TZrnrZ6Z7sWIzSaJ8RPyNdGLNlvlJeV
# vTZgh6wDlZl1iNpUW9IkiI95sfys1PGNHYVi2nOyPcUb4JYi1LluFR3kwCZ3OLNL
# n3tMFvjCO4WPIJKKg5M/qetidp+PqvFh35/IexOscsaaTZwLUlrDH5kd3vFfsVkH
# i+hfhBFxSJwy7r0Ajwnz5+4OXfSqQVsayVQ8v7eypbcOY80LF9GQM2I/lpxGx9KO
# FZpbWpzulFPDGj3XMMdiJkiikOKbLzBjSaQ1gBYLur1h493yDDXlF2PxH+fuzRbP
# UtVlKfgwcU3B6+ilC8OmKbsJsY37T39ZDFJ+y+a7gCIDTrmRHnnk/0fH0TxJUmHt
# D8bnsdt2Tw==
# SIG # End signature block
