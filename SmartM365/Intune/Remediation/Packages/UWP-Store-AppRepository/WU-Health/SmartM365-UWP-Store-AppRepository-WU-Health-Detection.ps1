<# 
    Name: SmartM365-UWP-Store-AppRepository-WU-Health-Detection.ps1
    Version: 1.1
    Description: Detects issues that may prevent Windows upgrade or UWP/AppX operations because Microsoft Store, AppRepository, WindowsApps, or Windows Update cache components are corrupted or unavailable.

    Intended use:
    - Microsoft Intune detection script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell when possible
#>

$ErrorActionPreference = "Stop"

$issues = New-Object System.Collections.Generic.List[string]

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

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $script:issues.Add($Message)
}

function Test-ServiceHealth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [bool]$RequireRunning = $false
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if ($null -eq $service) {
        Add-Issue "Required service is missing: $Name"
        return
    }

    $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

    if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
        Add-Issue "Required service is disabled: $Name"
        return
    }

    if ($RequireRunning -and $service.Status -ne "Running") {
        Add-Issue "Required service is not running: $Name"
        return
    }

    if ($service.Status -in @("StopPending", "PausePending", "Paused")) {
        Add-Issue "Required service is in an unhealthy state: $Name ($($service.Status))"
    }
}

try {
    # ---------------------------------------------------------
    # 1. WindowsApps folder health
    # ---------------------------------------------------------
    $windowsAppsPath = "C:\Program Files\WindowsApps"

    if (-not (Test-Path -Path $windowsAppsPath)) {
        Add-Issue "WindowsApps folder is missing"
    }
    else {
        try {
            $windowsAppsItems = Get-ChildItem -Path $windowsAppsPath -Directory -ErrorAction Stop

            if ($windowsAppsItems.Count -lt 50) {
                Add-Issue "WindowsApps contains too few package folders; possible corruption"
            }

            $knownPackages = @(
                "Microsoft.WindowsStore*",
                "Microsoft.DesktopAppInstaller*",
                "Microsoft.VCLibs*",
                "Microsoft.NET.Native.Runtime*",
                "Microsoft.NET.Native.Framework*"
            )

            foreach ($knownPackage in $knownPackages) {
                $match = $windowsAppsItems | Where-Object { $_.Name -like $knownPackage } | Select-Object -First 1
                if ($null -eq $match) {
                    Add-Issue "Expected WindowsApps package folder was not found: $knownPackage"
                }
            }
        }
        catch {
            Add-Issue "WindowsApps folder is unreadable or permissions appear broken"
        }
    }

    # ---------------------------------------------------------
    # 2. AppRepository health
    # ---------------------------------------------------------
    $appRepositoryPath = "C:\ProgramData\Microsoft\Windows\AppRepository"

    if (-not (Test-Path -Path $appRepositoryPath)) {
        Add-Issue "AppRepository folder is missing"
    }
    else {
        try {
            $appRepositoryItems = Get-ChildItem -Path $appRepositoryPath -ErrorAction Stop

            if ($appRepositoryItems.Count -lt 20) {
                Add-Issue "AppRepository contains too few files; possible corruption"
            }

            $requiredAppRepositoryFiles = @(
                "StateRepository-Machine.srd",
                "StateRepository-Machine.srd-shm",
                "StateRepository-Machine.srd-wal"
            )

            foreach ($requiredFile in $requiredAppRepositoryFiles) {
                $requiredFilePath = Join-Path -Path $appRepositoryPath -ChildPath $requiredFile
                if (-not (Test-Path -Path $requiredFilePath)) {
                    Add-Issue "Required AppRepository file is missing: $requiredFile"
                }
            }
        }
        catch {
            Add-Issue "AppRepository folder is unreadable or corrupted"
        }
    }

    # ---------------------------------------------------------
    # 3. Critical UWP/AppX services
    # ---------------------------------------------------------
    # These services do not always need to be running permanently.
    # The detection focuses on missing or disabled services, and unhealthy transient states.
    $criticalServices = @(
        "AppXSvc",
        "ClipSVC",
        "StateRepository"
    )

    foreach ($serviceName in $criticalServices) {
        Test-ServiceHealth -Name $serviceName -RequireRunning $false
    }

    # WpnService may be disabled by hardening baselines and is not always required for Windows upgrade.
    # It is checked as informational only when present.
    $wpnService = Get-Service -Name "WpnService" -ErrorAction SilentlyContinue
    if ($null -ne $wpnService -and $wpnService.Status -in @("StopPending", "PausePending", "Paused")) {
        Add-Issue "WpnService is in an unhealthy state: $($wpnService.Status)"
    }

    # ---------------------------------------------------------
    # 4. Microsoft Store package registration
    # ---------------------------------------------------------
    $storePackage = Get-AppxPackage -AllUsers -Name "Microsoft.WindowsStore" -ErrorAction SilentlyContinue

    if ($null -eq $storePackage) {
        Add-Issue "Microsoft Store package is missing or not registered"
    }
    else {
        $storePackageLocations = $storePackage | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.InstallLocation) -and
            (Test-Path -Path $_.InstallLocation)
        }

        if ($null -eq $storePackageLocations) {
            Add-Issue "Microsoft Store package is registered but its install location is missing"
        }
    }

    # ---------------------------------------------------------
    # 5. Essential UWP frameworks
    # ---------------------------------------------------------
    $frameworks = @(
        "Microsoft.NET.Native.Runtime",
        "Microsoft.NET.Native.Framework",
        "Microsoft.VCLibs",
        "Microsoft.UI.Xaml"
    )

    $allAppxPackages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue

    foreach ($framework in $frameworks) {
        $package = $allAppxPackages | Where-Object { $_.Name -like "$framework*" } | Select-Object -First 1

        if ($null -eq $package) {
            Add-Issue "Required UWP framework is missing: $framework"
        }
    }

    # ---------------------------------------------------------
    # 6. Windows Update SoftwareDistribution cache sanity
    # ---------------------------------------------------------
    $softwareDistributionDownloadPath = "C:\Windows\SoftwareDistribution\Download"
    $softwareDistributionIsEmpty = $false

    if (-not (Test-Path -Path $softwareDistributionDownloadPath)) {
        Add-Issue "Windows Update download cache folder is missing"
    }
    else {
        $downloadFiles = Get-ChildItem -Path $softwareDistributionDownloadPath -Recurse -File -ErrorAction SilentlyContinue
        $downloadDirectories = Get-ChildItem -Path $softwareDistributionDownloadPath -Directory -ErrorAction SilentlyContinue

        # An empty Download folder is not automatically corruption.
        # It becomes suspicious only if recent Windows Update/AppX errors are also present.
        $softwareDistributionIsEmpty = (($null -eq $downloadFiles -or $downloadFiles.Count -eq 0) -and ($null -eq $downloadDirectories -or $downloadDirectories.Count -eq 0))
    }

    # ---------------------------------------------------------
    # 7. WinRM health
    # ---------------------------------------------------------
    # WinRM is not required for Windows upgrade itself, but can be required by enterprise remediation tooling.
    $winRmService = Get-Service -Name "WinRM" -ErrorAction SilentlyContinue

    if ($null -eq $winRmService) {
        Add-Issue "WinRM service is missing"
    }
    else {
        $winRmServiceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'" -ErrorAction SilentlyContinue

        if ($null -ne $winRmServiceCim -and $winRmServiceCim.StartMode -eq "Disabled") {
            Add-Issue "WinRM service is disabled"
        }
    }

    # ---------------------------------------------------------
    # 8. Event Log checks for AppX / Store / UWP / Windows Update
    # ---------------------------------------------------------
    $appxErrors = Get-WinEvent -LogName "Microsoft-Windows-AppXDeploymentServer/Operational" -MaxEvents 100 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LevelDisplayName -eq "Error" -and
            $_.TimeCreated -ge (Get-Date).AddDays(-7)
        }

    if ($null -ne $appxErrors -and $appxErrors.Count -gt 0) {
        Add-Issue "Recent AppX deployment errors were detected in the event log"
    }

    $windowsUpdateErrors = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 100 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LevelDisplayName -eq "Error" -and
            $_.TimeCreated -ge (Get-Date).AddDays(-7)
        }

    if ($null -ne $windowsUpdateErrors -and $windowsUpdateErrors.Count -gt 0) {
        Add-Issue "Recent Windows Update errors were detected in the event log"
    }

    if ($softwareDistributionIsEmpty -and (($null -ne $windowsUpdateErrors -and $windowsUpdateErrors.Count -gt 0) -or ($null -ne $appxErrors -and $appxErrors.Count -gt 0))) {
        Add-Issue "SoftwareDistribution download cache is empty while recent Windows Update or AppX errors exist"
    }

    # ---------------------------------------------------------
    # Final result
    # ---------------------------------------------------------
    if ($issues.Count -gt 0) {
        $sampleIssues = @($issues | Select-Object -First 5)
        Write-IntuneResult -Status "IssuesDetected" -Data @{
            IssueCount = $issues.Count
            Samples = ($sampleIssues -join " | ")
        }
        exit 1
    }

    Write-IntuneResult -Status "Healthy" -Data @{
        Checks = "UWP,Store,AppRepository,WindowsApps,WinRM,WindowsUpdate"
    }
    exit 0
}
catch {
    Write-IntuneResult -Status "Error" -Data @{
        Message = $_.Exception.Message
    }
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAd+9IaTyIj6suX
# O6cGzVOEjzrxWnem/TlKTB9kUIQMTqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKCb3Q22UyAIX8VfhuqobWV7wbZVxAgRqEidaBKW1g/0MA0GCSqG
# SIb3DQEBAQUABIIBgGtsfnViGaQyHg8WKJrx6AJ7r5WvIbGhdYoXOmUINojDzM3x
# kFYJs4IdUi7ckHAmBf4B6vFWQsIP5LDY5rRdJRw+wXbjmr24Sa0EBTbjwgfqB1aS
# gepxdwCfM3fqXLVMNO7d+SjGHgs1NLgsLp8+Evis5o7001uUadqy0O4NnAIVtSf7
# KPESrxEO/BARa/rvzN+vKoApozYWsfVC58bkfi+WXYK/ZnQnm+X5GpG7/+15O+Hs
# N/dE3gIYrtEIkYYAmAv9aC7uHCqQVxhGGJuHrQ60UK5U+79hqOSL4IMLX+SMiecd
# wzFacBZdMx127lvf9tF3EJd4o6N4q8pD7TtZIi3cezPWi8keNBpwpg++G01OSpVP
# bNzQERn9rv+s2//dBlH5+OIG7tl4X3qIxilnYYiwCQjEsuA9SJJcL8S6JmE9TwSU
# EgcTzp2739vGHLjknoNJ/6pQWmW1jDVh3sUwvozjOGIiqj5o65MX5NtuOOOCF5Ia
# uT91VuaPlATRDEMkr6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# NTlaMC8GCSqGSIb3DQEJBDEiBCDiAlfpLganVaOFdMPf08AMxoVl9e8xVVVOqhA+
# ajnGojANBgkqhkiG9w0BAQEFAASCAgBjRymvHq3FPdlTP37gmXx7iAeGL8acVFqb
# 91GRd2UZeGJ9R127VT6M57UELs5NvPuDjYK/94d8r/Woaz0ax9s4rD4swRvsTtWS
# LIcYdykrN8U0XxYdLel2LIPvDZuN2eRhHLKHGRmPSBNJ1D4xfNclQQNsO4sHB4qp
# pNgIb/d0zwJdiaB3YuYM2fnfgTJA+7O5D3TP1MtpzP5YuHtve6S/U/sAV8QhzeMv
# 3YlDDWYrdw8wHX6zJGT/FYQWmWzBzoIp7mtFk+RaFeoWkl5Msk6oGBN2PFdDhDMv
# OX5uEjptbrNA9Pde2YNjtdUnsFAZBAYuDSfLEWDVKgZk1SyfZUF2ICMsKx1359Cx
# +I14+1BZiZVttY1KgT2sLByikQLqKiZ5GuidT0epZ9Nj5L9IQByC+QZICYxaK0Ws
# JWBHNv610AdgkZZdWRqj6laYCBJFPIDy9fqxVE2aWYob1Dk5e1nPJTbdFciG2Rnl
# DvkhMuTGo01Mrlwlp87VHrlzfKTpozi7JkS/XWeYm+24fXUgLWymkze4YNkOebtA
# WC0MCpbS89TEMHO35pRbgeMCHQMlzevv6eCJr0U1CJo6SShA48HBiUxgLiJwL+9r
# Ow86ys9gKvZ8Mt1z0u8z0tvTeIZeP63ttqSpvHW8Oc5u+oZSdFWht7VIpyelZfyk
# 6rsMup93ZA==
# SIG # End signature block
