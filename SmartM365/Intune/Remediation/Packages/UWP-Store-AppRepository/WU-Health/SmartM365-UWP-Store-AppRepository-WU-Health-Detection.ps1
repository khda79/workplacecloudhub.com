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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB6+aLA/o0oK82R
# +TANrwo7VUICY+zaKnQqtOzi7SPusqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDr7V2Is5vP2HzQyqRXVGMF32qqanXAfFvaDfiOFYIxzTANBgkqhkiG9w0B
# AQEFAASCAYAD1STPIRGVI5S1eLaaCFjOj1duzc/LkvWDeX4mBrYlD2eS/yzinRbr
# 98fQAiJiCmMD5ad00628OEQ0r7618KZdVv7hRLkXOoJebSZu1GKeh/cfNmr3+HEK
# be4nptg+T3b/iFxZNEU3rvH/LKSEoW9b6+aPVWXqprtYNsWQuzfOghRRkS5jOPFt
# XwU+1ek1Dk9txTyK3YBX2Cr+8vvBbJVF0F2KghSfFZ9L3LAd+nQVVm/DG6zTXZem
# LtdXG7F3cqKM3ftfhCadDK0TlwHSLg3ez7Mn1bjCC1K+lm5ndFRoW1RZ9RZnIxHQ
# D00pG6VsB5AfKOnFHEm04sIJRKhsK3sEaXVzGV6nnM5+EbSp5v2x7SoLQ2LhDOU4
# fPD5LZufw/9+81j+FS96PgUp5wcTABFRChx4DJIli+rtsGK9DcbqrJVRb9SSuvr/
# vZmCE1hWJhHd1yVP1/ZU59gYY0Dli93h1O8kjB7Cb9x287DVqss8cousP1c5lf5p
# mEBBDGAPo8c=
# SIG # End signature block
