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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB6+aLA/o0oK82R
# +TANrwo7VUICY+zaKnQqtOzi7SPusqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDr7V2Is5vP2HzQyqRX
# VGMF32qqanXAfFvaDfiOFYIxzTANBgkqhkiG9w0BAQEFAASCAYAv/XnbM/3hrPp/
# pZ+zR94HDN89BAaW124WvCznzYc0ayO6PMlubrrjPy18TUHkdWJVbVq5wRhreI+R
# dKFO6dDE0jvg57/9XBxCoJxvnJAJMlE0wvYx2nl35J5YQrqL/AgBc4kzxhqdeR9e
# qVRcAcOWlKIRCquCDr9dipMj6THidMZOyLY5CS7DqYmb8oxMgQva1lvQ96rRpVM9
# yybM3IzZieDmJq2fFxs7SaEqL6QxZpFxFzvptdb5g6yN53DJ141zqOEV1tXOgK/6
# lyJSHfmIXILS6EGLmEUxOftaBoEMjDJA846rhqYhUGUaNcEacYuLNGdgKWxDSacD
# Q/EHgZ7LSQMStz9zZDQerFyNigxUSZHms/WFv5MOW+UrCYyj+9kOunnHIuGwe90L
# ajamG6v7roRoYcaaH7hMLtTNxffxfRbJvD+YpID1Mao8mTzpSl2B8yYMmbyMj53G
# Y822Q0eteKdNx3KCiFuVdzaGxewfT5ifiNIUkkHhMv73z0jRTEE=
# SIG # End signature block
