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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCClNEknyzyLKmis
# OvxKjLExrp8Z8wTo2rMWXMwsRb4kyqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBovQKa/kny/IRfRIZA
# vlwfoRYYv44Vp9wrXq7eAbjdRzANBgkqhkiG9w0BAQEFAASCAYCsMpPpMEcXD6Pt
# jVnReNbguZWRqAnuRxclyVoyiqyUU7lenHBt4Y0FbWGC0x27V7K+YtFUa0s7lULy
# a1PgOC8f3Uba40EIKNvQrIirEvkLfjjPyUfXi3IuqaeTzuHICzu7mN6ZgB6J8Zpk
# FEDulcjnW4RdGOXqOrbqK8Ljs35WmA3M+C7Lzc8P2QSiRH9K+1Jw9nIbwLwMnkKD
# fG01cp1dBTH/R9zVmbB6RshAG7MoWBHpRxvHD50QsuiNuh1cDoaoR3HSjG5NMBUl
# X6QsDO/aPVZNxeMirNKt6ZInbzsAtNPaZD3dqg8zTL7OBS/+JE6yfyqzk0DuHEuB
# ENpXIVKHINQgQ2iJq3B6uuveqgWfQXLRxerIeaSObfT2kpOsZbMfdaM8EcB9mPVV
# miz1M5+fi6vXSAGsYasDJMhjfignwRQvOTVpXevyQVoeZypNg2iMDZf/rBdwSN63
# VOgrnyW5A8m8pmNrPVQIouS6vA8mZNnCpZGoDQm6lSLynS5UnYQ=
# SIG # End signature block
