<#
    Name: SmartM365-DeliveryOptimization-ContentEngine-Health-Remediation.ps1
    Version: 1.1
    Description: Safely remediates Delivery Optimization and BITS content engine state with compact Intune output and detailed local logging.

    Intended use:
    - Microsoft Intune remediation script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell when possible
    - No forced reboot

    Logs:
    - C:\ProgramData\SmartM365\IntuneRemediation\Logs\DeliveryOptimization-ContentEngine-Health\Remediate-DeliveryOptimization-ContentEngine-Health.log
#>

[CmdletBinding()]
param(
    [bool]$ResetBitsJobs = $true,
    [bool]$TriggerWindowsUpdateScan = $true,
    [bool]$CleanWindowsUpdateDownloadCache = $true
)

$ErrorActionPreference = "Stop"

$ScenarioName = "DeliveryOptimization-ContentEngine-Health"
$RemediationName = "Remediate-DeliveryOptimization-ContentEngine-Health"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$ScenarioName"
$LogPath = Join-Path -Path $LogRoot -ChildPath "$RemediationName.log"

if (-not (Test-Path -LiteralPath $LogRoot)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}

$ErrorFound = $false
$Actions = New-Object System.Collections.Generic.List[string]
$Skipped = New-Object System.Collections.Generic.List[string]
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
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 260
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

function Add-Action {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-SmartM365Log $Message
    $script:Actions.Add($Message)
}

function Add-SkippedAction {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-SmartM365Log $Message
    $script:Skipped.Add($Message)
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

function Invoke-ServiceStopIfRunning {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Add-SkippedAction "ServiceNotFound=$Name"
            return
        }

        if ($service.Status -ne "Stopped") {
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Add-Action "ServiceStopRequested=$Name"
        }
        else {
            Add-SkippedAction "ServiceAlreadyStopped=$Name"
        }
    }
    catch {
        Add-RemediationError "Failed to stop service ${Name}: $($_.Exception.Message)"
    }
}

function Invoke-ServiceStartIfAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Add-SkippedAction "ServiceNotFound=$Name"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
            Add-Action "ServiceStartupTypeChanged=$Name StartupType=Manual"
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Add-Action "ServiceStartRequested=$Name"
        }
        else {
            Add-SkippedAction "ServiceAlreadyRunning=$Name"
        }
    }
    catch {
        Add-RemediationError "Failed to start service ${Name}: $($_.Exception.Message)"
    }
}

function Clear-FolderContentSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            Add-SkippedAction "CleanupNotFound=$Path"
            return
        }

        $removedCount = 0
        $skippedCount = 0
        Write-SmartM365Log "CleanupStart=$Path"

        $items = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)

        foreach ($item in $items) {
            try {
                Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
                $removedCount++
            }
            catch {
                $skippedCount++
                Write-SmartM365Log "CleanupItemSkipped=$($item.FullName) Message=$($_.Exception.Message)"
            }
        }

        Add-Action "CleanupCompleted=$Path Removed=$removedCount Skipped=$skippedCount"
    }
    catch {
        Add-RemediationError "Failed to clean folder ${Path}: $($_.Exception.Message)"
    }
}

function Invoke-BitsTransferResetSafe {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$ResetBitsJobs
    )

    try {
        if (-not $ResetBitsJobs) {
            Add-SkippedAction "BITSReset=Skipped"
            return
        }

        $bitsAdminPath = Join-Path -Path $env:WINDIR -ChildPath "System32\bitsadmin.exe"

        if (-not (Test-Path -LiteralPath $bitsAdminPath -PathType Leaf)) {
            Add-SkippedAction "BITSReset=BitsadminNotFound"
            return
        }

        $process = Start-Process -FilePath $bitsAdminPath -ArgumentList "/reset", "/allusers" -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
        Add-Action "BITSReset=Completed ExitCode=$($process.ExitCode)"
    }
    catch {
        Add-RemediationError "Failed to reset BITS transfers: $($_.Exception.Message)"
    }
}

function Invoke-WindowsUpdateScanSafe {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$TriggerWindowsUpdateScan
    )

    try {
        if (-not $TriggerWindowsUpdateScan) {
            Add-SkippedAction "WindowsUpdateScan=Skipped"
            return
        }

        $usoClientPath = Join-Path -Path $env:WINDIR -ChildPath "System32\UsoClient.exe"

        if (-not (Test-Path -LiteralPath $usoClientPath -PathType Leaf)) {
            Add-SkippedAction "WindowsUpdateScan=UsoClientNotFound"
            return
        }

        Start-Process -FilePath $usoClientPath -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
        Add-Action "WindowsUpdateScan=Triggered"
    }
    catch {
        Add-RemediationError "Failed to trigger Windows Update scan: $($_.Exception.Message)"
    }
}

try {
    Write-SmartM365Log "===== Delivery Optimization remediation started ====="
    Write-SmartM365Log "ResetBitsJobs=$ResetBitsJobs TriggerWindowsUpdateScan=$TriggerWindowsUpdateScan CleanWindowsUpdateDownloadCache=$CleanWindowsUpdateDownloadCache"

    $deliveryOptimizationCachePaths = @(
        "C:\ProgramData\Microsoft\Windows\DeliveryOptimization\Cache",
        "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
    )

    foreach ($serviceName in @("DoSvc", "BITS")) {
        Invoke-ServiceStopIfRunning -Name $serviceName
    }

    Start-Sleep -Seconds 2

    foreach ($cachePath in $deliveryOptimizationCachePaths) {
        Clear-FolderContentSafe -Path $cachePath
    }

    if ($CleanWindowsUpdateDownloadCache) {
        foreach ($serviceName in @("wuauserv", "UsoSvc")) {
            Invoke-ServiceStopIfRunning -Name $serviceName
        }

        $windowsUpdateDownloadCache = Join-Path -Path $env:WINDIR -ChildPath "SoftwareDistribution\Download"
        Clear-FolderContentSafe -Path $windowsUpdateDownloadCache
    }
    else {
        Add-SkippedAction "WindowsUpdateDownloadCacheCleanup=Skipped"
    }

    Invoke-BitsTransferResetSafe -ResetBitsJobs $ResetBitsJobs

    foreach ($serviceName in @("BITS", "DoSvc", "wuauserv", "UsoSvc")) {
        Invoke-ServiceStartIfAvailable -Name $serviceName
    }

    Start-Sleep -Seconds 3

    Invoke-WindowsUpdateScanSafe -TriggerWindowsUpdateScan $TriggerWindowsUpdateScan

    Write-SmartM365Log "===== Delivery Optimization remediation finished ====="

    $sampleErrors = ""
    if ($RemediationErrors.Count -gt 0) {
        $sampleErrors = (($RemediationErrors | Select-Object -First 3) -join " | ")
    }

    $sampleActions = (($Actions | Select-Object -First 5) -join " | ")

    if ($ErrorFound) {
        Write-IntuneResult -Status "CompletedWithErrors" -Data @{
            ActionCount = $Actions.Count
            SkippedCount = $Skipped.Count
            ErrorCount = $RemediationErrors.Count
            Errors = $sampleErrors
            Log = $LogPath
        }

        exit 1
    }

    Write-IntuneResult -Status "Completed" -Data @{
        ActionCount = $Actions.Count
        SkippedCount = $Skipped.Count
        Actions = $sampleActions
        Log = $LogPath
    }

    exit 0
}
catch {
    try {
        Add-RemediationError $_.Exception.Message
    }
    catch {
        $null = $_
    }

    Write-IntuneResult -Status "TechnicalError" -Data @{
        Message = $_.Exception.Message
        Log = $LogPath
    }

    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASg82wXHB8KNCX
# tVFGTikfKK6Eh/sk0q1IIe9+V6/ZbKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAPVrcX8VGxfnLysaGM8aRI24Ce68s5gYNoqdNa+kSs8zANBgkqhkiG9w0B
# AQEFAASCAYBqvSKDYq07pxVnBskmWcxZ3SXf8SaAyDRg+7r5wMnPwvrIjD3uATEW
# bTSAlymqFG4th2V8gbIeBLN7GazKuzsD994SPTS/Gg/qkWaotTUIff89yHX+JmQP
# 33jd33zQSDG0P7UVdrByIVSkbvVp2uVZCeoJ5cMZn3CAMfii0dT9Te7xm0kFqNVB
# 9JiY5iPoUaN5xGYjnT2shK6n49SN5V8WTQFvi0h05lHXWaGB1hAOLh2fJX8uMY3B
# TFp8FN/rQUddMoQR1W9IulFtQVTCJx1PGOB/RcENnm+FgNqs2FswcnqrLu+Etpaa
# TbMiMaoP/Y0uCBcQz3CfV8Jha9/kGIRcuS0LpesJyFn539unTO9t4dLZRds79tkp
# xcTkxKbeEP+okPvViYRfQappfdhrjUN05P/qnOa5WG4qudpfTa8Sc6rnO7llfZC8
# grCd9sbA5O7MflCgenf6Vx+MUP9a3z5uyAnt/e5JScQ04J/VdPFfL8QeG6puojib
# QPwWQhtcfFc=
# SIG # End signature block
