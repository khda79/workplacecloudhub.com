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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASg82wXHB8KNCX
# tVFGTikfKK6Eh/sk0q1IIe9+V6/ZbKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAPVrcX8VGxfnLysaGM
# 8aRI24Ce68s5gYNoqdNa+kSs8zANBgkqhkiG9w0BAQEFAASCAYB4OA6o2pHVXsHK
# psHPAJDUT2imJn/j+8IEQWVp1tNHh6di9AfWf0CMrNyiPLrKzibiTll48MTSbUpH
# LhKm7EfjFDfmCfzpchriKhxH9erxyaEbb74C085kFvnmok2WOvIx9hKL30FamBoE
# OTHv2y5Fp9M8lqZIukU8+RzWROU/azqGpnK18r5aGha5fJ0HIiSmv2pught2GSH4
# 4J6bQIK1Y+hNkI+BRfsAZVdkOX97VO/FQ1+2nsq8Zz9kJ8tmTKOR0FQbS1kMdpK7
# D0pTPJ81D1c8U83XB39WAn+7JmnILgQ2xgU8dC10f3qXQDl+6ifGclB5kfRvCW7r
# 99yG7pGLMBCRMMPDolqNjNaK5LM+sW5Ejn51UqbTcS/gUEfnKCJ4xGWlzs3ckCv/
# L1XjX6+82qZ7doTo3hW0H9Ha62zYAMnL9Ld67EOOpgwZ5Vo5cvUcu32SqZhmKvdf
# OOnPrF3oYHBptKot2JbEpSMoPzuclyg8/lpi39RhQ+IFiNxnZxI=
# SIG # End signature block
