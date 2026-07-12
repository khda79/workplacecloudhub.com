<#
    Name: SmartM365-Upgrade-Diagnostics-Detection.ps1
    Version: 1.0
    Description: Consolidated SetupDiag and Windows upgrade blocker detection.
#>

[CmdletBinding()]
param(
    [int]$LookbackDays = 7,
    [int]$MaxIssuesToDisplay = 5,
    [int]$MaxIssueLength = 180,
    [int]$MaxAnalysisAgeDays = 7
)

$ErrorActionPreference = "Stop"
$SetupDiagResultPath = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Output\SetupDiag\SetupDiagResults.xml"
$SetupDiagToolPaths = @(
    (Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Tools\SetupDiag\SetupDiag.exe"),
    (Join-Path $env:ProgramData "SetupDiag\SetupDiag.exe")
)
$EnterpriseRegistryPath = "HKLM:\SOFTWARE\SmartM365\IntuneRemediation\SetupDiag"

$BlockingCodes = @(
    "0xC1900208",
    "0xC1900101",
    "0x80070070",
    "0x80240020",
    "0x80242016",
    "0x800F0922",
    "0xC1900200",
    "0x8007042B",
    "0x8007001F"
)

$RollbackPaths = @(
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
    (Join-Path $env:WINDIR "Panther\NewOS\Rollback"),
    (Join-Path $env:SystemDrive 'Windows.old\$Windows.~BT\Sources\Rollback')
)

$SetupLogRoots = @(
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
    (Join-Path $env:WINDIR "Panther"),
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Panther'),
    (Join-Path $env:SystemDrive 'Windows.old\$Windows.~BT\Sources\Rollback')
)

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Issues.Contains($Message)) {
        $Issues.Add($Message)
    }
}

function Add-EventIssue {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)][string]$Source,
        [AllowNull()]$Events,
        [Parameter(Mandatory = $true)][string[]]$Codes,
        [Parameter(Mandatory = $true)][int]$MaxIssueLength
    )

    foreach ($eventRecord in @($Events)) {
        if ($null -eq $eventRecord -or [string]::IsNullOrWhiteSpace($eventRecord.Message)) {
            continue
        }

        foreach ($code in $Codes) {
            if ($eventRecord.Message -match [regex]::Escape($code)) {
                $message = $eventRecord.Message -replace "`r|`n", " "
                if ($message.Length -gt $MaxIssueLength) {
                    $message = $message.Substring(0, $MaxIssueLength) + "..."
                }
                Add-Issue -Issues $Issues -Message "$Source Id=$($eventRecord.Id) Code=$code Message=$message"
            }
        }
    }
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $cutoffDate = (Get-Date).AddDays(-1 * $LookbackDays)
    $hasSetupDiagResult = Test-Path -LiteralPath $SetupDiagResultPath -PathType Leaf
    $hasSetupDiagTool = $false
    foreach ($setupDiagToolPath in $SetupDiagToolPaths) {
        if (Test-Path -LiteralPath $setupDiagToolPath -PathType Leaf) {
            $hasSetupDiagTool = $true
            break
        }
    }
    $hasSetupLogs = $false
    $hasRollbackLogs = $false

    foreach ($setupLogRoot in $SetupLogRoots) {
        if (Test-Path -LiteralPath $setupLogRoot) {
            $hasSetupLogs = $true
            break
        }
    }

    foreach ($rollbackPath in $RollbackPaths) {
        if (Test-Path -LiteralPath $rollbackPath) {
            $hasRollbackLogs = $true
            Add-Issue -Issues $issues -Message "RollbackDetected=$rollbackPath"
        }
    }

    if ($hasSetupLogs -and -not $hasSetupDiagResult) {
        if ($hasSetupDiagTool) {
            Add-Issue -Issues $issues -Message "SetupLogsDetectedWithoutSetupDiagAnalysis"
        }
        else {
            Add-Issue -Issues $issues -Message "SetupLogsDetectedSetupDiagToolMissing"
        }
    }

    if ($hasRollbackLogs -and -not $hasSetupDiagResult) {
        Add-Issue -Issues $issues -Message "RollbackLogsDetectedWithoutSetupDiagAnalysis"
    }

    if ($hasSetupDiagResult) {
        $xmlAgeDays = (New-TimeSpan -Start (Get-Item -LiteralPath $SetupDiagResultPath).LastWriteTimeUtc -End (Get-Date).ToUniversalTime()).TotalDays
        if ($xmlAgeDays -gt $MaxAnalysisAgeDays) {
            Add-Issue -Issues $issues -Message ("SetupDiagResultOutdated AgeDays={0:N1}" -f $xmlAgeDays)
        }

        $xmlContent = Get-Content -LiteralPath $SetupDiagResultPath -Raw -ErrorAction SilentlyContinue
        foreach ($code in $BlockingCodes) {
            if ($xmlContent -match [regex]::Escape($code)) {
                Add-Issue -Issues $issues -Message "SetupDiagDetected=$code"
            }
        }

        if ($xmlContent -match "Matching Profile found:\s*(.+)") {
            Add-Issue -Issues $issues -Message "SetupDiagProfile=$($matches[1].Trim())"
        }
    }
    elseif (Test-Path -Path $EnterpriseRegistryPath) {
        $setupDiagState = Get-ItemProperty -Path $EnterpriseRegistryPath -Name LastRunUtc -ErrorAction SilentlyContinue
        if ($setupDiagState -and -not [string]::IsNullOrWhiteSpace($setupDiagState.LastRunUtc)) {
            $analysisAgeDays = (New-TimeSpan -Start ([datetime]::Parse($setupDiagState.LastRunUtc)) -End (Get-Date).ToUniversalTime()).TotalDays
            if ($analysisAgeDays -gt $MaxAnalysisAgeDays) {
                Add-Issue -Issues $issues -Message ("SetupDiagRegistryStateOutdated AgeDays={0:N1}" -f $analysisAgeDays)
            }
        }
    }

    $wuEvents = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 500 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $cutoffDate }
    Add-EventIssue -Issues $issues -Source "WUEvent" -Events $wuEvents -Codes $BlockingCodes -MaxIssueLength $MaxIssueLength

    $setupEvents = Get-WinEvent -LogName "Setup" -MaxEvents 300 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $cutoffDate }
    Add-EventIssue -Issues $issues -Source "SetupEvent" -Events $setupEvents -Codes $BlockingCodes -MaxIssueLength $MaxIssueLength

    if ($issues.Count -gt 0) {
        Write-Output "Status=UpgradeDiagnosticsRequired"
        Write-Output "IssueCount=$($issues.Count)"
        $issues | Select-Object -First $MaxIssuesToDisplay | ForEach-Object { Write-Output $_ }
        exit 1
    }

    Write-Output "Status=Healthy"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKUu/if6iIZ8nq
# I/UAxiYk2/wExXsyT1opD/MGuw06eaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCD5pJK6aAvjh9llLqUVLSlQ2kw+H4Y+MXjDnDS4MA4K/zANBgkqhkiG9w0B
# AQEFAASCAYBWVRg/9gKKG0gjFwANnlnzDZ+vKBJ0s0MuG5z0tpf5rlJ2gAtynXbQ
# /t7We44yOBo+TpJoPkiDQJz/xJBB2th6OLL4YtVd53mwNMMQmut33eRxFD+oyQ1p
# DTL8j6pmIJn4HProvlibfaAt1PNVuB7t9TnvCBlXcX2FdWH8Jyem16+3eB7Zo7mo
# yxt1nScQF0JPAjVv464WJ8kRUKmslBLbtY+PdNN0TKU8YzkH+zIWc2OY+FSwpIfA
# mnfJzTr6gDWgJ2iUj4w/XDOFF+AYWmKQmk5gIv0bAj2qrvwwDBpjYAKDpN41s6OJ
# /45GLVQPpExKQhgcStAyllJGZGkX9r1dSgQh1uFvifnCUVt86HH3Og25nWYnuvme
# YmpOTgt/ewEzDdJuPMfv9lSEpgGQi6BfAF1ImgF5RYK9nmGS/CanmghiKBv+eLDS
# wn+oKz9VWAYxiB5Jj00u7N3jc288szGXBZD7/cXzrD7SfCWkHJhLvVLmiE2Xaxe9
# noa64etANxM=
# SIG # End signature block
