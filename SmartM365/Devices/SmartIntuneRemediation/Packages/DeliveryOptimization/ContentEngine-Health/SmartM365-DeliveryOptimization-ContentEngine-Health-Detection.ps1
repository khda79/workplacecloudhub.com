<#
    Name: SmartM365-DeliveryOptimization-ContentEngine-Health-Detection.ps1
    Version: 1.1
    Description: Detects actionable Delivery Optimization, Dynamic Download, BITS, and content download issues without flagging common informational event noise.

    Detection goals:
    - Keep Intune output compact and single-line
    - Does not treat a missing Delivery Optimization folder as an error
    - Does not treat empty event logs as an error
    - Ignores BITS peer helper event 310 with 0x80070032, which is common non-actionable noise
    - Detects Delivery Optimization and Dynamic Download errors
    - Detects BITS download errors related to content transfer
    - Detects Windows Update errors only when clearly related to Delivery Optimization or Dynamic Download

    Exit codes:
    0 = Healthy
    1 = Issues detected
#>

[CmdletBinding()]
param(
    [int]$MaxEvents = 200,
    [int]$LookbackDays = 7,
    [int]$MaxIssuesToDisplay = 5,
    [int]$MaxIssueLength = 180,
    [int]$MinimumCacheSizeBytes = 1024
)

$ErrorActionPreference = "Stop"

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
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 320
        $parts.Add(("{0}={1}" -f $key, $value))
    }

    Write-Output ($parts -join "; ")
}

function Get-EventMessage {
    param(
        [AllowNull()]
        [object]$EventRecord
    )

    if ($null -eq $EventRecord -or [string]::IsNullOrWhiteSpace($EventRecord.Message)) {
        return ""
    }

    return (($EventRecord.Message -replace "`r|`n", " ").Trim())
}

function Get-RecentEventsSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogName,

        [Parameter(Mandatory = $true)]
        [datetime]$CutoffDate,

        [Parameter(Mandatory = $true)]
        [int]$MaxEvents
    )

    try {
        $events = @(
            Get-WinEvent -LogName $LogName -MaxEvents $MaxEvents -ErrorAction SilentlyContinue |
            Where-Object { $_.TimeCreated -ge $CutoffDate }
        )

        if ($null -eq $events) {
            return @()
        }

        return $events
    }
    catch {
        return @()
    }
}

function Test-BitsPeerHelperNoise {
    param(
        [Parameter(Mandatory = $true)]
        [object]$EventRecord
    )

    $message = Get-EventMessage -EventRecord $EventRecord

    if ($EventRecord.Id -ne 310) {
        return $false
    }

    if ($message -match "0x80070032") {
        return $true
    }

    if ($message -match "peer helper|modules auxiliaires|Peerhilfs") {
        return $true
    }

    return $false
}

function Test-DoRelatedMessage {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    return (
        $Message -match "Delivery Optimization" -or
        $Message -match "DeliveryOptimization" -or
        $Message -match "DynamicDownload" -or
        $Message -match "0x80D0[0-9A-Fa-f]{4}"
    )
}

function Test-ActionableMessage {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    return (
        $Message -match "0x80D0[0-9A-Fa-f]{4}" -or
        $Message -match "\berror\b" -or
        $Message -match "\bfail(?:ed|ure)?\b" -or
        $Message -match "\btimeout\b" -or
        $Message -match "\bproxy\b" -or
        $Message -match "download.*fail" -or
        $Message -match "fail.*download"
    )
}

function Test-ActionableContentEvent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [object]$EventRecord
    )

    $message = Get-EventMessage -EventRecord $EventRecord

    if ([string]::IsNullOrWhiteSpace($message)) {
        return $false
    }

    switch ($Source) {
        "DO" {
            return (Test-ActionableMessage -Message $message)
        }
        "WU" {
            return ((Test-DoRelatedMessage -Message $message) -and (Test-ActionableMessage -Message $message))
        }
        "BITS" {
            if (Test-BitsPeerHelperNoise -EventRecord $EventRecord) {
                return $false
            }

            return (
                $message -match "0x80D0[0-9A-Fa-f]{4}" -or
                $message -match "0x802000[0-9A-Fa-f]{2}" -or
                $message -match "\bBG_E_" -or
                ((Test-DoRelatedMessage -Message $message) -and (Test-ActionableMessage -Message $message)) -or
                ($message -match "download.*fail" -or $message -match "fail.*download")
            )
        }
    }

    return $false
}

function Add-EventIssue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Events,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$IssueList,

        [Parameter(Mandatory = $true)]
        [int]$Limit,

        [Parameter(Mandatory = $true)]
        [int]$MaxIssueLength
    )

    if ($null -eq $Events -or $Events.Count -eq 0) {
        return
    }

    foreach ($eventRecord in $Events) {
        if ($null -eq $eventRecord) {
            continue
        }

        if ($IssueList.Count -ge $Limit) {
            return
        }

        if (-not (Test-ActionableContentEvent -Source $Source -EventRecord $eventRecord)) {
            continue
        }

        $message = Format-CompactText -Text (Get-EventMessage -EventRecord $eventRecord) -MaxLength $MaxIssueLength
        $IssueList.Add("${Source}:Time=$($eventRecord.TimeCreated);Id=$($eventRecord.Id);Message=$message")
    }
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $cutoffDate = (Get-Date).AddDays(-1 * $LookbackDays)
    $ignoredNoiseCount = 0
    $checkedEventCount = 0

    $deliveryOptimizationFolder = "C:\ProgramData\Microsoft\Windows\DeliveryOptimization"
    $deliveryOptimizationFolderStatus = "NotCreated"
    $deliveryOptimizationSizeMB = 0

    if (Test-Path -Path $deliveryOptimizationFolder) {
        $deliveryOptimizationFiles = @(Get-ChildItem -Path $deliveryOptimizationFolder -Recurse -File -ErrorAction SilentlyContinue)
        $deliveryOptimizationSizeBytes = 0

        if ($deliveryOptimizationFiles.Count -gt 0) {
            $deliveryOptimizationSizeBytes = ($deliveryOptimizationFiles | Measure-Object -Property Length -Sum).Sum

            if ($null -eq $deliveryOptimizationSizeBytes) {
                $deliveryOptimizationSizeBytes = 0
            }
        }

        $deliveryOptimizationSizeMB = [math]::Round(($deliveryOptimizationSizeBytes / 1MB), 2)
        $deliveryOptimizationFolderStatus = "Present"

        if ($deliveryOptimizationSizeBytes -lt $MinimumCacheSizeBytes) {
            $deliveryOptimizationFolderStatus = "EmptyOrSmall"
        }
    }

    $windowsUpdateEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    $deliveryOptimizationEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-DeliveryOptimization/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    $bitsEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-Bits-Client/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    $checkedEventCount = $windowsUpdateEvents.Count + $deliveryOptimizationEvents.Count + $bitsEvents.Count
    $ignoredNoiseCount = @($bitsEvents | Where-Object { Test-BitsPeerHelperNoise -EventRecord $_ }).Count

    Add-EventIssue `
        -Source "WU" `
        -Events $windowsUpdateEvents `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    Add-EventIssue `
        -Source "DO" `
        -Events $deliveryOptimizationEvents `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    Add-EventIssue `
        -Source "BITS" `
        -Events $bitsEvents `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    $uniqueIssues = @($issues | Select-Object -Unique)

    if ($uniqueIssues.Count -gt 0) {
        $sourceSummary = ($uniqueIssues | ForEach-Object {
            if ($_ -match "^([^:]+):") { $matches[1] } else { "Unknown" }
        } | Group-Object | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ","

        Write-IntuneResult -Status "DeliveryOptimizationContentEngineIssuesDetected" -Data @{
            IssueCount = $uniqueIssues.Count
            Sources = $sourceSummary
            CheckedEvents = $checkedEventCount
            IgnoredNoise = $ignoredNoiseCount
            DOFolderStatus = $deliveryOptimizationFolderStatus
            DOFolderSizeMB = $deliveryOptimizationSizeMB
            Examples = (($uniqueIssues | Select-Object -First $MaxIssuesToDisplay) -join " | ")
        }

        exit 1
    }

    Write-IntuneResult -Status "DeliveryOptimizationContentEngineHealthy" -Data @{
        CheckedEvents = $checkedEventCount
        IgnoredNoise = $ignoredNoiseCount
        DOFolderStatus = $deliveryOptimizationFolderStatus
        DOFolderSizeMB = $deliveryOptimizationSizeMB
    }

    exit 0
}
catch {
    Write-IntuneResult -Status "TechnicalError" -Data @{
        Message = $_.Exception.Message
    }

    exit 1
}
