<#
    Name: SmartM365-DeliveryOptimization-ContentEngine-Health-Detection.ps1
    Version: 1.0
    Description: Detects Delivery Optimization, Dynamic Download, BITS, and Windows Update content engine issues.

    Detection goals:
    - Does not treat a missing Delivery Optimization folder as an error
    - Does not treat empty event logs as an error
    - Detects recent Delivery Optimization errors
    - Detects 0x80D0xxxx errors
    - Detects DynamicDownload errors
    - Detects BITS download errors
    - Detects Windows Update errors related to Delivery Optimization or content download

    Exit codes:
    0 = Healthy
    1 = Issues detected
    2 = Technical error
#>

[CmdletBinding()]
param(
    [int]$MaxEvents = 200,
    [int]$LookbackDays = 7,
    [int]$MaxIssuesToDisplay = 5,
    [int]$MaxIssueLength = 220,
    [int]$MinimumCacheSizeBytes = 1024
)

$ErrorActionPreference = "Stop"

function Test-EventMessageMatch {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Event,

        [Parameter(Mandatory = $true)]
        [string[]]$Patterns
    )

    if ($null -eq $Event -or [string]::IsNullOrWhiteSpace($Event.Message)) {
        return $false
    }

    foreach ($pattern in $Patterns) {
        if ($Event.Message -match $pattern) {
            return $true
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
        [string[]]$Patterns,

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.List[string]]$IssueList,

        [Parameter(Mandatory = $true)]
        [int]$Limit,

        [Parameter(Mandatory = $true)]
        [int]$MaxIssueLength
    )

    if ($null -eq $IssueList) {
        return
    }

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

        if ([string]::IsNullOrWhiteSpace($eventRecord.Message)) {
            continue
        }

        if (Test-EventMessageMatch -Event $eventRecord -Patterns $Patterns) {
            $message = ($eventRecord.Message -replace "`r|`n", " ")

            if ($message.Length -gt $MaxIssueLength) {
                $message = $message.Substring(0, $MaxIssueLength) + "..."
            }

            $IssueList.Add("${Source} Time=$($eventRecord.TimeCreated) Id=$($eventRecord.Id) Message=$message")
        }
    }
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

try {
    Write-Output "=== Delivery Optimization / Content Engine Diagnostic ==="
    Write-Output "Version=1.0"
    Write-Output "LookbackDays=$LookbackDays"
    Write-Output "MaxEvents=$MaxEvents"

    $issues = New-Object System.Collections.Generic.List[string]
    $cutoffDate = (Get-Date).AddDays(-1 * $LookbackDays)

    # =========================================================
    # Delivery Optimization cache validation
    # =========================================================
    $deliveryOptimizationFolder = "C:\ProgramData\Microsoft\Windows\DeliveryOptimization"
    $deliveryOptimizationFolderStatus = "OK"

    if (-not (Test-Path -Path $deliveryOptimizationFolder)) {
        Write-Output "DOFolderStatus=NotCreated"
        Write-Output "Info=Delivery Optimization folder does not exist; this is not considered an error."
        $deliveryOptimizationFolderStatus = "NotCreated"
    }
    else {
        $deliveryOptimizationFiles = Get-ChildItem -Path $deliveryOptimizationFolder -Recurse -File -ErrorAction SilentlyContinue
        $deliveryOptimizationSizeBytes = 0

        if ($null -ne $deliveryOptimizationFiles -and $deliveryOptimizationFiles.Count -gt 0) {
            $deliveryOptimizationSizeBytes = (
                $deliveryOptimizationFiles |
                Measure-Object -Property Length -Sum
            ).Sum

            if ($null -eq $deliveryOptimizationSizeBytes) {
                $deliveryOptimizationSizeBytes = 0
            }
        }

        $deliveryOptimizationSizeMB = [math]::Round(($deliveryOptimizationSizeBytes / 1MB), 2)

        Write-Output "DOFolderStatus=Present"
        Write-Output "DOFolderSizeMB=$deliveryOptimizationSizeMB"

        if ($deliveryOptimizationSizeBytes -lt $MinimumCacheSizeBytes) {
            $deliveryOptimizationFolderStatus = "EmptyOrSmall"
            Write-Output "Info=Delivery Optimization cache is present but empty or very small; this is not an error by itself."
        }
    }

    # =========================================================
    # Windows Update events
    # =========================================================
    $windowsUpdatePatterns = @(
        "DynamicDownload",
        "Delivery Optimization",
        "DeliveryOptimization",
        "\bContent\b",
        "0x80D0[0-9A-Fa-f]{4}",
        "0x80070002",
        "0x8024000C",
        "download.*fail",
        "fail.*download"
    )

    $windowsUpdateEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    Write-Output "WUEventCount=$($windowsUpdateEvents.Count)"

    Add-EventIssue `
        -Source "WU" `
        -Events $windowsUpdateEvents `
        -Patterns $windowsUpdatePatterns `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    # =========================================================
    # Delivery Optimization events
    # =========================================================
    $deliveryOptimizationPatterns = @(
        "0x80D0[0-9A-Fa-f]{4}",
        "\berror\b",
        "\bfail",
        "\bfailed\b",
        "\btimeout\b",
        "\bproxy\b",
        "download.*fail",
        "fail.*download"
    )

    $deliveryOptimizationEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-DeliveryOptimization/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    Write-Output "DOEventCount=$($deliveryOptimizationEvents.Count)"

    Add-EventIssue `
        -Source "DO" `
        -Events $deliveryOptimizationEvents `
        -Patterns $deliveryOptimizationPatterns `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    # =========================================================
    # BITS events
    # =========================================================
    $bitsPatterns = @(
        "0x80D0[0-9A-Fa-f]{4}",
        "\berror\b",
        "\bfail",
        "\bfailed\b",
        "\btimeout\b",
        "download.*fail",
        "fail.*download"
    )

    $bitsEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-Bits-Client/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    Write-Output "BITSEventCount=$($bitsEvents.Count)"

    Add-EventIssue `
        -Source "BITS" `
        -Events $bitsEvents `
        -Patterns $bitsPatterns `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    # =========================================================
    # Correlation
    # =========================================================
    if ($deliveryOptimizationFolderStatus -eq "EmptyOrSmall" -and $issues.Count -gt 0) {
        $issues.Add("DOFolderStatus=EmptyOrSmall correlated with recent content engine errors")
    }

    # =========================================================
    # Deduplicate without breaking Generic.List type
    # =========================================================
    $uniqueIssues = @(
        $issues | Select-Object -Unique
    )

    # =========================================================
    # Final result
    # =========================================================
    if ($uniqueIssues.Count -gt 0) {
        Write-Output "Status=DeliveryOptimizationContentEngineIssuesDetected"
        Write-Output "IssueCount=$($uniqueIssues.Count)"
        Write-Output "----------------------------------------"

        $uniqueIssues |
            Select-Object -First $MaxIssuesToDisplay |
            ForEach-Object {
                Write-Output $_
            }

        exit 1
    }

    Write-Output "Status=DeliveryOptimizationContentEngineHealthy"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 2
}
