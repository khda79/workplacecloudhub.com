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

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC9/tVdUWhlp/M4
# ovZv1GummKtyGwiaL45FXn9u2cyNAqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBMNLPCJ0D6I5KJEzMv
# dFxlZ4y5FSp9i1r6wxZMRe1K/zANBgkqhkiG9w0BAQEFAASCAYB6ybJw9QC5FEKa
# 9lWWEvZw65wkNcXiW9diXq2UM82+DQ/eMDoGo8mlxPundo7wWLgjAqNeV5oq9nsV
# JfV5MZ1KgQhwJq3qVvHwXGo6PsRixgtF0lBtTlTaIyE+Icj/1SmAmrEuQZFZgWZC
# iAJj1QARPU5QC3PFSz8F8Ry+/v9sl+7/LczdCPAL1F+BHKKHOD3/WseiqooDX+oS
# TMkmvV4BtKE69+H9UU64JE8qG1azK5NmQk0PmmqjNO5Gy6OwNV2yXN0DJ2T0VctZ
# NH9oUbsotqGQMz4cUsUwfR3aAUcES+8O7q7RxicFfzDWVmje5zTpdIMr0hJG6OjX
# Z7BGOymzPdP8VKbfPZ24ydgjKludQPqOjuFl36lw8ld6EllLeDRJNO9kOkXLC3sP
# pteeJQFx2l2ZtkNBwgb1m3tjn8jADzdFIpWPP7MBnQNwJ0kHibnMILClWX/JhX8t
# ZVllbMH7QcF9gYZpMysgNBvbLET1EMtn9rNhHxJet+h8F+hWutQ=
# SIG # End signature block
