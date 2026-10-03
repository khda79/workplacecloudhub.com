<#
.SYNOPSIS
    Deletes SharePoint Online libraries that contain extra files compared with SP2019.

.DESCRIPTION
    Reads LibrarySummary.csv from a comparison folder and selects only libraries
    where ExtraInTarget is greater than zero. DifferentSize is ignored because
    ShareGate updates changed files during the final copy. The script uses
    interactive PnP authentication only.

    By default, this script is a dry run. Add -Execute to delete the libraries.

.EXAMPLE
    .\SmartM365-SharePointTarget-ExtraLibrariesRemove.ps1 -ComparisonDirectory ".\Migrations\MyMigration\comparisons\files\Migration-files-20260603-214544"

.EXAMPLE
    .\SmartM365-SharePointTarget-ExtraLibrariesRemove.ps1 -ComparisonDirectory ".\Migrations\MyMigration\comparisons\files\Migration-files-20260603-214544" -Execute
.VERSION
    1.0.0
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [ValidateNotNullOrEmpty()]
    [string]$ComparisonDirectory,

    [ValidateNotNullOrEmpty()]
    [string]$LibrarySummaryPath,

    [string]$ClientId = $env:SPO_INVENTORY_CLIENT_ID,

    [string]$Tenant = $env:SPO_INVENTORY_TENANT,

    [switch]$Execute,

    [switch]$ForceAuthentication,

    [switch]$Force,

    [switch]$AllowDuplicateKeys,

    [ValidateNotNullOrEmpty()]
    [string]$ResultPath,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'
Microsoft.PowerShell.Utility\Write-Host ('{0} Script  : {1} v{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $MyInvocation.MyCommand.Name, '1.0.0') -ForegroundColor Cyan
$script:TranscriptStarted = $false

function Get-ConsoleTimestamp {
    return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Write-Host {
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
        [object[]]$Object,

        [ConsoleColor]$ForegroundColor,

        [switch]$NoNewline
    )

    $message = if ($Object) { ($Object -join ' ') } else { '' }
    $line = if ($message -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ') { $message } else { "{0} {1}" -f (Get-ConsoleTimestamp), $message }
    $parameters = @{ Object = $line }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $parameters.ForegroundColor = $ForegroundColor }
    if ($NoNewline) { $parameters.NoNewline = $true }
    Microsoft.PowerShell.Utility\Write-Host @parameters
}

function Write-Warning {
    param(
        [Parameter(Position = 0)]
        [string]$Message
    )

    $line = if ($Message -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ') { $Message } else { "{0} WARNING: {1}" -f (Get-ConsoleTimestamp), $Message }
    Microsoft.PowerShell.Utility\Write-Host $line -ForegroundColor Yellow
}

function Read-Host {
    param(
        [Parameter(Position = 0)]
        [string]$Prompt
    )

    Microsoft.PowerShell.Utility\Read-Host ("{0} {1}" -f (Get-ConsoleTimestamp), $Prompt)
}

function Add-TimestampToLogFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $timestampPattern = '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} '
    $content = Get-Content -LiteralPath $Path
    $updated = foreach ($line in $content) {
        if ($line -match $timestampPattern -or [string]::IsNullOrWhiteSpace($line)) {
            $line
        }
        else {
            "{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line
        }
    }

    Set-Content -LiteralPath $Path -Value $updated -Encoding UTF8 -WhatIf:$false
}

function Stop-TimestampedTranscript {
    param([string]$Path)

    try {
        Stop-Transcript | Out-Null
    }
    catch {
        return
    }

    $script:TranscriptStarted = $false
    Add-TimestampToLogFile -Path $Path
}

function Import-PnPPowerShellModule {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "PnP.PowerShell is not installed. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }

    Import-Module PnP.PowerShell -ErrorAction Stop
}

function ConvertTo-NormalizedPath {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $text = [System.Uri]::UnescapeDataString($Value.Trim()).Replace('\', '/')
    if ($text -match '^https?://') {
        $text = ([System.Uri]$text).AbsolutePath
    }

    if (-not $text.StartsWith('/')) {
        $text = "/$text"
    }

    while ($text.Contains('//')) {
        $text = $text.Replace('//', '/')
    }

    return $text.TrimEnd('/').ToLowerInvariant()
}

function Get-RelativeLibraryPath {
    param(
        [string]$WebPath,
        [string]$LibraryPath
    )

    $cleanWebPath = ConvertTo-NormalizedPath -Value $WebPath
    $cleanLibraryPath = ConvertTo-NormalizedPath -Value $LibraryPath

    if ([string]::IsNullOrWhiteSpace($cleanLibraryPath)) {
        return $null
    }

    if ($cleanWebPath -and $cleanLibraryPath.StartsWith("$cleanWebPath/")) {
        return $cleanLibraryPath.Substring($cleanWebPath.Length + 1)
    }

    return $cleanLibraryPath.TrimStart('/')
}

function Get-TargetLibraryServerRelativePath {
    param(
        [string]$TargetWebUrl,
        [string]$WebPath,
        [string]$LibraryPath
    )

    if ([string]::IsNullOrWhiteSpace($TargetWebUrl)) {
        return $null
    }

    $targetWebPath = ConvertTo-NormalizedPath -Value ([System.Uri]$TargetWebUrl).AbsolutePath
    $relativeLibraryPath = Get-RelativeLibraryPath -WebPath $WebPath -LibraryPath $LibraryPath

    if ([string]::IsNullOrWhiteSpace($relativeLibraryPath) -or $relativeLibraryPath.Contains('|')) {
        return $null
    }

    return (ConvertTo-NormalizedPath -Value ("{0}/{1}" -f $targetWebPath.TrimEnd('/'), $relativeLibraryPath.Trim('/')))
}

function Connect-InteractiveOnly {
    param(
        [string]$Url,
        [switch]$UseForceAuthentication
    )

    $parameters = @{
        Url              = $Url
        Interactive      = $true
        ReturnConnection = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
        $parameters.ClientId = $ClientId
    }

    if (-not [string]::IsNullOrWhiteSpace($Tenant)) {
        $parameters.Tenant = $Tenant
    }

    if ($UseForceAuthentication) {
        $parameters.ForceAuthentication = $true
    }

    Connect-PnPOnline @parameters
}

function Find-TargetList {
    param(
        $Connection,
        [pscustomobject]$Item
    )

    $targetPath = Get-TargetLibraryServerRelativePath `
        -TargetWebUrl $Item.TargetWebUrl `
        -WebPath $Item.WebPath `
        -LibraryPath $Item.LibraryPath

    $lists = @(Get-PnPList -Includes Title,Id,RootFolder,BaseTemplate,Hidden -Connection $Connection |
        Where-Object { $_.BaseTemplate -eq 101 })

    if ($targetPath) {
        $pathMatches = @($lists | Where-Object {
                ConvertTo-NormalizedPath -Value $_.RootFolder.ServerRelativeUrl -eq $targetPath
            })

        if ($pathMatches.Count -eq 1) {
            return [pscustomobject]@{
                Status     = 'MatchedByPath'
                List       = $pathMatches[0]
                TargetPath = $targetPath
                Message    = ''
            }
        }

        if ($pathMatches.Count -gt 1) {
            return [pscustomobject]@{
                Status     = 'AmbiguousPath'
                List       = $null
                TargetPath = $targetPath
                Message    = "Multiple libraries match target path '$targetPath'."
            }
        }
    }

    $titleMatches = @($lists | Where-Object { $_.Title -eq $Item.TargetLibraryTitle })
    if ($titleMatches.Count -eq 1) {
        return [pscustomobject]@{
            Status     = 'MatchedByTitle'
            List       = $titleMatches[0]
            TargetPath = $targetPath
            Message    = ''
        }
    }

    if ($titleMatches.Count -gt 1) {
        return [pscustomobject]@{
            Status     = 'AmbiguousTitle'
            List       = $null
            TargetPath = $targetPath
            Message    = "Multiple libraries match target title '$($Item.TargetLibraryTitle)'."
        }
    }

    return [pscustomobject]@{
        Status     = 'NotFound'
        List       = $null
        TargetPath = $targetPath
        Message    = 'Target library was not found.'
    }
}

if ([string]::IsNullOrWhiteSpace($ComparisonDirectory) -and [string]::IsNullOrWhiteSpace($LibrarySummaryPath)) {
    $scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $localLibrarySummaryPath = Join-Path -Path $scriptDirectory -ChildPath 'LibrarySummary.csv'
    if (Test-Path -LiteralPath $localLibrarySummaryPath) {
        $ComparisonDirectory = $scriptDirectory
    }
    else {
        throw "Provide -ComparisonDirectory or -LibrarySummaryPath."
    }
}

if ([string]::IsNullOrWhiteSpace($LibrarySummaryPath)) {
    $LibrarySummaryPath = Join-Path -Path $ComparisonDirectory -ChildPath 'LibrarySummary.csv'
}

if (-not (Test-Path -LiteralPath $LibrarySummaryPath)) {
    throw "Library summary file not found: $LibrarySummaryPath"
}

if ([string]::IsNullOrWhiteSpace($ResultPath)) {
    $resultDirectory = if ([string]::IsNullOrWhiteSpace($ComparisonDirectory)) {
        Split-Path -Path $LibrarySummaryPath -Parent
    }
    else {
        $ComparisonDirectory
    }

    $ResultPath = Join-Path -Path $resultDirectory -ChildPath ("Libraries-To-Delete-Deletion-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $logBaseDirectory = if ([string]::IsNullOrWhiteSpace($ComparisonDirectory)) {
        Split-Path -Path $LibrarySummaryPath -Parent
    }
    else {
        $ComparisonDirectory
    }

    $LogPath = Join-Path -Path (Join-Path -Path $logBaseDirectory -ChildPath 'logs') -ChildPath ("Libraries-To-Delete-Deletion-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
}

$logDirectory = Split-Path -Path $LogPath -Parent
if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

trap {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
    throw $_
}

Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
$script:TranscriptStarted = $true
Write-Host ("Run log: {0}" -f $LogPath) -ForegroundColor Cyan

Import-PnPPowerShellModule

function Get-CsvDelimiter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 1
    if (([regex]::Matches($firstLine, ';')).Count -gt ([regex]::Matches($firstLine, ',')).Count) {
        return ';'
    }

    return ','
}

function Get-ComparisonDuplicateKeyCount {
    param(
        [string]$SummaryPath
    )

    if ([string]::IsNullOrWhiteSpace($SummaryPath) -or -not (Test-Path -LiteralPath $SummaryPath)) {
        return 0
    }

    $summary = Import-Csv -LiteralPath $SummaryPath -Delimiter (Get-CsvDelimiter -Path $SummaryPath) |
        Select-Object -First 1

    if (-not $summary) {
        return 0
    }

    return ([int]$summary.SourceDuplicateKeysIgnored + [int]$summary.TargetDuplicateKeysIgnored)
}

$summaryPath = Join-Path -Path (Split-Path -Path $LibrarySummaryPath -Parent) -ChildPath 'Summary.csv'
$duplicateKeyCount = Get-ComparisonDuplicateKeyCount -SummaryPath $summaryPath
if ($duplicateKeyCount -gt 0 -and -not $AllowDuplicateKeys) {
    $duplicatePath = Join-Path -Path (Split-Path -Path $LibrarySummaryPath -Parent) -ChildPath 'DuplicateKeys.csv'
    throw "Deletion blocked because duplicate inventory keys were found. Duplicate keys ignored: $duplicateKeyCount. Review $duplicatePath first, or rerun with -AllowDuplicateKeys after validation."
}

$items = @(Import-Csv -LiteralPath $LibrarySummaryPath -Delimiter (Get-CsvDelimiter -Path $LibrarySummaryPath) |
    Where-Object { [int]$_.ExtraInTarget -gt 0 } |
    Sort-Object TargetWebUrl, TargetLibraryTitle)

if ($items.Count -eq 0) {
    Write-Host "No libraries with ExtraInTarget > 0 were found."
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
    return
}

Write-Host ("Libraries selected: {0}" -f $items.Count)
Write-Host ("Mode: {0}" -f ($(if ($Execute) { 'EXECUTE - libraries will be deleted' } else { 'DRY RUN - no deletion' })))
Write-Host ("Result CSV: {0}" -f $ResultPath)

if ($Execute -and -not $Force) {
    Write-Warning "This will delete SharePoint Online document libraries. Deleted libraries should be restored only through SharePoint recycle bin or by rerunning migration."
    $confirmation = Read-Host "Type DELETE to continue"
    if ($confirmation -ne 'DELETE') {
        throw "Deletion cancelled."
    }
}

$results = New-Object System.Collections.Generic.List[object]
$forceAuthenticationPending = [bool]$ForceAuthentication

foreach ($group in ($items | Group-Object TargetWebUrl)) {
    $webUrl = $group.Name
    Write-Host ("Connecting interactively to: {0}" -f $webUrl)
    $connection = Connect-InteractiveOnly -Url $webUrl -UseForceAuthentication:$forceAuthenticationPending
    $forceAuthenticationPending = $false

    foreach ($item in $group.Group) {
        $match = Find-TargetList -Connection $connection -Item $item
        $list = $match.List
        $actionStatus = if ($Execute) { 'PendingDelete' } else { 'WouldDelete' }
        $message = $match.Message

        if ($null -eq $list) {
            $actionStatus = $match.Status
        }
        elseif ($Execute) {
            $targetName = "{0} | {1}" -f $webUrl, $list.Title
            if ($PSCmdlet.ShouldProcess($targetName, 'Delete SharePoint Online document library')) {
                try {
                    Remove-PnPList -Identity $list.Id -Force -Connection $connection
                    $actionStatus = 'Deleted'
                    $message = 'Library deleted.'
                }
                catch {
                    $actionStatus = 'DeleteFailed'
                    $message = $_.Exception.Message
                    Write-Warning ("Failed to delete library '{0}' in web '{1}': {2}" -f $list.Title, $webUrl, $_.Exception.Message)
                }
            }
            else {
                $actionStatus = 'SkippedByShouldProcess'
            }
        }
        else {
            Write-Host ("Would delete: {0} | {1} | ExtraInTarget={2}" -f $webUrl, $item.TargetLibraryTitle, $item.ExtraInTarget)
        }

        $results.Add([pscustomobject]@{
                Time                 = Get-Date
                ActionStatus         = $actionStatus
                MatchStatus          = $match.Status
                TargetWebUrl         = $webUrl
                TargetLibraryTitle   = $item.TargetLibraryTitle
                MatchedListTitle     = if ($list) { $list.Title } else { $null }
                MatchedListId        = if ($list) { $list.Id } else { $null }
                MatchedListUrl       = if ($list) { $list.RootFolder.ServerRelativeUrl } else { $null }
                ExpectedLibraryPath  = $match.TargetPath
                ExtraInTarget        = $item.ExtraInTarget
                ExtraInTargetBytes   = $item.ExtraInTargetBytes
                TargetFiles          = $item.TargetFiles
                TargetBytes          = $item.TargetBytes
                MissingInTarget      = $item.MissingInTarget
                SourceWebUrl         = $item.SourceWebUrl
                SourceLibraryTitle   = $item.SourceLibraryTitle
                SourceFiles          = $item.SourceFiles
                SourceBytes          = $item.SourceBytes
                LibraryPath          = $item.LibraryPath
                Message              = $message
            })
    }
}

$results | Export-Csv -Delimiter ';' -Path $ResultPath -NoTypeInformation -Encoding UTF8
Write-Host ("Deletion result written: {0}" -f $ResultPath)

if ($script:TranscriptStarted) {
    Stop-TimestampedTranscript -Path $LogPath
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDv2XX9iWQzu28w
# Sx7MnIjWa9anhddDJvOKlAaqRU+uJqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIMe/Bgy6fh2FUfF3DPDgCTJO/dF05xrx/jTvDtyKfwljMA0GCSqG
# SIb3DQEBAQUABIIBgKgZ3Cu+KDbNud+bYk2tSd4qCM/SiyuqaTPCj/Ublae5+Nfa
# f8hewEIhhmZCGmtwcl0b96R6wNES0wUYNrgeA5WCpxkeaM4leCotWBpsB2UbOZ7u
# ngA7h6DZmPHzoyndphSZGi4JFC4hB6YxM24Hd8YKygTSeUZ3125jfo5BybYvlpij
# F230umoI9ku+BcEky++SViwuDSZU6w4g/CbLm2S+p/I5Tlb3JJFqC0sJ5mf9EA1G
# zdUtSiShs8YhRX23T8NFfEyDMhDrSKdD2OHw02SWtzjjfI+KmNSA5PCJvFdMcW5D
# +fwJR9NeuTik708g/Dam54ItbKOTLCDyTGJLUb7QQm5yWjoemNKXixWY7ugn6KdR
# YQ9DPfdKQmcgN5yjPdTKgRAswGmkyBzXEw27+a6I3dAUM1asPpMhSLPvUMSVPku9
# xmrCwhXIUd5wt7J4j1V6DkBJcCoRopy4ydspaZ/wgudxeypWA8MVADYmIStTsZX6
# Ezh+FI8o+3o6v8CSt6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMwMDE5
# MDlaMC8GCSqGSIb3DQEJBDEiBCA10hXNh4h6S7u9kXwZRULEijh8LYUBff5d+h9T
# v3K7NjANBgkqhkiG9w0BAQEFAASCAgBkM4lQ6h+KTLovGTP8EXCoHZeIsohLwPY2
# VhnNZz7vAz7ltr+l/G3yj6+itOuT0q3cMIAlex1gH2UMaCbWzx8UGGAvU/+1+NAW
# UYOu2xnxWqFUHbr4gg/6hC5vazC6gQlsWRnEV0Dwxaf8bWyv2QgVM7Ns8ZuZ32FV
# JoHqqlzYHH6Ihk61HIjm9gM0JBxam2Rohshjh/tTdaewJHvhHS84bacwmlnvxstS
# s8ESKiIYgGwpwAHqi9Lse4/YVXGajzzzEE6aXIKZ0ZjUnFkqy7QNdNOn8Kw8rMOX
# qcvpFFt4EQ/U8l5Uw0RBB6da6gC1mLIVoQbjhT1NpRvmrCGz/zl91djWQlQw5a2E
# wq8PN1C8jDY85epKx0xcyYRYj+OkLs0RY7pdZ4+va/KC65dhhvP/7DZU8H80w9nT
# eofDetOY2iHGbu0/+CDYHFnkylRMdXx4/jcd745HAWDhgKZKeyPHiC9nQv/Utksf
# rrwDwHsnYpaTfVx6Dd0VTOxP6R+Qd+vkErR84Y70deqw5KF02NMm0tXQDci6Q//2
# 7rskeSfumncTkqXMFJ6bXM2M7MAGUz+VUrWa1u5Y72zlB5xjQwgRUschW90niQdi
# d1M7V9kFlOhpRe/AWrj1Qx5jlwSQqnVvT/eX+UHtsgedFpTKXxMpEMNUhA6d7ow9
# jfTlb/4RaA==
# SIG # End signature block
