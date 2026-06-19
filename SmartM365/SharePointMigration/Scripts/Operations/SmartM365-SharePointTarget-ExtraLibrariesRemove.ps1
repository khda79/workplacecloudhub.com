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
