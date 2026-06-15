import argparse
import builtins
import csv
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Export"))

from export_libraries_to_delete_excel import build_delete_rows


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if args:
        args = (f"{timestamp} {args[0]}", *args[1:])
    else:
        args = (timestamp,)
    builtins.print(*args, **kwargs)


SCRIPT_HEADER = r'''<#
.SYNOPSIS
    Deletes the SharePoint Online libraries listed in this script.

.DESCRIPTION
    This script is generated for one comparison result. The target libraries are
    embedded directly in the $LibrariesToDelete array.

    By default, this script is a dry run. Add -Execute to delete the libraries.
    Authentication is interactive only.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$ClientId = $env:SPO_INVENTORY_CLIENT_ID,

    [string]$Tenant = $env:SPO_INVENTORY_TENANT,

    [switch]$Execute,

    [switch]$ForceAuthentication,

    [switch]$Force,

    [ValidateNotNullOrEmpty()]
    [string]$ResultPath,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'
$script:TranscriptStarted = $false
$script:DeletionResults = $null
$script:DeletionResultPath = $null

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

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return
    }

    $timestampPrefixPattern = '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} '
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $temporaryPath = "{0}.timestamp.tmp" -f $Path

    Get-Content -LiteralPath $Path | ForEach-Object {
        if ($_ -match $timestampPrefixPattern) {
            $_
        }
        else {
            "{0} {1}" -f $timestamp, $_
        }
    } | Set-Content -LiteralPath $temporaryPath -Encoding UTF8 -WhatIf:$false

    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force -WhatIf:$false
}

function Stop-TimestampedTranscript {
    param([string]$Path)

    Stop-Transcript -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $false
    Add-TimestampToLogFile -Path $Path
}

function Export-DeletionResults {
    param([switch]$Quiet)

    if ($null -ne $script:DeletionResults -and -not [string]::IsNullOrWhiteSpace($script:DeletionResultPath)) {
        $script:DeletionResults | Export-Csv -Delimiter ';' -Path $script:DeletionResultPath -NoTypeInformation -Encoding UTF8
        if (-not $Quiet) {
            Write-Host ("Deletion result written: {0}" -f $script:DeletionResultPath) -ForegroundColor DarkCyan
        }
    }
}

trap {
    Export-DeletionResults

    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }

    throw
}

function Import-PnPPowerShellModule {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "PnP.PowerShell is not installed. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }

    Import-Module PnP.PowerShell -ErrorAction Stop
}

function ConvertTo-NormalizedPath {
    param([string]$Value)

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

function Connect-InteractiveOnly {
    param(
        [string]$Url,
        [switch]$UseForceAuthentication
    )

    if ($UseForceAuthentication) {
        try {
            Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction SilentlyContinue
            Write-Host "Cleared persisted PnP login before forced authentication." -ForegroundColor DarkCyan
        }
        catch {
            Write-Warning ("Could not clear persisted PnP login before forced authentication: {0}" -f $_.Exception.Message)
        }
    }

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

function Get-PnPConnectionIdentity {
    param(
        $Connection,
        [string]$Url
    )

    try {
        $context = Get-PnPContext -Connection $Connection
        $context.Load($context.Web.CurrentUser)
        $context.ExecuteQuery()

        $currentUser = $context.Web.CurrentUser
        $loginName = if ($currentUser.LoginName) { $currentUser.LoginName } else { '<unknown>' }
        $email = if ($currentUser.Email) { $currentUser.Email } else { '<no email>' }
        $title = if ($currentUser.Title) { $currentUser.Title } else { '<no title>' }

        return [pscustomobject]@{
            LoginName   = $loginName
            Email       = $email
            Title       = $title
            DisplayText = ("{0} | {1} | {2}" -f $loginName, $email, $title)
        }
    }
    catch {
        Write-Warning ("Could not determine connected account for '{0}': {1}" -f $Url, $_.Exception.Message)
        return [pscustomobject]@{
            LoginName   = '<unknown>'
            Email       = '<unknown>'
            Title       = '<unknown>'
            DisplayText = '<unknown>'
        }
    }
}

function Write-DeleteStatus {
    param(
        [string]$Status,
        [string]$WebUrl,
        [string]$TargetLibraryTitle,
        [string]$Message
    )

    switch ($Status) {
        'Deleted' {
            Write-Host ("OK Deleted library: {0} | {1}" -f $WebUrl, $TargetLibraryTitle) -ForegroundColor Green
        }
        'WouldDelete' {
            Write-Host ("DRY RUN Would delete library: {0} | {1}" -f $WebUrl, $TargetLibraryTitle) -ForegroundColor DarkCyan
        }
        'SkippedByShouldProcess' {
            Write-Host ("SKIP ShouldProcess: {0} | {1}" -f $WebUrl, $TargetLibraryTitle) -ForegroundColor Yellow
        }
        default {
            Write-Warning ("FAILED {0} | {1}: {2}" -f $WebUrl, $TargetLibraryTitle, $Message)
        }
    }
}

function Find-TargetList {
    param(
        $Connection,
        [pscustomobject]$Item
    )

    $lists = @(Get-PnPList -Includes Title,Id,RootFolder,BaseTemplate,Hidden -Connection $Connection |
        Where-Object { $_.BaseTemplate -eq 101 })

    if (-not [string]::IsNullOrWhiteSpace($Item.TargetLibraryServerRelativeUrl)) {
        $pathMatches = @($lists | Where-Object {
                ConvertTo-NormalizedPath -Value $_.RootFolder.ServerRelativeUrl -eq $Item.TargetLibraryServerRelativeUrl
            })

        if ($pathMatches.Count -eq 1) {
            return [pscustomobject]@{
                Status  = 'MatchedByPath'
                List    = $pathMatches[0]
                Message = ''
            }
        }

        if ($pathMatches.Count -gt 1) {
            return [pscustomobject]@{
                Status  = 'AmbiguousPath'
                List    = $null
                Message = "Multiple libraries match target path '$($Item.TargetLibraryServerRelativeUrl)'."
            }
        }
    }

    $titleMatches = @($lists | Where-Object { $_.Title -eq $Item.TargetLibraryTitle })
    if ($titleMatches.Count -eq 1) {
        return [pscustomobject]@{
            Status  = 'MatchedByTitle'
            List    = $titleMatches[0]
            Message = ''
        }
    }

    if ($titleMatches.Count -gt 1) {
        return [pscustomobject]@{
            Status  = 'AmbiguousTitle'
            List    = $null
            Message = "Multiple libraries match target title '$($Item.TargetLibraryTitle)'."
        }
    }

    return [pscustomobject]@{
        Status  = 'NotFound'
        List    = $null
        Message = 'Target library was not found.'
    }
}

'''


SCRIPT_FOOTER = r'''

if ([string]::IsNullOrWhiteSpace($ResultPath)) {
    $ResultPath = Join-Path -Path $PSScriptRoot -ChildPath ("Libraries-To-Delete-Deletion-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
}
$script:DeletionResultPath = $ResultPath

if ($Execute) {
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'logs') -ChildPath ("Libraries-To-Delete-Deletion-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
    }

    $logDirectory = Split-Path -Path $LogPath -Parent
    if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
}

Import-PnPPowerShellModule

if (-not $LibrariesToDelete -or $LibrariesToDelete.Count -eq 0) {
    Write-Host "No libraries are embedded in this script."
    return
}

Write-Host ("Libraries embedded: {0}" -f $LibrariesToDelete.Count)
$modeText = $(if ($Execute) { 'EXECUTE - libraries will be deleted' } else { 'DRY RUN - no deletion' })
$modeColor = $(if ($Execute) { 'Red' } else { 'DarkCyan' })
Write-Host ("Mode: {0}" -f $modeText) -ForegroundColor $modeColor
Write-Host ("Result CSV: {0}" -f $ResultPath) -ForegroundColor DarkCyan
if ($Execute) {
    Write-Host ("Run log: {0}" -f $LogPath) -ForegroundColor DarkCyan
}

if ($Execute -and -not $Force) {
    Write-Warning "This will delete SharePoint Online document libraries. Deleted libraries should be restored only through SharePoint recycle bin or by rerunning migration."
    $confirmation = Read-Host "Type DELETE to continue"
    if ($confirmation -ne 'DELETE') {
        throw "Deletion cancelled."
    }
}

$script:DeletionResults = New-Object System.Collections.Generic.List[object]
$results = $script:DeletionResults
$forceAuthenticationPending = [bool]$ForceAuthentication

foreach ($group in ($LibrariesToDelete | Group-Object TargetWebUrl)) {
    $webUrl = $group.Name
    Write-Host ("Connecting interactively to: {0}" -f $webUrl)
    $connection = Connect-InteractiveOnly -Url $webUrl -UseForceAuthentication:$forceAuthenticationPending
    $forceAuthenticationPending = $false
    $connectedAccount = Get-PnPConnectionIdentity -Connection $connection -Url $webUrl
    Write-Host ("Connected account for {0}: {1}" -f $webUrl, $connectedAccount.DisplayText) -ForegroundColor Cyan

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
                    Remove-PnPList -Identity $list.Id -Force -Connection $connection -ErrorAction Stop
                    $actionStatus = 'Deleted'
                    $message = 'Library deleted.'
                }
                catch {
                    $actionStatus = 'DeleteFailed'
                    $message = $_.Exception.Message
                }
            }
            else {
                $actionStatus = 'SkippedByShouldProcess'
            }
        }
        else {
        }

        Write-DeleteStatus -Status $actionStatus -WebUrl $webUrl -TargetLibraryTitle $item.TargetLibraryTitle -Message $message

        $results.Add([pscustomobject]@{
                Time                           = Get-Date
                ActionStatus                   = $actionStatus
                MatchStatus                    = $match.Status
                TargetWebUrl                   = $webUrl
                TargetLibraryTitle             = $item.TargetLibraryTitle
                TargetLibraryServerRelativeUrl = $item.TargetLibraryServerRelativeUrl
                ConnectedAccount               = $connectedAccount.DisplayText
                ConnectedAccountLogin          = $connectedAccount.LoginName
                ConnectedAccountEmail          = $connectedAccount.Email
                MatchedListTitle               = if ($list) { $list.Title } else { $null }
                MatchedListId                  = if ($list) { $list.Id } else { $null }
                MatchedListUrl                 = if ($list) { $list.RootFolder.ServerRelativeUrl } else { $null }
                ExtraInTarget                  = $item.ExtraInTarget
                ExtraInTargetPercent           = $item.ExtraInTargetPercent
                ExtraInTargetBytes             = $item.ExtraInTargetBytes
                ExtraInTargetGB                = $item.ExtraInTargetGB
                TargetFiles                    = $item.TargetFiles
                TargetSizeBytes                = $item.TargetSizeBytes
                TargetSizeGB                   = $item.TargetSizeGB
                MissingInTarget                = $item.MissingInTarget
                SourceWebUrl                   = $item.SourceWebUrl
                SourceLibraryTitle             = $item.SourceLibraryTitle
                SourceSizeBytes                = $item.SourceSizeBytes
                SourceSizeGB                   = $item.SourceSizeGB
                Message                        = $message
            })
    }
}

Export-DeletionResults

if ($script:TranscriptStarted) {
    Stop-TimestampedTranscript -Path $LogPath
}
'''


def ps_string(value):
    if value is None:
        return ""
    return str(value).replace("'", "''")


def as_int(value):
    if value is None or str(value).strip() == "":
        return 0
    return int(str(value).strip())


def detect_csv_dialect(handle):
    sample = handle.read(65536)
    handle.seek(0)
    if not sample:
        return csv.excel

    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        return csv.excel


def duplicate_key_count(comparison_directory):
    summary_path = comparison_directory / "Summary.csv"
    if not summary_path.exists():
        return 0

    with summary_path.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle, dialect=detect_csv_dialect(handle)))

    if not rows:
        return 0

    summary = rows[0]
    return as_int(summary.get("SourceDuplicateKeysIgnored")) + as_int(summary.get("TargetDuplicateKeysIgnored"))


def build_target_server_relative_url(row):
    url = row.get("TargetLibraryUrl", "")
    if not url:
        return ""
    try:
        from urllib.parse import urlparse, unquote

        return unquote(urlparse(url).path).rstrip("/").lower()
    except Exception:
        return ""


def render_library_array(rows):
    lines = ["$LibrariesToDelete = @("]
    for row in rows:
        lines.append("    [pscustomobject]@{")
        lines.append(f"        TargetWebUrl                   = '{ps_string(row.get('TargetWebUrl'))}'")
        lines.append(f"        TargetLibraryTitle             = '{ps_string(row.get('TargetLibraryTitle'))}'")
        lines.append(f"        TargetLibraryUrl               = '{ps_string(row.get('TargetLibraryUrl'))}'")
        lines.append(f"        TargetLibraryServerRelativeUrl = '{ps_string(build_target_server_relative_url(row))}'")
        lines.append(f"        ExtraInTarget                  = {int(row.get('ExtraInTarget') or 0)}")
        lines.append(f"        ExtraInTargetPercent           = {float(row.get('ExtraInTargetPercent') or 0)}")
        lines.append(f"        ExtraInTargetBytes             = {int(row.get('ExtraInTargetBytes') or 0)}")
        lines.append(f"        ExtraInTargetGB                = {float(row.get('ExtraInTargetGB') or 0)}")
        lines.append(f"        TargetFiles                    = {int(row.get('TargetFiles') or 0)}")
        lines.append(f"        TargetSizeBytes                = {int(row.get('TargetSizeBytes') or 0)}")
        lines.append(f"        TargetSizeGB                   = {float(row.get('TargetSizeGB') or 0)}")
        lines.append(f"        MissingInTarget                = {int(row.get('MissingInTarget') or 0)}")
        lines.append(f"        MatchedFiles                   = {int(row.get('MatchedFiles') or 0)}")
        lines.append(f"        SourceWebUrl                   = '{ps_string(row.get('SourceWebUrl'))}'")
        lines.append(f"        SourceLibraryTitle             = '{ps_string(row.get('SourceLibraryTitle'))}'")
        lines.append(f"        SourceFiles                    = {int(row.get('SourceFiles') or 0)}")
        lines.append(f"        SourceSizeBytes                = {int(row.get('SourceSizeBytes') or 0)}")
        lines.append(f"        SourceSizeGB                   = {float(row.get('SourceSizeGB') or 0)}")
        lines.append(f"        LibraryPath                    = '{ps_string(row.get('LibraryPath'))}'")
        lines.append(f"        Status                         = '{ps_string(row.get('Status'))}'")
        lines.append("    }")
    lines.append(")")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--comparison-directory", required=True)
    parser.add_argument("--output-ps1")
    parser.add_argument("--allow-duplicate-keys", action="store_true")
    args = parser.parse_args()

    comparison_directory = Path(args.comparison_directory)
    library_summary_path = comparison_directory / "LibrarySummary.csv"
    if not library_summary_path.exists():
        raise FileNotFoundError(f"LibrarySummary.csv not found: {library_summary_path}")

    output_ps1 = Path(args.output_ps1) if args.output_ps1 else comparison_directory / "SmartM365-SharePointTarget-ExtraLibrariesRemove-Custom.ps1"
    duplicate_count = duplicate_key_count(comparison_directory)
    if duplicate_count > 0 and not args.allow_duplicate_keys:
        duplicate_path = comparison_directory / "DuplicateKeys.csv"
        duplicate_excel_path = comparison_directory / "DuplicateKeys.xlsx"
        print("")
        print("Deletion script generation skipped.")
        print("Reason: duplicate inventory keys were found.")
        print(f"Duplicate keys ignored: {duplicate_count}")
        print(f"Review CSV: {duplicate_path}")
        if duplicate_excel_path.exists():
            print(f"Review Excel: {duplicate_excel_path}")
        print("After validation, rerun with --allow-duplicate-keys to generate the deletion script.")
        return 2

    rows = build_delete_rows(library_summary_path)
    if not rows:
        if output_ps1.exists():
            output_ps1.unlink()
        print("Deletion script generation skipped.")
        print("Reason: no libraries with extra files were found in LibrarySummary.csv.")
        print(f"PowerShell file not created: {output_ps1}")
        return 0

    content = SCRIPT_HEADER + render_library_array(rows) + SCRIPT_FOOTER
    output_ps1.write_text(content, encoding="utf-8")

    print(f"PowerShell file created: {output_ps1}")
    print(f"Libraries embedded: {len(rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
