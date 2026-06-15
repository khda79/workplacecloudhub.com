import argparse
import builtins
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Export"))

from export_folders_to_delete_excel import build_delete_folder_rows


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if args:
        args = (f"{timestamp} {args[0]}", *args[1:])
    else:
        args = (timestamp,)
    builtins.print(*args, **kwargs)


SCRIPT_HEADER = r'''<#
.SYNOPSIS
    Deletes SharePoint Online folders that no longer exist in the SP2019 source file paths.

.DESCRIPTION
    This script is generated for one comparison result. The target folders are
    embedded directly in the $FoldersToDelete array.

    By default, this script is a dry run. Add -Execute to delete folders.
    It always checks that each folder is empty before deletion.
    Run this after deleting extra files.
    Authentication is interactive only.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$ClientId = $env:SPO_INVENTORY_CLIENT_ID,

    [string]$Tenant = $env:SPO_INVENTORY_TENANT,

    [switch]$Execute,

    [switch]$Recycle,

    [switch]$ForceAuthentication,

    [switch]$Force,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$Test,

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
    param([Parameter(Position = 0)][string]$Message)

    $line = if ($Message -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ') { $Message } else { "{0} WARNING: {1}" -f (Get-ConsoleTimestamp), $Message }
    Microsoft.PowerShell.Utility\Write-Host $line -ForegroundColor Yellow
}

function Read-Host {
    param([Parameter(Position = 0)][string]$Prompt)

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
            Write-Host ("Folder deletion result written: {0}" -f $script:DeletionResultPath) -ForegroundColor DarkCyan
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

function Remove-TargetFolderIfEmpty {
    param(
        [Parameter(Mandatory = $true)]
        $Connection,

        [Parameter(Mandatory = $true)]
        [string]$ServerRelativeUrl,

        [switch]$Recycle
    )

    try {
        $context = Get-PnPContext -Connection $Connection
        $folder = $context.Web.GetFolderByServerRelativeUrl($ServerRelativeUrl)
        $context.Load($folder)
        $context.Load($folder.Files)
        $context.Load($folder.Folders)
        $context.ExecuteQuery()

        if ($folder.Files.Count -gt 0 -or $folder.Folders.Count -gt 0) {
            return [pscustomobject]@{
                Success = $false
                Status  = 'SkippedNotEmpty'
                Method  = 'CSOM folder check'
                Message = ("Folder is not empty. Files: {0}; Folders: {1}" -f $folder.Files.Count, $folder.Folders.Count)
            }
        }

        if ($Recycle) {
            $context.Load($folder.ListItemAllFields)
            $context.ExecuteQuery()
            [void]$folder.ListItemAllFields.Recycle()
            $method = 'Folder.ListItemAllFields.Recycle'
            $message = 'Folder moved to recycle bin.'
        }
        else {
            $folder.DeleteObject()
            $method = 'Folder.DeleteObject'
            $message = 'Folder permanently deleted.'
        }

        $context.ExecuteQuery()
        return [pscustomobject]@{
            Success = $true
            Status  = if ($Recycle) { 'Recycled' } else { 'PermanentlyDeleted' }
            Method  = $method
            Message = $message
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Status  = 'DeleteFailed'
            Method  = 'CSOM folder delete'
            Message = $_.Exception.Message
        }
    }
}

function Write-FolderDeleteStatus {
    param(
        [string]$Status,
        [string]$Method,
        [string]$WebUrl,
        [string]$ServerRelativeUrl,
        [string]$Message
    )

    switch ($Status) {
        'Recycled' {
            Write-Host ("OK Recycled [{0}]: {1} | {2}" -f $Method, $WebUrl, $ServerRelativeUrl) -ForegroundColor Green
        }
        'PermanentlyDeleted' {
            Write-Host ("OK Deleted [{0}]: {1} | {2}" -f $Method, $WebUrl, $ServerRelativeUrl) -ForegroundColor Green
        }
        'WouldDelete' {
            Write-Host ("DRY RUN Would delete empty folder if still empty: {0} | {1}" -f $WebUrl, $ServerRelativeUrl) -ForegroundColor DarkCyan
        }
        'SkippedByShouldProcess' {
            Write-Host ("SKIP ShouldProcess: {0} | {1}" -f $WebUrl, $ServerRelativeUrl) -ForegroundColor Yellow
        }
        'SkippedNotEmpty' {
            Write-Host ("SKIP Not empty [{0}]: {1} | {2}: {3}" -f $Method, $WebUrl, $ServerRelativeUrl, $Message) -ForegroundColor Yellow
        }
        default {
            Write-Warning ("FAILED [{0}] {1} | {2}: {3}" -f $Method, $WebUrl, $ServerRelativeUrl, $Message)
        }
    }
}

'''


SCRIPT_FOOTER = r'''

if ([string]::IsNullOrWhiteSpace($ResultPath)) {
    $ResultPath = Join-Path -Path $PSScriptRoot -ChildPath ("Folders-To-Delete-Deletion-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
}
$script:DeletionResultPath = $ResultPath

if ($Execute) {
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'logs') -ChildPath ("Folders-To-Delete-Deletion-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
    }

    $logDirectory = Split-Path -Path $LogPath -Parent
    if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
}

Import-PnPPowerShellModule

if (-not $FoldersToDelete -or $FoldersToDelete.Count -eq 0) {
    Write-Host "No folders are embedded in this script."
    return
}

if ($PSBoundParameters.ContainsKey('Test')) {
    $FoldersToDelete = @($FoldersToDelete | Select-Object -First $Test)
    Write-Host ("Test mode: only the first {0} embedded folders will be processed." -f $FoldersToDelete.Count) -ForegroundColor Yellow
}

Write-Host ("Folders embedded: {0}" -f $FoldersToDelete.Count)
$modeText = $(if ($Execute) { if ($Recycle) { 'EXECUTE - empty folders will be recycled' } else { 'EXECUTE - empty folders will be permanently deleted' } } else { 'DRY RUN - no deletion' })
$modeColor = $(if ($Execute) { if ($Recycle) { 'Yellow' } else { 'Red' } } else { 'DarkCyan' })
Write-Host ("Mode: {0}" -f $modeText) -ForegroundColor $modeColor
Write-Host ("Result CSV: {0}" -f $ResultPath) -ForegroundColor DarkCyan
Write-Host ("PnP ClientId parameter: {0}" -f ($(if ([string]::IsNullOrWhiteSpace($ClientId)) { '<not set>' } else { $ClientId }))) -ForegroundColor DarkCyan
Write-Host ("PnP Tenant parameter: {0}" -f ($(if ([string]::IsNullOrWhiteSpace($Tenant)) { '<not set>' } else { $Tenant }))) -ForegroundColor DarkCyan
if ($Execute) {
    Write-Host ("Run log: {0}" -f $LogPath) -ForegroundColor DarkCyan
}

if ($Execute -and -not $Force) {
    if ($Recycle) {
        Write-Warning "This will move empty SharePoint Online folders to the recycle bin."
        $expectedConfirmation = 'RECYCLE'
    }
    else {
        Write-Warning "This will permanently delete empty SharePoint Online folders. Recovery may not be possible."
        $expectedConfirmation = 'DELETE'
    }

    $confirmation = Read-Host ("Type {0} to continue" -f $expectedConfirmation)
    if ($confirmation -ne $expectedConfirmation) {
        throw "Folder deletion cancelled."
    }
}

$script:DeletionResults = New-Object System.Collections.Generic.List[object]
$results = $script:DeletionResults
$forceAuthenticationPending = [bool]$ForceAuthentication

foreach ($group in ($FoldersToDelete | Sort-Object TargetWebUrl, @{ Expression = 'Depth'; Descending = $true }, TargetFolderServerRelativeUrl | Group-Object TargetWebUrl)) {
    $webUrl = $group.Name
    Write-Host ("Connecting interactively to: {0}" -f $webUrl)
    $connection = Connect-InteractiveOnly -Url $webUrl -UseForceAuthentication:$forceAuthenticationPending
    $forceAuthenticationPending = $false
    $connectedAccount = Get-PnPConnectionIdentity -Connection $connection -Url $webUrl
    Write-Host ("Connected account for {0}: {1}" -f $webUrl, $connectedAccount.DisplayText) -ForegroundColor Cyan

    foreach ($item in ($group.Group | Sort-Object @{ Expression = 'Depth'; Descending = $true }, TargetFolderServerRelativeUrl)) {
        $actionStatus = if ($Execute) { if ($Recycle) { 'PendingRecycle' } else { 'PendingPermanentDelete' } } else { 'WouldDelete' }
        $deleteMethod = if ($Execute) { 'Pending' } else { 'DryRun' }
        $message = ''

        if ([string]::IsNullOrWhiteSpace($item.TargetFolderServerRelativeUrl)) {
            $actionStatus = 'MissingServerRelativeUrl'
            $message = 'Target folder server relative URL is empty.'
        }
        elseif ($Execute) {
            $targetName = "{0} | {1}" -f $webUrl, $item.TargetFolderServerRelativeUrl
            $operation = if ($Recycle) { 'Recycle empty SharePoint Online folder' } else { 'Permanently delete empty SharePoint Online folder' }
            if ($PSCmdlet.ShouldProcess($targetName, $operation)) {
                $deleteResult = Remove-TargetFolderIfEmpty -Connection $connection -ServerRelativeUrl $item.TargetFolderServerRelativeUrl -Recycle:$Recycle
                $actionStatus = $deleteResult.Status
                $deleteMethod = $deleteResult.Method
                $message = $deleteResult.Message
            }
            else {
                $actionStatus = 'SkippedByShouldProcess'
                $deleteMethod = 'SkippedByShouldProcess'
            }
        }

        Write-FolderDeleteStatus -Status $actionStatus -Method $deleteMethod -WebUrl $webUrl -ServerRelativeUrl $item.TargetFolderServerRelativeUrl -Message $message

        $results.Add([pscustomobject]@{
                Time                          = Get-Date
                ActionStatus                  = $actionStatus
                TargetWebUrl                  = $webUrl
                TargetLibraryTitle            = $item.TargetLibraryTitle
                TargetFolderServerRelativeUrl = $item.TargetFolderServerRelativeUrl
                FolderName                    = $item.FolderName
                Depth                         = $item.Depth
                ExtraFileCount                = $item.ExtraFileCount
                ExtraFileBytes                = $item.ExtraFileBytes
                ConnectedAccount              = $connectedAccount.DisplayText
                ConnectedAccountLogin         = $connectedAccount.LoginName
                ConnectedAccountEmail         = $connectedAccount.Email
                DeleteMethod                  = $deleteMethod
                Message                       = $message
            })
        Export-DeletionResults -Quiet
    }
}

Export-DeletionResults

if ($script:TranscriptStarted) {
    Stop-TimestampedTranscript -Path $LogPath
}
'''


def powershell_string(value):
    return "'" + str(value or "").replace("'", "''") + "'"


def render_folder_array(rows):
    lines = ["$FoldersToDelete = @("]
    for row in rows:
        lines.append("    [pscustomobject]@{")
        for key, value in row.items():
            if isinstance(value, int):
                lines.append(f"        {key} = {value}")
            else:
                lines.append(f"        {key} = {powershell_string(value)}")
        lines.append("    }")
    lines.append(")")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--comparison-directory", required=True)
    parser.add_argument("--output-ps1", required=True)
    args = parser.parse_args()

    output_ps1 = Path(args.output_ps1)
    rows = build_delete_folder_rows(Path(args.comparison_directory))

    if not rows:
        if output_ps1.exists():
            output_ps1.unlink()
        print("Folder deletion script generation skipped.")
        print("Reason: no SPO folder path was found missing from the SP2019 source file paths.")
        print(f"PowerShell file not created: {output_ps1}")
        return

    output_ps1.parent.mkdir(parents=True, exist_ok=True)
    content = SCRIPT_HEADER + render_folder_array(rows) + SCRIPT_FOOTER
    output_ps1.write_text(content, encoding="utf-8")

    print(f"PowerShell file created: {output_ps1}")
    print(f"Folders embedded: {len(rows)}")


if __name__ == "__main__":
    main()
