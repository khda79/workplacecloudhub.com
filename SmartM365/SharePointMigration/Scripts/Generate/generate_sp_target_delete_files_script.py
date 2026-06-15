import argparse
import builtins
import csv
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Export"))

from export_files_to_delete_excel import build_delete_file_rows
from export_libraries_to_delete_excel import as_int
from export_comparison_to_excel import detect_csv_dialect


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if args:
        args = (f"{timestamp} {args[0]}", *args[1:])
    else:
        args = (timestamp,)
    builtins.print(*args, **kwargs)


SCRIPT_HEADER = r'''<#
.SYNOPSIS
    Deletes the SharePoint Online files listed in this script.

.DESCRIPTION
    This script is generated for one comparison result. The target files are
    embedded directly in the $FilesToDelete array.

    By default, this script is a dry run. Add -Execute to permanently delete
    the files. Add -Recycle only if files must be moved to the recycle bin.
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

function ConvertFrom-JwtPayload {
    param([string]$Token)

    if ([string]::IsNullOrWhiteSpace($Token) -or $Token.Split('.').Count -lt 2) {
        return $null
    }

    try {
        $payload = $Token.Split('.')[1]
        while ($payload.Length % 4 -ne 0) {
            $payload += '='
        }

        $payload = $payload.Replace('-', '+').Replace('_', '/')
        $json = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload))
        return $json | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Get-PnPTokenDiagnostic {
    param($Connection)

    $diagnostic = [ordered]@{
        AppId  = '<unknown>'
        Scopes = '<unknown>'
        Roles  = '<unknown>'
    }

    if (-not (Get-Command -Name Get-PnPAccessToken -ErrorAction SilentlyContinue)) {
        $diagnostic.Scopes = '<Get-PnPAccessToken unavailable>'
        return [pscustomobject]$diagnostic
    }

    try {
        try {
            $token = Get-PnPAccessToken -ResourceTypeName SharePoint -Connection $Connection -ErrorAction Stop
        }
        catch {
            $token = Get-PnPAccessToken -Connection $Connection -ErrorAction Stop
        }

        $claims = ConvertFrom-JwtPayload -Token $token
        if ($null -eq $claims) {
            $diagnostic.Scopes = '<token claims unavailable>'
            return [pscustomobject]$diagnostic
        }

        $appId = $claims.appid
        if (-not $appId) { $appId = $claims.azp }
        if (-not $appId) { $appId = $claims.aud }

        $scopes = if ($claims.scp) { $claims.scp } else { '<none>' }
        $roles = if ($claims.roles) { ($claims.roles -join ' ') } else { '<none>' }

        $diagnostic.AppId = $appId
        $diagnostic.Scopes = $scopes
        $diagnostic.Roles = $roles
        return [pscustomobject]$diagnostic
    }
    catch {
        $diagnostic.Scopes = ("<token diagnostic failed: {0}>" -f $_.Exception.Message)
        return [pscustomobject]$diagnostic
    }
}

function Remove-TargetFile {
    param(
        [Parameter(Mandatory = $true)]
        $Connection,

        [Parameter(Mandatory = $true)]
        [string]$ServerRelativeUrl,

        [switch]$Recycle
    )

    $removeParameters = @{
        ServerRelativeUrl = $ServerRelativeUrl
        Force             = $true
        Connection        = $Connection
        ErrorAction       = 'Stop'
    }

    if ($Recycle) {
        $removeParameters.Recycle = $true
    }

    try {
        Remove-PnPFile @removeParameters
        return [pscustomobject]@{
            Success = $true
            Method  = if ($Recycle) { 'Remove-PnPFile -Recycle' } else { 'Remove-PnPFile permanent delete' }
            Message = if ($Recycle) { 'File moved to recycle bin.' } else { 'File permanently deleted.' }
        }
    }
    catch {
        $primaryMessage = $_.Exception.Message
    }

    try {
        $listItem = Get-PnPFile -Url $ServerRelativeUrl -AsListItem -Connection $Connection -ErrorAction Stop
        $context = Get-PnPContext -Connection $Connection

        if ($Recycle) {
            [void]$listItem.Recycle()
            $method = 'ListItem.Recycle fallback'
            $successMessage = 'File moved to recycle bin by list item fallback.'
        }
        else {
            $listItem.DeleteObject()
            $method = 'ListItem.DeleteObject fallback'
            $successMessage = 'File permanently deleted by list item fallback.'
        }

        $context.ExecuteQuery()
        return [pscustomobject]@{
            Success = $true
            Method  = $method
            Message = ("{0} Primary Remove-PnPFile error was: {1}" -f $successMessage, $primaryMessage)
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Method  = 'Remove-PnPFile + ListItem fallback'
            Message = ("Primary Remove-PnPFile error: {0}; List item fallback error: {1}" -f $primaryMessage, $_.Exception.Message)
        }
    }
}

function Write-DeleteStatus {
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
            Write-Host ("DRY RUN Would delete: {0} | {1}" -f $WebUrl, $ServerRelativeUrl) -ForegroundColor DarkCyan
        }
        'SkippedByShouldProcess' {
            Write-Host ("SKIP ShouldProcess: {0} | {1}" -f $WebUrl, $ServerRelativeUrl) -ForegroundColor Yellow
        }
        default {
            Write-Warning ("FAILED [{0}] {1} | {2}: {3}" -f $Method, $WebUrl, $ServerRelativeUrl, $Message)
        }
    }
}

'''


SCRIPT_FOOTER = r'''

if ([string]::IsNullOrWhiteSpace($ResultPath)) {
    $ResultPath = Join-Path -Path $PSScriptRoot -ChildPath ("Files-To-Delete-Deletion-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
}
$script:DeletionResultPath = $ResultPath

if ($Execute) {
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'logs') -ChildPath ("Files-To-Delete-Deletion-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
    }

    $logDirectory = Split-Path -Path $LogPath -Parent
    if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
}

Import-PnPPowerShellModule

if (-not $FilesToDelete -or $FilesToDelete.Count -eq 0) {
    Write-Host "No files are embedded in this script."
    return
}

if ($PSBoundParameters.ContainsKey('Test')) {
    $FilesToDelete = @($FilesToDelete | Select-Object -First $Test)
    Write-Host ("Test mode: only the first {0} embedded files will be processed." -f $FilesToDelete.Count) -ForegroundColor Yellow
}

Write-Host ("Files embedded: {0}" -f $FilesToDelete.Count)
$modeText = $(if ($Execute) { if ($Recycle) { 'EXECUTE - files will be recycled' } else { 'EXECUTE - files will be permanently deleted' } } else { 'DRY RUN - no deletion' })
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
        Write-Warning "This will move SharePoint Online files to the recycle bin."
        $expectedConfirmation = 'RECYCLE'
    }
    else {
        Write-Warning "This will permanently delete SharePoint Online files. Recovery may not be possible."
        $expectedConfirmation = 'DELETE'
    }

    $confirmation = Read-Host ("Type {0} to continue" -f $expectedConfirmation)
    if ($confirmation -ne $expectedConfirmation) {
        throw "Deletion cancelled."
    }
}

$script:DeletionResults = New-Object System.Collections.Generic.List[object]
$results = $script:DeletionResults
$forceAuthenticationPending = [bool]$ForceAuthentication

foreach ($group in ($FilesToDelete | Group-Object TargetWebUrl)) {
    $webUrl = $group.Name
    Write-Host ("Connecting interactively to: {0}" -f $webUrl)
    $connection = Connect-InteractiveOnly -Url $webUrl -UseForceAuthentication:$forceAuthenticationPending
    $forceAuthenticationPending = $false
    $connectedAccount = Get-PnPConnectionIdentity -Connection $connection -Url $webUrl
    Write-Host ("Connected account for {0}: {1}" -f $webUrl, $connectedAccount.DisplayText) -ForegroundColor Cyan
    $tokenDiagnostic = Get-PnPTokenDiagnostic -Connection $connection
    Write-Host ("PnP token app id: {0}" -f $tokenDiagnostic.AppId) -ForegroundColor Cyan
    $scopeColor = if ($tokenDiagnostic.Scopes -match 'AllSites\.(Write|Manage|FullControl)') { 'Green' } else { 'Yellow' }
    Write-Host ("PnP token scopes: {0}" -f $tokenDiagnostic.Scopes) -ForegroundColor $scopeColor
    Write-Host ("PnP token roles: {0}" -f $tokenDiagnostic.Roles) -ForegroundColor Cyan

    foreach ($item in $group.Group) {
        $actionStatus = if ($Execute) { if ($Recycle) { 'PendingRecycle' } else { 'PendingPermanentDelete' } } else { 'WouldDelete' }
        $message = ''

        if ([string]::IsNullOrWhiteSpace($item.TargetServerRelativeUrl)) {
            $actionStatus = 'MissingServerRelativeUrl'
            $message = 'Target server relative URL is empty.'
        }
        elseif ($Execute) {
            $targetName = "{0} | {1}" -f $webUrl, $item.TargetServerRelativeUrl
            $operation = if ($Recycle) { 'Recycle SharePoint Online file' } else { 'Permanently delete SharePoint Online file' }
            if ($PSCmdlet.ShouldProcess($targetName, $operation)) {
                $deleteResult = Remove-TargetFile -Connection $connection -ServerRelativeUrl $item.TargetServerRelativeUrl -Recycle:$Recycle
                if ($deleteResult.Success) {
                    $actionStatus = if ($Recycle) { 'Recycled' } else { 'PermanentlyDeleted' }
                    $deleteMethod = $deleteResult.Method
                    $message = $deleteResult.Message
                }
                else {
                    $actionStatus = 'DeleteFailed'
                    $deleteMethod = $deleteResult.Method
                    $message = $deleteResult.Message
                }
            }
            else {
                $actionStatus = 'SkippedByShouldProcess'
                $deleteMethod = 'SkippedByShouldProcess'
            }
        }
        else {
            $deleteMethod = 'DryRun'
        }

        Write-DeleteStatus -Status $actionStatus -Method $deleteMethod -WebUrl $webUrl -ServerRelativeUrl $item.TargetServerRelativeUrl -Message $message

        $results.Add([pscustomobject]@{
                Time                    = Get-Date
                ActionStatus            = $actionStatus
                TargetWebUrl            = $webUrl
                TargetLibraryTitle      = $item.TargetLibraryTitle
                FileName                = $item.FileName
                TargetFileUrl           = $item.TargetFileUrl
                TargetServerRelativeUrl = $item.TargetServerRelativeUrl
                ConnectedAccount        = $connectedAccount.DisplayText
                ConnectedAccountLogin   = $connectedAccount.LoginName
                ConnectedAccountEmail   = $connectedAccount.Email
                TokenAppId              = $tokenDiagnostic.AppId
                TokenScopes             = $tokenDiagnostic.Scopes
                TokenRoles              = $tokenDiagnostic.Roles
                DeleteMethod            = $deleteMethod
                SizeBytes               = $item.SizeBytes
                SizeMB                  = $item.SizeMB
                Modified                = $item.Modified
                ModifiedBy              = $item.ModifiedBy
                Message                 = $message
            })
        Export-DeletionResults -Quiet
    }
}

Export-DeletionResults

if ($script:TranscriptStarted) {
    Stop-TimestampedTranscript -Path $LogPath
}
'''


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


def render_file_array(rows):
    json_rows = json.dumps(rows, ensure_ascii=True, indent=2)
    return "$FilesToDeleteJson = @'\n" + json_rows + "\n'@\n$FilesToDelete = @($FilesToDeleteJson | ConvertFrom-Json)"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--comparison-directory", required=True)
    parser.add_argument("--output-ps1")
    parser.add_argument("--allow-duplicate-keys", action="store_true")
    args = parser.parse_args()

    comparison_directory = Path(args.comparison_directory)
    extra_in_target_path = comparison_directory / "ExtraInTarget.csv"
    if not extra_in_target_path.exists():
        raise FileNotFoundError(f"ExtraInTarget.csv not found: {extra_in_target_path}")

    output_ps1 = Path(args.output_ps1) if args.output_ps1 else comparison_directory / "SmartM365-SharePointTarget-ExtraFilesRemove-Custom.ps1"
    duplicate_count = duplicate_key_count(comparison_directory)
    if duplicate_count > 0 and not args.allow_duplicate_keys:
        duplicate_path = comparison_directory / "DuplicateKeys.csv"
        duplicate_excel_path = comparison_directory / "DuplicateKeys.xlsx"
        print("")
        print("File deletion script generation skipped.")
        print("Reason: duplicate inventory keys were found.")
        print(f"Duplicate keys ignored: {duplicate_count}")
        print(f"Review CSV: {duplicate_path}")
        if duplicate_excel_path.exists():
            print(f"Review Excel: {duplicate_excel_path}")
        print("After validation, rerun with --allow-duplicate-keys to generate the file deletion script.")
        return 2

    rows = build_delete_file_rows(extra_in_target_path)
    if not rows:
        if output_ps1.exists():
            output_ps1.unlink()
        print("File deletion script generation skipped.")
        print("Reason: no extra files were found in ExtraInTarget.csv.")
        print(f"PowerShell file not created: {output_ps1}")
        return 0

    content = SCRIPT_HEADER + render_file_array(rows) + SCRIPT_FOOTER
    output_ps1.write_text(content, encoding="utf-8")

    print(f"PowerShell file created: {output_ps1}")
    print(f"Files embedded: {len(rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
