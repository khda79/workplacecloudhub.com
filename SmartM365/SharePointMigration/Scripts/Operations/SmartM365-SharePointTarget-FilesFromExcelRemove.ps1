<#
.SYNOPSIS
    Deletes SharePoint Online files listed in an Excel file.

.DESCRIPTION
    The Excel file must contain a TargetServerRelativeUrl column.
    Authentication is explicit interactive authentication only. The script does
    not use app-only authentication, certificate authentication, or environment
    credentials.

    By default, files are permanently deleted. Use -Recycle to move files to the
    SharePoint recycle bin instead.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExcelPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [switch]$Recycle,

    [switch]$Force,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [ValidateNotNullOrEmpty()]
    [string]$ResultPath
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
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
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
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Stop-Transcript -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $false
    Add-TimestampToLogFile -Path $Path
}

function Get-DefaultOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$Suffix,

        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    $inputItem = Get-Item -LiteralPath $InputPath -ErrorAction Stop
    $directory = $inputItem.DirectoryName
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name)
    return Join-Path -Path $directory -ChildPath ("{0}-{1}-{2:yyyyMMdd-HHmmss}.{3}" -f $baseName, $Suffix, (Get-Date), $Extension)
}

function Get-DefaultLogPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$Suffix
    )

    $inputItem = Get-Item -LiteralPath $InputPath -ErrorAction Stop
    $directory = Join-Path -Path $inputItem.DirectoryName -ChildPath 'logs'
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name)
    return Join-Path -Path $directory -ChildPath ("{0}-{1}-{2:yyyyMMdd-HHmmss}.log" -f $baseName, $Suffix, (Get-Date))
}

function Import-RequiredModule {
    param([string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required PowerShell module '$Name' is not installed."
    }

    Import-Module $Name -ErrorAction Stop
}

function Connect-ExplicitInteractivePnP {
    param([string]$Url)

    $parameters = @{
        Url              = $Url
        Interactive      = $true
        ReturnConnection = $true
    }

    $connectCommand = Get-Command -Name Connect-PnPOnline -ErrorAction Stop
    if ($connectCommand.Parameters.ContainsKey('ForceAuthentication')) {
        $parameters.ForceAuthentication = $true
    }

    Write-Host ("Connecting with explicit interactive authentication: {0}" -f $Url)
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

function Add-Result {
    param(
        [System.Collections.Generic.List[object]]$Results,
        [string]$TargetServerRelativeUrl,
        [string]$Action,
        [string]$Status,
        [string]$Message,
        $ConnectedAccount
    )

    $Results.Add([pscustomobject]@{
            Time                    = Get-Date
            SiteUrl                 = $SiteUrl
            TargetServerRelativeUrl = $TargetServerRelativeUrl
            Action                  = $Action
            Status                  = $Status
            ConnectedAccount        = if ($ConnectedAccount) { $ConnectedAccount.DisplayText } else { '' }
            ConnectedAccountLogin   = if ($ConnectedAccount) { $ConnectedAccount.LoginName } else { '' }
            ConnectedAccountEmail   = if ($ConnectedAccount) { $ConnectedAccount.Email } else { '' }
            Message                 = $Message
        })
}

try {
    $ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).ProviderPath

    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Get-DefaultLogPath -InputPath $ExcelPath -Suffix 'Deletion-Run'
    }

    if ([string]::IsNullOrWhiteSpace($ResultPath)) {
        $ResultPath = Get-DefaultOutputPath -InputPath $ExcelPath -Suffix 'Deletion-Results' -Extension 'csv'
    }

    $logDirectory = Split-Path -Path $LogPath -Parent
    if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    $resultDirectory = Split-Path -Path $ResultPath -Parent
    if ($resultDirectory -and -not (Test-Path -LiteralPath $resultDirectory)) {
        New-Item -Path $resultDirectory -ItemType Directory -Force | Out-Null
    }

    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true

    Write-Host ("Excel input: {0}" -f $ExcelPath)
    Write-Host ("Site URL: {0}" -f $SiteUrl)
    Write-Host ("Run log: {0}" -f $LogPath)
    Write-Host ("Result CSV: {0}" -f $ResultPath)
    Write-Host ("Mode: {0}" -f ($(if ($Recycle) { 'RECYCLE' } else { 'PERMANENT DELETE' })))

    Import-RequiredModule -Name PnP.PowerShell
    Import-RequiredModule -Name ImportExcel

    $rows = @(Import-Excel -Path $ExcelPath)

    if (-not $rows -or $rows.Count -eq 0) {
        throw "Excel file contains no rows."
    }

    if (-not ($rows | Get-Member -Name TargetServerRelativeUrl -MemberType NoteProperty)) {
        throw "Column 'TargetServerRelativeUrl' not found in Excel file."
    }

    $files = @(
        $rows |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.TargetServerRelativeUrl) } |
            Select-Object -ExpandProperty TargetServerRelativeUrl -Unique
    )

    if ($files.Count -eq 0) {
        throw "No non-empty TargetServerRelativeUrl values were found."
    }

    Write-Host ("Files loaded: {0}" -f $files.Count)

    if (-not $Force) {
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

    $connection = Connect-ExplicitInteractivePnP -Url $SiteUrl
    $connectedAccount = Get-PnPConnectionIdentity -Connection $connection -Url $SiteUrl
    Write-Host ("Connected account for {0}: {1}" -f $SiteUrl, $connectedAccount.DisplayText) -ForegroundColor Cyan

    $results = New-Object 'System.Collections.Generic.List[object]'
    $action = if ($Recycle) { 'Recycle' } else { 'PermanentDelete' }

    foreach ($fileUrl in $files) {
        try {
            Write-Host ("Deleting: {0}" -f $fileUrl)

            if ($PSCmdlet.ShouldProcess($fileUrl, $action)) {
                $removeParameters = @{
                    ServerRelativeUrl = $fileUrl
                    Force             = $true
                    Connection        = $connection
                }

                if ($Recycle) {
                    $removeParameters.Recycle = $true
                }

                Remove-PnPFile @removeParameters
                Add-Result -Results $results -TargetServerRelativeUrl $fileUrl -Action $action -Status 'Success' -Message '' -ConnectedAccount $connectedAccount
            }
            else {
                Add-Result -Results $results -TargetServerRelativeUrl $fileUrl -Action $action -Status 'SkippedByShouldProcess' -Message '' -ConnectedAccount $connectedAccount
            }
        }
        catch {
            Write-Warning ("Failed: {0}" -f $fileUrl)
            Write-Warning $_.Exception.Message
            Write-Warning ("Connected account: {0}" -f $connectedAccount.DisplayText)
            Add-Result -Results $results -TargetServerRelativeUrl $fileUrl -Action $action -Status 'Failed' -Message $_.Exception.Message -ConnectedAccount $connectedAccount
        }
    }

    $results | Export-Csv -Delimiter ';' -Path $ResultPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Deletion result written: {0}" -f $ResultPath)
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue

    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
}
