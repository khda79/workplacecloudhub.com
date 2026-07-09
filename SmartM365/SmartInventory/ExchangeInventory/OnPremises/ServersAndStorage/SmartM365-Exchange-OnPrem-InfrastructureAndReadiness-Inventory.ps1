<#
.SYNOPSIS
    Inventory Exchange 2016 servers, CPU, RAM and disks using WMI/DCOM only.

.DESCRIPTION
    This script inventories Exchange 2016 servers from the Exchange organization and collects:
    - Exchange server identity
    - CPU and RAM inventory
    - Logical disks from Win32_LogicalDisk
    - Disk drives from Win32_DiskDrive
    - Optional Exchange mailbox database paths
    - Optional Exchange service health
    - Exchange schema and migration readiness configuration
    - Per-server decommissioning summary
    - Global CSV summary
    - HTML executive summary for slide integration

    This version does not use:
    - Get-CimInstance
    - New-CimSession
    - Invoke-Command
    - WinRM / PowerShell Remoting

.VERSION
    1.4.6

.NOTES
    Script Name : SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1
    Version     : 1.4.7
    Requirements:
      - Windows PowerShell 5.1 with Exchange 2016 Management Tools
      - Exchange read permissions
      - Remote WMI/DCOM access to Exchange servers
      - PowerShell 5.1 or later

.CHANGELOG
    1.4.6
      - Adds detailed readiness collector progress logs to identify long-running Exchange cmdlets.

    1.4.5
      - Writes LOG-ALL log/transcript, publishes DATA-LAST copies, and uploads weekly history.
      - Sends the summary mail with the shared SmartM365 template and SharePoint links.
      - Keeps launcher snap-in loading inside the script for script-scope validation.

    1.4.4
      - Loads the Exchange 2016 PowerShell snap-in directly when required.

    1.4.3
      - Allows readiness collectors to start with an empty row list.

    1.4.2
      - Renames the script to Exchange OnPrem Infrastructure and Readiness inventory.

    1.4.1
      - Renders the full Exchange migration readiness inventory in the HTML report.

    1.4.0
      - Adds AD Exchange schema and Exchange migration readiness configuration inventory.

    1.3.7
      - Applies the script local JSON overrides before reading mail and SharePoint settings.

    1.3.6
      - Fixes HTML email sending after prefixed SmartM365.Core import.

    1.3.5
      - Adds SendMailMode support for Graph, SMTP, or Graph with SMTP fallback.

    1.3.4
      - Keeps SharePoint upload available but disabled by default in local configuration.

    1.3.3
      - Restyled the HTML executive summary with Smart365 branding.

    1.3.2
      - Sends the HTML executive summary by email after report generation.

    1.3.1
      - WMI/DCOM only.
      - Fixed StrictMode-safe numeric aggregation.
      - Added HTML executive summary for slide integration.
      - Kept CPU, RAM, logical disks, disk drives, mailbox database paths and service health outputs.

    1.2.1
      - Fixed PowerShell variable interpolation before colon.

    1.2.0
      - Added CPU and RAM inventory.
      - Added Win32_DiskDrive inventory.
      - Added per-server infrastructure summary.
      - Added global decommissioning summary.
      - Added HTML executive summary.
#>

[CmdletBinding()]
param(
    [string]$Tenant = "test",

    [string]$OutputRoot,

    [switch]$IncludeServicesHealth,

    [switch]$IncludeMailboxDatabasePaths,

    [switch]$SkipFqdnAndUseServerName,

    [int]$LowFreeSpaceThresholdPercent = 15
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath

$ScriptName = "SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory"
$ScriptVersion = "1.4.7"
$RunId = (Get-Date).ToString("yyyyMMdd-HHmmss")

$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

function Merge-SmartM365ScriptLocalConfig {
    param(
        [Parameter(Mandatory)]
        [string]$ConfigPath
    )

    $localConfig = Read-SmartM365JsonConfig -Path $ConfigPath -Required
    if ($null -eq $localConfig -or $localConfig.Count -eq 0) { return }

    $effectiveConfig = ConvertTo-SmartM365Hashtable -InputObject $script:SmartM365EffectiveConfig
    foreach ($key in $localConfig.Keys) {
        $value = $localConfig[$key]
        if ($value -is [string] -and $value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { continue }
        $effectiveConfig[$key] = $value
    }

    $script:SmartM365EffectiveConfig = [pscustomobject]$effectiveConfig
}

Merge-SmartM365ScriptLocalConfig -ConfigPath (Join-Path -Path $PSScriptRoot -ChildPath "$ScriptName.local.json")

function Resolve-SmartM365ConfigTokenValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    $resolved = [string]$Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$tokenName]
            if ($null -eq $property -or $null -eq $property.Value) { continue }

            $tokenValue = Resolve-SmartM365ConfigTokenValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Resolve-SmartM365ConfigTokenValue -Value '{{DataAllRootPath}}\Exchange\OnPrem\ServersAndStorage'
}
else {
    $OutputRoot = Resolve-SmartM365ConfigTokenValue -Value $OutputRoot
}

$OutputFolder = Join-Path $OutputRoot $RunId
$LogRoot = Resolve-SmartM365ConfigTokenValue -Value '{{LogAllRootPath}}'
if ([string]::IsNullOrWhiteSpace($LogRoot) -or $LogRoot -eq '{{LogAllRootPath}}') {
    $LogRoot = Join-Path -Path (Split-Path -Path $OutputRoot -Parent) -ChildPath 'LOG-ALL'
}
$LogFolder = Join-Path -Path $LogRoot -ChildPath $ScriptName
$LogFile = Join-Path $LogFolder "$ScriptName-$RunId.log"
$TranscriptFile = Join-Path $LogFolder "$ScriptName-$RunId`_Transcript.log"
$script:TranscriptStarted = $false
$script:ServersAndStorageGeneratedCsvPaths = New-Object 'System.Collections.Generic.List[string]'
$script:ServersAndStorageSharePointUploads = New-Object System.Collections.ArrayList
$script:ServersAndStorageWarningCount = 0
$script:ServersAndStorageErrorCount = 0

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    if ($Level -eq 'WARN') { $script:ServersAndStorageWarningCount++ }
    if ($Level -eq 'ERROR') { $script:ServersAndStorageErrorCount++ }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { "$timestamp [$Level] $_" })
    $line | ForEach-Object { Write-Host $_ }
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Stop-ServersAndStorageTranscriptSafely {
    if (-not $script:TranscriptStarted) { return }
    try { Stop-Transcript | Out-Null } catch {}
    $script:TranscriptStarted = $false
}

function Complete-ServersAndStorageRun {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][datetime]$Started,
        [AllowNull()][string]$ErrorMessage
    )

    $ended = Get-Date
    Write-Log 'Execution summary:'
    Write-Log ("  Status: {0}" -f $Status)
    Write-Log ("  Duration: {0}" -f ($ended - $Started))
    Write-Log ("  Started: {0}" -f $Started.ToString('yyyy-MM-dd HH:mm:ss zzz'))
    Write-Log ("  Ended: {0}" -f $ended.ToString('yyyy-MM-dd HH:mm:ss zzz'))
    Write-Log ("  ScriptName: {0}" -f $ScriptName)
    Write-Log ("  ScriptVersion: {0}" -f $ScriptVersion)
    Write-Log ("  TenantKey: {0}" -f $Tenant)
    Write-Log ("  OutputPath: {0}" -f $OutputFolder)
    Write-Log ("  LogTextFile: {0}" -f $LogFile)
    Write-Log ("  TranscriptFile: {0}" -f $TranscriptFile)
    Write-Log ("  GeneratedCsvFiles: {0}" -f $script:ServersAndStorageGeneratedCsvPaths.Count)
    Write-Log ("  Warnings: {0}" -f $script:ServersAndStorageWarningCount)
    Write-Log ("  Errors: {0}" -f $script:ServersAndStorageErrorCount)
    if (-not [string]::IsNullOrWhiteSpace($ErrorMessage)) { Write-Log ("  ErrorMessage: {0}" -f $ErrorMessage) }

    Stop-ServersAndStorageTranscriptSafely
    Invoke-ServersAndStorageSharePointUpload -LocalFilePath $LogFile
    if (Test-Path -LiteralPath $TranscriptFile -PathType Leaf) {
        Invoke-ServersAndStorageSharePointUpload -LocalFilePath $TranscriptFile
    }
}
function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Get-SmartM365EffectiveConfigValue {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [AllowNull()]
        [object]$DefaultValue = $null
    )

    if ($null -ne $script:SmartM365EffectiveConfig) {
        $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$Name]
        if ($null -ne $property -and $null -ne $property.Value) {
            if ($property.Value -is [string] -and [string]::IsNullOrWhiteSpace($property.Value)) {
                return $DefaultValue
            }

            return Resolve-SmartM365ConfigTokenValue -Value $property.Value
        }
    }

    return $DefaultValue
}

function Import-SmartM365CoreModule {
    if ((Get-Command Publish-CoreSmartM365Csv -ErrorAction SilentlyContinue) -and (Get-Command CoreSendEmailHtmlReport -ErrorAction SilentlyContinue)) {
        return
    }

    $d = $PSScriptRoot
    while ($d) {
        $moduleCandidates = if ($PSVersionTable.PSVersion.Major -lt 6) {
            @(
                (Join-Path -Path $d -ChildPath 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'),
                (Join-Path -Path $d -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1')
            )
        }
        else {
            @(
                (Join-Path -Path $d -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1')
            )
        }

        foreach ($modulePath in $moduleCandidates) {
            if (Test-Path -LiteralPath $modulePath) {
                Import-Module $modulePath -Prefix Core -ErrorAction Stop
                return
            }
        }

        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }

    throw 'SmartM365.Core module not found.'
}

function Invoke-ServersAndStorageSharePointUpload {
    param(
        [Parameter(Mandatory)]
        [string]$LocalFilePath
    )

    $enabled = [bool](Get-SmartM365EffectiveConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false)
    if (-not $enabled) { return }

    $thumbprint = Get-SmartM365EffectiveConfigValue -Name 'Thumbprint' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($thumbprint)) {
        $thumbprint = Get-SmartM365EffectiveConfigValue -Name 'Thumb' -DefaultValue ''
    }

    try {
        Import-SmartM365CoreModule
        $record = Invoke-CoreSmartM365SharePointCsvUpload `
            -LocalFilePath $LocalFilePath `
            -Enabled $true `
            -SiteHostname (Get-SmartM365EffectiveConfigValue -Name 'SharePointSiteHostname' -DefaultValue '') `
            -SitePath (Get-SmartM365EffectiveConfigValue -Name 'SharePointSitePath' -DefaultValue '') `
            -LibraryDisplayName (Get-SmartM365EffectiveConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents') `
            -TargetFolderPath (Get-SmartM365EffectiveConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue '') `
            -AppId (Get-SmartM365EffectiveConfigValue -Name 'AppId' -DefaultValue '') `
            -TenantId (Get-SmartM365EffectiveConfigValue -Name 'TenantId' -DefaultValue '') `
            -Thumbprint $thumbprint
        if ($record) {
            [void]$script:ServersAndStorageSharePointUploads.Add($record)
        }
        Write-Log "SharePoint upload requested: $LocalFilePath"
        return $record
    }
    catch {
        Write-Log "SharePoint upload failed (non-blocking): $($_.Exception.Message)" "WARN"
    }
}

function ConvertTo-ServersAndStorageEmailHtmlText {
    param([AllowNull()]$Value)

    if (Get-Command ConvertTo-CoreSmartM365EmailHtmlText -ErrorAction SilentlyContinue) {
        return ConvertTo-CoreSmartM365EmailHtmlText -Value $Value
    }

    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-ServersAndStorageEmailLinkHtml {
    param(
        [AllowNull()][string]$Text,
        [AllowNull()][string]$Url
    )

    $safeText = ConvertTo-ServersAndStorageEmailHtmlText -Value $Text
    if (-not [string]::IsNullOrWhiteSpace($Url)) {
        $safeUrl = ConvertTo-ServersAndStorageEmailHtmlText -Value $Url
        return "<a href=`"$safeUrl`" style=`"color:#2563eb;text-decoration:underline;`">$safeText</a>"
    }

    return $safeText
}

function New-ServersAndStorageSharePointLinksSectionHtml {
    param([array]$UploadRecords)

    $records = @($UploadRecords | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.WebUrl) })
    if ($records.Count -eq 0) { return $null }

    $rowsHtml = foreach ($record in $records) {
        $label = if ($record.FileName) { [string]$record.FileName } else { [string]$record.SharePointPath }
        $pathText = if ($record.SharePointPath) { [string]$record.SharePointPath } else { [string]$record.WebUrl }
        $linkHtml = New-ServersAndStorageEmailLinkHtml -Text $pathText -Url ([string]$record.WebUrl)
        "<tr><td style=`"width:220px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $label)</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;word-break:break-all;`">$linkHtml</td></tr>"
    }

    return @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  $($rowsHtml -join "`n")
</table>
"@
}

function Send-ServersAndStorageHtmlReport {
    param(
        [Parameter(Mandatory)]
        [string]$HtmlReportPath,

        [Parameter(Mandatory)]
        [string]$Subject,

        [Parameter(Mandatory)]
        [string[]]$Attachments,

        [Parameter(Mandatory)]
        [pscustomobject]$Summary,

        [Parameter(Mandatory)]
        [array]$PerServerSummary,

        [Parameter(Mandatory)]
        [array]$ReadinessInventory
    )

    if (-not (Test-Path -LiteralPath $HtmlReportPath -PathType Leaf)) {
        throw "HTML report not found: $HtmlReportPath"
    }

    Import-SmartM365CoreModule

    $from = [string](Get-SmartM365EffectiveConfigValue -Name 'From' -DefaultValue '')
    $to = [string](Get-SmartM365EffectiveConfigValue -Name 'To' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($to)) {
        $to = [string](Get-SmartM365EffectiveConfigValue -Name 'ErrorMailTo' -DefaultValue '')
    }

    $smtpServer = [string](Get-SmartM365EffectiveConfigValue -Name 'SmtpServer' -DefaultValue '')
    $sendMailMode = [string](Get-SmartM365EffectiveConfigValue -Name 'SendMailMode' -DefaultValue '')
    $cc = [string](Get-SmartM365EffectiveConfigValue -Name 'Cc' -DefaultValue '')
    $existingAttachments = @($Attachments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Leaf) })

    $summaryRows = @(
        [pscustomobject]@{ Label = 'Exchange servers'; Value = $Summary.ExchangeServersCount }
        [pscustomobject]@{ Label = 'Total memory (TB)'; Value = $Summary.TotalMemoryTB }
        [pscustomobject]@{ Label = 'Total logical disk size (TB)'; Value = $Summary.TotalLogicalDiskSizeTB }
        [pscustomobject]@{ Label = 'Logical disk free (TB)'; Value = $Summary.TotalLogicalDiskFreeTB }
        [pscustomobject]@{ Label = 'Low space warnings'; Value = $Summary.LowSpaceWarnings }
        [pscustomobject]@{ Label = 'Readiness warnings'; Value = $Summary.ExchangeReadinessWarnings }
        [pscustomobject]@{ Label = 'Readiness errors'; Value = $Summary.ExchangeReadinessErrors }
    )

    $sections = @()
    $perServerRows = @($PerServerSummary | Sort-Object ExchangeServerName | Select-Object -First 50)
    if ($perServerRows.Count -gt 0) {
        $serverRowsHtml = foreach ($server in $perServerRows) {
            "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $server.ExchangeServerName)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $server.ServerRole)</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $server.MemoryGB)</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $server.LogicalDiskFreeGB)</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#92400e;font-weight:700;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $server.LowSpaceLogicalDiskCount)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $server.LogicalDiskCollectionStatus)</td></tr>"
        }
        $serversSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Server</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Role</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Memory GB</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Free GB</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Low disks</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Disk status</th>
  </tr>
  $($serverRowsHtml -join "`n")
</table>
"@
        $sections += [pscustomobject]@{ Title = 'Server preview'; Html = $serversSectionHtml }
    }

    $readinessRows = @($ReadinessInventory | Where-Object { $_.CollectionStatus -eq 'ERROR' -or $_.CollectionStatus -eq 'WARNING' -or $_.Importance -eq 'Error' -or $_.Importance -eq 'Warning' } | Select-Object -First 25)
    if ($readinessRows.Count -gt 0) {
        $readinessRowsHtml = foreach ($row in $readinessRows) {
            "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $row.Category)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $row.Setting)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $row.Value)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;font-weight:700;color:#92400e;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $row.Importance)</td></tr>"
        }
        $readinessSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Category</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Setting</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Value</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Importance</th>
  </tr>
  $($readinessRowsHtml -join "`n")
</table>
"@
        $sections += [pscustomobject]@{ Title = 'Top readiness warnings'; Html = $readinessSectionHtml }
    }

    $sharePointSectionHtml = New-ServersAndStorageSharePointLinksSectionHtml -UploadRecords $script:ServersAndStorageSharePointUploads
    if ($sharePointSectionHtml) {
        $sections += [pscustomobject]@{ Title = 'SharePoint links'; Html = $sharePointSectionHtml }
    }

    $latestCsvFolder = [string](Get-SmartM365EffectiveConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
    $pathRows = @(
        [pscustomobject]@{ Label = 'DATA-ALL run'; Path = $OutputFolder }
        [pscustomobject]@{ Label = 'DATA-LAST'; Path = $latestCsvFolder }
        [pscustomobject]@{ Label = 'LOG-ALL'; Path = $LogFolder }
        [pscustomobject]@{ Label = 'HTML report'; Path = $HtmlReportPath }
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Path) }

    $severity = if ($Summary.ExchangeReadinessErrors -gt 0 -or $Summary.LogicalDiskCollectionErrors -gt 0 -or $Summary.DiskDriveCollectionErrors -gt 0 -or $Summary.ComputeCollectionErrors -gt 0) { 'Warning' } else { 'Success' }
    $actionTitle = if ($severity -eq 'Warning') { 'Review required' } else { '' }
    $actionHtml = if ($severity -eq 'Warning') { 'Review readiness warnings/errors and low disk capacity before Exchange decommissioning or migration decisions.' } else { '' }
    $bodyHtml = New-CoreSmartM365EmailBody `
        -Title $Subject `
        -Category 'SmartM365 Exchange OnPrem' `
        -Severity $severity `
        -Tenant $Tenant `
        -HostName $env:COMPUTERNAME `
        -Message 'Exchange on-premises infrastructure and migration readiness inventory completed.' `
        -ActionTitle $actionTitle `
        -ActionHtml $actionHtml `
        -SummaryRows $summaryRows `
        -Sections $sections `
        -PathRows $pathRows `
        -Footer 'This automated message was generated by SmartM365. Use the exported CSV files and SharePoint links as the source of truth.'

    $mailParams = @{
        From        = $from
        To          = $to
        Subject     = $Subject
        BodyHtml    = $bodyHtml
        Attachments = $existingAttachments
    }
    if (-not [string]::IsNullOrWhiteSpace($smtpServer)) { $mailParams['SmtpServer'] = $smtpServer }
    if (-not [string]::IsNullOrWhiteSpace($sendMailMode)) { $mailParams['SendMailMode'] = $sendMailMode }
    if (-not [string]::IsNullOrWhiteSpace($cc)) { $mailParams['Cc'] = $cc }

    CoreSendEmailHtmlReport @mailParams
}
function Export-ServersAndStorageCsv {
    param(
        [AllowNull()][object[]]$Data,
        [Parameter(Mandatory)]
        [string]$Path
    )

    Import-SmartM365CoreModule
    Publish-CoreSmartM365Csv -Data @($Data) -TimestampedPath $Path -Delimiter ';' -NoSharePointUpload | Out-Null
    $script:ServersAndStorageGeneratedCsvPaths.Add($Path) | Out-Null
    Invoke-ServersAndStorageSharePointUpload -LocalFilePath $Path

    $latestCsvFolder = [string](Get-SmartM365EffectiveConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($latestCsvFolder)) {
        Ensure-Directory -Path $latestCsvFolder
        $latestPath = Join-Path -Path $latestCsvFolder -ChildPath (Split-Path -Path $Path -Leaf)
        Copy-Item -LiteralPath $Path -Destination $latestPath -Force -ErrorAction Stop
        $script:ServersAndStorageGeneratedCsvPaths.Add($latestPath) | Out-Null
        Write-Log ("CSV latest copy written to: {0}" -f $latestPath)
        Invoke-ServersAndStorageSharePointUpload -LocalFilePath $latestPath
    }
}

function Save-ServersAndStorageWeeklyHistory {
    $weeklyHistoryEnabled = [bool](Get-SmartM365EffectiveConfigValue -Name 'EnableWeeklyHistory' -DefaultValue $true)
    if (-not $weeklyHistoryEnabled) { return }

    $weeklyHistoryPath = [string](Get-SmartM365EffectiveConfigValue -Name 'WeeklyHistoryFolderPath' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($weeklyHistoryPath)) { return }

    $weeklyHistoryPath = Resolve-SmartM365ConfigTokenValue -Value $weeklyHistoryPath
    $sourceCsvPaths = @($script:ServersAndStorageGeneratedCsvPaths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -Unique)
    if ($sourceCsvPaths.Count -eq 0) { return }

    try {
        Import-SmartM365CoreModule
        if (Get-Command Add-CoreSmartM365WeeklyHistory -ErrorAction SilentlyContinue) {
            Add-CoreSmartM365WeeklyHistory -SourceCsvPaths $sourceCsvPaths -HistoryRootPath $weeklyHistoryPath -RetentionWeeks ([int](Get-SmartM365EffectiveConfigValue -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)) -HistoryLabel 'Exchange on-prem infrastructure and readiness' | Out-Null
        }
    }
    catch {
        Write-Log "Weekly history upload failed (non-blocking): $($_.Exception.Message)" 'WARN'
    }
}
function Test-ExchangeShell {
    if ($PSVersionTable.PSEdition -ne 'Desktop') {
        throw "Exchange 2016 snap-ins require Windows PowerShell 5.1. Run with C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe."
    }

    $snapinName = 'Microsoft.Exchange.Management.PowerShell.SnapIn'
    if (-not (Get-PSSnapin $snapinName -Registered -ErrorAction SilentlyContinue)) {
        throw "Exchange Management PSSnapin '$snapinName' is not registered. Run this script on an Exchange 2016 server with the Management Tools installed."
    }

    if (-not (Get-PSSnapin $snapinName -ErrorAction SilentlyContinue)) {
        Write-Log "Exchange PSSnapin '$snapinName' is not loaded. Loading it now."
        Add-PSSnapin $snapinName -ErrorAction Stop
    }

    if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
        throw "Exchange cmdlets are not available after loading the Exchange PSSnapin."
    }
}

function ConvertTo-GB {
    param(
        [AllowNull()]
        [object]$Bytes
    )

    if ($null -eq $Bytes) {
        return $null
    }

    return [math]::Round(([double]$Bytes / 1GB), 2)
}

function ConvertTo-TBFromGB {
    param(
        [AllowNull()]
        [object]$GB
    )

    if ($null -eq $GB) {
        return $null
    }

    return [math]::Round(([double]$GB / 1024), 2)
}

function Get-SafeSum {
    param(
        [AllowNull()]
        [object[]]$InputObject,

        [Parameter(Mandatory)]
        [string]$Property
    )

    $sum = 0.0

    foreach ($item in @($InputObject)) {
        if ($null -eq $item) {
            continue
        }

        $propertyValue = $null

        try {
            $propertyValue = $item.$Property
        }
        catch {
            $propertyValue = $null
        }

        if ($null -eq $propertyValue -or [string]::IsNullOrWhiteSpace([string]$propertyValue)) {
            continue
        }

        try {
            $sum += [double]$propertyValue
        }
        catch {
            continue
        }
    }

    return [math]::Round($sum, 2)
}

function Format-HtmlValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return "N/A"
    }

    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        return "N/A"
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-ComputerNamesToTry {
    param(
        [Parameter(Mandatory)]
        [object]$ExchangeServer
    )

    if ($SkipFqdnAndUseServerName) {
        return @($ExchangeServer.Name)
    }

    return @($ExchangeServer.Fqdn, $ExchangeServer.Name) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
}

function Invoke-RemoteWmi {
    param(
        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry,

        [Parameter(Mandatory)]
        [string]$ClassName,

        [string]$Filter
    )

    $errors = New-Object System.Collections.Generic.List[string]

    foreach ($computerName in $ComputerNamesToTry | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
        try {
            $params = @{
                Class        = $ClassName
                ComputerName = $computerName
                ErrorAction  = "Stop"
            }

            if (-not [string]::IsNullOrWhiteSpace($Filter)) {
                $params.Filter = $Filter
            }

            return [pscustomobject]@{
                ComputerNameTried = $computerName
                Method            = "WmiDcom"
                Data              = @(Get-WmiObject @params)
                ErrorMessage      = $null
                Success           = $true
            }
        }
        catch {
            $errors.Add("WmiDcom on $computerName failed for ${ClassName}: $($_.Exception.Message)")
        }
    }

    return [pscustomobject]@{
        ComputerNameTried = ($ComputerNamesToTry -join ";")
        Method            = "WmiDcom"
        Data              = @()
        ErrorMessage      = ($errors -join " | ")
        Success           = $false
    }
}

function Test-RemoteAccess {
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeServerName,

        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry
    )

    foreach ($computerName in $ComputerNamesToTry | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
        $pingOk = $false
        $wmiOk = $false
        $wmiError = $null

        try {
            $pingOk = Test-Connection -ComputerName $computerName -Count 1 -Quiet -ErrorAction Stop
        }
        catch {
            $pingOk = $false
        }

        try {
            $null = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computerName -ErrorAction Stop | Select-Object -First 1
            $wmiOk = $true
        }
        catch {
            $wmiError = $_.Exception.Message
        }

        [pscustomobject]@{
            ExchangeServerName = $ExchangeServerName
            ComputerNameTried  = $computerName
            PingOk             = $pingOk
            WmiDcomOk           = $wmiOk
            WmiDcomError        = $wmiError
        }

        if ($wmiOk) {
            return
        }
    }
}

function Get-ServerComputeInventory {
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeServerName,

        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry
    )

    $csResult = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_ComputerSystem"
    $cpuResult = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_Processor"
    $osResult = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_OperatingSystem"

    if (-not $csResult.Success -or -not $cpuResult.Success) {
        [pscustomobject]@{
            ExchangeServerName    = $ExchangeServerName
            ComputerNameTried     = (($csResult.ComputerNameTried, $cpuResult.ComputerNameTried) | Where-Object { $_ } | Select-Object -Unique) -join ";"
            Manufacturer          = $null
            Model                 = $null
            IsVirtualMachine      = $null
            Domain                = $null
            OSName                = $null
            OSVersion             = $null
            SocketCount           = $null
            PhysicalCoreCount     = $null
            LogicalProcessorCount = $null
            MemoryGB              = $null
            CollectionMethod      = "WmiDcom"
            CollectionStatus      = "ERROR"
            ErrorMessage          = "ComputerSystem: $($csResult.ErrorMessage) | Processor: $($cpuResult.ErrorMessage)"
        }
        return
    }

    $cs = @($csResult.Data) | Select-Object -First 1
    $cpus = @($cpuResult.Data)
    $os = @($osResult.Data) | Select-Object -First 1

    $manufacturer = [string]$cs.Manufacturer
    $model = [string]$cs.Model
    $isVm = if ($manufacturer -match "VMware|Microsoft|Xen|QEMU|Virtual" -or $model -match "Virtual|VMware|KVM|Hyper-V") { $true } else { $false }

    [pscustomobject]@{
        ExchangeServerName    = $ExchangeServerName
        ComputerNameTried     = $csResult.ComputerNameTried
        Manufacturer          = $cs.Manufacturer
        Model                 = $cs.Model
        IsVirtualMachine      = $isVm
        Domain                = $cs.Domain
        OSName                = if ($os) { $os.Caption } else { $null }
        OSVersion             = if ($os) { $os.Version } else { $null }
        SocketCount           = @($cpus).Count
        PhysicalCoreCount     = Get-SafeSum -InputObject $cpus -Property "NumberOfCores"
        LogicalProcessorCount = Get-SafeSum -InputObject $cpus -Property "NumberOfLogicalProcessors"
        MemoryGB              = ConvertTo-GB -Bytes $cs.TotalPhysicalMemory
        CollectionMethod      = "WmiDcom"
        CollectionStatus      = "OK"
        ErrorMessage          = $null
    }
}

function New-LogicalDiskInventoryObject {
    param(
        [string]$ExchangeServerName,
        [string]$ComputerNameTried,
        [object]$Disk,
        [string]$CollectionStatus = "OK",
        [string]$ErrorMessage = $null
    )

    if ($null -eq $Disk) {
        return [pscustomobject]@{
            ExchangeServerName = $ExchangeServerName
            ComputerNameTried  = $ComputerNameTried
            DeviceId           = $null
            VolumeName         = $null
            FileSystem         = $null
            SizeGB             = $null
            UsedGB             = $null
            FreeGB             = $null
            FreePercent        = $null
            LowSpaceWarning    = $null
            CollectionMethod   = "WmiDcom"
            CollectionStatus   = $CollectionStatus
            ErrorMessage       = $ErrorMessage
        }
    }

    $sizeGb = ConvertTo-GB -Bytes $Disk.Size
    $freeGb = ConvertTo-GB -Bytes $Disk.FreeSpace
    $usedGb = if ($null -ne $sizeGb -and $null -ne $freeGb) { [math]::Round(($sizeGb - $freeGb), 2) } else { $null }
    $freePercent = if ($Disk.Size -and [double]$Disk.Size -gt 0) {
        [math]::Round((([double]$Disk.FreeSpace / [double]$Disk.Size) * 100), 2)
    } else {
        $null
    }

    [pscustomobject]@{
        ExchangeServerName = $ExchangeServerName
        ComputerNameTried  = $ComputerNameTried
        DeviceId           = $Disk.DeviceID
        VolumeName         = $Disk.VolumeName
        FileSystem         = $Disk.FileSystem
        SizeGB             = $sizeGb
        UsedGB             = $usedGb
        FreeGB             = $freeGb
        FreePercent        = $freePercent
        LowSpaceWarning    = if ($null -ne $freePercent -and $freePercent -lt $LowFreeSpaceThresholdPercent) { $true } else { $false }
        CollectionMethod   = "WmiDcom"
        CollectionStatus   = $CollectionStatus
        ErrorMessage       = $ErrorMessage
    }
}

function Get-ServerLogicalDiskInventory {
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeServerName,

        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry
    )

    $result = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_LogicalDisk" -Filter "DriveType = 3"

    if (-not $result.Success) {
        New-LogicalDiskInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -Disk $null -CollectionStatus "ERROR" -ErrorMessage $result.ErrorMessage
        return
    }

    if (@($result.Data).Count -eq 0) {
        New-LogicalDiskInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -Disk $null -CollectionStatus "WARNING" -ErrorMessage "No fixed logical disk returned."
        return
    }

    foreach ($disk in @($result.Data)) {
        New-LogicalDiskInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -Disk $disk
    }
}

function New-DiskDriveInventoryObject {
    param(
        [string]$ExchangeServerName,
        [string]$ComputerNameTried,
        [object]$DiskDrive,
        [string]$CollectionStatus = "OK",
        [string]$ErrorMessage = $null
    )

    if ($null -eq $DiskDrive) {
        return [pscustomobject]@{
            ExchangeServerName = $ExchangeServerName
            ComputerNameTried  = $ComputerNameTried
            Index              = $null
            Name               = $null
            DeviceId           = $null
            Model              = $null
            InterfaceType      = $null
            MediaType          = $null
            SizeGB             = $null
            Partitions         = $null
            SerialNumber       = $null
            CollectionMethod   = "WmiDcom"
            CollectionStatus   = $CollectionStatus
            ErrorMessage       = $ErrorMessage
        }
    }

    [pscustomobject]@{
        ExchangeServerName = $ExchangeServerName
        ComputerNameTried  = $ComputerNameTried
        Index              = $DiskDrive.Index
        Name               = $DiskDrive.Name
        DeviceId           = $DiskDrive.DeviceID
        Model              = $DiskDrive.Model
        InterfaceType      = $DiskDrive.InterfaceType
        MediaType          = $DiskDrive.MediaType
        SizeGB             = ConvertTo-GB -Bytes $DiskDrive.Size
        Partitions         = $DiskDrive.Partitions
        SerialNumber       = $DiskDrive.SerialNumber
        CollectionMethod   = "WmiDcom"
        CollectionStatus   = $CollectionStatus
        ErrorMessage       = $ErrorMessage
    }
}

function Get-ServerDiskDriveInventory {
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeServerName,

        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry
    )

    $result = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_DiskDrive"

    if (-not $result.Success) {
        New-DiskDriveInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -DiskDrive $null -CollectionStatus "ERROR" -ErrorMessage $result.ErrorMessage
        return
    }

    if (@($result.Data).Count -eq 0) {
        New-DiskDriveInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -DiskDrive $null -CollectionStatus "WARNING" -ErrorMessage "No disk drive returned."
        return
    }

    foreach ($diskDrive in @($result.Data)) {
        New-DiskDriveInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -DiskDrive $diskDrive
    }
}

function Get-ExchangeServiceHealthSummary {
    param(
        [Parameter(Mandatory)]
        [string]$ServerName
    )

    try {
        $health = Test-ServiceHealth -Server $ServerName -ErrorAction Stop
        $notRunning = @($health.ServicesNotRunning)

        [pscustomobject]@{
            ServerName         = $ServerName
            ServicesRunning    = if ($notRunning.Count -eq 0) { $true } else { $false }
            ServicesNotRunning = ($notRunning -join ";")
            CollectionStatus   = "OK"
            ErrorMessage       = $null
        }
    }
    catch {
        [pscustomobject]@{
            ServerName         = $ServerName
            ServicesRunning    = $null
            ServicesNotRunning = $null
            CollectionStatus   = "ERROR"
            ErrorMessage       = $_.Exception.Message
        }
    }
}

function ConvertTo-ReadinessText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Array]) {
        return ((@($Value) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }) -join "; ")
    }

    return [string]$Value
}

function Get-ObjectPropertyValue {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Add-ExchangeReadinessRow {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Rows,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$Setting,
        [AllowNull()][object]$Value,
        [Parameter(Mandatory)][string]$MigrationFocus,
        [ValidateSet("Info", "Review", "Warning", "Error")][string]$Importance = "Info",
        [ValidateSet("OK", "WARNING", "ERROR")][string]$CollectionStatus = "OK",
        [AllowNull()][string]$ErrorMessage = $null
    )

    $Rows.Add([pscustomobject]@{
        Category         = $Category
        ObjectName       = $ObjectName
        Setting          = $Setting
        Value            = ConvertTo-ReadinessText -Value $Value
        MigrationFocus   = $MigrationFocus
        Importance       = $Importance
        CollectionStatus = $CollectionStatus
        ErrorMessage     = $ErrorMessage
    }) | Out-Null
}

function Add-ExchangeReadinessProperties {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Rows,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string[]]$Properties,
        [Parameter(Mandatory)][string]$MigrationFocus,
        [ValidateSet("Info", "Review", "Warning", "Error")][string]$Importance = "Info"
    )

    foreach ($propertyName in $Properties) {
        Add-ExchangeReadinessRow -Rows $Rows -Category $Category -ObjectName $ObjectName -Setting $propertyName -Value (Get-ObjectPropertyValue -InputObject $InputObject -Name $propertyName) -MigrationFocus $MigrationFocus -Importance $Importance
    }
}

function Invoke-ExchangeReadinessCollector {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Rows,
        [Parameter(Mandatory)][string]$CollectorName,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )

    $collectorStart = Get-Date
    $rowCountBefore = if ($Rows) { $Rows.Count } else { 0 }
    Write-Log ("Starting Exchange readiness collector: {0}" -f $CollectorName)

    try {
        & $ScriptBlock
        $rowCountAfter = if ($Rows) { $Rows.Count } else { 0 }
        $rowsAdded = $rowCountAfter - $rowCountBefore
        $duration = (Get-Date) - $collectorStart
        Write-Log ("Finished Exchange readiness collector: {0}. RowsAdded: {1}. Duration: {2}" -f $CollectorName, $rowsAdded, $duration)
    }
    catch {
        $duration = (Get-Date) - $collectorStart
        Write-Log ("Exchange readiness collector failed: {0}. Duration: {1}. Error: {2}" -f $CollectorName, $duration, $_.Exception.Message) -Level "WARN"
        Add-ExchangeReadinessRow -Rows $Rows -Category "CollectorError" -ObjectName $CollectorName -Setting "Collection" -Value "Failed" -MigrationFocus "Review this collection failure before migration readiness sign-off." -Importance "Warning" -CollectionStatus "ERROR" -ErrorMessage $_.Exception.Message
    }
}

function Get-ExchangeReadinessInventory {
    param(
        [string[]]$ExchangeServerNames = @()
    )

    $rows = New-Object System.Collections.Generic.List[object]

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "Active Directory Exchange schema" -ScriptBlock {
        $rootDse = [ADSI]"LDAP://RootDSE"
        $schemaNamingContext = [string]$rootDse.schemaNamingContext
        $configurationNamingContext = [string]$rootDse.configurationNamingContext
        Add-ExchangeReadinessRow -Rows $rows -Category "ADSchema" -ObjectName "RootDSE" -Setting "SchemaNamingContext" -Value $schemaNamingContext -MigrationFocus "Confirms the schema partition used to validate Exchange schema level before Exchange SE preparation."
        Add-ExchangeReadinessRow -Rows $rows -Category "ADSchema" -ObjectName "RootDSE" -Setting "ConfigurationNamingContext" -Value $configurationNamingContext -MigrationFocus "Confirms the configuration partition containing the Exchange organization."
        $schemaVersionObject = [ADSI]("LDAP://CN=ms-Exch-Schema-Version-Pt,{0}" -f $schemaNamingContext)
        Add-ExchangeReadinessRow -Rows $rows -Category "ADSchema" -ObjectName "ms-Exch-Schema-Version-Pt" -Setting "rangeUpper" -Value $schemaVersionObject.rangeUpper -MigrationFocus "Compare this value with the required Exchange SE setup /PrepareSchema level before migration." -Importance "Review"
        $exchangeServiceContainer = [ADSI]("LDAP://CN=Microsoft Exchange,CN=Services,{0}" -f $configurationNamingContext)
        Add-ExchangeReadinessRow -Rows $rows -Category "ADSchema" -ObjectName "Microsoft Exchange configuration container" -Setting "objectVersion" -Value $exchangeServiceContainer.objectVersion -MigrationFocus "Compare this value with the required Exchange SE /PrepareAD organization object version." -Importance "Review"
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "OrganizationConfig" -ScriptBlock {
        $item = Get-OrganizationConfig -ErrorAction Stop
        Add-ExchangeReadinessProperties -Rows $rows -Category "OrganizationConfig" -InputObject $item -ObjectName $item.Name -Properties @("AdminDisplayName", "OAuth2ClientProfileEnabled", "MapiHttpEnabled", "PublicFoldersEnabled", "ActivityBasedAuthenticationTimeoutEnabled", "DefaultAuthenticationPolicy", "RootPublicFolderMailbox", "DistributionGroupDefaultOU") -MigrationFocus "Core organization settings affecting client auth, coexistence, public folders, and migration behavior." -Importance "Review"
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "AuthConfig" -ScriptBlock {
        if (-not (Get-Command Get-AuthConfig -ErrorAction SilentlyContinue)) { return }
        $item = Get-AuthConfig -ErrorAction Stop
        Add-ExchangeReadinessProperties -Rows $rows -Category "AuthConfig" -InputObject $item -ObjectName "AuthConfig" -Properties @("CurrentCertificateThumbprint", "PreviousCertificateThumbprint", "NextCertificateThumbprint", "ServiceName") -MigrationFocus "Hybrid modern auth and OAuth trust readiness for Exchange Online coexistence." -Importance "Review"
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "TransportConfig" -ScriptBlock {
        $item = Get-TransportConfig -ErrorAction Stop
        Add-ExchangeReadinessProperties -Rows $rows -Category "TransportConfig" -InputObject $item -ObjectName "TransportConfig" -Properties @("MaxSendSize", "MaxReceiveSize", "ExternalPostmasterAddress", "InternalSMTPServers", "TLSReceiveDomainSecureList", "TLSSendDomainSecureList", "JournalingReportNdrTo") -MigrationFocus "Transport limits, TLS posture, journaling NDR behavior, and relay dependencies to review before EXO migration." -Importance "Review"
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "AcceptedDomain" -ScriptBlock {
        foreach ($item in @(Get-AcceptedDomain -ErrorAction Stop | Sort-Object Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "AcceptedDomain" -InputObject $item -ObjectName $item.Name -Properties @("DomainName", "DomainType", "Default", "MatchSubDomains", "AddressBookEnabled") -MigrationFocus "Accepted domain and authoritative/internal relay model must align with Exchange Online accepted domains and mail flow." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "EmailAddressPolicy" -ScriptBlock {
        foreach ($item in @(Get-EmailAddressPolicy -ErrorAction Stop | Sort-Object Priority, Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "EmailAddressPolicy" -InputObject $item -ObjectName $item.Name -Properties @("Priority", "EnabledEmailAddressTemplates", "RecipientFilter", "IncludedRecipients", "ConditionalCompany", "ConditionalDepartment", "ConditionalCustomAttribute1") -MigrationFocus "Proxy address generation and policy ownership must be understood before hybrid and EXO address management changes." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "SendConnector" -ScriptBlock {
        foreach ($item in @(Get-SendConnector -ErrorAction Stop | Sort-Object Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "SendConnector" -InputObject $item -ObjectName $item.Name -Properties @("AddressSpaces", "SmartHosts", "DNSRoutingEnabled", "RequireTLS", "TlsAuthLevel", "SourceTransportServers", "Enabled", "MaxMessageSize") -MigrationFocus "Outbound mail flow, smart hosts, TLS requirements, and source servers to map for coexistence and cutover." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "ReceiveConnector" -ScriptBlock {
        foreach ($item in @(Get-ReceiveConnector -ErrorAction Stop | Sort-Object Identity)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "ReceiveConnector" -InputObject $item -ObjectName $item.Identity -Properties @("Server", "TransportRole", "Bindings", "RemoteIPRanges", "AuthMechanism", "PermissionGroups", "RequireTLS", "Enabled", "MaxMessageSize") -MigrationFocus "Inbound SMTP, application relay, scanner/device relay, and anonymous/authenticated connector dependencies." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "RemoteDomain" -ScriptBlock {
        foreach ($item in @(Get-RemoteDomain -ErrorAction Stop | Sort-Object Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "RemoteDomain" -InputObject $item -ObjectName $item.Name -Properties @("DomainName", "AllowedOOFType", "AutoReplyEnabled", "AutoForwardEnabled", "TNEFEnabled", "TrustedMailOutboundEnabled") -MigrationFocus "Remote domain behavior can affect coexistence, forwarding, OOF, and external recipient experience after migration." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "MailboxDatabase" -ScriptBlock {
        foreach ($item in @(Get-MailboxDatabase -Status -ErrorAction Stop | Sort-Object Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "MailboxDatabase" -InputObject $item -ObjectName $item.Name -Properties @("Server", "Mounted", "DatabaseSize", "AvailableNewMailboxSpace", "EdbFilePath", "LogFolderPath", "CircularLoggingEnabled", "Recovery", "ReplicationType", "IsExcludedFromProvisioning", "IsSuspendedFromProvisioning") -MigrationFocus "Mailbox placement, storage, replication, and provisioning controls to plan moves and decommissioning." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "DatabaseAvailabilityGroup" -ScriptBlock {
        if (-not (Get-Command Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue)) { return }
        foreach ($item in @(Get-DatabaseAvailabilityGroup -Status -ErrorAction Stop | Sort-Object Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "DatabaseAvailabilityGroup" -InputObject $item -ObjectName $item.Name -Properties @("Servers", "WitnessServer", "WitnessDirectory", "DatacenterActivationMode", "DatabaseAvailabilityGroupIpAddresses", "ThirdPartyReplication", "OperationalServers") -MigrationFocus "DAG layout and witness settings are required for SE transition planning and decommission order." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "OfflineAddressBook" -ScriptBlock {
        foreach ($item in @(Get-OfflineAddressBook -ErrorAction Stop | Sort-Object Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "OfflineAddressBook" -InputObject $item -ObjectName $item.Name -Properties @("IsDefault", "GlobalWebDistributionEnabled", "VirtualDirectories", "AddressLists", "GeneratingMailbox") -MigrationFocus "OAB generation and distribution must be validated for Outlook during coexistence and mailbox moves." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "VirtualDirectories" -ScriptBlock {
        $virtualDirectoryCmdlets = @(
            @{ Name = "OWA"; Cmdlet = "Get-OwaVirtualDirectory" },
            @{ Name = "ECP"; Cmdlet = "Get-EcpVirtualDirectory" },
            @{ Name = "EWS"; Cmdlet = "Get-WebServicesVirtualDirectory" },
            @{ Name = "MAPI"; Cmdlet = "Get-MapiVirtualDirectory" },
            @{ Name = "OAB"; Cmdlet = "Get-OabVirtualDirectory" },
            @{ Name = "ActiveSync"; Cmdlet = "Get-ActiveSyncVirtualDirectory" },
            @{ Name = "PowerShell"; Cmdlet = "Get-PowerShellVirtualDirectory" },
            @{ Name = "Autodiscover"; Cmdlet = "Get-AutodiscoverVirtualDirectory" }
        )
        foreach ($definition in $virtualDirectoryCmdlets) {
            if (-not (Get-Command $definition.Cmdlet -ErrorAction SilentlyContinue)) {
                Write-Log ("Skipping virtual directory collector because cmdlet is unavailable: {0}" -f $definition.Cmdlet)
                continue
            }

            Write-Log ("Collecting virtual directories with {0}" -f $definition.Cmdlet)
            foreach ($item in @(& $definition.Cmdlet -ErrorAction Stop | Sort-Object Identity)) {
                Add-ExchangeReadinessProperties -Rows $rows -Category ("VirtualDirectory:{0}" -f $definition.Name) -InputObject $item -ObjectName $item.Identity -Properties @("Server", "InternalUrl", "ExternalUrl", "InternalAuthenticationMethods", "ExternalAuthenticationMethods", "BasicAuthentication", "WindowsAuthentication", "OAuthAuthentication", "ClientCertAuth") -MigrationFocus "Namespace, URL and authentication inventory for Exchange SE coexistence, load balancer, certificate, and client access planning." -Importance "Review"
            }
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "ExchangeCertificate" -ScriptBlock {
        $certificateServers = @($ExchangeServerNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
        if ($certificateServers.Count -eq 0) {
            $certificateServers = @($env:COMPUTERNAME)
        }

        foreach ($serverName in $certificateServers) {
            Write-Log ("Collecting Exchange certificates from {0}" -f $serverName)
            try {
                $serverCertificateCount = 0
                foreach ($item in @(Get-ExchangeCertificate -Server $serverName -ErrorAction Stop | Sort-Object NotAfter)) {
                    $serverCertificateCount++
                    $objectName = "{0}:{1}" -f $serverName, $item.Thumbprint
                    Add-ExchangeReadinessRow -Rows $rows -Category "ExchangeCertificate" -ObjectName $objectName -Setting "Server" -Value $serverName -MigrationFocus "Certificate coverage and expiration are critical for HTTPS, SMTP TLS, hybrid, and Exchange SE namespace readiness." -Importance "Review"
                    Add-ExchangeReadinessProperties -Rows $rows -Category "ExchangeCertificate" -InputObject $item -ObjectName $objectName -Properties @("Subject", "CertificateDomains", "Services", "NotBefore", "NotAfter", "IsSelfSigned", "Status") -MigrationFocus "Certificate coverage and expiration are critical for HTTPS, SMTP TLS, hybrid, and Exchange SE namespace readiness." -Importance "Review"
                }
                Write-Log ("Finished Exchange certificates from {0}. Certificates: {1}" -f $serverName, $serverCertificateCount)
            }
            catch {
                Write-Log ("Exchange certificate collection failed for {0}: {1}" -f $serverName, $_.Exception.Message) -Level "WARN"
                Add-ExchangeReadinessRow -Rows $rows -Category "ExchangeCertificate" -ObjectName $serverName -Setting "Collection" -Value "Failed" -MigrationFocus "Review certificate collection failure for this server before namespace, TLS, hybrid, or SE readiness sign-off." -Importance "Warning" -CollectionStatus "ERROR" -ErrorMessage $_.Exception.Message
            }
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "HybridConfiguration" -ScriptBlock {
        if (-not (Get-Command Get-HybridConfiguration -ErrorAction SilentlyContinue)) { return }
        foreach ($item in @(Get-HybridConfiguration -ErrorAction Stop)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "HybridConfiguration" -InputObject $item -ObjectName $item.Name -Properties @("ClientAccessServers", "ReceivingTransportServers", "SendingTransportServers", "Domains", "Features", "TlsCertificateName", "OnPremisesSmartHost", "ExternalIPAddresses") -MigrationFocus "Hybrid configuration drives Exchange Online coexistence, secure mail flow, free/busy, and migration endpoint planning." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "OrganizationRelationship" -ScriptBlock {
        if (-not (Get-Command Get-OrganizationRelationship -ErrorAction SilentlyContinue)) { return }
        foreach ($item in @(Get-OrganizationRelationship -ErrorAction Stop | Sort-Object Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "OrganizationRelationship" -InputObject $item -ObjectName $item.Name -Properties @("DomainNames", "FreeBusyAccessEnabled", "FreeBusyAccessLevel", "TargetAutodiscoverEpr", "TargetSharingEpr", "Enabled") -MigrationFocus "Free/busy and organization sharing dependencies for Exchange Online coexistence." -Importance "Review"
        }
    }

    Invoke-ExchangeReadinessCollector -Rows $rows -CollectorName "IntraOrganizationConnector" -ScriptBlock {
        if (-not (Get-Command Get-IntraOrganizationConnector -ErrorAction SilentlyContinue)) { return }
        foreach ($item in @(Get-IntraOrganizationConnector -ErrorAction Stop | Sort-Object Name)) {
            Add-ExchangeReadinessProperties -Rows $rows -Category "IntraOrganizationConnector" -InputObject $item -ObjectName $item.Name -Properties @("DiscoveryEndpoint", "TargetAddressDomains", "Enabled") -MigrationFocus "Hybrid intra-organization connectivity for Exchange Online coexistence." -Importance "Review"
        }
    }

    return $rows.ToArray()
}
function Get-MailboxDatabasePathInventory {
    param(
        [Parameter(Mandatory)]
        [string[]]$ExchangeServerNames
    )

    $serverSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($serverName in $ExchangeServerNames) {
        [void]$serverSet.Add($serverName)
    }

    try {
        $databases = Get-MailboxDatabase -Status -ErrorAction Stop

        foreach ($database in $databases) {
            $serverName = [string]$database.Server
            if (-not $serverSet.Contains($serverName)) {
                continue
            }

            [pscustomobject]@{
                ServerName               = $serverName
                DatabaseName             = $database.Name
                Mounted                  = $database.Mounted
                DatabaseSize             = if ($database.DatabaseSize) { $database.DatabaseSize.ToString() } else { $null }
                AvailableNewMailboxSpace = if ($database.AvailableNewMailboxSpace) { $database.AvailableNewMailboxSpace.ToString() } else { $null }
                EdbFilePath              = if ($database.EdbFilePath) { $database.EdbFilePath.PathName } else { $null }
                LogFolderPath            = if ($database.LogFolderPath) { $database.LogFolderPath.PathName } else { $null }
                CircularLoggingEnabled   = $database.CircularLoggingEnabled
                Recovery                 = $database.Recovery
                ReplicationType          = $database.ReplicationType
                ErrorMessage             = $null
            }
        }
    }
    catch {
        [pscustomobject]@{
            ServerName               = $null
            DatabaseName             = $null
            Mounted                  = $null
            DatabaseSize             = $null
            AvailableNewMailboxSpace = $null
            EdbFilePath              = $null
            LogFolderPath            = $null
            CircularLoggingEnabled   = $null
            Recovery                 = $null
            ReplicationType          = $null
            ErrorMessage             = $_.Exception.Message
        }
    }
}

function New-HtmlExecutiveSummary {
    param(
        [Parameter(Mandatory)]
        [object]$Summary,

        [Parameter(Mandatory)]
        [object[]]$PerServerSummary,

        [AllowNull()]
        [object[]]$ReadinessInventory = @(),

        [Parameter(Mandatory)]
        [string]$Path
    )

    $rowsHtml = foreach ($row in ($PerServerSummary | Sort-Object ExchangeServerName)) {
        "<tr><td>$(Format-HtmlValue $row.ExchangeServerName)</td><td>$(Format-HtmlValue $row.ServerRole)</td><td class='num'>$(Format-HtmlValue $row.LogicalProcessorCount)</td><td class='num'>$(Format-HtmlValue $row.MemoryGB)</td><td class='num'>$(Format-HtmlValue $row.DiskDriveCount)</td><td class='num'>$(Format-HtmlValue $row.DiskDriveTotalSizeGB)</td><td class='status'>$(Format-HtmlValue $row.ComputeCollectionStatus)</td></tr>"
    }

    $readinessRows = @($ReadinessInventory)
    $readinessErrorCount = @($readinessRows | Where-Object { $_.CollectionStatus -eq "ERROR" -or $_.Importance -eq "Error" }).Count
    $readinessWarningCount = @($readinessRows | Where-Object { $_.CollectionStatus -eq "WARNING" -or $_.Importance -eq "Warning" }).Count
    $readinessSortedRows = @($readinessRows | Sort-Object @{ Expression = { if ($_.CollectionStatus -eq "ERROR" -or $_.Importance -eq "Error") { 0 } elseif ($_.CollectionStatus -eq "WARNING" -or $_.Importance -eq "Warning") { 1 } elseif ($_.Importance -eq "Review") { 2 } else { 3 } } }, Category, ObjectName, Setting)
    $readinessCategoryBlocks = foreach ($categoryGroup in ($readinessSortedRows | Group-Object Category | Sort-Object Name)) {
        $categoryRowsHtml = foreach ($row in $categoryGroup.Group) {
            $statusClass = switch ([string]$row.CollectionStatus) {
                "ERROR" { "status-error"; break }
                "WARNING" { "status-warning"; break }
                default { "status-ok" }
            }
            $importanceClass = switch ([string]$row.Importance) {
                "Error" { "status-error"; break }
                "Warning" { "status-warning"; break }
                "Review" { "status-review"; break }
                default { "status-ok" }
            }
            "<tr><td>$(Format-HtmlValue $row.ObjectName)</td><td>$(Format-HtmlValue $row.Setting)</td><td class='cell-value'>$(Format-HtmlValue $row.Value)</td><td><span class='badge $importanceClass'>$(Format-HtmlValue $row.Importance)</span></td><td><span class='badge $statusClass'>$(Format-HtmlValue $row.CollectionStatus)</span></td><td>$(Format-HtmlValue $row.MigrationFocus)</td><td>$(Format-HtmlValue $row.ErrorMessage)</td></tr>"
        }

@"
      <h3>$(Format-HtmlValue $categoryGroup.Name) <span>$(Format-HtmlValue $categoryGroup.Count) records</span></h3>
      <table class="readiness-table">
        <thead>
          <tr>
            <th>Object</th>
            <th>Setting</th>
            <th>Value</th>
            <th>Importance</th>
            <th>Status</th>
            <th>Migration focus</th>
            <th>Error</th>
          </tr>
        </thead>
        <tbody>
          $($categoryRowsHtml -join "`r`n")
        </tbody>
      </table>
"@
    }

    $readinessSection = if ($readinessRows.Count -gt 0) {
@"
      <h2>Exchange migration readiness configuration</h2>
      <div class="summary">
        Collected <strong>$(Format-HtmlValue $readinessRows.Count)</strong> Exchange configuration readiness records for Exchange SE and Exchange Online planning. AD schema rangeUpper: <strong>$(Format-HtmlValue $Summary.ExchangeSchemaRangeUpper)</strong>. Exchange organization objectVersion: <strong>$(Format-HtmlValue $Summary.ExchangeOrgObjectVersion)</strong>. Review warnings: <strong>$(Format-HtmlValue $readinessWarningCount)</strong>. Collection errors: <strong>$(Format-HtmlValue $readinessErrorCount)</strong>. Every collected readiness record is rendered below and also attached as CSV.
      </div>
      $($readinessCategoryBlocks -join "`r`n")
"@
    } else {
        ""
    }
    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Exchange On-Premises Infrastructure and Migration Readiness</title>
<style>
body { margin: 0; padding: 0; background: #f6f8fb; color: #0f172a; font-family: Segoe UI, Arial, sans-serif; }
.shell { max-width: 1180px; margin: 0 auto; padding: 24px; }
.panel { background: #ffffff; border: 1px solid #dbe3ef; border-radius: 10px; overflow: hidden; }
.hero { background: linear-gradient(135deg, #0f766e, #2563eb); color: #ffffff; padding: 24px 28px; }
.eyebrow { font-size: 12px; text-transform: uppercase; letter-spacing: .08em; opacity: .9; font-weight: 700; }
h1 { font-size: 26px; line-height: 1.25; margin: 6px 0 0 0; color: #ffffff; }
.subtitle { margin-top: 10px; font-size: 13px; opacity: .95; }
.content { padding: 24px 28px; }
.cards { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px; margin-bottom: 22px; }
.card { border: 1px solid #dbe3ef; border-radius: 8px; padding: 14px; background: #f8fafc; }
.card.green { border-color: #bbf7d0; background: #f0fdf4; }
.card.blue { border-color: #bfdbfe; background: #eff6ff; }
.card.amber { border-color: #fed7aa; background: #fff7ed; }
.card.purple { border-color: #e9d5ff; background: #faf5ff; }
.card .label { font-size: 11px; text-transform: uppercase; color: #64748b; font-weight: 800; }
.card .value { font-size: 28px; font-weight: 800; margin-top: 4px; color: #0f172a; }
.card.green .value { color: #166534; }
.card.blue .value { color: #1d4ed8; }
.card.amber .value { color: #9a3412; }
.card.purple .value { color: #7e22ce; }
.card .unit { font-size: 12px; color: #64748b; margin-top: 2px; }
.summary { border: 1px solid #e5e7eb; border-radius: 8px; background: #ffffff; padding: 16px; font-size: 14px; line-height: 1.55; color: #334155; margin-bottom: 22px; }
h2 { font-size: 17px; margin: 0 0 10px 0; color: #0f172a; }
h3 { font-size: 14px; margin: 22px 0 8px 0; color: #0f172a; }
h3 span { color: #64748b; font-size: 12px; font-weight: 600; margin-left: 6px; }
table { width: 100%; border-collapse: collapse; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; font-size: 12px; }
th { text-align: left; background: #f8fafc; color: #475569; padding: 10px 12px; text-transform: uppercase; font-size: 11px; }
td { border-bottom: 1px solid #e5e7eb; padding: 9px 12px; color: #334155; }
td.num, th.num { text-align: right; }
.readiness-table { margin-bottom: 18px; table-layout: fixed; }
.readiness-table th:nth-child(1) { width: 16%; }
.readiness-table th:nth-child(2) { width: 14%; }
.readiness-table th:nth-child(3) { width: 22%; }
.readiness-table th:nth-child(4) { width: 8%; }
.readiness-table th:nth-child(5) { width: 8%; }
.readiness-table th:nth-child(6) { width: 22%; }
.readiness-table th:nth-child(7) { width: 10%; }
.readiness-table td { vertical-align: top; overflow-wrap: anywhere; }
.cell-value { font-family: Consolas, Courier New, monospace; font-size: 11px; }
.badge { display: inline-block; border-radius: 999px; padding: 2px 7px; font-size: 10px; font-weight: 800; text-transform: uppercase; }
.status-ok { color: #166534; background: #dcfce7; }
.status-review { color: #1d4ed8; background: #dbeafe; }
.status-warning { color: #9a3412; background: #ffedd5; }
.status-error { color: #991b1b; background: #fee2e2; }
.status { font-weight: 700; color: #047857; }
.footer { margin-top: 18px; padding-top: 14px; border-top: 1px solid #e5e7eb; font-size: 12px; color: #64748b; line-height: 1.5; }
@media (max-width: 900px) { .cards { grid-template-columns: repeat(2, minmax(0, 1fr)); } .shell { padding: 12px; } }
</style>
</head>
<body>
<div class="shell">
  <div class="panel">
    <div class="hero">
      <div class="eyebrow">Smart365 Exchange OnPrem</div>
      <h1>Exchange On-Premises Infrastructure and Migration Readiness</h1>
      <div class="subtitle">Executive summary generated on $(Format-HtmlValue $Summary.ExecutionDate) - RunId $(Format-HtmlValue $Summary.RunId)</div>
    </div>

    <div class="content">
      <div class="cards">
        <div class="card"><div class="label">Exchange VMs</div><div class="value">$(Format-HtmlValue $Summary.ExchangeServersCount)</div><div class="unit">servers identified</div></div>
        <div class="card blue"><div class="label">vCPU</div><div class="value">$(Format-HtmlValue $Summary.TotalLogicalProcessorCount)</div><div class="unit">logical processors</div></div>
        <div class="card green"><div class="label">RAM</div><div class="value">$(Format-HtmlValue $Summary.TotalMemoryGB)</div><div class="unit">GB</div></div>
        <div class="card amber"><div class="label">Disks</div><div class="value">$(Format-HtmlValue $Summary.TotalDiskDriveCount)</div><div class="unit">WMI disk drives</div></div>
        <div class="card purple"><div class="label">Provisioned Storage</div><div class="value">$(Format-HtmlValue $Summary.TotalDiskDriveSizeTB)</div><div class="unit">TB</div></div>
      </div>

      <div class="summary">
        The current Exchange on-premises footprint represents <strong>$(Format-HtmlValue $Summary.ExchangeServersCount) virtual machines</strong>, <strong>$(Format-HtmlValue $Summary.TotalLogicalProcessorCount) vCPU</strong>, <strong>$(Format-HtmlValue $Summary.TotalMemoryGB) GB RAM</strong>, and <strong>$(Format-HtmlValue $Summary.TotalDiskDriveCount) disks</strong>. The full readiness section below supports Exchange SE preparation, Exchange Online coexistence, dependency review, and decommissioning sign-off.
      </div>

      <h2>Per-server infrastructure summary</h2>
      <table>
        <thead>
          <tr>
            <th>Server</th>
            <th>Role</th>
            <th class="num">vCPU</th>
            <th class="num">RAM GB</th>
            <th class="num">Disks</th>
            <th class="num">Disk Size GB</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          $($rowsHtml -join "`r`n")
        </tbody>
      </table>

      $readinessSection

      <div class="footer">
        CPU, RAM and disk data are collected from the guest OS using WMI/DCOM only. Validate with Infrastructure if hypervisor-level figures are required for final capacity reclamation.
      </div>
    </div>
  </div>
</div>
</body>
</html>
"@

    Set-Content -Path $Path -Value $html -Encoding UTF8
}

$StartTime = Get-Date

try {
    Ensure-Directory -Path $OutputFolder
    Ensure-Directory -Path $LogFolder
    Start-Transcript -Path $TranscriptFile -Append | Out-Null
    $script:TranscriptStarted = $true
    Write-Log "Starting $ScriptName v$ScriptVersion. RunId: $RunId"
    Write-Log "Output folder: $OutputFolder"
    Write-Log "Collection method: WMI/DCOM only"

    Test-ExchangeShell

    Write-Log "Collecting Exchange servers."
    $exchangeServers = @(Get-ExchangeServer | Sort-Object Name)

    if ($exchangeServers.Count -eq 0) {
        throw "No Exchange servers were returned by Get-ExchangeServer."
    }

    $serverInventory = foreach ($server in $exchangeServers) {
        [pscustomobject]@{
            Name                = $server.Name
            Fqdn                = $server.Fqdn
            Site                = if ($server.Site) { $server.Site.ToString() } else { $null }
            ServerRole          = $server.ServerRole
            AdminDisplayVersion = $server.AdminDisplayVersion.ToString()
            Edition             = $server.Edition
            IsExchange2016      = if ($server.AdminDisplayVersion.ToString() -like "Version 15.1*") { $true } else { $false }
        }
    }

    $serverInventoryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_Inventory.csv"
    Export-ServersAndStorageCsv -Data @($serverInventory) -Path $serverInventoryPath
    Write-Log "Exchange server inventory exported to: $serverInventoryPath"

    Write-Log "Testing remote access with WMI/DCOM."
    $remoteAccessInventory = foreach ($server in $exchangeServers) {
        $namesToTry = Get-ComputerNamesToTry -ExchangeServer $server
        Write-Log "Testing WMI/DCOM access for $($server.Name)"
        Test-RemoteAccess -ExchangeServerName $server.Name -ComputerNamesToTry $namesToTry
    }

    $remoteAccessPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_RemoteAccess.csv"
    Export-ServersAndStorageCsv -Data @($remoteAccessInventory) -Path $remoteAccessPath
    Write-Log "Remote access report exported to: $remoteAccessPath"

    Write-Log "Collecting CPU and RAM inventory."
    $computeInventory = foreach ($server in $exchangeServers) {
        $namesToTry = Get-ComputerNamesToTry -ExchangeServer $server
        Write-Log "Collecting compute inventory from $($server.Name)"
        Get-ServerComputeInventory -ExchangeServerName $server.Name -ComputerNamesToTry $namesToTry
    }

    $computeInventoryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_Compute.csv"
    Export-ServersAndStorageCsv -Data @($computeInventory) -Path $computeInventoryPath
    Write-Log "Compute inventory exported to: $computeInventoryPath"

    Write-Log "Collecting logical disk inventory."
    $logicalDiskInventory = foreach ($server in $exchangeServers) {
        $namesToTry = Get-ComputerNamesToTry -ExchangeServer $server
        Write-Log "Collecting logical disks from $($server.Name)"
        Get-ServerLogicalDiskInventory -ExchangeServerName $server.Name -ComputerNamesToTry $namesToTry
    }

    $logicalDiskInventoryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_LogicalDisks.csv"
    Export-ServersAndStorageCsv -Data @($logicalDiskInventory) -Path $logicalDiskInventoryPath
    Write-Log "Logical disk inventory exported to: $logicalDiskInventoryPath"

    Write-Log "Collecting disk drive inventory."
    $diskDriveInventory = foreach ($server in $exchangeServers) {
        $namesToTry = Get-ComputerNamesToTry -ExchangeServer $server
        Write-Log "Collecting disk drives from $($server.Name)"
        Get-ServerDiskDriveInventory -ExchangeServerName $server.Name -ComputerNamesToTry $namesToTry
    }

    $diskDriveInventoryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_DiskDrives.csv"
    Export-ServersAndStorageCsv -Data @($diskDriveInventory) -Path $diskDriveInventoryPath
    Write-Log "Disk drive inventory exported to: $diskDriveInventoryPath"

    if ($IncludeServicesHealth) {
        Write-Log "Collecting Exchange services health."
        $serviceHealth = foreach ($server in $exchangeServers) {
            Write-Log "Collecting service health from $($server.Name)"
            Get-ExchangeServiceHealthSummary -ServerName $server.Name
        }

        $serviceHealthPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_ServiceHealth.csv"
        Export-ServersAndStorageCsv -Data @($serviceHealth) -Path $serviceHealthPath
        Write-Log "Service health inventory exported to: $serviceHealthPath"
    }

    if ($IncludeMailboxDatabasePaths) {
        Write-Log "Collecting mailbox database paths."
        $databasePaths = @(Get-MailboxDatabasePathInventory -ExchangeServerNames @($exchangeServers.Name))
        $databasePathsPath = Join-Path $OutputFolder "Exchange_OnPrem_MailboxDatabases_Paths.csv"
        Export-ServersAndStorageCsv -Data @($databasePaths) -Path $databasePathsPath
        Write-Log "Mailbox database path inventory exported to: $databasePathsPath"
    }

    Write-Log "Collecting Exchange schema and migration readiness configuration."
    $exchangeReadinessInventory = @(Get-ExchangeReadinessInventory -ExchangeServerNames @($exchangeServers.Name))
    $exchangeReadinessPath = Join-Path $OutputFolder "Exchange_OnPrem_MigrationReadiness_Config.csv"
    Export-ServersAndStorageCsv -Data @($exchangeReadinessInventory) -Path $exchangeReadinessPath
    Write-Log "Exchange schema and migration readiness configuration exported to: $exchangeReadinessPath"

    $logicalDiskRows = @($logicalDiskInventory)
    $successfulLogicalDiskRows = @($logicalDiskRows | Where-Object { $_.CollectionStatus -eq "OK" })
    $failedLogicalDiskRows = @($logicalDiskRows | Where-Object { $_.CollectionStatus -eq "ERROR" })
    $lowSpaceRows = @($successfulLogicalDiskRows | Where-Object { $_.LowSpaceWarning -eq $true })

    $diskDriveRows = @($diskDriveInventory)
    $successfulDiskDriveRows = @($diskDriveRows | Where-Object { $_.CollectionStatus -eq "OK" })
    $failedDiskDriveRows = @($diskDriveRows | Where-Object { $_.CollectionStatus -eq "ERROR" })

    $computeRows = @($computeInventory)
    $successfulComputeRows = @($computeRows | Where-Object { $_.CollectionStatus -eq "OK" })
    $failedComputeRows = @($computeRows | Where-Object { $_.CollectionStatus -eq "ERROR" })

    Write-Log "Building per-server infrastructure summary."
    $perServerSummary = foreach ($server in $exchangeServers) {
        $serverName = $server.Name
        $compute = $computeRows | Where-Object { $_.ExchangeServerName -eq $serverName } | Select-Object -First 1
        $serverLogicalDisks = @($successfulLogicalDiskRows | Where-Object { $_.ExchangeServerName -eq $serverName })
        $serverDiskDrives = @($successfulDiskDriveRows | Where-Object { $_.ExchangeServerName -eq $serverName })

        [pscustomobject]@{
            ExchangeServerName          = $serverName
            Fqdn                        = $server.Fqdn
            ServerRole                  = $server.ServerRole
            IsExchange2016              = if ($server.AdminDisplayVersion.ToString() -like "Version 15.1*") { $true } else { $false }
            IsVirtualMachine            = if ($compute) { $compute.IsVirtualMachine } else { $null }
            Manufacturer                = if ($compute) { $compute.Manufacturer } else { $null }
            Model                       = if ($compute) { $compute.Model } else { $null }
            SocketCount                 = if ($compute) { $compute.SocketCount } else { $null }
            PhysicalCoreCount           = if ($compute) { $compute.PhysicalCoreCount } else { $null }
            LogicalProcessorCount       = if ($compute) { $compute.LogicalProcessorCount } else { $null }
            MemoryGB                    = if ($compute) { $compute.MemoryGB } else { $null }
            DiskDriveCount              = $serverDiskDrives.Count
            DiskDriveTotalSizeGB        = Get-SafeSum -InputObject $serverDiskDrives -Property "SizeGB"
            LogicalDiskCount            = $serverLogicalDisks.Count
            LogicalDiskTotalSizeGB      = Get-SafeSum -InputObject $serverLogicalDisks -Property "SizeGB"
            LogicalDiskUsedGB           = Get-SafeSum -InputObject $serverLogicalDisks -Property "UsedGB"
            LogicalDiskFreeGB           = Get-SafeSum -InputObject $serverLogicalDisks -Property "FreeGB"
            LowSpaceLogicalDiskCount    = @($serverLogicalDisks | Where-Object { $_.LowSpaceWarning -eq $true }).Count
            ComputeCollectionStatus     = if ($compute) { $compute.CollectionStatus } else { "ERROR" }
            DiskDriveCollectionStatus   = if (@($diskDriveRows | Where-Object { $_.ExchangeServerName -eq $serverName -and $_.CollectionStatus -eq "ERROR" }).Count -gt 0) { "ERROR" } else { "OK" }
            LogicalDiskCollectionStatus = if (@($logicalDiskRows | Where-Object { $_.ExchangeServerName -eq $serverName -and $_.CollectionStatus -eq "ERROR" }).Count -gt 0) { "ERROR" } else { "OK" }
        }
    }

    $perServerSummaryPath = Join-Path $OutputFolder "Exchange_OnPrem_Infrastructure_PerServerSummary.csv"
    Export-ServersAndStorageCsv -Data @($perServerSummary) -Path $perServerSummaryPath
    Write-Log "Per-server decommissioning summary exported to: $perServerSummaryPath"

    $totalMemoryGB = Get-SafeSum -InputObject $successfulComputeRows -Property "MemoryGB"
    $totalDiskDriveSizeGB = Get-SafeSum -InputObject $successfulDiskDriveRows -Property "SizeGB"
    $totalLogicalDiskSizeGB = Get-SafeSum -InputObject $successfulLogicalDiskRows -Property "SizeGB"
    $totalLogicalDiskUsedGB = Get-SafeSum -InputObject $successfulLogicalDiskRows -Property "UsedGB"
    $totalLogicalDiskFreeGB = Get-SafeSum -InputObject $successfulLogicalDiskRows -Property "FreeGB"

    $exchangeReadinessRows = @($exchangeReadinessInventory)
    $exchangeReadinessErrorRows = @($exchangeReadinessRows | Where-Object { $_.CollectionStatus -eq "ERROR" -or $_.Importance -eq "Error" })
    $exchangeReadinessWarningRows = @($exchangeReadinessRows | Where-Object { $_.CollectionStatus -eq "WARNING" -or $_.Importance -eq "Warning" })
    $exchangeSchemaRangeUpperRow = $exchangeReadinessRows | Where-Object { $_.Category -eq "ADSchema" -and $_.Setting -eq "rangeUpper" } | Select-Object -First 1
    $exchangeOrgObjectVersionRow = $exchangeReadinessRows | Where-Object { $_.Category -eq "ADSchema" -and $_.Setting -eq "objectVersion" } | Select-Object -First 1
    $exchangeSchemaRangeUpper = if ($null -ne $exchangeSchemaRangeUpperRow) { $exchangeSchemaRangeUpperRow.Value } else { $null }
    $exchangeOrgObjectVersion = if ($null -ne $exchangeOrgObjectVersionRow) { $exchangeOrgObjectVersionRow.Value } else { $null }

    $summary = [pscustomobject]@{
        ScriptName                       = $ScriptName
        ScriptVersion                    = $ScriptVersion
        RunId                            = $RunId
        ExecutionDate                    = Get-Date
        OutputFolder                     = $OutputFolder
        ExchangeServersCount             = $exchangeServers.Count
        ComputeRowsCount                 = $computeRows.Count
        ComputeCollectionErrors          = $failedComputeRows.Count
        TotalSocketCount                 = Get-SafeSum -InputObject $successfulComputeRows -Property "SocketCount"
        TotalPhysicalCoreCount           = Get-SafeSum -InputObject $successfulComputeRows -Property "PhysicalCoreCount"
        TotalLogicalProcessorCount       = Get-SafeSum -InputObject $successfulComputeRows -Property "LogicalProcessorCount"
        TotalMemoryGB                    = $totalMemoryGB
        TotalMemoryTB                    = ConvertTo-TBFromGB -GB $totalMemoryGB
        DiskDriveRowsCount               = $diskDriveRows.Count
        SuccessfulDiskDriveRowsCount     = $successfulDiskDriveRows.Count
        DiskDriveCollectionErrors        = $failedDiskDriveRows.Count
        TotalDiskDriveCount              = $successfulDiskDriveRows.Count
        TotalDiskDriveSizeGB             = $totalDiskDriveSizeGB
        TotalDiskDriveSizeTB             = ConvertTo-TBFromGB -GB $totalDiskDriveSizeGB
        LogicalDiskRowsCount             = $logicalDiskRows.Count
        SuccessfulLogicalDiskRowsCount   = $successfulLogicalDiskRows.Count
        LogicalDiskCollectionErrors      = $failedLogicalDiskRows.Count
        LowSpaceWarnings                 = $lowSpaceRows.Count
        TotalLogicalDiskSizeGB           = $totalLogicalDiskSizeGB
        TotalLogicalDiskSizeTB           = ConvertTo-TBFromGB -GB $totalLogicalDiskSizeGB
        TotalLogicalDiskUsedGB           = $totalLogicalDiskUsedGB
        TotalLogicalDiskUsedTB           = ConvertTo-TBFromGB -GB $totalLogicalDiskUsedGB
        TotalLogicalDiskFreeGB           = $totalLogicalDiskFreeGB
        TotalLogicalDiskFreeTB           = ConvertTo-TBFromGB -GB $totalLogicalDiskFreeGB
        ExchangeSchemaRangeUpper         = $exchangeSchemaRangeUpper
        ExchangeOrgObjectVersion         = $exchangeOrgObjectVersion
        ExchangeReadinessRowsCount       = $exchangeReadinessRows.Count
        ExchangeReadinessWarnings        = $exchangeReadinessWarningRows.Count
        ExchangeReadinessErrors          = $exchangeReadinessErrorRows.Count
    }

    $summaryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_Inventory_Summary.csv"
    Export-ServersAndStorageCsv -Data @($summary) -Path $summaryPath
    Write-Log "Global summary exported to: $summaryPath"

    $htmlSummaryPath = Join-Path $OutputFolder "Exchange_OnPrem_InfrastructureAndReadiness_Report.html"
    New-HtmlExecutiveSummary -Summary $summary -PerServerSummary $perServerSummary -ReadinessInventory $exchangeReadinessInventory -Path $htmlSummaryPath
    Write-Log "HTML executive summary exported to: $htmlSummaryPath"
    Invoke-ServersAndStorageSharePointUpload -LocalFilePath $htmlSummaryPath
    Save-ServersAndStorageWeeklyHistory

    $mailSubject = "SmartM365 Exchange OnPrem Servers and Storage inventory - $RunId"
    Send-ServersAndStorageHtmlReport `
        -HtmlReportPath $htmlSummaryPath `
        -Subject $mailSubject `
        -Attachments @($htmlSummaryPath, $summaryPath, $perServerSummaryPath, $exchangeReadinessPath) `
        -Summary $summary `
        -PerServerSummary @($perServerSummary) `
        -ReadinessInventory @($exchangeReadinessInventory)
    Write-Log "HTML executive summary email sent."

    Write-Log "Completed successfully."
    Complete-ServersAndStorageRun -Status 'Success' -Started $StartTime -ErrorMessage $null
}
catch {
    $message = $_.Exception.Message
    try { Write-Log -Level "ERROR" -Message $message } catch {}
    try { Complete-ServersAndStorageRun -Status 'Failed' -Started $StartTime -ErrorMessage $message } catch {}
    throw
}
