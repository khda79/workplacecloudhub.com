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
    - Optional Microsoft CSS-Exchange Health Checker findings and per-server summary
      (vendored, SHA256 and Microsoft signature verified before every run). The collector server runs
      it locally; every other server runs it locally through its own scheduled task, started on
      demand over CIM/DCOM (see Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1)
    - Exchange schema and readiness configuration
    - Per-server decommissioning summary
    - Global CSV summary
    - HTML executive summary for slide integration

    This version does not use:
    - Get-CimInstance
    - Invoke-Command
    - WinRM / PowerShell Remoting
    New-CimSession is used over DCOM only, to start the Health Checker scheduled task.

.VERSION
1.8.4
.REQUIREMENTS
    Windows PowerShell 5.1 on an Exchange 2016/on-premises management host.
    Modules/snap-ins: SmartM365 WindowsPowerShell5 compatibility module; Exchange Management snap-in.
    Minimum permissions: Exchange on-premises read access for Get-ExchangeServer and virtual directory/readiness cmdlets; WMI/DCOM read access to target servers for compute, disk and storage inventory.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
    Conditional: -IncludeHealthChecker requires an elevated session (local administrator) on the Exchange server;
    for the other servers, the run account must be local administrator of each of them (C$ share, CIM/DCOM)
    and the Health Checker scheduled task must be installed there.
.NOTES
    Script Name : SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1
    Version     : 1.8.4
    Requirements:
      - Windows PowerShell 5.1 with Exchange 2016 Management Tools
      - Exchange 2016 read RBAC for Get-ExchangeServer, Get-MailboxDatabase,
        virtual directory, certificate, service health and readiness Get-* cmdlets
      - Active Directory / Exchange organization read access for schema and organization readiness checks
      - Remote WMI/DCOM read access to Exchange servers for Win32_Processor,
        Win32_ComputerSystem, Win32_LogicalDisk and Win32_DiskDrive
      - PowerShell 5.1 or later

.CHANGELOG
    1.8.4
      - An interrupted run (Ctrl+C) now shows CANCELLED (shared lifecycle-aware banner, TenantContext
        1.0.12) with a warning line, instead of FAILED; exit code 130 and no receipt as before.

    1.8.3
      - An interrupted run (Ctrl+C) is never reported as SUCCESS: the banner shows FAILED with an
        error line; no receipt or success status is written and the exit code stays 130.
      - Stops waiting for a remote task that is no longer running without a fresh status.json after
        3 minutes: TaskTriggerFailed with the last task result in hexadecimal and the likely causes
        (local administrator, batch logon right, password), instead of waiting for the full limit.

    1.8.2
      - New HealthCheckerExcludedServers setting (short names or FQDNs, case-insensitive): listed
        servers stay in the inventory but Health Checker never starts or contacts them; their summary
        row is NotRun "excluded by configuration", without warning, and they are not counted in
        ServersTotal.

    1.8.1
      - Fixes the CIM/DCOM session options (1.8.0 mixed two parameter sets of New-CimSessionOption and
        could not start any remote task).
      - Never deserializes Health Checker XML from another server: its runner (1.1.0) converts the
        result on the server into plain JSON lines, read here with a size limit and a strict schema.
      - The work folder parent check also refuses generic write or full-control rights.

    1.8.0
      - Runs Health Checker on remote servers through a local scheduled task started in parallel over
        CIM/DCOM (WinRM is not used); the collector server still runs it in a local process.
      - Never passes -Server to Health Checker; refreshes the server files when their SHA256 differs.
      - Uses the daily scheduled result when the on-demand start fails and the result is recent enough
        (HealthCheckerMaxResultAgeHours, default 26).
      - New run statuses RemoteConnectionFailed, TaskTriggerFailed, TaskTimeout, StaleResult; new
        Summary columns ResultSource and ResultAgeHours at the end.
      - Keeps only the last Health Checker run; HealthCheckerKeepFailedRawOutputDays is no longer used.
      - Moves the default work folder to %ProgramFiles%\SmartM365\HealthChecker\Work, owned by
        Administrators with an administrators-only ACL; refuses unsafe parents (UnsafeWorkFolder) and
        checks the copy again immediately before starting it.

    1.6.9
      - Uses English Excluded status in the HTML report and log.

    1.6.8
      - Excludes configured servers from per-server collection and marks them in the HTML email.

    1.6.3
      - Renders per-server ERROR status values in red in the HTML report and email body.

    1.5.3
      - Stops attaching the HTML report when it is already used as the email body.
      - Uses Outlook-friendly inline/table layout for the report header and KPI cards.

    1.5.1
      - Fixes report-only regeneration when a source CSV contains a single row under StrictMode.

    1.5.0
      - Adds report-only regeneration from existing DATA-ALL CSV files.
      - Adds priority issue and readiness status summary tables to the HTML report.

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
      - Renders the full Exchange readiness inventory in the HTML report.

    1.4.0
      - Adds AD Exchange schema and Exchange readiness configuration inventory.

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

    [switch]$IncludeHealthChecker,

    [switch]$SkipFqdnAndUseServerName,

    [string]$RegenerateHtmlReportFromFolder,

    [switch]$SendRegeneratedReportMail,

    [int]$LowFreeSpaceThresholdPercent = 15,
    [int]$MaxItems = 0
)
# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
    foreach ($smartM365LimitName in @('TopUsers','TopMailboxes','MaxDevices','MaxSites','MaxTeams','MaxApps','MaxPolicies','Limit','MaxPages')) {
        $smartM365LimitVariable = Get-Variable -Name $smartM365LimitName -Scope Script -ErrorAction SilentlyContinue
        if ($smartM365LimitVariable -and -not $PSBoundParameters.ContainsKey($smartM365LimitName) -and $null -ne $smartM365LimitVariable.Value) {
            Set-Variable -Name $smartM365LimitName -Value ([int]$MaxItems) -Scope Script
        }
    }
}


if ($MaxItems -gt 0) {
    throw "-MaxItems is not supported by SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory because infrastructure and readiness status must be calculated from the complete on-premises Exchange estate."
}
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
$ScriptVersion = "1.8.4"
$RunId = (Get-Date).ToString("yyyyMMdd-HHmmss")
# Vendored Microsoft CSS-Exchange Health Checker (see ThirdParty\CSS-Exchange\HealthChecker\PROVENANCE.md).
$HealthCheckerRelativePath = 'ThirdParty\CSS-Exchange\HealthChecker\HealthChecker.ps1'
$HealthCheckerVersion = '26.10.13.1644'
$HealthCheckerExpectedSha256 = 'BF8758E96E8760226E3094D45A84909AFFA9F91167DE11F739D8E3D9343229DC'
# Remote servers run Health Checker locally through a scheduled task (see Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1).
$HealthCheckerRunnerFileName = 'SmartM365-Exchange-OnPrem-HealthChecker-LocalRunner.ps1'
$HealthCheckerRemoteRootPath = 'C:\Program Files\SmartM365\HealthChecker'
$HealthCheckerTaskPath = '\WCH\'
$HealthCheckerTaskName = 'SmartM365 Exchange Health Checker'

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

if (-not [string]::IsNullOrWhiteSpace($RegenerateHtmlReportFromFolder)) {
    $OutputFolder = Resolve-SmartM365ConfigTokenValue -Value $RegenerateHtmlReportFromFolder
}
else {
    $OutputFolder = Join-Path $OutputRoot $RunId
}
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
    try { Stop-SmartM365RunTranscript | Out-Null } catch {}
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
    Write-Log ("  TenantProfile: {0}" -f $Tenant)
    Write-Log ("  TenantKey: {0}" -f $global:SmartM365TenantKey)
    Write-Log ("  OutputPath: {0}" -f $OutputFolder)
    Write-Log ("  LogTextFile: {0}" -f $LogFile)
    Write-Log ("  TranscriptFile: {0}" -f $TranscriptFile)
    Write-Log ("  GeneratedCsvFiles: {0}" -f $script:ServersAndStorageGeneratedCsvPaths.Count)
    Write-Log ("  Warnings: {0}" -f $script:ServersAndStorageWarningCount)
    Write-Log ("  Errors: {0}" -f $script:ServersAndStorageErrorCount)
    if (-not [string]::IsNullOrWhiteSpace($ErrorMessage)) { Write-Log ("  ErrorMessage: {0}" -f $ErrorMessage) }

    Stop-ServersAndStorageTranscriptSafely
    if(Get-Command Complete-CoreSmartM365SourceReceipt -ErrorAction SilentlyContinue){
        $sourceReceipt = Complete-CoreSmartM365SourceReceipt -Status $Status -ErrorCount $script:ServersAndStorageErrorCount
        if($sourceReceipt){Invoke-ServersAndStorageSharePointUpload -LocalFilePath $sourceReceipt}
    }
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

function Get-ConfiguredExcludedServers {
    $property = $script:SmartM365EffectiveConfig.PSObject.Properties['ExcludedServers']
    if ($null -eq $property -or $null -eq $property.Value) { return @() }
    $value = $property.Value
    if ($value -isnot [array]) {
        throw "ExcludedServers must be a JSON array of server names."
    }

    $names = foreach ($item in $value) {
        if ($item -isnot [string] -or [string]::IsNullOrWhiteSpace($item)) {
            throw "ExcludedServers must contain only non-empty server names."
        }
        $item.Trim()
    }

    return @($names | Sort-Object -Unique)
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
                $minimumVersion = if ($modulePath -like '*WindowsPowerShell5*') { '1.0.57' } else { '1.0.87' }
                Import-Module -Name $modulePath -MinimumVersion $minimumVersion -Prefix Core -ErrorAction Stop
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

    $reportHtml = Get-Content -LiteralPath $HtmlReportPath -Raw -ErrorAction Stop
    if ($reportHtml -notmatch 'SmartM365EmailTemplate:v1') {
        $reportHtml = "<!-- SmartM365EmailTemplate:v1 -->`r`n$reportHtml"
    }

    $mailParams = @{
        From        = $from
        To          = $to
        Subject     = $Subject
        BodyHtml    = $reportHtml
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
        [string]$Path,
        [string[]]$Columns = @()
    )

    Import-SmartM365CoreModule
    $latestPath = $null
    $latestCsvFolder = [string](Get-SmartM365EffectiveConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($latestCsvFolder)) {
        Ensure-Directory -Path $latestCsvFolder
        $latestPath = Join-Path -Path $latestCsvFolder -ChildPath (Split-Path -Path $Path -Leaf)
    }
    Publish-CoreSmartM365Csv -Data @($Data) -TimestampedPath $Path -LatestPath $latestPath -Columns $Columns -Delimiter ';' -NoSharePointUpload | Out-Null
    $script:ServersAndStorageGeneratedCsvPaths.Add($Path) | Out-Null
    Invoke-ServersAndStorageSharePointUpload -LocalFilePath $Path

    if (-not [string]::IsNullOrWhiteSpace($latestPath)) {
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

function Get-HealthCheckerCsvColumn {
    param([Parameter(Mandatory)][ValidateSet('Findings', 'Summary')][string]$Kind)

    if ($Kind -eq 'Findings') {
        return @('ServerName', 'Category', 'CategoryOrder', 'RowNumber', 'Check', 'Item', 'Status', 'Value', 'ValueMasked', 'Recommendation', 'HealthCheckerVersion', 'GenerationTime')
    }
    return @('ServerName', 'RunStatus', 'RunMessage', 'DurationSeconds', 'HealthCheckerVersion', 'GenerationTime', 'ExchangeVersion', 'BuildNumber', 'CumulativeUpdate', 'SecurityUpdate', 'LatestInstallTime', 'SecurityUpdateStatus', 'CumulativeUpdateOutOfDate', 'CumulativeUpdateDaysOld', 'EndOfLifeStatus', 'VulnerabilityCount', 'Vulnerabilities', 'ErrorCount', 'WarningCount', 'OkCount', 'InfoCount', 'ServersProcessed', 'ServersTotal', 'ResultSource', 'ResultAgeHours')
}

function Get-HealthCheckerMaskingRule {
    # Values of these checks can hold host names, addresses, certificates, domains or tenant/app identifiers.
    return [pscustomobject]@{
        Categories = @('Hybrid Information')
        Checks     = @('Address', 'DNS Server', 'Internet Web Proxy', 'Http Proxy Setting', 'Http Proxy By Pass List', 'FriendlyName', 'Thumbprint', 'Certificate', 'Current Auth Certificate', 'Next Auth Certificate', 'TLS Certificate Name', 'TlsCertificateName', 'Internal Transport Certificate', 'Namespaces', 'Domain(s) configured for Hybrid use', 'Domain Trusts', 'On-Premises Smart Host Domain', 'Open Relay Wild Card Domain', 'Exchange Server Membership', 'Exchange Server Token Groups', 'Legacy Exchange Security Groups', 'AD Site', 'Location', 'Description', 'OAuth between Exchange Server and Exchange Online', 'OAuth between Exchange Server and Microsoft Teams')
    }
}

function Get-HealthCheckerSetting {
    $perServerMinutes = [int](Get-SmartM365EffectiveConfigValue -Name 'HealthCheckerTimeoutMinutesPerServer' -DefaultValue 30)
    $budgetMinutes = [int](Get-SmartM365EffectiveConfigValue -Name 'HealthCheckerTotalBudgetMinutes' -DefaultValue 1320)
    $maxResultAgeHours = [int](Get-SmartM365EffectiveConfigValue -Name 'HealthCheckerMaxResultAgeHours' -DefaultValue 26)
    $workRoot = [string](Get-SmartM365EffectiveConfigValue -Name 'HealthCheckerWorkFolderPath' -DefaultValue '')
    if ($perServerMinutes -lt 1 -or $perServerMinutes -gt 240) { throw "HealthCheckerTimeoutMinutesPerServer must be between 1 and 240." }
    if ($budgetMinutes -lt 1 -or $budgetMinutes -gt 1440) { throw "HealthCheckerTotalBudgetMinutes must be between 1 and 1440." }
    if ($maxResultAgeHours -lt 1 -or $maxResultAgeHours -gt 168) { throw "HealthCheckerMaxResultAgeHours must be between 1 and 168." }
    if ([string]::IsNullOrWhiteSpace($workRoot)) {
        # Under Program Files, where standard users cannot create folders: the verified copy runs elevated from there.
        $workRoot = Join-Path -Path $env:ProgramFiles -ChildPath 'SmartM365\HealthChecker\Work'
    }
    # Raw Health Checker files hold the full server configuration: they stay on a local disk only.
    if ($workRoot.StartsWith('\\') -or -not [System.IO.Path]::IsPathRooted($workRoot)) {
        throw "HealthCheckerWorkFolderPath must be a local absolute path, not a share: $workRoot"
    }
    # Servers kept in the inventory but never contacted by Health Checker (short name or FQDN).
    $excludedValue = Get-SmartM365EffectiveConfigValue -Name 'HealthCheckerExcludedServers' -DefaultValue @()
    $excludedServers = @(foreach ($item in @($excludedValue)) {
            if ($null -eq $item) { continue }
            if ($item -isnot [string] -or [string]::IsNullOrWhiteSpace($item)) { throw 'HealthCheckerExcludedServers must be a JSON array of server names.' }
            $item.Trim()
        })
    # Only the last run is kept now; the former retention setting is ignored.
    if ($null -ne (Get-SmartM365EffectiveConfigValue -Name 'HealthCheckerKeepFailedRawOutputDays' -DefaultValue $null)) {
        Write-Log 'HealthCheckerKeepFailedRawOutputDays is no longer used and can be removed from the local configuration: only the last Health Checker run is kept.'
    }

    return [pscustomobject]@{
        TimeoutMinutesPerServer    = $perServerMinutes
        TotalBudgetMinutes         = $budgetMinutes
        MaxResultAgeHours          = $maxResultAgeHours
        ExcludedServers            = $excludedServers
        WorkFolderPath             = $workRoot
        RemoteRootPath             = $HealthCheckerRemoteRootPath
        TaskPath                   = $HealthCheckerTaskPath
        TaskName                   = $HealthCheckerTaskName
        PollSeconds                = 15
        TaskStartGraceSeconds      = 180
    }
}

function Test-HealthCheckerScriptIntegrity {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ExpectedSha256
    )

    $result = [pscustomobject]@{ IsValid = $false; Reason = $null; Sha256 = $null; SignatureStatus = $null; Signer = $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Reason = "Health Checker script not found: $Path"
        return $result
    }

    $result.Sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($result.Sha256 -ne $ExpectedSha256.ToUpperInvariant()) {
        $result.Reason = "SHA256 mismatch: expected $($ExpectedSha256.ToUpperInvariant()), found $($result.Sha256)."
        return $result
    }

    $signature = Get-AuthenticodeSignature -FilePath $Path
    $result.SignatureStatus = [string]$signature.Status
    if ($signature.SignerCertificate) { $result.Signer = [string]$signature.SignerCertificate.Subject }
    if ([string]$signature.Status -ne 'Valid' -or -not $signature.SignerCertificate) {
        $result.Reason = "Authenticode signature is not valid (status $($signature.Status))."
        return $result
    }
    $subject = [string]$signature.SignerCertificate.Subject
    $issuer = [string]$signature.SignerCertificate.Issuer
    if ($subject -notmatch '(^|,\s*)CN=Microsoft Corporation(,|$)' -or $subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)' -or $issuer -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)') {
        $result.Reason = "Authenticode signer is not Microsoft Corporation: $subject"
        return $result
    }

    $result.IsValid = $true
    return $result
}

function Test-HealthCheckerElevation {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-HealthCheckerUnsafeAclEntry {
    param([Parameter(Mandatory = $true)][System.Security.AccessControl.FileSystemSecurity]$Acl)

    # SYSTEM, Administrators, TrustedInstaller; CREATOR OWNER only grants rights to whoever creates a child.
    $trustedSids = @('S-1-5-18', 'S-1-5-32-544', 'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464')
    $writeRights = [System.Security.AccessControl.FileSystemRights]'WriteData, AppendData, WriteExtendedAttributes, DeleteSubdirectoriesAndFiles, WriteAttributes, Delete, ChangePermissions, TakeOwnership'
    $unsafe = New-Object System.Collections.Generic.List[string]
    $owner = [string]$Acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value
    if ($owner -notin $trustedSids) { $unsafe.Add("owner $owner") }
    foreach ($rule in @($Acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))) {
        $sid = [string]$rule.IdentityReference.Value
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow -or $sid -in $trustedSids -or $sid -eq 'S-1-3-0') { continue }
        # Generic rights (GENERIC_ALL 0x10000000, GENERIC_WRITE 0x40000000) are not mapped by .NET: test them too.
        $mask = [int64][int]$rule.FileSystemRights -band 4294967295
        if (($mask -band [int64]$writeRights) -ne 0 -or ($mask -band 0x50000000) -ne 0) { $unsafe.Add(('{0} {1}' -f $sid, $rule.FileSystemRights)) }
    }
    return $unsafe.ToArray()
}

function Get-HealthCheckerWorkFolderAncestor {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string[]]$TrustedBase = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:SystemRoot)
    )

    # Parents below the first system folder, nearest first; system folders and the drive root are trusted.
    $bases = @($TrustedBase | Where-Object { $_ } | ForEach-Object { [System.IO.Path]::GetFullPath($_).TrimEnd('\') })
    $ancestors = New-Object System.Collections.Generic.List[string]
    $current = Split-Path -Path ([System.IO.Path]::GetFullPath($Path).TrimEnd('\')) -Parent
    while ($current) {
        $normalized = $current.TrimEnd('\')
        if ($normalized.Length -le 2 -or $normalized -in $bases) { break }
        $ancestors.Add($normalized)
        $current = Split-Path -Path $normalized -Parent
    }
    return $ancestors.ToArray()
}

function Set-HealthCheckerWorkFolderAclRule {
    param([Parameter(Mandatory)][System.Security.AccessControl.DirectorySecurity]$Acl)

    # Administrators own the work folder; only SYSTEM and Administrators have access, inheritance removed.
    $Acl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')))
    $Acl.SetAccessRuleProtection($true, $false)
    foreach ($existingRule in @($Acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))) {
        [void]$Acl.RemoveAccessRuleSpecific($existingRule)
    }
    foreach ($sid in @('S-1-5-18', 'S-1-5-32-544')) {
        $identity = New-Object System.Security.Principal.SecurityIdentifier($sid)
        $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($identity, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
    }
    return $Acl
}

function Initialize-HealthCheckerWorkFolder {
    param([Parameter(Mandatory)][string]$WorkRoot)

    # The verified Health Checker copy runs elevated from here: only administrators may change this folder or its parents.
    $ancestors = @(Get-HealthCheckerWorkFolderAncestor -Path $WorkRoot)
    [array]::Reverse($ancestors)
    foreach ($ancestor in $ancestors) {
        if (-not (Test-Path -LiteralPath $ancestor -PathType Container)) {
            New-Item -ItemType Directory -Path $ancestor -Force | Out-Null
            $ancestorAcl = Get-Acl -LiteralPath $ancestor
            $ancestorAcl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')))
            Set-Acl -LiteralPath $ancestor -AclObject $ancestorAcl
        }
        $unsafe = @(Get-HealthCheckerUnsafeAclEntry -Acl (Get-Acl -LiteralPath $ancestor))
        if ($unsafe.Count -gt 0) {
            throw ("A parent of the Health Checker work folder is not restricted to administrators ({0}): {1}" -f ($unsafe -join '; '), $ancestor)
        }
    }
    if (-not (Test-Path -LiteralPath $WorkRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $WorkRoot -Force | Out-Null
    }
    Set-Acl -LiteralPath $WorkRoot -AclObject (Set-HealthCheckerWorkFolderAclRule -Acl (Get-Acl -LiteralPath $WorkRoot))
}

function Invoke-HealthCheckerProcess {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$OutputFolder,
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [string]$ExpectedSha256,
        [string]$PowerShellPath = (Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe')
    )

    Ensure-Directory -Path $OutputFolder
    $consolePath = Join-Path -Path $OutputFolder -ChildPath 'HealthChecker-Console.txt'
    $errorPath = Join-Path -Path $OutputFolder -ChildPath 'HealthChecker-Errors.txt'
    # A child process: Health Checker calls exit when it is not elevated, and a hung run can be killed.
    # No -Server: this runs for the local server only; remote servers run it through their own scheduled task.
    $arguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $ScriptPath), '-OutputFilePath', ('"{0}"' -f $OutputFolder), '-SkipVersionCheck')
    $started = Get-Date
    $result = [pscustomobject]@{ ServerName = $ServerName; Status = $null; Message = $null; ExitCode = $null; DurationSeconds = $null; XmlPath = $null }
    if ($ExpectedSha256) {
        # Checked again immediately before the start, so the file that runs is the one just verified.
        $integrity = Test-HealthCheckerScriptIntegrity -Path $ScriptPath -ExpectedSha256 $ExpectedSha256
        if (-not $integrity.IsValid) {
            $result.Status = 'IntegrityFailed'
            $result.Message = "Health Checker was not run: $($integrity.Reason)"
            $result.DurationSeconds = 0
            return $result
        }
    }
    try {
        $process = Start-Process -FilePath $PowerShellPath -ArgumentList $arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $consolePath -RedirectStandardError $errorPath -ErrorAction Stop
        $null = $process.Handle
    }
    catch {
        $result.Status = 'LaunchFailed'
        $result.Message = "Health Checker could not be started: $($_.Exception.Message)"
        $result.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
        return $result
    }

    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        & taskkill.exe /PID $process.Id /T /F 2>&1 | Out-Null
        try { [void]$process.WaitForExit(30000) } catch { Write-Verbose "Stopped Health Checker process already gone: $($_.Exception.Message)" }
        $result.Status = 'Timeout'
        $result.Message = "Health Checker exceeded the $([math]::Round($TimeoutSeconds / 60, 1)) minute limit and was stopped."
        $result.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
        return $result
    }

    $result.ExitCode = $process.ExitCode
    $result.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
    $xmlFiles = @(Get-ChildItem -LiteralPath $OutputFolder -Filter 'HealthChecker-*.xml' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike 'HealthChecker-ScriptDebugObject*' -and $_.Name -notlike 'HealthChecker-ExchangeDCCoreRatio-*' } | Sort-Object LastWriteTime -Descending)
    $consoleText = if (Test-Path -LiteralPath $consolePath -PathType Leaf) { [System.IO.File]::ReadAllText($consolePath) } else { '' }
    if ($xmlFiles.Count -gt 0) {
        $result.Status = 'Completed'
        $result.XmlPath = $xmlFiles[0].FullName
    }
    elseif ($consoleText -match 'needs to be executed in elevated mode') {
        $result.Status = 'NotElevated'
        $result.Message = 'Health Checker requires an elevated (administrator) session.'
    }
    else {
        $result.Status = 'NoOutput'
        $result.Message = "Health Checker produced no XML result (exit code $($result.ExitCode))."
    }
    return $result
}

function ConvertTo-HealthCheckerSafeValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $null }
    $text = if ($Value -is [string]) { $Value } else { (@($Value) | ForEach-Object { [string]$_ }) -join '; ' }
    $text = [regex]::Replace($text, '\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b', '[guid]')
    $text = [regex]::Replace($text, '\b(?:\d{1,3}\.){3}\d{1,3}\b', '[ip]')
    # IPv6 candidates are confirmed by the parser so clock times such as 06:00:00 are kept.
    $text = [regex]::Replace($text, '(?<![\w:.])(?:[0-9a-fA-F]{0,4}:){2,7}[0-9a-fA-F]{0,4}(?![\w:])', [System.Text.RegularExpressions.MatchEvaluator]{
            param($match)
            $address = $null
            if ($match.Value -notmatch '^\d{1,2}:\d{2}(:\d{2})?$' -and [System.Net.IPAddress]::TryParse($match.Value, [ref]$address) -and $address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) { return '[ip]' }
            return $match.Value
        })
    $text = [regex]::Replace($text, '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', '[email]')
    $text = [regex]::Replace($text, '\s+', ' ').Trim()
    if ($text.Length -gt 1000) { $text = $text.Substring(0, 997) + '...' }
    return $text
}

function Get-HealthCheckerLineValue {
    param(
        [AllowNull()][object]$Line,
        [Parameter(Mandatory)][string]$Name
    )

    # Serialized entries can be plain strings when Export-Clixml ran out of depth.
    if ($null -eq $Line) { return $null }
    if ($Line -is [string]) { return $(if ($Name -eq 'DisplayValue') { $Line } else { $null }) }
    $property = $Line.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function ConvertTo-HealthCheckerResultLine {
    param([Parameter(Mandatory)][string]$Path)

    # Import-Clixml is only used where the XML was produced (the server itself); the result lines
    # hold strings and numbers only, so they can cross to the collector as JSON.
    $data = Import-Clixml -LiteralPath $Path
    $displayResults = Get-HealthCheckerLineValue -Line $data -Name 'DisplayResults'
    if ($displayResults -isnot [System.Collections.IDictionary] -or $displayResults.Count -eq 0) {
        throw 'The Health Checker XML has no DisplayResults section.'
    }

    $lines = New-Object 'System.Collections.Generic.List[object]'
    $groups = @(foreach ($entry in $displayResults.GetEnumerator()) {
        $order = 0
        [void][int]::TryParse([string](Get-HealthCheckerLineValue -Line $entry.Key -Name 'DisplayOrder'), [ref]$order)
        $categoryName = if ($entry.Key -is [string]) { $entry.Key } else { [string](Get-HealthCheckerLineValue -Line $entry.Key -Name 'Name') }
        [pscustomobject]@{ Category = $categoryName; Order = $order; Lines = @($entry.Value) }
    })
    foreach ($group in @($groups | Sort-Object Order, Category)) {
        $parentCheck = $null
        foreach ($line in $group.Lines) {
            if ($null -eq $line) { continue }
            $name = [string](Get-HealthCheckerLineValue -Line $line -Name 'Name')
            $tab = 0
            [void][int]::TryParse([string](Get-HealthCheckerLineValue -Line $line -Name 'TabNumber'), [ref]$tab)
            $testingName = [string](Get-HealthCheckerLineValue -Line $line -Name 'TestingName')
            $isSubLine = $tab -gt 1 -or [string]::IsNullOrWhiteSpace($name) -or $name -in @('Error', 'Warning')
            if (-not $isSubLine -or [string]::IsNullOrWhiteSpace($parentCheck)) {
                $parentCheck = if ([string]::IsNullOrWhiteSpace($name)) { $testingName } else { $name }
            }
            # An Error/Warning line with its own testing name (for example "Out of Date") is its own check.
            $check = if ($isSubLine -and -not [string]::IsNullOrWhiteSpace($testingName) -and $testingName -ne $name -and $testingName -notin @('Error', 'Warning')) { $testingName } else { $parentCheck }
            $writeType = [string](Get-HealthCheckerLineValue -Line $line -Name 'WriteType')
            $displayValue = Get-HealthCheckerLineValue -Line $line -Name 'DisplayValue'
            $displayText = if ($writeType -eq 'OutColumns' -or $null -eq $displayValue) { $null } elseif ($displayValue -is [string]) { $displayValue } else { (@($displayValue) | ForEach-Object { [string]$_ }) -join '; ' }
            $lines.Add([pscustomobject][ordered]@{
                Category      = [string]$group.Category
                CategoryOrder = [int]$group.Order
                Check         = [string]$check
                Item          = $(if ($isSubLine -and -not [string]::IsNullOrWhiteSpace($name)) { $name } else { $null })
                Name          = $name
                WriteType     = $writeType
                DisplayValue  = $displayText
                TestingName   = [string](Get-HealthCheckerLineValue -Line $line -Name 'TestingName')
                CustomName    = [string](Get-HealthCheckerLineValue -Line $line -Name 'CustomName')
                CustomValue   = [string](Get-HealthCheckerLineValue -Line $line -Name 'CustomValue')
            }) | Out-Null
        }
    }
    return $lines.ToArray()
}

function ConvertFrom-HealthCheckerResultLine {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Lines,
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$HealthCheckerVersion
    )

    $masking = Get-HealthCheckerMaskingRule
    $statusMap = @{ Red = 'Error'; Yellow = 'Warning'; Green = 'OK' }
    $lines = $Lines
    $findings = New-Object 'System.Collections.Generic.List[object]'
    $generationTime = $null
    $rowNumber = 0
    foreach ($entry in $lines) {
        $rowNumber++
        $writeType = $entry.WriteType
        $status = if ($statusMap.ContainsKey($writeType)) { $statusMap[$writeType] } else { 'Info' }
        $effectiveName = if (-not [string]::IsNullOrWhiteSpace([string]$entry.Item) -and [string]$entry.Item -notin @('Error', 'Warning')) { [string]$entry.Item } else { [string]$entry.Check }
        $masked = $entry.Category -in $masking.Categories -or $effectiveName -in $masking.Checks
        $rawValue = if ($writeType -eq 'OutColumns') { '[table]' } else { $entry.DisplayValue }
        $safeValue = ConvertTo-HealthCheckerSafeValue -Value $rawValue
        $recommendation = if ($masked -or [string]::IsNullOrWhiteSpace($safeValue)) { $null } else { (@([regex]::Matches($safeValue, 'https?://[^\s;,)]+') | ForEach-Object { $_.Value.TrimEnd('.') }) | Select-Object -Unique) -join '; ' }
        if ($entry.Category -eq 'Exchange Information' -and $entry.Check -eq 'Generation Time') { $generationTime = $safeValue }
        $findings.Add([pscustomobject]@{
            ServerName           = $ServerName
            Category             = $entry.Category
            CategoryOrder        = $entry.CategoryOrder
            RowNumber            = $rowNumber
            Check                = $entry.Check
            Item                 = $entry.Item
            Status               = $status
            Value                = $(if ($masked) { '[masked]' } else { $safeValue })
            ValueMasked          = $masked
            Recommendation       = $(if ([string]::IsNullOrWhiteSpace($recommendation)) { $null } else { $recommendation })
            HealthCheckerVersion = $HealthCheckerVersion
            GenerationTime       = $null
        }) | Out-Null
    }
    foreach ($finding in $findings) { $finding.GenerationTime = $generationTime }

    # Same vulnerability matching as Health Checker's own -VulnerabilityReport.
    $vulnerabilities = New-Object 'System.Collections.Generic.List[string]'
    foreach ($entry in @($lines | Where-Object { $_.Category -eq 'Security Vulnerability' })) {
        $name = $entry.Name
        $cveName = ''
        if ($name -eq 'Security Vulnerability') {
            $cveName = if ($entry.CustomName -ne 'Security Vulnerability' -and -not [string]::IsNullOrWhiteSpace($entry.CustomName)) { $entry.CustomName } else { $entry.CustomValue }
        }
        elseif (($name -eq 'IIS module anomalies detected' -and $entry.CustomValue -eq 'True') -or (-not [string]::IsNullOrWhiteSpace($name) -and $name -ne 'IIS module anomalies detected')) {
            $cveName = $name
        }
        if (-not [string]::IsNullOrWhiteSpace($cveName) -and -not $vulnerabilities.Contains($cveName)) { $vulnerabilities.Add($cveName) }
    }

    $exchangeLines = @($lines | Where-Object { $_.Category -eq 'Exchange Information' })
    $firstValue = {
        param([string]$CheckName)
        $match = @($exchangeLines | Where-Object { $_.Check -eq $CheckName -and [string]::IsNullOrWhiteSpace([string]$_.Item) } | Select-Object -First 1)
        if ($match.Count -eq 0) { return $null }
        return ConvertTo-HealthCheckerSafeValue -Value $match[0].DisplayValue
    }
    $version = & $firstValue 'Version'
    $exchangeTexts = @($exchangeLines | ForEach-Object { [string]$_.DisplayValue })
    $outOfDateLine = @($exchangeLines | Where-Object { $_.TestingName -eq 'Out of Date' } | Select-Object -First 1)
    $daysOld = $null
    if ($outOfDateLine.Count -gt 0 -and [string]$outOfDateLine[0].DisplayValue -match 'build that is (\d+) days old') { $daysOld = [int]$Matches[1] }
    $securityStatus = if ($exchangeTexts -match 'out of support and no longer receives SUs') { 'OutOfSupportNoSU' }
        elseif ($exchangeTexts -match 'Not on the latest SU') { 'MissingLatestSU' }
        elseif ($exchangeTexts -match 'Not on the latest Exchange Patch') { 'NewerPatchAvailable' }
        elseif ($null -ne $version) { 'Latest' }
        else { 'Unknown' }
    $endOfLife = @($exchangeLines | Where-Object { $_.Check -eq 'End Of Life' } | Select-Object -First 1)
    $endOfLifeStatus = if ($endOfLife.Count -eq 0) { 'Supported' } elseif ($endOfLife[0].WriteType -eq 'Red') { 'EndOfLifeSoonOrReached' } else { 'EndOfLifeWithinOneYear' }
    $cumulativeUpdate = $null
    $securityUpdate = $null
    if ($version -match '\b(CU\s?\d+)\b') { $cumulativeUpdate = $Matches[1].Replace(' ', '') }
    if ($version -match '\bCU\s?\d+\s+(.+)$') { $securityUpdate = $Matches[1].Trim() }

    return [pscustomobject]@{
        Findings        = $findings.ToArray()
        Vulnerabilities = $vulnerabilities.ToArray()
        Facts           = [pscustomobject]@{
            GenerationTime            = $generationTime
            ExchangeVersion           = $version
            BuildNumber               = & $firstValue 'Build Number'
            CumulativeUpdate          = $cumulativeUpdate
            SecurityUpdate            = $securityUpdate
            LatestInstallTime         = & $firstValue 'Latest Install Time (SU/CU)'
            SecurityUpdateStatus      = $securityStatus
            CumulativeUpdateOutOfDate = ($outOfDateLine.Count -gt 0)
            CumulativeUpdateDaysOld   = $daysOld
            EndOfLifeStatus           = $endOfLifeStatus
            ErrorCount                = @($findings | Where-Object { $_.Status -eq 'Error' }).Count
            WarningCount              = @($findings | Where-Object { $_.Status -eq 'Warning' }).Count
            OkCount                   = @($findings | Where-Object { $_.Status -eq 'OK' }).Count
            InfoCount                 = @($findings | Where-Object { $_.Status -eq 'Info' }).Count
        }
    }
}

function ConvertFrom-HealthCheckerXml {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$HealthCheckerVersion
    )

    # Only for XML produced on this server; results from other servers arrive as JSON.
    $lines = @(ConvertTo-HealthCheckerResultLine -Path $Path)
    return ConvertFrom-HealthCheckerResultLine -Lines $lines -ServerName $ServerName -HealthCheckerVersion $HealthCheckerVersion
}

function Read-HealthCheckerResultJson {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxBytes = 20MB
    )

    # Data from another server: size capped, strict schema, only the expected fields kept as strings or numbers.
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($item.Length -gt $MaxBytes) { throw ("The Health Checker result is larger than {0} MB." -f [math]::Round($MaxBytes / 1MB)) }
    $document = [System.IO.File]::ReadAllText($Path) | ConvertFrom-Json
    if ($null -eq $document -or $null -eq $document.PSObject.Properties['SchemaVersion'] -or [string]$document.SchemaVersion -ne '1' -or $null -eq $document.PSObject.Properties['Lines']) {
        throw 'The Health Checker result does not follow the expected schema (version 1).'
    }
    $textFields = @('Category', 'Check', 'Item', 'Name', 'WriteType', 'DisplayValue', 'TestingName', 'CustomName', 'CustomValue')
    $lines = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in @($document.Lines)) {
        if ($null -eq $entry -or $entry -is [string] -or $entry -is [ValueType]) { throw 'The Health Checker result holds an unexpected line.' }
        $record = [ordered]@{}
        foreach ($field in $textFields) {
            $property = $entry.PSObject.Properties[$field]
            $value = if ($null -eq $property) { $null } else { $property.Value }
            if ($null -ne $value -and $value -isnot [string] -and $value -isnot [ValueType]) { throw "The Health Checker result field $field is not a plain value." }
            $record[$field] = $(if ($null -eq $value) { $null } else { [string]$value })
        }
        $order = 0
        $orderProperty = $entry.PSObject.Properties['CategoryOrder']
        if ($orderProperty) { [void][int]::TryParse([string]$orderProperty.Value, [ref]$order) }
        $record['CategoryOrder'] = $order
        $lines.Add([pscustomobject]$record) | Out-Null
    }
    if ($lines.Count -eq 0) { throw 'The Health Checker result has no lines.' }
    return $lines.ToArray()
}

function ConvertTo-HealthCheckerSummaryRow {
    param(
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$RunStatus,
        [AllowNull()][string]$RunMessage,
        [AllowNull()][object]$DurationSeconds,
        [Parameter(Mandatory)][string]$HealthCheckerVersion,
        [AllowNull()][object]$Parsed,
        [AllowNull()][string]$ResultSource,
        [AllowNull()][object]$ResultAgeHours
    )

    $facts = if ($Parsed) { $Parsed.Facts } else { $null }
    return [pscustomobject]@{
        ServerName                = $ServerName
        RunStatus                 = $RunStatus
        RunMessage                = $RunMessage
        DurationSeconds           = $DurationSeconds
        HealthCheckerVersion      = $HealthCheckerVersion
        GenerationTime            = $(if ($facts) { $facts.GenerationTime } else { $null })
        ExchangeVersion           = $(if ($facts) { $facts.ExchangeVersion } else { $null })
        BuildNumber               = $(if ($facts) { $facts.BuildNumber } else { $null })
        CumulativeUpdate          = $(if ($facts) { $facts.CumulativeUpdate } else { $null })
        SecurityUpdate            = $(if ($facts) { $facts.SecurityUpdate } else { $null })
        LatestInstallTime         = $(if ($facts) { $facts.LatestInstallTime } else { $null })
        SecurityUpdateStatus      = $(if ($facts) { $facts.SecurityUpdateStatus } else { 'Unknown' })
        CumulativeUpdateOutOfDate = $(if ($facts) { $facts.CumulativeUpdateOutOfDate } else { $null })
        CumulativeUpdateDaysOld   = $(if ($facts) { $facts.CumulativeUpdateDaysOld } else { $null })
        EndOfLifeStatus           = $(if ($facts) { $facts.EndOfLifeStatus } else { $null })
        VulnerabilityCount        = $(if ($Parsed) { @($Parsed.Vulnerabilities).Count } else { $null })
        Vulnerabilities           = $(if ($Parsed) { @($Parsed.Vulnerabilities) -join '; ' } else { $null })
        ErrorCount                = $(if ($facts) { $facts.ErrorCount } else { $null })
        WarningCount              = $(if ($facts) { $facts.WarningCount } else { $null })
        OkCount                   = $(if ($facts) { $facts.OkCount } else { $null })
        InfoCount                 = $(if ($facts) { $facts.InfoCount } else { $null })
        ServersProcessed          = $null
        ServersTotal              = $null
        ResultSource              = $(if ($Parsed) { $ResultSource } else { $null })
        ResultAgeHours            = $(if ($Parsed) { $ResultAgeHours } else { $null })
    }
}

function Remove-HealthCheckerPreviousWorkFolder {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$WorkRoot,
        [Parameter(Mandatory)][string]$CurrentRunId
    )

    if (-not (Test-Path -LiteralPath $WorkRoot -PathType Container)) { return }
    # Only the last run is kept: every earlier run folder created by this collector (RunId format) is removed.
    foreach ($folder in @(Get-ChildItem -LiteralPath $WorkRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{8}-\d{6}$' -and $_.Name -ne $CurrentRunId })) {
        if (-not $PSCmdlet.ShouldProcess($folder.FullName, 'Remove previous Health Checker work folder')) { continue }
        try { Remove-Item -LiteralPath $folder.FullName -Recurse -Force -ErrorAction Stop }
        catch { Write-Log ("Previous Health Checker work folder could not be removed: {0}. {1}" -f $folder.FullName, $_.Exception.Message) 'WARN' }
    }
}

function Test-HealthCheckerServerExcluded {
    param(
        [Parameter(Mandatory)][object]$Server,
        [AllowNull()][AllowEmptyCollection()][string[]]$ExcludedServers
    )

    # Case-insensitive, on the short name and on the FQDN; a configured FQDN also matches the short name.
    $name = [string]$Server.Name
    $fqdn = if ($Server.PSObject.Properties['Fqdn']) { [string]$Server.Fqdn } else { '' }
    foreach ($entry in @($ExcludedServers)) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $candidate = $entry.Trim()
        if ($candidate -ieq $name -or ($fqdn -and $candidate -ieq $fqdn) -or ($candidate.Contains('.') -and $candidate.Split('.')[0] -ieq $name)) { return $true }
    }
    return $false
}

function Get-HealthCheckerServerComputerName {
    param([Parameter(Mandatory)][object]$Server)

    $fqdn = if ($Server.PSObject.Properties['Fqdn']) { [string]$Server.Fqdn } else { '' }
    $skipFqdn = [bool](Get-Variable -Name 'SkipFqdnAndUseServerName' -ValueOnly -ErrorAction SilentlyContinue)
    if ($skipFqdn -or [string]::IsNullOrWhiteSpace($fqdn)) { return [string]$Server.Name }
    return $fqdn
}

function ConvertTo-HealthCheckerAdminSharePath {
    param(
        [Parameter(Mandatory)][string]$Computer,
        [Parameter(Mandatory)][string]$LocalPath
    )

    if ($LocalPath -notmatch '^(?<Drive>[A-Za-z]):\\(?<Rest>.+)$') {
        throw "Health Checker remote root must be a local absolute path with a drive letter: $LocalPath"
    }
    return ('\\{0}\{1}$\{2}' -f $Computer, $Matches['Drive'].ToUpperInvariant(), $Matches['Rest'])
}

function ConvertTo-HealthCheckerRemoteMessage {
    param(
        [AllowNull()][string]$Message,
        [Parameter(Mandatory)][string]$Computer
    )

    if ([string]::IsNullOrWhiteSpace($Message)) { return $null }
    $text = $Message
    $dot = $Computer.IndexOf('.')
    if ($dot -gt 0) {
        # Summary messages never carry the domain: host.domain becomes host.
        $text = [regex]::Replace($text, '\.' + [regex]::Escape($Computer.Substring($dot + 1)) + '\b', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    $text = ConvertTo-HealthCheckerSafeValue -Value $text
    if ($text.Length -gt 300) { $text = $text.Substring(0, 297) + '...' }
    return $text
}

function ConvertTo-HealthCheckerUtcTime {
    param([AllowNull()][object]$Value)

    # Windows PowerShell keeps ISO 8601 strings from ConvertFrom-Json; PowerShell 7 returns DateTime.
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
        return $parsed.ToUniversalTime()
    }
    return $null
}

function Read-HealthCheckerRemoteStatus {
    param([Parameter(Mandatory)][string]$OutputSharePath)

    $statusPath = Join-Path -Path $OutputSharePath -ChildPath 'status.json'
    if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { return $null }
    try { return ([System.IO.File]::ReadAllText($statusPath) | ConvertFrom-Json) }
    catch { return $null }
}

function Get-HealthCheckerTaskLogonFailure {
    param([AllowNull()][object]$LastTaskResult)

    if ($null -eq $LastTaskResult) { return $null }
    $code = [long]$LastTaskResult
    if ($code -lt 0) { $code += 4294967296 }
    # Task Scheduler results that mean the run account could not log on: the runner never started.
    $reasons = @{
        2147943726 = 'wrong user name or password'
        2147943727 = 'account restriction'
        2147943730 = 'password expired'
        2147943731 = 'account disabled'
        2147943785 = "'Log on as a batch job' right missing"
        2147944309 = 'account locked out'
    }
    if (-not $reasons.ContainsKey($code)) { return $null }
    return ("The Health Checker task could not log on (0x{0:X8}, {1}). If the service account password expired or changed, run the Health Checker task installer again." -f $code, $reasons[$code])
}

function ConvertFrom-HealthCheckerRunnerState {
    param([Parameter(Mandatory)][object]$RemoteStatus)

    $state = [string]$RemoteStatus.State
    $message = [string]$RemoteStatus.Message
    $status = switch ($state) {
        'Completed' { 'Completed' }
        'Timeout' { 'TaskTimeout' }
        'IntegrityFailed' { 'IntegrityFailed' }
        'NotElevated' { 'NotElevated' }
        'LaunchFailed' { 'LaunchFailed' }
        default { 'NoOutput' }
    }
    if ($status -eq 'NoOutput' -and [string]::IsNullOrWhiteSpace($message)) {
        $message = "The Health Checker task ended with state '$state' and no XML result."
    }
    return [pscustomobject]@{ Status = $status; Message = $(if ($status -eq 'Completed') { $null } else { $message }) }
}

function Resolve-HealthCheckerScheduledFallback {
    param(
        [Parameter(Mandatory)][string]$FailureStatus,
        [AllowNull()][string]$FailureMessage,
        [AllowNull()][object]$RemoteStatus,
        [Parameter(Mandatory)][datetime]$NowUtc,
        [Parameter(Mandatory)][int]$MaxResultAgeHours
    )

    # Only used when the on-demand start failed: the runner did not run, so the last scheduled result is intact.
    $result = [pscustomobject]@{ Status = $FailureStatus; Message = $FailureMessage; UseResult = $false; ResultAgeHours = $null }
    if ($null -eq $RemoteStatus -or [string]$RemoteStatus.State -ne 'Completed' -or [string]$RemoteStatus.ResultFileName -cne 'result.json') { return $result }
    $finished = ConvertTo-HealthCheckerUtcTime -Value $RemoteStatus.FinishedAtUtc
    if ($null -eq $finished) { return $result }

    $age = [math]::Max(0, [math]::Round(($NowUtc.ToUniversalTime() - $finished).TotalHours, 1))
    $result.ResultAgeHours = $age
    if ($age -le $MaxResultAgeHours) {
        $result.Status = 'Completed'
        $result.Message = "On-demand run failed ($FailureStatus): $FailureMessage Using the scheduled result from $age h ago."
        $result.UseResult = $true
    }
    else {
        $result.Status = 'StaleResult'
        $result.Message = "On-demand run failed ($FailureStatus): $FailureMessage The last scheduled result is $age h old (limit $MaxResultAgeHours h)."
    }
    return $result
}

function Copy-HealthCheckerRemoteFile {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestinationPath
    )

    $sourceHash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
    if ((Test-Path -LiteralPath $DestinationPath -PathType Leaf) -and (Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash -eq $sourceHash) { return $false }
    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    if ((Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash -ne $sourceHash) {
        throw "The deployed file does not match its source: $(Split-Path -Path $DestinationPath -Leaf)"
    }
    return $true
}

function New-HealthCheckerDcomSessionOption {
    # DCOM, not WinRM: PowerShell remoting is disabled on the Exchange servers. The DCOM parameter set
    # (-Impersonation/-PacketPrivacy/-PacketIntegrity) returns DComSessionOptions; it cannot be combined with -Protocol.
    return New-CimSessionOption -Impersonation Impersonate -PacketPrivacy -PacketIntegrity
}

function Start-HealthCheckerRemoteRun {
    param(
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$Computer,
        [Parameter(Mandatory)][string]$VerifiedScriptPath,
        [Parameter(Mandatory)][string]$RunnerSourcePath,
        [Parameter(Mandatory)][object]$Settings
    )

    $run = [pscustomobject]@{
        ServerName      = $ServerName
        Computer        = $Computer
        Session         = $null
        OutputSharePath = $null
        PreviousStatus  = $null
        WaitToken       = $null
        StartedAt       = Get-Date
        Done            = $false
        Status          = $null
        Message         = $null
        RemoteStatus    = $null
    }
    $fail = {
        param($Status, $Message)
        $run.Done = $true
        $run.Status = $Status
        $run.Message = $Message
        return $run
    }

    $rootShare = ConvertTo-HealthCheckerAdminSharePath -Computer $Computer -LocalPath $Settings.RemoteRootPath
    $binShare = Join-Path -Path $rootShare -ChildPath 'Bin'
    $run.OutputSharePath = Join-Path -Path $rootShare -ChildPath 'Output'
    $driveShare = ($rootShare -split '\\')[0..3] -join '\'
    $shareReachable = $false
    try { $shareReachable = Test-Path -LiteralPath $driveShare -PathType Container } catch { $shareReachable = $false }
    if (-not $shareReachable) {
        return (& $fail 'RemoteConnectionFailed' 'The administrative share of the server is not reachable (SMB 445 and administrator rights on the server are required).')
    }
    if (-not (Test-Path -LiteralPath $binShare -PathType Container)) {
        return (& $fail 'TaskTriggerFailed' 'The Health Checker task is not installed on the server: run the Health Checker task installer.')
    }

    try {
        # Files on the server are refreshed when they differ; the runner checks Health Checker again before running it.
        [void](Copy-HealthCheckerRemoteFile -SourcePath $VerifiedScriptPath -DestinationPath (Join-Path -Path $binShare -ChildPath 'HealthChecker.ps1'))
        [void](Copy-HealthCheckerRemoteFile -SourcePath $RunnerSourcePath -DestinationPath (Join-Path -Path $binShare -ChildPath (Split-Path -Path $RunnerSourcePath -Leaf)))
    }
    catch {
        return (& $fail 'TaskTriggerFailed' ("Health Checker files could not be deployed to the server: {0}" -f $_.Exception.Message))
    }
    $run.PreviousStatus = Read-HealthCheckerRemoteStatus -OutputSharePath $run.OutputSharePath

    try {
        # DCOM, not WinRM: PowerShell remoting is disabled on the Exchange servers.
        $run.Session = New-CimSession -ComputerName $Computer -SessionOption (New-HealthCheckerDcomSessionOption) -ErrorAction Stop
    }
    catch {
        return (& $fail 'RemoteConnectionFailed' ("The server could not be reached over CIM/DCOM (RPC 135 and dynamic ports): {0}" -f $_.Exception.Message))
    }

    try {
        $task = Get-ScheduledTask -CimSession $run.Session -TaskPath $Settings.TaskPath -TaskName $Settings.TaskName -ErrorAction SilentlyContinue
        if ($null -eq $task) {
            return (& $fail 'TaskTriggerFailed' 'The Health Checker task is not registered on the server: run the Health Checker task installer.')
        }
        if ([string]$task.State -eq 'Running' -and $run.PreviousStatus -and [string]$run.PreviousStatus.State -eq 'Running') {
            # A run is already in progress (for example the daily one): its result is awaited instead of starting another.
            $run.WaitToken = [string]$run.PreviousStatus.RunToken
        }
        else {
            Start-ScheduledTask -CimSession $run.Session -TaskPath $Settings.TaskPath -TaskName $Settings.TaskName -ErrorAction Stop
        }
    }
    catch {
        return (& $fail 'TaskTriggerFailed' ("The Health Checker task could not be started: {0}" -f $_.Exception.Message))
    }
    return $run
}

function Update-HealthCheckerRemoteRun {
    param(
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][object]$Settings
    )

    $current = Read-HealthCheckerRemoteStatus -OutputSharePath $Run.OutputSharePath
    $previousToken = if ($Run.PreviousStatus) { [string]$Run.PreviousStatus.RunToken } else { '' }
    $currentToken = if ($current) { [string]$current.RunToken } else { '' }
    $isThisRun = if ($Run.WaitToken) { $currentToken -eq $Run.WaitToken } else { $currentToken -and $currentToken -ne $previousToken }
    if ($isThisRun -and [string]$current.State -ne 'Running') {
        $Run.Done = $true
        $Run.RemoteStatus = $current
        return
    }
    if ($isThisRun -or ((Get-Date) - $Run.StartedAt).TotalSeconds -lt 60) { return }

    # No new run after a minute: a logon failure means the runner never started.
    try {
        $info = Get-ScheduledTaskInfo -CimSession $Run.Session -TaskPath $Settings.TaskPath -TaskName $Settings.TaskName -ErrorAction Stop
        $logonFailure = Get-HealthCheckerTaskLogonFailure -LastTaskResult $info.LastTaskResult
        if ($logonFailure) {
            $Run.Done = $true
            $Run.Status = 'TaskTriggerFailed'
            $Run.Message = $logonFailure
            return
        }
        # After the grace period, a task that is no longer running without a fresh status.json never ran the
        # runner (rights, batch logon, password): stop waiting instead of using the whole per-server limit.
        $graceSeconds = if ($Settings.PSObject.Properties['TaskStartGraceSeconds']) { [int]$Settings.TaskStartGraceSeconds } else { 180 }
        if (((Get-Date) - $Run.StartedAt).TotalSeconds -ge $graceSeconds) {
            $task = Get-ScheduledTask -CimSession $Run.Session -TaskPath $Settings.TaskPath -TaskName $Settings.TaskName -ErrorAction Stop
            if ([string]$task.State -ne 'Running') {
                $code = [long]$info.LastTaskResult
                if ($code -lt 0) { $code += 4294967296 }
                $Run.Done = $true
                $Run.Status = 'TaskTriggerFailed'
                $Run.Message = ("The Health Checker task is not running and wrote no result after {0} min (last task result 0x{1:X8}). Check that the service account is a local administrator of the server, may log on as a batch job, and that its password is current." -f [math]::Round($graceSeconds / 60, 1), $code)
            }
        }
    }
    catch {
        Write-Verbose "Health Checker task information unavailable for $($Run.ServerName): $($_.Exception.Message)"
    }
}

function Complete-HealthCheckerRemoteRun {
    param(
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][string]$LocalFolder,
        [Parameter(Mandatory)][object]$Settings
    )

    if ($Run.Session) { Remove-CimSession -CimSession $Run.Session -ErrorAction SilentlyContinue; $Run.Session = $null }
    $outcome = [pscustomobject]@{ Status = $Run.Status; Message = $Run.Message; DurationSeconds = $null; XmlPath = $null; JsonPath = $null; ResultSource = $null; ResultAgeHours = $null }
    $resultStatus = $null
    if ($Run.RemoteStatus) {
        $mapped = ConvertFrom-HealthCheckerRunnerState -RemoteStatus $Run.RemoteStatus
        $outcome.Status = $mapped.Status
        $outcome.Message = $mapped.Message
        $outcome.DurationSeconds = $Run.RemoteStatus.DurationSeconds
        $outcome.ResultSource = 'OnDemand'
        if ($mapped.Status -eq 'Completed') { $resultStatus = $Run.RemoteStatus }
    }
    elseif ($outcome.Status -in @('RemoteConnectionFailed', 'TaskTriggerFailed') -and $Run.PreviousStatus) {
        $fallback = Resolve-HealthCheckerScheduledFallback -FailureStatus $outcome.Status -FailureMessage $outcome.Message -RemoteStatus $Run.PreviousStatus -NowUtc ([DateTime]::UtcNow) -MaxResultAgeHours $Settings.MaxResultAgeHours
        $outcome.Status = $fallback.Status
        $outcome.Message = $fallback.Message
        if ($fallback.UseResult) {
            $resultStatus = $Run.PreviousStatus
            $outcome.ResultSource = 'Scheduled'
            $outcome.DurationSeconds = $Run.PreviousStatus.DurationSeconds
        }
    }
    $outcome.Message = ConvertTo-HealthCheckerRemoteMessage -Message $outcome.Message -Computer $Run.Computer

    if ($resultStatus) {
        $finished = ConvertTo-HealthCheckerUtcTime -Value $resultStatus.FinishedAtUtc
        if ($finished) { $outcome.ResultAgeHours = [math]::Max(0, [math]::Round(([DateTime]::UtcNow - $finished).TotalHours, 1)) }
        # Only the plain JSON result is copied, never the Health Checker XML: deserializing XML from
        # another server in the collector's Exchange session would trust that server with this account.
        if ([string]$resultStatus.ResultFileName -cne 'result.json') {
            $outcome.Status = 'NoOutput'
            $outcome.Message = 'The Health Checker task reported no result file or an unexpected one.'
            $outcome.ResultSource = $null
            return $outcome
        }
        try {
            Ensure-Directory -Path $LocalFolder
            $localJson = Join-Path -Path $LocalFolder -ChildPath 'result.json'
            Copy-Item -LiteralPath (Join-Path -Path $Run.OutputSharePath -ChildPath 'result.json') -Destination $localJson -Force -ErrorAction Stop
            $outcome.JsonPath = $localJson
        }
        catch {
            $outcome.Status = 'RemoteConnectionFailed'
            $outcome.Message = ConvertTo-HealthCheckerRemoteMessage -Message ("The Health Checker result could not be copied from the server: {0}" -f $_.Exception.Message) -Computer $Run.Computer
            $outcome.ResultSource = $null
        }
    }
    return $outcome
}

function Invoke-ExchangeHealthCheckerInventory {
    param(
        [Parameter(Mandatory)][object[]]$ExchangeServers,
        [Parameter(Mandatory)][string]$SourceScriptPath,
        [Parameter(Mandatory)][string]$RunnerSourcePath,
        [Parameter(Mandatory)][string]$ExpectedSha256,
        [Parameter(Mandatory)][string]$HealthCheckerVersion,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][object]$Settings,
        [string]$LocalComputerName = $env:COMPUTERNAME
    )

    $summaryRows = New-Object 'System.Collections.Generic.List[object]'
    $findingRows = New-Object 'System.Collections.Generic.List[object]'
    $serverNames = @($ExchangeServers | ForEach-Object { [string]$_.Name })
    # Excluded servers keep a summary row but are never started, contacted or counted as failures.
    $excludedList = if ($Settings.PSObject.Properties['ExcludedServers']) { @($Settings.ExcludedServers) } else { @() }
    $excludedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($server in $ExchangeServers) {
        if (Test-HealthCheckerServerExcluded -Server $server -ExcludedServers $excludedList) { [void]$excludedNames.Add([string]$server.Name) }
    }
    $activeServers = @($ExchangeServers | Where-Object { -not $excludedNames.Contains([string]$_.Name) })
    if ($excludedNames.Count -gt 0) {
        Write-Log ("Health Checker excluded by configuration (HealthCheckerExcludedServers): {0}." -f (@($excludedNames) -join ', '))
    }
    $runFolder = Join-Path -Path $Settings.WorkFolderPath -ChildPath $RunId
    $blockedStatus = $null
    $blockedMessage = $null

    $runScriptPath = Join-Path -Path $runFolder -ChildPath 'HealthChecker.ps1'
    $integrity = $null
    if (-not (Test-HealthCheckerElevation)) {
        $blockedStatus = 'NotElevated'
        $blockedMessage = 'Health Checker was not run: it requires an elevated (administrator) session on the Exchange server.'
    }
    else {
        try { Initialize-HealthCheckerWorkFolder -WorkRoot $Settings.WorkFolderPath }
        catch {
            $blockedStatus = 'UnsafeWorkFolder'
            $blockedMessage = "Health Checker was not run: its work folder cannot be restricted to administrators. $($_.Exception.Message)"
        }
    }
    if (-not $blockedStatus) {
        Remove-HealthCheckerPreviousWorkFolder -WorkRoot $Settings.WorkFolderPath -CurrentRunId $RunId
        $integrity = Test-HealthCheckerScriptIntegrity -Path $SourceScriptPath -ExpectedSha256 $ExpectedSha256
        if ($integrity.IsValid) {
            Ensure-Directory -Path $runFolder
            # The copy that runs is verified again, so a change after the first check is never executed.
            [System.IO.File]::WriteAllBytes($runScriptPath, [System.IO.File]::ReadAllBytes($SourceScriptPath))
            $integrity = Test-HealthCheckerScriptIntegrity -Path $runScriptPath -ExpectedSha256 $ExpectedSha256
        }
        if (-not $integrity.IsValid) {
            $blockedStatus = 'IntegrityFailed'
            $blockedMessage = "Health Checker was not run: $($integrity.Reason)"
        }
    }
    if ($blockedStatus) {
        Write-Log $blockedMessage 'WARN'
    }
    elseif (-not (Test-Path -LiteralPath $RunnerSourcePath -PathType Leaf)) {
        $blockedStatus = 'IntegrityFailed'
        $blockedMessage = 'Health Checker was not run: the local runner deployed to remote servers is missing from the collector folder.'
        Write-Log $blockedMessage 'WARN'
    }
    else {
        Write-Log ("Health Checker {0} verified (SHA256 and Microsoft signature). Per-server limit: {1} min; total budget: {2} min." -f $HealthCheckerVersion, $Settings.TimeoutMinutesPerServer, $Settings.TotalBudgetMinutes)
    }

    $deadline = (Get-Date).AddMinutes($Settings.TotalBudgetMinutes)
    $outcomes = @{}
    if (-not $blockedStatus) {
        $remoteRuns = New-Object 'System.Collections.Generic.List[object]'
        # 1. Remote servers start first, through their scheduled task, so they all run in parallel.
        foreach ($server in @($activeServers | Where-Object { [string]$_.Name -ne $LocalComputerName })) {
            $serverName = [string]$server.Name
            $computer = Get-HealthCheckerServerComputerName -Server $server
            Write-Log ("Starting the Health Checker task on {0}." -f $serverName)
            $remoteRuns.Add((Start-HealthCheckerRemoteRun -ServerName $serverName -Computer $computer -VerifiedScriptPath $runScriptPath -RunnerSourcePath $RunnerSourcePath -Settings $Settings)) | Out-Null
        }

        # 2. The collector's own server runs in a local child process meanwhile.
        foreach ($server in @($activeServers | Where-Object { [string]$_.Name -eq $LocalComputerName })) {
            $serverName = [string]$server.Name
            $remainingSeconds = [int][math]::Floor(($deadline - (Get-Date)).TotalSeconds)
            if ($remainingSeconds -lt 60) {
                $outcomes[$serverName] = [pscustomobject]@{ Status = 'NotRun'; Message = 'Health Checker was not run on this server: the total Health Checker time budget was used up.'; DurationSeconds = $null; XmlPath = $null; JsonPath = $null; ResultSource = $null; ResultAgeHours = $null }
                continue
            }
            $timeoutSeconds = [math]::Min($Settings.TimeoutMinutesPerServer * 60, $remainingSeconds)
            $serverFolder = Join-Path -Path $runFolder -ChildPath ($serverName -replace '[^A-Za-z0-9._-]', '_')
            Write-Log ("Running Health Checker locally for {0} (limit {1} min)." -f $serverName, [math]::Round($timeoutSeconds / 60, 1))
            $local = Invoke-HealthCheckerProcess -ScriptPath $runScriptPath -ServerName $serverName -OutputFolder $serverFolder -TimeoutSeconds $timeoutSeconds -ExpectedSha256 $ExpectedSha256
            $outcomes[$serverName] = [pscustomobject]@{ Status = $local.Status; Message = $local.Message; DurationSeconds = $local.DurationSeconds; XmlPath = $local.XmlPath; JsonPath = $null; ResultSource = 'Local'; ResultAgeHours = 0 }
        }

        # 3. Remote results are awaited together, each within its own limit and the total budget.
        while (@($remoteRuns | Where-Object { -not $_.Done }).Count -gt 0) {
            foreach ($run in @($remoteRuns | Where-Object { -not $_.Done })) {
                $runDeadline = $run.StartedAt.AddMinutes($Settings.TimeoutMinutesPerServer)
                if ($deadline -lt $runDeadline) { $runDeadline = $deadline }
                Update-HealthCheckerRemoteRun -Run $run -Settings $Settings
                if (-not $run.Done -and (Get-Date) -ge $runDeadline) {
                    $run.Done = $true
                    $run.Status = 'TaskTimeout'
                    $run.Message = ("No result from the Health Checker task within {0} min." -f [math]::Round(($runDeadline - $run.StartedAt).TotalMinutes, 1))
                }
            }
            if (@($remoteRuns | Where-Object { -not $_.Done }).Count -gt 0) { Start-Sleep -Seconds $Settings.PollSeconds }
        }
        foreach ($run in $remoteRuns) {
            $serverFolder = Join-Path -Path $runFolder -ChildPath ($run.ServerName -replace '[^A-Za-z0-9._-]', '_')
            $outcomes[$run.ServerName] = Complete-HealthCheckerRemoteRun -Run $run -LocalFolder $serverFolder -Settings $Settings
        }
    }

    foreach ($serverName in $serverNames) {
        if ($excludedNames.Contains($serverName)) {
            $summaryRows.Add((ConvertTo-HealthCheckerSummaryRow -ServerName $serverName -RunStatus 'NotRun' -RunMessage 'Health Checker was not run: excluded by configuration.' -HealthCheckerVersion $HealthCheckerVersion -Parsed $null)) | Out-Null
            continue
        }
        if ($blockedStatus) {
            $summaryRows.Add((ConvertTo-HealthCheckerSummaryRow -ServerName $serverName -RunStatus $blockedStatus -RunMessage $blockedMessage -HealthCheckerVersion $HealthCheckerVersion -Parsed $null)) | Out-Null
            continue
        }
        $outcome = $outcomes[$serverName]
        $serverFolder = Join-Path -Path $runFolder -ChildPath ($serverName -replace '[^A-Za-z0-9._-]', '_')
        $parsed = $null
        $status = $outcome.Status
        $message = $outcome.Message
        if ($status -eq 'Completed') {
            try {
                $parsed = if ($outcome.JsonPath) {
                    # Another server's result: plain JSON lines only.
                    ConvertFrom-HealthCheckerResultLine -Lines @(Read-HealthCheckerResultJson -Path $outcome.JsonPath) -ServerName $serverName -HealthCheckerVersion $HealthCheckerVersion
                }
                else {
                    ConvertFrom-HealthCheckerXml -Path $outcome.XmlPath -ServerName $serverName -HealthCheckerVersion $HealthCheckerVersion
                }
                foreach ($finding in $parsed.Findings) { $findingRows.Add($finding) | Out-Null }
                Write-Log ("Health Checker for {0} ({1}): {2} finding(s), {3} vulnerability(ies), {4} error(s), {5} warning(s), {6}s." -f $serverName, $outcome.ResultSource, @($parsed.Findings).Count, @($parsed.Vulnerabilities).Count, $parsed.Facts.ErrorCount, $parsed.Facts.WarningCount, $outcome.DurationSeconds)
                if ($message) { Write-Log ("{0}: {1}" -f $serverName, $message) 'WARN' }
            }
            catch {
                $status = 'ParseFailed'
                $message = "Health Checker XML could not be read: $($_.Exception.Message)"
            }
        }

        if ($status -eq 'Completed') {
            if (Test-Path -LiteralPath $serverFolder -PathType Container) {
                try { Remove-Item -LiteralPath $serverFolder -Recurse -Force -ErrorAction Stop }
                catch { Write-Log ("Health Checker raw output could not be removed: {0}. {1}" -f $serverFolder, $_.Exception.Message) 'WARN' }
            }
        }
        else {
            $kept = if ($outcome.ResultSource -eq 'Local') { " Raw output kept locally until the next run: $serverFolder" } else { '' }
            Write-Log ("{0}: Health Checker {1}. {2}{3}" -f $serverName, $status, $message, $kept) 'WARN'
        }
        $summaryRows.Add((ConvertTo-HealthCheckerSummaryRow -ServerName $serverName -RunStatus $status -RunMessage $message -DurationSeconds $outcome.DurationSeconds -HealthCheckerVersion $HealthCheckerVersion -Parsed $parsed -ResultSource $outcome.ResultSource -ResultAgeHours $outcome.ResultAgeHours)) | Out-Null
    }

    if (Test-Path -LiteralPath $runScriptPath -PathType Leaf) { Remove-Item -LiteralPath $runScriptPath -Force -ErrorAction SilentlyContinue }
    if ((Test-Path -LiteralPath $runFolder -PathType Container) -and @(Get-ChildItem -LiteralPath $runFolder -Force).Count -eq 0) { Remove-Item -LiteralPath $runFolder -Force -ErrorAction SilentlyContinue }

    $processed = @($summaryRows | Where-Object { $_.RunStatus -eq 'Completed' }).Count
    foreach ($row in $summaryRows) {
        $row.ServersProcessed = $processed
        $row.ServersTotal = $activeServers.Count
    }
    Write-Log ("Health Checker processed {0} of {1} server(s); {2} excluded by configuration." -f $processed, $activeServers.Count, $excludedNames.Count)

    return [pscustomobject]@{
        Summary  = $summaryRows.ToArray()
        Findings = $findingRows.ToArray()
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
            $virtualDirectoryCommand = Get-Command $definition.Cmdlet -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $virtualDirectoryCommand) {
                Write-Log ("Skipping virtual directory collector because cmdlet is unavailable: {0}" -f $definition.Cmdlet)
                continue
            }

            $virtualDirectoryParameters = @{ ErrorAction = 'Stop' }
            if ($virtualDirectoryCommand.Parameters.ContainsKey('ADPropertiesOnly')) {
                $virtualDirectoryParameters.ADPropertiesOnly = $true
                Write-Log ("Collecting virtual directories with {0} -ADPropertiesOnly" -f $definition.Cmdlet)
            } else {
                Write-Log ("Collecting virtual directories with {0}; ADPropertiesOnly is not supported by this cmdlet." -f $definition.Cmdlet)
            }
            foreach ($item in @(& $definition.Cmdlet @virtualDirectoryParameters | Sort-Object Identity)) {
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

function ConvertTo-ServersAndStorageDatabaseSectionHtml {
    param(
        [AllowNull()][object[]]$ReadinessInventory = @(),
        [Parameter(Mandatory)][object]$Summary
    )

    # Use only this report's readiness rows, including when regenerating an old run.
    $databaseRows = @($ReadinessInventory | Where-Object { $null -ne $_ -and $_.Category -eq 'MailboxDatabase' })
    $collectorErrors = @($ReadinessInventory | Where-Object { $null -ne $_ -and $_.Category -eq 'CollectorError' -and $_.ObjectName -eq 'MailboxDatabase' })
    $unnamedRows = @($databaseRows | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.ObjectName) })
    $databaseGroups = @($databaseRows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.ObjectName) } | Group-Object ObjectName | Sort-Object Name)
    $mountedCount = 0
    $notMountedCount = 0
    $unknownCount = 0
    $detailRows = @(foreach ($group in $databaseGroups) {
        $values = @{}
        $notes = New-Object 'System.Collections.Generic.List[string]'
        foreach ($setting in @('Server', 'Mounted', 'DatabaseSize', 'AvailableNewMailboxSpace')) {
            $settingRows = @($group.Group | Where-Object { $_.Setting -eq $setting })
            $values[$setting] = $null
            if ($settingRows.Count -gt 1) {
                $notes.Add("${setting}: ambiguous duplicate records.")
            }
            elseif ($settingRows.Count -eq 0 -or [string]::IsNullOrWhiteSpace([string]$settingRows[0].Value)) {
                $notes.Add("${setting}: missing.")
            }
            elseif ($settingRows[0].CollectionStatus -ne 'OK') {
                $notes.Add("${setting}: collection not OK.")
            }
            else {
                $values[$setting] = [string]$settingRows[0].Value
            }
        }

        $mounted = $false
        $status = 'Unknown'
        $statusColor = '#64748b'
        if ([bool]::TryParse([string]$values['Mounted'], [ref]$mounted)) {
            if ($mounted) {
                $status = 'Mounted'; $mountedCount++; $statusColor = '#166534'
            }
            else {
                $status = 'Not mounted'; $notMountedCount++; $statusColor = '#9a3412'
            }
        }
        else {
            $unknownCount++
            if ($null -ne $values['Mounted']) { $notes.Add('Mounted: unrecognized value.') }
        }

        "<tr><td>$(Format-HtmlValue $group.Name)</td><td>$(Format-HtmlValue $values['Server'])</td><td style='font-weight:700;color:$statusColor;'>$status</td><td>$(Format-HtmlValue $values['DatabaseSize'])</td><td>$(Format-HtmlValue $values['AvailableNewMailboxSpace'])</td><td>$(Format-HtmlValue $(if ($notes.Count) { $notes -join ' ' } else { 'OK' }))</td></tr>"
    })

    $coverageNote = 'Counters cover observed distinct database names, including recovery databases. Missing fields are N/A; missing or unrecognized Mounted values are Unknown.'
    if ($collectorErrors.Count -gt 0) {
        $coverageNote += ' MailboxDatabase collection failed: coverage is incomplete; observed counts are not estate totals.'
    }
    if ($unnamedRows.Count -gt 0) {
        $coverageNote += " $($unnamedRows.Count) record(s) have no database name and cannot be counted or associated. Coverage is incomplete."
    }
    $totalCount = $databaseGroups.Count
    if ($totalCount -eq 0) {
        $totalCount = 'N/A'; $mountedCount = 'N/A'; $notMountedCount = 'N/A'; $unknownCount = 'N/A'
        $coverageNote += ' No named database records are available; this does not prove that there are no databases.'
        $detailRows = @('<tr><td colspan="6">No database details available (N/A).</td></tr>')
    }

@"
      <h2>Mailbox databases</h2>
      <div class="summary">
        Total databases observed: <strong>$totalCount</strong> |
        Mounted: <strong>$mountedCount</strong> |
        Not mounted: <strong>$notMountedCount</strong> |
        Unknown status: <strong>$unknownCount</strong>.<br>
        $(Format-HtmlValue $coverageNote)<br>
        Source: <strong>Exchange_OnPrem_MigrationReadiness_Config.csv</strong>, category MailboxDatabase, collected with Get-MailboxDatabase -Status.
        Source run: <strong>$(Format-HtmlValue $Summary.RunId)</strong>; run date: <strong>$(Format-HtmlValue $Summary.ExecutionDate)</strong>. These values describe that run, not the time the email is read.
      </div>
      <table>
        <thead><tr><th>Database</th><th>Server (reported)</th><th>Mounted status</th><th>Database size</th><th>Internal reusable space</th><th>Data notes</th></tr></thead>
        <tbody>$($detailRows -join "`r`n")</tbody>
      </table>
      <div class="summary">
        Sizes retain the units returned by Exchange. AvailableNewMailboxSpace is reusable space inside the database, not free disk space.<br>
        Server is the association reported by Get-MailboxDatabase; detailed DAG copy states and per-copy servers are not collected and are not inferred here.<br>
        Mailboxes per database: N/A. Mailboxes are collected separately; this report has no qualified cross-collection database match with mailbox provenance and freshness. Zero observed mailboxes would not prove an empty database.
      </div>
"@
}

function ConvertTo-ServersAndStorageHealthCheckerSectionHtml {
    param(
        [AllowNull()][object[]]$HealthCheckerSummary = @(),
        [AllowNull()][object[]]$HealthCheckerFindings = @()
    )

    # No section when the run did not include Health Checker.
    $summaryRows = @($HealthCheckerSummary | Where-Object { $null -ne $_ })
    if ($summaryRows.Count -eq 0) { return '' }

    $toInt = { param($Value) $number = 0; if ([int]::TryParse([string]$Value, [ref]$number)) { $number } else { 0 } }
    $processed = @($summaryRows | Where-Object { [string]$_.RunStatus -eq 'Completed' }).Count
    $vulnerabilityTotal = 0; $errorTotal = 0; $warningTotal = 0
    foreach ($row in $summaryRows) {
        $vulnerabilityTotal += & $toInt $row.VulnerabilityCount
        $errorTotal += & $toInt $row.ErrorCount
        $warningTotal += & $toInt $row.WarningCount
    }
    $healthCheckerVersion = @($summaryRows | ForEach-Object { [string]$_.HealthCheckerVersion } | Where-Object { $_ } | Select-Object -Unique) -join ', '

    $sortedServers = @($summaryRows | Sort-Object @{ Expression = { & $toInt $_.VulnerabilityCount }; Descending = $true }, @{ Expression = { & $toInt $_.ErrorCount }; Descending = $true }, @{ Expression = { & $toInt $_.WarningCount }; Descending = $true }, ServerName)
    $serverRowsHtml = foreach ($row in $sortedServers) {
        $runStatus = [string]$row.RunStatus
        $runClass = if ($runStatus -eq 'Completed') { 'status-ok' } else { 'status-warning' }
        $vulnerabilityCount = & $toInt $row.VulnerabilityCount
        $vulnerabilityClass = if ($vulnerabilityCount -gt 0) { 'status-error' } else { 'status-ok' }
        $suStatus = [string]$row.SecurityUpdateStatus
        $suClass = switch ($suStatus) { 'Latest' { 'status-ok'; break } 'NewerPatchAvailable' { 'status-review'; break } 'Unknown' { 'status-review'; break } default { 'status-error' } }
        $vulnerabilityCell = if ($runStatus -eq 'Completed') { "<span class='badge $vulnerabilityClass'>$(Format-HtmlValue $vulnerabilityCount)</span>" } else { 'N/A' }
        "<tr><td>$(Format-HtmlValue $row.ServerName)</td><td><span class='badge $runClass'>$(Format-HtmlValue $runStatus)</span></td><td>$(Format-HtmlValue $row.ExchangeVersion)</td><td>$(Format-HtmlValue $row.BuildNumber)</td><td><span class='badge $suClass'>$(Format-HtmlValue $suStatus)</span></td><td class='num'>$vulnerabilityCell</td><td class='num'>$(Format-HtmlValue $row.ErrorCount)</td><td class='num'>$(Format-HtmlValue $row.WarningCount)</td><td>$(Format-HtmlValue $(if ($runStatus -eq 'Completed') { $row.Vulnerabilities } else { $row.RunMessage }))</td></tr>"
    }

    $errorFindings = @($HealthCheckerFindings | Where-Object { $null -ne $_ -and [string]$_.Status -eq 'Error' } | Sort-Object @{ Expression = { if ([string]$_.Category -eq 'Security Vulnerability') { 0 } else { 1 } } }, ServerName, @{ Expression = { & $toInt $_.RowNumber } })
    $shownErrors = @($errorFindings | Select-Object -First 100)
    $errorRowsHtml = foreach ($finding in $shownErrors) {
        $checkText = if ([string]::IsNullOrWhiteSpace([string]$finding.Item)) { $finding.Check } else { "$($finding.Check) - $($finding.Item)" }
        "<tr><td>$(Format-HtmlValue $finding.ServerName)</td><td>$(Format-HtmlValue $finding.Category)</td><td>$(Format-HtmlValue $checkText)</td><td class='cell-value'>$(Format-HtmlValue $finding.Value)</td><td>$(Format-HtmlValue $finding.Recommendation)</td></tr>"
    }
    $errorsBlock = if ($shownErrors.Count -gt 0) {
        $truncationNote = if ($errorFindings.Count -gt $shownErrors.Count) { "<div class='summary'>Showing the first $($shownErrors.Count) of $($errorFindings.Count) errors; all findings are in Exchange_OnPrem_HealthChecker_Findings.csv.</div>" } else { '' }
@"
      <h3>Vulnerabilities and errors <span>$(Format-HtmlValue $errorFindings.Count) records</span></h3>
      <table class="readiness-table">
        <thead><tr><th>Server</th><th>Area</th><th>Check</th><th>Value</th><th>More information</th></tr></thead>
        <tbody>$($errorRowsHtml -join "`r`n")</tbody>
      </table>
      $truncationNote
"@
    } else { "<div class='summary'>No Health Checker error was reported for the processed servers.</div>" }

@"
      <h2>Exchange Health Checker</h2>
      <div class="summary">
        Microsoft CSS-Exchange Health Checker <strong>$(Format-HtmlValue $healthCheckerVersion)</strong> processed <strong>$processed</strong> of <strong>$($summaryRows.Count)</strong> server(s).
        Vulnerabilities: <strong>$vulnerabilityTotal</strong> | Errors: <strong>$errorTotal</strong> | Warnings: <strong>$warningTotal</strong>.<br>
        Servers that were not processed show N/A and the reason; their state is not measured. Values that can reveal host names, addresses, certificates or domains are masked in the CSV and in this report.
      </div>
      <table>
        <thead><tr><th>Server</th><th>Run</th><th>Exchange version</th><th>Build</th><th>Security update</th><th class="num">Vulnerabilities</th><th class="num">Errors</th><th class="num">Warnings</th><th>Vulnerabilities / reason</th></tr></thead>
        <tbody>$($serverRowsHtml -join "`r`n")</tbody>
      </table>
      $errorsBlock
"@
}

function New-HtmlExecutiveSummary {
    param(
        [Parameter(Mandatory)]
        [object]$Summary,

        [Parameter(Mandatory)]
        [object[]]$PerServerSummary,

        [AllowNull()]
        [object[]]$ReadinessInventory = @(),

        [AllowNull()]
        [object[]]$HealthCheckerSummary = @(),

        [AllowNull()]
        [object[]]$HealthCheckerFindings = @(),

        [Parameter(Mandatory)]
        [string]$Path
    )

    $excludedServerNames = @()
    $excludedServersProperty = $Summary.PSObject.Properties['ExcludedServers']
    if ($null -ne $excludedServersProperty -and -not [string]::IsNullOrWhiteSpace([string]$excludedServersProperty.Value)) {
        $excludedServerNames = @(([string]$excludedServersProperty.Value -split ';') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $rowsHtml = @(foreach ($row in ($PerServerSummary | Sort-Object ExchangeServerName)) {
        $computeStatus = [string]$row.ComputeCollectionStatus
        $statusClass = switch ($computeStatus) {
            'ERROR' { 'status-error'; break }
            'WARNING' { 'status-warning'; break }
            default { 'status-ok' }
        }
        $statusColor = switch ($computeStatus) {
            'ERROR' { '#b91c1c'; break }
            'WARNING' { '#9a3412'; break }
            default { '#047857' }
        }
        "<tr><td>$(Format-HtmlValue $row.ExchangeServerName)</td><td>$(Format-HtmlValue $row.ServerRole)</td><td class='num'>$(Format-HtmlValue $row.LogicalProcessorCount)</td><td class='num'>$(Format-HtmlValue $row.MemoryGB)</td><td class='num'>$(Format-HtmlValue $row.DiskDriveCount)</td><td class='num'>$(Format-HtmlValue $row.DiskDriveTotalSizeGB)</td><td class='status $statusClass' style='font-weight:700;color:$statusColor;'>$(Format-HtmlValue $computeStatus)</td></tr>"
    })
    foreach ($serverName in $excludedServerNames) {
        $rowsHtml += "<tr><td>$(Format-HtmlValue $serverName)</td><td>N/A</td><td class='num'>N/A</td><td class='num'>N/A</td><td class='num'>N/A</td><td class='num'>N/A</td><td class='status' style='font-weight:700;color:#64748b;'>Excluded</td></tr>"
    }
    $exclusionNote = if ($excludedServerNames.Count -gt 0) {
        "<div class='summary'>Excluded by configuration: <strong>$(Format-HtmlValue $excludedServerNames.Count)</strong> server(s). Their status is <strong>Excluded</strong>; no per-server collection was attempted.</div>"
    } else { '' }

    $readinessRows = @($ReadinessInventory)
    $readinessErrorCount = @($readinessRows | Where-Object { $_.CollectionStatus -eq "ERROR" -or $_.Importance -eq "Error" }).Count
    $readinessWarningCount = @($readinessRows | Where-Object { $_.CollectionStatus -eq "WARNING" -or $_.Importance -eq "Warning" }).Count
    $readinessSortedRows = @($readinessRows | Sort-Object @{ Expression = { if ($_.CollectionStatus -eq "ERROR" -or $_.Importance -eq "Error") { 0 } elseif ($_.CollectionStatus -eq "WARNING" -or $_.Importance -eq "Warning") { 1 } elseif ($_.Importance -eq "Review") { 2 } else { 3 } } }, Category, ObjectName, Setting)
    $priorityRows = New-Object System.Collections.ArrayList
    foreach ($server in @($PerServerSummary | Sort-Object ExchangeServerName)) {
        if ([string]$server.ComputeCollectionStatus -eq 'ERROR') {
            [void]$priorityRows.Add([pscustomobject]@{ Severity = 'Error'; Category = 'Infrastructure'; ObjectName = $server.ExchangeServerName; Issue = 'Compute collection'; Status = $server.ComputeCollectionStatus; Details = 'CPU/RAM inventory failed or returned no data for this server.' })
        }
        if ([string]$server.DiskDriveCollectionStatus -eq 'ERROR') {
            [void]$priorityRows.Add([pscustomobject]@{ Severity = 'Error'; Category = 'Infrastructure'; ObjectName = $server.ExchangeServerName; Issue = 'Disk drive collection'; Status = $server.DiskDriveCollectionStatus; Details = 'Physical disk inventory failed or returned errors for this server.' })
        }
        if ([string]$server.LogicalDiskCollectionStatus -eq 'ERROR') {
            [void]$priorityRows.Add([pscustomobject]@{ Severity = 'Error'; Category = 'Infrastructure'; ObjectName = $server.ExchangeServerName; Issue = 'Logical disk collection'; Status = $server.LogicalDiskCollectionStatus; Details = 'Logical disk inventory failed or returned errors for this server.' })
        }
        if ([int]$server.LowSpaceLogicalDiskCount -gt 0) {
            [void]$priorityRows.Add([pscustomobject]@{ Severity = 'Warning'; Category = 'Capacity'; ObjectName = $server.ExchangeServerName; Issue = 'Low disk space'; Status = 'WARNING'; Details = ('{0} logical disk(s) below the configured free-space threshold.' -f $server.LowSpaceLogicalDiskCount) })
        }
    }
    foreach ($row in @($readinessSortedRows | Where-Object { $_.CollectionStatus -eq 'ERROR' -or $_.CollectionStatus -eq 'WARNING' -or $_.Importance -eq 'Error' -or $_.Importance -eq 'Warning' })) {
        $severity = if ($row.CollectionStatus -eq 'ERROR' -or $row.Importance -eq 'Error') { 'Error' } else { 'Warning' }
        $details = if (-not [string]::IsNullOrWhiteSpace([string]$row.ErrorMessage)) { $row.ErrorMessage } else { $row.MigrationFocus }
        [void]$priorityRows.Add([pscustomobject]@{ Severity = $severity; Category = $row.Category; ObjectName = $row.ObjectName; Issue = $row.Setting; Status = $row.CollectionStatus; Details = $details })
    }

    $priorityRowsSorted = @($priorityRows | Sort-Object @{ Expression = { if ($_.Severity -eq 'Error') { 0 } else { 1 } } }, Category, ObjectName, Issue | Select-Object -First 100)
    $priorityIssuesSection = if ($priorityRowsSorted.Count -gt 0) {
        $priorityRowsHtml = foreach ($row in $priorityRowsSorted) {
            $severityClass = if ($row.Severity -eq 'Error') { 'status-error' } else { 'status-warning' }
            "<tr><td><span class='badge $severityClass'>$(Format-HtmlValue $row.Severity)</span></td><td>$(Format-HtmlValue $row.Category)</td><td>$(Format-HtmlValue $row.ObjectName)</td><td>$(Format-HtmlValue $row.Issue)</td><td><span class='badge $severityClass'>$(Format-HtmlValue $row.Status)</span></td><td>$(Format-HtmlValue $row.Details)</td></tr>"
        }
@"
      <h2>Priority issues summary</h2>
      <table class="priority-table">
        <thead>
          <tr>
            <th>Severity</th>
            <th>Area</th>
            <th>Object</th>
            <th>Issue</th>
            <th>Status</th>
            <th>Details</th>
          </tr>
        </thead>
        <tbody>
          $($priorityRowsHtml -join "`r`n")
        </tbody>
      </table>
"@
    } else {
@"
      <h2>Priority issues summary</h2>
      <div class="summary">No priority infrastructure or readiness errors were detected in this run.</div>
"@
    }

    $statusSummaryRows = foreach ($categoryGroup in ($readinessRows | Group-Object Category | Sort-Object Name)) {
        $groupRows = @($categoryGroup.Group)
        [pscustomobject]@{
            Category         = $categoryGroup.Name
            Total            = $groupRows.Count
            Ok               = @($groupRows | Where-Object { $_.CollectionStatus -eq 'OK' -and $_.Importance -notin @('Review', 'Warning', 'Error') }).Count
            Review           = @($groupRows | Where-Object { $_.Importance -eq 'Review' -and $_.CollectionStatus -eq 'OK' }).Count
            Warning          = @($groupRows | Where-Object { $_.CollectionStatus -eq 'WARNING' -or $_.Importance -eq 'Warning' }).Count
            Error            = @($groupRows | Where-Object { $_.CollectionStatus -eq 'ERROR' -or $_.Importance -eq 'Error' }).Count
            CollectionErrors = @($groupRows | Where-Object { $_.CollectionStatus -eq 'ERROR' }).Count
        }
    }
    $statusSummarySection = if (@($statusSummaryRows).Count -gt 0) {
        $statusRowsHtml = foreach ($row in $statusSummaryRows) {
            "<tr><td>$(Format-HtmlValue $row.Category)</td><td class='num'>$(Format-HtmlValue $row.Total)</td><td class='num'>$(Format-HtmlValue $row.Ok)</td><td class='num'>$(Format-HtmlValue $row.Review)</td><td class='num'>$(Format-HtmlValue $row.Warning)</td><td class='num'>$(Format-HtmlValue $row.Error)</td><td class='num'>$(Format-HtmlValue $row.CollectionErrors)</td></tr>"
        }
@"
      <h2>Readiness status summary</h2>
      <table class="status-summary-table">
        <thead>
          <tr>
            <th>Area</th>
            <th class="num">Total</th>
            <th class="num">OK</th>
            <th class="num">Review</th>
            <th class="num">Warnings</th>
            <th class="num">Errors</th>
            <th class="num">Collection errors</th>
          </tr>
        </thead>
        <tbody>
          $($statusRowsHtml -join "`r`n")
        </tbody>
      </table>
"@
    } else { '' }
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
            <th>Readiness focus</th>
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
      <h2>Exchange readiness configuration</h2>
      <div class="summary">
        Collected <strong>$(Format-HtmlValue $readinessRows.Count)</strong> Exchange configuration readiness records for Exchange SE and Exchange Online planning. AD schema rangeUpper: <strong>$(Format-HtmlValue $Summary.ExchangeSchemaRangeUpper)</strong>. Exchange organization objectVersion: <strong>$(Format-HtmlValue $Summary.ExchangeOrgObjectVersion)</strong>. Review warnings: <strong>$(Format-HtmlValue $readinessWarningCount)</strong>. Collection errors: <strong>$(Format-HtmlValue $readinessErrorCount)</strong>. Every collected readiness record is rendered below and also attached as CSV.
      </div>
      $($readinessCategoryBlocks -join "`r`n")
"@
    } else {
        ""
    }
    $databaseSection = ConvertTo-ServersAndStorageDatabaseSectionHtml -ReadinessInventory $ReadinessInventory -Summary $Summary
    $healthCheckerSection = if (@($HealthCheckerSummary | Where-Object { $null -ne $_ }).Count -gt 0) { ConvertTo-ServersAndStorageHealthCheckerSectionHtml -HealthCheckerSummary $HealthCheckerSummary -HealthCheckerFindings $HealthCheckerFindings } else { '' }
    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Exchange On-Premises Infrastructure and Readiness</title>
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
<div class="shell" style="max-width:1180px;margin:0 auto;padding:24px;font-family:Segoe UI,Arial,sans-serif;color:#0f172a;">
  <div class="panel" style="background:#ffffff;border:1px solid #dbe3ef;border-radius:10px;overflow:hidden;">
    <div class="hero" style="background:#0f766e;color:#ffffff;padding:24px 28px;">
      <div class="eyebrow" style="font-size:12px;text-transform:uppercase;letter-spacing:.08em;opacity:.9;font-weight:700;color:#ffffff;">Smart365 Exchange OnPrem</div>
      <h1 style="font-size:26px;line-height:1.25;margin:6px 0 0 0;color:#ffffff;">Exchange On-Premises Infrastructure and Readiness</h1>
      <div class="subtitle" style="margin-top:10px;font-size:13px;opacity:.95;color:#ffffff;">Executive summary generated on $(Format-HtmlValue $Summary.ExecutionDate) - RunId $(Format-HtmlValue $Summary.RunId)</div>
    </div>

    <div class="content" style="padding:24px 28px;">
      <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:separate;border-spacing:10px;border:0;margin:0 0 22px 0;font-size:12px;">
        <tr>
          <td style="width:20%;vertical-align:top;border:1px solid #dbe3ef;border-radius:8px;padding:14px;background:#f8fafc;color:#334155;"><div style="font-size:11px;text-transform:uppercase;color:#64748b;font-weight:800;">Exchange VMs</div><div style="font-size:28px;font-weight:800;margin-top:4px;color:#0f172a;">$(Format-HtmlValue $Summary.ExchangeServersCount)</div><div style="font-size:12px;color:#64748b;margin-top:2px;">servers collected</div></td>
          <td style="width:20%;vertical-align:top;border:1px solid #bfdbfe;border-radius:8px;padding:14px;background:#eff6ff;color:#334155;"><div style="font-size:11px;text-transform:uppercase;color:#64748b;font-weight:800;">vCPU</div><div style="font-size:28px;font-weight:800;margin-top:4px;color:#1d4ed8;">$(Format-HtmlValue $Summary.TotalLogicalProcessorCount)</div><div style="font-size:12px;color:#64748b;margin-top:2px;">logical processors</div></td>
          <td style="width:20%;vertical-align:top;border:1px solid #bbf7d0;border-radius:8px;padding:14px;background:#f0fdf4;color:#334155;"><div style="font-size:11px;text-transform:uppercase;color:#64748b;font-weight:800;">RAM</div><div style="font-size:28px;font-weight:800;margin-top:4px;color:#166534;">$(Format-HtmlValue $Summary.TotalMemoryGB)</div><div style="font-size:12px;color:#64748b;margin-top:2px;">GB</div></td>
          <td style="width:20%;vertical-align:top;border:1px solid #fed7aa;border-radius:8px;padding:14px;background:#fff7ed;color:#334155;"><div style="font-size:11px;text-transform:uppercase;color:#64748b;font-weight:800;">Disks</div><div style="font-size:28px;font-weight:800;margin-top:4px;color:#9a3412;">$(Format-HtmlValue $Summary.TotalDiskDriveCount)</div><div style="font-size:12px;color:#64748b;margin-top:2px;">WMI disk drives</div></td>
          <td style="width:20%;vertical-align:top;border:1px solid #e9d5ff;border-radius:8px;padding:14px;background:#faf5ff;color:#334155;"><div style="font-size:11px;text-transform:uppercase;color:#64748b;font-weight:800;">Provisioned Storage</div><div style="font-size:28px;font-weight:800;margin-top:4px;color:#7e22ce;">$(Format-HtmlValue $Summary.TotalDiskDriveSizeTB)</div><div style="font-size:12px;color:#64748b;margin-top:2px;">TB</div></td>
        </tr>
      </table>
      <div class="summary">
        The collected Exchange on-premises footprint represents <strong>$(Format-HtmlValue $Summary.ExchangeServersCount) virtual machines</strong>, <strong>$(Format-HtmlValue $Summary.TotalLogicalProcessorCount) vCPU</strong>, <strong>$(Format-HtmlValue $Summary.TotalMemoryGB) GB RAM</strong>, and <strong>$(Format-HtmlValue $Summary.TotalDiskDriveCount) disks</strong>. The full readiness section below supports Exchange SE preparation, Exchange Online coexistence, dependency review, and decommissioning sign-off.
      </div>
      $exclusionNote

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

      $priorityIssuesSection

      $healthCheckerSection

      $databaseSection

      $statusSummarySection

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

function Import-ServersAndStorageExistingCsv {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Description
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required $Description CSV not found: $Path"
    }

    return @(Import-Csv -LiteralPath $Path -Delimiter ';')
}

function Invoke-ServersAndStorageHtmlReportRegeneration {
    param(
        [Parameter(Mandatory)]
        [string]$RunFolder,

        [switch]$SendMail
    )

    if (-not (Test-Path -LiteralPath $RunFolder -PathType Container)) {
        throw "Existing DATA-ALL run folder not found: $RunFolder"
    }

    $summaryPath = Join-Path $RunFolder "Exchange_OnPrem_Servers_Inventory_Summary.csv"
    $perServerSummaryPath = Join-Path $RunFolder "Exchange_OnPrem_Infrastructure_PerServerSummary.csv"
    $exchangeReadinessPath = Join-Path $RunFolder "Exchange_OnPrem_MigrationReadiness_Config.csv"
    $htmlSummaryPath = Join-Path $RunFolder "Exchange_OnPrem_InfrastructureAndReadiness_Report.html"

    Write-Log "Regenerating HTML executive summary from existing DATA-ALL CSV files."
    Write-Log "Existing run folder: $RunFolder"

    $summaryRows = @(Import-ServersAndStorageExistingCsv -Path $summaryPath -Description 'global summary')
    if ($summaryRows.Count -lt 1) { throw "Global summary CSV is empty: $summaryPath" }

    $summary = $summaryRows | Select-Object -First 1
    $perServerSummary = @(Import-ServersAndStorageExistingCsv -Path $perServerSummaryPath -Description 'per-server infrastructure summary')
    $exchangeReadinessInventory = @(Import-ServersAndStorageExistingCsv -Path $exchangeReadinessPath -Description 'Exchange readiness configuration')
    # Health Checker CSVs exist only for runs made with -IncludeHealthChecker.
    $healthCheckerSummaryPath = Join-Path $RunFolder "Exchange_OnPrem_HealthChecker_Summary.csv"
    $healthCheckerFindingsPath = Join-Path $RunFolder "Exchange_OnPrem_HealthChecker_Findings.csv"
    $healthCheckerSummary = if (Test-Path -LiteralPath $healthCheckerSummaryPath -PathType Leaf) { @(Import-ServersAndStorageExistingCsv -Path $healthCheckerSummaryPath -Description 'Health Checker summary') } else { @() }
    $healthCheckerFindings = if (Test-Path -LiteralPath $healthCheckerFindingsPath -PathType Leaf) { @(Import-ServersAndStorageExistingCsv -Path $healthCheckerFindingsPath -Description 'Health Checker findings') } else { @() }

    New-HtmlExecutiveSummary -Summary $summary -PerServerSummary @($perServerSummary) -ReadinessInventory @($exchangeReadinessInventory) -HealthCheckerSummary @($healthCheckerSummary) -HealthCheckerFindings @($healthCheckerFindings) -Path $htmlSummaryPath
    Write-Log "HTML executive summary regenerated to: $htmlSummaryPath"
    Invoke-ServersAndStorageSharePointUpload -LocalFilePath $htmlSummaryPath

    if ($SendMail) {
        $subjectRunId = if (-not [string]::IsNullOrWhiteSpace([string]$summary.RunId)) { [string]$summary.RunId } else { Split-Path -Path $RunFolder -Leaf }
        $mailSubject = "SmartM365 Exchange OnPrem Servers and Storage inventory - $subjectRunId"
        Send-ServersAndStorageHtmlReport `
            -HtmlReportPath $htmlSummaryPath `
            -Subject $mailSubject `
            -Attachments @($summaryPath, $perServerSummaryPath, $exchangeReadinessPath) `
            -Summary $summary `
            -PerServerSummary @($perServerSummary) `
            -ReadinessInventory @($exchangeReadinessInventory)
        Write-Log "Regenerated HTML executive summary email sent."
    }
    else {
        Write-Log "Regenerated report email was skipped. Use -SendRegeneratedReportMail to send it."
    }
}
$StartTime = Get-Date

$script:CompletionStatus = 'Auto'
try {
    Ensure-Directory -Path $OutputFolder
    Ensure-Directory -Path $LogFolder
    Start-SmartM365RunTranscript -Path $TranscriptFile -Append | Out-Null
    $script:TranscriptStarted = $true
    Write-Log "Starting $ScriptName v$ScriptVersion. RunId: $RunId"
    Write-Log "Output folder: $OutputFolder"

    if (-not [string]::IsNullOrWhiteSpace($RegenerateHtmlReportFromFolder)) {
        Write-Log "Mode: HTML report regeneration from existing DATA-ALL CSV files. Exchange collection is skipped."
        Invoke-ServersAndStorageHtmlReportRegeneration -RunFolder $OutputFolder -SendMail:$SendRegeneratedReportMail
        Write-Log "Completed regenerated report workflow."
        $completionStatus = if ($script:ServersAndStorageWarningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
        $script:CompletionStatus = $completionStatus
        Complete-ServersAndStorageRun -Status $completionStatus -Started $StartTime -ErrorMessage $null
        if ($completionStatus -eq 'CompletedWithWarnings') { $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(3); exit $smartM365RunGuard.ExitCode }
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    Write-Log "Collection method: WMI/DCOM only"

    $configuredExcludedServers = @(Get-ConfiguredExcludedServers)

    Import-SmartM365CoreModule
    Start-CoreSmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath ([string](Get-SmartM365EffectiveConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')) -ScopeParameters @{IncludeServicesHealth=$IncludeServicesHealth;IncludeMailboxDatabasePaths=$IncludeMailboxDatabasePaths;IncludeHealthChecker=$IncludeHealthChecker;ExcludedServers=($configuredExcludedServers -join ';')}
    Invoke-CoreSmartM365Preflight -ScriptName $ScriptName -OutputPaths @($OutputFolder,$LogFolder) | Out-Null

    Test-ExchangeShell
    Invoke-CoreSmartM365Preflight -ScriptName $ScriptName -RequireExchangeOnPrem | Out-Null

    Write-Log "Collecting Exchange servers."
    $discoveredExchangeServers = @(Get-ExchangeServer | Sort-Object Name)

    if ($discoveredExchangeServers.Count -eq 0) {
        throw "No Exchange servers were returned by Get-ExchangeServer."
    }

    $excludedSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($serverName in $configuredExcludedServers) { [void]$excludedSet.Add($serverName) }
    $excludedExchangeServers = @($discoveredExchangeServers | Where-Object { $excludedSet.Contains([string]$_.Name) -or $excludedSet.Contains([string]$_.Fqdn) })
    $exchangeServers = @($discoveredExchangeServers | Where-Object { -not ($excludedSet.Contains([string]$_.Name) -or $excludedSet.Contains([string]$_.Fqdn)) })
    Write-Log ("Exchange servers discovered: {0}; excluded by configuration: {1}; included in collection: {2}." -f $discoveredExchangeServers.Count, $excludedExchangeServers.Count, $exchangeServers.Count)
    foreach ($serverName in $configuredExcludedServers) { Write-Log ("Configured exclusion: {0} (Excluded)" -f $serverName) }
    if ($exchangeServers.Count -eq 0) {
        throw "No Exchange servers remain after applying ExcludedServers."
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

    $healthCheckerSummary = @()
    $healthCheckerFindings = @()
    if ($IncludeHealthChecker) {
        Write-Log "Running Microsoft CSS-Exchange Health Checker ${HealthCheckerVersion}: locally on this server, through the scheduled task on the other servers."
        $healthCheckerSettings = $null
        try { $healthCheckerSettings = Get-HealthCheckerSetting }
        catch {
            # A wrong Health Checker setting skips Health Checker only; the inventory continues.
            $configurationMessage = "Health Checker was not run: invalid configuration. $($_.Exception.Message)"
            Write-Log $configurationMessage 'WARN'
            $healthCheckerSummary = @(foreach ($server in $exchangeServers) {
                $row = ConvertTo-HealthCheckerSummaryRow -ServerName ([string]$server.Name) -RunStatus 'NotRun' -RunMessage $configurationMessage -HealthCheckerVersion $HealthCheckerVersion -Parsed $null
                $row.ServersProcessed = 0
                $row.ServersTotal = $exchangeServers.Count
                $row
            })
        }
        if ($healthCheckerSettings) {
            $healthCheckerResult = Invoke-ExchangeHealthCheckerInventory -ExchangeServers @($exchangeServers) -SourceScriptPath (Join-Path -Path $PSScriptRoot -ChildPath $HealthCheckerRelativePath) -RunnerSourcePath (Join-Path -Path $PSScriptRoot -ChildPath $HealthCheckerRunnerFileName) -ExpectedSha256 $HealthCheckerExpectedSha256 -HealthCheckerVersion $HealthCheckerVersion -RunId $RunId -Settings $healthCheckerSettings
            $healthCheckerSummary = @($healthCheckerResult.Summary)
            $healthCheckerFindings = @($healthCheckerResult.Findings)
        }
        $healthCheckerSummaryPath = Join-Path $OutputFolder "Exchange_OnPrem_HealthChecker_Summary.csv"
        Export-ServersAndStorageCsv -Data @($healthCheckerSummary) -Path $healthCheckerSummaryPath -Columns (Get-HealthCheckerCsvColumn -Kind Summary)
        $healthCheckerFindingsPath = Join-Path $OutputFolder "Exchange_OnPrem_HealthChecker_Findings.csv"
        Export-ServersAndStorageCsv -Data @($healthCheckerFindings) -Path $healthCheckerFindingsPath -Columns (Get-HealthCheckerCsvColumn -Kind Findings)
        Write-Log "Health Checker summary and findings exported to: $healthCheckerSummaryPath; $healthCheckerFindingsPath"
    }

    Write-Log "Collecting Exchange schema and readiness configuration."
    $exchangeReadinessInventory = @(Get-ExchangeReadinessInventory -ExchangeServerNames @($exchangeServers.Name))
    $exchangeReadinessPath = Join-Path $OutputFolder "Exchange_OnPrem_MigrationReadiness_Config.csv"
    Export-ServersAndStorageCsv -Data @($exchangeReadinessInventory) -Path $exchangeReadinessPath
    Write-Log "Exchange schema and readiness configuration exported to: $exchangeReadinessPath"

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
    $reportedWarningCount = $exchangeReadinessWarningRows.Count + $lowSpaceRows.Count
    if ($reportedWarningCount -gt $script:ServersAndStorageWarningCount) {
        $script:ServersAndStorageWarningCount = $reportedWarningCount
    }
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
        DiscoveredExchangeServersCount   = $discoveredExchangeServers.Count
        ExcludedServersCount             = $excludedExchangeServers.Count
        ExcludedServers                  = $configuredExcludedServers -join ';'
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
    New-HtmlExecutiveSummary -Summary $summary -PerServerSummary $perServerSummary -ReadinessInventory $exchangeReadinessInventory -HealthCheckerSummary @($healthCheckerSummary) -HealthCheckerFindings @($healthCheckerFindings) -Path $htmlSummaryPath
    Write-Log "HTML executive summary exported to: $htmlSummaryPath"
    Invoke-ServersAndStorageSharePointUpload -LocalFilePath $htmlSummaryPath
    Save-ServersAndStorageWeeklyHistory
    Remove-CoreSmartM365TimestampedDirectoriesOlderThan -RootPath $OutputRoot -RetentionDays 7 -ExcludeDirectories @($OutputFolder) -LogFile $LogFile

    $mailSubject = "SmartM365 Exchange OnPrem Servers and Storage inventory - $RunId"
    Send-ServersAndStorageHtmlReport `
        -HtmlReportPath $htmlSummaryPath `
        -Subject $mailSubject `
        -Attachments @($summaryPath, $perServerSummaryPath, $exchangeReadinessPath) `
        -Summary $summary `
        -PerServerSummary @($perServerSummary) `
        -ReadinessInventory @($exchangeReadinessInventory)
    Write-Log "HTML executive summary email sent."

    Write-Log "Completed successfully."
    $completionStatus = if ($script:ServersAndStorageWarningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
    $script:CompletionStatus = $completionStatus
    Complete-ServersAndStorageRun -Status $completionStatus -Started $StartTime -ErrorMessage $null
    if ($completionStatus -eq 'CompletedWithWarnings') { $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(3); exit $smartM365RunGuard.ExitCode }

    $smartM365RunGuard.ReachedTerminal=$true
}
catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_

    $script:CompletionStatus = 'Failed'
    $message = $_.Exception.Message
    try { Write-Log -Level "ERROR" -Message $message } catch {}
    try { Complete-ServersAndStorageRun -Status 'Failed' -Started $StartTime -ErrorMessage $message } catch {}
    throw
}
finally {
    $smartM365RunGuard.InCleanup=$true

    # 'Auto' here means the run never reached its end (Ctrl+C or a stopped pipeline): the shared banner shows
    # CANCELLED and the lifecycle finalizer exits with code 130. No receipt or success status was written.
    $bannerStatus = $script:CompletionStatus
    if ($bannerStatus -eq 'Auto') {
        try { Write-Log -Level 'WARN' -Message 'Run interrupted before completion (cancelled): no receipt or success status was written.' } catch { $script:ServersAndStorageWarningCount++ }
    }
    Write-SmartM365CompletionBanner -Status $bannerStatus -ScriptName $ScriptName -StartedAt $StartTime -WarningCount $script:ServersAndStorageWarningCount -ErrorCount $script:ServersAndStorageErrorCount
}
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBzazwp55lLi41u
# 85ClDQsVsYn5OAthMfM48dhANceq8qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBRsfahTdL7H5EeVFrYaImDNMybGt8/G0vodiA03Wtl4MA0GCSqG
# SIb3DQEBAQUABIIBgAQKoawFJPFZupoKKGAvI/zzvZSQ5rRdlqiPlrUDDLJTJgaM
# mIvDluISN5oz0lME1+1r3GvEv+P+GnKvX/pqdge+DKozTX7KHTswK/MHo4Xpm/HE
# u0773dvpXt1z/pNhbw9J/yv/oqZVEkSFMDAJM4NiwNN9r5H7FmafvrZcROnpMAi7
# IOOANAl0y2x4TZ/n7mCp88ceM0Nd50UI/VmzR49FuVKHYfMRkzgY6TE4bd2Xfbry
# ilycA0PNG7upNjBw7ZCKaFJiGD6bk6Ks6g/TLc/IS7eJyBdraBViMPHtgJ6snP4V
# dma9czJp/FJF4/3T0sFHXWoRAxslNvQv50PaND6EcW5bw8Ed52wAnekcws15O8Hy
# l/Q7rpYKD8/QcLIGVhlPE1vL8nIiE0PnMKuI2epmegaZ2N7llY08tmuUvKzQnqhi
# 6e5tIdQWrg6DjJ7ilga/hOn/FYn2nWwdB8TiJPlf/Nlq3VP2Gi5C4lOKr0b0XpVf
# LZF2tPqyRh4ExUx3y6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEyMDMy
# NDVaMC8GCSqGSIb3DQEJBDEiBCC/KTD2OEC5XBLrSea62TlPhEf0sBbAe01BoqpW
# AdXzyjANBgkqhkiG9w0BAQEFAASCAgBxANl+me/SpfarkSM3hZatEYucaRTxhNN/
# 6PHG+mGItERoOS46a6housMPrfur85jv6MRdtqRzJKIpM3d7PSNSjFXFvBJ7RlzJ
# vLETKu4s/r0lHvY68DPJGyIo7X26+j/jUH0W7d0W9z5BteYH8KLSvI5qNWPgDX7S
# JYJGrdsqij0+HCXlYx0BPpsdTGuIiOIa+pMws0mxlmptbv18q68GcTEWCbiHOdCv
# LL7cAoE5y98VnIF6TcUOeVbi7wLA2rVczslkYSATZyUiwt4oHF2P10gKjkKZmF68
# zTFkWGVeC9EXns3jDYYwuutMx7I5a/mWpKKU8ADIdWoBiZk284tHGeWAZyU45388
# kiSnEuvpokhJumnxHnctRYjlHlqfMeupFXzqKUty6Ts+DEQD5aHK33SmESFT4UZf
# Y7KphOZpFCmMxW+hoTMpxcHKkFpNfX3ii5hy3Mv3+o3c7rLHoqBzquAw+/7sNage
# vE6L9PTu1W8sIX3tBeV4SAJOlrfWYw5yU6Mjp5brZfDzAWZUlxqIKMBX3Y1PwCAk
# 0n41HfwTsuqcSzp2xuRkU+yu6fzDodKsyIGwALSSpCu7XKHcmC1YUaz+/9Y7VCZT
# hAoqSVtdXORABIDFrWGbSrAkdSVk9We2pFzbICzHyUSGtq/AXriVrXoh2HEPvV8w
# MKbZjiAl0Q==
# SIG # End signature block
