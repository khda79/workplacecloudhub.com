#Requires -Version 7.0
<#
.SYNOPSIS
    Azure AD Connect synchronization health inventory and freshness alert.

.DESCRIPTION
    Reads tenant organization synchronization state from Microsoft Graph, exports health checks to CSV,
    and sends an Alerts notification when Entra ID is not synchronized or when the last sync is older than
    the configured threshold. This protects downstream SmartInventory data consumers from stale hybrid data.

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER SyncAgeThresholdMinutes
    Maximum acceptable age in minutes for on-premises directory synchronization. Default: 180.

.PARAMETER Connect
    Kept for launcher consistency. The script always disconnects any existing Microsoft Graph session before connecting.

.PARAMETER InteractiveAuth
    Uses delegated interactive Microsoft Graph authentication for troubleshooting.

.PARAMETER OutputPath
    Optional output directory override. If omitted, ScriptCsvLogFolderPath from local JSON is used.

.VERSION
1.25
.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: Directory.Read.All.
    Conditional: Mail.Send is required when SyncHealth email notification is enabled; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Requires: PowerShell 7+, Microsoft.Graph.Authentication, SmartM365.Core.psd1
    Minimum application permissions: Directory.Read.All
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [int]$SyncAgeThresholdMinutes = 180,
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [string]$OutputPath,
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
    throw "-MaxItems is not supported by SmartM365-AzureADConnect-SyncHealth-Inventory because it produces tenant-level synchronization health checks, not row-based inventory."
}
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) { if (Test-Path -LiteralPath $p) { return $p } }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365GlobalConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MaximumFunctionCount = 32768
$ScriptVersion = "1.25"
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
$CurrentOperation = 'Initialize'

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host 'This script requires PowerShell 7 or later.' -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
}

function Import-SmartM365CoreModule {
    [CmdletBinding()]
    param()

    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $candidate = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) {
            Import-Module -Name $candidate -MinimumVersion '1.0.122' -Force -ErrorAction Stop
            return
        }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    throw 'SmartM365.Core module manifest not found.'
}

function Get-ScriptLocalConfig {
    param()
    Read-SmartM365ScriptConfiguration -ScriptPath $PSCommandPath
}

function Get-SmartM365GlobalConfig {
    param()
    return $script:SmartM365GlobalConfig
}

function Resolve-SmartM365ConfigValue {
    param([AllowNull()]$Value)
    Resolve-SmartM365ConfigurationTokens -Value $Value -Config $script:SmartM365GlobalConfig
}

function Get-ScriptLocalConfigValue {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)
    Get-SmartM365ConfigurationValue -Name $Name -LocalConfig $Config -EffectiveConfig $script:SmartM365GlobalConfig -DefaultValue $DefaultValue -EffectiveFallback Blank
}

function Ensure-GraphModule {
    [CmdletBinding()]
    param()

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

function Disconnect-GraphSafe {
    [CmdletBinding()]
    param()

    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $context = Get-MgContext -ErrorAction SilentlyContinue
            if ($context) {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false
                WriteLog -Message 'Disconnected from Microsoft Graph.' -Level SUCCESS
            }
        }
    }
    catch { WriteLog -Message ("Disconnect-MgGraph failed (non-fatal): {0}" -f $_.Exception.Message) -Level WARNING }
}

function Connect-GraphForSyncHealth {
    [CmdletBinding()]
    param(
        [switch]$UseInteractiveAuth,
        [string]$AppId,
        [string]$TenantId,
        [string]$Thumbprint
    )

    Disconnect-GraphSafe
    if ($UseInteractiveAuth) {
        WriteLog -Message 'Connecting to Microsoft Graph using interactive authentication.' -Level INFO
        Connect-MgGraph -Scopes @('Directory.Read.All') -NoWelcome -ErrorAction Stop | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $true
    }
    else {
        WriteLog -Message 'Connecting to Microsoft Graph using app-only certificate authentication.' -Level INFO
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $true
    }
    WriteLog -Message 'Connected to Microsoft Graph.' -Level SUCCESS
}

function Send-SyncHealthTeamsNotification {
    [CmdletBinding()]
    param(
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level,
        [string]$Title,
        [string]$Message,
        [hashtable]$Facts
    )

    if (Get-Command Send-SmartM365TeamsNotification -ErrorAction SilentlyContinue) {
        Send-SmartM365TeamsNotification -Title $Title -Message $Message -Level $Level -Facts $Facts | Out-Null
    }
}

function Get-SyncHealthSendMailMode {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Config)

    $configuredMode = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'SendMailMode' -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($configuredMode)) { $configuredMode = $configuredMode.Trim() }

    if ([string]::IsNullOrWhiteSpace($configuredMode) -or $configuredMode -in @('__USE_GLOBAL__', 'USE_GLOBAL')) {
        $smtpServer = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'SmtpServer' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($smtpServer)) { return 'Graph' }
        return 'SMTP'
    }

    switch ($configuredMode.ToLowerInvariant()) {
        'graph' { return 'Graph' }
        'smtp' { return 'SMTP' }
        'both' { return 'Both' }
        default { throw ("Invalid SendMailMode '{0}'. Use Graph, SMTP, or Both." -f $configuredMode) }
    }
}

function Resolve-SyncHealthMailValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Name,
        [string]$DefaultValue = ''
    )

    $value = [string](Get-ScriptLocalConfigValue -Config $Config -Name $Name -DefaultValue $DefaultValue)
    if ($value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return $DefaultValue }
    return $value
}

function ConvertTo-SyncHealthRecipientString {
    [CmdletBinding()]
    param([string]$Recipients)

    if ([string]::IsNullOrWhiteSpace($Recipients)) { return '' }
    return (@($Recipients -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ';')
}

function Send-SyncHealthMailNotification {
    [CmdletBinding()]
    param(
        [ValidateSet('WARNING','ERROR')][string]$Level,
        [string]$Title,
        [string]$Message,
        [hashtable]$Facts,
        [string[]]$Attachments
    )

    try {
        $from = Resolve-SyncHealthMailValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue ''
        if ([string]::IsNullOrWhiteSpace($from)) {
            throw 'Sync Health email notification requires From in configuration.'
        }
        $to = ConvertTo-SyncHealthRecipientString -Recipients (Resolve-SyncHealthMailValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($to)) {
            $to = ConvertTo-SyncHealthRecipientString -Recipients (Resolve-SyncHealthMailValue -Config $ScriptLocalConfig -Name 'ErrorMailTo' -DefaultValue '')
        }
        if ([string]::IsNullOrWhiteSpace($to)) {
            throw 'Sync Health email notification requires To or ErrorMailTo in configuration.'
        }

        $cc = ConvertTo-SyncHealthRecipientString -Recipients (Resolve-SyncHealthMailValue -Config $ScriptLocalConfig -Name 'Cc' -DefaultValue '')
        $overallStatus = if ($Facts -and $Facts.ContainsKey('OverallStatus')) { [string]$Facts['OverallStatus'] } else { $Level }
        $syncEnabled = if ($Facts -and $Facts.ContainsKey('SyncEnabled')) { [string]$Facts['SyncEnabled'] } else { '' }
        $syncAgeText = if ($Facts -and $Facts.ContainsKey('SyncAgeMinutes')) { [string]$Facts['SyncAgeMinutes'] } else { '' }
        $thresholdText = if ($Facts -and $Facts.ContainsKey('ThresholdMinutes')) { [string]$Facts['ThresholdMinutes'] } else { '' }
        $lastSyncText = if ($Facts -and $Facts.ContainsKey('LastSyncDateTimeUtc')) { [string]$Facts['LastSyncDateTimeUtc'] } else { '' }
        $latestCsvPathForMail = if ($Facts -and $Facts.ContainsKey('LatestCsvPath')) { [string]$Facts['LatestCsvPath'] } else { '' }
        $logFileForMail = if ($Facts -and $Facts.ContainsKey('LogFile')) { [string]$Facts['LogFile'] } else { '' }

        $syncEnabledStatus = if ($syncEnabled -eq 'True') { 'OK' } elseif ([string]::IsNullOrWhiteSpace($syncEnabled)) { 'WARNING' } else { 'ERROR' }
        $syncAgeStatus = if ($overallStatus -eq 'ERROR') { 'ERROR' } elseif ($overallStatus -eq 'WARNING') { 'WARNING' } else { 'OK' }

        $emailSeverity = switch ($Level) {
            'WARNING' { 'Warning' }
            'ERROR' { 'Error' }
            default { 'Info' }
        }

        function ConvertTo-SyncHealthMailHtmlText {
            param([AllowNull()]$Value)
            if ($null -eq $Value) { return '' }
            return [System.Net.WebUtility]::HtmlEncode([string]$Value)
        }

        function New-SyncHealthStatusBadgeHtml {
            param([string]$Status)
            $normalizedStatus = if ([string]::IsNullOrWhiteSpace($Status)) { 'INFO' } else { $Status.ToUpperInvariant() }
            $badge = switch ($normalizedStatus) {
                'OK' { @{ Text = 'OK'; Bg = '#dcfce7'; Border = '#86efac'; Color = '#166534' } }
                'SUCCESS' { @{ Text = 'OK'; Bg = '#dcfce7'; Border = '#86efac'; Color = '#166534' } }
                'WARNING' { @{ Text = 'WARNING'; Bg = '#fef3c7'; Border = '#fcd34d'; Color = '#92400e' } }
                'ERROR' { @{ Text = 'ERROR'; Bg = '#fee2e2'; Border = '#fca5a5'; Color = '#991b1b' } }
                default { @{ Text = $normalizedStatus; Bg = '#e0f2fe'; Border = '#7dd3fc'; Color = '#075985' } }
            }
            return '<span style="display:inline-block;min-width:72px;text-align:center;border-radius:999px;border:1px solid {0};background:{1};color:{2};font-size:11px;line-height:18px;font-weight:700;letter-spacing:.2px;">{3}</span>' -f $badge.Border, $badge.Bg, $badge.Color, $badge.Text
        }

        function New-SyncHealthTestRowHtml {
            param(
                [string]$Check,
                [string]$Status,
                [string]$Value,
                [string]$Detail
            )
            $safeCheck = ConvertTo-SyncHealthMailHtmlText $Check
            $safeValue = ConvertTo-SyncHealthMailHtmlText $Value
            $safeDetail = ConvertTo-SyncHealthMailHtmlText $Detail
            $badgeHtml = New-SyncHealthStatusBadgeHtml -Status $Status
            return '<tr><td style="border-bottom:1px solid #e5edf5;padding:10px 12px;font-size:13px;color:#334155;">{0}<div style="font-size:11px;line-height:16px;color:#64748b;margin-top:2px;">{1}</div></td><td align="right" style="border-bottom:1px solid #e5edf5;padding:10px 12px;font-size:13px;font-weight:700;color:#111827;">{2}</td><td align="right" style="border-bottom:1px solid #e5edf5;padding:10px 12px;">{3}</td></tr>' -f $safeCheck, $safeDetail, $safeValue, $badgeHtml
        }

        $testRows = @(
            New-SyncHealthTestRowHtml -Check 'Directory synchronization enabled' -Status $syncEnabledStatus -Value $syncEnabled -Detail 'Expected value: True'
            New-SyncHealthTestRowHtml -Check 'Last synchronization freshness' -Status $syncAgeStatus -Value ("{0} min" -f $syncAgeText) -Detail ("Threshold: {0} min" -f $thresholdText)
            New-SyncHealthTestRowHtml -Check 'Last synchronization timestamp' -Status 'INFO' -Value $lastSyncText -Detail 'UTC timestamp reported by Microsoft Graph'
        )
        $testsHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;border-radius:4px;overflow:hidden;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Check</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Value</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Status</th>
  </tr>
  $($testRows -join "`n")
</table>
"@

        $pathRows = @()
        if (-not [string]::IsNullOrWhiteSpace($latestCsvPathForMail)) {
            $pathRows += '<div style="margin-top:4px;"><span style="font-weight:700;color:#475569;">CSV:</span> <span style="font-family:Consolas,''Courier New'',monospace;word-break:break-all;">' + (ConvertTo-SyncHealthMailHtmlText $latestCsvPathForMail) + '</span></div>'
        }
        if (-not [string]::IsNullOrWhiteSpace($logFileForMail)) {
            $pathRows += '<div style="margin-top:4px;"><span style="font-weight:700;color:#475569;">Log:</span> <span style="font-family:Consolas,''Courier New'',monospace;word-break:break-all;">' + (ConvertTo-SyncHealthMailHtmlText $logFileForMail) + '</span></div>'
        }
        $technicalFilesHtml = ''
        if ($pathRows.Count -gt 0) {
            $technicalFilesHtml = @"
          <tr>
            <td style="padding:12px 24px 4px 24px;color:#64748b;font-size:11px;line-height:16px;">
              <div style="font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.3px;margin-bottom:4px;">Technical files</div>
              $($pathRows -join "`n")
            </td>
          </tr>
"@
        }

        $bodyHtml = New-SmartM365EmailBody -Title $Title -Category 'SmartM365 Sync Health' -Severity $emailSeverity -Message $Message -Sections @([pscustomobject]@{ Title = 'Checks'; Html = $testsHtml }) -BodyHtml $technicalFilesHtml
        if (Get-Command -Name Convert-SmartM365MailBodyLocalPathsToSharePointLinks -ErrorAction SilentlyContinue) {
            $bodyHtml = Convert-SmartM365MailBodyLocalPathsToSharePointLinks -BodyHtml $bodyHtml
        }
        $validAttachments = @($Attachments | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) })
        $sendMailMode = Get-SyncHealthSendMailMode -Config $ScriptLocalConfig

        $smtpServer = Resolve-SyncHealthMailValue -Config $ScriptLocalConfig -Name 'SmtpServer' -DefaultValue ''
        $smtpPort = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SmtpPort' -DefaultValue 25)
        $mailPurpose = if ($Level -eq 'ERROR') { 'Error' } else { 'Report' }
        Send-SmartM365Mail -DeliveryClass Operational -MailPurpose $mailPurpose -SendMailMode $sendMailMode -SmtpServer $smtpServer -SmtpPort $smtpPort -From $from -To $to -Cc $cc -Subject $Title -BodyHtml $bodyHtml -Attachments $validAttachments -AllowAttachments
        WriteLog -Message ("Sync Health mail sent from {0} to {1} via {2}" -f $from, $to, $sendMailMode) -Level SUCCESS
    }
    catch {
        WriteLog -Message ("Sync Health email notification failed: {0}" -f $_.Exception.Message) -Level ERROR
        throw
    }
}

Import-SmartM365CoreModule
$ScriptLocalConfig = Get-ScriptLocalConfig

$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
$ScriptCsvLogFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue ''
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''
$ConfiguredThreshold = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SyncAgeThresholdMinutes' -DefaultValue $SyncAgeThresholdMinutes)
if ($PSBoundParameters.ContainsKey('SyncAgeThresholdMinutes') -eq $false) { $SyncAgeThresholdMinutes = $ConfiguredThreshold }
$WeeklyHistoryFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryFolderPath' -DefaultValue ''

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $ScriptCsvLogFolderPath = $OutputPath }
if ([string]::IsNullOrWhiteSpace($ScriptCsvLogFolderPath)) { $ScriptCsvLogFolderPath = Join-Path -Path $PSScriptRoot -ChildPath 'Output' }
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $LatestCsvFolderPath = $ScriptCsvLogFolderPath }

Ensure-GraphModule

$runId = [guid]::NewGuid().ToString()
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvName = 'M365_Entra_AzureADConnect_SyncHealth.csv'
$timestampedCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("M365_Entra_AzureADConnect_SyncHealth_{0}.csv" -f $timestamp)
$latestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath $csvName
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath 'SmartM365-EntraConnect-SyncHealth' }
$logPath = Join-Path -Path $logRoot -ChildPath ("SmartM365-AzureADConnect-SyncHealth-Inventory_{0}.log" -f $timestamp)
foreach ($folder in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logRoot)) { if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null } }
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$global:LogTextFile = $logPath

function Stop-SmartM365SyncHealthTranscript {
    [CmdletBinding()]
    param()
    try {
        Stop-SmartM365RunTranscript | Out-Null
        $transcriptPath = $global:logTranscriptFile
        if (-not $transcriptPath) { $transcriptPath = $global:LogTranscriptFile }
        if ($transcriptPath -and (Get-Command Update-SmartM365TimestampedTranscript -ErrorAction SilentlyContinue)) {
            Update-SmartM365TimestampedTranscript -Path $transcriptPath
        }
    }
    catch {}
}

$script:CompletionStatus = 'Auto'
try {
    $CurrentOperation = 'InitializeScriptEnvironment'
    $initializedOutput = InitializeScriptEnvironment -OutputPath $ScriptCsvLogFolderPath -LogFileName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    . (Join-Path (Get-Module SmartM365.Core).ModuleBase 'SmartM365-SharePointUploadPolicy.ps1')
    Initialize-SmartM365CollectorUploadContext -LocalConfig $ScriptLocalConfig -GlobalConfig $global:SmartM365GlobalConfig -DataAllRoot $initializedOutput -DataLastRoot $LatestCsvFolderPath -LogAllRoot $global:LogPath -ExternalActionsAllowed $true
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $LatestCsvFolderPath
    $ScriptCsvLogFolderPath = $initializedOutput
    if ([string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath)) { $WeeklyHistoryFolderPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath 'WeeklyHistory' }
    $timestampedCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("M365_Entra_AzureADConnect_SyncHealth_{0}.csv" -f $timestamp)
    $latestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath $csvName
    Start-SmartM365RunTranscript -Path $global:logTranscriptFile -Append | Out-Null
    WriteLog -Message "Starting $TaskName. Threshold=$SyncAgeThresholdMinutes minute(s)." -Level INFO
    WriteLog -Message "Default WeeklyHistoryFolderPath: $WeeklyHistoryFolderPath" -Level INFO
    if ($Connect) { WriteLog -Message 'Connect switch specified; Graph connection will be established by this script.' -Level INFO }

    $CurrentOperation = 'ConnectGraph'
    Connect-GraphForSyncHealth -UseInteractiveAuth:$InteractiveAuth -AppId $AppId -TenantId $TenantId -Thumbprint $Thumb

    $CurrentOperation = 'Preflight'
    Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules @('Microsoft.Graph.Authentication') -OutputPaths @($ScriptCsvLogFolderPath, $LatestCsvFolderPath) -RequiredGraphApplicationPermissions @('Directory.Read.All') -GraphProbeUris @('https://graph.microsoft.com/v1.0/organization?$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime') | Out-Null

    $CurrentOperation = 'ReadOrganizationSyncState'
    $orgResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime' -OutputType PSObject -ErrorAction Stop
    $org = @($orgResponse.value) | Select-Object -First 1
    if (-not $org) { throw 'Graph /organization returned no tenant organization object.' }

    $syncEnabled = [bool]$org.onPremisesSyncEnabled
    $lastSyncText = [string]$org.onPremisesLastSyncDateTime
    $syncAgeMinutes = $null
    if (-not [string]::IsNullOrWhiteSpace($lastSyncText)) {
        $lastSyncUtc = [datetime]::Parse($lastSyncText, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()
        $syncAgeMinutes = [math]::Round(((Get-Date).ToUniversalTime() - $lastSyncUtc).TotalMinutes, 1)
    }

    $exportDateTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $rows = @(
        [pscustomobject][ordered]@{
            CheckName = 'SyncEnabled'
            Status = if ($syncEnabled) { 'OK' } else { 'ERROR' }
            Value = [string]$syncEnabled
            Threshold = 'true'
            Detail = if ($syncEnabled) { 'On-premises directory synchronization is enabled.' } else { 'On-premises directory synchronization is disabled. Hybrid SmartInventory data may be stale or incomplete.' }
            OrganizationId = [string]$org.id
            OrganizationName = [string]$org.displayName
            LastSyncDateTimeUtc = $lastSyncText
            SyncAgeMinutes = $syncAgeMinutes
            ExportDateTime = $exportDateTime
            RunId = $runId
        },
        [pscustomobject][ordered]@{
            CheckName = 'LastSyncAge'
            Status = if (-not $syncEnabled) { 'ERROR' } elseif ($null -eq $syncAgeMinutes) { 'WARNING' } elseif ($syncAgeMinutes -le $SyncAgeThresholdMinutes) { 'OK' } else { 'WARNING' }
            Value = if ($null -eq $syncAgeMinutes) { '' } else { "$syncAgeMinutes min" }
            Threshold = "$SyncAgeThresholdMinutes min"
            Detail = if (-not $syncEnabled) { 'Sync is disabled; last sync freshness cannot be trusted.' } elseif ($null -eq $syncAgeMinutes) { 'Last sync timestamp is empty.' } elseif ($syncAgeMinutes -le $SyncAgeThresholdMinutes) { "Last sync age is within threshold: $syncAgeMinutes minute(s)." } else { "Last sync is stale: $syncAgeMinutes minute(s), threshold is $SyncAgeThresholdMinutes minute(s). SmartInventory data may not be up to date." }
            OrganizationId = [string]$org.id
            OrganizationName = [string]$org.displayName
            LastSyncDateTimeUtc = $lastSyncText
            SyncAgeMinutes = $syncAgeMinutes
            ExportDateTime = $exportDateTime
            RunId = $runId
        }
    )

    $CurrentOperation = 'ExportCsv'
    Export-SmartM365Csv -Data $rows -TimestampedPath $timestampedCsvPath -LatestPath $latestCsvPath | Out-Null

    $overallStatus = if (@($rows | Where-Object Status -eq 'ERROR').Count -gt 0) { 'ERROR' } elseif (@($rows | Where-Object Status -eq 'WARNING').Count -gt 0) { 'WARNING' } else { 'OK' }
    $message = "Azure AD Connect sync health: $overallStatus. SyncEnabled=$syncEnabled; LastSyncAgeMinutes=$syncAgeMinutes; Threshold=$SyncAgeThresholdMinutes."
    $statusLogLevel = switch ($overallStatus) {
        'ERROR' { 'ERROR' }
        'WARNING' { 'WARNING' }
        default { 'SUCCESS' }
    }
    WriteLog -Message $message -Level $statusLogLevel

    $facts = @{
        Script = $MyInvocation.MyCommand.Name
        TenantOrOrganization = [string]$org.displayName
        OverallStatus = $overallStatus
        SyncEnabled = [string]$syncEnabled
        LastSyncDateTimeUtc = $lastSyncText
        SyncAgeMinutes = [string]$syncAgeMinutes
        ThresholdMinutes = [string]$SyncAgeThresholdMinutes
        LatestCsvPath = $latestCsvPath
        TimestampedCsvPath = $timestampedCsvPath
        WeeklyHistoryFolderPath = $WeeklyHistoryFolderPath
        LogFile = $global:LogTextFile
    }
    $CurrentOperation = 'SendNotifications'
    switch ($overallStatus) {
        'ERROR' {
            Send-SyncHealthTeamsNotification -Level ERROR -Title 'Azure AD Connect sync health error' -Message $message -Facts $facts
            Send-SyncHealthMailNotification -Level ERROR -Title 'Azure AD Connect sync health error' -Message $message -Facts $facts -Attachments @($latestCsvPath)
        }
        'WARNING' {
            Send-SyncHealthTeamsNotification -Level WARNING -Title 'Azure AD Connect sync health warning' -Message $message -Facts $facts
            Send-SyncHealthMailNotification -Level WARNING -Title 'Azure AD Connect sync health warning' -Message $message -Facts $facts -Attachments @($latestCsvPath)
        }
        default {
            WriteLog -Message 'Sync Health notifications skipped because overall status is OK.' -Level INFO
        }
    }

    Stop-SmartM365SyncHealthTranscript
    Disconnect-GraphSafe

    $smartM365RunGuard.ReachedTerminal=$true
}
catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_

    $script:CompletionStatus = 'Failed'
    $globalError = $_
    WriteLog -Message ("Global error during {0}: {1}" -f $CurrentOperation, $globalError.Exception.Message) -Level ERROR
    $errorFacts = @{
        Script = $MyInvocation.MyCommand.Name
        TenantOrOrganization = $OrgDomain
        Operation = $CurrentOperation
        LogFile = $global:LogTextFile
    }
    Send-SyncHealthTeamsNotification -Level ERROR -Title 'Azure AD Connect sync health failed' -Message $globalError.Exception.Message -Facts $errorFacts
    try {
        Send-SyncHealthMailNotification -Level ERROR -Title 'Azure AD Connect sync health failed' -Message $globalError.Exception.Message -Facts $errorFacts
    }
    catch {
        WriteLog -Message ("Failed to send Sync Health error email notification: {0}" -f $_.Exception.Message) -Level ERROR
    }
    Stop-SmartM365SyncHealthTranscript
    Disconnect-GraphSafe
    throw
}
finally {
    $smartM365RunGuard.InCleanup=$true

    try {
        if ($script:CompletionStatus -ne 'Failed') {
            Remove-SmartM365TimestampedFilesOlderThan -FolderPath $ScriptCsvLogFolderPath -FilePattern '*.csv' -RetentionDays 7 -LogFile $global:LogTextFile
        }
    } catch {}
    try { Stop-SmartM365SyncHealthTranscript } catch {}
    try { Complete-SmartM365ExecutionContext -Status $script:CompletionStatus -FailureStage $(if ($script:CompletionStatus -eq 'Failed') { $CurrentOperation } else { '' }) } catch {}
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCiD+J0Q/aV/E9y
# p8Ye3gavS+oS7oeqX6GziV27NZ68PKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDDVQvwk8CeoEUphxI3u71P
# ZYO5i2mC5iBEGctuI+vRajANBgkqhkiG9w0BAQEFAASCAYCaxfC7/VVuvKVJWrnq
# 9zWUabb0x67gAUOUOc1++zgdEWk+I6Z5MW+C4th03idBPRR6yocnOdkoP/gvhCM9
# wyCykdokBeRjVPNpMiquYXSpjcEj1ZN+MXpOHF+43d43vaEb9aUY+O5GFt+YrwkI
# YYBrh0om143zO6z6v6REL9lX6xzM3tyOUqIHe8FRO/yYyzIJ4v4Dtin6aaRIC1EI
# ugHYD/zxfpzI0Z/U6u1RtSVMhG5FDKDHeiTZYl27/APj2/h/hson11a1WX5AueVD
# S1voVEtQiMXReyhOLCmawo9Xs5PKKdN+Zt6viA6z8vzaINJhxw4pCNma9XKdyUhy
# Ha1jfLMbH0wtNUUnvZpJnVJ/MQ2/tGLIYwjee3BZQaAgiwdq+Czwjl9062AMM6KS
# KlYenOH/1ej7gTkfgeZ2zt+eMjd3SYRqiMQXCMIS0mFMVWkKiDBR88JTlrrRvdn8
# QiBPrBGiWTlp2YG8fAv54UOdhywloI1nCeuGI0UKmDbJLR8=
# SIG # End signature block
