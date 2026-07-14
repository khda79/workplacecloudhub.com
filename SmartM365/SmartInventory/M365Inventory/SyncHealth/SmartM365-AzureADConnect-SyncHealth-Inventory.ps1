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
1.14


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
$ScriptVersion = "1.14"
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
$CurrentOperation = 'Initialize'

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host 'This script requires PowerShell 7 or later.' -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

function Import-SmartM365CoreModule {
    [CmdletBinding()]
    param()

    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $candidate = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) {
            Import-Module -Name $candidate -MinimumVersion '1.0.24' -Force -ErrorAction Stop
            return
        }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    throw 'SmartM365.Core module manifest not found.'
}

function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
            Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null
        }
        else {
            if (-not (Test-Path -LiteralPath $templatePath)) { throw "Missing local config and template: $configPath" }
            Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
        }
    }
    return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}

function Get-SmartM365GlobalConfig {
    [CmdletBinding()]
    param()

    if ($null -ne $script:SmartM365GlobalConfig) { return $script:SmartM365GlobalConfig }
    $script:SmartM365GlobalConfig = [pscustomobject]@{}
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
        if (Test-Path -LiteralPath $globalConfigPath) {
            $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            break
        }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    return $script:SmartM365GlobalConfig
}

function Resolve-SmartM365ConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value -notmatch '\{\{[^}]+\}\}') { return $Value }
    $globalConfig = Get-SmartM365GlobalConfig
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $tokenProperty = $globalConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }
            $tokenValue = Resolve-SmartM365ConfigValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-ScriptLocalConfigValue {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Name, $DefaultValue)

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return Resolve-SmartM365ConfigValue -Value $property.Value }
        }
        else { return Resolve-SmartM365ConfigValue -Value $property.Value }
    }

    $globalConfig = Get-SmartM365GlobalConfig
    $globalProperty = $globalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) { return $DefaultValue }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
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
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
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
        Connect-MgGraph -Scopes @('Directory.Read.All') -NoWelcome -ErrorAction Stop | Out-Null
    }
    else {
        WriteLog -Message 'Connecting to Microsoft Graph using app-only certificate authentication.' -Level INFO
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null
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

        if ($sendMailMode -in @('Graph','Both')) {
            Send-SmartM365GraphMail -From $from -To $to -Cc $cc -Subject $Title -BodyHtml $bodyHtml -Attachments $validAttachments
        }

        if ($sendMailMode -in @('SMTP','Both')) {
            $smtpServer = Resolve-SyncHealthMailValue -Config $ScriptLocalConfig -Name 'SmtpServer' -DefaultValue ''
            if ([string]::IsNullOrWhiteSpace($smtpServer)) { throw 'SmtpServer is required when SendMailMode is SMTP or Both.' }
            $smtpPort = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SmtpPort' -DefaultValue 25)
            $smtpParams = @{
                SmtpServer = $smtpServer
                Port = $smtpPort
                From = $from
                To = @($to -split ';' | Where-Object { $_ })
                Subject = $Title
                Body = $bodyHtml
                BodyAsHtml = $true
                ErrorAction = 'Stop'
            }
            if (-not [string]::IsNullOrWhiteSpace($cc)) { $smtpParams['Cc'] = @($cc -split ';' | Where-Object { $_ }) }
            if ($validAttachments.Count -gt 0) { $smtpParams['Attachments'] = $validAttachments }
            Send-MailMessage @smtpParams
            WriteLog -Message ("SMTP mail sent from {0} to {1}" -f $from, $to) -Level SUCCESS
        }
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
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath 'AzureADConnect-SyncHealth' }
$logPath = Join-Path -Path $logRoot -ChildPath ("SmartM365-AzureADConnect-SyncHealth-Inventory_{0}.log" -f $timestamp)
foreach ($folder in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logRoot)) { if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null } }
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$global:LogTextFile = $logPath

function Stop-SmartM365SyncHealthTranscript {
    [CmdletBinding()]
    param()
    try {
        Stop-Transcript | Out-Null
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
    $ScriptCsvLogFolderPath = $initializedOutput
    if ([string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath)) { $WeeklyHistoryFolderPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath 'WeeklyHistory' }
    $timestampedCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("M365_Entra_AzureADConnect_SyncHealth_{0}.csv" -f $timestamp)
    $latestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath $csvName
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
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
}
catch {
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
    try { RemoveOldFiles -Path $ScriptCsvLogFolderPath -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile } catch {}
    try { Stop-SmartM365SyncHealthTranscript } catch {}
    try { Complete-SmartM365ExecutionContext -Status $script:CompletionStatus -FailureStage $(if ($script:CompletionStatus -eq 'Failed') { $CurrentOperation } else { '' }) } catch {}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBKyJYlOScJWU+x
# GVFII1pXSevsEFFF9xVb9v8WFtuAOKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIMxVUwaWBgRsu3dnYM6TFxZ01CJzeYsnVJ7Ioy9Yo2PkMA0GCSqG
# SIb3DQEBAQUABIIBgHZUu9TIAmCIPdEviy7o1WGQUHDUxxHd2jog+BbestiGngHk
# icYy0Jjy4gHAmys6/bVtoYAFvnDf/YxwQVEj6XNZU0PkJaIlM5vFRtcu99KxV20u
# ZwnMBr4aIBLVn8AeqsJ91DXIra1cTFeMCVUZDSaIGalVGTY18Tc/d5BWnDiZN2aE
# ZynInRafEVyCKVxSIJoOTmKNH0S0cH2IbZ+DtCUhi+wi/i9eCgAmevqvZ11orwTh
# fN4RXjodu9ln8NtzpYGOGuzojFvMLldGjPpmQeN2gEszkORecKFKCSP0V7bjbOQX
# P2d3mYT4w4NslRnzUKRIhxWTwvLrn5XdYudw4izLqOe9sYnETvEA0Qoe10AmeH2M
# xS4j992OKWsdCxqaEWUsTYwn/fn4/Fbb/TMQW0nOXtAY4r7x1DplWEGiJZVwEpxS
# 90NWPgE/Rj1fBkKHfKehbAEGIJ1TGnxfSTC9jkhsq9qfyQmiFOpiSD3+bTyIGiFg
# EvERISrrUqLhkpn7B6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTQwMTEy
# NTdaMC8GCSqGSIb3DQEJBDEiBCBG3zQt3Y8Io+c/g6ixc6jNK+kQBY0ZJzTqCFZq
# tx1bSTANBgkqhkiG9w0BAQEFAASCAgAjkGURCgeZweW4oB3r416oRfzNiVE6bdmM
# UD61aSPp7ltztlqnKZl4RFFBllKrQ8kvnjLaRcDRACXuM1K7PTGkFC3ZPcSk4E/5
# MQAmCm/zaRpCENrz6Tb9JKIW+dR9oCQp7sIGwE/LzSLkenMowZTjOjGMxmCP2APU
# 48RedBK84xy53itoLf7zV6Vh2SbFsBVqhT+Ut1XJU2PY3H4HkuXKKf0Td1QnWXR0
# HTYGy/5udH/ILvACv4KQyPgSFKfRrraK7qwFWTMFoiCRf9mjzMHrlulku3f3jsU0
# d1EtZjI+//7zUm13jliCEXfWTPw3WlFAaSqB8n9qr+Hko4Ep7bN3OEyi5EdM7CiB
# /9YiaUuPcvhfKiU3rfX8FnrKopa+7TA6+k6mu4k8ZBge6/qRTsB6zufTS7t1mlmL
# Sb/SWV65900zO62faw3ga0UV3Nnq6eAVBPC0b+UvOUjXs63JjaDbL+1Mu0YMDmOZ
# jOXdiuU4opCeNmDCf1gVVcHyQCw6azu5MRyo5zclVp8oUd335FkFT7fu8mFm70pJ
# pa3ZEVOBYualE0UNq/1vm8PyRNah4vvy2/YwNtrmEk4MnxJvrzSS4XQQSWTbHaJu
# 2s4s67W7Cak3ij4TjQZy5f+CzL2aaMfz/I0M1rXRLSs0+Ap9sKCTWKEDIsE60feE
# Ar/uqQnBMg==
# SIG # End signature block
