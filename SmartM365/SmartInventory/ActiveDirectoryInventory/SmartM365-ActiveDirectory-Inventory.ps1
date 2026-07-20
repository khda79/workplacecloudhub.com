<#
.SYNOPSIS
    Active Directory full inventory (OUs, Computers, Users, Groups, Contacts) across one or multiple domains.

.DESCRIPTION
    This script performs a comprehensive inventory of Active Directory across:
    - Organizational Units (OUs)
    - Computers
    - Users
    - Groups
    - Contacts

    It:
    - Discovers all domains in the forest or uses a subset passed via -TargetDomains
    - Exports detailed CSVs per domain in a "Not-CSV-Combined" folder
    - Combines all per-domain CSVs into global "AllDomains" CSVs
    - Analyzes duplicate UserPrincipalNames, SMTP proxy addresses, and remote mailbox routing consistency across all domains
    - Uses the shared framework (SmartM365.Core / InitializeScriptEnvironment)
    - Logs to text + transcript
    - Copies combined CSVs to a local configuration path (LatestCsvFolderPath)
    - Cleans old CSV/log files
    - Sends an email notification in case of a global error (SendEmailHtmlReport)

.VERSION
1.43
.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; ActiveDirectory RSAT/Windows Server module; ImportExcel for the diagnostic mail workbook.
    Minimum permissions: read access to all target AD domains and user/computer/group attributes collected by Get-AD* cmdlets.
    Conditional: Mail.Send is required only when Graph mail notifications are enabled; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Minimum permissions: PowerShell 7+, RSAT ActiveDirectory module, and read access to every targeted domain/Global Catalog.
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',

    [Parameter(Mandatory = $false)]
    [string[]]$TargetDomains,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [int]$DomainParallelThrottleLimit = 0,

    [Parameter(Mandatory = $false)]
    [switch]$DomainWorker,

    [Parameter(Mandatory = $false)]
    [string]$DomainWorkerTempFolder,

    [Parameter(Mandatory = $false)]
    [switch]$ReportOnly,

    [Parameter(Mandatory = $false)]
    [switch]$DuplicateAnalysisOnly,

    [Parameter(Mandatory = $false)]
    [switch]$ForceSendDuplicateNotification,

    [Parameter(Mandatory = $false)]
    [switch]$SkipDailyReport,

    [Parameter(Mandatory = $false)]
    [switch]$ForceSendDailySummary,
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
    throw "-MaxItems is not supported by SmartM365-ActiveDirectory-Inventory because duplicate detection and daily summary require a complete Active Directory snapshot."
}
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
$script:SmartM365GlobalConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

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
            if (-not (Test-Path -LiteralPath $templatePath)) {
                $message = @(
                    "Local configuration file not found: $configPath",
                    "Template to copy is also missing: $templatePath",
                    'Create the .local.json file from a safe template, then run the script again.'
                ) -join [Environment]::NewLine
                throw $message
            }

            Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
            Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
            Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
        }
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if (Get-Command Sync-SmartM365JsonConfigWithTemplate -ErrorAction SilentlyContinue) {
            return (Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $configPath)
        }
        return $config
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
}

function Resolve-SmartM365ConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    if ($Value -notmatch '\{\{[^}]+\}\}') {
        return $Value
    }

    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($ScriptRoot) { $ScriptRoot } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$tokenName]
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
function Test-SmartM365UseGlobalConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string]) { return $false }

    $text = $Value.Trim()
    return ($text -in @('__USE_GLOBAL__', 'USE_GLOBAL', '**USE_GLOBAL**', '**USE\_GLOBAL**'))
}

function Get-ScriptLocalConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and -not (Test-SmartM365UseGlobalConfigValue -Value $localValue)) {
                return Resolve-SmartM365ConfigValue -Value $property.Value
            }
        }
        else {
            return Resolve-SmartM365ConfigValue -Value $property.Value
        }
    }


    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $PSCommandPath -Parent }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $globalProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string]) {
            $globalValue = $globalProperty.Value.Trim()
            if ([string]::IsNullOrWhiteSpace($globalValue) -or (Test-SmartM365UseGlobalConfigValue -Value $globalValue)) {
                return $DefaultValue
            }
        }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}
function Get-SmartM365RemoteRoutingDomain {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Config)

    $remoteRoutingDomain = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'RemoteRoutingDomain' -DefaultValue '')
    $sourceName = 'RemoteRoutingDomain'
    if ([string]::IsNullOrWhiteSpace($remoteRoutingDomain)) {
        $orgDomain = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'OrgDomain' -DefaultValue '')
        $orgDomain = $orgDomain.Trim().TrimStart('@').ToLowerInvariant()
        $sourceName = 'OrgDomain fallback'

        if ($orgDomain.EndsWith('.mail.onmicrosoft.com', [System.StringComparison]::OrdinalIgnoreCase)) {
            $remoteRoutingDomain = $orgDomain
        }
        elseif ($orgDomain.EndsWith('.onmicrosoft.com', [System.StringComparison]::OrdinalIgnoreCase)) {
            $remoteRoutingDomain = $orgDomain.Substring(0, $orgDomain.Length - '.onmicrosoft.com'.Length) + '.mail.onmicrosoft.com'
        }
        else {
            throw 'RemoteRoutingDomain is not configured and OrgDomain cannot be converted to a tenant mail.onmicrosoft.com routing domain.'
        }
    }

    $remoteRoutingDomain = $remoteRoutingDomain.Trim().TrimStart('@').ToLowerInvariant()
    if ($remoteRoutingDomain -notmatch '^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+mail\.onmicrosoft\.com$') {
        throw ("Invalid RemoteRoutingDomain '{0}'. Expected a tenant-specific domain such as contoso.mail.onmicrosoft.com." -f $remoteRoutingDomain)
    }

    return [pscustomobject]@{
        Domain = $remoteRoutingDomain
        Source = $sourceName
    }
}


$ScriptLocalConfig = Get-ScriptLocalConfig

function Get-SmartM365AdInventorySendMailMode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config
    )

    $configuredMode = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'SendMailMode' -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($configuredMode)) {
        $configuredMode = $configuredMode.Trim()
    }

    if ([string]::IsNullOrWhiteSpace($configuredMode) -or (Test-SmartM365UseGlobalConfigValue -Value $configuredMode)) {
        $smtpServer = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'SmtpServer' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($smtpServer)) { return 'Graph' }
        return 'SMTP'
    }

    switch ($configuredMode.ToLowerInvariant()) {
        'graph' { return 'Graph' }
        'smtp'  { return 'SMTP' }
        'both'  { return 'Both' }
        default { throw ("Invalid SendMailMode '{0}'. Use Graph, SMTP, or Both." -f $configuredMode) }
    }
}

function New-SmartM365SharePointLinksSection {
    [CmdletBinding()]
    param([array]$UploadRecords)

    $records = @($UploadRecords | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.WebUrl) })
    if ($records.Count -eq 0) { return $null }

    $rows = foreach ($record in $records) {
        $label = ConvertTo-SmartM365EmailHtmlText $record.Label
        if ([string]::IsNullOrWhiteSpace($label)) { $label = ConvertTo-SmartM365EmailHtmlText $record.FileName }
        $url = ConvertTo-SmartM365EmailHtmlText $record.WebUrl
        $sharePointPath = ConvertTo-SmartM365EmailHtmlText $record.SharePointPath
        "<tr><td style=`"width:180px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;`">$label</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;word-break:break-all;`"><a href=`"$url`" style=`"color:#2563eb;text-decoration:underline;`">$sharePointPath</a></td></tr>"
    }

    $tableHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  $($rows -join "`n")
</table>
"@

    return [pscustomobject]@{ Title = 'SharePoint links'; Html = $tableHtml }
}

function Add-SmartM365SharePointUploadLabel {
    [CmdletBinding()]
    param(
        [AllowNull()]$UploadRecord,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not $UploadRecord) { return $null }
    $UploadRecord | Add-Member -NotePropertyName Label -NotePropertyValue $Label -Force
    return $UploadRecord
}

function New-SmartM365AdDuplicatePreviewSection {
    [CmdletBinding()]
    param(
        [array]$Rows,
        [Parameter(Mandatory = $true)]
        [ValidateSet('UPN','SMTP','REMOTE')]
        [string]$DuplicateType,
        [int]$Limit = 50
    )

    $items = @($Rows | Where-Object { $_ })
    if ($items.Count -eq 0) { return $null }

    $valueProperty = switch ($DuplicateType) {
        'UPN' { 'UserPrincipalName' }
        'SMTP' { 'SmtpAddress' }
        'REMOTE' { 'NormalizedRemoteRoutingAddress' }
    }
    $label = switch ($DuplicateType) {
        'UPN' { 'duplicate UPN' }
        'SMTP' { 'duplicate SMTP address' }
        'REMOTE' { 'duplicate remote routing address' }
    }
    $title = switch ($DuplicateType) {
        'UPN' { 'Top 50 duplicate UPN accounts' }
        'SMTP' { 'Top 50 duplicate SMTP entries' }
        'REMOTE' { 'Top 50 duplicate remote routing addresses' }
    }

    $groups = @(
        $items |
            Where-Object {
                $property = $_.PSObject.Properties[$valueProperty]
                $null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)
            } |
            Group-Object -Property $valueProperty |
            Where-Object { $_.Count -gt 1 } |
            Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, Name
    )
    if ($groups.Count -eq 0) { return $null }

    $topGroups = @($groups | Select-Object -First $Limit)
    $tableRows = foreach ($group in $topGroups) {
        $sortedRows = @($group.Group | Sort-Object DomainName, SamAccountName, UserPrincipalName)
        $firstInGroup = $true
        foreach ($row in $sortedRows) {
            $valueCell = if ($firstInGroup) { ConvertTo-SmartM365EmailHtmlText $group.Name } else { '' }
            $countCell = if ($firstInGroup) { ConvertTo-SmartM365EmailHtmlText $group.Count } else { '' }
            $firstInGroup = $false

            $domainNameShort = [string]$row.DomainNameShort
            $domainName = [string]$row.DomainName
            $samAccountName = [string]$row.SamAccountName
            $userPrincipalName = [string]$row.UserPrincipalName
            $accountText = ''
            if (-not [string]::IsNullOrWhiteSpace($domainNameShort) -and -not [string]::IsNullOrWhiteSpace($samAccountName)) { $accountText = '{0}\{1}' -f $domainNameShort, $samAccountName }
            elseif (-not [string]::IsNullOrWhiteSpace($samAccountName)) { $accountText = $samAccountName }
            elseif (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) { $accountText = $userPrincipalName }

            $enabledText = [string]$row.Enabled
            if ([string]::IsNullOrWhiteSpace($enabledText)) { $enabledText = 'Unknown' }
            $lastLogonText = [string]$row.LastLogonDate
            if ([string]::IsNullOrWhiteSpace($lastLogonText)) { $lastLogonText = 'Never / unavailable' }

            $accountHtml = ConvertTo-SmartM365EmailHtmlText $accountText
            $enabledHtml = ConvertTo-SmartM365EmailHtmlText $enabledText
            $lastLogonHtml = ConvertTo-SmartM365EmailHtmlText $lastLogonText
            $domainHtml = ConvertTo-SmartM365EmailHtmlText $domainName

            "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$valueCell</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:13px;font-weight:700;color:#111827;`">$countCell</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$accountHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$enabledHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;white-space:nowrap;`">$lastLogonHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$domainHtml</td></tr>"
        }
    }

    $caption = ConvertTo-SmartM365EmailHtmlText ('Showing accounts for top {0} of {1} {2} value(s). Full details are available in the attached Excel workbook.' -f $topGroups.Count, $groups.Count, $label)
    $html = @"
<div style="font-size:13px;color:#64748b;margin-bottom:8px;">$caption</div>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Value</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Count</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Account</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Enabled</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Last logon</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Domain</th>
  </tr>
  $($tableRows -join "`n")
</table>
"@

    return [pscustomobject]@{ Title = $title; Html = $html }
}
function New-SmartM365AdRemoteRoutingIssuePreviewSection {
    [CmdletBinding()]
    param(
        [array]$Rows,
        [int]$Limit = 50
    )

    $items = @($Rows | Where-Object { $_ } | Sort-Object Severity, IssueType, DomainName, SamAccountName | Select-Object -First $Limit)
    if ($items.Count -eq 0) { return $null }

    $tableRows = foreach ($row in $items) {
        $accountText = if (-not [string]::IsNullOrWhiteSpace([string]$row.DomainNameShort) -and -not [string]::IsNullOrWhiteSpace([string]$row.SamAccountName)) {
            '{0}\{1}' -f $row.DomainNameShort, $row.SamAccountName
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$row.UserPrincipalName)) { [string]$row.UserPrincipalName }
        else { [string]$row.DistinguishedName }

        $issueHtml = ConvertTo-SmartM365EmailHtmlText $row.IssueType
        $severityHtml = ConvertTo-SmartM365EmailHtmlText $row.Severity
        $accountHtml = ConvertTo-SmartM365EmailHtmlText $accountText
        $targetAddressHtml = ConvertTo-SmartM365EmailHtmlText $row.TargetAddress
        $expectedDomainHtml = ConvertTo-SmartM365EmailHtmlText $row.ExpectedRemoteRoutingDomain
        "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$issueHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;font-weight:700;color:#334155;`">$severityHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$accountHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$targetAddressHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$expectedDomainHtml</td></tr>"
    }

    $caption = ConvertTo-SmartM365EmailHtmlText ('Showing {0} of {1} remote routing issue row(s). Full details are available in the attached Excel workbook.' -f $items.Count, @($Rows).Count)
    $html = @"
<div style="font-size:13px;color:#64748b;margin-bottom:8px;">$caption</div>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Issue</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Severity</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Account</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">targetAddress</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Expected domain</th>
  </tr>
  $($tableRows -join "`n")
</table>
"@

    return [pscustomobject]@{ Title = 'Top 50 remote routing issues'; Html = $html }
}

function Send-SmartM365AdInventoryEmailHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$BodyHtml,
        [string]$To = '',
        [string[]]$Attachments
    )

    $effectiveSendMailMode = Get-SmartM365AdInventorySendMailMode -Config $ScriptLocalConfig
    $mailParams = @{
        Subject      = $Subject
        BodyHtml     = $BodyHtml
        SendMailMode = $effectiveSendMailMode
    }
    if (-not [string]::IsNullOrWhiteSpace($To)) {
        $mailParams['To'] = $To
    }
    $attachmentPaths = @($Attachments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($attachmentPaths.Count -gt 0) {
        $mailParams['Attachments'] = $attachmentPaths
        $mailParams['AllowAttachments'] = $true
    }

    SendEmailHtmlReport @mailParams
}

function Export-SmartM365AdDiagnosticsWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][array]$Sources,
        [Parameter(Mandatory = $true)][string]$Path
    )

    Import-Module ImportExcel -ErrorAction Stop

    $parentPath = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($parentPath) -and -not (Test-Path -LiteralPath $parentPath)) {
        New-Item -Path $parentPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
    }

    $excelPackage = $null
    try {
        $excelPackage = Open-ExcelPackage -Path $Path -Create
        foreach ($source in @($Sources)) {
            $csvPath = [string]$source.CsvPath
            $worksheetName = [string]$source.WorksheetName
            $tableName = [string]$source.TableName
            if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
                throw ("Diagnostic CSV not found: {0}" -f $csvPath)
            }

            $rows = @(Import-Csv -LiteralPath $csvPath)
            if ($rows.Count -gt 0) {
                $columns = @($rows[0].PSObject.Properties.Name)
            }
            else {
                $header = [string](Get-Content -LiteralPath $csvPath -TotalCount 1 -ErrorAction Stop)
                if ([string]::IsNullOrWhiteSpace($header)) {
                    throw ("Diagnostic CSV has no header: {0}" -f $csvPath)
                }
                $separatorCount = ([regex]::Matches($header, ',')).Count
                $fakeRow = (@('x') * ($separatorCount + 1)) -join ','
                $probe = @(@($header, $fakeRow) | ConvertFrom-Csv)
                $columns = if ($probe.Count -gt 0) { @($probe[0].PSObject.Properties.Name) } else { @() }
            }
            if ($columns.Count -eq 0) {
                throw ("Diagnostic CSV columns could not be resolved: {0}" -f $csvPath)
            }

            $worksheet = $excelPackage.Workbook.Worksheets.Add($worksheetName)
            $matrix = [System.Collections.Generic.List[object[]]]::new()
            $matrix.Add([object[]]$columns)
            foreach ($row in $rows) {
                $values = foreach ($column in $columns) {
                    $property = $row.PSObject.Properties[$column]
                    if ($null -eq $property -or $null -eq $property.Value) { '' } else { [string]$property.Value }
                }
                $matrix.Add([object[]]$values)
            }
            $null = $worksheet.Cells['A1'].LoadFromArrays($matrix)
            $worksheet.View.FreezePanes(2, 1)
            $worksheet.Cells.AutoFitColumns(10, 60)

            $headerAddress = [OfficeOpenXml.ExcelCellBase]::GetAddress(1, 1, 1, $columns.Count)
            $headerRange = $worksheet.Cells[$headerAddress]
            $headerRange.Style.Font.Bold = $true
            $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
            $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))

            if ($rows.Count -gt 0) {
                $tableAddress = [OfficeOpenXml.ExcelCellBase]::GetAddress(1, 1, $rows.Count + 1, $columns.Count)
                $tableRange = $worksheet.Cells[$tableAddress]
                $table = $worksheet.Tables.Add($tableRange, $tableName)
                $table.TableStyle = [OfficeOpenXml.Table.TableStyles]::Medium2
            }
            else {
                $headerRange.AutoFilter = $true
            }
        }

        Close-ExcelPackage -ExcelPackage $excelPackage
        $excelPackage = $null
    }
    finally {
        if ($null -ne $excelPackage) {
            Close-ExcelPackage -ExcelPackage $excelPackage -NoSave
        }
    }

    $workbookFile = Get-Item -LiteralPath $Path -ErrorAction Stop
    WriteLog -Message ("AD identity and mail routing workbook created: {0} ({1:N0} bytes)" -f $workbookFile.FullName, $workbookFile.Length)
    return $workbookFile.FullName
}



$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$DomainFriendlyNames = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DomainFriendlyNames' -DefaultValue ([pscustomobject]@{})
$AdEnrichmentWindowsUpdateAnchorPolicyId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AdEnrichmentWindowsUpdateAnchorPolicyId' -DefaultValue '38ad040b-08ca-41cd-bd86-5da5ef0b740e'
$AdEnrichmentWindowsUpdate24H2PolicyId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AdEnrichmentWindowsUpdate24H2PolicyId' -DefaultValue '82e1d3e6-bbc0-4ddd-b36d-415979dadec6'
$AdEnrichmentWindowsUpdate25H2PolicyId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AdEnrichmentWindowsUpdate25H2PolicyId' -DefaultValue '41046c77-bc66-44af-b4cd-7bbf2c7d343e'
$ConfiguredComputerGroupNamesRaw = @(Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ComputerMembershipGroupNames' -DefaultValue @())
$ConfiguredComputerGroupNames = for ($groupIndex = 0; $groupIndex -lt 10; $groupIndex++) {
    if ($groupIndex -lt $ConfiguredComputerGroupNamesRaw.Count -and $null -ne $ConfiguredComputerGroupNamesRaw[$groupIndex]) {
        ([string]$ConfiguredComputerGroupNamesRaw[$groupIndex]).Trim()
    }
    else {
        ''
    }
}
$EnableOuInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableOuInventory' -DefaultValue $true)
$EnableComputerInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableComputerInventory' -DefaultValue $true)
$EnableUserInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableUserInventory' -DefaultValue $true)
$EnableGroupInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableGroupInventory' -DefaultValue $true)
$EnableContactInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableContactInventory' -DefaultValue $true)
$EnableDuplicateAnalysis = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDuplicateAnalysis' -DefaultValue $true)
$EnableDuplicateNotification = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDuplicateNotification' -DefaultValue $true)
$DuplicateNotificationLastSentFilePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DuplicateNotificationLastSentFilePath' -DefaultValue ''
$DeleteTemporaryPerDomainCsv = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeleteTemporaryPerDomainCsv' -DefaultValue $true)
$TemporaryPerDomainRetentionDays = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TemporaryPerDomainRetentionDays' -DefaultValue 2)
$EnableWeeklyHistory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableWeeklyHistory' -DefaultValue $true)
$WeeklyHistoryFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryFolderPath' -DefaultValue ''
$WeeklyHistoryRetentionWeeks = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)
$EnableDailyReport = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDailyReport' -DefaultValue $true)
$DailyReportAllowedOS = @(Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyReportAllowedOS' -DefaultValue @('Windows 7*','Windows 8*','Windows 10*','Windows 11*'))
$DailyReportInactiveDays = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyReportInactiveDays' -DefaultValue 90)
$EnableDailyReportLock = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDailyReportLock' -DefaultValue $true)
$DailyReportLockRoot = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyReportLockRoot' -DefaultValue 'C:\ProgramData\SmartM365\Locks'
$DailyReportLockName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyReportLockName' -DefaultValue 'SmartM365-ActiveDirectory-Inventory-DailyReport'
$EnableDailySummaryEmail = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDailySummaryEmail' -DefaultValue $true)
$DailySummaryLastSentFilePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailySummaryLastSentFilePath' -DefaultValue ''
$ConfiguredDomainParallelThrottleLimit = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DomainParallelThrottleLimit' -DefaultValue 1)
$EffectiveDomainParallelThrottleLimit = if ($DomainParallelThrottleLimit -gt 0) { $DomainParallelThrottleLimit } else { $ConfiguredDomainParallelThrottleLimit }
if ($EffectiveDomainParallelThrottleLimit -lt 1) { $EffectiveDomainParallelThrottleLimit = 1 }
if ($DomainWorker) { $EffectiveDomainParallelThrottleLimit = 1 }

# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host ("Current PowerShell version: {0}" -f $PSVersionTable.PSVersion) -ForegroundColor Yellow
    exit 1
}

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.40' -ErrorAction Stop
} catch {
    Write-Host ("Failed to import SmartM365.Core module from '{0}' : {1}" -f $modulePath, $_) -ForegroundColor Red
    exit 1
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.43"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$defaultActiveDirectoryInventoryOutputPath = if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath } else { Resolve-SmartM365ConfigValue -Value '{{DataAllRootPath}}\ActiveDirectory\Inventory' }
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ActiveDirectoryInventoryCsvLogFolderPath' -DefaultValue $defaultActiveDirectoryInventoryOutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','') -CallerScriptPath $PSCommandPath
    Start-Transcript -Path $global:LogTranscriptFile -Append

    WriteLog -Message ("Script environment initialized at {0}" -f $InitializeOutputPath)
    $OutputPath = $InitializeOutputPath
    if ([string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath)) {
        $WeeklyHistoryFolderPath = Join-Path -Path $OutputPath -ChildPath 'WeeklyHistory'
    }
    WriteLog -Message ("Starting {0}" -f $TaskName)
    WriteLog -Message ("Configured computer membership group slots: {0}" -f (($ConfiguredComputerGroupNames | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_)) { '<empty>' } else { $_ } }) -join ', '))
}
catch {
    Write-Host ("Initialization failed: {0}" -f $_) -ForegroundColor Red
    exit 1
}

$script:SmartM365AdReferenceXlsxCache = @{}
$script:SmartM365AdReferenceXlsxCacheRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("SmartM365\ADReferenceData\{0}" -f $PID)

function Resolve-SmartM365AdReferenceXlsx {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)

    $leafName = [System.IO.Path]::GetFileName($Name)
    if ([string]::IsNullOrWhiteSpace($leafName) -or $leafName -ne $Name -or -not $leafName.EndsWith('.xlsx', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw ("Invalid AD reference workbook name: {0}" -f $Name)
    }

    if ($script:SmartM365AdReferenceXlsxCache.ContainsKey($leafName)) {
        $cachedPath = [string]$script:SmartM365AdReferenceXlsxCache[$leafName]
        if (Test-Path -LiteralPath $cachedPath -PathType Leaf) { return $cachedPath }
        $script:SmartM365AdReferenceXlsxCache.Remove($leafName)
    }

    if (-not (Get-Command Invoke-SmartM365SharePointFileDownload -ErrorAction SilentlyContinue)) {
        WriteLog -Message ("AD reference workbook cannot be recovered because the SharePoint download helper is unavailable: {0}" -f $leafName) -Level 'WARNING'
        return ''
    }

    if (-not (Test-Path -LiteralPath $script:SmartM365AdReferenceXlsxCacheRoot)) {
        New-Item -Path $script:SmartM365AdReferenceXlsxCacheRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $cachePath = Join-Path -Path $script:SmartM365AdReferenceXlsxCacheRoot -ChildPath $leafName
    $sharePointRelativePath = 'DATA-ALL/{0}' -f $leafName
    WriteLog -Message ("AD reference workbook not found locally; attempting SharePoint recovery: {0}" -f $sharePointRelativePath)
    $downloaded = Invoke-SmartM365SharePointFileDownload `
        -LocalFilePath $cachePath `
        -SharePointRelativePath $sharePointRelativePath `
        -Enabled $true `
        -Force
    if ($downloaded -and (Test-Path -LiteralPath $downloaded.FullName -PathType Leaf)) {
        $script:SmartM365AdReferenceXlsxCache[$leafName] = $downloaded.FullName
        return $downloaded.FullName
    }

    WriteLog -Message ("AD reference workbook not found in SharePoint; related enrichment columns will be blank: {0}" -f $sharePointRelativePath) -Level 'WARNING'
    return ''
}

function Clear-SmartM365AdReferenceXlsxCache {
    [CmdletBinding()]
    param()

    $script:SmartM365AdReferenceXlsxCache = @{}
    if (Test-Path -LiteralPath $script:SmartM365AdReferenceXlsxCacheRoot) {
        Remove-Item -LiteralPath $script:SmartM365AdReferenceXlsxCacheRoot -Recurse -Force -ErrorAction Stop
        WriteLog -Message ("AD reference workbook cache removed: {0}" -f $script:SmartM365AdReferenceXlsxCacheRoot)
    }
}

$adEnrichmentHelperPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-ActiveDirectory-Enrichment.ps1'
if (Test-Path -LiteralPath $adEnrichmentHelperPath) {
    . $adEnrichmentHelperPath
    WriteLog -Message ("Active Directory enrichment helper loaded: {0}" -f $adEnrichmentHelperPath)
}
else {
    WriteLog -Message ("WARNING: Active Directory enrichment helper not found. AD_Computers_AllDomains.csv will not be generated: {0}" -f $adEnrichmentHelperPath)
}

$adUsersEnrichmentHelperPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-ActiveDirectory-UsersEnrichment.ps1'
if (Test-Path -LiteralPath $adUsersEnrichmentHelperPath) {
    . $adUsersEnrichmentHelperPath
    WriteLog -Message ("Active Directory users enrichment helper loaded: {0}" -f $adUsersEnrichmentHelperPath)
}
else {
    WriteLog -Message ("WARNING: Active Directory users enrichment helper not found. AD_Users_AllDomains.csv will not be generated: {0}" -f $adUsersEnrichmentHelperPath)
}
# ==========================================================
# MAIN TRY / CATCH / FINALLY
# ==========================================================
try {
    if (-not $ReportOnly -and -not $DuplicateAnalysisOnly) {
        Import-Module ActiveDirectory -ErrorAction Stop
        WriteLog -Message "ActiveDirectory module imported successfully."
        Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredModules @('ActiveDirectory') -RequireActiveDirectoryRead | Out-Null
    }
    else {
        Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) | Out-Null
        if ($DuplicateAnalysisOnly) {
            WriteLog -Message "DuplicateAnalysisOnly mode enabled. Live Active Directory inventory collection will be skipped."
        }
        else {
            WriteLog -Message "ReportOnly mode enabled. Live Active Directory inventory collection will be skipped."
        }
    }
    $RemoteRoutingDomain = ''
    if (-not $ReportOnly -and -not $DomainWorker) {
        $remoteRoutingDomainResolution = Get-SmartM365RemoteRoutingDomain -Config $ScriptLocalConfig
        $RemoteRoutingDomain = [string]$remoteRoutingDomainResolution.Domain
        WriteLog -Message ("Remote routing domain resolved to '{0}' from {1}." -f $RemoteRoutingDomain, $remoteRoutingDomainResolution.Source)
    }


    # ----------------------------------------------------------
    # RETRY CONFIGURATION
    # ----------------------------------------------------------
    $MaxRetries        = 3
    $RetryDelaySeconds = 30

    # ----------------------------------------------------------
    # TRANSIENT AD ERROR DETECTION
    # ----------------------------------------------------------
    function Test-IsTransientADError {
        param(
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.ErrorRecord]$ErrorRecord
        )

        $ex = $ErrorRecord.Exception
        while ($null -ne $ex) {
            if ($ex -is [Microsoft.ActiveDirectory.Management.ADServerDownException]) { return $true }
            if ($ex.Message -match 'Unable to contact the server')                    { return $true }
            if ($ex.Message -match 'The server is not operational')                   { return $true }
            if ($ex.Message -match 'invalid enumeration context')                     { return $true }
            $ex = $ex.InnerException
        }
        return $false
    }

    $utcDate     = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $baseFolder  = Join-Path $OutputPath "Not-CSV-Combined"
    $tempFolder = $null
    if ($DuplicateAnalysisOnly) {
        WriteLog -Message "Temporary per-domain export folder skipped in DuplicateAnalysisOnly mode."
    }
    elseif ($DomainWorker -and -not [string]::IsNullOrWhiteSpace($DomainWorkerTempFolder)) {
        $tempFolder = $DomainWorkerTempFolder
    }
    else {
        $tempFolder = Join-Path $baseFolder $utcDate
    }

    if (-not [string]::IsNullOrWhiteSpace($tempFolder)) {
        $null = New-Item -ItemType Directory -Path $tempFolder -Force
        WriteLog -Message ("Temporary per-domain export folder ready: {0}" -f $tempFolder)
    }
    function Get-DomainNameShort {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainName
        )

        if ([string]::IsNullOrWhiteSpace($DomainName)) {
            return $null
        }

        $domainNameLower = $DomainName.Trim().ToLowerInvariant()
        $configuredName = $DomainFriendlyNames.PSObject.Properties[$domainNameLower]
        if ($null -ne $configuredName -and -not [string]::IsNullOrWhiteSpace([string]$configuredName.Value)) {
            return [string]$configuredName.Value
        }

        $dotPos = $domainNameLower.IndexOf('.')
        if ($dotPos -gt 0) {
            return $domainNameLower.Substring(0, $dotPos).ToUpperInvariant()
        }

        return $domainNameLower.ToUpperInvariant()
    }

    function Get-AdServerForDistinguishedName {
        param(
            [AllowNull()][string]$DistinguishedName,
            [Parameter(Mandatory = $true)][string]$FallbackServer
        )

        if ([string]::IsNullOrWhiteSpace($DistinguishedName)) {
            return $FallbackServer
        }

        $domainComponents = @([regex]::Matches($DistinguishedName, '(?i)(?:^|,)DC=([^,]+)') | ForEach-Object {
            $_.Groups[1].Value
        })
        if ($domainComponents.Count -eq 0) {
            return $FallbackServer
        }

        return ($domainComponents -join '.')
    }

    function Get-NormalizedDomainAndSam {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainNameShort,

            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$SamAccountName
        )

        $domainPart = if ([string]::IsNullOrWhiteSpace($DomainNameShort)) { '' } else { $DomainNameShort.Trim() }
        $samPart    = if ([string]::IsNullOrWhiteSpace($SamAccountName)) { '' } else { $SamAccountName.Trim() }

        if ($samPart.EndsWith('$')) {
            $samPart = $samPart.Substring(0, $samPart.Length - 1)
        }

        $value = "{0}\{1}" -f $domainPart, $samPart
        $value = $value.Trim().ToLowerInvariant()
        $value = $value -replace '\u00A0', ''
        $value = $value -replace ' ', ''
        $value = $value -replace '\t', ''

        return $value
    }

    function Convert-GuidToImmutableId {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$ObjectGuid
        )
        try {
            if ([string]::IsNullOrWhiteSpace($ObjectGuid)) {
                return $null
            }
            $guid  = [System.Guid]::Parse($ObjectGuid)
            $bytes = $guid.ToByteArray()
            return [System.Convert]::ToBase64String($bytes)
        }
        catch {
            return $null
        }
    }

    function Remove-OldFiles {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$Filter,

            [Parameter(Mandatory = $true)]
            [int]$OlderThanDays
        )

        if (-not (Test-Path -Path $Path)) {
            return
        }

        $limit = (Get-Date).AddDays(-$OlderThanDays)
        Get-ChildItem -Path $Path -Filter $Filter -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $limit } |
            ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -ErrorAction Stop
                    WriteLog -Message ("Deleted old file: {0}" -f $_.FullName)
                }
                catch {
                    WriteLog -Message ("Failed to delete old file '{0}': {1}" -f $_.FullName, $_)
                }
            }
    }

    function Invoke-SmartM365AdFileOperationWithRetry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$OperationName,
            [Parameter(Mandatory = $true)][string]$Path,
            [Parameter(Mandatory = $true)][scriptblock]$Action,
            [int]$MaxAttempts = 12,
            [int]$InitialDelaySeconds = 5,
            [int]$MaxDelaySeconds = 30
        )

        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            try {
                return & $Action
            }
            catch {
                if ($attempt -ge $MaxAttempts) { throw }

                $delaySeconds = [Math]::Min($MaxDelaySeconds, ($InitialDelaySeconds * $attempt))
                WriteLog -Message ("File operation retry: {0} for '{1}' failed on attempt {2}/{3}; waiting {4}s. Error: {5}" -f $OperationName, $Path, $attempt, $MaxAttempts, $delaySeconds, $_.Exception.Message) -Level 'INFO'
                Start-Sleep -Seconds $delaySeconds
            }
        }
    }

    function Remove-SmartM365AdFileWithRetry {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return }

        Invoke-SmartM365AdFileOperationWithRetry -OperationName 'remove file' -Path $Path -Action {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
        } | Out-Null
    }

    function Copy-SmartM365AdFileWithRetry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$SourcePath,
            [Parameter(Mandatory = $true)][string]$DestinationPath
        )

        Invoke-SmartM365AdFileOperationWithRetry -OperationName 'publish file' -Path $DestinationPath -Action {
            Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force -ErrorAction Stop
        } | Out-Null
    }

    function Add-SmartM365AdGeneratedCsvPath {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path
        )

        if ([string]::IsNullOrWhiteSpace($Path)) { return }

        if (-not $global:csvGeneratedPaths -or -not ($global:csvGeneratedPaths -is [System.Collections.Generic.HashSet[string]])) {
            $existingCsvGeneratedPaths = @($global:csvGeneratedPaths)
            $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($existingCsvGeneratedPath in $existingCsvGeneratedPaths) {
                if (-not [string]::IsNullOrWhiteSpace([string]$existingCsvGeneratedPath)) {
                    [void]$global:csvGeneratedPaths.Add([string]$existingCsvGeneratedPath)
                }
            }
        }

        [void]$global:csvGeneratedPaths.Add($Path)
    }
    function Combine-CsvFiles {
        param(
            [Parameter(Mandatory = $true)]
            [string]$SourceFolder,

            [Parameter(Mandatory = $true)]
            [string]$Filter,

            [Parameter(Mandatory = $true)]
            [string]$DestinationFile
        )

        $files = Get-ChildItem -Path $SourceFolder -Filter $Filter -File | Sort-Object Name
        if (-not $files) {
            if (Test-Path -LiteralPath $DestinationFile) {
                Remove-SmartM365AdFileWithRetry -Path $DestinationFile
                WriteLog -Message ("Removed stale combined CSV because no source files were found for filter '{0}': {1}" -f $Filter, $DestinationFile)
            }
            WriteLog -Message ("No CSV files found for filter '{0}' in '{1}'" -f $Filter, $SourceFolder)
            return
        }

        $destinationFolder = Split-Path -Path $DestinationFile -Parent
        if (-not [string]::IsNullOrWhiteSpace($destinationFolder) -and -not (Test-Path -LiteralPath $destinationFolder)) {
            New-Item -Path $destinationFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $destinationFileName = [System.IO.Path]::GetFileName($DestinationFile)
        $stagingName = "{0}.{1}.{2}.tmp" -f $destinationFileName, (Get-Date).ToString('yyyyMMddHHmmssfff'), $PID
        $stagingFile = Join-Path -Path $destinationFolder -ChildPath $stagingName

        $isFirstFile = $true
        [int64]$combinedRowCount = 0

        try {
            foreach ($file in $files) {
                $rows = @(Invoke-SmartM365AdCsvReadWithRetry -Path $file.FullName -ReadAction {
                    Import-Csv -LiteralPath $Path -ErrorAction Stop
                })
                if ($rows.Count -eq 0) { continue }

                if ($isFirstFile) {
                    $rows | Add-SmartM365TenantKey | Export-Csv -LiteralPath $stagingFile -NoTypeInformation -Encoding UTF8
                    $isFirstFile = $false
                }
                else {
                    $rows | Add-SmartM365TenantKey | Export-Csv -LiteralPath $stagingFile -NoTypeInformation -Encoding UTF8 -Append
                }

                $combinedRowCount += $rows.Count
            }

            if ($isFirstFile) {
                Remove-SmartM365AdFileWithRetry -Path $stagingFile
                if (Test-Path -LiteralPath $DestinationFile) {
                    Remove-SmartM365AdFileWithRetry -Path $DestinationFile
                    WriteLog -Message ("Removed stale combined CSV because no data rows were found for filter '{0}': {1}" -f $Filter, $DestinationFile)
                }
                WriteLog -Message ("No data rows found while combining filter '{0}' in '{1}'" -f $Filter, $SourceFolder)
                return
            }

            $combinedRowsForValidation = @(Invoke-SmartM365AdCsvReadWithRetry -Path $stagingFile -ReadAction {
                Import-Csv -LiteralPath $Path -ErrorAction Stop
            })
            Assert-SmartM365CsvDataCompleteness -Data $combinedRowsForValidation -TimestampedPath $stagingFile -LatestPath $DestinationFile

            Copy-SmartM365AdFileWithRetry -SourcePath $stagingFile -DestinationPath $DestinationFile
            Add-SmartM365AdGeneratedCsvPath -Path $DestinationFile
            WriteLog -Message ("Combined {0} file(s), {1} row(s), into '{2}'" -f $files.Count, $combinedRowCount, $DestinationFile)
        }
        finally {
            try {
                Remove-SmartM365AdFileWithRetry -Path $stagingFile
            }
            catch {
                WriteLog -Message ("Failed to delete temporary combined CSV staging file '{0}': {1}" -f $stagingFile, $_.Exception.Message) -Level 'WARNING'
            }
        }
    }
    function Publish-WeeklyInventoryHistoryToSharePoint {
        param(
            [Parameter(Mandatory = $true)]
            [string]$WeekName,

            [Parameter(Mandatory = $true)]
            [string]$WeekFolder,

            [Parameter(Mandatory = $true)]
            [string]$ManifestPath,

            [Parameter(Mandatory = $true)]
            [object]$Manifest,

            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [string[]]$RequiredFileNames
        )

        if (-not $global:EnableSharePointUpload) {
            return $false
        }
        if (-not (Get-Command Invoke-SmartM365SharePointCsvUpload -ErrorAction SilentlyContinue)) {
            WriteLog -Message "Weekly AD inventory history SharePoint publication skipped: upload helper is unavailable." -Level 'WARNING'
            return $false
        }
        if ([string]$Manifest.SharePointStatus -eq 'Complete') {
            WriteLog -Message ("Weekly AD inventory history SharePoint publication already complete for {0}. Upload skipped." -f $WeekName)
            return $true
        }

        $fileNames = @(
            @($Manifest.Files)
            @($Manifest.RequiredFiles)
            @($RequiredFileNames)
        ) |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            ForEach-Object { [System.IO.Path]::GetFileName([string]$_) } |
            Select-Object -Unique

        $missingFiles = @($fileNames | Where-Object {
            -not (Test-Path -LiteralPath (Join-Path -Path $WeekFolder -ChildPath $_))
        })
        if ($missingFiles.Count -gt 0) {
            $Manifest | Add-Member -NotePropertyName SharePointStatus -NotePropertyValue 'Incomplete' -Force
            $Manifest | Add-Member -NotePropertyName SharePointPublishedAt -NotePropertyValue $null -Force
            $Manifest | Add-Member -NotePropertyName SharePointFailedFiles -NotePropertyValue $missingFiles -Force
            $Manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
            WriteLog -Message ("Weekly AD inventory history SharePoint publication postponed because snapshot files are missing for {0}: {1}" -f $WeekName, ($missingFiles -join ', ')) -Level 'WARNING'
            return $false
        }

        WriteLog -Message ("Weekly AD inventory history SharePoint publication started for {0}: {1} data file(s) plus manifest." -f $WeekName, $fileNames.Count)
        $failedFiles = New-Object System.Collections.Generic.List[string]
        foreach ($fileName in $fileNames) {
            $filePath = Join-Path -Path $WeekFolder -ChildPath $fileName
            $uploadRecord = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $filePath
            if (-not $uploadRecord) {
                [void]$failedFiles.Add($fileName)
            }
        }

        if ($failedFiles.Count -gt 0) {
            $Manifest | Add-Member -NotePropertyName SharePointStatus -NotePropertyValue 'Incomplete' -Force
            $Manifest | Add-Member -NotePropertyName SharePointPublishedAt -NotePropertyValue $null -Force
            $Manifest | Add-Member -NotePropertyName SharePointFailedFiles -NotePropertyValue $failedFiles.ToArray() -Force
            $Manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
            WriteLog -Message ("Weekly AD inventory history SharePoint publication incomplete for {0}. Failed files: {1}" -f $WeekName, ($failedFiles -join ', ')) -Level 'WARNING'
            return $false
        }

        $Manifest | Add-Member -NotePropertyName SharePointStatus -NotePropertyValue 'Complete' -Force
        $Manifest | Add-Member -NotePropertyName SharePointPublishedAt -NotePropertyValue (Get-Date).ToString('o') -Force
        $Manifest | Add-Member -NotePropertyName SharePointFailedFiles -NotePropertyValue @() -Force
        $Manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8

        $manifestUploadRecord = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $ManifestPath
        if (-not $manifestUploadRecord) {
            $Manifest | Add-Member -NotePropertyName SharePointStatus -NotePropertyValue 'Incomplete' -Force
            $Manifest | Add-Member -NotePropertyName SharePointPublishedAt -NotePropertyValue $null -Force
            $Manifest | Add-Member -NotePropertyName SharePointFailedFiles -NotePropertyValue @('manifest.json') -Force
            $Manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
            WriteLog -Message ("Weekly AD inventory history data files were published for {0}, but manifest.json upload failed. Publication will be retried." -f $WeekName) -Level 'WARNING'
            return $false
        }

        WriteLog -Message ("Weekly AD inventory history SharePoint publication completed for {0}: {1} file(s)." -f $WeekName, ($fileNames.Count + 1))
        return $true
    }

    function Save-WeeklyInventoryHistory {
        param(
            [Parameter(Mandatory = $true)]
            [string[]]$SourceFiles,

            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [string[]]$RequiredSourceFiles,

            [Parameter(Mandatory = $true)]
            [string]$HistoryRootPath,

            [Parameter(Mandatory = $false)]
            [int]$RetentionWeeks = 52
        )

        if ([string]::IsNullOrWhiteSpace($HistoryRootPath)) {
            WriteLog -Message "Weekly AD inventory history skipped: WeeklyHistoryFolderPath is empty."
            return
        }

        $missingRequiredFiles = @($RequiredSourceFiles | Where-Object {
            [string]::IsNullOrWhiteSpace($_) -or -not (Test-Path -LiteralPath $_)
        })
        if ($missingRequiredFiles.Count -gt 0) {
            WriteLog -Message ("Weekly AD inventory history postponed because required canonical CSV files are missing: {0}" -f ($missingRequiredFiles -join ', ')) -Level 'WARNING'
            return
        }

        $existingSourceFiles = @($SourceFiles |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_) } |
            Select-Object -Unique)
        if ($existingSourceFiles.Count -eq 0) {
            WriteLog -Message "Weekly AD inventory history skipped: no source CSV file found."
            return
        }

        $now = Get-Date
        $isoYear = [System.Globalization.ISOWeek]::GetYear($now)
        $isoWeek = [System.Globalization.ISOWeek]::GetWeekOfYear($now)
        $weekName = "{0}-W{1:00}" -f $isoYear, $isoWeek
        $weekFolder = Join-Path -Path $HistoryRootPath -ChildPath $weekName
        $manifestPath = Join-Path -Path $weekFolder -ChildPath 'manifest.json'
        $requiredFileNames = @($RequiredSourceFiles | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Select-Object -Unique)

        if (Test-Path -LiteralPath $weekFolder) {
            $manifest = $null
            if (Test-Path -LiteralPath $manifestPath) {
                try {
                    $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    WriteLog -Message ("Existing weekly manifest is unreadable and will be rebuilt: {0}" -f $manifestPath) -Level 'WARNING'
                }
            }

            $requiredHistoryFilesExist = @($requiredFileNames | Where-Object {
                -not (Test-Path -LiteralPath (Join-Path -Path $weekFolder -ChildPath $_))
            }).Count -eq 0
            if ($null -ne $manifest -and $manifest.Status -eq 'Complete' -and $requiredHistoryFilesExist) {
                WriteLog -Message ("Weekly AD inventory history already exists for {0}. Snapshot skipped: {1}" -f $weekName, $weekFolder)
                Publish-WeeklyInventoryHistoryToSharePoint `
                    -WeekName $weekName `
                    -WeekFolder $weekFolder `
                    -ManifestPath $manifestPath `
                    -Manifest $manifest `
                    -RequiredFileNames $requiredFileNames | Out-Null
                return
            }

            WriteLog -Message ("Weekly AD inventory history for {0} is incomplete and will be rebuilt: {1}" -f $weekName, $weekFolder) -Level 'WARNING'
            if (Test-Path -LiteralPath $manifestPath) {
                Remove-Item -LiteralPath $manifestPath -Force -ErrorAction Stop
            }
        }

        New-Item -Path $weekFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        foreach ($rawHistoryName in @('AD_Users_AllDomains_Brut.csv', 'AD_Computers_AllDomains_Brut.csv')) {
            $rawHistoryPath = Join-Path -Path $weekFolder -ChildPath $rawHistoryName
            if (Test-Path -LiteralPath $rawHistoryPath) {
                Remove-Item -LiteralPath $rawHistoryPath -Force -ErrorAction Stop
            }
        }

        $copiedFiles = New-Object System.Collections.Generic.List[string]
        foreach ($sourceFile in $existingSourceFiles) {
            $destinationFile = Join-Path -Path $weekFolder -ChildPath ([System.IO.Path]::GetFileName($sourceFile))
            Copy-Item -LiteralPath $sourceFile -Destination $destinationFile -Force -ErrorAction Stop
            [void]$copiedFiles.Add($destinationFile)
        }

        $manifest = [PSCustomObject][ordered]@{
            Status           = 'Complete'
            CreatedAt        = (Get-Date).ToString('o')
            Week             = $weekName
            SourceOutputPath = $OutputPath
            RequiredFiles         = $requiredFileNames
            Files                 = @($copiedFiles | ForEach-Object { [System.IO.Path]::GetFileName($_) })
            SharePointStatus      = if ($global:EnableSharePointUpload) { 'Pending' } else { 'NotEnabled' }
            SharePointPublishedAt = $null
            SharePointFailedFiles = @()
        }
        $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

        WriteLog -Message ("Weekly AD inventory history saved for {0}: {1} file(s) in {2}" -f $weekName, $copiedFiles.Count, $weekFolder)
        Publish-WeeklyInventoryHistoryToSharePoint `
            -WeekName $weekName `
            -WeekFolder $weekFolder `
            -ManifestPath $manifestPath `
            -Manifest $manifest `
            -RequiredFileNames $requiredFileNames | Out-Null

        if ($RetentionWeeks -gt 0) {
            $oldWeekFolders = @(Get-ChildItem -LiteralPath $HistoryRootPath -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^\d{4}-W\d{2}$' } |
                Sort-Object Name -Descending |
                Select-Object -Skip $RetentionWeeks)

            foreach ($oldWeekFolder in $oldWeekFolders) {
                try {
                    Remove-Item -LiteralPath $oldWeekFolder.FullName -Recurse -Force -ErrorAction Stop
                    WriteLog -Message ("Deleted old weekly AD inventory history folder: {0}" -f $oldWeekFolder.FullName)
                }
                catch {
                    WriteLog -Message ("Failed to delete old weekly AD inventory history folder '{0}': {1}" -f $oldWeekFolder.FullName, $_) -Level "WARNING"
                }
            }
        }
    }

    function Remove-TemporaryInventoryFolder {
        param(
            [Parameter(Mandatory = $true)]
            [string]$TempFolder,

            [Parameter(Mandatory = $true)]
            [string]$BaseFolder
        )

        if (-not $DeleteTemporaryPerDomainCsv) {
            WriteLog -Message ("Temporary per-domain CSV folder kept because DeleteTemporaryPerDomainCsv is disabled: {0}" -f $TempFolder)
            return
        }

        if ([string]::IsNullOrWhiteSpace($TempFolder) -or -not (Test-Path -LiteralPath $TempFolder)) {
            return
        }

        try {
            $resolvedTemp = (Resolve-Path -LiteralPath $TempFolder -ErrorAction Stop).Path.TrimEnd('\')
            $resolvedBase = (Resolve-Path -LiteralPath $BaseFolder -ErrorAction Stop).Path.TrimEnd('\')
            $basePrefix = $resolvedBase + '\'

            if ($resolvedTemp -eq $resolvedBase -or -not $resolvedTemp.StartsWith($basePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                WriteLog -Message ("Temporary per-domain CSV cleanup skipped because the path is outside the expected base folder. Temp: {0}. Base: {1}" -f $resolvedTemp, $resolvedBase) -Level "WARNING"
                return
            }

            Remove-Item -LiteralPath $resolvedTemp -Recurse -Force -ErrorAction Stop
            WriteLog -Message ("Deleted temporary per-domain CSV folder: {0}" -f $resolvedTemp)
        }
        catch {
            WriteLog -Message ("Failed to delete temporary per-domain CSV folder '{0}': {1}" -f $TempFolder, $_) -Level "WARNING"
        }
    }

    function Remove-StaleTemporaryInventoryFolders {
        param(
            [Parameter(Mandatory = $true)]
            [string]$BaseFolder,

            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$CurrentTempFolder,

            [Parameter(Mandatory = $false)]
            [int]$RetentionDays = 2
        )

        if (-not $DeleteTemporaryPerDomainCsv) {
            WriteLog -Message "Stale temporary per-domain cleanup skipped because DeleteTemporaryPerDomainCsv is disabled."
            return
        }
        if ($RetentionDays -le 0) {
            WriteLog -Message "Stale temporary per-domain cleanup disabled because TemporaryPerDomainRetentionDays is 0 or less."
            return
        }
        if ([string]::IsNullOrWhiteSpace($BaseFolder) -or -not (Test-Path -LiteralPath $BaseFolder)) {
            return
        }

        $resolvedBasePath = Resolve-Path -LiteralPath $BaseFolder -ErrorAction Stop
        $resolvedBase = $resolvedBasePath.ProviderPath.TrimEnd('\')
        $basePrefix = $resolvedBase + '\'
        $resolvedCurrent = $null
        if (-not [string]::IsNullOrWhiteSpace($CurrentTempFolder) -and (Test-Path -LiteralPath $CurrentTempFolder)) {
            $resolvedCurrentPath = Resolve-Path -LiteralPath $CurrentTempFolder -ErrorAction Stop
            $resolvedCurrent = $resolvedCurrentPath.ProviderPath.TrimEnd('\')
        }
        $cutoffUtc = (Get-Date).ToUniversalTime().AddDays(-$RetentionDays)

        foreach ($folder in @(Get-ChildItem -LiteralPath $resolvedBase -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d{8}T\d{6}Z$' -and $_.LastWriteTimeUtc -lt $cutoffUtc })) {
            $resolvedFolder = $folder.FullName.TrimEnd('\')
            if ($resolvedFolder -eq $resolvedCurrent) { continue }
            if (-not $resolvedFolder.StartsWith($basePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                WriteLog -Message ("Stale temporary folder cleanup skipped outside the expected base folder: {0}" -f $resolvedFolder) -Level 'WARNING'
                continue
            }

            try {
                Remove-Item -LiteralPath $resolvedFolder -Recurse -Force -ErrorAction Stop
                WriteLog -Message ("Deleted abandoned temporary per-domain folder older than {0} day(s): {1}" -f $RetentionDays, $resolvedFolder)
            }
            catch {
                WriteLog -Message ("Failed to delete abandoned temporary per-domain folder '{0}': {1}" -f $resolvedFolder, $_) -Level 'WARNING'
            }
        }
    }

    if (-not $DomainWorker -and -not $DuplicateAnalysisOnly) {
        Remove-StaleTemporaryInventoryFolders -BaseFolder $baseFolder -CurrentTempFolder $tempFolder -RetentionDays $TemporaryPerDomainRetentionDays
    }

    function ConvertTo-DailyReportBoolean {
        [CmdletBinding()]
        param([AllowNull()]$Value)

        if ($Value -is [bool]) { return $Value }
        if ($null -eq $Value) { return $false }
        return ([string]$Value).Trim() -eq 'True'
    }

    function ConvertTo-DailyReportDate {
        [CmdletBinding()]
        param([AllowNull()]$Value)

        if ($Value -is [datetime]) { return $Value }
        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

        $parsedDate = [datetime]::MinValue
        if ([datetime]::TryParse([string]$Value, [ref]$parsedDate)) { return $parsedDate }
        return $null
    }

    function Get-DailyReportSimpleOS {
        [CmdletBinding()]
        param([AllowNull()][string]$OperatingSystem)

        switch -Wildcard ($OperatingSystem) {
            '*Windows 11*' { return 'Windows 11' }
            '*Windows 10*' { return 'Windows 10' }
            '*Windows 8*'  { return 'Windows 8' }
            '*Windows 7*'  { return 'Windows 7' }
            default        { return 'Unknown' }
        }
    }

    function Test-DailyReportWildcardMatch {
        [CmdletBinding()]
        param(
            [AllowNull()][string]$Value,
            [string[]]$Patterns
        )

        foreach ($pattern in @($Patterns)) {
            if ($Value -like $pattern) { return $true }
        }
        return $false
    }

    function Get-InventoryDailyReportSource {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$SourceFolder,
            [Parameter(Mandatory = $true)][string[]]$AllowedOS,
            [string[]]$TargetDomains
        )

        if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
            WriteLog -Message 'Daily report source skipped: source folder is empty.' -Level 'WARNING'
            return $null
        }

        $computersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Computers_AllDomains.csv'
        if (-not (Test-Path -LiteralPath $computersCsv)) { $computersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Computers_AllDomains_Brut.csv' }
        $usersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Users_AllDomains.csv'
        if (-not (Test-Path -LiteralPath $usersCsv)) { $usersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Users_AllDomains_Brut.csv' }

        foreach ($sourceCsv in @($computersCsv, $usersCsv)) {
            if (-not (Test-Path -LiteralPath $sourceCsv)) {
                WriteLog -Message ("Daily report source unavailable. Missing file: {0}" -f $sourceCsv) -Level 'WARNING'
                return $null
            }
        }

        $targetDomainSet = @{}
        foreach ($domain in @($TargetDomains | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            $targetDomainSet[$domain.ToLowerInvariant()] = $true
        }
        $hasTargetDomainFilter = $targetDomainSet.Count -gt 0

        $computers = @(Import-Csv -LiteralPath $computersCsv -Encoding UTF8 | Where-Object {
            $domainName = [string]$_.DomainName
            ((-not $hasTargetDomainFilter) -or $targetDomainSet.ContainsKey($domainName.ToLowerInvariant())) -and
            (Test-DailyReportWildcardMatch -Value ([string]$_.OperatingSystem) -Patterns $AllowedOS)
        } | ForEach-Object {
            [PSCustomObject]@{
                Enabled                = ConvertTo-DailyReportBoolean $_.Enabled
                SimpleOS               = Get-DailyReportSimpleOS -OperatingSystem ([string]$_.OperatingSystem)
                ADDomain               = [string]$_.DomainName
                OperatingSystem        = [string]$_.OperatingSystem
                OperatingSystemVersion = [string]$_.OperatingSystemVersion
                LastLogonDate          = ConvertTo-DailyReportDate $_.LastLogonDate
            }
        })

        $users = @(Import-Csv -LiteralPath $usersCsv -Encoding UTF8 | Where-Object {
            $domainName = [string]$_.DomainName
            ((-not $hasTargetDomainFilter) -or $targetDomainSet.ContainsKey($domainName.ToLowerInvariant())) -and
            (-not [string]::IsNullOrWhiteSpace([string]$_.UserPrincipalName))
        } | ForEach-Object {
            [PSCustomObject]@{
                ADDomain          = [string]$_.DomainName
                Enabled           = ConvertTo-DailyReportBoolean $_.Enabled
                LastLogonDate     = ConvertTo-DailyReportDate $_.LastLogonDate
                UserPrincipalName = [string]$_.UserPrincipalName
            }
        })

        if ($computers.Count -eq 0 -and $users.Count -eq 0) {
            WriteLog -Message 'Daily report source produced no reportable rows.' -Level 'WARNING'
            return $null
        }

        $domains = @(@($computers | Select-Object -ExpandProperty ADDomain) + @($users | Select-Object -ExpandProperty ADDomain) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique)

        WriteLog -Message ("Daily report source loaded. Computers: {0}; Users: {1}; Domains: {2}" -f $computers.Count, $users.Count, ($domains -join ', '))
        return [PSCustomObject]@{
            Domains   = $domains
            Computers = $computers
            Users     = $users
        }
    }

    function Write-DailyReportCsv {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]$Rows,
            [Parameter(Mandatory = $true)][string]$OutputFilePath,
            [string]$LatestFolderPath
        )

        $rowsArray = @($Rows)
        if ($rowsArray.Count -eq 0) {
            WriteLog -Message ("Daily report CSV skipped because there are no rows: {0}" -f $OutputFilePath) -Level 'WARNING'
            return
        }

        $rowsArray = @($rowsArray | Add-SmartM365TenantKey)
        if (Test-Path -LiteralPath $OutputFilePath) {
            Repair-SmartM365CsvTenantKeySchema -Path $OutputFilePath -Delimiter ';' -Encoding UTF8 | Out-Null
            $rowsArray | ConvertTo-Csv -NoTypeInformation -Delimiter ';' | Select-Object -Skip 1 | Add-Content -Path $OutputFilePath -Encoding UTF8
            WriteLog -Message ("Daily report rows appended to CSV: {0}" -f $OutputFilePath)
        }
        else {
            $rowsArray | Add-SmartM365TenantKey | Export-Csv -Path $OutputFilePath -NoTypeInformation -Delimiter ';' -Encoding UTF8
            WriteLog -Message ("Daily report CSV created: {0}" -f $OutputFilePath)
        }
        if (-not $global:csvGeneratedPaths -or -not ($global:csvGeneratedPaths -is [System.Collections.Generic.HashSet[string]])) {
            $existingCsvGeneratedPaths = @($global:csvGeneratedPaths)
            $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($existingCsvGeneratedPath in $existingCsvGeneratedPaths) {
                if (-not [string]::IsNullOrWhiteSpace([string]$existingCsvGeneratedPath)) {
                    [void]$global:csvGeneratedPaths.Add([string]$existingCsvGeneratedPath)
                }
            }
        }
        [void]$global:csvGeneratedPaths.Add($OutputFilePath)
        if (-not [string]::IsNullOrWhiteSpace($LatestFolderPath)) {
            if (-not (Test-Path -LiteralPath $LatestFolderPath)) {
                New-Item -Path $LatestFolderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                WriteLog -Message ("Created missing LatestCsvFolderPath directory: {0}" -f $LatestFolderPath)
            }

            $latestFilePath = Join-Path -Path $LatestFolderPath -ChildPath ([System.IO.Path]::GetFileName($OutputFilePath))
            Copy-Item -LiteralPath $OutputFilePath -Destination $latestFilePath -Force -ErrorAction Stop
            WriteLog -Message ("Daily report CSV copied to LatestCsvFolderPath: {0}" -f $latestFilePath)
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestFilePath | Out-Null
        }

        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $OutputFilePath | Out-Null
    }

    function Invoke-ActiveDirectoryDailyReport {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$SourceFolder,
            [Parameter(Mandatory = $true)][string]$ReportOutputPath,
            [string]$LatestFolderPath,
            [string[]]$AllowedOS = @('Windows 7*','Windows 8*','Windows 10*','Windows 11*'),
            [string[]]$TargetDomains,
            [int]$InactiveDays = 90,
            [bool]$UseDailyLock = $true,
            [string]$LockRoot,
            [string]$LockName
        )

        if ($SkipDailyReport) {
            WriteLog -Message 'Active Directory daily report skipped because -SkipDailyReport was specified.'
            return $false
        }

        if (-not $EnableDailyReport) {
            WriteLog -Message 'Active Directory daily report skipped because EnableDailyReport is disabled.'
            return $false
        }

        $today = Get-Date -Format 'yyyy-MM-dd'
        $dailyLockFile = $null
        if ($UseDailyLock) {
            if ([string]::IsNullOrWhiteSpace($LockRoot)) { $LockRoot = 'C:\ProgramData\SmartM365\Locks' }
            if ([string]::IsNullOrWhiteSpace($LockName)) { $LockName = 'SmartM365-ActiveDirectory-Inventory-DailyReport' }
            if (-not (Test-Path -LiteralPath $LockRoot)) { New-Item -Path $LockRoot -ItemType Directory -Force | Out-Null }
            $dailyLockFile = Join-Path -Path $LockRoot -ChildPath ("{0}-SUCCESS-{1}.lock" -f $LockName, $today)
            if (Test-Path -LiteralPath $dailyLockFile) {
                WriteLog -Message ("[{0}] Daily report already succeeded on {1}. Skipping." -f $LockName, $today)
                return $false
            }
        }

        $reportSource = Get-InventoryDailyReportSource -SourceFolder $SourceFolder -AllowedOS $AllowedOS -TargetDomains $TargetDomains
        if ($null -eq $reportSource) { return $false }

        $nowText = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $osReportColumns = @('Windows 7', 'Windows 8', 'Windows 10', 'Windows 11')

        $computerRows = @($reportSource.Computers | Group-Object -Property ADDomain | ForEach-Object {
            $domainName = $_.Name
            $computersInGroup = $_.Group
            $enabledComputersInGroup = @($computersInGroup | Where-Object { $_.Enabled -eq $true })
            $disabledComputersInGroup = @($computersInGroup | Where-Object { $_.Enabled -eq $false })
            $countsByEnabledOS = @($enabledComputersInGroup | Group-Object -Property SimpleOS -NoElement)
            $countsByDisabledOS = @($disabledComputersInGroup | Group-Object -Property SimpleOS -NoElement)

            $outputRow = [ordered]@{
                Date             = $nowText
                DomainName       = $domainName
                TotalComputers   = $computersInGroup.Count
                EnabledAccounts  = $enabledComputersInGroup.Count
                DisabledAccounts = $disabledComputersInGroup.Count
            }

            foreach ($os in $osReportColumns) {
                $enabledOsGroup = @($countsByEnabledOS | Where-Object { $_.Name -eq $os } | Select-Object -First 1)
                $disabledOsGroup = @($countsByDisabledOS | Where-Object { $_.Name -eq $os } | Select-Object -First 1)
                $outputRow["$os Enabled"] = if ($enabledOsGroup.Count -gt 0) { $enabledOsGroup[0].Count } else { 0 }
                $outputRow["$os Disabled"] = if ($disabledOsGroup.Count -gt 0) { $disabledOsGroup[0].Count } else { 0 }
            }

            [PSCustomObject]$outputRow
        })

        $inactiveThreshold = (Get-Date).AddDays(-1 * $InactiveDays)
        $userRows = @($reportSource.Users | Where-Object { -not [string]::IsNullOrEmpty($_.ADDomain) } | Group-Object -Property ADDomain | ForEach-Object {
            $domainName = $_.Name
            $usersInGroup = $_.Group
            $enabledUsers = @($usersInGroup | Where-Object { $_.Enabled -eq $true })
            $disabledUsers = @($usersInGroup | Where-Object { $_.Enabled -eq $false })
            $inactiveUsers = @($usersInGroup | Where-Object { $null -eq $_.LastLogonDate -or $_.LastLogonDate -lt $inactiveThreshold })

            [PSCustomObject]([ordered]@{
                Date = $nowText
                DomainName = $domainName
                TotalUsers = $usersInGroup.Count
                EnabledUsers = $enabledUsers.Count
                DisabledUsers = $disabledUsers.Count
                Inactive90DaysUsers = $inactiveUsers.Count
            })
        })

        $computerReportPath = Join-Path -Path $ReportOutputPath -ChildPath 'AD_Computers_DailyStats.csv'
        $userReportPath = Join-Path -Path $ReportOutputPath -ChildPath 'AD_Users_DailyStats.csv'
        Write-DailyReportCsv -Rows $computerRows -OutputFilePath $computerReportPath -LatestFolderPath $LatestFolderPath
        Write-DailyReportCsv -Rows $userRows -OutputFilePath $userReportPath -LatestFolderPath $LatestFolderPath

        if ($UseDailyLock -and -not [string]::IsNullOrWhiteSpace($dailyLockFile)) {
            New-Item -Path $dailyLockFile -ItemType File -Force | Out-Null
            WriteLog -Message ("[{0}] Daily report success lock created for {1}." -f $LockName, $today)
        }

        WriteLog -Message 'Active Directory daily reports generated successfully.'
        return $true
    }
    function Invoke-SmartM365AdCsvReadWithRetry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$Path,
            [Parameter(Mandatory = $true)][scriptblock]$ReadAction,
            [int]$MaxAttempts = 6,
            [int]$InitialDelaySeconds = 2
        )

        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            try {
                return & $ReadAction
            }
            catch {
                if ($attempt -ge $MaxAttempts) { throw }

                $delaySeconds = [Math]::Min(30, ($InitialDelaySeconds * $attempt))
                WriteLog -Message ("CSV read retry for '{0}' after transient read failure: attempt {1}/{2}; waiting {3}s. Error: {4}" -f $Path, $attempt, $MaxAttempts, $delaySeconds, $_.Exception.Message) -Level 'INFO'
                Start-Sleep -Seconds $delaySeconds
            }
        }
    }

    function Get-SmartM365AdCsvDataRowCount {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return 0 }

        [int64]$lineCount = Invoke-SmartM365AdCsvReadWithRetry -Path $Path -ReadAction {
            [int64]$count = 0
            foreach ($line in [System.IO.File]::ReadLines($Path)) { $count++ }
            return $count
        }
        if ($lineCount -le 0) { return 0 }
        return [Math]::Max(0, $lineCount - 1)
    }

    function Get-SmartM365AdCsvDistinctValueCount {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$Path,
            [Parameter(Mandatory = $true)][string]$ColumnName
        )

        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return 0 }

        Invoke-SmartM365AdCsvReadWithRetry -Path $Path -ReadAction {
            $values = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($row in @(Import-Csv -LiteralPath $Path -Encoding UTF8)) {
                $property = $row.PSObject.Properties[$ColumnName]
                if ($null -eq $property) { continue }

                $value = [string]$property.Value
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    [void]$values.Add($value.Trim())
                }
            }

            return $values.Count
        }
    }
    function ConvertTo-SmartM365AdSummaryInt64 {
        [CmdletBinding()]
        param([AllowNull()]$Value)

        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return 0 }
        [int64]$parsed = 0
        if ([int64]::TryParse(([string]$Value).Trim(), [ref]$parsed)) { return $parsed }
        return 0
    }

    function Format-SmartM365AdSummaryNumber {
        [CmdletBinding()]
        param([AllowNull()]$Value)

        $number = ConvertTo-SmartM365AdSummaryInt64 -Value $Value
        return $number.ToString('N0', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    function Format-SmartM365AdSummaryDelta {
        [CmdletBinding()]
        param(
            [AllowNull()]$Current,
            [AllowNull()]$Previous,
            [bool]$HasPrevious
        )

        if (-not $HasPrevious) { return 'n/a' }

        $delta = (ConvertTo-SmartM365AdSummaryInt64 -Value $Current) - (ConvertTo-SmartM365AdSummaryInt64 -Value $Previous)
        if ($delta -gt 0) { return ('+{0}' -f (Format-SmartM365AdSummaryNumber -Value $delta)) }
        if ($delta -lt 0) { return ('-{0}' -f (Format-SmartM365AdSummaryNumber -Value ([Math]::Abs($delta)))) }
        return '0'
    }

    function Get-SmartM365AdDailySummarySnapshot {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$SourceFolder)

        if ([string]::IsNullOrWhiteSpace($SourceFolder) -or -not (Test-Path -LiteralPath $SourceFolder)) {
            WriteLog -Message ("AD daily summary source folder unavailable: {0}" -f $SourceFolder) -Level 'WARNING'
            return $null
        }

        $usersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Users_AllDomains.csv'
        if (-not (Test-Path -LiteralPath $usersCsv)) { $usersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Users_AllDomains_Brut.csv' }
        $computersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Computers_AllDomains.csv'
        if (-not (Test-Path -LiteralPath $computersCsv)) { $computersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Computers_AllDomains_Brut.csv' }
        $groupsCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Groups_AllDomains.csv'
        $ousCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_OUs_AllDomains.csv'
        $contactsCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Contacts_AllDomains.csv'
        $duplicateUpnCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Users_DuplicateUPN.csv'
        $duplicateSmtpCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Users_DuplicateSMTP.csv'

        foreach ($requiredCsv in @($usersCsv, $computersCsv, $groupsCsv, $ousCsv, $contactsCsv)) {
            if (-not (Test-Path -LiteralPath $requiredCsv)) {
                WriteLog -Message ("AD daily summary skipped because required source CSV is missing: {0}" -f $requiredCsv) -Level 'WARNING'
                return $null
            }
        }

        return [PSCustomObject][ordered]@{
            SnapshotDate                   = (Get-Date).ToString('yyyy-MM-dd')
            GeneratedAt                    = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
            DomainCount                    = Get-SmartM365AdCsvDistinctValueCount -Path $usersCsv -ColumnName 'DomainName'
            TotalUsers                     = Get-SmartM365AdCsvDataRowCount -Path $usersCsv
            TotalComputers                 = Get-SmartM365AdCsvDataRowCount -Path $computersCsv
            TotalGroups                    = Get-SmartM365AdCsvDataRowCount -Path $groupsCsv
            TotalOUs                       = Get-SmartM365AdCsvDataRowCount -Path $ousCsv
            TotalContacts                  = Get-SmartM365AdCsvDataRowCount -Path $contactsCsv
            DistinctDuplicateUPNs          = Get-SmartM365AdCsvDistinctValueCount -Path $duplicateUpnCsv -ColumnName 'UserPrincipalName'
            AffectedDuplicateUPNAccounts   = Get-SmartM365AdCsvDataRowCount -Path $duplicateUpnCsv
            DistinctDuplicateSMTPAddresses = Get-SmartM365AdCsvDistinctValueCount -Path $duplicateSmtpCsv -ColumnName 'SmtpAddress'
            AffectedDuplicateSMTPEntries   = Get-SmartM365AdCsvDataRowCount -Path $duplicateSmtpCsv
        }
    }

    function Get-SmartM365AdPreviousDailySummarySnapshot {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$SummaryCsvPath)

        if ([string]::IsNullOrWhiteSpace($SummaryCsvPath) -or -not (Test-Path -LiteralPath $SummaryCsvPath)) { return $null }

        $today = (Get-Date).Date
        $previousRows = @(
            foreach ($row in @(Import-Csv -LiteralPath $SummaryCsvPath -Encoding UTF8)) {
                $snapshotDate = [datetime]::MinValue
                if (-not [datetime]::TryParse([string]$row.SnapshotDate, [ref]$snapshotDate)) { continue }
                if ($snapshotDate.Date -ge $today) { continue }

                [PSCustomObject]@{
                    SortDate = $snapshotDate.Date
                    Row      = $row
                }
            }
        )

        if ($previousRows.Count -eq 0) { return $null }
        return @($previousRows | Sort-Object SortDate -Descending | Select-Object -First 1)[0].Row
    }

    function New-SmartM365AdDailySummaryDiffSection {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]$Current,
            [AllowNull()]$Previous
        )

        $hasPrevious = $null -ne $Previous
        $metrics = @(
            [pscustomobject]@{ Label = 'Domains'; Property = 'DomainCount' }
            [pscustomobject]@{ Label = 'Users'; Property = 'TotalUsers' }
            [pscustomobject]@{ Label = 'Computers'; Property = 'TotalComputers' }
            [pscustomobject]@{ Label = 'Groups'; Property = 'TotalGroups' }
            [pscustomobject]@{ Label = 'OUs'; Property = 'TotalOUs' }
            [pscustomobject]@{ Label = 'Contacts'; Property = 'TotalContacts' }
            [pscustomobject]@{ Label = 'Distinct duplicate UPNs'; Property = 'DistinctDuplicateUPNs' }
            [pscustomobject]@{ Label = 'Affected duplicate UPN accounts'; Property = 'AffectedDuplicateUPNAccounts' }
            [pscustomobject]@{ Label = 'Distinct duplicate SMTP addresses'; Property = 'DistinctDuplicateSMTPAddresses' }
            [pscustomobject]@{ Label = 'Affected duplicate SMTP entries'; Property = 'AffectedDuplicateSMTPEntries' }
        )

        $rows = foreach ($metric in $metrics) {
            $currentProperty = $Current.PSObject.Properties[$metric.Property]
            $currentValue = if ($null -ne $currentProperty) { $currentProperty.Value } else { 0 }
            $previousProperty = if ($hasPrevious) { $Previous.PSObject.Properties[$metric.Property] } else { $null }
            $previousValue = if ($null -ne $previousProperty) { $previousProperty.Value } else { $null }
            $delta = Format-SmartM365AdSummaryDelta -Current $currentValue -Previous $previousValue -HasPrevious $hasPrevious
            $deltaColor = '#334155'
            if ($delta -like '+*') { $deltaColor = '#1d4ed8' }
            elseif ($delta -like '-*') { $deltaColor = '#b91c1c' }

            $labelHtml = ConvertTo-SmartM365EmailHtmlText $metric.Label
            $currentHtml = ConvertTo-SmartM365EmailHtmlText (Format-SmartM365AdSummaryNumber -Value $currentValue)
            $previousHtml = if ($hasPrevious) { ConvertTo-SmartM365EmailHtmlText (Format-SmartM365AdSummaryNumber -Value $previousValue) } else { 'n/a' }
            $deltaHtml = ConvertTo-SmartM365EmailHtmlText $delta

            "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$labelHtml</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#111827;`">$currentHtml</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;`">$previousHtml</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:$deltaColor;`">$deltaHtml</td></tr>"
        }

        $previousLabel = if ($hasPrevious) { ConvertTo-SmartM365EmailHtmlText ("Previous older scan: {0}" -f $Previous.SnapshotDate) } else { 'No previous older daily summary snapshot found. This run becomes the comparison baseline.' }
        $html = @"
<div style="font-size:13px;color:#64748b;margin-bottom:8px;">$previousLabel</div>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Metric</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Current</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Previous</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Delta</th>
  </tr>
  $($rows -join "`n")
</table>
"@

        return [pscustomobject]@{ Title = 'Diff since previous scan'; Html = $html }
    }

    function Invoke-SmartM365AdDailySummaryEmail {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$SourceFolder,
            [Parameter(Mandatory = $true)][string]$SummaryOutputPath,
            [string]$LatestFolderPath,
            [string]$LastSentFilePath,
            [bool]$ForceSend = $false
        )

        if (-not $EnableDailySummaryEmail) {
            WriteLog -Message 'Active Directory daily summary email skipped because EnableDailySummaryEmail is disabled.'
            return $false
        }

        if ([string]::IsNullOrWhiteSpace($LastSentFilePath)) {
            $LastSentFilePath = Join-Path -Path $SummaryOutputPath -ChildPath 'AD_DailySummary_LastSent.txt'
        }

        $todayStamp = (Get-Date).ToString('yyyy-MM-dd')
        $lastSentStamp = ''
        if (Test-Path -LiteralPath $LastSentFilePath) {
            $lastSentStamp = (Get-Content -LiteralPath $LastSentFilePath -Raw -ErrorAction SilentlyContinue).Trim()
        }

        if ($lastSentStamp -eq $todayStamp -and -not $ForceSend) {
            WriteLog -Message ("Active Directory daily summary email already sent today ({0}). Use -ForceSendDailySummary to resend." -f $todayStamp)
            return $false
        }

        $summarySnapshot = Get-SmartM365AdDailySummarySnapshot -SourceFolder $SourceFolder
        if ($null -eq $summarySnapshot) { return $false }

        if (-not (Test-Path -LiteralPath $SummaryOutputPath)) {
            New-Item -Path $SummaryOutputPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $summaryCsvPath = Join-Path -Path $SummaryOutputPath -ChildPath 'AD_Inventory_DailySummary.csv'
        $previousSnapshot = Get-SmartM365AdPreviousDailySummarySnapshot -SummaryCsvPath $summaryCsvPath

        $summarySnapshot = $summarySnapshot | Add-SmartM365TenantKey
        if (Test-Path -LiteralPath $summaryCsvPath) {
            Repair-SmartM365CsvTenantKeySchema -Path $summaryCsvPath -Delimiter ',' -Encoding UTF8 | Out-Null
            $summarySnapshot | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Add-Content -LiteralPath $summaryCsvPath -Encoding UTF8
            WriteLog -Message ("AD daily summary snapshot appended: {0}" -f $summaryCsvPath)
        }
        else {
            $summarySnapshot | Add-SmartM365TenantKey | Export-Csv -LiteralPath $summaryCsvPath -NoTypeInformation -Encoding UTF8
            WriteLog -Message ("AD daily summary snapshot created: {0}" -f $summaryCsvPath)
        }

        if (-not $global:csvGeneratedPaths -or -not ($global:csvGeneratedPaths -is [System.Collections.Generic.HashSet[string]])) {
            $existingCsvGeneratedPaths = @($global:csvGeneratedPaths)
            $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($existingCsvGeneratedPath in $existingCsvGeneratedPaths) {
                if (-not [string]::IsNullOrWhiteSpace([string]$existingCsvGeneratedPath)) {
                    [void]$global:csvGeneratedPaths.Add([string]$existingCsvGeneratedPath)
                }
            }
        }
        [void]$global:csvGeneratedPaths.Add($summaryCsvPath)

        $summarySharePointUploads = @()
        $sourceUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $summaryCsvPath
        $sourceUpload = Add-SmartM365SharePointUploadLabel -UploadRecord $sourceUpload -Label 'AD daily summary (DATA-ALL)'
        if ($sourceUpload) { $summarySharePointUploads += $sourceUpload }

        $latestSummaryPath = $null
        if (-not [string]::IsNullOrWhiteSpace($LatestFolderPath)) {
            if (-not (Test-Path -LiteralPath $LatestFolderPath)) {
                New-Item -Path $LatestFolderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                WriteLog -Message ("Created missing LatestCsvFolderPath directory: {0}" -f $LatestFolderPath)
            }

            $latestSummaryPath = Join-Path -Path $LatestFolderPath -ChildPath ([System.IO.Path]::GetFileName($summaryCsvPath))
            Copy-Item -LiteralPath $summaryCsvPath -Destination $latestSummaryPath -Force -ErrorAction Stop
            WriteLog -Message ("AD daily summary copied to LatestCsvFolderPath: {0}" -f $latestSummaryPath)

            $latestUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestSummaryPath
            $latestUpload = Add-SmartM365SharePointUploadLabel -UploadRecord $latestUpload -Label 'AD daily summary (DATA-LAST)'
            if ($latestUpload) { $summarySharePointUploads += $latestUpload }
        }

        $dailySummaryTo = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($dailySummaryTo)) {
            WriteLog -Message 'Active Directory daily summary email skipped because To is not configured in local or global configuration.' -Level 'WARNING'
            return $false
        }

        $diffSection = New-SmartM365AdDailySummaryDiffSection -Current $summarySnapshot -Previous $previousSnapshot
        $sharePointSection = New-SmartM365SharePointLinksSection -UploadRecords $summarySharePointUploads
        $sections = @($diffSection)
        if ($sharePointSection) { $sections += $sharePointSection }

        $hasDuplicateIdentities = ((ConvertTo-SmartM365AdSummaryInt64 -Value $summarySnapshot.AffectedDuplicateUPNAccounts) -gt 0) -or ((ConvertTo-SmartM365AdSummaryInt64 -Value $summarySnapshot.AffectedDuplicateSMTPEntries) -gt 0)
        $severity = if ($hasDuplicateIdentities) { 'Warning' } else { 'Success' }
        $actionTitle = if ($hasDuplicateIdentities) { 'Review required' } else { 'No duplicate identity conflict detected' }
        $actionHtml = if ($hasDuplicateIdentities) { 'Review duplicate UPN and SMTP counters before identity cleanup, migration, or synchronization decisions.' } else { 'Keep the generated CSV files as the daily Active Directory inventory baseline.' }

        $summaryRows = @(
            [pscustomobject]@{ Label = 'Domains'; Value = Format-SmartM365AdSummaryNumber -Value $summarySnapshot.DomainCount }
            [pscustomobject]@{ Label = 'Users'; Value = Format-SmartM365AdSummaryNumber -Value $summarySnapshot.TotalUsers }
            [pscustomobject]@{ Label = 'Computers'; Value = Format-SmartM365AdSummaryNumber -Value $summarySnapshot.TotalComputers }
            [pscustomobject]@{ Label = 'Groups'; Value = Format-SmartM365AdSummaryNumber -Value $summarySnapshot.TotalGroups }
            [pscustomobject]@{ Label = 'OUs'; Value = Format-SmartM365AdSummaryNumber -Value $summarySnapshot.TotalOUs }
            [pscustomobject]@{ Label = 'Contacts'; Value = Format-SmartM365AdSummaryNumber -Value $summarySnapshot.TotalContacts }
            [pscustomobject]@{ Label = 'Duplicate UPN accounts'; Value = Format-SmartM365AdSummaryNumber -Value $summarySnapshot.AffectedDuplicateUPNAccounts }
            [pscustomobject]@{ Label = 'Duplicate SMTP entries'; Value = Format-SmartM365AdSummaryNumber -Value $summarySnapshot.AffectedDuplicateSMTPEntries }
        )

        $pathRows = @(
            [pscustomobject]@{ Label = 'Source folder'; Path = $SourceFolder }
            [pscustomobject]@{ Label = 'Daily summary'; Path = $summaryCsvPath }
        )
        if (-not [string]::IsNullOrWhiteSpace($latestSummaryPath)) {
            $pathRows += [pscustomobject]@{ Label = 'Latest summary'; Path = $latestSummaryPath }
        }

        $emailBody = New-SmartM365EmailBody `
            -Title 'Active Directory daily summary' `
            -Category 'SmartM365 Active Directory' `
            -Severity $severity `
            -Tenant $Tenant `
            -HostName $env:COMPUTERNAME `
            -Message 'Daily Active Directory inventory summary with comparison against the latest available scan from a previous day.' `
            -ActionTitle $actionTitle `
            -ActionHtml $actionHtml `
            -SummaryRows $summaryRows `
            -PathRows $pathRows `
            -Sections $sections `
            -Footer 'This automated message was generated by SmartM365. Use the exported CSV paths and SharePoint links as the inventory source of truth.'

        Send-SmartM365AdInventoryEmailHtmlReport -Subject 'SmartM365 Active Directory daily summary' -BodyHtml $emailBody -To $dailySummaryTo

        $lastSentFolder = Split-Path -Path $LastSentFilePath -Parent
        if (-not [string]::IsNullOrWhiteSpace($lastSentFolder) -and -not (Test-Path -LiteralPath $lastSentFolder)) {
            New-Item -Path $lastSentFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        Set-Content -LiteralPath $LastSentFilePath -Value $todayStamp -Encoding UTF8
        WriteLog -Message ("Active Directory daily summary email sent. Last-sent marker updated: {0}" -f $LastSentFilePath)
        return $true
    }
    function Get-ADStringValue {
        param([object]$Value)
        if ($null -eq $Value) { return $null }
        if ($Value -is [string]) { return $Value }
        if ($Value.Count -eq 0) { return $null }
        return [string]($Value[0])
    }


    function Get-CountryNameFromCode {
        param(
            [Parameter(Mandatory = $true)]
            [string]$CountryCode
        )

        $CountryLookup = @{
            "AE" = "United Arab Emirates"
            "AT" = "Austria"
            "BE" = "Belgium"
            "BR" = "Brazil"
            "CH" = "Switzerland"
            "CL" = "Chile"
            "CN" = "China"
            "CZ" = "Czech Republic"
            "DE" = "Germany"
            "ES" = "Spain"
            "FR" = "France"
            "GB" = "United Kingdom"
            "HR" = "Croatia"
            "IE" = "Ireland"
            "IT" = "Italy"
            "LU" = "Luxembourg"
            "LV" = "Latvia"
            "MX" = "Mexico"
            "NL" = "Netherlands"
            "PL" = "Poland"
            "PT" = "Portugal"
            "SI" = "Slovenia"
            "UY" = "Uruguay"
        }

        if ($CountryLookup.ContainsKey($CountryCode.ToUpper())) {
            return $CountryLookup[$CountryCode.ToUpper()]
        }
        else {
            return "Country code not found"
        }
    }

    function Test-UserAccountControlFlag {
        param(
            [Parameter(Mandatory = $true)]
            [int]$UserAccountControlValue,

            [Parameter(Mandatory = $true)]
            [int]$FlagToCheck
        )

        return (($UserAccountControlValue -band $FlagToCheck) -eq $FlagToCheck)
    }


    function Test-GroupMembershipByName {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string[]]$GroupNames,

            [Parameter(Mandatory = $true)]
            [string]$GroupNameToFind
        )

        if ($null -eq $GroupNames -or $GroupNames.Count -eq 0) {
            return $false
        }

        return ($GroupNames -contains $GroupNameToFind)
    }


    function Get-ComputerGroupNames {
        param(
            [Parameter(Mandatory = $true)]
            [object]$Computer,

            [Parameter(Mandatory = $true)]
            [string]$Server,

            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainSid,

            [Parameter(Mandatory = $true)]
            [bool]$ResolveNestedGroups,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupNameByDNCache,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupParentsByDNCache,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupNameBySIDCache
        )

        if ($ResolveNestedGroups) {
            try {
                return @(Get-ADPrincipalGroupMembership -Identity $Computer -Server $Server | Select-Object -ExpandProperty Name)
            }
            catch {
                return @()
            }
        }

        $startDns = New-Object System.Collections.Generic.List[string]
        if ($Computer.MemberOf) {
            foreach ($memberDn in $Computer.MemberOf) {
                if (-not [string]::IsNullOrWhiteSpace([string]$memberDn)) {
                    [void]$startDns.Add([string]$memberDn)
                }
            }
        }

        if ($Computer.primaryGroupID -and $DomainSid) {
            try {
                $pgSid = ('{0}-{1}' -f $DomainSid, $Computer.primaryGroupID)
                $pgDn  = $null

                if ($GroupNameBySIDCache.ContainsKey($pgSid)) {
                    $pgDn = $GroupNameBySIDCache[$pgSid]
                }
                else {
                    $pgObj = Get-ADGroup -Identity $pgSid -Server $Server -Properties DistinguishedName
                    if ($pgObj -and $pgObj.DistinguishedName) {
                        $pgDn = [string]$pgObj.DistinguishedName
                        $GroupNameBySIDCache[$pgSid] = $pgDn
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($pgDn)) {
                    [void]$startDns.Add($pgDn)
                }
            }
            catch { }
        }

        if ($startDns.Count -eq 0) {
            return @()
        }

        $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $queue   = New-Object System.Collections.Queue
        $names   = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($dn in $startDns) {
            if ($visited.Add($dn)) {
                $queue.Enqueue($dn)
            }
        }

        while ($queue.Count -gt 0) {
            $gdn      = [string]$queue.Dequeue()
            $gName    = $null
            $gParents = $null

            if ($GroupNameByDNCache.ContainsKey($gdn)) {
                $gName = $GroupNameByDNCache[$gdn]
                if ($GroupParentsByDNCache.ContainsKey($gdn)) {
                    $gParents = $GroupParentsByDNCache[$gdn]
                }
            }
            else {
                try {
                    $groupServer = Get-AdServerForDistinguishedName -DistinguishedName $gdn -FallbackServer $Server
                    $gObj = Get-ADGroup -Identity $gdn -Server $groupServer -Properties Name, MemberOf -ErrorAction Stop
                    if ($gObj) {
                        $gName    = $gObj.Name
                        $gParents = @($gObj.MemberOf)
                        $GroupNameByDNCache[$gdn]    = $gName
                        $GroupParentsByDNCache[$gdn] = $gParents
                    }
                }
                catch {
                    $GroupNameByDNCache[$gdn] = ''
                    $GroupParentsByDNCache[$gdn] = @()
                    $gParents = @()
                    WriteLog -Message ("Computer group resolution failed once and was cached. Group='{0}'; Server='{1}'; Error='{2}'" -f $gdn, $groupServer, $_.Exception.Message) -Level 'WARNING'
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($gName)) {
                [void]$names.Add($gName)
            }

            if ($gParents) {
                foreach ($parentDn in $gParents) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$parentDn)) {
                        $parentDnString = [string]$parentDn
                        if ($visited.Add($parentDnString)) {
                            $queue.Enqueue($parentDnString)
                        }
                    }
                }
            }
        }

        return @($names)
    }


    # ----------------------------------------------------------
    # PATH VERIFICATION
    # ----------------------------------------------------------
    $destinationRootPath = $null
    try {
        $destinationRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''

        if ([string]::IsNullOrWhiteSpace($destinationRootPath)) {
            WriteLog -Message "WARNING: LatestCsvFolderPath not found in local configuration or returned empty. Combined CSV copy will be skipped."
            $destinationRootPath = $null
        }
        else {
            $destinationRootPath = $destinationRootPath.Trim()
            WriteLog -Message ("LatestCsvFolderPath resolved to: {0}" -f $destinationRootPath)
        }
    }
    catch {
        WriteLog -Message ("WARNING: Failed to resolve LatestCsvFolderPath: {0}. Combined CSV copy will be skipped." -f $_)
        $destinationRootPath = $null
    }

    if ($ReportOnly) {
        $reportSourceFolder = $destinationRootPath
        if ([string]::IsNullOrWhiteSpace($reportSourceFolder)) {
            $reportSourceFolder = $OutputPath
        }

        Invoke-ActiveDirectoryDailyReport -SourceFolder $reportSourceFolder `
            -ReportOutputPath $OutputPath `
            -LatestFolderPath $destinationRootPath `
            -AllowedOS $DailyReportAllowedOS `
            -TargetDomains $TargetDomains `
            -InactiveDays $DailyReportInactiveDays `
            -UseDailyLock $EnableDailyReportLock `
            -LockRoot $DailyReportLockRoot `
            -LockName $DailyReportLockName | Out-Null

        try {
            Invoke-SmartM365AdDailySummaryEmail -SourceFolder $reportSourceFolder `
                -SummaryOutputPath $OutputPath `
                -LatestFolderPath $destinationRootPath `
                -LastSentFilePath $DailySummaryLastSentFilePath `
                -ForceSend ([bool]$ForceSendDailySummary) | Out-Null
        }
        catch {
            WriteLog -Message ("Active Directory daily summary email failed in ReportOnly mode: {0}" -f $_) -Level 'WARNING'
        }

        Remove-OldFiles -Path $OutputPath -Filter "*.csv" -OlderThanDays 30
        Remove-OldFiles -Path $OutputPath -Filter "*.log" -OlderThanDays 30
        WriteLog -Message ("{0} completed successfully in ReportOnly mode." -f $TaskName)
        return
    }

    $combinedUsersCsv     = Join-Path $OutputPath "AD_Users_AllDomains_Brut.csv"
    $combinedComputersCsv = Join-Path $OutputPath "AD_Computers_AllDomains_Brut.csv"
    $combinedGroupsCsv    = Join-Path $OutputPath "AD_Groups_AllDomains.csv"
    $combinedOusCsv       = Join-Path $OutputPath "AD_OUs_AllDomains.csv"
    $combinedContactsCsv  = Join-Path $OutputPath "AD_Contacts_AllDomains.csv"
    $duplicateUpnCsv      = Join-Path $OutputPath "AD_Users_DuplicateUPN.csv"
    $duplicateSmtpCsv     = Join-Path $OutputPath "AD_Users_DuplicateSMTP.csv"
    $duplicateRemoteRoutingCsv = Join-Path $OutputPath "AD_Users_DuplicateRemoteRoutingAddress.csv"
    $remoteRoutingIssuesCsv    = Join-Path $OutputPath "AD_Users_RemoteRoutingIssues.csv"

    if (-not $DuplicateAnalysisOnly) {

    if ($TargetDomains -and $TargetDomains.Count -gt 0) {
        $DomainsToProcess = $TargetDomains
        WriteLog -Message ("Using explicitly provided target domains: {0}" -f ($DomainsToProcess -join ', '))
    }
    else {
        try {
            $forest = Get-ADForest -ErrorAction Stop
            $DomainsToProcess = $forest.Domains
            WriteLog -Message ("Discovered forest domains: {0}" -f ($DomainsToProcess -join ', '))
        }
        catch {
            throw "Unable to retrieve forest domains. $_"
        }
    }

    $skipDomainLoop = $false
    $DomainsToProcess = @($DomainsToProcess)

    if (-not $DomainWorker -and $EffectiveDomainParallelThrottleLimit -gt 1 -and $DomainsToProcess.Count -gt 1) {
        WriteLog -Message ("Starting parallel domain inventory with throttle limit {0}. Domains: {1}" -f $EffectiveDomainParallelThrottleLimit, ($DomainsToProcess -join ', '))

        $pwshPath = (Get-Process -Id $PID).Path
        $scriptPath = $PSCommandPath
        $workerJobs = @()
        $jobFailures = New-Object System.Collections.Generic.List[string]

        $receiveDomainJob = {
            param([System.Management.Automation.Job]$Job)

            $jobOutput = Receive-Job -Job $Job -ErrorAction SilentlyContinue 2>&1
            if ($jobOutput) {
                foreach ($line in $jobOutput) {
                    WriteLog -Message ("[{0}] {1}" -f $Job.Name, $line)
                }
            }

            if ($Job.State -ne 'Completed') {
                $reason = if ($Job.ChildJobs.Count -gt 0 -and $Job.ChildJobs[0].JobStateInfo.Reason) { $Job.ChildJobs[0].JobStateInfo.Reason.Message } else { $Job.State }
                [void]$jobFailures.Add(("{0}: {1}" -f $Job.Name, $reason))
            }

            Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
        }

        foreach ($domainName in $DomainsToProcess) {
            while (@($workerJobs | Where-Object { $_.State -eq 'Running' }).Count -ge $EffectiveDomainParallelThrottleLimit) {
                $completedJob = Wait-Job -Job $workerJobs -Any -Timeout 5
                if ($completedJob) {
                    & $receiveDomainJob $completedJob
                    $workerJobs = @($workerJobs | Where-Object { $_.Id -ne $completedJob.Id })
                }
            }

            $safeJobName = $domainName -replace '[^a-zA-Z0-9.-]', '_'
            $workerJobs += Start-Job -Name ("ADInventory-{0}" -f $safeJobName) -ScriptBlock {
                param($PwshPath, $ScriptPath, $TenantName, $DomainName, $OutputPathValue, $TempFolderPath)

                & $PwshPath -NoProfile -ExecutionPolicy Bypass -File $ScriptPath -Tenant $TenantName -TargetDomains $DomainName -OutputPath $OutputPathValue -DomainWorker -DomainWorkerTempFolder $TempFolderPath
                if ($LASTEXITCODE -ne 0) {
                    throw ("Domain worker failed for {0} with exit code {1}." -f $DomainName, $LASTEXITCODE)
                }
            } -ArgumentList $pwshPath, $scriptPath, $Tenant, $domainName, $OutputPath, $tempFolder
        }

        while ($workerJobs.Count -gt 0) {
            $completedJob = Wait-Job -Job $workerJobs -Any
            & $receiveDomainJob $completedJob
            $workerJobs = @($workerJobs | Where-Object { $_.Id -ne $completedJob.Id })
        }

        if ($jobFailures.Count -gt 0) {
            throw ("Parallel domain inventory failed: {0}" -f ($jobFailures -join ' | '))
        }

        WriteLog -Message "Parallel domain inventory workers completed successfully."
        $skipDomainLoop = $true
    }

    if (-not $skipDomainLoop) {
    foreach ($currentDomainName in $DomainsToProcess) {
        WriteLog -Message ("Starting inventory for domain '{0}'" -f $currentDomainName)

        $domainAttempt = 0
        $domainSuccess = $false

        while (-not $domainSuccess -and $domainAttempt -lt $MaxRetries) {
            $domainAttempt++
            if ($domainAttempt -gt 1) {
                WriteLog -Message ("Retrying inventory for domain '{0}' (attempt {1}/{2}) after {3}s delay..." -f $currentDomainName, $domainAttempt, $MaxRetries, $RetryDelaySeconds)
                Start-Sleep -Seconds $RetryDelaySeconds
            }

        try {

        $safeDomainFileName = $currentDomainName -replace '[^a-zA-Z0-9\.-]', '_'

        # ------------------------------------------------------
        # OU INVENTORY
        # ------------------------------------------------------
        if ($EnableOuInventory) {
        try {
            $CurrentObjectType = "OUs"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_OUs_{0}.csv" -f $safeDomainFileName)
            [int64]$ouCount = 0

            Get-ADOrganizationalUnit -Filter * -Server $currentDomainName -Properties Name, DistinguishedName, description, managedBy |
                ForEach-Object { [void]($ouCount++); $_ } |
                Select-Object `
                    @{Name = 'DomainName';   Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';   Expression = { $CurrentObjectType }},
                    Name,
                    DistinguishedName,
                    @{Name = 'Description'; Expression = { $_.Description -replace "`r", " -R " -replace "`n", " -N " }},
                    managedBy |
                Add-SmartM365TenantKey | Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported OUs for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $ouCount)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("OU inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("OU inventory skipped for domain '{0}' because EnableOuInventory is disabled." -f $currentDomainName)
        }

        # ------------------------------------------------------
        # COMPUTER INVENTORY
        # ------------------------------------------------------
        if ($EnableComputerInventory) {
        try {
            $CurrentObjectType = "Computers"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Computers_{0}.csv" -f $safeDomainFileName)

            $EnablePCFilter = $true
            if ($EnablePCFilter) {
                $computerFilter = {
                    (OperatingSystem -like "*Windows*" -and OperatingSystem -notlike "*Server*") -or (OperatingSystem -notlike "*")
                }
            }
            else {
                $computerFilter = { $true }
            }

            $GroupNameByDNCache     = @{}
            $GroupParentsByDNCache = @{}
            $GroupNameBySIDCache   = @{}

            try {
                $domainObj = Get-ADDomain -Server $currentDomainName
                $domainSid = $domainObj.DomainSID.Value
            }
            catch {
                $domainSid = $null
            }

            $ResolveNestedComputerGroups = $false
            [int64]$computerCount = 0

            Get-ADComputer -Filter $computerFilter -Server $currentDomainName -Properties SamAccountName, Name, DistinguishedName, Enabled, DNSHostName, OperatingSystem, operatingSystemHotfix, operatingSystemServicePack, operatingSystemVersion, LastLogonDate, LastLogonTimestamp, Description, IPv4Address, WhenCreated, WhenChanged, pwdLastSet, CanonicalName, MemberOf, primaryGroupID, ObjectGUID, ObjectSID, SIDHistory, extensionAttribute1, extensionAttribute2, extensionAttribute3, extensionAttribute4, extensionAttribute5, extensionAttribute6, extensionAttribute7, extensionAttribute8, extensionAttribute9, extensionAttribute10, extensionAttribute11, extensionAttribute12, extensionAttribute13, extensionAttribute14, extensionAttribute15 |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.DNSHostName) } |
                ForEach-Object {
                    [void]($computerCount++)
                    $computer = $_
                    $computerGroupNames = Get-ComputerGroupNames -Computer $computer -Server $currentDomainName -DomainSid $domainSid -ResolveNestedGroups:$ResolveNestedComputerGroups -GroupNameByDNCache $GroupNameByDNCache -GroupParentsByDNCache $GroupParentsByDNCache -GroupNameBySIDCache $GroupNameBySIDCache
                    $computerGroupNameSet = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
                    foreach ($computerGroupName in @($computerGroupNames)) {
                        if (-not [string]::IsNullOrWhiteSpace([string]$computerGroupName)) {
                            [void]$computerGroupNameSet.Add([string]$computerGroupName)
                        }
                    }
                    $configuredGroupFlags = New-Object System.Collections.Generic.List[bool]
                    $configuredGroupMatches = New-Object System.Collections.Generic.List[string]
                    foreach ($configuredGroupName in $ConfiguredComputerGroupNames) {
                        $isConfiguredGroupMember = (-not [string]::IsNullOrWhiteSpace($configuredGroupName)) -and $computerGroupNameSet.Contains($configuredGroupName)
                        [void]$configuredGroupFlags.Add([bool]$isConfiguredGroupMember)
                        if ($isConfiguredGroupMember) { [void]$configuredGroupMatches.Add($configuredGroupName) }
                    }

                    $computerRow = [ordered]@{
                        DomainName              = $currentDomainName
                        ObjectType              = $CurrentObjectType
                        SamAccountName          = $computer.SamAccountName
                        Name                    = $computer.Name
                        DistinguishedName       = $computer.DistinguishedName
                        Enabled                 = $computer.Enabled
                        DNSHostName             = $computer.DNSHostName
                        OperatingSystem         = $computer.OperatingSystem
                        operatingSystemHotfix   = $computer.operatingSystemHotfix
                        operatingSystemServicePack = $computer.operatingSystemServicePack
                        operatingSystemVersion  = $computer.operatingSystemVersion
                        LastLogonDate           = if ($computer.LastLogonTimestamp -ne $null -and $computer.LastLogonTimestamp -ne 0) { [datetime]::FromFileTime($computer.LastLogonTimestamp) } else { '' }
                        Description             = $computer.Description -replace "`r", " -R " -replace "`n", " -N "
                        IPv4Address             = $computer.IPv4Address
                        WhenCreated             = $computer.WhenCreated
                        WhenChanged             = $computer.WhenChanged
                        pwdLastSetDate          = if ($computer.pwdLastSet -ne $null -and $computer.pwdLastSet -ne 0) { [datetime]::FromFileTime($computer.pwdLastSet) } else { '' }
                        CanonicalName           = $computer.CanonicalName -replace "`r", " -R " -replace "`n", " -N "
                        MemberOfDNs             = if ($computer.MemberOf) { ($computer.MemberOf) -join ';' } else { '' }
                        PrimaryGroupName        = if ($computer.primaryGroupID -and $domainSid) {
                            $pgSid = ('{0}-{1}' -f $domainSid, $computer.primaryGroupID)
                            if ($GroupNameBySIDCache.ContainsKey($pgSid)) {
                                $pgDn = $GroupNameBySIDCache[$pgSid]
                                if (-not [string]::IsNullOrWhiteSpace($pgDn)) {
                                    if ($GroupNameByDNCache.ContainsKey($pgDn)) {
                                        $GroupNameByDNCache[$pgDn]
                                    }
                                    else {
                                        try {
                                            $g = Get-ADGroup -Identity $pgSid -Server $currentDomainName
                                            $GroupNameByDNCache[$pgDn] = $g.Name
                                            $g.Name
                                        }
                                        catch { '' }
                                    }
                                }
                                else {
                                    ''
                                }
                            }
                            else {
                                try {
                                    $g = Get-ADGroup -Identity $pgSid -Server $currentDomainName -Properties DistinguishedName
                                    if ($g -and $g.DistinguishedName) {
                                        $GroupNameBySIDCache[$pgSid] = [string]$g.DistinguishedName
                                        $GroupNameByDNCache[[string]$g.DistinguishedName] = $g.Name
                                    }
                                    $g.Name
                                }
                                catch { '' }
                            }
                        } else { '' }
                        ObjectGUID              = $computer.ObjectGUID
                        SID                     = $computer.ObjectSID.Value
                        SIDHistory              = if ($computer.SIDHistory) { ($computer.SIDHistory | ForEach-Object { $_.Value }) -join ';' } else { '' }
                        extensionAttribute1     = Get-ADStringValue $computer.extensionAttribute1
                        extensionAttribute2     = Get-ADStringValue $computer.extensionAttribute2
                        extensionAttribute3     = Get-ADStringValue $computer.extensionAttribute3
                        extensionAttribute4     = Get-ADStringValue $computer.extensionAttribute4
                        extensionAttribute5     = Get-ADStringValue $computer.extensionAttribute5
                        extensionAttribute6     = Get-ADStringValue $computer.extensionAttribute6
                        extensionAttribute7     = Get-ADStringValue $computer.extensionAttribute7
                        extensionAttribute8     = Get-ADStringValue $computer.extensionAttribute8
                        extensionAttribute9     = Get-ADStringValue $computer.extensionAttribute9
                        extensionAttribute10    = Get-ADStringValue $computer.extensionAttribute10
                        extensionAttribute11    = Get-ADStringValue $computer.extensionAttribute11
                        extensionAttribute12    = Get-ADStringValue $computer.extensionAttribute12
                        extensionAttribute13    = Get-ADStringValue $computer.extensionAttribute13
                        extensionAttribute14    = Get-ADStringValue $computer.extensionAttribute14
                        extensionAttribute15    = Get-ADStringValue $computer.extensionAttribute15
                        DomainNameShort         = Get-DomainNameShort -DomainName $currentDomainName
                        DomainAndSam            = Get-NormalizedDomainAndSam -DomainNameShort (Get-DomainNameShort -DomainName $currentDomainName) -SamAccountName $computer.SamAccountName
                        ImmutableId_AD          = Convert-GuidToImmutableId -ObjectGuid ([string]$computer.ObjectGUID)
                    }
                    for ($configuredGroupIndex = 0; $configuredGroupIndex -lt 10; $configuredGroupIndex++) {
                        $computerRow[("IsMemberOfConfiguredGroup{0:D2}" -f ($configuredGroupIndex + 1))] = [bool]$configuredGroupFlags[$configuredGroupIndex]
                    }
                    $computerRow['MatchedConfiguredGroups'] = ($configuredGroupMatches -join ';')
                    [PSCustomObject]$computerRow
                } |
                Add-SmartM365TenantKey | Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Computers for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $computerCount)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Computer inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("Computer inventory skipped for domain '{0}' because EnableComputerInventory is disabled." -f $currentDomainName)
        }

        # ------------------------------------------------------
        # USER INVENTORY
        # ------------------------------------------------------
        if ($EnableUserInventory) {
        try {
            $CurrentObjectType = "Users"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Users_{0}.csv" -f $safeDomainFileName)
            [int64]$userCount = 0

            [int]$DomainExcludedUsersNoUpn = 0
            try {
                $DomainExcludedUsersNoUpn = (Get-ADUser -LDAPFilter "(&(objectCategory=person)(objectClass=user)(!(userPrincipalName=*)))" -Server $currentDomainName -ResultSetSize $null).Count
            }
            catch {
                WriteLog -Message ("WARNING: Failed to count users without UPN for domain '{0}': {1}" -f $currentDomainName, $_)
                $DomainExcludedUsersNoUpn = 0
            }
            WriteLog -Message ("Users excluded because of missing UPN for domain '{0}': {1}" -f $currentDomainName, $DomainExcludedUsersNoUpn)

            Get-ADUser -LDAPFilter "(&(objectCategory=person)(objectClass=user)(userPrincipalName=*))" -Server $currentDomainName -ResultSetSize $null -Properties SamAccountName, sAMAccountType, Name, DistinguishedName, UserPrincipalName, Enabled, manager, LastLogonTimestamp, DisplayName, GivenName, Surname, Description, Department, Title, Company, Office, TelephoneNumber, MobilePhone, EmailAddress, StreetAddress, City, PostalCode, Country, WhenCreated, WhenChanged, AccountExpirationDate, pwdLastSet, badPwdCount, badPasswordTime, LogonCount, userAccountControl, msDS-ManagedPassword, ProxyAddresses, MemberOf, CanonicalName, ObjectGUID, targetAddress, msExchRemoteRecipientType, msExchRecipientTypeDetails, ObjectSID, SIDHistory, extensionAttribute1, extensionAttribute2, extensionAttribute3, extensionAttribute4, extensionAttribute5, extensionAttribute6, extensionAttribute7, extensionAttribute8, extensionAttribute9, extensionAttribute10, extensionAttribute11, extensionAttribute12, extensionAttribute13, extensionAttribute14, extensionAttribute15 |
                Select-Object `
                    @{Name = 'DomainName';           Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';           Expression = { $CurrentObjectType }},
                    SamAccountName,
                    sAMAccountType,
                    @{Name = 'Name';                Expression = { $_.Name        -replace "`r", " -R " -replace "`n", " -N " }},
                    DistinguishedName,
                    UserPrincipalName,
                    Enabled,
                    manager,
                    LastLogonTimestamp,
                    @{Name = 'LastLogonDate';       Expression = {
                        if ($_.LastLogonTimestamp -eq 0 -or $_.LastLogonTimestamp -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.LastLogonTimestamp)
                        }
                    }},
                    @{Name = 'DisplayName';          Expression = { $_.DisplayName  -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'GivenName';            Expression = { $_.GivenName    -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Surname';             Expression = { $_.Surname     -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Description';         Expression = { $_.Description -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Department';          Expression = { $_.Department  -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Title';               Expression = { $_.Title       -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Company';             Expression = { $_.Company     -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Office';              Expression = { $_.Office      -replace "`r", " -R " -replace "`n", " -N " }},
                    TelephoneNumber,
                    MobilePhone,
                    EmailAddress,
                    @{Name = 'StreetAddress';       Expression = { $_.StreetAddress -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'City';                Expression = { $_.City          -replace "`r", " -R " -replace "`n", " -N " }},
                    PostalCode,
                    Country,
                    @{Name = 'CountryName';         Expression = {
                        if (-not [string]::IsNullOrWhiteSpace($_.Country)) {
                            Get-CountryNameFromCode -CountryCode $_.Country
                        }
                        else {
                            "Unknown"
                        }
                    }},
                    WhenCreated,
                    WhenChanged,
                    AccountExpirationDate,
                    @{Name = 'PasswordLastSetDate'; Expression = {
                        if ($_.pwdLastSet -eq 0 -or $_.pwdLastSet -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.pwdLastSet)
                        }
                    }},
                    badPwdCount,
                    @{Name = 'BadPasswordDate';     Expression = {
                        if ($_.badPasswordTime -eq 0 -or $_.badPasswordTime -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.badPasswordTime)
                        }
                    }},
                    LogonCount,
                    userAccountControl,
                    @{Name = 'IsNormalAccount';         Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 64 }},
                    @{Name = 'IsScriptAccount';         Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 1 }},
                    @{Name = 'IsPasswordNeverExpires';  Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 65536 }},
                    @{Name = 'IsTrustedForDelegation';  Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 524288 }},
                    @{Name = 'MemberOfGroups';          Expression = { $_.MemberOf      -join ";" }},
                    @{Name = 'ProxyAddresses';          Expression = { $_.ProxyAddresses -join ";" }},
                    @{Name = 'CanonicalName';           Expression = { $_.CanonicalName -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'ObjectGUID';              Expression = { $_.ObjectGUID }},
                    @{Name = 'TargetAddress';           Expression = { $_.targetAddress }},
                    msExchRemoteRecipientType,
                    msExchRecipientTypeDetails,
                    @{Name = 'ObjectSID';               Expression = { $_.ObjectSID.Value }},
                    @{Name = 'ObjectSIDHistory';        Expression = {
                        if ($_.SIDHistory) {
                            ($_.SIDHistory | ForEach-Object { $_.Value }) -join ";"
                        }
                        else {
                            ""
                        }
                    }},
                    @{Name = 'extensionAttribute1';          Expression = { Get-ADStringValue $_.extensionAttribute1 }},
                    @{Name = 'extensionAttribute2';          Expression = { Get-ADStringValue $_.extensionAttribute2 }},
                    @{Name = 'extensionAttribute3';          Expression = { Get-ADStringValue $_.extensionAttribute3 }},
                    @{Name = 'extensionAttribute4';          Expression = { Get-ADStringValue $_.extensionAttribute4 }},
                    @{Name = 'extensionAttribute5';          Expression = { Get-ADStringValue $_.extensionAttribute5 }},
                    @{Name = 'extensionAttribute6';          Expression = { Get-ADStringValue $_.extensionAttribute6 }},
                    @{Name = 'extensionAttribute7';          Expression = { Get-ADStringValue $_.extensionAttribute7 }},
                    @{Name = 'extensionAttribute8';          Expression = { Get-ADStringValue $_.extensionAttribute8 }},
                    @{Name = 'extensionAttribute9';          Expression = { Get-ADStringValue $_.extensionAttribute9 }},
                    @{Name = 'extensionAttribute10';         Expression = { Get-ADStringValue $_.extensionAttribute10 }},
                    @{Name = 'extensionAttribute11';         Expression = { Get-ADStringValue $_.extensionAttribute11 }},
                    @{Name = 'extensionAttribute12';         Expression = { Get-ADStringValue $_.extensionAttribute12 }},
                    @{Name = 'extensionAttribute13';         Expression = { Get-ADStringValue $_.extensionAttribute13 }},
                    @{Name = 'extensionAttribute14';         Expression = { Get-ADStringValue $_.extensionAttribute14 }},
                    @{Name = 'extensionAttribute15';         Expression = { Get-ADStringValue $_.extensionAttribute15 }},
                    @{Name = 'MustChangePasswordAtNextLogon'; Expression = { ($_.pwdLastSet -eq 0) -and (-not (Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 65536)) }},
                    @{Name = 'DomainNameShort';         Expression = { Get-DomainNameShort -DomainName $currentDomainName }},
                    @{Name = 'DomainAndSam';            Expression = { Get-NormalizedDomainAndSam -DomainNameShort (Get-DomainNameShort -DomainName $currentDomainName) -SamAccountName $_.SamAccountName }},
                    @{Name = 'ImmutableId_AD';          Expression = { Convert-GuidToImmutableId -ObjectGuid ([string]$_.ObjectGUID) }} |
                ForEach-Object { [void]($userCount++); $_ } |
                Add-SmartM365TenantKey | Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Users for domain '{0}' to '{1}'. Count: {2}. Excluded without UPN: {3}" -f $currentDomainName, $outputCsvFilePath, $userCount, $DomainExcludedUsersNoUpn)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("User inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("User inventory skipped for domain '{0}' because EnableUserInventory is disabled." -f $currentDomainName)
        }

        # ------------------------------------------------------
        # GROUP INVENTORY
        # ------------------------------------------------------
        if ($EnableGroupInventory) {
        try {
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Groups_{0}.csv" -f $safeDomainFileName)
            $GroupData = @(Get-ADGroup -Filter * -Server $currentDomainName -Properties CanonicalName, CN, Created, createTimeStamp, Deleted, Description, DisplayName, DistinguishedName, GroupCategory, GroupScope, GroupType, HomePage, LastKnownParent, mail, ManagedBy, MemberOf, Members, Modified, modifyTimeStamp, Name, ObjectCategory, ObjectClass, ObjectGUID, objectSid, ProtectedFromAccidentalDeletion, SamAccountName, SIDHistory, whenChanged, whenCreated |
                Select-Object `
                    @{Name = 'DomainName'; Expression = { [string]$currentDomainName }},
                    @{Name = 'CanonicalName'; Expression = { Get-ADStringValue $_.CanonicalName }},
                    @{Name = 'CN'; Expression = { Get-ADStringValue $_.CN }},
                    @{Name = 'Created'; Expression = { $_.Created }},
                    @{Name = 'createTimeStamp'; Expression = { $_.createTimeStamp }},
                    @{Name = 'Deleted'; Expression = { $_.Deleted }},
                    @{Name = 'Description'; Expression = { Get-ADStringValue $_.Description }},
                    @{Name = 'DisplayName'; Expression = { Get-ADStringValue $_.DisplayName }},
                    @{Name = 'DistinguishedName'; Expression = { Get-ADStringValue $_.DistinguishedName }},
                    @{Name = 'GroupCategory'; Expression = { Get-ADStringValue $_.GroupCategory }},
                    @{Name = 'GroupScope'; Expression = { Get-ADStringValue $_.GroupScope }},
                    @{Name = 'GroupType'; Expression = { $_.GroupType }},
                    @{Name = 'HomePage'; Expression = { Get-ADStringValue $_.HomePage }},
                    @{Name = 'LastKnownParent'; Expression = { Get-ADStringValue $_.LastKnownParent }},
                    @{Name = 'mail'; Expression = { Get-ADStringValue $_.mail }},
                    @{Name = 'ManagedBy'; Expression = { Get-ADStringValue $_.ManagedBy }},
                    @{Name = 'MemberOf'; Expression = { if ($_.MemberOf) { ($_.MemberOf -join ';') } else { '' } }},
                    @{Name = 'Members'; Expression = { if ($_.Members) { ($_.Members -join ';') } else { '' } }},
                    @{Name = 'Modified'; Expression = { $_.Modified }},
                    @{Name = 'modifyTimeStamp'; Expression = { $_.modifyTimeStamp }},
                    @{Name = 'Name'; Expression = { Get-ADStringValue $_.Name }},
                    @{Name = 'ObjectCategory'; Expression = { Get-ADStringValue $_.ObjectCategory }},
                    @{Name = 'ObjectClass'; Expression = { Get-ADStringValue $_.ObjectClass }},
                    @{Name = 'ObjectGUID'; Expression = { if ($_.ObjectGUID) { $_.ObjectGUID.Guid } else { $null } }},
                    @{Name = 'objectSid'; Expression = { $_.objectSid.Value }},
                    @{Name = 'ProtectedFromAccidentalDeletion'; Expression = { $_.ProtectedFromAccidentalDeletion }},
                    @{Name = 'SamAccountName'; Expression = { Get-ADStringValue $_.SamAccountName }},
                    @{Name = 'SIDHistory'; Expression = { if ($_.SIDHistory) { ($_.SIDHistory.Value) -join ';' } else { '' } }},
                    @{Name = 'whenChanged'; Expression = { $_.whenChanged }},
                    @{Name = 'whenCreated'; Expression = { $_.whenCreated }}
            )
            $GroupData | Add-SmartM365TenantKey | Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8
            WriteLog -Message ("Exported Groups for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $GroupData.Count)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Group inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("Group inventory skipped for domain '{0}' because EnableGroupInventory is disabled." -f $currentDomainName)
        }

        # ------------------------------------------------------
        # CONTACT INVENTORY
        # ------------------------------------------------------
        if ($EnableContactInventory) {
        try {
            $CurrentObjectType = "Contacts"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Contacts_{0}.csv" -f $safeDomainFileName)
            [int64]$contactCount = 0

            Get-ADObject -Filter { ObjectClass -eq "contact" } -Server $currentDomainName -Properties DisplayName, ProxyAddresses, Mail, Name, DistinguishedName, ObjectGUID |
                ForEach-Object { [void]($contactCount++); $_ } |
                Select-Object `
                    @{Name = 'DomainName';      Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';      Expression = { $CurrentObjectType }},
                    Name,
                    DistinguishedName,
                    @{Name = 'ObjectGUID';       Expression = { $_.ObjectGUID.Guid }},
                    @{Name = 'DisplayName';      Expression = { $value = Get-ADStringValue $_.DisplayName; if ([string]::IsNullOrWhiteSpace($value)) { $value = Get-ADStringValue $_.Name }; if ([string]::IsNullOrWhiteSpace($value)) { $value = Get-ADStringValue $_.DistinguishedName }; if ([string]::IsNullOrWhiteSpace($value) -and $_.ObjectGUID) { $value = $_.ObjectGUID.Guid }; $value }},
                    @{Name = 'ProxyAddresses'; Expression = { $_.ProxyAddresses -join ";" }},
                    Mail |
                Add-SmartM365TenantKey | Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Contacts for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $contactCount)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Contact inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("Contact inventory skipped for domain '{0}' because EnableContactInventory is disabled." -f $currentDomainName)
        }

        $domainSuccess = $true

        } # end outer domain try
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) {
                if ($domainAttempt -lt $MaxRetries) {
                    WriteLog -Message ("WARNING: Transient AD connectivity error for domain '{0}' (attempt {1}/{2}): {3}" -f $currentDomainName, $domainAttempt, $MaxRetries, $_.Exception.Message)
                    # Loop continues: next iteration will wait and retry
                }
                else {
                    WriteLog -Message ("ERROR: Transient AD connectivity error for domain '{0}' persisted after {1} attempt(s). Skipping domain. Error: {2}" -f $currentDomainName, $domainAttempt, $_.Exception.Message)
                }
            }
            else {
                WriteLog -Message ("FATAL: Unhandled error for domain '{0}' (attempt {1}/{2}): {3}. Continuing with next domain." -f $currentDomainName, $domainAttempt, $MaxRetries, $_)
                break
            }
        }

        } # end while retry loop

        if (-not $domainSuccess) {
            WriteLog -Message ("WARNING: Domain '{0}' could not be fully inventoried after {1} attempt(s). Skipping." -f $currentDomainName, $domainAttempt)
        }
    }
    }

    if ($DomainWorker) {
        WriteLog -Message "Domain worker completed; combined CSV generation is handled by the parent process."
        return
    }

    # ------------------------------------------------------
    # COMBINE PER-DOMAIN CSV FILES
    # ------------------------------------------------------
    $combinedUsersCsv     = Join-Path $OutputPath "AD_Users_AllDomains_Brut.csv"
    $combinedComputersCsv = Join-Path $OutputPath "AD_Computers_AllDomains_Brut.csv"
    $combinedGroupsCsv    = Join-Path $OutputPath "AD_Groups_AllDomains.csv"
    $combinedOusCsv       = Join-Path $OutputPath "AD_OUs_AllDomains.csv"
    $combinedContactsCsv  = Join-Path $OutputPath "AD_Contacts_AllDomains.csv"

    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Users_*.csv"     -DestinationFile $combinedUsersCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Computers_*.csv" -DestinationFile $combinedComputersCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Groups_*.csv"    -DestinationFile $combinedGroupsCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_OUs_*.csv"       -DestinationFile $combinedOusCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Contacts_*.csv"  -DestinationFile $combinedContactsCsv

    $combinedUsersEnrichedCsv = $null
    if (Get-Command Invoke-SmartM365AdUsersEnrichedCsv -ErrorAction SilentlyContinue) {
        $combinedUsersEnrichedCsv = Invoke-SmartM365AdUsersEnrichedCsv `
            -CombinedUsersCsv $combinedUsersCsv `
            -OutputFolder $OutputPath `
            -LatestFolderPath $destinationRootPath `
            -RemoteRoutingDomain $RemoteRoutingDomain
    }
    else {
        WriteLog -Message "WARNING: AD users enrichment function is unavailable. AD_Users_AllDomains.csv will not be generated."
    }
    $combinedComputersEnrichedCsv = $null
    if (Get-Command Invoke-SmartM365AdComputersEnrichedCsv -ErrorAction SilentlyContinue) {
        $combinedComputersEnrichedCsv = Invoke-SmartM365AdComputersEnrichedCsv `
            -CombinedComputersCsv $combinedComputersCsv `
            -OutputFolder $OutputPath `
            -LatestFolderPath $destinationRootPath `
            -WindowsUpdateAnchorPolicyId $AdEnrichmentWindowsUpdateAnchorPolicyId `
            -WindowsUpdate24H2PolicyId $AdEnrichmentWindowsUpdate24H2PolicyId `
            -WindowsUpdate25H2PolicyId $AdEnrichmentWindowsUpdate25H2PolicyId
    }
    else {
        WriteLog -Message "WARNING: AD computers enrichment function is unavailable. AD_Computers_AllDomains.csv will not be generated."
    }

    # ------------------------------------------------------
    # Copy combined CSVs to the latest CSV folder
    # ------------------------------------------------------
    try {
        if ($destinationRootPath) {
            if (-not (Test-Path -LiteralPath $destinationRootPath)) {
                New-Item -Path $destinationRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                WriteLog -Message ("Created missing LatestCsvFolderPath directory: {0}" -f $destinationRootPath)
            }
            foreach ($combinedCsv in (@($combinedUsersCsv, $combinedUsersEnrichedCsv, $combinedComputersCsv, $combinedComputersEnrichedCsv, $combinedGroupsCsv, $combinedOusCsv, $combinedContactsCsv) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })) {
                if (Test-Path -Path $combinedCsv) {
                    $destinationFile = Join-Path $destinationRootPath ([System.IO.Path]::GetFileName($combinedCsv))
                    Copy-Item -LiteralPath $combinedCsv -Destination $destinationFile -Force -ErrorAction Stop
                    WriteLog -Message ("Copied combined CSV '{0}' to '{1}'" -f $combinedCsv, $destinationFile)
                    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $combinedCsv | Out-Null
                    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $destinationFile | Out-Null
                }
                else {
                    WriteLog -Message ("Combined CSV not found, skipping copy: {0}" -f $combinedCsv)
                }
            }
        }
        else {
            WriteLog -Message "LatestCsvFolderPath unavailable. Combined CSV copy skipped."
        }
    }
    catch {
        WriteLog -Message ("Copy to LatestCsvFolderPath failed: {0}" -f $_)
    }

    }
    else {
        WriteLog -Message "DuplicateAnalysisOnly mode: using existing combined users CSV for duplicate analysis."
        if ($destinationRootPath) {
            $latestUsersCsv = Join-Path $destinationRootPath "AD_Users_AllDomains.csv"
            if (-not (Test-Path -LiteralPath $latestUsersCsv)) { $latestUsersCsv = Join-Path $destinationRootPath "AD_Users_AllDomains_Brut.csv" }
            if (Test-Path -LiteralPath $latestUsersCsv) {
                $combinedUsersCsv = $latestUsersCsv
                WriteLog -Message ("DuplicateAnalysisOnly source CSV: {0}" -f $combinedUsersCsv)
            }
            else {
                WriteLog -Message ("Latest AD users CSV not found, falling back to OutputPath: {0}" -f $latestUsersCsv) -Level "WARNING"
            }

            $latestComputersCsv = Join-Path $destinationRootPath 'AD_Computers_AllDomains.csv'
            if (-not (Test-Path -LiteralPath $latestComputersCsv)) { $latestComputersCsv = Join-Path $destinationRootPath 'AD_Computers_AllDomains_Brut.csv' }
            if (Test-Path -LiteralPath $latestComputersCsv) { $combinedComputersCsv = $latestComputersCsv }

            foreach ($latestName in @('AD_Groups_AllDomains.csv', 'AD_OUs_AllDomains.csv', 'AD_Contacts_AllDomains.csv')) {
                $candidatePath = Join-Path $destinationRootPath $latestName
                if (-not (Test-Path -LiteralPath $candidatePath)) { continue }
                switch ($latestName) {
                    'AD_Groups_AllDomains.csv'    { $combinedGroupsCsv = $candidatePath }
                    'AD_OUs_AllDomains.csv'       { $combinedOusCsv = $candidatePath }
                    'AD_Contacts_AllDomains.csv'  { $combinedContactsCsv = $candidatePath }
                }
            }
        }
    }

    # ------------------------------------------------------
    # DUPLICATE USER IDENTITY ANALYSIS
    # ------------------------------------------------------
    if ($EnableDuplicateAnalysis) {
        try {
            WriteLog -Message "Starting duplicate UPN, SMTP proxy address, and remote routing analysis..."

            if (-not (Test-Path -Path $combinedUsersCsv)) {
                WriteLog -Message ("WARNING: Combined users CSV not found, skipping duplicate analysis: {0}" -f $combinedUsersCsv)
            }
            else {
                $allUsers = @(Import-Csv -Path $combinedUsersCsv -Encoding UTF8)
                WriteLog -Message ("Loaded {0} users from combined CSV for duplicate analysis." -f $allUsers.Count)

                $upnMap = @{}
                $smtpMap = @{}
                $remoteRoutingMap = @{}
                $remoteRoutingIssueRows = [System.Collections.Generic.List[object]]::new()
                $hasRemoteRecipientTypeColumn = ($allUsers.Count -eq 0) -or ($null -ne $allUsers[0].PSObject.Properties['msExchRemoteRecipientType'])
                if (-not $hasRemoteRecipientTypeColumn) {
                    WriteLog -Message 'The source users CSV does not contain msExchRemoteRecipientType. Missing targetAddress checks cannot run; remote routing checks will be limited to rows with a non-empty TargetAddress.' -Level 'WARNING'
                }

                $addRemoteRoutingIssue = {
                    param(
                        $User,
                        [string]$IssueType,
                        [string]$Severity,
                        [string]$TargetAddress,
                        [string]$NormalizedTargetAddress,
                        [bool]$HasMatchingProxyAddress,
                        [bool]$HasExpectedRoutingDomainProxyAddress
                    )

                    [void]$remoteRoutingIssueRows.Add([PSCustomObject][ordered]@{
                        IssueType                               = $IssueType
                        Severity                                = $Severity
                        TargetAddress                           = $TargetAddress
                        NormalizedTargetAddress                 = $NormalizedTargetAddress
                        ExpectedRemoteRoutingDomain             = $RemoteRoutingDomain
                        HasMatchingProxyAddress                 = $HasMatchingProxyAddress
                        HasExpectedRoutingDomainProxyAddress    = $HasExpectedRoutingDomainProxyAddress
                        msExchRemoteRecipientType               = $User.msExchRemoteRecipientType
                        msExchRecipientTypeDetails              = $User.msExchRecipientTypeDetails
                        UserPrincipalName                       = $User.UserPrincipalName
                        DomainName                              = $User.DomainName
                        DomainNameShort                         = $User.DomainNameShort
                        SamAccountName                          = $User.SamAccountName
                        DisplayName                             = $User.DisplayName
                        Enabled                                 = $User.Enabled
                        LastLogonDate                           = $User.LastLogonDate
                        DistinguishedName                       = $User.DistinguishedName
                    })
                }


                foreach ($u in $allUsers) {
                    $upnKey = ([string]$u.UserPrincipalName).Trim().ToLowerInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($upnKey)) {
                        if (-not $upnMap.ContainsKey($upnKey)) {
                            $upnMap[$upnKey] = [System.Collections.Generic.List[object]]::new()
                        }
                        [void]$upnMap[$upnKey].Add($u)
                    }

                    $proxyAddressSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    $hasExpectedRoutingDomainProxyAddress = $false
                    if (-not [string]::IsNullOrWhiteSpace([string]$u.ProxyAddresses)) {
                        foreach ($entry in ([string]$u.ProxyAddresses -split ';')) {
                            $entry = $entry.Trim()
                            if ($entry -notmatch '^(?i:smtp):(.+)$') { continue }

                            $smtpAddress = $Matches[1].Trim()
                            if ([string]::IsNullOrWhiteSpace($smtpAddress)) { continue }

                            $smtpKey = $smtpAddress.ToLowerInvariant()
                            [void]$proxyAddressSet.Add($smtpKey)
                            if ($smtpKey.EndsWith('@' + $RemoteRoutingDomain, [System.StringComparison]::OrdinalIgnoreCase)) {
                                $hasExpectedRoutingDomainProxyAddress = $true
                            }
                            if (-not $smtpMap.ContainsKey($smtpKey)) {
                                $smtpMap[$smtpKey] = [System.Collections.Generic.List[object]]::new()
                            }

                            [void]$smtpMap[$smtpKey].Add([PSCustomObject]@{
                                SmtpAddress       = $smtpAddress
                                IsUppercaseSMTP   = $entry.StartsWith('SMTP:', [System.StringComparison]::Ordinal)
                                UserPrincipalName = $u.UserPrincipalName
                                DomainName        = $u.DomainName
                                DomainNameShort   = $u.DomainNameShort
                                SamAccountName    = $u.SamAccountName
                                DisplayName       = $u.DisplayName
                                Enabled           = $u.Enabled
                                LastLogonDate     = $u.LastLogonDate
                                DistinguishedName = $u.DistinguishedName
                            })
                        }
                    }

                    $targetAddress = ([string]$u.TargetAddress).Trim()
                    $normalizedTargetAddress = $targetAddress
                    if ($normalizedTargetAddress -match '^(?i:smtp):(?<Address>.+)$') {
                        $normalizedTargetAddress = $Matches['Address'].Trim()
                    }
                    $normalizedTargetAddress = $normalizedTargetAddress.ToLowerInvariant()

                    [long]$remoteRecipientTypeValue = 0
                    $hasRemoteRecipientTypeValue = [long]::TryParse(([string]$u.msExchRemoteRecipientType).Trim(), [ref]$remoteRecipientTypeValue)
                    $isRemoteMailbox = ($hasRemoteRecipientTypeColumn -and $hasRemoteRecipientTypeValue -and $remoteRecipientTypeValue -ne 0) -or
                        (-not $hasRemoteRecipientTypeColumn -and -not [string]::IsNullOrWhiteSpace($normalizedTargetAddress))
                    if (-not $isRemoteMailbox) { continue }

                    if (-not [string]::IsNullOrWhiteSpace($normalizedTargetAddress)) {
                        if (-not $remoteRoutingMap.ContainsKey($normalizedTargetAddress)) {
                            $remoteRoutingMap[$normalizedTargetAddress] = [System.Collections.Generic.List[object]]::new()
                        }
                        [void]$remoteRoutingMap[$normalizedTargetAddress].Add($u)
                    }

                    $hasMatchingProxyAddress = -not [string]::IsNullOrWhiteSpace($normalizedTargetAddress) -and $proxyAddressSet.Contains($normalizedTargetAddress)
                    if ([string]::IsNullOrWhiteSpace($normalizedTargetAddress)) {
                        & $addRemoteRoutingIssue $u 'MissingRemoteRoutingAddress' 'Critical' $targetAddress $normalizedTargetAddress $false $hasExpectedRoutingDomainProxyAddress
                    }
                    else {
                        $isValidTargetAddress = $normalizedTargetAddress -match '^[^@\s]+@[^@\s]+$'
                        if (-not $isValidTargetAddress) {
                            & $addRemoteRoutingIssue $u 'InvalidRemoteRoutingAddress' 'Critical' $targetAddress $normalizedTargetAddress $hasMatchingProxyAddress $hasExpectedRoutingDomainProxyAddress
                        }
                        else {
                            $targetAddressDomain = $normalizedTargetAddress.Substring($normalizedTargetAddress.LastIndexOf('@') + 1)
                            if (-not $targetAddressDomain.Equals($RemoteRoutingDomain, [System.StringComparison]::OrdinalIgnoreCase)) {
                                & $addRemoteRoutingIssue $u 'UnexpectedRemoteRoutingDomain' 'Critical' $targetAddress $normalizedTargetAddress $hasMatchingProxyAddress $hasExpectedRoutingDomainProxyAddress
                            }
                            if (-not $hasMatchingProxyAddress) {
                                & $addRemoteRoutingIssue $u 'RemoteRoutingAddressMissingFromProxyAddresses' 'Warning' $targetAddress $normalizedTargetAddress $false $hasExpectedRoutingDomainProxyAddress
                            }
                        }
                    }

                    if (-not $hasExpectedRoutingDomainProxyAddress) {
                        & $addRemoteRoutingIssue $u 'MissingMailOnMicrosoftProxyAddress' 'Critical' $targetAddress $normalizedTargetAddress $hasMatchingProxyAddress $false
                    }
                }

                $duplicateUpnRows = @(
                    foreach ($upnKey in @($upnMap.Keys | Sort-Object)) {
                        $users = $upnMap[$upnKey]
                        if ($users.Count -le 1) { continue }

                        foreach ($u in @($users | Sort-Object DomainName, SamAccountName)) {
                            [PSCustomObject][ordered]@{
                                UserPrincipalName   = $u.UserPrincipalName
                                UPN_OccurrenceCount = $users.Count
                                DomainName          = $u.DomainName
                                DomainNameShort     = $u.DomainNameShort
                                SamAccountName      = $u.SamAccountName
                                DisplayName         = $u.DisplayName
                                Enabled             = $u.Enabled
                                LastLogonDate       = $u.LastLogonDate
                                DistinguishedName   = $u.DistinguishedName
                            }
                        }
                    }
                )

                $duplicateUpnColumns = @(
                    'UserPrincipalName',
                    'UPN_OccurrenceCount',
                    'DomainName',
                    'DomainNameShort',
                    'SamAccountName',
                    'DisplayName',
                    'Enabled',
                    'LastLogonDate',
                    'DistinguishedName'
                )
                Write-SmartM365CsvAtomically -Data $duplicateUpnRows -Path $duplicateUpnCsv -Columns $duplicateUpnColumns -Encoding UTF8
                Add-SmartM365AdGeneratedCsvPath -Path $duplicateUpnCsv
                $upnDuplicateCount = @($upnMap.Keys | Where-Object { $upnMap[$_].Count -gt 1 }).Count
                WriteLog -Message ("Duplicate UPN analysis complete. Distinct duplicate UPNs: {0}. Affected accounts: {1}. Output: {2}" -f $upnDuplicateCount, $duplicateUpnRows.Count, $duplicateUpnCsv)

                $duplicateSmtpRows = @(
                    foreach ($smtpKey in @($smtpMap.Keys | Sort-Object)) {
                        $entries = $smtpMap[$smtpKey]
                        if ($entries.Count -le 1) { continue }

                        foreach ($e in @($entries | Sort-Object DomainName, UserPrincipalName, SmtpAddress)) {
                            [PSCustomObject][ordered]@{
                                SmtpAddress          = $e.SmtpAddress
                                SMTP_OccurrenceCount = $entries.Count
                                IsUppercaseSMTP      = $e.IsUppercaseSMTP
                                UserPrincipalName    = $e.UserPrincipalName
                                DomainName           = $e.DomainName
                                DomainNameShort      = $e.DomainNameShort
                                SamAccountName       = $e.SamAccountName
                                DisplayName          = $e.DisplayName
                                Enabled              = $e.Enabled
                                LastLogonDate        = $e.LastLogonDate
                                DistinguishedName    = $e.DistinguishedName
                            }
                        }
                    }
                )

                $duplicateSmtpColumns = @(
                    'SmtpAddress',
                    'SMTP_OccurrenceCount',
                    'IsUppercaseSMTP',
                    'UserPrincipalName',
                    'DomainName',
                    'DomainNameShort',
                    'SamAccountName',
                    'DisplayName',
                    'Enabled',
                    'LastLogonDate',
                    'DistinguishedName'
                )
                Write-SmartM365CsvAtomically -Data $duplicateSmtpRows -Path $duplicateSmtpCsv -Columns $duplicateSmtpColumns -Encoding UTF8
                Add-SmartM365AdGeneratedCsvPath -Path $duplicateSmtpCsv
                $smtpDuplicateCount = @($smtpMap.Keys | Where-Object { $smtpMap[$_].Count -gt 1 }).Count
                WriteLog -Message ("Duplicate SMTP analysis complete. Distinct duplicate addresses: {0}. Affected entries: {1}. Output: {2}" -f $smtpDuplicateCount, $duplicateSmtpRows.Count, $duplicateSmtpCsv)

                $duplicateRemoteRoutingRows = @(
                    foreach ($remoteRoutingKey in @($remoteRoutingMap.Keys | Sort-Object)) {
                        $users = $remoteRoutingMap[$remoteRoutingKey]
                        if ($users.Count -le 1) { continue }

                        foreach ($u in @($users | Sort-Object DomainName, SamAccountName)) {
                            [PSCustomObject][ordered]@{
                                TargetAddress                         = $u.TargetAddress
                                NormalizedRemoteRoutingAddress        = $remoteRoutingKey
                                RemoteRoutingAddressOccurrenceCount   = $users.Count
                                ExpectedRemoteRoutingDomain           = $RemoteRoutingDomain
                                msExchRemoteRecipientType             = $u.msExchRemoteRecipientType
                                msExchRecipientTypeDetails            = $u.msExchRecipientTypeDetails
                                UserPrincipalName                     = $u.UserPrincipalName
                                DomainName                            = $u.DomainName
                                DomainNameShort                       = $u.DomainNameShort
                                SamAccountName                        = $u.SamAccountName
                                DisplayName                           = $u.DisplayName
                                Enabled                               = $u.Enabled
                                LastLogonDate                         = $u.LastLogonDate
                                DistinguishedName                     = $u.DistinguishedName
                            }
                        }
                    }
                )

                $duplicateRemoteRoutingColumns = @(
                    'TargetAddress',
                    'NormalizedRemoteRoutingAddress',
                    'RemoteRoutingAddressOccurrenceCount',
                    'ExpectedRemoteRoutingDomain',
                    'msExchRemoteRecipientType',
                    'msExchRecipientTypeDetails',
                    'UserPrincipalName',
                    'DomainName',
                    'DomainNameShort',
                    'SamAccountName',
                    'DisplayName',
                    'Enabled',
                    'LastLogonDate',
                    'DistinguishedName'
                )
                Write-SmartM365CsvAtomically -Data $duplicateRemoteRoutingRows -Path $duplicateRemoteRoutingCsv -Columns $duplicateRemoteRoutingColumns -Encoding UTF8
                Add-SmartM365AdGeneratedCsvPath -Path $duplicateRemoteRoutingCsv
                $remoteRoutingDuplicateCount = @($remoteRoutingMap.Keys | Where-Object { $remoteRoutingMap[$_].Count -gt 1 }).Count
                WriteLog -Message ("Duplicate remote routing address analysis complete. Distinct duplicate addresses: {0}. Affected accounts: {1}. Output: {2}" -f $remoteRoutingDuplicateCount, $duplicateRemoteRoutingRows.Count, $duplicateRemoteRoutingCsv)

                $remoteRoutingIssueRowsArray = @($remoteRoutingIssueRows.ToArray())
                $remoteRoutingIssueColumns = @(
                    'IssueType',
                    'Severity',
                    'TargetAddress',
                    'NormalizedTargetAddress',
                    'ExpectedRemoteRoutingDomain',
                    'HasMatchingProxyAddress',
                    'HasExpectedRoutingDomainProxyAddress',
                    'msExchRemoteRecipientType',
                    'msExchRecipientTypeDetails',
                    'UserPrincipalName',
                    'DomainName',
                    'DomainNameShort',
                    'SamAccountName',
                    'DisplayName',
                    'Enabled',
                    'LastLogonDate',
                    'DistinguishedName'
                )
                Write-SmartM365CsvAtomically -Data $remoteRoutingIssueRowsArray -Path $remoteRoutingIssuesCsv -Columns $remoteRoutingIssueColumns -Encoding UTF8
                Add-SmartM365AdGeneratedCsvPath -Path $remoteRoutingIssuesCsv
                $remoteRoutingIssueAccountCount = @($remoteRoutingIssueRowsArray | Select-Object -ExpandProperty DistinguishedName -Unique).Count
                $missingRemoteRoutingAddressCount = @($remoteRoutingIssueRowsArray | Where-Object IssueType -eq 'MissingRemoteRoutingAddress').Count
                $invalidRemoteRoutingAddressCount = @($remoteRoutingIssueRowsArray | Where-Object IssueType -eq 'InvalidRemoteRoutingAddress').Count
                $unexpectedRemoteRoutingDomainCount = @($remoteRoutingIssueRowsArray | Where-Object IssueType -eq 'UnexpectedRemoteRoutingDomain').Count
                $remoteRoutingAddressMissingFromProxyCount = @($remoteRoutingIssueRowsArray | Where-Object IssueType -eq 'RemoteRoutingAddressMissingFromProxyAddresses').Count
                $missingMailOnMicrosoftProxyCount = @($remoteRoutingIssueRowsArray | Where-Object IssueType -eq 'MissingMailOnMicrosoftProxyAddress').Count
                WriteLog -Message ("Remote routing validation complete. Issue rows: {0}. Affected accounts: {1}. Output: {2}" -f $remoteRoutingIssueRowsArray.Count, $remoteRoutingIssueAccountCount, $remoteRoutingIssuesCsv)

                $duplicateSharePointUploads = @()
                foreach ($duplicateCsv in @($duplicateUpnCsv, $duplicateSmtpCsv, $duplicateRemoteRoutingCsv, $remoteRoutingIssuesCsv)) {
                    if ($destinationRootPath -and (Test-Path -Path $duplicateCsv)) {
                        $duplicateLabel = switch ($duplicateCsv) {
                            $duplicateUpnCsv { 'Duplicate UPN'; break }
                            $duplicateSmtpCsv { 'Duplicate SMTP'; break }
                            $duplicateRemoteRoutingCsv { 'Duplicate remote routing address'; break }
                            $remoteRoutingIssuesCsv { 'Remote routing issues'; break }
                            default { [System.IO.Path]::GetFileNameWithoutExtension($duplicateCsv) }
                        }
                        $sourceUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $duplicateCsv
                        $sourceUpload = Add-SmartM365SharePointUploadLabel -UploadRecord $sourceUpload -Label ("{0} (DATA-ALL)" -f $duplicateLabel)
                        if ($sourceUpload) { $duplicateSharePointUploads += $sourceUpload }

                        $destinationFile = Join-Path $destinationRootPath ([System.IO.Path]::GetFileName($duplicateCsv))
                        Copy-Item -LiteralPath $duplicateCsv -Destination $destinationFile -Force -ErrorAction Stop
                        WriteLog -Message ("Copied '{0}' to '{1}'" -f $duplicateCsv, $destinationFile)
                        $latestUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $destinationFile
                        $latestUpload = Add-SmartM365SharePointUploadLabel -UploadRecord $latestUpload -Label ("{0} (DATA-LAST)" -f $duplicateLabel)
                        if ($latestUpload) { $duplicateSharePointUploads += $latestUpload }
                    }
                }

                $hasDuplicateIdentities = (($upnDuplicateCount -gt 0) -or ($smtpDuplicateCount -gt 0) -or ($remoteRoutingDuplicateCount -gt 0) -or ($remoteRoutingIssueRowsArray.Count -gt 0))
                if ($EnableDuplicateNotification -and $hasDuplicateIdentities) {
                    if ([string]::IsNullOrWhiteSpace($DuplicateNotificationLastSentFilePath)) {
                        $DuplicateNotificationLastSentFilePath = Join-Path $OutputPath 'AD_DuplicateNotification_LastSent.txt'
                    }

                    $todayStamp = (Get-Date).ToString('yyyy-MM-dd')
                    $lastSentStamp = ''
                    if (Test-Path -LiteralPath $DuplicateNotificationLastSentFilePath) {
                        $lastSentStamp = (Get-Content -LiteralPath $DuplicateNotificationLastSentFilePath -Raw -ErrorAction SilentlyContinue).Trim()
                    }

                    if ($lastSentStamp -eq $todayStamp -and -not $ForceSendDuplicateNotification) {
                        WriteLog -Message ("Duplicate identity notification already sent today ({0}). Use -ForceSendDuplicateNotification to resend." -f $todayStamp)
                    }
                    else {
                        $notificationFolder = Split-Path -Path $DuplicateNotificationLastSentFilePath -Parent
                        if (-not [string]::IsNullOrWhiteSpace($notificationFolder) -and -not (Test-Path -LiteralPath $notificationFolder)) {
                            New-Item -Path $notificationFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }

                        $emailSubject = "SmartM365 Active Directory identity and mail routing issues detected"
                        $duplicateSummaryRows = @(
                            [pscustomobject]@{ Label = 'Distinct duplicate UPNs'; Value = $upnDuplicateCount }
                            [pscustomobject]@{ Label = 'Affected UPN accounts'; Value = $duplicateUpnRows.Count }
                            [pscustomobject]@{ Label = 'Distinct duplicate SMTP addresses'; Value = $smtpDuplicateCount }
                            [pscustomobject]@{ Label = 'Affected SMTP entries'; Value = $duplicateSmtpRows.Count }
                            [pscustomobject]@{ Label = 'Distinct duplicate remote routing addresses'; Value = $remoteRoutingDuplicateCount }
                            [pscustomobject]@{ Label = 'Affected remote routing accounts'; Value = $duplicateRemoteRoutingRows.Count }
                            [pscustomobject]@{ Label = 'Remote routing issue accounts'; Value = $remoteRoutingIssueAccountCount }
                            [pscustomobject]@{ Label = 'Missing targetAddress'; Value = $missingRemoteRoutingAddressCount }
                            [pscustomobject]@{ Label = 'Invalid targetAddress'; Value = $invalidRemoteRoutingAddressCount }
                            [pscustomobject]@{ Label = 'Unexpected targetAddress domain'; Value = $unexpectedRemoteRoutingDomainCount }
                            [pscustomobject]@{ Label = 'targetAddress absent from proxyAddresses'; Value = $remoteRoutingAddressMissingFromProxyCount }
                            [pscustomobject]@{ Label = 'Missing tenant mail.onmicrosoft.com proxy'; Value = $missingMailOnMicrosoftProxyCount }
                        )
                        $diagnosticWorkbookPath = Export-SmartM365AdDiagnosticsWorkbook -Path (Join-Path $OutputPath 'AD_Users_IdentityAndMailRoutingIssues.xlsx') -Sources @(
                            [pscustomobject]@{ CsvPath = $duplicateUpnCsv; WorksheetName = 'Duplicate UPN'; TableName = 'DuplicateUPN' }
                            [pscustomobject]@{ CsvPath = $duplicateSmtpCsv; WorksheetName = 'Duplicate SMTP'; TableName = 'DuplicateSMTP' }
                            [pscustomobject]@{ CsvPath = $duplicateRemoteRoutingCsv; WorksheetName = 'Duplicate RemoteRouting'; TableName = 'DuplicateRemoteRouting' }
                            [pscustomobject]@{ CsvPath = $remoteRoutingIssuesCsv; WorksheetName = 'RemoteRouting Issues'; TableName = 'RemoteRoutingIssues' }
                        )
                        $duplicatePathRows = @(
                            [pscustomobject]@{ Label = 'Source users'; Path = $combinedUsersCsv }
                            [pscustomobject]@{ Label = 'Diagnostic workbook'; Path = $diagnosticWorkbookPath }
                        )
                        $duplicateUpnPreviewSection = New-SmartM365AdDuplicatePreviewSection -Rows $duplicateUpnRows -DuplicateType 'UPN' -Limit 50
                        $duplicateSmtpPreviewSection = New-SmartM365AdDuplicatePreviewSection -Rows $duplicateSmtpRows -DuplicateType 'SMTP' -Limit 50
                        $duplicateSharePointSection = New-SmartM365SharePointLinksSection -UploadRecords $duplicateSharePointUploads
                        $duplicateRemoteRoutingPreviewSection = New-SmartM365AdDuplicatePreviewSection -Rows $duplicateRemoteRoutingRows -DuplicateType 'REMOTE' -Limit 50
                        $remoteRoutingIssuePreviewSection = New-SmartM365AdRemoteRoutingIssuePreviewSection -Rows $remoteRoutingIssueRowsArray -Limit 50
                        $duplicateSections = @()
                        if ($duplicateUpnPreviewSection) { $duplicateSections += $duplicateUpnPreviewSection }
                        if ($duplicateSmtpPreviewSection) { $duplicateSections += $duplicateSmtpPreviewSection }
                        if ($duplicateSharePointSection) { $duplicateSections += $duplicateSharePointSection }

                        if ($duplicateRemoteRoutingPreviewSection) { $duplicateSections += $duplicateRemoteRoutingPreviewSection }
                        if ($remoteRoutingIssuePreviewSection) { $duplicateSections += $remoteRoutingIssuePreviewSection }
                        $emailBody = New-SmartM365EmailBody `
                            -Title 'Identity and mail routing issues detected' `
                            -Category 'SmartM365 Active Directory' `
                            -Severity Warning `
                            -Tenant $Tenant `
                            -HostName $env:COMPUTERNAME `
                            -Message ("Duplicate identity and remote routing analysis found conflicts in Active Directory user data. Expected remote routing domain: {0}." -f $RemoteRoutingDomain) `
                            -ActionTitle 'Action required' `
                            -ActionHtml 'Review the attached Excel workbook, which contains the duplicate UPN, SMTP, remote routing address, and remote routing issue details, before identity cleanup, migration, or synchronization decisions.' `
                            -SummaryRows $duplicateSummaryRows `
                            -PathRows $duplicatePathRows `
                            -Sections $duplicateSections `
                            -Footer 'This automated message was generated by SmartM365. Use the attached Excel workbook and SharePoint links above as the source of truth for remediation.'
                        $duplicateNotificationTo = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue '')
                        if ([string]::IsNullOrWhiteSpace($duplicateNotificationTo)) {
                            WriteLog -Message 'Duplicate identity notification skipped because To is not configured in local or global configuration. Duplicate CSV exports remain valid.' -Level 'WARNING'
                        }
                        else {
                            Send-SmartM365AdInventoryEmailHtmlReport -Subject $emailSubject -BodyHtml $emailBody -To $duplicateNotificationTo -Attachments @($diagnosticWorkbookPath)
                            Set-Content -LiteralPath $DuplicateNotificationLastSentFilePath -Value $todayStamp -Encoding UTF8
                            WriteLog -Message ("Duplicate identity notification sent. Last-sent marker updated: {0}" -f $DuplicateNotificationLastSentFilePath)
                        }
                    }
                }
                elseif (-not $EnableDuplicateNotification) {
                    WriteLog -Message "Duplicate identity notification skipped because EnableDuplicateNotification is disabled."
                }
            }
        }
        catch {
            $duplicateAnalysisError = $_
            $duplicateAnalysisErrorText = [string]$duplicateAnalysisError
            $isDuplicateNotificationError = $duplicateAnalysisErrorText -like 'Duplicate identity notification requires To in local or global configuration*'

            if ($isDuplicateNotificationError) {
                WriteLog -Message ("Duplicate identity notification failed: {0}" -f $duplicateAnalysisError) -Level 'ERROR'
            }
            else {
                WriteLog -Message ("Duplicate user identity analysis failed: {0}" -f $duplicateAnalysisError) -Level 'ERROR'
            }

            try {
                $emailSubject = if ($isDuplicateNotificationError) { "[ERROR] SmartM365 Active Directory duplicate identity notification" } else { "[ERROR] SmartM365 Active Directory duplicate identity analysis" }
                $errorTitle = if ($isDuplicateNotificationError) { 'Duplicate identity notification failed' } else { 'Duplicate identity analysis failed' }
                $actionTitle = if ($isDuplicateNotificationError) { 'Notification configuration error' } else { 'Error notification' }
                $actionHtml = if ($isDuplicateNotificationError) { 'The duplicate identity analysis completed, but the notification cannot be sent because To is not configured in local or global configuration. ErrorMailTo is reserved for error notifications.' } else { 'The duplicate identity analysis stopped before completion. Review the error details and the script log.' }
                $safeDuplicateAnalysisError = ConvertTo-SmartM365EmailHtmlText $duplicateAnalysisError
                $duplicateAnalysisErrorHtml = @"
          <tr>
            <td style="padding:18px 24px 0 24px;">
              <div style="font-size:15px;font-weight:700;color:#111827;margin-bottom:8px;">Error details</div>
              <pre style="margin:0;background:#f8fafc;border:1px solid #d9e2ec;border-radius:4px;padding:10px 12px;font-family:Consolas,'Courier New',monospace;font-size:12px;line-height:18px;color:#334155;white-space:pre-wrap;word-break:break-word;">$safeDuplicateAnalysisError</pre>
            </td>
          </tr>
"@
                $emailBody = New-SmartM365EmailBody `
                    -Title $errorTitle `
                    -Category 'SmartM365 Active Directory' `
                    -Severity Error `
                    -Tenant $Tenant `
                    -HostName $env:COMPUTERNAME `
                    -ActionTitle $actionTitle `
                    -ActionHtml $actionHtml `
                    -SummaryRows @(
                        [pscustomobject]@{ Label = 'Script'; Value = $MyInvocation.MyCommand.Name }
                        [pscustomobject]@{ Label = 'Version'; Value = $ScriptVersion }
                        [pscustomobject]@{ Label = 'Date'; Value = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz') }
                    ) `
                    -BodyHtml $duplicateAnalysisErrorHtml
                Send-SmartM365AdInventoryEmailHtmlReport -Subject $emailSubject -BodyHtml $emailBody
                if ($isDuplicateNotificationError) {
                    WriteLog -Message "Duplicate identity notification error email sent."
                }
                else {
                    WriteLog -Message "Duplicate identity analysis error email notification sent."
                }
            }
            catch {
                WriteLog -Message ("Failed to send duplicate identity analysis error email notification: {0}" -f $_)
            }
        }
    }
    else {
        WriteLog -Message "Duplicate UPN and SMTP proxy address analysis skipped because EnableDuplicateAnalysis is disabled."
    }

    # ------------------------------------------------------
    # DAILY ACTIVE DIRECTORY REPORTS
    # ------------------------------------------------------
    if (-not $DuplicateAnalysisOnly -and -not $SkipDailyReport) {
        Invoke-ActiveDirectoryDailyReport -SourceFolder $OutputPath `
            -ReportOutputPath $OutputPath `
            -LatestFolderPath $destinationRootPath `
            -AllowedOS $DailyReportAllowedOS `
            -TargetDomains $TargetDomains `
            -InactiveDays $DailyReportInactiveDays `
            -UseDailyLock $EnableDailyReportLock `
            -LockRoot $DailyReportLockRoot `
            -LockName $DailyReportLockName | Out-Null
    }
    elseif ($DuplicateAnalysisOnly) {
        WriteLog -Message "Daily Active Directory report skipped in DuplicateAnalysisOnly mode."
    }
    else {
        WriteLog -Message "Daily Active Directory report skipped because -SkipDailyReport was specified."
    }
    # ------------------------------------------------------
    # DAILY ACTIVE DIRECTORY SUMMARY EMAIL
    # ------------------------------------------------------
    if (-not $DuplicateAnalysisOnly -and -not $SkipDailyReport) {
        try {
            Invoke-SmartM365AdDailySummaryEmail -SourceFolder $OutputPath `
                -SummaryOutputPath $OutputPath `
                -LatestFolderPath $destinationRootPath `
                -LastSentFilePath $DailySummaryLastSentFilePath `
                -ForceSend ([bool]$ForceSendDailySummary) | Out-Null
        }
        catch {
            WriteLog -Message ("Active Directory daily summary email failed: {0}" -f $_) -Level 'WARNING'
        }
    }
    elseif ($DuplicateAnalysisOnly) {
        WriteLog -Message "Daily Active Directory summary email skipped in DuplicateAnalysisOnly mode."
    }
    else {
        WriteLog -Message "Daily Active Directory summary email skipped because -SkipDailyReport was specified."
    }
    # ------------------------------------------------------
    # WEEKLY INVENTORY HISTORY
    # ------------------------------------------------------
    if ($EnableWeeklyHistory) {
        try {
            $weeklySourceFiles = New-Object System.Collections.Generic.List[string]
            $weeklyRequiredSourceFiles = New-Object System.Collections.Generic.List[string]

            if ($EnableUserInventory) {
                $weeklyUsersCsv = if ([string]::IsNullOrWhiteSpace($combinedUsersEnrichedCsv)) { Join-Path $OutputPath 'AD_Users_AllDomains.csv' } else { $combinedUsersEnrichedCsv }
                [void]$weeklySourceFiles.Add($weeklyUsersCsv)
                [void]$weeklyRequiredSourceFiles.Add($weeklyUsersCsv)
            }
            if ($EnableComputerInventory) {
                $weeklyComputersCsv = if ([string]::IsNullOrWhiteSpace($combinedComputersEnrichedCsv)) { Join-Path $OutputPath 'AD_Computers_AllDomains.csv' } else { $combinedComputersEnrichedCsv }
                [void]$weeklySourceFiles.Add($weeklyComputersCsv)
                [void]$weeklyRequiredSourceFiles.Add($weeklyComputersCsv)
            }
            foreach ($weeklyInventory in @(
                [pscustomobject]@{ Enabled = $EnableGroupInventory; Path = $combinedGroupsCsv },
                [pscustomobject]@{ Enabled = $EnableOuInventory; Path = $combinedOusCsv },
                [pscustomobject]@{ Enabled = $EnableContactInventory; Path = $combinedContactsCsv }
            )) {
                if (-not $weeklyInventory.Enabled) { continue }
                [void]$weeklySourceFiles.Add($weeklyInventory.Path)
                [void]$weeklyRequiredSourceFiles.Add($weeklyInventory.Path)
            }
            if ($EnableDuplicateAnalysis -and $EnableUserInventory) {
                foreach ($weeklyDiagnosticCsv in @(
                    $duplicateUpnCsv,
                    $duplicateSmtpCsv,
                    $duplicateRemoteRoutingCsv,
                    $remoteRoutingIssuesCsv
                )) {
                    [void]$weeklySourceFiles.Add($weeklyDiagnosticCsv)
                    [void]$weeklyRequiredSourceFiles.Add($weeklyDiagnosticCsv)
                }
            }

            $weeklyHistoryParameters = @{
                SourceFiles         = $weeklySourceFiles.ToArray()
                RequiredSourceFiles = $weeklyRequiredSourceFiles.ToArray()
                HistoryRootPath     = $WeeklyHistoryFolderPath
                RetentionWeeks      = $WeeklyHistoryRetentionWeeks
            }
            Save-WeeklyInventoryHistory @weeklyHistoryParameters
        }
        catch {
            WriteLog -Message ("Weekly AD inventory history failed: {0}" -f $_) -Level "WARNING"
        }
    }
    else {
        WriteLog -Message "Weekly AD inventory history skipped because EnableWeeklyHistory is disabled."
    }

    # ------------------------------------------------------
    # CLEANUP TEMPORARY PER-DOMAIN CSV FILES
    # ------------------------------------------------------
    if (-not $DuplicateAnalysisOnly) {
        Remove-TemporaryInventoryFolder -TempFolder $tempFolder -BaseFolder $baseFolder
    }
    else {
        WriteLog -Message "Temporary per-domain cleanup skipped in DuplicateAnalysisOnly mode."
    }

    # ------------------------------------------------------
    # CLEANUP OLD FILES
    # ------------------------------------------------------
    Remove-OldFiles -Path $OutputPath -Filter "*.csv" -OlderThanDays 30
    Remove-OldFiles -Path $OutputPath -Filter "*.log" -OlderThanDays 30

    if ([int]$global:SmartM365ErrorCount -gt 0) {
        WriteLog -Message ("{0} processing completed, but {1} error(s) were recorded. See the execution summary and log for details." -f $TaskName, [int]$global:SmartM365ErrorCount) -Level 'WARNING'
    }
    else {
        WriteLog -Message ("{0} completed successfully." -f $TaskName)
    }
}
catch {
    $globalError = $_
    WriteLog -Message ("Fatal error in script: {0}" -f $globalError)

    try {
        $emailSubject = "[ERROR] $TaskName"
        $safeGlobalError = ConvertTo-SmartM365EmailHtmlText $globalError
        $globalErrorHtml = @"
          <tr>
            <td style="padding:18px 24px 0 24px;">
              <div style="font-size:15px;font-weight:700;color:#111827;margin-bottom:8px;">Error details</div>
              <pre style="margin:0;background:#f8fafc;border:1px solid #d9e2ec;border-radius:4px;padding:10px 12px;font-family:Consolas,'Courier New',monospace;font-size:12px;line-height:18px;color:#334155;white-space:pre-wrap;word-break:break-word;">$safeGlobalError</pre>
            </td>
          </tr>
"@
        $emailBody = New-SmartM365EmailBody `
            -Title 'Active Directory inventory failed' `
            -Category 'SmartM365 Active Directory' `
            -Severity Error `
            -Tenant $Tenant `
            -HostName $env:COMPUTERNAME `
            -ActionTitle 'Error notification' `
            -ActionHtml 'The Active Directory inventory stopped with a fatal error. Review the error details and the script log.' `
            -SummaryRows @(
                [pscustomobject]@{ Label = 'Script'; Value = $MyInvocation.MyCommand.Name }
                [pscustomobject]@{ Label = 'Version'; Value = $ScriptVersion }
                [pscustomobject]@{ Label = 'Date'; Value = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz') }
            ) `
            -BodyHtml $globalErrorHtml
        Send-SmartM365AdInventoryEmailHtmlReport -Subject $emailSubject -BodyHtml $emailBody
        WriteLog -Message "Global error email notification sent."
    }
    catch {
        WriteLog -Message ("Failed to send global error email notification: {0}" -f $_)
    }
}
finally {
    try {
        Clear-SmartM365AdReferenceXlsxCache
    }
    catch {
        WriteLog -Message ("AD reference workbook cache cleanup failed: {0}" -f $_.Exception.Message) -Level 'WARNING'
    }

    try {
        WriteLog -Message ("Stopping transcript for script '{0}'" -f $MyInvocation.MyCommand.Name)
        Stop-Transcript | Out-Null
        try {
            $smartM365TranscriptPath = $null
            $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue
            if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
                $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
            }
            else {
                $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue
                if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
                    $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
                }
            }
            if ($smartM365TranscriptPath) {
                Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath
            }
        }
        catch {
        }
    }
    catch {
        Write-Host ("Failed to stop transcript: {0}" -f $_) -ForegroundColor Yellow
    }

    try {
        Complete-SmartM365ExecutionContext -Status Auto -ErrorRecord $globalError
    }
    catch {
        Write-Host ("Failed to write execution summary: {0}" -f $_) -ForegroundColor Yellow
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQ6cBordhC9vQK
# HelEcrmjwXaD2KQa/qlg7UWWaqQhuKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILP+qSErr2h4f3yi0mssAnv20zYD3SlE8XRIGO5/DVx4MA0GCSqG
# SIb3DQEBAQUABIIBgAUpiqQ/W7aBsYsAA4LmIkC+Pxnd9JZAQb2xgVjZ24r2RwYh
# 4h6xJgbXI/LthwOK8Ufzf+MtgsQMEgqWtj4VOJSL3sA1+08UPQvilPieGeg1edmi
# Ytl6rQ3a1WiIAzOfC6sa/b16BF8rx3pMP/Ufo+w3UrPnExY1lF67/sVZLBes13+d
# Yv+6zs8jO5++nuyin6Hl/ZQUpo1nL3ERicgwB8FfPADfqTYKr2oZhXu7uwhehSXM
# y9UBTRMiMHeXhGsI3wU70/QYnBjEOEddAY+QOqfZW99vZ/mQVY8Di6EC1a/EfYO6
# eZGyVsP4qaYAel94yBq6kyr7aZ+OuApwZK5VBWLNRtln2MVPXzpHSW7Qkq4hWlXL
# wQKYKexuBchEElmL+gdjuI86KLmBXCxCch3TaI3Npn9KVhCUtoz0j5ScC22aV+dA
# egRHM/i8gx0QVAfxJxk3q4W/GVOzRPqPxuz/35WG3EMdv8RQXL5cAp76ev/iF1XS
# PNYRg3wSon0OLChgHaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTcxOTQw
# NDJaMC8GCSqGSIb3DQEJBDEiBCC5qOEoNShuY6ac1PHN1Ht97LXt9+JMfGne2M+p
# l9Kq2jANBgkqhkiG9w0BAQEFAASCAgC8tWIvXh3ocstXbE/+1hJuCQIXxo6NnAxD
# qQU5m8MAdE3f/1T1+44DNC0rymBYf6l9oK4Y78DPtQR+ETEcp+jFPzjrQTjYL2We
# oo3CLI0B3K5CJC2tBIkoq8BBimJdO6QTnzltElK9rkVBYz26bdl2EFdY0a7MWh+k
# EREcox/j6V6FsBg34XFGlLsUTFJUeDa0gaI5dgWk05JAvwnHa46Gka3x6+kk6/qH
# XJ4xpp5VcQ47In+R2x0mRnKWTeuamDs8zuJ+nuHpH98xEC+gTg9+wcVgBMYbZ3mU
# MoMlPT632tZf/up4YPjlpQxoLRJdJAJ0DfIKELNinAUWA2YKXTz7SvvLeCGDCDUs
# 92avU0Fhn1hY5LGZ4sLHLpFcw1A2Oq7z/cKuF8rTC+YcGBFXht3BwKVwmEn6jd+o
# TpKQetayZdXf6LCTFI/IyohJMzE4poEjqRbSpwRdEqxd3KOYi1s2SkYgE6/ithAm
# kWP5+zcQzRMZjw14UOXywFvYANFZtPqmG0pa0B+UErh1M3dCwV6NB/EF5zWWqNqT
# 9MRrP++bPh5hDRP/VrC3hCEGFwPQTfsHyaW3XR3ENS2jSDXdfH6c9WDYPtzgbfSa
# 72SOZ8+vhNlPetPf82VQvt9m58VP8nOomx1o3fOwriTra+s0YnA1tckl5oQRm3Pr
# 85AzGafNqA==
# SIG # End signature block
