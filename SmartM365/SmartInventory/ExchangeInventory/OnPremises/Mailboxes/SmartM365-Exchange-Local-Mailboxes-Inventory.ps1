<#
.SYNOPSIS
    Inventories local and optional remote Exchange mailboxes in an on-premises environment.

.DESCRIPTION
    This script scans and exports detailed information about Exchange mailboxes, including:
      - Full mailbox properties (name, alias, UPN, OU, etc.)
      - Permissions (FullAccess, SendAs, excluding NT AUTHORITY\SELF)
      - Mailbox statistics (size, item count, last logon, etc.)
      - Archive status and retention policies
      - Mobile device associations

    Supports both full-forest and targeted OU scans. Results are exported to CSV files.
    Can also generate the mailbox daily statistics report that used to live in the Report script.
    Includes robust logging, error handling, and backup of previous exports.
    Designed for Exchange 2016 servers with management tools and AD modules installed.
    Parameters allow customization of output paths, permission inclusion, and overwrite behavior.

.VERSION
1.29

.NOTES
    Version: 1.29
    Author: https://github.com/khda79/workplacecloudhub.com
    Requirements: Exchange 2016 Management Tools, Active Directory module
#>

[CmdletBinding()]
param (
    [string]$Tenant = 'test',
[Parameter(Mandatory = $false)]
    [string]$OutputPath,
	[Parameter(Mandatory = $false)]
    [string]$OutputPathOnlyADPermission,
    [Parameter(Mandatory = $false)]
    [string[]]$IncludedOrganizationalUnit = @(),
    [Parameter(Mandatory = $false)]
    [bool]$DetectAllDomains = $true,
    [Parameter(Mandatory = $false)]
    [bool]$IncludeADPermission = $false,
    [Parameter(Mandatory = $false)]
    [bool]$OnlyADPermission = $false,
    [Parameter(Mandatory = $false)]
    [bool]$ForceOverwriteCSV = $true,
    [Parameter(Mandatory = $false)]
    [bool]$GenerateReport = $true,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeRemoteMailboxes,
    [Parameter(Mandatory = $false)]
    [switch]$RemoteMailboxesOnly,
    [Parameter(Mandatory = $false)]
    [bool]$ReportOnly = $false,
    [Parameter(Mandatory = $false)]
    [bool]$DryRun = $false,
    [Parameter(Mandatory = $false)]
    [string[]]$TargetDomains = @(),
    [Parameter(Mandatory = $false)]
    [ValidateRange(1,168)]
    [int]$FileFreshnessHours = 4
)
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
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

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
            Write-Host 'Edit the local JSON now if needed. Press Enter to continue with the current file values.' -ForegroundColor Yellow
            Read-Host 'Press Enter to continue'
        }
    }

    try {
        return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
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
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
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
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) {
            return $DefaultValue
        }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

$ScriptLocalConfig = Get-ScriptLocalConfig



$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
#region Module Import and Initialization
$ScriptVersion = "1.29"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$EnableWeeklyHistory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableWeeklyHistory' -DefaultValue $true)
$WeeklyHistoryFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryFolderPath' -DefaultValue ''
$WeeklyHistoryRetentionWeeks = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCsvLogFolderPath' -DefaultValue $OutputPath
$RemoteMailboxOutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RemoteMailboxCsvLogFolderPath' -DefaultValue ''
$configuredIncludeRemoteMailboxes = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'IncludeRemoteMailboxes' -DefaultValue $true)
if (-not $PSBoundParameters.ContainsKey('IncludeRemoteMailboxes')) { $IncludeRemoteMailboxes = $configuredIncludeRemoteMailboxes }
if ($RemoteMailboxesOnly) { $IncludeRemoteMailboxes = $true }
function Resolve-SmartM365RemoteMailboxOutputPath {
    [CmdletBinding()]
    param(
        [string]$ConfiguredPath,
        [string]$LocalMailboxOutputPath
    )

    if ([string]::IsNullOrWhiteSpace($LocalMailboxOutputPath)) { return $ConfiguredPath }

    $defaultRemotePath = Join-Path -Path (Split-Path -Path $LocalMailboxOutputPath -Parent) -ChildPath 'RemoteMailboxes'
    if ([string]::IsNullOrWhiteSpace($ConfiguredPath)) { return $defaultRemotePath }

    $normalizedConfiguredPath = $ConfiguredPath.TrimEnd('\')
    if ($normalizedConfiguredPath -match '\\Exchange\\EXO\\Mailboxes$') {
        $legacyRemotePathMessage = ("RemoteMailboxCsvLogFolderPath used legacy EXO path; using on-prem remote mailbox path instead: {0}" -f $defaultRemotePath)
        if (Get-Command WriteLog -ErrorAction SilentlyContinue) { WriteLog -Message $legacyRemotePathMessage 'WARN' } else { Write-Warning $legacyRemotePathMessage }
        return $defaultRemotePath
    }

    return $ConfiguredPath
}
$RemoteMailboxOutputPath = Resolve-SmartM365RemoteMailboxOutputPath -ConfiguredPath $RemoteMailboxOutputPath -LocalMailboxOutputPath $OutputPath
function Ensure-SmartM365ExchangeScriptScope {
    [CmdletBinding()]
    param(
        [string[]]$RequiredCommands = @("Get-Mailbox", "Set-ADServerSettings"),
        [switch]$ViewEntireForest
    )

    $snapinName = "Microsoft.Exchange.Management.PowerShell.SnapIn"

    if (-not (Get-PSSnapin $snapinName -Registered -ErrorAction SilentlyContinue)) {
        WriteLog -Message "Exchange Management PSSnapin '$snapinName' is not registered on this server." "ERROR"
        return $false
    }

    if (-not (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue)) {
        try {
            if (-not (Get-PSSnapin $snapinName -ErrorAction SilentlyContinue)) {
                Add-PSSnapin $snapinName -ErrorAction Stop
                WriteLog -Message "Exchange snap-in loaded in script scope."
            }
        }
        catch {
            WriteLog -Message ("Failed to load Exchange snap-in in script scope: {0}" -f $_.Exception.Message) "ERROR"
            return $false
        }
    }

    $missingCommands = @($RequiredCommands | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
    if ($missingCommands.Count -gt 0) {
        WriteLog -Message ("Exchange cmdlet(s) not available in script scope: {0}" -f ($missingCommands -join ', ')) "ERROR"
        return $false
    }

    if ($ViewEntireForest) {
        try {
            Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
            WriteLog -Message "Set-ADServerSettings -ViewEntireForest True applied successfully."
        }
        catch {
            WriteLog -Message ("Failed to apply Set-ADServerSettings -ViewEntireForest True: {0}" -f $_.Exception.Message) "ERROR"
            return $false
        }
    }

    return $true
}
$LimitResultSize = $null
if ($LimitResultSize) {
    $TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion (LimitResultSize : $LimitResultSize)..."
}
$scriptdatamailbox = $false
$scriptdatamegewithperm = $true


function Register-SmartM365GeneratedCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not $global:csvGeneratedPaths) {
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }
    [void]$global:csvGeneratedPaths.Add($Path)
}

function Publish-SmartM365ExchangeLocalMailboxCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [string]$LatestFileName,
        [string]$HistoryLabel = 'Exchange on-prem mailboxes'
    )

    if ([string]::IsNullOrWhiteSpace($SourcePath) -or -not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { return }

    Register-SmartM365GeneratedCsv -Path $SourcePath
    $sourceUpload = $null
    try {
        $sourceUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $SourcePath
    } catch {
        WriteLog -Message ("Failed to upload source CSV to SharePoint for '{0}': {1}" -f $SourcePath, $_.Exception.Message) -Level 'WARNING'
    }

    $publishedPath = $SourcePath
    $latestUpload = $null
    $latestCsvFolder = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
    if (-not [string]::IsNullOrWhiteSpace($latestCsvFolder)) {
        try {
            if (-not (Test-Path -LiteralPath $latestCsvFolder)) { New-Item -ItemType Directory -Path $latestCsvFolder -Force | Out-Null }
            if ([string]::IsNullOrWhiteSpace($LatestFileName)) { $LatestFileName = Split-Path -Path $SourcePath -Leaf }
            $latestPath = Join-Path -Path $latestCsvFolder -ChildPath $LatestFileName
            Copy-Item -LiteralPath $SourcePath -Destination $latestPath -Force -ErrorAction Stop
            Register-SmartM365GeneratedCsv -Path $latestPath
            WriteLog -Message ("CSV latest copy written to: {0}" -f $latestPath)
            $latestUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestPath
            $publishedPath = $latestPath
        } catch {
            WriteLog -Message ("Failed to publish latest CSV copy for '{0}': {1}" -f $SourcePath, $_.Exception.Message) -Level 'WARNING'
        }
    }

    if ($EnableWeeklyHistory -and -not [string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath) -and (Get-Command Add-SmartM365WeeklyHistory -ErrorAction SilentlyContinue)) {
        Add-SmartM365WeeklyHistory -SourceCsvPaths @($publishedPath) -HistoryRootPath $WeeklyHistoryFolderPath -RetentionWeeks $WeeklyHistoryRetentionWeeks -HistoryLabel $HistoryLabel | Out-Null
    }

    return [pscustomobject]@{
        SourcePath   = $SourcePath
        LatestPath   = $publishedPath
        SourceUpload = $sourceUpload
        LatestUpload = $latestUpload
    }
}

function Get-SmartM365SharePointUploadRecordByLocalPath {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not $global:SmartM365SharePointUploadedFiles) { return $null }

    $resolvedPath = $Path
    try {
        if (Test-Path -LiteralPath $Path) {
            $resolvedPath = (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName
        }
    } catch {}

    $matches = @($global:SmartM365SharePointUploadedFiles | Where-Object {
        $_.LocalFilePath -and ([string]::Equals([string]$_.LocalFilePath, [string]$resolvedPath, [System.StringComparison]::OrdinalIgnoreCase))
    })

    if ($matches.Count -gt 0) { return $matches[-1] }
    return $null
}
# Atomic CSV export helper
function Export-CsvAtomic {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateSet("UTF8","UTF8BOM","Unicode","ASCII","Default")]
        [string]$Encoding = "UTF8",

        [Parameter(Mandatory = $false)]
        [string]$Delimiter = ","
    )

    Write-SmartM365CsvAtomically -Data @($InputObject) -Path $Path -Encoding $Encoding -Delimiter $Delimiter
    Register-SmartM365GeneratedCsv -Path $Path
}

function ConvertTo-SmartM365ExchangeMailboxSemicolonList {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    $items = @($Value | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    if ($items.Count -eq 0) { return '' }
    return (($items | Select-Object -Unique) -join ';')
}

function Get-SmartM365ExchangeRemoteMailboxDelegationSummary {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Identity)

    $result = [ordered]@{ FullAccessUsers = ''; SendAsUsers = '' }
    try {
        $fullAccess = Get-MailboxPermission -Identity $Identity -ErrorAction Stop | Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and -not $_.Deny -and $_.User -notlike 'NT AUTHORITY\SELF' } | Select-Object -ExpandProperty User
        $result.FullAccessUsers = ConvertTo-SmartM365ExchangeMailboxSemicolonList -Value $fullAccess
    }
    catch { $result.FullAccessUsers = "ERROR: $($_.Exception.Message)" }

    try {
        $sendAs = Get-ADPermission -Identity $Identity -ErrorAction Stop | Where-Object { $_.ExtendedRights -like '*Send-As*' -and -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\SELF' } | Select-Object -ExpandProperty User
        $result.SendAsUsers = ConvertTo-SmartM365ExchangeMailboxSemicolonList -Value $sendAs
    }
    catch { $result.SendAsUsers = "ERROR: $($_.Exception.Message)" }
    return [pscustomobject]$result
}

function ConvertTo-SmartM365ExchangeRemoteMailboxRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Mailbox,
        [switch]$IncludeDelegation
    )

    $domain = Get-SmartM365MailboxReportDomain -DistinguishedName ([string]$Mailbox.DistinguishedName)
    if ([string]::IsNullOrWhiteSpace($domain)) { $domain = 'Unknown' }

    $record = [ordered]@{
        DomainName                   = $domain
        Name                         = $Mailbox.Name
        DisplayName                  = $Mailbox.DisplayName
        Alias                        = $Mailbox.Alias
        PrimarySmtpAddress           = if ($Mailbox.PrimarySmtpAddress) { $Mailbox.PrimarySmtpAddress.ToString() } else { '' }
        WindowsEmailAddress          = if ($Mailbox.WindowsEmailAddress) { $Mailbox.WindowsEmailAddress.ToString() } else { '' }
        UserPrincipalName            = $Mailbox.UserPrincipalName
        SamAccountName               = $Mailbox.SamAccountName
        RecipientType                = if ($Mailbox.RecipientType) { $Mailbox.RecipientType.ToString() } else { '' }
        RecipientTypeDetails         = if ($Mailbox.RecipientTypeDetails) { $Mailbox.RecipientTypeDetails.ToString() } else { 'RemoteUserMailbox' }
        RemoteRoutingAddress         = if ($Mailbox.RemoteRoutingAddress) { $Mailbox.RemoteRoutingAddress.ToString() } else { '' }
        RemoteRecipientType          = if ($Mailbox.RemoteRecipientType) { $Mailbox.RemoteRecipientType.ToString() } else { '' }
        OnPremisesOrganizationalUnit = $Mailbox.OnPremisesOrganizationalUnit
        DistinguishedName            = $Mailbox.DistinguishedName
        ObjectCategory               = if ($Mailbox.ObjectCategory) { $Mailbox.ObjectCategory.ToString() } else { '' }
        WhenCreated                  = $Mailbox.WhenCreated
        WhenChanged                  = $Mailbox.WhenChanged
        MailboxRelease               = $Mailbox.MailboxRelease
        WhenMailboxCreated           = $Mailbox.WhenMailboxCreated
        AccountDisabled              = $Mailbox.AccountDisabled
        ExchangeUserAccountControl   = $Mailbox.ExchangeUserAccountControl
        ArchiveState                 = if ($Mailbox.ArchiveState) { $Mailbox.ArchiveState.ToString() } else { '' }
        ArchiveQuota                 = if ($Mailbox.ArchiveQuota) { $Mailbox.ArchiveQuota.ToString() } else { '' }
        ArchiveWarningQuota          = if ($Mailbox.ArchiveWarningQuota) { $Mailbox.ArchiveWarningQuota.ToString() } else { '' }
        DeliverToMailboxAndForward   = $Mailbox.DeliverToMailboxAndForward
        ForwardingAddress            = if ($Mailbox.ForwardingAddress) { $Mailbox.ForwardingAddress.ToString() } else { '' }
        IsValid                      = $Mailbox.IsValid
        MailboxMoveTargetMDB         = if ($Mailbox.MailboxMoveTargetMDB) { $Mailbox.MailboxMoveTargetMDB.ToString() } else { '' }
        MailboxMoveSourceMDB         = if ($Mailbox.MailboxMoveSourceMDB) { $Mailbox.MailboxMoveSourceMDB.ToString() } else { '' }
        MailboxMoveFlags             = if ($Mailbox.MailboxMoveFlags) { $Mailbox.MailboxMoveFlags.ToString() } else { '' }
        MailboxMoveRemoteHostName    = $Mailbox.MailboxMoveRemoteHostName
        MailboxMoveBatchName         = $Mailbox.MailboxMoveBatchName
        MailboxMoveStatus            = if ($Mailbox.MailboxMoveStatus) { $Mailbox.MailboxMoveStatus.ToString() } else { '' }
        SendOnBehalf                 = ConvertTo-SmartM365ExchangeMailboxSemicolonList -Value $Mailbox.GrantSendOnBehalfTo
    }

    if ($IncludeDelegation) {
        $delegation = Get-SmartM365ExchangeRemoteMailboxDelegationSummary -Identity ([string]$Mailbox.Identity)
        $record.FullAccessUsers = $delegation.FullAccessUsers
        $record.SendAsUsers = $delegation.SendAsUsers
    }

    $record.ExchangeGuid = if ($Mailbox.ExchangeGuid) { $Mailbox.ExchangeGuid.ToString() } else { '' }
    $immutableIdValue = ''
    if ($Mailbox.Guid) { try { $immutableIdValue = [System.Convert]::ToBase64String($Mailbox.Guid.ToByteArray()) } catch { $immutableIdValue = '' } }
    $record.ImmutableId = $immutableIdValue
    $record.ObjectGuid = if ($Mailbox.Guid) { $Mailbox.Guid.ToString() } else { '' }
    return [pscustomobject]$record
}

function ConvertFrom-SmartM365ExchangeRemoteMailboxWarnings {
    [CmdletBinding()]
    param([AllowNull()][object[]]$Warnings)

    $results = @()
    $seen = @{}
    $currentObjectPath = ''

    foreach ($warning in @($Warnings)) {
        if ($null -eq $warning) { continue }
        $message = [string]$warning
        if ([string]::IsNullOrWhiteSpace($message)) { continue }

        $issue = ''
        $suggestedAction = ''

        if ($message -match '^The object (?<Path>.+?) has been corrupted or isn''t compatible') {
            $currentObjectPath = $Matches['Path']
            $issue = 'ExchangeObjectInconsistentState'
            $suggestedAction = 'Open the recipient in Exchange Management Shell and repair the validation errors reported for this object.'
        }
        elseif ($message -match 'There is no primary SMTP address') {
            $issue = 'MissingPrimarySmtpAddress'
            $suggestedAction = 'Set or repair the primary SMTP address on the recipient before migration or synchronization decisions.'
        }
        elseif ($message -match 'ExternalEmailAddress is mandatory on MailUser') {
            $issue = 'MissingExternalEmailAddress'
            $suggestedAction = 'Set a valid ExternalEmailAddress or targetAddress for the MailUser or remote mailbox.'
        }
        elseif ($message -match 'mail contact and mail user must have a valid external e-mail address') {
            $issue = 'InvalidExternalEmailAddress'
            $suggestedAction = 'Review ExternalEmailAddress or targetAddress syntax and make sure it contains a valid routable SMTP address.'
        }
        elseif ($message -match 'The property "DisplayName".* is invalid') {
            $issue = 'InvalidDisplayName'
            $suggestedAction = 'Remove invalid leading or trailing whitespace or unsupported characters from DisplayName.'
        }

        if ([string]::IsNullOrWhiteSpace($issue)) { continue }

        $dedupeKey = ('{0}|{1}|{2}' -f $issue, $currentObjectPath, $message)
        if ($seen.ContainsKey($dedupeKey)) { continue }
        $seen[$dedupeKey] = $true

        $results += [pscustomobject]@{
            Issue           = $issue
            ObjectPath      = $currentObjectPath
            Warning         = $message
            SuggestedAction = $suggestedAction
        }
    }

    return @($results)
}
function Invoke-SmartM365ExchangeRemoteMailboxInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteOutputPath,
        [string[]]$IncludedLDAPPaths = @()
    )

    if ([string]::IsNullOrWhiteSpace($RemoteOutputPath)) { $RemoteOutputPath = Join-Path -Path (Split-Path -Path $OutputPath -Parent) -ChildPath 'RemoteMailboxes' }
    if (-not (Test-Path -LiteralPath $RemoteOutputPath)) { New-Item -ItemType Directory -Path $RemoteOutputPath -Force | Out-Null }
    WriteLog -Message ("Starting Exchange remote mailbox inventory. OutputPath: {0}" -f $RemoteOutputPath)
    Write-Host "Starting Exchange remote mailbox inventory... $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan

    $allRemoteMailboxes = @()
    $remoteWarningRecords = @()
    if ($IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0) {
        foreach ($scope in $IncludedLDAPPaths) {
            WriteLog -Message ("Retrieving remote mailboxes from scope: {0}" -f $scope)
            try {
                $remoteInScope = @(Get-RemoteMailbox -OnPremisesOrganizationalUnit $scope -ResultSize Unlimited -WarningVariable +remoteWarningRecords -ErrorAction Stop)
                $allRemoteMailboxes += $remoteInScope
                WriteLog -Message ("Found {0} remote mailboxes in scope: {1}" -f $remoteInScope.Count, $scope)
            }
            catch { WriteLog -Message ("WARNING: Failed to retrieve remote mailboxes from scope '{0}': {1}" -f $scope, $_.Exception.Message) }
        }
    }
    else { $allRemoteMailboxes = @(Get-RemoteMailbox -ResultSize Unlimited -WarningVariable remoteWarningRecords -ErrorAction Stop) }

    $dataQualityWarnings = @(ConvertFrom-SmartM365ExchangeRemoteMailboxWarnings -Warnings $remoteWarningRecords)
    $Global:SmartM365ExchangeRemoteMailboxDataQualityWarnings = @($dataQualityWarnings)
    if ($dataQualityWarnings.Count -gt 0) {
        WriteLog -Message ("Exchange remote mailbox data quality warnings captured: {0}" -f $dataQualityWarnings.Count) 'WARN'
    }

    $seenRemoteGuids = @{}
    $records = @()
    $index = 0
    $total = $allRemoteMailboxes.Count
    foreach ($remoteMailbox in $allRemoteMailboxes) {
        $index++
        $key = if ($remoteMailbox.Guid) { $remoteMailbox.Guid.ToString() } else { [string]$remoteMailbox.Identity }
        if (-not [string]::IsNullOrWhiteSpace($key) -and $seenRemoteGuids.ContainsKey($key)) { continue }
        if (-not [string]::IsNullOrWhiteSpace($key)) { $seenRemoteGuids[$key] = $true }
        if ($total -gt 0) { Write-Progress -Activity 'Processing Exchange remote mailboxes' -Status ("{0} of {1}: {2}" -f $index, $total, $remoteMailbox.Name) -PercentComplete (($index / $total) * 100) }
        $records += ConvertTo-SmartM365ExchangeRemoteMailboxRecord -Mailbox $remoteMailbox -IncludeDelegation:($IncludeADPermission -or $OnlyADPermission)
    }
    Write-Progress -Activity 'Processing Exchange remote mailboxes' -Completed

    if ($TargetDomains -and $TargetDomains.Count -gt 0) { $records = @($records | Where-Object { $_.DomainName -in $TargetDomains }) }
    if ($records.Count -eq 0) {
        WriteLog -Message 'No remote mailboxes were found. Remote mailbox CSV export skipped.'
        return [pscustomobject]@{ RecordCount = 0; CombinedCsv = $null; PerDomainCsvs = @(); PublishResult = $null; DataQualityWarnings = $dataQualityWarnings }
    }

    $suffix = if ($OnlyADPermission) { '_OnlyADPermission.csv' } else { '.csv' }
    $perDomainPaths = @()
    foreach ($group in ($records | Group-Object -Property DomainName)) {
        $safeDomain = if ([string]::IsNullOrWhiteSpace($group.Name)) { 'Unknown' } else { $group.Name -replace '[^a-zA-Z0-9.-]', '_' }
        $perDomainPath = Join-Path -Path $RemoteOutputPath -ChildPath ("Exchange_OnPrem_RemoteMailboxes_{0}{1}" -f $safeDomain, $suffix)
        Export-CsvAtomic -InputObject @($group.Group) -Path $perDomainPath -Encoding UTF8
        $perDomainPaths += $perDomainPath
    }

    $combinedPath = Join-Path -Path $RemoteOutputPath -ChildPath ("Exchange_OnPrem_RemoteMailboxes_AllDomains{0}" -f $suffix)
    Export-CsvAtomic -InputObject @($records) -Path $combinedPath -Encoding UTF8
    $publishResult = Publish-SmartM365ExchangeLocalMailboxCsv -SourcePath $combinedPath -LatestFileName (Split-Path -Path $combinedPath -Leaf) -HistoryLabel 'Exchange on-prem remote mailboxes'
    WriteLog -Message ("Exchange remote mailbox inventory completed. Records: {0}; CombinedCsv: {1}" -f $records.Count, $combinedPath)
    return [pscustomobject]@{ RecordCount = $records.Count; CombinedCsv = $combinedPath; PerDomainCsvs = $perDomainPaths; PublishResult = $publishResult; DataQualityWarnings = $dataQualityWarnings }
}
[string]$inputFolderCSVfiles
[string]$excludeMailboxesFile

[bool]$ExcludeAllDisabledAccounts = $false
[bool]$ExcludeDisabledAccountsExceptForSharedMailboxes = $false
[bool]$ExcludeDisabledMailboxes = $false
[bool]$IncludeSpecificLastLogonCriteria = $false
[string[]]$ExcludeSamAccountNamePatterns
[string[]]$excludeFullAccessSamAccounts
[string[]]$excludeSendAsSamAccounts
[string[]]$excludeSendOnBehalfToSamAccounts


[string]$SendFileListEmailReportFileName
$script:MailSmtpServer = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SmtpServer' -DefaultValue ''
$script:MailSmtpPort = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SmtpPort' -DefaultValue 25)
$script:MailSendMailMode = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SendMailMode' -DefaultValue ''
$script:MailFrom = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue ''
$script:MailTo = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue ''
$script:MailErrorTo = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ErrorMailTo' -DefaultValue ''
$script:MailCc = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Cc' -DefaultValue ''

function Send-SmartM365OptionalEmailHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BodyHtml,

        [Parameter(Mandatory = $false)]
        [string[]]$Attachments
    )

    $recipient = if (-not [string]::IsNullOrWhiteSpace($script:MailTo)) { $script:MailTo } elseif (-not [string]::IsNullOrWhiteSpace($script:MailErrorTo)) { $script:MailErrorTo } else { '' }
    if ([string]::IsNullOrWhiteSpace($script:MailFrom)) { throw 'Email notification requires From in configuration.' }
    if ([string]::IsNullOrWhiteSpace($recipient)) { throw 'Email notification requires To or ErrorMailTo in configuration.' }

    try {
        $mailParams = @{ BodyHtml = $BodyHtml; From = $script:MailFrom; To = $recipient }
        if (-not [string]::IsNullOrWhiteSpace($script:MailSmtpServer)) { $mailParams['SmtpServer'] = $script:MailSmtpServer }
        if ($script:MailSmtpPort -gt 0) { $mailParams['SmtpPort'] = $script:MailSmtpPort }
        if (-not [string]::IsNullOrWhiteSpace($script:MailSendMailMode)) { $mailParams['SendMailMode'] = $script:MailSendMailMode }
        if (-not [string]::IsNullOrWhiteSpace($script:MailCc)) { $mailParams['Cc'] = $script:MailCc }
        if ($Attachments -and $Attachments.Count -gt 0) { $mailParams['Attachments'] = $Attachments }
        SendEmailHtmlReport @mailParams
        WriteLog -Message ("Email notification sent. From: {0}; To: {1}" -f $script:MailFrom, $recipient)
    }
    catch {
        $message = "Email notification failed: $($_.Exception.Message)"
        if (Get-Command WriteLog -ErrorAction SilentlyContinue) { WriteLog -Message $message 'ERROR' }
        else { Write-Error $message }
        throw
    }
}

function Join-ModulePath {
param([Parameter(Mandatory)][string]$FileName)
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..')).Path
return (Join-Path (Join-Path (Join-Path (Join-Path $repoRoot 'Modules') 'SmartM365.Core') 'Compatibility\WindowsPowerShell5') $FileName)
}
function Stop-SmartM365TranscriptSafely {
    [CmdletBinding()]
    param()
    try { Stop-Transcript | Out-Null } catch {}
    try {
        $p = $null
        $v = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue
        if ($v -and $v.Value) { $p = $v.Value }
        else {
            $v = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue
            if ($v -and $v.Value) { $p = $v.Value }
        }
        if ($p) { Update-SmartM365TimestampedTranscript -Path $p }
    } catch {}
}

function Get-SmartM365MailboxReportValue {
    param($Row, [string[]]$Names)
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return $property.Value }
    }
    return $null
}

function Get-SmartM365MailboxReportDomain {
    param([string]$DistinguishedName)
    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return $null }
    $matches = [regex]::Matches($DistinguishedName, '(?i)(?:^|,)DC=([^,]+)')
    if ($matches.Count -eq 0) { return $null }
    $parts = @()
    foreach ($match in $matches) { $parts += $match.Groups[1].Value }
    return ($parts -join '.')
}

function ConvertTo-SmartM365MailboxReportRecord {
    param($Row, [switch]$Remote)
    $dn = Get-SmartM365MailboxReportValue -Row $Row -Names @('DistinguishedName', 'DN')
    $domain = Get-SmartM365MailboxReportValue -Row $Row -Names @('ADDomain', 'DomainName', 'Domain')
    if (-not $domain) { $domain = Get-SmartM365MailboxReportDomain -DistinguishedName $dn }
    $type = Get-SmartM365MailboxReportValue -Row $Row -Names @('RecipientTypeDetails', 'RecipientType', 'MailboxType')
    if (-not $type -and $Remote) { $type = 'RemoteUserMailbox' }
    $sizeRaw = Get-SmartM365MailboxReportValue -Row $Row -Names @('TotalItemSizeToMB', 'TotalItemSize-In-MB', 'TotalItemSizeMB', 'TotalSizeMB')
    $sizeMb = 0.0
    if ($sizeRaw) { [void][double]::TryParse(([string]$sizeRaw).Replace(',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$sizeMb) }
    $disabledRaw = Get-SmartM365MailboxReportValue -Row $Row -Names @('AccountDisabled', 'Disabled')
    [pscustomobject]@{
        DistinguishedName    = $dn
        AccountDisabled      = ([string]$disabledRaw -match '^(?i:true|1|yes|oui)$')
        RecipientTypeDetails = [string]$type
        ADDomain             = [string]$domain
        TotalItemSizeToMB    = [math]::Round($sizeMb, 2)
    }
}

function Import-SmartM365MailboxReportCsv {
    param([string]$Path)
    $header = Get-Content -LiteralPath $Path -First 1 -ErrorAction Stop
    if ($header -match ';') { return Import-Csv -LiteralPath $Path -Delimiter ';' }
    return Import-Csv -LiteralPath $Path
}

function Test-SmartM365MailboxReportCsvSchema {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "Mailbox report CSV not found: $Path" }
    $headerLine = Get-Content -LiteralPath $Path -First 1 -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($headerLine)) { throw "Mailbox report CSV is empty or has no header: $Path" }
    $delimiter = if ($headerLine -match ';') { ';' } else { ',' }
    $headers = @($headerLine -split [regex]::Escape($delimiter) | ForEach-Object { $_.Trim().Trim('"') })

    $missingGroups = New-Object 'System.Collections.Generic.List[string]'
    if (-not ($headers -contains 'DistinguishedName' -or $headers -contains 'DN')) { $missingGroups.Add('DistinguishedName or DN') }
    if (-not ($headers -contains 'RecipientTypeDetails' -or $headers -contains 'RecipientType' -or $headers -contains 'MailboxType')) { $missingGroups.Add('RecipientTypeDetails, RecipientType or MailboxType') }
    if ($missingGroups.Count -gt 0) {
        throw ("Mailbox report CSV schema is invalid for '{0}'. Missing column group(s): {1}" -f $Path, ($missingGroups -join '; '))
    }

    return $true
}
function ConvertTo-SmartM365MailboxReportHtml {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 'N/A' }
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return 'N/A' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-SmartM365MailboxReportCsvRowCount {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return 0 }
    try {
        $lineCount = @(Get-Content -LiteralPath $Path -ErrorAction Stop | Measure-Object -Line).Lines
        if ($lineCount -le 1) { return 0 }
        return ($lineCount - 1)
    }
    catch { return 0 }
}

function New-SmartM365ExchangeLocalMailboxReportEmailBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ReportRows,

        [Parameter(Mandatory = $true)]
        [string]$DailyStatsCsv,

        [string]$LatestDailyStatsCsv,

        [Parameter(Mandatory = $true)]
        [string]$SummaryCsv,

        [string]$LocalMailboxCsv,
        [string]$LatestLocalMailboxCsv,
        [string]$RemoteMailboxCsv,
        [string]$LatestRemoteMailboxCsv,

        [AllowNull()]$DailyStatsUpload,
        [AllowNull()]$SummaryUpload,
        [AllowNull()]$LocalMailboxUpload,
        [AllowNull()]$RemoteMailboxUpload,

        [object[]]$RemoteMailboxDataQualityWarnings = @(),

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    function Format-SmartM365ExchangeMailboxReportNumber {
        param([AllowNull()]$Value)
        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '0' }
        try {
            $number = [double]$Value
            if ([math]::Abs($number % 1) -lt 0.000001) { return $number.ToString('N0') }
            return $number.ToString('N2')
        }
        catch {
            return [string]$Value
        }
    }

    function New-SmartM365ExchangeMailboxReportLinkHtml {
        param(
            [string]$Text,
            [string]$Url
        )

        $safeText = ConvertTo-SmartM365EmailHtmlText $Text
        if ([string]::IsNullOrWhiteSpace($Url)) { return $safeText }
        $safeUrl = ConvertTo-SmartM365EmailHtmlText $Url
        return ('<a href="{0}" style="color:#075985;text-decoration:underline;">{1}</a>' -f $safeUrl, $safeText)
    }

    $rows = @($ReportRows)
    $domainCount = $rows.Count
    $totalMailboxes = ($rows | Measure-Object -Property TotalMailboxCount -Sum).Sum
    $totalLocalMailboxes = ($rows | Measure-Object -Property TotalLocalMailboxCount -Sum).Sum
    $totalRemoteMailboxes = ($rows | Measure-Object -Property TotalRemoteMailboxCount -Sum).Sum
    $totalEnabledAccounts = ($rows | Measure-Object -Property EnabledAccounts -Sum).Sum
    $totalDisabledAccounts = ($rows | Measure-Object -Property DisabledAccounts -Sum).Sum
    $totalSizeGb = ($rows | Measure-Object -Property TotalLocalMailboxSizeGB -Sum).Sum
    if ($null -eq $totalMailboxes) { $totalMailboxes = 0 }
    if ($null -eq $totalLocalMailboxes) { $totalLocalMailboxes = 0 }
    if ($null -eq $totalRemoteMailboxes) { $totalRemoteMailboxes = 0 }
    if ($null -eq $totalEnabledAccounts) { $totalEnabledAccounts = 0 }
    if ($null -eq $totalDisabledAccounts) { $totalDisabledAccounts = 0 }
    if ($null -eq $totalSizeGb) { $totalSizeGb = 0 }
    $totalSizeGb = [math]::Round([double]$totalSizeGb, 2)

    $summaryRows = @(
        [pscustomobject]@{ Label = 'Domains'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $domainCount }
        [pscustomobject]@{ Label = 'Total mailboxes'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalMailboxes }
        [pscustomobject]@{ Label = 'Local mailboxes'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalLocalMailboxes }
        [pscustomobject]@{ Label = 'Remote mailboxes'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalRemoteMailboxes }
        [pscustomobject]@{ Label = 'Enabled accounts'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalEnabledAccounts }
        [pscustomobject]@{ Label = 'Disabled accounts'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalDisabledAccounts }
        [pscustomobject]@{ Label = 'Local mailbox size GB'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalSizeGb }
    )

    $domainRowsHtml = foreach ($row in ($rows | Sort-Object -Property @{ Expression = 'TotalMailboxCount'; Descending = $true }, DomainName)) {
        @"
<tr>
  <td style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText $row.DomainName)</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.TotalMailboxCount))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.TotalLocalMailboxCount))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.TotalRemoteMailboxCount))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.EnabledAccounts))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.DisabledAccounts))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.TotalLocalMailboxSizeGB))</td>
</tr>
"@
    }

    $domainTableHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Domain</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Total</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Local</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Remote</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Enabled</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Disabled</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Local GB</th>
  </tr>
  $($domainRowsHtml -join "`r`n")
</table>
"@

    $fileRows = @()
    $localMailboxPathForMail = if (-not [string]::IsNullOrWhiteSpace($LatestLocalMailboxCsv)) { $LatestLocalMailboxCsv } else { $LocalMailboxCsv }
    if (-not [string]::IsNullOrWhiteSpace($localMailboxPathForMail)) {
        $fileRows += [pscustomobject]@{ Label = 'Mailbox inventory'; Path = $localMailboxPathForMail; WebUrl = if ($LocalMailboxUpload -and $LocalMailboxUpload.WebUrl) { [string]$LocalMailboxUpload.WebUrl } else { '' } }
    }
    $remoteMailboxPathForMail = if (-not [string]::IsNullOrWhiteSpace($LatestRemoteMailboxCsv)) { $LatestRemoteMailboxCsv } else { $RemoteMailboxCsv }
    if (-not [string]::IsNullOrWhiteSpace($remoteMailboxPathForMail) -and (Test-Path -LiteralPath $remoteMailboxPathForMail -PathType Leaf)) {
        $fileRows += [pscustomobject]@{ Label = 'Remote mailbox inventory'; Path = $remoteMailboxPathForMail; WebUrl = if ($RemoteMailboxUpload -and $RemoteMailboxUpload.WebUrl) { [string]$RemoteMailboxUpload.WebUrl } else { '' } }
    }
    if (-not [string]::IsNullOrWhiteSpace($LatestDailyStatsCsv)) {
        $fileRows += [pscustomobject]@{ Label = 'Daily stats'; Path = $LatestDailyStatsCsv; WebUrl = if ($DailyStatsUpload -and $DailyStatsUpload.WebUrl) { [string]$DailyStatsUpload.WebUrl } else { '' } }
    }
    else {
        $fileRows += [pscustomobject]@{ Label = 'Daily stats'; Path = $DailyStatsCsv; WebUrl = if ($DailyStatsUpload -and $DailyStatsUpload.WebUrl) { [string]$DailyStatsUpload.WebUrl } else { '' } }
    }
    $fileRows += [pscustomobject]@{ Label = 'Summary'; Path = $SummaryCsv; WebUrl = if ($SummaryUpload -and $SummaryUpload.WebUrl) { [string]$SummaryUpload.WebUrl } else { '' } }
    $fileRowsHtml = foreach ($fileRow in $fileRows) {
        $rowCount = Get-SmartM365MailboxReportCsvRowCount -Path $fileRow.Path
        $pathHtml = New-SmartM365ExchangeMailboxReportLinkHtml -Text $fileRow.Path -Url $fileRow.WebUrl
        @"
<tr>
  <td style="width:130px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText $fileRow.Label)</td>
  <td style="border-bottom:1px solid #eef2f7;padding:10px 12px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;">$pathHtml</td>
  <td align="right" style="width:80px;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $rowCount))</td>
</tr>
"@
    }

    $filesTableHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">File</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Path / SharePoint link</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Rows</th>
  </tr>
  $($fileRowsHtml -join "`r`n")
</table>
"@

    $sections = @(
        [pscustomobject]@{ Title = 'Domain summary'; Html = $domainTableHtml }
    )

    $exchangeDataQualityWarnings = @($RemoteMailboxDataQualityWarnings | Select-Object -First 100)
    if ($exchangeDataQualityWarnings.Count -gt 0) {
        $warningRowsHtml = foreach ($warningRow in $exchangeDataQualityWarnings) {
            $issueLabel = switch ([string]$warningRow.Issue) {
                'ExchangeObjectInconsistentState' { 'Inconsistent state'; break }
                'MissingPrimarySmtpAddress' { 'Missing primary SMTP'; break }
                'MissingExternalEmailAddress' { 'Missing external address'; break }
                'InvalidExternalEmailAddress' { 'Invalid external address'; break }
                'InvalidDisplayName' { 'Invalid display name'; break }
                default { [string]$warningRow.Issue }
            }
            @"
<tr>
  <td valign="top" style="width:165px;border-bottom:1px solid #fde68a;padding:10px 12px;font-size:12px;color:#92400e;font-weight:700;line-height:1.35;">$(ConvertTo-SmartM365EmailHtmlText $issueLabel)</td>
  <td valign="top" style="border-bottom:1px solid #fde68a;padding:10px 12px;font-size:12px;color:#78350f;line-height:1.4;">
    <div style="font-weight:700;margin:0 0 6px 0;color:#78350f;">$(ConvertTo-SmartM365EmailHtmlText $warningRow.Warning)</div>
    <div style="margin:0 0 6px 0;font-family:Consolas,'Courier New',monospace;font-size:11px;color:#92400e;word-break:break-word;overflow-wrap:anywhere;">$(ConvertTo-SmartM365EmailHtmlText $warningRow.ObjectPath)</div>
    <div style="margin:0;color:#92400e;">$(ConvertTo-SmartM365EmailHtmlText $warningRow.SuggestedAction)</div>
  </td>
</tr>
"@
        }
        $warningsTableHtml = @"
<p style="margin:0 0 10px 0;color:#78350f;font-size:13px;">Top 100 Exchange object data quality warnings captured during Get-RemoteMailbox.</p>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #f59e0b;font-size:12px;background:#fffbeb;table-layout:fixed;">
  <tr>
    <th align="left" style="width:165px;background:#fef3c7;border-bottom:1px solid #f59e0b;padding:10px 12px;color:#92400e;text-transform:uppercase;">Issue</th>
    <th align="left" style="background:#fef3c7;border-bottom:1px solid #f59e0b;padding:10px 12px;color:#92400e;text-transform:uppercase;">Details</th>
  </tr>
  $($warningRowsHtml -join "`r`n")
</table>
"@
        $sections += [pscustomobject]@{ Title = 'Exchange object data quality warnings - Top 100'; Html = $warningsTableHtml }
    }

    $sections += [pscustomobject]@{ Title = 'Files'; Html = $filesTableHtml }

    return New-SmartM365EmailBody `
        -Title 'Mailbox inventory summary' `
        -Category 'SmartM365 Exchange OnPrem' `
        -Severity 'Success' `
        -Tenant $Tenant `
        -Message 'Exchange on-premises mailbox inventory summary generated from the latest SmartM365 CSV outputs.' `
        -SummaryRows $summaryRows `
        -Sections $sections `
        -Footer 'This automated message was generated by SmartM365 from Exchange on-premises mailbox inventory data.'
}

function Invoke-SmartM365ExchangeLocalMailboxReport {
    [CmdletBinding()]
    param([bool]$UseCurrentInventoryData = $false)
    $reportOutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxReportCsvLogFolderPath' -DefaultValue (Join-Path -Path $OutputPath -ChildPath 'Reports')
    if (-not (Test-Path -LiteralPath $reportOutputPath)) { New-Item -ItemType Directory -Path $reportOutputPath -Force | Out-Null }
    $localMailboxFolder = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCsvLogFolderPath' -DefaultValue $OutputPath
    $remoteMailboxFolder = $RemoteMailboxOutputPath
    $latestCsvFolder = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
    $localCsv = Join-Path -Path $localMailboxFolder -ChildPath 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
    $latestLocalCsv = if ($latestCsvFolder) { Join-Path -Path $latestCsvFolder -ChildPath 'Exchange_OnPrem_Mailboxes_AllDomains.csv' } else { $null }
    $remoteCsv = if ($remoteMailboxFolder) { Join-Path -Path $remoteMailboxFolder -ChildPath 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv' } else { $null }
    $latestRemoteCsv = if ($latestCsvFolder) { Join-Path -Path $latestCsvFolder -ChildPath 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv' } else { $null }
    $dailyCsv = Join-Path -Path $reportOutputPath -ChildPath 'Exchange_OnPrem_Mailboxes_DailyStats.csv'
    $summaryCsv = Join-Path -Path $reportOutputPath -ChildPath 'Exchange_OnPrem_Mailboxes_DailyStats_Summary.csv'
    $latestDailyCsv = if ($latestCsvFolder) { Join-Path -Path $latestCsvFolder -ChildPath 'Exchange_OnPrem_Mailboxes_DailyStats.csv' } else { $null }
    $latestSummaryCsv = if ($latestCsvFolder) { Join-Path -Path $latestCsvFolder -ChildPath 'Exchange_OnPrem_Mailboxes_DailyStats_Summary.csv' } else { $null }
    $allowedTypes = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox', 'RemoteUserMailbox', 'RemoteSharedMailbox')
    $remoteTypes = @('RemoteUserMailbox', 'RemoteSharedMailbox')
    $records = @()
    $remoteCsvAvailable = ($remoteCsv -and (Test-Path -LiteralPath $remoteCsv -PathType Leaf))
    if ($UseCurrentInventoryData -and $Global:ScriptOverallMailboxData -and $Global:ScriptOverallMailboxData.Count -gt 0) {
        $records += $Global:ScriptOverallMailboxData | ForEach-Object { ConvertTo-SmartM365MailboxReportRecord -Row $_ }
    }
    elseif (Test-Path -LiteralPath $localCsv -PathType Leaf) {
        [void](Test-SmartM365MailboxReportCsvSchema -Path $localCsv)
        $records += Import-SmartM365MailboxReportCsv -Path $localCsv | ForEach-Object { ConvertTo-SmartM365MailboxReportRecord -Row $_ }
    }
    elseif (-not $remoteCsvAvailable) {
        throw ('Local mailbox CSV not found for report generation: {0}' -f $localCsv)
    }
    else {
        WriteLog -Message ('Local mailbox CSV not found; mailbox report will use remote mailbox CSV only. Missing local path: {0}' -f $localCsv) 'WARN'
    }
    if ($remoteCsvAvailable) { [void](Test-SmartM365MailboxReportCsvSchema -Path $remoteCsv); $records += Import-SmartM365MailboxReportCsv -Path $remoteCsv | ForEach-Object { ConvertTo-SmartM365MailboxReportRecord -Row $_ -Remote } }
    $records = @($records | Where-Object { $_.RecipientTypeDetails -and ($allowedTypes -contains $_.RecipientTypeDetails) })
    if ($TargetDomains -and $TargetDomains.Count -gt 0) { $records = @($records | Where-Object { $_.ADDomain -in $TargetDomains }) }
    if ($records.Count -eq 0) { WriteLog -Message 'No mailbox data available for report generation.' 'WARN'; return $null }
    $mailboxTypes = $allowedTypes | Sort-Object
    $report = @($records | Group-Object -Property ADDomain | ForEach-Object {
        $domain = $_.Name
        $items = @($_.Group)
        $totalMb = ($items | Measure-Object -Property TotalItemSizeToMB -Sum).Sum
        if ($null -eq $totalMb) { $totalMb = 0 }
        $row = [ordered]@{ Date = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); DomainName = $domain; TotalMailboxCount = $items.Count; TotalLocalMailboxCount = @($items | Where-Object { $_.RecipientTypeDetails -notin $remoteTypes }).Count; TotalRemoteMailboxCount = @($items | Where-Object { $_.RecipientTypeDetails -in $remoteTypes }).Count; EnabledAccounts = @($items | Where-Object { $_.AccountDisabled -eq $false }).Count; DisabledAccounts = @($items | Where-Object { $_.AccountDisabled -eq $true }).Count; TotalLocalMailboxSizeGB = [math]::Round(([double]$totalMb / 1024), 2) }
        $counts = @($items | Group-Object -Property RecipientTypeDetails -NoElement)
        foreach ($type in $mailboxTypes) { $row[$type] = @($counts | Where-Object { $_.Name -eq $type }).Count }
        [pscustomobject]$row
    })
    $dailyRows = @()
    if (Test-Path -LiteralPath $dailyCsv) { $dailyRows += Import-Csv -LiteralPath $dailyCsv -Delimiter ';' }
    $dailyRows += $report
    Export-CsvAtomic -InputObject $dailyRows -Path $dailyCsv -Encoding UTF8 -Delimiter ';'
    $summaryColumns = @('DomainName', 'TotalMailboxCount', 'TotalLocalMailboxCount', 'TotalRemoteMailboxCount', 'EnabledAccounts', 'DisabledAccounts', 'TotalLocalMailboxSizeGB')
    $summary = @($report | Select-Object $summaryColumns)
    $totalRow = [ordered]@{ DomainName = 'TOTAL' }
    foreach ($column in $summaryColumns) { if ($column -ne 'DomainName') { $sum = ($summary | Measure-Object -Property $column -Sum).Sum; if ($column -eq 'TotalLocalMailboxSizeGB') { $sum = [math]::Round([double]$sum, 2) }; $totalRow[$column] = $sum } }
    Export-CsvAtomic -InputObject @($summary + ([pscustomobject]$totalRow)) -Path $summaryCsv -Encoding UTF8 -Delimiter ';'
    $latestDailyUpload = $null
    if ($latestDailyCsv) {
        $latestDir = Split-Path -Path $latestDailyCsv -Parent
        if (-not (Test-Path -LiteralPath $latestDir)) { New-Item -ItemType Directory -Path $latestDir -Force | Out-Null }
        Copy-Item -LiteralPath $dailyCsv -Destination $latestDailyCsv -Force
        try {
            $latestDailyUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestDailyCsv
        }
        catch {
            WriteLog -Message ('Failed to upload latest daily stats CSV to SharePoint: {0}' -f $_.Exception.Message) 'WARN'
        }
    }
    $summaryUpload = $null
    if ($latestSummaryCsv) {
        $latestSummaryDir = Split-Path -Path $latestSummaryCsv -Parent
        if (-not (Test-Path -LiteralPath $latestSummaryDir)) { New-Item -ItemType Directory -Path $latestSummaryDir -Force | Out-Null }
        Copy-Item -LiteralPath $summaryCsv -Destination $latestSummaryCsv -Force
        try {
            $summaryUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestSummaryCsv
        }
        catch {
            WriteLog -Message ('Failed to upload latest mailbox summary CSV to SharePoint: {0}' -f $_.Exception.Message) 'WARN'
        }
    }
    if (-not $summaryUpload) {
        try {
            $summaryUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $summaryCsv
        }
        catch {
            WriteLog -Message ('Failed to upload mailbox summary CSV to SharePoint: {0}' -f $_.Exception.Message) 'WARN'
        }
    }
    $historyDailySource = if ($latestDailyCsv -and (Test-Path -LiteralPath $latestDailyCsv -PathType Leaf)) { $latestDailyCsv } else { $dailyCsv }
    $historySummarySource = if ($latestSummaryCsv -and (Test-Path -LiteralPath $latestSummaryCsv -PathType Leaf)) { $latestSummaryCsv } else { $summaryCsv }
    if ($EnableWeeklyHistory -and -not [string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath) -and (Get-Command Add-SmartM365WeeklyHistory -ErrorAction SilentlyContinue)) {
        Add-SmartM365WeeklyHistory -SourceCsvPaths @($historyDailySource, $historySummarySource) -HistoryRootPath $WeeklyHistoryFolderPath -RetentionWeeks $WeeklyHistoryRetentionWeeks -HistoryLabel 'Exchange on-prem mailbox daily stats' | Out-Null
    }
    RemoveOldFiles -Path $reportOutputPath -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
    $localMailboxUpload = if ($latestLocalCsv) { Get-SmartM365SharePointUploadRecordByLocalPath -Path $latestLocalCsv } else { $null }
    if (-not $localMailboxUpload) { $localMailboxUpload = Get-SmartM365SharePointUploadRecordByLocalPath -Path $localCsv }
    $remoteMailboxUpload = if ($latestRemoteCsv) { Get-SmartM365SharePointUploadRecordByLocalPath -Path $latestRemoteCsv } else { $null }
    if (-not $remoteMailboxUpload -and $remoteCsv) { $remoteMailboxUpload = Get-SmartM365SharePointUploadRecordByLocalPath -Path $remoteCsv }
    $summaryCsvForMail = if ($latestSummaryCsv -and (Test-Path -LiteralPath $latestSummaryCsv -PathType Leaf)) { $latestSummaryCsv } else { $summaryCsv }
    $remoteMailboxDataQualityWarnings = @($Global:SmartM365ExchangeRemoteMailboxDataQualityWarnings)
    try {
        $mailBody = New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows $report -DailyStatsCsv $dailyCsv -LatestDailyStatsCsv $latestDailyCsv -SummaryCsv $summaryCsvForMail -LocalMailboxCsv $localCsv -LatestLocalMailboxCsv $latestLocalCsv -RemoteMailboxCsv $remoteCsv -LatestRemoteMailboxCsv $latestRemoteCsv -DailyStatsUpload $latestDailyUpload -SummaryUpload $summaryUpload -LocalMailboxUpload $localMailboxUpload -RemoteMailboxUpload $remoteMailboxUpload -RemoteMailboxDataQualityWarnings $remoteMailboxDataQualityWarnings -Title ($TaskName + ' - Mailbox report')
        Send-SmartM365OptionalEmailHtmlReport -BodyHtml $mailBody
    }
    catch {
        WriteLog -Message ('Failed to send mailbox report email: {0}' -f $_.Exception.Message) 'ERROR'
        throw
    }
    return [pscustomobject]@{ DailyStatsCsv = $dailyCsv; SummaryCsv = $summaryCsvForMail; RowCount = $report.Count }
}
try {
    Write-Host "Loading module SmartM365-WindowsPowerShell5.psd1..."
    Import-Module (Join-ModulePath 'SmartM365-WindowsPowerShell5.psd1') -ErrorAction Stop
	$InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
	Start-Transcript -Path $global:logTranscriptFile -Append
	WriteLog -Message $MyInvocation.MyCommand.Name
	WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
	$OutputPath = $InitializeOutputPath
    if ([string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath)) { $WeeklyHistoryFolderPath = Join-Path -Path $OutputPath -ChildPath 'WeeklyHistory' }
	WriteLog -Message "Starting $TaskName..."
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    throw
}
#endregion

$StartTime = Get-Date

if ($DryRun) {
    $dryRunError = $null
    try {
        WriteLog -Message 'DryRun mode enabled. Inventory collection will be skipped.'
        Write-Host 'DryRun mode enabled. Inventory collection will be skipped.' -ForegroundColor Cyan
        Write-Host "OutputPath: $OutputPath"

        $reportPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxReportCsvLogFolderPath' -DefaultValue (Join-Path -Path $OutputPath -ChildPath 'Reports')
        Write-Host "ReportPath: $reportPath"

        if ($ReportOnly -or $GenerateReport) {
            $localMailboxFolder = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCsvLogFolderPath' -DefaultValue $OutputPath
            $localCsv = Join-Path -Path $localMailboxFolder -ChildPath 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
            if (Test-Path -LiteralPath $localCsv) { [void](Test-SmartM365MailboxReportCsvSchema -Path $localCsv); Write-Host "Report local CSV schema OK: $localCsv" -ForegroundColor Green }
            elseif ($ReportOnly) { throw "ReportOnly requires the local mailbox CSV or Exchange live collection: $localCsv" }
        }

        if (-not $ReportOnly) {
            if (-not (Ensure-SmartM365ExchangeScriptScope -RequiredCommands @(if ($IncludeRemoteMailboxes) { "Get-Mailbox"; "Get-RemoteMailbox"; "Set-ADServerSettings" } else { "Get-Mailbox"; "Set-ADServerSettings" }) -ViewEntireForest)) {
                throw 'Exchange Management Tools were not detected or the Exchange snap-in could not be loaded.'
            }

            $exchangeCmdletAvailable = [bool](Get-Command Get-Mailbox -ErrorAction SilentlyContinue)
            Write-Host "Exchange Get-Mailbox available: $exchangeCmdletAvailable"
            if (-not $exchangeCmdletAvailable) { throw 'Get-Mailbox is not available after Exchange snap-in load.' }
        }
    }
    catch {
        $dryRunError = $_
        WriteLog -Message ("DryRun failed: {0}" -f $_.Exception.Message) "ERROR"
        throw
    }
    finally {
        Stop-SmartM365TranscriptSafely
        if ($dryRunError) { try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $dryRunError -FailureStage 'DryRun' } catch {} }
        else { try { Complete-SmartM365ExecutionContext -Status Auto } catch {} }
    }
    return
}
if ($ReportOnly) {
    $reportOnlyError = $null
    try {
        WriteLog -Message 'ReportOnly mode enabled. Inventory collection will be skipped.'
        Invoke-SmartM365ExchangeLocalMailboxReport -UseCurrentInventoryData:$false | Out-Null
    }
    catch {
        $reportOnlyError = $_
        WriteLog -Message ("ReportOnly execution failed: {0}" -f $_.Exception.Message) "ERROR"
        throw
    }
    finally {
        Stop-SmartM365TranscriptSafely
        if ($reportOnlyError) { try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $reportOnlyError -FailureStage 'ReportOnly' } catch {} }
        else { try { Complete-SmartM365ExecutionContext -Status Auto } catch {} }
    }
    return
}

if (-not (Ensure-SmartM365ExchangeScriptScope -RequiredCommands @(if ($IncludeRemoteMailboxes) { "Get-Mailbox"; "Get-RemoteMailbox"; "Set-ADServerSettings" } else { "Get-Mailbox"; "Set-ADServerSettings" }) -ViewEntireForest)) {
    Write-Error "Exchange environment not ready. Exiting script."
    $errorMessage = "Exchange environment not ready. Exiting script."
    $body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
    Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
    exit 1
}
$scriptScopeExchangeCommands = @(
    "Get-Mailbox",
    "Get-MailboxDatabase",
    "Get-MailboxPermission",
    "Get-MailboxStatistics",
    "Get-MobileDevice",
    "Get-ADPermission"
)
if ($IncludeRemoteMailboxes) { $scriptScopeExchangeCommands += "Get-RemoteMailbox" }
if (-not (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue)) {
    WriteLog -Message "Exchange cmdlets are loaded in module scope; loading Exchange snap-in in script scope for mailbox processing."
    try {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
    }
    catch {
        WriteLog -Message ("Failed to load Exchange snap-in in script scope: {0}" -f $_.Exception.Message) "ERROR"
    }
}

$missingScriptScopeExchangeCommands = @($scriptScopeExchangeCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
if ($missingScriptScopeExchangeCommands.Count -gt 0) {
    $errorMessage = "Exchange cmdlet(s) not available in script scope after snap-in load: $($missingScriptScopeExchangeCommands -join ', ')"
    Write-Error $errorMessage
    WriteLog -Message $errorMessage "ERROR"
    $body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
    Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
    throw $errorMessage
}

$preflightOutputPaths = @($OutputPath)
if ($IncludeRemoteMailboxes) {
    if ([string]::IsNullOrWhiteSpace($RemoteMailboxOutputPath)) {
        $RemoteMailboxOutputPath = Join-Path -Path (Split-Path -Path $OutputPath -Parent) -ChildPath 'RemoteMailboxes'
    }
    if (-not (Test-Path -LiteralPath $RemoteMailboxOutputPath)) { New-Item -ItemType Directory -Path $RemoteMailboxOutputPath -Force | Out-Null }
    $preflightOutputPaths += $RemoteMailboxOutputPath
}
if ($OnlyADPermission -or $IncludeADPermission) {
    try {
        if ([string]::IsNullOrWhiteSpace($OutputPathOnlyADPermission)) {
            $OutputPathOnlyADPermission = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxOnlyAdPermissionCsvLogFolderPath' -DefaultValue ''
        }
    } catch {
        WriteLog -Message "ERROR retrieving LocalMailboxOnlyAdPermissionCsvLogFolderPath: $_" "ERROR"
        throw
    }

    if ([string]::IsNullOrWhiteSpace($OutputPathOnlyADPermission) -or -not (Test-Path $OutputPathOnlyADPermission)) {
        $errorMessage = "The share '$OutputPathOnlyADPermission' is not available. Stopping the script."
        Write-Host $errorMessage -ForegroundColor Red
        $body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
        Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
        throw $errorMessage
    }

    Write-Host "The network share '$OutputPathOnlyADPermission' is available. Continuing the script..." -ForegroundColor Green
    $preflightOutputPaths += $OutputPathOnlyADPermission
}

Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths $preflightOutputPaths -RequireExchangeOnPrem -RequireActiveDirectoryRead | Out-Null

# Refined parameter logic for permissions
if ($OnlyADPermission) {
    $IncludeADPermission = $true
    if ([string]::IsNullOrWhiteSpace($OutputPathOnlyADPermission)) {
        throw "OutputPathOnlyADPermission is required when OnlyADPermission is set"
    }
    $OutputPath = $OutputPathOnlyADPermission

    # In OnlyADPermission mode, do not merge with permissions inventory nor batch naming.
    $scriptdatamegewithperm = $false
}
if ($IncludedOrganizationalUnit.Count -ne 0) {
    $DetectAllDomains = $false
}
$TimeStampForLogOnly = Get-Date -Format "yyyyMMdd_HHmmss" # Timestamp for log files to ensure uniqueness per run

# Initialize global caches
if (-not $Global:DomainInfoCache) {
    $Global:DomainInfoCache = @{} # For NetBIOSName to DNSRoot (FQDN) mapping
}
if (-not $Global:ADObjectCache) {
    $Global:ADObjectCache = @{} # For SID to ADObject mapping (or $null if not found)
}
if (-not $Global:GroupMemberCache) {
    $Global:GroupMemberCache = @{} # For Group DN to member SamAccountNames list mapping (or error/empty marker)
}
if (-not $Global:DomainFQDNToNetBIOSCache) {
    $Global:DomainFQDNToNetBIOSCache = @{} # For Domain FQDN to NetBIOS Name mapping
}
$Global:ScriptOverallMailboxData = @() # To accumulate all mailbox data for a final combined CSV
$Script:MailboxesProcessingLogFile = $null # Initialize script-scoped variable for detailed log file path

# Variable to track successful script completion
$InventoryCompletedSuccessfully = $false


try { # Main try block for script execution and interruption handling
    WriteLog -Message "Starting script '$PSCommandPath' - Version $ScriptVersion"
    Write-Host "Starting script '$($MyInvocation.MyCommand.Name)' - Version $ScriptVersion ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
    Write-Host "Output Path for this run: $OutputPath"
    WriteLog -Message "Output Path for this run: $OutputPath"
WriteLog -Message "Effective permission flags: IncludeADPermission = $IncludeADPermission, OnlyADPermission = $OnlyADPermission, IncludeRemoteMailboxes = $IncludeRemoteMailboxes, RemoteMailboxesOnly = $RemoteMailboxesOnly, ForceOverwriteCSV = $ForceOverwriteCSV"
    Write-Host "Effective permission flags: IncludeADPermission = $IncludeADPermission, OnlyADPermission = $OnlyADPermission, IncludeRemoteMailboxes = $IncludeRemoteMailboxes, RemoteMailboxesOnly = $RemoteMailboxesOnly, ForceOverwriteCSV = $ForceOverwriteCSV"
    Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))

    # ViewEntireForest is applied during Exchange readiness validation.
    WriteLog -Message "Set-ADServerSettings -ViewEntireForest $true was applied during Exchange readiness validation."

    #region Function Definitions

    # This function contains the detailed mailbox processing logic
    function MailboxesProcessing2 {
        param (
            [Parameter(Mandatory=$true)]
            [string[]]$IncludedLDAPPaths
        )
        $FunctionStartTime = Get-Date

        # --- START NEW LOGIC FOR LOG FILE IDENTIFIER ---
        $logPathIdentifier = "AllMailboxes"
        if ($IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0) {
            if ($IncludedLDAPPaths.Count -eq 1) {
                $firstPath = $IncludedLDAPPaths[0]
                $pathParts = $firstPath -split ','
                $identifierPartAttempt = $pathParts[0]
                foreach ($part in $pathParts) {
                    if ($part.StartsWith("OU=", [System.StringComparison]::OrdinalIgnoreCase) -or $part.StartsWith("CN=", [System.StringComparison]::OrdinalIgnoreCase)) {
                        $identifierPartAttempt = $part
                        break
                    }
                }
                $identifierPartClean = $identifierPartAttempt -replace '^(OU=|CN=|DC=)', ''
                $identifierPartClean = $identifierPartClean -replace '[^a-zA-Z0-9_-]', ''
                if ($identifierPartClean.Length -gt 25) {
                    $identifierPartClean = $identifierPartClean.Substring(0, 25)
                }
                if (-not [string]::IsNullOrWhiteSpace($identifierPartClean)) {
                    $logPathIdentifier = $identifierPartClean
                } else {
                    $logPathIdentifier = "SinglePath_UnknownFormat"
                }
            } else {
                $logPathIdentifier = "MultiplePaths_($($IncludedLDAPPaths.Count))"
            }
        }
        # --- END NEW LOGIC FOR LOG FILE IDENTIFIER ---

        # Use script-scoped variable for the detailed log file path
        $Script:MailboxesProcessingLogFile = Join-Path -Path $logPath -ChildPath "MailboxesProcessingDetails_Local_${logPathIdentifier}_$TimeStampForLogOnly.log"

        # Logging function specific to MailboxesProcessing2
        function Write-LogMailboxesProcessing {
            param (
                [Parameter(Mandatory = $true)]
                [string]$Message
            )
            $ProcessingTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $ProcessingLogEntry = "$ProcessingTimestamp - $Message"
            try {
                Add-Content -Path $Script:MailboxesProcessingLogFile -Value $ProcessingLogEntry -ErrorAction Stop
            } catch {
                $mainLogErrorMessage = "ERROR (MailboxesProcessing2): Error writing to MailboxesProcessing log file ($($Script:MailboxesProcessingLogFile)): $($_.Exception.Message). Original Message: $Message"
                WriteLog -message $mainLogErrorMessage
                Write-Host -ForegroundColor Red $mainLogErrorMessage
            }
        }

        # --- START CSV File Overwrite Check (Combined CSV) ---
        # This check is now primarily handled in the main script logic before calling Process-SpecificDomain or MailboxesProcessing
        # However, for the case where !DetectAllDomains and a single OU is provided, the per-domain/path CSV check is relevant here.
        if (-not $DetectAllDomains -and $IncludedLDAPPaths -and $IncludedLDAPPaths.Count -eq 1) {
            $singlePath = $IncludedLDAPPaths[0]
            $singlePathDomainName = "UnknownPathDomain" # Default
            if ($singlePath -match "DC=([^,]+)") {
                $tempDcParts = @()
                $singlePath -split ',' | Where-Object {$_ -like "DC=*"} | ForEach-Object {$tempDcParts += $_ -replace "DC="}
                if ($tempDcParts.Count -gt 0) {
                    $singlePathDomainName = $tempDcParts -join "."
                }
            } else {
                $singlePathDomainName = $singlePath -replace '^(OU=|CN=)','' -replace '[^a-zA-Z0-9.-]','_'
                if ($singlePathDomainName.Length -gt 30) {$singlePathDomainName = $singlePathDomainName.Substring(0,30)}
            }

            $perDomainCsvFileName = "Exchange_OnPrem_Mailboxes_$($singlePathDomainName)$(if($OnlyADPermission){'_OnlyADPermission'}else{''}).csv"
			$perDomainCsvFileFullPath = Join-Path -Path $OutputPath -ChildPath $perDomainCsvFileName

			# Define the base backup directory (for example, a local ExchangeMailboxesInventory\Backup folder).
			$baseExchangeMailboxesInventoryPath = (Get-Item $OutputPath).Parent.FullName # Adjusted to get ExchangeMailboxesInventory base
			$backupBaseDir = Join-Path -Path $baseExchangeMailboxesInventoryPath -ChildPath "Backup"

			if (Test-Path $perDomainCsvFileFullPath) {
				if (-not $ForceOverwriteCSV) {
					# Create a timestamped directory inside the backup base directory
					$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
					$currentBackupDir = Join-Path -Path $backupBaseDir -ChildPath $timestamp

					# Create the full backup path including the timestamped directory
					if (-not (Test-Path $currentBackupDir)) {
						New-Item -ItemType Directory -Path $currentBackupDir -Force | Out-Null
						Write-LogMailboxesProcessing "INFO: Created timestamped backup directory: '$currentBackupDir'."
					}

					# The backup file name should be the same as the source file name
					$backupFileName = Split-Path -Path $perDomainCsvFileFullPath -Leaf
					$backupFilePath = Join-Path -Path $currentBackupDir -ChildPath $backupFileName

					$message = "Per-domain/path CSV file '$perDomainCsvFileFullPath' for '$singlePath' already exists and -ForceOverwriteCSV is `$false. Backing up the existing file to '$backupFilePath'."
					Write-LogMailboxesProcessing "WARNING: $message"
					Write-Host -ForegroundColor Yellow $message

					# Copy the file to the backup directory
					try {
						Copy-Item -Path $perDomainCsvFileFullPath -Destination $backupFilePath -Force -ErrorAction Stop
						Write-LogMailboxesProcessing "INFO: Successfully backed up '$perDomainCsvFileFullPath' to '$backupFilePath'."

						# After backing up, remove the original file
						Remove-Item -Path $perDomainCsvFileFullPath -ErrorAction Stop
						Write-LogMailboxesProcessing "INFO: Successfully removed original file '$perDomainCsvFileFullPath'."
					} catch {
						$errorMessage = "CRITICAL ERROR: Failed to process the per-domain CSV file '$perDomainCsvFileFullPath'. Error: $($_.Exception.Message)"
						Write-LogMailboxesProcessing $errorMessage
						Write-Host -ForegroundColor Red $errorMessage
						return # Stop processing for this domain if copy or removal fails
					}
				} else {
					$message = "Per-domain/path CSV file '$perDomainCsvFileFullPath' for '$singlePath' exists and -ForceOverwriteCSV is `$true. Existing file will be deleted."
					Write-LogMailboxesProcessing "INFO: $message"
					Write-Host -ForegroundColor Cyan $message
					try {
						Remove-Item -Path $perDomainCsvFileFullPath -Force -ErrorAction Stop
						Write-LogMailboxesProcessing "INFO: File '$perDomainCsvFileFullPath' deleted successfully."
					} catch {
						$errorMessage = "ERROR: Could not delete existing per-domain/path CSV file '$perDomainCsvFileFullPath'. Message: $($_.Exception.Message)"
						Write-LogMailboxesProcessing $errorMessage
						Write-Error $errorMessage
					}
				}
			}
        }
        # --- END Per-Path/Domain CSV Overwrite Check ---


        $output = @()
        $Domains = @()
        $MailboxesByDomain = @{}
        $ShowAll = $false
        $ExpandGroups = $true
        $i = 0
        $totalMailBox = 0
        $MailboxScanStartTime = $null

        Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
        Write-LogMailboxesProcessing "MailboxesProcessing2: Starting data collection and export. Log Identifier: $logPathIdentifier"
        if ($IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0) {
            Write-Host "Filtering by provided LDAP paths ($($IncludedLDAPPaths.Count) paths) ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-LogMailboxesProcessing "Filtering by LDAP paths: $($IncludedLDAPPaths -join '; ')"
        }
        Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))

        $AllMailbox = @()
        if ($IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0) {
            Write-Host -ForegroundColor:Cyan "Retrieving mailboxes from specified paths. Please wait, this may take some time... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            $totalOUs = $IncludedLDAPPaths.Count
            $ouCounter = 0
            foreach ($ou in $IncludedLDAPPaths) {
                $ouCounter++
                Write-Host -ForegroundColor:Yellow "Checking existence and retrieving mailboxes from '$ou' (Path $ouCounter of $totalOUs)..."
                Write-LogMailboxesProcessing "Attempting to retrieve mailboxes from path: '$ou' (Path $ouCounter of $totalOUs)"
                try {
                    $adObjectExists = $false
                    if ($ou.StartsWith("OU=", [System.StringComparison]::OrdinalIgnoreCase)) {
                        if (Get-ADOrganizationalUnit -Identity $ou -ErrorAction SilentlyContinue) {
                            $adObjectExists = $true
                        }
                    } elseif ($ou.StartsWith("DC=", [System.StringComparison]::OrdinalIgnoreCase)) {
                        if (Get-ADDomain -Identity ($ou -replace 'DC=','' -replace ',','.') -ErrorAction SilentlyContinue) {
                             $adObjectExists = $true
                        } else {
                             Write-LogMailboxesProcessing "Note: Could not verify domain DN '$ou' with Get-ADDomain. Get-Mailbox will be attempted."
                             $adObjectExists = $true # Assume it might be a valid scope for Get-Mailbox anyway
                        }
                    } else {
                         Write-LogMailboxesProcessing "Path '$ou' does not start with OU= or DC=. Assuming it's a specific DN Get-Mailbox might handle or it's an error. Attempting Get-Mailbox."
                         $adObjectExists = $true # Assume Get-Mailbox might handle it
                    }

                    if ($adObjectExists) {
                        Write-Host -ForegroundColor:Yellow "Retrieving mailboxes from '$ou'..."
                        Write-LogMailboxesProcessing "Retrieving mailboxes from '$ou'."
						if ($LimitResultSize) {
							$mailboxesInPath = Get-Mailbox -OrganizationalUnit $ou -ResultSize $LimitResultSize -ErrorAction Stop
						} else {
							$mailboxesInPath = Get-Mailbox -OrganizationalUnit $ou -ResultSize Unlimited -ErrorAction Stop
						}
                        $AllMailbox += $mailboxesInPath
                        Write-LogMailboxesProcessing "Found $($mailboxesInPath.Count) mailboxes in '$ou'."
                    } else {
                        $warningMessage = "The path '$ou' was not positively identified as an existing OU or resolvable Domain DN. It will be skipped."
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage"
                    }
                }
                catch {
                    $errorMessage = "Error while retrieving mailboxes from path '$ou': $($_.Exception.Message)."
                    Write-Error $errorMessage
                    Write-LogMailboxesProcessing "ERROR: $errorMessage"
                }
            }
        }
        else
        {
            Write-Host -ForegroundColor:Cyan "Retrieving ALL mailboxes (no specific paths provided). Please wait, this may take some time... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-LogMailboxesProcessing "Retrieving ALL mailboxes (no specific paths provided)..."
            try {
				if ($LimitResultSize) {
					$AllMailbox = Get-Mailbox -ResultSize $LimitResultSize -ErrorAction Stop
				} else {
					$AllMailbox = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
				}

            } catch {
                $errorMessage = "Error while retrieving all mailboxes: $($_.Exception.Message)."
                Write-Error $errorMessage
                Write-LogMailboxesProcessing "ERROR: $errorMessage"
            }
        }

        $totalMailBox = $AllMailbox.Count
        Write-Host -ForegroundColor Cyan "Total number of mailboxes to process for current scope: $totalMailBox"
        Write-LogMailboxesProcessing "Total number of mailboxes to process for current scope: $totalMailBox"

        if ($totalMailBox -eq 0) {
            $warningMessage = "No mailboxes were found to process for the current scope. $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-Warning $warningMessage
            Write-LogMailboxesProcessing $warningMessage
        }
        else
        {
            Write-Host -ForegroundColor:Cyan "Processing mailboxes. Please wait, this will likely take some time... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            if ($OnlyADPermission) {
                Write-Host -ForegroundColor:Cyan "OnlyADPermission mode is active!"
                Write-LogMailboxesProcessing "OnlyADPermission mode is active. Standard mailbox details will be skipped."
            }
            Write-LogMailboxesProcessing "Starting detailed processing of $totalMailBox mailboxes for current scope."
            $MailboxScanStartTime = Get-Date

            $LocalAllOutput = @() # Collection for data from this specific MailboxesProcessing2 call

            Foreach ($Mbx in $AllMailbox) {
                $i++
                $Domain = ""
                try {
                    $DNParts = $Mbx.DistinguishedName -split ","
                    $DomainParts = $DNParts | Where-Object { $_ -like "DC=*" } | ForEach-Object { $_ -replace "DC=", "" }
                    $Domain = $DomainParts -join "."
                } catch {
                    $warningMessage = "Could not determine domain for mailbox $($Mbx.Name) from DN: $($Mbx.DistinguishedName)"
                    Write-Warning $warningMessage
                    Write-LogMailboxesProcessing "WARNING: $warningMessage"
                    $Domain = "Unknown"
                }

                Write-LogMailboxesProcessing "Processing mailbox ($i of $totalMailBox): $($Mbx.Name) - ($Domain)"
                $MbxStartTime = Get-Date

                if ($Domains -notcontains $Domain) {
                    $Domains += $Domain
                    $MailboxesByDomain[$Domain] = @()
                }

                $userObj = New-Object PSObject
                $userObj | Add-Member NoteProperty -Name "DomainName" -Value $Domain
                $userObj | Add-Member NoteProperty -Name "Name" -Value $Mbx.Name
                $userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Mbx.DisplayName
                $userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
                $userObj | Add-Member NoteProperty -Name "UserPrincipalName" -Value $Mbx.UserPrincipalName
                $userObj | Add-Member NoteProperty -Name "SamAccountName" -Value $Mbx.SamAccountName

                if ($OnlyADPermission -eq $false)
                {
                    $userObj | Add-Member NoteProperty -Name "IsMailboxEnabled" -Value $Mbx.IsMailboxEnabled
                    $userObj | Add-Member NoteProperty -Name "IsValid" -Value $Mbx.IsValid
                    $userObj | Add-Member NoteProperty -Name "IsShared" -Value $Mbx.IsShared
                    $userObj | Add-Member NoteProperty -Name "IsLinked" -Value $Mbx.IsLinked
                    $userObj | Add-Member NoteProperty -Name "IsResource" -Value $Mbx.IsResource
                    $userObj | Add-Member NoteProperty -Name "AccountDisabled" -Value $Mbx.AccountDisabled
                    $userObj | Add-Member NoteProperty -Name "DistinguishedName" -Value $Mbx.DistinguishedName
                    $userObj | Add-Member NoteProperty -Name "RecipientType" -Value $(if($null -ne $Mbx.RecipientTypeDetails){$Mbx.RecipientTypeDetails.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "RecipientOU" -Value $Mbx.OrganizationalUnit
                    $userObj | Add-Member NoteProperty -Name "PrimarySMTPaddress" -Value $(if($null -ne $Mbx.PrimarySmtpAddress){$Mbx.PrimarySmtpAddress.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "EmailAddresses" -Value (($Mbx.EmailAddresses | Where-Object {$_.PrefixString -eq 'smtp'} | ForEach-Object {$_.SmtpAddress}) -join ";")
                    $userObj | Add-Member NoteProperty -Name "Database" -Value $(if($null -ne $Mbx.Database){$Mbx.Database.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ServerName" -Value $Mbx.ServerName
                    $userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
                    $userObj | Add-Member NoteProperty -Name "ArchiveName" -Value ($Mbx.ArchiveName -join ";")
                    $userObj | Add-Member NoteProperty -Name "ArchiveStatus" -Value $(if($null -ne $Mbx.ArchiveStatus){$Mbx.ArchiveStatus.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ArchiveState" -Value $(if($null -ne $Mbx.ArchiveState){$Mbx.ArchiveState.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ArchiveQuota" -Value $(if($null -ne $Mbx.ArchiveQuota){$Mbx.ArchiveQuota.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ForwardingAddress" -Value $(if($null -ne $Mbx.ForwardingAddress){$Mbx.ForwardingAddress.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ForwardingSmtpAddress" -Value $(if($null -ne $Mbx.ForwardingSmtpAddress){$Mbx.ForwardingSmtpAddress.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "DeliverToMailboxAndForward" -Value $Mbx.DeliverToMailboxAndForward
                    $userObj | Add-Member NoteProperty -Name "Department" -Value $Mbx.Department
                    $userObj | Add-Member NoteProperty -Name "Office" -Value $Mbx.Office
                    $userObj | Add-Member NoteProperty -Name "Manager" -Value $(if($null -ne $Mbx.Manager){$Mbx.Manager.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "WhenMailboxCreated" -Value $Mbx.WhenMailboxCreated
                    $userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $Mbx.WhenChanged
                    $userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $Mbx.WhenCreated
                    $userObj | Add-Member NoteProperty -Name "MailboxLocations" -Value ($Mbx.MailboxLocations -join ";")
                    $userObj | Add-Member NoteProperty -Name "Identity" -Value $(if($null -ne $Mbx.Identity){$Mbx.Identity.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ObjectCategory" -Value $(if($null -ne $Mbx.ObjectCategory){$Mbx.ObjectCategory.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "GrantSendOnBehalfTo" -Value (($Mbx.GrantSendOnBehalfTo | ForEach-Object {if($null -ne $_){$_.ToString()}else{""}}) -join ";")

                    $ProhibitSendReceiveQuota = "N/A"
                    try {
                        if (($Mbx.UseDatabaseQuotaDefaults -eq $true)) {
                            $db = Get-MailboxDatabase $Mbx.Database -ErrorAction Stop
                            if ($db.ProhibitSendReceiveQuota.IsUnlimited) { $ProhibitSendReceiveQuota = "Unlimited" }
                            else { $ProhibitSendReceiveQuota = $db.ProhibitSendReceiveQuota.Value.ToMB() }
                        } elseif ($Mbx.ProhibitSendReceiveQuota.IsUnlimited) {
                            $ProhibitSendReceiveQuota = "Unlimited"
                        } else {
                            $ProhibitSendReceiveQuota = $Mbx.ProhibitSendReceiveQuota.Value.ToMB()
                        }
                    } catch {
                        $warningMessage = "Could not determine ProhibitSendReceiveQuota for $($Mbx.Name): $($_.Exception.Message)"
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage"
                        $ProhibitSendReceiveQuota = "Error"
                    }
                    $userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $ProhibitSendReceiveQuota

                    # Initialize Mailbox Statistics properties
                    $userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value "N/A"
                    $userObj | Add-Member NoteProperty -Name "TotalItemSize-In-MB" -Value "N/A"
                    $userObj | Add-Member NoteProperty -Name "ItemCount" -Value "N/A"
                    $userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value "N/A"
                    $userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize-In-MB" -Value "N/A"

                    $MbxStatsLogStartTime = Get-Date
                    try {
                        $Stats = Get-MailboxStatistics -Identity $Mbx.DistinguishedName -WarningAction SilentlyContinue -ErrorAction Stop
                        if ($Stats) {
                            $userObj."LastLogonTime" = $Stats.LastLogonTime
							$userObj."TotalItemSize-In-MB" = $(if($null -ne $Stats.TotalItemSize -and $Stats.TotalItemSize.IsUnlimited -eq $false) {$Stats.TotalItemSize.Value.ToMB()} elseif($null -ne $Stats.TotalItemSize -and $Stats.TotalItemSize.IsUnlimited -eq $true) {$Stats.TotalItemSize.Value.ToMB()} else {"N/A"})
                            $userObj."ItemCount" = $Stats.ItemCount
                            $userObj."DeletedItemCount" = $Stats.DeletedItemCount
                            $userObj."TotalDeletedItemSize-In-MB" = $(if($null -ne $Stats.TotalDeletedItemSize -and $Stats.TotalDeletedItemSize.IsUnlimited -eq $false) {$Stats.TotalDeletedItemSize.Value.ToMB()} elseif($null -ne $Stats.TotalDeletedItemSize -and $Stats.TotalDeletedItemSize.IsUnlimited -eq $true) {$Stats.TotalDeletedItemSize.Value.ToMB()} else {"N/A"})
                        }
                    }
                    catch {
                        $warningMessage = "An error occurred while retrieving mailbox statistics for $($Mbx.DistinguishedName)."
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage Error: $($_.Exception.Message)"
                        $userObj."LastLogonTime" = "Error"
                        $userObj."TotalItemSize-In-MB" = "Error"
                        $userObj."ItemCount" = "Error"
                        $userObj."DeletedItemCount" = "Error"
                        $userObj."TotalDeletedItemSize-In-MB" = "Error"
                    }

                    $ArchiveTotalItemSizeMB = "N/A"
                    $ArchiveTotalItemCount = "N/A"
                    if ($Mbx.ArchiveGuid -ne [System.Guid]::Empty) {
                        try {
                            $MbxArchiveStats = Get-MailboxStatistics -Identity $Mbx.DistinguishedName -Archive -WarningAction SilentlyContinue -ErrorAction Stop
                            if ($MbxArchiveStats) {
                                $ArchiveTotalItemSizeMB = $(if($null -ne $MbxArchiveStats.TotalItemSize -and $MbxArchiveStats.TotalItemSize.IsUnlimited -eq $false) {$MbxArchiveStats.TotalItemSize.Value.ToMB()} elseif ($null -ne $MbxArchiveStats.TotalItemSize -and $MbxArchiveStats.TotalItemSize.IsUnlimited -eq $true) {"Unlimited"} else {"N/A"})
                                $ArchiveTotalItemCount = $MbxArchiveStats.ItemCount
                            }
                        }
                        catch {
                            $warningMessage = "An error occurred while retrieving archive mailbox statistics for $($Mbx.DistinguishedName)."
                            Write-Warning $warningMessage
                            Write-LogMailboxesProcessing "WARNING: $warningMessage Error: $($_.Exception.Message)"
                            $ArchiveTotalItemSizeMB = "Error"
                            $ArchiveTotalItemCount = "Error"
                        }
                    }
                    $userObj | Add-Member NoteProperty -Name "ArchiveTotalItemSize-In-MB" -Value $ArchiveTotalItemSizeMB
                    $userObj | Add-Member NoteProperty -Name "ArchiveItemCount" -Value $ArchiveTotalItemCount

                    $LargeItemCount = "N/A" # Logic for this is not implemented
                    $LargeItemThresholdMBValue = 35 # Example threshold
                    $userObj | Add-Member NoteProperty -Name "LargeItemCount-Over-$($LargeItemThresholdMBValue)MB" -Value $LargeItemCount

                    $MbxStatsLogEndTime = Get-Date
                    Write-LogMailboxesProcessing " -------- Processing time for MailboxStatistics for $($Mbx.Name): $($MbxStatsLogEndTime - $MbxStatsLogStartTime)"

                    $UserExceptions = @()
                    $AdditionalUserFilters = ""
                    if (-not $ShowAll) {
                        $UserExceptions = @(
                            'S-1-*', # Well-known SIDs often cause issues or are irrelevant
                            "*\Organization Management", "*\Domain Admins", "*\Enterprise Admins",
                            "*\Exchange Services", "*\Exchange Trusted Subsystem", "*\Exchange Servers",
                            "*\Exchange View-Only Administrators", "*\Exchange Admins",
                            "*\Managed Availability Servers", "*\Public Folder Administrators",
                            "*\Exchange Domain Servers", "*\Exchange Organization Administrators",
                            "NT AUTHORITY\*"
                        )
                    }

                    $ExceptionsRegex = '^(?!)$' # Regex that matches nothing
                    if ($UserExceptions.Count -gt 0) {
                        $ExceptionMatches = @($UserExceptions | ForEach-Object {[System.Text.RegularExpressions.Regex]::Escape($_)})
                        $ExceptionsRegex = '^(' + ($ExceptionMatches -join '|') + ')$'
                        $ExceptionsRegex = $ExceptionsRegex -replace '\\\*','.*' # Convert wildcard * to regex .*
                    }

                    $MbxPermsLogStartTime = Get-Date
                    $FullAccessUsersList = @()
                    $FullAccessUserCount = 0
                    try {
                        $fullPermsRaw = Get-MailboxPermission -Identity $Mbx.DistinguishedName -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'FullAccess' -and (-not $_.IsInherited) -and (-not $_.Deny) }
                        if ($fullPermsRaw) {
                            foreach ($permEntry in $fullPermsRaw) {
                                $UserIDString = if ($null -ne $permEntry.User) {$permEntry.User.ToString()} else {"<UnknownUserSID: $($permEntry.User.SecurityIdentifier.Value)>"}
                                $UserSID = $permEntry.User.SecurityIdentifier.Value
                                $adObject = $null
                                $resolvedBy = "N/A"
                                $formattedADObjectIdentity = $UserIDString # Default if not resolved

                                if ($UserIDString -notmatch $ExceptionsRegex) {
                                    if ($ExpandGroups) {
                                        if ($Global:ADObjectCache.ContainsKey($UserSID)) {
                                            $adObject = $Global:ADObjectCache[$UserSID]
                                            if ($adObject) {
                                                $resolvedBy = "Cache (Object Found)"
                                                Write-LogMailboxesProcessing "        Found SID $UserSID ($UserIDString) in ADObjectCache."
                                            } else {
                                                $resolvedBy = "Cache (Not Found Marker)"
                                                Write-LogMailboxesProcessing "        Found SID $UserSID ($UserIDString) in ADObjectCache as 'Not Found'. Skipping AD lookups."
                                            }
                                        } else { # Not in cache, try to resolve
                                            # Attempt 1: If UserIDString is DOMAIN\User format (and not a SID string)
                                            if ($UserIDString -match '.+\\.+' -and -not ($UserIDString -like "S-1-*")) {
                                                $NetBIOSDomainName, $sAMAccountNameToFind = $UserIDString.Split('\')
                                                Write-LogMailboxesProcessing "        Attempting to resolve '$UserIDString' via NetBIOS domain '$NetBIOSDomainName' and sAMAccountName '$sAMAccountNameToFind'."
                                                $TargetDomainDNSRoot = $null
                                                if ($Global:DomainInfoCache.ContainsKey($NetBIOSDomainName)) {
                                                    $TargetDomainDNSRoot = $Global:DomainInfoCache[$NetBIOSDomainName]
                                                    Write-LogMailboxesProcessing "          Found '$NetBIOSDomainName' in DomainInfoCache. DNSRoot: $TargetDomainDNSRoot"
                                                } else {
                                                    try {
                                                        $TargetDomainObj = Get-ADDomain -Identity $NetBIOSDomainName -ErrorAction Stop
                                                        if ($TargetDomainObj) {
                                                            $TargetDomainDNSRoot = $TargetDomainObj.DNSRoot
                                                            $Global:DomainInfoCache[$NetBIOSDomainName] = $TargetDomainDNSRoot
                                                            Write-LogMailboxesProcessing "          Retrieved and cached DNSRoot for '$NetBIOSDomainName': $TargetDomainDNSRoot"
                                                        }
                                                    } catch { Write-LogMailboxesProcessing "          Error retrieving domain info for NetBIOS '$NetBIOSDomainName': $($_.Exception.Message)" }
                                                }

                                                if ($TargetDomainDNSRoot) {
                                                    try {
                                                        $adObject = Get-ADObject -Filter "SamAccountName -eq '$sAMAccountNameToFind'" -Server $TargetDomainDNSRoot -Properties ObjectClass, SamAccountName, DistinguishedName -ErrorAction Stop
                                                        if ($adObject) { $resolvedBy = "TargetedDomain ($TargetDomainDNSRoot)" }
                                                    } catch { Write-LogMailboxesProcessing "          Error resolving '$UserIDString' via targeted domain search on '$TargetDomainDNSRoot': $($_.Exception.Message)" }
                                                }
                                            }

                                            # Attempt 2: Resolve by SID via GC (if not resolved yet)
                                            if (-not $adObject) {
                                                $gcServer = $null
                                                try {
                                                    $currentForest = Get-ADForest -ErrorAction SilentlyContinue
                                                    if ($currentForest) { $gcServer = $currentForest.GlobalCatalogs | Get-Random -ErrorAction SilentlyContinue }
                                                } catch { Write-LogMailboxesProcessing "          Warning: Could not get Global Catalog server list for SID $UserSID. Error: $($_.Exception.Message)" }

                                                if ($gcServer) {
                                                    Write-LogMailboxesProcessing "        Attempting to resolve SID $UserSID for '$UserIDString' via GC: $gcServer"
                                                    try {
                                                        $adObject = Get-ADObject -Identity $UserSID -Server $gcServer -Properties ObjectClass, SamAccountName, DistinguishedName -ErrorAction Stop
                                                        if ($adObject) { $resolvedBy = "GC ($gcServer)" }
                                                    } catch { Write-LogMailboxesProcessing "          GC lookup for SID $UserSID failed or object not found. Error: $($_.Exception.Message)" }
                                                } else { Write-LogMailboxesProcessing "          Skipping GC lookup for SID $UserSID as no GC server was found/available."}
                                            }

                                            # Attempt 3: Resolve by SID via default domain (if still not resolved)
                                            if (-not $adObject) {
                                                Write-LogMailboxesProcessing "        Attempting default domain search for SID $UserSID for '$UserIDString'."
                                                try {
                                                    $adObject = Get-ADObject -Identity $UserSID -Properties ObjectClass, SamAccountName, DistinguishedName -ErrorAction Stop
                                                    if ($adObject) { $resolvedBy = "DefaultDomain" }
                                                } catch { Write-LogMailboxesProcessing "          Default domain search for SID $UserSID failed or object not found. Error: $($_.Exception.Message)" }
                                            }
                                            $Global:ADObjectCache[$UserSID] = $adObject # Cache the result (object or $null)
                                        }

                                        if ($adObject) { # If object was resolved (from cache or new lookup)
                                            # Get NetBIOS domain name for the resolved object for consistent formatting
                                            $objectDomainDistinguishedName = ($adObject.DistinguishedName -replace "^.+?(?=DC=)","") # Extract domain part of DN
                                            $objectDomainFQDN = ($objectDomainDistinguishedName -replace 'DC=','' -replace ',','.')
                                            $objectNetBIOSDomain = $null

                                            if ($Global:DomainFQDNToNetBIOSCache.ContainsKey($objectDomainFQDN)) {
                                                $objectNetBIOSDomain = $Global:DomainFQDNToNetBIOSCache[$objectDomainFQDN]
                                            } else {
                                                try {
                                                    $adDomainForObject = Get-ADDomain -Identity $objectDomainFQDN -ErrorAction Stop # Query domain by FQDN
                                                    if ($adDomainForObject) {
                                                        $objectNetBIOSDomain = $adDomainForObject.NetBIOSName
                                                        $Global:DomainFQDNToNetBIOSCache[$objectDomainFQDN] = $objectNetBIOSDomain
                                                    } else { Write-LogMailboxesProcessing "          Warning: Get-ADDomain returned null for FQDN '$objectDomainFQDN' (object $($adObject.SamAccountName))" }
                                                } catch { Write-LogMailboxesProcessing "          Warning: Could not get NetBIOS for domain FQDN '$objectDomainFQDN' of object '$($adObject.SamAccountName)': $($_.Exception.Message)" }
                                            }
                                            $formattedADObjectIdentity = if ($objectNetBIOSDomain) { "$objectNetBIOSDomain\$($adObject.SamAccountName)" } else { $adObject.SamAccountName }

                                            if ($resolvedBy -ne "Cache (Not Found Marker)") { # Avoid logging for already known 'not found'
                                                 Write-LogMailboxesProcessing "        Successfully resolved SID $UserSID ($UserIDString) to '$formattedADObjectIdentity' (DN: '$($adObject.DistinguishedName)') via $resolvedBy."
                                            }

                                            if ($adObject.ObjectClass -contains 'group') {
                                                Write-LogMailboxesProcessing "        Expanding group $formattedADObjectIdentity (DN: $($adObject.DistinguishedName)) for FullAccess on $($Mbx.Name)"
                                                $groupMembersSam = @()
                                                $groupExpansionStatusString = "$formattedADObjectIdentity (Group - "

                                                if ($Global:GroupMemberCache.ContainsKey($adObject.DistinguishedName)) {
                                                    $cachedEntry = $Global:GroupMemberCache[$adObject.DistinguishedName]
                                                    if ($cachedEntry -is [array]) { # Successfully expanded before
                                                        $groupMembersSam = $cachedEntry
                                                        Write-LogMailboxesProcessing "          Found group '$($adObject.SamAccountName)' in GroupMemberCache. Members count: $($groupMembersSam.Count)"
                                                    } elseif ($cachedEntry -is [string] -and ($cachedEntry -like "*Error Expanding Members*" -or $cachedEntry -like "*No Members Listed*")) {
                                                        $FullAccessUsersList += $cachedEntry # Add error/status from cache
                                                    }
                                                } else { # Not in cache, try to expand
                                                    $groupDomainForExpansionFQDN = $objectDomainFQDN # Use the group's own domain FQDN
                                                    try {
                                                        $groupMembersADPrincipal = $null
                                                        if ($groupDomainForExpansionFQDN) {
                                                            Write-LogMailboxesProcessing "          Targeting domain '$groupDomainForExpansionFQDN' for Get-ADGroupMember for group '$($adObject.DistinguishedName)'"
                                                            $groupMembersADPrincipal = Get-ADGroupMember -Identity $adObject.DistinguishedName -Server $groupDomainForExpansionFQDN -Recursive -ErrorAction Stop
                                                        } else {
                                                            Write-LogMailboxesProcessing "          Could not determine specific domain FQDN for group '$($adObject.DistinguishedName)'. Attempting Get-ADGroupMember using DN without specifying server."
                                                            $groupMembersADPrincipal = Get-ADGroupMember -Identity $adObject.DistinguishedName -Recursive -ErrorAction Stop
                                                        }

                                                        if ($groupMembersADPrincipal) {
                                                            $groupMembersSam = $groupMembersADPrincipal | ForEach-Object {$_.SamAccountName}
                                                            $Global:GroupMemberCache[$adObject.DistinguishedName] = $groupMembersSam # Cache successful expansion
                                                        } else { # Group has no members
                                                            $statusMsg = "$groupExpansionStatusString" + "No Members Listed)"
                                                            $FullAccessUsersList += $statusMsg
                                                            $Global:GroupMemberCache[$adObject.DistinguishedName] = $statusMsg # Cache no members status
                                                        }
                                                    } catch {
                                                        $statusMsg = "$groupExpansionStatusString" + "Error Expanding Members: $($_.Exception.Message.Split([Environment]::NewLine)[0]))"
                                                        Write-LogMailboxesProcessing "          $statusMsg"
                                                        $FullAccessUsersList += $statusMsg
                                                        $Global:GroupMemberCache[$adObject.DistinguishedName] = $statusMsg # Cache error status
                                                    }
                                                }

                                                if ($groupMembersSam.Count -gt 0) {
                                                    foreach($sam in $groupMembersSam){
                                                        $memberNetBIOSDomain = $objectNetBIOSDomain # Assume members are in the same domain as the group for formatting
                                                        $memberFormattedIdentity = if ($memberNetBIOSDomain) { "$memberNetBIOSDomain\$sam" } else { $sam }
                                                        if ($memberFormattedIdentity -notmatch $ExceptionsRegex) { # Check exception for each member
                                                            $FullAccessUsersList += $memberFormattedIdentity
                                                        }
                                                    }
                                                }
                                            } else { # Not a group, just a user
                                                $FullAccessUsersList += $formattedADObjectIdentity
                                            }
                                        } else { # adObject is $null (could not be resolved)
                                            if ($resolvedBy -ne "Cache (Not Found Marker)") { # Avoid logging for already known 'not found'
                                                Write-LogMailboxesProcessing "        SID $UserSID ($UserIDString) could not be resolved to an AD object after all attempts."
                                            }
                                            $FullAccessUsersList += "$UserIDString (SID Not Found in AD)"
                                        }
                                    } else { # ExpandGroups is $false
                                        $FullAccessUsersList += $UserIDString
                                    }
                                } # End if notmatch $ExceptionsRegex
                            }
                            $FullAccessUsersList = $FullAccessUsersList | Select-Object -Unique
                            $FullAccessUserCount = $FullAccessUsersList.Count
                        }
                        $userObj | Add-Member NoteProperty -Name "FullAccessCount" -Value $FullAccessUserCount
                        $userObj | Add-Member NoteProperty -Name "FullAccess" -Value ($FullAccessUsersList -join ";")
                    } catch {
                        $userObj | Add-Member NoteProperty -Name "FullAccessCount" -Value "Error"
                        $userObj | Add-Member NoteProperty -Name "FullAccess" -Value "Error retrieving permissions"
                        $warningMessage = "An error occurred while retrieving MailboxPermission for mailbox $($Mbx.DistinguishedName): $($_.Exception.Message)"
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage"
                        Write-LogMailboxesProcessing "        Outer catch for MailboxPermission processing for '$UserIDString' on $($Mbx.Name): $($_.Exception.Message)"
                    }
                    $MbxPermsLogEndTime = Get-Date
                    Write-LogMailboxesProcessing " -------- Processing time for MailboxPermission for $($Mbx.Name): $($MbxPermsLogEndTime - $MbxPermsLogStartTime)"
                } # End if ($OnlyADPermission -eq $false)
				else
				{
					$userObj | Add-Member NoteProperty -Name "PrimarySMTPaddress" -Value $(if($null -ne $Mbx.PrimarySmtpAddress){$Mbx.PrimarySmtpAddress.ToString()}else{""})
				}

                If ($IncludeADPermission)
                {
					try {
						$MbxADPermsLogStartTime = Get-Date   # log start time

						# Make sure these are reset for each mailbox
						$SendAsUsersList  = @()
						$SendAsUserCount  = 0

						$sendAsPermissions = Get-ADPermission -Identity $Mbx.DistinguishedName | Where-Object {
							($_.ExtendedRights -like "*Send*") -and
							($_.IsInherited -eq $false) -and
							($_.User -notlike "NT AUTHORITY\SELF") -and
							($_.User -notmatch '^S-1-5-21-')   # Ignore unresolved SIDs
						}

						if ($sendAsPermissions) {
							Write-Host "`nMailbox: $($Mbx.DisplayName)"
							foreach ($perm in $sendAsPermissions) {
								$SendAsUsersList += $perm.User
								Write-Host "  SendAs granted to: $($perm.User)"
							}
							$SendAsUsersList = $SendAsUsersList | Select-Object -Unique
							$SendAsUserCount = $SendAsUsersList.Count
						}

						$userObj | Add-Member NoteProperty -Name "SendAsCount" -Value $SendAsUserCount -Force
						$userObj | Add-Member NoteProperty -Name "SendAs"      -Value ($SendAsUsersList -join ";") -Force
					}
					catch {
						$userObj | Add-Member NoteProperty -Name "SendAsCount" -Value "Error" -Force
						$userObj | Add-Member NoteProperty -Name "SendAs"      -Value "Error retrieving SendAs" -Force
						$warningMessage = "An error occurred while retrieving ADPermission (SendAs) for mailbox $($Mbx.DistinguishedName): $($_.Exception.Message)"
						Write-Warning $warningMessage
						Write-LogMailboxesProcessing "WARNING: $warningMessage"
					}
                    $MbxADPermsLogEndTime = Get-Date
                    Write-LogMailboxesProcessing " -------- Processing time for ADPermission (SendAs) for $($Mbx.Name): $($MbxADPermsLogEndTime - $MbxADPermsLogStartTime)"
                }
                else
                {
                    $userObj | Add-Member NoteProperty -Name "SendAsCount" -Value "NotChecked"
                    $userObj | Add-Member NoteProperty -Name "SendAs" -Value "NotChecked"
                }

                if ($OnlyADPermission -eq $false)
                {
                    $MbxMobileLogStartTime = Get-Date
                    $MobileDevicePresent = $false
                    $ClientTypes = @()
                    $FriendlyNames = @()
                    $DeviceTypes = @()
                    $LastSyncTimes = @()
                    $MobileDeviceError = ""

                    try {
                        $MobileDeviceIdentity = $Mbx.UserPrincipalName
                        if (-not $MobileDeviceIdentity -and $Mbx.SamAccountName) {
                            $MobileDeviceIdentity = $Mbx.SamAccountName
                        } elseif (-not $MobileDeviceIdentity) {
                            $MobileDeviceIdentity = $Mbx.DistinguishedName # Fallback, though UPN/SAM is preferred for Get-MobileDevice
                        }

                        if ($MobileDeviceIdentity) {
                            $MobileDevices = Get-MobileDevice -Mailbox $MobileDeviceIdentity -ErrorAction SilentlyContinue
                            if ($MobileDevices) {
                                 $MobileDevicePresent = $true
                                 $ClientTypes = $MobileDevices.ClientType | ForEach-Object {if($null -ne $_){$_.ToString()}else{""}}
                                 $FriendlyNames = $MobileDevices.FriendlyName
                                 $DeviceTypes = $MobileDevices.DeviceType
                                 $LastSyncTimes = $MobileDevices.LastSyncTime | ForEach-Object {if($null -ne $_){$_}else{""}} # Keep as datetime objects for now, will be stringified by join
                            }
                        } else {
                            $warningMessage = "Could not determine a suitable identity for Get-MobileDevice for $($Mbx.Name)."
                            Write-Warning $warningMessage
                            Write-LogMailboxesProcessing "WARNING: $warningMessage"
                            $MobileDeviceError = "No Identity"
                        }
                    } catch {
                        $warningMessage = "An error occurred while retrieving MobileDevice for mailbox $($Mbx.DistinguishedName): $($_.Exception.Message)"
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage"
                        $MobileDeviceError = "Error"
                    }
                    $userObj | Add-Member NoteProperty -Name "MobileDeviceAssociated" -Value $(if($MobileDeviceError){$MobileDeviceError} else {$MobileDevicePresent})
                    $userObj | Add-Member NoteProperty -Name "MobileClientTypes" -Value ($ClientTypes -join ";")
                    $userObj | Add-Member NoteProperty -Name "MobileFriendlyNames" -Value ($FriendlyNames -join ";")
                    $userObj | Add-Member NoteProperty -Name "MobileDeviceTypes" -Value ($DeviceTypes -join ";")
                    $userObj | Add-Member NoteProperty -Name "MobileLastSyncTimes" -Value ($LastSyncTimes -join ";")

					$userObj | Add-Member NoteProperty -Name "RemoteRoutingAddress" -Value $Mbx.RemoteRoutingAddress

                    $MbxMobileLogEndTime = Get-Date
                    Write-LogMailboxesProcessing " -------- Processing time for MobileDevice for $($Mbx.Name): $($MbxMobileLogEndTime - $MbxMobileLogStartTime)"
                }

                # ObjectGUID from AD (via Exchange Mailbox object property - no extra AD query needed)
                $userObj | Add-Member NoteProperty -Name "ObjectGUID" -Value $(if ($null -ne $Mbx.Guid) { $Mbx.Guid.ToString() } else { "" })

                # Add to the local collection for this function's scope
                $output += $userObj

                # Add to domain-specific collection (for per-domain CSVs if !DetectAllDomains and multiple paths)
                $MailboxesByDomain[$Domain] += $userObj

                $MbxEndTime = Get-Date
                $MbxTotalTimeTaken = $MbxEndTime - $MbxStartTime
                Write-LogMailboxesProcessing " ------ Total processing time for mailbox $($Mbx.Name): $($MbxTotalTimeTaken.ToString())"

                if ($totalMailBox -gt 0) {
                    $TimeElapsed = (Get-Date) - $MailboxScanStartTime
                    $AvgTimePerMailbox = if ($i -gt 0) { $TimeElapsed.TotalSeconds / $i } else { 0 }
                    $RemainingMailboxes = $totalMailBox - $i
                    $EstimatedTimeRemainingSeconds = if ($AvgTimePerMailbox -gt 0) { $RemainingMailboxes * $AvgTimePerMailbox } else { 0 }
                    $TimeSpan = [TimeSpan]::FromSeconds($EstimatedTimeRemainingSeconds)
                    $EstimatedTimeRemaining = "{0:d2}:{1:d2}:{2:d2}" -f $TimeSpan.Hours, $TimeSpan.Minutes, $TimeSpan.Seconds
                    Write-Progress -Activity "Processing mailbox $($Mbx.Name) - ($($Domain))" -Status "Scanned: $i of $totalMailBox ($([int](($i / $totalMailBox) * 100))%). Estimated time remaining: $EstimatedTimeRemaining" -PercentComplete (($i / $totalMailBox) * 100)
                }
            }
            Write-Progress -Activity "Mailbox Processing" -Status "Completed." -Completed
            Write-LogMailboxesProcessing "END - Detailed processing of mailboxes."
            Write-Host "`nEND - Detailed processing of mailboxes ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-LogMailboxesProcessing "DEBUG: Entering per-path/domain CSV export logic."
            Write-LogMailboxesProcessing "DEBUG: IncludedLDAPPaths = $($IncludedLDAPPaths)"
            Write-LogMailboxesProcessing "DEBUG: IncludedLDAPPaths.Count = $($IncludedLDAPPaths.Count)"
            Write-LogMailboxesProcessing "DEBUG: DetectAllDomains = $($DetectAllDomains)"

            if ((-not $DetectAllDomains -and $IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0)) {
                Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
                Write-LogMailboxesProcessing "Starting export of per-path/per-OU data to CSVs (as -DetectAllDomains is false)."
                Write-Host "Exporting per-path/per-OU data to CSVs ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"

                if ($IncludedLDAPPaths.Count -eq 1) { # Single OU/Path
                    $singlePathForCsv = $IncludedLDAPPaths[0]
                    $csvIdentifierForSinglePath = $logPathIdentifier # Use the already generated logPathIdentifier
                    $OutputDataForPath = $output # All output from this run corresponds to this single path

                    if ($OutputDataForPath -and $OutputDataForPath.Count -gt 0) {
                        $FileNameSuffix = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
        $PerPathCsvFile = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_$($csvIdentifierForSinglePath)${FileNameSuffix}"
                        try {
                            Export-CsvAtomic -InputObject $OutputDataForPath -Path $PerPathCsvFile -Encoding UTF8
                            Write-Host -ForegroundColor:Green "Mailbox data for path '$($singlePathForCsv)' (identified as '$csvIdentifierForSinglePath') exported to: $PerPathCsvFile"
                            Write-LogMailboxesProcessing "Mailbox data for path '$($singlePathForCsv)' exported to: $PerPathCsvFile"
                        } catch {
                            $errorMessage = "Failed to export CSV for path '$($singlePathForCsv)': $($_.Exception.Message)"
                            Write-Error $errorMessage
                            Write-LogMailboxesProcessing "ERROR: $errorMessage"
                        }
                    } else {
                        Write-LogMailboxesProcessing "No data to export for path '$singlePathForCsv'."
                    }
                } elseif ($IncludedLDAPPaths.Count -gt 1) { # Multiple OUs/Paths
                    foreach ($DomainToExportInOUContext in $Domains) {
                        $OutputDataForDomainInOU = $MailboxesByDomain[$DomainToExportInOUContext]
                        if ($OutputDataForDomainInOU -and $OutputDataForDomainInOU.Count -gt 0) {
                            $FileNameSuffix = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
                $csvBaseNameForOUContext = "Exchange_OnPrem_Mailboxes_TargetedScope_Domain_$($DomainToExportInOUContext)"
                            $PerDomainCsvFileInOUContext = Join-Path -Path $OutputPath -ChildPath "${csvBaseNameForOUContext}${FileNameSuffix}"

                            if ($ForceOverwriteCSV -and (Test-Path $PerDomainCsvFileInOUContext)) {
                                Write-LogMailboxesProcessing "INFO: Per-scope/domain CSV file '$PerDomainCsvFileInOUContext' exists and ForceOverwriteCSV is enabled. Deleting file."
                                try { Remove-Item -Path $PerDomainCsvFileInOUContext -Force -ErrorAction Stop }
                                catch { Write-LogMailboxesProcessing "ERROR: Could not delete existing per-scope/domain CSV file '$PerDomainCsvFileInOUContext'. Message: $($_.Exception.Message)" }
                            } elseif ((Test-Path $PerDomainCsvFileInOUContext) -and (-not $ForceOverwriteCSV)) {
                                Write-LogMailboxesProcessing "WARNING: CSV file '$PerDomainCsvFileInOUContext' exists and -ForceOverwriteCSV is false. Skipping export for this specific domain context within targeted OUs."
                                Write-Host -ForegroundColor Yellow "CSV file '$PerDomainCsvFileInOUContext' exists and -ForceOverwriteCSV is false. Skipping export for this domain context."
                                continue # Skip to next domain context
                            }

                            try {
                                Export-CsvAtomic -InputObject $OutputDataForDomainInOU -Path $PerDomainCsvFileInOUContext -Encoding UTF8
                                Write-Host -ForegroundColor:Green "Mailbox data for domain scope '$($DomainToExportInOUContext)' (from targeted OUs) exported to: $PerDomainCsvFileInOUContext"
                                Write-LogMailboxesProcessing "Mailbox data for domain scope '$($DomainToExportInOUContext)' (from targeted OUs) exported to: $PerDomainCsvFileInOUContext"
                            } catch {
                                $errorMessage = "Failed to export CSV for domain scope '$($DomainToExportInOUContext)' (from targeted OUs): $($_.Exception.Message)"
                                Write-Error $errorMessage
                                Write-LogMailboxesProcessing "ERROR: $errorMessage"
                            }
                        } else {
                            Write-LogMailboxesProcessing "No data to export for domain scope '$DomainToExportInOUContext' (from targeted OUs)."
                        }
                    }
                }
                Write-Host "END - Exporting per-path/per-OU data to CSVs ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
                Write-LogMailboxesProcessing "END - Exporting per-path/per-OU data to CSVs."
            }
        } # End else (if $totalMailbox -gt 0)

        Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
        $FunctionEndTime = Get-Date
        $FunctionTotalTimeTaken = $null

        if ($MailboxScanStartTime) { # If scan actually started
            $FunctionTotalTimeTaken = $FunctionEndTime - $MailboxScanStartTime
            if ($totalMailBox -gt 0) {
                $TotalSeconds = $FunctionTotalTimeTaken.TotalSeconds
                $AverageSecondsPerMailbox = $TotalSeconds / $totalMailBox
                $AverageTimePerMailbox = [TimeSpan]::FromSeconds($AverageSecondsPerMailbox)
                Write-Host "MailboxesProcessing2: Total processing time for $totalMailBox mailboxes: $($FunctionTotalTimeTaken.ToString())"
                Write-Host "MailboxesProcessing2: Average processing time per mailbox: $($AverageTimePerMailbox.ToString())"
                Write-LogMailboxesProcessing "MailboxesProcessing2: Total processing time for $totalMailBox mailboxes: $($FunctionTotalTimeTaken.ToString())"
                Write-LogMailboxesProcessing "MailboxesProcessing2: Average processing time per mailbox: $($AverageTimePerMailbox.ToString())"
            } else { # Scan started but no mailboxes found
                Write-LogMailboxesProcessing "MailboxesProcessing2: No mailboxes processed, cannot calculate average time."
                Write-Host "MailboxesProcessing2: Total execution time (scan started, no mailboxes processed): $($FunctionTotalTimeTaken.ToString())"
                Write-LogMailboxesProcessing "MailboxesProcessing2: Total execution time (scan started, no mailboxes processed): $($FunctionTotalTimeTaken.ToString())"
            }
        } else { # Mailbox scan phase was skipped (e.g., no mailboxes initially found)
             $ApproximateStartTime = $FunctionStartTime # Use function start time as best guess
             if (Test-Path $Script:MailboxesProcessingLogFile) { # Use script-scoped variable
                 try {
                     $ApproximateStartTime = (Get-Item $Script:MailboxesProcessingLogFile -ErrorAction Stop).CreationTime
                 } catch {
                     Write-LogMailboxesProcessing "Warning: Could not get CreationTime of MailboxesProcessingLogFile. Using function start time for duration calculation."
                 }
             }
             $FunctionTotalTimeTaken = $FunctionEndTime - $ApproximateStartTime
             Write-Host "MailboxesProcessing2: Total execution time (no mailbox scan phase or scan start time unavailable): $($FunctionTotalTimeTaken.ToString())"
             Write-LogMailboxesProcessing "MailboxesProcessing2: Total execution time (no mailbox scan phase or scan start time unavailable): $($FunctionTotalTimeTaken.ToString())"
        }
        return $output # Return the collected data for this scope
    } # End of MailboxesProcessing2 function

    # Wrapper function that calls MailboxesProcessing2
    function MailboxesProcessing {
        param (
            [Parameter(Mandatory=$true)]
            [string[]]$IncludedLDAPPaths
        )
        process {
            Write-Host "`n  [MailboxesProcessing] Function called." -ForegroundColor Magenta
            WriteLog -Message "  [MailboxesProcessing] Wrapper function called. Will call MailboxesProcessing2."
            WriteLog -Message "  [MailboxesProcessing] Effective permission flags for this call: IncludeADPermission = $IncludeADPermission, OnlyADPermission = $OnlyADPermission"

            # MailboxesProcessing2 will return its collected data
            $processedData = MailboxesProcessing2 -IncludedLDAPPaths $IncludedLDAPPaths

            WriteLog -Message "  [MailboxesProcessing] Finished call to MailboxesProcessing2."
            return $processedData # Pass the data up
        }
    }

    # Function to process a specific domain when DetectAllDomains is $true
    function Process-SpecificDomain {
        param (
            [Parameter(Mandatory=$true)]
            [System.DirectoryServices.ActiveDirectory.Domain]$CurrentDomain
        )
        process {
            $domainName = $CurrentDomain.Name
            $distinguishedName = ($domainName.Split('.') | ForEach-Object { "DC=$_" }) -join ','

            WriteLog -Message "Starting global processing for domain: '$domainName' (DN: '$distinguishedName')"

            $isForestRootDomain = $false
            if ($CurrentDomain.Forest -and $CurrentDomain.Forest.RootDomain -and ($CurrentDomain.Name -eq $CurrentDomain.Forest.RootDomain.Name)) {
                $isForestRootDomain = $true
                WriteLog -Message "  -> Domain '$domainName' IS the AD forest root."
            } else {
                if ($CurrentDomain.Forest -and $CurrentDomain.Forest.RootDomain) {
                     WriteLog -Message "  -> Domain '$domainName' IS NOT the AD forest root (The forest root is '$($CurrentDomain.Forest.RootDomain.Name)')."
                } else {
                     WriteLog -Message "  -> Domain '$domainName' IS NOT the AD forest root (Could not confirm the current forest root via this specific domain object)."
                }
            }

            Write-Host -ForegroundColor:Cyan "Processing domain '$domainName' $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-Host "  LDAP DN: $distinguishedName"

            # --- START Per-Domain CSV Overwrite/Skip/Load Check for DetectAllDomains mode ---
    $perDomainCsvBaseName = "Exchange_OnPrem_Mailboxes_$($domainName)"
            $perDomainCsvSuffix = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
            $perDomainCsvFullPath = Join-Path -Path $OutputPath -ChildPath "$($perDomainCsvBaseName)${perDomainCsvSuffix}"

            if (Test-Path $perDomainCsvFullPath) {
                if (-not $ForceOverwriteCSV) {
                    # MODIFICATION START: Load data from existing CSV instead of just skipping
                    $loadMessage = "Per-domain CSV file '$perDomainCsvFullPath' for domain '$domainName' already exists and -ForceOverwriteCSV is `$false. Loading data from this file."
                    WriteLog -Message "INFO: $loadMessage"
                    Write-Host -ForegroundColor Cyan $loadMessage

                    try {
                        $existingData = Import-Csv -Path $perDomainCsvFullPath -Encoding UTF8 -ErrorAction Stop
                        if ($existingData -and $existingData.Count -gt 0) {
                            $Global:ScriptOverallMailboxData += $existingData
                            $successMessage = "Successfully loaded $($existingData.Count) records from '$perDomainCsvFullPath' and added them to global data. Live processing for domain '$domainName' will be skipped."
                            WriteLog -Message "INFO: $successMessage"
                            Write-Host -ForegroundColor Green $successMessage
                        } elseif ($existingData) { # File exists but is empty
                             $emptyFileMessage = "CSV file '$perDomainCsvFullPath' exists but is empty. No records were loaded. Live processing for domain '$domainName' will be skipped."
                             WriteLog -Message "INFO: $emptyFileMessage" # Changed from WARNING to INFO as it's an expected scenario
                             Write-Host -ForegroundColor Cyan $emptyFileMessage
                        } else { # Should not happen if Import-Csv did not error but returned $null for $existingData
                            $nullDataMessage = "Import from '$perDomainCsvFullPath' returned no data (possibly an empty or malformed file not detected as an error). Live processing for domain '$domainName' will be skipped."
                            WriteLog -Message "WARNING: $nullDataMessage"
                            Write-Host -ForegroundColor Yellow $nullDataMessage
                        }
                    } catch {
                        $importErrorMessage = "ERROR: Failed to import data from existing CSV '$perDomainCsvFullPath' for domain '$domainName'. Message: $($_.Exception.Message). Live processing for this domain will be skipped."
                        WriteLog -message $importErrorMessage
                        Write-Error $importErrorMessage # Keep as Write-Error for visibility
                    }
                    return # Skip live processing for this domain as data is loaded or attempt was made
                    # MODIFICATION END
                } else { # ForceOverwriteCSV is $true
                    $message = "Per-domain CSV file '$perDomainCsvFullPath' for domain '$domainName' exists and -ForceOverwriteCSV is `$true. Existing file will be deleted before processing."
                    WriteLog -Message "INFO: $message"
                    Write-Host -ForegroundColor Cyan $message
                    try {
                        Remove-Item -Path $perDomainCsvFullPath -Force -ErrorAction Stop
                        WriteLog -Message "INFO: File '$perDomainCsvFullPath' for domain '$domainName' deleted successfully."
                    } catch {
                        $errorMessage = "ERROR: Could not delete existing per-domain CSV file '$perDomainCsvFullPath' for domain '$domainName'. Message: $($_.Exception.Message)"
                        WriteLog -Message "ERROR: $errorMessage"
                        Write-Error $errorMessage
                        # Decide if script should stop or continue if deletion fails. Original script implies continuation.
                    }
                }
            }
            # --- END Per-Domain CSV Overwrite/Skip/Load Check ---

            [string[]]$pathsForMailboxProcessing = @()
            $domainDataFromProcessing = $null # Initialize to null

            if ($isForestRootDomain) {
                Write-Host "  Domain Status: AD Forest Root." -ForegroundColor Yellow
                WriteLog -Message "  Domain '$domainName' is the forest root. Attempting to retrieve first-level OUs."

                try {
                    $firstLevelOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $distinguishedName -SearchScope OneLevel -Server $domainName -ErrorAction Stop | Select-Object -ExpandProperty DistinguishedName

                    if ($firstLevelOUs -and $firstLevelOUs.Count -gt 0) {
                        $pathsForMailboxProcessing = $firstLevelOUs
                        Write-Host "    -> $($firstLevelOUs.Count) first-level OUs found under '$distinguishedName'. They will be passed to MailboxesProcessing."
                        $firstLevelOUs | ForEach-Object { WriteLog -Message "      -> First-level OU to include for MailboxesProcessing: '$_'" }
                    } else {
                        WriteLog -Message "  WARNING: No first-level OUs found under the root domain '$domainName'. Processing the domain root ('$distinguishedName') itself."
                        Write-Host "  WARNING: No first-level OUs found under the root domain '$domainName'. Processing the domain root itself." -ForegroundColor Yellow
                        $pathsForMailboxProcessing = @($distinguishedName)
                    }
                } catch {
                    # If Get-ADOrganizationalUnit fails, fall back to processing the domain root itself.
                    $errorMessage = "ERROR retrieving first-level OUs for the root domain '$domainName'. Message: $($_.Exception.Message). Will attempt to process the domain root ('$distinguishedName') itself."
                    WriteLog -message $errorMessage
                    Write-Warning $errorMessage
                    $pathsForMailboxProcessing = @($distinguishedName)
                }
                $domainDataFromProcessing = MailboxesProcessing -IncludedLDAPPaths $pathsForMailboxProcessing
                if ($null -ne $domainDataFromProcessing) { $Global:ScriptOverallMailboxData += $domainDataFromProcessing }


            } else { # Not the forest root domain
                Write-Host "  Domain Status: Not the AD Forest Root."
                WriteLog -Message "  Domain '$domainName' is not the forest root. The domain's own DN ('$distinguishedName') will be used for MailboxesProcessing."
                $pathsForMailboxProcessing = @($distinguishedName)
                $domainDataFromProcessing = MailboxesProcessing -IncludedLDAPPaths $pathsForMailboxProcessing
                if ($null -ne $domainDataFromProcessing) { $Global:ScriptOverallMailboxData += $domainDataFromProcessing }
            }

            # Export data for THIS specific domain if it was processed live (not loaded from existing CSV)
            # This is the per-domain CSV that might be loaded in future runs.
            if ($domainDataFromProcessing -and $domainDataFromProcessing.Count -gt 0) {
                try {
                    Export-CsvAtomic -InputObject $domainDataFromProcessing -Path $perDomainCsvFullPath -Encoding UTF8
                    WriteLog -Message "INFO: Live processed data for domain '$domainName' exported to '$perDomainCsvFullPath'."
                    Write-Host -ForegroundColor Green "Live processed data for domain '$domainName' exported to '$perDomainCsvFullPath'."
                } catch {
                    $exportError = "ERROR: Failed to export live processed data for domain '$domainName' to '$perDomainCsvFullPath'. Message: $($_.Exception.Message)"
                    WriteLog -message $exportError
                    Write-Error $exportError
                }
            } elseif ($domainDataFromProcessing) { # Processed live, but no mailboxes found
                 WriteLog -Message "INFO: Live processing for domain '$domainName' yielded no mailboxes. Per-domain CSV '$perDomainCsvFullPath' will not be created or will be empty if pre-deleted."
            }
            # If data was loaded from existing CSV, $domainDataFromProcessing would be $null here.

            Write-Host "  -> Other processing logic for domain '$domainName' (outside MailboxesProcessing) completed."
            WriteLog -Message "Finished global processing for domain: '$domainName'. IsForestRoot: $isForestRootDomain"
        }
    }
    #endregion Function Definitions

    #region Main Script Block
    $InventoryCompletedSuccessfully = $false # Initialize completion flag
    try { # Main try block for script execution
        if ($RemoteMailboxesOnly)
        {
            WriteLog -Message 'RemoteMailboxesOnly mode is enabled. Local mailbox collection will be skipped.'
            Write-Host 'RemoteMailboxesOnly mode is enabled. Local mailbox collection will be skipped.' -ForegroundColor Cyan
        }
        elseif ($DetectAllDomains)
        {
            Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
            WriteLog -Message "DetectAllDomains mode enabled. Checking for Active Directory module..."
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                $errorMessage = "CRITICAL ERROR: The Active Directory module is not installed. Please install it before running this script. The script will stop."
                WriteLog -message $errorMessage
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
                throw $errorMessage
            }
            WriteLog -Message "Importing Active Directory module..."
            Write-Host "Importing Active Directory module... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Import-Module ActiveDirectory -ErrorAction Stop

            WriteLog -Message "Retrieving forest information..."
            $forest = $null
            try {
                $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            } catch {
                $errorMessage = "CRITICAL ERROR: Failed to contact Active Directory forest. Check connectivity and permissions. Message: $($_.Exception.Message). The script will stop."
                WriteLog -message $errorMessage
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
                throw $errorMessage
            }

            if (-not $forest) {
                $errorMessage = "CRITICAL ERROR: Failed to retrieve current Active Directory forest information (forest object is null). The script will stop."
                WriteLog -message $errorMessage
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
                throw $errorMessage
            }
            WriteLog -Message "Current forest: $($forest.Name)"

            if (-not $forest.RootDomain) {
                $errorMessage = "CRITICAL ERROR: Failed to determine the root domain of the Active Directory forest '$($forest.Name)'. The script will stop."
                WriteLog -message $errorMessage
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
                throw $errorMessage
            }
            WriteLog -Message "Detected forest root domain: $($forest.RootDomain.Name)"

			# --- START Check for existing COMBINED CSV file if DetectAllDomains is $true ---
			$combinedCsvFileSuffixGlobal = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
    $globalCombinedCsvFile = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_AllDomains${combinedCsvFileSuffixGlobal}"

			$baseExchangeMailboxesInventoryPath = (Get-Item $OutputPath).Parent.Parent.FullName # Go up two levels to ExchangeMailboxesInventory
			$backupBaseDir = Join-Path -Path $baseExchangeMailboxesInventoryPath -ChildPath "Backup"

			if (Test-Path $globalCombinedCsvFile) {
				if (-not $ForceOverwriteCSV) {
					# Create a timestamped directory inside the backup base directory
					$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
					$currentBackupDir = Join-Path -Path $backupBaseDir -ChildPath $timestamp

					# Create the full backup path including the timestamped directory
					if (-not (Test-Path $currentBackupDir)) {
						New-Item -ItemType Directory -Path $currentBackupDir -Force | Out-Null
						WriteLog -Message "INFO: Created timestamped backup directory: '$currentBackupDir'"
					}

					# The backup file name should be the same as the source file name
					$backupFileName = Split-Path -Path $globalCombinedCsvFile -Leaf
					$backupFilePath = Join-Path -Path $currentBackupDir -ChildPath $backupFileName

					WriteLog -Message "INFO: Combined CSV file '$globalCombinedCsvFile' already exists and -ForceOverwriteCSV is `$false. Backing up the existing file to '$backupFilePath'."
					Write-Host -ForegroundColor Yellow "WARNING: Combined CSV file '$globalCombinedCsvFile' already exists. Backing it up to '$backupFilePath' before continuing."

					# Copy the file to the backup directory
					try {
						# Copy the file
						Copy-Item -Path $globalCombinedCsvFile -Destination $backupFilePath -Force -ErrorAction Stop
						WriteLog -Message "INFO: Successfully backed up '$globalCombinedCsvFile' to '$backupFilePath'."

						# After successful backup, remove the original file
						Remove-Item -Path $globalCombinedCsvFile -ErrorAction Stop
						WriteLog -Message "INFO: Successfully removed original combined CSV file '$globalCombinedCsvFile'."

					} catch {
						# Handle errors during copy or removal
						$errorMessage = "CRITICAL ERROR: Failed to process the combined CSV file '$globalCombinedCsvFile' (copy or removal failed). Error: $($_.Exception.Message)"
						WriteLog -message $errorMessage
						Write-Host -ForegroundColor Red $errorMessage
						$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
						Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
throw $errorMessage
					}
				} else {
					WriteLog -Message "INFO: Combined CSV file '$globalCombinedCsvFile' exists and -ForceOverwriteCSV is `$true. It will be deleted before new combined export at the end of processing."
					# Deletion will happen just before the final export of the combined file.
				}
			}
            $domainsToProcess = $forest.Domains
            if (-not $domainsToProcess -or $domainsToProcess.Count -eq 0) {
                $errorMessage = "CRITICAL ERROR: No domains were found in the forest '$($forest.Name)'. Check Active Directory configuration. The script will stop."
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
                WriteLog -message $errorMessage
                throw $errorMessage
            }

            WriteLog -Message "Number of domains to process: $($domainsToProcess.Count)"

            if ($domainsToProcess.Count -gt 0) {
                Write-Host "`nList of domains to process:" -ForegroundColor Green
                $domainsToProcess | Select-Object Name, @{Name="ExistingCSV"; Expression={
                    $domainCsvBaseName = "Exchange_OnPrem_Mailboxes_$($_.Name)"
                    $domainCsvSuffix = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
                    $domainCsvFullPath = Join-Path -Path $OutputPath -ChildPath "$($domainCsvBaseName)${domainCsvSuffix}"
                    if (Test-Path $domainCsvFullPath) {"Yes"} else {"No"}
                }} | Format-Table -AutoSize
                Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
            }

            WriteLog -Message "Starting processing of each domain via the Process-SpecificDomain function..."
            foreach ($domain in $domainsToProcess) {
                Process-SpecificDomain -CurrentDomain $domain
            }
            WriteLog -Message "Finished processing all domains."
            # Export the globally accumulated data to the AllDomains CSV
            if ($Global:ScriptOverallMailboxData.Count -gt 0) {
                WriteLog -Message "Exporting combined data for all processed domains to '$globalCombinedCsvFile'..."
                if ((Test-Path $globalCombinedCsvFile) -and $ForceOverwriteCSV) { # File might have been created by a previous version or if script was interrupted
                    try {
                        Remove-Item -Path $globalCombinedCsvFile -Force -ErrorAction Stop
                        WriteLog -Message "INFO: Successfully deleted existing combined CSV file '$globalCombinedCsvFile' due to ForceOverwriteCSV."
                    } catch {
                        WriteLog -Message "ERROR: Failed to delete existing combined CSV file '$globalCombinedCsvFile'. Export may fail or append. Error: $($_.Exception.Message)"
                    }
                }
                try {
                    Export-CsvAtomic -InputObject $Global:ScriptOverallMailboxData -Path $globalCombinedCsvFile -Encoding UTF8
                    WriteLog -Message "Successfully exported combined data to '$globalCombinedCsvFile'."
                    $null = Publish-SmartM365ExchangeLocalMailboxCsv -SourcePath $globalCombinedCsvFile -LatestFileName (Split-Path -Path $globalCombinedCsvFile -Leaf)
                    Write-Host -ForegroundColor Green "All processed mailbox data exported to: $globalCombinedCsvFile"
					$InputCsvForDuplicateScan = $globalCombinedCsvFile
					$scriptdatamailbox = $true
					$SendFileListEmailReportFileName = $globalCombinedCsvFile
                } catch {
                    $errorMessage = "Failed to export combined data to '$globalCombinedCsvFile': $($_.Exception.Message)"
                    WriteLog -Message "ERROR: $errorMessage"
                    Write-Error $errorMessage
					$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Failed to export combined data to '$globalCombinedCsvFile': $($_.Exception.Message)"
					Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
                }
            } else {
                WriteLog -Message "No data accumulated in \$Global:ScriptOverallMailboxData. Combined 'AllDomains' CSV will not be created or will be empty if it was pre-deleted."
                Write-Host -ForegroundColor Yellow "No mailbox data was collected or loaded. The combined file '$globalCombinedCsvFile' will not be created or will be empty."
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : No mailbox data was collected or loaded. The combined file '$globalCombinedCsvFile' will not be created or will be empty."
				Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
            }

        }
        else # DetectAllDomains is $false
        {
            WriteLog -Message "DetectAllDomains mode is $false."
            Write-Host "DetectAllDomains mode is $false." -ForegroundColor Green

            $combinedCsvFileForNonDetectAll = $null
            $createCombinedCsvForNonDetectAll = $true # Assume we create a combined CSV unless it's a single OU
            $combinedCsvFileSuffixLocal = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }

            if ($IncludedOrganizationalUnit -and $IncludedOrganizationalUnit.Count -eq 1) {
                $createCombinedCsvForNonDetectAll = $false # For single OU, MailboxesProcessing2 handles its own CSV. No separate "combined" for one.
                WriteLog -Message "INFO: Combined CSV will not be created for non-DetectAllDomains mode as only one specific OU/path is processed. Its specific CSV will be generated by MailboxesProcessing2 if data is found."
            } elseif ($IncludedOrganizationalUnit -and $IncludedOrganizationalUnit.Count -gt 1) {
                $pathIdentifiers = foreach ($path in $IncludedOrganizationalUnit) {
                    $firstPart = ($path -split ',')[0] -replace '^(OU=|CN=|DC=)','' -replace '[^a-zA-Z0-9_-]',''
                    if ($firstPart.Length -gt 15) { $firstPart = $firstPart.Substring(0,15) }
                    $firstPart
                }
                $identifierString = ($pathIdentifiers | Select-Object -Unique) -join "_"
                if ($identifierString.Length -gt 50) { $identifierString = $identifierString.Substring(0,50)}
                $combinedCsvFileForNonDetectAll = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_AllTargetedOUs_${identifierString}${combinedCsvFileSuffixLocal}"
            } else { # No OUs specified, processing current scope
                $combinedCsvFileForNonDetectAll = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_AllCurrentScope${combinedCsvFileSuffixLocal}"
            }

            if ($createCombinedCsvForNonDetectAll -and $combinedCsvFileForNonDetectAll) {
                if (Test-Path $combinedCsvFileForNonDetectAll) {
                    if (-not $ForceOverwriteCSV) {
                        $errorMessage = "CRITICAL ERROR: Combined CSV file '$combinedCsvFileForNonDetectAll' for specified OUs/scope already exists and -ForceOverwriteCSV is `$false. Script will stop."
                        WriteLog -message $errorMessage
                        Write-Host -ForegroundColor Red $errorMessage
						$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
						Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
throw $errorMessage
                    } else {
                        WriteLog -Message "INFO: Combined CSV file '$combinedCsvFileForNonDetectAll' exists and -ForceOverwriteCSV is `$true. It will be deleted."
                        try { Remove-Item -Path $combinedCsvFileForNonDetectAll -Force -ErrorAction Stop }
                        catch { WriteLog -Message "ERROR: Could not delete '$combinedCsvFileForNonDetectAll'. Message: $($_.Exception.Message)" }
                    }
                }
            }

            if ($IncludedOrganizationalUnit -and $IncludedOrganizationalUnit.Count -gt 0) {
                WriteLog -Message "Processing specified Organizational Units. $($IncludedOrganizationalUnit.Count) OU(s) provided via -IncludedOrganizationalUnit parameter."
                Write-Host "Processing specified Organizational Units. $($IncludedOrganizationalUnit.Count) OU(s) provided." -ForegroundColor Green
                $IncludedOrganizationalUnit | ForEach-Object { WriteLog -Message "  - Will process OU: $_" }
            } else {
                WriteLog -Message "No Organizational Units specified via -IncludedOrganizationalUnit parameter. MailboxesProcessing2 will attempt to retrieve all mailboxes in the current Exchange scope."
                Write-Host "Warning: No Organizational Units specified via -IncludedOrganizationalUnit parameter. MailboxesProcessing2 will attempt to retrieve all mailboxes in the current Exchange scope." -ForegroundColor Yellow
            }

            $ouData = MailboxesProcessing -IncludedLDAPPaths $IncludedOrganizationalUnit # This will call MailboxesProcessing2
            if ($null -ne $ouData) {
                 $Global:ScriptOverallMailboxData += $ouData
            }
            # Export the combined data if applicable for !DetectAllDomains (multiple OUs or current scope)
            if ($createCombinedCsvForNonDetectAll -and $Global:ScriptOverallMailboxData.Count -gt 0 -and $combinedCsvFileForNonDetectAll) {
                WriteLog -Message "Exporting combined data for specified OUs/scope to '$combinedCsvFileForNonDetectAll'..."
                # Overwrite check already performed, or ForceOverwriteCSV is true (file deleted)
                try {
                    Export-CsvAtomic -InputObject $Global:ScriptOverallMailboxData -Path $combinedCsvFileForNonDetectAll -Encoding UTF8
                    WriteLog -Message "Successfully exported combined data to '$combinedCsvFileForNonDetectAll'."
                    $null = Publish-SmartM365ExchangeLocalMailboxCsv -SourcePath $combinedCsvFileForNonDetectAll -LatestFileName (Split-Path -Path $combinedCsvFileForNonDetectAll -Leaf)
                    Write-Host -ForegroundColor Green "All processed mailbox data for specified scope exported to: $combinedCsvFileForNonDetectAll"
					$InputCsvForDuplicateScan = $combinedCsvFileForNonDetectAll
					$scriptdatamailbox = $true
					$SendFileListEmailReportFileName = $combinedCsvFileForNonDetectAll
                } catch {
                    $errorMessage = "Failed to export combined data to '$combinedCsvFileForNonDetectAll': $($_.Exception.Message)"
                    WriteLog -Message "ERROR: $errorMessage"
                    Write-Error $errorMessage
                }
            } elseif ($createCombinedCsvForNonDetectAll -and (-not $Global:ScriptOverallMailboxData -or $Global:ScriptOverallMailboxData.Count -eq 0)) {
                WriteLog -Message "No data accumulated in \$Global:ScriptOverallMailboxData for specified OUs/scope. Combined CSV '$combinedCsvFileForNonDetectAll' will not be created or will be empty if pre-deleted."
                Write-Host -ForegroundColor Yellow "No mailbox data was collected for the specified OUs/scope. The combined file '$combinedCsvFileForNonDetectAll' will not be created or will be empty."
            }
        }
        if ($IncludeRemoteMailboxes) {
            $remoteResult = Invoke-SmartM365ExchangeRemoteMailboxInventory -RemoteOutputPath $RemoteMailboxOutputPath -IncludedLDAPPaths $IncludedOrganizationalUnit
            if ($remoteResult -and $remoteResult.CombinedCsv) {
                WriteLog -Message ("Remote mailbox inventory generated. Records: {0}; CombinedCsv: {1}" -f $remoteResult.RecordCount, $remoteResult.CombinedCsv)
            }
        }
        else {
            WriteLog -Message 'Remote mailbox inventory skipped. Use -IncludeRemoteMailboxes or set IncludeRemoteMailboxes to true in local config to enable it.'
        }

        WriteLog -Message "END of script main logic execution."
        Write-Host "END of script main logic execution... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
        $InventoryCompletedSuccessfully = $true # Set flag for successful completion
    #endregion Main Script Block
    }
    finally { # This is the finally for the main try block that starts after variable initialization
        $EndTime = Get-Date
        if ($InventoryCompletedSuccessfully) {
            WriteLog -Message "Inventory finished. Total duration: $($EndTime - $StartTime)."
            Write-Host "Inventory finished. Total duration: $($EndTime - $StartTime)."
        } else { # Script was interrupted or had an unhandled terminating error
            $interruptionMessage = "Inventory interrupted or terminated due to an error. Total duration: $($EndTime - $StartTime)."
            Write-Host -ForegroundColor Yellow $interruptionMessage
        }
    }
}
finally {

if ($scriptdatamailbox -eq $true) {
	Write-Host "-----------------------------------------------------------------------------------------"
	Write-Host "Remove Duplicate in Export Exchange 2016 mailboxes inventory ..."
	WriteLog -Message "Remove Duplicate in Export Exchange 2016 mailboxes inventory ..."

    # Define paths
	$TempOutputCsv = [System.IO.Path]::ChangeExtension($InputCsvForDuplicateScan, $null) + "_WithoutDuplicateSMTP.csv"
    $LogFileDuplicate = $global:LogTextFile -replace '\.log$', '_duplicate.log'

    # Load CSV
    WriteLog -Message "Load CSV : $InputCsvForDuplicateScan"
    $rows = Import-Csv -Path $InputCsvForDuplicateScan

    # Create hashtable to track unique addresses (case-insensitive)
    WriteLog -Message "Create hashtable to track unique addresses"
    $seen = @{}
    $uniqueRows = @()
    $duplicates = @()

    foreach ($row in $rows) {
        $address = $row.PrimarySMTPaddress.ToLower()
        if ($seen.ContainsKey($address)) {
            $duplicates += $row
        } else {
            $seen[$address] = $true
            $uniqueRows += $row
        }
    }

    # Check if duplicates exist
    if ($duplicates.Count -eq 0) {
        WriteLog -Message "No duplicates found. CSV file remains unchanged."

        # Clean up temporary file if it exists
        if (Test-Path $TempOutputCsv) {
            try {
                Remove-Item -Path $TempOutputCsv -Force
                WriteLog -Message "Temporary file removed: $TempOutputCsv"
            } catch {
                WriteLog -Message "Failed to remove temporary file: $_" "ERROR"
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : No Duplicates Primary SMTP found - Failed to remove temporary file: $_"
				Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
throw $errorMessage
            }
        }
    } else {
        WriteLog -Message "Duplicates Primary SMTP detected. Processing cleanup..."
        # Export cleaned CSV to temporary file
        try {
            Export-CsvAtomic -InputObject $uniqueRows -Path $TempOutputCsv -Encoding UTF8
        } catch {
            WriteLog -Message "Failed to export cleaned CSV: $_" "ERROR"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Duplicates Primary SMTP detected - Failed to export cleaned CSV: $_"
			Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
throw $errorMessage
        }

        # Log duplicates with timestamp
        $logEntries = @()
        foreach ($dup in $duplicates) {
            $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Duplicate found - PrimarySMTPaddress: $($dup.PrimarySMTPaddress), DistinguishedName: $($dup.DistinguishedName)"
            WriteLog -Message $entry
            $logEntries += $entry
        }
        try {
            $logEntries | Out-File -FilePath $LogFileDuplicate -Encoding UTF8
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Duplicates Primary SMTP detected."
			Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body -Attachments $LogFileDuplicate
        } catch {
            WriteLog -Message "Failed to write log file: $_" "ERROR"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Duplicates Primary SMTP detected - Failed to write log file: $_"
			Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
throw $errorMessage
        }

        # Rename original file by appending '-versionoriginal'
        $originalName = Split-Path -Leaf $InputCsvForDuplicateScan
        $originalFolder = Split-Path -Parent $InputCsvForDuplicateScan
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($originalName)
        $extension = [System.IO.Path]::GetExtension($originalName)
        $renamedPath = Join-Path $originalFolder "$baseName-versionoriginal$extension"

        try {
            if (Test-Path $renamedPath) {
                if ($ForceOverwriteCSV) {
                    WriteLog -Message "File already exists: $renamedPath. Removing due to ForceOverwriteCSV = $true"
                    Remove-Item -Path $renamedPath -Force
                } else {
                    WriteLog -Message "File already exists: $renamedPath. Rename aborted." "ERROR"
					$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Duplicates Primary SMTP detected - File already exists: $renamedPath. Rename aborted."
					Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
throw $errorMessage
                }
            }
            Rename-Item -Path $InputCsvForDuplicateScan -NewName $renamedPath
            WriteLog -Message "Original file renamed to: $renamedPath"
        } catch {
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Failed to rename original file: $_"
			Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
            WriteLog -Message "Failed to rename original file: $_" "ERROR"
throw $errorMessage
        }

        # Move cleaned file back to original name
        try {
            if (Test-Path $InputCsvForDuplicateScan) {
                if ($ForceOverwriteCSV) {
                    WriteLog -Message "File already exists: $InputCsvForDuplicateScan. Removing due to ForceOverwriteCSV = $true"
                    Remove-Item -Path $InputCsvForDuplicateScan -Force
                } else {
                    WriteLog -Message "File already exists: $InputCsvForDuplicateScan. Move aborted." "ERROR"
					$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : File already exists: $InputCsvForDuplicateScan. Move aborted."
					Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
throw $errorMessage
                }
            }
            Move-Item -Path $TempOutputCsv -Destination $InputCsvForDuplicateScan
            WriteLog -Message "Cleaned file moved to: $InputCsvForDuplicateScan"
        } catch {
            WriteLog -Message "Failed to move cleaned file: $_" "ERROR"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Failed to move cleaned file: $_"
			Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
throw $errorMessage
        }

        # Summary log
        WriteLog -Message "Processing complete."
        WriteLog -Message "Total rows processed: $($rows.Count)"
        WriteLog -Message "Unique rows retained: $($uniqueRows.Count)"
        WriteLog -Message "Duplicates found: $($duplicates.Count)"
        WriteLog -Message "Cleaned file saved to: $InputCsvForDuplicateScan"
        WriteLog -Message "Duplicates logged in: $LogFile"
        WriteLog -Message "Original file renamed to: $renamedPath"
    }
}

if ($scriptdatamailbox -eq $true -and $scriptdatamegewithperm -eq $true -and $IncludeADPermission -and (-not $OnlyADPermission)) {
	Write-Host "-----------------------------------------------------------------------------------------"
	Write-Host "Combine Export Exchange 2016 mailboxes inventory with last permission inventory ..."
	try {
		$SourceDirectory = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxOnlyAdPermissionCsvLogFolderPath' -DefaultValue ''
	} catch {
		$errorMessage = "Error retrieving local configuration path LastOutput $_"
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
		throw
	}

	if (-not (Test-Path $SourceDirectory)) {
		$warningMessage = "Permission merge skipped because the source directory '$SourceDirectory' is not available."
		WriteLog -Message $warningMessage "WARNING"
		Write-Warning $warningMessage
		$scriptdatamegewithperm = $false
	}

	if ($scriptdatamegewithperm -eq $true) {
	try {
		$DestinationDirectory = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCsvLogFolderPath' -DefaultValue ''
	} catch {
		$errorMessage = "Error retrieving local configuration path DestinationDirectory $_"
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
		throw
	}

	if (-not (Test-Path $DestinationDirectory)) {
		$errorMessage = "The share '$DestinationDirectory' is not available. Stopping the script."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
		throw
	}

	try {
		$ScriptCsvLogFolderPathectory = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCombinedPermissionCsvLogFolderPath' -DefaultValue ''
	} catch {
		$errorMessage = "Error retrieving local configuration path ScriptCsvLogFolderPathectory $_"
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
		throw
	}

	if (-not (Test-Path $ScriptCsvLogFolderPathectory)) {
		$errorMessage = "The share '$ScriptCsvLogFolderPathectory' is not available. Stopping the script."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
		throw
	}

	# --- Step 2: Process files based on mode (Consolidation vs. Standard) ---

	$allDomainsFile = Get-ChildItem -Path $DestinationDirectory -Filter "*_AllDomains.csv" -File | Select-Object -First 1
	WriteLog -Message "$allDomainsFile found in the Destination directory."

	$sourceFiles = Get-ChildItem -Path $SourceDirectory -Filter "*_OnlyADPermission.csv" -File

	if ($sourceFiles.Count -eq 0) {
		WriteLog -Message "No '*_OnlyADPermission.csv' files found in the source directory. Permission merge skipped." "WARNING"
		Write-Warning "No '*_OnlyADPermission.csv' files found in the source directory. Permission merge skipped."
		$scriptdatamegewithperm = $false
	}

	if ($scriptdatamegewithperm -eq $true) {
	if ($allDomainsFile) {
		# --- Consolidation Mode: Merge all sources into one destination ---
		WriteLog -Message "Found '$($allDomainsFile.Name)'. Activating consolidation mode."
		$fullPaths = $sourceFiles | Select-Object -ExpandProperty FullName
		WriteLog -Message "All $($sourceFiles.Count) source files will be merged to update this single destination file:`n$($fullPaths -join "`n")"
		try {
			Write-Verbose "Creating a consolidated lookup table from all source files..."
			$consolidatedLookupTable = @{}
			foreach ($sourceFile in $sourceFiles) {
				Write-Verbose "  -> Adding source: $($sourceFile.Name)"
				Import-Csv -Path $sourceFile.FullName | ForEach-Object {
					if (-not [string]::IsNullOrEmpty($_.UserPrincipalName)) {
						$consolidatedLookupTable[$_.UserPrincipalName] = $_
					}
				}
			}
			Write-Verbose "Consolidated lookup table created with $($consolidatedLookupTable.Count) unique entries."

			Write-Verbose "Importing data from '$($allDomainsFile.Name)'..."
			$destinationData = Import-Csv -Path $allDomainsFile.FullName

			$updatedCount = 0
			$notFoundCount = 0

			$destinationData | ForEach-Object {
				$currentUserPrincipalName = $_.UserPrincipalName
				if ($consolidatedLookupTable.ContainsKey($currentUserPrincipalName)) {
					$_.SendAsCount = $consolidatedLookupTable[$currentUserPrincipalName].SendAsCount
					$_.SendAs = $consolidatedLookupTable[$currentUserPrincipalName].SendAs
					$updatedCount++
				} else {
					$notFoundCount++
				}
			}

			$outputFilePath = Join-Path -Path $ScriptCsvLogFolderPathectory -ChildPath $allDomainsFile.Name
			$SendFileListEmailReportFileName = $outputFilePath
			WriteLog -Message "Exporting updated data to '$outputFilePath'..."
			if ($PSCmdlet.ShouldProcess($outputFilePath, "Export Updated Consolidated CSV")) {
				Export-CsvAtomic -InputObject $destinationData -Path $outputFilePath -Encoding UTF8
			}

			WriteLog -Message "--------------------------------------------------"
			WriteLog -Message "Consolidation complete."
			WriteLog -Message "File updated: $outputFilePath"
			WriteLog -Message "Users updated: $updatedCount. Not found: $notFoundCount."

		} catch {
			Write-Error "An unexpected error occurred while processing the consolidated file '$($allDomainsFile.Name)'. Details: $_"
			$errorMessage = "ERROR retrieving local configuration path: $_"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
			Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
			throw
		}

	} else {
		# --- Standard Mode: One-to-one file matching ---
		WriteLog -Message "No '*_AllDomains.csv' file found. Proceeding with standard one-to-one matching."

		$summary = @{ Processed = 0; Skipped = 0; Errors = 0; SkippedFiles = [System.Collections.Generic.List[string]]::new() }

		foreach ($sourceFile in $sourceFiles) {
			try {
				WriteLog -Message "Processing source file: $($sourceFile.Name)"
				$baseName = $sourceFile.BaseName.Replace('_OnlyADPermission', '')
				$destinationFile = Get-ChildItem -Path $DestinationDirectory -Filter "$baseName*.csv" -File | Select-Object -First 1

				if (-not $destinationFile) {
					WriteLog -Message "No matching destination file found for '$baseName*' for '$($sourceFile.Name)'. File skipped." "WARNING"
					$summary.Skipped++; $summary.SkippedFiles.Add($sourceFile.Name)
					continue
				}

				WriteLog -Message "Match found: '$($sourceFile.Name)' -> '$($destinationFile.Name)'"

				$lookupTable = @{}
				Import-Csv -Path $sourceFile.FullName | ForEach-Object {
					if (-not [string]::IsNullOrEmpty($_.UserPrincipalName)) {
						$lookupTable[$_.UserPrincipalName] = $_
					}
				}

				$destinationData = Import-Csv -Path $destinationFile.FullName
				$updatedCount = 0; $notFoundCount = 0

				$destinationData | ForEach-Object {
					if ($lookupTable.ContainsKey($_.UserPrincipalName)) {
						$_.SendAsCount = $lookupTable[$_.UserPrincipalName].SendAsCount
						$_.SendAs = $lookupTable[$_.UserPrincipalName].SendAs
						$updatedCount++
					} else {
						$notFoundCount++
					}
				}

				$outputFilePath = Join-Path -Path $ScriptCsvLogFolderPathectory -ChildPath $destinationFile.Name
				$SendFileListEmailReportFileName = $outputFilePath
				if ($PSCmdlet.ShouldProcess($outputFilePath, "Export Updated CSV")) {
					Export-CsvAtomic -InputObject $destinationData -Path $outputFilePath -Encoding UTF8
				}

				$summary.Processed++
				WriteLog -Message "Update for '$($destinationFile.Name)' finished. Users updated: $updatedCount. Not found: $notFoundCount."

			} catch {
				Write-Error "An unexpected error occurred while processing file '$($sourceFile.Name)'. Details: $_"
				$summary.Errors++
				$errorMessage = "ERROR retrieving local configuration path: $_"
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				Send-SmartM365OptionalEmailHtmlReport -BodyHtml $body
				throw
			}
		}

		# --- Final Summary for Standard Mode ---
		WriteLog -Message "--------------------------------------------------"
		WriteLog -Message "All files processed."
		WriteLog -Message "Files processed successfully: $($summary.Processed)"
		WriteLog -Message "Files skipped (no destination): $($summary.Skipped)"
		WriteLog -Message "Errors: $($summary.Errors)"

		if ($summary.SkippedFiles.Count -gt 0) {
			WriteLog -Message "List of skipped files:" "WARNING"
			$summary.SkippedFiles | ForEach-Object { Write-Warning "- $_" }
		}
	}
	}
	}
}

if ($InventoryCompletedSuccessfully -eq $false) {
    $EndTimeFinalRedundant = Get-Date

$interruptionMessageRedundant = "Script (outer finally) interrupted or terminated due to an error. Total duration: $($EndTimeFinalRedundant - $StartTime)."

Write-Host -ForegroundColor Yellow $interruptionMessageRedundant
    Stop-SmartM365TranscriptSafely
    try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage 'Inventory' } catch {}
}
Else
{
	    if ($GenerateReport -and -not $OnlyADPermission) {
        try {
            $reportResult = Invoke-SmartM365ExchangeLocalMailboxReport -UseCurrentInventoryData:$true
            if ($reportResult) { WriteLog -Message ("Mailbox report generated. DailyStatsCsv: {0}; SummaryCsv: {1}; Rows: {2}" -f $reportResult.DailyStatsCsv, $reportResult.SummaryCsv, $reportResult.RowCount) }
        }
        catch {
            WriteLog -Message ("Mailbox report generation failed: {0}" -f $_.Exception.Message) "ERROR"
            throw
        }
    }
    elseif ($OnlyADPermission) { WriteLog -Message 'Mailbox report generation skipped because OnlyADPermission mode is enabled.' }
    else { WriteLog -Message 'Mailbox report generation skipped because GenerateReport is disabled.' }

	#SendFileListEmailReport -Files @($SendFileListEmailReportFileName) -Title "$TaskName" -Message "$TaskName : All processed mailbox data exported to $SendFileListEmailReportFileName."

	    #region Run Summary
    try {
        $EndTimeSummary = Get-Date
        $durationSummary = $EndTimeSummary - $StartTime

        $includedOuText = ""
        if ($IncludedOrganizationalUnit -and $IncludedOrganizationalUnit.Count -gt 0) {
            $includedOuText = ($IncludedOrganizationalUnit -join "; ")
        }

        $mailboxRecordCount = $null
        if ($Global:ScriptOverallMailboxData) {
            $mailboxRecordCount = $Global:ScriptOverallMailboxData.Count
        }

        Write-Host
        Write-Host -ForegroundColor Cyan "Run Summary:"
        WriteLog -Message "Run Summary:"

        Write-Host "  ScriptVersion     : $ScriptVersion"
        Write-Host "  StartTime         : $($StartTime.ToString('o'))"
        Write-Host "  EndTime           : $($EndTimeSummary.ToString('o'))"
        Write-Host "  Duration          : $durationSummary"
        Write-Host "  Host              : $env:COMPUTERNAME"
        Write-Host "  User              : $env:USERNAME"
        Write-Host "  OnlyADPermission  : $OnlyADPermission"
        Write-Host "  IncludeADPermission: $IncludeADPermission"
        Write-Host "  DetectAllDomains  : $DetectAllDomains"
        Write-Host "  IncludedOUs       : $includedOuText"
        Write-Host "  OutputPath        : $OutputPath"
        if ($SendFileListEmailReportFileName) { Write-Host "  PrimaryOutputFile : $SendFileListEmailReportFileName" }
        if ($mailboxRecordCount -ne $null) { Write-Host "  MailboxRecords    : $mailboxRecordCount" }

        WriteLog -Message ("  ScriptVersion      : {0}" -f $ScriptVersion)
        WriteLog -Message ("  StartTime          : {0}" -f $StartTime.ToString('o'))
        WriteLog -Message ("  EndTime            : {0}" -f $EndTimeSummary.ToString('o'))
        WriteLog -Message ("  Duration           : {0}" -f $durationSummary)
        WriteLog -Message ("  OnlyADPermission   : {0}" -f $OnlyADPermission)
        WriteLog -Message ("  IncludeADPermission: {0}" -f $IncludeADPermission)
        WriteLog -Message ("  DetectAllDomains   : {0}" -f $DetectAllDomains)
        WriteLog -Message ("  IncludedOUs        : {0}" -f $includedOuText)
        WriteLog -Message ("  OutputPath         : {0}" -f $OutputPath)
        if ($SendFileListEmailReportFileName) { WriteLog -Message ("  PrimaryOutputFile  : {0}" -f $SendFileListEmailReportFileName) }
        if ($mailboxRecordCount -ne $null) { WriteLog -Message ("  MailboxRecords     : {0}" -f $mailboxRecordCount) }
    } catch {
        WriteLog -Message "Failed to build Run Summary: $_" "ERROR"
    }
    #endregion Run Summary
	#region Cleanup
	# Clean up old CSV files + old log files
	# Automatically excludes all generated CSVs via global:csvGeneratedPaths + current transcript and log files via global variables
	RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
    if ($IncludeRemoteMailboxes -and -not [string]::IsNullOrWhiteSpace($RemoteMailboxOutputPath) -and (Test-Path -LiteralPath $RemoteMailboxOutputPath)) {
        RemoveOldFiles -Path $RemoteMailboxOutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
    }
	RemoveOldFiles -Path $logPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
	WriteLog -Message "$TaskName completed."
    Stop-SmartM365TranscriptSafely
    try { Complete-SmartM365ExecutionContext -Status Auto } catch {}
	#endregion
}
}
#End of script
