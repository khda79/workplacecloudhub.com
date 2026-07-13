#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 Backup policy scope mailbox coverage inventory.

.DESCRIPTION
    Reads the configured Entra group used as the Microsoft 365 Backup mailbox policy scope,
    compares its user members with the latest Exchange Online mailbox inventory, and exports
    Power BI friendly CSV files. This is an expected-scope inventory, not proof that Microsoft
    365 Backup has protected each mailbox.

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER BackupPolicyScopeGroupId
    Entra group object id used by the Microsoft 365 Backup policy scope.

.PARAMETER BackupPolicyScopeGroupDisplayName
    Optional display name used in logs and CSVs.

.PARAMETER DataLastFolder
    Optional DATA-LAST folder override. Defaults to LatestCsvFolderPath from config.

.PARAMETER OutputPath
    Optional output directory override. Defaults to ScriptCsvLogFolderPath from config.

.PARAMETER DisableSharePointUpload
    Skips SharePoint upload for generated latest CSV files.

.PARAMETER MaxItems
    Limits group members and mailbox rows for smoke tests. Generated CSV names include _MAXITEMS.

.VERSION
1.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Minimum application permissions: GroupMember.Read.All, User.Read.All.
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$BackupPolicyScopeGroupId = '',
    [string]$BackupPolicyScopeGroupDisplayName = '',
    [string]$DataLastFolder = '',
    [string]$OutputPath = '',
    [switch]$DisableSharePointUpload,
    [int]$MaxItems = 0
)

if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
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
$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MaximumFunctionCount = 32768
$ScriptVersion = '1.0'
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
$script:SmartM365GlobalConfig = $null
$script:LogPath = ''

function Import-SmartM365CoreModule {
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
    if ($null -ne $script:SmartM365GlobalConfig) { return $script:SmartM365GlobalConfig }
    if ($null -ne $script:SmartM365EffectiveConfig) {
        $script:SmartM365GlobalConfig = $script:SmartM365EffectiveConfig
        return $script:SmartM365GlobalConfig
    }
    $script:SmartM365GlobalConfig = [pscustomobject]@{}
    return $script:SmartM365GlobalConfig
}

function Resolve-SmartM365ConfigValue {
    param([AllowNull()]$Value)
    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value -notmatch '\{\{[^}]+\}\}') { return $Value }
    $globalConfig = Get-SmartM365GlobalConfig
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $matches) {
            $property = $globalConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $tokenValue = Resolve-SmartM365ConfigValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-ScriptLocalConfigValue {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )
    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -isnot [string]) { return Resolve-SmartM365ConfigValue -Value $property.Value }
        $text = $property.Value.Trim()
        if ($text -and $text -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return Resolve-SmartM365ConfigValue -Value $property.Value }
    }
    $globalConfig = Get-SmartM365GlobalConfig
    $globalProperty = $globalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) { return $DefaultValue }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

function Write-ScopeLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
}

function Ensure-GraphAuthModule {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Microsoft.Graph.Authentication module is required.'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

function Connect-GraphAppOnly {
    param(
        [Parameter(Mandatory = $true)][string]$AppId,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-ScopeLog -Message 'Connecting to Microsoft Graph with app-only certificate authentication.'
    Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null
    Write-ScopeLog -Message 'Microsoft Graph connected.' -Level SUCCESS
}

function Invoke-GraphGetCollection {
    param([Parameter(Mandatory = $true)][string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
        foreach ($item in @($response.value)) { $items.Add($item) | Out-Null }
        $nextProperty = $response.PSObject.Properties['@odata.nextLink']
        $nextUri = if ($null -ne $nextProperty) { [string]$nextProperty.Value } else { '' }
        if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { break }
    }
    if ($MaxItems -gt 0 -and $items.Count -gt $MaxItems) { return @($items | Select-Object -First $MaxItems) }
    return @($items)
}

function Get-ObjectValue {
    param([AllowNull()]$Object, [Parameter(Mandatory = $true)][string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }
    return $null
}

function Normalize-MailAddress {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return $Value.Trim().ToLowerInvariant()
}

function Build-MailboxAddressIndex {
    param([object[]]$Mailboxes)
    $index = @{}
    foreach ($mailbox in $Mailboxes) {
        foreach ($propertyName in @('UserPrincipalName','UPN','PrimarySmtpAddress','WindowsEmailAddress','ExternalDirectoryObjectId')) {
            $value = Normalize-MailAddress ([string](Get-ObjectValue -Object $mailbox -Names @($propertyName)))
            if ($value -and -not $index.ContainsKey($value)) { $index[$value] = $mailbox }
        }
        $aliases = [string](Get-ObjectValue -Object $mailbox -Names @('EmailAddresses','ProxyAddresses'))
        foreach ($alias in @($aliases -split '[;|,]')) {
            $clean = (Normalize-MailAddress $alias) -replace '^smtp:', ''
            if ($clean -and -not $index.ContainsKey($clean)) { $index[$clean] = $mailbox }
        }
    }
    return $index
}

function ConvertTo-GroupMemberRow {
    param($Member, [string]$RunId, [string]$RunDateUtc, [string]$TenantName, [string]$GroupId, [string]$GroupDisplayName)
    $upn = [string](Get-ObjectValue -Object $Member -Names @('userPrincipalName'))
    $mail = [string](Get-ObjectValue -Object $Member -Names @('mail'))
    $enabled = Get-ObjectValue -Object $Member -Names @('accountEnabled')
    [pscustomobject]@{
        RunId = $RunId
        RunDateUtc = $RunDateUtc
        TenantName = $TenantName
        GroupId = $GroupId
        GroupDisplayName = $GroupDisplayName
        MemberId = [string](Get-ObjectValue -Object $Member -Names @('id'))
        MemberDisplayName = [string](Get-ObjectValue -Object $Member -Names @('displayName'))
        MemberUserPrincipalName = $upn
        MemberMail = $mail
        MemberType = [string](Get-ObjectValue -Object $Member -Names @('@odata.type'))
        AccountEnabled = if ($null -eq $enabled) { '' } else { [string]$enabled }
        Status = 'OK'
        NumericValue = 1
        TextValue = $upn
        Threshold = ''
        Details = 'Group member in expected Microsoft 365 Backup policy scope.'
    }
}

function ConvertTo-CoverageRows {
    param(
        [object[]]$Members,
        [hashtable]$MailboxIndex,
        [string]$RunId,
        [string]$RunDateUtc,
        [string]$TenantName,
        [string]$GroupId,
        [string]$GroupDisplayName
    )
    foreach ($member in $Members) {
        $keys = @(
            Normalize-MailAddress ([string](Get-ObjectValue -Object $member -Names @('userPrincipalName'))),
            Normalize-MailAddress ([string](Get-ObjectValue -Object $member -Names @('mail'))),
            Normalize-MailAddress ([string](Get-ObjectValue -Object $member -Names @('id')))
        ) | Where-Object { $_ }
        $mailbox = $null
        foreach ($key in $keys) {
            if ($MailboxIndex.ContainsKey($key)) { $mailbox = $MailboxIndex[$key]; break }
        }
        $hasMailbox = $null -ne $mailbox
        [pscustomobject]@{
            RunId = $RunId
            RunDateUtc = $RunDateUtc
            TenantName = $TenantName
            GroupId = $GroupId
            GroupDisplayName = $GroupDisplayName
            MemberId = [string](Get-ObjectValue -Object $member -Names @('id'))
            MemberDisplayName = [string](Get-ObjectValue -Object $member -Names @('displayName'))
            MemberUserPrincipalName = [string](Get-ObjectValue -Object $member -Names @('userPrincipalName'))
            MemberMail = [string](Get-ObjectValue -Object $member -Names @('mail'))
            MailboxFoundInInventory = $hasMailbox
            MailboxDisplayName = if ($hasMailbox) { [string](Get-ObjectValue -Object $mailbox -Names @('DisplayName','Name')) } else { '' }
            MailboxUserPrincipalName = if ($hasMailbox) { [string](Get-ObjectValue -Object $mailbox -Names @('UserPrincipalName','UPN')) } else { '' }
            PrimarySmtpAddress = if ($hasMailbox) { [string](Get-ObjectValue -Object $mailbox -Names @('PrimarySmtpAddress','WindowsEmailAddress')) } else { '' }
            RecipientTypeDetails = if ($hasMailbox) { [string](Get-ObjectValue -Object $mailbox -Names @('RecipientTypeDetails')) } else { '' }
            ExpectedPolicyScopeStatus = if ($hasMailbox) { 'InScopeMailbox' } else { 'GroupMemberWithoutMailboxInInventory' }
            ProtectionEvidenceStatus = 'NotVerifiedByThisScript'
            Status = if ($hasMailbox) { 'OK' } else { 'Warning' }
            NumericValue = if ($hasMailbox) { 1 } else { 0 }
            TextValue = if ($hasMailbox) { 'Mailbox matched in Exchange inventory' } else { 'No mailbox match in Exchange inventory' }
            Threshold = 'MailboxFoundInInventory=True'
            Details = 'Expected scope comparison only; use Microsoft 365 Backup protection-unit inventory for protection evidence.'
        }
    }
}

function Export-InventoryCsv {
    param(
        [Parameter(Mandatory = $true)][string]$BaseName,
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [string[]]$Columns = @()
    )
    $name = if ($MaxItems -gt 0) { '{0}_MAXITEMS' -f $BaseName } else { $BaseName }
    Export-SmartM365Csv -BaseFileName $name -OutputPath $ScriptCsvLogFolderPath -GlobalPath $LatestCsvFolderPath -Data $Rows -Columns $Columns -Encoding UTF8 -Delimiter ',' -NoSharePointUpload:($DisableSharePointUpload.IsPresent) | Out-Null
}

Import-SmartM365CoreModule
Ensure-GraphAuthModule
$ScriptLocalConfig = Get-ScriptLocalConfig

$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$AppId = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '')
$TenantId = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '')
$Thumb = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '')
$OrgDomain = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue $Tenant)
$ScriptCsvLogFolderPath = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue '')
$LatestCsvFolderPath = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '')
$LogAllRootPath = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue '')
$InputDataLastFolder = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'InputDataLastFolder' -DefaultValue $LatestCsvFolderPath)

if ([string]::IsNullOrWhiteSpace($BackupPolicyScopeGroupId)) { $BackupPolicyScopeGroupId = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'BackupPolicyScopeGroupId' -DefaultValue '') }
if ([string]::IsNullOrWhiteSpace($BackupPolicyScopeGroupDisplayName)) { $BackupPolicyScopeGroupDisplayName = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'BackupPolicyScopeGroupDisplayName' -DefaultValue '') }
if (-not [string]::IsNullOrWhiteSpace($DataLastFolder)) { $InputDataLastFolder = $DataLastFolder }
if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $ScriptCsvLogFolderPath = $OutputPath }
if ([string]::IsNullOrWhiteSpace($ScriptCsvLogFolderPath)) { $ScriptCsvLogFolderPath = Join-Path -Path $PSScriptRoot -ChildPath 'Output' }
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $LatestCsvFolderPath = $ScriptCsvLogFolderPath }
if ([string]::IsNullOrWhiteSpace($InputDataLastFolder)) { $InputDataLastFolder = $LatestCsvFolderPath }
if ([string]::IsNullOrWhiteSpace($BackupPolicyScopeGroupId)) { throw 'BackupPolicyScopeGroupId is required.' }

$runId = [guid]::NewGuid().ToString()
$runDateUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath 'SmartM365-M365-BackupPolicyScope-Inventory' }
$script:LogPath = Join-Path -Path $logRoot -ChildPath ("SmartM365-M365-BackupPolicyScope-Inventory_{0}.log" -f $timestamp)
foreach ($folder in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logRoot)) {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
}
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $script:LogPath
$global:LogTextFile = $script:LogPath

try {
    Write-ScopeLog -Message "Starting $TaskName. Tenant=$Tenant RunId=$runId"
    Write-ScopeLog -Message "This inventory reports expected backup policy scope membership; it does not verify Microsoft 365 Backup protection state."
    if ($MaxItems -gt 0) { Write-ScopeLog -Message "MaxItems test mode enabled: $MaxItems" -Level WARN }

    $mailboxCsvPath = Join-Path -Path $InputDataLastFolder -ChildPath 'Exchange_EXO_Mailboxes_AllDomains.csv'
    $resolvedMailboxCsvPath = Resolve-SmartM365CsvPathWithSharePointFallback -Path $mailboxCsvPath -Description 'Exchange Online mailbox inventory' -Required
    Write-ScopeLog -Message "Loading mailbox inventory: $resolvedMailboxCsvPath"
    $mailboxes = @(Import-Csv -LiteralPath $resolvedMailboxCsvPath -ErrorAction Stop)
    if ($MaxItems -gt 0) { $mailboxes = @($mailboxes | Select-Object -First $MaxItems) }
    Write-ScopeLog -Message "Mailbox inventory rows loaded: $($mailboxes.Count)"

    Connect-GraphAppOnly -AppId $AppId -TenantId $TenantId -Thumbprint $Thumb

    Invoke-SmartM365Preflight `
        -ScriptName $TaskName `
        -RequiredModules @('Microsoft.Graph.Authentication') `
        -OutputPaths @($ScriptCsvLogFolderPath, $LatestCsvFolderPath) `
        -RequiredGraphApplicationPermissions @('GroupMember.Read.All','User.Read.All') `
        -GraphProbeUris @("https://graph.microsoft.com/v1.0/groups/$BackupPolicyScopeGroupId/members?`$top=1") | Out-Null

    $select = 'id,displayName,userPrincipalName,mail,accountEnabled'
    $membersUri = "https://graph.microsoft.com/v1.0/groups/$BackupPolicyScopeGroupId/transitiveMembers/microsoft.graph.user?`$select=$select&`$top=999"
    Write-ScopeLog -Message "Retrieving transitive user members for group $BackupPolicyScopeGroupId."
    $members = @(Invoke-GraphGetCollection -Uri $membersUri)
    Write-ScopeLog -Message "Group user members retrieved: $($members.Count)"

    $mailboxIndex = Build-MailboxAddressIndex -Mailboxes $mailboxes
    $memberRows = @($members | ForEach-Object { ConvertTo-GroupMemberRow -Member $_ -RunId $runId -RunDateUtc $runDateUtc -TenantName $OrgDomain -GroupId $BackupPolicyScopeGroupId -GroupDisplayName $BackupPolicyScopeGroupDisplayName })
    $coverageRows = @(ConvertTo-CoverageRows -Members $members -MailboxIndex $mailboxIndex -RunId $runId -RunDateUtc $runDateUtc -TenantName $OrgDomain -GroupId $BackupPolicyScopeGroupId -GroupDisplayName $BackupPolicyScopeGroupDisplayName)
    $missingRows = @($coverageRows | Where-Object { $_.MailboxFoundInInventory -ne $true })
    $summaryRows = @(
        [pscustomobject]@{ RunId=$runId; RunDateUtc=$runDateUtc; TenantName=$OrgDomain; GroupId=$BackupPolicyScopeGroupId; GroupDisplayName=$BackupPolicyScopeGroupDisplayName; Metric='GroupUserMembers'; NumericValue=$members.Count; TextValue=[string]$members.Count; Threshold=''; Status='OK'; Details='Transitive user members in expected backup policy scope group.' },
        [pscustomobject]@{ RunId=$runId; RunDateUtc=$runDateUtc; TenantName=$OrgDomain; GroupId=$BackupPolicyScopeGroupId; GroupDisplayName=$BackupPolicyScopeGroupDisplayName; Metric='MatchedMailboxes'; NumericValue=(@($coverageRows | Where-Object { $_.MailboxFoundInInventory -eq $true })).Count; TextValue=[string](@($coverageRows | Where-Object { $_.MailboxFoundInInventory -eq $true })).Count; Threshold=''; Status='OK'; Details='Group users matched to Exchange Online mailbox inventory.' },
        [pscustomobject]@{ RunId=$runId; RunDateUtc=$runDateUtc; TenantName=$OrgDomain; GroupId=$BackupPolicyScopeGroupId; GroupDisplayName=$BackupPolicyScopeGroupDisplayName; Metric='MembersWithoutMailboxInInventory'; NumericValue=$missingRows.Count; TextValue=[string]$missingRows.Count; Threshold='0'; Status=$(if ($missingRows.Count -gt 0) { 'Warning' } else { 'OK' }); Details='Group users not matched to Exchange Online mailbox inventory.' }
    )

    Export-InventoryCsv -BaseName 'M365_BackupPolicyScope_GroupMembers' -Rows $memberRows
    Export-InventoryCsv -BaseName 'M365_BackupPolicyScope_MailboxCoverage' -Rows $coverageRows
    Export-InventoryCsv -BaseName 'M365_BackupPolicyScope_GroupMembersWithoutMailbox' -Rows $missingRows
    Export-InventoryCsv -BaseName 'M365_BackupPolicyScope_Summary' -Rows $summaryRows

    Write-ScopeLog -Message ("Backup policy scope inventory complete. Members={0}; MatchedMailboxes={1}; MissingMailbox={2}" -f $members.Count, (@($coverageRows | Where-Object { $_.MailboxFoundInInventory -eq $true })).Count, $missingRows.Count) -Level SUCCESS
    exit 0
}
catch {
    Write-ScopeLog -Message $_.Exception.Message -Level ERROR
    exit 1
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBg/68OnKaqiwc0
# 6fYy73tLG+EYVTqa8lZumjl3VXd0HKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPAciD5fN3LsfH0SKeVHViAbLrBMuHq4OBKd6V7NiZgxMA0GCSqG
# SIb3DQEBAQUABIIBgHGgzHbJVWnZ4udLi1gf8T2+E9+aLB9MqMXyQzexy+lNLT93
# SbUbjAV6WC4Ce/SljG3zsnh4gABudYqEERUci3Kbmi033CRP88MGsaO/uVx+pmf+
# vgnrYADJ/9//VVXSWSabzu9nQ2SoGXIJpZl+9NN1skT+GnxLFVsjgZ3BdCzgKHXe
# bPvpBy4zHtOH11SeQtWQhgEuOSaSKm0aafsM4+Ik5t0jqwZ7Kl+WprGCKBKeBXid
# m3buWvEHvXtD3oGoQA0Q4Hphgp3M8hJL8LjfW6ujzdT2zIGlD1RMG+AENpIU/Ljw
# Hr1y0YWLZvjTKkD7S7nNBURnYc7PXCBm2OqiSvq3byEWxCllyDuMEtsToQPM6T8F
# QK9tXSLoumSZhqE9UEVjXK4V9tgW7HXp/vQj+ITM5cQNiqxvV8iiAg8hKtEgvIdp
# S/MxfYY9R/598ZXkaQrGzwxL1iYGjWGIq2wJ8hag/02tY5uSAcm9ZN7pDU0EBydh
# vihWUXOoYxwk0EBSsKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MjNaMC8GCSqGSIb3DQEJBDEiBCAqtuLH/PvID2Sn9wbB0e5a58dgCnvsUMdeJ2cy
# 0DMafTANBgkqhkiG9w0BAQEFAASCAgAshny7mwnAlWVMBQdkrKT22SXVjqVmlgrz
# CCjHXdnMpayGUoweHgkJSQg5lFAUq5alHTpKKVyEDTg8YXpc0G6aOOU6MoYYDj7Z
# ZQtpEhxYGWGAucB1LmmuCofmhb3pm2aip/+MYqsz3QGlOdNO59Skdo9JrOzmIbNe
# A82MBDczZUvxAekhbMqNRVOesIV51GuioF1ps1SsQsLJsbaKImoKjq0VIPcVvt/M
# hQkc/eyvEqI2weM0t3GqPFT37WmNi/BwydeSj3a306YJC5VKVjwURBGdeIEsOxEN
# 2ZIwv7gKRhQcbrWUNEoY49DxZWCJDzQrywAtRmiX89HZW6YOQg+lnapGJQUpUgsu
# OJ49UZOanME7QGQ0vxlw4PcDYqm43yFADb7YeUasj801Dn2UyMPlDojltLcYo/a2
# mDaMwcy/0CvGDwAX7cUPWshpTP0KKnXKLjF/3pCXln40VkkhZQ8c/Bp10lImk2fc
# rY3m91bHt9vX9vP3R6J8QKSi8cq9UxzjZ05t5az00w9BRCcFQPioFM6mLjTpKyEJ
# 7b+o2bbpZWSxSx68Od05dU44xAwPBzoNBjnh4c9UkFKV/eUIP8g60xM5D4lLYoVz
# kH0FQAVdClYDobpwbK3kb+u5+u2/5ppNQyYBkIsyJJOD44kO6x6kmL4qjIrDVCwr
# KlBqtMJPIg==
# SIG # End signature block
