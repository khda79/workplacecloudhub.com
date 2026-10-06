<#
.SYNOPSIS
    Inventories SharePoint Online permissions to a dedicated CSV file.

.DESCRIPTION
    Uses PnP.PowerShell with interactive, device login, or certificate authentication.
    SiteUrl and WebUrlsFile inputs are treated as roots: descendant subsites
    are included by default for permission inventory.
    The output columns intentionally match SmartM365-SharePointSource-PermissionInventory.ps1 as closely
    as possible so both inventories can be compared.

.VERSION
    1.1.5
#>

[CmdletBinding(DefaultParameterSetName = 'WebUrlsFile')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Site')]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'WebUrlsFile')]
    [ValidateNotNullOrEmpty()]
    [string]$WebUrlsFile,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [ValidateNotNullOrEmpty()]
    [string]$ErrorPath,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [string]$ClientId = $env:SPO_INVENTORY_CLIENT_ID,

    [string]$Tenant = $env:SPO_INVENTORY_TENANT,

    [string]$TenantId = $env:SPO_INVENTORY_TENANT_ID,

    [string]$Thumbprint = $env:SPO_INVENTORY_THUMBPRINT,

    [switch]$Interactive,

    [switch]$DeviceLogin,

    [switch]$ForceAuthentication,

    [switch]$IncludeHiddenLists,

    [switch]$IncludeSystemLists,

    [switch]$DocumentLibrariesOnly,

    [switch]$IncludeItemPermissions = $true,

    [int]$PageSize = 2000,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$ItemProgressInterval = 500
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'CANCELLED'
try {

function Get-ConsoleTimestamp {
    return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Write-ConsoleMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [System.ConsoleColor]$ForegroundColor
    )

    $line = "{0} {1}" -f (Get-ConsoleTimestamp), $Message
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) {
        Microsoft.PowerShell.Utility\Write-Host $line -ForegroundColor $ForegroundColor
    }
    else {
        Microsoft.PowerShell.Utility\Write-Host $line
    }
}

function Write-ConsoleWarning {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-ConsoleMessage -Message ("WARNING: {0}" -f $Message) -ForegroundColor Yellow
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

    Stop-Transcript | Out-Null
    $script:TranscriptStarted = $false
    Add-TimestampToLogFile -Path $Path
}

$PermissionColumns = @(
    'SiteCollectionUrl',
    'WebUrl',
    'WebTitle',
    'AssociatedMemberGroup',
    'AssociatedOwnerGroup',
    'AssociatedVisitorGroup',
    'ObjectScope',
    'ObjectUrl',
    'ObjectServerRelativeUrl',
    'ObjectTitle',
    'ObjectId',
    'ParentObjectUrl',
    'ListTitle',
    'ListUrl',
    'ListBaseTemplate',
    'ListBaseType',
    'IsDocumentLibrary',
    'ItemId',
    'ItemFileSystemObjectType',
    'HasUniqueRoleAssignments',
    'InheritedFrom',
    'PrincipalType',
    'PrincipalName',
    'PrincipalLoginName',
    'PrincipalId',
    'PrincipalMemberCount',
    'PrincipalUserMemberCount',
    'PrincipalDomainGroupMemberCount',
    'PrincipalMemberLoginNames',
    'PrincipalMemberDisplayNames',
    'PrincipalMemberLookupStatus',
    'PermissionLevels',
    'IsLimitedAccessOnly'
)

function ConvertTo-SafeFileName {
    param(
        [string]$Name,
        [string]$Fallback = 'SharePointOnline'
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = $Fallback
    }

    $invalidCharacters = [regex]::Escape((-join [System.IO.Path]::GetInvalidFileNameChars()))
    $safeName = [regex]::Replace($Name.Trim(), "[$invalidCharacters]+", '-')
    $safeName = [regex]::Replace($safeName, '\s+', ' ')
    $safeName = $safeName.Trim(" .-".ToCharArray())

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = $Fallback
    }

    if ($safeName.Length -gt 80) {
        $safeName = $safeName.Substring(0, 80).Trim(" .-".ToCharArray())
    }

    return $safeName
}

function Format-InventoryDuration {
    param(
        [TimeSpan]$Elapsed
    )

    if ($Elapsed.Days -gt 0) {
        return ("{0}d {1:00}:{2:00}:{3:00}" -f $Elapsed.Days, $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds)
    }

    return ("{0:00}:{1:00}:{2:00}" -f $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds)
}

function Write-ItemPermissionHeartbeat {
    param(
        [string]$WebUrl,
        [string]$ListTitle,
        [int]$ProcessedItems,
        [int]$UniquePermissionItems,
        [int]$ExportedRows,
        [TimeSpan]$Elapsed,
        [string]$LastItem
    )

    Write-ConsoleMessage -Message ("  Item heartbeat: web='{0}' library='{1}' processed={2}; unique={3}; permission rows exported={4}; elapsed={5}; last='{6}'" -f `
            $WebUrl, `
            $ListTitle, `
            $ProcessedItems, `
            $UniquePermissionItems, `
            $ExportedRows, `
            (Format-InventoryDuration -Elapsed $Elapsed), `
            $LastItem) -ForegroundColor DarkCyan
}

function Write-PermissionHeaderOnlyCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $header = ($PermissionColumns | ForEach-Object { '"{0}"' -f ($_ -replace '"', '""') }) -join ';'
    Set-Content -LiteralPath $Path -Value $header -Encoding UTF8
}

function Export-PermissionRows {
    param(
        [object[]]$Rows,
        [string]$CsvPath
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        return
    }

    if ($script:CsvCreated) {
        $Rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $Rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8
        $script:CsvCreated = $true
    }
}

function Write-InventoryError {
    param(
        [string]$Scope,
        [string]$Url,
        [string]$Name,
        [string]$Message,
        [string]$ItemId = '',
        [string]$ItemUrl = ''
    )

    if ([string]::IsNullOrWhiteSpace($script:ErrorPath)) {
        return
    }

    $row = [pscustomobject]@{
        Time    = Get-Date
        Scope   = $Scope
        Url     = $Url
        Name    = $Name
        Message = $Message
        ItemId  = $ItemId
        ItemUrl = $ItemUrl
    }

    if ($script:ErrorCsvCreated) {
        $row | Export-Csv -Delimiter ';' -Path $script:ErrorPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $row | Export-Csv -Delimiter ';' -Path $script:ErrorPath -NoTypeInformation -Encoding UTF8
        $script:ErrorCsvCreated = $true
    }
}

function Import-RequiredModule {
    $module = Get-Module -ListAvailable -Name 'PnP.PowerShell' |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $module) {
        throw "Required module 'PnP.PowerShell' is not installed."
    }

    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-ConsoleMessage -Message ("PnP.PowerShell version: {0}" -f $module.Version)
    Write-ConsoleMessage -Message ("PnP.PowerShell path: {0}" -f $module.Path)
}

function Connect-ToSPOWeb {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    Write-ConsoleMessage -Message ("Connecting to: {0}" -f $Url)

    $parameters = @{
        Url         = $Url
        ErrorAction = 'Stop'
        ReturnConnection = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
        $parameters.ClientId = $ClientId
    }

    if (-not [string]::IsNullOrWhiteSpace($Tenant)) {
        $parameters.Tenant = $Tenant
    }

    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $parameters.Tenant = $TenantId
    }

    if ($Interactive) {
        $parameters.Interactive = $true
    }
    elseif ($DeviceLogin) {
        $parameters.DeviceLogin = $true
    }
    elseif (-not [string]::IsNullOrWhiteSpace($Thumbprint)) {
        $parameters.Thumbprint = $Thumbprint
    }
    else {
        $parameters.Interactive = $true
    }

    if ($ForceAuthentication -and -not $script:ForceAuthenticationUsed -and -not $parameters.ContainsKey('Thumbprint')) {
        try {
            Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction SilentlyContinue
            Write-ConsoleMessage -Message "Cleared persisted PnP login before forced authentication."
        }
        catch {
            Write-ConsoleWarning -Message ("Could not clear persisted PnP login: {0}" -f $_.Exception.Message)
        }

        $parameters.ForceAuthentication = $true
        $script:ForceAuthenticationUsed = $true
    }

    $connection = Connect-PnPOnline @parameters
    $script:SPOPermissionConnection = $connection
    Write-SPOConnectionIdentity -Connection $connection -Url $Url
    $tokenContext = '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f ([uri]$Url).Host, $Tenant, $TenantId, $ClientId, $Thumbprint, [bool]$DeviceLogin, $script:SPOConnectedAccount
    if ($script:TokenSummaryContexts.Add($tokenContext)) { Write-PnPTokenSummary -Connection $connection }
}

function Write-SPOConnectionIdentity {
    param(
        $Connection,
        [string]$Url
    )

    $script:SPOConnectedAccount = $Url
    try {
        $context = Get-PnPContext -Connection $Connection
        $context.Load($context.Web.CurrentUser)
        $context.ExecuteQuery()

        $currentUser = $context.Web.CurrentUser
        $script:SPOConnectedAccount = [string]$currentUser.LoginName
        $loginName = if ($currentUser.LoginName) { $currentUser.LoginName } else { '<unknown>' }
        $email = if ($currentUser.Email) { $currentUser.Email } else { '<no email>' }
        $title = if ($currentUser.Title) { $currentUser.Title } else { '<no title>' }

        Write-ConsoleMessage -Message ("Connected account for {0}: {1} | {2} | {3}" -f $Url, $loginName, $email, $title) -ForegroundColor Cyan
    }
    catch {
        Write-ConsoleWarning -Message ("Could not determine connected account for '{0}': {1}" -f $Url, $_.Exception.Message)
    }
}

function ConvertFrom-Base64Url {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $base64 = $Value.Replace('-', '+').Replace('_', '/')
    while ($base64.Length % 4 -ne 0) {
        $base64 += '='
    }

    return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
}

function Write-PnPTokenSummary {
    param(
        $Connection
    )

    try {
        $sharePointScopes = @(Get-PnPAccessToken -ResourceTypeName SharePoint -ListPermissionScopes -Connection $Connection -ErrorAction Stop)
        if ($sharePointScopes.Count -gt 0) {
            Write-ConsoleMessage -Message ("PnP SharePoint token scopes: {0}" -f ($sharePointScopes -join ' ')) -ForegroundColor DarkCyan
        }
        else {
            Write-ConsoleMessage -Message "PnP SharePoint token scopes: <none returned>" -ForegroundColor DarkYellow
        }
    }
    catch {
        Write-ConsoleWarning -Message ("Could not list SharePoint token scopes: {0}" -f $_.Exception.Message)
    }

    try {
        if ($null -eq $Connection -or [string]::IsNullOrWhiteSpace($Connection.AccessToken)) {
            Write-ConsoleMessage -Message "PnP token details: <not available>" -ForegroundColor DarkYellow
            return
        }

        $parts = $Connection.AccessToken.Split('.')
        if ($parts.Count -lt 2) {
            Write-ConsoleMessage -Message "PnP token details: <unrecognized token format>" -ForegroundColor DarkYellow
            return
        }

        $payload = ConvertFrom-Json -InputObject (ConvertFrom-Base64Url -Value $parts[1])
        $appId = if ($payload.appid) { $payload.appid } elseif ($payload.azp) { $payload.azp } else { '<none>' }
        $scopes = if ($payload.scp) { $payload.scp } else { '<none>' }
        $roles = if ($payload.roles) { ($payload.roles -join ' ') } else { '<none>' }

        Write-ConsoleMessage -Message ("PnP token app id: {0}" -f $appId) -ForegroundColor DarkCyan
        Write-ConsoleMessage -Message ("PnP token scopes: {0}" -f $scopes) -ForegroundColor DarkCyan
        Write-ConsoleMessage -Message ("PnP token roles: {0}" -f $roles) -ForegroundColor DarkCyan
    }
    catch {
        Write-ConsoleWarning -Message ("Could not decode PnP token details: {0}" -f $_.Exception.Message)
    }
}

function Get-SiteCollectionUrlFromWebUrl {
    param(
        [string]$WebUrl
    )

    $uri = [System.Uri]$WebUrl
    $segments = @($uri.AbsolutePath.Trim('/').Split('/') | Where-Object { $_ })
    if ($segments.Count -ge 2 -and $segments[0].Equals('sites', [System.StringComparison]::OrdinalIgnoreCase)) {
        return ("{0}://{1}/sites/{2}" -f $uri.Scheme, $uri.Authority, $segments[1])
    }

    return ("{0}://{1}" -f $uri.Scheme, $uri.Authority)
}

function ConvertTo-AbsoluteSharePointUrl {
    param(
        [string]$WebUrl,
        [string]$ServerRelativeUrl
    )

    if ([string]::IsNullOrWhiteSpace($ServerRelativeUrl)) {
        return $null
    }

    if ($ServerRelativeUrl -match '^https?://') {
        return $ServerRelativeUrl
    }

    if (-not $ServerRelativeUrl.StartsWith('/')) {
        $ServerRelativeUrl = "/$ServerRelativeUrl"
    }

    $uri = [System.Uri]$WebUrl
    return ("{0}://{1}{2}" -f $uri.Scheme, $uri.Authority, $ServerRelativeUrl)
}

function Test-SystemList {
    param(
        $List
    )

    $rootFolder = Get-PnPProperty -ClientObject $List -Property RootFolder -Connection $script:SPOPermissionConnection
    $systemUrls = @(
        '_catalogs/masterpage',
        '_catalogs/wp',
        '_catalogs/lt',
        'Style Library',
        'FormServerTemplates',
        'PreservationHoldLibrary',
        'Site Collection Documents',
        'Site Collection Images'
    )

    foreach ($url in $systemUrls) {
        if ($rootFolder.ServerRelativeUrl -like "*/$url*" -or $rootFolder.Name -eq $url) {
            return $true
        }
    }

    return $false
}

function Get-PrincipalInfo {
    param(
        $Member
    )

    $principalType = $Member.PrincipalType
    $name = $Member.Title
    $loginName = $Member.LoginName
    $id = $Member.Id

    [pscustomobject]@{
        PrincipalType      = $principalType
        PrincipalName      = $name
        PrincipalLoginName = $loginName
        PrincipalId        = $id
    }
}

function Get-PnPGroupTitle {
    param(
        $Group
    )

    if ($null -eq $Group) {
        return ''
    }

    # A single-property request returns its value, not the original group object.
    return [string](Get-PnPProperty -ClientObject $Group -Property Title -ErrorAction Stop -Connection $script:SPOPermissionConnection)
}

function Get-AssociatedWebGroupNames {
    param(
        $Web
    )

    $key = [string]$Web.Url
    if ($script:AssociatedWebGroupCache.ContainsKey($key)) { return $script:AssociatedWebGroupCache[$key] }
    $names = [ordered]@{}
    foreach ($property in @('AssociatedMemberGroup', 'AssociatedOwnerGroup', 'AssociatedVisitorGroup')) {
        $names[$property] = ''
        try {
            $names[$property] = Get-PnPGroupTitle -Group (Get-PnPProperty -ClientObject $Web -Property $property -ErrorAction Stop -Connection $script:SPOPermissionConnection)
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to read {0} for '{1}': {2}" -f $property, $key, $_.Exception.Message)
            Write-InventoryError -Scope 'AssociatedGroup' -Url $key -Name $property -Message $_.Exception.Message
        }
    }
    $result = [pscustomobject]$names
    $script:AssociatedWebGroupCache[$key] = $result
    return $result
}

$script:AssociatedWebGroupCache = @{}
$script:TokenSummaryContexts = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:SPOInheritanceLoad = $null
$script:SPOInheritanceItemType = $null
$script:SharePointGroupMembershipCache = @{}

function Join-PrincipalMemberValues {
    param(
        [object[]]$Values
    )

    if (-not $Values -or $Values.Count -eq 0) {
        return ''
    }

    return (($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique) -join ' || ')
}

function Get-EmptyPrincipalMembershipInfo {
    param(
        [string]$Status = ''
    )

    [pscustomobject]@{
        PrincipalMemberCount            = ''
        PrincipalUserMemberCount        = ''
        PrincipalDomainGroupMemberCount = ''
        PrincipalMemberLoginNames       = ''
        PrincipalMemberDisplayNames     = ''
        PrincipalMemberLookupStatus     = $Status
    }
}

function Get-PrincipalMembershipInfo {
    param(
        $Member,
        [string]$PrincipalType
    )

    if ($PrincipalType -ne 'SharePointGroup') {
        return Get-EmptyPrincipalMembershipInfo
    }

    $groupIdentity = if (-not [string]::IsNullOrWhiteSpace([string]$Member.Title)) { [string]$Member.Title } else { [string]$Member.LoginName }
    if ([string]::IsNullOrWhiteSpace($groupIdentity)) {
        return Get-EmptyPrincipalMembershipInfo -Status 'Skipped: group identity is empty'
    }

    $cacheKey = if ($null -ne $Member.Id) { [string]$Member.Id } else { $groupIdentity.ToLowerInvariant() }
    if ($script:SharePointGroupMembershipCache.ContainsKey($cacheKey)) {
        return $script:SharePointGroupMembershipCache[$cacheKey]
    }

    try {
        $members = @(Get-PnPGroupMember -Group $groupIdentity -ErrorAction Stop -Connection $script:SPOPermissionConnection)
        $loginNames = @($members | ForEach-Object { $_.LoginName })
        $displayNames = @($members | ForEach-Object { $_.Title })
        $domainGroupMembers = @($members | Where-Object { [string]$_.PrincipalType -match 'SecurityGroup|DistributionList|SharePointGroup' })
        $userMembers = @($members | Where-Object { [string]$_.PrincipalType -eq 'User' })

        $info = [pscustomobject]@{
            PrincipalMemberCount            = $members.Count
            PrincipalUserMemberCount        = $userMembers.Count
            PrincipalDomainGroupMemberCount = $domainGroupMembers.Count
            PrincipalMemberLoginNames       = Join-PrincipalMemberValues -Values $loginNames
            PrincipalMemberDisplayNames     = Join-PrincipalMemberValues -Values $displayNames
            PrincipalMemberLookupStatus     = 'OK'
        }
    }
    catch {
        $info = Get-EmptyPrincipalMembershipInfo -Status ("Failed: {0}" -f $_.Exception.Message)
    }

    $script:SharePointGroupMembershipCache[$cacheKey] = $info
    return $info
}
function Get-RoleAssignmentRows {
    param(
        [string]$SiteCollectionUrl,
        [string]$WebUrl,
        [string]$WebTitle,
        [string]$AssociatedMemberGroup,
        [string]$AssociatedOwnerGroup,
        [string]$AssociatedVisitorGroup,
        [string]$ObjectScope,
        [string]$ObjectUrl,
        [string]$ObjectServerRelativeUrl,
        [string]$ObjectTitle,
        [string]$ObjectId,
        [string]$ParentObjectUrl,
        [string]$ListTitle,
        [string]$ListUrl,
        [object]$ListBaseTemplate,
        [string]$ListBaseType,
        [object]$IsDocumentLibrary,
        [object]$ItemId,
        [string]$ItemFileSystemObjectType,
        [bool]$HasUniqueRoleAssignments,
        [string]$InheritedFrom,
        $RoleAssignments
    )

    foreach ($roleAssignment in $RoleAssignments) {
        try {
            $member = Get-PnPProperty -ClientObject $roleAssignment -Property Member -Connection $script:SPOPermissionConnection
            $bindings = @(Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings -Connection $script:SPOPermissionConnection)
            $permissionLevels = @($bindings | ForEach-Object { $_.Name })
            if ($permissionLevels.Count -eq 0) {
                continue
            }

            $principal = Get-PrincipalInfo -Member $member
            $membership = Get-PrincipalMembershipInfo -Member $member -PrincipalType $principal.PrincipalType
            [pscustomobject]@{
                SiteCollectionUrl        = $SiteCollectionUrl
                WebUrl                   = $WebUrl
                WebTitle                 = $WebTitle
                AssociatedMemberGroup    = $AssociatedMemberGroup
                AssociatedOwnerGroup     = $AssociatedOwnerGroup
                AssociatedVisitorGroup   = $AssociatedVisitorGroup
                ObjectScope              = $ObjectScope
                ObjectUrl                = $ObjectUrl
                ObjectServerRelativeUrl  = $ObjectServerRelativeUrl
                ObjectTitle              = $ObjectTitle
                ObjectId                 = $ObjectId
                ParentObjectUrl          = $ParentObjectUrl
                ListTitle                = $ListTitle
                ListUrl                  = $ListUrl
                ListBaseTemplate        = $ListBaseTemplate
                ListBaseType            = $ListBaseType
                IsDocumentLibrary       = $IsDocumentLibrary
                ItemId                   = $ItemId
                ItemFileSystemObjectType = $ItemFileSystemObjectType
                HasUniqueRoleAssignments = $HasUniqueRoleAssignments
                InheritedFrom            = $InheritedFrom
                PrincipalType            = $principal.PrincipalType
                PrincipalName            = $principal.PrincipalName
                PrincipalLoginName       = $principal.PrincipalLoginName
                PrincipalId              = $principal.PrincipalId
                PrincipalMemberCount         = $membership.PrincipalMemberCount
                PrincipalUserMemberCount     = $membership.PrincipalUserMemberCount
                PrincipalDomainGroupMemberCount = $membership.PrincipalDomainGroupMemberCount
                PrincipalMemberLoginNames    = $membership.PrincipalMemberLoginNames
                PrincipalMemberDisplayNames  = $membership.PrincipalMemberDisplayNames
                PrincipalMemberLookupStatus  = $membership.PrincipalMemberLookupStatus
                PermissionLevels         = ($permissionLevels -join '|')
                IsLimitedAccessOnly      = ($permissionLevels.Count -eq 1 -and $permissionLevels[0] -eq 'Limited Access')
            }
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to read role assignment on '{0}': {1}" -f $ObjectUrl, $_.Exception.Message)
            Write-InventoryError -Scope "$ObjectScope RoleAssignment" -Url $ObjectUrl -Name $ObjectTitle -ItemId ([string]$ItemId) -ItemUrl $(if ($ObjectScope -eq 'Item') { $ObjectUrl } else { '' }) -Message $_.Exception.Message
        }
    }
}

function Export-ItemPermissionInventory {
    param(
        $Web,
        $List,
        [string]$CsvPath,
        [int]$ProgressInterval
    )

    $associatedGroups = Get-AssociatedWebGroupNames -Web $Web
    $rootFolder = Get-PnPProperty -ClientObject $List -Property RootFolder -Connection $script:SPOPermissionConnection
    $listUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl $rootFolder.ServerRelativeUrl
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $state = @{ Processed=0; Unique=0; Exported=0; LastItem='<none>' }
    $script:SPOPermissionPageState = $state
    $script:SPOPermissionPageWeb = $Web
    $script:SPOPermissionPageList = $List
    $script:SPOPermissionPageCsv = $CsvPath
    $script:SPOPermissionPageInterval = $ProgressInterval
    $script:SPOPermissionPageStopwatch = $stopwatch
    try {
        # Process the callback before the next page is fetched; do not retain the whole library.
        Get-PnPListItem -List $List -PageSize $PageSize -Fields 'FileRef','FileLeafRef','FSObjType','UniqueId','ID' -Connection $script:SPOPermissionConnection -ScriptBlock {
            param($PageItems)
            Export-SPOPermissionPage -PageItems $PageItems -Web $script:SPOPermissionPageWeb -List $script:SPOPermissionPageList -CsvPath $script:SPOPermissionPageCsv -State $script:SPOPermissionPageState -ProgressInterval $script:SPOPermissionPageInterval -Stopwatch $script:SPOPermissionPageStopwatch
        } -ErrorAction Stop | Out-Null
    }
    catch {
        Write-ConsoleWarning -Message ("Failed to enumerate items for list '{0}' in web '{1}': {2}" -f $List.Title, $Web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'ListItems' -Url $listUrl -Name $List.Title -Message $_.Exception.Message
    }
    finally { $stopwatch.Stop() }
    Write-ItemPermissionHeartbeat -WebUrl $Web.Url -ListTitle $List.Title -ProcessedItems $state.Processed -UniquePermissionItems $state.Unique -ExportedRows $state.Exported -Elapsed $stopwatch.Elapsed -LastItem $state.LastItem
}

function Export-SPOPermissionPage {
    param($PageItems, $Web, $List, [string]$CsvPath, [hashtable]$State, [int]$ProgressInterval, $Stopwatch)
    $associatedGroups = Get-AssociatedWebGroupNames -Web $Web
    $rootFolder = Get-PnPProperty -ClientObject $List -Property RootFolder -Connection $script:SPOPermissionConnection
    $listUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl $rootFolder.ServerRelativeUrl
    $inheritance = Get-SPOPageInheritance -PageItems $PageItems -Web $Web -List $List -ListUrl $listUrl
    foreach ($item in $PageItems) {
        $State.Processed++
        $State.LastItem = "ID $($item.Id)"
        try {
            if ($State.Processed % $ProgressInterval -eq 0) {
                Write-ItemPermissionHeartbeat -WebUrl $Web.Url -ListTitle $List.Title -ProcessedItems $State.Processed -UniquePermissionItems $State.Unique -ExportedRows $State.Exported -Elapsed $Stopwatch.Elapsed -LastItem $State.LastItem
            }
            if (-not $inheritance.ContainsKey([int]$item.Id) -or -not $inheritance[[int]$item.Id]) { continue }
            $State.Unique++
            $roleAssignments = Get-PnPProperty -ClientObject $item -Property RoleAssignments -Connection $script:SPOPermissionConnection
            $serverRelativeUrl = [string]$item.FieldValues.FileRef
            $absoluteUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl $serverRelativeUrl
            $fsObjType = if ([string]$item.FieldValues.FSObjType -eq '1') { 'Folder' } else { 'FileOrItem' }
            $itemName = [string]$item.FieldValues.FileLeafRef
            if ([string]::IsNullOrWhiteSpace($itemName)) {
                $itemName = [string]$item.Id
            }

            $rows = @(Get-RoleAssignmentRows `
                    -SiteCollectionUrl (Get-SiteCollectionUrlFromWebUrl -WebUrl $Web.Url) `
                    -WebUrl $Web.Url `
                    -WebTitle $Web.Title `
                    -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                    -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                    -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                    -ObjectScope 'Item' `
                    -ObjectUrl $absoluteUrl `
                    -ObjectServerRelativeUrl $serverRelativeUrl `
                    -ObjectTitle $itemName `
                    -ObjectId ([string]$item.FieldValues.UniqueId) `
                    -ParentObjectUrl $listUrl `
                    -ListTitle $List.Title `
                    -ListUrl $listUrl `
                    -ListBaseTemplate ([int]$List.BaseTemplate) `
                    -ListBaseType $(if ([string]$List.BaseType -eq 'DocumentLibrary') { 'DocumentLibrary' } else { '' }) `
                    -IsDocumentLibrary ([string]$List.BaseType -eq 'DocumentLibrary') `
                    -ItemId $item.Id `
                    -ItemFileSystemObjectType $fsObjType `
                    -HasUniqueRoleAssignments $true `
                    -InheritedFrom '' `
                    -RoleAssignments $roleAssignments)

            Export-PermissionRows -Rows $rows -CsvPath $CsvPath
            $State.Exported += $rows.Count
        }
        catch {
            Write-SPOItemError -Item $item -Web $Web -List $List -ListUrl $listUrl -Message $_.Exception.Message
        }
    }
}

function Add-SPOInheritanceRead {
    param($Context, $Item)
    # Queue only HasUniqueRoleAssignments through the supported CSOM Load<T> API.
    $itemType = $Item.GetType()
    if (-not $script:SPOInheritanceLoad -or $script:SPOInheritanceItemType -ne $itemType) {
        $parameter = [System.Linq.Expressions.Expression]::Parameter($itemType, 'item')
        $property = [System.Linq.Expressions.Expression]::Property($parameter, 'HasUniqueRoleAssignments')
        $boxed = [System.Linq.Expressions.Expression]::Convert($property, [object])
        $delegateType = [System.Func`2].MakeGenericType($itemType, [object])
        $lambda = [System.Linq.Expressions.Expression]::Lambda($delegateType, $boxed, [System.Linq.Expressions.ParameterExpression[]]@($parameter))
        $expressionType = [System.Linq.Expressions.Expression`1].MakeGenericType($delegateType)
        $script:SPOInheritanceSelectors = [Array]::CreateInstance($expressionType, 1)
        $script:SPOInheritanceSelectors.SetValue($lambda, 0)
        $method = @($Context.GetType().GetMethods() | Where-Object { $_.Name -eq 'Load' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 2 })[0]
        $script:SPOInheritanceLoad = $method.MakeGenericMethod($itemType)
        $script:SPOInheritanceItemType = $itemType
    }
    [void]$script:SPOInheritanceLoad.Invoke($Context, [object[]]@($Item, $script:SPOInheritanceSelectors))
}

function Write-SPOItemError {
    param($Item, $Web, $List, [string]$ListUrl, [string]$Message)
    $itemUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl ([string]$Item.FieldValues.FileRef)
    if ($Message -match '(?i)item does not exist|does not exist.*item') {
        try {
            $probe = Get-PnPListItem -List $List -Id $Item.Id -Fields 'FileRef' -Connection $script:SPOPermissionConnection -ErrorAction Stop
            $Message += $(if ($null -eq $probe) { ' Existence check: no item returned.' } else { ' Existence check: item still exists; permission read failed.' })
        }
        catch { $Message += ' Existence check failed: ' + $_.Exception.Message }
    }
    Write-ConsoleWarning -Message ("Failed item permission read: list='{0}'; ID={1}; path='{2}'; {3}" -f $List.Title, $Item.Id, $itemUrl, $Message)
    Write-InventoryError -Scope 'Item' -Url $ListUrl -Name $List.Title -ItemId ([string]$Item.Id) -ItemUrl $itemUrl -Message $Message
}

function Get-SPOPageInheritance {
    param($PageItems, $Web, $List, [string]$ListUrl)
    $context = Get-PnPContext -Connection $script:SPOPermissionConnection
    $result = @{}
    # A 2,000-item page can exceed SPO's 2 MiB CSOM request limit. Keep the
    # enumeration page size, but load inheritance in smaller independent requests.
    $items = @($PageItems)
    $batchSize = 100
    for ($offset = 0; $offset -lt $items.Count; $offset += $batchSize) {
        $last = [Math]::Min($offset + $batchSize - 1, $items.Count - 1)
        $batch = @($items[$offset..$last])
        $label = "inheritance batch in {0} ({1} items; IDs {2}-{3})" -f $List.Title, $batch.Count, $batch[0].Id, $batch[-1].Id
        try {
            Invoke-SPORead -Label $label -Operation {
                foreach ($item in $batch) { Add-SPOInheritanceRead -Context $context -Item $item }
                $context.ExecuteQuery()
            } | Out-Null
            foreach ($item in $batch) { $result[[int]$item.Id] = [bool]$item.HasUniqueRoleAssignments }
        }
        catch {
            if ($_.Exception.Message -notmatch '(?i)item does not exist|does not exist.*item') {
                throw ("Failed {0}: {1}" -f $label, $_.Exception.Message)
            }
            # Isolate missing reads in this batch; continue collecting later batches.
            foreach ($item in $batch) {
                try { $result[[int]$item.Id] = [bool](Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments -ErrorAction Stop -Connection $script:SPOPermissionConnection) }
                catch { Write-SPOItemError -Item $item -Web $Web -List $List -ListUrl $ListUrl -Message $_.Exception.Message }
            }
        }
    }
    return $result
}

function Export-WebPermissionInventory {
    param(
        [string]$WebUrl,
        [string]$CsvPath
    )

    Connect-ToSPOWeb -Url $WebUrl
    $web = Invoke-SPORead -Label 'web' -Operation { Get-PnPWeb -Includes Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments -ErrorAction Stop -Connection $script:SPOPermissionConnection }
    $siteCollectionUrl = Get-SiteCollectionUrlFromWebUrl -WebUrl $web.Url

    Write-ConsoleMessage -Message ("Web: {0}" -f $web.Url)
    $associatedGroups = Get-AssociatedWebGroupNames -Web $web

    try {
        $webRoleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments -Connection $script:SPOPermissionConnection
        $rows = @(Get-RoleAssignmentRows `
                -SiteCollectionUrl $siteCollectionUrl `
                -WebUrl $web.Url `
                -WebTitle $web.Title `
                -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                -ObjectScope 'Web' `
                -ObjectUrl $web.Url `
                -ObjectServerRelativeUrl $web.ServerRelativeUrl `
                -ObjectTitle $web.Title `
                -ObjectId ([string]$web.Id) `
                -ParentObjectUrl '' `
                -ListTitle '' `
                -ListUrl '' `
                -ListBaseTemplate '' `
                -ListBaseType '' `
                -IsDocumentLibrary '' `
                -ItemId $null `
                -ItemFileSystemObjectType '' `
                -HasUniqueRoleAssignments $web.HasUniqueRoleAssignments `
                -InheritedFrom $(if ($web.HasUniqueRoleAssignments) { '' } else { 'ParentWeb' }) `
                -RoleAssignments $webRoleAssignments)
        Export-PermissionRows -Rows $rows -CsvPath $CsvPath
    }
    catch {
        Write-ConsoleWarning -Message ("Failed to inventory web permissions '{0}': {1}" -f $web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'Web' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
    }

    try {
        $lists = Invoke-SPORead -Label 'lists' -Operation { Get-PnPList -Includes Title,Hidden,BaseTemplate,BaseType,Id,RootFolder,HasUniqueRoleAssignments -ErrorAction Stop -Connection $script:SPOPermissionConnection }
    }
    catch {
        Write-ConsoleWarning -Message ("Failed to enumerate lists for web '{0}': {1}" -f $web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'WebLists' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
        return
    }

    foreach ($list in $lists) {
        $listTitle = '<unknown>'
        $listUrl = $web.Url

        try {
            $rootFolder = Get-PnPProperty -ClientObject $list -Property RootFolder -Connection $script:SPOPermissionConnection
            $listTitle = $list.Title
            $listUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $web.Url -ServerRelativeUrl $rootFolder.ServerRelativeUrl

            if (-not $IncludeHiddenLists -and $list.Hidden) {
                Write-ConsoleMessage -Message ("  Skipping list/library '{0}' ({1}): hidden; use -IncludeHiddenLists to include it." -f $listTitle, $listUrl)
                continue
            }

            if (-not $IncludeSystemLists -and (Test-SystemList -List $list)) {
                Write-ConsoleMessage -Message ("  Skipping list/library '{0}' ({1}): system; use -IncludeSystemLists to include it." -f $listTitle, $listUrl)
                continue
            }

            if ($DocumentLibrariesOnly -and [string]$list.BaseType -ne 'DocumentLibrary') {
                Write-ConsoleMessage -Message ("  Skipping list '{0}' ({1}): not a document library; -DocumentLibrariesOnly is enabled." -f $listTitle, $listUrl)
                continue
            }

            $isDocumentLibrary = ([string]$list.BaseType -eq 'DocumentLibrary')
            $listBaseType = if ($isDocumentLibrary) { 'DocumentLibrary' } else { '' }
            $roleAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments -Connection $script:SPOPermissionConnection
            $rows = @(Get-RoleAssignmentRows `
                    -SiteCollectionUrl $siteCollectionUrl `
                    -WebUrl $web.Url `
                    -WebTitle $web.Title `
                    -AssociatedMemberGroup $associatedGroups.AssociatedMemberGroup `
                    -AssociatedOwnerGroup $associatedGroups.AssociatedOwnerGroup `
                    -AssociatedVisitorGroup $associatedGroups.AssociatedVisitorGroup `
                    -ObjectScope 'List' `
                    -ObjectUrl $listUrl `
                    -ObjectServerRelativeUrl $rootFolder.ServerRelativeUrl `
                    -ObjectTitle $list.Title `
                    -ObjectId ([string]$list.Id) `
                    -ParentObjectUrl $web.Url `
                    -ListTitle $list.Title `
                    -ListUrl $listUrl `
                    -ListBaseTemplate ([int]$list.BaseTemplate) `
                    -ListBaseType $listBaseType `
                    -IsDocumentLibrary $isDocumentLibrary `
                    -ItemId $null `
                    -ItemFileSystemObjectType '' `
                    -HasUniqueRoleAssignments $list.HasUniqueRoleAssignments `
                    -InheritedFrom $(if ($list.HasUniqueRoleAssignments) { '' } else { 'Web' }) `
                    -RoleAssignments $roleAssignments)
            Export-PermissionRows -Rows $rows -CsvPath $CsvPath

            if ($IncludeItemPermissions) {
                Write-ConsoleMessage -Message ("  Item permissions: {0}" -f $list.Title)
                Export-ItemPermissionInventory -Web $web -List $list -CsvPath $CsvPath -ProgressInterval $ItemProgressInterval
            }
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to inventory list permissions '{0}' in web '{1}': {2}" -f $listTitle, $web.Url, $_.Exception.Message)
            Write-InventoryError -Scope 'List' -Url $listUrl -Name $listTitle -Message $_.Exception.Message
        }
    }
}

function Get-DefaultOutputPath {
    $scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $targetName = if ($PSCmdlet.ParameterSetName -eq 'Site') {
        ConvertTo-SafeFileName -Name ([System.Uri]$SiteUrl).Segments[-1].Trim('/')
    }
    else {
        ConvertTo-SafeFileName -Name ([System.IO.Path]::GetFileNameWithoutExtension($WebUrlsFile))
    }

    $targetDirectory = Join-Path -Path $scriptDirectory -ChildPath $targetName
    return Join-Path -Path $targetDirectory -ChildPath ("SPO-PermissionInventory-{0}-{1:yyyyMMdd-HHmmss}.csv" -f $targetName, (Get-Date))
}

function Get-WebUrlsFromFile {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Web URLs file not found: $Path"
    }

    Get-Content -LiteralPath $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        Sort-Object -Unique
}


function Write-SPOReadRetry { param([string]$Message) Write-ConsoleWarning -Message $Message }

function Invoke-SPORead {
    param([scriptblock]$Operation, [string]$Label, [int]$Attempts = 3)
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            # Buffer each attempt so a partial response is never emitted twice.
            $result = @(& $Operation)
            return $result
        }
        catch {
            $message = $_.Exception.ToString()
            $transient = $message -match '(?i)HttpClient.Timeout|timed? out|timeout|\b429\b|\b503\b|TooManyRequests|temporarily unavailable|connection.*(closed|reset)'
            if (-not $transient -or $attempt -eq $Attempts) { throw }
            $delay = if ($attempt -eq 1) { 5 } else { 15 }
            Write-SPOReadRetry -Message ("Retrying {0} after a transient read failure ({1}/{2}); waiting {3}s: {4}" -f $Label, $attempt, $Attempts, $delay, $_.Exception.Message)
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-WebUrlsFromSite {
    param(
        [string]$Url
    )

    $webUrls = New-Object System.Collections.Generic.List[string]
    $pendingWebUrls = New-Object System.Collections.Generic.Queue[string]
    $seenWebUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $pendingWebUrls.Enqueue($Url)

    while ($pendingWebUrls.Count -gt 0) {
        $currentWebUrl = $pendingWebUrls.Dequeue()
        if (-not $seenWebUrls.Add($currentWebUrl)) {
            continue
        }

        try {
            Connect-ToSPOWeb -Url $currentWebUrl
            $web = Invoke-SPORead -Label 'web' -Operation { Get-PnPWeb -Includes Title,Url,ServerRelativeUrl -ErrorAction Stop -Connection $script:SPOPermissionConnection }
            $webUrls.Add($web.Url)

            try {
                $subWebs = @(Invoke-SPORead -Label 'subsites' -Operation { Get-PnPSubWeb -Includes Title,Url,ServerRelativeUrl -ErrorAction Stop -Connection $script:SPOPermissionConnection })
                Write-ConsoleMessage -Message ("  Subsites found under {0}: {1}" -f $web.Url, $subWebs.Count)

                foreach ($subWeb in $subWebs) {
                    if ($subWeb.Url -and -not $seenWebUrls.Contains($subWeb.Url)) {
                        $pendingWebUrls.Enqueue($subWeb.Url)
                    }
                }
            }
            catch {
                Write-ConsoleWarning -Message ("Failed to enumerate immediate subsites for web '{0}': {1}" -f $web.Url, $_.Exception.Message)
                Write-InventoryError -Scope 'SubsiteEnumeration' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
            }
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to connect to or read web '{0}': {1}" -f $currentWebUrl, $_.Exception.Message)
            Write-InventoryError -Scope 'Web' -Url $currentWebUrl -Name $currentWebUrl -Message $_.Exception.Message
        }
    }

    $webUrls
}

Import-RequiredModule

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Get-DefaultOutputPath
}

if ([string]::IsNullOrWhiteSpace($ErrorPath)) {
    $errorBaseDirectory = Split-Path -Path $OutputPath -Parent
    $ErrorPath = Join-Path -Path $errorBaseDirectory -ChildPath ("{0}-Errors.csv" -f [System.IO.Path]::GetFileNameWithoutExtension($OutputPath))
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $logBaseDirectory = Split-Path -Path $OutputPath -Parent
    $LogPath = Join-Path -Path (Join-Path -Path $logBaseDirectory -ChildPath 'logs') -ChildPath ("{0}-Run.log" -f [System.IO.Path]::GetFileNameWithoutExtension($OutputPath))
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$TempOutputPath = "{0}.tmp" -f $OutputPath

foreach ($path in @($ErrorPath, $LogPath)) {
    $directory = Split-Path -Path $path -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
}

foreach ($path in @($OutputPath, $TempOutputPath, $ErrorPath, $LogPath)) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

$script:ErrorPath = $ErrorPath
$script:CsvCreated = $false
$script:ErrorCsvCreated = $false
$script:TranscriptStarted = $false
$script:InventoryStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$script:ForceAuthenticationUsed = $false

try {
    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
    Write-ConsoleMessage -Message ("Permission inventory output: {0}" -f $OutputPath)
    Write-ConsoleMessage -Message ("Temporary permission inventory output: {0}" -f $TempOutputPath)
    Write-ConsoleMessage -Message ("Error output: {0}" -f $ErrorPath)
    Write-ConsoleMessage -Message ("Run log: {0}" -f $LogPath)
    $permissionScope = if ($DocumentLibrariesOnly) { 'DocumentLibrariesOnly' } else { 'AllListsAndLibraries' }
    $authMode = if ($Interactive) { 'Interactive' } elseif ($DeviceLogin) { 'DeviceLogin' } elseif (-not [string]::IsNullOrWhiteSpace($Thumbprint)) { 'Certificate' } else { 'Interactive' }
    $inputScope = if ($PSCmdlet.ParameterSetName -eq 'Site') { $SiteUrl } else { $WebUrlsFile }
    Write-ConsoleMessage -Message ("Permission scan options: Scope={0}; IncludeItemPermissions={1}; IncludeHiddenLists={2}; IncludeSystemLists={3}; PageSize={4}; ItemProgressInterval={5}; AuthMode={6}; ForceAuthentication={7}; ParameterSet={8}; Input={9}" -f $permissionScope, [bool]$IncludeItemPermissions, [bool]$IncludeHiddenLists, [bool]$IncludeSystemLists, $PageSize, $ItemProgressInterval, $authMode, [bool]$ForceAuthentication, $PSCmdlet.ParameterSetName, $inputScope)
}
catch {
    Write-ConsoleWarning -Message ("Could not start transcript log '{0}': {1}" -f $LogPath, $_.Exception.Message)
}

try {
    $rootWebUrls = @(if ($PSCmdlet.ParameterSetName -eq 'Site') { $SiteUrl } else { Get-WebUrlsFromFile -Path $WebUrlsFile })
    Write-ConsoleMessage -Message ("Root web URLs loaded: {0}; descendant subsites are included." -f $rootWebUrls.Count)

    $expandedWebUrls = New-Object System.Collections.Generic.List[string]
    $seenExpandedWebUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($rootWebUrl in $rootWebUrls) {
        $rootExpandedWebUrls = @(Get-WebUrlsFromSite -Url $rootWebUrl)
        Write-ConsoleMessage -Message ("Web URLs expanded from root '{0}': {1}" -f $rootWebUrl, $rootExpandedWebUrls.Count)
        foreach ($expandedWebUrl in $rootExpandedWebUrls) {
            if ($seenExpandedWebUrls.Add($expandedWebUrl)) {
                $expandedWebUrls.Add($expandedWebUrl)
            }
        }
    }

    $webUrls = @($expandedWebUrls)
    Write-ConsoleMessage -Message ("Web URLs to inventory after expansion: {0}" -f $webUrls.Count)

    foreach ($webUrl in $webUrls) {
        try {
            Export-WebPermissionInventory -WebUrl $webUrl -CsvPath $TempOutputPath
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to inventory web permissions '{0}': {1}" -f $webUrl, $_.Exception.Message)
            Write-InventoryError -Scope 'Web' -Url $webUrl -Name $webUrl -Message $_.Exception.Message
        }
    }

    if (-not $script:CsvCreated) {
        Write-PermissionHeaderOnlyCsv -Path $TempOutputPath
        Write-ConsoleWarning -Message "No permission rows were exported. The final permission CSV contains headers only."
    }

    if ($script:ErrorCsvCreated) {
        throw ("Permission inventory errors were recorded. Final CSV was not published. Error details: {0}" -f $ErrorPath)
    }

    Move-Item -LiteralPath $TempOutputPath -Destination $OutputPath -Force
    $script:InventoryStopwatch.Stop()
    Write-ConsoleMessage -Message ("Permission inventory completed: {0}" -f $OutputPath)
    Write-ConsoleMessage -Message ("Scan duration: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
}
catch {
    if ($script:InventoryStopwatch -and $script:InventoryStopwatch.IsRunning) {
        $script:InventoryStopwatch.Stop()
    }
    Write-ConsoleMessage -Message ("ERROR: {0}" -f $_) -ForegroundColor Red
    if ($script:InventoryStopwatch) {
        Write-ConsoleWarning -Message ("Scan duration before failure: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
    }
    if (Test-Path -LiteralPath $TempOutputPath) {
        Write-ConsoleWarning -Message ("Permission inventory failed before final CSV publication. Partial temporary CSV kept: {0}" -f $TempOutputPath)
    }
    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
        Write-ConsoleWarning -Message ("Removed incomplete final permission CSV: {0}" -f $OutputPath)
    }
    $script:ConsoleLifecycleStatus = 'FAILED'
    exit 1
}
finally {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
}
$script:ConsoleLifecycleStatus = 'SUCCESS'
}
catch {
    $script:ConsoleLifecycleFailure = $_
    throw
}
finally {
    Complete-SmartM365MigrationConsoleLifecycle -Context $script:ConsoleLifecycleContext -Failure $script:ConsoleLifecycleFailure -Status $script:ConsoleLifecycleStatus
}


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDnSnH4NF1vhMVt
# HDb4zyzoCG7c6k3bYwyGNndrfnp0U6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAr0ltXNXDkPfKcXG6mygvS
# KspO7lRCnS2yOIC6AykdCzANBgkqhkiG9w0BAQEFAASCAYBvVGFyKFMTWxITjbtz
# zxPeGVLgYOStzcJWzNEOooxpHlU7SptrkQrUXYKb8B5HH4j723fY2hIRLqXNo3rP
# Ipoixfz6Fi/mUPnM1VgsE1XsWuAG26efMRVxo/pBEhFkhWx47Y55bU4um0s+vDHl
# k4sFS/cxrez32BCH2Ea5RdEvUhs8kyFpvAcqGPu5Nx11u0zKS6OaN/qwzuTHO382
# /bxR3cG8EQY3bJsgYHGdcG9a3TlwhagEtVpEt4I70cY1ju6UBaRRhtWRdg+rzRGM
# 8z5MKfP8GeHinB7xj0P/YQupC6hm8T4uls+LT1uNi7zZZWRNW5g4BunYudS7xTM4
# ISNJBXmNSebjce8Z1Yj35Euj8+obmDlF0EgBj4JCfTuvB7JaeYDvsNxL0VELMGfr
# SqzMqjFQsTf9aGz0g23kAdjIPNrbey2CbZEsxjYx8RJMCaqbNQcIXXpU3bHl0w6x
# EgMfjwOX91TC3rH5ceDTQbiGGroZWYipWTKsXjijd8Vt/Dg=
# SIG # End signature block
