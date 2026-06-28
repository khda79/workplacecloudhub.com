<#
.SYNOPSIS
    Inventories SharePoint Online permissions to a dedicated CSV file.

.DESCRIPTION
    Uses PnP.PowerShell with interactive, device login, or certificate authentication.
    The output columns intentionally match SmartM365-SharePointSource-PermissionInventory.ps1 as closely
    as possible so both inventories can be compared.

.VERSION
    1.0.0
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
    'ObjectScope',
    'ObjectUrl',
    'ObjectServerRelativeUrl',
    'ObjectTitle',
    'ObjectId',
    'ParentObjectUrl',
    'ListTitle',
    'ListUrl',
    'ItemId',
    'ItemFileSystemObjectType',
    'HasUniqueRoleAssignments',
    'InheritedFrom',
    'PrincipalType',
    'PrincipalName',
    'PrincipalLoginName',
    'PrincipalId',
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

    return ("{0:00}:{1:00}:{2:00}" -f [int]$Elapsed.TotalHours, $Elapsed.Minutes, $Elapsed.Seconds)
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
        [string]$Message
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
    Write-SPOConnectionIdentity -Connection $connection -Url $Url
    Write-PnPTokenSummary -Connection $connection
    return $connection
}

function Write-SPOConnectionIdentity {
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

    $rootFolder = Get-PnPProperty -ClientObject $List -Property RootFolder
    $systemUrls = @(
        '_catalogs/masterpage',
        '_catalogs/wp',
        '_catalogs/lt',
        'Style Library',
        'SiteAssets',
        'SitePages',
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

function Get-RoleAssignmentRows {
    param(
        [string]$SiteCollectionUrl,
        [string]$WebUrl,
        [string]$WebTitle,
        [string]$ObjectScope,
        [string]$ObjectUrl,
        [string]$ObjectServerRelativeUrl,
        [string]$ObjectTitle,
        [string]$ObjectId,
        [string]$ParentObjectUrl,
        [string]$ListTitle,
        [string]$ListUrl,
        [object]$ItemId,
        [string]$ItemFileSystemObjectType,
        [bool]$HasUniqueRoleAssignments,
        [string]$InheritedFrom,
        $RoleAssignments
    )

    foreach ($roleAssignment in $RoleAssignments) {
        try {
            $member = Get-PnPProperty -ClientObject $roleAssignment -Property Member
            $bindings = @(Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings)
            $permissionLevels = @($bindings | ForEach-Object { $_.Name })
            if ($permissionLevels.Count -eq 0) {
                continue
            }

            $principal = Get-PrincipalInfo -Member $member
            [pscustomobject]@{
                SiteCollectionUrl        = $SiteCollectionUrl
                WebUrl                   = $WebUrl
                WebTitle                 = $WebTitle
                ObjectScope              = $ObjectScope
                ObjectUrl                = $ObjectUrl
                ObjectServerRelativeUrl  = $ObjectServerRelativeUrl
                ObjectTitle              = $ObjectTitle
                ObjectId                 = $ObjectId
                ParentObjectUrl          = $ParentObjectUrl
                ListTitle                = $ListTitle
                ListUrl                  = $ListUrl
                ItemId                   = $ItemId
                ItemFileSystemObjectType = $ItemFileSystemObjectType
                HasUniqueRoleAssignments = $HasUniqueRoleAssignments
                InheritedFrom            = $InheritedFrom
                PrincipalType            = $principal.PrincipalType
                PrincipalName            = $principal.PrincipalName
                PrincipalLoginName       = $principal.PrincipalLoginName
                PrincipalId              = $principal.PrincipalId
                PermissionLevels         = ($permissionLevels -join '|')
                IsLimitedAccessOnly      = ($permissionLevels.Count -eq 1 -and $permissionLevels[0] -eq 'Limited Access')
            }
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to read role assignment on '{0}': {1}" -f $ObjectUrl, $_.Exception.Message)
            Write-InventoryError -Scope "$ObjectScope RoleAssignment" -Url $ObjectUrl -Name $ObjectTitle -Message $_.Exception.Message
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

    $rootFolder = Get-PnPProperty -ClientObject $List -Property RootFolder
    $listUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl $rootFolder.ServerRelativeUrl
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        $script:SPOPermissionListPageItemsRead = 0
        $script:SPOPermissionListPageLastHeartbeat = 0
        $script:SPOPermissionListPageWebUrl = $Web.Url
        $script:SPOPermissionListPageTitle = $List.Title
        $script:SPOPermissionListPageStopwatch = $stopwatch
        $script:SPOPermissionListPageInterval = $ProgressInterval

        $items = Get-PnPListItem `
            -List $List `
            -PageSize $PageSize `
            -Fields 'FileRef','FileLeafRef','FSObjType','UniqueId','ID' `
            -ScriptBlock {
                param($PageItems)

                $script:SPOPermissionListPageItemsRead += $PageItems.Count
                if (($script:SPOPermissionListPageItemsRead - $script:SPOPermissionListPageLastHeartbeat) -ge $script:SPOPermissionListPageInterval) {
                    $script:SPOPermissionListPageLastHeartbeat = $script:SPOPermissionListPageItemsRead
                    Write-ItemPermissionHeartbeat `
                        -WebUrl $script:SPOPermissionListPageWebUrl `
                        -ListTitle $script:SPOPermissionListPageTitle `
                        -ProcessedItems $script:SPOPermissionListPageItemsRead `
                        -UniquePermissionItems 0 `
                        -ExportedRows 0 `
                        -Elapsed $script:SPOPermissionListPageStopwatch.Elapsed `
                        -LastItem 'retrieving list items'
                }
            } `
            -ErrorAction Stop
    }
    catch {
        Write-ConsoleWarning -Message ("Failed to enumerate items for list '{0}' in web '{1}': {2}" -f $List.Title, $Web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'ListItems' -Url $listUrl -Name $List.Title -Message $_.Exception.Message
        return
    }

    $processedItems = 0
    $uniquePermissionItems = 0
    $exportedRows = 0
    $lastItem = '<none>'

    foreach ($item in $items) {
        $processedItems++
        try {
            $lastItem = if ($item.Id) { "ID $($item.Id)" } else { '<unknown>' }
            $hasUniqueRoleAssignments = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments
            if ($processedItems % $ProgressInterval -eq 0) {
                Write-ItemPermissionHeartbeat `
                    -WebUrl $Web.Url `
                    -ListTitle $List.Title `
                    -ProcessedItems $processedItems `
                    -UniquePermissionItems $uniquePermissionItems `
                    -ExportedRows $exportedRows `
                    -Elapsed $stopwatch.Elapsed `
                    -LastItem $lastItem
            }

            if (-not $hasUniqueRoleAssignments) {
                continue
            }

            $uniquePermissionItems++
            $roleAssignments = Get-PnPProperty -ClientObject $item -Property RoleAssignments
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
                    -ObjectScope 'Item' `
                    -ObjectUrl $absoluteUrl `
                    -ObjectServerRelativeUrl $serverRelativeUrl `
                    -ObjectTitle $itemName `
                    -ObjectId ([string]$item.FieldValues.UniqueId) `
                    -ParentObjectUrl $listUrl `
                    -ListTitle $List.Title `
                    -ListUrl $listUrl `
                    -ItemId $item.Id `
                    -ItemFileSystemObjectType $fsObjType `
                    -HasUniqueRoleAssignments $true `
                    -InheritedFrom '' `
                    -RoleAssignments $roleAssignments)

            Export-PermissionRows -Rows $rows -CsvPath $CsvPath
            $exportedRows += $rows.Count
        }
        catch {
            Write-ConsoleWarning -Message ("Failed to inventory item permissions in list '{0}' in web '{1}': {2}" -f $List.Title, $Web.Url, $_.Exception.Message)
            Write-InventoryError -Scope 'Item' -Url $listUrl -Name $List.Title -Message $_.Exception.Message
        }
    }

    $stopwatch.Stop()
    Write-ItemPermissionHeartbeat `
        -WebUrl $Web.Url `
        -ListTitle $List.Title `
        -ProcessedItems $processedItems `
        -UniquePermissionItems $uniquePermissionItems `
        -ExportedRows $exportedRows `
        -Elapsed $stopwatch.Elapsed `
        -LastItem $lastItem
}

function Export-WebPermissionInventory {
    param(
        [string]$WebUrl,
        [string]$CsvPath
    )

    Connect-ToSPOWeb -Url $WebUrl
    $web = Get-PnPWeb -Includes Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments -ErrorAction Stop
    $siteCollectionUrl = Get-SiteCollectionUrlFromWebUrl -WebUrl $web.Url

    Write-ConsoleMessage -Message ("Web: {0}" -f $web.Url)

    try {
        $webRoleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments
        $rows = @(Get-RoleAssignmentRows `
                -SiteCollectionUrl $siteCollectionUrl `
                -WebUrl $web.Url `
                -WebTitle $web.Title `
                -ObjectScope 'Web' `
                -ObjectUrl $web.Url `
                -ObjectServerRelativeUrl $web.ServerRelativeUrl `
                -ObjectTitle $web.Title `
                -ObjectId ([string]$web.Id) `
                -ParentObjectUrl '' `
                -ListTitle '' `
                -ListUrl '' `
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
        $lists = Get-PnPList -Includes Title,Hidden,BaseTemplate,Id,RootFolder,HasUniqueRoleAssignments -ErrorAction Stop
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
            $rootFolder = Get-PnPProperty -ClientObject $list -Property RootFolder
            $listTitle = $list.Title
            $listUrl = ConvertTo-AbsoluteSharePointUrl -WebUrl $web.Url -ServerRelativeUrl $rootFolder.ServerRelativeUrl

            if (-not $IncludeHiddenLists -and $list.Hidden) {
                continue
            }

            if (-not $IncludeSystemLists -and (Test-SystemList -List $list)) {
                continue
            }

            if ($DocumentLibrariesOnly -and $list.BaseTemplate -ne 101) {
                continue
            }

            $roleAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments
            $rows = @(Get-RoleAssignmentRows `
                    -SiteCollectionUrl $siteCollectionUrl `
                    -WebUrl $web.Url `
                    -WebTitle $web.Title `
                    -ObjectScope 'List' `
                    -ObjectUrl $listUrl `
                    -ObjectServerRelativeUrl $rootFolder.ServerRelativeUrl `
                    -ObjectTitle $list.Title `
                    -ObjectId ([string]$list.Id) `
                    -ParentObjectUrl $web.Url `
                    -ListTitle $list.Title `
                    -ListUrl $listUrl `
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
}
catch {
    Write-ConsoleWarning -Message ("Could not start transcript log '{0}': {1}" -f $LogPath, $_.Exception.Message)
}

try {
    $webUrls = if ($PSCmdlet.ParameterSetName -eq 'Site') { @($SiteUrl) } else { @(Get-WebUrlsFromFile -Path $WebUrlsFile) }
    Write-ConsoleMessage -Message ("Web URLs loaded: {0}" -f $webUrls.Count)

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

    Move-Item -LiteralPath $TempOutputPath -Destination $OutputPath -Force
    $script:InventoryStopwatch.Stop()
    Write-ConsoleMessage -Message ("Permission inventory completed: {0}" -f $OutputPath)
    Write-ConsoleMessage -Message ("Scan duration: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
    if ($script:ErrorCsvCreated) {
        Write-ConsoleWarning -Message ("Some permissions could not be inventoried. Error details: {0}" -f $ErrorPath)
    }
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
    exit 1
}
finally {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
}
