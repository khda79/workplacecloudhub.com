<#
.SYNOPSIS
    Inventories SharePoint 2019 permissions to a dedicated CSV file.

.DESCRIPTION
    Run this script on a SharePoint 2019 server from the SharePoint Management
    Shell, or from a PowerShell console started with farm administrator rights.

    The script exports role assignments for webs and lists/libraries. Item and
    folder permissions are optional because they can create very large outputs.
    When -IncludeItemPermissions is used, only items and folders with unique
    permissions are exported.

.EXAMPLE
    .\SmartM365-SharePointSource-PermissionInventory.ps1 -WebApplicationUrl "https://intranet"

.EXAMPLE
    .\SmartM365-SharePointSource-PermissionInventory.ps1 -WebApplicationUrl "https://intranet" -OutputPath "C:\Temp\SP2019PermissionScans"

.EXAMPLE
    .\SmartM365-SharePointSource-PermissionInventory.ps1 -WebApplicationUrl "https://intranet" -UseSiteUrlFilter -SiteUrlsFile "C:\Temp\sp2019-site-urls.txt"

.EXAMPLE
    .\SmartM365-SharePointSource-PermissionInventory.ps1 -WebUrl "https://intranet/sites/finance" -DocumentLibrariesOnly -IncludeItemPermissions
#>

[CmdletBinding(DefaultParameterSetName = 'WebApplication')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'WebApplication')]
    [ValidateNotNullOrEmpty()]
    [string]$WebApplicationUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Site')]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Web')]
    [ValidateNotNullOrEmpty()]
    [string]$WebUrl,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [ValidateNotNullOrEmpty()]
    [string]$ErrorPath,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [ValidateNotNullOrEmpty()]
    [string]$SiteUrlsFile,

    [switch]$UseSiteUrlFilter,

    [switch]$IncludeHiddenLists,

    [switch]$IncludeSystemLists,

    [switch]$DocumentLibrariesOnly,

    [switch]$IncludeItemPermissions = $true,

    [int]$RowLimit = 2000,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$ItemProgressInterval = 500
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

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

function Add-SharePointSnapIn {
    if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }
}

function ConvertTo-SafeFileName {
    param(
        [string]$Name,
        [string]$Fallback = 'SharePoint2019'
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

    Write-Host ("{0}   Item heartbeat: web='{1}' library='{2}' processed={3}; unique={4}; permission rows exported={5}; elapsed={6}; last='{7}'" -f `
            (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), `
            $WebUrl, `
            $ListTitle, `
            $ProcessedItems, `
            $UniquePermissionItems, `
            $ExportedRows, `
            (Format-InventoryDuration -Elapsed $Elapsed), `
            $LastItem) -ForegroundColor DarkCyan
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

function Write-PermissionHeaderOnlyCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $header = ($PermissionColumns | ForEach-Object { '"{0}"' -f ($_ -replace '"', '""') }) -join ';'
    Set-Content -LiteralPath $Path -Value $header -Encoding UTF8
}

function Get-NameFromUrl {
    param(
        [string]$Url
    )

    try {
        $uri = New-Object System.Uri($Url)
        $path = $uri.AbsolutePath.Trim('/')
        if ($path) {
            $segments = @($path.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($segments.Count -gt 0) {
                return $segments[$segments.Count - 1]
            }
        }

        return $uri.Host
    }
    catch {
        return $Url
    }
}

function Get-NormalizedUrlPath {
    param(
        [string]$Url
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $null
    }

    $cleanUrl = $Url.Trim()
    try {
        $path = (New-Object System.Uri($cleanUrl)).AbsolutePath
    }
    catch {
        $path = $cleanUrl
    }

    $path = ($path -replace '[?#].*$', '').Trim()
    if (-not $path.StartsWith('/')) {
        $path = "/$path"
    }

    $path = $path.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($path)) {
        $path = '/'
    }

    return $path.ToLowerInvariant()
}

function Import-SiteUrlFilter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Site URL filter file not found: $Path"
    }

    $filterPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Get-Content -LiteralPath $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        ForEach-Object {
            $normalizedPath = Get-NormalizedUrlPath -Url $_
            if ($normalizedPath) {
                [void]$filterPaths.Add($normalizedPath)
            }
        }

    return $filterPaths
}

function Test-WebMatchesSiteUrlFilter {
    param(
        [Microsoft.SharePoint.SPWeb]$Web
    )

    if (-not $script:UseSiteUrlFilter) {
        return $true
    }

    $webPath = Get-NormalizedUrlPath -Url $Web.Url
    return $script:SiteUrlFilterPaths.Contains($webPath)
}

function Test-SystemList {
    param(
        [Microsoft.SharePoint.SPList]$List
    )

    $systemUrls = @(
        '_catalogs/masterpage',
        '_catalogs/wp',
        '_catalogs/lt',
        'Style Library',
        'SiteAssets',
        'SitePages',
        'FormServerTemplates'
    )

    foreach ($url in $systemUrls) {
        if ($List.RootFolder.Url -like "$url*" -or $List.RootFolder.ServerRelativeUrl -like "*/$url*") {
            return $true
        }
    }

    return $false
}

function ConvertTo-AbsoluteSharePointUrl {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
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

    $webUri = New-Object System.Uri($Web.Url)
    return ("{0}://{1}{2}" -f $webUri.Scheme, $webUri.Authority, $ServerRelativeUrl)
}

function Get-PrincipalInfo {
    param(
        $Member
    )

    $principalType = $Member.GetType().Name
    $name = $Member.Name
    $loginName = $null
    $id = $null

    if ($Member -is [Microsoft.SharePoint.SPGroup]) {
        $principalType = 'SharePointGroup'
        $loginName = $Member.LoginName
        $id = $Member.ID
    }
    elseif ($Member -is [Microsoft.SharePoint.SPUser]) {
        if ($Member.IsDomainGroup) {
            $principalType = 'DomainGroup'
        }
        else {
            $principalType = 'User'
        }

        $loginName = $Member.LoginName
        $id = $Member.ID
    }

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
        $permissionLevels = @($roleAssignment.RoleDefinitionBindings | ForEach-Object { $_.Name })
        if ($permissionLevels.Count -eq 0) {
            continue
        }

        $principal = Get-PrincipalInfo -Member $roleAssignment.Member
        [pscustomobject]@{
            SiteCollectionUrl       = $SiteCollectionUrl
            WebUrl                  = $WebUrl
            WebTitle                = $WebTitle
            ObjectScope             = $ObjectScope
            ObjectUrl               = $ObjectUrl
            ObjectServerRelativeUrl = $ObjectServerRelativeUrl
            ObjectTitle             = $ObjectTitle
            ObjectId                = $ObjectId
            ParentObjectUrl         = $ParentObjectUrl
            ListTitle               = $ListTitle
            ListUrl                 = $ListUrl
            ItemId                  = $ItemId
            ItemFileSystemObjectType = $ItemFileSystemObjectType
            HasUniqueRoleAssignments = $HasUniqueRoleAssignments
            InheritedFrom           = $InheritedFrom
            PrincipalType           = $principal.PrincipalType
            PrincipalName           = $principal.PrincipalName
            PrincipalLoginName      = $principal.PrincipalLoginName
            PrincipalId             = $principal.PrincipalId
            PermissionLevels        = ($permissionLevels -join '|')
            IsLimitedAccessOnly     = ($permissionLevels.Count -eq 1 -and $permissionLevels[0] -eq 'Limited Access')
        }
    }
}

function Export-PermissionRows {
    param(
        [object[]]$Rows,
        [string]$CsvPath,
        [switch]$Append
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        return
    }

    if ($Append -or $script:CsvCreated) {
        $Rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $Rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8
        $script:CsvCreated = $true
    }
}

function Get-ItemServerRelativeUrl {
    param(
        [Microsoft.SharePoint.SPListItem]$Item
    )

    if ($Item.File -and $Item.File.ServerRelativeUrl) {
        return $Item.File.ServerRelativeUrl
    }

    if ($Item.Folder -and $Item.Folder.ServerRelativeUrl) {
        return $Item.Folder.ServerRelativeUrl
    }

    if ($Item.Url) {
        return $Item.Url
    }

    return $null
}

function Export-ItemPermissionInventory {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
        [Microsoft.SharePoint.SPList]$List,
        [string]$CsvPath,
        [int]$BatchSize,
        [int]$ProgressInterval
    )

    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.ViewAttributes = "Scope='RecursiveAll'"
    $query.RowLimit = [uint32]$BatchSize

    $processedItems = 0
    $uniquePermissionItems = 0
    $exportedRows = 0
    $lastItem = '<none>'
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    do {
        try {
            $items = $List.GetItems($query)
            $query.ListItemCollectionPosition = $items.ListItemCollectionPosition
        }
        catch {
            $listUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl $List.RootFolder.ServerRelativeUrl
            Write-Warning ("Failed to enumerate items for list '{0}' in web '{1}': {2}" -f $List.Title, $Web.Url, $_.Exception.Message)
            Write-InventoryError -Scope 'ListItems' -Url $listUrl -Name $List.Title -Message $_.Exception.Message
            break
        }

        foreach ($item in $items) {
            $processedItems++
            try {
                $lastItem = if ($item.ID) { "ID $($item.ID)" } else { '<unknown>' }
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

                if (-not $item.HasUniqueRoleAssignments) {
                    continue
                }

                $uniquePermissionItems++
                $serverRelativeUrl = Get-ItemServerRelativeUrl -Item $item
                $absoluteUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl $serverRelativeUrl
                $fsObjType = if ([string]$item['FSObjType'] -eq '1') { 'Folder' } else { 'FileOrItem' }
                $listUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl $List.RootFolder.ServerRelativeUrl
                $rows = @(Get-RoleAssignmentRows `
                        -SiteCollectionUrl $Web.Site.Url `
                        -WebUrl $Web.Url `
                        -WebTitle $Web.Title `
                        -ObjectScope 'Item' `
                        -ObjectUrl $absoluteUrl `
                        -ObjectServerRelativeUrl $serverRelativeUrl `
                        -ObjectTitle ([string]$item.Name) `
                        -ObjectId ([string]$item.UniqueId) `
                        -ParentObjectUrl $listUrl `
                        -ListTitle $List.Title `
                        -ListUrl $listUrl `
                        -ItemId $item.ID `
                        -ItemFileSystemObjectType $fsObjType `
                        -HasUniqueRoleAssignments $true `
                        -InheritedFrom '' `
                        -RoleAssignments $item.RoleAssignments)

                Export-PermissionRows -Rows $rows -CsvPath $CsvPath
                $exportedRows += $rows.Count
            }
            catch {
                $itemUrl = $null
                $itemName = $null
                try {
                    $itemUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl (Get-ItemServerRelativeUrl -Item $item)
                    $itemName = [string]$item.Name
                }
                catch {
                    $itemUrl = $Web.Url
                    $itemName = $List.Title
                }

                Write-Warning ("Failed to inventory item permissions in list '{0}' in web '{1}': {2}" -f $List.Title, $Web.Url, $_.Exception.Message)
                Write-InventoryError -Scope 'Item' -Url $itemUrl -Name $itemName -Message $_.Exception.Message
            }
        }
    }
    while ($null -ne $query.ListItemCollectionPosition)

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
        [Microsoft.SharePoint.SPWeb]$Web,
        [string]$CsvPath,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [switch]$LibrariesOnly,
        [switch]$IncludeItems,
        [int]$BatchSize
    )

    Write-Host ("Web: {0}" -f $Web.Url)

    try {
        $webHasUniqueRoleAssignments = $Web.HasUniqueRoleAssignments
        $webRows = @(Get-RoleAssignmentRows `
                -SiteCollectionUrl $Web.Site.Url `
                -WebUrl $Web.Url `
                -WebTitle $Web.Title `
                -ObjectScope 'Web' `
                -ObjectUrl $Web.Url `
                -ObjectServerRelativeUrl $Web.ServerRelativeUrl `
                -ObjectTitle $Web.Title `
                -ObjectId ([string]$Web.ID) `
                -ParentObjectUrl '' `
                -ListTitle '' `
                -ListUrl '' `
                -ItemId $null `
                -ItemFileSystemObjectType '' `
                -HasUniqueRoleAssignments $webHasUniqueRoleAssignments `
                -InheritedFrom $(if ($webHasUniqueRoleAssignments) { '' } else { 'ParentWeb' }) `
                -RoleAssignments $Web.RoleAssignments)
        Export-PermissionRows -Rows $webRows -CsvPath $CsvPath
    }
    catch {
        Write-Warning ("Failed to inventory web permissions '{0}': {1}" -f $Web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'Web' -Url $Web.Url -Name $Web.Title -Message $_.Exception.Message
    }

    try {
        $lists = @($Web.Lists)
    }
    catch {
        Write-Warning ("Failed to enumerate lists for web '{0}': {1}" -f $Web.Url, $_.Exception.Message)
        Write-InventoryError -Scope 'WebLists' -Url $Web.Url -Name $Web.Title -Message $_.Exception.Message
        return
    }

    foreach ($list in $lists) {
        $listTitle = '<unknown>'
        $listUrl = $Web.Url

        try {
            $listTitle = $list.Title
            $listUrl = ConvertTo-AbsoluteSharePointUrl -Web $Web -ServerRelativeUrl $list.RootFolder.ServerRelativeUrl

            if (-not $IncludeHidden -and $list.Hidden) {
                continue
            }

            if (-not $IncludeSystem -and (Test-SystemList -List $list)) {
                continue
            }

            if ($LibrariesOnly -and $list.BaseTemplate -ne [Microsoft.SharePoint.SPListTemplateType]::DocumentLibrary) {
                continue
            }

            $listHasUniqueRoleAssignments = $list.HasUniqueRoleAssignments
            $inheritedFrom = if ($listHasUniqueRoleAssignments) { '' } else { 'Web' }
            $listRows = @(Get-RoleAssignmentRows `
                    -SiteCollectionUrl $Web.Site.Url `
                    -WebUrl $Web.Url `
                    -WebTitle $Web.Title `
                    -ObjectScope 'List' `
                    -ObjectUrl $listUrl `
                    -ObjectServerRelativeUrl $list.RootFolder.ServerRelativeUrl `
                    -ObjectTitle $list.Title `
                    -ObjectId ([string]$list.ID) `
                    -ParentObjectUrl $Web.Url `
                    -ListTitle $list.Title `
                    -ListUrl $listUrl `
                    -ItemId $null `
                    -ItemFileSystemObjectType '' `
                    -HasUniqueRoleAssignments $listHasUniqueRoleAssignments `
                    -InheritedFrom $inheritedFrom `
                    -RoleAssignments $list.RoleAssignments)
            Export-PermissionRows -Rows $listRows -CsvPath $CsvPath

            if ($IncludeItems) {
                Write-Host ("  Item permissions: {0}" -f $list.Title)
                Export-ItemPermissionInventory -Web $Web -List $list -CsvPath $CsvPath -BatchSize $BatchSize -ProgressInterval $ItemProgressInterval
            }
        }
        catch {
            Write-Warning ("Failed to inventory list permissions '{0}' in web '{1}': {2}" -f $listTitle, $Web.Url, $_.Exception.Message)
            Write-InventoryError -Scope 'List' -Url $listUrl -Name $listTitle -Message $_.Exception.Message
        }
    }
}

function Export-SitePermissionInventory {
    param(
        [Microsoft.SharePoint.SPSite]$Site,
        [string]$CsvPath,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [switch]$LibrariesOnly,
        [switch]$IncludeItems,
        [int]$BatchSize
    )

    Write-Host ("Site collection: {0}" -f $Site.Url)
    foreach ($web in $Site.AllWebs) {
        try {
            if (Test-WebMatchesSiteUrlFilter -Web $web) {
                Export-WebPermissionInventory `
                    -Web $web `
                    -CsvPath $CsvPath `
                    -IncludeHidden:$IncludeHidden `
                    -IncludeSystem:$IncludeSystem `
                    -LibrariesOnly:$LibrariesOnly `
                    -IncludeItems:$IncludeItems `
                    -BatchSize $BatchSize
            }
            else {
                Write-Host ("Skipping web not in filter: {0}" -f $web.Url)
            }
        }
        finally {
            $web.Dispose()
        }
    }
}

function Get-TargetSafeName {
    param(
        [string]$ParameterSetName,
        [string]$WebApplicationUrl,
        [string]$SiteUrl,
        [string]$WebUrl
    )

    $targetName = $null

    switch ($ParameterSetName) {
        'WebApplication' {
            try {
                $targetName = (Get-SPWebApplication -Identity $WebApplicationUrl).Name
            }
            catch {
                $targetName = Get-NameFromUrl -Url $WebApplicationUrl
            }
        }

        'Site' {
            $targetName = Get-NameFromUrl -Url $SiteUrl
        }

        'Web' {
            $targetName = Get-NameFromUrl -Url $WebUrl
        }
    }

    return (ConvertTo-SafeFileName -Name $targetName)
}

function Get-DefaultOutputPath {
    param(
        [string]$ParameterSetName,
        [string]$WebApplicationUrl,
        [string]$SiteUrl,
        [string]$WebUrl,
        [string]$BaseDirectory
    )

    if ([string]::IsNullOrWhiteSpace($BaseDirectory)) {
        $BaseDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    }

    $safeTargetName = Get-TargetSafeName `
        -ParameterSetName $ParameterSetName `
        -WebApplicationUrl $WebApplicationUrl `
        -SiteUrl $SiteUrl `
        -WebUrl $WebUrl

    if ([string]::IsNullOrWhiteSpace($safeTargetName)) {
        $safeTargetName = 'SharePoint'
    }

    $safeTargetName = ConvertTo-SafeFileName -Name $safeTargetName
    $targetDirectory = Join-Path -Path $BaseDirectory -ChildPath $safeTargetName
    return Join-Path -Path $targetDirectory -ChildPath ("SP2019-PermissionInventory-{0}-{1:yyyyMMdd-HHmmss}.csv" -f $safeTargetName, (Get-Date))
}

function Test-OutputPathIsDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path -PathType Container) {
        return $true
    }

    if ($Path.EndsWith('\') -or $Path.EndsWith('/')) {
        return $true
    }

    $extension = [System.IO.Path]::GetExtension($Path)
    return [string]::IsNullOrWhiteSpace($extension)
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

Add-SharePointSnapIn

$script:UseSiteUrlFilter = $UseSiteUrlFilter -or (-not [string]::IsNullOrWhiteSpace($SiteUrlsFile))
$script:SiteUrlFilterPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

if ($script:UseSiteUrlFilter) {
    if ([string]::IsNullOrWhiteSpace($SiteUrlsFile)) {
        throw "UseSiteUrlFilter requires -SiteUrlsFile."
    }

    $script:SiteUrlFilterPaths = Import-SiteUrlFilter -Path $SiteUrlsFile
    Write-Host ("Site URL filter enabled: {0} paths loaded from {1}" -f $script:SiteUrlFilterPaths.Count, $SiteUrlsFile)
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Get-DefaultOutputPath `
        -ParameterSetName $PSCmdlet.ParameterSetName `
        -WebApplicationUrl $WebApplicationUrl `
        -SiteUrl $SiteUrl `
        -WebUrl $WebUrl
}
elseif (Test-OutputPathIsDirectory -Path $OutputPath) {
    $OutputPath = Get-DefaultOutputPath `
        -ParameterSetName $PSCmdlet.ParameterSetName `
        -WebApplicationUrl $WebApplicationUrl `
        -SiteUrl $SiteUrl `
        -WebUrl $WebUrl `
        -BaseDirectory $OutputPath
}

if ([string]::IsNullOrWhiteSpace($ErrorPath)) {
    $errorBaseDirectory = Split-Path -Path $OutputPath -Parent
    if ([string]::IsNullOrWhiteSpace($errorBaseDirectory)) {
        $errorBaseDirectory = (Get-Location).Path
    }

    $ErrorPath = Join-Path -Path $errorBaseDirectory -ChildPath ("{0}-Errors.csv" -f [System.IO.Path]::GetFileNameWithoutExtension($OutputPath))
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $logBaseDirectory = Split-Path -Path $OutputPath -Parent
    if ([string]::IsNullOrWhiteSpace($logBaseDirectory)) {
        $logBaseDirectory = (Get-Location).Path
    }

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

try {
    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
    Write-Host ("Permission inventory output: {0}" -f $OutputPath)
    Write-Host ("Temporary permission inventory output: {0}" -f $TempOutputPath)
    Write-Host ("Error output: {0}" -f $ErrorPath)
    Write-Host ("Run log: {0}" -f $LogPath)
}
catch {
    Write-Warning ("Could not start transcript log '{0}': {1}" -f $LogPath, $_.Exception.Message)
}

try {
    switch ($PSCmdlet.ParameterSetName) {
        'WebApplication' {
            $sites = Get-SPSite -WebApplication $WebApplicationUrl -Limit All
            foreach ($site in $sites) {
                try {
                    Export-SitePermissionInventory `
                        -Site $site `
                        -CsvPath $TempOutputPath `
                        -IncludeHidden:$IncludeHiddenLists `
                        -IncludeSystem:$IncludeSystemLists `
                        -LibrariesOnly:$DocumentLibrariesOnly `
                        -IncludeItems:$IncludeItemPermissions `
                        -BatchSize $RowLimit
                }
                catch {
                    Write-Warning ("Failed to inventory site collection permissions '{0}': {1}" -f $site.Url, $_.Exception.Message)
                    Write-InventoryError -Scope 'SiteCollection' -Url $site.Url -Name $site.Url -Message $_.Exception.Message
                }
                finally {
                    $site.Dispose()
                }
            }
        }

        'Site' {
            $site = Get-SPSite -Identity $SiteUrl
            try {
                Export-SitePermissionInventory `
                    -Site $site `
                    -CsvPath $TempOutputPath `
                    -IncludeHidden:$IncludeHiddenLists `
                    -IncludeSystem:$IncludeSystemLists `
                    -LibrariesOnly:$DocumentLibrariesOnly `
                    -IncludeItems:$IncludeItemPermissions `
                    -BatchSize $RowLimit
            }
            finally {
                $site.Dispose()
            }
        }

        'Web' {
            $site = New-Object Microsoft.SharePoint.SPSite($WebUrl)
            try {
                $web = $site.OpenWeb()
                try {
                    if (Test-WebMatchesSiteUrlFilter -Web $web) {
                        Export-WebPermissionInventory `
                            -Web $web `
                            -CsvPath $TempOutputPath `
                            -IncludeHidden:$IncludeHiddenLists `
                            -IncludeSystem:$IncludeSystemLists `
                            -LibrariesOnly:$DocumentLibrariesOnly `
                            -IncludeItems:$IncludeItemPermissions `
                            -BatchSize $RowLimit
                    }
                    else {
                        Write-Host ("Skipping web not in filter: {0}" -f $web.Url)
                    }
                }
                finally {
                    $web.Dispose()
                }
            }
            finally {
                $site.Dispose()
            }
        }
    }

    if (-not $script:CsvCreated) {
        Write-PermissionHeaderOnlyCsv -Path $TempOutputPath
        Write-Warning "No permission rows were exported. The final permission CSV contains headers only."
    }

    Move-Item -LiteralPath $TempOutputPath -Destination $OutputPath -Force
    $script:InventoryStopwatch.Stop()
    Write-Host ("Permission inventory completed: {0}" -f $OutputPath)
    Write-Host ("Scan duration: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
    if ($script:ErrorCsvCreated) {
        Write-Warning ("Some permissions could not be inventoried. Error details: {0}" -f $ErrorPath)
    }
}
catch {
    if ($script:InventoryStopwatch -and $script:InventoryStopwatch.IsRunning) {
        $script:InventoryStopwatch.Stop()
    }
    Write-Error $_
    if ($script:InventoryStopwatch) {
        Write-Warning ("Scan duration before failure: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
    }
    if (Test-Path -LiteralPath $TempOutputPath) {
        Write-Warning ("Permission inventory failed before final CSV publication. Partial temporary CSV kept: {0}" -f $TempOutputPath)
    }
    exit 1
}
finally {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
}
