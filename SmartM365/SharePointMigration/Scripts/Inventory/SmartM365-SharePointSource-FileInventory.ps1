<#
.SYNOPSIS
    Inventories files from SharePoint 2019 document libraries.

.DESCRIPTION
    Run this script on a SharePoint 2019 server from the SharePoint Management
    Shell, or from a PowerShell console started with farm administrator rights.

    The script scans document libraries and exports file metadata to a CSV file.
    It can target a web application, a site collection, or a specific web.
    If OutputPath is not specified, the script creates a folder in the script
    directory using the target site name, then writes the CSV, run log, and
    error CSV in that folder.
    Permission inventory is not included by default. Use
    -IncludePermissionInventory to start SmartM365-SharePointSource-PermissionInventory.ps1
    after a successful file inventory and write a dedicated permission CSV next
    to the file inventory CSV.

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventory.ps1 -WebApplicationUrl "https://intranet"

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventory.ps1 -SiteUrl "https://intranet/sites/finance" -OutputPath "C:\Temp\Finance-Files.csv"

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventory.ps1 -WebApplicationUrl "https://intranet" -OutputPath "C:\Temp\SP2019Scans"

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventory.ps1 -WebUrl "https://intranet/sites/finance/compta" -OutputPath "C:\Temp\Compta-Files.csv" -IncludeHiddenLibraries

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventory.ps1 -WebApplicationUrl "https://intranet" -UseSiteUrlFilter -SiteUrlsFile "C:\Temp\sp2019-site-urls.txt"

.EXAMPLE
    .\SmartM365-SharePointSource-FileInventory.ps1 -WebApplicationUrl "https://intranet" -IncludePermissionInventory

.VERSION
    1.0.0
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

    [switch]$IncludeHiddenLibraries,

    [switch]$IncludeSystemLibraries,

    [switch]$IncludePermissionInventory,

    [switch]$IncludePermissionItemPermissions,

    [int]$RowLimit = 2000
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

function Get-ObjectCount {
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return 0
    }

    $countProperty = $Value.GetType().GetProperty('Count')
    if ($countProperty) {
        return [int]$countProperty.GetValue($Value, $null)
    }

    return @($Value).Count
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

$InventoryColumns = @(
    'SiteCollectionUrl',
    'WebUrl',
    'WebTitle',
    'LibraryTitle',
    'LibraryUrl',
    'ItemId',
    'UniqueId',
    'FileName',
    'FileUrl',
    'ServerRelativeUrl',
    'Extension',
    'SizeBytes',
    'SizeMB',
    'Created',
    'CreatedBy',
    'Modified',
    'ModifiedBy',
    'ContentType',
    'Version',
    'VersionsCount',
    'CheckedOutBy'
)

function Add-SharePointSnapIn {
    if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }
}

function Convert-SPFieldUserValueToString {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $Value,

        [Microsoft.SharePoint.SPWeb]$Web
    )

    if ($null -eq $Value) {
        return $null
    }

    try {
        $userValue = New-Object Microsoft.SharePoint.SPFieldUserValue($Web, [string]$Value)
        if ($userValue.User -and $userValue.User.Email) {
            return ("{0} <{1}>" -f $userValue.User.Name, $userValue.User.Email)
        }

        if ($userValue.LookupValue) {
            return $userValue.LookupValue
        }
    }
    catch {
        return [string]$Value
    }

    return [string]$Value
}

function Test-SystemLibrary {
    param(
        [Microsoft.SharePoint.SPList]$List
    )

    $systemLibraryUrls = @(
        '_catalogs/masterpage',
        '_catalogs/wp',
        '_catalogs/lt',
        'Style Library',
        'SiteAssets',
        'SitePages',
        'FormServerTemplates'
    )

    foreach ($url in $systemLibraryUrls) {
        if ($List.RootFolder.ServerRelativeUrl -like "*/$url" -or $List.RootFolder.ServerRelativeUrl -like "*/$url/*") {
            return $true
        }
    }

    return $false
}

function Get-DocumentLibraries {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem
    )

    foreach ($list in $Web.Lists) {
        if ($list.BaseType -ne [Microsoft.SharePoint.SPBaseType]::DocumentLibrary) {
            continue
        }

        if (-not $IncludeHidden -and $list.Hidden) {
            continue
        }

        if (-not $IncludeSystem -and (Test-SystemLibrary -List $list)) {
            continue
        }

        if ($list.ItemCount -eq 0) {
            continue
        }

        $list
    }
}

function Get-FileInventoryFromLibrary {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
        [Microsoft.SharePoint.SPList]$Library,
        [int]$BatchSize
    )

    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.ViewAttributes = "Scope='RecursiveAll'"
    $query.RowLimit = [uint32]$BatchSize
    $query.Query = @"
<Where>
  <Eq>
    <FieldRef Name='FSObjType' />
    <Value Type='Integer'>0</Value>
  </Eq>
</Where>
"@

    try {
        $query.QueryThrottleMode = [Microsoft.SharePoint.SPQueryThrottleOption]::Override
    }
    catch {
        # This property depends on the installed SharePoint/.NET patch level.
    }

    do {
        $items = $Library.GetItems($query)
        $query.ListItemCollectionPosition = $items.ListItemCollectionPosition

        foreach ($item in $items) {
            $file = $item.File
            if ($null -eq $file) {
                continue
            }

            $fileSizeBytes = $null
            if ($item.Fields.ContainsField('File Size') -and $null -ne $item['File_x0020_Size']) {
                [long]$fileSizeBytes = $item['File_x0020_Size']
            }
            elseif ($file.Length -ge 0) {
                [long]$fileSizeBytes = $file.Length
            }

            $checkedOutBy = $null
            if ($file.CheckedOutByUser) {
                $checkedOutBy = if ($file.CheckedOutByUser.Email) {
                    "{0} <{1}>" -f $file.CheckedOutByUser.Name, $file.CheckedOutByUser.Email
                }
                else {
                    $file.CheckedOutByUser.Name
                }
            }

            [pscustomobject]@{
                SiteCollectionUrl = $Web.Site.Url
                WebUrl            = $Web.Url
                WebTitle          = $Web.Title
                LibraryTitle      = $Library.Title
                LibraryUrl        = $Library.RootFolder.ServerRelativeUrl
                ItemId            = $item.ID
                UniqueId          = $item.UniqueId
                FileName          = [string]$item['FileLeafRef']
                FileUrl           = $Web.Site.MakeFullUrl([string]$item['FileRef'])
                ServerRelativeUrl = [string]$item['FileRef']
                Extension         = [System.IO.Path]::GetExtension([string]$item['FileLeafRef'])
                SizeBytes         = $fileSizeBytes
                SizeMB            = if ($null -ne $fileSizeBytes) { [math]::Round($fileSizeBytes / 1MB, 2) } else { $null }
                Created           = $item['Created']
                CreatedBy         = Convert-SPFieldUserValueToString -Value $item['Author'] -Web $Web
                Modified          = $item['Modified']
                ModifiedBy        = Convert-SPFieldUserValueToString -Value $item['Editor'] -Web $Web
                ContentType       = $item.ContentType.Name
                Version           = $file.UIVersionLabel
                VersionsCount     = Get-ObjectCount -Value $file.Versions
                CheckedOutBy      = $checkedOutBy
            }
        }
    }
    while ($null -ne $query.ListItemCollectionPosition)
}

function Export-WebInventory {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
        [string]$CsvPath,
        [switch]$Append,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [int]$BatchSize
    )

    Write-Host ("Web: {0}" -f $Web.Url)

    foreach ($library in (Get-DocumentLibraries -Web $Web -IncludeHidden:$IncludeHidden -IncludeSystem:$IncludeSystem)) {
        try {
            Write-Host ("  Library: {0} ({1} items)" -f $library.Title, $library.ItemCount)

            $rows = Get-FileInventoryFromLibrary -Web $Web -Library $library -BatchSize $BatchSize
            if ($Append) {
                $rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Append
            }
            else {
                $rows | Export-Csv -Delimiter ';' -Path $CsvPath -NoTypeInformation -Encoding UTF8
                $script:CsvCreated = $true
                $Append = $true
            }
        }
        catch {
            if ($script:InventoryStopRequested) {
                throw
            }
            Write-Warning ("Failed to inventory library '{0}' in web '{1}': {2}" -f $library.Title, $Web.Url, $_.Exception.Message)
            Stop-InventoryAfterError -Scope 'Library' -Url $Web.Url -Name $library.Title -Message $_.Exception.Message
        }
    }
}

function Export-SiteInventory {
    param(
        [Microsoft.SharePoint.SPSite]$Site,
        [string]$CsvPath,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [int]$BatchSize
    )

    Write-Host ("Site collection: {0}" -f $Site.Url)

    try {
        foreach ($web in $Site.AllWebs) {
            try {
                if (-not (Test-WebMatchesSiteUrlFilter -Web $web)) {
                    Write-Host ("Skipping web not in filter: {0}" -f $web.Url)
                    continue
                }

                Export-WebInventory `
                    -Web $web `
                    -CsvPath $CsvPath `
                    -Append:$script:CsvCreated `
                    -IncludeHidden:$IncludeHidden `
                    -IncludeSystem:$IncludeSystem `
                    -BatchSize $BatchSize
            }
            catch {
                if ($script:InventoryStopRequested) {
                    throw
                }
                Write-Warning ("Failed to inventory web '{0}': {1}" -f $web.Url, $_.Exception.Message)
                Stop-InventoryAfterError -Scope 'Web' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
            }
            finally {
                $web.Dispose()
            }
        }
    }
    catch {
        if ($script:InventoryStopRequested) {
            throw
        }
        Write-Warning ("Failed to enumerate webs for site collection '{0}': {1}" -f $Site.Url, $_.Exception.Message)
        Stop-InventoryAfterError -Scope 'SiteCollection' -Url $Site.Url -Name $Site.Url -Message $_.Exception.Message
    }
}

function ConvertTo-SafeFileName {
    param(
        [string]$Name,
        [string]$Fallback = 'SharePointSite'
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

function Write-InventoryHeaderOnlyCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $header = ($InventoryColumns | ForEach-Object { '"{0}"' -f ($_ -replace '"', '""') }) -join ';'
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
        $uri = New-Object System.Uri($cleanUrl)
        $path = $uri.AbsolutePath
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
                $webApplication = Get-SPWebApplication -Identity $WebApplicationUrl
                $targetName = $webApplication.Name
            }
            catch {
                $targetName = Get-NameFromUrl -Url $WebApplicationUrl
            }
        }

        'Site' {
            $site = Get-SPSite -Identity $SiteUrl
            try {
                $rootWeb = $site.RootWeb
                try {
                    $targetName = $rootWeb.Title
                }
                finally {
                    $rootWeb.Dispose()
                }
            }
            finally {
                $site.Dispose()
            }

            if ([string]::IsNullOrWhiteSpace($targetName)) {
                $targetName = Get-NameFromUrl -Url $SiteUrl
            }
        }

        'Web' {
            $site = New-Object Microsoft.SharePoint.SPSite($WebUrl)
            try {
                $web = $site.OpenWeb()
                try {
                    $targetName = $web.Title
                }
                finally {
                    $web.Dispose()
                }
            }
            finally {
                $site.Dispose()
            }

            if ([string]::IsNullOrWhiteSpace($targetName)) {
                $targetName = Get-NameFromUrl -Url $WebUrl
            }
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
    return Join-Path -Path $targetDirectory -ChildPath ("SP2019-FileInventory-{0}-{1:yyyyMMdd-HHmmss}.csv" -f $safeTargetName, (Get-Date))
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

function Stop-InventoryAfterError {
    param(
        [string]$Scope,
        [string]$Url,
        [string]$Name,
        [string]$Message
    )

    Write-InventoryError -Scope $Scope -Url $Url -Name $Name -Message $Message
    $script:InventoryStopRequested = $true
    throw ("Inventory stopped after {0} error at '{1}'. Final CSV was not published. Details: {2}" -f $Scope, $Url, $Message)
}

function Get-PermissionInventoryOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileInventoryOutputPath
    )

    $directory = Split-Path -Path $FileInventoryOutputPath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileInventoryOutputPath)
    if ($baseName -match 'FileInventory') {
        $permissionBaseName = $baseName -replace 'FileInventory', 'PermissionInventory'
    }
    else {
        $permissionBaseName = "{0}-Permissions" -f $baseName
    }

    return Join-Path -Path $directory -ChildPath ("{0}.csv" -f $permissionBaseName)
}

function Invoke-PermissionInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileInventoryOutputPath
    )

    $scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $permissionScriptPath = Join-Path -Path $scriptDirectory -ChildPath 'SmartM365-SharePointSource-PermissionInventory.ps1'
    if (-not (Test-Path -LiteralPath $permissionScriptPath)) {
        throw "Permission inventory script not found: $permissionScriptPath"
    }

    $permissionOutputPath = Get-PermissionInventoryOutputPath -FileInventoryOutputPath $FileInventoryOutputPath
    $parameters = @{
        OutputPath            = $permissionOutputPath
        RowLimit              = $RowLimit
        DocumentLibrariesOnly = $true
        IncludeHiddenLists    = $IncludeHiddenLibraries
        IncludeSystemLists    = $IncludeSystemLibraries
    }

    switch ($PSCmdlet.ParameterSetName) {
        'WebApplication' {
            $parameters.WebApplicationUrl = $WebApplicationUrl
        }

        'Site' {
            $parameters.SiteUrl = $SiteUrl
        }

        'Web' {
            $parameters.WebUrl = $WebUrl
        }
    }

    if ($script:UseSiteUrlFilter) {
        $parameters.UseSiteUrlFilter = $true
        $parameters.SiteUrlsFile = $SiteUrlsFile
    }

    if ($IncludePermissionItemPermissions) {
        $parameters.IncludeItemPermissions = $true
    }

    Write-Host ("Starting permission inventory: {0}" -f $permissionOutputPath)
    & $permissionScriptPath @parameters
    if (-not $?) {
        throw "Permission inventory failed."
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

$errorDirectory = Split-Path -Path $ErrorPath -Parent
if ($errorDirectory -and -not (Test-Path -LiteralPath $errorDirectory)) {
    New-Item -Path $errorDirectory -ItemType Directory -Force | Out-Null
}

$logDirectory = Split-Path -Path $LogPath -Parent
if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

if (Test-Path -LiteralPath $TempOutputPath) {
    Remove-Item -LiteralPath $TempOutputPath -Force
}

if (Test-Path -LiteralPath $ErrorPath) {
    Remove-Item -LiteralPath $ErrorPath -Force
}

if (Test-Path -LiteralPath $LogPath) {
    Remove-Item -LiteralPath $LogPath -Force
}

$script:ErrorPath = $ErrorPath
$script:CsvCreated = $false
$script:ErrorCsvCreated = $false
$script:TranscriptStarted = $false
$script:InventoryStopRequested = $false
$script:InventoryExitCode = 0
$script:InventoryStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
    Write-Host ("Inventory output: {0}" -f $OutputPath)
    Write-Host ("Temporary inventory output: {0}" -f $TempOutputPath)
    Write-Host ("Error output: {0}" -f $ErrorPath)
    Write-Host ("Run log: {0}" -f $LogPath)
    $inputScope = switch ($PSCmdlet.ParameterSetName) {
        'WebApplication' { $WebApplicationUrl }
        'Site' { $SiteUrl }
        'Web' { $WebUrl }
    }
    $siteUrlsFileLabel = if ([string]::IsNullOrWhiteSpace($SiteUrlsFile)) { '<none>' } else { $SiteUrlsFile }
    Write-Host ("File inventory options: IncludeHiddenLibraries={0}; IncludeSystemLibraries={1}; IncludePermissionInventory={2}; IncludePermissionItemPermissions={3}; RowLimit={4}; ParameterSet={5}; Input={6}; UseSiteUrlFilter={7}; SiteUrlsFile={8}" -f [bool]$IncludeHiddenLibraries, [bool]$IncludeSystemLibraries, [bool]$IncludePermissionInventory, [bool]$IncludePermissionItemPermissions, $RowLimit, $PSCmdlet.ParameterSetName, $inputScope, [bool]$script:UseSiteUrlFilter, $siteUrlsFileLabel)
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
                    Export-SiteInventory `
                        -Site $site `
                        -CsvPath $TempOutputPath `
                        -IncludeHidden:$IncludeHiddenLibraries `
                        -IncludeSystem:$IncludeSystemLibraries `
                        -BatchSize $RowLimit
                }
                catch {
                    if ($script:InventoryStopRequested) {
                        throw
                    }
                    Write-Warning ("Failed to inventory site collection '{0}': {1}" -f $site.Url, $_.Exception.Message)
                    Stop-InventoryAfterError -Scope 'SiteCollection' -Url $site.Url -Name $site.Url -Message $_.Exception.Message
                }
                finally {
                    $site.Dispose()
                }
            }
        }

        'Site' {
            $site = Get-SPSite -Identity $SiteUrl
            try {
                Export-SiteInventory `
                    -Site $site `
                    -CsvPath $TempOutputPath `
                    -IncludeHidden:$IncludeHiddenLibraries `
                    -IncludeSystem:$IncludeSystemLibraries `
                    -BatchSize $RowLimit
            }
            catch {
                if ($script:InventoryStopRequested) {
                    throw
                }
                Write-Warning ("Failed to inventory site collection '{0}': {1}" -f $site.Url, $_.Exception.Message)
                Stop-InventoryAfterError -Scope 'SiteCollection' -Url $site.Url -Name $site.Url -Message $_.Exception.Message
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
                        Export-WebInventory `
                            -Web $web `
                            -CsvPath $TempOutputPath `
                            -Append:$script:CsvCreated `
                            -IncludeHidden:$IncludeHiddenLibraries `
                            -IncludeSystem:$IncludeSystemLibraries `
                            -BatchSize $RowLimit
                    }
                    else {
                        Write-Host ("Skipping web not in filter: {0}" -f $web.Url)
                    }
                }
                catch {
                    if ($script:InventoryStopRequested) {
                        throw
                    }
                    Write-Warning ("Failed to inventory web '{0}': {1}" -f $web.Url, $_.Exception.Message)
                    Stop-InventoryAfterError -Scope 'Web' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
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
        Write-InventoryHeaderOnlyCsv -Path $TempOutputPath
        Write-Warning "No file rows were exported. The final inventory CSV contains headers only."
    }

    if ($script:ErrorCsvCreated) {
        throw ("Inventory errors were recorded. Final CSV was not published. Error details: {0}" -f $ErrorPath)
    }

    Move-Item -LiteralPath $TempOutputPath -Destination $OutputPath -Force
    $script:InventoryStopwatch.Stop()
    Write-Host ("Inventory completed: {0}" -f $OutputPath)
    Write-Host ("Scan duration: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
    if ($script:ErrorCsvCreated) {
        Write-Warning ("Some items could not be inventoried. Error details: {0}" -f $ErrorPath)
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
        Write-Warning ("Inventory failed before final CSV publication. Partial temporary CSV kept: {0}" -f $TempOutputPath)
    }
    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
        Write-Warning ("Removed incomplete final CSV: {0}" -f $OutputPath)
    }
    $script:InventoryExitCode = 1
}
finally {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
}

if ($script:InventoryExitCode -ne 0) {
    exit $script:InventoryExitCode
}

if ($IncludePermissionInventory) {
    try {
        Invoke-PermissionInventory -FileInventoryOutputPath $OutputPath
    }
    catch {
        Write-Error $_
        exit 1
    }
}
