<#
.SYNOPSIS
    Inventories files from SharePoint Online document libraries.

.DESCRIPTION
    Uses the PnP.PowerShell module to scan SharePoint Online document libraries
    and export file metadata to a CSV file.

    The script can target:
    - A whole tenant, through the SharePoint admin URL
    - A site collection, including all subsites
    - A single web
    - A list of web URLs from a text or CSV file
      Each listed web URL is treated as a root and its descendant subsites
      are included by default.

    If OutputPath is not specified, the script creates a folder in the script
    directory using the target site or tenant name, then writes the CSV, run
    log, and error CSV in that folder.

.NOTES
    PnP.PowerShell now requires your own Entra ID app registration for most
    interactive scenarios. Provide -ClientId, or configure a default client ID
    through PnP.PowerShell environment/default-client-id settings.

.EXAMPLE
    .\SmartM365-SharePointTarget-FileInventory.ps1 -TenantAdminUrl "https://yourtenant-admin.sharepoint.com"

.EXAMPLE
    .\SmartM365-SharePointTarget-FileInventory.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/finance"

.EXAMPLE
    .\SmartM365-SharePointTarget-FileInventory.ps1 -WebUrl "https://yourtenant.sharepoint.com/sites/finance/accounting"

.EXAMPLE
    .\SmartM365-SharePointTarget-FileInventory.ps1 -WebUrlsFile "C:\Temp\spo-web-urls.txt" -Interactive

.EXAMPLE
    .\SmartM365-SharePointTarget-FileInventory.ps1 -TenantAdminUrl "https://yourtenant-admin.sharepoint.com" -UseEnvironmentVariables

.VERSION
    1.0.2
#>

[CmdletBinding(DefaultParameterSetName = 'Tenant')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Tenant')]
    [ValidateNotNullOrEmpty()]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Site')]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Web')]
    [ValidateNotNullOrEmpty()]
    [string]$WebUrl,

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

    [switch]$DeviceLogin,

    [switch]$Interactive,

    [switch]$PersistLogin,

    [switch]$ForceAuthentication,

    [switch]$UseEnvironmentVariables,

    [switch]$ManagedIdentity,

    [ValidateNotNullOrEmpty()]
    [string]$CertificatePath,

    [System.Security.SecureString]$CertificatePassword,

    [string]$Thumbprint = $env:SPO_INVENTORY_CERT_THUMBPRINT,

    [switch]$IncludeHiddenLibraries,

    [switch]$IncludeSystemLibraries,

    [switch]$IncludeOneDriveSites,

    [int]$PageSize = 2000
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

function Import-PnPPowerShellModule {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "PnP.PowerShell is not installed. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }

    Import-Module PnP.PowerShell -ErrorAction Stop
}

function Write-Info {
    param(
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    Write-Host $Message -ForegroundColor $Color
}

function Test-CertificateThumbprint {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint
    )

    $normalizedThumbprint = ($CertificateThumbprint -replace '\s', '').ToUpperInvariant()
    $certificate = Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
        Where-Object { ($_.Thumbprint -replace '\s', '').ToUpperInvariant() -eq $normalizedThumbprint } |
        Select-Object -First 1

    if (-not $certificate) {
        throw "Certificate thumbprint '$CertificateThumbprint' was not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My."
    }

    if (-not $certificate.HasPrivateKey) {
        throw "Certificate thumbprint '$CertificateThumbprint' was found, but it does not include a private key. Import the .pfx file, not only the .cer file."
    }

    if ($certificate.NotAfter -lt (Get-Date)) {
        throw "Certificate thumbprint '$CertificateThumbprint' expired on $($certificate.NotAfter)."
    }
}

function Connect-SPOInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $parameters = @{
        Url              = $Url
        ReturnConnection = $true
        ErrorAction      = 'Stop'
    }

    $forceAuthenticationForThisConnection = $ForceAuthentication -and -not $script:ForceAuthenticationAlreadyUsed

    if ($UseEnvironmentVariables) {
        $parameters.EnvironmentVariable = $true
        $authenticationMode = 'Environment variables'
    }
    elseif ($ManagedIdentity) {
        $parameters.ManagedIdentity = $true
        $authenticationMode = 'Managed identity'
    }
    elseif ($Interactive) {
        $parameters.Interactive = $true
        $authenticationMode = 'Interactive'
        if (($PersistLogin -or $script:UsePersistedLoginForRun) -and -not $forceAuthenticationForThisConnection) {
            $parameters.PersistLogin = $true
        }

        if ($ClientId) {
            $parameters.ClientId = $ClientId
        }
    }
    elseif ($DeviceLogin) {
        if (-not $Tenant) {
            throw "Device login requires -Tenant, for example yourtenant.onmicrosoft.com."
        }

        $parameters.DeviceLogin = $true
        $parameters.Tenant = $Tenant
        $authenticationMode = 'Device login'
        if ($PersistLogin) {
            $parameters.PersistLogin = $true
        }

        if ($ClientId) {
            $parameters.ClientId = $ClientId
        }
    }
    elseif ($CertificatePath) {
        if (-not $ClientId -or -not $Tenant) {
            throw "Certificate authentication requires -ClientId and -Tenant."
        }

        $parameters.ClientId = $ClientId
        $parameters.Tenant = if ($TenantId) { $TenantId } else { $Tenant }
        $parameters.CertificatePath = $CertificatePath
        $authenticationMode = 'Certificate path'
        if ($CertificatePassword) {
            $parameters.CertificatePassword = $CertificatePassword
        }
    }
    elseif ($Thumbprint) {
        if (-not $ClientId -or -not $Tenant) {
            throw "Certificate thumbprint authentication requires -ClientId and -Tenant."
        }

        Test-CertificateThumbprint -CertificateThumbprint $Thumbprint

        $parameters.ClientId = $ClientId
        $parameters.Tenant = if ($TenantId) { $TenantId } else { $Tenant }
        $parameters.Thumbprint = $Thumbprint
        $authenticationMode = 'Certificate thumbprint'
    }
    else {
        $parameters.Interactive = $true
        $authenticationMode = 'Interactive'
        if (($PersistLogin -or $script:UsePersistedLoginForRun) -and -not $forceAuthenticationForThisConnection) {
            $parameters.PersistLogin = $true
        }

        if ($ClientId) {
            $parameters.ClientId = $ClientId
        }
    }

    if ($forceAuthenticationForThisConnection -and ($parameters.ContainsKey('Interactive') -or $parameters.ContainsKey('OSLogin'))) {
        $parameters.ForceAuthentication = $true
    }

    if ($parameters.ContainsKey('ForceAuthentication') -and -not $script:PersistedLoginCleared) {
        try {
            Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction SilentlyContinue
            $script:PersistedLoginCleared = $true
            Write-Info -Color DarkCyan -Message "Cleared persisted PnP login before forced authentication."
        }
        catch {
            Write-Warning ("Could not clear persisted PnP login before forced authentication: {0}" -f $_.Exception.Message)
        }
    }

    Write-Info -Color Cyan -Message ("Authentication mode for {0}: {1}" -f $Url, $authenticationMode)

    $connection = Connect-PnPOnline @parameters
    if ($parameters.ContainsKey('ForceAuthentication')) {
        $script:ForceAuthenticationAlreadyUsed = $true
    }

    Write-SPOConnectionIdentity -Connection $connection -Url $Url
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

        Write-Info -Color Cyan -Message ("Connected account for {0}: {1} | {2} | {3}" -f $Url, $loginName, $email, $title)
    }
    catch {
        Write-Warning ("Could not determine connected account for '{0}': {1}" -f $Url, $_.Exception.Message)
    }
}

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

    $uri = New-Object System.Uri($WebUrl)
    return ("{0}://{1}{2}" -f $uri.Scheme, $uri.Host, $ServerRelativeUrl)
}

function Get-WebUrlsFromFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Web URLs file not found: $Path"
    }

    $extension = [System.IO.Path]::GetExtension($Path)
    if ($extension -ieq '.csv') {
        $firstLine = Get-Content -LiteralPath $Path -TotalCount 1
        $delimiter = if (([regex]::Matches($firstLine, ';')).Count -gt ([regex]::Matches($firstLine, ',')).Count) { ';' } else { ',' }
        $rows = Import-Csv -LiteralPath $Path -Delimiter $delimiter
        foreach ($row in $rows) {
            $propertyName = @('Url', 'WebUrl', 'SiteUrl') | Where-Object { $row.PSObject.Properties.Name -contains $_ } | Select-Object -First 1
            if ($propertyName -and -not [string]::IsNullOrWhiteSpace($row.$propertyName)) {
                $row.$propertyName.Trim()
            }
        }
    }
    else {
        Get-Content -LiteralPath $Path |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') }
    }
}

function Convert-PnPFieldUserValueToString {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [array]) {
        return (($Value | ForEach-Object { Convert-PnPFieldUserValueToString -Value $_ }) -join '; ')
    }

    if ($Value.PSObject.Properties.Name -contains 'Email' -and $Value.Email) {
        if ($Value.PSObject.Properties.Name -contains 'LookupValue' -and $Value.LookupValue) {
            return ("{0} <{1}>" -f $Value.LookupValue, $Value.Email)
        }

        return $Value.Email
    }

    if ($Value.PSObject.Properties.Name -contains 'LookupValue' -and $Value.LookupValue) {
        return $Value.LookupValue
    }

    return [string]$Value
}

function Test-SystemLibrary {
    param(
        $List
    )

    $systemLibraryUrls = @(
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

    $rootFolderUrl = $null
    if ($List.RootFolder -and ($List.RootFolder.PSObject.Properties.Name -contains 'ServerRelativeUrl')) {
        $rootFolderUrl = $List.RootFolder.ServerRelativeUrl
    }

    foreach ($url in $systemLibraryUrls) {
        if ($List.Title -eq $url -or $rootFolderUrl -like "*/$url" -or $rootFolderUrl -like "*/$url/*") {
            return $true
        }
    }

    return $false
}

function Get-DocumentLibraries {
    param(
        $Connection,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [switch]$WriteSummary
    )

    $lists = Get-PnPList -Includes BaseType,Hidden,Title,ItemCount,RootFolder,IsSystemList -Connection $Connection
    $totalDocumentLibraries = 0
    $selectedDocumentLibraries = 0
    $hiddenSkipped = 0
    $systemSkipped = 0
    $emptySkipped = 0

    foreach ($list in $lists) {
        if ([string]$list.BaseType -ne 'DocumentLibrary') {
            continue
        }

        $totalDocumentLibraries++

        if (-not $IncludeHidden -and $list.Hidden) {
            $hiddenSkipped++
            continue
        }

        if (-not $IncludeSystem -and (Test-SystemLibrary -List $list)) {
            $systemSkipped++
            continue
        }

        if ($list.ItemCount -eq 0) {
            $emptySkipped++
            continue
        }

        $selectedDocumentLibraries++
        $list
    }

    if ($WriteSummary) {
        Write-Info -Color DarkCyan -Message ("  Document libraries visible: {0}; selected for inventory: {1}; skipped hidden: {2}; skipped system: {3}; skipped empty: {4}" -f $totalDocumentLibraries, $selectedDocumentLibraries, $hiddenSkipped, $systemSkipped, $emptySkipped)
    }
}

function Get-FileInventoryFromLibrary {
    param(
        $Connection,
        [string]$SiteCollectionUrl,
        $Web,
        $Library,
        [int]$BatchSize
    )

    $lastItemId = 0
    $itemsScanned = 0
    $filesReturned = 0
    $useSimplePagingFallback = $false

    do {
        if ($useSimplePagingFallback) {
            $items = @(Get-PnPListItem `
                    -List $Library `
                    -PageSize $BatchSize `
                    -Fields 'ID', 'FSObjType', 'FileLeafRef', 'FileRef', 'UniqueId', 'File_x0020_Size', 'Created', 'Author', 'Modified', 'Editor', 'ContentType', '_UIVersionString', 'CheckoutUser' `
                    -IncludeContentType `
                    -Connection $Connection)
        }
        else {
            $query = @"
<View Scope='RecursiveAll'>
  <ViewFields>
    <FieldRef Name='ID' />
    <FieldRef Name='FSObjType' />
    <FieldRef Name='FileLeafRef' />
    <FieldRef Name='FileRef' />
    <FieldRef Name='UniqueId' />
    <FieldRef Name='File_x0020_Size' />
    <FieldRef Name='Created' />
    <FieldRef Name='Author' />
    <FieldRef Name='Modified' />
    <FieldRef Name='Editor' />
    <FieldRef Name='ContentType' />
    <FieldRef Name='_UIVersionString' />
    <FieldRef Name='CheckoutUser' />
  </ViewFields>
  <Query>
    <Where>
      <Gt>
        <FieldRef Name='ID' />
        <Value Type='Counter'>$lastItemId</Value>
      </Gt>
    </Where>
    <OrderBy Override='TRUE'>
      <FieldRef Name='ID' Ascending='TRUE' />
    </OrderBy>
  </Query>
  <RowLimit Paged='TRUE'>$BatchSize</RowLimit>
</View>
"@

            $items = @(Get-PnPListItem -List $Library -Query $query -PageSize $BatchSize -IncludeContentType -Connection $Connection)
            if ($items.Count -eq 0 -and $lastItemId -eq 0 -and $Library.ItemCount -gt 0) {
                Write-Warning ("    ID CAML paging returned 0 items for non-empty library '{0}'. Retrying with simple PnP paging." -f $Library.Title)
                $useSimplePagingFallback = $true
                continue
            }
        }

        if ($items.Count -eq 0) {
            break
        }

        foreach ($item in $items) {
            if ($item.Id -gt $lastItemId) {
                $lastItemId = $item.Id
            }

            $values = $item.FieldValues
            $itemsScanned++

            if ($values.ContainsKey('FSObjType') -and [string]$values['FSObjType'] -ne '0') {
                continue
            }

            $fileName = [string]$values['FileLeafRef']
            $serverRelativeUrl = [string]$values['FileRef']
            $fileSizeBytes = $null

            if ([string]::IsNullOrWhiteSpace($serverRelativeUrl)) {
                continue
            }

            if ($values.ContainsKey('File_x0020_Size') -and $null -ne $values['File_x0020_Size']) {
                [long]$fileSizeBytes = $values['File_x0020_Size']
            }

            $filesReturned++

            [pscustomobject]@{
                SiteCollectionUrl = $SiteCollectionUrl
                WebUrl            = $Web.Url
                WebTitle          = $Web.Title
                LibraryTitle      = $Library.Title
                LibraryUrl        = $Library.RootFolder.ServerRelativeUrl
                ItemId            = $item.Id
                UniqueId          = if ($values.ContainsKey('UniqueId')) { $values['UniqueId'] } else { $null }
                FileName          = $fileName
                FileUrl           = ConvertTo-AbsoluteSharePointUrl -WebUrl $Web.Url -ServerRelativeUrl $serverRelativeUrl
                ServerRelativeUrl = $serverRelativeUrl
                Extension         = [System.IO.Path]::GetExtension($fileName)
                SizeBytes         = $fileSizeBytes
                SizeMB            = if ($null -ne $fileSizeBytes) { [math]::Round($fileSizeBytes / 1MB, 2) } else { $null }
                Created           = if ($values.ContainsKey('Created')) { $values['Created'] } else { $null }
                CreatedBy         = if ($values.ContainsKey('Author')) { Convert-PnPFieldUserValueToString -Value $values['Author'] } else { $null }
                Modified          = if ($values.ContainsKey('Modified')) { $values['Modified'] } else { $null }
                ModifiedBy        = if ($values.ContainsKey('Editor')) { Convert-PnPFieldUserValueToString -Value $values['Editor'] } else { $null }
                ContentType       = if (($item.PSObject.Properties.Name -contains 'ContentType') -and $item.ContentType) { $item.ContentType.Name } elseif ($values.ContainsKey('ContentType')) { $values['ContentType'] } else { $null }
                Version           = if ($values.ContainsKey('_UIVersionString')) { $values['_UIVersionString'] } else { $null }
                VersionsCount     = $null
                CheckedOutBy      = if ($values.ContainsKey('CheckoutUser')) { Convert-PnPFieldUserValueToString -Value $values['CheckoutUser'] } else { $null }
            }
        }

        if ($useSimplePagingFallback) {
            break
        }
    }
    while ($items.Count -gt 0)

    Write-Info -Color DarkGray -Message ("    Items scanned by ID paging: {0}; files exported: {1}" -f $itemsScanned, $filesReturned)
}

function Get-ConnectedSiteCollectionUrl {
    param(
        $Connection,
        [string]$FallbackUrl
    )

    try {
        $context = Get-PnPContext -Connection $Connection
        $site = $context.Site
        $context.Load($site)
        $context.ExecuteQuery()
        if (-not [string]::IsNullOrWhiteSpace($site.Url)) {
            return $site.Url
        }
    }
    catch {
        Write-Warning ("Could not determine site collection URL for '{0}': {1}" -f $FallbackUrl, $_.Exception.Message)
    }

    return $FallbackUrl
}

function Export-WebInventory {
    param(
        [string]$Url,
        [string]$SiteCollectionUrl,
        [string]$CsvPath,
        [switch]$Append,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [int]$BatchSize
    )

    $connection = Connect-SPOInventory -Url $Url
    $web = Get-PnPWeb -Includes Title,Url,ServerRelativeUrl -Connection $connection
    if ([string]::IsNullOrWhiteSpace($SiteCollectionUrl)) {
        $SiteCollectionUrl = $Url
    }
    $SiteCollectionUrl = Get-ConnectedSiteCollectionUrl -Connection $connection -FallbackUrl $SiteCollectionUrl

    Write-Info -Color Blue -Message ("Web: {0}" -f $web.Url)

    $libraries = @(Get-DocumentLibraries -Connection $connection -IncludeHidden:$IncludeHidden -IncludeSystem:$IncludeSystem -WriteSummary)
    if ($libraries.Count -eq 0) {
        Write-Warning ("No non-empty document libraries selected for web '{0}'." -f $web.Url)
    }

    foreach ($library in $libraries) {
        try {
            Write-Info -Color Yellow -Message ("  Library: {0} ({1} items)" -f $library.Title, $library.ItemCount)

            $rows = Get-FileInventoryFromLibrary -Connection $connection -SiteCollectionUrl $SiteCollectionUrl -Web $web -Library $library -BatchSize $BatchSize
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
            Write-Warning ("Failed to inventory library '{0}' in web '{1}': {2}" -f $library.Title, $web.Url, $_.Exception.Message)
            Stop-InventoryAfterError -Scope 'Library' -Url $web.Url -Name $library.Title -Message $_.Exception.Message
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
            $connection = Connect-SPOInventory -Url $currentWebUrl
            $web = Get-PnPWeb -Includes Title,Url,ServerRelativeUrl -Connection $connection
            $webUrls.Add($web.Url)

            try {
                $subWebs = @(Get-PnPSubWeb -Includes Title,Url,ServerRelativeUrl -Connection $connection)
                Write-Info -Color DarkCyan -Message ("  Subsites found under {0}: {1}" -f $web.Url, $subWebs.Count)

                foreach ($subWeb in $subWebs) {
                    if ($subWeb.Url -and -not $seenWebUrls.Contains($subWeb.Url)) {
                        $pendingWebUrls.Enqueue($subWeb.Url)
                    }
                }
            }
            catch {
                if ($script:InventoryStopRequested) {
                    throw
                }
                Write-Warning ("Failed to enumerate immediate subsites for web '{0}': {1}" -f $web.Url, $_.Exception.Message)
                Stop-InventoryAfterError -Scope 'SubsiteEnumeration' -Url $web.Url -Name $web.Title -Message $_.Exception.Message
            }
        }
        catch {
            if ($script:InventoryStopRequested) {
                throw
            }
            Write-Warning ("Failed to connect to or read web '{0}': {1}" -f $currentWebUrl, $_.Exception.Message)
            Stop-InventoryAfterError -Scope 'Web' -Url $currentWebUrl -Name $currentWebUrl -Message $_.Exception.Message
        }
    }

    $webUrls
}

function Export-SiteInventory {
    param(
        [string]$Url,
        [string]$CsvPath,
        [switch]$IncludeHidden,
        [switch]$IncludeSystem,
        [int]$BatchSize
    )

    Write-Info -Color Magenta -Message ("Site collection: {0}" -f $Url)

    try {
        foreach ($targetWebUrl in (Get-WebUrlsFromSite -Url $Url)) {
            try {
                Export-WebInventory `
                    -Url $targetWebUrl `
                    -SiteCollectionUrl $Url `
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
                Write-Warning ("Failed to inventory web '{0}': {1}" -f $targetWebUrl, $_.Exception.Message)
                Stop-InventoryAfterError -Scope 'Web' -Url $targetWebUrl -Name $targetWebUrl -Message $_.Exception.Message
            }
        }
    }
    catch {
        if ($script:InventoryStopRequested) {
            throw
        }
        Write-Warning ("Failed to enumerate webs for site collection '{0}': {1}" -f $Url, $_.Exception.Message)
        Stop-InventoryAfterError -Scope 'SiteCollection' -Url $Url -Name $Url -Message $_.Exception.Message
    }
}

function Get-DefaultOutputPath {
    param(
        [string]$ParameterSetName,
        [string]$TenantAdminUrl,
        [string]$SiteUrl,
        [string]$WebUrl,
        [string]$WebUrlsFile
    )

    $scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $targetName = $null

    switch ($ParameterSetName) {
        'Tenant' {
            $targetName = Get-NameFromUrl -Url $TenantAdminUrl
            $targetName = $targetName -replace '-admin\.sharepoint\.com$', ''
        }

        'Site' {
            try {
                $connection = Connect-SPOInventory -Url $SiteUrl
                $web = Get-PnPWeb -Includes Title,Url -Connection $connection
                $targetName = $web.Title
            }
            catch {
                $targetName = Get-NameFromUrl -Url $SiteUrl
            }
        }

        'Web' {
            try {
                $connection = Connect-SPOInventory -Url $WebUrl
                $web = Get-PnPWeb -Includes Title,Url -Connection $connection
                $targetName = $web.Title
            }
            catch {
                $targetName = Get-NameFromUrl -Url $WebUrl
            }
        }

        'WebUrlsFile' {
            $targetName = [System.IO.Path]::GetFileNameWithoutExtension($WebUrlsFile)
        }
    }

    $safeTargetName = ConvertTo-SafeFileName -Name $targetName
    $targetDirectory = Join-Path -Path $scriptDirectory -ChildPath $safeTargetName
    return Join-Path -Path $targetDirectory -ChildPath ("SPO-FileInventory-{0}-{1:yyyyMMdd-HHmmss}.csv" -f $safeTargetName, (Get-Date))
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

Import-PnPPowerShellModule

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Get-DefaultOutputPath `
        -ParameterSetName $PSCmdlet.ParameterSetName `
        -TenantAdminUrl $TenantAdminUrl `
        -SiteUrl $SiteUrl `
        -WebUrl $WebUrl `
        -WebUrlsFile $WebUrlsFile
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
$script:ForceAuthenticationAlreadyUsed = $false
$script:PersistedLoginCleared = $false
$script:UsePersistedLoginForRun = ($PSCmdlet.ParameterSetName -eq 'WebUrlsFile')
$script:InventoryStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
    $script:TranscriptStarted = $true
    Write-Info -Color Cyan -Message ("Inventory output: {0}" -f $OutputPath)
    Write-Info -Color Cyan -Message ("Temporary inventory output: {0}" -f $TempOutputPath)
    Write-Info -Color Cyan -Message ("Error output: {0}" -f $ErrorPath)
    Write-Info -Color Cyan -Message ("Run log: {0}" -f $LogPath)
    $authMode = if ($ManagedIdentity) { 'ManagedIdentity' } elseif (-not [string]::IsNullOrWhiteSpace($CertificatePath)) { 'CertificatePath' } elseif (-not [string]::IsNullOrWhiteSpace($Thumbprint)) { 'Certificate' } elseif ($DeviceLogin) { 'DeviceLogin' } elseif ($Interactive) { 'Interactive' } elseif ($UseEnvironmentVariables) { 'EnvironmentVariables' } else { 'Interactive' }
    $inputScope = switch ($PSCmdlet.ParameterSetName) {
        'Tenant' { $TenantAdminUrl }
        'Site' { $SiteUrl }
        'Web' { $WebUrl }
        'WebUrlsFile' { $WebUrlsFile }
    }
    Write-Info -Color Cyan -Message ("File inventory options: IncludeHiddenLibraries={0}; IncludeSystemLibraries={1}; IncludeOneDriveSites={2}; PageSize={3}; AuthMode={4}; ForceAuthentication={5}; PersistLogin={6}; ParameterSet={7}; Input={8}" -f [bool]$IncludeHiddenLibraries, [bool]$IncludeSystemLibraries, [bool]$IncludeOneDriveSites, $PageSize, $authMode, [bool]$ForceAuthentication, [bool]$PersistLogin, $PSCmdlet.ParameterSetName, $inputScope)
}
catch {
    Write-Warning ("Could not start transcript log '{0}': {1}" -f $LogPath, $_.Exception.Message)
}

try {
    switch ($PSCmdlet.ParameterSetName) {
        'Tenant' {
            $adminConnection = Connect-SPOInventory -Url $TenantAdminUrl
            $tenantSiteParameters = @{
                Detailed   = $true
                Connection = $adminConnection
            }

            if ($IncludeOneDriveSites) {
                $tenantSiteParameters.IncludeOneDriveSites = $true
            }

            $tenantSites = Get-PnPTenantSite @tenantSiteParameters

            if (-not $IncludeOneDriveSites) {
                $tenantSites = $tenantSites | Where-Object { $_.Url -notmatch '-my\.sharepoint\.com/personal/' }
            }

            foreach ($site in $tenantSites) {
                try {
                    Export-SiteInventory `
                        -Url $site.Url `
                        -CsvPath $TempOutputPath `
                        -IncludeHidden:$IncludeHiddenLibraries `
                        -IncludeSystem:$IncludeSystemLibraries `
                        -BatchSize $PageSize
                }
                catch {
                    if ($script:InventoryStopRequested) {
                        throw
                    }
                    Write-Warning ("Failed to inventory site collection '{0}': {1}" -f $site.Url, $_.Exception.Message)
                    $siteName = if ($site.PSObject.Properties.Name -contains 'Title') { $site.Title } else { $site.Url }
                    Stop-InventoryAfterError -Scope 'SiteCollection' -Url $site.Url -Name $siteName -Message $_.Exception.Message
                }
            }
        }

        'Site' {
            try {
                Export-SiteInventory `
                    -Url $SiteUrl `
                    -CsvPath $TempOutputPath `
                    -IncludeHidden:$IncludeHiddenLibraries `
                    -IncludeSystem:$IncludeSystemLibraries `
                    -BatchSize $PageSize
            }
            catch {
                if ($script:InventoryStopRequested) {
                    throw
                }
                Write-Warning ("Failed to inventory site collection '{0}': {1}" -f $SiteUrl, $_.Exception.Message)
                Stop-InventoryAfterError -Scope 'SiteCollection' -Url $SiteUrl -Name $SiteUrl -Message $_.Exception.Message
            }
        }

        'Web' {
            try {
                Export-WebInventory `
                    -Url $WebUrl `
                    -SiteCollectionUrl $WebUrl `
                    -CsvPath $TempOutputPath `
                    -Append:$script:CsvCreated `
                    -IncludeHidden:$IncludeHiddenLibraries `
                    -IncludeSystem:$IncludeSystemLibraries `
                    -BatchSize $PageSize
            }
            catch {
                if ($script:InventoryStopRequested) {
                    throw
                }
                Write-Warning ("Failed to inventory web '{0}': {1}" -f $WebUrl, $_.Exception.Message)
                Stop-InventoryAfterError -Scope 'Web' -Url $WebUrl -Name $WebUrl -Message $_.Exception.Message
            }
        }

        'WebUrlsFile' {
            $rootWebUrls = @(Get-WebUrlsFromFile -Path $WebUrlsFile | Select-Object -Unique)
            Write-Info -Color Green -Message ("Root web URLs loaded from file: {0}; descendant subsites are included." -f $rootWebUrls.Count)

            foreach ($rootWebUrl in $rootWebUrls) {
                $webUrls = @(Get-WebUrlsFromSite -Url $rootWebUrl | Select-Object -Unique)
                Write-Info -Color Green -Message ("Web URLs expanded from root '{0}': {1}" -f $rootWebUrl, $webUrls.Count)

                foreach ($targetWebUrl in $webUrls) {
                    try {
                        Export-WebInventory `
                            -Url $targetWebUrl `
                            -SiteCollectionUrl $rootWebUrl `
                            -CsvPath $TempOutputPath `
                            -Append:$script:CsvCreated `
                            -IncludeHidden:$IncludeHiddenLibraries `
                            -IncludeSystem:$IncludeSystemLibraries `
                            -BatchSize $PageSize
                    }
                    catch {
                        if ($script:InventoryStopRequested) {
                            throw
                        }
                        Write-Warning ("Failed to inventory web '{0}': {1}" -f $targetWebUrl, $_.Exception.Message)
                        Stop-InventoryAfterError -Scope 'Web' -Url $targetWebUrl -Name $targetWebUrl -Message $_.Exception.Message
                    }
                }
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
    Write-Info -Color Green -Message ("Inventory completed: {0}" -f $OutputPath)
    $script:InventoryStopwatch.Stop()
    Write-Info -Color Green -Message ("Scan duration: {0}" -f (Format-InventoryDuration -Elapsed $script:InventoryStopwatch.Elapsed))
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCdCr8lhXmJ17k0
# VlsQaQDY3oZq8AHSHDS/Gbtht3aJaqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIH8X++2oDZLQMeTyH3XPTq6ljQElQPH4sZ2JeKWdQ7hNMA0GCSqG
# SIb3DQEBAQUABIIBgI5bK0RP3wE156h3IfAYEcF6bZ2LToJ664yWrNLi8hYjbws0
# 81HNfTq3wTBOO+NHrVJPUuC/9Haq+ZANRq0s1sSW6PoJLoR772k3D8BtXD+kDxQS
# zP3r7oq/bKE/9nDSsxBBdJluO759ES+C3DIxYs85jVbwR7KmJoibcQOUfZeuGSQ4
# qJGWW73ojFEsYtnUKZPu4gsaoe4weVNlX2ir5BxgZqGZ/22l8jqybsil8Io2Zm2G
# qwznumLwVWpFJO90hBiIc5ouECimVA0/pPx36Ozg7Sn5RkujIg0PX3vQ6qcnyBdv
# wkVHWXX+Bysk4iBUsGR7iwasIE9t39Vx2Fs5e/q72Qo+PU5aHitJcOM5RH4zz2NE
# WddtYwnhimIviGEU7DtVH7CHSYbgvy4R+6DKad3kt5Z8Rnx82ZjGsKceMX31c35e
# 1HYQKXUn4adKTD+qqur1zIXr4mfLnNy1f2XjLaaB+XDSK8jMrcqwzEpl4xn5wSSs
# HMuRoGZj9PHoL1dClKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MTVaMC8GCSqGSIb3DQEJBDEiBCAJew1LUludNt16+s+ttnO0y6S/6OYbCxD3HgU1
# UUjnTzANBgkqhkiG9w0BAQEFAASCAgCgIA7ZGW7iOar1THg0GGy/DnKLJ21PjXmB
# aF4Xepz1JrvI2KXEoBxgnpPqLt9pdeNjY3iRKhlnUmCoNhg7GOC/6f0OeiGOKb/j
# mZYjgVnTCPXvf7aMqUCCyD8OdwegILm9jCDohdEnX1ilwjhQGHJelB8b357ueGdB
# 1FkyThxZohw5d7d64Uc/z5co009D6ryyPFN1RA+FQ3ez1y/fYXXPLFucNeA1mDIF
# L8WQn8UqKQdtrNuXotCRDU0Jk5nf5DezSBZyMPMF7BCifKeeWgU+4Q07FW+7rfsU
# 5BfVdnvfQqWVP/sk65iTC5B/YD2eyczMAUtl1YZwhIuJpkGPLJFRVTX6oYTRSwGQ
# /lS3KQ9NmtcoADyLUad977Cxsohzw5Cjxqdd20mEnz+QykgV3nH7CeSWfTnEu8PH
# 97Ib7v5sOGadYV05SvcKv6undT3xtEosz2pMg6u8lkVQXMQsQEOljXk8S+8GI8VH
# tGtSIv9vpw89pXllvZk0mPk3eEKfKH0IUcrVSY+y154UiYzBhPzJX4rHJoE5gyDs
# 1qXE4KXE/JvZzf18agrDa/5kgrKrlIUDSaOaQ6u+eQuYTzonJohQEqD5Nb4+3fap
# pAITBFo4dNEiKiqmAEhQhiQdel/pyy4gyuuUGcshK2C32ic142BTjNu9p4kODwq7
# 4EEo1xc1DQ==
# SIG # End signature block
