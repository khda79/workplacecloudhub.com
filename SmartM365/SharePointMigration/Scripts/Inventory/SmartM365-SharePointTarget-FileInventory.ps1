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
            $webUrls = @(Get-WebUrlsFromFile -Path $WebUrlsFile | Select-Object -Unique)
            Write-Info -Color Green -Message ("Web URLs loaded from file: {0}" -f $webUrls.Count)

            foreach ($targetWebUrl in $webUrls) {
                try {
                    Export-WebInventory `
                        -Url $targetWebUrl `
                        -SiteCollectionUrl $targetWebUrl `
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
