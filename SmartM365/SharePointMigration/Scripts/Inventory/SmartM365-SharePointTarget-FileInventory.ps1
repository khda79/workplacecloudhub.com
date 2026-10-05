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
    1.0.10
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
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'CANCELLED'
try {

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

    return ("{0:00}:{1:00}:{2:00}" -f $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds)
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

    $lists = Invoke-SPORead -Label 'lists' -Operation { Get-PnPList -Includes BaseType,Hidden,Title,ItemCount,RootFolder,IsSystemList -Connection $Connection -ErrorAction Stop }
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
            Write-Info -Color DarkCyan -Message ("  Skipping library '{0}' ({1}): hidden; use -IncludeHiddenLibraries to include it." -f $list.Title, $list.RootFolder.ServerRelativeUrl)
            continue
        }

        if (-not $IncludeSystem -and (Test-SystemLibrary -List $list)) {
            $systemSkipped++
            Write-Info -Color DarkCyan -Message ("  Skipping library '{0}' ({1}): system; use -IncludeSystemLibraries to include it." -f $list.Title, $list.RootFolder.ServerRelativeUrl)
            continue
        }

        if ($list.ItemCount -eq 0) {
            $emptySkipped++
            Write-Info -Color DarkCyan -Message ("  Skipping library '{0}' ({1}): zero items." -f $list.Title, $list.RootFolder.ServerRelativeUrl)
            continue
        }

        $selectedDocumentLibraries++
        $list
    }

    if ($WriteSummary) {
        Write-Info -Color DarkCyan -Message ("  Document libraries visible: {0}; selected for inventory: {1}; skipped hidden: {2}; skipped system: {3}; skipped empty: {4}" -f $totalDocumentLibraries, $selectedDocumentLibraries, $hiddenSkipped, $systemSkipped, $emptySkipped)
    }
}

function Invoke-FileInventoryItems {
    param(
        $Connection,
        $Library,
        [int]$LastItemId,
        [int]$BatchSize,
        [switch]$SimplePaging
    )

    if ($SimplePaging) {
        Get-PnPListItem `
            -List $Library `
            -PageSize $BatchSize `
            -Fields 'ID', 'FSObjType', 'FileLeafRef', 'FileRef', 'UniqueId', 'File_x0020_Size', 'Created', 'Author', 'Modified', 'Editor', 'ContentType', '_UIVersionString', 'CheckoutUser' `
            -IncludeContentType `
            -Connection $Connection `
            -ErrorAction Stop
        return
    }

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
        <Value Type='Counter'>$LastItemId</Value>
      </Gt>
    </Where>
    <OrderBy Override='TRUE'>
      <FieldRef Name='ID' Ascending='TRUE' />
    </OrderBy>
  </Query>
  <RowLimit Paged='TRUE'>$BatchSize</RowLimit>
</View>
"@
    Get-PnPListItem -List $Library -Query $query -PageSize $BatchSize -IncludeContentType -Connection $Connection -ErrorAction Stop
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
    $effectiveBatchSize = $BatchSize
    $timeoutsWithoutProgress = 0

    while ($true) {
        $startItemId = $lastItemId
        $itemsThisAttempt = 0
        try {
            Invoke-FileInventoryItems -Connection $Connection -Library $Library `
                -LastItemId $startItemId -BatchSize $effectiveBatchSize -SimplePaging:$useSimplePagingFallback |
                ForEach-Object {
                    $item = $_
                    if ($item.Id -le $lastItemId) { return }
                    $lastItemId = $item.Id
                    $itemsThisAttempt++

            $values = $item.FieldValues
            $itemsScanned++

            if ($values.ContainsKey('FSObjType') -and [string]$values['FSObjType'] -ne '0') {
                return
            }

            $fileName = [string]$values['FileLeafRef']
            $serverRelativeUrl = [string]$values['FileRef']
            $fileSizeBytes = $null

            if ([string]::IsNullOrWhiteSpace($serverRelativeUrl)) {
                return
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
            if (-not $useSimplePagingFallback -and $itemsThisAttempt -eq 0 -and $startItemId -eq 0 -and $Library.ItemCount -gt 0) {
                Write-Warning ("    ID CAML paging returned 0 items for non-empty library '{0}'. Retrying with simple PnP paging." -f $Library.Title)
                $useSimplePagingFallback = $true
                continue
            }
            break
        }
        catch {
            if ($useSimplePagingFallback) { throw }
            $messages = [System.Collections.Generic.List[string]]::new()
            $exception = $_.Exception
            while ($exception) {
                $messages.Add($exception.Message)
                $exception = $exception.InnerException
            }
            if (($messages -join ' | ') -notmatch 'HttpClient\.Timeout') { throw }
            $timeoutsWithoutProgress = if ($lastItemId -gt $startItemId) { 1 } else { $timeoutsWithoutProgress + 1 }
            if ($timeoutsWithoutProgress -ge 3) {
                Write-Warning ("    PnP request for library '{0}' after item ID {1} timed out on 3 attempts without further progress." -f $Library.Title, $lastItemId)
                throw
            }
            $effectiveBatchSize = [Math]::Min($effectiveBatchSize, $(if ($timeoutsWithoutProgress -eq 1) { 100 } else { 25 }))
            $delay = if ($timeoutsWithoutProgress -eq 1) { 5 } else { 15 }
            Write-Warning ("    PnP request for library '{0}' timed out after item ID {1} (attempt {2}/3 without further progress). Retrying in {3} s with page size {4}." -f $Library.Title, $lastItemId, $timeoutsWithoutProgress, $delay, $effectiveBatchSize)
            Start-Sleep -Seconds $delay
        }
    }

    Write-Info -Color DarkGray -Message ("    Items scanned: {0}; files exported: {1}" -f $itemsScanned, $filesReturned)
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

function Add-FileInventoryMetric {
    param($Row)
    $script:MetricRows++
    $path = ([string]$Row.ServerRelativeUrl).Trim().Replace('\','/').TrimEnd('/')
    if (-not $path) { $script:MetricMissingPathRows++; return }
    $library = ([string]$Row.LibraryUrl).Trim().Replace('\','/').TrimEnd('/')
    [long]$size = 0
    $hasSize = [long]::TryParse([string]$Row.SizeBytes, [ref]$size) -and $size -ge 0
    if ($script:MetricFileSizes.ContainsKey($path)) {
        $script:MetricDuplicateRows++
        $previous = $script:MetricFileSizes[$path]
        if ($null -ne $previous -and $hasSize -and $previous -ne $size) { $script:MetricConflictingSizeRows++ }
        if ($null -eq $previous -and $hasSize) {
            $script:MetricFileSizes[$path] = $size
            $script:MetricKnownSizeBytes += $size
        }
    }
    else {
        $script:MetricFileSizes[$path] = if ($hasSize) { $size } else { $null }
        if ($hasSize) { $script:MetricKnownSizeBytes += $size }
    }
    if (-not $library -or -not $path.StartsWith($library + '/', [StringComparison]::OrdinalIgnoreCase)) {
        $script:MetricInvalidLibraryRows++
        return
    }
    $parent = $path.Substring(0, $path.LastIndexOf('/'))
    while ($parent.StartsWith($library + '/', [StringComparison]::OrdinalIgnoreCase)) {
        [void]$script:MetricFolders.Add($parent)
        $slash = $parent.LastIndexOf('/')
        if ($slash -lt 0) { break }
        $parent = $parent.Substring(0, $slash)
    }
}

function Write-FileInventoryMetrics {
    param([string]$CsvPath)
    $csv = Get-Item -LiteralPath $CsvPath -ErrorAction Stop
    $missingSizes = @($script:MetricFileSizes.Values | Where-Object { $null -eq $_ }).Count
    $payload = [ordered]@{
        SchemaVersion = 1
        InventoryFile = $csv.Name
        CsvLengthBytes = [long]$csv.Length
        CsvSha256 = (Get-FileHash -LiteralPath $CsvPath -Algorithm SHA256 -ErrorAction Stop).Hash
        CsvLastWriteTimeUtcTicks = [long]$csv.LastWriteTimeUtc.Ticks
        CompletedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Rows = [long]$script:MetricRows
        Files = [long]$script:MetricFileSizes.Count
        FoldersWithFiles = [long]$script:MetricFolders.Count
        KnownSizeBytes = [long]$script:MetricKnownSizeBytes
        MissingSizeFiles = [long]$missingSizes
        MissingPathRows = [long]$script:MetricMissingPathRows
        InvalidLibraryRows = [long]$script:MetricInvalidLibraryRows
        DuplicateRows = [long]$script:MetricDuplicateRows
        ConflictingSizeRows = [long]$script:MetricConflictingSizeRows
    }
    $metricsPath = "$CsvPath.metrics.json.txt"
    $tempPath = "$metricsPath.tmp"
    $payload | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $tempPath -Encoding UTF8
    Move-Item -LiteralPath $tempPath -Destination $metricsPath -Force
    Write-Info -Color DarkCyan -Message ("File inventory metrics: {0} files; {1} folders with files; {2} known bytes; details: {3}" -f $payload.Files, $payload.FoldersWithFiles, $payload.KnownSizeBytes, $metricsPath)
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
    $web = Invoke-SPORead -Label 'web' -Operation { Get-PnPWeb -Includes Title,Url,ServerRelativeUrl -Connection $connection -ErrorAction Stop }
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
            foreach ($row in $rows) { Add-FileInventoryMetric -Row $row }
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

function Write-SPOReadRetry { param([string]$Message) Write-Warning $Message }

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
            $connection = Connect-SPOInventory -Url $currentWebUrl
            $web = Invoke-SPORead -Label 'web' -Operation { Get-PnPWeb -Includes Title,Url,ServerRelativeUrl -Connection $connection -ErrorAction Stop }
            $webUrls.Add($web.Url)

            try {
                $subWebs = @(Invoke-SPORead -Label 'subsites' -Operation { Get-PnPSubWeb -Includes Title,Url,ServerRelativeUrl -Connection $connection -ErrorAction Stop })
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
                $web = Invoke-SPORead -Label 'web' -Operation { Get-PnPWeb -Includes Title,Url -Connection $connection -ErrorAction Stop }
                $targetName = $web.Title
            }
            catch {
                $targetName = Get-NameFromUrl -Url $SiteUrl
            }
        }

        'Web' {
            try {
                $connection = Connect-SPOInventory -Url $WebUrl
                $web = Invoke-SPORead -Label 'web' -Operation { Get-PnPWeb -Includes Title,Url -Connection $connection -ErrorAction Stop }
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
$script:MetricFileSizes = @{}
$script:MetricFolders = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$script:MetricRows = [long]0
$script:MetricKnownSizeBytes = [long]0
$script:MetricMissingPathRows = [long]0
$script:MetricInvalidLibraryRows = [long]0
$script:MetricDuplicateRows = [long]0
$script:MetricConflictingSizeRows = [long]0
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
    Write-FileInventoryMetrics -CsvPath $OutputPath
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
    $script:ConsoleLifecycleStatus = 'FAILED'
    exit $script:InventoryExitCode
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDg2DT9P3wnyutc
# NwSYWfyp94ym6L/4g4wlRjTMWICUYqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCvj8pvO45XcGjda4aIeiWm
# 3mi0PvH53dLsI47+dB+7UjANBgkqhkiG9w0BAQEFAASCAYAetvv7/MYx5wFOR114
# g8/5D0BUgfcY1Il6Xbe0jp7vi6rXHlXoEbDFjeNOb3Wb30qNbt9lpoWaQ9mi+cE3
# HZ7547XAd2Mk4K/QOGQKP8qSx/L+SWVrjkCEyakOo6Vqq7P3xnIBVFZYvDxLAiQi
# o2j6cQbmJL4qWU/TqNpLNQEUFq5CmMWvMscBKUhI3K4OLkAq48QT4eeA96+AuP2I
# 9QfVyF2M3jsm6xxqiNaFcfzyuaRSIwjalKXcpmZondE1eUk+dvNfCL3Ojz6w/zlq
# TK22MDTiKWZ7qsnnk2hByrBqlQqXCPwki8c11erPJNn9qqPrGY4wZQpVMqLWRKuq
# mSLNBPPKpGnvLV31cvxJFZVw7q12Fp2Zy/3mGF8XV7jtjgGfZZeNyKAbAMdSKvzL
# HynltQB0DU1NlvvS3RBAm0PH37ioRSbWo8B7j0Xl9Q2XspdVodVUf2pOr87zVzDT
# g4/rUrxNvo9Y8mRiHD95FQMmWqYlys+BxUPfrq0qJb6+lyM=
# SIG # End signature block
