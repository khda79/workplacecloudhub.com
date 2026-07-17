Set-StrictMode -Version 2.0

$script:SemrVersion = '1.1.1'
$script:OnPremisesSession = $null
$script:InventoryContext = $null
$script:ConnectionState = [ordered]@{
    ActiveDirectory = $false
    OnPremisesExchange = $false
    ExchangeOnline = $false
    MicrosoftGraph = $false
    EntraConnect = $false
}

function Get-SemrVersion {
    return $script:SemrVersion
}

function ConvertTo-SemrHashtable {
    param([Parameter(Mandatory)]$InputObject)

    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $result[$key] = ConvertTo-SemrHashtable -InputObject $InputObject[$key]
        }
        return $result
    }

    if ($InputObject -is [pscustomobject]) {
        $result = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $result[$property.Name] = ConvertTo-SemrHashtable -InputObject $property.Value
        }
        return $result
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return @($InputObject | ForEach-Object { ConvertTo-SemrHashtable -InputObject $_ })
    }

    return $InputObject
}

function Merge-SemrConfigDefault {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Runtime,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Template,
        [string]$Prefix = ''
    )

    $added = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $Template.Keys) {
        $path = if ($Prefix) { "$Prefix.$key" } else { [string]$key }
        if (-not $Runtime.Contains($key)) {
            $Runtime[$key] = $Template[$key]
            [void]$added.Add($path)
            continue
        }

        if ($Runtime[$key] -is [System.Collections.IDictionary] -and $Template[$key] -is [System.Collections.IDictionary]) {
            foreach ($child in @(Merge-SemrConfigDefault -Runtime $Runtime[$key] -Template $Template[$key] -Prefix $path)) {
                [void]$added.Add($child)
            }
        }
    }
    return @($added)
}

function Resolve-SemrConfigPath {
    param(
        [string]$Path,
        [Parameter(Mandatory)][string]$BasePath
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

function Get-SemrConfig {
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path $PSScriptRoot 'Config\SmartM365-ExchangeMigrationReadiness.local.json')
    )

    $templatePath = "$Path.template"
    if (-not (Test-Path -LiteralPath $templatePath -PathType Leaf)) {
        throw "Configuration template not found: $templatePath"
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $parent = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $parent)) {
            New-Item -Path $parent -ItemType Directory -Force | Out-Null
        }
        Copy-Item -LiteralPath $templatePath -Destination $Path
    }

    $template = ConvertTo-SemrHashtable -InputObject (Get-Content -LiteralPath $templatePath -Raw | ConvertFrom-Json)
    $runtime = ConvertTo-SemrHashtable -InputObject (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json)
    $added = @(Merge-SemrConfigDefault -Runtime $runtime -Template $template)
    $configurationChanged = $added.Count -gt 0
    foreach ($deprecatedKey in @('Tenant', 'OnPremises', 'EntraConnect', 'SmartM365', 'EvidenceSources')) {
        if ($runtime.Contains($deprecatedKey)) {
            $runtime.Remove($deprecatedKey)
            $configurationChanged = $true
        }
    }
    if ($runtime.Contains('MicrosoftGraph') -and $runtime['MicrosoftGraph'] -is [System.Collections.IDictionary] -and $runtime['MicrosoftGraph'].Contains('UseDeviceCode')) {
        $runtime['MicrosoftGraph'].Remove('UseDeviceCode')
        $configurationChanged = $true
    }
    if ($configurationChanged) {
        $runtime | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $Path -Encoding utf8
    }

    $mode = [string]$runtime['Mode']
    if ($mode -notin @('Live', 'CacheOnly')) {
        throw "Invalid Mode '$mode'. Expected Live or CacheOnly."
    }
    $tenantProfile = $runtime['TenantProfile']
    if ([string]::IsNullOrWhiteSpace([string]$tenantProfile['TenantId'])) {
        throw 'TenantProfile.TenantId is required to keep authentication in the intended tenant.'
    }
    if ([string]::IsNullOrWhiteSpace([string]$runtime['Hybrid']['TargetDeliveryDomain'])) {
        $runtime['Hybrid']['TargetDeliveryDomain'] = [string]$tenantProfile['RemoteRoutingDomain']
    }
    $cacheRootPath = Resolve-SemrConfigPath -Path ([string]$runtime['Cache']['RootPath']) -BasePath $PSScriptRoot

    $runtime['_RuntimePath'] = $Path
    $runtime['_AddedKeys'] = $added
    $runtime['_TenantProfileKey'] = [string]$tenantProfile['ProfileKey']
    $runtime['_TenantId'] = [string]$tenantProfile['TenantId']
    $runtime['_RemoteRoutingDomain'] = [string]$tenantProfile['RemoteRoutingDomain']
    $runtime['_CacheRootPath'] = $cacheRootPath
    $runtime['_InventoryDataLastPath'] = if ($cacheRootPath) { Join-Path $cacheRootPath 'DATA-LAST' } else { '' }
    return $runtime
}
function Get-SemrConnectionState {
    [CmdletBinding()]
    param()

    $copy = [ordered]@{}
    foreach ($key in $script:ConnectionState.Keys) {
        $copy[$key] = [bool]$script:ConnectionState[$key]
    }
    return [pscustomobject]$copy
}

function Test-SemrCommand {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Connect-SemrActiveDirectory {
    [CmdletBinding()]
    param()

    Import-Module ActiveDirectory -ErrorAction Stop
    [void](Get-ADDomain -ErrorAction Stop)
    $script:ConnectionState.ActiveDirectory = $true
    return Get-SemrConnectionState
}

function Connect-SemrOnPremisesExchange {
    [CmdletBinding()]
    param(
        [string]$ConnectionUri = '',
        [ValidateSet('Kerberos', 'Negotiate')]
        [string]$Authentication = 'Kerberos',
        [pscredential]$Credential
    )

    if (Test-SemrCommand -Name 'Get-OnPremMailbox') {
        $script:ConnectionState.OnPremisesExchange = $true
        return Get-SemrConnectionState
    }

    if ((Test-SemrCommand -Name 'Get-Mailbox') -and (Test-SemrCommand -Name 'Get-RemoteMailbox')) {
        $script:ConnectionState.OnPremisesExchange = $true
        return Get-SemrConnectionState
    }

    if ([string]::IsNullOrWhiteSpace($ConnectionUri)) {
        throw 'Exchange on-premises cmdlets are unavailable. The application will use the configured CSV fallback.'
    }

    $sessionParameters = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri = $ConnectionUri
        Authentication = $Authentication
        ErrorAction = 'Stop'
    }
    if ($Credential) {
        $sessionParameters.Credential = $Credential
    }

    $script:OnPremisesSession = New-PSSession @sessionParameters
    Import-PSSession -Session $script:OnPremisesSession -Prefix OnPrem -DisableNameChecking -AllowClobber -Global | Out-Null
    if (-not (Test-SemrCommand -Name 'Get-OnPremMailbox')) {
        throw 'The Exchange on-premises session was imported, but Get-OnPremMailbox is unavailable.'
    }

    $script:ConnectionState.OnPremisesExchange = $true
    return Get-SemrConnectionState
}

function Connect-SemrExchangeOnline {
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName = '',
        [bool]$DisableWam = $true,
        [string]$TenantId = ''
    )

    Import-Module ExchangeOnlineManagement -MinimumVersion 3.0.0 -ErrorAction Stop
    if (Test-SemrCommand -Name 'Disconnect-ExchangeOnline') {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }

    $parameters = @{
        ShowBanner = $false
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        $parameters.UserPrincipalName = $UserPrincipalName
    }
    if ($DisableWam) {
        $parameters.DisableWAM = $true
    }
    Connect-ExchangeOnline @parameters
    if ($TenantId -and (Test-SemrCommand -Name 'Get-ConnectionInformation')) {
        $connection = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1)
        $connectedTenantId = if ($connection.Count -eq 1) { [string](Get-SemrPropertyValue -InputObject $connection[0] -Names @('TenantID', 'TenantId') -Default '') } else { '' }
        if ($connectedTenantId -and $connectedTenantId -ne $TenantId) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            throw "Exchange Online connected to tenant '$connectedTenantId' instead of configured tenant '$TenantId'."
        }
    }
    $script:ConnectionState.ExchangeOnline = $true
    return Get-SemrConnectionState
}

function Connect-SemrMicrosoftGraph {
    [CmdletBinding()]
    param(
        [string[]]$Scopes = @('User.Read.All', 'Directory.Read.All', 'Organization.Read.All'),
        [string]$TenantId = ''
    )

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    if (Test-SemrCommand -Name 'Disconnect-MgGraph') {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }

    $parameters = @{
        Scopes = $Scopes
        ContextScope = 'Process'
        NoWelcome = $true
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $parameters.TenantId = $TenantId
    }
    Connect-MgGraph @parameters | Out-Null
    $context = Get-MgContext
    if (-not $context) {
        throw 'Microsoft Graph authentication did not return a context.'
    }
    if ($TenantId -and [string]$context.TenantId -ne $TenantId) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        throw "Microsoft Graph connected to tenant '$($context.TenantId)' instead of configured tenant '$TenantId'."
    }
    $script:ConnectionState.MicrosoftGraph = $true
    return Get-SemrConnectionState
}
function Disconnect-SemrSession {
    [CmdletBinding()]
    param()

    if (Test-SemrCommand -Name 'Disconnect-ExchangeOnline') {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    if (Test-SemrCommand -Name 'Disconnect-MgGraph') {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    if ($script:OnPremisesSession) {
        Remove-PSSession -Session $script:OnPremisesSession -ErrorAction SilentlyContinue
        $script:OnPremisesSession = $null
    }
    foreach ($key in $script:ConnectionState.Keys) {
        $script:ConnectionState[$key] = $false
    }
}

function Initialize-SemrLiveSourceConnections {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Config)

    $result = [ordered]@{
        Mode = [string]$Config['Mode']
        ActiveDirectoryLive = $false
        ActiveDirectoryMessage = ''
        ExchangeOnPremisesLive = $false
        ExchangeOnPremisesMessage = ''
    }
    if ([string]$Config['Mode'] -eq 'CacheOnly') {
        $result.ActiveDirectoryMessage = 'CacheOnly: live Active Directory connection was not attempted.'
        $result.ExchangeOnPremisesMessage = 'CacheOnly: live Exchange on-premises connection was not attempted.'
        return [pscustomobject]$result
    }

    if (-not $script:ConnectionState.ActiveDirectory) {
        try {
            Connect-SemrActiveDirectory | Out-Null
            $result.ActiveDirectoryMessage = 'Live Active Directory connection succeeded.'
        }
        catch {
            $result.ActiveDirectoryMessage = "Live Active Directory unavailable; CSV fallback selected. $($_.Exception.Message)"
        }
    }
    else {
        $result.ActiveDirectoryMessage = 'Live Active Directory was already connected.'
    }
    $result.ActiveDirectoryLive = [bool]$script:ConnectionState.ActiveDirectory

    if (-not $script:ConnectionState.OnPremisesExchange) {
        try {
            Connect-SemrOnPremisesExchange | Out-Null
            $result.ExchangeOnPremisesMessage = 'Live Exchange on-premises connection succeeded.'
        }
        catch {
            $result.ExchangeOnPremisesMessage = "Live Exchange on-premises unavailable; CSV fallback selected. $($_.Exception.Message)"
        }
    }
    else {
        $result.ExchangeOnPremisesMessage = 'Live Exchange on-premises was already connected.'
    }
    $result.ExchangeOnPremisesLive = [bool]$script:ConnectionState.OnPremisesExchange
    return [pscustomobject]$result
}

function Read-SemrTextFile {
    param([Parameter(Mandatory)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $strictUtf8 = [System.Text.UTF8Encoding]::new($false, $true)
    try {
        return [pscustomobject]@{
            Text = $strictUtf8.GetString($bytes)
            Encoding = 'UTF-8'
        }
    }
    catch {
        [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
        return [pscustomobject]@{
            Text = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)
            Encoding = 'Windows-1252'
        }
    }
}

function Get-SemrDelimiter {
    param([Parameter(Mandatory)][string]$Header)

    $candidates = @(',', ';', "`t")
    $best = ','
    $bestCount = -1
    foreach ($candidate in $candidates) {
        $count = ([regex]::Matches($Header, [regex]::Escape($candidate))).Count
        if ($count -gt $bestCount) {
            $best = $candidate
            $bestCount = $count
        }
    }
    return $best
}

function Get-SemrPropertyValue {
    param(
        $InputObject,
        [Parameter(Mandatory)][string[]]$Names,
        $Default = $null
    )

    if ($null -eq $InputObject) { return $Default }
    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($property -and $null -ne $property.Value) {
            return $property.Value
        }
    }
    return $Default
}

function ConvertTo-SemrAddressList {
    param($Value)

    if ($null -eq $Value) { return @() }
    $values = if ($Value -is [string]) { $Value -split '[;,|]' } else { @($Value) }
    return @(
        $values |
            ForEach-Object { ([string]$_).Trim() -replace '^(?i)smtp:', '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Test-SemrSmtpAddress {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    try {
        $address = [System.Net.Mail.MailAddress]::new($Value)
        return $address.Address -ieq $Value.Trim()
    }
    catch {
        return $false
    }
}

function Import-SemrBatchCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "CSV file not found: $Path"
    }

    $file = Read-SemrTextFile -Path $Path
    $lines = @($file.Text -split "\r?\n" | Where-Object { $_ -ne '' })
    if ($lines.Count -eq 0) {
        throw 'The CSV file is empty.'
    }
    $delimiter = Get-SemrDelimiter -Header $lines[0]
    $rows = @($file.Text | ConvertFrom-Csv -Delimiter $delimiter)
    if ($rows.Count -eq 0) {
        throw 'The CSV contains no data rows.'
    }

    $identityAliases = @('EmailAddress', 'PrimarySmtp', 'PrimarySmtpAddress', 'PrimarySMTPAddress', 'UserPrincipalName', 'UPN', 'Mailbox')
    $identityColumn = $null
    foreach ($alias in $identityAliases) {
        if ($rows[0].PSObject.Properties.Name -contains $alias) {
            $identityColumn = $alias
            break
        }
    }
    if (-not $identityColumn) {
        throw "Missing mailbox identity column. Expected one of: $($identityAliases -join ', ')."
    }

    $headers = @($rows[0].PSObject.Properties.Name)
    $allowedHeaders = @($identityAliases + @('MailboxType','TargetSku','TargetSkuPartNumber','SkuPartNumber','BadItemLimit','LargeItemLimit'))
    $unknownHeaders = @($headers | Where-Object { $_ -notin $allowedHeaders })
    $normalized = [System.Collections.Generic.List[object]]::new()
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $email = ([string](Get-SemrPropertyValue -InputObject $row -Names @($identityColumn) -Default '')).Trim()
        [void]$normalized.Add([pscustomobject]@{
            RowNumber = $rowNumber
            EmailAddress = $email
            MailboxType = ([string](Get-SemrPropertyValue -InputObject $row -Names @('MailboxType') -Default '')).Trim()
            TargetSku = ([string](Get-SemrPropertyValue -InputObject $row -Names @('TargetSku', 'TargetSkuPartNumber', 'SkuPartNumber') -Default '')).Trim()
            BadItemLimit = ([string](Get-SemrPropertyValue -InputObject $row -Names @('BadItemLimit') -Default '')).Trim()
            LargeItemLimit = ([string](Get-SemrPropertyValue -InputObject $row -Names @('LargeItemLimit') -Default '')).Trim()
            RawRow = $row
        })
    }

    return [pscustomobject]@{
        Path = (Resolve-Path -LiteralPath $Path).Path
        Encoding = $file.Encoding
        Delimiter = if ($delimiter -eq "`t") { 'TAB' } else { $delimiter }
        IdentityColumn = $identityColumn
        Headers = $headers
        UnknownHeaders = $unknownHeaders
        Rows = @($normalized)
    }
}

function ConvertTo-SemrFinding {
    param(
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$EmailAddress,
        [Parameter(Mandatory)][string]$CheckId,
        [Parameter(Mandatory)][string]$Category,
        [ValidateSet('Critical', 'Warning', 'Information')][string]$Severity,
        [ValidateSet('PASS', 'FAIL', 'WARN', 'UNKNOWN', 'SKIPPED')][string]$Result,
        [bool]$IsBlocking,
        [string]$ObservedValue = '',
        [string]$ExpectedValue = '',
        [string]$EvidenceSource = '',
        [string]$Message = '',
        [string]$RecommendedAction = ''
    )

    return [pscustomobject][ordered]@{
        RunId = $RunId
        EmailAddress = $EmailAddress
        CheckId = $CheckId
        Category = $Category
        Severity = $Severity
        Result = $Result
        IsBlocking = $IsBlocking
        ObservedValue = $ObservedValue
        ExpectedValue = $ExpectedValue
        EvidenceSource = $EvidenceSource
        SourceTimestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Message = $Message
        RecommendedAction = $RecommendedAction
    }
}

function Add-SemrFinding {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$List,
        [Parameter(Mandatory)][hashtable]$Parameters
    )
    [void]$List.Add((ConvertTo-SemrFinding @Parameters))
}

function Get-SemrOnPremCommandName {
    param([Parameter(Mandatory)][string]$Name)
    $prefixed = $Name -replace '^Get-', 'Get-OnPrem'
    if (Test-SemrCommand -Name $prefixed) { return $prefixed }
    if (Test-SemrCommand -Name $Name) { return $Name }
    return ''
}

function Invoke-SemrCommandSafe {
    param(
        [string]$CommandName,
        [hashtable]$Parameters = @{}
    )
    if ([string]::IsNullOrWhiteSpace($CommandName)) { return @() }
    try {
        return @(& $CommandName @Parameters -ErrorAction Stop)
    }
    catch {
        return @()
    }
}

function ConvertTo-SemrBoolean {
    param($Value)

    if ($Value -is [bool]) { return [bool]$Value }
    $text = ([string]$Value).Trim()
    return $text -match '^(?i:true|1|yes|oui)$'
}

function Split-SemrInventoryValue {
    param($Value)

    if ($null -eq $Value) { return @() }
    return @(
        ([string]$Value -split '[;|\r\n]+' | ForEach-Object { $_.Trim() }) |
            Where-Object { $_ -and $_ -notmatch '^(?i:NotChecked|None|N/A)$' }
    )
}

function Get-SemrEvidenceSourceMode {
    param(
        [System.Collections.IDictionary]$Config,
        [Parameter(Mandatory)][string]$Name
    )

    $mode = if ($Config -and $Config.Contains('Mode')) { [string]$Config['Mode'] } else { 'Live' }
    if ($mode -eq 'CacheOnly') { return 'Cache' }
    switch ($Name) {
        'ActiveDirectory' { return $(if ($script:ConnectionState.ActiveDirectory) { 'Live' } else { 'CacheFallback' }) }
        'ExchangeOnPremises' { return $(if ($script:ConnectionState.OnPremisesExchange) { 'Live' } else { 'CacheFallback' }) }
        'EntraConnect' { return $(if ($script:ConnectionState.EntraConnect) { 'Live' } else { 'CacheFallback' }) }
        default { return 'Live' }
    }
}
function Get-SemrInventoryFileState {
    param(
        [Parameter(Mandatory)][string]$Path,
        [double]$MaximumAgeHours = 48
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ Available = $false; Path = $Path; Timestamp = $null; AgeHours = $null; Message = "Inventory file not found: $Path" }
    }
    $item = Get-Item -LiteralPath $Path
    $ageHours = ((Get-Date) - $item.LastWriteTime).TotalHours
    $fresh = $MaximumAgeHours -le 0 -or $ageHours -le $MaximumAgeHours
    return [pscustomobject]@{
        Available = $fresh
        Path = $item.FullName
        Timestamp = $item.LastWriteTime
        AgeHours = [math]::Round($ageHours, 2)
        Message = if ($fresh) { "Inventory available; age $([math]::Round($ageHours, 2)) hour(s)." } else { "Inventory is stale; age $([math]::Round($ageHours, 2)) hour(s), maximum $MaximumAgeHours." }
    }
}

function Import-SemrInventoryMatches {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$EmailAddresses,
        [Parameter(Mandatory)][string[]]$IdentityColumns
    )

    $targets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $result = @{}
    foreach ($emailAddress in $EmailAddresses) {
        $normalized = ([string]$emailAddress).Trim().ToLowerInvariant()
        if (-not $normalized) { continue }
        [void]$targets.Add($normalized)
        $result[$normalized] = [System.Collections.Generic.List[object]]::new()
    }

    Import-Csv -LiteralPath $Path | ForEach-Object {
        $row = $_
        $matchedTargets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($column in $IdentityColumns) {
            $value = Get-SemrPropertyValue -InputObject $row -Names @($column) -Default ''
            foreach ($candidate in @(([string]$value -split '[;|\r\n]+'))) {
                $normalizedCandidate = ($candidate.Trim() -replace '^(?i:smtp:)', '').ToLowerInvariant()
                if ($targets.Contains($normalizedCandidate)) {
                    [void]$matchedTargets.Add($normalizedCandidate)
                }
            }
        }
        foreach ($matchedTarget in $matchedTargets) {
            [void]$result[$matchedTarget].Add($row)
        }
    }
    return $result
}

function Initialize-SemrInventoryContext {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Config,
        [Parameter(Mandatory)][object[]]$BatchRows
    )

    $script:InventoryContext = [ordered]@{
        Mode = [string]$Config['Mode']
        DataLastPath = [string]$Config['_InventoryDataLastPath']
        ActiveDirectoryAvailable = $false
        ExchangeOnPremisesAvailable = $false
        EntraConnectAvailable = $false
        ActiveDirectoryMessage = 'Tenant CSV cache mode is not enabled.'
        ExchangeOnPremisesMessage = 'Tenant CSV cache mode is not enabled.'
        EntraConnectMessage = 'Tenant CSV cache mode is not enabled.'
        AdRowsByEmail = @{}
        MailboxRowsByEmail = @{}
        AdTimestamp = $null
        MailboxTimestamp = $null
        EntraTimestamp = $null
        EntraRows = @()
        CloudUsersAvailable = $false
        ExchangeOnlineAvailable = $false
        MigrationDataAvailable = $false
        LicenseDataAvailable = $false
        CloudUserRowsByEmail = @{}
        ExoMailboxRowsByEmail = @{}
        MigrationRowsByEmail = @{}
        LicenseRows = @()
        CloudTimestamp = $null
        ExoTimestamp = $null
        MigrationTimestamp = $null
        LicenseTimestamp = $null
    }

    # Always preload these inventories: they are mandatory in CacheOnly and automatic fallbacks in Live.
    $usesAdInventory = $true
    $usesExchangeInventory = $true
    $usesEntraInventory = $true

    $dataLastPath = [string]$Config['_InventoryDataLastPath']
    $cacheRootPath = [string]$Config['_CacheRootPath']
    if (-not (Test-Path -LiteralPath $dataLastPath -PathType Container) -and (Test-Path -LiteralPath $cacheRootPath -PathType Container)) {
        $dataLastPath = $cacheRootPath
        $script:InventoryContext.DataLastPath = $dataLastPath
    }
    if ([string]::IsNullOrWhiteSpace($dataLastPath) -or -not (Test-Path -LiteralPath $dataLastPath -PathType Container)) {
        $message = "Tenant CSV cache folder not found: $dataLastPath"
        $script:InventoryContext.ActiveDirectoryMessage = $message
        $script:InventoryContext.ExchangeOnPremisesMessage = $message
        $script:InventoryContext.EntraConnectMessage = $message
        return [pscustomobject]$script:InventoryContext
    }

    $maximumAgeHours = [double]$Config['Cache']['MaximumAgeHours']
    $emails = @($BatchRows | ForEach-Object { [string]$_.EmailAddress } | Where-Object { $_ })
    $adPath = Join-Path $dataLastPath 'AD_Users_AllDomains.csv'
    $mailboxPath = Join-Path $dataLastPath 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
    $entraPath = Join-Path $dataLastPath 'M365_Entra_AzureADConnect_SyncHealth.csv'
    $adState = Get-SemrInventoryFileState -Path $adPath -MaximumAgeHours $maximumAgeHours
    $mailboxState = Get-SemrInventoryFileState -Path $mailboxPath -MaximumAgeHours $maximumAgeHours
    $entraState = Get-SemrInventoryFileState -Path $entraPath -MaximumAgeHours $maximumAgeHours

    if ($usesAdInventory -and $adState.Available) {
        $script:InventoryContext.AdRowsByEmail = Import-SemrInventoryMatches -Path $adPath -EmailAddresses $emails -IdentityColumns @('EmailAddress', 'UserPrincipalName', 'PrimarySmtpAddress', 'ProxyAddresses')
        $script:InventoryContext.ActiveDirectoryAvailable = $true
        $script:InventoryContext.ActiveDirectoryMessage = $adState.Message
        $script:InventoryContext.AdTimestamp = $adState.Timestamp
    }
    elseif ($usesAdInventory) {
        $script:InventoryContext.ActiveDirectoryMessage = $adState.Message
    }

    if ($usesExchangeInventory -and $adState.Available -and $mailboxState.Available) {
        if (-not $script:InventoryContext.ActiveDirectoryAvailable) {
            $script:InventoryContext.AdRowsByEmail = Import-SemrInventoryMatches -Path $adPath -EmailAddresses $emails -IdentityColumns @('EmailAddress', 'UserPrincipalName', 'PrimarySmtpAddress', 'ProxyAddresses')
            $script:InventoryContext.AdTimestamp = $adState.Timestamp
        }
        $script:InventoryContext.MailboxRowsByEmail = Import-SemrInventoryMatches -Path $mailboxPath -EmailAddresses $emails -IdentityColumns @('PrimarySMTPaddress', 'UserPrincipalName', 'EmailAddresses')
        $script:InventoryContext.ExchangeOnPremisesAvailable = $true
        $script:InventoryContext.ExchangeOnPremisesMessage = "$($adState.Message) $($mailboxState.Message)"
        $script:InventoryContext.MailboxTimestamp = $mailboxState.Timestamp
    }
    elseif ($usesExchangeInventory) {
        $script:InventoryContext.ExchangeOnPremisesMessage = "AD inventory: $($adState.Message) Exchange inventory: $($mailboxState.Message)"
    }

    if ($usesEntraInventory -and $entraState.Available) {
        $script:InventoryContext.EntraRows = @(Import-Csv -LiteralPath $entraPath)
        $script:InventoryContext.EntraConnectAvailable = $script:InventoryContext.EntraRows.Count -gt 0
        $script:InventoryContext.EntraConnectMessage = $entraState.Message
        $script:InventoryContext.EntraTimestamp = $entraState.Timestamp
    }
    elseif ($usesEntraInventory) {
        $script:InventoryContext.EntraConnectMessage = $entraState.Message
    }

    if ([string]$Config['Mode'] -eq 'CacheOnly') {
        $cloudUserPath = Join-Path $dataLastPath 'M365_Users_Active.csv'
        $exoMailboxPath = Join-Path $dataLastPath 'Exchange_EXO_Mailboxes_AllDomains.csv'
        $migrationPath = Join-Path $dataLastPath 'Exchange_EXO_MigrationJobs.csv'
        $licensePath = Join-Path $dataLastPath 'M365_Licenses_Tenant.csv'
        $cloudState = Get-SemrInventoryFileState -Path $cloudUserPath -MaximumAgeHours $maximumAgeHours
        $exoState = Get-SemrInventoryFileState -Path $exoMailboxPath -MaximumAgeHours $maximumAgeHours
        $migrationState = Get-SemrInventoryFileState -Path $migrationPath -MaximumAgeHours $maximumAgeHours
        $licenseState = Get-SemrInventoryFileState -Path $licensePath -MaximumAgeHours $maximumAgeHours

        if ($cloudState.Available) {
            $script:InventoryContext.CloudUserRowsByEmail = Import-SemrInventoryMatches -Path $cloudUserPath -EmailAddresses $emails -IdentityColumns @('User principal name', 'Proxy addresses')
            $script:InventoryContext.CloudUsersAvailable = $true
            $script:InventoryContext.CloudTimestamp = $cloudState.Timestamp
        }
        if ($exoState.Available) {
            $script:InventoryContext.ExoMailboxRowsByEmail = Import-SemrInventoryMatches -Path $exoMailboxPath -EmailAddresses $emails -IdentityColumns @('PrimarySmtpAddress', 'UserPrincipalName', 'EmailAddresses')
            $script:InventoryContext.ExchangeOnlineAvailable = $true
            $script:InventoryContext.ExoTimestamp = $exoState.Timestamp
        }
        if ($migrationState.Available) {
            $script:InventoryContext.MigrationRowsByEmail = Import-SemrInventoryMatches -Path $migrationPath -EmailAddresses $emails -IdentityColumns @('EmailAddress', 'MigrationUser')
            $script:InventoryContext.MigrationDataAvailable = $true
            $script:InventoryContext.MigrationTimestamp = $migrationState.Timestamp
        }
        if ($licenseState.Available) {
            $script:InventoryContext.LicenseRows = @(Import-Csv -LiteralPath $licensePath)
            $script:InventoryContext.LicenseDataAvailable = $script:InventoryContext.LicenseRows.Count -gt 0
            $script:InventoryContext.LicenseTimestamp = $licenseState.Timestamp
        }
    }

    return [pscustomobject]$script:InventoryContext
}

function Get-SemrInventoryOnPremisesEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    $key = $EmailAddress.Trim().ToLowerInvariant()
    $adRows = @()
    if ($script:InventoryContext.AdRowsByEmail.ContainsKey($key)) { $adRows = @($script:InventoryContext.AdRowsByEmail[$key] | ForEach-Object { $_ }) }
    $mailboxRows = @()
    if ($script:InventoryContext.MailboxRowsByEmail.ContainsKey($key)) { $mailboxRows = @($script:InventoryContext.MailboxRowsByEmail[$key] | ForEach-Object { $_ }) }
    $adRow = @($adRows | Select-Object -First 1)
    $mailboxes = [System.Collections.Generic.List[object]]::new()
    $remoteMailboxes = [System.Collections.Generic.List[object]]::new()
    $statistics = [System.Collections.Generic.List[object]]::new()
    $permissions = [System.Collections.Generic.List[object]]::new()
    $permissionsAvailable = $mailboxRows.Count -gt 0

    foreach ($mailboxRow in $mailboxRows) {
        $relatedAdRow = if ($adRow.Count -eq 1) { $adRow[0] } else { $null }
        $proxyAddresses = if ($relatedAdRow) { @(Split-SemrInventoryValue -Value $relatedAdRow.ProxyAddresses) } else { @() }
        $primarySmtp = [string](Get-SemrPropertyValue -InputObject $mailboxRow -Names @('PrimarySMTPaddress', 'UserPrincipalName') -Default $EmailAddress)
        [void]$mailboxes.Add([pscustomobject]@{
            Identity = [string](Get-SemrPropertyValue -InputObject $mailboxRow -Names @('Identity', 'DistinguishedName') -Default $primarySmtp)
            DistinguishedName = [string](Get-SemrPropertyValue -InputObject $mailboxRow -Names @('DistinguishedName') -Default '')
            PrimarySmtpAddress = $primarySmtp
            RecipientTypeDetails = [string](Get-SemrPropertyValue -InputObject $mailboxRow -Names @('RecipientType') -Default 'UserMailbox')
            EmailAddresses = $proxyAddresses
            ExternalEmailAddress = if ($relatedAdRow) { [string]$relatedAdRow.TargetAddress } else { '' }
            WindowsEmailAddress = if ($relatedAdRow) { [string]$relatedAdRow.TargetAddress } else { '' }
            GrantSendOnBehalfTo = @(Split-SemrInventoryValue -Value $mailboxRow.GrantSendOnBehalfTo)
            LitigationHoldEnabled = $null
            InPlaceHolds = @()
        })
        $sizeMb = [string](Get-SemrPropertyValue -InputObject $mailboxRow -Names @('TotalItemSize-In-MB') -Default '')
        if ($sizeMb) {
            [void]$statistics.Add([pscustomobject]@{ TotalItemSize = "$sizeMb MB" })
        }

        foreach ($delegate in @(Split-SemrInventoryValue -Value $mailboxRow.FullAccess)) {
            [void]$permissions.Add([pscustomobject]@{ PermissionType = 'FullAccess'; Delegate = $delegate; IsInherited = $false; Source = 'TenantCache-ExchangeOnPrem' })
        }
        $sendAsRaw = [string]$mailboxRow.SendAs
        if ($sendAsRaw -match '^(?i:NotChecked)$') {
            $permissionsAvailable = $false
        }
        else {
            foreach ($delegate in @(Split-SemrInventoryValue -Value $sendAsRaw)) {
                [void]$permissions.Add([pscustomobject]@{ PermissionType = 'SendAs'; Delegate = $delegate; IsInherited = $false; Source = 'TenantCache-ExchangeOnPrem' })
            }
        }
        foreach ($delegate in @(Split-SemrInventoryValue -Value $mailboxRow.GrantSendOnBehalfTo)) {
            [void]$permissions.Add([pscustomobject]@{ PermissionType = 'SendOnBehalf'; Delegate = $delegate; IsInherited = $false; Source = 'TenantCache-ExchangeOnPrem' })
        }
    }

    if ($mailboxRows.Count -eq 0 -and $adRow.Count -eq 1 -and (ConvertTo-SemrBoolean -Value $adRow[0].HasMailbox)) {
        $proxyAddresses = @(Split-SemrInventoryValue -Value $adRow[0].ProxyAddresses)
        [void]$mailboxes.Add([pscustomobject]@{
            Identity = [string]$adRow[0].DistinguishedName
            DistinguishedName = [string]$adRow[0].DistinguishedName
            PrimarySmtpAddress = [string]$adRow[0].PrimarySmtpAddress
            RecipientTypeDetails = [string]$adRow[0].MailboxRecipientType
            EmailAddresses = $proxyAddresses
            ExternalEmailAddress = [string]$adRow[0].TargetAddress
            WindowsEmailAddress = [string]$adRow[0].TargetAddress
            GrantSendOnBehalfTo = @()
            LitigationHoldEnabled = $null
            InPlaceHolds = @()
        })
        if ([string]$adRow[0].MailboxSizeGB) {
            [void]$statistics.Add([pscustomobject]@{ TotalItemSize = "$($adRow[0].MailboxSizeGB) GB" })
        }
        $permissionsAvailable = $false
    }
    if ($adRow.Count -eq 1 -and (ConvertTo-SemrBoolean -Value $adRow[0].HasRemoteMailbox)) {
        [void]$remoteMailboxes.Add([pscustomobject]@{
            PrimarySmtpAddress = [string]$adRow[0].PrimarySmtpAddress
            RecipientTypeDetails = 'RemoteUserMailbox'
            RemoteRoutingAddress = [string]$adRow[0].RemoteMailboxRoutingAddress
        })
    }

    return [pscustomobject]@{
        Available = [bool]$script:InventoryContext.ExchangeOnPremisesAvailable
        Source = 'Tenant CSV cache inventories'
        SourceTimestamp = $script:InventoryContext.MailboxTimestamp
        Message = [string]$script:InventoryContext.ExchangeOnPremisesMessage
        Mailboxes = @($mailboxes)
        RemoteMailboxes = @($remoteMailboxes)
        MailUsers = @()
        Recipients = @($mailboxes) + @($remoteMailboxes)
        Statistics = @($statistics)
        Permissions = @($permissions)
        PermissionsAvailable = $permissionsAvailable
        HoldDataAvailable = $false
    }
}
function Get-SemrOnPremisesEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    if ($script:InventoryContext -and ($script:InventoryContext.Mode -eq 'CacheOnly' -or -not $script:ConnectionState.OnPremisesExchange)) {
        return Get-SemrInventoryOnPremisesEvidence -EmailAddress $EmailAddress
    }

    $mailboxCommand = Get-SemrOnPremCommandName -Name 'Get-Mailbox'
    $remoteMailboxCommand = Get-SemrOnPremCommandName -Name 'Get-RemoteMailbox'
    $mailUserCommand = Get-SemrOnPremCommandName -Name 'Get-MailUser'
    $recipientCommand = Get-SemrOnPremCommandName -Name 'Get-Recipient'
    $statisticsCommand = Get-SemrOnPremCommandName -Name 'Get-MailboxStatistics'
    $permissionCommand = Get-SemrOnPremCommandName -Name 'Get-MailboxPermission'
    $adPermissionCommand = Get-SemrOnPremCommandName -Name 'Get-ADPermission'

    $mailboxes = @(Invoke-SemrCommandSafe -CommandName $mailboxCommand -Parameters @{ Identity = $EmailAddress })
    $remoteMailboxes = @(Invoke-SemrCommandSafe -CommandName $remoteMailboxCommand -Parameters @{ Identity = $EmailAddress })
    $mailUsers = @(Invoke-SemrCommandSafe -CommandName $mailUserCommand -Parameters @{ Identity = $EmailAddress })
    $recipients = @(Invoke-SemrCommandSafe -CommandName $recipientCommand -Parameters @{ Identity = $EmailAddress })

    $statistics = @()
    if ($mailboxes.Count -eq 1 -and $statisticsCommand) {
        $statistics = @(Invoke-SemrCommandSafe -CommandName $statisticsCommand -Parameters @{ Identity = $mailboxes[0].Identity })
    }

    $permissions = [System.Collections.Generic.List[object]]::new()
    if ($mailboxes.Count -eq 1 -and $permissionCommand) {
        foreach ($permission in @(Invoke-SemrCommandSafe -CommandName $permissionCommand -Parameters @{ Identity = $mailboxes[0].Identity })) {
            $rights = @($permission.AccessRights | ForEach-Object { [string]$_ })
            if ($permission.IsInherited -or $rights -notcontains 'FullAccess') { continue }
            $delegate = [string]$permission.User
            if ($delegate -match 'NT AUTHORITY|S-1-5-|SELF') { continue }
            [void]$permissions.Add([pscustomobject]@{
                PermissionType = 'FullAccess'
                Delegate = $delegate
                IsInherited = [bool]$permission.IsInherited
                Source = 'ExchangeOnPrem'
            })
        }
    }
    if ($mailboxes.Count -eq 1 -and $adPermissionCommand) {
        foreach ($permission in @(Invoke-SemrCommandSafe -CommandName $adPermissionCommand -Parameters @{ Identity = $mailboxes[0].DistinguishedName })) {
            if ($permission.IsInherited -or @($permission.ExtendedRights) -notcontains 'Send-As' -or $permission.Deny) { continue }
            [void]$permissions.Add([pscustomobject]@{
                PermissionType = 'SendAs'
                Delegate = [string]$permission.User
                IsInherited = [bool]$permission.IsInherited
                Source = 'ExchangeOnPrem'
            })
        }
    }
    if ($mailboxes.Count -eq 1) {
        foreach ($delegate in @($mailboxes[0].GrantSendOnBehalfTo)) {
            [void]$permissions.Add([pscustomobject]@{
                PermissionType = 'SendOnBehalf'
                Delegate = [string]$delegate
                IsInherited = $false
                Source = 'ExchangeOnPrem'
            })
        }
    }

    return [pscustomobject]@{
        Available = [bool]$script:ConnectionState.OnPremisesExchange
        Source = 'Live Exchange on-premises'
        SourceTimestamp = Get-Date
        Message = 'Live Exchange on-premises evidence collected.'
        Mailboxes = $mailboxes
        RemoteMailboxes = $remoteMailboxes
        MailUsers = $mailUsers
        Recipients = $recipients
        Statistics = $statistics
        Permissions = @($permissions)
        PermissionsAvailable = $true
        HoldDataAvailable = $true
    }
}

function Get-SemrActiveDirectoryEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    if (($script:InventoryContext -and $script:InventoryContext.Mode -eq 'CacheOnly') -or -not $script:ConnectionState.ActiveDirectory -or -not (Test-SemrCommand -Name 'Get-ADUser')) {
        if ($script:InventoryContext) {
            $key = $EmailAddress.Trim().ToLowerInvariant()
            $users = @()
            if ($script:InventoryContext.AdRowsByEmail.ContainsKey($key)) { $users = @($script:InventoryContext.AdRowsByEmail[$key] | ForEach-Object { $_ }) }
            return [pscustomobject]@{
                Available = [bool]$script:InventoryContext.ActiveDirectoryAvailable
                Source = 'Tenant CSV cache AD inventory'
                SourceTimestamp = $script:InventoryContext.AdTimestamp
                Message = [string]$script:InventoryContext.ActiveDirectoryMessage
                Users = $users
            }
        }
        return [pscustomobject]@{ Available = $false; Source = 'Active Directory'; SourceTimestamp = $null; Message = 'Active Directory is not connected.'; Users = @() }
    }

    $escaped = $EmailAddress.Replace("'", "''")
    $users = @()
    try {
        $users = @(
            Get-ADUser -Filter "UserPrincipalName -eq '$escaped' -or mail -eq '$escaped' -or proxyAddresses -eq 'smtp:$escaped'" `
                -Properties Enabled, UserPrincipalName, mail, proxyAddresses, targetAddress, msDS-ConsistencyGuid, ObjectGuid, whenChanged `
                -ErrorAction Stop
        )
    }
    catch { $null = $_ }
    return [pscustomobject]@{
        Available = $true
        Source = 'Live Active Directory'
        SourceTimestamp = Get-Date
        Message = 'Live Active Directory evidence collected.'
        Users = $users
    }
}
function Get-SemrInventoryExchangeOnlineEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    $key = $EmailAddress.Trim().ToLowerInvariant()
    $cloudUsers = @()
    if ($script:InventoryContext.CloudUserRowsByEmail.ContainsKey($key)) { $cloudUsers = @($script:InventoryContext.CloudUserRowsByEmail[$key] | ForEach-Object { $_ }) }
    $mailboxes = @()
    if ($script:InventoryContext.ExoMailboxRowsByEmail.ContainsKey($key)) { $mailboxes = @($script:InventoryContext.ExoMailboxRowsByEmail[$key] | ForEach-Object { $_ }) }
    $migrationUsers = @()
    if ($script:InventoryContext.MigrationRowsByEmail.ContainsKey($key)) { $migrationUsers = @($script:InventoryContext.MigrationRowsByEmail[$key] | ForEach-Object { $_ }) }
    $recipients = [System.Collections.Generic.List[object]]::new()
    if ($mailboxes.Count -gt 0) {
        foreach ($mailbox in $mailboxes) { [void]$recipients.Add($mailbox) }
    }
    else {
        foreach ($user in $cloudUsers) {
            [void]$recipients.Add([pscustomobject]@{
                PrimarySmtpAddress = $EmailAddress
                RecipientType = 'MailUser'
                RecipientTypeDetails = 'MailUser'
                ExternalEmailAddress = $EmailAddress
                EmailAddresses = @(Split-SemrInventoryValue -Value (Get-SemrPropertyValue -InputObject $user -Names @('Proxy addresses') -Default ''))
            })
        }
    }
    $statistics = @(
        $mailboxes | ForEach-Object {
            [pscustomobject]@{
                TotalItemSize = "$([string](Get-SemrPropertyValue -InputObject $_ -Names @('TotalItemSizeGB') -Default '0')) GB"
                ItemCount = Get-SemrPropertyValue -InputObject $_ -Names @('ItemCount') -Default ''
            }
        }
    )
    return [pscustomobject]@{
        Available = [bool]($script:InventoryContext.CloudUsersAvailable -and $script:InventoryContext.ExchangeOnlineAvailable)
        Source = 'CacheOnly CSV inventories'
        SourceTimestamp = $script:InventoryContext.ExoTimestamp
        Recipients = @($recipients)
        Mailboxes = $mailboxes
        Statistics = $statistics
        SoftDeleted = @()
        SoftDeletedAvailable = $false
        MigrationUsers = $migrationUsers
        MoveRequests = @()
        MoveDataAvailable = [bool]$script:InventoryContext.MigrationDataAvailable
    }
}

function Get-SemrInventoryGraphEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    $key = $EmailAddress.Trim().ToLowerInvariant()
    $sourceRows = @()
    if ($script:InventoryContext.CloudUserRowsByEmail.ContainsKey($key)) { $sourceRows = @($script:InventoryContext.CloudUserRowsByEmail[$key] | ForEach-Object { $_ }) }
    $users = @(
        $sourceRows | ForEach-Object {
            [pscustomobject]@{
                Id = [string](Get-SemrPropertyValue -InputObject $_ -Names @('Object Id') -Default '')
                DisplayName = [string](Get-SemrPropertyValue -InputObject $_ -Names @('Display name') -Default '')
                UserPrincipalName = [string](Get-SemrPropertyValue -InputObject $_ -Names @('User principal name') -Default '')
                Mail = $EmailAddress
                ProxyAddresses = @(Split-SemrInventoryValue -Value (Get-SemrPropertyValue -InputObject $_ -Names @('Proxy addresses') -Default ''))
                AccountEnabled = ConvertTo-SemrBoolean -Value (Get-SemrPropertyValue -InputObject $_ -Names @('AccountEnabled') -Default $false)
                OnPremisesSyncEnabled = ConvertTo-SemrBoolean -Value (Get-SemrPropertyValue -InputObject $_ -Names @('OnPremisesSyncEnabled', 'DirSyncEnabled') -Default $false)
                OnPremisesImmutableId = [string](Get-SemrPropertyValue -InputObject $_ -Names @('OnPremisesImmutableId') -Default '')
                OnPremisesLastSyncDateTime = [string](Get-SemrPropertyValue -InputObject $_ -Names @('Last dirsync time') -Default '')
                UsageLocation = [string](Get-SemrPropertyValue -InputObject $_ -Names @('Usage location') -Default '')
                AssignedLicenses = @(Split-SemrInventoryValue -Value (Get-SemrPropertyValue -InputObject $_ -Names @('Licenses') -Default ''))
                AssignedPlans = @()
            }
        }
    )
    $licenseDetails = if ($users.Count -eq 1) { @($users[0].AssignedLicenses | ForEach-Object { [pscustomobject]@{ SkuPartNumber = $_ } }) } else { @() }
    return [pscustomobject]@{
        Available = [bool]$script:InventoryContext.CloudUsersAvailable
        Source = 'CacheOnly M365 user inventory'
        SourceTimestamp = $script:InventoryContext.CloudTimestamp
        Users = $users
        LicenseDetails = $licenseDetails
    }
}
function Get-SemrExchangeOnlineEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    if ($script:InventoryContext -and $script:InventoryContext.Mode -eq 'CacheOnly') {
        return Get-SemrInventoryExchangeOnlineEvidence -EmailAddress $EmailAddress
    }
    if (-not $script:ConnectionState.ExchangeOnline) {
        return [pscustomobject]@{
            Available = $false
            Recipients = @()
            Mailboxes = @()
            Statistics = @()
            SoftDeleted = @()
            MigrationUsers = @()
            MoveRequests = @()
        }
    }

    $recipients = if (Test-SemrCommand -Name 'Get-EXORecipient') {
        Invoke-SemrCommandSafe -CommandName 'Get-EXORecipient' -Parameters @{
            Identity = $EmailAddress
            Properties = @('RecipientType', 'RecipientTypeDetails', 'EmailAddresses', 'ExternalEmailAddress')
        }
    } else { @() }
    $mailboxes = if (Test-SemrCommand -Name 'Get-EXOMailbox') {
        Invoke-SemrCommandSafe -CommandName 'Get-EXOMailbox' -Parameters @{
            Identity = $EmailAddress
            Properties = @(
                'RecipientTypeDetails', 'EmailAddresses', 'ExchangeGuid', 'ArchiveGuid',
                'LitigationHoldEnabled', 'InPlaceHolds', 'DelayHoldApplied',
                'DelayReleaseHoldApplied', 'ProhibitSendReceiveQuota', 'GrantSendOnBehalfTo'
            )
        }
    } else { @() }
    $statistics = if ($mailboxes.Count -eq 1 -and (Test-SemrCommand -Name 'Get-EXOMailboxStatistics')) {
        Invoke-SemrCommandSafe -CommandName 'Get-EXOMailboxStatistics' -Parameters @{
            Identity = $EmailAddress
            Properties = @('TotalItemSize', 'ItemCount', 'TotalDeletedItemSize', 'StorageLimitStatus')
        }
    } else { @() }
    $softDeleted = @()
    if (Test-SemrCommand -Name 'Get-Mailbox') {
        try {
            $softDeleted = @(Get-Mailbox -Identity $EmailAddress -SoftDeletedMailbox -IncludeInactiveMailbox -ErrorAction Stop)
        }
        catch { $null = $_ }
    }
    $migrationUsers = if (Test-SemrCommand -Name 'Get-MigrationUser') {
        Invoke-SemrCommandSafe -CommandName 'Get-MigrationUser' -Parameters @{ Identity = $EmailAddress }
    } else { @() }
    $moveRequests = if (Test-SemrCommand -Name 'Get-MoveRequest') {
        Invoke-SemrCommandSafe -CommandName 'Get-MoveRequest' -Parameters @{ Identity = $EmailAddress }
    } else { @() }

    return [pscustomobject]@{
        Available = $true
        Source = 'Live Exchange Online'
        SourceTimestamp = Get-Date
        Recipients = $recipients
        Mailboxes = $mailboxes
        Statistics = $statistics
        SoftDeleted = $softDeleted
        SoftDeletedAvailable = $true
        MigrationUsers = $migrationUsers
        MoveRequests = $moveRequests
        MoveDataAvailable = $true
    }
}

function Get-SemrGraphEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    if ($script:InventoryContext -and $script:InventoryContext.Mode -eq 'CacheOnly') {
        return Get-SemrInventoryGraphEvidence -EmailAddress $EmailAddress
    }
    if (-not $script:ConnectionState.MicrosoftGraph) {
        return [pscustomobject]@{ Available = $false; Users = @(); LicenseDetails = @() }
    }

    $escaped = $EmailAddress.Replace("'", "''")
    $users = @()
    try {
        $users = @(Get-MgUser -Filter "userPrincipalName eq '$escaped' or mail eq '$escaped'" -Property @(
            'id', 'displayName', 'userPrincipalName', 'mail', 'proxyAddresses', 'accountEnabled',
            'onPremisesSyncEnabled', 'onPremisesImmutableId', 'onPremisesLastSyncDateTime',
            'usageLocation', 'assignedLicenses', 'assignedPlans'
        ) -All -ErrorAction Stop)
    }
    catch { $null = $_ }

    $licenseDetails = @()
    if ($users.Count -eq 1 -and (Test-SemrCommand -Name 'Get-MgUserLicenseDetail')) {
        try {
            $licenseDetails = @(Get-MgUserLicenseDetail -UserId $users[0].Id -ErrorAction Stop)
        }
        catch { $null = $_ }
    }
    return [pscustomobject]@{
        Available = $true
        Source = 'Live Microsoft Graph'
        SourceTimestamp = Get-Date
        Users = $users
        LicenseDetails = $licenseDetails
    }
}

function Get-SemrTenantLicenseEvidence {
    param(
        [string]$TargetSku,
        [System.Collections.IDictionary]$Config
    )

    $result = [ordered]@{
        Available = $false
        TargetSku = $TargetSku
        Found = $false
        Enabled = 0
        Consumed = 0
        AvailableUnits = 0
        Message = ''
    }
    if ($Config -and [string]$Config['Mode'] -eq 'CacheOnly') {
        if (-not $script:InventoryContext -or -not $script:InventoryContext.LicenseDataAvailable) {
            $result.Message = 'Cached tenant license inventory is unavailable.'
            return [pscustomobject]$result
        }
        $result.Available = $true
        $sku = @($script:InventoryContext.LicenseRows | Where-Object { [string]$_.TenantSkuPartNumber -ieq $TargetSku } | Select-Object -First 1)
        if ($sku.Count -eq 0) {
            $result.Message = "Target SKU '$TargetSku' was not found in the cached tenant license inventory."
            return [pscustomobject]$result
        }
        $result.Found = $true
        $result.Enabled = [int]$sku[0].TenantPrepaidEnabled
        $result.Consumed = [int]$sku[0].TenantConsumedUnits
        $result.AvailableUnits = [math]::Max(0, $result.Enabled - $result.Consumed)
        $result.Message = "Cached: Enabled=$($result.Enabled); Consumed=$($result.Consumed); Available=$($result.AvailableUnits)"
        return [pscustomobject]$result
    }

    if (-not $script:ConnectionState.MicrosoftGraph -or -not (Test-SemrCommand -Name 'Get-MgSubscribedSku')) {
        $result.Message = 'Microsoft Graph subscribed SKU data is unavailable.'
        return [pscustomobject]$result
    }
    $result.Available = $true
    try {
        $sku = @(Get-MgSubscribedSku -All -ErrorAction Stop | Where-Object { $_.SkuPartNumber -ieq $TargetSku } | Select-Object -First 1)
        if ($sku.Count -eq 0) {
            $result.Message = "Target SKU '$TargetSku' was not found in subscribed SKUs."
            return [pscustomobject]$result
        }
        $result.Found = $true
        $result.Enabled = [int]$sku[0].PrepaidUnits.Enabled
        $result.Consumed = [int]$sku[0].ConsumedUnits
        $result.AvailableUnits = [math]::Max(0, $result.Enabled - $result.Consumed)
        $result.Message = "Enabled=$($result.Enabled); Consumed=$($result.Consumed); Available=$($result.AvailableUnits)"
    }
    catch {
        $result.Message = $_.Exception.Message
    }
    return [pscustomobject]$result
}
function Get-SemrExchangeOnlinePermission {
    param([Parameter(Mandatory)][string]$EmailAddress)

    $permissions = [System.Collections.Generic.List[object]]::new()
    if (-not $script:ConnectionState.ExchangeOnline) { return @() }

    $fullAccessCommand = if (Test-SemrCommand -Name 'Get-EXOMailboxPermission') { 'Get-EXOMailboxPermission' } elseif (Test-SemrCommand -Name 'Get-MailboxPermission') { 'Get-MailboxPermission' } else { '' }
    if ($fullAccessCommand) {
        foreach ($permission in @(Invoke-SemrCommandSafe -CommandName $fullAccessCommand -Parameters @{ Identity = $EmailAddress })) {
            $rights = @($permission.AccessRights | ForEach-Object { [string]$_ })
            if ($permission.IsInherited -or $rights -notcontains 'FullAccess') { continue }
            $delegate = [string]$permission.User
            if ($delegate -match 'NT AUTHORITY|S-1-5-|SELF') { continue }
            [void]$permissions.Add([pscustomobject]@{ EmailAddress = $EmailAddress; PermissionType = 'FullAccess'; Delegate = $delegate; IsInherited = [bool]$permission.IsInherited; Source = 'ExchangeOnline' })
        }
    }
    if (Test-SemrCommand -Name 'Get-RecipientPermission') {
        foreach ($permission in @(Invoke-SemrCommandSafe -CommandName 'Get-RecipientPermission' -Parameters @{ Identity = $EmailAddress })) {
            if ($permission.IsInherited -or $permission.Deny -or @($permission.AccessRights) -notcontains 'SendAs') { continue }
            $delegate = [string]$permission.Trustee
            if ($delegate -match 'NT AUTHORITY|S-1-5-|SELF') { continue }
            [void]$permissions.Add([pscustomobject]@{ EmailAddress = $EmailAddress; PermissionType = 'SendAs'; Delegate = $delegate; IsInherited = [bool]$permission.IsInherited; Source = 'ExchangeOnline' })
        }
    }
    if (Test-SemrCommand -Name 'Get-EXOMailbox') {
        $mailbox = @(Invoke-SemrCommandSafe -CommandName 'Get-EXOMailbox' -Parameters @{ Identity = $EmailAddress; Properties = @('GrantSendOnBehalfTo') } | Select-Object -First 1)
        if ($mailbox.Count -eq 1) {
            foreach ($delegate in @($mailbox[0].GrantSendOnBehalfTo)) {
                [void]$permissions.Add([pscustomobject]@{ EmailAddress = $EmailAddress; PermissionType = 'SendOnBehalf'; Delegate = [string]$delegate; IsInherited = $false; Source = 'ExchangeOnline' })
            }
        }
    }
    return @($permissions)
}

function Get-SemrPostMigrationPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Batch,
        [scriptblock]$ProgressCallback
    )

    if (-not $script:ConnectionState.ExchangeOnline) {
        throw 'Exchange Online must be connected before collecting post-migration permissions.'
    }
    $rows = @($Batch.Rows)
    $results = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($row in $rows) {
        $index++
        if ($ProgressCallback) { & $ProgressCallback $index $rows.Count $row.EmailAddress 'Collecting post-migration permissions' }
        foreach ($permission in @(Get-SemrExchangeOnlinePermission -EmailAddress $row.EmailAddress)) {
            [void]$results.Add([pscustomobject][ordered]@{
                EmailAddress = $permission.EmailAddress
                PermissionType = $permission.PermissionType
                Delegate = $permission.Delegate
                IsInherited = $permission.IsInherited
                Source = $permission.Source
                CapturedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            })
        }
    }
    return @($results)
}
function ConvertFrom-SemrByteSize {
    param($Value)
    if ($null -eq $Value) { return 0.0 }
    if ($Value.PSObject.Properties['Value'] -and $Value.Value.PSObject.Properties['ToBytes']) {
        return [math]::Round(([double]$Value.Value.ToBytes() / 1GB), 3)
    }
    $text = [string]$Value
    if ($text -match '\((?<bytes>[\d,]+) bytes\)') {
        return [math]::Round(([double](($Matches.bytes -replace ',', '')) / 1GB), 3)
    }
    if ($text -match '(?<number>[\d.,]+)\s*(?<unit>TB|GB|MB|KB)') {
        $number = 0.0
        [void][double]::TryParse(($Matches.number -replace ',', '.'), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)
        switch ($Matches.unit) {
            'TB' { return [math]::Round($number * 1024, 3) }
            'GB' { return [math]::Round($number, 3) }
            'MB' { return [math]::Round($number / 1024, 3) }
            'KB' { return [math]::Round($number / 1MB, 3) }
        }
    }
    return 0.0
}

function Get-SemrTargetQuotaGb {
    param(
        [string]$TargetSku,
        [System.Collections.IDictionary]$Config,
        [string]$MailboxType
    )
    if ($MailboxType -match 'Shared' -and [string]::IsNullOrWhiteSpace($TargetSku)) {
        return 50.0
    }
    $map = $Config['TargetQuotaGbBySku']
    if ($map -is [System.Collections.IDictionary] -and $TargetSku) {
        foreach ($key in $map.Keys) {
            if ([string]$key -ieq $TargetSku) {
                return [double]$map[$key]
            }
        }
    }
    return [double]$Config['DefaultTargetQuotaGb']
}

function Test-SemrHybridReadiness {
    [CmdletBinding()]
    param([System.Collections.IDictionary]$Config)

    $mode = if ($Config.Contains('Mode')) { [string]$Config['Mode'] } else { 'Live' }
    $hybridConfig = if ($Config.Contains('Hybrid')) { $Config['Hybrid'] } else { $Config }
    $result = [ordered]@{
        Available = $false
        EndpointFound = $false
        EndpointName = ''
        ConnectivitySuccess = $false
        Source = if ($mode -eq 'CacheOnly') { 'CacheOnly' } else { 'Live Exchange Online' }
        Message = ''
    }
    if ($mode -eq 'CacheOnly') {
        $result.Message = 'CacheOnly mode cannot execute a live migration endpoint test.'
        return [pscustomobject]$result
    }
    if (-not $script:ConnectionState.ExchangeOnline) {
        $result.Message = 'Exchange Online is not connected.'
        return [pscustomobject]$result
    }
    if (-not (Test-SemrCommand -Name 'Get-MigrationEndpoint')) {
        $result.Message = 'Get-MigrationEndpoint is unavailable.'
        return [pscustomobject]$result
    }

    $result.Available = $true
    $configuredName = [string]$hybridConfig['MigrationEndpointName']
    $endpoints = @()
    try {
        $endpoints = if ($configuredName) {
            @(Get-MigrationEndpoint -Identity $configuredName -ErrorAction Stop)
        }
        else {
            @(Get-MigrationEndpoint -ErrorAction Stop | Where-Object { [string]$_.EndpointType -match 'ExchangeRemoteMove' })
        }
    }
    catch {
        $result.Message = $_.Exception.Message
        return [pscustomobject]$result
    }

    if ($endpoints.Count -eq 0) {
        $result.Message = 'No ExchangeRemoteMove migration endpoint was found.'
        return [pscustomobject]$result
    }
    $endpoint = $endpoints[0]
    $result.EndpointFound = $true
    $result.EndpointName = [string]$endpoint.Identity
    if (Test-SemrCommand -Name 'Test-MigrationServerAvailability') {
        try {
            $test = Test-MigrationServerAvailability -Endpoint $endpoint.Identity -ErrorAction Stop
            $result.ConnectivitySuccess = [bool](Get-SemrPropertyValue -InputObject $test -Names @('Result', 'Success') -Default $false)
            if (-not $result.ConnectivitySuccess -and [string]$test.Result -eq 'Success') {
                $result.ConnectivitySuccess = $true
            }
            $result.Message = [string](Get-SemrPropertyValue -InputObject $test -Names @('Message', 'ErrorDetail') -Default $test.Result)
        }
        catch {
            $result.Message = $_.Exception.Message
        }
    }
    else {
        $result.Message = 'Endpoint found; Test-MigrationServerAvailability is unavailable.'
    }
    return [pscustomobject]$result
}

function Test-SemrEntraConnect {
    [CmdletBinding()]
    param([System.Collections.IDictionary]$Config)

    $mode = if ($Config -and $Config.Contains('Mode')) { [string]$Config['Mode'] } else { 'Live' }
    $liveFailure = ''
    if ($mode -eq 'Live') {
        if (-not (Test-SemrCommand -Name 'Get-ADSyncScheduler')) {
            try {
                Import-Module ADSync -ErrorAction Stop
            }
            catch {
                if ($PSVersionTable.PSEdition -eq 'Core') {
                    try { Import-Module ADSync -UseWindowsPowerShell -ErrorAction Stop } catch { $liveFailure = $_.Exception.Message }
                }
                else { $liveFailure = $_.Exception.Message }
            }
        }

        if (Test-SemrCommand -Name 'Get-ADSyncScheduler') {
            try {
                $scheduler = Get-ADSyncScheduler -ErrorAction Stop
                $runStatus = if (Test-SemrCommand -Name 'Get-ADSyncConnectorRunStatus') {
                    @(Get-ADSyncConnectorRunStatus -ErrorAction Stop | Select-Object ConnectorName, RunState)
                }
                else { @() }
                $script:ConnectionState.EntraConnect = $true
                return [pscustomobject][ordered]@{
                    Available = $true
                    Server = $env:COMPUTERNAME
                    Source = 'Live Microsoft Entra Connect (local ADSync)'
                    SourceTimestamp = Get-Date
                    SyncCycleEnabled = [bool]$scheduler.SyncCycleEnabled
                    SchedulerSuspended = [bool]$scheduler.SchedulerSuspended
                    StagingModeEnabled = [bool]$scheduler.StagingModeEnabled
                    NextSyncCycleStartTimeInUTC = $scheduler.NextSyncCycleStartTimeInUTC
                    ConnectorRunStatus = ($runStatus | ConvertTo-Json -Compress)
                    Message = 'Live Microsoft Entra Connect scheduler state collected from local ADSync cmdlets.'
                }
            }
            catch { $liveFailure = $_.Exception.Message }
        }
        elseif (-not $liveFailure) {
            $liveFailure = 'ADSync cmdlets are not available on this computer.'
        }
    }

    $result = [ordered]@{
        Available = $false
        Server = 'CSV cache'
        Source = if ($mode -eq 'CacheOnly') { 'Tenant CSV cache' } else { 'Tenant CSV cache fallback' }
        SourceTimestamp = $null
        SyncCycleEnabled = $null
        SchedulerSuspended = $null
        StagingModeEnabled = $null
        NextSyncCycleStartTimeInUTC = $null
        ConnectorRunStatus = ''
        Message = ''
    }
    if (-not $script:InventoryContext -or -not $script:InventoryContext.EntraConnectAvailable) {
        $cacheMessage = if ($script:InventoryContext) { [string]$script:InventoryContext.EntraConnectMessage } else { 'Tenant CSV cache context is not initialized.' }
        $result.Message = if ($mode -eq 'Live' -and $liveFailure) { "Live Entra Connect unavailable ($liveFailure). CSV fallback unavailable: $cacheMessage" } else { $cacheMessage }
        return [pscustomobject]$result
    }

    $rows = @($script:InventoryContext.EntraRows)
    $syncEnabledRow = @($rows | Where-Object CheckName -EQ 'SyncEnabled' | Select-Object -First 1)
    $lastSyncAgeRow = @($rows | Where-Object CheckName -EQ 'LastSyncAge' | Select-Object -First 1)
    if ($syncEnabledRow.Count -ne 1 -or $lastSyncAgeRow.Count -ne 1) {
        $result.Message = 'Tenant Entra sync-health cache is missing SyncEnabled or LastSyncAge.'
        return [pscustomobject]$result
    }
    $healthOk = @($rows | Where-Object { [string]$_.Status -ne 'OK' }).Count -eq 0
    $result.Available = $true
    $result.SourceTimestamp = $script:InventoryContext.EntraTimestamp
    $result.SyncCycleEnabled = ConvertTo-SemrBoolean -Value $syncEnabledRow[0].Value
    $result.SchedulerSuspended = -not $healthOk
    $result.NextSyncCycleStartTimeInUTC = $syncEnabledRow[0].LastSyncDateTimeUtc
    $result.ConnectorRunStatus = ($rows | Select-Object CheckName, Status, Value, Detail | ConvertTo-Json -Compress)
    $cacheSummary = "CSV cache: SyncEnabled=$($result.SyncCycleEnabled); SyncHealthOK=$healthOk; $($lastSyncAgeRow[0].Detail)"
    $result.Message = if ($mode -eq 'Live' -and $liveFailure) { "Live Entra Connect unavailable ($liveFailure). $cacheSummary" } else { $cacheSummary }
    return [pscustomobject]$result
}
function Invoke-SemrAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Batch,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Config,
        [scriptblock]$ProgressCallback,
        [scriptblock]$CancellationCheck
    )

    $runId = "SEMR-{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    $findings = [System.Collections.Generic.List[object]]::new()
    $permissionRows = [System.Collections.Generic.List[object]]::new()
    $evidenceRows = [System.Collections.Generic.List[object]]::new()
    $rows = @($Batch.Rows)
    $duplicates = @(
        $rows |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.EmailAddress) } |
            Group-Object { $_.EmailAddress.ToLowerInvariant() } |
            Where-Object Count -gt 1 |
            ForEach-Object Name
    )

    if ($ProgressCallback) {
        & $ProgressCallback 0 $rows.Count '' 'Loading tenant CSV fallback inventories'
    }
    [void](Initialize-SemrInventoryContext -Config $Config -BatchRows $rows)
    if ($ProgressCallback -and [string]$Config['Mode'] -eq 'Live') {
        & $ProgressCallback 0 $rows.Count '' 'Trying live AD and Exchange on-premises sources'
    }
    $sourceInitialization = Initialize-SemrLiveSourceConnections -Config $Config
    $hybrid = Test-SemrHybridReadiness -Config $Config
    $entraConnect = Test-SemrEntraConnect -Config $Config
    $licenseCapacityCache = @{}
    $index = 0

    foreach ($row in $rows) {
        $index++
        if ($CancellationCheck -and (& $CancellationCheck)) { break }
        if ($ProgressCallback) {
            & $ProgressCallback $index $rows.Count $row.EmailAddress "Collecting evidence"
        }

        $email = [string]$row.EmailAddress
        $base = @{
            RunId = $runId
            EmailAddress = $email
        }

        if ([string]::IsNullOrWhiteSpace($email)) {
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'CSV-EMPTY-IDENTITY'; Category = 'CSV'; Severity = 'Critical'; Result = 'FAIL'; IsBlocking = $true
                Message = "CSV row $($row.RowNumber) has an empty mailbox identity."; RecommendedAction = 'Populate EmailAddress and reload the CSV.'
            })
            continue
        }

        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'CSV-SMTP-FORMAT'; Category = 'CSV'; Severity = if (Test-SemrSmtpAddress $email) { 'Information' } else { 'Critical' }
            Result = if (Test-SemrSmtpAddress $email) { 'PASS' } else { 'FAIL' }; IsBlocking = -not (Test-SemrSmtpAddress $email)
            ObservedValue = $email; ExpectedValue = 'Valid SMTP address'; EvidenceSource = 'Batch CSV'
            Message = if (Test-SemrSmtpAddress $email) { 'Mailbox identity is a valid SMTP address.' } else { 'Mailbox identity is not a valid SMTP address.' }
            RecommendedAction = if (Test-SemrSmtpAddress $email) { '' } else { 'Correct the EmailAddress value.' }
        })
        $isDuplicate = $duplicates -contains $email.ToLowerInvariant()
        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'CSV-DUPLICATE'; Category = 'CSV'; Severity = if ($isDuplicate) { 'Critical' } else { 'Information' }
            Result = if ($isDuplicate) { 'FAIL' } else { 'PASS' }; IsBlocking = $isDuplicate
            ObservedValue = $email; ExpectedValue = 'One row per mailbox'; EvidenceSource = 'Batch CSV'
            Message = if ($isDuplicate) { 'Mailbox identity occurs more than once in the batch.' } else { 'Mailbox identity is unique in the batch.' }
            RecommendedAction = if ($isDuplicate) { 'Remove duplicate rows from the batch.' } else { '' }
        })

        $unknownHeaderText = @($Batch.UnknownHeaders) -join ';'
        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'CSV-COLUMNS'; Category = 'CSV'; Severity = if ($unknownHeaderText) { 'Warning' } else { 'Information' }
            Result = if ($unknownHeaderText) { 'WARN' } else { 'PASS' }; IsBlocking = $false
            ObservedValue = if ($unknownHeaderText) { $unknownHeaderText } else { $Batch.Headers -join ';' }; ExpectedValue = 'Only recognized migration readiness columns'
            EvidenceSource = 'Batch CSV'; Message = if ($unknownHeaderText) { 'The CSV contains columns that are preserved but not interpreted by this release.' } else { 'All CSV columns are recognized.' }
            RecommendedAction = if ($unknownHeaderText) { 'Confirm that ignored columns are intentional and do not carry required migration semantics.' } else { '' }
        })
        $mailboxTypeAllowed = [string]::IsNullOrWhiteSpace($row.MailboxType) -or $row.MailboxType -in @('PrimaryOnly','ArchiveOnly','PrimaryAndArchive')
        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'CSV-MAILBOX-TYPE'; Category = 'CSV'; Severity = if ($mailboxTypeAllowed) { 'Information' } else { 'Critical' }
            Result = if ($mailboxTypeAllowed) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $mailboxTypeAllowed
            ObservedValue = $row.MailboxType; ExpectedValue = 'Blank, PrimaryOnly, ArchiveOnly, or PrimaryAndArchive'; EvidenceSource = 'Batch CSV'
            Message = if ($mailboxTypeAllowed) { 'MailboxType is blank or recognized.' } else { 'MailboxType is not recognized.' }
            RecommendedAction = if ($mailboxTypeAllowed) { '' } else { 'Correct the MailboxType value for the selected migration scenario.' }
        })
        foreach ($limitDefinition in @(@('BadItemLimit',$row.BadItemLimit),@('LargeItemLimit',$row.LargeItemLimit))) {
            $limitName = [string]$limitDefinition[0]
            $limitValue = [string]$limitDefinition[1]
            $limitValid = [string]::IsNullOrWhiteSpace($limitValue) -or $limitValue -match '^\d+$'
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = "CSV-$($limitName.ToUpperInvariant())"; Category = 'CSV'; Severity = if ($limitValid) { 'Information' } else { 'Critical' }
                Result = if ($limitValid) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $limitValid
                ObservedValue = $limitValue; ExpectedValue = 'Blank or non-negative integer'; EvidenceSource = 'Batch CSV'
                Message = if ($limitValid) { "$limitName is blank or valid." } else { "$limitName is invalid." }
                RecommendedAction = if ($limitValid) { '' } else { "Correct $limitName before creating the migration batch." }
            })
        }
        $ad = Get-SemrActiveDirectoryEvidence -EmailAddress $email
        $onPrem = Get-SemrOnPremisesEvidence -EmailAddress $email
        $exo = Get-SemrExchangeOnlineEvidence -EmailAddress $email
        $graph = Get-SemrGraphEvidence -EmailAddress $email
        $adSource = [string](Get-SemrPropertyValue -InputObject $ad -Names @('Source') -Default 'Active Directory')
        $onPremSource = [string](Get-SemrPropertyValue -InputObject $onPrem -Names @('Source') -Default 'Exchange on-premises')

        if (-not $ad.Available) {
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'AD-SOURCE'; Category = 'ActiveDirectory'; Severity = 'Critical'; Result = 'UNKNOWN'; IsBlocking = $true
                ObservedValue = [string]$ad.Message; EvidenceSource = $adSource
                Message = 'Active Directory cache evidence is unavailable.'; RecommendedAction = 'Refresh or restore the configured AD inventory CSV, then rerun.'
            })
        }
        else {
            $adCount = @($ad.Users).Count
            $adValid = $adCount -eq 1
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'AD-IDENTITY-UNIQUE'; Category = 'ActiveDirectory'; Severity = if ($adValid) { 'Information' } else { 'Critical' }
                Result = if ($adValid) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $adValid
                ObservedValue = "$adCount matching AD user(s)"; ExpectedValue = '1'; EvidenceSource = $adSource
                Message = if ($adValid) { 'Exactly one Active Directory user was resolved.' } else { 'The Active Directory identity is missing or ambiguous.' }
                RecommendedAction = if ($adValid) { '' } else { 'Resolve AD mail/UPN/proxy address identity uniqueness.' }
            })
            if ($adValid) {
                $adEnabled = ConvertTo-SemrBoolean -Value $ad.Users[0].Enabled
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'AD-ACCOUNT-ENABLED'; Category = 'ActiveDirectory'; Severity = if ($adEnabled) { 'Information' } else { 'Warning' }
                    Result = if ($adEnabled) { 'PASS' } else { 'WARN' }; IsBlocking = $false
                    ObservedValue = "Enabled=$adEnabled"; ExpectedValue = 'Enabled account unless disabled migration is intentional'; EvidenceSource = $adSource
                    Message = if ($adEnabled) { 'The Active Directory account is enabled.' } else { 'The Active Directory account is disabled.' }
                    RecommendedAction = if ($adEnabled) { '' } else { 'Confirm that migrating a disabled account is intentional and supported by the migration wave.' }
                })
            }
        }

        if (-not $onPrem.Available) {
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'ONPREM-SOURCE'; Category = 'ExchangeOnPrem'; Severity = 'Critical'; Result = 'UNKNOWN'; IsBlocking = $true
                ObservedValue = [string]$onPrem.Message; EvidenceSource = $onPremSource
                Message = 'Exchange on-premises cache evidence is unavailable.'; RecommendedAction = 'Refresh or restore the configured Exchange on-premises inventory CSV, then rerun.'
            })
        }
        else {
            $mailboxCount = @($onPrem.Mailboxes).Count
            $remoteCount = @($onPrem.RemoteMailboxes).Count
            $mailUserCount = @($onPrem.MailUsers).Count
            $validPreMigrationState = $mailboxCount -eq 1 -and $remoteCount -eq 0
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'ONPREM-RECIPIENT-STATE'; Category = 'ExchangeOnPrem'; Severity = if ($validPreMigrationState) { 'Information' } else { 'Critical' }
                Result = if ($validPreMigrationState) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $validPreMigrationState
                ObservedValue = "UserMailbox=$mailboxCount; RemoteMailbox=$remoteCount; MailUser=$mailUserCount"
                ExpectedValue = 'UserMailbox=1; RemoteMailbox=0'; EvidenceSource = $onPremSource
                Message = if ($validPreMigrationState) { 'The on-premises recipient state is consistent with a pre-onboarding mailbox.' } else { 'The on-premises recipient state is not consistent with a pre-onboarding mailbox.' }
                RecommendedAction = if ($validPreMigrationState) { '' } else { 'Review mailbox type and hybrid recipient provisioning before migration.' }
            })

            if ($mailboxCount -eq 1) {
                $mailbox = $onPrem.Mailboxes[0]
                $targetAddress = [string](Get-SemrPropertyValue -InputObject $mailbox -Names @('ExternalEmailAddress', 'WindowsEmailAddress') -Default '')
                $addresses = @(ConvertTo-SemrAddressList -Value (Get-SemrPropertyValue -InputObject $mailbox -Names @('EmailAddresses') -Default @()))
                $targetDomain = [string]$Config['Hybrid']['TargetDeliveryDomain']
                $targetPresent = -not [string]::IsNullOrWhiteSpace($targetAddress)
                $targetValid = -not $targetPresent
                if ($targetPresent) {
                    $targetValid = -not $targetDomain -or $targetAddress.EndsWith("@$targetDomain", [StringComparison]::OrdinalIgnoreCase)
                }
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'ONPREM-TARGET-ADDRESS'; Category = 'HybridIdentity'; Severity = if ($targetValid) { 'Information' } else { 'Critical' }
                    Result = if ($targetValid) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $targetValid
                    ObservedValue = $targetAddress; ExpectedValue = if ($targetDomain) { "Empty on source UserMailbox, or *@$targetDomain when populated" } else { 'Empty or valid configured routing address' }
                    EvidenceSource = $onPremSource; Message = if (-not $targetPresent) { 'No target address is stamped on the source UserMailbox; this is acceptable before onboarding.' } elseif ($targetValid) { 'The populated target routing address uses the configured delivery domain.' } else { 'The populated target routing address does not use the configured target delivery domain.' }
                    RecommendedAction = if ($targetValid) { '' } else { 'Correct targetAddress/ExternalEmailAddress using supported Exchange tools.' }
                })

                $primarySmtp = [string](Get-SemrPropertyValue -InputObject $mailbox -Names @('PrimarySmtpAddress') -Default $email)
                $primaryCount = @((Get-SemrPropertyValue -InputObject $mailbox -Names @('EmailAddresses') -Default @()) | Where-Object { [string]$_ -cmatch '^SMTP:' }).Count
                $proxyValid = $primaryCount -eq 1 -and ($addresses -contains $primarySmtp)
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'ONPREM-PROXY-INTEGRITY'; Category = 'HybridIdentity'; Severity = if ($proxyValid) { 'Information' } else { 'Critical' }
                    Result = if ($proxyValid) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $proxyValid
                    ObservedValue = "PrimarySMTP=$primarySmtp; PrimaryProxyCount=$primaryCount"; ExpectedValue = 'One primary SMTP present in proxyAddresses'
                    EvidenceSource = $onPremSource; Message = if ($proxyValid) { 'Primary SMTP and proxyAddresses are consistent.' } else { 'Primary SMTP and proxyAddresses are inconsistent.' }
                    RecommendedAction = if ($proxyValid) { '' } else { 'Ensure proxyAddresses contains exactly one uppercase SMTP primary address.' }
                })

                $statistics = @($onPrem.Statistics | Select-Object -First 1)
                $sizeGb = if ($statistics.Count -eq 1) { ConvertFrom-SemrByteSize -Value $statistics[0].TotalItemSize } else { 0.0 }
                $targetSku = if ($row.TargetSku) { $row.TargetSku } else { [string]$Config['DefaultTargetSku'] }
                $quotaGb = Get-SemrTargetQuotaGb -TargetSku $targetSku -Config $Config -MailboxType ([string]$mailbox.RecipientTypeDetails)
                $bufferPercent = [double]$Config['QuotaSafetyBufferPercent']
                $safeQuota = $quotaGb * (1.0 - ($bufferPercent / 100.0))
                $sizePass = $statistics.Count -eq 1 -and $sizeGb -lt $safeQuota
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'MAILBOX-TARGET-QUOTA'; Category = 'Mailbox'; Severity = if ($sizePass) { 'Information' } else { 'Critical' }
                    Result = if ($statistics.Count -eq 0) { 'UNKNOWN' } elseif ($sizePass) { 'PASS' } else { 'FAIL' }
                    IsBlocking = -not $sizePass; ObservedValue = "$sizeGb GB"; ExpectedValue = "< $([math]::Round($safeQuota, 2)) GB ($targetSku / $quotaGb GB less $bufferPercent% buffer)"
                    EvidenceSource = "$onPremSource mailbox statistics"; Message = if ($sizePass) { 'Mailbox size fits the target quota with the configured safety buffer.' } elseif ($statistics.Count -eq 0) { 'Mailbox statistics could not be collected.' } else { 'Mailbox size exceeds the safe target quota.' }
                    RecommendedAction = if ($sizePass) { '' } else { 'Reduce mailbox size, enable/archive content, or select a target SKU with sufficient quota.' }
                })

                $holdDataAvailable = [bool](Get-SemrPropertyValue -InputObject $onPrem -Names @('HoldDataAvailable') -Default $false)
                $litigationHold = ConvertTo-SemrBoolean -Value (Get-SemrPropertyValue -InputObject $mailbox -Names @('LitigationHoldEnabled') -Default $false)
                $inPlaceHolds = @(Get-SemrPropertyValue -InputObject $mailbox -Names @('InPlaceHolds') -Default @())
                $holdCount = $inPlaceHolds.Count + $(if ($litigationHold) { 1 } else { 0 })
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'MAILBOX-HOLDS'; Category = 'Compliance'; Severity = if (-not $holdDataAvailable -or $holdCount -gt 0) { 'Warning' } else { 'Information' }
                    Result = if (-not $holdDataAvailable) { 'UNKNOWN' } elseif ($holdCount -gt 0) { 'WARN' } else { 'PASS' }; IsBlocking = $false
                    ObservedValue = if ($holdDataAvailable) { "LitigationHold=$litigationHold; InPlaceHolds=$($inPlaceHolds.Count)" } else { 'Hold properties are not present in the on-premises cache.' }
                    ExpectedValue = 'Hold state documented'; EvidenceSource = $onPremSource
                    Message = if (-not $holdDataAvailable) { 'Hold state requires a live Exchange validation or a richer inventory export.' } elseif ($holdCount -gt 0) { 'Mailbox hold state was detected. Hybrid moves normally preserve supported holds and Recoverable Items.' } else { 'No Litigation Hold or In-Place Hold was returned.' }
                    RecommendedAction = if (-not $holdDataAvailable) { 'Validate hold state during the live EXO phase or refresh the Exchange inventory with hold properties.' } elseif ($holdCount -gt 0) { 'Confirm hold preservation requirements and monitor Recoverable Items quota; do not remove holds solely to satisfy this precheck.' } else { '' }
                })

                foreach ($permission in @($onPrem.Permissions)) {
                    [void]$permissionRows.Add([pscustomobject][ordered]@{
                        RunId = $runId
                        EmailAddress = $email
                        PermissionType = $permission.PermissionType
                        Delegate = $permission.Delegate
                        IsInherited = $permission.IsInherited
                        Source = $permission.Source
                        CapturedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    })
                }
                $permissionCount = @($onPrem.Permissions).Count
                $permissionsAvailable = [bool](Get-SemrPropertyValue -InputObject $onPrem -Names @('PermissionsAvailable') -Default $false)
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'PERMISSIONS-BASELINE'; Category = 'Permissions'; Severity = if (-not $permissionsAvailable -or $permissionCount -gt 0) { 'Warning' } else { 'Information' }
                    Result = if (-not $permissionsAvailable) { 'UNKNOWN' } elseif ($permissionCount -gt 0) { 'WARN' } else { 'PASS' }; IsBlocking = $false
                    ObservedValue = "$permissionCount explicit permission grant(s)"; ExpectedValue = 'Baseline captured'
                    EvidenceSource = $onPremSource
                    Message = if ($permissionsAvailable) { 'Explicit Full Access, Send As, and Send on Behalf permissions were captured.' } else { 'The cache does not contain a complete delegated-permission baseline.' }
                    RecommendedAction = if (-not $permissionsAvailable) { 'Refresh the on-premises permission inventory before final approval.' } elseif ($permissionCount -gt 0) { 'Migrate dependent delegates in a coordinated wave and compare permissions after migration.' } else { '' }
                })
            }
        }

        $exoSource = [string](Get-SemrPropertyValue -InputObject $exo -Names @('Source') -Default 'Exchange Online')
        $graphSource = [string](Get-SemrPropertyValue -InputObject $graph -Names @('Source') -Default 'Microsoft Graph')
        if (-not $exo.Available) {
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'EXO-SOURCE'; Category = 'ExchangeOnline'; Severity = 'Critical'; Result = 'UNKNOWN'; IsBlocking = $true
                EvidenceSource = $exoSource; Message = if ([string]$Config['Mode'] -eq 'CacheOnly') { 'The cached Exchange Online inventories are unavailable.' } else { 'Exchange Online is not connected.' }
                RecommendedAction = if ([string]$Config['Mode'] -eq 'CacheOnly') { 'Refresh or restore the tenant CSV cache, then run the assessment again.' } else { 'Connect Exchange Online and run the assessment again.' }
            })
        }
        else {
            $cloudRecipientCount = @($exo.Recipients).Count
            $cloudMailboxCount = @($exo.Mailboxes).Count
            $cloudStateValid = $cloudRecipientCount -eq 1 -and $cloudMailboxCount -eq 0
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'EXO-RECIPIENT-STATE'; Category = 'ExchangeOnline'; Severity = if ($cloudStateValid) { 'Information' } else { 'Critical' }
                Result = if ($cloudStateValid) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $cloudStateValid
                ObservedValue = "Recipients=$cloudRecipientCount; ActiveMailboxes=$cloudMailboxCount"; ExpectedValue = 'One synchronized MailUser and no active EXO mailbox'
                EvidenceSource = $exoSource; Message = if ($cloudStateValid) { 'The cloud recipient state is consistent with hybrid pre-onboarding.' } else { 'The cloud recipient is missing, ambiguous, or already has an active mailbox.' }
                RecommendedAction = if ($cloudStateValid) { '' } else { 'Resolve synchronization, duplicate recipient, or split-brain mailbox state before migration.' }
            })
            $softDeletedAvailable = [bool](Get-SemrPropertyValue -InputObject $exo -Names @('SoftDeletedAvailable') -Default $false)
            $softDeletedCount = @($exo.SoftDeleted).Count
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'EXO-SOFT-DELETED-CONFLICT'; Category = 'KnownErrors'; Severity = if (-not $softDeletedAvailable) { 'Warning' } elseif ($softDeletedCount -gt 0) { 'Critical' } else { 'Information' }
                Result = if (-not $softDeletedAvailable) { 'UNKNOWN' } elseif ($softDeletedCount -gt 0) { 'FAIL' } else { 'PASS' }; IsBlocking = $softDeletedAvailable -and $softDeletedCount -gt 0
                ObservedValue = if ($softDeletedAvailable) { "$softDeletedCount soft-deleted/inactive mailbox match(es)" } else { 'Soft-deleted and inactive mailbox data is not present in the cache.' }; ExpectedValue = '0'
                EvidenceSource = $exoSource; Message = if (-not $softDeletedAvailable) { 'This conflict check requires Live mode or an enriched cache export.' } elseif ($softDeletedCount -gt 0) { 'A soft-deleted or inactive mailbox may conflict with the migration identity.' } else { 'No soft-deleted or inactive mailbox conflict was returned.' }
                RecommendedAction = if (-not $softDeletedAvailable) { 'Run the final validation in Live mode before creating the migration batch.' } elseif ($softDeletedCount -gt 0) { 'Investigate the soft-deleted/inactive object and GUID ownership before creating the batch.' } else { '' }
            })
            $moveDataAvailable = [bool](Get-SemrPropertyValue -InputObject $exo -Names @('MoveDataAvailable') -Default $false)
            $activeMoveCount = @($exo.MigrationUsers).Count + @($exo.MoveRequests).Count
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'EXO-EXISTING-MOVE'; Category = 'Migration'; Severity = if (-not $moveDataAvailable) { 'Warning' } elseif ($activeMoveCount -gt 0) { 'Critical' } else { 'Information' }
                Result = if (-not $moveDataAvailable) { 'UNKNOWN' } elseif ($activeMoveCount -gt 0) { 'FAIL' } else { 'PASS' }; IsBlocking = $moveDataAvailable -and $activeMoveCount -gt 0
                ObservedValue = if ($moveDataAvailable) { "$activeMoveCount migration object(s)" } else { 'Migration job data is not present in the cache.' }; ExpectedValue = '0'
                EvidenceSource = $exoSource; Message = if (-not $moveDataAvailable) { 'Existing move state could not be evaluated from the cache.' } elseif ($activeMoveCount -gt 0) { 'An existing migration user or move request already exists.' } else { 'No existing migration user or move request was returned.' }
                RecommendedAction = if (-not $moveDataAvailable) { 'Refresh the migration-jobs inventory or run the final validation in Live mode.' } elseif ($activeMoveCount -gt 0) { 'Review and clean up the previous migration object only after confirming its state and history.' } else { '' }
            })
        }

        if (-not $graph.Available) {
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'GRAPH-SOURCE'; Category = 'MicrosoftGraph'; Severity = 'Critical'; Result = 'UNKNOWN'; IsBlocking = $true
                EvidenceSource = $graphSource; Message = if ([string]$Config['Mode'] -eq 'CacheOnly') { 'The cached Microsoft 365 user inventory is unavailable.' } else { 'Microsoft Graph is not connected.' }
                RecommendedAction = if ([string]$Config['Mode'] -eq 'CacheOnly') { 'Refresh or restore the tenant CSV cache, then run the assessment again.' } else { 'Connect Microsoft Graph and run the assessment again.' }
            })
        }
        else {
            $graphCount = @($graph.Users).Count
            $graphValid = $graphCount -eq 1
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'GRAPH-USER-STATE'; Category = 'MicrosoftGraph'; Severity = if ($graphValid) { 'Information' } else { 'Critical' }
                Result = if ($graphValid) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $graphValid
                ObservedValue = "$graphCount matching user(s)"; ExpectedValue = '1'
                EvidenceSource = $graphSource; Message = if ($graphValid) { 'Exactly one Microsoft Entra user was resolved.' } else { 'The Microsoft Entra user is missing or ambiguous.' }
                RecommendedAction = if ($graphValid) { '' } else { 'Resolve Microsoft Entra synchronization and identity uniqueness.' }
            })
            if ($graphValid) {
                $user = $graph.Users[0]
                $syncEnabled = [bool]$user.OnPremisesSyncEnabled
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'GRAPH-DIRSYNC'; Category = 'MicrosoftGraph'; Severity = if ($syncEnabled) { 'Information' } else { 'Critical' }
                    Result = if ($syncEnabled) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $syncEnabled
                    ObservedValue = "OnPremisesSyncEnabled=$syncEnabled; LastSync=$($user.OnPremisesLastSyncDateTime)"
                    ExpectedValue = 'Synchronized user'; EvidenceSource = $graphSource
                    Message = if ($syncEnabled) { 'The user is synchronized from on-premises.' } else { 'The user is not marked as synchronized from on-premises.' }
                    RecommendedAction = if ($syncEnabled) { '' } else { 'Confirm the object is in Microsoft Entra Connect scope and successfully exported.' }
                })
                $licenseCount = @($graph.LicenseDetails).Count
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'LICENSE-PRE-MIGRATION'; Category = 'Licensing'; Severity = if ($licenseCount -gt 0) { 'Warning' } else { 'Information' }
                    Result = if ($licenseCount -gt 0) { 'WARN' } else { 'PASS' }; IsBlocking = $false
                    ObservedValue = "$licenseCount assigned license(s)"; ExpectedValue = 'License available; Exchange license assigned after migration'
                    EvidenceSource = $graphSource; Message = if ($licenseCount -gt 0) { 'One or more licenses are already assigned before migration.' } else { 'No assigned license was returned before migration.' }
                    RecommendedAction = if ($licenseCount -gt 0) { 'Verify that an Exchange service plan has not provisioned an unintended cloud mailbox.' } else { 'Assign the approved Exchange Online license after migration completion and within the operational deadline.' }
                })
                $usageLocationMissing = [string]::IsNullOrWhiteSpace([string]$user.UsageLocation)
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'LICENSE-USAGE-LOCATION'; Category = 'Licensing'; Severity = if ($usageLocationMissing) { 'Warning' } else { 'Information' }
                    Result = if ($usageLocationMissing) { 'WARN' } else { 'PASS' }; IsBlocking = $false
                    ObservedValue = [string]$user.UsageLocation; ExpectedValue = 'Two-letter usage location'
                    EvidenceSource = $graphSource; Message = if ($usageLocationMissing) { 'UsageLocation is empty and can block license assignment.' } else { 'UsageLocation is populated.' }
                    RecommendedAction = if ($usageLocationMissing) { 'Populate UsageLocation before the post-migration license assignment.' } else { '' }
                })

                $targetSku = if ($row.TargetSku) { $row.TargetSku } else { [string]$Config['DefaultTargetSku'] }
                if (-not $licenseCapacityCache.ContainsKey($targetSku.ToUpperInvariant())) {
                    $licenseCapacityCache[$targetSku.ToUpperInvariant()] = Get-SemrTenantLicenseEvidence -TargetSku $targetSku -Config $Config
                }
                $licenseCapacity = $licenseCapacityCache[$targetSku.ToUpperInvariant()]
                $capacityPass = $licenseCapacity.Found -and $licenseCapacity.AvailableUnits -gt 0
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'LICENSE-CAPACITY'; Category = 'Licensing'; Severity = if ($capacityPass) { 'Information' } else { 'Critical' }
                    Result = if (-not $licenseCapacity.Available) { 'UNKNOWN' } elseif ($capacityPass) { 'PASS' } else { 'FAIL' }
                    IsBlocking = -not $capacityPass; ObservedValue = $licenseCapacity.Message; ExpectedValue = "At least one available $targetSku license"
                    EvidenceSource = "$graphSource subscribedSkus"; Message = if ($capacityPass) { 'The target license SKU has available capacity.' } else { 'The target license SKU is missing or has no available capacity.' }
                    RecommendedAction = if ($capacityPass) { '' } else { 'Add license capacity or select an approved target SKU before the migration wave.' }
                })
            }
        }

        $hybridRequiredLive = [string]$Config['Mode'] -eq 'Live'
        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'HYBRID-ENDPOINT'; Category = 'HybridConnectivity'; Severity = if ($hybrid.ConnectivitySuccess) { 'Information' } elseif ($hybridRequiredLive) { 'Critical' } else { 'Warning' }
            Result = if (-not $hybrid.Available) { 'UNKNOWN' } elseif ($hybrid.ConnectivitySuccess) { 'PASS' } else { 'FAIL' }
            IsBlocking = $hybridRequiredLive -and -not $hybrid.ConnectivitySuccess; ObservedValue = "$($hybrid.EndpointName): $($hybrid.Message)"
            ExpectedValue = 'ExchangeRemoteMove endpoint test succeeds'; EvidenceSource = [string]$hybrid.Source
            Message = if ($hybrid.ConnectivitySuccess) { 'The migration endpoint connectivity test succeeded.' } elseif ($hybridRequiredLive) { 'The migration endpoint could not be validated.' } else { 'CacheOnly mode does not execute a live migration endpoint test.' }
            RecommendedAction = if ($hybrid.ConnectivitySuccess) { '' } elseif ($hybridRequiredLive) { 'Validate the endpoint, MRSProxy, WSSecurity, certificate, DNS and HTTPS connectivity.' } else { 'Run the final validation in Live mode before creating the migration batch.' }
        })
        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'ENTRA-CONNECT-SCHEDULER'; Category = 'EntraConnect'; Severity = if ($entraConnect.Available -and $entraConnect.SyncCycleEnabled -and -not $entraConnect.SchedulerSuspended) { 'Information' } else { 'Critical' }
            Result = if (-not $entraConnect.Available) { 'UNKNOWN' } elseif ($entraConnect.SyncCycleEnabled -and -not $entraConnect.SchedulerSuspended) { 'PASS' } else { 'FAIL' }
            IsBlocking = -not ($entraConnect.Available -and $entraConnect.SyncCycleEnabled -and -not $entraConnect.SchedulerSuspended)
            ObservedValue = $entraConnect.Message; ExpectedValue = 'Synchronization enabled and healthy'; EvidenceSource = [string]$entraConnect.Source
            Message = if ($entraConnect.Available) { "Microsoft Entra Connect sync health was collected from $($entraConnect.Source)." } else { 'Microsoft Entra Connect sync health could not be collected live or from the tenant cache.' }
            RecommendedAction = if ($entraConnect.Available) { 'Review synchronization health before migration if this check failed.' } else { 'Run on an Entra Connect server with ADSync cmdlets or refresh M365_Entra_AzureADConnect_SyncHealth.csv, then rerun.' }
        })
        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'KNOWN-ENHANCED-RESTORE'; Category = 'KnownErrors'; Severity = 'Warning'; Result = 'UNKNOWN'; IsBlocking = $false
            ObservedValue = 'No deterministic public pre-move property evaluated'; ExpectedValue = 'No enhanced restore mailbox restriction'
            EvidenceSource = 'Known error catalog'; Message = 'CannotMoveEnhancedRestoreMailboxesCrossOrgPermanentException cannot always be predicted from standard recipient properties.'
            RecommendedAction = 'Review prior migration reports and restore history. Treat a prior occurrence as NO-GO until Microsoft-supported remediation is confirmed.'
        })

        $mailboxEvidence = [pscustomobject][ordered]@{
            RunId = $runId
            EmailAddress = $email
            AdUserCount = @($ad.Users).Count
            OnPremMailboxCount = @($onPrem.Mailboxes).Count
            OnPremRemoteMailboxCount = @($onPrem.RemoteMailboxes).Count
            ExoRecipientCount = @($exo.Recipients).Count
            ExoMailboxCount = @($exo.Mailboxes).Count
            GraphUserCount = @($graph.Users).Count
            PermissionCount = @($onPrem.Permissions).Count
            CollectedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        }
        [void]$evidenceRows.Add($mailboxEvidence)
    }

    $summary = [System.Collections.Generic.List[object]]::new()
    $summaryRows = @($rows | Group-Object { ([string]$_.EmailAddress).ToLowerInvariant() } | ForEach-Object { $_.Group[0] })
    foreach ($row in $summaryRows) {
        $email = [string]$row.EmailAddress
        $mailFindings = @($findings | Where-Object EmailAddress -EQ $email)
        $blocking = @($mailFindings | Where-Object { $_.IsBlocking -and $_.Result -in @('FAIL', 'UNKNOWN') })
        $warnings = @($mailFindings | Where-Object Result -EQ 'WARN')
        $unknown = @($mailFindings | Where-Object Result -EQ 'UNKNOWN')
        $decision = if ($blocking.Count -gt 0) { 'NO-GO' } elseif ($unknown.Count -gt 0 -or $warnings.Count -gt 0) { 'GO-WARNING' } else { 'GO' }
        $blockingCodes = @($blocking | ForEach-Object { $_.CheckId } | Sort-Object -Unique)
        $actionableFindings = if ($blocking.Count -gt 0) { @($blocking) } else { @($unknown) + @($warnings) }
        $recommended = @($actionableFindings | ForEach-Object { $_.RecommendedAction } | Where-Object { $_ } | Sort-Object -Unique)
        [void]$summary.Add([pscustomobject][ordered]@{
            RunId = $runId
            BatchName = [System.IO.Path]::GetFileNameWithoutExtension($Batch.Path)
            EmailAddress = $email
            Decision = $decision
            BlockingCount = $blocking.Count
            WarningCount = $warnings.Count
            UnknownCount = $unknown.Count
            DataCoverage = if (@($blocking | Where-Object CheckId -match 'SOURCE').Count -gt 0) { 'Incomplete' } elseif ($unknown.Count -gt 0) { 'Partial' } else { 'Complete' }
            BlockingCodes = ($blockingCodes -join ';')
            RecommendedAction = ($recommended -join ' | ')
            CheckedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        })
    }

    return [pscustomobject]@{
        RunId = $runId
        Batch = $Batch
        Summary = @($summary)
        Findings = @($findings)
        PermissionsBaseline = @($permissionRows)
        Evidence = @($evidenceRows)
        Hybrid = $hybrid
        EntraConnect = $entraConnect
        SourceInitialization = $sourceInitialization
        Cancelled = ($CancellationCheck -and (& $CancellationCheck))
    }
}

function Export-SemrCsvFile {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$Columns
    )

    if (@($Data).Count -gt 0) {
        @($Data) | Select-Object -Property $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8
        return
    }
    $header = ($Columns | ForEach-Object { '"{0}"' -f ($_ -replace '"', '""') }) -join ','
    Set-Content -LiteralPath $Path -Value $header -Encoding utf8
}
function Export-SemrAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Assessment,
        [Parameter(Mandatory)][string]$OutputRoot
    )

    $runFolder = Join-Path $OutputRoot $Assessment.RunId
    New-Item -Path $runFolder -ItemType Directory -Force | Out-Null
    $summaryPath = Join-Path $runFolder 'Summary.csv'
    $findingsPath = Join-Path $runFolder 'Findings.csv'
    $permissionsPath = Join-Path $runFolder 'Permissions-Baseline.csv'
    $evidencePath = Join-Path $runFolder 'Evidence.csv'

    Export-SemrCsvFile -Data @($Assessment.Summary) -Path $summaryPath -Columns @('RunId','BatchName','EmailAddress','Decision','BlockingCount','WarningCount','UnknownCount','DataCoverage','BlockingCodes','RecommendedAction','CheckedAt')
    Export-SemrCsvFile -Data @($Assessment.Findings) -Path $findingsPath -Columns @('RunId','EmailAddress','CheckId','Category','Severity','Result','IsBlocking','ObservedValue','ExpectedValue','EvidenceSource','SourceTimestamp','Message','RecommendedAction')
    Export-SemrCsvFile -Data @($Assessment.PermissionsBaseline) -Path $permissionsPath -Columns @('RunId','EmailAddress','PermissionType','Delegate','IsInherited','Source','CapturedAt')
    Export-SemrCsvFile -Data @($Assessment.Evidence) -Path $evidencePath -Columns @('RunId','EmailAddress','AdUserCount','OnPremMailboxCount','OnPremRemoteMailboxCount','ExoRecipientCount','ExoMailboxCount','GraphUserCount','PermissionCount','CollectedAt')

    return [pscustomobject]@{
        RunFolder = $runFolder
        SummaryPath = $summaryPath
        FindingsPath = $findingsPath
        PermissionsPath = $permissionsPath
        EvidencePath = $evidencePath
    }
}

function Compare-SemrPermissionsBaseline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BaselinePath,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$CurrentPermissions
    )

    $baseline = @(Import-Csv -LiteralPath $BaselinePath)
    $key = {
        param($Row)
        "{0}|{1}|{2}" -f ([string]$Row.EmailAddress).ToLowerInvariant(), ([string]$Row.PermissionType).ToLowerInvariant(), ([string]$Row.Delegate).ToLowerInvariant()
    }
    $baselineKeys = @{}
    foreach ($row in $baseline) { $baselineKeys[(& $key $row)] = $row }
    $currentKeys = @{}
    foreach ($row in $CurrentPermissions) { $currentKeys[(& $key $row)] = $row }
    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($item in $baselineKeys.GetEnumerator()) {
        [void]$results.Add([pscustomobject]@{
            EmailAddress = $item.Value.EmailAddress
            PermissionType = $item.Value.PermissionType
            Delegate = $item.Value.Delegate
            Comparison = if ($currentKeys.ContainsKey($item.Key)) { 'MATCH' } else { 'MISSING-IN-CURRENT' }
        })
    }
    foreach ($item in $currentKeys.GetEnumerator()) {
        if ($baselineKeys.ContainsKey($item.Key)) { continue }
        [void]$results.Add([pscustomobject]@{
            EmailAddress = $item.Value.EmailAddress
            PermissionType = $item.Value.PermissionType
            Delegate = $item.Value.Delegate
            Comparison = 'NEW-IN-CURRENT'
        })
    }
    return @($results)
}

Export-ModuleMember -Function @(
    'Get-SemrVersion',
    'Get-SemrConfig',
    'Get-SemrConnectionState',
    'Connect-SemrActiveDirectory',
    'Connect-SemrOnPremisesExchange',
    'Connect-SemrExchangeOnline',
    'Connect-SemrMicrosoftGraph',
    'Disconnect-SemrSession',
    'Import-SemrBatchCsv',
    'Test-SemrHybridReadiness',
    'Test-SemrEntraConnect',
    'Invoke-SemrAssessment',
    'Export-SemrAssessment',
    'Get-SemrPostMigrationPermission',
    'Compare-SemrPermissionsBaseline'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDhVrQKzYP06HOo
# yv2YBbPyT221yU2tc+toh2vSAewxw6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINgJK2KB4iaL3YIND0ieVR35FZrCRaeqYG45ZKYWgEP3MA0GCSqG
# SIb3DQEBAQUABIIBgBsgGTxabdaryHlcrR11gJ85vC46hhsUTIDZYMu5UFQpeEIt
# 5ZGZPNJb+cljDExHukD4Z38KbK/Gb1CNZF927E8FSXa9VDA8lN6RFhdraAA7xRR6
# UnI+IFSfehk1yqYGQb71syENpzmUuFmn0N4n32DlOZBApsHQXOqch/EOyKUdA+gA
# 3eOlBtPEGwG0aDPWh66C0VMoHPhFLbHpSPtFWw/E8iTP7FnfQLF9vhZ1a7MHZGaf
# pKAXGlKfUBzOBkAq1umNODSL9DKyVVinDVlx0Y4LnEYwcpZ6I/IYM7hWn7QjHLnG
# WlfA5ITuhXTMJf8Y//ReIUqt+TFrENDGZzYtC3J2+MG47aRQVcJ4GzjKEigV1UVh
# 4kewwYnoGpIbcO6RZfmX0bQIUH3Z/76abp7tFKbM8gARh5E6MoXu0cT56ps6J9Xx
# pm4Eywc38w12MrYhUH4uI1hQqRyNApOwOo5gWebtaIdKpPR+HiRPPcE2jAKDgfZV
# Pkh953wQp83eqOIF66GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTcyMDAz
# MzNaMC8GCSqGSIb3DQEJBDEiBCDyQWY7f3/kLfuExYZUBIdk0z7+kQdH1rJHgUB6
# Ymhb4TANBgkqhkiG9w0BAQEFAASCAgA5GaDQcMUIuvwFK+OStqqu4eabjgIi+bsX
# eB/j874IOn9xQTgMWBWdkX9JdhJl23GxEiFQ8sWHgWFONZLc4+VdZQsUWvAUEyge
# F3363xMn9N37SM7XE9I/kfOjvLC3LivEtb8hZVh3FordQh0VX0aP916vKiDnTv4h
# 05LRt9i7pclsCj9qX5QiMjzkFDKMn/2WnK85SjJ9JTo4cROOoekVoXtjFHcdPDYb
# TJJBS/rBkoKrAlhjLHfBWHLv4CK/74ltuVOWEFNsnMggrRrCklKz0GSqYx0l8fSs
# ENk1KbpOrSNDzIEcGVCENY9iiLN1Sgleb/t4xe+eQegO2s+6J4X5QTGPssV0Uf28
# xo498FAyhNOKK49U/V1rJw27Y/BzGyObOmp/P16fqUEVyjky9ycrftTeKnJE2fmj
# +x3QDcpLBFbVJUS53vBZCiawsXjJ3C6uHSSQKwvGoS6KzGMdEzVSPo/Avm2sCku/
# FTux+dg+dUiBP6WqfnmR06bYad3xuJolKmHuX5LOcYBlxICY11xPLrbG5PPgedVR
# h+3jer2QAq65cRgxAvTNt5Dgevwy4ExIVzps2HEw9XYTi8CVQsoKKY9CbZ6nt6ia
# 8s1uf7QGB/JfnLfD4PoYj6rWnberkzkOZ5H40HubO2NxV5/rtdr2BLsmDqrJ/jvc
# URDlB9xIZw==
# SIG # End signature block
