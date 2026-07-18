Set-StrictMode -Version 2.0

$script:SemrVersion = '1.6.0'
$script:OnPremisesSession = $null
$script:InventoryContext = $null
$script:ActiveDirectoryDomains = @()
$script:ActiveDirectoryFallbackUsed = $false
$script:GraphEvidenceByEmail = @{}
$script:GraphSubscribedSkus = @()
$script:GraphSubscribedSkuError = ''
$script:CurrentSourceTimestamps = @{}
$script:ActiveDirectoryFallbackReasons = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:DisabledChecks = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
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

function Get-SemrCheckCatalog {
    $mandatory = @('CSV-EMPTY-IDENTITY','CSV-SMTP-FORMAT','CSV-DUPLICATE','AD-SOURCE','ONPREM-SOURCE','EXO-SOURCE','GRAPH-SOURCE')
    $definitions = @(
        @('CSV-EMPTY-IDENTITY','CSV','Mailbox identity is present','Reject empty mailbox identities.'),
        @('CSV-SMTP-FORMAT','CSV','SMTP syntax','Validate mailbox identity SMTP syntax.'),
        @('CSV-DUPLICATE','CSV','Batch duplicates','Reject duplicate mailbox rows.'),
        @('CSV-COLUMNS','CSV','Recognized columns','Report columns not interpreted by the application.'),
        @('CSV-MAILBOX-TYPE','CSV','Migration mailbox type','Validate PrimaryOnly, ArchiveOnly or PrimaryAndArchive.'),
        @('CSV-BADITEMLIMIT','CSV','Bad item limit','Validate the optional bad item limit.'),
        @('CSV-LARGEITEMLIMIT','CSV','Large item limit','Validate the optional large item limit.'),
        @('AD-SOURCE','Active Directory','AD source availability','Require live or cached Active Directory evidence.'),
        @('AD-IDENTITY-UNIQUE','Active Directory','Unique AD identity','Require exactly one matching AD user.'),
        @('AD-ACCOUNT-ENABLED','Active Directory','AD account enabled','Report disabled source accounts.'),
        @('PROXY-SMTP-GLOBAL-UNIQUE','Hybrid identity','Global SMTP uniqueness','Detect SMTP ownership conflicts across directory recipients.'),
        @('PROXY-INTERNAL-DUPLICATE','Hybrid identity','Internal proxy duplicates','Detect duplicate or malformed proxyAddresses values.'),
        @('TARGET-ADDRESS-GLOBAL-UNIQUE','Hybrid identity','targetAddress uniqueness','Detect duplicate or invalid target routing addresses.'),
        @('X500-LEGACYEXCHANGEDN','Hybrid identity','X500 preservation','Verify LegacyExchangeDN is preserved as an X500 proxy.'),
        @('SMTP-ACCEPTED-DOMAIN','Hybrid identity','Accepted SMTP domains','Verify every mailbox SMTP domain is accepted in Exchange Online.'),
        @('ONPREM-SOURCE','Exchange on-premises','On-premises source availability','Require live or cached Exchange on-premises evidence.'),
        @('ONPREM-RECIPIENT-STATE','Exchange on-premises','Recipient state','Require one UserMailbox and no pre-existing RemoteMailbox.'),
        @('RECIPIENT-TYPE-SUPPORTED','Exchange on-premises','Supported recipient type','Reject system and unsupported mailbox types.'),
        @('ONPREM-TARGET-ADDRESS','Exchange on-premises','Routing address domain','Validate the configured remote routing domain.'),
        @('ONPREM-PROXY-INTEGRITY','Exchange on-premises','Primary SMTP integrity','Require one primary SMTP proxy consistent with the mailbox.'),
        @('EXCHANGE-GUID-CONSISTENCY','Exchange identity','Exchange GUID consistency','Verify source and synchronized cloud recipient GUIDs.'),
        @('ARCHIVE-GUID-CONSISTENCY','Exchange identity','Archive GUID consistency','Verify archive GUID ownership and consistency.'),
        @('MAILBOX-TARGET-QUOTA','Mailbox','Target mailbox quota','Compare primary mailbox size with the safe target SKU quota.'),
        @('ARCHIVE-READINESS','Mailbox','Archive readiness','Validate archive state, size and target capability.'),
        @('MAILBOX-RECOVERABLE-ITEMS-QUOTA','Mailbox','Recoverable Items quota','Detect Recoverable Items quota saturation risk.'),
        @('MAILBOX-FOLDER-LIMITS','Mailbox','Folder and item limits','Detect folders or item counts near Exchange Online limits.'),
        @('MAILBOX-LARGE-ITEMS','Mailbox','Large items','Report oversized items against migration policy.'),
        @('CUSTOM-SOURCE-QUOTA','Mailbox','Custom source quota','Report custom on-premises mailbox quotas.'),
        @('MAILBOX-HOLDS','Compliance','Mailbox holds','Document Litigation Hold and In-Place Hold state.'),
        @('PERMISSIONS-BASELINE','Permissions','Permissions baseline','Capture Full Access, Send As and Send on Behalf.'),
        @('DELEGATE-MIGRATION-DEPENDENCY','Permissions','Delegate migration dependencies','Detect unresolved or cross-premises delegate dependencies.'),
        @('MAIL-FORWARDING','Mail flow','Mailbox forwarding','Document mailbox-level forwarding configuration.'),
        @('INBOX-FORWARDING-RULES','Mail flow','Inbox forwarding rules','Detect inbox rules that redirect or forward messages.'),
        @('DELIVERY-RESTRICTIONS','Mail flow','Delivery restrictions','Document moderation and delivery restrictions.'),
        @('EXCHANGE-DATABASE-HEALTH','Exchange infrastructure','Mailbox database health','Validate source mailbox database availability.'),
        @('EXO-SOURCE','Exchange Online','EXO source availability','Require live or cached Exchange Online evidence.'),
        @('EXO-RECIPIENT-STATE','Exchange Online','Cloud recipient state','Require one synchronized MailUser and no active cloud mailbox.'),
        @('EXO-SOFT-DELETED-CONFLICT','Exchange Online','Soft-deleted conflicts','Detect soft-deleted or inactive mailbox conflicts.'),
        @('EXO-EXISTING-MOVE','Migration','Existing move objects','Detect existing migration users and move requests.'),
        @('MOVE-HISTORY','Migration','Previous move history','Report failed, suspended or abandoned previous moves.'),
        @('GRAPH-SOURCE','Microsoft Graph','Graph source availability','Require live or cached Microsoft Graph evidence.'),
        @('GRAPH-USER-STATE','Microsoft Graph','Unique Entra user','Require exactly one matching Entra user.'),
        @('GRAPH-DIRSYNC','Microsoft Graph','Directory synchronization','Verify the user is synchronized from on-premises.'),
        @('ENTRA-OBJECT-SYNC-ERROR','Microsoft Graph','Object synchronization errors','Detect stale synchronization and identity anchor issues.'),
        @('LICENSE-PRE-MIGRATION','Licensing','Pre-migration licenses','Report licenses already assigned before migration.'),
        @('LICENSE-USAGE-LOCATION','Licensing','Usage location','Verify UsageLocation is populated.'),
        @('LICENSE-CAPACITY','Licensing','Target SKU capacity','Require available target SKU capacity.'),
        @('LICENSE-EXCHANGE-SERVICE-PLAN','Licensing','Exchange service plan','Require a mailbox-bearing Exchange service plan.'),
        @('HYBRID-ENDPOINT','Hybrid connectivity','Migration endpoint','Test the ExchangeRemoteMove endpoint.'),
        @('HYBRID-MRSPROXY','Hybrid connectivity','MRSProxy readiness','Verify MRSProxy and EWS migration readiness.'),
        @('HYBRID-CERTIFICATE-EXPIRY','Hybrid connectivity','Hybrid certificate expiry','Detect missing or expiring hybrid certificates.'),
        @('HYBRID-ENDPOINT-CAPACITY','Hybrid connectivity','Endpoint capacity','Report active migration load and endpoint limits.'),
        @('HYBRID-AUTODISCOVER-OAUTH','Hybrid connectivity','Autodiscover and OAuth','Validate hybrid organization relationship and OAuth configuration.'),
        @('ENTRA-CONNECT-SCHEDULER','Entra Connect','Entra Connect health','Require enabled and healthy synchronization.'),
        @('KNOWN-ENHANCED-RESTORE','Known errors','Enhanced Restore risk','Document the known Enhanced Restore cross-organization move risk.')
    )
    return @($definitions | ForEach-Object {
        [pscustomobject][ordered]@{
            CheckId = $_[0]
            Category = $_[1]
            Name = $_[2]
            Description = $_[3]
            Mandatory = $_[0] -in $mandatory
            CanDisable = $_[0] -notin $mandatory
            DefaultEnabled = $true
        }
    })
}

function Set-SemrAssessmentCheckOptions {
    param([System.Collections.IDictionary]$Config)
    $script:DisabledChecks.Clear()
    if (-not $Config -or -not $Config.Contains('DisabledChecks')) { return }
    foreach ($checkId in @($Config['DisabledChecks'])) {
        if (-not [string]::IsNullOrWhiteSpace([string]$checkId)) { [void]$script:DisabledChecks.Add(([string]$checkId).Trim()) }
    }
}

function Test-SemrCheckEnabled {
    param([Parameter(Mandatory)][string]$CheckId)
    $definition = @(Get-SemrCheckCatalog | Where-Object CheckId -EQ $CheckId | Select-Object -First 1)
    if ($definition.Count -eq 1 -and $definition[0].Mandatory) { return $true }
    return -not $script:DisabledChecks.Contains($CheckId)
}

function ConvertTo-SemrHashtable {
    param([Parameter(Mandatory)][AllowNull()]$InputObject)

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [string] -or $InputObject -is [ValueType]) {
        return $InputObject
    }

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

function Find-SemrUpwardFile {
    param(
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)][string]$StartPath
    )

    $current = $StartPath
    while ($current) {
        $candidate = Join-Path -Path $current -ChildPath $RelativePath
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    return ''
}

function Import-SemrSmartM365TenantProfile {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Runtime)

    $globalPath = Find-SemrUpwardFile -RelativePath 'Config\SmartM365.global.local.json' -StartPath $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($globalPath)) { return $false }

    $globalConfig = ConvertTo-SemrHashtable -InputObject (Get-Content -LiteralPath $globalPath -Raw | ConvertFrom-Json)
    $tenantProfile = $Runtime['TenantProfile']
    $profileKey = [string]$tenantProfile['ProfileKey']
    if ([string]::IsNullOrWhiteSpace($profileKey) -or $profileKey -eq 'tenant') {
        $profileKey = [string]$globalConfig['DefaultTenant']
    }
    if ([string]::IsNullOrWhiteSpace($profileKey) -or $profileKey -notmatch '^[a-zA-Z0-9-]+$') { return $false }

    $tenantPath = Join-Path -Path (Split-Path -Path $globalPath -Parent) -ChildPath ("Tenants\{0}.local.json" -f $profileKey)
    if (-not (Test-Path -LiteralPath $tenantPath -PathType Leaf)) { return $false }

    $smartM365Profile = ConvertTo-SemrHashtable -InputObject (Get-Content -LiteralPath $tenantPath -Raw | ConvertFrom-Json)
    $tenantId = [string]$smartM365Profile['TenantId']
    if ([string]::IsNullOrWhiteSpace($tenantId)) { return $false }

    $tenantProfile['ProfileKey'] = $profileKey
    if ([string]::IsNullOrWhiteSpace([string]$tenantProfile['TenantId'])) {
        $tenantProfile['TenantId'] = $tenantId
    }
    $remoteRoutingDomain = [string]$smartM365Profile['RemoteRoutingDomain']
    if (([string]::IsNullOrWhiteSpace([string]$tenantProfile['RemoteRoutingDomain']) -or [string]$tenantProfile['RemoteRoutingDomain'] -eq 'tenant.mail.onmicrosoft.com') -and -not [string]::IsNullOrWhiteSpace($remoteRoutingDomain)) {
        $tenantProfile['RemoteRoutingDomain'] = $remoteRoutingDomain
    }
    return $true
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
    foreach ($deprecatedKey in @('Tenant', 'OnPremises', 'EntraConnect', 'SmartM365', 'EvidenceSources', 'Authentication')) {
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
    $assessmentPhase = [string]$runtime['AssessmentPhase']
    if ($assessmentPhase -notin @('PreCreation', 'ExistingBatch')) {
        throw "Invalid AssessmentPhase '$assessmentPhase'. Expected PreCreation or ExistingBatch."
    }
    $tenantProfile = $runtime['TenantProfile']
    [void](Import-SemrSmartM365TenantProfile -Runtime $runtime)
    if ([string]::IsNullOrWhiteSpace([string]$tenantProfile['TenantId'])) {
        throw 'TenantProfile.TenantId is required. Configure it in the application JSON or provide the SmartM365 default tenant profile under Config\Tenants.'
    }
    $tenantId = ([string]$tenantProfile['TenantId']).Trim()
    $tenantProfile['TenantId'] = $tenantId
    try { [void][guid]::Parse($tenantId) } catch { throw "TenantProfile.TenantId '$tenantId' is not a valid GUID." }
    $targetDeliveryDomain = [string]$runtime['Hybrid']['TargetDeliveryDomain']
    if ([string]::IsNullOrWhiteSpace($targetDeliveryDomain) -or $targetDeliveryDomain -eq 'tenant.mail.onmicrosoft.com') {
        $runtime['Hybrid']['TargetDeliveryDomain'] = [string]$tenantProfile['RemoteRoutingDomain']
    }
    $cacheRootPath = Resolve-SemrConfigPath -Path ([string]$runtime['Cache']['RootPath']) -BasePath $PSScriptRoot
    $alternativeCacheRootPath = Resolve-SemrConfigPath -Path ([string]$runtime['Cache']['AlternativeRootPath']) -BasePath $PSScriptRoot
    $cacheDataLastPath = if ($cacheRootPath -and [IO.Path]::GetFileName($cacheRootPath.TrimEnd('\')) -ieq 'DATA-LAST') { $cacheRootPath } elseif ($cacheRootPath) { Join-Path $cacheRootPath 'DATA-LAST' } else { '' }
    $alternativeDataLastPath = if ($alternativeCacheRootPath -and [IO.Path]::GetFileName($alternativeCacheRootPath.TrimEnd('\')) -ieq 'DATA-LAST') { $alternativeCacheRootPath } elseif ($alternativeCacheRootPath) { Join-Path $alternativeCacheRootPath 'DATA-LAST' } else { '' }

    $runtime['_RuntimePath'] = $Path
    $runtime['_AddedKeys'] = $added
    $runtime['_TenantProfileKey'] = [string]$tenantProfile['ProfileKey']
    $runtime['_TenantId'] = $tenantId
    $runtime['_RemoteRoutingDomain'] = [string]$tenantProfile['RemoteRoutingDomain']
    $runtime['_CacheRootPath'] = $cacheRootPath
    $runtime['_InventoryDataLastPath'] = $cacheDataLastPath
    $runtime['_AlternativeCacheRootPath'] = $alternativeCacheRootPath
    $runtime['_AlternativeInventoryDataLastPath'] = $alternativeDataLastPath
    $runtime['_AllowStaleCache'] = $false
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

function Test-SemrObjectNotFoundError {
    param([string]$Message)

    return -not [string]::IsNullOrWhiteSpace($Message) -and
        $Message -match '(?i)not found|couldn''t be found|could not be found|does not exist|no such request exists in (?:the )?specified index'
}

function Invoke-SemrIsolatedPowerShell {
    param(
        [Parameter(Mandatory)][string]$ScriptText,
        [scriptblock]$ProgressCallback
    )

    $pwshCommand = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    if (-not $pwshCommand) {
        throw 'PowerShell 7 (pwsh.exe) is required for Microsoft Graph module validation and installation.'
    }
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptText))
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $pwshCommand.Source
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-EncodedCommand',$encodedCommand)) {
        [void]$startInfo.ArgumentList.Add([string]$argument)
    }

    $process = $null
    try {
        $process = [Diagnostics.Process]::Start($startInfo)
        if (-not $process) { throw 'The isolated PowerShell 7 process could not be started.' }
        $outputTask = $process.StandardOutput.ReadToEndAsync()
        $errorTask = $process.StandardError.ReadToEndAsync()
        while (-not $process.WaitForExit(500)) {
            if ($ProgressCallback) { & $ProgressCallback }
        }
        $process.WaitForExit()
        return [pscustomobject]@{
            ExitCode = $process.ExitCode
            StandardOutput = $outputTask.GetAwaiter().GetResult().Trim()
            StandardError = $errorTask.GetAwaiter().GetResult().Trim()
        }
    }
    finally {
        if ($process) { $process.Dispose() }
    }
}

function Get-SemrMicrosoftGraphModuleState {
    [CmdletBinding()]
    param(
        [string[]]$RequiredModules = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Identity.DirectoryManagement'
        )
    )

    $modules = @($RequiredModules | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $moduleLiteral = ($modules | ForEach-Object { "'$($_.Replace("'", "''"))'" }) -join ','
    $probeScript = "`$ErrorActionPreference='Stop'; `$required=@($moduleLiteral); `$required | Where-Object { -not (Get-Module -ListAvailable -Name `$_) } | Sort-Object -Unique | ForEach-Object { [Console]::Out.WriteLine(`$_) }"
    $probe = Invoke-SemrIsolatedPowerShell -ScriptText $probeScript
    if ($probe.ExitCode -ne 0) {
        $detail = if ($probe.StandardError) { $probe.StandardError } else { $probe.StandardOutput }
        throw "Microsoft Graph module validation failed in PowerShell 7: $detail"
    }
    $missingModules = @($probe.StandardOutput -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    return [pscustomobject]@{
        Available = $missingModules.Count -eq 0
        RequiredModules = $modules
        MissingModules = $missingModules
    }
}

function Install-SemrMicrosoftGraphModule {
    [CmdletBinding()]
    param(
        [string[]]$ModuleNames = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Identity.DirectoryManagement'
        ),
        [scriptblock]$ProgressCallback
    )

    $modules = @($ModuleNames | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    if ($modules.Count -eq 0) { return Get-SemrMicrosoftGraphModuleState }
    $moduleLiteral = ($modules | ForEach-Object { "'$($_.Replace("'", "''"))'" }) -join ','
    $installScript = "`$ErrorActionPreference='Stop'; `$ProgressPreference='SilentlyContinue'; if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) { throw 'Install-Module is unavailable in PowerShell 7.' }; foreach (`$moduleName in @($moduleLiteral)) { if (-not (Get-Module -ListAvailable -Name `$moduleName)) { Install-Module -Name `$moduleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop } }"
    $install = Invoke-SemrIsolatedPowerShell -ScriptText $installScript -ProgressCallback $ProgressCallback
    if ($install.ExitCode -ne 0) {
        $detail = if ($install.StandardError) { $install.StandardError } else { $install.StandardOutput }
        throw "Microsoft Graph module installation failed in PowerShell 7: $detail"
    }
    return Get-SemrMicrosoftGraphModuleState
}

function Connect-SemrActiveDirectory {
    [CmdletBinding()]
    param()

    Import-Module ActiveDirectory -ErrorAction Stop
    $forest = Get-ADForest -ErrorAction Stop
    $script:ActiveDirectoryDomains = @(
        $forest.Domains |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
    if ($script:ActiveDirectoryDomains.Count -eq 0) {
        throw "No domain was returned for Active Directory forest '$([string]$forest.Name)'."
    }
    $domainFailures = [System.Collections.Generic.List[string]]::new()
    foreach ($domain in $script:ActiveDirectoryDomains) {
        try {
            [void](Get-ADDomain -Identity $domain -Server $domain -ErrorAction Stop)
        }
        catch {
            [void]$domainFailures.Add("${domain}: $($_.Exception.Message)")
        }
    }
    if ($domainFailures.Count -gt 0) {
        $script:ActiveDirectoryDomains = @()
        throw "Active Directory forest coverage is incomplete ($($domainFailures.Count) domain(s) unavailable: $($domainFailures -join '; '))."
    }
    $script:ActiveDirectoryFallbackUsed = $false
    $script:ConnectionState.ActiveDirectory = $true
    return Get-SemrConnectionState
}

function Enable-SemrExchangeForestView {
    $commandName = if (Test-SemrCommand -Name 'Set-OnPremADServerSettings') {
        'Set-OnPremADServerSettings'
    }
    elseif (Test-SemrCommand -Name 'Set-ADServerSettings') {
        'Set-ADServerSettings'
    }
    else {
        ''
    }
    if (-not $commandName) {
        throw 'Set-ADServerSettings is unavailable in the Exchange on-premises session; ViewEntireForest cannot be enabled.'
    }
    & $commandName -ViewEntireForest $true -ErrorAction Stop | Out-Null
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
        Enable-SemrExchangeForestView
        $script:ConnectionState.OnPremisesExchange = $true
        return Get-SemrConnectionState
    }

    if ((Test-SemrCommand -Name 'Get-Mailbox') -and (Test-SemrCommand -Name 'Get-RemoteMailbox')) {
        Enable-SemrExchangeForestView
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
    Enable-SemrExchangeForestView

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
        [string]$TenantId = '',
        [string[]]$EmailAddresses = @(),
        [scriptblock]$ProgressCallback
    )

    $normalizedScopes = @($Scopes | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $invalidScopes = @($normalizedScopes | Where-Object { $_ -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]+$' })
    if ($normalizedScopes.Count -eq 0 -or $invalidScopes.Count -gt 0) {
        throw "Microsoft Graph scopes are invalid: $($normalizedScopes -join ', ')"
    }
    $workerPath = Join-Path $PSScriptRoot 'SmartM365-ExchangeMigrationReadiness-GraphWorker.ps1'
    if (-not (Test-Path -LiteralPath $workerPath -PathType Leaf)) {
        throw "Microsoft Graph worker is missing: $workerPath"
    }
    $pwshCommand = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    if (-not $pwshCommand) {
        throw 'PowerShell 7 (pwsh.exe) is required for the isolated Microsoft Graph connection.'
    }

    $runtimeRoot = Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-ExchangeMigrationReadiness\Graph-{0}" -f [guid]::NewGuid().ToString('N'))
    $scopesPath = Join-Path $runtimeRoot 'scopes.clixml'
    $inputPath = Join-Path $runtimeRoot 'mailboxes.clixml'
    $outputPath = Join-Path $runtimeRoot 'evidence.clixml'
    $errorPath = Join-Path $runtimeRoot 'error.txt'
    $progressPath = Join-Path $runtimeRoot 'progress.txt'
    $process = $null
    try {
        [void](New-Item -ItemType Directory -Path $runtimeRoot -Force)
        $normalizedScopes | Export-Clixml -LiteralPath $scopesPath -Force
        @($EmailAddresses | Where-Object { $_ } | Sort-Object -Unique) | Export-Clixml -LiteralPath $inputPath -Force

        $startInfo = [Diagnostics.ProcessStartInfo]::new()
        $startInfo.FileName = $pwshCommand.Source
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        foreach ($argument in @(
            '-NoLogo', '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', $workerPath,
            '-TenantId', $TenantId,
            '-ScopesPath', $scopesPath, '-InputPath', $inputPath,
            '-OutputPath', $outputPath, '-ErrorPath', $errorPath, '-ProgressPath', $progressPath
        )) {
            [void]$startInfo.ArgumentList.Add([string]$argument)
        }

        $process = [Diagnostics.Process]::Start($startInfo)
        if (-not $process) {
            throw 'The isolated Microsoft Graph process could not be started.'
        }
        $lastProgress = ''
        while (-not $process.WaitForExit(250)) {
            if ($ProgressCallback -and (Test-Path -LiteralPath $progressPath -PathType Leaf)) {
                try {
                    $progress = [IO.File]::ReadAllText($progressPath)
                    if ($progress -and $progress -ne $lastProgress) {
                        $lastProgress = $progress
                        $parts = @($progress.Split('|', 3))
                        $message = if ($parts.Count -eq 3) { "Microsoft Graph [$($parts[0])/$($parts[1])] - $($parts[2])" } else { 'Collecting Microsoft Graph evidence...' }
                        & $ProgressCallback $message
                    }
                }
                catch { $null = $_ }
            }
        }
        if ($process.ExitCode -ne 0) {
            $workerError = if (Test-Path -LiteralPath $errorPath -PathType Leaf) { [IO.File]::ReadAllText($errorPath) } else { "Worker exit code $($process.ExitCode)." }
            throw "Microsoft Graph isolated authentication or collection failed: $workerError"
        }
        if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
            throw 'Microsoft Graph worker completed without returning evidence.'
        }

        $workerResult = Import-Clixml -LiteralPath $outputPath
        if ($TenantId -and [string]$workerResult.TenantId -ne $TenantId) {
            throw "Microsoft Graph connected to tenant '$($workerResult.TenantId)' instead of configured tenant '$TenantId'."
        }
        $script:GraphEvidenceByEmail = @{}
        foreach ($entry in @($workerResult.Evidence)) {
            $key = ([string]$entry.EmailAddress).Trim().ToLowerInvariant()
            if ($key) { $script:GraphEvidenceByEmail[$key] = $entry }
        }
        $script:GraphSubscribedSkus = @($workerResult.SubscribedSkus)
        $script:GraphSubscribedSkuError = [string]$workerResult.SubscribedSkuError
        $script:ConnectionState.MicrosoftGraph = $true
        return Get-SemrConnectionState
    }
    finally {
        if ($process) { $process.Dispose() }
        if (Test-Path -LiteralPath $runtimeRoot) {
            Remove-Item -LiteralPath $runtimeRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
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
    foreach ($key in @($script:ConnectionState.Keys)) {
        $script:ConnectionState[$key] = $false
    }
    $script:ActiveDirectoryDomains = @()
    $script:ActiveDirectoryFallbackUsed = $false
    $script:GraphEvidenceByEmail = @{}
    $script:GraphSubscribedSkus = @()
    $script:GraphSubscribedSkuError = ''
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
            $result.ActiveDirectoryMessage = "Live Active Directory forest connection succeeded. $($script:ActiveDirectoryDomains.Count) domain(s) will be searched."
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
            $result.ExchangeOnPremisesMessage = 'Live Exchange on-premises connection succeeded with ViewEntireForest enabled.'
        }
        catch {
            $result.ExchangeOnPremisesMessage = "Live Exchange on-premises unavailable; CSV fallback selected. $($_.Exception.Message)"
        }
    }
    else {
        $result.ExchangeOnPremisesMessage = 'Live Exchange on-premises was already connected with ViewEntireForest enabled.'
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
        SourceTimestamp = (Get-Item -LiteralPath $Path).LastWriteTime
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
        $SourceTimestamp = $null,
        [string]$Message = '',
        [string]$RecommendedAction = ''
    )

    $sourceTimestampText = ''
    if ($null -ne $SourceTimestamp -and -not [string]::IsNullOrWhiteSpace([string]$SourceTimestamp)) {
        try { $sourceTimestampText = ([datetime]$SourceTimestamp).ToString('yyyy-MM-dd HH:mm:ss') }
        catch { $sourceTimestampText = [string]$SourceTimestamp }
    }

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
        SourceTimestamp = $sourceTimestampText
        Message = $Message
        RecommendedAction = $RecommendedAction
    }
}

function Add-SemrFinding {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$List,
        [Parameter(Mandatory)][hashtable]$Parameters
    )
    if (-not (Test-SemrCheckEnabled -CheckId ([string]$Parameters.CheckId))) { return }
    if (-not $Parameters.ContainsKey('SourceTimestamp')) {
        $resolvedTimestamp = $null
        $evidenceSource = ([string]$Parameters.EvidenceSource).Trim().ToLowerInvariant()
        if ($evidenceSource -and $script:CurrentSourceTimestamps) {
            foreach ($sourceKey in @($script:CurrentSourceTimestamps.Keys | Sort-Object Length -Descending)) {
                if ($evidenceSource.Contains(([string]$sourceKey).ToLowerInvariant())) {
                    $resolvedTimestamp = $script:CurrentSourceTimestamps[$sourceKey]
                    break
                }
            }
        }
        $Parameters.SourceTimestamp = $resolvedTimestamp
    }
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
        [double]$MaximumAgeHours = 48,
        [switch]$AllowStale
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ Available = $false; Present = $false; Fresh = $false; AcceptedStale = $false; Path = $Path; Timestamp = $null; AgeHours = $null; Message = "Inventory file not found: $Path" }
    }
    $item = Get-Item -LiteralPath $Path
    $ageHours = ((Get-Date) - $item.LastWriteTime).TotalHours
    $fresh = $MaximumAgeHours -le 0 -or $ageHours -le $MaximumAgeHours
    $acceptedStale = -not $fresh -and $AllowStale
    return [pscustomobject]@{
        Available = $fresh -or $acceptedStale
        Present = $true
        Fresh = $fresh
        AcceptedStale = $acceptedStale
        Path = $item.FullName
        Timestamp = $item.LastWriteTime
        AgeHours = [math]::Round($ageHours, 2)
        Message = if ($fresh) { "Inventory available; age $([math]::Round($ageHours, 2)) hour(s)." } elseif ($acceptedStale) { "Inventory is stale but accepted for this run; age $([math]::Round($ageHours, 2)) hour(s), maximum $MaximumAgeHours." } else { "Inventory is stale; age $([math]::Round($ageHours, 2)) hour(s), maximum $MaximumAgeHours." }
    }
}

function Get-SemrInventoryFolders {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Config)

    $folders = [System.Collections.Generic.List[string]]::new()
    foreach ($folder in @(
        [string]$Config['_InventoryDataLastPath'],
        [string]$Config['_CacheRootPath'],
        [string]$Config['_AlternativeInventoryDataLastPath'],
        [string]$Config['_AlternativeCacheRootPath']
    )) {
        if ([string]::IsNullOrWhiteSpace($folder)) { continue }
        if (-not $folders.Contains($folder)) { [void]$folders.Add($folder) }
    }
    return @($folders)
}

function Resolve-SemrInventoryFilePath {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Config,
        [Parameter(Mandatory)][string]$FileName
    )

    $candidates = [System.Collections.Generic.List[string]]::new()
    foreach ($folder in @(Get-SemrInventoryFolders -Config $Config)) {
        $candidate = Join-Path $folder $FileName
        if (-not $candidates.Contains($candidate)) { [void]$candidates.Add($candidate) }
    }
    $existing = @($candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | ForEach-Object { Get-Item -LiteralPath $_ } | Sort-Object LastWriteTime -Descending)
    if ($existing.Count -gt 0) { return $existing[0].FullName }
    if ($candidates.Count -gt 0) { return $candidates[0] }
    return $FileName
}

function Get-SemrCsvSourceInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Config,
        [AllowNull()]$Batch,
        [AllowNull()]$EntraConnect,
        [switch]$AssessmentCompleted
    )

    Set-SemrAssessmentCheckOptions -Config $Config
    $maximumAgeHours = [double]$Config['Cache']['MaximumAgeHours']
    $allowStale = [bool]$Config['_AllowStaleCache']
    $mode = [string]$Config['Mode']
    $connectionState = Get-SemrConnectionState
    $results = [System.Collections.Generic.List[object]]::new()

    if ($Batch -and -not [string]::IsNullOrWhiteSpace([string]$Batch.Path)) {
        $batchState = Get-SemrInventoryFileState -Path ([string]$Batch.Path) -MaximumAgeHours 0
        [void]$results.Add([pscustomobject][ordered]@{
            FileName = [IO.Path]::GetFileName([string]$Batch.Path)
            Category = 'Migration batch'
            ExpectedUse = 'Required input'
            Path = [string]$batchState.Path
            Present = [bool]$batchState.Present
            Fresh = [bool]$batchState.Present
            Used = [bool]$batchState.Present
            Status = if ($batchState.Present) { 'Used' } else { 'Missing' }
            LastWriteTime = if ($batchState.Timestamp) { ([datetime]$batchState.Timestamp).ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            AgeHours = $batchState.AgeHours
            Details = 'Mailbox identities selected for this assessment.'
        })
    }

    $definitions = @(
        [pscustomobject]@{ FileName = 'AD_Users_AllDomains.csv'; Category = 'Active Directory'; ExpectedUse = 'Live fallback / CacheOnly'; SourceName = 'ActiveDirectory' },
        [pscustomobject]@{ FileName = 'Exchange_OnPrem_Mailboxes_AllDomains.csv'; Category = 'Exchange on-premises'; ExpectedUse = 'Live fallback / CacheOnly'; SourceName = 'ExchangeOnPremises' },
        [pscustomobject]@{ FileName = 'M365_Entra_AzureADConnect_SyncHealth.csv'; Category = 'Microsoft Entra Connect'; ExpectedUse = 'Live fallback / CacheOnly'; SourceName = 'EntraConnect' },
        [pscustomobject]@{ FileName = 'M365_Users_Active.csv'; Category = 'Microsoft Graph users'; ExpectedUse = 'CacheOnly'; SourceName = 'CloudCacheOnly' },
        [pscustomobject]@{ FileName = 'Exchange_EXO_Mailboxes_AllDomains.csv'; Category = 'Exchange Online recipients'; ExpectedUse = 'CacheOnly'; SourceName = 'CloudCacheOnly' },
        [pscustomobject]@{ FileName = 'Exchange_EXO_MigrationJobs.csv'; Category = 'Exchange Online migration jobs'; ExpectedUse = 'CacheOnly'; SourceName = 'CloudCacheOnly'; CheckIds = @('EXO-EXISTING-MOVE','MOVE-HISTORY') },
        [pscustomobject]@{ FileName = 'M365_Licenses_Tenant.csv'; Category = 'Microsoft 365 licensing'; ExpectedUse = 'CacheOnly'; SourceName = 'CloudCacheOnly'; CheckIds = @('LICENSE-CAPACITY') },
        [pscustomobject]@{ FileName = 'M365_Licenses_ServicePlans.csv'; Category = 'Exchange license service plans'; ExpectedUse = 'CacheOnly'; SourceName = 'CloudCacheOnly'; CheckIds = @('LICENSE-EXCHANGE-SERVICE-PLAN') },
        [pscustomobject]@{ FileName = 'AD_Users_DuplicateSMTP.csv'; Category = 'Global SMTP uniqueness'; ExpectedUse = 'Live fallback / CacheOnly'; SourceName = 'ActiveDirectory'; CheckIds = @('PROXY-SMTP-GLOBAL-UNIQUE') },
        [pscustomobject]@{ FileName = 'AD_Users_DuplicateRemoteRoutingAddress.csv'; Category = 'Remote routing uniqueness'; ExpectedUse = 'Live fallback / CacheOnly'; SourceName = 'ActiveDirectory'; CheckIds = @('TARGET-ADDRESS-GLOBAL-UNIQUE') },
        [pscustomobject]@{ FileName = 'Exchange_OnPrem_ProxyAddresses_Check.csv'; Category = 'Exchange proxy and routing readiness'; ExpectedUse = 'Live fallback / CacheOnly'; SourceName = 'ExchangeOnPremises'; CheckIds = @('PROXY-SMTP-GLOBAL-UNIQUE','TARGET-ADDRESS-GLOBAL-UNIQUE') },
        [pscustomobject]@{ FileName = 'Exchange_EXO_AcceptedDomains.csv'; Category = 'Exchange Online accepted domains'; ExpectedUse = 'CacheOnly'; SourceName = 'CloudCacheOnly'; CheckIds = @('SMTP-ACCEPTED-DOMAIN') },
        [pscustomobject]@{ FileName = 'Exchange_EXO_Mailboxes_AllDomains_Archive.csv'; Category = 'Exchange Online archives'; ExpectedUse = 'CacheOnly'; SourceName = 'CloudCacheOnly'; CheckIds = @('ARCHIVE-READINESS','ARCHIVE-GUID-CONSISTENCY') },
        [pscustomobject]@{ FileName = 'Exchange_OnPrem_MigrationReadiness_Config.csv'; Category = 'Hybrid configuration'; ExpectedUse = 'Live fallback / CacheOnly'; SourceName = 'ExchangeOnPremises'; CheckIds = @('HYBRID-MRSPROXY','HYBRID-CERTIFICATE-EXPIRY','HYBRID-AUTODISCOVER-OAUTH','EXCHANGE-DATABASE-HEALTH') }
    )

    foreach ($definition in $definitions) {
        $filePath = Resolve-SemrInventoryFilePath -Config $Config -FileName $definition.FileName
        $state = Get-SemrInventoryFileState -Path $filePath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        $used = $false
        $status = ''
        $checkIds = @(Get-SemrPropertyValue -InputObject $definition -Names @('CheckIds') -Default @())
        $relatedChecksEnabled = $checkIds.Count -eq 0 -or @($checkIds | Where-Object { Test-SemrCheckEnabled -CheckId $_ }).Count -gt 0

        if (-not $relatedChecksEnabled) {
            $status = 'Not used - related checks disabled'
        }
        elseif ($mode -eq 'CacheOnly') {
            $used = [bool]$state.Available
            $status = if (-not $state.Present) { 'Required - Missing' } elseif ($state.AcceptedStale) { 'Used - Stale accepted' } elseif (-not $state.Fresh) { 'Required - Stale' } else { 'Used' }
        }
        elseif ($definition.SourceName -eq 'CloudCacheOnly') {
            $status = 'Not used - Live EXO/Graph source'
        }
        elseif (-not $AssessmentCompleted) {
            $status = if (-not $state.Present) { 'Fallback missing' } elseif (-not $state.Fresh) { 'Fallback stale' } else { 'Fallback available' }
        }
        else {
            $liveSourceUsed = switch ($definition.SourceName) {
                'ActiveDirectory' { [bool]$connectionState.ActiveDirectory -and -not $script:ActiveDirectoryFallbackUsed }
                'ExchangeOnPremises' { [bool]$connectionState.OnPremisesExchange }
                'EntraConnect' { if ($EntraConnect) { [string]$EntraConnect.Source -notmatch 'CSV' } else { [bool]$connectionState.EntraConnect } }
                default { $true }
            }
            if ($liveSourceUsed) {
                $status = 'Not used - Live source succeeded'
            }
            else {
                $used = [bool]$state.Available
                $status = if (-not $state.Present) { 'Fallback required - Missing' } elseif ($state.AcceptedStale) { 'Used - Stale CSV fallback accepted' } elseif (-not $state.Fresh) { 'Fallback required - Stale' } else { 'Used - CSV fallback' }
            }
        }

        [void]$results.Add([pscustomobject][ordered]@{
            FileName = $definition.FileName
            Category = $definition.Category
            ExpectedUse = $definition.ExpectedUse
            Path = [string]$state.Path
            Present = [bool]$state.Present
            Fresh = [bool]$state.Fresh
            Used = $used
            Status = $status
            LastWriteTime = if ($state.Timestamp) { ([datetime]$state.Timestamp).ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            AgeHours = $state.AgeHours
            Details = [string]$state.Message
        })
    }

    return @($results)
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

function Import-SemrFlexibleCsv {
    param([Parameter(Mandatory)][string]$Path)
    $file = Read-SemrTextFile -Path $Path
    $firstLine = @($file.Text -split '\r?\n' | Select-Object -First 1)
    if ($firstLine.Count -eq 0 -or [string]::IsNullOrWhiteSpace($firstLine[0])) { return @() }
    $delimiter = Get-SemrDelimiter -Header $firstLine[0]
    return @($file.Text | ConvertFrom-Csv -Delimiter $delimiter)
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
        LicenseServicePlanDataAvailable = $false
        DuplicateSmtpDataAvailable = $false
        DuplicateRoutingDataAvailable = $false
        ProxyCheckDataAvailable = $false
        AcceptedDomainDataAvailable = $false
        ArchiveDataAvailable = $false
        HybridConfigDataAvailable = $false
        CloudUserRowsByEmail = @{}
        ExoMailboxRowsByEmail = @{}
        MigrationRowsByEmail = @{}
        LicenseRows = @()
        LicenseServicePlanRows = @()
        DuplicateSmtpRows = @()
        DuplicateRoutingRows = @()
        ProxyCheckRowsByEmail = @{}
        AcceptedDomainRows = @()
        ArchiveRowsByEmail = @{}
        HybridConfigRows = @()
        CloudTimestamp = $null
        ExoTimestamp = $null
        MigrationTimestamp = $null
        LicenseTimestamp = $null
        LicenseServicePlanTimestamp = $null
        AdvancedIdentityTimestamp = $null
        AcceptedDomainTimestamp = $null
        ArchiveTimestamp = $null
        HybridConfigTimestamp = $null
    }

    # Always preload these inventories: they are mandatory in CacheOnly and automatic fallbacks in Live.
    $usesAdInventory = $true
    $usesExchangeInventory = $true
    $usesEntraInventory = $true

    $availableCacheFolders = @(Get-SemrInventoryFolders -Config $Config | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
    if ($availableCacheFolders.Count -gt 0) {
        $script:InventoryContext.DataLastPath = $availableCacheFolders[0]
    }
    if ($availableCacheFolders.Count -eq 0) {
        $configuredFolders = @(Get-SemrInventoryFolders -Config $Config)
        $message = "No configured CSV cache folder is accessible: $($configuredFolders -join '; ')"
        $script:InventoryContext.ActiveDirectoryMessage = $message
        $script:InventoryContext.ExchangeOnPremisesMessage = $message
        $script:InventoryContext.EntraConnectMessage = $message
        return [pscustomobject]$script:InventoryContext
    }

    $maximumAgeHours = [double]$Config['Cache']['MaximumAgeHours']
    $allowStale = [bool]$Config['_AllowStaleCache']
    $emails = @($BatchRows | ForEach-Object { [string]$_.EmailAddress } | Where-Object { $_ })
    $adPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'AD_Users_AllDomains.csv'
    $mailboxPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
    $entraPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'M365_Entra_AzureADConnect_SyncHealth.csv'
    $adState = Get-SemrInventoryFileState -Path $adPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
    $mailboxState = Get-SemrInventoryFileState -Path $mailboxPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
    $entraState = Get-SemrInventoryFileState -Path $entraPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale

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

    $advancedIdentityChecks = @('PROXY-SMTP-GLOBAL-UNIQUE','TARGET-ADDRESS-GLOBAL-UNIQUE','PROXY-INTERNAL-DUPLICATE')
    if (@($advancedIdentityChecks | Where-Object { Test-SemrCheckEnabled -CheckId $_ }).Count -gt 0) {
        $duplicateSmtpPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'AD_Users_DuplicateSMTP.csv'
        $duplicateRoutingPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'AD_Users_DuplicateRemoteRoutingAddress.csv'
        $proxyCheckPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'Exchange_OnPrem_ProxyAddresses_Check.csv'
        $duplicateSmtpState = Get-SemrInventoryFileState -Path $duplicateSmtpPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        $duplicateRoutingState = Get-SemrInventoryFileState -Path $duplicateRoutingPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        $proxyCheckState = Get-SemrInventoryFileState -Path $proxyCheckPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        if ($duplicateSmtpState.Available) {
            $script:InventoryContext.DuplicateSmtpRows = @(Import-Csv -LiteralPath $duplicateSmtpPath)
            $script:InventoryContext.DuplicateSmtpDataAvailable = $true
            $script:InventoryContext.AdvancedIdentityTimestamp = $duplicateSmtpState.Timestamp
        }
        if ($duplicateRoutingState.Available) {
            $script:InventoryContext.DuplicateRoutingRows = @(Import-SemrFlexibleCsv -Path $duplicateRoutingPath)
            $script:InventoryContext.DuplicateRoutingDataAvailable = $true
            if (-not $script:InventoryContext.AdvancedIdentityTimestamp) { $script:InventoryContext.AdvancedIdentityTimestamp = $duplicateRoutingState.Timestamp }
        }
        if ($proxyCheckState.Available) {
            $script:InventoryContext.ProxyCheckRowsByEmail = Import-SemrInventoryMatches -Path $proxyCheckPath -EmailAddresses $emails -IdentityColumns @('PrimaryAddress')
            $script:InventoryContext.ProxyCheckDataAvailable = $true
            if (-not $script:InventoryContext.AdvancedIdentityTimestamp) { $script:InventoryContext.AdvancedIdentityTimestamp = $proxyCheckState.Timestamp }
        }
    }

    if (Test-SemrCheckEnabled -CheckId 'SMTP-ACCEPTED-DOMAIN') {
        $acceptedDomainPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'Exchange_EXO_AcceptedDomains.csv'
        $acceptedDomainState = Get-SemrInventoryFileState -Path $acceptedDomainPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        if ($acceptedDomainState.Available) {
            $script:InventoryContext.AcceptedDomainRows = @(Import-SemrFlexibleCsv -Path $acceptedDomainPath)
            $script:InventoryContext.AcceptedDomainDataAvailable = $true
            $script:InventoryContext.AcceptedDomainTimestamp = $acceptedDomainState.Timestamp
        }
    }

    if (@(@('ARCHIVE-READINESS','ARCHIVE-GUID-CONSISTENCY') | Where-Object { Test-SemrCheckEnabled -CheckId $_ }).Count -gt 0) {
        $archivePath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'Exchange_EXO_Mailboxes_AllDomains_Archive.csv'
        $archiveState = Get-SemrInventoryFileState -Path $archivePath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        if ($archiveState.Available) {
            $script:InventoryContext.ArchiveRowsByEmail = Import-SemrInventoryMatches -Path $archivePath -EmailAddresses $emails -IdentityColumns @('UserPrincipalName','PrimarySmtpAddress')
            $script:InventoryContext.ArchiveDataAvailable = $true
            $script:InventoryContext.ArchiveTimestamp = $archiveState.Timestamp
        }
    }

    $hybridAdvancedChecks = @('HYBRID-MRSPROXY','HYBRID-CERTIFICATE-EXPIRY','HYBRID-ENDPOINT-CAPACITY','HYBRID-AUTODISCOVER-OAUTH','EXCHANGE-DATABASE-HEALTH')
    if (@($hybridAdvancedChecks | Where-Object { Test-SemrCheckEnabled -CheckId $_ }).Count -gt 0) {
        $hybridConfigPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'Exchange_OnPrem_MigrationReadiness_Config.csv'
        $hybridConfigState = Get-SemrInventoryFileState -Path $hybridConfigPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        if ($hybridConfigState.Available) {
            $script:InventoryContext.HybridConfigRows = @(Import-SemrFlexibleCsv -Path $hybridConfigPath)
            $script:InventoryContext.HybridConfigDataAvailable = $true
            $script:InventoryContext.HybridConfigTimestamp = $hybridConfigState.Timestamp
        }
    }

    if ([string]$Config['Mode'] -eq 'CacheOnly') {
        $cloudUserPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'M365_Users_Active.csv'
        $exoMailboxPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'Exchange_EXO_Mailboxes_AllDomains.csv'
        $migrationPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'Exchange_EXO_MigrationJobs.csv'
        $licensePath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'M365_Licenses_Tenant.csv'
        $licenseServicePlanPath = Resolve-SemrInventoryFilePath -Config $Config -FileName 'M365_Licenses_ServicePlans.csv'
        $cloudState = Get-SemrInventoryFileState -Path $cloudUserPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        $exoState = Get-SemrInventoryFileState -Path $exoMailboxPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale
        $migrationChecksEnabled = (Test-SemrCheckEnabled 'EXO-EXISTING-MOVE') -or (Test-SemrCheckEnabled 'MOVE-HISTORY')
        $licenseCapacityEnabled = Test-SemrCheckEnabled 'LICENSE-CAPACITY'
        $licenseServicePlanEnabled = Test-SemrCheckEnabled 'LICENSE-EXCHANGE-SERVICE-PLAN'
        $migrationState = if($migrationChecksEnabled){Get-SemrInventoryFileState -Path $migrationPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale}else{[pscustomobject]@{Available=$false}}
        $licenseState = if($licenseCapacityEnabled){Get-SemrInventoryFileState -Path $licensePath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale}else{[pscustomobject]@{Available=$false}}
        $licenseServicePlanState = if($licenseServicePlanEnabled){Get-SemrInventoryFileState -Path $licenseServicePlanPath -MaximumAgeHours $maximumAgeHours -AllowStale:$allowStale}else{[pscustomobject]@{Available=$false}}

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
        if ($licenseServicePlanState.Available) {
            $configuredQuotaSkus = @($Config['TargetQuotaGbBySku'].Keys | ForEach-Object { ([string]$_).ToUpperInvariant() })
            $targetSkus = @(
                @($BatchRows | ForEach-Object { [string]$_.TargetSku }) + @([string]$Config['DefaultTargetSku']) |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    ForEach-Object { $_.Trim().ToUpperInvariant() } |
                    Where-Object { $_ -in $configuredQuotaSkus } |
                    Sort-Object -Unique
            )
            $servicePlanRows = [System.Collections.Generic.List[object]]::new()
            foreach ($targetSku in $targetSkus) {
                $exchangePlan = @(
                    Import-Csv -LiteralPath $licenseServicePlanPath |
                        Where-Object {
                            ([string]$_.SkuPartNumber).Trim().ToUpperInvariant() -eq $targetSku -and
                            (Test-SemrMailboxServicePlanName -Name ([string]$_.PlanName)) -and
                            (ConvertTo-SemrBoolean -Value $_.IsEnabled) -and
                            ([string]::IsNullOrWhiteSpace([string]$_.PlanStatus) -or [string]$_.PlanStatus -ieq 'Success')
                        } |
                        Select-Object -First 1
                )
                if ($exchangePlan.Count -eq 1) { [void]$servicePlanRows.Add($exchangePlan[0]) }
            }
            $script:InventoryContext.LicenseServicePlanRows = @($servicePlanRows)
            $script:InventoryContext.LicenseServicePlanDataAvailable = $true
            $script:InventoryContext.LicenseServicePlanTimestamp = $licenseServicePlanState.Timestamp
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
            LegacyExchangeDN = ''
            ExchangeGuid = ''
            ArchiveGuid = ''
            ArchiveStatus = [string]$mailboxRow.ArchiveStatus
            ArchiveState = [string]$mailboxRow.ArchiveState
            ArchiveQuota = [string]$mailboxRow.ArchiveQuota
            ArchiveTotalItemSize = [string]$mailboxRow.'ArchiveTotalItemSize-In-MB'
            ForwardingAddress = [string]$mailboxRow.ForwardingAddress
            ForwardingSmtpAddress = [string]$mailboxRow.ForwardingSmtpAddress
            DeliverToMailboxAndForward = ConvertTo-SemrBoolean -Value $mailboxRow.DeliverToMailboxAndForward
            UseDatabaseQuotaDefaults = $(
                $quotaDefaultRaw = [string](Get-SemrPropertyValue -InputObject $mailboxRow -Names @('UseDatabaseQuotaDefaults') -Default '')
                if ([string]::IsNullOrWhiteSpace($quotaDefaultRaw)) { $null } else { ConvertTo-SemrBoolean -Value $quotaDefaultRaw }
            )
            ProhibitSendReceiveQuota = [string]$mailboxRow.'ProhibitSendReceiveQuota-In-MB'
            Database = [string]$mailboxRow.Database
            ServerName = [string]$mailboxRow.ServerName
            LargeItemCount = [string]$mailboxRow.'LargeItemCount-Over-35MB'
            ModerationEnabled = $null
        })
        $sizeMb = [string](Get-SemrPropertyValue -InputObject $mailboxRow -Names @('TotalItemSize-In-MB') -Default '')
        if ($sizeMb) {
            [void]$statistics.Add([pscustomobject]@{
                TotalItemSize = "$sizeMb MB"
                ItemCount = [string]$mailboxRow.ItemCount
                TotalDeletedItemSize = "$([string]$mailboxRow.'TotalDeletedItemSize-In-MB') MB"
            })
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
            LegacyExchangeDN = ''
            ExchangeGuid = ''
            ArchiveGuid = ''
            ArchiveStatus = [string](Get-SemrPropertyValue -InputObject $adRow[0] -Names @('ArchiveStatus') -Default '')
            ArchiveState = ''
            ArchiveQuota = ''
            ArchiveTotalItemSize = [string](Get-SemrPropertyValue -InputObject $adRow[0] -Names @('ArchiveMailboxSizeGB') -Default '')
            ForwardingAddress = ''
            ForwardingSmtpAddress = ''
            DeliverToMailboxAndForward = $false
            UseDatabaseQuotaDefaults = $null
            ProhibitSendReceiveQuota = ''
            Database = ''
            ServerName = ''
            LargeItemCount = ''
            ModerationEnabled = $null
        })
        $adMailboxSizeGb = [string](Get-SemrPropertyValue -InputObject $adRow[0] -Names @('MailboxSizeGB') -Default '')
        if ($adMailboxSizeGb) {
            [void]$statistics.Add([pscustomobject]@{ TotalItemSize = "$adMailboxSizeGb GB" })
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
        FolderStatistics = @()
        FolderStatisticsAvailable = $false
        InboxRules = @()
        InboxRulesAvailable = $false
        DatabaseHealth = @()
        DatabaseHealthAvailable = $false
        DeliveryRestrictionsAvailable = $false
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
    $folderStatisticsCommand = Get-SemrOnPremCommandName -Name 'Get-MailboxFolderStatistics'
    $inboxRuleCommand = Get-SemrOnPremCommandName -Name 'Get-InboxRule'
    $databaseCommand = Get-SemrOnPremCommandName -Name 'Get-MailboxDatabase'

    $mailboxes = @(Invoke-SemrCommandSafe -CommandName $mailboxCommand -Parameters @{ Identity = $EmailAddress })
    $remoteMailboxes = @(Invoke-SemrCommandSafe -CommandName $remoteMailboxCommand -Parameters @{ Identity = $EmailAddress })
    $mailUsers = @(Invoke-SemrCommandSafe -CommandName $mailUserCommand -Parameters @{ Identity = $EmailAddress })
    $recipients = @(Invoke-SemrCommandSafe -CommandName $recipientCommand -Parameters @{ Identity = $EmailAddress })

    $statistics = @()
    if ($mailboxes.Count -eq 1 -and $statisticsCommand) {
        $statistics = @(Invoke-SemrCommandSafe -CommandName $statisticsCommand -Parameters @{ Identity = $mailboxes[0].Identity })
    }

    $folderStatistics = @()
    $folderStatisticsAvailable = $false
    if ($mailboxes.Count -eq 1 -and $folderStatisticsCommand -and (
        (Test-SemrCheckEnabled -CheckId 'MAILBOX-RECOVERABLE-ITEMS-QUOTA') -or
        (Test-SemrCheckEnabled -CheckId 'MAILBOX-FOLDER-LIMITS')
    )) {
        $folderStatistics = @(Invoke-SemrCommandSafe -CommandName $folderStatisticsCommand -Parameters @{ Identity = $mailboxes[0].Identity })
        $folderStatisticsAvailable = $folderStatistics.Count -gt 0
    }

    $inboxRules = @()
    $inboxRulesAvailable = $false
    if ($mailboxes.Count -eq 1 -and $inboxRuleCommand -and (Test-SemrCheckEnabled -CheckId 'INBOX-FORWARDING-RULES')) {
        $inboxRules = @(Invoke-SemrCommandSafe -CommandName $inboxRuleCommand -Parameters @{ Mailbox = $mailboxes[0].Identity; IncludeHidden = $true })
        $inboxRulesAvailable = $true
    }

    $databaseHealth = @()
    $databaseHealthAvailable = $false
    if ($mailboxes.Count -eq 1 -and $databaseCommand -and (Test-SemrCheckEnabled -CheckId 'EXCHANGE-DATABASE-HEALTH')) {
        $databaseIdentity = [string](Get-SemrPropertyValue -InputObject $mailboxes[0] -Names @('Database') -Default '')
        if ($databaseIdentity) {
            $databaseHealth = @(Invoke-SemrCommandSafe -CommandName $databaseCommand -Parameters @{ Identity = $databaseIdentity; Status = $true })
            $databaseHealthAvailable = $databaseHealth.Count -gt 0
        }
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
        FolderStatistics = $folderStatistics
        FolderStatisticsAvailable = $folderStatisticsAvailable
        InboxRules = $inboxRules
        InboxRulesAvailable = $inboxRulesAvailable
        DatabaseHealth = $databaseHealth
        DatabaseHealthAvailable = $databaseHealthAvailable
        DeliveryRestrictionsAvailable = $true
    }
}

function Get-SemrActiveDirectoryEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    $getCachedEvidence = {
        param([string]$FallbackReason = '')
        if ($script:InventoryContext) {
            $key = $EmailAddress.Trim().ToLowerInvariant()
            $users = @()
            if ($script:InventoryContext.AdRowsByEmail.ContainsKey($key)) { $users = @($script:InventoryContext.AdRowsByEmail[$key] | ForEach-Object { $_ }) }
            $message = [string]$script:InventoryContext.ActiveDirectoryMessage
            if ($FallbackReason) {
                $script:ActiveDirectoryFallbackUsed = $true
                if ($script:ActiveDirectoryFallbackReasons) { [void]$script:ActiveDirectoryFallbackReasons.Add($FallbackReason) }
                $message = "$FallbackReason CSV fallback: $message"
            }
            return [pscustomobject]@{
                Available = [bool]$script:InventoryContext.ActiveDirectoryAvailable
                Source = if ($FallbackReason) { 'Tenant CSV cache AD inventory (live forest fallback)' } else { 'Tenant CSV cache AD inventory' }
                SourceTimestamp = $script:InventoryContext.AdTimestamp
                Message = $message
                Users = $users
            }
        }
        return [pscustomobject]@{
            Available = $false
            Source = 'Active Directory'
            SourceTimestamp = $null
            Message = if ($FallbackReason) { $FallbackReason } else { 'Active Directory is not connected.' }
            Users = @()
        }
    }

    if (($script:InventoryContext -and $script:InventoryContext.Mode -eq 'CacheOnly') -or -not $script:ConnectionState.ActiveDirectory -or -not (Test-SemrCommand -Name 'Get-ADUser')) {
        return & $getCachedEvidence
    }

    $escaped = $EmailAddress.Replace("'", "''")
    $domains = @($script:ActiveDirectoryDomains)
    if ($domains.Count -eq 0) {
        try {
            $domains = @(
                (Get-ADForest -ErrorAction Stop).Domains |
                    ForEach-Object { ([string]$_).Trim() } |
                    Where-Object { $_ } |
                    Sort-Object -Unique
            )
            $script:ActiveDirectoryDomains = $domains
        }
        catch {
            return & $getCachedEvidence "Live Active Directory forest discovery failed: $($_.Exception.Message)."
        }
    }

    $users = [System.Collections.Generic.List[object]]::new()
    $domainErrors = [System.Collections.Generic.List[string]]::new()
    foreach ($domain in $domains) {
        try {
            $queryParameters = @{
                Server = $domain
                Filter = "UserPrincipalName -eq '$escaped' -or mail -eq '$escaped' -or proxyAddresses -eq 'smtp:$escaped'"
                Properties = @('Enabled','UserPrincipalName','mail','proxyAddresses','targetAddress','msDS-ConsistencyGuid','ObjectGuid','whenChanged')
                ErrorAction = 'Stop'
            }
            $domainUsers = @(Get-ADUser @queryParameters)
            foreach ($domainUser in $domainUsers) { [void]$users.Add($domainUser) }
        }
        catch {
            [void]$domainErrors.Add("${domain}: $($_.Exception.Message)")
        }
    }
    if ($domainErrors.Count -gt 0) {
        return & $getCachedEvidence "Live Active Directory forest search was incomplete ($($domainErrors.Count)/$($domains.Count) domain(s) failed: $($domainErrors -join '; '))."
    }

    $uniqueUsers = @(
        $users |
            Group-Object {
                $objectGuid = [string](Get-SemrPropertyValue -InputObject $_ -Names @('ObjectGuid') -Default '')
                if ($objectGuid) { $objectGuid } else { [string](Get-SemrPropertyValue -InputObject $_ -Names @('DistinguishedName') -Default $_.UserPrincipalName) }
            } |
            ForEach-Object { $_.Group[0] }
    )
    return [pscustomobject]@{
        Available = $true
        Source = "Live Active Directory forest ($($domains.Count) domains)"
        SourceTimestamp = Get-Date
        Message = "Live Active Directory evidence collected across $($domains.Count) forest domain(s)."
        Users = $uniqueUsers
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
        RecipientDataAvailable = [bool]$script:InventoryContext.CloudUsersAvailable
        RecipientLookupMethods = @('CacheOnly CSV inventories')
        RecipientLookupErrors = @()
        MailboxDataAvailable = [bool]$script:InventoryContext.ExchangeOnlineAvailable
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
                OnPremisesProvisioningErrors = @(Split-SemrInventoryValue -Value (Get-SemrPropertyValue -InputObject $_ -Names @('OnPremisesProvisioningErrors','Provisioning errors') -Default ''))
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
            Available = $false; Source = 'Exchange Online unavailable'; SourceTimestamp = $null
            Recipients = @(); RecipientDataAvailable = $false; RecipientLookupMethods = @(); RecipientLookupErrors = @()
            Mailboxes = @(); MailboxDataAvailable = $false; Statistics = @()
            SoftDeleted = @(); SoftDeletedAvailable = $false
            MigrationUsers = @(); MoveRequests = @(); MoveDataAvailable = $false
        }
    }

    $collectedAt = Get-Date
    $lookupErrors = [Collections.Generic.List[string]]::new()
    $lookupMethods = [Collections.Generic.List[string]]::new()
    $recipientList = [Collections.Generic.List[object]]::new()
    $recipientKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $addRecipients = {
        param([object[]]$Items, [string]$Method)
        if (-not $lookupMethods.Contains($Method)) { [void]$lookupMethods.Add($Method) }
        foreach ($item in @($Items)) {
            $key = [string](Get-SemrPropertyValue -InputObject $item -Names @('ExternalDirectoryObjectId','Guid','Identity','PrimarySmtpAddress') -Default '')
            if (-not $key) { $key = [string]$item }
            if ($recipientKeys.Add($key)) { [void]$recipientList.Add($item) }
        }
    }
    $escapedAddress = $EmailAddress.Replace("'", "''")
    $recipientProperties = @('RecipientType','RecipientTypeDetails','EmailAddresses','ExternalEmailAddress','ExchangeGuid','ArchiveGuid','ExternalDirectoryObjectId')
    $recipientDataAvailable = $false

    if (Test-SemrCommand -Name 'Get-EXORecipient') {
        try {
            $items = @(Get-EXORecipient -Filter "EmailAddresses -eq 'smtp:$escapedAddress'" -ResultSize Unlimited -Properties $recipientProperties -ErrorAction Stop)
            $recipientDataAvailable = $true
            & $addRecipients $items 'Get-EXORecipient -Filter'
        }
        catch { [void]$lookupErrors.Add("Get-EXORecipient filter: $($_.Exception.Message)") }
        if ($recipientList.Count -eq 0) {
            try {
                $items = @(Get-EXORecipient -Identity $EmailAddress -Properties $recipientProperties -ErrorAction Stop)
                $recipientDataAvailable = $true
                & $addRecipients $items 'Get-EXORecipient -Identity'
            }
            catch {
                if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $recipientDataAvailable = $true }
                else { [void]$lookupErrors.Add("Get-EXORecipient identity: $($_.Exception.Message)") }
            }
        }
    }
    if ($recipientList.Count -eq 0 -and (Test-SemrCommand -Name 'Get-Recipient')) {
        try {
            $items = @(Get-Recipient -Identity $EmailAddress -ErrorAction Stop)
            $recipientDataAvailable = $true
            & $addRecipients $items 'Get-Recipient -Identity'
        }
        catch {
            if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $recipientDataAvailable = $true }
            else { [void]$lookupErrors.Add("Get-Recipient identity: $($_.Exception.Message)") }
        }
    }
    if ($recipientList.Count -eq 0 -and (Test-SemrCommand -Name 'Get-MailUser')) {
        try {
            $items = @(Get-MailUser -Identity $EmailAddress -ErrorAction Stop)
            $recipientDataAvailable = $true
            & $addRecipients $items 'Get-MailUser -Identity'
        }
        catch {
            if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $recipientDataAvailable = $true }
            else { [void]$lookupErrors.Add("Get-MailUser identity: $($_.Exception.Message)") }
        }
    }

    $mailboxes = @()
    $mailboxDataAvailable = $false
    if (Test-SemrCommand -Name 'Get-EXOMailbox') {
        try {
            $mailboxes = @(Get-EXOMailbox -Identity $EmailAddress -Properties @('RecipientTypeDetails','EmailAddresses','ExchangeGuid','ArchiveGuid','LitigationHoldEnabled','InPlaceHolds','DelayHoldApplied','ProhibitSendReceiveQuota','GrantSendOnBehalfTo','ArchiveStatus','ArchiveQuota','ForwardingAddress','ForwardingSmtpAddress','DeliverToMailboxAndForward','ModerationEnabled') -ErrorAction Stop)
            $mailboxDataAvailable = $true
        }
        catch {
            if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $mailboxDataAvailable = $true }
            else { [void]$lookupErrors.Add("Get-EXOMailbox identity: $($_.Exception.Message)") }
        }
    }
    if (-not $mailboxDataAvailable -and (Test-SemrCommand -Name 'Get-Mailbox')) {
        try { $mailboxes = @(Get-Mailbox -Identity $EmailAddress -ErrorAction Stop); $mailboxDataAvailable = $true }
        catch {
            if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $mailboxDataAvailable = $true }
            else { [void]$lookupErrors.Add("Get-Mailbox identity: $($_.Exception.Message)") }
        }
    }
    if (-not $mailboxDataAvailable -and $recipientList.Count -eq 1) {
        $recipientType = [string](Get-SemrPropertyValue -InputObject $recipientList[0] -Names @('RecipientTypeDetails','RecipientType') -Default '')
        if ($recipientType -match 'MailUser|RemoteUserMailbox') { $mailboxDataAvailable = $true }
    }

    $statistics = @(if ($mailboxes.Count -eq 1 -and (Test-SemrCommand -Name 'Get-EXOMailboxStatistics')) {
        Invoke-SemrCommandSafe -CommandName 'Get-EXOMailboxStatistics' -Parameters @{ Identity = $EmailAddress; Properties = @('TotalItemSize','ItemCount','TotalDeletedItemSize','StorageLimitStatus') }
    })
    $softDeleted = @(); $softDeletedAvailable = $false
    if (Test-SemrCommand -Name 'Get-Mailbox') {
        try { $softDeleted = @(Get-Mailbox -Identity $EmailAddress -SoftDeletedMailbox -IncludeInactiveMailbox -ErrorAction Stop); $softDeletedAvailable = $true }
        catch {
            if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $softDeletedAvailable = $true }
            else { [void]$lookupErrors.Add("Soft-deleted mailbox lookup: $($_.Exception.Message)") }
        }
    }

    $migrationUsers = @(); $moveRequests = @(); $migrationLookupAvailable = $false; $moveRequestLookupAvailable = $false
    if (Test-SemrCommand -Name 'Get-MigrationUser') {
        try { $migrationUsers = @(Get-MigrationUser -Identity $EmailAddress -ErrorAction Stop); $migrationLookupAvailable = $true }
        catch {
            if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $migrationLookupAvailable = $true }
            else { [void]$lookupErrors.Add("Get-MigrationUser identity: $($_.Exception.Message)") }
        }
    }
    if (Test-SemrCommand -Name 'Get-MoveRequest') {
        try { $moveRequests = @(Get-MoveRequest -Identity $EmailAddress -ErrorAction Stop); $moveRequestLookupAvailable = $true }
        catch {
            if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $moveRequestLookupAvailable = $true }
            else { [void]$lookupErrors.Add("Get-MoveRequest identity: $($_.Exception.Message)") }
        }
    }

    return [pscustomobject]@{
        Available = $true; Source = 'Live Exchange Online'; SourceTimestamp = $collectedAt
        Recipients = @($recipientList); RecipientDataAvailable = $recipientDataAvailable
        RecipientLookupMethods = @($lookupMethods); RecipientLookupErrors = @($lookupErrors)
        Mailboxes = $mailboxes; MailboxDataAvailable = $mailboxDataAvailable; Statistics = $statistics
        SoftDeleted = $softDeleted; SoftDeletedAvailable = $softDeletedAvailable
        MigrationUsers = $migrationUsers; MoveRequests = $moveRequests
        MoveDataAvailable = $migrationLookupAvailable -and $moveRequestLookupAvailable
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

    $key = $EmailAddress.Trim().ToLowerInvariant()
    if (-not $script:GraphEvidenceByEmail.ContainsKey($key)) {
        return [pscustomobject]@{
            Available = $true
            Source = 'Live Microsoft Graph (isolated process)'
            SourceTimestamp = Get-Date
            Users = @()
            LicenseDetails = @()
            QueryError = ''
        }
    }

    $entry = $script:GraphEvidenceByEmail[$key]
    return [pscustomobject]@{
        Available = $true
        Source = 'Live Microsoft Graph (isolated process)'
        SourceTimestamp = $entry.SourceTimestamp
        Users = @($entry.Users)
        LicenseDetails = @($entry.LicenseDetails)
        QueryError = [string]$entry.QueryError
    }
}
function Test-SemrMailboxServicePlanName {
    param([string]$Name)

    return -not [string]::IsNullOrWhiteSpace($Name) -and $Name -match '^(?i:EXCHANGE_S_(ENTERPRISE|STANDARD|DESKLESS)|EXCHANGE_(ENTERPRISE|STANDARD|DESKLESS))$'
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
        ServicePlanDataAvailable = $false
        ExchangeServicePlanFound = $false
        ExchangeServicePlanName = ''
        ExchangeServicePlanStatus = ''
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
        if ($script:InventoryContext.LicenseServicePlanDataAvailable) {
            $result.ServicePlanDataAvailable = $true
            $exchangePlan = @(
                $script:InventoryContext.LicenseServicePlanRows |
                    Where-Object {
                        [string]$_.SkuPartNumber -ieq $TargetSku -and
                        (Test-SemrMailboxServicePlanName -Name ([string]$_.PlanName)) -and
                        (ConvertTo-SemrBoolean -Value $_.IsEnabled) -and
                        ([string]::IsNullOrWhiteSpace([string]$_.PlanStatus) -or [string]$_.PlanStatus -ieq 'Success')
                    } |
                    Select-Object -First 1
            )
            if ($exchangePlan.Count -eq 1) {
                $result.ExchangeServicePlanFound = $true
                $result.ExchangeServicePlanName = [string]$exchangePlan[0].PlanName
                $result.ExchangeServicePlanStatus = [string]$exchangePlan[0].PlanStatus
            }
        }
        $result.Message = "Cached: Enabled=$($result.Enabled); Consumed=$($result.Consumed); Available=$($result.AvailableUnits)"
        return [pscustomobject]$result
    }

    if (-not $script:ConnectionState.MicrosoftGraph) {
        $result.Message = 'Microsoft Graph subscribed SKU data is unavailable.'
        return [pscustomobject]$result
    }
    if ($script:GraphSubscribedSkuError) {
        $result.Message = $script:GraphSubscribedSkuError
        return [pscustomobject]$result
    }

    $result.Available = $true
    $sku = @($script:GraphSubscribedSkus | Where-Object { $_.SkuPartNumber -ieq $TargetSku } | Select-Object -First 1)
    if ($sku.Count -eq 0) {
        $result.Message = "Target SKU '$TargetSku' was not found in subscribed SKUs."
        return [pscustomobject]$result
    }
    $result.Found = $true
    $result.Enabled = [int]$sku[0].Enabled
    $result.Consumed = [int]$sku[0].Consumed
    $result.AvailableUnits = [math]::Max(0, $result.Enabled - $result.Consumed)
    if ($sku[0].PSObject.Properties['ServicePlans']) {
        $result.ServicePlanDataAvailable = $true
        $exchangePlan = @(
            @($sku[0].ServicePlans) |
                Where-Object {
                    (Test-SemrMailboxServicePlanName -Name ([string]$_.ServicePlanName)) -and
                    ([string]::IsNullOrWhiteSpace([string]$_.ProvisioningStatus) -or [string]$_.ProvisioningStatus -ieq 'Success')
                } |
                Select-Object -First 1
        )
        if ($exchangePlan.Count -eq 1) {
            $result.ExchangeServicePlanFound = $true
            $result.ExchangeServicePlanName = [string]$exchangePlan[0].ServicePlanName
            $result.ExchangeServicePlanStatus = [string]$exchangePlan[0].ProvisioningStatus
        }
    }
    $result.Message = "Enabled=$($result.Enabled); Consumed=$($result.Consumed); Available=$($result.AvailableUnits)"
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

function Get-SemrTargetQuotaEvidence {
    param(
        [string]$TargetSku,
        [System.Collections.IDictionary]$Config,
        [string]$MailboxType,
        [switch]$TargetSkuExplicit
    )

    $result = [ordered]@{
        TargetSku = $TargetSku
        Known = $false
        Eligible = $false
        RequiresLicense = $true
        QuotaGb = 0.0
        Message = ''
    }
    if ($MailboxType -match 'Shared' -and -not $TargetSkuExplicit) {
        $result.TargetSku = 'UNLICENSED_SHARED_MAILBOX'
        $result.Known = $true
        $result.Eligible = $true
        $result.RequiresLicense = $false
        $result.QuotaGb = 50.0
        $result.Message = 'Unlicensed shared mailbox policy: 50 GB maximum.'
        return [pscustomobject]$result
    }
    if ([string]::IsNullOrWhiteSpace($TargetSku)) {
        $result.Message = 'No target SKU is configured for this mailbox.'
        return [pscustomobject]$result
    }

    $ineligibleSkus = @($Config['MailboxIneligibleTargetSkus'] | ForEach-Object { [string]$_ })
    if (@($ineligibleSkus | Where-Object { $_ -ieq $TargetSku }).Count -gt 0) {
        $result.Known = $true
        $result.Message = "Target SKU '$TargetSku' does not grant an Exchange Online user mailbox."
        return [pscustomobject]$result
    }

    $map = $Config['TargetQuotaGbBySku']
    if ($map -is [System.Collections.IDictionary] -and $TargetSku) {
        foreach ($key in $map.Keys) {
            if ([string]$key -ieq $TargetSku) {
                $result.Known = $true
                $result.Eligible = $true
                $result.QuotaGb = [double]$map[$key]
                $result.Message = "Configured primary mailbox quota for '$TargetSku': $($result.QuotaGb) GB."
                return [pscustomobject]$result
            }
        }
    }
    $result.Message = "Target SKU '$TargetSku' is not present in TargetQuotaGbBySku; quota eligibility cannot be assumed."
    return [pscustomobject]$result
}

function Get-SemrAcceptedDomainEvidence {
    $result = [ordered]@{ Available = $false; Source = 'Unavailable'; SourceTimestamp = $null; Domains = @(); Message = 'Accepted-domain evidence is unavailable.' }
    if (-not (Test-SemrCheckEnabled -CheckId 'SMTP-ACCEPTED-DOMAIN')) { return [pscustomobject]$result }
    if ($script:ConnectionState.ExchangeOnline -and (Test-SemrCommand -Name 'Get-AcceptedDomain')) {
        try {
            $rows = @(Get-AcceptedDomain -ErrorAction Stop)
            $result.Available = $true
            $result.Source = 'Live Exchange Online accepted domains'
            $result.SourceTimestamp = Get-Date
            $result.Domains = @($rows | ForEach-Object { [string](Get-SemrPropertyValue -InputObject $_ -Names @('DomainName','Name') -Default '') } | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
            $result.Message = "$($result.Domains.Count) accepted domain(s) collected live."
            return [pscustomobject]$result
        }
        catch { $result.Message = "Live accepted-domain collection failed: $($_.Exception.Message)" }
    }
    if ($script:InventoryContext -and $script:InventoryContext.AcceptedDomainDataAvailable) {
        $result.Available = $true
        $result.Source = 'Tenant CSV cache accepted domains'
        $result.SourceTimestamp = $script:InventoryContext.AcceptedDomainTimestamp
        $result.Domains = @($script:InventoryContext.AcceptedDomainRows | ForEach-Object { [string](Get-SemrPropertyValue -InputObject $_ -Names @('DomainName','Name') -Default '') } | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
        $result.Message = "$($result.Domains.Count) accepted domain(s) collected from CSV cache."
    }
    return [pscustomobject]$result
}

function Get-SemrProxyConflictEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress, [Parameter(Mandatory)][string[]]$Addresses)

    $conflicts = [System.Collections.Generic.List[string]]::new()
    $available = $false
    $source = 'Unavailable'
    if (-not (Test-SemrCheckEnabled -CheckId 'PROXY-SMTP-GLOBAL-UNIQUE')) { return [pscustomobject]@{ Available=$false; Source='Disabled'; Conflicts=@() } }
    if ($script:InventoryContext -and $script:InventoryContext.DuplicateSmtpDataAvailable) {
        $available = $true
        $source = 'AD duplicate SMTP CSV cache'
        $normalized = @($Addresses | ForEach-Object { ([string]$_ -replace '^(?i:smtp:)','').Trim().ToLowerInvariant() } | Where-Object { $_ })
        foreach ($row in @($script:InventoryContext.DuplicateSmtpRows | Where-Object { $normalized -contains ([string]$_.SmtpAddress).ToLowerInvariant() })) {
            $owner = [string](Get-SemrPropertyValue -InputObject $row -Names @('UserPrincipalName','DistinguishedName') -Default 'unknown owner')
            if ($owner -and $owner -ine $EmailAddress) { [void]$conflicts.Add("$([string]$row.SmtpAddress) -> $owner") }
        }
    }
    if ($script:InventoryContext -and $script:InventoryContext.ProxyCheckDataAvailable) {
        $available = $true
        if ($source -eq 'Unavailable') { $source = 'Exchange proxy-address CSV cache' }
        $key = $EmailAddress.ToLowerInvariant()
        $proxyRows = if ($script:InventoryContext.ProxyCheckRowsByEmail.ContainsKey($key)) { @($script:InventoryContext.ProxyCheckRowsByEmail[$key]) } else { @() }
        foreach ($row in $proxyRows) {
            if (ConvertTo-SemrBoolean -Value (Get-SemrPropertyValue -InputObject $row -Names @('ExpectedAddressConflict') -Default $false)) {
                $address = [string](Get-SemrPropertyValue -InputObject $row -Names @('ExpectedAddress') -Default '')
                $owners = [string](Get-SemrPropertyValue -InputObject $row -Names @('ExpectedAddressConflictOwners','ExpectedAddressDuplicatePeers') -Default '')
                [void]$conflicts.Add("$address -> $owners")
            }
        }
    }
    if (-not $available -and $script:ConnectionState.OnPremisesExchange) {
        $recipientCommand = Get-SemrOnPremCommandName -Name 'Get-Recipient'
        if ($recipientCommand) {
            $available = $true
            $source = 'Live Exchange on-premises recipient directory'
            foreach ($address in @($Addresses | ForEach-Object { ([string]$_ -replace '^(?i:smtp:)','').Trim() } | Where-Object { Test-SemrSmtpAddress $_ } | Sort-Object -Unique)) {
                $safeAddress = $address.Replace("'", "''")
                foreach ($owner in @(Invoke-SemrCommandSafe -CommandName $recipientCommand -Parameters @{ Filter = "EmailAddresses -eq 'smtp:$safeAddress'"; ResultSize = 'Unlimited' })) {
                    $ownerAddress = [string](Get-SemrPropertyValue -InputObject $owner -Names @('PrimarySmtpAddress','WindowsEmailAddress') -Default '')
                    if ($ownerAddress -and $ownerAddress -ine $EmailAddress) { [void]$conflicts.Add("$address -> $ownerAddress") }
                }
            }
        }
    }
    return [pscustomobject]@{ Available = $available; Source = $source; Conflicts = @($conflicts | Sort-Object -Unique) }
}

function Add-SemrIdentityAdvancedFindings {
    param([System.Collections.Generic.List[object]]$Findings,[hashtable]$Base,$OnPrem,$Exo,$AcceptedDomains)
    $email = [string]$Base.EmailAddress
    $source = [string](Get-SemrPropertyValue $OnPrem @('Source') 'Exchange on-premises')
    $cloudSource = [string](Get-SemrPropertyValue $Exo @('Source') 'Exchange Online')
    $mailbox = @($OnPrem.Mailboxes | Select-Object -First 1)
    if ($mailbox.Count -ne 1) { return }

    $raw = @(Get-SemrPropertyValue $mailbox[0] @('EmailAddresses') @())
    $smtp = @($raw | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^(?i:smtp:)' })
    $normalized = @($smtp | ForEach-Object { ($_ -replace '^(?i:smtp:)','').Trim().ToLowerInvariant() })
    $invalid = @($normalized | Where-Object { -not (Test-SemrSmtpAddress $_) })
    $duplicates = @($normalized | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name)
    $pass = $invalid.Count -eq 0 -and $duplicates.Count -eq 0
    Add-SemrFinding $Findings ($Base + @{CheckId='PROXY-INTERNAL-DUPLICATE';Category='HybridIdentity';Severity=if($pass){'Information'}else{'Critical'};Result=if($pass){'PASS'}else{'FAIL'};IsBlocking=-not $pass;ObservedValue=if($pass){"$($normalized.Count) unique SMTP proxy address(es)"}else{"Invalid=$($invalid -join ';'); Duplicate=$($duplicates -join ';')"};ExpectedValue='Unique, valid SMTP proxy addresses';EvidenceSource=$source;Message=if($pass){'No internal SMTP proxy duplicate or malformed address was found.'}else{'A malformed or duplicated SMTP proxy address was found.'};RecommendedAction=if($pass){''}else{'Correct proxyAddresses with supported Exchange tools before migration.'}})

    $conflict = Get-SemrProxyConflictEvidence -EmailAddress $email -Addresses $smtp
    if ($conflict.Source -and $script:InventoryContext -and $script:InventoryContext.AdvancedIdentityTimestamp) { $script:CurrentSourceTimestamps[[string]$conflict.Source] = $script:InventoryContext.AdvancedIdentityTimestamp }
    $conflicts = @($conflict.Conflicts)
    Add-SemrFinding $Findings ($Base + @{CheckId='PROXY-SMTP-GLOBAL-UNIQUE';Category='HybridIdentity';Severity=if(-not $conflict.Available){'Warning'}elseif($conflicts.Count){'Critical'}else{'Information'};Result=if(-not $conflict.Available){'UNKNOWN'}elseif($conflicts.Count){'FAIL'}else{'PASS'};IsBlocking=$conflict.Available -and $conflicts.Count -gt 0;ObservedValue=if($conflict.Available){if($conflicts.Count){$conflicts -join ' | '}else{'No conflicting owner'}}else{'Global recipient index unavailable'};ExpectedValue='One owner per SMTP proxy';EvidenceSource=$conflict.Source;Message=if(-not $conflict.Available){'Global SMTP ownership could not be evaluated.'}elseif($conflicts.Count){'At least one SMTP proxy is owned by another recipient.'}else{'No conflicting SMTP owner was found.'};RecommendedAction=if(-not $conflict.Available){'Refresh duplicate/proxy inventories or validate against live Exchange.'}elseif($conflicts.Count){'Resolve every duplicate SMTP ownership conflict.'}else{''}})

    $target = ([string](Get-SemrPropertyValue $mailbox[0] @('ExternalEmailAddress','WindowsEmailAddress') '') -replace '^(?i:smtp:)','').Trim().ToLowerInvariant()
    $targetEvidence = -not $target -or ($script:InventoryContext -and ($script:InventoryContext.DuplicateRoutingDataAvailable -or $script:InventoryContext.ProxyCheckDataAvailable))
    $targetMatches = @(if($target -and $script:InventoryContext -and $script:InventoryContext.DuplicateRoutingDataAvailable){$script:InventoryContext.DuplicateRoutingRows | Where-Object { (([string](Get-SemrPropertyValue $_ @('RemoteRoutingAddress','TargetAddress','SmtpAddress') '')) -replace '^(?i:smtp:)','').Trim().ToLowerInvariant() -eq $target }})
    $targetConflict = $target -and $targetMatches.Count -gt 1
    if ($targetEvidence -and $script:InventoryContext -and $script:InventoryContext.AdvancedIdentityTimestamp) { $script:CurrentSourceTimestamps['Hybrid identity evidence'] = $script:InventoryContext.AdvancedIdentityTimestamp }
    Add-SemrFinding $Findings ($Base + @{CheckId='TARGET-ADDRESS-GLOBAL-UNIQUE';Category='HybridIdentity';Severity=if(-not $targetEvidence){'Warning'}elseif($targetConflict){'Critical'}else{'Information'};Result=if(-not $targetEvidence){'UNKNOWN'}elseif($targetConflict){'FAIL'}else{'PASS'};IsBlocking=[bool]$targetConflict;ObservedValue=if($target){"$target; owners=$($targetMatches.Count)"}else{'No targetAddress on source UserMailbox'};ExpectedValue='Empty before onboarding or one unique routing address';EvidenceSource=if($targetEvidence){'Hybrid identity evidence'}else{'Unavailable'};Message=if(-not $targetEvidence){'Global targetAddress uniqueness could not be evaluated.'}elseif($targetConflict){'The target routing address is duplicated.'}else{'No duplicated target routing address was detected.'};RecommendedAction=if($targetConflict){'Resolve the duplicate remote routing address.'}elseif(-not $targetEvidence){'Refresh the duplicate routing and proxy inventories.'}else{''}})

    $legacyDn = [string](Get-SemrPropertyValue $mailbox[0] @('LegacyExchangeDN') '')
    $x500 = $legacyDn -and @($raw | Where-Object { (([string]$_) -replace '^(?i:x500:)','') -ieq $legacyDn }).Count -gt 0
    Add-SemrFinding $Findings ($Base + @{CheckId='X500-LEGACYEXCHANGEDN';Category='HybridIdentity';Severity=if(-not $legacyDn){'Warning'}elseif($x500){'Information'}else{'Critical'};Result=if(-not $legacyDn){'UNKNOWN'}elseif($x500){'PASS'}else{'FAIL'};IsBlocking=[bool]($legacyDn -and -not $x500);ObservedValue=if($legacyDn){"LegacyExchangeDN=$legacyDn; X500Present=$x500"}else{'LegacyExchangeDN unavailable'};ExpectedValue='LegacyExchangeDN preserved as X500';EvidenceSource=$source;Message=if(-not $legacyDn){'LegacyExchangeDN is unavailable.'}elseif($x500){'LegacyExchangeDN is preserved in proxyAddresses.'}else{'LegacyExchangeDN is not preserved as an X500 proxy.'};RecommendedAction=if(-not $legacyDn){'Validate against live Exchange or enrich the cache.'}elseif(-not $x500){'Add the exact legacyExchangeDN as an X500 proxy.'}else{''}})

    $domains = @($normalized | ForEach-Object { ($_ -split '@',2)[1] } | Where-Object { $_ } | Sort-Object -Unique)
    $unaccepted = @(if($AcceptedDomains.Available){$domains | Where-Object { $AcceptedDomains.Domains -notcontains $_ }})
    Add-SemrFinding $Findings ($Base + @{CheckId='SMTP-ACCEPTED-DOMAIN';Category='HybridIdentity';Severity=if(-not $AcceptedDomains.Available){'Warning'}elseif($unaccepted.Count){'Critical'}else{'Information'};Result=if(-not $AcceptedDomains.Available){'UNKNOWN'}elseif($unaccepted.Count){'FAIL'}else{'PASS'};IsBlocking=$AcceptedDomains.Available -and $unaccepted.Count -gt 0;ObservedValue=if($AcceptedDomains.Available){if($unaccepted.Count){'Not accepted: '+($unaccepted -join ';')}else{$domains -join ';'}}else{$AcceptedDomains.Message};ExpectedValue='Every SMTP domain accepted in Exchange Online';EvidenceSource=$AcceptedDomains.Source;Message=if(-not $AcceptedDomains.Available){'Accepted domains could not be evaluated.'}elseif($unaccepted.Count){'One or more SMTP domains are not accepted.'}else{'Every SMTP proxy domain is accepted.'};RecommendedAction=if($unaccepted.Count){'Add/verify accepted domains or remove invalid proxies.'}elseif(-not $AcceptedDomains.Available){'Refresh the accepted-domain inventory or rerun in Live mode.'}else{''}})

    $type = [string](Get-SemrPropertyValue $mailbox[0] @('RecipientTypeDetails') '')
    $supported = @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','LinkedMailbox')
    $typePass = $supported -contains $type
    Add-SemrFinding $Findings ($Base + @{CheckId='RECIPIENT-TYPE-SUPPORTED';Category='ExchangeOnPrem';Severity=if($typePass){'Information'}else{'Critical'};Result=if($typePass){'PASS'}else{'FAIL'};IsBlocking=-not $typePass;ObservedValue=$type;ExpectedValue=$supported -join ', ';EvidenceSource=$source;Message=if($typePass){'The source recipient type is supported.'}else{'The source recipient type is unsupported or ambiguous.'};RecommendedAction=if($typePass){''}else{'Exclude system mailboxes or validate another migration method.'}})

    $cloudRecipient = @($Exo.Recipients | Select-Object -First 1)
    foreach($definition in @(@('EXCHANGE-GUID-CONSISTENCY','ExchangeGuid','Exchange GUID'),@('ARCHIVE-GUID-CONSISTENCY','ArchiveGuid','Archive GUID'))){
        $checkId=$definition[0];$property=$definition[1];$label=$definition[2]
        $sourceGuid=[string](Get-SemrPropertyValue $mailbox[0] @($property) '')
        $cloudGuid=if($cloudRecipient.Count){[string](Get-SemrPropertyValue $cloudRecipient[0] @($property) '')}else{''}
        $archiveNotRequired=$checkId -eq 'ARCHIVE-GUID-CONSISTENCY' -and [string]::IsNullOrWhiteSpace([string](Get-SemrPropertyValue $mailbox[0] @('ArchiveStatus','ArchiveState') '')) -and -not $sourceGuid
        $available=$archiveNotRequired -or ($sourceGuid -and $cloudGuid);$guidPass=$archiveNotRequired -or ($available -and $sourceGuid -ieq $cloudGuid)
        Add-SemrFinding $Findings ($Base + @{CheckId=$checkId;Category='ExchangeIdentity';Severity=if(-not $available){'Warning'}elseif($guidPass){'Information'}else{'Critical'};Result=if(-not $available){'UNKNOWN'}elseif($guidPass){'PASS'}else{'FAIL'};IsBlocking=$available -and -not $guidPass;ObservedValue=if($archiveNotRequired){'No active source archive'}else{"Source=$sourceGuid; Cloud=$cloudGuid"};ExpectedValue="Matching $label";EvidenceSource="$source + $cloudSource";Message=if(-not $available){"$label comparison evidence is incomplete."}elseif($guidPass){"$label state is consistent."}else{"Source and cloud $label values differ."};RecommendedAction=if(-not $available){'Run live validation or enrich GUID inventories.'}elseif(-not $guidPass){'Resolve GUID ownership and soft-deleted conflicts.'}else{''}})
    }
}

function Add-SemrMailboxRiskFindings {
    param([System.Collections.Generic.List[object]]$Findings,[hashtable]$Base,$Row,$OnPrem,[System.Collections.IDictionary]$Config)
    $email=[string]$Base.EmailAddress
    $source=[string](Get-SemrPropertyValue $OnPrem @('Source') 'Exchange on-premises')
    $mailbox=@($OnPrem.Mailboxes | Select-Object -First 1)
    if($mailbox.Count -ne 1){return}

    $archiveStatus=[string](Get-SemrPropertyValue $mailbox[0] @('ArchiveStatus','ArchiveState') '')
    $archiveGuid=[string](Get-SemrPropertyValue $mailbox[0] @('ArchiveGuid') '')
    $hasArchive=$archiveStatus -match 'Active|Local|Hosted' -or ($archiveGuid -and $archiveGuid -notmatch '^0{8}-')
    $archiveSize=ConvertFrom-SemrByteSize (Get-SemrPropertyValue $mailbox[0] @('ArchiveTotalItemSize') '')
    if($archiveSize -eq 0 -and $script:InventoryContext -and $script:InventoryContext.ArchiveDataAvailable -and $script:InventoryContext.ArchiveRowsByEmail.ContainsKey($email.ToLowerInvariant())){
        $archiveRow=@($script:InventoryContext.ArchiveRowsByEmail[$email.ToLowerInvariant()] | Select-Object -First 1)
        if($archiveRow.Count){$archiveSize=[double](Get-SemrPropertyValue $archiveRow[0] @('Archive_TotalItemSizeGB') 0)}
    }
    $targetSku=if($Row.TargetSku){[string]$Row.TargetSku}else{[string]$Config['DefaultTargetSku']}
    $unsupported=$hasArchive -and $targetSku -match 'F1|F3|DESKLESS'
    $tooLarge=$hasArchive -and $archiveSize -ge 95
    $archivePass=-not $unsupported -and -not $tooLarge
    Add-SemrFinding $Findings ($Base+@{CheckId='ARCHIVE-READINESS';Category='Mailbox';Severity=if(-not $archivePass){'Critical'}elseif($hasArchive){'Warning'}else{'Information'};Result=if(-not $archivePass){'FAIL'}elseif($hasArchive){'WARN'}else{'PASS'};IsBlocking=-not $archivePass;ObservedValue="Archive=$hasArchive; Size=$archiveSize GB; TargetSku=$targetSku";ExpectedValue='Archive-capable SKU and archive below 95 GB safety threshold';EvidenceSource="$source + archive inventory";Message=if($unsupported){'The selected target SKU is not archive-capable.'}elseif($tooLarge){'The archive is at or above the safe target threshold.'}elseif($hasArchive){'An archive is present and must be included in the migration plan.'}else{'No active source archive was detected.'};RecommendedAction=if($unsupported){'Select an archive-capable Exchange Online license.'}elseif($tooLarge){'Reduce archive content or validate auto-expanding archive requirements.'}elseif($hasArchive){'Confirm that MailboxType includes the archive.'}else{''}})

    $folderAvailable=[bool](Get-SemrPropertyValue $OnPrem @('FolderStatisticsAvailable') $false)
    $folders=@($OnPrem.FolderStatistics)
    $recoverable=@($folders | Where-Object {[string](Get-SemrPropertyValue $_ @('FolderType','Name','FolderPath') '') -match 'Recoverable|Purges|DiscoveryHolds|Versions'})
    $recoverableGb=0.0
    foreach($folder in $recoverable){$recoverableGb+=ConvertFrom-SemrByteSize (Get-SemrPropertyValue $folder @('FolderAndSubfolderSize','FolderSize') '')}
    $recoverableGb=[math]::Round($recoverableGb,2);$recoverableFail=$folderAvailable -and $recoverableGb -ge 90
    Add-SemrFinding $Findings ($Base+@{CheckId='MAILBOX-RECOVERABLE-ITEMS-QUOTA';Category='Mailbox';Severity=if(-not $folderAvailable){'Warning'}elseif($recoverableFail){'Critical'}elseif($recoverableGb -ge 80){'Warning'}else{'Information'};Result=if(-not $folderAvailable){'UNKNOWN'}elseif($recoverableFail){'FAIL'}elseif($recoverableGb -ge 80){'WARN'}else{'PASS'};IsBlocking=$recoverableFail;ObservedValue=if($folderAvailable){"$recoverableGb GB"}else{'Folder statistics unavailable'};ExpectedValue='< 90 GB safety threshold';EvidenceSource=$source;Message=if(-not $folderAvailable){'Recoverable Items usage could not be evaluated.'}elseif($recoverableFail){'Recoverable Items usage reached the safety threshold.'}elseif($recoverableGb -ge 80){'Recoverable Items usage is approaching the safety threshold.'}else{'Recoverable Items usage is below the safety threshold.'};RecommendedAction=if(-not $folderAvailable){'Run the final validation with live folder statistics.'}elseif($recoverableGb -ge 80){'Review hold-related growth and Recoverable Items quota.'}else{''}})

    $maxItems=0.0
    if($folderAvailable){foreach($folder in $folders){$value=[double](Get-SemrPropertyValue $folder @('ItemsInFolder','ItemsInFolderAndSubfolders') 0);if($value -gt $maxItems){$maxItems=$value}}}
    $folderFail=$folderAvailable -and $maxItems -ge 1000000
    Add-SemrFinding $Findings ($Base+@{CheckId='MAILBOX-FOLDER-LIMITS';Category='Mailbox';Severity=if(-not $folderAvailable){'Warning'}elseif($folderFail){'Critical'}elseif($maxItems -ge 900000){'Warning'}else{'Information'};Result=if(-not $folderAvailable){'UNKNOWN'}elseif($folderFail){'FAIL'}elseif($maxItems -ge 900000){'WARN'}else{'PASS'};IsBlocking=$folderFail;ObservedValue=if($folderAvailable){"Largest folder item count=$maxItems"}else{'Folder statistics unavailable'};ExpectedValue='< 1,000,000 items per folder';EvidenceSource=$source;Message=if(-not $folderAvailable){'Folder item limits could not be evaluated.'}elseif($folderFail){'At least one folder reaches the item limit.'}elseif($maxItems -ge 900000){'At least one folder is approaching the item limit.'}else{'Folder item counts are below the safety threshold.'};RecommendedAction=if(-not $folderAvailable){'Run the final validation with live folder statistics.'}elseif($maxItems -ge 900000){'Reduce oversized folders before migration.'}else{''}})

    $largeRaw=[string](Get-SemrPropertyValue $mailbox[0] @('LargeItemCount') '')
    $largeKnown=$largeRaw -match '^\d+$';$largeCount=if($largeKnown){[int]$largeRaw}else{0}
    Add-SemrFinding $Findings ($Base+@{CheckId='MAILBOX-LARGE-ITEMS';Category='Mailbox';Severity=if(-not $largeKnown -or $largeCount){'Warning'}else{'Information'};Result=if(-not $largeKnown){'UNKNOWN'}elseif($largeCount){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if($largeKnown){"$largeCount item(s) over inventory threshold"}else{'Large-item evidence unavailable'};ExpectedValue='0, or explicit migration exception policy';EvidenceSource=$source;Message=if(-not $largeKnown){'Large-item exposure could not be evaluated.'}elseif($largeCount){'Large items may require an explicit migration policy.'}else{'No oversized item was reported.'};RecommendedAction=if($largeCount){'Review sizes and explicitly approve BadItemLimit/LargeItemLimit.'}elseif(-not $largeKnown){'Refresh the mailbox inventory or run a targeted scan.'}else{''}})

    $quotaDefaults=Get-SemrPropertyValue $mailbox[0] @('UseDatabaseQuotaDefaults') $null
    $quotaKnown=$null -ne $quotaDefaults;$custom=$quotaKnown -and -not (ConvertTo-SemrBoolean $quotaDefaults)
    Add-SemrFinding $Findings ($Base+@{CheckId='CUSTOM-SOURCE-QUOTA';Category='Mailbox';Severity=if(-not $quotaKnown -or $custom){'Warning'}else{'Information'};Result=if(-not $quotaKnown){'UNKNOWN'}elseif($custom){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if($quotaKnown){"UseDatabaseQuotaDefaults=$quotaDefaults; ProhibitSendReceiveQuota=$([string](Get-SemrPropertyValue $mailbox[0] @('ProhibitSendReceiveQuota') ''))"}else{'Source quota evidence unavailable'};ExpectedValue='Source quota documented';EvidenceSource=$source;Message=if(-not $quotaKnown){'Source quota customization could not be evaluated.'}elseif($custom){'The mailbox uses a custom source quota.'}else{'The mailbox uses database quota defaults.'};RecommendedAction=if($custom){'Confirm the target SKU quota independently of the custom source quota.'}elseif(-not $quotaKnown){'Enrich the cache or validate live.'}else{''}})
}

function Get-SemrMoveState {
    param([Parameter(Mandatory)]$Exo)

    $states = [Collections.Generic.List[object]]::new()
    foreach ($entry in @($Exo.MigrationUsers)) {
        [void]$states.Add([pscustomobject]@{ Source = 'MigrationUser'; Status = [string](Get-SemrPropertyValue -InputObject $entry -Names @('Status','StatusSummary') -Default 'Unknown') })
    }
    foreach ($entry in @($Exo.MoveRequests)) {
        [void]$states.Add([pscustomobject]@{ Source = 'MoveRequest'; Status = [string](Get-SemrPropertyValue -InputObject $entry -Names @('Status','StatusSummary') -Default 'Unknown') })
    }
    $activeStates = @($states | Where-Object { $_.Status -notmatch '^(?i:Completed|CompletedWithWarning|Removed)$' })
    $badStates = @($states | Where-Object { $_.Status -match '(?i)Failed|Suspended|Stopped|Corrupt|Error' })
    return [pscustomobject]@{
        ObjectCount = $states.Count
        ActiveOperationCount = if ($activeStates.Count -gt 0) { 1 } else { 0 }
        ActiveStatuses = @($activeStates | ForEach-Object { [string]$_.Status } | Sort-Object -Unique)
        ActiveSources = @($activeStates | ForEach-Object { [string]$_.Source } | Sort-Object -Unique)
        BadStatuses = @($badStates | ForEach-Object { [string]$_.Status } | Sort-Object -Unique)
        TerminalStatuses = @($states | Where-Object { $_.Status -match '^(?i:Completed|CompletedWithWarning|Removed)$' } | ForEach-Object Status | Sort-Object -Unique)
    }
}


function Add-SemrFlowAndSyncFindings {
    param([System.Collections.Generic.List[object]]$Findings,[hashtable]$Base,[object[]]$BatchRows,$OnPrem,$Exo,$Graph)
    $source=[string](Get-SemrPropertyValue $OnPrem @('Source') 'Exchange on-premises')
    $exoSource=[string](Get-SemrPropertyValue $Exo @('Source') 'Exchange Online')
    $graphSource=[string](Get-SemrPropertyValue $Graph @('Source') 'Microsoft Graph')
    $mailbox=@($OnPrem.Mailboxes | Select-Object -First 1)
    if($mailbox.Count -eq 1){
        $batchEmails=@($BatchRows | ForEach-Object {([string]$_.EmailAddress).ToLowerInvariant()})
        $permissionsAvailable=[bool](Get-SemrPropertyValue $OnPrem @('PermissionsAvailable') $false)
        $delegates=@($OnPrem.Permissions | ForEach-Object {[string]$_.Delegate} | Where-Object {$_} | Sort-Object -Unique)
        $external=@($delegates | Where-Object {$batchEmails -notcontains $_.ToLowerInvariant()})
        Add-SemrFinding $Findings ($Base+@{CheckId='DELEGATE-MIGRATION-DEPENDENCY';Category='Permissions';Severity=if(-not $permissionsAvailable -or $external.Count){'Warning'}else{'Information'};Result=if(-not $permissionsAvailable){'UNKNOWN'}elseif($external.Count){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if(-not $permissionsAvailable){'Complete permission evidence unavailable'}elseif($external.Count){$external -join ';'}else{'All resolvable delegates are in the batch or none exist'};ExpectedValue='Delegates coordinated or documented';EvidenceSource=$source;Message=if(-not $permissionsAvailable){'Delegate migration dependencies could not be evaluated completely.'}elseif($external.Count){'One or more delegates are outside this migration batch.'}else{'No cross-wave delegate dependency was identified.'};RecommendedAction=if(-not $permissionsAvailable){'Refresh the permission inventory before final approval.'}elseif($external.Count){'Coordinate these delegates or document temporary cross-premises limitations.'}else{''}})

        $forwarding=[string](Get-SemrPropertyValue $mailbox[0] @('ForwardingSmtpAddress','ForwardingAddress') '')
        Add-SemrFinding $Findings ($Base+@{CheckId='MAIL-FORWARDING';Category='MailFlow';Severity=if($forwarding){'Warning'}else{'Information'};Result=if($forwarding){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if($forwarding){$forwarding}else{'No mailbox forwarding'};ExpectedValue='Forwarding documented';EvidenceSource=$source;Message=if($forwarding){'Mailbox-level forwarding is configured.'}else{'No mailbox-level forwarding was detected.'};RecommendedAction=if($forwarding){'Confirm the forwarding target and post-migration behavior.'}else{''}})

        $rulesAvailable=[bool](Get-SemrPropertyValue $OnPrem @('InboxRulesAvailable') $false)
        $forwardRules=@($OnPrem.InboxRules | Where-Object {(ConvertTo-SemrBoolean (Get-SemrPropertyValue $_ @('Enabled') $true)) -and (@(Get-SemrPropertyValue $_ @('ForwardTo','RedirectTo','ForwardAsAttachmentTo') @()).Count -gt 0)})
        Add-SemrFinding $Findings ($Base+@{CheckId='INBOX-FORWARDING-RULES';Category='MailFlow';Severity=if(-not $rulesAvailable -or $forwardRules.Count){'Warning'}else{'Information'};Result=if(-not $rulesAvailable){'UNKNOWN'}elseif($forwardRules.Count){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if($rulesAvailable){"$($forwardRules.Count) forwarding rule(s)"}else{'Inbox rule evidence unavailable'};ExpectedValue='Forwarding rules documented';EvidenceSource=$source;Message=if(-not $rulesAvailable){'Inbox forwarding rules could not be evaluated.'}elseif($forwardRules.Count){'Enabled forwarding or redirect rules were detected.'}else{'No enabled inbox forwarding rule was detected.'};RecommendedAction=if($forwardRules.Count){'Review destinations and validate each rule after migration.'}elseif(-not $rulesAvailable){'Run the final validation with live inbox-rule access.'}else{''}})

        $deliveryAvailable=[bool](Get-SemrPropertyValue $OnPrem @('DeliveryRestrictionsAvailable') $false)
        $moderated=ConvertTo-SemrBoolean (Get-SemrPropertyValue $mailbox[0] @('ModerationEnabled') $false)
        $restrictionCount=0
        foreach($name in @('AcceptMessagesOnlyFrom','AcceptMessagesOnlyFromDLMembers','RejectMessagesFrom','RejectMessagesFromDLMembers')){if(@(Get-SemrPropertyValue $mailbox[0] @($name) @()).Count){$restrictionCount++}}
        $restricted=$moderated -or $restrictionCount -gt 0
        Add-SemrFinding $Findings ($Base+@{CheckId='DELIVERY-RESTRICTIONS';Category='MailFlow';Severity=if(-not $deliveryAvailable -or $restricted){'Warning'}else{'Information'};Result=if(-not $deliveryAvailable){'UNKNOWN'}elseif($restricted){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if($deliveryAvailable){"Moderation=$moderated; RestrictionTypes=$restrictionCount"}else{'Delivery restriction evidence unavailable'};ExpectedValue='Delivery restrictions documented';EvidenceSource=$source;Message=if(-not $deliveryAvailable){'Delivery restrictions could not be evaluated from the selected cache.'}elseif($restricted){'Moderation or delivery restrictions are configured.'}else{'No moderation or explicit delivery restriction was detected.'};RecommendedAction=if(-not $deliveryAvailable){'Run the final validation against live Exchange on-premises.'}elseif($restricted){'Capture and validate restriction owners and membership after migration.'}else{''}})

        $dbAvailable=[bool](Get-SemrPropertyValue $OnPrem @('DatabaseHealthAvailable') $false)
        $db=@($OnPrem.DatabaseHealth | Select-Object -First 1)
        $mounted=$dbAvailable -and $db.Count -eq 1 -and (ConvertTo-SemrBoolean (Get-SemrPropertyValue $db[0] @('Mounted') $false))
        Add-SemrFinding $Findings ($Base+@{CheckId='EXCHANGE-DATABASE-HEALTH';Category='ExchangeInfrastructure';Severity=if(-not $dbAvailable){'Warning'}elseif($mounted){'Information'}else{'Critical'};Result=if(-not $dbAvailable){'UNKNOWN'}elseif($mounted){'PASS'}else{'FAIL'};IsBlocking=$dbAvailable -and -not $mounted;ObservedValue=if($dbAvailable){"Database=$([string](Get-SemrPropertyValue $mailbox[0] @('Database') '')); Mounted=$mounted"}else{'Database health unavailable'};ExpectedValue='Source mailbox database mounted';EvidenceSource=$source;Message=if(-not $dbAvailable){'Database health could not be evaluated.'}elseif($mounted){'The source mailbox database is mounted.'}else{'The source mailbox database is not mounted.'};RecommendedAction=if(-not $dbAvailable){'Validate database health immediately before the batch.'}elseif(-not $mounted){'Restore database availability before migration.'}else{''}})
    }

    $moveAvailable=[bool](Get-SemrPropertyValue $Exo @('MoveDataAvailable') $false)
    $moveState=Get-SemrMoveState -Exo $Exo
    $badStatuses=@($moveState.BadStatuses)
    Add-SemrFinding $Findings ($Base+@{CheckId='MOVE-HISTORY';Category='Migration';Severity=if(-not $moveAvailable){'Warning'}elseif($badStatuses.Count){'Critical'}else{'Information'};Result=if(-not $moveAvailable){'UNKNOWN'}elseif($badStatuses.Count){'FAIL'}else{'PASS'};IsBlocking=$moveAvailable -and $badStatuses.Count -gt 0;ObservedValue=if($moveAvailable){if($badStatuses.Count){$badStatuses -join ';'}else{'No failed or suspended prior move'}}else{'Move history unavailable'};ExpectedValue='No unresolved failed or suspended move';EvidenceSource=$exoSource;Message=if(-not $moveAvailable){'Previous move history could not be evaluated.'}elseif($badStatuses.Count){'An unresolved failed or suspended migration object was detected.'}else{'No unresolved failed move was returned.'};RecommendedAction=if($badStatuses.Count){'Review the migration report and resolve or clean up the failed move deliberately.'}elseif(-not $moveAvailable){'Refresh migration-jobs inventory or validate live.'}else{''}})

    $user=@($Graph.Users | Select-Object -First 1)
    if($user.Count -eq 1){
        $errors=@(Get-SemrPropertyValue $user[0] @('OnPremisesProvisioningErrors') @())
        $anchor=[string](Get-SemrPropertyValue $user[0] @('OnPremisesImmutableId') '')
        $lastSyncText=[string](Get-SemrPropertyValue $user[0] @('OnPremisesLastSyncDateTime') '')
        $lastSync=$null;if($lastSyncText){try{$lastSync=[datetime]$lastSyncText}catch{}}
        $timestampOld=$lastSync -and $lastSync -lt (Get-Date).AddHours(-24)
        $syncEnabled=ConvertTo-SemrBoolean (Get-SemrPropertyValue $user[0] @('OnPremisesSyncEnabled') $false)
        $syncError=$errors.Count -gt 0 -or ($syncEnabled -and [string]::IsNullOrWhiteSpace($anchor))
        Add-SemrFinding $Findings ($Base+@{CheckId='ENTRA-OBJECT-SYNC-ERROR';Category='MicrosoftGraph';Severity=if($syncError){'Critical'}else{'Information'};Result=if($syncError){'FAIL'}else{'PASS'};IsBlocking=$syncError;ObservedValue="ProvisioningErrors=$($errors.Count); SyncEnabled=$syncEnabled; ImmutableIdPresent=$(-not [string]::IsNullOrWhiteSpace($anchor)); ObjectLastSync=$lastSyncText";ExpectedValue='No provisioning errors and an identity anchor for synchronized users';EvidenceSource=$graphSource;Message=if($errors.Count){'Microsoft Entra reports provisioning errors.'}elseif($syncEnabled -and -not $anchor){'The synchronized identity anchor is missing.'}elseif($timestampOld){'The per-object synchronization timestamp is older than 24 hours; it is informational and is not used as a proxy for global Entra Connect scheduler health.'}else{'No object-level synchronization issue was detected.'};RecommendedAction=if($syncError){'Resolve Entra Connect export errors and identity anchoring.'}elseif($timestampOld){'Use the ENTRA-CONNECT-SCHEDULER result to assess current synchronization service health.'}else{''}})
    }
}

function Get-SemrHybridAdvancedEvidence {
    param([ValidateSet('Live','CacheOnly')][string]$Mode,[System.Collections.IDictionary]$Config)
    $capacityWarningThreshold = if($Config -and $Config.Contains('Hybrid') -and $Config['Hybrid'].Contains('ActiveMigrationWarningThreshold')){[int]$Config['Hybrid']['ActiveMigrationWarningThreshold']}else{100}
    $result=[ordered]@{
        MrsProxyAvailable=$false;MrsProxyEnabled=$false;MrsProxyMessage='MRSProxy evidence unavailable.';MrsProxySource='Unavailable';MrsProxySourceTimestamp=$null
        CertificateAvailable=$false;CertificateHealthy=$false;CertificateDaysRemaining=$null;CertificateMessage='Hybrid certificate evidence unavailable.';CertificateSource='Unavailable';CertificateSourceTimestamp=$null
        CapacityAvailable=$false;ActiveMigrationCount=$null;CapacityHealthy=$false;CapacityMessage='Migration load evidence unavailable.';CapacitySource='Unavailable';CapacitySourceTimestamp=$null
        OAuthAvailable=$false;OAuthHealthy=$false;OAuthMessage='Autodiscover/OAuth evidence unavailable.';OAuthSource='Unavailable';OAuthSourceTimestamp=$null
    }

    if($Mode -eq 'Live' -and $script:ConnectionState.OnPremisesExchange){
        $ewsCommand=Get-SemrOnPremCommandName 'Get-WebServicesVirtualDirectory'
        if((Test-SemrCheckEnabled 'HYBRID-MRSPROXY') -and $ewsCommand){
            $ews=@(Invoke-SemrCommandSafe $ewsCommand @{});$result.MrsProxyAvailable=$ews.Count -gt 0
            $enabled=@($ews | Where-Object {ConvertTo-SemrBoolean (Get-SemrPropertyValue $_ @('MRSProxyEnabled') $false)})
            $result.MrsProxyEnabled=$enabled.Count -gt 0;$result.MrsProxyMessage=if($result.MrsProxyEnabled){"MRSProxy enabled on $($enabled.Count) EWS virtual directorie(s)."}elseif($result.MrsProxyAvailable){'No EWS virtual directory has MRSProxy enabled.'}else{'EWS virtual directories could not be queried.'};$result.MrsProxySource='Live Exchange on-premises EWS virtual directories';$result.MrsProxySourceTimestamp=Get-Date
        }
        $certificateCommand=Get-SemrOnPremCommandName 'Get-ExchangeCertificate'
        if((Test-SemrCheckEnabled 'HYBRID-CERTIFICATE-EXPIRY') -and $certificateCommand){
            $certificates=@(Invoke-SemrCommandSafe $certificateCommand @{})
            $iis=@($certificates | Where-Object {[string](Get-SemrPropertyValue $_ @('Services') '') -match 'IIS' -and [string](Get-SemrPropertyValue $_ @('Status') 'Valid') -match 'Valid'})
            if($iis.Count){
                $result.CertificateAvailable=$true
                $days=@($iis | ForEach-Object {try{[math]::Floor((([datetime](Get-SemrPropertyValue $_ @('NotAfter') (Get-Date))) -(Get-Date)).TotalDays)}catch{-1}} | Sort-Object -Descending)
                $result.CertificateDaysRemaining=$days[0];$result.CertificateHealthy=$days[0] -ge 60
                $result.CertificateMessage="Best valid IIS certificate has $($days[0]) day(s) remaining.";$result.CertificateSource='Live Exchange on-premises certificates';$result.CertificateSourceTimestamp=Get-Date
            }
        }
        $orgCommand=Get-SemrOnPremCommandName 'Get-OrganizationConfig'
        $iocCommand=Get-SemrOnPremCommandName 'Get-IntraOrganizationConnector'
        if((Test-SemrCheckEnabled 'HYBRID-AUTODISCOVER-OAUTH') -and ($orgCommand -or $iocCommand)){
            $org=if($orgCommand){@(Invoke-SemrCommandSafe $orgCommand @{} | Select-Object -First 1)}else{@()}
            $ioc=if($iocCommand){@(Invoke-SemrCommandSafe $iocCommand @{})}else{@()}
            $oauth=$org.Count -and (ConvertTo-SemrBoolean (Get-SemrPropertyValue $org[0] @('OAuth2ClientProfileEnabled') $false))
            $enabledIoc=@($ioc | Where-Object {ConvertTo-SemrBoolean (Get-SemrPropertyValue $_ @('Enabled') $false)}).Count -gt 0
            $result.OAuthAvailable=$org.Count -gt 0 -or $ioc.Count -gt 0;$result.OAuthHealthy=$oauth -and $enabledIoc
            $result.OAuthMessage="OAuth2ClientProfileEnabled=$oauth; EnabledIntraOrganizationConnector=$enabledIoc.";$result.OAuthSource='Live Exchange on-premises hybrid configuration';$result.OAuthSourceTimestamp=Get-Date
        }
    }

    if((Test-SemrCheckEnabled 'HYBRID-ENDPOINT-CAPACITY') -and $Mode -eq 'Live' -and $script:ConnectionState.ExchangeOnline -and (Test-SemrCommand 'Get-MigrationUser')){
        try{
            $moves=@(Get-MigrationUser -ResultSize Unlimited -ErrorAction Stop | Where-Object {[string]$_.Status -notmatch 'Completed|Synced'})
            $result.CapacityAvailable=$true;$result.ActiveMigrationCount=$moves.Count;$result.CapacityHealthy=$moves.Count -lt $capacityWarningThreshold
            $result.CapacityMessage="$($moves.Count) active/non-terminal migration user(s); advisory threshold is $capacityWarningThreshold.";$result.CapacitySource='Live Exchange Online migration users';$result.CapacitySourceTimestamp=Get-Date
        }catch{$result.CapacityMessage=$_.Exception.Message}
    }

    if($script:InventoryContext -and $script:InventoryContext.HybridConfigDataAvailable){
        $rows=@($script:InventoryContext.HybridConfigRows)
        if((Test-SemrCheckEnabled 'HYBRID-MRSPROXY') -and -not $result.MrsProxyAvailable){
            $mrs=@($rows | Where-Object {$_.Setting -eq 'MRSProxyEnabled'})
            if($mrs.Count){$result.MrsProxyAvailable=$true;$result.MrsProxyEnabled=@($mrs | Where-Object {ConvertTo-SemrBoolean $_.Value}).Count -gt 0;$result.MrsProxyMessage="CSV cache: $(@($mrs | Where-Object {ConvertTo-SemrBoolean $_.Value}).Count)/$($mrs.Count) EWS virtual directorie(s) have MRSProxy enabled.";$result.MrsProxySource='Tenant CSV cache hybrid configuration';$result.MrsProxySourceTimestamp=$script:InventoryContext.HybridConfigTimestamp}
        }
        if((Test-SemrCheckEnabled 'HYBRID-CERTIFICATE-EXPIRY') -and -not $result.CertificateAvailable){
            $groups=@($rows | Where-Object {$_.Category -eq 'ExchangeCertificate'} | Group-Object ObjectName)
            $days=[System.Collections.Generic.List[int]]::new()
            foreach($group in $groups){
                $services=[string](Get-SemrPropertyValue ($group.Group | Where-Object Setting -eq 'Services' | Select-Object -First 1) @('Value') '')
                $status=[string](Get-SemrPropertyValue ($group.Group | Where-Object Setting -eq 'Status' | Select-Object -First 1) @('Value') '')
                $notAfter=[string](Get-SemrPropertyValue ($group.Group | Where-Object Setting -eq 'NotAfter' | Select-Object -First 1) @('Value') '')
                if($services -match 'IIS' -and $status -match 'Valid' -and $notAfter){try{[void]$days.Add([math]::Floor((([datetime]$notAfter)-(Get-Date)).TotalDays))}catch{}}
            }
            if($days.Count){$best=@($days | Sort-Object -Descending)[0];$result.CertificateAvailable=$true;$result.CertificateDaysRemaining=$best;$result.CertificateHealthy=$best -ge 60;$result.CertificateMessage="CSV cache: best valid IIS certificate has $best day(s) remaining.";$result.CertificateSource='Tenant CSV cache hybrid configuration';$result.CertificateSourceTimestamp=$script:InventoryContext.HybridConfigTimestamp}
        }
        if((Test-SemrCheckEnabled 'HYBRID-AUTODISCOVER-OAUTH') -and -not $result.OAuthAvailable){
            $oauthRow=@($rows | Where-Object {$_.Category -eq 'OrganizationConfig' -and $_.Setting -eq 'OAuth2ClientProfileEnabled'} | Select-Object -First 1)
            $iocRows=@($rows | Where-Object {$_.Category -eq 'IntraOrganizationConnector' -and $_.Setting -eq 'Enabled'})
            if($oauthRow.Count -or $iocRows.Count){$oauth=$oauthRow.Count -and (ConvertTo-SemrBoolean $oauthRow[0].Value);$enabledIoc=@($iocRows | Where-Object {ConvertTo-SemrBoolean $_.Value}).Count -gt 0;$result.OAuthAvailable=$true;$result.OAuthHealthy=$oauth -and $enabledIoc;$result.OAuthMessage="CSV cache: OAuth2ClientProfileEnabled=$oauth; EnabledIntraOrganizationConnector=$enabledIoc.";$result.OAuthSource='Tenant CSV cache hybrid configuration';$result.OAuthSourceTimestamp=$script:InventoryContext.HybridConfigTimestamp}
        }
    }
    return [pscustomobject]$result
}

function Get-SemrMigrationEndpointOption {
    [CmdletBinding()]
    param()

    if (-not $script:ConnectionState.ExchangeOnline) {
        throw 'Exchange Online is not connected.'
    }
    if (-not (Test-SemrCommand -Name 'Get-MigrationEndpoint')) {
        throw 'Get-MigrationEndpoint is unavailable in the Exchange Online session.'
    }
    return @(
        Get-MigrationEndpoint -ErrorAction Stop |
            Where-Object { [string]$_.EndpointType -match '^ExchangeRemoteMove$' } |
            ForEach-Object {
                [pscustomobject][ordered]@{
                    Name = [string]$_.Identity
                    EndpointType = [string]$_.EndpointType
                    RemoteServer = [string](Get-SemrPropertyValue -InputObject $_ -Names @('RemoteServer','RemoteServerName') -Default '')
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } |
            Sort-Object Name -Unique
    )
}

function Test-SemrHybridReadiness {
    [CmdletBinding()]
    param([System.Collections.IDictionary]$Config)

    $mode = if ($Config.Contains('Mode')) { [string]$Config['Mode'] } else { 'Live' }
    $hybridConfig = if ($Config.Contains('Hybrid')) { $Config['Hybrid'] } else { $Config }
    $advanced = Get-SemrHybridAdvancedEvidence -Mode $mode -Config $Config
    $result = [ordered]@{
        Available = $false
        EndpointFound = $false
        EndpointName = ''
        ConnectivitySuccess = $false
        ConnectivityTestAvailable = $false
        SourceTimestamp = if ($mode -eq 'Live') { Get-Date } else { $null }
        Source = if ($mode -eq 'CacheOnly') { 'CacheOnly' } else { 'Live Exchange Online' }
        Message = ''
        MrsProxyAvailable = $advanced.MrsProxyAvailable
        MrsProxyEnabled = $advanced.MrsProxyEnabled
        MrsProxyMessage = $advanced.MrsProxyMessage
        MrsProxySource = $advanced.MrsProxySource
        MrsProxySourceTimestamp = $advanced.MrsProxySourceTimestamp
        CertificateAvailable = $advanced.CertificateAvailable
        CertificateHealthy = $advanced.CertificateHealthy
        CertificateDaysRemaining = $advanced.CertificateDaysRemaining
        CertificateMessage = $advanced.CertificateMessage
        CertificateSource = $advanced.CertificateSource
        CertificateSourceTimestamp = $advanced.CertificateSourceTimestamp
        CapacityAvailable = $advanced.CapacityAvailable
        ActiveMigrationCount = $advanced.ActiveMigrationCount
        CapacityHealthy = $advanced.CapacityHealthy
        CapacityMessage = $advanced.CapacityMessage
        CapacitySource = $advanced.CapacitySource
        CapacitySourceTimestamp = $advanced.CapacitySourceTimestamp
        OAuthAvailable = $advanced.OAuthAvailable
        OAuthHealthy = $advanced.OAuthHealthy
        OAuthMessage = $advanced.OAuthMessage
        OAuthSource = $advanced.OAuthSource
        OAuthSourceTimestamp = $advanced.OAuthSourceTimestamp
    }
    if (-not (Test-SemrCheckEnabled -CheckId 'HYBRID-ENDPOINT')) {
        $result.Message = 'Migration endpoint check disabled for this run.'
        return [pscustomobject]$result
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
        $result.Available = $false
        $result.Message = $_.Exception.Message
        return [pscustomobject]$result
    }

    if ($endpoints.Count -eq 0) {
        $result.Message = 'No ExchangeRemoteMove migration endpoint was found.'
        return [pscustomobject]$result
    }
    if (-not $configuredName -and $endpoints.Count -gt 1) {
        $endpointNames = @($endpoints | ForEach-Object { [string]$_.Identity } | Where-Object { $_ })
        $result.Message = "Multiple ExchangeRemoteMove endpoints were found ($($endpointNames -join ', ')). Configure Hybrid.MigrationEndpointName explicitly."
        return [pscustomobject]$result
    }
    $endpoint = $endpoints[0]
    $result.EndpointFound = $true
    $result.EndpointName = [string]$endpoint.Identity
    if (Test-SemrCommand -Name 'Test-MigrationServerAvailability') {
        $result.ConnectivityTestAvailable = $true
        try {
            $test = Test-MigrationServerAvailability -Endpoint $endpoint.Identity -ErrorAction Stop
            $result.ConnectivitySuccess = [bool](Get-SemrPropertyValue -InputObject $test -Names @('Result', 'Success') -Default $false)
            if (-not $result.ConnectivitySuccess -and [string]$test.Result -eq 'Success') {
                $result.ConnectivitySuccess = $true
            }
            $result.Message = [string](Get-SemrPropertyValue -InputObject $test -Names @('Message', 'ErrorDetail') -Default $test.Result)
            if ($result.ConnectivitySuccess -and [string]::IsNullOrWhiteSpace($result.Message)) {
                $result.Message = 'Connectivity test succeeded.'
            }
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
    $maximumLastSyncAgeMinutes = if ($Config -and $Config.Contains('EntraConnectHealth')) { [double]$Config['EntraConnectHealth']['MaximumLastSyncAgeMinutes'] } else { 120.0 }
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
                $lastSyncValue = Get-SemrPropertyValue -InputObject $scheduler -Names @('LastSyncCycleEndTimeInUTC','LastSyncCycleStartTimeInUTC') -Default $null
                $lastSyncDateTimeUtc = if ($lastSyncValue) {
                    $parsedLastSync = [datetime]$lastSyncValue
                    if ($parsedLastSync.Kind -eq [DateTimeKind]::Unspecified) {
                        [datetime]::SpecifyKind($parsedLastSync, [DateTimeKind]::Utc)
                    }
                    else { $parsedLastSync.ToUniversalTime() }
                }
                else { $null }
                $lastSyncAgeMinutes = if ($lastSyncDateTimeUtc) { [math]::Round(((Get-Date).ToUniversalTime() - $lastSyncDateTimeUtc).TotalMinutes, 1) } else { $null }
                $lastSyncFresh = $null -ne $lastSyncAgeMinutes -and ($maximumLastSyncAgeMinutes -le 0 -or $lastSyncAgeMinutes -le $maximumLastSyncAgeMinutes)
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
                    LastSyncDateTimeUtc = $lastSyncDateTimeUtc
                    LastSyncAgeMinutes = $lastSyncAgeMinutes
                    LastSyncFresh = $lastSyncFresh
                    ConnectorRunStatus = ($runStatus | ConvertTo-Json -Compress)
                    Message = "Live Microsoft Entra Connect scheduler state collected from local ADSync cmdlets; last sync age=$lastSyncAgeMinutes minute(s), maximum=$maximumLastSyncAgeMinutes."
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
        LastSyncDateTimeUtc = $null
        LastSyncAgeMinutes = $null
        LastSyncFresh = $false
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
    $result.NextSyncCycleStartTimeInUTC = Get-SemrPropertyValue -InputObject $lastSyncAgeRow[0] -Names @('NextSyncCycleStartTimeInUTC') -Default $null
    $result.ConnectorRunStatus = ($rows | Select-Object CheckName, Status, Value, Detail | ConvertTo-Json -Compress)
    $lastSyncDescription = [string](Get-SemrPropertyValue -InputObject $lastSyncAgeRow[0] -Names @('Detail') -Default 'Last synchronization timestamp unavailable.')
    $lastSyncValue = Get-SemrPropertyValue -InputObject $lastSyncAgeRow[0] -Names @('LastSyncDateTimeUtc','LastSyncDateTime') -Default $null
    if ($lastSyncValue) {
        try {
            $parsedLastSync = [datetime]$lastSyncValue
            $result.LastSyncDateTimeUtc = if ($parsedLastSync.Kind -eq [DateTimeKind]::Unspecified) {
                [datetime]::SpecifyKind($parsedLastSync, [DateTimeKind]::Utc)
            }
            else { $parsedLastSync.ToUniversalTime() }
            $result.LastSyncAgeMinutes = [math]::Round(((Get-Date).ToUniversalTime() - $result.LastSyncDateTimeUtc).TotalMinutes, 1)
            $result.LastSyncFresh = $maximumLastSyncAgeMinutes -le 0 -or $result.LastSyncAgeMinutes -le $maximumLastSyncAgeMinutes
            $lastSyncDescription = "Last sync age $($result.LastSyncAgeMinutes) minute(s), maximum $maximumLastSyncAgeMinutes, calculated from $lastSyncValue."
        }
        catch { $null = $_ }
    }
    $snapshotAgeHours = if ($script:InventoryContext.EntraTimestamp) { [math]::Round(((Get-Date) - [datetime]$script:InventoryContext.EntraTimestamp).TotalHours, 2) } else { $null }
    $cacheSummary = "CSV cache age=$snapshotAgeHours hour(s); SyncEnabled=$($result.SyncCycleEnabled); SyncHealthOK=$healthOk; $lastSyncDescription"
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

    Set-SemrAssessmentCheckOptions -Config $Config
    $script:ActiveDirectoryFallbackReasons.Clear()
    $script:ActiveDirectoryFallbackUsed = $false
    $runId = "SEMR-{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    $findings = [System.Collections.Generic.List[object]]::new()
    $globalFindings = [System.Collections.Generic.List[object]]::new()
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
    if ($ProgressCallback) { & $ProgressCallback 0 $rows.Count '' 'Testing hybrid migration endpoint; please wait' }
    $hybrid = Test-SemrHybridReadiness -Config $Config
    if ($ProgressCallback) { & $ProgressCallback 0 $rows.Count '' 'Collecting accepted domains and advanced identity evidence; please wait' }
    $acceptedDomains = Get-SemrAcceptedDomainEvidence
    if ($ProgressCallback) { & $ProgressCallback 0 $rows.Count '' 'Checking Microsoft Entra Connect health; please wait' }
    $entraConnect = Test-SemrEntraConnect -Config $Config
    $csvSources = @()
    $licenseCapacityCache = @{}
    $index = 0

    foreach ($row in $rows) {
        $index++
        if ($CancellationCheck -and (& $CancellationCheck)) { break }
        if ($ProgressCallback) {
            & $ProgressCallback $index $rows.Count $row.EmailAddress "Collecting evidence"
        }

        $email = [string]$row.EmailAddress
        $resolvedRecipientType = ''
        $mailboxTargetPolicy = $null
        $base = @{
            RunId = $runId
            EmailAddress = $email
        }
        $script:CurrentSourceTimestamps = @{}
        if ($Batch.SourceTimestamp) { $script:CurrentSourceTimestamps['Batch CSV'] = $Batch.SourceTimestamp }
        if ($hybrid.Source -and $hybrid.SourceTimestamp) { $script:CurrentSourceTimestamps[[string]$hybrid.Source] = $hybrid.SourceTimestamp }
        if ($entraConnect.Source -and $entraConnect.SourceTimestamp) { $script:CurrentSourceTimestamps[[string]$entraConnect.Source] = $entraConnect.SourceTimestamp }
        if ($acceptedDomains.Source -and $acceptedDomains.SourceTimestamp) { $script:CurrentSourceTimestamps[[string]$acceptedDomains.Source] = $acceptedDomains.SourceTimestamp }

        if ([string]::IsNullOrWhiteSpace($email)) {
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'CSV-EMPTY-IDENTITY'; Category = 'CSV'; Severity = 'Critical'; Result = 'FAIL'; IsBlocking = $true
                Message = "CSV row $($row.RowNumber) has an empty mailbox identity."; RecommendedAction = 'Populate EmailAddress and reload the CSV.'
            })
            continue
        }

        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'CSV-EMPTY-IDENTITY'; Category = 'CSV'; Severity = 'Information'; Result = 'PASS'; IsBlocking = $false
            ObservedValue = $email; ExpectedValue = 'Non-empty mailbox identity'; EvidenceSource = 'Batch CSV'
            Message = 'The CSV row contains a mailbox identity.'; RecommendedAction = ''
        })

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
        if ($ad.SourceTimestamp) { $script:CurrentSourceTimestamps[$adSource] = $ad.SourceTimestamp }
        if ($onPrem.SourceTimestamp) { $script:CurrentSourceTimestamps[$onPremSource] = $onPrem.SourceTimestamp }

        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'AD-SOURCE'; Category = 'ActiveDirectory'; Severity = if ($ad.Available) { 'Information' } else { 'Critical' }
            Result = if ($ad.Available) { 'PASS' } else { 'UNKNOWN' }; IsBlocking = -not $ad.Available
            ObservedValue = if ($ad.Available) { 'Available' } else { [string]$ad.Message }; ExpectedValue = 'Available AD evidence'; EvidenceSource = $adSource
            Message = if ($ad.Available) { 'Active Directory evidence is available.' } else { 'Active Directory evidence is unavailable.' }
            RecommendedAction = if ($ad.Available) { '' } else { 'Restore live AD access or refresh the configured AD inventory CSV, then rerun.' }
        })
        if ($ad.Available) {
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

        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'ONPREM-SOURCE'; Category = 'ExchangeOnPrem'; Severity = if ($onPrem.Available) { 'Information' } else { 'Critical' }
            Result = if ($onPrem.Available) { 'PASS' } else { 'UNKNOWN' }; IsBlocking = -not $onPrem.Available
            ObservedValue = if ($onPrem.Available) { 'Available' } else { [string]$onPrem.Message }; ExpectedValue = 'Available Exchange on-premises evidence'; EvidenceSource = $onPremSource
            Message = if ($onPrem.Available) { 'Exchange on-premises evidence is available.' } else { 'Exchange on-premises evidence is unavailable.' }
            RecommendedAction = if ($onPrem.Available) { '' } else { 'Restore live Exchange on-premises access or refresh the configured inventory CSV, then rerun.' }
        })
        if ($onPrem.Available) {
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
                $resolvedRecipientType = [string]$mailbox.RecipientTypeDetails
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
                $targetSkuExplicit = -not [string]::IsNullOrWhiteSpace([string]$row.TargetSku)
                $mailboxTargetPolicy = Get-SemrTargetQuotaEvidence -TargetSku $targetSku -Config $Config -MailboxType $resolvedRecipientType -TargetSkuExplicit:$targetSkuExplicit
                $bufferPercent = [double]$Config['QuotaSafetyBufferPercent']
                $safeQuota = if ($mailboxTargetPolicy.Known -and $mailboxTargetPolicy.Eligible) { $mailboxTargetPolicy.QuotaGb * (1.0 - ($bufferPercent / 100.0)) } else { 0.0 }
                $sizePass = $mailboxTargetPolicy.Known -and $mailboxTargetPolicy.Eligible -and $statistics.Count -eq 1 -and $sizeGb -lt $safeQuota
                $quotaResult = if (-not $mailboxTargetPolicy.Known) { 'UNKNOWN' } elseif (-not $mailboxTargetPolicy.Eligible) { 'FAIL' } elseif ($statistics.Count -eq 0) { 'UNKNOWN' } elseif ($sizePass) { 'PASS' } else { 'FAIL' }
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'MAILBOX-TARGET-QUOTA'; Category = 'Mailbox'; Severity = if ($sizePass) { 'Information' } else { 'Critical' }
                    Result = $quotaResult
                    IsBlocking = -not $sizePass; ObservedValue = "$sizeGb GB"; ExpectedValue = if ($mailboxTargetPolicy.Known -and $mailboxTargetPolicy.Eligible) { "< $([math]::Round($safeQuota, 2)) GB ($($mailboxTargetPolicy.TargetSku) / $($mailboxTargetPolicy.QuotaGb) GB less $bufferPercent% buffer)" } else { 'A mailbox-eligible target SKU with an explicitly configured quota' }
                    EvidenceSource = "$onPremSource mailbox statistics + configured target SKU policy"; Message = if (-not $mailboxTargetPolicy.Known) { $mailboxTargetPolicy.Message } elseif (-not $mailboxTargetPolicy.Eligible) { $mailboxTargetPolicy.Message } elseif ($sizePass) { 'Mailbox size fits the target quota with the configured safety buffer.' } elseif ($statistics.Count -eq 0) { 'Mailbox statistics could not be collected.' } else { 'Mailbox size exceeds the safe target quota.' }
                    RecommendedAction = if ($sizePass) { '' } elseif (-not $mailboxTargetPolicy.Known) { 'Add the approved SKU and its mailbox quota to TargetQuotaGbBySku, or set TargetSku correctly in the batch.' } elseif (-not $mailboxTargetPolicy.Eligible) { 'Select a target SKU that grants an Exchange Online mailbox.' } else { 'Reduce mailbox size, enable/archive content, or select a target SKU with sufficient quota.' }
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
        if ($exo.SourceTimestamp) { $script:CurrentSourceTimestamps[$exoSource] = $exo.SourceTimestamp }
        if ($graph.SourceTimestamp) { $script:CurrentSourceTimestamps[$graphSource] = $graph.SourceTimestamp }
        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'EXO-SOURCE'; Category = 'ExchangeOnline'; Severity = if ($exo.Available) { 'Information' } else { 'Critical' }
            Result = if ($exo.Available) { 'PASS' } else { 'UNKNOWN' }; IsBlocking = -not $exo.Available
            ObservedValue = if ($exo.Available) { 'Available' } else { 'Unavailable' }; ExpectedValue = 'Available Exchange Online evidence'; EvidenceSource = $exoSource
            Message = if ($exo.Available) { 'Exchange Online evidence is available.' } elseif ([string]$Config['Mode'] -eq 'CacheOnly') { 'The cached Exchange Online inventories are unavailable.' } else { 'Exchange Online is not connected.' }
            RecommendedAction = if ($exo.Available) { '' } elseif ([string]$Config['Mode'] -eq 'CacheOnly') { 'Refresh or restore the tenant CSV cache, then run the assessment again.' } else { 'Connect Exchange Online interactively and run the assessment again.' }
        })
        if ($exo.Available) {
            $cloudRecipientCount = @($exo.Recipients).Count
            $cloudMailboxCount = @($exo.Mailboxes).Count
            $recipientDataAvailable = [bool](Get-SemrPropertyValue -InputObject $exo -Names @('RecipientDataAvailable') -Default $false)
            $mailboxDataAvailable = [bool](Get-SemrPropertyValue -InputObject $exo -Names @('MailboxDataAvailable') -Default $false)
            $cloudStateEvidenceAvailable = $recipientDataAvailable -and $mailboxDataAvailable
            $cloudStateValid = $cloudStateEvidenceAvailable -and $cloudRecipientCount -eq 1 -and $cloudMailboxCount -eq 0
            $lookupMethods = @((Get-SemrPropertyValue -InputObject $exo -Names @('RecipientLookupMethods') -Default @())) -join ','
            $lookupErrors = @((Get-SemrPropertyValue -InputObject $exo -Names @('RecipientLookupErrors') -Default @())) -join ' | '
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'EXO-RECIPIENT-STATE'; Category = 'ExchangeOnline'; Severity = if ($cloudStateValid) { 'Information' } else { 'Critical' }
                Result = if (-not $cloudStateEvidenceAvailable) { 'UNKNOWN' } elseif ($cloudStateValid) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $cloudStateValid
                ObservedValue = "Recipients=$cloudRecipientCount; ActiveMailboxes=$cloudMailboxCount; Methods=$lookupMethods; Errors=$lookupErrors"; ExpectedValue = 'One synchronized MailUser and no active EXO mailbox'
                EvidenceSource = $exoSource; Message = if (-not $cloudStateEvidenceAvailable) { 'The Exchange Online recipient or mailbox lookup did not complete reliably.' } elseif ($cloudStateValid) { 'The cloud recipient state is consistent with hybrid pre-onboarding.' } else { 'The cloud recipient is missing, ambiguous, or already has an active mailbox.' }
                RecommendedAction = if (-not $cloudStateEvidenceAvailable) { 'Verify delegated Exchange Online RBAC and cmdlet availability, then rerun.' } elseif ($cloudStateValid) { '' } else { 'Resolve synchronization, duplicate recipient, or split-brain mailbox state before migration.' }
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
            $moveState = Get-SemrMoveState -Exo $exo
            $activeMoveCount = [int]$moveState.ActiveOperationCount
            $activeMoveDescription = if ($activeMoveCount -gt 0) { "ActiveOperation=1; Sources=$($moveState.ActiveSources -join ','); Statuses=$($moveState.ActiveStatuses -join ',')" } elseif ($moveState.TerminalStatuses.Count -gt 0) { "No active operation; terminal history=$($moveState.TerminalStatuses -join ',')" } else { 'No migration object' }
            $existingBatchPhase = [string]$Config['AssessmentPhase'] -eq 'ExistingBatch'
            $existingMoveResult = if (-not $moveDataAvailable) { 'UNKNOWN' } elseif ($existingBatchPhase -and $activeMoveCount -eq 0) { 'WARN' } elseif ($existingBatchPhase -or $activeMoveCount -eq 0) { 'PASS' } else { 'FAIL' }
            $existingMoveBlocking = -not $existingBatchPhase -and $moveDataAvailable -and $activeMoveCount -gt 0
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'EXO-EXISTING-MOVE'; Category = 'Migration'; Severity = if ($existingMoveResult -eq 'FAIL') { 'Critical' } elseif ($existingMoveResult -in @('WARN','UNKNOWN')) { 'Warning' } else { 'Information' }
                Result = $existingMoveResult; IsBlocking = $existingMoveBlocking
                ObservedValue = if ($moveDataAvailable) { $activeMoveDescription } else { 'Migration job data is unavailable.' }; ExpectedValue = if ($existingBatchPhase) { 'Existing active operation expected and evaluated as part of the current batch' } else { 'No active or unresolved migration operation before batch creation' }
                EvidenceSource = $exoSource; Message = if (-not $moveDataAvailable) { 'Existing move state could not be evaluated.' } elseif ($existingBatchPhase -and $activeMoveCount -gt 0) { 'The active migration operation is expected in ExistingBatch phase. MigrationUser and MoveRequest representations were deduplicated.' } elseif ($existingBatchPhase) { 'No active migration operation was returned for this ExistingBatch assessment.' } elseif ($activeMoveCount -gt 0) { 'An active or unresolved migration operation already exists before batch creation.' } else { 'No active migration operation was returned; terminal history is not treated as an active move.' }
                RecommendedAction = if (-not $moveDataAvailable) { 'Refresh the migration-jobs inventory or run the final validation in Live mode.' } elseif ($existingBatchPhase -and $activeMoveCount -eq 0) { 'Confirm that the mailbox still belongs to the intended migration batch or review its terminal history.' } elseif (-not $existingBatchPhase -and $activeMoveCount -gt 0) { 'Review the active operation and clean it up only after confirming its state and history.' } else { '' }
            })
        }

        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'GRAPH-SOURCE'; Category = 'MicrosoftGraph'; Severity = if ($graph.Available) { 'Information' } else { 'Critical' }
            Result = if ($graph.Available) { 'PASS' } else { 'UNKNOWN' }; IsBlocking = -not $graph.Available
            ObservedValue = if ($graph.Available) { 'Available' } else { 'Unavailable' }; ExpectedValue = 'Available Microsoft Graph evidence'; EvidenceSource = $graphSource
            Message = if ($graph.Available) { 'Microsoft Graph evidence is available.' } elseif ([string]$Config['Mode'] -eq 'CacheOnly') { 'The cached Microsoft 365 user inventory is unavailable.' } else { 'Microsoft Graph is not connected.' }
            RecommendedAction = if ($graph.Available) { '' } elseif ([string]$Config['Mode'] -eq 'CacheOnly') { 'Refresh or restore the tenant CSV cache, then run the assessment again.' } else { 'Connect Microsoft Graph interactively and run the assessment again.' }
        })
        if ($graph.Available) {
            $graphCount = @($graph.Users).Count
            $graphQueryError = [string](Get-SemrPropertyValue -InputObject $graph -Names @('QueryError') -Default '')
            $graphEvidenceAvailable = [string]::IsNullOrWhiteSpace($graphQueryError)
            $graphValid = $graphEvidenceAvailable -and $graphCount -eq 1
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = 'GRAPH-USER-STATE'; Category = 'MicrosoftGraph'; Severity = if ($graphValid) { 'Information' } else { 'Critical' }
                Result = if (-not $graphEvidenceAvailable) { 'UNKNOWN' } elseif ($graphValid) { 'PASS' } else { 'FAIL' }; IsBlocking = -not $graphValid
                ObservedValue = if ($graphEvidenceAvailable) { "$graphCount matching user(s)" } else { "QueryError=$graphQueryError" }; ExpectedValue = '1'
                EvidenceSource = $graphSource; Message = if (-not $graphEvidenceAvailable) { 'The Microsoft Graph user query did not complete reliably.' } elseif ($graphValid) { 'Exactly one Microsoft Entra user was resolved.' } else { 'The Microsoft Entra user is missing or ambiguous.' }
                RecommendedAction = if (-not $graphEvidenceAvailable) { 'Verify delegated Graph permissions and rerun the assessment.' } elseif ($graphValid) { '' } else { 'Resolve Microsoft Entra synchronization and identity uniqueness.' }
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
                $existingBatchPhase = [string]$Config['AssessmentPhase'] -eq 'ExistingBatch'
                $licenseResult = if ($existingBatchPhase) { if ($licenseCount -gt 0) { 'PASS' } else { 'WARN' } } else { if ($licenseCount -gt 0) { 'WARN' } else { 'PASS' } }
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'LICENSE-PRE-MIGRATION'; Category = 'Licensing'; Severity = if ($licenseResult -eq 'WARN') { 'Warning' } else { 'Information' }
                    Result = $licenseResult; IsBlocking = $false
                    ObservedValue = "$licenseCount assigned license(s)"; ExpectedValue = if ($existingBatchPhase) { 'Approved target license assigned for the existing migration batch' } else { 'License available; Exchange license assigned after migration' }
                    EvidenceSource = $graphSource; Message = if ($existingBatchPhase -and $licenseCount -gt 0) { 'The assigned license is expected for the existing migration batch.' } elseif ($existingBatchPhase) { 'No assigned license was returned for the existing migration batch.' } elseif ($licenseCount -gt 0) { 'One or more licenses are already assigned before migration.' } else { 'No assigned license was returned before migration.' }
                    RecommendedAction = if ($existingBatchPhase -and $licenseCount -eq 0) { 'Confirm the target licensing plan and assignment timing for this active migration.' } elseif (-not $existingBatchPhase -and $licenseCount -gt 0) { 'Verify that an Exchange service plan has not provisioned an unintended cloud mailbox.' } elseif (-not $existingBatchPhase) { 'Assign the approved Exchange Online license after migration completion and within the operational deadline.' } else { '' }
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
                if (-not $mailboxTargetPolicy) {
                    $mailboxTargetPolicy = Get-SemrTargetQuotaEvidence -TargetSku $targetSku -Config $Config -MailboxType $resolvedRecipientType -TargetSkuExplicit:($null -ne $row.TargetSku -and -not [string]::IsNullOrWhiteSpace([string]$row.TargetSku))
                }
                $licenseCapacity = $null
                if ($mailboxTargetPolicy.RequiresLicense -and $mailboxTargetPolicy.Known -and $mailboxTargetPolicy.Eligible) {
                    $licenseKey = ([string]$mailboxTargetPolicy.TargetSku).ToUpperInvariant()
                    if (-not $licenseCapacityCache.ContainsKey($licenseKey)) {
                        $licenseCapacityCache[$licenseKey] = Get-SemrTenantLicenseEvidence -TargetSku $mailboxTargetPolicy.TargetSku -Config $Config
                    }
                    $licenseCapacity = $licenseCapacityCache[$licenseKey]
                }
                $capacityPass = -not $mailboxTargetPolicy.RequiresLicense -or ($licenseCapacity -and $licenseCapacity.Found -and $licenseCapacity.AvailableUnits -gt 0)
                $capacityResult = if (-not $mailboxTargetPolicy.RequiresLicense) { 'PASS' } elseif (-not $mailboxTargetPolicy.Known -or -not $mailboxTargetPolicy.Eligible) { 'FAIL' } elseif (-not $licenseCapacity.Available) { 'UNKNOWN' } elseif ($capacityPass) { 'PASS' } else { 'FAIL' }
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'LICENSE-CAPACITY'; Category = 'Licensing'; Severity = if ($capacityPass) { 'Information' } else { 'Critical' }
                    Result = $capacityResult
                    IsBlocking = -not $capacityPass; ObservedValue = if (-not $mailboxTargetPolicy.RequiresLicense) { $mailboxTargetPolicy.Message } elseif ($licenseCapacity) { $licenseCapacity.Message } else { $mailboxTargetPolicy.Message }; ExpectedValue = if ($mailboxTargetPolicy.RequiresLicense) { "At least one available $($mailboxTargetPolicy.TargetSku) license" } else { 'No direct license required below the shared mailbox quota' }
                    EvidenceSource = if ($mailboxTargetPolicy.RequiresLicense) { "$graphSource subscribedSkus" } else { 'Configured shared mailbox policy' }; Message = if (-not $mailboxTargetPolicy.RequiresLicense) { 'This shared mailbox does not require a direct license while it stays within the 50 GB limit and uses no licensed-only feature.' } elseif ($capacityPass) { 'The target license SKU has available capacity.' } elseif (-not $mailboxTargetPolicy.Known -or -not $mailboxTargetPolicy.Eligible) { $mailboxTargetPolicy.Message } else { 'The target license SKU is missing or has no available capacity.' }
                    RecommendedAction = if ($capacityPass) { '' } elseif (-not $mailboxTargetPolicy.Known -or -not $mailboxTargetPolicy.Eligible) { 'Select an approved mailbox-eligible target SKU.' } else { 'Add license capacity or select an approved target SKU before the migration wave.' }
                })

                $servicePlanPass = -not $mailboxTargetPolicy.RequiresLicense -or ($licenseCapacity -and $licenseCapacity.ExchangeServicePlanFound)
                $servicePlanResult = if (-not $mailboxTargetPolicy.RequiresLicense) { 'PASS' } elseif (-not $mailboxTargetPolicy.Known -or -not $mailboxTargetPolicy.Eligible) { 'FAIL' } elseif (-not $licenseCapacity.ServicePlanDataAvailable) { 'UNKNOWN' } elseif ($servicePlanPass) { 'PASS' } else { 'FAIL' }
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'LICENSE-EXCHANGE-SERVICE-PLAN'; Category = 'Licensing'; Severity = if ($servicePlanPass) { 'Information' } else { 'Critical' }
                    Result = $servicePlanResult; IsBlocking = -not $servicePlanPass
                    ObservedValue = if (-not $mailboxTargetPolicy.RequiresLicense) { 'Not required for an unlicensed shared mailbox within policy' } elseif ($licenseCapacity -and $licenseCapacity.ExchangeServicePlanFound) { "$($licenseCapacity.ExchangeServicePlanName): $($licenseCapacity.ExchangeServicePlanStatus)" } elseif ($licenseCapacity -and -not $licenseCapacity.ServicePlanDataAvailable) { 'Exchange service plan evidence is unavailable' } else { 'No enabled mailbox-bearing Exchange service plan found' }
                    ExpectedValue = if ($mailboxTargetPolicy.RequiresLicense) { 'Enabled Exchange Online mailbox service plan in the target SKU' } else { 'Not required' }
                    EvidenceSource = if ($mailboxTargetPolicy.RequiresLicense) { "$graphSource target SKU service plans" } else { 'Configured shared mailbox policy' }
                    Message = if (-not $mailboxTargetPolicy.RequiresLicense) { 'No Exchange service plan is required for this unlicensed shared mailbox scenario.' } elseif ($servicePlanPass) { 'The target SKU contains an enabled Exchange Online mailbox service plan.' } elseif (-not $mailboxTargetPolicy.Known -or -not $mailboxTargetPolicy.Eligible) { $mailboxTargetPolicy.Message } elseif ($licenseCapacity -and -not $licenseCapacity.ServicePlanDataAvailable) { 'The target SKU service plans could not be verified.' } else { 'The target SKU does not expose an enabled Exchange Online mailbox service plan.' }
                    RecommendedAction = if ($servicePlanPass) { '' } elseif ($licenseCapacity -and -not $licenseCapacity.ServicePlanDataAvailable) { 'Refresh M365_Licenses_ServicePlans.csv or rerun in Live mode.' } else { 'Select a mailbox-eligible SKU with an enabled Exchange Online service plan.' }
                })
            }
        }

        Add-SemrIdentityAdvancedFindings -Findings $findings -Base $base -OnPrem $onPrem -Exo $exo -AcceptedDomains $acceptedDomains
        Add-SemrMailboxRiskFindings -Findings $findings -Base $base -Row $row -OnPrem $onPrem -Config $Config
        Add-SemrFlowAndSyncFindings -Findings $findings -Base $base -BatchRows $rows -OnPrem $onPrem -Exo $exo -Graph $graph

        if ($index -eq 1) {
            $globalBase = @{ RunId = $runId; EmailAddress = '' }
        $hybridEndpointResult = if (-not $hybrid.Available) { 'UNKNOWN' } elseif (-not $hybrid.EndpointFound) { 'FAIL' } elseif (-not $hybrid.ConnectivityTestAvailable) { 'UNKNOWN' } elseif ($hybrid.ConnectivitySuccess) { 'PASS' } else { 'FAIL' }
        $hybridEndpointBlocking = $hybridEndpointResult -eq 'FAIL'
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId = 'HYBRID-ENDPOINT'; Category = 'HybridConnectivity'; Severity = if ($hybridEndpointResult -eq 'FAIL') { 'Critical' } elseif ($hybridEndpointResult -eq 'UNKNOWN') { 'Warning' } else { 'Information' }
            Result = $hybridEndpointResult
            IsBlocking = $hybridEndpointBlocking; ObservedValue = "$($hybrid.EndpointName): $($hybrid.Message)"
            ExpectedValue = 'ExchangeRemoteMove endpoint exists and its connectivity test succeeds'; EvidenceSource = [string]$hybrid.Source; SourceTimestamp = $hybrid.SourceTimestamp
            Message = if ($hybridEndpointResult -eq 'PASS') { 'The migration endpoint connectivity test succeeded.' } elseif (-not $hybrid.Available) { 'Migration endpoint evidence is unavailable.' } elseif (-not $hybrid.EndpointFound) { 'No unique configured ExchangeRemoteMove endpoint could be resolved.' } elseif (-not $hybrid.ConnectivityTestAvailable) { 'The endpoint exists, but the connectivity cmdlet is unavailable; this is not treated as a failed endpoint.' } else { 'The migration endpoint connectivity test failed.' }
            RecommendedAction = if ($hybridEndpointResult -eq 'PASS') { '' } elseif (-not $hybrid.Available) { 'Restore Exchange Online endpoint cmdlets or run the final validation in Live mode.' } elseif (-not $hybrid.EndpointFound) { 'Select or repair the intended ExchangeRemoteMove endpoint.' } elseif (-not $hybrid.ConnectivityTestAvailable) { 'Run Test-MigrationServerAvailability from a supported delegated Exchange Online session before creating the batch.' } else { 'Validate the endpoint, MRSProxy, WSSecurity, certificate, DNS and HTTPS connectivity.' }
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-MRSPROXY';Category='HybridConnectivity';Severity=if(-not $hybrid.MrsProxyAvailable){'Warning'}elseif($hybrid.MrsProxyEnabled){'Information'}else{'Critical'};Result=if(-not $hybrid.MrsProxyAvailable){'UNKNOWN'}elseif($hybrid.MrsProxyEnabled){'PASS'}else{'FAIL'};IsBlocking=$hybrid.MrsProxyAvailable -and -not $hybrid.MrsProxyEnabled
            ObservedValue=$hybrid.MrsProxyMessage;ExpectedValue='MRSProxy enabled on a published EWS virtual directory';EvidenceSource=$hybrid.MrsProxySource;SourceTimestamp=$hybrid.MrsProxySourceTimestamp;Message=$hybrid.MrsProxyMessage;RecommendedAction=if(-not $hybrid.MrsProxyAvailable){'Validate MRSProxy live before creating the batch.'}elseif(-not $hybrid.MrsProxyEnabled){'Enable and publish MRSProxy on the intended EWS endpoint.'}else{''}
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-CERTIFICATE-EXPIRY';Category='HybridConnectivity';Severity=if(-not $hybrid.CertificateAvailable){'Warning'}elseif($hybrid.CertificateHealthy){'Information'}else{'Critical'};Result=if(-not $hybrid.CertificateAvailable){'UNKNOWN'}elseif($hybrid.CertificateHealthy){'PASS'}else{'FAIL'};IsBlocking=$hybrid.CertificateAvailable -and -not $hybrid.CertificateHealthy
            ObservedValue=$hybrid.CertificateMessage;ExpectedValue='Valid IIS hybrid certificate with at least 60 days remaining';EvidenceSource=$hybrid.CertificateSource;SourceTimestamp=$hybrid.CertificateSourceTimestamp;Message=$hybrid.CertificateMessage;RecommendedAction=if(-not $hybrid.CertificateAvailable){'Collect Exchange certificate evidence before migration.'}elseif(-not $hybrid.CertificateHealthy){'Renew and deploy the hybrid IIS certificate before migration.'}else{''}
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-ENDPOINT-CAPACITY';Category='HybridConnectivity';Severity=if(-not $hybrid.CapacityAvailable){'Warning'}elseif($hybrid.CapacityHealthy){'Information'}else{'Warning'};Result=if(-not $hybrid.CapacityAvailable){'UNKNOWN'}elseif($hybrid.CapacityHealthy){'PASS'}else{'WARN'};IsBlocking=$false
            ObservedValue=$hybrid.CapacityMessage;ExpectedValue='Active migration load below advisory threshold';EvidenceSource=$hybrid.CapacitySource;SourceTimestamp=$hybrid.CapacitySourceTimestamp;Message=$hybrid.CapacityMessage;RecommendedAction=if(-not $hybrid.CapacityAvailable){'Review current migration load before starting the batch.'}elseif(-not $hybrid.CapacityHealthy){'Reduce or stagger concurrent migration waves.'}else{''}
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-AUTODISCOVER-OAUTH';Category='HybridConnectivity';Severity=if($hybrid.OAuthHealthy){'Information'}else{'Warning'};Result=if(-not $hybrid.OAuthAvailable){'UNKNOWN'}elseif($hybrid.OAuthHealthy){'PASS'}else{'WARN'};IsBlocking=$false
            ObservedValue=$hybrid.OAuthMessage;ExpectedValue='OAuth enabled and an enabled intra-organization connector for hybrid collaboration features';EvidenceSource=$hybrid.OAuthSource;SourceTimestamp=$hybrid.OAuthSourceTimestamp;Message=if($hybrid.OAuthHealthy){$hybrid.OAuthMessage}else{"$($hybrid.OAuthMessage) This does not block ExchangeRemoteMove migrations."};RecommendedAction=if(-not $hybrid.OAuthAvailable){'Validate hybrid OAuth and Autodiscover for free/busy and other hybrid collaboration features.'}elseif(-not $hybrid.OAuthHealthy){'Repair hybrid OAuth and the intra-organization connector for hybrid collaboration features; migration readiness remains non-blocking.'}else{''}
        })
        $entraCoreHealthy = $entraConnect.Available -and $entraConnect.SyncCycleEnabled -and -not $entraConnect.SchedulerSuspended
        $entraLastSyncStale = $entraCoreHealthy -and -not [bool]$entraConnect.LastSyncFresh
        $entraResult = if (-not $entraConnect.Available) { 'UNKNOWN' } elseif (-not $entraCoreHealthy) { 'FAIL' } elseif ($entraLastSyncStale) { 'WARN' } else { 'PASS' }
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId = 'ENTRA-CONNECT-SCHEDULER'; Category = 'EntraConnect'; Severity = if ($entraResult -eq 'FAIL' -or $entraResult -eq 'UNKNOWN') { 'Critical' } elseif ($entraResult -eq 'WARN') { 'Warning' } else { 'Information' }
            Result = $entraResult; IsBlocking = $entraResult -in @('FAIL','UNKNOWN')
            ObservedValue = $entraConnect.Message; ExpectedValue = "Synchronization enabled and healthy; last sync no older than $([double]$Config['EntraConnectHealth']['MaximumLastSyncAgeMinutes']) minutes"; EvidenceSource = [string]$entraConnect.Source; SourceTimestamp = $entraConnect.SourceTimestamp
            Message = if (-not $entraConnect.Available) { 'Microsoft Entra Connect sync health could not be collected live or from the tenant cache.' } elseif ($entraLastSyncStale) { "The Entra Connect scheduler is enabled, but the last synchronization is $($entraConnect.LastSyncAgeMinutes) minutes old." } elseif ($entraCoreHealthy) { "Microsoft Entra Connect scheduler and last-sync freshness are healthy from $($entraConnect.Source)." } else { 'Microsoft Entra Connect synchronization is disabled or suspended.' }
            RecommendedAction = if (-not $entraConnect.Available) { 'Run on an Entra Connect server with ADSync cmdlets or refresh M365_Entra_AzureADConnect_SyncHealth.csv, then rerun.' } elseif ($entraLastSyncStale) { 'Investigate the synchronization backlog or scheduler history before starting another migration wave.' } elseif (-not $entraCoreHealthy) { 'Restore the Entra Connect scheduler and resolve synchronization health failures before migration.' } else { '' }
        })
        }
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

    $sourceInitialization | Add-Member -NotePropertyName ActiveDirectoryFallbackUsed -NotePropertyValue ([bool]$script:ActiveDirectoryFallbackUsed) -Force
    $sourceInitialization | Add-Member -NotePropertyName ActiveDirectoryFallbackReasons -NotePropertyValue @($script:ActiveDirectoryFallbackReasons) -Force

    $csvSources = @(Get-SemrCsvSourceInventory -Config $Config -Batch $Batch -EntraConnect $entraConnect -AssessmentCompleted)
    $summary = [System.Collections.Generic.List[object]]::new()
    $summaryRows = @($rows | Group-Object { ([string]$_.EmailAddress).ToLowerInvariant() } | ForEach-Object { $_.Group[0] })
    $globalBlocking = @($globalFindings | Where-Object { $_.IsBlocking -and $_.Result -in @('FAIL','UNKNOWN') })
    $globalWarnings = @($globalFindings | Where-Object Result -EQ 'WARN')
    $globalUnknown = @($globalFindings | Where-Object Result -EQ 'UNKNOWN')
    foreach ($row in $summaryRows) {
        $email = [string]$row.EmailAddress
        $mailFindings = @($findings | Where-Object EmailAddress -EQ $email)
        $mailboxBlocking = @($mailFindings | Where-Object { $_.IsBlocking -and $_.Result -in @('FAIL', 'UNKNOWN') })
        $blocking = @($mailboxBlocking) + @($globalBlocking)
        $warnings = @($mailFindings | Where-Object Result -EQ 'WARN')
        $unknown = @($mailFindings | Where-Object Result -EQ 'UNKNOWN')
        $decision = if ($blocking.Count -gt 0) { 'NO-GO' } elseif ($unknown.Count -gt 0 -or $warnings.Count -gt 0 -or $globalUnknown.Count -gt 0 -or $globalWarnings.Count -gt 0) { 'GO-WARNING' } else { 'GO' }
        $blockingCodes = @($blocking | ForEach-Object { $_.CheckId } | Sort-Object -Unique)
        $actionableFindings = if ($blocking.Count -gt 0) { @($blocking) } else { @($unknown) + @($warnings) + @($globalUnknown) + @($globalWarnings) }
        $recommended = @($actionableFindings | ForEach-Object { $_.RecommendedAction } | Where-Object { $_ } | Sort-Object -Unique)
        [void]$summary.Add([pscustomobject][ordered]@{
            RunId = $runId
            BatchName = [System.IO.Path]::GetFileNameWithoutExtension($Batch.Path)
            AssessmentPhase = [string]$Config['AssessmentPhase']
            EmailAddress = $email
            Decision = $decision
            BlockingCount = $blocking.Count
            MailboxBlockingCount = $mailboxBlocking.Count
            GlobalBlockingCount = $globalBlocking.Count
            WarningCount = $warnings.Count + $globalWarnings.Count
            MailboxWarningCount = $warnings.Count
            GlobalWarningCount = $globalWarnings.Count
            UnknownCount = $unknown.Count + $globalUnknown.Count
            MailboxUnknownCount = $unknown.Count
            GlobalUnknownCount = $globalUnknown.Count
            DataCoverage = if (@($blocking | Where-Object CheckId -match 'SOURCE').Count -gt 0) { 'Incomplete' } elseif ($unknown.Count -gt 0 -or $globalUnknown.Count -gt 0) { 'Partial' } else { 'Complete' }
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
        GlobalFindings = @($globalFindings)
        PermissionsBaseline = @($permissionRows)
        Evidence = @($evidenceRows)
        Hybrid = $hybrid
        EntraConnect = $entraConnect
        CsvSources = @($csvSources)
        CheckOptions = @(Get-SemrCheckCatalog | ForEach-Object {
            [pscustomobject][ordered]@{
                CheckId = $_.CheckId; Category = $_.Category; Name = $_.Name
                Mandatory = $_.Mandatory; Enabled = Test-SemrCheckEnabled -CheckId $_.CheckId
                Description = $_.Description
            }
        })
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
    $globalFindingsPath = Join-Path $runFolder 'Global-Findings.csv'
    $permissionsPath = Join-Path $runFolder 'Permissions-Baseline.csv'
    $evidencePath = Join-Path $runFolder 'Evidence.csv'
    $csvSourcesPath = Join-Path $runFolder 'Csv-Sources.csv'
    $checkOptionsPath = Join-Path $runFolder 'Check-Options.csv'

    Export-SemrCsvFile -Data @($Assessment.Summary) -Path $summaryPath -Columns @('RunId','BatchName','AssessmentPhase','EmailAddress','Decision','BlockingCount','MailboxBlockingCount','GlobalBlockingCount','WarningCount','MailboxWarningCount','GlobalWarningCount','UnknownCount','MailboxUnknownCount','GlobalUnknownCount','DataCoverage','BlockingCodes','RecommendedAction','CheckedAt')
    Export-SemrCsvFile -Data @($Assessment.Findings) -Path $findingsPath -Columns @('RunId','EmailAddress','CheckId','Category','Severity','Result','IsBlocking','ObservedValue','ExpectedValue','EvidenceSource','SourceTimestamp','Message','RecommendedAction')
    Export-SemrCsvFile -Data @($Assessment.GlobalFindings) -Path $globalFindingsPath -Columns @('RunId','EmailAddress','CheckId','Category','Severity','Result','IsBlocking','ObservedValue','ExpectedValue','EvidenceSource','SourceTimestamp','Message','RecommendedAction')
    Export-SemrCsvFile -Data @($Assessment.PermissionsBaseline) -Path $permissionsPath -Columns @('RunId','EmailAddress','PermissionType','Delegate','IsInherited','Source','CapturedAt')
    Export-SemrCsvFile -Data @($Assessment.Evidence) -Path $evidencePath -Columns @('RunId','EmailAddress','AdUserCount','OnPremMailboxCount','OnPremRemoteMailboxCount','ExoRecipientCount','ExoMailboxCount','GraphUserCount','PermissionCount','CollectedAt')
    Export-SemrCsvFile -Data @($Assessment.CsvSources) -Path $csvSourcesPath -Columns @('FileName','Category','ExpectedUse','Path','Present','Fresh','Used','Status','LastWriteTime','AgeHours','Details')
    Export-SemrCsvFile -Data @($Assessment.CheckOptions) -Path $checkOptionsPath -Columns @('CheckId','Category','Name','Mandatory','Enabled','Description')

    return [pscustomobject]@{
        RunFolder = $runFolder
        SummaryPath = $summaryPath
        FindingsPath = $findingsPath
        GlobalFindingsPath = $globalFindingsPath
        PermissionsPath = $permissionsPath
        EvidencePath = $evidencePath
        CsvSourcesPath = $csvSourcesPath
        CheckOptionsPath = $checkOptionsPath
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
    'Get-SemrCheckCatalog',
    'Get-SemrConfig',
    'Get-SemrConnectionState',
    'Connect-SemrActiveDirectory',
    'Connect-SemrOnPremisesExchange',
    'Connect-SemrExchangeOnline',
    'Connect-SemrMicrosoftGraph',
    'Get-SemrMicrosoftGraphModuleState',
    'Install-SemrMicrosoftGraphModule',
    'Disconnect-SemrSession',
    'Import-SemrBatchCsv',
    'Get-SemrCsvSourceInventory',
    'Get-SemrMigrationEndpointOption',
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVoX52QbaUAxPa
# 661X50FI/o9SvU+nBpFPz5FDuOSYpqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINJLRdlD2gZ6lkHFcS/EziUiEC0X89UcMZbevvqa4SWuMA0GCSqG
# SIb3DQEBAQUABIIBgHZeS4BnU0oJ3IuVEv/FmBkqF5HnPbYtf1Q4oc30ExVnCa22
# Tqp9EoPFanD/ORl3uNGCAvIjf9pFogbWKzhBID9Zt1TkzFirnlbNwz5rmrJsX5nL
# /MiuJtFgUXTeI/LVJySkGmH59eOupZkapG2lrDvmScfMZ38esUeNfpiEKA/EL6U3
# xo27LJGY7KT69Zox76lHlOB/TczgJw4NewY2KAMaTGrNs7l+2hDdXfnBt4jgOvMw
# EHP1CjJoDPempjVC5LLSmmg9OE3vM1hb9ZdFBaU9gnycpJ0AyDoFH/QaNu/WDGSL
# cXO1+0SeYR4/eE3q+1WwYWxxWxuSuGs8OAMt70o9aW3m5X6aHwhWJI1UBpStlDap
# 9A2IhsgiU//qkfUWtUy5hlv7Ur4l0fO7+h1AJcaFYnq0up2uRe7KG0eWj++cWAyi
# OmzLgZ8qRwQOPUnP2Buvihmw2AnUpqN60d6XjPjqZ0jr4pXUszeD27sEqymU7VSB
# KwLGZBGfilvQI9cXBKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTgxNTE5
# NDhaMC8GCSqGSIb3DQEJBDEiBCDPrnCLDK+dX+fqhw8x3Oq6E8StoDfARrx9EFIb
# Ldb8zTANBgkqhkiG9w0BAQEFAASCAgA78HVKAGQoSK/lW1CMpwYDV2rJtUqIvbjf
# ttsaXNZ28YUGCtv4YJQdQxui93D8a6hxiLZSDIa981KfderxG8bXkp0UVKU1I3Kv
# 77WyZoOAxxJNO/lmrcv6NTNqN8LxVWDtc4q5XaGuohbuaphFF55Jv6FKin9+eAw4
# vrqXYubRfHXSwjpa+ZwfBarJ9m8yfKf06do4nlK15ip7ravxT8x9xuR97Sn49a1e
# z1l2CIV43/k9tiiHUoH7TqFOgWbzwZxbtqbB3+9OmbEVA8BFtFCYKl9Jb62JnS/U
# TSZP3g0xPO7mEzot6Nk8jGhIZGopAsaBgEYxfOisHNzES4FEvDDCtPICy2PwAOVS
# Xd1FZL5YYV19e3/aqAvCgYxTuchIL4bjq0DsWFkUroOAQtdPhfDV66eCorbAwCCh
# zPHpZtGBuEGykumnT3vnOUTQgz0Jbx9+aw64QGXNooxJAGbL2lDzJcFwAtxOlK5U
# WvpAwMjtwmV2HddoVQdMcY7V6A31JP49SVeweQLGQInfOTzRLHqHP4BY4k6Tu198
# rFytFRqzErEgCkLe1EFvE3bujzLm3OPhfzPyJ0OfWkOzl/IyU4FkgrqNt/BbN0Fu
# bb44o7MAr+OKVLmucw6XfmYI0be1l/mBlm7bO+b8zJxg8Jx6yo3ZvXmG9pg1w053
# BnUvMiUiXA==
# SIG # End signature block
