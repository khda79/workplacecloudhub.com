Set-StrictMode -Version 2.0

$script:SemrVersion = '1.11.4'
$script:ActiveDirectoryDomains = @()
$script:Exchange2016EvidenceByEmail = @{}
$script:Exchange2016HybridEvidence = $null
$script:Exchange2016WorkerMessage = ''
$script:Exchange2016WorkerCollectedAt = $null
$script:Exchange2016PreflightAttempted = $false
$script:GraphEvidenceByEmail = @{}
$script:GraphSubscribedSkus = @()
$script:GraphSubscribedSkuError = ''
$script:GraphOrganization = $null
$script:GraphOrganizationError = ''
$script:CurrentSourceTimestamps = @{}
$script:BatchActiveDirectoryEvidenceByEmail = @{}
$script:BatchDelegateIdentityToEmail = @{}
$script:BatchEmailSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:AllMigrationUsers = @()
$script:MigrationUsersLoaded = $false
$script:MigrationUsersByEmail = @{}
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
    $mandatory = @('CSV-EMPTY-IDENTITY','CSV-SMTP-FORMAT','CSV-DUPLICATE','AD-SOURCE','ONPREM-SOURCE','EXO-SOURCE','GRAPH-SOURCE','ENTRA-CONNECT-SCHEDULER')
    $definitions = @(
        @('CSV-EMPTY-IDENTITY','CSV','Mailbox identity is present','Reject empty mailbox identities.'),
        @('CSV-SMTP-FORMAT','CSV','SMTP syntax','Validate mailbox identity SMTP syntax.'),
        @('CSV-DUPLICATE','CSV','Batch duplicates','Reject duplicate mailbox rows.'),
        @('CSV-COLUMNS','CSV','Recognized columns','Report columns not interpreted by the application.'),
        @('CSV-MAILBOX-TYPE','CSV','Migration mailbox type','Validate PrimaryOnly, ArchiveOnly or PrimaryAndArchive.'),
        @('CSV-BADITEMLIMIT','CSV','Bad item limit','Validate the optional bad item limit.'),
        @('CSV-LARGEITEMLIMIT','CSV','Large item limit','Validate the optional large item limit.'),
        @('AD-SOURCE','Active Directory','AD source availability','Require live Active Directory evidence across the complete forest.'),
        @('AD-IDENTITY-UNIQUE','Active Directory','Unique AD identity','Require exactly one matching AD user.'),
        @('AD-ACCOUNT-ENABLED','Active Directory','AD account enabled','Report disabled source accounts.'),
        @('PROXY-SMTP-GLOBAL-UNIQUE','Hybrid identity','Global SMTP uniqueness','Detect SMTP ownership conflicts across directory recipients.'),
        @('PROXY-INTERNAL-DUPLICATE','Hybrid identity','Internal proxy duplicates','Detect duplicate or malformed proxyAddresses values.'),
        @('TARGET-ADDRESS-GLOBAL-UNIQUE','Hybrid identity','targetAddress uniqueness','Detect duplicate or invalid target routing addresses.'),
        @('X500-LEGACYEXCHANGEDN','Hybrid identity','X500 preservation','Verify LegacyExchangeDN is preserved as an X500 proxy.'),
        @('SMTP-ACCEPTED-DOMAIN','Hybrid identity','Accepted SMTP domains','Verify every mailbox SMTP domain is accepted in Exchange Online.'),
        @('ONPREM-SOURCE','Exchange on-premises','On-premises source availability','Require live Exchange on-premises evidence with ViewEntireForest enabled.'),
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
        @('EXO-SOURCE','Exchange Online','EXO source availability','Require live Exchange Online evidence.'),
        @('EXO-RECIPIENT-STATE','Exchange Online','Cloud recipient state','Require one synchronized MailUser and no active cloud mailbox.'),
        @('EXO-SOFT-DELETED-CONFLICT','Exchange Online','Soft-deleted conflicts','Detect soft-deleted or inactive mailbox conflicts.'),
        @('EXO-EXISTING-MOVE','Migration','Existing move objects','Detect existing migration users and move requests.'),
        @('MOVE-HISTORY','Migration','Current migration failure state','Report failed, suspended or abandoned current migration objects.'),
        @('GRAPH-SOURCE','Microsoft Graph','Graph source availability','Require live Microsoft Graph evidence.'),
        @('GRAPH-USER-STATE','Microsoft Graph','Unique Entra user','Require exactly one matching Entra user.'),
        @('GRAPH-DIRSYNC','Microsoft Graph','Directory synchronization','Verify the user is synchronized from on-premises.'),
        @('ENTRA-UPN-VERIFIED-DOMAIN','Microsoft Graph','Verified UPN domain','Verify the Entra user UPN domain is verified in the tenant.'),
        @('ENTRA-OBJECT-SYNC-ERROR','Microsoft Graph','Object synchronization errors','Detect stale synchronization and identity anchor issues.'),
        @('LICENSE-PRE-MIGRATION','Licensing','Pre-migration licenses','Report licenses already assigned before migration.'),
        @('LICENSE-ASSIGNED-MAILBOX-QUOTA','Licensing','Assigned license mailbox quota','Compare mailbox size with the quota of the currently assigned mailbox license.'),
        @('LICENSE-USAGE-LOCATION','Licensing','Usage location','Verify UsageLocation is populated.'),
        @('LICENSE-CAPACITY','Licensing','Target SKU capacity','Require available target SKU capacity.'),
        @('LICENSE-BATCH-CAPACITY','Licensing','Batch target SKU capacity','Require enough available target SKU licenses for the complete batch.'),
        @('LICENSE-EXCHANGE-SERVICE-PLAN','Licensing','Exchange service plan','Require a mailbox-bearing Exchange service plan.'),
        @('HYBRID-ENDPOINT','Hybrid connectivity','Migration endpoint','Test the ExchangeRemoteMove endpoint.'),
        @('HYBRID-MRSPROXY','Hybrid connectivity','MRSProxy readiness','Verify MRSProxy and EWS migration readiness.'),
        @('HYBRID-CERTIFICATE-EXPIRY','Hybrid connectivity','Hybrid certificate expiry','Detect missing or expiring hybrid certificates.'),
        @('HYBRID-ENDPOINT-CAPACITY','Hybrid connectivity','Tenant migration load','Report tenant-wide active migration load.'),
        @('HYBRID-MIGRATION-BACKLOG','Hybrid connectivity','Migration failure backlog','Report failed, stopped or corrupted migration users separately from active migration load.'),
        @('HYBRID-AUTODISCOVER-OAUTH','Hybrid connectivity','Autodiscover and OAuth','Validate hybrid organization relationship and OAuth configuration.'),
        @('ENTRA-CONNECT-SCHEDULER','Microsoft Entra','Tenant synchronization health','Require enabled and recent tenant synchronization.'),
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
    foreach ($deprecatedKey in @('Mode', 'Cache', 'Tenant', 'OnPremises', 'ExchangeOnPremises', 'EntraConnect', 'SmartM365', 'EvidenceSources', 'Authentication')) {
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

    # Internal compatibility marker. Live is the only supported evidence mode.
    $runtime['Mode'] = 'Live'
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

    $runtime['_RuntimePath'] = $Path
    $runtime['_AddedKeys'] = $added
    $runtime['_TenantProfileKey'] = [string]$tenantProfile['ProfileKey']
    $runtime['_TenantId'] = $tenantId
    $runtime['_RemoteRoutingDomain'] = [string]$tenantProfile['RemoteRoutingDomain']
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
    $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8
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
        $script:ConnectionState.ActiveDirectory = $true
    return Get-SemrConnectionState
}

function Get-SemrWindowsPowerShellPath {
    [CmdletBinding()]
    param()

    $path = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Windows PowerShell 5.1 was not found at '$path'."
    }
    return $path
}

function Get-SemrExchange2016WorkerPath {
    [CmdletBinding()]
    param()

    $path = Join-Path $PSScriptRoot 'SmartM365-ExchangeMigrationReadiness-Exchange2016Worker.ps1'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "The local Exchange 2016 worker is missing: $path"
    }
    return $path
}

function Test-SemrExchange2016WorkerSerialization {
    [CmdletBinding()]
    param()

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = Get-SemrWindowsPowerShellPath
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8
    foreach ($argument in @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', (Get-SemrExchange2016WorkerPath), '-SelfTest')) {
        [void]$startInfo.ArgumentList.Add([string]$argument)
    }
    $process = [Diagnostics.Process]::Start($startInfo)
    if (-not $process) { throw 'The Windows PowerShell 5.1 Exchange worker self-test could not be started.' }
    try {
        if (-not $process.WaitForExit(30000)) {
            try { $process.Kill() } catch { $null = $_ }
            throw 'The Windows PowerShell 5.1 Exchange worker self-test exceeded 30 seconds.'
        }
        $standardOutput = $process.StandardOutput.ReadToEnd().Trim()
        $standardError = $process.StandardError.ReadToEnd().Trim()
        if ($process.ExitCode -ne 0 -or $standardOutput -notmatch '^SELFTEST_OK\|') {
            $detail = if ($standardError) { $standardError } elseif ($standardOutput) { $standardOutput } else { "Worker exit code $($process.ExitCode)." }
            throw "Windows PowerShell 5.1 Exchange worker serialization self-test failed: $detail"
        }
        return $standardOutput
    }
    finally {
        $process.Dispose()
    }
}
function Connect-SemrOnPremisesExchange {
    [CmdletBinding()]
    param()

    $script:Exchange2016PreflightAttempted = $true
    [void](Test-SemrExchange2016WorkerSerialization)
    $powershellPath = Get-SemrWindowsPowerShellPath
    $workerPath = Get-SemrExchange2016WorkerPath
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $powershellPath
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8
    foreach ($argument in @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $workerPath, '-ValidateOnly')) {
        [void]$startInfo.ArgumentList.Add([string]$argument)
    }

    $process = [Diagnostics.Process]::Start($startInfo)
    if (-not $process) {
        throw 'The local Windows PowerShell 5.1 Exchange 2016 validation process could not be started.'
    }
    try {
        if (-not $process.WaitForExit(120000)) {
            try { $process.Kill() } catch { $null = $_ }
            throw 'The local Exchange 2016 preflight exceeded 120 seconds.'
        }
        $standardOutput = $process.StandardOutput.ReadToEnd().Trim()
        $standardError = $process.StandardError.ReadToEnd().Trim()
        if ($process.ExitCode -ne 0 -or $standardOutput -notmatch '^VALIDATION_OK\|') {
            $detail = if ($standardError) { $standardError } elseif ($standardOutput) { $standardOutput } else { "Worker exit code $($process.ExitCode)." }
            throw "Local Exchange 2016 Management Shell preflight failed: $detail"
        }
        $parts = @($standardOutput.Split('|'))
        $computerName = if ($parts.Count -gt 2) { $parts[2] } else { $env:COMPUTERNAME }
        $psVersion = if ($parts.Count -gt 3) { $parts[3] } else { '5.1' }
        $script:Exchange2016WorkerMessage = "Local Exchange 2016 Management Shell validated on $computerName with Windows PowerShell $psVersion; ViewEntireForest enabled."
        $script:ConnectionState.OnPremisesExchange = $true
        return Get-SemrConnectionState
    }
    catch {
        $script:ConnectionState.OnPremisesExchange = $false
        $script:Exchange2016WorkerMessage = "Local Exchange 2016 Management Shell unavailable. $($_.Exception.Message)"
        throw
    }
    finally {
        $process.Dispose()
    }
}

function Initialize-SemrExchange2016Evidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$EmailAddresses,
        [scriptblock]$ProgressCallback,
        [scriptblock]$CancellationCheck,
        [string]$DiagnosticsDirectory = '',
        [ValidateRange(1,50)][int]$SmtpUniquenessBatchSize = 25,
        [ValidateRange(5,300)][int]$SmtpUniquenessBatchTimeoutSeconds = 60
    )

    if (-not $script:ConnectionState.OnPremisesExchange) {
        return [pscustomobject]@{ Available = $false; Message = 'Local Exchange 2016 Management Shell preflight has not succeeded.'; MailboxCount = 0 }
    }
    $powershellPath = Get-SemrWindowsPowerShellPath
    $workerPath = Get-SemrExchange2016WorkerPath
    $runtimeRoot = Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-ExchangeMigrationReadiness\Exchange2016-{0}" -f [guid]::NewGuid().ToString('N'))
    $inputPath = Join-Path $runtimeRoot 'mailboxes.clixml'
    $outputPath = Join-Path $runtimeRoot 'evidence.clixml'
    $errorPath = Join-Path $runtimeRoot 'error.txt'
    $progressPath = Join-Path $runtimeRoot 'progress.txt'
    $cancelPath = Join-Path $runtimeRoot 'cancel.requested'
    $process = $null
    try {
        [void](New-Item -ItemType Directory -Path $runtimeRoot -Force)
        $workerInput = [pscustomobject][ordered]@{
            EmailAddresses = @($EmailAddresses | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
            EnabledChecks = @(Get-SemrCheckCatalog | Where-Object { Test-SemrCheckEnabled -CheckId $_.CheckId } | ForEach-Object { [string]$_.CheckId })
            SmtpUniquenessBatchSize = $SmtpUniquenessBatchSize
            SmtpUniquenessBatchTimeoutSeconds = $SmtpUniquenessBatchTimeoutSeconds
        }
        $workerInput | Export-Clixml -LiteralPath $inputPath -Depth 4 -Force

        $startInfo = [Diagnostics.ProcessStartInfo]::new()
        $startInfo.FileName = $powershellPath
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        foreach ($argument in @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $workerPath, '-InputPath', $inputPath, '-OutputPath', $outputPath, '-ErrorPath', $errorPath, '-ProgressPath', $progressPath, '-CancelPath', $cancelPath, '-DiagnosticsDirectory', $DiagnosticsDirectory)) {
            [void]$startInfo.ArgumentList.Add([string]$argument)
        }
        $process = [Diagnostics.Process]::Start($startInfo)
        if (-not $process) { throw 'The local Exchange 2016 evidence worker could not be started.' }

        $lastProgress = ''
        $cancelRequestedAt = $null
        while (-not $process.WaitForExit(250)) {
            if ($CancellationCheck -and (& $CancellationCheck)) {
                if (-not $cancelRequestedAt) {
                    $cancelRequestedAt = Get-Date
                    [IO.File]::WriteAllText($cancelPath, 'Cancellation requested', [Text.UTF8Encoding]::new($false))
                    if ($ProgressCallback) { & $ProgressCallback 0 0 'Exchange 2016 cancellation requested; stopping the worker...' }
                }
                elseif (((Get-Date) - $cancelRequestedAt).TotalSeconds -ge 3) {
                    try { $process.Kill($true) } catch { try { $process.Kill() } catch { $null = $_ } }
                }
            }
            if ($ProgressCallback -and (Test-Path -LiteralPath $progressPath -PathType Leaf)) {
                try {
                    $progress = [IO.File]::ReadAllText($progressPath)
                    if ($progress -and $progress -ne $lastProgress) {
                        $lastProgress = $progress
                        $parts = @($progress.Split('|', 3))
                        if ($parts.Count -eq 3) {
                            $current = 0; $total = 0
                            [void][int]::TryParse($parts[0], [ref]$current)
                            [void][int]::TryParse($parts[1], [ref]$total)
                            & $ProgressCallback $current $total ("Exchange 2016 - {0}" -f $parts[2])
                        }
                        else { & $ProgressCallback 0 0 'Collecting local Exchange 2016 evidence...' }
                    }
                }
                catch { $null = $_ }
            }
        }
        if ($cancelRequestedAt -or ($CancellationCheck -and (& $CancellationCheck))) {
            throw [OperationCanceledException]::new('Local Exchange 2016 evidence collection was cancelled by the operator.')
        }
        if ($process.ExitCode -ne 0) {
            $workerError = if (Test-Path -LiteralPath $errorPath -PathType Leaf) { [IO.File]::ReadAllText($errorPath) } else { "Worker exit code $($process.ExitCode)." }
            throw "Local Exchange 2016 evidence collection failed: $workerError"
        }
        if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
            throw 'The local Exchange 2016 worker completed without returning evidence.'
        }

        $workerResult = Import-Clixml -LiteralPath $outputPath
        if ($ProgressCallback) { $completeCount = @($workerResult.Evidence).Count; & $ProgressCallback $completeCount $completeCount 'Exchange 2016 - Complete' }
        $script:Exchange2016EvidenceByEmail = @{}
        foreach ($entry in @($workerResult.Evidence)) {
            $key = ([string]$entry.EmailAddress).Trim().ToLowerInvariant()
            if ($key) { $script:Exchange2016EvidenceByEmail[$key] = $entry }
        }
        $script:Exchange2016HybridEvidence = $workerResult.Hybrid
        $script:Exchange2016WorkerCollectedAt = $workerResult.CollectedAt
        $script:Exchange2016WorkerMessage = "Collected $($script:Exchange2016EvidenceByEmail.Count) mailbox evidence set(s), $($workerResult.MailboxObjectCount) mailbox object(s) and $($workerResult.PermissionCount) delegated permission(s) in $($workerResult.DurationSeconds) second(s) through direct local Exchange 2016 cmdlets on $($workerResult.ComputerName) with Windows PowerShell $($workerResult.PowerShellVersion); ViewEntireForest enabled; per-mailbox errors=$($workerResult.ErrorCount); SMTP uniqueness candidates=$($workerResult.SmtpUniquenessCandidateAddressCount), batches=$($workerResult.SmtpUniquenessBatchCount), timeouts=$($workerResult.SmtpUniquenessTimeoutCount), query errors=$($workerResult.SmtpUniquenessErrorCount), duration=$($workerResult.SmtpUniquenessDurationSeconds) second(s); child logs=$($workerResult.SmtpUniquenessDiagnosticsDirectory)."
        return [pscustomobject]@{ Available = $true; Message = $script:Exchange2016WorkerMessage; MailboxCount = $script:Exchange2016EvidenceByEmail.Count; ErrorCount = [int]$workerResult.ErrorCount; DurationSeconds = $workerResult.DurationSeconds }
    }
    finally {
        if ($process) { $process.Dispose() }
        if (Test-Path -LiteralPath $runtimeRoot) { Remove-Item -LiteralPath $runtimeRoot -Recurse -Force -ErrorAction SilentlyContinue }
    }
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
        $script:GraphOrganization = $workerResult.Organization
        $script:GraphOrganizationError = [string]$workerResult.OrganizationError
        $script:ConnectionState.EntraConnect = $null -ne $script:GraphOrganization
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

    foreach ($key in @($script:ConnectionState.Keys)) {
        $script:ConnectionState[$key] = $false
    }
    $script:ActiveDirectoryDomains = @()
    $script:Exchange2016EvidenceByEmail = @{}
    $script:Exchange2016HybridEvidence = $null
    $script:Exchange2016WorkerMessage = ''
    $script:Exchange2016WorkerCollectedAt = $null
    $script:Exchange2016PreflightAttempted = $false
    $script:GraphEvidenceByEmail = @{}
    $script:GraphSubscribedSkus = @()
    $script:GraphSubscribedSkuError = ''
    $script:BatchActiveDirectoryEvidenceByEmail = @{}
    $script:BatchDelegateIdentityToEmail = @{}
    $script:BatchEmailSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $script:AllMigrationUsers = @()
    $script:MigrationUsersLoaded = $false
    $script:MigrationUsersByEmail = @{}
    $script:GraphOrganization = $null
    $script:GraphOrganizationError = ''
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

    if (-not $script:ConnectionState.ActiveDirectory) {
        try {
            Connect-SemrActiveDirectory | Out-Null
            $result.ActiveDirectoryMessage = "Live Active Directory forest connection succeeded. $($script:ActiveDirectoryDomains.Count) domain(s) will be searched."
        }
        catch {
            $result.ActiveDirectoryMessage = "Live Active Directory unavailable. The assessment will be INCOMPLETE. $($_.Exception.Message)"
        }
    }
    else {
        $result.ActiveDirectoryMessage = 'Live Active Directory was already connected.'
    }
    $result.ActiveDirectoryLive = [bool]$script:ConnectionState.ActiveDirectory

    if (-not $script:ConnectionState.OnPremisesExchange) {
        if ($script:Exchange2016PreflightAttempted) {
            $result.ExchangeOnPremisesMessage = "$($script:Exchange2016WorkerMessage) The assessment will be INCOMPLETE."
        }
        else {
            try {
                Connect-SemrOnPremisesExchange | Out-Null
                $result.ExchangeOnPremisesMessage = $script:Exchange2016WorkerMessage
            }
            catch {
                $result.ExchangeOnPremisesMessage = "Live Exchange on-premises unavailable. The assessment will be INCOMPLETE. $($_.Exception.Message)"
            }
        }
    }
    else {
        $result.ExchangeOnPremisesMessage = $script:Exchange2016WorkerMessage
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
        [Parameter(Mandatory)][AllowEmptyString()][string]$EmailAddress,
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
        $matchingTimestamps = [System.Collections.Generic.List[datetime]]::new()
        $evidenceSource = ([string]$Parameters.EvidenceSource).Trim().ToLowerInvariant()
        if ($evidenceSource -and $script:CurrentSourceTimestamps) {
            foreach ($sourceKey in @($script:CurrentSourceTimestamps.Keys)) {
                if ($evidenceSource.Contains(([string]$sourceKey).ToLowerInvariant())) {
                    $timestamp = $script:CurrentSourceTimestamps[$sourceKey]
                    if ($timestamp) {
                        try { [void]$matchingTimestamps.Add([datetime]$timestamp) } catch {}
                    }
                }
            }
        }
        $Parameters.SourceTimestamp = if ($matchingTimestamps.Count -gt 0) { @($matchingTimestamps | Sort-Object | Select-Object -First 1)[0] } else { $null }
    }
    [void]$List.Add((ConvertTo-SemrFinding @Parameters))
}

function Get-SemrMailboxDecision {
    param(
        [object[]]$MailboxFindings = @(),
        [object[]]$GlobalFindings = @()
    )

    $blockingFindings = @($MailboxFindings) + @($GlobalFindings | Where-Object IsBlocking)
    if (@($blockingFindings | Where-Object Result -EQ 'FAIL').Count -gt 0) {
        return 'NO-GO'
    }
    if (@($blockingFindings | Where-Object Result -EQ 'UNKNOWN').Count -gt 0) {
        return 'UNKNOWN'
    }
    if (@($MailboxFindings | Where-Object { $_.Result -in @('WARN','UNKNOWN') }).Count -gt 0) {
        return 'GO-WARNING'
    }
    return 'GO'
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

function Invoke-SemrCommandResult {
    param(
        [string]$CommandName,
        [hashtable]$Parameters = @{}
    )

    if ([string]::IsNullOrWhiteSpace($CommandName)) {
        return [pscustomobject]@{ Success = $false; Rows = @(); ErrorMessage = 'Required command is unavailable.' }
    }
    try {
        return [pscustomobject]@{
            Success = $true
            Rows = @(& $CommandName @Parameters -ErrorAction Stop)
            ErrorMessage = ''
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Rows = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function ConvertTo-SemrBoolean {
    param($Value)

    if ($Value -is [bool]) { return [bool]$Value }
    $text = ([string]$Value).Trim()
    return $text -match '^(?i:true|1|yes|oui)$'
}

function Get-SemrOnPremisesEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    $key = $EmailAddress.Trim().ToLowerInvariant()
    if ($script:ConnectionState.OnPremisesExchange -and $script:Exchange2016EvidenceByEmail.ContainsKey($key)) {
        return $script:Exchange2016EvidenceByEmail[$key]
    }
    return [pscustomobject]@{
        Available = $false
        Source = 'Local Exchange 2016 Management Shell unavailable'
        SourceTimestamp = $null
        Message = if ($script:Exchange2016WorkerMessage) { $script:Exchange2016WorkerMessage } else { 'No Exchange 2016 evidence was returned by the local Windows PowerShell 5.1 worker.' }
        Mailboxes = @()
        RemoteMailboxes = @()
        MailUsers = @()
        Recipients = @()
        RecipientLookupAvailable = $false
        Statistics = @()
        StatisticsAvailable = $false
        ArchiveStatistics = @()
        ArchiveStatisticsAvailable = $false
        Permissions = @()
        PermissionsAvailable = $false
        HoldDataAvailable = $false
        FolderStatistics = @()
        FolderStatisticsAvailable = $false
        InboxRules = @()
        InboxRulesAvailable = $false
        DatabaseHealth = @()
        DatabaseHealthAvailable = $false
        DatabaseHealthSource = 'Local Exchange 2016 Management Shell unavailable'
        DatabaseHealthSourceTimestamp = $null
        DeliveryRestrictionsAvailable = $false
        AddressConflicts = @()
    }
}

function Initialize-SemrBatchActiveDirectoryEvidence {
    param([Parameter(Mandatory)][string[]]$EmailAddresses)

    $emails = @(
        $EmailAddresses |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
    if ($emails.Count -eq 0) {
        return [pscustomobject]@{ Available = $true; Message = 'No mailbox identity requires Active Directory prefetch.'; DomainCount = 0; QueryCount = 0 }
    }

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
            return [pscustomobject]@{ Available = $false; Message = "Live Active Directory forest discovery failed: $($_.Exception.Message)"; DomainCount = 0; QueryCount = 0 }
        }
    }

    $usersByEmail = @{}
    foreach ($email in $emails) { $usersByEmail[$email] = [System.Collections.Generic.List[object]]::new() }
    $domainErrors = [System.Collections.Generic.List[string]]::new()
    $queryCount = 0
    $chunkSize = 25
    foreach ($domain in $domains) {
        for ($offset = 0; $offset -lt $emails.Count; $offset += $chunkSize) {
            $last = [math]::Min($offset + $chunkSize - 1, $emails.Count - 1)
            $chunk = @($emails[$offset..$last])
            $clauses = @(
                foreach ($email in $chunk) {
                    $escaped = $email.Replace("'", "''")
                    "(UserPrincipalName -eq '$escaped' -or mail -eq '$escaped' -or proxyAddresses -eq 'smtp:$escaped')"
                }
            )
            try {
                $queryCount++
                $domainUsers = @(
                    Get-ADUser -Server $domain -Filter ($clauses -join ' -or ') -Properties @(
                        'Enabled','UserPrincipalName','mail','proxyAddresses','targetAddress',
                        'mS-DS-ConsistencyGuid','ObjectGuid','whenChanged','SamAccountName','DistinguishedName'
                    ) -ErrorAction Stop
                )
                foreach ($domainUser in $domainUsers) {
                    $addresses = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    foreach ($address in @(
                        [string](Get-SemrPropertyValue -InputObject $domainUser -Names @('UserPrincipalName') -Default ''),
                        [string](Get-SemrPropertyValue -InputObject $domainUser -Names @('mail','PrimarySmtpAddress') -Default '')
                    )) {
                        if ($address) { [void]$addresses.Add($address.Trim()) }
                    }
                    foreach ($proxy in @(Get-SemrPropertyValue -InputObject $domainUser -Names @('proxyAddresses') -Default @())) {
                        $proxyText = ([string]$proxy).Trim()
                        if ($proxyText -match '^(?i:smtp):(.+)$') { [void]$addresses.Add($Matches[1]) }
                    }
                    foreach ($email in $chunk) {
                        if ($addresses.Contains($email)) { [void]$usersByEmail[$email].Add($domainUser) }
                    }
                }
            }
            catch {
                [void]$domainErrors.Add("${domain}: $($_.Exception.Message)")
                break
            }
        }
    }

    if ($domainErrors.Count -gt 0) {
        return [pscustomobject]@{
            Available = $false
            Message = "Live Active Directory forest batch search was incomplete ($($domainErrors.Count)/$($domains.Count) domain(s) failed: $($domainErrors -join '; '))."
            DomainCount = $domains.Count
            QueryCount = $queryCount
        }
    }

    $collectedAt = Get-Date
    foreach ($email in $emails) {
        $uniqueUsers = @(
            $usersByEmail[$email] |
                Group-Object {
                    $objectGuid = [string](Get-SemrPropertyValue -InputObject $_ -Names @('ObjectGuid') -Default '')
                    if ($objectGuid) { $objectGuid } else { [string](Get-SemrPropertyValue -InputObject $_ -Names @('DistinguishedName') -Default $_.UserPrincipalName) }
                } |
                ForEach-Object { $_.Group[0] }
        )
        $evidence = [pscustomobject]@{
            Available = $true
            Source = "Live Active Directory forest batch query ($($domains.Count) domains)"
            SourceTimestamp = $collectedAt
            Message = "Live Active Directory evidence collected with $queryCount batched forest queries."
            Users = $uniqueUsers
        }
        $script:BatchActiveDirectoryEvidenceByEmail[$email] = $evidence
        foreach ($adUser in $uniqueUsers) {
            foreach ($identity in @(
                $email,
                [string](Get-SemrPropertyValue -InputObject $adUser -Names @('SamAccountName') -Default ''),
                [string](Get-SemrPropertyValue -InputObject $adUser -Names @('UserPrincipalName') -Default ''),
                [string](Get-SemrPropertyValue -InputObject $adUser -Names @('Mail','PrimarySmtpAddress') -Default ''),
                [string](Get-SemrPropertyValue -InputObject $adUser -Names @('DistinguishedName') -Default '')
            )) {
                if (-not [string]::IsNullOrWhiteSpace($identity)) {
                    $script:BatchDelegateIdentityToEmail[$identity.Trim().ToLowerInvariant()] = $email
                }
            }
        }
    }

    return [pscustomobject]@{
        Available = $true
        Message = "Active Directory batch prefetch completed across $($domains.Count) domain(s) with $queryCount queries."
        DomainCount = $domains.Count
        QueryCount = $queryCount
    }
}


function Get-SemrActiveDirectoryEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)

    $evidenceKey = $EmailAddress.Trim().ToLowerInvariant()
    if ($script:BatchActiveDirectoryEvidenceByEmail.ContainsKey($evidenceKey)) {
        return $script:BatchActiveDirectoryEvidenceByEmail[$evidenceKey]
    }
    return [pscustomobject]@{
        Available = $false
        Source = 'Live Active Directory unavailable'
        SourceTimestamp = $null
        Message = if ($script:ConnectionState.ActiveDirectory) { 'Live Active Directory batch evidence is unavailable for this identity.' } else { 'Live Active Directory forest is not connected.' }
        Users = @()
    }
}
function Get-SemrExchangeOnlineEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)


    if (-not $script:ConnectionState.ExchangeOnline) {
        return [pscustomobject]@{
            Available = $false; Source = 'Exchange Online unavailable'; SourceTimestamp = $null
            Recipients = @(); RecipientDataAvailable = $false; RecipientLookupMethods = @(); RecipientLookupErrors = @()
            Mailboxes = @(); MailboxDataAvailable = $false; Statistics = @()
            SoftDeleted = @(); SoftDeletedAvailable = $false
            MigrationUsers = @(); MoveRequests = @(); MoveDataAvailable = $false
            MigrationStatistics = @(); MigrationReportDataAvailable = $false; MigrationReportText = ''
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
    $recipientType = if ($recipientList.Count -eq 1) { [string](Get-SemrPropertyValue -InputObject $recipientList[0] -Names @('RecipientTypeDetails','RecipientType') -Default '') } else { '' }
    if ($recipientList.Count -eq 1 -and $recipientType -match 'MailUser|RemoteUserMailbox') {
        $mailboxDataAvailable = $true
    }
    elseif (Test-SemrCommand -Name 'Get-EXOMailbox') {
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
    $migrationStatistics = [System.Collections.Generic.List[object]]::new()
    $migrationReportDataAvailable = $false
    $migrationReportText = ''
    $migrationKey = $EmailAddress.Trim().ToLowerInvariant()
    if ($script:MigrationUsersLoaded) {
        $migrationLookupAvailable = $true
        if ($script:MigrationUsersByEmail.ContainsKey($migrationKey)) {
            $migrationUsers = @($script:MigrationUsersByEmail[$migrationKey])
        }
    }
    elseif (Test-SemrCommand -Name 'Get-MigrationUser') {
        try { $migrationUsers = @(Get-MigrationUser -Identity $EmailAddress -ErrorAction Stop); $migrationLookupAvailable = $true }
        catch {
            if (Test-SemrObjectNotFoundError -Message $_.Exception.Message) { $migrationLookupAvailable = $true }
            else { [void]$lookupErrors.Add("Get-MigrationUser identity: $($_.Exception.Message)") }
        }
    }
    if ($migrationUsers.Count -eq 0) {
        $migrationReportDataAvailable = $migrationLookupAvailable
    }
    elseif (Test-SemrCommand -Name 'Get-MigrationUserStatistics') {
        foreach ($migrationUser in $migrationUsers) {
            try {
                $statisticsIdentity = Get-SemrPropertyValue -InputObject $migrationUser -Names @('Identity','EmailAddress') -Default $EmailAddress
                $migrationStatistic = Get-MigrationUserStatistics -Identity $statisticsIdentity -IncludeReport -ErrorAction Stop
                [void]$migrationStatistics.Add($migrationStatistic)
            }
            catch { [void]$lookupErrors.Add("Get-MigrationUserStatistics: $($_.Exception.Message)") }
        }
        $migrationReportDataAvailable = $migrationStatistics.Count -eq $migrationUsers.Count
        if ($migrationStatistics.Count -gt 0) {
            $migrationReportText = @($migrationStatistics | ConvertTo-Json -Depth 8 -Compress) -join ' '
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
        MigrationStatistics = @($migrationStatistics)
        MigrationReportDataAvailable = $migrationReportDataAvailable
        MigrationReportText = $migrationReportText
        MoveDataAvailable = $migrationLookupAvailable -and $moveRequestLookupAvailable
    }
}

function Get-SemrGraphEvidence {
    param([Parameter(Mandatory)][string]$EmailAddress)


    if (-not $script:ConnectionState.MicrosoftGraph) {
        return [pscustomobject]@{ Available = $false; Source = 'Live Microsoft Graph unavailable'; SourceTimestamp = $null; Users = @(); LicenseDetails = @(); QueryError = 'Microsoft Graph is not connected.' }
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
        Source = 'Unavailable'
        SourceTimestamp = $null
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
    if (-not $script:ConnectionState.MicrosoftGraph) {
        $result.Message = 'Microsoft Graph subscribed SKU data is unavailable.'
        return [pscustomobject]$result
    }
    if ($script:GraphSubscribedSkuError) {
        $result.Message = $script:GraphSubscribedSkuError
        return [pscustomobject]$result
    }

    $result.Available = $true
    $result.Source = 'Live Microsoft Graph subscribedSkus'
    $result.SourceTimestamp = if ($script:GraphOrganization) { Get-SemrPropertyValue -InputObject $script:GraphOrganization -Names @('CollectedAt') -Default (Get-Date) } else { Get-Date }
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

    return [pscustomobject]$result
}

function Get-SemrUpnVerifiedDomainEvidence {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$UserPrincipalName)

    $upnDomain = if ($UserPrincipalName -match '@(?<Domain>[^@]+)$') { $Matches.Domain.Trim().ToLowerInvariant() } else { '' }
    $organization = $script:GraphOrganization
    $verifiedDomainRows = if ($organization) { @(Get-SemrPropertyValue -InputObject $organization -Names @('VerifiedDomains') -Default @()) } else { @() }
    $verifiedDomainNames = @(
        $verifiedDomainRows |
            ForEach-Object { [string](Get-SemrPropertyValue -InputObject $_ -Names @('Name') -Default '') } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim().ToLowerInvariant() } |
            Sort-Object -Unique
    )
    $available = $null -ne $organization -and $verifiedDomainNames.Count -gt 0
    $verified = $available -and -not [string]::IsNullOrWhiteSpace($upnDomain) -and $verifiedDomainNames -contains $upnDomain
    $sourceTimestamp = if ($organization) { Get-SemrPropertyValue -InputObject $organization -Names @('CollectedAt') -Default $null } else { $null }
    return [pscustomobject][ordered]@{
        Available = $available
        Verified = $verified
        UserPrincipalName = $UserPrincipalName
        UpnDomain = $upnDomain
        VerifiedDomains = $verifiedDomainNames
        Source = if ($organization) { 'Live Microsoft Graph organization verifiedDomains' } else { 'Unavailable' }
        SourceTimestamp = $sourceTimestamp
        Message = if (-not $available) { 'Microsoft Entra verified-domain evidence is unavailable.' } elseif ([string]::IsNullOrWhiteSpace($upnDomain)) { 'The Microsoft Entra user UPN is empty or malformed.' } elseif ($verified) { "UPN domain '$upnDomain' is verified in Microsoft Entra." } else { "UPN domain '$upnDomain' is not present in the tenant verifiedDomains collection." }
    }
}

function Get-SemrProxyConflictEvidence {
    param(
        [Parameter(Mandatory)][string]$EmailAddress,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Addresses
    )

    if (-not (Test-SemrCheckEnabled -CheckId 'PROXY-SMTP-GLOBAL-UNIQUE') -and -not (Test-SemrCheckEnabled -CheckId 'TARGET-ADDRESS-GLOBAL-UNIQUE')) {
        return [pscustomobject]@{ Available = $false; Source = 'Disabled'; SourceTimestamp = $null; Conflicts = @(); PlannedWarnings = @() }
    }
    $key = $EmailAddress.Trim().ToLowerInvariant()
    if (-not $script:ConnectionState.OnPremisesExchange -or -not $script:Exchange2016EvidenceByEmail.ContainsKey($key)) {
        return [pscustomobject]@{ Available = $false; Source = 'Local Exchange 2016 Management Shell unavailable'; SourceTimestamp = $null; Conflicts = @(); PlannedWarnings = @() }
    }

    $entry = $script:Exchange2016EvidenceByEmail[$key]
    $wanted = @($Addresses | ForEach-Object { ([string]$_ -replace '^(?i:smtp:)', '').Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
    $queries = @($entry.AddressConflicts | Where-Object { $wanted -contains ([string]$_.Address).ToLowerInvariant() })
    $available = $wanted.Count -eq 0 -or ($queries.Count -eq $wanted.Count -and @($queries | Where-Object { -not $_.Available }).Count -eq 0)
    return [pscustomobject]@{
        Available = $available
        Source = [string]$entry.Source
        SourceTimestamp = $entry.SourceTimestamp
        Conflicts = @($queries | ForEach-Object { @($_.Conflicts) } | Sort-Object -Unique)
        PlannedWarnings = @()
    }
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
    if ($conflict.Source -and $conflict.SourceTimestamp) { $script:CurrentSourceTimestamps[[string]$conflict.Source] = $conflict.SourceTimestamp }
    $conflicts = @($conflict.Conflicts)
    $plannedWarnings = @($conflict.PlannedWarnings)
    $proxyResult = if (-not $conflict.Available) { 'UNKNOWN' } elseif ($conflicts.Count) { 'FAIL' } elseif ($plannedWarnings.Count) { 'WARN' } else { 'PASS' }
    Add-SemrFinding $Findings ($Base + @{
        CheckId = 'PROXY-SMTP-GLOBAL-UNIQUE'; Category = 'HybridIdentity'
        Severity = if ($proxyResult -eq 'FAIL') { 'Critical' } elseif ($proxyResult -in @('UNKNOWN','WARN')) { 'Warning' } else { 'Information' }
        Result = $proxyResult
        IsBlocking = $proxyResult -eq 'FAIL'
        ObservedValue = if (-not $conflict.Available) { 'Global recipient index unavailable' } elseif ($conflicts.Count) { $conflicts -join ' | ' } elseif ($plannedWarnings.Count) { 'Planned address warning: ' + ($plannedWarnings -join ' | ') } else { 'No conflicting owner' }
        ExpectedValue = 'No current SMTP owner conflict; planned remote routing addresses remain unique'
        EvidenceSource = $conflict.Source; SourceTimestamp = $conflict.SourceTimestamp
        Message = if (-not $conflict.Available) { 'Global SMTP ownership could not be evaluated.' } elseif ($conflicts.Count) { 'At least one SMTP proxy is currently owned by another recipient.' } elseif ($plannedWarnings.Count) { 'A future expected routing address is duplicated, but no current SMTP owner conflict was reported.' } else { 'No current or planned SMTP conflict was found.' }
        RecommendedAction = if (-not $conflict.Available) { 'Restore live Exchange on-premises recipient lookup and rerun.' } elseif ($conflicts.Count) { 'Resolve every current SMTP ownership conflict.' } elseif ($plannedWarnings.Count) { 'Review the planned remote routing address and coordinate a unique value before onboarding the affected recipients.' } else { '' }
    })

    $target = ([string](Get-SemrPropertyValue $mailbox[0] @('ExternalEmailAddress','WindowsEmailAddress') '') -replace '^(?i:smtp:)','').Trim().ToLowerInvariant()
    $targetConflictEvidence = if ($target) { Get-SemrProxyConflictEvidence -EmailAddress $email -Addresses @($target) } else { $null }
    $targetEvidence = -not $target -or ($targetConflictEvidence -and $targetConflictEvidence.Available)
    $targetConflicts = if ($targetConflictEvidence) { @($targetConflictEvidence.Conflicts) } else { @() }
    $targetConflict = $targetConflicts.Count -gt 0
    $targetSource = if ($targetConflictEvidence) { [string]$targetConflictEvidence.Source } else { 'Live Exchange on-premises recipient directory' }
    $targetTimestamp = if ($targetConflictEvidence) { $targetConflictEvidence.SourceTimestamp } else { Get-Date }
    Add-SemrFinding $Findings ($Base + @{CheckId='TARGET-ADDRESS-GLOBAL-UNIQUE';Category='HybridIdentity';Severity=if(-not $targetEvidence){'Warning'}elseif($targetConflict){'Critical'}else{'Information'};Result=if(-not $targetEvidence){'UNKNOWN'}elseif($targetConflict){'FAIL'}else{'PASS'};IsBlocking=[bool]$targetConflict;ObservedValue=if($target){if($targetConflict){$targetConflicts -join ' | '}else{"$target; no conflicting owner"}}else{'No targetAddress on source UserMailbox'};ExpectedValue='Empty before onboarding or one unique routing address';EvidenceSource=$targetSource;SourceTimestamp=$targetTimestamp;Message=if(-not $targetEvidence){'Global targetAddress uniqueness could not be evaluated live.'}elseif($targetConflict){'The target routing address is owned by another recipient.'}else{'No duplicated target routing address was detected.'};RecommendedAction=if($targetConflict){'Resolve the duplicate remote routing address.'}elseif(-not $targetEvidence){'Restore live Exchange on-premises recipient lookup and rerun.'}else{''}})

    $legacyDn = [string](Get-SemrPropertyValue $mailbox[0] @('LegacyExchangeDN') '')
    $x500 = $legacyDn -and @($raw | Where-Object { (([string]$_) -replace '^(?i:x500:)','') -ieq $legacyDn }).Count -gt 0
    Add-SemrFinding $Findings ($Base + @{CheckId='X500-LEGACYEXCHANGEDN';Category='HybridIdentity';Severity=if(-not $legacyDn){'Warning'}elseif($x500){'Information'}else{'Critical'};Result=if(-not $legacyDn){'UNKNOWN'}elseif($x500){'PASS'}else{'FAIL'};IsBlocking=[bool]($legacyDn -and -not $x500);ObservedValue=if($legacyDn){"LegacyExchangeDN=$legacyDn; X500Present=$x500"}else{'LegacyExchangeDN unavailable'};ExpectedValue='LegacyExchangeDN preserved as X500';EvidenceSource=$source;Message=if(-not $legacyDn){'LegacyExchangeDN is unavailable.'}elseif($x500){'LegacyExchangeDN is preserved in proxyAddresses.'}else{'LegacyExchangeDN is not preserved as an X500 proxy.'};RecommendedAction=if(-not $legacyDn){'Restore the local Exchange 2016 worker property collection and rerun.'}elseif(-not $x500){'Add the exact legacyExchangeDN as an X500 proxy.'}else{''}})

    $domains = @($normalized | ForEach-Object { ($_ -split '@',2)[1] } | Where-Object { $_ } | Sort-Object -Unique)
    $unaccepted = @(if($AcceptedDomains.Available){$domains | Where-Object { $AcceptedDomains.Domains -notcontains $_ }})
    Add-SemrFinding $Findings ($Base + @{CheckId='SMTP-ACCEPTED-DOMAIN';Category='HybridIdentity';Severity=if(-not $AcceptedDomains.Available){'Warning'}elseif($unaccepted.Count){'Critical'}else{'Information'};Result=if(-not $AcceptedDomains.Available){'UNKNOWN'}elseif($unaccepted.Count){'FAIL'}else{'PASS'};IsBlocking=$AcceptedDomains.Available -and $unaccepted.Count -gt 0;ObservedValue=if($AcceptedDomains.Available){if($unaccepted.Count){'Not accepted: '+($unaccepted -join ';')}else{$domains -join ';'}}else{$AcceptedDomains.Message};ExpectedValue='Every SMTP domain accepted in Exchange Online';EvidenceSource=$AcceptedDomains.Source;Message=if(-not $AcceptedDomains.Available){'Accepted domains could not be evaluated.'}elseif($unaccepted.Count){'One or more SMTP domains are not accepted.'}else{'Every SMTP proxy domain is accepted.'};RecommendedAction=if($unaccepted.Count){'Add/verify accepted domains or remove invalid proxies.'}elseif(-not $AcceptedDomains.Available){'Restore the live Exchange Online accepted-domain query and rerun.'}else{''}})

    $type = [string](Get-SemrPropertyValue $mailbox[0] @('RecipientTypeDetails') '')
    $supported = @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','LinkedMailbox')
    $typePass = $supported -contains $type
    Add-SemrFinding $Findings ($Base + @{CheckId='RECIPIENT-TYPE-SUPPORTED';Category='ExchangeOnPrem';Severity=if($typePass){'Information'}else{'Critical'};Result=if($typePass){'PASS'}else{'FAIL'};IsBlocking=-not $typePass;ObservedValue=$type;ExpectedValue=$supported -join ', ';EvidenceSource=$source;Message=if($typePass){'The source recipient type is supported.'}else{'The source recipient type is unsupported or ambiguous.'};RecommendedAction=if($typePass){''}else{'Exclude system mailboxes or validate another migration method.'}})

    $cloudRecipient = @($Exo.Recipients | Select-Object -First 1)
    foreach($definition in @(@('EXCHANGE-GUID-CONSISTENCY','ExchangeGuid','Exchange GUID'),@('ARCHIVE-GUID-CONSISTENCY','ArchiveGuid','Archive GUID'))){
        $checkId=$definition[0];$property=$definition[1];$label=$definition[2]
        $sourceGuid=[string](Get-SemrPropertyValue $mailbox[0] @($property) '')
        $cloudGuid=if($cloudRecipient.Count){[string](Get-SemrPropertyValue $cloudRecipient[0] @($property) '')}else{''}
        $archiveStateText = [string](Get-SemrPropertyValue $mailbox[0] @('ArchiveStatus','ArchiveState') '')
        $archiveNotRequired=$checkId -eq 'ARCHIVE-GUID-CONSISTENCY' -and ([string]::IsNullOrWhiteSpace($archiveStateText) -or $archiveStateText -match '^(?i:None|NotPresent|Disabled)$') -and -not $sourceGuid
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
    $hasArchive=$archiveStatus -match '^(?i:Active|HostedPending|Local)$' -or ($archiveGuid -and $archiveGuid -notmatch '^0{8}-')
    $archiveStatistics=@($OnPrem.ArchiveStatistics | Select-Object -First 1)
    $archiveSize=if($archiveStatistics.Count){ConvertFrom-SemrByteSize (Get-SemrPropertyValue $archiveStatistics[0] @('TotalItemSize') '')}else{0}

    $targetSku=if($Row.TargetSku){[string]$Row.TargetSku}else{[string]$Config['DefaultTargetSku']}
    $unsupported=$hasArchive -and $targetSku -match 'F1|F3|DESKLESS'
    $tooLarge=$hasArchive -and $archiveSize -ge 95
    $archivePass=-not $unsupported -and -not $tooLarge
    Add-SemrFinding $Findings ($Base+@{CheckId='ARCHIVE-READINESS';Category='Mailbox';Severity=if(-not $archivePass){'Critical'}elseif($hasArchive){'Warning'}else{'Information'};Result=if(-not $archivePass){'FAIL'}elseif($hasArchive){'WARN'}else{'PASS'};IsBlocking=-not $archivePass;ObservedValue="Archive=$hasArchive; Size=$archiveSize GB; TargetSku=$targetSku";ExpectedValue='Archive-capable SKU and archive below 95 GB safety threshold';EvidenceSource="$source + live archive statistics";Message=if($unsupported){'The selected target SKU is not archive-capable.'}elseif($tooLarge){'The archive is at or above the safe target threshold.'}elseif($hasArchive){'An archive is present and must be included in the migration plan.'}else{'No active source archive was detected.'};RecommendedAction=if($unsupported){'Select an archive-capable Exchange Online license.'}elseif($tooLarge){'Reduce archive content or validate auto-expanding archive requirements.'}elseif($hasArchive){'Confirm that MailboxType includes the archive.'}else{''}})

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
    $largeCollectionStatus=[string](Get-SemrPropertyValue $mailbox[0] @('LargeItemCollectionStatus') '')
    $largeKnown=$largeRaw -match '^\d+$';$largeCount=if($largeKnown){[int]$largeRaw}else{0}
    $largeUnknownAction = 'Run a targeted live large-item scan or review the migration report before approving LargeItemLimit.'
    Add-SemrFinding $Findings ($Base+@{CheckId='MAILBOX-LARGE-ITEMS';Category='Mailbox';Severity=if(-not $largeKnown -or $largeCount){'Warning'}else{'Information'};Result=if(-not $largeKnown){'UNKNOWN'}elseif($largeCount){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if($largeKnown){"$largeCount item(s) over inventory threshold; CollectionStatus=$largeCollectionStatus"}else{"Large-item evidence unavailable; CollectionStatus=$largeCollectionStatus"};ExpectedValue='0, or explicit migration exception policy';EvidenceSource=$source;Message=if(-not $largeKnown -and $largeCollectionStatus -eq 'NotCollected'){'Large-item statistics are not exposed by the available live mailbox cmdlets.'}elseif(-not $largeKnown){'Large-item exposure could not be evaluated.'}elseif($largeCount){'Large items may require an explicit migration policy.'}else{'No oversized item was reported.'};RecommendedAction=if($largeCount){'Review sizes and explicitly approve BadItemLimit/LargeItemLimit.'}elseif(-not $largeKnown){$largeUnknownAction}else{''}})

    $quotaDefaults=Get-SemrPropertyValue $mailbox[0] @('UseDatabaseQuotaDefaults') $null
    $quotaKnown=$null -ne $quotaDefaults;$custom=$quotaKnown -and -not (ConvertTo-SemrBoolean $quotaDefaults)
    Add-SemrFinding $Findings ($Base+@{CheckId='CUSTOM-SOURCE-QUOTA';Category='Mailbox';Severity=if(-not $quotaKnown -or $custom){'Warning'}else{'Information'};Result=if(-not $quotaKnown){'UNKNOWN'}elseif($custom){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if($quotaKnown){"UseDatabaseQuotaDefaults=$quotaDefaults; ProhibitSendReceiveQuota=$([string](Get-SemrPropertyValue $mailbox[0] @('ProhibitSendReceiveQuota') ''))"}else{'Source quota evidence unavailable'};ExpectedValue='Source quota documented';EvidenceSource=$source;Message=if(-not $quotaKnown){'Source quota customization could not be evaluated.'}elseif($custom){'The mailbox uses a custom source quota.'}else{'The mailbox uses database quota defaults.'};RecommendedAction=if($custom){'Confirm the target SKU quota independently of the custom source quota.'}elseif(-not $quotaKnown){'Restore the required property in the live Exchange on-premises query.'}else{''}})
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
        $resolvedDelegates=@($delegates | ForEach-Object {$delegateKey=$_.Trim().ToLowerInvariant();if($script:BatchDelegateIdentityToEmail.ContainsKey($delegateKey)){$script:BatchDelegateIdentityToEmail[$delegateKey]}else{$_}} | Sort-Object -Unique)
        $external=@($resolvedDelegates | Where-Object {$batchEmails -notcontains $_.ToLowerInvariant()})
        $knownText=if($resolvedDelegates.Count){$resolvedDelegates -join ';'}else{'None observed'}
        $dependencyObserved=if(-not $permissionsAvailable){"Partial baseline: $($resolvedDelegates.Count) known delegate(s): $knownText"}elseif($external.Count){$external -join ';'}else{'All resolvable delegates are in the batch or none exist'}
        $dependencyMessage=if(-not $permissionsAvailable -and $external.Count){'Permission evidence is incomplete, but known delegates outside this migration batch were identified.'}elseif(-not $permissionsAvailable){'Delegate migration dependencies could not be evaluated completely; known grants are still listed.'}elseif($external.Count){'One or more delegates are outside this migration batch.'}else{'No cross-wave delegate dependency was identified.'}
        $dependencyAction=@()
        if(-not $permissionsAvailable){$dependencyAction+='Restore Full Access and Send As permission queries on live Exchange on-premises, then rerun.'}
        if($external.Count){$dependencyAction+='Coordinate these delegates or document temporary cross-premises limitations.'}
        Add-SemrFinding $Findings ($Base+@{CheckId='DELEGATE-MIGRATION-DEPENDENCY';Category='Permissions';Severity=if(-not $permissionsAvailable -or $external.Count){'Warning'}else{'Information'};Result=if(-not $permissionsAvailable){'UNKNOWN'}elseif($external.Count){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=$dependencyObserved;ExpectedValue='Delegates coordinated or documented';EvidenceSource=$source;Message=$dependencyMessage;RecommendedAction=$dependencyAction -join ' | '})

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
        Add-SemrFinding $Findings ($Base+@{CheckId='DELIVERY-RESTRICTIONS';Category='MailFlow';Severity=if(-not $deliveryAvailable -or $restricted){'Warning'}else{'Information'};Result=if(-not $deliveryAvailable){'UNKNOWN'}elseif($restricted){'WARN'}else{'PASS'};IsBlocking=$false;ObservedValue=if($deliveryAvailable){"Moderation=$moderated; RestrictionTypes=$restrictionCount"}else{'Delivery restriction evidence unavailable'};ExpectedValue='Delivery restrictions documented';EvidenceSource=$source;Message=if(-not $deliveryAvailable){'Delivery restrictions could not be evaluated from live Exchange on-premises.'}elseif($restricted){'Moderation or delivery restrictions are configured.'}else{'No moderation or explicit delivery restriction was detected.'};RecommendedAction=if(-not $deliveryAvailable){'Restore the live Exchange on-premises delivery-restriction query.'}elseif($restricted){'Capture and validate restriction owners and membership after migration.'}else{''}})

        $dbAvailable = [bool](Get-SemrPropertyValue $OnPrem @('DatabaseHealthAvailable') $false)
        $db = @($OnPrem.DatabaseHealth | Select-Object -First 1)
        $databaseName = [string](Get-SemrPropertyValue $mailbox[0] @('Database') '')
        $mounted = $dbAvailable -and $db.Count -eq 1 -and (ConvertTo-SemrBoolean (Get-SemrPropertyValue $db[0] @('Mounted') $false))
        $dbSource = [string](Get-SemrPropertyValue $OnPrem @('DatabaseHealthSource') $source)
        $dbSourceTimestamp = Get-SemrPropertyValue $OnPrem @('DatabaseHealthSourceTimestamp') $null
        Add-SemrFinding $Findings ($Base + @{
            CheckId = 'EXCHANGE-DATABASE-HEALTH'; Category = 'ExchangeInfrastructure'
            Severity = if (-not $dbAvailable) { 'Warning' } elseif ($mounted) { 'Information' } else { 'Critical' }
            Result = if (-not $dbAvailable) { 'UNKNOWN' } elseif ($mounted) { 'PASS' } else { 'FAIL' }
            IsBlocking = $dbAvailable -and -not $mounted
            ObservedValue = if ($dbAvailable) { "Database=$databaseName; Mounted=$mounted; Server=$([string](Get-SemrPropertyValue $db[0] @('Server') '')); ReplicationType=$([string](Get-SemrPropertyValue $db[0] @('ReplicationType') ''))" } else { "Database=$databaseName; health unavailable" }
            ExpectedValue = 'Source mailbox database mounted'
            EvidenceSource = $dbSource; SourceTimestamp = $dbSourceTimestamp
            Message = if (-not $dbAvailable) { 'Database health could not be evaluated.' } elseif ($mounted) { 'The source mailbox database is mounted.' } else { 'The source mailbox database is not mounted.' }
            RecommendedAction = if (-not $dbAvailable) { 'Validate database health immediately before the batch.' } elseif (-not $mounted) { 'Restore database availability before migration.' } else { '' }
        })
    }

    $moveAvailable=[bool](Get-SemrPropertyValue $Exo @('MoveDataAvailable') $false)
    $moveState=Get-SemrMoveState -Exo $Exo
    $badStatuses=@($moveState.BadStatuses)
    Add-SemrFinding $Findings ($Base+@{CheckId='MOVE-HISTORY';Category='Migration';Severity=if(-not $moveAvailable){'Warning'}elseif($badStatuses.Count){'Critical'}else{'Information'};Result=if(-not $moveAvailable){'UNKNOWN'}elseif($badStatuses.Count){'FAIL'}else{'PASS'};IsBlocking=$moveAvailable -and $badStatuses.Count -gt 0;ObservedValue=if($moveAvailable){if($badStatuses.Count){$badStatuses -join ';'}else{'No failed or suspended current migration object'}}else{'Current migration failure evidence unavailable'};ExpectedValue='No unresolved failed or suspended current migration object';EvidenceSource=$exoSource;Message=if(-not $moveAvailable){'Current migration failure state could not be evaluated.'}elseif($badStatuses.Count){'An unresolved failed or suspended migration object was detected.'}else{'No unresolved failed current migration object was returned.'};RecommendedAction=if($badStatuses.Count){'Review the migration report and resolve or clean up the failed move deliberately.'}elseif(-not $moveAvailable){'Restore the live Exchange Online migration-user query.'}else{''}})

    $user=@($Graph.Users | Select-Object -First 1)
    if($user.Count -eq 1){
        $errors=@(Get-SemrPropertyValue $user[0] @('OnPremisesProvisioningErrors') @())
        $anchor=[string](Get-SemrPropertyValue $user[0] @('OnPremisesImmutableId') '')
        $lastSyncText=[string](Get-SemrPropertyValue $user[0] @('OnPremisesLastSyncDateTime') '')
        $lastSync=$null;if($lastSyncText){try{$lastSync=[datetime]$lastSyncText}catch{}}
        $timestampOld=$lastSync -and $lastSync -lt (Get-Date).AddHours(-24)
        $syncEnabled=ConvertTo-SemrBoolean (Get-SemrPropertyValue $user[0] @('OnPremisesSyncEnabled') $false)
        $syncError=$errors.Count -gt 0 -or ($syncEnabled -and [string]::IsNullOrWhiteSpace($anchor))
        Add-SemrFinding $Findings ($Base+@{CheckId='ENTRA-OBJECT-SYNC-ERROR';Category='MicrosoftGraph';Severity=if($syncError){'Critical'}else{'Information'};Result=if($syncError){'FAIL'}else{'PASS'};IsBlocking=$syncError;ObservedValue="ProvisioningErrors=$($errors.Count); SyncEnabled=$syncEnabled; ImmutableIdPresent=$(-not [string]::IsNullOrWhiteSpace($anchor)); ObjectLastSync=$lastSyncText";ExpectedValue='No provisioning errors and an identity anchor for synchronized users';EvidenceSource=$graphSource;Message=if($errors.Count){'Microsoft Entra reports provisioning errors.'}elseif($syncEnabled -and -not $anchor){'The synchronized identity anchor is missing.'}elseif($timestampOld){'The per-object synchronization timestamp is older than 24 hours; it is informational and is not used as a proxy for tenant-wide synchronization freshness.'}else{'No object-level synchronization issue was detected.'};RecommendedAction=if($syncError){'Resolve Entra Connect export errors and identity anchoring.'}elseif($timestampOld){'Use the ENTRA-CONNECT-SCHEDULER result to assess current synchronization service health.'}else{''}})
    }
}

function Get-SemrHybridAdvancedEvidence {
    param([System.Collections.IDictionary]$Config)
    $capacityWarningThreshold = if($Config -and $Config.Contains('Hybrid') -and $Config['Hybrid'].Contains('ActiveMigrationWarningThreshold')){[int]$Config['Hybrid']['ActiveMigrationWarningThreshold']}else{100}
    $result=[ordered]@{
        MrsProxyAvailable=$false;MrsProxyEnabled=$false;MrsProxyMessage='MRSProxy evidence unavailable.';MrsProxySource='Unavailable';MrsProxySourceTimestamp=$null
        CertificateAvailable=$false;CertificateHealthy=$false;CertificateDaysRemaining=$null;CertificateMessage='Hybrid certificate evidence unavailable.';CertificateSource='Unavailable';CertificateSourceTimestamp=$null
        CapacityAvailable=$false;ActiveMigrationCount=$null;CapacityHealthy=$false;CapacityMessage='Migration load evidence unavailable.';CapacitySource='Unavailable';CapacitySourceTimestamp=$null
        FailureBacklogAvailable=$false;FailureBacklogCount=$null;FailureBacklogHealthy=$false;FailureBacklogMessage='Migration failure backlog evidence unavailable.';FailureBacklogSource='Unavailable';FailureBacklogSourceTimestamp=$null
        OAuthAvailable=$false;OAuthHealthy=$false;OAuthMessage='Autodiscover/OAuth evidence unavailable.';OAuthSource='Unavailable';OAuthSourceTimestamp=$null
    }

    if ($script:ConnectionState.OnPremisesExchange -and $script:Exchange2016HybridEvidence) {
        $workerHybrid = $script:Exchange2016HybridEvidence
        if (Test-SemrCheckEnabled 'HYBRID-MRSPROXY') {
            foreach ($name in @('MrsProxyAvailable','MrsProxyEnabled','MrsProxyMessage','MrsProxySource','MrsProxySourceTimestamp')) {
                $result[$name] = Get-SemrPropertyValue -InputObject $workerHybrid -Names @($name) -Default $result[$name]
            }
        }
        if (Test-SemrCheckEnabled 'HYBRID-AUTODISCOVER-OAUTH') {
            foreach ($name in @('OAuthAvailable','OAuthHealthy','OAuthMessage','OAuthSource','OAuthSourceTimestamp')) {
                $result[$name] = Get-SemrPropertyValue -InputObject $workerHybrid -Names @($name) -Default $result[$name]
            }
        }
    }
    if(((Test-SemrCheckEnabled 'HYBRID-ENDPOINT-CAPACITY') -or (Test-SemrCheckEnabled 'HYBRID-MIGRATION-BACKLOG') -or (Test-SemrCheckEnabled 'EXO-EXISTING-MOVE') -or (Test-SemrCheckEnabled 'MOVE-HISTORY') -or (Test-SemrCheckEnabled 'KNOWN-ENHANCED-RESTORE')) -and $script:ConnectionState.ExchangeOnline -and (Test-SemrCommand 'Get-MigrationUser')){
        try{
            $allMoves = @(Get-MigrationUser -ResultSize Unlimited -ErrorAction Stop)
            $script:AllMigrationUsers = $allMoves
            $script:MigrationUsersLoaded = $true
            $script:MigrationUsersByEmail = @{}
            foreach($move in $allMoves){
                foreach($identityValue in @(
                    (Get-SemrPropertyValue -InputObject $move -Names @('EmailAddress') -Default ''),
                    (Get-SemrPropertyValue -InputObject $move -Names @('RecipientIdentifier') -Default ''),
                    (Get-SemrPropertyValue -InputObject $move -Names @('Identity') -Default '')
                )){
                    $identityText = ([string]$identityValue).Trim()
                    if($identityText -match '^(?i:smtp):(.+)$'){ $identityText = $Matches[1] }
                    if($identityText -notmatch '@'){ continue }
                    $key = $identityText.ToLowerInvariant()
                    if(-not $script:MigrationUsersByEmail.ContainsKey($key)){ $script:MigrationUsersByEmail[$key] = [System.Collections.Generic.List[object]]::new() }
                    [void]$script:MigrationUsersByEmail[$key].Add($move)
                }
            }
            $activeStatuses = @('Queued','Syncing','IncrementalSyncing','Completing','Validating','Provisioning','Starting','Removing','Stopping')
            $failureStatuses = @('Failed','Stopped','Corrupted','Error')
            $activeMoves = @($allMoves | Where-Object { [string]$_.Status -in $activeStatuses })
            $failedMoves = @($allMoves | Where-Object { [string]$_.Status -in $failureStatuses -or [string]$_.Status -match 'Fail|Error|Corrupt' })
            $statusBreakdown = @($allMoves | Group-Object Status | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join '; '
            $collectedAt = Get-Date
            $result.CapacityAvailable=$true;$result.ActiveMigrationCount=$activeMoves.Count;$result.CapacityHealthy=$activeMoves.Count -lt $capacityWarningThreshold
            $result.CapacityMessage="$($activeMoves.Count) active migration user(s); advisory threshold is $capacityWarningThreshold. Statuses: $statusBreakdown";$result.CapacitySource='Live Exchange Online migration users';$result.CapacitySourceTimestamp=$collectedAt
            $result.FailureBacklogAvailable=$true;$result.FailureBacklogCount=$failedMoves.Count;$result.FailureBacklogHealthy=$failedMoves.Count -eq 0
            $result.FailureBacklogMessage="$($failedMoves.Count) failed/stopped/corrupted migration user(s). Statuses: $statusBreakdown";$result.FailureBacklogSource='Live Exchange Online migration users';$result.FailureBacklogSourceTimestamp=$collectedAt
        }catch{
            $result.CapacityMessage=$_.Exception.Message
            $result.FailureBacklogMessage=$_.Exception.Message
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

function Get-SemrTlsCertificateEvidence {
    param(
        [Parameter(Mandatory)][string]$RemoteServer,
        [int]$TimeoutMilliseconds = 10000
    )

    $result = [ordered]@{
        Available = $false
        Healthy = $false
        DaysRemaining = $null
        Message = ''
        Source = "Live TLS endpoint probe: $RemoteServer"
        SourceTimestamp = Get-Date
    }
    $hostName = $RemoteServer.Trim()
    $port = 443
    try {
        $uri = $null
        if ([uri]::TryCreate($hostName, [UriKind]::Absolute, [ref]$uri)) {
            $hostName = $uri.DnsSafeHost
            if (-not $uri.IsDefaultPort) { $port = $uri.Port }
        }
        elseif ($hostName -match '^([^:]+):(\d+)$') {
            $hostName = $Matches[1]
            $port = [int]$Matches[2]
        }
        if ([string]::IsNullOrWhiteSpace($hostName)) { throw 'The migration endpoint RemoteServer is empty.' }

        $tcpClient = [System.Net.Sockets.TcpClient]::new()
        try {
            $connect = $tcpClient.BeginConnect($hostName, $port, $null, $null)
            if (-not $connect.AsyncWaitHandle.WaitOne($TimeoutMilliseconds)) { throw "TLS connection timed out after $TimeoutMilliseconds ms." }
            $tcpClient.EndConnect($connect)
            $state = [pscustomobject]@{ Certificate = $null; PolicyErrors = [System.Net.Security.SslPolicyErrors]::None }
            $callback = {
                param($sender,$certificate,$chain,$sslPolicyErrors)
                $state.Certificate = if ($certificate) { [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certificate) } else { $null }
                $state.PolicyErrors = $sslPolicyErrors
                return $true
            }.GetNewClosure()
            $sslStream = [System.Net.Security.SslStream]::new($tcpClient.GetStream(),$false,$callback)
            try {
                $sslStream.ReadTimeout = $TimeoutMilliseconds
                $sslStream.WriteTimeout = $TimeoutMilliseconds
                $sslStream.AuthenticateAsClient($hostName)
                if (-not $state.Certificate) { throw 'The remote endpoint did not present a TLS certificate.' }
                $daysRemaining = [math]::Floor(($state.Certificate.NotAfter.ToUniversalTime() - (Get-Date).ToUniversalTime()).TotalDays)
                $policyHealthy = $state.PolicyErrors -eq [System.Net.Security.SslPolicyErrors]::None
                $result.Available = $true
                $result.DaysRemaining = $daysRemaining
                $result.Healthy = $policyHealthy -and $daysRemaining -ge 60
                $result.Message = "Endpoint $($hostName):$port presented certificate '$($state.Certificate.Subject)' with $daysRemaining day(s) remaining; TLS policy errors=$($state.PolicyErrors)."
            }
            finally { $sslStream.Dispose() }
        }
        finally { $tcpClient.Dispose() }
    }
    catch {
        $result.Message = "TLS certificate probe failed for ${RemoteServer}: $($_.Exception.Message)"
    }
    return [pscustomobject]$result
}


function Test-SemrHybridReadiness {
    [CmdletBinding()]
    param([System.Collections.IDictionary]$Config)

    $mode = if ($Config.Contains('Mode')) { [string]$Config['Mode'] } else { 'Live' }
    $hybridConfig = if ($Config.Contains('Hybrid')) { $Config['Hybrid'] } else { $Config }
    $advanced = Get-SemrHybridAdvancedEvidence -Config $Config
    $result = [ordered]@{
        Available = $false
        EndpointFound = $false
        EndpointName = ''
        ConnectivitySuccess = $false
        ConnectivityTestAvailable = $false
        SourceTimestamp = if ($mode -eq 'Live') { Get-Date } else { $null }
        Source = 'Live Exchange Online'
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
        FailureBacklogAvailable = $advanced.FailureBacklogAvailable
        FailureBacklogCount = $advanced.FailureBacklogCount
        FailureBacklogHealthy = $advanced.FailureBacklogHealthy
        FailureBacklogMessage = $advanced.FailureBacklogMessage
        FailureBacklogSource = $advanced.FailureBacklogSource
        FailureBacklogSourceTimestamp = $advanced.FailureBacklogSourceTimestamp
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
    if (Test-SemrCheckEnabled -CheckId 'HYBRID-CERTIFICATE-EXPIRY') {
        $endpointRemoteServer = [string](Get-SemrPropertyValue -InputObject $endpoint -Names @('RemoteServer','RemoteServerName') -Default '')
        if ($endpointRemoteServer) {
            $tlsEvidence = Get-SemrTlsCertificateEvidence -RemoteServer $endpointRemoteServer
            $result.CertificateAvailable = $tlsEvidence.Available
            $result.CertificateHealthy = $tlsEvidence.Healthy
            $result.CertificateDaysRemaining = $tlsEvidence.DaysRemaining
            $result.CertificateMessage = $tlsEvidence.Message
            $result.CertificateSource = $tlsEvidence.Source
            $result.CertificateSourceTimestamp = $tlsEvidence.SourceTimestamp
        }
        else {
            $result.CertificateAvailable = $false
            $result.CertificateHealthy = $false
            $result.CertificateDaysRemaining = $null
            $result.CertificateMessage = 'The selected migration endpoint does not expose RemoteServer; its presented TLS certificate could not be probed.'
            $result.CertificateSource = 'Selected Exchange Online migration endpoint'
            $result.CertificateSourceTimestamp = Get-Date
        }
    }

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
            if ($result.ConnectivitySuccess) {
                $result.MrsProxyAvailable = $true
                $result.MrsProxyEnabled = $true
                $result.MrsProxyMessage = 'The successful ExchangeRemoteMove endpoint test functionally validates published MRSProxy connectivity.'
                $result.MrsProxySource = 'Live Exchange Online migration endpoint connectivity test'
                $result.MrsProxySourceTimestamp = Get-Date
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

    $maximumLastSyncAgeMinutes = if ($Config -and $Config.Contains('EntraConnectHealth')) { [double]$Config['EntraConnectHealth']['MaximumLastSyncAgeMinutes'] } else { 120.0 }
    $graphSource = 'Live Microsoft Graph organization'
    if (-not $script:ConnectionState.MicrosoftGraph -or -not $script:GraphOrganization) {
        $graphError = if ($script:GraphOrganizationError) { $script:GraphOrganizationError } elseif (-not $script:ConnectionState.MicrosoftGraph) { 'Microsoft Graph evidence has not been collected.' } else { 'Microsoft Graph returned no organization object.' }
        return [pscustomobject][ordered]@{
            Available=$false; Authoritative=$true
            Server='Microsoft Graph'; Source=$graphSource; SourceTimestamp=$null
            SyncCycleEnabled=$null; SchedulerSuspended=$null; StagingModeEnabled=$null; NextSyncCycleStartTimeInUTC=$null; ConnectorRunStatus=''
            LastSyncDateTimeUtc=$null; LastSyncAgeMinutes=$null; LastSyncKnown=$false; LastSyncFresh=$false
            Message="Authoritative tenant synchronization health is unavailable from Microsoft Graph: $graphError"
        }
    }

    $organization = $script:GraphOrganization
    $syncEnabledValue = Get-SemrPropertyValue -InputObject $organization -Names @('OnPremisesSyncEnabled') -Default $null
    $syncEnabled = if ($null -eq $syncEnabledValue) { $null } else { [bool]$syncEnabledValue }
    $lastSyncValue = Get-SemrPropertyValue -InputObject $organization -Names @('OnPremisesLastSyncDateTime') -Default $null
    $lastSyncDateTimeUtc = $null
    if ($lastSyncValue) {
        try {
            $parsedLastSync = [datetime]$lastSyncValue
            $lastSyncDateTimeUtc = if ($parsedLastSync.Kind -eq [DateTimeKind]::Unspecified) { [datetime]::SpecifyKind($parsedLastSync,[DateTimeKind]::Utc) } else { $parsedLastSync.ToUniversalTime() }
        }
        catch { $lastSyncDateTimeUtc = $null }
    }
    $lastSyncAgeMinutes = if ($lastSyncDateTimeUtc) { [math]::Round(((Get-Date).ToUniversalTime() - $lastSyncDateTimeUtc).TotalMinutes,1) } else { $null }
    $lastSyncFresh = $null -ne $lastSyncAgeMinutes -and ($maximumLastSyncAgeMinutes -le 0 -or $lastSyncAgeMinutes -le $maximumLastSyncAgeMinutes)
    $sourceTimestamp = Get-SemrPropertyValue -InputObject $organization -Names @('CollectedAt') -Default (Get-Date)
    $script:ConnectionState.EntraConnect = $true
    return [pscustomobject][ordered]@{
        Available=$true; Authoritative=$true
        Server='Microsoft Graph'; Source=$graphSource; SourceTimestamp=$sourceTimestamp
        SyncCycleEnabled=$syncEnabled; SchedulerSuspended=$null; StagingModeEnabled=$null; NextSyncCycleStartTimeInUTC=$null; ConnectorRunStatus=''
        LastSyncDateTimeUtc=$lastSyncDateTimeUtc; LastSyncAgeMinutes=$lastSyncAgeMinutes; LastSyncKnown=$null -ne $lastSyncDateTimeUtc; LastSyncFresh=$lastSyncFresh
        Message="Tenant organization synchronization state collected from Microsoft Graph; OnPremisesSyncEnabled=$syncEnabled; last sync UTC=$lastSyncDateTimeUtc; age=$lastSyncAgeMinutes minute(s); maximum=$maximumLastSyncAgeMinutes."
    }
}
function Invoke-SemrAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Batch,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Config,
        [scriptblock]$ProgressCallback,
        [scriptblock]$CancellationCheck,
        [string]$DiagnosticsRoot = ''
    )

    Set-SemrAssessmentCheckOptions -Config $Config
    $startedAt = Get-Date
    $runId = "SEMR-{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    $exchangeDiagnosticsDirectory = if ([string]::IsNullOrWhiteSpace($DiagnosticsRoot)) { '' } else { Join-Path $DiagnosticsRoot $runId }
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
    $script:AllMigrationUsers = @()
    $script:MigrationUsersByEmail = @{}
    $script:MigrationUsersLoaded = $false
    $script:BatchActiveDirectoryEvidenceByEmail = @{}
    $script:BatchDelegateIdentityToEmail = @{}
    $script:BatchEmailSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($batchRow in $rows) {
        if ($batchRow.EmailAddress) { [void]$script:BatchEmailSet.Add(([string]$batchRow.EmailAddress).Trim()) }
    }
    if ($ProgressCallback) {
        & $ProgressCallback 0 $rows.Count '' 'Connecting to live AD and Exchange on-premises sources'
    }
    $sourceInitialization = Initialize-SemrLiveSourceConnections -Config $Config
    if ($script:ConnectionState.ActiveDirectory) {
        if ($ProgressCallback) { & $ProgressCallback 0 $rows.Count '' 'Prefetching Active Directory evidence with batched forest queries' }
        $batchAd = Initialize-SemrBatchActiveDirectoryEvidence -EmailAddresses @($rows | ForEach-Object { [string]$_.EmailAddress })
        if (-not $batchAd.Available) {
            $script:ConnectionState.ActiveDirectory = $false
            $sourceInitialization.ActiveDirectoryLive = $false
            $sourceInitialization.ActiveDirectoryMessage = "$($batchAd.Message) The assessment will be INCOMPLETE."
        }
    }

    if ($script:ConnectionState.OnPremisesExchange) {
        try {
            if ($ProgressCallback) { & $ProgressCallback 0 $rows.Count '' 'Collecting Exchange 2016 evidence through local Windows PowerShell 5.1' }
            $exchangeProgress = if ($ProgressCallback) {
                { param($Current,$Total,$Message) & $ProgressCallback $Current $Total '' $Message }.GetNewClosure()
            }
            else { $null }
            $batchExchange = Initialize-SemrExchange2016Evidence -EmailAddresses @($rows | ForEach-Object { [string]$_.EmailAddress }) -ProgressCallback $exchangeProgress -CancellationCheck $CancellationCheck -DiagnosticsDirectory $exchangeDiagnosticsDirectory
            $sourceInitialization.ExchangeOnPremisesLive = [bool]$batchExchange.Available
            $sourceInitialization.ExchangeOnPremisesMessage = [string]$batchExchange.Message
        }
        catch {
            if ($CancellationCheck -and (& $CancellationCheck)) {
                throw [OperationCanceledException]::new('Exchange migration readiness assessment cancelled by the operator.')
            }
            $script:ConnectionState.OnPremisesExchange = $false
            $script:Exchange2016EvidenceByEmail = @{}
            $sourceInitialization.ExchangeOnPremisesLive = $false
            $sourceInitialization.ExchangeOnPremisesMessage = "Local Exchange 2016 evidence collection failed. The assessment will be INCOMPLETE. $($_.Exception.Message)"
        }
    }
    if ($ProgressCallback) { & $ProgressCallback 0 $rows.Count '' 'Testing hybrid migration endpoint' }
    $hybrid = Test-SemrHybridReadiness -Config $Config
    if ($ProgressCallback) { & $ProgressCallback 0 $rows.Count '' 'Collecting accepted domains and advanced identity evidence' }
    $acceptedDomains = Get-SemrAcceptedDomainEvidence
    if ($ProgressCallback) { & $ProgressCallback 0 $rows.Count '' 'Checking latest tenant synchronization through Microsoft Graph' }
    $entraConnect = Test-SemrEntraConnect -Config $Config
    $checkedAt = Get-Date
    $liveSources = @(
        [pscustomobject][ordered]@{ Source='Active Directory'; Required=$true; Available=[bool]$script:ConnectionState.ActiveDirectory; Status=if($script:ConnectionState.ActiveDirectory){'Connected'}else{'Unavailable'}; Details=[string]$sourceInitialization.ActiveDirectoryMessage; CheckedAt=$checkedAt }
        [pscustomobject][ordered]@{ Source='Exchange on-premises'; Required=$true; Available=[bool]$script:ConnectionState.OnPremisesExchange; Status=if($script:ConnectionState.OnPremisesExchange){'Connected'}else{'Unavailable'}; Details=[string]$sourceInitialization.ExchangeOnPremisesMessage; CheckedAt=$checkedAt }
        [pscustomobject][ordered]@{ Source='Exchange Online'; Required=$true; Available=[bool]$script:ConnectionState.ExchangeOnline; Status=if($script:ConnectionState.ExchangeOnline){'Connected'}else{'Unavailable'}; Details=if($script:ConnectionState.ExchangeOnline){'Interactive delegated Exchange Online session connected.'}else{'Exchange Online is not connected.'}; CheckedAt=$checkedAt }
        [pscustomobject][ordered]@{ Source='Microsoft Graph'; Required=$true; Available=[bool]$script:ConnectionState.MicrosoftGraph; Status=if($script:ConnectionState.MicrosoftGraph){'Connected'}else{'Unavailable'}; Details=if($script:ConnectionState.MicrosoftGraph){'Interactive delegated Microsoft Graph evidence collected.'}else{'Microsoft Graph is not connected.'}; CheckedAt=$checkedAt }
        [pscustomobject][ordered]@{ Source='Microsoft Entra tenant synchronization'; Required=$true; Available=[bool]$entraConnect.Available; Status=if($entraConnect.Available){'Collected'}else{'Unavailable'}; Details=[string]$entraConnect.Message; CheckedAt=$checkedAt }
    )
    $assessmentStatus = if (@($liveSources | Where-Object { $_.Required -and -not $_.Available }).Count -eq 0) { 'COMPLETE' } else { 'INCOMPLETE' }
    $licenseCapacityCache = @{}
    $batchLicenseRequirements = @{}
    $globalCheckIds = @('HYBRID-ENDPOINT','HYBRID-MRSPROXY','HYBRID-CERTIFICATE-EXPIRY','HYBRID-ENDPOINT-CAPACITY','HYBRID-MIGRATION-BACKLOG','HYBRID-AUTODISCOVER-OAUTH','ENTRA-CONNECT-SCHEDULER','LICENSE-BATCH-CAPACITY')
    $enabledCheckDefinitions = @(Get-SemrCheckCatalog | Where-Object { Test-SemrCheckEnabled -CheckId $_.CheckId })
    $enabledMailboxCheckDefinitions = @($enabledCheckDefinitions | Where-Object { $_.CheckId -notin $globalCheckIds })
    $enabledGlobalCheckDefinitions = @($enabledCheckDefinitions | Where-Object { $_.CheckId -in $globalCheckIds })
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
        $targetLicenseAssigned = $false
        $mailboxSizeGb = $null
        $targetLicense = if ($row.TargetSku) { [string]$row.TargetSku } else { [string]$Config['DefaultTargetSku'] }
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
            RecommendedAction = if ($ad.Available) { '' } else { 'Restore live AD forest access, then rerun.' }
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
            RecommendedAction = if ($onPrem.Available) { '' } else { 'Restore the local Exchange 2016 worker and ViewEntireForest collection, then rerun.' }
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
                $mailboxSizeGb = if ($statistics.Count -eq 1) { [math]::Round([double]$sizeGb, 2) } else { $null }
                $targetLicense = [string]$mailboxTargetPolicy.TargetSku
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
                    ObservedValue = if ($holdDataAvailable) { "LitigationHold=$litigationHold; InPlaceHolds=$($inPlaceHolds.Count)" } else { 'Hold properties were not returned by the live Exchange on-premises query.' }
                    ExpectedValue = 'Hold state documented'; EvidenceSource = $onPremSource
                    Message = if (-not $holdDataAvailable) { 'Hold state could not be collected from live Exchange on-premises.' } elseif ($holdCount -gt 0) { 'Mailbox hold state was detected. Hybrid moves normally preserve supported holds and Recoverable Items.' } else { 'No Litigation Hold or In-Place Hold was returned.' }
                    RecommendedAction = if (-not $holdDataAvailable) { 'Validate hold state through live Exchange on-premises before migration.' } elseif ($holdCount -gt 0) { 'Confirm hold preservation requirements and monitor Recoverable Items quota; do not remove holds solely to satisfy this precheck.' } else { '' }
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
                $knownPermissionText = @($onPrem.Permissions | ForEach-Object { "$($_.PermissionType):$($_.Delegate)" } | Sort-Object -Unique) -join '; '
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'PERMISSIONS-BASELINE'; Category = 'Permissions'; Severity = if (-not $permissionsAvailable -or $permissionCount -gt 0) { 'Warning' } else { 'Information' }
                    Result = if (-not $permissionsAvailable) { 'UNKNOWN' } elseif ($permissionCount -gt 0) { 'WARN' } else { 'PASS' }; IsBlocking = $false
                    ObservedValue = if ($knownPermissionText) { "$permissionCount known explicit permission grant(s): $knownPermissionText" } else { "$permissionCount explicit permission grant(s)" }; ExpectedValue = 'Baseline captured'
                    EvidenceSource = $onPremSource
                    Message = if ($permissionsAvailable) { 'Explicit Full Access, Send As, and Send on Behalf permissions were captured.' } elseif ($permissionCount -gt 0) { 'The baseline is incomplete, but the known delegated permissions are preserved in the report.' } else { 'The live Exchange on-premises permission queries did not return a complete baseline.' }
                    RecommendedAction = if (-not $permissionsAvailable) { 'Restore Full Access and Send As permission queries on live Exchange on-premises, then rerun.' } elseif ($permissionCount -gt 0) { 'Migrate dependent delegates in a coordinated wave and compare permissions after migration.' } else { '' }
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
            Message = if ($exo.Available) { 'Exchange Online evidence is available.' } else { 'Exchange Online is not connected.' }
            RecommendedAction = if ($exo.Available) { '' } else { 'Connect Exchange Online interactively and run the assessment again.' }
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
                ObservedValue = if ($softDeletedAvailable) { "$softDeletedCount soft-deleted/inactive mailbox match(es)" } else { 'The live soft-deleted and inactive mailbox query is unavailable.' }; ExpectedValue = '0'
                EvidenceSource = $exoSource; Message = if (-not $softDeletedAvailable) { 'The live Exchange Online soft-deleted mailbox query did not complete.' } elseif ($softDeletedCount -gt 0) { 'A soft-deleted or inactive mailbox may conflict with the migration identity.' } else { 'No soft-deleted or inactive mailbox conflict was returned.' }
                RecommendedAction = if (-not $softDeletedAvailable) { 'Restore the live Exchange Online query before creating the migration batch.' } elseif ($softDeletedCount -gt 0) { 'Investigate the soft-deleted/inactive object and GUID ownership before creating the batch.' } else { '' }
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
                RecommendedAction = if (-not $moveDataAvailable) { 'Restore the live Exchange Online migration-user and move-request queries.' } elseif ($existingBatchPhase -and $activeMoveCount -eq 0) { 'Confirm that the mailbox still belongs to the intended migration batch or review its terminal history.' } elseif (-not $existingBatchPhase -and $activeMoveCount -gt 0) { 'Review the active operation and clean it up only after confirming its state and history.' } else { '' }
            })
        }

        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'GRAPH-SOURCE'; Category = 'MicrosoftGraph'; Severity = if ($graph.Available) { 'Information' } else { 'Critical' }
            Result = if ($graph.Available) { 'PASS' } else { 'UNKNOWN' }; IsBlocking = -not $graph.Available
            ObservedValue = if ($graph.Available) { 'Available' } else { 'Unavailable' }; ExpectedValue = 'Available Microsoft Graph evidence'; EvidenceSource = $graphSource
            Message = if ($graph.Available) { 'Microsoft Graph evidence is available.' } else { 'Microsoft Graph is not connected.' }
            RecommendedAction = if ($graph.Available) { '' } else { 'Connect Microsoft Graph interactively and run the assessment again.' }
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
                $upnDomainEvidence = Get-SemrUpnVerifiedDomainEvidence -UserPrincipalName ([string]$user.UserPrincipalName)
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'ENTRA-UPN-VERIFIED-DOMAIN'; Category = 'MicrosoftGraph'; Severity = if (-not $upnDomainEvidence.Available) { 'Warning' } elseif ($upnDomainEvidence.Verified) { 'Information' } else { 'Critical' }
                    Result = if (-not $upnDomainEvidence.Available) { 'UNKNOWN' } elseif ($upnDomainEvidence.Verified) { 'PASS' } else { 'FAIL' }; IsBlocking = $upnDomainEvidence.Available -and -not $upnDomainEvidence.Verified
                    ObservedValue = if ($upnDomainEvidence.Available) { "UPN=$($upnDomainEvidence.UserPrincipalName); Domain=$($upnDomainEvidence.UpnDomain)" } else { $upnDomainEvidence.Message }
                    ExpectedValue = 'The Entra user UPN domain is present in organization.verifiedDomains'; EvidenceSource = $upnDomainEvidence.Source; SourceTimestamp = $upnDomainEvidence.SourceTimestamp
                    Message = $upnDomainEvidence.Message
                    RecommendedAction = if (-not $upnDomainEvidence.Available) { 'Restore the Microsoft Graph organization verifiedDomains query.' } elseif (-not $upnDomainEvidence.Verified) { 'Verify the UPN domain in Microsoft Entra or correct the synchronized userPrincipalName before migration.' } else { '' }
                })
                $licenseNames = @($graph.LicenseDetails | ForEach-Object { [string]$_.SkuPartNumber } | Where-Object { $_ } | Sort-Object -Unique)
                $licenseCount = $licenseNames.Count
                $existingBatchPhase = [string]$Config['AssessmentPhase'] -eq 'ExistingBatch'
                $targetLicenseAssigned = $licenseNames -contains $targetLicense
                $activeCloudMailboxCount = @($exo.Mailboxes).Count
                $stagedTargetLicense = -not $existingBatchPhase -and $targetLicenseAssigned -and $activeCloudMailboxCount -eq 0
                $licenseResult = if ($existingBatchPhase) { if ($licenseCount -gt 0) { 'PASS' } else { 'WARN' } } elseif ($licenseCount -eq 0 -or $stagedTargetLicense) { 'PASS' } else { 'WARN' }
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'LICENSE-PRE-MIGRATION'; Category = 'Licensing'; Severity = if ($licenseResult -eq 'WARN') { 'Warning' } else { 'Information' }
                    Result = $licenseResult; IsBlocking = $false
                    ObservedValue = if ($licenseCount -gt 0) { "$licenseCount assigned license(s): $($licenseNames -join '; '); Target=$targetLicense; ActiveCloudMailbox=$activeCloudMailboxCount" } else { '0 assigned licenses' }; ExpectedValue = if ($existingBatchPhase) { 'Approved target license assigned for the existing migration batch' } else { 'No license, or the approved target license staged without an active cloud mailbox' }
                    EvidenceSource = $graphSource; Message = if ($existingBatchPhase -and $licenseCount -gt 0) { 'The assigned license is expected for the existing migration batch.' } elseif ($existingBatchPhase) { 'No assigned license was returned for the existing migration batch.' } elseif ($stagedTargetLicense) { 'The approved target license is staged and no active cloud mailbox was returned.' } elseif ($licenseCount -gt 0) { 'Assigned licenses differ from the planned target or require review before migration.' } else { 'No license is assigned before migration; this is valid during pre-creation.' }
                    RecommendedAction = if ($existingBatchPhase -and $licenseCount -eq 0) { 'Confirm the target licensing plan and assignment timing for this active migration.' } elseif (-not $existingBatchPhase -and $licenseCount -gt 0 -and -not $targetLicenseAssigned) { "Align the assigned mailbox license with the approved target '$targetLicense', then rerun the assessment." } elseif (-not $existingBatchPhase -and $licenseCount -gt 0 -and $activeCloudMailboxCount -gt 0) { 'Resolve the active cloud mailbox state before migration.' } else { '' }
                })

                $assignedQuotaCandidates = @()
                $licensePlanEvidenceAvailable = @($graph.LicenseDetails | Where-Object { $_.PSObject.Properties['ServicePlans'] }).Count -gt 0
                $mailboxBearingLicenseNames = if ($licensePlanEvidenceAvailable) {
                    @($graph.LicenseDetails | Where-Object {
                        $_.PSObject.Properties['ServicePlans'] -and @($_.ServicePlans | Where-Object {
                            (Test-SemrMailboxServicePlanName -Name ([string]$_.ServicePlanName)) -and
                            ([string]::IsNullOrWhiteSpace([string]$_.ProvisioningStatus) -or [string]$_.ProvisioningStatus -ieq 'Success')
                        }).Count -gt 0
                    } | ForEach-Object { [string]$_.SkuPartNumber } | Sort-Object -Unique)
                } else { @() }
                $assignedQuotaLicenseNames = if ($licensePlanEvidenceAvailable) { $mailboxBearingLicenseNames } else { $licenseNames }
                $quotaMap = $Config['TargetQuotaGbBySku']
                if ($quotaMap -is [System.Collections.IDictionary]) {
                    $assignedQuotaCandidates = @(
                        foreach ($licenseName in $assignedQuotaLicenseNames) {
                            foreach ($quotaKey in $quotaMap.Keys) {
                                if ([string]$quotaKey -ieq $licenseName) {
                                    [pscustomobject]@{
                                        SkuPartNumber = $licenseName
                                        QuotaGb = [double]$quotaMap[$quotaKey]
                                    }
                                    break
                                }
                            }
                        }
                    )
                }
                $assignedQuotaKnown = $assignedQuotaCandidates.Count -gt 0
                $assignedQuotaGb = if ($assignedQuotaKnown) {
                    [double](@($assignedQuotaCandidates | Sort-Object QuotaGb -Descending | Select-Object -First 1)[0].QuotaGb)
                }
                else {
                    0.0
                }
                $assignedQuotaSkus = if ($assignedQuotaKnown) {
                    @($assignedQuotaCandidates | Where-Object { [double]$_.QuotaGb -eq $assignedQuotaGb } | ForEach-Object { [string]$_.SkuPartNumber } | Sort-Object -Unique)
                }
                else {
                    @()
                }
                $assignedSafeQuotaGb = $assignedQuotaGb * (1.0 - ([double]$Config['QuotaSafetyBufferPercent'] / 100.0))
                $assignedQuotaPass = $assignedQuotaKnown -and $null -ne $mailboxSizeGb -and [double]$mailboxSizeGb -lt $assignedSafeQuotaGb
                $assignedQuotaResult = if ($licenseCount -eq 0) {
                    'PASS'
                }
                elseif ($null -eq $mailboxSizeGb -or -not $assignedQuotaKnown) {
                    'UNKNOWN'
                }
                elseif ($assignedQuotaPass) {
                    'PASS'
                }
                else {
                    'FAIL'
                }
                $assignedQuotaBlocking = $assignedQuotaResult -eq 'FAIL'
                Add-SemrFinding -List $findings -Parameters ($base + @{
                    CheckId = 'LICENSE-ASSIGNED-MAILBOX-QUOTA'; Category = 'Licensing'
                    Severity = if ($assignedQuotaBlocking) { 'Critical' } elseif ($assignedQuotaResult -eq 'UNKNOWN') { 'Warning' } else { 'Information' }
                    Result = $assignedQuotaResult; IsBlocking = $assignedQuotaBlocking
                    ObservedValue = if ($licenseCount -eq 0) {
                        'No currently assigned license'
                    }
                    elseif (-not $assignedQuotaKnown) {
                        "AssignedLicenses=$($licenseNames -join '; '); mapped mailbox quota unavailable"
                    }
                    elseif ($null -eq $mailboxSizeGb) {
                        "AssignedLicenses=$($licenseNames -join '; '); mailbox size unavailable"
                    }
                    else {
                        "Mailbox=$mailboxSizeGb GB; AssignedLicenses=$($licenseNames -join '; '); EffectiveQuota=$assignedQuotaGb GB; SafeQuota=$([math]::Round($assignedSafeQuotaGb, 2)) GB"
                    }
                    ExpectedValue = if ($assignedQuotaKnown) {
                        "Mailbox below $([math]::Round($assignedSafeQuotaGb, 2)) GB for assigned mailbox license $($assignedQuotaSkus -join '; ')"
                    }
                    else {
                        'Assigned mailbox license represented in TargetQuotaGbBySku'
                    }
                    EvidenceSource = "$graphSource assigned licenses + $onPremSource mailbox statistics"
                    Message = if ($licenseCount -eq 0) {
                        'No current mailbox-license quota applies; the planned target-license checks remain authoritative.'
                    }
                    elseif (-not $assignedQuotaKnown) {
                        'The quota of the currently assigned licenses could not be determined.'
                    }
                    elseif ($null -eq $mailboxSizeGb) {
                        'Mailbox size is unavailable, so the currently assigned license quota could not be validated.'
                    }
                    elseif ($assignedQuotaPass) {
                        'Mailbox size fits the quota of the currently assigned mailbox license.'
                    }
                    else {
                        'Mailbox size exceeds the safe quota of the currently assigned mailbox license.'
                    }
                    RecommendedAction = if ($assignedQuotaBlocking) {
                        if ($targetLicenseAssigned) { "Select a mailbox license with sufficient quota, or reduce the mailbox below $([math]::Round($assignedSafeQuotaGb, 2)) GB, then rerun the assessment." }
                        else { "Assign the approved target license '$targetLicense' before starting the migration, or reduce the mailbox below $([math]::Round($assignedSafeQuotaGb, 2)) GB, then rerun the assessment." }
                    }
                    elseif ($licenseCount -gt 0 -and -not $assignedQuotaKnown) {
                        'Add the mailbox-bearing assigned SKU and its quota to TargetQuotaGbBySku, then rerun the assessment.'
                    }
                    elseif ($null -eq $mailboxSizeGb) {
                        'Collect mailbox statistics and rerun the assessment.'
                    }
                    else {
                        ''
                    }
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
                    RecommendedAction = if ($servicePlanPass) { '' } elseif ($licenseCapacity -and -not $licenseCapacity.ServicePlanDataAvailable) { 'Restore the live Microsoft Graph subscribedSkus service-plan query.' } else { 'Select a mailbox-eligible SKU with an enabled Exchange Online service plan.' }
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
            RecommendedAction = if ($hybridEndpointResult -eq 'PASS') { '' } elseif (-not $hybrid.Available) { 'Restore the live Exchange Online migration endpoint cmdlets.' } elseif (-not $hybrid.EndpointFound) { 'Select or repair the intended ExchangeRemoteMove endpoint.' } elseif (-not $hybrid.ConnectivityTestAvailable) { 'Run Test-MigrationServerAvailability from a supported delegated Exchange Online session before creating the batch.' } else { 'Validate the endpoint, MRSProxy, WSSecurity, certificate, DNS and HTTPS connectivity.' }
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-MRSPROXY';Category='HybridConnectivity';Severity=if(-not $hybrid.MrsProxyAvailable){'Warning'}elseif($hybrid.MrsProxyEnabled){'Information'}else{'Critical'};Result=if(-not $hybrid.MrsProxyAvailable){'UNKNOWN'}elseif($hybrid.MrsProxyEnabled){'PASS'}else{'FAIL'};IsBlocking=$hybrid.MrsProxyAvailable -and -not $hybrid.MrsProxyEnabled
            ObservedValue=$hybrid.MrsProxyMessage;ExpectedValue='MRSProxy enabled on a published EWS virtual directory';EvidenceSource=$hybrid.MrsProxySource;SourceTimestamp=$hybrid.MrsProxySourceTimestamp;Message=$hybrid.MrsProxyMessage;RecommendedAction=if(-not $hybrid.MrsProxyAvailable){'Validate MRSProxy live before creating the batch.'}elseif(-not $hybrid.MrsProxyEnabled){'Enable and publish MRSProxy on the intended EWS endpoint.'}else{''}
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-CERTIFICATE-EXPIRY';Category='HybridConnectivity';Severity=if(-not $hybrid.CertificateAvailable){'Warning'}elseif($hybrid.CertificateHealthy){'Information'}else{'Critical'};Result=if(-not $hybrid.CertificateAvailable){'UNKNOWN'}elseif($hybrid.CertificateHealthy){'PASS'}else{'FAIL'};IsBlocking=$hybrid.CertificateAvailable -and -not $hybrid.CertificateHealthy
            ObservedValue=$hybrid.CertificateMessage;ExpectedValue='Valid TLS certificate presented by the selected migration endpoint with at least 60 days remaining';EvidenceSource=$hybrid.CertificateSource;SourceTimestamp=$hybrid.CertificateSourceTimestamp;Message=$hybrid.CertificateMessage;RecommendedAction=if(-not $hybrid.CertificateAvailable){'Restore direct TLS access to the selected endpoint and rerun the certificate probe.'}elseif(-not $hybrid.CertificateHealthy){'Renew or correct the TLS certificate presented by the selected migration endpoint.'}else{''}
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-ENDPOINT-CAPACITY';Category='HybridConnectivity';Severity=if(-not $hybrid.CapacityAvailable){'Warning'}elseif($hybrid.CapacityHealthy){'Information'}else{'Warning'};Result=if(-not $hybrid.CapacityAvailable){'UNKNOWN'}elseif($hybrid.CapacityHealthy){'PASS'}else{'WARN'};IsBlocking=$false
            ObservedValue=$hybrid.CapacityMessage;ExpectedValue='Active migration load below advisory threshold';EvidenceSource=$hybrid.CapacitySource;SourceTimestamp=$hybrid.CapacitySourceTimestamp;Message=$hybrid.CapacityMessage;RecommendedAction=if(-not $hybrid.CapacityAvailable){'Review current migration load before starting the batch.'}elseif(-not $hybrid.CapacityHealthy){'Reduce or stagger concurrent migration waves.'}else{''}
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-MIGRATION-BACKLOG';Category='HybridConnectivity';Severity=if(-not $hybrid.FailureBacklogAvailable){'Warning'}elseif($hybrid.FailureBacklogHealthy){'Information'}else{'Warning'};Result=if(-not $hybrid.FailureBacklogAvailable){'UNKNOWN'}elseif($hybrid.FailureBacklogHealthy){'PASS'}else{'WARN'};IsBlocking=$false
            ObservedValue=$hybrid.FailureBacklogMessage;ExpectedValue='No failed, stopped or corrupted migration user backlog';EvidenceSource=$hybrid.FailureBacklogSource;SourceTimestamp=$hybrid.FailureBacklogSourceTimestamp;Message=$hybrid.FailureBacklogMessage;RecommendedAction=if(-not $hybrid.FailureBacklogAvailable){'Review failed migration objects before starting the batch.'}elseif(-not $hybrid.FailureBacklogHealthy){'Review and resolve failed, stopped or corrupted migration users separately from active capacity.'}else{''}
        })
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId='HYBRID-AUTODISCOVER-OAUTH';Category='HybridConnectivity';Severity=if($hybrid.OAuthHealthy){'Information'}else{'Warning'};Result=if(-not $hybrid.OAuthAvailable){'UNKNOWN'}elseif($hybrid.OAuthHealthy){'PASS'}else{'WARN'};IsBlocking=$false
            ObservedValue=$hybrid.OAuthMessage;ExpectedValue='OAuth enabled and an enabled intra-organization connector for hybrid collaboration features';EvidenceSource=$hybrid.OAuthSource;SourceTimestamp=$hybrid.OAuthSourceTimestamp;Message=if($hybrid.OAuthHealthy){$hybrid.OAuthMessage}else{"$($hybrid.OAuthMessage) This does not block ExchangeRemoteMove migrations."};RecommendedAction=if(-not $hybrid.OAuthAvailable){'Validate hybrid OAuth and Autodiscover for free/busy and other hybrid collaboration features.'}elseif(-not $hybrid.OAuthHealthy){'Repair hybrid OAuth and the intra-organization connector for hybrid collaboration features; migration readiness remains non-blocking.'}else{''}
        })
        $entraSyncEnabledKnown = $null -ne $entraConnect.SyncCycleEnabled
        $entraLastSyncKnown = $entraConnect.PSObject.Properties['LastSyncKnown'] -and [bool]$entraConnect.LastSyncKnown
        $entraResult = if (-not $entraConnect.Available) { 'UNKNOWN' }
            elseif (-not $entraSyncEnabledKnown) { 'UNKNOWN' }
            elseif (-not [bool]$entraConnect.SyncCycleEnabled) { 'FAIL' }
            elseif ($entraConnect.SchedulerSuspended -eq $true) { 'FAIL' }
            elseif (-not $entraLastSyncKnown) { 'UNKNOWN' }
            elseif (-not [bool]$entraConnect.LastSyncFresh) { 'FAIL' }
            else { 'PASS' }
        Add-SemrFinding -List $globalFindings -Parameters ($globalBase + @{
            CheckId = 'ENTRA-CONNECT-SCHEDULER'; Category = 'MicrosoftEntra'; Severity = if ($entraResult -in @('FAIL','UNKNOWN')) { 'Critical' } else { 'Information' }
            Result = $entraResult; IsBlocking = $entraResult -ne 'PASS'
            ObservedValue = $entraConnect.Message; ExpectedValue = "Tenant synchronization enabled; last sync no older than $([double]$Config['EntraConnectHealth']['MaximumLastSyncAgeMinutes']) minutes"; EvidenceSource = [string]$entraConnect.Source; SourceTimestamp = $entraConnect.SourceTimestamp
            Message = if (-not $entraConnect.Available) { 'Authoritative tenant synchronization health could not be collected from Microsoft Graph.' } elseif (-not $entraSyncEnabledKnown) { 'Microsoft Graph did not return the tenant synchronization-enabled state.' } elseif (-not [bool]$entraConnect.SyncCycleEnabled) { 'On-premises directory synchronization is disabled for the tenant.' } elseif ($entraConnect.SchedulerSuspended -eq $true) { 'The tenant synchronization service reports a suspended or unhealthy state.' } elseif (-not $entraLastSyncKnown) { 'The last tenant synchronization timestamp is unavailable.' } elseif (-not [bool]$entraConnect.LastSyncFresh) { "The last tenant synchronization is $($entraConnect.LastSyncAgeMinutes) minutes old, beyond the configured threshold." } else { "Tenant synchronization is enabled and recent according to $($entraConnect.Source)." }
            RecommendedAction = if (-not $entraConnect.Available) { 'Restore Microsoft Graph organization read access and rerun.' } elseif (-not $entraSyncEnabledKnown -or -not $entraLastSyncKnown) { 'Verify the tenant synchronization state in Microsoft Entra and rerun after the next successful synchronization.' } elseif (-not [bool]$entraConnect.SyncCycleEnabled -or $entraConnect.SchedulerSuspended -eq $true) { 'Restore tenant directory synchronization before migration.' } elseif (-not [bool]$entraConnect.LastSyncFresh) { 'Investigate the synchronization backlog or connector health before starting the migration wave.' } else { '' }
        })
        }
        $enhancedRestoreReportAvailable = [bool](Get-SemrPropertyValue -InputObject $exo -Names @('MigrationReportDataAvailable') -Default $false)
        $enhancedRestoreReportText = [string](Get-SemrPropertyValue -InputObject $exo -Names @('MigrationReportText') -Default '')
        $enhancedRestoreDetected = $enhancedRestoreReportText -match 'CannotMoveEnhancedRestoreMailboxesCrossOrgPermanentException|EnhancedRestore'
        Add-SemrFinding -List $findings -Parameters ($base + @{
            CheckId = 'KNOWN-ENHANCED-RESTORE'; Category = 'KnownErrors'
            Severity = if ($enhancedRestoreDetected) { 'Critical' } elseif ($enhancedRestoreReportAvailable) { 'Information' } else { 'Warning' }
            Result = if ($enhancedRestoreDetected) { 'FAIL' } elseif ($enhancedRestoreReportAvailable) { 'PASS' } else { 'UNKNOWN' }
            IsBlocking = $enhancedRestoreDetected
            ObservedValue = if ($enhancedRestoreDetected) { 'Enhanced Restore cross-organization move exception found in migration statistics/report.' } elseif ($enhancedRestoreReportAvailable) { 'No Enhanced Restore exception found in available current migration evidence.' } else { 'Migration statistics/report unavailable for an existing migration object.' }
            ExpectedValue = 'No enhanced restore mailbox restriction'
            EvidenceSource = if ($enhancedRestoreReportAvailable) { "$($exo.Source) migration statistics/report" } else { 'Known error catalog + unavailable migration report' }
            Message = if ($enhancedRestoreDetected) { 'A migration report contains the Enhanced Restore cross-organization move exception.' } elseif ($enhancedRestoreReportAvailable) { 'Available current migration evidence contains no Enhanced Restore exception.' } else { 'Enhanced Restore risk could not be evaluated because detailed migration report evidence is unavailable.' }
            RecommendedAction = if ($enhancedRestoreDetected) { 'Treat this mailbox as NO-GO and confirm Microsoft-supported remediation before retrying the cross-organization move.' } elseif (-not $enhancedRestoreReportAvailable) { 'Collect Get-MigrationUserStatistics -IncludeReport evidence for the current migration object.' } else { '' }
        })

        $mailboxCheckIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($mailboxFinding in @($findings | Where-Object { [string]$_.EmailAddress -ieq $email })) { [void]$mailboxCheckIds.Add([string]$mailboxFinding.CheckId) }
        foreach ($missingDefinition in @($enabledMailboxCheckDefinitions | Where-Object { -not $mailboxCheckIds.Contains([string]$_.CheckId) })) {
            Add-SemrFinding -List $findings -Parameters ($base + @{
                CheckId = [string]$missingDefinition.CheckId
                Category = [string]$missingDefinition.Category
                Severity = 'Warning'
                Result = 'UNKNOWN'
                IsBlocking = [bool]$missingDefinition.Mandatory
                ObservedValue = 'Enabled check did not produce a finding'
                ExpectedValue = [string]$missingDefinition.Name
                EvidenceSource = 'Assessment execution coverage control'
                Message = "The enabled check '$($missingDefinition.CheckId)' did not produce a mailbox finding. Required evidence may be unavailable or the evaluation path may be incomplete."
                RecommendedAction = 'Review the required source evidence and rerun. If this persists with complete sources, treat it as an application defect.'
            })
        }
        $graphUserForReport = @($graph.Users | Select-Object -First 1)
        $assignedLicenseNames = @($graph.LicenseDetails | ForEach-Object { [string]$_.SkuPartNumber } | Where-Object { $_ } | Sort-Object -Unique)
        $graphReportAvailable = [string]::IsNullOrWhiteSpace([string](Get-SemrPropertyValue -InputObject $graph -Names @('QueryError') -Default '')) -and $graphUserForReport.Count -eq 1
        if ($mailboxTargetPolicy -and $mailboxTargetPolicy.RequiresLicense -and -not $targetLicenseAssigned) {
            $requirementSku = ([string]$mailboxTargetPolicy.TargetSku).Trim().ToUpperInvariant()
            if ($requirementSku) {
                if (-not $batchLicenseRequirements.ContainsKey($requirementSku)) {
                    $batchLicenseRequirements[$requirementSku] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                }
                [void]$batchLicenseRequirements[$requirementSku].Add($email)
            }
        }

        $mailboxEvidence = [pscustomobject][ordered]@{
            RunId = $runId
            EmailAddress = $email
            UserPrincipalName = if ($graphUserForReport.Count -eq 1) { [string]$graphUserForReport[0].UserPrincipalName } else { '' }
            MailboxSizeGb = if ($null -ne $mailboxSizeGb) { [double]$mailboxSizeGb } else { $null }
            TargetLicense = $targetLicense
            AssignedLicenses = if (-not $graphReportAvailable) { 'Not available' } elseif ($assignedLicenseNames.Count -gt 0) { $assignedLicenseNames -join '; ' } else { 'None' }
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


    foreach ($licenseSku in @($batchLicenseRequirements.Keys | Sort-Object)) {
        $requiredCount = $batchLicenseRequirements[$licenseSku].Count
        if ($requiredCount -le 0) { continue }
        if (-not $licenseCapacityCache.ContainsKey($licenseSku)) {
            $licenseCapacityCache[$licenseSku] = Get-SemrTenantLicenseEvidence -TargetSku $licenseSku -Config $Config
        }
        $batchCapacity = $licenseCapacityCache[$licenseSku]
        $batchCapacityResult = if (-not $batchCapacity.Available) { 'UNKNOWN' } elseif (-not $batchCapacity.Found -or $batchCapacity.AvailableUnits -lt $requiredCount) { 'FAIL' } else { 'PASS' }
        $batchCapacityShortfall = if ($batchCapacity.Available) { [math]::Max(0, $requiredCount - [int]$batchCapacity.AvailableUnits) } else { $null }
        Add-SemrFinding -List $globalFindings -Parameters @{
            RunId = $runId; EmailAddress = ''; CheckId = 'LICENSE-BATCH-CAPACITY'; Category = 'Licensing'
            Severity = if ($batchCapacityResult -eq 'PASS') { 'Information' } else { 'Critical' }
            Result = $batchCapacityResult; IsBlocking = $batchCapacityResult -ne 'PASS'
            ObservedValue = if ($batchCapacity.Available) { "SKU=$licenseSku; Required=$requiredCount; Available=$($batchCapacity.AvailableUnits); Shortfall=$batchCapacityShortfall" } else { "SKU=$licenseSku; Required=$requiredCount; capacity evidence unavailable" }
            ExpectedValue = "At least $requiredCount available $licenseSku license(s) for unlicensed batch recipients"
            EvidenceSource = "$($batchCapacity.Source) + complete batch requirement"; SourceTimestamp = $batchCapacity.SourceTimestamp
            Message = if ($batchCapacityResult -eq 'PASS') { "The tenant has enough $licenseSku capacity for the complete batch." } elseif ($batchCapacityResult -eq 'FAIL') { "The complete batch requires $requiredCount $licenseSku license(s), but only $($batchCapacity.AvailableUnits) are available." } else { "Complete-batch license capacity could not be verified for $licenseSku." }
            RecommendedAction = if ($batchCapacityResult -eq 'FAIL') { "Add at least $batchCapacityShortfall $licenseSku license(s), reduce the wave, or assign another approved mailbox-eligible SKU." } elseif ($batchCapacityResult -eq 'UNKNOWN') { 'Restore subscribed SKU evidence and rerun before approving the complete batch.' } else { '' }
        }
    }

    if ((Test-SemrCheckEnabled -CheckId 'LICENSE-BATCH-CAPACITY') -and $batchLicenseRequirements.Count -eq 0) {
        Add-SemrFinding -List $globalFindings -Parameters @{
            RunId = $runId; EmailAddress = ''; CheckId = 'LICENSE-BATCH-CAPACITY'; Category = 'Licensing'
            Severity = 'Information'; Result = 'PASS'; IsBlocking = $false
            ObservedValue = 'No unlicensed batch recipient requires target SKU capacity'
            ExpectedValue = 'Sufficient target SKU capacity for all recipients that require a license'
            EvidenceSource = 'Complete batch target-license requirements'
            Message = 'No additional target license capacity is required for this batch.'
            RecommendedAction = ''
        }
    }
    $globalCheckIdsFound = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($globalFinding in @($globalFindings)) { [void]$globalCheckIdsFound.Add([string]$globalFinding.CheckId) }
    foreach ($missingDefinition in @($enabledGlobalCheckDefinitions | Where-Object { -not $globalCheckIdsFound.Contains([string]$_.CheckId) })) {
        Add-SemrFinding -List $globalFindings -Parameters @{
            RunId = $runId
            EmailAddress = ''
            CheckId = [string]$missingDefinition.CheckId
            Category = [string]$missingDefinition.Category
            Severity = 'Warning'
            Result = 'UNKNOWN'
            IsBlocking = [bool]$missingDefinition.Mandatory
            ObservedValue = 'Enabled global check did not produce a finding'
            ExpectedValue = [string]$missingDefinition.Name
            EvidenceSource = 'Assessment execution coverage control'
            Message = "The enabled global check '$($missingDefinition.CheckId)' did not produce a finding."
            RecommendedAction = 'Review the required tenant evidence and rerun. If this persists with complete sources, treat it as an application defect.'
        }
    }
    $cancelled = [bool]($CancellationCheck -and (& $CancellationCheck))
    if ($cancelled) { $assessmentStatus = 'INCOMPLETE' }
    $summary = [System.Collections.Generic.List[object]]::new()
    $summaryRows = @($rows | Group-Object { ([string]$_.EmailAddress).ToLowerInvariant() } | ForEach-Object { $_.Group[0] })
    $globalBlockingFailed = @($globalFindings | Where-Object { $_.IsBlocking -and $_.Result -eq 'FAIL' })
    $globalBlockingUnknown = @($globalFindings | Where-Object { $_.IsBlocking -and $_.Result -eq 'UNKNOWN' })
    $globalBlocking = @($globalBlockingFailed) + @($globalBlockingUnknown)
    $globalWarnings = @($globalFindings | Where-Object Result -EQ 'WARN')
    $globalUnknown = @($globalFindings | Where-Object Result -EQ 'UNKNOWN')
    $evidenceByEmail = @{}


    $findingsByEmail = @{}
    foreach ($finding in @($findings)) {
        $findingKey = ([string]$finding.EmailAddress).ToLowerInvariant()
        if (-not $findingsByEmail.ContainsKey($findingKey)) { $findingsByEmail[$findingKey] = [System.Collections.Generic.List[object]]::new() }
        [void]$findingsByEmail[$findingKey].Add($finding)
    }
    foreach ($evidenceRow in @($evidenceRows)) {
        $evidenceKey = ([string]$evidenceRow.EmailAddress).ToLowerInvariant()
        if (-not $evidenceByEmail.ContainsKey($evidenceKey)) { $evidenceByEmail[$evidenceKey] = $evidenceRow }
    }
    foreach ($row in $summaryRows) {
        $email = [string]$row.EmailAddress
        $emailKey = $email.ToLowerInvariant()
        $mailboxReport = if ($evidenceByEmail.ContainsKey($emailKey)) { $evidenceByEmail[$emailKey] } else { $null }
        $mailFindings = if ($findingsByEmail.ContainsKey($emailKey)) { @($findingsByEmail[$emailKey]) } else { @() }
        $mailboxBlockingFailed = @($mailFindings | Where-Object { $_.IsBlocking -and $_.Result -eq 'FAIL' })
        $mailboxBlockingUnknown = @($mailFindings | Where-Object { $_.IsBlocking -and $_.Result -eq 'UNKNOWN' })
        $mailboxBlocking = @($mailboxBlockingFailed) + @($mailboxBlockingUnknown)
        $blocking = @($mailboxBlocking) + @($globalBlocking)
        $warnings = @($mailFindings | Where-Object Result -EQ 'WARN')
        $unknown = @($mailFindings | Where-Object Result -EQ 'UNKNOWN')
        $decision = Get-SemrMailboxDecision -MailboxFindings $mailFindings -GlobalFindings @($globalFindings)
        if ($assessmentStatus -eq 'INCOMPLETE' -and $decision -in @('GO','GO-WARNING')) { $decision = 'UNKNOWN' }
        $blockingCodes = @($blocking | ForEach-Object { $_.CheckId } | Sort-Object -Unique)
        $actionableFindings = @($mailboxBlockingFailed) + @($globalBlockingFailed) + @($mailboxBlockingUnknown) + @($globalBlockingUnknown) + @($warnings) + @($unknown | Where-Object { -not $_.IsBlocking })
        $recommendedList = [System.Collections.Generic.List[string]]::new()
        $recommendedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($actionableFinding in $actionableFindings) {
            foreach ($recommendedAction in @(([string]$actionableFinding.RecommendedAction -split '\s*\|\s*') | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
                if ($recommendedSet.Add($recommendedAction)) { [void]$recommendedList.Add($recommendedAction) }
            }
        }
        $recommended = @($recommendedList)
        [void]$summary.Add([pscustomobject][ordered]@{
            RunId = $runId
            BatchName = [System.IO.Path]::GetFileNameWithoutExtension($Batch.Path)
            AssessmentPhase = [string]$Config['AssessmentPhase']
            AssessmentStatus = $assessmentStatus
            EmailAddress = $email
            UserPrincipalName = if ($mailboxReport) { [string]$mailboxReport.UserPrincipalName } else { '' }
            MailboxSizeGb = if ($mailboxReport) { $mailboxReport.MailboxSizeGb } else { $null }
            TargetLicense = if ($mailboxReport) { [string]$mailboxReport.TargetLicense } elseif ($row.TargetSku) { [string]$row.TargetSku } else { [string]$Config['DefaultTargetSku'] }
            AssignedLicenses = if ($mailboxReport) { [string]$mailboxReport.AssignedLicenses } else { 'Not available' }
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
            DataCoverage = if ($assessmentStatus -eq 'INCOMPLETE' -or @($mailboxBlocking | Where-Object CheckId -match 'SOURCE').Count -gt 0) { 'Incomplete' } elseif ($unknown.Count + $globalUnknown.Count -gt 0) { 'Partial' } else { 'Complete' }
            BlockingCodes = ($blockingCodes -join ';')
            RecommendedAction = ($recommended -join ' | ')
            CheckedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        })
    }

    $summary = @(
        $summary | Sort-Object `
            @{ Expression = {
                switch ([string]$_.Decision) {
                    'NO-GO' { 0 }
                    'UNKNOWN' { 1 }
                    'GO-WARNING' { 2 }
                    'GO' { 3 }
                    default { 4 }
                }
            } },
            @{ Expression = { [int]$_.BlockingCount }; Descending = $true },
            @{ Expression = { [int]$_.WarningCount }; Descending = $true },
            @{ Expression = { [string]$_.EmailAddress } }
    )

    $mailboxCoverageExpected = $summaryRows.Count * $enabledMailboxCheckDefinitions.Count
    $mailboxCoverageActual = @($findings).Count
    $globalCoverageExpected = $enabledGlobalCheckDefinitions.Count
    $globalCoverageActual = @($globalFindings | Group-Object CheckId).Count
    $duplicateMailboxCheckCount = @($findings | Group-Object { "{0}|{1}" -f ([string]$_.EmailAddress).ToLowerInvariant(), [string]$_.CheckId } | Where-Object Count -gt 1).Count
    $duplicateGlobalCheckCount = @($globalFindings | Group-Object CheckId | Where-Object Count -gt 1).Count
    $coverageDiagnostics = [pscustomobject][ordered]@{
        RunId = $runId
        MailboxCount = $summaryRows.Count
        EnabledMailboxCheckCount = $enabledMailboxCheckDefinitions.Count
        ExpectedMailboxFindingCount = $mailboxCoverageExpected
        ActualMailboxFindingCount = $mailboxCoverageActual
        MaterializedMailboxUnknownCount = @($findings | Where-Object EvidenceSource -EQ 'Assessment execution coverage control').Count
        EnabledGlobalCheckCount = $globalCoverageExpected
        ActualGlobalFindingCount = $globalCoverageActual
        MaterializedGlobalUnknownCount = @($globalFindings | Where-Object EvidenceSource -EQ 'Assessment execution coverage control').Count
        DuplicateMailboxCheckCount = $duplicateMailboxCheckCount
        DuplicateGlobalCheckCount = $duplicateGlobalCheckCount
        Status = if ($mailboxCoverageActual -eq $mailboxCoverageExpected -and $globalCoverageActual -eq $globalCoverageExpected -and $duplicateMailboxCheckCount -eq 0 -and $duplicateGlobalCheckCount -eq 0) { 'PASS' } else { 'REVIEW' }
    }
    return [pscustomobject]@{
        Mode = 'Live'
        AssessmentStatus = $assessmentStatus
        AssessmentPhase = [string]$Config['AssessmentPhase']
        StartedAt = $startedAt
        CompletedAt = Get-Date
        RunId = $runId
        Batch = $Batch
        Summary = @($summary)
        Findings = @($findings)
        GlobalFindings = @($globalFindings)
        PermissionsBaseline = @($permissionRows)
        Evidence = @($evidenceRows)
        Hybrid = $hybrid
        EntraConnect = $entraConnect
        LiveSources = @($liveSources)
        CheckCoverage = $coverageDiagnostics
        CheckOptions = @(Get-SemrCheckCatalog | ForEach-Object {
            [pscustomobject][ordered]@{
                CheckId = $_.CheckId; Category = $_.Category; Name = $_.Name
                Mandatory = $_.Mandatory; Enabled = Test-SemrCheckEnabled -CheckId $_.CheckId
                Description = $_.Description
            }
        })
        SourceInitialization = $sourceInitialization
        Cancelled = $cancelled
    }
}

function Export-SemrCsvFile {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$Columns
    )

    if (@($Data).Count -gt 0) {
        $originalCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::InvariantCulture
            @($Data) | Select-Object -Property $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8
        }
        finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $originalCulture
        }
        return
    }
    $header = ($Columns | ForEach-Object { '"{0}"' -f ($_ -replace '"', '""') }) -join ','
    Set-Content -LiteralPath $Path -Value $header -Encoding utf8
}
function ConvertTo-SemrXmlText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value) -replace '[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]', ''
    return [System.Security.SecurityElement]::Escape($text)
}

function ConvertTo-SemrHtmlText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-SemrExcelColumnName {
    param([Parameter(Mandatory)][int]$Index)
    $name = ''
    $value = $Index
    while ($value -gt 0) {
        $value--
        $name = [char](65 + ($value % 26)) + $name
        $value = [math]::Floor($value / 26)
    }
    return $name
}

function Write-SemrZipEntry {
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Archive,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content
    )
    $entry = $Archive.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    try { $writer.Write($Content) }
    finally { $writer.Dispose(); $stream.Dispose() }
}

function ConvertTo-SemrWorksheetXml {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data,
        [Parameter(Mandatory)][string[]]$Columns
    )
    $rows = @($Data)
    $lastColumn = Get-SemrExcelColumnName -Index ([math]::Max(1, $Columns.Count))
    $lastRow = [math]::Max(1, $rows.Count + 1)
    $builder = [Text.StringBuilder]::new()
    [void]$builder.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$builder.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:')
    [void]$builder.Append($lastColumn).Append($lastRow).Append('"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>')
    for ($columnIndex = 0; $columnIndex -lt $Columns.Count; $columnIndex++) {
        $width = [math]::Min(52, [math]::Max(11, ([string]$Columns[$columnIndex]).Length + 2))
        foreach ($row in $rows | Select-Object -First 250) {
            $property = $row.PSObject.Properties[$Columns[$columnIndex]]
            $length = if ($property -and $null -ne $property.Value) { ([string]$property.Value).Length + 2 } else { 0 }
            if ($length -gt $width) { $width = [math]::Min(52, $length) }
        }
        $position = $columnIndex + 1
        [void]$builder.Append('<col min="').Append($position).Append('" max="').Append($position).Append('" width="').Append($width.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture)).Append('" customWidth="1"/>')
    }
    [void]$builder.Append('</cols><sheetData><row r="1" ht="24" customHeight="1">')
    for ($columnIndex = 0; $columnIndex -lt $Columns.Count; $columnIndex++) {
        $cellReference = "$(Get-SemrExcelColumnName -Index ($columnIndex + 1))1"
        [void]$builder.Append('<c r="').Append($cellReference).Append('" t="inlineStr" s="1"><is><t xml:space="preserve">').Append((ConvertTo-SemrXmlText $Columns[$columnIndex])).Append('</t></is></c>')
    }
    [void]$builder.Append('</row>')
    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $excelRow = $rowIndex + 2
        [void]$builder.Append('<row r="').Append($excelRow).Append('">')
        for ($columnIndex = 0; $columnIndex -lt $Columns.Count; $columnIndex++) {
            $columnName = $Columns[$columnIndex]
            $property = $rows[$rowIndex].PSObject.Properties[$columnName]
            $value = if ($property) { $property.Value } else { $null }
            $cellReference = "$(Get-SemrExcelColumnName -Index ($columnIndex + 1))$excelRow"
            $textValue = if ($null -eq $value) { '' } else { [string]$value }
            $style = switch -Regex ($textValue.ToUpperInvariant()) {
                '^(GO|PASS|SUCCESS|TRUE)$' { 2; break }
                '^(GO-WARNING|WARN|WARNING)$' { 3; break }
                '^(NO-GO|FAIL|FAILED|FALSE)$' { 4; break }
                '^UNKNOWN$' { 5; break }
                default { if ($textValue.Length -gt 60) { 6 } else { 0 } }
            }
            $numericTypes = @([byte],[sbyte],[int16],[uint16],[int32],[uint32],[int64],[uint64],[single],[double],[decimal])
            $isNumeric = $false
            foreach ($numericType in $numericTypes) { if ($value -is $numericType) { $isNumeric = $true; break } }
            if ($isNumeric) {
                $number = [Convert]::ToString($value, [Globalization.CultureInfo]::InvariantCulture)
                [void]$builder.Append('<c r="').Append($cellReference).Append('" s="').Append($style).Append('"><v>').Append($number).Append('</v></c>')
            }
            elseif ($value -is [bool]) {
                [void]$builder.Append('<c r="').Append($cellReference).Append('" t="inlineStr" s="').Append($style).Append('"><is><t>').Append($(if ($value) { 'TRUE' } else { 'FALSE' })).Append('</t></is></c>')
            }
            else {
                [void]$builder.Append('<c r="').Append($cellReference).Append('" t="inlineStr" s="').Append($style).Append('"><is><t xml:space="preserve">').Append((ConvertTo-SemrXmlText $textValue)).Append('</t></is></c>')
            }
        }
        [void]$builder.Append('</row>')
    }
    [void]$builder.Append('</sheetData><autoFilter ref="A1:').Append($lastColumn).Append($lastRow).Append('"/><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>')
    return $builder.ToString()
}

function New-SemrExcelReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Assessment,
        [Parameter(Mandatory)][string]$Path,
        [string]$RunFolder = ''
    )
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    $actionPlan = @(
        foreach ($finding in @($Assessment.Findings) + @($Assessment.GlobalFindings)) {
            foreach ($action in @(([string]$finding.RecommendedAction -split '\s*\|\s*') | Where-Object { $_ })) {
                [pscustomobject][ordered]@{
                    Scope = if ([string]::IsNullOrWhiteSpace([string]$finding.EmailAddress)) { 'Tenant' } else { 'Mailbox' }
                    EmailAddress = [string]$finding.EmailAddress
                    Priority = if ($finding.IsBlocking) { 'Blocking' } elseif ($finding.Result -eq 'WARN') { 'Warning' } else { 'Follow-up' }
                    CheckId = [string]$finding.CheckId
                    Result = [string]$finding.Result
                    RecommendedAction = $action.Trim()
                }
            }
        }
    )
    $sheets = @(
        [pscustomobject]@{ Name='Summary'; Data=@($Assessment.Summary); Columns=@('RunId','BatchName','AssessmentPhase','AssessmentStatus','EmailAddress','UserPrincipalName','MailboxSizeGb','TargetLicense','AssignedLicenses','Decision','BlockingCount','MailboxBlockingCount','GlobalBlockingCount','WarningCount','MailboxWarningCount','GlobalWarningCount','UnknownCount','MailboxUnknownCount','GlobalUnknownCount','DataCoverage','BlockingCodes','RecommendedAction','CheckedAt') },
        [pscustomobject]@{ Name='Mailbox Findings'; Data=@($Assessment.Findings); Columns=@('RunId','EmailAddress','CheckId','Category','Severity','Result','IsBlocking','ObservedValue','ExpectedValue','EvidenceSource','SourceTimestamp','Message','RecommendedAction') },
        [pscustomobject]@{ Name='Tenant Checks'; Data=@($Assessment.GlobalFindings); Columns=@('RunId','EmailAddress','CheckId','Category','Severity','Result','IsBlocking','ObservedValue','ExpectedValue','EvidenceSource','SourceTimestamp','Message','RecommendedAction') },
        [pscustomobject]@{ Name='Permissions'; Data=@($Assessment.PermissionsBaseline); Columns=@('RunId','EmailAddress','PermissionType','Delegate','IsInherited','Source','CapturedAt') },
        [pscustomobject]@{ Name='Evidence'; Data=@($Assessment.Evidence); Columns=@('RunId','EmailAddress','UserPrincipalName','MailboxSizeGb','TargetLicense','AssignedLicenses','AdUserCount','OnPremMailboxCount','OnPremRemoteMailboxCount','ExoRecipientCount','ExoMailboxCount','GraphUserCount','PermissionCount','CollectedAt') },
        [pscustomobject]@{ Name='Live Sources'; Data=@($Assessment.LiveSources); Columns=@('Source','Required','Available','Status','Details','CheckedAt') },
        [pscustomobject]@{ Name='Check Coverage'; Data=@($Assessment.CheckCoverage); Columns=@('RunId','MailboxCount','EnabledMailboxCheckCount','ExpectedMailboxFindingCount','ActualMailboxFindingCount','MaterializedMailboxUnknownCount','EnabledGlobalCheckCount','ActualGlobalFindingCount','MaterializedGlobalUnknownCount','DuplicateMailboxCheckCount','DuplicateGlobalCheckCount','Status') },
        [pscustomobject]@{ Name='Action Plan'; Data=@($actionPlan); Columns=@('Scope','EmailAddress','Priority','CheckId','Result','RecommendedAction') },
        [pscustomobject]@{ Name='Check Options'; Data=@($Assessment.CheckOptions); Columns=@('CheckId','Category','Name','Mandatory','Enabled','Description') }
    )
    if ($RunFolder -and (Test-Path -LiteralPath $RunFolder -PathType Container)) {
        $knownCsvFiles = @('Summary.csv','Findings.csv','Global-Findings.csv','Permissions-Baseline.csv','Evidence.csv','Live-Sources.csv','Check-Coverage.csv','Check-Options.csv')
        $usedSheetNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($sheet in $sheets) { [void]$usedSheetNames.Add([string]$sheet.Name) }
        foreach ($csvFile in @(Get-ChildItem -LiteralPath $RunFolder -Filter '*.csv' -File | Where-Object Name -NotIn $knownCsvFiles | Sort-Object Name)) {
            $csvRows = @(Import-Csv -LiteralPath $csvFile.FullName)
            if ($csvRows.Count -gt 0) { $csvColumns = @($csvRows[0].PSObject.Properties.Name) }
            else {
                $headerLine = @(Get-Content -LiteralPath $csvFile.FullName -TotalCount 1)
                $delimiter = if ($headerLine.Count -eq 1) { Get-SemrDelimiter -Header $headerLine[0] } else { ',' }
                $csvColumns = if ($headerLine.Count -eq 1) { @($headerLine[0].Split($delimiter) | ForEach-Object { $_.Trim().Trim('"') }) } else { @('Value') }
            }
            $sheetName = ([IO.Path]::GetFileNameWithoutExtension($csvFile.Name) -replace '[\[\]:*?/\\]', ' ').Trim()
            if ([string]::IsNullOrWhiteSpace($sheetName)) { $sheetName = 'Additional CSV' }
            if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0,31) }
            $baseName = $sheetName; $suffix = 2
            while (-not $usedSheetNames.Add($sheetName)) {
                $suffixText = " $suffix"; $maximumBaseLength = 31 - $suffixText.Length
                $sheetName = $baseName.Substring(0,[math]::Min($baseName.Length,$maximumBaseLength)) + $suffixText
                $suffix++
            }
            $sheets += [pscustomobject]@{ Name=$sheetName; Data=$csvRows; Columns=$csvColumns }
        }
    }
    $fileStream = [IO.File]::Open($Path, [IO.FileMode]::Create, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $archive = [IO.Compression.ZipArchive]::new($fileStream, [IO.Compression.ZipArchiveMode]::Create, $false)
    try {
        $overrides = for ($index = 1; $index -le $sheets.Count; $index++) { '<Override PartName="/xl/worksheets/sheet'+$index+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' }
        Write-SemrZipEntry $archive '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+($overrides -join '')+'</Types>')
        Write-SemrZipEntry $archive '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        $sheetNodes = for ($index = 0; $index -lt $sheets.Count; $index++) { '<sheet name="'+(ConvertTo-SemrXmlText $sheets[$index].Name)+'" sheetId="'+($index+1)+'" r:id="rId'+($index+1)+'"/>' }
        Write-SemrZipEntry $archive 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>'+($sheetNodes -join '')+'</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        $relationships = for ($index = 0; $index -lt $sheets.Count; $index++) { '<Relationship Id="rId'+($index+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+($index+1)+'.xml"/>' }
        $relationships += '<Relationship Id="rId'+($sheets.Count+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        Write-SemrZipEntry $archive 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+($relationships -join '')+'</Relationships>')
        $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="10"/><name val="Segoe UI"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Segoe UI"/></font><font><b/><sz val="10"/><name val="Segoe UI"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0078D4"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDFF3E4"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF1CC"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFDE2E1"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EDF3"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FFD9E2EC"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
        Write-SemrZipEntry $archive 'xl/styles.xml' $styles
        for ($index = 0; $index -lt $sheets.Count; $index++) {
            Write-SemrZipEntry $archive ("xl/worksheets/sheet$($index+1).xml") (ConvertTo-SemrWorksheetXml -Data @($sheets[$index].Data) -Columns @($sheets[$index].Columns))
        }
    }
    finally { $archive.Dispose(); $fileStream.Dispose() }
    return $Path
}

function New-SemrHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Assessment,
        [Parameter(Mandatory)][string]$Path
    )
    $summary = @(
        $Assessment.Summary | Sort-Object `
            @{ Expression = {
                switch ([string]$_.Decision) {
                    'NO-GO' { 0 }
                    'UNKNOWN' { 1 }
                    'GO-WARNING' { 2 }
                    'GO' { 3 }
                    default { 4 }
                }
            } },
            @{ Expression = { [int]$_.BlockingCount }; Descending = $true },
            @{ Expression = { [int]$_.WarningCount }; Descending = $true },
            @{ Expression = { [string]$_.EmailAddress } }
    )
    $globalFindings = @($Assessment.GlobalFindings)
    $go = @($summary | Where-Object Decision -EQ 'GO').Count
    $warning = @($summary | Where-Object Decision -EQ 'GO-WARNING').Count
    $noGo = @($summary | Where-Object Decision -EQ 'NO-GO').Count
    $unknown = @($summary | Where-Object Decision -EQ 'UNKNOWN').Count
    $duration = if ($Assessment.StartedAt -and $Assessment.CompletedAt) { [math]::Round(([datetime]$Assessment.CompletedAt - [datetime]$Assessment.StartedAt).TotalSeconds,1) } else { $null }
    $endpoint = if ($Assessment.Hybrid -and $Assessment.Hybrid.EndpointName) { [string]$Assessment.Hybrid.EndpointName } else { 'Not available' }
    $entra = if ($Assessment.EntraConnect) { [string]$Assessment.EntraConnect.Message } else { 'Not available' }
    $summaryRows = foreach ($row in $summary) {
        $decisionClass = ([string]$row.Decision).ToLowerInvariant().Replace('-','')
        $sizeDisplay = if ($row.PSObject.Properties['MailboxSizeGb'] -and $null -ne $row.MailboxSizeGb -and [string]$row.MailboxSizeGb -ne '') { ([double]$row.MailboxSizeGb).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) + ' GB' } else { 'Not available' }
        $searchText = "$($row.EmailAddress) $($row.UserPrincipalName) $sizeDisplay $($row.TargetLicense) $($row.AssignedLicenses) $($row.Decision) $($row.BlockingCodes) $($row.RecommendedAction)"
        $actionItems = @(([string]$row.RecommendedAction -split '\s*\|\s*') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $actionHtml = if ($actionItems.Count -gt 0) {
            '<ul class="actions"><li>' + (($actionItems | ForEach-Object { ConvertTo-SemrHtmlText $_ }) -join '</li><li>') + '</li></ul>'
        }
        else {
            ''
        }
        '<tr data-search="'+(ConvertTo-SemrHtmlText $searchText)+'"><td class="mailbox-col">'+(ConvertTo-SemrHtmlText $row.EmailAddress)+'</td><td class="mailbox-col">'+(ConvertTo-SemrHtmlText $row.UserPrincipalName)+'</td><td class="number">'+(ConvertTo-SemrHtmlText $sizeDisplay)+'</td><td class="license-col">'+(ConvertTo-SemrHtmlText $row.TargetLicense)+'</td><td class="license-col">'+(ConvertTo-SemrHtmlText $row.AssignedLicenses)+'</td><td><span class="badge '+$decisionClass+'">'+(ConvertTo-SemrHtmlText $row.Decision)+'</span></td><td>'+$row.BlockingCount+'</td><td>'+$row.WarningCount+'</td><td>'+$row.UnknownCount+'</td><td>'+(ConvertTo-SemrHtmlText $row.DataCoverage)+'</td><td class="codes-col">'+(ConvertTo-SemrHtmlText $row.BlockingCodes)+'</td><td class="action-col">'+$actionHtml+'</td></tr>'
    }
    $tenantRows = foreach ($row in $globalFindings) {
        $resultClass = ([string]$row.Result).ToLowerInvariant().Replace('-','')
        '<tr><td>'+(ConvertTo-SemrHtmlText $row.CheckId)+'</td><td>'+(ConvertTo-SemrHtmlText $row.Category)+'</td><td><span class="badge '+$resultClass+'">'+(ConvertTo-SemrHtmlText $row.Result)+'</span></td><td>'+(ConvertTo-SemrHtmlText $row.Message)+'</td><td>'+(ConvertTo-SemrHtmlText $row.RecommendedAction)+'</td></tr>'
    }
    $sourceRows = foreach ($row in @($Assessment.LiveSources)) {
        $sourceClass = if ($row.Available) { 'pass' } else { 'fail' }
        '<tr><td>'+(ConvertTo-SemrHtmlText $row.Source)+'</td><td><span class="badge '+$sourceClass+'">'+(ConvertTo-SemrHtmlText $row.Status)+'</span></td><td>'+(ConvertTo-SemrHtmlText $row.Required)+'</td><td>'+(ConvertTo-SemrHtmlText $row.Details)+'</td><td>'+(ConvertTo-SemrHtmlText $row.CheckedAt)+'</td></tr>'
    }
    $findingRows = foreach ($row in @($Assessment.Findings | Where-Object { $_.Result -in @('FAIL','UNKNOWN','WARN') })) {
        '<tr><td>'+(ConvertTo-SemrHtmlText $row.EmailAddress)+'</td><td>'+(ConvertTo-SemrHtmlText $row.CheckId)+'</td><td>'+(ConvertTo-SemrHtmlText $row.Result)+'</td><td>'+(ConvertTo-SemrHtmlText $row.Message)+'</td><td>'+(ConvertTo-SemrHtmlText $row.RecommendedAction)+'</td></tr>'
    }
    $html = @"
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Smart Exchange Migration Readiness - $($Assessment.RunId)</title><style>
:root{--blue:#0078d4;--navy:#18324a;--muted:#5f6b7a;--line:#d9e2ec;--bg:#f4f8fb;--green:#146c43;--amber:#8a5a00;--red:#b42318}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:#1f2937;font:14px Segoe UI,Arial,sans-serif}.wrap{max-width:1500px;margin:auto;padding:24px}.hero,.panel{background:white;border:1px solid var(--line);border-radius:12px;padding:20px;margin-bottom:16px}.hero h1{margin:0 0 6px;font-size:28px}.sub{color:var(--muted)}.overview{display:flex;flex-wrap:wrap;gap:10px;margin-top:16px}.overview .card{flex:1 1 110px;min-width:105px;border:1px solid var(--line);border-radius:9px;padding:11px 12px;background:#f9fbfd}.overview .card strong{display:block;font-size:24px}.go{color:var(--green)}.gowarning,.warn,.warning{color:var(--amber)}.nogo,.fail{color:var(--red)}.unknown{color:#586477}.badge{display:inline-block;border-radius:999px;padding:3px 9px;font-weight:700;background:#e8edf3}.badge.go,.badge.pass{background:#dff3e4;color:var(--green)}.badge.gowarning,.badge.warn{background:#fff1cc;color:var(--amber)}.badge.nogo,.badge.fail{background:#fde2e1;color:var(--red)}h2{margin:0 0 14px}input{width:100%;max-width:620px;padding:10px;border:1px solid #b8c5d1;border-radius:7px;margin-bottom:12px}table{border-collapse:collapse;width:100%;font-size:13px;min-width:1280px}th{position:sticky;top:0;background:var(--blue);color:white;text-align:left;padding:9px;white-space:nowrap}td{border-bottom:1px solid #e5ebf1;padding:8px;vertical-align:top}.mailbox-table{min-width:1900px}.mailbox-col{min-width:245px}.license-col{min-width:125px;white-space:nowrap}.codes-col{min-width:220px}.action-col{min-width:420px;max-width:540px;line-height:1.5;white-space:normal}.actions{margin:0;padding-left:19px}.actions li{margin:0 0 8px}.actions li:last-child{margin-bottom:0}.number{text-align:right;white-space:nowrap}tbody tr:hover{background:#f4f9fd}.scroll{overflow:auto;max-height:620px}details{background:white;border:1px solid var(--line);border-radius:12px;margin-bottom:16px;padding:14px}summary{cursor:pointer;font-size:18px;font-weight:600}.note{padding:12px;background:#eef6fc;border-left:4px solid var(--blue);margin-top:12px}@media(max-width:900px){.overview .card{flex-basis:145px}}@media print{body{background:white}.wrap{max-width:none;padding:0}.scroll{max-height:none;overflow:visible}th{position:static}}
</style></head><body><main class="wrap"><section class="hero"><h1>Smart Exchange Migration Readiness</h1><div class="sub">Read-only assessment $($Assessment.RunId)</div><div class="overview"><div class="card"><b>Status</b><div>$(ConvertTo-SemrHtmlText $Assessment.AssessmentStatus)</div></div><div class="card"><b>Mode</b><div>Live strict</div></div><div class="card"><b>Phase</b><div>$(ConvertTo-SemrHtmlText $Assessment.AssessmentPhase)</div></div><div class="card"><b>Endpoint</b><div>$(ConvertTo-SemrHtmlText $endpoint)</div></div><div class="card"><b>Duration</b><div>$duration s</div></div><div class="card go"><b>GO</b><strong>$go</strong></div><div class="card gowarning"><b>GO-WARNING</b><strong>$warning</strong></div><div class="card nogo"><b>NO-GO</b><strong>$noGo</strong></div><div class="card unknown"><b>UNKNOWN</b><strong>$unknown</strong></div></div><div class="note"><b>Tenant synchronization:</b> $(ConvertTo-SemrHtmlText $entra)</div></section>
<section class="panel"><h2>Mailbox decisions</h2><input id="mailFilter" placeholder="Filter by mailbox, UPN, size, license, decision, code or action"><div class="scroll"><table id="mailTable" class="mailbox-table"><thead><tr><th class="mailbox-col">Mailbox</th><th class="mailbox-col">UPN</th><th>Size</th><th class="license-col">Target license</th><th class="license-col">Assigned licenses</th><th>Decision</th><th>Blocking</th><th>Warnings</th><th>Unknown</th><th>Coverage</th><th class="codes-col">Blocking codes</th><th class="action-col">Recommended action</th></tr></thead><tbody>$($summaryRows -join '')</tbody></table></div></section>
<details open><summary>Tenant checks</summary><div class="scroll"><table><thead><tr><th>Check</th><th>Category</th><th>Result</th><th>Message</th><th>Recommended action</th></tr></thead><tbody>$($tenantRows -join '')</tbody></table></div></details>
<details><summary>Blocking, warning and unknown details</summary><div class="scroll"><table><thead><tr><th>Mailbox</th><th>Check</th><th>Result</th><th>Message</th><th>Recommended action</th></tr></thead><tbody>$($findingRows -join '')</tbody></table></div></details>
<details open><summary>Live source status</summary><div class="scroll"><table><thead><tr><th>Source</th><th>Status</th><th>Required</th><th>Details</th><th>Checked at</th></tr></thead><tbody>$($sourceRows -join '')</tbody></table></div></details>
</main><script>document.getElementById('mailFilter').addEventListener('input',function(){var q=this.value.toLowerCase();document.querySelectorAll('#mailTable tbody tr').forEach(function(r){r.style.display=(r.getAttribute('data-search')||'').toLowerCase().includes(q)?'':'none';});});</script></body></html>
"@
    [IO.File]::WriteAllText($Path, $html, [Text.UTF8Encoding]::new($false))
    return $Path
}
function Export-SemrAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Assessment,
        [Parameter(Mandatory)][string]$OutputRoot,
        [scriptblock]$ProgressCallback
    )

    $runFolder = Join-Path $OutputRoot $Assessment.RunId
    New-Item -Path $runFolder -ItemType Directory -Force | Out-Null
    $summaryPath = Join-Path $runFolder 'Summary.csv'
    $findingsPath = Join-Path $runFolder 'Findings.csv'
    $globalFindingsPath = Join-Path $runFolder 'Global-Findings.csv'
    $permissionsPath = Join-Path $runFolder 'Permissions-Baseline.csv'
    $evidencePath = Join-Path $runFolder 'Evidence.csv'
    $liveSourcesPath = Join-Path $runFolder 'Live-Sources.csv'
    $checkCoveragePath = Join-Path $runFolder 'Check-Coverage.csv'
    $checkOptionsPath = Join-Path $runFolder 'Check-Options.csv'
    $excelPath = Join-Path $runFolder ("SmartM365-ExchangeMigrationReadiness-$($Assessment.RunId).xlsx")
    $htmlPath = Join-Path $runFolder ("SmartM365-ExchangeMigrationReadiness-$($Assessment.RunId).html")
    if ($ProgressCallback) { & $ProgressCallback 'CSV reports' 'Writing detailed CSV files' 1 3 }

    Export-SemrCsvFile -Data @($Assessment.Summary) -Path $summaryPath -Columns @('RunId','BatchName','AssessmentPhase','AssessmentStatus','EmailAddress','UserPrincipalName','MailboxSizeGb','TargetLicense','AssignedLicenses','Decision','BlockingCount','MailboxBlockingCount','GlobalBlockingCount','WarningCount','MailboxWarningCount','GlobalWarningCount','UnknownCount','MailboxUnknownCount','GlobalUnknownCount','DataCoverage','BlockingCodes','RecommendedAction','CheckedAt')
    Export-SemrCsvFile -Data @($Assessment.Findings) -Path $findingsPath -Columns @('RunId','EmailAddress','CheckId','Category','Severity','Result','IsBlocking','ObservedValue','ExpectedValue','EvidenceSource','SourceTimestamp','Message','RecommendedAction')
    Export-SemrCsvFile -Data @($Assessment.GlobalFindings) -Path $globalFindingsPath -Columns @('RunId','EmailAddress','CheckId','Category','Severity','Result','IsBlocking','ObservedValue','ExpectedValue','EvidenceSource','SourceTimestamp','Message','RecommendedAction')
    Export-SemrCsvFile -Data @($Assessment.PermissionsBaseline) -Path $permissionsPath -Columns @('RunId','EmailAddress','PermissionType','Delegate','IsInherited','Source','CapturedAt')
    Export-SemrCsvFile -Data @($Assessment.Evidence) -Path $evidencePath -Columns @('RunId','EmailAddress','UserPrincipalName','MailboxSizeGb','TargetLicense','AssignedLicenses','AdUserCount','OnPremMailboxCount','OnPremRemoteMailboxCount','ExoRecipientCount','ExoMailboxCount','GraphUserCount','PermissionCount','CollectedAt')
    Export-SemrCsvFile -Data @($Assessment.LiveSources) -Path $liveSourcesPath -Columns @('Source','Required','Available','Status','Details','CheckedAt')
    Export-SemrCsvFile -Data @($Assessment.CheckCoverage) -Path $checkCoveragePath -Columns @('RunId','MailboxCount','EnabledMailboxCheckCount','ExpectedMailboxFindingCount','ActualMailboxFindingCount','MaterializedMailboxUnknownCount','EnabledGlobalCheckCount','ActualGlobalFindingCount','MaterializedGlobalUnknownCount','DuplicateMailboxCheckCount','DuplicateGlobalCheckCount','Status')
    Export-SemrCsvFile -Data @($Assessment.CheckOptions) -Path $checkOptionsPath -Columns @('CheckId','Category','Name','Mandatory','Enabled','Description')

    if ($ProgressCallback) { & $ProgressCallback 'Excel report' 'Building the autonomous OpenXML workbook' 2 3 }
    [void](New-SemrExcelReport -Assessment $Assessment -Path $excelPath -RunFolder $runFolder)
    if ($ProgressCallback) { & $ProgressCallback 'HTML report' 'Building the self-contained interactive report' 3 3 }
    [void](New-SemrHtmlReport -Assessment $Assessment -Path $htmlPath)

    return [pscustomobject]@{
        RunFolder = $runFolder
        SummaryPath = $summaryPath
        FindingsPath = $findingsPath
        GlobalFindingsPath = $globalFindingsPath
        PermissionsPath = $permissionsPath
        EvidencePath = $evidencePath
        LiveSourcesPath = $liveSourcesPath
        CheckCoveragePath = $checkCoveragePath
        CheckOptionsPath = $checkOptionsPath
        ExcelPath = $excelPath
        HtmlPath = $htmlPath
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
    'Test-SemrExchange2016WorkerSerialization',
    'Connect-SemrOnPremisesExchange',
    'Connect-SemrExchangeOnline',
    'Connect-SemrMicrosoftGraph',
    'Get-SemrMicrosoftGraphModuleState',
    'Install-SemrMicrosoftGraphModule',
    'Disconnect-SemrSession',
    'Import-SemrBatchCsv',
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxGTTN6YsKP2DE
# GIIddDQCAZ6hZ9GS6BPYxIwSGvbtaKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMskh4YxmNopTH5lawcfve7dBPrG2W6z/M9OYhxSHHTvMA0GCSqG
# SIb3DQEBAQUABIIBgJWIKpbTF5+4jw3mizHFrrge0jY+ANTIs1KUiM+OHc4iGoie
# Dv83wP+28yYfF//W2324UCD8A1WhD/btG0VyuIBVGIENfK0YUWKVg52tflWHRLuD
# Ofn42K/GocNaE5tDQReBkpnYDQ+koXqJAH94XDo2jBRiI7fjLAYnessF+FDOEiPP
# Yx0nn7czYES+kBcran/xDnQf0giPH1b6uHg/sShKQYAspc9TR+gzIYqyNF/UEQip
# uYyA7u+4v/NIa5S8jjnE+lRskoJotvlgDIBoNPidQO7speTaQEk/mAHTS25/mHL/
# /grRQjvN0VkvjBlOQTE3xZXqr3lWjD8ltPS4QdASx9i/XKXh2V0yjzpHoOcbutQE
# DX572IJ1eTmswIFjb8AmBVHLnAQXDDRCTGPKr9govgpH02qNnOY5GNKXjileewgk
# IAt5DWkQHL4NWkd0v/Tktp41ix3EiwF4C3rTiMaFvt+YqQGEqqwtstS4wei2gYa+
# aiwWHMbD2ULIkbyncaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxMjM4
# NDhaMC8GCSqGSIb3DQEJBDEiBCAtSJf1mVgP7WFf9U02HmR7QO+KOf6m8k6dupzY
# Sz53ojANBgkqhkiG9w0BAQEFAASCAgCZWZFIqyNCXgGhAZydSgDOanQOCX/tJzqw
# 8pxatgFAn9ky0fB903FCuHDuL7/QI1w4lFfGz9d/a2zpbQDsUyNfcqDfTrE0z50C
# 61PyDC2qFkcHgkump6CXXQ7n1VGiZk1/F2yOm1jb/Di+9aU1IQvb1OXQTzOkp7xd
# fDQX0mP6qO0Z8VqKEYA9K2X49vLZ4dWdYHKWJnQi+ggc3I19JDxstNEz8LFQbrCj
# YUzTKEpN/K0sWNZY9BofaEORpYfHtWmvcaw1sxtixRHbrqJAvKvOUaQJ0VmsxL8z
# Ro31Qmp/Fw/v9L0qjp5SxTsgG7OyNHHRBti3IYXL1aAlR3HNzXxqKZ5Sa150v29h
# fLV3UHjYhZjyazr+Ul5myoUIWwDZNF11+mzN9u7HB51c79ssPRblEbDHJ3YmP5eu
# 4McLLysnDgF+iaRdBUinTNMaL9UQzVRNdYu8Fen/HFwkg/1JgeZ0ets6pJxuT9qp
# FkoLhTSXhhvzek41hmZE8BOXie1xIoU33eKM16pVtDUQhgPQFU9q0imhx7KYMcVw
# 6e7X0CKVIGHWRMijufrbE9hWiZcOATm6Tx80Y8rRMDFdX0RJopL2Pg5KKtyHRy3t
# 4U7GobSEB2XrF1ctu5KkDP87zHvao/YUEKcfZMY8xOFgQ5+TZ/41I4OUaIwBnnND
# QCLUJyKOmQ==
# SIG # End signature block
