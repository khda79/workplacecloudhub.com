<#
.SYNOPSIS
    Read-only SharePoint Server 2016/2019 site collection and web inventory.
.VERSION
    1.0.8
.REQUIREMENTS
    Windows PowerShell 5.1 x64 on a SharePoint farm server; SharePoint Shell and content read access.
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$OutputRoot,
    [switch]$ValidateOnly,
    [switch]$ForceSendDailySummary,
    [ValidateRange(0, [int]::MaxValue)][int]$MaxItems = 0,
    [ValidateRange(1, 10080)][int]$GlobalTimeoutMinutes = 720,
    [ValidateRange(1, 1440)][int]$CollectionTimeoutMinutes = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ScriptName = 'SmartM365-SharePoint-OnPrem-Content-Inventory'
$script:RunId = [guid]::NewGuid().ToString('N')
$script:CollectedAtUtc = [datetime]::UtcNow.ToString('o')
$script:StartedAtUtc = [datetime]::UtcNow
$script:GlobalDeadlineUtc = $script:StartedAtUtc.AddMinutes($GlobalTimeoutMinutes)

$root = $PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Config\SmartM365-TenantContext.ps1'))) {
    $parent = Split-Path $root -Parent
    if (-not $parent -or $parent -eq $root) { throw 'SmartM365 root was not found.' }
    $root = $parent
}
. (Join-Path $root 'Config\SmartM365-TenantContext.ps1')
$script:EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$localConfig = Read-SmartM365JsonConfig -Path (Join-Path $PSScriptRoot "$script:ScriptName.local.json") -Required
$merged = ConvertTo-SmartM365Hashtable -InputObject $script:EffectiveConfig
foreach ($key in $localConfig.Keys) {
    $value = $localConfig[$key]
    if ($value -is [string] -and ($value -eq '' -or $value -in @('__USE_GLOBAL__','USE_GLOBAL'))) { continue }
    $merged[$key] = $value
}
if ($null -eq $merged['EnableSharePointUpload']) { $merged['EnableSharePointUpload'] = $true }
$script:EffectiveConfig = [pscustomobject]$merged
$global:SmartM365GlobalConfig = $script:EffectiveConfig
$global:EnableSharePointUpload = [bool]$merged['EnableSharePointUpload'] -and -not $ValidateOnly -and $MaxItems -eq 0

function Get-InventoryConfigValue {
    param([string]$Name, $DefaultValue = $null)
    $property = $script:EffectiveConfig.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    $value = $property.Value
    if ($value -is [string] -and ($value -in @('', '__USE_GLOBAL__', 'USE_GLOBAL'))) { return $DefaultValue }
    if ($value -isnot [string]) { return $value }
    return Resolve-InventoryConfigTokens -Value $value -Name $Name
}

function Resolve-InventoryConfigTokens {
    param([string]$Value, [string]$Name)
    $resolved = $Value
    for ($pass = 0; $pass -lt 10; $pass++) {
        $tokens = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokens.Count -eq 0) { break }
        $previous = $resolved
        foreach ($token in $tokens) {
            $property = $script:EffectiveConfig.PSObject.Properties[$token.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { throw "Unresolved configuration value: $Name" }
            $replacement = [string]$property.Value
            if ([string]::IsNullOrWhiteSpace($replacement) -or $replacement -in @('__USE_GLOBAL__','USE_GLOBAL')) { throw "Unresolved configuration value: $Name" }
            $resolved = $resolved.Replace($token.Value, $replacement)
        }
        if ($resolved -eq $previous) { break }
    }
    if ($resolved -match '\{\{' -or $resolved -in @('__USE_GLOBAL__','USE_GLOBAL')) { throw "Unresolved configuration value: $Name" }
    return $resolved
}

$global:SharePointSiteHostname = [string](Get-InventoryConfigValue 'SharePointSiteHostname' '')
$global:SharePointSitePath = [string](Get-InventoryConfigValue 'SharePointSitePath' '')
$global:SharePointLibraryDisplayName = [string](Get-InventoryConfigValue 'SharePointLibraryDisplayName' 'Documents')
$global:SharePointTargetFolderPath = [string](Get-InventoryConfigValue 'SharePointTargetFolderPath' '')
$global:AppId = [string](Get-InventoryConfigValue 'AppId' '')
$global:TenantId = [string](Get-InventoryConfigValue 'TenantId' '')
$global:Thumbprint = [string](Get-InventoryConfigValue 'Thumbprint' (Get-InventoryConfigValue 'Thumb' ''))

function Assert-InventoryPath {
    param([string]$Path, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not [IO.Path]::IsPathRooted($Path) -or $Path -match '\{\{') {
        throw "$Name must be an absolute resolved path."
    }
    $resolved = [IO.Path]::GetFullPath($Path)
    if ($root -match '[\\/]LauncherCache[\\/]') {
        $cacheScriptRoot = [IO.Path]::GetFullPath($root).TrimEnd([char[]]@('\','/'))
        if ($resolved.Equals($cacheScriptRoot, [StringComparison]::OrdinalIgnoreCase) -or
            $resolved.StartsWith(($cacheScriptRoot + [IO.Path]::DirectorySeparatorChar), [StringComparison]::OrdinalIgnoreCase)) {
            throw "$Name resolves inside the disposable launcher cache. Configure a persistent tenant path outside LauncherCache."
        }
    }
    return $resolved
}

function Get-ConfiguredWebApplications {
    $all = @(Get-SPWebApplication -ErrorAction Stop)
    $included = @(Get-InventoryConfigValue -Name 'IncludedWebApplicationUrls' -DefaultValue @())
    $excluded = @(Get-InventoryConfigValue -Name 'ExcludedWebApplicationUrls' -DefaultValue @())
    foreach ($entry in @($included) + @($excluded)) {
        if ($entry -isnot [string] -or [string]::IsNullOrWhiteSpace($entry)) { throw 'Web application filters must contain non-empty URLs.' }
    }
    $selected = @($all | Where-Object {
        $url = ([string]$_.Url).TrimEnd('/')
        ($included.Count -eq 0 -or @($included | Where-Object { ([string]$_).TrimEnd('/') -ieq $url }).Count -gt 0) -and
        @($excluded | Where-Object { ([string]$_).TrimEnd('/') -ieq $url }).Count -eq 0
    })
    foreach ($url in $included) {
        if (@($all | Where-Object { ([string]$_.Url).TrimEnd('/') -ieq ([string]$url).TrimEnd('/') }).Count -eq 0) {
            throw "Included web application was not found: $url"
        }
    }
    if ($selected.Count -eq 0) { throw 'No content web application remains after filtering.' }
    return $selected
}

function Test-SharePointPrerequisites {
    if ($PSVersionTable.PSVersion.Major -ne 5 -or -not [Environment]::Is64BitProcess) {
        throw 'Windows PowerShell 5.1 x64 is required on a SharePoint farm server.'
    }
    if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
    }
    $farm = Get-SPFarm -ErrorAction Stop
    if ($null -eq $farm -or -not $farm.Id) { throw 'SharePoint farm access is unavailable.' }
    return $farm
}

function New-InventoryRow {
    param([string]$FarmId)
    return [ordered]@{ TenantKey=$global:SmartM365TenantKey; FarmId=$FarmId; RunId=$script:RunId; CollectedAtUtc=$script:CollectedAtUtc }
}

function Get-ObservedProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return '' }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return '' }
    try { if ($null -eq $property.Value) { return '' }; return $property.Value }
    catch { throw "Cannot read $Name`: $($_.Exception.Message)" }
}

function Test-MissingObservation {
    param($Value)
    return ($null -eq $Value -or ($Value -is [string] -and $Value.Length -eq 0))
}

function Convert-ToUtcText {
    param($Value)
    if ($null -eq $Value -or $Value -eq '') { return '' }
    return ([datetime]$Value).ToUniversalTime().ToString('o')
}

function Resolve-SiteLockObservation {
    param($Site, $ContentDatabase, [string]$FilteredLockState = '', [bool]$FilteredLockConflict = $false)
    $rawText = [string](Get-ObservedProperty $Site 'LockState')
    $directState = switch -Regex ($rawText) {
        '^(Unlock|Unlocked)$' { 'Unlock'; break }
        '^(NoAdditions|Content)$' { 'NoAdditions'; break }
        '^(ReadOnly|Readonly)$' { 'ReadOnly'; break }
        '^(NoAccess|Noaccess)$' { 'NoAccess'; break }
        default { '' }
    }
    $filteredState = switch -Regex ($FilteredLockState) {
        '^(Unlock|Unlocked)$' { 'Unlock'; break }
        '^(NoAdditions|Content)$' { 'NoAdditions'; break }
        '^(ReadOnly|Readonly)$' { 'ReadOnly'; break }
        '^(NoAccess|Noaccess)$' { 'NoAccess'; break }
        default { '' }
    }
    $conflict = $FilteredLockConflict -or ($directState -and $filteredState -and $directState -ne $filteredState)
    $state = if ($conflict) { '' } elseif ($filteredState) { $filteredState } else { $directState }
    $readLocked = Get-ObservedProperty $Site 'ReadLocked'
    $writeLocked = Get-ObservedProperty $Site 'WriteLocked'
    $isReadOnly = Get-ObservedProperty $Site 'ReadOnly'
    if (Test-MissingObservation $isReadOnly) { $isReadOnly = Get-ObservedProperty $Site 'IsReadOnly' }
    $issue = Get-ObservedProperty $Site 'LockIssue'
    $databaseReadOnly = Get-ObservedProperty $ContentDatabase 'IsReadOnly'
    # Boolean properties cannot safely distinguish all four SharePoint lock states.
    return [pscustomobject]@{
        LockState = $state
        IsReadOnly = if (Test-MissingObservation $isReadOnly) { '' } else { [string][bool]$isReadOnly }
        ReadLocked = if (Test-MissingObservation $readLocked) { '' } else { [string][bool]$readLocked }
        WriteLocked = if (Test-MissingObservation $writeLocked) { '' } else { [string][bool]$writeLocked }
        LockIssue = [string]$issue
        ContentDatabaseIsReadOnly = if (Test-MissingObservation $databaseReadOnly) { '' } else { [string][bool]$databaseReadOnly }
        LockStatus = if ($conflict) { 'Conflict' } elseif ($state) { 'Observed' } else { 'Unverified' }
    }
}

function Get-LockStateLookup {
    param($ContentDatabase)
    $filters = [ordered]@{
        Unlock = { $_.LockState -eq 'Unlock' }
        NoAdditions = { $_.LockState -eq 'NoAdditions' }
        ReadOnly = { $_.LockState -eq 'ReadOnly' }
        NoAccess = { $_.LockState -eq 'NoAccess' }
    }
    $states = @{}
    $conflicts = @{}
    $queryErrors = New-Object 'System.Collections.Generic.List[string]'
    foreach ($state in $filters.Keys) {
        Assert-Deadline ([datetime]::MaxValue)
        $query = @{ ContentDatabase=$ContentDatabase; Filter=$filters[$state]; Limit='All'; ErrorAction='Stop' }
        try {
            Get-SPSite @query | ForEach-Object {
                $candidate = $_
                try {
                    $id = [string](Get-ObservedProperty $candidate 'Id')
                    if (-not $id) { throw 'The filtered site collection ID is unavailable.' }
                    if (-not $conflicts.ContainsKey($id)) {
                        if ($states.ContainsKey($id) -and $states[$id] -ne $state) {
                            $states.Remove($id)
                            $conflicts[$id] = $true
                        } else { $states[$id] = $state }
                    }
                } finally { if ($null -ne $candidate) { $candidate.Dispose() } }
            }
        } catch { $queryErrors.Add(("{0}: {1}" -f $state, $_.Exception.Message)) }
    }
    return [pscustomobject]@{ States=$states; Conflicts=$conflicts; Errors=$queryErrors.ToArray() }
}

function Assert-Deadline {
    param([datetime]$CollectionDeadlineUtc)
    if ([datetime]::UtcNow -ge $script:GlobalDeadlineUtc) { throw 'GlobalTimeout' }
    if ([datetime]::UtcNow -ge $CollectionDeadlineUtc) { throw 'CollectionTimeout' }
}

function Get-CollectionFailureStatus {
    param([string]$Message, $ErrorRecord)
    if ($Message -eq 'CollectionTimeout') { return 'TimedOut' }
    if ($Message -eq 'GlobalTimeout') { return 'GlobalTimedOut' }
    if ($ErrorRecord -and $ErrorRecord.Exception -and ($ErrorRecord.Exception -is [UnauthorizedAccessException] -or $ErrorRecord.Exception.HResult -eq -2147024891)) { return 'AccessDenied' }
    if ($Message -match 'access|denied|unauthorized|permission|acc[eè]s|refus[eé]') { return 'AccessDenied' }
    return 'Failed'
}

function Get-DatabaseCoverageStatus {
    param($ExpectedCount, [long]$ObservedCount)
    if (Test-MissingObservation $ExpectedCount) { return 'DatabaseCountUnavailable' }
    $expected = [long]0
    if (-not [long]::TryParse([string]$ExpectedCount, [ref]$expected)) { return 'DatabaseCountUnavailable' }
    if ($expected -ne $ObservedCount) { return 'DatabaseCoverageMismatch' }
    return 'Collected'
}

function Get-ContentPublicationDecision {
    param([int]$FailureCount, [bool]$GlobalTimedOut, [int]$MaxItems)
    if ($MaxItems -gt 0) { return 'TestOnly' }
    if ($GlobalTimedOut) { return 'Blocked' }
    if ($FailureCount -gt 0) { return 'Partial' }
    return 'Complete'
}

function Write-RunCsv {
    param([string]$Name, [object[]]$Rows, [string[]]$Columns, [string]$Folder)
    $path = Join-Path $Folder $Name
    Write-SmartM365CsvAtomically -Data @($Rows) -Path $path -Columns $Columns -Delimiter ';' -NoTenantKey
    if (-not (Get-Variable -Name csvGeneratedPaths -Scope Global -ErrorAction SilentlyContinue)) {
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }
    [void]$global:csvGeneratedPaths.Add($path)
    return $path
}

function Ensure-GraphAuthenticationModule {
    $moduleName = 'Microsoft.Graph.Authentication'
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) { throw 'Install-Module is unavailable.' }
        $previousProtocol = [Net.ServicePointManager]::SecurityProtocol
        try {
            [Net.ServicePointManager]::SecurityProtocol = $previousProtocol -bor [Net.SecurityProtocolType]::Tls12
            WriteLog -Message 'Installing Microsoft.Graph.Authentication from PSGallery for CurrentUser.' -Level INFO
            Install-Module -Name $moduleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        } finally {
            [Net.ServicePointManager]::SecurityProtocol = $previousProtocol
        }
    }
    Import-Module -Name $moduleName -ErrorAction Stop
    foreach ($commandName in @('Connect-MgGraph','Get-MgContext','Invoke-MgGraphRequest')) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) { throw "Graph Authentication command is unavailable: $commandName" }
    }
    WriteLog -Message 'Microsoft.Graph.Authentication is ready for SharePoint upload and Graph mail.' -Level INFO
}

function Invoke-DailySummaryMail {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$MarkerPath, [Parameter(Mandatory)][scriptblock]$SendAction, [switch]$Force)
    $today = (Get-Date).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $markerParent = Split-Path -Path $MarkerPath -Parent
    if (-not (Test-Path -LiteralPath $markerParent -PathType Container)) {
        New-Item -Path $markerParent -ItemType Directory -Force | Out-Null
    }
    $lockStream = $null
    try {
        try { $lockStream = [IO.File]::Open("$MarkerPath.lock", [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch [IO.IOException] {
            WriteLog -Message 'Daily content summary email is being evaluated by another run; email skipped.' -Level INFO
            return $false
        }
        $lastSentDate = if (Test-Path -LiteralPath $MarkerPath -PathType Leaf) { [string](Get-Content -LiteralPath $MarkerPath -Raw -ErrorAction Stop) } else { '' }
        if (-not $Force -and $lastSentDate.Trim() -eq $today) {
            WriteLog -Message "Daily content summary email already sent for $today; email skipped." -Level INFO
            return $false
        }
        $null = & $SendAction
        [IO.File]::WriteAllText($MarkerPath, $today, [Text.UTF8Encoding]::new($false))
        return $true
    } finally {
        if ($null -ne $lockStream) { $lockStream.Dispose() }
    }
}

function Send-InventorySummaryMail {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][string]$BodyHtml)
    $from = [string](Get-InventoryConfigValue 'From' '')
    $to = [string](Get-InventoryConfigValue 'To' '')
    if ([string]::IsNullOrWhiteSpace($to)) { $to = [string](Get-InventoryConfigValue 'ErrorMailTo' '') }
    if ([string]::IsNullOrWhiteSpace($from) -or [string]::IsNullOrWhiteSpace($to)) { throw 'Daily summary email requires From and To or ErrorMailTo.' }
    $mailParams = @{ From=$from; To=$to; Subject=$Subject; BodyHtml=$BodyHtml; MailPurpose='Report'; SendMailMode='Graph'; ErrorAction='Stop' }
    foreach ($name in @('Cc')) {
        $value = [string](Get-InventoryConfigValue $name '')
        if (-not [string]::IsNullOrWhiteSpace($value)) { $mailParams[$name] = $value }
    }
    $null = SendEmailHtmlReport @mailParams
    WriteLog -Message "Daily content summary email sent to $to."
}

function Publish-QualifiedCsvUploads {
    param([string[]]$Paths)
    if (-not $global:EnableSharePointUpload) { return 0 }
    $failures = 0
    foreach ($path in @($Paths | Sort-Object -Unique)) {
        try {
            if (-not (Invoke-SmartM365SharePointCsvUpload -LocalFilePath $path)) { throw 'No upload receipt was returned.' }
        } catch {
            $failures++
            WriteLog -Message ("SharePoint CSV upload incomplete for {0}: {1}" -f $path, $_.Exception.Message) -Level WARNING
        }
    }
    return $failures
}

function Flush-RunRows {
    param([string]$Kind, [hashtable]$Buffers, [hashtable]$Paths, [int]$MinimumCount = 1)
    $buffer = $Buffers[$Kind]
    if ($buffer.Count -lt $MinimumCount) { return }
    @($buffer.ToArray()) | Export-Csv -LiteralPath $Paths[$Kind] -Append -NoTypeInformation -Encoding UTF8 -Delimiter ';' -ErrorAction Stop
    $buffer.Clear()
}

$coreManifest = Join-Path $root 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
Import-Module -Name $coreManifest -MinimumVersion '1.0.55' -ErrorAction Stop
$outputBase = if ($PSBoundParameters.ContainsKey('OutputRoot')) { Resolve-InventoryConfigTokens -Value $OutputRoot -Name 'OutputRoot' } else { Get-InventoryConfigValue -Name 'OutputRoot' }
$outputBase = Assert-InventoryPath -Path $outputBase -Name 'OutputRoot'
$latestRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue -Name 'LatestCsvFolderPath') -Name 'LatestCsvFolderPath'
$runBase = if ($MaxItems -gt 0) { Join-Path $outputBase 'TEST' } else { $outputBase }
$runFolder = Join-Path $runBase ((Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:RunId.Substring(0,8))
$script:Completed = $false
try {
    $null = InitializeScriptEnvironment -OutputPath $runFolder -LogFileName $script:ScriptName -CallerScriptPath $PSCommandPath
    WriteLog -Message ("SharePoint content inventory starting. SharePoint upload enabled for this run: {0}." -f $global:EnableSharePointUpload)
    if (-not $ValidateOnly -and $MaxItems -eq 0) {
        try { Ensure-GraphAuthenticationModule }
        catch { WriteLog -Message ("Graph Authentication bootstrap incomplete; local inventory will continue: {0}" -f $_.Exception.Message) -Level WARNING }
    }
    $farm = Test-SharePointPrerequisites
    $webApplications = @(Get-ConfiguredWebApplications)
    $databaseGroups = New-Object 'System.Collections.Generic.List[object]'
    foreach ($webApplication in $webApplications) {
        foreach ($database in @(Get-SPContentDatabase -WebApplication $webApplication -ErrorAction Stop)) {
            $databaseGroups.Add([pscustomobject]@{ WebApplication=$webApplication; ContentDatabase=$database })
        }
    }
    WriteLog -Message ("Farm {0}: {1} web applications and {2} content databases selected." -f $farm.Id, $webApplications.Count, $databaseGroups.Count)
    if ($ValidateOnly) {
        WriteLog -Message 'ValidateOnly completed. No site collection was traversed and no CSV or receipt was published.'
        Complete-SmartM365ExecutionContext -Status Success
        $script:Completed = $true
        return
    }
    if ($MaxItems -gt 0) {
        WriteLog -Message ("MaxItems={0}: test run; DATA-LAST, weekly history and source receipt will not be changed." -f $MaxItems) -Level WARNING
    }
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $latestRoot -ReadOnly:($MaxItems -gt 0) -ScopeParameters @{ IncludedWebApplicationUrls=@(Get-InventoryConfigValue 'IncludedWebApplicationUrls' @()); ExcludedWebApplicationUrls=@(Get-InventoryConfigValue 'ExcludedWebApplicationUrls' @()); MaxItems=$MaxItems }
    $farmId = [string]$farm.Id
    $rows = @{
        SiteCollections = New-Object 'System.Collections.Generic.List[object]'
        SiteAdministrators = New-Object 'System.Collections.Generic.List[object]'
        Webs = New-Object 'System.Collections.Generic.List[object]'
        CollectionCoverage = New-Object 'System.Collections.Generic.List[object]'
    }
    $schemas = [ordered]@{
        SiteCollections = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','ContentDatabaseId','SiteCollectionId','Url','OwnerLogin','SecondaryOwnerLogin','StorageBytes','QuotaLimitBytes','LastContentModifiedUtc','RootWebTemplate','LanguageLcid','LockState','IsReadOnly','ReadLocked','WriteLocked','LockIssue','ContentDatabaseIsReadOnly','LockStatus','CollectionStatus')
        SiteAdministrators = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','ContentDatabaseId','SiteCollectionId','AdministratorLogin','IsPrimary','IsSecondary')
        Webs = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','ContentDatabaseId','SiteCollectionId','WebId','Url','Title','WebTemplate','LanguageLcid','ListCount','LibraryCount','HasUniqueRoleAssignments')
        CollectionCoverage = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','ContentDatabaseId','SiteCollectionId','Url','Status','WebsCollected','ErrorStage','ErrorMessage')
    }
    $runPaths = New-Object 'System.Collections.Generic.List[string]'
    $spoolPaths = @{}
    foreach ($kind in $schemas.Keys) {
        $path = Write-RunCsv -Name "SharePoint_OnPrem_$kind.csv" -Rows @() -Columns $schemas[$kind] -Folder $runFolder
        $runPaths.Add($path)
        $spoolPaths[$kind] = $path
    }
    $failureCount = 0
    $processed = 0
    $lockObservedCount = 0
    $lockUnverifiedCount = 0
    $limited = $MaxItems -gt 0
    $globalTimedOut = $false
    foreach ($group in $databaseGroups) {
        if ([datetime]::UtcNow -ge $script:GlobalDeadlineUtc) { $globalTimedOut = $true; break }
        if ($MaxItems -gt 0 -and $processed -ge $MaxItems) { $limited = $true; break }
        $database = $group.ContentDatabase
        $webApplication = $group.WebApplication
        $processedBeforeDatabase = $processed
        $remaining = if ($MaxItems -gt 0) { $MaxItems - $processed } else { 0 }
        try {
            $databaseSiteCount = Get-ObservedProperty $database 'CurrentSiteCount'
            $databaseLockLookup = Get-LockStateLookup -ContentDatabase $database
            foreach ($queryError in $databaseLockLookup.Errors) {
                WriteLog -Message ("Content database {0} lock-state filter failed: {1}" -f $database.Id, $queryError) -Level WARNING
            }
            if ($databaseLockLookup.Conflicts.Count -gt 0) {
                WriteLog -Message ("Content database {0} returned {1} conflicting lock-state filter result(s)." -f $database.Id, $databaseLockLookup.Conflicts.Count) -Level WARNING
            }
            $processSite = {
                $site = $_
                $processed++
                $collectionDeadline = [datetime]::UtcNow.AddMinutes($CollectionTimeoutMinutes)
                $siteId = ''
                $siteUrl = ''
                $coverage = New-InventoryRow $farmId
                $coverage['WebApplicationId'] = [string]$webApplication.Id
                $coverage['ContentDatabaseId'] = [string]$database.Id
                $coverage['SiteCollectionId'] = ''
                $coverage['Url'] = ''
                $coverage['Status'] = 'Collected'
                $coverage['WebsCollected'] = 0
                $coverage['ErrorStage'] = ''
                $coverage['ErrorMessage'] = ''
                $lock = $null
                $siteRow = $null
                try {
                    Assert-Deadline $collectionDeadline
                    $siteId = [string](Get-ObservedProperty $site 'Id')
                    $siteUrl = [string](Get-ObservedProperty $site 'Url')
                    $coverage['SiteCollectionId'] = $siteId
                    $coverage['Url'] = $siteUrl
                    if ([string]::IsNullOrWhiteSpace($siteId) -or [string]::IsNullOrWhiteSpace($siteUrl)) { throw 'Site collection identity is unavailable.' }
                    $filteredLockState = ''
                    if ($databaseLockLookup.States.ContainsKey($siteId)) { $filteredLockState = [string]$databaseLockLookup.States[$siteId] }
                    $lock = Resolve-SiteLockObservation -Site $site -ContentDatabase $database -FilteredLockState $filteredLockState -FilteredLockConflict $databaseLockLookup.Conflicts.ContainsKey($siteId)
                    if ($lock.LockStatus -eq 'Observed') { $lockObservedCount++ } else { $lockUnverifiedCount++ }
                    $rootWeb = $site.RootWeb
                    try {
                        $collection = New-InventoryRow $farmId
                        $collection['WebApplicationId'] = [string]$webApplication.Id
                        $collection['ContentDatabaseId'] = [string]$database.Id
                        $collection['SiteCollectionId'] = $siteId
                        $collection['Url'] = $siteUrl
                        $collection['OwnerLogin'] = [string](Get-ObservedProperty (Get-ObservedProperty $site 'Owner') 'LoginName')
                        $collection['SecondaryOwnerLogin'] = [string](Get-ObservedProperty (Get-ObservedProperty $site 'SecondaryContact') 'LoginName')
                        $collection['StorageBytes'] = [string](Get-ObservedProperty (Get-ObservedProperty $site 'Usage') 'Storage')
                        $collection['QuotaLimitBytes'] = [string](Get-ObservedProperty (Get-ObservedProperty $site 'Quota') 'StorageMaximumLevel')
                        $collection['LastContentModifiedUtc'] = Convert-ToUtcText (Get-ObservedProperty $site 'LastContentModifiedDate')
                        $collection['RootWebTemplate'] = [string](Get-ObservedProperty $rootWeb 'WebTemplate')
                        $collection['LanguageLcid'] = [string](Get-ObservedProperty $rootWeb 'Language')
                        foreach ($name in @('LockState','IsReadOnly','ReadLocked','WriteLocked','LockIssue','ContentDatabaseIsReadOnly','LockStatus')) { $collection[$name] = $lock.$name }
                        $collection['CollectionStatus'] = 'Collected'
                        $missingFields = New-Object 'System.Collections.Generic.List[string]'
                        if ($lock.LockStatus -ne 'Observed') { $missingFields.Add('LockState') }
                        foreach ($field in @('StorageBytes','QuotaLimitBytes','LastContentModifiedUtc','OwnerLogin','RootWebTemplate','LanguageLcid','IsReadOnly','ReadLocked','WriteLocked','ContentDatabaseIsReadOnly')) {
                            if (Test-MissingObservation $collection[$field]) { $missingFields.Add($field) }
                        }
                        if ($missingFields.Count -gt 0) {
                            $collection['CollectionStatus'] = 'MetadataIncomplete'
                            $coverage['Status'] = 'MetadataIncomplete'
                            $coverage['ErrorStage'] = 'SiteCollectionMetadata'
                            $coverage['ErrorMessage'] = 'Missing or unverified observations: ' + ($missingFields -join ', ')
                        }
                        $siteRow = [pscustomobject]$collection
                        $rows.SiteCollections.Add($siteRow)
                        $administrators = Get-ObservedProperty $rootWeb 'SiteAdministrators'
                        if (Test-MissingObservation $administrators) { throw 'Site administrator collection is unavailable.' }
                        $administratorCount = 0
                        foreach ($administrator in $administrators) {
                            Assert-Deadline $collectionDeadline
                            $adminRow = New-InventoryRow $farmId
                            $adminRow['WebApplicationId'] = [string]$webApplication.Id
                            $adminRow['ContentDatabaseId'] = [string]$database.Id
                            $adminRow['SiteCollectionId'] = $siteId
                            $adminRow['AdministratorLogin'] = [string](Get-ObservedProperty $administrator 'LoginName')
                            if (-not $adminRow['AdministratorLogin']) { throw 'Site administrator login is unavailable.' }
                            $adminRow['IsPrimary'] = if ($collection['OwnerLogin']) { [string]($adminRow['AdministratorLogin'] -ieq $collection['OwnerLogin']) } else { '' }
                            $adminRow['IsSecondary'] = if ($collection['SecondaryOwnerLogin']) { [string]($adminRow['AdministratorLogin'] -ieq $collection['SecondaryOwnerLogin']) } else { '' }
                            $rows.SiteAdministrators.Add([pscustomobject]$adminRow)
                            $administratorCount++
                            Flush-RunRows -Kind SiteAdministrators -Buffers $rows -Paths $spoolPaths -MinimumCount 500
                        }
                        if ($administratorCount -eq 0) { throw 'No site administrator was returned.' }
                        Assert-Deadline $collectionDeadline
                    } finally { if ($null -ne $rootWeb) { $rootWeb.Dispose() } }
                    foreach ($web in $site.AllWebs) {
                        try {
                            Assert-Deadline $collectionDeadline
                            $listCount = 0
                            $libraryCount = 0
                            foreach ($list in $web.Lists) {
                                $listCount++
                                if ([string](Get-ObservedProperty $list 'BaseType') -eq 'DocumentLibrary') { $libraryCount++ }
                            }
                            Assert-Deadline $collectionDeadline
                            $webRow = New-InventoryRow $farmId
                            $webRow['WebApplicationId'] = [string]$webApplication.Id
                            $webRow['ContentDatabaseId'] = [string]$database.Id
                            $webRow['SiteCollectionId'] = $siteId
                            $webRow['WebId'] = [string](Get-ObservedProperty $web 'Id')
                            $webRow['Url'] = [string](Get-ObservedProperty $web 'Url')
                            $webRow['Title'] = [string](Get-ObservedProperty $web 'Title')
                            $webRow['WebTemplate'] = [string](Get-ObservedProperty $web 'WebTemplate')
                            $webRow['LanguageLcid'] = [string](Get-ObservedProperty $web 'Language')
                            $webRow['ListCount'] = $listCount
                            $webRow['LibraryCount'] = $libraryCount
                            $webRow['HasUniqueRoleAssignments'] = [string](Get-ObservedProperty $web 'HasUniqueRoleAssignments')
                            if (-not $webRow['WebId'] -or -not $webRow['Url'] -or -not $webRow['WebTemplate'] -or -not $webRow['LanguageLcid'] -or -not $webRow['HasUniqueRoleAssignments']) {
                                $coverage['Status'] = 'MetadataIncomplete'
                                if (-not $coverage['ErrorStage']) { $coverage['ErrorStage'] = 'WebMetadata'; $coverage['ErrorMessage'] = 'A web has missing identity, template, language or permission observations.' }
                            }
                            $rows.Webs.Add([pscustomobject]$webRow)
                            Flush-RunRows -Kind Webs -Buffers $rows -Paths $spoolPaths -MinimumCount 500
                            $coverage['WebsCollected'] = [int]$coverage['WebsCollected'] + 1
                        } finally { if ($null -ne $web) { $web.Dispose() } }
                    }
                    Assert-Deadline $collectionDeadline
                    if ($coverage['Status'] -ne 'Collected') {
                        $siteRow.CollectionStatus = $coverage['Status']
                        $failureCount++
                    }
                } catch {
                    $message = $_.Exception.Message
                    $coverage['Status'] = Get-CollectionFailureStatus -Message $message -ErrorRecord $_
                    $coverage['ErrorStage'] = 'SiteCollection'
                    $coverage['ErrorMessage'] = $message
                    if ($null -ne $siteRow) {
                        $siteRow.CollectionStatus = $coverage['Status']
                    } else {
                        $fallback = New-InventoryRow $farmId
                        $fallback['WebApplicationId'] = [string]$webApplication.Id
                        $fallback['ContentDatabaseId'] = [string]$database.Id
                        $fallback['SiteCollectionId'] = $siteId
                        $fallback['Url'] = $siteUrl
                        foreach ($name in @('OwnerLogin','SecondaryOwnerLogin','StorageBytes','QuotaLimitBytes','LastContentModifiedUtc','RootWebTemplate','LanguageLcid','LockState','IsReadOnly','ReadLocked','WriteLocked','LockIssue','ContentDatabaseIsReadOnly','LockStatus')) { $fallback[$name] = '' }
                        if ($null -ne $lock) { foreach ($name in @('LockState','IsReadOnly','ReadLocked','WriteLocked','LockIssue','ContentDatabaseIsReadOnly','LockStatus')) { $fallback[$name] = $lock.$name } }
                        $fallback['CollectionStatus'] = $coverage['Status']
                        $rows.SiteCollections.Add([pscustomobject]$fallback)
                    }
                    $failureCount++
                    WriteLog -Message ("Site collection {0} failed: {1}" -f $siteUrl, $message) -Level ERROR
                    if ($message -eq 'GlobalTimeout') { $globalTimedOut = $true }
                } finally {
                    $rows.CollectionCoverage.Add([pscustomobject]$coverage)
                    Flush-RunRows -Kind SiteCollections -Buffers $rows -Paths $spoolPaths -MinimumCount 1
                    Flush-RunRows -Kind SiteAdministrators -Buffers $rows -Paths $spoolPaths -MinimumCount 1
                    Flush-RunRows -Kind Webs -Buffers $rows -Paths $spoolPaths -MinimumCount 1
                    Flush-RunRows -Kind CollectionCoverage -Buffers $rows -Paths $spoolPaths -MinimumCount 1
                    if ($null -ne $site) { $site.Dispose() }
                }
                if ($globalTimedOut) { throw 'GlobalTimeout' }
            }
            if ($remaining -gt 0) {
                Get-SPSite -ContentDatabase $database -Limit All -ErrorAction Stop | Select-Object -First $remaining | ForEach-Object -Process $processSite
            } else {
                Get-SPSite -ContentDatabase $database -Limit All -ErrorAction Stop | ForEach-Object -Process $processSite
            }
            if ($MaxItems -eq 0) {
                $databaseCoverageStatus = Get-DatabaseCoverageStatus -ExpectedCount $databaseSiteCount -ObservedCount ($processed - $processedBeforeDatabase)
                if ($databaseCoverageStatus -ne 'Collected') { throw ("{0}: expected={1}; observed={2}" -f $databaseCoverageStatus,$databaseSiteCount,($processed - $processedBeforeDatabase)) }
            }
        } catch {
            if ($_.Exception.Message -eq 'GlobalTimeout') { $globalTimedOut = $true; break }
            $failureCount++
            $coverage = New-InventoryRow $farmId
            $coverage['WebApplicationId'] = [string]$webApplication.Id
            $coverage['ContentDatabaseId'] = [string]$database.Id
            $coverage['SiteCollectionId'] = ''
            $coverage['Url'] = ''
            $coverage['Status'] = if ($_.Exception.Message -match '^(DatabaseCountUnavailable|DatabaseCoverageMismatch):') { $Matches[1] } else { 'DatabaseEnumerationFailed' }
            $coverage['WebsCollected'] = ''
            $coverage['ErrorStage'] = 'ContentDatabase'
            $coverage['ErrorMessage'] = $_.Exception.Message
            $rows.CollectionCoverage.Add([pscustomobject]$coverage)
            Flush-RunRows -Kind CollectionCoverage -Buffers $rows -Paths $spoolPaths -MinimumCount 1
            if ($coverage['Status'] -in @('DatabaseCountUnavailable','DatabaseCoverageMismatch')) {
                WriteLog -Message ("Content database {0} coverage incomplete: {1}" -f $database.Id, $_.Exception.Message) -Level WARNING
            } else {
                WriteLog -Message ("Content database {0} enumeration failed: {1}" -f $database.Id, $_.Exception.Message) -Level ERROR
            }
        }
        if ($globalTimedOut) { break }
    }
    if ($MaxItems -gt 0 -and $processed -ge $MaxItems) { $limited = $true }
    if ([datetime]::UtcNow -ge $script:GlobalDeadlineUtc) { $globalTimedOut = $true }
    foreach ($kind in $schemas.Keys) { Flush-RunRows -Kind $kind -Buffers $rows -Paths $spoolPaths -MinimumCount 1 }
    WriteLog -Message ("Lock observation: observed={0}; unverifiedOrConflicting={1}." -f $lockObservedCount,$lockUnverifiedCount)
    $coverageLevel = if ($globalTimedOut) { 'ERROR' } elseif ($failureCount -gt 0) { 'WARNING' } else { 'INFO' }
    WriteLog -Message ("Collection coverage: processed={0}; failed={1}; limited={2}; globalTimeout={3}." -f $processed,$failureCount,$limited,$globalTimedOut) -Level $coverageLevel
    $publicationDecision = Get-ContentPublicationDecision -FailureCount $failureCount -GlobalTimedOut $globalTimedOut -MaxItems $MaxItems
    if ($publicationDecision -eq 'TestOnly') {
        if ($failureCount -gt 0 -or $globalTimedOut) {
            WriteLog -Message 'Limited content run has collection or coverage failures.' -Level ERROR
            Complete-SmartM365ExecutionContext -Status Failed
            $script:Completed = $true
            exit 1
        }
        Complete-SmartM365ExecutionContext -Status CompletedWithWarnings
        $script:Completed = $true
        exit 3
    }
    if ($publicationDecision -eq 'Blocked') {
        WriteLog -Message 'Content collection timed out globally. DATA-LAST and source receipt were not updated.' -Level ERROR
        Complete-SmartM365ExecutionContext -Status Failed
        $script:Completed = $true
        exit 1
    }
    $partialInventory = $publicationDecision -eq 'Partial'
    if ($partialInventory) {
        WriteLog -Message ("Partial content inventory: {0} coverage or metadata issue(s) are recorded in CollectionCoverage.csv. Current CSVs will be published with a partial source receipt; weekly history will be skipped." -f $failureCount) -Level WARNING
    }
    if (-not $partialInventory -and [bool](Get-InventoryConfigValue 'EnableWeeklyHistory' $true)) {
        $historyRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue 'WeeklyHistoryFolderPath') -Name 'WeeklyHistoryFolderPath'
        Add-SmartM365WeeklyHistory -SourceCsvPaths @($runPaths) -HistoryRootPath $historyRoot -RetentionWeeks ([int](Get-InventoryConfigValue 'WeeklyHistoryRetentionWeeks' 52)) -HistoryLabel 'SharePoint on-prem content' | Out-Null
    }
    $publishedCsvPaths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($path in $runPaths) {
        $latestPath = Join-Path $latestRoot (Split-Path $path -Leaf)
        Copy-SmartM365FileAtomically -SourcePath $path -DestinationPath $latestPath
        [void]$global:csvGeneratedPaths.Add($latestPath)
        $publishedCsvPaths.Add($path)
        $publishedCsvPaths.Add($latestPath)
    }
    $receiptPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    if (-not $receiptPath -or $receiptPath -notlike '*.current.json.txt') { throw 'Content source receipt qualification failed.' }
    WriteLog -Message $(if ($partialInventory) { 'Partial content collection and current publication completed.' } else { 'Content collection and current publication completed.' })
    $uploadFailures = Publish-QualifiedCsvUploads -Paths $publishedCsvPaths.ToArray()
    try {
        $runStatus = if ($partialInventory) { 'Partial' } else { 'Qualified' }
        $body = "<h2>SharePoint on-prem content inventory</h2><p>$runStatus run: $($script:RunId)</p><p>Coverage or metadata issues: $failureCount. Details: SharePoint_OnPrem_CollectionCoverage.csv.</p><p>SharePoint upload failures: $uploadFailures</p><table><tr><th>Web applications</th><th>Content databases</th><th>Site collections</th></tr><tr><td>$($webApplications.Count)</td><td>$($databaseGroups.Count)</td><td>$processed</td></tr></table>"
        $mailMarkerPath = Join-Path $outputBase 'SmartM365-SharePoint-OnPrem-Content-DailySummary.sent'
        $null = Invoke-DailySummaryMail -MarkerPath $mailMarkerPath -Force:$ForceSendDailySummary -SendAction {
            Send-InventorySummaryMail -Subject 'SharePoint on-prem content inventory' -BodyHtml $body
        }
    } catch {
        WriteLog -Message ("Daily content summary email failed: {0}" -f $_.Exception.Message) -Level ERROR
        Complete-SmartM365ExecutionContext -Status CompletedWithWarnings
        $script:Completed = $true
        exit 3
    }
    if ($partialInventory -or $uploadFailures -gt 0) {
        Complete-SmartM365ExecutionContext -Status CompletedWithWarnings
        $script:Completed = $true
        exit 3
    }
    Complete-SmartM365ExecutionContext -Status Success
    $script:Completed = $true
} catch {
    try { WriteLog -Message ("Content inventory failed: {0}" -f $_.Exception.Message) -Level ERROR } catch { Write-Error $_ }
    if (-not $script:Completed) { try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $_ -FailureStage 'ContentInventory' } catch {} }
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBM/yUc/Uox8mpJ
# pfgq6PvG1lIYoWvFbBrbYFbQ3z2Rz6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIJOfJVxQNsIfHZEyUe9zy078/ofqEMCj97imfJZGrAh4MA0GCSqG
# SIb3DQEBAQUABIIBgE3JgEkNnQhMbzVFXjRKKp5/d8whk7BtWluR2nWbxxp/kX38
# jF5BLkoYQ2T9HoU9umUZpizYx/oQUdp6k/fTvgAOn8qZVXgZ8Yk8y5Ulrf72PbIv
# OlL/HEFDs8OKzkw0Nj+SD9PP5Mhpj9MmBrmfk1nGG0EqVpDp0B16LMvs1+4OAl92
# CX7Pd5U+CeaX1LIEs+6C5WgW+MS4IUrN759JTbfjc5ov7Bn0GCrdwDSCopAO8bju
# K1+4PILdKDQcqOiyZo7cLiYyx0Vx4pcPfIj3sWf7G+MkR21jvowwzqxnDS8go7fO
# En0JKThbQcXjmKNP17jm/oJghAZ2weYBfr1g3hnb8EVgltiUh69AqNGe0sLnsjiN
# QLtLCL3HYdqxHtMyU0gvtG8dGigK4LzZ+AAt/CPxFOcnVcJ8K/3PsycAfXw7KDPe
# ZOOifOKyJfQx6FSxMEiQRR7lvGPlVH+fEeZiRWrXdM0vJq/P+MoqWviGYlV4Jh+q
# DDcaSApusUqIZmWh9KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDkyMjA0
# MTRaMC8GCSqGSIb3DQEJBDEiBCC3AbPRN2QWWLGGovIyEmx7ukAQHQ3Ex8f0E4Q6
# uxa5HTANBgkqhkiG9w0BAQEFAASCAgBbkymSEE+oJ1yOeG2Bn3KqTNjhK/rhimFh
# wCha54uMyctL861wdiBWFptcXP93fS9UWhmMMBiAvYB+HjHpFu2o29lMzCR0QqVO
# ckV+1FtC7Agfxh2Y4C7sBFCQcVk1yerIrV9aTl0NReXnN73/wcn0xN0UDdp5yq9F
# IgKobC60c0XC5Q5QZs2M7LOMJLzK/5kFWI6eWSyt4TneRzfEOghN9XKoI95IO3N0
# e2tvAnpd99HuyDOiKHnz9PXFWW4pQbKJThBSXBHLu5YxtyOu7o3CJv4o+2M/lRe/
# 4eWNAyC7lGjPMh8go5mSRyl7E+T+v+5Wr+TjAwAmGla9CKiDyd6tQdTp+plEvpVn
# tfLkXkrwCznJvgAQLgQZjRwMZV9Wnuh67lpQuKyDjweQenIZ2S08z0dnI0EuhEL3
# ygg3IUde3fBTrjxknRv3yZXORYC+bzYb/N4wrYq7EiBHZLx1vOvlRRQfr9khNh6S
# yw1rdFSo/h0Eh0V7fNVWAvRNmGCOUYxSXzgizvhEmyQh010GrqxIRDGatnLgLTgI
# DiGUwEe07FOhI4W8pgLW5hgEF5u239nI6IVIYcOlf7MtBocjcVDBEYEF/uKZ+7ik
# oJb4buXqP4t7jFBb81mdehqP1qPFL5H1MCm+fah199icW1zGpxUhgkXuvF7WkLrF
# qwubCxUgwA==
# SIG # End signature block
