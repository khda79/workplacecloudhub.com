<#
.SYNOPSIS
    Read-only SharePoint Server 2016/2019 site collection and web inventory.
.VERSION
    1.0.0
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
$merged['EnableSharePointUpload'] = $false
$script:EffectiveConfig = [pscustomobject]$merged
$global:SmartM365GlobalConfig = $script:EffectiveConfig
$global:EnableSharePointUpload = $false

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

function Assert-InventoryPath {
    param([string]$Path, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not [IO.Path]::IsPathRooted($Path) -or $Path -match '\{\{') {
        throw "$Name must be an absolute resolved path."
    }
    return [IO.Path]::GetFullPath($Path)
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
    param($Site, $ContentDatabase)
    $raw = Get-ObservedProperty $Site 'LockState'
    $rawText = [string]$raw
    $state = switch -Regex ($rawText) {
        '^(Unlock|Unlocked)$' { 'Unlock'; break }
        '^(NoAdditions|Content)$' { 'NoAdditions'; break }
        '^(ReadOnly|Readonly)$' { 'ReadOnly'; break }
        '^(NoAccess|Noaccess)$' { 'NoAccess'; break }
        default { '' }
    }
    $readLocked = Get-ObservedProperty $Site 'ReadLocked'
    $writeLocked = Get-ObservedProperty $Site 'WriteLocked'
    $isReadOnly = Get-ObservedProperty $Site 'ReadOnly'
    if (Test-MissingObservation $isReadOnly) { $isReadOnly = Get-ObservedProperty $Site 'IsReadOnly' }
    $issue = Get-ObservedProperty $Site 'LockIssue'
    $databaseReadOnly = Get-ObservedProperty $ContentDatabase 'IsReadOnly'
    # Boolean properties cannot distinguish Unlock from NoAdditions on all builds.
    if (-not $state -and -not (Test-MissingObservation $readLocked) -and [bool]$readLocked) { $state = 'NoAccess' }
    if (-not $state -and ((-not (Test-MissingObservation $isReadOnly) -and [bool]$isReadOnly) -or (-not (Test-MissingObservation $writeLocked) -and [bool]$writeLocked))) { $state = 'ReadOnly' }
    return [pscustomobject]@{
        LockState = $state
        IsReadOnly = if (Test-MissingObservation $isReadOnly) { '' } else { [string][bool]$isReadOnly }
        ReadLocked = if (Test-MissingObservation $readLocked) { '' } else { [string][bool]$readLocked }
        WriteLocked = if (Test-MissingObservation $writeLocked) { '' } else { [string][bool]$writeLocked }
        LockIssue = [string]$issue
        ContentDatabaseIsReadOnly = if (Test-MissingObservation $databaseReadOnly) { '' } else { [string][bool]$databaseReadOnly }
        LockStatus = if ($state) { 'Observed' } else { 'Unverified' }
    }
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

function Write-RunCsv {
    param([string]$Name, [object[]]$Rows, [string[]]$Columns, [string]$Folder)
    $path = Join-Path $Folder $Name
    Write-SmartM365CsvAtomically -Data @($Rows) -Path $path -Columns $Columns -Delimiter ';' -NoTenantKey
    return $path
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
    $mailParams = @{ From=$from; To=$to; Subject=$Subject; BodyHtml=$BodyHtml; MailPurpose='Report'; ErrorAction='Stop' }
    foreach ($name in @('SmtpServer','SendMailMode','Cc')) {
        $value = [string](Get-InventoryConfigValue $name '')
        if (-not [string]::IsNullOrWhiteSpace($value)) { $mailParams[$name] = $value }
    }
    $smtpPort = Get-InventoryConfigValue 'SmtpPort' $null
    if ($null -ne $smtpPort -and [string]$smtpPort -ne '') { $mailParams['SmtpPort'] = [int]$smtpPort }
    $null = SendEmailHtmlReport @mailParams
    WriteLog -Message "Daily content summary email sent to $to."
}

function Flush-RunRows {
    param([string]$Kind, [hashtable]$Buffers, [hashtable]$Paths, [int]$MinimumCount = 1)
    $buffer = $Buffers[$Kind]
    if ($buffer.Count -lt $MinimumCount) { return }
    @($buffer.ToArray()) | Export-Csv -LiteralPath $Paths[$Kind] -Append -NoTypeInformation -Encoding UTF8 -Delimiter ';' -ErrorAction Stop
    $buffer.Clear()
}

$coreManifest = Join-Path $root 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
Import-Module -Name $coreManifest -MinimumVersion '1.0.50' -ErrorAction Stop
$outputBase = if ($PSBoundParameters.ContainsKey('OutputRoot')) { Resolve-InventoryConfigTokens -Value $OutputRoot -Name 'OutputRoot' } else { Get-InventoryConfigValue -Name 'OutputRoot' }
$outputBase = Assert-InventoryPath -Path $outputBase -Name 'OutputRoot'
$latestRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue -Name 'LatestCsvFolderPath') -Name 'LatestCsvFolderPath'
$runBase = if ($MaxItems -gt 0) { Join-Path $outputBase 'TEST' } else { $outputBase }
$runFolder = Join-Path $runBase ((Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:RunId.Substring(0,8))
$script:Completed = $false
try {
    $null = InitializeScriptEnvironment -OutputPath $runFolder -LogFileName $script:ScriptName -CallerScriptPath $PSCommandPath
    WriteLog -Message 'SharePoint content inventory starting. Automatic upload is disabled.'
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
                    $lock = Resolve-SiteLockObservation -Site $site -ContentDatabase $database
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
                        if ($lock.LockStatus -ne 'Observed' -or $collection['StorageBytes'] -eq '' -or $collection['QuotaLimitBytes'] -eq '' -or $collection['LastContentModifiedUtc'] -eq '' -or $collection['OwnerLogin'] -eq '' -or $collection['RootWebTemplate'] -eq '' -or $collection['LanguageLcid'] -eq '' -or $lock.IsReadOnly -eq '' -or $lock.ReadLocked -eq '' -or $lock.WriteLocked -eq '' -or $lock.ContentDatabaseIsReadOnly -eq '') {
                            $collection['CollectionStatus'] = 'MetadataIncomplete'
                            $coverage['Status'] = 'MetadataIncomplete'
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
                            if (-not $webRow['WebId'] -or -not $webRow['Url'] -or -not $webRow['WebTemplate'] -or -not $webRow['LanguageLcid'] -or -not $webRow['HasUniqueRoleAssignments']) { $coverage['Status'] = 'MetadataIncomplete' }
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
                if ($databaseCoverageStatus -ne 'Collected') { throw $databaseCoverageStatus }
            }
        } catch {
            if ($globalTimedOut -and $_.Exception.Message -eq 'GlobalTimeout') { break }
            $failureCount++
            $coverage = New-InventoryRow $farmId
            $coverage['WebApplicationId'] = [string]$webApplication.Id
            $coverage['ContentDatabaseId'] = [string]$database.Id
            $coverage['SiteCollectionId'] = ''
            $coverage['Url'] = ''
            $coverage['Status'] = if ($_.Exception.Message -in @('DatabaseCountUnavailable','DatabaseCoverageMismatch')) { $_.Exception.Message } else { 'DatabaseEnumerationFailed' }
            $coverage['WebsCollected'] = ''
            $coverage['ErrorStage'] = 'ContentDatabase'
            $coverage['ErrorMessage'] = $_.Exception.Message
            $rows.CollectionCoverage.Add([pscustomobject]$coverage)
            Flush-RunRows -Kind CollectionCoverage -Buffers $rows -Paths $spoolPaths -MinimumCount 1
            WriteLog -Message ("Content database {0} enumeration failed: {1}" -f $database.Id, $_.Exception.Message) -Level ERROR
        }
        if ($globalTimedOut) { break }
    }
    if ($MaxItems -gt 0 -and $processed -ge $MaxItems) { $limited = $true }
    if ([datetime]::UtcNow -ge $script:GlobalDeadlineUtc) { $globalTimedOut = $true }
    foreach ($kind in $schemas.Keys) { Flush-RunRows -Kind $kind -Buffers $rows -Paths $spoolPaths -MinimumCount 1 }
    WriteLog -Message ("Collection coverage: processed={0}; failed={1}; limited={2}; globalTimeout={3}." -f $processed,$failureCount,$limited,$globalTimedOut)
    if ($MaxItems -gt 0) {
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
    if ($failureCount -gt 0 -or $globalTimedOut) {
        WriteLog -Message 'Content coverage is incomplete. DATA-LAST and source receipt were not updated.' -Level ERROR
        Complete-SmartM365ExecutionContext -Status Failed
        $script:Completed = $true
        exit 1
    }
    if ([bool](Get-InventoryConfigValue 'EnableWeeklyHistory' $true)) {
        $historyRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue 'WeeklyHistoryFolderPath') -Name 'WeeklyHistoryFolderPath'
        Add-SmartM365WeeklyHistory -SourceCsvPaths @($runPaths) -HistoryRootPath $historyRoot -RetentionWeeks ([int](Get-InventoryConfigValue 'WeeklyHistoryRetentionWeeks' 52)) -HistoryLabel 'SharePoint on-prem content' | Out-Null
    }
    foreach ($path in $runPaths) {
        Copy-SmartM365FileAtomically -SourcePath $path -DestinationPath (Join-Path $latestRoot (Split-Path $path -Leaf))
    }
    $receiptPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    if (-not $receiptPath -or $receiptPath -notlike '*.current.json.txt') { throw 'Content source receipt qualification failed.' }
    WriteLog -Message 'Content collection and current publication completed.'
    try {
        $body = "<h2>SharePoint on-prem content inventory</h2><p>Qualified run: $($script:RunId)</p><table><tr><th>Web applications</th><th>Content databases</th><th>Site collections</th></tr><tr><td>$($webApplications.Count)</td><td>$($databaseGroups.Count)</td><td>$processed</td></tr></table>"
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
    Complete-SmartM365ExecutionContext -Status Success
    $script:Completed = $true
} catch {
    try { WriteLog -Message ("Content inventory failed: {0}" -f $_.Exception.Message) -Level ERROR } catch { Write-Error $_ }
    if (-not $script:Completed) { try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $_ -FailureStage 'ContentInventory' } catch {} }
    throw
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBaI+PyZrAud3bP
# H5Xkg2bk3r4VUmnyGHIUB9AOSAQI46CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCcHHRsshtGnAGHo8e12ksL
# vlb4eremsDFlzsH/ji3StjANBgkqhkiG9w0BAQEFAASCAYAaDA7hTr1uDEj2IO2b
# n4SgZzbrWpmCgW598SyYiia9JDowQyVLIb0479dQXKG+ny8Y1R+HuCuzDAUmxLGB
# kuj0UBf3kKhC8GbgzvGt/8On/5JzSZ47jKySr481iuU8/uCcwaSNOStJ00BbU2XI
# 26FBriwOGCZbz9oX2EURT08J5VTMhlmD/fFnnR6zaeWIXbLptmIZs2AdNM2juyoy
# TgngoP5AcE+oC0vPdXxt/9ZNUfiFHbhp6FUPA4oxBCo9YnES71ieerSNk96WTuqd
# osITuiM+7LnJeUE9/OOmFqYMYhaOM4c1modLv2R8BjnApXiB46FVM8YlEG6NvLsg
# /PYI+ES0nTqiEPHloULlP7E/jZU2DOVc3vIeGprBFXzqhrfMYtYVgfGcQYbStnQ6
# KJZmYv+o0E9jFmuDJsIKRTIrEJzA+vwxMH75VzBJTvql5fjSQFK1wk5DeN6H4JO8
# zEXB8w1OMtCkqw1N48Rh66Crur6AbtYmHatogBnh2PLrRa4=
# SIG # End signature block
