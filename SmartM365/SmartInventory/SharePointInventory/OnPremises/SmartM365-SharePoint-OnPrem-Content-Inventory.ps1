<#
.SYNOPSIS
    Read-only SharePoint Server 2016/2019 site collection and web inventory.
.VERSION
    1.0.18
.REQUIREMENTS
    Windows PowerShell 5.1 x64 on a SharePoint farm server; SharePoint Shell and content read access.
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$OutputRoot,
    [switch]$ValidateOnly,
    [Alias('ForceMail')][switch]$ForceSendDailySummary,
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
$script:BuildEditionMapPath = Join-Path $PSScriptRoot 'SharePoint-OnPrem-BuildEditions.psd1'
$script:GlobalDeadlineUtc = $script:StartedAtUtc.AddMinutes($GlobalTimeoutMinutes)

$root = $PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Config\SmartM365-TenantContext.ps1'))) {
    $parent = Split-Path $root -Parent
    if (-not $parent -or $parent -eq $root) { throw 'SmartM365 root was not found.' }
    $root = $parent
}
. (Join-Path $root 'Config\SmartM365-TenantContext.ps1')
$script:EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$localConfig = Read-SmartM365ScriptConfiguration -Path (Join-Path $PSScriptRoot "$script:ScriptName.local.json") -AsHashtable
$merged = ConvertTo-SmartM365Hashtable -InputObject (Merge-SmartM365Configuration -EffectiveConfig $script:EffectiveConfig -LocalConfig $localConfig -Mode Evidence)
if ($null -eq $merged['EnableSharePointUpload']) { $merged['EnableSharePointUpload'] = $true }
$script:EffectiveConfig = [pscustomobject]$merged
$global:SmartM365GlobalConfig = $script:EffectiveConfig
$global:EnableSharePointUpload = [bool]$merged['EnableSharePointUpload'] -and -not $ValidateOnly -and $MaxItems -eq 0

function Get-InventoryConfigValue {
    param([Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)
    Get-SmartM365ConfigurationValue -Name $Name -EffectiveConfig $script:EffectiveConfig -DefaultValue $DefaultValue -Strict
}

function Resolve-InventoryConfigTokens {
    param([string]$Value, [string]$Name)
    Resolve-SmartM365ConfigurationTokens -Value $Value -Config $script:EffectiveConfig -Strict -Name $Name
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
    $included = @((Get-InventoryConfigValue -Name 'IncludedWebApplicationUrls' -DefaultValue @()))
    $excluded = @((Get-InventoryConfigValue -Name 'ExcludedWebApplicationUrls' -DefaultValue @()))
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

function Get-MySiteHostNames {
    $seen = @{}
    foreach ($entry in @((Get-InventoryConfigValue -Name 'MySiteHostNames' -DefaultValue @()))) {
        if ($entry -isnot [string] -or [string]::IsNullOrWhiteSpace($entry) -or $entry -ne $entry.Trim() -or
            [uri]::CheckHostName($entry) -eq [UriHostNameType]::Unknown) {
            throw 'MySiteHostNames must contain host names without a scheme, port, path or whitespace.'
        }
        $hostName = $entry.ToLowerInvariant()
        if (-not $seen.ContainsKey($hostName)) {
            $seen[$hostName] = $true
            $hostName
        }
    }
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

function Resolve-SiteCategory {
    param([string]$Url, [string]$RootWebTemplate, [string[]]$MySiteHostNames)
    if ($RootWebTemplate -ieq 'SPSPERS') { return [pscustomobject]@{ Category='Personal'; Status='Observed' } }
    if ([string]::IsNullOrWhiteSpace($RootWebTemplate)) { return [pscustomobject]@{ Category=''; Status='TemplateUnavailable' } }
    $siteUri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$siteUri) -or
        $siteUri.Scheme -notin @('http','https') -or [string]::IsNullOrWhiteSpace($siteUri.Host)) {
        return [pscustomobject]@{ Category=''; Status='InvalidUrl' }
    }
    if (@($MySiteHostNames).Count -eq 0) { return [pscustomobject]@{ Category=''; Status='HostListUnconfigured' } }
    foreach ($hostName in $MySiteHostNames) {
        if ($siteUri.Host -ieq $hostName) { return [pscustomobject]@{ Category='OtherOnMyHost'; Status='Observed' } }
    }
    return [pscustomobject]@{ Category='Other'; Status='Observed' }
}

function Resolve-QuotaUsage {
    param($StorageBytes, $QuotaLimitBytes)
    if (Test-MissingObservation $StorageBytes) { return [pscustomobject]@{ Percent=''; Status='StorageUnavailable' } }
    if (Test-MissingObservation $QuotaLimitBytes) { return [pscustomobject]@{ Percent=''; Status='QuotaUnavailable' } }
    [decimal]$storage = 0
    [decimal]$quota = 0
    $numberStyle = [Globalization.NumberStyles]::Integer
    $culture = [Globalization.CultureInfo]::InvariantCulture
    if (-not [decimal]::TryParse([string]$StorageBytes, $numberStyle, $culture, [ref]$storage) -or $storage -lt 0) {
        return [pscustomobject]@{ Percent=''; Status='InvalidStorage' }
    }
    if (-not [decimal]::TryParse([string]$QuotaLimitBytes, $numberStyle, $culture, [ref]$quota) -or $quota -lt 0) {
        return [pscustomobject]@{ Percent=''; Status='InvalidQuota' }
    }
    if ($quota -eq 0) { return [pscustomobject]@{ Percent=''; Status='NoQuota' } }
    $percent = ($storage * 100 / $quota).ToString('0.########', $culture)
    return [pscustomobject]@{ Percent=$percent; Status='Observed' }
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
    if ($Message -match 'access|denied|unauthorized|permission|acc[e\u00E8]s|refus[e\u00E9]') { return 'AccessDenied' }
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

function Save-ContentWeeklyHistory {
    param([string[]]$RunPaths, [string]$HistoryRoot, [int]$RetentionWeeks, [bool]$Partial)
    $destinationRoot = if ($Partial) { Join-Path $HistoryRoot 'Partial' } else { $HistoryRoot }
    $label = if ($Partial) { 'SharePoint on-prem content (partial; incomplete coverage)' } else { 'SharePoint on-prem content' }
    WriteLog -Message ("Saving {0} weekly history in {1}." -f $label, $destinationRoot)
    Add-SmartM365WeeklyHistory -SourceCsvPaths $RunPaths -HistoryRootPath $destinationRoot -RetentionWeeks $RetentionWeeks -HistoryLabel $label | Out-Null
    return $destinationRoot
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

function Get-ContentMailFarmEdition {
    param([string]$BuildVersion)
    try {
        $mapping = Import-PowerShellDataFile -Path $script:BuildEditionMapPath -ErrorAction Stop
        if ($mapping.Builds.ContainsKey($BuildVersion) -and $mapping.Builds[$BuildVersion] -in @('2016','2019')) {
            return [string]$mapping.Builds[$BuildVersion]
        }
    } catch { WriteLog -Message ("Farm edition mapping unavailable for the combined email: {0}" -f $_.Exception.Message) -Level WARNING }
    return 'Unknown'
}

function Get-VerifiedInfrastructureMailData {
    param([string]$LatestRoot, [string]$FarmId, [string]$TenantKey, [string[]]$ExpectedWebApplicationIds = @(), [datetime]$NowUtc = [datetime]::UtcNow)
    $result = [ordered]@{ Status='Unavailable'; Detail='Infrastructure receipt or CSVs unavailable'; CollectedAtUtc=''; Rows=$null }
    $names = @('Farms','Servers','ServiceApplications','WebApplications','WebApplicationZones','ContentDatabases')
    try {
        $receiptPath = Join-Path $LatestRoot 'SmartInventory_SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.current.json.txt'
        $receipt = Get-Content -LiteralPath $receiptPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($receipt.Owner -ne 'SmartInventory-SourceReceipt' -or $receipt.Status -ne 'Completed' -or
            $receipt.TenantKey -ne $TenantKey -or @($receipt.Files).Count -ne $names.Count) {
            throw 'Infrastructure receipt is not a completed six-file receipt for this tenant.'
        }
        $collectedAt = ([datetimeoffset]$receipt.CompletedAtUtc).UtcDateTime
        if ($collectedAt -eq [datetime]::MinValue -or $collectedAt -gt $NowUtc.AddMinutes(5)) { throw 'Infrastructure receipt timestamp is invalid.' }
        $result.CollectedAtUtc = $collectedAt.ToString('o')
        if ($collectedAt -lt $NowUtc.AddHours(-24)) {
            $result.Status = 'Stale'
            $result.Detail = 'Last completed infrastructure inventory is older than 24 hours.'
            return [pscustomobject]$result
        }
        $runPath = Join-Path $LatestRoot 'SmartInventory_SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.run.json.txt'
        if (Test-Path -LiteralPath $runPath -PathType Leaf) {
            $lastRun = Get-Content -LiteralPath $runPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($lastRun.Status -ne 'Completed' -and ([datetimeoffset]$lastRun.StartedAtUtc).UtcDateTime -gt $collectedAt) {
                $result.Status = 'Stale'
                $result.Detail = 'A newer infrastructure run failed; the last completed CSVs were not used.'
                return [pscustomobject]$result
            }
        }
        $rows = @{}
        $csvRunId = $null
        foreach ($name in $names) {
            $fileName = "SharePoint_OnPrem_$name.csv"
            $proof = @($receipt.Files | Where-Object { $_.File -eq $fileName -and $_.Status -eq 'Success' -and $_.RunId -eq $receipt.RunId })
            if ($proof.Count -ne 1) { throw "Infrastructure receipt entry invalid: $fileName" }
            $path = Join-Path $LatestRoot $fileName
            if ((Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash -ne $proof[0].SHA256) { throw "Infrastructure CSV hash mismatch: $fileName" }
            $csvRows = @(Import-Csv -LiteralPath $path -Delimiter ';' -ErrorAction Stop)
            if ($csvRows.Count -ne [int]$proof[0].Rows) { throw "Infrastructure CSV row count mismatch: $fileName" }
            foreach ($row in $csvRows) {
                if ($row.TenantKey -ne $TenantKey -or $row.FarmId -ne $FarmId -or [string]::IsNullOrWhiteSpace([string]$row.RunId)) {
                    throw "Infrastructure CSV scope mismatch: $fileName"
                }
                if ($null -eq $csvRunId) { $csvRunId = [string]$row.RunId }
                elseif ($row.RunId -ne $csvRunId) { throw "Infrastructure CSV run mismatch: $fileName" }
            }
            $rows[$name] = $csvRows
        }
        if ($rows.Farms.Count -ne 1 -or -not $csvRunId) { throw 'Infrastructure farm row or CSV run ID is missing.' }
        $expectedIds = @($ExpectedWebApplicationIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        if ($expectedIds.Count -gt 0) {
            $observedIds = @($rows.WebApplications | ForEach-Object { [string]$_.WebApplicationId } | Sort-Object -Unique)
            if (($expectedIds -join '|') -ne ($observedIds -join '|')) { throw 'Infrastructure and Content selected different web applications.' }
        }
        $result.Status = 'Qualified'
        $result.Detail = 'Six current CSVs verified against the completed source receipt.'
        $result.Rows = $rows
    } catch {
        $result.Detail = $_.Exception.Message
        WriteLog -Message ("Infrastructure data omitted from the combined email: {0}" -f $result.Detail) -Level WARNING
    }
    return [pscustomobject]$result
}

function New-ContentMailHtml {
    param([string]$Status, [hashtable]$Stats, [string[]]$Failures = @(), [int]$UploadFailures = 0, [string]$TenantName = '', [Nullable[timespan]]$Duration = $null, [string]$HostName = $env:COMPUTERNAME, $Infrastructure = $null, [string]$ContentStatus = '')
    if (-not $ContentStatus) { $ContentStatus = $Status }
    $specs = @(
        @('Web applications','WebApplications','#f8fafc','#dbe3ef','#0f172a'),
        @('Content databases','ContentDatabases','#eff6ff','#bfdbfe','#1d4ed8'),
        @('Site collections','SiteCollections','#f0fdf4','#bbf7d0','#166534'),
        @('Webs','Webs','#faf5ff','#e9d5ff','#7e22ce'),
        @('Coverage issues','CoverageIssues','#fff7ed','#fed7aa','#9a3412'),
        @('Read-only collections','ReadOnlyCollections','#f0fdfa','#99f6e4','#0f766e'),
        @('Unverified locks','UnverifiedLocks','#fef2f2','#fecaca','#b91c1c'),
        @('Content upload failures','UploadFailures','#eef2ff','#c7d2fe','#4338ca')
    )
    $cards = New-Object 'System.Collections.Generic.List[object]'
    foreach ($spec in $specs) {
        $value = if ($spec[1] -eq 'UploadFailures') { if ($Status -eq 'Failed') { 'n/a' } else { [string]$UploadFailures } }
                 elseif ($null -eq $Stats -or $null -eq $Stats[$spec[1]]) { 'n/a' } else { [string]$Stats[$spec[1]] }
        $detail = if ($spec[1] -eq 'UploadFailures' -and $Status -eq 'Failed') { 'not attempted' } elseif ($value -eq 'n/a') { 'not observed' } else { 'observed rows' }
        $cards.Add([pscustomobject]@{Label=$spec[0];Value=$value;Detail=$detail;Background=$spec[2];Border=$spec[3];Accent=$spec[4];Span=1})
    }
    $sections = New-Object 'System.Collections.Generic.List[object]'
    $sections.Add([pscustomobject]@{Title='Summary';Html=(New-SmartM365EmailKpiGridHtml -Cards $cards.ToArray() -RowSizes @(4,4))})
    if ($Infrastructure) {
        $sourceRows = @(
            @('Content',$ContentStatus,'Current inventory run'),
            @('Infrastructure',[string]$Infrastructure.Status,([string]$Infrastructure.CollectedAtUtc + ' ' + [string]$Infrastructure.Detail).Trim())
        )
        $sections.Add([pscustomobject]@{Title='Source status';Html=(New-SmartM365EmailTableHtml -Headers @('Source','Status','Detail') -Rows $sourceRows)})
        if ($Infrastructure.Status -eq 'Qualified' -and $Infrastructure.Rows) {
            $infraRows = $Infrastructure.Rows
            $serverCount = @($infraRows.Servers | Where-Object { $_.Role -ne 'Invalid' }).Count
            $externalCount = @($infraRows.Servers | Where-Object { $_.Role -eq 'Invalid' }).Count
            $farmBuild = [string]$infraRows.Farms[0].BuildVersion
            $infraCards = @(
                [pscustomobject]@{Label='SharePoint servers';Value=[string]$serverCount;Detail="$externalCount external";Background='#f8fafc';Border='#dbe3ef';Accent='#0f172a';Span=1},
                [pscustomobject]@{Label='Service applications';Value=[string]$infraRows.ServiceApplications.Count;Detail='inventory rows';Background='#eff6ff';Border='#bfdbfe';Accent='#1d4ed8';Span=1},
                [pscustomobject]@{Label='Farm edition';Value=(Get-ContentMailFarmEdition -BuildVersion $farmBuild);Detail=$farmBuild;Background='#f0fdfa';Border='#99f6e4';Accent='#0f766e';Span=1}
            )
            $sections.Add([pscustomobject]@{Title='Infrastructure';Html=(New-SmartM365EmailKpiGridHtml -Cards $infraCards -RowSizes @(3))})
        }
    }
    $alerts = New-Object 'System.Collections.Generic.List[object]'
    foreach ($failure in $Failures) { $alerts.Add(@('Run failure',[string]$failure,'')) }
    if ($Stats -and $null -ne $Stats.CoverageIssues -and $Stats.CoverageIssues -gt 0) { $alerts.Add(@('Coverage incomplete',"$($Stats.CoverageIssues) issue(s)",'See CollectionCoverage.csv')) }
    if ($Stats -and $null -ne $Stats.ReadOnlyCollections -and $Stats.ReadOnlyCollections -gt 0) { $alerts.Add(@('Read-only collections',"$($Stats.ReadOnlyCollections) collection(s)",'See SiteCollections.csv')) }
    if ($Stats -and $null -ne $Stats.UnverifiedLocks -and $Stats.UnverifiedLocks -gt 0) { $alerts.Add(@('Lock state unverified',"$($Stats.UnverifiedLocks) collection(s)",'See SiteCollections.csv')) }
    if ($UploadFailures -gt 0) { $alerts.Add(@('SharePoint upload incomplete',"$UploadFailures file(s)",'')) }
    if ($Infrastructure) {
        if ($Infrastructure.Status -ne 'Qualified') {
            $alerts.Add(@('Infrastructure data unavailable',[string]$Infrastructure.Status,[string]$Infrastructure.Detail))
        } elseif ($Infrastructure.Rows) {
            foreach ($server in @($Infrastructure.Rows.Servers | Where-Object { $_.Role -ne 'Invalid' })) {
                if ($server.Status -and $server.Status -ne 'Online') { $alerts.Add(@('Server not online',[string]$server.ServerName,[string]$server.Status)) }
            }
            foreach ($service in $Infrastructure.Rows.ServiceApplications) {
                if ($service.Status -and $service.Status -notin @('Online','Started')) { $alerts.Add(@('Service application not started',[string]$service.Name,[string]$service.Status)) }
            }
            foreach ($database in $Infrastructure.Rows.ContentDatabases) {
                if ($database.Status -and $database.Status -ne 'Online') { $alerts.Add(@('Content database not online',[string]$database.Name,[string]$database.Status)) }
                if ($database.IsReadOnly -eq 'True') { $alerts.Add(@('Content database read-only',[string]$database.Name,'True')) }
                if ((Get-ObservedProperty $database 'NeedsUpgradeIncludeChildren') -eq 'True' -and (Get-ObservedProperty $database 'NeedsUpgradeStatus') -eq 'Collected') {
                    $alerts.Add(@('Content database needs upgrade',[string]$database.Name,'True'))
                }
            }
            $unverifiedUpgradeCount = @($Infrastructure.Rows.ContentDatabases | Where-Object { (Get-ObservedProperty $_ 'NeedsUpgradeStatus') -notin @('Collected','') }).Count
            if ($unverifiedUpgradeCount -gt 0) { $alerts.Add(@('Upgrade state unavailable',"$unverifiedUpgradeCount content databases",'Check NeedsUpgradeStatus in CSV')) }
        }
    }
    if ($alerts.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Alerts';Html=(New-SmartM365EmailTableHtml -Headers @('Issue','Object','Detail') -Rows $alerts.ToArray())}) }
    if ($Infrastructure -and $Infrastructure.Status -eq 'Qualified' -and $Infrastructure.Rows) {
        $serverRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($server in @($Infrastructure.Rows.Servers | Where-Object { $_.Role -ne 'Invalid' } | Sort-Object ServerName)) {
            $serverRows.Add(@([string]$server.ServerName,[string]$server.Role,[string]$server.Status))
        }
        if ($serverRows.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Servers';Html=(New-SmartM365EmailTableHtml -Headers @('Name','Role','Status') -Rows $serverRows.ToArray())}) }
        $webRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($web in @($Infrastructure.Rows.WebApplications | Sort-Object DefaultUrl)) {
            $zones = @($Infrastructure.Rows.WebApplicationZones | Where-Object { $_.WebApplicationId -eq $web.WebApplicationId } | Sort-Object Zone | ForEach-Object { '{0}: {1}' -f $_.Zone,$_.AuthenticationMode }) -join '; '
            $webRows.Add(@([string]$web.DefaultUrl,[string]$web.ApplicationPool,[string]$web.ContentDatabaseCount,[string]$zones))
        }
        if ($webRows.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Web applications';Html=(New-SmartM365EmailTableHtml -Headers @('URL','Application pool','Content databases','Authentication by zone') -Rows $webRows.ToArray() -NoWrapColumns @(0,1))}) }
        $sized = @($Infrastructure.Rows.ContentDatabases | Where-Object { (Get-ObservedProperty $_ 'DatabaseSizeStatus') -eq 'Collected' -and (Get-ObservedProperty $_ 'DatabaseSizeBytes') -match '^\d+$' } | Sort-Object @{Expression={[decimal]$_.DatabaseSizeBytes};Descending=$true} | Select-Object -First 5)
        if ($sized.Count -gt 0) {
            $sizeRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($database in $sized) { $sizeRows.Add(@([string]$database.Name,[string]$database.DatabaseServer,[string]$database.DatabaseSizeBytes)) }
            $sections.Add([pscustomobject]@{Title='Top 5 largest content databases';Html=(New-SmartM365EmailTableHtml -Headers @('Name','Server','Size (bytes)') -Rows $sizeRows.ToArray())})
        }
    }
    if ($Stats -and $Stats.CoverageSamples.Count -gt 0) {
        $issueRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $Stats.CoverageSamples) { $issueRows.Add(@([string]$item.Status,[string]$item.Url,[string]$item.ErrorMessage)) }
        $sections.Add([pscustomobject]@{Title='Collection coverage examples';Html=(New-SmartM365EmailTableHtml -Headers @('Status','Collection or database','Detail') -Rows $issueRows.ToArray())})
    }
    if ($Stats -and $Stats.ReadOnlySamples.Count -gt 0) {
        $readOnlyRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $Stats.ReadOnlySamples) { $readOnlyRows.Add(@([string]$item.Url,[string]$item.LockState,[string]$item.ContentDatabaseIsReadOnly)) }
        $sections.Add([pscustomobject]@{Title='Read-only collection examples';Html=(New-SmartM365EmailTableHtml -Headers @('URL','Lock state','Database read-only') -Rows $readOnlyRows.ToArray() -NoWrapColumns @(0))})
    }
    if ($Stats -and $Status -in @('Qualified','Partial') -and $null -ne $Stats.SiteCollections) {
        $fileRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in @(@('SiteCollections',$Stats.SiteCollections),@('SiteAdministrators',$Stats.SiteAdministrators),@('Webs',$Stats.Webs),@('CollectionCoverage',$Stats.CollectionCoverageRows))) {
            $fileRows.Add(@("SharePoint_OnPrem_$($item[0]).csv",[string]$item[1]))
        }
        if ($Infrastructure -and $Infrastructure.Status -eq 'Qualified' -and $Infrastructure.Rows) {
            foreach ($name in @('Farms','Servers','ServiceApplications','WebApplications','WebApplicationZones','ContentDatabases')) {
                $fileRows.Add(@("SharePoint_OnPrem_$name.csv",[string]$Infrastructure.Rows[$name].Count))
            }
        }
        $sections.Add([pscustomobject]@{Title='Files';Html=(New-SmartM365EmailTableHtml -Headers @('CSV file','Rows') -Rows $fileRows.ToArray())})
    }
    $severity = switch ($Status) { 'Qualified' {'Success'} 'Partial' {'Warning'} default {'Error'} }
    $durationText = if ($null -ne $Duration) { ([timespan]$Duration).ToString('hh\:mm\:ss') } else { '' }
    $message = if ($Status -eq 'Failed') { 'SharePoint on-premises content inventory stopped before full publication. Values below reflect observations made before the failure.' } else { 'SharePoint on-premises content inventory summary generated from the current SmartM365 CSV outputs.' }
    return New-SmartM365EmailBody -Title 'SharePoint on-prem inventory summary' -Category 'SmartM365 SharePoint OnPrem' -Severity $severity -Tenant $TenantName -HostName $HostName -StatusBadge $Status.ToUpperInvariant() -Duration $durationText -Message $message -Sections $sections.ToArray() -Footer 'This automated message was generated by SmartM365 from SharePoint on-premises inventory data.'
}

function Send-ContentRunMail {
    param([string]$Status, [string[]]$Failures = @(), [int]$UploadFailures = 0)
    try {
        $infrastructure = [pscustomobject]@{Status='Unavailable';Detail='Content run could not establish farm identity.';CollectedAtUtc='';Rows=$null}
        $farmVariable = Get-Variable -Name farm -Scope Script -ErrorAction SilentlyContinue
        if ($farmVariable -and $farmVariable.Value -and $farmVariable.Value.Id) {
            $webAppsVariable = Get-Variable -Name webApplications -Scope Script -ErrorAction SilentlyContinue
            $selectedIds = if ($webAppsVariable) { @($webAppsVariable.Value | ForEach-Object { [string]$_.Id }) } else { @() }
            $infrastructure = Get-VerifiedInfrastructureMailData -LatestRoot $latestRoot -FarmId ([string]$farmVariable.Value.Id) -TenantKey $global:SmartM365TenantKey -ExpectedWebApplicationIds $selectedIds
        }
        $mailStatus = if ($Status -eq 'Failed') { 'Failed' } elseif ($Status -eq 'Qualified' -and $infrastructure.Status -eq 'Qualified' -and $UploadFailures -eq 0) { 'Qualified' } else { 'Partial' }
        $markerName = if ($mailStatus -eq 'Qualified') { 'SmartM365-SharePoint-OnPrem-Content-DailySummary.sent' } else { 'SmartM365-SharePoint-OnPrem-Content-IssueDailySummary.sent' }
        $markerPath = Join-Path $outputBase $markerName
        $duration = (Get-Date) - $script:StartedAt
        $body = New-ContentMailHtml -Status $mailStatus -ContentStatus $Status -Infrastructure $infrastructure -Stats $script:MailStats -Failures $Failures -UploadFailures $UploadFailures -TenantName $Tenant -Duration $duration
        $null = Invoke-DailySummaryMail -MarkerPath $markerPath -Force:$ForceSendDailySummary -SendAction { Send-InventorySummaryMail -Subject 'SharePoint on-prem inventory' -BodyHtml $body }
    } catch {
        try { WriteLog -Message ("Content summary email failed without affecting inventory qualification: {0}" -f $_.Exception.Message) -Level WARNING }
        catch { [Console]::Error.WriteLine('Content summary email failed; its warning could not be logged.') }
    }
}

function Publish-QualifiedCsvUploads {
    param([string[]]$Paths)
    if (-not $global:EnableSharePointUpload) { return 0 }
    $failures = 0
    foreach ($path in @($Paths | Sort-Object -Unique)) {
        $policy=Get-Variable SmartM365CollectorUploadPolicies -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        if ($null -ne $policy) {
            $category=Get-SmartM365SharePointUploadCategory -Path $path -DataAllRoot $outputBase -DataLastRoot $latestRoot -LogAllRoot $global:LogPath
            if ($category -and -not (Test-SmartM365CollectorUploadEnabled -Category $category -Policies $policy -RunEnabled $global:EnableSharePointUpload)) { continue }
        }
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
Import-Module -Name $coreManifest -MinimumVersion '1.0.70' -ErrorAction Stop
. (Join-Path $root 'Modules\SmartM365.Core\SmartM365-SharePointUploadPolicy.ps1')
$oldPolicyVariable=Get-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue
$oldPolicy=if($oldPolicyVariable){@{Value=$oldPolicyVariable.Value}}else{$null}
$oldRootsVariable=Get-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue
$oldRoots=if($oldRootsVariable){@{Value=$oldRootsVariable.Value}}else{$null}
$legacyUpload=[bool]$global:EnableSharePointUpload
$outputBase = if ($PSBoundParameters.ContainsKey('OutputRoot')) { Resolve-InventoryConfigTokens -Value $OutputRoot -Name 'OutputRoot' } else { Get-InventoryConfigValue -Name 'OutputRoot' }
$outputBase = Assert-InventoryPath -Path $outputBase -Name 'OutputRoot'
$latestRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue -Name 'LatestCsvFolderPath') -Name 'LatestCsvFolderPath'
$runBase = if ($MaxItems -gt 0) { Join-Path $outputBase 'TEST' } else { $outputBase }
$runFolder = Join-Path $runBase ((Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:RunId.Substring(0,8))
$script:StartedAt = Get-Date
$script:MailStats = @{
    WebApplications=$null; ContentDatabases=$null; SiteCollections=$null; SiteAdministrators=$null
    Webs=$null; CollectionCoverageRows=$null; CoverageIssues=$null; ReadOnlyCollections=$null
    UnverifiedLocks=$null; CoverageSamples=(New-Object 'System.Collections.Generic.List[object]')
    ReadOnlySamples=(New-Object 'System.Collections.Generic.List[object]')
}
$script:Completed = $false
$script:FailureHandled = $false
try {
    $null = InitializeScriptEnvironment -OutputPath $runFolder -LogFileName $script:ScriptName -CallerScriptPath $PSCommandPath
    $global:SmartM365CollectorUploadPolicies=Resolve-SmartM365CollectorUploadPolicies -LocalConfig $localConfig -GlobalConfig $script:EffectiveConfig -LegacyEnabled ([bool]$merged['EnableSharePointUpload']) -ExternalActionsAllowed (-not $ValidateOnly -and $MaxItems -eq 0)
    $global:SmartM365CollectorUploadRoots=[pscustomobject]@{DataAllRoot=$outputBase;DataLastRoot=$latestRoot;LogAllRoot=$global:LogPath}
    Assert-SmartM365CollectorUploadRoots -DataAllRoot $outputBase -DataLastRoot $latestRoot -LogAllRoot $global:LogPath -Policies $global:SmartM365CollectorUploadPolicies
    $global:EnableSharePointUpload=@($global:SmartM365CollectorUploadPolicies.Values | Where-Object Enabled).Count -gt 0
    WriteLog -Message ("SharePoint content inventory starting. SharePoint upload enabled for this run: {0}." -f $global:EnableSharePointUpload)
    if (-not $ValidateOnly -and $MaxItems -eq 0) {
        try { Ensure-GraphAuthenticationModule }
        catch { WriteLog -Message ("Graph Authentication bootstrap incomplete; local inventory will continue: {0}" -f $_.Exception.Message) -Level WARNING }
    }
    $farm = Test-SharePointPrerequisites
    $webApplications = @(Get-ConfiguredWebApplications)
    $mySiteHostNames = @(Get-MySiteHostNames)
    $excludePersonalCollections = Get-InventoryConfigValue -Name 'ExcludePersonalSiteCollections' -DefaultValue $true
    if ($excludePersonalCollections -isnot [bool]) { throw 'ExcludePersonalSiteCollections must be a JSON boolean.' }
    if ($mySiteHostNames.Count -eq 0) { WriteLog -Message 'MySiteHostNames is empty; non-personal site category will remain unclassified.' -Level WARNING }
    WriteLog -Message ("ExcludePersonalSiteCollections={0}." -f $excludePersonalCollections)
    $script:MailStats.WebApplications = $webApplications.Count
    $databaseGroups = New-Object 'System.Collections.Generic.List[object]'
    foreach ($webApplication in $webApplications) {
        foreach ($database in @(Get-SPContentDatabase -WebApplication $webApplication -ErrorAction Stop)) {
            $databaseGroups.Add([pscustomobject]@{ WebApplication=$webApplication; ContentDatabase=$database })
        }
    }
    $script:MailStats.ContentDatabases = $databaseGroups.Count
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
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $latestRoot -ReadOnly:($MaxItems -gt 0) -ScopeParameters @{ IncludedWebApplicationUrls=@((Get-InventoryConfigValue 'IncludedWebApplicationUrls' @())); ExcludedWebApplicationUrls=@((Get-InventoryConfigValue 'ExcludedWebApplicationUrls' @())); MySiteHostNames=$mySiteHostNames; ExcludePersonalSiteCollections=$excludePersonalCollections; MaxItems=$MaxItems }
    $farmId = [string]$farm.Id
    $rows = @{
        SiteCollections = New-Object 'System.Collections.Generic.List[object]'
        SiteAdministrators = New-Object 'System.Collections.Generic.List[object]'
        Webs = New-Object 'System.Collections.Generic.List[object]'
        CollectionCoverage = New-Object 'System.Collections.Generic.List[object]'
    }
    $schemas = [ordered]@{
        SiteCollections = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','ContentDatabaseId','SiteCollectionId','Url','OwnerLogin','SecondaryOwnerLogin','StorageBytes','QuotaLimitBytes','QuotaUsagePercent','QuotaUsageStatus','LastContentModifiedUtc','RootWebTemplate','SiteCategory','SiteCategoryStatus','LanguageLcid','LockState','IsReadOnly','ReadLocked','WriteLocked','LockIssue','ContentDatabaseIsReadOnly','LockStatus','CollectionStatus')
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
    foreach ($name in @('SiteCollections','SiteAdministrators','Webs','CollectionCoverageRows','CoverageIssues','ReadOnlyCollections','UnverifiedLocks')) { $script:MailStats[$name] = 0 }
    $failureCount = 0
    $processed = 0
    $excludedPersonalCount = 0
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
                $script:MailStats.SiteCollections++
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
                    if ($lock.LockStatus -ne 'Observed') { $script:MailStats.UnverifiedLocks++ }
                    $rootWeb = $site.RootWeb
                    try {
                        $rootWebTemplate = [string](Get-ObservedProperty $rootWeb 'WebTemplate')
                        if ($excludePersonalCollections -and $rootWebTemplate -ieq 'SPSPERS') {
                            $coverage['Status'] = 'ExcludedPersonal'
                            $script:MailStats.SiteCollections--
                            $excludedPersonalCount++
                            return
                        }
                        $collection = New-InventoryRow $farmId
                        $collection['WebApplicationId'] = [string]$webApplication.Id
                        $collection['ContentDatabaseId'] = [string]$database.Id
                        $collection['SiteCollectionId'] = $siteId
                        $collection['Url'] = $siteUrl
                        $collection['OwnerLogin'] = [string](Get-ObservedProperty (Get-ObservedProperty $site 'Owner') 'LoginName')
                        $collection['SecondaryOwnerLogin'] = [string](Get-ObservedProperty (Get-ObservedProperty $site 'SecondaryContact') 'LoginName')
                        $collection['StorageBytes'] = [string](Get-ObservedProperty (Get-ObservedProperty $site 'Usage') 'Storage')
                        $collection['QuotaLimitBytes'] = [string](Get-ObservedProperty (Get-ObservedProperty $site 'Quota') 'StorageMaximumLevel')
                        $quotaUsage = Resolve-QuotaUsage -StorageBytes $collection['StorageBytes'] -QuotaLimitBytes $collection['QuotaLimitBytes']
                        $collection['QuotaUsagePercent'] = $quotaUsage.Percent
                        $collection['QuotaUsageStatus'] = $quotaUsage.Status
                        $collection['LastContentModifiedUtc'] = Convert-ToUtcText (Get-ObservedProperty $site 'LastContentModifiedDate')
                        $collection['RootWebTemplate'] = $rootWebTemplate
                        $category = Resolve-SiteCategory -Url $siteUrl -RootWebTemplate $collection['RootWebTemplate'] -MySiteHostNames $mySiteHostNames
                        $collection['SiteCategory'] = $category.Category
                        $collection['SiteCategoryStatus'] = $category.Status
                        $collection['LanguageLcid'] = [string](Get-ObservedProperty $rootWeb 'Language')
                        foreach ($name in @('LockState','IsReadOnly','ReadLocked','WriteLocked','LockIssue','ContentDatabaseIsReadOnly','LockStatus')) { $collection[$name] = $lock.$name }
                        $collection['CollectionStatus'] = 'Collected'
                        $missingFields = New-Object 'System.Collections.Generic.List[string]'
                        if ($lock.LockStatus -ne 'Observed') { $missingFields.Add('LockState') }
                        foreach ($field in @('StorageBytes','QuotaLimitBytes','LastContentModifiedUtc','OwnerLogin','RootWebTemplate','LanguageLcid','IsReadOnly','ReadLocked','WriteLocked','ContentDatabaseIsReadOnly')) {
                            if (Test-MissingObservation $collection[$field]) { $missingFields.Add($field) }
                        }
                        if ($quotaUsage.Status -in @('InvalidStorage','InvalidQuota')) { $missingFields.Add('QuotaUsagePercent') }
                        if ($category.Status -eq 'InvalidUrl') { $missingFields.Add('SiteCategory') }
                        if ($missingFields.Count -gt 0) {
                            $collection['CollectionStatus'] = 'MetadataIncomplete'
                            $coverage['Status'] = 'MetadataIncomplete'
                            $coverage['ErrorStage'] = 'SiteCollectionMetadata'
                            $coverage['ErrorMessage'] = 'Missing or unverified observations: ' + ($missingFields -join ', ')
                        }
                        $siteRow = [pscustomobject]$collection
                        $rows.SiteCollections.Add($siteRow)
                        if ($lock.LockState -eq 'ReadOnly' -or $lock.IsReadOnly -eq 'True' -or $lock.ContentDatabaseIsReadOnly -eq 'True') {
                            $script:MailStats.ReadOnlyCollections++
                            if ($script:MailStats.ReadOnlySamples.Count -lt 5) { $script:MailStats.ReadOnlySamples.Add([pscustomobject]@{Url=$siteUrl;LockState=$lock.LockState;ContentDatabaseIsReadOnly=$lock.ContentDatabaseIsReadOnly}) }
                        }
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
                            $script:MailStats.SiteAdministrators++
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
                            $script:MailStats.Webs++
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
                        foreach ($name in @('OwnerLogin','SecondaryOwnerLogin','StorageBytes','QuotaLimitBytes','QuotaUsagePercent','LastContentModifiedUtc','RootWebTemplate','SiteCategory','LanguageLcid','LockState','IsReadOnly','ReadLocked','WriteLocked','LockIssue','ContentDatabaseIsReadOnly','LockStatus')) { $fallback[$name] = '' }
                        $fallback['QuotaUsageStatus'] = 'NotCollected'
                        $fallback['SiteCategoryStatus'] = 'NotCollected'
                        if ($null -ne $lock) { foreach ($name in @('LockState','IsReadOnly','ReadLocked','WriteLocked','LockIssue','ContentDatabaseIsReadOnly','LockStatus')) { $fallback[$name] = $lock.$name } }
                        $fallback['CollectionStatus'] = $coverage['Status']
                        $rows.SiteCollections.Add([pscustomobject]$fallback)
                    }
                    $failureCount++
                    WriteLog -Message ("Site collection {0} failed: {1}" -f $siteUrl, $message) -Level ERROR
                    if ($message -eq 'GlobalTimeout') { $globalTimedOut = $true }
                } finally {
                    $rows.CollectionCoverage.Add([pscustomobject]$coverage)
                    $script:MailStats.CollectionCoverageRows++
                    if ($coverage['Status'] -notin @('Collected','ExcludedPersonal') -and $script:MailStats.CoverageSamples.Count -lt 5) { $script:MailStats.CoverageSamples.Add([pscustomobject]@{Status=$coverage['Status'];Url=$coverage['Url'];ErrorMessage=$coverage['ErrorMessage']}) }
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
            $script:MailStats.CollectionCoverageRows++
            if ($script:MailStats.CoverageSamples.Count -lt 5) { $script:MailStats.CoverageSamples.Add([pscustomobject]@{Status=$coverage['Status'];Url=([string]$database.Id);ErrorMessage=$coverage['ErrorMessage']}) }
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
    $script:MailStats.CoverageIssues = $failureCount
    foreach ($kind in $schemas.Keys) { Flush-RunRows -Kind $kind -Buffers $rows -Paths $spoolPaths -MinimumCount 1 }
    WriteLog -Message ("Lock observation: observed={0}; unverifiedOrConflicting={1}." -f $lockObservedCount,$lockUnverifiedCount)
    $coverageLevel = if ($globalTimedOut) { 'ERROR' } elseif ($failureCount -gt 0) { 'WARNING' } else { 'INFO' }
    WriteLog -Message ("Collection coverage: processed={0}; excludedPersonal={1}; failed={2}; limited={3}; globalTimeout={4}." -f $processed,$excludedPersonalCount,$failureCount,$limited,$globalTimedOut) -Level $coverageLevel
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
        Send-ContentRunMail -Status Failed -Failures @('Global collection timeout; current CSVs and source receipt were not promoted.')
        Complete-SmartM365ExecutionContext -Status Failed
        $script:Completed = $true
        exit 1
    }
    $partialInventory = $publicationDecision -eq 'Partial'
    if ($partialInventory) {
        WriteLog -Message ("Partial content inventory: {0} coverage or metadata issue(s) are recorded in CollectionCoverage.csv. Current CSVs will be published with a partial source receipt and a separate partial weekly snapshot." -f $failureCount) -Level WARNING
    }
    if ([bool](Get-InventoryConfigValue 'EnableWeeklyHistory' $true)) {
        $historyRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue 'WeeklyHistoryFolderPath') -Name 'WeeklyHistoryFolderPath'
        try {
            Save-ContentWeeklyHistory -RunPaths @($runPaths) -HistoryRoot $historyRoot -RetentionWeeks ([int](Get-InventoryConfigValue 'WeeklyHistoryRetentionWeeks' 52)) -Partial $partialInventory | Out-Null
        } catch {
            if (-not $partialInventory) { throw }
            WriteLog -Message ("Partial weekly history could not be saved or uploaded; current CSV publication will continue: {0}" -f $_.Exception.Message) -Level WARNING
        }
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
    $runStatus = if ($partialInventory) { 'Partial' } else { 'Qualified' }
    Send-ContentRunMail -Status $runStatus -UploadFailures $uploadFailures
    if ($partialInventory -or $uploadFailures -gt 0) {
        Complete-SmartM365ExecutionContext -Status CompletedWithWarnings
        $script:Completed = $true
        exit 3
    }
    Complete-SmartM365ExecutionContext -Status Success
    $script:Completed = $true
} catch {
    $script:FailureHandled = $true
    if (-not $ValidateOnly -and $MaxItems -eq 0 -and -not $script:Completed) { Send-ContentRunMail -Status Failed -Failures @($_.Exception.Message) }
    try { WriteLog -Message ("Content inventory failed: {0}" -f $_.Exception.Message) -Level ERROR } catch { Write-Error $_ }
    if (-not $script:Completed) { try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $_ -FailureStage 'ContentInventory' } catch {} }
    throw
} finally {
    if (-not $script:Completed -and -not $script:FailureHandled) {
        $global:EnableSharePointUpload = $false
        try { WriteLog -Message ("Content inventory interrupted by Ctrl+C. RunId={0}; no further publication or mail will be attempted." -f $script:RunId) -Level WARNING }
        catch { [Console]::Error.WriteLine('Content inventory interrupted; the run log could not be updated.') }
        try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage 'Interrupted' }
        catch { [Console]::Error.WriteLine('Content inventory interrupted; execution cleanup failed: ' + $_.Exception.Message) }
    }
    $priorPolicy=Get-Variable oldPolicy -ErrorAction SilentlyContinue
    if($priorPolicy){if($priorPolicy.Value){Set-Variable SmartM365CollectorUploadPolicies -Scope Global -Value $priorPolicy.Value.Value}else{Remove-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue}}
    $priorRoots=Get-Variable oldRoots -ErrorAction SilentlyContinue
    if($priorRoots){if($priorRoots.Value){Set-Variable SmartM365CollectorUploadRoots -Scope Global -Value $priorRoots.Value.Value}else{Remove-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue}}
    $priorUpload=Get-Variable legacyUpload -ErrorAction SilentlyContinue
    if (($script:Completed -or $script:FailureHandled) -and $priorUpload) { $global:EnableSharePointUpload=[bool]$priorUpload.Value }
}

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA7B7AnzTO8YR4F
# mVo0MccI1bSSfr2OER7udHWxMW2kW6CCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgLxIrINEoCRhj
# foYUT/L2RhWP2pcWC8whpU/GKAr4I/swDQYJKoZIhvcNAQEBBQAEggGAWQM4q22a
# Q2cn1Dos/hiX5JFmPxHTJBvUUOTv58XYcgvjO/N1dNffIeuPvfx5Hi4ReHKwda/G
# N/W7sYKF/NsU74CibDt8zww0YAJkRaWPsP/28uyGfSZyq9uxZnCpp7h2nlWcCHN/
# qmGRwyiRRM33B3HQchPb3i3L9yMqBleVKnbDoiXvLqRJ61F5k2BysH7Vi2qYY7Ig
# Yq/FnPkTokPYn1y7LzHuKBT78SCk5fIkHgj5P0gYtja23RwCfjqjZ3u5uQDTPbvu
# QT6O2wMC7mibNCbJFAzjquecj3lDSc6q4d8ONgS9jwLFryZovRTQgMj5h0ihAQmA
# vGPExTVV/DkTFkg/YRUEshlS1CKUnOUxx6Z8UyYhm0FH40qf8C8PMSeNd7SaQKA8
# uBZCtYRgqBXELvpXlycv+7SN57c51dbYNHScSMTgoh6XItn46o0r8pQ37lJr4mvU
# hkxYoVvWdJ45E9Sv8hmy2HN6repYZ/HN2mTT4RoxlnwQfWZncyTw6Tn/oYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTAxMTY1OTIwWjA/BgkqhkiG9w0BCQQxMgQwBByFTooSAA0wUskLWwLAJiym
# U1S1DnsE/18DouxkQKm8d5OulmY5Od73MaDNkB53MA0GCSqGSIb3DQEBAQUABIIC
# AFV03w2Nlr0tIlwKEQWFyN4j7pOYY2tvRuH8FgmwGhvfrz4XWXOHmUBpI2mNsXoV
# xRc8EQsMgLcdHvsiMKenmDdWXLiMslqPnbJbi5jEQAhvuEsqfOcs7PMJ76QkhaPz
# h0njWz2Y8UEgW+1GtGuauRoLkw/0xuyzi5+mcrMi6e2KQu+BHuIWjJqU6woNSUaT
# LiNzeDlvSZP4EkNajI9iK7BUaMscyAlv0h+lEzFIXPRU8UGA6tnTk6rTjUa3ILpW
# XudilwjuDNDSudhq9h+dYiC5YFOebn03mdl51Il6g0Ul3/tbQZ3hwAcGBK+CQrvm
# ebLClQve+7tJSafTSMWnvdFaeRgS4CrkGYRxOGxgB/AMtWKcj2s9qJBZzmEEmvNu
# mr+4SY772dgeECV/kZoAkd8nWxGw/Eu2ZJDpAEh10LJK/m0NcY/BKa1S+x4ZVEfF
# m27Mq/WMYrOj+eIGSPaznw9ddn07Y/tNasVb7DBWmPRU+riH7E7BXpZTO8+ouc7Z
# DRTVOZepnIG5Cf/jH7m53/TRbHiLoGNXT7mduAfrJrEJrMcBVjVeZZd3Ru8e1Iea
# B/ZxuM+P/P3Uv/2PNFmcMPCdsOkMZ5ZouMGQ7MCQzUnpXnHSf5kBmYDtoO1a0ehD
# 3yryi4JFaM7kcSW84YZc5IzLoRqOEvw1GBthqhWELjHg
# SIG # End signature block
