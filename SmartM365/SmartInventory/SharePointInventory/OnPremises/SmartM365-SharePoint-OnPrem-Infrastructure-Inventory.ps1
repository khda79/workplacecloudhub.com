<#
.SYNOPSIS
    Read-only SharePoint Server 2016/2019 farm infrastructure inventory.
.VERSION
    1.0.5
.REQUIREMENTS
    Windows PowerShell 5.1 x64 on a SharePoint farm server; SharePoint Shell access.
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$OutputRoot,
    [switch]$ValidateOnly,
    [switch]$ForceSendDailySummary
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ScriptName = 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory'
$script:RunId = [guid]::NewGuid().ToString('N')
$script:CollectedAtUtc = [datetime]::UtcNow.ToString('o')

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
$global:EnableSharePointUpload = [bool]$merged['EnableSharePointUpload'] -and -not $ValidateOnly

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
            WriteLog -Message 'Daily infrastructure summary email is being evaluated by another run; email skipped.' -Level INFO
            return $false
        }
        $lastSentDate = if (Test-Path -LiteralPath $MarkerPath -PathType Leaf) { [string](Get-Content -LiteralPath $MarkerPath -Raw -ErrorAction Stop) } else { '' }
        if (-not $Force -and $lastSentDate.Trim() -eq $today) {
            WriteLog -Message "Daily infrastructure summary email already sent for $today; email skipped." -Level INFO
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
    WriteLog -Message "Daily infrastructure summary email sent to $to."
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

$coreManifest = Join-Path $root 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
Import-Module -Name $coreManifest -MinimumVersion '1.0.50' -ErrorAction Stop
$outputBase = if ($PSBoundParameters.ContainsKey('OutputRoot')) { Resolve-InventoryConfigTokens -Value $OutputRoot -Name 'OutputRoot' } else { Get-InventoryConfigValue -Name 'OutputRoot' }
$outputBase = Assert-InventoryPath -Path $outputBase -Name 'OutputRoot'
$latestRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue -Name 'LatestCsvFolderPath') -Name 'LatestCsvFolderPath'
$runFolder = Join-Path $outputBase ((Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:RunId.Substring(0,8))
$script:StartedAt = Get-Date
$script:ReceiptStarted = $false
$script:Completed = $false
try {
    $null = InitializeScriptEnvironment -OutputPath $runFolder -LogFileName $script:ScriptName -CallerScriptPath $PSCommandPath
    WriteLog -Message ("SharePoint infrastructure inventory starting. SharePoint upload enabled for this run: {0}." -f $global:EnableSharePointUpload)
    if (-not $ValidateOnly) {
        try { Ensure-GraphAuthenticationModule }
        catch { WriteLog -Message ("Graph Authentication bootstrap incomplete; local inventory will continue: {0}" -f $_.Exception.Message) -Level WARNING }
    }
    $farm = Test-SharePointPrerequisites
    $webApplications = @(Get-ConfiguredWebApplications)
    WriteLog -Message ("Farm {0}: {1} selected content web applications." -f $farm.Id, $webApplications.Count)
    if ($ValidateOnly) {
        WriteLog -Message 'ValidateOnly completed. No farm objects were traversed and no CSV or receipt was published.'
        Complete-SmartM365ExecutionContext -Status Success
        $script:Completed = $true
        return
    }

    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $latestRoot -ScopeParameters @{ IncludedWebApplicationUrls=@(Get-InventoryConfigValue 'IncludedWebApplicationUrls' @()); ExcludedWebApplicationUrls=@(Get-InventoryConfigValue 'ExcludedWebApplicationUrls' @()) }
    $script:ReceiptStarted = $true
    $farmId = [string]$farm.Id
    $rows = @{
        Farms = New-Object 'System.Collections.Generic.List[object]'
        Servers = New-Object 'System.Collections.Generic.List[object]'
        ServiceApplications = New-Object 'System.Collections.Generic.List[object]'
        WebApplications = New-Object 'System.Collections.Generic.List[object]'
        WebApplicationZones = New-Object 'System.Collections.Generic.List[object]'
        ContentDatabases = New-Object 'System.Collections.Generic.List[object]'
    }
    $failures = New-Object 'System.Collections.Generic.List[string]'
    $servers = @()
    $serviceApplications = @()
    try { $servers = @(Get-SPServer -ErrorAction Stop) }
    catch { $failures.Add("Servers: $($_.Exception.Message)"); WriteLog -Message "Server inventory failed: $($_.Exception.Message)" -Level ERROR }
    try { $serviceApplications = @(Get-SPServiceApplication -ErrorAction Stop) }
    catch { $failures.Add("ServiceApplications: $($_.Exception.Message)"); WriteLog -Message "Service application inventory failed: $($_.Exception.Message)" -Level ERROR }

    $farmRow = New-InventoryRow $farmId
    $farmBuild = Get-ObservedProperty $farm 'BuildVersion'
    $farmRow['FarmVersion'] = [string](Get-ObservedProperty $farmBuild 'Major')
    $farmRow['BuildVersion'] = [string]$farmBuild
    $farmRow['ConfigurationDatabase'] = [string](Get-ObservedProperty (Get-ObservedProperty $farm 'ConfigurationDatabase') 'Name')
    $farmRow['ServerCount'] = if (@($failures | Where-Object { $_ -like 'Servers:*' }).Count) { '' } else { $servers.Count }
    $farmRow['ServiceApplicationCount'] = if (@($failures | Where-Object { $_ -like 'ServiceApplications:*' }).Count) { '' } else { $serviceApplications.Count }
    if ($servers.Count -eq 0) { $failures.Add('Servers: no farm server was returned') }
    if (-not $farmRow['FarmVersion'] -or -not $farmRow['BuildVersion'] -or -not $farmRow['ConfigurationDatabase']) { $failures.Add('Farm: required metadata is unavailable') }
    $farmRow['CollectionStatus'] = if ($failures.Count) { 'Partial' } else { 'Collected' }
    $rows.Farms.Add([pscustomobject]$farmRow)

    foreach ($server in $servers) {
        try {
            $row = New-InventoryRow $farmId
            $row['ServerId'] = [string](Get-ObservedProperty $server 'Id')
            $row['ServerName'] = [string](Get-ObservedProperty $server 'Name')
            $row['Role'] = [string](Get-ObservedProperty $server 'Role')
            $row['Status'] = [string](Get-ObservedProperty $server 'Status')
            $row['CollectionStatus'] = if (-not $row['ServerId'] -or -not $row['ServerName'] -or -not $row['Role'] -or -not $row['Status']) { 'PropertyUnavailable' } else { 'Collected' }
            if ($row['CollectionStatus'] -ne 'Collected') { $failures.Add("ServerProperty: $($row['ServerId'])") }
            $rows.Servers.Add([pscustomobject]$row)
        } catch { $failures.Add("Server: $($_.Exception.Message)"); WriteLog -Message "Server row failed: $($_.Exception.Message)" -Level ERROR }
    }
    foreach ($service in $serviceApplications) {
        try {
            $row = New-InventoryRow $farmId
            $row['ServiceApplicationId'] = [string](Get-ObservedProperty $service 'Id')
            $row['Name'] = [string](Get-ObservedProperty $service 'Name')
            $row['TypeName'] = [string](Get-ObservedProperty $service 'TypeName')
            $row['Status'] = [string](Get-ObservedProperty $service 'Status')
            $row['ApplicationPool'] = [string](Get-ObservedProperty (Get-ObservedProperty $service 'ApplicationPool') 'Name')
            $row['CollectionStatus'] = if (-not $row['ServiceApplicationId'] -or -not $row['Name'] -or -not $row['TypeName'] -or -not $row['Status']) { 'PropertyUnavailable' } else { 'Collected' }
            if ($row['CollectionStatus'] -ne 'Collected') { $failures.Add("ServiceApplicationProperty: $($row['ServiceApplicationId'])") }
            $rows.ServiceApplications.Add([pscustomobject]$row)
        } catch { $failures.Add("ServiceApplication: $($_.Exception.Message)"); WriteLog -Message "Service application row failed: $($_.Exception.Message)" -Level ERROR }
    }
    foreach ($webApplication in $webApplications) {
        $webApplicationId = [string]$webApplication.Id
        try {
            $databases = @(Get-SPContentDatabase -WebApplication $webApplication -ErrorAction Stop)
            $row = New-InventoryRow $farmId
            $row['WebApplicationId'] = $webApplicationId
            $row['Name'] = [string](Get-ObservedProperty $webApplication 'Name')
            $row['DefaultUrl'] = [string](Get-ObservedProperty $webApplication 'Url')
            $row['ApplicationPool'] = [string](Get-ObservedProperty (Get-ObservedProperty $webApplication 'ApplicationPool') 'Name')
            $row['ContentDatabaseCount'] = $databases.Count
            $row['CollectionStatus'] = if (-not $row['WebApplicationId'] -or -not $row['Name'] -or -not $row['DefaultUrl'] -or -not $row['ApplicationPool']) { 'PropertyUnavailable' } else { 'Collected' }
            if ($row['CollectionStatus'] -ne 'Collected') { $failures.Add("WebApplicationProperty: $webApplicationId") }
            $rows.WebApplications.Add([pscustomobject]$row)
            foreach ($database in $databases) {
                try {
                    $dbRow = New-InventoryRow $farmId
                    $dbRow['WebApplicationId'] = $webApplicationId
                    $dbRow['ContentDatabaseId'] = [string](Get-ObservedProperty $database 'Id')
                    $dbRow['Name'] = [string](Get-ObservedProperty $database 'Name')
                    $dbRow['DatabaseServer'] = [string](Get-ObservedProperty $database 'Server')
                    $dbRow['Status'] = [string](Get-ObservedProperty $database 'Status')
                    $dbRow['SiteCollectionCount'] = [string](Get-ObservedProperty $database 'CurrentSiteCount')
                    $dbRow['IsReadOnly'] = [string](Get-ObservedProperty $database 'IsReadOnly')
                    $dbRow['CollectionStatus'] = if (-not $dbRow['ContentDatabaseId'] -or -not $dbRow['Name'] -or -not $dbRow['DatabaseServer'] -or -not $dbRow['Status'] -or $dbRow['SiteCollectionCount'] -eq '' -or $dbRow['IsReadOnly'] -eq '') { 'PropertyUnavailable' } else { 'Collected' }
                    if ($dbRow['CollectionStatus'] -ne 'Collected') { $failures.Add("ContentDatabaseProperty: $($dbRow['ContentDatabaseId'])") }
                    $rows.ContentDatabases.Add([pscustomobject]$dbRow)
                } catch { $failures.Add("ContentDatabase: $($_.Exception.Message)"); WriteLog -Message "Content database row failed: $($_.Exception.Message)" -Level ERROR }
            }
            foreach ($zone in @($webApplication.IisSettings.Keys)) {
                try {
                    $iis = $webApplication.IisSettings[$zone]
                    $zoneRow = New-InventoryRow $farmId
                    $zoneRow['WebApplicationId'] = $webApplicationId
                    $zoneRow['Zone'] = [string]$zone
                    $zoneRow['PublicUrl'] = [string]$webApplication.AlternateUrls.GetResponseUrl($zone)
                    $zoneRow['AuthenticationMode'] = [string](Get-ObservedProperty $iis 'AuthenticationMode')
                    $zoneRow['ClaimsAuthentication'] = [string](Get-ObservedProperty $iis 'UseClaimsAuthentication')
                    $providers = @(Get-SPAuthenticationProvider -WebApplication $webApplication -Zone $zone -ErrorAction Stop)
                    $zoneRow['AuthenticationProviders'] = @($providers | ForEach-Object { [string](Get-ObservedProperty $_ 'DisplayName') }) -join '|'
                    $zoneRow['CollectionStatus'] = if (-not $zoneRow['PublicUrl'] -or -not $zoneRow['AuthenticationMode'] -or -not $zoneRow['ClaimsAuthentication'] -or -not $zoneRow['AuthenticationProviders']) { 'PropertyUnavailable' } else { 'Collected' }
                    if ($zoneRow['CollectionStatus'] -ne 'Collected') { $failures.Add("WebApplicationZoneProperty: $webApplicationId/$zone") }
                    $rows.WebApplicationZones.Add([pscustomobject]$zoneRow)
                } catch { $failures.Add("WebApplicationZone: $($_.Exception.Message)"); WriteLog -Message "Web application zone failed: $($_.Exception.Message)" -Level ERROR }
            }
        } catch { $failures.Add("WebApplication: $($_.Exception.Message)"); WriteLog -Message "Web application failed: $($_.Exception.Message)" -Level ERROR }
    }

    if ($failures.Count -gt 0) { $rows.Farms[0].CollectionStatus = 'Partial' }
    $schemas = [ordered]@{
        Farms = @('TenantKey','FarmId','RunId','CollectedAtUtc','FarmVersion','BuildVersion','ConfigurationDatabase','ServerCount','ServiceApplicationCount','CollectionStatus')
        Servers = @('TenantKey','FarmId','RunId','CollectedAtUtc','ServerId','ServerName','Role','Status','CollectionStatus')
        ServiceApplications = @('TenantKey','FarmId','RunId','CollectedAtUtc','ServiceApplicationId','Name','TypeName','Status','ApplicationPool','CollectionStatus')
        WebApplications = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','Name','DefaultUrl','ApplicationPool','ContentDatabaseCount','CollectionStatus')
        WebApplicationZones = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','Zone','PublicUrl','AuthenticationMode','ClaimsAuthentication','AuthenticationProviders','CollectionStatus')
        ContentDatabases = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','ContentDatabaseId','Name','DatabaseServer','Status','SiteCollectionCount','IsReadOnly','CollectionStatus')
    }
    $runPaths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($kind in $schemas.Keys) {
        $name = "SharePoint_OnPrem_$kind.csv"
        $runPaths.Add((Write-RunCsv -Name $name -Rows $rows[$kind].ToArray() -Columns $schemas[$kind] -Folder $runFolder))
    }
    if ($failures.Count -gt 0) {
        WriteLog -Message ("Infrastructure coverage is incomplete: {0} failure(s). DATA-LAST was not updated." -f $failures.Count) -Level ERROR
        Complete-SmartM365ExecutionContext -Status Failed
        $script:Completed = $true
        exit 1
    }
    if ([bool](Get-InventoryConfigValue -Name 'EnableWeeklyHistory' -DefaultValue $true)) {
        $historyRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue 'WeeklyHistoryFolderPath') -Name 'WeeklyHistoryFolderPath'
        Add-SmartM365WeeklyHistory -SourceCsvPaths @($runPaths) -HistoryRootPath $historyRoot -RetentionWeeks ([int](Get-InventoryConfigValue 'WeeklyHistoryRetentionWeeks' 52)) -HistoryLabel 'SharePoint on-prem infrastructure' | Out-Null
    }
    $qualifiedCsvPaths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($path in $runPaths) {
        $latestPath = Join-Path $latestRoot (Split-Path $path -Leaf)
        Copy-SmartM365FileAtomically -SourcePath $path -DestinationPath $latestPath
        [void]$global:csvGeneratedPaths.Add($latestPath)
        $qualifiedCsvPaths.Add($path)
        $qualifiedCsvPaths.Add($latestPath)
    }
    $receiptPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    if (-not $receiptPath -or $receiptPath -notlike '*.current.json.txt') { throw 'Infrastructure source receipt qualification failed.' }
    WriteLog -Message 'Infrastructure collection and current publication completed.'
    $uploadFailures = Publish-QualifiedCsvUploads -Paths $qualifiedCsvPaths.ToArray()
    try {
        $farmVersion = [System.Net.WebUtility]::HtmlEncode([string]$rows.Farms[0].FarmVersion)
        $body = "<h2>SharePoint on-prem infrastructure inventory</h2><p>Qualified run: $($script:RunId)</p><p>Farm version: $farmVersion</p><p>SharePoint upload failures: $uploadFailures</p><table><tr><th>Servers</th><th>Service applications</th><th>Web applications</th><th>Content databases</th></tr><tr><td>$($rows.Servers.Count)</td><td>$($rows.ServiceApplications.Count)</td><td>$($rows.WebApplications.Count)</td><td>$($rows.ContentDatabases.Count)</td></tr></table>"
        $mailMarkerPath = Join-Path $outputBase 'SmartM365-SharePoint-OnPrem-Infrastructure-DailySummary.sent'
        $null = Invoke-DailySummaryMail -MarkerPath $mailMarkerPath -Force:$ForceSendDailySummary -SendAction {
            Send-InventorySummaryMail -Subject 'SharePoint on-prem infrastructure inventory' -BodyHtml $body
        }
    } catch {
        WriteLog -Message ("Daily infrastructure summary email failed: {0}" -f $_.Exception.Message) -Level ERROR
        Complete-SmartM365ExecutionContext -Status CompletedWithWarnings
        $script:Completed = $true
        exit 3
    }
    if ($uploadFailures -gt 0) {
        Complete-SmartM365ExecutionContext -Status CompletedWithWarnings
        $script:Completed = $true
        exit 3
    }
    Complete-SmartM365ExecutionContext -Status Success
    $script:Completed = $true
} catch {
    try { WriteLog -Message ("Infrastructure inventory failed: {0}" -f $_.Exception.Message) -Level ERROR } catch { Write-Error $_ }
    if (-not $script:Completed) { try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $_ -FailureStage 'InfrastructureInventory' } catch {} }
    throw
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCKWyH08wp4uvUM
# udOSk3GHE/ulqcerwiCyP+wOb44GpaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBwMTOUOhK/QSKsj1bwN6oC
# WZ0jSuF338zedxvE/JKxNjANBgkqhkiG9w0BAQEFAASCAYA2S0sIp7aaWdPDbjxV
# gyljuJXnKjguK0+6hMtyLjYF1G81dWgJJlmLWbSZm6IEdss/j2ZdueKiJ/qQm5kr
# cnm5IddxE8K/EO8D2pPj2gtnPAYaz0AKyh/NC54Sjvzg3S+Y48yHprI98YsJoSES
# a6LyS4tjhN09BfoHQNVqbw7MK33/YZd9LCwLp5dXYnO/FDL5wSfyXkKa99WmuwgV
# F4cFSj/JFTMi8JOFiFIIUh07vurb+x03v615jvsuGMKuUeoaR9L+FYNXwLW0RqRn
# LnZZXL6LtBmF3vnckcxJQTDM/x7BV+mM71cRMWnVDwc35s1kkRfVF4yhrX61e/qE
# RBnEXqdN1c+BH3ksehRPEZipoNOdOyKVDoB1uIHGWahh+jKulMfu9hvw1MVyGOJX
# QniVjJXIBDoqEB0j/T243DFzuxtpATj45/fFc1fSC4+2nVxx5ZtGwdoxgcTGDxmE
# vfyQ37ddq/WQri5jIg/9oaaQL0IuFKyNHTM3HLsYCJ5tBAc=
# SIG # End signature block
