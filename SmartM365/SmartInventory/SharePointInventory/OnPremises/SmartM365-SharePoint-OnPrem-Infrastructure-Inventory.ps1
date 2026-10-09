<#
.SYNOPSIS
    Read-only SharePoint Server 2016/2019 farm infrastructure inventory.
.VERSION
    1.0.10
.REQUIREMENTS
    Windows PowerShell 5.1 x64 on a SharePoint farm server; SharePoint Shell access.
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$OutputRoot,
    [switch]$ValidateOnly,
    [Alias('ForceSendDailySummary')][switch]$ForceMail
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ScriptName = 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory'
$script:RunId = [guid]::NewGuid().ToString('N')
$script:CollectedAtUtc = [datetime]::UtcNow.ToString('o')
$script:BuildEditionMapPath = Join-Path $PSScriptRoot 'SharePoint-OnPrem-BuildEditions.psd1'

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

function Get-FarmConfigurationDatabaseName {
    param($Farm)
    $directName = [string](Get-ObservedProperty (Get-ObservedProperty $Farm 'ConfigurationDatabase') 'Name')
    if (-not [string]::IsNullOrWhiteSpace($directName)) { return $directName }
    try {
        $configurationDatabases = @(Get-SPDatabase -ErrorAction Stop | Where-Object { [string](Get-ObservedProperty $_ 'Type') -eq 'Configuration Database' })
        if ($configurationDatabases.Count -ne 1) {
            WriteLog -Message ("Farm configuration database lookup returned {0} candidates; expected one." -f $configurationDatabases.Count) -Level WARNING
            return ''
        }
        $name = [string](Get-ObservedProperty $configurationDatabases[0] 'Name')
        if ([string]::IsNullOrWhiteSpace($name)) {
            WriteLog -Message 'Farm configuration database name is unavailable.' -Level WARNING
            return ''
        }
        WriteLog -Message 'Farm configuration database name resolved through Get-SPDatabase.' -Level INFO
        return $name
    } catch {
        WriteLog -Message ("Farm configuration database lookup failed: {0}" -f $_.Exception.Message) -Level WARNING
        return ''
    }
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

function Get-InfrastructureFarmEdition {
    param([string]$BuildVersion)
    try {
        $mapping = Import-PowerShellDataFile -Path $script:BuildEditionMapPath -ErrorAction Stop
        if ($mapping.Builds.ContainsKey($BuildVersion) -and $mapping.Builds[$BuildVersion] -in @('2016','2019')) { return [string]$mapping.Builds[$BuildVersion] }
    } catch { WriteLog -Message ("Farm edition mapping unavailable: {0}" -f $_.Exception.Message) -Level WARNING }
    return 'Unknown'
}

function Get-ContentDatabaseExtendedFields {
    param($Database)
    $result = [ordered]@{ DatabaseSizeBytes=''; DatabaseSizeStatus='Unavailable'; NeedsUpgradeIncludeChildren=''; NeedsUpgradeStatus='PropertyUnavailable' }
    # SPContentDatabase.DiskSizeRequired estimates backup space, not database size.
    # The actual byte size has no approved SharePoint object-model source in this lot.
    try {
        $upgrade = Get-ObservedProperty $Database 'NeedsUpgradeIncludeChildren'
        if ($upgrade -is [bool]) {
            $result['NeedsUpgradeIncludeChildren'] = [string]$upgrade
            $result['NeedsUpgradeStatus'] = 'Collected'
        }
    } catch { $result['NeedsUpgradeStatus'] = 'ReadFailed'; WriteLog -Message ("Content database upgrade-state read failed: {0}" -f $_.Exception.Message) -Level WARNING }
    return [pscustomobject]$result
}

function Get-QualifiedPreviousInfrastructureCounts {
    param([string]$LatestRoot, [string]$FarmId, [string]$TenantKey)
    $receiptPath = Join-Path $LatestRoot 'SmartInventory_SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.current.json.txt'
    $names = @('Farms','Servers','ServiceApplications','WebApplications','WebApplicationZones','ContentDatabases')
    try {
        $receipt = Get-Content -LiteralPath $receiptPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($receipt.Owner -ne 'SmartInventory-SourceReceipt' -or $receipt.Status -ne 'Completed' -or -not $receipt.RunId) { return $null }
        $counts = @{}
        $csvRunId = $null
        foreach ($name in $names) {
            $fileName = "SharePoint_OnPrem_$name.csv"
            $fileReceipts = @($receipt.Files | Where-Object { $_.File -eq $fileName -and $_.Status -eq 'Success' -and $_.RunId -eq $receipt.RunId })
            if ($fileReceipts.Count -ne 1) { return $null }
            $path = Join-Path $LatestRoot $fileName
            if ((Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash -ne $fileReceipts[0].SHA256) { return $null }
            $csvRows = @(Import-Csv -LiteralPath $path -Delimiter ';' -ErrorAction Stop)
            if ($csvRows.Count -ne [int]$fileReceipts[0].Rows) { return $null }
            foreach ($csvRow in $csvRows) {
                if (-not $csvRow.RunId -or $csvRow.FarmId -ne $FarmId -or $csvRow.TenantKey -ne $TenantKey) { return $null }
                if ($null -eq $csvRunId) { $csvRunId = [string]$csvRow.RunId }
                elseif ($csvRow.RunId -ne $csvRunId) { return $null }
            }
            $counts[$name] = if ($name -eq 'Servers') { @($csvRows | Where-Object { $_.Role -ne 'Invalid' }).Count } else { $csvRows.Count }
        }
        if ($counts.Farms -ne 1) { return $null }
        return $counts
    } catch {
        WriteLog -Message ("Previous qualified infrastructure run is unavailable for comparison: {0}" -f $_.Exception.Message) -Level INFO
        return $null
    }
}

function New-InfrastructureMailHtml {
    param([string]$Status, $Rows, [string[]]$Failures = @(), [int]$UploadFailures = 0, $PreviousCounts = $null, [string]$TenantName = '', [Nullable[timespan]]$Duration = $null, [string]$HostName = $env:COMPUTERNAME)
    $kinds = @('Farms','Servers','ServiceApplications','WebApplications','WebApplicationZones','ContentDatabases')
    $counts = @{}
    foreach ($kind in $kinds) { $counts[$kind] = if ($Rows -and $Rows.ContainsKey($kind)) { [int]$Rows[$kind].Count } else { $null } }
    $externalServerCount = if ($Rows -and $Rows.ContainsKey('Servers')) { @($Rows.Servers | Where-Object { $_.Role -eq 'Invalid' }).Count } else { 0 }
    if ($null -ne $counts.Servers) { $counts.Servers -= $externalServerCount }
    $farmBuild = if ($Rows -and $counts.Farms -gt 0) { [string]$Rows.Farms[0].BuildVersion } else { '' }
    $edition = Get-InfrastructureFarmEdition -BuildVersion $farmBuild
    $cardSpecs = @(
        @('Servers','Servers','#f8fafc','#dbe3ef','#0f172a'),
        @('Service applications','ServiceApplications','#eff6ff','#bfdbfe','#1d4ed8'),
        @('Web applications','WebApplications','#f0fdf4','#bbf7d0','#166534'),
        @('Content databases','ContentDatabases','#faf5ff','#e9d5ff','#7e22ce')
    )
    $cards = New-Object 'System.Collections.Generic.List[object]'
    foreach ($spec in $cardSpecs) {
        $failurePrefix = switch ($spec[1]) { 'Servers' {'Servers:'} 'ServiceApplications' {'ServiceApplications:'} 'WebApplications' {'WebApplication:'} 'ContentDatabases' {'ContentDatabase:'} }
        $areaFailed = @($Failures | Where-Object { $_.StartsWith($failurePrefix, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
        $value = if ($null -eq $counts[$spec[1]] -or ($counts[$spec[1]] -eq 0 -and $areaFailed)) { 'n/a' } else { [string]$counts[$spec[1]] }
        if ($null -ne $PreviousCounts -and $value -ne 'n/a' -and $PreviousCounts.ContainsKey($spec[1])) {
            $delta = [int]$counts[$spec[1]] - [int]$PreviousCounts[$spec[1]]
            $value += (' ({0}{1})' -f $(if ($delta -ge 0) { '+' } else { '' }), $delta)
        }
        $detail = if ($spec[1] -eq 'Servers') { if ($areaFailed) { 'external count unavailable' } else { "$externalServerCount external" } } elseif ($Status -eq 'Qualified') { 'inventory rows' } else { 'observed rows' }
        $cards.Add([pscustomobject]@{ Label=$spec[0]; Value=$value; Detail=$detail; Background=$spec[2]; Border=$spec[3]; Accent=$spec[4]; Span=1 })
    }
    $cards.Add([pscustomobject]@{Label='Farm edition';Value=$edition;Detail=$(if ($farmBuild) { $farmBuild } else { 'build unavailable' });Background='#f0fdfa';Border='#99f6e4';Accent='#0f766e';Span=2})
    $uploadValue = if ($Status -eq 'Qualified') { [string]$UploadFailures } else { 'n/a' }
    $cards.Add([pscustomobject]@{Label='Upload failures';Value=$uploadValue;Detail=$(if ($Status -eq 'Qualified') { 'SharePoint upload' } else { 'not attempted' });Background='#eef2ff';Border='#c7d2fe';Accent='#4338ca';Span=2})
    $sections = New-Object 'System.Collections.Generic.List[object]'
    $sections.Add([pscustomobject]@{Title='Summary';Html=(New-SmartM365EmailKpiGridHtml -Cards $cards.ToArray() -RowSizes @(4,2))})

    $alerts = New-Object 'System.Collections.Generic.List[object]'
    if ($Rows) {
        foreach ($server in ($Rows.Servers | Where-Object { $_.Role -ne 'Invalid' })) { if ($server.Status -and $server.Status -ne 'Online') { $alerts.Add(@('Server not online',[string]$server.ServerName,[string]$server.Status)) } }
        foreach ($service in $Rows.ServiceApplications) { if ($service.Status -and $service.Status -notin @('Online','Started')) { $alerts.Add(@('Service application not started',[string]$service.Name,[string]$service.Status)) } }
        foreach ($database in $Rows.ContentDatabases) {
            if ($database.Status -and $database.Status -ne 'Online') { $alerts.Add(@('Content database not online',[string]$database.Name,[string]$database.Status)) }
            if ($database.IsReadOnly -eq 'True') { $alerts.Add(@('Content database read-only',[string]$database.Name,'True')) }
            if ((Get-ObservedProperty $database 'NeedsUpgradeIncludeChildren') -eq 'True' -and (Get-ObservedProperty $database 'NeedsUpgradeStatus') -eq 'Collected') { $alerts.Add(@('Content database needs upgrade',[string]$database.Name,'True')) }
        }
        $unverifiedUpgradeCount = @($Rows.ContentDatabases | Where-Object { (Get-ObservedProperty $_ 'NeedsUpgradeStatus') -notin @('Collected','') }).Count
        if ($unverifiedUpgradeCount -gt 0) { $alerts.Add(@('Upgrade state unavailable',"$unverifiedUpgradeCount content databases",'Check NeedsUpgradeStatus in CSV')) }
    }
    foreach ($failure in $Failures) { $alerts.Add(@('Collection issue',[string]$failure,'')) }
    if ($alerts.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Alerts';Html=(New-SmartM365EmailTableHtml -Headers @('Issue','Object','Status') -Rows $alerts.ToArray())}) }
    if ($Rows) {
        $serverRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($server in ($Rows.Servers | Where-Object { $_.Role -ne 'Invalid' } | Sort-Object ServerName)) {
            $serverRows.Add(@([string]$server.ServerName,[string]$server.Role,[string]$server.Status))
        }
        if ($serverRows.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Servers';Html=(New-SmartM365EmailTableHtml -Headers @('Name','Role','Status') -Rows $serverRows.ToArray())}) }
        $webRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($web in ($Rows.WebApplications | Sort-Object DefaultUrl)) {
            $zones = @($Rows.WebApplicationZones | Where-Object { $_.WebApplicationId -eq $web.WebApplicationId } | Sort-Object Zone | ForEach-Object { '{0}: {1}' -f $_.Zone,$_.AuthenticationMode }) -join '; '
            $webRows.Add(@([string]$web.DefaultUrl,[string]$web.ApplicationPool,[string]$web.ContentDatabaseCount,[string]$zones))
        }
        if ($webRows.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Web applications';Html=(New-SmartM365EmailTableHtml -Headers @('URL','Application pool','Content databases','Authentication by zone') -Rows $webRows.ToArray() -NoWrapColumns @(0,1))}) }
        $sized = @($Rows.ContentDatabases | Where-Object { (Get-ObservedProperty $_ 'DatabaseSizeStatus') -eq 'Collected' -and (Get-ObservedProperty $_ 'DatabaseSizeBytes') -match '^\d+$' } | Sort-Object @{Expression={[decimal]$_.DatabaseSizeBytes};Descending=$true} | Select-Object -First 5)
        if ($sized.Count -gt 0) {
            $sizeRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($database in $sized) { $sizeRows.Add(@([string]$database.Name,[string]$database.DatabaseServer,[string]$database.DatabaseSizeBytes)) }
            $sections.Add([pscustomobject]@{Title='Top 5 largest content databases';Html=(New-SmartM365EmailTableHtml -Headers @('Name','Server','Size (bytes)') -Rows $sizeRows.ToArray())})
        }
        if ($Status -in @('Qualified','Preview')) {
            $fileRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($kind in $kinds) { $fileRows.Add(@("SharePoint_OnPrem_$kind.csv",[string]$Rows[$kind].Count)) }
            $sections.Add([pscustomobject]@{Title='Files';Html=(New-SmartM365EmailTableHtml -Headers @('CSV file','Rows') -Rows $fileRows.ToArray())})
        }
    }
    $severity = switch ($Status) { 'Qualified' {'Success'} 'Partial' {'Warning'} 'Preview' {'Info'} default {'Error'} }
    $durationText = if ($null -ne $Duration) { ([timespan]$Duration).ToString('hh\:mm\:ss') } else { '' }
    return New-SmartM365EmailBody -Title 'SharePoint infrastructure inventory summary' -Category 'SmartM365 SharePoint OnPrem' -Severity $severity -Tenant $TenantName -HostName $HostName -StatusBadge $Status.ToUpperInvariant() -Duration $durationText -Message 'SharePoint on-premises infrastructure inventory summary generated from the latest SmartM365 CSV outputs.' -Sections $sections.ToArray() -Footer 'This automated message was generated by SmartM365 from SharePoint on-premises infrastructure inventory data.'
}

function Send-InfrastructureRunMail {
    param([string]$Status, $Rows, [string[]]$Failures = @(), [int]$UploadFailures = 0, $PreviousCounts = $null)
    try {
        $markerName = if ($Status -eq 'Qualified') { 'SmartM365-SharePoint-OnPrem-Infrastructure-DailySummary.sent' } else { 'SmartM365-SharePoint-OnPrem-Infrastructure-IssueDailySummary.sent' }
        $markerPath = Join-Path $outputBase $markerName
        $duration = (Get-Date) - $script:StartedAt
        $body = New-InfrastructureMailHtml -Status $Status -Rows $Rows -Failures $Failures -UploadFailures $UploadFailures -PreviousCounts $PreviousCounts -TenantName $Tenant -Duration $duration
        $null = Invoke-DailySummaryMail -MarkerPath $markerPath -Force:$ForceMail -SendAction { Send-InventorySummaryMail -Subject 'SharePoint on-prem infrastructure inventory' -BodyHtml $body }
    } catch { WriteLog -Message ("Infrastructure summary email failed without affecting inventory qualification: {0}" -f $_.Exception.Message) -Level WARNING }
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
Import-Module -Name $coreManifest -MinimumVersion '1.0.54' -ErrorAction Stop
$outputBase = if ($PSBoundParameters.ContainsKey('OutputRoot')) { Resolve-InventoryConfigTokens -Value $OutputRoot -Name 'OutputRoot' } else { Get-InventoryConfigValue -Name 'OutputRoot' }
$outputBase = Assert-InventoryPath -Path $outputBase -Name 'OutputRoot'
$latestRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue -Name 'LatestCsvFolderPath') -Name 'LatestCsvFolderPath'
$runFolder = Join-Path $outputBase ((Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:RunId.Substring(0,8))
$script:StartedAt = Get-Date
$script:ReceiptStarted = $false
$script:Completed = $false
$script:FailureHandled = $false
$script:CurrentStage = 'Initialization'
$script:CurrentStagePath = ''
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
    $previousCounts = Get-QualifiedPreviousInfrastructureCounts -LatestRoot $latestRoot -FarmId $farmId -TenantKey $global:SmartM365TenantKey
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
    $farmRow['ConfigurationDatabase'] = Get-FarmConfigurationDatabaseName -Farm $farm
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
                    $extended = Get-ContentDatabaseExtendedFields -Database $database
                    $dbRow['DatabaseSizeBytes'] = $extended.DatabaseSizeBytes
                    $dbRow['DatabaseSizeStatus'] = $extended.DatabaseSizeStatus
                    $dbRow['NeedsUpgradeIncludeChildren'] = $extended.NeedsUpgradeIncludeChildren
                    $dbRow['NeedsUpgradeStatus'] = $extended.NeedsUpgradeStatus
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
        ContentDatabases = @('TenantKey','FarmId','RunId','CollectedAtUtc','WebApplicationId','ContentDatabaseId','Name','DatabaseServer','Status','SiteCollectionCount','IsReadOnly','CollectionStatus','DatabaseSizeBytes','DatabaseSizeStatus','NeedsUpgradeIncludeChildren','NeedsUpgradeStatus')
    }
    $runPaths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($kind in $schemas.Keys) {
        $name = "SharePoint_OnPrem_$kind.csv"
        $runPaths.Add((Write-RunCsv -Name $name -Rows $rows[$kind].ToArray() -Columns $schemas[$kind] -Folder $runFolder))
    }
    if ($failures.Count -gt 0) {
        foreach ($failure in $failures) { WriteLog -Message ("Infrastructure coverage detail: {0}" -f $failure) -Level WARNING }
        WriteLog -Message ("Infrastructure coverage is incomplete: {0} failure(s). DATA-LAST was not updated." -f $failures.Count) -Level ERROR
        Send-InfrastructureRunMail -Status Partial -Rows $rows -Failures $failures.ToArray()
        Complete-SmartM365ExecutionContext -Status Failed
        $script:Completed = $true
        exit 1
    }
    if ([bool](Get-InventoryConfigValue -Name 'EnableWeeklyHistory' -DefaultValue $true)) {
        $historyRoot = Assert-InventoryPath -Path (Get-InventoryConfigValue 'WeeklyHistoryFolderPath') -Name 'WeeklyHistoryFolderPath'
        $script:CurrentStage = 'WeeklyHistory'
        $script:CurrentStagePath = $historyRoot
        WriteLog -Message ("Saving weekly infrastructure history in {0}." -f $historyRoot)
        Add-SmartM365WeeklyHistory -SourceCsvPaths @($runPaths) -HistoryRootPath $historyRoot -RetentionWeeks ([int](Get-InventoryConfigValue 'WeeklyHistoryRetentionWeeks' 52)) -HistoryLabel 'SharePoint on-prem infrastructure' | Out-Null
    }
    $script:CurrentStage = 'CurrentCsvPublication'
    $qualifiedCsvPaths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($path in $runPaths) {
        $latestPath = Join-Path $latestRoot (Split-Path $path -Leaf)
        $script:CurrentStagePath = $latestPath
        Copy-SmartM365FileAtomically -SourcePath $path -DestinationPath $latestPath
        [void]$global:csvGeneratedPaths.Add($latestPath)
        $qualifiedCsvPaths.Add($path)
        $qualifiedCsvPaths.Add($latestPath)
    }
    $script:CurrentStage = 'SourceReceipt'
    $script:CurrentStagePath = $latestRoot
    $receiptPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    if (-not $receiptPath -or $receiptPath -notlike '*.current.json.txt') { throw 'Infrastructure source receipt qualification failed.' }
    WriteLog -Message 'Infrastructure collection and current publication completed.'
    $uploadFailures = Publish-QualifiedCsvUploads -Paths $qualifiedCsvPaths.ToArray()
    Send-InfrastructureRunMail -Status Qualified -Rows $rows -UploadFailures $uploadFailures -PreviousCounts $previousCounts
    if ($uploadFailures -gt 0) {
        Complete-SmartM365ExecutionContext -Status CompletedWithWarnings
        $script:Completed = $true
        exit 3
    }
    Complete-SmartM365ExecutionContext -Status Success
    $script:Completed = $true
} catch {
    $script:FailureHandled = $true
    $failure = $_
    if (-not $ValidateOnly -and -not $script:Completed) {
        $mailRows = if (Get-Variable -Name rows -ErrorAction SilentlyContinue) { $rows } else { $null }
        Send-InfrastructureRunMail -Status Failed -Rows $mailRows -Failures @($failure.Exception.Message)
    }
    try {
        WriteLog -Message ("Infrastructure inventory failed during {0} at {1}: {2}" -f $script:CurrentStage, $script:CurrentStagePath, $failure.Exception.Message) -Level ERROR
        if ($failure.InvocationInfo -and $failure.InvocationInfo.PositionMessage) {
            WriteLog -Message ("Infrastructure failure location: {0}" -f ($failure.InvocationInfo.PositionMessage.Trim() -replace '\r?\n', ' | ')) -Level ERROR
        }
        if ($failure.ScriptStackTrace) {
            WriteLog -Message ("Infrastructure failure stack: {0}" -f ($failure.ScriptStackTrace -replace '\r?\n', ' | ')) -Level ERROR
        }
    } catch { Write-Error $failure }
    if (-not $script:Completed) { try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $failure -FailureStage 'InfrastructureInventory' } catch {} }
    throw
} finally {
    if (-not $script:Completed -and -not $script:FailureHandled) {
        $global:EnableSharePointUpload = $false
        try { WriteLog -Message ("Infrastructure inventory interrupted by Ctrl+C. RunId={0}; no further publication or mail will be attempted." -f $script:RunId) -Level WARNING }
        catch { [Console]::Error.WriteLine('Infrastructure inventory interrupted; the run log could not be updated.') }
        try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage 'Interrupted' }
        catch { [Console]::Error.WriteLine('Infrastructure inventory interrupted; execution cleanup failed: ' + $_.Exception.Message) }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA5n/BCtsyzT5hM
# Nk3upIzeigz2vrJC1oTp6bSdVnrJ5qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINLu9Ts1FE1cCY4AsvX7JFxcwtRJxn63DiYiq7YbTnonMA0GCSqG
# SIb3DQEBAQUABIIBgGIH6zpFM19ir2hv0EtdP5pCaOAUJYbhm3cRYgkgJLBssK6I
# cDGS2SWOv67AL+Emxk4vwI/ghiqimrKG6D/X+YHu10PxqfqZytf5BlCoYE9y6Ka0
# pqTmms8IaVkc84mwgt/cQLoOJ12dOAlWy2w08b9W0rsifJ3Yancr9LY5P8NApbu1
# Pi5rftjEmZxJxXf4bWDx0PxDb4iiNTVyYuViyU2riBXkXBCLoYQGZn360b4wiA1/
# PdVCa3cfxLGC/d0JafV5lrxjn47hqpbPgzpV0TrLQrHpnpoAVcp7XG2JwFRjqD2o
# KcqH930CKfWMcLVSVrqUabP8MrXdatMvp74qyknenbvIYQn6TzrfK9XdjdRvTH3i
# TvXpvWCl9jvzw3Ei1wm6Et6Eq4JBKt8iT7Pti2y8f/3Jv+JIJEv9mX4m+Z+sxoff
# ATylCKLLsdVtrRB2R/pB2n8gyIlK7grMPCJo1GJAgNWFc74ztmJ/9QwNPQQ/1SbF
# ny5sxxrdlLBrPGBY4qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDkyMjMx
# NTdaMC8GCSqGSIb3DQEJBDEiBCBXOtwf2CsMHU1gITnHgZTp57wF/OJlEJMzi6zN
# kCw0NDANBgkqhkiG9w0BAQEFAASCAgBYdLxd8nXtaeCgnkjm/smGKSQJoPYCZLT8
# 5pJq7FNtUol885+fsAtARH1AFlghuDspRoEXvotjjXbI69OJC0QtdxIejsdQ6Y7O
# nrh6KusZZpNn2ooeFHd28yW+all37UL8CEujfEZa7hik91+NYyhyNC+npJm4GnoQ
# UISP8+9ugOGEZ4rtlZJBSfb4zjlxmUeyTP27gD66Kd9Yel3+xJ3MKHhGX8JuIEO8
# t/HOt9wG+6C3GoVr3jn9HtFtPlky3uFQSVJJHl/Hsk2WZazGUV1Rx/u2F1ampJIB
# 8sPHyyURsHLVXcbiYKamLkIskRTF6t/oI0hiOiqtNpM1y6wxd2koWci5HJBAkY9B
# 7+qDFd3GOxMGI2GBhDstN6j9zAyNJrmKrcaXTnfLQwkEVKFytkia0qNmyso/SHRg
# IQTXoQTTuKPShIzUyeig7g4082wIXNIUqID53iNvaTK2OzvO38zdWPLZUzpkhm/X
# +2ETaWtAoob46s/KiMvD9Y9DY2MvRpRfyu08ba6tjmA8pCwkxRGfgaD6eyEdMjy/
# Jkb/p0GRHiBnhZKcrw/BC69xdzSaD2+Mgj/F8UrYejlei6C847WCUL1A1srXFJKZ
# 4k6+PGXVnZydA1lN94VVXEiLmycDqzSVhsRLilCiVPXDC6V/Zvo9YcKolGMrX16t
# l/Tp+ql7zA==
# SIG # End signature block
