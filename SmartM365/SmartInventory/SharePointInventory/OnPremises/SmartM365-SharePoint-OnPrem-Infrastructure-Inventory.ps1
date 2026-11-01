<#
.SYNOPSIS
    Read-only SharePoint Server 2016/2019 farm infrastructure inventory.
.VERSION
    1.0.16
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
$localConfig = Read-SmartM365ScriptConfiguration -Path (Join-Path $PSScriptRoot "$script:ScriptName.local.json") -AsHashtable
$merged = ConvertTo-SmartM365Hashtable -InputObject (Merge-SmartM365Configuration -EffectiveConfig $script:EffectiveConfig -LocalConfig $localConfig -Mode Evidence)
if ($null -eq $merged['EnableSharePointUpload']) { $merged['EnableSharePointUpload'] = $true }
$script:EffectiveConfig = [pscustomobject]$merged
$global:SmartM365GlobalConfig = $script:EffectiveConfig
$global:EnableSharePointUpload = [bool]$merged['EnableSharePointUpload'] -and -not $ValidateOnly

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

function Get-SharePointServerCount {
    param([object[]]$Servers)
    return @($Servers | Where-Object { [string](Get-ObservedProperty $_ 'Role') -ne 'Invalid' }).Count
}

function Get-WebApplicationZonePublicUrl {
    param($WebApplication, $Zone)
    try {
        $responseUri = $WebApplication.GetResponseUri($Zone)
        if ($null -eq $responseUri) { return '' }
        $uri = [uri]$responseUri
        if (-not $uri.IsAbsoluteUri -or $uri.Scheme -notin @('http','https')) { return '' }
        return $uri.AbsoluteUri
    } catch {
        WriteLog -Message ("Web application zone response URI is unavailable for {0}/{1}: {2}" -f $WebApplication.Id, $Zone, $_.Exception.Message) -Level WARNING
        return ''
    }
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
$runFolder = Join-Path $outputBase ((Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:RunId.Substring(0,8))
$script:StartedAt = Get-Date
$script:ReceiptStarted = $false
$script:Completed = $false
$script:FailureHandled = $false
$script:CurrentStage = 'Initialization'
$script:CurrentStagePath = ''
try {
    $null = InitializeScriptEnvironment -OutputPath $runFolder -LogFileName $script:ScriptName -CallerScriptPath $PSCommandPath
    $global:SmartM365CollectorUploadPolicies=Resolve-SmartM365CollectorUploadPolicies -LocalConfig $localConfig -GlobalConfig $script:EffectiveConfig -LegacyEnabled ([bool]$merged['EnableSharePointUpload']) -ExternalActionsAllowed (-not $ValidateOnly)
    $global:SmartM365CollectorUploadRoots=[pscustomobject]@{DataAllRoot=$outputBase;DataLastRoot=$latestRoot;LogAllRoot=$global:LogPath}
    Assert-SmartM365CollectorUploadRoots -DataAllRoot $outputBase -DataLastRoot $latestRoot -LogAllRoot $global:LogPath -Policies $global:SmartM365CollectorUploadPolicies
    $global:EnableSharePointUpload=@($global:SmartM365CollectorUploadPolicies.Values | Where-Object Enabled).Count -gt 0
    WriteLog -Message ("SharePoint infrastructure inventory starting. SharePoint upload enabled for this run: {0}." -f $global:EnableSharePointUpload)
    if ($ForceMail) { WriteLog -Message 'ForceMail is retained for command compatibility; the combined summary email is sent by the Content collector.' -Level INFO }
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

    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $latestRoot -ScopeParameters @{ IncludedWebApplicationUrls=@((Get-InventoryConfigValue 'IncludedWebApplicationUrls' @())); ExcludedWebApplicationUrls=@((Get-InventoryConfigValue 'ExcludedWebApplicationUrls' @())) }
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
    $farmRow['ConfigurationDatabase'] = Get-FarmConfigurationDatabaseName -Farm $farm
    $farmRow['ServerCount'] = if (@($failures | Where-Object { $_ -like 'Servers:*' }).Count) { '' } else { Get-SharePointServerCount -Servers $servers }
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
                    $zoneRow['PublicUrl'] = Get-WebApplicationZonePublicUrl -WebApplication $webApplication -Zone $zone
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgdDhIWYJYlLfZ
# xnw5ZlGp8xEB8l0MPYTjz8BbIiQmYqCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgbnzayvQx1wsR
# 803ntQsRJdDkwewX/+CwzQ1n+fSs1jAwDQYJKoZIhvcNAQEBBQAEggGAsOXUuFDm
# p5lWKFpm5U/SLNz3UEYmK2lbX5LxAFZTksmPVlfmEErfYAfEMUbCf3w8xIBs173b
# 1QzmJehgGGFbtNE9j2MVfMiZsxVGwkHwhkbhwzXrLbgYh0CWkNYr9zvQ6Bpy4Tjt
# uhQbMscHH7Fci6upOToI8kGClLyT4neQgJAxLfDo9ePJXsZJ2viRvva1+r6jAQbR
# nTi8s20A376HhyMag9Bm3dN4RkDVhVDFZ3oX3FO6mlwvI2798uxDTSiHIcM3G6bu
# O2HBzh9mNVW8ZCCbmBzEOBlob/DMCWSgvmdzP53ejPbigllb3t9wfD6Y3bzB5b9g
# Lupgl8MfRQmxywc9xJ8R+kew3yeoqd+zqsTLKLGZ5jybg8ydizQWuplBPVVh7qml
# eVK40UvneERDmiQEfS9EIkZESfnAwwHVcuGTbP9BR7qZ+WYTEGoSBAcN/ffwnj1G
# +zyfYtS9ZU6vX7pq5fmlOfif/pVTGTMnoVhJc85zSoXsaj4jkk5gMnOxoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTAxMTY1OTIzWjA/BgkqhkiG9w0BCQQxMgQwvGNjUAoNduGh61pbYvXCii3b
# tbeNKHri+YmewuH27f9YsvSG9Qyujknb05dcnHntMA0GCSqGSIb3DQEBAQUABIIC
# ABtmUegIcnUJFt1xXC7QQJr5IOKHYXtdYkbRVhfpIeAE8gFTEQvg2mcP6W+KGdkJ
# Q/ssiN+s9Lvx8RZY3U1IzU2a//mkTj2h0WV2j6h/wyE7jDKugGWDTE9UMF/gKC/o
# dFRsSIgj059hj8+5bIbfHqIOSKvLZXe2PwY55h8n4RWP+WT6OoeEjWDpP8XVfKWz
# 66v1PJ7azOHf8KK9E1J3E4tirWtvW7AFzYqDuK6uzYbbSLFKYVMJ6W2FeF1/UdUh
# c9rZHUSA+2ShQgQkd4eK2IbGb8ECwHKkbWABhQGvQQEAe1RVvmOQl6b/5RVtaN4C
# 8HMQzXMp0jPARK65+cZ3sXqqAOFvfFoOxUntm2Ttt+1DprxqssF4CGmbFNrz7fDw
# lJEBbI83FdSFf6tFyapTYE/VVQmUlO+1LfDiE2qfYuXOGiQ/Ob630Wl6hSnS7B3B
# +i7/bfUsDay4TmnEqgKXXQmkXv8dzLBYa0USeBg0Z7t0SpfgWp5Ulpx4384Tu7q8
# Wkoz+uFrCSx9XeVOck6EDA+Xatb9VeEyUkCGgPVEYwpZkMNpJj0tUZLRXqU+s/aI
# EGokaDO//pIV0kXkK4u4l4ehwEdxm0RxRq3XzhxdJ9OvB4lj2kGJ2BwqvS9vRym7
# PmtGVovNgZabObPrYCMtnl6wmI0Ec93JqWmWQCq8EYoc
# SIG # End signature block
