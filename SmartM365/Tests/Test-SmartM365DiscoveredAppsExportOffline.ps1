<#
.SYNOPSIS
Offline regression tests for the Discovered Apps AppInvRawData relation source.
.DESCRIPTION
Loads the actual collector helpers through AST. The export job, the export file,
Graph relation pages and log sinks are synthetic; no authentication, tenant config
or CSV publication. Checks strict export parsing, catalog filtering, the Graph
verification sample, byte-identical relation rows and fallback triggers.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.99' -Force -Global -WarningAction SilentlyContinue
$sourcePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartInventory/M365Inventory/IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
foreach ($name in @('Get-ShortGraphErrorMessage', 'Get-DiscoveredAppsGraphHttpErrorMessage', 'Get-GraphRetryDelaySeconds',
    'Invoke-GraphPagedRequest', 'Get-DiscoveredAppsValidatedDeviceIds', 'Get-DiscoveredAppsStableDeviceIds',
    'Get-DiscoveredAppsAppDeviceCount', 'Get-DiscoveredAppsCsvDataRowCount', 'Write-DiscoveredAppsCsvRows',
    'Invoke-DiscoveredAppsGraphSingleRequest', 'Get-DiscoveredAppsRelationExportCsv', 'Read-DiscoveredAppsRelationExportCsv',
    'Select-DiscoveredAppsRandomItems', 'Test-DiscoveredAppsRelationExportSample', 'Write-DiscoveredAppsRelationExportRows',
    'Invoke-DiscoveredAppsRelationExportSource')) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if ($null -eq $node) { throw "Missing collector helper: $name" }
    Invoke-Expression $node.Extent.Text
}
$script:GraphMaxRetryAttempts = 3
$script:GraphRetryDefaultSeconds = 1
$script:GraphRetryMaxSeconds = 1
$script:Stat_GraphCalls = 0
$script:Stat_ThrottleRetries = 0
$script:testCount = 0
$script:logs = [Collections.Generic.List[string]]::new()

function WriteLog { param([string]$Message, [string]$Level) $script:logs.Add("$Level|$Message") }
function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) }
function Add-SmartM365TenantKey {
    process { [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = [string]$_.AppId; DeviceId = [string]$_.DeviceId } }
}
function Check {
    param([bool]$Condition, [string]$Message)
    $script:testCount++
    if (-not $Condition) { throw "FAIL: $Message" }
}
function Assert-Throws {
    param([scriptblock]$Script, [string]$Pattern, [string]$Message)
    $caught = $null
    try { & $Script } catch { $caught = $_.Exception.Message }
    Check ($null -ne $caught -and $caught -match $Pattern) "$Message (got: $caught)"
}

# Synthetic Graph: relation truth per app, export job state, scripted transient failures.
function Reset-Graph {
    $script:graphDevices = @{}
    $script:jobPolls = 0
    $script:jobFinalStatus = 'completed'
    $script:failQueue = [Collections.Generic.Queue[int]]::new()
    $script:requests = [Collections.Generic.List[string]]::new()
    $script:exportText = ''
    $script:exportExtraFile = $false
    $script:downloadCount = 0
}
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param($Method, $Uri, $OutputType, [switch]$SkipHttpErrorCheck, $StatusCodeVariable, $ResponseHeadersVariable, $Body, $ContentType)
    if ($ResponseHeadersVariable) { Set-Variable -Name $ResponseHeadersVariable -Value @{} -Scope 1 }
    $script:requests.Add("$Method $Uri")
    $status = 200
    $result = $null
    if ($script:failQueue.Count -gt 0) {
        $status = $script:failQueue.Dequeue()
        $result = [pscustomobject]@{ error = [pscustomobject]@{ message = "synthetic $status" } }
    }
    elseif ($Method -eq 'POST' -and $Uri -like '*/reports/exportJobs') {
        $request = $Body | ConvertFrom-Json
        Check ($request.reportName -eq 'AppInvRawData') 'Export job must request AppInvRawData.'
        Check ((@($request.select) -join ',') -eq 'ApplicationKey,DeviceId') 'Export job must select only ApplicationKey and DeviceId.'
        $result = [pscustomobject]@{ id = 'AppInvRawData_job1'; status = 'notStarted'; url = $null; requestDateTime = '2026-11-10T00:00:00Z' }
    }
    elseif ($Uri -like '*/reports/exportJobs/*') {
        $script:jobPolls++
        $result = if ($script:jobPolls -lt 2) { [pscustomobject]@{ id = 'AppInvRawData_job1'; status = 'inProgress'; url = $null } }
        elseif ($script:jobFinalStatus -eq 'completed') { [pscustomobject]@{ id = 'AppInvRawData_job1'; status = 'completed'; url = 'https://blob.example.invalid/export.zip?sig=x' } }
        else { [pscustomobject]@{ id = 'AppInvRawData_job1'; status = $script:jobFinalStatus; url = $null } }
    }
    elseif ($Uri -match '/detectedApps/([^/?]+)/managedDevices') {
        $appId = $Matches[1]
        $ids = if ($script:graphDevices.ContainsKey($appId)) { @($script:graphDevices[$appId]) } else { @() }
        $result = [pscustomobject]@{ value = @($ids | ForEach-Object { [pscustomobject]@{ id = $_ } }) }
    }
    elseif ($Uri -match '/detectedApps/([^/?]+)\?') {
        $appId = [uri]::UnescapeDataString($Matches[1])
        if ($script:graphDevices.ContainsKey($appId)) { $result = [pscustomobject]@{ id = $appId } } else { $status = 404; $result = [pscustomobject]@{ error = [pscustomobject]@{ message = 'not found' } } }
    }
    else { throw "Unexpected synthetic Graph request: $Method $Uri" }
    Set-Variable -Name $StatusCodeVariable -Value $status -Scope 1
    return $result
}
function Invoke-SmartM365ExportFileDownload {
    param([string]$Url, [string]$OutFile, [string]$Operation)
    $script:downloadCount++
    $staging = Join-Path ([IO.Path]::GetDirectoryName($OutFile)) 'staging'
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $staging 'AppInvRawData_job1.csv'), $script:exportText, [Text.UTF8Encoding]::new($true))
    if ($script:exportExtraFile) { [IO.File]::WriteAllText((Join-Path $staging 'extra.csv'), 'x') }
    Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $OutFile -Force
    Remove-Item -LiteralPath $staging -Recurse -Force
    return $OutFile
}
function New-App { param([string]$Id, [int]$Count) [pscustomobject]@{ id = $Id; displayName = $Id; version = '1'; publisher = 'p'; platform = 'windows'; deviceCount = $Count } }
function New-Guids { param([int]$Count) return ,([string[]]@(1..$Count | ForEach-Object { [guid]::NewGuid().ToString() })) }
function New-ExportText {
    param([hashtable]$Pairs, [string]$Header = '"ApplicationKey","DeviceId"', [switch]$DeviceFirst)
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add($Header)
    foreach ($key in $Pairs.Keys) {
        foreach ($device in @($Pairs[$key])) {
            if ($DeviceFirst) { $lines.Add(('"{0}","{1}"' -f $device, $key)) } else { $lines.Add(('"{0}","{1}"' -f $key, $device)) }
        }
    }
    ($lines -join "`r`n") + "`r`n"
}

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DiscoveredAppsExportTest-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    # --- Strict export parsing ------------------------------------------------
    $d1 = New-Guids 3; $d2 = New-Guids 2; $dx = New-Guids 2
    $csv = Join-Path $testRoot 'export.csv'
    $text = New-ExportText -Pairs ([ordered]@{ 'APP-A' = $d1; 'app-b' = $d2; 'app-unknown' = $dx })
    [IO.File]::WriteAllText($csv, $text, [Text.UTF8Encoding]::new($true))
    $read = Read-DiscoveredAppsRelationExportCsv -Path $csv -CatalogAppIds @('app-a', 'app-b', 'app-c')
    Check ($read.DataRows -eq 7 -and $read.AcceptedRows -eq 5 -and $read.UnknownKeyRows -eq 2) 'Export row counters are wrong.'
    Check ($read.DevicesByAppId['app-a'].Count -eq 3 -and $read.DevicesByAppId.ContainsKey('APP-A')) 'Catalog keys must match case-insensitively.'
    Check (@($read.DevicesByAppId.Keys) -ccontains 'app-a') 'Device lists must be keyed by the canonical catalog id.'
    Check ($read.DevicesByAppId['app-c'].Count -eq 0) 'Catalog apps absent from the export must have an empty list.'
    Check ($read.UnknownKeys.Count -eq 1 -and $read.UnknownKeys[0] -eq 'app-unknown') 'Non-catalog keys must be reported.'

    [IO.File]::WriteAllText($csv, (New-ExportText -Pairs ([ordered]@{ 'app-a' = $d1 }) -Header '"DeviceId","ApplicationKey"' -DeviceFirst), [Text.UTF8Encoding]::new($false))
    Check ((Read-DiscoveredAppsRelationExportCsv -Path $csv -CatalogAppIds @('app-a')).AcceptedRows -eq 3) 'Column order must follow the header.'
    [IO.File]::WriteAllText($csv, "ApplicationKey,DeviceId`r`napp-a,$($d1[0])`r`n")
    Check ((Read-DiscoveredAppsRelationExportCsv -Path $csv -CatalogAppIds @('app-a')).AcceptedRows -eq 1) 'Unquoted simple fields must be accepted.'

    $badCases = @(
        @{ Text = ''; Pattern = 'empty'; Label = 'empty file' }
        @{ Text = "`"ApplicationKey`",`"DeviceName`"`r`n"; Pattern = 'header is unexpected'; Label = 'wrong header' }
        @{ Text = "`"ApplicationKey`",`"DeviceId`",`"Platform`"`r`n"; Pattern = 'header is not'; Label = 'extra column' }
        @{ Text = (New-ExportText -Pairs @{ 'app-a' = @($d1[0], $d1[0]) }); Pattern = 'Repeated app/device pair'; Label = 'duplicate pair' }
        @{ Text = (New-ExportText -Pairs @{ 'app-a' = @($d1[0].ToUpperInvariant(), $d1[0]) }); Pattern = 'Repeated app/device pair'; Label = 'duplicate pair with other case' }
        @{ Text = (New-ExportText -Pairs @{ 'app-a' = @('not-a-guid') }); Pattern = 'invalid ApplicationKey or DeviceId'; Label = 'invalid DeviceId' }
        @{ Text = (New-ExportText -Pairs @{ ' app-a' = @($d1[0]) }); Pattern = 'invalid ApplicationKey'; Label = 'padded key' }
        @{ Text = (New-ExportText -Pairs @{ '' = @($d1[0]) }); Pattern = 'invalid ApplicationKey'; Label = 'empty key' }
        @{ Text = "`"ApplicationKey`",`"DeviceId`"`r`n`"app-a`",`"$($d1[0])`",`"x`"`r`n"; Pattern = 'malformed'; Label = 'extra field in a row' }
        @{ Text = "`"ApplicationKey`",`"DeviceId`"`r`n`r`n"; Pattern = 'malformed'; Label = 'blank row' }
    )
    foreach ($case in $badCases) {
        [IO.File]::WriteAllText($csv, $case.Text)
        Assert-Throws { Read-DiscoveredAppsRelationExportCsv -Path $csv -CatalogAppIds @('app-a') } $case.Pattern "Export parsing must reject: $($case.Label)"
    }
    Assert-Throws { Read-DiscoveredAppsRelationExportCsv -Path $csv -CatalogAppIds @('app-a', 'APP-A') } 'repeated AppId' 'A repeated catalog id must be rejected.'

    # --- Random selection -----------------------------------------------------
    $items = @(1..20)
    $pick1 = @(Select-DiscoveredAppsRandomItems -Items $items -Count 5 -Random ([Random]::new(7)))
    $pick2 = @(Select-DiscoveredAppsRandomItems -Items $items -Count 5 -Random ([Random]::new(7)))
    Check ($pick1.Count -eq 5 -and (@($pick1 | Sort-Object -Unique).Count -eq 5)) 'Random selection must return distinct items.'
    Check (($pick1 -join ',') -eq ($pick2 -join ',')) 'Random selection must be reproducible from the seed.'
    Check (@(Select-DiscoveredAppsRandomItems -Items $items -Count 50 -Random ([Random]::new(1))).Count -eq 20) 'Random selection must be capped.'
    Check (@(Select-DiscoveredAppsRandomItems -Items @() -Count 5 -Random ([Random]::new(1))).Count -eq 0) 'Random selection of nothing must be empty.'

    # --- Graph single request -------------------------------------------------
    Reset-Graph
    $script:failQueue.Enqueue(429)
    $script:graphDevices['app-a'] = @()
    $single = Invoke-DiscoveredAppsGraphSingleRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/app-z?$select=id' -AcceptedStatusCodes @(404)
    Check ($single.StatusCode -eq 404 -and $script:requests.Count -eq 2) 'A 429 must be retried and an accepted 404 returned.'
    Reset-Graph
    $script:failQueue.Enqueue(403)
    Assert-Throws { Invoke-DiscoveredAppsGraphSingleRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/app-z?$select=id' } 'Status=403' 'A 403 must not be retried.'
    Check ($script:requests.Count -eq 1) 'A 403 must be sent once.'
    Reset-Graph
    1..3 | ForEach-Object { $script:failQueue.Enqueue(503) }
    Assert-Throws { Invoke-DiscoveredAppsGraphSingleRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/app-z?$select=id' } 'Attempts=3' 'Exhausted retries must throw.'

    # --- Graph verification sample --------------------------------------------
    function New-Fixture {
        Reset-Graph
        $script:apps = [Collections.Generic.List[object]]::new()
        $script:exportPairs = [ordered]@{}
        foreach ($index in 1..30) {
            $id = 'app-{0:00}' -f $index
            $devices = New-Guids ($index % 4 + 1)
            $script:graphDevices[$id] = $devices
            $script:exportPairs[$id] = $devices
            # Reported catalog counts are volatile: a few differ from the real relations.
            $reported = if ($index -in 10, 20, 30) { $devices.Count + 1 } else { $devices.Count }
            $script:apps.Add((New-App -Id $id -Count $reported))
        }
        # Catalog app with no relation on either side (seen in production).
        $script:graphDevices['app-empty'] = @()
        $script:apps.Add((New-App -Id 'app-empty' -Count 1))
        $script:exportPairs['gone-key'] = New-Guids 2
    }
    function Get-FixtureDevices {
        [IO.File]::WriteAllText($csv, (New-ExportText -Pairs $script:exportPairs))
        (Read-DiscoveredAppsRelationExportCsv -Path $csv -CatalogAppIds ([string[]]@($script:apps | ForEach-Object { $_.id }))).DevicesByAppId
    }
    function Invoke-Sample { param($Devices, [int]$Random = 100, [int]$MaxAbsent = 500)
        Test-DiscoveredAppsRelationExportSample -Apps $script:apps.ToArray() -DevicesByAppId $Devices -UnknownKeys @('gone-key') -RandomApps $Random -TopApps 3 -MismatchApps 2 -MaxAbsentApps $MaxAbsent -Seed 42
    }

    New-Fixture
    $sample = Invoke-Sample (Get-FixtureDevices)
    Check ($sample.Passed -and $sample.CheckedApps -eq 31 -and $sample.AbsentApps -eq 1 -and $sample.CheckedUnknownKeys -eq 1) "Identical export and Graph must pass ($($sample.Reason))."
    $partial = Invoke-Sample (Get-FixtureDevices) -Random 0
    Check ($partial.Passed -and $partial.CheckedApps -eq 6) "Absent, top and mismatch apps are always checked ($($partial.CheckedApps))."

    New-Fixture
    $script:graphDevices['app-05'] = @($script:graphDevices['app-05']) + (New-Guids 1)
    $sample = Invoke-Sample (Get-FixtureDevices)
    Check (-not $sample.Passed -and $sample.Reason -match 'Mismatch: AppId=app-05;.*OnlyGraph=1; OnlyExport=0') "A device missing from the export must fail ($($sample.Reason))."

    New-Fixture
    $script:exportPairs['app-06'] = @($script:exportPairs['app-06']) + (New-Guids 1)
    $sample = Invoke-Sample (Get-FixtureDevices)
    Check (-not $sample.Passed -and $sample.Reason -match 'AppId=app-06;.*OnlyExport=1') "An extra export device must fail ($($sample.Reason))."

    New-Fixture
    $script:graphDevices['app-empty'] = New-Guids 1
    $sample = Invoke-Sample (Get-FixtureDevices) -Random 0
    Check (-not $sample.Passed -and $sample.Reason -match 'AppId=app-empty; Check=AbsentFromExport') "A catalog app absent from the export but present in Graph must fail ($($sample.Reason))."

    New-Fixture
    $script:graphDevices['gone-key'] = @()
    $sample = Invoke-Sample (Get-FixtureDevices) -Random 0
    Check (-not $sample.Passed -and $sample.Reason -match '^UnknownKey: .*HTTP 200') "A non-catalog key known to Graph must fail ($($sample.Reason))."

    New-Fixture
    $sample = Invoke-Sample (Get-FixtureDevices) -MaxAbsent 0
    Check (-not $sample.Passed -and $sample.Reason -match '^Coverage' -and $sample.CheckedApps -eq 0) 'Too many absent catalog apps must fail before any Graph read.'

    # --- Relation rows are byte-identical to the Graph path writer ------------
    New-Fixture
    $devices = Get-FixtureDevices
    $graphPath = Join-Path $testRoot 'graph.partial.csv'
    $exportPath = Join-Path $testRoot 'export.partial.csv'
    Write-DiscoveredAppsCsvRows -Path $graphPath -Rows @()
    Write-DiscoveredAppsCsvRows -Path $exportPath -Rows @()
    foreach ($app in $script:apps) {
        $rows = @($devices[$app.id] | ForEach-Object { [pscustomobject]@{ AppId = $app.id; DeviceId = $_ } })
        Write-DiscoveredAppsCsvRows -Path $graphPath -Rows $rows
    }
    $written = Write-DiscoveredAppsRelationExportRows -Path $exportPath -Apps $script:apps.ToArray() -DevicesByAppId $devices -TenantKey 'tenant-test'
    $expectedRows = 0; foreach ($app in $script:apps) { $expectedRows += $devices[$app.id].Count }
    Check ($written.Rows -eq $expectedRows -and $written.CountsByAppId['app-empty'] -eq 0 -and $written.CountsByAppId.Count -eq $script:apps.Count) 'Export writer counters are wrong.'
    Check ((Get-FileHash -LiteralPath $graphPath).Hash -eq (Get-FileHash -LiteralPath $exportPath).Hash) 'Export rows must be byte-identical to the Graph path CSV rows.'
    Assert-Throws { Write-DiscoveredAppsRelationExportRows -Path $exportPath -Apps $script:apps.ToArray() -DevicesByAppId $devices -TenantKey 'tenant-test' } 'empty partial CSV' 'The export writer must refuse a partial CSV that already has rows.'

    # --- End-to-end source: job, download, temporary files --------------------
    New-Fixture
    $script:exportText = New-ExportText -Pairs $script:exportPairs
    $before = @(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'SmartM365-DiscoveredApps-Export-*').Count
    $source = Invoke-DiscoveredAppsRelationExportSource -Apps $script:apps.ToArray() -TimeoutSeconds 60 -RandomApps 5 -TopApps 2 -MismatchApps 1 -Seed 3
    $after = @(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'SmartM365-DiscoveredApps-Export-*').Count
    Check ($source.Passed -and $source.JobId -eq 'AppInvRawData_job1' -and $source.DevicesByAppId['app-01'].Count -gt 0) "End-to-end export source must pass ($($source.Reason))."
    Check ($script:downloadCount -eq 1 -and $script:jobPolls -eq 2) 'The job must be polled until completed, then downloaded once.'
    Check ($after -eq $before) 'Temporary export files must be removed.'

    New-Fixture
    $script:exportText = New-ExportText -Pairs $script:exportPairs
    $script:jobFinalStatus = 'failed'
    Assert-Throws { Invoke-DiscoveredAppsRelationExportSource -Apps $script:apps.ToArray() -TimeoutSeconds 60 -Seed 3 } 'export job failed' 'A failed export job must throw (caller falls back to Graph).'
    Check ($script:downloadCount -eq 0) 'A failed job must not be downloaded.'

    New-Fixture
    $script:exportText = New-ExportText -Pairs $script:exportPairs
    $script:exportExtraFile = $true
    Assert-Throws { Invoke-DiscoveredAppsRelationExportSource -Apps $script:apps.ToArray() -TimeoutSeconds 60 -Seed 3 } 'exactly one CSV' 'An export archive with several files must be rejected.'

    New-Fixture
    $script:exportText = New-ExportText -Pairs $script:exportPairs
    $script:graphDevices['app-10'] = New-Guids 1
    $source = Invoke-DiscoveredAppsRelationExportSource -Apps $script:apps.ToArray() -TimeoutSeconds 60 -RandomApps 100 -Seed 3
    Check (-not $source.Passed -and $source.Reason -match 'app-10') 'A verification difference must be returned as not passed.'
    Check (@($script:logs | Where-Object { $_ -like 'WARNING|Relation export Graph verification: Passed=False*' }).Count -ge 1) 'A failed verification must be logged as a warning.'
}
finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: $script:testCount Discovered Apps export relation checks; PowerShell $($PSVersionTable.PSVersion)."
