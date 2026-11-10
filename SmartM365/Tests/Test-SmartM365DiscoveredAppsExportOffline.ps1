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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0tkM5RhAPE5td
# l7BLN4e8vlisaHgI3cUzIYgAHjuET6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMdvpH7NnQS1qojK51ENcAeRoYJuVVlKc+3cHtrkXnhbMA0GCSqG
# SIb3DQEBAQUABIIBgKfCy0cU2lQwVYV8sHCZhg3kNJ+VZioucTlciMsR44gjoLp4
# PYJzaISQTTAGk8lwde5nc5+rE7VYbkOQEet+DQ+5Mq90jBYBihEJkKoWAG6dfDCM
# 4UJ+95LsC4JvDV5NbC1aASfXXmyA9dEb5O7ZBJxMUr3qt62D4icdt58tfeDXJoS0
# go+Pc0wDuHC9oaeaELfHXJGSPybj4KQxSt1Uy4YJOkHLj7Pes1yCihOk0HX7nBls
# pyNdeJfQozCNlIApp8wlSzcsYO2o7uBmmb7H9L+Eo8uL8ZdCpqf+KFDFo5E4U5C0
# I0LrZS0/aMRg95+u4GDPEHI2O++0TnVacjzvuY9mLWaGWJxihz50crwaQDylTSgB
# XZ0yegQXTd1uKIyoEQYhyrfv2kEw26MU29BNzvvw2zRZSf9DPXivEQVkrQEJG3hE
# V8z8+pRMBaHuKAIwzkjmEb4K8+JZ+6cXIaj/2LWE4V0GohUCBJKNOSTwa19fmGbA
# IKtKX/ughldinJGlXqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTAxNTMx
# NTFaMC8GCSqGSIb3DQEJBDEiBCDvMcuY0MhAQRtHda3Di7qc8XYhnYtaIYbv9544
# snQq4zANBgkqhkiG9w0BAQEFAASCAgBsAWyztwqPXWu2TS09HRjN7Ww5GqY6tukI
# G0sufjoEVMKbPEtLbRdoS39bAJhmPtTMm/+xEO6FVFWepE7+Gk7nge07GGBFVVMM
# bKLeQPo+25boNFPsHk9uB4GtoxXG+XyJg6ypmZ4GdbKYtm5PWqMx8gz0YfNgLNIf
# zXq1nBiDW0q+KEAxZn+FX/LcM5dH5i12V93oXqbZyr/KoDTFymWXcjA97bhjXCHH
# QLp/qs/JeFummTzxvSHgcuUcYrBeXQS1dawhHbUaq6WXL5SORzlPY57kWY3+koAh
# IlxPwe8GgzwVa/40QS+e+jciwaWsWJO+qn+4LoWLaI4cxFPW/MlIrndjQnDbyCC1
# L/juRYEgEx3069ZOuJdfdYk+TU08oOXntxCOFLAVXh/vNPCfxkVrAGoDmNG6I5Rz
# T0m2ymp+9OmvG/fTPWrKNcBd7fiX+s/I0scRqlLu1qfK00sez1j83UfKNqRj/ZXB
# j3QlGTspQzQLSy3YSQKAnsL6kTtEmf2O94fz5iArn8f4APiwIJ39oBIVTBio010K
# RdyTz69H1FiYZbQDezPG1mOTmVuuKMiF2PPEBhngdPGbtDMVK3apTnJRJ6bYr4Gd
# jYRhB+uC48UH0CffUHlqly6WXVgm+4eMlZpACtfzNdt6VTO+hoMZWY7Hu/u1sgaQ
# ma6lVVnT1w==
# SIG # End signature block
