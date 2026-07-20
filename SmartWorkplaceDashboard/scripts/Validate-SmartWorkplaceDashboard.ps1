# Version: 4.3.0
[CmdletBinding()]
param(
    [string]$DashboardRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$CheckSourceFiles,
    [switch]$SkipPbirValidation
)
$ErrorActionPreference = 'Stop'
function Read-Json([string]$Path) { if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){throw "Missing file: $Path"}; Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -Depth 100 }
function Get-CsvHeader([string]$Path,[string]$Delimiter){ Add-Type -AssemblyName Microsoft.VisualBasic; $p=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path,[Text.Encoding]::UTF8,$true); try{$p.TextFieldType=[Microsoft.VisualBasic.FileIO.FieldType]::Delimited;$p.SetDelimiters($Delimiter);$p.HasFieldsEnclosedInQuotes=$true;@($p.ReadFields())}finally{$p.Close()} }
function Get-Delimiter([string]$Line){$comma=0;$semi=0;$quoted=$false;for($i=0;$i-lt$Line.Length;$i++){if($Line[$i]-eq '"'){if($quoted-and$i+1-lt$Line.Length-and$Line[$i+1]-eq '"'){$i++}else{$quoted=-not$quoted}}elseif(-not$quoted){if($Line[$i]-eq ','){$comma++}elseif($Line[$i]-eq ';'){$semi++}}};if($semi-gt$comma){';'}else{','}}
function Visit-Node($Node,[scriptblock]$Action){if($null-eq$Node){return};&$Action $Node;if($Node-is[Collections.IDictionary]){foreach($v in $Node.Values){Visit-Node $v $Action}}elseif($Node-is[Management.Automation.PSCustomObject]){foreach($p in $Node.PSObject.Properties){Visit-Node $p.Value $Action}}elseif($Node-is[Collections.IEnumerable]-and$Node-isnot[string]){foreach($v in $Node){Visit-Node $v $Action}}}
$pbip=Join-Path $DashboardRoot 'pbip';$report=Join-Path $pbip 'SmartWorkplaceDashboard.Report';$semantic=Join-Path $pbip 'SmartWorkplaceDashboard.SemanticModel';$modelPath=Join-Path $semantic 'model.bim';$definition=Join-Path $report 'definition';$selectionPath=Join-Path $DashboardRoot 'source-selection.json'
$required=@((Join-Path $pbip 'SmartWorkplaceDashboard.pbip'),(Join-Path $report '.platform'),(Join-Path $report 'definition.pbir'),(Join-Path $definition 'version.json'),(Join-Path $definition 'report.json'),(Join-Path $definition 'pages\pages.json'),(Join-Path $semantic 'definition.pbism'),$selectionPath,$modelPath);foreach($p in $required){$null=Read-Json $p}
$model=Read-Json $modelPath;$selection=Read-Json $selectionPath
if($model.compatibilityLevel-ne1600){throw "compatibilityLevel must be 1600; found $($model.compatibilityLevel)"}
$files=@($selection.includedFiles);if($files.Count-ne89){throw "Expected 89 selected CSV files; found $($files.Count)"};if(($files|Sort-Object -Unique).Count-ne$files.Count){throw 'Duplicate selected CSV file'}
$approvedSummary=@($selection.includeOverrides);foreach($f in $files){if(-not$f.EndsWith('.csv',[StringComparison]::OrdinalIgnoreCase)){throw "Not a CSV: $f"};if($f-match'_MAXITEMS'){throw "Bounded export selected: $f"};if($f-match'Summary' -and$approvedSummary-notcontains$f){throw "Unapproved Summary selected: $f"}}
$expectedNames=@($files|ForEach-Object{[IO.Path]::GetFileNameWithoutExtension($_)});$tables=@($model.model.tables);$actualNames=@($tables.name)
$missing=@($expectedNames|Where-Object{$actualNames-notcontains$_});$extra=@($actualNames|Where-Object{$expectedNames-notcontains$_});if($missing.Count){throw "Missing source tables: $($missing-join', ')"};if($extra.Count){throw "Non-source tables found: $($extra-join', ')"};if($tables.Count-ne89){throw "Expected exactly 89 source tables; found $($tables.Count)"}
$old=@('DimUser','DimDevice','DimLicenseSku','FactAction','FactBackupMailbox','FactDataQuality','FactDeviceCompliance','FactMailbox','FactSharePointSite','FactSourceFreshness','FactSyncHealth','FactTeam','FactUpgradeEligibility','FactUserActivity','FactUserLicense','FactWindowsUpdate','HistoryCoverage','DashboardMetrics','Measures');$present=@($actualNames|Where-Object{$old-contains$_});if($present.Count){throw "Legacy work tables remain: $($present-join', ')"}
$metadata=@('__SnapshotDate','__SnapshotDateTime','__SnapshotPeriod','__IsCurrent','__SourceFile','__SourceFolder');$technical=@('__SkuPlanKey');$derived=@('IsEnabled','PlanStatus','IsServicePlanActive');$dataRoot=if($env:SMART_M365_DATA_ROOT){$env:SMART_M365_DATA_ROOT}else{'C:\SmartM365\DATA'};$dataLast=Join-Path $dataRoot 'DATA-LAST'
foreach ($f in $files) {
    $name = [IO.Path]::GetFileNameWithoutExtension($f)
    $t = $tables | Where-Object name -eq $name | Select-Object -First 1
    if ($null -eq $t) { throw "Missing source table: $name" }
    if ($t.isHidden -eq $true) { throw "Source table is hidden: $name" }
    $partition = [string]$t.partitions.source.expression
    if ($partition -notmatch 'fnLoadSourceTable\(' -or $partition -notmatch [regex]::Escape($f)) {
        throw "Direct source contract missing for $name"
    }
    $csv = Join-Path $dataLast $f
    $cols = @($t.columns.name)
    $base = @($cols | Where-Object { $metadata -notcontains $_ -and $technical -notcontains $_ -and $derived -notcontains $_ })
    $overrideProperty = if ($null -ne $selection.schemaOverrides) { $selection.schemaOverrides.PSObject.Properties[$f] } else { $null }
    if ($null -ne $overrideProperty) {
        $overrideHeader = @($overrideProperty.Value.columns)
        if (($overrideHeader -join "`u{001f}") -ne ($base -join "`u{001f}")) { throw "Schema override/model columns differ for $name" }
    }
    if ($CheckSourceFiles) {
        if (-not (Test-Path -LiteralPath $csv)) { throw "Missing source CSV: $csv" }
        $line = Get-Content -LiteralPath $csv -TotalCount 1
        $delimiter = Get-Delimiter $line
        $header = @(Get-CsvHeader $csv $delimiter)
        if (($header -join "`u{001f}") -ne ($base -join "`u{001f}")) { throw "CSV/model columns differ for $name" }
    }
    elseif ($null -eq $overrideProperty -and (Test-Path -LiteralPath $csv)) {
        $line = Get-Content -LiteralPath $csv -TotalCount 1
        $delimiter = Get-Delimiter $line
        $header = @(Get-CsvHeader $csv $delimiter)
        if (($header -join "`u{001f}") -ne ($base -join "`u{001f}")) { throw "CSV/model columns differ for $name" }
    }
    foreach ($m in $metadata) { if ($cols -notcontains $m) { throw "Missing snapshot metadata $name[$m]" } }}
$relationships=@($model.model.relationships)
if($relationships.Count-ne2){throw "Expected exactly 2 validated service-plan relationships; found $($relationships.Count)"}
$expectedRelationships=@{
    'rel_ServicePlanStates_Catalog' = @{ fromTable='M365_Licenses_UserServicePlanStates'; fromColumn='__SkuPlanKey'; toTable='M365_Licenses_ServicePlans_Catalog'; toColumn='__SkuPlanKey' }
    'rel_ServicePlanStates_Users' = @{ fromTable='M365_Licenses_UserServicePlanStates'; fromColumn='UserId'; toTable='M365_Users_Active'; toColumn='Object Id' }
}
foreach($relationshipName in $expectedRelationships.Keys){
    $relationship=$relationships|Where-Object name -eq $relationshipName|Select-Object -First 1
    if($null-eq$relationship){throw "Missing validated relationship: $relationshipName"}
    foreach($propertyName in @('fromTable','fromColumn','toTable','toColumn')){
        if([string]$relationship.$propertyName -ne [string]$expectedRelationships[$relationshipName][$propertyName]){throw "Relationship regression: $relationshipName.$propertyName"}
    }
}
foreach($tableName in @('M365_Licenses_UserServicePlanStates','M365_Licenses_ServicePlans_Catalog')){
    $table=$tables|Where-Object name -eq $tableName|Select-Object -First 1
    $keyColumn=$table.columns|Where-Object name -eq '__SkuPlanKey'|Select-Object -First 1
    if($null-eq$keyColumn -or $keyColumn.dataType-ne'string' -or $keyColumn.isHidden-ne$true){throw "Missing hidden service-plan key on $tableName"}
    if([string]$table.partitions.source.expression -notmatch 'Table.AddColumn\([^,]+, "__SkuPlanKey"'){throw "Service-plan key is not generated in Power Query for $tableName"}
}
$stateTable=$tables|Where-Object name -eq 'M365_Licenses_UserServicePlanStates'|Select-Object -First 1
$stateColumns=@{};foreach($column in $stateTable.columns){$stateColumns[[string]$column.name]=$column}
foreach($expected in @{'StateCode'='string';'IsEnabled'='boolean';'PlanStatus'='string';'IsServicePlanActive'='boolean'}.GetEnumerator()){
    if(-not$stateColumns.ContainsKey($expected.Key)-or$stateColumns[$expected.Key].dataType-ne$expected.Value){throw "Service-plan compatibility field regression: $($expected.Key)"}
}
$stateExpression=[string]$stateTable.partitions.source.expression
foreach($fragment in @('Table.AddColumn(Source, "IsEnabled"','Table.AddColumn(WithIsEnabled, "PlanStatus"','Table.AddColumn(WithPlanStatus, "IsServicePlanActive"','List.Contains({"A","PA","PI","PP","E"},State)','State="A" then "Success"','State="D" then "Disabled"','State="PA" then "PendingActivation"','State="PI" then "PendingInput"','State="PP" then "PendingProvisioning"','State="E" then "Error"','Text.StartsWith(State,"EN:"','Text.StartsWith(State,"DIS:"','Text.Upper(Text.Trim(Text.From([StateCode])))="A"')){
    if(-not$stateExpression.Contains($fragment,[StringComparison]::Ordinal)){throw "StateCode compatibility mapping regression: $fragment"}
}
$typedExpectations = @{
    'M365_Users_Active' = @{ 'AccountEnabled' = 'boolean' }
    'M365_Users_Activity' = @{ 'HasAnyM365Activity' = 'boolean'; 'LastActivityDate' = 'dateTime' }
    'Intune_Devices_Inventory' = @{ 'IsCompliant' = 'boolean'; 'LastSyncDateTime' = 'dateTime' }
    'Intune_Windows11_Readiness_Issues' = @{ 'IsBlocking' = 'boolean' }
    'M365_SPO_Sites' = @{ 'IsInactive' = 'boolean'; 'IsOrphaned' = 'boolean' }
    'M365_SPO_Tenant' = @{ 'StorageUsedTB' = 'double'; 'StorageCapacityTB' = 'double'; 'StorageUtilizationPercent' = 'double'; 'IsPartialInventory' = 'boolean' }
    'M365_Mailbox_Usage' = @{ 'Storage Used (Byte)' = 'int64'; 'Prohibit Send/Receive Quota (Byte)' = 'int64'; 'Last Activity Date' = 'dateTime' }
    'Exchange_OnPrem_Servers_Compute' = @{ 'IsVirtualMachine' = 'boolean'; 'MemoryGB' = 'double'; 'PhysicalCoreCount' = 'int64' }
    'Exchange_OnPrem_Servers_LogicalDisks' = @{ 'LowSpaceWarning' = 'boolean'; 'FreePercent' = 'double' }
    'Exchange_OnPrem_Servers_RemoteAccess' = @{ 'PingOk' = 'boolean'; 'WmiDcomOk' = 'boolean' }
    'Exchange_OnPrem_Servers_Inventory_Summary' = @{ 'ExecutionDate' = 'dateTime'; 'ExchangeServersCount' = 'int64'; 'TotalMemoryTB' = 'double'; 'ExchangeReadinessErrors' = 'int64' }
}
foreach ($tableName in $typedExpectations.Keys) {
    $typedTable = $tables | Where-Object name -eq $tableName | Select-Object -First 1
    $partitionText = [string]$typedTable.partitions.source.expression
    foreach ($columnName in $typedExpectations[$tableName].Keys) {
        $expectedType = $typedExpectations[$tableName][$columnName]
        $typedColumn = $typedTable.columns | Where-Object name -eq $columnName | Select-Object -First 1
        if ($typedColumn.dataType -ne $expectedType) { throw "Typed field regression: $tableName[$columnName] is $($typedColumn.dataType), expected $expectedType" }
        $mKind = if ($expectedType -eq 'boolean') { 'logical' } elseif ($expectedType -eq 'dateTime') { 'datetime' } elseif ($expectedType -eq 'double') { 'number' } else { $expectedType }
        if ($partitionText -notmatch [regex]::Escape("{`"$columnName`", `"$mKind`"}")) { throw "Power Query type map regression: $tableName[$columnName] must use $mKind" }
    }
}
$measureTable=$tables|Where-Object name -eq 'M365_Users_Active'|Select-Object -First 1;$measureNames=@($measureTable.measures.name);if($measureNames.Count -lt 60){throw "Expected at least 60 measures; found $($measureNames.Count)"};foreach($n in @('Enabled Users','Enabled Users Inactive 90d','Purchased Licenses','M365 SKU Utilization','Service Plan Assignments','Service Plan Assigned Users','Enabled Service Plan Assignments','Disabled Service Plan Assignments','Users with Disabled Service Plans','Service Plan Enablement Rate','Managed Devices','Device Compliance Rate','Windows 11 Affected Devices','Mailboxes','Mailbox Quota Utilization','Teams','SharePoint Sites','SharePoint Storage License Units','SharePoint Estimated Capacity TB','SharePoint Estimated Remaining TB','SharePoint Estimated Utilization','SharePoint Approx Project Visio Capacity TB','Migration Failed Items','Protected Mailboxes','Backup Coverage Rate','Hybrid Identity Affected Users','Model Device Count','Devices per Application','AD Enabled Users Trend','Managed Devices Trend','Selected Source Tables','On-prem Hosted Mailboxes','Online Placement Rate','Remote Mailboxes without EXO','Potential Dual-hosted SMTP','Exchange Servers','Server Inventory Coverage','Low Space Volumes')){if($measureNames -notcontains $n){throw "Missing measure: $n"}}
$tableMap=@{};foreach($t in $tables){$tableMap[$t.name]=@($t.columns.name)};$measureSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase);foreach($n in $measureNames){$null=$measureSet.Add($n)}
foreach($m in $measureTable.measures){$matches=[regex]::Matches([string]$m.expression,"'([^']+)'\[([^\]]+)\]");foreach($x in $matches){$tn=$x.Groups[1].Value;$cn=$x.Groups[2].Value;if(-not $tableMap.ContainsKey($tn)){throw "Measure $($m.name) references missing table $tn"};if($tableMap[$tn] -notcontains $cn){throw "Measure $($m.name) references missing field $tn[$cn]"}}}
$pagesMeta=Read-Json (Join-Path $definition 'pages\pages.json');if(@($pagesMeta.pageOrder).Count-ne15){throw "Expected 15 report pages; found $(@($pagesMeta.pageOrder).Count)"}
$visualFiles=@(Get-ChildItem -LiteralPath (Join-Path $definition 'pages') -Filter visual.json -File -Recurse);if($visualFiles.Count-lt120){throw "Expected at least 120 visuals; found $($visualFiles.Count)"}
$logoVisuals=@($visualFiles|ForEach-Object{Read-Json $_.FullName}|Where-Object{$_.visual.visualType-eq'image'})
if($logoVisuals.Count-ne15){throw "Expected one logo image per page; found $($logoVisuals.Count)"}
foreach($logo in $logoVisuals){$url=[string]$logo.visual.objects.general[0].properties.imageUrl.expr.Literal.Value;if(-not$url.StartsWith("'data:image/png;base64,")){throw 'Logo image is not embedded as PNG data'}}
$slicers=@($visualFiles|ForEach-Object{Read-Json $_.FullName}|Where-Object{$_.visual.visualType-eq'slicer'})
if($slicers.Count-lt1){throw 'Ownership slicer is missing'}
foreach($vf in $visualFiles){$v=Read-Json $vf.FullName;Visit-Node $v {param($node);foreach($kind in @('Column','Measure')){$exact=$node.PSObject.Properties[$kind];if($null -ne $exact){$entry=$exact.Value;$entity=$entry.Expression.SourceRef.Entity;$property=$entry.Property;if(-not $entity -or -not $property){throw "Incomplete $kind reference in $($vf.FullName)"};if(-not $tableMap.ContainsKey([string]$entity)){throw "Visual references missing table $entity"};if($kind -eq 'Column' -and $tableMap[[string]$entity] -notcontains [string]$property){throw "Visual references missing field $entity.$property"};if($kind -eq 'Measure' -and -not $measureSet.Contains([string]$property)){throw "Visual references missing measure $property"}}}}}
foreach($n in @('fnGetSourceFiles','fnLoadSourceTable','fnToLogical','fnToDateTime','fnToNumber','fnToInt64')){if(-not ($model.model.expressions|Where-Object name -eq $n)){throw "Missing Power Query expression: $n"}}
$invalidDaxMeasures=@($tables.measures|Where-Object { [string]$_.expression -match '&&\s*VAR\b' })
if($invalidDaxMeasures.Count -gt 0){throw "Invalid DAX variable placement in measures: $($invalidDaxMeasures.name -join ', ')"}
$loaderExpression=[string](($model.model.expressions|Where-Object name -eq 'fnLoadSourceTable'|Select-Object -First 1).expression)
if(([regex]::Matches($loaderExpression,'\(')).Count -ne ([regex]::Matches($loaderExpression,'\)')).Count){throw 'Unbalanced parentheses in fnLoadSourceTable'}
$forbidden=@('SmartWorkplaceCMDB','LocalDateTable_','C:\Users\');$publishable=@(Get-ChildItem -LiteralPath $DashboardRoot -File -Recurse|Where-Object{$_.FullName -notmatch '\\\.pbi\\' -and $_.Name -ne 'LOCAL_MEMORY.md' -and $_.FullName -ne $PSCommandPath -and $_.Extension -in @('.js','.ps1','.json','.bim','.pbip','.pbir','.pbism','.pq','.md')});foreach($p in $publishable){$text=Get-Content -LiteralPath $p.FullName -Raw;foreach($term in $forbidden){if($text.Contains($term,[StringComparison]::OrdinalIgnoreCase)){throw "Forbidden dependency '$term' in $($p.FullName)"}}}
$builder=Join-Path $DashboardRoot 'scripts\Build-SmartWorkplaceDashboard.js';& node --check $builder;if($LASTEXITCODE-ne0){throw 'Builder JavaScript syntax check failed'}
Write-Host "OK SmartWorkplaceDashboard: 89 source tables, $($measureNames.Count) measures, 15 pages, $($visualFiles.Count) visuals, StateCode service-plan compatibility validated."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAhX+aAmSRfV94s
# wOlMVqyuD5iC03KuwJACJli/P6cc6KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICI8RMH9SLxFkhtfb5cmhLfYdA4BN/ntibIPbhc+ZlrYMA0GCSqG
# SIb3DQEBAQUABIIBgJGqwHthd0NwS+X+LVRhb47NJ2WqofE98P89SvVQvQFimUCJ
# poifVcVdJQcQ7RQViCci1/pw9aMbiuq90ol1ehTnpmUPymr+LHQ62ucC74a27snx
# Y+I1IxUsoyavhx62E2rtlRXRRJUVvLX+Y3R0bb5pLyrydzu/OxI7c5kwzqNseytm
# wLJJKyJVGL92Wab0mKTPILik5qMNzXwZkAmwFJDUnvPH69JBy9DZJQHBMzoyBEf2
# f8J22O3O3EcBLEUCmxhrfnWgCKQkXWlN/scHCpwR24byvL28sHXutIHZ0PpWVte0
# 6u6aKwsOgCkQ3EzSFcLPhCKrgGWtk/chMnPbcDajO+VYs2D8YZ+L5j0Hj05sPhy9
# JIOx7eJAGe24dxRCeRwQzNQKqxQrU10jlXEGOCVnh66aoqeEBXNp+sQ9mo1YsKga
# 792DasrbMFX494axgDCgJAxBsSMGAJX7sa6MQomDCA2ik1OhqfOik3nkNJI2M5GG
# 7/42C4uUF0C53WUxXqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjAxNTUw
# MzVaMC8GCSqGSIb3DQEJBDEiBCAbm6McdFhC0zRXRYJkif6YIqN/hJM7vDOO0f2D
# dFC2iDANBgkqhkiG9w0BAQEFAASCAgCebQ+dpxrLwWxk8/8hi56CAprd/49xQ5Ol
# g8pgKXi+TED45OoSa8u8yEcipJ9QzsTlOaoSyIhQ6Rb4Qy+kcPC0N/qC0IvvcQbM
# Ljp/+ULo2SIBTAGvWBOsjy0LI/WLBkMEt4dLFqJ09y8/3ZY0MirLPH0Zj3dEzP6T
# DoXtea7ZDES3mAR7z/Mck0jc2r4d/Qu65IlduOqYTjKKSQ80OM9ECQaWSiCBS2Lp
# pluKbT7SB58utCcFB+SOVQu8nfvMz+GPtz8ypfgxv/QGUHCXE+Q4sJZutpgeDGJW
# n2sXyWVCXVsD+3qVagDnGCfgc+9bXaMUIKC2hiUqMoMhur7uTcpKzDsWYYgLXmx5
# MrhQGg2e61Zz7oYECoT14MmHWichW8/4NohC8WetWNDInVVgmX6+yIKauRqUps67
# J25QPiShglcggt0rot7DoYbmoYJOwnX3OhDnCKa90G24SLYQo6KIb6O4TdYo76Pd
# xZn9+XK7nEt2bBxP5VbKKXs0YaqkuYYBKNUOPOGh4ku001nHh5l5CMp22gDxzONm
# bZ6PDvVE3PXqPr1qDBWop1NYPwj3pwECp+FbA68hK9dA2tjROFu70ri5gu0WA+hO
# aTTXR1ZYU792b8V9Ea96qoXwuGvsSIocPDI+u/bD+Owpzrxj7ZnG1DZ4i0oFpt/d
# jHUAlnP9Dw==
# SIG # End signature block
