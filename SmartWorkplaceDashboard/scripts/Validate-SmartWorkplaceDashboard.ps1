#Requires -Version 7.0
# Version: 3.8.0-beta.1 BETA
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
$files=@($selection.includedFiles);if($files.Count-ne90){throw "Expected 90 selected CSV files; found $($files.Count)"};if(($files|Sort-Object -Unique).Count-ne$files.Count){throw 'Duplicate selected CSV file'}
$approvedSummary=@($selection.includeOverrides);foreach($f in $files){if(-not$f.EndsWith('.csv',[StringComparison]::OrdinalIgnoreCase)){throw "Not a CSV: $f"};if($f-match'_MAXITEMS'){throw "Bounded export selected: $f"};if($f-match'Summary' -and$approvedSummary-notcontains$f){throw "Unapproved Summary selected: $f"}}
$expectedNames=@($files|ForEach-Object{[IO.Path]::GetFileNameWithoutExtension($_)});$derivedTableNames=@('DeviceDetail','UserDetail');$tables=@($model.model.tables);$actualNames=@($tables.name);$sourceTables=@($tables|Where-Object{$derivedTableNames-notcontains$_.name});$actualSourceNames=@($sourceTables.name)
$missing=@($expectedNames|Where-Object{$actualSourceNames-notcontains$_});$extra=@($actualSourceNames|Where-Object{$expectedNames-notcontains$_});if($missing.Count){throw "Missing source tables: $($missing-join', ')"};if($extra.Count){throw "Unexpected source tables found: $($extra-join', ')"};if($sourceTables.Count-ne90){throw "Expected exactly 90 source tables; found $($sourceTables.Count)"};if($tables.Count-ne92){throw "Expected 90 source tables and two derived tables; found $($tables.Count) tables"}
$old=@('DimUser','DimDevice','DimLicenseSku','FactAction','FactBackupMailbox','FactDataQuality','FactDeviceCompliance','FactMailbox','FactSharePointSite','FactSourceFreshness','FactSyncHealth','FactTeam','FactUpgradeEligibility','FactUserActivity','FactUserLicense','FactWindowsUpdate','HistoryCoverage','DashboardMetrics','Measures');$present=@($actualNames|Where-Object{$old-contains$_});if($present.Count){throw "Legacy work tables remain: $($present-join', ')"}
$metadata=@('__SnapshotDate','__SnapshotDateTime','__SnapshotPeriod','__IsCurrent','__SourceFile','__SourceFolder');$technical=@('__SkuPlanKey','__TenantAppKey');$derived=@('IsEnabled','PlanStatus','IsServicePlanActive','__WeekStart','WindowsVersionFamily','BIOSDateKnown','BIOSOlderThan5Years');$dataRoot=if($env:SMART_M365_DATA_ROOT){$env:SMART_M365_DATA_ROOT}else{'C:\SmartM365\DATA'};$dataLast=Join-Path $dataRoot 'DATA-LAST'
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
    foreach ($m in $metadata) { if ($cols -notcontains $m) { throw "Missing snapshot metadata $name[$m]" } }}
$adComputers=$tables|Where-Object name -eq 'AD_Computers_AllDomains'|Select-Object -First 1
if($null-eq$adComputers){throw 'Missing AD_Computers_AllDomains table'}
$adComputerTypes=@{Enabled='boolean';ExistsInIntune='boolean';IntuneDiskFreeGB='double';PhysicalMemoryGB='double';IntunePrimaryUserMailboxSizeGB='double'}
foreach($entry in $adComputerTypes.GetEnumerator()){
    $column=$adComputers.columns|Where-Object name -eq $entry.Key|Select-Object -First 1
    if($null-eq$column){throw "Missing canonical AD computer field: $($entry.Key)"}
    if([string]$column.dataType-ne[string]$entry.Value){throw "AD computer type regression: $($entry.Key) is $($column.dataType), expected $($entry.Value)"}
}
$adComputerPartition=[string]$adComputers.partitions[0].source.expression
foreach($fragment in @('{"Enabled", "logical"}','{"ExistsInIntune", "logical"}','{"IntuneDiskFreeGB", "number"}','{"PhysicalMemoryGB", "number"}','{"IntunePrimaryUserMailboxSizeGB", "number"}')){
    if(-not$adComputerPartition.Contains($fragment,[StringComparison]::Ordinal)){throw "AD computer typed-source contract regression: $fragment"}
}
$adEnabledMeasure=$tables.measures|Where-Object name -eq 'AD Enabled Computers'|Select-Object -First 1
if($null-eq$adEnabledMeasure){throw 'Missing AD Enabled Computers measure'}
$adEnabledExpression=[string]$adEnabledMeasure.expression
if(-not$adEnabledExpression.Contains("'AD_Computers_AllDomains'[Enabled]=TRUE()",[StringComparison]::Ordinal)){throw 'AD Enabled Computers must filter the canonical boolean field'}
if($adEnabledExpression.Contains('[Enabled]="True"',[StringComparison]::Ordinal)){throw 'AD Enabled Computers still compares Enabled as text'}
$deviceDetail=$tables|Where-Object name -eq 'DeviceDetail'|Select-Object -First 1
if($null-eq$deviceDetail){throw 'Missing derived DeviceDetail table'}
if(@($deviceDetail.partitions).Count-ne1-or[string]$deviceDetail.partitions[0].source.type-ne'm'){throw 'DeviceDetail must have one M partition'}
$deviceDetailColumns=@{};foreach($column in $deviceDetail.columns){$deviceDetailColumns[[string]$column.name]=$column}
$requiredDeviceDetailFields=@{
    'CanonicalDeviceKey'='string';'DeviceName'='string';'PrimaryUserUPN'='string';'Country'='string';'SourceCoverage'='string';
    'InAD'='boolean';'InEntra'='boolean';'InIntune'='boolean';'ADEnabled'='boolean';'EntraAccountEnabled'='boolean';'EntraIsManaged'='boolean';'EntraIsCompliant'='boolean';'IntuneIsCompliant'='boolean';
    'ADLastLogonDateTime'='dateTime';'EntraLastSignInDateTime'='dateTime';'IntuneLastSyncDateTime'='dateTime';'LastActivityDateTime'='dateTime';
    'DaysSinceLastActivity'='int64';'PhysicalMemoryGB'='double';'StorageTotalGB'='double';'StorageFreeGB'='double';'ActionPriority'='int64';'ActionRequired'='boolean';'RecommendedAction'='string'
}
foreach($entry in $requiredDeviceDetailFields.GetEnumerator()){
    if(-not$deviceDetailColumns.ContainsKey($entry.Key)){throw "Missing DeviceDetail field: $($entry.Key)"}
    if([string]$deviceDetailColumns[$entry.Key].dataType-ne[string]$entry.Value){throw "DeviceDetail type regression: $($entry.Key) is $($deviceDetailColumns[$entry.Key].dataType), expected $($entry.Value)"}
}
foreach($hiddenField in @('CanonicalDeviceKey','ADObjectGUID','EntraObjectId','EntraDeviceId','IntuneDeviceId','ActionPriority')){if($deviceDetailColumns[$hiddenField].isHidden-ne$true){throw "DeviceDetail technical field must be hidden: $hiddenField"}}
$deviceDetailExpression=[string]$deviceDetail.partitions[0].source.expression
foreach($fragment in @('#"AD_Computers_AllDomains"','#"M365_Entra_Devices"','#"Intune_Devices_Inventory"','#"M365_Users_Active"','[__IsCurrent]=true','A_EntraObjectId','A_EntraDeviceId','A_IntuneDeviceId','I_EntraObjectId','I_AzureADDeviceId','"AD:"&Norm([A_ObjectGUID])','"INTUNE:"&Norm([I_DeviceId])','"ENTRA:"&[E_ObjectKey]')){if(-not$deviceDetailExpression.Contains($fragment,[StringComparison]::Ordinal)){throw "DeviceDetail exact-key contract regression: $fragment"}}
foreach($joinLine in [regex]::Matches($deviceDetailExpression,'Table\.NestedJoin\([^)]*\)')){if($joinLine.Value-match 'DisplayName|DeviceName|A_Name|I_DeviceName'){throw "DeviceDetail must not join devices by name: $($joinLine.Value)"}}$userDetail=$tables|Where-Object name -eq 'UserDetail'|Select-Object -First 1
if($null-eq$userDetail){throw 'Missing derived UserDetail table'}
if(@($userDetail.partitions).Count-ne1-or[string]$userDetail.partitions[0].source.type-ne'm'){throw 'UserDetail must have one M partition'}
$userDetailColumns=@{};foreach($column in $userDetail.columns){$userDetailColumns[[string]$column.name]=$column}
$requiredUserDetailFields=@{
    'CanonicalUserKey'='string';'DisplayName'='string';'UserPrincipalName'='string';'PrimarySmtpAddress'='string';'Country'='string';'SourceCoverage'='string';'IdentityMatchStatus'='string';'HybridIdentityType'='string';
    'InAD'='boolean';'InEntra'='boolean';'InExchangeOnPrem'='boolean';'InExchangeOnline'='boolean';'ADAccountEnabled'='boolean';'EntraAccountEnabled'='boolean';'IsEnabled'='boolean';'AccountStatus'='string';
    'DirSyncEnabled'='boolean';'OnPremisesSyncEnabled'='boolean';'IsLikelyServiceAccount'='boolean';'EXOMailboxSizeGB'='double';'OnPremMailboxSizeGB'='double';'TotalMailboxSizeGB'='double';'ArchiveMailboxSizeGB'='double';'MailboxItemCount'='int64';
    'ADLastLogonDateTime'='dateTime';'EntraLastSignInDateTime'='dateTime';'EXOLastUserActionDateTime'='dateTime';'OnPremLastLogonDateTime'='dateTime';'LastActivityDateTime'='dateTime';'DaysSinceLastActivity'='int64';
    'ActionPriority'='int64';'ActionRequired'='boolean';'RecommendedAction'='string'
}
foreach($entry in $requiredUserDetailFields.GetEnumerator()){
    if(-not$userDetailColumns.ContainsKey($entry.Key)){throw "Missing UserDetail field: $($entry.Key)"}
    if([string]$userDetailColumns[$entry.Key].dataType-ne[string]$entry.Value){throw "UserDetail type regression: $($entry.Key) is $($userDetailColumns[$entry.Key].dataType), expected $($entry.Value)"}
}
foreach($hiddenField in @('CanonicalUserKey','ADObjectGUID','EntraObjectId','OnPremExchangeGuid','EXOMailboxGuid','ActionPriority')){if($userDetailColumns[$hiddenField].isHidden-ne$true){throw "UserDetail technical field must be hidden: $hiddenField"}}
$userDetailExpression=[string]$userDetail.partitions[0].source.expression
foreach($fragment in @('#"AD_Users_AllDomains"','#"M365_Users_Active"','#"Exchange_OnPrem_Mailboxes_AllDomains"','#"Exchange_EXO_Mailboxes_AllDomains"','[__IsCurrent]=true','A_ImmutableId','A_SID','A_UPN','O_ObjectGUID','X_ExternalDirectoryObjectId','CandidateCount({[CanonicalByImmutable],[CanonicalBySID],[CanonicalByUPN]})>1','FirstText({[CanonicalByImmutable],[CanonicalBySID],[CanonicalByUPN]','FirstText({[CanonicalByADObject],[CanonicalByUPN],[CanonicalBySmtp]','FirstText({[CanonicalByEntraObject],[CanonicalByImmutable],[CanonicalByUPN],[CanonicalBySmtp]','"AD:"&Norm([A_ObjectGUID])','"ONPREM:"&Norm([O_ObjectGUID])','"EXO:"&Norm([X_MailboxGuid])')){if(-not$userDetailExpression.Contains($fragment,[StringComparison]::Ordinal)){throw "UserDetail exact-key contract regression: $fragment"}}
foreach($joinLine in [regex]::Matches($userDetailExpression,'Table\.NestedJoin\([^)]*\)')){if($joinLine.Value-match 'DisplayName|GivenName|Surname'){throw "UserDetail must not join users by display name: $($joinLine.Value)"}}
if($userDetailExpression-match 'ProxyAddresses|EmailAddresses'){throw 'UserDetail must not merge users from alias or proxy-address collections'}$relationships=@($model.model.relationships)
if($relationships.Count-ne16){throw "Expected exactly 16 validated relationships; found $($relationships.Count)"}
$expectedRelationships=@{
    'rel_ServicePlanStates_Catalog' = @{ fromTable='M365_Licenses_UserServicePlanStates'; fromColumn='__SkuPlanKey'; toTable='M365_Licenses_ServicePlans_Catalog'; toColumn='__SkuPlanKey' }
    'rel_ServicePlanStates_Users' = @{ fromTable='M365_Licenses_UserServicePlanStates'; fromColumn='UserId'; toTable='M365_Users_Active'; toColumn='Object Id' }
    'rel_DiscoveredAppRelations_Summary' = @{ fromTable='Intune_DiscoveredApps_AppDeviceRelations'; fromColumn='__TenantAppKey'; toTable='Intune_DiscoveredApps_Summary'; toColumn='__TenantAppKey' }
    'rel_UserActivity_Users' = @{ fromTable='M365_Users_Activity'; fromColumn='UserPrincipalName'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_IntuneInventory_Users' = @{ fromTable='Intune_Devices_Inventory'; fromColumn='Primary user UPN'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_LocalSystem_Users' = @{ fromTable='Intune_Devices_LocalSystem'; fromColumn='UserPrincipalName'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_WindowsUpdate_Users' = @{ fromTable='Intune_WindowsUpdate_Status'; fromColumn='UPN'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_EXO_Users' = @{ fromTable='Exchange_EXO_Mailboxes_AllDomains'; fromColumn='UserPrincipalName'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_OnPremMailboxes_Users' = @{ fromTable='Exchange_OnPrem_Mailboxes_AllDomains'; fromColumn='UserPrincipalName'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_TeamsActivity_Users' = @{ fromTable='M365_Teams_UserActivity'; fromColumn='User Principal Name'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_TeamsGuests_Users' = @{ fromTable='M365_Teams_Guests'; fromColumn='UserId'; toTable='M365_Users_Active'; toColumn='Object Id' }
    'rel_MailboxUsage_Users' = @{ fromTable='M365_Mailbox_Usage'; fromColumn='User Principal Name'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_CopilotUsage_Users' = @{ fromTable='M365_Copilot_UserUsage'; fromColumn='UserPrincipalName'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_BackupProtected_Users' = @{ fromTable='M365_Backup_ProtectedMailboxes'; fromColumn='UserPrincipalName'; toTable='M365_Users_Active'; toColumn='User principal name' }
    'rel_BackupScope_Users' = @{ fromTable='M365_BackupPolicyScope_MailboxCoverage'; fromColumn='MemberId'; toTable='M365_Users_Active'; toColumn='Object Id' }
    'rel_HybridIssues_Users' = @{ fromTable='Exchange_HybridIdentity_Issues'; fromColumn='ObjectGUID'; toTable='M365_Users_Active'; toColumn='Object Id' }
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
foreach($tableName in @('Intune_DiscoveredApps_AppDeviceRelations','Intune_DiscoveredApps_Summary')){
    $table=$tables|Where-Object name -eq $tableName|Select-Object -First 1
    $keyColumn=$table.columns|Where-Object name -eq '__TenantAppKey'|Select-Object -First 1
    if($null-eq$keyColumn -or $keyColumn.dataType-ne'string' -or $keyColumn.isHidden-ne$true){throw "Missing hidden tenant/application key on $tableName"}
    if([string]$table.partitions.source.expression -notmatch 'Table.AddColumn\([^,]+, "__TenantAppKey"'){throw "Tenant/application key is not generated in Power Query for $tableName"}
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
    'Intune_Devices_LocalSystem' = @{ 'IsEncrypted' = 'boolean' }
    'Intune_Windows11_Readiness_Issues' = @{ 'IsBlocking' = 'boolean' }
    'M365_SPO_Sites' = @{ 'IsInactive' = 'boolean'; 'IsOrphaned' = 'boolean' }
    'M365_SPO_Tenant' = @{ 'StorageUsedTB' = 'double'; 'StorageCapacityTB' = 'double'; 'StorageUtilizationPercent' = 'double'; 'IsPartialInventory' = 'boolean' }
    'M365_Mailbox_Usage' = @{ 'Is Deleted' = 'boolean'; 'Storage Used (Byte)' = 'int64'; 'Prohibit Send/Receive Quota (Byte)' = 'int64'; 'Last Activity Date' = 'dateTime' }
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
$measureTable=$tables|Where-Object name -eq 'M365_Users_Active'|Select-Object -First 1;$measureNames=@($measureTable.measures.name);if($measureNames.Count -lt 100){throw "Expected at least 100 measures; found $($measureNames.Count)"};foreach($n in @('Enabled Users','Enabled Users Inactive 90d','M365 E3 Purchased','M365 E3 Consumed','M365 E3 Available','M365 E3 Utilization','M365 E5 Purchased','M365 E5 Consumed','M365 E5 Available','M365 E5 Utilization','M365 F3 Purchased','M365 F3 Consumed','M365 F3 Available','M365 F3 Utilization','M365 SKU Utilization','Service Plan Assignments','Users with Disabled Service Plans','Managed Devices','Intune Managed Devices','Entra Device Objects','AD Computer Objects','AD Enabled Computers','Unified Devices','Devices in All Three Sources','Device Detail Intune Managed','Device Detail Noncompliant','Device Detail Stale 30d','Device Detail Action Required','Unified Users','User Detail Enabled Users','User Detail Inactive 90d','User Detail AD Users','User Detail Entra Users','User Detail Exchange Online','User Detail Exchange On-prem','User Detail Action Required','Device Compliance Rate','Windows 11 Devices','Windows 10 Devices','Windows 11 Device Rate','Windows 10 Device Rate','Hardware Refresh Candidates','Hardware Refresh Candidate Rate','BIOS Older Than 5 Years Devices','BIOS Older Than 5 Years Rate','Mailboxes','Mailbox Storage Used GB','Mailbox Quota GB','Mailbox Storage TB','Mailbox Storage Utilization','On-prem Mailbox Storage GB','On-prem Mailbox Storage TB','Online Placement Rate','On-prem Placement Rate','Teams Storage Used GB','Teams Storage Used TB','Teams Storage Quota TB','Teams Storage Utilization','SharePoint Storage Used GB','SharePoint Storage Used TB','SharePoint Estimated Capacity TB','SharePoint Estimated Utilization','AD Enabled Users Weekly Snapshot','Migration Failed Items','Protected Mailboxes','Hybrid Identity Affected Users','Model Device Count','Devices per Application','Selected Source Tables','Potential Dual-hosted SMTP','Exchange Servers')){if($measureNames -notcontains $n){throw "Missing measure: $n"}}
$tableMap=@{};foreach($t in $tables){$tableMap[$t.name]=@($t.columns.name)};$measureSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase);foreach($n in $measureNames){$null=$measureSet.Add($n)}
foreach($m in $measureTable.measures){$matches=[regex]::Matches([string]$m.expression,"'([^']+)'\[([^\]]+)\]");foreach($x in $matches){$tn=$x.Groups[1].Value;$cn=$x.Groups[2].Value;if(-not $tableMap.ContainsKey($tn)){throw "Measure $($m.name) references missing table $tn"};if($tableMap[$tn] -notcontains $cn){throw "Measure $($m.name) references missing field $tn[$cn]"}}}
$invalidMeasureFormats=@($measureTable.measures|Where-Object { [string]$_.formatString -match '\[>=' })
if($invalidMeasureFormats.Count-gt0){throw "Unsupported conditional measure formats: $($invalidMeasureFormats.name -join ', ')"}
foreach($sku in @('E3','E5','F3')){$utilization=$measureTable.measures|Where-Object name -eq "M365 $sku Utilization"|Select-Object -First 1;if([string]$utilization.expression-notmatch 'BLANK\(\)'){throw "M365 $sku Utilization must be blank when no licenses are purchased"}}$pagesMeta=Read-Json (Join-Path $definition 'pages\pages.json');if(@($pagesMeta.pageOrder).Count-ne19){throw "Expected 19 report pages; found $(@($pagesMeta.pageOrder).Count)"}
$visualFiles=@(Get-ChildItem -LiteralPath (Join-Path $definition 'pages') -Filter visual.json -File -Recurse);if($visualFiles.Count-ne501){throw "Expected exactly 501 visuals; found $($visualFiles.Count)"}
$visualObjects=@($visualFiles|ForEach-Object{Read-Json $_.FullName})
$logoVisuals=@($visualFiles|ForEach-Object{Read-Json $_.FullName}|Where-Object{$_.visual.visualType-eq'image'})
if($logoVisuals.Count-ne19){throw "Expected one logo image per page; found $($logoVisuals.Count)"}
foreach($logo in $logoVisuals){$url=[string]$logo.visual.objects.general[0].properties.imageUrl.expr.Literal.Value;if(-not$url.StartsWith("'data:image/png;base64,")){throw 'Logo image is not embedded as PNG data'}}
$slicers=@($visualFiles|ForEach-Object{Read-Json $_.FullName}|Where-Object{$_.visual.visualType-eq'slicer'})
if($slicers.Count-lt39){throw 'Expected country slicers plus operational slicers'};foreach($slicer in $slicers){$headerVisible=[string]$slicer.visual.objects.header[0].properties.show.expr.Literal.Value;if($headerVisible-ne'false'){throw "Technical slicer header must be hidden: $($slicer.name)"}}
$countrySlicers=@($visualObjects|Where-Object{$_.name-match'country$'})
if($countrySlicers.Count-ne19){throw "Expected one country slicer per page; found $($countrySlicers.Count)"};foreach($countrySlicer in $countrySlicers){$countryTitle=[string]$countrySlicer.visual.visualContainerObjects.title[0].properties.text.expr.Literal.Value;if($countryTitle-ne"'Country'"-or[double]$countrySlicer.position.height-lt80){throw "Country slicer layout regression: $($countrySlicer.name) title=$countryTitle height=$($countrySlicer.position.height)"}}
$cardVisuals=@($visualObjects|Where-Object{$_.visual.visualType-eq'cardVisual'})
$kpiLabels=@($visualObjects|Where-Object{$_.name-match'kpi\d+label$'})
if($cardVisuals.Count-ne161){throw "Expected 161 individual KPI cards; found $($cardVisuals.Count)"}
if($kpiLabels.Count-ne161){throw "Expected 161 wrapped KPI labels; found $($kpiLabels.Count)"}
$visualByName=@{};foreach($visualObject in $visualObjects){if($visualByName.ContainsKey([string]$visualObject.name)){throw "Duplicate visual name: $($visualObject.name)"};$visualByName[[string]$visualObject.name]=$visualObject;$right=[double]$visualObject.position.x+[double]$visualObject.position.width;$bottom=[double]$visualObject.position.y+[double]$visualObject.position.height;if($right-gt1280.001-or$bottom-gt720.001){throw "Visual outside 1280x720 canvas: $($visualObject.name) right=$right bottom=$bottom"}}
$requiredDeviceDetailVisuals=@('devicedetail000018country','devicedetaildevice','devicedetailuser','devicedetailsources','devicedetailcompliance','devicedetailownership','devicedetailtrust','devicedetailactivity','devicedetailaction','devicedetailtable')
foreach($visualName in $requiredDeviceDetailVisuals){if(-not$visualByName.ContainsKey($visualName)){throw "Missing Devices Detail visual: $visualName"}}
$deviceCountryProjection=$visualByName['devicedetail000018country'].visual.query.queryState.Values.projections[0].field.Column
if([string]$deviceCountryProjection.Expression.SourceRef.Entity-ne'DeviceDetail'-or[string]$deviceCountryProjection.Property-ne'Country'){throw 'Devices Detail country slicer must use DeviceDetail[Country]'}
$deviceTableFields=@($visualByName['devicedetailtable'].visual.query.queryState.Values.projections|ForEach-Object{[string]$_.field.Column.Property})
foreach($fieldName in @('DeviceName','PrimaryUserUPN','Country','SourceCoverage','ActionSeverity','RecommendedAction','ComplianceStatus','LastActivityDateTime','Windows11Readiness','UpdateRisk')){if($deviceTableFields-notcontains$fieldName){throw "Devices Detail table is missing operator field: $fieldName"}}$requiredUserDetailVisuals=@('userdetail000019country','userdetailuser','userdetaildepartment','userdetailaccount','userdetailtype','userdetailsources','userdetailplacement','userdetaillicense','userdetailactivity','userdetailaction','userdetailtable')
foreach($visualName in $requiredUserDetailVisuals){if(-not$visualByName.ContainsKey($visualName)){throw "Missing Users Detail visual: $visualName"}}
$userCountryProjection=$visualByName['userdetail000019country'].visual.query.queryState.Values.projections[0].field.Column
if([string]$userCountryProjection.Expression.SourceRef.Entity-ne'UserDetail'-or[string]$userCountryProjection.Property-ne'Country'){throw 'Users Detail country slicer must use UserDetail[Country]'}
$userTableFields=@($visualByName['userdetailtable'].visual.query.queryState.Values.projections|ForEach-Object{[string]$_.field.Column.Property})
foreach($fieldName in @('DisplayName','UserPrincipalName','PrimarySmtpAddress','Country','SourceCoverage','IdentityMatchStatus','AccountStatus','LicenseSummary','MailboxPlacement','LastActivityDateTime','ActionSeverity','RecommendedAction')){if($userTableFields-notcontains$fieldName){throw "Users Detail table is missing operator field: $fieldName"}}foreach($kpiCard in $cardVisuals){if([string]$kpiCard.name-notmatch'kpi\d+$'){throw "Legacy multi-measure card remains: $($kpiCard.name)"};$projections=@($kpiCard.visual.query.queryState.Data.projections);if($projections.Count-ne1){throw "KPI card must expose exactly one measure: $($kpiCard.name)"};$measureName=[string]$projections[0].field.Measure.Property;$builtInLabel=[string]$kpiCard.visual.objects.label[0].properties.show.expr.Literal.Value;if($builtInLabel-ne'false'){throw "KPI built-in label must be hidden: $($kpiCard.name)"};$labelName=[string]$kpiCard.name+'label';if(-not$visualByName.ContainsKey($labelName)){throw "Missing wrapped KPI label: $labelName"};$label=$visualByName[$labelName];$labelText=[string]$label.visual.objects.general[0].properties.paragraphs[0].textRuns[0].value;if($labelText-ne$measureName){throw "KPI label mismatch: $($kpiCard.name) expected '$measureName', found '$labelText'"}}
foreach($kpiCard in $cardVisuals){$units=[string]$kpiCard.visual.objects.value[0].properties.displayUnits.expr.Literal.Value;$decimals=[string]$kpiCard.visual.objects.value[0].properties.decimalPlaces.expr.Literal.Value;if($units-ne'0D'-or$decimals-ne'1D'){throw "KPI numeric formatting regression: $($kpiCard.name) displayUnits=$units decimalPlaces=$decimals"}}
$barVisuals=@($visualObjects|Where-Object{$_.visual.visualType-eq'barChart'})
foreach($barVisual in $barVisuals){$units=[string]$barVisual.visual.objects.labels[0].properties.labelDisplayUnits.expr.Literal.Value;$precision=[string]$barVisual.visual.objects.labels[0].properties.labelPrecision.expr.Literal.Value;if($units-ne'0D'-or$precision-ne'1D'){throw "Bar label formatting regression: $($barVisual.name) displayUnits=$units precision=$precision"}}foreach($vf in $visualFiles){$v=Read-Json $vf.FullName;Visit-Node $v {param($node);foreach($kind in @('Column','Measure')){$exact=$node.PSObject.Properties[$kind];if($null -ne $exact){$entry=$exact.Value;$entity=$entry.Expression.SourceRef.Entity;$property=$entry.Property;if(-not $entity -or -not $property){throw "Incomplete $kind reference in $($vf.FullName)"};if(-not $tableMap.ContainsKey([string]$entity)){throw "Visual references missing table $entity"};if($kind -eq 'Column' -and $tableMap[[string]$entity] -notcontains [string]$property){throw "Visual references missing field $entity.$property"};if($kind -eq 'Measure' -and -not $measureSet.Contains([string]$property)){throw "Visual references missing measure $property"}}}}}
foreach($n in @('fnGetSourceFiles','fnLoadSourceTable','fnToLogical','fnToDateTime','fnToNumber','fnToInt64')){if(-not ($model.model.expressions|Where-Object name -eq $n)){throw "Missing Power Query expression: $n"}}
$invalidDaxMeasures=@($tables.measures|Where-Object { [string]$_.expression -match '&&\s*VAR\b' })
if($invalidDaxMeasures.Count -gt 0){throw "Invalid DAX variable placement in measures: $($invalidDaxMeasures.name -join ', ')"}
$loaderExpression=[string](($model.model.expressions|Where-Object name -eq 'fnLoadSourceTable'|Select-Object -First 1).expression)
if(([regex]::Matches($loaderExpression,'\(')).Count -ne ([regex]::Matches($loaderExpression,'\)')).Count){throw 'Unbalanced parentheses in fnLoadSourceTable'}
$forbidden=@('SmartWorkplaceCMDB','LocalDateTable_','C:\Users\');$publishable=@(Get-ChildItem -LiteralPath $DashboardRoot -File -Recurse|Where-Object{$_.FullName -notmatch '\\\.pbi\\' -and $_.Name -ne 'LOCAL_MEMORY.md' -and $_.FullName -ne $PSCommandPath -and $_.Extension -in @('.js','.ps1','.json','.bim','.pbip','.pbir','.pbism','.pq','.md')});foreach($p in $publishable){$text=Get-Content -LiteralPath $p.FullName -Raw;foreach($term in $forbidden){if($text.Contains($term,[StringComparison]::OrdinalIgnoreCase)){throw "Forbidden dependency '$term' in $($p.FullName)"}}}
$builder=Join-Path $DashboardRoot 'scripts\Build-SmartWorkplaceDashboard.js';& node --check $builder;if($LASTEXITCODE-ne0){throw 'Builder JavaScript syntax check failed'}
& node (Join-Path $DashboardRoot 'scripts/Test-SmartWorkplaceDashboard.cjs') $DashboardRoot
if($LASTEXITCODE-ne0){throw 'Beta generated-code regression contracts failed'}
Write-Host "OK SmartWorkplaceDashboard: 90 source tables, 2 derived tables, $($measureNames.Count) measures, 19 pages, $($visualFiles.Count) visuals, 161 readable KPI labels, StateCode and discovered-app relationships validated."


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAoNq2djV5HSqRE
# SZRtlOf/tndMH2Msh3TUU9A4zfr09KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDj48WgoOCwzTifUNRFPmBy
# C/Kyu5m/QOZ426ZNQboBPTANBgkqhkiG9w0BAQEFAASCAYACdH0onwozEhDsiu7M
# pnFFyZgYvhYU8C5pfCVewoG1+hwu+zGWgBWchyp++OM+4vgECKMcc1qZc3+BwG9x
# 1pJ7OZ4lrHYMkGBYom8kWJcpvbM4FNVYq7ytcgNx2EURxK56apyBcjiTbjjafGGV
# tCY9AaHAjDYXRT9qGE312Zqapiuptg6W6jbYk/bcZXNzjbSQfD2R5T4xgokK9Wxs
# PKeCSa7zpk5KMFB47Aujf+Kfe7EsTGX5n6uS13zEjDskxeZdREpwcFiWXEEn3Y6N
# hMPcG6gpnq2U4ttmQIwwil5atEL7ZKxxQAdVyT9HLE/IY2IT6VZvc7gh0D4F8Y9c
# Ve7QzdliMQBn/HH6fKctPGLOFtH6z7/5F4b5EMqDkwuqE54Ut3x3gJHAwWrkNfdz
# rBlAw/CzaNqv/TwLAZ1YXXbSh0hpyNxPnnI97eUMZ6u1knq6shzMBCtcldw5Ane3
# MOJkvpt66Xq9ZqkfRrTm12m6Yy6ZBuAgCdztCA1FZYiyFeo=
# SIG # End signature block
