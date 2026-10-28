# Offline adapters for reported Intune scope. No network, tenant bootstrap or publication.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:IntuneContractPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'config/intune-operational-contract.json.txt'

function Read-IntuneOperationalJson {
    param([string]$Path)
    $arguments = @{AsHashtable=$true}
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $arguments.DateKind='String' }
    Get-Content -LiteralPath $Path -Raw -Encoding utf8 | ConvertFrom-Json @arguments
}

function Get-IntuneOperationalContract { Read-IntuneOperationalJson $script:IntuneContractPath }

function New-IntuneEvidenceKey {
    param([string[]]$Values)
    ($Values | ForEach-Object { ([string]$_.Length)+':'+$_ }) -join '|'
}

function ConvertTo-IntuneEvidenceTime {
    param([AllowEmptyString()][string]$Value, [switch]$SourceFieldUtc,
        [datetimeoffset]$AsOfUtc=[datetimeoffset]::UtcNow)
    $result=[ordered]@{UTC='';Basis='Missing'}
    if ([string]::IsNullOrWhiteSpace($Value)) { return [pscustomobject]$result }
    $explicit=$Value -match '^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$'
    $offsetless=$Value -match '^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?$'
    if (-not $explicit -and -not $offsetless) { $result.Basis='Invalid'; return [pscustomobject]$result }
    $candidate=if($offsetless){$Value+'Z'}else{$Value}
    $date=[datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse($candidate,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$date)) {
        $result.Basis='Invalid'; return [pscustomobject]$result
    }
    if ($offsetless -and -not $SourceFieldUtc) { $result.Basis='Time zone not supplied'; return [pscustomobject]$result }
    if ($date -gt $AsOfUtc) { $result.Basis='Future'; return [pscustomobject]$result }
    $result.UTC=$date.UtcDateTime.ToString('O')
    $result.Basis=if($explicit){'Explicit offset'}else{'Source field UTC (offsetless)'}
    [pscustomobject]$result
}

function Get-IntuneOperationalSourceGroups {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LastRoot, [string]$TenantKey,
        [datetimeoffset]$AsOfUtc=[datetimeoffset]::UtcNow,
        [ValidateRange(1,8760)][int]$MaxSourceAgeHours=168, [hashtable]$AgeOverrides=@{})
    foreach ($group in (Get-IntuneOperationalContract).groups) {
        $receiptPath=Join-Path $LastRoot $group.receipt
        if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) {
            $legacy=$receiptPath.Substring(0,$receiptPath.Length-4)
            if (Test-Path -LiteralPath $legacy -PathType Leaf) { $receiptPath=$legacy }
        }
        $present=@($group.files | Where-Object { Test-Path -LiteralPath (Join-Path $LastRoot $_.file) -PathType Leaf })
        $exists=Test-Path -LiteralPath $receiptPath -PathType Leaf
        if (-not $exists -and $present.Count -eq 0) {
            [pscustomobject]@{Spec=$group;Available=$false;Receipt=$null;ReceiptPath='';ReceiptSHA256='';Files=@();BusinessRunId=''}
            continue
        }
        if (-not $exists -or $present.Count -ne $group.files.Count) { throw "Incomplete Intune source group: $($group.name). No partial group is admitted." }
        $receipt=Read-IntuneOperationalJson $receiptPath
        if ($receipt.Owner -ne 'SmartInventory-SourceReceipt' -or $receipt.ContractVersion -ne '1.2' -or
            $receipt.Producer -ne $group.producer -or $receipt.Status -ne 'Completed' -or
            $receipt.Errors -ne 0 -or [string]::IsNullOrWhiteSpace($receipt.RunId) -or
            [string]::IsNullOrWhiteSpace($receipt.TenantKey) -or ($TenantKey -and $receipt.TenantKey -ne $TenantKey)) {
            throw "Incompatible Intune receipt: $($group.name)."
        }
        foreach($name in 'TenantId','OrganizationKey','EnvironmentKey'){
            if(-not $receipt.ContainsKey($name) -or [string]::IsNullOrWhiteSpace([string]$receipt[$name])){throw "Missing receipt identity field: $name."}
        }
        $version=[version]'0.0'
        if (-not [version]::TryParse([string]$receipt.ScriptVersion,[ref]$version) -or $version -lt [version]$group.minVersion) { throw "Intune producer version below qualified contract: $($group.name)." }
        foreach ($flag in 'FullInventoryQualified','ConsumerScopeQualified') {
            if (-not $receipt.ContainsKey($flag) -or $receipt[$flag] -isnot [bool]) { throw "Missing scope qualification: $flag." }
        }
        $completed=ConvertTo-IntuneEvidenceTime ([string]$receipt.CompletedAtUtc) -AsOfUtc $AsOfUtc
        $started=ConvertTo-IntuneEvidenceTime ([string]$receipt.StartedAtUtc) -AsOfUtc $AsOfUtc
        if (-not $completed.UTC -or -not $started.UTC -or [datetimeoffset]$started.UTC -gt [datetimeoffset]$completed.UTC) { throw "Unqualified receipt timestamps: $($group.name)." }
        $expected=@($group.files.file | Sort-Object)
        $required=@($receipt.RequiredFiles | Sort-Object)
        if ($required.Count -ne $expected.Count -or @(Compare-Object -ReferenceObject $expected -DifferenceObject $required).Count -ne 0 -or @($receipt.Files).Count -ne $expected.Count) { throw "Unexpected receipt file set: $($group.name)." }
        $files=[Collections.Generic.List[object]]::new()
        foreach ($file in $group.files) {
            $entry=@($receipt.Files | Where-Object { $_.File -eq $file.file })
            if ($entry.Count -ne 1) { throw "Duplicate or absent receipt file: $($file.file)." }
            $entry=$entry[0]
            $rows=0L
            if ($entry.Status -ne 'Success' -or $entry.Errors -ne 0 -or $entry.Delimiter -ne ',' -or $entry.RunId -ne $receipt.RunId -or
                $entry.Producer -ne $receipt.Producer -or $entry.ScriptVersion -ne $receipt.ScriptVersion -or
                $entry.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or -not [long]::TryParse([string]$entry.Rows,[ref]$rows) -or $rows -lt 0) { throw "Unqualified receipt file: $($file.file)." }
            $limit=if($AgeOverrides.ContainsKey($file.file)){[int]$AgeOverrides[$file.file]}else{$MaxSourceAgeHours}
            if ($limit -lt 1 -or $limit -gt 8760) { throw "Invalid Intune receipt-age limit: $($file.file)." }
            # Optional group: an old complete collection is still the latest available data. It is used
            # and reported; only incomplete or unqualified groups are refused.
            $receiptAge=($AsOfUtc-[datetimeoffset]$completed.UTC).TotalHours
            if ($receiptAge -gt $limit) { Write-Warning ("Intune source group used although stale: {0} ({1}). Older than its {2} h target (age {3:N1} h); used as the latest available data." -f $group.name, $file.file, $limit, $receiptAge) }
            $files.Add([pscustomobject]@{Spec=$file;Path=(Join-Path $LastRoot $file.file);Rows=$rows;SHA256=$entry.SHA256;Stats=$null})
        }
        [pscustomobject]@{Spec=$group;Available=$true;Receipt=$receipt;ReceiptPath=$receiptPath;ReceiptSHA256=(Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash;Files=$files.ToArray();BusinessRunId=''}
    }
}

function Invoke-IntuneCsvRows {
    param([string]$Path, [string[]]$RequiredColumns, [scriptblock]$Action,[hashtable]$Profile)
    $parser=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path,[Text.Encoding]::UTF8,$true)
    try {
        $parser.SetDelimiters(','); $parser.HasFieldsEnclosedInQuotes=$true; $parser.TrimWhiteSpace=$false
        $header=$parser.ReadFields()
        if (-not $header -or @($header | Sort-Object -Unique).Count -ne $header.Count -or @($header | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count) { throw "Invalid CSV header: $([IO.Path]::GetFileName($Path))." }
        if (@($RequiredColumns | Where-Object { $_ -notin $header }).Count) { throw "Intune source schema mismatch: $([IO.Path]::GetFileName($Path))." }
        if($null -ne $Profile){$Profile.Columns=$header.Count}
        $ordinal=0L
        while (-not $parser.EndOfData) {
            $values=$parser.ReadFields(); $ordinal++
            if ($values.Count -ne $header.Count) { throw "Malformed Intune CSV record: $ordinal." }
            $row=@{}
            for($index=0;$index -lt $header.Count;$index++){ $row[$header[$index]]=$values[$index] }
            $null=& $Action $row $ordinal
        }
        $ordinal
    } finally { $parser.Dispose() }
}

function Test-IntuneOperationalSnapshot {
    param([Parameter(Mandatory)][string]$LastRoot, [Parameter(Mandatory)][string]$TenantKey,
        [datetimeoffset]$AsOfUtc=[datetimeoffset]::UtcNow,
        [ValidateRange(1,8760)][int]$MaxSourceAgeHours=168, [hashtable]$AgeOverrides=@{})
    $groups=@(Get-IntuneOperationalSourceGroups -LastRoot $LastRoot -TenantKey $TenantKey -AsOfUtc $AsOfUtc -MaxSourceAgeHours $MaxSourceAgeHours -AgeOverrides $AgeOverrides)
    foreach ($group in $groups | Where-Object Available) {
        $businessRuns=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $collectionTimes=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $timeCache=@{}
        foreach ($file in $group.Files) {
            if ((Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash -ne $file.SHA256) { throw "Intune source/receipt hash mismatch: $($file.Spec.file)." }
            $keys=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $stats=@{EarliestCollectionUTC='';LatestCollectionUTC='';Rows=0L;Columns=$file.Spec.columns.Count}
            $count=Invoke-IntuneCsvRows $file.Path $file.Spec.columns {
                param($row,$ordinal)
                if ($row.TenantKey -ne $TenantKey -or [string]::IsNullOrWhiteSpace($row.RunId)) { throw "Tenant or business run mismatch: $($file.Spec.file), record $ordinal." }
                foreach($name in 'TenantId','OrganizationKey','EnvironmentKey'){if($row[$name] -ne $group.Receipt[$name]){throw "Source/receipt identity mismatch: $($file.Spec.file), record $ordinal."}}
                $null=$businessRuns.Add($row.RunId)
                if ($businessRuns.Count -gt 1) { throw "Mixed business runs: $($group.Spec.name)." }
                if(-not $timeCache.ContainsKey($row.CollectedAtUtc)){$timeCache[$row.CollectedAtUtc]=ConvertTo-IntuneEvidenceTime $row.CollectedAtUtc -AsOfUtc $AsOfUtc}
                $time=$timeCache[$row.CollectedAtUtc]
                if (-not $time.UTC) { throw "Unqualified collection timestamp: $($file.Spec.file), record $ordinal." }
                $null=$collectionTimes.Add($time.UTC)
                if($collectionTimes.Count -gt 1){throw 'Mixed runtime collection timestamps in Intune source group.'}
                if ([datetimeoffset]$time.UTC -gt [datetimeoffset]$group.Receipt.CompletedAtUtc) { throw 'Collection observation after receipt completion.' }
                if (-not $stats.EarliestCollectionUTC -or $time.UTC -lt $stats.EarliestCollectionUTC) {$stats.EarliestCollectionUTC=$time.UTC}
                if (-not $stats.LatestCollectionUTC -or $time.UTC -gt $stats.LatestCollectionUTC) {$stats.LatestCollectionUTC=$time.UTC}
                if ($file.Spec.key.Count) {
                    $values=@($file.Spec.key | ForEach-Object { [string]$row[$_] })
                    if (@($values | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -or -not $keys.Add((New-IntuneEvidenceKey $values))) { throw "Missing or duplicate Intune key: $($file.Spec.file), record $ordinal." }
                }
                if ($file.Spec.file -like 'Intune_Configuration*States.csv' -and ([string]::IsNullOrWhiteSpace($row.PolicyId) -or [string]::IsNullOrWhiteSpace($row.IntuneDeviceId))) { throw 'Configuration observation requires technical policy and device IDs.' }
                if ($file.Spec.file -eq 'Intune_ManagedAppDeviceStates.csv' -and [string]::IsNullOrWhiteSpace($row.ApplicationId)) { throw 'Application observation requires ApplicationId.' }
            } -Profile $stats
            if ($count -ne $file.Rows) { throw "Receipt row count mismatch: $($file.Spec.file)." }
            $stats.Rows=$count; $file.Stats=$stats
            if ((Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash -ne $file.SHA256) { throw "Intune source changed while reading: $($file.Spec.file)." }
        }
        $group.BusinessRunId=if($businessRuns.Count){@($businessRuns)[0]}else{''}
        if ((Get-FileHash -LiteralPath $group.ReceiptPath -Algorithm SHA256).Hash -ne $group.ReceiptSHA256) { throw 'Intune receipt changed while validating group.' }
    }
    $groups
}

function Write-IntuneCsvRecord {
    param([IO.StreamWriter]$Writer,[string[]]$Columns,[Collections.IDictionary]$Row)
    $fields=foreach($column in $Columns){ '"'+([string]$Row[$column]).Replace('"','""')+'"' }
    $Writer.WriteLine(($fields -join ','))
}

function New-IntuneOperationalEvidence {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LastRoot,[Parameter(Mandatory)][string]$OutputRoot,
        [Parameter(Mandatory)][string]$TenantKey,[datetimeoffset]$AsOfUtc=[datetimeoffset]::UtcNow,
        [ValidateRange(1,8760)][int]$MaxSourceAgeHours=168,[hashtable]$AgeOverrides=@{})
    $source=[IO.Path]::GetFullPath($LastRoot).TrimEnd('\','/')
    $destination=[IO.Path]::GetFullPath($OutputRoot).TrimEnd('\','/')
    $dataRoot=Split-Path -Parent $source
    if ($destination -eq $dataRoot -or $destination.StartsWith($dataRoot+'\',[StringComparison]::OrdinalIgnoreCase) -or $source.StartsWith($destination+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Intune outputs must be isolated staging, outside the raw data tree.' }
    # Validate all governed source bytes and rows before creating any derived output.
    $groups=@(Test-IntuneOperationalSnapshot -LastRoot $source -TenantKey $TenantKey -AsOfUtc $AsOfUtc -MaxSourceAgeHours $MaxSourceAgeHours -AgeOverrides $AgeOverrides)
    $contract=Get-IntuneOperationalContract
    $inventory=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $ambiguous=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $inventoryPath=Join-Path $source 'Intune_Devices_Inventory.csv'
    if (@($groups | Where-Object Available).Count -gt 0) {
        if (-not (Test-Path -LiteralPath $inventoryPath -PathType Leaf)) { throw 'Device reference inventory is required for operational evidence.' }
        $null=Invoke-IntuneCsvRows $inventoryPath @('TenantKey','Device ID') {
            param($row,$ordinal)
            if ($row.TenantKey -ne $TenantKey) { throw 'Device reference tenant mismatch.' }
            if (-not [string]::IsNullOrWhiteSpace($row['Device ID']) -and -not $inventory.Add($row['Device ID'])) { $null=$ambiguous.Add($row['Device ID']) }
        }
    }
    $catalogs=@{Configurations=@{};ManagedApps=@{}}
    $inputFiles=@{}
    foreach($group in $groups | Where-Object Available){foreach($file in $group.Files){$inputFiles[$file.Spec.file]=@{Group=$group;File=$file}}}
    foreach($name in 'Intune_ConfigurationPolicies.csv','Intune_ManagedApps.csv'){
        if ($inputFiles.ContainsKey($name)) {
            $context=$inputFiles[$name]; $catalog=$catalogs[$context.Group.Spec.name]
            $null=Invoke-IntuneCsvRows $context.File.Path $context.File.Spec.columns {
                param($row,$ordinal)
                $id=if($row.ContainsKey('PolicyId')){$row.PolicyId}else{$row.ApplicationId}
                $catalog[$id]=$row
            }
        }
    }
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    $aggregate=@{}; $summary=@{}; $assigned=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $configCounts=@{Current=0L;Outside=0L}
    foreach($table in $contract.tables | Where-Object { $_.ContainsKey('source') }) {
        $path=Join-Path $destination $table.file
        if (Test-Path -LiteralPath $path) { throw "Intune staging output already exists: $($table.file)." }
        $writer=[IO.StreamWriter]::new($path,$false,[Text.UTF8Encoding]::new($false))
        try {
            $header=@{}; foreach($name in $table.columns.name){$header[$name]=$name}
            Write-IntuneCsvRecord $writer $table.columns.name $header
            if (-not $inputFiles.ContainsKey($table.source)) { continue }
            $context=$inputFiles[$table.source]; $file=$context.File; $group=$context.Group
            $null=Invoke-IntuneCsvRows $file.Path $file.Spec.columns {
                param($row,$ordinal)
                $output=@{}
                foreach($column in $table.columns.name){if($row.ContainsKey($column)){$output[$column]=$row[$column]}}
                $sourceKey=if($file.Spec.key.Count){@($file.Spec.key | ForEach-Object {[string]$row[$_]})}else{@($TenantKey,$row.RunId,$file.SHA256,[string]$ordinal)}
                $output.EvidenceKey=New-IntuneEvidenceKey $sourceKey
                $output.ReceiptRunId=$group.Receipt.RunId; $output.BusinessRunId=$row.RunId
                $output.SourceFile=$file.Spec.file; $output.SourceSHA256=$file.SHA256; $output.SourceRecord=$ordinal
                $output.CollectionUTC=$file.Stats.LatestCollectionUTC
                $output.HealthQualification='Reported deployment is not antivirus or EDR health'
                $id=if($row.ContainsKey('PolicyId')){$row.PolicyId}elseif($row.ContainsKey('ApplicationId')){$row.ApplicationId}else{''}
                if ($id) { $output.CatalogMatch=if($catalogs[$group.Spec.name].ContainsKey($id)){'Matched'}else{'Outside collected catalog'} }
                $device=if($row.ContainsKey('IntuneDeviceId')){$row.IntuneDeviceId}elseif($row.ContainsKey('ManagedDeviceIdentifier')){if($row.HasManagedDeviceIdentifier -eq 'True' -and $row.ManagedDeviceIdentifier -ne '00000000-0000-0000-0000-000000000000'){$row.ManagedDeviceIdentifier}else{''}}elseif($row.ContainsKey('DeviceId')){$row.DeviceId}else{''}
                $output.InventoryMatch=if(-not $device){'Identifier not supplied'}elseif($ambiguous.Contains($device)){'Ambiguous reference identifier'}elseif($inventory.Contains($device)){'Matched technical identifier'}else{'Unmatched technical identifier'}
                if($table.source -eq 'Intune_ConfigurationDeviceCurrentStates.csv'){
                    $date=ConvertTo-IntuneEvidenceTime $row.LastModifiedTimeUtc -SourceFieldUtc -AsOfUtc $AsOfUtc
                    $output.SourceStatusUTC=$date.UTC; $output.SourceStatusTimeBasis=$date.Basis
                    if($row.ReportedApplicationState -notin @('ReportedApplied','ReportedIssue','Pending','NotApplicable','Unknown')){$output.ReportedApplicationState='Unknown'}
                    $configCounts.Current++; if($output.CatalogMatch -ne 'Matched'){$configCounts.Outside++}
                }
                # Producer scope is apps with assignment records, including exclusions;
                # this is not proof of an effective required-install target.
                if($table.source -eq 'Intune_ManagedAppAssignments.csv'){$null=$assigned.Add($id)}
                if($table.source -eq 'Intune_ManagedAppDeviceStates.csv'){
                    $date=ConvertTo-IntuneEvidenceTime $row.LastModifiedDateTime -AsOfUtc $AsOfUtc
                    $output.SourceStatusUTC=$date.UTC; $output.SourceStatusTimeBasis=$date.Basis
                    if($row.ReportedDeploymentState -notin @('ReportedInstalled','ReportedFailure','ReportedNotInstalled','Pending','NotApplicable','Unknown')){$output.ReportedDeploymentState='Unknown'}
                    if(-not $aggregate.ContainsKey($id)){$aggregate[$id]=@{ObservedRows=0L;ReportedInstalledRows=0L;ReportedFailureRows=0L;ReportedNotInstalledRows=0L;PendingRows=0L;NotApplicableRows=0L;UnknownRows=0L}}
                    $aggregate[$id].ObservedRows++; $aggregate[$id][$output.ReportedDeploymentState+'Rows']++
                }
                if($table.source -eq 'Intune_ManagedAppInstallSummary.csv'){
                    $summary[$id]=$row
                    if($row.StatusSource -notin @('DeviceInstallStatusByApp','NoReportRows','NotAssigned')){throw 'Unknown managed-app summary source.'}
                    $output.MeasurementState=if($row.StatusSource -eq 'DeviceInstallStatusByApp'){'Reported observations only'}else{'Not measured'}
                    if($catalogs.ManagedApps.ContainsKey($id)){foreach($name in 'DisplayName','Publisher','AppType','DisplayVersion'){$output[$name]=$catalogs.ManagedApps[$id][$name]}}
                }
                if($table.source -eq 'Intune_EnrollmentFailures.csv'){
                    $date=ConvertTo-IntuneEvidenceTime $row.EventDateTime -AsOfUtc $AsOfUtc
                    $output.SourceStatusUTC=$date.UTC; $output.SourceStatusTimeBasis=$date.Basis
                }
                Write-IntuneCsvRecord $writer $table.columns.name $output
            }
        } finally { $writer.Dispose() }
    }
    foreach($id in $summary.Keys){
        $row=$summary[$id]; $total=0L
        foreach($name in 'ReportedInstalledRows','ReportedFailureRows','ReportedNotInstalledRows','PendingRows','NotApplicableRows','UnknownRows'){
            $n=0L; if(-not [long]::TryParse($row[$name],[ref]$n) -or $n -lt 0){throw 'Invalid managed-app summary count.'}; $total+=$n
            $observed=if($aggregate.ContainsKey($id)){$aggregate[$id][$name]}else{0L}
            if($n -ne $observed){throw 'Managed-app summary/deployment reconciliation failed.'}
        }
        if([string]$total -ne $row.ObservedRows -or ($row.StatusSource -eq 'DeviceInstallStatusByApp' -and $total -eq 0) -or ($row.StatusSource -ne 'DeviceInstallStatusByApp' -and $total -ne 0)){throw 'Managed-app measurement state contradicts report counts.'}
        if(($row.StatusSource -eq 'NotAssigned' -and $assigned.Contains($id)) -or ($row.StatusSource -ne 'NotAssigned' -and -not $assigned.Contains($id))){throw 'Managed-app assignment/status-source reconciliation failed.'}
    }
    foreach($id in $aggregate.Keys){if(-not $summary.ContainsKey($id)){throw 'Application observations missing from summary.'}}
    foreach($id in $catalogs.ManagedApps.Keys){if(-not $summary.ContainsKey($id)){throw 'Catalog application missing from summary.'}}
    if($inputFiles.ContainsKey('Intune_ConfigurationStatusCoverage.csv')){
        $coverage=$inputFiles['Intune_ConfigurationStatusCoverage.csv']
        if($coverage.File.Rows -ne 1){throw 'Configuration coverage must contain exactly one observation.'}
        $null=Invoke-IntuneCsvRows $coverage.File.Path $coverage.File.Spec.columns {
            param($row,$ordinal)
            foreach($pair in @(@('AssignmentStatusRows','Intune_ConfigurationAssignmentStates.csv'),@('DeviceStatusHistoryRows','Intune_ConfigurationDeviceStates.csv'),@('DeviceStatusCurrentRows','Intune_ConfigurationDeviceCurrentStates.csv'))){if([string]$inputFiles[$pair[1]].File.Rows -ne $row[$pair[0]]){throw 'Configuration coverage/source count mismatch.'}}
            if([string]$configCounts.Current -ne $row.DeviceStatusCurrentRows -or [string]$configCounts.Outside -ne $row.CurrentStatesOutsideCollectedCatalog){throw 'Configuration current/catalog coverage mismatch.'}
        }
    }
    $coverageTable=$contract.tables | Where-Object file -eq 'IntuneOperationalSourceEvidence.csv'
    $sourceRows=foreach($group in $groups){
        $r=$group.Receipt
        @{TenantKey=$TenantKey;EvidenceKey=(New-IntuneEvidenceKey @($TenantKey,$group.Spec.name));SourceGroup=$group.Spec.name;Availability=$(if($group.Available){'Receipt and files verified'}else{'Not supplied'});Producer=$group.Spec.producer;ScriptVersion=$(if($r){$r.ScriptVersion}else{''});ReceiptRunId=$(if($r){$r.RunId}else{''});BusinessRunId=$group.BusinessRunId;ReceiptSHA256=$group.ReceiptSHA256;CompletedUTC=$(if($r){(ConvertTo-IntuneEvidenceTime $r.CompletedAtUtc -AsOfUtc $AsOfUtc).UTC}else{''});ScopeQualification=$(if($r){$r.ScopeQualification}else{'Not measured'});FullInventoryQualified=$(if($r){[string]$r.FullInventoryQualified}else{'Unknown'});ConsumerScopeQualified=$(if($r){[string]$r.ConsumerScopeQualified}else{'Unknown'});FileCount=$group.Files.Count;RowCount=$(if($group.Files.Count){($group.Files | Measure-Object Rows -Sum).Sum}else{0});Qualification='Reported source scope only; no complete tenant coverage or endpoint health inference'}
    }
    $coveragePath=Join-Path $destination $coverageTable.file
    if(Test-Path -LiteralPath $coveragePath){throw 'Source coverage staging output already exists.'}
    $writer=[IO.StreamWriter]::new($coveragePath,$false,[Text.UTF8Encoding]::new($false))
    try{ $header=@{};foreach($name in $coverageTable.columns.name){$header[$name]=$name};Write-IntuneCsvRecord $writer $coverageTable.columns.name $header;foreach($row in $sourceRows){Write-IntuneCsvRecord $writer $coverageTable.columns.name $row} }finally{$writer.Dispose()}
    # Detect changed inputs during the second (projection) pass as well.
    foreach($group in $groups | Where-Object Available){foreach($file in $group.Files){if((Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash -ne $file.SHA256){throw 'Source changed while projecting Intune evidence.'}};if((Get-FileHash -LiteralPath $group.ReceiptPath -Algorithm SHA256).Hash -ne $group.ReceiptSHA256){throw 'Receipt changed while projecting Intune evidence.'}}
    [pscustomobject]@{Tables=$contract.tables.Count;Groups=$sourceRows;Files=@($groups | Where-Object Available | ForEach-Object {$_.Files})}
}

function ConvertFrom-IntuneOperationalAudit {
    param([Parameter(Mandatory)][Collections.IDictionary]$Audit,[datetimeoffset]$AsOfUtc=[datetimeoffset]::UtcNow)
    $spec=Get-IntuneOperationalContract
    foreach($group in $Audit.Groups){
        $domain=if($group.SourceGroup -eq 'ManagedApps'){'Applications'}else{'Endpoint'}
        $definition=@($spec.groups | Where-Object name -eq $group.SourceGroup)
        if($definition.Count -ne 1){throw 'Unknown audited Intune source group.'}
        $files=@($Audit.Files | Where-Object {$_.Spec.file -in $definition[0].files.file})
        if($group.Availability -eq 'Not supplied'){
            [pscustomobject][ordered]@{
                'Source Key'='IntuneOperations|'+$group.SourceGroup;'Evidence Domain'=$domain;'Source Name'='Intune reported '+$group.SourceGroup;
                'File Name'='';'File State'='Missing';'Schema State'='Not evaluated';'Expected Columns'='';'Missing Columns'='';
                'Row Count'=$null;'Column Count'=$null;'Snapshot Date'='';'Snapshot Basis'='Not supplied';'Evidence Age Days'=$null;
                'Freshness State'='Unknown';'Decision Readiness'='Optional unavailable';'Evidence Status'='Not measured';
                'Recommended Action'='Supply the complete governed source group; absence is not a zero failure count.';'Read Error'=''
            }
            continue
        }
        if($files.Count -ne $definition[0].files.Count){throw 'Incomplete Intune source audit.'}
        foreach($file in $files){
            $raw=if($file.Rows -eq 0){$group.CompletedUTC}else{$file.Stats.LatestCollectionUTC}
            $time=ConvertTo-IntuneEvidenceTime $raw -AsOfUtc $AsOfUtc
            if(-not $time.UTC){throw 'Unqualified audited source timestamp.'}
            $snapshot=[datetimeoffset]$time.UTC; $age=($AsOfUtc-$snapshot).TotalDays
            [pscustomobject][ordered]@{
                'Source Key'='IntuneOperations|'+$file.Spec.file;'Evidence Domain'=$domain;
                'Source Name'='Intune reported '+$group.SourceGroup+' / '+$file.Spec.file;'File Name'=$file.Spec.file;
                'File State'=$(if($file.Rows -eq 0){'Empty allowed'}else{'Loaded'});'Schema State'='Qualified';
                'Expected Columns'=($file.Spec.columns -join ', ');'Missing Columns'='';'Row Count'=$file.Rows;'Column Count'=$file.Stats.Columns;
                'Snapshot Date'=$snapshot.ToString('yyyy-MM-dd');'Snapshot Basis'='Explicit collection UTC; receipt/hash verified';
                # Export completed days as required by Data Trust; classify using the exact age.
                'Evidence Age Days'=[long][math]::Floor($age);'Freshness State'=$(if($age -le 7){'Fresh'}elseif($age -le 30){'Aging'}else{'Stale'});
                'Decision Readiness'=$(if($age -gt 30){'Refresh required'}else{'Ready'});'Evidence Status'='Reported source scope only';
                'Recommended Action'='File readiness only; scope is '+$group.ScopeQualification+'. Do not infer complete tenant coverage or endpoint health.';'Read Error'=''
            }
        }
    }
}

Export-ModuleMember -Function Read-IntuneOperationalJson,Get-IntuneOperationalContract,Get-IntuneOperationalSourceGroups,Test-IntuneOperationalSnapshot,ConvertTo-IntuneEvidenceTime,New-IntuneEvidenceKey,New-IntuneOperationalEvidence,ConvertFrom-IntuneOperationalAudit

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCfNQf0ghUC6np7
# WcLeDQTMeZlnLn4L+tPwYTsxwlHWkqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIE8q9nAsNY7k9P1nVBg7w1QRCbOFCRMiF8UOvUlXq5ceMA0GCSqG
# SIb3DQEBAQUABIIBgJDz9iRp9tVLofsjF89ngeex9MIBRrbE/bGkbfzfr+uXPRrf
# Q4TDtyQlQSC8ak8No4nbYIPnglWQqlKnIOyGDKZ6VkoOjM+xQCt9WxBVtjo9bKnw
# H+cha+Q3ZCjJ6b7PUr+HUoSVKBPGa8sx5JvCrxhvW6x4fweNQ1CVfOIK/m8SdC0+
# moihqWkbj+FUMYbcM0seUb65fpdZJL4UrlCE69qcs6mGXc5FAZJt/AglwYiIJQ3V
# brzg//hitlR/uW6KourNtd1v6EOlY0ggVyWajJqn8e//4exym1/V9ycw8DsSDpD1
# 2G3/nxHR32bh3QETKPrSBOW850j8xM6/n0Z7m0oWfwe7AMjmuHQanaF+0HlztjCO
# hxa+Iix0h+fo7LQZUHIwh1gfuLgPvLvGh8zy2as2S2/ykzpK1HZEI5gjCytq0kyu
# wEe+jKoCDdIByFq46ieoDckXs9b6NhhjwuYCPaGIuCgIdzEhTH5e/lP17F8RCadV
# ulUyMzDP7wL2Oqt9ZKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTQxMTI3
# NTBaMC8GCSqGSIb3DQEJBDEiBCCeMW/F+8rmw1M9ZLiZ3e6tvjWVUUROqw03bCer
# 4cZKgzANBgkqhkiG9w0BAQEFAASCAgCpdfloALKLZpFeUBiuK460rt71miYKTXFq
# gRQzHvGZRGaieJYAneliw14h2Pif9xfmFWOlJZa7sqnZhkqhAtmRpHnPID+VmTg4
# VZ9SfkRfvPcrK/0He4xjoTNrIDllx74sKrUpGZpReuWKH1v7a/4vKwpVsTBRJAj0
# BJtJ+/g1EEp61YGtjcsfuwFbNDBYFvKyEDN7FNEhi3XMK812hJfCog8LNPIFnCfd
# TT1CidjdF+cZRSIYmvl2gsz6XU4L0FnNAyp9EUofdaS1G+XriUx14IJSL4WZ6PrR
# KUEtsFKnFsx9wYVijidliecY7Gg+banr1F7F5q5teBbPsB9NaGwhA/IIrhAj0E9Y
# dKxYCvgKMvChckOzDpW8yJYhq1SYUxhw5rfpYB60NwtcJrCi1yWKQ75KSDP325Zu
# 4FyAu24XgIFAE2g1LuhcSG3vsNh9PexFnVSgTNsJXKYlBOnMpsumoQo39GzjvCdB
# +bwWoQeA40qEk1mhpk7L1LQ+G9up28hQ4+AVIK7rYXsakrO/s3/OXOKANQ1ZTM+/
# 8OZTn7+5z0KSQsHSUAkFVPrNFuPhSyvweiBBR0WJSLejkBE4D4QbUKWMXoxADu8y
# Qkk3n1S8CbW5svWhPMLlXwhs75wJJopAksvnS0/JU+In0qXTAwpeBFJ277YW2ODQ
# j+EZhP4CwQ==
# SIG # End signature block
