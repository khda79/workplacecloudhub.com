[CmdletBinding()]
param()
# Synthetic fixtures only. Never reads tenant data or starts a collector/preparation pipeline.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'IntuneOperationalEvidence.psm1') -Force
$spec=Get-IntuneOperationalContract
$script:Assertions=0
function Assert([bool]$Condition,[string]$Message){$script:Assertions++;if(-not $Condition){throw "Assertion failed: $Message"}}
function Reject([scriptblock]$Action,[string]$Message){$failed=$false;try{$null=& $Action}catch{$failed=$true};Assert $failed $Message}
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('intune-operational-offline-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $testRoot
$data=Join-Path $testRoot 'source';$last=Join-Path $data 'DATA-LAST';$null=New-Item -ItemType Directory -Path $last -Force
$now=[datetimeoffset]'2026-01-02T12:00:00Z'
$jsonArgs=@{AsHashtable=$true};if((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')){$jsonArgs.DateKind='String'}
$script:Fixtures=@{}
function New-Row($File,[hashtable]$Values){
    $row=[ordered]@{};foreach($name in $File.columns){$row[$name]=''}
    $row.TenantKey='synthetic-test';$row.OrganizationKey='synthetic';$row.EnvironmentKey='test';$row.TenantId='00000000-0000-0000-0000-000000000001'
    $row.RunId='business-'+($File.file.Split('_')[1]);$row.CollectedAtUtc='2026-01-02T10:00:00Z'
    foreach($key in $Values.Keys){$row[$key]=[string]$Values[$key]}
    [pscustomobject]$row
}
function Reset-Fixtures {
    foreach($group in $spec.groups){
        foreach($file in $group.files){
            $v=@{RunId='business-'+$group.name;PolicyId='policy-1';ApplicationId='app-1';AssignmentId='assign-1';SettingId='setting-1';EventId='event-1';IntuneDeviceId='device-1';DeviceId='device-1';TargetType='groupAssignmentTarget';CollectedAtUtc='2026-01-02T10:00:00Z'}
            # Restrict values to the raw schema; do not manufacture extra source fields.
            $v2=@{};foreach($k in $v.Keys){if($k -in $file.columns){$v2[$k]=$v[$k]}}
            $row=New-Row $file $v2
            switch($file.file){
                'Intune_ConfigurationPolicies.csv'{$row.DisplayName='Synthetic policy';$row.NativePolicyJson='secret-value-must-not-project'}
                'Intune_ConfigurationSettings.csv'{$row.SettingInstanceJson='secret-setting-must-not-project'}
                'Intune_ConfigurationAssignments.csv'{$row.TargetType='exclusionGroupAssignmentTarget'}
                'Intune_ConfigurationDeviceStates.csv'{$row.LastModifiedTimeUtc='2026-01-02 09:00:00.0000000';$row.ReportedApplicationState='ReportedApplied';$row.PolicyInCollectedCatalog='True'}
                'Intune_ConfigurationDeviceCurrentStates.csv'{$row.LastModifiedTimeUtc='2026-01-02 09:00:00.0000000';$row.ReportedApplicationState='ReportedApplied';$row.PolicyInCollectedCatalog='True'}
                'Intune_ConfigurationStatusCoverage.csv'{foreach($k in @('AssignmentStatusRows','DeviceStatusHistoryRows','DeviceStatusCurrentRows')){$row.$k='1'};foreach($k in @('UnknownDeviceStatusRows','UndatedDeviceStatusRows','CurrentStatesOutsideCollectedCatalog')){$row.$k='0'};$row.AssignmentStatusReportState='Reported'}
                'Intune_ManagedApps.csv'{$row.DisplayName='Synthetic app'}
                'Intune_ManagedAppAssignments.csv'{$row.Intent='available'}
                'Intune_ManagedAppInstallSummary.csv'{foreach($k in @('ObservedRows','ReportedInstalledRows')){$row.$k='1'};foreach($k in @('ReportedFailureRows','ReportedNotInstalledRows','PendingRows','NotApplicableRows','UnknownRows')){$row.$k='0'};$row.StatusSource='DeviceInstallStatusByApp';$row.InCollectedCatalog='True'}
                'Intune_ManagedAppDeviceStates.csv'{$row.ReportedDeploymentState='ReportedInstalled';$row.LastModifiedDateTime='2026-01-02 09:00:00.0000000'}
                'Intune_EnrollmentFailures.csv'{$row.EventDateTime='2026-01-02T09:00:00Z';$row.HasManagedDeviceIdentifier='False';$row.ManagedDeviceIdentifier='00000000-0000-0000-0000-000000000000';$row.DeviceId='device-1';$row.FailureCategory='UserValidation'}
            }
            $script:Fixtures[$file.file]=@($row)
        }
    }
}
function Write-Fixtures {
    foreach($group in $spec.groups){
        $entries=foreach($file in $group.files){
            $path=Join-Path $last $file.file
            $rows=@($script:Fixtures[$file.file])
            if($rows.Count){$rows | Select-Object $file.columns | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8}
            else{($file.columns | ForEach-Object {'"'+$_+'"'}) -join ',' | Set-Content -LiteralPath $path -Encoding utf8}
            @{File=$file.file;Rows=$rows.Count;SHA256=(Get-FileHash -LiteralPath $path).Hash;Errors=0;Status='Success';Delimiter=',';RunId='receipt-'+$group.name;Producer=$group.producer;ScriptVersion=$group.minVersion}
        }
        @{Owner='SmartInventory-SourceReceipt';ContractVersion='1.2';Producer=$group.producer;ScriptVersion=$group.minVersion;Status='Completed';Errors=0;TenantKey='synthetic-test';TenantId='00000000-0000-0000-0000-000000000001';OrganizationKey='synthetic';EnvironmentKey='test';RunId='receipt-'+$group.name;StartedAtUtc='2026-01-02T10:00:00Z';CompletedAtUtc='2026-01-02T11:00:00Z';FullInventoryQualified=$false;ConsumerScopeQualified=$false;ScopeQualification='ConfiguredOutputsOnly';RequiredFiles=@($group.files.file);Files=@($entries)} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $last $group.receipt) -Encoding utf8
    }
    @([pscustomobject]@{TenantKey='synthetic-test';'Device ID'='device-1'}) | Export-Csv -LiteralPath (Join-Path $last 'Intune_Devices_Inventory.csv') -NoTypeInformation
}
function Build([string]$Suffix){New-IntuneOperationalEvidence -LastRoot $last -OutputRoot (Join-Path $testRoot $Suffix) -TenantKey 'synthetic-test' -AsOfUtc $now}
function Validate {Test-IntuneOperationalSnapshot -LastRoot $last -TenantKey 'synthetic-test' -AsOfUtc $now}
function Edit-Receipt($Group,[scriptblock]$Action){$p=Join-Path $last $Group.receipt;$r=Get-Content $p -Raw | ConvertFrom-Json -AsHashtable;& $Action $r;$r | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $p -Encoding utf8}
function Test-TrustExport($Rows,[string]$Name,[bool]$ExpectedPass=$true) {
    $root=Join-Path $testRoot ('trust-'+$Name)
    $null=New-Item -ItemType Directory -Path $root
    $Rows | Export-Csv -LiteralPath (Join-Path $root 'DataTrustEvidence.csv') -NoTypeInformation -Encoding utf8
    $report=Join-Path $root 'validation.json.txt'
    & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root $root -Tables 'Data Trust Evidence' -ReportPath $report | Out-Null
    $exitCode=$LASTEXITCODE
    $validation=Read-IntuneOperationalJson $report
    Assert ($validation.Passed -eq $ExpectedPass -and $exitCode -eq $(if($ExpectedPass){0}else{1})) ('Final Data Trust contract '+$Name+': '+($validation.Errors -join '; '))
    return $validation
}
try{
    Reset-Fixtures;Write-Fixtures
    $result=Build 'valid'
    Assert ($result.Tables -eq 9) 'Nine output tables'
    Assert (@($result.Groups).Count -eq 3) 'Three source groups'
    Assert ($result.Groups[0].ReceiptRunId -ne $result.Groups[0].BusinessRunId) 'Dual run provenance is not equated'
    Assert (@($result.Groups | Where-Object FullInventoryQualified -ne 'False').Count -eq 0) 'Source scope not promoted'
    $config=Import-Csv (Join-Path $testRoot 'valid/IntuneConfigurationDeviceEvidence.csv')
    Assert ($config.SourceStatusTimeBasis -eq 'Source field UTC (offsetless)') 'Configuration source UTC qualified'
    Assert ($config.SourceStatusUTC -eq '2026-01-02T09:00:00.0000000Z') 'Configuration invariant UTC'
    Assert ($config.InventoryMatch -eq 'Matched technical identifier') 'Exact device identifier only'
    $app=Import-Csv (Join-Path $testRoot 'valid/IntuneManagedAppDeploymentEvidence.csv')
    Assert ($app.SourceStatusUTC -eq '' -and $app.SourceStatusTimeBasis -eq 'Time zone not supplied') 'App offsetless source time remains unqualified'
    Assert ($app.HealthQualification -like '*not antivirus or EDR*') 'Deployment is not health'
    $event=Import-Csv (Join-Path $testRoot 'valid/IntuneEnrollmentFailureEvidence.csv')
    Assert ($event.InventoryMatch -eq 'Identifier not supplied') 'Event DeviceId never substituted for managed identifier'
    Assert ($event.SourceStatusUTC -ne '') 'Enrollment explicit UTC qualified'
    foreach($table in $spec.tables){
        $p=Join-Path $testRoot ('valid/'+$table.file)
        Assert (Test-Path -LiteralPath $p) ('Output '+$table.file)
        $header=Get-Content -LiteralPath $p -TotalCount 1
        Assert ($header -eq (($table.columns.name | ForEach-Object {'"'+$_+'"'}) -join ',')) ('Header contract '+$table.file)
        Assert ((Get-Content -LiteralPath $p -Raw) -notmatch 'secret-value|secret-setting|NativePolicyJson|SettingInstanceJson') ('No configuration value leak '+$table.file)
    }
    $audit=$result | ConvertTo-Json -Depth 15 | ConvertFrom-Json @jsonArgs
    $trust=@(ConvertFrom-IntuneOperationalAudit -Audit $audit -AsOfUtc $now)
    Assert ($trust.Count -eq 12) 'Data Trust reuses all twelve verified profiles'
    Assert (@($trust | Where-Object 'Evidence Status' -ne 'Reported source scope only').Count -eq 0) 'Data Trust has bounded source status'
    $null=Test-TrustExport $trust 'fractional-day'
    Assert (@($trust | Where-Object {$_.'Evidence Age Days' -isnot [long] -or $_.'Evidence Age Days' -ne 0}).Count -eq 0) 'Two-hour observations export zero completed days as Int64'
    # Integer display must not round freshness/readiness across the existing exact thresholds.
    $ageCases=@(
        @{Age=0.0;Days=0L;Freshness='Fresh';Readiness='Ready'},
        @{Age=(1.0-1.0/1440);Days=0L;Freshness='Fresh';Readiness='Ready'},
        @{Age=1.0;Days=1L;Freshness='Fresh';Readiness='Ready'},
        @{Age=(7.0-1.0/1440);Days=6L;Freshness='Fresh';Readiness='Ready'},
        @{Age=7.0;Days=7L;Freshness='Fresh';Readiness='Ready'},
        @{Age=(7.0+1.0/1440);Days=7L;Freshness='Aging';Readiness='Ready'},
        @{Age=(30.0-1.0/1440);Days=29L;Freshness='Aging';Readiness='Ready'},
        @{Age=30.0;Days=30L;Freshness='Aging';Readiness='Ready'},
        @{Age=(30.0+1.0/1440);Days=30L;Freshness='Stale';Readiness='Refresh required'},
        @{Age=31.0;Days=31L;Freshness='Stale';Readiness='Refresh required'}
    )
    foreach($culture in @('en-US','fr-FR','tr-TR')) {
        $before=[Globalization.CultureInfo]::CurrentCulture
        try {
            [Globalization.CultureInfo]::CurrentCulture=[Globalization.CultureInfo]$culture
            $allTrust=[Collections.Generic.List[object]]::new()
            foreach($case in $ageCases) {
                $ageAudit=$audit | ConvertTo-Json -Depth 15 | ConvertFrom-Json @jsonArgs
                foreach($file in $ageAudit.Files){$file.Stats.LatestCollectionUTC=$now.AddDays(-$case.Age).ToString('O')}
                $ageTrust=@(ConvertFrom-IntuneOperationalAudit -Audit $ageAudit -AsOfUtc $now)
                Assert ($ageTrust.Count -eq 12) ('All source profiles retained at age '+$case.Age)
                foreach($row in $ageTrust) {
                    Assert ($row.'Evidence Age Days' -is [long] -and $row.'Evidence Age Days' -eq $case.Days) ('Completed-day export '+$culture+'/'+$case.Age)
                    Assert ($row.'Freshness State' -eq $case.Freshness -and $row.'Decision Readiness' -eq $case.Readiness) ('Unrounded age thresholds '+$culture+'/'+$case.Age)
                    $allTrust.Add($row)
                }
            }
            $null=Test-TrustExport $allTrust ('boundaries-'+$culture)
        } finally {[Globalization.CultureInfo]::CurrentCulture=$before}
    }
    # The unchanged validator must still reject the pre-fix fractional export.
    $legacyTrust=$audit | ConvertTo-Json -Depth 15 | ConvertFrom-Json @jsonArgs
    $legacyRows=@(ConvertFrom-IntuneOperationalAudit -Audit $legacyTrust -AsOfUtc $now)
    foreach($row in $legacyRows){$row.'Evidence Age Days'=0.08}
    $legacyValidation=Test-TrustExport $legacyRows 'legacy-fractional' $false
    Assert ($legacyValidation.Errors.Count -eq 1 -and $legacyValidation.Errors[0] -eq 'Invalid DataTrustEvidence.csv: duplicate keys=0, invalid values=12') 'Pre-fix twelve invalid ages remain rejected'
    $emptyFileAudit=$audit | ConvertTo-Json -Depth 15 | ConvertFrom-Json @jsonArgs
    foreach($group in $emptyFileAudit.Groups){$group.CompletedUTC=$now.AddHours(-23).ToString('O')}
    foreach($file in $emptyFileAudit.Files){$file.Rows=0;$file.Stats.LatestCollectionUTC=''}
    $emptyFileTrust=@(ConvertFrom-IntuneOperationalAudit -Audit $emptyFileAudit -AsOfUtc $now)
    Assert ($emptyFileTrust.Count -eq 12 -and @($emptyFileTrust | Where-Object {$_.'Evidence Age Days' -isnot [long] -or $_.'Evidence Age Days' -ne 0 -or $_.'Freshness State' -ne 'Fresh'}).Count -eq 0) 'Allowed empty files use receipt completion, with integer days'
    $null=Test-TrustExport $emptyFileTrust 'empty-files'
    # Exercise source selection/capture with isolated fixtures only. No pipeline run.
    Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
    $sourceContract=Join-Path $testRoot 'sources.json.txt'
    @{currentFiles=@();mappingFiles=@();dailyFiles=@();history=@();optionalSourceGroups='intune-operational-contract.json.txt'} | ConvertTo-Json | Set-Content -LiteralPath $sourceContract
    $plan=@(Get-PreparedSourcePlan -DataRoot $data -SourceContractPath $sourceContract -MaxSourceAgeHours 8760 -MetadataOnly)
    Assert ($plan.Count -eq 15) 'Grouped source selection includes twelve CSVs and three receipts'
    Assert (@($plan | Where-Object Kind -eq 'Receipt').Count -eq 3) 'Receipts included in stable source capture'
    Assert (@($plan | Where-Object IsOptional).Count -eq 0) 'Partial selected groups never silently skip individual files'
    $sourceSchema=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/source-schema.json') -Raw | ConvertFrom-Json -AsHashtable
    $baseContract=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/prepared-evidence-contract.json.txt') -Raw | ConvertFrom-Json -AsHashtable
    foreach($group in $spec.groups){foreach($file in $group.files){Assert (($sourceSchema.files[$file.file].columns -join '|') -eq ($file.columns -join '|')) ('Pinned raw schema '+$file.file)}}
    foreach($table in $spec.tables){$declared=@($baseContract.tables | Where-Object file -eq $table.file);Assert ($declared.Count -eq 1 -and ($declared[0].columns.name -join '|') -eq ($table.columns.name -join '|')) ('Base output contract '+$table.file)}
    $policyPath=Join-Path $testRoot 'source-policy.json.txt'
    @{MaxSourceAgeHours=168;AgeOverrides=@{}} | ConvertTo-Json | Set-Content -LiteralPath $policyPath
    $entryRoot=Join-Path $testRoot 'entry-point'
    & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'New-IntuneOperationalEvidence.ps1') -DataRoot $data -OutputRoot $entryRoot -TenantKey 'synthetic-test' -SourcePolicyPath $policyPath -AsOfUtc ($now.ToString('O')) | Out-Null
    Assert ($LASTEXITCODE -eq 0) 'Staging entry point accepts isolated synthetic source and policy'
    $entryAudit=Read-IntuneOperationalJson (Join-Path $entryRoot 'IntuneOperationalAudit.json.txt')
    Assert (@(ConvertFrom-IntuneOperationalAudit -Audit $entryAudit -AsOfUtc $now).Count -eq 12) 'Serialized staging audit preserves UTC for Data Trust'
    $null=Test-TrustExport @(ConvertFrom-IntuneOperationalAudit -Audit $entryAudit -AsOfUtc $now) 'serialized-audit'
    foreach($value in @('2026-01-02T09:00:00Z','2026-01-02T10:00:00+01:00','2026-01-02T04:00:00-05:00')){Assert ((ConvertTo-IntuneEvidenceTime $value -AsOfUtc $now).UTC -eq '2026-01-02T09:00:00.0000000Z') 'Explicit offsets equivalent'}
    foreach($culture in @('en-US','fr-FR','tr-TR')){
        $before=[Globalization.CultureInfo]::CurrentCulture
        try{[Globalization.CultureInfo]::CurrentCulture=[Globalization.CultureInfo]$culture;Assert ((ConvertTo-IntuneEvidenceTime '2026-01-02 09:00:00' -SourceFieldUtc -AsOfUtc $now).UTC -eq '2026-01-02T09:00:00.0000000Z') ('Invariant timestamp '+$culture)}finally{[Globalization.CultureInfo]::CurrentCulture=$before}
    }
    foreach($value in @('01/02/2026 09:00:00','2026-02-30T09:00:00Z','invalid')){Assert ((ConvertTo-IntuneEvidenceTime $value -AsOfUtc $now).Basis -eq 'Invalid') 'Ambiguous or invalid timestamp'}
    Assert ((ConvertTo-IntuneEvidenceTime '2026-01-03T09:00:00Z' -AsOfUtc $now).Basis -eq 'Future') 'Future source date not qualified'
    Assert ((ConvertTo-IntuneEvidenceTime '' -AsOfUtc $now).Basis -eq 'Missing') 'Missing timestamp explicit'
    $early=(ConvertTo-IntuneEvidenceTime '2025-10-26T02:30:00+02:00' -AsOfUtc $now).UTC
    $late=(ConvertTo-IntuneEvidenceTime '2025-10-26T02:30:00+01:00' -AsOfUtc $now).UTC
    Assert ($early -eq '2025-10-26T00:30:00.0000000Z' -and $late -eq '2025-10-26T01:30:00.0000000Z') 'DST ambiguity resolved only by explicit source offsets'
    Assert ((New-IntuneEvidenceKey @('a|b','c')) -ne (New-IntuneEvidenceKey @('a','b|c'))) 'Key delimiters unambiguous'
    Reject {Build 'valid'} 'No overwrite of existing staging output'
    Reject {New-IntuneOperationalEvidence -LastRoot $last -OutputRoot (Join-Path $data 'unsafe') -TenantKey 'synthetic-test' -AsOfUtc $now} 'Raw tree outputs refused'
    foreach($group in $spec.groups){
        foreach($mutation in @(
            @{Field='TenantKey';Value='another-tenant'},@{Field='Producer';Value='other.ps1'},@{Field='ScriptVersion';Value='0.1.0'},
            @{Field='Status';Value='Failed'},@{Field='Errors';Value=1},@{Field='RunId';Value=''},@{Field='FullInventoryQualified';Value='false'},
            @{Field='CompletedAtUtc';Value='2026-01-04T11:00:00Z'},@{Field='StartedAtUtc';Value='2026-01-02 10:00:00'},@{Field='ContractVersion';Value='0.0'})){
            Write-Fixtures;Edit-Receipt $group {param($r)$r[$mutation.Field]=$mutation.Value};Reject {Validate} ('Receipt gate '+$group.name+'/'+$mutation.Field)
        }
        Write-Fixtures;Edit-Receipt $group {param($r)$r.Files[0].Rows=999};Reject {Validate} ('Receipt count '+$group.name)
        Write-Fixtures;Edit-Receipt $group {param($r)$r.Files[0].SHA256='0'*64};Reject {Validate} ('Receipt hash '+$group.name)
        Write-Fixtures;Edit-Receipt $group {param($r)$r.Files[0].RunId='mixed'};Reject {Validate} ('Receipt file run '+$group.name)
        Write-Fixtures;Remove-Item -LiteralPath (Join-Path $last $group.files[0].file);Reject {Validate} ('Incomplete group '+$group.name)
        Write-Fixtures;Remove-Item -LiteralPath (Join-Path $last $group.receipt);Reject {Validate} ('Receipt missing '+$group.name)
    }
    Reset-Fixtures;Write-Fixtures
    $script:Fixtures['Intune_ConfigurationSettings.csv'][0].RunId='mixed-business';Write-Fixtures;Reject {Validate} 'Mixed business run rejected without comparing receipt ID'
    Reset-Fixtures;$script:Fixtures['Intune_ManagedApps.csv'][0].TenantKey='another-tenant';Write-Fixtures;Reject {Validate} 'Every source row tenant checked'
    Reset-Fixtures;$script:Fixtures['Intune_ManagedApps.csv'][0].CollectedAtUtc='2026-01-02 10:00:00';Write-Fixtures;Reject {Validate} 'Collection timestamps require explicit offsets'
    Reset-Fixtures;$script:Fixtures['Intune_ManagedApps.csv'][0].TenantId='other-directory';Write-Fixtures;Reject {Validate} 'Directory identity not equated solely on tenant alias'
    Reset-Fixtures;$script:Fixtures['Intune_ConfigurationSettings.csv'][0].CollectedAtUtc='2026-01-02T10:00:01Z';Write-Fixtures;Reject {Validate} 'Mixed runtime collection timestamps rejected'
    Reset-Fixtures;$script:Fixtures['Intune_ManagedApps.csv']+= $script:Fixtures['Intune_ManagedApps.csv'][0];Write-Fixtures;Reject {Validate} 'Duplicate app catalog key rejected'
    Reset-Fixtures;$script:Fixtures['Intune_ManagedApps.csv'][0].ApplicationId='';Write-Fixtures;Reject {Validate} 'Missing required technical key rejected'
    # Optional group: a complete but old collection is still the latest available data (warning only).
    Reset-Fixtures;Write-Fixtures;Assert (@(Test-IntuneOperationalSnapshot -LastRoot $last -TenantKey 'synthetic-test' -AsOfUtc $now.AddDays(8) -WarningAction SilentlyContinue -WarningVariable staleWarnings).Count -eq 3 -and @($staleWarnings | Where-Object { [string]$_ -like '*used although stale*' }).Count -gt 0) 'Old complete Intune groups kept as the latest available data'
    $longLimits=@{};foreach($group in $spec.groups){foreach($file in $group.files){$longLimits[$file.file]=744}}
    Assert (@(Test-IntuneOperationalSnapshot -LastRoot $last -TenantKey 'synthetic-test' -AsOfUtc $now.AddDays(8) -AgeOverrides $longLimits).Count -eq 3) 'Configured per-file age overrides retained'
    Reset-Fixtures;$script:Fixtures['Intune_ManagedAppInstallSummary.csv'][0].ReportedInstalledRows='2';Write-Fixtures;Reject {Build 'bad-summary'} 'Summary/deployment counts reconcile'
    Reset-Fixtures;$script:Fixtures['Intune_ConfigurationStatusCoverage.csv'][0].DeviceStatusHistoryRows='999';Write-Fixtures;Reject {Build 'bad-coverage'} 'Coverage/history counts reconcile'
    Reset-Fixtures
    $appFile=$spec.groups[1].files | Where-Object file -eq 'Intune_ManagedApps.csv'
    $summaryFile=$spec.groups[1].files | Where-Object file -eq 'Intune_ManagedAppInstallSummary.csv'
    $assignmentFile=$spec.groups[1].files | Where-Object file -eq 'Intune_ManagedAppAssignments.csv'
    foreach($id in @('app-2','app-3')){
        $script:Fixtures[$appFile.file]+=New-Row $appFile @{ApplicationId=$id;RunId='business-ManagedApps'}
        $script:Fixtures[$summaryFile.file]+=New-Row $summaryFile @{ApplicationId=$id;RunId='business-ManagedApps';StatusSource=$(if($id -eq 'app-2'){'NoReportRows'}else{'NotAssigned'});ObservedRows='0';ReportedInstalledRows='0';ReportedFailureRows='0';ReportedNotInstalledRows='0';PendingRows='0';NotApplicableRows='0';UnknownRows='0'}
    }
    $script:Fixtures[$assignmentFile.file]+=New-Row $assignmentFile @{ApplicationId='app-2';AssignmentId='assign-2';TargetType='groupAssignmentTarget';Intent='uninstall';RunId='business-ManagedApps'}
    $state=$script:Fixtures['Intune_ManagedAppDeviceStates.csv'][0].PSObject.Copy();$state.ReportedDeploymentState='brand-new-state';$state.DeviceId='missing-reference'
    $script:Fixtures['Intune_ManagedAppDeviceStates.csv']+=$state
    $script:Fixtures[$summaryFile.file][0].ObservedRows='2';$script:Fixtures[$summaryFile.file][0].UnknownRows='1'
    $script:Fixtures['Intune_ConfigurationSettings.csv'][0].PolicyId='outside-catalog'
    $script:Fixtures['Intune_ConfigurationDeviceCurrentStates.csv'][0].IntuneDeviceId='unmatched-device'
    Write-Fixtures;$null=Build 'limits'
    $apps=@(Import-Csv (Join-Path $testRoot 'limits/IntuneManagedAppEvidence.csv'))
    Assert (@($apps | Where-Object MeasurementState -eq 'Not measured').Count -eq 2) 'NoReportRows and NotAssigned not failures or healthy'
    $deploy=@(Import-Csv (Join-Path $testRoot 'limits/IntuneManagedAppDeploymentEvidence.csv'))
    Assert ($deploy.Count -eq 2 -and @($deploy.EvidenceKey | Sort-Object -Unique).Count -eq 2) 'Nullable UPN/version do not collapse report records'
    Assert ($deploy[1].ReportedDeploymentState -eq 'Unknown') 'Unknown producer enum cannot become failure or health'
    Assert ($deploy[1].InventoryMatch -eq 'Unmatched technical identifier') 'Orphan deployment retained'
    $setting=Import-Csv (Join-Path $testRoot 'limits/IntuneConfigurationSettingEvidence.csv')
    Assert ($setting.CatalogMatch -eq 'Outside collected catalog') 'Uncatalogued policy references retained'
    $assignment=Import-Csv (Join-Path $testRoot 'limits/IntuneConfigurationAssignmentEvidence.csv')
    Assert ($assignment.TargetType -like '*exclusion*') 'Exclusion assignment retained without applied inference'
    Reset-Fixtures
    $script:Fixtures['Intune_ManagedAppAssignments.csv'][0].TargetType='exclusionGroupAssignmentTarget'
    $script:Fixtures['Intune_ManagedAppDeviceStates.csv']=@()
    $script:Fixtures['Intune_ManagedAppInstallSummary.csv'][0].StatusSource='NoReportRows'
    $script:Fixtures['Intune_ManagedAppInstallSummary.csv'][0].ObservedRows='0';$script:Fixtures['Intune_ManagedAppInstallSummary.csv'][0].ReportedInstalledRows='0'
    Write-Fixtures;$null=Build 'exclusion-only'
    $appSummary=Import-Csv (Join-Path $testRoot 'exclusion-only/IntuneManagedAppEvidence.csv')
    Assert ($appSummary.MeasurementState -eq 'Not measured') 'Exclusion-only assignment scope matches producer, not install success'
    Reset-Fixtures;Write-Fixtures
    @([pscustomobject]@{TenantKey='synthetic-test';'Device ID'='device-1'},[pscustomobject]@{TenantKey='synthetic-test';'Device ID'='device-1'}) | Export-Csv -LiteralPath (Join-Path $last 'Intune_Devices_Inventory.csv') -NoTypeInformation
    $null=Build 'ambiguous-reference'
    $config=Import-Csv (Join-Path $testRoot 'ambiguous-reference/IntuneConfigurationDeviceEvidence.csv')
    Assert ($config.InventoryMatch -eq 'Ambiguous reference identifier') 'Ambiguous inventory references not merged or expanded'
    Reset-Fixtures
    $script:Fixtures['Intune_EnrollmentFailures.csv'][0].ManagedDeviceIdentifier='device-1';$script:Fixtures['Intune_EnrollmentFailures.csv'][0].HasManagedDeviceIdentifier='True'
    Write-Fixtures;$null=Build 'linked-event'
    $event=Import-Csv (Join-Path $testRoot 'linked-event/IntuneEnrollmentFailureEvidence.csv')
    Assert ($event.InventoryMatch -eq 'Matched technical identifier') 'Enrollment supplied managed ID exact join'
    Reset-Fixtures
    $script:Fixtures['Intune_ManagedAppDeviceStates.csv']=@(1..5000 | ForEach-Object {$script:Fixtures['Intune_ManagedAppDeviceStates.csv'][0].PSObject.Copy()})
    $script:Fixtures['Intune_ManagedAppInstallSummary.csv'][0].ObservedRows='5000';$script:Fixtures['Intune_ManagedAppInstallSummary.csv'][0].ReportedInstalledRows='5000'
    Write-Fixtures;$null=Build 'streamed'
    $streamed=@(Import-Csv (Join-Path $testRoot 'streamed/IntuneManagedAppDeploymentEvidence.csv'))
    Assert ($streamed.Count -eq 5000 -and @($streamed.EvidenceKey | Sort-Object -Unique).Count -eq 5000) 'Streaming projection preserves all logical records, including repeated observations'
    Assert ($streamed[-1].SourceRecord -eq '5000') 'Technical observation keys use logical CSV ordinal'
    Reset-Fixtures;Write-Fixtures
    # Remove only explicitly named synthetic source files; optional groups are absent, not zero healthy devices.
    foreach($group in $spec.groups){foreach($file in $group.files){Remove-Item -LiteralPath (Join-Path $last $file.file)};Remove-Item -LiteralPath (Join-Path $last $group.receipt)}
    $empty=Build 'absent'
    Assert (@($empty.Groups | Where-Object Availability -eq 'Not supplied').Count -eq 3) 'Absent optional groups remain Not supplied'
    $emptyAudit=$empty | ConvertTo-Json -Depth 15 | ConvertFrom-Json @jsonArgs
    $trust=@(ConvertFrom-IntuneOperationalAudit -Audit $emptyAudit -AsOfUtc $now)
    Assert ($trust.Count -eq 3 -and @($trust | Where-Object 'Decision Readiness' -ne 'Optional unavailable').Count -eq 0) 'Missing groups not blocking unrelated sources'
    Assert (@($trust | Where-Object {$null -ne $_.'Evidence Age Days'}).Count -eq 0) 'Unavailable source ages remain null, not zero'
    $null=Test-TrustExport $trust 'absent-groups'
    foreach($table in $spec.tables | Where-Object {$_.ContainsKey('source')}){Assert (@(Import-Csv (Join-Path $testRoot ('absent/'+$table.file))).Count -eq 0) 'Absent group tables header-only'}
    $contractPath=Join-Path $testRoot 'output-contract.json.txt';@{tables=$spec.tables} | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $contractPath
    foreach($name in @('valid','limits','absent')){
        & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root (Join-Path $testRoot $name) -ContractPath $contractPath | Out-Null
        Assert ($LASTEXITCODE -eq 0) ('Prepared output contract validated '+$name)
    }
    [pscustomobject]@{Passed=$true;Assertions=$script:Assertions;Data='Synthetic only';Collectors=$false;OperationalPreparation=$false;Transfer=$false;PowerBIRefresh=$false} | ConvertTo-Json
}finally{
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/')
    if($resolved.StartsWith($temp+'\',[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved) -match '^intune-operational-offline-[a-f0-9]{32}$') {Remove-Item -LiteralPath $resolved -Recurse -Force}
}
