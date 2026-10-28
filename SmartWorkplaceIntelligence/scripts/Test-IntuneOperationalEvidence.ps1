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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBiz0njBVhE11V1
# 4hDrkMDULMoNnkr1sntERvmzHeQ13qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHYn21J85Ge+YlGCFOifs2pha5N1Q2gmS/OovWSOdJxHMA0GCSqG
# SIb3DQEBAQUABIIBgKBkBhg5P2WbBpR08yxYS6nDvPAjrt9x+OdjGxms5mXTzNPV
# ulZuUpnIb7bCIu1eoMkr1dnDoPY5DwuoAUQF3XPDjmW0bA/mheH/Zj8t+IAmS0VC
# YRWP+hskGBWBhPf0mhqNWYWjR6MpDkvG9nLi0Vlt7engywHfJbZgLumqbORYwZaD
# lpQbabrcVleXJUNyVtbmA81zvQxm/58jYHFyrgrvGJ1C+xfJbaMBEQ3WHGREvuXC
# MC0HZK2Ep2XP1bWv+9kkw/PP49YAa/4sDwokSykON5h1LcdFP/06GvNbq/SJ/bf3
# on4vWXdQkjwhVqF1Bxo6LhGUfLgt8otKwgPBZABkAyFg18ss2mop/sJSCcqoxwhL
# d+w0EmohkzxnBYGpxn2AN9CgGGL3s2Zi+yIqQmxBzlSc5HoLmhsqvEdo+UZ9HHHa
# HYrwaUeyF1961xj3I82iDoehk87JaIRamQNgaqVBWyrX7QOXZeIOVBwmcVSHZSi2
# jGde2DSGizpqK0w+haGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxODIw
# MjJaMC8GCSqGSIb3DQEJBDEiBCCKsRS3HdWPo5VtkOnKJcjSOjtq/UftG6mKe3sY
# CAU32zANBgkqhkiG9w0BAQEFAASCAgCasEk+6lOveN+QSv+uB6pKXDU1XVg0lsjJ
# qPzfBbE6pZARRv4cyNkmIMBJr5zKuI6MtsffhOlYgG6T6ybg02zO4OsByOobRBnP
# 6fRWHDYVPoAToh1jjCIzeV/ALKH5mA7mNa4LZLBDsMT+0yzRNWBW1w6uKfpsrQGh
# BmLInw0BQMVnbp3Rmwn7IXf2E+DNi9u69yDmI71b17ItcuGU8G5daZlSaPaYq7w0
# uMB9RapuH5XdwO32sBxhDnNZ3XiywfQ1y4MotgBUcHHpCAgS8XBC7Z6jmc1fl3rx
# 1aA6bqBBWF57g8wzN1t20CMqM4sH4uGZFE4xWt9V0eh8fjqx8mWcMlALXAzZNt5o
# nJISxu8EOwF0xl7Vbyy13qoP8D9t0D9xyotoM+oDAf3P3t+jWYFdtT1alZCXi/Pv
# Q9gh41JaejZRUHEBdoCMm154uZ9oOe1FaeoT4FergTBsrWB6eb9qpaYCm7MYZAbt
# RRlfrxEmaYEOKjcGtNhOaHgifp8LdUY0yWl2+MgX95b0REjvCGFtmTv64+gsbB/k
# 6G49WyYn2CxdXHCuTvX8i85CPPb1+07f5oB6dtdesVjXoRA8DJ8q8UbA1wM39w92
# Rbr8SIipcFeXXKXJ8ba4QSAEc3uwbyZe26RG96vdH3IX0nsvfdC9+okM2IE1ms4N
# SgeyIwrdKg==
# SIG # End signature block
