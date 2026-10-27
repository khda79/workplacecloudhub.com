# Version 1.1.0
#Requires -Version 7.0
[CmdletBinding()]
param()
# Synthetic license report qualification checks: the Licences report is consumed only when the current
# licensing receipt records it. No collector, tenant, cloud or Power BI access.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedArtifacts.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('LicenseQualificationTest-'+[guid]::NewGuid().ToString('N'))
$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
function Json([string]$Path,$Value){[IO.File]::WriteAllText($Path,($Value|ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))}
$spec=(Get-Content (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw|ConvertFrom-Json)
$license=$spec.licenseReport
$tenant='synthetic'
$snapshotId=('b'*32)
$runId='11111111-2222-3333-4444-555555555555'

function New-Capture {
    param([string]$Name,[switch]$WithDecisions,[switch]$NoReceipt,[string]$ReceiptStatus='Completed',[string]$SnapshotRunId=$runId,
        [switch]$NoReceiptRun,[string]$UnpublishedFile,[string]$TamperedFile,[string]$ForeignSnapshotFile,[string]$ForeignRunFile,[string]$RowTenant=$tenant)
    $capture=Join-Path $root $Name
    $last=Join-Path $capture 'DATA-LAST'
    New-Item -ItemType Directory -Path $last -Force|Out-Null
    foreach($name in $spec.currentFiles){
        $path=Join-Path $last $name
        if($name -ceq $license.snapshot){continue}
        if($name -like '*.csv'){
            $sid=if($name -ceq $ForeignSnapshotFile){'c'*32}else{$snapshotId}
            $rowTenant=if($name -like 'M365_Licenses_Report*'){$RowTenant}else{$tenant}
            [pscustomobject]@{TenantKey=$rowTenant;SnapshotId=$sid}|Export-Csv -LiteralPath $path -NoTypeInformation
        }
    }
    $snapshot=[ordered]@{TenantKey=$tenant;SnapshotId=$snapshotId;SchemaVersion=2;Products=@(1,2,3,4);OtherProducts=@(1,2,3);AssignmentIssues=@{Available=$false;Rows=@()}}
    if(-not $NoReceiptRun){$snapshot.ReceiptRun=@{Producer=$license.producer;RunId=$SnapshotRunId;StartedAtUtc='2026-10-17T22:00:00Z'}}
    Json (Join-Path $last $license.snapshot) $snapshot
    $published=@($license.reportFiles)
    if($WithDecisions){
        foreach($name in $license.decisionFiles){
            $rid=if($name -ceq $ForeignRunFile){'other-run'}else{$runId}
            $sid=if($name -ceq $ForeignSnapshotFile){'c'*32}else{$snapshotId}
            [pscustomobject]@{SnapshotId=$sid;TenantKey=$RowTenant;ReceiptRunId=$rid;SkuPartNumber='SPE_E3'}|Export-Csv -LiteralPath (Join-Path $last $name) -NoTypeInformation
        }
        $published+=@($license.decisionFiles)
    }
    if(-not $NoReceipt){
        $files=foreach($name in $published){
            $hash=(Get-FileHash -LiteralPath (Join-Path $last $name)).Hash
            if($name -ceq $TamperedFile){$hash='0'*64}
            [ordered]@{File=$name;Required=$false;Status='Success';SHA256=$hash;RunId=$runId}
        }
        $optional=foreach($name in @($license.reportFiles)+@($license.decisionFiles)){
            [ordered]@{File=$name;Published=($published -contains $name -and $name -cne $UnpublishedFile)}
        }
        Json (Join-Path $last $license.receipt) ([ordered]@{Owner='SmartInventory-SourceReceipt';ContractVersion='1.2';Producer=$license.producer;TenantKey=$tenant;Status=$ReceiptStatus;RunId=$runId;Files=@($files);OptionalFiles=@($optional)})
    }
    foreach($history in $spec.history){
        $dir=Join-Path $capture "$($history.root)/2026-W40";New-Item -ItemType Directory -Path $dir -Force|Out-Null
        @{TenantKey=$tenant;Id='one'}|Export-Csv -LiteralPath (Join-Path $dir $history.file) -NoTypeInformation
        Json (Join-Path $dir 'manifest.json.txt') @{Week='2026-W40';SnapshotCreatedAtUtc='2026-10-01T00:00:00Z'}
    }
    return $capture
}
function Get-Qualification([string]$Capture){
    $batch=Join-Path $root ('batch-'+[guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $batch|Out-Null
    $entries=@(Copy-PreparedArtifacts -SnapshotRoot $Capture -Destination $batch -TenantKey $tenant)
    Assert-PreparedArtifacts -BatchRoot $batch -Entries $entries -TenantKey $tenant
    [pscustomobject]@{Entries=$entries;Result=(Get-PreparedLicenseReportQualification -SnapshotRoot $Capture -TenantKey $tenant -Artifacts $entries)}
}
function FileState($Result,[string]$Name){@($Result.Files | Where-Object File -ceq $Name)[0].State}

try {
    New-Item -ItemType Directory -Path $root|Out-Null
    Import-Module (Join-Path $PSScriptRoot 'WorkforceLicenseEvidence.psm1') -Force
    $capture=New-Capture 'workforce-decisions' -WithDecisions
    $decisionPath=Join-Path $capture 'DATA-LAST/M365_Licenses_ReportSkuCandidates.csv'
    function Set-WorkforceDecisions($Rows,[switch]$Empty) {
        if ($Empty) {
            '"TenantKey","SnapshotId","ReceiptRunId","UserId","CandidateType","Recommendation"' | Set-Content -LiteralPath $decisionPath
        } else { $Rows | Export-Csv -LiteralPath $decisionPath -NoTypeInformation }
        $receiptPath=Join-Path $capture ('DATA-LAST/'+$license.receipt)
        $receipt=Get-Content $receiptPath -Raw | ConvertFrom-Json
        ($receipt.Files | Where-Object File -eq 'M365_Licenses_ReportSkuCandidates.csv').SHA256=(Get-FileHash $decisionPath).Hash
        Json $receiptPath $receipt
    }
    $row=[pscustomobject]@{TenantKey=$tenant;SnapshotId=$snapshotId;ReceiptRunId=$runId;UserId='contoso|one';CandidateType='Disabled account';Recommendation='Recommended'}
    Set-WorkforceDecisions @($row,$row)
    $signal=Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant
    Check ($signal.State -eq 'Observed' -and $signal.Count -eq 1) 'Workforce must deduplicate accounts across paid SKUs.'
    $row.Recommendation='Review';Set-WorkforceDecisions @($row)
    Check ((Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant).Count -eq 1) 'Review candidates disappeared.'
    $row.Recommendation='Excluded';Set-WorkforceDecisions @($row)
    Check ((Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant).Count -eq 0) 'Excluded candidate counted.'
    Set-WorkforceDecisions @() -Empty
    $signal=Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant
    Check ($signal.State -eq 'Observed' -and $signal.Count -eq 0) 'Qualified empty decisions must be measured zero.'
    $row.TenantKey='foreign';Set-WorkforceDecisions @($row)
    $signal=Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant
    Check ($signal.State -eq 'Not measured' -and $null -eq $signal.Count) 'Foreign tenant must not become zero.'
    $row.TenantKey=$tenant;$row.Recommendation='Recommended';$row.UserId='';Set-WorkforceDecisions @($row)
    Check ($null -eq (Get-WorkforceDisabledLicenseEvidence -DataRoot $capture -TenantKey $tenant).Count) 'Missing candidate key accepted.'
    $captureMissing=New-Capture 'workforce-no-receipt' -WithDecisions -NoReceipt
    Check ($null -eq (Get-WorkforceDisabledLicenseEvidence -DataRoot $captureMissing -TenantKey $tenant).Count) 'Missing receipt became measured zero.'
    # Qualified report with both decision files.
    $q=Get-Qualification (New-Capture 'qualified' -WithDecisions)
    Check ($q.Result.State -eq 'Qualified' -and $q.Result.ReceiptRunId -eq $runId -and $q.Result.SnapshotId -eq $snapshotId) 'Recorded report must be qualified.'
    Check (@($q.Result.Files | Where-Object State -ne 'Qualified').Count -eq 0 -and @($q.Result.Files).Count -eq 7) 'All seven license files must be qualified.'
    Check (@($q.Entries | Where-Object File -in @($license.decisionFiles | ForEach-Object {"Source-$_"})).Count -eq 2) 'Decision files must be copied as optional artifacts.'
    Check (Test-PreparedArtifactName 'Source-M365_Licenses_ReportSkuCapacity.csv') 'Optional decision artifact name refused.'
    # Licences 1.48 output: decision files absent, report still qualified.
    $q=Get-Qualification (New-Capture 'nodecisions')
    Check ($q.Result.State -eq 'Qualified' -and (FileState $q.Result 'M365_Licenses_ReportSkuCandidates.csv') -eq 'Absent') 'Absent decision files must not disqualify the report.'
    # Lineage defects: published batch, Not qualified, decision files never shown alone.
    $cases=@(
        @{Name='noreceipt';Args=@{NoReceipt=$true};Reason='*not part of the capture*'}
        @{Name='failedreceipt';Args=@{ReceiptStatus='Failed'};Reason='*Failed, not Completed*'}
        @{Name='otherrun';Args=@{SnapshotRunId='99999999-0000-0000-0000-000000000000'};Reason='*another licensing receipt run*'}
        @{Name='noreceiptrun';Args=@{NoReceiptRun=$true};Reason='*names no licensing receipt run*'}
        @{Name='unpublished';Args=@{UnpublishedFile='M365_Licenses_ReportGaps.csv'};Reason='*M365_Licenses_ReportGaps.csv*not recorded as published*'}
        @{Name='tampered';Args=@{TamperedFile='M365_Licenses_ReportSummary.csv'};Reason='*M365_Licenses_ReportSummary.csv*hash differs*'}
        @{Name='foreignsnapshot';Args=@{ForeignSnapshotFile='M365_Licenses_ReportCandidates.csv'};Reason='*M365_Licenses_ReportCandidates.csv*another report snapshot*'}
    )
    foreach($case in $cases){
        $arguments=$case.Args
        $q=Get-Qualification (New-Capture $case.Name -WithDecisions @arguments)
        Check ($q.Result.State -eq 'Not qualified' -and $q.Result.Reason -like $case.Reason) "Lineage defect must be Not qualified: $($case.Name) ($($q.Result.Reason))"
        Check (@($q.Result.Files | Where-Object {$_.Kind -eq 'Decision' -and $_.State -eq 'Qualified'}).Count -eq 0) "Decision files shown without a qualified report: $($case.Name)"
    }
    # A decision file of another run is Not qualified on its own; the report stays qualified.
    $q=Get-Qualification (New-Capture 'decisionrun' -WithDecisions -ForeignRunFile 'M365_Licenses_ReportSkuCapacity.csv')
    Check ($q.Result.State -eq 'Qualified' -and (FileState $q.Result 'M365_Licenses_ReportSkuCapacity.csv') -eq 'Not qualified' -and (FileState $q.Result 'M365_Licenses_ReportSkuCandidates.csv') -eq 'Qualified') 'Decision file of another run must be disqualified alone.'
    # Another tenant stays fatal.
    $foreign=New-Capture 'foreigntenant' -WithDecisions -RowTenant 'foreign'
    $failed=$false
    try{Get-Qualification $foreign|Out-Null}catch{$failed=$_.Exception.Message -like '*tenant mismatch*'}
    Check $failed 'Another tenant in license report rows must stay fatal.'
    # Artifact contract compatibility: current and reviewed previous hashes only.
    $contractPath=Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt'
    Check (Test-PreparedArtifactContractCompatible -ContractSHA256 (Get-FileHash -LiteralPath $contractPath).Hash) 'Current artifact contract refused.'
    foreach($hash in @($spec.compatibleContractSHA256)){Check (Test-PreparedArtifactContractCompatible -ContractSHA256 $hash) 'Reviewed previous artifact contract refused.'}
    Check (-not (Test-PreparedArtifactContractCompatible -ContractSHA256 ('F'*64))) 'Unknown artifact contract accepted.'
    'Prepared license report qualification offline checks passed: '+$checks
} finally {
    $full=[IO.Path]::GetFullPath($root);$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if($full.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $full -Leaf) -match '^LicenseQualificationTest-[a-f0-9]{32}$'){
        if(Test-Path -LiteralPath $full){Remove-Item -LiteralPath $full -Recurse -Force}
    }
}
