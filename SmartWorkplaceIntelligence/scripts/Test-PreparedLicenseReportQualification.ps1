# Version 1.2.1
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
    $snapshot.GeneratedAtUtc='2026-10-17T23:00:00Z'
    $snapshot.Sources=@(foreach($group in @(
        @{Producer='SmartM365-M365UserActivity-Inventory.ps1';Files=@('M365_Users_Activity.csv','M365_Mailbox_Usage.csv','M365_Email_Activity.csv','M365_Apps_Usage_180D.csv','M365_OneDrive_Usage.csv')},
        @{Producer='SmartM365-Devices-Inventory.ps1';Files=@('Intune_Devices_Inventory.csv')}
    )) {
        $records=@(foreach($name in $group.Files) {
            $raw=Join-Path $last $name
            [pscustomobject]@{UserId='contoso-one';Value='synthetic'} | Export-Csv -LiteralPath $raw -NoTypeInformation
            $hash=(Get-FileHash -LiteralPath $raw).Hash
            @{File=$name;RunId='source-run';Status='Success';Errors=0;IsPartialInventory=$false;SHA256=$hash}
        })
        $stem='SmartInventory_'+[IO.Path]::GetFileNameWithoutExtension($group.Producer)
        Json (Join-Path $last ($stem+'.current.json.txt')) @{Owner='SmartInventory-SourceReceipt';Producer=$group.Producer;TenantKey=$tenant;RunId='source-run';Status='Completed';Errors=0;IsPartialInventory=$false;ConsumerScopeQualified=$true;PublicationProtocol=1;CompletedAtUtc='2026-10-17T22:00:00Z';Files=$records}
        Json (Join-Path $last ($stem+'.run.json.txt')) @{Owner='SmartInventory-SourceRun';Producer=$group.Producer;TenantKey=$tenant;RunId='source-run';Status='Completed';UnqualifiedPublication=$false}
        foreach($record in $records) { @{Name=$record.File;Ready=$true;Provisional=$false;SHA256=$record.SHA256} }
    })
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
function Get-Qualification([string]$Capture,[switch]$RequireCurrentSources){
    $batch=Join-Path $root ('batch-'+[guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $batch|Out-Null
    $entries=@(Copy-PreparedArtifacts -SnapshotRoot $Capture -Destination $batch -TenantKey $tenant)
    Assert-PreparedArtifacts -BatchRoot $batch -Entries $entries -TenantKey $tenant
    [pscustomobject]@{Entries=$entries;Result=(Get-PreparedLicenseReportQualification -SnapshotRoot $Capture -TenantKey $tenant -Artifacts $entries -RequireCurrentSources:$RequireCurrentSources)}
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
    # Frozen-source matching: a valid report must not be reused after its dependencies change.
    $q=Get-Qualification (New-Capture 'strict-valid' -WithDecisions) -RequireCurrentSources
    Check ($q.Result.State -eq 'Qualified') ('Current qualified dependency snapshot refused: '+$q.Result.Reason)
    # Licences writes GeneratedAtUtc with '+00:00'; any explicit offset is the same instant.
    foreach($offsetCase in @(@{Name='strict-plus00';Value='2026-10-17T23:00:00+00:00'},@{Name='strict-plus02';Value='2026-10-18T01:00:00+02:00'})) {
        $current=New-Capture $offsetCase.Name -WithDecisions
        $snapshotPath=Join-Path (Join-Path $current 'DATA-LAST') $license.snapshot
        $snapshot=Get-Content -LiteralPath $snapshotPath -Raw | ConvertFrom-Json
        $snapshot.GeneratedAtUtc=$offsetCase.Value
        Json $snapshotPath $snapshot
        $q=Get-Qualification $current -RequireCurrentSources
        Check ($q.Result.State -eq 'Qualified') ('Explicit offset refused: '+$offsetCase.Name+' '+$q.Result.Reason)
    }
    $strictCases=@(
        @{Name='source-unqualified';Change={$snapshot.Sources[0].Ready=$false}},
        @{Name='source-string-ready';Change={$snapshot.Sources[0].Ready='true'}},
        @{Name='source-provisional';Change={$snapshot.Sources[0].Provisional=$true}},
        @{Name='source-missing';Change={$snapshot.Sources=@($snapshot.Sources | Select-Object -Skip 1)}},
        @{Name='source-duplicate';Change={$snapshot.Sources+=@($snapshot.Sources[0])}},
        @{Name='source-hash';Change={$snapshot.Sources[0].SHA256='0'*64}},
        @{Name='source-legacy';Change={$snapshot.PSObject.Properties.Remove('Sources')}},
        @{Name='source-older';Change={$snapshot.GeneratedAtUtc='2026-10-17T21:00:00Z'}},
        @{Name='source-offsetless';Change={$snapshot.GeneratedAtUtc='2026-10-17T23:00:00'}},
        @{Name='receipt-missing';Change={Remove-Item -LiteralPath $proofPath}},
        @{Name='receipt-foreign';Change={$proof.TenantKey='foreign';Json $proofPath $proof}},
        @{Name='receipt-failed';Change={$proof.Status='Failed';Json $proofPath $proof}},
        @{Name='receipt-partial';Change={$proof.IsPartialInventory=$true;Json $proofPath $proof}},
        @{Name='receipt-unqualified';Change={$proof.ConsumerScopeQualified=$false;Json $proofPath $proof}},
        @{Name='receipt-hash';Change={$proof.Files[0].SHA256='0'*64;Json $proofPath $proof}},
        @{Name='receipt-other-run';Change={$proof.Files[0].RunId='other-run';Json $proofPath $proof}},
        @{Name='run-active';Change={$run.Status='Running';Json $runPath $run}},
        @{Name='run-unqualified';Change={$run.UnqualifiedPublication=$true;Json $runPath $run}},
        @{Name='run-other';Change={$run.RunId='other-run';Json $runPath $run}},
        @{Name='run-missing';Change={Remove-Item -LiteralPath $runPath}},
        @{Name='raw-changed';Change={Add-Content -LiteralPath (Join-Path $last 'M365_Users_Activity.csv') -Value 'changed'}}
    )
    foreach($case in $strictCases) {
        $current=New-Capture $case.Name -WithDecisions
        $last=Join-Path $current 'DATA-LAST'
        $snapshotPath=Join-Path $last $license.snapshot
        $snapshot=Get-Content -LiteralPath $snapshotPath -Raw | ConvertFrom-Json
        $proofPath=Join-Path $last 'SmartInventory_SmartM365-M365UserActivity-Inventory.current.json.txt'
        $runPath=Join-Path $last 'SmartInventory_SmartM365-M365UserActivity-Inventory.run.json.txt'
        $proof=Get-Content -LiteralPath $proofPath -Raw | ConvertFrom-Json
        $run=Get-Content -LiteralPath $runPath -Raw | ConvertFrom-Json
        & $case.Change
        Json $snapshotPath $snapshot
        $q=Get-Qualification $current -RequireCurrentSources
        Check ($q.Result.State -eq 'Not qualified' -and $q.Result.Reason -like '*dependency mismatch*') ('Stale source accepted: '+$case.Name+' '+$q.Result.Reason)
        Check (@($q.Result.Files | Where-Object State -eq 'Qualified').Count -eq 0) ('Unqualified snapshot exposed report or decision values: '+$case.Name)
        Check ((Get-WorkforceDisabledLicenseEvidence -DataRoot $current -TenantKey $tenant).State -eq 'Not measured') ('Workforce reused unqualified snapshot: '+$case.Name)
    }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+AoC3mTnE/UrZ
# 7bDC4m2PazpBJbSO9ARTgDRVr/Hbn6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPym6J6tzidOHnAnt9ylLN6nIKC61vzPeAKb/NtYnjcLMA0GCSqG
# SIb3DQEBAQUABIIBgFVT8sE7pmLQw8tJKn7p0Fa8t+HfY+0SyvcHyw0U9pPhSfz5
# 9/8q1L1/hqJpUv+V0d9CchyEzh+zxsjFggVUUO7+TQWRv/HyEcOdtGPoAMSQS7je
# m3wWHTg4UUtRbau6eopvddWg9dwqcj8VymK9eE9IDR/2t++1ZLAb6ePXeCGwewSN
# bY7ZU3SKcHoW+3UZVsFMINI4GUbqAQumOeB8UQWzd69XX8SNjN9l+SptI61/KVCv
# t2r+blcIC1olod3jhpOrbkfmCeTfwb7XfYVckI5QiBgkY34Xmz1+J0GtmwlfTEg+
# OeIBAfVXnVjISEFXQQqsyPf5KGGum33RjOtf/fYxvhP0X0mRIlJ/4YrHuCGXnHS+
# l4ctWjdPsOLk4vbh8W1aFXNeiNlkinoxYPPO6EvQOgExlkt2sbNT9UXjUIH8Xu3a
# CH2DWMPrSdLF2g8Hbdlfpc2P2RHP4IKcz9GGSwpWIEPq4MofMRp9LbK8DinUJZCM
# 2pIuINRJIsvI2V97AqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcyMTEx
# MDVaMC8GCSqGSIb3DQEJBDEiBCCxDU85Wv0qdAJCERXqCB3+zNSExZDXgYLA6JG/
# jPbxhTANBgkqhkiG9w0BAQEFAASCAgBb9IgaTIyPL9p5YpZCvsXEjUxhJeRrtyxr
# fO1WUxzIbZ5saw/psHAa3jH09Ru6EWPM0QFGxiq5zWJ/1kxAd6PjPm3syOpL3KwV
# GL5etUhv0JNbfa0Ye7QnjHvjQxz+htqjplbRTOXYLhMRspidnjcfKlgD19EBpANZ
# DkiRunDGLe3kdt2kvah76anx0XpWXu/DzcFqzTNLpKww1nzZjkPV1KRs8LIPg/WF
# V1PQZnmM/lE50P1PPR1KcQAtnWXKGXEPb4rogotXtfPCmnWlS35yq0W3kWEdFS1U
# 3JR0j0rlBfjlqcYpnExrTkvGZ6fqRZxLxUDKrHT8MapifYKA81l3kLQrt2bISdeP
# yXzhitihQe6zPpvM1q1zi5tamsiQIC3FJ7gr5/VHe8FqWpbphTz2Z/tlvmILpE25
# OL22uiyzmB7vWXSSY+YZTgylr+juOReH7L3B8M/dPjMB3K91gV+q+e3/rtN21Hc4
# yDqBYDjg1VN297LMrdVi7N7kx+MxJBSPI5/NsV24BAb3z4iUaO7v4WI3y48bIq9e
# 0AN4dU7TwQWCvvbIwOhbvfYfN+/P/5bwddNkhC92jCJjaAiiQCtmq6uQ9EdDBVB/
# NB3cHcx0Dbqq+kHnjFLWdYwKxqVCUcpR7jehnDnfRMBm1+aZfkitOUOcmaETS29d
# TQz0Amrbyg==
# SIG # End signature block
