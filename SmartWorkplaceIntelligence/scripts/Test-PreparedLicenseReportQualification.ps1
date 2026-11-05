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
    param([string]$Name,[switch]$WithDecisions,[switch]$WithIntegrity,[switch]$NoReceipt,[string]$ReceiptStatus='Completed',[string]$SnapshotRunId=$runId,
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
    $integrityName='M365_Licenses_ReportSnapshotIntegrity.csv'
    if($WithIntegrity){
        [pscustomobject]@{SnapshotId=$snapshotId;TenantKey=$tenant;ReceiptRunId=$runId;SnapshotSHA256=(Get-FileHash -LiteralPath (Join-Path $last $license.snapshot)).Hash} |
            Export-Csv -LiteralPath (Join-Path $last $integrityName) -NoTypeInformation
        $published+=@($integrityName)
    }
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
        $optional=foreach($name in @($license.reportFiles)+@($license.decisionFiles)+@($integrityName)){
            [ordered]@{File=$name;Published=($published -contains $name -and $name -cne $UnpublishedFile)}
        }
        Json (Join-Path $last $license.receipt) ([ordered]@{Owner='SmartInventory-SourceReceipt';ContractVersion='1.2';Producer=$license.producer;TenantKey=$tenant;Status=$ReceiptStatus;RunId=$runId;Errors=0;IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@($files);OptionalFiles=@($optional)})
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
    Check (@($q.Result.Files | Where-Object { $_.Kind -ne 'Integrity' -and $_.State -ne 'Qualified' }).Count -eq 0 -and (FileState $q.Result 'M365_Licenses_ReportSnapshotIntegrity.csv') -eq 'Absent') 'Legacy report files must be qualified without assuming snapshot integrity.'
    Check (@($q.Entries | Where-Object File -in @($license.decisionFiles | ForEach-Object {"Source-$_"})).Count -eq 2) 'Decision files must be copied as optional artifacts.'
    Check (Test-PreparedArtifactName 'Source-M365_Licenses_ReportSkuCapacity.csv') 'Optional decision artifact name refused.'
    $intCapture=New-Capture 'snapshot-integrity' -WithDecisions -WithIntegrity
    $q=Get-Qualification $intCapture
    Check ((FileState $q.Result 'M365_Licenses_ReportSnapshotIntegrity.csv') -eq 'Qualified') 'Snapshot integrity evidence should be qualified.'
    $snapshotPath=Join-Path (Join-Path $intCapture 'DATA-LAST') $license.snapshot
    Add-Content -LiteralPath $snapshotPath -Value ' '
    $q=Get-Qualification $intCapture
    Check ($q.Result.State -eq 'Qualified' -and (FileState $q.Result 'M365_Licenses_ReportSnapshotIntegrity.csv') -eq 'Not qualified') 'Changed snapshot must invalidate only the new integrity evidence.'
    $malformed=New-Capture 'snapshot-integrity-schema' -WithIntegrity
    $last=Join-Path $malformed 'DATA-LAST'
    $integrityPath=Join-Path $last 'M365_Licenses_ReportSnapshotIntegrity.csv'
    [pscustomobject]@{SnapshotId=$snapshotId;TenantKey=$tenant;ReceiptRunId=$runId;Other='bad'} | Export-Csv -LiteralPath $integrityPath -NoTypeInformation
    $receiptPath=Join-Path $last $license.receipt
    $receipt=Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
    ($receipt.Files | Where-Object File -ceq 'M365_Licenses_ReportSnapshotIntegrity.csv').SHA256=(Get-FileHash -LiteralPath $integrityPath).Hash
    Json $receiptPath $receipt
    $q=Get-Qualification $malformed
    Check ($q.Result.State -eq 'Qualified' -and (FileState $q.Result 'M365_Licenses_ReportSnapshotIntegrity.csv') -eq 'Not qualified') 'Malformed integrity evidence must be unavailable without breaking the existing report.'
    $partial=New-Capture 'snapshot-integrity-partial' -WithIntegrity
    $receiptPath=Join-Path (Join-Path $partial 'DATA-LAST') $license.receipt
    $receipt=Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
    $receipt.IsPartialInventory=$true
    Json $receiptPath $receipt
    $q=Get-Qualification $partial
    Check ($q.Result.State -eq 'Qualified' -and (FileState $q.Result 'M365_Licenses_ReportSnapshotIntegrity.csv') -eq 'Not qualified') 'Partial licensing receipt must not qualify snapshot integrity.'
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
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBm3cCjb8n7rqlq
# LqHWPGokyS1yc4tsFbYMmJM/umuW5qCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgx3CmGEeWqKOV
# tjDaRvcLbjoHZRlOGc0yYhPXqXByeJUwDQYJKoZIhvcNAQEBBQAEggGAKVRNu137
# e3nYht7geoymDrnTTGU+HfRxtZaPmO60DoNfyELyn6xFxJtzi1u57bWN21w2x7Ix
# 8fUHD2i6yE8JOt7zIZ7yuMs2bo3FM1XHP6wnNrtZACio0U/H/3QVxqjZ+heVgA6t
# dScYGpFiWc9La0eaYf31qQ8B12jauVi4jD0Znq7a7UGDNQBmutkosvgPwE/QJu4W
# nrNqySTgyvghxVuVwXTyYTZMmamtzYGfnhghRsqfHi33joll7LSqnNU6jeiKzG+b
# ZkdO6H4GnGQCdPmsSdP8ZtrCYq1CDuGYqtQzPQQgefdFWzST0tp7rjIC/ONUelDH
# lBMqL+svBSLR5lM2nK5xYnE/0AWNi+U9OvNhMsrkeFN48Hd8t/yFceyP1Y/pJBfs
# Vc2O7VlZCH1tu74G7UQKoHMZxW/wcKHOUW0FXCtKL6Ppnoct2663Xg1AKELOJUwg
# 0oG+9IDovgTUfsrQWa5AdFNMI9ve62VCLh4TMb2gGXw4z4mpAId+rKaxoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTA1MTU1MjQyWjA/BgkqhkiG9w0BCQQxMgQw6EWlt5Xaw2ViOFyTdR+0Q/OI
# vMawC9ZfloiYrL4DU1JCJMx+Z4S2p6+pQe3WDwKcMA0GCSqGSIb3DQEBAQUABIIC
# AGZ3yisPYgWCT8uTCFjLhIX6KGohRz+ozIXJ+PUX0qV2bDvRJ3M6PDFMeqOEMKZo
# G6ziQ1UoZdOOmpMc8wEI2KHOFpXYTXWoVHNsvbdmDvRqtookGxv9DB68ocUpSK5I
# GvVdzdBGKYVNbKKONCJdh62K5WAbj04MbBIK+YKXyaGJ0GbcE21JFMSGvBPLN6xQ
# yH7qkY50IwRj5A29smxsmDmvrtXZgxi4J5D7iF/IELEIA9oxMVh1GMUa/1xEhMEr
# MxwhxBZ8WNSuswOEcalUGflPr0x0YS2CuRvRnu1JvmXnNak0U6sNsjHTFYu/7t76
# HyGm03sk1S9I4xw13mB1sFdCkJbfO4vkKGcm6cwTgKQ9tpl+TqeoKzU9pdQelE8r
# UP6Nkoonln7D70g8/CvwOLT0nBE/JNyJgbUUVHjEaMpHUKsniDSN/OdJDdMNdL8F
# znFxTnKQYkBdDSiuPeexibpJhliL9D/aTgh4E+aYdvUUxyOujmLAw9Xj4O5Im4u/
# E/MAGKkMqRI+wlL9pL6qC2rfzM5pPNyPtwvkBpzu6Md/n7jcPJjgN8g6NrVuvs0+
# FMW2a0t4iBKCl/F0iLnze17yhkqeT5Gf9Pr8ywZq/PLMUv2VSdQX+iJ4o2lUnmqX
# ksbKepcnpTKxzDtjYmrnLtMwLuTymt7/ysEhYdNTYloQ
# SIG # End signature block
