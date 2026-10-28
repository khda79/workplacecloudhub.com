#Requires -Version 7.0
# Version: 0.1.3
[CmdletBinding()]
param()
# Synthetic inputs and local callbacks only: no collector, tenant or cloud connection.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'NativeEvidence.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'EndpointProtectionEvidence.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PreparedSharePointTransfer.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('UnifiedTransferTest-'+[guid]::NewGuid().ToString('N'))
$checks=0
function Check([bool]$Value,[string]$Message){if(-not $Value){throw $Message};$script:checks++}
function Json([string]$Path,$Value){[IO.File]::WriteAllText($Path,($Value|ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))}
try {
    $staging=Join-Path $root 'staging';$capture=Join-Path $root 'capture';$output=Join-Path $root 'data/DATA-POWERBI'
    New-Item -ItemType Directory -Path $staging,(Join-Path $capture 'DATA-LAST') -Force|Out-Null
    $basePath=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/prepared-evidence-contract.json')).Path
    $base=(Read-SmartM365JsonDocument $basePath).Document
    $nativePath=Join-Path $PSScriptRoot '../config/native-output-contract.json.txt'
    $contract=Merge-NativeEvidenceContract $base
    $migrationPath=Join-Path $PSScriptRoot '../config/sharepoint-migration-contract.json.txt'
    $migration=(Read-SmartM365JsonDocument $migrationPath).Document
    $contract.tables=@($contract.tables)+@($migration.tables)
    $extension=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/endpoint-protection-contract.json.txt')).Document
    $contract=Merge-EndpointProtectionContract $contract $extension
    $contractPath=Join-Path $root 'contract.json.txt';Json $contractPath $contract
    foreach($table in $contract.tables){
        $row=[ordered]@{}
        $tenantAware=$table.PSObject.Properties['stagingIncludesTenantKey'] -and $table.stagingIncludesTenantKey
        foreach($column in $table.columns){
            if($column.name -eq 'TenantKey'){if($tenantAware){$row[$column.name]='synthetic-unified'};continue}
            $row[$column.name]=switch($column.type){
                boolean {'True'} int64 {'1'} double {'0.5'} decimal {'0.5'} dateTime {'2026-01-01T00:00:00Z'}
                default {'Synthetic '+$column.name}
            }
        }
        [pscustomobject]$row|Export-Csv -LiteralPath (Join-Path $staging $table.file) -NoTypeInformation -Encoding utf8
    }
    $artifacts=Get-Content (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw|ConvertFrom-Json
    foreach($name in $artifacts.currentFiles){
        $path=Join-Path $capture "DATA-LAST/$name"
        if($name -like '*.csv'){@{TenantKey='synthetic-unified';SnapshotId=('a'*32)}|Export-Csv -LiteralPath $path -NoTypeInformation}
        else{Json $path @{TenantKey='synthetic-unified';SnapshotId=('a'*32);SchemaVersion=2;Products=@(1,2,3,4);OtherProducts=@(1,2,3);AssignmentIssues=@{Available=$false;Rows=@()}}}
    }
    foreach($history in $artifacts.history){
        $dir=Join-Path $capture "$($history.root)/2026-W40";New-Item -ItemType Directory -Path $dir -Force|Out-Null
        @{TenantKey='synthetic-unified';Id='one'}|Export-Csv -LiteralPath (Join-Path $dir $history.file) -NoTypeInformation
        Json (Join-Path $dir 'manifest.json.txt') @{Week='2026-W40';SnapshotCreatedAtUtc='2026-10-01T00:00:00Z'}
    }
    $provenance=@{
        NativeEvidence=@{ContractSHA256=(Get-FileHash $nativePath).Hash}
        SharePointMigration=@{ContractSHA256=(Get-FileHash $migrationPath).Hash}
        IntuneOperations=@{ContractSHA256=(Get-FileHash (Join-Path $PSScriptRoot '../config/intune-operational-contract.json.txt')).Hash}
        EndpointProtectionPolicy=@{PolicyVersion='synthetic-v1';SHA256=('a'*64)}
    }
    $batch=Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $output -TenantKey 'synthetic-unified' -Provenance $provenance -ContractPath $contractPath -ArtifactSnapshotRoot $capture -SkipRetention
    $manifest=Get-Content (Join-Path $batch.BatchPath 'batch.json.txt') -Raw|ConvertFrom-Json
    Check ($manifest.Files.Count -eq $contract.tables.Count) 'Unified table set incomplete'
    Check ($manifest.Artifacts.Count -eq $artifacts.currentFiles.Count+5) 'Artifact set incomplete'
    Check ($manifest.LicenseReportQualification.State -eq 'Not qualified' -and $manifest.LicenseReportQualification.Reason -like '*receipt*') 'Unproven license report must be declared Not qualified in the manifest'
    # Batches built before New-LicensingEvidence was retired keep their two license tables and stay transferable.
    $legacyLicense=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/legacy-license-evidence-tables.json.txt')).Document
    $legacyContract=$contract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    $legacyTables=[Collections.Generic.List[object]]::new();foreach($table in @($legacyContract.tables)){$legacyTables.Add($table)}
    foreach($entry in @($legacyLicense.Tables | Sort-Object {[int]$_.index})){$legacyTables.Insert([int]$entry.index,$entry.table)}
    $legacyContract.tables=$legacyTables.ToArray()
    $legacyPath=Join-Path $root 'legacy-license-contract.json.txt';Json $legacyPath $legacyContract
    $legacyManifest=$manifest | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    $legacyManifest.Files=@($legacyManifest.Files)+@($legacyLicense.Tables | ForEach-Object {[pscustomobject]@{File=$_.table.file}})
    $legacyManifest.ContractSHA256=(Get-FileHash -LiteralPath $legacyPath).Hash
    $resolved=& (Get-Module PreparedSharePointTransfer) {param($m,$p) Resolve-PreparedTransferContract -Manifest $m -ContractPath $p} $legacyManifest $basePath
    Check ($resolved.tables.Count -eq $contract.tables.Count+2 -and @($resolved.tables | Where-Object table -eq 'License Evidence').Count -eq 1) 'Pre-retirement license cohort no longer transferable'
    $partialManifest=$legacyManifest | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    $partialManifest.Files=@($partialManifest.Files | Where-Object File -ne 'LicenseOptimizationEvidence.csv')
    $rejected=$false;try{& (Get-Module PreparedSharePointTransfer) {param($m,$p) Resolve-PreparedTransferContract -Manifest $m -ContractPath $p} $partialManifest $basePath|Out-Null}catch{$rejected=$_.Exception.Message -like '*legacy license*'}
    Check $rejected 'Partial legacy license cohort accepted'
    # The deployed contract must still resolve a full native/endpoint cohort
    # published before the two daily trend tables existed.
    $preDailyContract=$contract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    $preDailyContract.tables=@($preDailyContract.tables | Where-Object table -notin @('Daily Directory Trends','Daily Mailbox Trends'))
    $preDailyPath=Join-Path $root 'pre-daily-contract.json.txt';Json $preDailyPath $preDailyContract
    $preDailyManifest=$manifest | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    $preDailyManifest.Files=@($preDailyManifest.Files | Where-Object File -notin @('DailyDirectoryTrends.csv','DailyMailboxTrends.csv'))
    foreach($name in @('Daily Directory Trends','Daily Mailbox Trends')){$preDailyManifest.HistoryKeyVersions.PSObject.Properties.Remove($name)}
    $preDailyManifest.ContractSHA256=(Get-FileHash -LiteralPath $preDailyPath).Hash
    $resolved=& (Get-Module PreparedSharePointTransfer) {param($m,$p) Resolve-PreparedTransferContract -Manifest $m -ContractPath $p} $preDailyManifest $basePath
    Check ($resolved.tables.Count -eq $contract.tables.Count-2) 'Pre-daily unified contract no longer transferable'
    $events=[Collections.Generic.List[string]]::new();$remote=Join-Path $root 'simulated-remote'
    $state=@{Corrupt=$false}
    $upload={param($local,$relative)
        $events.Add('U:'+$relative);$target=Join-Path $remote $relative
        New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force|Out-Null
        Copy-Item -LiteralPath $local -Destination $target;$true
    }.GetNewClosure()
    $download={param($destination,$relative)
        $events.Add('D:'+$relative);Copy-Item -LiteralPath (Join-Path $remote $relative) -Destination $destination
        if($state.Corrupt -and $relative.EndsWith('/Source-M365_Licenses_ReportSnapshot.json.txt')){Add-Content -LiteralPath $destination 'tampered'}
        $true
    }.GetNewClosure()
    $parameters=@{OutputRoot=$output;TenantKey='synthetic-unified';ContractPath=$basePath;WorkRoot=(Join-Path $root 'scratch');ExpectedBatchId=$batch.BatchId;UploadFile=$upload;DownloadFile=$download}
    $result=Send-PreparedEvidenceBatch @parameters 6>$null
    Check ($result.PointerVerified -and -not $result.CsvRecalculated) 'Unverified transfer result'
    Assert-PreparedTransferResult -Result $result -ExpectedBatchId $batch.BatchId
    foreach($mode in 'ExitOnly','Partial','FalseText'){
        $badReceipt=if($mode -eq 'ExitOnly'){[pscustomobject]@{ExitCode=0}}else{$result|ConvertTo-Json|ConvertFrom-Json}
        if($mode -eq 'Partial'){$badReceipt.VerifiedFiles--}
        if($mode -eq 'FalseText'){$badReceipt.PointerVerified='false'}
        $rejected=$false;try{Assert-PreparedTransferResult -Result $badReceipt -ExpectedBatchId $batch.BatchId}catch{$rejected=$true}
        Check $rejected 'Zero exit or partial publication receipt accepted'
    }
    $rejected=$false;try{Assert-PreparedTransferResult -Result $result -ExpectedBatchId 'different-batch'}catch{$rejected=$true}
    Check $rejected 'Wrong-batch transfer receipt accepted'
    Check ($result.VerifiedFiles -eq $manifest.Files.Count+$manifest.Artifacts.Count+4) 'Missing unified read-back'
    Check ($events[-2] -eq 'U:current.json.txt' -and $events[-1] -eq 'D:current.json.txt') 'Pointer was not last'
    foreach($entry in @($manifest.Files)+@($manifest.Artifacts)){
        Check ((Get-FileHash -LiteralPath (Join-Path $remote "batches/$($batch.BatchId)/$($entry.File)")).Hash -eq $entry.SHA256) 'Simulated remote bytes differ'
    }
    $oldPointer=(Get-FileHash -LiteralPath (Join-Path $remote 'current.json.txt')).Hash
    $events.Clear();$state.Corrupt=$true;$failed=$false
    try{Send-PreparedEvidenceBatch @parameters 6>$null|Out-Null}catch{$failed=$true}
    Check $failed 'Read-back corruption accepted'
    Check (-not $events.Contains('U:current.json.txt')) 'Corrupt artifact advanced root pointer'
    Check ((Get-FileHash -LiteralPath (Join-Path $remote 'current.json.txt')).Hash -eq $oldPointer) 'Last valid remote pointer changed'
    # Required artifact failure must preserve the local pointer too.
    $localPointer=(Get-FileHash -LiteralPath (Join-Path $output 'current.json.txt')).Hash
    Json (Join-Path $capture 'DATA-LAST/M365_Licenses_ReportSnapshot.json.txt') @{TenantKey='foreign';SnapshotId=('a'*32);SchemaVersion=2}
    $failed=$false
    try{Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $output -TenantKey 'synthetic-unified' -Provenance $provenance -ContractPath $contractPath -ArtifactSnapshotRoot $capture -SkipRetention|Out-Null}catch{$failed=$true}
    Check $failed 'Foreign snapshot published'
    Check ((Get-FileHash -LiteralPath (Join-Path $output 'current.json.txt')).Hash -eq $localPointer) 'Failed publication advanced local pointer'
    Check (@(Get-ChildItem (Join-Path $output 'batches') -Directory).Count -eq 1) 'Failed candidate payload was not cleaned'
    # A Not qualified license report after a qualified one is a manifest state, never a coverage regression.
    $last=Join-Path $capture 'DATA-LAST'
    foreach($name in $artifacts.currentFiles){
        $path=Join-Path $last $name
        if($name -like '*.csv'){@{TenantKey='synthetic-unified';SnapshotId=('a'*32)}|Export-Csv -LiteralPath $path -NoTypeInformation}
    }
    $runId='synthetic-licensing-run'
    $sources=@(foreach($group in @(
        @{Producer='SmartM365-M365UserActivity-Inventory.ps1';Files=@('M365_Users_Activity.csv','M365_Mailbox_Usage.csv','M365_Email_Activity.csv','M365_Apps_Usage_180D.csv','M365_OneDrive_Usage.csv')},
        @{Producer='SmartM365-Devices-Inventory.ps1';Files=@('Intune_Devices_Inventory.csv')}
    )) {
        $records=@(foreach($name in $group.Files) {
            [pscustomobject]@{UserId='contoso-one'} | Export-Csv -LiteralPath (Join-Path $last $name) -NoTypeInformation
            @{File=$name;RunId='source-run';Status='Success';Errors=0;IsPartialInventory=$false;SHA256=(Get-FileHash (Join-Path $last $name)).Hash}
        })
        $stem='SmartInventory_'+[IO.Path]::GetFileNameWithoutExtension($group.Producer)
        Json (Join-Path $last ($stem+'.current.json.txt')) @{Owner='SmartInventory-SourceReceipt';Producer=$group.Producer;TenantKey='synthetic-unified';RunId='source-run';Status='Completed';Errors=0;IsPartialInventory=$false;ConsumerScopeQualified=$true;PublicationProtocol=1;CompletedAtUtc='2026-10-17T22:00:00Z';Files=$records}
        Json (Join-Path $last ($stem+'.run.json.txt')) @{Owner='SmartInventory-SourceRun';Producer=$group.Producer;TenantKey='synthetic-unified';RunId='source-run';Status='Completed';UnqualifiedPublication=$false}
        foreach($record in $records) { @{Name=$record.File;Ready=$true;Provisional=$false;SHA256=$record.SHA256} }
    })
    # This test writes the dependencies straight into the capture; Prepare only copies contract files.
    $sourceContract=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/prepared-source-contract.json.txt')).Document
    $capturedFiles=@($sourceContract.currentFiles)+@($sourceContract.optionalCurrentFiles | ForEach-Object file)
    foreach($source in $sources){Check ($capturedFiles -ccontains $source.Name) "License dependency never captured by Prepare: $($source.Name)"}

    Json (Join-Path $last 'M365_Licenses_ReportSnapshot.json.txt') @{GeneratedAtUtc='2026-10-17T23:00:00Z';Sources=$sources;TenantKey='synthetic-unified';SnapshotId=('a'*32);SchemaVersion=2;Products=@(1,2,3,4);OtherProducts=@(1,2,3);AssignmentIssues=@{Available=$false;Rows=@()};ReceiptRun=@{Producer='SmartM365-Licences-Inventory.ps1';RunId=$runId}}
    $reportFiles=@($artifacts.licenseReport.reportFiles)
    Json (Join-Path $last $artifacts.licenseReport.receipt) @{Producer='SmartM365-Licences-Inventory.ps1';TenantKey='synthetic-unified';Status='Completed';RunId=$runId
        Files=@($reportFiles | ForEach-Object {@{File=$_;SHA256=(Get-FileHash -LiteralPath (Join-Path $last $_)).Hash}});OptionalFiles=@($reportFiles | ForEach-Object {@{File=$_;Published=$true}})}
    $qualified=Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $output -TenantKey 'synthetic-unified' -Provenance $provenance -ContractPath $contractPath -ArtifactSnapshotRoot $capture -SkipRetention
    $qualifiedManifest=Get-Content (Join-Path $qualified.BatchPath 'batch.json.txt') -Raw|ConvertFrom-Json
    Check ($qualifiedManifest.LicenseReportQualification.State -eq 'Qualified' -and $qualifiedManifest.LicenseReportQualification.ReceiptRunId -eq $runId) 'Recorded license report not qualified'
    Add-Content -LiteralPath (Join-Path $last 'Intune_Devices_Inventory.csv') -Value 'newer device capture'
    $stale=Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $output -TenantKey 'synthetic-unified' -Provenance $provenance -ContractPath $contractPath -ArtifactSnapshotRoot $capture -SkipRetention 3>$null
    $staleManifest=Get-Content (Join-Path $stale.BatchPath 'batch.json.txt') -Raw|ConvertFrom-Json
    Check ($staleManifest.LicenseReportQualification.State -eq 'Not qualified' -and $staleManifest.LicenseReportQualification.Reason -like '*different source capture*') 'Old licensing snapshot survived a newer device capture'
    Check ($staleManifest.Files.Count -eq $qualifiedManifest.Files.Count -and (Get-Content (Join-Path $output 'current.json.txt') -Raw|ConvertFrom-Json).BatchId -eq $stale.BatchId) 'A stale licensing report blocked unrelated batch evidence'
    Remove-Item -LiteralPath (Join-Path $last $artifacts.licenseReport.receipt)
    $unqualified=Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $output -TenantKey 'synthetic-unified' -Provenance $provenance -ContractPath $contractPath -ArtifactSnapshotRoot $capture -SkipRetention 3>$null
    $unqualifiedManifest=Get-Content (Join-Path $unqualified.BatchPath 'batch.json.txt') -Raw|ConvertFrom-Json
    Check ($unqualifiedManifest.LicenseReportQualification.State -eq 'Not qualified' -and (Get-Content (Join-Path $output 'current.json.txt') -Raw|ConvertFrom-Json).BatchId -eq $unqualified.BatchId) 'Not qualified license report blocked the batch'
    Check (@($contract.tables | Where-Object {$_.file -like 'License*' -and $_.PSObject.Properties['historyKey']}).Count -eq 0) 'A license table declares a historical key'
    'Unified publication/transfer offline checks passed: '+$checks
} finally {
    $full=[IO.Path]::GetFullPath($root);$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if($full.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $full -Leaf) -match '^UnifiedTransferTest-[a-f0-9]{32}$'){
        if(Test-Path -LiteralPath $full){Remove-Item -LiteralPath $full -Recurse -Force}
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD7VeHrkWldFEq8
# h2V+Fcf7QTPGv5p7pm88iLikgk8QuaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICQZPuDvvGOv0zeDRgWxylJ0glgK5p+u/wu/Zp/O3d2oMA0GCSqG
# SIb3DQEBAQUABIIBgAK6eXvg2pyqRxZXm7aPjWMgueszreGHSvxpHWax4G4GnaSl
# fjzyuodyfqkBzFUv2ohD1LbsbfWqHf/LY9urp3XVfDxbVx+nUf3+2HsLE3OSDv5/
# ExrwUTYy/2dosHIUFuwHUYux2VsQI8v/Kb1pRCEz7pc+iHRIb2ikTQ7Vu/7AMVVL
# iDIHc+zVoU9D2T7vTOQjBi+Fu843EzFZqVCINYGuhagjGS3UBeXn4ZjTbgqwZnOF
# sR1fvQoPScHD/XVyD50mWi+UpWiHn5IS2rvP4H0vs6xEijISs+P6Fd2TbvQkneN1
# ++VMHCDf/+3WAEVqwntjkTYLn2PHOqMjiTNrvCUCCS0uTIrwftq1HETWVvUSkvqo
# FD+sKiPVJShec48AUAGbiojznpF9mcqtggd1YMNt2qOBaJEFMPKDClOig4ZvSCWs
# ALMISHcaAfcUfwtYive3WGkUfcuqkU08AGPTj2p2M0845MAMF6vjkpwsbtVgjIbU
# orRAofQvfn3w1TCZkaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxNTQ4
# NTNaMC8GCSqGSIb3DQEJBDEiBCByWFWT8GhTzsmyVbab5rvD8bKJ1g2thT+kd8br
# 7Pc3DjANBgkqhkiG9w0BAQEFAASCAgAjss5SObeZxPjRzrh8aLN4jYfoq95W2cqN
# V/rhKgIET28iUx0VhFObGH9XWcWtc4keD415QlhIyUCtnOWdxL7p4RI3gnXQ0qRp
# s9senu2z24SnARnIZWDTmvk2laLIlqNm/ahEhLhKMuxC86heP9jwgRQmH6DgvnII
# YF/4Yhifv4rowL8wdltJ3tHEl5Eovnx1gDdMOY2iJ3Y4RR/8pel4QC7TJbzCrQSq
# kxHMW/DLTpbCiWu/LZx336dOwkRURQ4n52NTDFbmAN0hlNLA6tZullEFNI5r8S66
# ya/tA0xaP1LA5qIcfa+Tcd7BFwmmQmwZLmW4mSIBxA7Akjt5BH+DlD6M5jqDb+3u
# VNhRiQcG8S/wDoIL9CIisij0VOSDwz3dqJbHZfgF0gUhm/b49tjbumqekYSBJNKj
# t8TF/TV1TQ39trekLXdgpc1Dctr5B8FpN3p6IHev6pQU75INCuqgIohbpTDaT64T
# C1CRpiZo4OeQCnxPTrlCdAyywXn3fDB/8v3IhUorLE+2gDbNk4+5dlILxupTt3ch
# u8QDNcQwMHsDf+pMBpeszQWoJaMCrg2yhVJilI5G41ozDsEm8K5pn7NMmOdaNh3R
# th4DyF/MVhNsLhs69hSSCJK1FxSKc+jG8AiESqQSu04PwPWwhqfHgmtzZD1H31SG
# 8j9YPsh5hg==
# SIG # End signature block
