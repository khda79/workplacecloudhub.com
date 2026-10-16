#Requires -Version 7.0
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
        IntuneOperations=@{ContractSHA256=(Get-FileHash (Join-Path $PSScriptRoot '../config/intune-operational-contract.json.txt')).Hash}
        EndpointProtectionPolicy=@{PolicyVersion='synthetic-v1';SHA256=('a'*64)}
    }
    $batch=Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $output -TenantKey 'synthetic-unified' -Provenance $provenance -ContractPath $contractPath -ArtifactSnapshotRoot $capture -SkipRetention
    $manifest=Get-Content (Join-Path $batch.BatchPath 'batch.json.txt') -Raw|ConvertFrom-Json
    Check ($manifest.Files.Count -eq $contract.tables.Count) 'Unified table set incomplete'
    Check ($manifest.Artifacts.Count -eq $artifacts.currentFiles.Count+5) 'Artifact set incomplete'
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
    'Unified publication/transfer offline checks passed: '+$checks
} finally {
    $full=[IO.Path]::GetFullPath($root);$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if($full.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $full -Leaf) -match '^UnifiedTransferTest-[a-f0-9]{32}$'){
        if(Test-Path -LiteralPath $full){Remove-Item -LiteralPath $full -Recurse -Force}
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOV0N+Ux38EHXZ
# odr95AHSy1zGp8N6yW0jjYnvRTbfKaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDMHPwS0YkYIEXJ0NftFZvY
# b8a29fAenoB0pEgL4cpUrzANBgkqhkiG9w0BAQEFAASCAYAfst3eB+ec0uYHZ3y9
# qebrITIijN5Czpok7ky07JkBLV6djm3E9Sj12Fz0u4ub78DsFYOFkOtEzENUcJYO
# nCAD/vd5YGxRkLW68vx6Mep/FXNg9r3FU4IpuCljZ3ZzbuAU4dYRB1u8ImXrvfI7
# 5JzSlwX2pdWuReEryZy0NvaoiPt+DyIKMneoivzk5CeoMXj/k3/ZF70me+p8hwud
# DzjKT+L8NurVbDGkdHCzbVN5JwofVh7f1XugKoBy7XnkhxEWcD3qmVK0mmFosotT
# 9hTqe0S7BuOa3Dk4LsK8bSYNTx2W8o5lJ0Fj7LbDv79M3IhRgN17HvdAynBRlNMh
# VkfZ9rSQWDB6Fz2MF61JTAFdQAd+aQfTfGrMNIsFrLPRGCzwI+H+X0AiwGR0umJM
# E+GlEAgx+O3EN9ZOqY4rr5Xb2iyDKcIrmFJSp6595QoGGFQC3HtJj0oexsXMQHnd
# U2DCI6U2EVhguyFoW9N38f5ZteJaogez0iTrSzBIImtlVGQ=
# SIG # End signature block
