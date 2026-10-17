#Requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Synthetic fixtures and filesystem callbacks only. No tenant, Graph or SharePoint calls.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PreparedSharePointTransfer.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'EndpointProtectionEvidence.psm1') -Force
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$basePath=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/prepared-evidence-contract.json')).Path
$baseContract=(Read-SmartM365JsonDocument $basePath).Document
$extension=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/endpoint-protection-contract.json.txt')).Document
$extendedContract=Merge-EndpointProtectionContract $baseContract $extension
$extendedPath=Join-Path $root 'extended-contract.json.txt'
function Write-TestJson([string]$Path,$Value){
    [IO.File]::WriteAllBytes($Path,[Text.UTF8Encoding]::new($false).GetBytes(($Value | ConvertTo-Json -Depth 20)))
}
Write-TestJson $extendedPath $extendedContract
$script:checks=0
function Check([bool]$Value,[string]$Message){$script:checks++;if(-not $Value){throw $Message}}
function Reject([scriptblock]$Action,[string]$Pattern){
    $caught=$null;try{& $Action | Out-Null}catch{$caught=$_}
    Check ($null -ne $caught) 'Expected rejection'
    Check ($caught.Exception.Message -match $Pattern) "Expected '$Pattern', got: $caught"
}
$intuneSpec=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/intune-operational-contract.json.txt')).Document
$baseCount=@($baseContract.tables).Count
$extendedCount=@($extendedContract.tables).Count
Check ($baseCount -eq 29+$intuneSpec.tables.Count) 'Expected legacy tables plus governed Intune and daily trend tables (SWI license evidence retired)'
Check (@($baseContract.tables | Where-Object table -in @('Daily Directory Trends','Daily Mailbox Trends')).Count -eq 2) 'Expected both governed daily trend tables'
Check ($extendedCount -eq $baseCount+2) 'Expected endpoint protection extension'
function New-Fixture([string]$Name,$Contract,[string]$ContractPath,[bool]$Endpoint){
    $fixtureRoot=Join-Path $root $Name
    $staging=Join-Path $fixtureRoot 'staging'
    $output=Join-Path $fixtureRoot 'data/DATA-POWERBI'
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    foreach($table in $Contract.tables){
        $row=[ordered]@{}
        $tenantAware=$table.PSObject.Properties['stagingIncludesTenantKey'] -and $table.stagingIncludesTenantKey
        foreach($column in $table.columns){
            if($column.name -eq 'TenantKey'){
                if($tenantAware){$row[$column.name]='synthetic-transfer'}
                continue
            }
            $row[$column.name]=switch($column.type){
                'boolean' {'True'}
                'int64' {'1'}
                'double' {'0.5'}
                'decimal' {'0.5'}
                'dateTime' {'2026-01-01T00:00:00Z'}
                default {'Synthetic '+$column.name}
            }
        }
        [pscustomobject]$row | Export-Csv -LiteralPath (Join-Path $staging $table.file) -NoTypeInformation -Encoding utf8
    }
    $provenance=@{Synthetic=$true}
    if(@($Contract.tables | Where-Object file -in $intuneSpec.tables.file).Count){$provenance.IntuneOperations=@{ContractSHA256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot '../config/intune-operational-contract.json.txt')).Hash}}
    if($Endpoint){$provenance.EndpointProtectionPolicy=@{PolicyVersion='synthetic-v1';SHA256=('A'*64);Qualification='Synthetic only'}}
    $batch=Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $output -TenantKey 'synthetic-transfer' -Provenance $provenance -ContractPath $ContractPath -SkipRetention
    [pscustomobject]@{
        Root=$fixtureRoot;Output=$output;Batch=$batch
        Manifest=Join-Path $batch.BatchPath 'batch.json.txt'
        Validation=Join-Path $batch.BatchPath 'validation.json.txt'
        Pointer=Join-Path $output 'current.json.txt'
        Receipt=Join-Path $batch.BatchPath 'current.json.txt'
    }
}
function Snapshot-Fixture($Fixture){
    @(@($Fixture.Pointer)+@(Get-ChildItem -LiteralPath $Fixture.Batch.BatchPath -File | ForEach-Object FullName) |
        Sort-Object | ForEach-Object {$_+'='+(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash}) -join '|'
}
function New-Transport($Fixture,[string]$Mode='ok'){
    $remote=Join-Path $root ('simulated-remote/'+[guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $remote -Force | Out-Null
    $state=@{Remote=$remote;Mode=$Mode;Events=[Collections.Generic.List[string]]::new()}
    $upload={param($local,$relative)
        $state.Events.Add('U:'+$relative)
        $target=Join-Path $state.Remote $relative
        New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $local -Destination $target
        $true
    }.GetNewClosure()
    $download={param($destination,$relative)
        $state.Events.Add('D:'+$relative)
        Copy-Item -LiteralPath (Join-Path $state.Remote $relative) -Destination $destination
        if($state.Mode -eq 'corrupt-endpoint' -and $relative.EndsWith('/EndpointProtectionEvidence.csv')){
            Add-Content -LiteralPath $destination -Value 'Synthetic corruption'
        }
        $true
    }.GetNewClosure()
    @{State=$state;Params=@{
        OutputRoot=$Fixture.Output;TenantKey='synthetic-transfer';ContractPath=$basePath
        WorkRoot=Join-Path $Fixture.Root 'scratch';ExpectedBatchId=$Fixture.Batch.BatchId
        UploadFile=$upload;DownloadFile=$download
    }}
}
function Check-Success($Fixture,[int]$Count,[string]$ContractPath=$basePath){
    $before=Snapshot-Fixture $Fixture
    $t=New-Transport $Fixture;$params=$t.Params;$params.ContractPath=$ContractPath
    $result=Send-PreparedEvidenceBatch @params 6>$null
    Check ($result.CsvFiles -eq $Count -and $result.VerifiedFiles -eq ($Count+4) -and $result.PointerVerified) 'Unexpected successful transfer receipt'
    Check (-not $result.CsvRecalculated) 'Transfer recalculated CSVs'
    Check ($t.State.Events[-2] -eq 'U:current.json.txt' -and $t.State.Events[-1] -eq 'D:current.json.txt') 'Root pointer not transferred last'
    Check ((Snapshot-Fixture $Fixture) -eq $before) 'Transfer changed local batch or pointer bytes'
    Check ((Get-FileHash -LiteralPath (Join-Path $t.State.Remote 'current.json.txt')).Hash -eq (Get-FileHash -LiteralPath $Fixture.Pointer).Hash) 'Remote pointer differs'
    if(Test-Path -LiteralPath (Join-Path $Fixture.Batch.BatchPath 'EndpointProtectionEvidence.csv')){
        foreach($name in 'EndpointProtectionPolicy.csv','EndpointProtectionEvidence.csv'){
            Check (Test-Path -LiteralPath (Join-Path $t.State.Remote "batches/$($Fixture.Batch.BatchId)/$name")) "Endpoint CSV absent from simulated transfer: $name"
        }
    }
    $result
}
function Check-RejectedBeforeTransport($Fixture,[string]$Pattern,[hashtable]$Overrides=@{}){
    $before=Snapshot-Fixture $Fixture
    $t=New-Transport $Fixture;$params=$t.Params
    foreach($key in $Overrides.Keys){$params[$key]=$Overrides[$key]}
    Reject {Send-PreparedEvidenceBatch @params} $Pattern
    Check ($t.State.Events.Count -eq 0) 'Rejected batch invoked a transport callback'
    Check (-not (Test-Path -LiteralPath (Join-Path $t.State.Remote 'current.json.txt'))) 'Rejected batch advanced remote pointer'
    Check ((Snapshot-Fixture $Fixture) -eq $before) 'Rejected transfer changed local batch or pointer'
}
function With-MutatedReceipt($Fixture,[scriptblock]$Change,[scriptblock]$Action){
    $saved=@{}
    foreach($path in @($Fixture.Manifest,$Fixture.Validation,$Fixture.Pointer,$Fixture.Receipt)){$saved[$path]=[IO.File]::ReadAllBytes($path)}
    try{
        $manifest=Get-Content -LiteralPath $Fixture.Manifest -Raw | ConvertFrom-Json
        $validation=Get-Content -LiteralPath $Fixture.Validation -Raw | ConvertFrom-Json
        & $Change $manifest $validation
        Write-TestJson $Fixture.Manifest $manifest
        Write-TestJson $Fixture.Validation $validation
        $pointer=Get-Content -LiteralPath $Fixture.Pointer -Raw | ConvertFrom-Json
        $pointer.ManifestSHA256=(Get-FileHash -LiteralPath $Fixture.Manifest).Hash
        Write-TestJson $Fixture.Pointer $pointer
        Write-TestJson $Fixture.Receipt $pointer
        & $Action
    }finally{foreach($path in $saved.Keys){[IO.File]::WriteAllBytes($path,$saved[$path])}}
}
$base=New-Fixture 'base' $baseContract $basePath $false
$extended=New-Fixture 'extended' $extendedContract $extendedPath $true
$null=Check-Success $base $baseCount
$first=Check-Success $extended $extendedCount
$again=Check-Success $extended $extendedCount
Check ($again.BatchId -eq $first.BatchId) 'Retry switched batches'
$null=Check-Success $extended $extendedCount $extendedPath
# Base batches remain transferable after extended batches; no live tenant policy is read.
$null=Check-Success $base $baseCount
$beforeDaily=$baseContract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
$beforeDaily.tables=@($beforeDaily.tables | Where-Object table -notin @('Daily Directory Trends','Daily Mailbox Trends'))
$beforeDailyPath=Join-Path $root 'before-daily-contract.json.txt';Write-TestJson $beforeDailyPath $beforeDaily
$beforeDailyFixture=New-Fixture -Name 'before-daily' -Contract $beforeDaily -ContractPath $beforeDailyPath -Endpoint $false
$null=Check-Success -Fixture $beforeDailyFixture -Count ($baseCount-2)
$legacyBase=$baseContract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
# Pre-Intune cohorts were built before New-LicensingEvidence was retired: restore its two tables at their original positions.
$legacyLicense=Get-Content (Join-Path $PSScriptRoot '../config/legacy-license-evidence-tables.json.txt') -Raw | ConvertFrom-Json
$legacyTables=[Collections.Generic.List[object]]::new();foreach($table in @($legacyBase.tables)){$legacyTables.Add($table)}
foreach($entry in @($legacyLicense.Tables | Sort-Object {[int]$_.index})){$legacyTables.Insert([int]$entry.index,$entry.table)}
$legacyBase.tables=@($legacyTables | Where-Object { $_.file -notin $intuneSpec.tables.file -and $_.table -notin @('Daily Directory Trends','Daily Mailbox Trends') })
$legacyPath=Join-Path $root 'legacy-contract.json.txt';Write-TestJson $legacyPath $legacyBase
Import-Module (Join-Path $PSScriptRoot 'EndpointProtectionEvidence.psm1') -Global
$legacyExtended=Merge-EndpointProtectionContract $legacyBase $extension
$legacyExtendedPath=Join-Path $root 'legacy-extended-contract.json.txt';Write-TestJson $legacyExtendedPath $legacyExtended
$oldBase=New-Fixture 'legacy-base' $legacyBase $legacyPath $false
$oldExtended=New-Fixture 'legacy-extended' $legacyExtended $legacyExtendedPath $true
$null=Check-Success $oldBase 29
$null=Check-Success $oldExtended 31
With-MutatedReceipt -Fixture $base -Change {param($m,$v)
    $m.Files=@($m.Files | Where-Object File -ne 'DailyMailboxTrends.csv')
    $v.Files=@($v.Files | Where-Object File -ne 'DailyMailboxTrends.csv')
} -Action {Check-RejectedBeforeTransport $base 'Incomplete daily trends file set'}
With-MutatedReceipt -Fixture $base -Change {param($m,$v)
    $m.Files=@($m.Files | Where-Object File -notin @('DailyDirectoryTrends.csv','DailyMailboxTrends.csv'))
    $v.Files=@($v.Files | Where-Object File -notin @('DailyDirectoryTrends.csv','DailyMailboxTrends.csv'))
} -Action {Check-RejectedBeforeTransport $base 'Incomplete daily trends file set'}
With-MutatedReceipt $base {param($m,$v)$m.Provenance.PSObject.Properties.Remove('IntuneOperations')} {
    Check-RejectedBeforeTransport $base 'Intune operational batch lacks'
}
With-MutatedReceipt $base {param($m,$v)$m.Files=@($m.Files | Where-Object File -ne 'IntuneManagedAppDeploymentEvidence.csv');$v.Files=@($v.Files | Where-Object File -ne 'IntuneManagedAppDeploymentEvidence.csv')} {
    Check-RejectedBeforeTransport $base 'Incomplete Intune operational file set'
}
With-MutatedReceipt $extended {param($m,$v) $m.Provenance.PSObject.Properties.Remove('EndpointProtectionPolicy')} {
    Check-RejectedBeforeTransport $extended 'policy provenance'
}
With-MutatedReceipt $extended {param($m,$v) $m.Provenance.EndpointProtectionPolicy.SHA256='invalid'} {
    Check-RejectedBeforeTransport $extended 'policy provenance'
}
With-MutatedReceipt $extended {param($m,$v) $m.PSObject.Properties.Remove('ContractSHA256')} {
    Check-RejectedBeforeTransport $extended 'lacks a contract hash'
}
With-MutatedReceipt $extended {param($m,$v) $m.ContractSHA256=('0'*64)} {
    Check-RejectedBeforeTransport $extended 'contract hash mismatch'
}
With-MutatedReceipt $extended {param($m,$v)
    $m.Files=@($m.Files | Where-Object File -ne 'EndpointProtectionEvidence.csv')
    $v.Files=@($v.Files | Where-Object File -ne 'EndpointProtectionEvidence.csv')
} {Check-RejectedBeforeTransport $extended 'Incomplete endpoint protection file set'}
With-MutatedReceipt $base {param($m,$v)
    $m.Provenance | Add-Member NoteProperty EndpointProtectionPolicy ([pscustomobject]@{PolicyVersion='synthetic-v1';SHA256=('A'*64)})
} {Check-RejectedBeforeTransport $base 'Incomplete endpoint protection file set'}
With-MutatedReceipt $extended {param($m,$v)
    ($m.Files | Where-Object File -eq 'EndpointProtectionEvidence.csv').File='EndpointProtectionPolicy.csv'
} {Check-RejectedBeforeTransport $extended 'Incomplete endpoint protection file set'}
With-MutatedReceipt $extended {param($m,$v)
    ($m.Files | Where-Object File -eq 'DeviceSystemEvidence.csv').File='Unreviewed.csv'
    ($v.Files | Where-Object File -eq 'DeviceSystemEvidence.csv').File='Unreviewed.csv'
} {Check-RejectedBeforeTransport $extended 'Unexpected, unsafe or duplicate CSV name'}
With-MutatedReceipt $extended {param($m,$v) $v.Files=@($v.Files | Select-Object -Skip 1)} {
    Check-RejectedBeforeTransport $extended 'file count differs'
}
With-MutatedReceipt $extended {param($m,$v) $v.Passed=$false} {
    Check-RejectedBeforeTransport $extended 'successful full validation receipt'
}
Check-RejectedBeforeTransport $extended 'pointer identity|tenant ownership mismatch' @{TenantKey='other-synthetic'}
Check-RejectedBeforeTransport $extended 'ExpectedBatchId' @{ExpectedBatchId='20260101T000000000Z-aaaaaaaa'}
$changedBase=$baseContract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
$changedBase.tables[0].columns[0].type='boolean'
$changedPath=Join-Path $root 'changed-base.json.txt';Write-TestJson $changedPath $changedBase
Check-RejectedBeforeTransport $extended 'contract hash mismatch' @{ContractPath=$changedPath}
$incomplete=$baseContract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
$incomplete.tables=@($incomplete.tables)+@($extension.tables[0])
$incompletePath=Join-Path $root 'incomplete-contract.json.txt';Write-TestJson $incompletePath $incomplete
Check-RejectedBeforeTransport $extended 'Incomplete supplied endpoint protection contract' @{ContractPath=$incompletePath}
$localEndpoint=Join-Path $extended.Batch.BatchPath 'EndpointProtectionEvidence.csv'
$original=[IO.File]::ReadAllBytes($localEndpoint)
try{
    Add-Content -LiteralPath $localEndpoint -Value 'Synthetic corruption'
    Check-RejectedBeforeTransport $extended 'Local CSV integrity mismatch'
}finally{[IO.File]::WriteAllBytes($localEndpoint,$original)}
$before=Snapshot-Fixture $extended
$t=New-Transport $extended 'corrupt-endpoint';$params=$t.Params
Reject {Send-PreparedEvidenceBatch @params} 'Remote SHA256/size mismatch'
Check (-not(Test-Path -LiteralPath (Join-Path $t.State.Remote 'current.json.txt'))) 'Corrupt endpoint read-back advanced pointer'
Check ((Snapshot-Fixture $extended) -eq $before) 'Failed endpoint read-back modified local batch'
Write-Host "PASS: $script:checks synthetic base/endpoint contract, retry, provenance, hash, file-set and integrity checks. No tenant or network calls."

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCClLMl9Qk4Y1zD3
# cUAR7E1jT/VmRepBDl4GOhbg0YudxKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBHvyMWGr6nocj7aJrKw/SV
# jpKCtSqXqkYqp5THeRrK0TANBgkqhkiG9w0BAQEFAASCAYB0ls11xp66cs8DWv45
# dM3in+Mos/SndKTq0BCcWOuXvlFUutha/OwPQ4vVDyw0GimE58mKHGij6SYUx8/w
# PgB7O6X7zdhUcRmunm57zrxjTjo3Lp/9lCiCF4XJwQAjACkut1CM/hnfqOc4W3U8
# JZvy8t2D3XqOyZHtYP0CSIctJ7gi1pcMKAe1Q7632A0EoPNeL+G+mfIcRk3A7gCU
# TnwwvM2tMRmYuvuDkwpXZLl1000yUzH8K2zK3wjnPRemhAyFEyhiwz14AB8LwvNX
# IfwGjGFvJoE3c/NmJMctI+jlNL5/+pCPCKiK5QAkH4wThJyph/yP9oQeyK6yw7oi
# vXohHEj/sDR6NZ2bfWApgx6wONMXG5B1DoYuI6EVIga3xafQGxxIg4qAY5TcUdN0
# XnMwYG9ls50fnzkySw8deM0Nl/BR6lCDqRFQMX308uaPETjFEqbgAQLHV3UWWXN9
# k41be2NTgFrVV2h6JyorNWb05c9qMBjVi4rq7rdiEnsj2mo=
# SIG # End signature block
