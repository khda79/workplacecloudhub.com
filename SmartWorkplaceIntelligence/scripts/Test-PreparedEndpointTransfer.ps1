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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBFUxp9/8/Fi32M
# DuMGJDeAcvkQTkcz/tSyTfYzbrEn+qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIK/y//EUqImzLipHTni9uDnU3z6qs8FidN1ZJ/wvpbjDMA0GCSqG
# SIb3DQEBAQUABIIBgCBtnRxU56SHiCuVI2Bx91/153xkhOlDlVNvZiEiwAwJy7CZ
# ipRwuPFKPhXisq7hlMBxP1xMswS0LQ3wsNMhxwxyiLJDRzo3dpJguPxDjXWsaxjw
# 3dBvLmsxsT3tARS8lMFEBXdaJh1EvxtHi+WedgohapToytMjhuVGm7//Go8aVzqD
# vjKvLbByWQwYQeaAxeC81GhWIGbVyCsuch61PTZG5jXHHnkT2ut1hghObTJBnJ5K
# xo3H5aPiteB78do1PUCwIBZq8n6NbkWWj7ShW3Mu0977AOTl6KKeelG9w8HYE3IH
# mAYYo18lhNgNnvSs2Ig/PkUEHkP/q8F12WM6iWIUuQOokD/XEXYUWjYilO6oEPXl
# x5kIKaF+S2daN+K4P7yzUnXvA8oBFURTuQDgsAxTXVRrA5MCfyqNvY2NCUbu1TzG
# XfyFtMI+OlyJ0GqxgXdDkttDc0zId5ZvBRApjMv/N2wKuW8qzTZVPU2UqfFumPPh
# Iz+RsH8jIreYePkO+qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTcyMzE1
# NTdaMC8GCSqGSIb3DQEJBDEiBCAxbUEhzhfjYyv6pyIhbsSRnhur0aC+nMrRJtVx
# Eo5wTTANBgkqhkiG9w0BAQEFAASCAgBG0zSSQMYD+hA6afrvJsrt0DugEpsg4WiZ
# 5QihobWA1FsFmtmfcea1ueIK/tKETo6WLWpAkwgp7Kbh+yjnE+P3lmIQ9VtmpXlL
# yfWrr0R1+d7EckG9Q4QWc0De7WzKXkm0KbHcpDmftrw/4GStdMSZprQbaX75lhYK
# bPNEmXDm646ADnQiSXA9LScE9l2TJht2G7lCVt9/XUSdk6wy1AG4aeHu0IWhDYTE
# jWii/zAy7ZeCc041apqkKDiMr6u/dDH0kNFs3F/xrRj0Z2xOlhYh7oB4dMkYNyX2
# Av16ZZpPoMhfiPOcufDZLrns4cBgZ9ynWhqFdFsfU4KyIic7/fevw4DX3Bo3Pi14
# Kr6FhEZR3EfZYJq6dfy9OGN2UKJq/6QjbWr3q/63lKXP91g2YLOOn1EoLSgdvqvP
# eWdrarK2nmiU9gZwBf/g9i1VKO5kFDZSaEKrM/u1e8R+Z2BcLq6LR42e5rOuOsUt
# pIXl7xTe3LJr6H4USahe8KWPd9WQ5u2f1eKviDmEIcNZHfHeUjm0HCiKPBCWOK2I
# 7nX7aH7yzeFRiKuRz8+x0zxcT3fNJR3EVrKb7vQ5cvC6YGqc6oclKUIjmdORSadd
# XSYnM6ki4oNyBiv+So3SKlrYjqz6ZDzRgbrSOQvJ73gv8FZff000UJkdBJHPIsZm
# tjoNtKlukA==
# SIG # End signature block
