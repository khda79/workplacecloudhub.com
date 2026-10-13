[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PreparedSharePointTransfer.psm1') -Force
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$staging=Join-Path $root 'staging';$output=Join-Path $root 'DATA/DATA-POWERBI';$scratch=Join-Path $root 'scratch'
New-Item -ItemType Directory -Path $staging -Force | Out-Null
$contract=Join-Path $root 'contract.json'
@{tables=@(@{table='Example';file='Example.csv';uniqueKey=@('Name');columns=@(@{name='Name';type='string'})})} | ConvertTo-Json -Depth 8 | Set-Content $contract
[pscustomobject]@{Name='Synthetic'} | Export-Csv (Join-Path $staging 'Example.csv') -NoTypeInformation
$batch=Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $output -TenantKey 'synthetic' -Provenance @{Test=$true} -ContractPath $contract
$script:checks=0
function Check($value,[string]$message){if(-not $value){throw $message};$script:checks++}
function Reject([scriptblock]$action,[string]$pattern){$caught=$null;try{& $action | Out-Null}catch{$caught=$_};Check ($null -ne $caught) 'Expected rejection';Check ($caught.Exception.Message -match $pattern) "Unexpected rejection: $caught"}
function Transport([string]$mode='ok'){
    $localCsvPath=Join-Path $batch.BatchPath 'Example.csv'
    $remote=Join-Path $root ('cloud/'+[guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $remote -Force | Out-Null
    $state=@{Remote=$remote;Mode=$mode;Events=[Collections.Generic.List[string]]::new()}
    $upload={param($local,$relative)
        $state.Events.Add('U:'+$relative)
        if($state.Mode -eq 'upload-fail' -and $relative.EndsWith('.csv')){return $null}
        $path=Join-Path $state.Remote $relative;New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $local -Destination $path
        return $true
    }.GetNewClosure()
    $download={param($destination,$relative)
        $state.Events.Add('D:'+$relative)
        if($state.Mode -eq 'read-fail' -or ($state.Mode -eq 'pointer-read-fail' -and $relative -eq 'current.json.txt')){return $null}
        Copy-Item -LiteralPath (Join-Path $state.Remote $relative) -Destination $destination
        if($state.Mode -eq 'corrupt'){Add-Content -LiteralPath $destination 'bad'}
        if($state.Mode -eq 'mutate-local' -and $relative.EndsWith('.csv')){Add-Content -LiteralPath $localCsvPath 'changed'}
        return $true
    }.GetNewClosure()
    @{State=$state;Params=@{OutputRoot=$output;TenantKey='synthetic';ContractPath=$contract;WorkRoot=$scratch;ExpectedBatchId=$batch.BatchId;UploadFile=$upload;DownloadFile=$download}}
}
Check ((Resolve-PreparedSharePointFolder '' 'Group/DATA') -ceq 'Group/DATA/DATA-POWERBI') 'Default cloud path'
Check ((Resolve-PreparedSharePointFolder 'Other\DATA-POWERBI' 'Group/DATA') -ceq 'Other/DATA-POWERBI') 'Explicit path override'
Reject {Resolve-PreparedSharePointFolder '' ''} 'missing'
foreach($unsafe in 'https://example/DATA-POWERBI','DATA/../DATA-POWERBI','DATA/DATA-LAST/DATA-POWERBI','DATA//DATA-POWERBI') {Reject {Resolve-PreparedSharePointFolder $unsafe ''} 'library-relative'}
$t=Transport;$before=(Get-FileHash (Join-Path $batch.BatchPath 'Example.csv')).Hash
$params=$t.Params;$result=Send-PreparedEvidenceBatch @params
Check ($result.PointerVerified -and $result.VerifiedFiles -eq 5 -and $result.CsvFiles -eq 1 -and -not $result.CsvRecalculated) 'Success counts'
Check ($t.State.Events[-2] -eq 'U:current.json.txt' -and $t.State.Events[-1] -eq 'D:current.json.txt') 'Pointer was not last'
Check ((Get-FileHash (Join-Path $t.State.Remote 'current.json.txt')).Hash -eq (Get-FileHash (Join-Path $output 'current.json.txt')).Hash) 'Pointer bytes changed'
Check ((Get-FileHash (Join-Path $batch.BatchPath 'Example.csv')).Hash -eq $before) 'Transfer changed local CSV'
Check ((Get-Content $result.AuditPath -Raw | ConvertFrom-Json).Status -eq 'Success') 'Success audit missing'
Check (@(Get-ChildItem $scratch -Recurse -Filter readback.tmp).Count -eq 0) 'Temporary payload retained'
$params=$t.Params; $again=Send-PreparedEvidenceBatch @params
Check ($again.BatchId -eq $result.BatchId -and $again.PointerVerified) 'Retry changed batch'
foreach($mode in 'upload-fail','read-fail','corrupt'){
    $t=Transport $mode;$params=$t.Params
    Reject {Send-PreparedEvidenceBatch @params} '(Upload failed|Read-back failed|Remote SHA256)'
    Check (-not(Test-Path (Join-Path $t.State.Remote 'current.json.txt'))) 'Failure advanced pointer'
}
$t=Transport 'pointer-read-fail';$params=$t.Params
Reject {Send-PreparedEvidenceBatch @params} 'publication was attempted but completion is not confirmed'
Check (Test-Path (Join-Path $t.State.Remote 'current.json.txt')) 'Pointer failure fixture did not exercise post-upload uncertainty'
$t=Transport;$params=$t.Params;$params.TenantKey='other'
Reject {Send-PreparedEvidenceBatch @params} 'pointer identity|Audit tenant ownership mismatch'
Check ($t.State.Events.Count -eq 0) 'Cross-tenant transfer made a network call'
$params.TenantKey='synthetic';$params.ExpectedBatchId='20260101T000000000Z-aaaaaaaa'
Reject {Send-PreparedEvidenceBatch @params} 'ExpectedBatchId'
$params.ExpectedBatchId=$batch.BatchId
$lock=[IO.File]::Open((Join-Path $output '.publication.lock'),'OpenOrCreate','ReadWrite','None')
try{Reject {Send-PreparedEvidenceBatch @params} 'being used|utilis|Open'}finally{$lock.Dispose()}
$localCsv=Join-Path $batch.BatchPath 'Example.csv';$original=[IO.File]::ReadAllBytes($localCsv)
try{
    $t=Transport 'mutate-local';$params=$t.Params
    Reject {Send-PreparedEvidenceBatch @params} 'Local batch changed'
    Check (-not(Test-Path (Join-Path $t.State.Remote 'current.json.txt'))) 'Source mutation advanced pointer'
}finally{[IO.File]::WriteAllBytes($localCsv,$original)}
$t=Transport;$params=$t.Params
try{Add-Content $localCsv 'bad';Reject {Send-PreparedEvidenceBatch @params} 'Local CSV integrity';Check ($t.State.Events.Count -eq 0) 'Local corrupt CSV uploaded'}finally{[IO.File]::WriteAllBytes($localCsv,$original)}
$params.WorkRoot=Join-Path $output 'scratch'
Reject {Send-PreparedEvidenceBatch @params} 'scratch must be outside'
Check (@(Get-ChildItem $scratch -Recurse -Filter readback.tmp).Count -eq 0) 'Failed transfer retained temporary payload'
# Exercise the actual entry-point adapters against an in-memory Core stub.
$repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$entryPath=Join-Path $repo 'SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1'
$tokens=$null;$errors=$null;$entryAst=[Management.Automation.Language.Parser]::ParseFile($entryPath,[ref]$tokens,[ref]$errors)
Check ($errors.Count -eq 0) 'Entry point parse errors'
$assignment=$entryAst.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$transferParameters'},$true)
$corePath=Join-Path $repo 'SmartM365/Modules/SmartM365.Core/SmartM365.Core.psm1'
$coreAst=[Management.Automation.Language.Parser]::ParseFile($corePath,[ref]$tokens,[ref]$errors)
$normalizer=$coreAst.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertTo-SmartM365SharePointDataRootPath'},$true)
. ([scriptblock]::Create($normalizer.Extent.Text))
Check ((Resolve-PreparedSharePointFolder '' (ConvertTo-SmartM365SharePointDataRootPath 'Example/CSV')) -eq 'Example/DATA/DATA-POWERBI') 'Shared CSV-to-DATA convention changed'
$global:PreparedTransferAdapterEvents=[Collections.Generic.List[object]]::new()
New-Module -Name SmartM365.Core -ScriptBlock {
    function Get-SmartM365SharePointRelativeFilePath {param($LocalFilePath) [IO.Path]::GetFileName($LocalFilePath)}
    function Invoke-SmartM365SharePointCsvUpload {param($LocalFilePath,$TargetFolderPath,[switch]$EnsureParentFolders,$Enabled,$SiteHostname,$SitePath,$LibraryDisplayName,$AppId,$TenantId,$Thumbprint)
        $global:PreparedTransferAdapterEvents.Add(@{Operation='Upload';Target=$TargetFolderPath;Ensure=[bool]$EnsureParentFolders;Enabled=$Enabled});$true}
    function Invoke-SmartM365SharePointFileDownload {param($LocalFilePath,$SharePointRelativePath,$TargetFolderPath,[switch]$Force,$Enabled,$SiteHostname,$SitePath,$LibraryDisplayName,$AppId,$TenantId,$Thumbprint)
        $global:PreparedTransferAdapterEvents.Add(@{Operation='Download';Target=$TargetFolderPath;Relative=$SharePointRelativePath;Force=[bool]$Force});$true}
    Export-ModuleMember -Function *
} | Import-Module -Force
try{
    $cloudRoot='Example/DATA/DATA-POWERBI';$cloud=@{Enabled=$true;SiteHostname='synthetic';SitePath='/sites/test';LibraryDisplayName='Documents';AppId='test';TenantId='test';Thumbprint='test'}
    $effective=@{TenantKey='synthetic'};$work=$scratch;$product=Split-Path $PSScriptRoot -Parent
    . ([scriptblock]::Create($assignment.Extent.Text))
    $null=& $transferParameters.UploadFile (Join-Path $batch.BatchPath 'Example.csv') ('batches/'+$batch.BatchId+'/Example.csv')
    $null=& $transferParameters.UploadFile (Join-Path $batch.BatchPath 'current.json.txt') 'current.json.txt'
    $null=& $transferParameters.DownloadFile (Join-Path $scratch 'read.tmp') 'current.json.txt'
    Check ($global:PreparedTransferAdapterEvents[0].Target -eq ($cloudRoot+'/batches/'+$batch.BatchId)) 'Batch upload adapter target'
    Check ($global:PreparedTransferAdapterEvents[0].Ensure -and $global:PreparedTransferAdapterEvents[0].Enabled) 'Upload folder provisioning missing'
    Check ($global:PreparedTransferAdapterEvents[1].Target -eq $cloudRoot) 'Pointer upload adapter target'
    Check ($global:PreparedTransferAdapterEvents[2].Target -eq $cloudRoot -and $global:PreparedTransferAdapterEvents[2].Force) 'Read-back adapter target/cache behavior'
}finally{Remove-Module SmartM365.Core;Remove-Variable PreparedTransferAdapterEvents -Scope Global}
$entry=[IO.File]::ReadAllText($entryPath)
$transferBranch=$entryAst.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith('if ($TransferOnly)')},$true)
Check ($transferBranch.Extent.EndOffset -lt $entry.IndexOf("`$phase='Download SharePoint classification workbooks'") -and $transferBranch.Extent.Text -match '\breturn\b') 'Transfer-only falls through to mapping download/preparation'
Check ($entry -match 'if \(\$TransferOnly -and \(\$Offline') 'Offline/transfer mode exclusion missing'
Write-Host "PASS: $script:checks synthetic transfer, integrity, pointer ordering, retry, locking and failure checks. No tenant or network calls."

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC4GmOXhnuB6UTP
# 9xE9UsN+11SQuM6t8NWDVx0YVSKwLqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCY9AEfMRynnNeyRnx9bOGQ
# JnxCZm6mFS8/GwIeQbBCSzANBgkqhkiG9w0BAQEFAASCAYCCsFQv8uU32qWAGkvX
# 6e1PnNsF482TKa745LDnty+alBlHGNopGNi9e2DQej/tSFEEdm8zr7Zds4zYuutz
# gZ1aGdJxaLY02Mbz0Bzvhw/RrhjT56btBXgbH3iF+djgNWX4JUOcZ8boEhlxdGCd
# R1tGdJsHxTda/mKnfIYRm5hh0xyizzIDOWpQGGOhoNy2Yq0lQpOCsAHmuTzX4u3Z
# V6IJbWsQQGOq7VdI0a7TKt1Qb0v3pH+9x24LuZsx3DpmvEQwexC4gjRzAASHevSr
# bDN/dcpAKvFHaR3daDw9ah2cdNvPa2IuXgUMA96Pn3lZNat1E/fkiQ4GAl6gYhyB
# gf+21bdXJlL4CbMAaLzNdW+6jTqaLgKZdU7aPlbcbVOezYCIf59YSEcoKsXCok3j
# KUlPa5wohmZeKM6/jsweMrXbSlsdwCwrb2ICKWkGqM828+UPAr/VKCbuLRZjIjyx
# de+ZjalSS7aVyJb4/JCGu1CRcONYYQwEbip30/YdGkJ/FW0=
# SIG # End signature block
