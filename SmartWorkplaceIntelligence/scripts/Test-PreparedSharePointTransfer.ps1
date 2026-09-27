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
Reject {Send-PreparedEvidenceBatch @params} 'pointer identity'
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAwTZ7xWhOS7qhp
# M/SBQN1+o/gRuumG8yrEjJz/MBC5LKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOm2w9+/AMgLPjiyu8olL9VTHcciJ5er+cggyYXwbiiPMA0GCSqG
# SIb3DQEBAQUABIIBgHrpfwig42l0Szt2OzjTeVc2RUhCnMlXMdMkSIv4CZUBDfBB
# AP5ml8zlK6uVzRJ4oISw8Bwnqo/m7FgNzVXtqURtG7aC8KsAzPn++D958HIoNf1Z
# Ee7AzmPTMTG2cuhmMpkNVj1a07CwGMCHHDASmzEFgUJ0eW6hsOkMizlWVHpIYCig
# aFb2+j1Pjg/pfaS/zm6Csde2DlZ+p/NvuyADcKaJoXbHuKev/YZJAxllg5a/Ic8/
# 8i8wzZQn1UUbcnXJBsV/4Tk41fAavkG/4Y8PeJIlO4y29bhDMf3yNM31zux6HVYK
# Z//qAsUyNXKBc3VmcayeM2//ijH56cA+Tg3CwAe/sK0UQwMjsxcyhChElrvW252B
# BmbPFkSC2E2Mz9UI/jb0OhnobOV/hbUVmIjfRFkWdr9kBJO8DBsOOC1mkJ+dPfFY
# +7DxvC5YBiE/FwCQ4Aaw9eDqKyUp2iBCM9xzmccRuna+L8ETm7y7BlByxe3JtSiv
# q5AUnVgRlLwrlAeTHaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxMzU0
# MTBaMC8GCSqGSIb3DQEJBDEiBCDlIoWlaAjaB4mYpugQIH2j+F36sO8/Hl9wSEAH
# ApGiNDANBgkqhkiG9w0BAQEFAASCAgAFBZdAJJlh8RabtxavB7O8/iZEq+76m+Yy
# hc4iJ0oTRe7SqCRQZ6FEYc7T0YkEylT+1mSw5En1M+Fj0E45HYMc7W7QKtn/HMhV
# R0pzMp7yzboKuSYho8fJwz88F4rCJv64TMpxbiZjEY+F+b0y3xMJG6jevl+cbMW9
# kvLvc/0DJCuu5DM93eaDYHjl0o9yQrty/iDDKH2aKDR/EiWylT8inNW2XtMBl7nK
# hWrlwuTYp82jr4gKablMgYFe5MCwHkJueCDMPQ7oulNR2d61aisXeM7RRFUuTtd4
# wFgTY+X5LOee/+oC5KfSIvrQP93At8LxY7wchJxtZrcynm/NwZeIans+o+S/TGHk
# GzpBHnameD5XUWiSF15cAETGWpownLgdakFxL1qW1ucdjzhY28ts39t9in7v1h1h
# tlywIa8oh/aYehBqozVAZFeXvV1XupyAuNW+HJpJxIlCi8qp71gKYN8WvhRkCK6p
# 4IHDNsXTqIpdDFPALdlusFhGL/FC6aDQELQAxHZHZ0OF47cAzNcrUIS9cxLLrU/k
# nw+eI/8IDZfzmcRakyqCRzaAglry2bZyc9WKMJKLgaVn+AGpUmdEUeuzlWUv3ton
# Ii2/X3dtxGzC8wRvBNnhXMTwxArwJ7235GN/7/m+ZztmFb+zOx1dIDxMfiWdwxC3
# 7uDLmo5Ftg==
# SIG # End signature block
