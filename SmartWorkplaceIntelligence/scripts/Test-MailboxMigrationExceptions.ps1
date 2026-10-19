# Version 0.1.0. Synthetic mapping rollback, contract and transfer guards only.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot, [string]$PythonCommand='python')
$ErrorActionPreference='Stop'
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ('mailbox-exceptions-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PreparedSharePointTransfer.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'NativeEvidence.psm1') -Force
$script:checks=0
function Check([bool]$Pass,[string]$Message){if(-not $Pass){throw $Message};$script:checks++}
function Reject([scriptblock]$Action){$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Check $failed 'Expected fail-closed rejection'}
$fixtures=Join-Path $root 'fixtures';$data=Join-Path $root 'data'
New-Item -ItemType Directory -Path $fixtures,$data | Out-Null
function MinimalWorkbook([string]$Path,[string[]]$Sheets){
    $archive=[IO.Compression.ZipFile]::Open($Path,[IO.Compression.ZipArchiveMode]::Create)
    try{$writer=[IO.StreamWriter]::new($archive.CreateEntry('xl/workbook.xml').Open());try{$writer.Write('<workbook><sheets>'+ (($Sheets|ForEach-Object {'<sheet name="'+$_+'"/>'}) -join '')+'</sheets></workbook>')}finally{$writer.Dispose()}}finally{$archive.Dispose()}
}
$persona='SmartWorkplaceIntelligence-PersonaClassification.xlsx';$site='SmartWorkplaceIntelligence-SiteClassification.xlsx';$exceptions='SmartWorkplaceIntelligence-MailboxMigrationExceptions.xlsx'
MinimalWorkbook (Join-Path $fixtures $persona) @('Personas','Rules','Exclusions')
MinimalWorkbook (Join-Path $fixtures $site) @('Settings','SiteTypes','Sites')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "../config/templates/$exceptions") -Destination (Join-Path $fixtures $exceptions)
$download={param($destination,$name) Copy-Item -LiteralPath (Join-Path $fixtures $name) -Destination $destination;Get-Item -LiteralPath $destination}.GetNewClosure()
$accepted=Receive-PreparedMappingWorkbooks -DataRoot $data -DownloadFile $download -MailboxMigrationExceptionsEnabled -PythonCommand $PythonCommand
Check (@(Get-ChildItem -LiteralPath $accepted -File).Count -eq 3) 'Enabled mapping set contains all three fresh workbooks'
$hashes=@{};foreach($name in @($persona,$site,$exceptions)){$hashes[$name]=(Get-FileHash -LiteralPath (Join-Path $data $name)).Hash}
$failThird={param($destination,$name) if($name -eq $exceptions){return $null}; & $download $destination $name}.GetNewClosure()
Reject {Receive-PreparedMappingWorkbooks -DataRoot $data -DownloadFile $failThird -MailboxMigrationExceptionsEnabled -PythonCommand $PythonCommand}
foreach($name in $hashes.Keys){Check ((Get-FileHash -LiteralPath (Join-Path $data $name)).Hash -eq $hashes[$name]) ('Third download failure preserves '+$name)}
$invalidThird={param($destination,$name) if($name -eq $exceptions){Copy-Item -LiteralPath (Join-Path $fixtures $site) -Destination $destination;return Get-Item $destination}; & $download $destination $name}.GetNewClosure()
Reject {Receive-PreparedMappingWorkbooks -DataRoot $data -DownloadFile $invalidThird -MailboxMigrationExceptionsEnabled -PythonCommand $PythonCommand}
foreach($name in $hashes.Keys){Check ((Get-FileHash -LiteralPath (Join-Path $data $name)).Hash -eq $hashes[$name]) ('Invalid third workbook preserves '+$name)}
$sourceContract=Join-Path $root 'source-contract.json.txt'
@{currentFiles=@();mappingFiles=@($persona,$site);dailyFiles=@();history=@();mailboxMigrationExceptionsFile=$exceptions}|ConvertTo-Json|Set-Content -LiteralPath $sourceContract
$disabled=@(Get-PreparedSourcePlan -DataRoot $data -SourceContractPath $sourceContract -MappingRoot $accepted)
$enabled=@(Get-PreparedSourcePlan -DataRoot $data -SourceContractPath $sourceContract -MappingRoot $accepted -MailboxMigrationExceptionsEnabled)
Check ($disabled.Count -eq 2 -and $enabled.Count -eq 3) 'Feature flag controls capture, with no cached third-file fallback when disabled'
Check (@($enabled|Where-Object Relative -EQ $exceptions).SHA256 -eq $hashes[$exceptions]) 'Workbook hash is frozen in capture provenance'
Remove-PreparedMappingWorkbooks -DataRoot $data -MappingRoot $accepted
foreach($name in $hashes.Keys){Check (Test-Path -LiteralPath (Join-Path $data $name)) ('Cleanup retains operational mapping '+$name)}
# Transfer old/new immutable cohorts; no filesystem transport or production data.
$product=Split-Path $PSScriptRoot -Parent
$base=Get-Content -LiteralPath (Join-Path $product 'config/prepared-evidence-contract.json.txt') -Raw|ConvertFrom-Json
$base=Merge-NativeEvidenceContract $base
$extension=Get-Content -LiteralPath (Join-Path $product 'config/mailbox-migration-exceptions-contract.json.txt') -Raw|ConvertFrom-Json
$intunePath=Join-Path $product 'config/intune-operational-contract.json.txt'
$nativePath=Join-Path $product 'config/native-output-contract.json.txt'
$exceptionPath=Join-Path $product 'config/mailbox-migration-exceptions-contract.json.txt'
$transferModule=Get-Module PreparedSharePointTransfer
function Manifest($Contract,[bool]$Enabled){
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($Contract|ConvertTo-Json -Depth 20))
    $hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    $proof=@{NativeEvidence=@{ContractSHA256=(Get-FileHash $nativePath).Hash};IntuneOperations=@{ContractSHA256=(Get-FileHash $intunePath).Hash}}
    if($Enabled){$proof.MailboxMigrationExceptions=@{ContractSHA256=(Get-FileHash $exceptionPath).Hash}}
    @{Files=@($Contract.tables|ForEach-Object {@{File=$_.file}});Provenance=$proof;ContractSHA256=$hash}|ConvertTo-Json -Depth 20|ConvertFrom-Json
}
$path=Join-Path $root 'base.json.txt';[IO.File]::WriteAllText($path,($base|ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))
$old=Manifest $base $false
$resolved=& $transferModule {param($m,$p) Resolve-PreparedTransferContract $m $p} $old $path
Check ($resolved.tables.Count -eq $base.tables.Count) 'Existing immutable batch contract remains unchanged'
$expanded=$base|ConvertTo-Json -Depth 30|ConvertFrom-Json;$expanded.tables=@($expanded.tables)+@($extension.tables)
$new=Manifest $expanded $true
$resolved=& $transferModule {param($m,$p) Resolve-PreparedTransferContract $m $p} $new $path
Check ($resolved.tables.Count -eq $expanded.tables.Count) 'New exception tables participate in the exact unified contract hash'
$partial=Manifest $expanded $true;$partial.Files=@($partial.Files|Where-Object File -NE 'MailboxMigrationExceptions.csv')
Reject {& $transferModule {param($m,$p) Resolve-PreparedTransferContract $m $p} $partial $path}
$wrong=Manifest $expanded $true;$wrong.Provenance.MailboxMigrationExceptions.ContractSHA256=('0'*64)
Reject {& $transferModule {param($m,$p) Resolve-PreparedTransferContract $m $p} $wrong $path}
# Exercise the production config loader on a private synthetic file, never initialize a tenant.
. (Join-Path $PSScriptRoot '../../SmartM365/Config/SmartM365-TenantContext.ps1')
$localConfig=Join-Path $root 'fixture.local.json'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot '../../SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.local.json.template') -Destination ($localConfig+'.template')
@{PreparedPythonCommand='keep-python';Nested=@{Keep='private-value'}}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $localConfig
$loaded=Read-SmartM365JsonConfig -Path $localConfig -Required
Check ($loaded.PreparedMailboxMigrationExceptionsEnabled -eq $false -and $loaded.PreparedMailboxMigrationBindingsPath -eq '') 'New default keys are added by the real runtime loader'
Check ($loaded.PreparedPythonCommand -eq 'keep-python' -and $loaded.Nested.Keep -eq 'private-value') 'Runtime template merge preserves existing local and nested values'
$saved=(Read-SmartM365JsonDocument -Path $localConfig).Document
Check ($saved.PSObject.Properties['PreparedMailboxMigrationExceptionsEnabled'] -and $saved.PreparedPythonCommand -eq 'keep-python') 'Merged configuration is saved without overwriting local values'
# Execute the real Data Trust adapter against a synthetic audit and summary.
$auditPath=Join-Path $root 'MailboxMigrationExceptionsAudit.json.txt'
@{'Source Rows'=0;'Snapshot Date'='2026-10-19';'Workbook SHA256'=('0'*64);'Captured UTC'='2026-10-19T08:00:00Z'}|ConvertTo-Json|Set-Content -LiteralPath $auditPath
[pscustomobject]@{'Domain Key'='__all__';'Review Count'=2}|Export-Csv -LiteralPath (Join-Path $root 'MailboxMigrationExceptionSummary.csv') -NoTypeInformation
$trustOutput=Join-Path $root 'trust.csv'
& (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'New-DataTrustEvidence.ps1') -DataRoot $data -OutputPath $trustOutput -AsOfDate '2026-10-19' -MailboxMigrationExceptionsAuditPath $auditPath | Out-Null
Check ($LASTEXITCODE -eq 0) 'Data Trust adapter executes successfully'
$trust=@(Import-Csv -LiteralPath $trustOutput|Where-Object 'Source Name' -EQ 'Mailbox migration exception register')
Check ($trust.Count -eq 1 -and $trust[0].'Decision Readiness' -eq 'Review required' -and $trust[0].'Evidence Age Days' -eq '0') 'Data Trust exposes one dated, review-qualified registry row'
[pscustomobject]@{Status='Passed';Checks=$script:checks;TestRoot=$root}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDf6JXDQOGjXpko
# QzGbMFRCHwFOd7fgYdB4RiQAQaOqbaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMY3tnA4lT8ScOYetDx1Yye9waV8T62KJN1k91ziqGgRMA0GCSqG
# SIb3DQEBAQUABIIBgD/EnQ61NAN7GqOa6iRwioc0Htx+ctDbfnZ2rMjnZD7KpO3r
# 44es+jR9vR9d8LP3mvU7hlwZ9M1Q59zQG7pkyAAJGE8bVPLpRHQeZfiLbXTckWeJ
# AMfqhGDeOcFYxXjiWYWksRx5iCdJvx0fEmB1eQ1aot5jQaKwzb76tyQ4Ph5i8hJf
# E13YELuEZ0xTtbOAAydN0mDuCTFFBtE0u2fNrNdZPRLo09m21pByYa4GyallwSnG
# 5tJlhx7/5hogTEdVi/gX95ckzfamDhX2a4Rb27uaA+1EU+IKvBo0yT4UyLRZknjm
# QqR9YhyWEScZilqpvAz8tt4ef5LGcczkv2KsPxT6mRRJjLzIdlUhTYBuX/sJIFOK
# i+HSoJ/7HuRTjZLaldAnHaWYru55N4lXjfEr4zleRCiR8S13E11/gRR8mDLT1Odx
# MxB646OfboZj956DLhkD+sKS8sNHbYm8COE/eCaDL3EG+sGSQNNEckMC4Uc2oLls
# sDon+29nDAKF6mcNU6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxNTAz
# MzRaMC8GCSqGSIb3DQEJBDEiBCA+iYMIY+WIrOVX3/WCo9/8B91VaJAGuxXTqLAs
# VKef6DANBgkqhkiG9w0BAQEFAASCAgBtkkBzb+9NjQcSMoATVC9l1gobgQX+wsnd
# 9XZvzhA5Wob032j0Pu6XPVwWzALfiwyxAlAnDgUzdwvN2fnnXe3umpfV9LXGx7H/
# LWr6saZE4O1AiGyCON8boSHXxW3Kc0+YByJyQrH1nGKZKkDjR9aTIU7966ZdNx4O
# Ifyzts++mfGADPLOqi1KmWibnCTTtLJu6BTN/QJ1ZtQgh9VtR16qDxj/fPTrS2WS
# +dTxp2SyLTFhFSsAOyLYmBR4ExmRFNntcyG7Ex/jKc6xQy5UhJ24F2ii5uwUbLw1
# 3z2dSOQJvtLisuR07BbVT249O7Wdcd3dHoyPGcwgeuMKt0HXum7zURb1VG0eEcC0
# GMzkKMqIWVqurUp7+0ADFJlDICsk7o82IhcfD/LNDZS/PQjCt3NOfgOytIckdNcn
# bXWPtTfyPM9LC235KwNTAqwPKVmPVeNPsImuUziG4gH8RzPlLbCd8OaUZVAoyrnP
# kGMLcHSJizNU5bZDM0Sst0UioudwZqIq/lDNLbN2tdl7pdSV8Qft078w3gdbNP84
# 8Y++CjbsBZmXx9hoy+fiUTSk3J+b5rJD0F1pHkCwm83A37+qJHsPYWSKd8janyxi
# OjRh6lrh2gvwHBJy6Xoz8b2DZNSzLUY4nHy5ptacPKzy/3J9qYQgoitw2lLub3Q1
# v6D2R/csiw==
# SIG # End signature block
