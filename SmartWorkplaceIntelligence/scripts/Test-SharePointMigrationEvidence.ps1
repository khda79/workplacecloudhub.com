[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot,[string]$PythonCommand='python')
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'SharePointMigrationEvidence.psm1') -Force
$taskRoot=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $taskRoot -Force | Out-Null
$python=Get-Command $PythonCommand -CommandType Application -ErrorAction Stop | Select-Object -First 1
& $python.Source -B (Join-Path $PSScriptRoot 'spo_migration_fixture.py') --root $taskRoot
if($LASTEXITCODE -ne 0){throw 'Synthetic fixture creation failed.'}
$raw=Join-Path $taskRoot 'raw'
$staged=Join-Path $taskRoot 'prepared'
$fixture=Get-Content -LiteralPath (Join-Path $taskRoot 'fixture.json.txt') -Raw | ConvertFrom-Json
$sourceItems=@(Get-SharePointMigrationSourceItems -DataRoot $raw)
if($sourceItems.Count -ne 12 -or 'DATA-LAST/SharePoint_OnPrem_Webs.csv' -notin $sourceItems.Relative){throw 'Migration receipt groups were not added to the source capture.'}
$args=@{DataRoot=$raw;OutputRoot=$staged;TenantKey='sample-prod';Enabled=$true;AsOfUtc=[datetimeoffset]'2026-10-19T12:00:00Z';PythonCommand=$PythonCommand;TargetTenantUrl='https://sample.sharepoint.com'}
& (Join-Path $PSScriptRoot 'New-SharePointMigrationEvidence.ps1') @args
& (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root $staged -TenantKey 'sample-prod' -ContractPath (Join-Path $PSScriptRoot '../config/sharepoint-migration-contract.json.txt') -ReportPath (Join-Path $taskRoot 'validation.json.txt') | Out-Null
if($LASTEXITCODE -ne 0){throw 'Synthetic prepared CSV validation failed.'}
$plans=@(Import-Csv -LiteralPath (Join-Path $staged 'SharePointMigrationPlanEvidence.csv'))
if($plans.Count -ne 2 -or @($plans.'Operation Key' | Sort-Object -Unique).Count -ne 2){throw 'Cross-flow operation identity failed.'}
if(@($plans | Where-Object 'Target Observation' -EQ 'Site observed').Count -ne 2){throw 'SPO target observation was not qualified.'}
if(@($plans | Where-Object 'Is Validated' -EQ 'True').Count){throw 'Declared DONE was promoted to migration validation.'}
if((Get-FileHash -LiteralPath (Join-Path $raw 'SmartWorkplaceIntelligence-Migration-to-SPO.xlsx')).Hash -ne $fixture.WorkbookSHA256){throw 'Source workbook was changed.'}

# Missing configured sources and tampered receipts fail; there is no cached fallback.
$receipt=Join-Path $raw 'DATA-LAST/SmartInventory_SmartM365-SharePoint-OnPrem-Content-Inventory.current.json.txt'
$saved=[IO.File]::ReadAllBytes($receipt)
try {
    [IO.File]::WriteAllText($receipt,'{}')
    $failed=$false
    try { & (Join-Path $PSScriptRoot 'New-SharePointMigrationEvidence.ps1') @args 2>$null } catch {$failed=$true}
    if(-not $failed){throw 'Invalid receipt was accepted.'}
} finally {[IO.File]::WriteAllBytes($receipt,$saved)}

$empty=Join-Path $taskRoot 'disabled'
& (Join-Path $PSScriptRoot 'New-SharePointMigrationEvidence.ps1') -DataRoot $raw -OutputRoot $empty -TenantKey 'sample-prod' -PythonCommand $PythonCommand
if((Import-Csv -LiteralPath (Join-Path $empty 'SharePointMigrationSourceSummary.csv')).Qualification -ne 'Not configured'){throw 'Disabled feature produced a healthy zero.'}
$unsafeFailed=$false
try{Assert-SharePointMigrationFileName '../private.xlsx'}catch{$unsafeFailed=$true}
if(-not $unsafeFailed){throw 'Unsafe configured name accepted.'}
$templateSheets=@{'SmartWorkplaceIntelligence-PersonaClassification.xlsx'=@('Personas','Rules','Exclusions');'SmartWorkplaceIntelligence-SiteClassification.xlsx'=@('Settings','SiteTypes','Sites')}
foreach($name in $templateSheets.Keys){
    $zip=[IO.Compression.ZipFile]::Open((Join-Path $raw $name),[IO.Compression.ZipArchiveMode]::Create)
    try{
        $entry=$zip.CreateEntry('xl/workbook.xml')
        $writer=[IO.StreamWriter]::new($entry.Open())
        try{$writer.Write('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets>'+ (($templateSheets[$name] | ForEach-Object {'<sheet name="'+$_+'"/>'}) -join '') +'</sheets></workbook>')}finally{$writer.Dispose()}
    }finally{$zip.Dispose()}
}
$bindingName='MigrationBindings.local.json.txt'
@{SchemaVersion=1;Identity=@{TenantKey='sample-prod';ProjectKey='sample-project'};Bindings=@()} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $raw $bindingName)
$mappingData=Join-Path $taskRoot 'mapping-data'
New-Item -ItemType Directory -Path $mappingData | Out-Null
$download={param($destination,$name) Copy-Item -LiteralPath (Join-Path $raw $name) -Destination $destination; Get-Item -LiteralPath $destination}.GetNewClosure()
$mapping=Receive-PreparedMappingWorkbooks -DataRoot $mappingData -DownloadFile $download -SharePointMigrationEnabled -SharePointMigrationBindingsFileName $bindingName -PythonCommand $PythonCommand
if(@(Get-ChildItem -LiteralPath $mapping -File).Count -ne 4){throw 'Migration workbook and registry did not join the mapping capture.'}
$sourceContract=Join-Path $taskRoot 'minimal-source-contract.json.txt'
@{currentFiles=@();mappingFiles=@('SmartWorkplaceIntelligence-PersonaClassification.xlsx','SmartWorkplaceIntelligence-SiteClassification.xlsx');dailyFiles=@();history=@();sharePointMigrationContract='sharepoint-migration-contract.json.txt'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $sourceContract
$sourcePlan=@(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $sourceContract -MappingRoot $mapping -MetadataOnly -SharePointMigrationEnabled -SharePointMigrationBindingsFileName $bindingName)
if($sourcePlan.Count -ne 16 -or 'DATA-LAST/SharePoint_OnPrem_Webs.csv' -notin $sourcePlan.Relative){throw 'Full pipeline source plan omitted migration inventory inputs.'}
$before=(Get-FileHash -LiteralPath (Join-Path $mappingData 'SmartWorkplaceIntelligence-Migration-to-SPO.xlsx')).Hash
$broken={param($destination,$name) if($name -ceq $bindingName){return $null}; Copy-Item -LiteralPath (Join-Path $raw $name) -Destination $destination; Get-Item -LiteralPath $destination}.GetNewClosure()
$failed=$false
try{$null=Receive-PreparedMappingWorkbooks -DataRoot $mappingData -DownloadFile $broken -SharePointMigrationEnabled -SharePointMigrationBindingsFileName $bindingName -PythonCommand $PythonCommand}catch{$failed=$true}
if(-not $failed -or (Get-FileHash -LiteralPath (Join-Path $mappingData 'SmartWorkplaceIntelligence-Migration-to-SPO.xlsx')).Hash -ne $before){throw 'Missing registry did not preserve the validated mapping set.'}

$published=Join-Path $taskRoot 'predecessor'
$batch=Join-Path $published 'batches/test-batch'
New-Item -ItemType Directory -Path $batch -Force | Out-Null
$historyName='SharePointMigrationHistoryEvidence.csv'
Copy-Item -LiteralPath (Join-Path $staged $historyName) -Destination (Join-Path $batch $historyName)
$hash=(Get-FileHash -LiteralPath (Join-Path $batch $historyName)).Hash
$entry=@{File=$historyName;Rows=2;Bytes=(Get-Item -LiteralPath (Join-Path $batch $historyName)).Length;SHA256=$hash}
@{BatchId='test-batch';TenantKey='sample-prod';Files=@($entry)} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $batch 'batch.json.txt')
@{Passed=$true;SchemaOnly=$false;Files=@($entry)} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $batch 'validation.json.txt')
@{BatchId='test-batch';TenantKey='sample-prod';ManifestSHA256=(Get-FileHash -LiteralPath (Join-Path $batch 'batch.json.txt')).Hash} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $published 'current.json.txt')
$capture=Get-SharePointMigrationHistoryCapture -OutputRoot $published -TenantKey 'sample-prod'
if($capture.BatchId -cne 'test-batch' -or $capture.Entry.SHA256 -cne $hash){throw 'Validated predecessor history was not captured.'}
$disabledHistory=Join-Path $taskRoot 'disabled-retaining-history'
& (Join-Path $PSScriptRoot 'New-SharePointMigrationEvidence.ps1') -DataRoot $raw -OutputRoot $disabledHistory -TenantKey 'sample-prod' -PythonCommand $PythonCommand -PreviousHistoryPath $capture.Entry.SourcePath
if(@(Import-Csv -LiteralPath (Join-Path $disabledHistory $historyName)).Count -ne 2){throw 'Disabling the feature discarded its prior history.'}
[IO.File]::AppendAllText((Join-Path $batch $historyName),'tampered')
$failed=$false
try{$null=Get-SharePointMigrationHistoryCapture -OutputRoot $published -TenantKey 'sample-prod'}catch{$failed=$true}
if(-not $failed){throw 'Tampered predecessor history was accepted.'}
Write-Host 'PASS: synthetic migration generator, nine CSV contracts, identities, target observations, missing/invalid source rejection, source preservation, disabled compatibility and safe filenames.'
Write-Host 'PASS: mapping acquisition and missing-registry rollback, validated predecessor capture and tamper rejection, history retained while disabled.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAlrnsyS105vprR
# TjyWFKnGBv7cqjQreXfTcHVMN0iDtaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL/wPCeVMARpPVpg+lcvMqhm2POovE10HcS/3dBJAESXMA0GCSqG
# SIb3DQEBAQUABIIBgD0X5rDshcmKcm64cvdk9RIi80dXoB4Ru8wxWnysddO54eOR
# zaC+6r4Bx40iGxiorlbrpRYQ69MVMRgvgtu8Nk0aLW2vZw3p4e+cLIwbd9BJP4p5
# sZIU/Q3kvhWQGjCsSZ41CkRtbpBQSGGu4FoHLggJbF8K+M1ckiacdb//obSaEO1r
# TRQaELPG2bGHZfcexuA0dner9OgNaNKqtOvngKuJ+Ufh/YMgWRdr0ZA7sn4BKeLG
# z4oTECVoylufR9LepsO+s9qmsh9TGO80bIj3Zpy5QnKUKbmEVe6BJal8Zkqq5cqv
# sXcK2AhyisCPVz0ldzW/OvGzK5LHPFTQUQWfBckgk6Nf3TpCOizB4Zemh4QBSJ6E
# aj4vK828sgycMf87RKXXJE8BFZZU5Bxr4dZgLyNOigNj/bSg23ADMssrkGOjAIf4
# ynfVjBcMnm5kYGWRdTKVMm2FeUUQkJs4K0gERy6zgQiklq3iQJ66aZXnO38xjqtn
# 7qYNPnVpiNlo7At5f6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkyMDU4
# MTRaMC8GCSqGSIb3DQEJBDEiBCC+p3kqd/sUVWHe4GKpqDZOAywa8GWEESPlT3wt
# pDDHVTANBgkqhkiG9w0BAQEFAASCAgCmKycjB6lPsZWodwaO1kHSPixPnfAt/zjn
# cRfXx3gV4rNhMOj6QzHlFAr+Wn85+7qNxM1a0CnbfRiS2DStRqkidaJmxTHwJV97
# aUMnTvO6uqlZ/WkzhdaLqajudmSZGQiPbcMeXDdGUiPn6PHmTXmfCkX37iUuFTKL
# 2hzsLU0y7JqVTGzpH5emZAYNwoVoccUwQhspXK44xQDI7d6HQsrazWGVpnfqEjni
# 6y21vdCMO58f523ie9JhNZUakPE4bDSGQzspyz54o24UiOd+ZcTKeXw0DHXb+1od
# Z8yrT2kbXEog64LZMhhXEuWZWEfgig9vMgVvwWcuUOB+Lor3eB3LPiNZty6caYLM
# EQVDw3bpBnHzsvmX1bZWac4nmnYIiIWN1y8YwStbE8Ij5a4LhMqJr1dsrxxgmbtg
# cLOHiF1Vi3mtqVWRuRUA3tUNjXOkaNTaUCjkkvuoHZOwFyjQZ72xWKCMb5nFI1gg
# a7Kuy7Jja5mCs0798Ltv91aWQwROiqGfRzZObAjeX0zpfKZLPn4fl3K9g7ELtL8I
# 7pjFBLCIB14wCN4TiIQzrlkZUNRgQnM0dXrk80OCpiYVr3ls6Dw/Spi5irGAYSYz
# 5BFkgdGsRRZ0rUwc7sfjxNHLysJTgBLRiu6CCsdj3y5wKYVc6AKgOMhrJXArscl5
# ihWfuy11hg==
# SIG # End signature block
