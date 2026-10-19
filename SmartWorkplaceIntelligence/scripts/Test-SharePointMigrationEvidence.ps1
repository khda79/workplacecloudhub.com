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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAg5La3WfoZ3sRb
# 1zhMja2K+HVs+KDkZ0ckPb2epH70QaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCkc3ejRtZHOSP0xcSkQTUf
# w1Z2jZUkcLltZ9RXIOHdOzANBgkqhkiG9w0BAQEFAASCAYBzbw+GCK7EsX0nDTQ3
# lc0gLxvdGmFPwowZ3DSCzpFU2fsvBg7nXKM3YLvJpGcbrEAWzXGU3K7s3UDHFO0z
# dJeftThfLu9o8B03bXbfPx2ev50WLrJQfDGxkGfvZSZ79VcvJsyjP8aLgS3Bpr9U
# F2pM6P/K6JWSUR5R5WLv+9LtFZCXpc5CTsoKEZEHP8nVLRUmBpRkd8i+h6lv9jgI
# cnwIRgAWs2MfbDEwQSYDyJBM9yhNiPSwm9EUFNGObL8espyNm23dMJ8bzq/t5GfH
# O8L7zLf+8hUS9cTz4BcuPB5B2LrEWMDWvuh0I7S76w3y25fqzRz8+QkV0cNaQvZZ
# whnashlIxtgSWRzuP1QU/v4b7AQR5+ta5g7cEp1ueC4jOLlDp2QmLq2ITrV3VkBM
# X2weN1ywCtLa2fpIIdNPiOd4u5YOMMGbvQvtMDwOft6vQ73slussfktVsrLnuM1m
# Nb4aCljL0tFzyEMbZyngIHuVWdqVhzdtZZrUm3nuH46w8E0=
# SIG # End signature block
