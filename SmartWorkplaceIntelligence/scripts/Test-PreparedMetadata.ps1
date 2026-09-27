[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedMetadata.psm1') -Force
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$script:checks=0
function Check($condition,[string]$message){if(-not $condition){throw $message};$script:checks++}
function Reject([scriptblock]$action){$failed=$false;try{& $action | Out-Null}catch{$failed=$true};Check $failed 'Expected rejection'}
function Fixture([string]$name){
    $output=Join-Path $root ($name+'/DATA-POWERBI')
    $id='20260101T010101001Z-1234abcd';$folder=Join-Path $output ('batches/'+$id)
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $csv=Join-Path $folder 'Trend.csv'
    [pscustomobject]@{TenantKey='synthetic';Date='2026-01-01';Value=1} | Export-Csv -LiteralPath $csv -NoTypeInformation
    $file=@{File='Trend.csv';Rows=1;Bytes=(Get-Item $csv).Length;SHA256=(Get-FileHash $csv).Hash}
    @{SchemaVersion=1;TenantKey='synthetic';BatchId=$id;Files=@($file)} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $folder 'batch.json')
    @{Passed=$true;SchemaOnly=$false;Files=@($file);Errors=@()} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $folder 'validation.json')
    $pointer=@{SchemaVersion=1;TenantKey='synthetic';BatchId=$id;PreviousBatchId=$null;ManifestSHA256=(Get-FileHash (Join-Path $folder 'batch.json')).Hash}
    $pointer | ConvertTo-Json | Set-Content (Join-Path $folder 'current.json')
    Copy-Item (Join-Path $folder 'current.json') (Join-Path $output 'current.json')
    [pscustomobject]@{Root=$output;Folder=$folder;Csv=$csv}
}
$fixture=Fixture 'success';$hash=(Get-FileHash $fixture.Csv).Hash
$before=@(Get-ChildItem $fixture.Root -Recurse -Filter '*.json' | ForEach-Object {[pscustomobject]@{Path=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
$preview=Convert-PreparedMetadataNames $fixture.Root 'synthetic'
Check (-not $preview.Applied -and $preview.MetadataFiles -eq 4 -and $preview.CheckedCsvFiles -eq 1) 'Preview result'
Check (@(Get-ChildItem $fixture.Root -Recurse -Filter '*.json.txt').Count -eq 0) 'Preview changed metadata'
$applied=Convert-PreparedMetadataNames $fixture.Root 'synthetic' -Apply
Check ($applied.Applied -and $applied.MetadataFiles -eq 4) 'Conversion result'
foreach($item in $before){Check ((Get-FileHash ($item.Path+'.txt')).Hash -eq $item.Hash) 'Metadata bytes changed';Check (-not(Test-Path $item.Path)) 'Legacy name retained'}
Check ((Get-FileHash $fixture.Csv).Hash -eq $hash) 'CSV changed'
Check ((Convert-PreparedMetadataNames $fixture.Root 'synthetic' -Apply).MetadataFiles -eq 0) 'Not idempotent'
Check ((Get-PreparedMetadataPath $fixture.Root 'current') -eq (Join-Path $fixture.Root 'current.json.txt')) 'New name not preferred'
Reject {Convert-PreparedMetadataNames $fixture.Root 'other' -Apply}
$lock=[IO.File]::Open((Join-Path $fixture.Root '.publication.lock'),'OpenOrCreate','ReadWrite','None')
try{Reject {Convert-PreparedMetadataNames $fixture.Root 'synthetic' -Apply}}finally{$lock.Dispose()}
$bad=Fixture 'tampered';Add-Content -LiteralPath $bad.Csv -Value 'corrupt'
Reject {Convert-PreparedMetadataNames $bad.Root 'synthetic' -Apply}
Check (Test-Path (Join-Path $bad.Root 'current.json')) 'Failure changed root'
Check (@(Get-ChildItem $bad.Root -Recurse -Filter '*.json.txt').Count -eq 0) 'Failure partly renamed metadata'
$mixed=Fixture 'resume';Move-Item (Join-Path $mixed.Folder 'batch.json') (Join-Path $mixed.Folder 'batch.json.txt')
Check ((Convert-PreparedMetadataNames $mixed.Root 'synthetic' -Apply).MetadataFiles -eq 3) 'Interrupted conversion could not resume'
$both=Fixture 'ambiguous';Copy-Item (Join-Path $both.Root 'current.json') (Join-Path $both.Root 'current.json.txt')
Check ((Convert-PreparedMetadataNames $both.Root 'synthetic' -Apply).MetadataFiles -eq 4) 'Identical pair did not converge'
Check (-not(Test-Path (Join-Path $both.Root 'current.json'))) 'Identical legacy pointer retained'
$conflict=Fixture 'conflict';Copy-Item (Join-Path $conflict.Root 'current.json') (Join-Path $conflict.Root 'current.json.txt')
Add-Content (Join-Path $conflict.Root 'current.json.txt') ' '
Reject {Convert-PreparedMetadataNames $conflict.Root 'synthetic' -Apply}
$malformed=Fixture 'malformed';[IO.File]::WriteAllText((Join-Path $malformed.Root 'current.json.txt'),'invalid')
Reject {Convert-PreparedMetadataNames $malformed.Root 'synthetic' -Apply}
Reject {Get-PreparedMetadataPath $malformed.Root 'current'}
$directory=Fixture 'preferred-directory';New-Item -ItemType Directory (Join-Path $directory.Root 'current.json.txt') | Out-Null
Reject {Get-PreparedMetadataPath $directory.Root 'current'}
Reject {Convert-PreparedMetadataNames $root 'synthetic' -Apply}
$chain=Fixture 'chain';$previousId='20251201T010101001Z-1234abcd'
$previousFolder=Join-Path $chain.Root ('batches/'+$previousId)
Copy-Item $chain.Folder $previousFolder -Recurse
$manifestPath=Join-Path $previousFolder 'batch.json'
$manifest=Get-Content $manifestPath -Raw | ConvertFrom-Json
$manifest.BatchId=$previousId;$manifest | ConvertTo-Json -Depth 8 | Set-Content $manifestPath
$receipt=Get-Content (Join-Path $previousFolder 'current.json') -Raw | ConvertFrom-Json
$receipt.BatchId=$previousId;$receipt.ManifestSHA256=(Get-FileHash $manifestPath).Hash
$receipt | ConvertTo-Json | Set-Content (Join-Path $previousFolder 'current.json')
$pointer=Get-Content (Join-Path $chain.Root 'current.json') -Raw | ConvertFrom-Json
$pointer.PreviousBatchId=$previousId
$pointer | ConvertTo-Json | Set-Content (Join-Path $chain.Root 'current.json')
$pointer | ConvertTo-Json | Set-Content (Join-Path $chain.Folder 'current.json')
$result=Convert-PreparedMetadataNames $chain.Root 'synthetic' -Apply
Check ($result.Batches -eq 2 -and $result.MetadataFiles -eq 7 -and $result.CheckedCsvFiles -eq 2) 'Retained history chain not migrated'
Check ($result.Files[-1].To -eq (Join-Path $chain.Root 'current.json.txt')) 'Root pointer not last'
$hashBad=Fixture 'manifest-tampered';Add-Content (Join-Path $hashBad.Folder 'batch.json') ' '
Reject {Convert-PreparedMetadataNames $hashBad.Root 'synthetic' -Apply}
Check (Test-Path (Join-Path $hashBad.Root 'current.json')) 'Manifest mismatch changed pointer'
$archives=Fixture 'archives'
$retired=Join-Path $archives.Root ('retired/'+(Split-Path $archives.Folder -Leaf))
New-Item -ItemType Directory $retired -Force | Out-Null
foreach($name in 'batch','current','validation'){Copy-Item (Join-Path $archives.Folder ($name+'.json')) $retired}
$failedFolder=Join-Path $archives.Root 'failed/20260102T010101001Z-1234abcd'
New-Item -ItemType Directory $failedFolder -Force | Out-Null
@{TenantKey='synthetic';BatchId='20260102T010101001Z-1234abcd';Error='Synthetic failure';Utc='2026-01-02'}|ConvertTo-Json|Set-Content (Join-Path $failedFolder 'failure.json')
$result=Convert-PreparedMetadataNames $archives.Root 'synthetic' -Apply
Check ($result.MetadataFiles -eq 8) 'Retired or failed metadata not migrated'
Check (-not(Test-Path (Join-Path $retired 'Trend.csv'))) 'Retired CSV payload was recreated'
$transport=Get-Module SmartM365.JsonTransport;$metadata=Get-Module PreparedMetadata
$originalPolicy=& $transport {(Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock}
try{
    & $transport {function script:Get-SmartM365JsonTransportPolicy {@{Mode='JsonText';QualifiedUncRoots=@()}}}
    & $metadata {function script:Get-SmartM365JsonTransportPolicy {@{Mode='JsonText';QualifiedUncRoots=@()}}}
    foreach($family in 'transfers','DATA-REPAIR-BACKUPS','workforce-diagnostics'){
        $id=if($family -eq 'transfers'){[guid]::NewGuid().ToString('N')}else{'20260101T010101001Z-1234abcd'}
        $folder=Join-Path $root ($family+'/'+$id);New-Item -ItemType Directory $folder -Force|Out-Null
        $name=switch($family){'transfers'{'transfer'} 'DATA-REPAIR-BACKUPS'{'repair'} default{'environment'}}
        $value=switch($family){'transfers'{@{TenantKey='synthetic';BatchId='20260101T010101001Z-1234abcd';Status='Failed'}} 'DATA-REPAIR-BACKUPS'{@{TenantKey='synthetic';SchemaVersion=1;Files=@();Status='Failed'}} default{@{Publication=$false;WorkerSHA256=('A'*64);MonitorSHA256=('B'*64)}}}
        $path=Join-Path $folder ($name+'.json');$value|ConvertTo-Json|Set-Content $path
        $before=(Get-FileHash $path).Hash
        Convert-PreparedAuditNames -Root $root -Family $family -TenantKey synthetic
        Check ((Get-FileHash ($path+'.txt')).Hash -eq $before -and -not(Test-Path $path)) "Audit bytes changed: $family"
        Convert-PreparedAuditNames -Root $root -Family $family -TenantKey synthetic
        Check ((Get-FileHash ($path+'.txt')).Hash -eq $before) "Audit not idempotent: $family"
    }
    $auto=Fixture 'automatic';Initialize-PreparedMetadataNames -OutputRoot $auto.Root -TenantKey synthetic
    Check (Test-Path (Join-Path $auto.Root 'current.json.txt')) 'Owner initialization did not integrate conversion'
}finally{
    & $transport {param($p)Set-Item Function:script:Get-SmartM365JsonTransportPolicy $p} $originalPolicy
    & $metadata {Remove-Item Function:script:Get-SmartM365JsonTransportPolicy -ErrorAction SilentlyContinue}
}
& (Get-Module PreparedMetadata) {
    $script:seen=[Collections.Generic.List[string]]::new()
    # Mock filesystem inspection: no network access during this UNC regression check.
    function Get-Item {param([string]$LiteralPath,[switch]$Force) $script:seen.Add($LiteralPath);[pscustomobject]@{Attributes=[IO.FileAttributes]::Directory}}
    Assert-PreparedUnlinkedPath '\\synthetic-server\synthetic-share\DATA\DATA-POWERBI'
    if($script:seen.Count -ne 3 -or $script:seen[-1] -ne '\\synthetic-server\synthetic-share'){throw 'UNC traversal escaped share root'}
}
$script:checks++
Write-Host "PASS: $script:checks metadata conversion, compatibility, integrity, preview, idempotence and safety checks. Synthetic data only."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1+iO4uemRThEt
# WRB0RN4NUeGWNTiSo/qvh4DxDvqQ+qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGhFCrJVt5bVusqz3FXQK7e0OytN98QpGbbmIC9QPl/TMA0GCSqG
# SIb3DQEBAQUABIIBgDntLp4kR3a2lAZzWT4GT7EduG3NFVm6HlwIC/BazL+9xjGE
# 6zu8aO+xL5t1ASY3EWcxC4ZP1EUEs438mQ++74UaDuAdaloVtPv/fHVfPnlzCKBM
# 1eGBopieQTKaKvrosaCGPmYGrOsGLhDEo58ACFDVHofzRNN6r9iicdoL7dFvVzYU
# frlrpfsDK0aD4Rv5dPgnVPJFTglgvEijfvJevkMfN6cGcgJxiL7ONjbSUgQ0QhmR
# KSanq589Ab3X9fBCaJKELjX9+SxvW9jIvDPIVarRhv8EN89t4EyqggwoJriFQTLg
# 2Y1YHurOtfaUdnezZj7OpmWGe+UfXiAQDSJXB/gbIxpKoCE4I6dRMIdoXFnvVrGM
# LxfMWEnAQSGDMxoSXuW3ErnadXb3ENB9YeNqK6ZB3Nv6uEwiyfrIWFGcOM8l9NJC
# jHn4ygkR30IxpnhwrxtELmvRLqVtdSYrsi6wRxsyomdU4v/0XC2GOmCCmw2JDpDk
# xezvweFI2aU1NwnakaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU4
# MDBaMC8GCSqGSIb3DQEJBDEiBCDVJTVvN5HZi4zePSJfCTEhouPQrc0mybIuWK0i
# qTbHrzANBgkqhkiG9w0BAQEFAASCAgCsUoagmp01Ulw/YHbi/NK8svfxahgZYCj6
# HgCqEEBOn7R/H8efuNPykXj1Gxb+RwTyKTnXiK07wRTGf2blx1l8v/DeCvYegghS
# IYSuIM+M2Rb49YlkQGxkBHbFxujtS5fMZCHL71p2xWowpt9bnDGIniYDg/NbCRPS
# GdQu68nLcdgrFq9sG/JKrbEgQoMiqchmim8ZEFMT0RmbKCybyJEZDAqC0I/cqVcR
# u7qCqqxJx5LIzFtpnP9hyAfpL1s7c1y2Urfs/KHjPFGoFZJgr306Vl4Fx5Pdl2Gp
# S6wbRV1P2uhWcYz8pkokNXUKxHa4TFH7+RVgIrFXe9s7sw4tc8ese+6uvoXkDgqr
# zq+5ZbZS9JSqR3UQvXZ4ZRFNYZqLigkZWz8zhjU+rYTn9d0GVZExgzke5Y7l7/7V
# 7L6NKfp1bW0esrsKH3Kpe7wadMtz685hNzD0Aszb8B1KYejmHFudYS8T54MeO0a9
# u+9d51qtPD2vAZ/Beko2SOaTMNpbE1+qzJb1a9Z/0p6LeWV/Vszv/ewXdXU4xtzf
# DOQGGHHGTlM/U0JzUATllt50dpAdfIJmnuaIJQNSqxH3qcVv+q9OfOFr/3invbh3
# y1NIkQL6vPfDMasuRKsFR6s4vuXJSr7o77SNAoLf0RSkAp1pCHbDn8T1E0I3339+
# hH+HzLVftQ==
# SIG # End signature block
