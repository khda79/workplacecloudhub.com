# Version 1.1.0
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'PreparedArtifacts.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('WorkplaceArtifactsTest-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null
$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
function Reject([scriptblock]$Action,[string]$Message){$failed=$false;try{& $Action}catch{$failed=$true};Check $failed $Message}
try {
    $source=Join-Path $root 'source';$dest=Join-Path $root 'batch'
    New-Item -ItemType Directory -Path (Join-Path $source 'DATA-LAST'),$dest | Out-Null
    $spec=Get-Content (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw|ConvertFrom-Json
    foreach($file in $spec.currentFiles){
        $path=Join-Path $source "DATA-LAST/$file"
        if($file -like '*.csv'){@{TenantKey='synthetic';SnapshotId=('a'*32)}|Export-Csv -LiteralPath $path -NoTypeInformation}
        else{@{TenantKey='synthetic';SnapshotId=('a'*32);SchemaVersion=2;Products=@(1,2,3,4);OtherProducts=@(1,2,3);AssignmentIssues=@{Available=$false;Rows=@()}}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $path}
    }
    foreach($history in $spec.history){
        $dir=Join-Path $source "$($history.root)/2026-W40";New-Item -ItemType Directory -Path $dir -Force|Out-Null
        @{TenantKey='synthetic';Id='one'}|Export-Csv -LiteralPath (Join-Path $dir $history.file) -NoTypeInformation
        @{Week='2026-W40';SnapshotCreatedAtUtc='2026-10-01T00:00:00Z'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'manifest.json')
    }
    $plan=@(Get-PreparedArtifactSources -DataRoot $source)
    Check ($plan.Count -eq ($spec.currentFiles.Count+4+@($spec.optionalCurrentFiles).Count)) 'Incomplete capture plan.'
    Check (@($plan | Where-Object { $_.ContainsKey('IsOptional') -and $_.IsOptional }).Count -eq @($spec.optionalCurrentFiles).Count) 'Optional artifact files are not optional capture items.'
    $entries=@(Copy-PreparedArtifacts -SnapshotRoot $source -Destination $dest -TenantKey 'synthetic')
    Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $entries
    Check ($entries.Count -eq ($spec.currentFiles.Count+5)) 'Incomplete artifact manifest.'
    foreach($entry in $entries|Where-Object SourceRelative -NE 'GeneratedFromVerifiedCapture'){
        Check ((Get-FileHash (Join-Path $source $entry.SourceRelative)).Hash -eq $entry.SHA256) 'Capture bytes changed.'
    }
    Check (-not @($entries | Where-Object File -Like 'Source-Exchange_OnPrem_Servers_*').Count) 'Absent optional infrastructure was fabricated.'
    $inventory=Join-Path $source 'DATA-LAST/Exchange_OnPrem_Servers_Inventory.csv'
    $compute=Join-Path $source 'DATA-LAST/Exchange_OnPrem_Servers_Compute.csv'
    '"TenantKey";"Name";"Description"', '"synthetic";"mail-one";"quoted; semicolon"' | Set-Content -LiteralPath $inventory
    '"TenantKey";"ExchangeServerName";"MemoryGB"', '"synthetic";"mail-one";"64"' | Set-Content -LiteralPath $compute
    $infraDest=Join-Path $root 'infrastructure';New-Item -ItemType Directory -Path $infraDest|Out-Null
    $infraEntries=@(Copy-PreparedArtifacts -SnapshotRoot $source -Destination $infraDest -TenantKey 'synthetic')
    Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $infraDest -Entries $infraEntries
    Check (@($infraEntries | Where-Object File -Like 'Source-Exchange_OnPrem_Servers_*').Count -eq 2) 'Semicolon infrastructure was not captured.'
    Check ($infraEntries.Count -eq $entries.Count+2) 'Optional infrastructure changed required membership.'
    $goodCompute=Get-Content -LiteralPath $compute -Raw
    foreach($bad in @(
        ('"TenantKey";"ExchangeServerName"'+"`n"+'"foreign";"mail-one"'),
        ('"TenantKey";"ExchangeServerName"'+"`n"+'"";"mail-one"'),
        ('"OtherKey";"ExchangeServerName"'+"`n"+'"synthetic";"mail-one"'),
        ('"TenantKey";"TenantKey"'+"`n"+'"synthetic";"synthetic"'),
        ('"TenantKey";"ExchangeServerName"'+"`n"+'"synthetic";"mail-one";"extra"')
    )){
        Set-Content -LiteralPath $compute -Value $bad
        $badDest=Join-Path $root ([guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $badDest|Out-Null
        Reject {Copy-PreparedArtifacts -SnapshotRoot $source -Destination $badDest -TenantKey 'synthetic'} 'Invalid optional infrastructure tenant/schema accepted.'
        Check (-not (Test-Path -LiteralPath (Join-Path $badDest 'Source-Exchange_OnPrem_Servers_Compute.csv'))) 'Rejected optional infrastructure was copied.'
    }
    Set-Content -LiteralPath $compute -Value $goodCompute
    $copied=Join-Path $infraDest 'Source-Exchange_OnPrem_Servers_Compute.csv'
    '"TenantKey";"ExchangeServerName"', '"foreign";"mail-one"' | Set-Content -LiteralPath $copied
    $copiedEntry=@($infraEntries | Where-Object File -CEQ 'Source-Exchange_OnPrem_Servers_Compute.csv')[0]
    $copiedEntry.Bytes=(Get-Item -LiteralPath $copied).Length
    $copiedEntry.SHA256=(Get-FileHash -LiteralPath $copied).Hash
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $infraDest -Entries $infraEntries} 'Foreign optional tenant accepted despite matching size/hash.'
    Check (-not (Test-PreparedArtifactName '../current.json.txt')) 'Traversal accepted.'
    Check (-not (Test-PreparedArtifactName 'Source-secret.json.txt')) 'Unknown artifact accepted.'
    Check (-not (Test-PreparedArtifactName 'History-Windows-2026-W54-manifest.json.txt')) 'Invalid ISO week accepted.'
    Reject {Copy-PreparedArtifacts -SnapshotRoot $source -Destination $dest -TenantKey 'synthetic'} 'Existing immutable artifacts overwritten.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries @($entries+$entries[0])} 'Duplicate entry accepted.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries @($entries|Where-Object File -NE 'PreparedHistoryIndex.json.txt')} 'Missing index accepted.'
    # Retired artifacts are accepted only for explicitly reviewed historical contract hashes.
    $legacyName='Source-Exchange_HybridIdentity_Issues.csv'
    $oldHash=$spec.legacyCurrentFiles[0].contractSHA256[-1]
    Check (-not (Test-PreparedArtifactName $legacyName)) 'Retired artifact accepted for a new batch.'
    Check (Test-PreparedArtifactName $legacyName -ContractSHA256 $oldHash) 'Reviewed historical artifact refused.'
    $legacyPath=Join-Path $dest $legacyName
    [pscustomobject]@{TenantKey='synthetic';ObjectGUID='contoso-account'} | Export-Csv -LiteralPath $legacyPath -NoTypeInformation
    $legacyEntry=[pscustomobject]@{File=$legacyName;Bytes=(Get-Item $legacyPath).Length;SHA256=(Get-FileHash $legacyPath).Hash}
    $legacyEntries=@($entries)+@($legacyEntry)
    Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $legacyEntries -ContractSHA256 $oldHash
    Check $true 'Historical artifacts accepted with their intact hashes.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $entries -ContractSHA256 $oldHash} 'Historical required artifact silently lost.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $legacyEntries} 'Stale legacy file admitted to a new batch.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $legacyEntries -ContractSHA256 ('0'*64)} 'Unknown historical contract accepted.'
    Add-Content -LiteralPath $legacyPath 'tampered'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $legacyEntries -ContractSHA256 $oldHash} 'Tampered historical artifact accepted.'
    $changed=Join-Path $dest $entries[0].File;Add-Content -LiteralPath $changed 'tampered'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $entries} 'Changed bytes accepted.'
    $dest2=Join-Path $root 'batch2';New-Item -ItemType Directory -Path $dest2|Out-Null
    Reject {Copy-PreparedArtifacts -SnapshotRoot $source -Destination $dest2 -TenantKey 'foreign'} 'Wrong tenant accepted.'
    Check (@(Get-ChildItem $dest2 -File).Count -eq 0) 'Identity failure wrote output.'
    'Prepared artifact offline checks passed: '+$checks
} finally {
    $resolved=[IO.Path]::GetFullPath($root)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -match '^WorkplaceArtifactsTest-[a-f0-9]{32}$'){
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDVEL7qdA4VPnk3
# 5P54W94UVvX559+aLKG8Cr8uaHwXRKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMQ7y3pUQVF+sTWgV6juqTpylg00QaHHJnve00fVlNOsMA0GCSqG
# SIb3DQEBAQUABIIBgFzm0UGl/Os4Ab2BVF/0x89UbjSwfuENVrv5WLz0AcbuUEW5
# g/kGmWhkNzjD9DaokxHBWlYlrFOlnY4REznzJ5jPZjH4GgLElYsdhJ8b37mzrUkt
# iTaG44dNRbz1JO47e06jzI71UUHm4NKOHVQEcGOzW//BkN84OAPEMe8H1IP+0BtW
# 4gL7URtvsI/6NvU/HJ8biZR634Ytvstm/sCxe0WbRVnQpRtP2TvbPB9eDzT5aFSK
# cUltR9D0uVgZZ+Il40TALsvY5PAeEwRhmMGQjKUBG6S4P6p7HLxAmrpaMG2hIDDz
# WbA+KoVQ+p2ieLRL4tzVPymVlHpZ195U3XeL+rLW3quYsS2c+h9j94aFvYoibUjM
# D45KuT/yP20I2k71FrVPau2Vcb/uBSiGOgi1lNuJojIOIER+uFfphAdxt4EmWJOc
# LbP9Rx+ukPk0V26jyCaAg7MTwQrh9zkWb8c5LWzniFd7IDEi4w8lEL4yGVEY3elV
# CdANq4pDZc/6hSyek6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcwOTQx
# NTRaMC8GCSqGSIb3DQEJBDEiBCAxhb56PKOTCe6ZOP6LW7yAqsUBjS43VaDbir/t
# P/bQ1TANBgkqhkiG9w0BAQEFAASCAgBN48OXyTLNp54AVnfqvkPuFnYMRuFmL4eT
# vxw3tAUELyM+N07AB48u6iCsSglniDLpT/ULOJLYF+2V9NRd5QlYm8swPnjD690f
# H1LqQ9+ppNOKKDjZeGiDRVc7Js+jQIBFi4KxX108lW10vj+lAOZh+RI6MWuumCjJ
# FGvhhcut4G8JhvlMqvwPmga894OX6TP8P2QjmXRB8bxP3N2+5Ts/IAqSpYXWcsj+
# gSuYB6pcVB1MhJc+6bg7y6hq36AvVpw+DsWKJjRgo4TsJ41giwBcSY4JJ17O2a+8
# iKBCmXv2w6R0msoGZn+biqzHPDLjVjfRbDQa/h/NQwcku0R93H2gS18OSUYFgbnA
# q0JqbDN4K79P3QBNlWNz1aAXMtf0uhIeBrzceIqgChje0M9MA/E9LkfILHOfgvIK
# HZS+eK+qCBWzfzsLor9bFu3i0cuQmTgngNL9oTtkV1jyUKa1OaeBsdfgSrrKuZDr
# C8ESsyGfCQl87cHnntG2ofkwPBGNKrli2sRD3HcTqIciOFr8AMLwkworKz3j/jM5
# 4vLd6WGXXMCvqwOVjddPe2Ngt2PJ6zQ8vAMtRwmBYCLBIEG3PblUzQ1CP3RML+Ys
# QJz3eBjCb3W0Nry2JYQ3g/Ug2xYjJHLKrDrB92Rr2Dcx1yeBKoOPFTrWl2BkB1oN
# qjiqZz+KAA==
# SIG # End signature block
