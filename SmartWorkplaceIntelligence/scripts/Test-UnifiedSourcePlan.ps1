#Requires -Version 7.0
# Version 1.1.0
[CmdletBinding()]
param()
# Synthetic source-selection test only; no preparation, collectors or cloud calls.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('UnifiedSourcePlan-'+[guid]::NewGuid().ToString('N'))
$checks=0
function Check([bool]$Value,[string]$Message){if(-not $Value){throw $Message};$script:checks++}
function Put([string]$Relative){
    $path=Join-Path $root $Relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
    [IO.File]::WriteAllText($path,'Synthetic selection fixture; not business data.')
}
function Reject([scriptblock]$Action,[string]$Pattern){
    $message='';try{& $Action | Out-Null}catch{$message=$_.Exception.Message}
    Check ($message -match $Pattern) "Expected $Pattern; got $message"
}
try {
    $base=Get-Content (Join-Path $PSScriptRoot '../config/prepared-source-contract.json.txt') -Raw | ConvertFrom-Json
    $native=Get-Content (Join-Path $PSScriptRoot '../config/native-evidence-contract.json.txt') -Raw | ConvertFrom-Json
    $artifacts=Get-Content (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw | ConvertFrom-Json
    $registry=Get-Content (Join-Path $PSScriptRoot '../../SmartM365/Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt') -Raw | ConvertFrom-Json
    foreach($name in @($base.currentFiles)+@($native.sources.file)+@($artifacts.currentFiles)){Put "DATA-LAST/$name"}
    foreach($producer in $registry.Producers){Put "DATA-LAST/$($producer.Receipt)";Put ('DATA-LAST/'+$producer.Receipt.Replace('.current.json.txt','.run.json.txt'))}
    [IO.File]::WriteAllText((Join-Path $root 'DATA-LAST/SmartInventory_SmartM365-Devices-Inventory.current.json.txt'),'{}')
    foreach($name in @($base.mappingFiles)+@($base.dailyFiles)){Put $name}
    foreach($spec in $base.history){Put "$($spec.root)/WeeklyHistory/2026-W41/$($spec.file)"}
    foreach($spec in $base.archivedHistory){
        Put "$($spec.root)/2026-W40/$($spec.file)"
        Put "$($spec.root)/2026-W40/$($spec.manifest)"
    }
    foreach($spec in $artifacts.history){
        Put "$($spec.root)/2026-W41/$($spec.file)"
        Put "$($spec.root)/2026-W41/manifest.json.txt"
        if($spec.PSObject.Properties['optionalFile']){Put "$($spec.root)/2026-W41/$($spec.optionalFile)"}
    }
    # Intune receipt qualification has its own suite; omit that independent group
    # to keep this test focused on the composed required source selection.
    $base.PSObject.Properties.Remove('optionalSourceGroups')
    $base.PSObject.Properties.Remove('optionalCurrentFiles')
    $contract=Join-Path $root 'selection.json.txt'
    $base | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $contract -Encoding utf8NoBOM
    $plan=@(Get-PreparedSourcePlan -DataRoot $root -SourceContractPath $contract -MetadataOnly)
    Check ($plan.Count -gt 90) 'Composed source plan unexpectedly small'
    $deviceReceipt=Join-Path $root 'DATA-LAST/SmartInventory_SmartM365-Devices-Inventory.current.json.txt'
    [IO.File]::WriteAllText($deviceReceipt,'{"RequiredFiles":["Intune_DeviceIdentityEvidence.csv"]}')
    Put 'DATA-LAST/Intune_DeviceIdentityEvidence.csv'
    $modern=@(Get-PreparedSourcePlan -DataRoot $root -SourceContractPath $contract -MetadataOnly)
    Check (@($modern | Where-Object Relative -CEQ 'DATA-LAST/Intune_DeviceIdentityEvidence.csv').Count -eq 1) 'Declared device semantic evidence is missing or duplicated.'
    Remove-Item -LiteralPath (Join-Path $root 'DATA-LAST/Intune_DeviceIdentityEvidence.csv')
    Reject {Get-PreparedSourcePlan -DataRoot $root -SourceContractPath $contract -MetadataOnly} 'does not exist|Cannot find|introuvable|n.existe pas'
    [IO.File]::WriteAllText($deviceReceipt,'{}')
    Check (@($plan.SourcePath | Sort-Object -Unique).Count -eq $plan.Count) 'Physical source captured more than once'
    Check (@($plan | Where-Object Relative -Match '\\').Count -eq 0) 'Source separator normalization missing'
    foreach($name in @($base.currentFiles)+@($native.sources.file)+@($artifacts.currentFiles)){
        Check (@($plan | Where-Object Relative -CEQ "DATA-LAST/$name").Count -eq 1) "Missing or duplicate input: $name"
    }
    foreach($producer in $registry.Producers){
        Check (@($plan | Where-Object Relative -CEQ "DATA-LAST/$($producer.Receipt)").Count -eq 1) 'Native acquisition receipt not captured'
    }
    foreach($spec in $artifacts.history){
        Check (@($plan | Where-Object Relative -CEQ "$($spec.root)/2026-W41/$($spec.file)").Count -eq 1) 'Shared migration history captured twice'
    }
    foreach($spec in $base.archivedHistory){
        Check (@($plan | Where-Object Relative -CEQ "$($spec.root)/2026-W40/$($spec.file)").Count -eq 1) 'Archived application summary was not captured'
        Check (@($plan | Where-Object Relative -CEQ "$($spec.root)/2026-W40/$($spec.manifest)").Count -eq 1) 'Archived application summary manifest was not captured'
        $missingManifest=Join-Path $root "$($spec.root)/2026-W40/$($spec.manifest)"
        Remove-Item -LiteralPath $missingManifest -ErrorAction Stop
        Reject {Get-PreparedSourcePlan -DataRoot $root -SourceContractPath $contract -MetadataOnly} 'Archived history input missing'
        Put "$($spec.root)/2026-W40/$($spec.manifest)"
        $archiveFile=Join-Path $root "$($spec.root)/2026-W40/$($spec.file)"
        [IO.File]::WriteAllText($archiveFile,"AppId,DeviceCount$([Environment]::NewLine)synthetic-app,2$([Environment]::NewLine)")
        $manifest=@{Week='2026-W40';FileName=$spec.file;SHA256=(Get-FileHash -LiteralPath $archiveFile -Algorithm SHA256).Hash}
        $manifest | ConvertTo-Json | Set-Content -LiteralPath $missingManifest -Encoding utf8NoBOM
        $archivePlan=@($plan | Where-Object { $_.Kind -in 'ArchivedHistory','ArchiveManifest' })
        $module=Get-Module PreparedEvidencePipeline
        $identity=& $module {param($selection,$tenant,$snapshot) Test-PreparedSourceTenants -Plan $selection -TenantKey $tenant -SnapshotRoot $snapshot} $archivePlan 'synthetic-tenant' $root
        Check ($identity.LegacyArchivedSummaries -eq 1 -and $identity.CsvFiles -eq 1 -and $identity.Rows -eq 1) 'Manifest-verified tenantless archive was not qualified narrowly'
        $manifest.SHA256='0'*64
        $manifest | ConvertTo-Json | Set-Content -LiteralPath $missingManifest -Encoding utf8NoBOM
        Reject {& $module {param($selection,$tenant,$snapshot) Test-PreparedSourceTenants -Plan $selection -TenantKey $tenant -SnapshotRoot $snapshot} $archivePlan 'synthetic-tenant' $root} 'provenance does not match'
        $manifest.SHA256=(Get-FileHash -LiteralPath $archiveFile -Algorithm SHA256).Hash
        $manifest | ConvertTo-Json | Set-Content -LiteralPath $missingManifest -Encoding utf8NoBOM
        [IO.File]::WriteAllText($archiveFile,"TenantKey,AppId$([Environment]::NewLine)other-tenant,synthetic-app$([Environment]::NewLine)")
        $manifest.SHA256=(Get-FileHash -LiteralPath $archiveFile -Algorithm SHA256).Hash
        $manifest | ConvertTo-Json | Set-Content -LiteralPath $missingManifest -Encoding utf8NoBOM
        Reject {& $module {param($selection,$tenant,$snapshot) Test-PreparedSourceTenants -Plan $selection -TenantKey $tenant -SnapshotRoot $snapshot} $archivePlan 'synthetic-tenant' $root} 'incompatible TenantKey'
    }
    $normalPlan=@($plan | Where-Object Relative -CEQ "DATA-LAST/$($base.currentFiles[0])")
    Reject {& $module {param($selection,$tenant,$snapshot) Test-PreparedSourceTenants -Plan $selection -TenantKey $tenant -SnapshotRoot $snapshot} $normalPlan 'synthetic-tenant' $root} 'TenantKey column missing'
    $overlap='AD_Users_AllDomains.csv'
    Check (@($plan | Where-Object Relative -CEQ "DATA-LAST/$overlap")[0].Kind -eq 'Current') 'Native overlap weakened Intelligence publication freshness'
    $stale=Get-Item -LiteralPath (Join-Path $root "DATA-LAST/$overlap")
    $stale.LastWriteTimeUtc=[datetime]::UtcNow.AddHours(-200)
    Reject {Get-PreparedSourcePlan -DataRoot $root -SourceContractPath $contract -MetadataOnly} 'publication-age limit'
    $stale.LastWriteTimeUtc=[datetime]::UtcNow
    # Delete only this synthetic input; a required native receipt must fail closed.
    $missing=Join-Path $root ('DATA-LAST/'+$registry.Producers[0].Receipt)
    Remove-Item -LiteralPath $missing -ErrorAction Stop
    Reject {Get-PreparedSourcePlan -DataRoot $root -SourceContractPath $contract -MetadataOnly} 'does not exist|Cannot find|introuvable|n.existe pas'
    Write-Host "Unified source selection checks passed: $checks. No preparation or collection."
} finally {
    $resolved=[IO.Path]::GetFullPath($root)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/')
    if($resolved.StartsWith($temp+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -and
       [IO.Path]::GetFileName($resolved) -match '^UnifiedSourcePlan-[a-f0-9]{32}$' -and (Test-Path -LiteralPath $resolved)){
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJO+A1lF1YIsVY
# beMFFqvUtfJJ+GPgDZ75BqerHnsV0qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINC7R1QhRllC9mYZBrNGrThk8OG8HhjITPRFSfswiqhvMA0GCSqG
# SIb3DQEBAQUABIIBgJ/EJkGYQUoHqWEGQcR7clE2U4JuxnzHkR4tA3gYklnarpdt
# CXw4NKevJjhhDSD98/EsyrNhpaMDgkrBmrzzAe3mX73l7Zfwxu+gk5Th68f58YSC
# eh2QO3Ft+ORpXFMkiG5WeCQUrneGyWOlJ5mmBUwL3kRtz05K6Emg9WJwFnAzsIeF
# grJ8+KF9hIA0yogY+k6/7nORppURbaGaUxKZ2ohSZ0Isu3Cfdukz3tCTfXerWTA0
# fCs+G0VX7Bu87R5q2uaUsJrbCG6gs/N5ZwfKrUoIqhecgCZubP0CkZu2qI9wMlkA
# GBpAMqHUwyK7AGNnQmfi7gbEBS/YZ5aIijpaxV2tVSuFYJMxwN6Xz0vsmvjzrgol
# llHPzfdxAxH8tjdG2RsNx5MQB/EiqXbPRuXv8Rkq2Mu/G9qiib5kNRIqiXpL4f4I
# KzeFBNwrStm8j738+AtARcYQYr/GK1IyE7If9QV3Mn2aAXtU7pqFDlTZa0tnL5y0
# JTPSD55OJip5YGDkT6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMzIz
# NTVaMC8GCSqGSIb3DQEJBDEiBCBwCJkI9Jonj/L73SBBm+FgqqKoeDuZwZzMsNk+
# Dqqg5jANBgkqhkiG9w0BAQEFAASCAgBMxSoQ7qAlSDxiWj+5Dz/CGSV4WXdNRp0G
# yOhaqt45GO4odJdBzPcOBq6UNZbAGHPUrCVbHX/n5jxH1iRExptAEc6D6l1R556u
# dKZ1kjuUtA2AH5VG2UzeC5eKt1qpqQMjNPm0nTiRF5FS1/Yyh6bXTrtNlyY5uCCg
# 4S4MjgLQeBCQ/8kbeZ3G4vKyLeVJrt/gQmrfHukL2+Bvx+IpYeit5CoPydxR8vDm
# 8MQDrArsOrukPBt7k+/JssyccIHQztqRZ/68zx8RZH5Ap4B4c3Y9rqd+IBLenHD1
# VyoAnjDZKC/8wU37R0Vp9vdnQy7tJu8ksBy6yEtF1dwrN5HR31lREGtUV+RqrnLF
# qybtxtX8kWCzHYuX05zZt6szWnXaY7/2+YT5eEtI7CLJUgANtTMQH5UHudqHS6AW
# rsioA2Gv+yultym7L2UaSj91vbD7L1mwj54Vi1uiudr6Ycjby+i99ZU8w8IqOfLR
# qIxVM4bEZHgAflfEt6bvkzkSnrit6d2tKXUhNmYg5Qldz1HTGLJmnqFEdsvATqGU
# NFjeFUNHCiM68MwTu90nAzhWHd+SnpYhLP9lV/NRhc7WDkvuN+hnNA2rxl+kykao
# blFjAp8C0xzhyT2Uj5X4Vz3EbAbrZtwh47X8gD2Vo3YeJJNFP4qifx0qYa27IUFA
# GsDXq0y9/Q==
# SIG # End signature block
