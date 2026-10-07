#Requires -Version 7.0
<#
.SYNOPSIS
Offline legacy state relocation and two real generation/transfer cycles.
.VERSION
1.0.0
.NOTES
Synthetic files and mocked SharePoint only. No tenant, mail or production writes.
#>
[CmdletBinding()]
param([string]$PythonCommand='python')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CmdbTransferTest-'+[guid]::NewGuid().ToString('N'))
$base=Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartInventory/PreparedEvidence'
$python=(Get-Command $PythonCommand -CommandType Application -ErrorAction Stop|Select-Object -First 1).Source
$checks=0
function Check([bool]$Value,[string]$Message){if(-not $Value){throw $Message};$script:checks++}
function Rejected([scriptblock]$Action){$failed=$false;try{& $Action|Out-Null}catch{$failed=$true};Check $failed 'Unsafe transition recovery accepted.'}
try {
    $null=New-Item -Path $root -ItemType Directory
    & $python -B (Join-Path $PSScriptRoot 'test_cmdb_transfer.py') --fixture-root $root
    if($LASTEXITCODE -ne 0){throw 'Synthetic preparation failed.'}
    Import-Module (Join-Path $base 'SmartM365-CmdbSharePointTransfer.psm1') -Force
    Import-Module (Join-Path (Split-Path (Split-Path $base -Parent) -Parent) 'Modules/SmartM365.Core/SmartM365.SharePointJsonTransition.psd1') -Force
    $transition=Get-Module SmartM365.SharePointJsonTransition
    & $transition {function script:Get-SmartM365JsonTransportPolicy {@{Mode='JsonText';QualifiedSharePointDrives=@('synthetic')}}}
    $prepared=Join-Path $root 'DATA-POWERBI-CMDB'
    $stateRoot=Join-Path $root 'LOG-ALL/Publication/CMDB/SharePointTransition'
    $manifest=Join-Path $prepared 'current.json.txt'
    $legacyLock=$manifest+'.sharepoint-transition.lock'
    $identity=@{TenantKey='synthetic';OrganizationKey='test';EnvironmentKey='test';TenantId='synthetic-tenant'}
    $before=(Get-FileHash -LiteralPath $manifest).Hash
    [IO.File]::WriteAllBytes($legacyLock,[byte[]]@())
    $active=[IO.File]::Open($legacyLock,'Open','ReadWrite','None')
    try{Rejected {Move-SmartM365CmdbTransitionState -PreparedRoot $prepared -StateRoot $stateRoot -Identity $identity}}
    finally{$active.Dispose()}
    Check (Test-Path -LiteralPath $legacyLock) 'Active lock moved.'
    [IO.File]::WriteAllText($legacyLock,'invalid')
    Rejected {Move-SmartM365CmdbTransitionState -PreparedRoot $prepared -StateRoot $stateRoot -Identity $identity}
    Check ((Get-Item -LiteralPath $legacyLock).Length -eq 7) 'Nonempty lock changed.'
    [IO.File]::WriteAllBytes($legacyLock,[byte[]]@())
    Rejected {Move-SmartM365CmdbTransitionState -PreparedRoot $prepared -StateRoot $prepared -Identity $identity}
    $foreign=@{}+$identity;$foreign.TenantId='foreign'
    Rejected {Move-SmartM365CmdbTransitionState -PreparedRoot $prepared -StateRoot $stateRoot -Identity $foreign}
    $guard=[IO.File]::Open((Join-Path $root '.cmdb-preparation.lock'),'Open','ReadWrite','ReadWrite')
    try{$guard.Lock(0,1);Rejected {Move-SmartM365CmdbTransitionState -PreparedRoot $prepared -StateRoot $stateRoot -Identity $identity}}
    finally{$guard.Unlock(0,1);$guard.Dispose()}
    $legacyJournal=$manifest+'.sharepoint-transition.log'
    $journal='{"Phase":"Completed","DriveId":"synthetic","Target":"Folder/current.json.txt","Record":{}}'
    [IO.File]::WriteAllText($legacyJournal,$journal)
    $moved=Move-SmartM365CmdbTransitionState -PreparedRoot $prepared -StateRoot $stateRoot -Identity $identity
    Check ($moved -eq 2) 'Expected lock and journal relocation.'
    Check ([IO.File]::ReadAllText((Join-Path $stateRoot 'current.json.txt.sharepoint-transition.log')) -ceq $journal) 'Recovery journal bytes changed.'
    Check ((Get-FileHash -LiteralPath $manifest).Hash -ceq $before) 'Relocation changed the validated manifest.'
    Check (@(Get-ChildItem -LiteralPath $prepared -Force).Count -eq 47) 'Relocation did not restore the strict cohort.'
    Check ((Move-SmartM365CmdbTransitionState -PreparedRoot $prepared -StateRoot $stateRoot -Identity $identity) -eq 0) 'Relocation is not idempotent.'
    # Existing destination must never be overwritten, even for an empty lock.
    [IO.File]::WriteAllBytes($legacyLock,[byte[]]@())
    Rejected {Move-SmartM365CmdbTransitionState -PreparedRoot $prepared -StateRoot $stateRoot -Identity $identity}
    Check (Test-Path -LiteralPath $legacyLock) 'Conflicting legacy lock was lost.'
    Remove-Item -LiteralPath $legacyLock
    $cloud=@{};$uploaded=[Collections.Generic.List[string]]::new()
    $request={param($method,$uri,$body,$headers)
        if($method -ne 'GET'){throw 'Unexpected remote mutation.'}
        $error=[IO.IOException]::new('Synthetic remote 404');$error.Data['StatusCode']=404;throw $error
    }
    $remoteDownload={throw 'No remote legacy item should require downloading.'}
    $upload={param($local,$name)
        if($name -ceq 'current.json.txt'){
            $result=Invoke-SmartM365SharePointJsonNameTransition -LocalFilePath $local -StateFolderPath $stateRoot -DriveId synthetic -EncodedTargetPath 'Folder/current.json.txt' -Request $request -Download $remoteDownload
            if($result.Status -ne 'NoLegacy'){throw 'Unexpected transition result.'}
        }
        $uploaded.Add($name);$cloud[$name]=[IO.File]::ReadAllBytes($local);return $true
    }.GetNewClosure()
    $download={param($destination,$name) [IO.File]::WriteAllBytes($destination,$cloud[$name]);return $true}.GetNewClosure()
    $code='import sys; from pathlib import Path; sys.path.insert(0,sys.argv[1]); import cmdb_prepare; root=Path(sys.argv[2]); identity={"TenantKey":"synthetic","OrganizationKey":"test","EnvironmentKey":"test","TenantId":"synthetic-tenant"}; result=cmdb_prepare.prepare(root/"DATA-LAST",root/"DATA-POWERBI-CMDB","synthetic",identity); assert result["GeneratedTables"]==46; print(result["Status"])'
    foreach($cycle in 1,2){
        & $python -B -c $code $base $root
        if($LASTEXITCODE -ne 0){throw "Real synthetic generation failed: cycle $cycle."}
        $hash=(Get-FileHash -LiteralPath $manifest).Hash
        $parameters=@{PreparedRoot=$prepared;Identity=$identity;PythonPath=$python;ExpectedManifestSHA256=$hash}
        $uploaded.Clear()
        $result=Send-SmartM365CmdbPreparedSnapshot @parameters -UploadFile $upload -DownloadFile $download
        Check ($result.Status -eq 'PublishedAndReadBack' -and $result.VerifiedFiles -eq 47) "Cycle $cycle transfer failed."
        Check ($uploaded.Count -eq 47 -and $uploaded[-1] -ceq 'current.json.txt') 'Manifest was not last.'
        Check (@(Get-ChildItem -LiteralPath $prepared -Force).Count -eq 47) 'Shared transition polluted the cohort.'
        Check ((Get-SmartM365CmdbTransferPlan @parameters).CsvFiles -eq 46) 'Post-transfer cohort rejected.'
        Check (Test-Path -LiteralPath (Join-Path $stateRoot 'current.json.txt.sharepoint-transition.lock')) 'Persistent transition guard missing.'
    }
    $extra=Join-Path $prepared 'unexpected.txt'
    [IO.File]::WriteAllText($extra,'unrecognized')
    Rejected {Get-SmartM365CmdbTransferPlan @parameters}
    Remove-Item -LiteralPath $extra
    $tokens=$null;$errors=$null
    $core=[Management.Automation.Language.Parser]::ParseFile((Join-Path (Split-Path (Split-Path $base -Parent) -Parent) 'Modules/SmartM365.Core/SmartM365.Core.psm1'),[ref]$tokens,[ref]$errors)
    $function=$core.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-SmartM365SharePointCsvUpload'},$true)
    Check ($function.Extent.Text.Contains('-StateFolderPath $JsonTransitionStateFolderPath')) 'Core did not forward optional transition state.'
    $wrapper=Get-Content -LiteralPath (Join-Path $base 'SmartM365-CmdbEvidence-Publish.ps1') -Raw
    Check ($wrapper.Contains('-JsonTransitionStateFolderPath $transitionRoot')) 'Publisher did not isolate transition state.'
    Write-Host "PASS: $checks offline transition-state checks, including two real preparation/mock publication cycles. No production access."
} finally {
    $resolved=[IO.Path]::GetFullPath($root)
    if((Split-Path $resolved -Parent).TrimEnd('\','/') -cne ([IO.Path]::GetTempPath()).TrimEnd('\','/') -or
       (Split-Path $resolved -Leaf) -notmatch '^SmartM365-CmdbTransferTest-[a-f0-9]{32}$'){throw 'Unsafe fixture cleanup target.'}
    if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBR44e7CBvQT4Pw
# a/yodvJ3qS8eopBcLrtDEJcpAvUVVKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIxbK4hBwjSJxMcLwopwRwdPqQtzKmvDM9HZPrOLNPZXMA0GCSqG
# SIb3DQEBAQUABIIBgDBYi1CtdvrS/hstQpsDjtRvRcPesKxwTrKOkNhKB+ZFtG6E
# LqfdDgV/iRbJi3hE4sg8vJT17Yp1hxQv/VvWa+5x9q/E5K2+VaxpP2RbupfFO4BN
# 2BYkhZhKQq5xsSCbdBiS95sAlSot0DwtU2ab+4sU4jA4VxU4StWjdo8rxnt8hyqy
# grIt+L4K5X/9W8Um6ZQmp5DN+wchDCtUg5DTDM+IvZpULN8ASM+yb1UTjgMnAX7/
# H1/aKCy+VBbkjoVWlt0bR4rjncoFBnyfRQMtE04wiuo4Sk1hyoxpUSUxVL07U6eV
# XJGDuELaE4WE5kb5fpJm3x5T6GuzTfZDncMzTPDvmeig7FyL01WA6BgIhPnb7pQq
# EPmw7PQXFjykLLOMXxjCOah0pwe5aUjCEfDEEsmnfplZBTT6DXrRIbl7OD/XsJ/W
# C73/6zNzziNMIyBjenYNsw1TKcsIDduUIS8WAz/P7sh56qRE9j1rYGiFY8tnJRWQ
# la309Cyf2G6SJPKO86GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxMTM2
# NTFaMC8GCSqGSIb3DQEJBDEiBCC9b7rKR1J+D01QS2r/vmHKVHSso0D/LYOdU3DC
# qdxXdDANBgkqhkiG9w0BAQEFAASCAgBJEjYtVkGP1ppc35M9OXEoXN+vkJec3OGV
# JoksoLvedD/GMNbFwmFCmewiz8gTIbex352sJ7qJcHqt01rmqjVZ6QNs/D2uI7HS
# zqSTpBt88WtEZAXB2Xsso3kPJubnWYpuKphwjmslBEtNtBdoz+s2aNIpjDJ4LZdW
# OC4BLgFM+u8TBD3gRpzUZugQrW5rC8qBTOZwGZOdADfn4GV7KKTrw/bNJ2f4F2jk
# vVI311MFCRZlP5Ch6CL+OKUkhoGcMq/A5H8ILSuKBCo4x4IQ/VAIuO7FDY9jINLc
# n7rEVOKR9ya3FYNs2okOuixu+Sgy07uwsEuxFWpoXNVcEC1VtTXDzgCzuGM7ISIz
# 6CG5P2LYLSTWgYlTCVh0ffbrsnO81MQEDapy0iVEBCiwFK9V3v4pKtESxfA8lokq
# Fpoe9wSMYM0uwUwwYGlQyUG9Bb8ReJkpj37qILn20MMhits6FVd50xon4ctjIB+f
# RBmh/bmFFy23/EzYaOinzBplsvbhmDzVYM5ZXB5/ZicjNgEONT2eMjJHEDg13osV
# CgA+ztGRFznKeCZERw9rJurW3ugg8LN803BiSqDkVDLHiyTrRgWo64qtbn6MbJ5J
# +tTh2O7y7VyUF9BxyBIRHs58s5lmQP7oHc8KdmaHeuM9PavW5q8hhzLPoPpbC/bb
# UStXXyYb3w==
# SIG # End signature block
