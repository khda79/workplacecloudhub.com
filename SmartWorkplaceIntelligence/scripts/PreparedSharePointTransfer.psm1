Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Resolve-PreparedSharePointFolder {
    param([string]$ConfiguredPath,[string]$NormalizedDataRoot)
    $path=if([string]::IsNullOrWhiteSpace($ConfiguredPath)){
        if([string]::IsNullOrWhiteSpace($NormalizedDataRoot)){throw 'SharePoint data root is missing.'}
        $NormalizedDataRoot.TrimEnd('/','\')+'/DATA-POWERBI'
    }else{$ConfiguredPath}
    $path=($path -replace '\\','/').Trim('/')
    if($path -match '[:%?#]' -or @($path.Split('/') | Where-Object {$_ -in @('','.','..','DATA-LAST','DATA-ALL','LOG-ALL')}).Count -or $path.Split('/')[-1] -cne 'DATA-POWERBI'){
        throw 'Prepared SharePoint folder must be a library-relative DATA-POWERBI path, outside raw/log folders.'
    }
    $path
}

function Assert-TransferPath([string]$Path){
    $cursor=[IO.Path]::GetFullPath($Path);$volume=[IO.Path]::GetPathRoot($cursor).TrimEnd('\','/')
    while($cursor){
        if((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Transfer refuses linked paths.'}
        if($cursor.TrimEnd('\','/') -eq $volume){break};$cursor=Split-Path $cursor -Parent
    }
}

function Get-PreparedTransferPlan {
    param([string]$Root,[string]$TenantKey,[string]$ContractPath,[string]$ExpectedBatchId)
    if(Test-Path -LiteralPath (Join-Path $Root 'current.json')){throw 'Convert legacy metadata before transfer.'}
    $pointerPath=Join-Path $Root 'current.json.txt';Assert-TransferPath $pointerPath
    $pointer=Get-Content -LiteralPath $pointerPath -Raw | ConvertFrom-Json
    $id=[string]$pointer.BatchId
    if($id -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$' -or $pointer.SchemaVersion -ne 1 -or $pointer.TenantKey -ne $TenantKey){throw 'Invalid local pointer identity.'}
    if($ExpectedBatchId -and $id -cne $ExpectedBatchId){throw 'Current batch differs from ExpectedBatchId; transfer not started.'}
    $folder=Join-Path $Root ('batches/'+$id);Assert-TransferPath $folder
    $metadata=@{}
    foreach($name in 'batch','validation','current'){
        $path=Join-Path $folder ($name+'.json.txt');Assert-TransferPath $path
        $metadata[$name]=[pscustomobject]@{Path=$path;Bytes=(Get-Item -LiteralPath $path).Length;SHA256=(Get-FileHash -LiteralPath $path).Hash;Value=(Get-Content -LiteralPath $path -Raw | ConvertFrom-Json)}
    }
    $manifest=$metadata.batch.Value;$validation=$metadata.validation.Value;$receipt=$metadata.current.Value
    if($manifest.SchemaVersion -ne 1 -or $manifest.TenantKey -ne $TenantKey -or $manifest.BatchId -ne $id -or $metadata.batch.SHA256 -ne $pointer.ManifestSHA256 -or $metadata.current.SHA256 -ne (Get-FileHash -LiteralPath $pointerPath).Hash){throw 'Manifest or publication receipt mismatch.'}
    if($receipt.SchemaVersion -ne 1 -or $validation.Passed -ne $true -or $validation.SchemaOnly -ne $false){throw 'Batch lacks a successful full validation receipt.'}
    $contract=Get-Content -LiteralPath $ContractPath -Raw | ConvertFrom-Json
    $expected=@($contract.tables.file)
    if(@($manifest.Files).Count -ne $expected.Count -or @($validation.Files).Count -ne $expected.Count -or $expected.Count -eq 0){throw 'Prepared file count differs from the contract.'}
    $files=[Collections.Generic.List[object]]::new();$seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($entry in $manifest.Files){
        $name=[string]$entry.File
        if($name -notmatch '^[A-Za-z0-9_-]+\.csv$' -or $name -cnotin $expected -or -not $seen.Add($name)){throw 'Unexpected, unsafe or duplicate CSV name.'}
        $path=Join-Path $folder $name;Assert-TransferPath $path
        $validated=@($validation.Files | Where-Object File -CEQ $name)
        if($entry.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or $validated.Count -ne 1 -or $validated[0].SHA256 -ne $entry.SHA256 -or $validated[0].Rows -ne $entry.Rows -or (Get-Item -LiteralPath $path).Length -ne $entry.Bytes -or (Get-FileHash -LiteralPath $path).Hash -ne $entry.SHA256){throw "Local CSV integrity mismatch: $name"}
        $files.Add([pscustomobject]@{Path=$path;Relative="batches/$id/$name";Bytes=$entry.Bytes;SHA256=$entry.SHA256})
    }
    foreach($name in 'batch','validation','current'){
        $m=$metadata[$name];$files.Add([pscustomobject]@{Path=$m.Path;Relative="batches/$id/$name.json.txt";Bytes=$m.Bytes;SHA256=$m.SHA256})
    }
    # Immutable batch receipt is also the source for the final root pointer.
    $files.Add([pscustomobject]@{Path=$metadata.current.Path;Relative='current.json.txt';Bytes=$metadata.current.Bytes;SHA256=$metadata.current.SHA256})
    [pscustomobject]@{BatchId=$id;CsvCount=$expected.Count;PointerPath=$pointerPath;PointerHash=$metadata.current.SHA256;Files=@($files)}
}

function Send-PreparedEvidenceBatch {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputRoot,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][string]$ContractPath,[Parameter(Mandatory)][string]$WorkRoot,
        [Parameter(Mandatory)][scriptblock]$UploadFile,[Parameter(Mandatory)][scriptblock]$DownloadFile,
        [string]$ExpectedBatchId)
    $root=(Resolve-Path -LiteralPath $OutputRoot).ProviderPath.TrimEnd('\','/')
    if((Split-Path $root -Leaf) -ne 'DATA-POWERBI'){throw 'Transfer requires a dedicated DATA-POWERBI directory.'}
    Assert-TransferPath $root
    $scratch=[IO.Path]::GetFullPath($WorkRoot).TrimEnd('\','/')
    $data=Split-Path $root -Parent
    if($scratch -eq $data -or $scratch.StartsWith($data+'\',[StringComparison]::OrdinalIgnoreCase) -or $data.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)){
        throw 'Transfer scratch must be outside the tenant DATA tree and not its ancestor.'
    }
    $lock=[IO.File]::Open((Join-Path $root '.publication.lock'),'OpenOrCreate','ReadWrite','None')
    $run=$null;$temp=$null;$journal=$null
    try {
        $plan=Get-PreparedTransferPlan $root $TenantKey $ContractPath $ExpectedBatchId
        $run=Join-Path $scratch ('transfers/'+[guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $run -Force | Out-Null;Assert-TransferPath $run
        $temp=Join-Path $run 'readback.tmp'
        $journal=[ordered]@{TenantKey=$TenantKey;BatchId=$plan.BatchId;StartedUtc=[datetime]::UtcNow.ToString('O');Status='Transferring';PointerAttempted=$false;VerifiedFiles=[Collections.Generic.List[string]]::new();Error=$null}
        foreach($file in $plan.Files){
            if($file.Relative -eq 'current.json.txt'){
                # Verify local inputs again before activating the cloud batch.
                if((Get-FileHash -LiteralPath $plan.PointerPath).Hash -ne $plan.PointerHash){throw 'Local current pointer changed during transfer.'}
                foreach($source in $plan.Files){if((Get-FileHash -LiteralPath $source.Path).Hash -ne $source.SHA256){throw 'Local batch changed during transfer; cloud pointer not advanced.'}}
                $journal.PointerAttempted=$true
            }
            if((Get-FileHash -LiteralPath $file.Path).Hash -ne $file.SHA256){throw "Local file changed: $($file.Relative)"}
            if(-not (& $UploadFile $file.Path $file.Relative)){throw "Upload failed: $($file.Relative)"}
            if(-not (& $DownloadFile $temp $file.Relative)){throw "Read-back failed: $($file.Relative)"}
            Assert-TransferPath $temp
            if((Get-Item -LiteralPath $temp).Length -ne $file.Bytes -or (Get-FileHash -LiteralPath $temp).Hash -ne $file.SHA256){throw "Remote SHA256/size mismatch: $($file.Relative)"}
            Remove-Item -LiteralPath $temp -Force
            $journal.VerifiedFiles.Add($file.Relative)
            Write-Host ('SharePoint verified {0}/{1}: {2}' -f $journal.VerifiedFiles.Count,$plan.Files.Count,$file.Relative)
        }
        $journal.Status='Success'
        [pscustomobject]@{BatchId=$plan.BatchId;CsvFiles=$plan.CsvCount;VerifiedFiles=$journal.VerifiedFiles.Count;PointerVerified=$true;CsvRecalculated=$false;AuditPath=(Join-Path $run 'transfer.json')}
    } catch {
        if($journal){$journal.Status='Failed';$journal.Error=$_.Exception.Message}
        if($journal -and $journal.PointerAttempted){throw "Cloud pointer publication was attempted but completion is not confirmed. Do not refresh Power BI; retry the same batch. $($_.Exception.Message)"}
        throw
    } finally {
        try{
            if($temp -and (Test-Path -LiteralPath $temp)){Assert-TransferPath $temp;Remove-Item -LiteralPath $temp -Force}
            if($journal){$journal['EndedUtc']=[datetime]::UtcNow.ToString('O');$journal | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $run 'transfer.json') -Encoding utf8}
        }finally{$lock.Dispose()}
    }
}
Export-ModuleMember -Function Resolve-PreparedSharePointFolder,Send-PreparedEvidenceBatch

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAeLmzlgec442rk
# WRlasOua4OUSMXSX5Y12hEGPwKnnxKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFKvSIzvIoAZ+o1foZawnfhZcdaE/OGZGhYzjT3bRMmCMA0GCSqG
# SIb3DQEBAQUABIIBgJmJtPkV4EKLTMqUUU8vceoEZiCTIr8z4dQqGam1xOa+pLDG
# QFUc6ANGRMk8zfznHyIyfWHGZ2Jut3nGqeSKacdhU/SRH01eplG7vQpnj901of3T
# rbZy5cl6l8swyEZCIOkajbj1KP+WLTCqGReOYnmm4RDErkzxnDzYSH0XqCwHdVpo
# mvB6jCtEHn4wmGDWB8ZO7UsAvay46uPbsEFkuTN5Ag/XgWbNXlx2wOO6A9QtWjf8
# ESaRNNcfKuojiRhEJVILQSAYA8aXgFwEV0feQzRvMvHngyS1YMD5KMZGBPvAhEMJ
# dEhtmZBte1VURLz82MvCvNRrdyZ4ggP60qeTLKwV415G1Wyc2hU3vkMwVJkktnr4
# laEjgDsDezKOxx9Dbe6SE/kFNiQjeFRnv0kFhUDrX+XA/iJLcy8TyXaQDoAkSaVn
# ejlpy408HNDPIFQ2lVY8VJZBfqZoykjfMPWSY2Ex4u4EoRPJzH8nd1aawA6MG9Nd
# STn2b1WEH27qPhrQ+6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxMzU0
# MTBaMC8GCSqGSIb3DQEJBDEiBCAFqMcgbeVPxSyArVzqalUINpBC9VGRSlbjh/Mq
# jB/cZzANBgkqhkiG9w0BAQEFAASCAgBCuR6NFkEggsFo59VGxb769LmHDRhSZCLh
# d3H1bePMNeLdJjzWkXqE20wNVFgihYCbu/FzBV5MqYgkm8pGV4CrTXKPek9wao2+
# QGiNqeIKgOAgnFhUHuKsHOPoFquCkO5KjaGTpMPn6KFdvKgrILQNBeSzFRP1tpKf
# jH9kObcIGIW91EJUmAJbx9V2lM5FVgvL4UrePnb5MzTFvEEJBm4DKPV73h4CBogX
# VSc5GhokHP/8ObvgRbc63MOpfsipLDkwUXwFSlJndlqOhmFHbOJt8tko6Tl0XLFs
# ycDbFmEaMs8rIp/v+3G+FAkvz7wx0WxR2CYvvJ3oUExSVKLUcSGzWtObr9iLcPV4
# FKC2yA85vWCwJ1hzlRStQD9UcpZHnBUHPXCidDg+GNn7zX7tY3XZBO0a7SEzy4V3
# jhAqeBIEXFUYYSy7KDJPsDJNcCJ9FQhJWwd5Ed3ot4kVcYetA6SGbA4kFFHiNHjR
# CCy24MWyHjFtnqxM5kjWX71sGSJoJ2LZ8yX/r8OY6JR/BoToKwizinSUmbB/E6mp
# 3f+GF+eV9e590UokYSEjNm8FHphThB0Jj1zQQh6/ki/1FeICJti/K960E0nJaj0/
# OWOExtFdylFdWT1L9R1SKf7hsD9xhTujSU6EmHHn0jAZOXUFnu9WaMvBOTcmJjfR
# 5PWnyhojXg==
# SIG # End signature block
