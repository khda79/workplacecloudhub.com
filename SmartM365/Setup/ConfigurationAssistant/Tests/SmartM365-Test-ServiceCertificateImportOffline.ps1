$ErrorActionPreference='Stop'
$assistantRoot=Split-Path -Parent $PSScriptRoot
$root=Split-Path -Parent (Split-Path -Parent $assistantRoot)
. (Join-Path $root 'SmartInventory/Orchestrator/SmartM365-ServiceCertificateImport.ps1')
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('si-pfx-'+[guid]::NewGuid().ToString('N'))
try {
    $folder=Join-Path $fixture 'Config/Certificats';$null=New-Item -ItemType Directory -Path $folder -Force
    $thumb='A'*40;$sid='S-1-5-21-1-2-3-1001'
    [IO.File]::WriteAllText((Join-Path $folder ('SmartM365-App-'+$thumb+'-'+('b'*32)+'.pfx')),'synthetic')
    [IO.File]::WriteAllText((Join-Path $folder ('SmartM365-App-'+('C'*40)+'-'+('d'*32)+'.pfx')),'synthetic')
    $files=@(Get-SiServiceCertificatePfx -Root $fixture -Thumbprint $thumb)
    if ($files.Count -ne 1 -or $files[0].Name -notmatch ('^SmartM365-App-'+$thumb)) {throw 'PFX discovery did not match the configured certificate.'}
    $valid=[pscustomobject]@{State='Imported';Sid=$sid;Thumbprint=$thumb;HasPrivateKey=$true}
    if (-not (Test-SiServiceCertificateImportResult -ExpectedSid $sid -ExpectedThumbprint $thumb -Result $valid)) {throw 'Verified import was rejected.'}
    foreach ($invalid in @(
        [pscustomobject]@{State='Imported';Sid='S-1-5-21-1-2-3-1002';Thumbprint=$thumb;HasPrivateKey=$true},
        [pscustomobject]@{State='Imported';Sid=$sid;Thumbprint=('B'*40);HasPrivateKey=$true},
        [pscustomobject]@{State='Imported';Sid=$sid;Thumbprint=$thumb;HasPrivateKey=$false},
        [pscustomobject]@{State='Failed';Sid=$sid;Thumbprint=$thumb;HasPrivateKey=$true}
    )) {
        $rejected=$false;try {Test-SiServiceCertificateImportResult -ExpectedSid $sid -ExpectedThumbprint $thumb -Result $invalid | Out-Null} catch {$rejected=$true}
        if (-not $rejected) {throw 'An unverified import was accepted.'}
    }
    # Synthetic child exercises the real private-pipe and PID handshake, without importing a PFX.
    $currentSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $fixturePassword=ConvertTo-SecureString 'synthetic-password' -AsPlainText -Force
    $credential=[Management.Automation.PSCredential]::new('TEST\fixture',$fixturePassword)
    $launcher={param($exe,$encoded,$pipeName,$expectedSid)
        if (('"'+$exe+'" -NoProfile -NonInteractive -EncodedCommand '+$encoded).Length -ge 1024) {throw 'Worker launch command exceeds the Windows limit.'}
        $child=@'
$pipe=[IO.Pipes.NamedPipeClientStream]::new('.', '__PIPE__', [IO.Pipes.PipeDirection]::InOut)
$pipe.Connect(10000)
$reader=[IO.StreamReader]::new($pipe,[Text.Encoding]::UTF8,$false,4096,$true)
$writer=[IO.StreamWriter]::new($pipe,[Text.UTF8Encoding]::new($false),4096,$true);$writer.AutoFlush=$true
$null=$reader.ReadLine()
$writer.WriteLine('READY|'+[Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
$data=$reader.ReadLine() | ConvertFrom-Json
if ($data.Password -cne 'synthetic-password' -or [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($data.Pfx)) -cne 'synthetic') {throw 'Incorrect synthetic payload.'}
$writer.WriteLine((@{State='Imported';Sid=$data.Sid;Thumbprint=$data.Thumbprint;HasPrivateKey=$true}|ConvertTo-Json -Compress))
$reader.Dispose();$writer.Dispose();$pipe.Dispose()
'@
        $child=$child.Replace('__PIPE__',$pipeName)
        $code=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($child))
        $process=Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$code) -WindowStyle Hidden -PassThru
        $process.Id
    }
    try {
        $result=Invoke-SiServiceCertificateImport -PfxPath $files[0].FullName -PfxPassword $fixturePassword -Credential $credential -ExpectedSid $currentSid -Thumbprint $thumb -PowerShell (Get-Process -Id $PID).Path -LaunchWorker $launcher
        if ($result.State -ne 'Imported' -or $result.Sid -cne $currentSid) {throw 'Synthetic pipe protocol failed.'}
        $wrongPidLauncher={param($exe,$encoded,$pipeName,$expectedSid)
            $null=& $launcher $exe $encoded $pipeName $expectedSid
            $PID
        }.GetNewClosure()
        $rejected=$false
        try {Invoke-SiServiceCertificateImport -PfxPath $files[0].FullName -PfxPassword $fixturePassword -Credential $credential -ExpectedSid $currentSid -Thumbprint $thumb -PowerShell (Get-Process -Id $PID).Path -LaunchWorker $wrongPidLauncher | Out-Null} catch {$rejected=$_.Exception.Message -match 'identity could not be confirmed'}
        if (-not $rejected) {throw 'An unexpected pipe client was accepted.'}
        $realWorkerLauncher={param($exe,$encoded,$pipeName,$expectedSid)
            $process=Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru
            $process.Id
        }
        $rejected=$false
        try {Invoke-SiServiceCertificateImport -PfxPath $files[0].FullName -PfxPassword $fixturePassword -Credential $credential -ExpectedSid $currentSid -Thumbprint $thumb -PowerShell (Get-Process -Id $PID).Path -LaunchWorker $realWorkerLauncher | Out-Null} catch {$rejected=$_.Exception.Message -match 'failed or is uncertain'}
        if (-not $rejected) {throw 'The real worker accepted an invalid synthetic PFX.'}
        $rsa=[Security.Cryptography.RSA]::Create(2048)
        $request=[Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=Offline Fixture',$rsa,[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $syntheticCert=$request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddMinutes(-1),[DateTimeOffset]::UtcNow.AddDays(1))
        try {
            $syntheticThumb=$syntheticCert.Thumbprint
            $syntheticPath=Join-Path $folder ('SmartM365-App-'+$syntheticThumb+'-'+('e'*32)+'.pfx')
            [IO.File]::WriteAllBytes($syntheticPath,$syntheticCert.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx,'synthetic-password'))
            # Keep all keys ephemeral and replace only the OS store with an in-memory fixture.
            $inMemoryLauncher={param($exe,$encoded,$pipeName,$expectedSid)
                $child=[Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($encoded))
                $originalStore="`$store=[Security.Cryptography.X509Certificates.X509Store]::new('My',[Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)"
                $fakeStore="`$store=[pscustomobject]@{Certificates=[Collections.Generic.List[object]]::new()};`$store | Add-Member ScriptMethod Open {param(`$flags)};`$store | Add-Member ScriptMethod Add {param(`$cert) `$this.Certificates.Add(`$cert)};`$store | Add-Member ScriptMethod Close {}"
                $original64=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($originalStore))
                $fake64=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($fakeStore))
                $transform='$s=$r.ReadLine();$w=[Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($s));$a=[Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('''+$original64+'''));$b=[Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('''+$fake64+'''));if(-not $w.Contains($a)){throw ''Fixture store missing''};$w=$w.Replace($a,$b).Replace(''[Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet -bor [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet'',''[Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet'');$s=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($w));'
                if (-not $child.Contains('$s=$r.ReadLine();')) {throw 'Fixture bootstrap was not found.'}
                $child=$child.Replace('$s=$r.ReadLine();',$transform)
                $code=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($child))
                $process=Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$code) -WindowStyle Hidden -PassThru
                $process.Id
            }
            $result=Invoke-SiServiceCertificateImport -PfxPath $syntheticPath -PfxPassword $fixturePassword -Credential $credential -ExpectedSid $currentSid -Thumbprint $syntheticThumb -PowerShell (Get-Process -Id $PID).Path -LaunchWorker $inMemoryLauncher
            if ($result.State -ne 'Imported' -or $result.Thumbprint -ne $syntheticThumb) {throw 'Synthetic PFX import did not complete.'}
            $wrongPassword=ConvertTo-SecureString 'incorrect-password' -AsPlainText -Force
            try {
                $rejected=$false
                try {Invoke-SiServiceCertificateImport -PfxPath $syntheticPath -PfxPassword $wrongPassword -Credential $credential -ExpectedSid $currentSid -Thumbprint $syntheticThumb -PowerShell (Get-Process -Id $PID).Path -LaunchWorker $inMemoryLauncher | Out-Null} catch {$rejected=$_.Exception.Message -match 'failed or is uncertain'}
                if (-not $rejected) {throw 'An incorrect PFX password was accepted.'}
            } finally {$wrongPassword.Dispose()}
        } finally {$syntheticCert.Dispose();$rsa.Dispose()}
    } finally {$fixturePassword.Dispose()}
    $source=Get-Content -LiteralPath (Join-Path $root 'SmartInventory/Orchestrator/SmartM365-ServiceCertificateImport.ps1') -Raw
    if ($source -notmatch 'CreateProcessWithLogonW\(user,String\.IsNullOrEmpty\(domain\)\?null:domain,secret,1,executable,command,0x00000010' -or $source -notmatch 'LogonUserW\(user,String\.IsNullOrEmpty\(domain\)\?null:domain,secret,2,0,out token\)' -or $source -match 'CreateProcessWithTokenW') {throw 'Certificate worker must use a verified interactive logon with the user profile.'}
    if ($source -notmatch "(?s)\`$worker=@'\r?\n(.*?)\r?\n'@") {throw 'Certificate worker source was not found.'}
    $workerText=$Matches[1];$tokens=$null;$parseErrors=$null
    [Management.Automation.Language.Parser]::ParseInput($workerText,[ref]$tokens,[ref]$parseErrors) | Out-Null
    if ($parseErrors.Count) {throw 'Certificate worker source is invalid.'}
    'Service certificate import preparation offline fixtures passed.'
} finally {
    if (Test-Path -LiteralPath $fixture) {
        $resolved=(Resolve-Path -LiteralPath $fixture).Path
        if ($resolved -ne $fixture -or [IO.Path]::GetFileName($resolved) -notmatch '^si-pfx-[0-9a-f]{32}$') {throw 'Unexpected fixture path.'}
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCmZvCzqxB9pfBQ
# JuWYa0QL9IpHi5Fq96iz/YgFOPxLOqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEUbFLw0rFTPrf6wEvCwii0AfxF8om3BXz6Tp1bMIrAuMA0GCSqG
# SIb3DQEBAQUABIIBgBS0rycyATIiYaoP9oN4BNRImxp1kSGhfmW0GBmHl65GgCpo
# U7/JrVqYXpjuNJukupVHVQlcRyN52DMss24kZcJBbByQJsgdmMUA+HOgQXY9dMbN
# mZRXeXRrKC3+irGhh11GxzSRLU7jIo9WAcO6RMGnkAwhc3WxudbFVJON4AlI4GUK
# za7wbKEh+RP8eggrfcWAgW4ig3hbwvuja10OTUxUu39QgJejHl6N7l9l9wdXT/5T
# sWtNEJ41vsivBR8uiv4qMPpV3ncIBcn5x+BfLedd/W4GJ3jbcBRlueBnCk+f8Lgp
# X2DsIk2noqjXV3frwdtR2kEyuFA/EDHkA3QrNoSg78TZg1r9SDf4OasOYEk3izG5
# Yq/13cc96GP7b0pe0OcPdFhC013Z3SixWNM7C04H7AzVhblTkfAf7QXa2muSzfk5
# 53kTLRTy7pwTKJEo7+TQyLxEbrNZQyV4y3Xp2jAIyY8+ROja6mNiVVnJWGX+6of0
# J0TK7ubtCTvPBa4k2aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MDlaMC8GCSqGSIb3DQEJBDEiBCBVCNM+1g3pd8bKbObz55TXRXTbC3ElXmDybEVq
# t1KRmzANBgkqhkiG9w0BAQEFAASCAgAKPcoqgG7sZBwb6bFWswD9Z72mxiGmuq7j
# hi/7RVsoNBHg3yPGXWpKMbBgHKGl2+rutZQ4gjJSInuX5I2tLLWBrlBq7tyXoYsO
# wXkt4u7C7RPB6z+eoNAJFJQxbJOfQEfeZ/rWG28DqYref/ySSy+WYeqNun8KCyWX
# 2j2t2Bi6R6rLqxULG5Ecz3vFZeahw1zG3DjWZTdG+zgBuayQE590fxqXr/DglGZs
# dtkUX5mjw0DcgeL6QmXLjo8k2xGVXgGVc9J8G8lYickeQJkbSZypWnMF4kaPBAyM
# Z7jPRycPNj7MwhW8yh94al86R2b9kGgWgUmNhObyEUQHcrCkP7I6bDNpJPVaoDF3
# ZOtVkwYUurzUqvznQxftAv0HWYmEYzXiFRRMUd5ZOKXAuiahWzSjGScs1n+UMP6D
# 36/QkTp/1Q/GizAdIJRVQfyGRVn50p5Nd8G+W4KLR4VD0xJ/up+liiHHceIrij+N
# /zaX8R+9QizMQo9atsOBKT1nVUnhHHRBmMRJ7QB+fcQ0xd1N/OXO3oXwgWZwQoFa
# IRB3bW23rh2GgSlKH7nygrfNcNE/gATD1SvKCgrMhnQqNJdCaf+f7nD4aI1tTz0K
# XxsUKAjYcmvcb90tZZz623dbiO3jJFtKI+y1ogRCutbmVWah+Fu9lE/9qVaccLok
# SO85cBNXBA==
# SIG # End signature block
