# Interactive read-only permission verification process, version 0.1.0.
function Start-SiPermissionCheck {
    param([Parameter(Mandatory)]$Plan,[string]$Executable=(Join-Path $PSHOME 'pwsh.exe'),[string]$Parent=([IO.Path]::GetTempPath()),[scriptblock]$Launch)
    $fresh=New-SiPermissionCheckPlan -Root $Plan.Root -Tenant $Plan.Tenant -CatalogPath $Plan.CatalogPath -Outside $Plan.Outside -ExpectedScope $Plan.ExpectedScope -Worker $Plan.Worker
    foreach ($key in @('TenantId','AppId','Sender','ProfilePath','ProfileHash','CatalogPath','CatalogHash','UndeclaredJobs')) {if ($fresh.$key -cne $Plan.$key) {throw 'Permission-check inputs changed; review again.'}}
    if (($fresh.RequiredRoles -join "`n") -cne ($Plan.RequiredRoles -join "`n") -or $fresh.Sources.Count -ne $Plan.Sources.Count) {throw 'Permission-check inputs changed; review again.'}
    if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) {throw 'PowerShell is unavailable.'}
    $parentPath=Assert-SiLocalAssessmentPath $Parent
    if (-not (Test-Path -LiteralPath $parentPath -PathType Container)) {throw 'A local temporary directory is required.'}
    $directory=Join-Path $parentPath ('SmartM365-permission-'+[guid]::NewGuid().ToString('N'))
    $security=[Security.AccessControl.DirectorySecurity]::new(); $security.SetAccessRuleProtection($true,$false)
    $inherit=[Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit'
    foreach ($sid in @([Security.Principal.WindowsIdentity]::GetCurrent().User,[Security.Principal.SecurityIdentifier]::new('S-1-5-18'),[Security.Principal.SecurityIdentifier]::new('S-1-5-32-544'))) { $security.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($sid,[Security.AccessControl.FileSystemRights]::FullControl,$inherit,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Allow)) }
    [IO.FileSystemAclExtensions]::Create([IO.DirectoryInfo]::new($directory),$security)
    $locks=[Collections.Generic.List[object]]::new();$attempted=$false;$process=$null
    try {
        for ($i=0;$i -lt $fresh.Sources.Count;$i++) {
            if ($fresh.Sources[$i].Path -cne $Plan.Sources[$i].Path -or $fresh.Sources[$i].Hash -cne $Plan.Sources[$i].Hash) {throw 'Permission-check components changed; review again.'}
            $locks.Add([IO.File]::Open($fresh.Sources[$i].Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read))
            if ((Get-FileHash -LiteralPath $fresh.Sources[$i].Path -Algorithm SHA256).Hash -cne $fresh.Sources[$i].Hash) {throw 'Permission-check components changed; review again.'}
        }
        $inputPath=Join-Path $directory 'input.json';$outputPath=Join-Path $directory 'result.json'
        $payload=@{TenantId=$fresh.TenantId;AppId=$fresh.AppId;Sender=$fresh.Sender;Outside=$fresh.Outside;ExpectedScope=$fresh.ExpectedScope;RequiredRoles=@($fresh.RequiredRoles)} | ConvertTo-Json -Depth 10
        $inputStream=[IO.File]::Open($inputPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try { $writer=[IO.StreamWriter]::new($inputStream,[Text.UTF8Encoding]::new($false)); try {$writer.Write($payload)} finally {$writer.Dispose()} } finally {$inputStream.Dispose()}
        $quote={param([string]$s) "'"+$s.Replace("'","''")+"'"}
        $command='& '+(& $quote $fresh.Worker)+' -InputPath '+(& $quote $inputPath)+' -OutputPath '+(& $quote $outputPath)+'; exit $LASTEXITCODE'
        $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
        if ($null -eq $Launch) {$Launch={param($exe,$encodedCommand) Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-EncodedCommand',$encodedCommand) -WindowStyle Normal -PassThru -ErrorAction Stop}}
        $attempted=$true;$process=& $Launch $Executable $encoded
        if ($null -eq $process) {throw 'Process not returned.'}
        [pscustomobject]@{Process=$process;Locks=$locks;Parent=$parentPath;Directory=$directory;InputPath=$inputPath;OutputPath=$outputPath;Plan=$fresh;Stopped=$false}
    } catch {
        foreach ($lock in $locks) {$lock.Dispose()}
        if ($attempted) {throw ('Permission-check launch is uncertain. Inspect the connection and private temporary folder before retrying: '+$directory)}
        throw ('Permission check did not start. The private temporary folder was retained: '+$directory)
    }
}
function Stop-SiPermissionCheck {param([Parameter(Mandatory)]$Run) $Run.Stopped=$true; if (-not $Run.Process.HasExited) {$Run.Process.Kill($true)} }
function Complete-SiPermissionCheck {
    param([Parameter(Mandatory)]$Run)
    if (-not $Run.Process.HasExited) {throw 'Permission check is still running.'}
    try {
        if ($Run.Stopped -or $Run.Process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $Run.OutputPath -PathType Leaf)) {return [pscustomobject]@{State='Uncertain';Detail=('Check interrupted or failed. Inspect the connection and retained private folder before retrying: '+$Run.Directory);OperationallyQualified=$false}}
        $fresh=New-SiPermissionCheckPlan -Root $Run.Plan.Root -Tenant $Run.Plan.Tenant -CatalogPath $Run.Plan.CatalogPath -Outside $Run.Plan.Outside -ExpectedScope $Run.Plan.ExpectedScope -Worker $Run.Plan.Worker
        if ($fresh.ProfileHash -cne $Run.Plan.ProfileHash -or $fresh.ProfilePath -cne $Run.Plan.ProfilePath -or $fresh.CatalogHash -cne $Run.Plan.CatalogHash) {return [pscustomobject]@{State='Uncertain';Detail=('Profile or catalog changed during the check. Results retained for review: '+$Run.Directory);OperationallyQualified=$false}}
        $result=Get-Content -LiteralPath $Run.OutputPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($result.OperationallyQualified -isnot [bool] -or $result.OperationallyQualified -ne $false -or $result.Graph.OperationallyQualified -isnot [bool] -or $result.Graph.OperationallyQualified -ne $false -or $result.Exchange.OperationallyQualified -isnot [bool] -or $result.Exchange.OperationallyQualified -ne $false -or $result.Graph.State -notin @('ReviewRequired','Uncertain','ObservedReviewRequired') -or $result.Exchange.State -notin @('ReviewRequired','Uncertain','ObservedReviewRequired')) {throw 'Invalid permission-check result.'}
        if ($result.Graph.Missing -isnot [array] -or $result.Graph.Missing.Count -gt 500 -or @($result.Graph.Missing | Where-Object {$_ -isnot [string] -or $_ -notmatch '^[A-Za-z0-9._-]{1,128}$'}).Count -or $result.Graph.BroadGraphMailSend -isnot [bool] -or $result.Exchange.UnexpectedOutsideAllow -isnot [bool] -or ($result.Graph.UnmappedAssignments -isnot [long] -and $result.Graph.UnmappedAssignments -isnot [int]) -or $result.Graph.UnmappedAssignments -lt 0 -or ($result.Exchange.OtherApplicationAssignments -isnot [long] -and $result.Exchange.OtherApplicationAssignments -isnot [int]) -or $result.Exchange.OtherApplicationAssignments -lt 0) {throw 'Invalid permission-check details.'}
        if ($result.Graph.State -eq 'ObservedReviewRequired' -and ($result.Graph.BroadGraphMailSend -or $result.Graph.Missing.Count -or $result.Graph.UnmappedAssignments)) {throw 'Inconsistent Graph result.'}
        if ($result.Exchange.State -eq 'ObservedReviewRequired' -and ($result.Exchange.UnexpectedOutsideAllow -or $result.Exchange.OtherApplicationAssignments)) {throw 'Inconsistent Exchange result.'}
        $detail='Graph: {0}; missing declared roles: {1}; broad Graph Mail.Send: {2}; unmapped assignments: {3}. Exchange: {4}; unexpected outside authorization: {5}; other application assignments: {6}. Enabled catalog jobs without role declarations: {7}. The selected catalog may not be effective. Propagation, cached authorization, other permissions, delivery and workload execution remain unverified.' -f $result.Graph.State,(@($result.Graph.Missing) -join ', '),$result.Graph.BroadGraphMailSend,$result.Graph.UnmappedAssignments,$result.Exchange.State,$result.Exchange.UnexpectedOutsideAllow,$result.Exchange.OtherApplicationAssignments,$fresh.UndeclaredJobs
        $state=if ($result.Graph.State -eq 'ReviewRequired' -or $result.Exchange.State -eq 'ReviewRequired') {'ReviewRequired'} elseif ($result.Graph.State -eq 'Uncertain' -or $result.Exchange.State -eq 'Uncertain') {'Uncertain'} else {'ObservedReviewRequired'}
        $retained=''
        try {
            $root=[IO.Path]::GetFullPath($Run.Directory)
            if ($root -ine $Run.Directory -or [IO.Path]::GetFileName($root) -notmatch '^SmartM365-permission-[0-9a-f]{32}$' -or (Split-Path -Parent $root) -ine $Run.Parent -or (Assert-SiLocalAssessmentPath $root) -ine $root -or $Run.InputPath -ine (Join-Path $root 'input.json') -or $Run.OutputPath -ine (Join-Path $root 'result.json')) {throw 'Unexpected temporary directory.'}
            [IO.File]::Delete((Join-Path $root 'input.json'));[IO.File]::Delete((Join-Path $root 'result.json'));[IO.Directory]::Delete($root,$false)
        } catch {$retained=$Run.Directory; $detail+=' Temporary files require manual review: '+$retained}
        [pscustomobject]@{State=$state;Detail=$detail;OperationallyQualified=$false;RetainedPath=$retained}
    } catch { [pscustomobject]@{State='Uncertain';Detail=('Permission result could not be interpreted. Private files retained for review: '+$Run.Directory);OperationallyQualified=$false} }
    finally {$Run.Process.Dispose();foreach ($lock in $Run.Locks) {$lock.Dispose()}}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDZkk4yx75Xa2pO
# QuD+kwb0Y/sS0+j2Zt+hve88PxbOcaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICNsDL9vdsUT40ajI24uLiU49Qo9+JL8EKRlDCES4qMPMA0GCSqG
# SIb3DQEBAQUABIIBgIrdX7J0glPVBIZjM2QzqqsE5BE+jUY4XiMadSfu0msg6yRt
# 36/BQ/Y+nmQ0JGj/IcOoMx4Xj6Iow2LLz3zIvKt1q0FblmM/SxQ+md4+IGfSsWMf
# /v0XrGP0KRJR01BFj87HIzEhjwbpuSEYAl5gem1IyU3HeGFjjBMB3P2xLRAYXJcT
# 0KoS5GDyn2Kc6KjgEiBFe6Nsst4HwF8k3RTILi95091gEih+0CwU4UH7F4PPhOLa
# OZ4zv6DXMf2lSQpsxeJxAWVjbM50VLXSah/Pek22ozMBUM0wrAx3f82csIWsLm9Q
# 0KihOlaSvawZbSD9eGuNlBVe8rDQ7f9TCdUjXu21CDVJtUKNaM8gSMbr2VMErGK+
# eR3u7YSPZx5NQrY3cRC++Sq/LlrshRjGZaCFlI3PEM+kYkzSCYfOlX6Hj55ntnsB
# xqcYGB+tvdCZoSgNwNURGN1vr/GRk68CKOQlwvlTZINW0ehW6yopCAHQToIAJ3vq
# 7qncLNVpnY9NNKbA16GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MTlaMC8GCSqGSIb3DQEJBDEiBCDRgX4ton2c2XKGKy5P5zUdVg448O82CKYC5Ggt
# Ff9SlzANBgkqhkiG9w0BAQEFAASCAgBEv+DLeNhpqhnmNLQNHZBAxoy6DCRu0wSC
# uvOhJjtE2JzEA1P4NUb70D4vTqZaeAvSlyu2VpMsfOevFceW9k0K0O/GvYsMrjSn
# I8eazdDtE0TCrvNctSj7FylZVtIMF0jwWrbvtEJOjyQCLvdT2EegN4ycC6WAUFmI
# xkV4EFZyxGCjAEzQ+BIEAxlbNxWjSSuo/eIIPQILZ2wMJuWJS46fcvQ0wplQfbKx
# NMPAzXcMv4MDOLbrhLWitZ0pqZkW0qZFTLj7rHCvyNI1pAy/rkaUNLT8yFA7S4mI
# 5giBnbqVYBdqvnXdqmBpXrEidJfCzkJ0Hb341HUoTm7AwzX510+uvo1U5tRB3CSB
# OzZ3XFzUbOhUu3zfFTPXTYWdmH/KlpOgBNp1jDJpdeMYHDyzYUH8kxRDUMywVHg5
# 7DcAtGlc/RPBmzqjp/NbU9BKrMKQPhZnh7k3LQ+lqLTPu2p5miNmqWVFcDOSTyO7
# 7uDBkc0W6Nwgp32YhMGS7Fi98b5WOIAz2abBC50G0r4miPtQ8mBKjnBl9sDETFYN
# cifeyQZi7YJm8octpWgqWFR4URns26bV+HnS6H+R6iiR/JfVLr68Cs0mWftbUC5Z
# SCaePCPE86ZBtojLe2qEBVQOaMMgD0MIcrB0W+juKfM4E/HTC7DkNkv4AcsIYbfT
# JO8PQ07jIA==
# SIG # End signature block
