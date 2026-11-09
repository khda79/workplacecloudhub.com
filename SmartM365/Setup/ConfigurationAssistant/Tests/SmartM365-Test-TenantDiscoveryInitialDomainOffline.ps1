$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$temp=Join-Path ([IO.Path]::GetTempPath()) ('si-tenant-fixture-'+[guid]::NewGuid().ToString('N'))
$module=Join-Path $temp 'Microsoft.Graph.Authentication'
$null=New-Item -ItemType Directory -Path $module -Force
try {
    @'
function Connect-MgGraph { param($Scopes,[switch]$UseDeviceAuthentication,$ContextScope,[switch]$NoWelcome) if ($UseDeviceAuthentication) {throw 'Device code flow was requested.'} }
function Get-MgContext { if ($env:SI_FIXTURE_MISSING -eq '1') {return @{TenantId='11111111-2222-3333-4444-555555555555';Account='user@synthetic.invalid';Scopes=@('User.Read')}}; if ($env:SI_FIXTURE_BAD -eq '1') {return @{TenantId='invalid';Account='user@synthetic.invalid'}}; return @{TenantId='11111111-2222-3333-4444-555555555555';Account='user@synthetic.invalid';Scopes=@('User.Read','Policy.Read.All','Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All','Channel.Create','Channel.ReadBasic.All','Directory.Read.All','Group.ReadWrite.All','RoleManagement.ReadWrite.Directory','Sites.FullControl.All')} }
function Invoke-MgGraphRequest { param($Method,$Uri) if ($Uri -like '*/v1.0/domains*') {
 if ($env:SI_FIXTURE_DOMAIN -eq 'bad-next') {return @{value=@(@{id='synthetic.onmicrosoft.com';isInitial=$true;isVerified=$true});'@odata.nextLink'='https://evil.invalid/v1.0/domains'}}
 if ($env:SI_FIXTURE_DOMAIN -eq 'pages' -and $Uri -notlike '*skiptoken*') {return @{value=@(@{id='other.example';isInitial=$false;isVerified=$true});'@odata.nextLink'='https://graph.microsoft.com/v1.0/domains?$skiptoken=2'}}
 if ($env:SI_FIXTURE_DOMAIN -eq 'error') {throw 'Synthetic domain failure.'}
 if ($env:SI_FIXTURE_DOMAIN -eq 'duplicate') {return @{value=@(@{id='one.onmicrosoft.com';isInitial=$true;isVerified=$true},@{id='two.onmicrosoft.com';isInitial=$true;isVerified=$true})}}
 if ($env:SI_FIXTURE_DOMAIN -eq 'unverified') {return @{value=@(@{id='synthetic.onmicrosoft.com';isInitial=$true;isVerified=$false})}}
 return @{value=@(@{id='synthetic.onmicrosoft.com';isInitial=$true;isVerified=$true},@{id='other.example';isInitial=$false;isVerified=$true})}
 }; if ($Uri -like '*transitiveMemberOf*') {return @{value=@(@{'@odata.type'='#microsoft.graph.directoryRole';roleTemplateId='62e90394-69f5-4237-9190-012177145e10'})}};if ($Uri -like '*groupSettings*') {return @{value=@()}};if ($Uri -like '*/sites/root?*') {if ($env:SI_FIXTURE_NOSITE -eq '1') {throw 'Synthetic site failure.'};return @{webUrl='https://synthetic.sharepoint.invalid/'}}; if ($env:SI_FIXTURE_NOMAIL -eq '1') {throw 'Synthetic Graph failure.'}; if ($Uri -notlike '*/me?*') {throw 'Unexpected Graph URI.'}; return @{mail='admin@synthetic.invalid'} }
function Disconnect-MgGraph { }
Export-ModuleMember -Function Connect-MgGraph,Get-MgContext,Invoke-MgGraphRequest,Disconnect-MgGraph
'@ | Set-Content -LiteralPath (Join-Path $module 'Microsoft.Graph.Authentication.psm1') -Encoding utf8
    $old=$env:PSModulePath;$env:PSModulePath=$temp
    $result=Join-Path $temp 'result.json'
    $source=Get-Content -LiteralPath (Join-Path $root 'Lib/SmartM365.TenantDiscoveryWorker.InitialDomain.ps1') -Raw
    $ui=Get-Content -LiteralPath (Join-Path $root 'Lib/SmartM365.TenantSignInUi.InitialDomain.ps1') -Raw
    if ($ui -notmatch '-WindowStyle Hidden' -or $ui -match '-WindowStyle Normal') {throw 'The sign-in worker must not open a PowerShell console.'}
    if ($source -match 'UseDevice(?:Code|Authentication)') {throw 'Device-code authentication must not be used.'}
    $mockPath=(Join-Path $module 'Microsoft.Graph.Authentication.psm1').Replace("'","''")
    $source=$source.Replace('Import-Module Microsoft.Graph.Authentication -ErrorAction Stop',("Import-Module '"+$mockPath+"' -ErrorAction Stop"))
    Copy-Item -LiteralPath (Join-Path $root 'Lib/SmartM365.GraphSetupReadiness.ps1') -Destination (Join-Path $temp 'SmartM365.GraphSetupReadiness.ps1');$worker=Join-Path $temp 'Worker.ps1';[IO.File]::WriteAllText($worker,$source,[Text.UTF8Encoding]::new($false))
    & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $worker -ResultPath $result
    if ($LASTEXITCODE -ne 0) {throw 'Synthetic detection failed.'}
    $value=Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
    if ($value.TenantId -ne '11111111-2222-3333-4444-555555555555' -or $value.Account -ne 'user@synthetic.invalid' -or $value.Mail -ne 'admin@synthetic.invalid' -or $value.InitialDomain -ne 'synthetic.onmicrosoft.com' -or $value.SiteHost -ne 'synthetic.sharepoint.invalid' -or -not $value.Readiness.Ready -or @($value.MissingScopes).Count -ne 0) {throw 'Unexpected synthetic context.'}
    Remove-Item -LiteralPath $result
    $env:SI_FIXTURE_NOMAIL='1'
    & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $worker -ResultPath $result
    $noMail=Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0 -or $noMail.Mail -ne '') {throw 'Unavailable mail was incorrectly inferred or sign-in was blocked.'}
    Remove-Item -LiteralPath $result;Remove-Item Env:SI_FIXTURE_NOMAIL
    $env:SI_FIXTURE_NOSITE='1'
    & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $worker -ResultPath $result
    $noSite=Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0 -or $noSite.SiteHost -ne '') {throw 'Unavailable site host must not be invented.'}
    Remove-Item -LiteralPath $result;Remove-Item Env:SI_FIXTURE_NOSITE
    $env:SI_FIXTURE_MISSING='1'
    & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $worker -ResultPath $result
    $missing=Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0 -or 'Group.ReadWrite.All' -notin @($missing.MissingScopes)) {throw 'Missing setup permission was not reported.'}
    Remove-Item -LiteralPath $result;Remove-Item Env:SI_FIXTURE_MISSING
    foreach ($scenario in @('error','duplicate','unverified','bad-next')) {
        $env:SI_FIXTURE_DOMAIN=$scenario
        & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $worker -ResultPath $result
        $missingDomain=Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
        if ($LASTEXITCODE -ne 0 -or $missingDomain.InitialDomain -ne '') {throw ('An unavailable or ambiguous initial domain was accepted: '+$scenario)}
        Remove-Item -LiteralPath $result
    }
    $env:SI_FIXTURE_DOMAIN='pages'
    & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $worker -ResultPath $result
    $paged=Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0 -or $paged.InitialDomain -ne 'synthetic.onmicrosoft.com') {throw 'A paged initial domain was not detected.'}
    Remove-Item -LiteralPath $result
    Remove-Item Env:SI_FIXTURE_DOMAIN
    $env:SI_FIXTURE_BAD='1'
    & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $worker -ResultPath $result
    if ($LASTEXITCODE -eq 0 -or (Test-Path -LiteralPath $result)) {throw 'Invalid context was accepted.'}
    'Offline tenant discovery fixtures passed.'
} finally {$env:PSModulePath=$old;Remove-Item Env:SI_FIXTURE_BAD,Env:SI_FIXTURE_NOMAIL,Env:SI_FIXTURE_NOSITE,Env:SI_FIXTURE_MISSING,Env:SI_FIXTURE_DOMAIN -ErrorAction SilentlyContinue;Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC9DQSFdpBnvZdx
# WCSanGv4O8UHmEMJQeOqdXh2VbMCW6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOii3Q9Xh5yOWe58Nmu+yIdPoDUAAZskoExDqcaiVeCFMA0GCSqG
# SIb3DQEBAQUABIIBgKfdKoUFfJ551SMHotJdXyJdqm5Ds6i7GcieTMAorN21cf60
# pnx440eYOia6oUBoYRygAYOR1y9PYULz+QJNbjjdf+aFsRvAzONfB6fb9LcfLswK
# 4WBXRyHF9wM8d+rusexPemxxBwkLTlS6sHufY+zApH8SkcP5Dwhc6y2bFykqbFMW
# Ba/drboxycTEk64LC97yXbrJEIAZ59uGpajy//HW3PKZ/Q1WqS2buzh5x40ZdOGD
# NroFN0tuAJ4QSk6CqRGRh531G39/2A6GVJZd/oEV+Zb9wB60ihJZQSjyDRWL55/2
# /6KSh4Z1AiP6TmskaJWY8kf7cOBUv/6ze2okIIifxLqMutqPw/htrBDN6sYIa8oQ
# 5pzVbf5Cf4XQl894XsKzr/HGshlAUa++NhdsrrDwppb4qitopKaKtA+OiydoutOE
# fM9TO/CKQ01Ck4cOSt9GemjT7RjXs318uK8SZsWPEvgo7CUcA0BmmZN0f1ZvqnOP
# GxDCasRl9mJz2W3Ey6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MTFaMC8GCSqGSIb3DQEJBDEiBCCh6ycwVU2coqlJcSdeFc6ByKEqCjq02/sO1r4k
# 2bcPrzANBgkqhkiG9w0BAQEFAASCAgApFJulyG/N48V5YvdtaQSxjRCuQ+VKmXZK
# 3qZF2LmwfgMpInDiD3Pn8p4XJVvC2Tf2bK10euasTWewbUnYHv0Gb0nEQznJu+AE
# cm7W4r4ZPz9sdHxdHF48m7vkzowC/Qr9MUWB0vQ1wI6dVC0RmVwGs4Fu2tX5+f5z
# swQ9SpwuEB+5VZ1E4AFDCExVKcGbYh01/KDG20k3+E5lw55oh9dBsPVmBgZS721Y
# MdEmaTTjyVwiD8rqpvyue1pQVIqvoArGtbuqJjbz1PMTqwfwwJSp7JUKmO9/jwRV
# 6QVGqcUMjIXYBgsm2QilPq5CnqMVfTK9wvY2j2saONyzbiseFKLhYuvmg1WqSc8g
# 4dn9NtNEX+8pfKyPG3EMh2FFp6tdpfl27YUCndJ8fLamxu+hmCf4Io3jutyfnX2W
# 1yip1n/M2WxbThLNr1fhuLjk1vOnCKfP0OY5Xy19RV7fyWHAiq6Ee3lFOvzDQn9z
# u2VLfSDIW38w9F9dqFATSMJBuWMWLMQuIZKTM/hNPZXS54BFdYF39DxO4oTCxe0s
# DjoS/I+1+7La7b/fDFWdP8nMf9CM3HcJa6H6qvEjtMwSQmH0jZ93OfXf3PFZHDTN
# RUtfdqT3p8R6r7c4fMRW/BQRNB+Meq/UrB+9BWS1w5rDe1rMW3jaQVlb6K6G5L1S
# EqasjHAwbg==
# SIG # End signature block
