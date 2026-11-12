<# .VERSION 1.1.0
Isolated optional-license checks. No native diagnostic, authentication or network calls.
#>
$ErrorActionPreference='Stop'
$app=Join-Path (Split-Path $PSScriptRoot -Parent) 'Devices\DeviceRegistrationTool'
. (Join-Path $app 'SmartM365-DeviceRegistration.Licenses.ps1')
$script:checks=0
function Assert-License($Condition,$Message){if(-not $Condition){throw $Message};$script:checks++}
$upn='user01@contoso.example';$sku='00000000-0000-4000-8000-000000000001'
$profile=[pscustomobject]@{userPrincipalName=$upn;assignedLicenses=@([pscustomobject]@{skuId=$sku})}
$before=$profile|ConvertTo-Json -Depth 5
$r=ConvertTo-DrtLicenseResult $profile $upn
Assert-License ($r.Status -eq 'Assigned' -and @($r.SkuIds).Count -eq 1 -and $r.SkuIds[0] -eq $sku) 'Valid assignment'
Assert-License (($profile|ConvertTo-Json -Depth 5) -ceq $before) 'Profile must not be mutated'
Assert-License ((ConvertTo-DrtLicenseResult $profile $upn.ToUpper()).Status -eq 'Assigned') 'Case-insensitive identity'
Assert-License ((ConvertTo-DrtLicenseResult $profile 'other@contoso.example').Status -eq 'Unverifiable') 'Other identity must not be accepted'
Assert-License ((ConvertTo-DrtLicenseResult ([pscustomobject]@{userPrincipalName=$upn;assignedLicenses=@()}) $upn).Status -eq 'None') 'Verified explicit empty assignment'
foreach($p in @($null,[pscustomobject]@{userPrincipalName=$upn},[pscustomobject]@{userPrincipalName=$upn;assignedLicenses=$null},[pscustomobject]@{userPrincipalName=$upn;assignedLicenses=@([pscustomobject]@{skuId='not-a-guid'})},[pscustomobject]@{userPrincipalName=$upn;assignedLicenses=@([pscustomobject]@{skuId=[guid]::Empty})})){
 Assert-License ((ConvertTo-DrtLicenseResult $p $upn).Status -eq 'Unverifiable') 'Incomplete response is not absence'
}
# Shadow module loading and SDK entry points in this isolated test process.
function Import-Module {param($Name,$RequiredVersion,$ErrorAction)}
$script:ConnectCalls=0;$script:ReadCalls=0;$script:DisconnectCalls=0;$script:Failure='';$script:WrongContext=$false
function Disconnect-MgGraph {param($ErrorAction);$script:DisconnectCalls++}
function Connect-MgGraph {
 param($Scopes,$ContextScope,$LoginHint,[switch]$NoWelcome,[switch]$UseDeviceCode,$ClientTimeout,$ErrorAction)
 $script:ConnectCalls++
 Assert-License (@($Scopes).Count -eq 1 -and $Scopes[0] -eq 'User.Read') 'Only User.Read requested'
 Assert-License ($ContextScope -eq 'Process' -and $LoginHint -eq $upn) 'Isolated context and account verification'
 if($script:Failure -in @('cancel','timeout','consent')){throw ('secret-'+$script:Failure)}
}
function Get-MgContext {param($ErrorAction);[pscustomobject]@{AuthType='Delegated';ContextScope='Process';Account=$(if($script:WrongContext){'admin@contoso.example'}else{$upn})}}
function Invoke-MgGraphRequest {
 param($Method,$Uri,$OutputType,$ErrorAction)
 $script:ReadCalls++
 Assert-License ($Method -ceq 'GET' -and $Uri -ceq 'https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,assignedLicenses') 'Only fixed own-profile GET'
 if($script:Failure){throw ('secret-'+$script:Failure)}
 $profile
}
$r=Invoke-DrtLicenseCheck $upn
Assert-License ($r.Status -eq 'Assigned' -and $script:ConnectCalls -eq 1 -and $script:ReadCalls -eq 1 -and $script:DisconnectCalls -eq 2) 'Delegated success and cleanup'
$script:WrongContext=$true;$n=$script:ReadCalls
Assert-License ((Invoke-DrtLicenseCheck $upn).Status -eq 'Unverifiable' -and $script:ReadCalls -eq $n) 'Wrong cached account cannot query'
$script:WrongContext=$false
foreach($failure in '403','401','cancel','timeout','consent','network','429'){
 $script:Failure=$failure;$r=Invoke-DrtLicenseCheck $upn
 Assert-License ($r.Status -eq 'Unverifiable' -and @($r.SkuIds).Count -eq 0) ('Failure cannot be no-license: '+$failure)
 Assert-License (($r|ConvertTo-Json -Depth 5) -notmatch 'secret-') 'Raw SDK errors must not be persisted'
}
$n=$script:ConnectCalls
Assert-License ((Invoke-DrtLicenseCheck '').Reason -eq 'AccountUnavailable' -and $script:ConnectCalls -eq $n) 'Missing account does not authenticate'
function Import-Module {param($Name,$RequiredVersion,$ErrorAction);throw 'Missing SDK'}
Assert-License ((Invoke-DrtLicenseCheck $upn).Reason -eq 'DependencyUnavailable' -and $script:ConnectCalls -eq $n) 'Missing dependency is explicit and no automatic install'
function Import-Module {param($Name,$RequiredVersion,$ErrorAction)}
function Disconnect-MgGraph {param($ErrorAction);throw 'No context to disconnect'}
$script:Failure=''
Assert-License ((Invoke-DrtLicenseCheck $upn).Status -eq 'Assigned') 'Disconnect failure must not discard a valid response'
$exception=[Exception]::new('AADSTS65001 consent required SECRET_TOKEN and private identifiers')
$record=[Management.Automation.ErrorRecord]::new($exception,'MockFailure',[Management.Automation.ErrorCategory]::PermissionDenied,$null)
$r=New-DrtLicenseFailure -ErrorRecord $record -Phase Authentication -Account $upn
Assert-License ($r.Diagnostic.Phase -eq 'Authentication' -and $r.Diagnostic.AadstsCode -eq '65001') 'Safe diagnostic phase and AADSTS code'
Assert-License (($r|ConvertTo-Json -Depth 5) -notmatch 'SECRET_TOKEN|private identifiers|consent required') 'Never persist raw exception text'
$record=[Management.Automation.ErrorRecord]::new([TimeoutException]::new('Authentication timed out due to inactivity'),'MockTimeout',[Management.Automation.ErrorCategory]::OperationTimeout,$null)
Assert-License ((New-DrtLicenseFailure $record Authentication $upn).Reason -eq 'AuthenticationTimeout') 'Authentication timeout is distinguishable from access denial'
$temp=Join-Path ([IO.Path]::GetTempPath()) ('WCH-LicenseWorker-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $temp|Out-Null
try{
 $out=Join-Path $temp 'Result.json';$worker=Join-Path $app 'SmartM365-DeviceRegistration.Licenses.ps1'
 $child=Start-Process powershell.exe -WindowStyle Hidden -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "'+$worker+'" -Check -ExpectedUpn invalid -ResultPath "'+$out+'"') -Wait -PassThru
 Assert-License ($child.ExitCode -eq 0 -and (Test-Path $out)) 'Real hidden worker entry serializes a safe result without SDK or network'
 Assert-License ((Get-Content $out -Raw|ConvertFrom-Json).Reason -eq 'AccountUnavailable') 'Worker entry preserves unavailable account'
 Assert-License ((Get-Content ($out+'.progress.json') -Raw|ConvertFrom-Json).Phase -eq 'Completed') 'Worker phase reaches completion'
}finally{Remove-Item -LiteralPath $temp -Recurse -Force}
$cases=@(
 @('66b55226-6b4f-492c-910c-a3b7a3c9d993','Microsoft 365 F3'),
 @('05e9a617-0261-4cee-bb44-138d3ef5d965','Microsoft 365 E3'),
 @('06ebc4ee-1bb5-47dd-8120-11324bc54e06','Microsoft 365 E5'),
 @('6fd2c87f-b296-42f0-b197-1e91e994b900','Office 365 E3'),
 @('c7df2760-2c81-4ef7-b578-5b5392b571df','Office 365 E5'),
 @('18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e','Microsoft 365 E5 (no Teams)'),
 @('3b555118-da6a-4418-894f-7df1e2096870','Microsoft 365 Business Basic'),
 @('f245ecc8-75af-4f8e-b61f-27d8114de5f3','Microsoft 365 Business Standard'),
 @('cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46','Microsoft 365 Business Premium')
)
foreach($case in $cases){$resolved=@(Resolve-DrtLicenseProducts -SkuIds @($case[0]));Assert-License ($resolved.Count -eq 1 -and $resolved[0].DisplayName -ceq $case[1] -and $resolved[0].CatalogMatch) ('Catalog identity: '+$case[1])}
$unknown=@(Resolve-DrtLicenseProducts -SkuIds @($sku))
Assert-License (-not $unknown[0].CatalogMatch -and $unknown[0].SkuId -eq $sku) 'Unknown product stays unknown with exact identifier'
$missing=@(Resolve-DrtLicenseProducts -SkuIds @($cases[0][0]) -CatalogPath (Join-Path $app 'does-not-exist.json'))
Assert-License (-not $missing[0].CatalogMatch) 'Missing catalog does not invent names'
$multiple=@(Resolve-DrtLicenseProducts -SkuIds @($cases[0][0],$cases[1][0],$sku))
Assert-License ($multiple.Count -eq 3 -and @($multiple|Where-Object CatalogMatch).Count -eq 2) 'Multiple assignments retain every product and unknown entry'
$display=Get-DrtLicenseText (New-DrtLicenseResult -Status Assigned -Account $upn -SkuIds @($cases[5][0]))
Assert-License ($display -match 'Microsoft 365 E5 \(no Teams\)' -and $display -match $cases[5][0]) 'Readable product and technical SKU both retained'
$main=Get-Content (Join-Path $app 'SmartM365-DeviceRegistration-Tool.ps1') -Raw
Assert-License ($main -match 'licenseButton.Add_Click' -and $main -match 'WindowStyle Hidden' -and $main -match 'DispatcherTimer') 'Explicit optional async hidden worker'
Assert-License ($main -match 'CreateNoWindow=\$false' -and $main -match '-NoProfile -STA -WindowStyle Hidden' -and $main -match 'ReadToEndAsync') 'STA worker retains a hidden WAM parent and drains streams in memory'
Assert-License ($main -match 'LicenseExpander.*IsExpanded="False"') 'Optional section collapsed'
Assert-License ($main -notmatch '\$script:lastResult\.License') 'Cloud check does not rewrite local result'
Add-Type -AssemblyName PresentationFramework
$t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $app 'SmartM365-DeviceRegistration-Tool.ps1'),[ref]$t,[ref]$e)
Assert-License (-not $e) 'Main parses'
$x=$ast.FindAll({param($a)$a -is [Management.Automation.Language.AssignmentStatementAst] -and $a.Left.Extent.Text -eq '$xaml'},$true)|Select-Object -First 1
$markup=Invoke-Expression $x.Right.Extent.Text
$w=[Windows.Markup.XamlReader]::Parse($markup)
Assert-License ($w.FindName('LicenseExpander').IsExpanded -eq $false -and $w.FindName('LicenseButton')) 'Actual native WPF optional controls'
'PASS: '+$script:checks+' optional-license assertions'





# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCuFB2QVhMXKeN6
# g7FYZ6Mhb0t6AADjJMBE0Haivo1yMaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMR6rzCCEG/fgD374NhrymJbOqlL+w9QZ/NeS2977MnUMA0GCSqG
# SIb3DQEBAQUABIIBgA5Zs+qPAbeYb80nQyaArCch8g+piVEmLY0d2uMIJT5uWS4P
# nb8I+nolZQpvFNcaPy8rw8ajE7hnBCvDKnTQprzNOV8WRnDLpu0LyBqfaX7Oj/zK
# r9xs69aCpHkLkKuk/MpPkdDkWMn5q9zSGU+qniV/mtwCnhXnNQJejjXvYyCcxdlf
# KSpGXnyUEAdPU9MR76yFiSvtt36++z44g4P6uPMWelP+MWRoiuRVCRiFtpCrU2DL
# Pjke+OhKvK+7k+A5RPSR9VGKLNG0cSb+dfxBvkR0rgIS3+NuMQ+GdC1KT8TMatJP
# sNUH+DfUbuFYZPUuw7RxzOVRoT7saQGBKc7DcxbwvN0VTcj0yhdKiaLaFUDeBurL
# yotvB/KA1e5zQwNFv5cFWqQf5lWdG9haU2Cd9mQ9AR+Z5QvjjBYmYSbyFUmpBM1t
# Oscv7ZHkhVSsmjc/PZv6GaEpJDLcdEQx1lLXKOqRT6sTf1RdQXj8h1W0930ds15U
# t8sCCcT99KhFGsre/6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTIyMjQy
# NDlaMC8GCSqGSIb3DQEJBDEiBCCLmaj/348+WiTtMpbRs1qivqK64g9QYbbv/b5Z
# sw7XjzANBgkqhkiG9w0BAQEFAASCAgBGJ2/eTqcFJctJP77qf7LjPwjQKp33nKEI
# n29QsE6aDsqOvjERh2PZvftkPWR/icFtVPOXQyPbhl2mGxSEsOA7AqingTSXkDU3
# ZB9dnXvIV400KWA/x/rj9N7yxzo0hu/Juyl81Qonan7uLU6unc03IH4bzsSd/czz
# uuQJEz//GSfyS29grnz+ZxvQdiWVulOREflaGT8Ej61HxWnJH4eWbbvoeyNzombz
# BW8TYLOGlXONboqDKpI8l5sRd1r5/rD3JE/KXEe5jNnR3+DPItHu1pGF2Rq2R/89
# 0FIac4Nnsd7Xc7ovF6+/mYqJtV+31otPB8XIgQv6p9OcNhL7ahsHWnnVug7udl85
# 0+HPnjchyrefVXKHuZhEgVA0L3mcgd0BioBLQ0KkeNJhM/XGcdZD7tqRuFjJHEig
# HE2f7jA95y0gbXeirAka+Ov+MtITgU+NIymTmVYYbeSeZ+F0JM3cC7txzTUgH/tk
# gOMe53j/JhQ4dgxpdakcta4e0DwjwAazHd2qbt1k3Gw1FxF8pMuzjjqEJp6LxZTQ
# 1ps3005nMDaYQGnZcrHw8kifvDpFhmXgk4JLuepHAccK502lavsjTLe0CoLKDK2n
# i2PqcdU0anNfqmOVro/H5LkjGymiKB6tfgbpj3DnW6JA26zTT/0PT/HcUEIuuEVI
# lXEvDXzr1g==
# SIG # End signature block
