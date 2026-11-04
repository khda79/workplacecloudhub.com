<#
.SYNOPSIS
    Test the PowerShell 7 to Windows PowerShell identity bridge with synthetic data.
.VERSION
    1.0.2
#>
#Requires -Version 7.4
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath $PSScriptRoot).ProviderPath
$scratch=Join-Path $root ('.identity-bridge-test-'+[guid]::NewGuid().ToString('N'))
if(-not $scratch.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe test path.'}
$oldPath=$env:PSModulePath
$bridge=$null
try {
    $module=Join-Path $scratch 'PnP.PowerShell'
    New-Item -ItemType Directory -Path $module -Force|Out-Null
    @'
@{RootModule='PnP.PowerShell.psm1';ModuleVersion='99.0.0';GUID='a481ac4f-9521-4708-a0b6-812b24fab22b';FunctionsToExport='*'}
'@ | Set-Content -LiteralPath (Join-Path $module 'PnP.PowerShell.psd1') -Encoding utf8
    @'
function Get-PnPList { [CmdletBinding()]param($Identity,$Includes,$Connection) [pscustomobject]@{Title='Documents';BaseType='DocumentLibrary';Id=[guid]'c285d922-b71a-4b41-b546-76a25c44a173';RootFolder=[pscustomobject]@{ServerRelativeUrl='/sites/source/Shared Documents'}} }
function Get-PnPFile { [CmdletBinding()]param($Url,[switch]$AsListItem,$Connection) @{Id=8;FileRef='/sites/source/Shared Documents/Folder/file.txt';ParentList=[pscustomobject]@{Id=[guid]'c285d922-b71a-4b41-b546-76a25c44a173'}} }
function Get-PnPProperty { [CmdletBinding()]param($ClientObject,$Property,$Connection) }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath (Join-Path $module 'PnP.PowerShell.psm1') -Encoding utf8
    $env:PSModulePath=$scratch+[IO.Path]::PathSeparator+$oldPath
    . (Join-Path $root '../Scripts/Diagnostics/SmartM365-SharePointMigration-SpoIdentityBridge.ps1')
    $bridge=Start-SmartM365SpoIdentityBridge -Connection ([pscustomobject]@{Url='https://tenant.sharepoint.com/sites/source'})
    $client=Join-Path $scratch 'client.ps1'
    @'
param($Helper,$PipeName)
$ErrorActionPreference='Stop'
. $Helper
$args=@{PipeName=$PipeName;SiteUrl='https://tenant.sharepoint.com/sites/source';ListName='Documents';SourceFilePath='Folder/file.txt';ExpectedItemId=8}
$result=Get-SmartM365BridgedSpoSourceIdentity @args
if($result.SourceItemId -ne 8){throw 'Valid identity was rejected.'}
$bad=$args.Clone();$bad.ExpectedItemId=9
try {Get-SmartM365BridgedSpoSourceIdentity @bad|Out-Null;throw 'Wrong ID accepted.'} catch {if($_.Exception.Message -notmatch 'verification failed'){throw}}
$bad=$args.Clone();$bad.SiteUrl='https://other.sharepoint.com/sites/source'
try {Get-SmartM365BridgedSpoSourceIdentity @bad|Out-Null;throw 'Wrong site accepted.'} catch {if($_.Exception.Message -notmatch 'verification failed'){throw}}
Write-Output 'PASS: bridge accepts exact identity, rejects incorrect ID and site, and reuses its connection.'
'@ | Set-Content -LiteralPath $client -Encoding utf8
    $helper=Join-Path $root '../Scripts/Diagnostics/SmartM365-SharePointMigration-SpoIdentityBridgeClient.ps1'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $client -Helper $helper -PipeName $bridge.PipeName
    if($LASTEXITCODE -ne 0){throw 'Windows PowerShell bridge client failed.'}
    $sharegate=Join-Path $scratch 'ShareGate'
    New-Item -ItemType Directory -Path $sharegate -Force|Out-Null
    @'
@{RootModule='ShareGate.psm1';ModuleVersion='99.0.0';GUID='b4be4596-bf51-4d18-83c0-d022d6995b19';FunctionsToExport='*'}
'@ | Set-Content -LiteralPath (Join-Path $sharegate 'ShareGate.psd1') -Encoding utf8
    @'
function Connect-Site { [CmdletBinding()]param([string]$Url,[switch]$Browser,[switch]$ModernAuth) [pscustomobject]@{Address=$Url} }
function Get-List { [CmdletBinding()]param($Site,[string]$Name) [pscustomobject]@{Title=$Name;Site=$Site} }
function Get-Folder { [CmdletBinding()]param($List,[string]$Path) [pscustomobject]@{Path=$Path} }
function Get-File { [CmdletBinding()]param($List,[string]$Path) if($Path -eq 'file.txt'){throw 'not found'};[pscustomobject]@{Address=($List.Site.Address+'/'+$List.Title+'/'+$Path)} }
function New-CopySettings { [CmdletBinding()]param([string]$OnContentItemExists) [pscustomobject]@{Mode=$OnContentItemExists} }
function Copy-Content { [CmdletBinding()]param($SourceList,$DestinationList,$SourceItemId,$DestinationFolder,$CopySettings,$TaskName) Set-Content -LiteralPath $env:SMARTM365_BRIDGE_TEST_COPY -Value 'called';throw 'COPY_REACHED' }
function Export-Report { [CmdletBinding()]param($CopyResult,[string]$Path) }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath (Join-Path $sharegate 'ShareGate.psm1') -Encoding utf8
    $integration=Join-Path $scratch 'integration.ps1'
    @'
param($Qualification,$Scratch,$PipeName)
$ErrorActionPreference='Stop'
$zone=@([TimeZoneInfo]::GetSystemTimeZones()|Where-Object {$hour=[TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow,$_).Hour;$hour -ge 10 -and $hour -le 14}|Select-Object -First 1)[0]
if(-not $zone){throw 'No safe test timezone.'}
foreach($id in @(8,9)){
 $project=Join-Path $Scratch ('project-'+$id);$analysis=Join-Path $project 'ShareGate\Diagnostics\Analysis'
 New-Item -ItemType Directory -Path $analysis -Force|Out-Null
 $row=[pscustomobject]@{SessionId='261103-1';RuleId='SG-ACCESS-SOURCE';State='To fix';AccessSide='Source';ObjectType='File';ItemKey=('file-'+$id);SourceItemId=$id;SourceUrl='https://tenant.sharepoint.com/sites/source';SourceList='Documents';DestinationUrl='https://tenant.sharepoint.com/sites/target';DestinationList='Documents';'Raw: Source path'='Folder/file.txt';'Raw: Destination path'='Folder/file.txt'}
 $csv=Join-Path $analysis 'ClassifiedRows.csv';$row|Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8
 $hash=(Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash
 $env:SMARTM365_BRIDGE_TEST_COPY=Join-Path $project 'copy-called.txt';$env:SMARTM365_BRIDGE_TEST_ID=[string]$id
 function global:Read-Host { param($Prompt) 'QUALIFY ONE ITEM '+$env:SMARTM365_BRIDGE_TEST_ID+' v1.0.7' }
 $errorText=''
 try { & $Qualification -ProjectRoot $project -AnalysisDirectory $analysis -SessionId '261103-1' -SourceItemId $id -FarmTimeZoneId $zone.Id -ExpectedAnalysisHash $hash -SpoIdentityPipe $PipeName -Run -ConfirmQualification | Out-Null }
 catch { $errorText=$_.Exception.Message }
 if($id -eq 8){if($errorText -notmatch 'COPY_REACHED' -or -not (Test-Path -LiteralPath $env:SMARTM365_BRIDGE_TEST_COPY)){throw ('Correct ID did not reach simulated copy: '+$errorText)}}
 else {if($errorText -notmatch 'verification failed' -or (Test-Path -LiteralPath $env:SMARTM365_BRIDGE_TEST_COPY)){throw ('Wrong ID was not blocked before simulated copy: '+$errorText)}}
}
Write-Output 'PASS: qualification reaches mocked copy only for the verified source ID.'
'@ | Set-Content -LiteralPath $integration -Encoding utf8
    $qualification=Join-Path $root '../Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGatePathQualification.ps1'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $integration -Qualification $qualification -Scratch $scratch -PipeName $bridge.PipeName
    if($LASTEXITCODE -ne 0){throw 'Simulated ShareGate identity integration failed.'}
}
finally {
    if($bridge){Stop-SmartM365SpoIdentityBridge -Bridge $bridge}
    $env:PSModulePath=$oldPath
    if(Test-Path -LiteralPath $scratch){Remove-Item -LiteralPath $scratch -Recurse -Force}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDN3Ur840LjsbSp
# V+vRrr6GjlfspLgyAWI7zIYmJOPC0qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJXSGj8ywJJIi4dKTftL21+PAp4dRpcW6BUdst79oUGqMA0GCSqG
# SIb3DQEBAQUABIIBgHPl66SjkmL5UgXLwC0RAaeAMH90bZm3ueLir65PeYiLf7Sz
# VoOqEbId1tcR8f5sddGpEnnr+0DzdPfY53R5QuVe/kwUEbjAWrcBgsV9/w6pg7+A
# lnSKt3oeucwnRhwBz6i9t8roXNv99bLVl80WEyBNj4vkoqwAudCNcgnJuHk9VEsx
# xTryReYWdWv8nWhmKF7MuYqwvoqEIz3DMK17W2AZ5fZaRkwd3FBh+RRBnQgprARs
# 67i+4uG0dxO7V1US4sYoBtdTqy6Vyo+/R2B53psvkUEOsyzziNhG6Mzh8pX5tqhq
# uEdDolhQf04eQMqLChEGvrWdWl6/o94hvhD5RkxVEa7Ycwrxq0A+t7PUT4O1Eahy
# U+91RFTLdlNcjH+WpMnobe3233qjVb5u+BciCqKxDnUhv9o1k2/5laVc+GNNOASG
# Fn2jtu0sopeTq4Ul/KQNDo1YE3gWkdSG/03WddRL4+qHe5Rczp1ElZ5/KdeimMYa
# AujWMppLQSw3vw5+HqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDQyMjM4
# MzNaMC8GCSqGSIb3DQEJBDEiBCD3NZkKg92BMUj22/q8e3TAgQKL94eMrXxEjp+Q
# 8bNYLDANBgkqhkiG9w0BAQEFAASCAgCIYl81Actu4a78rR3FNpxoM6OfrptKBuNy
# LNe+92cHdnc61RtZSvKzAxw5Sqsct+Znr2gMydD2+9YBmUQCnwMNhLT3uTQwMP97
# CSrPOsVOk7wKwCaUBPMRdiOgs8ABKlAuOZI20I18x+sdxMA8pTm8dsQWYantvHnS
# 9Q2VypxTX0c+LFwc+rfKuN4L2Uf4+QeIFmoQAQp2D5quW7h/y1aVU3JBD70ogEiX
# jvebOpyRA3dE9fFBsU4I0yNQzHbQlA0KHqoCPw6CEvgutNMWu12Pn6qRBSR+lViA
# sWI2x0x3Klr8F3H29bm40Cg25BQU3IM1h3Iyx9NUYdhfADN8ljAYuJ4YO28xycQ4
# FcLb40FoKqdo1I0tTX05vDZTJl3w8GXgcwsyYkI/Wr4iiR+4XqvugAeASh+ojjBx
# fNs89XHmynmwsT/3/TTywoO1gdI7wyeygCrjlNEHjuXL24wrMWGdO4W/zOP/WmIF
# 4+R2zYXLCwcKHZANYyChCSQDcCWWiWKiRabokrETORzpa66WaVx90FiRt6gYWXnN
# 7MpRjbu73pQAdxtnok4rci3DHU3BxOXGv2WRFwNVa5HK8efKgjCUei/OsYqVgwOA
# jig0ZZaEHuMmY+wdtTBaHEAMf7oUkFx8zG1fxoIQ599tD7oNrSjjs1kSiqsannkl
# iZL5+jGPwg==
# SIG # End signature block
