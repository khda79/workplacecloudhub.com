#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic bounded identity reconciliation tests. All acquisition and waits are mocked.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$directory=Join-Path $PSScriptRoot '../SmartInventory/M365Inventory/IntuneInventory/Devices'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $directory 'SmartM365-Devices-Inventory.ps1'),[ref]$tokens,[ref]$errors)
$getter=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-SafeProperty'},$true)
Set-Item Function:script:Get-SafeProperty ([scriptblock]::Create($getter.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
. (Join-Path $directory 'SmartM365-DeviceHardware.ps1')
. (Join-Path $directory 'SmartM365-DeviceIdentity.ps1')
$script:checks=0
function Check([bool]$pass,[string]$message){if(-not $pass){throw $message};$script:checks++}
function Reject([scriptblock]$body){$failed=$false;try{& $body | Out-Null}catch{$failed=$true};Check $failed 'Expected blocking acquisition failure.'}
function WriteLog {param($Message,$Level) $script:messages+=@($Message)}
function Start-Sleep {param($Seconds) $script:delays+=@($Seconds)}
function New-Device([string]$Id,[string]$Correlation){
    @{Id=$Id;AzureADDeviceId=$Correlation;DeviceName='Synthetic device';OperatingSystem='Windows';OsVersion='10.0.26100';UserId='fake-user';SerialNumber='serial';Manufacturer='Vendor';Model='Original';totalStorageSpaceInBytes=100L;freeStorageSpaceInBytes=10L;physicalMemoryInBytes=20L}
}
function Get-MgDeviceManagementManagedDevice {
    param($ManagedDeviceId,$ErrorAction)
    $script:managedReads++
    Check ($ManagedDeviceId -eq 'd1' -and $ErrorAction -eq 'Stop') 'Unexpected managed acquisition.'
    if($script:mode -in @('graph-error','404')){throw ('Synthetic '+$script:mode)}
    $result=New-Device $ManagedDeviceId ('attempt-'+$script:managedReads)
    $result.Model='Fresh-'+$script:managedReads
    if($script:mode -eq 'wrong-managed'){$result.Id='wrong'}
    $result
}
function Invoke-MgGraphRequest {
    param($Method,$Uri,$ErrorAction)
    $script:hardwareReads++
    Check ($Method -eq 'GET' -and $Uri -match '/managedDevices/d1\?\$select=id,azureADDeviceId,' -and $ErrorAction -eq 'Stop') 'Hardware re-read must be an independent targeted GET.'
    if($script:mode -eq 'hardware-error'){throw 'Synthetic hardware transport failure'}
    $corr=if($script:convergeAt -eq $script:hardwareReads){'attempt-'+$script:hardwareReads}else{'still-conflicting'}
    $result=New-Device d1 $corr
    if($script:mode -eq 'wrong-hardware'){$result.Id='wrong'}
    if($script:mode -eq 'invalid-hardware'){$result.freeStorageSpaceInBytes=999L}
    $result
}
function Run-Scenario([int]$ConvergeAt=0,[string]$Mode='ok',[switch]$Consistent){
    $script:mode=$Mode;$script:convergeAt=$ConvergeAt;$script:managedReads=0;$script:hardwareReads=0;$script:delays=@();$script:messages=@()
    $script:devices=@((New-Device d1 original),(New-Device d2 matched))
    $script:hardware=@{d1=(ConvertTo-SmartM365HardwareRecord d1 (New-Device d1 $(if($Consistent){'original'}else{'different'})));d2=(ConvertTo-SmartM365HardwareRecord d2 (New-Device d2 matched))}
    Resolve-SmartM365DeviceIdentity -Devices $script:devices -Hardware $script:hardware -CollectedAtUtc '2026-01-01T00:00:00Z' -TenantKey synthetic -RunId fake-run
}
$result=Run-Scenario -Consistent
Check ($script:managedReads -eq 0 -and $script:hardwareReads -eq 0 -and $result.EvidenceRows.Count -eq 0 -and $result.ManagedRows.Count -eq 2) 'Consistent rows must require zero additional acquisition.'
$empty=Resolve-SmartM365DeviceIdentity -Devices @() -Hardware @{} -CollectedAtUtc '2026-01-01T00:00:00Z' -TenantKey synthetic -RunId empty-run
Check ($empty.ManagedRows.Count -eq 0 -and $empty.HardwareRows.Count -eq 0 -and $empty.EvidenceRows.Count -eq 0) 'Empty complete scope fabricated an observation.'
foreach($attempt in 1..3){
    $result=Run-Scenario -ConvergeAt $attempt
    Check ($script:managedReads -eq $attempt -and $script:hardwareReads -eq $attempt) 'Unbounded or missing targeted reads.'
    Check ($result.ConflictIds.Count -eq 0 -and $result.EvidenceRows[0].Resolution -eq 'Converged') 'Convergence missing.'
    $managed=@($result.ManagedRows | Where-Object ManagedDeviceId -eq d1)[0]
    $device=@($result.Devices | Where-Object Id -eq d1)[0]
    Check ($managed.Model -eq "Fresh-$attempt" -and $device.Model -eq "Fresh-$attempt" -and $managed.AzureADDeviceId -eq "attempt-$attempt") 'Full managed and Windows observation was not replaced.'
    Check ($managed.CollectedAtUtc -ne '2026-01-01T00:00:00Z' -and $result.HardwareRows[0].azureADDeviceId -eq $managed.AzureADDeviceId) 'Independent acquisition dates or hardware replacement lost.'
    Check (($script:delays -join ',') -eq (@(2,4 | Select-Object -First ($attempt-1)) -join ',')) 'Expected only 2 then 4 second pauses.'
    $doc=$result.EvidenceRows[0].EvidenceJson | ConvertFrom-Json
    Check ($doc.Observations.Count -eq ($attempt+1) -and $doc.Observations[0].Managed.AzureADDeviceId -eq 'original') 'Initial observations or complete attempt history lost.'
}
$result=Run-Scenario
Check ($result.ConflictIds.ContainsKey('d1') -and $result.ManagedRows.Count -eq 2 -and $result.HardwareRows.Count -eq 2) 'Persistent conflict lost an identity.'
Check ($result.ManagedRows[0].AzureADDeviceId -eq 'original' -and $result.HardwareRows[0].azureADDeviceId -eq 'different') 'Quarantine rewrote contested observations.'
Check ($script:managedReads -eq 3 -and $script:hardwareReads -eq 3 -and ($script:delays -join ',') -eq '2,4') 'Persistent conflict is not bounded.'
Check (($result.EvidenceRows[0].EvidenceJson | ConvertFrom-Json).Observations.Count -eq 4) 'Quarantine lacks all observations.'
foreach($mode in @('graph-error','404','hardware-error','wrong-managed','wrong-hardware','invalid-hardware')){Reject {Run-Scenario -Mode $mode}}
Reject {Resolve-SmartM365DeviceIdentity -Devices @($script:devices[0],$script:devices[0]) -Hardware $script:hardware -CollectedAtUtc '2026-01-01T00:00:00Z' -TenantKey synthetic -RunId fake-run}
$script:hardware.d1.CollectionStatus='Failed'
Reject {Resolve-SmartM365DeviceIdentity -Devices $script:devices -Hardware $script:hardware -CollectedAtUtc '2026-01-01T00:00:00Z' -TenantKey synthetic -RunId fake-run}
[pscustomobject]@{Status='Passed';Checks=$script:checks;LiveAPIsExecuted=0;CollectorsExecuted=0}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCpEtq3US4dBHi2
# 0BCdnCSbdx0s5OevBOuFUBeN3OtU2qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIISsok8aoRmHXjvMo2i7porbgI0fpMrH1S8XuQD9SNZDMA0GCSqG
# SIb3DQEBAQUABIIBgKFjU/I+jtNa9RE0yR6hkAMILqlhHVsbgEEtPxK2VId1QrED
# mKcRsWYTMRi3INE/EDsn5SZGR32tEAjVIoi9XVEsiUwnkFlutRQcHwmO6Nl6ai+y
# B18A5TE6CPPVLryOwDG600iNjU2Co0TLTjEKXOmpx6OlVS7VTgDB+pJNwL6wM2V+
# 1almJqiNWKs2p3hzuu8XscsMd2H0zq/t2WcLNFc72lZSKdX5nrmcMEuCKlw6NhUV
# WGgbqUBi21GvOsLzwdvNWK6cOxqttrA7LrOiNBQ43vI/afQBM5Bxi14ZRZPoL+l0
# 1lSTXXsbfPQtFOE0nu6qitk7VgqwuMigdt+gddw80I5UfZqwWcuE+lIQvjMLJrDA
# L+Mb6FaOIBRu0zKJpV7b2TVaRxRtmMDhQKAFhCj50DTDZgZErVyJfLmDYrIgmXn3
# LP4FgDA6hFmVMMgs2Rx6IJxo11oMyGmpKZ5YoEczA0pTymywWv3QvCFC+UoJ2qM8
# jWdWnjuNLaU3EjdBW6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMzIz
# NTJaMC8GCSqGSIb3DQEJBDEiBCAg2xut9reR91hw0tsUr2X4TLip+SnyD5bK5D2l
# Hpzz+DANBgkqhkiG9w0BAQEFAASCAgCbayhN4ta0/uAHPNZ16Jy69Iv8DvFzDk+a
# I9eiVNMeygT+uC5YNzRYFhCnJ5G7CWwDz569T2/s0E3DR3CfNsPi/NKNflD8Uydx
# eeuMJaiPA3l+4ZpRKk9SQbJnuGW8cX2YMX1dOtdyY4SEhb1Da3Ytx2lPcgSLIMSB
# yHFtI7+y0ZClKfqWlnB9LG9WVzWKgYuIY/vU2Lw0s9tfm+eTko6yHxCh6KrRH/Xk
# 5c/H9IGlmxmMdlq9qiPiiBmepLLeuXXjej6POUs3jjmeTzhRdmxAg8iyKZMggdOK
# UAQAWH+WR6lFdstsmB+CeQ4+iCB1l4yIuF956h4d/BoVY1mcRh3bHw2oyg7q0DZp
# sPaqe03BEoA4e85TCRRqx2LoJUhOGuvw5NY2xe2KPKK1iv2lBj25yqi3ggIWqkFr
# c89TgfowxMGNLKETAOvbhsv+Wq1DFQFZr7hOMvzkDcRbm7oZatUPDbcmLSEorjYT
# yy+YK8AH6J9qx1bgjkLqcAGx2P+Ca4rTEJOBaFzdXGmzwzXT1fs5CnM9+LHxjrvK
# BVUZ0a/V8ck4GYf4V1mxB/AFU70fpg2r6YwrfoO5hRjjhRswdAaPViH3/zx2ZflB
# 5RlBzHQ3nPrPr2hzitOpSrJUxwE1dR10HbDN0AIc27D0wEFCHCwBBNXbnWanxkqw
# yvuyg4LHbg==
# SIG # End signature block
