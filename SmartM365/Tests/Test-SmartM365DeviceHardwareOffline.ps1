#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic hardware acquisition tests. No module import, API or collector execution.
.VERSION
1.0.0
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Start-Sleep is explicitly mocked to prevent real waits in synthetic tests.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Functions manipulate in-memory synthetic fixtures only; no filesystem or API side effects.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$collector=Join-Path $root 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Collector syntax errors.'}
$getter=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-SafeProperty'},$true)
Set-Item -Path 'Function:script:Get-SafeProperty' -Value ([scriptblock]::Create($getter.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
. (Join-Path (Split-Path $collector) 'SmartM365-DeviceHardware.ps1')
$script:checks=0
function Assert-Hardware {param([bool]$Pass,[string]$Message) if(-not $Pass){throw $Message};$script:checks++}
function Assert-HardwareRejection {param([scriptblock]$Body) $failed=$false;try{& $Body | Out-Null}catch{$failed=$true};Assert-Hardware $failed 'Expected rejection.'}
function WriteLog {param($Message,$Level) $script:log=@{Message=$Message;Level=$Level}}
function Start-Sleep {param($Seconds) $script:delays+=@($Seconds)}
function New-SyntheticHardware {param($Id)
    return @{id=$Id;azureADDeviceId='cloud-'+$Id;serialNumber='serial';manufacturer='Vendor';model='Test';totalStorageSpaceInBytes=100L;freeStorageSpaceInBytes=10L;physicalMemoryInBytes=20L}
}
function Invoke-MgGraphRequest {
    param($Method,$Uri,$Body,$ContentType,$ErrorAction)
    Assert-Hardware ($Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/$batch') 'Unexpected API route in mock.'
    Assert-Hardware ($ContentType -eq 'application/json' -and $ErrorAction -eq 'Stop') 'Batch request options changed.'
    $script:batches++
    $requests=@(($Body | ConvertFrom-Json).requests)
    Assert-Hardware ($requests.Count -le 20) 'Batch exceeds 20.'
    $parts=@()
    foreach($request in $requests){
        Assert-Hardware ($request.url -match '\?\$select=.*totalStorageSpaceInBytes.*physicalMemoryInBytes') 'Explicit hardware select missing.'
        $id=[uri]::UnescapeDataString(($request.url -split '/')[3].Split('?')[0])
        $status=if($script:mode -in @('throttle','fallback','fail')){503}else{200}
        $parts+=@{id=$request.id;status=$status;body=(New-SyntheticHardware $id);headers=@{'Retry-After'='19'}}
    }
    if($script:mode -eq 'throw'){throw 'Synthetic envelope failure'}
    if($script:mode -eq 'missing'){$parts=@()}
    if($script:mode -eq 'duplicate'){$parts+=@($parts[0])}
    if($script:mode -eq 'wrong'){$parts[0].body.id='foreign'}
    return @{responses=$parts}
}
function Get-MgDeviceManagementManagedDevice {
    param($ManagedDeviceId,$Property,$ErrorAction)
    $script:units++
    Assert-Hardware ($ErrorAction -eq 'Stop') 'SDK error boundary missing.'
    Assert-Hardware ($Property -match 'physicalMemoryInBytes' -and $Property -match 'totalStorageSpaceInBytes') 'Unit select missing.'
    if($script:mode -eq 'fail'){throw 'Synthetic permanent failure'}
    return New-SyntheticHardware $ManagedDeviceId
}
function Reset-HardwareMock {param($Mode) $script:mode=$Mode;$script:batches=0;$script:units=0;$script:delays=@()}

$native=New-SyntheticHardware 'd1'
$row=ConvertTo-SmartM365HardwareRecord -ManagedDeviceId d1 -Device $native
Assert-Hardware ($row.CollectionStatus -eq 'Collected' -and $row.physicalMemoryInBytes -eq 20) 'Detailed values missing.'
Assert-Hardware ($row.physicalMemoryInBytesStatus -eq 'Reported') 'Reported status missing.'
Assert-Hardware ($row.CollectedAtUtc -match 'Z$|\+00:00$') 'Acquisition date not UTC.'
$native.physicalMemoryInBytes=0
$row=ConvertTo-SmartM365HardwareRecord -ManagedDeviceId d1 -Device $native
Assert-Hardware ($row.physicalMemoryInBytesStatus -eq 'ZeroReported') 'Zero status lost.'
$native.Remove('physicalMemoryInBytes')
$row=ConvertTo-SmartM365HardwareRecord -ManagedDeviceId d1 -Device $native
Assert-Hardware ($row.physicalMemoryInBytes -eq '' -and $row.physicalMemoryInBytesStatus -eq 'Missing') 'Missing became zero.'
foreach($bad in @('-1','1.2','NaN','9223372036854775808')){
    $native.physicalMemoryInBytes=$bad
    Assert-HardwareRejection {ConvertTo-SmartM365HardwareRecord -ManagedDeviceId d1 -Device $native}
}
$native=New-SyntheticHardware 'foreign'
Assert-HardwareRejection {ConvertTo-SmartM365HardwareRecord -ManagedDeviceId d1 -Device $native}
$native=New-SyntheticHardware 'd1';$native.freeStorageSpaceInBytes=101
Assert-HardwareRejection {ConvertTo-SmartM365HardwareRecord -ManagedDeviceId d1 -Device $native}
Reset-HardwareMock 'ok'
$map=Get-SmartM365ManagedDeviceHardware -ManagedDeviceIds @()
Assert-Hardware ($map.Count -eq 0 -and $script:batches -eq 0) 'Empty fleet invoked a request.'
$map=Get-SmartM365ManagedDeviceHardware -ManagedDeviceIds @(1..21 | ForEach-Object {'d'+$_})
Assert-Hardware ($map.Count -eq 21 -and $script:batches -eq 2 -and $script:units -eq 0) 'Batch fleet coverage differs.'
foreach($mode in @('fallback','missing','duplicate','wrong','throw')){
    Reset-HardwareMock $mode
    $map=Get-SmartM365ManagedDeviceHardware -ManagedDeviceIds @('d1') -MaxAttempts 1
    Assert-Hardware ($map.Count -eq 1 -and $map.d1.CollectionStatus -eq 'Collected' -and $script:units -eq 1) ('SDK fallback not verified: '+$mode)
}
Reset-HardwareMock 'fail'
$map=Get-SmartM365ManagedDeviceHardware -ManagedDeviceIds @('d1') -MaxAttempts 1
Assert-Hardware ($map.d1.CollectionStatus -eq 'Failed' -and -not $map.d1.PSObject.Properties['totalStorageSpaceInBytes']) 'Failed acquisition fabricated values.'
Reset-HardwareMock 'throttle'
$map=Get-SmartM365ManagedDeviceHardware -ManagedDeviceIds @('d1') -MaxAttempts 2
Assert-Hardware ($script:batches -eq 2 -and $script:delays[0] -eq 19 -and $script:units -eq 1) 'Retry-After or bounded fallback differs.'
[pscustomobject]@{Status='Passed';Checks=$script:checks;APIsExecuted=0;CollectorsExecuted=0}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDkGFWXzBoLJ2zu
# yu9x+xKakAsqXiTIqBrr25nZbtay1qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEIh9n72CIa1tIjDdtYxvkwF2oenekOII5byxS1n2AdiMA0GCSqG
# SIb3DQEBAQUABIIBgHTGhT/vxR2FnD5BvaPuH5iHOO7KuQP6pdnzbpCMxRQcXQ+7
# TGx6uGSOs5WgvzYbL7jqbDLhK7REBBmFyNBqB/souIpP1HPN6rbyLS84bzy4eW5P
# y2MPxryeEFg+SvRgZF7TSF2bgLpzd0Md4q6UQnxb66eDRxrytcfnr2Jp4OEQeMF3
# P0TyEBElPQv+1O9WN8rNMa2PmtQLp/d67eER9JzV1CMpbs/RfF6G8WZWpmLaE8zO
# GNuwCOyNOlfDz3/hxKbhkzGg1D9HpFO63+7E1qDINoCQUwqTCloba1uo3Eqn6bMD
# 2snLMwnu0beCgxbGKvvZVZEepKy9Vrp6kiwOPSkgXHT9dT7bOtaHkL2S09DlWmzm
# lDo+X7GOV6CRftxfHqP8RIWpeaxjf1Y6WXuwaa9eGX/CeVHKMI6cAQzOCsa/H3pc
# sNALxeM6rSNKGEZ+Q7fHRbQ0GM2NvwC4wwG5ybz5wAOLaZWiRHnfmtHi62IAulEe
# tCoBUPuqeTbN5PKoO6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDIxOTQx
# MDlaMC8GCSqGSIb3DQEJBDEiBCAbMI8vrGg6yptOAvMpsYueB6xcvXE2Fru1E4WV
# GYPB7zANBgkqhkiG9w0BAQEFAASCAgBadcnKPgMNVj40ZCzN00YpuzvxCf6qL3uF
# +agvykcYa8uJbceRpS281NZjIv16eYNx3B7BZQMqh5Rzev+9JY49ZvB31BaaSZz6
# xCCmAz374XCeFpegOvYJiYmQOk+ahni3e2wOYxvLM1QpsKxPLnr4mlkszJBElcpp
# hs8mljOz88LQN5koX3vEXUrT1ley5pDyUG9FtSVASGeG1iA8Wlno/t64I4VVn+ij
# RAqif+WOO4+TDtKD0+2bqut+Q5zu7lCuiwow28+0YPB9IFSLK39/8DzI0qkhj0Je
# 4+3Z8YrBe2si/S5zvmt3yR/S3usX5wPd6kFwtx5yAXz1drhgfcBYR6vEDpDaYGQs
# gCAPZEYEXmNh60F5KqUSlGFAO5BWGgAq1kCQsNWvnzmgMmNVCBIXbjoOw+V2xxTL
# QpSBTWLv4b0wEOR0o53QSjppYleEUryYyLtalvn8kK4HOH4ia3cQrTfoH5p3PPj+
# 9RGNqujWLZPTpodDf43dlWP/0LDS79VbBqLkDMTOWX/IQI3yMAHbTZxRbMNjG+Dc
# TxhE5AQsPC4u3eFs1VGlhN5fBGKGNzMqH5cmDlNKyq7O3anf8qETsoUta4fILMLc
# VCztJGzMMgHApUEwr/+bmiNTcv76ctNiw2gmLl2SzgFv9jli/RWxUV/r3206EDfc
# ezo/WzPayQ==
# SIG # End signature block
