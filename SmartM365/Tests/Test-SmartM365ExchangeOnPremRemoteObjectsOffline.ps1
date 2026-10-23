#Requires -Version 5.1
<# .SYNOPSIS Client-side Exchange wire-value regression tests. .VERSION 1.0.0 #>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$smartRoot=Split-Path $PSScriptRoot -Parent
. (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-ExchangeOnPremSession.ps1')
$checks=0
function Check([bool]$Condition,[string]$Message) { if (-not $Condition) { throw $Message }; $script:checks++ }
function Wire($Value) { [Management.Automation.PSSerializer]::Deserialize([Management.Automation.PSSerializer]::Serialize($Value,10)) }
$savedCulture=[Threading.Thread]::CurrentThread.CurrentCulture
try {
    foreach ($culture in @('en-US','fr-FR','de-DE')) {
        [Threading.Thread]::CurrentThread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo($culture)
        foreach ($case in @(
            @{Value='1.5 GB (1,610,612,736 bytes)';Bytes=1610612736;MB=1536},
            @{Value='1,5 GB (1.610.612.736 bytes)';Bytes=1610612736;MB=1536},
            @{Value='163.7 MB (171,699,411 bytes)';Bytes=171699411;MB=163},
            @{Value='0 B (0 bytes)';Bytes=0;MB=0},
            @{Value='1 B (1 bytes)';Bytes=1;MB=0},
            @{Value=[long]0;Bytes=0;MB=0},
            @{Value=[pscustomobject]@{IsUnlimited=$false;Value='2 MB (2 097 152 octets)'};Bytes=2097152;MB=2}
        )) {
            $wire=Wire $case.Value
            $result=ConvertFrom-SmartM365ExchangeSize $wire
            Check ($result.Status -eq 'Known' -and $result.Bytes -eq $case.Bytes -and $result.Megabytes -eq $case.MB) ('Wire bytes/culture failed: '+$culture+' / '+[string]$case.Value)
            Check ((ConvertTo-SmartM365ExchangeSizeMB $wire) -eq $case.MB) 'CSV MB conversion failed'
        }
        foreach ($value in @('Unlimited',[pscustomobject]@{IsUnlimited=$true;Value=$null})) {
            Check ((ConvertTo-SmartM365ExchangeSizeMB (Wire $value)) -eq 'Unlimited') 'Unlimited became zero or unknown'
        }
        foreach ($value in @($null,'','N/A','Unknown','1.5 GB','1.5 GB (1,23 bytes)')) {
            Check ((ConvertTo-SmartM365ExchangeSizeMB (Wire $value)) -eq 'N/A') 'Unavailable exact bytes became a measured number'
        }
        foreach ($value in @('Error','Failed','9 GB (999999999999999999999999 bytes)',[long]-1)) {
            Check ((ConvertTo-SmartM365ExchangeSizeMB (Wire $value)) -eq 'Error') 'Error became zero or unknown'
        }
        $guid=[guid]'12345678-1234-5678-9abc-123456789abc'
        foreach ($value in @($guid,$guid.ToString('B'),$guid.ToString('D'))) {
            Check ((ConvertTo-SmartM365ExchangeGuid (Wire $value)) -eq $guid.ToString('D')) 'Wire GUID changed'
            Check ((ConvertTo-SmartM365ExchangeGuid (Wire $value) -Base64) -eq [Convert]::ToBase64String($guid.ToByteArray())) 'Wire ImmutableId changed'
        }
        foreach ($value in @($null,'','invalid',[guid]::Empty)) {
            Check ((ConvertTo-SmartM365ExchangeGuid (Wire $value)) -eq '') 'Invalid GUID became a native identity'
        }
        foreach ($value in @('2026-10-05T09:08:07Z','2026-10-05T11:08:07+02:00')) {
            $iso=ConvertTo-SmartM365ExchangeDateIso (Wire $value)
            Check ($iso -and ([datetimeoffset]$iso).UtcDateTime -eq ([datetimeoffset]'2026-10-05T09:08:07Z').UtcDateTime) 'Explicit wire date offset changed'
        }
        foreach ($value in @($null,'','05/10/2026 09:08:07','2026-10-05T09:08:07','invalid','2026-02-30T09:08:07Z')) {
            Check ((ConvertTo-SmartM365ExchangeDateIso (Wire $value)) -eq '') 'Ambiguous/invalid text gained date proof'
        }
        $native=[datetime]::SpecifyKind([datetime]'2026-10-05T09:08:07',[DateTimeKind]::Unspecified)
        Check ((ConvertTo-SmartM365ExchangeDateIso (Wire $native)) -eq '2026-10-05T09:08:07.0000000') 'Unspecified native date gained a timezone'
    }

    # Exercise the actual remote mailbox projection, not a duplicated test projection.
    $collector=Join-Path $smartRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$errors)
    foreach ($name in @('Get-SmartM365MailboxReportDomain','ConvertTo-SmartM365ExchangeMailboxSemicolonList','ConvertTo-SmartM365ExchangeRemoteMailboxRecord')) {
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        Check ($null -ne $node) ('Projection dependency missing: '+$name)
        . ([scriptblock]::Create($node.Extent.Text))
    }
    $script:ExchangeConnectionMode='Remote'
    $source=Wire ([pscustomobject]@{Guid='12345678-1234-5678-9abc-123456789abc';ExchangeGuid='22345678-1234-5678-9abc-123456789abc';WhenMailboxCreated='2026-10-05T09:08:07Z';DistinguishedName='CN=Synthetic,DC=example,DC=invalid';Identity='Synthetic';PrimarySmtpAddress='synthetic@example.invalid'})
    $record=ConvertTo-SmartM365ExchangeRemoteMailboxRecord -Mailbox $source
    Check ($record.ObjectGuid -eq $source.Guid -and $record.ImmutableId -eq [Convert]::ToBase64String(([guid]$source.Guid).ToByteArray())) 'Actual remote projection lost GUID/ImmutableId'
    Check ($record.WhenMailboxCreated -eq $source.WhenMailboxCreated -and $record.WhenMailboxCreatedIso8601 -match 'Z$|\+00:00$') 'Actual projection lost raw date/ISO'
    $source.WhenMailboxCreated='05/10/2026 09:08:07'
    $record=ConvertTo-SmartM365ExchangeRemoteMailboxRecord -Mailbox $source
    Check ($record.WhenMailboxCreated -eq $source.WhenMailboxCreated -and $record.WhenMailboxCreatedIso8601 -eq '') 'Actual projection promoted display text'

    $statsNode=$ast.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '^if \(\$Stats\)'},$true)
    $quotaNode=$ast.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '\$quotaValue = \$Mbx.ProhibitSendReceiveQuota' -and $n.Extent.Text -match '^if \(\$script:ExchangeConnectionMode'},$true)
    Check ($null -ne $statsNode -and $null -ne $quotaNode) 'Actual local mailbox projections missing'
    foreach ($sizeCase in @(
        @{Input='0 B (0 bytes)';Expected=0},
        @{Input='2 MB (2,097,152 bytes)';Expected=2},
        @{Input='Unlimited';Expected='Unlimited'},
        @{Input=$null;Expected='N/A'},
        @{Input='Error';Expected='Error'}
    )) {
        $Stats=Wire ([pscustomobject]@{LastLogonTime='05/10/2026 09:08:07';TotalItemSize=$sizeCase.Input;TotalDeletedItemSize=$sizeCase.Input;ItemCount=0;DeletedItemCount=0})
        $userObj=[pscustomobject]@{LastLogonTime='N/A';LastLogonTimeIso8601='';'TotalItemSize-In-MB'='N/A';'TotalDeletedItemSize-In-MB'='N/A';ItemCount='N/A';DeletedItemCount='N/A'}
        & ([scriptblock]::Create($statsNode.Extent.Text))
        Check ($userObj.'TotalItemSize-In-MB' -eq $sizeCase.Expected -and $userObj.'TotalDeletedItemSize-In-MB' -eq $sizeCase.Expected) 'Actual statistics projection collapsed unknown/error/zero'
        Check ($userObj.ItemCount -eq 0 -and $userObj.LastLogonTime -eq $Stats.LastLogonTime -and $userObj.LastLogonTimeIso8601 -eq '') 'Actual statistics projection changed counts or promoted display date'
        $Mbx=Wire ([pscustomobject]@{UseDatabaseQuotaDefaults=$false;ProhibitSendReceiveQuota=$sizeCase.Input})
        . ([scriptblock]::Create($quotaNode.Extent.Text))
        Check ($ProhibitSendReceiveQuota -eq $sizeCase.Expected) 'Actual mailbox quota projection failed'
    }
    function Get-MailboxDatabase { [CmdletBinding()] param($Identity) Check ($Identity -eq 'SyntheticDatabase') 'Database identity lost'; [pscustomobject]@{ProhibitSendReceiveQuota='1 GB (1,073,741,824 bytes)'} }
    $Mbx=Wire ([pscustomobject]@{UseDatabaseQuotaDefaults=$true;Database='SyntheticDatabase';ProhibitSendReceiveQuota='Unlimited'})
    . ([scriptblock]::Create($quotaNode.Extent.Text))
    Check ($ProhibitSendReceiveQuota -eq 1024) 'Inherited database quota was replaced by mailbox quota'

    # Execute the unchanged native branches with a method-bearing local size value.
    $script:ExchangeConnectionMode='LocalSnapin'
    $nativeSize=[pscustomobject]@{}
    $nativeSize | Add-Member ScriptMethod ToMB { 7 }
    $nativeQuota=[pscustomobject]@{IsUnlimited=$false;Value=$nativeSize}
    $Mbx=[pscustomobject]@{UseDatabaseQuotaDefaults=$false;ProhibitSendReceiveQuota=$nativeQuota}
    . ([scriptblock]::Create($quotaNode.Extent.Text))
    Check ($ProhibitSendReceiveQuota -eq 7) 'LocalSnapin native quota behavior changed'
    . (Join-Path $smartRoot 'SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1')
    $Stats=[pscustomobject]@{LastLogonTime=[datetime]::SpecifyKind([datetime]'2026-10-05T09:08:07',[DateTimeKind]::Unspecified);TotalItemSize=$nativeQuota;TotalDeletedItemSize=$nativeQuota;ItemCount=3;DeletedItemCount=1}
    & ([scriptblock]::Create($statsNode.Extent.Text))
    Check ($userObj.'TotalItemSize-In-MB' -eq 7 -and $userObj.ItemCount -eq 3 -and $userObj.LastLogonTimeIso8601 -eq '2026-10-05T09:08:07.0000000') 'LocalSnapin native statistics behavior changed'
    Write-Output ('PASS: {0} remote-object checks; {1}; PSSerializer fixtures only.' -f $checks,$PSVersionTable.PSVersion)
}
finally { [Threading.Thread]::CurrentThread.CurrentCulture=$savedCulture }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCmZYSDVNIVwbWY
# DtF7DcA1p/AfIooOzq76tZYoOe9vd6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICcfvG8HYAhCZjq+QvyHU/wRW6k6OmKlgRCh51Mis+5aMA0GCSqG
# SIb3DQEBAQUABIIBgI/lSdQ/DHwR6jUOBIt6feDwEIrA4TgFgPaxKq7FZeKT8cYO
# nsAZ2ZxYVOu+TDegwlpn7iLmw46plOjpNDuddN2pPPyB26aPCmcZhXXpe2yOESd7
# OvBKGEN+5IjEpkJhfXZrVus0h0yqbTiDMXNDEpli6TQdQcRYUTYmeXXg5zGschlx
# vssBfybojcHxve1vHhZ3KOq++OBlm3ntQp9s+yC3ziOvv5+bS9hynl9QQBB62wxx
# rh54zKe7X2TLuE6XwZ0SDoKbJEIsYg4dCAj25zEWVtFzQvjZlqNsiamULpsPffOC
# du8dJ7P5ZBFR/+Q5zYYOCA/kAq4Nx/GlY3ouy6eCYqzdRVFvX1Yu7I7U6n6RnXj4
# tHOzwSkpoJcckR92g6p1Fyhw/5KSdTAoF6zyjLIhX2Mp/pjkQbR7BgMXSZAPdl7t
# zlfHEpnlE6lgSGHRZzRDTGoLvtzQUEK6V5Ef9kf9Gk+o4ycqZvQcqKd/cDgiyE+3
# l1AxWQxrjYu8v2kAX6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxMTQ3
# MDBaMC8GCSqGSIb3DQEJBDEiBCB+9Lbg06bPX/HIrriVTz+tX0jGVGmNgHSTEWbH
# beq6CTANBgkqhkiG9w0BAQEFAASCAgBOblQ36fUAOd6Yq5D+cVacRx3lDoNuz35Z
# xC45Wb0rUScwKMbFqBL3FcSXiVyZ4eFJIjSJi9BLg94AxeA5baQx+vP5eQpvi4me
# be3M1/21JKvVi6bSj25GTF7Ckvo46Fl9wr32tihGaJG9U3tPLplygAPwXD37OVzV
# xkWReymFpxt5lapQQ6+C/zIZJZgKf5gcLkKvJuOihVeRP6KUVdPRXDck254xqSBa
# kApqVDktnmXdNmR5kA5+pwVw75y8LYJzGI/oXdSFIODVh+vfrlcX6S93Mn2J5H4t
# /GR5TMMlyCPv1QAEDjtfdJG+TpWRKHLsdWZG8w25CK9KudjM8WfBLgKG5MG1DZKa
# V5SyeIKn0igINUhKrx9HDQYQgtVvTQkOdpIbHrLEFeN10lnC3JuzVYocWkjnprHB
# gdrtwJNKNWwwwzu/LCnGRHZ6Lm2Y636ZT9523KjzXab+uTM9Vd+sO00/tInrAKhm
# vkTXscCPA+vGcOnZnAvwfGN/GqvBjFkCSq5/kzb1ImgWAGSimoi0w4Rg7Hxxb0NX
# 2YOx/inntRy8h1U8BFOcxA2H/zYg4pdHHgUjDoSODamUsPCXe4/e4EXzNH2f2CjE
# rEpqXc1ZfCuhjGXbZvSBcquLfzacncNNk26KP6UoEQ9nJqja28EiYy1cnEOHpiB9
# jirp3TZAqw==
# SIG # End signature block
