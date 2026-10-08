#Requires -Version 7.0
<#
.SYNOPSIS
Offline CMDB run-log policy and actual terminal-cleanup checks.
.VERSION
1.0.0
.NOTES
Extracts code only. Synthetic settings and completion callbacks; no tenant
initialization, collector, real transcript, authentication or SharePoint request.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$testCount=0
function Assert-True {param([bool]$Value,[string]$Label) if(-not $Value){throw $Label};$script:testCount++}
$path=Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartInventory/PreparedEvidence/SmartM365-CmdbEvidence-Prepare.ps1'
$tokens=$null; $parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$parseErrors)
Assert-True ($parseErrors.Count -eq 0) 'Preparation parser failed.'
$helpers=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-SmartM365CmdbLogUploadEnabled'},$true))
Assert-True ($helpers.Count -eq 1) 'Missing log upload policy.'
. ([scriptblock]::Create($helpers[0].Extent.Text))
$tries=@($ast.EndBlock.Statements | Where-Object {$_ -is [Management.Automation.Language.TryStatementAst]})
Assert-True ($tries.Count -eq 1) 'Expected one preparation lifecycle.'
$text=$tries[0].Finally.Extent.Text
$finish=[scriptblock]::Create($text.Substring(1,$text.Length-2))
$core=New-Module -Name CmdbSyntheticLogCore -ScriptBlock {
    $script:SmartM365TeamsNotificationInProgress=$false
    function Get-ModuleLocalConfigValue {
        param($Config,$Name,$DefaultValue)
        $p=$Config.PSObject.Properties[$Name]
        if($p -and $null -ne $p.Value -and $p.Value -notin @('','__USE_GLOBAL__','USE_GLOBAL')){return $p.Value}
        return $script:ConfiguredUpload
    }
}
foreach($inherited in @($false,$true)){
    & $core {param($v) $script:ConfiguredUpload=$v} $inherited
    foreach($value in @($null,'','__USE_GLOBAL__','USE_GLOBAL',$false,$true,'false','true')){
        $config=@{EnableSharePointUpload=$value}
        $expected=if($null -eq $value -or $value -in @('','__USE_GLOBAL__','USE_GLOBAL')){$inherited}else{[bool]::Parse([string]$value)}
        Assert-True ((Get-SmartM365CmdbLogUploadEnabled -CoreModule $core -Configuration $config) -eq $expected) 'Config override/inheritance failed.'
        Assert-True (-not (Get-SmartM365CmdbLogUploadEnabled -CoreModule $core -Configuration $config -ValidateOnly)) 'ValidateOnly enabled upload.'
    }
}
$bad=@{EnableSharePointUpload='not-a-boolean'}
Assert-True (-not (Get-SmartM365CmdbLogUploadEnabled -CoreModule $core -Configuration $bad -ValidateOnly)) 'Offline validation resolved transport configuration.'
$rejected=$false
try {Get-SmartM365CmdbLogUploadEnabled -CoreModule $core -Configuration $bad | Out-Null}catch{$rejected=$true}
Assert-True $rejected 'Invalid Boolean silently enabled logs.'

function Complete-SmartM365ExecutionContext {
    param($Status,$ErrorRecord,$FailureStage,[switch]$CloseTranscriptBeforeUpload)
    $script:observedUpload=$global:EnableSharePointUpload
    $script:observedStatus=$Status
    $script:observedClosed=$CloseTranscriptBeforeUpload.IsPresent
    if($script:throwCompletion){throw 'Synthetic completion failure.'}
}
function Stop-Transcript { $script:fallbackStops++ }
$names=@('EnableSharePointUpload','EnableTeamsNotifications','SmtpServer','From','To','ErrorMailTo',
    'SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath','AppId','TenantId','Thumb','Thumbprint')
$original=@{}
foreach($name in $names){
    $v=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $original[$name]=@{Exists=($null -ne $v);Value=$(if($v){$v.Value}else{$null})}
}
try {
    foreach($validate in @($false,$true)){
        foreach($enabled in @($false,$true)){
            foreach($outcome in @('Success','Failed','CompletedWithWarnings','CompletionFailure','InitializationFailure')){
                $runtimeInitialized=$outcome -ne 'InitializationFailure'; $teamsGuardInstalled=$true; $previousTeamsGuard=$true
                & $core {$script:SmartM365TeamsNotificationInProgress=$true}
                $ValidateOnly=$validate; $logUploadEnabled=$enabled; $transcriptStarted=$runtimeInitialized
                $failure=if($outcome -eq 'Failed'){'Synthetic preparation failure'}else{$null}
                $preparationWarning=$outcome -eq 'CompletedWithWarnings'
                $script:throwCompletion=$outcome -eq 'CompletionFailure'
                $script:observedUpload=$null; $script:observedStatus=''; $script:observedClosed=$false; $script:fallbackStops=0
                $savedOfflineGlobals=@{}
                foreach($name in $names){
                    $savedOfflineGlobals[$name]=@{Exists=($name -ne 'From');Value=('Synthetic-before-'+$name)}
                    Set-Variable -Name $name -Scope Global -Value ('Synthetic-current-'+$name)
                }
                $global:EnableSharePointUpload=$false
                $caught=$false
                try {. $finish}catch{$caught=$true}
                Assert-True ($caught -eq $script:throwCompletion) 'Cleanup swallowed or fabricated a completion failure.'
                if($runtimeInitialized){
                    Assert-True ($script:observedUpload -eq ($enabled -and -not $validate)) 'Terminal transfer gate differs from policy.'
                    Assert-True $script:observedClosed 'Transcript not handed to Core for closure before transfer.'
                    $expectedStatus=if($outcome -eq 'CompletionFailure'){'Success'}else{$outcome}
                    Assert-True ($script:observedStatus -eq $expectedStatus) 'Preparation status changed by transfer policy.'
                }else{
                    Assert-True ($null -eq $script:observedUpload) 'Uninitialized preparation invoked completion transport.'
                }
                Assert-True ($script:fallbackStops -eq [int]$script:throwCompletion) 'Transcript fallback cleanup failed.'
                Assert-True (& $core {$script:SmartM365TeamsNotificationInProgress}) 'Teams callback guard was not restored.'
                foreach($name in $names){
                    $v=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
                    if($name -eq 'From'){Assert-True ($null -eq $v) 'Absent global leaked.'}
                    else{Assert-True ($v.Value -ceq ('Synthetic-before-'+$name)) "Global not restored: $name"}
                }
            }
        }
    }
} finally {
    foreach($name in $names){
        if($original[$name].Exists){Set-Variable -Name $name -Scope Global -Value $original[$name].Value}
        else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}
    }
}
$source=Get-Content -LiteralPath $path -Raw
Assert-True ($source -match 'if\(\$Publish -and -not \$ValidateOnly\)') 'Explicit prepared publication gate changed.'
Assert-True ($source -match '\$global:EnableSharePointUpload=\$false') 'Preparation-time automatic transfer guard removed.'
Assert-True ($source -match "Get-SmartM365CmdbLogUploadEnabled -CoreModule \`$core -Configuration \`$config -ValidateOnly:\`$ValidateOnly") 'Policy not wired to entry point.'
$template=Get-Content -LiteralPath ($path -replace '\.ps1$','.local.json.txt.template') -Raw | ConvertFrom-Json
Assert-True ($template.EnableSharePointUpload -eq '__USE_GLOBAL__') 'Template does not inherit the configured log policy.'
Write-Output "PASS: $testCount CMDB log policy offline checks. Synthetic callbacks only; no live transfers."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDM4MQOZkiwudDM
# BL0kelZPZCTE+EiAZXwRQmQ3iMACm6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC0N+ybdPXE/+lmJuUynrEUfjJX4CC+NExQwQD48vAhdMA0GCSqG
# SIb3DQEBAQUABIIBgF9izjWt0NZ81Oe3lBUo6aNfoVw5OBJ8MMbqHF0K4nmKXCEG
# /FODuwW29Q5xycC7p1aKAXuH0ZUhGjalx1ughkvnq7U0qP6tNiYQpztFL4K3zSdL
# xsqtk2PodLU1hUPQoALHm+0IssjIMgBj6P3/Dzq4nh4JNvApb15vlmt1vE4fwU3z
# KsOsGZuIu+ZMFmT1B2b3YGUML64pWBvJQ5TfS33gEKQ0qWOtJjjD+oLeTZnxx7cz
# G7jspM5eKJLmBkj+OiWS/1a2H4wNraH9TZk7SrGE7Kc8Kw3frMFLWh0HwZoNTPgP
# Ief+8UnAPwpTA3YffZZls31NWq9AP8grITN2G0nvo1qSf8FfCUTZZtp+OVT3LJzK
# QbYPhABBrvUCRvOSYbvz4WNmMaKB0tphyf6t2kWsgcwjDVDPQ7EtaUZbjoI5fT1q
# bIwr5gYhPRIpa05LIWsnwx1I5ufnbIKh6bcTv1qMGgIE4fVSjZj4LOnYhAfj2sKT
# Mg9Ngn1dd6R5UwHfYqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgwNzQ0
# MjVaMC8GCSqGSIb3DQEJBDEiBCAjbmx8lcDXh6ftyTZ6D8ktt2GDES53xxxUtB/V
# BuLCtDANBgkqhkiG9w0BAQEFAASCAgBijXfWY8Da7gwozFWroakVG1dX++JFL4X0
# /OFpWi9jGUYZXGD2DAx75YyCdat42U7iyvmiKjqj5tYzTEi7WYXhu1kQIJhXN77Q
# Qc+Ekq/RnAVoKNQkNcuIRpUarOZPUZR9+ur6N1qJite+QIR9Hp2gZzWPsjqUt48D
# +Ko+PfxesFqijQ11xhhbtwAY215fo79bVOAgsQ/r5qDdVx+2VbMJjTFlpnUEh23t
# bU3Muy+bohIwHdyITuGAqeUajGbSPQ0lxvE36PPIDdIhd7VjX9mfGR2+x37nPSJQ
# d1MAyRdIZ2wCRWXpe/YavQLd5m+T/7MYRiKiON//NiEKXWSXc5ve4d6FecMYBc3j
# /TtmBfKyxhkodzl7h9YvjkKRCOVcpC3VAmErgLxaVSN+o6DjXrnHR3LaJKDGbRjj
# GC2ypnCpU4KUY0sEpbWuj7/Hkuzz68pWWzPLJIxarAbHkEzOOCMfZfqbdQItwn2R
# DndCQ6U+ImHHXif47Kh6WLTHpVN00tY+7BdL26DSc1Nfhj+TQJ9nDCZq+NBiWK+z
# lgQY+npKdmOFmlT5Ok22JPQozx4M5xp3lG5nxY5+TNzGDyrL6UbWmDCn+pwckG65
# Gxqv3PI+ueaKEfMtJBraqz8+pS8D79DyTTbJ/jB02+9erDXAlsAlHGr2GVKGG+nv
# sBMC7lhEHQ==
# SIG # End signature block
