#Requires -Version 7.0
<#
.SYNOPSIS
Offline source-path regression checks for automatic CMDB preparation.
.VERSION
1.0.0
.NOTES
Uses the actual wrapper helper and Core token resolver extracted through AST.
Only synthetic configuration is supplied; no operational module import, collector,
filesystem write, tenant call or publication is performed.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$wrapper=Join-Path $root 'SmartInventory/PreparedEvidence/SmartM365-CmdbEvidence-Prepare.ps1'
$corePath=Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psm1'
function Read-ParsedScript {
    param([string]$Path)
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw ($errors | Out-String)}
    return $ast
}
$wrapperAst=Read-ParsedScript $wrapper
$coreAst=Read-ParsedScript $corePath
$helper=$wrapperAst.Find({param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Resolve-SmartM365CmdbPreparationSourcePath'
},$true)
$resolver=$coreAst.Find({param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Resolve-SmartM365ConfigValue'
},$true)
if(-not $helper -or -not $resolver){throw 'CMDB source path or Core resolver missing.'}
. ([scriptblock]::Create($helper.Extent.Text))
$mockModule=New-Module -Name SyntheticCmdbSourcePaths -ScriptBlock {
    param($definition)
    . ([scriptblock]::Create($definition))
    $script:SyntheticConfig=[pscustomobject]@{
        WorkspaceRootPath='C:\Synthetic CMDB'
        ProfileKey='test'
        LatestCsvFolderPath='{{WorkspaceRootPath}}\Tenants\{{ProfileKey}}\DATA-LAST'
        DataAllRootPath='{{WorkspaceRootPath}}\Tenants\{{ProfileKey}}\DATA-ALL'
    }
    function Get-SmartM365EffectiveModuleGlobalConfig {return $script:SyntheticConfig}
} -ArgumentList $resolver.Extent.Text
$checks=0
function Assert-Equal {
    param($Actual,$Expected,[string]$Label)
    if($Actual -cne $Expected){throw "Failed: $Label"}
    $script:checks++
}
try {
    $inherited='{{WorkspaceRootPath}}\Tenants\{{ProfileKey}}\DATA-LAST'
    $expected='C:\Synthetic CMDB\Tenants\test\DATA-LAST'
    # The orchestrator supplies no SourceRootPath: inherited profile tokens must resolve.
    foreach($marker in @($null,'',' ','__USE_GLOBAL__','USE_GLOBAL',' __USE_GLOBAL__ ')){
        $actual=Resolve-SmartM365CmdbPreparationSourcePath -CoreModule $mockModule -Configuration @{LatestCsvFolderPath=$marker} -LatestCsvFolderPath $inherited
        Assert-Equal $actual $expected 'Inherited tenant source path'
        Assert-Equal (Join-Path (Split-Path $actual -Parent) 'DATA-POWERBI-CMDB') 'C:\Synthetic CMDB\Tenants\test\DATA-POWERBI-CMDB' 'Output sibling uses resolved source'
    }
    $actual=Resolve-SmartM365CmdbPreparationSourcePath -CoreModule $mockModule -Configuration @{} -LatestCsvFolderPath $inherited
    Assert-Equal $actual $expected 'Missing local key inherits tenant path'
    foreach($case in @(
        @{Local='{{LatestCsvFolderPath}}';Explicit='';Expected=$expected},
        @{Local='C:\Local\DATA-LAST';Explicit='';Expected='C:\Local\DATA-LAST'},
        @{Local='{{UnknownRoot}}\DATA-LAST';Explicit='C:\Explicit\DATA-LAST';Expected='C:\Explicit\DATA-LAST'},
        @{Local='C:\Local\DATA-LAST';Explicit='{{LatestCsvFolderPath}}';Expected=$expected},
        @{Local='\\synthetic-server\synthetic-share\DATA-LAST';Explicit='';Expected='\\synthetic-server\synthetic-share\DATA-LAST'},
        @{Local='C:\Local\DATA-LAST\';Explicit='';Expected='C:\Local\DATA-LAST'},
        @{Local=' C:\Local\DATA-LAST ';Explicit='';Expected='C:\Local\DATA-LAST'},
        @{Local='{{DataAllRootPath}}\..\DATA-LAST';Explicit='';Expected=$expected},
        @{Local='{{UnknownRoot}}\DATA-LAST';Explicit='\\synthetic-server\synthetic-share\DATA-LAST\';Expected='\\synthetic-server\synthetic-share\DATA-LAST'}
    )){
        $actual=Resolve-SmartM365CmdbPreparationSourcePath -CoreModule $mockModule -Configuration @{LatestCsvFolderPath=$case.Local} -SourceRootPath $case.Explicit -LatestCsvFolderPath $inherited
        Assert-Equal $actual $case.Expected 'Source priority and token normalization'
    }
    foreach($invalid in @('',' ','__USE_GLOBAL__','USE_GLOBAL','relative\DATA-LAST','C:relative\DATA-LAST','\drive-relative\DATA-LAST',
        '{{MissingRoot}}\DATA-LAST','C:\Root\{{Unknown}}\DATA-LAST','C:\Root\{{broken\DATA-LAST','C:\Root\broken}}\DATA-LAST',
        'C:\Root\DATA-POWERBI','C:\Root\DATA-POWERBI-CMDB','C:\Root\DATA-LAST\nested')){
        # Test every selection tier: explicit invalid values must not silently fall back.
        $tiers=if([string]::IsNullOrWhiteSpace($invalid)){@('Inherited')}elseif($invalid -in @('__USE_GLOBAL__','USE_GLOBAL')){@('Explicit','Inherited')}else{@('Explicit','Local','Inherited')}
        foreach($tier in $tiers){
            $parameters=@{CoreModule=$mockModule;Configuration=@{LatestCsvFolderPath='__USE_GLOBAL__'};LatestCsvFolderPath=$inherited}
            switch($tier){
                Explicit {$parameters.SourceRootPath=$invalid}
                Local {$parameters.Configuration.LatestCsvFolderPath=$invalid}
                Inherited {$parameters.LatestCsvFolderPath=$invalid}
            }
            $rejected=$false
            try{Resolve-SmartM365CmdbPreparationSourcePath @parameters | Out-Null}
            catch{
                if($_.Exception.Message -notlike 'CMDB preparation requires a fully resolved absolute DATA-LAST*' -and
                   $_.Exception.Message -ne 'Use the authoritative SmartInventory DATA-LAST, not DATA-POWERBI.'){throw}
                $rejected=$true
            }
            Assert-Equal $rejected $true "Reject invalid $tier source"
        }
    }
    $text=$wrapperAst.Extent.Text
    $call='$source=Resolve-SmartM365CmdbPreparationSourcePath -CoreModule $core -SourceRootPath $SourceRootPath -Configuration $config -LatestCsvFolderPath ([string]$effective.LatestCsvFolderPath)'
    Assert-Equal ($text.Contains($call)) $true 'Entry point uses source resolver with tenant context'
    Assert-Equal ($text.IndexOf($call) -lt $text.IndexOf('$output=Join-Path (Split-Path $source -Parent)')) $true 'Source resolved before output derivation'
    Assert-Equal ($text.IndexOf($call) -lt $text.IndexOf('InitializeScriptEnvironment -OutputPathInit')) $true 'Source rejected before directory initialization'
    Assert-Equal ($helper.Extent.Text.IndexOf('IsPathFullyQualified') -lt $helper.Extent.Text.IndexOf('$source=[IO.Path]::GetFullPath')) $true 'Reject relative paths before GetFullPath'
    Write-Output "PASS: $checks CMDB source path offline checks. Actual Core resolver; synthetic configuration only; no live writes."
} finally {
    Remove-Module $mockModule -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDxNUNWrBT4JvXM
# ehctC6KVxShdCWd5ZCFU0invsE7C8aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKg3wIDzdIYO1oorI955cvFvALlXimeBDJwELhRK5TMWMA0GCSqG
# SIb3DQEBAQUABIIBgHgQQk/Nt0vmr65DTbKmycoqvt7yByIJqwBEZGvP5PGk74+s
# By/pZt0OBjzP61tZLqErzJ+2c/jR4Gi7lgg9PWrjHaPlfmI6AIG1MS96z3IymPmf
# 7CyiMRfllxuZqxryDt7B63/FAkIxm+0urWZLA3yhoMNqoYqDj7FzEjHjEFUKUZsN
# jx5FaYjK3RWegIqEJSA49KJ6Sr11/RywKeezhDp75giG1IG1SBpW7wJluuyaNiEk
# LwVuLDbhuFWZY2yBixo4RsT8mTOEKW/uuCV83Tef9YI7xnpBR4PL3DpZo7AcGqS8
# 9TpUHbm41GYv9pvLvp/yyhsITuHjnMqetFQuRolCCun5DJNsapJeHL2OhI9O+wV+
# HcRlert2oEGEjI6nE0MxfCxwBIuWorXNHHp8zJDyQOF/pvjobKuUh/9bV7VHabOb
# WLD+RDxsGh6FlV3FjGxa6IPn+BcOXSqFCHtQ+wxQ+1r/2PteRDaUjyp+jsS93w0Q
# VePlaBImCvqFaohFS6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgwOTI4
# NThaMC8GCSqGSIb3DQEJBDEiBCBkNAcMm33hyuBdJjLdI7GQ9cddfRzCqwYSXTmW
# a63DKjANBgkqhkiG9w0BAQEFAASCAgBvL4dviX/tPclfzcIB0ZrFrDxiZ0I1P9Cc
# lgnoxthC4liIJtw1gZzKE9+4tsdnjWtgfMsbLhuPshvymMToyY75sE2ZAlC/Wvp3
# 6UvfnHbOy+25/VfJV83mns6V+v1h8NLJe0gC2bnAEPyG/OhSNCBUwM64En0XaWg4
# gkNcgtIOaiJnZU5/Iy8zW3eKneM0+SiBQzBuYwgb+D8m6X0oRjIkN8gE9+Tcpgd7
# 66eRn+N3Rep+v/oa0zW4nuLowagdzVcSPCOUn7wG4T/nEkvTVwj42a2Ux6lW3GdY
# +C2Xgr1I3A55IykSzSb42VyoYJE+YUDN+yMqgsVPgR0Mk6XC+hCPghwP4XhvnLoX
# Ej6rXvLw1kD64n6uonc8ho1EQguIpV55tSR8d46Q9n067uhXKf7r9wb33arzbTWE
# anjDhWCNLCnz3bEGZiskZ9ChQraRbKpTl4v80CFQgbbfplQQxOO15vJ8cd6EDmKA
# QUWhe81F4+tc9Is4vLeT3nkGKJPa5UqQdq2IpWT/L/Y8Q9P0PMWsuB3DM0BX3Ei/
# W3gn3GnP1aTYqB3wwAM42apPq51I7+FYN1LgDYV1H5Fa8XpVBnc/6QhK8vP5ae1J
# Kxw28GrjfY6WFQ+neSkd25a11Z9NUxS/CuwjSCH8NQe2GsT88uoj+wvdKsqGHjKp
# x2ppAtrymA==
# SIG # End signature block
