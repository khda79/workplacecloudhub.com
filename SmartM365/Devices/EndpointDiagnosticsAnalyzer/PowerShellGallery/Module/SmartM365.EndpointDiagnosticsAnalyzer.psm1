Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ModuleName = 'SmartM365.EndpointDiagnosticsAnalyzer'

function Get-SmartM365EndpointDiagnosticsAnalyzer {
    [CmdletBinding()]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$InstallScope = 'CurrentUser',
        [string]$InstallPath = ''
    )
    if ([string]::IsNullOrWhiteSpace($InstallPath)) {
        $InstallPath = if ($InstallScope -eq 'AllUsers') { Join-Path $env:ProgramFiles 'SmartM365\EndpointDiagnosticsAnalyzer' } else { Join-Path $env:LOCALAPPDATA 'Programs\SmartM365\EndpointDiagnosticsAnalyzer' }
    }
    $metadataPath = Join-Path $InstallPath 'SmartM365-EndpointDiagnosticsAnalyzer.installation.json'
    $metadata = $null
    if (Test-Path -LiteralPath $metadataPath -PathType Leaf) { try { $metadata=Get-Content $metadataPath -Raw|ConvertFrom-Json } catch {} }
    $required = @('SmartM365-EndpointDiagnosticsAnalyzer.version.json','SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1','HardwareReadiness.ps1','SmartM365.GuiSplash.ps1','Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd','WorkplaceCloudHub-lockup-WPF.png','WorkplaceCloudHub.ico')
    $missing = @($required | Where-Object { -not (Test-Path -LiteralPath (Join-Path $InstallPath $_) -PathType Leaf) })
    $integrityIssues = @()
    try {
        $versionPath=Join-Path $InstallPath 'SmartM365-EndpointDiagnosticsAnalyzer.version.json'
        $version=Get-Content -LiteralPath $versionPath -Raw | ConvertFrom-Json
        if(-not$version.PSObject.Properties['RuntimeHashes']){throw 'Runtime hashes missing.'}
        if(-not$metadata -or -not$metadata.PSObject.Properties['VersionManifestHash'] -or $metadata.VersionManifestHash -ne (Get-FileHash -LiteralPath $versionPath).Hash){throw 'Version manifest integrity mismatch.'}
        foreach($file in $required | Where-Object { $_ -notlike '*.json' }){
            $expected=$version.RuntimeHashes.PSObject.Properties[$file]
            if(-not$expected -or (Get-FileHash -LiteralPath (Join-Path $InstallPath $file)).Hash -ne [string]$expected.Value){$integrityIssues += $file}
        }
    } catch { $integrityIssues += $_.Exception.Message }
    [pscustomobject]@{
        ProductName='Smart Endpoint Diagnostics Analyzer'
        Installed=((Test-Path -LiteralPath $InstallPath -PathType Container) -and $missing.Count -eq 0 -and $null -ne $metadata -and $integrityIssues.Count -eq 0)
        IntegrityIssues=$integrityIssues
        InstallScope=$InstallScope
        InstallPath=$InstallPath
        PackageVersion=if($metadata){[string]$metadata.PackageVersion}else{''}
        PackageSource=if($metadata){[string]$metadata.PackageSource}else{''}
        AutomaticUpdateEnabled=if($metadata){[bool]$metadata.AutomaticUpdateEnabled}else{$false}
        MissingRuntimeFiles=$missing
        MetadataPath=$metadataPath
    }
}

function Install-SmartM365EndpointDiagnosticsAnalyzer {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$InstallScope = 'CurrentUser',
        [string]$InstallPath = '',
        [bool]$EnableAutomaticUpdate = $false,
        [switch]$SkipShortcut
    )
    if ($EnableAutomaticUpdate) { throw 'Automatic Gallery updates are disabled. Use Update-SmartM365EndpointDiagnosticsAnalyzer explicitly.' }
    $installer=Join-Path $PSScriptRoot 'Runtime\Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1'
    if(-not(Test-Path $installer -PathType Leaf)){throw "Installer not found: $installer"}
    $target=if($InstallPath){$InstallPath}else{$InstallScope}
    if($PSCmdlet.ShouldProcess($target,'Install Smart Endpoint Diagnostics Analyzer')){& $installer -InstallScope $InstallScope -InstallPath $InstallPath -PackageSource PowerShellGallery -SkipShortcut:$SkipShortcut}
}

function Update-SmartM365EndpointDiagnosticsAnalyzer {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$InstallScope = 'CurrentUser',
        [string]$InstallPath = '',
        [bool]$IncludePrerelease = $false
    )
    $installed=Get-SmartM365EndpointDiagnosticsAnalyzer -InstallScope $InstallScope -InstallPath $InstallPath
    if($installed.PackageSource -eq 'Intune'){throw 'Intune owns this installation. Update it through Intune.'}
    if($PSCmdlet.ShouldProcess($script:ModuleName,'Update module and redeploy runtime')){
        if(Get-Command Update-PSResource -ErrorAction SilentlyContinue){Update-PSResource -Name $script:ModuleName -Repository PSGallery -Scope $InstallScope -Prerelease:$IncludePrerelease -ErrorAction Stop}
        elseif(Get-Command Update-Module -ErrorAction SilentlyContinue){Update-Module -Name $script:ModuleName -Scope $InstallScope -AllowPrerelease:$IncludePrerelease -Force -ErrorAction Stop}
        else{throw 'Install Microsoft.PowerShell.PSResourceGet or PowerShellGet to update the package.'}
        $latest=Get-Module -ListAvailable -Name $script:ModuleName|Sort-Object Version -Descending|Select-Object -First 1
        if(-not $latest){throw 'Updated module could not be located.'}
        $sameVersion=@(Get-Module -ListAvailable -Name $script:ModuleName | Where-Object Version -eq $latest.Version)
        if($sameVersion.Count -gt 1){throw 'Multiple installed module paths share the latest version. Import the intended package explicitly before redeploying; automatic selection is unsafe.'}
        $installer=Join-Path $latest.ModuleBase 'Runtime\Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1'
        & $installer -InstallScope $InstallScope -InstallPath $InstallPath -PackageSource PowerShellGallery
    }
}

function Uninstall-SmartM365EndpointDiagnosticsAnalyzer {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$InstallScope = 'CurrentUser',
        [string]$InstallPath = '',
        [switch]$RemoveUserData
    )
    $uninstaller=Join-Path $PSScriptRoot 'Runtime\Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Uninstall.ps1'
    if(-not(Test-Path $uninstaller -PathType Leaf)){throw "Uninstaller not found: $uninstaller"}
    $target=if($InstallPath){$InstallPath}else{$InstallScope}
    if($PSCmdlet.ShouldProcess($target,'Uninstall Smart Endpoint Diagnostics Analyzer')){& $uninstaller -InstallScope $InstallScope -InstallPath $InstallPath -RemoveUserData:$RemoveUserData -Confirm:$false}
}

Export-ModuleMember -Function @('Get-SmartM365EndpointDiagnosticsAnalyzer','Install-SmartM365EndpointDiagnosticsAnalyzer','Update-SmartM365EndpointDiagnosticsAnalyzer','Uninstall-SmartM365EndpointDiagnosticsAnalyzer')

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD7BT0cpRjCVTfL
# dNcJrb3oV10zlVM5xmmr19jMygxKEaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAqPIm5rRAODUc0p6TObURerBzP33gC64OQX6c53I9iPMA0GCSqG
# SIb3DQEBAQUABIIBgCj9/a6UMz8rODZ3frT05RbJusURjmRtukDM7+s2gfnFJzeG
# CpaqOTWPaxuPDMcgXUyb+PiL0POJRKW8Ggwh0fAzZWYkjRrsPk4p8AVn41c+1gw3
# CWg3TWLm5zwp/3PuWvVsjWgdJJLCi98WCYYgtlrpX+9+U21H0ZDuP/HPqEqjlBkP
# xqvoX+NMBmo+nL0DiWdeNU47peUNQCvn8hMRQIWOjlbvjAFzo6UbkyaeWIAQoBCO
# 4i5axnIOvh8T9LFjKWt2T42qvysnxrcBq1XtP0I8Ev/8YdEESaIgD4z/EC8KOOoI
# IjZAShEo58sBOKnMSzWtQ4o6anQtZPuQi9ved5dgM7omOL41XzfiqxvAA6BwMnh4
# CzvKA79WT6D1lGfrPnoltvy21zOX0tl8Kd6cx2IuVwHqF7Xcq6xm5stdPS1P/nSg
# muZA0cc5ud1VMQ1IVyx0doO39LGCFiFuYwxWA36SE5olqbFgNxe3+CP7fd/CDHl7
# NFgRboHoejLTbVZ0BaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDQwOTEz
# MzFaMC8GCSqGSIb3DQEJBDEiBCC8L4lNcmC12SDwx/fKsB3Frw6+TeWEk3qMynGh
# 722aWDANBgkqhkiG9w0BAQEFAASCAgCpN8HOdj8Jj3Gshfjjgdm0q0JThVGmAmIG
# JqM+ZxbTpS4LumzNeNqLCkHU4ZrGN8fGppcqrhZn+BkgaiykkGQngtJCeWpzyUWs
# x4RgpmJsX5frH7wGJP4kT9yqK/Bdpjh1ChtWpGemEAUC+pSNB1kOEUb/TIOldhE+
# tcXvyfs0keBsG4/+MhMD4O9KLaG7UsCjAW2N3naZHgwePLWoe1gmSTZlyOUORWdE
# xx/ArRvIJI+Zle7gW57uWXItR52uuXBIsTJh9uRYbEsssvHc4PdJgIoPmu2wRx2J
# gZh0rED4zSVwintLI1qp4GQivsQEg/+89ys4uob+Jkkbab2VOax6PEAsAVgojDtP
# Yncatrr6CNML03XSgIYDlp9A1vRg++qnYzV/brehKTrPl86Vd5Dkt6osGLposA/r
# ubCBtRAX6xAYpEpjhP/hbfI7CSVoE8Ep7zqz8q9bRlxd2hiNJUl069WA61vf8lAO
# gdMotEB/r4EwcrzZPjxpThWMdfDx6Tay9yWSm6SVTMHtjPB3WHY9XWQTQrxrkPxv
# XfEorXQwrOR+RyRB77dNvfIEj9oLFaouL8FFM/45eWEokEraHuA8/D+Eg8Je3d1G
# ycJMpL0ujF/tIYJVtWXgGLf1vNEYcX7vUqoX6BhEMTF4bbohQSB8EhecdQvRpYQt
# G/t1EwUTjQ==
# SIG # End signature block
