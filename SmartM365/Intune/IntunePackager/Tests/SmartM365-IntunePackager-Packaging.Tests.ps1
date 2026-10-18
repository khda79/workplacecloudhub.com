#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
.SYNOPSIS
    Pester 5 tests for SmartM365.IntunePackager.Packaging.psm1 (offline, synthetic data).

.DESCRIPTION
    The optional integration test runs the real Win32 Content Prep Tool only when the
    environment variable IP_TEST_CONTENT_PREP_TOOL points to a verified IntuneWinAppUtil.exe.

.VERSION
    1.0.0
#>

BeforeDiscovery {
    $script:RealToolPath = [string]$env:IP_TEST_CONTENT_PREP_TOOL
}

BeforeAll {
    $script:AppRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Core.psm1') -Force
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Packaging.psm1') -Force
    . (Join-Path $PSScriptRoot 'SmartM365-IntunePackager-TestHelpers.ps1')
    Set-IpLogSink -Sink $null -Console $false
    Set-IpLogPath -Path ''
    $script:Template = Read-IpJsonFile -Path (Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.config.json.template')
}

Describe 'Engine registry' {
    It 'registers the IntuneWinAppUtil engine' {
        (Get-IpPackagingEngine -Name 'IntuneWinAppUtil').Name | Should -Be 'IntuneWinAppUtil'
    }

    It 'fails clearly for an unknown engine' {
        { Get-IpPackagingEngine -Name 'Missing' } | Should -Throw "*Unknown packaging engine 'Missing'*"
    }

    It 'reports missing tool settings as engine issues' {
        $engine = Get-IpPackagingEngine -Name 'IntuneWinAppUtil'
        $settings = '{"ContentPrepTool":{"Version":"1.8.7","DownloadUrl":"","Sha256":"","StoragePath":"x"}}' | ConvertFrom-Json
        $issues = @(& $engine.Test @{ Settings = $settings })
        $issues | Should -Contain 'Configuration error: ContentPrepTool.DownloadUrl is required.'
        $issues | Should -Contain 'Configuration error: ContentPrepTool.Sha256 is required.'
    }
}

Describe 'Build request validation' {
    BeforeAll {
        $script:Source = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'Source')
        Set-Content -LiteralPath (Join-Path $script:Source 'setup.cmd') -Value 'exit 0'
    }

    It 'accepts a valid request' {
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'setup.cmd' -OutputFolder (Join-Path $TestDrive 'Out') | Should -BeNullOrEmpty
    }

    It 'rejects a missing source folder' {
        Test-IpPackageBuildRequest -SourceFolder (Join-Path $TestDrive 'None') -SetupFile 'setup.cmd' -OutputFolder (Join-Path $TestDrive 'Out') | Should -Match 'Source folder not found'
    }

    It 'rejects a setup file outside the source folder' {
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile '..\setup.cmd' -OutputFolder (Join-Path $TestDrive 'Out') | Should -Match 'relative path inside the source folder'
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'C:\setup.cmd' -OutputFolder (Join-Path $TestDrive 'Out') | Should -Match 'relative path inside the source folder'
    }

    It 'rejects a missing setup file' {
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'missing.exe' -OutputFolder (Join-Path $TestDrive 'Out') | Should -Match 'Setup file not found'
    }

    It 'rejects an output folder inside the source folder' {
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'setup.cmd' -OutputFolder (Join-Path $script:Source 'Out') | Should -Match 'must not be the source folder'
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'setup.cmd' -OutputFolder $script:Source | Should -Match 'must not be the source folder'
    }

    It 'builds the documented IntuneWinAppUtil arguments' {
        Get-IpIntuneWinAppUtilArguments -SourceFolder 'C:\Src' -SetupFile 'setup.msi' -OutputFolder 'C:\Out' | Should -Be @('-c', 'C:\Src', '-s', 'setup.msi', '-o', 'C:\Out', '-q')
    }
}

Describe 'Content Prep Tool verification' {
    It 'stores the tool per version under the configured path' {
        $settings = [pscustomobject]@{ Version = '1.8.7'; StoragePath = '%TEMP%\IpTools' }
        Get-IpContentPrepToolPath -ToolSettings $settings -AppRoot $TestDrive | Should -Be ([IO.Path]::GetFullPath((Join-Path $env:TEMP 'IpTools\1.8.7\IntuneWinAppUtil.exe')))
    }

    It 'rejects a missing file' {
        $result = Test-IpContentPrepTool -Path (Join-Path $TestDrive 'none.exe') -ExpectedSha256 $script:Template.ContentPrepTool.Sha256
        $result.IsValid | Should -BeFalse
        $result.Reasons | Should -Match 'File not found'
    }

    It 'rejects an unsigned file with a different hash' {
        $path = Join-Path $TestDrive 'IntuneWinAppUtil.exe'
        [IO.File]::WriteAllBytes($path, [byte[]](1..64))
        $result = Test-IpContentPrepTool -Path $path -ExpectedSha256 $script:Template.ContentPrepTool.Sha256
        $result.IsValid | Should -BeFalse
        ($result.Reasons -join ' ') | Should -Match 'SHA256 mismatch'
        ($result.Reasons -join ' ') | Should -Match 'Authenticode signature is not valid'
    }

    It 'refuses a download URL outside the official Microsoft repository' {
        $settings = [pscustomobject]@{ Version = '1.8.7'; StoragePath = (Join-Path $TestDrive 'Tools1'); DownloadUrl = 'https://downloads.contoso.example/IntuneWinAppUtil.exe'; Sha256 = $script:Template.ContentPrepTool.Sha256 }
        Mock -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest { throw 'Download must not be attempted.' }
        { Install-IpContentPrepTool -ToolSettings $settings -AppRoot $TestDrive -Confirm:$false } | Should -Throw '*official Microsoft GitHub repository*'
        Should -Invoke -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest -Times 0
    }

    It 'reuses a verified tool without downloading' {
        $settings = [pscustomobject]@{ Version = '1.8.7'; StoragePath = (Join-Path $TestDrive 'Tools2'); DownloadUrl = $script:Template.ContentPrepTool.DownloadUrl; Sha256 = $script:Template.ContentPrepTool.Sha256 }
        $toolPath = Get-IpContentPrepToolPath -ToolSettings $settings -AppRoot $TestDrive
        New-Item -ItemType Directory -Path (Split-Path $toolPath) -Force | Out-Null
        Set-Content -LiteralPath $toolPath -Value 'placeholder'
        Mock -ModuleName SmartM365.IntunePackager.Packaging Test-IpContentPrepTool { [pscustomobject]@{ IsValid = $true; Reasons = @() } }
        Mock -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest { throw 'Download must not be attempted.' }
        Install-IpContentPrepTool -ToolSettings $settings -AppRoot $TestDrive -Confirm:$false | Should -Be $toolPath
        Should -Invoke -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest -Times 0
    }

    It 'deletes a downloaded file that fails verification' {
        $settings = [pscustomobject]@{ Version = '1.8.7'; StoragePath = (Join-Path $TestDrive 'Tools3'); DownloadUrl = $script:Template.ContentPrepTool.DownloadUrl; Sha256 = $script:Template.ContentPrepTool.Sha256 }
        Mock -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest { [IO.File]::WriteAllBytes($OutFile, [byte[]](1..16)) }
        { Install-IpContentPrepTool -ToolSettings $settings -AppRoot $TestDrive -Confirm:$false } | Should -Throw '*failed verification*'
        $toolPath = Get-IpContentPrepToolPath -ToolSettings $settings -AppRoot $TestDrive
        Test-Path -LiteralPath $toolPath | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Split-Path $toolPath) -Force).Count | Should -Be 0
    }
}

Describe 'Package build orchestration' {
    BeforeAll {
        Register-IpPackagingEngine -Name 'TestEngine' -Description 'Synthetic engine for tests.' -Test { param($Request) @() } -Build {
            param($Request)
            $path = Join-Path $Request.OutputFolder ([IO.Path]::GetFileNameWithoutExtension($Request.SetupFile) + '.intunewin')
            New-TestIntuneWin -Path $path -XmlText (Get-TestDetectionXml -SetupFile $Request.SetupFile) -PayloadSize 1024 | Out-Null
            [pscustomobject]@{ IntuneWinPath = $path; EngineVersion = 'test' }
        }
        $script:BuildSource = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'BuildSource')
        Set-Content -LiteralPath (Join-Path $script:BuildSource 'setup.cmd') -Value 'exit 0'
    }

    AfterAll { Unregister-IpPackagingEngine -Name 'TestEngine' }

    It 'returns the package path, hash and metadata' {
        $result = Invoke-IpPackageBuild -EngineName 'TestEngine' -SourceFolder $script:BuildSource -SetupFile 'setup.cmd' -OutputFolder (Join-Path $TestDrive 'BuildOut') -Settings $script:Template -AppRoot $script:AppRoot
        $result.IntuneWinPath | Should -Be (Join-Path $TestDrive 'BuildOut\setup.intunewin')
        $result.Sha256 | Should -Match '^[0-9A-F]{64}$'
        $result.Metadata.SetupFile | Should -Be 'setup.cmd'
        $result.Metadata.SizeEncrypted | Should -Be 1024
        $result.Engine | Should -Be 'TestEngine'
    }

    It 'stops before building when the request is invalid' {
        { Invoke-IpPackageBuild -EngineName 'TestEngine' -SourceFolder $script:BuildSource -SetupFile 'missing.exe' -OutputFolder (Join-Path $TestDrive 'BuildOut') -Settings $script:Template -AppRoot $script:AppRoot } | Should -Throw '*Setup file not found*'
    }

    It 'fails when the engine does not produce the package' {
        Register-IpPackagingEngine -Name 'BrokenEngine' -Description 'Produces nothing.' -Test { param($Request) @() } -Build { param($Request) [pscustomobject]@{ IntuneWinPath = (Join-Path $Request.OutputFolder 'none.intunewin'); EngineVersion = 'x' } }
        try {
            { Invoke-IpPackageBuild -EngineName 'BrokenEngine' -SourceFolder $script:BuildSource -SetupFile 'setup.cmd' -OutputFolder (Join-Path $TestDrive 'BrokenOut') -Settings $script:Template -AppRoot $script:AppRoot } | Should -Throw '*did not produce the expected file*'
        }
        finally { Unregister-IpPackagingEngine -Name 'BrokenEngine' }
    }
}

Describe 'IntuneWinAppUtil integration' -Tag 'Integration' -Skip:([string]::IsNullOrWhiteSpace($script:RealToolPath)) {
    It 'builds a real package from a small folder' {
        $check = Test-IpContentPrepTool -Path $env:IP_TEST_CONTENT_PREP_TOOL -ExpectedSha256 $script:Template.ContentPrepTool.Sha256
        $check.IsValid | Should -BeTrue
        $source = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'RealSource')
        Set-Content -LiteralPath (Join-Path $source 'install.cmd') -Value 'exit 0'
        $output = Join-Path $TestDrive 'RealOut'
        New-Item -ItemType Directory -Path $output | Out-Null
        $path = Invoke-IpIntuneWinAppUtil -ToolPath $env:IP_TEST_CONTENT_PREP_TOOL -SourceFolder $source -SetupFile 'install.cmd' -OutputFolder $output -TimeoutMinutes 5
        (Read-IpIntuneWinMetadata -Path $path).SetupFile | Should -Be 'install.cmd'
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDsDgqTg96NKvPq
# I2/Fqe2m+TnKsy4zZLGOrr1X/+jYyKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOXC0ho+wL2q/PUJoZ86dnGpS9VtAY+/s1BiT1ozZxjhMA0GCSqG
# SIb3DQEBAQUABIIBgAnHBxxfcVJc2HpHcCzYDl6mRoy490ftVxejgHOP5IX51GOM
# etDklB5HSkdS9bjQh+2zFlkek7rKVTLd3TJw5xkOIDiDmzacsh7mEUFkVGxX6YTA
# 8ERo8eArZebku34wvdTiSTKqxP5v+MogMd5WK3mZUMvmmohM0lNWnb7GIIP/SDJG
# Bkzj11EebYcr1dB2jjtkcQc3rZmqMVTOPa60keYnrb+uWR6YH9UC6NE8Yg6g0Dj/
# /7/VXio535ScEeRHAV3RV+QhM1NeVPa0sNfhaE7ZcD6KN7tMRZx3Zy588ofk2v8P
# 0zfqjJdXwOHmgRQEjpNMbIfNbBoGvsXJIBFdIt2NGqwtYWsvVfeMEjcGZ5nbe53r
# NdxxiEP/ymzXUN5bVgdNu4KCaE4RBfaVx3ph1iRkqDPlqb7v8JYW8gwvvClEIIhA
# ctfWQwxkdSJPZHdMd285VMVilIqXM3BP1VOeolr6o0W1m+WdUWfyAfXlswEf/QVF
# gK8tv+islc8Jr4toEaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzdaMC8GCSqGSIb3DQEJBDEiBCDYSBJ4O8HrrVPhMKbHZv/NnL2PpKSDH2kBW9zo
# rLQL4TANBgkqhkiG9w0BAQEFAASCAgAIGrEudmZe7KmobxUIsTDXbX+JwMd06/YG
# 5rvzGQw/b9N+Q/pqnsovgtMUHG/9H7cqcEgZCbNW01GneNTQqXiGXo7rVWkCw/Wn
# /yQ9oe0TUDW0ICxVrkirDhmGh6/v3Skp9rYjlGD9iv950t1+OO/a07m6ciavBOWL
# V+4/7sQuIZbsv7GcbUWnhw8mJ9F0pWbc3MwwfSPqGOJmmIk4Lo0toy4EZptg/d99
# yii1qS86ZFIu9OlDVe0OXvX+/fGwNi/deK7miwESAXKUYT0pIFOnB8LS5vKXwtwb
# o2L6mrFUVaNBdF8drsbk5eMlW5qxgNNQKWMv7kumEDwT8HTJzLeAbbQ9Qy/T1ygG
# oYUiQzniJnFoOwvQnz+/2lLDru3N6XOhPGrtVliE6i1uMX39E5YKjMBJUiboVDZc
# KDx7BDUdYgNTByBdGW+vewVNSHhXBDF+DFKvFhlFPuXnlPX/SZl38HURO/cRdLtG
# /ZlKuOc3/DOhfoguA7QS/d4x9BB9LVt3TD6QQxfJ9UmbLhQqcNA/sY7L7w3a1alX
# avy9MO8VKVya/WmRL9r3OfjlJxBCWWKOGWbrRaWbP5HT0mcEV+Ls2QbmK/fNOyRx
# NtSfBA0MHnmrBF3zPkIps/vkzcW9FIElfu+dItxrfysQeOWojkw+4eYEW9XIEkfX
# Ix1/c8GgQA==
# SIG # End signature block
