#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$KeepTestArtifacts,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param([bool]$Condition,[string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

$applicationRoot = Split-Path -Parent $PSScriptRoot
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365\Tests\EndpointDiagnosticsAnalyzer\' + [guid]::NewGuid().ToString('N'))
$galleryRoot = Join-Path $testRoot 'Gallery'
$installPath = Join-Path $testRoot 'Install'
$module = $null
try {
    $version = Get-Content -LiteralPath (Join-Path $applicationRoot 'SmartM365-EndpointDiagnosticsAnalyzer.version.json') -Raw | ConvertFrom-Json
    Assert-True ([string]$version.PackageVersion -eq '0.3.0') 'Unexpected package version.'

    $guiPath = Join-Path $applicationRoot 'SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'
    $guiContent = Get-Content -LiteralPath $guiPath -Raw
    Assert-True ($guiContent -match "AppVersion = '0\.3\.0'") 'GUI version is not aligned.'
    Assert-True ($guiContent -match 'LocalDataRoot = Join-Path .*LOCALAPPDATA.*SmartM365\\EndpointDiagnosticsAnalyzer') 'Local data root is not under LocalAppData.'
    Assert-True ($guiContent -match 'api_key_dpapi') 'DPAPI-backed AI configuration is missing.'
    Assert-True ($guiContent -notmatch 'api_key\s*=\s*\$Config\.ApiKey') 'Plaintext AI key persistence is present.'
    Assert-True ($guiContent -match "entry\.Key -in @\('ValidateOnly', 'AIApiKey'\)") 'Plaintext API key is not excluded from elevation arguments.'

    $scripts = @(Get-ChildItem -LiteralPath $applicationRoot -Recurse -File | Where-Object { $_.Extension -in @('.ps1','.psm1','.psd1') })
    foreach ($scriptFile in $scripts) {
        $tokens = $null
        $errors = $null
        [void][Management.Automation.Language.Parser]::ParseFile($scriptFile.FullName,[ref]$tokens,[ref]$errors)
        Assert-True (@($errors).Count -eq 0) ("PowerShell parse error in {0}: {1}" -f $scriptFile.FullName,(@($errors | ForEach-Object { $_.Message }) -join ' | '))
    }
    $scheduledTaskReferences = @($scripts | Where-Object { $_.FullName -notlike (Join-Path $PSScriptRoot '*') -and (Get-Content -LiteralPath $_.FullName -Raw) -match '(?i)Register-ScheduledTask|New-ScheduledTask' })
    Assert-True ($scheduledTaskReferences.Count -eq 0) 'The application package must not create scheduled tasks.'

    $builder = Join-Path $applicationRoot 'PowerShellGallery\SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1'
    $build = & $builder -OutputRoot $galleryRoot -Force -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True ([bool]$build.Ready) 'Gallery build is not ready.'
    Assert-True ([string]$build.PackageVersion -eq [string]$version.PackageVersion) 'Gallery build version mismatch.'

    $module = Import-Module -Name (Join-Path $build.PackagePath 'SmartM365.EndpointDiagnosticsAnalyzer.psd1') -Force -PassThru
    foreach ($commandName in @(
        'Get-SmartM365EndpointDiagnosticsAnalyzer',
        'Install-SmartM365EndpointDiagnosticsAnalyzer',
        'Update-SmartM365EndpointDiagnosticsAnalyzer',
        'Uninstall-SmartM365EndpointDiagnosticsAnalyzer'
    )) {
        Assert-True ($null -ne (Get-Command -Name $commandName -Module $module.Name -ErrorAction SilentlyContinue)) "Missing exported command: $commandName"
    }

    $install = Install-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath -SkipShortcut -Confirm:$false
    Assert-True ([string]$install.Result -eq 'PASS') 'Gallery installation failed.'
    Assert-True (-not [bool]$install.AutomaticUpdateEnabled) 'Automatic updates must be disabled by default.'
    $status = Get-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath
    Assert-True ([bool]$status.Installed) 'Installed runtime was not detected by the module.'
    Assert-True ([string]$status.PackageVersion -eq [string]$version.PackageVersion) 'Installed runtime version mismatch.'

    $detection = Join-Path $build.IntuneSourcePath 'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Detection.ps1'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $detection -InstallPath $installPath -ExpectedVersion ([string]$version.PackageVersion) -ExpectedPackageSource PowerShellGallery | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) 'Detection script rejected the controlled Gallery installation.'

    # Preflight rejection must leave the existing installation byte-for-byte intact.
    $installedGui=Join-Path $installPath 'SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'
    $before=(Get-FileHash -LiteralPath $installedGui).Hash
    $broken=Join-Path $testRoot 'BrokenRuntime'
    Copy-Item -LiteralPath $build.IntuneSourcePath -Destination $broken -Recurse
    $asset=Join-Path $broken 'WorkplaceCloudHub.ico'
    [IO.File]::WriteAllBytes($asset,[byte[]]@(0))
    $rejected=$false
    try { & (Join-Path $broken 'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1') -InstallScope CurrentUser -InstallPath $installPath -SkipShortcut | Out-Null } catch { $rejected=$_.Exception.Message -match 'hash|integrity|signature' }
    Assert-True ($rejected -and (Get-FileHash -LiteralPath $installedGui).Hash -eq $before) 'Broken package modified the previous installation.'

    # Simulate activation failing after the previous tree has been moved aside.
    function Move-Item {
        param([string]$LiteralPath,[string]$Destination)
        if($LiteralPath -like ($installPath+'.stage-*') -and $Destination -eq $installPath){throw 'Synthetic activation failure'}
        Microsoft.PowerShell.Management\Move-Item -LiteralPath $LiteralPath -Destination $Destination
    }
    $activationRejected=$false
    try { & (Join-Path $build.IntuneSourcePath 'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1') -InstallScope CurrentUser -InstallPath $installPath -PackageSource PowerShellGallery -SkipShortcut | Out-Null } catch { $activationRejected=$_.Exception.Message -match 'Synthetic activation failure' }
    finally { Remove-Item Function:\Move-Item }
    Assert-True ($activationRejected -and (Get-FileHash -LiteralPath $installedGui).Hash -eq $before) 'Failed activation did not restore the previous installation.'

    # Detection and the public status command must reject damaged payloads.
    $installedAsset=Join-Path $installPath 'WorkplaceCloudHub.ico'
    $assetBytes=[IO.File]::ReadAllBytes($installedAsset)
    [IO.File]::WriteAllBytes($installedAsset,[byte[]]@(0))
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $detection -InstallPath $installPath -ExpectedVersion ([string]$version.PackageVersion) -ExpectedPackageSource PowerShellGallery | Out-Null
    Assert-True ($LASTEXITCODE -ne 0) 'Detection accepted a corrupt runtime.'
    Assert-True (-not(Get-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath).Installed) 'Status accepted a corrupt runtime.'
    [IO.File]::WriteAllBytes($installedAsset,$assetBytes)

    # Intune ownership must stop a Gallery update before any repository operation.
    $metadataPath=Join-Path $installPath 'SmartM365-EndpointDiagnosticsAnalyzer.installation.json'
    $metadataText=[IO.File]::ReadAllText($metadataPath)
    $metadata=$metadataText|ConvertFrom-Json;$metadata.PackageSource='Intune'
    [IO.File]::WriteAllText($metadataPath,($metadata|ConvertTo-Json -Depth 5),[Text.Encoding]::UTF8)
    $ownedRejected=$false
    try { Update-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath -Confirm:$false | Out-Null } catch { $ownedRejected=$_.Exception.Message -match 'Intune' }
    Assert-True $ownedRejected 'Gallery update did not reject the Intune-owned target.'
    [IO.File]::WriteAllText($metadataPath,$metadataText,[Text.Encoding]::UTF8)

    Uninstall-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath -Confirm:$false | Out-Null
    Assert-True (-not (Test-Path -LiteralPath $installPath)) 'Controlled installation was not removed.'

    [pscustomobject]@{
        Result = 'PASS'
        PackageVersion = [string]$version.PackageVersion
        PowerShellEdition = $PSVersionTable.PSEdition
        PowerShellVersion = [string]$PSVersionTable.PSVersion
        ParsedFileCount = $scripts.Count
        GalleryPackagePath = [string]$build.PackagePath
        DefaultAutomaticUpdate = $false
        ScheduledTaskCreated = $false
        UserDataRemovalRequested = $false
    }
}
finally {
    if ($module) { Remove-Module -ModuleInfo $module -Force -ErrorAction SilentlyContinue }
    if (-not $KeepTestArtifacts -and (Test-Path -LiteralPath $testRoot)) {
        $resolved=[IO.Path]::GetFullPath($testRoot)
        $allowed=[IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\Tests\EndpointDiagnosticsAnalyzer')).TrimEnd('\')+'\'
        if(-not$resolved.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $resolved) -notmatch '^[0-9a-f]{32}$'){throw 'Unsafe fixture cleanup target'}
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAR7CdLnwp9/pPs
# C6NQlwa4lGHnBhqoCyIoQn20nn8QfqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL9zIa0upWxXBhPW7wzDVWWjep+fkG4vbRNe+kLnmMdwMA0GCSqG
# SIb3DQEBAQUABIIBgIT0YDiVFFvaktN+Eswk7D5yY2w0xQS2LLq/bhhgXplrgqKx
# CWVS9E1x90X1Ct4jP8tYda+CXlqwidP9EaslXQKKCI44arq2EBUL21l4rebDtWep
# hrM71t4d+KSHFRRMc6m2n15HNdLP4kbHq1mNNDvKSmhr7rtRueRzyFBsxpIeek93
# L78VUBGA/A4WjCgB17JZ2KqpV6F7dDjLpIKFexwBc0CqL+9olm37BI6iwIuiik8L
# PRKmq16YDa6GmbtuGiXNS6toPQbxb9asughAh+CDgR0KNdy9QUBvv13ITtfwaSAS
# VfG1XUJsnDNOvmiVn9b/LDejRTJdqqrqXovdLo8JW/lnQsZqEODel0WiFpp05rdj
# 4iX8yuRPno7gLniiKUlz+g4JCz12qMXgI3SKaPYhnes1MIVRkg24HzRg7GO8VwVh
# XJxaFf0ON0+7FWuHlCwFA6XlJA1HRBbP4tE/o9ag2KonZUjNejz0ss2+bGKuNUb2
# 1/XMArkPIc5D3HuCO6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDQxNTQw
# MzhaMC8GCSqGSIb3DQEJBDEiBCCU3WX3GYxvZkIGrUFD5AnZc/6KI2zTQArsqvM1
# 8Oi9YTANBgkqhkiG9w0BAQEFAASCAgBfyPQ4kq5kQpLbGDHzNNK0McND2iX2jGgg
# aMjZEAVx2b4KuZ9Hiz3RV7+9nkANpRlhJYhlWRR1JovoxejVz9RvErjbNhRnOU7K
# ST5ukNKfQGCpbyFKq3V8+/BDA9lPh8IkyHv6PNdQBukUsjSMk7liEikgxYUi43W8
# ATNLW381LX7EpiQEOTKrljOf+aS4qQ/85z9fGS+l1l6SbOna8LAlAPT5mWKssEm3
# vd+V3aO1wBYmizBiSpliQbHdwP3b5xKd7MUezvprU9zq77tyinFC7h4uLsuiCPL3
# kNQYVRUWFAeDv7PpX44YAuf9bJz0q323DfE8d0kr5sWd+/C7d21yTHZxjHbzxeq+
# bEA44jzKZivrdmyFFaLD+co37L0+/geMlECRgDeh3wQ0pORkMBF34rC5OZt5HBRm
# 4Nw/a/Pp06lu83BcWhrRT6e2hVO3uiXLuDsu7uH5H2zz87nb3doQB19PPLYaPIWp
# rNstTZrjbGbu19rIfbxX3OOXsUFDMShqKW7IzcwVR3H9pux3ETKo0s8iMyFLqVB9
# vDQDskq3R+AvCs6VZb4EyCCOOlfbuXa0hPa6+3hOG+GzUTEEHpaD5ReBDQQ4/iJZ
# R06AOAPHlIq/vU2TcJ3wsFQPFqz3IbQf+vvzK+8HaO45d05vzDQzr++Z9rdkQrUS
# womQlasJZw==
# SIG # End signature block
