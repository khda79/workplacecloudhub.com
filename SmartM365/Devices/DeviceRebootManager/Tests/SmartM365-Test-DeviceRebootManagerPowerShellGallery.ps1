#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) {
        throw "Assertion failed: $Message"
    }
}

$productRoot = Split-Path -Path $PSScriptRoot -Parent
$galleryRoot = Join-Path -Path $productRoot -ChildPath 'PowerShellGallery'
$moduleSourceRoot = Join-Path -Path $galleryRoot -ChildPath 'Module'
$testRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("SmartM365-GalleryTest-{0}" -f [guid]::NewGuid().ToString('N'))

try {
    $sourcePowerShellFiles = @(Get-ChildItem -LiteralPath $galleryRoot -Recurse -File |
        Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
    Assert-True ($sourcePowerShellFiles.Count -ge 5) 'Expected Gallery source files were not found.'

    foreach ($file in $sourcePowerShellFiles) {
        $parseErrors = $null
        [void][Management.Automation.PSParser]::Tokenize(
            (Get-Content -LiteralPath $file.FullName -Raw),
            [ref]$parseErrors
        )
        Assert-True ($parseErrors.Count -eq 0) ("PowerShell syntax errors found in {0}" -f $file.FullName)
    }

    $moduleSourcePath = Join-Path -Path $moduleSourceRoot -ChildPath 'SmartM365.DeviceRebootManager.psm1'
    $moduleSourceText = Get-Content -LiteralPath $moduleSourcePath -Raw
    $updateHelperPath = Join-Path -Path $moduleSourceRoot -ChildPath 'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
    $updateHelperText = Get-Content -LiteralPath $updateHelperPath -Raw
    Assert-True ($moduleSourceText -match '\[bool\]\$EnableAutomaticUpdate\s*=\s*\$false') 'Installation must disable automatic updates by default.'
    Assert-True ($updateHelperText -match '\[bool\]\$EnableAutomaticUpdate\s*=\s*\$false') 'Update helper must disable automatic updates by default.'
    Assert-True ($moduleSourceText -match '\[Nullable\[bool\]\]\$EnableAutomaticUpdate\s*=\s*\$null') `
        'Manual update must preserve the existing automatic-update state by default.'

    $scheduledTaskScriptPath = Join-Path -Path $productRoot -ChildPath 'Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    $scheduledTaskScriptText = Get-Content -LiteralPath $scheduledTaskScriptPath -Raw
    Assert-True ($scheduledTaskScriptText -notmatch '\.Repetition\.(Interval|Duration)\s*=') `
        'Main scheduled-task script must not mutate a missing Repetition property.'
    Assert-True ($updateHelperText -notmatch '\.Repetition\.(Interval|Duration)\s*=') `
        'Gallery update helper must not mutate a missing Repetition property.'
    Assert-True ($scheduledTaskScriptText -match '(?s)New-ScheduledTaskTrigger\s+`\s*-Once.*-RepetitionInterval.*-RepetitionDuration') `
        'Main repeat trigger must use the compatible Once parameter set.'
    Assert-True ($scheduledTaskScriptText -notmatch 'LeastPrivilege') `
        'Main scheduled-task script must use a valid RunLevel enumerator.'
    Assert-True ($scheduledTaskScriptText -notmatch '-GroupId[^\r\n]+-LogonType') `
        'Group principal must not mix incompatible GroupId and LogonType parameter sets.'
    Assert-True ($scheduledTaskScriptText -match "-GroupId 'S-1-5-32-545' -RunLevel Limited") `
        'Main task principal must use the Users group with Limited run level.'
    Assert-True ($updateHelperText -match '(?s)New-ScheduledTaskTrigger\s+`\s*-Once.*-RepetitionInterval.*-RepetitionDuration') `
        'Gallery update trigger must use the compatible Once parameter set.'

    $sourceManifestPath = Join-Path -Path $moduleSourceRoot -ChildPath 'SmartM365.DeviceRebootManager.psd1'
    $sourceManifest = Test-ModuleManifest -Path $sourceManifestPath
    Assert-True ($sourceManifest.Name -eq 'SmartM365.DeviceRebootManager') 'Unexpected module name.'
    Assert-True ([string]$sourceManifest.PrivateData.PSData.Prerelease -eq 'preview4') 'Unexpected prerelease label.'

    $versionManifestPath = Join-Path -Path $productRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json'
    $versionManifest = Get-Content -LiteralPath $versionManifestPath -Raw | ConvertFrom-Json
    Assert-True ([int]$versionManifest.SchemaVersion -eq 1) 'Unexpected package version schema.'
    Assert-True ([string]$versionManifest.PackageVersion -eq '0.1.0-preview4') 'Version manifest and module prerelease differ.'

    $buildScript = Join-Path -Path $galleryRoot -ChildPath 'SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1'
    $buildResult = & $buildScript `
        -OutputRoot $testRoot `
        -Force `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True $buildResult.Ready 'Build did not report Ready.'
    Assert-True ($buildResult.PackageVersion -eq '0.1.0-preview4') 'Build package version is incorrect.'
    Assert-True (Test-Path -LiteralPath $buildResult.IntuneSourcePath -PathType Container) 'Intune source path was not created.'

    $requiredPackagePaths = @(
        'SmartM365.DeviceRebootManager.psd1'
        'SmartM365.DeviceRebootManager.psm1'
        'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
        'Runtime\SmartM365-DeviceRebootManager.version.json'
        'Runtime\SmartM365-DeviceRebootManager-GUI.ps1'
        'Runtime\SmartM365-DeviceRebootManager-GUI.config.json.template'
        'Runtime\Deploy\SmartM365-DeviceRebootManager-Install.ps1'
        'Runtime\Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1'
    )
    foreach ($relativePath in $requiredPackagePaths) {
        Assert-True (Test-Path -LiteralPath (Join-Path $buildResult.PackagePath $relativePath) -PathType Leaf) `
            ("Required package file is missing: {0}" -f $relativePath)
    }

    $forbiddenNames = @(
        'LOCAL_MEMORY.md'
        'SmartM365-DeviceRebootManager-GUI.config.json'
    )
    $packagedFiles = @(Get-ChildItem -LiteralPath $buildResult.PackagePath -Recurse -File)
    foreach ($forbiddenName in $forbiddenNames) {
        Assert-True (-not ($packagedFiles.Name -contains $forbiddenName)) ("Private file was packaged: {0}" -f $forbiddenName)
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $buildResult.PackagePath 'Logs'))) 'Logs directory was packaged.'

    $vmValidationScript = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Test-DeviceRebootManagerGalleryVm.ps1'
    $vmPreview = & $vmValidationScript -PackagePath $buildResult.PackagePath
    Assert-True ($vmPreview.Mode -eq 'Preview') 'VM validation script did not stay in Preview mode.'
    Assert-True (-not $vmPreview.ChangesAttempted) 'VM preview unexpectedly attempted changes.'
    Assert-True (-not $vmPreview.AutomaticUpdateDefault) 'VM preview does not report automatic updates disabled by default.'
    Assert-True $vmPreview.AllPinnedSignerMatches 'VM preview did not validate the pinned package signer.'
    Assert-True $vmPreview.ScheduledTaskApiCompatible 'VM preview ScheduledTasks API compatibility preflight failed.'

    $builtManifestPath = Join-Path $buildResult.PackagePath 'SmartM365.DeviceRebootManager.psd1'
    Import-Module -Name $builtManifestPath -Force
    $exportedNames = @((Get-Module SmartM365.DeviceRebootManager).ExportedFunctions.Keys)
    foreach ($expectedCommand in @(
        'Get-SmartM365DeviceRebootManager'
        'Install-SmartM365DeviceRebootManager'
        'Uninstall-SmartM365DeviceRebootManager'
        'Update-SmartM365DeviceRebootManager'
    )) {
        Assert-True ($exportedNames -contains $expectedCommand) ("Exported command is missing: {0}" -f $expectedCommand)
    }

    $fakeInstallPath = Join-Path -Path $testRoot -ChildPath 'NotInstalled'
    $status = Get-SmartM365DeviceRebootManager -InstallPath $fakeInstallPath
    Assert-True (-not $status.Installed) 'Status should report an absent installation.'

    $publishScript = Join-Path -Path $galleryRoot -ChildPath 'SmartM365-Publish-DeviceRebootManagerGalleryPackage.ps1'
    $preview = & $publishScript -PackagePath $buildResult.PackagePath
    Assert-True ($preview.Mode -eq 'Preview') 'Publish script did not stay in Preview mode.'
    Assert-True (-not $preview.PublicationAttempted) 'Preview unexpectedly attempted publication.'
    Assert-True (-not $preview.PublicMetadataComplete) 'LicenseUri guard should block public publication in this pilot.'

    $publicationGuardTriggered = $false
    try {
        & $publishScript -PackagePath $buildResult.PackagePath -Execute -AllowPrereleasePublication -Confirm:$false | Out-Null
    }
    catch {
        $publicationGuardTriggered = $_.Exception.Message -like '*public metadata is incomplete*'
        if (-not $publicationGuardTriggered) { throw }
    }
    Assert-True $publicationGuardTriggered 'Execute mode was not blocked by the missing LicenseUri.'

    [pscustomobject]@{
        Result                 = 'PASS'
        ModuleName             = $buildResult.ModuleName
        Version                = $buildResult.Version
        Prerelease             = $buildResult.Prerelease
        PackageFileCount       = $buildResult.FileCount
        PowerShellFileCount    = $buildResult.PowerShellFileCount
        SignaturesValidated    = $buildResult.SignaturesValidated
        PublicationAttempted   = $preview.PublicationAttempted
        PublicMetadataComplete = $preview.PublicMetadataComplete
        PublicationGuard       = $publicationGuardTriggered
        VmPreviewChangesAttempted = $vmPreview.ChangesAttempted
    }
}
finally {
    Remove-Module -Name SmartM365.DeviceRebootManager -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $testRoot) {
        $resolvedTestRoot = [IO.Path]::GetFullPath($testRoot)
        $resolvedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
        if ($resolvedTestRoot.StartsWith($resolvedTempRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
        }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAFyy5/wBF1TmRm
# v6gMmZo6Rp3s2L+t5G5F2Gu+yeme9qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIKAu/3KgJpD91vNvAy9UqSCYA9ZaiMU0ka4EvLtzGc9QMA0GCSqG
# SIb3DQEBAQUABIIBgJxgCPHzBtrpTF5iOFzC8RL0Z6Dzkj6uR6LvHsNUKAz7V0UA
# nydjvx4m5yqKATR+jUF4+lw3QAoduD5VJJ9v2Mj3IzfebgVdce3Wa6Q5Kd3ZKOII
# 90GWQ/hkISHeVgICxqkx/GDyEGHTAamNDHhkI0o8HDQLNCmvkgpmOo0+tWNn/Jyh
# k65fcCMn7zK5E2kRpqLk+Tma+sTm+gpyaqkxQW1s3TuEi1p/YCPAj6cCfdOobg10
# s9R0pGUjj4fNVwpgIa7V8L+8aBYaIGDQ1j4ri4rMQ2lB9xdXhvOuLR7rOy0O4Heo
# XQ3/7JOlAC/HV5SEMPdiCDHARKe+SI7u5eU90CA7G83bvwZkQ5Tf+66n5RKR0rNL
# h91cXmovqEdn6EUSjW1b5E2rGldWamEFHeF9C7RSqtYo0d47oqTHMtNK8MTRvGPZ
# 5yqyp4Uabto3HONWwYt2m0I4/eEbrdAeg/3QS5ZUBmjBIkVCpX1MQIirqx22m+cs
# I/halAUR/ZIZ+lePD6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjgyMjIz
# MDlaMC8GCSqGSIb3DQEJBDEiBCACmRXune2iF25wyxwfBEiX8Y5auZibka3bxo04
# 1ZKyTTANBgkqhkiG9w0BAQEFAASCAgCK95jBpbyeZDqKvWQYLiIQ4jpGTXb/S5Wr
# Zfy+wvMFldsytKgX/YToItLj/iykxg+wrWlVxW0Our4UDjw/+OfX387+XY+zmMu0
# ccZDstboJfVav8/wQv2H+ms/ONbnwezG04+S7Kwj3E8XC+IESDnkZ7mgrLCLlfRS
# 9V0qCPKqDyq024OmzXeG5JFAgkWBftP7xABaVFl+IzJe1pS929du7MvLiSQGZdGh
# Au8QH8/MXuO4GlqM2vpLcNbQbCcSKLKqEgh5aVoo2Tar1ksWSHQlUNaWTwZSMnNP
# RkZCg+uhvQqPguwLIFFNBGRRUpxEdW9Vq0z5nruQ/AK3QGzDAAjGXjFqBkOgj4XQ
# J5r9Mf2ewTHrJuwZReMcbQLwcb504nRIGnr1AVH+SihAOW+MSpMsTB//GJ41PjFl
# NyA3Z1cv5b8TI+rIj9uNWNqUqNicQ5J6c0Gvu9Nshv0GdIpEly/mX01KY65F8zkM
# tU2xflZ/lW4jC/owjYploU1jh1hmXJHEf2LxxV+SP7sotJnppMP/p7XQEJLHV3ok
# 3o743Vez83j+ZLm5yLqtsinzIM/Q3knycda68L/Lwqi9y65b5p1oQJGsxAEckQWE
# 9ttarRasxQtW9G2FZQhy0HnrKSzKzwtvNmMVb6q5XxsabwCGM926s6AGlYDlaLk4
# Y0w2QY5xBg==
# SIG # End signature block
