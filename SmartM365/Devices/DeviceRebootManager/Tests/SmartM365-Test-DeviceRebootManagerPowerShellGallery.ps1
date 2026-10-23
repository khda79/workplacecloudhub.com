#Requires -Version 5.1
<#
.SYNOPSIS
Validates or packages the complete application runtime, including offline demo dependencies.
.VERSION
1.0.2
#>

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

function Assert-PackagedDemo {
    param([string]$RuntimeRoot)
    $entry=(Join-Path $RuntimeRoot 'SmartM365-DeviceRebootManager-GUI.ps1').Replace("'","''")
    $command='$ErrorActionPreference=''Stop''; $state=. '''+$entry+''' -Demo -ValidateOnly; if($state.Mode -ne ''Demo'' -or $null -eq $state.Window -or $state.Controls.Count -eq 0){throw ''Packaged demo did not return a window and controls''}; $state.Window.Close(); Write-Output ''DEMO-PASS'''
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $output=& powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $encoded 2>&1
    Assert-True ($LASTEXITCODE -eq 0 -and $output -contains 'DEMO-PASS') ('Packaged demo failed: '+($output -join ' | '))
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
    Assert-True ($moduleSourceText -match '\[bool\]\$IncludePrerelease\s*=\s*\$false') `
        'Stable module must not select prerelease updates by default.'
    Assert-True ($updateHelperText -match '\[bool\]\$IncludePrerelease\s*=\s*\$false') `
        'Stable update helper must not select prerelease updates by default.'

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
    $sourcePrerelease = [string]$sourceManifest.PrivateData.PSData['Prerelease']
    Assert-True ([string]::IsNullOrWhiteSpace($sourcePrerelease)) 'Stable manifest must not contain a prerelease label.'

    $versionManifestPath = Join-Path -Path $productRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
    $versionManifest = Get-Content -LiteralPath $versionManifestPath -Raw | ConvertFrom-Json
    Assert-True ([int]$versionManifest.SchemaVersion -eq 1) 'Unexpected package version schema.'
    Assert-True ([string]$versionManifest.PackageVersion -eq '1.0.0') 'Version manifest and stable module version differ.'

    $buildScript = Join-Path -Path $galleryRoot -ChildPath 'SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1'
    $buildResult = & $buildScript `
        -OutputRoot $testRoot `
        -Force `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True $buildResult.Ready 'Build did not report Ready.'
    Assert-True ($buildResult.PackageVersion -eq '1.0.0') 'Stable build package version is incorrect.'
    Assert-True (Test-Path -LiteralPath $buildResult.IntuneSourcePath -PathType Container) 'Intune source path was not created.'

    foreach ($runtimeRoot in @((Join-Path $buildResult.PackagePath 'Runtime'), $buildResult.IntuneSourcePath)) {
        Assert-PackagedDemo -RuntimeRoot $runtimeRoot
    }

    $requiredPackagePaths = @(
        'SmartM365.DeviceRebootManager.psd1'
        'SmartM365.DeviceRebootManager.psm1'
        'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
        'Runtime\SmartM365-DeviceRebootManager.version.json.txt'
        'Runtime\SmartM365-DeviceRebootManager-GUI.ps1'
        'Runtime\SmartM365-DeviceRebootManager-Demo.ps1'
        'Runtime\Start-SmartM365-DeviceRebootManager-Demo.cmd'
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
    Assert-True $preview.PublicMetadataComplete 'Stable package must include approved public license metadata.'

    $missingApiKeyVariable = "SMARTM365_TEST_MISSING_API_KEY_$([guid]::NewGuid().ToString('N'))"
    $publicationGuardTriggered = $false
    try {
        & $publishScript `
            -PackagePath $buildResult.PackagePath `
            -ApiKeyEnvironmentVariable $missingApiKeyVariable `
            -Execute `
            -Confirm:$false | Out-Null
    }
    catch {
        $publicationGuardTriggered = $_.Exception.Message -like '*API key environment variable is empty*'
        if (-not $publicationGuardTriggered) { throw }
    }
    Assert-True $publicationGuardTriggered 'Execute mode was not blocked when the API key was absent.'

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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBZj7h6q6dw9ELY
# j7dzQy7oZ/8oyb1Qxp6es7pfrnU58qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMHzQEoqFMBXcYeTBSw9saEou8cDtlFdOGIN+oW7J3+yMA0GCSqG
# SIb3DQEBAQUABIIBgE005XfVBUafX6dhNRPamUWBsuzAYFafb0A8eWP1g3sS7gVn
# 4LEQxONgX+tUeCe3HYD13yxSVJQsEYxV5Fb9ixr0OFfF8A5J7J54VDQ59B7ltslo
# n8HtkuIink7pEJD5BoKECTLB+lagwyyLBCT5hLDZuVs9jnXnXvvjHW1v1+oNvL6d
# uuWmpKd02Pyeo/kYMhjLuJ2GOBLWH7Wh8n7ly51vLF8wyfuSNBq/d06IbsBUN6KH
# xXEdt6vS6YQZb0rA+HUQMSej+dhoCcwc0/rMPZ52no+sMuRigwvMo93YXTYCUQuv
# rWPujyxuK0ou8WbrnUsX4e0n7LDNfLZZmcL9VNYc7/BV3rDC4i28F7WBu4S6yM4X
# 1nNEw0EwgUJoM3CDaDTXhg6pK/TQR7FPPnNK2zJCRahrutdnbz7qHdSMyZmKZddG
# hQOUVTBO8sckz8VAvctc85wV6BrTLZWcEMMX9Ie7zeudUGJD3cEfUEjGTBuUjsQG
# FKg4g0zkCVQ5qan1+6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxODAy
# MTJaMC8GCSqGSIb3DQEJBDEiBCCBEx45PRoVGjFhyjFgYDFmKWfRf/Ykaor/TXwe
# TXUsYTANBgkqhkiG9w0BAQEFAASCAgCLL8DzkUqJDvbl1JkCBWYpk6MFoUHh7azN
# X5TbWpXfE1FTVnxy06g5wI5eb2axwLiRalTbU6n1tsNTzCIRTMCk4uDqKc9edNOh
# 9ML6oq8EIVvxZTCHkBJaE897I5jD9pf0uLDj6SJ+9a8y5nHmVq57wnT+KH7QMYzZ
# oi4Rpu15MMDwT1eV7AvrDS5JFIMyBHj3zu9ClwBOKoXVFDx8fKvlmDIBGan3aIVP
# yCsl2uhAkzZaKGWTrRGe+nJI1kZZDGTXQ8DRwNyfeQ+M6zItU3XHMbN8lK7rYb56
# yfla+2jymcKELV0qdd09i2bxvzLMtUDmSIzReyNQAYVCYjxTw4MrR2l1cvq8uLZm
# 1SR7VHbv5PMKXlVaM6LPyKkJTfXpEzI7U/pqQ1t5YbWcZh1BUmXkToBIv66vYP33
# 0q27tReWzHq1lqWM0Lx2BPp9oxL9LMNeR5XxJ5Nkcgn2gnB8FOcLSLu/i5UjXqNN
# 0xyNxkc/eAJv6XjOR5Bbiu0KL7vcenkJOOTU8X3K94yKYDoBzr5FMyeMKIEnFR1y
# HySLXyyGivycBGCsLjb8id+mdjAgSIo/ZI+vo7FOfpooZ8tjffCmqnvGLC9/AQaU
# +F5F4QDhyvr97ltth9+jbSwYReTobm/PVWb+KJMuNReb8GgmsMib5dHodyOmtR3u
# L1LM5K9YrQ==
# SIG # End signature block
