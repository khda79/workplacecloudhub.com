<#
.SYNOPSIS
PowerShell Gallery module that installs, updates, reports and removes SmartM365 Device Reboot Manager.
.DESCRIPTION
Install deploys the runtime carried by this module version. Update first updates
the module from PowerShell Gallery, then deploys the newest installed version.
The default installation path is the admin-only Program Files\WCH\SmartDeviceRebootManager.
.VERSION
1.1.1
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ModuleName = 'SmartM365.DeviceRebootManager'
$script:GalleryUpdateScriptName = 'SmartM365-DeviceRebootManager-GalleryUpdate.ps1'

function Get-SmartM365DeviceRebootManager {
    [CmdletBinding()]
    param(
        [string]$InstallPath = (Join-Path -Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) -ChildPath 'WCH\SmartDeviceRebootManager'),
        [string]$TaskPath = '\SmartM365\',
        [string]$TaskName = 'Device Reboot Manager',
        [string]$UpdateTaskPath = '\SmartM365\',
        [string]$UpdateTaskName = 'Device Reboot Manager Update'
    )

    $metadataPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager.installation.json'
    $metadata = $null
    if (Test-Path -LiteralPath $metadataPath -PathType Leaf) {
        try {
            $metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
        }
        catch {
            Write-Warning ("Unable to read installation metadata. {0}" -f $_.Exception.Message)
        }
    }

    $runtimeFiles = @(
        'SmartM365-DeviceRebootManager.version.json.txt'
        'SmartM365-DeviceRebootManager.installation.json'
        'SmartM365-DeviceRebootManager-GUI.ps1'
        'SmartM365-DeviceRebootManager-Demo.ps1'
        'Start-SmartM365-DeviceRebootManager-Demo.cmd'
        'SmartM365-DeviceRebootManager-GUI.strings.psd1'
        'SmartM365-DeviceRebootManager-GUI.config.json'
        'SmartM365.GuiSplash.ps1'
        'WorkplaceCloudHub.ico'
        'WorkplaceCloudHub-lockup-WPF.png'
    )
    $missingFiles = @($runtimeFiles | Where-Object {
        -not (Test-Path -LiteralPath (Join-Path -Path $InstallPath -ChildPath $_) -PathType Leaf)
    })

    $guiTask = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
    $updateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
    $installedModules = @(Get-Module -ListAvailable -Name $script:ModuleName | Sort-Object Version -Descending)

    [pscustomobject]@{
        ProductName             = 'Smart Device Reboot Manager'
        Installed               = ((Test-Path -LiteralPath $InstallPath -PathType Container) -and $missingFiles.Count -eq 0 -and $null -ne $guiTask)
        InstallPath             = $InstallPath
        DeployedPackageVersion  = if ($metadata) { [string]$metadata.PackageVersion } else { '' }
        PackageSource           = if ($metadata) { [string]$metadata.PackageSource } else { '' }
        InstalledModuleVersions = @($installedModules | ForEach-Object { [string]$_.Version })
        MissingRuntimeFiles     = $missingFiles
        GuiTaskRegistered       = ($null -ne $guiTask)
        UpdateTaskRegistered    = ($null -ne $updateTask)
        MetadataPath            = $metadataPath
    }
}

function Install-SmartM365DeviceRebootManager {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$InstallPath = (Join-Path -Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) -ChildPath 'WCH\SmartDeviceRebootManager'),
        [string]$ConfigSourcePath = '',
        [switch]$ForceConfig,
        [switch]$SkipScheduledTask,
        [string]$TaskPath = '\SmartM365\',
        [string]$TaskName = 'Device Reboot Manager',
        [ValidateRange(15, 10080)]
        [int]$RepeatIntervalMinutes = 240,
        [bool]$EnableAutomaticUpdate = $false,
        [string]$UpdateTaskPath = '\SmartM365\',
        [string]$UpdateTaskName = 'Device Reboot Manager Update',
        [ValidateRange(1, 168)]
        [int]$UpdateIntervalHours = 24,
        [bool]$IncludePrerelease = $false
    )

    $toolPath = Join-Path -Path $PSScriptRoot -ChildPath ("Tools\{0}" -f $script:GalleryUpdateScriptName)
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        throw "Package deployment helper not found: $toolPath"
    }

    # Deploy exactly this module version, even when a newer one is installed.
    $currentModuleVersion = [string]$MyInvocation.MyCommand.Module.Version

    if ($PSCmdlet.ShouldProcess($InstallPath, 'Install SmartM365 Device Reboot Manager')) {
        & $toolPath `
            -ModuleName $script:ModuleName `
            -InstallPath $InstallPath `
            -ConfigSourcePath $ConfigSourcePath `
            -ForceConfig:$ForceConfig `
            -SkipScheduledTask:$SkipScheduledTask `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -RepeatIntervalMinutes $RepeatIntervalMinutes `
            -EnableAutomaticUpdate:$EnableAutomaticUpdate `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName `
            -UpdateIntervalHours $UpdateIntervalHours `
            -IncludePrerelease:$IncludePrerelease `
            -ModuleVersion $currentModuleVersion `
            -SkipPackageUpdate
    }
}

function Update-SmartM365DeviceRebootManager {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$InstallPath = (Join-Path -Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) -ChildPath 'WCH\SmartDeviceRebootManager'),
        [string]$TaskPath = '\SmartM365\',
        [string]$TaskName = 'Device Reboot Manager',
        [ValidateRange(15, 10080)]
        [int]$RepeatIntervalMinutes = 240,
        [Nullable[bool]]$EnableAutomaticUpdate = $null,
        [string]$UpdateTaskPath = '\SmartM365\',
        [string]$UpdateTaskName = 'Device Reboot Manager Update',
        [ValidateRange(1, 168)]
        [int]$UpdateIntervalHours = 24,
        [bool]$IncludePrerelease = $false
    )

    $toolPath = Join-Path -Path $PSScriptRoot -ChildPath ("Tools\{0}" -f $script:GalleryUpdateScriptName)
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        throw "Package update helper not found: $toolPath"
    }

    $effectiveEnableAutomaticUpdate = if ($null -ne $EnableAutomaticUpdate) {
        [bool]$EnableAutomaticUpdate
    }
    else {
        $existingUpdateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
        $null -ne $existingUpdateTask
    }

    if ($PSCmdlet.ShouldProcess($script:ModuleName, 'Update package from PowerShell Gallery and redeploy runtime')) {
        & $toolPath `
            -ModuleName $script:ModuleName `
            -InstallPath $InstallPath `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -RepeatIntervalMinutes $RepeatIntervalMinutes `
            -EnableAutomaticUpdate:$effectiveEnableAutomaticUpdate `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName `
            -UpdateIntervalHours $UpdateIntervalHours `
            -IncludePrerelease:$IncludePrerelease
    }
}

function Uninstall-SmartM365DeviceRebootManager {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [string]$InstallPath = (Join-Path -Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) -ChildPath 'WCH\SmartDeviceRebootManager'),
        [string]$TaskPath = '\SmartM365\',
        [string]$TaskName = 'Device Reboot Manager',
        [string]$UpdateTaskPath = '\SmartM365\',
        [string]$UpdateTaskName = 'Device Reboot Manager Update',
        [switch]$KeepConfig
    )

    $uninstaller = Join-Path -Path $PSScriptRoot -ChildPath 'Runtime\Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1'
    if (-not (Test-Path -LiteralPath $uninstaller -PathType Leaf)) {
        throw "Runtime uninstaller not found: $uninstaller"
    }

    if ($PSCmdlet.ShouldProcess($InstallPath, 'Uninstall SmartM365 Device Reboot Manager')) {
        $updateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
        if ($null -ne $updateTask) {
            Unregister-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -Confirm:$false
        }

        & $uninstaller `
            -InstallPath $InstallPath `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName `
            -KeepConfig:$KeepConfig
    }
}

Export-ModuleMember -Function @(
    'Get-SmartM365DeviceRebootManager'
    'Install-SmartM365DeviceRebootManager'
    'Uninstall-SmartM365DeviceRebootManager'
    'Update-SmartM365DeviceRebootManager'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCeFNy91jbLwvJ2
# RgNZGTAaaBKaE2SxzwS89EjaqMADgqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL7uQV+0jlAMFFUadq6XQgWvzhqQM+AMhkrDTjnt2O6fMA0GCSqG
# SIb3DQEBAQUABIIBgKwj4RInI9/3j6juyovjMPbno8j/DTN4TsQ8ZMVNsZwXxZTV
# xvXWYN4i2eglyTqktLI3i7wJz4B4lQBKb+qaeVgAaljcLMonCa/4y4c3qwicmehZ
# XdeHpvCkDhwMyuGEfTmxhDhBgMRniBMO++w9ExvBKuPe3hK5XIlUla+OEUnXZQQj
# Cb8hhv2PpiM77Drutsw7kDxcNzH+iKZfcOX3n+E0GrsEHTp+RqWmT0s5gUHZxHJU
# g6ZFuVUpz41Kp5hKjqGJ1t6DpC1BRttuNDP6aqdDZMu80lpmb6RNwlfp3l51bFoJ
# VnXmxJKWo4Aw7sAmb9e0LGA5u2LG6OxHFNezcY3FhQpFosewDQyMY7V4jI2oAuIH
# G/4eDRSgo/UFuCDkWMx8CD0vZoPDVofiJTeQ4uoANuH3Oo+vG/hIFVA6KusbForo
# ZMApY2Tw+vPlwMF/TjsO47uoauItHpj9jUSxNxsdckEUVQMEoBEUJilIKTP099ac
# d+tNO6tBIQClY3MQoKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTgxNjMz
# NDZaMC8GCSqGSIb3DQEJBDEiBCB9z3d6nFxtduvJJMOlSRbZVWWrLRJxEc5TWwOG
# B2RIHjANBgkqhkiG9w0BAQEFAASCAgAIbmI/ti81ouDEoZ7bpD2tOK4kQ5KiU8CE
# P9CQN9k4c7NCogmOELcglHLlkQk7iOhMIYxqyjt9zCjiSX8IHxTeEIL/Ar9rt16U
# cHUPSk+dP+pVc0Jn9ikcXcoe2iYgq3eeW57GcuNpmzrkqqyr7+dfMXNMIlKMgp3E
# gi3lakVOhkmjSkfE5HoECx300LsKo3xmjlGaPIpb0d84L+cPMiC8hBJfniJD3wF2
# j9mcN4bm+FZ6qLVvu02NJVr1WCkI+xrPrlzD6xw//YnQDgDl9l6DZ8ksI5Gg9beE
# p+Rpd442Ya5rn/IQuujDbS6tUJCV1CMjGvXNm2OQ32hanAZ7MNINVIVcjTYGi2xN
# sB+8gin7MkhN7s6O6s7vr3VAg2sXgHWducWWNc0MXW2pn+1Pl24u04ReL7Wmf68W
# hvZquBY82BkErG6oZSXWms/X+zmamQ2C34gEDn8RdnEe8yagrEtqbbUMlWixEq7y
# agLPuqWFdq1GqCovQHyf5DFIfAE12NC8g3nOEvbX+pL48Fw63zDykvr4HEvJl6Rz
# lUf/3dGzcELagW4YWjoawCjZfMomqC0Ge/j047S3zSXG+QP7NWwoEXG5fGKoVJTJ
# Gcsir0eHrS16CI3vd/VpYF5bsTAdoTMi/xr2CbgRm7+MiLscxLTgFxjxSF5msPCd
# wwubZfaAqw==
# SIG # End signature block
