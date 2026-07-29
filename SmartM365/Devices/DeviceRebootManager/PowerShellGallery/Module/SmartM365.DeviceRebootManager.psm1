Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ModuleName = 'SmartM365.DeviceRebootManager'
$script:GalleryUpdateScriptName = 'SmartM365-DeviceRebootManager-GalleryUpdate.ps1'

function Get-SmartM365DeviceRebootManager {
    [CmdletBinding()]
    param(
        [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
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
        'SmartM365-DeviceRebootManager.version.json'
        'SmartM365-DeviceRebootManager.installation.json'
        'SmartM365-DeviceRebootManager-GUI.ps1'
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
        [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
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
            -SkipPackageUpdate
    }
}

function Update-SmartM365DeviceRebootManager {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
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
        [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBz2RBv/hObYe8Z
# YwpDjBVX2/P9odmIWHNKdkhZrVB5AqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDYJF66Jsg5yUtC28xPOlNKTMGbLoSofEm9DZ6QSY0uxMA0GCSqG
# SIb3DQEBAQUABIIBgAh7erDW35IWGWfQZqHHFi2QTEOjziH/jnjfYXTsxQvd8HjC
# 3op8n+Qi4BKxG8J729Vv7uRHM01vId19gGVG5uSQ0nuckz7sQquNvDr3zgDqGZoG
# 6CevkB54Z54XP4y5JJ49CQL1wbF4actoNFvS/Qbx3fGbauSe7KZ5IE26UtjsAnoO
# EgHN/S1nXt93bvBzYJaY9jgvNg/xBeQbdl24I5ueK5+inpU0UtB8gD29Ge0h3Xkr
# jTUEAwGdffRQMABphFsd2/o9vZDiBABsUTmhA3bfcWt5TxcxOHF1QqSaKqPKlOCP
# HO53l54pXJU2AFTOpml0VF1BPkoLP+QBOPvPAO4T+DQoA7C2CIFYuJjy4FTHD6NT
# U/K48BikleCW+x9gxdEYq2a71TNb7WcWIL3wklyvz4wldL8DOQ4VxkZtPZbuwX5S
# d4mYT2XCO5HGI/5v/MegXB/jf5B/5byYR48CmWIt94pV+6uQt7Xc0kIQA/gjW8kO
# MMeL4wUAwX4DsjF9e6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjkwNzU0
# MTdaMC8GCSqGSIb3DQEJBDEiBCC/edge+qy5FWwb4kiIcalIxgBxlFa+5puUL+Rg
# ldsJpTANBgkqhkiG9w0BAQEFAASCAgAaOuuXVuy9bwuYCLnrhxmTynyuy2mAVMgR
# SWjfZUFbAeHssTKtnuQlKsB1vTJK3JFfWmXeUDtEfFuP84oTOUBN12dtKUCYxh2d
# OolFEAKND62ZdiELBZ5u6igD1yTknaZunPLtrr3/K2DNsd9ffzPxcdA7NwTQwAzy
# QxoRqyzTS0iZjFvBx1Bv3KZRp+1fEro9OUiYIQ+NDD97Py8X/+W58TDB2Taig+ZJ
# 0SSN4OpV9tRs+L3yFVWV7saWIaUk00XLyGsUPD4FcterfLOdly+LlL9a6vB7AUvJ
# nUzoeRtN0RbcbTxYHeRpN2jjftbrc43AW5RJcDpOSHXhW8E+cVFFnP5VbJ2XVHVQ
# zlaZEK5PwaS3fbzGTglk4PIgG6nKtANnrLWKObOrosinzi4BA0UvJArAgZV0byd5
# +t+zHSBM07pYhvsPeEt9aYO86ErWPvcuMWxoHALsZmZKNtIKriUNFGQD22ioe8yq
# fyLbQr0m+GDTpYURN+JvkkV1CII5QSGOpFXVoiSx2eM50qw5W3QgL24owDb07MVJ
# z5YEp4fvGqL4dbnClGwez0M7OItbbYIwsH2gYuiw/cZI+1M9ya7rrLuuhgXTS3lk
# kVPk9Jign+uB2zPi3+ongCF6fXKx+ugi74w2FRAsaAln1hIrBjrzPdr/LaWckW0H
# nwcOauqL+w==
# SIG # End signature block
