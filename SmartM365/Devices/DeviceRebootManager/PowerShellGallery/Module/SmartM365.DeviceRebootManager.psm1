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
        'SmartM365-DeviceRebootManager.version.json.txt'
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBrq7G9s3VvZOJN
# HmX5onmvH/RQpsK4R+5iheoX2vpqPaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCC8zMyUA7U2uD5vf07XZWZ
# 6dWVSUNtsFRFVV75jpeW8jANBgkqhkiG9w0BAQEFAASCAYChRkYfNUyIzwP8rOmR
# IP1/oHe33RWN2Qz78DUV3vL3lEW5+zZ/GjWH8xUXgaxtdsm7eUcjStdTM4TZdZ3w
# TtX8ctL5aGTDXYxO3McNPDKgCn1P9ghz9cT9ZbAgX2m9ckKQ/AGDETAIK3o0blNp
# VpYhwD3NlCHpP1/hnNAWZK3C2jo5EMMhMFuQ+XG6WIQvEqiw2q6DeA9SOXBPse9+
# FtSnEIY4yN1PmJZT+NouqeSc0UhVeRHdQ0OkSI7+sSSCIce39q5X3LrUSMv4h+vJ
# VvC9oy2dj/v4aYAcQUf1kHVXAsthyq0a9fSxC2ucaufEkgcHP9lWtvZh430v1NWF
# cUpizQv9v8bEJvHU7J01fd2cFfF6GtRBtMjwmHUrEJOEuxNIsZrHCmiGWeY4NP6k
# qNRDz2ffeW8r5rRdZA4oQXpF8aV3/Q8fv0v65uky9paDi4YZQpBxdTlTshNbk/9S
# rbfm6hfTHTVz2Q/379V4HQMv/zKYMplvca0LjgcVtvK1TNI=
# SIG # End signature block
