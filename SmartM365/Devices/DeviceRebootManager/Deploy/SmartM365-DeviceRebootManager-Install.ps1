#Requires -Version 5.1

<#
.SYNOPSIS
    Installs SmartM365 Device Reboot Manager for Intune Win32 deployment.

.DESCRIPTION
    Copies runtime files to Program Files and creates the scheduled task used to
    launch the GUI in the interactive user session. Intune is the default
    deployment channel and removes the optional PowerShell Gallery update task.
    The pre-1.1.0 ProgramData installation is removed: standard users can
    create files there, so it must not host code run in other sessions.
.VERSION
1.1.0
#>

[CmdletBinding()]
param(
    # ProgramW6432 keeps the 64-bit Program Files when Intune starts a 32-bit PowerShell.
    [string]$InstallPath = (Join-Path -Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) -ChildPath 'WCH\SmartDeviceRebootManager'),
    [string]$LegacyInstallPath = (Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\DeviceRebootManager'),
    [string]$ConfigSourcePath = '',
    [switch]$ForceConfig,
    [switch]$SkipScheduledTask,
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [int]$RepeatIntervalMinutes = 240,
    [ValidateSet('Intune', 'PowerShellGallery', 'Local')]
    [string]$PackageSource = 'Intune',
    [string]$UpdateTaskPath = '\SmartM365\',
    [string]$UpdateTaskName = 'Device Reboot Manager Update'
)

$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Remove-LegacyInstallation {
    <#
    Removes the files written by a pre-1.1.0 installation, by exact name only, then the
    folder when it is left empty. The folder is never walked: standard users can create
    entries under ProgramData, so a recursive SYSTEM deletion could be redirected by a
    junction swapped in between the link check and the delete. Nothing is removed unless
    the folder and its parent are owned by SYSTEM or Administrators, which keeps users
    from renaming or replacing them while the named files are deleted.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string[]]$TrustedOwnerSids = @('S-1-5-18', 'S-1-5-32-544')
    )

    $legacyFileNames = @(
        'SmartM365-DeviceRebootManager.version.json.txt'
        'SmartM365-DeviceRebootManager.installation.json'
        'SmartM365-DeviceRebootManager-GUI.ps1'
        'SmartM365-DeviceRebootManager-Demo.ps1'
        'Start-SmartM365-DeviceRebootManager-Demo.cmd'
        'SmartM365-DeviceRebootManager-GUI.strings.psd1'
        'SmartM365-DeviceRebootManager-GUI.config.json'
        'SmartM365-DeviceRebootManager-GUI.config.json.template'
        'SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
        'SmartM365.GuiSplash.ps1'
        'WorkplaceCloudHub.ico'
        'WorkplaceCloudHub-lockup-WPF.png'
        'Start-SmartM365-DeviceRebootManager-GUI.cmd'
        'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
    )

    $root = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    if ($null -eq $root) { return 'Absent' }
    if ($root -isnot [IO.DirectoryInfo]) { return 'Skipped: not a directory' }
    if ($root.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        # A single non-recursive call on the link removes the link itself, never its target.
        [IO.Directory]::Delete($root.FullName, $false)
        return 'Removed: legacy path was a link'
    }

    foreach ($directory in @($root, $root.Parent)) {
        if ($directory.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            return ("Skipped: {0} is a link" -f $directory.FullName)
        }
        $owner = (Get-Acl -LiteralPath $directory.FullName).GetOwner([Security.Principal.SecurityIdentifier]).Value
        if ($TrustedOwnerSids -notcontains $owner) {
            return ("Skipped: {0} is owned by {1}" -f $directory.FullName, $owner)
        }
    }

    $removedCount = 0
    foreach ($fileName in $legacyFileNames) {
        $file = Get-Item -LiteralPath (Join-Path -Path $root.FullName -ChildPath $fileName) -Force -ErrorAction SilentlyContinue
        if ($null -eq $file -or $file -isnot [IO.FileInfo]) { continue }
        if (-not ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -and $file.IsReadOnly) {
            $file.IsReadOnly = $false
        }
        # Deleting a file link or hard link removes that name only.
        [IO.File]::Delete($file.FullName)
        $removedCount++
    }

    try {
        [IO.Directory]::Delete($root.FullName, $false)
        return ("Removed: {0} files and the empty folder" -f $removedCount)
    }
    catch {
        return ("Partially removed: {0} files; other entries were left in place" -f $removedCount)
    }
}

function Add-MissingTemplateProperties {
    param(
        [Parameter(Mandatory = $true)][object]$Target,
        [Parameter(Mandatory = $true)][object]$Template,
        [string]$Prefix = ''
    )

    $added = New-Object System.Collections.Generic.List[string]
    foreach ($templateProperty in $Template.PSObject.Properties) {
        $existing = $Target.PSObject.Properties[$templateProperty.Name]
        if ($null -eq $existing) {
            $Target | Add-Member -NotePropertyName $templateProperty.Name -NotePropertyValue $templateProperty.Value
            $added.Add($Prefix + $templateProperty.Name)
        }
        elseif ($templateProperty.Value -is [System.Management.Automation.PSCustomObject] -and $existing.Value -is [System.Management.Automation.PSCustomObject]) {
            foreach ($nestedName in (Add-MissingTemplateProperties -Target $existing.Value -Template $templateProperty.Value -Prefix ($Prefix + $templateProperty.Name + '.'))) {
                $added.Add($nestedName)
            }
        }
    }

    return $added.ToArray()
}

if (-not (Test-IsAdministrator)) {
    throw 'Installation must run elevated. Intune Win32 install should run as SYSTEM.'
}

$deployRoot = $PSScriptRoot
$sourceRoot = Split-Path -Path $deployRoot -Parent

$requiredFiles = @(
    'SmartM365-DeviceRebootManager.version.json.txt'
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-Demo.ps1'
    'Start-SmartM365-DeviceRebootManager-Demo.cmd'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json.template'
    'SmartM365.GuiSplash.ps1'
    'WorkplaceCloudHub.ico'
    'WorkplaceCloudHub-lockup-WPF.png'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
)

$versionSourcePath = Join-Path -Path $sourceRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
if (-not (Test-Path -LiteralPath $versionSourcePath -PathType Leaf)) {
    throw "Package version manifest not found: $versionSourcePath"
}

try {
    $packageVersionManifest = Get-Content -LiteralPath $versionSourcePath -Raw | ConvertFrom-Json
}
catch {
    throw "Package version manifest is invalid: $($_.Exception.Message)"
}

if ([int]$packageVersionManifest.SchemaVersion -ne 1 -or
    [string]::IsNullOrWhiteSpace([string]$packageVersionManifest.ProductName) -or
    [string]::IsNullOrWhiteSpace([string]$packageVersionManifest.PackageVersion)) {
    throw 'Package version manifest is incomplete or uses an unsupported schema.'
}

New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null

foreach ($fileName in $requiredFiles) {
    $sourcePath = Join-Path -Path $sourceRoot -ChildPath $fileName
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Required package file not found: $sourcePath"
    }

    Copy-Item -LiteralPath $sourcePath -Destination (Join-Path -Path $InstallPath -ChildPath $fileName) -Force
}

$runtimeConfigPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
if (-not [string]::IsNullOrWhiteSpace($ConfigSourcePath)) {
    if (-not (Test-Path -LiteralPath $ConfigSourcePath)) {
        throw "Config source file not found: $ConfigSourcePath"
    }

    Copy-Item -LiteralPath $ConfigSourcePath -Destination $runtimeConfigPath -Force
}
elseif ($ForceConfig -or -not (Test-Path -LiteralPath $runtimeConfigPath)) {
    $templatePath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json.template'
    Copy-Item -LiteralPath $templatePath -Destination $runtimeConfigPath -Force
}

# Upgrades keep local values and gain the keys introduced by the new template.
$installedTemplatePath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json.template'
$runtimeConfig = Get-Content -LiteralPath $runtimeConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$configTemplate = Get-Content -LiteralPath $installedTemplatePath -Raw -Encoding UTF8 | ConvertFrom-Json
$addedConfigKeys = @(Add-MissingTemplateProperties -Target $runtimeConfig -Template $configTemplate)
if ($addedConfigKeys.Count -gt 0) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($runtimeConfigPath, ($runtimeConfig | ConvertTo-Json -Depth 8), $utf8NoBom)
    Write-Output ("Configuration keys added from template: {0}" -f ($addedConfigKeys -join ', '))
}

$normalizedInstallPath = [IO.Path]::GetFullPath($InstallPath).TrimEnd('\')
$normalizedLegacyPath = if ([string]::IsNullOrWhiteSpace($LegacyInstallPath)) { '' } else { [IO.Path]::GetFullPath($LegacyInstallPath).TrimEnd('\') }
if ($normalizedLegacyPath -and $normalizedLegacyPath -ne $normalizedInstallPath -and (Test-Path -LiteralPath $normalizedLegacyPath)) {
    # Its configuration is not migrated: the folder may have been modified by a standard user.
    # A failed cleanup must not block the install; the scheduled task no longer points there.
    try {
        $legacyResult = Remove-LegacyInstallation -Path $normalizedLegacyPath
        Write-Output ("Legacy installation {0}: {1}" -f $normalizedLegacyPath, $legacyResult)
    }
    catch {
        Write-Warning ("Unable to remove legacy installation {0}: {1}" -f $normalizedLegacyPath, $_.Exception.Message)
    }
}

if ($PackageSource -eq 'Intune') {
    $galleryUpdateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
    if ($null -ne $galleryUpdateTask) {
        Unregister-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -Confirm:$false
    }
}

if (-not $SkipScheduledTask) {
    $taskScriptPath = Join-Path -Path $deployRoot -ChildPath 'SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    if (-not (Test-Path -LiteralPath $taskScriptPath)) {
        throw "Scheduled task helper not found: $taskScriptPath"
    }

    & $taskScriptPath `
        -InstallPath $InstallPath `
        -TaskPath $TaskPath `
        -TaskName $TaskName `
        -RepeatIntervalMinutes $RepeatIntervalMinutes `
        -ConfigPath $runtimeConfigPath
}

$installationMetadata = [ordered]@{
    SchemaVersion  = 1
    ProductName    = [string]$packageVersionManifest.ProductName
    PackageVersion = [string]$packageVersionManifest.PackageVersion
    PackageSource  = $PackageSource
    InstalledAtUtc = [DateTime]::UtcNow.ToString('o')
}
$installationMetadata |
    ConvertTo-Json -Depth 3 |
    Set-Content -LiteralPath (Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager.installation.json') -Encoding UTF8

Write-Output ("SmartM365 Device Reboot Manager {0} installed to: {1} (source={2})" -f
    $packageVersionManifest.PackageVersion,$InstallPath,$PackageSource)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBXBFisA7UPxDAV
# rY81N2hJH7X/R2fLPjF++5hZHQ9XKaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMvTtofx/l+57Ks6ziX+kj12+HC5EblJP16N6DkZ8974MA0GCSqG
# SIb3DQEBAQUABIIBgH/bHy+uWXyzz9x19OHRs6iOfLA2mAkSqd/rDbHE/oSUp+YX
# eER8W3kTwUOT6JS63Kf16x0c+Q3RtpeILsvYZfA9a5Vo+qN0rnZ1ylwZF9DwypNZ
# 28a0tjUcy8zQT0P0xEg+I3qhXpQT5G8sqUm3W2TB+JSIY4WVRHsDDVchjGPuTlAM
# Ntii2eNvLawOntSk43t/4r/waXH072anuq2Sj5JaEgG+mJS+1X+psPF5YciQFPlp
# wDcjhhrkSKJxq3L1LiDRgkXdtmiwWfiYvxuuhgJ7QNrof47JZzZvAy04+KRxg5eI
# X9dDxqx+k6A5mK7Nz2CvObUmahlOvz9VK27X7qpsMUH9IJYmVBzGrkOL+GNQJBR4
# OaaEA+mkzS7fshxvzoax4dYjwqujY6YpDbmF121Fyp45875ER75tkwe5G77+50bq
# 0XXB07c/Rt+uJTSG1C/zKH7Hw4whKdK8QQV23jrFaEtsLepgMZAcruoDB+YPbKFk
# FSbceU03l3xwDlVarKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMjU2
# NTZaMC8GCSqGSIb3DQEJBDEiBCAxGIYRtuYNMDYjbzubIw1XP51TBwaWKwfjhh3K
# /c8QjDANBgkqhkiG9w0BAQEFAASCAgCszLfZepwmsOClKgUbIqT2MFNjdYv6VMLN
# 1MlmTl/cq5HV55k5wR7QDRp60ldOaGYiBbMQ5Xq8+/+NNnTWFGdNLJkxPVgrIWbd
# ipPzNNuUA4vJ4R3yagkUOfKm5V26aB9CxdgTwEhjBWSHTIta2okhDAGXRKay9eMj
# gyfmWMptCRGGD2355vhUFA3c9L/FNceUiwiqgr9ZliB8/U7R7H9sWEPmE1M3l6Jp
# g1XRc9ol3RIPYFmE4wEwrueFK+/8OwyzoLAMYKqhXGoZV0EWZHY7gGBDHYYLNI+I
# S2SSswionW20WkSEEX/bGqejPaqn8Ok4PqhiRBU13ZvfPwnmgMojm54X0iSDxtyi
# 3OTyUpYv0FXwdZnotq+bIArQJaeG+ZiDnTB+TQlugnHgrbEhlx6viW6+jQf75NBO
# uDHj41YQVpMWcVAq1eh4UUTOvj7Z3PD59yCJ4hmNrjALixHDtsZbkFXzJ+kVnTDP
# 6ia0YrR6aM7nEcsd29LeAWbrIXwfKNu1HRIyBpreLVLlIyJlozf1IcwnfUUZ4Tmh
# Ybsvq/M3Qfs2HbPskUDGy73jZLlcnlQsfi5Bf5/l9MiAIWb6ZboIHO1LcMn4jycS
# Z9/kwNZj2/5s5L1eTUUbNMnZ1wgNOLTNKtGA4/KJTjMTaQHvtgp3aX9MoVjkji71
# +3ZImHbVFQ==
# SIG # End signature block
