#Requires -Version 5.1

<#
.SYNOPSIS
Validates the SmartM365 Device Reboot Manager Gallery package on a test VM.

.DESCRIPTION
Preview is the default and never installs the product. With -Execute, this
script validates a clean VM installation, confirms that automatic updates are
disabled by default, tests explicit opt-in, disables the update task again,
and verifies that the local configuration is preserved.

The current WorkplaceCloudHub signing certificate is self-signed. On an
isolated test VM, -TrustSignerCertificate explicitly imports the bundled public
certificate into LocalMachine Root and TrustedPublisher when required.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$PackagePath = '',
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [string]$UpdateTaskPath = '\SmartM365\',
    [string]$UpdateTaskName = 'Device Reboot Manager Update',
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [string]$SignerCertificatePath = '',
    [string]$ReportPath = (Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath 'SmartM365-DeviceRebootManager-VMValidation.json'),
    [switch]$Execute,
    [switch]$TrustSignerCertificate,
    [switch]$AllowExistingInstallation,
    [switch]$CleanupAfterTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($SignerCertificatePath)) {
    $SignerCertificatePath = Join-Path -Path $PSScriptRoot -ChildPath 'WorkplaceCloudHub-CodeSigning.cer'
}

function Get-NormalizedThumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Invoke-IntuneDetection {
    param(
        [Parameter(Mandatory = $true)][string]$DetectionPath,
        [Parameter(Mandatory = $true)][string]$TargetInstallPath,
        [Parameter(Mandatory = $true)][string]$ExpectedVersion,
        [Parameter(Mandatory = $true)][string]$GuiTaskPath,
        [Parameter(Mandatory = $true)][string]$GuiTaskName,
        [Parameter(Mandatory = $true)][string]$GalleryTaskPath,
        [Parameter(Mandatory = $true)][string]$GalleryTaskName
    )

    $powerShellPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) {
        $powerShellPath = 'powershell.exe'
    }

    $arguments = @(
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        $DetectionPath
        '-InstallPath'
        $TargetInstallPath
        '-TaskPath'
        $GuiTaskPath
        '-TaskName'
        $GuiTaskName
        '-ExpectedVersion'
        $ExpectedVersion
        '-UpdateTaskPath'
        $GalleryTaskPath
        '-UpdateTaskName'
        $GalleryTaskName
    )
    $output = @(& $powerShellPath @arguments 2>&1 | ForEach-Object { [string]$_ })
    [pscustomobject]@{
        ExitCode = $LASTEXITCODE
        Output   = $output
    }
}

function Test-ScheduledTaskApiCompatibility {
    try {
        $triggerCommand = Get-Command New-ScheduledTaskTrigger -ErrorAction Stop
        $principalCommand = Get-Command New-ScheduledTaskPrincipal -ErrorAction Stop
        $settingsCommand = Get-Command New-ScheduledTaskSettingsSet -ErrorAction Stop

        $runLevelNames = [Enum]::GetNames($principalCommand.Parameters['RunLevel'].ParameterType)
        $logonTypeNames = [Enum]::GetNames($principalCommand.Parameters['LogonType'].ParameterType)
        $multipleInstanceNames = [Enum]::GetNames($settingsCommand.Parameters['MultipleInstances'].ParameterType)

        $groupPrincipalSet = @($principalCommand.ParameterSets | Where-Object {
            $parameterNames = @($_.Parameters.Name)
            ($parameterNames -contains 'GroupId') -and
                ($parameterNames -contains 'RunLevel') -and
                (-not ($parameterNames -contains 'LogonType'))
        })
        $repetitionTriggerSet = @($triggerCommand.ParameterSets | Where-Object {
            $parameterNames = @($_.Parameters.Name)
            ($parameterNames -contains 'Once') -and
                ($parameterNames -contains 'RepetitionInterval') -and
                ($parameterNames -contains 'RepetitionDuration')
        })

        return (($runLevelNames -contains 'Limited') -and
            ($runLevelNames -contains 'Highest') -and
            ($logonTypeNames -contains 'ServiceAccount') -and
            ($multipleInstanceNames -contains 'IgnoreNew') -and
            $groupPrincipalSet.Count -gt 0 -and
            $repetitionTriggerSet.Count -gt 0)
    }
    catch {
        return $false
    }
}


function Get-PackageSignatureState {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $expected = Get-NormalizedThumbprint $Thumbprint
    $files = @(Get-ChildItem -LiteralPath $Root -Recurse -File |
        Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
    if ($files.Count -eq 0) {
        throw "No PowerShell files found in package: $Root"
    }

    $rows = @(foreach ($file in $files) {
        $signature = Get-AuthenticodeSignature -LiteralPath $file.FullName
        $actual = if ($signature.SignerCertificate) {
            Get-NormalizedThumbprint $signature.SignerCertificate.Thumbprint
        }
        else { '' }

        [pscustomobject]@{
            Path              = $file.FullName
            Status            = [string]$signature.Status
            SignerThumbprint  = $actual
            PinnedSignerMatch = ($actual -eq $expected)
            Valid             = ($signature.Status -eq 'Valid' -and $actual -eq $expected)
        }
    })

    [pscustomobject]@{
        Files                  = $rows
        AllPinnedSignerMatches = (-not ($rows.PinnedSignerMatch -contains $false))
        AllValid               = (-not ($rows.Valid -contains $false))
    }
}

function Import-TestSignerTrust {
    param(
        [Parameter(Mandatory = $true)][string]$CertificatePath,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    if (-not (Test-Path -LiteralPath $CertificatePath -PathType Leaf)) {
        throw "Bundled signer certificate not found: $CertificatePath"
    }

    $certificate = New-Object Security.Cryptography.X509Certificates.X509Certificate2($CertificatePath)
    $actual = Get-NormalizedThumbprint $certificate.Thumbprint
    $expected = Get-NormalizedThumbprint $Thumbprint
    if ($actual -ne $expected) {
        throw "Signer certificate thumbprint mismatch: expected=$expected; actual=$actual"
    }

    $importedStores = New-Object Collections.Generic.List[string]
    foreach ($storeName in @('Root', 'TrustedPublisher')) {
        $storePath = "Cert:\LocalMachine\$storeName\$expected"
        if (-not (Test-Path -LiteralPath $storePath)) {
            Import-Certificate -FilePath $CertificatePath -CertStoreLocation "Cert:\LocalMachine\$storeName" | Out-Null
            $importedStores.Add($storeName)
        }
    }
    return @($importedStores)
}

function Remove-TestSignerTrust {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][string[]]$StoreNames,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $expected = Get-NormalizedThumbprint $Thumbprint
    foreach ($storeName in $StoreNames) {
        $certificatePath = "Cert:\LocalMachine\$storeName\$expected"
        if (Test-Path -LiteralPath $certificatePath) {
            if ($PSCmdlet.ShouldProcess($certificatePath, 'Remove test signer trust')) {
                Remove-Item -LiteralPath $certificatePath -Force
            }
        }
    }
}

if ([string]::IsNullOrWhiteSpace($PackagePath)) {
    $bundledModuleRoot = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.DeviceRebootManager'
    $versionDirectories = @(Get-ChildItem -LiteralPath $bundledModuleRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -as [version] } |
        Sort-Object { [version]$_.Name } -Descending)
    if ($versionDirectories.Count -eq 0) {
        throw "No bundled module version found under: $bundledModuleRoot"
    }
    $PackagePath = $versionDirectories[0].FullName
}
$resolvedPackagePath = (Resolve-Path -LiteralPath $PackagePath -ErrorAction Stop).Path
$manifestPath = Join-Path -Path $resolvedPackagePath -ChildPath 'SmartM365.DeviceRebootManager.psd1'
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Module manifest not found: $manifestPath"
}

$manifest = Test-ModuleManifest -Path $manifestPath
if ($manifest.Name -ne 'SmartM365.DeviceRebootManager') {
    throw "Unexpected module name: $($manifest.Name)"
}
$prerelease = [string]$manifest.PrivateData.PSData.Prerelease
$expectedPackageVersion = if ([string]::IsNullOrWhiteSpace($prerelease)) {
    [string]$manifest.Version
}
else {
    '{0}-{1}' -f $manifest.Version,$prerelease
}
$detectionPath = Join-Path -Path $resolvedPackagePath -ChildPath 'Runtime\Deploy\SmartM365-DeviceRebootManager-Detection.ps1'
$runtimeInstallerPath = Join-Path -Path $resolvedPackagePath -ChildPath 'Runtime\Deploy\SmartM365-DeviceRebootManager-Install.ps1'
foreach ($requiredPath in @($detectionPath,$runtimeInstallerPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required Intune validation file not found: $requiredPath"
    }
}

$signatureState = Get-PackageSignatureState -Root $resolvedPackagePath -Thumbprint $ExpectedSignerThumbprint
if (-not $signatureState.AllPinnedSignerMatches) {
    $failedFiles = $signatureState.Files | Where-Object { -not $_.PinnedSignerMatch } | Select-Object -ExpandProperty Path
    throw "Package signer validation failed: $($failedFiles -join ', ')"
}

$scheduledTaskApiCompatible = Test-ScheduledTaskApiCompatibility
$existingGuiTask = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
$existingUpdateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
$existingInstall = Test-Path -LiteralPath $InstallPath -PathType Container

$preview = [pscustomobject]@{
    Mode                      = if ($Execute) { 'Execute' } else { 'Preview' }
    ModuleName                = $manifest.Name
    Version                   = [string]$manifest.Version
    Prerelease                = $prerelease
    PackageVersion            = $expectedPackageVersion
    PackagePath               = $resolvedPackagePath
    PowerShellFileCount       = $signatureState.Files.Count
    AllPinnedSignerMatches    = $signatureState.AllPinnedSignerMatches
    AllSignaturesTrusted      = $signatureState.AllValid
    ScheduledTaskApiCompatible = $scheduledTaskApiCompatible
    ExistingInstallation      = $existingInstall
    ExistingGuiTask           = ($null -ne $existingGuiTask)
    ExistingUpdateTask        = ($null -ne $existingUpdateTask)
    AutomaticUpdateDefault    = $false
    TrustImportRequested      = [bool]$TrustSignerCertificate
    CleanupRequested          = [bool]$CleanupAfterTest
    ChangesAttempted          = $false
}

if (-not $Execute) {
    return $preview
}
if (-not $scheduledTaskApiCompatible) {
    throw 'ScheduledTasks API compatibility preflight failed on this VM.'
}
if (-not (Test-IsAdministrator)) {
    throw 'VM validation with -Execute must run from an elevated Windows PowerShell session.'
}
if (($existingInstall -or $existingGuiTask -or $existingUpdateTask) -and -not $AllowExistingInstallation) {
    throw 'Existing installation or scheduled task detected. Use a clean VM or explicitly pass -AllowExistingInstallation.'
}

$importedCertificateStores = @()
$automaticTaskEnabledByThisRun = $false
$originalModulePath = $env:PSModulePath
$validationResult = $null

try {
    if (-not $signatureState.AllValid) {
        if (-not $TrustSignerCertificate) {
            throw 'The pinned signer is not trusted on this VM. Rerun with -TrustSignerCertificate only on an isolated test VM.'
        }
        $importedCertificateStores = @(Import-TestSignerTrust -CertificatePath $SignerCertificatePath -Thumbprint $ExpectedSignerThumbprint)
        $signatureState = Get-PackageSignatureState -Root $resolvedPackagePath -Thumbprint $ExpectedSignerThumbprint
        if (-not $signatureState.AllValid) {
            throw 'Package signatures are still not valid after importing the test signer trust.'
        }
    }

    $moduleNameRoot = Split-Path -Path $resolvedPackagePath -Parent
    $galleryRoot = Split-Path -Path $moduleNameRoot -Parent
    $env:PSModulePath = $galleryRoot + [IO.Path]::PathSeparator + $originalModulePath
    Import-Module -Name $manifestPath -Force

    if ($PSCmdlet.ShouldProcess($InstallPath, 'Run Device Reboot Manager VM installation and opt-in validation')) {
        $preview.ChangesAttempted = $true

        Install-SmartM365DeviceRebootManager `
            -InstallPath $InstallPath `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName `
            -EnableAutomaticUpdate:$false `
            -Confirm:$false

        $guiTaskAfterDefaultInstall = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
        $updateTaskAfterDefaultInstall = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
        $configPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
        if ($null -eq $guiTaskAfterDefaultInstall) { throw 'Main GUI task was not created.' }
        if ($null -ne $updateTaskAfterDefaultInstall) { throw 'Automatic update task was created by the default installation.' }
        if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { throw 'Runtime configuration was not created.' }
        $configHashBefore = (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash

        $galleryDetection = Invoke-IntuneDetection `
            -DetectionPath $detectionPath `
            -TargetInstallPath $InstallPath `
            -ExpectedVersion $expectedPackageVersion `
            -GuiTaskPath $TaskPath `
            -GuiTaskName $TaskName `
            -GalleryTaskPath $UpdateTaskPath `
            -GalleryTaskName $UpdateTaskName
        if ($galleryDetection.ExitCode -eq 0 -or
            -not ($galleryDetection.Output -match 'Unexpected installation source: PowerShellGallery')) {
            throw 'Intune detection did not reject the Gallery-managed installation.'
        }

        Install-SmartM365DeviceRebootManager `
            -InstallPath $InstallPath `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName `
            -EnableAutomaticUpdate:$true `
            -Confirm:$false

        $updateTaskAfterOptIn = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
        if ($null -eq $updateTaskAfterOptIn) { throw 'Automatic update task was not created after explicit opt-in.' }
        $automaticTaskEnabledByThisRun = $true

        $updateTaskDriftDetection = Invoke-IntuneDetection `
            -DetectionPath $detectionPath `
            -TargetInstallPath $InstallPath `
            -ExpectedVersion $expectedPackageVersion `
            -GuiTaskPath $TaskPath `
            -GuiTaskName $TaskName `
            -GalleryTaskPath $UpdateTaskPath `
            -GalleryTaskName $UpdateTaskName
        if ($updateTaskDriftDetection.ExitCode -eq 0 -or
            -not ($updateTaskDriftDetection.Output -match 'Unexpected PowerShell Gallery update task')) {
            throw 'Intune detection did not report the local automatic-update task.'
        }

        Install-SmartM365DeviceRebootManager `
            -InstallPath $InstallPath `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName `
            -EnableAutomaticUpdate:$false `
            -Confirm:$false

        $updateTaskAfterOptOut = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
        if ($null -ne $updateTaskAfterOptOut) { throw 'Automatic update task remains after explicit opt-out.' }
        $automaticTaskEnabledByThisRun = $false

        & $runtimeInstallerPath `
            -InstallPath $InstallPath `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName

        $intuneDetection = Invoke-IntuneDetection `
            -DetectionPath $detectionPath `
            -TargetInstallPath $InstallPath `
            -ExpectedVersion $expectedPackageVersion `
            -GuiTaskPath $TaskPath `
            -GuiTaskName $TaskName `
            -GalleryTaskPath $UpdateTaskPath `
            -GalleryTaskName $UpdateTaskName
        if ($intuneDetection.ExitCode -ne 0) {
            throw ("Intune exact-version detection failed: {0}" -f ($intuneDetection.Output -join '; '))
        }

        $configHashAfter = (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash
        if ($configHashAfter -ne $configHashBefore) { throw 'Runtime configuration changed during reinstall validation.' }

        $status = Get-SmartM365DeviceRebootManager `
            -InstallPath $InstallPath `
            -TaskPath $TaskPath `
            -TaskName $TaskName `
            -UpdateTaskPath $UpdateTaskPath `
            -UpdateTaskName $UpdateTaskName
        if (-not $status.Installed) { throw 'Final product status is not installed.' }
        if ($status.UpdateTaskRegistered) { throw 'Final product status reports an automatic update task.' }
        if ($status.PackageSource -ne 'Intune') { throw 'Final product status is not Intune-managed.' }
        if ($status.DeployedPackageVersion -ne $expectedPackageVersion) { throw 'Final installed version is incorrect.' }

        $validationResult = [pscustomobject]@{
            Result                         = 'PASS'
            ModuleName                     = $manifest.Name
            Version                        = [string]$manifest.Version
            DefaultInstallCreatedGuiTask   = $true
            DefaultInstallCreatedUpdateTask = $false
            ExplicitOptInCreatedUpdateTask = $true
            ExplicitOptOutRemovedUpdateTask = $true
            ConfigurationPreserved         = $true
            GalleryInstallRejectedByIntuneDetection = $true
            IntuneDetectedUpdateTaskDrift  = $true
            IntuneExactVersionDetectionPassed = $true
            FinalPackageSource             = 'Intune'
            FinalPackageVersion            = $expectedPackageVersion
            FinalAutomaticUpdateEnabled    = $false
            CleanupPerformed               = [bool]$CleanupAfterTest
            ReportPath                     = $ReportPath
        }

        if ($CleanupAfterTest) {
            Uninstall-SmartM365DeviceRebootManager `
                -InstallPath $InstallPath `
                -TaskPath $TaskPath `
                -TaskName $TaskName `
                -UpdateTaskPath $UpdateTaskPath `
                -UpdateTaskName $UpdateTaskName `
                -Confirm:$false
        }

        $reportParent = Split-Path -Path $ReportPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($reportParent)) {
            New-Item -ItemType Directory -Path $reportParent -Force | Out-Null
        }
        $validationResult | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
        $validationResult
    }
}
finally {
    if ($automaticTaskEnabledByThisRun) {
        $remainingUpdateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
        if ($remainingUpdateTask) {
            Unregister-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -Confirm:$false
        }
    }
    Remove-Module -Name SmartM365.DeviceRebootManager -Force -ErrorAction SilentlyContinue
    $env:PSModulePath = $originalModulePath

    if ($CleanupAfterTest -and $importedCertificateStores.Count -gt 0) {
        Remove-TestSignerTrust -StoreNames $importedCertificateStores -Thumbprint $ExpectedSignerThumbprint -Confirm:$false
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDpwrCFgie+qD0V
# cO6qAuM+sg87WzUg5UNyUXmRX89e26CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIF422LswGosEKQgR901LUkXVUxvGgjdsrrAkHd+OOR70MA0GCSqG
# SIb3DQEBAQUABIIBgDqRzfq+Q+PLASXU8M/j8ANqchQJM5hjqq5jvSAXEnkz6LSz
# 7Y2Qxrs3qkvTUMUq0RchKSAoULFb4kZfcMh0awKxZRlf/DftJTtNVv4kRWEkNegw
# OiRmOAFR5mqd+IHMa7v37E9sMGQofvlsQxWuQG4j0xaCgMiphzgcO+0Fc85Zkn+c
# 8C9fVgaDWZMtUYrhthDpWAhaBT9Na3MK4HEv22RYsHZjU70whDuKExksWfaaaQYj
# EOh4Bdm0Ifgsp6lBlI+3/7IzYt7ovf9kUFX7joYSgjTo0jKX3OyXnTR5dsC+kbCP
# v8LN1ut7TTlgWV88g2Z4a1LE3lJKXwxOWgfLny4GCZxG2jZxZBK6ZaWYpmZ8fY1A
# Y2sa/UXrnkwGGbJFqZgZmCQAMhfEUEQx0Iq2KiHkDD7LGqiuTjOom845osm8Nkrg
# x/zpUUsupcsy+DN78Tya5ZTEKObWTs4dlNrcJBnfivy1a/gEeciFRZNiuGMwWKW9
# mJO0z58vpNeDf0deGaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjgyMjIz
# MDhaMC8GCSqGSIb3DQEJBDEiBCBm/WFWDmZeKIheW2eyvyAw+kUoF1UrzOHwNn59
# N93UujANBgkqhkiG9w0BAQEFAASCAgCUVL41P2zwO6MEAYkMvCdgFCGaqwwReltm
# modKDN/cXA+9b/2pl6wMfTt4lHp3H+gZkrQRC6uaCpoXHTzMdapek9L7uWC1N53D
# l7MjKZFCJz1Lg+QdVnD2FJbanv5OEJMUQG3nPJQxVc2SAbNP9X24/NmMRZHeAmHb
# vcQ1lBejTjAAY6GA/IiUksY7jd1p1y/TplwSIx9MEdyBnWyHq69n/4GoyG2cvbmN
# n/mHcmx/njBCjTw50/2NMCva9T3NU1hFeK/7Y27CSIf9T71K6dWUGm/ui/2YAzI6
# CGZ/j+mMgHkd95Gfjx8hVmP9L8HvMpNv+fOvUZ1qIdD2+aZc4yf9UuEOirUwaHOR
# HNPvHZPlRv8MiuU+5/47LTo9n512vt0RdwS2Q9ebpZJMA53QAwySOIRYk3lrhzVg
# f8RYaTkWyBxHmHzte9UlLMrsi8pcelH4tvyPo9wzXox8Y2GWnHuMqp2XEiyHp/As
# Of7BGwbGyaD5Z6m1aRBd2aD0BLOoGKzwtJAqZ4rLpo14TYaK+5BnQeeXjhObRdYR
# O98M3ANBiol6hM9FE/4eN9p0e5vWadBnuk3whSFHX43kB7pnFgAkBKlaoMaMixsE
# 2N+QfSnAAYXW8ZJwyNuaK6txJGHPIVcxSN46pE+CheANhFe9q41PG0W71uGHghVS
# tRX5QXchVg==
# SIG # End signature block
