#Requires -Version 5.1

<#
.SYNOPSIS
Validates the SmartM365 Device Reboot Manager Gallery package on a test VM.

.DESCRIPTION
Preview is the default and never installs the product. With -Execute, this
script validates a clean VM installation, confirms that automatic updates are
disabled by default, tests explicit opt-in, disables the update task again,
and verifies that the local configuration is preserved.

Package integrity uses the same rule as the Gallery update helper: the signed
file catalog must match every packaged file, and every signature must have an
intact hash and the pinned signer thumbprint. The current WorkplaceCloudHub
certificate is self-signed, so an untrusted chain is accepted with a warning.
On an isolated test VM, -TrustSignerCertificate optionally imports the bundled
public certificate into LocalMachine Root and TrustedPublisher.

After the explicit automatic-update opt-in, the registered task action is
replayed with its own executable against a stub that carries the helper's exact
param block, and every parameter must bind to the expected value.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$PackagePath = '',
    [string]$InstallPath = (Join-Path -Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) -ChildPath 'WCH\SmartDeviceRebootManager'),
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

function Get-WinVerifyTrustResult {
    # Direct WinVerifyTrust (WINTRUST_ACTION_GENERIC_VERIFY_V2, no UI, no
    # revocation check, file choice, VERIFY then CLOSE). Returns the HRESULT.
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not ('SmartM365.DeviceRebootManagerWinTrust' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace SmartM365
{
    public static class DeviceRebootManagerWinTrust
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WINTRUST_FILE_INFO
        {
            public uint cbStruct;
            public string pcwszFilePath;
            public IntPtr hFile;
            public IntPtr pgKnownSubject;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WINTRUST_DATA
        {
            public uint cbStruct;
            public IntPtr pPolicyCallbackData;
            public IntPtr pSIPClientData;
            public uint dwUIChoice;
            public uint fdwRevocationChecks;
            public uint dwUnionChoice;
            public IntPtr pFile;
            public uint dwStateAction;
            public IntPtr hWVTStateData;
            public IntPtr pwszURLReference;
            public uint dwProvFlags;
            public uint dwUIContext;
            public IntPtr pSignatureSettings;
        }

        [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int WinVerifyTrust(IntPtr hwnd, [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID, IntPtr pWVTData);

        private const uint WTD_UI_NONE = 2;
        private const uint WTD_REVOKE_NONE = 0;
        private const uint WTD_CHOICE_FILE = 1;
        private const uint WTD_STATEACTION_VERIFY = 1;
        private const uint WTD_STATEACTION_CLOSE = 2;
        private const uint WTD_REVOCATION_CHECK_NONE = 0x10;

        public static int Verify(string path)
        {
            Guid action = new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");
            WINTRUST_FILE_INFO fileInfo = new WINTRUST_FILE_INFO();
            fileInfo.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_FILE_INFO));
            fileInfo.pcwszFilePath = path;
            IntPtr filePointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)));
            IntPtr dataPointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_DATA)));
            try
            {
                Marshal.StructureToPtr(fileInfo, filePointer, false);
                WINTRUST_DATA data = new WINTRUST_DATA();
                data.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_DATA));
                data.dwUIChoice = WTD_UI_NONE;
                data.fdwRevocationChecks = WTD_REVOKE_NONE;
                data.dwUnionChoice = WTD_CHOICE_FILE;
                data.pFile = filePointer;
                data.dwStateAction = WTD_STATEACTION_VERIFY;
                data.dwProvFlags = WTD_REVOCATION_CHECK_NONE;
                Marshal.StructureToPtr(data, dataPointer, false);
                int result = WinVerifyTrust(IntPtr.Zero, action, dataPointer);

                data = (WINTRUST_DATA)Marshal.PtrToStructure(dataPointer, typeof(WINTRUST_DATA));
                data.dwStateAction = WTD_STATEACTION_CLOSE;
                Marshal.StructureToPtr(data, dataPointer, true);
                WinVerifyTrust(IntPtr.Zero, action, dataPointer);
                return result;
            }
            finally
            {
                Marshal.DestroyStructure(filePointer, typeof(WINTRUST_FILE_INFO));
                Marshal.FreeHGlobal(filePointer);
                Marshal.FreeHGlobal(dataPointer);
            }
        }
    }
}
'@
    }
    return [SmartM365.DeviceRebootManagerWinTrust]::Verify($Path)
}

function Get-PinnedSignatureState {
    # Signature rule: WinVerifyTrust must return 0, CERT_E_UNTRUSTEDROOT or
    # CERT_E_CHAINING (self-signed root not trusted locally) and the signer
    # thumbprint must equal the pinned one. Any other result, including
    # TRUST_E_BAD_DIGEST, TRUST_E_NOSIGNATURE or an expired untimestamped
    # certificate, is rejected.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    $actual = if ($signature.SignerCertificate) {
        ([string]$signature.SignerCertificate.Thumbprint).Replace(' ', '').ToUpperInvariant()
    }
    else {
        ''
    }
    $pinnedSignerMatch = ($actual -ne '' -and $actual -eq ([string]$Thumbprint).Replace(' ', '').ToUpperInvariant())
    $trustResult = Get-WinVerifyTrustResult -Path ([IO.Path]::GetFullPath($Path))
    # 0x800B0109 CERT_E_UNTRUSTEDROOT, 0x800B010A CERT_E_CHAINING
    $untrustedChainResults = @([int]-2146762487, [int]-2146762486)
    $chainTrusted = ($trustResult -eq 0)

    [pscustomobject]@{
        Path              = $Path
        Status            = [string]$signature.Status
        TrustResult       = ('0x{0:X8}' -f $trustResult)
        SignerThumbprint  = $actual
        PinnedSignerMatch = $pinnedSignerMatch
        ChainTrusted      = $chainTrusted
        Accepted          = ($pinnedSignerMatch -and ($chainTrusted -or $untrustedChainResults -contains $trustResult))
    }
}

function Get-PackageSignatureState {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $catalogPath = Join-Path -Path $Root -ChildPath 'SmartM365.DeviceRebootManager.Files.cat'
    if (-not (Test-Path -LiteralPath $catalogPath -PathType Leaf)) {
        throw "Package file catalog not found: $catalogPath"
    }
    $catalogStatus = [string](Test-FileCatalog -Path $Root -CatalogFilePath $catalogPath -FilesToSkip '*.cat')

    $files = @(Get-ChildItem -LiteralPath $Root -Recurse -File |
        Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1', '.cat') })
    if (@($files | Where-Object { $_.Extension -ne '.cat' }).Count -eq 0) {
        throw "No PowerShell files found in package: $Root"
    }

    $rows = @(foreach ($file in $files) {
        $state = Get-PinnedSignatureState -Path $file.FullName -Thumbprint $Thumbprint
        if ($state.Accepted -and -not $state.ChainTrusted) {
            Write-Warning ("Pinned signer accepted but its certificate chain is not trusted on this VM: file={0}; winVerifyTrust={1}" -f $file.FullName,$state.TrustResult)
        }

        [pscustomobject]@{
            Path              = $file.FullName
            Status            = $state.Status
            TrustResult       = $state.TrustResult
            SignerThumbprint  = $state.SignerThumbprint
            PinnedSignerMatch = $state.PinnedSignerMatch
            Accepted          = $state.Accepted
            Valid             = ($state.Accepted -and $state.ChainTrusted)
        }
    })

    [pscustomobject]@{
        Files                  = $rows
        CatalogStatus          = $catalogStatus
        CatalogValid           = ($catalogStatus -eq 'Valid')
        AllPinnedSignerMatches = (-not ($rows.PinnedSignerMatch -contains $false))
        AllAccepted            = (-not ($rows.Accepted -contains $false))
        AllValid               = (-not ($rows.Valid -contains $false))
    }
}

function Test-GalleryUpdateTaskBinding {
    param(
        [Parameter(Mandatory = $true)][object]$Task,
        [Parameter(Mandatory = $true)][string]$HelperPath,
        [Parameter(Mandatory = $true)][hashtable]$Expected
    )

    # Replays the registered action with its own executable against a stub that
    # has the helper's exact param block, so binding errors surface offline.
    $action = @($Task.Actions)[0]
    $arguments = [string]$action.Arguments
    if (-not $arguments.Contains($HelperPath)) {
        throw "Update task does not run the installed helper: $arguments"
    }

    $tokens = $null
    $parseErrors = $null
    $helperAst = [Management.Automation.Language.Parser]::ParseFile($HelperPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0 -or $null -eq $helperAst.ParamBlock) {
        throw "Installed update helper param block could not be parsed: $HelperPath"
    }
    $stubLines = New-Object Collections.Generic.List[string]
    foreach ($attribute in @($helperAst.ParamBlock.Attributes)) {
        $stubLines.Add($attribute.Extent.Text)
    }
    $stubLines.Add($helperAst.ParamBlock.Extent.Text)
    $stubLines.Add('$bound = [ordered]@{}')
    $stubLines.Add('foreach ($key in @($PSBoundParameters.Keys)) {')
    $stubLines.Add('    $value = $PSBoundParameters[$key]')
    $stubLines.Add('    if ($value -is [Management.Automation.SwitchParameter]) { $value = $value.IsPresent }')
    $stubLines.Add('    $bound[$key] = $value')
    $stubLines.Add('}')
    $stubLines.Add('''BINDING='' + ($bound | ConvertTo-Json -Compress)')

    $stubRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("SmartM365 TaskBinding {0}" -f [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stubRoot -Force | Out-Null
    try {
        $stubPath = Join-Path -Path $stubRoot -ChildPath 'Gallery Update Stub.ps1'
        Set-Content -LiteralPath $stubPath -Value $stubLines.ToArray() -Encoding UTF8

        $startInfo = New-Object Diagnostics.ProcessStartInfo
        $startInfo.FileName = [string]$action.Execute
        $startInfo.Arguments = $arguments.Replace($HelperPath, $stubPath)
        $startInfo.UseShellExecute = $false
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $process = [Diagnostics.Process]::Start($startInfo)
        $standardOutput = $process.StandardOutput.ReadToEnd()
        $standardError = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $bindingLine = @($standardOutput -split "`r?`n" | Where-Object { $_ -like 'BINDING=*' }) | Select-Object -First 1
        if ($process.ExitCode -ne 0 -or -not $bindingLine) {
            throw ("Update task arguments do not bind: exit={0}; stderr={1}; arguments={2}" -f $process.ExitCode,$standardError.Trim(),$arguments)
        }
        $bound = $bindingLine.Substring(8) | ConvertFrom-Json
        foreach ($name in $Expected.Keys) {
            $boundProperty = $bound.PSObject.Properties[$name]
            $actualValue = if ($boundProperty) { $boundProperty.Value } else { $null }
            if ([string]$actualValue -ne [string]$Expected[$name]) {
                throw ("Update task parameter {0} bound to '{1}' instead of '{2}'. Arguments={3}" -f $name,$actualValue,$Expected[$name],$arguments)
            }
        }
        return $true
    }
    finally {
        Remove-Item -LiteralPath $stubRoot -Recurse -Force -ErrorAction SilentlyContinue
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
$prerelease = [string]$manifest.PrivateData.PSData['Prerelease']
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
if (-not $signatureState.CatalogValid) {
    throw "Package file catalog validation failed: status=$($signatureState.CatalogStatus)"
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
    SignedFileCount           = $signatureState.Files.Count
    CatalogValid              = $signatureState.CatalogValid
    AllPinnedSignerMatches    = $signatureState.AllPinnedSignerMatches
    AllSignaturesAccepted     = $signatureState.AllAccepted
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
    if (-not $signatureState.AllAccepted) {
        $failedFiles = $signatureState.Files | Where-Object { -not $_.Accepted } | Select-Object -ExpandProperty Path
        throw "Package signatures are not accepted by the pinned-signer rule: $($failedFiles -join ', ')"
    }
    if ($TrustSignerCertificate -and -not $signatureState.AllValid) {
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

        $installedHelperPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
        $updateTaskArgumentsBind = Test-GalleryUpdateTaskBinding `
            -Task $updateTaskAfterOptIn `
            -HelperPath $installedHelperPath `
            -Expected @{
                ModuleName               = 'SmartM365.DeviceRebootManager'
                InstallPath              = $InstallPath
                TaskPath                 = $TaskPath
                TaskName                 = $TaskName
                RepeatIntervalMinutes    = 240
                EnableAutomaticUpdate    = $true
                UpdateTaskPath           = $UpdateTaskPath
                UpdateTaskName           = $UpdateTaskName
                UpdateIntervalHours      = 24
                IncludePrerelease        = $null
                ExpectedSignerThumbprint = Get-NormalizedThumbprint $ExpectedSignerThumbprint
            }

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
            UpdateTaskArgumentsBind        = $updateTaskArgumentsBind
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJNNI6ZPg1E87C
# oe9PRLSL87+y6sR4rA6rWKYYoifxgqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII7rjEfOdk9/dmq+r8J/q6utNKOI6xPeJXEMTWSA/BCPMA0GCSqG
# SIb3DQEBAQUABIIBgKRi93JwS0KYmuXL1NeZBjLgrRGDuNlkwkvvTqq3novJZmFZ
# JpOU8s7xpDv/wMoC8Ru5lT1zM6MN80bHcM1sWh2CNGygaFK/1+wB5DN797TIThSA
# 0QOM3gDFrzlionsx23zaSeipGFeEdr5Ev+voMp8qzBBTM9HHwwVdLXGyQ2k7QbyO
# IauQIpYETTCVvTpj1X+lxwEJLgQh2IDicEUBoQA4MqcnuZcw/MZH5qHMGU4HaM1c
# hcLhd2ATvrLKB9GO1t2xcK+3f9RergkGZC8UD+OtHnoQLWWedl4tfJAPYHjRzUV1
# OOej4lh6PoObXFNBPc9F9QS8ncmB/t60ViHhhkP0niLwr+OrnNr11PIKJA9FE9eN
# 6bDd/leMKj8VdFEIFGkGwXv8ASYz5lzb3aMG1GWaDWgFL+Q8dW9sa5LpKsdLC0n4
# nDgd93JTmTz4RWhmw5E4qLb4Q5cVxv9CS2DcBLNiLTFojDfRhdzeNPWJXUXRMg8D
# HYva4aPVEBUVxhel36GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMjQy
# MjNaMC8GCSqGSIb3DQEJBDEiBCAAXYCh0/xjydvnO3WrHubbjbgYmk6gu5T72EdP
# O+v0kjANBgkqhkiG9w0BAQEFAASCAgCAFgRE45JC7p+q9tEBDlfXYd4p5ar2rwnr
# DXeFQdlCqRzlj8mHan1uoQQ1ZIyfMS55Yrw7MLO+S84fPmx3FQ6yD5TVM86WLM7Y
# iRg6Io3LPE59G6j/Ezbp8Zq9W7R9aWGTDo4IeNr01PYbawhdP3kJugbwQvVkHEj6
# n2RxV/gQhQ7onRMWGCZWq2vzF81t5b55qeYMJPOWemAEXB/vsp1XFKYzfaJB170k
# rujmmB84uAE7bBR95bg+gZ78QDpKzMNnwz+jHEnm+DPMlP3+x4zsxsI3mMICjJNb
# HiCGVv4mZaRBRbY6fMXbPFtneG6zEJ+dT48nSvwz0uAQpyM6/CpKtVqufuyFx/EI
# bMB9oDEiXJcI1S3zvyE5KWdoNYtjaA7211ykYiA3Clx5eNEcP3FTNYGHVra7lMJO
# W1AtLYasWSotxB6GuDRXYCCMs6pTvHfeeYAnAVgzPOa585HCEixT7N8fnG2r1Yct
# k6sMqWOEF3DgsHgganhZe2fKTCbz+soAgvrGhpVM8sWjXC816E6pcSFP0tbtk2ej
# PD5XuKvpzp8W7YcDd1Db4liDgdb8a+ciEwbRr1xCAgTkVCx48at5h7RyP+clMUyw
# IEPKyjBze4CE/jO52OvLDdMbJGFagUarK/0uF31mPLExwa8T8OGP2tXpa6YXqyRf
# AkvFMVarhQ==
# SIG # End signature block
