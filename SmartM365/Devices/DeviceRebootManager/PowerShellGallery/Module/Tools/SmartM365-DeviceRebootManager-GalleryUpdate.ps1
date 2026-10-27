#Requires -Version 5.1

<#
.SYNOPSIS
Updates and deploys SmartM365 Device Reboot Manager from PowerShell Gallery.

.DESCRIPTION
This helper is packaged with the PowerShell Gallery module and copied to the
admin-only Program Files installation. It updates the module, validates the
signed file catalog and the pinned WorkplaceCloudHub signer, deploys the
runtime, preserves local configuration, and optionally registers the automatic
update task.

Without -SkipPackageUpdate, the newest installed module version is deployed
after the Gallery update. With -SkipPackageUpdate, the module that contains
this script (or the version named by -ModuleVersion) is deployed, so an older
installed version can be redeployed for rollback.

Signature rule (WinVerifyTrust, WINTRUST_ACTION_GENERIC_VERIFY_V2, no revocation
check): a signature is accepted only when the result is 0, CERT_E_UNTRUSTEDROOT
or CERT_E_CHAINING (self-signed root not trusted locally, logged as a warning)
AND the signer thumbprint equals the pinned thumbprint. Any other result, such
as TRUST_E_BAD_DIGEST, TRUST_E_NOSIGNATURE or an expired untimestamped
certificate, is rejected.
#>

[CmdletBinding()]
param(
    [string]$ModuleName = 'SmartM365.DeviceRebootManager',
    [string]$InstallPath = (Join-Path -Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) -ChildPath 'WCH\SmartDeviceRebootManager'),
    [string]$ConfigSourcePath = '',
    [switch]$ForceConfig,
    [switch]$SkipScheduledTask,
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [ValidateRange(15, 10080)]
    [int]$RepeatIntervalMinutes = 240,
    [switch]$EnableAutomaticUpdate,
    [string]$UpdateTaskPath = '\SmartM365\',
    [string]$UpdateTaskName = 'Device Reboot Manager Update',
    [ValidateRange(1, 168)]
    [int]$UpdateIntervalHours = 24,
    [switch]$IncludePrerelease,
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [string]$ModuleVersion = '',
    [switch]$SkipPackageUpdate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:PackageCatalogFileName = 'SmartM365.DeviceRebootManager.Files.cat'
# Package-manager metadata added at install time (PowerShellGet/PSResourceGet).
# The catalog itself is excluded because it cannot list its own hash.
$script:CatalogFilesToSkip = @(
    '*.cat'
    'PSGetModuleInfo.xml'
    '*.nupkg'
    '*.nupkg.sha512'
    '*.nuspec'
    '*.psmdcp'
    '.rels'
    '`[Content_Types`].xml'
)

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-UpdateLog {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')

    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Output $line

    try {
        $logRoot = Join-Path -Path $InstallPath -ChildPath 'Logs'
        if (-not (Test-Path -LiteralPath $logRoot -PathType Container)) {
            New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
        }
        Add-Content -LiteralPath (Join-Path $logRoot 'PowerShellGalleryUpdate.log') -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning ("Unable to write update log. {0}" -f $_.Exception.Message)
    }
}

function Get-NormalizedThumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Get-LatestModuleInfo {
    param([Parameter(Mandatory = $true)][string]$Name)

    $module = Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if (-not $module) {
        throw "Installed module not found: $Name"
    }
    return $module
}

function Resolve-DeploymentModuleInfo {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$RequestedVersion,
        [bool]$PreferLatest
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedVersion)) {
        $parsedVersion = $null
        if (-not [version]::TryParse($RequestedVersion, [ref]$parsedVersion)) {
            throw "Invalid -ModuleVersion value: $RequestedVersion"
        }
        $module = Get-Module -ListAvailable -Name $Name |
            Where-Object { $_.Version -eq $parsedVersion } |
            Select-Object -First 1
        if (-not $module) {
            throw "Requested module version is not installed: $Name $RequestedVersion"
        }
        return $module
    }

    if (-not $PreferLatest) {
        # Deploy the module that contains this helper (Tools\ under the module base).
        $ownManifestPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath "$Name.psd1"
        if (Test-Path -LiteralPath $ownManifestPath -PathType Leaf) {
            $module = Test-ModuleManifest -Path $ownManifestPath -ErrorAction Stop
            if ($module.Name -eq $Name) {
                return $module
            }
        }
        Write-UpdateLog -Level WARN -Message 'Helper is not running from a module folder and no -ModuleVersion was supplied; deploying the newest installed version.'
    }

    return Get-LatestModuleInfo -Name $Name
}

function Invoke-GalleryPackageUpdate {
    param([Parameter(Mandatory = $true)][string]$Name)

    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $powerShellGetInstall = $null
    if (Get-Command Get-InstalledModule -ErrorAction SilentlyContinue) {
        $powerShellGetInstall = Get-InstalledModule -Name $Name -ErrorAction SilentlyContinue
    }

    if ($powerShellGetInstall -and (Get-Command Update-Module -ErrorAction SilentlyContinue)) {
        Write-UpdateLog "Updating $Name with PowerShellGet."
        $updateModuleParameters = @{
            Name        = $Name
            Force       = $true
            ErrorAction = 'Stop'
        }
        if ($IncludePrerelease) {
            $updateModuleParameters.AllowPrerelease = $true
        }
        Update-Module @updateModuleParameters
        return
    }

    if (Get-Command Update-PSResource -ErrorAction SilentlyContinue) {
        Write-UpdateLog "Updating $Name with PSResourceGet."
        $updateResourceParameters = @{
            Name            = $Name
            Repository      = 'PSGallery'
            Scope           = 'AllUsers'
            TrustRepository = $true
            Quiet           = $true
            ErrorAction     = 'Stop'
        }
        if ($IncludePrerelease) {
            $updateResourceParameters.Prerelease = $true
        }
        Update-PSResource @updateResourceParameters
        return
    }

    throw 'No supported PowerShell Gallery update command is available.'
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

function Assert-PinnedSignature {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $state = Get-PinnedSignatureState -Path $Path -Thumbprint $Thumbprint
    if (-not $state.Accepted) {
        throw ("Package signature validation failed: file={0}; status={1}; winVerifyTrust={2}; signer={3}" -f $Path,$state.Status,$state.TrustResult,$state.SignerThumbprint)
    }
    if (-not $state.ChainTrusted) {
        Write-UpdateLog -Level WARN -Message ("Pinned signer accepted but its certificate chain is not trusted on this device: file={0}; winVerifyTrust={1}" -f $Path,$state.TrustResult)
    }
}

function Get-FileCatalogDifference {
    param([Parameter(Mandatory = $true)][object]$CatalogResult)

    $differences = New-Object Collections.Generic.List[string]
    $catalogItems = $CatalogResult.CatalogItems
    $pathItems = $CatalogResult.PathItems
    foreach ($key in @($catalogItems.Keys)) {
        if (-not $pathItems.ContainsKey($key)) {
            $differences.Add("missing=$key")
        }
        elseif ([string]$pathItems[$key] -ne [string]$catalogItems[$key]) {
            $differences.Add("changed=$key")
        }
    }
    foreach ($key in @($pathItems.Keys)) {
        if (-not $catalogItems.ContainsKey($key)) {
            $differences.Add("extra=$key")
        }
    }
    return @($differences | Sort-Object)
}

function Assert-ModulePackageIntegrity {
    param(
        [Parameter(Mandatory = $true)][string]$ModuleBase,
        [Parameter(Mandatory = $true)][string]$Thumbprint,
        [switch]$SkipSignatureValidation
    )

    # The signed catalog covers every packaged file, including .cmd launchers,
    # templates, images and version metadata that cannot carry Authenticode.
    $catalogPath = Join-Path -Path $ModuleBase -ChildPath $script:PackageCatalogFileName
    if (-not (Test-Path -LiteralPath $catalogPath -PathType Leaf)) {
        throw "Package file catalog not found: $catalogPath"
    }
    if (-not $SkipSignatureValidation) {
        Assert-PinnedSignature -Path $catalogPath -Thumbprint $Thumbprint
    }

    $catalogResult = Test-FileCatalog `
        -Path $ModuleBase `
        -CatalogFilePath $catalogPath `
        -FilesToSkip $script:CatalogFilesToSkip `
        -Detailed
    if ([string]$catalogResult.Status -ne 'Valid') {
        $differences = @(Get-FileCatalogDifference -CatalogResult $catalogResult)
        throw ("Package file catalog validation failed: status={0}; differences={1}" -f
            $catalogResult.Status,(@($differences | Select-Object -First 20) -join ', '))
    }

    $powerShellFiles = @(Get-ChildItem -LiteralPath $ModuleBase -Recurse -File |
        Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
    if ($powerShellFiles.Count -eq 0) {
        throw "No PowerShell files found in package: $ModuleBase"
    }
    if (-not $SkipSignatureValidation) {
        foreach ($file in $powerShellFiles) {
            Assert-PinnedSignature -Path $file.FullName -Thumbprint $Thumbprint
        }
    }
}

function ConvertTo-ProcessArgument {
    param([AllowEmptyString()][string]$Value)

    # Quotes one argument for the MSVCRT/CommandLineToArgvW rules used by
    # powershell.exe -File: embedded quotes are escaped and every backslash run
    # that precedes a quote, including the closing quote, is doubled.
    $builder = New-Object Text.StringBuilder
    [void]$builder.Append('"')
    $backslashCount = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq [char]'\') {
            $backslashCount++
            continue
        }
        if ($character -eq [char]'"') {
            [void]$builder.Append([char]'\', (2 * $backslashCount) + 1)
        }
        elseif ($backslashCount -gt 0) {
            [void]$builder.Append([char]'\', $backslashCount)
        }
        [void]$builder.Append($character)
        $backslashCount = 0
    }
    if ($backslashCount -gt 0) {
        [void]$builder.Append([char]'\', 2 * $backslashCount)
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Get-GalleryUpdateTaskArgumentString {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$TargetModuleName,
        [Parameter(Mandatory = $true)][string]$TargetInstallPath,
        [Parameter(Mandatory = $true)][string]$GuiTaskPath,
        [Parameter(Mandatory = $true)][string]$GuiTaskName,
        [Parameter(Mandatory = $true)][int]$GuiRepeatIntervalMinutes,
        [Parameter(Mandatory = $true)][string]$ScheduledTaskPath,
        [Parameter(Mandatory = $true)][string]$ScheduledTaskName,
        [Parameter(Mandatory = $true)][int]$IntervalHours,
        [bool]$Prerelease,
        [Parameter(Mandatory = $true)][string]$SignerThumbprint
    )

    # Switches are passed bare: powershell.exe -File cannot convert the strings
    # '$true'/'$false' to [bool] parameters.
    $arguments = New-Object Collections.Generic.List[string]
    foreach ($token in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File')) {
        $arguments.Add($token)
    }
    $arguments.Add((ConvertTo-ProcessArgument $ScriptPath))
    $arguments.Add('-ModuleName')
    $arguments.Add((ConvertTo-ProcessArgument $TargetModuleName))
    $arguments.Add('-InstallPath')
    $arguments.Add((ConvertTo-ProcessArgument $TargetInstallPath))
    $arguments.Add('-TaskPath')
    $arguments.Add((ConvertTo-ProcessArgument $GuiTaskPath))
    $arguments.Add('-TaskName')
    $arguments.Add((ConvertTo-ProcessArgument $GuiTaskName))
    $arguments.Add('-RepeatIntervalMinutes')
    $arguments.Add([string]$GuiRepeatIntervalMinutes)
    $arguments.Add('-EnableAutomaticUpdate')
    $arguments.Add('-UpdateTaskPath')
    $arguments.Add((ConvertTo-ProcessArgument $ScheduledTaskPath))
    $arguments.Add('-UpdateTaskName')
    $arguments.Add((ConvertTo-ProcessArgument $ScheduledTaskName))
    $arguments.Add('-UpdateIntervalHours')
    $arguments.Add([string]$IntervalHours)
    if ($Prerelease) {
        $arguments.Add('-IncludePrerelease')
    }
    $arguments.Add('-ExpectedSignerThumbprint')
    $arguments.Add((ConvertTo-ProcessArgument $SignerThumbprint))
    return ($arguments -join ' ')
}

function Register-GalleryUpdateTask {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$ScheduledTaskPath,
        [Parameter(Mandatory = $true)][string]$ScheduledTaskName,
        [Parameter(Mandatory = $true)][int]$IntervalHours
    )

    $powerShellPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) {
        $powerShellPath = 'powershell.exe'
    }

    $argumentString = Get-GalleryUpdateTaskArgumentString `
        -ScriptPath $ScriptPath `
        -TargetModuleName $ModuleName `
        -TargetInstallPath $InstallPath `
        -GuiTaskPath $TaskPath `
        -GuiTaskName $TaskName `
        -GuiRepeatIntervalMinutes $RepeatIntervalMinutes `
        -ScheduledTaskPath $ScheduledTaskPath `
        -ScheduledTaskName $ScheduledTaskName `
        -IntervalHours $IntervalHours `
        -Prerelease ([bool]$IncludePrerelease) `
        -SignerThumbprint $ExpectedSignerThumbprint

    $action = New-ScheduledTaskAction -Execute $powerShellPath -Argument $argumentString
    $now = Get-Date
    $firstRun = $now.Date.AddHours(3)
    if ($firstRun -le $now) {
        $firstRun = $firstRun.AddDays(1)
    }
    $trigger = New-ScheduledTaskTrigger `
        -Once `
        -At $firstRun `
        -RepetitionInterval (New-TimeSpan -Hours $IntervalHours) `
        -RepetitionDuration (New-TimeSpan -Days 3650)
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1)

    Register-ScheduledTask `
        -TaskPath $ScheduledTaskPath `
        -TaskName $ScheduledTaskName `
        -Action $action `
        -Trigger $trigger `
        -Principal $principal `
        -Settings $settings `
        -Description 'Updates SmartM365 Device Reboot Manager from PowerShell Gallery.' `
        -Force | Out-Null
}

if (-not (Test-IsAdministrator)) {
    throw 'Device Reboot Manager Gallery deployment must run elevated.'
}

try {
    Write-UpdateLog "Starting Device Reboot Manager Gallery deployment. Module=$ModuleName; SkipPackageUpdate=$SkipPackageUpdate; RequestedVersion=$ModuleVersion"

    if (-not [string]::IsNullOrWhiteSpace($ModuleVersion) -and -not $SkipPackageUpdate) {
        throw '-ModuleVersion selects an installed version and requires -SkipPackageUpdate.'
    }

    if (-not $SkipPackageUpdate) {
        Invoke-GalleryPackageUpdate -Name $ModuleName
    }

    $moduleInfo = Resolve-DeploymentModuleInfo `
        -Name $ModuleName `
        -RequestedVersion $ModuleVersion `
        -PreferLatest (-not $SkipPackageUpdate)
    $moduleBase = $moduleInfo.ModuleBase
    Write-UpdateLog ("Selected module version {0} from {1}" -f $moduleInfo.Version,$moduleBase)
    Assert-ModulePackageIntegrity -ModuleBase $moduleBase -Thumbprint $ExpectedSignerThumbprint

    $latestToolPath = Join-Path -Path $moduleBase -ChildPath 'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
    if (-not $SkipPackageUpdate -and $latestToolPath -ne $PSCommandPath) {
        & $latestToolPath `
            -ModuleName $ModuleName `
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
            -ExpectedSignerThumbprint $ExpectedSignerThumbprint `
            -ModuleVersion ([string]$moduleInfo.Version) `
            -SkipPackageUpdate
        return
    }

    $installerPath = Join-Path -Path $moduleBase -ChildPath 'Runtime\Deploy\SmartM365-DeviceRebootManager-Install.ps1'
    if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
        throw "Runtime installer not found in package: $installerPath"
    }

    $installParameters = @{
        InstallPath           = $InstallPath
        ConfigSourcePath      = $ConfigSourcePath
        ForceConfig           = $ForceConfig
        SkipScheduledTask     = $SkipScheduledTask
        TaskPath              = $TaskPath
        TaskName              = $TaskName
        RepeatIntervalMinutes = $RepeatIntervalMinutes
        PackageSource         = 'PowerShellGallery'
        UpdateTaskPath        = $UpdateTaskPath
        UpdateTaskName        = $UpdateTaskName
    }
    & $installerPath @installParameters

    $installedUpdateScript = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
    Copy-Item -LiteralPath $latestToolPath -Destination $installedUpdateScript -Force

    if ($EnableAutomaticUpdate) {
        Register-GalleryUpdateTask `
            -ScriptPath $installedUpdateScript `
            -ScheduledTaskPath $UpdateTaskPath `
            -ScheduledTaskName $UpdateTaskName `
            -IntervalHours $UpdateIntervalHours
        Write-UpdateLog ("Automatic update task registered: {0}{1}" -f $UpdateTaskPath,$UpdateTaskName)
    }
    else {
        $existingUpdateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
        if ($existingUpdateTask) {
            Unregister-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -Confirm:$false
        }
        Write-UpdateLog 'Automatic update task is disabled.'
    }

    $metadataPath = Join-Path $InstallPath 'SmartM365-DeviceRebootManager.installation.json'
    $metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
    $metadata | Add-Member -NotePropertyName ModuleName -NotePropertyValue $ModuleName -Force
    $metadata | Add-Member -NotePropertyName ModuleVersion -NotePropertyValue ([string]$moduleInfo.Version) -Force
    $metadata | Add-Member -NotePropertyName SignerThumbprint `
        -NotePropertyValue (Get-NormalizedThumbprint $ExpectedSignerThumbprint) -Force
    $metadata | Add-Member -NotePropertyName AutomaticUpdateEnabled `
        -NotePropertyValue ([bool]$EnableAutomaticUpdate) -Force
    $metadata | Add-Member -NotePropertyName UpdateIntervalHours `
        -NotePropertyValue $UpdateIntervalHours -Force
    $metadata | ConvertTo-Json -Depth 4 |
        Set-Content -LiteralPath $metadataPath -Encoding UTF8

    Write-UpdateLog ("Device Reboot Manager deployed successfully. PackageVersion={0}" -f $metadata.PackageVersion)
}
catch {
    Write-UpdateLog -Level ERROR -Message $_.Exception.Message
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCvni7Sfl+AXCu6
# cvj85hm3USY/pwuPfLBCC5tFH42SC6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFAYjZoF/n4+osx9ozbdmm2WFH6qZ+TZSC4QNbjRMOMTMA0GCSqG
# SIb3DQEBAQUABIIBgFRKxvx7YajpnisjLegGNRoRfidtT95snWPXhjC0V1L91i8r
# M7bq1Jr6FyOtrcFUgFZQ3VIubgzQjo4YCZxVRpMzM1sZ3CT6QWAfddxJgyGL5cTN
# RvUyZRYn0ULLA4geNPwzgoGTNcTs1NypmbuLFRPMD8g29iM/zbllv8jyPssRrGi5
# V0dU9Rz6Dqpgf2+ThPtmCrpYmZZxqQ0+Evn5imaLtCof2Hm4ocLyyu/yjwnbfC2D
# a1Rv/4NaeVIJ4WwCKEv666hEFadBN+ZsFinrCj/xhLovww0Hm/hmH4UDT+y3d56h
# n4ofVPoyVBanNWlML0TPp0bUAxPjgl3TUn0KNiV4riaY9ct6mccVNfKio0ZNPcKa
# M4RqEgLNPkNUhjy2S51NY8F2bypgN7nZIsVvrG8snJ1tRZ/BIumPZtdQ05aJqtwf
# POmWKoSmBoBubCpTjI38JWlmQam+FkyDEHglJ75hTLiohPOsNVIWBYUXDEZojmA/
# BwvUjRtvzrcoFVUocaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMjQy
# MjJaMC8GCSqGSIb3DQEJBDEiBCC2jD6etXlfsyaTRxlxrZidL+EtMPHucXSYYXIO
# E5I+pzANBgkqhkiG9w0BAQEFAASCAgCKmcWITdAhmLP9XJG264yVwInfsCed0QFG
# 0EtJcGeJmMsP4mQRkvFOyjQA5eFimN83tpFgwGaoAiQd/IDneNs4nQw7nEKZH1gq
# y2SM52TJSnxVMWJyA4rdNY78NDMT8/NBMO6aLXDySxAPb+u0ctM7tLZ0JY8WcHax
# VGEc/rY1BeK89e33pliVWEw9mtam+dNDFOVmBgod8cMYqIeNlvAMBrSTPpo9rhCk
# sW9sbaBTShZZ4prSYEqqo2xMgm4UBNwA2GHpFa5p0yog6aqmjptTbKD2bCfnjR/c
# MV4aoGrITCtNKwgtMXmdMt+5InddnSH3dtR+Jg9TeTmBubPT1ve2mn2LlacoBEew
# QlzaYZQCdu1FiyduciwenaCl2RfH4EV+mowlFcWTB0zhYP/VdAxCefxY8oOD1vwZ
# NgXRxVeXmcrDc+ny72cUcy+MDgefAdfm2YGvSwvj1Cu3R4u6Hq+TzGaVj8z6t7ca
# QL/rkiSk9BP9Q/ZfmRjZ2iR7wGAzR2nupgHPKIA2clAqCkMnJfI/ObGVT+j0+1zV
# BCYmqBebIe1DlJdtcYfj+5IgUcb2AN6wtZwEfDPIEj8ODrPkLnbekbnwtHg0Zclr
# RiUDhkb/s4/DmfFb46BRVQJmN53+Y/17t1d0VO6j93ZLTnLJ14RjAnx6+Efo5opg
# vZwb5BeksA==
# SIG # End signature block
