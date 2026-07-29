<#
.SYNOPSIS
Validates scoped endpoint run-guard retries and generated GUI CMD launchers.

.VERSION
1.7.2
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ([string]$Actual -ne [string]$Expected) {
        throw ("ASSERT FAILED: {0}. Expected={1}; Actual={2}" -f $Message,$Expected,$Actual)
    }
}

function Import-ScriptFunction {
    param([string]$Path, [string]$Name)

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) { throw ($errors | ForEach-Object Message | Out-String) }
    $functionAst = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name }, $true) | Select-Object -First 1
    if (-not $functionAst) { throw "Function not found: $Name" }
    Set-Item -Path ("Function:\global:{0}" -f $Name) -Value $functionAst.Body.GetScriptBlock()
}

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$sourceToolkitRoot = $toolkitRoot
$orchestrator = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1'
$endpoint = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-Windows11UpgradeRepair.ps1'
$gui = Join-Path $toolkitRoot 'Scripts\SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1'
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SmartM365-W11UT-RunGuardCmd-{0}' -f [guid]::NewGuid().ToString('N'))

try {
    foreach ($functionName in @(
        'Get-SystemInstallLanguageTag',
        'Resolve-SetupLanguageRequirement',
        'Get-SetupMediaLanguages',
        'Resolve-SetupSourceMediaPath'
    )) {
        Import-ScriptFunction -Path $endpoint -Name $functionName
    }

    $global:SyntheticInstallLanguage = '040C'
    $global:SyntheticLanguageLogs = New-Object System.Collections.ArrayList
    function global:Get-RegistryValue {
        param([string]$Path, [string]$Name)
        if ($Name -eq 'InstallLanguage') { return $global:SyntheticInstallLanguage }
        return $null
    }
    function global:Get-WinSystemLocale {
        [CmdletBinding()]
        param()
        return [pscustomobject]@{ Name = 'fr-CH' }
    }
    function global:Write-SmartLog {
        param([string]$Message, [string]$Level = 'INFO')
        [void]$global:SyntheticLanguageLogs.Add(('{0}:{1}' -f $Level,$Message))
    }

    Assert-Equal -Actual (Resolve-SetupLanguageRequirement -RequestedLanguage 'MatchSystem') -Expected 'fr-FR' -Message 'MatchSystem prioritizes InstallLanguage 040C over SystemLocale fr-CH'
    Assert-True -Condition (@($global:SyntheticLanguageLogs | Where-Object { $_ -match 'InstallLanguage=040C; InstallCulture=fr-FR; InstalledUICulture=.*; SystemLocale=fr-CH' }).Count -gt 0) -Message 'language diagnostics retain InstallLanguage, UI culture, and SystemLocale evidence'

    $global:SyntheticInstallLanguage = ''
    Assert-Equal -Actual (Get-SystemInstallLanguageTag) -Expected ([System.Globalization.CultureInfo]::InstalledUICulture.Name) -Message 'InstalledUICulture is used when InstallLanguage is unavailable'

    $languageMediaRoot = Join-Path $testRoot 'language-media'
    $frMediaRoot = Join-Path $languageMediaRoot 'fr-FR'
    New-Item -ItemType Directory -Path (Join-Path $frMediaRoot 'sources') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $frMediaRoot 'sources\lang.ini') -Encoding ASCII -Value @(
        '[Available UI Languages]'
        'fr-FR = 3'
    )
    Assert-Equal -Actual (Resolve-SetupSourceMediaPath -SourcePath $languageMediaRoot -ExpectedLanguage 'fr-FR') -Expected $frMediaRoot -Message 'exact language in sources lang.ini selects the media folder'
    $strictLanguageMismatchRejected = $false
    try {
        Resolve-SetupSourceMediaPath -SourcePath $languageMediaRoot -ExpectedLanguage 'fr-CH' | Out-Null
    }
    catch {
        $strictLanguageMismatchRejected = $_.Exception.Message -match 'contains language fr-CH in sources\\lang\.ini'
    }
    Assert-True -Condition $strictLanguageMismatchRejected -Message 'sources lang.ini validation remains strict and rejects fr-CH when only fr-FR is present'

    Remove-Item -Path Function:\global:Get-RegistryValue -Force
    Remove-Item -Path Function:\global:Get-WinSystemLocale -Force
    Remove-Item -Path Function:\global:Write-SmartLog -Force
    Remove-Variable -Name SyntheticInstallLanguage -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name SyntheticLanguageLogs -Scope Global -ErrorAction SilentlyContinue

    Import-ScriptFunction -Path $orchestrator -Name 'ConvertTo-TechnicianRunGuardUtcDateTime'
    Import-ScriptFunction -Path $orchestrator -Name 'Test-TechnicianRunGuardEndpointBypassEligible'
    foreach ($functionName in @(
        'Get-ComputerListKey',
        'Test-HardwareNotCapableCycleResult',
        'Get-ComputerListMutexName',
        'Invoke-WithComputerListMutex',
        'Move-CycleStatusComputersFromList',
        'Move-HardwareNotCapableComputersFromList'
    )) {
        Import-ScriptFunction -Path $orchestrator -Name $functionName
    }


    $hardwareListRoot = Join-Path $testRoot 'hardware-list'
    New-Item -ItemType Directory -Path $hardwareListRoot -Force | Out-Null
    $hardwareListPath = Join-Path $hardwareListRoot 'Computers.txt'
    Set-Content -LiteralPath $hardwareListPath -Encoding ASCII -Value @(
        '# preserved comment'
        'FR-HARDWARE-001.fr.example.test'
        'FR-HARDWARE-002.fr.example.test'
        'FR-UNDETERMINED-001.fr.example.test'
        'FR-COMPAT-001.fr.example.test'
    )
    $remoteHardwareResult = [pscustomobject]@{
        ComputerName = 'FR-HARDWARE-001'
        LauncherStatus = 'ERROR'
        RemoteStatus = 'WINDOWS11_HARDWARE_NOT_CAPABLE'
    }
    $remoteHardwareMove = Move-HardwareNotCapableComputersFromList -ComputerListPath $hardwareListPath -CycleSummary @($remoteHardwareResult)
    Assert-Equal -Actual $remoteHardwareMove.Moved -Expected 1 -Message 'remote hardware-not-capable result is removed immediately'
    Assert-True -Condition ([string]$remoteHardwareMove.HardwareNotCapablePath -match 'ComputersHardwareNotCapable\.txt$') -Message 'hardware archive uses the dedicated file'
    $launcherHardwareResult = [pscustomobject]@{
        ComputerName = 'FR-HARDWARE-002.fr.example.test'
        LauncherStatus = 'WINDOWS11_HARDWARE_NOT_CAPABLE'
        RemoteStatus = ''
    }
    $launcherHardwareMove = Move-HardwareNotCapableComputersFromList -ComputerListPath $hardwareListPath -CycleSummary @($launcherHardwareResult)
    Assert-Equal -Actual $launcherHardwareMove.Moved -Expected 1 -Message 'launcher hardware-not-capable result is removed immediately'
    $remainingHardwareList = @(Get-Content -LiteralPath $hardwareListPath)
    Assert-True -Condition ($remainingHardwareList -notcontains 'FR-HARDWARE-001.fr.example.test') -Message 'short-name result matches and removes the FQDN from Computers.txt'
    Assert-True -Condition ($remainingHardwareList -contains 'FR-UNDETERMINED-001.fr.example.test') -Message 'undetermined readiness remains in Computers.txt'
    Assert-True -Condition ($remainingHardwareList -contains 'FR-COMPAT-001.fr.example.test') -Message 'other compatibility status remains in Computers.txt'
    $hardwareArchiveLines = @(Get-Content -LiteralPath $remoteHardwareMove.HardwareNotCapablePath)
    Assert-Equal -Actual $hardwareArchiveLines.Count -Expected 2 -Message 'hardware archive contains each removed computer once'
    $duplicateHardwareMove = Move-HardwareNotCapableComputersFromList -ComputerListPath $hardwareListPath -CycleSummary @($remoteHardwareResult)
    Assert-Equal -Actual $duplicateHardwareMove.Moved -Expected 0 -Message 'repeated result does not duplicate an archived computer'
    Assert-Equal -Actual @(Get-Content -LiteralPath $remoteHardwareMove.HardwareNotCapablePath).Count -Expected 2 -Message 'hardware archive remains deduplicated'
    $undeterminedResult = [pscustomobject]@{
        ComputerName = 'FR-UNDETERMINED-001'
        LauncherStatus = 'WINDOWS11_HARDWARE_READINESS_UNDETERMINED'
        RemoteStatus = 'WINDOWS11_HARDWARE_READINESS_UNDETERMINED'
    }
    Assert-Equal -Actual (Move-HardwareNotCapableComputersFromList -ComputerListPath $hardwareListPath -CycleSummary @($undeterminedResult)).Moved -Expected 0 -Message 'undetermined readiness is not removed'
    $compatibilityResult = [pscustomobject]@{
        ComputerName = 'FR-COMPAT-001'
        LauncherStatus = 'WINDOWS11_COMPAT_BLOCKER'
        RemoteStatus = 'WINDOWS11_COMPAT_BLOCKER'
    }
    Assert-Equal -Actual (Move-HardwareNotCapableComputersFromList -ComputerListPath $hardwareListPath -CycleSummary @($compatibilityResult)).Moved -Expected 0 -Message 'generic compatibility blocker is not removed'

    $dueUtc = (Get-Date).ToUniversalTime().AddMinutes(-1).ToString('o')
    $futureUtc = (Get-Date).ToUniversalTime().AddMinutes(10).ToString('o')
    $dueSetup = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'SetupFailure'; RemoteStatus = 'DIRECT_SETUP_UPGRADE_FAILED' }
    Assert-True -Condition (Test-TechnicianRunGuardEndpointBypassEligible -Entry $dueSetup) -Message 'due setup failure permits a one-attempt endpoint bypass'

    $futureSetup = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $futureUtc; FailureCategory = 'SetupFailure'; RemoteStatus = 'DIRECT_SETUP_UPGRADE_FAILED' }
    Assert-True -Condition (-not (Test-TechnicianRunGuardEndpointBypassEligible -Entry $futureSetup)) -Message 'future retry stays blocked'

    $networkFailure = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'NetworkTransient'; RemoteStatus = 'ADMIN_SHARE_UNREACHABLE' }
    Assert-True -Condition (-not (Test-TechnicianRunGuardEndpointBypassEligible -Entry $networkFailure)) -Message 'network failure never bypasses endpoint guard'

    $pendingReboot = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'OperatorAction'; RemoteStatus = 'PENDING_REBOOT_USER_CONNECTED' }
    Assert-True -Condition (-not (Test-TechnicianRunGuardEndpointBypassEligible -Entry $pendingReboot)) -Message 'pending reboot never bypasses endpoint guard'

    $ambiguousExecution = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'ExecutionTransient'; RemoteStatus = 'REMOTE_RESULT_STALE' }
    Assert-True -Condition (-not (Test-TechnicianRunGuardEndpointBypassEligible -Entry $ambiguousExecution)) -Message 'ambiguous execution state never bypasses endpoint guard'

    $endpointTransient = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'ExecutionTransient'; RemoteStatus = 'SETUP_CACHE_LOCKED' }
    Assert-True -Condition (Test-TechnicianRunGuardEndpointBypassEligible -Entry $endpointTransient) -Message 'recognized endpoint transient can retry after backoff'

    . $gui -ValidateOnly | Out-Null
    Import-ScriptFunction -Path $gui -Name 'Get-AutomaticInventoryFileInfo'
    Import-ScriptFunction -Path $gui -Name 'Get-AutomaticInventorySnapshot'
    Import-ScriptFunction -Path $gui -Name 'Get-GuiPowerShellProcessLogDetail'
    Import-ScriptFunction -Path $gui -Name 'Get-AutomaticGeneratedLotName'
    Import-ScriptFunction -Path $gui -Name 'Update-AutomaticGeneratedLotName'
    Import-ScriptFunction -Path $gui -Name 'New-AutomaticLotPreviewWork'
    Import-ScriptFunction -Path $gui -Name 'Get-AutomaticSourceSelection'
    Import-ScriptFunction -Path $gui -Name 'Test-AutomaticFilterSourceCompatibility'
    foreach ($functionName in @('Get-LotConfigPath','Get-LotsRoot','Get-SafeLotName','Invoke-LotWrapperRefresh','New-ToolkitLotFolder')) {
        Import-ScriptFunction -Path $gui -Name $functionName
    }

    $manualToolkitRoot = Join-Path $testRoot 'ManualLotToolkit'
    $manualScriptsRoot = Join-Path $manualToolkitRoot 'Scripts'
    New-Item -ItemType Directory -Path $manualScriptsRoot -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $manualScriptsRoot 'SmartM365-Windows11Upgrade-Update-LotCmdWrappers.ps1') -Encoding UTF8 -Value @(
        'param([string]$ToolkitRoot)'
        'Write-Output "Synthetic wrapper refresh chatter"'
        'exit 0'
    )
    $wrapperReturn = @(Invoke-LotWrapperRefresh -RootPath $manualToolkitRoot)
    Assert-Equal -Actual $wrapperReturn.Count -Expected 0 -Message 'successful wrapper refresh does not leak child process output'
    $manualLotResults = @(New-ToolkitLotFolder -RootPath $manualToolkitRoot -Name 'LOT-MANUAL-RETURN-TEST')
    Assert-Equal -Actual $manualLotResults.Count -Expected 1 -Message 'manual LOT creation returns exactly one result object'
    Assert-True -Condition ($null -ne $manualLotResults[0].PSObject.Properties['ComputersPath']) -Message 'manual LOT result exposes ComputersPath'
    Assert-True -Condition (Test-Path -LiteralPath $manualLotResults[0].ComputersPath -PathType Leaf) -Message 'manual LOT result points to the created Computers.txt'

    $successfulProgressState = [pscustomobject]@{
        Window = [pscustomobject]@{ IsVisible = $false }
        Result = $null
        ErrorRecord = $null
    }
    $successfulWork = New-AutomaticLotPreviewWork -ProgressState $successfulProgressState -Operation { 'SYNTHETIC_PREVIEW_OK' }
    & $successfulWork
    Assert-Equal -Actual $successfulProgressState.Result -Expected 'SYNTHETIC_PREVIEW_OK' -Message 'preview work closure preserves the progress state result property'
    Assert-True -Condition ($null -eq $successfulProgressState.ErrorRecord) -Message 'successful preview work closure does not create an error record'

    $failedProgressState = [pscustomobject]@{
        Window = [pscustomobject]@{ IsVisible = $false }
        Result = $null
        ErrorRecord = $null
    }
    $failedWork = New-AutomaticLotPreviewWork -ProgressState $failedProgressState -Operation { throw 'SYNTHETIC_PREVIEW_FAILURE' }
    & $failedWork
    Assert-True -Condition ($null -ne $failedProgressState.ErrorRecord) -Message 'failed preview work closure preserves the original error record'
    Assert-True -Condition ($failedProgressState.ErrorRecord.Exception.Message -match 'SYNTHETIC_PREVIEW_FAILURE') -Message 'failed preview work closure preserves the original error message'

    $global:controls = [pscustomobject]@{
        AutomaticLotNameText = [pscustomobject]@{ Text = '' }
        AutomaticNamePrefixText = [pscustomobject]@{ Text = '' }
        AutomaticSourceCombo = [pscustomobject]@{ SelectedItem = 'AD' }
        AutomaticNameContainsText = [pscustomobject]@{ Text = '' }
        AutomaticExcludeIntuneCheck = [pscustomobject]@{ IsChecked = $false }
        AutomaticExcludeStaleAdCheck = [pscustomobject]@{ IsChecked = $false }
        AutomaticLastLogonDaysText = [pscustomobject]@{ Text = '45'; IsEnabled = $false }
    }
    $script:AutomaticLotNameTimestamp = '20260726-120000'
    $script:AutomaticGeneratedLotName = ''
    Update-AutomaticGeneratedLotName
    Assert-Equal -Actual $global:controls.AutomaticLotNameText.Text -Expected 'LOT-AUTO-W10-20260726-120000' -Message 'automatic LOT name starts without a prefix segment'
    $global:controls.AutomaticNamePrefixText.Text = 'fr-'
    Update-AutomaticGeneratedLotName
    Assert-Equal -Actual $global:controls.AutomaticLotNameText.Text -Expected 'LOT-AUTO-W10-FR-20260726-120000' -Message 'automatic LOT name includes one normalized prefix'
    $global:controls.AutomaticNamePrefixText.Text = 'fr-;be-'
    Update-AutomaticGeneratedLotName
    Assert-Equal -Actual $global:controls.AutomaticLotNameText.Text -Expected 'LOT-AUTO-W10-FR-BE-20260726-120000' -Message 'automatic LOT name preserves multiple prefix order'
    $global:controls.AutomaticLotNameText.Text = 'LOT-MANUAL-PILOT'
    $global:controls.AutomaticNamePrefixText.Text = 'de-'
    Update-AutomaticGeneratedLotName
    Assert-Equal -Actual $global:controls.AutomaticLotNameText.Text -Expected 'LOT-MANUAL-PILOT' -Message 'manual LOT name is not overwritten by a prefix change'
    Test-AutomaticFilterSourceCompatibility
    $global:controls.AutomaticExcludeIntuneCheck.IsChecked = $true
    $intuneExclusionSourceRejected = $false
    try { Test-AutomaticFilterSourceCompatibility } catch { $intuneExclusionSourceRejected = ($_.Exception.Message -match 'AD \+ Intune') }
    Assert-True -Condition $intuneExclusionSourceRejected -Message 'Intune-presence exclusion requires the combined inventory source'
    $global:controls.AutomaticSourceCombo.SelectedItem = 'Both'
    Test-AutomaticFilterSourceCompatibility
    $global:controls.AutomaticExcludeIntuneCheck.IsChecked = $false
    $global:controls.AutomaticExcludeStaleAdCheck.IsChecked = $true
    $global:controls.AutomaticSourceCombo.SelectedItem = 'Intune'
    $staleAdSourceRejected = $false
    try { Test-AutomaticFilterSourceCompatibility } catch { $staleAdSourceRejected = ($_.Exception.Message -match 'AD or AD \+ Intune') }
    Assert-True -Condition $staleAdSourceRejected -Message 'AD LastLogon exclusion rejects an Intune-only source'
    $global:controls.AutomaticSourceCombo.SelectedItem = 'Both'
    Test-AutomaticFilterSourceCompatibility
    Remove-Variable -Name controls -Scope Global -Force
    function global:Get-ConfiguredValue { param([string]$Name); return '' }
    $global:SyntheticInventoryRefreshFails = $false
    $global:SyntheticRootFallbackAccepted = $false
    $global:SyntheticInventoryRefreshCalls = 0
    function global:Show-GuiWarningYesNo { param([string]$Title, [string]$Message); return [bool]$global:SyntheticRootFallbackAccepted }
    $automaticToolkitRoot = Join-Path $testRoot 'AutomaticToolkit'
    $automaticScriptsRoot = Join-Path $automaticToolkitRoot 'Scripts'
    New-Item -ItemType Directory -Path $automaticScriptsRoot -Force | Out-Null
    $rootAdPath = Join-Path $automaticToolkitRoot 'DevicesAD.csv'
    $rootIntunePath = Join-Path $automaticToolkitRoot 'DevicesIntune.csv'
    [pscustomobject]@{ ComputerName = 'ROOT-AD'; DeviceName = 'ROOT-AD'; Marker = 'ROOT'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0.19045' } | Export-Csv -LiteralPath $rootAdPath -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{ DeviceName = 'ROOT-INTUNE'; Marker = 'ROOT'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045'; InventoryTenantId = 'tenant-001'; InventoryAuthenticationMode = 'DelegatedInteractive'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $rootIntunePath -NoTypeInformation -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $automaticScriptsRoot 'SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1') -Value '# synthetic exporter' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $automaticScriptsRoot 'SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1') -Value '# synthetic exporter' -Encoding UTF8
    $toolkitRoot = $automaticToolkitRoot

    function global:Invoke-GuiPowerShellProcess {
        param([string]$ScriptPath, [string[]]$Arguments, [string]$LogPath, [string]$Activity, [switch]$Interactive, [scriptblock]$ProgressCallback)
        $global:SyntheticInventoryRefreshCalls++
        if ($global:SyntheticInventoryRefreshFails) { throw 'Synthetic inventory refresh failure.' }
        $outputIndex = [array]::IndexOf($Arguments, '-OutputPath')
        if ($outputIndex -lt 0) { throw 'Synthetic exporter did not receive -OutputPath.' }
        $outputPath = $Arguments[$outputIndex + 1]
        if ($ScriptPath -like '*Export-ADDevicesCsv.ps1') {
            [pscustomobject]@{ ComputerName = 'REFRESHED-AD'; DeviceName = 'REFRESHED-AD'; Marker = 'SNAPSHOT'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0.19045' } | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        }
        else {
            if (-not $Interactive) { throw 'Synthetic Intune export was not launched interactively.' }
            [pscustomobject]@{ DeviceName = 'REFRESHED-INTUNE'; Marker = 'SNAPSHOT'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045'; InventoryTenantId = 'tenant-001'; InventoryAuthenticationMode = 'DelegatedInteractive'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        }
        Set-Content -LiteralPath $LogPath -Value $Activity -Encoding UTF8
    }

    $automaticSnapshot = Get-AutomaticInventorySnapshot -Source Both
    Assert-True -Condition ($automaticSnapshot.AdInventoryCsv -ne $rootAdPath) -Message 'fresh AD root cache is copied to isolated evidence instead of used directly'
    Assert-True -Condition ($automaticSnapshot.IntuneInventoryCsv -ne $rootIntunePath) -Message 'fresh Intune root cache is copied to isolated evidence instead of used directly'
    Assert-Equal -Actual (Import-Csv -LiteralPath $automaticSnapshot.AdInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'ROOT-AD' -Message 'fresh AD root cache content is reused'
    Assert-Equal -Actual (Import-Csv -LiteralPath $automaticSnapshot.IntuneInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'ROOT-INTUNE' -Message 'fresh Intune root cache content is reused'
    Assert-Equal -Actual $global:SyntheticInventoryRefreshCalls -Expected 0 -Message 'fresh root caches avoid all exporter calls'
    Assert-True -Condition (($automaticSnapshot.SourceDetails -join ' ') -match 'reused verified recent root cache') -Message 'automatic preview reports root cache reuse'
    Assert-Equal -Actual $automaticSnapshot.TenantId -Expected 'tenant-001' -Message 'reused Intune cache preserves tenant provenance'
    Assert-Equal -Actual $automaticSnapshot.AuthenticationMode -Expected 'DelegatedInteractive' -Message 'reused Intune cache preserves delegated authentication'

    $automaticRunPath = Join-Path $testRoot 'AutomaticRun'
    New-Item -ItemType Directory -Path $automaticRunPath -Force | Out-Null
    $copiedSnapshots = @(Copy-AutomaticInventorySnapshotToRun -InventoryContext $automaticSnapshot -RunPath $automaticRunPath)
    Assert-Equal -Actual $copiedSnapshots.Count -Expected 2 -Message 'both automatic snapshots were copied into the LOT run'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $automaticRunPath 'DevicesAD.csv') -PathType Leaf) -Message 'run-local AD snapshot exists'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $automaticRunPath 'DevicesIntune.csv') -PathType Leaf) -Message 'run-local Intune snapshot exists'

    (Get-Item -LiteralPath $rootAdPath).LastWriteTime = (Get-Date).AddHours(-13)
    $staleAdSnapshot = Get-AutomaticInventorySnapshot -Source AD
    Assert-Equal -Actual $global:SyntheticInventoryRefreshCalls -Expected 1 -Message 'stale AD root cache triggers one targeted refresh'
    Assert-Equal -Actual (Import-Csv -LiteralPath $staleAdSnapshot.AdInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'REFRESHED-AD' -Message 'stale AD cache is replaced only in isolated evidence'
    Assert-Equal -Actual (Import-Csv -LiteralPath $rootAdPath | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'ROOT-AD' -Message 'AD root cache remains read-only'
    (Get-Item -LiteralPath $rootAdPath).LastWriteTime = Get-Date

    $forcedSnapshot = Get-AutomaticInventorySnapshot -Source AD -ForceRefresh
    Assert-Equal -Actual $global:SyntheticInventoryRefreshCalls -Expected 2 -Message 'force refresh bypasses a fresh AD root cache'
    Assert-Equal -Actual (Import-Csv -LiteralPath $forcedSnapshot.AdInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'REFRESHED-AD' -Message 'forced AD refresh uses isolated output'

    $global:SyntheticInventoryRefreshFails = $true
    $forcedFallbackRejected = $false
    try { Get-AutomaticInventorySnapshot -Source AD -ForceRefresh | Out-Null } catch { $forcedFallbackRejected = $true }
    Assert-True -Condition $forcedFallbackRejected -Message 'failed forced refresh remains blocked when fresh root fallback is rejected'

    $global:SyntheticRootFallbackAccepted = $true
    $acceptedFallback = Get-AutomaticInventorySnapshot -Source AD -ForceRefresh
    Assert-True -Condition ($acceptedFallback.AdInventoryCsv -ne $rootAdPath) -Message 'accepted root fallback is copied to isolated evidence'
    Assert-Equal -Actual (Import-Csv -LiteralPath $acceptedFallback.AdInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'ROOT-AD' -Message 'accepted fallback preserves root cache content'
    Assert-True -Condition (($acceptedFallback.SourceDetails -join ' ') -match 'explicitly accepted') -Message 'accepted root fallback is marked in source details'

    [pscustomobject]@{ DeviceName = 'ROOT-INTUNE'; Marker = 'ROOT'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045'; InventoryTenantId = 'tenant-001'; InventoryAuthenticationMode = 'AppOnly'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $rootIntunePath -NoTypeInformation -Encoding UTF8
    $refreshCallsBeforeInvalidIntune = $global:SyntheticInventoryRefreshCalls
    $appOnlyCacheRejected = $false
    try { Get-AutomaticInventorySnapshot -Source Intune | Out-Null } catch { $appOnlyCacheRejected = $true }
    Assert-True -Condition $appOnlyCacheRejected -Message 'app-only Intune root cache is rejected'
    Assert-Equal -Actual $global:SyntheticInventoryRefreshCalls -Expected ($refreshCallsBeforeInvalidIntune + 1) -Message 'invalid Intune provenance triggers an interactive refresh attempt'

    [pscustomobject]@{ DeviceName = 'ROOT-INTUNE'; Marker = 'ROOT'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045'; InventoryTenantId = 'tenant-001'; InventoryAuthenticationMode = 'DelegatedInteractive'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $rootIntunePath -NoTypeInformation -Encoding UTF8
    $refreshCallsBeforeDelegatedReuse = $global:SyntheticInventoryRefreshCalls
    $delegatedCache = Get-AutomaticInventorySnapshot -Source Intune
    Assert-True -Condition ($delegatedCache.IntuneInventoryCsv -ne $rootIntunePath) -Message 'delegated Intune root cache is copied to isolated evidence'
    Assert-Equal -Actual $delegatedCache.AuthenticationMode -Expected 'DelegatedInteractive' -Message 'delegated cache authentication provenance is preserved'
    Assert-Equal -Actual $global:SyntheticInventoryRefreshCalls -Expected $refreshCallsBeforeDelegatedReuse -Message 'delegated fresh Intune root cache avoids Graph refresh'

    $emptyStdout = Join-Path $testRoot 'empty.stdout.log'
    $emptyStderr = Join-Path $testRoot 'empty.stderr.log'
    [IO.File]::WriteAllText($emptyStdout, '')
    [IO.File]::WriteAllText($emptyStderr, '')
    Assert-Equal -Actual (Get-GuiPowerShellProcessLogDetail -Paths @($emptyStderr, $emptyStdout)) -Expected '' -Message 'empty inventory logs are read without a null Trim failure'

    $invalidAdCache = Join-Path $testRoot 'invalid-ad-cache.csv'
    [pscustomobject]@{ Marker = 'MISSING_REQUIRED_COLUMNS' } | Export-Csv -LiteralPath $invalidAdCache -NoTypeInformation -Encoding UTF8
    $invalidAdCacheInfo = Get-AutomaticInventoryFileInfo -Path $invalidAdCache -FreshnessHours 12 -SourceName 'AD'
    Assert-True -Condition (-not [bool]$invalidAdCacheInfo.Fresh) -Message 'recent AD cache without identity and OS columns is rejected'
    Assert-True -Condition (-not [bool]$invalidAdCacheInfo.ContentVerified) -Message 'invalid AD cache content is reported explicitly'

    $guiText = Get-Content -LiteralPath $gui -Raw
    Assert-True -Condition ($guiText -match 'InventoryProgressBar.{0,160}IsIndeterminate="True"') -Message 'automatic inventory wait window uses an indeterminate progress bar'
    Assert-True -Condition ($guiText -match 'AutomaticForceRefreshCheck') -Message 'automatic inventory force-refresh option is present'
    Assert-True -Condition ($guiText -match 'Update-AutomaticGeneratedLotName') -Message 'automatic computer-prefix changes update the generated LOT name'
    Assert-True -Condition ($guiText -match 'AutomaticNameContainsText') -Message 'automatic literal contains filter is present'
    Assert-True -Condition ($guiText -match 'AutomaticExcludeIntuneCheck') -Message 'automatic Intune-presence exclusion is present'
    Assert-True -Condition ($guiText -match 'AutomaticExcludeStaleAdCheck') -Message 'automatic AD LastLogon exclusion is present'
    Assert-True -Condition ($guiText -match 'AutomaticLastLogonDaysText.+Text="45"') -Message 'automatic AD LastLogon threshold defaults to 45 days'
    Assert-True -Condition ($guiText -match 'x:Name="AutomaticCreateButton"\s+Content="Create"') -Message 'automatic LOT action is labelled Create'
    Assert-True -Condition ($guiText -notmatch 'AutomaticModeCombo|Create and launch') -Message 'automatic LOT tab no longer exposes launch controls'
    $automaticCreateHandler = [regex]::Match($guiText, '(?s)\$controls\.AutomaticCreateButton\.Add_Click\(\{(?<Body>.*?)\r?\n\}\)\r?\n\$script:SyncingGlobalLimitText')
    Assert-True -Condition $automaticCreateHandler.Success -Message 'automatic LOT Create handler is present'
    Assert-True -Condition ($automaticCreateHandler.Groups['Body'].Value -match 'Invoke-AutomaticLotCreateWithProgress') -Message 'automatic LOT Create handler opens the creation progress workflow'
    Assert-True -Condition ($automaticCreateHandler.Groups['Body'].Value -notmatch 'Start-ToolkitLot|Confirm-UnlimitedCycleLaunch|Test-SetupSourceBeforeLaunch|Get-ToolkitOptionEnvironment') -Message 'automatic LOT Create handler cannot launch the LOT'
    $automaticCreateProgressFunction = [regex]::Match($guiText, '(?s)function Invoke-AutomaticLotCreateWithProgress\s*\{(?<Body>.*?)\r?\n\}\r?\n\r?\nfunction Confirm-AutomaticLotCreate')
    Assert-True -Condition $automaticCreateProgressFunction.Success -Message 'automatic LOT creation progress workflow is present'
    Assert-True -Condition ($automaticCreateProgressFunction.Groups['Body'].Value -match 'Invoke-AutomaticLotSelection.+-Create.+-ProgressCallback') -Message 'creation progress workflow creates the LOT with live stage callbacks'
    Assert-True -Condition ($automaticCreateProgressFunction.Groups['Body'].Value -match 'Refresh-LotList') -Message 'creation progress includes the final LOT-list refresh'
    $automaticSelectionFunction = [regex]::Match($guiText, '(?s)function Invoke-AutomaticLotSelection\s*\{(?<Body>.*?)\r?\n\}\r?\n\r?\nfunction Format-AutomaticLotSummary')
    Assert-True -Condition $automaticSelectionFunction.Success -Message 'automatic LOT selection function is present'
    Assert-True -Condition ($automaticSelectionFunction.Groups['Body'].Value -match 'ComputerNameContains|ExcludeIntunePresent|ExcludeStaleAd|AdLastLogonMaxAgeDays') -Message 'automatic LOT selection passes every advanced filter to the engine'
    Assert-True -Condition ($automaticSelectionFunction.Groups['Body'].Value -notmatch 'SkipWrapperRefresh') -Message 'automatic Create lets the engine generate launch wrappers'
    Assert-True -Condition ($automaticSelectionFunction.Groups['Body'].Value -match 'engineOutput[\s\S]+engineResults') -Message 'automatic LOT GUI isolates the engine Summary result from incidental console output'
    Assert-True -Condition ($automaticSelectionFunction.Groups['Body'].Value -match 'ProgressCallback') -Message 'automatic LOT GUI forwards progress callbacks to the engine'
    $automaticEngineText = Get-Content -LiteralPath (Join-Path $sourceToolkitRoot 'Scripts\SmartM365-Windows11Upgrade-New-AutomaticLot.ps1') -Raw
    Assert-True -Condition ($automaticEngineText -match "Refreshing LOT command wrappers") -Message 'automatic LOT engine reports the long wrapper-refresh stage'

    $global:SyntheticInventoryRefreshFails = $false
    $global:SyntheticRootFallbackAccepted = $false
    Assert-Equal -Actual (ConvertTo-CmdArgument -Value 'C:\Program Files\SmartM365') -Expected '"C:\Program Files\SmartM365"' -Message 'CMD path argument quoting'
    Assert-Equal -Actual (ConvertTo-CmdSetCommand -Name 'W11UT_TEST' -Value 'value&safe') -Expected 'set "W11UT_TEST=value&safe"' -Message 'quoted SET command keeps ampersand literal'

    $exporterText = Get-Content -LiteralPath (Join-Path $sourceToolkitRoot 'Scripts\SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1') -Raw
    Assert-True -Condition ($exporterText -match 'DeviceManagementManagedDevices\.Read\.All') -Message 'Intune exporter requests the delegated read scope'
    Assert-True -Condition ($exporterText -notmatch 'Initialize-SmartM365TenantContext|CertificateThumbprint|\[string\]\$Tenant\s*=') -Message 'Intune exporter has no app-only tenant profile dependency'
    Assert-True -Condition ($exporterText -match 'InventoryAuthenticationMode') -Message 'Intune exporter records authentication provenance'
    $rootCmdText = Get-Content -LiteralPath (Join-Path $sourceToolkitRoot 'Export-IntuneDevicesCsv.cmd') -Raw
    Assert-True -Condition ($rootCmdText -notmatch 'W11UT_INTUNE_TENANT_PROFILE|Profile\s*:\s*test') -Message 'root Intune CMD does not force the test profile'
    $orchestratorText = Get-Content -LiteralPath $orchestrator -Raw
    Assert-True -Condition ($orchestratorText -notmatch 'IntuneTenantProfile|InventoryTenantProfile') -Message 'LOT orchestrator has no app-only profile dependency'
    Assert-True -Condition ((Get-Content -LiteralPath $gui -Raw) -match 'W11UT_SKIP_INTUNE_INVENTORY_REFRESH.{0,200}=\s*''1''') -Message 'automatic LOT launch freezes the interactive Intune snapshot for the run'

    $unsafeRejected = $false
    try { ConvertTo-CmdSetCommand -Name 'W11UT_TEST' -Value '%TEMP%' | Out-Null } catch { $unsafeRejected = $true }
    Assert-True -Condition $unsafeRejected -Message 'percent expansion is rejected in generated CMD values'

    $workingDirectory = Join-Path $testRoot 'Run With Spaces'
    $evidenceDirectory = Join-Path $workingDirectory 'Logs'
    New-Item -ItemType Directory -Path $evidenceDirectory -Force | Out-Null
    $launchPath = New-GuiLaunchCommandFile -WorkingDirectory $workingDirectory -Commands @('echo CMD_TEST_OK') -NamePrefix 'Synthetic' -EvidenceDirectory $evidenceDirectory
    Assert-True -Condition (Test-Path -LiteralPath $launchPath -PathType Leaf) -Message 'temporary GUI launcher exists'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $evidenceDirectory 'GuiLaunchCommand.cmd') -PathType Leaf) -Message 'run-local GUI launcher evidence exists'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $evidenceDirectory 'GuiLaunchCommandEvidence.txt') -PathType Leaf) -Message 'GUI launcher evidence metadata exists'

    $cmdArguments = '/d /s /c ""{0}""' -f $launchPath
    $cmdProcess = Start-Process -FilePath $env:ComSpec -ArgumentList $cmdArguments -WorkingDirectory $workingDirectory -Wait -PassThru
    Assert-Equal -Actual $cmdProcess.ExitCode -Expected 0 -Message 'generated CMD launcher execution through Start-Process quoting'

    Write-Output 'SmartM365 Windows 11 run guard and CMD synthetic tests passed.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCm+4Ep+z2s+mzC
# spesKvCfGCHPSvEHT8PaGIXrquBy/KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGqUjdVQTtuKi7NQYGosfRkoZLislQndGVU7RKfNSyo8MA0GCSqG
# SIb3DQEBAQUABIIBgGFvLvf+TwybFbqhr5fA6ceovrz6S/3lFqvwOsGTdJk7iREf
# sjEVF+IeqLaEpxOFq9MqZav4RgMQjGy9tiVKBZQN7AqNhgaxVBkVKvElAUGooQZt
# Hl8yRAjJlZvoTBzoki/oDzFpXQdnFbJQTWEkvt+RI6xo1+PcnxgVpw/SJQ4zfqe4
# /HGwdUvr6R97Q+SsL6yIv9ugCPn4OydG7oDH3stMPPdCMMY7Mv3w1jd6Ck7kspaF
# fknWiS1Vfff+nRr4uMgAJAXgrQfe4t8Z/pCK9eH6GHgRkuMM6v/nJ6L2U3rmK/R8
# oa88Pteh1SMWNW05RvndCEitK99X//oD+WTYBn5pW1hAatxWNUn3RllLj1ie6KVS
# TncT/EcivkC1b0gcjP4RQ55J3+QhhhNCUdFwS2VlVVohmFvvAMp0vZvgVsesDmhU
# zmuLcgFZjznzGWd9L7aTmwyCXwoCMQO6DAdz2SFkDfSwMVP+rpglbo3nm4OvxxSn
# 0uy3unOWbqFz9H7BI6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjkyMDEy
# MjVaMC8GCSqGSIb3DQEJBDEiBCBkF3nwsIxNPz/3jufHp8fEK8sX3BLcfx5i7LzL
# 37UrJDANBgkqhkiG9w0BAQEFAASCAgCnUh0ZPxByVGFrmscfvkVxk59GT9tIfmim
# OoD+XDFU7JxdVYWRZuTKAnxiAkZvVn0R7AuoWUY6Svt0W31NmDwznm5YPU4NrlNz
# ElhUQM4ElWW5CezMyrKAOBTLNMJms4JGwbPqQriskKao29/2u5wBzCXHq2F8G2JY
# 1QFabKsLL+QA2ILtHRov8RGE6WS1oTPcO8Sstt0IRmun86X3vmsD6isyxI/jcRdo
# e6CHyL9H+8ZBfPxifafiZWxDcb+9niRhNL21O7malz2B2BX5DGGKmakD3T0WQQqD
# Fz78r3iYIS916zGRkDyD7rygTNk/v079/Qi8sJ/lUT4gJh30X0m+g9Vhddn9z+Vl
# i3xD4emYSKaXqAkZt7Jg5XiAnvhSVFQV21TIotiEpnOkKJM6U5jUPC8jXWztR+xs
# CHbv6zNXLtwZwvjeAb9yeUkHP0XuFgogqT4Bk6dCGGeJKa7jboi4HmQQV5nR6IFt
# PcQj9eAo+NY4wSJtvYc1GIKyA++8c5vimTGESJjXjPmhLhNypFBJjjE/1j7O7rvQ
# 5ijpwFzxxx7zCTCC1NpHDOFXUV7wNGnwWZaiGbncnlzniM2CIHzXPsMRVHo1Lmno
# QFGyv+oXDoH3n14Y8jfQzMX+mnUSIGojj/TzwohSdh56SP86uvC/o4X7zXyB6zbV
# MGpLvfV8DA==
# SIG # End signature block
