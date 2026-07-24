<#
.SYNOPSIS
Validates scoped endpoint run-guard retries and generated GUI CMD launchers.

.VERSION
1.0.0
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
$orchestrator = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1'
$gui = Join-Path $toolkitRoot 'Scripts\SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1'
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SmartM365-W11UT-RunGuardCmd-{0}' -f [guid]::NewGuid().ToString('N'))

try {
    Import-ScriptFunction -Path $orchestrator -Name 'ConvertTo-TechnicianRunGuardUtcDateTime'
    Import-ScriptFunction -Path $orchestrator -Name 'Test-TechnicianRunGuardEndpointBypassEligible'

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
    function global:Get-ConfiguredValue { param([string]$Name); return '' }
    $global:SyntheticInventoryRefreshFails = $false
    $global:SyntheticRootFallbackAccepted = $false
    function global:Show-GuiWarningYesNo { param([string]$Title, [string]$Message); return [bool]$global:SyntheticRootFallbackAccepted }
    $automaticToolkitRoot = Join-Path $testRoot 'AutomaticToolkit'
    $automaticScriptsRoot = Join-Path $automaticToolkitRoot 'Scripts'
    New-Item -ItemType Directory -Path $automaticScriptsRoot -Force | Out-Null
    [pscustomobject]@{ DeviceName = 'ROOT-AD'; Marker = 'ROOT' } | Export-Csv -LiteralPath (Join-Path $automaticToolkitRoot 'DevicesAD.csv') -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{ DeviceName = 'ROOT-INTUNE'; Marker = 'ROOT'; InventoryTenantProfile = 'prod'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath (Join-Path $automaticToolkitRoot 'DevicesIntune.csv') -NoTypeInformation -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $automaticScriptsRoot 'SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1') -Value '# synthetic exporter' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $automaticScriptsRoot 'SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1') -Value '# synthetic exporter' -Encoding UTF8
    $toolkitRoot = $automaticToolkitRoot

    function global:Invoke-GuiPowerShellProcess {
        param([string]$ScriptPath, [string[]]$Arguments, [string]$LogPath, [string]$Activity)
        if ($global:SyntheticInventoryRefreshFails) { throw 'Synthetic inventory refresh failure.' }
        $outputIndex = [array]::IndexOf($Arguments, '-OutputPath')
        if ($outputIndex -lt 0) { throw 'Synthetic exporter did not receive -OutputPath.' }
        $outputPath = $Arguments[$outputIndex + 1]
        if ($ScriptPath -like '*Export-ADDevicesCsv.ps1') {
            [pscustomobject]@{ DeviceName = 'REFRESHED-AD'; Marker = 'SNAPSHOT' } | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        }
        else {
            [pscustomobject]@{ DeviceName = 'REFRESHED-INTUNE'; Marker = 'SNAPSHOT'; InventoryTenantProfile = 'prod'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        }
        Set-Content -LiteralPath $LogPath -Value $Activity -Encoding UTF8
    }

    $automaticSnapshot = Get-AutomaticInventorySnapshot -Source Both -TenantProfile 'prod'
    Assert-True -Condition ($automaticSnapshot.AdInventoryCsv -ne (Join-Path $automaticToolkitRoot 'DevicesAD.csv')) -Message 'automatic AD selection does not silently reuse root CSV'
    Assert-True -Condition ($automaticSnapshot.IntuneInventoryCsv -ne (Join-Path $automaticToolkitRoot 'DevicesIntune.csv')) -Message 'automatic Intune selection does not silently reuse root CSV'
    Assert-Equal -Actual (Import-Csv -LiteralPath $automaticSnapshot.AdInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'REFRESHED-AD' -Message 'automatic AD snapshot was freshly generated'
    Assert-Equal -Actual (Import-Csv -LiteralPath $automaticSnapshot.IntuneInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'REFRESHED-INTUNE' -Message 'automatic Intune snapshot was freshly generated'

    $automaticRunPath = Join-Path $testRoot 'AutomaticRun'
    New-Item -ItemType Directory -Path $automaticRunPath -Force | Out-Null
    $copiedSnapshots = @(Copy-AutomaticInventorySnapshotToRun -InventoryContext $automaticSnapshot -RunPath $automaticRunPath)
    Assert-Equal -Actual $copiedSnapshots.Count -Expected 2 -Message 'both automatic snapshots were copied into the LOT run'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $automaticRunPath 'DevicesAD.csv') -PathType Leaf) -Message 'run-local AD snapshot exists'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $automaticRunPath 'DevicesIntune.csv') -PathType Leaf) -Message 'run-local Intune snapshot exists'

    $global:SyntheticInventoryRefreshFails = $true
    $fallbackRejected = $false
    try { Get-AutomaticInventorySnapshot -Source AD -TenantProfile 'prod' | Out-Null } catch { $fallbackRejected = $true }
    Assert-True -Condition $fallbackRejected -Message 'fresh root cache is not reused when explicit fallback is rejected'

    $global:SyntheticRootFallbackAccepted = $true
    $acceptedFallback = Get-AutomaticInventorySnapshot -Source AD -TenantProfile 'prod'
    Assert-Equal -Actual $acceptedFallback.AdInventoryCsv -Expected (Join-Path $automaticToolkitRoot 'DevicesAD.csv') -Message 'fresh root cache is used only after explicit fallback acceptance'
    Assert-True -Condition (($acceptedFallback.SourceDetails -join ' ') -match 'explicitly accepted') -Message 'accepted root fallback is marked in source details'
    $global:SyntheticInventoryRefreshFails = $false
    $global:SyntheticRootFallbackAccepted = $false
    Assert-Equal -Actual (ConvertTo-CmdArgument -Value 'C:\Program Files\SmartM365') -Expected '"C:\Program Files\SmartM365"' -Message 'CMD path argument quoting'
    Assert-Equal -Actual (ConvertTo-CmdSetCommand -Name 'W11UT_TEST' -Value 'value&safe') -Expected 'set "W11UT_TEST=value&safe"' -Message 'quoted SET command keeps ampersand literal'

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