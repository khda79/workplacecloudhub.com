<#
.SYNOPSIS
    Offline tests of the legacy configuration import (Scripts\SmartM365-EndpointMigration-Import-LegacyConfig.ps1).
.DESCRIPTION
    Synthetic previous-toolkit folders (fictional values only) and a copy of the SEMT Config and
    Modules folders are created under %TEMP%. The import runs as a child process of the same
    PowerShell edition, with process-scope environment variables only.
    1. Preview writes nothing and leaves the previous toolkit files unchanged.
    2. Apply writes the global file and the per-Campaign GUI state with the expected placement,
       the values are accepted by the SEMT option resolution, and a report is written.
    3. Decisions: previous GUI defaults, launch choices, values without equivalent, unknown and
       invalid values, the old DISM alias, LOT values only counted, setup path warnings.
    4. A second Apply writes nothing; an explicit SEMT value is kept unless
       -ReplaceExistingValues, which backs the file up first; -WhatIf writes nothing;
       -IncludeLegacyDefaults carries previous GUI defaults; a missing legacy folder fails.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$packageRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$importScript = Join-Path $packageRoot 'Scripts\SmartM365-EndpointMigration-Import-LegacyConfig.ps1'
$powerShellExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Script)
    try { Write-TestResult -Condition ([bool](& $Script)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

function Invoke-Child {
    param([string[]]$Arguments)
    # Windows PowerShell 5.1 turns child stderr into a terminating error under Stop.
    $ErrorActionPreference = 'Continue'
    $output = & $powerShellExe @Arguments 2>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
}

function Invoke-Import {
    param([string[]]$Arguments,[string]$LegacyRoot = $script:LegacyRoot)
    return (Invoke-Child -Arguments (@('-NoProfile','-ExecutionPolicy','Bypass','-File',$importScript,'-PackageRoot',$script:SemtRoot,'-LegacyDevicesRoot',$LegacyRoot,'-LegacyEnvironmentScope','Process') + $Arguments))
}

function Get-TreeHash {
    param([string]$Path)
    return (@(Get-ChildItem -LiteralPath $Path -Recurse -File | Sort-Object FullName | ForEach-Object { '{0}={1}' -f $_.FullName,(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join ';')
}

function Read-Json {
    param([string]$Path)
    return ([System.IO.File]::ReadAllText($Path) | ConvertFrom-Json)
}

function Get-Row {
    param($Report,[string]$Campaign,[string]$Key)
    return @($Report.Rows | Where-Object { $_.Campaign -eq $Campaign -and $_.LegacyKey -eq $Key })[0]
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SemtImportLegacy-' + [guid]::NewGuid().ToString('N'))
$script:LegacyRoot = Join-Path $tempRoot 'Legacy\Devices'
$script:SemtRoot = Join-Path $tempRoot 'Semt'
try {
    New-Item -ItemType Directory -Path (Join-Path $script:SemtRoot 'Config') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $packageRoot 'Config\SmartEndpointMigrationToolkit.options.json.txt') -Destination (Join-Path $script:SemtRoot 'Config')
    Copy-Item -LiteralPath (Join-Path $packageRoot 'Config\SmartEndpointMigrationToolkit.local.json.template') -Destination (Join-Path $script:SemtRoot 'Config')
    Copy-Item -LiteralPath (Join-Path $packageRoot 'Modules') -Destination $script:SemtRoot -Recurse
    $configRoot = Join-Path $script:SemtRoot 'Config'

    $hjFolder = Join-Path $script:LegacyRoot 'IntuneHybridJoinToolkit'
    $w11Folder = Join-Path $script:LegacyRoot 'Windows11UpgradeToolkit'
    New-Item -ItemType Directory -Path (Join-Path $hjFolder 'Lots\LOT-TEMPLATE'),(Join-Path $w11Folder 'Lots\LOT-A'),(Join-Path $w11Folder 'Lots\LOT-TEMPLATE'),(Join-Path $w11Folder 'SetupSourceGates') -Force | Out-Null
    # The previous Hybrid Join GUI writes UTF-8 with BOM; the Windows 11 file is written without BOM here.
    [System.IO.File]::WriteAllLines((Join-Path $hjFolder 'IntuneHybridJoinToolkit.config'), [string[]]@(
        '# SmartM365 Intune Hybrid Join Toolkit GUI options.'
        'EHJIR_THROTTLE=20'
        'EHJIR_PSEXEC_TIMEOUT_MINUTES=60'
        'EHJIR_STALE_CLEANUP_DELAY_SECONDS=30'
        'EHJIR_REBOOT_DELAY_SECONDS=30'
        'EHJIR_DISABLE_NIGHT_PAUSE=1'
        'EHJIR_AD_DOMAIN=contoso.test'
        'EHJIR_GUI_DRY_RUN=1'
        'EHJIR_SKIP_PRE_RUN_ARCHIVE=1'
        'EHJIR_GUI_REMOVE_NON_INTUNE_MDM=1'
        'EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE=0'
        'EHJIR_FICTIONAL_SETTING=1'
    ), (New-Object System.Text.UTF8Encoding($true)))
    [System.IO.File]::WriteAllLines((Join-Path $w11Folder 'Windows11UpgradeToolkit.config'), [string[]]@(
        'W11UT_SETUP_SOURCE=\\srv-demo.contoso.test\SEMT-Demo$\Windows11'
        ('W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT={0}' -f (Join-Path $w11Folder 'SetupSourceGates'))
        'W11UT_AD_DOMAIN=contoso.test'
        'W11UT_THROTTLE=10'
        'W11UT_DIRECT_SETUP_UPGRADE=1'
        'W11UT_ALLOW_REBOOT=0'
        'W11UT_SKIP_INTUNE_INVENTORY_REFRESH=0'
        'W11UT_DELAY_BETWEEN_CYCLES_MINUTES=5'
        'W11UT_PSEXEC_TIMEOUT_MINUTES=480'
        'W11UT_ALLOW_ADVANCED_DISK_CLEANUP=0'
        'W11UT_ALLOW_DISM_COMPONENT_CLEANUP=1'
        'W11UT_SETUP_MEDIA_ID=Win11'
    ), (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllLines((Join-Path $w11Folder 'Lots\LOT-A\Windows11UpgradeToolkit.config'), [string[]]@('W11UT_SETUP_SOURCE=\\srv-lot.contoso.test\Share\Windows11'))
    [System.IO.File]::WriteAllLines((Join-Path $w11Folder 'Lots\LOT-A\Computers.txt'), [string[]]@('PC-DEMO-001'))
    [System.IO.File]::WriteAllLines((Join-Path $w11Folder 'Lots\LOT-TEMPLATE\Windows11UpgradeToolkit.config'), [string[]]@('W11UT_SETUP_SOURCE=\\template\Share'))
    $legacyHash = Get-TreeHash -Path $script:LegacyRoot
    $env:EHJIR_INTUNE_TENANT_ID = '00000000-0000-4000-8000-000000000001'

    # 1. Preview
    $previewReport = Join-Path $tempRoot 'preview-report.json'
    $preview = Invoke-Import -Arguments @()
    Test-Step 'preview exits 0' { $preview.ExitCode -eq 0 }
    Test-Step 'preview writes nothing in the SEMT Config folder' { @(Get-ChildItem -LiteralPath $configRoot -File | Where-Object { $_.Name -like '*.local.json*' -and $_.Name -notlike '*.template' }).Count -eq 0 }
    Test-Step 'preview announces the files -Apply would write' { $preview.Output -match 'Preview: nothing written' -and $preview.Output -match 'GuiState\.HybridJoin\.local\.json' }
    $previewWithReport = Invoke-Import -Arguments @('-ReportPath',$previewReport)
    Test-Step 'preview with -ReportPath writes only that report' { $previewWithReport.ExitCode -eq 0 -and (Test-Path -LiteralPath $previewReport) -and @(Get-ChildItem -LiteralPath $configRoot -File | Where-Object { $_.Name -like '*.local.json*' -and $_.Name -notlike '*.template' }).Count -eq 0 }
    $report = Read-Json -Path $previewReport
    Test-Step 'report mode is Preview' { $report.Mode -eq 'Preview' }

    # 3. Decisions
    Test-Step 'previous GUI default is not carried (EHJIR_PSEXEC_TIMEOUT_MINUTES=60)' { (Get-Row $report 'HybridJoin' 'EHJIR_PSEXEC_TIMEOUT_MINUTES').Decision -eq 'LegacyDefault' }
    Test-Step 'value below the SEMT minimum is Invalid (EHJIR_REBOOT_DELAY_SECONDS=30)' { (Get-Row $report 'HybridJoin' 'EHJIR_REBOOT_DELAY_SECONDS').Decision -eq 'Invalid' }
    Test-Step 'dry run is a launch choice, never carried' { (Get-Row $report 'HybridJoin' 'EHJIR_GUI_DRY_RUN').Decision -eq 'LaunchChoice' -and @($report.Warnings | Where-Object { $_ -like 'EHJIR_GUI_DRY_RUN=1*' }).Count -eq 1 }
    Test-Step 'EHJIR_SKIP_PRE_RUN_ARCHIVE has no equivalent' { (Get-Row $report 'HybridJoin' 'EHJIR_SKIP_PRE_RUN_ARCHIVE').Decision -eq 'NoEquivalent' }
    Test-Step 'unknown legacy key is reported' { (Get-Row $report 'HybridJoin' 'EHJIR_FICTIONAL_SETTING').Decision -eq 'Unknown' }
    Test-Step 'environment value wins and is read (process scope in test)' { (Get-Row $report 'HybridJoin' 'EHJIR_INTUNE_TENANT_ID').LegacySource -eq 'Environment:Process' }
    Test-Step 'old DISM alias supersedes W11UT_ALLOW_ADVANCED_DISK_CLEANUP' { (Get-Row $report 'Windows11Upgrade' 'W11UT_ALLOW_ADVANCED_DISK_CLEANUP').Decision -eq 'Superseded' -and (Get-Row $report 'Windows11Upgrade' 'W11UT_ALLOW_DISM_COMPONENT_CLEANUP').Decision -eq 'ToGlobal' }
    Test-Step 'inverted key: W11UT_SKIP_INTUNE_INVENTORY_REFRESH=0 gives PreCycleIntuneInventoryRefresh=true' { $r = Get-Row $report 'Windows11Upgrade' 'W11UT_SKIP_INTUNE_INVENTORY_REFRESH'; $r.SemtValue -eq 'true' -and $r.Decision -eq 'ToGlobal' }
    Test-Step 'LOT values are counted, not converted (LOT-TEMPLATE ignored)' { $s = @($report.Sources | Where-Object { $_.Campaign -eq 'Windows11Upgrade' })[0]; $s.LotCount -eq 1 -and $s.LotWithOwnValuesCount -eq 1 -and @($report.Rows | Where-Object { $_.LegacyValue -like '*srv-lot*' }).Count -eq 0 }
    Test-Step 'warning: gate root points into the previous toolkit folder' { @($report.Warnings | Where-Object { $_ -like 'Windows11Upgrade.SetupSubnetConcurrencyGateRoot points into the previous*' }).Count -eq 1 }
    Test-Step 'warning: gate root is not a UNC path' { @($report.Warnings | Where-Object { $_ -like 'Windows11Upgrade.SetupSubnetConcurrencyGateRoot is not a UNC path*' }).Count -eq 1 }
    Test-Step 'warning: DirectSetup with AllowReboot off no longer reboots' { @($report.Warnings | Where-Object { $_ -like 'DirectSetupUpgrade is on and AllowReboot is off*' }).Count -eq 1 }
    Test-Step 'common value equal in both toolkits goes to the global file (AdDomain)' { (Get-Row $report 'HybridJoin' 'EHJIR_AD_DOMAIN').Decision -eq 'ToGlobal' -and (Get-Row $report 'Windows11Upgrade' 'W11UT_AD_DOMAIN').Decision -eq 'ToGlobal' }
    Test-Step 'common value that differs per Campaign goes to the GUI state (ThrottleLimit, PsExecTimeoutMinutes)' { (Get-Row $report 'HybridJoin' 'EHJIR_THROTTLE').Target -eq 'GuiState.HybridJoin.local.json' -and (Get-Row $report 'Windows11Upgrade' 'W11UT_PSEXEC_TIMEOUT_MINUTES').Target -eq 'GuiState.Windows11Upgrade.local.json' }
    Test-Step 'common value equal to the other Campaign default goes global (DelayBetweenCyclesMinutes=5, NightPauseEnabled=false)' { (Get-Row $report 'Windows11Upgrade' 'W11UT_DELAY_BETWEEN_CYCLES_MINUTES').Decision -eq 'ToGlobal' -and (Get-Row $report 'HybridJoin' 'EHJIR_DISABLE_NIGHT_PAUSE').Decision -eq 'ToGlobal' }
    Test-Step 'previous toolkit files unchanged after preview' { (Get-TreeHash -Path $script:LegacyRoot) -eq $legacyHash }

    # 2. Apply
    $apply = Invoke-Import -Arguments @('-Apply')
    Test-Step 'apply exits 0' { $apply.ExitCode -eq 0 }
    $globalPath = Join-Path $configRoot 'SmartEndpointMigrationToolkit.local.json'
    $hjStatePath = Join-Path $configRoot 'GuiState.HybridJoin.local.json'
    $w11StatePath = Join-Path $configRoot 'GuiState.Windows11Upgrade.local.json'
    Test-Step 'apply writes the global file and both GUI states' { (Test-Path $globalPath) -and (Test-Path $hjStatePath) -and (Test-Path $w11StatePath) }
    Test-Step 'written files are UTF-8 with BOM' { $bytes = [System.IO.File]::ReadAllBytes($globalPath); $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF }
    $global = Read-Json -Path $globalPath
    Test-Step 'global file keeps every template key' { $template = Read-Json -Path (Join-Path $configRoot 'SmartEndpointMigrationToolkit.local.json.template'); @($template.Options.PSObject.Properties.Name | Where-Object { $global.Options.PSObject.Properties.Name -notcontains $_ }).Count -eq 0 }
    Test-Step 'global values: AdDomain, IntuneTenantId, setup source, advanced cleanup, MDM removal' {
        $global.Options.AdDomain -eq 'contoso.test' -and $global.Options.IntuneTenantId -eq '00000000-0000-4000-8000-000000000001' -and
        $global.Options.'Windows11Upgrade.SetupSourcePath' -eq '\\srv-demo.contoso.test\SEMT-Demo$\Windows11' -and
        $global.Options.'Windows11Upgrade.AllowAdvancedDiskCleanup' -eq $true -and $global.Options.'HybridJoin.AllowRemoveNonIntuneMdmEnrollment' -eq $true -and
        $global.Options.'Windows11Upgrade.DirectSetupUpgrade' -eq $true -and $global.Options.'Windows11Upgrade.AllowReboot' -eq $false
    }
    Test-Step 'values not carried stay inherited in the global file' { $global.Options.PsExecTimeoutMinutes -eq '__INHERIT__' -and $global.Options.ThrottleLimit -eq '__INHERIT__' -and $global.Options.'HybridJoin.RebootDelaySeconds' -eq '__INHERIT__' -and $global.Options.'Windows11Upgrade.SetupMediaId' -eq '__INHERIT__' }
    Test-Step 'GUI state values per Campaign' { (Read-Json $hjStatePath).Options.ThrottleLimit -eq 20 -and (Read-Json $hjStatePath).Campaign -eq 'HybridJoin' -and (Read-Json $w11StatePath).Options.PsExecTimeoutMinutes -eq 480 }
    Test-Step 'report written next to the configuration with mode Apply' { $files = @(Get-ChildItem -LiteralPath $configRoot -Filter 'LegacyConfigImport.*.local.json'); $files.Count -eq 1 -and (Read-Json $files[0].FullName).Mode -eq 'Apply' }

    Import-Module (Join-Path $script:SemtRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -Force
    $catalog = Import-SemtOptionCatalog -PackageRoot $script:SemtRoot
    $hjEffective = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign 'HybridJoin' -GlobalLayer (Read-Json $globalPath) -GuiLayer (Read-Json $hjStatePath)
    $w11Effective = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign 'Windows11Upgrade' -GlobalLayer (Read-Json $globalPath) -GuiLayer (Read-Json $w11StatePath)
    $w11Cmd = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign 'Windows11Upgrade' -GlobalLayer (Read-Json $globalPath)
    Test-Step 'SEMT resolution (GUI): HJ ThrottleLimit 20, night pause off, PsExec timeout SEMT default 120' { $hjEffective['ThrottleLimit'].Value -eq 20 -and $hjEffective['NightPauseEnabled'].Value -eq $false -and $hjEffective['PsExecTimeoutMinutes'].Value -eq 120 }
    Test-Step 'SEMT resolution (GUI): W11 ThrottleLimit 10, cycle delay 5, PsExec timeout 480' { $w11Effective['ThrottleLimit'].Value -eq 10 -and $w11Effective['DelayBetweenCyclesMinutes'].Value -eq 5 -and $w11Effective['PsExecTimeoutMinutes'].Value -eq 480 }
    Test-Step 'SEMT resolution (CMD, global only): W11 PsExec timeout is the SEMT default 360' { $w11Cmd['PsExecTimeoutMinutes'].Value -eq 360 -and $w11Cmd['PsExecTimeoutMinutes'].Source -eq 'Default' }
    Test-Step 'previous toolkit files unchanged after apply' { (Get-TreeHash -Path $script:LegacyRoot) -eq $legacyHash }

    # 4. Idempotence, existing values, backup, WhatIf, legacy defaults, missing folder
    $globalHashBefore = (Get-FileHash -LiteralPath $globalPath).Hash
    $again = Invoke-Import -Arguments @('-Apply')
    Test-Step 'second apply writes nothing and creates no backup' { $again.ExitCode -eq 0 -and $again.Output -match 'Nothing to write' -and (Get-FileHash -LiteralPath $globalPath).Hash -eq $globalHashBefore -and @(Get-ChildItem -LiteralPath $configRoot -Filter '*.bak').Count -eq 0 }

    $edited = ConvertTo-SemtHashtable -InputObject (Read-SemtJsonFile -Path $globalPath)
    $edited['Options']['AdDomain'] = 'other.test'
    Write-SemtJsonFileAtomic -Path $globalPath -InputObject $edited
    $editedHash = (Get-FileHash -LiteralPath $globalPath).Hash
    Start-Sleep -Seconds 1
    $kept = Invoke-Import -Arguments @('-Apply','-ReportPath',(Join-Path $tempRoot 'kept.json'))
    Test-Step 'existing explicit SEMT value is kept by default' { $kept.ExitCode -eq 0 -and (Read-Json $globalPath).Options.AdDomain -eq 'other.test' -and (Get-FileHash -LiteralPath $globalPath).Hash -eq $editedHash -and (Get-Row (Read-Json (Join-Path $tempRoot 'kept.json')) 'HybridJoin' 'EHJIR_AD_DOMAIN').Decision -eq 'KeptExisting' }

    Start-Sleep -Seconds 1
    $whatIf = Invoke-Import -Arguments @('-Apply','-ReplaceExistingValues','-WhatIf')
    Test-Step '-WhatIf with -Apply writes nothing' { $whatIf.ExitCode -eq 0 -and (Get-FileHash -LiteralPath $globalPath).Hash -eq $editedHash -and @(Get-ChildItem -LiteralPath $configRoot -Filter '*.bak').Count -eq 0 }

    Start-Sleep -Seconds 1
    $replace = Invoke-Import -Arguments @('-Apply','-ReplaceExistingValues')
    $backups = @(Get-ChildItem -LiteralPath $configRoot -Filter 'SmartEndpointMigrationToolkit.local.json.*.bak')
    Test-Step '-ReplaceExistingValues replaces after a backup of the previous file' { $replace.ExitCode -eq 0 -and (Read-Json $globalPath).Options.AdDomain -eq 'contoso.test' -and $backups.Count -eq 1 -and (Get-FileHash -LiteralPath $backups[0].FullName).Hash -eq $editedHash }

    Start-Sleep -Seconds 1
    $defaults = Invoke-Import -Arguments @('-IncludeLegacyDefaults','-ReportPath',(Join-Path $tempRoot 'defaults.json'))
    $defaultsReport = Read-Json (Join-Path $tempRoot 'defaults.json')
    Test-Step '-IncludeLegacyDefaults carries previous GUI defaults (HJ PsExec timeout 60 to the HJ GUI state)' { $defaults.ExitCode -eq 0 -and (Get-Row $defaultsReport 'HybridJoin' 'EHJIR_PSEXEC_TIMEOUT_MINUTES').Target -eq 'GuiState.HybridJoin.local.json' }

    $missing = Invoke-Import -Arguments @() -LegacyRoot (Join-Path $tempRoot 'Empty')
    Test-Step 'missing legacy folder fails' { $missing.ExitCode -ne 0 -and $missing.Output -match 'Legacy devices folder not found' }
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'Empty') | Out-Null
    $none = Invoke-Import -Arguments @() -LegacyRoot (Join-Path $tempRoot 'Empty')
    Test-Step 'folder without any previous toolkit fails' { $none.ExitCode -ne 0 -and $none.Output -match 'Neither IntuneHybridJoinToolkit nor Windows11UpgradeToolkit' }

    Test-Step 'previous toolkit files unchanged at the end' { (Get-TreeHash -Path $script:LegacyRoot) -eq $legacyHash }
}
finally {
    Remove-Item Env:\EHJIR_INTUNE_TENANT_ID -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCVCDkgNWzlkAAx
# /Li25yhUwIDrszrCqKWtdZgoXHWwdqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIW8btjFIXgHNsMbyQoZN6qMlJffF7odOMJcPDES8b85MA0GCSqG
# SIb3DQEBAQUABIIBgDEDUje+WyC1jrdd3HA2Bq9hWFsVzRXzT1bZ4E5hBrD1KGz1
# LioPAfZUZvJTMRi3raQzs+Hc98g/wZtWLzmsTTLbaFpxxpvspbT1VTSezpXdmXzz
# 7tdhEe8jiccHNl1rS8iZ68V2fDp6thXb0EHltOg6pOFoVF4WRlmr+A4E2WBzTnMs
# 3RZANthWW4Ym3tK4mFchRDPEoMz2bhXhVpkN63K0xCWpGQdLgT60EmYOY/DJ/0bs
# kIGrOx9t5Wdb5H8jqSh1ROtmiWuhRp1/YYIIagGtd+x//fiS8LLHOSAwvgNjigNO
# 6uKcpTaVsRWuZiwLkwTqbG5pN6U7jGhmCQ+nQ4IMTCRV3fDwIeiPJWcbbM4Z9ZjP
# vWrrKkY4rJNFpP30hvWNHQX35gS30hzC5+kZuusRS6bQltn+KJnHS9IpJSVZWfNC
# GW5/cN+wecxRsGmG5KsjoHO8KfGZoUpTnmFCtXgIO3Jb53APC1xYEjLz6HTQ4vbA
# l84G04g3+sbYEIKy0KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxNDM5
# MDhaMC8GCSqGSIb3DQEJBDEiBCDz/IeJFr1oN59ZexNbLXS9K5yIxqRnV8K7sGC3
# 2yBKgTANBgkqhkiG9w0BAQEFAASCAgClMbdWj66lBO2rbAsz7mEs+PZnfItPFyJp
# f6YfmgUvjDcvRKZK+ov2H3+FM9ejoRe3tTFdVjDn2ZJ4JQQuCTymM7PEpY7/YCh7
# TEuBzAjxCyJ7Osxy+ADCZztwSAJNEzkDcOn1AKyeaZreGqKK1U3ESdS4dCrXWA67
# b/qKc3djZFoEMsbBs50i3F+6mjF/nxSInqfgM616NwDU4QZQQwN3etUcctEBhmJp
# fkO0tEHHuEHAFjR7FrpTXYFJFXBpiBuWk4nyNsBRZbP1xDuQ4mUZvqIQQ8kku7Ps
# RTPH+1lqbHPbwXIENxXp7GHjZoUrJ4qNhF1vJ3t57iqTU8ZGLkqCE2Hr8xJIzoZS
# YQU9gktLby4i9A4eC5BfBjroXJGCLnsdP1w7nbz0EfloES/36/ma/hixl42ZjyA+
# dsanehLVmM3kdnmor6QAqbFvnqSHn620BwW/2CTy2Y8dH19gVHYU9zKJOvmFYdzu
# dSDuAgvnM+fKWBGsKcTy/bl+6zVJskA+Nax5JwGYoupgUeSGoWkSDshpKZkg8VeL
# rtrWCgXq3MOZKZTb/Cz8HZrvDAtrtqcmT728rg+mqCgZgxQaHfK6EblNaI8sK1Ga
# MD8k66JkOn5px80X4sYnB316zj72XmOCbL+qWA2o015iL4DAbSsQD4Up98deiu6R
# /mwRurqfFw==
# SIG # End signature block
