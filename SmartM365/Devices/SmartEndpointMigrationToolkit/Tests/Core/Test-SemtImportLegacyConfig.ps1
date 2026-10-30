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

    Import-Module (Join-Path $script:SemtRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAxbwopaUmZdXxD
# 6ld9iEP+7lh8lqJZ7m69E7wy76vmFKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB6l29tnn2c+oGsP8Z4xNjR
# Al1MGzYpKTCPDxKDUWSUJTANBgkqhkiG9w0BAQEFAASCAYAFPDXo98kaSDiSLzTo
# tc8g1fMhOgoiMf82a0W504AVnhEPP8ZDDgcWTuzO0iVHFqs48lx/Qjd/7Vi0SBl2
# giJE71xzzSUA+D7jvMPvaVRt4qXZGQ2YLnXu+y0lurcDu6ffVW3WGoeKYLE9qnxm
# nz8J3hUatjRroB1zilkgIk8E+dy8GKPSRRdySD+UuG5XrJbCnOX+6icudkhs4cmb
# 08rEwK6sZ7nLXIIEQ+JyQooZhXhsjjfi92G6lwYgHJ3RrkGMHdCpMqBV5lkNTePY
# LSgXsnyJ6nOjNuyywAJx3MM/wpDnxzOUeDw7/RHbLnVv7KBwnHIjfEPgLgDcTago
# Sx/vj7lBkJl+Pj44pFHIHs9Xlb73yVRcYXCDVOf+ZMCuxrfp4nrrITIh10mIy074
# peNgLcgsYS1QDpmjZzAHUPybCYQ5HEbvY/qGzwMSirMKbFFFA+Oc64lkltjZBsa+
# AwauNIkBrlXhOxak7BjbAg8csfv/nF1jcWIL0frpIIv1VfM=
# SIG # End signature block
