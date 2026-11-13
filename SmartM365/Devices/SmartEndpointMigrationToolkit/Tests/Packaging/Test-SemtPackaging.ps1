<#
.SYNOPSIS
    Offline packaging tests for the Smart Endpoint Migration Toolkit release.
.DESCRIPTION
    Copies the package to a temp folder, plants private and operational files (Wave, Run log,
    local configuration, inventory, setup media, memory), builds an unsigned qualification
    release twice, and checks: exclusions, deterministic ZIP, checksum file, distribution
    manifest, independence of the extracted package (engine -ValidateOnly and GUI
    -ValidateOnly run from the extracted folder only, no reference to the previous toolkit
    folders), and that a modified component is detected by the manifest. Nothing is
    published, tagged or deployed.
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
$sourceRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try { Write-TestResult -Condition ([bool](& $Check)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

function Invoke-Child {
    # Child output (including its error stream) is captured as text; the exit code is kept in $script:ChildExitCode.
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $ErrorActionPreference = 'Continue'
    $lines = @(& $powershell @Arguments 2>&1 | ForEach-Object { [string]$_ })
    $script:ChildExitCode = $LASTEXITCODE
    return $lines
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-PkgTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$copyRoot = Join-Path $tempRoot 'Source\SmartEndpointMigrationToolkit'
$previousOverride = $env:SEMT_RUNTIME_ROOT_OVERRIDE
$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
try {
    New-Item -ItemType Directory -Path $copyRoot -Force | Out-Null
    foreach ($item in @(Get-ChildItem -LiteralPath $sourceRoot -Force | Where-Object { $_.Name -notin @('Runs','Inventory','DemoAssets') })) { Copy-Item -LiteralPath $item.FullName -Destination $copyRoot -Recurse -Force }
    Get-ChildItem -LiteralPath (Join-Path $copyRoot 'Release') -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -in @('Output','Work') } | Remove-Item -Recurse -Force
    $planted = @(
        'Waves\HybridJoin\WAVE-PRIVATE\Computers.txt',
        'Waves\HybridJoin\WAVE-PRIVATE\Wave.local.json.txt',
        'Waves\HybridJoin\Computers.AlreadyEnrolled.txt',
        'Waves\Windows11Upgrade\Computers.HardwareNotCapable.txt',
        'Waves\Windows11Upgrade\Campaign.Notes.txt',
        'Waves\Windows11Upgrade\Backup\WAVE-OLD\Computers.txt',
        'Runs\HybridJoin\WAVE-PRIVATE\R1\Logs\SEMT-launcher.log',
        'Config\SmartEndpointMigrationToolkit.local.json.txt',
        'Config\GuiState.HybridJoin.local.json.txt',
        'Inventory\DevicesIntune.csv',
        'Campaigns\Windows11Upgrade\SetupSource\Win11-fr-FR\setup.exe',
        'Campaigns\Windows11Upgrade\IntuneWin32\Output\package.intunewin',
        'DemoAssets\capture.png',
        'LOCAL_MEMORY.md'
    )
    foreach ($relative in $planted) {
        $path = Join-Path $copyRoot $relative
        New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
        Set-Content -LiteralPath $path -Value 'PRIVATE-CONTENT-MUST-NOT-SHIP' -Encoding ASCII
    }
    Set-Content -LiteralPath (Join-Path $copyRoot 'Scripts\UnsignedTestFixture.ps1') -Value "<#`n.VERSION`n1.0.0`n#>" -Encoding ASCII
    $builder = Join-Path $copyRoot 'Release\SmartM365-EndpointMigration-Build-Release.ps1'

    $preview = Invoke-Child -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$builder)
    Test-Step 'preview without -AllowUnsigned refuses unsigned sources' { $script:ChildExitCode -ne 0 -and ($preview -join "`n") -match 'not signed' }

    $out1 = Join-Path $tempRoot 'Out1'
    $out2 = Join-Path $tempRoot 'Out2'
    $build1 = Invoke-Child -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$builder,'-Build','-AllowUnsigned','-OutputRoot',$out1)
    Test-Step 'qualification build succeeds' { $script:ChildExitCode -eq 0 -and (Test-Path (Join-Path $out1 'SmartEndpointMigrationToolkit-1.0.0.zip')) }
    $null = Invoke-Child -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$builder,'-Build','-AllowUnsigned','-OutputRoot',$out2)
    $zip1 = Join-Path $out1 'SmartEndpointMigrationToolkit-1.0.0.zip'
    $zip2 = Join-Path $out2 'SmartEndpointMigrationToolkit-1.0.0.zip'
    Test-Step 'ZIP is deterministic (two builds, same SHA256)' { (Get-FileHash $zip1 -Algorithm SHA256).Hash -eq (Get-FileHash $zip2 -Algorithm SHA256).Hash }
    $sha = Get-Content -LiteralPath (Join-Path $out1 'SmartEndpointMigrationToolkit-1.0.0.sha256')
    Test-Step 'checksum file lists the ZIP with its real SHA256' { ($sha[0] -split ' ')[0] -eq (Get-FileHash $zip1 -Algorithm SHA256).Hash }
    $notes = [IO.File]::ReadAllText((Join-Path $out1 'RELEASE-NOTES-1.0.0.md'))
    Test-Step 'release notes mark the build UNSIGNED and name the tag without creating it' { $notes -match 'UNSIGNED QUALIFICATION BUILD' -and $notes -match 'endpoint-migration-toolkit-v1\.0\.0' -and $notes -match 'not created by this script' }
    Test-Step 'no build work folder is left behind' { @(Get-ChildItem -LiteralPath $out1 -Directory).Count -eq 0 }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    # Fresh extraction folder with spaces and non-ASCII characters, outside the repository.
    $extractRoot = Join-Path $tempRoot ('Recette package ' + [char]0x00E9 + [char]0x00E8 + ' ' + [char]0x4E2D + [char]0x6587)
    [IO.Compression.ZipFile]::ExtractToDirectory($zip1, $extractRoot)
    Test-Step 'extraction folder has spaces and non-ASCII characters and is outside the repository' { $extractRoot -match ' ' -and $extractRoot -match '[^\x00-\x7F]' -and -not $extractRoot.StartsWith($sourceRoot, [StringComparison]::OrdinalIgnoreCase) }
    $pkg = Join-Path $extractRoot 'SmartEndpointMigrationToolkit'
    $shipped = @(Get-ChildItem -LiteralPath $pkg -Recurse -File | ForEach-Object { $_.FullName.Substring($pkg.Length).TrimStart('\') })
    foreach ($relative in $planted) { Write-TestResult -Condition ($shipped -notcontains $relative) -Message ("private/operational file excluded: {0}" -f $relative) }
    Test-Step 'no file of the package contains the planted private marker' { @(Get-ChildItem -LiteralPath $pkg -Recurse -File | Where-Object { $_.Name -ne 'Test-SemtPackaging.ps1' -and [IO.File]::ReadAllText($_.FullName).Contains('PRIVATE-CONTENT-MUST-NOT-SHIP') }).Count -eq 0 }
    foreach ($required in @('README.md','Start-SmartEndpointMigrationToolkit-GUI.cmd','Start-SmartEndpointMigrationToolkit-Demo.cmd','Export-UserLicensesCsv.cmd','SmartEndpointMigrationToolkit.version.json.txt','SmartEndpointMigrationToolkit.components.json.txt','Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1','Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1','Scripts\SmartM365-EndpointMigration-GUI.ps1','Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1','Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1','Docs\PROVENANCE.json.txt','Assets\WorkplaceCloudHub-lockup-WPF.png')) {
        Write-TestResult -Condition ($shipped -contains $required) -Message ("shipped: {0}" -f $required)
    }
    $manifest = Get-Content -LiteralPath (Join-Path $pkg 'SmartEndpointMigrationToolkit.components.json.txt') -Raw | ConvertFrom-Json
    Test-Step 'distribution manifest lists every shipped file with its SHA256' {
        $listed = @($manifest.Files | ForEach-Object { $_.Path.Replace('/','\') })
        $missing = @($shipped | Where-Object { $_ -ne 'SmartEndpointMigrationToolkit.components.json.txt' -and $listed -notcontains $_ })
        $bad = @($manifest.Files | Where-Object { (Get-FileHash -LiteralPath (Join-Path $pkg $_.Path.Replace('/','\')) -Algorithm SHA256).Hash -ne $_.Sha256 })
        $missing.Count -eq 0 -and $bad.Count -eq 0
    }
    Test-Step 'unsigned qualification manifest declares RequireSignature=false' { $manifest.RequireSignature -eq $false -and $manifest.PackageVersion -eq '1.0.0' -and $manifest.WorkerContractVersion -eq 1 }

    $legacyReference = '(?i)(Devices[\\/](IntuneHybridJoinToolkit|Windows11UpgradeToolkit)|\.\.[\\/]+\.\.[\\/]+(IntuneHybridJoinToolkit|Windows11UpgradeToolkit)|\.\.[\\/]+(IntuneHybridJoinToolkit|Windows11UpgradeToolkit))'
    $offending = New-Object System.Collections.Generic.List[string]
    foreach ($file in @(Get-ChildItem -LiteralPath $pkg -Recurse -File | Where-Object { $_.Extension -in @('.ps1','.psm1','.psd1','.cmd') -and $_.Name -notlike 'SmartM365.Gui*.ps1' })) {
        # SmartM365.Gui* are verbatim shared libraries; their sibling-product logo fallback is never reached because the GUI passes its own logo and icon.
        $lineNumber = 0
        foreach ($line in [IO.File]::ReadAllLines($file.FullName)) {
            $lineNumber++
            if ($line -match $legacyReference -and $line -notmatch '(?i)(provenance|ported from|copied from|ported|source:|legacy)') { $offending.Add(('{0}:{1}' -f $file.FullName.Substring($pkg.Length),$lineNumber)) }
        }
    }
    Test-Step 'extracted package has no code reference to the previous toolkit folders' { $offending.Count -eq 0 }
    if ($offending.Count -gt 0) { $offending | Select-Object -First 10 | ForEach-Object { Write-Host ("INFO: legacy folder reference {0}" -f $_) } }
    Test-Step 'module imports inside the package use -MinimumVersion' {
        @(Get-ChildItem -LiteralPath $pkg -Recurse -File | Where-Object { $_.Extension -in @('.ps1','.psm1') } | Where-Object { [IO.File]::ReadAllText($_.FullName) -match 'Import-Module[^\r\n]*SmartEndpointMigration\.(Core|Campaign)[^\r\n]*' -and [IO.File]::ReadAllText($_.FullName) -notmatch 'Import-Module[^\r\n]*SmartEndpointMigration\.[A-Za-z.]+\.psd1[^\r\n]*-MinimumVersion' }).Count -eq 0
    }

    $env:SEMT_RUNTIME_ROOT_OVERRIDE = Join-Path $tempRoot 'Runtime'
    $engine = Join-Path $pkg 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'
    $validateHj = Invoke-Child -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$engine,'-Campaign','HybridJoin','-ComputerName','SEMT-QUALIF-PC','-RunMode','Once','-ValidateOnly')
    $validateHjText = $validateHj -join "`n"
    Test-Step 'extracted package: Hybrid Join engine -ValidateOnly succeeds' { $script:ChildExitCode -eq 0 -and $validateHjText -match 'ValidateOnly: options, rules and component integrity are valid' }
    Test-Step 'extracted package: integrity verified against the shipped manifest' { $validateHjText -match 'Component integrity: Files=\d+; Verified=\d+; Mismatch=0; Manifest=True' }
    $validateW11 = Invoke-Child -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$engine,'-Campaign','Windows11Upgrade','-ComputerName','SEMT-QUALIF-PC','-RunMode','Once','-DryRun','-ValidateOnly')
    Test-Step 'extracted package: Windows 11 engine -ValidateOnly (dry run) succeeds' { $script:ChildExitCode -eq 0 -and ($validateW11 -join "`n") -match 'ValidateOnly: options, rules and component integrity are valid' }
    Test-Step 'extracted package: effective options of the Single PC Run record values and sources (rules applied)' {
        $optionsFile = @(Get-ChildItem -LiteralPath (Join-Path $pkg 'Runs\HybridJoin\_SinglePC') -Recurse -Filter 'EffectiveOptions.json.txt' | Sort-Object LastWriteTime | Select-Object -First 1)[0]
        $effective = Get-Content -LiteralPath $optionsFile.FullName -Raw | ConvertFrom-Json
        $rows = @{}; foreach ($row in @($effective.Options)) { $rows[[string]$row.Option] = $row }
        $effective.Campaign -eq 'HybridJoin' -and [string]$rows['ThrottleLimit'].Value -eq '1' -and [string]$rows['ThrottleLimit'].Source -eq 'Rule:SinglePcOneWorker' -and [string]$rows['MaxCycles'].Source -eq 'Rule:RunModeOnce' -and [string]$rows['HybridJoin.AllowRebootAfterDsregLeave'].Value -eq 'False'
    }
    Test-Step 'extracted package: Windows 11 dry run records DryRun from the CLI layer' {
        $optionsFile = @(Get-ChildItem -LiteralPath (Join-Path $pkg 'Runs\Windows11Upgrade\_SinglePC') -Recurse -Filter 'EffectiveOptions.json.txt' | Select-Object -First 1)[0]
        $rows = @{}; foreach ($row in @((Get-Content -LiteralPath $optionsFile.FullName -Raw | ConvertFrom-Json).Options)) { $rows[[string]$row.Option] = $row }
        [string]$rows['DryRun'].Value -eq 'True' -and [string]$rows['DryRun'].Source -eq 'Cli'
    }
    $refused = Invoke-Child -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$engine,'-Campaign','Prerequisites','-ComputerName','SEMT-QUALIF-PC','-RunMode','Once','-ValidateOnly')
    Test-Step 'extracted package: the Prerequisites Campaign is refused by default (SEMT_CAMPAIGN_DISABLED) and no Run is created' { $script:ChildExitCode -ne 0 -and ($refused -join "`n") -match 'SEMT_CAMPAIGN_DISABLED' -and -not (Test-Path -LiteralPath (Join-Path $pkg 'Runs\Prerequisites')) }
    $env:SEMT_PREREQUISITES_ENABLED = '1'
    try { $validatePrereq = Invoke-Child -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$engine,'-Campaign','Prerequisites','-ComputerName','SEMT-QUALIF-PC','-RunMode','Once','-ValidateOnly') }
    finally { Remove-Item Env:\SEMT_PREREQUISITES_ENABLED -ErrorAction SilentlyContinue }
    Test-Step 'extracted package: Prerequisites engine -ValidateOnly succeeds once enabled (endpoint verified by the manifest)' { $script:ChildExitCode -eq 0 -and ($validatePrereq -join "`n") -match 'ValidateOnly: options, rules and component integrity are valid' -and ($validatePrereq -join "`n") -match 'SmartM365-Invoke-EndpointMigration-Prerequisites\.ps1' }
    Test-Step 'Single PC Run is created under Runs\<Campaign>\_SinglePC of the extracted package' { @(Get-ChildItem -LiteralPath (Join-Path $pkg 'Runs\HybridJoin\_SinglePC') -Directory -ErrorAction SilentlyContinue).Count -ge 1 }
    $gui = Invoke-Child -Arguments @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',(Join-Path $pkg 'Scripts\SmartM365-EndpointMigration-GUI.ps1'),'-ValidateOnly')
    # The GUI version comes from the package version manifest: no hard-coded version to forget.
    $guiVersion = [string](Get-Content -LiteralPath (Join-Path $pkg 'SmartEndpointMigrationToolkit.version.json.txt') -Raw | ConvertFrom-Json).Components.GUI
    Test-Step "extracted package: GUI -ValidateOnly succeeds (GUI $guiVersion)" { $script:ChildExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($guiVersion) -and @($gui | Where-Object { $_ -match ('^ValidateOnly: GUI ' + [regex]::Escape($guiVersion) + ' loaded') }).Count -eq 2 }

    $workerPath = Join-Path $pkg 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'
    Add-Content -LiteralPath $workerPath -Value '# tampered' -Encoding ASCII
    $tampered = Invoke-Child -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$engine,'-Campaign','HybridJoin','-ComputerName','SEMT-QUALIF-PC','-RunMode','Once','-ValidateOnly')
    Test-Step 'modified component is blocked by the shipped manifest (hash not supplied by the launcher)' { $script:ChildExitCode -ne 0 -and ($tampered -join "`n") -match 'SEMT_COMPONENT_INTEGRITY_BLOCKED' -and ($tampered -join "`n") -match 'PsExecWorker' }
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    $env:SEMT_RUNTIME_ROOT_OVERRIDE = $previousOverride
    try { if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force } } catch { Write-Host ("WARN: temp folder not fully removed: {0}" -f $tempRoot) }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA4NmRiNxG+mmuV
# Ke4G7JQp+fzQpZ1vugMLmbBFBjHrKaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAh7PAHUj/iw/bseXyAX5sR4OQUw5FYkQbKhTCwWI6T+MA0GCSqG
# SIb3DQEBAQUABIIBgIqelU2yuDohMVqVFTbO3kxxp126QZlP2UbkMtXy0YPMTPCJ
# gQ3ulWQ9fDutWCFE33l7r5VVJe/NvhU9kcf55RVPDVPOxS6balTK12wBMxglcvRq
# GPkBhXA2RLwYJzRXSl8nT2ntx81ghfSjuJrICM4LJdkzwLjSw0FjJLpnLCVYFe+W
# DQKccRhNx94voWgjLVCrWLvT8dyO/MTmBAFvD08HJg+aAPCBL7keLPk03bs1XPc9
# GWB4+IGSp87Yd2OF7rsBuvHZe6DnMOVbmOUW9jx2VF9tO+VjU7k7N7ExS67oOHZG
# Fk5yX/J/swI2rJgPKVacnpLQsoGOILVFuDwUFUo37/7LP/eKnmMwqhS+hh4VRoV8
# waKfQuB5yxHFw0qPYUeyfN0zdWSiVVGttAbXgVNozgfC537rRWrDfTNeBWSpTU1D
# 2dN9v5buDz8DXFcBFIPgGTE8QyIogwzb2j/Zp+4Ypd4sC7jLw4d//gkmdx6Zcibi
# NQGDyhpCCK173uQJHqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxNDU2
# NTdaMC8GCSqGSIb3DQEJBDEiBCAVAT3yOpDSAk/Jijhbc7GqRvw0dy7QoAl59gW1
# 0wQOqzANBgkqhkiG9w0BAQEFAASCAgBxsYR6Ds8LDQ2pR984ralYvuanvY/+EgTy
# VVBMOed5xzNARbMnFwf3zyuWnk7A7bKYI+tijT4nFpulm6gkoJQ6kSMbj1+v8rJ/
# 0ZolNnvesJzf2No06472U8Y1DrrEMwapTD8s9BKN9hF7/IaQ64LNwvBHs02oV7lv
# yYbAvmnBls7hOPbovl1EImvPhrhsrj0Uq7VnILhuGXUWXy0JcmaXo68w+dFHjd/K
# fIgQZ7xveRy5cWwmXSwZCWBk/RjA7VMG8yKVyOhE4zGFBU29922E90H6DmtMY/B+
# 0EjUsGeIFcuZ5xDq9O3GB/VoE7lQHQux7Ov5RW6kzPNMmzUfgP69oNkXm0Kx5n/d
# nDhoG+QU2l89VlbZ7mF+wwC6i/m2t3NjpHj3aT4+yzV+cfNDohQTjdA70ciPTDy+
# D+Y4t/T8gUPB+Wmw+tVAfuR5vMQv92RJw30Mq4PyqTnf7bH60DsnOERlOvKifbl6
# IIF98Nug/W6mM8ybG9Rl2WTwR2hJhszE9HUWqqpdtFkEKybK3nkBvbDgXtwhd+t8
# UY2ElN6FhSfq6p4LPqCakN0N/kcokjwU+iJ6bZuyzS/6e39Yo+VGf4J0RzthinA8
# yKMoLTYf6sC7GCdo0ykppYGuqrAupGafDrUSpLr9jWM6xchXOVFzti20PR1mIIuZ
# J6B2n9gyFw==
# SIG # End signature block
