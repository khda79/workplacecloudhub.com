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
        'Waves\HybridJoin\WAVE-PRIVATE\Wave.local.json',
        'Runs\HybridJoin\WAVE-PRIVATE\R1\Logs\SEMT-launcher.log',
        'Config\SmartEndpointMigrationToolkit.local.json',
        'Config\GuiState.HybridJoin.local.json',
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
        $optionsFile = @(Get-ChildItem -LiteralPath (Join-Path $pkg 'Runs\HybridJoin\_SinglePC') -Recurse -Filter 'EffectiveOptions.json' | Sort-Object LastWriteTime | Select-Object -First 1)[0]
        $effective = Get-Content -LiteralPath $optionsFile.FullName -Raw | ConvertFrom-Json
        $rows = @{}; foreach ($row in @($effective.Options)) { $rows[[string]$row.Option] = $row }
        $effective.Campaign -eq 'HybridJoin' -and [string]$rows['ThrottleLimit'].Value -eq '1' -and [string]$rows['ThrottleLimit'].Source -eq 'Rule:SinglePcOneWorker' -and [string]$rows['MaxCycles'].Source -eq 'Rule:RunModeOnce' -and [string]$rows['HybridJoin.AllowRebootAfterDsregLeave'].Value -eq 'False'
    }
    Test-Step 'extracted package: Windows 11 dry run records DryRun from the CLI layer' {
        $optionsFile = @(Get-ChildItem -LiteralPath (Join-Path $pkg 'Runs\Windows11Upgrade\_SinglePC') -Recurse -Filter 'EffectiveOptions.json' | Select-Object -First 1)[0]
        $rows = @{}; foreach ($row in @((Get-Content -LiteralPath $optionsFile.FullName -Raw | ConvertFrom-Json).Options)) { $rows[[string]$row.Option] = $row }
        [string]$rows['DryRun'].Value -eq 'True' -and [string]$rows['DryRun'].Source -eq 'Cli'
    }
    Test-Step 'Single PC Run is created under Runs\<Campaign>\_SinglePC of the extracted package' { @(Get-ChildItem -LiteralPath (Join-Path $pkg 'Runs\HybridJoin\_SinglePC') -Directory -ErrorAction SilentlyContinue).Count -ge 1 }
    $gui = Invoke-Child -Arguments @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',(Join-Path $pkg 'Scripts\SmartM365-EndpointMigration-GUI.ps1'),'-ValidateOnly')
    Test-Step 'extracted package: GUI -ValidateOnly succeeds' { $script:ChildExitCode -eq 0 -and @($gui | Where-Object { $_ -match '^ValidateOnly: GUI 1\.0\.0 loaded' }).Count -eq 2 }

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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAtQ2Gt2l+VzUvO
# 1z8h/ahU/RTBJgHHYgMSzn/XdA0fbKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAKKiQoFchjf6sL+bDQOfcy
# GoaJNdEdOLDiz0e5JuNX0zANBgkqhkiG9w0BAQEFAASCAYAw62KAFZkGE7OQ1Jkm
# OTLTSFOdNbIaPG42FtilyeP0K6sPPJV19iZotzLhhbJRnIjFFF6PuyuugGN6h4Lm
# dMYkyjvCjaJbkDi2DqlHLoRHUOKcKHosUguK46ZqYGZ3K4UqVWp0S+EvlHmi6TDE
# 0XgqsRsDpYTmecj0MBtt+JsY+xUpy5Ws5dPLazFQnkMumgN+OpI3YYSqrRuV1scl
# +z0bpTs1kqJWLmkPK8QuabntaWpe5m5xdko/Z3dZWmj7b0kjGQXVyZp0UqbfBHXt
# 6nL8lPz55Mp7ZuvPJq2lWqOPXfSpn/HV9YxBhshzPKf2R/44iOfNBvhmsnF9z6Cc
# 8LrzdU0sdiJY8TuOwVkKPdKcZGbibDskqywq8QsP6b7R2jyuydaXuXSavpioQ/Ig
# xbjJ2bIxIvylKtmbvXQHgfOp29ePEXIi3m2EhTlRY6LL1LTsmU6o1/ZX33WbeppP
# cwEb1+2olaBJIrzFYLXhgAnYVmLvYP/5PcVfZQV8nprcDUg=
# SIG # End signature block
