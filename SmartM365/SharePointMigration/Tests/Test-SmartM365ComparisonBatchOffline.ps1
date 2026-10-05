<#
.SYNOPSIS
    Verify comparison batch discovery, failure handling, logging and CMD forwarding offline.
.VERSION
    1.0.0
#>
#requires -Version 7.4
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
function Invoke-TestCmd {
    param([string]$Command)
    $info = [Diagnostics.ProcessStartInfo]::new()
    $info.FileName = $env:ComSpec
    $info.Arguments = '/d /s /c ' + $Command
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    # This child uses unsigned synthetic launchers only; leave host policy unchanged.
    $info.Environment['PSExecutionPolicyPreference'] = 'Bypass'
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($info)
    try {
        $output = $process.StandardOutput.ReadToEnd() + $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $script:CmdExitCode = $process.ExitCode
        return $output
    }
    finally { $process.Dispose() }
}
$project = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root = [IO.Path]::GetFullPath((Join-Path $env:TEMP ('SmartM365 comparison test ' + [guid]::NewGuid().ToString('N'))))
if (-not $root.StartsWith([IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test root.' }
try {
    $generic = Join-Path $root 'Scripts\Launchers\Generic'
    [void](New-Item -ItemType Directory -Path $generic -Force)
    $batch = Join-Path $generic 'SmartM365-SharePointMigration-ComparisonBatch.ps1'
    Copy-Item -LiteralPath (Join-Path $project 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-ComparisonBatch.ps1') -Destination $batch
    $cmd = Join-Path $root 'Start-SmartM365-SharePointMigration-ComparisonBatch.cmd'
    Copy-Item -LiteralPath (Join-Path $project 'Start-SmartM365-SharePointMigration-ComparisonBatch.cmd') -Destination $cmd
    foreach ($name in @('SiteA','Site B','SiteC','BrokenConfig','NewMigration')) {
        $dir = Join-Path $root "Migrations\$name"
        [void](New-Item -ItemType Directory -Path $dir -Force)
        $configName = if ($name -eq 'BrokenConfig') { 'DifferentName' } else { $name }
        "@{ Name='$configName' }" | Set-Content -LiteralPath (Join-Path $dir 'migration.config.psd1')
    }
    $stub = Join-Path $generic 'SmartM365-SharePointMigration-Launcher.ps1'
    [IO.File]::WriteAllText($stub, @'
param([string]$MigrationName,[string]$Action,[switch]$NonInteractive,[switch]$UseCertificate,[switch]$Force)
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$started = [datetime]::UtcNow
Start-Sleep -Milliseconds 100
[pscustomobject]@{ Name=$MigrationName; Action=$Action; NonInteractive=[bool]$NonInteractive; Certificate=[bool]$UseCertificate; Force=[bool]$Force; Start=$started.ToString('o'); End=[datetime]::UtcNow.ToString('o') } |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root "$MigrationName-$Action.event.json")
Write-Output 'First output line'
Write-Output 'Second output line'
if ($MigrationName -eq 'Site B' -and $Action -eq 'CompareFiles') { [Console]::Error.WriteLine('Synthetic missing source inventory'); exit 9 }
'@)
    $pwsh = (Get-Command pwsh).Source
    $argsList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',$batch,'-ProjectRoot',$root)
    $output = & $pwsh @argsList -PlanOnly -MigrationNames 'SiteA,SiteC' 2>&1
    if ($LASTEXITCODE -ne 0 -or (Test-Path -LiteralPath (Join-Path $root 'Migrations\logs')) -or
        @(Get-ChildItem -LiteralPath $root -Filter '*.event.json').Count -or -not ($output -match '4 comparisons')) { throw 'PlanOnly created output or planned the wrong jobs.' }
    $output = & $pwsh @argsList -PlanOnly -MigrationNames BrokenConfig 2>&1
    if ($LASTEXITCODE -ne 1 -or (Test-Path -LiteralPath (Join-Path $root 'Migrations\logs'))) { throw 'Invalid plan was accepted or created logs.' }
    $output = & $pwsh @argsList 2>&1
    if ($LASTEXITCODE -ne 1) { throw "Batch did not propagate a comparison failure: $output" }
    $events = @(Get-ChildItem -LiteralPath $root -Filter '*.event.json' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json } | Sort-Object { [datetime]$_.Start })
    if ($events.Count -ne 6) { throw 'Discovery skipped valid migrations or ran a template/invalid configuration.' }
    for ($i=1; $i -lt $events.Count; $i++) {
        if ([datetime]$events[$i].Start -lt [datetime]$events[$i-1].End) { throw 'Comparisons overlapped.' }
    }
    if (@($events[0..2] | Where-Object Action -NE CompareFiles).Count -or
        @($events[3..5] | Where-Object Action -NE ComparePermissions).Count) { throw 'File/permission phase order changed.' }
    if (@($events | Where-Object { -not $_.NonInteractive -or $_.Certificate -or $_.Force }).Count) { throw 'Incorrect default launcher flags.' }
    $summaries = @(Get-ChildItem -LiteralPath (Join-Path $root 'Migrations\logs') -Recurse -Filter summary.csv)
    $summary = @(Import-Csv -LiteralPath $summaries[0].FullName)
    if ($summary.Count -ne 8 -or @($summary | Where-Object Status -EQ FAILED).Count -ne 3 -or
        @($summary | Where-Object ExitCode -EQ 9).Count -ne 1 -or
        @($summary | Where-Object { $_.Migration -eq 'BrokenConfig' -and $_.Error -like '*name differs*' }).Count -ne 2) { throw 'Incorrect summary or lost child exit code/configuration error.' }
    foreach ($row in $summary) {
        if (-not (Test-Path -LiteralPath $row.ConsoleLog)) { throw 'Per-comparison console log missing.' }
        foreach ($line in (Get-Content -LiteralPath $row.ConsoleLog)) {
            if ($line -notmatch '^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] ') { throw 'Console log line has no timestamp.' }
        }
    }
    $failureLog = ($summary | Where-Object ExitCode -EQ 9).ConsoleLog
    if (-not (Select-String -LiteralPath $failureLog -SimpleMatch 'Synthetic missing source inventory' -Quiet)) { throw 'Native stderr was not captured.' }
    $output = & $pwsh @argsList -ComparisonMode PermissionsOnly -MigrationNames 'SiteC,sitec' -AuthMode Certificate -Force 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Selected comparison failed: $output" }
    $event = Get-Content -LiteralPath (Join-Path $root 'SiteC-ComparePermissions.event.json') -Raw | ConvertFrom-Json
    if (-not $event.Certificate -or -not $event.Force -or -not $event.NonInteractive) { throw 'Certificate or explicit Force flag was lost.' }
    $summaries = @(Get-ChildItem -LiteralPath (Join-Path $root 'Migrations\logs') -Recurse -Filter summary.csv | Sort-Object LastWriteTimeUtc -Descending)
    $selected = @(Import-Csv -LiteralPath $summaries[0].FullName)
    if ($selected.Count -ne 1 -or $selected[0].Migration -ne 'SiteC') { throw 'Selection/deduplication failed.' }
    $output = Invoke-TestCmd -Command ('""{0}" -PlanOnly -MigrationNames SiteA,SiteC -ComparisonMode FilesOnly"' -f $cmd)
    if ($script:CmdExitCode -ne 0 -or -not ($output -match '2 comparisons')) { throw "CMD plan argument forwarding failed: $output" }
    $output = Invoke-TestCmd -Command ('""{0}" -MigrationNames "Site B" -ComparisonMode FilesOnly"' -f $cmd)
    if ($script:CmdExitCode -ne 1) { throw "CMD batch failure exit was lost: $output" }
    $latest = Get-ChildItem -LiteralPath (Join-Path $root 'Migrations\logs') -Recurse -Filter summary.csv | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    $cmdSummary = @(Import-Csv -LiteralPath $latest.FullName)
    if ($cmdSummary.Count -ne 1 -or $cmdSummary[0].Migration -ne 'Site B' -or $cmdSummary[0].Action -ne 'CompareFiles' -or $cmdSummary[0].ExitCode -ne '9') { throw 'CMD did not run the selected failing comparison.' }
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $project 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1'),[ref]$tokens,[ref]$errors)
    $definition = $ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Open-DirectoryInExplorer'},$true)
    Invoke-Expression $definition.Extent.Text
    $script:LauncherNonInteractive = $true
    $script:ExplorerTouched = $false
    function Resolve-Path { $script:ExplorerTouched = $true; throw 'Explorer path resolution must not run in batch mode.' }
    function Invoke-Item { $script:ExplorerTouched = $true; throw 'Explorer must not open in batch mode.' }
    function Start-Process { $script:ExplorerTouched = $true; throw 'Explorer must not start in batch mode.' }
    Open-DirectoryInExplorer -Path (Join-Path $root 'missing')
    if ($script:ExplorerTouched) { throw 'Explorer helper was not suppressed in batch mode.' }
    'Offline comparison batch tests passed: discovery, phases, sequential jobs, failures, logs, selection, authentication flags, and CMD forwarding.'
}
finally {
    if (Test-Path -LiteralPath $root -PathType Container) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA3/FdQ9DwsVa2T
# ycIysypUACCoet6wvSyUzwiKjXlDuaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDRlCPIH8q64oE/SZhcPjWh
# Q26U1KvoBJ+z33kkqhG9ZTANBgkqhkiG9w0BAQEFAASCAYBroGWmCxiGkIx0MTc5
# ZZj7RTsULSsO8XVLyTZRyxGvHuce0x0fMQM+3maKamvHVvTPZTAKjWPujxCN4rfK
# 3nOSFW+Fwl1d3KKXXLfCf+xYIc2CnzrbtEV7JUu9u9OokLguXwWWtG00TFpbiGBP
# wXviHCI/pEWOKXQflY39a/FE+t5KRh/ohp5lvNvOdjy6UuKXLOKihFrVk0fVzy/7
# bqVYMbhCoR72EVUpprfpqKjT03JUSATfAXhrGqD6/an9T2vR9Gohvvn7jmHZkPvI
# 2Fa/J8EFspxmERn495+RXNNasmXye1UwYnSj5boFEvGX+zpnA4YCeTxEGm5ijN27
# KmumT3ke4baihi1bYnmIUzL/JLxDDgj5E7jcgHjknclEPYZUpcI/FdiVMLZ31M6P
# 9NHJHr05Tdg5zwyF9rPvdIaQejimAmVmnAA8LmNPataBjFqWmaAxZEPCmrzt89c6
# koiqYLAcClgpIkSWiJW2mRw6qTKu4bfe9aAjGgLiBHatEMw=
# SIG # End signature block
