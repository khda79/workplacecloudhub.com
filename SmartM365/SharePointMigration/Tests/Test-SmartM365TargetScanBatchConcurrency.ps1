<#
.SYNOPSIS
    Verify destination batch concurrency with offline child processes.
.VERSION
    1.0.1
#>
#requires -Version 7.4
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$testRoot = Join-Path $env:TEMP ('SmartM365-TargetBatchTest-' + [guid]::NewGuid().ToString('N'))
$testRoot = [IO.Path]::GetFullPath($testRoot)
if (-not $testRoot.StartsWith([IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Unsafe test directory.'
}
$batchPath = Join-Path $PSScriptRoot '..\Scripts\Launchers\Generic\SmartM365-SharePointMigration-TargetScanBatch.ps1'
$source = ([IO.File]::ReadAllText($batchPath) -split '# SIG # Begin signature block')[0]
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$errors)
if (@($errors).Count) { throw 'Batch parser errors.' }
# Authentication is excluded from this scheduler test; no tenant or store is accessed.
$certificateFunction = $ast.Find({ param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Assert-CertificateAuth'
}, $true)
$source = $source.Replace($certificateFunction.Extent.Text, 'function Assert-CertificateAuth { param([string]$Root) }')

try {
    [void](New-Item -ItemType Directory -Path $testRoot)
    $runnerPath = Join-Path $testRoot 'runner.ps1'
    $batchCopy = Join-Path $testRoot 'batch.ps1'
    [IO.File]::WriteAllText($batchCopy, $source)
    [IO.File]::WriteAllText($runnerPath, @'
param([string]$Batch, [string]$Root, [string]$Mode, [int]$Limit, [int]$Delay, [switch]$Plan)
$ErrorActionPreference = 'Stop'
function Get-Module {
    param([switch]$ListAvailable, [string]$Name)
    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = $Name } }
    Microsoft.PowerShell.Core\Get-Module @PSBoundParameters
}
function Start-Sleep { param([int]$Seconds) Microsoft.PowerShell.Utility\Start-Sleep -Milliseconds 100 }
function Start-Process {
    param([string]$FilePath, [string[]]$ArgumentList, [string]$WorkingDirectory, [string]$WindowStyle,
        [switch]$PassThru, [string]$RedirectStandardOutput, [string]$RedirectStandardError)
    $id = [guid]::NewGuid().ToString('N')
    [pscustomobject]@{
        Args = $ArgumentList; Window = $WindowStyle
        Redirected = $PSBoundParameters.ContainsKey('RedirectStandardOutput'); Time = [datetime]::UtcNow.ToString('o')
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Root "launch-$id.json")
    # Exercise real processes without opening test consoles or authentication prompts.
    Microsoft.PowerShell.Management\Start-Process -FilePath $FilePath -ArgumentList $ArgumentList `
        -WorkingDirectory $WorkingDirectory -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput (Join-Path $Root "$id.out") -RedirectStandardError (Join-Path $Root "$id.err")
}
$options = @{ ProjectRoot=$Root; LauncherPath=(Join-Path $Root 'stub.ps1'); AuthMode=$Mode; MaxParallel=$Limit; LaunchDelaySeconds=$Delay }
if ($Plan) { $options.PlanOnly = $true }
$global:LASTEXITCODE = 0
& $Batch @options
exit $LASTEXITCODE
'@)
    foreach ($scenario in @(
        @{ Name='interactive-default'; Mode='Interactive'; Limit=0; Expected=2; Delay=0; Fail=$false; Plan=$false },
        @{ Name='certificate-default'; Mode='Certificate'; Limit=0; Expected=2; Delay=0; Fail=$false; Plan=$false },
        @{ Name='interactive-serial'; Mode='Interactive'; Limit=1; Expected=1; Delay=0; Fail=$false; Plan=$false },
        @{ Name='interactive-failure'; Mode='Interactive'; Limit=2; Expected=2; Delay=1; Fail=$true; Plan=$false },
        @{ Name='interactive-plan'; Mode='Interactive'; Limit=2; Expected=2; Delay=0; Fail=$false; Plan=$true },
        @{ Name='certificate-plan'; Mode='Certificate'; Limit=2; Expected=2; Delay=0; Fail=$false; Plan=$true },
        @{ Name='interactive-reject-three'; Mode='Interactive'; Limit=3; Expected=0; Delay=0; Fail=$false; Plan=$false }
    )) {
        $root = Join-Path $testRoot $scenario.Name
        foreach ($name in @('SiteA','SiteB','SiteC')) {
            $dir = Join-Path $root "Migrations\$name"
            [void](New-Item -ItemType Directory -Path $dir -Force)
            "@{ Name='$name'; Target=@{ Type='SPO'; SiteUrl='https://target.example.test/sites/$name' } }" |
                Set-Content -LiteralPath (Join-Path $dir 'migration.config.psd1')
        }
        [IO.File]::WriteAllText((Join-Path $root 'stub.ps1'), @'
param([string]$MigrationName, [string]$Action, [switch]$UseCertificate, [string]$RunResultPath, [string]$RunResultId)
$start = [datetime]::UtcNow
Start-Sleep -Milliseconds 1800
[pscustomobject]@{ Migration=$MigrationName; Action=$Action; Start=$start.ToString('o'); End=[datetime]::UtcNow.ToString('o'); Certificate=[bool]$UseCertificate } |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot "$Action-$MigrationName.event.json")
$log=Join-Path $PSScriptRoot "$Action-$MigrationName.log"
$csv=Join-Path $PSScriptRoot "$Action-$MigrationName.csv"
Set-Content -LiteralPath $log 'offline scan log'
Set-Content -LiteralPath $csv 'offline scan output'
@{ RunId=$RunResultId; Migration=$MigrationName; Action=$Action; LogPath=$log; OutputCsv=$csv } |
    ConvertTo-Json | Set-Content -LiteralPath $RunResultPath
if ((Test-Path -LiteralPath (Join-Path $PSScriptRoot 'fail')) -and $MigrationName -eq 'SiteB' -and $Action -eq 'ScanTargetFiles') { exit 7 }
'@)
        if ($scenario.Fail) { [IO.File]::WriteAllText((Join-Path $root 'fail'), '') }
        $argsList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',$runnerPath,'-Batch',$batchCopy,'-Root',$root,
            '-Mode',$scenario.Mode,'-Limit',[string]$scenario.Limit,'-Delay',[string]$scenario.Delay)
        if ($scenario.Plan) { $argsList += '-Plan' }
        $output = & (Get-Command pwsh).Source @argsList 2>&1
        $code = $LASTEXITCODE
        if ($scenario.Expected -eq 0) {
            if ($code -eq 0 -or @(Get-ChildItem -LiteralPath $root -Filter 'launch-*.json').Count) { throw 'Limit above two was accepted.' }
            continue
        }
        if ($code -ne $(if ($scenario.Fail) { 1 } else { 0 })) { throw "Unexpected exit for $($scenario.Name): $code; $output" }
        if ($scenario.Plan) {
            if (@(Get-ChildItem -LiteralPath $root -Filter 'launch-*.json').Count -or
                (Test-Path -LiteralPath (Join-Path $root 'Migrations\logs')) -or
                -not ($output -match 'maximum parallel=2')) { throw 'PlanOnly started work or reported the wrong limit.' }
            continue
        }
        $events = @(Get-ChildItem -LiteralPath $root -Filter '*.event.json' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json })
        $launches = @(Get-ChildItem -LiteralPath $root -Filter 'launch-*.json' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json })
        if ($events.Count -ne 6 -or $launches.Count -ne 6) { throw 'Not all scans completed.' }
        $points = foreach ($event in $events) {
            [pscustomobject]@{ Time=[datetime]$event.Start; Delta=1 }
            [pscustomobject]@{ Time=[datetime]$event.End; Delta=-1 }
            if ($event.Certificate -ne ($scenario.Mode -eq 'Certificate')) { throw 'Wrong child authentication mode.' }
        }
        $active = 0; $peak = 0
        foreach ($point in ($points | Sort-Object Time,Delta)) { $active += $point.Delta; $peak = [Math]::Max($peak,$active) }
        if ($peak -ne $scenario.Expected) { throw "Unexpected concurrency for $($scenario.Name): $peak" }
        $filesEnd = ($events | Where-Object Action -EQ ScanTargetFiles | ForEach-Object { [datetime]$_.End } | Sort-Object | Select-Object -Last 1)
        $permissionsStart = ($events | Where-Object Action -EQ ScanTargetPermissions | ForEach-Object { [datetime]$_.Start } | Sort-Object | Select-Object -First 1)
        if ($permissionsStart -lt $filesEnd) { throw 'Permissions started before the file phase finished.' }
        foreach ($launch in $launches) {
            if ($scenario.Mode -eq 'Interactive') {
                if ($launch.Window -ne 'Normal' -or $launch.Redirected -or $launch.Args -contains '-NonInteractive' -or $launch.Args -contains '-UseCertificate') {
                    throw 'Interactive scan did not retain a visible console and interactive authentication.'
                }
            }
            elseif ($launch.Window -ne 'Hidden' -or -not $launch.Redirected -or $launch.Args -notcontains '-NonInteractive' -or $launch.Args -notcontains '-UseCertificate') {
                throw 'Certificate scan did not retain background execution.'
            }
        }
        if ($scenario.Delay) {
            foreach ($phase in @('ScanTargetFiles','ScanTargetPermissions')) {
                $times = @($launches | Where-Object { $_.Args -contains $phase } | ForEach-Object { [datetime]$_.Time } | Sort-Object)
                for ($i=1; $i -lt $times.Count; $i++) {
                    if (($times[$i]-$times[$i-1]).TotalSeconds -lt $scenario.Delay) { throw 'Launch spacing was not respected.' }
                }
            }
        }
        $summary = @(Get-ChildItem -LiteralPath (Join-Path $root 'Migrations\logs') -Recurse -Filter summary.csv | Import-Csv)
        if ($summary.Count -ne 6) { throw 'Incomplete batch summary.' }
        foreach ($row in $summary) {
            if (-not $row.RunLog -or -not $row.OutputCsv -or $row.OutputLog -ne $row.RunLog -or
                -not (Test-Path -LiteralPath $row.RunLog) -or -not (Test-Path -LiteralPath $row.OutputCsv)) {
                throw 'Interactive or certificate batch lost the actual scan log/output receipt.'
            }
        }
        $failures = @($summary | Where-Object Status -EQ FAILED)
        if ($failures.Count -ne $(if ($scenario.Fail) { 1 } else { 0 }) -or
            ($scenario.Fail -and $failures[0].ExitCode -ne '7')) { throw 'Child failure was not propagated.' }
        "PASS $($scenario.Name): peak=$peak; completed=$($summary.Count); failures=$($failures.Count)"
    }
    'Offline destination batch concurrency tests passed.'
}
finally {
    if (Test-Path -LiteralPath $testRoot -PathType Container) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDo2CH4M9SNjCmp
# j8z/UaCPh492yUrVuVqkOS1yPL/rRqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCpCkCCzQ40Tws+93llyfI9
# 1OirQFrXlVOoawn1RXzS5zANBgkqhkiG9w0BAQEFAASCAYASS97q/QQOFCkQ4CyU
# 7pKKMYJ0ydZQtjY7XQpXxpaQtc1eogeU9E5EPhj8aK4YiVIepMiCDr9lceRBCQnH
# AIxFtxjI+aonSWhRlrShcR27vq+2Gsi77gmPuBSDzykveWTZLVZQQ2om4w/bZBuN
# 3/rD/1aWn9zvwclF9BhLyLFujLu/MhbUylOpPsFRMd20LS+mCl98jsH+CrAk+KFL
# PWdnU68qmQG9KkOpuLgdtUDh8XIBRBrfIVGVXuDSFgI7Y02W1E9e0572z+TU7cIw
# H6qhWGIIppCcO9qml/e9i1AYg1JUgvwri+nUG8w3QW7UV+o5TEU5yh6JDYfS5XGW
# h/GioFy2s479tKHlX7XJkt22QSVyEntPwBYr8Nmp+QpNCBtQzrhLBsc40fJ2QJ1j
# ILK1mTQxocMls04RPz3jNcuFkWD5GqNaDZgGnqJ+D0zQ198pUvliqXKv8O3fmtjF
# lWxAvW9fDGmeRCSP7M8UziTV/LtM8+lsBOxjMVCeUf2xcWo=
# SIG # End signature block
