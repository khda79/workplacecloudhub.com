<#
.SYNOPSIS
    Verify comparison batch discovery, failure handling, logging and CMD forwarding offline.
.VERSION
    1.0.1
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
    Copy-Item -LiteralPath (Join-Path $project 'Scripts\Launchers\SmartM365-SharePointMigration-Storage.ps1') -Destination (Split-Path $generic)
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDTlEEZasrqRzzp
# XJQA+aSLqq/+bOwpyj3vSIiZqdLpMaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGCqi07EbqaT5W8AE4TbkaBZP/y/yc8MaOzac6VFjPbeMA0GCSqG
# SIb3DQEBAQUABIIBgGK7uqENTHtbQw+PO93qH4i4F+2WzMt0huB4INmHPwN3Y3at
# 5NbSrHEDiB6WbTBI9sJs1p5RoNfNlluChsKUiCZl2glZRBXQAuYIVmt4u78f71Cs
# g1EoyB9KUXgKZWf2j2u6bwB0RDtqWa0KJNUyeoS+3fCPhj3gtC+GsnXxdj+TTGxw
# wNam7rfKzxOx0PKy1S9PStXd26pLIglFUFw+ieRxq9qVvt2Jh3TTJUCRpwPDJyfs
# 7CkN3xRYztmDZT6g/YpAexnCc8D4MLIm5B5VzwpoZWq+hlTjBw3qJyehvPHiw0kM
# /SUI+xelf04+sH62mYyO7dhmsx17lMQPT/cgPOg4f2+bPpUcchGxGLEc+qD9maLB
# Uokaczo5p1hx/OSxCVibBLEmJnVuadjPlWKa5LoFOXndrXd8HWUigpWsU43O/q0c
# 89fwiZo9cm7Ti9496aKJ+KU8O/SDsCUymJ58D2o2Y9QAIsGWdyKD/Meoo1cHn+8V
# XVvlB19FJeBXJYrZwqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzdaMC8GCSqGSIb3DQEJBDEiBCAmMngw/qEPOYTXAPYuRNtAixU2eTM5KscXCsT3
# ojWCdzANBgkqhkiG9w0BAQEFAASCAgAwgoJ0TdD7uWTjdGPACLOOszMViViunOvF
# P358WFdpuzziwCJ+iR8u+F5AMhgr2xDCylZW3y+P6aeR68tsxWqsFk+TLDuzu1Nj
# lntuBGZmnbjdLmfBlevSdKWU8mc5YtmYTk2Z5QcBdqDTOvn8EuOdzGa1eFgQR5u8
# wo5j4+3EkqMhAiA9qwbF5EFzEs6Lr4Z+2VItgmLKrmiASHwoId7Qj168FInnEyfq
# k0PCkanTjlksttzI9h3HCZs4RVR4dR8U2KICQSt7joPvsWyN/xrJkv5dE77aSauP
# 6HsFarT8kEJ/F14FI8Ftxq3yetWtdT0G2rmwGE77LPCzhU2DK1nacz3yUtHYZWte
# 1K7VEnWs7GxeSe4nIAJCULhXzXsI/jj4ok5MvDQvJFmKsCSdoM3C42YOk4RlmLD2
# E9nzn4JNzRdKLip1rpUJ4vLz/akE1VR9PH5CjlL4A2fAkMCsPqXu9+ZVUchqcl0m
# A2C1iACbYvPmjQ1dTJ9+n2J1m2G7M0mJsyVoaP3l0A+HWCuRk2PVmflighNcqHUa
# 4DSmVmJtil3N+7LXOFMbwIHQEfEh2Hl9qqUb8bx28ROQ+wDvfRuQbyW18lJkoSxS
# FNsRAHXzpRjIDJmL0ziD8QpyoXhnMGxGzTwRQYdBANJAl7O/jRvhZGpj+p1Lf4ng
# AV8FB7YzVA==
# SIG # End signature block
