<#
.SYNOPSIS
    Compare source and destination inventories for all configured migrations.
.DESCRIPTION
    Runs file comparisons first, then permission comparisons, sequentially.
    Uses the existing migration launcher to select the latest inventories,
    validate scan evidence, refresh the Entra cache, and generate reports.
    Continues after individual failures and returns a nonzero batch exit code.
.VERSION
    1.0.2
.EXAMPLE
    .\SmartM365-SharePointMigration-ComparisonBatch.ps1 -PlanOnly
.EXAMPLE
    .\SmartM365-SharePointMigration-ComparisonBatch.ps1 -ComparisonMode Both
#>
#requires -Version 7.4
[CmdletBinding()]
param(
    [string[]]$MigrationNames = @(),
    [ValidateSet('FilesOnly', 'PermissionsOnly', 'Both')]
    [string]$ComparisonMode = 'Both',
    [ValidateSet('Interactive', 'Certificate')]
    [string]$AuthMode = 'Interactive',
    [switch]$PlanOnly,
    [ValidatePattern('^\d{8}-\d{6}-[a-f0-9]{8}$')]
    [string]$BatchId = '',
    [switch]$Force,
    [string]$ProjectRoot = (Join-Path $PSScriptRoot '..\..\..'),
    [string]$LauncherPath = ''
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$script:BatchLogPath = $null
$script:Results = [Collections.Generic.List[object]]::new()

function Write-Host {
    param([Parameter(Position = 0, ValueFromRemainingArguments = $true)][object[]]$Object)
    foreach ($part in $Object) {
        foreach ($line in ([string]$part -split "`r?`n")) {
            Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line)
        }
    }
}

function Write-BatchLine {
    param([string]$Message, [string]$ConsoleLog = '')
    foreach ($line in ($Message -split "`r?`n")) {
        $entry = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line
        Microsoft.PowerShell.Utility\Write-Host $entry
        if ($script:BatchLogPath) { Add-Content -LiteralPath $script:BatchLogPath -Value $entry -Encoding utf8 }
        if ($ConsoleLog) { Add-Content -LiteralPath $ConsoleLog -Value $entry -Encoding utf8 }
    }
}

function Get-ComparisonMigrations {
    param([string]$Root, [string[]]$Names)
    $directory = Get-SmartM365MigrationsRoot -ProjectRoot $Root
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { throw "Migrations directory not found: $directory" }
    $available = @{}
    foreach ($item in (Get-ChildItem -LiteralPath $directory -Directory)) {
        if ($item.Name -eq 'NewMigration') { continue }
        $path = Join-Path $item.FullName 'migration.config.psd1'
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $configError = ''
        try {
            $config = Import-PowerShellDataFile -LiteralPath $path
            if ($config.Name -ne $item.Name) { throw "Migration name differs from folder name: $path" }
        }
        catch { $configError = $_.Exception.Message }
        $available[$item.Name] = [pscustomobject]@{ Name=$item.Name; ConfigError=$configError }
    }
    if ($Names.Count) {
        $seen = @{}
        foreach ($name in $Names) {
            if (-not $available.ContainsKey($name)) { throw "Unknown migration: $name" }
            if ($seen.ContainsKey($name)) { continue }
            $seen[$name] = $true
            $available[$name]
        }
    }
    else { $available.Values | Sort-Object Name }
}

$ProjectRoot = [IO.Path]::GetFullPath($ProjectRoot)
. (Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-Storage.ps1')
if ([string]::IsNullOrWhiteSpace($LauncherPath)) {
    $LauncherPath = Join-Path $ProjectRoot 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1'
}
$requested = @($MigrationNames | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$actions = @()
if ($ComparisonMode -in @('FilesOnly','Both')) { $actions += 'CompareFiles' }
if ($ComparisonMode -in @('PermissionsOnly','Both')) { $actions += 'ComparePermissions' }
$summaryPath = $null
$interrupted = $true
try {
    if (-not $PlanOnly) {
        $id = if ($BatchId) { $BatchId } else { '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N').Substring(0,8) }
        $batchRoot = Get-SmartM365MigrationLogsRoot -ProjectRoot $ProjectRoot -ChildPath "comparison-batches\$id"
        Write-BatchLine ("Creating comparison batch logs: {0}" -f $batchRoot)
        [void](New-Item -ItemType Directory -Path $batchRoot)
        $script:BatchLogPath = Join-Path $batchRoot 'batch.log'
        $summaryPath = Join-Path $batchRoot 'summary.csv'
    }
    Write-BatchLine 'Comparison batch preflight started.'
    $pwsh = (Get-Command pwsh -ErrorAction Stop).Source
    $LauncherPath = [IO.Path]::GetFullPath($LauncherPath)
    if (-not (Test-Path -LiteralPath $LauncherPath -PathType Leaf)) { throw "Migration launcher not found: $LauncherPath" }
    $migrations = @(Get-ComparisonMigrations -Root $ProjectRoot -Names $requested)
    if ($migrations.Count -eq 0) { throw 'No configured migrations found.' }
    Write-BatchLine ("Comparison plan: {0} migrations; {1} comparisons; authentication={2}; sequential execution." -f
        $migrations.Count, ($migrations.Count * $actions.Count), $AuthMode)
    foreach ($action in $actions) {
        foreach ($migration in $migrations) {
            Write-BatchLine ("PLAN {0} {1}" -f $migration.Name, $action)
            if ($migration.ConfigError) { Write-BatchLine ("INVALID {0}: {1}" -f $migration.Name, $migration.ConfigError) }
        }
    }
    if ($PlanOnly) {
        Write-BatchLine 'Plan only: no comparisons, authentication, or output created.'
        if (@($migrations | Where-Object ConfigError).Count) { exit 1 }
        return
    }
    if ($Force) { Write-BatchLine 'WARNING: -Force allows comparisons outside the configured scan age limits.' }
    Write-BatchLine 'Using the latest source/target inventories selected by each migration launcher. Missing or invalid evidence will fail that comparison.'
    if ($AuthMode -eq 'Interactive' -and $ComparisonMode -ne 'FilesOnly') {
        Write-BatchLine 'Permission comparisons may request Entra sign-in if the users cache needs refreshing.'
    }
    foreach ($action in $actions) {
        Write-BatchLine ("Starting phase: {0}" -f $action)
        foreach ($migration in $migrations) {
            $started = Get-Date
            $consoleLog = Join-Path $batchRoot ("{0}-{1}.console.log" -f $migration.Name,$action)
            $status = 'FAILED'; $code = 1; $message = ''
            Write-BatchLine -Message ("START {0} {1}" -f $migration.Name,$action) -ConsoleLog $consoleLog
            try {
                if ($migration.ConfigError) { throw $migration.ConfigError }
                $childArguments = @('-NoLogo','-NoProfile')
                if ($AuthMode -eq 'Certificate') { $childArguments += '-NonInteractive' }
                $childArguments += @('-File',$LauncherPath,'-MigrationName',$migration.Name,'-Action',$action,'-NonInteractive')
                if ($AuthMode -eq 'Certificate') { $childArguments += '-UseCertificate' }
                if ($Force) { $childArguments += '-Force' }
                $global:LASTEXITCODE = 0
                & $pwsh @childArguments 2>&1 | ForEach-Object {
                    Write-BatchLine -Message ("{0} {1}: {2}" -f $migration.Name,$action,[string]$_) -ConsoleLog $consoleLog
                }
                $code = $LASTEXITCODE
                if ($null -eq $code) { $code = 1 }
                if ($code -eq 0) { $status = 'SUCCESS' }
                else { $message = "Comparison launcher returned exit $code. See console and migration logs." }
            }
            catch {
                $message = $_.Exception.Message
                Write-BatchLine -Message ("ERROR {0} {1}: {2}" -f $migration.Name,$action,$message) -ConsoleLog $consoleLog
            }
            finally {
                $script:Results.Add([pscustomobject]@{
                    Migration=$migration.Name; Action=$action; Status=$status; ExitCode=$code
                    Started=$started.ToString('o'); Finished=(Get-Date).ToString('o'); ConsoleLog=$consoleLog; Error=$message
                })
                $script:Results.ToArray() | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding utf8
                Write-BatchLine ("{0}: {1} {2} (exit {3})" -f $status,$migration.Name,$action,$code)
            }
        }
    }
    $interrupted = $false
}
catch {
    Write-BatchLine ("Comparison batch failed: {0}" -f $_.Exception.Message)
    throw
}
finally {
    if ($summaryPath) {
        $script:Results.ToArray() | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding utf8
        $failed = @($script:Results | Where-Object Status -NE SUCCESS).Count
        Write-BatchLine ("Batch finished: {0} successful; {1} failed; interrupted={2}; summary: {3}" -f
            ($script:Results.Count-$failed),$failed,$interrupted,$summaryPath)
    }
}
if ($interrupted -or @($script:Results | Where-Object Status -NE SUCCESS).Count) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBx7j6x3HJxeFeS
# CdXwc4/x7qUaBT8BxSf7ktmmwHEz4KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIi5TVEZDdut7HntXJjBfhrwTTpmTL+TTiBh1DbZ8G0XMA0GCSqG
# SIb3DQEBAQUABIIBgFEPxDYHzGVO21VyT+H3IYkTHGbJQE9tjfYXJCsYprROAx/x
# md8aqEWlZBn8gIFz7ZT17Xk0TS0nraj76offbF2krKIeQUbiiPW1Ez4g2U01YjdR
# JGZuoT2q+A4uK/DG+Pc6SiYjdRlmYk5vEjEzZeweOXL8hHRLcQhhL1GADfs5MB0d
# 6vh5xHYm9G8zbNZUwDbGhyZRUI4sH1JSOQHx/g/HH8ABrHg6Re878z0w4f6HPLAD
# 98JMUXVanIQ73zfCnXbmdH8uf0vL8pAmVmIlwup2bK20kCbIFx0NmorfMbTAXoD8
# W0sK/0teCK/rgFsuF32gPiFqzChNtsMgbEaUzs3XdP8ahhYu5gjPl5ZZDs2cllpg
# VyP1p7loCFGJVd5KtugdAhXOPIE4pektJW0BA2d8AJWhsYX6kfvG0FP6u0P07Fau
# CiCfohtRG7MxG2Y/hQUAUEOBSP3e2wezhrTQzPj377SWR1FTMDof/jFNIzYBpg4F
# XC/jHjNhDEYNt9c3wKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzNaMC8GCSqGSIb3DQEJBDEiBCC2J+/jt6RrF1MN7uDrDB5vJXPQNMUitu92phlN
# aNe8xDANBgkqhkiG9w0BAQEFAASCAgAYuax80Yc/B+I70ewml7O0wsDIyckSdW8X
# df+Tv1VGzlnaB/jp64x9hmcNLO0guCskXJeNF4dtSrsy/zis090TnvQzGBqNIHDq
# DVqb/ru2l3eUyCE0UP8wmYf2hiAYwwk23fDZ0N8Y2lsLRdWmpvEXsZ0/IMoChs6q
# B3yGanqiuBD48zurpF6bq83veol13XXG5VVTK917OrixqR6Y1AIkDreMjgAyZmRk
# axM1wPpWfMpgHd+znwvDoxG02DHZdxAM6I08YV7mPl2IlGs8xY0IVBEzojEyGONy
# Zc+jdzQXN8ZGkgiCGR0gq+k4nQbqiVME+lV3sLLT6VSwJc4AChHNwog8oJPiKP/i
# /JIMRfgq2mK/9HkEO+ELHCMnSiV2KN4eHFYPPCYShMPIHatDdaw8NwsbIxckHVlm
# ZXRK8aJ7KYBTj5EivSokowE1aQvC8baVcdRVE8r5oel6AbJVCK+HtQLt0sHaIDAA
# IDdfFqa7vD36y76jDk/dgfek9Y8N7rNC0t4wIW7AKZ8SI+1WEA2GOptumhas2G7O
# QRvFSvVSbohcasBSI9tW1rmyIbpJoMSZwmYajy9SxtdcEl1ALqQQjWyTtUfmXhMM
# ifWhPC5yKeJgMxpzQ9YuK/LqTBEPOzgbda+NwUBI5TI3nx+9j9QTYkRkTbMSLjYm
# mY38qJOPjQ==
# SIG # End signature block
