<#
.SYNOPSIS
    Start only the read-only farm diagnostic in Windows PowerShell 5.1.
.DESCRIPTION
    Lists migration projects when -Project is omitted, then finds the selected
    project's latest ShareGate five-minute access CSV. DryRun is the default.
.VERSION
    1.0.7
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Project = '',
    [string]$ToolkitRoot = '',
    [string]$ShareGatePeaksCsv = '',
    [ValidateRange(1,1440)][int]$WindowMinutes = 30,
    [switch]$Run,
    [switch]$PreviewOnly
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:FarmLauncherLifecycle = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath -Action $(if ($Run) { 'Farm diagnostic Run' } else { 'Farm diagnostic DryRun' }) -Migration $Project
$script:FarmLauncherFailure = $null
$script:FarmLauncherStatus = 'CANCELLED'
$script:FarmLauncherPreviousExpectedId = [string]$env:SPMIG_FARM_DIAG_EXPECTED_ID
$script:FarmLauncherPreviousCompletedId = [string]$env:SPMIG_FARM_DIAG_COMPLETED_ID

function Write-FarmLauncherInfo {
    param([string]$Message)
    Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Get-FarmLauncherPeakCount {
    param([string]$Path)
    $count = 0
    $unparsed = 0
    foreach ($row in @(Import-Csv -LiteralPath $Path -ErrorAction Stop)) {
        if ($row.WindowUtc -eq '(unparsed timestamp)') { $unparsed++; continue }
        if (-not $row.WindowUtc) { continue }
        $stamp = [datetime]::MinValue
        $value = ([string]$row.WindowUtc -replace '\s+UTC$','')
        if (-not [datetime]::TryParseExact($value,'yyyy-MM-dd HH:mm',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$stamp)) {
            throw "Invalid WindowUtc in ShareGate CSV: $($row.WindowUtc)"
        }
        $count++
    }
    if ($count -eq 0 -and $unparsed) { throw 'CSV has no UTC windows: ShareGate report dates were not converted from the exporter time zone.' }
    if ($count -eq 0) { throw "No valid five-minute UTC window in ShareGate CSV: $Path" }
    return $count
}

function Get-FarmLauncherProjectInfo {
    param([IO.DirectoryInfo]$Directory)
    $diagnostics = Join-Path $Directory.FullName 'ShareGate\Diagnostics'
    $info = [pscustomobject]@{
        Name = $Directory.Name
        CsvPath = ''
        PeakCount = 0
        Status = 'CSV missing'
        Usable = $false
    }
    if (-not (Test-Path -LiteralPath $diagnostics -PathType Container)) { return $info }
    try {
        $latest = Get-ChildItem -LiteralPath $diagnostics -Recurse -File -Filter 'AccessFailures-5min.csv' -ErrorAction Stop |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 1
        if (-not $latest) { return $info }
        $info.CsvPath = $latest.FullName
        $info.PeakCount = Get-FarmLauncherPeakCount -Path $latest.FullName
        $windowLabel = if ($info.PeakCount -eq 1) { 'window' } else { 'windows' }
        $info.Status = 'CSV {0} UTC ({1} {2})' -f $latest.LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm:ss'),$info.PeakCount,$windowLabel
        $info.Usable = $true
    }
    catch {
        $info.Status = if ($_.Exception.Message -like 'CSV has no UTC windows:*') { 'CSV has no UTC windows' } else { 'CSV invalid or inaccessible' }
    }
    return $info
}

try {
    if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1) {
        throw 'Windows PowerShell 5.1 is required for the farm diagnostic launcher.'
    }
    $root = if ($ToolkitRoot) { $ToolkitRoot } else { Join-Path $PSScriptRoot '..\..' }
    $root = [IO.Path]::GetFullPath($root)
    $migrationsRoot = Join-Path $root 'Migrations'
    if (-not (Test-Path -LiteralPath $migrationsRoot -PathType Container)) {
        throw "Migrations folder is missing: $migrationsRoot. Launch from the shared toolkit or supply -ToolkitRoot."
    }
    if ($Project -and $Project -notmatch '^(?!\.{1,2}$)[A-Za-z0-9][A-Za-z0-9 ._-]*$') {
        throw 'Project must be a migration folder name without path separators.'
    }
    if (-not $Project -and $ShareGatePeaksCsv) {
        throw '-Project is required when -ShareGatePeaksCsv is supplied.'
    }
    $projects = @(Get-ChildItem -LiteralPath $migrationsRoot -Directory -ErrorAction Stop |
        Where-Object { $_.Name -match '^(?!\.{1,2}$)[A-Za-z0-9][A-Za-z0-9 ._-]*$' -and $_.Name -notin @('logs','reports') } |
        Sort-Object Name |
        ForEach-Object { Get-FarmLauncherProjectInfo -Directory $_ })
    if (-not $projects.Count) { throw "No migration project folder found in $migrationsRoot" }
    $selected = $null
    if ($Project) {
        $selected = @($projects | Where-Object { $_.Name -ieq $Project } | Select-Object -First 1)
        if (-not $selected.Count) { throw "Project is not in the migration list: $Project" }
        $selected = $selected[0]
    }
    else {
        Write-FarmLauncherInfo 'Available migration projects:'
        for ($index = 0; $index -lt $projects.Count; $index++) {
            Write-FarmLauncherInfo ('{0,2}. {1} | {2}' -f ($index + 1),$projects[$index].Name,$projects[$index].Status)
        }
        Write-FarmLauncherInfo ' 0. Cancel'
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            $answer = Read-Host ('[{0}] Select a project number' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
            $number = -1
            if ([int]::TryParse([string]$answer,[ref]$number)) {
                if ($number -eq 0) {
                    $script:FarmLauncherStatus = 'CANCELLED'
                    Write-FarmLauncherInfo 'Selection cancelled. No diagnostic started.'
                    return
                }
                if ($number -ge 1 -and $number -le $projects.Count) {
                    $selected = $projects[$number - 1]
                    break
                }
            }
            Write-FarmLauncherInfo "Invalid selection ($attempt/3)."
        }
        if (-not $selected) { throw 'No valid project number was selected.' }
        $Project = $selected.Name
    }
    $script:FarmLauncherLifecycle.Migration = $Project
    if (-not $selected.Usable -and -not $ShareGatePeaksCsv) {
        throw "Project $Project cannot run: $($selected.Status). CSV: $($selected.CsvPath)"
    }
    $diagnosticScript = Join-Path $root 'Scripts\Diagnostics\SmartM365-SharePointMigration-FarmDiagnostic.ps1'
    if (-not (Test-Path -LiteralPath $diagnosticScript -PathType Leaf)) {
        throw "Farm diagnostic script is missing: $diagnosticScript"
    }
    if ($ShareGatePeaksCsv) {
        $peaksPath = [IO.Path]::GetFullPath($ShareGatePeaksCsv)
    }
    else {
        $peaksPath = $selected.CsvPath
    }
    if (-not (Test-Path -LiteralPath $peaksPath -PathType Leaf)) {
        throw "ShareGate peaks CSV is missing: $peaksPath"
    }
    $peakCount = Get-FarmLauncherPeakCount -Path $peaksPath
    $mode = if ($Run) { 'Read-only collection' } else { 'DryRun' }
    Write-FarmLauncherInfo "Host: Windows PowerShell $($PSVersionTable.PSVersion)"
    Write-FarmLauncherInfo "Project: $Project"
    Write-FarmLauncherInfo "Toolkit root: $root"
    $peakWindowLabel = if ($peakCount -eq 1) { 'window' } else { 'windows' }
    Write-FarmLauncherInfo "ShareGate peaks: $peaksPath ($peakCount UTC $peakWindowLabel)"
    Write-FarmLauncherInfo "Mode: $mode; margin: $WindowMinutes minutes"
    if ($PreviewOnly) {
        Write-FarmLauncherInfo 'PreviewOnly completed; no farm command was run.'
        $script:FarmLauncherStatus = 'SUCCESS'
        return
    }
    $parameters = @{
        Project = $Project
        ToolkitRoot = $root
        ShareGatePeaksCsv = $peaksPath
        WindowMinutes = $WindowMinutes
    }
    if (-not $Run) { $parameters['DryRun'] = $true }
    $expectedId = [guid]::NewGuid().ToString('N')
    $env:SPMIG_FARM_DIAG_EXPECTED_ID = $expectedId
    Remove-Item Env:SPMIG_FARM_DIAG_COMPLETED_ID -ErrorAction SilentlyContinue
    $global:LASTEXITCODE = 0
    & $diagnosticScript @parameters
    if ($LASTEXITCODE -ne 0) { throw "Farm diagnostic failed with exit code $LASTEXITCODE." }
    if ([string]$env:SPMIG_FARM_DIAG_COMPLETED_ID -ne $expectedId) {
        Write-FarmLauncherInfo 'Farm diagnostic stopped before completion. No successful result was confirmed.'
        exit 130
    }
    $script:FarmLauncherStatus = 'SUCCESS'
    Write-FarmLauncherInfo 'Farm diagnostic completed.'
}
catch {
    $script:FarmLauncherFailure = $_
    Write-FarmLauncherInfo "ERROR: $($_.Exception.Message)"
    exit 1
}
finally {
    try { Complete-SmartM365MigrationConsoleLifecycle -Context $script:FarmLauncherLifecycle -Failure $script:FarmLauncherFailure -Status $script:FarmLauncherStatus }
    finally {
        if ($script:FarmLauncherPreviousExpectedId) { $env:SPMIG_FARM_DIAG_EXPECTED_ID = $script:FarmLauncherPreviousExpectedId }
        else { Remove-Item Env:SPMIG_FARM_DIAG_EXPECTED_ID -ErrorAction SilentlyContinue }
        if ($script:FarmLauncherPreviousCompletedId) { $env:SPMIG_FARM_DIAG_COMPLETED_ID = $script:FarmLauncherPreviousCompletedId }
        else { Remove-Item Env:SPMIG_FARM_DIAG_COMPLETED_ID -ErrorAction SilentlyContinue }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCFdnRwcWAANeYp
# xnBBFuyXdo8tFzPfg/++dvcn9hWO1aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPmRfTqK3feWYuX+Rzj+1aJUG6CKPZIYrHzo8skwYV5cMA0GCSqG
# SIb3DQEBAQUABIIBgK2BxT1KzoX8iePW3XXzQTgHim2GkQjN5fC62rgZd0ceYXJH
# 8lb7CxWvtwcIb+iR+1xrPlE585CmijaQdwPnsnIfi3bc/JxvIF88RVweND1FYkcB
# r3m07hcIc2bi/kxd14RIZF68YiMfEd3b1HKDXuYcAufVOvdeDCADDd5wOwhMolZD
# 3GQ56tMuYultICpjLLljPOfunhJQY6YwOe6JlRjy/1Gj/06DeEZW5SikEXs2dvIb
# SldoZUhaVBjuSmZrzfANqUr4z/MSDNNNGcNvfxQJ0qEDKZZiZKMQrVExUUgoOMcu
# NaSAX81IvD5LXBtoFShgXX+VrcfRiCemhz2sXUDIMz2efKXNeDu+rRFxCXt8BYcw
# y1tKc+2AEinsnCTYHtcGJSHLXCPyUCu3Rb5nlAs2mh9mTxmZjn7F7HpeUkeDXAzv
# rqTmDPyoLx+gBnjnqg214T02ynxbYnFXy1QSH2p02sfbNMJsscmxTPVLBRUDyNnG
# NUQczK3dh4r6Bq/YwaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgyMDQx
# MDlaMC8GCSqGSIb3DQEJBDEiBCAGAiLAMcRivHvrXDL+hoduBLZkn1Qpx2Y5GNke
# /tTupTANBgkqhkiG9w0BAQEFAASCAgAo3/+dOWF9fOBo4OiRLBQG1WUrHy8R+DbG
# C5rOVj3Ma8JR2cUQXB6AkR7AksHtslL6GWMn700sFO38JpkOf9jFykBV0rFbiA3O
# RxRNtkKY3sNj7Bcb2BptTWX+X2JCTpoXdh5nIbFQ7tb9V6c+SIRYIQXSneEC1YN6
# zxfybM46IQtjO/Ws73KcoeJ+eBhBgFdgllurdGnnkeMRcOpc2N/Q36PkMTHRP23q
# 4tIptb8ClL4Z7mAX/aRf17PfBHCGi6XlJPGQwjt/h+nZ4vYM7zRJZawYBGMsvx/F
# 4VcX800izROJPFMQiLGI7B4chNSu3RBU52dHenXrhibVr6W0bliraglo0n6hkvz1
# Lv7h8eo4/q8BDrQeAGKb2/qDFGmTAo63StWPiKEdCA+5atgflPKhj3mjcjGVtC2f
# 6nAcX6R2zU4xsjzUJ2XbPRevVXdUwV+ByfeAh6naz0FnRPzu4BPHqjiN2Ca335SK
# cg20+PqXxyQ39to8ppNkB3xHXTfThQ8SrOsvQUF3bGIy2lFlWxckfSkEP6WesI9S
# 2/cJTeRlnbznzs2rMZ3ifxw1AK7m8BrRGB6PIloORDlkZAeD1ie165xaQV6i2tRM
# kg7c/VX4n7zUgYim5P3DvskwToGPXPJFnCfPTiDPQFUPl2KlrAWdknk6amG8HrSW
# u8udA7Wsqw==
# SIG # End signature block
