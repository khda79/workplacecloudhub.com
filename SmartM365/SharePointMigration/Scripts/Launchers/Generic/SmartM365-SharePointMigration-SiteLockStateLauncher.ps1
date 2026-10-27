<#
.SYNOPSIS
    Selects one migration site collection and an explicit lock state.
.DESCRIPTION
    Reads migration mapping pairs, shows a review, and requires EXECUTE before
    calling the SharePoint Server or SharePoint Online lock operation.
.VERSION
    1.0.1
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ToolkitRoot = '',
    [string]$MigrationName = '',
    [ValidateSet('Source', 'Destination')][string]$Side,
    [ValidateSet('ReadOnly', 'Unlock', 'NoAccess')][string]$LockState,
    [switch]$PreviewOnly
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:LockLauncherContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath -Action 'Set site collection lock state'
$script:LockLauncherFailure = $null
$script:LockLauncherStatus = 'CANCELLED'

function Write-LockLauncherMessage {
    param([string]$Message)
    Microsoft.PowerShell.Utility\Write-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Read-LockLauncherChoice {
    param([string]$Prompt, [int]$Maximum)
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $answer = Read-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Prompt)
        $number = -1
        if ([int]::TryParse([string]$answer, [ref]$number) -and $number -ge 0 -and $number -le $Maximum) {
            return $number
        }
        Write-LockLauncherMessage ('Invalid selection ({0}/3).' -f $attempt)
    }
    throw 'No valid selection was entered.'
}

function Get-MappingSiteUrl {
    param([string]$Value, [string]$MappingPath)
    $uri = $null
    if (-not [uri]::TryCreate($Value, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -notin @('http', 'https') -or $uri.UserInfo -or $uri.Query -or $uri.Fragment) {
        throw "Invalid site URL in mapping: $MappingPath"
    }
    return $uri.AbsoluteUri.TrimEnd('/')
}

try {
    if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1) {
        throw 'Windows PowerShell 5.1 is required for the site lock launcher.'
    }
    $root = if ($ToolkitRoot) { $ToolkitRoot } else { Join-Path $PSScriptRoot '..\..\..' }
    $root = [IO.Path]::GetFullPath($root)
    $storageHelper = Join-Path $root 'Scripts\Launchers\SmartM365-SharePointMigration-Storage.ps1'
    if (-not (Test-Path -LiteralPath $storageHelper -PathType Leaf)) { $storageHelper = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-Storage.ps1' }
    . $storageHelper
    $migrationsRoot = Get-SmartM365MigrationsRoot -ProjectRoot $root
    if (-not (Test-Path -LiteralPath $migrationsRoot -PathType Container)) {
        throw "Migrations folder is missing: $migrationsRoot"
    }
    $migrations = @(Get-ChildItem -LiteralPath $migrationsRoot -Directory -ErrorAction Stop |
        Where-Object {
            $_.Name -notin @('logs', 'reports', '_Template') -and
            (Test-Path -LiteralPath (Join-Path $_.FullName 'migration.config.psd1') -PathType Leaf) -and
            (Test-Path -LiteralPath (Join-Path $_.FullName 'migration.mapping.txt') -PathType Leaf)
        } | Sort-Object Name)
    if (-not $migrations.Count) { throw "No configured migration found in $migrationsRoot" }

    $selectedMigration = $null
    if ($MigrationName) {
        $selectedMigration = @($migrations | Where-Object { $_.Name -ieq $MigrationName } | Select-Object -First 1)
        if (-not $selectedMigration.Count) { throw "Migration is not in the configured list: $MigrationName" }
        $selectedMigration = $selectedMigration[0]
    }
    else {
        Write-LockLauncherMessage 'Available migrations:'
        for ($index = 0; $index -lt $migrations.Count; $index++) {
            Write-LockLauncherMessage ('{0,2}. {1}' -f ($index + 1), $migrations[$index].Name)
        }
        Write-LockLauncherMessage ' 0. Cancel'
        $choice = Read-LockLauncherChoice -Prompt 'Select a migration number' -Maximum $migrations.Count
        if ($choice -eq 0) { Write-LockLauncherMessage 'Cancelled; no site change was made.'; return }
        $selectedMigration = $migrations[$choice - 1]
    }
    $script:LockLauncherContext.Migration = $selectedMigration.Name
    $config = Import-PowerShellDataFile -LiteralPath (Join-Path $selectedMigration.FullName 'migration.config.psd1')
    if ([string]$config.Name -eq 'NewMigration') { throw 'The migration template cannot be executed.' }

    $mappingPath = Join-Path $selectedMigration.FullName 'migration.mapping.txt'
    $sites = @()
    foreach ($line in @(Get-Content -LiteralPath $mappingPath -ErrorAction Stop)) {
        $trimmed = $line.Trim()
        if (-not $trimmed -or $trimmed.StartsWith('#')) { continue }
        $fields = @($trimmed -split '[\s;,]+' | Where-Object { $_ })
        if ($fields.Count -ne 2) { throw "Each active mapping line must contain one source and one destination URL: $mappingPath" }
        $sites += [pscustomobject]@{ Side = 'Source'; Url = (Get-MappingSiteUrl -Value $fields[0] -MappingPath $mappingPath) }
        $sites += [pscustomobject]@{ Side = 'Destination'; Url = (Get-MappingSiteUrl -Value $fields[1] -MappingPath $mappingPath) }
    }
    if (-not $sites.Count) { throw "No site mapping found: $mappingPath" }

    Write-LockLauncherMessage ('Sites for migration {0}:' -f $selectedMigration.Name)
    for ($index = 0; $index -lt $sites.Count; $index++) {
        $entry = $sites[$index]
        $type = if ($entry.Side -eq 'Source') { [string]$config.Source.Type } else { [string]$config.Target.Type }
        Write-LockLauncherMessage ('{0,2}. {1} ({2}): {3}' -f ($index + 1), $entry.Side, $type, $entry.Url)
    }
    Write-LockLauncherMessage ' 0. Cancel'
    $selectedSite = $null
    if ($Side) {
        $matches = @($sites | Where-Object { $_.Side -eq $Side })
        if ($matches.Count -ne 1) { throw "-Side $Side requires exactly one mapped site on that side." }
        $selectedSite = $matches[0]
    }
    else {
        $choice = Read-LockLauncherChoice -Prompt 'Select a site number' -Maximum $sites.Count
        if ($choice -eq 0) { Write-LockLauncherMessage 'Cancelled; no site change was made.'; return }
        $selectedSite = $sites[$choice - 1]
    }

    $modes = @('ReadOnly', 'Unlock', 'NoAccess')
    Write-LockLauncherMessage 'Lock states:'
    for ($index = 0; $index -lt $modes.Count; $index++) {
        Write-LockLauncherMessage ('{0,2}. {1}' -f ($index + 1), $modes[$index])
    }
    Write-LockLauncherMessage ' 0. Cancel'
    if (-not $LockState) {
        $choice = Read-LockLauncherChoice -Prompt 'Select a lock state number' -Maximum $modes.Count
        if ($choice -eq 0) { Write-LockLauncherMessage 'Cancelled; no site change was made.'; return }
        $LockState = $modes[$choice - 1]
    }

    $section = if ($selectedSite.Side -eq 'Source') { $config.Source } else { $config.Target }
    $type = [string]$section.Type
    if ($type -notin @('SP2016', 'SP2019', 'SPO')) { throw "Unsupported SharePoint type: $type" }
    $operationName = if ($type -eq 'SPO') { 'SmartM365-SharePointTarget-SiteLockStateSet.ps1' } else { 'SmartM365-SharePointSource-SiteLockStateSet.ps1' }
    $operationPath = Join-Path $root (Join-Path 'Scripts\Operations' $operationName)
    if (-not (Test-Path -LiteralPath $operationPath -PathType Leaf)) { throw "Operation script is missing: $operationPath" }
    $operationArgs = @{ SiteUrl = $selectedSite.Url; LockState = $LockState }
    if ($type -eq 'SPO') {
        if (-not [string]$section.TenantAdminUrl) { throw "$($selectedSite.Side).TenantAdminUrl is required for a SPO site." }
        $operationArgs.TenantAdminUrl = [string]$section.TenantAdminUrl
    }

    Write-LockLauncherMessage 'Selection summary:'
    Write-LockLauncherMessage ('Migration: {0}' -f $selectedMigration.Name)
    Write-LockLauncherMessage ('Side and platform: {0} ({1})' -f $selectedSite.Side, $type)
    Write-LockLauncherMessage ('Selected mapped URL: {0}' -f $selectedSite.Url)
    Write-LockLauncherMessage ('Requested lock state: {0}' -f $LockState)
    if ($type -eq 'SPO') {
        Write-LockLauncherMessage ('SPO admin URL: {0}' -f $operationArgs.TenantAdminUrl)
    }
    else {
        Write-LockLauncherMessage 'Execution host: a server in the selected SharePoint farm is required.'
        Write-LockLauncherMessage 'The operation will reject a URL that is not the exact site collection URL.'
    }
    if ($PreviewOnly) {
        Write-LockLauncherMessage 'PreviewOnly completed; no SharePoint command was run.'
        $script:LockLauncherStatus = 'SUCCESS'
        return
    }
    $confirmation = Read-Host ('{0} Type EXECUTE to apply this change; anything else cancels' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    if ($confirmation -cne 'EXECUTE') {
        Write-LockLauncherMessage 'Cancelled; no site change was made.'
        return
    }
    Write-LockLauncherMessage ('Starting {0} for {1}.' -f $operationName, $selectedSite.Url)
    & $operationPath @operationArgs
    if (-not $?) { throw 'Site lock operation did not complete successfully.' }
    $script:LockLauncherStatus = 'SUCCESS'
}
catch {
    $script:LockLauncherFailure = $_
    throw
}
finally {
    Complete-SmartM365MigrationConsoleLifecycle -Context $script:LockLauncherContext -Failure $script:LockLauncherFailure -Status $script:LockLauncherStatus
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBFyDoZmAkhxOA
# CwvX+EnMpDxj1HYemDn/Q/EWLMw8VaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDVjQiwUn3zuNwFtcz4cwmzgWc2xO9azxZ9nDgBZvigUMA0GCSqG
# SIb3DQEBAQUABIIBgHYQYt21yrCz/lxLHm76TBQS9LVQNUBIW4lj7841WElmjQ0M
# ddtpADFvwd+I4BLY5yA0z+eqkp2Qbodgflr5AK1TDkUz53NRbJxSqfaYpyOk4y1g
# INKbBpM0AjCnsJnfosOH944be+/KYlSq2kp++yuRwTBpXGbat6HPqnQqShDReU9j
# uhvIFDJU7pmwjsXRoFzed1E2L69gqQIM5gWwW1QU/tur2gi0AOU8qJzU+BFBUK/O
# dsan+SCvX/cpWJ55ApYluikzkH7XwCJWWvnvQKFel1Q+D+e86Y9YPdDND60kXx4I
# 0WJ0ImwatKSzIZxoxfexPmzgWurQRJPg91+6JE9isQpczVy3A6fnclHHURSb4RUf
# 0kQLGuX6+0O7E/m7V1de1XUzMw9b2M8XphWrLSJpI9+Q0ndu3vnG/waej0Poc9Sf
# yYige/MEBH6LPgKGiZT1ICePnW4GPK0lZf80uD8E5r752+4BdlrHREsDJYfvVyoQ
# 8aP0tcBHlMpp49er/qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTEwODU4
# NDZaMC8GCSqGSIb3DQEJBDEiBCCV9agABrYHmyuXdNfoBXHvlRmwxvevdjAOe66r
# ZPZjXzANBgkqhkiG9w0BAQEFAASCAgAIyNTwY96Q3D9s4MqplBQhi8WndB234oAU
# +SqHrk9Pwj/BgD/NUOmYcOi4wb6i02l5PVbbyOu3LGUiWSVeVJZuay+LVxNIprS3
# NlFOoRfUnYhJWZBVwiDcEFlyWrTe4y70vJRNI4GzUHBEzr58521BQYD+MraD1Uw4
# rsrybbqCoSacCsxXyDShoGOCdgkvO1vrP54zTV7DT1q7ksSM/CS3rtGxL4cnqaBU
# aVz0+FfpvSWduRJM9+41C1J8m2x0OMQ3lnpqbv9qS6XxeEcqCU5Qn2dxSB6Clu0B
# 50yk63IiLZU5WPxlXERmDWNr2FsIsqjzJz5NplA8tDteUGlzGhIp3HU4jCSxJmOx
# biL1xdesZvgNyzog1pSYhUnKF3fbabIDS+V35RONpvesafRFQ/wHyDJJa5ywLjXP
# 1xoxjEjIlPqg+FDGkfB7iaYLoiwLtjwt9nfqhQMPa4e/itFHE2Ti/9WkLh4kcQ1W
# zSE+VRzEdtl/e8DuJEFf+vA6s5yYqv7LM9TOwZVeS9KWTBkolBWosVe8pDa52ctn
# GGE8f183AcWYYu5M6SdN6EmPGFTWlQenCLAsO2p4DSF7JFhnSrjYpB97QXjAZ3Ce
# VLfI0aPJzwtQZfiMmz7ycF3z3J6A4RhdqRbnzMz4Qq/0miRFFmC1sHodXpLZ7biU
# axqRkO5z8A==
# SIG # End signature block
