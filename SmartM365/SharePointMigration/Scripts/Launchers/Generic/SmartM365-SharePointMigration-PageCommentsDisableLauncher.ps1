<#
.SYNOPSIS
    Selects a migration destination before disabling SharePoint page comments.
.DESCRIPTION
    Shows source and destination sites from the migration mapping. Only the
    SharePoint Online destination can be selected for this target operation.
.VERSION
    1.0.1
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ToolkitRoot = '',
    [switch]$PreviewOnly
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:PageCommentsContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath -Action 'Disable SharePoint Online page comments'
$script:PageCommentsFailure = $null
$script:PageCommentsStatus = 'CANCELLED'

function Write-PageCommentsMessage {
    param([string]$Message)
    Microsoft.PowerShell.Utility\Write-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Read-PageCommentsChoice {
    param([string]$Prompt, [int]$Maximum)
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $answer = Read-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Prompt)
        $number = -1
        if ([int]::TryParse([string]$answer, [ref]$number) -and $number -ge 0 -and $number -le $Maximum) {
            return $number
        }
        Write-PageCommentsMessage ('Invalid selection ({0}/3).' -f $attempt)
    }
    throw 'No valid selection was entered.'
}

function ConvertTo-MappingUrl {
    param([string]$Value, [string]$MappingPath)
    $uri = $null
    if (-not [uri]::TryCreate($Value, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -notin @('http', 'https') -or $uri.UserInfo -or $uri.Query -or $uri.Fragment) {
        throw "Invalid site URL in mapping: $MappingPath"
    }
    return $uri.AbsoluteUri.TrimEnd('/')
}

try {
    if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1) {
        throw 'Windows PowerShell 5.1 is required for the page comments launcher.'
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

    Write-PageCommentsMessage 'Available migrations:'
    for ($index = 0; $index -lt $migrations.Count; $index++) {
        Write-PageCommentsMessage ('{0,2}. {1}' -f ($index + 1), $migrations[$index].Name)
    }
    Write-PageCommentsMessage ' 0. Cancel'
    $choice = Read-PageCommentsChoice -Prompt 'Select a migration number' -Maximum $migrations.Count
    if ($choice -eq 0) {
        Write-PageCommentsMessage 'Cancelled; no site change was made.'
        return
    }
    $migration = $migrations[$choice - 1]
    $script:PageCommentsContext.Migration = $migration.Name

    $config = Import-PowerShellDataFile -LiteralPath (Join-Path $migration.FullName 'migration.config.psd1')
    if ([string]$config.Name -eq 'NewMigration') { throw 'The migration template cannot be executed.' }
    if ([string]$config.Target.Type -ine 'SPO') { throw 'Disabling page comments requires a SharePoint Online destination.' }
    if ([string]::IsNullOrWhiteSpace([string]$config.Target.SiteUrl) -or
        [string]::IsNullOrWhiteSpace([string]$config.Target.TenantAdminUrl)) {
        throw 'Target.SiteUrl and Target.TenantAdminUrl are required in migration.config.psd1.'
    }

    $mappingPath = Join-Path $migration.FullName 'migration.mapping.txt'
    $sites = @()
    foreach ($line in @(Get-Content -LiteralPath $mappingPath -ErrorAction Stop)) {
        $trimmed = $line.Trim()
        if (-not $trimmed -or $trimmed.StartsWith('#')) { continue }
        $fields = @($trimmed -split '[\s;,]+' | Where-Object { $_ })
        if ($fields.Count -ne 2) {
            throw "Each active mapping line must contain one source and one destination URL: $mappingPath"
        }
        $sites += [pscustomobject]@{
            Source = ConvertTo-MappingUrl -Value $fields[0] -MappingPath $mappingPath
            Destination = ConvertTo-MappingUrl -Value $fields[1] -MappingPath $mappingPath
        }
    }
    if (-not $sites.Count) { throw "No site mapping found: $mappingPath" }

    $configuredDestination = ConvertTo-MappingUrl -Value ([string]$config.Target.SiteUrl) -MappingPath (Join-Path $migration.FullName 'migration.config.psd1')
    $destinations = @($sites | ForEach-Object { $_.Destination } | Select-Object -Unique)
    if ($destinations.Count -ne 1 -or $destinations[0] -ine $configuredDestination) {
        throw 'The mapped destination must match Target.SiteUrl before this operation can run.'
    }

    Write-PageCommentsMessage ('Sites for migration {0}:' -f $migration.Name)
    foreach ($site in $sites) {
        Write-PageCommentsMessage ('    Source ({0}, unavailable for this operation): {1}' -f $config.Source.Type, $site.Source)
    }
    Write-PageCommentsMessage (' 1. Destination (SPO): {0}' -f $destinations[0])
    Write-PageCommentsMessage ' 0. Cancel'
    $choice = Read-PageCommentsChoice -Prompt 'Select a destination number' -Maximum 1
    if ($choice -eq 0) {
        Write-PageCommentsMessage 'Cancelled; no site change was made.'
        return
    }

    Write-PageCommentsMessage 'Selection summary:'
    Write-PageCommentsMessage ('Migration: {0}' -f $migration.Name)
    Write-PageCommentsMessage ('Destination site: {0}' -f $destinations[0])
    Write-PageCommentsMessage ('SPO admin URL: {0}' -f $config.Target.TenantAdminUrl)
    Write-PageCommentsMessage 'Action: disable page comments on the destination site collection and its subsites.'

    if ($PreviewOnly) {
        Write-PageCommentsMessage 'PreviewOnly completed; no SharePoint command was run.'
        $script:PageCommentsStatus = 'SUCCESS'
        return
    }

    $confirmation = Read-Host ('{0} Type EXECUTE to apply this change; anything else cancels' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    if ($confirmation -cne 'EXECUTE') {
        Write-PageCommentsMessage 'Cancelled; no site change was made.'
        return
    }

    $operationLauncher = Join-Path $root 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-OperationLauncher.ps1'
    if (-not (Test-Path -LiteralPath $operationLauncher -PathType Leaf)) {
        throw "Operation launcher is missing: $operationLauncher"
    }
    Write-PageCommentsMessage ('Starting page comments operation for {0}.' -f $destinations[0])
    & $operationLauncher -MigrationRoot $migration.FullName -Operation DisablePageComments -Execute -NonInteractive
    if (-not $?) { throw 'Page comments operation did not complete successfully.' }
    $script:PageCommentsStatus = 'SUCCESS'
}
catch {
    $script:PageCommentsFailure = $_
    throw
}
finally {
    Complete-SmartM365MigrationConsoleLifecycle -Context $script:PageCommentsContext -Failure $script:PageCommentsFailure -Status $script:PageCommentsStatus
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB6CFC7J+JAEJgS
# vG/pheWZXY8hUd+Todpz3mxEtpcVoqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICeGUkMn5VwNvreeqEoQHOaRNccHKLNb9mT8MR+PAYi4MA0GCSqG
# SIb3DQEBAQUABIIBgJy9V+tjQXFdZUmfvbyKSRe1tviZhJba/11xFREXn2i+AjNC
# qEv/cixt628Joe4gEGIC6j5zw24PLREdU1FR1EOIvq22L5fHT3aTVk7qiDGXsCLg
# RMHU7be1w0e4s6be5yNFBrpUkh+qqArvXiIQUgn5VkTtuk5Le/SVVR5IwAei4jkJ
# 2AZATFodhCGHnF1ZasohCigWc0HyDEN2mNEv0VBGg02BHyZydr93pWHKSQUkRDZ/
# fK8Abtus8sVZkG2ovNNkW6hgUvHfib0nkvs9lFLFhIh/ItIxfBep8h/wG+vq4/Jj
# E3e0Rf1yC8uF4VIyQcUbCRspyb02Bxv4Wymtk/Hs22NHpOC7jcJ9vDjuUZaypo1V
# jZt0TdH7co3dK282N+GyZFnJO6NmiLW5mS5VvJmQOnTwAq8p/1yFg456kxCCDzFv
# tWBoPwFtmJdWvkTsYLrCaJG/i4GY4Z1mKdLI3LdMf7YMLq1R8zKgKj05EC8YwnjJ
# M+LiibH0zeR+91Alt6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTExMjEy
# NTVaMC8GCSqGSIb3DQEJBDEiBCA6VB/erqnOcjoGor2T2va8lBk4l6+49CuhPgmR
# uY/GWzANBgkqhkiG9w0BAQEFAASCAgCv0iJT/CQH25yRij/86Q9mdB4U/mJZgcyp
# 4S2N3Pfrd/dJ39HU1j8/iCmulGGQ5zoCTN2jaOscnPmMmGyf7OodIzTv+9s3l/2i
# hOlyhPQSC+hirqy7+unjB4mRA2fbw0V+MUTzdTXoKIlsOXqjNwi5V683p85+h4hx
# TubbtGYeDKu3KqKpf2VD0M/ZrRtZv21DXIWvAHwVIF2aLqOeA9cGuenIxvJctHHu
# XOYD+rXuvfaaTssg7WHCV1R2dpZ0J/8oEoz/Uxo4rvcQasp7CYMpS6J8CVlz/SS/
# rqDRh5w5FkFYNFMh8u02NGjNwbwJbwnOdtMD8BjSS23DdCQv3b74b8Y2aS0yHbI0
# 3cH0qENTcx4gzj6cblLyNEUSPs8YUG/HG7kzMxNLDZ6sXpn1yt0MpuPWaOLlW+oX
# wwOdxRfrvN1all83gK2LwIUi0XgrQiilwa9s+p3iBNm4YXkVDsmEOusTrAiMVlMf
# 8rGYTV+B73/e9y9AMwIOHr86KEJoxl1mrWBOF5G/V9c9CHNcMJsmbV4uya1YHAIx
# uV7JfLL/LtJgB3J4LURJ2ehd0dAx9b9XRFqmFqLcpmArdgKlcPlC5gtUYwvNqE9b
# HqY/cIBvy4KBB3tYly3+hQtQG38TbqSxz8itmA8ClnasMZMWjz1V2/vVv1PCqMHv
# IGG4glwb8A==
# SIG # End signature block
