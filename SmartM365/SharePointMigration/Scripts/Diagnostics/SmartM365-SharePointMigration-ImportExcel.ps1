<#
.SYNOPSIS
    Prepare ImportExcel for XLSX report analysis in the current user context.
.VERSION
    1.0.0
.DESCRIPTION
    Function-only helper. Installs from the official PSGallery when missing.
    DryRun never installs modules or registers repositories.
#>

function Initialize-SmartM365ImportExcel {
    param([switch]$DryRun, [scriptblock]$Progress)
    $module = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
    if ($DryRun) { return $module }
    try {
        if (-not $module) {
            if ($Progress) { & $Progress 'InstallingImportExcel' 'Installing ImportExcel from PSGallery for the current user…' | Out-Null }
            $resourceInstaller = Get-Command -Name Install-PSResource -ErrorAction SilentlyContinue
            if ($resourceInstaller) {
                $repository = Get-PSResourceRepository -Name PSGallery -ErrorAction SilentlyContinue
                if (-not $repository) {
                    Register-PSResourceRepository -PSGallery -ErrorAction Stop
                    $repository = Get-PSResourceRepository -Name PSGallery -ErrorAction Stop
                }
                if (([string]$repository.Uri).TrimEnd('/') -ne 'https://www.powershellgallery.com/api/v2') {
                    throw 'PSGallery must point to the official https://www.powershellgallery.com/api/v2 repository.'
                }
                $options = @{ Name='ImportExcel'; Scope='CurrentUser'; Repository='PSGallery'; TrustRepository=$true; AcceptLicense=$true; ErrorAction='Stop' }
                if ($resourceInstaller.Parameters.ContainsKey('Quiet')) { $options.Quiet=$true }
                Install-PSResource @options | Out-Null
            }
            else {
                $installer = Get-Command -Name Install-Module -ErrorAction SilentlyContinue
                if (-not $installer) { throw 'Install-PSResource or Install-Module is required to install ImportExcel automatically.' }
                $repository = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
                if (-not $repository) {
                    Register-PSRepository -Default -ErrorAction Stop
                    $repository = Get-PSRepository -Name PSGallery -ErrorAction Stop
                }
                if (([string]$repository.SourceLocation).TrimEnd('/') -ne 'https://www.powershellgallery.com/api/v2') {
                    throw 'PSGallery must point to the official https://www.powershellgallery.com/api/v2 repository.'
                }
                $options = @{ Name='ImportExcel'; Scope='CurrentUser'; Repository='PSGallery'; Force=$true; Confirm=$false; ErrorAction='Stop' }
                if ($installer.Parameters.ContainsKey('AcceptLicense')) { $options.AcceptLicense=$true }
                Install-Module @options | Out-Null
            }
            $module = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
            if (-not $module) { throw 'Installation returned without a discoverable ImportExcel module.' }
        }
        Import-Module -Name $module.Path -Global -ErrorAction Stop
        foreach ($name in @('Import-Excel','Get-ExcelSheetInfo')) {
            if (-not (Get-Command -Name $name -Module ImportExcel -ErrorAction SilentlyContinue)) { throw "ImportExcel did not provide the required command: $name" }
        }
        if ($Progress) { & $Progress 'ImportExcelReady' ("ImportExcel v{0} is ready. Preparing the XLSX report…" -f $module.Version) | Out-Null }
        return $module
    }
    catch {
        throw "ImportExcel preparation failed for the current user: $($_.Exception.Message) Check access to PSGallery and the current user's PowerShell module folder, then retry Analyze latest report."
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA5eNiV/EMEhsyC
# jYMqUsUl2ohXYIwUEhSBCNBfVFRH/qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCQBxgEVqnbE9K+3ee3Pd5D
# jOVRrysW7Xq9kIf0LqRXOjANBgkqhkiG9w0BAQEFAASCAYAloMCf+Qtg81SlfXe+
# 8MzsW6qps23C4kH4tsbsK8BNLV2VPyV/nufiAfuhbyYgQ7Ih5Z0GSa8bIdCa+zi0
# 6wl3T+1afns+F2TWHfeeIhvQ2+T3RCXVcqrAM1qoH8pYbzoJFQUTb7eo3jUhlMMf
# 5gSUaQJRtcpOdDp3SI6vAZ06C25hhkgP9XUuzDmli+bSnQxAOTCCiGN2Cm9lZ3q6
# enp2K7W7zhhcNgYJlJOWQfMCroWqLatmyferbPq6Go0gHiefk02KQSi7sevG3OnP
# //bNnjldAopjoTWVUjzzygrMfIl68Rar/AcCriI196b6ikYkrMGmTptTbLtuZ+XE
# W/eHyRhfRlhEH5Ik1twMX3/LM64QHMvHzuR/0fkzRVh5CraKwEHcEk2wuAuejSgX
# Y1+7+aHapyjoBoq4wz2LIMOaePMYm30vbv2g2mrr/nTnljKDBLA5SxnQUyAopfOr
# 7jv4+F478cq91wn+h0UzMNnIGsDbr5FMzbSRpmPdt5pukDE=
# SIG # End signature block
