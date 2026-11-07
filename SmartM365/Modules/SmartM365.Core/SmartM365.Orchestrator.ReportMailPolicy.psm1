<#
.SYNOPSIS
    Testable admission policy for report mail. No production caller is wired to it.
#>
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.ReportMailSchedule.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.ReportMailQueue.psm1') -Force -ErrorAction Stop

function Submit-SmartM365ReportMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Report','Operational','Error','Maintenance')][string]$DeliveryClass,
        [Parameter(Mandatory)][bool]$Enabled,
        [Parameter(Mandatory)][datetimeoffset]$Now,
        [Parameter(Mandatory)][hashtable]$Message,
        [Parameter(Mandatory)][scriptblock]$SendAction,
        [string]$NotBefore = '07:30',
        [string]$TimeZoneId = 'Romance Standard Time',
        [string]$QueuePath,
        [scriptblock]$AssertTrustedDirectory,
        [string]$TenantKey,
        [string]$DeduplicationKey
    )
    if ($Enabled -and $DeliveryClass -eq 'Report') {
        $eligible = Get-SmartM365ReportMailNotBeforeUtc -Now $Now -NotBefore $NotBefore -TimeZoneId $TimeZoneId
        if ($eligible -gt $Now.ToUniversalTime()) {
            foreach ($required in @('QueuePath','TenantKey','DeduplicationKey')) {
                if ([string]::IsNullOrWhiteSpace([string](Get-Variable -Name $required -ValueOnly))) { throw "Queue $required is required." }
            }
            if ($null -eq $AssertTrustedDirectory) { throw 'A trusted-directory verifier is required.' }
            $queued = Add-SmartM365ReportMailQueueItem -QueuePath $QueuePath -AssertTrustedDirectory $AssertTrustedDirectory -TenantKey $TenantKey -DeduplicationKey $DeduplicationKey -NotBeforeUtc $eligible -Message $Message
            return [pscustomobject]@{ Status = if ($queued.Status -eq 'Sent') { 'Sent' } else { 'Queued' }; Id = $queued.Id }
        }
    }
    $result = @(& $SendAction $Message)
    if ($result.Count -ne 1 -or $result[0] -isnot [bool] -or -not $result[0]) { throw 'Transport did not explicitly confirm acceptance.' }
    return [pscustomobject]@{ Status = 'Sent'; Id = '' }
}

Export-ModuleMember -Function Submit-SmartM365ReportMail

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAvE35N5P3ox51g
# so0tbTMYMK/FcJ5aFYbliTfGJJltTaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBRDvP+W5CurkyWj4BB8r4+
# DEIzLdEOK/47LHQl68ZMXTANBgkqhkiG9w0BAQEFAASCAYCBD8TDoWULO7IZeeZ3
# IUA8ooEPsvcpkZcn13xXtmJM0v7rEaQSVlALOy030Xsp190PHmSfwtgHOvZ1PIhN
# doVbUQFGT6Tl+9OP6B3pqqw6OsSeFEmG2PLkRlKCAEa9r1gXRGNNxsree7NiT1AN
# /yMaJmA8TGudLjgflu+ot1DGFw4WfdYDbeI0BDBU1p15ugV73ErwTXc1iVcUpn9i
# FhIhQcD0oeg/+6NzvV9Y6l8ylN/d4b4286GiWodXezbZKblZxRHEVWnIdNMV2+5r
# 10rL+hiFdozzJGc2YryXoFj4yTuOFRLHsCHQt+fW+AwEltd2nojnrwhM/iOii7yp
# OsHdP10Y023msIEDNZc5dbsXXeJ5fgcKphuboly2EOP36UL6uf3sm84L+QR9d4eq
# BxjGTsxC2nq378csYYanVH4Owk6cNFlMjIRvdPXAvMufsq68/cl1Xqb+7yi88ZMv
# Ybl7onM4jwGdAdK19lRrZAXU1+1CkoWvjV0C+s/4hAhEHKs=
# SIG # End signature block
