# SmartWorkplaceCMDB.Graph
# Version: 0.2.0

$script:SmartWorkplaceCMDBGraphVersion = '0.2.0'

function Get-SmartWorkplaceCMDBGraphObjectValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }
    return $null
}

function Test-SmartWorkplaceCMDBGraphAppOnlyReadiness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [string]$CertificateThumbprint
    )

    $parsedTenantId = [guid]::Empty
    if (-not [guid]::TryParse($TenantId, [ref]$parsedTenantId)) {
        throw 'MicrosoftGraph.TenantId must contain an Entra tenant GUID.'
    }
    $parsedClientId = [guid]::Empty
    if (-not [guid]::TryParse($ClientId, [ref]$parsedClientId)) {
        throw 'MicrosoftGraph.ClientId must contain an application registration GUID.'
    }
    if ($CertificateThumbprint -notmatch '^[a-fA-F0-9]{40,64}$') {
        throw 'MicrosoftGraph.CertificateThumbprint must contain a 40 to 64 character hexadecimal certificate thumbprint.'
    }

    $authenticationModule = Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if ($null -eq $authenticationModule) {
        throw 'Microsoft.Graph.Authentication is required for live Microsoft Graph collection.'
    }

    $certificate = Get-ChildItem -Path Cert:\CurrentUser\My |
        Where-Object Thumbprint -eq $CertificateThumbprint |
        Select-Object -First 1
    if ($null -eq $certificate) {
        throw 'The configured Microsoft Graph certificate was not found in Cert:\CurrentUser\My.'
    }
    if (-not $certificate.HasPrivateKey) {
        throw 'The configured Microsoft Graph certificate does not have an accessible private key.'
    }
    if ($certificate.NotAfter -le (Get-Date)) {
        throw 'The configured Microsoft Graph certificate is expired.'
    }

    return [pscustomobject]@{
        TenantId                   = $parsedTenantId.ToString()
        ClientId                   = $parsedClientId.ToString()
        CertificateThumbprint      = $certificate.Thumbprint
        AuthenticationModule       = $authenticationModule.Name
        AuthenticationModuleVersion = $authenticationModule.Version.ToString()
    }
}

function Get-SmartWorkplaceCMDBGraphRetryStatusCode {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    $response = $ErrorRecord.Exception.Response
    if ($null -ne $response -and $null -ne $response.StatusCode) {
        try { return [int]$response.StatusCode }
        catch { }
    }

    $message = [string]$ErrorRecord.Exception.Message
    $match = [regex]::Match($message, '(?<!\d)(408|429|500|502|503|504)(?!\d)')
    if ($match.Success) {
        return [int]$match.Groups[1].Value
    }
    return 0
}

function Get-SmartWorkplaceCMDBGraphRetryDelaySeconds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [Parameter(Mandatory)][ValidateRange(1, 30)][int]$Attempt,
        [ValidateRange(1, 60)][int]$BaseDelaySeconds = 2,
        [ValidateRange(1, 300)][int]$MaximumDelaySeconds = 120
    )

    $retryAfter = $null
    if ($null -ne $ErrorRecord.Exception.Data -and
        $ErrorRecord.Exception.Data.Contains('Retry-After')) {
        $retryAfter = [string]$ErrorRecord.Exception.Data['Retry-After']
    }
    if ([string]::IsNullOrWhiteSpace($retryAfter)) {
        $response = $ErrorRecord.Exception.Response
        if ($null -ne $response -and $null -ne $response.Headers) {
            try { $retryAfter = [string]$response.Headers['Retry-After'] }
            catch { }
            if ([string]::IsNullOrWhiteSpace($retryAfter)) {
                try { $retryAfter = [string]$response.Headers.RetryAfter.Delta.TotalSeconds }
                catch { }
            }
        }
    }

    $seconds = 0
    if (-not [string]::IsNullOrWhiteSpace($retryAfter) -and
        [int]::TryParse($retryAfter.Trim(), [ref]$seconds) -and
        $seconds -gt 0) {
        return [Math]::Min($seconds, $MaximumDelaySeconds)
    }

    $delay = [int]($BaseDelaySeconds * [Math]::Pow(2, $Attempt - 1))
    return [Math]::Min($delay, $MaximumDelaySeconds)
}

function Invoke-SmartWorkplaceCMDBGraphRequestWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateRange(0, 10)][int]$MaximumRetryCount = 5,
        [ValidateRange(1, 60)][int]$BaseDelaySeconds = 2,
        [ValidateRange(1, 300)][int]$MaximumDelaySeconds = 120,
        [scriptblock]$RequestScript = {
            param($RequestUri)
            Invoke-MgGraphRequest -Method GET -Uri $RequestUri -ErrorAction Stop
        },
        [scriptblock]$SleepScript = {
            param($Seconds)
            Start-Sleep -Seconds $Seconds
        }
    )

    $attempt = 0
    while ($true) {
        try {
            return & $RequestScript $Uri
        }
        catch {
            $statusCode = Get-SmartWorkplaceCMDBGraphRetryStatusCode -ErrorRecord $_
            $transient = $statusCode -in @(408, 429, 500, 502, 503, 504)
            if (-not $transient -or $attempt -ge $MaximumRetryCount) {
                throw
            }
            $attempt++
            $delay = Get-SmartWorkplaceCMDBGraphRetryDelaySeconds `
                -ErrorRecord $_ `
                -Attempt $attempt `
                -BaseDelaySeconds $BaseDelaySeconds `
                -MaximumDelaySeconds $MaximumDelaySeconds
            Write-Warning (
                'Microsoft Graph request returned transient status {0}. Retry {1}/{2} in {3} second(s).' -f
                $statusCode, $attempt, $MaximumRetryCount, $delay
            )
            & $SleepScript $delay
        }
    }
}

function Invoke-SmartWorkplaceCMDBGraphPagedRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [string]$CertificateThumbprint,

        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [string]$RequiredPermission,

        [ValidateRange(0, 2147483647)]
        [int]$MaxItems = 0
    )

    $readiness = Test-SmartWorkplaceCMDBGraphAppOnlyReadiness `
        -TenantId $TenantId `
        -ClientId $ClientId `
        -CertificateThumbprint $CertificateThumbprint
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $connected = $false
    try {
        Connect-MgGraph `
            -TenantId $readiness.TenantId `
            -ClientId $readiness.ClientId `
            -CertificateThumbprint $readiness.CertificateThumbprint `
            -ContextScope Process `
            -NoWelcome `
            -ErrorAction Stop | Out-Null
        $connected = $true

        $context = Get-MgContext -ErrorAction Stop
        if ($null -eq $context -or
            [string]::IsNullOrWhiteSpace([string]$context.TenantId) -or
            [string]$context.TenantId -ne $readiness.TenantId) {
            throw 'Microsoft Graph connected to an unexpected tenant.'
        }

        $items = New-Object System.Collections.Generic.List[object]
        $nextUri = $Uri
        while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
            try {
                $response = Invoke-SmartWorkplaceCMDBGraphRequestWithRetry -Uri $nextUri
            }
            catch {
                if ($_.Exception.Message -match '403|Forbidden|Authorization_RequestDenied|Insufficient') {
                    throw (
                        "Microsoft Graph denied the query. Grant and admin-consent the '{0}' application permission." -f
                        $RequiredPermission
                    )
                }
                throw
            }

            Assert-SmartWorkplaceCMDBCollectionPage -Response $response
            foreach ($item in @(Get-SmartWorkplaceCMDBGraphObjectValue -InputObject $response -Name 'value')) {
                if ($null -ne $item) {
                    $items.Add($item)
                    if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) {
                        break
                    }
                }
            }
            if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) {
                break
            }
            $nextUri = [string](Get-SmartWorkplaceCMDBGraphObjectValue -InputObject $response -Name '@odata.nextLink')
        }

        return @($items.ToArray())
    }
    finally {
        if ($connected -and (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

Export-ModuleMember -Function @(
    'Get-SmartWorkplaceCMDBGraphObjectValue',
    'Test-SmartWorkplaceCMDBGraphAppOnlyReadiness',
    'Invoke-SmartWorkplaceCMDBGraphRequestWithRetry',
    'Invoke-SmartWorkplaceCMDBGraphPagedRequest'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB2Hmqy9dF5Ov+2
# Hgm0AOE7aVAku8PFWz/jzk+F2TqtmqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILtxQYdy8FApBinbowA8i1wT0cDXicr5utoWgRQ5owRDMA0GCSqG
# SIb3DQEBAQUABIIBgEPiSTYzLAKsh8YUT4FjdagFuh8YOyFpv813u54or6I2b9kr
# 4sv5LwevW07cURuMhazfw9tka8YLguF2u3NhfA+TbABpsIuuQu3TZ5FA0XzCR91f
# LGibDYCbLIM+s/z++Vxe4l4FHEAdIFweDGMRuXcq7o+dE8lbBF+UHH7RAjr9C7R8
# MX39o9rNxdHCzgEDh+z34Hud1hGl4jWQ+TBVSNSvGhOPiuaaS8IwoAfv9dYGHijY
# GO5stoE8s0MkU7FoeGrwfKxRBooeQdloSernxytGOIaH1PaIbTU3np2AKxHcTcl+
# E9t/YVCwVWWTSjdc+wXA/qSfFBvngpjsE5kJ5MBCw/jZJs7/Oo3wjtSMv59DceTH
# f31bSplL5+jUZxFTfeDsRhubN2SsSBPB4SvNN3LmnmiFbqrB0O+W6xmdcf3WwUss
# xzxNDwi35TaR7eWP9QP4UMBimVT+zyyuYkjP2tN5qx3ZdAGmX0VZgk9OOilV6AOI
# rcpGs42QW25hkcBGJKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTExNTE1
# MTVaMC8GCSqGSIb3DQEJBDEiBCByV2581n6+y6VKV5mtlBAk1DwNNsLpzTbs+PCa
# RLpJOTANBgkqhkiG9w0BAQEFAASCAgBpIJQYMOztCprwh/Q8RA/Nt9I4mjYw9oUE
# kOVY+vW5BgYab9L+AcA16TdZEc+2sjrQztmCi2QUnkJpUmqprHx0r3Qz95c13RQj
# xJRn/GBk/qHF0nnTcowwTCM8f9h5M7FQb8xNoJeJDDxabl/IIKTqOMWyDdS5lu1r
# Tf8vzrrfkknVKj0GO4Aanh1GVoTO3J7WgKXrdZM9MSk6v1rO3RtgqsJgKOPID6HN
# 2cukMaJtczypC4H/RaWyF27vBXhhgKHCNd6BEBvSKqETWc/M5mEVyziy5iKPSYSH
# S5VVHBM7MdEiupJVZ1X5hkqWP4ArtL8VOfkoDK4iR6OYI10CLBQkQv+ByYLxPFkT
# yH4e1H5RR8MrVRvXf2vIicmj9zwaaThkhqtcwX8I1D0HY29jRFSofRYhEUUra1+I
# +7tDlmlzTbTDgffmJhS97JIVJPdYwg69kwTwjypi0uN9ISvS9oSduPWWDc0lqSzS
# X6ovrXJHSWqXLMlUnysW77yuuU9zb5/7ouCEfSUbdJB6rwxJ4UxkvQlQjV9RjrBf
# gjHbX+g9aB9W1kZVj7Z7nVeEiJdMclv4zsWh/5r9/NjZgKODDXZq9y8SOGwrRtJD
# m1+xKKELRUpKdoZCtRP4cxa0aWX59TDOLv01BxfQyYZ2/+y3I9HthmnkaITJQUlG
# +P1n+MzGiw==
# SIG # End signature block
