#Requires -Version 7.0
<#
.SYNOPSIS
    Creates the SmartM365 Device System detection-only Intune Remediation.

.DESCRIPTION
    Uses delegated interactive Microsoft Graph authentication and the selected
    SmartM365 tenant profile. Creates the package without any assignment or run
    schedule. A second run refuses to overwrite an existing package.

.PARAMETER Tenant
    SmartM365 tenant profile key. Defaults to test; use prod explicitly for production.

.PARAMETER DisplayName
    Name of the new Intune Remediation package.

.PARAMETER ConfigRoot
    SmartM365 root containing the local tenant profile. Useful when running
    this script from an isolated Git worktree on the same machine.

.PARAMETER EnforceSignatureCheck
    Require trusted Authenticode signatures on assigned devices. The bundled
    detection file is signed, but the signer must also be trusted by each device.

.NOTES
    Version: 1.0.0
    Delegated Graph permission: DeviceManagementScripts.ReadWrite.All.
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [ValidateNotNullOrEmpty()][string]$DisplayName = 'SmartM365 - Device System Inventory',
    [string]$ConfigRoot = '',
    [switch]$EnforceSignatureCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$startedAt = Get-Date
$completed = $false
$graphConnected = $false
$graphRoot = 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts'
$detectionPath = Join-Path $PSScriptRoot 'SmartM365-Detect-DeviceSystemInfo.ps1'

function Write-Step {
    param([Parameter(Mandatory)][string]$Message)
    Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Get-GraphProperty {
    param([AllowNull()]$InputObject, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) { return ,$InputObject[$Name] }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return ,$property.Value }
    return $null
}

function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    while ($Uri) {
        if (-not $visited.Add($Uri)) { throw 'Graph returned a repeated pagination link.' }
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        $value = Get-GraphProperty -InputObject $page -Name 'value'
        if ($null -eq $value) { throw 'Graph returned a collection without a value property.' }
        foreach ($item in @($value)) { if ($null -ne $item) { $items.Add($item) } }
        $Uri = [string](Get-GraphProperty -InputObject $page -Name '@odata.nextLink')
    }
    return $items.ToArray()
}

function Get-DeviceSystemPackageBody {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$Name,
        [bool]$CheckSignature
    )
    if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) { throw "Detection script not found: $ScriptPath" }
    $signature = Get-AuthenticodeSignature -FilePath $ScriptPath -ErrorAction Stop
    if ($signature.Status -ne 'Valid') { throw "Detection script signature is not valid: $($signature.Status)" }
    $bytes = [IO.File]::ReadAllBytes($ScriptPath)
    $source = [Text.Encoding]::UTF8.GetString($bytes)
    $versionMatch = [regex]::Match($source, '(?m)^\s*Version:\s*([0-9]+(?:\.[0-9]+)+)\s*$')
    if (-not $versionMatch.Success) { throw 'Detection script version was not found.' }
    if ($source -notmatch 'DeviceSystemPayloadEnd:1\.5') { throw 'Detection script does not contain the expected recurring payload marker.' }
    return [ordered]@{
        '@odata.type'            = '#microsoft.graph.deviceHealthScript'
        publisher                = 'SmartM365 by WorkplaceCloudHub'
        version                  = $versionMatch.Groups[1].Value
        displayName              = $Name
        description              = 'Recurring local Device System inventory. Detection only; created without assignments.'
        detectionScriptContent   = [Convert]::ToBase64String($bytes)
        remediationScriptContent = ''
        runAsAccount             = 'system'
        enforceSignatureCheck    = $CheckSignature
        runAs32Bit               = $false
        roleScopeTagIds          = @('0')
        deviceHealthScriptType   = 'deviceHealthScript'
    }
}

function Invoke-CreateDeviceSystemPackage {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Body,
        [Parameter(Mandatory)][string]$GraphRoot
    )
    $escapedName = ([string]$Body.displayName).Replace("'", "''")
    $filter = [uri]::EscapeDataString("displayName eq '$escapedName'")
    $existing = @(Get-GraphCollection -Uri "${GraphRoot}?`$filter=$filter")
    if ($existing.Count -gt 0) {
        $ids = @($existing | ForEach-Object { [string](Get-GraphProperty -InputObject $_ -Name 'id') }) -join ', '
        throw "A Device System Remediation named '$($Body.displayName)' already exists ($ids). No package was changed."
    }

    Write-Step "Creating unassigned detection-only package '$($Body.displayName)'."
    $created = Invoke-MgGraphRequest -Method POST -Uri $GraphRoot -Body ($Body | ConvertTo-Json -Depth 10 -Compress) -ContentType 'application/json' -ErrorAction Stop
    $createdId = [string](Get-GraphProperty -InputObject $created -Name 'id')
    if ($createdId -notmatch '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$') {
        throw 'Graph did not return a valid package ID. Check Intune before retrying to avoid a duplicate.'
    }
    try {
        $assignments = @(Get-GraphCollection -Uri "$GraphRoot/$createdId/assignments")
    }
    catch {
        throw "Package $createdId was created, but assignments could not be verified: $($_.Exception.Message). Check Intune before retrying."
    }
    if ($assignments.Count -ne 0) {
        throw "Package $createdId was created, but Graph reported $($assignments.Count) assignment(s). Review Intune immediately."
    }
    Write-Step "Created package ID: $createdId"
    Write-Step 'Verified assignments: 0. No device or user was assigned.'
    return $createdId
}

try {
    $configRootPath = if ($ConfigRoot) { [IO.Path]::GetFullPath($ConfigRoot) } else {
        [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\..'))
    }
    $tenantContextPath = Join-Path $configRootPath 'Config\SmartM365-TenantContext.ps1'
    if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "Tenant context not found: $tenantContextPath" }
    . $tenantContextPath
    $tenantContext = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $configRootPath
    $tenantId = [string]$tenantContext.TenantId
    if ($tenantId -notmatch '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$') {
        throw "Tenant profile '$Tenant' has no valid TenantId."
    }

    $body = Get-DeviceSystemPackageBody -ScriptPath $detectionPath -Name $DisplayName.Trim() -CheckSignature ([bool]$EnforceSignatureCheck)
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if (Get-MgContext) { Disconnect-MgGraph -ErrorAction Stop | Out-Null }
    Write-Step "Connecting interactively to tenant profile '$Tenant' ($tenantId)."
    Connect-MgGraph -TenantId $tenantId -Scopes @('DeviceManagementScripts.ReadWrite.All') -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
    $graphConnected = $true
    $context = Get-MgContext
    if (-not $context -or [string]$context.TenantId -ine $tenantId -or [string]$context.AuthType -ine 'Delegated' -or
        @($context.Scopes | Where-Object { $_ -ieq 'DeviceManagementScripts.ReadWrite.All' }).Count -eq 0) {
        throw 'The interactive Graph context does not match the configured tenant and delegated permission.'
    }
    Write-Step "Authenticated as $($context.Account)."

    $createdId = Invoke-CreateDeviceSystemPackage -Body $body -GraphRoot $graphRoot
    $completed = $true
}
finally {
    if ($graphConnected) {
        try { Disconnect-MgGraph -ErrorAction Stop | Out-Null }
        catch { Write-Step "Graph disconnect failed: $($_.Exception.Message)" }
    }
    if (Get-Command Write-SmartM365CompletionBanner -ErrorAction SilentlyContinue) {
        Write-SmartM365CompletionBanner -Status $(if ($completed) { 'Success' } else { 'Failed' }) -ScriptName 'SmartM365-Create-DeviceSystemRemediation' -StartedAt $startedAt -EndedAt (Get-Date)
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgdij4YS1Fj/V/
# y2UxvWUhUGY+uc7tFJf7lLE0dGdzc6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICLGbkw9zelfL01MZKgL4jk53ILaISKqheHOytP1lA9uMA0GCSqG
# SIb3DQEBAQUABIIBgIUzEG7REluBiAgHmYgkdsSI3RcGkDN0+W9fohsfW7ewij/R
# VfsuRACj3sAGzlGkvWVjUO0gqwVmmK4xrE3CW6uiyfHPKk8j51HtGIf9Xc8GVQfz
# l3xL89YvzihZJ+GUK7YixrLZQduTAdU2pJdmqIGbFS6STJVcwvP52dw41yvikguw
# XiTKZISI8B03ChSSHrm1rX9VGxH49urleUCZ2NRf/lzwO0zGfXRkC7T+n68/uUhq
# 5IrfbK9LouFSmHLB1qq+UxsSt5NP6dukUzvfMU67QQgo7fkfZtjeJCdBtH3wFy+G
# njhJ1NWJvKEm8NFV+8u9zZMNM196i9YKqfQHYcOwlM+NLXtROrdBxhPWanbuIz9V
# ShdZwwor5FZDEGzD+EgZmeGaS4vIutA7irFW2JR7PVaFBVBKUhJzNwffDQEMqvCu
# nlcvWjyhiIck/YeEVdvAFsuOaoWvy8rgE+ifpRKPwotiB7I4b12bR9sgCzqXxLtM
# hZ1FRF2eVSkrGZSRdaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTMxMzIz
# MzZaMC8GCSqGSIb3DQEJBDEiBCChCopPCuKx9GwTGibhNY+8L/aX/ZnquSExxlor
# ynHrGjANBgkqhkiG9w0BAQEFAASCAgCDI/eJiIU92OMJTgz+OjSMynXP8lDsMiTy
# DHjchbwfuW1qavbMUpKH9w4gAahUHNT6Jdc7i6z05l3nZByIB0hQ9OszfthXRKxf
# OcW4nP4HO9te17YlCwIQm7tqSknThENDYMNSSe1wPPx9y8IonmfqkIeAcYClXjV5
# 9pr/5+XTA82YfBtQ8a1d+dws2B/4U3IxzV8GKv/iglStHLqJe2GWCblgoBRtai5H
# MCcjZFUCdRhoTCwF8fc+Xa5ehagBN7A9jqQb9POlgAVZW032F+BYGOvDFPNg40eu
# SuAzBqCn8+4+YClKdKNIBfCGosS4YQcBXEvl+iSFfWEHaeFSOmI7jhb9A6LPvH2d
# WWz6T3c/89DmYctKFFyjuMyasAeqYFnYnhNxudVoebj443ria9MHByh710IPe1xl
# 4jb5KgR8HdljFaKzzMNE1GTpbz850GXgRvO9c9aoCcJcBMaDmQYKfYEGWSTjLRgJ
# Bn5BeupuqvM0BsxquNYim9G0GVRelW36vrfOoxtW8H4khae5tnLArGHrECOECvaT
# 4zf+dhK2GUXCNa+rlua+53yMaPupiYUQDpA1KkMF+Tq2SruiCBypw5N+Gko9BbX/
# L7JE4QsVuvmKUDhfKDTIEv0J1vpPVDX2FYMgh4Uwo7TyBpw8NNBP1AXVY/wLIiQS
# PQS/12DXSQ==
# SIG # End signature block
