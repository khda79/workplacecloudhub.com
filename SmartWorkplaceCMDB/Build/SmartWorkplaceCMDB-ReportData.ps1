<#
.SYNOPSIS
Builds the current Power BI report-only tables after a complete CMDB build.

.DESCRIPTION
Keeps the 28 collector contract tables unchanged. The 19 report tables and
their validated CI hardware evidence are replaced as one current directory;
no dated report-data history is retained.
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant = 'default',
    [string]$OrganizationKey, [string]$EnvironmentKey, [string]$TenantKey,
    [string]$TenantId, [string]$DataRootPath, [string]$DataAllRootPath,
    [string]$LatestOutputRootPath, [string]$LogRootPath,
    [string]$GlobalConfigPath, [string]$TenantConfigPath,
    [switch]$NoConfigWrite, [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$projectRoot = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1') -Force
$bound = @{}
foreach ($key in $PSBoundParameters.Keys) { $bound[$key] = $PSBoundParameters[$key] }
$context = Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound `
    -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath `
    -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths = $context.Paths
$reportConfig = $context.Configuration['ReportData']
if ($null -eq $reportConfig -or -not [bool]$reportConfig['Enabled']) {
    throw 'ReportData.Enabled must be true for the configured Full report-data step.'
}
$inventoryRoot = [string]$reportConfig['SmartInventoryLatestOutputRootPath']
if ([string]::IsNullOrWhiteSpace($inventoryRoot)) {
    throw 'ReportData.SmartInventoryLatestOutputRootPath is required.'
}
$inventoryRoot = [IO.Path]::GetFullPath($inventoryRoot)
$localMailboxes = Join-Path $inventoryRoot 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
$remoteMailboxes = Join-Path $inventoryRoot 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv'
$source = [IO.Path]::GetFullPath($paths.LatestOutputRootPath)
$powerBI = Join-Path $source 'PowerBI'
$destination = Join-Path $powerBI 'Report'
$raw = Join-Path $source 'Raw'
$hardware = Join-Path $raw 'Intune\Intune_DeviceHardware.csv'
foreach ($path in @($source, $powerBI, $raw, $hardware, $localMailboxes, $remoteMailboxes)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required report-data input is missing: '$path'." }
}
$python = Get-Command python -CommandType Application -ErrorAction Stop |
    Select-Object -First 1
$versionText = & $python.Source -c 'import sys; print(".".join(map(str, sys.version_info[:3])))'
if ($LASTEXITCODE -ne 0 -or [version]$versionText -lt [version]'3.10') {
    throw 'Python 3.10 or later is required for report-data preparation.'
}
$generator = Join-Path $projectRoot 'PowerBI\prepare_current_report_data.py'
if ($ValidateOnly) {
    [pscustomobject]@{Status='Validated'; PythonVersion=$versionText;
        SmartInventoryRoot=$inventoryRoot; ReportRoot=$destination}
    return
}

$runId = [guid]::NewGuid().ToString('N')
$ciWorkRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartWorkplaceCMDB-CI-' + $runId)
$ciOutput = Join-Path $ciWorkRoot 'CI'
$stage = Join-Path $powerBI ('Report.stage.' + $runId)
$previous = Join-Path $powerBI ('Report.previous.' + $runId)
foreach ($target in @($destination, $stage, $previous)) {
    if ([IO.Path]::GetFullPath((Split-Path $target -Parent)) -ne
        [IO.Path]::GetFullPath($powerBI)) {
        throw "Unsafe report-data target: '$target'."
    }
}
$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
if (-not [IO.Path]::GetFullPath($ciWorkRoot).StartsWith(
        $tempRoot + [IO.Path]::DirectorySeparatorChar,
        [StringComparison]::OrdinalIgnoreCase)) {
    throw "Unsafe CI staging target: '$ciWorkRoot'."
}
$promoted = $false
$movedPrevious = $false
$promotionValidated = $false
try {
    New-Item -ItemType Directory -Path $ciWorkRoot | Out-Null
    Import-Module (Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.CI\SmartWorkplaceCMDB.CI.psm1') -Force
    Export-SmartWorkplaceCMDBCIRegistry -InputRootPath (Join-Path $source 'CMDB') `
        -RawRootPath $raw -HardwareInputPath $hardware -IncludeContext `
        -OrganizationKey $paths.OrganizationKey -EnvironmentKey $paths.EnvironmentKey `
        -TenantKey $paths.TenantKey -TenantId $paths.TenantId `
        -OutputDirectory $ciOutput | Out-Null
    $arguments = @($generator, '--data-root', $source, '--output', $stage,
        '--final-output', $destination,
        '--ci-hardware', (Join-Path $ciOutput 'CMDB_CIDeviceHardware.csv'),
        '--exchange-onprem-local', $localMailboxes,
        '--exchange-onprem-remote', $remoteMailboxes)
    & $python.Source @arguments
    if ($LASTEXITCODE -ne 0) { throw "Report-data generator failed with exit code $LASTEXITCODE." }
    & $python.Source $generator --data-root $source --output $stage --validate-only
    if ($LASTEXITCODE -ne 0) { throw 'Staged report-data validation failed.' }
    if (Test-Path -LiteralPath $destination) {
        $oldManifestPath = Join-Path $destination 'report-data.manifest.json.txt'
        if (-not (Test-Path -LiteralPath $oldManifestPath -PathType Leaf)) {
            throw "Existing Report folder has no CMDB report-data manifest: '$destination'."
        }
        $oldManifest = Get-Content -LiteralPath $oldManifestPath -Raw | ConvertFrom-Json
        if ([string]$oldManifest.identity.TenantKey -cne [string]$paths.TenantKey -or
            @($oldManifest.outputHashes.PSObject.Properties).Count -ne 19) {
            throw 'Existing Report folder is not the expected 19-table tenant snapshot.'
        }
        [IO.Directory]::Move($destination, $previous)
        $movedPrevious = $true
    }
    [IO.Directory]::Move($stage, $destination)
    $promoted = $true
    & $python.Source $generator --data-root $source --output $destination --validate-only
    if ($LASTEXITCODE -ne 0) { throw 'Promoted report-data validation failed.' }
    $promotionValidated = $true
    if ($movedPrevious) {
        try {
            [IO.Directory]::Delete($previous, $true)
            $movedPrevious = $false
        }
        catch {
            Write-Warning ("Validated Report is current, but the old Report directory could not be removed: {0}. {1}" -f
                $previous, $_.Exception.Message)
        }
    }
    [pscustomobject]@{Status='Completed'; ReportRoot=$destination; Tables=19;
        PythonVersion=$versionText}
}
catch {
    if ($promoted -and -not $promotionValidated -and $movedPrevious -and
        (Test-Path -LiteralPath $destination -PathType Container)) {
        $failed = Join-Path $powerBI ('Report.failed.' + $runId)
        [IO.Directory]::Move($destination, $failed)
        [IO.Directory]::Move($previous, $destination)
    }
    throw
}
finally {
    if (Test-Path -LiteralPath $stage -PathType Container) {
        try { [IO.Directory]::Delete($stage, $true) }
        catch { Write-Warning ("Temporary report stage could not be removed: {0}" -f $stage) }
    }
    if (Test-Path -LiteralPath $ciWorkRoot -PathType Container) {
        try { [IO.Directory]::Delete($ciWorkRoot, $true) }
        catch { Write-Warning ("Temporary CI stage could not be removed: {0}" -f $ciWorkRoot) }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCMrefWS+3rd987
# j1DLPE8FG0jKjb+UAlsTnT6DTEXXE6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIO9Sljay5V3ro+K3gf5dC+G9McCIsvqIcN7uGAl0+M1TMA0GCSqG
# SIb3DQEBAQUABIIBgD1BFh6sIp20BEjktj+cAZ2eDIFxctqBHyTdUA+lfl+ozfsu
# eBrDYqhZfIR/aFMORZb9LhMVX/8sHrsQCB8s0KSPeTjUGsNt5wSSNseQO0JN6Fu5
# 3TX7sxW2srxk62VlJZ2TFSpdRAUYBQ/eHon1sWCMWdPDM50+NL73VT+B/qSB6qfC
# tOLRa2d+V8R0xyGhXd/Fqslixj/mF5CKcFk0YqE0GiVqhuUNhXOJJQgeKBLKwGGY
# eg5p0B90Jp//df3q2NASjAgd7wdGu9iTB6c2F4Rx3J2ZEozKSswxPrt0Qq/eDsnV
# nhNkDpV5NFqWMvxWB3q7UqEslFGasv8o7S/rwy6+Nk8ks5Mj0nRXW4hYhiLy4wmH
# g/cyVtMEdTDclX0/VdbNls+drtgFnBU/HUTGoARurSnQ145xH0UAQOw55/je+OzE
# Pw0rc+QQX2X5qStcADo3PzlgcIzuyP/tiG7E8hOUy5vQj9IS7b/x7jBHQZnRzc4n
# bZuQ1GD8+hhL5viCrKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDExMzMy
# NDNaMC8GCSqGSIb3DQEJBDEiBCA7YKsWzRCtLMt9y/ccrNt7wBqxBvU6fTgtqXJQ
# J7oeXDANBgkqhkiG9w0BAQEFAASCAgB6QEUISMdGY5vcxP+84Cn8JjhLLkFUukdz
# fCwQYFcj8dcUMFpAqesS8qCFskiggWnpVS8UPIy6+cYUAhz8lBz/APn/EHlYVaLx
# ys7u7jvBPphMBq/81nvceWWqkwsF7EfShdvIFDniAyLnCYAFwsDjKm0KHH/nLqOd
# WmAuB2JlOu9jKaFZRafBspnY9xjLG2t/5CtTzayX9bmF+E3jILn0ypxH9o6Xq8ZW
# W3IMbPOW4aiuKHHbFx+ucy6QFAPq18koPobI59AMg9UjqBJNbEgIEy2dUeUWy0Vd
# Wo8hyAbK15Qj2U24BV6VADwM1esg+odGFgpNWCPbjw41LZBRS6rdvDff1Mr4MLb2
# s7wcHXshxDb7Cd9Pf4pW81aoXiL1EFmeQXoreseg0OKphqDYwzhE3LmpjC0BLxrN
# kO6G25tP8/FayxRohoS3SM95G2mYeLeHD9kFjRR3mw5/EW4qXqL6TXaFDyLCfaCh
# C1KOzvg07w5DCb2yEdXvqf12FlaRVb5NvVj+i8lNvY82ojg6mSf9Fjd7BWrxI08S
# DQiGema/Jbw1fXly1FocVGhQFLq0sn3WghhHK8euQKoRt4X8jixpwoZI5WCc3ahP
# JzCN37V5DKJGOtqvWUtQ0viGx92BInMGVcs3DO/BRIkrsG8Ba7KT1nIa2kNwaQFU
# g4DboJZIDQ==
# SIG # End signature block
