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
$licenseSnapshot = Join-Path $inventoryRoot 'M365_Licenses_ReportSnapshot.json.txt'
$source = [IO.Path]::GetFullPath($paths.LatestOutputRootPath)
$powerBI = Join-Path $source 'PowerBI'
$destination = Join-Path $powerBI 'Report'
$priorReportManifest = Join-Path $destination 'report-data.manifest.json.txt'
if (-not (Test-Path -LiteralPath $licenseSnapshot -PathType Leaf) -and
    (Test-Path -LiteralPath $priorReportManifest -PathType Leaf)) {
    $priorReport = Get-Content -LiteralPath $priorReportManifest -Raw -ErrorAction Stop |
        ConvertFrom-Json -ErrorAction Stop
    if ($priorReport.PSObject.Properties['licenseReportSnapshot'] -and
        $null -ne $priorReport.licenseReportSnapshot) {
        throw 'Previously prepared SmartInventory license report snapshot is now missing; retaining the prior CMDB ReportData.'
    }
}
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
    if (Test-Path -LiteralPath $licenseSnapshot -PathType Leaf) {
        $arguments += @('--license-report-snapshot', $licenseSnapshot)
    }
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
        $oldTableCount = @($oldManifest.outputHashes.PSObject.Properties).Count
        $hasOldLicenseSnapshot = $oldManifest.PSObject.Properties['licenseReportSnapshot'] -and
            $null -ne $oldManifest.licenseReportSnapshot
        if ([string]$oldManifest.identity.TenantKey -cne [string]$paths.TenantKey -or
            $oldTableCount -ne $(if ($hasOldLicenseSnapshot) { 22 } else { 19 })) {
            throw 'Existing Report folder does not match the expected tenant snapshot table contract.'
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAelIuh1CeWBKMr
# XVayCr+jamD3NynclZy4EMD6qUlCDKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDnww3VI+9i15fxwIAAt+oAHH9UzuCkQcSsmmiqURb3tMA0GCSqG
# SIb3DQEBAQUABIIBgLGXET0IPQV3nWF5LD6Jzo7dyLpc/3ZkLCxQVhQ83kVhDwo7
# xvuE/HRIPP/xq/5Eli6vq+qRWWEDKE0VK5W5tz9JD17MDVLvwbhucV+hTCFN5vl5
# f/QDANaXAE5ZnZJcIMEQbUFt9vhwMeQbieJphceZ5Sn1OJ+bnSPfrlCuZ3rtjjQn
# SkIgLCkuFfvBpNBruMnCkSDi1F9/cpyaSnGDxy1vELz1z4Svqs2OgeRMVpNlIsU7
# 7lRpS0BBUDmnOhrwnIW96HaSksFu0Jdxh/piJWy3GA74C20DqfkpM5uZ523+hUjG
# 287uOTz2W/GaNnuS1V0rH6ry8qJOpvOTxr8hDcInRVxKh2Y6/hI8ZZjFthuL+f6m
# bLYyphRUX/62mpJ71p84R7VsVBtfEhvU1K9PQa9Y2oCGomssKlesPTAQgqnSQOKw
# YeWT6lvK2lPMWzdf2FgvE0qjbBfejj5Ibcud3gpfM4mcAuxvYkZoQ3wBh1JAJ/yq
# Tve4YYnzABIPy12pUKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxNTE2
# MzdaMC8GCSqGSIb3DQEJBDEiBCCRAMJL4f4xGvbyrw5YCGaTBpFd0swrR96QtaGE
# Wvv3+DANBgkqhkiG9w0BAQEFAASCAgBBEIC9awcnHDo5UVaw+bPA6qlozhMfyqJa
# m5B13PZwlNHB7YYjg/gxnd73LBJBd3/fHNJJ0OoriboCsJv/xpjGg959OlOiVDWD
# utmAAMKFRYLOUVllsHzRCxZklODCpk5qs1OkpISUgP0/102Jyno4+R5jyCdQ3zvj
# gag/a6/7DsDRWJVZONBhpADBVLaCK8Zm8nfVpC9CjFbX/1RzdA7aMFD3ngaNmaqu
# ns/0Uq2W489bo9wMCRj3s9Oa3Ga6SIN80A43zAhWTCD8hpEvrkrQoaY2Wjj57gCh
# iz25DCWGKjmy0Hddnxie6jrk7qpKC80E3OkEM7bTqEVzPdGCEJmodzGP3o4nHzW1
# 5izpd9ApKUMBt8CJjmILrJKsIqQ7UwGiVC8QN1YUppmCFSo8I4Y23R1AEKdaFMRj
# QVg4W41WGSnwQ30SzATrxZgbkEwAXCinXjVd7+ESFXZOwBqyKacnRCsv57m3GuTh
# hHJe06W/3jzIzATmMCGQ+owbi2tHKr1qw42+TGfB4BuE1rvO8eesu8kwkz+R7CAC
# BQly0oA41mp+BKFe9TO8woetEN2K3ZeuGXTcrlc+Avez15ADv5DyS0c0yXOQZbDQ
# QTfLyWaWVYhOPRNnlkHIiDFuwrq5SLtY51hsFGuHrH7Jje9TAhp28yml0yJFDe32
# /r0DsgEsAw==
# SIG # End signature block
