#Requires -Version 7.0
<#
.SYNOPSIS
    Creates the fictional "Contoso Sample App" MSI used by the example profile.

.DESCRIPTION
    Builds a minimal per-machine MSI with the Windows Installer COM API. The package contains
    no files: it only writes HKEY_LOCAL_MACHINE\SOFTWARE\Contoso\IntunePackagerSample\Version
    (64-bit view) so it installs and uninstalls cleanly on a test device. ProductCode and
    UpgradeCode are fixed fictional GUIDs that match Contoso-SampleApp.profile.json.template;
    the package code is new for each build.

.VERSION
    1.1.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com

.EXAMPLE
    pwsh -NoProfile -ExecutionPolicy Bypass -File .\Examples\SmartM365-IntunePackager-New-SampleMsi.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputFolder = (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'SmartM365\IntunePackager\Samples\Contoso-SampleApp\Source'),
    [ValidatePattern('^\d+\.\d+\.\d+$')][string]$ProductVersion = '1.0.0',
    # Adds an incompressible ContosoSampleApp-Payload.bin of this size next to the MSI (large upload tests).
    [ValidateRange(0, 4096)][int]$PayloadSizeMB = 0,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$productCode = '{5F0C3B2A-6D41-4E8B-9A57-2C1D7E3F9B10}'
$upgradeCode = '{A3E1D6C4-2B7F-4A90-8C35-6E9F1B2D4C78}'
$componentId = '{0B6E2D4F-8A13-4C57-9E21-7D3F5A9C1E64}'
$registryKey = 'SOFTWARE\Contoso\IntunePackagerSample'

function Write-SampleLog {
    param([string]$Message)
    Microsoft.PowerShell.Utility\Write-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Invoke-MsiSql {
    param($Database, [string]$Sql)
    $view = $Database.OpenView($Sql)
    try { [void]$view.Execute(); [void]$view.Close() }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
}

function Add-MsiRows {
    <#
    .SYNOPSIS
        Inserts rows with a parameterized statement. Integers use IntegerData, $null stays null.
    #>
    param($Installer, $Database, [string]$Table, [string[]]$Columns, [object[][]]$Rows)

    $columnList = ($Columns | ForEach-Object { '`{0}`' -f $_ }) -join ', '
    $placeholders = (@('?') * $Columns.Count) -join ', '
    $view = $Database.OpenView(('INSERT INTO `{0}` ({1}) VALUES ({2})' -f $Table, $columnList, $placeholders))
    try {
        foreach ($row in $Rows) {
            $record = $Installer.CreateRecord($Columns.Count)
            try {
                for ($i = 0; $i -lt $Columns.Count; $i++) {
                    $value = $row[$i]
                    if ($null -eq $value) { continue }
                    if ($value -is [int]) { $record.IntegerData($i + 1) = $value }
                    else { $record.StringData($i + 1) = [string]$value }
                }
                [void]$view.Execute($record)
            }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
        }
        [void]$view.Close()
    }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
}

[IO.Directory]::CreateDirectory($OutputFolder) | Out-Null
$msiPath = Join-Path -Path $OutputFolder -ChildPath 'ContosoSampleApp.msi'
if (Test-Path -LiteralPath $msiPath) {
    if (-not $Force) { throw "The sample MSI already exists. Use -Force to replace it: $msiPath" }
    Remove-Item -LiteralPath $msiPath -Force
}

$installer = New-Object -ComObject WindowsInstaller.Installer
$database = $null
try {
    # 3 = msiOpenDatabaseModeCreate
    $database = $installer.OpenDatabase($msiPath, 3)

    Invoke-MsiSql $database 'CREATE TABLE `Property` (`Property` CHAR(72) NOT NULL, `Value` LONGCHAR NOT NULL LOCALIZABLE PRIMARY KEY `Property`)'
    Invoke-MsiSql $database 'CREATE TABLE `Directory` (`Directory` CHAR(72) NOT NULL, `Directory_Parent` CHAR(72), `DefaultDir` CHAR(255) NOT NULL LOCALIZABLE PRIMARY KEY `Directory`)'
    Invoke-MsiSql $database 'CREATE TABLE `Component` (`Component` CHAR(72) NOT NULL, `ComponentId` CHAR(38), `Directory_` CHAR(72) NOT NULL, `Attributes` SHORT NOT NULL, `Condition` CHAR(255), `KeyPath` CHAR(72) PRIMARY KEY `Component`)'
    Invoke-MsiSql $database 'CREATE TABLE `Feature` (`Feature` CHAR(38) NOT NULL, `Feature_Parent` CHAR(38), `Title` CHAR(64) LOCALIZABLE, `Description` CHAR(255) LOCALIZABLE, `Display` SHORT, `Level` SHORT NOT NULL, `Directory_` CHAR(72), `Attributes` SHORT NOT NULL PRIMARY KEY `Feature`)'
    Invoke-MsiSql $database 'CREATE TABLE `FeatureComponents` (`Feature_` CHAR(38) NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Feature_`, `Component_`)'
    Invoke-MsiSql $database 'CREATE TABLE `Registry` (`Registry` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` LONGCHAR LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Registry`)'
    Invoke-MsiSql $database 'CREATE TABLE `InstallExecuteSequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`)'
    Invoke-MsiSql $database 'CREATE TABLE `InstallUISequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`)'

    Add-MsiRows $installer $database 'Property' @('Property', 'Value') @(
        , @('ProductCode', $productCode)
        , @('UpgradeCode', $upgradeCode)
        , @('ProductName', 'Contoso Sample App')
        , @('ProductVersion', $ProductVersion)
        , @('Manufacturer', 'Contoso Ltd')
        , @('ProductLanguage', '1033')
        , @('ALLUSERS', '1')
        , @('ARPNOMODIFY', '1')
        , @('ARPNOREPAIR', '1')
    )
    Add-MsiRows $installer $database 'Directory' @('Directory', 'Directory_Parent', 'DefaultDir') @(, @('TARGETDIR', $null, 'SourceDir'))
    # Attributes 260 = registry key path (4) + 64-bit component (256).
    Add-MsiRows $installer $database 'Component' @('Component', 'ComponentId', 'Directory_', 'Attributes', 'Condition', 'KeyPath') @(, @('ContosoSampleRegistry', $componentId, 'TARGETDIR', 260, $null, 'ContosoSampleVersion'))
    Add-MsiRows $installer $database 'Feature' @('Feature', 'Feature_Parent', 'Title', 'Description', 'Display', 'Level', 'Directory_', 'Attributes') @(, @('Main', $null, 'Contoso Sample App', 'Fictional sample package.', 1, 1, 'TARGETDIR', 0))
    Add-MsiRows $installer $database 'FeatureComponents' @('Feature_', 'Component_') @(, @('Main', 'ContosoSampleRegistry'))
    # Root 2 = HKEY_LOCAL_MACHINE.
    Add-MsiRows $installer $database 'Registry' @('Registry', 'Root', 'Key', 'Name', 'Value', 'Component_') @(, @('ContosoSampleVersion', 2, $registryKey, 'Version', '[ProductVersion]', 'ContosoSampleRegistry'))
    Add-MsiRows $installer $database 'InstallExecuteSequence' @('Action', 'Condition', 'Sequence') @(
        , @('CostInitialize', $null, 800)
        , @('FileCost', $null, 900)
        , @('CostFinalize', $null, 1000)
        , @('InstallValidate', $null, 1400)
        , @('InstallInitialize', $null, 1500)
        , @('ProcessComponents', $null, 1600)
        , @('UnpublishFeatures', $null, 1800)
        , @('RemoveRegistryValues', $null, 2600)
        , @('WriteRegistryValues', $null, 5000)
        , @('RegisterProduct', $null, 6100)
        , @('PublishFeatures', $null, 6300)
        , @('PublishProduct', $null, 6400)
        , @('InstallFinalize', $null, 6600)
    )
    Add-MsiRows $installer $database 'InstallUISequence' @('Action', 'Condition', 'Sequence') @(
        , @('CostInitialize', $null, 800)
        , @('FileCost', $null, 900)
        , @('CostFinalize', $null, 1000)
        , @('ExecuteAction', $null, 1300)
    )

    # 20 = maximum number of summary properties updated.
    $summary = $database.SummaryInformation(20)
    try {
        $summary.Property(2) = 'Installation Database'
        $summary.Property(3) = 'Contoso Sample App'
        $summary.Property(4) = 'Contoso Ltd'
        $summary.Property(7) = 'x64;1033'
        $summary.Property(9) = '{' + [guid]::NewGuid().ToString().ToUpperInvariant() + '}'
        $summary.Property(14) = 200
        $summary.Property(15) = 2
        $summary.Property(19) = 2
        [void]$summary.Persist()
    }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($summary) }

    [void]$database.Commit()
}
finally {
    if ($null -ne $database) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database) }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer)
}

$payloadPath = Join-Path -Path $OutputFolder -ChildPath 'ContosoSampleApp-Payload.bin'
if ($PayloadSizeMB -gt 0) {
    # Random bytes cannot be compressed, so the .intunewin stays larger than the requested size.
    $buffer = [byte[]]::new(1MB)
    $stream = [IO.File]::Create($payloadPath)
    try {
        for ($i = 0; $i -lt $PayloadSizeMB; $i++) {
            [Security.Cryptography.RandomNumberGenerator]::Fill($buffer)
            $stream.Write($buffer, 0, $buffer.Length)
        }
    }
    finally { $stream.Dispose() }
    Write-SampleLog ("Created incompressible payload: {0} ({1} MB)" -f $payloadPath, $PayloadSizeMB)
}
elseif (Test-Path -LiteralPath $payloadPath) {
    Remove-Item -LiteralPath $payloadPath -Force
}

Write-SampleLog "Created fictional sample MSI: $msiPath"
Write-SampleLog "ProductCode=$productCode; ProductVersion=$ProductVersion; Registry=HKEY_LOCAL_MACHINE\$registryKey\Version"

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB2R6ofYHAIyoWu
# Ukw6pBY3n6dA41+68Oa6tLBd+sRBO6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMvukLuIxHMhwqZj00hYGJnsc4T7N1dkea9pxWBkyu1RMA0GCSqG
# SIb3DQEBAQUABIIBgGdiOqCdg5KxyYxQG5f2Co1XHM68cj0Pqp6G1s9nIA6N7GSu
# nKedcTENVK00zj6JgMPxgTado6cGm36KT+ez4kELT9/O5XQHQNYw4ksp2EPN5Uje
# ei72pl0xToSIBccDI3/Na5t8QiBGPb1mISUs5Bj943D+JEVwgBtxX9a2FdKA+sSn
# nk+V1Mrk3wQwn2bdih0t9gLjhWrDYNC7OQ9W/Is7nxeVRSDvgoza2t3E/F6rq3Ci
# jM82tgB1gH3oeojL1O2tv0M0lqNUY1qq21Q1CETo6CCkbEEum6eYzVJk/e5fDkxk
# 1yg7aSUgfhfmr0cvSXJpg2apI9ZITRhNR+H9KagS0AucJRrQWRtJjQuUBXGpwxSE
# t7Fr28jljwUSwybA5cUdJ28FJkCZL20ZFhGEh1ORkgIWGCaZD0lrPNPemzFzMEn8
# 766LRqimbDrh5D6ekDEIStpDiptgi4rcZZlB2O3+a7RewpnCeS+OzKK618Co+r+g
# 5gYXlLmWv23wlElgqqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzJaMC8GCSqGSIb3DQEJBDEiBCCPJuryrXU6clP+miuCl99Wef/9JzMRDSHFYP6u
# H2n4JjANBgkqhkiG9w0BAQEFAASCAgAmdEBfMhTiBmRf9NGb82wgnuHNBDnbKzP/
# //kddP2S1h0FN0mtkDR7kOBdQVcj+n9WdxHgN7d8h5YiHvRvBxXn7I6c8lxP/IW0
# A6eyDqD+vswkj21tdKGjeacOapS/JoIilaNokU/Y9PQNFxWdI0NjEBRwBijRJes+
# tiod4LA5IDNowXTT+EpCvRqVHET3I8iVNGxJ4Gh5F7DmQTmNzJ5m0kBwQeX+EJ7l
# ixB6P7EZcIq7CjzcPw8LsJH1937y+YwRf6Y9gnoy1lkT/4adIk6oMjFCHBGdLo5O
# /kapAqLzUEifja+EuF06ZGkOTJK3COnTbHMGxz/2FRsTevuzTD6agfNwvPWjpwxN
# 4wjJjr9nBJsv1EWdTz/5OR2Kf4udW/KIAiihUDxH8DFKWCa0yiEc+Of0mAtxzCOc
# HOYmnp5moT7/og389/MK6n6RUr8FyMEOIoSMTQTEIYf5UH1NSjp27/cc+QF1K6gY
# iIhCYm0DNtrEEAdmnxqY24UM0Jb/qZ4VCPLBb72JosgamPfP0OIp9NBpr/Za0zxR
# dDGMlzJ90Oh65QJjzwgcDe1BznCg+RSnwMLEXV2t53Uv5/a22b54+YAVuRPOvvDB
# 6mVTM4LVdHQuGwbx2Jo2OZk51u4hF5rcu6bjLsBkucPGzU83HXidMJqJ1x7c549/
# ETLo4iV2rw==
# SIG # End signature block
