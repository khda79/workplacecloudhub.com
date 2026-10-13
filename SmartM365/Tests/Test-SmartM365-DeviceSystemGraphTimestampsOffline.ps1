#Requires -Version 7.0
<#
.SYNOPSIS
    Synthetic-only UTC serialization and prepared-consumer regression tests.
    Parses collector functions without executing its initialization or Graph calls.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$collectorPath = Join-Path $repoRoot 'SmartM365/SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Device-System-Inventory.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($collectorPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'Collector source did not parse.' }
$definition = $ast.Find({ param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'ConvertTo-DeviceSystemGraphUtcTimestamp'
}, $true)
if ($null -eq $definition) { throw 'UTC serializer was not found.' }
. ([scriptblock]::Create($definition.Extent.Text))

$script:checks = 0
$script:warnings = [Collections.Generic.List[string]]::new()
function Assert-Offline([bool]$Condition, [string]$Message) {
    $script:checks++
    if (-not $Condition) { throw $Message }
}
function WriteLog([string]$Message, [string]$Level) {
    Assert-Offline ($Level -eq 'WARNING') 'An unqualified timestamp must produce a warning.'
    $script:warnings.Add($Message)
    # Also verifies that incidental logger output cannot contaminate the CSV value.
    'synthetic logger output'
}

# Execute only the actual record-building assignment with synthetic inputs.
$recordAssignment = $ast.Find({ param($node)
    $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$record'
}, $true)
if ($null -eq $recordAssignment) { throw 'Collector record assignment was not found.' }
function Get-ExpectedThirdPartyAVStatus { 'Detected' }
$ExpectedThirdPartyAV = [pscustomobject]@{ DisplayName = 'Synthetic AV' }
$psData = [pscustomobject]@{
    SecureBoot = 'Enabled'; BIOSVersion = 'synthetic'; BIOSDate = '2026-01-01'; FirmwareType = 'Uefi'
    LastBootUpTime = '2026-01-01 10:00:00'; ThirdPartyAVStatus = 'Detected'; ThirdPartyAVProducts = 'Synthetic AV'
    WindowsUpdateRebootRequired = 'False'; BitLockerOSProtectionStatus = 'On'; BitLockerOSConversionStatus = 'FullyEncrypted'
    LastConsoleLogonStatus = 'Observed'; LastConsoleLogonUser = 'SYNTHETIC\\user'; LastConsoleLogonSid = 'S-1-5-21-1-2-3-1001'; LastConsoleLogonUtc = '2026-10-12T10:00:00Z'
}
$hw = $null
$device = [pscustomobject]@{
    deviceName = 'synthetic-device'; id = 'synthetic-id'; azureADDeviceId = 'synthetic-cloud-id'
    userPrincipalName = 'user@example.invalid'; managementState = 'managed'; complianceState = 'compliant'
    lastSyncDateTime = [datetimeoffset]'2026-10-12T14:13:32+02:00'
}
$psLastUpdate = [datetime]::new(2026, 10, 12, 14, 13, 32, [DateTimeKind]::Utc)
$localObservationUtc = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $psLastUpdate -FieldName 'LocalObservationUtc'
$localFreshness = 'Fresh'
$UseRecurringDeviceSystem = $true
. ([scriptblock]::Create($recordAssignment.Extent.Text))
Assert-Offline ($record.PlatformScriptLastUpdate -ceq '2026-10-12T14:13:32.0000000Z') 'PlatformScriptLastUpdate bypasses the serializer.'
Assert-Offline ($record.LocalObservationUtc -ceq $record.PlatformScriptLastUpdate -and $record.LocalDataFreshness -eq 'Fresh') 'Recurring observation alias or freshness was lost.'
Assert-Offline ($record.LastConsoleLogonUser -eq 'SYNTHETIC\\user' -and $record.LastConsoleLogonUtc -eq '2026-10-12T10:00:00.0000000Z') 'Local console logon was not serialized.'
Assert-Offline ($record.LastConsoleLogonSid -eq 'S-1-5-21-1-2-3-1001') 'The console SID was not exported.'
Assert-Offline ($record.LastSyncDateTime -ceq '2026-10-12T12:13:32.0000000Z') 'LastSyncDateTime bypasses the serializer.'
Assert-Offline ($record.PlatformScriptLastUpdate -is [string] -and $record.LastSyncDateTime -is [string]) 'The exported timestamps must be strings, not date objects.'

$originalCulture = [Threading.Thread]::CurrentThread.CurrentCulture
$originalUICulture = [Threading.Thread]::CurrentThread.CurrentUICulture
try {
    foreach ($culture in @('fr-FR', 'en-US')) {
        [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($culture)
        [Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo($culture)
        $knownUtc = [datetime]::new(2026, 10, 12, 14, 13, 32, [DateTimeKind]::Utc)
        $cases = @(
            @{Value = $knownUtc; Expected = '2026-10-12T14:13:32.0000000Z'},
            @{Value = $knownUtc.ToLocalTime(); Expected = '2026-10-12T14:13:32.0000000Z'},
            @{Value = ([datetimeoffset]'2026-10-25T02:30:00+02:00').UtcDateTime.ToLocalTime(); Expected = '2026-10-25T00:30:00.0000000Z'},
            @{Value = ([datetimeoffset]'2026-10-25T02:30:00+01:00').UtcDateTime.ToLocalTime(); Expected = '2026-10-25T01:30:00.0000000Z'},
            @{Value = [datetimeoffset]'2026-10-12T14:13:32+02:00'; Expected = '2026-10-12T12:13:32.0000000Z'},
            @{Value = '2026-10-12T14:13:32Z'; Expected = '2026-10-12T14:13:32.0000000Z'},
            @{Value = '2026-10-12T14:13:32.1234567+02:00'; Expected = '2026-10-12T12:13:32.1234567Z'},
            @{Value = '2026-10-12T14:13:32-07:00'; Expected = '2026-10-12T21:13:32.0000000Z'},
            @{Value = '2026-10-25T02:30:00+02:00'; Expected = '2026-10-25T00:30:00.0000000Z'},
            @{Value = '2026-10-25T02:30:00+01:00'; Expected = '2026-10-25T01:30:00.0000000Z'},
            @{Value = '2026-03-29T01:30:00+01:00'; Expected = '2026-03-29T00:30:00.0000000Z'},
            @{Value = '2026-03-29T03:30:00+02:00'; Expected = '2026-03-29T01:30:00.0000000Z'}
        )
        foreach ($case in $cases) {
            $before = $script:warnings.Count
            $actual = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $case.Value -FieldName 'PlatformScriptLastUpdate'
            Assert-Offline ($actual -ceq $case.Expected) "Known timestamp did not preserve its instant in $culture."
            Assert-Offline ($script:warnings.Count -eq $before) 'A qualified date raised a warning.'
            $roundTrip = [pscustomobject]@{PlatformScriptLastUpdate = $actual} | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv
            Assert-Offline ($roundTrip.PlatformScriptLastUpdate -ceq $case.Expected) "CSV serialization depends on $culture."
        }
        foreach ($empty in @($null, '', '   ')) {
            $before = $script:warnings.Count
            $actual = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $empty -FieldName 'LastSyncDateTime'
            Assert-Offline ($null -eq $actual -and $script:warnings.Count -eq $before) 'Missing observations were fabricated or logged as malformed.'
        }
        foreach ($invalid in @(
            [datetime]::new(2026, 10, 25, 2, 30, 0, [DateTimeKind]::Unspecified),
            '12/10/2026 14:13:32', '10/12/2026 2:13:32 PM', '2026-10-12T14:13:32',
            '2026-02-30T14:13:32Z', '2026-10-12T14:13:32+99:00', 'not-a-date', 123, $false
        )) {
            $before = $script:warnings.Count
            $actual = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $invalid -FieldName 'PlatformScriptLastUpdate'
            Assert-Offline ($null -eq $actual) 'An unqualified date acquired an invented time zone or leaked logger output.'
            Assert-Offline ($script:warnings.Count -eq $before + 1) 'An unqualified date needs exactly one warning.'
            Assert-Offline ($script:warnings[$before] -match '^PlatformScriptLastUpdate is not a qualified Graph timestamp:') 'Warning does not identify the affected field.'
        }
    }
}
finally {
    [Threading.Thread]::CurrentThread.CurrentCulture = $originalCulture
    [Threading.Thread]::CurrentThread.CurrentUICulture = $originalUICulture
}

# Pure consumer functions only: no entry script, tenant loader, output writer or preparation.
Import-Module (Join-Path $repoRoot 'SmartWorkplaceIntelligence/scripts/DeviceSystemEvidence.psm1') -Force
Import-Module (Join-Path $repoRoot 'SmartWorkplaceIntelligence/scripts/EndpointProtectionEvidence.psm1') -Force
$syntheticSystem = [pscustomobject]@{
    TenantKey = 'synthetic'; DeviceId = 'device-one'; DeviceName = 'synthetic-device'; AzureADDeviceId = 'cloud-id'
    UserPrincipalName = 'user@example.invalid'; PlatformScriptLastUpdate = $record.PlatformScriptLastUpdate
    LastSyncDateTime = $record.LastSyncDateTime; ExpectedThirdPartyAVName = 'Synthetic AV'
    ExpectedThirdPartyAVStatus = 'Detected'; ThirdPartyAVStatus = 'Detected'; ThirdPartyAVProducts = 'Synthetic AV'
}
$inventory = [pscustomobject]@{TenantKey = 'synthetic'; 'Device ID' = 'device-one'; 'Primary user UPN' = 'user@example.invalid'; Encrypted = ''; Compliance = ''}
$preparedInventory = [pscustomobject]@{'Device Source ID' = 'device-one'; Country = 'Unknown'}
$policy = ConvertTo-EndpointProtectionPolicy -Configuration @{
    EndpointProtectionPolicy = @{
        Enabled = $true; PolicyVersion = 'synthetic-v1'; MaxObservationAgeHours = 72
        Requirements = @(
            @{Role = 'Antivirus'; Product = 'SyntheticAV'; ExpectedMode = 'Active'; EvidenceProductName = 'Synthetic AV'; Reason = ''},
            @{Role = 'EDR'; Product = 'SyntheticAV'; ExpectedMode = 'Active'; EvidenceProductName = ''; Reason = ''}
        ); ScopeRules = @()
    }
} -TenantKey 'synthetic'
foreach ($case in @(
    @{Hours = 24; System = 'Within policy'; Endpoint = 'Current'; Registration = 'Detected'},
    @{Hours = 72; System = 'Within policy'; Endpoint = 'Current'; Registration = 'Detected'},
    @{Hours = 72.01; System = 'Stale'; Endpoint = 'Stale'; Registration = 'Unknown'},
    @{Hours = -1; System = 'Future observation'; Endpoint = 'Future observation'; Registration = 'Unknown'}
)) {
    $asOf = ([datetimeoffset]$record.PlatformScriptLastUpdate).AddHours($case.Hours)
    $systemEvidence = @(New-DeviceSystemEvidence -SystemRows @($syntheticSystem) -InventoryRows @($inventory) -UserRows @() -PreparedInventoryRows @($preparedInventory) -MaxObservationAgeHours 72 -AsOfUtc $asOf)
    Assert-Offline ($systemEvidence.Count -eq 1 -and $systemEvidence[0].'Observation Time Basis' -eq 'Explicit offset') 'Prepared consumer lost the explicit UTC timestamp.'
    Assert-Offline ($systemEvidence[0].Freshness -eq $case.System) 'Device System freshness boundary regressed.'
    $endpointEvidence = @(New-EndpointProtectionEvidence -Policy $policy -InventoryRows @($preparedInventory) -SystemRows $systemEvidence -AsOfUtc $asOf)
    $av = @($endpointEvidence | Where-Object Role -eq 'Antivirus')[0]
    $edr = @($endpointEvidence | Where-Object Role -eq 'EDR')[0]
    Assert-Offline ($av.Freshness -eq $case.Endpoint -and $av.'Registration State' -eq $case.Registration) 'Endpoint freshness/registration boundary regressed.'
    Assert-Offline ($av.'Evaluation State' -eq 'Unknown' -and $edr.'Evaluation State' -eq 'Unknown') 'Timestamp repair must not invent AV or EDR health.'
}
$syntheticSystem.PlatformScriptLastUpdate = ''
$missingEvidence = @(New-DeviceSystemEvidence -SystemRows @($syntheticSystem) -InventoryRows @($inventory) -UserRows @() -PreparedInventoryRows @($preparedInventory) -MaxObservationAgeHours 72)
Assert-Offline ($missingEvidence[0].Freshness -eq 'Missing' -and -not $missingEvidence[0].'Observation UTC') 'Missing timestamps were repaired using preparation time.'
$missingEndpoint = @(New-EndpointProtectionEvidence -Policy $policy -InventoryRows @($preparedInventory) -SystemRows $missingEvidence)
Assert-Offline (@($missingEndpoint | Where-Object {$_.'Evaluation State' -ne 'Unknown' -or $_.'Registration State' -ne 'Unknown'}).Count -eq 0) 'Missing timestamps became qualified endpoint health or registration.'
[void](ConvertTo-DeviceSystemGraphUtcTimestamp -Value 'date-without-offset' -FieldName 'LastSyncDateTime')
Assert-Offline ($script:warnings[$script:warnings.Count - 1] -match '^LastSyncDateTime is not a qualified Graph timestamp:') 'LastSyncDateTime warning does not identify the affected field.'
[pscustomobject]@{Passed = $true; Checks = $script:checks; Sources = 'Synthetic only'; CollectorExecuted = $false; PreparationExecuted = $false; Transfer = $false; PowerBIRefresh = $false; OutputFilesWritten = 0} | ConvertTo-Json

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBdwGRUxljmxbs
# BB+/bT/ASZSwyF+TTeWWG2bDmi+tnqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIER8qxblT3V6B4LdHI4lkjSyXh+iz2WtT9qT2RyLj5oaMA0GCSqG
# SIb3DQEBAQUABIIBgE28krjidUDTiuUY8+1d52fZCjDCRT+288cOxHb0gTjjFKYu
# 9BJ+p+kyynqHmqq7rtXUqqH6QO024BlyxGIpNO7HkMGaApwf2h5fJCbU3kSpPnqU
# pU99LxoOQ/47NYuodtOXp06JenGMMncFp/MgUQy/0ORwVw2PSGymcIFp2qbTyCgY
# eXVg4LapviQSZR6yQ1S0ebKleiGA4M3ve13RZrLxQ8sTvYytpVHXliK2nF50/Zqp
# iMHOgPdWJqTn/+wREITC33Utll0cdJDUn9sMIFrLSOf9BY0YBLpz787niY9iGqDY
# 53DcKPA4Les6V9JFsNa3eirG2+FqtFONuzJ+fn09qV+aikTA52Ew1ymNtXFgzUDy
# 0d1wiXaHtlC2Tr5p2wccObolpvSeDUFpR/FYBWrqzAq/boHGHS3W5ajNSgF9vEyI
# CsOTzptFrx5KDmVl27w3MUuBB0OBLqdwSGJcR1c9umtxpjbCJS1GhHfHg5bNezz1
# kGe9u6+h2TrqdCy7fqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTMxNTQ1
# MjhaMC8GCSqGSIb3DQEJBDEiBCBJUTgebe6a37x9nvyPC97Tf/0Tu/iLJOaYl9qE
# vb+BYjANBgkqhkiG9w0BAQEFAASCAgB5kKPlgzK9qZUdD+fJi9Jjm5GH27IOTtQO
# ahYBIOOoeI5obvSbX+Q78DetLdOMfW+VnLu35n0H1KLqG8I476VTioPR1p7dXS0z
# PB+ZAYncsJQqFmZJ4mGnVRyv/p2rzDH9oVnn90XKWHfGI43R5UIfWSrrpTwAdi3v
# u/QdHT/cxMdTwytZIb85t8P3adhptBSGa3ugmQS9fZ9SyzMtEpzHasF3vD8m0CgR
# Mx6+DgBkD1wpTcFBaqMMFe1NVaDG9JHfiiCVbER5TBkfaYn8687gb2iHty/miQRS
# Da25Rq23MMdx2xXDIioYa77dX6rlytWg5c+OmSwWaOpYP/jerCcDAcEaKYRceQ5M
# WFi6ygc/k8XR838KbQM2epW7Y4BGoIRuW7ZUTTJbI8rM3i1xN1Wm+W7CvYcN/+y6
# Fig1+hM9iKJMxCXBkCx+NecO+3jCXSH7I3I5/RuwakD+Xbl8zCqkk8XB6Imc1/1B
# s1CmN7yPqkpXihhID6HngJgNp5ub+MT9/ACKdTh/dtFUG10SvyUJAxrROusRfzLe
# MWUkpN9UBWi8EYWiYAOGLAySvZ9+63oN5B9obSxwvc3dywxolCu8LJcbEdY9GjRv
# cHW3lfP2VRVaPq6uQ4k2BWEASdZHEBBQmht4IegmUiZKp2eUOsL6Uoqi7Gbuzb4A
# dTIYllgZ/g==
# SIG # End signature block
