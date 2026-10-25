#Requires -Version 7.0
# Version 1.0.0
[CmdletBinding()]
param()
# Synthetic fixtures only. No tenant initialization, collectors, production files or Power BI.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:checks = 0
function Check([bool]$Pass, [string]$Message) { $script:checks++; if (-not $Pass) { throw $Message } }
Import-Module (Join-Path $PSScriptRoot 'DeviceReadinessGapEvidence.psm1') -Force
$reference = [datetimeoffset]'2026-10-25T08:00:00Z'

# Reboot time: offset-less values are never converted to UTC.
$fact = ConvertTo-RebootFact -Value '2026-10-15 23:30:00' -Reference $reference
Check ($fact.Basis -eq 'Offset-less source time' -and $fact.Utc -eq '' -and $fact.Days -eq 10 -and $fact.AgeBasis -match 'offset unknown') 'Offset-less time was converted or miscounted'
$fact = ConvertTo-RebootFact -Value '2026-10-26 01:00:00' -Reference $reference
Check ($fact.Days -eq 0 -and $fact.AgeBasis -ne 'Future') 'A local date one day ahead of UTC is not in the future'
$fact = ConvertTo-RebootFact -Value '2026-10-27 00:00:00' -Reference $reference
Check ($fact.AgeBasis -eq 'Future' -and $null -eq $fact.Days) 'Future offset-less value produced an age'
$fact = ConvertTo-RebootFact -Value '2026-10-24T23:30:00-02:00' -Reference $reference
Check ($fact.Basis -eq 'Explicit offset' -and $fact.Utc -eq '2026-10-25T01:30:00Z' -and $fact.Days -eq 0) 'Explicit offset was not converted to UTC'
$fact = ConvertTo-RebootFact -Value '2026-10-15T10:00:00Z' -Reference $reference
Check ($fact.Days -eq 10 -and $fact.AgeBasis -eq 'UTC calendar days') 'UTC value miscounted'
$fact = ConvertTo-RebootFact -Value '2026-10-25T09:00:00Z' -Reference $reference
Check ($fact.AgeBasis -eq 'Future' -and $null -eq $fact.Days) 'Future UTC value produced an age'
foreach ($value in 'Error', '31/02/2026 10:00:00', 'not a date') {
    $fact = ConvertTo-RebootFact -Value $value -Reference $reference
    Check ($fact.Basis -eq 'Invalid' -and $null -eq $fact.Days -and $fact.Utc -eq '') "Invalid value '$value' produced a date"
}
$fact = ConvertTo-RebootFact -Value '' -Reference $reference
Check ($fact.Basis -eq 'Missing' -and $null -eq $fact.Days) 'Missing value produced a date'

# Facts: AD priority, LocalSystem only when AD has no value, AD computers without Intune kept.
$intune = @(
    [pscustomobject]@{TenantKey = 't'; 'Device ID' = 'i-ad'; 'Device name' = 'PC-AD'; 'Azure AD Device ID' = 'E-AD'; PhysicalMemoryGB = '7,5'}
    [pscustomobject]@{TenantKey = 't'; 'Device ID' = 'i-local'; 'Device name' = 'PC-LOCAL'; 'Azure AD Device ID' = 'e-local'; PhysicalMemoryGB = 'Error'}
    [pscustomobject]@{TenantKey = 't'; 'Device ID' = 'i-invalid'; 'Device name' = 'PC-INVALID'; 'Azure AD Device ID' = 'e-invalid'; PhysicalMemoryGB = '16'})
$ad = @(
    [pscustomobject]@{TenantKey = 't'; ObjectGUID = 'G-AD'; ObjectGuidNormalized = '{E-AD}'; IntuneDeviceId = 'i-ad'; Name = 'PC-AD$'; LastRebootDateTime = '2026-10-01 08:00:00'}
    [pscustomobject]@{TenantKey = 't'; ObjectGUID = 'g-local'; ObjectGuidNormalized = 'e-local'; IntuneDeviceId = ''; Name = 'PC-LOCAL$'; LastRebootDateTime = ''}
    [pscustomobject]@{TenantKey = 't'; ObjectGUID = 'g-invalid'; ObjectGuidNormalized = 'e-invalid'; IntuneDeviceId = 'i-invalid'; Name = 'PC-INVALID$'; LastRebootDateTime = 'Error'}
    [pscustomobject]@{TenantKey = 't'; ObjectGUID = 'g-adonly'; ObjectGuidNormalized = 'e-adonly'; IntuneDeviceId = ''; Name = 'PC-ADONLY$'; LastRebootDateTime = '2026-09-01 08:00:00'})
$entra = @(
    [pscustomobject]@{DeviceId = 'e-ad'; IsPending = 'False'}
    [pscustomobject]@{DeviceId = 'e-adonly'; IsPending = 'True'})
$local = @(
    [pscustomobject]@{AzureADDeviceId = 'e-ad'; FirmwareType = 'Bios'; SecureBootStatus = 'NotSupported'; 'Last Reboot Date' = '2026-10-24 08:00:00'}
    [pscustomobject]@{AzureADDeviceId = 'e-local'; FirmwareType = 'Uefi'; SecureBootStatus = 'Enabled'; 'Last Reboot Date' = '2026-10-20 08:00:00'}
    [pscustomobject]@{AzureADDeviceId = 'e-invalid'; FirmwareType = 'Uefi'; SecureBootStatus = 'Error'; 'Last Reboot Date' = '2026-10-20 08:00:00'})
$rows = @(New-DeviceReadinessGapEvidence -IntuneRows $intune -EntraRows $entra -AdRows $ad -LocalSystemRows $local -Reference $reference)
Check ($rows.Count -eq 4) "Expected three Intune devices and one AD-only computer, got $($rows.Count)"
$byName = @{}; foreach ($row in $rows) { $byName[$row.'Device Name'] = $row }
$row = $byName['PC-AD']
Check ($row.'Inventory Scope' -eq 'Intune device' -and $row.'AD Object GUID' -eq 'e-ad' -and $row.'Azure AD Device ID' -eq 'e-ad') 'Intune/AD/Entra keys not normalized'
Check ($row.'Last Reboot Source' -eq 'AD computer' -and $row.'Days Since Last Reboot' -eq '24') 'AD reboot value lost its priority'
Check ($row.'Physical Memory (GB)' -eq '7.5' -and $row.'Firmware Type (source)' -eq 'Bios' -and $row.'Secure Boot Status (source)' -eq 'NotSupported' -and $row.'Entra Registration Pending' -eq 'False') 'Raw device facts changed'
$row = $byName['PC-LOCAL']
Check ($row.'Last Reboot Source' -eq 'Intune local system' -and $row.'Days Since Last Reboot' -eq '5' -and $row.'AD Object GUID' -eq 'e-local') 'LocalSystem fallback or AD match by Entra ID failed'
Check ($row.'Physical Memory (GB)' -eq '' -and $row.'Entra Registration Pending' -eq '') 'Unknown memory or pending state was not left blank'
$row = $byName['PC-INVALID']
Check ($row.'Last Reboot Source' -eq 'AD computer' -and $row.'Last Reboot Time Basis' -eq 'Invalid' -and $row.'Days Since Last Reboot' -eq '') 'An invalid AD value must not fall back nor produce an age'
$row = $byName['PC-ADONLY']
Check ($row.'Inventory Scope' -eq 'AD computer only' -and $row.'Device Source ID' -eq '' -and $row.'Entra Registration Pending' -eq 'True' -and $row.'Days Since Last Reboot' -eq '54' -and $row.TenantKey -eq 't') 'AD-only computer facts missing'
Check (@($rows | ForEach-Object { $_.PSObject.Properties.Name -join ',' } | Sort-Object -Unique).Count -eq 1 -and (($rows[0].PSObject.Properties.Name) -join ',') -ceq ((Get-DeviceReadinessGapColumns) -join ',')) 'Row columns differ from the declared columns'
$contract = (Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/prepared-evidence-contract.json.txt') -Raw | ConvertFrom-Json).tables | Where-Object file -EQ 'DeviceReadinessGapEvidence.csv'
Check (($contract.columns.name -join ',') -ceq ((Get-DeviceReadinessGapColumns) -join ',')) 'Declared columns differ from the contract'

# Complete generator: the existing outputs are byte-identical with and without the new output.
$root = Join-Path ([IO.Path]::GetTempPath()) ('DeviceReadinessGap-' + [guid]::NewGuid().ToString('N'))
try {
    $last = Join-Path $root 'DATA-LAST'; New-Item -ItemType Directory -Path $last -Force | Out-Null
    $schema = (Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/source-schema.json') -Raw | ConvertFrom-Json).files
    function Write-Source([string]$Name, [hashtable[]]$Rows) {
        # Declared source columns plus the extra exported columns a fixture row needs.
        $columns = [Collections.Generic.List[string]]::new(); foreach ($column in @($schema.$Name.columns) + @($Rows | ForEach-Object { $_.Keys })) { if ($column -and -not $columns.Contains($column)) { $columns.Add($column) } }
        $objects = foreach ($values in $Rows) { $row = [ordered]@{}; foreach ($column in $columns) { $row[$column] = $(if ($values.ContainsKey($column)) { $values[$column] } else { '' }) }; [pscustomobject]$row }
        if (@($objects).Count) { @($objects) | Export-Csv -LiteralPath (Join-Path $last $Name) -NoTypeInformation -Encoding utf8NoBOM }
        else { [IO.File]::WriteAllText((Join-Path $last $Name), (($columns | ForEach-Object { '"' + $_ + '"' }) -join ',') + [Environment]::NewLine) }
    }
    Write-Source 'Intune_Devices_Inventory.csv' @(@{TenantKey = 'synthetic-test'; 'Device ID' = 'i-1'; 'Device name' = 'PC-1'; 'Azure AD Device ID' = 'e-1'; OS = 'Windows'; 'OS version' = '10.0.19045.1'; PhysicalMemoryGB = '8'}, @{TenantKey = 'synthetic-test'; 'Device ID' = 'i-2'; 'Device name' = 'PC-2'; 'Azure AD Device ID' = 'e-2'; OS = 'Windows'; 'OS version' = '10.0.26100.1'; PhysicalMemoryGB = '4'})
    Write-Source 'M365_Users_Active.csv' @()
    Write-Source 'M365_Entra_Devices.csv' @(@{TenantKey = 'synthetic-test'; DeviceId = 'e-1'; IsPending = 'False'}, @{TenantKey = 'synthetic-test'; DeviceId = 'e-3'; IsPending = 'True'})
    Write-Source 'AD_Computers_AllDomains.csv' @(@{TenantKey = 'synthetic-test'; ObjectGUID = 'g-1'; ObjectGuidNormalized = 'e-1'; IntuneDeviceId = 'i-1'; Name = 'PC-1$'; Enabled = 'True'; OperatingSystem = 'Windows 10 Enterprise'; LastRebootDateTime = '2026-10-01 08:00:00'}, @{TenantKey = 'synthetic-test'; ObjectGUID = 'g-3'; ObjectGuidNormalized = 'e-3'; Name = 'PC-3$'; Enabled = 'True'; OperatingSystem = 'Windows 10 Enterprise'})
    foreach ($name in 'Intune_Devices_UpgradeEligibility.csv', 'Intune_WindowsUpdate_Status.csv', 'Intune_EndpointAnalytics_DevicePerformance.csv', 'Intune_EndpointAnalytics_StartupDevices.csv', 'M365_Entra_Devices_HardwareIdConflicts.csv', 'M365_Entra_Devices_RemovalCandidates.csv') { Write-Source $name @() }
    Write-Source 'Intune_Devices_LocalSystem.csv' @(@{TenantKey = 'synthetic-test'; AzureADDeviceId = 'e-2'; DeviceId = 'i-2'; FirmwareType = 'Uefi'; SecureBootStatus = 'Enabled'; 'Last Reboot Date' = '2026-10-20 08:00:00'; PlatformScriptLastUpdate = '2026-10-24T08:00:00Z'})
    $generator = Join-Path $PSScriptRoot 'New-DeviceInventoryEvidence.ps1'
    $outputs = @{}
    foreach ($mode in 'without', 'with') {
        $out = Join-Path $root $mode; New-Item -ItemType Directory -Path $out -Force | Out-Null
        $arguments = @{DataRoot = $root; OutputPath = (Join-Path $out 'DeviceInventoryEvidence.csv'); SignalsOutputPath = (Join-Path $out 'DeviceRiskEvidenceSummary.csv'); DirectorySummaryOutputPath = (Join-Path $out 'DeviceDirectorySummary.csv'); SystemEvidenceOutputPath = (Join-Path $out 'DeviceSystemEvidence.csv'); AsOfUtc = $reference}
        if ($mode -eq 'with') { $arguments.ReadinessGapOutputPath = Join-Path $out 'DeviceReadinessGapEvidence.csv' }
        & $generator @arguments | Out-Null
        $outputs[$mode] = $out
    }
    foreach ($name in 'DeviceInventoryEvidence.csv', 'DeviceRiskEvidenceSummary.csv', 'DeviceDirectorySummary.csv', 'DeviceSystemEvidence.csv') {
        Check ((Get-FileHash (Join-Path $outputs.without $name)).Hash -eq (Get-FileHash (Join-Path $outputs.with $name)).Hash) "Existing output changed: $name"
    }
    Check (-not (Test-Path -LiteralPath (Join-Path $outputs.without 'DeviceReadinessGapEvidence.csv'))) 'Gap facts written without being requested'
    $gap = @(Import-Csv -LiteralPath (Join-Path $outputs.with 'DeviceReadinessGapEvidence.csv'))
    Check ($gap.Count -eq 3 -and @($gap | Where-Object 'Inventory Scope' -EQ 'AD computer only').Count -eq 1) 'Generator gap facts scope changed'
    Check ((@($gap | Where-Object 'Device Name' -EQ 'PC-3')[0]).'Entra Registration Pending' -eq 'True') 'AD-only pending registration lost'
    Check ((@($gap | Where-Object 'Device Name' -EQ 'PC-2')[0]).'Last Reboot Source' -eq 'Intune local system') 'Generator LocalSystem fallback lost'
    Check ((Get-Content -LiteralPath (Join-Path $outputs.with 'DeviceReadinessGapEvidence.csv') -TotalCount 1) -ceq (((Get-DeviceReadinessGapColumns) | ForEach-Object { '"' + $_ + '"' }) -join ',')) 'Generator header differs from the declared columns'
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host "PASS: $script:checks device readiness gap checks (reboot time basis, AD priority, AD-only scope, byte-identical existing outputs). Synthetic fixtures only."
