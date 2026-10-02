<#
.SYNOPSIS
Offline regression tests for SmartInventory sources required by CMDB.
.DESCRIPTION
Loads only selected AST function definitions. Never starts a collector, tenant
authentication, notification, upload, or production-file write.
.VERSION
1.0.2
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$sourceRoot = Split-Path $PSScriptRoot -Parent
$tests = [System.Collections.Generic.List[object]]::new()
function Assert-True { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message } }
function Assert-Throws {
    param([scriptblock]$Body, [string]$MessagePattern)
    $caught = $null
    try { & $Body | Out-Null } catch { $caught = $_ }
    if ($null -eq $caught -or $caught.Exception.Message -notlike $MessagePattern) { throw "Expected failure: $MessagePattern" }
}
function Test-Case {
    param([string]$Name, [scriptblock]$Body)
    try { & $Body; $tests.Add([pscustomobject]@{ Name = $Name; Passed = $true; Error = '' }) }
    catch { $tests.Add([pscustomobject]@{ Name = $Name; Passed = $false; Error = $_.Exception.Message }) }
}
function Read-TestAst {
    param([string]$RelativePath)
    $parseTokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $sourceRoot $RelativePath), [ref]$parseTokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
    return $ast
}
function Import-TestFunction {
    param($Ast, [string]$Name)
    $node = $Ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $node) { throw "Missing function: $Name" }
    Set-Item -Path "Function:script:$Name" -Value ([scriptblock]::Create($node.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
}
$readinessAst = Read-TestAst 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-UpgradeEligibility.ps1'
$licenseAst = Read-TestAst 'SmartInventory/M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1'
Import-TestFunction $readinessAst 'Select-UpgradeEligibilityIdentityRows'
Import-TestFunction $licenseAst 'ConvertTo-LicensesAssignmentPathRows'
Import-TestFunction $licenseAst 'ConvertTo-LicensesTenantSkuEvidence'
$coreAst = Read-TestAst 'Modules/SmartM365.Core/SmartM365.Core.psm1'
foreach ($functionName in @('Get-SmartM365CsvValidationBaseName','Get-SmartM365CsvValidationRule','Assert-SmartM365CsvDataCompleteness','Add-SmartM365CsvValidationRule','Initialize-SmartM365DefaultCsvValidationRules')) {
    Import-TestFunction $coreAst $functionName
}
function WriteLog { param([string]$Message, [string]$Level) }
$global:SmartM365CsvValidationRules = @{}
$global:SmartM365RequireCsvValidationRules = $true
Initialize-SmartM365DefaultCsvValidationRules

$autopilotAst = Read-TestAst 'SmartInventory/M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1'
Import-TestFunction $autopilotAst 'Get-InventoryColumns'
Import-TestFunction $autopilotAst 'Select-AutopilotIdentityRows'
function New-TestAutopilotRow {
    param([string]$Id='native-1',[string]$Serial='same')
    $row = [ordered]@{}
    foreach ($column in Get-InventoryColumns) { $row[$column] = '' }
    $row['Autopilot ID'] = $Id; $row['Serial number'] = $Serial
    [pscustomobject]$row
}
Test-Case 'Autopilot keeps separate native IDs with the same serial' {
    Assert-True (@(Select-AutopilotIdentityRows @((New-TestAutopilotRow 'native-1'),(New-TestAutopilotRow 'native-2'))).Count -eq 2) 'Native device was lost by serial.'
}
Test-Case 'Autopilot collapses only identical repeated native IDs' {
    $row = New-TestAutopilotRow
    Assert-True (@(Select-AutopilotIdentityRows @($row,$row)).Count -eq 1) 'Exact repeat retained.'
}
Test-Case 'Autopilot conflicting native IDs fail rather than pick the latest' {
    Assert-Throws { Select-AutopilotIdentityRows @((New-TestAutopilotRow 'native-1' 'serial-1'),(New-TestAutopilotRow 'native-1' 'serial-2')) } '*Conflicting Autopilot*'
}
Test-Case 'Autopilot missing native ID is not replaced by serial' {
    Assert-Throws { Select-AutopilotIdentityRows @((New-TestAutopilotRow '')) } '*lacks its native*'
}
Test-Case 'Autopilot blank serial is retained with its native ID' {
    $row = New-TestAutopilotRow 'native-1' ''
    Assert-SmartM365CsvDataCompleteness -BaseFileName 'Intune_Autopilot_Devices' -Data @($row)
    Assert-True (@(Select-AutopilotIdentityRows @($row)).Count -eq 1) 'Blank serial lost.'
}
Test-Case 'Autopilot zero population keeps all native headers' {
    Assert-True (@(Select-AutopilotIdentityRows @()).Count -eq 0) 'Empty population was fabricated.'
    Assert-SmartM365CsvDataCompleteness -BaseFileName 'Intune_Autopilot_Devices' -Data @() -Columns (Get-InventoryColumns)
}
Test-Case 'Autopilot duplicate native identity fails the export gate' {
    $row = New-TestAutopilotRow
    Assert-Throws { Assert-SmartM365CsvDataCompleteness -BaseFileName 'Intune_Autopilot_Devices' -Data @($row,$row) } '*duplicate immutable*'
}
Test-Case 'Verified domains empty success requires its actual complete schema' {
    Assert-SmartM365CsvDataCompleteness -BaseFileName 'M365_Entra_VerifiedDomains' -Data @() -Columns @('Id','IsVerified','IsDefault','IsInitial','AuthenticationType','SupportedServices','AvailabilityStatus')
    Assert-Throws { Assert-SmartM365CsvDataCompleteness -BaseFileName 'M365_Entra_VerifiedDomains' -Data @() -Columns @('TenantKey') } '*missing required column*'
}

Test-Case 'Different immutable device IDs sharing a name are both retained' {
    $rows = @([pscustomobject]@{ GraphId = 'device-1'; DeviceName = 'PC-1'; NormalizedDeviceName = 'pc-1'; UpgradeEligibility = 'capable' },
              [pscustomobject]@{ GraphId = 'device-2'; DeviceName = 'PC-1'; NormalizedDeviceName = 'pc-1'; UpgradeEligibility = 'notCapable' })
    Assert-True (@(Select-UpgradeEligibilityIdentityRows $rows).Count -eq 2) 'A device was dropped by display name.'
}
Test-Case 'Only identical repeated readiness IDs are collapsed' {
    $row = [pscustomobject]@{ GraphId = 'device-1'; DeviceName = 'PC-1'; NormalizedDeviceName = 'pc-1'; UpgradeEligibility = 'capable' }
    Assert-True (@(Select-UpgradeEligibilityIdentityRows @($row, $row)).Count -eq 1) 'Repeated ID was not collapsed.'
}
Test-Case 'Conflicting readiness IDs block publication' {
    $rows = @([pscustomobject]@{ GraphId = 'device-1'; UpgradeEligibility = 'capable' }, [pscustomobject]@{ GraphId = 'device-1'; UpgradeEligibility = 'notCapable' })
    Assert-Throws { Select-UpgradeEligibilityIdentityRows $rows } '*Conflicting readiness*'
}
Test-Case 'Missing readiness ID is not replaced with a name' {
    Assert-Throws { Select-UpgradeEligibilityIdentityRows @([pscustomobject]@{ GraphId = ''; DeviceName = 'PC-1' }) } '*lacks GraphId*'
}
Test-Case 'Error and disabled license paths are exported with group IDs' {
    $user = [pscustomobject]@{ Id = 'user-1'; LicenseAssignmentStates = @(
        [pscustomobject]@{ SkuId = 'sku-1'; AssignedByGroup = $null; State = 'Disabled'; Error = $null; DisabledPlans = @('plan-2','plan-1'); LastUpdatedDateTime = $null },
        [pscustomobject]@{ SkuId = 'sku-1'; AssignedByGroup = 'group-1'; State = 'Error'; Error = 'CountViolation'; DisabledPlans = @(); LastUpdatedDateTime = '2026-10-01T12:00:00+02:00' }) }
    $rows = @(ConvertTo-LicensesAssignmentPathRows $user)
    Assert-True ($rows.Count -eq 2 -and $rows[0].AssignmentState -eq 'Disabled') 'Ineffective assignment lost.'
    Assert-True ($rows[1].AssignedByGroupId -eq 'group-1' -and $rows[1].AssignmentError -eq 'CountViolation') 'Immutable group/error evidence lost.'
    Assert-True ($rows[1].LastUpdatedDateTime -eq '2026-10-01T10:00:00.0000000+00:00') 'Assignment date is not UTC qualified.'
    Assert-True ($rows[0].DisabledPlanIds -eq 'plan-1;plan-2') 'Disabled plans are not deterministic.'
}
Test-Case 'Conflicting paths cannot silently become one assignment' {
    $state = [pscustomobject]@{ SkuId = 'sku-1'; AssignedByGroup = 'group-1'; State = 'Active'; Error = $null; DisabledPlans = @(); LastUpdatedDateTime = $null }
    $other = [pscustomobject]@{ SkuId = 'sku-1'; AssignedByGroup = 'group-1'; State = 'Error'; Error = 'CountViolation'; DisabledPlans = @(); LastUpdatedDateTime = $null }
    Assert-Throws { ConvertTo-LicensesAssignmentPathRows ([pscustomobject]@{ Id = 'user-1'; LicenseAssignmentStates = @($state, $other) }) } '*Conflicting license*'
}
Test-Case 'Unknown capacity remains null rather than zero' {
    $sku = [pscustomobject]@{ SkuPartNumber = 'TEST'; PrepaidUnits = $null; ConsumedUnits = $null; CapabilityStatus = $null; AppliesTo = $null; SubscriptionIds = @() }
    $row = ConvertTo-LicensesTenantSkuEvidence $sku
    Assert-True ($null -eq $row.PrepaidEnabled -and $null -eq $row.ConsumedUnits) 'Missing capacity was fabricated as zero.'
}
Test-Case 'Real zero capacity is preserved' {
    $sku = [pscustomobject]@{ SkuPartNumber = 'TEST'; PrepaidUnits = [pscustomobject]@{ Enabled = 0; Warning = 2; Suspended = 1 }; ConsumedUnits = 0; CapabilityStatus = 'Suspended'; AppliesTo = 'User'; SubscriptionIds = @() }
    $row = ConvertTo-LicensesTenantSkuEvidence $sku
    Assert-True ($row.PrepaidEnabled -eq 0 -and $row.ConsumedUnits -eq 0 -and $row.PrepaidWarning -eq 2 -and $null -eq $row.PrepaidLockedOut) 'Capacity states were conflated.'
}
Test-Case 'AD user collection does not filter on UPN presence' {
    $ast = Read-TestAst 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1'
    $commands = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-ADUser' }, $true))
    Assert-True (@($commands | Where-Object { $_.Extent.Text -match '\(userPrincipalName=\*\)' }).Count -eq 0) 'UPN-less AD users remain excluded.'
    Assert-True (@($commands | Where-Object { $_.Extent.Text -match 'primaryGroupID' -and $_.Extent.Text -match 'ObjectGUID' -and $_.Extent.Text -match 'ObjectSID' }).Count -gt 0) 'AD immutable/group identity missing.'
}
Test-Case 'License path extraction precedes the effective-detail empty guard' {
    $text = $licenseAst.Extent.Text
    Assert-True ($text.IndexOf('foreach ($assignmentPath') -lt $text.IndexOf('if (-not $licenseDetails')) 'Error-only paths are hidden by effective-license filtering.'
    Assert-True ($text.IndexOf('if ($userResolutionFailures -gt 0)') -lt $text.IndexOf("`$currentOperation = 'Export immutable license assignment paths'")) 'Partial licensing data could be published as complete.'
}
Test-Case 'Full Entra inventory is exported before diagnostic filters' {
    $ast = Read-TestAst 'SmartInventory/M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1'
    Assert-True ($ast.Extent.Text.IndexOf("-BaseFileName 'M365_EntraDevices_All'") -lt $ast.Extent.Text.IndexOf('Applying OperatingSystem filter')) 'CMDB source is still diagnostic-filtered.'
}
Test-Case 'Full Intune source retrieves every platform while legacy stays Windows' {
    $ast = Read-TestAst 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1'
    $commands = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-MgDeviceManagementManagedDevice' -and $n.Extent.Text -match '-All' }, $true))
    Assert-True (@($commands | Where-Object { $_.Extent.Text -match '-Filter' }).Count -eq 0) 'All-platform source is filtered.'
    Assert-True ($ast.Extent.Text.Contains('Where-Object { (Get-SafeProperty $_ ''OperatingSystem'') -eq ''Windows'' }')) 'Legacy Windows view not retained.'
}
Test-Case 'Full source exports reject duplicate immutable keys' {
    $row = [pscustomobject]@{ ObjectId = 'device-1'; DeviceId = ''; OnPremisesSecurityIdentifier = ''; OperatingSystem = 'Windows' }
    Assert-Throws { Assert-SmartM365CsvDataCompleteness -BaseFileName 'M365_EntraDevices_All' -Data @($row, $row) } '*duplicate immutable*'
}
Test-Case 'A single missing immutable key blocks even a small source' {
    $row = [pscustomobject]@{ ObjectId = ''; DeviceId = ''; OnPremisesSecurityIdentifier = ''; OperatingSystem = 'Windows' }
    Assert-Throws { Assert-SmartM365CsvDataCompleteness -BaseFileName 'M365_EntraDevices_All' -Data @($row) } '*DATA-LAST publication is blocked*'
}
Test-Case 'Direct assignments may have an empty group ID' {
    $row = [pscustomobject]@{ UserId = 'user-1'; SkuId = 'sku-1'; AssignmentRoute = 'Direct'; AssignedByGroupId = ''; AssignmentState = ''; AssignmentError = ''; DisabledPlanIds = ''; LastUpdatedDateTime = '' }
    Assert-SmartM365CsvDataCompleteness -BaseFileName 'M365_Licenses_AssignmentPaths' -Data @($row)
}
Test-Case 'Successful empty native inventories retain their explicit schema' {
    Assert-SmartM365CsvDataCompleteness -BaseFileName 'M365_EntraGroups_All' -Data @() -Columns @('GroupId','DisplayName','GroupTypes','SecurityEnabled')
}
Test-Case 'SMTP duplicates do not require the optional AD UPN' {
    $row = [pscustomobject]@{ SmtpAddress = 'shared@example.invalid'; SamAccountName = 'shared'; DistinguishedName = 'CN=shared,DC=example,DC=invalid'; UserPrincipalName = '' }
    Assert-SmartM365CsvDataCompleteness -BaseFileName 'AD_Users_DuplicateSMTP' -Data @($row)
}
Test-Case 'Legacy Windows inventory also retains duplicate display names' {
    $ast = Read-TestAst 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1'
    Assert-True (-not $ast.Extent.Text.Contains('$seenNames.Add')) 'Legacy Windows inventory still drops same-name devices.'
    $row1 = [pscustomobject]@{ 'Device ID' = 'id-1'; 'Device name' = 'PC-1'; 'Azure AD Device ID' = '' }
    $row2 = [pscustomobject]@{ 'Device ID' = 'id-2'; 'Device name' = 'PC-1'; 'Azure AD Device ID' = '' }
    Assert-SmartM365CsvDataCompleteness -BaseFileName 'Intune_Devices_Inventory' -Data @($row1, $row2)
    Assert-Throws { Assert-SmartM365CsvDataCompleteness -BaseFileName 'Intune_Devices_Inventory' -Data @($row1, $row1) } '*duplicate immutable*'
}
Test-Case 'A legitimate empty Windows scope does not start RAM requests' {
    $ast = Read-TestAst 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-DeviceHardware.ps1'
    Import-TestFunction $ast 'Get-SmartM365ManagedDeviceHardware'
    $map = Get-SmartM365ManagedDeviceHardware -ManagedDeviceIds @()
    Assert-True ($map.Count -eq 0) 'Empty Windows scope was not supported.'
    Assert-SmartM365CsvDataCompleteness -BaseFileName 'Intune_Devices_Inventory' -Data @() -Columns @('Device ID','Device name','Azure AD Device ID')
}
$tests | Format-Table -AutoSize
if (@($tests | Where-Object { -not $_.Passed }).Count) { throw 'CMDB source completeness regression failed.' }
[pscustomobject]@{ Status = 'Passed'; TestCount = $tests.Count; ProductionActions = 0 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcYSO1sJKELMEq
# JUgMb2rfUB9W7dcwNXxm9INjiIwaEaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJ0S5zjZIdjFZxp8IkuFOuHRMkZXa6UxKbZJPPSaQLiUMA0GCSqG
# SIb3DQEBAQUABIIBgJ/0EpnzD7AEcuJKzXhLny6cdugtnqIBmkuUXjR1Zghchry4
# w8w6/n9aBTQIsydpjlabphmxxwsR5sZAOsiSSna/DIbmr4uZIFRufqRiNXfRc22G
# +LRih0LoR9MPFZjO+8vFu6aWXydqucqBPnDNrtMq9sr+ena9BF8yJNc2Gx1xG4a6
# 60ZKByOYqRsdZSgk6pM+5Svbj6RNnfl59m7ovw67FT/zPjt8oo5KN/bUIvvCtZPa
# grbFMx+SPDGTsJTLxPZ5XWt/IOu5suzfy/Z1476T0fFrTbgRHHxo6FVJ+ftsscNa
# hWkUxRsUh2t64oV8qoVu4Rhr/Fzt6m9PZKjiHjGi0m/nmF82N0ZJIDz69wQW22V0
# P1mMW0S0MUyXfwbGwsGKLdl9Cm0PdQ9NJ//eCLv9vefb8M/9CUtsZGTbtmhX1y/Q
# x4gyxzjOOGPRP0OZFZwn6XfFYy2TRUe56Y8/0hbSDrgHBSDyVcH6Hl9BpZ1Q/4sv
# KIL+xUxecjSl9ZQJo6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDIxOTM4
# NDVaMC8GCSqGSIb3DQEJBDEiBCD5zFSzk/DuYxak3b/f+0be4F+NTnseIpfpHkWf
# KOPPJzANBgkqhkiG9w0BAQEFAASCAgCTvjy8nFITVcbyyKfNUUIUrCoB4DKepMIH
# bOeGuRsfU+rJIu0kKMwxO3p2eax0lGWOZ4W2E0cA/4FGKaZW+DFSdkyb6VHc+XGU
# lxmZ135kxHZE0w+tzySnUkvIYLo7ku4qB3rbr0bfQXRjGdfBh8DGdy/+BcP5wpfM
# 285JXLg0M+eG9FsdHAoE1uq4Pvu+Yn7aQu6L8Fu+BubMFVZuXQqqwgQ5P74hnw2N
# Tp7wirTs7Kj9MAiaLj/PtM8DH6ZTI5eruBdQKoG0qYguJ6horZww8qyEhT1cLilO
# cstPWv8BozA4KJo4zyVPos2d+/Bcwq2Xes0eNOGFPSFAPsGk+SndMdD4cqUWB3qO
# pkPFEr4YHPxZRYyrPVJp2rlHF6Fh+wZe1eI/2238yoHW8lpYJKAYcH+BGa6h5OY8
# GK3Ww1ECU94NzqxI0L0zTDdUm8epJIPS9YP4DE4dU+DcUUUSXVMcug3STlHtU/03
# ReSg/uU1ThKTAGTltxjZJsW+Igl515Po5jF/tlt/FteIFYT2cie9GzObP3zwoaZY
# ZMho8uEmAeBRQ+N2WgDPDbQaTxGLn7iNuo63sFaMRfxUcd6PfZy7IMHLMja57tQ7
# f9jPMecDTpd04jWXOsF36VWxeY8tF+Ns4rznYF0ZH9/UDtq+ccoE0Xn6upoofvk1
# 0k462lWrVQ==
# SIG # End signature block
