<#
.SYNOPSIS
    AD, Intune and Entra inventory caches, scoped refreshes and result enrichment.

.DESCRIPTION
    Shared inventories live in <Package>\Inventory and are reused READ-ONLY while they are
    fresh (and, for Intune, when their delegated provenance matches). A refresh scoped to a
    Wave is always written inside the Run folder and never replaces a shared inventory.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

$script:SemtInventoryKinds = @('AD','Intune','Entra')
$script:SemtInventoryFields = @{
    AD = @('ADInventoryPresent','ADDomain','ADEnabled','ADDNSHostName','ADDistinguishedName','ADOperatingSystem','ADOperatingSystemVersion','ADLastLogonTimestampUtc','ADInventoryCsv')
    Intune = @('IntuneInventoryPresent','IntuneDeviceName','IntuneManagedDeviceName','IntuneManagedDeviceId','IntuneAzureADDeviceId','IntuneOperatingSystem','IntuneOSVersion','IntuneLastSyncDateTime','IntuneUserPrincipalName','IntuneComplianceState','IntuneManagementState','IntuneInventoryCsv')
    Entra = @('EntraInventoryPresent','EntraRegisteredState','EntraAlternativeSecurityIdCount','EntraPendingReason','EntraRegistrationDateTime','EntraTrustType','EntraDeviceId','EntraObjectId','EntraInventoryCsv')
}
$script:SemtInventorySourceColumns = @{
    AD = [ordered]@{ ADDomain = 'ADDomain'; ADEnabled = 'Enabled'; ADDNSHostName = 'DNSHostName'; ADDistinguishedName = 'DistinguishedName'; ADOperatingSystem = 'OperatingSystem'; ADOperatingSystemVersion = 'OperatingSystemVersion'; ADLastLogonTimestampUtc = 'LastLogonTimestampUtc' }
    Intune = [ordered]@{ IntuneDeviceName = 'DeviceName'; IntuneManagedDeviceName = 'ManagedDeviceName'; IntuneManagedDeviceId = 'IntuneManagedDeviceId'; IntuneAzureADDeviceId = 'AzureADDeviceId'; IntuneOperatingSystem = 'OperatingSystem'; IntuneOSVersion = 'OSVersion'; IntuneLastSyncDateTime = 'LastSyncDateTime'; IntuneUserPrincipalName = 'UserPrincipalName'; IntuneComplianceState = 'ComplianceState'; IntuneManagementState = 'ManagementState' }
    Entra = [ordered]@{ EntraRegisteredState = 'EntraRegisteredState'; EntraAlternativeSecurityIdCount = 'AlternativeSecurityIdCount'; EntraPendingReason = 'EntraPendingReason'; EntraRegistrationDateTime = 'RegistrationDateTime'; EntraTrustType = 'TrustType'; EntraDeviceId = 'DeviceId'; EntraObjectId = 'EntraObjectId' }
}

function Get-SemtInventoryFieldNames {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateSet('AD','Intune','Entra')][string]$Kind)
    return @($script:SemtInventoryFields[$Kind])
}

function Find-SemtGraphAuthenticationModule {
    <#
    Newest Microsoft.Graph.Authentication visible to this process. When PSModulePath misses the folder
    that Install-Module -Scope CurrentUser wrote to (redirected Documents, PSModulePath set by policy),
    the PowerShellGet install record gives that folder; it is added to PSModulePath for this process only.
    #>
    [CmdletBinding()]
    param()

    $name = 'Microsoft.Graph.Authentication'
    $module = Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
    if ($module) { return $module }
    if (-not (Get-Command -Name Get-InstalledModule -ErrorAction SilentlyContinue)) { return $null }
    $installed = @(Get-InstalledModule -Name $name -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending)
    foreach ($item in $installed) {
        $location = [string]$item.InstalledLocation
        if ([string]::IsNullOrWhiteSpace($location) -or -not (Test-Path -LiteralPath (Join-Path $location ($name + '.psd1')) -PathType Leaf)) { continue }
        $modulesRoot = Split-Path -Path (Split-Path -Path $location -Parent) -Parent
        if (@($env:PSModulePath -split ';' | ForEach-Object { $_.TrimEnd('\') }) -notcontains $modulesRoot.TrimEnd('\')) { $env:PSModulePath = $modulesRoot + ';' + $env:PSModulePath }
        $module = Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
        if ($module) { return $module }
    }
    return $null
}

function Get-SemtGraphModuleDiagnostic {
    [CmdletBinding()]
    param()

    $installed = @(if (Get-Command -Name Get-InstalledModule -ErrorAction SilentlyContinue) { Get-InstalledModule -Name 'Microsoft.Graph.Authentication' -AllVersions -ErrorAction SilentlyContinue | ForEach-Object { '{0} in {1}' -f $_.Version,$_.InstalledLocation } })
    return ('Documents folder: {0}. PSModulePath: {1}. PowerShellGet install record: {2}. Install it for all users from an elevated Windows PowerShell: Install-Module Microsoft.Graph.Authentication -Scope AllUsers' -f [Environment]::GetFolderPath('MyDocuments'),$env:PSModulePath,$(if ($installed.Count -gt 0) { $installed -join '; ' } else { 'none' }))
}

function Get-SemtSharedInventoryPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][ValidateSet('AD','Intune','Entra')][string]$Kind
    )

    return (Join-Path -Path (Join-Path -Path $PackageRoot -ChildPath 'Inventory') -ChildPath ("Devices{0}.csv" -f $Kind))
}

function Get-SemtInventoryMap {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Path,
        [Parameter(Mandatory = $true)][ValidateSet('AD','Intune','Entra')][string]$Kind,
        [string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $map }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }
    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidates = switch ($Kind) {
            'AD' { @('ComputerName','DNSHostName','Name') }
            'Intune' { @('ComputerName','DeviceName','ManagedDeviceName','Name') }
            'Entra' { @('ComputerName','DisplayName','DeviceName','Name') }
        }
        $first = $rows[0]
        foreach ($candidate in $candidates) {
            if ($first.PSObject.Properties.Name -contains $candidate) { $NameColumn = $candidate; break }
        }
    }
    if ([string]::IsNullOrWhiteSpace($NameColumn)) { throw ("Unable to infer the {0} inventory device name column: {1}" -f $Kind,$Path) }
    $presenceColumn = '{0}InventoryPresent' -f $Kind
    foreach ($row in $rows) {
        $key = Get-SemtComputerKey -ComputerName ([string]$row.$NameColumn)
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if ($row.PSObject.Properties.Name -contains $presenceColumn -and -not (Test-SemtBooleanLikeTrue -Value $row.$presenceColumn)) { continue }
        if (-not $map.ContainsKey($key)) { $map[$key] = $row; continue }
        if ($Kind -eq 'Intune') {
            $currentSync = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $map[$key] -Name 'LastSyncDateTime' -Default '')
            $candidateSync = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $row -Name 'LastSyncDateTime' -Default '')
            if ($null -ne $candidateSync -and ($null -eq $currentSync -or $candidateSync -gt $currentSync)) { $map[$key] = $row }
        }
    }
    return $map
}

function Test-SemtSharedInventoryUsable {
    <#
    A shared inventory is reused only when it is fresh and, for Intune, when its delegated
    tenant provenance and AllManagedDevices scope match. The result explains the decision.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Path,
        [Parameter(Mandatory = $true)][ValidateSet('AD','Intune','Entra')][string]$Kind,
        [Parameter(Mandatory = $true)][int]$MaxAgeHours,
        [AllowNull()][string]$TenantId
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ Usable = $false; Reason = 'Missing'; AgeHours = $null; Provenance = '' }
    }
    $item = Get-Item -LiteralPath $Path
    $ageHours = [math]::Round(((Get-Date) - $item.LastWriteTime).TotalHours, 2)
    if ($ageHours -gt $MaxAgeHours) {
        return [pscustomobject]@{ Usable = $false; Reason = ('Stale:{0}h>{1}h' -f $ageHours,$MaxAgeHours); AgeHours = $ageHours; Provenance = '' }
    }
    if ($Kind -ne 'Intune') {
        return [pscustomobject]@{ Usable = $true; Reason = 'Fresh'; AgeHours = $ageHours; Provenance = '' }
    }
    $first = Import-Csv -LiteralPath $Path | Select-Object -First 1
    $rowTenant = [string](Get-SemtPropertyValue -InputObject $first -Name 'InventoryTenantId' -Default '')
    $rowAuth = [string](Get-SemtPropertyValue -InputObject $first -Name 'InventoryAuthenticationMode' -Default '')
    $rowScope = [string](Get-SemtPropertyValue -InputObject $first -Name 'InventoryScope' -Default '')
    $tenantOk = (-not [string]::IsNullOrWhiteSpace($rowTenant) -and ([string]::IsNullOrWhiteSpace($TenantId) -or $rowTenant -ieq $TenantId))
    $authOk = ($rowAuth -in @('DelegatedInteractive','DelegatedExistingSession'))
    $scopeOk = ($rowScope -eq 'AllManagedDevices')
    $provenance = 'TenantId={0};Auth={1};Scope={2}' -f $rowTenant,$rowAuth,$rowScope
    if ($tenantOk -and $authOk -and $scopeOk) {
        return [pscustomobject]@{ Usable = $true; Reason = 'FreshVerifiedProvenance'; AgeHours = $ageHours; Provenance = $provenance }
    }
    return [pscustomobject]@{ Usable = $false; Reason = 'ProvenanceRejected'; AgeHours = $ageHours; Provenance = $provenance }
}

function Invoke-SemtInventoryExport {
    <#
    Runs an exporter in Windows PowerShell 5.1 (ActiveDirectory and Graph modules) and loads
    the produced CSV. A scoped export must target a Run folder, never the shared inventory.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('AD','Intune','Entra')][string]$Kind,
        [Parameter(Mandatory = $true)][string]$ExportScriptPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [AllowNull()][string]$ComputerListPath,
        [AllowNull()][string]$Domain,
        [AllowNull()][string]$TenantId,
        [ValidateRange(1, 999)][int]$PageSize = 999,
        [AllowNull()][string]$SharedInventoryPath
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath -PathType Leaf)) { throw "Inventory exporter not found: $ExportScriptPath" }
        if (-not [string]::IsNullOrWhiteSpace($SharedInventoryPath) -and -not [string]::IsNullOrWhiteSpace($ComputerListPath)) {
            if ([System.IO.Path]::GetFullPath($OutputPath) -ieq [System.IO.Path]::GetFullPath($SharedInventoryPath)) {
                throw "SEMT_SCOPED_INVENTORY_TARGETS_SHARED_CACHE: a Wave-scoped export cannot overwrite the shared inventory $SharedInventoryPath."
            }
        }
        New-SemtDirectory -Path (Split-Path -Path $OutputPath -Parent)
        New-SemtDirectory -Path (Split-Path -Path $LogPath -Parent)
        $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',$ExportScriptPath,'-OutputPath',$OutputPath,'-ForceRefresh')
        if (-not [string]::IsNullOrWhiteSpace($ComputerListPath)) { $arguments += @('-ComputerListPath',$ComputerListPath) }
        if ($Kind -eq 'AD' -and -not [string]::IsNullOrWhiteSpace($Domain)) { $arguments += @('-Domain',$Domain) }
        if ($Kind -ne 'AD') {
            $arguments += @('-PageSize',[string]$PageSize)
            if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $arguments += @('-TenantId',$TenantId) }
        }
        $powershell = Resolve-SemtWindowsPowerShellPath
        $output = & $powershell @arguments 2>&1
        $exitCode = $LASTEXITCODE
        $output | ForEach-Object { '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_ } | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force
        if ($exitCode -ne 0) { throw ("{0} inventory exporter exited with code {1}. Log={2}" -f $Kind,$exitCode,$LogPath) }
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) { throw ("{0} inventory CSV was not created: {1}" -f $Kind,$OutputPath) }
        $map = Get-SemtInventoryMap -Path $OutputPath -Kind $Kind -NameColumn 'ComputerName'
        return [pscustomobject]@{ Success = $true; Kind = $Kind; CsvPath = $OutputPath; LogPath = $LogPath; Map = $map; Error = '' }
    }
    catch {
        return [pscustomobject]@{ Success = $false; Kind = $Kind; CsvPath = $OutputPath; LogPath = $LogPath; Map = @{}; Error = $_.Exception.Message }
    }
}

function Merge-SemtScopedInventoryMap {
    [CmdletBinding()]
    param(
        [AllowNull()][hashtable]$ExistingMap,
        [AllowNull()][hashtable]$RefreshedMap,
        [AllowEmptyCollection()][string[]]$ScopedComputers = @()
    )

    if ($null -eq $ExistingMap) { $ExistingMap = @{} }
    foreach ($computer in @($ScopedComputers)) {
        $key = Get-SemtComputerKey -ComputerName $computer
        if (-not [string]::IsNullOrWhiteSpace($key)) { [void]$ExistingMap.Remove($key) }
    }
    if ($null -ne $RefreshedMap) {
        foreach ($entry in $RefreshedMap.GetEnumerator()) { $ExistingMap[[string]$entry.Key] = $entry.Value }
    }
    return $ExistingMap
}

function Set-SemtResultField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Result,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Value
    )

    if ($Result -is [System.Collections.IDictionary]) { $Result[$Name] = $Value }
    else { $Result | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force }
}

function Add-SemtInventoryFieldsToResult {
    <#
    Adds the AD, Intune or Entra columns of a computer to a result. When an inventory is loaded
    but the computer is absent, <Kind>InventoryPresent=False keeps the evidence explicit.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Result,
        [Parameter(Mandatory = $true)][ValidateSet('AD','Intune','Entra')][string]$Kind,
        [AllowNull()][hashtable]$Map,
        [AllowNull()][string]$CsvPath
    )

    if ($null -eq $Map -or $Map.Count -eq 0) { return $Result }
    $computer = [string](Get-SemtPropertyValue -InputObject $Result -Name 'ComputerName' -Default '')
    $key = Get-SemtComputerKey -ComputerName $computer
    $row = if (-not [string]::IsNullOrWhiteSpace($key) -and $Map.ContainsKey($key)) { $Map[$key] } else { $null }
    Set-SemtResultField -Result $Result -Name ('{0}InventoryPresent' -f $Kind) -Value ([bool]($null -ne $row))
    foreach ($entry in $script:SemtInventorySourceColumns[$Kind].GetEnumerator()) {
        $value = if ($row) { [string](Get-SemtPropertyValue -InputObject $row -Name $entry.Value -Default '') } else { '' }
        Set-SemtResultField -Result $Result -Name $entry.Key -Value $value
    }
    Set-SemtResultField -Result $Result -Name ('{0}InventoryCsv' -f $Kind) -Value ([string]$CsvPath)
    return $Result
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDLiSDVWaz3+g4c
# cYESs/DMcJ1hMZ7mCZPfHGEoo/7xE6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPiwC1eOWGd+fGfooqBaCy/NdQnzWjx2id9T/LyWSdrjMA0GCSqG
# SIb3DQEBAQUABIIBgGrjrA0+H7DUQ3vlzKJFBGIUsyJ2Ev3nKm49UiZLNUG3HqgJ
# ObvFds9fxKThOlHOt1qpnBh2yhIplfCL6yBPKz7WPpwQ59nEqo+EHxblisg4/mwA
# ujCpB+2TtHuk/1GVYsyhXThnl7ASSI2QbUmhgeCy0xBLjudGRVfqp+Ga4sxBOz+f
# Ak5O8oxmkwaDdTjG4y9kZNWWNsZ+MJppbiSDwqIc6x9yJNDKEOdIwB6mkBTqTfwX
# 6IFEdr6TQVSlrC44CUEfrIRpdbPxsmV1biotAe9EOt/HXKLUb8oFVIWIs8Z7uU1u
# 2kPqBK1RtkYQ1Xy3KjzhKRYtkBXD/zWNL8PUzsHD+9OxQjAMIHzifzeVDHP0UpOe
# b/ZCZ+wHKYJHMPVaqdxy2c8siCSOkP+Rc55H/wG+fwvpW4nBdEFMxYyyTJQI4/V8
# 1ZrrNN4g9k9MZfpncvcStq5yyo7i4AlQiwlC1QAiSwlPXB5pXDrw2LAYt8I95WPO
# 4oPfQidjt1MVhEqIwqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxNTU4
# NDRaMC8GCSqGSIb3DQEJBDEiBCAAKV7YlcSqnUg4RZOwP/gChMoSCx9khiHw+MAR
# xTcVDDANBgkqhkiG9w0BAQEFAASCAgA7fi0P3bzY+1iE7ZSL9G2+sFvQuE6UDiHa
# 3VKUQR+h5Crrm0Xp89SZ1nDPjtQVFSQ6bv0Gjw3JRE43yZKCETEtD3yKF73WbeZ2
# 85Axnp+rPaakCGWSktOBvIrn3qCzVPQaUaZQCqvZ14xBDmHG4I0WiaEGhNcOOS5r
# mjI89v8VZKFHi0lHG97AxERcOZYw+UdjO+DtiNO7PaVmPY9f6paX7Zsn7RgAsv0/
# M9Bbx3aVyDKkdZNYsV1RW3nUsxVAwT2yN7mOLPx+uO1AGNEViaDz/ObUa0Eizd7f
# tTyRBvQypY89z3iKfABSpAhbssseUy9t0aUVAhJydZDfbwvgSGlOKSKGXWzapPDZ
# 47Ylyrgggg0WST8yK4h1DqA/Alc6nSn9CrjDbKnoASpadtGw791sz4IPdfTFSYw7
# m17oDxWXeHv0Idc3J6T6pCZOFEO0oJc4lmqu1IeXZlkwY/DZjo2zEJatXeOv3j3O
# vKpGMNW+4sxLbqjYs5dT3Td2FgeIRo2/FcoC2sLybdzcQVSnipxxoPTifZYBP/Em
# I0L8yE13/u1hmSR/XjggtzBaTq7vmBX0KYqEkXlSYFhYJfjWBabRjNXr7X/j7tuQ
# ssL/k81ZaVReNRqod60O6ygGbEPf3FSMm79uXTZBFeYd60BxcRocQhVH+RiMIKOS
# xkYjkJNY5Q==
# SIG # End signature block
