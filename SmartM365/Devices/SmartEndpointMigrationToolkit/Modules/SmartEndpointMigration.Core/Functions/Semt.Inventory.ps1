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
