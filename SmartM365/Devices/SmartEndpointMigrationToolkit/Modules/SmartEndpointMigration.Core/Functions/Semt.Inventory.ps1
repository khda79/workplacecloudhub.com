<#
.SYNOPSIS
    AD, Intune and Entra inventory caches, scoped refreshes and result enrichment.

.DESCRIPTION
    Shared inventories live in <Inventory> and are reused READ-ONLY while they are
    <Inventory> is Storage.InventoryPath of Config\SmartEndpointMigrationToolkit.local.json.txt (default <Package>\Inventory).
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

function Get-SemtUserLicensePreflightDecision {
    [CmdletBinding()]
    param([AllowNull()]$Evidence,[object[]]$Users)
    $unknown = [pscustomobject]@{ State='Unknown'; Detail='Preflight identity, policy or inventory is incomplete'; Block=$false }
    if ($null -eq $Evidence -or $Evidence.Tenants.Count -ne 1 -or -not $Evidence.Sid) { return $unknown }
    $tenant = [string]$Evidence.Tenants[0]
    $userMatches = @($Users | Where-Object { $_.InventoryTenantId -ieq $tenant -and $_.InventoryScope -eq 'AllUsers' -and $_.InventoryComplete -eq 'True' -and $_.OnPremisesSecurityIdentifier -ieq $Evidence.Sid })
    if ($userMatches.Count -ne 1) { return $unknown }
    if ($null -eq $userMatches[0].PSObject.Properties['AssignedSkuIds'] -or $null -eq $userMatches[0].PSObject.Properties['CollectedUtc'] -or $null -eq $userMatches[0].AssignedSkuIds) { return $unknown }
    $collectedText = ([string]$userMatches[0].CollectedUtc).Trim()
    if ($collectedText -notmatch '(?i)(Z|[+-]\d{2}:\d{2})$') { return $unknown }
    $collected = ConvertTo-SemtUtcDateTime -Value $collectedText
    if ($null -eq $collected) { return $unknown }
    $age = ([datetime]::UtcNow - $collected).TotalHours
    if ($age -lt 0 -or $age -ge 24) { return $unknown }
    if ([string]::IsNullOrWhiteSpace([string]$userMatches[0].AssignedSkuIds)) {
        return [pscustomobject]@{ State='NoAssignedLicense'; Detail='Current console user has no assigned license in the fresh same-tenant inventory. Assign an appropriate Intune license before a new Run. Tenant is inferred from registry and must be confirmed.'; Block=$true; UserPrincipalName=[string](Get-SemtPropertyValue -InputObject $userMatches[0] -Name 'UserPrincipalName' -Default ''); CollectedUtc=$collectedText }
    }
    return [pscustomobject]@{ State='HasAssignedLicense'; Detail='A user SKU exists; Intune plan and MDM scope are not confirmed by this preflight.'; Block=$false }
}

function Invoke-SemtUserLicensePreflight {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$ComputerName,[Parameter(Mandatory=$true)][string]$InventoryPath,[int]$TimeoutSeconds=25)
    $unknown = [pscustomobject]@{ State='Unknown'; Detail='User inventory is missing, stale, or preflight could not complete'; Block=$false }
    if (-not (Test-Path -LiteralPath $InventoryPath -PathType Leaf)) { return $unknown }
    $age = ([datetime]::UtcNow - (Get-Item -LiteralPath $InventoryPath).LastWriteTimeUtc).TotalHours
    if ($age -lt 0 -or $age -ge 24) { return $unknown }
    try { $users = @(Import-Csv -LiteralPath $InventoryPath) } catch { return $unknown }
    if ($users.Count -eq 0) { return $unknown }
    $job=Start-Job -ArgumentList $ComputerName -ScriptBlock {
        param($target)
        $s=$null
        try {
            $s=New-CimSession -ComputerName $target -SessionOption (New-CimSessionOption -Protocol Dcom) -ErrorAction Stop
            $cs=Get-CimInstance -CimSession $s -ClassName Win32_ComputerSystem -OperationTimeoutSec 15 -ErrorAction Stop
            $user=[string]$cs.UserName
            if (-not $user -or $user -notmatch '^([^\\]+)\\(.+)$') { return }
            $domain=$Matches[1]; $name=$Matches[2]
            $sids=@()
            foreach ($process in @(Get-CimInstance -CimSession $s -ClassName Win32_Process -Filter "Name='explorer.exe'" -OperationTimeoutSec 15 -ErrorAction Stop)) {
                $owner=Invoke-CimMethod -InputObject $process -MethodName GetOwner -OperationTimeoutSec 15 -ErrorAction Stop
                if ($owner.ReturnValue -ne 0 -or $owner.Domain -ine $domain -or $owner.User -ine $name) { continue }
                $sid=Invoke-CimMethod -InputObject $process -MethodName GetOwnerSid -OperationTimeoutSec 15 -ErrorAction Stop
                if ($sid.ReturnValue -eq 0 -and $sid.Sid) { $sids += [string]$sid.Sid }
            }
            $sids=@($sids | Select-Object -Unique)
            if ($sids.Count -ne 1) { return }
            $reg=Get-CimClass -CimSession $s -Namespace root/default -ClassName StdRegProv -OperationTimeoutSec 15 -ErrorAction Stop
            $hive=[uint32]2147483650
            $tenantKeys=Invoke-CimMethod -CimSession $s -CimClass $reg -MethodName EnumKey -Arguments @{hDefKey=$hive;sSubKeyName='SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo'} -OperationTimeoutSec 15 -ErrorAction Stop
            if ($tenantKeys.ReturnValue -ne 0) { return }
            [pscustomobject]@{ Sid=$sids[0]; Tenants=@($tenantKeys.sNames) }
        } finally { if ($s) { Remove-CimSession $s -ErrorAction SilentlyContinue } }
    }
    try {
        if (-not (Wait-Job $job -Timeout $TimeoutSeconds)) { Stop-Job $job; return $unknown }
        if ($job.State -ne 'Completed') { return $unknown }
        $data=@(Receive-Job $job -ErrorAction SilentlyContinue)
        if ($data.Count -ne 1) { return $unknown }
        return (Get-SemtUserLicensePreflightDecision -Evidence $data[0] -Users $users)
    } catch { return $unknown } finally { Remove-Job $job -Force -ErrorAction SilentlyContinue }
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
        [Parameter(Mandatory = $true)][ValidateSet('AD','Intune','Entra','UserLicenses')][string]$Kind
    )

    # The shared inventory folder is Storage.InventoryPath of the local configuration (default <Package>\Inventory).
    $fileName = if ($Kind -eq 'UserLicenses') { 'UsersLicenses.csv' } else { "Devices{0}.csv" -f $Kind }
    return (Join-Path -Path (Get-SemtStorageRoot -PackageRoot $PackageRoot -Kind Inventory) -ChildPath $fileName)
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
    $elapsedHours = ((Get-Date) - $item.LastWriteTime).TotalHours
    $ageHours = [math]::Round($elapsedHours, 2)
    if ($elapsedHours -gt $MaxAgeHours) {
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
        [AllowNull()][string]$SharedInventoryPath,
        [AllowNull()][string]$DeviceIdListPath
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
            if (-not [string]::IsNullOrWhiteSpace($DeviceIdListPath) -and -not [string]::IsNullOrWhiteSpace($ComputerListPath)) { $arguments += @('-DeviceIdListPath',$DeviceIdListPath) }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCuGgrbcXML5G5H
# 2rOjCSir/5JECUAPVeCgOQjR1fQhLqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIITQeuBG8FuWdIObdtAKKTFNoE0Xt4HJE27ZLXDkbTyhMA0GCSqG
# SIb3DQEBAQUABIIBgBiwLH0U+huW0RM/gyidgx8mEZhNe31NoZexQDgjzA515NUa
# 04FU4nbgujkuZIVQ8zJ2+JKZUkla28DFqzto7z1mgFMM5yJA+1g7Lyn89seKipzo
# XN3M9gAhnkSTCjoc/d4TwXVYPRZoTqFwM/yZTo2KKbd1ydTiLUG4ikldf4YxH4db
# RYh6tnTydZgi570y8IEFJ/rD4gMwBkXZwAJIKKKby0tIc6fWv6dx62vwSkY0rFJ2
# sP9xTNiHULKSvgmPx2HwpxVo9LxrnAJvm81sve/Kypo4+4HG75pzdmBAq7iolRmJ
# NeDMjSsms9hzRxGkqN4iC1NqcKwHEJp53TLT2dt1YbxYcXwHPY50pve4k4lo4wX3
# HHHtZ9KX3bDcvt2fFYZBPFC6t5+2l8ZN8ICDHh0oEL+j3SZgL/Y0jI/+2UZxJuqJ
# pTFTuEFlySmPkpXJzy2fQfa94nwmZdrlATr8Qx/EsIwrvRtwsDw7Onna0z1CTGrP
# k9wHnDqjqJU93ftMGaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDUyMTQ4
# MTVaMC8GCSqGSIb3DQEJBDEiBCC6Y/xd5q5b3IHzQRrfRUPRMCR9pTQufx9I+Qs+
# arPbYTANBgkqhkiG9w0BAQEFAASCAgBcDmwyjF6YH7Vwyf/2rfirl9KI/26GLyCe
# 3O3LC+/qonoAAH9oZNaomUV3K9rs6GtQcO0VvgaRZoJEwhb2XpJ7lzmf42omBYXR
# 3AMZh7gTH+s1asYmCg7FW95pwuhsJdwnf2zl6mUvPGVYeT00RkHAZ1BT6zLjfCPa
# MZgRM3Q8Sa3XakrRDqfdE98ShQpuqJVLQ3vk8OlbeYc2dgfkRpypf2dcluDuiwKL
# La2BDYC5me1Wn+UHvP76LtkxVAkNR2p9fVnR9RhCSkRIVmsAbQX6T4dFpqL07dAM
# 0abiRDDpPT0VJeI+MhRHheG0GiD+8sDjp/5TE4Rz8ibNTi0LRPLif+9P5LkJhav4
# PGVB6aZTNgpz8vFPnp92aNx0pBd8l/70y3HGzD4NlV4U0OSizUtkcE1ZoWjsCIUL
# lVjrQ61j2xPQwDpxrDUGmlnT/DibTlQFBjdekQXDddeoN68cS6IjfTQB+BoJAld4
# 6AzyWu5mhV10LlBAOZBjGewuDDC1snwITUVLQLRUySyuVCDYkvdDXaugGf3aMeSN
# uy/YIT04byjkkrjNsv+WDerDaZyotFmg78n/VCiq2GIlHy1yX4v53k8HK8MFS8XS
# ejh8Kzx2rS5qvvc5zvCn5LmKuonlmfu809bM8fabumDl6rh5DhEJUiQ9GWZbLOfQ
# dnZB0DPMcQ==
# SIG # End signature block
