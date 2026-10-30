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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAXfBxeMBceuss5
# XJeZM1YYtCwQgd+syb0EsqXXX8imV6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD7HtUxoCSgr0AygJL3qCxU
# 5/g4FpKNCA2v5ZkxP1ROmTANBgkqhkiG9w0BAQEFAASCAYCn6OmTysl+3+zCk/BH
# NuCwnX2N/zmJ9YzY+CeyuqucLa2ieSY8SGNCcFESbwT5NLCrXaNF8wp+XXLS/38H
# iJuT5VbA5P+vod49P2pDFrFyMGsABvEnRFjRhCsCPN2isA0dujx71gPOzRXPWWil
# HAI8JZfEHrFz2BvQBcP0L3vkkj189VEn9XX0wbF/F1gocouU10D69l8doSLZs/25
# lU4cVxDBH6ClF0QY/aQ8zxu78M7SsiO6zrK8DAfifw3A8ahWN1GkQKFWlZQfVdo/
# CKGOABxjZJlV5xUvbctk2UIiwiF+LJqwB1smHwWVvNOkaIoEAcFO0Y1TyaQOs8kS
# SkEhvLMk8E50ceO7fKzJarC/ncMV5ovs5NUYdr69fkcwz0h677N6eH4tkYpLCBVv
# gxRV6JMfIcbMaZOgpgK+Zjta2dM5MAXvS3vtuC/LcZPqSd73oM6eJQRw2LlI5WEb
# 9u12FOBoJs/HEM4o0UvpsDcT4bOzjhoHoiJmx0nTgBEaz24=
# SIG # End signature block
