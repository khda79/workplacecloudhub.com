<#
.SYNOPSIS
    Active Directory multi-domain pre-check before Exchange SE preparation.

.DESCRIPTION
    Read-only Active Directory readiness checks before Exchange SE /PrepareSchema,
    /PrepareAD and /PrepareAllDomains. Adapted for SmartM365 SmartInventory with
    tenant-aware configuration, DATA-ALL/DATA-LAST CSV exports, LOG-ALL logs,
    SharePoint upload and weekly history through SmartM365.Core.

.VERSION
    1.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Requirements: RSAT ActiveDirectory module, AD read access, repadmin/dcdiag for full checks.
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string[]]$TargetDomains,
    [string]$OutputRoot,
    [int]$EventLookbackDays = 7,
    [switch]$SkipDcdiag,
    [switch]$SkipEvents,
    [switch]$SkipPortChecks,
    [switch]$IncludeAllDcEvents,
    [switch]$SkipRecommendedCommands
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptName = 'SmartM365-ExchangeSE-ADPreCheck-Inventory'
$ScriptVersion = '1.0'
$RunId = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:FailureStage = 'Initialization'
$script:TranscriptStarted = $false
$script:CheckResults = New-Object System.Collections.Generic.List[object]
$script:Rows = @{
    Domains = New-Object System.Collections.Generic.List[object]
    DomainControllers = New-Object System.Collections.Generic.List[object]
    ExchangeVersions = New-Object System.Collections.Generic.List[object]
    Replication = New-Object System.Collections.Generic.List[object]
    DnsSrv = New-Object System.Collections.Generic.List[object]
    Sysvol = New-Object System.Collections.Generic.List[object]
    SchemaConnectivity = New-Object System.Collections.Generic.List[object]
    Events = New-Object System.Collections.Generic.List[object]
    CurrentUserGroups = New-Object System.Collections.Generic.List[object]
    PrivilegedGroups = New-Object System.Collections.Generic.List[object]
    Readiness = New-Object System.Collections.Generic.List[object]
}
$script:EffectiveConfig = $null
$script:Forest = $null
$script:TargetDomainNames = @()
$script:OutputFolder = ''
$script:LatestCsvFolder = ''

function Resolve-SmartM365TenantContextPath {
    $d = $PSScriptRoot
    while ($d) {
        foreach ($candidate in @((Join-Path $d 'SmartM365-TenantContext.ps1'), (Join-Path $d 'Config\SmartM365-TenantContext.ps1'))) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}

. (Resolve-SmartM365TenantContextPath)
$script:EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

function ConvertTo-LocalHashtable {
    param([AllowNull()]$InputObject)
    $table = @{}
    if ($null -eq $InputObject) { return $table }
    if ($InputObject -is [hashtable]) { foreach ($key in $InputObject.Keys) { $table[$key] = $InputObject[$key] }; return $table }
    foreach ($property in $InputObject.PSObject.Properties) { $table[$property.Name] = $property.Value }
    return $table
}

function Resolve-ConfigTokenValue {
    param([AllowNull()]$Value)
    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return '' }
    $resolved = [string]$Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $matches) {
            $property = $script:EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $tokenValue = Resolve-ConfigTokenValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Initialize-ScriptLocalConfig {
    $configPath = Join-Path $PSScriptRoot ("{0}.local.json" -f $ScriptName)
    $templatePath = "{0}.template" -f $configPath
    if (-not (Test-Path -LiteralPath $configPath) -and (Test-Path -LiteralPath $templatePath)) {
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
            Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null
        } else { Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop }
    }
    if (-not (Test-Path -LiteralPath $configPath)) { return }
    $localConfig = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    $effective = ConvertTo-LocalHashtable -InputObject $script:EffectiveConfig
    foreach ($property in $localConfig.PSObject.Properties) {
        if ($property.Value -is [string] -and $property.Value -in @('__USE_GLOBAL__', 'USE_GLOBAL', '')) { continue }
        $effective[$property.Name] = $property.Value
    }
    $script:EffectiveConfig = [pscustomobject]$effective
}

function Get-ConfigValue {
    param([Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)
    $property = if ($script:EffectiveConfig) { $script:EffectiveConfig.PSObject.Properties[$Name] } else { $null }
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    if ($property.Value -is [string] -and $property.Value -in @('__USE_GLOBAL__', 'USE_GLOBAL', '')) { return $DefaultValue }
    return Resolve-ConfigTokenValue -Value $property.Value
}

Initialize-ScriptLocalConfig

$modulePath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365.Core module not found.'
}
Import-Module $modulePath -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = [string](Get-ConfigValue -Name 'OutputRoot' -DefaultValue '{{DataAllRootPath}}\Exchange\OnPrem\ExchangeSEPreCheck') }
$OutputRoot = [string](Resolve-ConfigTokenValue -Value $OutputRoot)
$script:OutputFolder = Join-Path $OutputRoot $RunId
$script:LatestCsvFolder = [string](Resolve-ConfigTokenValue -Value (Get-ConfigValue -Name 'LatestCsvFolderPath' -DefaultValue ''))

function New-SafeFileName { param([Parameter(Mandatory)][string]$Name) return ($Name -replace '[\\/:*?"<>|]', '_') }
function Test-CommandAvailable { param([Parameter(Mandatory)][string]$CommandName) return [bool](Get-Command -Name $CommandName -ErrorAction SilentlyContinue) }
function Get-ObjectValue { param([AllowNull()]$Object, [Parameter(Mandatory)][string]$Name) if ($null -eq $Object) { return $null }; $p = $Object.PSObject.Properties[$Name]; if ($null -eq $p) { return $null }; return $p.Value }

function Add-ExchangeSECheckResult {
    param([Parameter(Mandatory)][string]$Scope, [Parameter(Mandatory)][string]$Check, [Parameter(Mandatory)][ValidateSet('OK','WARNING','ERROR','SKIPPED')][string]$Status, [string]$Details = '')
    $row = [pscustomobject]@{ Time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'); Scope = $Scope; Check = $Check; Status = $Status; Details = $Details }
    $script:CheckResults.Add($row) | Out-Null
    $level = switch ($Status) { 'ERROR' { 'ERROR' } 'WARNING' { 'WARNING' } default { 'INFO' } }
    WriteLog -Message ("[{0}] {1} - {2}: {3}" -f $Status, $Scope, $Check, $Details) -Level $level
}

function Invoke-RawCommand {
    param([Parameter(Mandatory)][string]$CommandName, [Parameter(Mandatory)][string[]]$Arguments, [Parameter(Mandatory)][string]$OutputFile)
    WriteLog -Message ("Running command: {0} {1}" -f $CommandName, ($Arguments -join ' '))
    & $CommandName @Arguments 2>&1 | Tee-Object -FilePath $OutputFile | Out-Null
    return $LASTEXITCODE
}

function Export-ExchangeSECsv {
    param([Parameter(Mandatory)][string]$BaseName, [AllowNull()][object[]]$Data, [string[]]$Columns = @())
    $timestampedPath = Join-Path $script:OutputFolder ("{0}_{1}.csv" -f $BaseName, $RunId)
    $latestPath = if ([string]::IsNullOrWhiteSpace($script:LatestCsvFolder)) { '' } else { Join-Path $script:LatestCsvFolder ("{0}.csv" -f $BaseName) }
    Publish-SmartM365Csv -Data @($Data) -TimestampedPath $timestampedPath -LatestPath $latestPath -Columns $Columns | Out-Null
}

function Test-TcpPortQuiet {
    param([Parameter(Mandatory)][string]$ComputerName, [Parameter(Mandatory)][int]$Port)
    try {
        if (Get-Command Test-NetConnection -ErrorAction SilentlyContinue) { return [bool](Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -WarningAction SilentlyContinue) }
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        $ok = $async.AsyncWaitHandle.WaitOne(3000, $false)
        if ($ok) { $client.EndConnect($async) }
        $client.Close()
        return [bool]$ok
    } catch { return $false }
}

try {
    $null = InitializeScriptEnvironment -OutputPathInit $script:OutputFolder -LogFileName $ScriptName
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    $script:TranscriptStarted = $true
    WriteLog -Message ("Starting {0} v{1}" -f $ScriptName, $ScriptVersion)
    WriteLog -Message ("Output folder: {0}" -f $script:OutputFolder)

    $script:FailureStage = 'Preflight'
    Import-Module ActiveDirectory -ErrorAction Stop
    Invoke-SmartM365Preflight -ScriptName $ScriptName -OutputPaths @($script:OutputFolder) -RequiredModules @('ActiveDirectory') -RequireActiveDirectoryRead | Out-Null
    Add-ExchangeSECheckResult -Scope 'Local' -Check 'ActiveDirectory module' -Status 'OK' -Details 'ActiveDirectory module loaded.'

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]$identity
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ($isAdmin) { Add-ExchangeSECheckResult -Scope 'Local' -Check 'PowerShell elevated' -Status 'OK' -Details 'PowerShell is running as Administrator.' }
    else { Add-ExchangeSECheckResult -Scope 'Local' -Check 'PowerShell elevated' -Status 'WARNING' -Details 'PowerShell is not running as Administrator. Some remote checks may fail.' }

    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_PowerShellEnvironment' -Data @([pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        UserName = $identity.Name
        PowerShell = $PSVersionTable.PSVersion.ToString()
        PSEdition = $PSVersionTable.PSEdition
        OS = if ($os) { $os.Caption } else { [Environment]::OSVersion.VersionString }
        OSVersion = if ($os) { $os.Version } else { [Environment]::OSVersion.Version.ToString() }
        IsAdmin = $isAdmin
        ReportPath = $script:OutputFolder
    })

    $script:FailureStage = 'ForestDiscovery'
    $rootDse = Get-ADRootDSE -ErrorAction Stop
    $script:Forest = Get-ADForest -ErrorAction Stop
    $script:TargetDomainNames = if ($TargetDomains -and $TargetDomains.Count -gt 0) { @($TargetDomains) } else { @($script:Forest.Domains) }
    Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Forest discovery' -Status 'OK' -Details ("Forest discovered: {0}; domains in scope: {1}." -f $script:Forest.Name, $script:TargetDomainNames.Count)
    Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Schema Master' -Status 'OK' -Details ("Schema Master: {0}." -f $script:Forest.SchemaMaster)
    Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Domain Naming Master' -Status 'OK' -Details ("Domain Naming Master: {0}." -f $script:Forest.DomainNamingMaster)

    foreach ($domainName in $script:TargetDomainNames) {
        try {
            $domain = Get-ADDomain -Identity $domainName -ErrorAction Stop
            $script:Rows.Domains.Add([pscustomobject]@{ Forest = $script:Forest.Name; Domain = $domain.DNSRoot; NetBIOSName = $domain.NetBIOSName; DistinguishedName = $domain.DistinguishedName; DomainMode = $domain.DomainMode; PDCEmulator = $domain.PDCEmulator; RIDMaster = $domain.RIDMaster; InfrastructureMaster = $domain.InfrastructureMaster }) | Out-Null
            Add-ExchangeSECheckResult -Scope $domainName -Check 'Domain discovery' -Status 'OK' -Details ("Domain mode: {0}." -f $domain.DomainMode)
        } catch { Add-ExchangeSECheckResult -Scope $domainName -Check 'Domain discovery' -Status 'ERROR' -Details $_.Exception.Message }
    }

    $script:FailureStage = 'GroupMembership'
    try {
        foreach ($sid in $identity.Groups) {
            $groupName = try { $sid.Translate([Security.Principal.NTAccount]).Value } catch { $sid.Value }
            $script:Rows.CurrentUserGroups.Add([pscustomobject]@{ UserName = $identity.Name; Group = $groupName; Sid = $sid.Value }) | Out-Null
        }
        Add-ExchangeSECheckResult -Scope 'Local' -Check 'Current user token groups' -Status 'OK' -Details ("{0} group SID(s) exported." -f $script:Rows.CurrentUserGroups.Count)
    } catch { Add-ExchangeSECheckResult -Scope 'Local' -Check 'Current user token groups' -Status 'WARNING' -Details $_.Exception.Message }

    foreach ($groupName in @('Schema Admins','Enterprise Admins','Domain Admins','Organization Management')) {
        try {
            $group = Get-ADGroup -Filter "Name -eq '$groupName'" -Properties DistinguishedName -ErrorAction Stop | Select-Object -First 1
            if ($null -eq $group) { Add-ExchangeSECheckResult -Scope 'Forest' -Check "Privileged group - $groupName" -Status 'WARNING' -Details 'Group not found from current AD context.'; continue }
            $members = @(Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -ErrorAction Stop)
            foreach ($member in $members) { $script:Rows.PrivilegedGroups.Add([pscustomobject]@{ GroupName = $groupName; MemberName = $member.Name; SamAccountName = $member.SamAccountName; ObjectClass = $member.ObjectClass; DistinguishedName = $member.DistinguishedName }) | Out-Null }
            Add-ExchangeSECheckResult -Scope 'Forest' -Check "Privileged group - $groupName" -Status 'OK' -Details ("{0} member(s) exported." -f $members.Count)
        } catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check "Privileged group - $groupName" -Status 'WARNING' -Details $_.Exception.Message }
    }

    $script:FailureStage = 'DomainControllers'
    foreach ($domainName in $script:TargetDomainNames) {
        try {
            $dcs = @(Get-ADDomainController -Filter * -Server $domainName -ErrorAction Stop)
            foreach ($dc in $dcs) {
                $script:Rows.DomainControllers.Add([pscustomobject]@{ Domain = $domainName; HostName = $dc.HostName; Name = $dc.Name; Site = $dc.Site; IPv4Address = $dc.IPv4Address; IsGlobalCatalog = [bool]$dc.IsGlobalCatalog; IsReadOnly = [bool]$dc.IsReadOnly; OperatingSystem = $dc.OperatingSystem; OperatingSystemVersion = $dc.OperatingSystemVersion }) | Out-Null
            }
            $gcCount = @($dcs | Where-Object { $_.IsGlobalCatalog }).Count
            if ($gcCount -gt 0) { Add-ExchangeSECheckResult -Scope $domainName -Check 'Global Catalog availability' -Status 'OK' -Details ("{0} Global Catalog server(s) found." -f $gcCount) }
            else { Add-ExchangeSECheckResult -Scope $domainName -Check 'Global Catalog availability' -Status 'WARNING' -Details 'No Global Catalog server found in this domain.' }
        } catch { Add-ExchangeSECheckResult -Scope $domainName -Check 'Domain controllers' -Status 'ERROR' -Details $_.Exception.Message }
    }

    $script:FailureStage = 'ExchangeAdVersions'
    $schemaNc = [string]$rootDse.schemaNamingContext
    $configNc = [string]$rootDse.configurationNamingContext
    try {
        $schemaVersion = Get-ADObject "CN=ms-Exch-Schema-Version-Pt,$schemaNc" -Properties rangeUpper -ErrorAction Stop
        $script:Rows.ExchangeVersions.Add([pscustomobject]@{ Scope = 'Forest'; ObjectType = 'ExchangeSchemaVersion'; Name = $schemaVersion.Name; DistinguishedName = $schemaVersion.DistinguishedName; RangeUpper = $schemaVersion.rangeUpper; ObjectVersion = $null; MsExchProductId = $null }) | Out-Null
        Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange schema version' -Status 'OK' -Details ("rangeUpper={0}." -f $schemaVersion.rangeUpper)
    } catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange schema version' -Status 'WARNING' -Details $_.Exception.Message }

    try {
        $exchangeConfig = Get-ADObject "CN=Microsoft Exchange,CN=Services,$configNc" -Properties objectVersion, msExchProductId -ErrorAction Stop
        $script:Rows.ExchangeVersions.Add([pscustomobject]@{ Scope = 'Forest'; ObjectType = 'ExchangeConfigurationObject'; Name = $exchangeConfig.Name; DistinguishedName = $exchangeConfig.DistinguishedName; RangeUpper = $null; ObjectVersion = $exchangeConfig.objectVersion; MsExchProductId = $exchangeConfig.msExchProductId }) | Out-Null
        Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange configuration object' -Status 'OK' -Details ("objectVersion={0}; msExchProductId={1}." -f $exchangeConfig.objectVersion, $exchangeConfig.msExchProductId)
    } catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange configuration object' -Status 'WARNING' -Details $_.Exception.Message }

    try {
        $orgContainers = @(Get-ADObject -LDAPFilter '(objectClass=msExchOrganizationContainer)' -SearchBase $configNc -Properties objectVersion, msExchProductId -ErrorAction Stop)
        foreach ($org in $orgContainers) { $script:Rows.ExchangeVersions.Add([pscustomobject]@{ Scope = 'Forest'; ObjectType = 'ExchangeOrganizationContainer'; Name = $org.Name; DistinguishedName = $org.DistinguishedName; RangeUpper = $null; ObjectVersion = $org.objectVersion; MsExchProductId = $org.msExchProductId }) | Out-Null }
        if ($orgContainers.Count -gt 0) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange organization container' -Status 'OK' -Details ("{0} Exchange organization container(s) found." -f $orgContainers.Count) }
        else { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange organization container' -Status 'WARNING' -Details 'No Exchange organization container found.' }
    } catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange organization container' -Status 'WARNING' -Details $_.Exception.Message }

    foreach ($domainName in $script:TargetDomainNames) {
        try {
            $domain = Get-ADDomain -Identity $domainName -ErrorAction Stop
            $obj = Get-ADObject "CN=Microsoft Exchange System Objects,$($domain.DistinguishedName)" -Properties objectVersion -ErrorAction Stop
            $script:Rows.ExchangeVersions.Add([pscustomobject]@{ Scope = $domainName; ObjectType = 'ExchangeDomainObject'; Name = $obj.Name; DistinguishedName = $obj.DistinguishedName; RangeUpper = $null; ObjectVersion = $obj.objectVersion; MsExchProductId = $null }) | Out-Null
            Add-ExchangeSECheckResult -Scope $domainName -Check 'Exchange domain object' -Status 'OK' -Details ("objectVersion={0}." -f $obj.objectVersion)
        } catch { Add-ExchangeSECheckResult -Scope $domainName -Check 'Exchange domain object' -Status 'WARNING' -Details $_.Exception.Message }
    }

    $script:FailureStage = 'ReplicationChecks'
    if (Test-CommandAvailable -CommandName 'repadmin.exe') {
        try {
            Invoke-RawCommand 'repadmin.exe' @('/replsummary') (Join-Path $script:OutputFolder 'repadmin-replsummary.txt') | Out-Null
            Invoke-RawCommand 'repadmin.exe' @('/replsummary','/bysrc','/bydest','/sort:delta') (Join-Path $script:OutputFolder 'repadmin-replsummary-bysrc-bydest.txt') | Out-Null
            Invoke-RawCommand 'repadmin.exe' @('/showrepl','*') (Join-Path $script:OutputFolder 'repadmin-showrepl.txt') | Out-Null
            Invoke-RawCommand 'repadmin.exe' @('/showrepl','*','/csv') (Join-Path $script:OutputFolder 'repadmin-showrepl.csv') | Out-Null
            Invoke-RawCommand 'repadmin.exe' @('/queue','*') (Join-Path $script:OutputFolder 'repadmin-queue.txt') | Out-Null
            Add-ExchangeSECheckResult -Scope 'Forest' -Check 'repadmin outputs' -Status 'OK' -Details 'Forest-wide repadmin outputs collected.'
        } catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'repadmin outputs' -Status 'ERROR' -Details $_.Exception.Message }
    } else { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'repadmin' -Status 'ERROR' -Details 'repadmin.exe not found.' }

    foreach ($domainName in $script:TargetDomainNames) {
        try {
            $metadata = @(Get-ADReplicationPartnerMetadata -Target $domainName -Scope Domain -ErrorAction Stop)
            foreach ($item in $metadata) {
                $failureCount = [int](Get-ObjectValue -Object $item -Name 'ConsecutiveReplicationFailures')
                $script:Rows.Replication.Add([pscustomobject]@{ Domain = $domainName; Server = $item.Server; Partner = $item.Partner; Partition = $item.Partition; LastReplicationSuccess = $item.LastReplicationSuccess; LastReplicationAttempt = $item.LastReplicationAttempt; ConsecutiveReplicationFailures = $failureCount; LastReplicationResult = $item.LastReplicationResult; LastReplicationResultCode = $item.LastReplicationResultCode }) | Out-Null
            }
            $failures = @($metadata | Where-Object { [int](Get-ObjectValue -Object $_ -Name 'ConsecutiveReplicationFailures') -gt 0 })
            if ($failures.Count -gt 0) { Add-ExchangeSECheckResult -Scope $domainName -Check 'AD replication metadata' -Status 'ERROR' -Details ("{0} replication issue(s) found." -f $failures.Count) }
            else { Add-ExchangeSECheckResult -Scope $domainName -Check 'AD replication metadata' -Status 'OK' -Details 'No replication failure found.' }
        } catch { Add-ExchangeSECheckResult -Scope $domainName -Check 'AD replication metadata' -Status 'WARNING' -Details $_.Exception.Message }
    }

    $script:FailureStage = 'DcdiagChecks'
    if ($SkipDcdiag) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'DCDIAG' -Status 'SKIPPED' -Details 'Skipped by parameter.' }
    elseif (Test-CommandAvailable -CommandName 'dcdiag.exe') {
        try {
            $fullPath = Join-Path $script:OutputFolder 'dcdiag-full-forest.txt'
            $quickPath = Join-Path $script:OutputFolder 'dcdiag-quick-forest.txt'
            $dnsPath = Join-Path $script:OutputFolder 'dcdiag-dns-forest.txt'
            Invoke-RawCommand 'dcdiag.exe' @('/e','/c','/v',"/f:$fullPath") (Join-Path $script:OutputFolder 'dcdiag-full-console.txt') | Out-Null
            Invoke-RawCommand 'dcdiag.exe' @('/e','/c','/q',"/f:$quickPath") (Join-Path $script:OutputFolder 'dcdiag-quick-console.txt') | Out-Null
            Invoke-RawCommand 'dcdiag.exe' @('/e','/test:DNS','/v',"/f:$dnsPath") (Join-Path $script:OutputFolder 'dcdiag-dns-console.txt') | Out-Null
            $quickContent = if (Test-Path -LiteralPath $quickPath) { Get-Content -LiteralPath $quickPath -Raw -ErrorAction SilentlyContinue } else { '' }
            if ([string]::IsNullOrWhiteSpace($quickContent)) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'DCDIAG quick' -Status 'OK' -Details 'dcdiag /q returned no issue.' }
            else { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'DCDIAG quick' -Status 'WARNING' -Details 'dcdiag /q returned content. Review dcdiag-quick-forest.txt.' }
        } catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'DCDIAG' -Status 'ERROR' -Details $_.Exception.Message }
    } else { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'DCDIAG' -Status 'ERROR' -Details 'dcdiag.exe not found.' }

    $script:FailureStage = 'DnsSrvChecks'
    foreach ($domainName in $script:TargetDomainNames) {
        foreach ($prefix in @('_ldap._tcp.dc._msdcs','_ldap._tcp.gc._msdcs','_kerberos._tcp','_kerberos._udp','_ldap._tcp')) {
            $recordName = "$prefix.$domainName"
            try {
                $records = @(Resolve-DnsName -Name $recordName -Type SRV -ErrorAction Stop | Where-Object { $_.Type -eq 'SRV' })
                if ($records.Count -eq 0) { $script:Rows.DnsSrv.Add([pscustomobject]@{ Domain = $domainName; RecordName = $recordName; Status = 'ERROR'; NameTarget = ''; Port = ''; Priority = ''; Weight = ''; Details = 'No SRV record returned.' }) | Out-Null }
                foreach ($record in $records) { $script:Rows.DnsSrv.Add([pscustomobject]@{ Domain = $domainName; RecordName = $recordName; Status = 'OK'; NameTarget = $record.NameTarget; Port = $record.Port; Priority = $record.Priority; Weight = $record.Weight; Details = '' }) | Out-Null }
            } catch { $script:Rows.DnsSrv.Add([pscustomobject]@{ Domain = $domainName; RecordName = $recordName; Status = 'ERROR'; NameTarget = ''; Port = ''; Priority = ''; Weight = ''; Details = $_.Exception.Message }) | Out-Null }
        }
        $dnsErrors = @($script:Rows.DnsSrv | Where-Object { $_.Domain -eq $domainName -and $_.Status -eq 'ERROR' })
        if ($dnsErrors.Count -gt 0) { Add-ExchangeSECheckResult -Scope $domainName -Check 'DNS SRV records' -Status 'ERROR' -Details ("{0} DNS SRV lookup(s) failed." -f $dnsErrors.Count) }
        else { Add-ExchangeSECheckResult -Scope $domainName -Check 'DNS SRV records' -Status 'OK' -Details 'DNS SRV lookups succeeded.' }
    }

    $script:FailureStage = 'AdSitesSubnets'
    try {
        $sites = @(Get-ADObject -LDAPFilter '(objectClass=site)' -SearchBase "CN=Sites,$configNc" -Properties name -ErrorAction Stop)
        $subnets = @(Get-ADObject -LDAPFilter '(objectClass=subnet)' -SearchBase "CN=Subnets,CN=Sites,$configNc" -Properties name, siteObject -ErrorAction Stop)
        $siteRows = @($sites | ForEach-Object { [pscustomobject]@{ ObjectType = 'Site'; Name = $_.Name; DistinguishedName = $_.DistinguishedName; SiteObject = '' } })
        $subnetRows = @($subnets | ForEach-Object { [pscustomobject]@{ ObjectType = 'Subnet'; Name = $_.Name; DistinguishedName = $_.DistinguishedName; SiteObject = [string]$_.siteObject } })
        Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_ADSitesSubnets' -Data @($siteRows + $subnetRows) -Columns @('ObjectType','Name','DistinguishedName','SiteObject')
        Add-ExchangeSECheckResult -Scope 'Forest' -Check 'AD sites and subnets' -Status 'OK' -Details ("{0} site(s), {1} subnet(s) exported." -f $sites.Count, $subnets.Count)
    } catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'AD sites and subnets' -Status 'WARNING' -Details $_.Exception.Message }

    $script:FailureStage = 'SysvolNetlogon'
    foreach ($dc in $script:Rows.DomainControllers) {
        $sysvolPath = "\\{0}\SYSVOL" -f $dc.HostName
        $netlogonPath = "\\{0}\NETLOGON" -f $dc.HostName
        $script:Rows.Sysvol.Add([pscustomobject]@{ Domain = $dc.Domain; HostName = $dc.HostName; SysvolPath = $sysvolPath; SysvolAvailable = (Test-Path -LiteralPath $sysvolPath); NetlogonPath = $netlogonPath; NetlogonAvailable = (Test-Path -LiteralPath $netlogonPath) }) | Out-Null
    }
    $sysvolFailures = @($script:Rows.Sysvol | Where-Object { -not $_.SysvolAvailable -or -not $_.NetlogonAvailable })
    if ($sysvolFailures.Count -gt 0) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'SYSVOL/NETLOGON' -Status 'ERROR' -Details ("{0} DC(s) have SYSVOL or NETLOGON issue." -f $sysvolFailures.Count) }
    else { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'SYSVOL/NETLOGON' -Status 'OK' -Details 'SYSVOL and NETLOGON are available on all discovered DCs.' }

    $script:FailureStage = 'DfsrState'
    if (Test-CommandAvailable -CommandName 'dfsrdiag.exe') {
        try { Invoke-RawCommand 'dfsrdiag.exe' @('ReplicationState') (Join-Path $script:OutputFolder 'dfsrdiag-replicationstate.txt') | Out-Null; Add-ExchangeSECheckResult -Scope 'Forest' -Check 'DFSR state' -Status 'OK' -Details 'DFSR replication state collected.' }
        catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'DFSR state' -Status 'WARNING' -Details $_.Exception.Message }
    } else { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'DFSR state' -Status 'WARNING' -Details 'dfsrdiag.exe not found.' }

    $script:FailureStage = 'EventLogs'
    if ($SkipEvents) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'AD event logs' -Status 'SKIPPED' -Details 'Skipped by parameter.' }
    else {
        $startTime = (Get-Date).AddDays(-1 * [Math]::Abs($EventLookbackDays))
        $eventDcs = if ($IncludeAllDcEvents) { @($script:Rows.DomainControllers) } else { @($script:Rows.DomainControllers | Group-Object Domain | ForEach-Object { $_.Group | Select-Object -First 5 }) }
        if (-not $IncludeAllDcEvents) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'AD event logs scope' -Status 'WARNING' -Details 'Event collection is limited to up to 5 DCs per domain. Use -IncludeAllDcEvents for all DCs.' }
        foreach ($dc in $eventDcs) {
            foreach ($logName in @('Directory Service','DNS Server','DFS Replication','System')) {
                try {
                    $events = @(Get-WinEvent -ComputerName $dc.HostName -FilterHashtable @{ LogName = $logName; Level = 1,2,3; StartTime = $startTime } -ErrorAction Stop)
                    $status = if ($events.Count -gt 0) { 'WARNING' } else { 'OK' }
                    $script:Rows.Events.Add([pscustomobject]@{ Domain = $dc.Domain; HostName = $dc.HostName; LogName = $logName; LookbackDays = $EventLookbackDays; EventCount = $events.Count; Status = $status; Details = '' }) | Out-Null
                    $eventDetails = if ($events.Count -gt 0) { "{0} warning/error/critical event(s) found." -f $events.Count } else { 'No warning/error/critical event found.' }`r`n                    Add-ExchangeSECheckResult -Scope $dc.Domain -Check "Events - $($dc.HostName) - $logName" -Status $status -Details $eventDetails
                } catch {
                    $script:Rows.Events.Add([pscustomobject]@{ Domain = $dc.Domain; HostName = $dc.HostName; LogName = $logName; LookbackDays = $EventLookbackDays; EventCount = $null; Status = 'WARNING'; Details = $_.Exception.Message }) | Out-Null
                    Add-ExchangeSECheckResult -Scope $dc.Domain -Check "Events - $($dc.HostName) - $logName" -Status 'WARNING' -Details $_.Exception.Message
                }
            }
        }
    }

    $script:FailureStage = 'SchemaBackupEvents'
    try {
        $schemaMaster = [string]$script:Forest.SchemaMaster
        $backupEvents = @(Get-WinEvent -ComputerName $schemaMaster -LogName 'Microsoft-Windows-Backup' -MaxEvents 50 -ErrorAction Stop)
        $backupRows = @($backupEvents | ForEach-Object { [pscustomobject]@{ ComputerName = $schemaMaster; TimeCreated = $_.TimeCreated; Id = $_.Id; LevelDisplayName = $_.LevelDisplayName; ProviderName = $_.ProviderName; Message = (($_.Message -replace '[\r\n]+',' ') -replace '\s+',' ') } })
        Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_SchemaMasterBackupEvents' -Data $backupRows
        if ($backupEvents.Count -gt 0) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Schema Master backup events' -Status 'OK' -Details ("{0} Windows Backup event(s) exported from {1}." -f $backupEvents.Count, $schemaMaster) }
        else { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Schema Master backup events' -Status 'WARNING' -Details ("No Windows Backup event found on {0}. Verify System State backup manually." -f $schemaMaster) }
    } catch { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Schema Master backup events' -Status 'WARNING' -Details $_.Exception.Message }

    $script:FailureStage = 'SchemaMasterConnectivity'
    if ($SkipPortChecks) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Schema Master connectivity' -Status 'SKIPPED' -Details 'Skipped by parameter.' }
    else {
        $schemaMaster = [string]$script:Forest.SchemaMaster
        $conn = [pscustomobject]@{
            SchemaMaster = $schemaMaster
            Ping = (Test-Connection -ComputerName $schemaMaster -Count 2 -Quiet -ErrorAction SilentlyContinue)
            LDAP389 = (Test-TcpPortQuiet -ComputerName $schemaMaster -Port 389)
            GC3268 = (Test-TcpPortQuiet -ComputerName $schemaMaster -Port 3268)
            SMB445 = (Test-TcpPortQuiet -ComputerName $schemaMaster -Port 445)
            RPC135 = (Test-TcpPortQuiet -ComputerName $schemaMaster -Port 135)
        }
        $script:Rows.SchemaConnectivity.Add($conn) | Out-Null
        if ($conn.Ping -and $conn.LDAP389 -and $conn.SMB445 -and $conn.RPC135) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Schema Master connectivity' -Status 'OK' -Details 'Schema Master is reachable on key ports.' }
        else { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Schema Master connectivity' -Status 'ERROR' -Details 'Schema Master connectivity issue detected.' }
    }

    $script:FailureStage = 'RecommendedCommands'
    if ($SkipRecommendedCommands) { Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange preparation commands' -Status 'SKIPPED' -Details 'Skipped by parameter.' }
    else {
        $commandsPath = Join-Path $script:OutputFolder 'ExchangeSE-Preparation-Commands.txt'
        $commands = New-Object System.Collections.Generic.List[string]
        foreach ($line in @(
            'Run only after reviewing this pre-check report and confirming a recent System State backup of the Schema Master.',
            '',
            '1. Prepare schema:',
            'Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareSchema',
            '',
            '2. Wait for AD replication:',
            'repadmin /syncall /AdeP',
            'repadmin /replsummary',
            '',
            '3. Prepare AD:',
            'Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD',
            '',
            '4. Wait for AD replication again:',
            'repadmin /syncall /AdeP',
            'repadmin /replsummary',
            '',
            '5. Multi-domain forest preparation:',
            'Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAllDomains',
            '',
            'Alternative if you prefer domain by domain:'
        )) { $commands.Add($line) | Out-Null }
        foreach ($domainName in $script:TargetDomainNames) { $commands.Add("Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareDomain:$domainName") | Out-Null }
        foreach ($line in @('', '6. Final replication check:', 'repadmin /replsummary', 'dcdiag /e /c /q')) { $commands.Add($line) | Out-Null }
        $commands | Set-Content -LiteralPath $commandsPath -Encoding UTF8
        Add-ExchangeSECheckResult -Scope 'Forest' -Check 'Exchange preparation commands' -Status 'OK' -Details 'Command reference exported to ExchangeSE-Preparation-Commands.txt.'
    }

    $script:FailureStage = 'CsvExports'
    $errorRows = @($script:CheckResults | Where-Object { $_.Status -eq 'ERROR' })
    $warningRows = @($script:CheckResults | Where-Object { $_.Status -eq 'WARNING' })
    $skippedRows = @($script:CheckResults | Where-Object { $_.Status -eq 'SKIPPED' })
    $gcCount = @($script:Rows.DomainControllers | Where-Object { $_.IsGlobalCatalog }).Count
    $readinessStatus = if ($errorRows.Count -gt 0) { 'NOT_READY' } elseif ($warningRows.Count -gt 0) { 'REVIEW_REQUIRED' } else { 'READY_FROM_SCRIPT_PERSPECTIVE' }
    $script:Rows.Readiness.Add([pscustomobject]@{
        RunId = $RunId
        Forest = $script:Forest.Name
        DomainsInScope = $script:TargetDomainNames.Count
        DomainControllers = $script:Rows.DomainControllers.Count
        GlobalCatalogs = $gcCount
        Errors = $errorRows.Count
        Warnings = $warningRows.Count
        Skipped = $skippedRows.Count
        EventLookbackDays = $EventLookbackDays
        IncludeAllDcEvents = [bool]$IncludeAllDcEvents
        SkipDcdiag = [bool]$SkipDcdiag
        SkipEvents = [bool]$SkipEvents
        SkipPortChecks = [bool]$SkipPortChecks
        ReadinessStatus = $readinessStatus
        OutputFolder = $script:OutputFolder
        LogFile = $global:LogTextFile
        TranscriptFile = $global:logTranscriptFile
    }) | Out-Null

    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_CheckResults' -Data @($script:CheckResults) -Columns @('Time','Scope','Check','Status','Details')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_Domains' -Data @($script:Rows.Domains) -Columns @('Forest','Domain','NetBIOSName','DistinguishedName','DomainMode','PDCEmulator','RIDMaster','InfrastructureMaster')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_DomainControllers' -Data @($script:Rows.DomainControllers) -Columns @('Domain','HostName','Name','Site','IPv4Address','IsGlobalCatalog','IsReadOnly','OperatingSystem','OperatingSystemVersion')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_ExchangeAdVersions' -Data @($script:Rows.ExchangeVersions) -Columns @('Scope','ObjectType','Name','DistinguishedName','RangeUpper','ObjectVersion','MsExchProductId')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_Replication' -Data @($script:Rows.Replication) -Columns @('Domain','Server','Partner','Partition','LastReplicationSuccess','LastReplicationAttempt','ConsecutiveReplicationFailures','LastReplicationResult','LastReplicationResultCode')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_DnsSrv' -Data @($script:Rows.DnsSrv) -Columns @('Domain','RecordName','Status','NameTarget','Port','Priority','Weight','Details')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_SysvolNetlogon' -Data @($script:Rows.Sysvol) -Columns @('Domain','HostName','SysvolPath','SysvolAvailable','NetlogonPath','NetlogonAvailable')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_SchemaMasterConnectivity' -Data @($script:Rows.SchemaConnectivity) -Columns @('SchemaMaster','Ping','LDAP389','GC3268','SMB445','RPC135')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_EventSummary' -Data @($script:Rows.Events) -Columns @('Domain','HostName','LogName','LookbackDays','EventCount','Status','Details')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_ReadinessSummary' -Data @($script:Rows.Readiness) -Columns @('RunId','Forest','DomainsInScope','DomainControllers','GlobalCatalogs','Errors','Warnings','Skipped','EventLookbackDays','IncludeAllDcEvents','SkipDcdiag','SkipEvents','SkipPortChecks','ReadinessStatus','OutputFolder','LogFile','TranscriptFile')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_CurrentUserGroups' -Data @($script:Rows.CurrentUserGroups) -Columns @('UserName','Group','Sid')
    Export-ExchangeSECsv -BaseName 'Exchange_OnPrem_ExchangeSEPreCheck_PrivilegedGroupMembers' -Data @($script:Rows.PrivilegedGroups) -Columns @('GroupName','MemberName','SamAccountName','ObjectClass','DistinguishedName')

    $summaryTextPath = Join-Path $script:OutputFolder 'ExchangeSE-ADPreCheck-Summary.txt'
    @("RunId: $RunId", "Forest: $($script:Forest.Name)", "ReadinessStatus: $readinessStatus", "Errors: $($errorRows.Count)", "Warnings: $($warningRows.Count)", "OutputFolder: $($script:OutputFolder)", "LogFile: $($global:LogTextFile)", "TranscriptFile: $($global:logTranscriptFile)") | Set-Content -LiteralPath $summaryTextPath -Encoding UTF8

    if ([bool](Get-ConfigValue -Name 'SendSummaryEmail' -DefaultValue $true)) {
        $script:FailureStage = 'SummaryEmail'
        $topIssues = @($script:CheckResults | Where-Object { $_.Status -in @('ERROR','WARNING') } | Select-Object -First 15)
        $issueRows = if ($topIssues.Count -gt 0) {
            ($topIssues | ForEach-Object { '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>' -f [System.Net.WebUtility]::HtmlEncode($_.Status), [System.Net.WebUtility]::HtmlEncode($_.Scope), [System.Net.WebUtility]::HtmlEncode($_.Check), [System.Net.WebUtility]::HtmlEncode($_.Details) }) -join "`n"
        } else { '<tr><td colspan="4">No error or warning detected.</td></tr>' }
        $bodyHtml = @"
<h1>Exchange SE AD pre-check $readinessStatus</h1>
<p>Forest: $($script:Forest.Name)</p>
<table>
<tr><th>Metric</th><th>Value</th></tr>
<tr><td>Domains in scope</td><td>$($script:TargetDomainNames.Count)</td></tr>
<tr><td>Domain controllers</td><td>$($script:Rows.DomainControllers.Count)</td></tr>
<tr><td>Global Catalogs</td><td>$gcCount</td></tr>
<tr><td>Errors</td><td>$($errorRows.Count)</td></tr>
<tr><td>Warnings</td><td>$($warningRows.Count)</td></tr>
<tr><td>Skipped</td><td>$($skippedRows.Count)</td></tr>
</table>
<h2>Top issues</h2>
<table>
<tr><th>Status</th><th>Scope</th><th>Check</th><th>Details</th></tr>
$issueRows
</table>
<p style="font-size:11px;color:#64748b;">Output: $($script:OutputFolder)<br/>Log: $($global:LogTextFile)</p>
"@
        SendEmailHtmlReport -Subject ("SmartM365 Exchange SE AD pre-check - {0}" -f $readinessStatus) -BodyHtml $bodyHtml -Attachments @($summaryTextPath) -VerboseLog
    } else { WriteLog -Message 'Summary email skipped because SendSummaryEmail is false.' -Level 'INFO' }

    WriteLog -Message ("Exchange SE AD pre-check completed with status {0}." -f $readinessStatus) -Level 'SUCCESS'
}
catch {
    $globalError = $_
    WriteLog -Message ("Exchange SE AD pre-check failed at stage {0}: {1}" -f $script:FailureStage, $_.Exception.Message) -Level 'ERROR'
    throw
}
finally {
    if ($script:TranscriptStarted) { try { Stop-Transcript | Out-Null } catch {}; $script:TranscriptStarted = $false }
    $finalError = if (Get-Variable -Name globalError -Scope Local -ErrorAction SilentlyContinue) { $globalError } else { $null }
    Complete-SmartM365ExecutionContext -Status Auto -ErrorRecord $finalError -FailureStage $script:FailureStage
}

