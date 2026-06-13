<#
.SYNOPSIS
    Inventory Exchange 2016 servers, CPU, RAM and disks using WMI/DCOM only.

.DESCRIPTION
    This script inventories Exchange 2016 servers from the Exchange organization and collects:
    - Exchange server identity
    - CPU and RAM inventory
    - Logical disks from Win32_LogicalDisk
    - Disk drives from Win32_DiskDrive
    - Optional Exchange mailbox database paths
    - Optional Exchange service health
    - Per-server decommissioning summary
    - Global CSV summary
    - HTML executive summary for slide integration

    This version does not use:
    - Get-CimInstance
    - New-CimSession
    - Invoke-Command
    - WinRM / PowerShell Remoting

.NOTES
    Script Name : SmartM365-Exchange-OnPrem-ServersAndStorage-Inventory.ps1
    Version     : 1.3.1
    Requirements:
      - Run from Exchange Management Shell on Exchange 2016
      - Exchange read permissions
      - Remote WMI/DCOM access to Exchange servers
      - PowerShell 5.1 or later

.CHANGELOG
    1.3.1
      - WMI/DCOM only.
      - Fixed StrictMode-safe numeric aggregation.
      - Added HTML executive summary for slide integration.
      - Kept CPU, RAM, logical disks, disk drives, mailbox database paths and service health outputs.

    1.2.1
      - Fixed PowerShell variable interpolation before colon.

    1.2.0
      - Added CPU and RAM inventory.
      - Added Win32_DiskDrive inventory.
      - Added per-server infrastructure summary.
      - Added global decommissioning summary.
      - Added HTML executive summary.
#>

[CmdletBinding()]
param(
    [string]$Tenant = "test",

    [string]$OutputRoot,

    [switch]$IncludeServicesHealth,

    [switch]$IncludeMailboxDatabasePaths,

    [switch]$SkipFqdnAndUseServerName,

    [int]$LowFreeSpaceThresholdPercent = 15
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

$ScriptName = "SmartM365-Exchange-OnPrem-ServersAndStorage-Inventory"
$ScriptVersion = "1.3.1"
$RunId = (Get-Date).ToString("yyyyMMdd-HHmmss")

function Resolve-SmartM365ConfigTokenValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    $resolved = [string]$Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$tokenName]
            if ($null -eq $property -or $null -eq $property.Value) { continue }

            $tokenValue = Resolve-SmartM365ConfigTokenValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Resolve-SmartM365ConfigTokenValue -Value '{{DataAllRootPath}}\Exchange\OnPrem\ServersAndStorage'
}
else {
    $OutputRoot = Resolve-SmartM365ConfigTokenValue -Value $OutputRoot
}

$OutputFolder = Join-Path $OutputRoot $RunId
$LogFile = Join-Path $OutputFolder "$ScriptName-v$ScriptVersion.log"

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { "$timestamp [$Level] $_" })
    $line | ForEach-Object { Write-Host $_ }
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Test-ExchangeShell {
    if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
        throw "Exchange cmdlets are not available. Run this script from the Exchange Management Shell."
    }
}

function ConvertTo-GB {
    param(
        [AllowNull()]
        [object]$Bytes
    )

    if ($null -eq $Bytes) {
        return $null
    }

    return [math]::Round(([double]$Bytes / 1GB), 2)
}

function ConvertTo-TBFromGB {
    param(
        [AllowNull()]
        [object]$GB
    )

    if ($null -eq $GB) {
        return $null
    }

    return [math]::Round(([double]$GB / 1024), 2)
}

function Get-SafeSum {
    param(
        [AllowNull()]
        [object[]]$InputObject,

        [Parameter(Mandatory)]
        [string]$Property
    )

    $sum = 0.0

    foreach ($item in @($InputObject)) {
        if ($null -eq $item) {
            continue
        }

        $propertyValue = $null

        try {
            $propertyValue = $item.$Property
        }
        catch {
            $propertyValue = $null
        }

        if ($null -eq $propertyValue -or [string]::IsNullOrWhiteSpace([string]$propertyValue)) {
            continue
        }

        try {
            $sum += [double]$propertyValue
        }
        catch {
            continue
        }
    }

    return [math]::Round($sum, 2)
}

function Format-HtmlValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return "N/A"
    }

    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        return "N/A"
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-ComputerNamesToTry {
    param(
        [Parameter(Mandatory)]
        [object]$ExchangeServer
    )

    if ($SkipFqdnAndUseServerName) {
        return @($ExchangeServer.Name)
    }

    return @($ExchangeServer.Fqdn, $ExchangeServer.Name) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
}

function Invoke-RemoteWmi {
    param(
        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry,

        [Parameter(Mandatory)]
        [string]$ClassName,

        [string]$Filter
    )

    $errors = New-Object System.Collections.Generic.List[string]

    foreach ($computerName in $ComputerNamesToTry | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
        try {
            $params = @{
                Class        = $ClassName
                ComputerName = $computerName
                ErrorAction  = "Stop"
            }

            if (-not [string]::IsNullOrWhiteSpace($Filter)) {
                $params.Filter = $Filter
            }

            return [pscustomobject]@{
                ComputerNameTried = $computerName
                Method            = "WmiDcom"
                Data              = @(Get-WmiObject @params)
                ErrorMessage      = $null
                Success           = $true
            }
        }
        catch {
            $errors.Add("WmiDcom on $computerName failed for ${ClassName}: $($_.Exception.Message)")
        }
    }

    return [pscustomobject]@{
        ComputerNameTried = ($ComputerNamesToTry -join ";")
        Method            = "WmiDcom"
        Data              = @()
        ErrorMessage      = ($errors -join " | ")
        Success           = $false
    }
}

function Test-RemoteAccess {
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeServerName,

        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry
    )

    foreach ($computerName in $ComputerNamesToTry | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
        $pingOk = $false
        $wmiOk = $false
        $wmiError = $null

        try {
            $pingOk = Test-Connection -ComputerName $computerName -Count 1 -Quiet -ErrorAction Stop
        }
        catch {
            $pingOk = $false
        }

        try {
            $null = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computerName -ErrorAction Stop | Select-Object -First 1
            $wmiOk = $true
        }
        catch {
            $wmiError = $_.Exception.Message
        }

        [pscustomobject]@{
            ExchangeServerName = $ExchangeServerName
            ComputerNameTried  = $computerName
            PingOk             = $pingOk
            WmiDcomOk           = $wmiOk
            WmiDcomError        = $wmiError
        }

        if ($wmiOk) {
            return
        }
    }
}

function Get-ServerComputeInventory {
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeServerName,

        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry
    )

    $csResult = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_ComputerSystem"
    $cpuResult = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_Processor"
    $osResult = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_OperatingSystem"

    if (-not $csResult.Success -or -not $cpuResult.Success) {
        [pscustomobject]@{
            ExchangeServerName    = $ExchangeServerName
            ComputerNameTried     = (($csResult.ComputerNameTried, $cpuResult.ComputerNameTried) | Where-Object { $_ } | Select-Object -Unique) -join ";"
            Manufacturer          = $null
            Model                 = $null
            IsVirtualMachine      = $null
            Domain                = $null
            OSName                = $null
            OSVersion             = $null
            SocketCount           = $null
            PhysicalCoreCount     = $null
            LogicalProcessorCount = $null
            MemoryGB              = $null
            CollectionMethod      = "WmiDcom"
            CollectionStatus      = "ERROR"
            ErrorMessage          = "ComputerSystem: $($csResult.ErrorMessage) | Processor: $($cpuResult.ErrorMessage)"
        }
        return
    }

    $cs = @($csResult.Data) | Select-Object -First 1
    $cpus = @($cpuResult.Data)
    $os = @($osResult.Data) | Select-Object -First 1

    $manufacturer = [string]$cs.Manufacturer
    $model = [string]$cs.Model
    $isVm = if ($manufacturer -match "VMware|Microsoft|Xen|QEMU|Virtual" -or $model -match "Virtual|VMware|KVM|Hyper-V") { $true } else { $false }

    [pscustomobject]@{
        ExchangeServerName    = $ExchangeServerName
        ComputerNameTried     = $csResult.ComputerNameTried
        Manufacturer          = $cs.Manufacturer
        Model                 = $cs.Model
        IsVirtualMachine      = $isVm
        Domain                = $cs.Domain
        OSName                = if ($os) { $os.Caption } else { $null }
        OSVersion             = if ($os) { $os.Version } else { $null }
        SocketCount           = @($cpus).Count
        PhysicalCoreCount     = Get-SafeSum -InputObject $cpus -Property "NumberOfCores"
        LogicalProcessorCount = Get-SafeSum -InputObject $cpus -Property "NumberOfLogicalProcessors"
        MemoryGB              = ConvertTo-GB -Bytes $cs.TotalPhysicalMemory
        CollectionMethod      = "WmiDcom"
        CollectionStatus      = "OK"
        ErrorMessage          = $null
    }
}

function New-LogicalDiskInventoryObject {
    param(
        [string]$ExchangeServerName,
        [string]$ComputerNameTried,
        [object]$Disk,
        [string]$CollectionStatus = "OK",
        [string]$ErrorMessage = $null
    )

    if ($null -eq $Disk) {
        return [pscustomobject]@{
            ExchangeServerName = $ExchangeServerName
            ComputerNameTried  = $ComputerNameTried
            DeviceId           = $null
            VolumeName         = $null
            FileSystem         = $null
            SizeGB             = $null
            UsedGB             = $null
            FreeGB             = $null
            FreePercent        = $null
            LowSpaceWarning    = $null
            CollectionMethod   = "WmiDcom"
            CollectionStatus   = $CollectionStatus
            ErrorMessage       = $ErrorMessage
        }
    }

    $sizeGb = ConvertTo-GB -Bytes $Disk.Size
    $freeGb = ConvertTo-GB -Bytes $Disk.FreeSpace
    $usedGb = if ($null -ne $sizeGb -and $null -ne $freeGb) { [math]::Round(($sizeGb - $freeGb), 2) } else { $null }
    $freePercent = if ($Disk.Size -and [double]$Disk.Size -gt 0) {
        [math]::Round((([double]$Disk.FreeSpace / [double]$Disk.Size) * 100), 2)
    } else {
        $null
    }

    [pscustomobject]@{
        ExchangeServerName = $ExchangeServerName
        ComputerNameTried  = $ComputerNameTried
        DeviceId           = $Disk.DeviceID
        VolumeName         = $Disk.VolumeName
        FileSystem         = $Disk.FileSystem
        SizeGB             = $sizeGb
        UsedGB             = $usedGb
        FreeGB             = $freeGb
        FreePercent        = $freePercent
        LowSpaceWarning    = if ($null -ne $freePercent -and $freePercent -lt $LowFreeSpaceThresholdPercent) { $true } else { $false }
        CollectionMethod   = "WmiDcom"
        CollectionStatus   = $CollectionStatus
        ErrorMessage       = $ErrorMessage
    }
}

function Get-ServerLogicalDiskInventory {
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeServerName,

        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry
    )

    $result = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_LogicalDisk" -Filter "DriveType = 3"

    if (-not $result.Success) {
        New-LogicalDiskInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -Disk $null -CollectionStatus "ERROR" -ErrorMessage $result.ErrorMessage
        return
    }

    if (@($result.Data).Count -eq 0) {
        New-LogicalDiskInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -Disk $null -CollectionStatus "WARNING" -ErrorMessage "No fixed logical disk returned."
        return
    }

    foreach ($disk in @($result.Data)) {
        New-LogicalDiskInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -Disk $disk
    }
}

function New-DiskDriveInventoryObject {
    param(
        [string]$ExchangeServerName,
        [string]$ComputerNameTried,
        [object]$DiskDrive,
        [string]$CollectionStatus = "OK",
        [string]$ErrorMessage = $null
    )

    if ($null -eq $DiskDrive) {
        return [pscustomobject]@{
            ExchangeServerName = $ExchangeServerName
            ComputerNameTried  = $ComputerNameTried
            Index              = $null
            Name               = $null
            DeviceId           = $null
            Model              = $null
            InterfaceType      = $null
            MediaType          = $null
            SizeGB             = $null
            Partitions         = $null
            SerialNumber       = $null
            CollectionMethod   = "WmiDcom"
            CollectionStatus   = $CollectionStatus
            ErrorMessage       = $ErrorMessage
        }
    }

    [pscustomobject]@{
        ExchangeServerName = $ExchangeServerName
        ComputerNameTried  = $ComputerNameTried
        Index              = $DiskDrive.Index
        Name               = $DiskDrive.Name
        DeviceId           = $DiskDrive.DeviceID
        Model              = $DiskDrive.Model
        InterfaceType      = $DiskDrive.InterfaceType
        MediaType          = $DiskDrive.MediaType
        SizeGB             = ConvertTo-GB -Bytes $DiskDrive.Size
        Partitions         = $DiskDrive.Partitions
        SerialNumber       = $DiskDrive.SerialNumber
        CollectionMethod   = "WmiDcom"
        CollectionStatus   = $CollectionStatus
        ErrorMessage       = $ErrorMessage
    }
}

function Get-ServerDiskDriveInventory {
    param(
        [Parameter(Mandatory)]
        [string]$ExchangeServerName,

        [Parameter(Mandatory)]
        [string[]]$ComputerNamesToTry
    )

    $result = Invoke-RemoteWmi -ComputerNamesToTry $ComputerNamesToTry -ClassName "Win32_DiskDrive"

    if (-not $result.Success) {
        New-DiskDriveInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -DiskDrive $null -CollectionStatus "ERROR" -ErrorMessage $result.ErrorMessage
        return
    }

    if (@($result.Data).Count -eq 0) {
        New-DiskDriveInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -DiskDrive $null -CollectionStatus "WARNING" -ErrorMessage "No disk drive returned."
        return
    }

    foreach ($diskDrive in @($result.Data)) {
        New-DiskDriveInventoryObject -ExchangeServerName $ExchangeServerName -ComputerNameTried $result.ComputerNameTried -DiskDrive $diskDrive
    }
}

function Get-ExchangeServiceHealthSummary {
    param(
        [Parameter(Mandatory)]
        [string]$ServerName
    )

    try {
        $health = Test-ServiceHealth -Server $ServerName -ErrorAction Stop
        $notRunning = @($health.ServicesNotRunning)

        [pscustomobject]@{
            ServerName         = $ServerName
            ServicesRunning    = if ($notRunning.Count -eq 0) { $true } else { $false }
            ServicesNotRunning = ($notRunning -join ";")
            CollectionStatus   = "OK"
            ErrorMessage       = $null
        }
    }
    catch {
        [pscustomobject]@{
            ServerName         = $ServerName
            ServicesRunning    = $null
            ServicesNotRunning = $null
            CollectionStatus   = "ERROR"
            ErrorMessage       = $_.Exception.Message
        }
    }
}

function Get-MailboxDatabasePathInventory {
    param(
        [Parameter(Mandatory)]
        [string[]]$ExchangeServerNames
    )

    $serverSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($serverName in $ExchangeServerNames) {
        [void]$serverSet.Add($serverName)
    }

    try {
        $databases = Get-MailboxDatabase -Status -ErrorAction Stop

        foreach ($database in $databases) {
            $serverName = [string]$database.Server
            if (-not $serverSet.Contains($serverName)) {
                continue
            }

            [pscustomobject]@{
                ServerName               = $serverName
                DatabaseName             = $database.Name
                Mounted                  = $database.Mounted
                DatabaseSize             = if ($database.DatabaseSize) { $database.DatabaseSize.ToString() } else { $null }
                AvailableNewMailboxSpace = if ($database.AvailableNewMailboxSpace) { $database.AvailableNewMailboxSpace.ToString() } else { $null }
                EdbFilePath              = if ($database.EdbFilePath) { $database.EdbFilePath.PathName } else { $null }
                LogFolderPath            = if ($database.LogFolderPath) { $database.LogFolderPath.PathName } else { $null }
                CircularLoggingEnabled   = $database.CircularLoggingEnabled
                Recovery                 = $database.Recovery
                ReplicationType          = $database.ReplicationType
                ErrorMessage             = $null
            }
        }
    }
    catch {
        [pscustomobject]@{
            ServerName               = $null
            DatabaseName             = $null
            Mounted                  = $null
            DatabaseSize             = $null
            AvailableNewMailboxSpace = $null
            EdbFilePath              = $null
            LogFolderPath            = $null
            CircularLoggingEnabled   = $null
            Recovery                 = $null
            ReplicationType          = $null
            ErrorMessage             = $_.Exception.Message
        }
    }
}

function New-HtmlExecutiveSummary {
    param(
        [Parameter(Mandatory)]
        [object]$Summary,

        [Parameter(Mandatory)]
        [object[]]$PerServerSummary,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $rowsHtml = foreach ($row in ($PerServerSummary | Sort-Object ExchangeServerName)) {
        "<tr><td>$(Format-HtmlValue $row.ExchangeServerName)</td><td>$(Format-HtmlValue $row.ServerRole)</td><td class='num'>$(Format-HtmlValue $row.LogicalProcessorCount)</td><td class='num'>$(Format-HtmlValue $row.MemoryGB)</td><td class='num'>$(Format-HtmlValue $row.DiskDriveCount)</td><td class='num'>$(Format-HtmlValue $row.DiskDriveTotalSizeGB)</td><td class='status'>$(Format-HtmlValue $row.ComputeCollectionStatus)</td></tr>"
    }

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Exchange On-Premises Decommissioning Summary</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2937; background: #ffffff; }
h1 { font-size: 28px; margin: 0 0 8px 0; color: #111827; }
.subtitle { font-size: 14px; color: #4b5563; margin-bottom: 24px; }
.cards { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 22px; }
.card { border: 1px solid #d1d5db; border-radius: 12px; padding: 16px; background: #f9fafb; }
.card .label { font-size: 12px; text-transform: uppercase; letter-spacing: .04em; color: #6b7280; }
.card .value { font-size: 30px; font-weight: 700; margin-top: 6px; color: #111827; }
.card .unit { font-size: 13px; color: #6b7280; margin-top: 2px; }
.section { margin-top: 18px; }
.summary { border-left: 4px solid #374151; padding: 10px 14px; background: #f3f4f6; font-size: 15px; line-height: 1.45; }
table { width: 100%; border-collapse: collapse; margin-top: 14px; font-size: 12px; }
th { text-align: left; background: #111827; color: white; padding: 8px; }
td { border-bottom: 1px solid #e5e7eb; padding: 7px 8px; }
td.num, th.num { text-align: right; }
.status { font-weight: 600; }
.footer { margin-top: 18px; font-size: 11px; color: #6b7280; }
</style>
</head>
<body>
<h1>Exchange On-Premises Infrastructure Decommissioning</h1>
<div class="subtitle">Executive summary generated on $(Format-HtmlValue $Summary.ExecutionDate) — RunId $(Format-HtmlValue $Summary.RunId)</div>

<div class="cards">
  <div class="card"><div class="label">Exchange VMs</div><div class="value">$(Format-HtmlValue $Summary.ExchangeServersCount)</div><div class="unit">servers identified</div></div>
  <div class="card"><div class="label">vCPU</div><div class="value">$(Format-HtmlValue $Summary.TotalLogicalProcessorCount)</div><div class="unit">logical processors</div></div>
  <div class="card"><div class="label">RAM</div><div class="value">$(Format-HtmlValue $Summary.TotalMemoryGB)</div><div class="unit">GB</div></div>
  <div class="card"><div class="label">Disks</div><div class="value">$(Format-HtmlValue $Summary.TotalDiskDriveCount)</div><div class="unit">WMI disk drives</div></div>
  <div class="card"><div class="label">Provisioned Storage</div><div class="value">$(Format-HtmlValue $Summary.TotalDiskDriveSizeTB)</div><div class="unit">TB</div></div>
</div>

<div class="section summary">
The current legacy Exchange 2016 on-premises footprint represents <strong>$(Format-HtmlValue $Summary.ExchangeServersCount) virtual machines</strong>, <strong>$(Format-HtmlValue $Summary.TotalLogicalProcessorCount) vCPU</strong>, <strong>$(Format-HtmlValue $Summary.TotalMemoryGB) GB RAM</strong>, and <strong>$(Format-HtmlValue $Summary.TotalDiskDriveCount) disks</strong>. These assets are candidates for decommissioning after Exchange SE migration validation and dependency sign-off.
</div>

<div class="section">
<table>
<thead>
<tr>
<th>Server</th>
<th>Role</th>
<th class="num">vCPU</th>
<th class="num">RAM GB</th>
<th class="num">Disks</th>
<th class="num">Disk Size GB</th>
<th>Status</th>
</tr>
</thead>
<tbody>
$($rowsHtml -join "`r`n")
</tbody>
</table>
</div>

<div class="footer">
Note: CPU, RAM and disk data are collected from the guest OS using WMI/DCOM only. Validate with Infrastructure if hypervisor-level figures are required for final capacity reclamation.
</div>
</body>
</html>
"@

    Set-Content -Path $Path -Value $html -Encoding UTF8
}

try {
    Ensure-Directory -Path $OutputFolder
    Write-Log "Starting $ScriptName v$ScriptVersion. RunId: $RunId"
    Write-Log "Output folder: $OutputFolder"
    Write-Log "Collection method: WMI/DCOM only"

    Test-ExchangeShell

    Write-Log "Collecting Exchange servers."
    $exchangeServers = @(Get-ExchangeServer | Sort-Object Name)

    if ($exchangeServers.Count -eq 0) {
        throw "No Exchange servers were returned by Get-ExchangeServer."
    }

    $serverInventory = foreach ($server in $exchangeServers) {
        [pscustomobject]@{
            Name                = $server.Name
            Fqdn                = $server.Fqdn
            Site                = if ($server.Site) { $server.Site.ToString() } else { $null }
            ServerRole          = $server.ServerRole
            AdminDisplayVersion = $server.AdminDisplayVersion.ToString()
            Edition             = $server.Edition
            IsExchange2016      = if ($server.AdminDisplayVersion.ToString() -like "Version 15.1*") { $true } else { $false }
        }
    }

    $serverInventoryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_Inventory.csv"
    $serverInventory | Export-Csv -Path $serverInventoryPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Exchange server inventory exported to: $serverInventoryPath"

    Write-Log "Testing remote access with WMI/DCOM."
    $remoteAccessInventory = foreach ($server in $exchangeServers) {
        $namesToTry = Get-ComputerNamesToTry -ExchangeServer $server
        Write-Log "Testing WMI/DCOM access for $($server.Name)"
        Test-RemoteAccess -ExchangeServerName $server.Name -ComputerNamesToTry $namesToTry
    }

    $remoteAccessPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_RemoteAccess.csv"
    $remoteAccessInventory | Export-Csv -Path $remoteAccessPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Remote access report exported to: $remoteAccessPath"

    Write-Log "Collecting CPU and RAM inventory."
    $computeInventory = foreach ($server in $exchangeServers) {
        $namesToTry = Get-ComputerNamesToTry -ExchangeServer $server
        Write-Log "Collecting compute inventory from $($server.Name)"
        Get-ServerComputeInventory -ExchangeServerName $server.Name -ComputerNamesToTry $namesToTry
    }

    $computeInventoryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_Compute.csv"
    $computeInventory | Export-Csv -Path $computeInventoryPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Compute inventory exported to: $computeInventoryPath"

    Write-Log "Collecting logical disk inventory."
    $logicalDiskInventory = foreach ($server in $exchangeServers) {
        $namesToTry = Get-ComputerNamesToTry -ExchangeServer $server
        Write-Log "Collecting logical disks from $($server.Name)"
        Get-ServerLogicalDiskInventory -ExchangeServerName $server.Name -ComputerNamesToTry $namesToTry
    }

    $logicalDiskInventoryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_LogicalDisks.csv"
    $logicalDiskInventory | Export-Csv -Path $logicalDiskInventoryPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Logical disk inventory exported to: $logicalDiskInventoryPath"

    Write-Log "Collecting disk drive inventory."
    $diskDriveInventory = foreach ($server in $exchangeServers) {
        $namesToTry = Get-ComputerNamesToTry -ExchangeServer $server
        Write-Log "Collecting disk drives from $($server.Name)"
        Get-ServerDiskDriveInventory -ExchangeServerName $server.Name -ComputerNamesToTry $namesToTry
    }

    $diskDriveInventoryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_DiskDrives.csv"
    $diskDriveInventory | Export-Csv -Path $diskDriveInventoryPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Disk drive inventory exported to: $diskDriveInventoryPath"

    if ($IncludeServicesHealth) {
        Write-Log "Collecting Exchange services health."
        $serviceHealth = foreach ($server in $exchangeServers) {
            Write-Log "Collecting service health from $($server.Name)"
            Get-ExchangeServiceHealthSummary -ServerName $server.Name
        }

        $serviceHealthPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_ServiceHealth.csv"
        $serviceHealth | Export-Csv -Path $serviceHealthPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Write-Log "Service health inventory exported to: $serviceHealthPath"
    }

    if ($IncludeMailboxDatabasePaths) {
        Write-Log "Collecting mailbox database paths."
        $databasePaths = @(Get-MailboxDatabasePathInventory -ExchangeServerNames @($exchangeServers.Name))
        $databasePathsPath = Join-Path $OutputFolder "Exchange_OnPrem_MailboxDatabases_Paths.csv"
        $databasePaths | Export-Csv -Path $databasePathsPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Write-Log "Mailbox database path inventory exported to: $databasePathsPath"
    }

    $logicalDiskRows = @($logicalDiskInventory)
    $successfulLogicalDiskRows = @($logicalDiskRows | Where-Object { $_.CollectionStatus -eq "OK" })
    $failedLogicalDiskRows = @($logicalDiskRows | Where-Object { $_.CollectionStatus -eq "ERROR" })
    $lowSpaceRows = @($successfulLogicalDiskRows | Where-Object { $_.LowSpaceWarning -eq $true })

    $diskDriveRows = @($diskDriveInventory)
    $successfulDiskDriveRows = @($diskDriveRows | Where-Object { $_.CollectionStatus -eq "OK" })
    $failedDiskDriveRows = @($diskDriveRows | Where-Object { $_.CollectionStatus -eq "ERROR" })

    $computeRows = @($computeInventory)
    $successfulComputeRows = @($computeRows | Where-Object { $_.CollectionStatus -eq "OK" })
    $failedComputeRows = @($computeRows | Where-Object { $_.CollectionStatus -eq "ERROR" })

    Write-Log "Building per-server infrastructure summary."
    $perServerSummary = foreach ($server in $exchangeServers) {
        $serverName = $server.Name
        $compute = $computeRows | Where-Object { $_.ExchangeServerName -eq $serverName } | Select-Object -First 1
        $serverLogicalDisks = @($successfulLogicalDiskRows | Where-Object { $_.ExchangeServerName -eq $serverName })
        $serverDiskDrives = @($successfulDiskDriveRows | Where-Object { $_.ExchangeServerName -eq $serverName })

        [pscustomobject]@{
            ExchangeServerName          = $serverName
            Fqdn                        = $server.Fqdn
            ServerRole                  = $server.ServerRole
            IsExchange2016              = if ($server.AdminDisplayVersion.ToString() -like "Version 15.1*") { $true } else { $false }
            IsVirtualMachine            = if ($compute) { $compute.IsVirtualMachine } else { $null }
            Manufacturer                = if ($compute) { $compute.Manufacturer } else { $null }
            Model                       = if ($compute) { $compute.Model } else { $null }
            SocketCount                 = if ($compute) { $compute.SocketCount } else { $null }
            PhysicalCoreCount           = if ($compute) { $compute.PhysicalCoreCount } else { $null }
            LogicalProcessorCount       = if ($compute) { $compute.LogicalProcessorCount } else { $null }
            MemoryGB                    = if ($compute) { $compute.MemoryGB } else { $null }
            DiskDriveCount              = $serverDiskDrives.Count
            DiskDriveTotalSizeGB        = Get-SafeSum -InputObject $serverDiskDrives -Property "SizeGB"
            LogicalDiskCount            = $serverLogicalDisks.Count
            LogicalDiskTotalSizeGB      = Get-SafeSum -InputObject $serverLogicalDisks -Property "SizeGB"
            LogicalDiskUsedGB           = Get-SafeSum -InputObject $serverLogicalDisks -Property "UsedGB"
            LogicalDiskFreeGB           = Get-SafeSum -InputObject $serverLogicalDisks -Property "FreeGB"
            LowSpaceLogicalDiskCount    = @($serverLogicalDisks | Where-Object { $_.LowSpaceWarning -eq $true }).Count
            ComputeCollectionStatus     = if ($compute) { $compute.CollectionStatus } else { "ERROR" }
            DiskDriveCollectionStatus   = if (@($diskDriveRows | Where-Object { $_.ExchangeServerName -eq $serverName -and $_.CollectionStatus -eq "ERROR" }).Count -gt 0) { "ERROR" } else { "OK" }
            LogicalDiskCollectionStatus = if (@($logicalDiskRows | Where-Object { $_.ExchangeServerName -eq $serverName -and $_.CollectionStatus -eq "ERROR" }).Count -gt 0) { "ERROR" } else { "OK" }
        }
    }

    $perServerSummaryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_Decommissioning_PerServerSummary.csv"
    $perServerSummary | Export-Csv -Path $perServerSummaryPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Per-server decommissioning summary exported to: $perServerSummaryPath"

    $totalMemoryGB = Get-SafeSum -InputObject $successfulComputeRows -Property "MemoryGB"
    $totalDiskDriveSizeGB = Get-SafeSum -InputObject $successfulDiskDriveRows -Property "SizeGB"
    $totalLogicalDiskSizeGB = Get-SafeSum -InputObject $successfulLogicalDiskRows -Property "SizeGB"
    $totalLogicalDiskUsedGB = Get-SafeSum -InputObject $successfulLogicalDiskRows -Property "UsedGB"
    $totalLogicalDiskFreeGB = Get-SafeSum -InputObject $successfulLogicalDiskRows -Property "FreeGB"

    $summary = [pscustomobject]@{
        ScriptName                       = $ScriptName
        ScriptVersion                    = $ScriptVersion
        RunId                            = $RunId
        ExecutionDate                    = Get-Date
        OutputFolder                     = $OutputFolder
        ExchangeServersCount             = $exchangeServers.Count
        ComputeRowsCount                 = $computeRows.Count
        ComputeCollectionErrors          = $failedComputeRows.Count
        TotalSocketCount                 = Get-SafeSum -InputObject $successfulComputeRows -Property "SocketCount"
        TotalPhysicalCoreCount           = Get-SafeSum -InputObject $successfulComputeRows -Property "PhysicalCoreCount"
        TotalLogicalProcessorCount       = Get-SafeSum -InputObject $successfulComputeRows -Property "LogicalProcessorCount"
        TotalMemoryGB                    = $totalMemoryGB
        TotalMemoryTB                    = ConvertTo-TBFromGB -GB $totalMemoryGB
        DiskDriveRowsCount               = $diskDriveRows.Count
        SuccessfulDiskDriveRowsCount     = $successfulDiskDriveRows.Count
        DiskDriveCollectionErrors        = $failedDiskDriveRows.Count
        TotalDiskDriveCount              = $successfulDiskDriveRows.Count
        TotalDiskDriveSizeGB             = $totalDiskDriveSizeGB
        TotalDiskDriveSizeTB             = ConvertTo-TBFromGB -GB $totalDiskDriveSizeGB
        LogicalDiskRowsCount             = $logicalDiskRows.Count
        SuccessfulLogicalDiskRowsCount   = $successfulLogicalDiskRows.Count
        LogicalDiskCollectionErrors      = $failedLogicalDiskRows.Count
        LowSpaceWarnings                 = $lowSpaceRows.Count
        TotalLogicalDiskSizeGB           = $totalLogicalDiskSizeGB
        TotalLogicalDiskSizeTB           = ConvertTo-TBFromGB -GB $totalLogicalDiskSizeGB
        TotalLogicalDiskUsedGB           = $totalLogicalDiskUsedGB
        TotalLogicalDiskUsedTB           = ConvertTo-TBFromGB -GB $totalLogicalDiskUsedGB
        TotalLogicalDiskFreeGB           = $totalLogicalDiskFreeGB
        TotalLogicalDiskFreeTB           = ConvertTo-TBFromGB -GB $totalLogicalDiskFreeGB
    }

    $summaryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_Inventory_Summary.csv"
    $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Global summary exported to: $summaryPath"

    $htmlSummaryPath = Join-Path $OutputFolder "Exchange_OnPrem_Servers_Decommissioning_ExecutiveSummary.html"
    New-HtmlExecutiveSummary -Summary $summary -PerServerSummary $perServerSummary -Path $htmlSummaryPath
    Write-Log "HTML executive summary exported to: $htmlSummaryPath"

    Write-Log "Completed successfully."
}
catch {
    Write-Log -Level "ERROR" -Message $_.Exception.Message
    throw
}
