<#
.SYNOPSIS
    Common primitives of the Smart Endpoint Migration Toolkit engine.

.DESCRIPTION
    Product constants, runtime paths, JSON files with template self-healing, hashing,
    script versions, file locks, local ACLs, technician identity and timestamped logs.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

$script:SemtProductName = 'SmartEndpointMigrationToolkit'
$script:SemtProductDisplayName = 'Smart Endpoint Migration Toolkit'
$script:SemtProductVersion = '1.0.0'
$script:SemtProductSubtitle = 'Hybrid Join, Intune Enrollment & Windows 11 Upgrade'
$script:SemtCampaigns = @('HybridJoin','Windows11Upgrade')
$script:SemtWorkerContractVersion = 1
$script:SemtLastRunSchemaVersion = 2
$script:SemtTargetRuntimeRoot = 'C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit'
# Endpoint scripts started as SYSTEM on targets are stored here (administrators only), never under
# the ProgramData runtime root, which keeps only data.
$script:SemtTargetProgramFilesRoot = 'C:\Program Files'

function Get-SemtProductInfo {
    [CmdletBinding()]
    param()

    return [pscustomobject]@{
        Name = $script:SemtProductName
        DisplayName = $script:SemtProductDisplayName
        Subtitle = $script:SemtProductSubtitle
        Version = $script:SemtProductVersion
        EngineVersion = $script:SemtModuleVersion
        WorkerContractVersion = $script:SemtWorkerContractVersion
        LastRunSchemaVersion = $script:SemtLastRunSchemaVersion
        Campaigns = @($script:SemtCampaigns)
        TargetRuntimeRoot = $script:SemtTargetRuntimeRoot
    }
}

function Assert-SemtCampaign {
    [CmdletBinding()]
    param([AllowNull()][string]$Campaign)

    if ([string]::IsNullOrWhiteSpace($Campaign) -or $script:SemtCampaigns -notcontains $Campaign) {
        throw ("Unsupported Campaign '{0}'. Expected one of: {1}." -f $Campaign,($script:SemtCampaigns -join ', '))
    }
    return $Campaign
}

function Get-SemtRuntimeRoot {
    [CmdletBinding()]
    param()

    # Offline tests and demos isolate technician state in a temporary folder. Never set in production.
    $override = [Environment]::GetEnvironmentVariable('SEMT_RUNTIME_ROOT_OVERRIDE', 'Process')
    if (-not [string]::IsNullOrWhiteSpace($override)) { return [System.IO.Path]::GetFullPath($override) }
    $programData = [Environment]::GetFolderPath('CommonApplicationData')
    if ([string]::IsNullOrWhiteSpace($programData)) { $programData = 'C:\ProgramData' }
    return (Join-Path -Path $programData -ChildPath 'WCH\SmartM365\SmartEndpointMigrationToolkit')
}

function Get-SemtLauncherStateRoot {
    [CmdletBinding()]
    param([string]$RuntimeRoot)

    if ([string]::IsNullOrWhiteSpace($RuntimeRoot)) { $RuntimeRoot = Get-SemtRuntimeRoot }
    return (Join-Path -Path $RuntimeRoot -ChildPath 'LauncherState')
}

function Get-SemtTargetCampaignRoot {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Campaign)

    [void](Assert-SemtCampaign -Campaign $Campaign)
    return (Join-Path -Path $script:SemtTargetRuntimeRoot -ChildPath $Campaign)
}

function Get-SemtTargetProgramRoot {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Campaign)

    [void](Assert-SemtCampaign -Campaign $Campaign)
    return (Join-Path -Path $script:SemtTargetProgramFilesRoot -ChildPath ('WCH\SmartM365\SmartEndpointMigrationToolkit\{0}' -f $Campaign))
}

function Get-SemtTargetProgramFilesRoot {
    [CmdletBinding()]
    param()
    return $script:SemtTargetProgramFilesRoot
}

function Get-SemtPackageRoot {
    [CmdletBinding()]
    param([string]$StartPath)

    if ([string]::IsNullOrWhiteSpace($StartPath)) { $StartPath = $script:SemtModuleRoot }
    $current = [System.IO.Path]::GetFullPath($StartPath)
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        $marker = Join-Path -Path $current -ChildPath 'Config\SmartEndpointMigrationToolkit.options.json.txt'
        if (Test-Path -LiteralPath $marker -PathType Leaf) { return $current }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    throw "Smart Endpoint Migration Toolkit package root not found from '$StartPath'."
}

function New-SemtDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Get-SemtUtcNowText {
    [CmdletBinding()]
    param()
    return (Get-Date).ToUniversalTime().ToString('o')
}

function ConvertTo-SemtUtcDateTime {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return ([datetime]$Value).ToUniversalTime() }
    if ($Value -is [datetimeoffset]) { return ([datetimeoffset]$Value).UtcDateTime }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
        return $parsed.ToUniversalTime()
    }
    return $null
}

function Get-SemtLongPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    $full = [System.IO.Path]::GetFullPath($Path)
    if ($full.StartsWith('\\?\')) { return $full }
    if ($full.StartsWith('\\')) { return ('\\?\UNC\' + $full.Substring(2)) }
    return ('\\?\' + $full)
}

function Get-SemtFileSha256 {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = $null
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        try { $stream = [System.IO.File]::OpenRead($Path) }
        catch { $stream = [System.IO.File]::OpenRead((Get-SemtLongPath -Path $Path)) }
        return (([System.BitConverter]::ToString($sha.ComputeHash($stream))) -replace '-','')
    }
    finally {
        if ($stream) { $stream.Dispose() }
        $sha.Dispose()
    }
}

function Get-SemtStringSha256 {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
        return (([System.BitConverter]::ToString($sha.ComputeHash($bytes))) -replace '-','')
    }
    finally { $sha.Dispose() }
}

function Get-SemtScriptVersion {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    $text = [System.IO.File]::ReadAllText($Path)
    $match = [regex]::Match($text, '(?im)^[ \t]*\.VERSION(?:[ \t]+(?<v>[^\r\n]+)|[ \t]*\r?\n[ \t]*(?<v>[^\r\n]+))')
    if ($match.Success) { return $match.Groups['v'].Value.Trim() }
    return ''
}

function Get-SemtSafeName {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Value,
        [string]$Fallback = 'Unknown',
        [ValidateRange(8, 200)][int]$MaxLength = 80
    )

    $safe = [regex]::Replace(([string]$Value).Trim(), '[^A-Za-z0-9._-]', '_').Trim('.','_','-',' ')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = $Fallback }
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength) }
    return $safe
}

function New-SemtRunId {
    [CmdletBinding()]
    param()
    return ('{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'),([guid]::NewGuid().ToString('N').Substring(0, 6)))
}

function ConvertTo-SemtHashtable {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $copy = [ordered]@{}
        foreach ($key in @($InputObject.Keys)) { $copy[[string]$key] = ConvertTo-SemtHashtable -InputObject $InputObject[$key] }
        return $copy
    }
    if ($InputObject -is [pscustomobject]) {
        $copy = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) { $copy[$property.Name] = ConvertTo-SemtHashtable -InputObject $property.Value }
        return $copy
    }
    if (($InputObject -is [System.Collections.IEnumerable]) -and -not ($InputObject -is [string])) {
        $items = New-Object System.Collections.ArrayList
        foreach ($item in $InputObject) { [void]$items.Add((ConvertTo-SemtHashtable -InputObject $item)) }
        return ,($items.ToArray())
    }
    return $InputObject
}

#region Supported PowerShell and JSON dates

function Test-SemtSupportedPowerShell {
    # Windows PowerShell 5.1, or PowerShell 7.2 and later. Older engines are refused before any
    # state is written (PowerShell 7.0 and 7.1 are out of support and not qualified).
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Edition = [string]$PSVersionTable['PSEdition'],
        [AllowNull()][version]$Version = $PSVersionTable['PSVersion']
    )
    if ($null -eq $Version) { return $false }
    if ([string]::IsNullOrWhiteSpace($Edition) -or $Edition -eq 'Desktop') { return ($Version -ge [version]'5.1') }
    if ($Edition -eq 'Core') { return ($Version -ge [version]'7.2') }
    return $false
}

function Assert-SemtSupportedPowerShell {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Edition = [string]$PSVersionTable['PSEdition'],
        [AllowNull()][version]$Version = $PSVersionTable['PSVersion']
    )
    if (-not (Test-SemtSupportedPowerShell -Edition $Edition -Version $Version)) {
        throw ("SEMT_UNSUPPORTED_POWERSHELL: {0} {1} is not supported. Use Windows PowerShell 5.1 or PowerShell 7.2 and later." -f $Edition,$Version)
    }
}

function ConvertTo-SemtJsonDateText {
    # PowerShell 7 converts ISO 8601 strings to DateTime while reading JSON (without -DateKind
    # before 7.5). Every such value is turned back into round-trip text so all engines see the
    # same values as Windows PowerShell 5.1: UTC and offset values as UTC text ('o', same instant),
    # values without zone as text without zone.
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        $date = [datetime]$Value
        if ($date.Kind -eq [System.DateTimeKind]::Unspecified) { return $date.ToString('yyyy-MM-ddTHH:mm:ss.FFFFFFF', [Globalization.CultureInfo]::InvariantCulture) }
        return $date.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [datetimeoffset]) { return ([datetimeoffset]$Value).UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        foreach ($property in @($Value.PSObject.Properties)) { $property.Value = ConvertTo-SemtJsonDateText -Value $property.Value }
        return $Value
    }
    if ($Value -is [System.Collections.IList]) {
        for ($index = 0; $index -lt $Value.Count; $index++) { $Value[$index] = ConvertTo-SemtJsonDateText -Value $Value[$index] }
        return ,$Value
    }
    return $Value
}

#endregion

$script:SemtJsonDateKindSupported = (Get-Command -Name ConvertFrom-Json).Parameters.ContainsKey('DateKind')

function Read-SemtJsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [AllowNull()]$Default = $null,
        [switch]$Required
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        if ($Required) { throw "Required JSON file not found: $Path" }
        return $Default
    }
    $lastError = $null
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $raw = [System.IO.File]::ReadAllText($Path)
            if ([string]::IsNullOrWhiteSpace($raw)) {
                if ($Required) { throw "Required JSON file is empty: $Path" }
                return $Default
            }
            # PowerShell 7 converts ISO dates to DateTime; keep them as text so the UTC round-trip
            # values used by reservations and PID start-time checks stay exact (as on 5.1).
            if ($script:SemtJsonDateKindSupported) { return ($raw | ConvertFrom-Json -DateKind String -ErrorAction Stop) }
            return (ConvertTo-SemtJsonDateText -Value ($raw | ConvertFrom-Json -ErrorAction Stop))
        }
        catch [System.IO.IOException] {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(1000, 100 * $attempt))
        }
        catch {
            throw ("Invalid JSON file {0}: {1}" -f $Path,$_.Exception.Message)
        }
    }
    throw ("JSON file could not be read after retries: {0}; Error={1}" -f $Path,$lastError.Exception.Message)
}

function Write-SemtJsonFileAtomic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowNull()]$InputObject,
        [ValidateRange(2, 30)][int]$Depth = 12
    )

    $parent = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent)) { New-SemtDirectory -Path $parent }
    $json = $InputObject | ConvertTo-Json -Depth $Depth
    $encoding = New-Object System.Text.UTF8Encoding($true)
    $lastError = $null
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        $tempPath = Join-Path -Path $parent -ChildPath ('.{0}.{1}.tmp' -f (Split-Path -Path $Path -Leaf),([guid]::NewGuid().ToString('N')))
        try {
            [System.IO.File]::WriteAllText($tempPath, $json, $encoding)
            Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
            return
        }
        catch {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(2000, 150 * $attempt))
        }
        finally {
            if (Test-Path -LiteralPath $tempPath -PathType Leaf) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
        }
    }
    throw ("JSON file could not be written: {0}; Error={1}" -f $Path,$lastError.Exception.Message)
}

function Add-SemtMissingTemplateKeys {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Target,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Template,
        [string]$Prefix = '',
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[string]]$AddedKeys
    )

    foreach ($key in @($Template.Keys)) {
        $keyPath = if ([string]::IsNullOrWhiteSpace($Prefix)) { [string]$key } else { '{0}.{1}' -f $Prefix,$key }
        $templateValue = $Template[$key]
        if (-not $Target.Contains($key)) {
            $Target[$key] = $templateValue
            $AddedKeys.Add($keyPath)
            continue
        }
        if (($templateValue -is [System.Collections.IDictionary]) -and ($Target[$key] -is [System.Collections.IDictionary])) {
            Add-SemtMissingTemplateKeys -Target $Target[$key] -Template $templateValue -Prefix $keyPath -AddedKeys $AddedKeys
        }
    }
}

function Sync-SemtJsonWithTemplate {
    <#
    Adds keys present in the committed template but missing from the runtime JSON
    (recursively), saves the runtime file and preserves every existing local value.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$TemplatePath,
        [switch]$CreateWhenMissing
    )

    $template = ConvertTo-SemtHashtable -InputObject (Read-SemtJsonFile -Path $TemplatePath -Required)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        if (-not $CreateWhenMissing) {
            return [pscustomobject]@{ Path = $Path; Exists = $false; AddedKeys = @(); Data = $template }
        }
        Write-SemtJsonFileAtomic -Path $Path -InputObject $template
        return [pscustomobject]@{ Path = $Path; Exists = $true; AddedKeys = @('<created from template>'); Data = $template }
    }

    $runtime = ConvertTo-SemtHashtable -InputObject (Read-SemtJsonFile -Path $Path -Required)
    if ($null -eq $runtime) { $runtime = [ordered]@{} }
    $added = New-Object 'System.Collections.Generic.List[string]'
    Add-SemtMissingTemplateKeys -Target $runtime -Template $template -AddedKeys $added
    if ($added.Count -gt 0) { Write-SemtJsonFileAtomic -Path $Path -InputObject $runtime }
    return [pscustomobject]@{ Path = $Path; Exists = $true; AddedKeys = @($added.ToArray()); Data = $runtime }
}

function Invoke-SemtWithFileLock {
    <#
    Cross-session and cross-account critical section based on an exclusive file handle.
    The lock is released automatically if the owning process dies.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LockPath,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @(),
        [ValidateRange(1, 3600)][int]$TimeoutSeconds = 60
    )

    $parent = Split-Path -Path $LockPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent)) { New-SemtDirectory -Path $parent }
    $deadline = (Get-Date).ToUniversalTime().AddSeconds($TimeoutSeconds)
    $stream = $null
    while ($null -eq $stream) {
        try {
            $stream = [System.IO.File]::Open($LockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] { $stream = $null }
        catch [System.UnauthorizedAccessException] { $stream = $null }
        if ($null -eq $stream) {
            if ((Get-Date).ToUniversalTime() -ge $deadline) {
                throw ("SEMT_LOCK_TIMEOUT: lock not acquired within {0} second(s): {1}" -f $TimeoutSeconds,$LockPath)
            }
            Start-Sleep -Milliseconds 150
        }
    }
    try { return (& $ScriptBlock @ArgumentList) }
    finally { $stream.Dispose() }
}

function Protect-SemtLocalDirectoryAcl {
    <#
    Restricts a technician-side state folder to SYSTEM and Administrators (full control),
    optionally granting Users read access. Inheritance is removed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$AllowUsersRead
    )

    New-SemtDirectory -Path $Path
    if (-not [string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable('SEMT_RUNTIME_ROOT_OVERRIDE', 'Process'))) { return 'SkippedRuntimeOverride' }
    $elevated = $false
    try { $elevated = ([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator) } catch { $elevated = $false }
    if (-not $elevated) { return 'SkippedNotElevated' }
    $icacls = Join-Path -Path $env:SystemRoot -ChildPath 'System32\icacls.exe'
    if (-not (Test-Path -LiteralPath $icacls -PathType Leaf)) { return 'IcaclsNotFound' }
    $grants = @('*S-1-5-18:(OI)(CI)F','*S-1-5-32-544:(OI)(CI)F')
    if ($AllowUsersRead) { $grants += '*S-1-5-32-545:(OI)(CI)RX' }
    $output = & $icacls $Path '/inheritance:r' '/grant:r' $grants 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw ("ACL hardening failed for {0}: {1}" -f $Path,(($output | Out-String).Trim()))
    }
    return 'Protected'
}

function Get-SemtTechnicianIdentity {
    [CmdletBinding()]
    param()

    $identity = $null
    try { $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent() } catch { $identity = $null }
    $upn = ''
    try {
        $whoamiUpn = & whoami.exe /upn 2>$null
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$whoamiUpn)) { $upn = ([string]$whoamiUpn).Trim() }
    }
    catch { $upn = '' }
    $isAdmin = $false
    try { $isAdmin = ([System.Security.Principal.WindowsPrincipal]$identity).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator) } catch { $isAdmin = $false }
    $sessionId = -1
    try { $sessionId = [int][System.Diagnostics.Process]::GetCurrentProcess().SessionId } catch { $sessionId = -1 }

    return [pscustomobject]@{
        Account = if ($identity -and -not [string]::IsNullOrWhiteSpace($identity.Name)) { [string]$identity.Name } else { [Environment]::UserName }
        UserPrincipalName = $upn
        Sid = if ($identity -and $identity.User) { [string]$identity.User.Value } else { '' }
        AuthenticationType = if ($identity) { [string]$identity.AuthenticationType } else { '' }
        IsElevatedAdministrator = [bool]$isAdmin
        SessionId = $sessionId
        ComputerName = $env:COMPUTERNAME
    }
}

function Get-SemtProcessStartUtc {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][int]$ProcessId)

    try { return ((Get-Process -Id $ProcessId -ErrorAction Stop).StartTime.ToUniversalTime().ToString('o')) }
    catch { return '' }
}

function Test-SemtProcessAlive {
    <#
    True when the process exists and, when known, has the same start time (PID reuse guard).
    #>
    [CmdletBinding()]
    param(
        [int]$ProcessId,
        [AllowNull()][string]$ExpectedStartUtc
    )

    if ($ProcessId -le 0) { return $false }
    $process = $null
    try { $process = Get-Process -Id $ProcessId -ErrorAction Stop } catch { return $false }
    if ([string]::IsNullOrWhiteSpace($ExpectedStartUtc)) { return $true }
    $expected = ConvertTo-SemtUtcDateTime -Value $ExpectedStartUtc
    if ($null -eq $expected) { return $true }
    try {
        $actual = $process.StartTime.ToUniversalTime()
        return ([math]::Abs(($actual - $expected).TotalSeconds) -lt 2)
    }
    catch { return $true }
}

function Write-SemtLogLine {
    <#
    Appends one or more timestamped lines to every target file. Multi-line messages are split
    so each physical line keeps its own yyyy-MM-dd HH:mm:ss prefix.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Path,
        [AllowNull()][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )

    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $lines = @(([string]$Message) -split "`r?`n")
    foreach ($target in @($Path)) {
        if ([string]::IsNullOrWhiteSpace($target)) { continue }
        try {
            $parent = Split-Path -Path $target -Parent
            if (-not [string]::IsNullOrWhiteSpace($parent)) { New-SemtDirectory -Path $parent }
            $content = @($lines | ForEach-Object { '{0} [{1}] {2}' -f $stamp,$Level,$_ })
            Add-Content -LiteralPath $target -Value $content -Encoding UTF8 -ErrorAction Stop
        }
        catch { }
    }
}

function Resolve-SemtWindowsPowerShellPath {
    [CmdletBinding()]
    param()

    $windowsPowerShell = Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $windowsPowerShell -PathType Leaf) { return $windowsPowerShell }
    foreach ($commandName in @('powershell.exe','pwsh.exe')) {
        $command = Get-Command -Name $commandName -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) { return [string]$command.Source }
    }
    throw 'No supported PowerShell executable was found.'
}

function Test-SemtBooleanLikeTrue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -is [bool]) { return [bool]$Value }
    $text = ([string]$Value).Trim()
    return ($text -in @('True','true','TRUE','1','YES','Yes','yes','OUI','Oui','oui'))
}

function Get-SemtPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Default = ''
    )

    if ($null -eq $InputObject) { return $Default }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $Default
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $Default
}
