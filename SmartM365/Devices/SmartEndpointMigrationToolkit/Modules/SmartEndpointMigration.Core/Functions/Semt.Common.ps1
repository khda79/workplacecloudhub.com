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
        # Short temporary name: the final path length is not increased (260-character limit in PowerShell 5.1).
        $tempPath = Join-Path -Path $parent -ChildPath ('.{0}.tmp' -f [guid]::NewGuid().ToString('N').Substring(0, 12))
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA5nlaOe2e9hbN1
# 97G3zGyjlXXA73A2u079t0AAb3a0paCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID8AAiFgTz7dguJvSKTxv63CnYB0PU+mHzR96Cd0Fk0/MA0GCSqG
# SIb3DQEBAQUABIIBgDay5cST+dBAv2UC4O/Qiu2eN1jcuu58kYCvimtopy8KU0po
# 1bPeysA9rVL9lcr0Px0E5EJCZeS3Ymac67btmzCSNjTtc8SfD3AG6VQA0Mnntfur
# D4VVsAe49VOLMTyFgiRiJP2UD/OyKcZRHaZyKlFhjcEo9EvbxkNKjpn8IeFxVwCK
# /dLAYZr2I17SbZlCESu/h2pCx75YyJNZyfetJNJ2UPwgWBnPAdu8vhF7MdIUVDk+
# AzMMHbjHKDj1DU0X48HqMzUcDa731CMSXTk8GUrB4OHIz2uaq5m1aJM0luuGlzrR
# +cPZJH8nypZy14ncW2YomAF0yFLKSa7FON4C4dR5C0T6dsroMYuxRR8nE6pTkedp
# FjO+6K5UFM19hbuSW/kI4FlXmha6/A3Vr0tL8paIkWejvdz8iPQU3AY73dhku5iv
# VpYKtv4yWWysBt2nd3rc63zKShL22S3ne4lfJhihXoghg2r7QA8fWazdC0vwXGSt
# 5jHBN0sUXbtZQkmGi6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTFaMC8GCSqGSIb3DQEJBDEiBCAU3cncIQc3WQzr5exe10XRTZaRXqvIg5WjksMY
# xK9JITANBgkqhkiG9w0BAQEFAASCAgAuEzoS/HYM2jvA8t7pGMk2uNizdM9jz9+4
# ckIuzzOCPMFLgImzcNbhJUsFlC66RO5jM1YEx+ckIGUeVeki67LCsf/0fiw+54hT
# lXmDXaSIe1Ook+rXUfV1/mydpE/kAPevKaXeFkX3QPZo01dAQJFr17TLkomOq/gM
# fp0odg9cfCOa7P1fJVURlhPKO9xnKZnO19/j9KhBWgCoeb437QvsqLTJCWRTAJof
# vTPpiCJJ8X8zOxlCbKwWDbZcYQ40DQo9NdcAqULAfddzIXZvdeu9ibmgCaVFNb3+
# XEdcMAQePtPEnG2X2YZN9S72QFl9dN0rbitxG/ukLQqrlzCRjDdS5coenLs095+j
# 8FkVLc8M6vBPt/VFE11QD8WpGCBsY+TZCLcLKyzmRPqSk2zNNqFe7DRw9x02MT4g
# rZyIpghcug1qD1nlP+d5W1WQKR1L5nD7kQ3ecXuqnxIn96tGUygU/b9lbGwZ9VAP
# Mg1+ByhBLYEjm8IMJtYOV7AxVJqEX2y1ILiKCchTr1tQi4SsVm0P3jzUmeWh2P/l
# cbNxYFH/fcBB9+EeBUyMKG49Ya4RzCpLXFuFlXbrMYr3akXn8YPggmFW7pdS8U6h
# mJyU/YiI/zF+tiLEw+ubR78NbLxrTjVYzFNZqdxjws5jPyzPAeDN3y7BxIK4nnct
# QK6DWoKmWA==
# SIG # End signature block
