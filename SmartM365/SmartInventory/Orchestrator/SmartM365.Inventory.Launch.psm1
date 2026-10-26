<#
.SYNOPSIS
Resolves and runs SmartInventory jobs and manual launches from the jobs catalog.

.DESCRIPTION
The catalog is Orchestrator-Jobs.json.template: its Jobs array (orchestrator jobs) and its
Launches array (manual entries, "Manual": true). Launch fields are read from the committed
template; Arguments and Execution.Mode of a job may be overridden by a runtime jobs manifest.

Launch fields of an entry:
- Runtime: PowerShell7 or WindowsPowerShell (default: PowerShellEdition, then PowerShell7).
- Execution: { Mode = Direct|LocalCache|Auto (optional), CacheName, Include = [...] }.
  Only entries that declare Execution can run from a local cache.
- LaunchArguments (jobs): fixed arguments that the legacy launcher added before the job Arguments.
- WorkingDirectory: folder relative to SmartInventory, used as the current directory.
- Interactive: the entry needs a user console (interactive sign-in or prompts).
- RequiresConfirmation: the user must type WRITE before the script starts.
- Tokens: { Name = { Value; MustExist } } used as {{Name}} in Arguments.
- LegacyLaunchers: the .cmd files this entry replaces (relative to SmartInventory).
- PassTenant: false when the script asks for its tenant itself (default: -Tenant is passed
  when the script declares it).
- HideInGui: true to keep a manual entry out of the orchestrator GUI (the GUI already offers it
  as a button, or it is a setup tool); it stays in the SmartInventory.cmd menu and command line.

LocalCache reproduces the legacy on-premises launchers: first usable folder among
ProgramData\SmartM365\LauncherCache\<CacheName>, Windows\Temp, C:\Temp and %TEMP% (the orchestrator
uses only ProgramData and %TEMP%); robocopy /MIR of the script folder, Config and
Modules\SmartM365.Core; failure from exit code 8; PrepareCache.log; Include files checked before and
after the copy; Unblock-File. Outputs that resolve inside the cache are refused.
The cache is hardened: a folder or parent that is a junction or link, or that is owned by another
account than SYSTEM, Administrators, TrustedInstaller or the launching account, is refused; the
cache folder gets an explicit ACL (SYSTEM, Administrators and the launching account only) and is
emptied; right before the launch, every code file of the cache must be byte-identical to the source.

Compatible with Windows PowerShell 5.1 and PowerShell 7.

.VERSION
1.4.0
#>
Set-StrictMode -Version 2.0

$script:LaunchModuleVersion = '1.4.0'
$script:ValidRuntimes = @('PowerShell7', 'WindowsPowerShell')
$script:ValidModes = @('Direct', 'LocalCache', 'Auto')
# Test hook: replaced by the offline tests to simulate the zone of a share.
$script:PathZoneResolver = $null
$script:LaunchLogPath = ''
# Test hook: replaced by the offline tests to simulate the owner (SID) of a cache folder.
$script:OwnerResolver = $null
$script:CodeFileExtensions = @('.ps1', '.psm1', '.psd1', '.ps1xml', '.dll', '.exe', '.cs', '.cmd', '.bat', '.vbs', '.js')

function Write-SmartM365InventoryLaunchMessage {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $color = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } default { 'Gray' } }
    foreach ($line in ($Message -split "\r?\n")) {
        $text = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $line
        # Orchestrated launches write into the job log instead of the hidden console.
        if ($script:LaunchLogPath) { [IO.File]::AppendAllText($script:LaunchLogPath, $text + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false))) }
        else { Microsoft.PowerShell.Utility\Write-Host $text -ForegroundColor $color }
    }
}

function Get-SmartM365InventoryProperty {
    param($Object, [Parameter(Mandatory)][string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] }
        return $Default
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) { return $property.Value }
    return $Default
}

function Get-SmartM365InventoryRoots {
    [CmdletBinding()]
    param([string]$SmartInventoryRoot = '')
    if ([string]::IsNullOrWhiteSpace($SmartInventoryRoot)) { $SmartInventoryRoot = Split-Path -Path $PSScriptRoot -Parent }
    $inventoryRoot = [IO.Path]::GetFullPath($SmartInventoryRoot).TrimEnd('\', '/')
    [pscustomobject]@{
        SmartInventoryRoot = $inventoryRoot
        SmartM365Root = (Split-Path -Path $inventoryRoot -Parent)
        TemplatePath = (Join-Path -Path $inventoryRoot -ChildPath 'Orchestrator\Orchestrator-Jobs.json.template')
    }
}

function Read-SmartM365InventoryJsonFile {
    param([Parameter(Mandatory)][string]$Path)
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    return (ConvertFrom-Json -InputObject $text)
}

function Get-SmartM365InventoryCatalog {
    <#
    Loads the catalog template and, when given, the runtime manifest whose job Arguments and
    Execution.Mode override the template values.
    #>
    [CmdletBinding()]
    param(
        [string]$SmartInventoryRoot = '',
        [AllowEmptyString()][string]$OverlayPath = ''
    )
    $roots = Get-SmartM365InventoryRoots -SmartInventoryRoot $SmartInventoryRoot
    if (-not (Test-Path -LiteralPath $roots.TemplatePath -PathType Leaf)) { throw "Jobs catalog not found: $($roots.TemplatePath)" }
    $template = Read-SmartM365InventoryJsonFile -Path $roots.TemplatePath
    $overlayJobs = @{}
    if (-not [string]::IsNullOrWhiteSpace($OverlayPath)) {
        if (-not (Test-Path -LiteralPath $OverlayPath -PathType Leaf)) { throw "Jobs manifest not found: $OverlayPath" }
        $overlay = Read-SmartM365InventoryJsonFile -Path $OverlayPath
        foreach ($job in @(Get-SmartM365InventoryProperty -Object $overlay -Name 'Jobs' -Default @())) {
            $overlayJobs[[string]$job.Name] = $job
        }
    }

    $entries = New-Object System.Collections.Generic.List[object]
    $names = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($job in @(Get-SmartM365InventoryProperty -Object $template -Name 'Jobs' -Default @())) {
        $name = [string]$job.Name
        if (-not $names.Add($name)) { throw "Duplicate catalog name: $name" }
        $overlayJob = $null
        if ($overlayJobs.ContainsKey($name)) { $overlayJob = $overlayJobs[$name] }
        $entries.Add([pscustomobject]@{ Name = $name; Kind = 'Job'; Definition = $job; Overlay = $overlayJob })
    }
    foreach ($launch in @(Get-SmartM365InventoryProperty -Object $template -Name 'Launches' -Default @())) {
        $name = [string]$launch.Name
        if (-not $names.Add($name)) { throw "Duplicate catalog name: $name" }
        if (-not [bool](Get-SmartM365InventoryProperty -Object $launch -Name 'Manual' -Default $false)) { throw "Launches entry '$name' must declare ""Manual"": true." }
        $entries.Add([pscustomobject]@{ Name = $name; Kind = 'Manual'; Definition = $launch; Overlay = $null })
    }
    [pscustomobject]@{
        SmartInventoryRoot = $roots.SmartInventoryRoot
        SmartM365Root = $roots.SmartM365Root
        TemplatePath = $roots.TemplatePath
        OverlayPath = $OverlayPath
        Entries = $entries.ToArray()
    }
}

function Find-SmartM365InventoryCatalogEntry {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Catalog, [Parameter(Mandatory)][string]$Name)
    $match = @($Catalog.Entries | Where-Object { $_.Name -ieq $Name })
    if ($match.Count -eq 0) { throw "Unknown job or launch '$Name'. Use -List to show the catalog." }
    return $match[0]
}

function ConvertTo-SmartM365InventoryArgumentList {
    <# Splits a catalog Arguments string; double quotes group a value that contains spaces. #>
    [CmdletBinding()]
    param([AllowEmptyString()][AllowNull()][string]$Arguments)
    $result = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Arguments)) { return }
    foreach ($match in [regex]::Matches($Arguments, '"(?<Quoted>[^"]*)"|(?<Plain>[^\s"]+)')) {
        if ($match.Groups['Quoted'].Success) { $result.Add($match.Groups['Quoted'].Value) } else { $result.Add($match.Groups['Plain'].Value) }
    }
    return $result.ToArray()
}

function ConvertTo-SmartM365InventoryCommandLineArgument {
    <# Quotes one argument with the CommandLineToArgvW rules (backslashes before a quote are doubled). #>
    param([AllowEmptyString()][string]$Value)
    if ($Value.Length -gt 0 -and $Value -notmatch '[\s"]') { return $Value }
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $backslashes = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq '\') { $backslashes++; continue }
        if ($character -eq '"') {
            [void]$builder.Append([char]'\', (2 * $backslashes) + 1)
            [void]$builder.Append('"')
        }
        else {
            if ($backslashes -gt 0) { [void]$builder.Append([char]'\', $backslashes) }
            [void]$builder.Append($character)
        }
        $backslashes = 0
    }
    if ($backslashes -gt 0) { [void]$builder.Append([char]'\', 2 * $backslashes) }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Join-SmartM365InventoryCommandLine {
    param([AllowEmptyCollection()][string[]]$Arguments)
    return ((@($Arguments) | ForEach-Object { ConvertTo-SmartM365InventoryCommandLineArgument -Value $_ }) -join ' ')
}

function Get-SmartM365InventoryScriptParameters {
    <# Names declared in the param block of a script (AST, nothing is executed). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and @($parseErrors).Count -gt 0) { throw "Unable to parse $Path : $(@($parseErrors)[0].Message)" }
    if ($null -eq $ast.ParamBlock) { return }
    return @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
}

function Resolve-SmartM365InventoryTenant {
    <# Tenant: explicit value, then SMARTM365_ORCHESTRATOR_TENANT, then DefaultTenant of the global configuration. #>
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Tenant = '', [Parameter(Mandatory)][string]$SmartM365Root)
    if (-not [string]::IsNullOrWhiteSpace($Tenant)) { return [pscustomobject]@{ Value = $Tenant.Trim(); Source = 'Parameter' } }
    $environmentTenant = [Environment]::GetEnvironmentVariable('SMARTM365_ORCHESTRATOR_TENANT')
    if (-not [string]::IsNullOrWhiteSpace($environmentTenant)) { return [pscustomobject]@{ Value = $environmentTenant.Trim(); Source = 'SMARTM365_ORCHESTRATOR_TENANT' } }
    $value = Get-SmartM365InventoryGlobalSetting -SmartM365Root $SmartM365Root -Name 'DefaultTenant'
    if ([string]::IsNullOrWhiteSpace([string]$value)) { throw 'No tenant: pass -Tenant or set DefaultTenant in SmartM365.global.local.json.' }
    return [pscustomobject]@{ Value = ([string]$value).Trim(); Source = 'DefaultTenant' }
}

function Get-SmartM365InventoryGlobalSetting {
    <# Reads one value of SmartM365.global.local.json (or its .json.txt transport name), then of its template. Read-only. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SmartM365Root, [Parameter(Mandatory)][string]$Name)
    $configFolder = Join-Path -Path $SmartM365Root -ChildPath 'Config'
    foreach ($fileName in @('SmartM365.global.local.json.txt', 'SmartM365.global.local.json', 'SmartM365.global.local.json.template')) {
        $path = Join-Path -Path $configFolder -ChildPath $fileName
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $value = Get-SmartM365InventoryProperty -Object (Read-SmartM365InventoryJsonFile -Path $path) -Name $Name
        if ($null -ne $value -and -not ($value -is [string] -and ([string]::IsNullOrWhiteSpace($value) -or $value -in @('__USE_GLOBAL__', 'USE_GLOBAL')))) { return $value }
    }
    return $null
}

function Get-SmartM365InventoryPathZone {
    <# Internet Explorer security zone of a path: MyComputer, Intranet, Trusted, Internet or Restricted. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if ($null -ne $script:PathZoneResolver) { return [string](& $script:PathZoneResolver $Path) }
    if (-not ('SmartM365.InventoryLaunch.ZoneMapper' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace SmartM365.InventoryLaunch {
    [ComImport, Guid("79EAC9EE-BAF9-11CE-8C82-00AA004BA90B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IInternetSecurityManager {
        [PreserveSig] int SetSecuritySite(IntPtr site);
        [PreserveSig] int GetSecuritySite(out IntPtr site);
        [PreserveSig] int MapUrlToZone([MarshalAs(UnmanagedType.LPWStr)] string url, out uint zone, uint flags);
    }
    public static class ZoneMapper {
        [DllImport("urlmon.dll")]
        static extern int CoInternetCreateSecurityManager(IntPtr provider, [MarshalAs(UnmanagedType.Interface)] out object manager, uint reserved);
        public static int MapPath(string path) {
            object raw;
            int hr = CoInternetCreateSecurityManager(IntPtr.Zero, out raw, 0);
            if (hr != 0) { Marshal.ThrowExceptionForHR(hr); }
            IInternetSecurityManager manager = (IInternetSecurityManager)raw;
            uint zone;
            hr = manager.MapUrlToZone(new Uri(path).AbsoluteUri, out zone, 0);
            if (hr != 0) { Marshal.ThrowExceptionForHR(hr); }
            return (int)zone;
        }
    }
}
'@
    }
    $zone = [SmartM365.InventoryLaunch.ZoneMapper]::MapPath([IO.Path]::GetFullPath($Path))
    switch ($zone) { 0 { 'MyComputer' } 1 { 'Intranet' } 2 { 'Trusted' } 3 { 'Internet' } 4 { 'Restricted' } default { "Zone$zone" } }
}

function Resolve-SmartM365InventoryExecutionMode {
    <#
    Mode: parameter, then Execution.Mode of the job (runtime manifest, then template), then
    OnPremExecutionMode of the global configuration (default LocalCache). An entry without
    Execution always runs Direct. Auto runs Direct when the share is in the Intranet or Trusted
    zone (or is local), LocalCache otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Entry,
        [Parameter(Mandatory)]$Catalog,
        [AllowEmptyString()][string]$ExecutionMode = '',
        [AllowEmptyString()][string]$DefaultOnPremExecutionMode = '',
        [string]$ExecutionModeSource = 'Parameter'
    )
    $execution = Get-SmartM365InventoryProperty -Object $Entry.Definition -Name 'Execution'
    $requested = ''; $source = ''
    if (-not [string]::IsNullOrWhiteSpace($ExecutionMode)) { $requested = $ExecutionMode; $source = $ExecutionModeSource }
    if (-not $requested) {
        $overlayExecution = Get-SmartM365InventoryProperty -Object $Entry.Overlay -Name 'Execution'
        $overlayMode = [string](Get-SmartM365InventoryProperty -Object $overlayExecution -Name 'Mode' -Default '')
        if ($overlayMode) { $requested = $overlayMode; $source = 'Job (runtime manifest)' }
    }
    if (-not $requested -and $null -ne $execution) {
        $templateMode = [string](Get-SmartM365InventoryProperty -Object $execution -Name 'Mode' -Default '')
        if ($templateMode) { $requested = $templateMode; $source = 'Job' }
    }
    if (-not $requested -and $null -eq $execution) { return [pscustomobject]@{ Mode = 'Direct'; Requested = 'Direct'; Source = 'No Execution definition'; Zone = '' } }
    if (-not $requested) {
        $requested = $DefaultOnPremExecutionMode
        if ([string]::IsNullOrWhiteSpace($requested)) { $requested = [string](Get-SmartM365InventoryGlobalSetting -SmartM365Root $Catalog.SmartM365Root -Name 'OnPremExecutionMode') }
        if ([string]::IsNullOrWhiteSpace($requested)) { $requested = 'LocalCache' }
        $source = 'OnPremExecutionMode'
    }
    $requested = @($script:ValidModes | Where-Object { $_ -ieq $requested }) | Select-Object -First 1
    if (-not $requested) { throw "Entry '$($Entry.Name)': execution mode must be Direct, LocalCache or Auto." }
    if ($requested -ne 'Direct' -and $null -eq $execution) { throw "Entry '$($Entry.Name)' has no Execution definition and can only run Direct." }
    if ($requested -ne 'Auto') { return [pscustomobject]@{ Mode = $requested; Requested = $requested; Source = $source; Zone = '' } }
    $zone = Get-SmartM365InventoryPathZone -Path $Catalog.SmartInventoryRoot
    $mode = if ($zone -in @('MyComputer', 'Intranet', 'Trusted')) { 'Direct' } else { 'LocalCache' }
    return [pscustomobject]@{ Mode = $mode; Requested = 'Auto'; Source = $source; Zone = $zone }
}

function Resolve-SmartM365InventoryLaunch {
    <#
    Builds the launch plan of a catalog entry without side effects: engine, script, ordered
    arguments, tenant, working directory, execution mode and cache definition.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Catalog,
        [Parameter(Mandatory)][string[]]$Name,
        [ValidateSet('Now', 'Request')][string]$Action = 'Now',
        [AllowEmptyString()][string]$Tenant = '',
        [AllowEmptyString()][string]$ExecutionMode = '',
        [string]$ExecutionModeSource = 'Parameter',
        [AllowEmptyString()][string]$DefaultOnPremExecutionMode = '',
        [AllowEmptyCollection()][string[]]$ExtraArguments = @(),
        # Orchestrator: the Arguments of its (runtime) job replace the catalog Arguments.
        [AllowNull()]$JobArguments = $null,
        # Orchestrator started without -Connect: the job script does not get -Connect either.
        [switch]$NoAutoConnect
    )
    $names = @($Name | ForEach-Object { ([string]$_) -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($names.Count -eq 0) { throw 'A job or launch name is required.' }
    $entries = @($names | ForEach-Object { Find-SmartM365InventoryCatalogEntry -Catalog $Catalog -Name $_ })
    $tenantInfo = Resolve-SmartM365InventoryTenant -Tenant $Tenant -SmartM365Root $Catalog.SmartM365Root
    $extra = @($ExtraArguments | Where-Object { $null -ne $_ })

    if ($Action -eq 'Request') {
        foreach ($entry in $entries) {
            if ($entry.Kind -ne 'Job') { throw "'$($entry.Name)' is a manual launch: it cannot be requested through the orchestrators." }
        }
        $pipelineScript = Join-Path -Path $Catalog.SmartInventoryRoot -ChildPath 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        $arguments = @('-Job', ((@($entries | ForEach-Object { $_.Name })) -join ','), '-Collect')
        return [pscustomobject]@{
            Name = ((@($entries | ForEach-Object { $_.Name })) -join ',')
            Kind = 'Request'; Action = 'Request'; Runtime = 'PowerShell7'
            ScriptPath = $pipelineScript; ScriptRelativePath = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
            Arguments = $arguments; ExtraArguments = $extra
            Tenant = $tenantInfo.Value; TenantSource = $tenantInfo.Source; PassTenant = $true
            WorkingDirectory = ''; Mode = 'Direct'; ModeInfo = $null; CacheName = ''; Include = @()
            Interactive = $false; RequiresConfirmation = $false; Description = 'Request through the resident orchestrators.'
            Tokens = @{}; Entry = $null
        }
    }
    if ($entries.Count -gt 1) { throw 'Run one job at a time with -Now; use -Request for several jobs.' }
    $entry = $entries[0]
    $definition = $entry.Definition
    $scriptRelativePath = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'ScriptPath' -Default '')
    if (-not $scriptRelativePath) { throw "Entry '$($entry.Name)' has no ScriptPath." }
    $scriptPath = [IO.Path]::GetFullPath((Join-Path -Path $Catalog.SmartInventoryRoot -ChildPath $scriptRelativePath))

    $runtime = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'Runtime' -Default '')
    if (-not $runtime) { $runtime = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'PowerShellEdition' -Default 'PowerShell7') }
    if ($runtime -notin $script:ValidRuntimes) { throw "Entry '$($entry.Name)': Runtime must be PowerShell7 or WindowsPowerShell." }

    $modeInfo = Resolve-SmartM365InventoryExecutionMode -Entry $entry -Catalog $Catalog -ExecutionMode $ExecutionMode -ExecutionModeSource $ExecutionModeSource -DefaultOnPremExecutionMode $DefaultOnPremExecutionMode
    $execution = Get-SmartM365InventoryProperty -Object $definition -Name 'Execution'
    $cacheName = [string](Get-SmartM365InventoryProperty -Object $execution -Name 'CacheName' -Default '')
    $include = @(Get-SmartM365InventoryProperty -Object $execution -Name 'Include' -Default @() | ForEach-Object { [string]$_ })

    # Job arguments: legacy launcher arguments, then the runtime (or template) Arguments.
    $argumentText = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'Arguments' -Default '')
    if ($null -ne $entry.Overlay -and $null -ne $entry.Overlay.PSObject.Properties['Arguments']) { $argumentText = [string]$entry.Overlay.Arguments }
    if ($null -ne $JobArguments) { $argumentText = [string]$JobArguments }
    $arguments = @()
    if ($entry.Kind -eq 'Job') { $arguments += @(ConvertTo-SmartM365InventoryArgumentList -Arguments ([string](Get-SmartM365InventoryProperty -Object $definition -Name 'LaunchArguments' -Default ''))) }
    $arguments += @(ConvertTo-SmartM365InventoryArgumentList -Arguments $argumentText)

    $parameters = @()
    if (Test-Path -LiteralPath $scriptPath -PathType Leaf) { $parameters = @(Get-SmartM365InventoryScriptParameters -Path $scriptPath) }
    # Like the orchestrator, a job connects when its script declares -Connect.
    if ($entry.Kind -eq 'Job' -and -not $NoAutoConnect -and $parameters -contains 'Connect' -and -not (@($arguments + $extra) | Where-Object { $_ -match '^-Connect(:|$)' })) { $arguments += '-Connect' }

    $tokens = @{}
    $tokenDefinitions = Get-SmartM365InventoryProperty -Object $definition -Name 'Tokens'
    if ($null -ne $tokenDefinitions) {
        foreach ($property in $tokenDefinitions.PSObject.Properties) {
            $value = $property.Value
            if ($value -is [string]) { $value = [pscustomobject]@{ Value = $value } }
            $tokens[$property.Name] = [pscustomobject]@{
                Value = [string](Get-SmartM365InventoryProperty -Object $value -Name 'Value' -Default '')
                MustExist = [bool](Get-SmartM365InventoryProperty -Object $value -Name 'MustExist' -Default $false)
            }
        }
    }

    [pscustomobject]@{
        Name = $entry.Name; Kind = $entry.Kind; Action = 'Now'; Runtime = $runtime
        ScriptPath = $scriptPath; ScriptRelativePath = $scriptRelativePath
        Arguments = $arguments; ExtraArguments = $extra
        Tenant = $tenantInfo.Value; TenantSource = $tenantInfo.Source; PassTenant = (($parameters -contains 'Tenant') -and [bool](Get-SmartM365InventoryProperty -Object $definition -Name 'PassTenant' -Default $true))
        WorkingDirectory = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'WorkingDirectory' -Default '')
        Mode = $modeInfo.Mode; ModeInfo = $modeInfo; CacheName = $cacheName; Include = $include
        Interactive = [bool](Get-SmartM365InventoryProperty -Object $definition -Name 'Interactive' -Default $false)
        RequiresConfirmation = [bool](Get-SmartM365InventoryProperty -Object $definition -Name 'RequiresConfirmation' -Default $false)
        Description = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'Description' -Default '')
        Tokens = $tokens; Entry = $entry
    }
}

function Resolve-SmartM365InventoryTokenValue {
    param([AllowEmptyString()][string]$Text, [hashtable]$Values)
    $resolved = $Text
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            if (-not $Values.ContainsKey($tokenName)) { throw "Argument token '$($match.Value)' is not defined." }
            $resolved = $resolved.Replace($match.Value, [string]$Values[$tokenName])
        }
    }
    return $resolved
}

function Get-SmartM365InventoryEnginePath {
    <# pwsh: ProgramFiles, ProgramFiles(x86), then PATH. powershell.exe 5.1: Sysnative, then System32. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('PowerShell7', 'WindowsPowerShell')][string]$Runtime)
    $checked = New-Object System.Collections.Generic.List[string]
    if ($Runtime -eq 'WindowsPowerShell') {
        foreach ($candidate in @((Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'), (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'))) {
            $checked.Add($candidate)
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
        }
        throw ("Windows PowerShell 5.1 (powershell.exe) was not found. Checked:{0}  {1}" -f [Environment]::NewLine, ($checked -join ([Environment]::NewLine + '  ')))
    }
    $candidates = @((Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'))
    if (${env:ProgramFiles(x86)}) { $candidates += (Join-Path ${env:ProgramFiles(x86)} 'PowerShell\7\pwsh.exe') }
    foreach ($candidate in $candidates) {
        $checked.Add($candidate)
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    $checked.Add('PATH')
    $command = Get-Command -Name 'pwsh.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $command) { return $command.Source }
    throw ("PowerShell 7 (pwsh.exe) was not found. Install PowerShell 7 or add pwsh.exe to PATH. Checked:{0}  {1}" -f [Environment]::NewLine, ($checked -join ([Environment]::NewLine + '  ')))
}

function Get-SmartM365InventoryCacheCandidates {
    <# Cache folders in order. The orchestrator (-Orchestrated) skips the shared Windows\Temp and C:\Temp. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CacheName, [switch]$Orchestrated)
    $candidates = @((Join-Path $env:ProgramData ('SmartM365\LauncherCache\' + $CacheName)))
    if (-not $Orchestrated) { $candidates += @((Join-Path $env:SystemRoot ('Temp\SmartM365LauncherCache\' + $CacheName)), ('C:\Temp\SmartM365LauncherCache\' + $CacheName)) }
    if ($env:TEMP) { $candidates += (Join-Path $env:TEMP ('SmartM365LauncherCache\' + $CacheName)) }
    return $candidates
}

function Get-SmartM365InventoryTrustedOwnerSids {
    # SYSTEM, Administrators, TrustedInstaller and the account that launches.
    return @('S-1-5-18', 'S-1-5-32-544', 'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464', [Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
}

function Get-SmartM365InventoryFolderOwner {
    param([Parameter(Mandatory)][string]$Path)
    if ($null -ne $script:OwnerResolver) { return [string](& $script:OwnerResolver $Path) }
    return (Get-Acl -LiteralPath $Path).GetOwner([Security.Principal.SecurityIdentifier]).Value
}

function Test-SmartM365InventoryCacheLocation {
    <#
    Returns why a cache folder cannot be used, or an empty string: the folder and its existing
    parents (drive root excepted) must not be junctions or links and must be owned by a trusted
    account, otherwise another account could replace the files between the copy and the launch.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $trusted = Get-SmartM365InventoryTrustedOwnerSids
    $directory = New-Object System.IO.DirectoryInfo ([IO.Path]::GetFullPath($Path))
    while ($null -ne $directory -and $null -ne $directory.Parent) {
        if (Test-Path -LiteralPath $directory.FullName) {
            $item = Get-Item -LiteralPath $directory.FullName -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return ('{0} is a junction or a link' -f $directory.FullName) }
            $owner = Get-SmartM365InventoryFolderOwner -Path $directory.FullName
            if ($owner -notin $trusted) { return ('{0} is owned by {1}' -f $directory.FullName, $owner) }
        }
        $directory = $directory.Parent
    }
    return ''
}

function Protect-SmartM365InventoryCacheFolder {
    <# Explicit ACL without inheritance: SYSTEM, Administrators and the launching account only. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $acl = New-Object System.Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true, $false)
    $inheritance = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
    foreach ($sid in @('S-1-5-18', 'S-1-5-32-544', [Security.Principal.WindowsIdentity]::GetCurrent().User.Value)) {
        $identity = New-Object System.Security.Principal.SecurityIdentifier $sid
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($identity, 'FullControl', $inheritance, 'None', 'Allow')))
    }
    # Only the modified DACL is written (Set-Acl would also write the SACL, which needs SeSecurityPrivilege).
    $directory = New-Object System.IO.DirectoryInfo $Path
    if ('System.IO.FileSystemAclExtensions' -as [type]) { [System.IO.FileSystemAclExtensions]::SetAccessControl($directory, $acl) }
    else { $directory.SetAccessControl($acl) }
}

function Assert-SmartM365InventoryCacheIntegrity {
    <# Every code file of a cached copy (file or folder) must be byte-identical to its source. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination)
    if (Test-Path -LiteralPath $Source -PathType Leaf) {
        $pairs = @([pscustomobject]@{ Relative = (Split-Path $Source -Leaf); Source = $Source; Destination = $Destination })
        if (-not (Test-Path -LiteralPath $Destination -PathType Leaf)) { throw "Cache integrity: $Destination is missing." }
    }
    else {
        $list = {
            param([string]$Root)
            $full = [IO.Path]::GetFullPath($Root).TrimEnd('\')
            $map = @{}
            foreach ($file in @(Get-ChildItem -LiteralPath $full -File -Recurse -Force -ErrorAction Stop)) {
                if ($script:CodeFileExtensions -contains $file.Extension.ToLowerInvariant()) { $map[$file.FullName.Substring($full.Length + 1).ToLowerInvariant()] = $file.FullName }
            }
            return $map
        }
        $sourceFiles = & $list $Source
        $cachedFiles = & $list $Destination
        foreach ($relative in @($cachedFiles.Keys)) { if (-not $sourceFiles.ContainsKey($relative)) { throw "Cache integrity: unexpected code file in the cache: $($cachedFiles[$relative])" } }
        $pairs = foreach ($relative in @($sourceFiles.Keys)) {
            if (-not $cachedFiles.ContainsKey($relative)) { throw "Cache integrity: $relative is missing from $Destination." }
            [pscustomobject]@{ Relative = $relative; Source = $sourceFiles[$relative]; Destination = $cachedFiles[$relative] }
        }
    }
    foreach ($pair in @($pairs)) {
        if ((Get-FileHash -LiteralPath $pair.Source -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $pair.Destination -Algorithm SHA256).Hash) {
            throw "Cache integrity: $($pair.Destination) differs from its source $($pair.Source)."
        }
    }
}

function Test-SmartM365InventoryPathInside {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root)
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    return ($full.Equals($rootFull, [StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($rootFull + '\', [StringComparison]::OrdinalIgnoreCase))
}

function Invoke-SmartM365InventoryProcess {
    <# Runs an executable in the current console and returns its exit code. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FilePath, [AllowEmptyCollection()][string[]]$Arguments = @(), [string]$WorkingDirectory = '')
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FilePath
    $startInfo.Arguments = Join-SmartM365InventoryCommandLine -Arguments $Arguments
    $startInfo.UseShellExecute = $false
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $startInfo.WorkingDirectory = $WorkingDirectory }
    elseif ((Get-Location).Provider.Name -eq 'FileSystem') { $startInfo.WorkingDirectory = (Get-Location).ProviderPath }
    $process = [System.Diagnostics.Process]::Start($startInfo)
    try { $process.WaitForExit(); return $process.ExitCode }
    finally { $process.Dispose() }
}

function Invoke-SmartM365InventoryRobocopy {
    # Same options as the legacy launchers; the console output (the log file line) is dropped,
    # everything is in PrepareCache.log.
    param([string]$Source, [string]$Destination, [string[]]$Options, [string]$LogOption)
    $arguments = @($Source.TrimEnd('\'), $Destination.TrimEnd('\')) + @($Options) + @('/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP', $LogOption)
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Join-Path $env:SystemRoot 'System32\robocopy.exe'
    $startInfo.Arguments = Join-SmartM365InventoryCommandLine -Arguments $arguments
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.CreateNoWindow = $true
    $process = [System.Diagnostics.Process]::Start($startInfo)
    try {
        [void]$process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        return $process.ExitCode
    }
    finally { $process.Dispose() }
}

function Initialize-SmartM365InventoryLocalCache {
    <#
    Prepares the local runtime cache exactly like the legacy on-premises launchers and returns
    the cache folders. Throws on any failure; nothing is launched.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Plan,
        [Parameter(Mandatory)]$Catalog,
        [string[]]$CacheBaseCandidates = @(),
        [switch]$Orchestrated
    )
    if (-not $Plan.CacheName) { throw "Entry '$($Plan.Name)' declares no Execution.CacheName." }
    if ($CacheBaseCandidates.Count -eq 0) { $CacheBaseCandidates = Get-SmartM365InventoryCacheCandidates -CacheName $Plan.CacheName -Orchestrated:$Orchestrated }
    $cacheBase = ''
    $refusals = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in $CacheBaseCandidates) {
        try {
            $reason = Test-SmartM365InventoryCacheLocation -Path $candidate
            if ($reason) { $refusals.Add(('{0}: {1}' -f $candidate, $reason)); continue }
            if ((Test-Path -LiteralPath $candidate) -and -not (Test-Path -LiteralPath $candidate -PathType Container)) { $refusals.Add(('{0}: not a folder' -f $candidate)); continue }
            New-Item -ItemType Directory -Path $candidate -Force -ErrorAction Stop | Out-Null
            # Checked again once created: another account may have created it in the meantime.
            $reason = Test-SmartM365InventoryCacheLocation -Path $candidate
            if ($reason) { $refusals.Add(('{0}: {1}' -f $candidate, $reason)); continue }
            Protect-SmartM365InventoryCacheFolder -Path $candidate
            # The cache is rebuilt for every run (same as rmdir /s /q in the legacy launchers), so nothing
            # left by an earlier run, or by another account before the ACL was set, is reused.
            foreach ($child in @(Get-ChildItem -LiteralPath $candidate -Force -ErrorAction Stop)) {
                if ($child.PSIsContainer -and ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) { & (Join-Path $env:SystemRoot 'System32\cmd.exe') /d /c rmdir /s /q "$($child.FullName)" 2>$null | Out-Null }
                else { Remove-Item -LiteralPath $child.FullName -Force -Recurse -ErrorAction SilentlyContinue }
            }
            if (@(Get-ChildItem -LiteralPath $candidate -Force -ErrorAction Stop).Count -gt 0) { $refusals.Add(('{0}: could not be emptied' -f $candidate)); continue }
            $probe = Join-Path $candidate '.smartm365-cache-write-test'
            [IO.File]::WriteAllText($probe, '')
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
            $cacheBase = $candidate
            break
        }
        catch { $refusals.Add(('{0}: {1}' -f $candidate, $_.Exception.Message)); continue }
    }
    if (-not $cacheBase) { throw ("Failed to prepare a safe local cache base folder. Tried:{0}  {1}" -f [Environment]::NewLine, ($refusals -join ([Environment]::NewLine + '  '))) }
    if ($refusals.Count -gt 0) { Write-SmartM365InventoryLaunchMessage -Level WARN -Message ("Cache folder(s) skipped:{0}  {1}" -f [Environment]::NewLine, ($refusals -join ([Environment]::NewLine + '  '))) }

    $smartM365Root = $Catalog.SmartM365Root
    $scriptFolder = Split-Path -Path $Plan.ScriptPath -Parent
    if (-not (Test-SmartM365InventoryPathInside -Path $scriptFolder -Root $smartM365Root)) { throw "Script folder is outside SmartM365: $scriptFolder" }
    $scriptFolderRelative = $scriptFolder.Substring($smartM365Root.TrimEnd('\').Length).TrimStart('\')
    $localRoot = Join-Path $cacheBase 'SmartM365'
    $localScriptFolder = Join-Path $localRoot $scriptFolderRelative
    $cacheLog = Join-Path $cacheBase 'PrepareCache.log'

    foreach ($folder in @($localScriptFolder, (Join-Path $localRoot 'Config'), (Join-Path $localRoot 'Modules\SmartM365.Core'))) {
        New-Item -ItemType Directory -Path $folder -Force -ErrorAction SilentlyContinue | Out-Null
    }
    Write-SmartM365InventoryLaunchMessage -Message ("Preparing local SmartM365 PowerShell runtime cache.{0}Source script folder: {1}{0}Local cache folder: {2}" -f [Environment]::NewLine, $scriptFolder, $localRoot)

    $copies = @(
        [pscustomobject]@{ Label = 'script'; Source = $scriptFolder; Destination = $localScriptFolder; Log = ('/LOG:' + $cacheLog) },
        [pscustomobject]@{ Label = 'Config'; Source = (Join-Path $smartM365Root 'Config'); Destination = (Join-Path $localRoot 'Config'); Log = ('/LOG+:' + $cacheLog) },
        [pscustomobject]@{ Label = 'SmartM365.Core'; Source = (Join-Path $smartM365Root 'Modules\SmartM365.Core'); Destination = (Join-Path $localRoot 'Modules\SmartM365.Core'); Log = ('/LOG+:' + $cacheLog) }
    )
    foreach ($copy in $copies) {
        $exitCode = Invoke-SmartM365InventoryRobocopy -Source $copy.Source -Destination $copy.Destination -Options @('/MIR') -LogOption $copy.Log
        if ($exitCode -ge 8) { throw ("Failed to prepare local {0} cache from: {1}; robocopy exit code: {2}; robocopy log: {3}" -f $copy.Label, $copy.Source, $exitCode, $cacheLog) }
    }

    $unblockFolders = New-Object System.Collections.Generic.List[string]
    $unblockFolders.Add($localScriptFolder)
    $integrity = New-Object System.Collections.Generic.List[object]
    foreach ($copy in $copies) { $integrity.Add([pscustomobject]@{ Source = $copy.Source; Destination = $copy.Destination }) }
    foreach ($item in $Plan.Include) {
        $relative = ([string]$item).Trim().TrimStart('\', '/')
        $source = Join-Path $smartM365Root $relative
        $destination = Join-Path $localRoot $relative
        if (-not (Test-Path -LiteralPath $source)) { throw ("Include source missing: {0}; destination: {1}; cache log: {2}" -f $source, $destination, $cacheLog) }
        if (Test-Path -LiteralPath $source -PathType Container) {
            $exitCode = Invoke-SmartM365InventoryRobocopy -Source $source -Destination $destination -Options @('/MIR') -LogOption ('/LOG+:' + $cacheLog)
            $unblockFolders.Add($destination)
        }
        else {
            $exitCode = Invoke-SmartM365InventoryRobocopy -Source (Split-Path $source -Parent) -Destination (Split-Path $destination -Parent) -Options @((Split-Path $source -Leaf)) -LogOption ('/LOG+:' + $cacheLog)
            $unblockFolders.Add((Split-Path $destination -Parent))
        }
        if ($exitCode -ge 8) { throw ("Include copy failed: {0}; destination: {1}; robocopy exit: {2}; cache log: {3}" -f $source, $destination, $exitCode, $cacheLog) }
        if (-not (Test-Path -LiteralPath $destination)) { throw ("Include absent after copy: {0}; source: {1}; cache log: {2}" -f $destination, $source, $cacheLog) }
        $integrity.Add([pscustomobject]@{ Source = $source; Destination = $destination })
    }
    $unblockFolders.Add((Join-Path $localRoot 'Config'))
    $unblockFolders.Add((Join-Path $localRoot 'Modules\SmartM365.Core'))
    Get-ChildItem -LiteralPath @($unblockFolders | Select-Object -Unique) -Include '*.ps1', '*.psm1', '*.psd1' -File -Recurse -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
    # Last check before the cached code is loaded (output check, then the launch).
    foreach ($pair in $integrity) { Assert-SmartM365InventoryCacheIntegrity -Source $pair.Source -Destination $pair.Destination }

    [pscustomobject]@{
        CacheBase = $cacheBase
        LocalSmartM365Root = $localRoot
        LocalScriptPath = (Join-Path $localScriptFolder (Split-Path $Plan.ScriptPath -Leaf))
        CacheLog = $cacheLog
    }
}

function Get-SmartM365InventoryOutputRoots {
    <# Data and log roots that a script started from $SmartM365Root resolves for the tenant. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SmartM365Root, [Parameter(Mandatory)][string]$Tenant)
    . (Join-Path $SmartM365Root 'Config\SmartM365-TenantContext.ps1')
    $config = Get-SmartM365EffectiveGlobalConfig -StartPath $SmartM365Root -ProfileKey $Tenant
    $values = @{}
    foreach ($property in $config.PSObject.Properties) { $values[$property.Name] = $property.Value }
    $result = [ordered]@{}
    foreach ($key in @('WorkspaceRootPath', 'DataAllRootPath', 'LatestCsvFolderPath', 'LogAllRootPath')) {
        if (-not $values.ContainsKey($key) -or $values[$key] -isnot [string] -or [string]::IsNullOrWhiteSpace($values[$key])) { continue }
        $resolved = [string]$values[$key]
        for ($i = 0; $i -lt 10 -and $resolved -match '\{\{'; $i++) {
            foreach ($match in [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')) {
                $tokenName = $match.Groups['Name'].Value
                if ($values.ContainsKey($tokenName) -and $null -ne $values[$tokenName]) { $resolved = $resolved.Replace($match.Value, [string]$values[$tokenName]) }
            }
        }
        if ([IO.Path]::IsPathRooted($resolved)) { $result[$key] = [IO.Path]::GetFullPath($resolved) }
        elseif ($resolved -notmatch '\{\{') { $result[$key] = [IO.Path]::GetFullPath((Join-Path $SmartM365Root $resolved)) }
    }
    return [pscustomobject]$result
}

function Assert-SmartM365InventoryOutputOutsideCache {
    <# Refuses a LocalCache run whose data or log root, or a path argument, resolves inside the cache. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)]$Cache, [AllowEmptyCollection()][string[]]$Arguments = @())
    $violations = New-Object System.Collections.Generic.List[string]
    # The configuration helpers report the template keys they add to the cached copy; not useful here.
    $roots = Get-SmartM365InventoryOutputRoots -SmartM365Root $Cache.LocalSmartM365Root -Tenant $Plan.Tenant 6>$null
    foreach ($property in $roots.PSObject.Properties) {
        if (Test-SmartM365InventoryPathInside -Path $property.Value -Root $Cache.CacheBase) { $violations.Add(('{0} = {1}' -f $property.Name, $property.Value)) }
    }
    foreach ($argument in $Arguments) {
        if ([string]::IsNullOrWhiteSpace($argument) -or $argument.StartsWith('-')) { continue }
        try { $rooted = [IO.Path]::IsPathRooted($argument) } catch { $rooted = $false }
        if ($rooted -and (Test-SmartM365InventoryPathInside -Path $argument -Root $Cache.CacheBase)) { $violations.Add(('argument {0}' -f $argument)) }
    }
    if ($violations.Count -gt 0) {
        throw ("Output resolves inside the disposable launcher cache {0}: {1}. Configure a persistent path outside LauncherCache." -f $Cache.CacheBase, ($violations -join '; '))
    }
    return $roots
}

function Get-SmartM365InventoryCommand {
    <# Final engine and argument list of a plan, for a script root (source or cache). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][string]$EnginePath, [Parameter(Mandatory)][string]$ScriptPath, [Parameter(Mandatory)][string]$SmartInventoryRoot)
    $tokenValues = @{
        Tenant = $Plan.Tenant
        ScriptDirectory = (Split-Path -Path $ScriptPath -Parent)
        SmartInventoryRoot = $SmartInventoryRoot
    }
    foreach ($key in @($Plan.Tokens.Keys)) {
        $tokenValues[$key] = Resolve-SmartM365InventoryTokenValue -Text $Plan.Tokens[$key].Value -Values $tokenValues
    }
    foreach ($key in @($Plan.Tokens.Keys)) {
        if ($Plan.Tokens[$key].MustExist -and -not (Test-Path -LiteralPath $tokenValues[$key])) { throw ("Token {0}: path not found: {1}" -f $key, $tokenValues[$key]) }
    }
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath)
    if ($Plan.PassTenant) { $arguments += @('-Tenant', $Plan.Tenant) }
    $arguments += @($Plan.Arguments | ForEach-Object { Resolve-SmartM365InventoryTokenValue -Text $_ -Values $tokenValues })
    $arguments += @($Plan.ExtraArguments)
    $workingDirectory = ''
    if ($Plan.WorkingDirectory) { $workingDirectory = [IO.Path]::GetFullPath((Join-Path $SmartInventoryRoot $Plan.WorkingDirectory)) }
    [pscustomobject]@{
        EnginePath = $EnginePath
        Arguments = $arguments
        WorkingDirectory = $workingDirectory
        CommandLine = ((ConvertTo-SmartM365InventoryCommandLineArgument -Value $EnginePath) + ' ' + (Join-SmartM365InventoryCommandLine -Arguments $arguments))
    }
}

function Format-SmartM365InventoryPlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)]$Catalog)
    $enginePath = '<{0} not found>' -f $Plan.Runtime
    try { $enginePath = Get-SmartM365InventoryEnginePath -Runtime $Plan.Runtime } catch { Write-Verbose $_.Exception.Message }
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(('Entry: {0} ({1}, action {2})' -f $Plan.Name, $Plan.Kind, $Plan.Action))
    if ($Plan.Description) { $lines.Add(('Description: {0}' -f $Plan.Description)) }
    $lines.Add(('Tenant: {0} (source: {1})' -f $Plan.Tenant, $Plan.TenantSource))
    $lines.Add(('Runtime: {0}' -f $Plan.Runtime))
    $scriptPath = $Plan.ScriptPath
    $inventoryRoot = $Catalog.SmartInventoryRoot
    if ($Plan.Mode -eq 'LocalCache') {
        $candidates = Get-SmartM365InventoryCacheCandidates -CacheName $Plan.CacheName
        $lines.Add(('Execution mode: LocalCache (requested {0}, source {1}{2})' -f $Plan.ModeInfo.Requested, $Plan.ModeInfo.Source, $(if ($Plan.ModeInfo.Zone) { ', zone ' + $Plan.ModeInfo.Zone } else { '' })))
        $lines.Add(('Cache folder: first writable of {0}' -f ($candidates -join ' | ')))
        if (@($Plan.Include).Count -gt 0) { $lines.Add(('Include: {0}' -f (@($Plan.Include) -join ', '))) }
        $relativeFolder = (Split-Path $Plan.ScriptPath -Parent).Substring($Catalog.SmartM365Root.TrimEnd('\').Length).TrimStart('\')
        $scriptPath = Join-Path (Join-Path (Join-Path $candidates[0] 'SmartM365') $relativeFolder) (Split-Path $Plan.ScriptPath -Leaf)
        $inventoryRoot = Join-Path (Join-Path $candidates[0] 'SmartM365') 'SmartInventory'
        $lines.Add(('Source script: {0}' -f $Plan.ScriptPath))
    }
    elseif ($null -ne $Plan.ModeInfo) {
        $lines.Add(('Execution mode: Direct (requested {0}, source {1}{2})' -f $Plan.ModeInfo.Requested, $Plan.ModeInfo.Source, $(if ($Plan.ModeInfo.Zone) { ', zone ' + $Plan.ModeInfo.Zone } else { '' })))
    }
    if ($Plan.Interactive) { $lines.Add('Interactive: yes (needs a user console)') }
    if ($Plan.RequiresConfirmation) { $lines.Add('Confirmation: the user must type WRITE') }
    try {
        $command = Get-SmartM365InventoryCommand -Plan $Plan -EnginePath $enginePath -ScriptPath $scriptPath -SmartInventoryRoot $inventoryRoot
        if ($command.WorkingDirectory) { $lines.Add(('Working directory: {0}' -f $command.WorkingDirectory)) }
        $lines.Add(('Command: {0}' -f $command.CommandLine))
    }
    catch { $lines.Add(('Command: unresolved ({0})' -f $_.Exception.Message)) }
    return ($lines -join [Environment]::NewLine)
}

function Invoke-SmartM365InventoryLaunch {
    <#
    Runs a resolved plan and returns the exit code of the launched script. Preparation errors
    throw; the caller decides how to report them.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Plan,
        [Parameter(Mandatory)]$Catalog,
        [string[]]$CacheBaseCandidates = @()
    )
    $command = Initialize-SmartM365InventoryLaunch -Plan $Plan -Catalog $Catalog -CacheBaseCandidates $CacheBaseCandidates
    if ($null -eq $command) { return 1 }
    return (Invoke-SmartM365InventoryProcess -FilePath $command.EnginePath -Arguments $command.Arguments -WorkingDirectory $command.WorkingDirectory)
}

function Invoke-SmartM365InventoryOrchestratedLaunch {
    <#
    Used by the orchestrator supervisor: resolves the job from the catalog, prepares it (local
    cache, output checks) and runs it with its output appended to the job log. Returns the exit
    code of the script, or 1 when the launch could not be prepared.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SmartInventoryRoot,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Tenant,
        [Parameter(Mandatory)][string]$LogPath,
        [AllowEmptyString()][string]$ExecutionMode = '',
        [AllowNull()]$JobArguments = $null,
        [switch]$NoAutoConnect,
        # Script validated by the orchestrator (Authenticode); the launch is refused if the catalog resolves another one.
        [AllowEmptyString()][string]$ExpectedScriptPath = ''
    )
    $script:LaunchLogPath = $LogPath
    $ErrorActionPreference = 'Continue'
    try {
        try {
            $catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot $SmartInventoryRoot
            $plan = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name $Name -Tenant $Tenant -ExecutionMode $ExecutionMode -ExecutionModeSource 'Job (runtime manifest)' -JobArguments $JobArguments -NoAutoConnect:$NoAutoConnect
            if ($ExpectedScriptPath -and [IO.Path]::GetFullPath($plan.ScriptPath) -ine [IO.Path]::GetFullPath($ExpectedScriptPath)) {
                throw ("the catalog script {0} differs from the validated script {1}" -f $plan.ScriptPath, $ExpectedScriptPath)
            }
            $command = Initialize-SmartM365InventoryLaunch -Plan $plan -Catalog $catalog -Orchestrated
        }
        catch {
            Write-SmartM365InventoryLaunchMessage -Level ERROR -Message ("Catalog engine could not prepare {0}: {1}" -f $Name, $_.Exception.Message)
            return 1
        }
        if ($null -eq $command) { return 1 }
        $encoding = New-Object System.Text.UTF8Encoding($false)
        $pushed = $false
        if ($command.WorkingDirectory) { Push-Location -LiteralPath $command.WorkingDirectory; $pushed = $true }
        try {
            # Same capture as the ScriptPath launch: every stream of the script goes to the job log.
            & $command.EnginePath @($command.Arguments) 2>&1 | ForEach-Object {
                [IO.File]::AppendAllText($LogPath, ([string]$_) + [Environment]::NewLine, $encoding)
            }
            $exitCode = $LASTEXITCODE
        }
        finally { if ($pushed) { Pop-Location } }
        if ($null -eq $exitCode) { $exitCode = 1 }
        return [int]$exitCode
    }
    finally { $script:LaunchLogPath = '' }
}

function Initialize-SmartM365InventoryLaunch {
    <#
    Checks and prepares a plan (console, confirmation, local cache, output checks) and returns
    the final command, or $null when the user cancelled the confirmation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Plan,
        [Parameter(Mandatory)]$Catalog,
        [string[]]$CacheBaseCandidates = @(),
        [switch]$Orchestrated
    )
    $inputRedirected = $false
    try { $inputRedirected = [Console]::IsInputRedirected } catch { $inputRedirected = $true }
    $interactiveConsole = [Environment]::UserInteractive -and -not $inputRedirected
    if (($Plan.Interactive -or $Plan.RequiresConfirmation) -and -not $interactiveConsole) {
        throw "'$($Plan.Name)' needs an interactive console and cannot run unattended."
    }
    $enginePath = Get-SmartM365InventoryEnginePath -Runtime $Plan.Runtime
    if (-not (Test-Path -LiteralPath $Plan.ScriptPath -PathType Leaf)) { throw "Script not found: $($Plan.ScriptPath)" }

    if ($Plan.RequiresConfirmation) {
        Write-SmartM365InventoryLaunchMessage -Level WARN -Message ("WARNING: '{0}' can modify data.{1}It runs {2} {3}." -f $Plan.Name, [Environment]::NewLine, (Split-Path $Plan.ScriptPath -Leaf), (@($Plan.Arguments) -join ' '))
        if ($Plan.Description) { Write-SmartM365InventoryLaunchMessage -Level WARN -Message $Plan.Description }
        $answer = Read-Host -Prompt 'Type WRITE to continue'
        if ([string]$answer -ine 'WRITE') {
            Write-SmartM365InventoryLaunchMessage -Message 'Write mode cancelled.'
            return $null
        }
    }

    $scriptPath = $Plan.ScriptPath
    $inventoryRoot = $Catalog.SmartInventoryRoot
    $outputCheckArguments = @($Plan.Arguments) + @($Plan.ExtraArguments)
    if ($Plan.Mode -eq 'LocalCache') {
        $cache = Initialize-SmartM365InventoryLocalCache -Plan $Plan -Catalog $Catalog -CacheBaseCandidates $CacheBaseCandidates -Orchestrated:$Orchestrated
        Write-SmartM365InventoryLaunchMessage -Message ("Execution mode: LocalCache; cache folder: {0}" -f $cache.LocalSmartM365Root)
        Assert-SmartM365InventoryOutputOutsideCache -Plan $Plan -Cache $cache -Arguments $outputCheckArguments | Out-Null
        $scriptPath = $cache.LocalScriptPath
        $inventoryRoot = Join-Path $cache.LocalSmartM365Root 'SmartInventory'
    }
    else {
        Write-SmartM365InventoryLaunchMessage -Message ("Execution mode: Direct; script: {0}" -f $Plan.ScriptPath)
    }
    $command = Get-SmartM365InventoryCommand -Plan $Plan -EnginePath $enginePath -ScriptPath $scriptPath -SmartInventoryRoot $inventoryRoot
    if ($command.WorkingDirectory -and -not (Test-Path -LiteralPath $command.WorkingDirectory -PathType Container)) { throw "Working directory not found: $($command.WorkingDirectory)" }
    Write-SmartM365InventoryLaunchMessage -Message ("Starting {0} (tenant {1}): {2}" -f $Plan.Name, $Plan.Tenant, $command.CommandLine)
    return $command
}

function Get-SmartM365InventoryMenu {
    <#
    Numbered menu of the catalog for SmartInventory.cmd: the jobs sorted by group, the manual
    launches, then the orchestrator entries (Orchestrator-*).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Catalog)
    $sections = @(
        [pscustomobject]@{ Title = 'Jobs: run now on this computer, or request through the orchestrators'; Entries = @($Catalog.Entries | Where-Object { $_.Kind -eq 'Job' } | Sort-Object -Property @{ Expression = { [string](Get-SmartM365InventoryProperty -Object $_.Definition -Name 'Group' -Default '') } }, Name) },
        [pscustomobject]@{ Title = 'Manual launches: run on this computer'; Entries = @($Catalog.Entries | Where-Object { $_.Kind -eq 'Manual' -and $_.Name -notlike 'Orchestrator-*' } | Sort-Object -Property Name) },
        [pscustomobject]@{ Title = 'Orchestrator'; Entries = @($Catalog.Entries | Where-Object { $_.Kind -eq 'Manual' -and $_.Name -like 'Orchestrator-*' }) }
    )
    $number = 0
    foreach ($section in $sections) {
        foreach ($entry in $section.Entries) {
            $number++
            $definition = $entry.Definition
            $runtime = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'Runtime' -Default (Get-SmartM365InventoryProperty -Object $definition -Name 'PowerShellEdition' -Default 'PowerShell7'))
            $description = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'Description' -Default '')
            if (-not $description) { $description = 'Group ' + [string](Get-SmartM365InventoryProperty -Object $definition -Name 'Group' -Default '-') }
            [pscustomobject]@{
                Number = $number
                Section = $section.Title
                Name = $entry.Name
                Kind = $entry.Kind
                Runtime = $(if ($runtime -eq 'WindowsPowerShell') { 'PS 5.1' } else { 'PS 7' })
                Cache = $null -ne (Get-SmartM365InventoryProperty -Object $definition -Name 'Execution')
                Interactive = [bool](Get-SmartM365InventoryProperty -Object $definition -Name 'Interactive' -Default $false)
                RequiresConfirmation = [bool](Get-SmartM365InventoryProperty -Object $definition -Name 'RequiresConfirmation' -Default $false)
                Description = $description
            }
        }
    }
}

function Resolve-SmartM365InventoryMenuChoice {
    <# Menu item for a typed number or entry name (case-insensitive), or $null. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Menu, [AllowEmptyString()][string]$Choice)
    $value = ([string]$Choice).Trim()
    if (-not $value) { return $null }
    $number = 0
    if ([int]::TryParse($value, [ref]$number)) { return (@($Menu | Where-Object { $_.Number -eq $number }) | Select-Object -First 1) }
    return (@($Menu | Where-Object { $_.Name -ieq $value }) | Select-Object -First 1)
}

function Test-SmartM365InventoryCatalog {
    <#
    Validates the catalog without executing anything: names, scripts, runtimes, execution,
    includes, tokens, working directories, legacy launchers, and every named argument against
    the param block of its script (AST).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Catalog)
    $errors = New-Object System.Collections.Generic.List[string]
    $legacy = @{}
    foreach ($entry in $Catalog.Entries) {
        $definition = $entry.Definition
        $name = $entry.Name
        if ($name -notmatch '^[A-Za-z0-9._-]+$') { $errors.Add("$name : invalid name.") }
        $scriptRelativePath = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'ScriptPath' -Default '')
        $scriptPath = if ($scriptRelativePath) { Join-Path $Catalog.SmartInventoryRoot $scriptRelativePath } else { '' }
        $parameters = @()
        if (-not $scriptPath -or -not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) { $errors.Add("$name : script not found: $scriptRelativePath") }
        else {
            try { $parameters = @(Get-SmartM365InventoryScriptParameters -Path $scriptPath) } catch { $errors.Add("$name : $($_.Exception.Message)") }
        }
        $runtime = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'Runtime' -Default '')
        if ($runtime -and $runtime -notin $script:ValidRuntimes) { $errors.Add("$name : invalid Runtime '$runtime'.") }
        $execution = Get-SmartM365InventoryProperty -Object $definition -Name 'Execution'
        if ($null -ne $execution) {
            $mode = [string](Get-SmartM365InventoryProperty -Object $execution -Name 'Mode' -Default '')
            if ($mode -and $mode -notin $script:ValidModes) { $errors.Add("$name : invalid Execution.Mode '$mode'.") }
            $cacheName = [string](Get-SmartM365InventoryProperty -Object $execution -Name 'CacheName' -Default '')
            if ($cacheName -notmatch '^[A-Za-z0-9._-]+(\\[A-Za-z0-9._-]+)*$') { $errors.Add("$name : Execution.CacheName is required (letters, digits, . _ - and \).") }
            foreach ($item in @(Get-SmartM365InventoryProperty -Object $execution -Name 'Include' -Default @())) {
                if (-not (Test-Path -LiteralPath (Join-Path $Catalog.SmartM365Root ([string]$item)))) { $errors.Add("$name : Include not found: $item") }
            }
        }
        $hideInGui = Get-SmartM365InventoryProperty -Object $definition -Name 'HideInGui'
        if ($null -ne $hideInGui -and ($hideInGui -isnot [bool] -or $entry.Kind -ne 'Manual')) { $errors.Add("$name : HideInGui must be true or false, on manual entries only.") }
        $workingDirectory = [string](Get-SmartM365InventoryProperty -Object $definition -Name 'WorkingDirectory' -Default '')
        if ($workingDirectory -and -not (Test-Path -LiteralPath (Join-Path $Catalog.SmartInventoryRoot $workingDirectory) -PathType Container)) { $errors.Add("$name : WorkingDirectory not found: $workingDirectory") }
        $tokenNames = @('Tenant', 'ScriptDirectory', 'SmartInventoryRoot')
        $tokens = Get-SmartM365InventoryProperty -Object $definition -Name 'Tokens'
        if ($null -ne $tokens) { $tokenNames += @($tokens.PSObject.Properties | ForEach-Object { $_.Name }) }
        $argumentText = ([string](Get-SmartM365InventoryProperty -Object $definition -Name 'LaunchArguments' -Default '')) + ' ' + ([string](Get-SmartM365InventoryProperty -Object $definition -Name 'Arguments' -Default ''))
        if ($entry.Kind -eq 'Manual' -and [string](Get-SmartM365InventoryProperty -Object $definition -Name 'LaunchArguments' -Default '')) { $errors.Add("$name : LaunchArguments is reserved for jobs; use Arguments.") }
        foreach ($match in [regex]::Matches($argumentText, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')) {
            if ($match.Groups['Name'].Value -notin $tokenNames) { $errors.Add("$name : token $($match.Value) is not defined.") }
        }
        if ($parameters.Count -gt 0) {
            foreach ($argument in @(ConvertTo-SmartM365InventoryArgumentList -Arguments $argumentText)) {
                if ($argument -notmatch '^-(?<Name>[A-Za-z][A-Za-z0-9_]*)(:.*)?$') { continue }
                $parameterName = $Matches['Name']
                if (-not (@($parameters) | Where-Object { $_ -ieq $parameterName })) { $errors.Add("$name : parameter -$parameterName is not declared by $scriptRelativePath.") }
            }
        }
        foreach ($launcher in @(Get-SmartM365InventoryProperty -Object $definition -Name 'LegacyLaunchers' -Default @())) {
            $key = ([string]$launcher).ToLowerInvariant()
            if ($legacy.ContainsKey($key)) { $errors.Add("$name : legacy launcher $launcher is also mapped to $($legacy[$key]).") } else { $legacy[$key] = $name }
        }
    }
    return $errors.ToArray()
}

Export-ModuleMember -Function Get-SmartM365InventoryCatalog, Find-SmartM365InventoryCatalogEntry, ConvertTo-SmartM365InventoryArgumentList,
    ConvertTo-SmartM365InventoryCommandLineArgument, Get-SmartM365InventoryScriptParameters, Resolve-SmartM365InventoryTenant,
    Get-SmartM365InventoryGlobalSetting, Get-SmartM365InventoryPathZone, Resolve-SmartM365InventoryExecutionMode, Resolve-SmartM365InventoryLaunch,
    Get-SmartM365InventoryEnginePath, Get-SmartM365InventoryCacheCandidates, Initialize-SmartM365InventoryLocalCache, Get-SmartM365InventoryOutputRoots,
    Assert-SmartM365InventoryOutputOutsideCache, Get-SmartM365InventoryCommand, Format-SmartM365InventoryPlan, Invoke-SmartM365InventoryLaunch,
    Initialize-SmartM365InventoryLaunch, Invoke-SmartM365InventoryOrchestratedLaunch, Test-SmartM365InventoryCacheLocation,
    Assert-SmartM365InventoryCacheIntegrity, Get-SmartM365InventoryMenu, Resolve-SmartM365InventoryMenuChoice,
    Test-SmartM365InventoryCatalog, Write-SmartM365InventoryLaunchMessage

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB/iwYsd58KpU/e
# 18iUT9t/11Lv61suZ9TnD62oFBO3+aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP3zT3HqnMs+G4qUWAXkk1oX678M73wGGg4/55SlKsZaMA0GCSqG
# SIb3DQEBAQUABIIBgJeBD41LARWYFMwNOZeDMxGFBJ1QRzcJXrcJm169xw9/RcGW
# 2dhKEZbwVZjn56R2g650EdAE9TaR+sEj0VuuHS2mjCNtFMmGTPKJ9e1hOqT8u+/i
# A+N/qLDDe+DtmKHNCUiQAK347dZmOHLqkj3y53c+5Cl+uSy87/WphrmiLy9JYXO3
# smzCKSDX6X6733GLX9AHQl2UihDNDQDqtWYsPFlyFt5Vc+UqquKRw/Up9yNIpnGh
# q1EZZt0b0H6UDjpMQdOHEVbaYf2x4IFOO18CSZE/V+/e6HFrBlSUGPvRnk0goao8
# Ni3TOP/LcisDEOlyG2BmAzOrcIZ+83uVnjhJYZUVHHW/Ez5oKDLDByLFiqYGV4XV
# 9/j7HSKK3EIlkWCUMl+vNI2yrpJNoVR3l7+Kz0TZBk2Qxh/OkT8rJyVkg8vNiu+t
# KgiGkYf0zmgpQTjfuKIMfu329I4+LZcybV6CunTj2wKIQSiUStYg2/vI3jM7r6HY
# KNG9nddwWM2fbzJhdqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYxMDU0
# NDZaMC8GCSqGSIb3DQEJBDEiBCAaQIbUo+rZ8PJeByOCWdVz4DeOqJRRrnR5MKRO
# 96EDUzANBgkqhkiG9w0BAQEFAASCAgAkgFEdyOXM5bV7kxDSf3JvVIZpZVvS5/b2
# NFMbXnJLAriwH/B3nb3JqifH39UA9/ERZSvPod3TQLL5J6Ikhc0M1GTUIHCS+1pW
# FhFy5mQMIPCQpzhPFO9LYWh+UDMG0Iqb7pDDW5l1G1lAVti/jA7EGI5rn8ZeHZT2
# XXdcO0YCOFfKiOH/pBv/DSyIwufU02sMlDg+GPT+aHJILKOc0utwbOBEAi+A7NCr
# InSrgDe6SWoCcLyCvmVYmyOTJ60dOdmC8gOjU2Us/Un3r3rFluSrimePX7GEhb/c
# rUPCDHd1c2oYbB6fp+dGsV6SrTRdQcAPwBYVht3gqS9axhO4Sk2W6dLW7b8x2pG3
# SMzRBPXYdtuJGXAiUFmwoPqdRbsNAW8TPWrGuSjGUOevStQ4WsEsd6o43DSDBhp7
# OGRi/epVwx0zHI/XdYy/NompwpG+QvwJU10SyTRp/OGn6u4bWVDcTUB5w9IMtkVl
# kiagVkS/1TpHZy8KIYvo2UF1Lb6ZSZDTCD5iw+5V5mZFsz3vcyNkGsKUz14pygGi
# QY4PH/bBPI25RuFUBzGoqgXSLSApOaywVsTie3UNbn0i4UTX3jDRKD62Vm3v9Xod
# ynH2R4o/J0AK/f2s30W1mZ+CN+HmRjRlnPu/TKhFi9osKlHm/dkVOnBfKR+5eikn
# 8fMkcSF0SA==
# SIG # End signature block
