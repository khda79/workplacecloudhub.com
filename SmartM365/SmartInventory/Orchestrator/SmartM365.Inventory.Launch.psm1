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

LocalCache reproduces the legacy on-premises launchers: first writable folder among
ProgramData\SmartM365\LauncherCache\<CacheName>, Windows\Temp, C:\Temp and %TEMP%; cache
emptied; robocopy /MIR of the script folder, Config and Modules\SmartM365.Core; failure from
exit code 8; PrepareCache.log; Include files checked before and after the copy; Unblock-File.
Outputs that resolve inside the cache are refused.

Compatible with Windows PowerShell 5.1 and PowerShell 7.

.VERSION
1.2.0
#>
Set-StrictMode -Version 2.0

$script:LaunchModuleVersion = '1.2.0'
$script:ValidRuntimes = @('PowerShell7', 'WindowsPowerShell')
$script:ValidModes = @('Direct', 'LocalCache', 'Auto')
# Test hook: replaced by the offline tests to simulate the zone of a share.
$script:PathZoneResolver = $null
$script:LaunchLogPath = ''

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
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CacheName)
    $candidates = @((Join-Path $env:ProgramData ('SmartM365\LauncherCache\' + $CacheName)), (Join-Path $env:SystemRoot ('Temp\SmartM365LauncherCache\' + $CacheName)), ('C:\Temp\SmartM365LauncherCache\' + $CacheName))
    if ($env:TEMP) { $candidates += (Join-Path $env:TEMP ('SmartM365LauncherCache\' + $CacheName)) }
    return $candidates
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
        [string[]]$CacheBaseCandidates = @()
    )
    if (-not $Plan.CacheName) { throw "Entry '$($Plan.Name)' declares no Execution.CacheName." }
    if ($CacheBaseCandidates.Count -eq 0) { $CacheBaseCandidates = Get-SmartM365InventoryCacheCandidates -CacheName $Plan.CacheName }
    $cacheBase = ''
    foreach ($candidate in $CacheBaseCandidates) {
        try {
            New-Item -ItemType Directory -Path $candidate -Force -ErrorAction Stop | Out-Null
            $probe = Join-Path $candidate '.smartm365-cache-write-test'
            [IO.File]::WriteAllText($probe, '')
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
            $cacheBase = $candidate
            break
        }
        catch { continue }
    }
    if (-not $cacheBase) { throw ("Failed to create a writable local cache base folder. Tried:{0}  {1}" -f [Environment]::NewLine, ($CacheBaseCandidates -join ([Environment]::NewLine + '  '))) }

    $smartM365Root = $Catalog.SmartM365Root
    $scriptFolder = Split-Path -Path $Plan.ScriptPath -Parent
    if (-not (Test-SmartM365InventoryPathInside -Path $scriptFolder -Root $smartM365Root)) { throw "Script folder is outside SmartM365: $scriptFolder" }
    $scriptFolderRelative = $scriptFolder.Substring($smartM365Root.TrimEnd('\').Length).TrimStart('\')
    $localRoot = Join-Path $cacheBase 'SmartM365'
    $localScriptFolder = Join-Path $localRoot $scriptFolderRelative
    $cacheLog = Join-Path $cacheBase 'PrepareCache.log'

    # The cache is rebuilt for every run (same as rmdir /s /q in the legacy launchers).
    if (Test-Path -LiteralPath $localRoot) { & (Join-Path $env:SystemRoot 'System32\cmd.exe') /d /c rmdir /s /q "$localRoot" 2>$null | Out-Null }
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
    }
    $unblockFolders.Add((Join-Path $localRoot 'Config'))
    $unblockFolders.Add((Join-Path $localRoot 'Modules\SmartM365.Core'))
    Get-ChildItem -LiteralPath @($unblockFolders | Select-Object -Unique) -Include '*.ps1', '*.psm1', '*.psd1' -File -Recurse -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue

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
        [switch]$NoAutoConnect
    )
    $script:LaunchLogPath = $LogPath
    $ErrorActionPreference = 'Continue'
    try {
        try {
            $catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot $SmartInventoryRoot
            $plan = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name $Name -Tenant $Tenant -ExecutionMode $ExecutionMode -ExecutionModeSource 'Job (runtime manifest)' -JobArguments $JobArguments -NoAutoConnect:$NoAutoConnect
            $command = Initialize-SmartM365InventoryLaunch -Plan $plan -Catalog $catalog
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
        [string[]]$CacheBaseCandidates = @()
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
        $cache = Initialize-SmartM365InventoryLocalCache -Plan $Plan -Catalog $Catalog -CacheBaseCandidates $CacheBaseCandidates
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
    Initialize-SmartM365InventoryLaunch, Invoke-SmartM365InventoryOrchestratedLaunch,
    Test-SmartM365InventoryCatalog, Write-SmartM365InventoryLaunchMessage
