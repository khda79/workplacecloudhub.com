#Requires -Version 7.0
<#
.SYNOPSIS
Launches synthetic Windows PowerShell 5.1 and PowerShell 7 jobs through the real orchestrator job launcher
with the PowerShell 7 module path first, and checks that the 5.1 job still loads its own modules.

.DESCRIPTION
No collector, tenant or network is used. The orchestrator's Start-InventoryJob, Get-JobEngine and
Get-OrchestratorWindowsPowerShellModulePathCommand run unchanged from pwsh, with $env:PSModulePath starting with
the PowerShell 7 module folders (as in the resident orchestrator). Checks:
- a WindowsPowerShell job loads Microsoft.PowerShell.Security from Windows PowerShell and Get-Acl works;
- without the prefix the same job fails (the production failure is reproduced);
- the PSModulePath keeps the other entries (Documents\WindowsPowerShell\Modules, Program Files\WindowsPowerShell\Modules);
- a PowerShell 7 job is unchanged: no prefix, it still loads the PowerShell 7 module;
- the orchestrator process environment is not changed.

.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Orchestrator parse failed.' }
$definitions = foreach ($name in @('Start-InventoryJob', 'Get-JobEngine', 'Get-PwshPath', 'Get-OrchestratorWindowsPowerShellModulePathCommand')) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $node) { throw "Missing orchestrator function: $name" }
    $node.Extent.Text
}

$script:Checks = 0
function Assert-ModulePath {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ModulePath-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $root
$collectorPath = Join-Path $root 'SyntheticCollector.ps1'
[IO.File]::WriteAllText($collectorPath, @'
param([string]$Tenant)
try {
    $null = Get-Acl -LiteralPath $PSScriptRoot
    'ACL OK edition=' + $PSVersionTable.PSEdition + ' security=' + (Get-Module Microsoft.PowerShell.Security).ModuleBase
    'PATH=' + $env:PSModulePath
    exit 0
}
catch {
    'ACL FAILED edition=' + $PSVersionTable.PSEdition + ' ' + $_.Exception.Message
    exit 7
}
'@, [Text.UTF8Encoding]::new($false))

$windowsPowerShellModules = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\Modules'
$pwshModules = Join-Path $env:ProgramFiles 'PowerShell\7\Modules'
$savedModulePath = $env:PSModulePath
$pollutedModulePath = (@($pwshModules, (Join-Path $env:USERPROFILE 'Documents\PowerShell\Modules'), (Join-Path $env:ProgramFiles 'PowerShell\Modules'), (Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules'), (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'), $windowsPowerShellModules) -join ';')

function Invoke-SyntheticJob {
    param([string]$Name, [string]$Edition, [switch]$WithoutPrefix)
    $folder = Join-Path $root $Name
    $null = New-Item -ItemType Directory -Path $folder
    $module = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
    & $module {
        param($folder, $jobName, $withoutPrefix)
        $script:Settings = @{ JobLogFolderPath = $folder; SharedDataFolderPath = $folder }
        $script:RunningJobs = @{}
        $script:State = @{ Jobs = @{ $jobName = @{ Running = $null } } }
        $script:JobState = $script:State.Jobs[$jobName]
        $script:ScriptName = 'SyntheticOrchestrator'
        $script:Tenant = 'test'
        $script:Connect = $false
        function script:Get-JobState { param($JobName) $script:JobState }
        function script:Resolve-OrchestratorJobPath { param($Path) $Path }
        function script:Invoke-OrchestratorAuthenticodeValidation { param($Job, $ScriptFullPath) [pscustomobject]@{ Allowed = $true } }
        function script:ConvertTo-StateTime { param($Value) ([datetime]$Value).ToString('o') }
        function script:Save-OrchestratorState {}
        function script:Write-OrchestratorLog { param($Message, $Level) }
        function script:Complete-JobRun { throw 'Synthetic launch failed before a child was started.' }
        if ($withoutPrefix) { function script:Get-OrchestratorWindowsPowerShellModulePathCommand { '' } }
    } $folder $Name $WithoutPrefix.IsPresent
    $job = [pscustomobject]@{ Name = $Name; ScriptPath = $collectorPath; Arguments = ''; PowerShellEdition = $Edition; TimeoutMinutes = 2 }
    & $module { param($job) Start-InventoryJob -Job $job -Occurrence (Get-Date) } $job
    $process = & $module { param($n) $script:RunningJobs[$n].Process } $Name
    if (-not $process.WaitForExit(60000)) { throw "Synthetic job $Name did not exit." }
    $record = & $module { param($n) $script:State.Jobs[$n].Running } $Name
    $result = Get-Content -LiteralPath $record.ResultPath -Raw | ConvertFrom-Json
    $log = Get-ChildItem -LiteralPath $folder -Filter '*.log' -File -Recurse | Select-Object -First 1
    return [pscustomobject]@{ ExitCode = $result.ExitCode; Log = [IO.File]::ReadAllText($log.FullName) }
}

try {
    $env:PSModulePath = $pollutedModulePath
    $windowsJob = Invoke-SyntheticJob -Name 'WindowsPowerShellJob' -Edition 'WindowsPowerShell'
    $controlJob = Invoke-SyntheticJob -Name 'WindowsPowerShellControl' -Edition 'WindowsPowerShell' -WithoutPrefix
    $coreJob = Invoke-SyntheticJob -Name 'PowerShell7Job' -Edition 'PowerShell7'
    Assert-ModulePath ($env:PSModulePath -ceq $pollutedModulePath) 'The orchestrator process keeps its own PSModulePath.'

    Assert-ModulePath ($windowsJob.ExitCode -eq 0 -and $windowsJob.Log -match ('ACL OK edition=Desktop security=' + [regex]::Escape((Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0')))) "A Windows PowerShell job loads its own security module and Get-Acl works: $($windowsJob.Log)"
    $windowsPath = ([regex]::Match($windowsJob.Log, 'PATH=(?<Path>[^\r\n]+)')).Groups['Path'].Value -split ';'
    Assert-ModulePath ($windowsPath[0] -ieq $windowsPowerShellModules) 'Windows PowerShell modules come first.'
    Assert-ModulePath (@($windowsPath | Where-Object { $_ -match '(?i)(^|\\)PowerShell\\(7[^\\]*\\)?Modules\\?$' }).Count -eq 0) "PowerShell 7 module folders are removed: $($windowsPath -join ';')"
    Assert-ModulePath (($windowsPath -contains (Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules')) -and ($windowsPath -contains (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'))) 'Windows PowerShell user and machine module folders are kept.'
    Assert-ModulePath ($windowsJob.Log -match 'Child command: \$smartM365ModulePath = ') 'The repair is visible at the start of the logged child command.'

    if (Test-Path -LiteralPath (Join-Path $pwshModules 'Microsoft.PowerShell.Security')) {
        Assert-ModulePath ($controlJob.ExitCode -eq 7 -and $controlJob.Log -match 'ACL FAILED edition=Desktop') "Without the repair the same job fails, as in production: $($controlJob.Log)"
    }

    Assert-ModulePath ($coreJob.ExitCode -eq 0 -and $coreJob.Log -match 'ACL OK edition=Core security=[^\r\n]*PowerShell\\7') "A PowerShell 7 job still loads its PowerShell 7 module: $($coreJob.Log)"
    Assert-ModulePath ($coreJob.Log -notmatch 'smartM365ModulePath') 'A PowerShell 7 job gets no module path prefix.'
}
finally {
    $env:PSModulePath = $savedModulePath
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
}

"PASS: $script:Checks orchestrator Windows PowerShell module path checks."
