#Requires -Version 5.1
<# Replays only the actual prerequisite against synthetic modules. No collector, installation or connection. #>
[CmdletBinding()]
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'
$checks = 0
function Check([bool]$Ok, [string]$Message) { if (-not $Ok) { throw $Message }; $script:checks++ }
function WriteLog { param($Message, $Level); $script:messages += [string]$Message }
function Install-Module { throw 'Installation forbidden in offline test' }
function Connect-MgGraph { throw 'Connection forbidden in offline test' }
$path = Join-Path $PSScriptRoot '../SmartInventory/ExchangeInventory/OnPremises/ProxyAddresses/SmartM365-Check-ProxyAddresses-Exchange.ps1'
if ($SourcePath) { $path = $SourcePath }
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
Check ($errors.Count -eq 0) 'Collector parser failed'
$definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Assert-ProxyAddressesGraphPrerequisite' }, $true)
Check ($null -ne $definition) 'Prerequisite function missing'
. ([scriptblock]::Create($definition.Extent.Text))
$call = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Assert-ProxyAddressesGraphPrerequisite' }, $true)
$snapin = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-PSSnapin' }, $true)
Check ($call.Extent.StartOffset -lt $snapin.Extent.StartOffset) 'Graph check must precede Exchange loading/collection'
Check ($call.Extent.Text -match '-EnableSharePointUpload \$global:EnableSharePointUpload -SendMailMode \$SendMailMode -SmtpServer \$SmtpServer') 'Call must use effective configuration'
Check ($call.Extent.StartOffset -gt $ast.Extent.Text.IndexOf('$ScriptLocalConfig = Get-ScriptLocalConfig')) 'Graph check must follow config resolution'
$savedPath = $env:PSModulePath
$root = Join-Path ([IO.Path]::GetTempPath()) ('proxy-graph-' + [guid]::NewGuid().ToString('N'))
$moduleBase = Join-Path $root 'Microsoft.Graph.Authentication'
$null = New-Item -ItemType Directory -Path $moduleBase -Force
$manifest = Join-Path $moduleBase 'Microsoft.Graph.Authentication.psd1'
$body = Join-Path $moduleBase 'Microsoft.Graph.Authentication.psm1'
try {
    Check (@(Get-Module Microsoft.Graph.Authentication).Count -eq 0) 'Run in a fresh process without a loaded Graph module'
    $env:PSModulePath = $root
    $cases = @(
        @{Name='SharePoint'; Upload=$true; Mode='SMTP'; Smtp='relay'; Needs=$true},
        @{Name='Graph mail'; Upload=$false; Mode='Graph'; Smtp='relay'; Needs=$true},
        @{Name='Both'; Upload=$false; Mode='Both'; Smtp='relay'; Needs=$true},
        @{Name='Implicit Graph'; Upload=$false; Mode=''; Smtp=''; Needs=$true},
        @{Name='Explicit SMTP'; Upload=$false; Mode='SMTP'; Smtp='relay'; Needs=$false},
        @{Name='Implicit SMTP'; Upload=$false; Mode=$null; Smtp='relay'; Needs=$false}
    )
    foreach ($case in $cases) {
        $messages = @(); $failure = $null
        try { Assert-ProxyAddressesGraphPrerequisite $case.Upload $case.Mode $case.Smtp } catch { $failure = $_ }
        Check (($null -ne $failure) -eq $case.Needs) ($case.Name + ': missing module gate')
        if ($case.Needs) {
            Check ($failure.Exception.Message -match 'Graph Authentication prerequisite failed.*PSModulePath=') ($case.Name + ': actionable diagnostic')
            Check ($failure.Exception.Message -match 'No module installation or Graph connection was attempted') ($case.Name + ': no automatic repair')
        }
    }
    "@{RootModule='Microsoft.Graph.Authentication.psm1';ModuleVersion='1.2.3';GUID='ead3110d-d8f9-4a6b-976c-2d9ea09cf0fb';FunctionsToExport=@('Get-MgContext');CmdletsToExport=@();AliasesToExport=@();VariablesToExport=@()}" | Set-Content -LiteralPath $manifest -Encoding UTF8
    "function Get-MgContext { throw 'Graph context must not be queried by module prerequisite' }" | Set-Content -LiteralPath $body -Encoding UTF8
    foreach ($case in $cases | Where-Object Needs) {
        $messages = @()
        Assert-ProxyAddressesGraphPrerequisite $case.Upload $case.Mode $case.Smtp
        $loaded = Get-Module Microsoft.Graph.Authentication
        Check ($loaded.ModuleBase -eq $moduleBase) ($case.Name + ': only fixture loaded')
        Check (($messages -join '|') -match 'prerequisite OK; PowerShell=.*Version=1.2.3; Path=') ($case.Name + ': runtime/version/path logged')
        Check ($null -ne (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) ($case.Name + ': module remains available to later publication')
        Remove-Module Microsoft.Graph.Authentication -Force
    }
    "throw 'Synthetic incompatible Graph module'" | Set-Content -LiteralPath $body -Encoding UTF8
    $failure = $null
    try { Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay' } catch { $failure = $_ }
    Check ($null -ne $failure -and $failure.Exception.Message -match 'Synthetic incompatible Graph module') 'Import failure must stop prerequisite'
    Check ($env:PSModulePath -ceq $root) 'Prerequisite must not add PS7 paths'
    $messages = @()
    Assert-ProxyAddressesGraphPrerequisite $false 'SMTP' 'relay'
    Check (($messages -join '|') -match 'prerequisite skipped') 'SMTP-only skips even an incompatible module'
}
finally {
    Get-Module Microsoft.Graph.Authentication | Where-Object { $_.ModuleBase -eq $moduleBase } | Remove-Module -Force
    $env:PSModulePath = $savedPath
    if (-not ([IO.Path]::GetFullPath($root)).StartsWith(([IO.Path]::GetFullPath([IO.Path]::GetTempPath())).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe fixture cleanup path' }
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
Write-Output ("PASS: {0} ProxyAddresses Graph prerequisite checks; {1}; synthetic only." -f $checks, $PSVersionTable.PSVersion)
