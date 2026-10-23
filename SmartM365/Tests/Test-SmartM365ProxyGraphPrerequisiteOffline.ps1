#Requires -Version 5.1
<# Replays the prerequisite against synthetic modules and a mocked CurrentUser installer. No real installation or connection. #>
[CmdletBinding()]
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'
$checks = 0
function Check([bool]$Ok, [string]$Message) { if (-not $Ok) { throw $Message }; $script:checks++ }
function WriteLog { param($Message, $Level); $script:messages += [string]$Message }
function Install-Module {
    [CmdletBinding()]
    param($Name, $RequiredVersion, $Repository, $Scope, [switch]$Force, [switch]$AllowClobber)
    $script:installCalls++
    Check ($Name -eq 'Microsoft.Graph.Authentication' -and $RequiredVersion -eq '2.41.2') 'Bootstrap module/version must be constrained'
    Check ($Repository -eq 'PSGallery' -and $Scope -eq 'CurrentUser') 'Only CurrentUser installation from PSGallery is allowed'
    Check ($Force -and $AllowClobber -and $PSBoundParameters.ErrorAction -eq 'Stop') 'Installer must fail explicitly'
    if ($script:installMode -eq 'Denied') { throw 'Synthetic CurrentUser access denied' }
    if ($script:installMode -eq 'Unavailable') { throw 'Synthetic PSGallery unavailable' }
    if ($script:installMode -eq 'Invisible') { return }
    New-GraphFixture -Version $RequiredVersion
}
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
function New-GraphFixture([string]$Version) {
    "@{RootModule='Microsoft.Graph.Authentication.psm1';ModuleVersion='$Version';GUID='ead3110d-d8f9-4a6b-976c-2d9ea09cf0fb';FunctionsToExport=@('Get-MgContext');CmdletsToExport=@();AliasesToExport=@();VariablesToExport=@()}" | Set-Content -LiteralPath $manifest -Encoding UTF8
    "function Get-MgContext { throw 'Graph context must not be queried by module prerequisite' }" | Set-Content -LiteralPath $body -Encoding UTF8
}
$script:installCalls = 0
$script:installMode = 'Unavailable'
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
        $script:installCalls = 0
        $messages = @(); $failure = $null
        try { Assert-ProxyAddressesGraphPrerequisite $case.Upload $case.Mode $case.Smtp } catch { $failure = $_ }
        Check (($null -ne $failure) -eq $case.Needs) ($case.Name + ': missing module gate')
        Check ($script:installCalls -eq [int]$case.Needs) ($case.Name + ': install only if required and absent')
        if ($case.Needs) {
            Check ($failure.Exception.Message -match 'Graph Authentication prerequisite failed.*PSModulePath=') ($case.Name + ': actionable diagnostic')
            Check ($failure.Exception.Message -match 'Synthetic PSGallery unavailable.*Scope=CurrentUser') ($case.Name + ': installation failure preserves cause and scope')
        }
    }
    New-GraphFixture -Version '1.2.3'
    foreach ($case in $cases | Where-Object Needs) {
        $script:installCalls = 0
        $messages = @()
        Assert-ProxyAddressesGraphPrerequisite $case.Upload $case.Mode $case.Smtp
        $loaded = Get-Module Microsoft.Graph.Authentication
        Check ($loaded.ModuleBase -eq $moduleBase) ($case.Name + ': only fixture loaded')
        Check (($messages -join '|') -match 'prerequisite OK; PowerShell=.*Version=1.2.3; Path=') ($case.Name + ': runtime/version/path logged')
        Check ($null -ne (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) ($case.Name + ': module remains available to later publication')
        Check ($script:installCalls -eq 0 -and $loaded.Version -eq [version]'1.2.3') ($case.Name + ': available module reused without upgrade')
        Remove-Module Microsoft.Graph.Authentication -Force
    }
    $script:installCalls = 0
    Import-Module -Name $manifest -Global -ErrorAction Stop
    # A loaded module remains usable even if its discovery path is no longer listed.
    $env:PSModulePath = Join-Path $root 'not-a-module-root'
    Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay'
    Check ($script:installCalls -eq 0 -and (Get-Module Microsoft.Graph.Authentication).Version -eq [version]'1.2.3') 'Loaded module reused without installation'
    Remove-Module Microsoft.Graph.Authentication -Force
    $env:PSModulePath = $root
    "throw 'Synthetic incompatible Graph module'" | Set-Content -LiteralPath $body -Encoding UTF8
    $failure = $null
    try { Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay' } catch { $failure = $_ }
    Check ($null -ne $failure -and $failure.Exception.Message -match 'Synthetic incompatible Graph module') 'Import failure must stop prerequisite'
    Check ($script:installCalls -eq 0) 'An incompatible available module must not be replaced automatically'
    Check ($env:PSModulePath -ceq $root) 'Prerequisite must not add PS7 paths'
    $messages = @()
    Assert-ProxyAddressesGraphPrerequisite $false 'SMTP' 'relay'
    Check (($messages -join '|') -match 'prerequisite skipped') 'SMTP-only skips even an incompatible module'
    Remove-Item -LiteralPath $manifest, $body -Force
    foreach ($mode in @('Denied','Invisible','Success')) {
        $script:installMode = $mode; $script:installCalls = 0
        $messages = @(); $failure = $null
        try { Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay' } catch { $failure = $_ }
        Check ($script:installCalls -eq 1) ($mode + ': one installation attempt')
        if ($mode -eq 'Success') {
            Check ($null -eq $failure -and (Get-Module Microsoft.Graph.Authentication).Version -eq [version]'2.41.2') 'Missing module installed then imported'
            Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay'
            Check ($script:installCalls -eq 1) 'Second check does not reinstall'
            Check (($messages -join '|') -match 'absent; installing version 2.41.2.*Scope=CurrentUser') 'Installation scope and version logged'
        }
        else {
            $pattern = if ($mode -eq 'Denied') { 'Synthetic CurrentUser access denied' } else { 'still not discoverable' }
            Check ($null -ne $failure -and $failure.Exception.Message -match $pattern) ($mode + ': collection must remain blocked')
            Check ($null -eq (Get-Module Microsoft.Graph.Authentication)) ($mode + ': no false successful import')
        }
    }
}
finally {
    Get-Module Microsoft.Graph.Authentication | Where-Object { $_.ModuleBase -eq $moduleBase } | Remove-Module -Force
    $env:PSModulePath = $savedPath
    if (-not ([IO.Path]::GetFullPath($root)).StartsWith(([IO.Path]::GetFullPath([IO.Path]::GetTempPath())).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe fixture cleanup path' }
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
Write-Output ("PASS: {0} ProxyAddresses Graph prerequisite checks; {1}; synthetic only." -f $checks, $PSVersionTable.PSVersion)
