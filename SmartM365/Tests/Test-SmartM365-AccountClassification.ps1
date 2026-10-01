<#
.SYNOPSIS
    Validates the published account-classification template and, when present, the private
    AccountClassification.local.json(.txt) rules.
.VERSION
1.1
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$smartInventoryRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartInventory'
Import-Module (Join-Path $smartInventoryRoot 'Common\SmartM365.AccountClassification.psd1') -MinimumVersion '1.0.0' -Force -ErrorAction Stop
$templatePath = Join-Path $smartInventoryRoot 'Config\AccountClassification.local.json.template'
$expectedAccountTypes = @(
    'Named Account', 'Ext Account', 'Service Account', 'Shared Mailbox',
    'Room Mailbox', 'System Account', 'Admin Account', 'Generic Account',
    'Unclassified Account'
)
$expectedRules = @('SharedMailbox', 'RoomMailbox', 'ExternalAccount', 'SystemAccount', 'ServiceAccount', 'AdminAccount', 'GenericAccount', 'NamedAccount')

function Test-AccountClassificationFile {
    param([Parameter(Mandatory)][string]$Path)
    $config = Read-SmartM365AccountClassification -Path $Path
    $populationNames = @('Human', 'NonHuman', 'ReviewRequired')
    $membership = @{}
    foreach ($populationName in $populationNames) {
        foreach ($accountType in @($config.Population[$populationName])) {
            if ($membership.ContainsKey($accountType)) {
                throw "Account type '$accountType' belongs to both '$($membership[$accountType])' and '$populationName'."
            }
            $membership[$accountType] = $populationName
        }
    }
    $missingAccountTypes = @($expectedAccountTypes | Where-Object { -not $membership.ContainsKey($_) })
    if ($missingAccountTypes.Count -gt 0) { throw "Unmapped account types: $($missingAccountTypes -join ', ')." }
    $missingRules = @($expectedRules | Where-Object { -not $config.AccountTypeRules.ContainsKey($_) })
    if ($missingRules.Count -gt 0) { throw "Missing account type rules: $($missingRules -join ', ')." }
    if (-not $config.ContainsKey('LegacyLikelyServiceAccount')) { throw 'LegacyLikelyServiceAccount is required.' }
    [pscustomobject]@{
        ConfigPath = $config.SourcePath
        SchemaVersion = [string]$config.SchemaVersion
        RuleVersion = [string]$config.RuleVersion
        AccountTypes = $membership.Count
        Populations = $populationNames.Count
        Result = 'Passed'
    }
}

# The published template must stay valid; the private rules are checked when available.
# Copy the template to a temporary .json so the shared reader validates its JSON as a configuration.
$templateCopy = Join-Path ([IO.Path]::GetTempPath()) ('AccountClassification-template-{0}.json' -f [guid]::NewGuid().ToString('N'))
try {
    Copy-Item -LiteralPath $templatePath -Destination $templateCopy
    Test-AccountClassificationFile -Path $templateCopy
}
finally { Remove-Item -LiteralPath $templateCopy -Force -ErrorAction SilentlyContinue }

$privatePath = if ($ConfigPath) { $ConfigPath } else { Get-SmartM365AccountClassificationDefaultPath }
$privateExists = $true
try { $null = Resolve-SmartM365AccountClassificationPath -Path $privatePath } catch { $privateExists = $false }
if ($privateExists) { Test-AccountClassificationFile -Path $privatePath }
elseif ($ConfigPath) { throw "Account classification configuration not found: $ConfigPath" }
else { Write-Output 'Private AccountClassification.local.json not present; only the published template was validated.' }

# Missing private rules must fail closed, never fall back to the template.
$missingRejected = $false
try { $null = Read-SmartM365AccountClassification -Path (Join-Path ([IO.Path]::GetTempPath()) ('absent-{0}.local.json' -f [guid]::NewGuid().ToString('N'))) }
catch { $missingRejected = $_.Exception.Message -like '*AccountClassification.local.json.template*' }
if (-not $missingRejected) { throw 'A missing private configuration did not fail closed.' }
