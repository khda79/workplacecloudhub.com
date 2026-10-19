<#
.SYNOPSIS
  Offline checks: the five usage reports read by the license report are qualified only by the UserActivity source receipt.
.VERSION
1.0.0
.NOTES
  Receipts are written by the real SmartM365.Core receipt commands in an isolated temporary folder.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$userActivityPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\Usage\SmartM365-M365UserActivity-Inventory.ps1')).Path
$coreModulePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1')).Path
Import-Module -Name $coreModulePath -MinimumVersion '1.0.90' -Force -ErrorAction Stop

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
$names = @(
  'Get-LicensesSuiteProduct', 'Get-LicensesOptionalProperty', 'ConvertTo-LicensesActivityDate', 'ConvertTo-LicensesMailboxSizeGb',
  'Get-LicensesQualifiedPartialAdCoverage', 'Get-LicensesCsvSource', 'Import-LicensesSourceCsv', 'Read-LicensesIndexedSource',
  'Get-LicensesMailboxGapSummary', 'Get-LicensesAdAccountActivitySummary', 'Get-LicensesFocusedUsageRows'
)
$definitions = @($ast.FindAll({
  param($node)
  $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name
}, $true))
if ($definitions.Count -ne $names.Count) { throw 'License usage functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }
function WriteLog { param([string]$Message, [string]$Level) }

$script:Checks = 0
function Check {
  param([bool]$Condition, [string]$Label)
  if (-not $Condition) { throw "FAIL: $Label" }
  $script:Checks++
}

$usageFiles = @('M365_Users_Activity.csv','M365_Mailbox_Usage.csv','M365_Email_Activity.csv','M365_Apps_Usage_180D.csv','M365_OneDrive_Usage.csv')
$receiptName = 'SmartInventory_SmartM365-M365UserActivity-Inventory.current.json.txt'
$today = [datetime]::UtcNow.Date
$refresh = $today.AddDays(-3).ToString('yyyy-MM-dd')

$saved = @{}
foreach ($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')) {
  $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
  $saved[$name] = @{ Exists=($null -ne $variable); Value=$(if ($variable) { $variable.Value } else { $null }) }
}

function Set-TestTenant {
  param([string]$TenantKey)
  $global:SmartM365TenantKey = $TenantKey; $global:SmartM365OrganizationKey = 'contoso'
  $global:SmartM365EnvironmentKey = 'prod'; $global:SmartM365TenantId = 'synthetic-tenant'
}

function New-UsageFixture {
  param([string]$Name, [string]$TenantKey = 'prod', [string]$OneDrivePeriod = '180')
  $folder = Join-Path $script:Root $Name
  New-Item -ItemType Directory -Path $folder -Force | Out-Null
  $t = $TenantKey
  $upn = 'u1@example.invalid'
  @([pscustomobject]@{TenantKey='prod';UserId='u1';SkuPartNumber='SPE_E3'}) | Export-Csv -LiteralPath (Join-Path $folder 'M365_Licenses_Users.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;UserPrincipalName=$upn;ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$refresh;IsDeleted='False'}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_Users_Activity.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;'User Principal Name'=$upn;'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$refresh;'Is Deleted'='False'}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_Mailbox_Usage.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;'User Principal Name'=$upn;'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$refresh;'Is Deleted'='False';'Send Count'=1;'Read Count'=1}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_Email_Activity.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;'User Principal Name'=$upn;'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$refresh;Windows='No';Mac='No'}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_Apps_Usage_180D.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;'Owner Principal Name'=$upn;'Is Deleted'='False';'Storage Used (Byte)'='1073741824';'Report Refresh Date'=$refresh;'Report Period'=$OneDrivePeriod}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_OneDrive_Usage.csv') -NoTypeInformation -Encoding utf8NoBOM
  return $folder
}

function New-UserActivityReceipt {
  param([string]$Folder, [string[]]$Files, [string]$TenantKey = 'prod', [switch]$LeavePublishing)
  # Same producer, registry and scope as SmartM365-M365UserActivity-Inventory.ps1.
  Set-TestTenant $TenantKey
  $global:csvGeneratedPaths = $null
  Start-SmartM365CmdbSourceReceipt -ScriptPath $userActivityPath -SourceRootPath $Folder
  foreach ($file in $Files) { Register-SmartM365SourceCsv -Path (Join-Path $Folder $file) }
  if ($LeavePublishing) { return }
  Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope 'CMDB:activity'
  return (Complete-SmartM365CmdbSourceReceipt -Status 'Completed')
}

function Get-UsageSources {
  param([string]$Folder)
  # The license file is proven by a current-run publication, as during the licensing run.
  $licensePath = [IO.Path]::GetFullPath((Join-Path $Folder 'M365_Licenses_Users.csv'))
  $hashes = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
  $hashes[$licensePath] = (Get-FileHash -LiteralPath $licensePath -Algorithm SHA256).Hash
  $proof = [pscustomobject]@{ RunId='license-run'; Hashes=$hashes }
  $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $Folder -ExpectedTenantKey 'prod' -AsOfUtc $today -CurrentRunPublication $proof
  $byName = @{}
  foreach ($source in $usage.Sources) { if ($usageFiles -contains $source.Name) { $byName[$source.Name] = $source } }
  Check ($byName.Count -eq 5) 'The five usage reports are evaluated'
  return [pscustomobject]@{ Usage=$usage; Sources=$byName }
}

function Test-Unqualified {
  param($Source, [string]$Reason, [string]$Label)
  Check (-not $Source.Ready -and -not $Source.Forced -and $Source.Reason -ceq $Reason) "$Label ($($Source.Name)): expected '$Reason', got Ready=$($Source.Ready) '$($Source.Reason)'"
}

$script:Root = Join-Path $env:TEMP ('SmartM365-LicenseUsageReceipts-' + [guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Path $script:Root -Force | Out-Null

  # 1. Every report listed by a completed receipt with the same hash is qualified, never provisional.
  $folder = New-UsageFixture 'listed'
  Check ((New-UserActivityReceipt -Folder $folder -Files $usageFiles) -like '*.current.json.txt') 'Core completes the UserActivity receipt'
  $listed = Get-UsageSources $folder
  foreach ($name in $usageFiles) {
    $source = $listed.Sources[$name]
    Check ($source.Ready -and -not $source.Forced -and $source.Provenance -ceq 'collector completed' -and $source.Reason -eq '') "Listed report is qualified: $name"
    Check ($source.Date -eq $refresh) "Report date is its Report Refresh Date: $name"
  }
  Check ($listed.Usage.DowngradeReview.Unknown -eq 1) 'E3 review stays N/D without the other sources'

  # 2. A report present in DATA-LAST but absent from the receipt is not qualified.
  foreach ($name in @($usageFiles | Where-Object { $_ -ne 'M365_Users_Activity.csv' })) {
    $folder = New-UsageFixture ('absent-' + [IO.Path]::GetFileNameWithoutExtension($name))
    [void](New-UserActivityReceipt -Folder $folder -Files @($usageFiles | Where-Object { $_ -ne $name }))
    $result = Get-UsageSources $folder
    Test-Unqualified $result.Sources[$name] 'file missing from collector receipt' 'File absent from a Core receipt'
    foreach ($other in @($usageFiles | Where-Object { $_ -ne $name })) { Check $result.Sources[$other].Ready "Other listed report stays qualified: $other without $name" }
  }
  # M365_Users_Activity.csv is required by the producer registry: Core refuses a receipt without it,
  # so its absence is checked on a completed Core receipt whose record for that file is removed.
  $folder = New-UsageFixture 'absent-M365_Users_Activity'
  Check ($null -eq (New-UserActivityReceipt -Folder $folder -Files @($usageFiles | Where-Object { $_ -ne 'M365_Users_Activity.csv' })) -or
    -not (Test-Path -LiteralPath (Join-Path $folder $receiptName))) 'Core refuses a receipt without the required activity report'
  Test-Unqualified (Get-UsageSources $folder).Sources['M365_Users_Activity.csv'] 'collector receipt is missing' 'Refused Core receipt'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  $receiptPath = Join-Path $folder $receiptName
  $document = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
  $document.Files = @($document.Files | Where-Object File -ne 'M365_Users_Activity.csv')
  $document | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $receiptPath -Encoding utf8NoBOM
  Test-Unqualified (Get-UsageSources $folder).Sources['M365_Users_Activity.csv'] 'file missing from collector receipt' 'File absent from a Core receipt'

  # 3. A report whose bytes differ from the receipt is not qualified.
  $folder = New-UsageFixture 'hash'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  foreach ($name in $usageFiles) { Add-Content -LiteralPath (Join-Path $folder $name) -Value '' -NoNewline:$false }
  $hash = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $hash.Sources[$name] 'file hash differs from collector receipt' 'Changed file' }

  # 4. No receipt: fail-closed, the CSV alone is never used.
  $folder = New-UsageFixture 'missing'
  $missing = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $missing.Sources[$name] 'collector receipt is missing' 'Missing receipt' }

  # 5. A run that is publishing (or failed after publication) withholds the previous completed receipt.
  $folder = New-UsageFixture 'publishing'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  New-UserActivityReceipt -Folder $folder -Files @('M365_OneDrive_Usage.csv') -LeavePublishing
  try {
    $publishing = Get-UsageSources $folder
    foreach ($name in $usageFiles) { Test-Unqualified $publishing.Sources[$name] 'collector receipt is unreadable' 'Publishing run' }
  }
  finally { [void](Complete-SmartM365CmdbSourceReceipt -Status 'Failed') }
  $failed = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $failed.Sources[$name] 'collector receipt is unreadable' 'Run failed after publication' }

  # 6. A current receipt that is not Completed is not accepted.
  $folder = New-UsageFixture 'running'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  $receiptPath = Join-Path $folder $receiptName
  $document = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
  $document.Status = 'Running'
  $document | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $receiptPath -Encoding utf8NoBOM
  $running = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $running.Sources[$name] 'collector receipt is Running or partial' 'Running receipt' }

  # 7. Rows of another tenant, even under their own completed receipt, are not used.
  $folder = New-UsageFixture 'tenant' -TenantKey 'other'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles -TenantKey 'other')
  $tenant = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $tenant.Sources[$name] 'another tenant is present' 'Foreign tenant' }

  # 8. OneDrive keeps the 180-day period of the other usage reports; another period is refused.
  $folder = New-UsageFixture 'period' -OneDrivePeriod '30'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  $period = Get-UsageSources $folder
  Test-Unqualified $period.Sources['M365_OneDrive_Usage.csv'] 'unexpected report period' 'OneDrive period'
  Check $period.Sources['M365_Apps_Usage_180D.csv'].Ready 'Other reports are not affected by the OneDrive period'

  Write-Host ("PASS: license usage reports are qualified by the UserActivity receipt ({0} checks)." -f $script:Checks)
}
finally {
  foreach ($name in $saved.Keys) {
    if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
    else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
  }
  if (Test-Path -LiteralPath $script:Root) { Remove-Item -LiteralPath $script:Root -Recurse -Force }
}
