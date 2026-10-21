#Requires -Version 7.0
<#
.SYNOPSIS
  Offline checks that license assignment paths and user service-plan states come from one snapshot per user:
  a license assigned between the user list and the license details is read again for that user, an
  assignment is never inferred, and publication stops when a service-plan pair has no assignment path.
.VERSION
1.0.0
.NOTES
  Synthetic identifiers only (GUIDs). Graph reads are replaced by script blocks; no tenant access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Synthetic record builders change no state.')]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$null, [ref]$null)
$names = @('Get-LicensesUncoveredSkuId', 'Resolve-LicensesUserSnapshot', 'Assert-LicensesPlanStateCoverage', 'ConvertTo-LicensesAssignmentPathRows')
$definitions = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name }, $true))
if ($definitions.Count -ne $names.Count) { throw 'License snapshot functions are missing.' }
foreach ($definition in $definitions) { . ([scriptblock]::Create($definition.Extent.Text)) }

$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$userId = [guid]::NewGuid().ToString()
$e3 = [guid]::NewGuid().ToString()
$f1 = [guid]::NewGuid().ToString()
$group = [guid]::NewGuid().ToString()
function New-State([string]$Sku, [string]$Group = '') { [pscustomobject]@{ SkuId = $Sku; AssignedByGroup = $Group; State = 'Active'; Error = 'None'; DisabledPlans = @(); LastUpdatedDateTime = $null } }
function New-Detail([string]$Sku) { [pscustomobject]@{ SkuId = $Sku; ServicePlans = @([pscustomobject]@{ ServicePlanId = [guid]::NewGuid().ToString(); ServicePlanName = 'EXCHANGE_S'; ProvisioningStatus = 'Success' }) } }

# 1. User list read at the start: E3 only. License details read later: E3 and F1 (F1 assigned in between).
$listed = [pscustomobject]@{ Id = $userId; LicenseAssignmentStates = @(New-State $e3); AssignedLicenses = @() }
$details = @((New-Detail $e3), (New-Detail $f1))
$script:userReads = 0; $script:detailReads = 0
$fresh = [pscustomobject]@{ Id = $userId; LicenseAssignmentStates = @((New-State $e3), (New-State $f1 $group)); AssignedLicenses = @() }
$snapshot = Resolve-LicensesUserSnapshot -User $listed -LicenseDetails $details -ReadUser { $script:userReads++; $fresh } -ReadLicenseDetails { $script:detailReads++; $details }
Check ($snapshot.Refreshed -and $script:userReads -eq 1 -and $script:detailReads -eq 1) 'A license assigned during collection was not read again for the user.'
Check ((@($snapshot.UncoveredSkuIds) -join ',') -eq $f1) 'The changed SKU is not reported.'
$paths = [System.Collections.Generic.List[object]]::new()
foreach ($row in @(ConvertTo-LicensesAssignmentPathRows -User $snapshot.User)) { $paths.Add($row) }
$pairs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($detail in $snapshot.LicenseDetails) { [void]$pairs.Add(('{0}|{1}' -f $userId, $detail.SkuId)) }
Assert-LicensesPlanStateCoverage -AssignmentPaths $paths -PlanStatePairs $pairs
$script:checks++
Check (@($paths | Where-Object { $_.SkuId -eq $f1 -and $_.AssignmentRoute -eq 'Group' -and $_.AssignedByGroupId -eq $group }).Count -eq 1) 'The fresh assignment state is not the one published.'

# 2. Without the fresh read (previous behavior), the same data breaks the preparation rule.
$stalePaths = [System.Collections.Generic.List[object]]::new()
foreach ($row in @(ConvertTo-LicensesAssignmentPathRows -User $listed)) { $stalePaths.Add($row) }
$failure = ''
try { Assert-LicensesPlanStateCoverage -AssignmentPaths $stalePaths -PlanStatePairs $pairs } catch { $failure = $_.Exception.Message }
Check ($failure -like "License collection is inconsistent: 1 user/SKU pair(s)*$userId|$f1*Current CSV publication is stopped.") "An uncovered service-plan pair was not stopped clearly: $failure"

# 3. Consistent user: no extra Graph read.
$script:userReads = 0; $script:detailReads = 0
$steady = Resolve-LicensesUserSnapshot -User $fresh -LicenseDetails $details -ReadUser { $script:userReads++; $fresh } -ReadLicenseDetails { $script:detailReads++; $details }
Check (-not $steady.Refreshed -and $script:userReads -eq 0 -and $script:detailReads -eq 0) 'A consistent user was read again.'

# 4. A license removed in between (assigned in the list, gone from the details) needs no re-read.
$removed = Resolve-LicensesUserSnapshot -User $fresh -LicenseDetails @(New-Detail $e3) -ReadUser { throw 'unexpected read' } -ReadLicenseDetails { throw 'unexpected read' }
Check (-not $removed.Refreshed) 'A removed license forced a re-read.'

# 5. The fresh read still lacks the SKU: the user fails; no assignment is inferred from license details.
$failure = ''
try { $null = Resolve-LicensesUserSnapshot -User $listed -LicenseDetails $details -ReadUser { $listed } -ReadLicenseDetails { $details } } catch { $failure = $_.Exception.Message }
Check ($failure -like "*user '$userId' list SKU(s) $f1 without a license assignment state after a fresh read; no assignment is inferred.") "A persistent gap was not reported: $failure"

# 6. Static wiring of the collector: snapshot right after the license details, pairs recorded with the plan
#    states, coverage checked before the assignment paths are exported; one version.
$text = [IO.File]::ReadAllText($inventoryPath)
Check ($text -match '(?s)\$licenseDetails = Invoke-GraphWithRetry \{ Get-MgUserLicenseDetail[^\r\n]*\r?\n(\s*#[^\r\n]*\r?\n)*\s*\$snapshot = Resolve-LicensesUserSnapshot') 'The collector does not resolve one snapshot after the license details.'
Check ($text -match "\[void\]\`$servicePlanStatePairs\.Add\(\('\{0\}\|\{1\}' -f \`$uid, \`$skuId\)\)") 'Service-plan pairs are not recorded.'
$assertIndex = $text.IndexOf('Assert-LicensesPlanStateCoverage -AssignmentPaths $assignmentPaths -PlanStatePairs $servicePlanStatePairs', [StringComparison]::Ordinal)
$exportIndex = $text.IndexOf('Export-SmartM365Csv -Data $assignmentPaths.ToArray()', [StringComparison]::Ordinal)
Check ($assertIndex -gt 0 -and $exportIndex -gt $assertIndex) 'Coverage is not checked before the assignment paths are published.'
$version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
Check ($version -and [version]$version -ge [version]'1.58' -and $text -match ('\$ScriptVersion = "' + [regex]::Escape($version) + '"')) 'Header and runtime versions differ or predate 1.58.'

"PASS: License assignment snapshot offline checks ($script:checks checks)."
