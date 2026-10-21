#Requires -Version 7.0
<#
.SYNOPSIS
  Offline checks for users deleted from the directory while the licensing run collects them.
.VERSION
1.0.0
.NOTES
  Synthetic identifiers only (GUIDs and example.invalid). Graph calls are replaced by mocks; no tenant access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Offline mocks keep the production command signatures.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Retry delays and warnings are silenced for the offline run.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Offline mocks change no state.')]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$null, [ref]$null)
$names = @('Invoke-GraphWithRetry', 'Test-LicensesGraphNotFound', 'Register-LicensesUserFailure', 'Assert-LicensesUserResolution')
$definitions = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name }, $true))
if ($definitions.Count -ne $names.Count) { throw 'Deleted-user functions are missing.' }
foreach ($definition in $definitions) { . ([scriptblock]::Create($definition.Extent.Text)) }

$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$script:Log = [System.Collections.Generic.List[string]]::new()
function WriteLog { param([string]$Message, [string]$Level) $script:Log.Add("$Level $Message") }
function Start-Sleep { param([int]$Seconds) $null = $Seconds }
function Write-Warning { param([string]$Message) $null = $Message }
$script:LookupCalls = 0
$script:LookupBehavior = 'NotFound'
function Get-MgUser {
  param([string]$UserId, [string[]]$Property, $ErrorAction)
  $script:LookupCalls++
  switch ($script:LookupBehavior) {
    'NotFound' { throw [Exception]::new("Resource '$UserId' does not exist or one of its queried reference-property objects are not present. Status: 404 (NotFound) ErrorCode: Request_ResourceNotFound") }
    'Exists' { return [pscustomobject]@{ Id = $UserId } }
    'Throttled' { throw [Exception]::new('Status: 429 (TooManyRequests) ErrorCode: TooManyRequests') }
    'Forbidden' { throw [Exception]::new('Status: 403 (Forbidden) ErrorCode: Authorization_RequestDenied') }
  }
}
function Get-SyntheticErrorRecord([string]$Message, [Exception]$Inner) {
  try { if ($Inner) { throw [Exception]::new($Message, $Inner) } else { throw [Exception]::new($Message) } } catch { return $_ }
}
$notFound = Get-SyntheticErrorRecord "Resource 'synthetic' does not exist or one of its queried reference-property objects are not present. Status: 404 (NotFound) ErrorCode: Request_ResourceNotFound"
$forbidden = Get-SyntheticErrorRecord 'Status: 403 (Forbidden) ErrorCode: Authorization_RequestDenied'
function Get-SyntheticPathList { $paths = [System.Collections.Generic.List[object]]::new(); $paths.Add([pscustomobject]@{ UserId = 'kept' }); return ,$paths }
function Invoke-Case([string]$Behavior, $ErrorRecord, [bool]$DetailStage = $true) {
  $script:LookupBehavior = $Behavior; $script:LookupCalls = 0
  $paths = Get-SyntheticPathList
  $before = $paths.Count
  $paths.Add([pscustomobject]@{ UserId = '00000000-0000-0000-0000-000000000001' })
  $state = @{ UserResolutionFailures = 0; UsersDeletedDuringCollection = 0 }
  $result = Register-LicensesUserFailure -ErrorRecord $ErrorRecord -UserId '00000000-0000-0000-0000-000000000001' -LicenseDetailStage $DetailStage `
    -AssignmentPaths $paths -AssignmentPathCountBefore $before -State $state
  return [pscustomobject]@{ Result = $result; State = $state; Paths = $paths; Calls = $script:LookupCalls }
}

# Recognition of a Graph 404, also through an inner exception or a status property.
Check (Test-LicensesGraphNotFound -ErrorRecord $notFound) 'A Request_ResourceNotFound error was not recognized.'
Check (Test-LicensesGraphNotFound -ErrorRecord (Get-SyntheticErrorRecord 'Wrapper' ([Exception]::new('Status: 404 (NotFound)')))) 'An inner 404 was not recognized.'
$statusException = [Exception]::new('Graph error'); $statusException | Add-Member -NotePropertyName ResponseStatusCode -NotePropertyValue 404
Check (Test-LicensesGraphNotFound -ErrorRecord $statusException) 'A 404 status property was not recognized.'
Check (-not (Test-LicensesGraphNotFound -ErrorRecord $forbidden)) 'A 403 was taken for a 404.'

# 1. Confirmed deletion: excluded, counted once, its assignment paths removed, run allowed.
$case = Invoke-Case 'NotFound' $notFound
Check ($case.Result -eq 'Deleted' -and $case.State.UsersDeletedDuringCollection -eq 1 -and $case.State.UserResolutionFailures -eq 0) 'A confirmed deletion was not counted as deleted.'
Check ($case.Calls -eq 1) 'The deletion was not confirmed by a directory lookup.'
Check ($case.Paths.Count -eq 1 -and $case.Paths[0].UserId -eq 'kept') 'The deleted user assignment paths were not removed, or other paths were removed.'
$qualification = Assert-LicensesUserResolution -State $case.State -ListedUsers 47096
Check ($qualification -eq 'UsersDeletedDuringCollection=1') 'A confirmed deletion within the limit stopped the run.'
Check (@($script:Log | Where-Object { $_ -like 'WARNING*deleted from the directory during collection*' }).Count -ge 1) 'The deletion is not logged as a warning.'
Check (@($script:Log | Where-Object { $_ -match '@' }).Count -eq 0) 'A user principal name reached the deleted-user log.'

# 2. 404 not confirmed: the user still exists, the run fails.
$case = Invoke-Case 'Exists' $notFound
Check ($case.Result -eq 'Failure' -and $case.State.UserResolutionFailures -eq 1 -and $case.State.UsersDeletedDuringCollection -eq 0) 'An existing user with a 404 was excluded.'
$failed = $false; try { Assert-LicensesUserResolution -State $case.State -ListedUsers 47096 | Out-Null } catch { $failed = $_.Exception.Message -like '*could not be resolved*' }
Check $failed 'An unconfirmed 404 did not stop the run.'

# 3. The confirmation lookup itself fails (throttling, permission): not confirmed, the run fails.
foreach ($behavior in @('Throttled', 'Forbidden')) {
  $case = Invoke-Case $behavior $notFound
  Check ($case.Result -eq 'Failure' -and $case.State.UserResolutionFailures -eq 1) "A $behavior confirmation lookup excluded the user."
}

# 4. Any other error, or a 404 outside the license-detail call, fails without a deletion lookup.
$case = Invoke-Case 'NotFound' $forbidden
Check ($case.Result -eq 'Failure' -and $case.Calls -eq 0 -and $case.State.UserResolutionFailures -eq 1) 'A permission error was treated as a deletion.'
$case = Invoke-Case 'NotFound' $notFound $false
Check ($case.Result -eq 'Failure' -and $case.Calls -eq 0) 'A 404 outside the license-detail call was treated as a deletion.'
$case = Invoke-Case 'NotFound' (Get-SyntheticErrorRecord 'Status: 503 (ServiceUnavailable)')
Check ($case.Result -eq 'Failure' -and $case.Calls -eq 0) 'A service error was treated as a deletion.'

# 5. Mass deletion limit: 1 % of listed users, capped at 100 by default; both are configurable.
foreach ($limitCase in @(
    @{ Deleted = 100; Listed = 47096; Fails = $false }, @{ Deleted = 101; Listed = 47096; Fails = $true },
    @{ Deleted = 5; Listed = 500; Fails = $false }, @{ Deleted = 6; Listed = 500; Fails = $true },
    @{ Deleted = 11; Listed = 47096; MaxCount = 10; Fails = $true }, @{ Deleted = 30; Listed = 1000; MaxPercent = 5; Fails = $false })) {
  $state = @{ UserResolutionFailures = 0; UsersDeletedDuringCollection = $limitCase.Deleted }
  $arguments = @{ State = $state; ListedUsers = $limitCase.Listed }
  if ($limitCase.ContainsKey('MaxCount')) { $arguments.MaxDeletedCount = $limitCase.MaxCount }
  if ($limitCase.ContainsKey('MaxPercent')) { $arguments.MaxDeletedPercent = $limitCase.MaxPercent }
  $threw = $false; try { Assert-LicensesUserResolution @arguments | Out-Null } catch { $threw = $_.Exception.Message -like '*above the limit*' }
  Check ($threw -eq $limitCase.Fails) ("Deleted-user limit misapplied: {0} deleted of {1}." -f $limitCase.Deleted, $limitCase.Listed)
}
Check ($null -eq (Assert-LicensesUserResolution -State @{ UserResolutionFailures = 0; UsersDeletedDuringCollection = 0 } -ListedUsers 10)) 'A run without deletion returned a qualification.'

# 6. Collector wiring: stage flag around the license details, classification in the catch, limit before export,
#    qualification in the receipt, counter in the completion summary, one version.
$text = [IO.File]::ReadAllText($inventoryPath)
Check ($text -match '\$licenseDetailStage = \$true\s+\$licenseDetails = Invoke-GraphWithRetry \{ Get-MgUserLicenseDetail[^\r\n]*\s+\$licenseDetailStage = \$false') 'The license-detail call is not marked as the deletion stage.'
Check ($text -match 'Register-LicensesUserFailure -ErrorRecord \$_ ') 'The user loop does not classify its failures.'
Check ($text -match '\$usersDeletedQualification = Assert-LicensesUserResolution') 'The run does not apply the deleted-user rules.'
Check ($text -match "-Qualifications @\(\`$usersDeletedQualification") 'The deleted-user count does not reach the receipt.'
Check ($text -match 'Add-SmartM365SourceRunNote -Message \(''\{0\}: users deleted') 'The deleted-user count does not reach the run state.'
Check ($text -match '-UsersDeletedDuringCollection \(\[int\]\$userResolution\.UsersDeletedDuringCollection\)') 'The completion summary does not show the deleted-user count.'
$version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
Check ($version -eq '1.55' -and $text -match '\$ScriptVersion = "1\.55"') 'Header and runtime versions are not 1.55.'

"PASS: License deleted-user offline checks ($script:checks checks)."
