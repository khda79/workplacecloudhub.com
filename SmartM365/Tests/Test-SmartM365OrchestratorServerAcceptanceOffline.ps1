<#
.SYNOPSIS
    Offline checks of the orchestrator server acceptance script: dataset checks, statuses,
    module assessment, child output parsing, Graph roles and job eligibility. Nothing is read
    from a directory, a tenant or a share.
.VERSION
    1.0.0
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartInventory\Maintenance\SmartM365-OrchestratorServerAcceptance.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) { throw 'The acceptance script has PowerShell parse errors.' }
foreach ($definition in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -like '*-SmartM365Acceptance*' }, $true))) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
foreach ($assignment in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -in @('$script:GuidPattern', '$script:SmtpPattern', '$script:ResultMarker', '$script:ExpectedModules') }, $false))) {
    . ([scriptblock]::Create($assignment.Extent.Text))
}

$script:Checks = 0
function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "Assertion failed: $Message" }
    $script:Checks++
}

# Worst status and latency/clock thresholds.
Assert-True ((Get-SmartM365AcceptanceWorstStatus -Status @('PASS', 'WARN', 'SKIPPED')) -eq 'WARN') 'WARN wins over PASS and SKIPPED.'
Assert-True ((Get-SmartM365AcceptanceWorstStatus -Status @('WARN', 'FAIL')) -eq 'FAIL') 'FAIL wins.'
Assert-True ((Get-SmartM365AcceptanceWorstStatus -Status @()) -eq 'PASS') 'No status is PASS.'
Assert-True ((Get-SmartM365AcceptancePercentile -Values @(1, 2, 3, 4, 100) -Percent 50) -eq 3) 'Median of five values.'
Assert-True ((Get-SmartM365AcceptancePercentile -Values @(1, 2, 3, 4, 100) -Percent 95) -eq 100) 'p95 of five values.'
Assert-True ($null -eq (Get-SmartM365AcceptancePercentile -Values @())) 'No sample has no percentile.'
Assert-True ((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds 9 -WarningMilliseconds 15 -FailureMilliseconds 50) -eq 'PASS') 'Fast share passes.'
Assert-True ((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds 23 -WarningMilliseconds 15 -FailureMilliseconds 50) -eq 'WARN') 'The 23 ms incident share warns.'
Assert-True ((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds 50 -WarningMilliseconds 15 -FailureMilliseconds 50) -eq 'FAIL') 'Very slow share fails.'
Assert-True ((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds $null -WarningMilliseconds 15 -FailureMilliseconds 50) -eq 'FAIL') 'Unmeasured latency fails.'
Assert-True ((Get-SmartM365AcceptanceClockStatus -OffsetSeconds -12) -eq 'PASS') 'Small clock offset passes.'
Assert-True ((Get-SmartM365AcceptanceClockStatus -OffsetSeconds 90) -eq 'WARN') 'Clock offset of 90 s warns.'
Assert-True ((Get-SmartM365AcceptanceClockStatus -OffsetSeconds -301) -eq 'FAIL') 'Clock offset beyond 5 minutes fails.'

# Dataset checks.
function New-AdRow([string]$Guid, [string]$Sam, [string]$Upn, [string]$Created = '2015-03-01T08:00:00.0000000Z') {
    [pscustomobject]@{ ObjectGUID = $Guid; SamAccountName = $Sam; UserPrincipalName = $Upn; WhenCreated = $Created }
}
$spec = Get-SmartM365AcceptanceDatasetSpec -Name 'ADUsers'
$good = @(1..5 | ForEach-Object { New-AdRow ([guid]::NewGuid().ToString()) "user$_" "user$_@example.test" })
$check = Test-SmartM365AcceptanceDataset -Rows $good -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'PASS' -and $check.Count -eq 5) "Clean AD rows pass ($($check.Issues -join '; '))."
$check = Test-SmartM365AcceptanceDataset -Rows @() -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL') 'No AD row fails.'
$check = Test-SmartM365AcceptanceDataset -Rows $good -SampleSize 4 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL' -and ($check.Issues -join ' ') -match 'limit of 4') 'More rows than the limit fails.'
$duplicate = @($good + (New-AdRow $good[0].ObjectGUID 'other' 'other@example.test'))
$check = Test-SmartM365AcceptanceDataset -Rows $duplicate -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL' -and ($check.Issues -join ' ') -match 'ObjectGUID has 1 duplicated') 'Duplicated ObjectGUID fails.'
$emptyGuid = @($good + (New-AdRow '' 'nogid' 'nogid@example.test'))
$check = Test-SmartM365AcceptanceDataset -Rows $emptyGuid -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL' -and ($check.Issues -join ' ') -match 'ObjectGUID empty in 1') 'Empty ObjectGUID fails.'
$badGuid = @($good + (New-AdRow 'not-a-guid' 'badgid' 'badgid@example.test'))
$check = Test-SmartM365AcceptanceDataset -Rows $badGuid -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL') 'Malformed required GUID fails.'
$fewUpn = @(1..4 | ForEach-Object { New-AdRow ([guid]::NewGuid().ToString()) "svc$_" '' }) + $good[0]
$check = Test-SmartM365AcceptanceDataset -Rows $fewUpn -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'WARN' -and ($check.Issues -join ' ') -match 'UserPrincipalName filled in 20 %') 'Mostly empty optional UPN warns.'
$future = @($good + (New-AdRow ([guid]::NewGuid().ToString()) 'future' 'future@example.test' '2099-01-01T00:00:00Z'))
$check = Test-SmartM365AcceptanceDataset -Rows $future -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'WARN' -and ($check.Issues -join ' ') -match 'implausible date') 'Future creation date warns.'
$sites = Get-SmartM365AcceptanceDatasetSpec -Name 'SharePointSites'
Assert-True ((Test-SmartM365AcceptanceDataset -Rows @() -SampleSize 10 -Columns $sites.Columns -AllowEmpty:([bool]$sites.AllowEmpty)).Status -eq 'WARN') 'An empty SharePoint farm only warns.'
$mailboxes = Get-SmartM365AcceptanceDatasetSpec -Name 'EXOMailboxes'
$mailboxRows = @([pscustomobject]@{ Guid = [guid]::NewGuid().ToString(); PrimarySmtpAddress = 'no-at-sign'; ExternalDirectoryObjectId = [guid]::NewGuid().ToString(); RecipientTypeDetails = 'UserMailbox' })
Assert-True ((Test-SmartM365AcceptanceDataset -Rows $mailboxRows -SampleSize 100 -Columns $mailboxes.Columns).Status -eq 'FAIL') 'Malformed SMTP address fails.'
foreach ($name in @('ADUsers', 'GraphUsers', 'EntraDevices', 'IntuneDevices', 'EXOMailboxes', 'OnPremMailboxes', 'SharePointSites', 'TeamsUsers', 'AdaxesUsers')) {
    Assert-True (@((Get-SmartM365AcceptanceDatasetSpec -Name $name).Columns | Where-Object { $_.PSObject.Properties['Required'] -and $_.Required }).Count -ge 1) "$name declares a required key column."
}

# Child output parsing and assessment.
$json = ConvertTo-Json -Compress -Depth 8 -InputObject @{ Ready = $true; Detail = 'forest example.test'; Datasets = @(
        @{ Name = 'ADUsers'; Scope = 'a.example.test'; Error = ''; DurationSeconds = 1.2; Rows = @($good) }
        @{ Name = 'ADUsers'; Scope = 'b.example.test'; Error = 'The server is not operational'; DurationSeconds = 21; Rows = @() }) }
$child = ConvertFrom-SmartM365AcceptanceChildOutput -StandardOutput ("WARNING: noise`r`n" + $script:ResultMarker + $json + "`r`n")
Assert-True ([bool]$child.Ready -and @($child.Datasets).Count -eq 2) 'Marker line is parsed among other output.'
$assessment = Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize 100
Assert-True ($assessment.Status -eq 'WARN' -and $assessment.Detail -match 'b\.example\.test: read failed') 'One failing domain out of two warns.'
$child.Datasets[0].Error = 'denied'
Assert-True ((Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize 100).Status -eq 'FAIL') 'Every domain failing fails.'
$missing = ConvertFrom-SmartM365AcceptanceChildOutput -StandardOutput ($script:ResultMarker + '{"Ready":false,"ComponentMissing":true,"Detail":"Adaxes PowerShell module is not installed on this server."}')
Assert-True ((Get-SmartM365AcceptanceChildAssessment -Child $missing -SampleSize 100).Status -eq 'SKIPPED') 'Missing component is skipped.'
Assert-True ((Get-SmartM365AcceptanceChildAssessment -Child $missing -SampleSize 100 -Requested).Status -eq 'FAIL') 'Missing component fails when requested.'
$broken = ConvertFrom-SmartM365AcceptanceChildOutput -StandardOutput '' -StandardError 'Import-Module : boom' -ExitCode 1
Assert-True (-not [bool]$broken.Ready -and $broken.Detail -match 'boom') 'No marker reports stderr.'
Assert-True ((Get-SmartM365AcceptanceChildAssessment -Child $broken -SampleSize 100).Status -eq 'FAIL') 'Child without result fails.'

# Modules.
$installed = @(
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = '2.25.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Users'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = '3.9.0' }
    [pscustomobject]@{ Name = 'ImportExcel'; Version = '7.8.10' }
    [pscustomobject]@{ Name = 'MicrosoftTeams'; Version = '7.1.0' }
    [pscustomobject]@{ Name = 'ActiveDirectory'; Version = '1.0.1.0' }
    [pscustomobject]@{ Name = 'Adaxes'; Version = '2025.1' }
)
Assert-True ((Get-SmartM365AcceptanceModuleAssessment -Installed $installed -Expected $script:ExpectedModules).Status -eq 'PASS') 'Expected modules pass.'
$oldExo = @($installed | Where-Object Name -ne 'ExchangeOnlineManagement') + [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = '3.4.0' }
$assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $oldExo -Expected $script:ExpectedModules
Assert-True ($assessment.Status -eq 'FAIL' -and $assessment.Detail -match 'ExchangeOnlineManagement 3\.4\.0 below 3\.7\.2') 'Old required module fails.'
$noAdaxes = @($installed | Where-Object Name -ne 'Adaxes')
Assert-True ((Get-SmartM365AcceptanceModuleAssessment -Installed $noAdaxes -Expected $script:ExpectedModules).Status -eq 'WARN') 'Missing optional module warns.'
$mixedGraph = @($installed) + [pscustomobject]@{ Name = 'Microsoft.Graph.Groups'; Version = '2.30.0' }
$assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $mixedGraph -Expected $script:ExpectedModules
Assert-True ($assessment.Status -eq 'WARN' -and $assessment.Detail -match 'Microsoft\.Graph\.Groups 2\.30\.0') 'Graph submodule at another version warns.'

# Graph roles and job eligibility.
$jobs = @(
    [pscustomobject]@{ Name = 'Domains'; Enabled = $true; RequiredCapabilities = @('Graph', 'SharedRuntime'); RequiredGraphAppRoles = 'Directory.Read.All' }
    [pscustomobject]@{ Name = 'Devices'; Enabled = $true; RequiredCapabilities = @('Graph', 'SharedRuntime'); RequiredGraphAppRoles = @('Device.Read.All', 'Directory.Read.All') }
    [pscustomobject]@{ Name = 'Exchange2016'; Enabled = $true; RequiredCapabilities = @('AD', 'ExchangeOnPrem', 'SharedRuntime'); RequiredGraphAppRoles = @() }
    [pscustomobject]@{ Name = 'Disabled'; Enabled = $false; RequiredCapabilities = @('Graph'); RequiredGraphAppRoles = @('Policy.Read.All') }
)
$required = @(Get-SmartM365AcceptanceRequiredGraphRole -Jobs $jobs -AdditionalRoles @('Mail.Send'))
Assert-True ((@($required.Role) -join ',') -eq 'Device.Read.All,Directory.Read.All,Mail.Send') 'Roles of enabled jobs plus configuration roles.'
Assert-True ((@($required | Where-Object Role -eq 'Directory.Read.All')[0].Jobs -join ',') -eq 'Domains,Devices') 'A role lists the jobs that need it.'
$eligibility = @(Get-SmartM365AcceptanceJobEligibility -Jobs $jobs -ReadyCapabilities @('Graph', 'SharedRuntime') -GraphRoles @('Directory.Read.All'))
Assert-True ($eligibility.Count -eq 3) 'Disabled jobs are ignored.'
Assert-True ([bool]@($eligibility | Where-Object Name -eq 'Domains')[0].Eligible) 'Job with every capability and role is eligible.'
$devices = @($eligibility | Where-Object Name -eq 'Devices')[0]
Assert-True (-not $devices.Eligible -and ($devices.MissingGraphAppRoles -join ',') -eq 'Device.Read.All') 'Missing role is reported.'
$exchange = @($eligibility | Where-Object Name -eq 'Exchange2016')[0]
Assert-True (-not $exchange.Eligible -and ($exchange.MissingCapabilities -join ',') -eq 'AD,ExchangeOnPrem') 'Missing capabilities are reported.'

# Tokens.
$tokenValues = @{ WorkspaceRootPath = 'D:\Data'; TenantKey = 'contoso-test'; DataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL' }
Assert-True ((Resolve-SmartM365AcceptanceTokens -Value '{{DataAllRootPath}}' -Tokens $tokenValues) -eq 'D:\Data\Data\Tenants\contoso-test\DATA-ALL') 'Nested tokens are resolved.'
$rejected = $false
try { $null = Resolve-SmartM365AcceptanceTokens -Value '{{Unknown}}\x' -Tokens $tokenValues } catch { $rejected = $true }
Assert-True $rejected 'Unknown token is refused.'

# The script must never export data.
$text = Get-Content -LiteralPath $scriptPath -Raw
Assert-True ($text -notmatch '(?i)\bExport-Csv\b|Export-SmartM365Csv|Publish-SmartM365Csv|Out-File') 'The acceptance script writes no data file.'
Assert-True ($text -match '\$global:EnableSharePointUpload = \$false') 'The SharePoint log upload is disabled.'

# Catalog entry used by Manual launches and the GUI buttons.
$catalog = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartInventory\Orchestrator\Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
$entry = @($catalog.Launches | Where-Object Name -eq 'Orchestrator-ServerAcceptance')
Assert-True ($entry.Count -eq 1 -and $entry[0].Manual -eq $true -and $entry[0].ScriptPath -eq 'Maintenance\SmartM365-OrchestratorServerAcceptance.ps1') 'Catalog manual entry points to the script.'
Assert-True (-not $entry[0].PSObject.Properties['HideInGui']) 'The entry is visible in the GUI Manual launches.'

Write-Host ("Orchestrator server acceptance offline test passed ({0} checks)." -f $script:Checks) -ForegroundColor Green
