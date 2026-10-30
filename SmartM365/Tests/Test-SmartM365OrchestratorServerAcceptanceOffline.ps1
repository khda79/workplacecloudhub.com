<#
.SYNOPSIS
    Offline checks of the orchestrator server acceptance script: dataset checks, statuses,
    module assessment, child output parsing, Graph roles and job eligibility. Nothing is read
    from a directory, a tenant or a share.
.VERSION
    1.1.0
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartInventory\Maintenance\SmartM365-OrchestratorServerAcceptance.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) { throw 'The acceptance script has PowerShell parse errors.' }
foreach ($definition in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and ($node.Name -like '*-SmartM365Acceptance*' -or $node.Name -eq 'Invoke-AcceptanceChildProcess') }, $true))) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
foreach ($assignment in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -in @('$script:GuidPattern', '$script:SmtpPattern', '$script:ResultMarker', '$script:ExpectedModules', '$script:ChildPreamble') }, $false))) {
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
foreach ($name in @('ADUsers', 'GraphUsers', 'EntraDevices', 'IntuneDevices', 'EXOMailboxes', 'OnPremMailboxes', 'SharePointSites', 'TeamsPhoneNumbers', 'AdaxesUsers')) {
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

# Steps announced by a child, and the step reported after a real time-out.
$stdout = "SMARTM365_ACCEPTANCE_STAGE:importing the Adaxes module (09:40:01)`r`nnoise`r`nSMARTM365_ACCEPTANCE_STAGE:Get-AdmUser through service adx1 on a.example.test (09:40:03)`r`n"
Assert-True ((Get-SmartM365AcceptanceLastStage -StandardOutput $stdout) -eq 'Get-AdmUser through service adx1 on a.example.test (09:40:03)') 'Last announced step is found.'
Assert-True ((Get-SmartM365AcceptanceLastStage -StandardOutput 'no stage') -eq '') 'No step gives an empty text.'
$hung = Invoke-AcceptanceChildProcess -EnginePath (Get-Process -Id $PID).Path -ScriptText "Write-AcceptanceStage 'waiting on a hung call'; Start-Sleep -Seconds 60" -TimeoutSeconds 8
Assert-True (-not [bool]$hung.Ready -and [bool]$hung.TimedOut -and $hung.Detail -match 'during: waiting on a hung call') "Time-out names the hung step ($($hung.Detail))."
$quick = Invoke-AcceptanceChildProcess -EnginePath (Get-Process -Id $PID).Path -ScriptText "Invoke-AcceptanceDataset 'X' '' { [pscustomobject]@{ A = 1 } } | Out-Null; Write-AcceptanceResult @{ Ready = `$true; Detail = 'ok'; Datasets = @(Invoke-AcceptanceDataset 'TeamsPhoneNumbers' '' { [pscustomobject]@{ TelephoneNumber = '+33100000000'; NumberType = 'CallingPlan'; ActivationState = 'Activated' } }) }" -TimeoutSeconds 60
Assert-True ([bool]$quick.Ready -and @($quick.Datasets[0].Rows).Count -eq 1) 'Step lines do not leak into the rows of a data set.'

# Inner exception messages reach the detail (Teams only said 'see inner exception' in 1.0.3).
$inner = Invoke-AcceptanceChildProcess -EnginePath (Get-Process -Id $PID).Path -ScriptText "try { throw [System.Net.Http.HttpRequestException]::new('The SSL connection could not be established, see inner exception.', [System.Security.Authentication.AuthenticationException]::new('The remote certificate is invalid')) } catch { Write-AcceptanceResult @{ Ready = `$false; Detail = (Get-AcceptanceErrorText `$_) } }" -TimeoutSeconds 60
Assert-True ($inner.Detail -match 'see inner exception\. -> The remote certificate is invalid') "Inner exception is reported ($($inner.Detail))."

# Adaxes child with a fake module: an unreadable objectClass is excluded by DN and the domain read again (like the collector).
$fakeRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365AcceptanceAdaxes-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path (Join-Path $fakeRoot 'Adaxes') -Force
Set-Content -LiteralPath (Join-Path $fakeRoot 'Adaxes\Adaxes.psm1') -Encoding utf8 -Value @'
function Get-AdmUser { param($AdaxesService, $Server, $LDAPFilter, $Properties, $ResultSetSize)
    if ($LDAPFilter -notlike '*svc\28x\29*') { throw "An instance of 'ADUser' can't be created because the attribute 'objectClass' of the directory object 'CN=svc(x),OU=Service,DC=example,DC=test' is unavailable." }
    1..3 | ForEach-Object { [pscustomobject]@{ ObjectGUID = [guid]::NewGuid(); SamAccountName = "user$_"; UserPrincipalName = "user$_@example.test" } }
}
'@
try {
    $adaxesScript = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$script:AdaxesChildScript' }, $false))[0].Right.Expression.Value
    $environment = @{ PSModulePath = $fakeRoot + [IO.Path]::PathSeparator + $env:PSModulePath; SMART_SAMPLE_SIZE = '100'; SMART_SMALL_SAMPLE_SIZE = '10'
        SMART_ADAXES_SERVICES = '["adx1"]'; SMART_ADAXES_DOMAINS = '["example.test"]'; SMART_ADAXES_LDAP_FILTER = '(sAMAccountType=805306368)' }
    $adaxes = Invoke-AcceptanceChildProcess -EnginePath (Get-Process -Id $PID).Path -ScriptText $adaxesScript -Environment $environment -TimeoutSeconds 120
    Assert-True ([bool]$adaxes.Ready -and @($adaxes.Datasets[0].Rows).Count -eq 3 -and [int]$adaxes.Info.ExcludedObjects -eq 1) "Unreadable object is excluded and the domain read again ($($adaxes.Detail) $(@($adaxes.Datasets)[0].Error))."
}
finally { Remove-Item -LiteralPath $fakeRoot -Recurse -Force -ErrorAction SilentlyContinue }

# Final summary: short results, tables and verdict.
$sample = @(
    [pscustomobject]@{ Capability = 'RunAccount'; Status = 'PASS'; DurationSeconds = 0; Detail = 'running as EXAMPLE\svc-test; elevated=True (account of the orchestrator scheduled task)' }
    [pscustomobject]@{ Capability = 'Modules'; Status = 'FAIL'; DurationSeconds = 11.7; Detail = 'PowerShell 7.6.6; Microsoft.Graph.Authentication 2.32.0; ImportExcel missing; Windows PowerShell 5.1 (ActiveDirectory=True)' }
    [pscustomobject]@{ Capability = 'Network'; Status = 'WARN'; DurationSeconds = 1.3; Detail = 'login.microsoftonline.com: HTTP 200 in 496 ms (direct); teams.microsoft.com (Teams): unreachable through direct: reset' }
    [pscustomobject]@{ Capability = 'SharedRuntime'; Status = 'PASS'; DurationSeconds = 1.4; Detail = 'DataAllRootPath: read/write OK; DATA witness write/read/metadata/delete: median 1.2 ms, p95 2.7 ms over 80 call(s); code share read/metadata/list: median 1.1 ms, p95 2 ms over 60 call(s); DATA free 39.2 % (49.5 GiB of 126.4 GiB)' }
    [pscustomobject]@{ Capability = 'Certificate'; Status = 'PASS'; DurationSeconds = 0.2; Detail = 'Usable certificate 9875...21D4 found in Cert:\CurrentUser\My; private key usable; expires 2027-10-15 19:18 UTC (350 day(s)). (tested as svc-test)' }
    [pscustomobject]@{ Capability = 'AD'; Status = 'WARN'; DurationSeconds = 16.7; Detail = 'forest example.test | ADUsers a.example.test: 100 row(s) in 0.5 s | ADUsers b.example.test: 100 row(s) in 0.3 s (UserPrincipalName has 2 value(s) in an unexpected format) | ADUsers c.example.test: read failed: denied' }
    [pscustomobject]@{ Capability = 'ExchangeOnPrem'; Status = 'SKIPPED'; DurationSeconds = 9; Detail = 'snap-in is not registered on this server.' }
)
$short = @{}; foreach ($item in $sample) { $short[$item.Capability] = Get-SmartM365AcceptanceShortResult -Capability $item.Capability -Status $item.Status -Detail $item.Detail }
Assert-True ($short.RunAccount -eq 'EXAMPLE\svc-test (orchestrator service account)') "RunAccount short result ($($short.RunAccount))."
Assert-True ($short.Modules -eq 'ImportExcel missing') "Modules short result lists the problems ($($short.Modules))."
Assert-True ($short.Network -eq '1 of 2 endpoint(s) reachable; unreachable: teams.microsoft.com') "Network short result ($($short.Network))."
Assert-True ($short.SharedRuntime -eq 'DATA SMB median 1.2 ms, code 1.1 ms, DATA free 39.2 %') "SharedRuntime short result ($($short.SharedRuntime))."
Assert-True ($short.Certificate -eq 'private key usable, expires 2027-10-15 (350 days)') "Certificate short result ($($short.Certificate))."
Assert-True ($short.AD -eq '2 of 3 data set(s) read, 200 row(s) (1 read failure(s), format warnings)') "Extraction short result ($($short.AD))."
Assert-True ($short.ExchangeOnPrem -eq 'component not installed on this server') 'Skipped short result.'
$table = @(Format-SmartM365AcceptanceTable -Title 'T' -Columns @('Status', 'Capability', 'Seconds') -Rows @([pscustomobject]@{ Status = 'WARN'; Capability = 'AD'; Seconds = [double]16.7 }))
Assert-True ($table.Count -eq 5 -and $table[2].Text -eq 'Status  Capability  Seconds' -and $table[4].Text -eq 'WARN    AD          16.7' -and $table[4].Status -eq 'WARN') "Single-row table is aligned ($(($table | ForEach-Object Text) -join ' / '))."
$table = @(Format-SmartM365AcceptanceTable -Title 'T' -Columns @('Item', 'Value') -Rows @([pscustomobject]@{ Item = 'a'; Value = 'x' * 200 }, [pscustomobject]@{ Item = 'bb'; Value = 'y' }) -MaxColumnWidth 20)
Assert-True ($table.Count -eq 6 -and $table[4].Text -eq ('a     ' + ('x' * 17) + '...') -and $table[5].Text -eq 'bb    y') "Long values are cut and columns aligned ($(($table | ForEach-Object Text) -join ' / '))."
Assert-True ((@(Format-SmartM365AcceptanceTable -Title 'T' -Columns @('A') -Rows @()) | Select-Object -Last 1).Text -eq '(none)') 'Empty table says none.'
Assert-True ((Get-SmartM365AcceptanceVerdict -Results $sample).Text -match '^NOT READY: fix Modules') 'A FAIL gives NOT READY.'
Assert-True ((Get-SmartM365AcceptanceVerdict -Results @($sample | Where-Object Status -ne 'FAIL')).Status -eq 'WARN') 'Warnings give READY WITH WARNINGS.'
Assert-True ((Get-SmartM365AcceptanceVerdict -Results @($sample | Where-Object Status -in @('PASS', 'SKIPPED'))).Status -eq 'PASS') 'PASS and SKIPPED give READY.'
$assessment = Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize 100
Assert-True (@($assessment.Extractions).Count -eq 2 -and @($assessment.Extractions | Where-Object Rows -eq '-').Count -eq 2) 'Failed scopes are listed in the extraction table.'

# Result published for the orchestrator console: one file per server, service account apart from other accounts.
$published = @(
    [pscustomobject]@{ Capability = 'Graph'; Status = 'PASS'; PublishedStatus = 'PASS'; Detail = 'Graph app-only sign-in' }
    [pscustomobject]@{ Capability = 'TeamsPowerShell'; Status = 'WARN'; PublishedStatus = 'FAIL'; Detail = 'non-blocking failure | blocked' }
    [pscustomobject]@{ Capability = 'RunAccount'; Status = 'PASS'; Detail = 'x' * 1500 }
)
$document = New-SmartM365AcceptanceResultDocument -Results $published -ServerName 'srv-test' -AcceptanceVersion '1.1.0' -RunAccount 'EXAMPLE\svc-test' -ServiceAccount $true -Python ([pscustomobject]@{ Status = 'PASS'; Detail = 'Python 3.12.1' }) -GeneratedAtUtc ([datetime]::new(2026, 10, 30, 10, 0, 0, [DateTimeKind]::Utc))
Assert-True ($document.SchemaVersion -eq 1 -and $document.ServerName -eq 'SRV-TEST' -and $document.GeneratedAtUtc -eq '2026-10-30T10:00:00.0000000Z' -and $document.ServiceAccount) 'Result document header.'
$names = @($document.Results | ForEach-Object { $_.Name })
Assert-True (($names -join ',') -eq 'Graph,TeamsPowerShell,RunAccount,Python') "Result entries ($($names -join ','))."
Assert-True (@($document.Results | Where-Object { $_.Name -eq 'TeamsPowerShell' })[0].Status -eq 'FAIL') 'A non-blocking failure is published as FAIL for the console.'
Assert-True (@($document.Results | Where-Object { $_.Name -eq 'RunAccount' })[0].Detail.Length -eq 1000) 'Long details are cut.'
$resultFolder = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365AcceptanceResult-' + [guid]::NewGuid().ToString('N'))
try {
    $path = Write-SmartM365AcceptanceResultFile -ServerFolder (Join-Path $resultFolder 'SRV-TEST') -Document $document
    Assert-True ([IO.Path]::GetFileName($path) -eq 'Orchestrator-Acceptance.json') 'Service account result file name.'
    $bytes = [IO.File]::ReadAllBytes($path)
    Assert-True (-not ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) 'Result file has no BOM.'
    $read = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    Assert-True ($read.ServerName -eq 'SRV-TEST' -and @($read.Results).Count -eq 4 -and [bool]$read.ServiceAccount) 'Result file reads back.'
    $document.ServiceAccount = $false
    $other = Write-SmartM365AcceptanceResultFile -ServerFolder (Join-Path $resultFolder 'SRV-TEST') -Document $document
    Assert-True ([IO.Path]::GetFileName($other) -eq 'Orchestrator-Acceptance.OtherAccount.json' -and [bool](Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).ServiceAccount) 'Another account never overwrites the service account result.'
    $null = Write-SmartM365AcceptanceResultFile -ServerFolder (Join-Path $resultFolder 'SRV-TEST') -Document $document
    Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $resultFolder 'SRV-TEST') -File).Count -eq 2) 'Rewriting leaves no temporary file.'
}
finally { Remove-Item -LiteralPath $resultFolder -Recurse -Force -ErrorAction SilentlyContinue }

# Modules.
$installed = @(
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = '2.25.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Users'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Groups'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Identity.DirectoryManagement'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Reports'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.DeviceManagement'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Beta.Groups'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Beta.DeviceManagement'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = '3.9.0' }
    [pscustomobject]@{ Name = 'ImportExcel'; Version = '7.8.10' }
    [pscustomobject]@{ Name = 'MicrosoftTeams'; Version = '7.1.0' }
    [pscustomobject]@{ Name = 'ActiveDirectory'; Version = '1.0.1.0' }
    [pscustomobject]@{ Name = 'Adaxes'; Version = '2025.1' }
)
Assert-True ((Get-SmartM365AcceptanceModuleAssessment -Installed $installed -Expected $script:ExpectedModules).Status -eq 'PASS') 'Expected modules pass.'
$oldTeams = @($installed | Where-Object Name -ne 'MicrosoftTeams') + [pscustomobject]@{ Name = 'MicrosoftTeams'; Version = '4.5.0' }
$assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $oldTeams -Expected $script:ExpectedModules
Assert-True ($assessment.Status -eq 'WARN' -and $assessment.Detail -match 'MicrosoftTeams 4\.5\.0 below 4\.7\.1') 'Optional module below its minimum warns.'
$authenticationOnly = @($installed | Where-Object { $_.Name -notlike 'Microsoft.Graph.*' -or $_.Name -eq 'Microsoft.Graph.Authentication' })
$assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $authenticationOnly -Expected $script:ExpectedModules
Assert-True ($assessment.Status -eq 'FAIL' -and $assessment.Detail -match 'Microsoft\.Graph\.Users missing') 'Graph submodules the jobs import are required.'
$production = @($installed | ForEach-Object { if ($_.Name -like 'Microsoft.Graph*') { [pscustomobject]@{ Name = $_.Name; Version = '2.32.0' } } else { $_ } })
Assert-True ((Get-SmartM365AcceptanceModuleAssessment -Installed $production -Expected $script:ExpectedModules).Status -eq 'PASS') 'Graph 2.32.0 used by the active servers passes.'
$noAdaxes = @($installed | Where-Object Name -ne 'Adaxes')
Assert-True ((Get-SmartM365AcceptanceModuleAssessment -Installed $noAdaxes -Expected $script:ExpectedModules).Status -eq 'WARN') 'Missing optional module warns.'
$mixedGraph = @($installed) + [pscustomobject]@{ Name = 'Microsoft.Graph.Sites'; Version = '2.30.0' }
$assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $mixedGraph -Expected $script:ExpectedModules
Assert-True ($assessment.Status -eq 'WARN' -and $assessment.Detail -match 'Microsoft.Graph.Sites 2.30.0') 'Graph submodule at another version warns.'

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

# Empty configuration lists stay arrays under strict mode (Adaxes Domains = [] gave a null Count in 1.0.1).
Assert-True (@(ConvertTo-SmartM365AcceptanceStringList @()).Count -eq 0 -and @(ConvertTo-SmartM365AcceptanceStringList 'a').Count -eq 1) 'String lists wrapped in @() always have a Count.'
Assert-True (-not ([regex]::IsMatch((Get-Content -LiteralPath $scriptPath -Raw), '=\s*ConvertTo-SmartM365AcceptanceStringList'))) 'Every assignment from a string list is wrapped in @().'

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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCjU3s5j6aD3pyW
# jCUi/e9HXCKvGlH8MnkxJkUlB1PLqqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAzglu7StEv1FvdLD9PKn8Dll4/72Ii1ZDrco7behKVcMA0GCSqG
# SIb3DQEBAQUABIIBgHtDs7MMj5V6n6XSDfgoYneSczS3RMxyS2ziJVfP97yaARop
# dpRqo3k8p5XdQyUTZYceK0r2aaKWKg/ZVyaVGzYgnv6KAcc9ZrDkJSsHAWKayxQv
# E5eRUrRMRTWmXADrcVZPjo+9IyGa5wQUVd8kMBixzjgh0e9bzrm1zX4YY1+O/1KL
# fye8tZclrt24GANYzlWB2jbPX6q+U8qlxnfaBiYuFBwe1FYiWOsLwoEJV0HKmqsc
# sQ6aXkGd4ZiSo044gzMJXfUrCfyIWeAM5LAqXntB8+3SU2MogyFDJgvBAEk8fuji
# RR2QoymQP70W0aVKyHGw6wrLrCoKhV0WU9b2BLAFhqELs9IbuBtKgP74+dztxsEE
# V0J5n1AxUFmiKCzgN7I/6C4faEBpXm6Uo73ev/64TVhhbyAD31kPgqWEoGMmoHIa
# 4gUKX2rwOP6QLd+ual6PxS+/qcdpbsggPn5+bS6pmsqaEWwVcDslSG51kxf6llLO
# YPZm9a8ramS1NK44X6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAxMDQ4
# MzBaMC8GCSqGSIb3DQEJBDEiBCB8zZVhXPJw880YWFZOjnWQoDaVjOuP/zpmcKeO
# ukXWQDANBgkqhkiG9w0BAQEFAASCAgAMYpnnA5dO6GirYhF5j/VIU7S3MXC8QSoW
# N6ZEfbnXmVyxjHAzPiK/KuqowHCMTLYmxMiCnBFBZXpkn+Inpi8+tm2Q6zXIC2Nr
# XOJj7JNBe9izRj0IQvoGWl5O2gFdmJMibhFyr9le1SmpGclDavcX7CdImcFRmR4M
# BTbpVrX14cKlQnla7tlxZ+XdMxqOLeTYOV41vOyAJRovfqJddXo8kropnfEdUllY
# p5Uy5fKwEbWbkutxbZ+0BwZ6Xdz/8RlVJ8wBeI8putPNtllU2kJqN5pxqO5S9VLa
# /hnvvl+2Rd0ze4lt+50HIEdV9gepMWD1r0dzfW6HoeJsEzv6la5254+BD6/hQ1jw
# vpwIKoZFJUcNGBeL0Ty9qHF8/gGbB6jyOHudMSo2QUfqNKfWnIkDW5GXC2mvA7y6
# bqwf+Dx5VKOphL+FJYKgp8+hn1TLiOevVOnVgFewMPlaE9+me2ZhKNhIfDacrbxF
# MSsXusmJyH75TAV9JyiC02NfyM6ddJvbdxguCYTIgWl6f/BSzfv2WBJNb7AKBiKL
# 9KOgrfH8NPjFXz52IY82UcXomdTdtDUtDJvBSmrxrS7PJv2svTqoCpLPa4+mEwRB
# BJCx1kOgGYy6OgdxCLaNw43Iv5mtGIECLtwyFVMjzUDorDnsrp53Pq6Jk4yV7Co4
# xKe+m/cimA==
# SIG # End signature block
