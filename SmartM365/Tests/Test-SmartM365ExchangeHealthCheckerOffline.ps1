<#
.SYNOPSIS
Offline tests of the Health Checker integration of the Exchange on-premises infrastructure collector.

.DESCRIPTION
Extracts the Health Checker functions from the collector without running it. Microsoft's Health
Checker is never executed: process runs use a synthetic stand-in script. Covers:
- reading a synthetic Health Checker XML exported like Health Checker does (Export-Clixml -Depth 2);
- masking and scrubbing of sensitive values, status mapping, vulnerability matching, summary facts;
- integrity check of the vendored file (valid, SHA256 mismatch, non-Microsoft signer);
- child process run: arguments (never -Server), completed, not elevated, timeout with process tree stop;
- per-run orchestration: local server in a process, other servers through a stubbed task start,
  result sources, scheduled fallback, StaleResult, TaskTimeout, integrity failure, not elevated,
  time budget, only the last run kept, unsafe work folder, check again just before the start;
- work folder: parents checked below Program Files, ProgramData or the drive root, Administrators
  owner and administrators-only ACL, same parent check as the installer;
- remote helpers: administrative share path, messages without domain, task logon failures,
  runner state mapping, run token progress, result file name check, copy when changed;
- local runner: constants and integrity check identical to the collector, no -Server, standalone,
  status.json content, previous output removed, Exchange Management Shell failure;
- installer: same task, folder and hash as the collector, classic account only, password handling,
  CIM over DCOM, task settings, explicit folder permissions;
- remote results: plain JSON lines only (size limit, strict schema, same findings as the XML),
  Import-Clixml used only for XML produced on the collector server, generic ACL rights refused;
- HealthCheckerExcludedServers: short name or FQDN, case-insensitive, no start, no CIM contact,
  NotRun without warning, not counted in ServersTotal;
- interrupted run: never SUCCESS (CANCELLED banner, warning line, exit code 130), no receipt;
- task that does not start: TaskTriggerFailed after the grace period with the hexadecimal result;
- installer preflight: local administrator membership (direct, through groups, unknown);
- module path: a Windows PowerShell child started with the PowerShell 7 module path still loads
  Microsoft.PowerShell.Security (Get-Acl, Get-AuthenticodeSignature) in the collector, installer and runner;
- activation config/defaults, explicit false/true switch precedence and disabled report section.

.VERSION
1.8.0
#>
[CmdletBinding()]
param([string]$SourceRoot)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$collectorFolder = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage'
$collectorPath = Join-Path $collectorFolder 'SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
$vendoredPath = Join-Path $collectorFolder 'ThirdParty/CSS-Exchange/HealthChecker/HealthChecker.ps1'
$repositoryRoot = Split-Path $SourceRoot -Parent
$nonMicrosoftSignedPath = Join-Path $repositoryRoot 'Sign-WorkplaceCloudHubRepositoryPowerShellScripts.ps1'

$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($collectorPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw "Collector parse failed: $parseErrors" }
$functionNames = @('Ensure-Directory', 'Format-HtmlValue', 'Get-HealthCheckerCsvColumn', 'Get-HealthCheckerMaskingRule', 'Get-HealthCheckerSetting', 'Resolve-SmartM365InfrastructureHealthCheckerEnabled',
    'Test-HealthCheckerScriptIntegrity', 'Test-HealthCheckerElevation', 'Invoke-HealthCheckerProcess', 'ConvertTo-HealthCheckerSafeValue',
    'Get-HealthCheckerLineValue', 'ConvertTo-HealthCheckerResultLine', 'ConvertFrom-HealthCheckerResultLine', 'ConvertFrom-HealthCheckerXml', 'Read-HealthCheckerResultJson', 'ConvertTo-HealthCheckerSummaryRow', 'Remove-HealthCheckerPreviousWorkFolder',
    'Get-HealthCheckerServerComputerName', 'Test-HealthCheckerServerExcluded', 'ConvertTo-HealthCheckerAdminSharePath', 'ConvertTo-HealthCheckerRemoteMessage', 'ConvertTo-HealthCheckerUtcTime',
    'Read-HealthCheckerRemoteStatus', 'Get-HealthCheckerTaskLogonFailure', 'ConvertFrom-HealthCheckerRunnerState', 'Resolve-HealthCheckerScheduledFallback',
    'Copy-HealthCheckerRemoteFile', 'Start-HealthCheckerRemoteRun', 'Update-HealthCheckerRemoteRun', 'Complete-HealthCheckerRemoteRun',
    'Get-HealthCheckerUnsafeAclEntry', 'Get-HealthCheckerWorkFolderAncestor', 'Set-HealthCheckerWorkFolderAclRule', 'Initialize-HealthCheckerWorkFolder',
    'Invoke-ExchangeHealthCheckerInventory', 'ConvertTo-ServersAndStorageHealthCheckerSectionHtml', 'ConvertTo-ServersAndStorageDatabaseSectionHtml',
    'New-HtmlExecutiveSummary')
$definitions = foreach ($name in $functionNames) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if ($null -eq $node) { throw "Missing function: $name" }
    $node.Extent.Text
}
$constants = @{}
foreach ($constantName in @('HealthCheckerVersion', 'HealthCheckerExpectedSha256', 'HealthCheckerRelativePath', 'HealthCheckerRunnerFileName', 'HealthCheckerRemoteRootPath', 'HealthCheckerTaskPath', 'HealthCheckerTaskName')) {
    $assignment = $ast.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq "`$$constantName" }, $true)
    if ($null -eq $assignment) { throw "Missing collector constant: $constantName" }
    $constants[$constantName] = $assignment.Right.Extent.Text.Trim("'", '"')
}

$module = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
& $module {
    param($Constants)
    Set-StrictMode -Version Latest
    $script:LogLines = New-Object 'System.Collections.Generic.List[string]'
    $script:Config = @{}
    $script:LocalMode = $null
    $script:HealthCheckerRemoteRootPath = $Constants.HealthCheckerRemoteRootPath
    $script:HealthCheckerTaskPath = $Constants.HealthCheckerTaskPath
    $script:HealthCheckerTaskName = $Constants.HealthCheckerTaskName
    function script:Write-Log { param($Message, $Level = 'INFO') $script:LogLines.Add(("[{0}] {1}" -f $Level, $Message)) | Out-Null }
    function script:Get-SmartM365EffectiveConfigValue { param($Name, $DefaultValue) if ($script:Config.ContainsKey($Name)) { return $script:Config[$Name] } return $DefaultValue }
} $constants

$script:Checks = 0
function Assert-HealthChecker {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-HealthChecker-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

function Get-SyntheticGroupingKey { param([string]$Name, [int]$Order) [pscustomobject]@{ Name = $Name; DisplayGroupName = $true; DisplayOrder = $Order; DefaultTabNumber = 1 } }
function Get-SyntheticLine {
    param([string]$Name, [AllowNull()][object]$Value, [string]$WriteType = 'Grey', [int]$Tab = 1, [string]$TestingName = '', [string]$CustomName = '', [AllowNull()][object]$CustomValue = $null)
    [pscustomobject]@{
        DisplayValue = $Value; Name = $Name; TestingName = $(if ($TestingName) { $TestingName } else { $Name }); CustomName = $(if ($CustomName) { $CustomName } else { $Name })
        TabNumber = $Tab; TestingValue = $Value; CustomValue = $(if ($null -ne $CustomValue) { $CustomValue } else { $Value }); OutColumns = $null; WriteType = $WriteType
    }
}

function Export-SyntheticHealthCheckerXml {
    param([string]$Path, [string]$ServerName)

    # Same object shapes as Health Checker's Add-AnalyzedResultInformation and Get-DisplayResultsGroupingKey.
    $displayResults = @{}
    $add = { param($Key, [object[]]$Lines) $list = New-Object System.Collections.Generic.List[object]; foreach ($l in $Lines) { $list.Add($l) }; $displayResults.Add($Key, $list) }
    & $add (Get-SyntheticGroupingKey 'Exchange Information' 1) @(
        (Get-SyntheticLine 'Name' $ServerName),
        (Get-SyntheticLine 'Generation Time' '10/21/2026 06:00:00'),
        (Get-SyntheticLine 'Version' 'Exchange 2016 CU23 Nov24SUv2'),
        (Get-SyntheticLine 'Build Number' '15.1.2507.44'),
        (Get-SyntheticLine 'Latest Install Time (SU/CU)' '11/27/2024 10:00:00'),
        (Get-SyntheticLine -Name 'Error' -Value 'Out of date Cumulative Update. Please upgrade to one of the two most recently released Cumulative Updates. Currently running on a build that is 400 days old.' -WriteType 'Red' -Tab 2 -TestingName 'Out of Date'),
        (Get-SyntheticLine -Name '' -Value 'Not on the latest SU. More Information: https://aka.ms/HC-ExBuilds' -WriteType 'Yellow' -Tab 2)
    )
    & $add (Get-SyntheticGroupingKey 'Hybrid Information' 2) @(
        (Get-SyntheticLine 'Domain(s) configured for Hybrid use' 'contoso.example')
    )
    & $add (Get-SyntheticGroupingKey 'NIC Settings Per Active Adapter' 3) @(
        (Get-SyntheticLine 'Interface Description' 'Synthetic adapter'),
        (Get-SyntheticLine -Name 'Address' -Value 'IPv4 192.0.2.10/24' -Tab 2),
        (Get-SyntheticLine -Name 'Sleepy NIC Disabled' -Value 'True' -WriteType 'Green' -Tab 2)
    )
    $tableLine = Get-SyntheticLine 'Database copies' 'ignored'
    $tableLine.WriteType = 'OutColumns'
    & $add (Get-SyntheticGroupingKey 'Security Settings' 4) @(
        (Get-SyntheticLine -Name 'Telemetry enabled' -Value 'Owner admin@contoso.example id 0f8fad5b-d9cb-469f-a165-70867728950e host 2001:db8::1 More Information: https://aka.ms/HC-Telemetry.' -WriteType 'Yellow'),
        (Get-SyntheticLine -Name 'SMB1 Installed' -Value 'False' -WriteType 'Green'),
        $tableLine
    )
    & $add (Get-SyntheticGroupingKey 'Security Vulnerability' 5) @(
        (Get-SyntheticLine -Name 'Security Vulnerability' -Value 'See: https://portal.msrc.microsoft.com/security-guidance/advisory/CVE-2024-21410 for more information.' -WriteType 'Red' -CustomName 'CVE-2024-21410' -CustomValue $true),
        (Get-SyntheticLine -Name 'Security Vulnerability' -Value 'CVE-2023-36745' -WriteType 'Red' -CustomName 'Security Vulnerability' -CustomValue 'CVE-2023-36745'),
        (Get-SyntheticLine -Name 'IIS module anomalies detected' -Value 'False' -CustomValue $false)
    )

    $result = [pscustomobject]@{
        HealthCheckerExchangeServer = [pscustomobject]@{ ServerName = $ServerName; ExchangeInformation = [pscustomobject]@{ BuildInformation = [pscustomobject]@{ VersionInformation = [pscustomobject]@{ BuildVersion = '15.1.2507.44' } } } }
        HtmlServerValues            = @{ ServerDetails = (New-Object System.Collections.Generic.List[object]) }
        DisplayResults              = $displayResults
    }
    $result | Export-Clixml -Path $Path -Encoding utf8 -Depth 2 -Force
}

try {
    # 1. Synthetic XML parsing.
    $fixture = Join-Path $tempRoot 'HealthChecker-EXSRV01-20261021060000.xml'
    Export-SyntheticHealthCheckerXml -Path $fixture -ServerName 'EXSRV01'
    $parsed = & $module { param($p, $v) ConvertFrom-HealthCheckerXml -Path $p -ServerName 'EXSRV01' -HealthCheckerVersion $v } $fixture $constants.HealthCheckerVersion
    $findings = @($parsed.Findings)
    Assert-HealthChecker ($findings.Count -eq 17) "Every result line becomes a finding (found $($findings.Count))."
    Assert-HealthChecker (@($findings | Where-Object { $_.ServerName -ne 'EXSRV01' }).Count -eq 0) 'Findings carry the server name.'
    $findingColumns = & $module { Get-HealthCheckerCsvColumn -Kind Findings }
    Assert-HealthChecker ((@($findings[0].PSObject.Properties.Name) -join ',') -eq ($findingColumns -join ',')) 'Finding rows follow the Findings CSV columns.'
    $outOfDate = @($findings | Where-Object { $_.Check -eq 'Out of Date' -and $_.Item -eq 'Error' })
    Assert-HealthChecker ($outOfDate.Count -eq 1 -and $outOfDate[0].Status -eq 'Error') 'An Error line with its own testing name is its own check, with Error status.'
    $suLine = @($findings | Where-Object { $_.Status -eq 'Warning' -and $_.Category -eq 'Exchange Information' })
    Assert-HealthChecker ($suLine.Count -eq 1 -and $suLine[0].Recommendation -eq 'https://aka.ms/HC-ExBuilds') 'Warning sub-line keeps its More Information link as recommendation.'
    $hybrid = @($findings | Where-Object { $_.Category -eq 'Hybrid Information' })
    Assert-HealthChecker ($hybrid.Count -eq 1 -and $hybrid[0].Value -eq '[masked]' -and $hybrid[0].ValueMasked -eq $true) 'Hybrid Information values are masked.'
    $address = @($findings | Where-Object { $_.Item -eq 'Address' })
    Assert-HealthChecker ($address.Count -eq 1 -and $address[0].Value -eq '[masked]') 'NIC address values are masked.'
    Assert-HealthChecker (@($findings | Where-Object { $_.Item -eq 'Sleepy NIC Disabled' -and $_.Status -eq 'OK' -and $_.Value -eq 'True' }).Count -eq 1) 'Green lines map to OK and keep their value.'
    $telemetry = @($findings | Where-Object { $_.Check -eq 'Telemetry enabled' })[0]
    Assert-HealthChecker ($telemetry.Value -notmatch 'contoso|0f8fad5b|2001:db8' -and $telemetry.Value -match '\[email\]' -and $telemetry.Value -match '\[guid\]' -and $telemetry.Value -match '\[ip\]') "Mail, GUID and IP values are scrubbed: $($telemetry.Value)"
    Assert-HealthChecker ($telemetry.Recommendation -eq 'https://aka.ms/HC-Telemetry') 'Trailing punctuation is removed from links.'
    Assert-HealthChecker (@($findings | Where-Object { $_.Check -eq 'Database copies' -and $_.Value -eq '[table]' }).Count -eq 1) 'Table lines are reported without their content.'
    $allText = ($findings | ForEach-Object { $_.Value }) -join ' '
    Assert-HealthChecker ($allText -notmatch '192\.0\.2\.10|contoso\.example') 'No raw address or domain value reaches the findings.'
    Assert-HealthChecker ((@($parsed.Vulnerabilities) -join ',') -eq 'CVE-2024-21410,CVE-2023-36745') "Vulnerability matching follows Health Checker (found $(@($parsed.Vulnerabilities) -join ','))."
    $facts = $parsed.Facts
    Assert-HealthChecker ($facts.ExchangeVersion -eq 'Exchange 2016 CU23 Nov24SUv2' -and $facts.BuildNumber -eq '15.1.2507.44') 'Version and build are read.'
    Assert-HealthChecker ($facts.CumulativeUpdate -eq 'CU23' -and $facts.SecurityUpdate -eq 'Nov24SUv2') 'CU and SU are extracted from the version.'
    Assert-HealthChecker ($facts.SecurityUpdateStatus -eq 'MissingLatestSU') 'Missing latest SU is detected.'
    Assert-HealthChecker ($facts.CumulativeUpdateOutOfDate -eq $true -and $facts.CumulativeUpdateDaysOld -eq 400) 'CU age is read from the out of date message.'
    Assert-HealthChecker ($facts.ErrorCount -eq 3 -and $facts.WarningCount -eq 2 -and $facts.OkCount -eq 2) "Status counts (E=$($facts.ErrorCount) W=$($facts.WarningCount) OK=$($facts.OkCount))."
    Assert-HealthChecker ($facts.GenerationTime -eq '10/21/2026 06:00:00' -and $findings[0].GenerationTime -eq '10/21/2026 06:00:00') 'Generation time is read and copied to findings.'
    Assert-HealthChecker ($facts.EndOfLifeStatus -eq 'Supported') 'No End Of Life line means supported.'

    # Depth-truncated and malformed exports.
    $truncated = Join-Path $tempRoot 'HealthChecker-TRUNC-1.xml'
    [pscustomobject]@{ DisplayResults = @{ 'Exchange Information' = @('plain text entry') } } | Export-Clixml -Path $truncated -Depth 2
    $truncatedParsed = & $module { param($p) ConvertFrom-HealthCheckerXml -Path $p -ServerName 'TRUNC' -HealthCheckerVersion 'x' } $truncated
    Assert-HealthChecker (@($truncatedParsed.Findings).Count -eq 1 -and $truncatedParsed.Facts.SecurityUpdateStatus -eq 'Unknown') 'A plain string entry is kept as an Info finding and the SU status is Unknown.'
    $empty = Join-Path $tempRoot 'HealthChecker-EMPTY-1.xml'
    [pscustomobject]@{ HtmlServerValues = @{} } | Export-Clixml -Path $empty -Depth 2
    $emptyFailed = $false
    try { & $module { param($p) ConvertFrom-HealthCheckerXml -Path $p -ServerName 'EMPTY' -HealthCheckerVersion 'x' } $empty } catch { $emptyFailed = $_.Exception.Message -match 'no DisplayResults' }
    Assert-HealthChecker $emptyFailed 'An XML without DisplayResults is refused.'

    # Summary row.
    $summaryRow = & $module { param($p) ConvertTo-HealthCheckerSummaryRow -ServerName 'EXSRV01' -RunStatus 'Completed' -RunMessage $null -DurationSeconds 42 -HealthCheckerVersion '26.10.13.1644' -Parsed $p } $parsed
    $summaryColumns = & $module { Get-HealthCheckerCsvColumn -Kind Summary }
    Assert-HealthChecker ((@($summaryRow.PSObject.Properties.Name) -join ',') -eq ($summaryColumns -join ',')) 'Summary rows follow the Summary CSV columns.'
    Assert-HealthChecker ($summaryRow.VulnerabilityCount -eq 2 -and $summaryRow.Vulnerabilities -eq 'CVE-2024-21410; CVE-2023-36745' -and $summaryRow.ErrorCount -eq 3) 'Summary row carries vulnerabilities and counts.'
    $notRunRow = & $module { ConvertTo-HealthCheckerSummaryRow -ServerName 'EXSRV02' -RunStatus 'Timeout' -RunMessage 'stopped' -HealthCheckerVersion 'v' -Parsed $null }
    Assert-HealthChecker ($null -eq $notRunRow.VulnerabilityCount -and $notRunRow.SecurityUpdateStatus -eq 'Unknown') 'A server without result has no measured counts.'

    # 2. Integrity of the vendored file.
    Assert-HealthChecker ((Get-FileHash -LiteralPath $vendoredPath -Algorithm SHA256).Hash -eq $constants.HealthCheckerExpectedSha256) 'The collector expected SHA256 matches the vendored file.'
    Assert-HealthChecker ($constants.HealthCheckerRelativePath -eq 'ThirdParty\CSS-Exchange\HealthChecker\HealthChecker.ps1') 'The collector points at the vendored file.'
    $valid = & $module { param($p, $h) Test-HealthCheckerScriptIntegrity -Path $p -ExpectedSha256 $h } $vendoredPath $constants.HealthCheckerExpectedSha256
    Assert-HealthChecker ($valid.IsValid -and $valid.SignatureStatus -eq 'Valid' -and $valid.Signer -match 'CN=Microsoft Corporation') "The vendored file passes the integrity check ($($valid.Reason))."
    $tampered = Join-Path $tempRoot 'Tampered.ps1'
    $bytes = [IO.File]::ReadAllBytes($vendoredPath); $bytes[2000] = if ($bytes[2000] -eq 0x41) { [byte]0x42 } else { [byte]0x41 }; [IO.File]::WriteAllBytes($tampered, $bytes)
    $tamperedResult = & $module { param($p, $h) Test-HealthCheckerScriptIntegrity -Path $p -ExpectedSha256 $h } $tampered $constants.HealthCheckerExpectedSha256
    Assert-HealthChecker (-not $tamperedResult.IsValid -and $tamperedResult.Reason -match 'SHA256 mismatch') 'A modified file fails on SHA256.'
    $tamperedHash = (Get-FileHash -LiteralPath $tampered -Algorithm SHA256).Hash
    $tamperedSigned = & $module { param($p, $h) Test-HealthCheckerScriptIntegrity -Path $p -ExpectedSha256 $h } $tampered $tamperedHash
    Assert-HealthChecker (-not $tamperedSigned.IsValid -and $tamperedSigned.Reason -match 'not valid') 'A modified file with a matching expected hash still fails on its signature.'
    $otherSignerHash = (Get-FileHash -LiteralPath $nonMicrosoftSignedPath -Algorithm SHA256).Hash
    $otherSigner = & $module { param($p, $h) Test-HealthCheckerScriptIntegrity -Path $p -ExpectedSha256 $h } $nonMicrosoftSignedPath $otherSignerHash
    Assert-HealthChecker (-not $otherSigner.IsValid -and $otherSigner.Reason -match 'not Microsoft|not valid') "A file signed by another publisher is refused ($($otherSigner.Reason))."
    $missing = & $module { param($p) Test-HealthCheckerScriptIntegrity -Path $p -ExpectedSha256 'AA' } (Join-Path $tempRoot 'missing.ps1')
    Assert-HealthChecker (-not $missing.IsValid -and $missing.Reason -match 'not found') 'A missing file is refused.'

    # 3. Settings.
    $settings = & $module { $script:Config = @{}; Get-HealthCheckerSetting }
    Assert-HealthChecker ($settings.TimeoutMinutesPerServer -eq 30 -and $settings.TotalBudgetMinutes -eq 1320 -and $settings.MaxResultAgeHours -eq 26 -and $settings.WorkFolderPath -eq (Join-Path $env:ProgramFiles 'SmartM365\HealthChecker\Work')) 'Default settings: the work folder is under Program Files.'
    Assert-HealthChecker ($settings.RemoteRootPath -eq 'C:\Program Files\SmartM365\HealthChecker' -and $settings.TaskPath -eq '\WCH\' -and $settings.TaskName -eq 'SmartM365 Exchange Health Checker' -and $settings.PollSeconds -eq 15) 'Remote task settings come from the collector constants.'
    Assert-HealthChecker (-not $settings.PSObject.Properties['KeepFailedRawOutputDays']) 'The failed raw output retention setting is gone.'
    & $module { $script:LogLines.Clear(); $script:Config = @{ HealthCheckerKeepFailedRawOutputDays = 7 } }
    [void](& $module { Get-HealthCheckerSetting })
    $obsoleteLog = (& $module { $script:LogLines }) -join "`n"
    Assert-HealthChecker ($obsoleteLog -match '\[INFO\] HealthCheckerKeepFailedRawOutputDays is no longer used') 'A former retention setting is ignored with an information line, not a warning.'
    $ageRefused = $false
    try { & $module { $script:Config = @{ HealthCheckerMaxResultAgeHours = 0 }; Get-HealthCheckerSetting } } catch { $ageRefused = $true }
    Assert-HealthChecker $ageRefused 'An out-of-range result age limit is refused.'
    & $module { $script:LogLines.Clear() }
    $uncRefused = $false
    try { & $module { $script:Config = @{ HealthCheckerWorkFolderPath = '\\server\share\hc' }; Get-HealthCheckerSetting } } catch { $uncRefused = $_.Exception.Message -match 'local absolute path' }
    Assert-HealthChecker $uncRefused 'A share as work folder is refused.'
    $rangeRefused = $false
    try { & $module { $script:Config = @{ HealthCheckerTimeoutMinutesPerServer = 0 }; Get-HealthCheckerSetting } } catch { $rangeRefused = $true }
    Assert-HealthChecker $rangeRefused 'An out-of-range per-server limit is refused.'
    & $module { $script:Config = @{} }

    # 4. Child process with a synthetic stand-in (Microsoft's script is never run here).
    $fakeScript = Join-Path $tempRoot 'FakeHealthChecker.ps1'
    [IO.File]::WriteAllText($fakeScript, @'
param([string[]]$Server, [string]$OutputFilePath, [switch]$SkipVersionCheck)
$mode = $env:SMARTM365_FAKE_HC_MODE
Set-Content -LiteralPath (Join-Path $OutputFilePath 'fake-args.txt') -Value ("Server={0};SkipVersionCheck={1};PID={2}" -f ($Server -join ','), $SkipVersionCheck.IsPresent, $PID)
if ($mode -eq 'sleep') { Start-Sleep -Seconds 120; exit 0 }
if ($mode -eq 'notadmin') { Write-Warning 'The script needs to be executed in elevated mode. Start the Exchange Management Shell as an Administrator.'; exit }
if ($mode -eq 'nooutput') { exit 1 }
if ($mode -eq 'noshell') { Write-Warning 'Failed to Load Exchange PowerShell Module...'; exit 1 }
$name = if ($Server) { $Server[0].ToUpper() } else { $env:COMPUTERNAME.ToUpper() }
Copy-Item -LiteralPath $env:SMARTM365_FAKE_HC_FIXTURE -Destination (Join-Path $OutputFilePath ("HealthChecker-{0}-20261021060000.xml" -f $name))
'@, [Text.UTF8Encoding]::new($false))
    $env:SMARTM365_FAKE_HC_FIXTURE = $fixture
    try {
        $env:SMARTM365_FAKE_HC_MODE = 'ok'
        $okFolder = Join-Path $tempRoot 'run-ok'
        $ok = & $module { param($s, $o) Invoke-HealthCheckerProcess -ScriptPath $s -ServerName 'EXSRV01' -OutputFolder $o -TimeoutSeconds 120 } $fakeScript $okFolder
        Assert-HealthChecker ($ok.Status -eq 'Completed' -and $ok.XmlPath -like '*HealthChecker-*.xml') "A completed run returns its XML ($($ok.Status) $($ok.Message))."
        $fakeArgs = Get-Content -LiteralPath (Join-Path $okFolder 'fake-args.txt') -Raw
        Assert-HealthChecker ($fakeArgs -match 'Server=;SkipVersionCheck=True') "Health Checker gets -SkipVersionCheck and never -Server ($fakeArgs)."

        # Checked again just before the start: a file swapped after the first check never runs.
        $swappedFolder = Join-Path $tempRoot 'run-swapped'
        $swapped = & $module { param($s, $o, $h) Invoke-HealthCheckerProcess -ScriptPath $s -ServerName 'EXSRV01' -OutputFolder $o -TimeoutSeconds 120 -ExpectedSha256 $h } $fakeScript $swappedFolder $constants.HealthCheckerExpectedSha256
        Assert-HealthChecker ($swapped.Status -eq 'IntegrityFailed' -and $swapped.Message -match 'SHA256 mismatch' -and -not (Test-Path -LiteralPath (Join-Path $swappedFolder 'fake-args.txt'))) 'A file that no longer matches the expected SHA256 is not started.'

        $env:SMARTM365_FAKE_HC_MODE = 'notadmin'
        $notAdmin = & $module { param($s, $o) Invoke-HealthCheckerProcess -ScriptPath $s -ServerName 'EXSRV01' -OutputFolder $o -TimeoutSeconds 120 } $fakeScript (Join-Path $tempRoot 'run-notadmin')
        Assert-HealthChecker ($notAdmin.Status -eq 'NotElevated') "The elevation warning is recognized ($($notAdmin.Status))."

        $env:SMARTM365_FAKE_HC_MODE = 'nooutput'
        $noOutput = & $module { param($s, $o) Invoke-HealthCheckerProcess -ScriptPath $s -ServerName 'EXSRV01' -OutputFolder $o -TimeoutSeconds 120 } $fakeScript (Join-Path $tempRoot 'run-nooutput')
        Assert-HealthChecker ($noOutput.Status -eq 'NoOutput' -and $noOutput.ExitCode -eq 1) "A run without XML is reported with its exit code ($($noOutput.Status) $($noOutput.ExitCode))."

        $env:SMARTM365_FAKE_HC_MODE = 'sleep'
        $sleepFolder = Join-Path $tempRoot 'run-sleep'
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $timeout = & $module { param($s, $o) Invoke-HealthCheckerProcess -ScriptPath $s -ServerName 'EXSRV01' -OutputFolder $o -TimeoutSeconds 5 } $fakeScript $sleepFolder
        $watch.Stop()
        Assert-HealthChecker ($timeout.Status -eq 'Timeout' -and $watch.Elapsed.TotalSeconds -lt 60) "A hung run is stopped at its limit ($($timeout.Status), $([int]$watch.Elapsed.TotalSeconds)s)."
        $sleepPid = [int]([regex]::Match((Get-Content -LiteralPath (Join-Path $sleepFolder 'fake-args.txt') -Raw), 'PID=(\d+)').Groups[1].Value)
        Assert-HealthChecker ($null -eq (Get-Process -Id $sleepPid -ErrorAction SilentlyContinue)) 'The stopped run leaves no process behind.'
    }
    finally {
        Remove-Item Env:\SMARTM365_FAKE_HC_MODE -ErrorAction SilentlyContinue
        Remove-Item Env:\SMARTM365_FAKE_HC_FIXTURE -ErrorAction SilentlyContinue
    }

    # 5. Orchestration: the local server runs in a stubbed process, remote servers through a stubbed start.
    $servers = @([pscustomobject]@{ Name = 'EXSRV01' }, [pscustomobject]@{ Name = 'EXSRV02'; Fqdn = 'EXSRV02.corp.contoso.example' })
    $remoteShare = Join-Path $tempRoot 'remote-EXSRV02'
    New-Item -ItemType Directory -Path (Join-Path $remoteShare 'Raw') -Force | Out-Null
    $runnerSource = Join-Path $collectorFolder $constants.HealthCheckerRunnerFileName
    & $module {
        param($FixturePath, $RemoteShare)
        $script:FixturePath = $FixturePath
        $script:ResultJson = [pscustomobject]@{ SchemaVersion = 1; HealthCheckerVersion = 'x'; Lines = @(ConvertTo-HealthCheckerResultLine -Path $FixturePath) } | ConvertTo-Json -Depth 4 -Compress
        $script:RemoteShare = $RemoteShare
        $script:ProcessCalls = New-Object 'System.Collections.Generic.List[object]'
        $script:RemoteStarts = New-Object 'System.Collections.Generic.List[object]'
        $script:Elevated = $true
        $script:RemoteMode = 'Completed'
        $script:UnsafeWorkFolder = $false
        $script:WorkFolderInitializations = New-Object 'System.Collections.Generic.List[string]'
        function script:Test-HealthCheckerElevation { return $script:Elevated }
        # Setting the Administrators owner needs elevation: the folder preparation itself is tested separately.
        function script:Initialize-HealthCheckerWorkFolder {
            param($WorkRoot)
            $script:WorkFolderInitializations.Add($WorkRoot) | Out-Null
            if ($script:UnsafeWorkFolder) { throw 'A parent of the Health Checker work folder is not restricted to administrators (S-1-5-32-545 Modify): C:\Unsafe' }
            New-Item -ItemType Directory -Path $WorkRoot -Force | Out-Null
        }
        function script:Invoke-HealthCheckerProcess {
            param($ScriptPath, $ServerName, $OutputFolder, $TimeoutSeconds, $ExpectedSha256)
            $script:ProcessCalls.Add([pscustomobject]@{ ScriptPath = $ScriptPath; ServerName = $ServerName; TimeoutSeconds = $TimeoutSeconds; ExpectedSha256 = $ExpectedSha256; ScriptExists = (Test-Path -LiteralPath $ScriptPath) }) | Out-Null
            Ensure-Directory -Path $OutputFolder
            if ($script:LocalMode -eq 'Timeout') { return [pscustomobject]@{ ServerName = $ServerName; Status = 'Timeout'; Message = 'stopped'; ExitCode = $null; DurationSeconds = 5; XmlPath = $null } }
            $xml = Join-Path $OutputFolder ("HealthChecker-{0}-1.xml" -f $ServerName)
            Copy-Item -LiteralPath $script:FixturePath -Destination $xml
            return [pscustomobject]@{ ServerName = $ServerName; Status = 'Completed'; Message = $null; ExitCode = 0; DurationSeconds = 12; XmlPath = $xml }
        }
        function script:Start-HealthCheckerRemoteRun {
            param($ServerName, $Computer, $VerifiedScriptPath, $RunnerSourcePath, $Settings)
            $script:RemoteStarts.Add([pscustomobject]@{ ServerName = $ServerName; Computer = $Computer; VerifiedScriptPath = $VerifiedScriptPath; RunnerSourcePath = $RunnerSourcePath }) | Out-Null
            $now = [DateTime]::UtcNow
            $previous = [pscustomobject]@{ RunToken = 'old'; State = 'Completed'; Message = $null; StartedAtUtc = $now.AddHours(-3).ToString('o'); FinishedAtUtc = $now.AddHours(-2).ToString('o'); DurationSeconds = 150; ResultFileName = 'result.json' }
            $run = [pscustomobject]@{ ServerName = $ServerName; Computer = $Computer; Session = $null; OutputSharePath = $script:RemoteShare; PreviousStatus = $previous; WaitToken = $null; StartedAt = (Get-Date); Done = $false; Status = $null; Message = $null; RemoteStatus = $null }
            [IO.File]::WriteAllText((Join-Path $script:RemoteShare 'result.json'), $script:ResultJson)
            switch ($script:RemoteMode) {
                'Completed' {
                    $current = [pscustomobject]@{ RunToken = 'new'; State = 'Completed'; Message = $null; StartedAtUtc = $now.ToString('o'); FinishedAtUtc = $now.ToString('o'); DurationSeconds = 140; ResultFileName = 'result.json' }
                    [IO.File]::WriteAllText((Join-Path $script:RemoteShare 'status.json'), ($current | ConvertTo-Json))
                }
                'RunnerTimeout' {
                    $current = [pscustomobject]@{ RunToken = 'new'; State = 'Timeout'; Message = 'Health Checker exceeded the 30 minute limit on the server and was stopped.'; StartedAtUtc = $now.ToString('o'); FinishedAtUtc = $now.ToString('o'); DurationSeconds = 1800; ResultFileName = $null }
                    [IO.File]::WriteAllText((Join-Path $script:RemoteShare 'status.json'), ($current | ConvertTo-Json))
                }
                'TriggerFailed' { $run.Done = $true; $run.Status = 'TaskTriggerFailed'; $run.Message = 'The Health Checker task could not be started: access denied on EXSRV02.corp.contoso.example.' }
                'TriggerFailedStale' { $previous.FinishedAtUtc = $now.AddHours(-30).ToString('o'); $run.Done = $true; $run.Status = 'TaskTriggerFailed'; $run.Message = 'start refused' }
                'NoResult' { [IO.File]::WriteAllText((Join-Path $script:RemoteShare 'status.json'), ($previous | ConvertTo-Json)) }
            }
            return $run
        }
    } $fixture $remoteShare
    $workRoot = Join-Path $tempRoot 'work'
    $runSettings = [pscustomobject]@{ TimeoutMinutesPerServer = 30; TotalBudgetMinutes = 1320; MaxResultAgeHours = 26; WorkFolderPath = $workRoot; RemoteRootPath = 'C:\Program Files\SmartM365\HealthChecker'; TaskPath = '\WCH\'; TaskName = 'SmartM365 Exchange Health Checker'; PollSeconds = 0 }
    $invoke = {
        param($Servers, $Source, $Runner, $Hash, $Version, $Settings)
        Invoke-ExchangeHealthCheckerInventory -ExchangeServers $Servers -SourceScriptPath $Source -RunnerSourcePath $Runner -ExpectedSha256 $Hash -HealthCheckerVersion $Version -RunId '20261021-060000' -Settings $Settings -LocalComputerName 'EXSRV01'
    }
    $result = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $runSettings
    $calls = @(& $module { $script:ProcessCalls })
    $starts = @(& $module { $script:RemoteStarts })
    Assert-HealthChecker ($calls.Count -eq 1 -and $calls[0].ServerName -eq 'EXSRV01' -and $calls[0].ScriptPath -like "$workRoot*" -and $calls[0].ScriptExists) 'Only the collector server runs in a local process, from the verified copy.'
    Assert-HealthChecker ($calls[0].TimeoutSeconds -eq 1800) 'The per-server limit is applied.'
    Assert-HealthChecker ($calls[0].ExpectedSha256 -eq $constants.HealthCheckerExpectedSha256) 'The local run checks the copy again just before starting it.'
    Assert-HealthChecker (@(& $module { $script:WorkFolderInitializations }) -contains $workRoot) 'The work folder is restricted to administrators before the copy is written.'
    Assert-HealthChecker ($starts.Count -eq 1 -and $starts[0].ServerName -eq 'EXSRV02' -and $starts[0].Computer -eq 'EXSRV02.corp.contoso.example' -and $starts[0].VerifiedScriptPath -like "$workRoot*" -and $starts[0].RunnerSourcePath -eq $runnerSource) 'Other servers start their task by FQDN, with the verified copy and the runner.'
    $summary = @($result.Summary)
    Assert-HealthChecker ($summary.Count -eq 2 -and $summary[0].RunStatus -eq 'Completed' -and $summary[1].RunStatus -eq 'Completed') 'Both servers complete.'
    Assert-HealthChecker ($summary[0].ResultSource -eq 'Local' -and $summary[0].ResultAgeHours -eq 0 -and $summary[1].ResultSource -eq 'OnDemand' -and $summary[1].DurationSeconds -eq 140) 'Result sources and remote duration are recorded.'
    Assert-HealthChecker ($summary[0].ServersProcessed -eq 2 -and $summary[0].ServersTotal -eq 2) 'Processed and total server counts are on every row.'
    Assert-HealthChecker (@($result.Findings).Count -eq 34) 'Findings of both servers are returned.'
    $runFolder = Join-Path $workRoot '20261021-060000'
    Assert-HealthChecker (-not (Test-Path -LiteralPath (Join-Path $runFolder 'EXSRV01')) -and -not (Test-Path -LiteralPath (Join-Path $runFolder 'EXSRV02'))) 'Raw output of completed servers is deleted.'
    Assert-HealthChecker (-not (Test-Path -LiteralPath (Join-Path $runFolder 'HealthChecker.ps1'))) 'The run copy of the script is removed.'
    Assert-HealthChecker (Test-Path -LiteralPath (Join-Path $remoteShare 'result.json')) 'Files on the server are left to its own next run.'

    & $module { $script:ProcessCalls.Clear(); $script:LogLines.Clear(); $script:RemoteMode = 'RunnerTimeout'; $script:LocalMode = 'Timeout' }
    $failed = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $runSettings
    $failedSummary = @($failed.Summary)
    Assert-HealthChecker ($failedSummary[0].RunStatus -eq 'Timeout' -and $failedSummary[1].RunStatus -eq 'TaskTimeout' -and $failedSummary[1].RunMessage -match 'on the server') 'A local timeout is Timeout and a runner timeout is TaskTimeout.'
    Assert-HealthChecker ($null -eq $failedSummary[1].ResultSource -and $failedSummary[0].ServersProcessed -eq 0) 'A failed server has no result source.'
    Assert-HealthChecker (Test-Path -LiteralPath (Join-Path $runFolder 'EXSRV01')) 'Raw output of a failed local server is kept until the next run.'
    $logText = (& $module { $script:LogLines }) -join "`n"
    Assert-HealthChecker ($logText -match '\[WARN\] EXSRV02: Health Checker TaskTimeout' -and $logText -match 'until the next run' -and $logText -match 'processed 0 of 2') 'Failures are logged as warnings with the processed count.'

    & $module { $script:RemoteMode = 'TriggerFailed'; $script:LocalMode = $null }
    $fallback = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $runSettings
    $fallbackRow = @($fallback.Summary)[1]
    Assert-HealthChecker ($fallbackRow.RunStatus -eq 'Completed' -and $fallbackRow.ResultSource -eq 'Scheduled' -and $fallbackRow.ResultAgeHours -ge 1.9 -and $fallbackRow.ResultAgeHours -le 2.1) "A failed start uses a recent scheduled result ($($fallbackRow.RunStatus) $($fallbackRow.ResultSource) $($fallbackRow.ResultAgeHours))."
    Assert-HealthChecker ($fallbackRow.RunMessage -match 'TaskTriggerFailed' -and $fallbackRow.RunMessage -match 'EXSRV02' -and $fallbackRow.RunMessage -notmatch 'contoso') "The fallback message keeps the cause without the domain: $($fallbackRow.RunMessage)"
    Assert-HealthChecker ($fallbackRow.DurationSeconds -eq 150 -and @($fallback.Findings | Where-Object { $_.ServerName -eq 'EXSRV02' }).Count -eq 17) 'The scheduled result is read.'

    & $module { $script:RemoteMode = 'TriggerFailedStale' }
    $stale = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $runSettings
    $staleRow = @($stale.Summary)[1]
    Assert-HealthChecker ($staleRow.RunStatus -eq 'StaleResult' -and $staleRow.RunMessage -match '30 h old \(limit 26 h\)' -and $null -eq $staleRow.ResultSource -and @($stale.Findings | Where-Object { $_.ServerName -eq 'EXSRV02' }).Count -eq 0) "An old scheduled result is StaleResult and is not read ($($staleRow.RunMessage))."

    & $module { $script:RemoteMode = 'NoResult' }
    $shortSettings = [pscustomobject]@{ TimeoutMinutesPerServer = 0.05; TotalBudgetMinutes = 1320; MaxResultAgeHours = 26; WorkFolderPath = $workRoot; RemoteRootPath = 'C:\Program Files\SmartM365\HealthChecker'; TaskPath = '\WCH\'; TaskName = 'SmartM365 Exchange Health Checker'; PollSeconds = 1 }
    $noResult = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $shortSettings
    Assert-HealthChecker (@($noResult.Summary)[1].RunStatus -eq 'TaskTimeout' -and @($noResult.Summary)[1].RunMessage -match 'No result from the Health Checker task') 'No new run within the limit is TaskTimeout, never the previous result.'
    & $module { $script:RemoteMode = 'Completed' }

    & $module { $script:ProcessCalls.Clear(); $script:RemoteStarts.Clear(); $script:LogLines.Clear() }
    $integrityResult = & $module $invoke $servers $vendoredPath $runnerSource ('0' * 64) $constants.HealthCheckerVersion $runSettings
    Assert-HealthChecker (@(& $module { $script:ProcessCalls }).Count -eq 0 -and @(& $module { $script:RemoteStarts }).Count -eq 0) 'Nothing runs or starts when the integrity check fails.'
    Assert-HealthChecker (@($integrityResult.Summary | Where-Object { $_.RunStatus -eq 'IntegrityFailed' -and $_.RunMessage -match 'SHA256 mismatch' }).Count -eq 2 -and @($integrityResult.Findings).Count -eq 0) 'Every server reports IntegrityFailed.'
    Assert-HealthChecker (((& $module { $script:LogLines }) -join "`n") -match '\[WARN\] Health Checker was not run') 'The integrity failure is a warning.'

    $missingRunner = & $module $invoke $servers $vendoredPath (Join-Path $tempRoot 'no-runner.ps1') $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $runSettings
    Assert-HealthChecker (@(& $module { $script:RemoteStarts }).Count -eq 0 -and @($missingRunner.Summary | Where-Object { $_.RunStatus -eq 'IntegrityFailed' -and $_.RunMessage -match 'runner' }).Count -eq 2) 'A missing runner blocks the run.'

    & $module { $script:ProcessCalls.Clear(); $script:RemoteStarts.Clear(); $script:LogLines.Clear(); $script:UnsafeWorkFolder = $true }
    $unsafeWorkRoot = Join-Path $tempRoot 'work-unsafe'
    $unsafeSettings = [pscustomobject]@{ TimeoutMinutesPerServer = 30; TotalBudgetMinutes = 1320; MaxResultAgeHours = 26; WorkFolderPath = $unsafeWorkRoot; RemoteRootPath = 'C:\Program Files\SmartM365\HealthChecker'; TaskPath = '\WCH\'; TaskName = 'SmartM365 Exchange Health Checker'; PollSeconds = 0 }
    $unsafeResult = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $unsafeSettings
    Assert-HealthChecker (@(& $module { $script:ProcessCalls }).Count -eq 0 -and @(& $module { $script:RemoteStarts }).Count -eq 0 -and -not (Test-Path -LiteralPath (Join-Path $unsafeWorkRoot '20261021-060000'))) 'Nothing is copied, run or started when the work folder cannot be secured.'
    Assert-HealthChecker (@($unsafeResult.Summary | Where-Object { $_.RunStatus -eq 'UnsafeWorkFolder' -and $_.RunMessage -match 'not restricted to administrators' }).Count -eq 2) 'Every server reports UnsafeWorkFolder.'
    Assert-HealthChecker (((& $module { $script:LogLines }) -join "`n") -match '\[WARN\] Health Checker was not run: its work folder') 'An unsafe work folder is a warning.'
    & $module { $script:UnsafeWorkFolder = $false }

    & $module { $script:Elevated = $false }
    $notElevatedResult = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $runSettings
    Assert-HealthChecker (@(& $module { $script:ProcessCalls }).Count -eq 0 -and @($notElevatedResult.Summary | Where-Object { $_.RunStatus -eq 'NotElevated' }).Count -eq 2) 'Nothing runs without elevation and every server reports NotElevated.'
    & $module { $script:Elevated = $true }

    $localOnly = @([pscustomobject]@{ Name = 'EXSRV01' })
    $budgetSettings = [pscustomobject]@{ TimeoutMinutesPerServer = 30; TotalBudgetMinutes = 1.04; MaxResultAgeHours = 26; WorkFolderPath = (Join-Path $tempRoot 'work-budget'); RemoteRootPath = 'C:\Program Files\SmartM365\HealthChecker'; TaskPath = '\WCH\'; TaskName = 'SmartM365 Exchange Health Checker'; PollSeconds = 0 }
    [void](& $module $invoke $localOnly $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $budgetSettings)
    $budgetCalls = @(& $module { $script:ProcessCalls })
    Assert-HealthChecker ($budgetCalls.Count -eq 1 -and $budgetCalls[0].TimeoutSeconds -le 63) "The limit never exceeds the remaining budget ($($budgetCalls[0].TimeoutSeconds)s)."
    $budgetSettings.TotalBudgetMinutes = 0.5
    $budgetResult = & $module $invoke $localOnly $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $budgetSettings
    Assert-HealthChecker (@($budgetResult.Summary)[0].RunStatus -eq 'NotRun' -and @($budgetResult.Summary)[0].RunMessage -match 'budget') 'A server left when the budget is used up is NotRun.'

    # Only the last run is kept: earlier run folders are removed, other folders are left.
    $purgeRoot = Join-Path $tempRoot 'purge'
    foreach ($name in @('20261001-010101', '20261020-010101', '20261021-060000', 'Keep-Me')) { New-Item -ItemType Directory -Path (Join-Path $purgeRoot $name) -Force | Out-Null }
    & $module { param($r) Remove-HealthCheckerPreviousWorkFolder -WorkRoot $r -CurrentRunId '20261021-060000' } $purgeRoot
    Assert-HealthChecker (-not (Test-Path (Join-Path $purgeRoot '20261001-010101')) -and -not (Test-Path (Join-Path $purgeRoot '20261020-010101')) -and (Test-Path (Join-Path $purgeRoot '20261021-060000')) -and (Test-Path (Join-Path $purgeRoot 'Keep-Me'))) 'Only the current run folder is kept.'

    # HealthCheckerExcludedServers: kept in the summary, never started or contacted, no warning.
    $defaultExcluded = & $module { $script:Config = @{}; (Get-HealthCheckerSetting).ExcludedServers }
    Assert-HealthChecker (@($defaultExcluded).Count -eq 0) 'No server is excluded by default.'
    $configuredExcluded = & $module { $script:Config = @{ HealthCheckerExcludedServers = @(' EXSRV02.corp.contoso.example ', 'EXSRV09') }; (Get-HealthCheckerSetting).ExcludedServers }
    Assert-HealthChecker ((@($configuredExcluded) -join '|') -eq 'EXSRV02.corp.contoso.example|EXSRV09') 'Configured exclusions are read and trimmed.'
    foreach ($badValue in @(@(''), @(1))) {
        $badRefused = $false
        try { & $module { param($v) $script:Config = @{ HealthCheckerExcludedServers = $v }; Get-HealthCheckerSetting } $badValue } catch { $badRefused = $_.Exception.Message -match 'HealthCheckerExcludedServers must be' }
        Assert-HealthChecker $badRefused 'An empty or non-text exclusion is refused.'
    }
    & $module { $script:Config = @{} }
    $exServer = [pscustomobject]@{ Name = 'EXSRV02'; Fqdn = 'EXSRV02.corp.contoso.example' }
    $exNoFqdn = [pscustomobject]@{ Name = 'EXSRV02' }
    Assert-HealthChecker ((& $module { param($s) Test-HealthCheckerServerExcluded -Server $s -ExcludedServers @('exsrv02') } $exServer) -and (& $module { param($s) Test-HealthCheckerServerExcluded -Server $s -ExcludedServers @('exsrv02.CORP.contoso.example') } $exServer)) 'A server is excluded by short name or FQDN, ignoring case.'
    Assert-HealthChecker (& $module { param($s) Test-HealthCheckerServerExcluded -Server $s -ExcludedServers @('EXSRV02.corp.contoso.example') } $exNoFqdn) 'A configured FQDN also matches a server known by its short name only.'
    Assert-HealthChecker (-not (& $module { param($s) Test-HealthCheckerServerExcluded -Server $s -ExcludedServers @('EXSRV0', 'EXSRV021', 'EXSRV20.corp.contoso.example') } $exServer) -and -not (& $module { param($s) Test-HealthCheckerServerExcluded -Server $s -ExcludedServers @() } $exServer)) 'Other names never exclude a server.'

    & $module { $script:ProcessCalls.Clear(); $script:RemoteStarts.Clear(); $script:LogLines.Clear(); $script:RemoteMode = 'Completed'; $script:LocalMode = $null }
    $excludeSettings = [pscustomobject]@{ TimeoutMinutesPerServer = 30; TotalBudgetMinutes = 1320; MaxResultAgeHours = 26; ExcludedServers = @('exsrv02.CORP.contoso.example'); WorkFolderPath = $workRoot; RemoteRootPath = 'C:\Program Files\SmartM365\HealthChecker'; TaskPath = '\WCH\'; TaskName = 'SmartM365 Exchange Health Checker'; PollSeconds = 0 }
    $excludedRun = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $excludeSettings
    $excludedRows = @($excludedRun.Summary)
    $excludedLog = (& $module { $script:LogLines }) -join "`n"
    Assert-HealthChecker (@(& $module { $script:RemoteStarts }).Count -eq 0 -and @(& $module { $script:ProcessCalls }).Count -eq 1) 'An excluded server is never started or contacted; the others still run.'
    Assert-HealthChecker ($excludedRows.Count -eq 2 -and $excludedRows[1].ServerName -eq 'EXSRV02' -and $excludedRows[1].RunStatus -eq 'NotRun' -and $excludedRows[1].RunMessage -eq 'Health Checker was not run: excluded by configuration.' -and $null -eq $excludedRows[1].ResultSource) 'The excluded server keeps a NotRun summary row with a short reason.'
    Assert-HealthChecker ($excludedRows[0].RunStatus -eq 'Completed' -and $excludedRows[0].ServersProcessed -eq 1 -and $excludedRows[0].ServersTotal -eq 1 -and @($excludedRun.Findings | Where-Object { $_.ServerName -eq 'EXSRV02' }).Count -eq 0) 'The excluded server is not counted in the processed total and has no findings.'
    Assert-HealthChecker ($excludedLog -notmatch '\[WARN\]' -and $excludedLog -match '\[INFO\] Health Checker excluded by configuration \(HealthCheckerExcludedServers\): EXSRV02\.' -and $excludedLog -match 'processed 1 of 1 server\(s\); 1 excluded') "An exclusion is information, not a warning: $excludedLog"
    & $module { $script:ProcessCalls.Clear(); $script:RemoteStarts.Clear(); $script:LogLines.Clear() }
    $excludeSettings.ExcludedServers = @('exsrv01', 'EXSRV02')
    $allExcluded = & $module $invoke $servers $vendoredPath $runnerSource $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $excludeSettings
    Assert-HealthChecker (@(& $module { $script:ProcessCalls }).Count -eq 0 -and @(& $module { $script:RemoteStarts }).Count -eq 0 -and @($allExcluded.Summary | Where-Object { $_.RunStatus -eq 'NotRun' }).Count -eq 2 -and ((& $module { $script:LogLines }) -join "`n") -notmatch '\[WARN\]') 'The collector server itself can be excluded by short name.'

    $collectorText = Get-Content -LiteralPath $collectorPath -Raw
    # Exercise the actual switch binder and production resolution statement.
    $activationNode = $ast.Find({
        param($n)
        $n -is [Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left.Extent.Text -eq '$IncludeHealthChecker' -and
        $n.Right.Extent.Text -like 'Resolve-SmartM365InfrastructureHealthCheckerEnabled*'
    }, $true)
    Assert-HealthChecker ($null -ne $activationNode) 'Activation is resolved in the entry workflow.'
    $activationProbe = [scriptblock]::Create(
        'param([switch]$IncludeHealthChecker) ' + $activationNode.Extent.Text +
        '; [pscustomobject]@{ Enabled=[bool]$IncludeHealthChecker; Scope=@{IncludeHealthChecker=$IncludeHealthChecker} }'
    )
    foreach ($configured in @('Missing', $true, $false)) {
        foreach ($explicit in @('Missing', $true, $false)) {
            $result = & $module {
                param($cfg, $arg, $probe)
                $script:Config = @{}
                if ($cfg -is [bool]) { $script:Config['IncludeHealthChecker'] = $cfg }
                $arguments = @{}
                if ($arg -is [bool]) { $arguments['IncludeHealthChecker'] = $arg }
                & $probe @arguments
            } $configured $explicit $activationProbe
            $expected = if ($explicit -is [bool]) { $explicit } elseif ($configured -is [bool]) { $configured } else { $true }
            Assert-HealthChecker ($result.Enabled -eq $expected) "Activation config=$configured parameter=$explicit."
            Assert-HealthChecker ([bool]$result.Scope.IncludeHealthChecker -eq $expected) 'Receipt scope records effective activation.'
        }
    }
    foreach ($bad in @('false', 'true', 0, 1, @())) {
        $rejected = $false
        try {
            & $module { param($v, $probe) $script:Config = @{IncludeHealthChecker=$v}; & $probe } $bad $activationProbe
        } catch { $rejected = $_.Exception.Message -match 'must be a JSON boolean' }
        Assert-HealthChecker $rejected 'Non-boolean configuration is rejected.'
    }
    $template = Get-Content -LiteralPath ($collectorPath -replace '\.ps1$', '.local.json.template') -Raw | ConvertFrom-Json
    Assert-HealthChecker ($template.IncludeHealthChecker -is [bool] -and $template.IncludeHealthChecker) 'Template enables Health Checker.'
    $tenMinutes = & $module { $script:Config=@{HealthCheckerTimeoutMinutesPerServer=10}; (Get-HealthCheckerSetting).TimeoutMinutesPerServer }
    Assert-HealthChecker ($tenMinutes -eq 10) 'Explicit 10-minute timeout is preserved.'
    $receiptCall = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Start-CoreSmartM365SourceReceipt' }, $true)
    Assert-HealthChecker ($activationNode.Extent.StartOffset -lt $receiptCall.Extent.StartOffset) 'Activation resolves before receipt scope.'
    Assert-HealthChecker ($collectorText.IndexOf('Invoke-ServersAndStorageHtmlReportRegeneration -RunFolder') -lt $activationNode.Extent.StartOffset) 'Report regeneration precedes activation and returns without collection.'
    & $module { $script:Config=@{} }

    # 6. Report section and disabled option.
    $sectionSummary = @($summaryRow, $notRunRow)
    $section = & $module { param($s, $f) ConvertTo-ServersAndStorageHealthCheckerSectionHtml -HealthCheckerSummary $s -HealthCheckerFindings $f } $sectionSummary $findings
    Assert-HealthChecker ($section -match '<h2>Exchange Health Checker</h2>' -and $section -match 'processed <strong>1</strong> of <strong>2</strong>') 'The section shows the processed server count.'
    Assert-HealthChecker ($section -match 'Vulnerabilities: <strong>2</strong> \| Errors: <strong>3</strong>') 'The section totals vulnerabilities and errors.'
    Assert-HealthChecker ($section.IndexOf('EXSRV01') -lt $section.IndexOf('EXSRV02')) 'Servers with vulnerabilities come first.'
    Assert-HealthChecker ($section.IndexOf('CVE-2024-21410') -lt $section.IndexOf('Out of date Cumulative Update')) 'Vulnerabilities are listed before other errors.'
    Assert-HealthChecker ($section -match 'Timeout' -and $section -match 'stopped') 'A server that did not complete shows its reason.'
    $noSection = & $module { ConvertTo-ServersAndStorageHealthCheckerSectionHtml -HealthCheckerSummary @() -HealthCheckerFindings @() }
    Assert-HealthChecker ($noSection -eq '') 'Without Health Checker data there is no section.'
    $reportSummary = [pscustomobject]@{ RunId = 'R1'; ExecutionDate = '2026-10-21'; ExchangeServersCount = 1; TotalLogicalProcessorCount = 4; TotalMemoryGB = 16; TotalDiskDriveCount = 1; TotalDiskDriveSizeTB = 1; ExchangeSchemaRangeUpper = 1; ExchangeOrgObjectVersion = 1 }
    $perServer = @([pscustomobject]@{ ExchangeServerName = 'EXSRV01'; ServerRole = 'Mailbox'; LogicalProcessorCount = 4; MemoryGB = 16; DiskDriveCount = 1; DiskDriveTotalSizeGB = 100; ComputeCollectionStatus = 'OK'; DiskDriveCollectionStatus = 'OK'; LogicalDiskCollectionStatus = 'OK'; LowSpaceLogicalDiskCount = 0 })
    $reportOff = Join-Path $tempRoot 'report-off.html'
    & $module { param($s, $p, $o) New-HtmlExecutiveSummary -Summary $s -PerServerSummary $p -ReadinessInventory @() -Path $o } $reportSummary $perServer $reportOff
    Assert-HealthChecker ((Get-Content -LiteralPath $reportOff -Raw) -notmatch 'Exchange Health Checker') 'The report has no Health Checker section when the option is off.'
    $reportOn = Join-Path $tempRoot 'report-on.html'
    & $module { param($s, $p, $o, $hs, $hf) New-HtmlExecutiveSummary -Summary $s -PerServerSummary $p -ReadinessInventory @() -HealthCheckerSummary $hs -HealthCheckerFindings $hf -Path $o } $reportSummary $perServer $reportOn $sectionSummary $findings
    $reportOnText = Get-Content -LiteralPath $reportOn -Raw
    Assert-HealthChecker ($reportOnText.IndexOf('Priority issues summary') -lt $reportOnText.IndexOf('Exchange Health Checker') -and $reportOnText.IndexOf('Exchange Health Checker') -lt $reportOnText.IndexOf('Mailbox databases')) 'The section follows the priority issues.'

    $collectorText = [IO.File]::ReadAllText($collectorPath)
    $hcCalls = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Invoke-ExchangeHealthCheckerInventory' }, $true))
    Assert-HealthChecker ($hcCalls.Count -eq 1) 'Health Checker is started from one place only.'
    $guarded = $false
    $ancestor = $hcCalls[0].Parent
    while ($ancestor) {
        if ($ancestor -is [Management.Automation.Language.IfStatementAst] -and $ancestor.Clauses[0].Item1.Extent.Text -eq '$IncludeHealthChecker') { $guarded = $true }
        $ancestor = $ancestor.Parent
    }
    Assert-HealthChecker $guarded 'Health Checker runs only when the resolved IncludeHealthChecker value is true.'
    Assert-HealthChecker ($collectorText -match '\[switch\]\$IncludeHealthChecker' -and $collectorText -match 'IncludeHealthChecker=\$IncludeHealthChecker') 'The switch exists and is recorded in the receipt scope.'
    Assert-HealthChecker ($collectorText -notmatch 'HealthChecker\.ps1[^\r\n]*-ScriptUpdateOnly') 'The collector never asks Health Checker to update itself.'
    $receiptContract = [IO.File]::ReadAllText((Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365-SourceReceipts.json.txt'))
    Assert-HealthChecker ($receiptContract -match 'Exchange_OnPrem_HealthChecker_Summary\.csv' -and $receiptContract -match 'Exchange_OnPrem_HealthChecker_Findings\.csv') 'Both CSVs are optional receipt outputs.'

    # 7. Remote helpers: paths, messages, task logon failures, runner states and the scheduled fallback.
    Assert-HealthChecker ((& $module { ConvertTo-HealthCheckerAdminSharePath -Computer 'EXSRV02.corp.contoso.example' -LocalPath 'C:\Program Files\SmartM365\HealthChecker' }) -eq '\\EXSRV02.corp.contoso.example\C$\Program Files\SmartM365\HealthChecker') 'The remote root maps to the administrative share.'
    $shareRefused = $false
    try { & $module { ConvertTo-HealthCheckerAdminSharePath -Computer 'EXSRV02' -LocalPath '\\server\share' } } catch { $shareRefused = $true }
    Assert-HealthChecker $shareRefused 'A remote root without a drive letter is refused.'
    $cleanMessage = & $module { ConvertTo-HealthCheckerRemoteMessage -Message 'Access to \\EXSRV02.corp.contoso.example\C$ from 192.0.2.4 failed for EXSRV09.CORP.contoso.example.' -Computer 'EXSRV02.corp.contoso.example' }
    Assert-HealthChecker ($cleanMessage -notmatch 'contoso|192\.0' -and $cleanMessage -match 'EXSRV02' -and $cleanMessage -match 'EXSRV09') "Remote messages keep host names without the domain or addresses: $cleanMessage"
    Assert-HealthChecker ($null -eq (& $module { ConvertTo-HealthCheckerRemoteMessage -Message '' -Computer 'EXSRV02' })) 'An empty message stays empty.'
    $logon = & $module { Get-HealthCheckerTaskLogonFailure -LastTaskResult ([uint32]2147943726) }
    Assert-HealthChecker ($logon -match '0x8007052E' -and $logon -match 'wrong user name or password' -and $logon -match 'run the Health Checker task installer again') "A wrong password is reported with the installer action: $logon"
    Assert-HealthChecker ((& $module { Get-HealthCheckerTaskLogonFailure -LastTaskResult ([int]-2147023511) }) -match "Log on as a batch job") 'A signed result code is read as the batch logon right failure.'
    Assert-HealthChecker ($null -eq (& $module { Get-HealthCheckerTaskLogonFailure -LastTaskResult 0 }) -and $null -eq (& $module { Get-HealthCheckerTaskLogonFailure -LastTaskResult 267011 })) 'Success and not-yet-run results are not logon failures.'
    $stateMap = @{ Completed = 'Completed'; Timeout = 'TaskTimeout'; IntegrityFailed = 'IntegrityFailed'; NotElevated = 'NotElevated'; LaunchFailed = 'LaunchFailed'; ExchangeShellFailed = 'NoOutput'; NoOutput = 'NoOutput'; Running = 'NoOutput' }
    foreach ($state in $stateMap.Keys) {
        $mapped = & $module { param($s) ConvertFrom-HealthCheckerRunnerState -RemoteStatus ([pscustomobject]@{ State = $s; Message = '' }) } $state
        Assert-HealthChecker ($mapped.Status -eq $stateMap[$state]) "Runner state $state maps to $($stateMap[$state]) (found $($mapped.Status))."
    }
    $nowUtc = [DateTime]::UtcNow
    $fresh = [pscustomobject]@{ State = 'Completed'; ResultFileName = 'result.json'; FinishedAtUtc = $nowUtc.AddHours(-5).ToString('o') }
    $freshResult = & $module { param($r, $n) Resolve-HealthCheckerScheduledFallback -FailureStatus 'RemoteConnectionFailed' -FailureMessage 'unreachable.' -RemoteStatus $r -NowUtc $n -MaxResultAgeHours 26 } $fresh $nowUtc
    Assert-HealthChecker ($freshResult.Status -eq 'Completed' -and $freshResult.UseResult -and $freshResult.ResultAgeHours -eq 5) 'A recent scheduled result is used.'
    $fresh.FinishedAtUtc = $nowUtc.AddHours(-26.04).ToString('o')
    $limitResult = & $module { param($r, $n) Resolve-HealthCheckerScheduledFallback -FailureStatus 'RemoteConnectionFailed' -FailureMessage 'x' -RemoteStatus $r -NowUtc $n -MaxResultAgeHours 26 } $fresh $nowUtc
    Assert-HealthChecker ($limitResult.Status -eq 'Completed') 'A result at the age limit is still used.'
    $fresh.FinishedAtUtc = [datetime]$nowUtc.AddHours(-48)
    $oldResult = & $module { param($r, $n) Resolve-HealthCheckerScheduledFallback -FailureStatus 'TaskTriggerFailed' -FailureMessage 'x' -RemoteStatus $r -NowUtc $n -MaxResultAgeHours 26 } $fresh $nowUtc
    Assert-HealthChecker ($oldResult.Status -eq 'StaleResult' -and -not $oldResult.UseResult -and $oldResult.ResultAgeHours -eq 48) 'An old result (DateTime from PowerShell 7 JSON) is StaleResult.'
    foreach ($unusable in @($null, [pscustomobject]@{ State = 'NoOutput'; ResultFileName = $null; FinishedAtUtc = $nowUtc.ToString('o') }, [pscustomobject]@{ State = 'Completed'; ResultFileName = 'result.json'; FinishedAtUtc = 'not a date' }, [pscustomobject]@{ State = 'Completed'; ResultFileName = 'Raw\HealthChecker-X-1.xml'; FinishedAtUtc = $nowUtc.ToString('o') })) {
        $kept = & $module { param($r, $n) Resolve-HealthCheckerScheduledFallback -FailureStatus 'TaskTriggerFailed' -FailureMessage 'kept' -RemoteStatus $r -NowUtc $n -MaxResultAgeHours 26 } $unusable $nowUtc
        Assert-HealthChecker ($kept.Status -eq 'TaskTriggerFailed' -and $kept.Message -eq 'kept' -and -not $kept.UseResult) 'Without a usable result the failure status is kept.'
    }

    # Progress of a remote run is read from status.json: a new run token, or the awaited one.
    $progressShare = Join-Path $tempRoot 'progress'
    New-Item -ItemType Directory -Path $progressShare -Force | Out-Null
    $writeStatus = { param($Token, $State) [IO.File]::WriteAllText((Join-Path $progressShare 'status.json'), ([pscustomobject]@{ RunToken = $Token; State = $State } | ConvertTo-Json)) }
    $progressRun = [pscustomobject]@{ ServerName = 'EXSRV02'; Computer = 'EXSRV02'; Session = $null; OutputSharePath = $progressShare; PreviousStatus = [pscustomobject]@{ RunToken = 'old'; State = 'Completed' }; WaitToken = $null; StartedAt = (Get-Date); Done = $false; Status = $null; Message = $null; RemoteStatus = $null }
    & $writeStatus 'old' 'Completed'
    & $module { param($r, $s) Update-HealthCheckerRemoteRun -Run $r -Settings $s } $progressRun $runSettings
    Assert-HealthChecker (-not $progressRun.Done) 'The previous run result is never taken for the new run.'
    & $writeStatus 'new' 'Running'
    & $module { param($r, $s) Update-HealthCheckerRemoteRun -Run $r -Settings $s } $progressRun $runSettings
    Assert-HealthChecker (-not $progressRun.Done) 'A running run is awaited.'
    & $writeStatus 'new' 'NoOutput'
    & $module { param($r, $s) Update-HealthCheckerRemoteRun -Run $r -Settings $s } $progressRun $runSettings
    Assert-HealthChecker ($progressRun.Done -and $progressRun.RemoteStatus.State -eq 'NoOutput') 'A finished new run is taken.'
    $waitRun = [pscustomobject]@{ ServerName = 'EXSRV02'; Computer = 'EXSRV02'; Session = $null; OutputSharePath = $progressShare; PreviousStatus = [pscustomobject]@{ RunToken = 'daily'; State = 'Running' }; WaitToken = 'daily'; StartedAt = (Get-Date); Done = $false; Status = $null; Message = $null; RemoteStatus = $null }
    & $writeStatus 'daily' 'Completed'
    & $module { param($r, $s) Update-HealthCheckerRemoteRun -Run $r -Settings $s } $waitRun $runSettings
    Assert-HealthChecker ($waitRun.Done -and $waitRun.RemoteStatus.State -eq 'Completed') 'A run already in progress at start time is awaited instead of a new one.'

    # A task that does not start is detected after the grace period, not at the end of the per-server limit.
    & $module {
        $script:TaskInfoResult = 267011
        $script:TaskState = 'Ready'
        function script:Get-ScheduledTaskInfo { param($CimSession, $TaskPath, $TaskName, $ErrorAction) [pscustomobject]@{ LastTaskResult = $script:TaskInfoResult; LastRunTime = (Get-Date).AddMinutes(-4) } }
        function script:Get-ScheduledTask { param($CimSession, $TaskPath, $TaskName, $ErrorAction) [pscustomobject]@{ State = $script:TaskState } }
    }
    & $writeStatus 'old' 'Completed'
    $graceSettings = [pscustomobject]@{ TaskPath = '\WCH\'; TaskName = 'SmartM365 Exchange Health Checker'; TaskStartGraceSeconds = 180 }
    $newGraceRun = { param([int]$AgeSeconds) [pscustomobject]@{ ServerName = 'EXSRV02'; Computer = 'EXSRV02'; Session = $null; OutputSharePath = $progressShare; PreviousStatus = [pscustomobject]@{ RunToken = 'old'; State = 'Completed' }; WaitToken = $null; StartedAt = (Get-Date).AddSeconds(-$AgeSeconds); Done = $false; Status = $null; Message = $null; RemoteStatus = $null } }
    $earlyRun = & $newGraceRun 90
    & $module { param($r, $s) Update-HealthCheckerRemoteRun -Run $r -Settings $s } $earlyRun $graceSettings
    Assert-HealthChecker (-not $earlyRun.Done) 'Within the grace period the task is still awaited.'
    $deadRun = & $newGraceRun 200
    & $module { param($r, $s) Update-HealthCheckerRemoteRun -Run $r -Settings $s } $deadRun $graceSettings
    Assert-HealthChecker ($deadRun.Done -and $deadRun.Status -eq 'TaskTriggerFailed' -and $deadRun.Message -match '0x00041303' -and $deadRun.Message -match 'local administrator' -and $deadRun.Message -match 'batch job' -and $deadRun.Message -match 'password') "A task not running without a fresh status after 3 min is TaskTriggerFailed with its hexadecimal result: $($deadRun.Message)"
    & $module { $script:TaskState = 'Running' }
    $runningRun = & $newGraceRun 200
    & $module { param($r, $s) Update-HealthCheckerRemoteRun -Run $r -Settings $s } $runningRun $graceSettings
    Assert-HealthChecker (-not $runningRun.Done) 'A task still running is awaited after the grace period.'
    & $module { $script:TaskState = 'Ready'; $script:TaskInfoResult = 2147943726 }
    $logonRun = & $newGraceRun 70
    & $module { param($r, $s) Update-HealthCheckerRemoteRun -Run $r -Settings $s } $logonRun $graceSettings
    Assert-HealthChecker ($logonRun.Done -and $logonRun.Message -match 'could not log on \(0x8007052E') 'A logon failure is still reported after one minute.'
    & $module { Remove-Item function:script:Get-ScheduledTaskInfo, function:script:Get-ScheduledTask -ErrorAction SilentlyContinue }
    Assert-HealthChecker (((& $module { (Get-HealthCheckerSetting).TaskStartGraceSeconds })) -eq 180) 'The grace period is three minutes.'
    $badName = [pscustomobject]@{ ServerName = 'EXSRV02'; Computer = 'EXSRV02'; Session = $null; OutputSharePath = $progressShare; PreviousStatus = $null; WaitToken = $null; StartedAt = (Get-Date); Done = $true; Status = $null; Message = $null; RemoteStatus = [pscustomobject]@{ State = 'Completed'; ResultFileName = '..\..\Windows\win.ini'; FinishedAtUtc = $nowUtc.ToString('o'); DurationSeconds = 1 } }
    $badOutcome = & $module { param($r, $f, $s) Complete-HealthCheckerRemoteRun -Run $r -LocalFolder $f -Settings $s } $badName (Join-Path $tempRoot 'bad-name') $runSettings
    Assert-HealthChecker ($badOutcome.Status -eq 'NoOutput' -and $null -eq $badOutcome.JsonPath -and $null -eq $badOutcome.XmlPath) 'Only result.json is ever copied from a server.'

    # Server files are copied only when they differ.
    $copySource = Join-Path $tempRoot 'copy-source.txt'
    $copyTarget = Join-Path $tempRoot 'copy-target.txt'
    [IO.File]::WriteAllText($copySource, 'v1')
    Assert-HealthChecker ((& $module { param($s, $d) Copy-HealthCheckerRemoteFile -SourcePath $s -DestinationPath $d } $copySource $copyTarget) -eq $true) 'A missing file is copied.'
    Assert-HealthChecker ((& $module { param($s, $d) Copy-HealthCheckerRemoteFile -SourcePath $s -DestinationPath $d } $copySource $copyTarget) -eq $false) 'An identical file is not copied again.'
    [IO.File]::WriteAllText($copyTarget, 'tampered')
    Assert-HealthChecker ((& $module { param($s, $d) Copy-HealthCheckerRemoteFile -SourcePath $s -DestinationPath $d } $copySource $copyTarget) -eq $true -and [IO.File]::ReadAllText($copyTarget) -eq 'v1') 'A file that differs is replaced.'

    # Remote results: plain JSON lines, size limit, strict schema, same findings as the XML.
    $plainLines = @(& $module { param($p) ConvertTo-HealthCheckerResultLine -Path $p } $fixture)
    $jsonPath = Join-Path $tempRoot 'result.json'
    [IO.File]::WriteAllText($jsonPath, ([pscustomobject]@{ SchemaVersion = 1; HealthCheckerVersion = 'x'; Lines = $plainLines } | ConvertTo-Json -Depth 4 -Compress))
    $jsonLines = @(& $module { param($p) Read-HealthCheckerResultJson -Path $p } $jsonPath)
    $fromJson = & $module { param($l) ConvertFrom-HealthCheckerResultLine -Lines $l -ServerName 'EXSRV01' -HealthCheckerVersion '26.10.13.1644' } $jsonLines
    $fromXml = & $module { param($p) ConvertFrom-HealthCheckerXml -Path $p -ServerName 'EXSRV01' -HealthCheckerVersion '26.10.13.1644' } $fixture
    Assert-HealthChecker ((@($fromJson.Findings) | ConvertTo-Json -Depth 3 -Compress) -ceq (@($fromXml.Findings) | ConvertTo-Json -Depth 3 -Compress) -and ($fromJson.Facts | ConvertTo-Json -Compress) -ceq ($fromXml.Facts | ConvertTo-Json -Compress)) 'Findings and facts from result.json equal those from the XML.'
    Assert-HealthChecker (@($plainLines | Where-Object { $_.DisplayValue -isnot [string] -and $null -ne $_.DisplayValue }).Count -eq 0 -and @($plainLines | Where-Object { $_.WriteType -eq 'OutColumns' -and $null -ne $_.DisplayValue }).Count -eq 0) 'Result lines hold text only, and tables are not carried.'
    $refusedJson = {
        param([string]$Content, [int]$MaxBytes = 20MB)
        $path = Join-Path $tempRoot ('bad-' + [guid]::NewGuid().ToString('N') + '.json')
        [IO.File]::WriteAllText($path, $Content)
        $refused = $false
        try { [void](& $module { param($p, $m) Read-HealthCheckerResultJson -Path $p -MaxBytes $m } $path $MaxBytes) } catch { $refused = $true }
        return $refused
    }
    Assert-HealthChecker (& $refusedJson ([IO.File]::ReadAllText($jsonPath)) 100) 'A result over the size limit is refused before it is read.'
    Assert-HealthChecker (& $refusedJson '{"SchemaVersion":2,"Lines":[{"Category":"A"}]}') 'Another schema version is refused.'
    Assert-HealthChecker (& $refusedJson '{"SchemaVersion":1,"Lines":[{"Category":"A","DisplayValue":{"Nested":"object"}}]}') 'A nested object in a field is refused.'
    Assert-HealthChecker (& $refusedJson '{"SchemaVersion":1,"Lines":["text"]}') 'A line that is not an object is refused.'
    Assert-HealthChecker (& $refusedJson '{"SchemaVersion":1,"Lines":[]}') 'A result without lines is refused.'
    Assert-HealthChecker (& $refusedJson ([IO.File]::ReadAllText($fixture))) 'A CliXml document is never accepted as a remote result.'
    $extraPath = Join-Path $tempRoot 'extra.json'
    [IO.File]::WriteAllText($extraPath, '{"SchemaVersion":1,"Lines":[{"Category":"Exchange Information","CategoryOrder":"1","Check":"Name","Name":"Name","WriteType":"Grey","DisplayValue":42,"Unexpected":{"Type":"x"}}]}')
    $extraLines = @(& $module { param($p) Read-HealthCheckerResultJson -Path $p } $extraPath)
    Assert-HealthChecker ($extraLines.Count -eq 1 -and $extraLines[0].DisplayValue -ceq '42' -and $extraLines[0].CategoryOrder -eq 1 -and -not $extraLines[0].PSObject.Properties['Unexpected']) 'Only the expected fields are kept, as strings and numbers.'
    $clixmlCalls = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Import-Clixml' }, $true))
    $clixmlOwner = if ($clixmlCalls.Count -eq 1) { $node = $clixmlCalls[0].Parent; while ($node -and $node -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $node = $node.Parent }; $node.Name } else { '' }
    Assert-HealthChecker ($clixmlOwner -eq 'ConvertTo-HealthCheckerResultLine') 'Import-Clixml is used in one place only.'
    $xmlParsers = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in @('ConvertFrom-HealthCheckerXml', 'ConvertTo-HealthCheckerResultLine') }, $true) | ForEach-Object { $node = $_.Parent; while ($node -and $node -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $node = $node.Parent }; $node.Name } | Sort-Object -Unique)
    Assert-HealthChecker (($xmlParsers -join ',') -eq 'ConvertFrom-HealthCheckerXml,Invoke-ExchangeHealthCheckerInventory') "XML is only read by the local-run path ($($xmlParsers -join ','))."
    $remoteCompleteText = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Complete-HealthCheckerRemoteRun' }, $true).Extent.Text
    Assert-HealthChecker ($remoteCompleteText -notmatch '(?i)\.xml|Clixml') 'Nothing XML is copied from a server.'

    # Generic rights are not mapped by .NET: GENERIC_WRITE or GENERIC_ALL for a non-administrator is refused.
    foreach ($generic in @('GW', 'GA')) {
        $genericAcl = New-Object System.Security.AccessControl.DirectorySecurity
        $genericAcl.SetSecurityDescriptorSddlForm(('O:BAG:BAD:P(A;OICI;FA;;;BA)(A;OICIIO;{0};;;AU)' -f $generic))
        Assert-HealthChecker (@(& $module { param($a) Get-HealthCheckerUnsafeAclEntry -Acl $a } $genericAcl).Count -eq 1) ("Generic right {0} for Authenticated Users is refused." -f $generic)
    }

    # 8. Local runner deployed to the servers.
    $runnerPath = Join-Path $collectorFolder $constants.HealthCheckerRunnerFileName
    $runnerAst = [Management.Automation.Language.Parser]::ParseFile($runnerPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Runner parse failed: $parseErrors" }
    $runnerText = [IO.File]::ReadAllText($runnerPath)
    foreach ($constantName in @('HealthCheckerVersion', 'HealthCheckerExpectedSha256')) {
        $assignment = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq "`$$constantName" }, $true)
        Assert-HealthChecker ($null -ne $assignment -and $assignment.Right.Extent.Text.Trim("'", '"') -eq $constants[$constantName]) "The runner $constantName matches the collector."
    }
    $collectorIntegrity = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-HealthCheckerScriptIntegrity' }, $true).Extent.Text
    $runnerIntegrity = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-HealthCheckerScriptIntegrity' }, $true).Extent.Text
    Assert-HealthChecker ($runnerIntegrity -ceq $collectorIntegrity) 'The runner checks Health Checker exactly like the collector.'
    foreach ($sharedName in @('Get-HealthCheckerLineValue', 'ConvertTo-HealthCheckerResultLine')) {
        $collectorShared = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $sharedName }, $true).Extent.Text
        $runnerShared = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $sharedName }, $true)
        Assert-HealthChecker ($null -ne $runnerShared -and $runnerShared.Extent.Text -ceq $collectorShared) "The runner converts results with the collector's $sharedName."
    }
    Assert-HealthChecker (((@($runnerAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })) -join ',') -eq 'TimeoutMinutes') 'The runner has no -Server parameter.'
    Assert-HealthChecker ($runnerText -notmatch "'-Server'" -and $runnerText -notmatch 'SmartM365\.Core|Import-Module|InventoryLifecycle|TenantContext') 'The runner never passes -Server and is standalone (no repository module).'
    Assert-HealthChecker ($runnerText -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' -and $runnerText -match 'Global\\SmartM365-Exchange-HealthChecker-LocalRunner') 'The runner forces Windows PowerShell 5.1 and takes a machine-wide lock.'
    $runnerFunctions = foreach ($node in $runnerAst.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) { $node.Extent.Text }
    $runnerModule = New-Module -ScriptBlock ([scriptblock]::Create(($runnerFunctions -join "`n")))
    & $runnerModule {
        param($Version, $Hash)
        Set-StrictMode -Version Latest
        $script:ScriptVersion = '1.1.1'; $script:HealthCheckerVersion = $Version; $script:HealthCheckerExpectedSha256 = $Hash
        $script:IntegrityOk = $true
        $script:RunnerElevated = $true
    } $constants.HealthCheckerVersion $constants.HealthCheckerExpectedSha256
    $runnerArgs = & $runnerModule { Get-LocalHealthCheckerArgument -ScriptPath 'C:\Root\Bin\HealthChecker.ps1' -RawFolder 'C:\Root\Output\Raw' }
    Assert-HealthChecker (($runnerArgs -join ' ') -notmatch '-Server' -and ($runnerArgs -join ' ') -match '-OutputFilePath "C:\\Root\\Output\\Raw" -SkipVersionCheck') "The runner command line has no -Server ($($runnerArgs -join ' '))."

    $env:SMARTM365_FAKE_HC_FIXTURE = $fixture
    try {
        $runnerRoot = Join-Path $tempRoot 'runner-root'
        New-Item -ItemType Directory -Path (Join-Path $runnerRoot 'Bin'), (Join-Path $runnerRoot 'Output\Raw') -Force | Out-Null
        Copy-Item -LiteralPath $fakeScript -Destination (Join-Path $runnerRoot 'Bin\HealthChecker.ps1')
        [IO.File]::WriteAllText((Join-Path $runnerRoot 'Output\Raw\previous-run.txt'), 'old')

        $integrityRun = & $runnerModule { param($r) Invoke-LocalHealthCheckerRun -RootFolder $r -TimeoutSeconds 120 } $runnerRoot
        $statusFile = Get-Content -LiteralPath (Join-Path $runnerRoot 'Output\status.json') -Raw | ConvertFrom-Json
        Assert-HealthChecker ($integrityRun.State -eq 'IntegrityFailed' -and $statusFile.State -eq 'IntegrityFailed' -and $statusFile.Message -match 'SHA256 mismatch') 'The runner refuses a file that is not the expected Health Checker.'
        Assert-HealthChecker (-not (Test-Path -LiteralPath (Join-Path $runnerRoot 'Output\Raw\previous-run.txt'))) 'The previous run output is removed first.'
        Assert-HealthChecker ($statusFile.Message -notmatch [regex]::Escape($runnerRoot) -and $statusFile.RunToken -match '^[0-9a-f-]{36}$' -and $statusFile.FinishedAtUtc -and $statusFile.RunnerVersion -eq '1.1.1') 'The status holds a run token, times and versions without local paths.'
        $statusProperties = @($statusFile.PSObject.Properties.Name) -join ','
        Assert-HealthChecker ($statusProperties -eq 'SchemaVersion,RunToken,State,Message,RunnerVersion,HealthCheckerVersion,StartedAtUtc,FinishedAtUtc,DurationSeconds,ExitCode,ResultFileName') "The status has no host or account field ($statusProperties)."

        # The real check is replaced to run the synthetic stand-in; Microsoft's script is never run here.
        & $runnerModule {
            function script:Test-HealthCheckerScriptIntegrity { param($Path, $ExpectedSha256) [pscustomobject]@{ IsValid = $true; Reason = $null } }
            function script:Test-LocalRunnerElevation { return $true }
        }
        $env:SMARTM365_FAKE_HC_MODE = 'ok'
        $okRun = & $runnerModule { param($r) Invoke-LocalHealthCheckerRun -RootFolder $r -TimeoutSeconds 120 } $runnerRoot
        $runnerJsonPath = Join-Path $runnerRoot 'Output\result.json'
        Assert-HealthChecker ($okRun.State -eq 'Completed' -and $okRun.ResultFileName -eq 'result.json' -and (Test-Path -LiteralPath $runnerJsonPath)) "A completed run converts its XML to result.json on the server ($($okRun.State) $($okRun.Message))."
        $runnerLines = @(& $module { param($p) Read-HealthCheckerResultJson -Path $p } $runnerJsonPath)
        $runnerParsed = & $module { param($l) ConvertFrom-HealthCheckerResultLine -Lines $l -ServerName 'EXSRV01' -HealthCheckerVersion 'v' } $runnerLines
        Assert-HealthChecker (@($runnerParsed.Findings).Count -eq 17 -and @($runnerParsed.Vulnerabilities).Count -eq 2) 'The collector reads the runner result.json into the same findings.'
        $runnerFakeArgs = Get-Content -LiteralPath (Join-Path $runnerRoot 'Output\Raw\fake-args.txt') -Raw
        Assert-HealthChecker ($runnerFakeArgs -match 'Server=;SkipVersionCheck=True') "Health Checker gets no -Server from the runner ($runnerFakeArgs)."

        $env:SMARTM365_FAKE_HC_MODE = 'noshell'
        $shellRun = & $runnerModule { param($r) Invoke-LocalHealthCheckerRun -RootFolder $r -TimeoutSeconds 120 } $runnerRoot
        Assert-HealthChecker ($shellRun.State -eq 'ExchangeShellFailed' -and $shellRun.Message -match 'task account') 'An Exchange Management Shell load failure is reported explicitly.'
        Assert-HealthChecker (@(Get-ChildItem -LiteralPath (Join-Path $runnerRoot 'Output\Raw') -Filter '*.xml').Count -eq 0) 'Only the last run output is kept on the server.'
    }
    finally {
        Remove-Item Env:\SMARTM365_FAKE_HC_MODE -ErrorAction SilentlyContinue
        Remove-Item Env:\SMARTM365_FAKE_HC_FIXTURE -ErrorAction SilentlyContinue
    }

    # 9. Scheduled-task installer.
    $installerPath = Join-Path $collectorFolder 'Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1'
    $installerAst = [Management.Automation.Language.Parser]::ParseFile($installerPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Installer parse failed: $parseErrors" }
    $installerText = [IO.File]::ReadAllText($installerPath)
    $installerValue = { param($Name) $node = $installerAst.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq "`$$Name" }, $true); if ($node) { $node.Right.Extent.Text.Trim("'", '"') } }
    Assert-HealthChecker ((& $installerValue 'taskPath') -eq $constants.HealthCheckerTaskPath -and (& $installerValue 'taskName') -eq $constants.HealthCheckerTaskName -and (& $installerValue 'remoteRootPath') -eq $constants.HealthCheckerRemoteRootPath) 'The installer registers the task and folders the collector uses.'
    Assert-HealthChecker ((& $installerValue 'healthCheckerExpectedSha256') -eq $constants.HealthCheckerExpectedSha256 -and (& $installerValue 'runnerFileName') -eq $constants.HealthCheckerRunnerFileName) 'The installer deploys the expected Health Checker and runner.'
    $installerIntegrity = $installerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-HealthCheckerScriptIntegrity' }, $true).Extent.Text
    Assert-HealthChecker ($installerIntegrity -ceq $collectorIntegrity) 'The installer checks Health Checker exactly like the collector before deploying it.'
    Assert-HealthChecker ($installerText -notmatch '(?i)gmsa|managed service account|-LogonType') 'Only a classic service account is supported.'
    $passwordLines = @($installerText -split "`r?`n" | Where-Object { $_ -match 'plainTextPassword' } | ForEach-Object { $_.Trim() })
    $allowedPasswordLines = @('$plainTextPassword = $null', '$plainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)', 'Password = $plainTextPassword')
    Assert-HealthChecker (@($passwordLines | Where-Object { $_ -notin $allowedPasswordLines }).Count -eq 0 -and $installerText -match 'ZeroFreeBSTR\(\$passwordPointer\)') "The password is only passed to Register-ScheduledTask and cleared ($($passwordLines -join ' | '))."
    Assert-HealthChecker ($installerText -match "Get-Credential @credentialParameters" -and $installerText -match "'s-1-5-18'" -and $installerText -match 'credentialSid -ne \$serviceAccountSid') 'The password is typed in Get-Credential for the named account, and SYSTEM is refused.'
    Assert-HealthChecker ($installerText -match 'New-HealthCheckerDcomSessionOption' -and $installerText -notmatch '-Protocol Dcom' -and $installerText -notmatch 'Invoke-Command|New-PSSession|Enter-PSSession|schtasks') 'The installer uses CIM over DCOM, never WinRM or schtasks.'
    Assert-HealthChecker ($installerText -match "MultipleInstances IgnoreNew" -and $installerText -match "RunLevel = 'Highest'" -and $installerText -match "New-TimeSpan -Hours 1" -and $installerText -match '\[string\]\$DailyAt = ''01:00''') 'The task runs daily at 01:00, elevated, one run at a time, at most one hour.'
    Assert-HealthChecker ((& $installerValue 'remoteRootPath') -like 'C:\Program Files\*') 'The elevated task runs from Program Files, where standard users cannot create files.'
    Assert-HealthChecker ($installerText -match 'Assert-HealthCheckerParentFolder -Path \(Split-Path -Path \$rootShare -Parent\)\r?\n\s*Set-HealthCheckerFolder -Path \$rootShare') 'The parent folder is checked before the task folders are created.'
    $installerFunctions = foreach ($name in @('Set-HealthCheckerFolderAclRule', 'ConvertTo-AdministrativeSharePath', 'Get-HealthCheckerUnsafeAclEntry')) { $installerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true).Extent.Text }
    $installerModule = New-Module -ScriptBlock ([scriptblock]::Create(($installerFunctions -join "`n")))
    Assert-HealthChecker ((& $installerModule { ConvertTo-AdministrativeSharePath -Computer 'EXSRV02' -LocalPath 'C:\Program Files\SmartM365\HealthChecker' }) -eq '\\EXSRV02\C$\Program Files\SmartM365\HealthChecker') 'The installer maps the remote root like the collector.'
    $aclFolder = Join-Path $tempRoot 'acl'
    New-Item -ItemType Directory -Path $aclFolder -Force | Out-Null
    $currentSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $aclResult = & $installerModule { param($p, $s) Set-HealthCheckerFolderAclRule -Acl (Get-Acl -LiteralPath $p) -ServiceAccountSid $s -ServiceAccountRight 'Modify' } $aclFolder $currentSid
    $aclRules = @($aclResult.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))
    $ruleText = @($aclRules | ForEach-Object { '{0}={1}' -f $_.IdentityReference.Value, $_.FileSystemRights }) -join ';'
    Assert-HealthChecker ($aclResult.AreAccessRulesProtected -and $aclRules.Count -eq 3 -and $ruleText -match 'S-1-5-18=FullControl' -and $ruleText -match 'S-1-5-32-544=FullControl' -and $ruleText -match ([regex]::Escape($currentSid) + '=Modify')) "Folder permissions are explicit: SYSTEM, Administrators and the service account only ($ruleText)."
    $installerUnsafe = $installerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-HealthCheckerUnsafeAclEntry' }, $true).Extent.Text
    $collectorUnsafe = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-HealthCheckerUnsafeAclEntry' }, $true).Extent.Text
    Assert-HealthChecker ($installerUnsafe -ceq $collectorUnsafe) 'The collector work folder and the server folders use the same parent check.'
    $ancestors = @(& $module { Get-HealthCheckerWorkFolderAncestor -Path 'C:\Program Files\SmartM365\HealthChecker\Work' -TrustedBase @('C:\Program Files', 'C:\ProgramData') })
    Assert-HealthChecker (($ancestors -join '|') -eq 'C:\Program Files\SmartM365\HealthChecker|C:\Program Files\SmartM365') "Every parent below Program Files is checked, nearest first ($($ancestors -join '|'))."
    $dataAncestors = @(& $module { Get-HealthCheckerWorkFolderAncestor -Path 'C:\ProgramData\SmartM365\HealthChecker' -TrustedBase @('C:\Program Files', 'C:\ProgramData') })
    Assert-HealthChecker (($dataAncestors -join '|') -eq 'C:\ProgramData\SmartM365') 'A configured ProgramData work folder has its SmartM365 parent checked.'
    $otherAncestors = @(& $module { Get-HealthCheckerWorkFolderAncestor -Path 'D:\Tools\HealthChecker' -TrustedBase @('C:\Program Files') })
    Assert-HealthChecker (($otherAncestors -join '|') -eq 'D:\Tools') 'Parents are checked up to the drive root.'
    $workAcl = & $module { param($p) Set-HealthCheckerWorkFolderAclRule -Acl (Get-Acl -LiteralPath $p) } $aclFolder
    $workRules = @($workAcl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]) | ForEach-Object { '{0}={1}' -f $_.IdentityReference.Value, $_.FileSystemRights })
    Assert-HealthChecker ($workAcl.AreAccessRulesProtected -and $workAcl.GetOwner([Security.Principal.SecurityIdentifier]).Value -eq 'S-1-5-32-544' -and ($workRules -join ';') -eq 'S-1-5-18=FullControl;S-1-5-32-544=FullControl') "The work folder is owned by Administrators and only SYSTEM and Administrators have access ($($workRules -join ';'))."    Assert-HealthChecker ($aclResult.GetOwner([Security.Principal.SecurityIdentifier]).Value -eq 'S-1-5-32-544') 'Task folders are owned by Administrators.'
    $newAcl = {
        param([string]$Owner, [string[]]$Rules)
        $sec = New-Object System.Security.AccessControl.DirectorySecurity
        $sec.SetOwner((New-Object Security.Principal.SecurityIdentifier($Owner)))
        foreach ($rule in $Rules) {
            $sid, $right = $rule -split '='
            $sec.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier($sid)), [System.Security.AccessControl.FileSystemRights]$right, 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
        }
        $sec
    }
    $safeParent = & $newAcl 'S-1-5-32-544' @('S-1-5-18=FullControl', 'S-1-5-32-544=FullControl', 'S-1-5-32-545=ReadAndExecute', 'S-1-3-0=FullControl', 'S-1-15-2-1=ReadAndExecute')
    Assert-HealthChecker (@(& $installerModule { param($a) Get-HealthCheckerUnsafeAclEntry -Acl $a } $safeParent).Count -eq 0) 'A Program Files-like parent is accepted.'
    $writableParent = & $newAcl 'S-1-5-32-544' @('S-1-5-32-544=FullControl', 'S-1-5-32-545=Modify')
    $writableIssues = @(& $installerModule { param($a) Get-HealthCheckerUnsafeAclEntry -Acl $a } $writableParent)
    Assert-HealthChecker ($writableIssues.Count -eq 1 -and $writableIssues[0] -match '^S-1-5-32-545 ') "A parent writable by Users is refused ($($writableIssues -join '; '))."
    $userOwnedParent = & $newAcl 'S-1-5-32-545' @('S-1-5-32-544=FullControl')
    Assert-HealthChecker (@(& $installerModule { param($a) Get-HealthCheckerUnsafeAclEntry -Acl $a } $userOwnedParent) -contains 'owner S-1-5-32-545') 'A parent owned by a non-administrator is refused.'
    $createOnly = & $newAcl 'S-1-5-32-544' @('S-1-5-32-544=FullControl', 'S-1-5-11=CreateDirectories')
    Assert-HealthChecker (@(& $installerModule { param($a) Get-HealthCheckerUnsafeAclEntry -Acl $a } $createOnly).Count -eq 1) 'A parent where others can create folders is refused: they could plant the task folder first.'
    $inheritedFromProgramData = & $newAcl 'S-1-5-32-544' @('S-1-5-18=FullControl', 'S-1-5-32-544=FullControl', 'S-1-5-32-545=CreateDirectories, AppendData')
    Assert-HealthChecker (@(& $module { param($a) Get-HealthCheckerUnsafeAclEntry -Acl $a } $inheritedFromProgramData).Count -eq 1) 'A parent with the ProgramData rights for Users is refused for the work folder.'

    # CIM over DCOM: the session options are really built (no mock), from the code of both scripts.
    foreach ($source in @(@{ Name = 'collector'; Ast = $ast }, @{ Name = 'installer'; Ast = $installerAst })) {
        $optionFunction = $source.Ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-HealthCheckerDcomSessionOption' }, $true)
        Assert-HealthChecker ($null -ne $optionFunction) "The $($source.Name) builds its CIM options in New-HealthCheckerDcomSessionOption."
        $optionModule = New-Module -ScriptBlock ([scriptblock]::Create($optionFunction.Extent.Text))
        $option = & $optionModule { New-HealthCheckerDcomSessionOption }
        Assert-HealthChecker ($option -is [Microsoft.Management.Infrastructure.Options.DComSessionOptions] -and $option.PacketPrivacy -and $option.PacketIntegrity -and [string]$option.Impersonation -eq 'Impersonate') "The $($source.Name) CIM options are DCOM with packet privacy and integrity ($($option.GetType().Name))."
        $optionCalls = @($source.Ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-CimSessionOption' }, $true))
        $outside = @($optionCalls | Where-Object { $node = $_.Parent; while ($node -and $node -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $node = $node.Parent }; -not $node -or $node.Name -ne 'New-HealthCheckerDcomSessionOption' })
        Assert-HealthChecker ($optionCalls.Count -eq 1 -and $outside.Count -eq 0) "The $($source.Name) builds CIM options in one place only."
        $sessionCalls = @($source.Ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-CimSession' }, $true))
        Assert-HealthChecker ($sessionCalls.Count -eq 1 -and $sessionCalls[0].Extent.Text -match '-SessionOption \(New-HealthCheckerDcomSessionOption\)') "The $($source.Name) opens its CIM session with those options."
    }
    $collectorOption = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-HealthCheckerDcomSessionOption' }, $true).Extent.Text
    $installerOption = $installerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-HealthCheckerDcomSessionOption' }, $true).Extent.Text
    Assert-HealthChecker ($collectorOption -ceq $installerOption) 'The collector and the installer use the same CIM options.'

    # Windows PowerShell 5.1 only: under PowerShell 7 the ScheduledTasks module runs through Windows
    # compatibility and receives CIM sessions deserialized.
    $launcherText = [IO.File]::ReadAllText((Join-Path $SourceRoot 'SmartInventory/Launchers/OnPremises/Start-SmartM365-Exchange-HealthChecker-ScheduledTask-Installer.cmd'))
    Assert-HealthChecker ($launcherText -match '"%WINDOWS_POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER_PATH%" -Interactive' -and $launcherText -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' -and $launcherText -notmatch '(?i)pwsh') 'The installer launcher starts Windows PowerShell 5.1, never PowerShell 7.'
    Assert-HealthChecker ($installerText -match "\`$isWindowsPowerShell = \`$PSVersionTable\.PSEdition -eq 'Desktop'" -and $installerText -match 'if \(-not \$isAdministrator -or -not \$isWindowsPowerShell\) \{' -and $installerText -match 'FilePath\s+= \$windowsPowerShellPath') 'The installer restarts itself in Windows PowerShell 5.1 when started from PowerShell 7.'
    $relaunchIf = $installerAst.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '-not $isAdministrator -or -not $isWindowsPowerShell' }, $true)
    $firstCimUse = @($installerAst.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in @('New-HealthCheckerCimSession', 'Register-ScheduledTask', 'Get-Credential') }, $true) | Sort-Object { $_.Extent.StartOffset })[0]
    Assert-HealthChecker ($null -ne $relaunchIf -and $relaunchIf.Extent.EndOffset -lt $firstCimUse.Extent.StartOffset -and $relaunchIf.Clauses[0].Item2.Extent.Text -match 'exit \$smartM365RunGuard\.ExitCode') 'The restart happens before any credential or CIM use and ends the PowerShell 7 instance.'
    $shellCheck = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-ExchangeShell' }, $true)
    $shellCall = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Test-ExchangeShell' }, $true))
    Assert-HealthChecker ($shellCheck.Extent.Text -match "PSEdition -ne 'Desktop'" -and $shellCall.Count -eq 1 -and $shellCall[0].Extent.StartOffset -lt $hcCalls[0].Extent.StartOffset) 'The collector refuses PowerShell 7 before Health Checker can start.'
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $shellModule = New-Module -ScriptBlock ([scriptblock]::Create($shellCheck.Extent.Text))
        $refusedCore = $false
        try { & $shellModule { Test-ExchangeShell } } catch { $refusedCore = $_.Exception.Message -match 'require Windows PowerShell 5\.1' }
        Assert-HealthChecker $refusedCore 'Run under PowerShell 7, the collector stops with the Windows PowerShell 5.1 requirement.'
    }

    # Server names: powershell -File passes 'A,B' as one string.
    $listFunction = $installerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertTo-HealthCheckerComputerList' }, $true).Extent.Text
    $listModule = New-Module -ScriptBlock ([scriptblock]::Create($listFunction))
    Assert-HealthChecker ((@(& $listModule { ConvertTo-HealthCheckerComputerList -Value @('EXSRV01,EXSRV02,EXSRV03') }) -join '|') -eq 'EXSRV01|EXSRV02|EXSRV03') 'A comma-separated single value gives one name per server.'
    Assert-HealthChecker ((@(& $listModule { ConvertTo-HealthCheckerComputerList -Value @(' EXSRV01 ; EXSRV02,,', 'EXSRV01', 'EXSRV03') }) -join '|') -eq 'EXSRV01|EXSRV02|EXSRV03') 'Semicolons, spaces, empty entries and duplicates are handled.'
    Assert-HealthChecker (@(& $listModule { ConvertTo-HealthCheckerComputerList -Value @() }).Count -eq 0 -and @(& $listModule { ConvertTo-HealthCheckerComputerList -Value $null }).Count -eq 0) 'No value gives no server.'
    Assert-HealthChecker ($installerText -match '\$ComputerName = @\(ConvertTo-HealthCheckerComputerList -Value \$ComputerName\)' -and $installerText -match '\$ComputerName = @\(ConvertTo-HealthCheckerComputerList -Value \$candidateComputers\)' -and $installerText -notmatch "must not contain quotes, commas") 'Parameter and interactive names go through the same split, commas are allowed.'

    # Every failure is counted, so the completion banner reports Failed with its error count.
    $errorFunction = $installerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Write-InstallerError' }, $true).Extent.Text
    $errorModule = New-Module -ScriptBlock ([scriptblock]::Create($errorFunction + "`nfunction Write-InstallerMessage { param(`$Message) `$script:Last = `$Message }"))
    $global:SmartM365ErrorCount = 0
    & $errorModule { Write-InstallerError 'EXSRV01: installation failed: x' }
    Assert-HealthChecker ($global:SmartM365ErrorCount -eq 1 -and (& $errorModule { $script:Last }) -eq 'ERROR: EXSRV01: installation failed: x') 'An installer error is written and counted for the banner.'
    Remove-Variable -Name SmartM365ErrorCount -Scope Global -ErrorAction SilentlyContinue
    Assert-HealthChecker ($installerText -match 'Write-InstallerError \("\{0\}: installation failed' -and $installerText -match 'Write-InstallerError \("\{0\}: uninstall failed' -and $installerText -match 'Write-InstallerError \$_\.Exception\.Message' -and $installerText -match 'first run ended with') 'Validation errors, failed servers and failed first runs are counted.'
    Assert-HealthChecker ($installerText -notmatch "InstallerCompletionStatus = 'CompletedWithWarnings' \}" -and ([regex]::Matches($installerText, "if \(\`$failedServers -gt 0\) \{ \`$script:InstallerCompletionStatus = 'Failed' \}")).Count -eq 2) 'A failed server makes the installation Failed.'

    # Installer preflight: the service account must be a local administrator of each server.
    $preflightNames = @('ConvertFrom-HealthCheckerSidBytes', 'Test-HealthCheckerAdministratorMembership')
    $preflightModule = New-Module -ScriptBlock ([scriptblock]::Create((@($preflightNames | ForEach-Object { $n = $_; $installerAst.Find({ param($x) $x -is [Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true).Extent.Text }) -join "`n")))
    $sampleSid = New-Object Security.Principal.SecurityIdentifier('S-1-5-21-1111-2222-3333-1104')
    $sampleBytes = New-Object byte[] $sampleSid.BinaryLength
    $sampleSid.GetBinaryForm($sampleBytes, 0)
    Assert-HealthChecker ((& $preflightModule { param($b) ConvertFrom-HealthCheckerSidBytes -Bytes $b } $sampleBytes) -eq 'S-1-5-21-1111-2222-3333-1104') 'tokenGroups bytes are read as SIDs.'
    $accountSids = @('S-1-5-21-1111-2222-3333-1104', 'S-1-5-21-1111-2222-3333-513', 'S-1-5-21-1111-2222-3333-7001')
    $membership = { param($Sids, $Members) & $preflightModule { param($s, $m) Test-HealthCheckerAdministratorMembership -AccountSids $s -Members $m } $Sids $Members }
    $domainAdmins = [pscustomobject]@{ Sid = 'S-1-5-21-1111-2222-3333-512'; IsLocalGroup = $false }
    $localAdmin = [pscustomobject]@{ Sid = 'S-1-5-21-9999-500'; IsLocalGroup = $false }
    Assert-HealthChecker ((& $membership $accountSids @($localAdmin, [pscustomobject]@{ Sid = 'S-1-5-21-1111-2222-3333-1104'; IsLocalGroup = $false })) -eq 'Member') 'A directly added service account is a local administrator.'
    Assert-HealthChecker ((& $membership $accountSids @($domainAdmins, [pscustomobject]@{ Sid = 'S-1-5-21-1111-2222-3333-7001'; IsLocalGroup = $false })) -eq 'Member') 'Membership through a (nested) domain group is recognized.'
    Assert-HealthChecker ((& $membership $accountSids @($domainAdmins, $localAdmin)) -eq 'NotMember') 'An account in none of the groups is proven not to be an administrator.'
    Assert-HealthChecker ((& $membership $accountSids @($localAdmin, [pscustomobject]@{ Sid = 'S-1-5-21-9999-1010'; IsLocalGroup = $true })) -eq 'Unknown') 'A local group inside Administrators leaves the result undetermined.'
    Assert-HealthChecker ((& $membership $null @($localAdmin)) -eq 'Unknown') 'Without the account groups the result is undetermined.'
    $preflightIf = $installerAst.Find({ param($x) $x -is [Management.Automation.Language.IfStatementAst] -and $x.Clauses[0].Item1.Extent.Text -eq "`$membership -eq 'NotMember'" }, $true)
    $firstFolder = $installerAst.Find({ param($x) $x -is [Management.Automation.Language.CommandAst] -and $x.GetCommandName() -eq 'Assert-HealthCheckerParentFolder' }, $true)
    $register = $installerAst.Find({ param($x) $x -is [Management.Automation.Language.CommandAst] -and $x.GetCommandName() -eq 'Register-ScheduledTask' }, $true)
    Assert-HealthChecker ($null -ne $preflightIf -and $preflightIf.Extent.EndOffset -lt $firstFolder.Extent.StartOffset -and $preflightIf.Extent.EndOffset -lt $register.Extent.StartOffset -and $preflightIf.Clauses[0].Item2.Extent.Text -match 'throw' -and $preflightIf.Clauses[0].Item2.Extent.Text -match 'is not a local administrator on') 'A server where the account is not an administrator is refused before any folder or task change.'
    Assert-HealthChecker ($installerText -match "Write-InstallerWarning \(""\{0\}: local administrator membership of \{1\} could not be confirmed" -and $installerText -match 'Write-InstallerError \("\{0\}: installation failed') 'An undetermined membership is a warning; a refused server is an error.'
    Assert-HealthChecker (@($installerAst.FindAll({ param($x) $x -is [Management.Automation.Language.CommandAst] -and $x.GetCommandName() -eq 'Get-Credential' }, $true)).Count -eq 1 -and $installerText -match "DirectoryEntry\(\('LDAP://<SID=\{0\}>' -f \`$AccountSid\)\)" -and $installerText -match "RefreshCache\(\[string\[\]\]@\('tokenGroups'\)\)") 'Active Directory is read once, read-only, without any credential prompt.'

    # Module path inherited from PowerShell 7: the Windows PowerShell child must still load its own modules.
    $modulePathTexts = foreach ($scriptAst in @($ast, $installerAst, $runnerAst)) { $scriptAst.Find({ param($x) $x -is [Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq 'Repair-HealthCheckerModulePath' }, $true).Extent.Text }
    Assert-HealthChecker (@($modulePathTexts | Select-Object -Unique).Count -eq 1) 'The collector, installer and runner repair the module path the same way.'
    $collectorBody = @($ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State') })[0].Body.Statements
    $installerBody = @($installerAst.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State') })[0].Body.Statements
    Assert-HealthChecker ($collectorBody[0].Name -eq 'Repair-HealthCheckerModulePath' -and $collectorBody[1].Extent.Text -eq 'Repair-HealthCheckerModulePath' -and $installerBody[0].Name -eq 'Repair-HealthCheckerModulePath' -and $installerBody[1].Extent.Text -eq 'Repair-HealthCheckerModulePath') 'The collector and the installer repair the module path before anything else.'
    $runnerEntry = $runnerAst.Find({ param($x) $x -is [Management.Automation.Language.IfStatementAst] -and $x.Clauses[0].Item1.Extent.Text -eq "`$MyInvocation.InvocationName -ne '.'" }, $true)
    Assert-HealthChecker ($runnerEntry.Clauses[0].Item2.Statements[0].Extent.Text -eq 'Repair-HealthCheckerModulePath') 'The runner repairs the module path when it starts.'
    $pwshModules = Join-Path $env:ProgramFiles 'PowerShell\7\Modules'
    $modulePathProbe = Join-Path $tempRoot 'module-path-probe.ps1'
    [IO.File]::WriteAllText($modulePathProbe, (@(
        'param([string]$Mode, [string]$Folder)',
        $modulePathTexts[0],
        "if (`$Mode -eq 'Repaired') { Repair-HealthCheckerModulePath }",
        "try { `$acl = Get-Acl -LiteralPath `$Folder; `$sig = Get-AuthenticodeSignature -FilePath `$PSCommandPath; 'LOADED ' + `$acl.GetType().Name + ' ' + (Get-Module Microsoft.PowerShell.Security).ModuleBase + ' FIRST=' + ((`$env:PSModulePath -split ';')[0]) } catch { 'FAILED ' + `$_.Exception.Message }"
    ) -join "`r`n"), (New-Object Text.UTF8Encoding($false)))
    $savedModulePath = $env:PSModulePath
    try {
        $env:PSModulePath = $pwshModules + ';' + (Join-Path $env:USERPROFILE 'Documents\PowerShell\Modules') + ';' + $savedModulePath
        $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        # Start-Process like the orchestrator: pwsh's call operator would clean the module path of the child itself.
        $runProbe = {
            param([string]$Mode)
            $outFile = Join-Path $tempRoot "module-path-$Mode.out"
            [void](Start-Process -FilePath $windowsPowerShell -ArgumentList @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $modulePathProbe), '-Mode', $Mode, '-Folder', ('"{0}"' -f $tempRoot)) -RedirectStandardOutput $outFile -RedirectStandardError ($outFile + '.err') -WindowStyle Hidden -Wait -PassThru)
            ((Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath ($outFile + '.err') -Raw)).Trim()
        }
        $plainOutput = & $runProbe 'Plain'
        $repairedOutput = & $runProbe 'Repaired'
    }
    finally { $env:PSModulePath = $savedModulePath }
    if (Test-Path -LiteralPath (Join-Path $pwshModules 'Microsoft.PowerShell.Security')) {
        Assert-HealthChecker ($plainOutput -match '^FAILED') "With the PowerShell 7 module path first, Windows PowerShell fails to load its security module (reproduced): $plainOutput"
    }
    Assert-HealthChecker ($repairedOutput -match '^LOADED DirectorySecurity' -and $repairedOutput -match ([regex]::Escape((Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\Modules'))) -and $repairedOutput -notmatch 'PowerShell\\7') "After the repair, Get-Acl and Get-AuthenticodeSignature load from Windows PowerShell: $repairedOutput"

    # Interrupted run (Ctrl+C): the collector's own end-of-run code, run for real in a child process.
    $collectorTry = $ast.Find({ param($x) $x -is [Management.Automation.Language.TryStatementAst] -and $x.Finally -and $x.Finally.Extent.Text.Contains('$bannerStatus') }, $true)
    Assert-HealthChecker ($null -ne $collectorTry -and $collectorTry.CatchClauses[0].Body.Statements[0].Extent.Text -match 'PipelineStoppedException.*\{ throw \}') 'An interruption is rethrown before any receipt or completion is written.'
    $probePath = Join-Path $tempRoot 'interrupted-probe.ps1'
    $probeTemplate = @'
param([string]$Root, [string]$Mode)
. (Join-Path $Root 'Config/SmartM365-TenantContext.ps1')
. (Join-Path $Root 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$script:ServersAndStorageWarningCount = 0
$script:ServersAndStorageErrorCount = 0
function Write-Log { param($Message, $Level = 'INFO') if ($Level -eq 'ERROR') { $script:ServersAndStorageErrorCount++ }; if ($Level -eq 'WARN') { $script:ServersAndStorageWarningCount++ }; Write-Output ("[{0}] {1}" -f $Level, $Message) }
$ScriptName = 'Probe'
$StartTime = Get-Date
$script:CompletionStatus = 'Auto'
$smartM365RunGuard = Start-SmartM365InventoryRun -ScriptPath (Join-Path $Root 'probe.ps1') -RequireReturn
try {
    try {
        if ($Mode -eq 'Interrupted') { throw [Management.Automation.PipelineStoppedException]::new() }
        $script:CompletionStatus = 'Success'
        $smartM365RunGuard.ReachedTerminal = $true
    }
    catch { if ($_.Exception -is [Management.Automation.PipelineStoppedException]) { throw } }
    FINALLY_BODY
    $smartM365RunGuard.ReachedTerminal = $true
    $smartM365RunGuard.ReturnedNormally = $true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException]) { $smartM365RunGuard.ErrorRecord = $_ }
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if ($smartM365RunGuard.ExitCode -eq 130) { exit 130 }
}
'@
    [IO.File]::WriteAllText($probePath, $probeTemplate.Replace('FINALLY_BODY', 'finally ' + $collectorTry.Finally.Extent.Text), (New-Object Text.UTF8Encoding($false)))
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $interruptedOutput = (& $windowsPowerShell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $probePath -Root $SourceRoot -Mode Interrupted 2>&1 | Out-String)
    $interruptedExit = $LASTEXITCODE
    Assert-HealthChecker ($interruptedExit -eq 130 -and $interruptedOutput -cmatch 'Status\s*:\s*CANCELLED' -and $interruptedOutput -cnotmatch 'SUCCESS' -and $interruptedOutput -match '\[WARN\] Run interrupted before completion' -and $interruptedOutput -match 'Warnings\s*:\s*1' -and $interruptedOutput -match 'Errors\s*:\s*0') "An interrupted run is CANCELLED, never SUCCESS, and exits with 130 (exit $interruptedExit): $interruptedOutput"
    $completedOutput = (& $windowsPowerShell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $probePath -Root $SourceRoot -Mode Completed 2>&1 | Out-String)
    $completedExit = $LASTEXITCODE
    Assert-HealthChecker ($completedExit -eq 0 -and $completedOutput -cmatch 'Status\s*:\s*SUCCESS' -and $completedOutput -notmatch 'interrupted') "A completed run keeps its SUCCESS banner (exit $completedExit)."
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ("PASS: {0} Health Checker checks; PowerShell {1}." -f $script:Checks, $PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAYxrpvVAn3LXwN
# dAwhPumOjxo4KLHRnMcZ6AgP6NNdEKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC1gLjE3JAye8T60W8sQJ2625mitc/0n7ZPcgxvI1953MA0GCSqG
# SIb3DQEBAQUABIIBgKpzf2LLD6IKGFmzWVZwQzvlPUfhCsjK07CpXm+fUd2m/M+K
# JEb+gdJcDqW1VEhVTUvvttAyd6Eyv8KQ9MNZPFe3DR1iM8J2hW/yhwXPBkL+87Ix
# uCWSzFkAXstZ4k7YDXHNcbmWMaAnD4jNyCetz703ljeVI0CFXQpISHwOhnfzvPNI
# zaBX33EpRI624h4LV8687R/S+HWzfJ8FEIkBpIIdA3ysdZQgR2fu5qG2BD8XsbYM
# 9XrHjiQNnzu2LOOB7IZ+oVQwNVl/96u1rXuvEhFTGNVZk1kUrx9uK7Nk+hzBYNYM
# laLyDU8RnkncJaa+sm2Dpm3rhu7z7Gfs2pevTSgnQ0zAD/jSa5KICdCogfY5euXa
# n0zLbZzm3E3py+HQp/rQTcsaFbjFTWarK439dJzs3JdLnG3cWMEmDBV7qbw2d9Zi
# G+inW4EXRti/R2KhDgPvf4Vrn28NYgGngp1SHfX5Fvi6b+rAQ/AcEryl671n9RUP
# N6htIcgD6hAzVZAbYaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjQxMzM3
# MzZaMC8GCSqGSIb3DQEJBDEiBCDKp+CAqVTYVZmo4TOhsdB/pAqnSLE4R0eT8O24
# ij1nLjANBgkqhkiG9w0BAQEFAASCAgCpqcOlOhhbmzYkZ3zMkdf43ei2wanGtn8X
# /P4NPhSM84ybXQI4QRJ/b/IxPh82SWdkYE1hekluJ+KysoWmwn8BNUUJ1RhjoIIO
# yreEso+2IR+IXAin+mZD4ZqadzA+dh5Fm2+tT1vA/y5aJUafdsykQUZEQNE8HuqA
# Xt5eorRhiTDpYd21iP0DJAIksi7qm2Y2OAZgDkGsstlWr8pxM8f56kHwezKgCcHN
# fLgOU+EsXeQ1j3vJ45YJrApNEObHQSdJoC3qUvwiq9jEWvNSJv8f8k17lUKC39V9
# GtXTBydTjoDQdBbz6MFMt6yh99ezZlarTz7V7+78ckgUiDE7wKsEDOKbf8o0U/hQ
# dMQye0TS8a6s89ERFxait7DSM9SBb73/PxZXeG0OcWQwiBeEJ/wMczbZvamGxelL
# /griHfKpXKz0a2JJJTZFVdq3YjAyO+ooj2/s55O0Tb17rPI6C1mjwE32gKGJTwmq
# bKVTflq3AdR9tgOP8+BpiJuFQlqBn6QAaOLW+pMe9v1n/BUY31huWbkcg5tdeGWU
# sc26BG8vWKFgNZiafpYQAlaPrxUV2OBgPiybTOSSSjVkliS7dM49uglfaesv76wh
# Z1hm8wlgZW/ficLYDIngRWJ7XvXy9OvB8PGS7eIeiJYLdy2OIwEi65TUHinh6Uo5
# XEc9mJTS3w==
# SIG # End signature block
