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
- option disabled: no Health Checker call outside -IncludeHealthChecker and no report section.

.VERSION
1.1.0
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
$functionNames = @('Ensure-Directory', 'Format-HtmlValue', 'Get-HealthCheckerCsvColumn', 'Get-HealthCheckerMaskingRule', 'Get-HealthCheckerSetting',
    'Test-HealthCheckerScriptIntegrity', 'Test-HealthCheckerElevation', 'Invoke-HealthCheckerProcess', 'ConvertTo-HealthCheckerSafeValue',
    'Get-HealthCheckerLineValue', 'ConvertFrom-HealthCheckerXml', 'ConvertTo-HealthCheckerSummaryRow', 'Remove-HealthCheckerPreviousWorkFolder',
    'Get-HealthCheckerServerComputerName', 'ConvertTo-HealthCheckerAdminSharePath', 'ConvertTo-HealthCheckerRemoteMessage', 'ConvertTo-HealthCheckerUtcTime',
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
            $previous = [pscustomobject]@{ RunToken = 'old'; State = 'Completed'; Message = $null; StartedAtUtc = $now.AddHours(-3).ToString('o'); FinishedAtUtc = $now.AddHours(-2).ToString('o'); DurationSeconds = 150; XmlFileName = 'Raw\HealthChecker-EXSRV02-1.xml' }
            $run = [pscustomobject]@{ ServerName = $ServerName; Computer = $Computer; Session = $null; OutputSharePath = $script:RemoteShare; PreviousStatus = $previous; WaitToken = $null; StartedAt = (Get-Date); Done = $false; Status = $null; Message = $null; RemoteStatus = $null }
            Copy-Item -LiteralPath $script:FixturePath -Destination (Join-Path $script:RemoteShare 'Raw\HealthChecker-EXSRV02-1.xml') -Force
            switch ($script:RemoteMode) {
                'Completed' {
                    $current = [pscustomobject]@{ RunToken = 'new'; State = 'Completed'; Message = $null; StartedAtUtc = $now.ToString('o'); FinishedAtUtc = $now.ToString('o'); DurationSeconds = 140; XmlFileName = 'Raw\HealthChecker-EXSRV02-1.xml' }
                    [IO.File]::WriteAllText((Join-Path $script:RemoteShare 'status.json'), ($current | ConvertTo-Json))
                }
                'RunnerTimeout' {
                    $current = [pscustomobject]@{ RunToken = 'new'; State = 'Timeout'; Message = 'Health Checker exceeded the 30 minute limit on the server and was stopped.'; StartedAtUtc = $now.ToString('o'); FinishedAtUtc = $now.ToString('o'); DurationSeconds = 1800; XmlFileName = $null }
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
    Assert-HealthChecker (Test-Path -LiteralPath (Join-Path $remoteShare 'Raw\HealthChecker-EXSRV02-1.xml')) 'Files on the server are left to its own next run.'

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
    Assert-HealthChecker $guarded 'Health Checker only runs under -IncludeHealthChecker.'
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
    $fresh = [pscustomobject]@{ State = 'Completed'; XmlFileName = 'Raw\HealthChecker-X-1.xml'; FinishedAtUtc = $nowUtc.AddHours(-5).ToString('o') }
    $freshResult = & $module { param($r, $n) Resolve-HealthCheckerScheduledFallback -FailureStatus 'RemoteConnectionFailed' -FailureMessage 'unreachable.' -RemoteStatus $r -NowUtc $n -MaxResultAgeHours 26 } $fresh $nowUtc
    Assert-HealthChecker ($freshResult.Status -eq 'Completed' -and $freshResult.UseResult -and $freshResult.ResultAgeHours -eq 5) 'A recent scheduled result is used.'
    $fresh.FinishedAtUtc = $nowUtc.AddHours(-26.04).ToString('o')
    $limitResult = & $module { param($r, $n) Resolve-HealthCheckerScheduledFallback -FailureStatus 'RemoteConnectionFailed' -FailureMessage 'x' -RemoteStatus $r -NowUtc $n -MaxResultAgeHours 26 } $fresh $nowUtc
    Assert-HealthChecker ($limitResult.Status -eq 'Completed') 'A result at the age limit is still used.'
    $fresh.FinishedAtUtc = [datetime]$nowUtc.AddHours(-48)
    $oldResult = & $module { param($r, $n) Resolve-HealthCheckerScheduledFallback -FailureStatus 'TaskTriggerFailed' -FailureMessage 'x' -RemoteStatus $r -NowUtc $n -MaxResultAgeHours 26 } $fresh $nowUtc
    Assert-HealthChecker ($oldResult.Status -eq 'StaleResult' -and -not $oldResult.UseResult -and $oldResult.ResultAgeHours -eq 48) 'An old result (DateTime from PowerShell 7 JSON) is StaleResult.'
    foreach ($unusable in @($null, [pscustomobject]@{ State = 'NoOutput'; XmlFileName = $null; FinishedAtUtc = $nowUtc.ToString('o') }, [pscustomobject]@{ State = 'Completed'; XmlFileName = 'Raw\a.xml'; FinishedAtUtc = 'not a date' })) {
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
    $badName = [pscustomobject]@{ ServerName = 'EXSRV02'; Computer = 'EXSRV02'; Session = $null; OutputSharePath = $progressShare; PreviousStatus = $null; WaitToken = $null; StartedAt = (Get-Date); Done = $true; Status = $null; Message = $null; RemoteStatus = [pscustomobject]@{ State = 'Completed'; XmlFileName = '..\..\Windows\win.ini'; FinishedAtUtc = $nowUtc.ToString('o'); DurationSeconds = 1 } }
    $badOutcome = & $module { param($r, $f, $s) Complete-HealthCheckerRemoteRun -Run $r -LocalFolder $f -Settings $s } $badName (Join-Path $tempRoot 'bad-name') $runSettings
    Assert-HealthChecker ($badOutcome.Status -eq 'NoOutput' -and $null -eq $badOutcome.XmlPath) 'A result file name outside Raw\HealthChecker-*.xml is refused.'

    # Server files are copied only when they differ.
    $copySource = Join-Path $tempRoot 'copy-source.txt'
    $copyTarget = Join-Path $tempRoot 'copy-target.txt'
    [IO.File]::WriteAllText($copySource, 'v1')
    Assert-HealthChecker ((& $module { param($s, $d) Copy-HealthCheckerRemoteFile -SourcePath $s -DestinationPath $d } $copySource $copyTarget) -eq $true) 'A missing file is copied.'
    Assert-HealthChecker ((& $module { param($s, $d) Copy-HealthCheckerRemoteFile -SourcePath $s -DestinationPath $d } $copySource $copyTarget) -eq $false) 'An identical file is not copied again.'
    [IO.File]::WriteAllText($copyTarget, 'tampered')
    Assert-HealthChecker ((& $module { param($s, $d) Copy-HealthCheckerRemoteFile -SourcePath $s -DestinationPath $d } $copySource $copyTarget) -eq $true -and [IO.File]::ReadAllText($copyTarget) -eq 'v1') 'A file that differs is replaced.'

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
    Assert-HealthChecker (((@($runnerAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })) -join ',') -eq 'TimeoutMinutes') 'The runner has no -Server parameter.'
    Assert-HealthChecker ($runnerText -notmatch "'-Server'" -and $runnerText -notmatch 'SmartM365\.Core|Import-Module|InventoryLifecycle|TenantContext') 'The runner never passes -Server and is standalone (no repository module).'
    Assert-HealthChecker ($runnerText -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' -and $runnerText -match 'Global\\SmartM365-Exchange-HealthChecker-LocalRunner') 'The runner forces Windows PowerShell 5.1 and takes a machine-wide lock.'
    $runnerFunctions = foreach ($node in $runnerAst.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) { $node.Extent.Text }
    $runnerModule = New-Module -ScriptBlock ([scriptblock]::Create(($runnerFunctions -join "`n")))
    & $runnerModule {
        param($Version, $Hash)
        Set-StrictMode -Version Latest
        $script:ScriptVersion = '1.0.0'; $script:HealthCheckerVersion = $Version; $script:HealthCheckerExpectedSha256 = $Hash
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
        Assert-HealthChecker ($statusFile.Message -notmatch [regex]::Escape($runnerRoot) -and $statusFile.RunToken -match '^[0-9a-f-]{36}$' -and $statusFile.FinishedAtUtc -and $statusFile.RunnerVersion -eq '1.0.0') 'The status holds a run token, times and versions without local paths.'
        $statusProperties = @($statusFile.PSObject.Properties.Name) -join ','
        Assert-HealthChecker ($statusProperties -eq 'SchemaVersion,RunToken,State,Message,RunnerVersion,HealthCheckerVersion,StartedAtUtc,FinishedAtUtc,DurationSeconds,ExitCode,XmlFileName') "The status has no host or account field ($statusProperties)."

        # The real check is replaced to run the synthetic stand-in; Microsoft's script is never run here.
        & $runnerModule {
            function script:Test-HealthCheckerScriptIntegrity { param($Path, $ExpectedSha256) [pscustomobject]@{ IsValid = $true; Reason = $null } }
            function script:Test-LocalRunnerElevation { return $true }
        }
        $env:SMARTM365_FAKE_HC_MODE = 'ok'
        $okRun = & $runnerModule { param($r) Invoke-LocalHealthCheckerRun -RootFolder $r -TimeoutSeconds 120 } $runnerRoot
        Assert-HealthChecker ($okRun.State -eq 'Completed' -and $okRun.XmlFileName -match '^Raw\\HealthChecker-.+\.xml$' -and (Test-Path -LiteralPath (Join-Path (Join-Path $runnerRoot 'Output') $okRun.XmlFileName))) "A completed run records its XML relative to Output ($($okRun.XmlFileName))."
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
    Assert-HealthChecker ($installerText -match "-Protocol Dcom" -and $installerText -notmatch 'Invoke-Command|New-PSSession|Enter-PSSession|schtasks') 'The installer uses CIM over DCOM, never WinRM or schtasks.'
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
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ("PASS: {0} Health Checker checks; PowerShell {1}." -f $script:Checks, $PSVersionTable.PSVersion)
