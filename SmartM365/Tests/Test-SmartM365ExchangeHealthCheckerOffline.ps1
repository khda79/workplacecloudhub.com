<#
.SYNOPSIS
Offline tests of the Health Checker integration of the Exchange on-premises infrastructure collector.

.DESCRIPTION
Extracts the Health Checker functions from the collector without running it. Microsoft's Health
Checker is never executed: process runs use a synthetic stand-in script. Covers:
- reading a synthetic Health Checker XML exported like Health Checker does (Export-Clixml -Depth 2);
- masking and scrubbing of sensitive values, status mapping, vulnerability matching, summary facts;
- integrity check of the vendored file (valid, SHA256 mismatch, non-Microsoft signer);
- child process run: arguments, completed, not elevated, timeout with process tree stop;
- per-run orchestration: integrity failure, not elevated, time budget, raw output cleanup;
- option disabled: no Health Checker call outside -IncludeHealthChecker and no report section.

.VERSION
1.0.0
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
    'Get-HealthCheckerLineValue', 'ConvertFrom-HealthCheckerXml', 'ConvertTo-HealthCheckerSummaryRow', 'Remove-HealthCheckerExpiredWorkFolder',
    'Invoke-ExchangeHealthCheckerInventory', 'ConvertTo-ServersAndStorageHealthCheckerSectionHtml', 'ConvertTo-ServersAndStorageDatabaseSectionHtml',
    'New-HtmlExecutiveSummary')
$definitions = foreach ($name in $functionNames) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if ($null -eq $node) { throw "Missing function: $name" }
    $node.Extent.Text
}
$constants = @{}
foreach ($constantName in @('HealthCheckerVersion', 'HealthCheckerExpectedSha256', 'HealthCheckerRelativePath')) {
    $assignment = $ast.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq "`$$constantName" }, $true)
    if ($null -eq $assignment) { throw "Missing collector constant: $constantName" }
    $constants[$constantName] = $assignment.Right.Extent.Text.Trim("'", '"')
}

$module = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
& $module {
    Set-StrictMode -Version Latest
    $script:LogLines = New-Object 'System.Collections.Generic.List[string]'
    $script:Config = @{}
    function script:Write-Log { param($Message, $Level = 'INFO') $script:LogLines.Add(("[{0}] {1}" -f $Level, $Message)) | Out-Null }
    function script:Get-SmartM365EffectiveConfigValue { param($Name, $DefaultValue) if ($script:Config.ContainsKey($Name)) { return $script:Config[$Name] } return $DefaultValue }
}

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
    Assert-HealthChecker ($settings.TimeoutMinutesPerServer -eq 30 -and $settings.TotalBudgetMinutes -eq 1320 -and $settings.KeepFailedRawOutputDays -eq 7 -and $settings.WorkFolderPath -like '*\SmartM365\HealthChecker') 'Default settings.'
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
Copy-Item -LiteralPath $env:SMARTM365_FAKE_HC_FIXTURE -Destination (Join-Path $OutputFilePath ("HealthChecker-{0}-20261021060000.xml" -f $Server[0].ToUpper()))
'@, [Text.UTF8Encoding]::new($false))
    $env:SMARTM365_FAKE_HC_FIXTURE = $fixture
    try {
        $env:SMARTM365_FAKE_HC_MODE = 'ok'
        $okFolder = Join-Path $tempRoot 'run-ok'
        $ok = & $module { param($s, $o) Invoke-HealthCheckerProcess -ScriptPath $s -ServerName 'EXSRV01' -OutputFolder $o -TimeoutSeconds 120 } $fakeScript $okFolder
        Assert-HealthChecker ($ok.Status -eq 'Completed' -and $ok.XmlPath -like '*HealthChecker-EXSRV01-*.xml') "A completed run returns its XML ($($ok.Status) $($ok.Message))."
        $fakeArgs = Get-Content -LiteralPath (Join-Path $okFolder 'fake-args.txt') -Raw
        Assert-HealthChecker ($fakeArgs -match 'Server=EXSRV01;SkipVersionCheck=True') "Health Checker gets -Server and -SkipVersionCheck ($fakeArgs)."

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

    # 5. Orchestration with a stubbed runner.
    $servers = @([pscustomobject]@{ Name = 'EXSRV01' }, [pscustomobject]@{ Name = 'EXSRV02' })
    & $module {
        param($FixturePath)
        $script:FixturePath = $FixturePath
        $script:ProcessCalls = New-Object 'System.Collections.Generic.List[object]'
        $script:Elevated = $true
        $script:StubDelaySeconds = 0
        function script:Test-HealthCheckerElevation { return $script:Elevated }
        function script:Invoke-HealthCheckerProcess {
            param($ScriptPath, $ServerName, $OutputFolder, $TimeoutSeconds)
            $script:ProcessCalls.Add([pscustomobject]@{ ScriptPath = $ScriptPath; ServerName = $ServerName; TimeoutSeconds = $TimeoutSeconds; ScriptExists = (Test-Path -LiteralPath $ScriptPath) }) | Out-Null
            Ensure-Directory -Path $OutputFolder
            if ($script:StubDelaySeconds -gt 0) { Start-Sleep -Seconds $script:StubDelaySeconds }
            if ($ServerName -eq 'EXSRV02') { return [pscustomobject]@{ ServerName = $ServerName; Status = 'Timeout'; Message = 'stopped'; ExitCode = $null; DurationSeconds = 5; XmlPath = $null } }
            $xml = Join-Path $OutputFolder 'HealthChecker-EXSRV01-1.xml'
            Copy-Item -LiteralPath $script:FixturePath -Destination $xml
            return [pscustomobject]@{ ServerName = $ServerName; Status = 'Completed'; Message = $null; ExitCode = 0; DurationSeconds = 12; XmlPath = $xml }
        }
    } $fixture
    $workRoot = Join-Path $tempRoot 'work'
    $runSettings = [pscustomobject]@{ TimeoutMinutesPerServer = 30; TotalBudgetMinutes = 1320; KeepFailedRawOutputDays = 7; WorkFolderPath = $workRoot }
    $invoke = {
        param($Servers, $Source, $Hash, $Version, $Settings)
        Invoke-ExchangeHealthCheckerInventory -ExchangeServers $Servers -SourceScriptPath $Source -ExpectedSha256 $Hash -HealthCheckerVersion $Version -RunId '20261021-060000' -Settings $Settings
    }
    $result = & $module $invoke $servers $vendoredPath $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $runSettings
    $calls = @(& $module { $script:ProcessCalls })
    Assert-HealthChecker ($calls.Count -eq 2 -and $calls[0].ScriptPath -like "$workRoot*" -and $calls[0].ScriptExists) 'Each server runs the verified copy in the local work folder.'
    Assert-HealthChecker ($calls[0].TimeoutSeconds -eq 1800) 'The per-server limit is applied.'
    $summary = @($result.Summary)
    Assert-HealthChecker ($summary.Count -eq 2 -and $summary[0].RunStatus -eq 'Completed' -and $summary[1].RunStatus -eq 'Timeout') 'Every server has a summary row with its run status.'
    Assert-HealthChecker ($summary[0].ServersProcessed -eq 1 -and $summary[0].ServersTotal -eq 2 -and $summary[1].ServersProcessed -eq 1) 'Processed and total server counts are on every row.'
    Assert-HealthChecker (@($result.Findings).Count -eq 17) 'Findings of the completed server are returned.'
    $runFolder = Join-Path $workRoot '20261021-060000'
    Assert-HealthChecker (-not (Test-Path -LiteralPath (Join-Path $runFolder 'EXSRV01'))) 'Raw output of a completed server is deleted.'
    Assert-HealthChecker (Test-Path -LiteralPath (Join-Path $runFolder 'EXSRV02')) 'Raw output of a failed server is kept for diagnosis.'
    Assert-HealthChecker (-not (Test-Path -LiteralPath (Join-Path $runFolder 'HealthChecker.ps1'))) 'The run copy of the script is removed.'
    $logText = (& $module { $script:LogLines }) -join "`n"
    Assert-HealthChecker ($logText -match '\[WARN\] EXSRV02: Health Checker Timeout' -and $logText -match 'processed 1 of 2') 'Failures are logged as warnings with the processed count.'

    & $module { $script:ProcessCalls.Clear(); $script:LogLines.Clear() }
    $integrityResult = & $module $invoke $servers $vendoredPath ('0' * 64) $constants.HealthCheckerVersion $runSettings
    Assert-HealthChecker (@(& $module { $script:ProcessCalls }).Count -eq 0) 'Nothing runs when the integrity check fails.'
    Assert-HealthChecker (@($integrityResult.Summary | Where-Object { $_.RunStatus -eq 'IntegrityFailed' -and $_.RunMessage -match 'SHA256 mismatch' }).Count -eq 2 -and @($integrityResult.Findings).Count -eq 0) 'Every server reports IntegrityFailed.'
    Assert-HealthChecker (((& $module { $script:LogLines }) -join "`n") -match '\[WARN\] Health Checker was not run') 'The integrity failure is a warning.'

    & $module { $script:ProcessCalls.Clear(); $script:Elevated = $false }
    $notElevatedResult = & $module $invoke $servers $vendoredPath $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $runSettings
    Assert-HealthChecker (@(& $module { $script:ProcessCalls }).Count -eq 0 -and @($notElevatedResult.Summary | Where-Object { $_.RunStatus -eq 'NotElevated' }).Count -eq 2) 'Nothing runs without elevation and every server reports NotElevated.'

    & $module { $script:ProcessCalls.Clear(); $script:Elevated = $true; $script:StubDelaySeconds = 3 }
    $budgetSettings = [pscustomobject]@{ TimeoutMinutesPerServer = 30; TotalBudgetMinutes = 1.04; KeepFailedRawOutputDays = 7; WorkFolderPath = (Join-Path $tempRoot 'work-budget') }
    $budgetResult = & $module $invoke $servers $vendoredPath $constants.HealthCheckerExpectedSha256 $constants.HealthCheckerVersion $budgetSettings
    $budgetCalls = @(& $module { $script:ProcessCalls })
    Assert-HealthChecker ($budgetCalls.Count -eq 1 -and $budgetCalls[0].TimeoutSeconds -le 63) "The limit never exceeds the remaining budget ($($budgetCalls[0].TimeoutSeconds)s)."
    Assert-HealthChecker (@($budgetResult.Summary)[1].RunStatus -eq 'NotRun' -and @($budgetResult.Summary)[1].RunMessage -match 'budget') 'Servers left after the budget is used up are NotRun.'
    & $module { $script:StubDelaySeconds = 0 }

    # Expired work folders: only run folders older than the retention are removed.
    $purgeRoot = Join-Path $tempRoot 'purge'
    foreach ($name in @('20261001-010101', '20261020-010101', 'Keep-Me')) { New-Item -ItemType Directory -Path (Join-Path $purgeRoot $name) -Force | Out-Null }
    (Get-Item (Join-Path $purgeRoot '20261001-010101')).LastWriteTime = (Get-Date).AddDays(-20)
    (Get-Item (Join-Path $purgeRoot 'Keep-Me')).LastWriteTime = (Get-Date).AddDays(-20)
    & $module { param($r) Remove-HealthCheckerExpiredWorkFolder -WorkRoot $r -KeepDays 7 } $purgeRoot
    Assert-HealthChecker (-not (Test-Path (Join-Path $purgeRoot '20261001-010101')) -and (Test-Path (Join-Path $purgeRoot '20261020-010101')) -and (Test-Path (Join-Path $purgeRoot 'Keep-Me'))) 'Only expired run folders are purged.'

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
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ("PASS: {0} Health Checker checks; PowerShell {1}." -f $script:Checks, $PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB6hrDEgXso2wvX
# wc9RlTITwvxk6Cgw9vmB0ApsoaPhAKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIk///gX/4F6W69dHlm5k2WO1a7gjm5V/j+3inmTtcPTMA0GCSqG
# SIb3DQEBAQUABIIBgDSa2YHPzmAdhvCEWAK0H5nrybJvhOjqjZXosea9Y7rPTjoD
# Yd0+mmgNl2ufEl3s4yJDzkIUS8dPajP546mHrkgr2E+OilyhxgS7FY2Iqi8tn8jK
# sNrIfNV0FBGS4RXazuPbgfUJ+FrN+VZpCRCTT7VfiyHgF5KQC208O/4lvagSoJQj
# Tr+S72EZORHT9tAnSAOSAO6jr3d+pabJVBdBn0YA2rj+amgy1a+0Nn8F89cHddr+
# C5YWzr0cjhZNZY5SbEVLnH5UCUt2Bd25+w1h8KvM4bWdM+dj4qF6FJLmbrRizzgs
# zgm/KwOwL9aTfOlKldVsOBw+l/v8HF+xSvWWjOZd5U+tiAQjfjtpKMLbGRlGxvAK
# PrYbrSAbbMgw5WHpXIJVItOMwjImY8atXOOpK7VpP7/fMhsuVPWy5BIF4DxfxdcQ
# 1nMfM/v+JA7ncV5jAJU+dq7YmNPRLCWNIwznr3xQqAO7jlR66C44JpjtbqHmXPBQ
# cUM6IyXxA1X6dFCZgaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExNTU1
# MjJaMC8GCSqGSIb3DQEJBDEiBCAoBtbL/7i3wFzC25MsrTQf7X3J3uClmUhIG/E/
# vJ5C6DANBgkqhkiG9w0BAQEFAASCAgAU5JgXlYIcjyCBo6v2ckFmYsgDfcCnc6Cx
# 4yWT4DW3TOhvhXP+psqYC1F3HUioG7Y7nwHM0SH598y1Cvam20G1nGIA44/xwaJn
# QNCNZ2zJ3VyRA4HtGHnAo0hd/4e5sfJg7uxRV2LbUOl1tcHIscMrlJyjm/+avl8x
# whC9E2PsJyE3slscmskO76q7HGrH9eXdEatSkvjzaQ6hYkfun98WgZIhFR5BywGF
# MAumkc6P/GdHNwqx9lYuizToyFnddq54NbZggkQr7Y3imTuerzS+MVSJLLyA84XS
# nyWR6aCLbcQProQgz/U2uNLQhtvsmRGHzG4W1Lef9aRbXUINsCO/c02V29xWJSvu
# OC0LT0s1NM1FJnCpam9b9F/Rwd0fbpbK0VPGjNjMxbwOPQ96hrAAqOYTcEAXHC0a
# nupUNIJgoP0NVyROWnu95Anh+CJj0LRsRMMYh14Y3Ply4UavbQj54N6f78+lh5pJ
# D/xaxR1f1D3ZEs4r4xrygbyivcc8QsiuMB2lNrFps+znxu4MTFZBevvOf3Al2P3n
# qpgtPd9GDdJrQ2eqPYU77au31JHMknGvj1VAY7PpSDRm8Knvzmjj2cPCZSlD9Cum
# NtCA5kEHHasBYryhBVnymDfds/C163N/aRIyaRY36U9+NybgkKU+9zzRJKkDD+4J
# JJMeQjWZdg==
# SIG # End signature block
