#Requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param([bool]$Condition,[string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Write-SedaLog { param($Level,$Message,$Exception) }

$applicationRoot = Split-Path -Parent $PSScriptRoot
$guiPath = Join-Path $applicationRoot 'SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath,[ref]$tokens,[ref]$errors)
Assert-True (@($errors).Count -eq 0) 'The GUI script contains parse errors.'

$guiContent = Get-Content -LiteralPath $guiPath -Raw
$launcherContent = Get-Content -LiteralPath (Join-Path $applicationRoot 'Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd') -Raw
Assert-True ($launcherContent -match '%ProgramFiles%\\PowerShell\\7\\pwsh\.exe') 'The GUI launcher does not prefer PowerShell 7.'
Assert-True ($guiContent -match '\$runner = \(Get-Process -Id \$PID\)\.Path') 'Elevation does not preserve the current PowerShell runtime.'
Assert-True ($guiContent -match 'function run-collect-command\([^\r\n]+\$CommandArguments') 'The native collector does not use the safe command argument parameter.'
Assert-True ($guiContent -notmatch 'function run-collect-command\([^\r\n]+\$Args') 'The native collector still collides with the automatic $args variable.'
Assert-True ($guiContent -match '-CommandArguments @\(\$cmd\.Args\)') 'Native command arguments are not forwarded.'
Assert-True ($guiContent -match '\$controls\.SummaryErrors\.Text = \[string\]\$analysis\.Health\.ErrorCount') 'The overview still mixes historical events into the actionable error count.'
Assert-True ($guiContent -match '\$controls\.SummaryWarnings\.Text = \[string\]\$analysis\.Health\.WarningCount') 'The overview still mixes historical events into the actionable warning count.'
Assert-True ($guiContent -match '\$message = ''\{0\}\|\{1\}\|\{2\}\|\{3\}''') 'The collector progress file does not include current and total steps.'
Assert-True ($guiContent -match 'Total = 57 \+ \$optionalMdmStepCount \+ \$optionalReadinessStepCount') 'The collection total does not account for optional steps.'
Assert-True ($guiContent -match 'Step \{0\}/\{1\} - \{2\} - elapsed \{3\}s') 'The GUI does not display numbered collection progress.'
Assert-True ($guiContent -match 'Text="Device Health Errors"') 'The overview does not distinguish device health from the overall diagnostic score.'
Assert-True ($guiContent -match 'Overall diagnostic score') 'The overall diagnostic score is not clearly labelled.'
Assert-True ($guiContent -match 'Explicit update result groups') 'Windows Update explicit results are not separated from ETL diagnostic traces.'
Assert-True ($guiContent -match '\$controls\.OverlayProgress\.IsIndeterminate = -not \$isDeterminate') 'The overlay progress bar is not switched to determinate mode.'
Assert-True ($guiContent -match '-ReuseProgress \| Out-Null') 'The Group Policy fallback incorrectly consumes an extra progress step.'
Assert-True ($guiContent -match 'function Assert-SedaSafeZip') 'ZIP safety validation is missing.'
Assert-True ($guiContent -match 'ProgressPath = \$workingPath \+ ''\.progress\.txt''') 'Collection progress is not stored outside the directory being archived.'
Assert-True ($guiContent -match 'Result=''PASS''') 'The local collector does not write an explicit PASS result marker.'
Assert-True ($guiContent -match '\$script:SedaAnalysisContext') 'ZIP analysis is not isolated from the WPF UI thread.'
Assert-True ($guiContent -match 'ExportAnalysisClixmlPath') 'The isolated analysis result handoff is missing.'
Assert-True ($guiContent -match 'Get-Command pwsh\.exe') 'The GUI analysis worker does not prefer PowerShell 7.'
Assert-True ($guiContent -match 'Remove-SedaAnalysisExtraction -Path \$analysis\.ExtractDir') 'CLI extraction cleanup is not wired to the extracted directory.'
Assert-True ($guiContent -notmatch 'Remove-SedaAnalysisExtraction -Analysis') 'Extraction cleanup still passes an unsupported parameter.'
Assert-True ($guiContent -match 'AssessmentName = ''Local configuration assessment''') 'The local assessment is not clearly distinguished from official Intune compliance.'
Assert-True ($guiContent -match 'ScoreComponents = \$scoreComponents') 'The diagnostic score breakdown is missing.'
Assert-True ($guiContent -match "'INFO'.*historical Modern Auth error code") 'Historical Modern Auth evidence is not informational.'
Assert-True ($guiContent -match "Keyword 'Managed policies' -MatchMode Exact") 'Managed and unmanaged MDM headings are not separated exactly.'
Assert-True ($guiContent.Contains("if (`$raw -match '^hex\((2|7|b)\):(.+)')")) 'REG_MULTI_SZ and REG_QWORD decoding is missing.'
Assert-True ($guiContent -match 'Collector runtime: Windows PowerShell 5\.1') 'The collector runtime is not explained in the GUI.'
Assert-True ($guiContent -match 'ETL-only signatures do not reduce the score') 'The diagnostic score scope is not documented in the analysis result.'
Assert-True ($guiContent -notmatch 'in \(Get-SedaTextContent -Path [^\r\n]+ -split') 'A text parser still passes -split as an argument instead of splitting the returned content.'
Assert-True ($guiContent -match 'Technical evidence \(kept out of the main grid\)') 'Insights do not expose technical evidence outside the main grid.'
Assert-True ($guiContent -match 'New-SedaHtmlTechnicalDetails') 'The HTML report does not collapse technical evidence.'
Assert-True ($guiContent -match 'Get-SedaWindowsUpdateCachePath') 'Windows Update ETL conversion is not cached per diagnostics ZIP.'
Assert-True ($guiContent -match 'isFlightingFallback') 'IME flighting fallback classification is missing.'
Assert-True ($guiContent -match "Keyword 'mdm_enrollment'") 'The singular local MDM enrollment registry export is not detected.'
Assert-True ($guiContent -match '\[string\[\]\]\$HtmlFiles' -and $guiContent -match '\[string\[\]\]\$XmlFiles') 'Direct local MDM HTML/XML reports are not accepted.'
Assert-True ($guiContent -match 'ConnectionInfo = \$connectionInfo') 'The combined local connection model is missing from the analysis result.'
Assert-True ($guiContent -match '<WrapPanel x:Name="HeaderPanel"') 'The tab navigation can still reorder wrapped tab rows.'
Assert-True ($guiContent -match 'x:Name="WuHistoryGrid"') 'The structured Windows Update history view is missing.'

foreach ($functionName in @(
    'New-SedaObject',
    'Get-SedaTextContent',
    'Repair-SedaTextMojibake',
    'Get-SedaOperatorSummary',
    'Get-SedaWindowsUpdateCodeInfo',
    'Get-SedaWindowsUpdateLogEvents',
    'Get-SedaExtendedDirectory',
    'Get-SedaModernAuthEvidence',
    'Set-SedaImeContext',
    'Expand-SedaCabFile',
    'ConvertFrom-SedaRegFile',
    'ConvertFrom-SedaRegValue',
    'ConvertFrom-SedaHtmlText',
    'Get-SedaMdmDiagSection',
    'Get-SedaImeLogEvents',
    'Read-SedaLogRecords',
    'Get-SedaImeOperationIdentity',
    'Get-SedaEvidenceCoverage',
    'Get-SedaWin11CompatibilityIndicators',
    'ConvertFrom-SedaDword',
    'Get-SedaEnrollments',
    'Group-SedaWindowsUpdateReportingEvents',
    'Get-SedaInsights',
    'ConvertTo-SedaTimelineDate',
    'Find-SedaInventoryFile',
    'Get-SedaExtraSummary',
    'Get-SedaComplianceSummary',
    'Get-SedaHardwareReadiness',
    'ConvertTo-SedaProcessArgument',
    'Stop-SedaProcessTree',
    'Invoke-SedaProcessWithTimeout'
)) {
    $functionAst = $ast.Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
    },$true)
    Assert-True ($null -ne $functionAst) "Missing function: $functionName"
    Invoke-Expression $functionAst.Extent.Text
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('seda_analysis_regression_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    $systemInfoPath = Join-Path $tempRoot 'ps_system_info.txt'
    $systemInfo = @(
        'CsName                  : TEST-PC-01',
        'OsName                  : Microsoft Windows 10 Enterprise',
        'OsVersion               : 10.0.19045',
        'OsBuildNumber           : 19045'
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($systemInfoPath,$systemInfo,[Text.UTF8Encoding]::new($false))
    $ipconfigPath = Join-Path $tempRoot 'ipconfig_exe_all.txt'
    $ipconfig = @(
        'Carte Ethernet VMware Network Adapter VMnet1 :',
        '   Adresse IPv4. . . . . . . . . . . . . .: 192.168.64.1',
        '',
        'Carte réseau sans fil Wi-Fi :',
        '   Adresse IPv4. . . . . . . . . . . . . .: 192.168.1.106',
        '   Passerelle par défaut. . . . . . . . . .: 192.168.1.254'
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($ipconfigPath,$ipconfig,[Text.UTF8Encoding]::new($false))
    $inventory = @{ all_files = @($systemInfoPath,$ipconfigPath) }
    $summary = Get-SedaExtraSummary -Inventory $inventory
    Assert-True ([string]$summary.Hostname -eq 'TEST-PC-01') 'Local computer name fallback failed.'
    Assert-True ([string]$summary.IPAddress -eq '192.168.1.106') 'A virtual adapter was selected instead of the physical network adapter.'
    Assert-True ([string]$summary.OSVersion -match 'Microsoft Windows 10 Enterprise') 'Local OS name fallback failed.'
    Assert-True ([string]$summary.OSVersion -match '10\.0\.19045') 'Local OS version fallback failed.'

    $enrollmentPath = Join-Path $tempRoot 'MDM_Enrollment.reg'
    $enrollmentReg = @(
        'Windows Registry Editor Version 5.00',
        '',
        '[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Enrollments\11111111-2222-3333-4444-555555555555]',
        '"EnrollmentState"=dword:00000001',
        '"EnrollmentType"=dword:00000001',
        '"ProviderID"="MS DM Server"',
        '"UPN"="test.user@test.invalid"',
        '"EnrollmentURL"="https://example.invalid/enrollment"'
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($enrollmentPath,$enrollmentReg,[Text.UTF8Encoding]::new($false))
    $enrollmentResult = Get-SedaEnrollments -Path $enrollmentPath
    Assert-True (@($enrollmentResult.Enrollments).Count -eq 1) 'The enrollment registry record was not parsed.'
    Assert-True ([string]$enrollmentResult.Summary['Active enrollment records'] -eq '1') 'The active enrollment summary is incorrect.'
    Assert-True ([string]$enrollmentResult.Summary['UPNs'] -eq 'test.user@test.invalid') 'The enrollment UPN summary is missing.'
    Assert-True (-not $enrollmentResult.Summary.Contains('Hostname')) 'Generic blank device fields leaked into the enrollment summary.'
    $readinessPath = Join-Path $tempRoot 'win11_readiness.json'
    Assert-True ([string](ConvertFrom-SedaRegValue -Value 'hex(7):4e,00,6f,00,6e,00,65,00,00,00,00,00') -eq 'None') 'REG_MULTI_SZ None was not decoded.'
    Assert-True ([string](ConvertFrom-SedaRegValue -Value 'hex(7):43,00,70,00,75,00,46,00,6d,00,73,00,00,00,54,00,70,00,6d,00,00,00,00,00') -eq 'CpuFms; Tpm') 'REG_MULTI_SZ blocker list was not decoded.'
    Assert-True ([string](ConvertFrom-SedaRegValue -Value 'hex(b):2a,00,00,00,00,00,00,00') -eq '42') 'REG_QWORD was not decoded.'

    $win11RegPath = Join-Path $tempRoot 'TargetVersionUpgradeExperienceIndicators.reg'
    $win11Reg = @(
        'Windows Registry Editor Version 5.00',
        '',
        '[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators\21H2]',
        '"RedReason"=hex(7):4e,00,6f,00,6e,00,65,00,00,00,00,00',
        '',
        '[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators\GE25H2]',
        '"RedReason"=hex(7):43,00,70,00,75,00,46,00,6d,00,73,00,00,00,54,00,70,00,6d,00,00,00,00,00'
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($win11RegPath,$win11Reg,[Text.UTF8Encoding]::new($false))
    $win11Indicators = Get-SedaWin11CompatibilityIndicators -Path $win11RegPath
    Assert-True (-not [bool](@($win11Indicators.Indicators | Where-Object TargetVersion -eq '21H2')[0].IsBlocking)) 'Decoded RedReason=None is still treated as a blocker.'
    Assert-True ([bool](@($win11Indicators.Indicators | Where-Object TargetVersion -eq 'GE25H2')[0].IsBlocking)) 'Concrete CpuFms/TPM reasons are not treated as blockers.'

    $managedSection = '<span class="SectionTitle">Managed policies</span><table><tr><td>AreaA</td><td>PolicyA</td><td>1</td></tr></table>'
    $unmanagedSection = '<span class="SectionTitle">Unmanaged policies</span><table><tr><td>AreaB</td><td>PolicyB</td></tr></table>'
    $selectedManagedSection = Get-SedaMdmDiagSection -Sections @($unmanagedSection,$managedSection) -Keyword 'Managed policies' -MatchMode Exact
    $selectedUnmanagedSection = Get-SedaMdmDiagSection -Sections @($managedSection,$unmanagedSection) -Keyword 'Unmanaged policies' -MatchMode StartsWith
    Assert-True ($selectedManagedSection -eq $managedSection) 'Managed policies incorrectly matched the Unmanaged policies section.'
    Assert-True ($selectedUnmanagedSection -eq $unmanagedSection) 'Unmanaged policies section selection failed.'

    $script:MdmErrorCodes = @{}
    $script:ImeThemes = @('intunemanagementextension')
    $imeLogPath = Join-Path $tempRoot 'IntuneManagementExtension.log'
    $dnsMessage = ('Diagnostic context ' * 24) + 'System.Net.Http.WinHttpException: Error 12007: The server name or address could not be resolved.'
    $imeLog = @(
        '<![LOG[[Flighting] Key: SendHeartbeatReport, value not found. Falling back to default value: True.]LOG]!><time="10:00:00.000" date="2026-08-02" component="IntuneManagementExtension" context="" type="2" thread="1">',
        "<![LOG[$dnsMessage]LOG]!><time=`"10:01:00.000`" date=`"2026-08-02`" component=`"IntuneManagementExtension`" context=`"`" type=`"3`" thread=`"1`">",
        '<![LOG[Send request failed, HTTP request error: System.Net.Http.WinHttpException: Error 12030 calling WINHTTP_CALLBACK_STATUS_REQUEST_ERROR.]LOG]!><time="10:02:00.001" date="2026-08-02" component="IntuneManagementExtension" context="" type="3" thread="1">',
        '<![LOG[[Flighting] Failed to get flighting information for key EnableIC3Feature, error System.Net.Http.WinHttpException: Error 12030, ''La connexion avec le serveur a ├®t├® interrompue anormalement''.  à System.Threading.Tasks.RendezvousAwaitable`1.GetResult(). Falling back to default value: False.]LOG]!><time="10:02:00.000" date="2026-08-02" component="IntuneManagementExtension" context="" type="3" thread="1">'
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($imeLogPath,$imeLog,[Text.UTF8Encoding]::new($false))
    $imeResult = Get-SedaImeLogEvents -ImeThemes @{ intunemanagementextension=@($imeLogPath) }
    $tiePath = Join-Path $tempRoot 'equal-time.log'
    $tieTemplate = '<![LOG[{0} System.Net.Http.WinHttpException: Error 12007]LOG]!><time="10:01:00.000" date="2026-08-02" component="IntuneManagementExtension" type="3">'
    [IO.File]::WriteAllLines($tiePath,[string[]]@(($tieTemplate -f 'B'),($tieTemplate -f 'A')))
    $tieFirst = Get-SedaImeLogEvents -ImeThemes @{intunemanagementextension=@($tiePath)}
    [IO.File]::WriteAllLines($tiePath,[string[]]@(($tieTemplate -f 'A'),($tieTemplate -f 'B')))
    $tieSecond = Get-SedaImeLogEvents -ImeThemes @{intunemanagementextension=@($tiePath)}
    Assert-True ($tieFirst.Events[0].FullMessage -eq $tieSecond.Events[0].FullMessage -and $tieFirst.Events[0].Occurrences -eq 2) 'Equal-time IME grouping is not deterministic.'
    $flightEvent = @($imeResult.Events | Where-Object { $_.FullMessage -match 'SendHeartbeatReport' })[0]
    $dnsEvent = @($imeResult.Events | Where-Object ErrorCode -eq '12007')[0]
    Assert-True ([string]$flightEvent.Severity -eq 'INFO' -and -not [bool]$flightEvent.IsActionable) 'Expected SendHeartbeatReport fallback remains actionable.'
    Assert-True ([string]$dnsEvent.KnownCode -eq 'WINHTTP_NAME_NOT_RESOLVED') 'WinHTTP 12007 was not classified.'
    $fallbackEvent = @($imeResult.Events | Where-Object { $_.FullMessage -match 'EnableIC3Feature' })[0]
    Assert-True ([string]$dnsEvent.ActionKey -eq 'NETWORK_DNS_RESOLUTION_FAILED') 'WinHTTP 12007 does not have a stable semantic action key.'
    Assert-True ([string]$dnsEvent.FullMessage -match 'could not be resolved') 'The full IME error message was truncated before classification.'
    Assert-True ([string]$dnsEvent.Message -notmatch 'could not be resolved') 'The display-message truncation fixture is not exercising the full-message path.'

    Assert-True ([string]$fallbackEvent.Severity -eq 'WARNING' -and [bool]$fallbackEvent.IsExpected -and -not [bool]$fallbackEvent.IsActionable) 'Flighting fallback 12030 should be a non-actionable warning.'
    Assert-True (@($imeResult.Events|Where-Object{$_.ErrorCode -eq '12030'-and$_.IsActionable}).Count -eq 0) 'Correlated flighting transport trace remains actionable.'
    Assert-True ([string]$fallbackEvent.FullMessage -match 'été interrompue') 'IME mojibake was not repaired.'
    $fallbackSummary = Get-SedaOperatorSummary -Text $fallbackEvent.FullMessage
    Assert-True ($fallbackSummary -match '12030' -and $fallbackSummary -notmatch 'System\.Threading') 'Operator summary still exposes a technical stack frame.'
    $ambiguousEnrollments = New-SedaObject @{ Enrollments=@(New-SedaObject @{ State='Enrolled (active)'; ProviderID='WMI_Bridge_SCCM_Server'; EnrollmentURL=''; DiscoveryServiceFullURL='' }) }
    $byodDsReg = New-SedaObject @{ Sections=@{ 'Device State'=[ordered]@{ AzureAdJoined='NO'; WorkplaceJoined='YES' }; 'Device Details'=[ordered]@{}; 'SSO State'=[ordered]@{}; 'User State'=[ordered]@{ WorkplaceJoined='YES' } }; RawText='' }
    $ambiguousAssessment = Get-SedaComplianceSummary -DsReg $byodDsReg -Enrollments $ambiguousEnrollments -Firewall (New-SedaObject @{ Profiles=@{} }) -Results (New-SedaObject @{ Errors=@() })
    Assert-True (@($ambiguousAssessment.PolicyStatuses | Where-Object { $_.Area -eq 'MDM Enrollment' -and $_.Status -eq 'NOT_EVALUATED' }).Count -eq 1) 'Ambiguous registry enrollment evidence still claims Intune compliance.'

    $readinessJson = '{"returnCode":1,"returnReason":"TPM, Processor, ","logging":"Storage: OSDiskSize=476GB. PASS; Memory: System_Memory=16GB. PASS; TPM: TPMVersion=False. FAIL; Processor: Family=6. FAIL; SecureBoot: Capable. PASS; ","returnResult":"NOT CAPABLE"}'
    [IO.File]::WriteAllText($readinessPath,$readinessJson,[Text.UTF8Encoding]::new($true))
    $readiness = Get-SedaHardwareReadiness -Path $readinessPath
    Assert-True ([bool]$readiness.Parsed) 'BOM-prefixed HardwareReadiness JSON was not parsed.'
    Assert-True ([string]$readiness.Status -eq 'NOT_CAPABLE') 'Hardware readiness status is incorrect.'
    Assert-True ([int]$readiness.ReturnCode -eq 1) 'Hardware readiness return code is incorrect.'
    Assert-True (@($readiness.Checks | Where-Object Status -eq 'FAIL').Count -ge 2) 'Hardware readiness failed checks were not extracted.'
    $wuLogPath = Join-Path $tempRoot 'WindowsUpdate.generated.log'
    $wuLog = @(
        '2026/08/01 08:00:00.0000000 100 200 Agent completed successfully',
        '2026/08/01 08:01:00.0000000 100 200 ComApi *FAILED* [80246007] ISusInternal:: IsCommitRequired',
        '2026/08/01 08:02:00.0000000 101 201 ComApi *FAILED* [80246007] ISusInternal:: IsCommitRequired',
        '2026/08/01 08:03:00.0000000 100 200 Agent Processing pending service registrations',
        '2026/08/01 08:04:00.0000000 100 200 Handler WARNING retry scheduled',
        '2026/08/01 08:05:00.0000000 100 200 Agent Callback to client with code Call complete and error 0',
        '2026/08/01 08:06:00.0000000 100 200 DownloadManager Subscribing to GDR Retry due to async handler trigger'
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($wuLogPath,$wuLog,[Text.UTF8Encoding]::new($false))
    $wuEvents = @(Get-SedaWindowsUpdateLogEvents -Path $wuLogPath)
    Assert-True ($wuEvents.Count -eq 2) 'Windows Update log events were not filtered and grouped correctly.'
    $groupedWuError = @($wuEvents | Where-Object ErrorCode -eq '0x80246007')[0]
    Assert-True ([int]$groupedWuError.Occurrences -eq 2) 'Repeated Windows Update HRESULT events were not grouped.'
    Assert-True (@($wuEvents | Where-Object Level -eq 'Warning').Count -eq 1) 'The Windows Update warning event was not retained.'
    Assert-True ([string]$groupedWuError.Meaning -match 'not downloaded') 'The Windows Update HRESULT does not expose an operator-friendly meaning.'

    $reportingRows = @(
        (New-SedaObject @{ Level='Error'; Timestamp='2026-08-01 08:00:00'; Occurrences=1; ErrorCode='0x80073D23'; Meaning='Deployment blocked for a special profile'; Source='AGENT_INSTALLING_FAILED'; Message='Installation Failure: Windows failed to install update with error 0x80073D23: Contoso.App.'; Recommendation='Use an interactive user profile.'; File='ReportingEvents.log' }),
        (New-SedaObject @{ Level='Error'; Timestamp='2026-08-01 08:01:00'; Occurrences=1; ErrorCode='0x80073D23'; Meaning='Deployment blocked for a special profile'; Source='AGENT_INSTALLING_FAILED'; Message='Installation Failure: Windows failed to install update with error 0x80073D23: Contoso.App.'; Recommendation='Use an interactive user profile.'; File='ReportingEvents.log' })
    )
    $reportingGroups = @(Group-SedaWindowsUpdateReportingEvents -Events $reportingRows)
    Assert-True ($reportingGroups.Count -eq 1) 'Explicit Windows Update results were not grouped.'
    Assert-True ([int]$reportingGroups[0].Occurrences -eq 2) 'Explicit Windows Update result occurrences are incorrect.'
    Assert-True ([string]$reportingGroups[0].Level -eq 'Warning') 'The special-profile AppX result should be contextualized as a warning.'
    Assert-True ([string]$reportingGroups[0].Message -match 'Contoso\.App') 'The grouped Windows Update result does not identify the affected item.'

    Assert-True ([string](Get-SedaWindowsUpdateCodeInfo -Code '0x80246010').Meaning -match 'sandbox') 'Windows Update download-sandbox code mapping is missing.'
    Assert-True ([string](Get-SedaWindowsUpdateCodeInfo -Code '0x80244022').Meaning -match '503') 'Windows Update HTTP 503 code mapping is missing.'
    Assert-True ([string](Get-SedaWindowsUpdateCodeInfo -Code '0x8024000B').Meaning -match 'cancelled') 'Windows Update cancelled-operation code mapping is missing.'
    Assert-True ([string](Get-SedaWindowsUpdateCodeInfo -Code '0x80073D02').Meaning -match 'in use') 'App package resources-in-use code mapping is missing.'
    $extendedRoot = Join-Path $tempRoot 'capture with spaces\extended'
    New-Item -ItemType Directory -Path $extendedRoot -Force | Out-Null
    $extendedEvidence = Join-Path $extendedRoot 'ps_disk_usage.txt'
    [IO.File]::WriteAllText($extendedEvidence,'test',[Text.UTF8Encoding]::new($false))
    Assert-True ([string](Get-SedaExtendedDirectory -Paths @($extendedEvidence)) -eq $extendedRoot) 'The extended evidence directory was not resolved from a file path.'

    $auth = Get-SedaModernAuthEvidence -Text 'AADSTS70043 AADSTS65002 AADSTS50011 AADSTS'
    Assert-True (@($auth.Codes).Count -eq 3) 'Modern Auth extraction did not return the three concrete AADSTS codes.'
    Assert-True (@($auth.Codes) -notcontains 'AADSTS') 'The generic AADSTS token was incorrectly counted as an error code.'

    $script:ImeThemes = @('intunemanagementextension','healthscripts')
    $syntheticIme = New-SedaObject @{
        Events = @(
            (New-SedaObject @{ Severity='ERROR'; IsExpected=$false; IsActionable=$true; IsRecent=$true; Theme='intunemanagementextension'; ActionKey='USER_TOKEN_FALLBACK'; Message='AAD User check failed, now fallback to the Graph audience.' }),
            (New-SedaObject @{ Severity='ERROR'; IsExpected=$false; IsActionable=$true; IsRecent=$true; Theme='intunemanagementextension'; ActionKey='USER_TOKEN_UNAVAILABLE'; Message='LogonUser failed with error code : 1008' })
        )
        Summary = New-SedaObject @{ ErrorCount=2; WarningCount=0; ActionableEventCount=2; InformationalCount=0; ThemeCounts=[ordered]@{} }
    }
    $syntheticDsReg = New-SedaObject @{ Sections=@{ 'Device State'=[ordered]@{ AzureAdJoined='YES'; DomainJoined='YES'; WorkplaceJoined='NO' }; 'User State'=[ordered]@{} } }
    $syntheticIme = Set-SedaImeContext -ImeResult $syntheticIme -DsReg $syntheticDsReg
    Assert-True ([int]$syntheticIme.Summary.ErrorCount -eq 0) 'Recovered IME fallback or missing session token is still counted as an error.'
    Assert-True ([int]$syntheticIme.Summary.WarningCount -eq 0) 'A standalone missing user-session token should not be actionable without a separate user-workload failure.'
    Assert-True (@($syntheticIme.Events | Where-Object Severity -eq 'INFO').Count -eq 2) 'Recovered fallback and the standalone session-token condition should be informational.'
    Assert-True (@($syntheticIme.Events | Where-Object { $_.ActionKey -eq 'USER_TOKEN_UNAVAILABLE' -and -not $_.IsActionable }).Count -eq 1) 'The standalone session-token condition was not classified as non-actionable.'

    $cabSourceRoot = Join-Path $tempRoot 'CAB source with spaces'
    $cabOutputRoot = Join-Path $tempRoot 'CAB output with spaces'
    New-Item -ItemType Directory -Path $cabSourceRoot,$cabOutputRoot -Force | Out-Null
    $cabSource = Join-Path $cabSourceRoot 'MDMDiagReport.xml'
    $cabPath = Join-Path $cabSourceRoot 'mdm report.cab'
    [IO.File]::WriteAllText($cabSource,'<MDMEnterpriseDiagnosticsReport />',[Text.UTF8Encoding]::new($false))
    & makecab.exe $cabSource $cabPath | Out-Null
    Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $cabPath -PathType Leaf)) 'The CAB regression fixture could not be created.'
    $cabResult = Expand-SedaCabFile -CabPath $cabPath -DestinationPath $cabOutputRoot
    Assert-True ([bool]$cabResult.Success) 'CAB extraction failed for a path containing spaces.'
    Assert-True (@($cabResult.Files | Where-Object { (Get-Content -LiteralPath $_ -Raw) -match 'MDMEnterpriseDiagnosticsReport' }).Count -eq 1) 'The expected MDM report content was not extracted from the CAB fixture.'

    $native = Invoke-SedaProcessWithTimeout -FilePath 'cmd.exe' -ArgumentList @('/d','/c','echo','seda-argument-ok') -TimeoutSeconds 10
    Assert-True (-not [bool]$native.TimedOut) 'Native argument regression command timed out.'
    Assert-True ([int]$native.ExitCode -eq 0) 'Native argument regression command failed.'
    Assert-True ([string]$native.StandardOutput -match 'seda-argument-ok') 'Native command arguments were not preserved.'
    $systemDsReg = New-SedaObject @{
        Sections = @{
            'Device State' = [ordered]@{ AzureAdJoined='YES'; WorkplaceJoined='NO' }
            'Device Details' = [ordered]@{ DeviceAuthStatus='SUCCESS'; TpmProtected='YES' }
            'SSO State' = [ordered]@{ AzureAdPrt='NO'; NgcSet='NO'; WamDefaultSet='ERROR' }
            'User State' = [ordered]@{ WorkplaceJoined='NO' }
        }
        RawText = "Executing Account Name : TEST\TEST-PC-01$, TEST-PC-01$@test.invalid"
    }
    $systemAssessment = Get-SedaComplianceSummary -DsReg $systemDsReg -Enrollments (New-SedaObject @{ Enrollments=@() }) -Firewall (New-SedaObject @{ Profiles=@{} }) -Results (New-SedaObject @{ Errors=@() })
    Assert-True ([int]$systemAssessment.NotEvaluatedCount -ge 1) 'SYSTEM-context user SSO evidence was not marked NOT_EVALUATED.'
    Assert-True ([string]$systemAssessment.OverallStatus -eq 'REVIEW') 'Incomplete local evidence should require review rather than claim Intune compliance.'

    $scoreFixture = New-SedaObject @{
        DeviceSummary = New-SedaObject @{ ComputerName='Synthetic device identity' }
        CriticalIssues = @(
            (New-SedaObject @{ Severity='WARNING'; Title='Reboot required'; Detail='RebootRequired=Yes'; Recommendation='Restart.'; Source='Windows Update' }),
            (New-SedaObject @{ Severity='INFO'; Title='Monitoring evidence unavailable'; Detail='No local evidence'; Recommendation='Run local analysis.'; Source='Device Health' })
        )
        Compliance = New-SedaObject @{ PolicyStatuses=@() }
        ImeEvents = @()
        WindowsUpdate = New-SedaObject @{
            Issues=@(New-SedaObject @{ Severity='WARNING'; Title='Reboot required'; Detail='RebootRequired=Yes'; Recommendation='Restart.'; Source='Windows Update' })
            ReportingGroups=@()
            ReportingEvents=@()
            EtlEvents=@(New-SedaObject @{ Level='Error'; ErrorCode='0x80246007'; Meaning='Content was not downloaded'; Occurrences=9; LastSeen='2026-08-01'; Source='DownloadManager'; Message='Trace signature'; EtlFile='WindowsUpdate.log' })
            Info=[ordered]@{ 'Reboot Required'='Yes' }
            Policies=@()
        }
        Win11Compatibility = New-SedaObject @{ Status='NOT_APPLICABLE'; BlockingIndicators=@(); HardwareReadiness=(New-SedaObject @{ Checks=@() }) }
        Health = New-SedaObject @{ Findings=@() }
        EventLogs = New-SedaObject @{ Events=@() }
    }
    $scoreResult = Get-SedaInsights -Analysis $scoreFixture
    Assert-True ([int]$scoreResult.Score -eq 95) 'INFO evidence, duplicate reboot evidence or ETL-only signatures still reduce the diagnostic score.'
    Assert-True (@($scoreResult.RootCauses | Where-Object { $_.Title -like 'Windows Update ETL*' -and $_.Severity -eq 'INFO' }).Count -eq 1) 'ETL diagnostic evidence should remain visible as informational context.'
    Assert-True (@($scoreResult.ScoreComponents).Count -ge 7) 'The score breakdown does not expose all scoring categories.'
    $appliedPenaltyTotal = [int](($scoreResult.ScoreComponents | Measure-Object AppliedPenalty -Sum).Sum)
    Assert-True ($appliedPenaltyTotal -eq 5) 'The score breakdown does not reconcile to the overall score.'

    $wingetTechnical = '[Win32App][WinGetApp][AppPackageManager] Installer error code: 0x80004004 (E_ABORT). Exception: System.Exception: Operation aborted. at Microsoft.Management.Clients.Workload.Run()'
    $scoreFixture.ImeEvents = @(
        New-SedaObject @{ Severity='ERROR'; IsActionable=$true; IsRecent=$true; IsExpected=$false; Theme='intunemanagementextension'; Category='IntuneManagementExtension'; ActionKey='WINGET_INSTALL_ABORTED'; Message=$wingetTechnical; FullMessage=$wingetTechnical; SourceFile='IntuneManagementExtension.log'; Timestamp='2026-08-12 12:00:00'; Occurrences=1; ErrorCode='0x80004004'; KnownCode='' }
    )
    $wingetResult = Get-SedaInsights -Analysis $scoreFixture
    $wingetAction = @($wingetResult.TopActions | Where-Object Title -match 'Winget')[0]
    Assert-True ($null -ne $wingetAction) 'The actionable Winget failure is missing from Insights.'
    Assert-True ([string]$wingetAction.Detail -notmatch 'System\.Exception|Microsoft\.Management') 'The Winget operator summary still exposes technical exception detail.'
    Assert-True ([string]$wingetAction.TechnicalDetail -match 'System\.Exception') 'The full Winget technical evidence was not retained.'

    [pscustomobject]@{
        Result = 'PASS'
        PackageVersion = '0.3.0'
        PowerShellEdition = $PSVersionTable.PSEdition
        PowerShellVersion = [string]$PSVersionTable.PSVersion
        ComputerName = [string]$summary.Hostname
        OSVersion = [string]$summary.OSVersion
        IPAddress = [string]$summary.IPAddress
        HardwareReadiness = [string]$readiness.Status
        FailedReadinessChecks = @($readiness.Checks | Where-Object Status -eq 'FAIL').Count
        NativeArgumentsPreserved = $true
        OverviewUsesHealthFindings = $true
        NumberedCollectionProgress = $true
        DeterminateProgressBars = $true
    }
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDVA8tzagBmcLAe
# bPTZJ7DdH9ahAp3hPnyqS+QdIyzwtqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDZsPSp5usRfl4SJG0GyJ77p5wzXmjZcSAE5L9wZ1hTWMA0GCSqG
# SIb3DQEBAQUABIIBgDKBYFY07O/8KNS1sCUUON2zFxO0bcZxOZGl7f9cE6Nxs830
# 98gjERaYyhzv2n+JBKHLGQBe5dRVcA2fuTp0cTyCU8JDlKIhvH2Dr9y7kQAWMcS3
# ngFKtzz/IOcdzBy1bzqLuiVrp0mKYasbOzfm7CJwoZaDJvYyNfE6xDYfUvXr9Gs0
# 4PRfMa8iO5xfYi25A/epNWYkKD9m/7DIg0ntFWriC9/wSCOef/c+z6YgdkUs8L0A
# adlbV3UICLIzuZ3Jbz5SFMZHmy8y+0NGtSS03wzHAmkDbY75RxjxH0L9Pi7ixiRP
# ijmyDwjm2EQgL/mXCZ1KTH2FpvoALEX00YWI4HLS0FjCWlgqxETF3T4kJBqiTs+U
# +Fwo3Ry3s8Uc6TjkoP85peZaV00u5AKtYNz8uLJMzpkyCTlMR86ufII4krsx/a89
# eYOyv6l7gQFjysQeQO8ANknKclczybvcdDpBMvd1lYNJKObNMgIwAWyApeB1kzxn
# UxQA216Bc1QZVvs3u6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDQxNTQw
# MzdaMC8GCSqGSIb3DQEJBDEiBCDXwmOoVZg7CfkPtzhqrrNxAqWMq/rlMtvX980X
# VMHwbjANBgkqhkiG9w0BAQEFAASCAgBG/4BWny61aCH+GOHw32sh8QxWFnrHToB4
# jd1hJOUSC5pUp5Etor9gp8fCDMrjCAWnfwIMlKQk1tjDIvT31pLboNyt+3ubFZVJ
# si5quJDVak0+TRXXHTLTJNpylyHVVfLb69DBDfdNxX5nt6vFYZAoVrKjZXdSYeCw
# I5maKG5Pm/FaY8pjE9xuAzdchkJZV6LeSG6fHTOK1E6pUO2A63ruO1YlwPSDcF7d
# snEErgeE+EVCiPASJps0Vp8GJVWC62FZq0Z4+3uigV5cB4S5ms4HHHyc8QkBjTkD
# Lwq0HDQNTquIppIxWDxI9ETVV+mJTy4CwfrSS0BKVM9h5pZl9Vib52MdZuesZQXE
# 4b2npjUFrGO0E79VYaGMKfHUDOKpFUBeftWyTet4YHJpHekfbJmexwNjiPVvtSCU
# hguh52sIvN18rGjGsnJ2U1IR6rGpMKXrw0a2AMbn32Bq13zK2t9a6va1jpf2Ccca
# STnlKaRz+cD0LjO3Ye1s0vaXKZka/WkEw0XAvsh96M0NS+368Ho3ikgrhnFsiocJ
# SbJswVLJbC4TSpvyn07bdwmdTia4S8X/6/4IvsLOHmMBKq6XbNnaq2YWHcpKuClJ
# SV3eZi3kyLEnxS0OsuReQiNwIMmPEGcDGJl4KDn5cni4tdb1KBv7hrpbncfyUGFh
# ZwEn3OsM4g==
# SIG # End signature block
