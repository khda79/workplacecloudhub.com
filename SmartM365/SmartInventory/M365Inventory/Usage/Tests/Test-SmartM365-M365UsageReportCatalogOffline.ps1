<#
.SYNOPSIS
    Offline checks for the usage report catalog and the Copilot user count trend.
.DESCRIPTION
    Runs the collectors' actual catalog, URI, normalization and trend export
    functions on synthetic CSVs that use the column headers documented by
    Microsoft Graph. Graph and CSV publication are mocked. No tenant collection
    or publication is performed.
.VERSION
1.0
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Import-CollectorFunction {
    param([string]$Path,[string[]]$Names)
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
    foreach ($name in $Names) {
        $definition = @($ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
        },$true))
        if ($definition.Count -ne 1) { throw "Expected one $name function in $Path." }
        Invoke-Expression ('function global:{0} {1}' -f $name,$definition[0].Body.Extent.Text)
    }
    return $ast
}

$script:checks = 0
function Assert-Case { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

$usageCollector = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-M365UserActivity-Inventory.ps1')).Path
$copilotCollector = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-CopilotUsage-Inventory.ps1')).Path
$usageAst = Import-CollectorFunction -Path $usageCollector -Names @('Get-M365UsageReportDefinition','Resolve-M365UsageReportSelection','ConvertFrom-M365GenericReport','Invoke-M365UsageReportDownload')
$null = Import-CollectorFunction -Path $copilotCollector -Names @('Get-SourcePropertyValue','ConvertTo-CopilotDateText','ConvertTo-CopilotCounter','Get-CopilotUserCountTrendColumns','ConvertFrom-CopilotUserCountTrendReport','Get-CopilotUsageReportUri','Get-GraphHttpStatusCode','Invoke-CopilotUserCountTrendExport')

$global:runId = '20261028_120000'
$global:Period = 'D180'
$script:warnings = [System.Collections.Generic.List[string]]::new()
$script:uris = [System.Collections.Generic.List[string]]::new()
$script:graphResponse = $null
$script:exports = [System.Collections.Generic.List[object]]::new()
function global:WriteLog { param([string]$Message,[string]$Level) if ($Level -eq 'WARNING') { $script:warnings.Add($Message) | Out-Null } }
function global:Invoke-MgGraphRequest {
    param([string]$Method,[string]$Uri,[string]$OutputFilePath,$ProgressAction,$ErrorAction)
    $script:uris.Add($Uri) | Out-Null
    if ($null -eq $script:graphResponse) { throw 'Response status code does not indicate success: NotFound (Not Found).' }
    Set-Content -LiteralPath $OutputFilePath -Value $script:graphResponse -Encoding UTF8
}
function global:Export-SmartM365Csv {
    param([string]$BaseFileName,[string]$OutputPath,[string]$GlobalPath,$Data,[string[]]$Columns)
    $script:exports.Add([pscustomobject]@{ BaseFileName = $BaseFileName; Data = @($Data); Columns = $Columns }) | Out-Null
    [pscustomobject]@{ LatestPath = Join-Path $GlobalPath "$BaseFileName.csv" }
}

# UserActivity catalog: the three added reports, selectable by name and through All.
$definitions = Get-M365UsageReportDefinition
$expectedNew = [ordered]@{
    EmailAppUsageUserDetail = @('getEmailAppUsageUserDetail','M365_Email_AppUsage','Get-MgReportEmailAppUsageUserDetail')
    OneDriveActivityUserDetail = @('getOneDriveActivityUserDetail','M365_OneDrive_UserActivity','Get-MgReportOneDriveActivityUserDetail')
    Office365GroupsActivityDetail = @('getOffice365GroupsActivityDetail','M365_Groups_Activity','Get-MgReportOffice365GroupActivityDetail')
}
foreach ($name in $expectedNew.Keys) {
    Assert-Case $definitions.Contains($name) "Report $name is missing from the catalog."
    $definition = $definitions[$name]
    Assert-Case ($definition.Endpoint -ceq $expectedNew[$name][0]) "$name endpoint is $($definition.Endpoint)."
    Assert-Case ($definition.BaseFileName -ceq $expectedNew[$name][1]) "$name file is $($definition.BaseFileName)."
    Assert-Case ($definition.Command -ceq $expectedNew[$name][2]) "$name fallback command is $($definition.Command)."
    Assert-Case ($definition.SupportsPeriod -and $definition.Normalize -eq 'Generic') "$name must be a period-based generic report."
    Assert-Case (-not ($definition.PSObject.Properties['Required'] -and $definition.Required)) "$name must not be required: an empty answer keeps the previous CSV."
}
$fileNames = @($definitions.Values | ForEach-Object BaseFileName)
Assert-Case ($fileNames.Count -eq @($fileNames | Sort-Object -Unique).Count) 'Two reports publish the same CSV name.'
Assert-Case (@(Resolve-M365UsageReportSelection -RequestedReports @('All')).Count -eq $definitions.Count) '-Reports All does not select every catalog report.'
Assert-Case (@(Resolve-M365UsageReportSelection -RequestedReports @('EmailAppUsageUserDetail','Office365GroupsActivityDetail')).Count -eq 2) 'Named selection of the added reports failed.'

$reportsParameter = $usageAst.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'Reports' }
$validateSet = $reportsParameter.Attributes | Where-Object { $_.TypeName.Name -eq 'ValidateSet' }
$allowed = @($validateSet.PositionalArguments | ForEach-Object { $_.Value })
Assert-Case ((@($allowed | Where-Object { $_ -ne 'All' } | Sort-Object) -join ',') -ceq (@($definitions.Keys | Sort-Object) -join ',')) '-Reports ValidateSet and the catalog differ.'

# Graph URI built for each added report with the run period.
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-UsageCatalog-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $rawPath = Join-Path $tempRoot 'raw.csv'
    $script:graphResponse = 'Report Refresh Date,User Principal Name'
    foreach ($name in $expectedNew.Keys) {
        $script:uris.Clear()
        Invoke-M365UsageReportDownload -ReportDefinition $definitions[$name] -RawPath $rawPath
        Assert-Case ($script:uris[0] -ceq "https://graph.microsoft.com/v1.0/reports/$($expectedNew[$name][0])(period='D180')") "Unexpected URI for ${name}: $($script:uris[0])"
    }

    # Generic export keeps every documented source column after the three run columns.
    $documented = [ordered]@{
        EmailAppUsageUserDetail = 'Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Last Activity Date,Mail For Mac,Outlook For Mac,Outlook For Windows,Outlook For Mobile,Other For Mobile,Outlook For Web,POP3 App,IMAP4 App,SMTP App,Report Period'
        OneDriveActivityUserDetail = 'Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Assigned Products,Report Period'
    }
    foreach ($name in $documented.Keys) {
        $header = $documented[$name]
        $values = @($header -split ',' | ForEach-Object { 'v' })
        $rows = @(("$header`n" + ($values -join ',')) | ConvertFrom-Csv)
        $converted = @(ConvertFrom-M365GenericReport -Rows $rows -ReportName $name -RequestedPeriod 'D180')
        $columns = @($converted[0].PSObject.Properties.Name)
        Assert-Case (($columns -join ',') -ceq ('RunId,ReportName,ReportPeriodRequested,' + $header)) "$name generic columns changed: $($columns -join ',')"
        Assert-Case ($converted[0].ReportName -ceq $name -and $converted[0].RunId -ceq $global:runId) "$name run columns are wrong."
    }

    # Copilot trend URI: v1 without the version parameter, v2 with it, user detail unchanged.
    Assert-Case ((Get-CopilotUsageReportUri -Version v2 -RequestedPeriod D180 -FunctionName getMicrosoft365CopilotUserCountTrend) -ceq "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUserCountTrend(period='D180', version='v2')?`$format=text/csv") 'Trend v2 URI is wrong.'
    Assert-Case ((Get-CopilotUsageReportUri -Version v1 -RequestedPeriod D30 -FunctionName getMicrosoft365CopilotUserCountTrend) -ceq "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUserCountTrend(period='D30')?`$format=text/csv") 'Trend v1 URI is wrong.'
    Assert-Case ((Get-CopilotUsageReportUri -Version v2 -RequestedPeriod D28) -ceq "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='D28', version='v2')?`$format=text/csv") 'User detail URI changed.'

    # Copilot trend normalization: v1 documented header, then v2 with its extra columns.
    $trendColumns = @(Get-CopilotUserCountTrendColumns)
    $v1Header = 'Report Refresh Date,Report Date,Microsoft Teams Enabled Users,Microsoft Teams Active Users,Word Enabled Users,Word Active Users,PowerPoint Enabled Users,PowerPoint Active Users,Outlook Enabled Users,Outlook Active Users,Excel Enabled Users,Excel Active Users,OneNote Enabled Users,OneNote Active Users,Loop Enabled Users,Loop Active Users,Any App Enabled Users,Any App Active Users,Copilot Chat Enabled Users,Copilot Chat Active Users,Report Period'
    $v1Rows = @("$v1Header`n2026-07-03,2026-07-03,25,1,25,2,25,3,25,4,25,5,25,6,25,7,25,8,25,9,7" | ConvertFrom-Csv)
    $v1 = @(ConvertFrom-CopilotUserCountTrendReport -Rows $v1Rows -EffectiveReportVersion v1)
    Assert-Case ((@($v1[0].PSObject.Properties.Name) -join ',') -ceq ($trendColumns -join ',')) 'Trend row columns differ from the stable column list.'
    Assert-Case ($v1[0].ReportVersion -ceq 'v1' -and $v1[0].ReportPeriod -ceq 'D180' -and $v1[0].ReportDate -ceq '2026-07-03' -and $v1[0].ReportRefreshDate -ceq '2026-07-03') 'Trend v1 identity columns are wrong.'
    Assert-Case ($v1[0].TeamsActiveUsers -eq 1 -and $v1[0].WordActiveUsers -eq 2 -and $v1[0].PowerPointActiveUsers -eq 3 -and $v1[0].OutlookActiveUsers -eq 4 -and $v1[0].ExcelActiveUsers -eq 5 -and $v1[0].OneNoteActiveUsers -eq 6 -and $v1[0].LoopActiveUsers -eq 7 -and $v1[0].AnyAppActiveUsers -eq 8 -and $v1[0].CopilotChatActiveUsers -eq 9 -and $v1[0].AnyAppEnabledUsers -eq 25) 'Trend v1 counters were mapped to the wrong columns.'
    foreach ($v2Only in @('EdgeEnabledUsers','EdgeActiveUsers','Microsoft365CopilotEnabledUsers','Microsoft365CopilotActiveUsers','CopilotChatWorkEnabledUsers','CopilotChatWorkActiveUsers','CopilotChatWebEnabledUsers','CopilotChatWebActiveUsers','PromptsSubmitted')) {
        Assert-Case ([string]$v1[0].$v2Only -ceq '') "v1 trend invented a value for $v2Only."
    }
    $v2Header = $v1Header.Replace(',Report Period',',Edge Enabled Users,Edge Active Users,Microsoft 365 Copilot Enabled Users,Microsoft 365 Copilot Active Users,Copilot Chat (work) Enabled Users,Copilot Chat (work) Active Users,Copilot Chat (web) Enabled Users,Copilot Chat (web) Active Users,Prompts submitted,Report Period')
    $v2Csv = "$v2Header`n2026-10-26,2026-10-25,25,1,25,2,25,3,25,4,25,5,25,6,25,7,25,8,25,9,20,10,21,11,22,12,23,13,140,28`n2026-10-26,2026-10-26,25,1,25,2,25,3,25,4,25,5,25,6,25,7,25,8,25,9,20,10,21,11,22,12,23,13,150,28"
    $v2 = @(ConvertFrom-CopilotUserCountTrendReport -Rows @($v2Csv | ConvertFrom-Csv) -EffectiveReportVersion v2)
    Assert-Case ($v2.Count -eq 2 -and $v2[1].ReportDate -ceq '2026-10-26') 'Trend v2 did not keep one row per report date.'
    Assert-Case ($v2[0].EdgeEnabledUsers -eq 20 -and $v2[0].EdgeActiveUsers -eq 10 -and $v2[0].Microsoft365CopilotActiveUsers -eq 11 -and $v2[0].CopilotChatWorkActiveUsers -eq 12 -and $v2[0].CopilotChatWebEnabledUsers -eq 23 -and $v2[0].CopilotChatWebActiveUsers -eq 13 -and $v2[0].PromptsSubmitted -eq 140) 'Trend v2 extra counters were mapped to the wrong columns.'

    # Trend export: success publishes its own CSV; Graph failure and empty report keep the previous CSV.
    $trendRaw = Join-Path $tempRoot 'trend.csv'
    $global:runOutputRoot = Join-Path $tempRoot 'run'
    $global:LatestCsvFolderPath = Join-Path $tempRoot 'DATA-LAST'
    $script:graphResponse = $v2Csv
    $script:exports.Clear(); $script:warnings.Clear(); $script:uris.Clear()
    $result = Invoke-CopilotUserCountTrendExport -DestinationPath $trendRaw -Version v2 -Columns $trendColumns
    Assert-Case ($null -ne $result -and $result.Rows -eq 2 -and $script:warnings.Count -eq 0) 'Populated trend was not exported.'
    Assert-Case ($script:exports.Count -eq 1 -and $script:exports[0].BaseFileName -ceq 'M365_Copilot_UserCountTrend' -and (($script:exports[0].Columns) -join ',') -ceq ($trendColumns -join ',') -and $script:exports[0].Data.Count -eq 2) 'Trend export call is wrong.'
    Assert-Case ($script:uris[0] -match "getMicrosoft365CopilotUserCountTrend\(period='D180', version='v2'\)") 'Trend export did not request the run version.'

    $script:graphResponse = $null
    $script:exports.Clear(); $script:warnings.Clear()
    Remove-Item -LiteralPath $trendRaw -Force -ErrorAction SilentlyContinue
    # The collector runs without strict mode; its status-code probe reads optional exception properties.
    $result = & { Set-StrictMode -Off; Invoke-CopilotUserCountTrendExport -DestinationPath $trendRaw -Version v2 -Columns $trendColumns }
    Assert-Case ($null -eq $result -and $script:exports.Count -eq 0) 'Failed trend download published a CSV.'
    Assert-Case ($script:warnings.Count -eq 1 -and $script:warnings[0] -match 'download failed; previous M365_Copilot_UserCountTrend\.csv kept unchanged') 'Failed trend download warning missing.'

    $script:graphResponse = $v2Header
    $script:exports.Clear(); $script:warnings.Clear()
    $result = Invoke-CopilotUserCountTrendExport -DestinationPath $trendRaw -Version v2 -Columns $trendColumns
    Assert-Case ($null -eq $result -and $script:exports.Count -eq 0) 'Empty trend published a CSV.'
    Assert-Case ($script:warnings.Count -eq 1 -and $script:warnings[0] -match 'trend is empty; previous M365_Copilot_UserCountTrend\.csv kept unchanged') 'Empty trend warning missing.'
}
finally {
    if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
}

Write-Host ("PASS: usage report catalog and Copilot trend offline checks ({0} checks)." -f $script:checks)
