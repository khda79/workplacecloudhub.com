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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC3al/hxQwy8BKW
# GcZbPyNEFjHlv8Vvx6tY59xjy7EKQKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMUdb8J7WqE+BqcUQNzMrg/JbN+PLwbu7PapQNTs68ndMA0GCSqG
# SIb3DQEBAQUABIIBgCmZsKe0CBuLcGoSygeWKEp7YcrKMFA7TPjMdInfMzH6VF5e
# m0MMK1fTQnZM108a1wHybHGchyW3y/rr4BY6bHfhOZU+eCota4Q7J955ij0KK2HI
# ZT6mGsEmPr3/4oCKAo/sGrSMVdSFWq0Rb4ZtsrFFTnohgQ9PcXWJiAxFGm6Q76zG
# lHS5Do5Lq924V4uTXprcVUK2I+X9RBazWUtAQlA9ELRSnQKj0P2DuRm5KNB8LMbm
# HbGXWRjNhbpocWjbWmBVMBmIG8Arq0iJetMZ8Fr49S/e7m5+V4sUOzMK+y2I9jYN
# h1DAyV2USTD8AXaCd/mxcbFFd5uf0h2UT9Ng4PsLAGnhQaexVi3FGDT7pIgkA/+q
# 4Xi5s55LBuzHXh/UY7ta2Yr1d71ZTF2C4TyBgMjj5LbmYgTZ/y64FMi7QW4fglJ/
# bBdfwJBR7c3ABcuVRydPD06BhyzlOlqHR1J1+JRgUsJf2CSEQwDhVYaCw5oTmQEn
# ZnYC8EpAYXSN8zTmmKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgwNzUx
# NDhaMC8GCSqGSIb3DQEJBDEiBCACX659dGylOaZDvyXTA5EeAo4ts6XwD8LmBY0v
# zaVh/jANBgkqhkiG9w0BAQEFAASCAgBqOjHqeEeyMWsDUDTscs3ztxqQ75iNGBsx
# 4Qa5BEePIyl61Ta9zP6xhusaa0uBCfrYSVu6s4tPagw/JvAsXGQuJ6SCB1RhMsL4
# U/5qn/scY323Mu5mkYkaNzxpCGieEq3hasgIkPbPIEa3fYpHXOQRebVmUgcVIncs
# CUknK47SDE8Kh4Crlm2kQHS1HXDDdE9gcp6WpYwpczuqqzkBVevbisaorACSZp6T
# rER1uYAruK6NqIVVgBJi2FjHYDmf2T02Cc+8qDfRdy0OCxkyyjgTdCP0g7jhf/qb
# 0RECFQp+ku5h1K2tOjR3AP6aXLkGuy0lRgiAkYnUUuDcGuMe/gLwcu0F4AdokPlQ
# iu455tXGWdF+h7tNBjuhGvZye7tYOf3R8JlD4gyYYO3sGlblPbDM9YAKd5nUlpO0
# yEeSRwBPHeiUEtuxj7MORGm5zPstmV7Oh8TkwevcEb/2liF8HPN1e0cVDUEZA48i
# IrUGNvwdJ5Veub3Ay8RKyGVlUamQVIkHEFaZgXZzY6Icb3yuTSZj/xhPSettnKnu
# J/eVQpdlsqgjEToRKrCRNucwQzqAQ3qFw5LVunWB7/ga4rXm6DfZPJsPy6xF6qkI
# z1SnszfXKQ4sVODGbxZ6r9J07nFuh3htKqdp61U6FHC5O0OElf6eadsQFBf8bxT9
# 7RwjlgXaow==
# SIG # End signature block
