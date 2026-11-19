#Requires -Version 5.1
<#
.SYNOPSIS
Synthetic Exchange infrastructure and readiness summary mail tests; extracts functions without running the collector.
.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    # Optional: writes the synthetic mail body rendered with the Core module here, for visual review.
    [string]$RenderPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$collectorPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
$coreVariants = [ordered]@{
    Core                = Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365.Core.psm1'
    WindowsPowerShell5  = Join-Path $SourceRoot 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1'
}

$reportAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $SourceRoot 'SmartInventory/Common/SmartM365.ReportRenderers.ps1'),[ref]$null,[ref]$null)
function Get-FunctionDefinition {
    param([string]$Path, [string[]]$Names)
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Parse failed: $Path" }
    foreach ($name in $Names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $node) {
            $node = $reportAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        }
        if ($null -eq $node) { throw "Missing function $name in $Path" }
        $node.Extent.Text
    }
}

$collectorFunctions = @(Get-FunctionDefinition -Path $collectorPath -Names @(
    'Get-ObjectPropertyValue', 'Format-HtmlValue', 'ConvertTo-ServersAndStorageDatabaseSectionHtml',
    'ConvertTo-ServersAndStorageHealthCheckerSectionHtml', 'New-HtmlExecutiveSummary',
    'ConvertTo-ServersAndStorageEmailHtmlText', 'New-ServersAndStorageEmailLinkHtml',
    'ConvertTo-ServersAndStorageMailNumber', 'Format-ServersAndStorageMailNumber', 'Limit-ServersAndStorageMailText',
    'ConvertTo-ServersAndStorageMailDate', 'New-ServersAndStorageMailTableHtml', 'New-ServersAndStorageMailNoteHtml',
    'Get-ServersAndStorageMailFile', 'New-ServersAndStorageMailBodyHtml', 'Send-ServersAndStorageHtmlReport'))
$coreNames = @('ConvertTo-SmartM365EmailHtmlText', 'New-SmartM365EmailKpiGridHtml', 'New-SmartM365EmailTableHtml', 'New-SmartM365EmailBody')

$script:Checks = 0
function Assert-InfraMail {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:Checks++
}

function New-TestModule {
    param([string]$CorePath)
    $coreFunctions = @(Get-FunctionDefinition -Path $CorePath -Names $coreNames)
    $stubs = @'
Set-StrictMode -Version Latest
# The collector imports Core with -Prefix Core; the extracted module functions are exposed under those names.
Set-Alias -Name ConvertTo-CoreSmartM365EmailHtmlText -Value ConvertTo-SmartM365EmailHtmlText
Set-Alias -Name New-CoreSmartM365EmailKpiGridHtml -Value New-SmartM365EmailKpiGridHtml
Set-Alias -Name New-CoreSmartM365EmailTableHtml -Value New-SmartM365EmailTableHtml
Set-Alias -Name New-CoreSmartM365EmailBody -Value New-SmartM365EmailBody
$script:CapturedMail = $null
$script:Tenant = 'test'
$script:ServersAndStorageWarningCount = 0
$script:ServersAndStorageErrorCount = 0
$script:ServersAndStorageSharePointUploads = New-Object System.Collections.ArrayList
function Import-SmartM365CoreModule { }
function Write-Log { param($Message, $Level) }
function Get-SmartM365EffectiveConfigValue {
    param($Name, $DefaultValue)
    switch ($Name) { 'From' { 'sender@example.invalid' } 'To' { 'recipient@example.invalid' } default { $DefaultValue } }
}
function CoreSendEmailHtmlReport {
    param($From, $To, $Subject, $BodyHtml, $Attachments, $SmtpServer, $SendMailMode, $Cc)
    $script:CapturedMail = [pscustomobject]@{ Subject = $Subject; Body = $BodyHtml; Bound = @($PSBoundParameters.Keys) }
}
function Get-MailboxDatabase { throw 'A synthetic mail must not collect Exchange data.' }
'@
    return New-Module -ScriptBlock ([scriptblock]::Create($stubs + "`n" + (($coreFunctions + $collectorFunctions) -join "`n")))
}

function New-Row {
    param([string]$Category, [string]$Object, [string]$Setting, [AllowNull()]$Value, [string]$Status = 'OK', [string]$Importance = 'Review', [string]$ErrorMessage = '')
    [pscustomobject]@{ Category = $Category; ObjectName = $Object; Setting = $Setting; Value = $Value; MigrationFocus = 'Synthetic fixture'; Importance = $Importance; CollectionStatus = $Status; ErrorMessage = $ErrorMessage }
}

function Get-InnerDataTable {
    param([string]$Html)
    # Innermost tables that carry a header row (data tables built for the mail).
    @([regex]::Matches($Html, '(?is)<table(?<Attrs>[^>]*)>(?<Inner>(?:(?!<table).)*?)</table>') | Where-Object { $_.Groups['Inner'].Value -match '<th' })
}

function Assert-FrameWidth {
    param([string]$Html, [string]$Label)
    foreach ($match in [regex]::Matches($Html, '(?i)(?:max-)?width\s*:\s*(\d+)px')) { Assert-InfraMail ([int]$match.Groups[1].Value -le 760) "$Label has a CSS width above 760px: $($match.Value)" }
    foreach ($match in [regex]::Matches($Html, '(?i)\swidth="(\d+)"')) { Assert-InfraMail ([int]$match.Groups[1].Value -le 760) "$Label has a width attribute above 760: $($match.Value)" }
    Assert-InfraMail ($Html -notmatch '(?i)min-width') "$Label uses min-width."
    Assert-InfraMail ($Html -notmatch '(?i)1180px|<style') "$Label still carries the standalone report layout."
    $tables = @(Get-InnerDataTable -Html $Html)
    Assert-InfraMail ($tables.Count -ge 3) "$Label has too few data tables."
    foreach ($table in $tables) {
        Assert-InfraMail ($table.Groups['Attrs'].Value -match 'table-layout:fixed') "$Label data table without table-layout:fixed: $($table.Groups['Attrs'].Value)"
        $widths = @([regex]::Matches($table.Value, '<th[^>]*\swidth="(\d+)%"') | ForEach-Object { [int]$_.Groups[1].Value })
        $headers = @([regex]::Matches($table.Value, '<th\b')).Count
        Assert-InfraMail ($widths.Count -eq $headers -and ($widths | Measure-Object -Sum).Sum -eq 100) "$Label table column widths are not explicit or do not total 100%."
    }
}

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-InfraMail-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
try {
    $now = Get-Date
    $longToken = 'CN=' + ('averylongsyntheticnamespace' * 6) + '.example.invalid'
    $summary = [pscustomobject]@{
        ScriptVersion = 'synthetic'; RunId = 'SYNTHETIC-RUN'; ExecutionDate = '2026-01-02 03:04:05'; ExchangeServersCount = 3; ExcludedServers = 'EX-09'
        ComputeCollectionErrors = 0; DiskDriveCollectionErrors = 1; LogicalDiskCollectionErrors = 0; LowSpaceWarnings = 1
        TotalSocketCount = 6; TotalPhysicalCoreCount = 24; TotalLogicalProcessorCount = 48; TotalMemoryGB = 384; TotalMemoryTB = 0.38
        TotalDiskDriveCount = 9; TotalDiskDriveSizeGB = 6144; TotalDiskDriveSizeTB = 6; TotalLogicalDiskSizeTB = 5.9; TotalLogicalDiskUsedTB = 3.25; TotalLogicalDiskFreeTB = 2.65
        ExchangeSchemaRangeUpper = 'N/A'; ExchangeOrgObjectVersion = 'N/A'; ExchangeReadinessRowsCount = 0; ExchangeReadinessWarnings = 0; ExchangeReadinessErrors = 0
    }
    $servers = @(
        [pscustomobject]@{ ExchangeServerName = 'EX-01'; ServerRole = 'Mailbox'; IsVirtualMachine = $true; LogicalProcessorCount = 16; MemoryGB = 128; DiskDriveCount = 3; DiskDriveTotalSizeGB = 2048; LowSpaceLogicalDiskCount = 1; ComputeCollectionStatus = 'OK'; DiskDriveCollectionStatus = 'OK'; LogicalDiskCollectionStatus = 'OK' }
        [pscustomobject]@{ ExchangeServerName = 'EX-02'; ServerRole = 'Mailbox'; IsVirtualMachine = 'True'; LogicalProcessorCount = 16; MemoryGB = 128; DiskDriveCount = 3; DiskDriveTotalSizeGB = 2048; LowSpaceLogicalDiskCount = 0; ComputeCollectionStatus = 'OK'; DiskDriveCollectionStatus = 'ERROR'; LogicalDiskCollectionStatus = 'OK' }
        [pscustomobject]@{ ExchangeServerName = 'EX-03'; ServerRole = 'Mailbox'; IsVirtualMachine = $false; LogicalProcessorCount = 16; MemoryGB = 128; DiskDriveCount = 3; DiskDriveTotalSizeGB = 2048; LowSpaceLogicalDiskCount = 0; ComputeCollectionStatus = 'OK'; DiskDriveCollectionStatus = 'OK'; LogicalDiskCollectionStatus = 'OK' }
    )
    $soon = $now.AddDays(20).ToString([Globalization.CultureInfo]::InvariantCulture)
    $later = $now.AddDays(400).ToString([Globalization.CultureInfo]::InvariantCulture)
    $expired = $now.AddDays(-3).ToString([Globalization.CultureInfo]::InvariantCulture)
    $readiness = @(
        New-Row 'ExchangeCertificate' 'EX-01:AAAA' 'Server' 'EX-01'
        New-Row 'ExchangeCertificate' 'EX-01:AAAA' 'Subject' 'CN=<script>alert(1)</script>, O=Synthetic'
        New-Row 'ExchangeCertificate' 'EX-01:AAAA' 'Services' 'IIS, SMTP'
        New-Row 'ExchangeCertificate' 'EX-01:AAAA' 'NotAfter' $soon
        New-Row 'ExchangeCertificate' 'EX-01:BBBB' 'Server' 'EX-01'
        New-Row 'ExchangeCertificate' 'EX-01:BBBB' 'Subject' $longToken
        New-Row 'ExchangeCertificate' 'EX-01:BBBB' 'NotAfter' $expired
        New-Row 'ExchangeCertificate' 'EX-01:CCCC' 'Server' 'EX-01'
        New-Row 'ExchangeCertificate' 'EX-01:CCCC' 'NotAfter' $later
        New-Row 'ExchangeCertificate' 'EX-02' 'CollectionPath' 'RemoteRetry'
        New-Row 'ExchangeCertificate' 'EX-02:DDDD' 'Server' 'EX-02'
        New-Row 'ExchangeCertificate' 'EX-02:DDDD' 'NotAfter' $later
        New-Row 'ExchangeCertificate' 'EX-02:DDDD' 'Services' 'Unknown' 'WARNING' 'Warning' 'Property unavailable after Exchange remoting; no value was inferred.'
        New-Row 'ExchangeCertificate' 'EX-02:DDDD' 'Status' 'Unknown' 'WARNING' 'Warning' 'Property unavailable after Exchange remoting; no value was inferred.'
        New-Row 'ExchangeCertificate' 'EX-03' 'Collection' 'Failed' 'ERROR' 'Warning' 'Budget=120s; elapsed=120.004s; Local: Timeout; stage=Read; synthetic'
        New-Row 'CollectorError' 'Exchange virtual directories' 'Collection' 'Failed' 'ERROR' 'Warning' 'Synthetic collector failure.'
    )
    $newHcRow = {
        param($Server, $Status, $Message, $Vulnerabilities, $Errors, $Warnings)
        [pscustomobject]@{ ServerName = $Server; RunStatus = $Status; RunMessage = $Message; HealthCheckerVersion = '26.10.13.1644'; ExchangeVersion = 'Exchange 2016 CU23'; BuildNumber = '15.1.2507.0'
            SecurityUpdateStatus = 'Latest'; VulnerabilityCount = $Vulnerabilities; Vulnerabilities = ''; ErrorCount = $Errors; WarningCount = $Warnings }
    }
    $hcSummary = @(
        & $newHcRow 'EX-01' 'Completed' '' 2 3 5
        & $newHcRow 'EX-02' 'TaskTimeout' 'Synthetic remote task timeout.' $null $null $null
        & $newHcRow 'EX-03' 'Completed' '' 0 1 2
        & $newHcRow 'EX-09' 'NotRun' 'Health Checker excluded by configuration.' $null $null $null
    )
    $hcFindings = @(
        [pscustomobject]@{ ServerName = 'EX-01'; Category = 'Security Vulnerability'; RowNumber = 1; Check = 'Synthetic'; Item = ''; Status = 'Error'; Value = 'x'; Recommendation = '' }
        [pscustomobject]@{ ServerName = 'EX-01'; Category = 'Security'; RowNumber = 2; Check = 'Synthetic'; Item = ''; Status = 'Error'; Value = 'x'; Recommendation = '' }
        [pscustomobject]@{ ServerName = 'EX-03'; Category = 'Security'; RowNumber = 3; Check = 'Synthetic'; Item = ''; Status = 'Error'; Value = 'x'; Recommendation = '' }
        [pscustomobject]@{ ServerName = 'EX-03'; Category = 'Security'; RowNumber = 4; Check = 'Synthetic'; Item = ''; Status = 'Warning'; Value = 'x'; Recommendation = '' }
    )

    $reportPath = Join-Path $testRoot 'Exchange_OnPrem_InfrastructureAndReadiness_Report.html'
    $csvPaths = foreach ($name in @('Exchange_OnPrem_Servers_Inventory_Summary.csv', 'Exchange_OnPrem_Infrastructure_PerServerSummary.csv', 'Exchange_OnPrem_MigrationReadiness_Config.csv', 'Exchange_OnPrem_HealthChecker_Summary.csv')) {
        $path = Join-Path $testRoot $name
        Set-Content -LiteralPath $path -Value 'TenantKey;Synthetic' -Encoding ASCII
        $path
    }
    $missingPath = Join-Path $testRoot 'Exchange_OnPrem_HealthChecker_Findings.csv'

    $renderedBody = $null
    foreach ($variant in $coreVariants.Keys) {
        $module = New-TestModule -CorePath $coreVariants[$variant]
        try {
            & $module { param($s, $p, $r, $hs, $hf, $path) New-HtmlExecutiveSummary -Summary $s -PerServerSummary $p -ReadinessInventory $r -HealthCheckerSummary $hs -HealthCheckerFindings $hf -Path $path } $summary $servers $readiness $hcSummary $hcFindings $reportPath
            $reportBefore = Get-Content -LiteralPath $reportPath -Raw
            $hashBefore = (Get-FileHash -LiteralPath $reportPath -Algorithm SHA256).Hash
            Assert-InfraMail ($reportBefore.Contains('1180px') -and $reportBefore.Contains('<style')) "[$variant] Full report layout changed."
            & $module { param($path) [void]$script:ServersAndStorageSharePointUploads.Add([pscustomobject]@{ LocalFilePath = $path; FileName = 'report.html'; SharePointPath = 'SmartM365/Exchange/Report.html'; WebUrl = 'https://sharepoint.example.invalid/sites/smartm365/Report.html' }) } $reportPath

            & $module { param($path, $files, $s, $p, $r, $hs, $hf) Send-ServersAndStorageHtmlReport -HtmlReportPath $path -Subject 'SmartM365 Exchange OnPrem Servers and Storage inventory - SYNTHETIC-RUN' -Attachments $files -Summary $s -PerServerSummary $p -ReadinessInventory $r -HealthCheckerSummary $hs -HealthCheckerFindings $hf } $reportPath (@($csvPaths) + $missingPath) $summary $servers $readiness $hcSummary $hcFindings
            $mail = & $module { $script:CapturedMail }
            $body = [string]$mail.Body

            Assert-InfraMail ($mail.Subject -eq 'SmartM365 Exchange OnPrem Servers and Storage inventory - SYNTHETIC-RUN') "[$variant] Subject changed."
            Assert-InfraMail ($body.Contains('<!-- SmartM365EmailTemplate:v1 -->') -and $body.Contains('width="760"') -and $body.Contains('background:#0f172a')) "[$variant] Body does not use the shared template."
            Assert-InfraMail ($body.Contains('SmartM365 Exchange OnPrem') -and $body.Contains('Infrastructure and readiness summary')) "[$variant] Category or title missing."
            Assert-InfraMail ($body.Contains('Completed with errors') -and $body.Contains('#fef2f2')) "[$variant] Error severity not derived from collection errors."
            Assert-InfraMail ($mail.Bound -notcontains 'Attachments') "[$variant] Files passed to Core would be appended outside the 760px frame."
            Assert-FrameWidth -Html $body -Label "[$variant] mail"

            foreach ($expected in @('Key figures', 'Virtual machines', 'vCPU', '384 GB', '6.00 TB', '3.25 TB used', 'Health Checker', '2/3', '2 vulnerabilities, 3 errors',
                'Per-server summary', 'EX-09', 'Excluded', 'ERROR: disks', 'Low space', '1 (RemoteRetry)', 'Failed', 'TaskTimeout',
                'Collection issues', 'Collection failed', 'Budget=120s', 'RemoteRetry', 'Unknown metadata', 'Services, Status unavailable', 'Collector failed', 'Synthetic remote task timeout.',
                'critical findings (Status=Error): 3', 'Excluded by configuration: 1 server(s)', '4 certificate(s) read; 2 expiring within 60 days or expired; 1 server(s) failed; 1 RemoteRetry', 'per-server budget 120 s', 'Expired', '>20<',
                'Files', 'Exchange_OnPrem_InfrastructureAndReadiness_Report.html', 'href="https://sharepoint.example.invalid/sites/smartm365/Report.html"', 'Exchange_OnPrem_HealthChecker_Summary.csv')) {
                Assert-InfraMail ($body.Contains($expected)) "[$variant] Missing mail content: $expected"
            }
            Assert-InfraMail (-not $body.Contains('Exchange_OnPrem_HealthChecker_Findings.csv')) "[$variant] Missing file listed."
            Assert-InfraMail (-not $body.Contains('<script>') -and $body.Contains('&lt;script&gt;')) "[$variant] Certificate subject not escaped."
            Assert-InfraMail ($body.IndexOf('Exchange_OnPrem_InfrastructureAndReadiness_Report.html') -lt $body.IndexOf('This automated message was generated by SmartM365')) "[$variant] Files section is outside the frame."
            Assert-InfraMail ($body -match 'word-break:break-all;overflow-wrap:anywhere;') "[$variant] Long columns cannot break."
            Assert-InfraMail ($body.IndexOf('Expired') -lt $body.IndexOf('>20<')) "[$variant] Expiring certificates not sorted by remaining days."
            Assert-InfraMail (-not $body.Contains($longToken)) "[$variant] Long certificate subject not shortened for the mail."
            Assert-InfraMail ((Get-FileHash -LiteralPath $reportPath -Algorithm SHA256).Hash -eq $hashBefore) "[$variant] Full report modified by mail generation."

            # Health Checker disabled and clean run.
            $cleanSummary = $summary.PSObject.Copy(); $cleanSummary.DiskDriveCollectionErrors = 0; $cleanSummary.LowSpaceWarnings = 0; $cleanSummary.ExcludedServers = ''
            $cleanServers = @($servers | ForEach-Object { $copy = $_.PSObject.Copy(); $copy.DiskDriveCollectionStatus = 'OK'; $copy.LowSpaceLogicalDiskCount = 0; $copy })
            $cleanRows = @($readiness | Where-Object { $_.ObjectName -like 'EX-01:CCCC' })
            $clean = & $module { param($s, $p, $r) New-ServersAndStorageMailBodyHtml -Summary $s -PerServerSummary $p -ReadinessInventory $r -HealthCheckerSummary @() -HealthCheckerFindings @() -Files @() } $cleanSummary $cleanServers $cleanRows
            Assert-InfraMail ($clean.Contains('>Completed<') -and -not $clean.Contains('Completed with') -and $clean.Contains('#ecfdf5')) "[$variant] Clean run is not Success."
            Assert-InfraMail ($clean.Contains('Health Checker was not run in this run') -and $clean.Contains('Not run') -and $clean.Contains('disabled for this run')) "[$variant] Disabled Health Checker not explicit."
            Assert-InfraMail ($clean.Contains('No collection failure was detected') -and $clean.Contains('1 certificate(s) read; 0 expiring') -and -not $clean.Contains('Expired')) "[$variant] Clean run shows issues."
            $warning = & $module { param($s, $p, $r, $hs) New-ServersAndStorageMailBodyHtml -Summary $s -PerServerSummary $p -ReadinessInventory $r -HealthCheckerSummary $hs -Files @() } $cleanSummary $cleanServers $cleanRows @($hcSummary[0])
            Assert-InfraMail ($warning.Contains('Completed with warnings')) "[$variant] Vulnerabilities do not raise a warning."
            $noData = & $module { param($s) New-ServersAndStorageMailBodyHtml -Summary $s } $cleanSummary
            Assert-InfraMail ($noData.Contains('No per-server data is available') -and $noData.Contains('no certificate data') -and $noData.Contains('No output file is available')) "[$variant] Empty inputs not handled."
            $csvSummary = [pscustomobject]@{ RunId = 'R'; ExecutionDate = 'D'; ExchangeServersCount = '3'; TotalLogicalProcessorCount = '48'; TotalMemoryGB = '384.6'; TotalDiskDriveCount = '9'; TotalDiskDriveSizeTB = '6,5'; LowSpaceWarnings = '0'; ExcludedServers = '' }
            $csvBody = & $module { param($s) New-ServersAndStorageMailBodyHtml -Summary $s } $csvSummary
            Assert-InfraMail ($csvBody.Contains('385 GB') -and $csvBody.Contains('>48<')) "[$variant] CSV text numbers not parsed."

            if ($variant -eq 'Core') { $renderedBody = $body }
        }
        finally {
            Remove-Module $module -Force
        }
    }
    if ($RenderPath) {
        [IO.File]::WriteAllText($RenderPath, $renderedBody, (New-Object Text.UTF8Encoding($false)))
        Write-Host "Synthetic mail written to: $RenderPath"
    }
    Write-Host "PASS: $script:Checks synthetic Exchange infrastructure mail checks (PowerShell $($PSVersionTable.PSVersion)). No collector or real mail executed."
}
finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SmartM365-InfraMail-*') { throw 'Unsafe test cleanup path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZsvBKfuKNm8wg
# E7NKXhD12M0YaEtWfjSof140AWPdIaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgzjTbazQli5tZ
# P/w5neB9MLLjMEOIFqoTyp9IQjHDNCcwDQYJKoZIhvcNAQEBBQAEggGABWDt1vXT
# x0XxJgzJq9+QeomvoAeG4zevPYD6cR7FTX8Y1rwydJs4zp5JhjmjKX9aAKbphWRW
# i1CXOKKjZTJDDg/xsXEaTuRwb40GhhGEYD8hCw0C7uyDbgnum9y9vks+f7fbhUmj
# 0VLJgrLp4x2xb5zMEo6M54oAuNJA8sDjGQVgnJi5Drxa1OXutym4VjK4oEuiLHpJ
# amIS7pTI5z/V6Bpc+gnduWyhoP7igfgiT10hON5mlSiVAYA3vR1uDMhTKgITqAiV
# tutK+kea+yTc2NWxegNHr0Z/vE4SaaJ/f0CaRLzlum32VTefj847aZgBtafGBh9T
# uEo9pwPPURhsYsKsNZv4ONFD4VsazQcBL/EKciE7k9Iy4VlTZczBeVi+L4XePvQE
# 09p786nFtDxXvDhcRfYyPoXfMg00ivsdzFDN8sWlmyHgfluP5LF0+6Kiakj5bBuf
# 9KTrw21u+XeyfEQsAPv2uKGkXmqhNC2NuLmi/eoBLyfdTNyovL2PjlXloYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNTE1WjA/BgkqhkiG9w0BCQQxMgQwAM9ch1hMRDsCek40qg6VCtTz
# kTgTwaaxcTjmmSbgqp4Gp5O3pIrprSwuqNosDVB/MA0GCSqGSIb3DQEBAQUABIIC
# AFSVMfst1O09VsCNdfeMMoOzySz6K5/f4PnWp72TvubnSzckkMqY/haJyflDfKQQ
# JcTmNxZCy6hJYZzThShsLyC03vd0rDw/QSBBa918B/8s1f4SXsEVNOdqEo5Ap27n
# 0Twzi34epsPUFk7MBU9Rdzua3KIspH7m5WaTNzAC/CYxmTQ3ZIZCN1v4Umde0Tbu
# /P7IziGrjWaVpdbI8h4GG2pQRqjmkVGWTrfoOzjMXMkZWWTO6oX0YhUQSph3E/DK
# KJBYUkgQtEAGKMrCmOjB9BO02TzqKtYDlQJqAxwAca8Jdsw0cKc6pOep4lvkbcW8
# k4UvYQCKKrNjUeDu2tdIsW4VtLsuzk5/HzFxdyOoDY2pmQtAaWhJdNMTAuMRb9Zt
# w6ZgtduadKmCgN33KDa1XOZYeRudlzR3rYlDvRpTg2FaoxEpldtUK1gamMZzon0c
# lidPvt7AETYaGjnJ2Tq2hJsSRWkR4JVDoL/Nn3+AAGxLCL6jxhOZNe0G2NEKta8P
# 2UsPTcEtsAd5dRPNXcnb3Af5/2Sn/XJoarho+1YgTA3i/9cWuxrS+5T9pZpxwLxy
# AAPUgiHeR9G2fL1uoruVt72KCu7VMMW7XX4/YCP8DFGrID7ttXgIavbIOgeheZw1
# fnr349cGdDkY3Bt9MLuiutTGnnStjPZt+Igx5dbXiDY4
# SIG # End signature block
