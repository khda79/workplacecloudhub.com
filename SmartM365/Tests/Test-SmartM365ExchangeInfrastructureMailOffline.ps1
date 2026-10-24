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

function Get-FunctionDefinition {
    param([string]$Path, [string[]]$Names)
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Parse failed: $Path" }
    foreach ($name in $Names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAxOZUmBPU0NKou
# R2ynQmIBFEP+aw2gabYbzRz9cijSwaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDOZ+sXbBqwOuZTnAbNXVzl0GPvSP7vMBfHKQaJdhylpMA0GCSqG
# SIb3DQEBAQUABIIBgDloQ1TCj6K+cp5OfayeWRy3h5DhoIPXw0jqCJlUJ2oiM6fR
# 98FnF5o4lG4Uqko4dzEH3nP/pyxwz+t6CG+aLCB4sYyvEqv8iw5q5C7oc48Dvtmd
# Qfmm4hyD1P9QubZMYRm7Z2Bs0h7a3f0lBXauKWqoF4fBLtjkrtyPv3h+jXLoG//Y
# O4PqMOKCBp6KekFFzQtfM/LE4e/PD7MKOmH9SD2a4F8KTLTJMRejBoIHSMOt6ZNQ
# JSwjnPhmscHQjYLt8qR1Gtu6TV1bZxbujU/LFYBztJzGDJP9cyX9e1iUZ7yPo59/
# csTzMgEGkxJbtivBDGZ49b2IpLCLCR1nYnqh/ngkcy52BVnkLDgop7ieCUvYmv8H
# 7uNN1MEBs24bv0Nzdv2vMKtMvHnUyja6oZP46pcgWwaPNpeirgOvv5mCL6Ay/wBf
# I8T8D+kzAE8NViiDUN97tJEPkwB8i8+zxs41r+cRtI1JnM+8XfBpBbi994hDlbJR
# z0zkzbh/Nzf5cPF2x6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjQxMzU5
# MzdaMC8GCSqGSIb3DQEJBDEiBCACzSm4JYqOd288mf3BIKs1h6/QvXu0T3erEYKh
# 8zdEwjANBgkqhkiG9w0BAQEFAASCAgAxuQczjDmP8oRwk67O3SqCeBGhMasnjmYA
# UzB3+2/YYeoO4iZ3kBP70DJ+2cEQAElgDAGA36GR1x9OgAH119wcnDMpP6nZX/k1
# oiAZ+jd8qdcWca2D/Q5pw2gi+UTvDc8+uQO9UiIS+TAS2srMr+4rtA/6cQ9F9d1b
# qja3+7O1yPXq1PBWJzCgA8ZSfaJiwtgHnBfOxshad3MJQeKWalYJ51Fa/XUO9aDh
# 1oimwxShsxgoEH3PzjHcmIgMvuGcmz36WRjy9WCtZkPLTLwJ7NkBIazX0vEtH8kh
# BbzfMJ4NkORsLRcHWmqqj3kqUjT/aXKmA2KLmAXwRYGRt1HggrRJQUpWN0D2oeOV
# MzzPTa3BdJ46HYFJKV9x1dwWw1CTrrnDW4xxWRFFm+D4tGFtgx94Spah6ANt8rSD
# +P3oBAh4FVtgGNDzbRuV3bXuo4bGm+zdFEyDiPDDMReH3VnXTjDFM4sazWd7wI/V
# HRc6N5qlExuZSstS37w6z/ve6maYb3PikVYZEJXFlBTCD6QSvFFbNMekC35vEuMr
# yfVF4DEc20tkVsyUJbZYOgIetj6Oh1daabPSDDMhxjema69PKW4m4RJ7aV0edus9
# yDrw9Fv+xEnElukW7PQa3y+Ku8PAzARPZ0dNgp9bWBvyf7+VKTu+VCQYYAW1wGn4
# JAclNK326Q==
# SIG # End signature block
