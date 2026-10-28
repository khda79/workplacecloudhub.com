# Version 1.0.0
[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-OfflineTests'))
# Synthetic local files only. No collectors, tenant calls, production data or uploads.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
$script:checks = 0
function Check([bool]$Pass, [string]$Message) { if (-not $Pass) { throw $Message }; $script:checks++ }

$product = Split-Path -Parent $PSScriptRoot
$sourceNames = @('M365_Email_AppUsage.csv', 'M365_OneDrive_UserActivity.csv', 'M365_Groups_Activity.csv', 'M365_Copilot_UserCountTrend.csv')
$tableNames = @('Email Client Usage Evidence', 'OneDrive Activity Evidence', 'Group Activity Evidence', 'Copilot Daily Adoption Evidence')
$now = [datetime]::new(2026, 10, 28, 12, 0, 0, [DateTimeKind]::Utc)
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$raw = Join-Path $fixture 'raw'
$last = Join-Path $raw 'DATA-LAST'
$output = Join-Path $fixture 'output'
$shell = (Get-Process -Id $PID).Path

function Write-Source([string]$Name, [object[]]$Rows) {
    $Rows | Export-Csv -LiteralPath (Join-Path $last $Name) -NoTypeInformation -Encoding utf8NoBOM
}
function Invoke-Signals {
    & $shell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'New-M365UsageSignalsEvidence.ps1') -DataRoot $raw `
        -EmailClientOutputPath (Join-Path $output 'EmailClientUsageEvidence.csv') -OneDriveActivityOutputPath (Join-Path $output 'OneDriveActivityEvidence.csv') `
        -GroupActivityOutputPath (Join-Path $output 'GroupActivityEvidence.csv') -CopilotDailyOutputPath (Join-Path $output 'CopilotDailyAdoptionEvidence.csv') `
        -AsOfUtc $now.ToString('o') | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Usage signals generator failed with exit code $LASTEXITCODE." }
    @{
        Email = @(Import-Csv -LiteralPath (Join-Path $output 'EmailClientUsageEvidence.csv'))
        OneDrive = @(Import-Csv -LiteralPath (Join-Path $output 'OneDriveActivityEvidence.csv'))
        Groups = @(Import-Csv -LiteralPath (Join-Path $output 'GroupActivityEvidence.csv'))
        Copilot = @(Import-Csv -LiteralPath (Join-Path $output 'CopilotDailyAdoptionEvidence.csv'))
    }
}
function Get-Trust {
    $path = Join-Path $fixture 'DataTrustEvidence.csv'
    & $shell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'New-DataTrustEvidence.ps1') -DataRoot $raw -OutputPath $path -AsOfDate $now.Date -AsOfUtc $now.ToString('yyyy-MM-ddTHH:mm:ss') | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Data Trust failed with exit code $LASTEXITCODE." }
    $rows = @(Import-Csv -LiteralPath $path | Where-Object { $_.'File Name' -in $sourceNames })
    Check ($rows.Count -eq 4) 'Data Trust must carry exactly one row per optional usage report'
    $rows
}
function Test-Outputs {
    # Publication shape: TenantKey prefix, then the governed contract validator.
    $root = Join-Path $fixture ('validation-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $root | Out-Null
    foreach ($table in $script:evidenceTables) {
        $source = Join-Path $output $table.file
        $target = Join-Path $root $table.file
        $rows = @(Import-Csv -LiteralPath $source)
        if ($rows.Count) { $rows | Select-Object @{n = 'TenantKey'; e = { 'synthetic-test' } }, * | Export-Csv -LiteralPath $target -NoTypeInformation -Encoding utf8 }
        else { [IO.File]::WriteAllText($target, '"TenantKey",' + (Get-Content -LiteralPath $source -TotalCount 1) + [Environment]::NewLine) }
    }
    # -File cannot pass a string array to the child; -Command keeps the table filter.
    $command = "& '{0}' -Root '{1}' -ContractPath '{2}' -Tables @({3}) -TenantKey 'synthetic-test'; exit `$LASTEXITCODE" -f (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1'), $root, (Join-Path $product 'config/prepared-evidence-contract.json.txt'), (($tableNames | ForEach-Object { "'$_'" }) -join ',')
    $report = & $shell -NoProfile -Command $command | Out-String
    Check ($LASTEXITCODE -eq 0) "Generated outputs violate the evidence contract: $report"
}

try {
    New-Item -ItemType Directory -Path $last, $output -Force | Out-Null
    # Data Trust qualifies device identity from the Intune inventory; a legacy minimal one is enough here.
    Write-Source 'Intune_Devices_Inventory.csv' @([pscustomobject]@{ TenantKey = 'synthetic-test'; 'Device ID' = 'device-1'; 'Device name' = 'CONTOSO-PC1' })

    # Contracts: optional at 744 hours, never required, never a history table.
    $sourceContract = Get-Content -LiteralPath (Join-Path $product 'config/prepared-source-contract.json.txt') -Raw | ConvertFrom-Json
    foreach ($name in $sourceNames) {
        $entry = @($sourceContract.optionalCurrentFiles | Where-Object file -CEQ $name)
        Check ($entry.Count -eq 1 -and [int]$entry[0].maxAgeHours -eq 744) "$name must be optional with a 744-hour limit"
        Check ($name -notin @($sourceContract.currentFiles)) "$name must not be a required current file"
    }
    $evidenceContract = Get-Content -LiteralPath (Join-Path $product 'config/prepared-evidence-contract.json.txt') -Raw | ConvertFrom-Json
    $script:evidenceTables = @($evidenceContract.tables | Where-Object { $_.table -in $tableNames })
    Check ($script:evidenceTables.Count -eq 4) 'Every usage signal table must be declared once in the evidence contract'
    Check (@($script:evidenceTables | Where-Object { $_.table -match 'History|Trend' }).Count -eq 0) 'A rolling Graph window must not become a history-coverage table'
    Check (@($script:evidenceTables | Where-Object { $_.PSObject.Properties['allowEmpty'] -and $_.allowEmpty }).table -join '|' -eq 'Group Activity Evidence|Copilot Daily Adoption Evidence') 'Only per-object and per-date tables may be empty'
    $build = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Invoke-PreparedEvidenceBuild.ps1') -Raw
    Check ($build -match "Add-Task 'M365UsageSignals' 'Collaboration' 'New-M365UsageSignalsEvidence\.ps1'") 'Build does not run the usage signals generator'

    # Selection: absent optional reports are diagnostics, never a blocked plan.
    $planContract = Join-Path $fixture 'prepared-source-contract.json.txt'
    @{ currentFiles = @(); optionalCurrentFiles = @($sourceContract.optionalCurrentFiles | Where-Object { $_.file -in $sourceNames }); mappingFiles = @(); dailyFiles = @(); history = @() } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $planContract
    $module = Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
    & $module { param($root) $script:ProductRoot = $root } $fixture
    $issues = [Collections.Generic.List[object]]::new()
    $plan = @(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $planContract -MetadataOnly -OptionalSourceIssues $issues)
    Check ($plan.Count -eq 0 -and $issues.Count -eq 4) 'Absent optional usage reports blocked selection or lost diagnostics'

    # 1. Never collected: stable dimension rows say Not collected, never zero.
    $absent = Invoke-Signals
    Check ($absent.Email.Count -eq 9 -and @($absent.Email | Where-Object { $_.'Evidence Status' -ne 'Not collected' -or $_.Users -ne '' }).Count -eq 0) 'Absent email report must give nine Not collected clients without counts'
    Check (@($absent.Email | Where-Object 'Legacy Protocol' -EQ 'Yes').Client -join ',' -eq 'POP3,IMAP4,SMTP') 'Legacy protocol flags changed'
    Check ($absent.OneDrive.Count -eq 5 -and @($absent.OneDrive | Where-Object { $_.'Evidence Status' -ne 'Not collected' -or $_.Users -ne '' }).Count -eq 0) 'Absent OneDrive report must give five Not collected metrics'
    Check ($absent.Groups.Count -eq 0 -and $absent.Copilot.Count -eq 0) 'Absent group and Copilot reports must publish schema only'
    Test-Outputs
    foreach ($row in Get-Trust) {
        Check ($row.'File State' -eq 'Missing' -and $row.'Decision Readiness' -eq 'Optional unavailable' -and $row.'Evidence Status' -eq 'Not collected' -and $row.'Freshness State' -eq 'Unknown') "Absent $($row.'File Name') must be Optional unavailable, not Blocked"
    }

    # 2. Kept empty by Microsoft (header only): same as not collected.
    foreach ($name in $sourceNames) { [IO.File]::WriteAllText((Join-Path $last $name), '"TenantKey","Report Refresh Date"' + [Environment]::NewLine) }
    $empty = Invoke-Signals
    Check (@($empty.Email | Where-Object 'Evidence Status' -EQ 'Not collected').Count -eq 9 -and $empty.Groups.Count -eq 0) 'Header-only reports must stay Not collected'
    foreach ($row in Get-Trust) {
        Check ($row.'File State' -eq 'Unexpected empty' -and $row.'Decision Readiness' -eq 'Optional unavailable') "Empty $($row.'File Name') must be Optional unavailable"
    }

    # 3. Populated reports.
    $refresh = '2026-10-26'
    $mail = { param($Upn, $Deleted, $Flags) $row = [ordered]@{ TenantKey = 'synthetic-test'; 'Report Refresh Date' = $refresh; 'User Principal Name' = $Upn; 'Display Name' = 'Contoso user'; 'Is Deleted' = $Deleted; 'Deleted Date' = ''; 'Last Activity Date' = $refresh }
        foreach ($column in @('Mail For Mac', 'Outlook For Mac', 'Outlook For Windows', 'Outlook For Mobile', 'Other For Mobile', 'Outlook For Web', 'POP3 App', 'IMAP4 App', 'SMTP App')) { $row[$column] = if ($column -in $Flags) { 'Yes' } else { '' } }
        $row['Report Period'] = '180'; [pscustomobject]$row }
    Write-Source 'M365_Email_AppUsage.csv' @(
        (& $mail 'a@contoso.example' 'False' @('Outlook For Windows', 'Outlook For Web')),
        (& $mail 'b@contoso.example' 'False' @('Outlook For Mobile', 'IMAP4 App')),
        (& $mail 'c@contoso.example' 'False' @()),
        (& $mail 'd@contoso.example' 'True' @('POP3 App'))
    )
    $drive = { param($Upn, $Deleted, $Last, $Viewed, $Synced, $Internal, $External) [pscustomobject][ordered]@{ TenantKey = 'synthetic-test'; 'Report Refresh Date' = $refresh; 'User Principal Name' = $Upn; 'Is Deleted' = $Deleted; 'Deleted Date' = ''; 'Last Activity Date' = $Last
        'Viewed Or Edited File Count' = $Viewed; 'Synced File Count' = $Synced; 'Shared Internally File Count' = $Internal; 'Shared Externally File Count' = $External; 'Assigned Products' = 'MICROSOFT 365 E3'; 'Report Period' = '180' } }
    Write-Source 'M365_OneDrive_UserActivity.csv' @(
        (& $drive 'a@contoso.example' 'False' '2026-10-20' '12' '3' '2' '1'),
        (& $drive 'b@contoso.example' 'False' '2026-09-01' '4' '0' '0' '0'),
        (& $drive 'c@contoso.example' 'False' '' '0' '0' '0' '0'),
        (& $drive 'd@contoso.example' 'True' '2026-10-20' '99' '99' '99' '99')
    )
    $group = { param($Id, $Name, $Deleted, $Owner, $Last, $External) [pscustomobject][ordered]@{ TenantKey = 'synthetic-test'; 'Report Refresh Date' = $refresh; 'Group Display Name' = $Name; 'Is Deleted' = $Deleted; 'Owner Principal Name' = $Owner; 'Last Activity Date' = $Last
        'Group Type' = 'Private'; 'Member Count' = '10'; 'External Member Count' = $External; 'Exchange Received Email Count' = '5'; 'SharePoint Active File Count' = '2'; 'Yammer Posted Message Count' = '0'; 'Yammer Read Message Count' = '0'; 'Yammer Liked Message Count' = '0'
        'Exchange Mailbox Total Item Count' = '50'; 'Exchange Mailbox Storage Used (Byte)' = [string](1GB); 'SharePoint Total File Count' = '20'; 'SharePoint Site Storage Used (Byte)' = [string](2GB); 'Group Id' = $Id; 'Report Period' = '180' } }
    Write-Source 'M365_Groups_Activity.csv' @(
        (& $group 'g-team' 'Contoso Team' 'False' 'owner@contoso.example' '2026-10-25' '0'),
        (& $group 'g-orphan' 'Contoso Orphan' 'False' '' '2026-10-25' '0'),
        (& $group 'g-old' 'Contoso Archive' 'False' 'owner@contoso.example' '2025-01-01' '3'),
        (& $group 'g-old' 'Contoso Archive duplicate' 'False' 'owner@contoso.example' '2025-01-01' '3'),
        (& $group 'g-gone' 'Contoso Deleted' 'True' '' '' '0')
    )
    Write-Source 'M365_Teams_Teams.csv' @([pscustomobject]@{ TenantKey = 'synthetic-test'; TeamId = 'g-team'; TeamDisplayName = 'Contoso Team'; RunDateUtc = $now.ToString('o') })
    $copilotColumns = @('TenantKey', 'RunId', 'ReportVersion', 'ReportPeriod', 'ReportRefreshDate', 'ReportDate')
    foreach ($prefix in @('AnyApp', 'Microsoft365Copilot', 'Teams', 'Word', 'Excel', 'PowerPoint', 'Outlook', 'OneNote', 'Loop', 'Edge', 'CopilotChat', 'CopilotChatWork', 'CopilotChatWeb')) { $copilotColumns += @("${prefix}EnabledUsers", "${prefix}ActiveUsers") }
    $copilotColumns += 'PromptsSubmitted'
    $day = { param($Date, $Version, $Values) $row = [ordered]@{}; foreach ($column in $copilotColumns) { $row[$column] = '' }
        $row.TenantKey = 'synthetic-test'; $row.RunId = 'run-1'; $row.ReportVersion = $Version; $row.ReportPeriod = 'D28'; $row.ReportRefreshDate = $refresh; $row.ReportDate = $Date
        foreach ($key in $Values.Keys) { $row[$key] = $Values[$key] }; [pscustomobject]$row }
    Write-Source 'M365_Copilot_UserCountTrend.csv' @(
        (& $day '2026-10-25' 'v2' @{ AnyAppEnabledUsers = '40'; AnyAppActiveUsers = '10'; TeamsEnabledUsers = '40'; TeamsActiveUsers = '4'; EdgeEnabledUsers = '0'; EdgeActiveUsers = '0'; PromptsSubmitted = '120' }),
        (& $day '2026-10-26' 'v2' @{ AnyAppEnabledUsers = '40'; AnyAppActiveUsers = '12'; TeamsEnabledUsers = '40'; TeamsActiveUsers = '5'; PromptsSubmitted = '150' })
    )
    $observed = Invoke-Signals
    $email = @{}; foreach ($row in $observed.Email) { $email[$row.Client] = $row }
    Check ($email['Outlook for Windows'].Users -eq '1' -and $email['IMAP4'].Users -eq '1' -and $email['POP3'].Users -eq '0') 'Email client counts are wrong or include deleted users'
    Check ($email['IMAP4'].'Email Client Users' -eq '2' -and [double]$email['IMAP4'].'User Share (%)' -eq 0.5) 'Email share must use users with at least one client'
    Check ($email['SMTP'].'Report Period Days' -eq '180' -and $email['SMTP'].'Usage Refresh Date' -eq $refresh -and $email['SMTP'].'Evidence Age Days' -eq '2' -and $email['SMTP'].'Freshness State' -eq 'Current' -and $email['SMTP'].'Evidence Status' -eq 'Observed') 'Email report freshness is wrong'
    $drive = @{}; foreach ($row in $observed.OneDrive) { $drive[$row.Metric] = $row }
    Check ($drive['Viewed or edited files'].Users -eq '2' -and $drive['Viewed or edited files'].Files -eq '16' -and $drive['Shared files externally'].Users -eq '1') 'OneDrive activity aggregates are wrong or include deleted users'
    Check ($drive['Any OneDrive activity'].Users -eq '2' -and $drive['Any OneDrive activity'].Files -eq '' -and $drive['Any OneDrive activity'].'Population Users' -eq '3') 'OneDrive population or activity count is wrong'
    $groups = @{}; foreach ($row in $observed.Groups) { $groups[$row.'Group ID'] = $row }
    Check ($observed.Groups.Count -eq 3 -and -not $groups.ContainsKey('g-gone')) 'Deleted or duplicate groups must be left out'
    Check ($groups['g-team'].'Teams State' -eq 'Teams-enabled' -and $groups['g-team'].'Risk Signal' -eq 'No current signal' -and $groups['g-team'].'Storage Used GB' -eq '3.0') 'Teams-enabled group or storage is wrong'
    Check ($groups['g-orphan'].'Teams State' -eq 'Group only' -and $groups['g-orphan'].'Risk Priority' -eq 'High' -and $groups['g-orphan'].'Risk Signal' -eq 'Group without owner') 'Group without owner was not prioritized'
    Check ($groups['g-old'].'Activity State' -eq 'Inactive 180D+' -and $groups['g-old'].'Risk Signal' -eq 'Inactive group' -and $groups['g-old'].'External Access State' -eq 'External members present' -and $groups['g-old'].'Group Name' -eq 'Contoso Archive') 'Inactive group signal is wrong'
    Check (@($observed.Copilot | Where-Object App -EQ 'Edge').Count -eq 1 -and @($observed.Copilot | Where-Object App -EQ 'Word').Count -eq 0) 'Unreported apps must leave no row; reported zeros must stay'
    $any = @($observed.Copilot | Where-Object { $_.App -eq 'Any app' -and $_.'Date Key' -eq '20261026' })
    Check ($any.Count -eq 1 -and [double]$any[0].'Active User Rate (%)' -eq 0.3 -and $any[0].'Prompts Submitted' -eq '150' -and $any[0].'Report Version' -eq 'v2') 'Copilot daily any-app row is wrong'
    Check (@($observed.Copilot | Where-Object { $_.App -ne 'Any app' -and $_.'Prompts Submitted' -ne '' }).Count -eq 0) 'Prompts belong to the any-app row only'
    Test-Outputs
    foreach ($row in Get-Trust) {
        Check ($row.'File State' -eq 'Populated' -and $row.'Decision Readiness' -eq 'Ready' -and $row.'Evidence Status' -eq 'Observed' -and $row.'Snapshot Date' -eq $refresh) "Populated $($row.'File Name') must be Ready"
    }

    # 4. Schema drift: the source exists but cannot be trusted.
    Write-Source 'M365_Email_AppUsage.csv' @(Import-Csv -LiteralPath (Join-Path $last 'M365_Email_AppUsage.csv') | Select-Object * -ExcludeProperty 'SMTP App')
    $drift = Invoke-Signals
    Check (@($drift.Email | Where-Object { $_.'Evidence Status' -ne 'Schema mismatch' -or $_.Users -ne '' }).Count -eq 0) 'Drifted email report must be Schema mismatch without counts'
    Check (@($drift.OneDrive | Where-Object 'Evidence Status' -EQ 'Observed').Count -eq 5) 'Other reports must stay observed'
    Test-Outputs

    Write-Host "M365 usage signals evidence offline test passed ($script:checks checks)."
}
finally {
    Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAIUvs+dfO350CZ
# AO0wCNdowgdLAadfG5diJSBV/mCjB6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIO3WtiYwEZNHJXmUp+9STme8E+NcGtsWtEAlDv4pGeNjMA0GCSqG
# SIb3DQEBAQUABIIBgFMZEnv8+XpkFmB9Kx2gJ7H8PpU5R+OXk2oWS36lFEqfB1uE
# V3ucyxMg9p8XzgnQtMF6BiPuba520gTHG2FwYdP2U+Xgnkfwkh0m3RMXACMXsY62
# 1hAY6rs9viFG0kJl9dl9v/M12AyMQ11J1jX+xJN5FISvW1O5KZPCzRJsdYKm+BFU
# fU2/90T+9WRZHquSad56KY8pD/JQfDuoAWGJamvn+OdIpMLAUcMy/6FUoWr4W1EN
# eOD831ETBOpRqF57OOCOKdt0WKc0a58CfJXOw51uTflP8EMw63dCkXUYQPZs8FPJ
# gOgFPRWLl5Qun6Reod6c1kcR13HCJTMqvMd7LtsfVExbJDesc8Q6NVL8uxsT8TpO
# V2Yf6WcwrJ8ac+XfZ8e30+HzohKuU2rYQUNaERza6UUv5VbzLgoXrjBvZsK2ei5s
# fz3PpfWSJXzpLsaW/XwIUUXLm/S8s6TpgEYWHhOlCimN5+nDVV/kv2uwgHCarjBz
# R7dm3jinh6qX2VXRBaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgwOTE4
# MDZaMC8GCSqGSIb3DQEJBDEiBCD+vB00e/RL0G6fMgkitZGHdy/1anM7UMMeoiSB
# nBmttDANBgkqhkiG9w0BAQEFAASCAgBLx7kBK49vEwOJMGAlFmwA+UHBclVoUG2e
# UquzfwbInVLLR85rviySkxIJx76/9KDCgyY84bI1+CnGbuValytghOs15YHYbD9w
# G/TySAv6iFZ8M655bLvpyikwjOTjJPEn1WS7bPYMWt/FUlsjy7yFnP/uiqgK09E/
# F2+JdVrdtAZxnGCgBpGpzXsWgcYIizhDaWj+wfYh7RUPmDVo/PRA/vC6J7wW/C9z
# 7KO/51lY9XSh2VO7Fdg2kbL0UkOWHr0P3LMR5z4wCeM7SAbGFYuw2MitwryfxHuJ
# RsOAkfTuC78WL1Dk+7rkEpGjJcc0wRq2NN+wupN221dnQR3IOKPQR0nWBCKiFSQQ
# 1i+CQgStqXL8zN2GvmrF9T6UZMEBQqSux8caaoyF646S4SU9H+cVo58ZVIHIQX1+
# G5as1nKIbiKld3p5iLGXwCG5tgTl2GJQiCTBWAbZ4jJU99q1m22dTFe9S7rWoti7
# b2t7SOggT2Yq48SD4TR5RMKDd2/VOPaw9R2eNhT/c87pUg+AAv7jGWAQKLDEZmaM
# t/P8D9mqsNZMS2bgLQOYXvnLSh04mIlQQHt2ukoQYrgwp+FIvF0nA//J8HCjNQN6
# LqJSygAZWE4DJzQO6o0RN7dO+/ek3FK0N4eRMzFRdaOyKlPKKecUb9E7CE7PXRbd
# QCjlXCspEQ==
# SIG # End signature block
