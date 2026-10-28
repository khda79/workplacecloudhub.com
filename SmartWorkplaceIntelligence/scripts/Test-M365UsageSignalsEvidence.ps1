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
