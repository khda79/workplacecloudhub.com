[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-OfflineTests'))

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$last = Join-Path $fixture 'DATA-LAST'
$contract = Join-Path $fixture 'prepared-source-contract.json.txt'
$manualNames = @(
    'M365_Backup_ProtectedMailboxes.csv',
    'M365_Backup_ProtectedSharePointSites.csv',
    'M365_Backup_ProtectedOneDriveAccounts.csv'
)
$policyNames = @('M365_BackupPolicyScope_MailboxCoverage.csv', 'M365_BackupPolicyScope_GroupMembers.csv')
$now = [datetime]::UtcNow
$checks = 0

function Check([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}

function Get-Plan {
    param([Collections.Generic.List[object]]$Issues)
    @(Get-PreparedSourcePlan -DataRoot $fixture -SourceContractPath $contract -MetadataOnly -OptionalSourceIssues $Issues)
}

function Run-Generator {
    param([string]$Output)
    & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'New-BackupResilienceEvidence.ps1') -DataRoot $last -OutputRoot $Output -AsOfDate $now.Date | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Backup generator failed with exit code $LASTEXITCODE." }
}

try {
    New-Item -ItemType Directory -Path $last -Force | Out-Null
    @{
        currentFiles = $policyNames
        optionalCurrentFiles = @($manualNames | ForEach-Object { @{ file = $_; maxAgeHours = 168 } })
        mappingFiles = @()
        dailyFiles = @()
        history = @()
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $contract

    @([pscustomobject]@{
        TenantKey = 'synthetic-test'; MemberId = 'user@example.invalid';
        PrimarySmtpAddress = 'user@example.invalid'; MailboxFoundInInventory = 'true';
        RunDateUtc = $now.ToString('O'); GroupDisplayName = 'Backup policy';
        MailboxUserPrincipalName = 'user@example.invalid'; MemberUserPrincipalName = 'user@example.invalid';
        MemberMail = 'user@example.invalid'; MailboxDisplayName = 'Example User';
        MemberDisplayName = 'Example User'; RecipientTypeDetails = 'UserMailbox'
    }) | Export-Csv -LiteralPath (Join-Path $last $policyNames[0]) -NoTypeInformation
    @([pscustomobject]@{
        TenantKey = 'synthetic-test'; MemberId = 'user@example.invalid';
        MemberMail = 'user@example.invalid'; MemberUserPrincipalName = 'user@example.invalid';
        AccountEnabled = 'true'
    }) | Export-Csv -LiteralPath (Join-Path $last $policyNames[1]) -NoTypeInformation

    $module = Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
    & $module { param($root) $script:ProductRoot = $root } $fixture

    $issues = [Collections.Generic.List[object]]::new()
    $plan = Get-Plan -Issues $issues
    Check ($plan.Count -eq 2 -and $issues.Count -eq 3) 'Missing interactive sources blocked selection or lost diagnostics.'

    $output = Join-Path $fixture 'output'
    Run-Generator -Output $output
    $rows = @(Import-Csv -LiteralPath (Join-Path $output 'BackupMailboxEvidence.csv'))
    Check ($rows.Count -eq 1 -and $rows[0].'Protection State' -eq 'Unknown' -and
        $rows[0].'Risk Priority' -eq 'Unknown' -and $rows[0].'Evidence Status' -eq 'Not measured') 'Missing interactive protection was misclassified as an unprotected mailbox.'

    @([pscustomobject]@{
        TenantKey = 'synthetic-test'; ProtectionUnitId = 'unit-1'; MailboxId = 'user@example.invalid';
        Mail = 'user@example.invalid'; UserPrincipalName = 'user@example.invalid';
        DisplayName = 'Example User'; MailboxType = 'UserMailbox'; ProtectionPolicyName = 'Backup policy';
        ErrorCode = ''; ErrorMessage = ''; ExportDateTime = $now.ToString('O'); Status = 'Protected'
    }) | Export-Csv -LiteralPath (Join-Path $last $manualNames[0]) -NoTypeInformation
    foreach ($name in $manualNames[1..2]) {
        'TenantKey,Workload,ProtectionUnitId,DirectoryObjectId,Status' | Set-Content -LiteralPath (Join-Path $last $name)
    }
    foreach ($name in $manualNames) {
        [IO.File]::SetLastWriteTimeUtc((Join-Path $last $name), $now.AddHours(-169))
    }
    $issues = [Collections.Generic.List[object]]::new()
    $plan = Get-Plan -Issues $issues
    Check ($plan.Count -eq 5 -and $issues.Count -eq 3 -and @($issues | Where-Object Phase -ne 'StaleUsed').Count -eq 0) 'Stale interactive sources were not kept as the latest available data, or blocked preparation.'

    foreach ($name in $manualNames) {
        [IO.File]::SetLastWriteTimeUtc((Join-Path $last $name), $now.AddHours(-1))
    }
    $issues = [Collections.Generic.List[object]]::new()
    $plan = Get-Plan -Issues $issues
    Check ($plan.Count -eq 5 -and $issues.Count -eq 0) 'Fresh interactive sources were not selected.'
    Run-Generator -Output $output
    $rows = @(Import-Csv -LiteralPath (Join-Path $output 'BackupMailboxEvidence.csv'))
    Check ($rows.Count -eq 1 -and $rows[0].'Protection State' -eq 'Protected') 'Fresh protected-mailbox evidence was not preserved.'

    $unmatched = @(Import-Csv -LiteralPath (Join-Path $last $manualNames[0]))
    $unmatched[0].MailboxId = 'other@example.invalid'
    $unmatched[0].Mail = 'other@example.invalid'
    $unmatched[0].UserPrincipalName = 'other@example.invalid'
    $unmatched | Export-Csv -LiteralPath (Join-Path $last $manualNames[0]) -NoTypeInformation
    Run-Generator -Output $output
    $rows = @(Import-Csv -LiteralPath (Join-Path $output 'BackupMailboxEvidence.csv'))
    Check (@($rows | Where-Object { $_.'Reconciliation State' -eq 'Expected not protected' -and $_.'Risk Priority' -eq 'High' }).Count -eq 1) 'Available protection evidence stopped identifying actual gaps.'

    [IO.File]::SetLastWriteTimeUtc((Join-Path $last $policyNames[0]), $now.AddHours(-169))
    $failure = ''
    try { Get-Plan -Issues ([Collections.Generic.List[object]]::new()) | Out-Null } catch { $failure = $_.Exception.Message }
    Check ($failure -match 'publication-age limit.*M365_BackupPolicyScope_MailboxCoverage') 'Required noninteractive policy scope lost its freshness gate.'

    "PASS: $checks optional manual backup source and unknown-protection checks."
}
finally {
    $resolvedRoot = [IO.Path]::GetFullPath($TestRoot).TrimEnd('\', '/')
    $resolvedFixture = [IO.Path]::GetFullPath($fixture)
    if ($resolvedFixture.StartsWith($resolvedRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path $resolvedFixture -Leaf) -match '^[0-9a-f]{32}$') {
        Remove-Item -LiteralPath $resolvedFixture -Recurse -Force -ErrorAction SilentlyContinue
    }
}
