# Version 1.0.0: directory cleanup decisions, boundaries, matching, holds, scope and source qualification.
[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-OfflineTests'))
# Synthetic local files only. No collectors, tenant calls, production data or uploads.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
$script:checks = 0
function Check([bool]$Pass, [string]$Message) { if (-not $Pass) { throw $Message }; $script:checks++ }

$product = Split-Path -Parent $PSScriptRoot
$tableNames = @('User Cleanup Evidence', 'Computer Cleanup Evidence', 'Mailbox Cleanup Evidence', 'Directory Cleanup Source Qualification')
$now = [datetimeoffset]::new(2026, 10, 28, 12, 0, 0, [TimeSpan]::Zero)
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$raw = Join-Path $fixture 'raw'
$last = Join-Path $raw 'DATA-LAST'
$prepared = Join-Path $fixture 'prepared'
$output = Join-Path $fixture 'output'
$shell = (Get-Process -Id $PID).Path
$rules = Join-Path $fixture 'AccountClassification.json.txt'
$tenant = 'synthetic-test'

function Iso([datetimeoffset]$Value) { $Value.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ', [Globalization.CultureInfo]::InvariantCulture) }
function Day([datetimeoffset]$Value) { $Value.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) }
function Guid([int]$Number) { '00000000-0000-4000-8000-{0:D12}' -f $Number }
function Write-Csv([string]$Path, [object[]]$Rows) { $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8NoBOM }
$producers = @{
    'AD_Users_AllDomains.csv' = 'SmartM365-ActiveDirectory-Inventory.ps1'; 'AD_Computers_AllDomains.csv' = 'SmartM365-ActiveDirectory-Inventory.ps1'
    'M365_Users_Active.csv' = 'SmartM365-ActiveUsers-Inventory.ps1'; 'M365_Entra_Devices.csv' = 'SmartM365-EntraDevices-Inventory.ps1'
    'Intune_ManagedDevices_All.csv' = 'SmartM365-Devices-Inventory.ps1'; 'Exchange_EXO_Mailboxes_AllDomains.csv' = 'SmartM365-EXO-Mailboxes-Inventory.ps1'
    'Exchange_OnPrem_Mailboxes_AllDomains.csv' = 'SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
    'M365_Users_Activity.csv' = 'SmartM365-M365UserActivity-Inventory.ps1'; 'M365_Mailbox_Usage.csv' = 'SmartM365-M365UserActivity-Inventory.ps1'
}
function Write-Receipts([hashtable]$CompletedOverride = @{}, [datetimeoffset]$At = $now) {
    # One completed, non-partial SmartInventory receipt per producer, hashing each captured file.
    foreach ($producer in @($producers.Values | Sort-Object -Unique)) {
        $completed = if ($CompletedOverride.ContainsKey($producer)) { $CompletedOverride[$producer] } else { Iso $At.AddHours(-6) }
        $files = @(foreach ($file in @($producers.Keys | Where-Object { $producers[$_] -eq $producer } | Sort-Object)) {
            $path = Join-Path $last $file
            if (Test-Path -LiteralPath $path) { @{File = $file; Status = 'Success'; IsPartialInventory = $false; SHA256 = (Get-FileHash -LiteralPath $path).Hash; CompletedAtUtc = $completed} }
        })
        $stem = 'SmartInventory_' + [IO.Path]::GetFileNameWithoutExtension($producer)
        @{Owner = 'SmartInventory-SourceReceipt'; Producer = $producer; TenantKey = $tenant; Status = 'Completed'; IsPartialInventory = $false; CompletedAtUtc = $completed; Files = $files} |
            ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $last "$stem.current.json.txt")
    }
}
function Invoke-Cleanup([datetimeoffset]$AsOf = $now) {
    Remove-Item -LiteralPath $output -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $output | Out-Null
    $log = & $shell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'New-DirectoryCleanupEvidence.ps1') -DataRoot $raw -TenantKey $tenant `
        -UserInventoryPath (Join-Path $prepared 'UserInventoryEvidence.csv') -AllDevicesPath (Join-Path $prepared 'AllDevicesEvidence.csv') `
        -UserOutputPath (Join-Path $output 'UserCleanupEvidence.csv') -ComputerOutputPath (Join-Path $output 'ComputerCleanupEvidence.csv') `
        -MailboxOutputPath (Join-Path $output 'MailboxCleanupEvidence.csv') -QualificationOutputPath (Join-Path $output 'DirectoryCleanupSourceQualification.csv') `
        -AccountClassificationConfigPath $rules -AsOfUtc $AsOf.ToString('o') 3>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "Directory cleanup generator failed with exit code $LASTEXITCODE. $log" }
    @{
        Log = $log
        Users = @(Import-Csv -LiteralPath (Join-Path $output 'UserCleanupEvidence.csv'))
        Computers = @(Import-Csv -LiteralPath (Join-Path $output 'ComputerCleanupEvidence.csv'))
        Mailboxes = @(Import-Csv -LiteralPath (Join-Path $output 'MailboxCleanupEvidence.csv'))
        Sources = @(Import-Csv -LiteralPath (Join-Path $output 'DirectoryCleanupSourceQualification.csv'))
    }
}
function Test-Outputs {
    # Publication shape: TenantKey prefix, then the governed contract validator.
    $root = Join-Path $fixture ('validation-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $root | Out-Null
    foreach ($table in $script:evidenceTables) {
        $source = Join-Path $output $table.file
        $target = Join-Path $root $table.file
        $rows = @(Import-Csv -LiteralPath $source)
        if ($rows.Count) { $rows | Select-Object @{n = 'TenantKey'; e = { $tenant } }, * | Export-Csv -LiteralPath $target -NoTypeInformation -Encoding utf8 }
        else { [IO.File]::WriteAllText($target, '"TenantKey",' + (Get-Content -LiteralPath $source -TotalCount 1) + [Environment]::NewLine) }
    }
    $command = "& '{0}' -Root '{1}' -ContractPath '{2}' -Tables @({3}) -TenantKey '{4}'; exit `$LASTEXITCODE" -f (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1'), $root, (Join-Path $product 'config/prepared-evidence-contract.json.txt'), (($tableNames | ForEach-Object { "'$_'" }) -join ','), $tenant
    $report = & $shell -NoProfile -Command $command | Out-String
    Check ($LASTEXITCODE -eq 0) "Generated outputs violate the evidence contract: $report"
}

# --- Synthetic estate ----------------------------------------------------------------------------
$old = $now.AddYears(-1).AddDays(-200)          # 1-2 years
$older = $now.AddYears(-3)                      # 2+ years
$recent = $now.AddDays(-20)
$created = Iso $now.AddYears(-6)
function New-AdUser([int]$N, [hashtable]$Values = @{}) {
    $row = [ordered]@{TenantKey = $tenant; DomainName = 'contoso.example'; ObjectGUID = (Guid $N); ObjectSID = "S-1-5-21-1-1-1-$N"; ImmutableId_AD = "imm$N=="; SamAccountName = "user$N"
        UserPrincipalName = "user$N@contoso.example"; DisplayName = "Synthetic User $N"; Enabled = 'True'; AccountType = 'Named Account'; HasMailbox = 'False'; Country = 'XA'; CountryName = 'Synthetic Country A'
        LastLogonDateIso8601 = (Iso $old); WhenCreatedIso8601 = $created; PasswordLastSetDateIso8601 = (Iso $old); IsPasswordNeverExpires = 'False'; OnPremMailboxSizeMB = ''}
    foreach ($k in $Values.Keys) { $row[$k] = $Values[$k] }
    [pscustomobject]$row
}
function New-EntraUser([int]$N, [hashtable]$Values = @{}) {
    $row = [ordered]@{'Object Id' = (Guid (1000 + $N)); 'User principal name' = "user$N@contoso.example"; OnPremisesSecurityIdentifier = "S-1-5-21-1-1-1-$N"; OnPremisesImmutableId = "imm$N=="; AccountEnabled = 'True'
        LastSignInDateTimeIso8601 = (Iso $old); LastNonInteractiveSignInDateTimeIso8601 = ''; LastSuccessfulSignInDateTimeIso8601 = (Iso $old)}
    foreach ($k in $Values.Keys) { $row[$k] = $Values[$k] }
    [pscustomobject]$row
}
function New-Exo([int]$Mailbox, [int]$Owner, [hashtable]$Values = @{}) {
    $row = [ordered]@{MailboxGuid = (Guid (5000 + $Mailbox)); RecipientTypeDetails = 'UserMailbox'; UserPrincipalName = "user$Owner@contoso.example"; PrimarySmtpAddress = "mbx$Mailbox@contoso.example"
        ExternalDirectoryObjectId = (Guid (1000 + $Owner)); OnPremisesImmutableId = "imm$Owner=="; LitigationHoldEnabled = 'False'; RetentionHoldEnabled = 'False'; LastUserActionTimeIso8601 = ''; LastLogonTimeIso8601 = (Iso $old)
        WhenMailboxCreatedIso8601 = $created}
    foreach ($k in $Values.Keys) { $row[$k] = $Values[$k] }
    [pscustomobject]$row
}
function New-OnPrem([int]$Mailbox, [int]$Owner, [hashtable]$Values = @{}) {
    $row = [ordered]@{ExchangeGuid = (Guid (5000 + $Mailbox)); ObjectGUID = (Guid $Owner); DomainName = 'contoso.example'; RecipientType = 'UserMailbox'; PrimarySMTPaddress = "mbx$Mailbox@contoso.example"
        LitigationHoldEnabled = 'False'; RetentionHoldEnabled = 'False'; InPlaceHolds = ''; LastLogonTimeIso8601 = (Iso $older); WhenMailboxCreatedIso8601 = $created}
    foreach ($k in $Values.Keys) { $row[$k] = $Values[$k] }
    [pscustomobject]$row
}
function New-Computer([int]$N, [hashtable]$Values = @{}) {
    $row = [ordered]@{ObjectGUID = (Guid (3000 + $N)); DomainName = 'contoso.example'; Name = "PC$N"; SamAccountName = "PC$N$"; Enabled = 'True'; OperatingSystem = 'Windows 11 Enterprise'; OperatingSystemShortName = 'Windows 11'
        LastLogonDateIso8601 = (Iso $old); WhenCreatedIso8601 = $created; pwdLastSetDateIso8601 = (Iso $old); EntraMatchCountByObjectGuid = '0'; ExistsInEntraByObjectGuid = 'False'; EntraApproximateLastSignInDateTime = ''; IsInDisabledObjectsOU = 'False'}
    foreach ($k in $Values.Keys) { $row[$k] = $Values[$k] }
    [pscustomobject]$row
}
function Write-Estate {
    $adUsers = @(
        New-AdUser 1 @{LastLogonDateIso8601 = (Iso $recent)}                                                    # Active
        New-AdUser 2 @{HasMailbox = 'True'}                                                                     # Candidate 1-2 years, enabled
        New-AdUser 3 @{Enabled = 'False'; HasMailbox = 'True'; LastLogonDateIso8601 = (Iso $older)}            # Candidate 2+ years, disabled
        New-AdUser 4 @{LastLogonDateIso8601 = (Iso $older)}                                                     # Recent Entra sign-in
        New-AdUser 5 @{HasMailbox = 'True'}                                                                     # Mailbox hold
        New-AdUser 6 @{AccountType = 'Service Account'}                                                         # Non-human
        New-AdUser 7 @{ObjectSID = 'S-1-5-21-9-9-9-7'; ImmutableId_AD = 'other7=='}                             # UPN-only Entra match
        New-AdUser 8 @{LastLogonDateIso8601 = (Iso $now.AddYears(-1).AddDays(-5)); UserPrincipalName = 'solo8@contoso.example'; ObjectSID = 'S-1-5-21-8-8-8-8'; ImmutableId_AD = 'solo8=='}
        New-AdUser 9 @{LastLogonDateIso8601 = ''; UserPrincipalName = 'solo9@contoso.example'; ObjectSID = 'S-1-5-21-8-8-8-9'; ImmutableId_AD = 'solo9=='}
        New-AdUser 10 @{LastLogonDateIso8601 = ''; WhenCreatedIso8601 = (Iso $now.AddMonths(-5)); UserPrincipalName = 'solo10@contoso.example'; ObjectSID = 'S-1-5-21-8-8-8-10'; ImmutableId_AD = 'solo10=='}
        New-AdUser 11 @{LastLogonDateIso8601 = (Iso $now.AddDays(40)); UserPrincipalName = 'solo11@contoso.example'; ObjectSID = 'S-1-5-21-8-8-8-11'; ImmutableId_AD = 'solo11=='}
        New-AdUser 12 @{LastLogonDateIso8601 = (Iso $now.AddHours(2)); UserPrincipalName = 'solo12@contoso.example'; ObjectSID = 'S-1-5-21-8-8-8-12'; ImmutableId_AD = 'solo12=='}
        New-AdUser 13 @{HasMailbox = 'True'; UserPrincipalName = 'solo13@contoso.example'; ObjectSID = 'S-1-5-21-8-8-8-13'; ImmutableId_AD = 'solo13=='}
        New-AdUser 14 @{HasMailbox = 'True'}
        New-AdUser 15 @{HasMailbox = 'True'}
        New-AdUser 16 @{AccountType = 'Generic Account'; LastLogonDateIso8601 = (Iso $older); UserPrincipalName = 'solo16@contoso.example'; ObjectSID = 'S-1-5-21-8-8-8-16'; ImmutableId_AD = 'solo16=='}
        New-AdUser 17 @{LastLogonDateIso8601 = '2026-01-01 10:00:00'; UserPrincipalName = 'solo17@contoso.example'; ObjectSID = 'S-1-5-21-8-8-8-17'; ImmutableId_AD = 'solo17=='}
    )
    Write-Csv (Join-Path $last 'AD_Users_AllDomains.csv') $adUsers
    $entra = @(
        foreach ($n in 1, 2, 5, 6, 14, 15) { New-EntraUser $n }
        New-EntraUser 3 @{LastSignInDateTimeIso8601 = (Iso $older); LastSuccessfulSignInDateTimeIso8601 = (Iso $older)}
        New-EntraUser 4 @{LastSignInDateTimeIso8601 = (Iso $recent)}
        New-EntraUser 7 @{OnPremisesSecurityIdentifier = ''; OnPremisesImmutableId = ''}
    )
    Write-Csv (Join-Path $last 'M365_Users_Active.csv') $entra
    Write-Csv (Join-Path $last 'M365_Users_Activity.csv') @(
        [pscustomobject]@{UserPrincipalName = 'user2@contoso.example'; IsDeleted = 'False'; LastActivityDate = (Day $old); ReportRefreshDate = (Day $now.AddDays(-3))}
        [pscustomobject]@{UserPrincipalName = 'user14@contoso.example'; IsDeleted = 'False'; LastActivityDate = ''; ReportRefreshDate = (Day $now.AddDays(-3))}
    )
    Write-Csv (Join-Path $last 'Exchange_EXO_Mailboxes_AllDomains.csv') @(
        New-Exo 1 2                                                       # m1: user 2, old, enabled owner
        New-Exo 2 2 @{RecipientTypeDetails = 'SharedMailbox'}             # m2: shared, out of scope
        New-Exo 3 5 @{LitigationHoldEnabled = 'True'}                     # m3: hold
        New-Exo 5 99                                                      # m5: owner unknown
        New-Exo 6 14                                                      # m6 EXO side: owner 14
    )
    Write-Csv (Join-Path $last 'Exchange_OnPrem_Mailboxes_AllDomains.csv') @(
        New-OnPrem 4 3                                                    # m4: user 3, disabled owner
        New-OnPrem 6 15                                                   # m6 on-prem side: owner 15 -> conflict
        New-OnPrem 7 98 @{ExchangeGuid = (Guid 5007)}                      # m7 duplicated GUID
        New-OnPrem 8 98 @{ExchangeGuid = (Guid 5007)}
    )
    Write-Csv (Join-Path $last 'M365_Mailbox_Usage.csv') @(
        [pscustomobject]@{'User Principal Name' = 'user2@contoso.example'; 'Is Deleted' = 'False'; 'Last Activity Date' = (Day $old); 'Storage Used (Byte)' = '1073741824'; 'Report Refresh Date' = (Day $now.AddDays(-3))}
    )
    Write-Csv (Join-Path $last 'AD_Computers_AllDomains.csv') @(
        New-Computer 1 @{LastLogonDateIso8601 = (Iso $recent)}                                                                                     # Active
        New-Computer 2 @{EntraMatchCountByObjectGuid = '1'; ExistsInEntraByObjectGuid = 'True'; EntraApproximateLastSignInDateTime = (Iso $old)}     # Candidate, not enrolled
        New-Computer 3 @{OperatingSystem = 'Windows Server 2022 Standard'; OperatingSystemShortName = 'Windows Server 2022'}                           # Server: out of scope
        New-Computer 4 @{OperatingSystem = ''; OperatingSystemShortName = ''}                                                                       # Undetermined
        New-Computer 5 @{EntraMatchCountByObjectGuid = '1'; ExistsInEntraByObjectGuid = 'True'}                                                      # Two Entra objects
        New-Computer 6 @{EntraMatchCountByObjectGuid = '1'; ExistsInEntraByObjectGuid = 'True'; Enabled = 'False'}                                   # Intune recent
    )
    Write-Csv (Join-Path $last 'M365_Entra_Devices.csv') @(
        [pscustomobject]@{DeviceId = (Guid 3002)}; [pscustomobject]@{DeviceId = (Guid 3005)}; [pscustomobject]@{DeviceId = (Guid 3005)}; [pscustomobject]@{DeviceId = (Guid 3006)}
    )
    Write-Csv (Join-Path $last 'Intune_ManagedDevices_All.csv') @([pscustomobject]@{ManagedDeviceId = 'intune-6'; AzureADDeviceId = (Guid 3006); LastSyncDateTime = (Iso $recent)})
    New-Item -ItemType Directory -Path $prepared -Force | Out-Null
    Write-Csv (Join-Path $prepared 'UserInventoryEvidence.csv') @(
        [pscustomobject]@{'User Source ID' = (Guid 1002); Country = 'Synthetic Country B'; 'User Type' = 'Member'; 'Workforce Persona' = 'Synthetic persona'; 'Console Last Activity UTC' = ''}
    )
    Write-Csv (Join-Path $prepared 'AllDevicesEvidence.csv') @([pscustomobject]@{'AD Object GUID' = (Guid 3002); Country = 'Synthetic Country C'; 'Country Basis' = 'Primary-user proxy: Entra country'})
}

try {
    New-Item -ItemType Directory -Path $last -Force | Out-Null
    '{"SchemaVersion":"1.0","RuleVersion":"SYNTHETIC-1","Population":{"Human":["Named Account","Ext Account"],"NonHuman":["Service Account","Shared Mailbox"]}}' | Set-Content -LiteralPath $rules

    # Contracts and build wiring.
    $evidenceContract = Get-Content -LiteralPath (Join-Path $product 'config/prepared-evidence-contract.json.txt') -Raw | ConvertFrom-Json
    $script:evidenceTables = @($evidenceContract.tables | Where-Object { $_.table -in $tableNames })
    Check ($script:evidenceTables.Count -eq 4) 'Every directory cleanup table must be declared once in the evidence contract'
    Check (@($script:evidenceTables | Where-Object { $_.table -match 'History|Trend' }).Count -eq 0) 'Cleanup tables are current snapshots, never history-coverage tables'
    $build = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Invoke-PreparedEvidenceBuild.ps1') -Raw
    Check ($build -match "Add-Task 'DirectoryCleanup' 'DirectoryCleanup' 'New-DirectoryCleanupEvidence\.ps1' \`$cleanupArgs -DependsOn 'Workforce','Estate'") 'Build does not run directory cleanup after Workforce and Estate'

    # 1. Qualified estate.
    Write-Estate; Write-Receipts
    $run = Invoke-Cleanup
    Test-Outputs
    $u = @{}; foreach ($r in $run.Users) { $u[$r.SamAccountName] = $r }
    $c = @{}; foreach ($r in $run.Computers) { $c[$r.'Computer Name'] = $r }
    $m = @{}; foreach ($r in $run.Mailboxes) { $m[$r.'Primary SMTP Address' + '|' + $r.'Hosting Location'] = $r }
    Check ($run.Users.Count -eq 17) 'Every AD user must give one row'
    Check ($u.user1.'Cleanup State' -eq 'Active' -and $u.user1.'Inactivity Bucket' -eq '< 1 year' -and $u.user1.'Recommended Action' -eq 'Keep') 'Recent AD logon must be Active'
    Check ($u.user2.'Cleanup State' -eq 'Candidate' -and $u.user2.'Inactivity Bucket' -eq '1-2 years' -and $u.user2.'Recommended Action' -eq 'Disable account') 'Old enabled account must be a 1-2 year candidate to disable'
    Check ($u.user2.'Identity Match Method' -eq 'SID' -and $u.user2.'Workforce Persona' -eq 'Synthetic persona' -and $u.user2.'Country Basis' -eq 'User Inventory Evidence') 'Strong match must reuse the User Inventory classification'
    Check ($u.user2.'Secondary Reasons' -match 'Organization retention policies not assessed' -and $u.user2.'Hold State' -eq 'No mailbox-level hold') 'EXO mailbox-level hold check must state the unassessed retention policies'
    Check ($u.user2.'Mailbox Usage Last Activity Quality' -eq 'Qualified' -and $u.user2.'M365 Last Activity Quality' -eq 'Qualified') 'Graph reports must be read through the strongly matched Entra UPN'
    Check ($u.user3.'Cleanup State' -eq 'Candidate' -and $u.user3.'Inactivity Bucket' -eq '2+ years' -and $u.user3.'Recommended Action' -eq 'Delete account') 'Old disabled account must be a 2+ year candidate to delete'
    Check ($u.user4.'Cleanup State' -eq 'Active' -and $u.user4.'Last Activity Source' -eq 'Entra Last Sign-In') 'Recent Entra sign-in must keep an old AD account active'
    Check ($u.user5.'Cleanup State' -eq 'Excluded - hold' -and $u.user5.'Litigation Hold' -eq 'True') 'Linked mailbox hold must exclude the user'
    Check ($u.user6.'Cleanup State' -eq 'Excluded - non-human' -and $u.user6.'Account Population' -eq 'Non-human') 'Non-human population must be excluded'
    Check ($u.user7.'Cleanup State' -eq 'Review' -and $u.user7.'Identity Match Status' -eq 'Weak' -and $u.user7.'Entra Last Sign-In Quality' -eq 'Weak match') 'UPN-only match must be Review and never count as activity'
    Check ($u.user8.'Cleanup State' -eq 'Review' -and $u.user8.'Decision Reason' -eq 'Near threshold (AD replication latency)' -and $u.user8.'Inactivity Bucket' -eq '1-2 years') 'AD logon inside the latency margin must be Review'
    Check ($u.user9.'Cleanup State' -eq 'Candidate' -and $u.user9.'Inactivity Bucket' -eq 'No observed activity' -and $u.user9.'Identity Match Status' -eq 'No Entra object' -and $u.user9.'Has Mailbox' -eq 'No') 'Qualified no-activity account older than a year must be a candidate'
    Check ($u.user10.'Cleanup State' -eq 'Review' -and $u.user10.'Decision Reason' -match 'created less than one year ago') 'Recent account without activity must be Review'
    Check ($u.user11.'Cleanup State' -eq 'Review' -and $u.user11.'AD Last Logon Quality' -eq 'Future date' -and $u.user11.'Decision Reason' -match 'Future date') 'Future date must be Review'
    Check ($u.user12.'Cleanup State' -eq 'Active' -and $u.user12.'AD Last Logon Quality' -eq 'Clock skew' -and $u.user12.'Days Since Last Activity' -eq '0') 'Date within one day ahead must be clamped to the reference'
    Check ($u.user13.'Cleanup State' -eq 'Review' -and $u.user13.'Has Mailbox' -eq 'Unknown' -and $u.user13.'Hold State' -eq 'Unknown') 'Mailbox presence without mailbox must be Review'
    Check ($u.user16.'Cleanup State' -eq 'Candidate' -and $u.user16.'Account Population' -eq 'Review required' -and $u.user16.'Secondary Reasons' -match 'classification review') 'Population review is a secondary reason, not a block'
    Check ($u.user17.'AD Last Logon Quality' -eq 'Unqualified timezone' -and $u.user17.'Cleanup State' -eq 'Review') 'A date without timezone is never read as UTC'
    Check ($u.user2.'User Cleanup Key' -like "$tenant|cleanup-ad-user|*" -and @($run.Users.'User Cleanup Key' | Sort-Object -Unique).Count -eq 17) 'User keys must be tenanted and unique'
    Check ([int]$u.user1.'State Sort' -eq 1 -and [int]$u.user2.'State Sort' -eq 2 -and [int]$u.user7.'State Sort' -eq 3 -and [int]$u.user9.'Bucket Sort' -eq 4) 'Stable sort orders changed'

    Check ($run.Computers.Count -eq 5 -and -not $c.ContainsKey('PC3')) 'Servers must be out of scope'
    Check ($c.PC1.'Cleanup State' -eq 'Active') 'Recent workstation must be Active'
    Check ($c.PC2.'Cleanup State' -eq 'Candidate' -and $c.PC2.'Intune Match Status' -eq 'Not enrolled' -and $c.PC2.'Recommended Action' -eq 'Disable computer' -and $c.PC2.Country -eq 'Synthetic Country C') 'Old hybrid workstation must be a candidate with its device country'
    Check ($c.PC2.'Account Population' -eq 'Not applicable' -and $c.PC2.'Workforce Persona' -eq 'Not applicable') 'User classification is not applicable to computers'
    Check ($c.PC4.'Cleanup State' -eq 'Review' -and $c.PC4.'Scope State' -eq 'Undetermined') 'Unknown OS must be Review, not excluded'
    Check ($c.PC5.'Cleanup State' -eq 'Review' -and $c.PC5.'Entra Match Status' -eq 'Ambiguous') 'Duplicate Entra device must be Review'
    Check ($c.PC6.'Cleanup State' -eq 'Active' -and $c.PC6.'Last Activity Source' -eq 'Intune Last Sync') 'Recent Intune sync must keep a workstation active'
    Check (@($run.Sources | Where-Object File -eq 'AD_Computers_AllDomains.csv').'Out Of Scope Rows' -eq '1') 'Excluded servers must be counted'

    Check (-not $m.ContainsKey('mbx2@contoso.example|Exchange Online')) 'Shared mailboxes must be out of scope'
    Check ($m['mbx1@contoso.example|Exchange Online'].'Cleanup State' -eq 'Candidate' -and $m['mbx1@contoso.example|Exchange Online'].'Recommended Action' -eq 'Review owner' -and $m['mbx1@contoso.example|Exchange Online'].'Mailbox Size GB' -eq '1') 'Old mailbox of an enabled owner must ask for an owner review'
    Check ($m['mbx3@contoso.example|Exchange Online'].'Cleanup State' -eq 'Excluded - hold') 'Mailbox on hold must be excluded'
    Check ($m['mbx4@contoso.example|On-premises'].'Cleanup State' -eq 'Candidate' -and $m['mbx4@contoso.example|On-premises'].'Recommended Action' -eq 'Remove mailbox after retention check' -and $m['mbx4@contoso.example|On-premises'].'Owner Account State' -eq 'Disabled') 'Old mailbox of a disabled owner must be removable after a retention check'
    Check ($m['mbx5@contoso.example|Exchange Online'].'Cleanup State' -eq 'Review' -and $m['mbx5@contoso.example|Exchange Online'].'Owner Match Status' -eq 'Not found' -and $m['mbx5@contoso.example|Exchange Online'].'Recommended Action' -eq 'Review owner') 'Mailbox without owner must be Review owner'
    $conflict = @($run.Mailboxes | Where-Object 'Mailbox GUID' -eq (Guid 5006))
    Check ($conflict.Count -eq 2 -and @($conflict | Where-Object { $_.'Cleanup State' -ne 'Review' -or $_.'Decision Reason' -ne 'Conflicting hybrid identity' }).Count -eq 0 -and @($conflict.'Mailbox Cleanup Key' | Sort-Object -Unique).Count -eq 2) 'Conflicting hybrid sides must stay two Review rows'
    Check (@($run.Mailboxes | Where-Object 'Mailbox GUID' -eq (Guid 5007)).Count -eq 0 -and $run.Log -match 'mailbox rows without a unique mailbox GUID') 'Duplicated mailbox GUID must not be published silently'
    Check (@($run.Sources | Where-Object Source -like 'User mailboxes*').'Unkeyable Rows' -eq '2') 'Unkeyable mailbox rows must be counted'

    # 2. Leap-day boundaries: the one-year limit of 29 February is 28 February.
    $leap = [datetimeoffset]::new(2028, 2, 29, 12, 0, 0, [TimeSpan]::Zero)
    $boundary = @(New-AdUser 20 @{LastLogonDateIso8601 = '2027-02-28T12:00:00Z'; UserPrincipalName = 'x20@contoso.example'; ObjectSID = 'S-1-5-21-7-7-7-20'; ImmutableId_AD = 'x20=='}
        New-AdUser 21 @{LastLogonDateIso8601 = '2027-02-28T12:00:01Z'; UserPrincipalName = 'x21@contoso.example'; ObjectSID = 'S-1-5-21-7-7-7-21'; ImmutableId_AD = 'x21=='}
        New-AdUser 22 @{LastLogonDateIso8601 = '2026-02-28T12:00:00Z'; UserPrincipalName = 'x22@contoso.example'; ObjectSID = 'S-1-5-21-7-7-7-22'; ImmutableId_AD = 'x22=='})
    Write-Csv (Join-Path $last 'AD_Users_AllDomains.csv') $boundary
    Write-Receipts @{} $leap; $leapRun = Invoke-Cleanup $leap
    $b = @{}; foreach ($r in $leapRun.Users) { $b[$r.SamAccountName] = $r }
    Check ($b.user20.'Inactivity Bucket' -eq '1-2 years' -and $b.user21.'Inactivity Bucket' -eq '< 1 year' -and $b.user22.'Inactivity Bucket' -eq '2+ years') 'Calendar year boundaries changed around 29 February'

    # 3. Source qualification: stale Entra, missing mailbox usage and a receipt that no longer matches.
    Write-Estate
    Remove-Item -LiteralPath (Join-Path $last 'M365_Mailbox_Usage.csv')
    Write-Receipts @{'SmartM365-ActiveUsers-Inventory.ps1' = (Iso $now.AddHours(-400))}
    Add-Content -LiteralPath (Join-Path $last 'M365_Entra_Devices.csv') -Value ('"' + (Guid 3999) + '"')
    $degraded = Invoke-Cleanup
    $d = @{}; foreach ($r in $degraded.Users) { $d[$r.SamAccountName] = $r }
    $s = @{}; foreach ($r in $degraded.Sources) { $s[$r.File] = $r }
    Check ($s['M365_Users_Active.csv'].Status -eq 'Stale' -and $s['M365_Mailbox_Usage.csv'].Status -eq 'Not collected' -and $s['M365_Entra_Devices.csv'].Status -eq 'Freshness unknown') 'Source statuses must come from receipts and report dates'
    Check ($d.user2.'Cleanup State' -eq 'Review' -and $d.user2.'Identity Match Status' -eq 'Unknown') 'A stale Entra source must stop candidates'
    Check (@($degraded.Users | Where-Object 'Cleanup State' -eq 'Candidate').Count -eq 0) 'No candidate may survive a stale identity source'
    Check (@($degraded.Computers | Where-Object { $_.'Cleanup State' -eq 'Candidate' }).Count -eq 0) 'An unqualified Entra device capture must stop computer candidates'
    Check ($degraded.Log -match 'Directory cleanup source Entra users: Stale') 'Unqualified sources must be reported as warnings'
    Test-Outputs

    # 4. Nothing collected: headers only, still contract-valid.
    Get-ChildItem -LiteralPath $last -Filter '*.csv' | Where-Object Name -ne 'AD_Users_AllDomains.csv' | Remove-Item
    Write-Csv (Join-Path $last 'AD_Users_AllDomains.csv') @(New-AdUser 1)
    Write-Receipts
    $minimal = Invoke-Cleanup
    Check ($minimal.Computers.Count -eq 0 -and $minimal.Mailboxes.Count -eq 0 -and $minimal.Users.Count -eq 1 -and $minimal.Users[0].'Cleanup State' -eq 'Review') 'Missing sources must give empty tables and Review, never candidates'
    Test-Outputs

    Write-Host "Directory cleanup evidence offline test passed ($script:checks checks)."
} finally {
    Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD7+3ShCd4F8uAE
# L7iYdaQC0fVONqgccXtrNKaeNULw/qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOuQ+4OGmuLVbd91kODw5fVeNjKYaubZQjAAKXWs64ZzMA0GCSqG
# SIb3DQEBAQUABIIBgJLwwYUAvkqv5yuYkwR/KLL+Ls64mu0AGeG5GOWnDygyr6Lz
# dPZCVc0QrfmtK5rwF7LtZaptIEGC6s3jVxtCv0x3r2bjbJjd25fETqQN0+lUw75B
# kDjlFqnWk/WPGNvoRatven0f+u41obOX7uZ3fUTZYOUFSQk3Ha4n0bQyRnQwdmsP
# a6KGxkHftdlRsKcXsFUZ7F4FHvzMdaAOLR9UbMT4HoELAyZpOA6WSkqGX+Sc9MZk
# OEDYfJgIOCpT0NJsDIJIISG5WrAfdCcJVedSmhhdNceRNWWi2s/uy+vEBxsXgc9V
# 2cXsNbhINxNWwB7Tx/LnSxVkreFeB8apNfIpsrlZg7hMtKeDy20OmtQthvsMWm9E
# nZJ3a2yr3DTc7wC+H/9M4aowCY/xwM06Z0nAmlexrZ4KqHZHbk3VU5MCuFxZ+Vvp
# /N0afqDjfXjAMqAyE0SXAsjWw5lxbnRjHG4UTGSmH608UCAHev7ioiXswmLMYJF0
# 4iojsgWvVi81k2iF+6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxOTUx
# NDNaMC8GCSqGSIb3DQEJBDEiBCD4wZGJT/hD9IuOCp/QZzhpuwx+6cMto/oB6dwt
# A4NV0TANBgkqhkiG9w0BAQEFAASCAgBVqkIbFaACdx015CVVGJ3ZeE73RSUW5X9M
# p1phPmr3DD81E209lCrHUupG3eWNrPYhlpHBLHR2RLFG2p/nNMzzpcKi9FNzqXUA
# RHYmIEE1U7ziRjWZuy+wA2yIGbamgAgbHQ/JerEeK+J7woX9Bp4XcXBGQcVVvBba
# WT5dME9jr2w/1UZEWoDBBNE02PYHLMNKw1VQgMM7PzMMDIqZOVF3j8M87mYnRGQT
# lNLLy9k4UQdYnav+M2JOExJIOhY/UtFKj5mls3NMdshBNGYjO7USjqD0/0zyEKui
# r6KN/1xWLBGeAmyxelvHVdlBnqxXw23mP1xBnEEyKx2YZevl6EgbqUUq3ixANBA2
# /v2p18SyZH3LXu4u8qZE8El1jMYf4EY1nUTknlk3YhqHMAGY0ONPRWoV9lkGSJoL
# qWxl+pn78G9p6lHUcfHgWuX2iImoFIx1JQ3kQGm0juluDhP+iUTw6DsTFV+snqDT
# iCZGZ83wPC7Ix2IEZ9n77Ka0kJ9OALgSBm7V9DoiWTxa4CNRiCTZS9fuKruNzJiv
# /tWn2jLTLdy+yRihgHOWRhAWRvhZCbANwbzfecv8jgXVBqucZa/0exgjyHSJ25zf
# xvUYoqDpBkyGgeCj7O7pRX63qpujOyTU2rP11L2/eyixsYuXvymlYkwSD5+q1qHA
# Ps/FhHyhWg==
# SIG # End signature block
