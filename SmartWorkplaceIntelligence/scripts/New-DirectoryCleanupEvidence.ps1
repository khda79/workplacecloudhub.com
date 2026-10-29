# Version 1.1.0: Persona Basis (User Inventory, else the AD inventory persona) and Account Type Source; 1.0.0: directory cleanup candidates for AD user accounts, AD workstation accounts and user mailboxes.
# Read-only evidence: an object is a candidate only when none of its qualified activity signals
# (AD, Entra, Exchange, Intune, Microsoft 365 usage, console) is more recent than one year and
# every applicable source is qualified. Weak matches, missing or stale sources, unknown holds and
# undetermined scope lead to Review, never to Candidate. Servers and shared or resource mailboxes
# are out of scope. No account, computer or mailbox is ever changed.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantKey,
    [Parameter(Mandatory = $true)][string]$UserInventoryPath,
    [Parameter(Mandatory = $true)][string]$AllDevicesPath,
    [string]$UserOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\UserCleanupEvidence.csv'),
    [string]$ComputerOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ComputerCleanupEvidence.csv'),
    [string]$MailboxOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\MailboxCleanupEvidence.csv'),
    [string]$QualificationOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\DirectoryCleanupSourceQualification.csv'),
    [string]$AccountClassificationConfigPath,
    # Collection age above which a receipt-qualified source is Stale (same default as Prepare's publication-age limit).
    [ValidateRange(1, 8760)][int]$SourceMaxAgeHours = 336,
    # Report Refresh Date age above which a Microsoft Graph usage report is Stale.
    [ValidateRange(1, 366)][int]$ReportMaxAgeDays = 31,
    # LastLogonTimestamp replication latency; a decision at the one-year boundary inside this margin is Review.
    [ValidateRange(0, 60)][int]$AdLatencyDays = 14,
    # Optional reference clock for reproducible offline preparation; defaults to the current time.
    [Nullable[datetimeoffset]]$AsOfUtc
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'SmartM365\SmartInventory\Common\SmartM365.AccountClassification.psd1') -MinimumVersion '1.2.0' -ErrorAction Stop
$classification = if ($AccountClassificationConfigPath) { Read-SmartM365AccountClassification -Path $AccountClassificationConfigPath } else { Read-SmartM365AccountClassification }
$populationRules = $classification.Population
$classificationRuleVersion = [string]$classification.RuleVersion

$R = if ($null -ne $AsOfUtc) { $AsOfUtc.ToUniversalTime() } else { [datetimeoffset]::UtcNow }
$referenceText = $R.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ', [Globalization.CultureInfo]::InvariantCulture)
$evidenceDate = $R.UtcDateTime.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
$oneYear = $R.AddYears(-1); $twoYears = $R.AddYears(-2); $latencyFloor = $oneYear.AddDays(-$AdLatencyDays)
$last = Join-Path $DataRoot 'DATA-LAST'
$sha = [Security.Cryptography.SHA256]::Create()

# Per-object helpers run several million times on a large estate. PowerShell function calls cost
# about a hundred microseconds each, so the hot helpers are compiled once as static methods.
if (-not ('SmartM365.DirectoryCleanup.Fast' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Management.Automation;
using System.Text.RegularExpressions;
namespace SmartM365.DirectoryCleanup {
    public sealed class Signal { public string Name; public DateTimeOffset? Utc; public string Quality; }
    public sealed class Hit { public int Count; public object Row; }
    public static class Fast {
        static readonly Regex Offset = new Regex(@"(Z|[+-]\d{2}:\d{2})$", RegexOptions.Compiled);
        static readonly Regex DayOnly = new Regex(@"^\d{4}-\d{2}-\d{2}$", RegexOptions.Compiled);
        public static string T(object value) { return value == null ? "" : value.ToString().Trim(); }
        public static string K(object value) { return T(value).ToLowerInvariant(); }
        // The first present, non-empty column wins; an absent row or column reads as empty.
        public static string F(object row, params string[] names) {
            if (row == null) return "";
            PSObject ps = row as PSObject ?? PSObject.AsPSObject(row);
            foreach (string name in names) {
                PSPropertyInfo p = ps.Properties[name];
                if (p == null || p.Value == null) continue;
                string s = p.Value.ToString().Trim();
                if (s.Length > 0) return s;
            }
            return "";
        }
        public static string KF(object row, params string[] names) { return F(row, names).ToLowerInvariant(); }
        public static bool B(object value) { string s = K(value); return s == "true" || s == "1" || s == "yes"; }
        public static bool BF(object row, params string[] names) { return B(F(row, names)); }
        public static string Utc(object value) {
            if (value == null) return "";
            return ((DateTimeOffset)value).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        }
        public static void Add(Hashtable map, string key, object row) {
            if (string.IsNullOrEmpty(key)) return;
            List<object> list = map[key] as List<object>;
            if (list == null) { list = new List<object>(); map[key] = list; }
            list.Add(row);
        }
        public static Hit One(Hashtable map, string key) {
            Hit hit = new Hit();
            if (string.IsNullOrEmpty(key)) return hit;
            List<object> list = map[key] as List<object>;
            if (list == null) return hit;
            hit.Count = list.Count;
            if (list.Count == 1) hit.Row = list[0];
            return hit;
        }
        public static object[] List(Hashtable map, string key) {
            if (string.IsNullOrEmpty(key)) return new object[0];
            List<object> list = map[key] as List<object>;
            return list == null ? new object[0] : list.ToArray();
        }
        // Quality: why a signal counts or not. Only Qualified and Clock skew values count as activity.
        public static Signal NewSignal(string name, string raw, string sourceStatus, string match, DateTimeOffset reference) {
            Signal s = new Signal(); s.Name = name;
            if (match == "Not applicable") { s.Quality = "Not applicable"; return s; }
            if (sourceStatus != "Observed") { s.Quality = sourceStatus; return s; }
            if (match != "Qualified") { s.Quality = match; return s; }
            if (string.IsNullOrEmpty(raw)) { s.Quality = "No value"; return s; }
            if (DayOnly.IsMatch(raw)) {
                // Graph report dates are UTC calendar days, compared with the reference day.
                DateTimeOffset day = DateTimeOffset.ParseExact(raw, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);
                if (day.UtcDateTime.Date > reference.UtcDateTime.Date.AddDays(1)) { s.Quality = "Future date"; return s; }
                s.Utc = day > reference ? reference : day; s.Quality = "Qualified"; return s;
            }
            DateTimeOffset parsed;
            if (!Offset.IsMatch(raw) || !DateTimeOffset.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed)) { s.Quality = "Unqualified timezone"; return s; }
            parsed = parsed.ToUniversalTime();
            if (parsed > reference.AddDays(1)) { s.Quality = "Future date"; return s; }
            if (parsed > reference) { s.Utc = reference; s.Quality = "Clock skew"; return s; }
            s.Utc = parsed; s.Quality = "Qualified"; return s;
        }
        // Fixed order breaks ties: the first signal of the list with the latest date names the source.
        public static Signal Latest(IEnumerable signals) {
            Signal best = null;
            foreach (object o in signals) {
                Signal s = o as Signal ?? (o is PSObject ? ((PSObject)o).BaseObject as Signal : null);
                if (s == null || !s.Utc.HasValue || (s.Quality != "Qualified" && s.Quality != "Clock skew")) continue;
                if (best == null || s.Utc.Value > best.Utc.Value) best = s;
            }
            return best;
        }
    }
}
'@
}
$Fast = [SmartM365.DirectoryCleanup.Fast]
function T([AllowNull()][object]$Value) { $Fast::T($Value) }
function K([AllowNull()][object]$Value) { $Fast::K($Value) }
function F([AllowNull()][object]$Row, [string[]]$Names) { $Fast::F($Row, $Names) }
function B([AllowNull()][object]$Value) { $Fast::B($Value) }
function Get-Prop([AllowNull()][object]$Object, [string]$Name) { if ($null -eq $Object) { return $null }; $p = $Object.PSObject.Properties[$Name]; if ($null -eq $p) { $null } else { $p.Value } }
function New-Key([string]$Kind, [string]$Value) {
    $digest = [Convert]::ToHexString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value.ToLowerInvariant()))).ToLowerInvariant()
    "$TenantKey|$Kind|$digest"
}
function Format-Utc([AllowNull()][object]$Value) {
    if ($null -eq $Value) { return '' }
    ([datetimeoffset]$Value).UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
}
function Read-ReceiptUtc([AllowNull()][object]$Value) {
    # ConvertFrom-Json yields Utc for 'Z' and Local for an explicit offset; Unspecified has no timezone.
    if ($Value -is [datetimeoffset]) { return $Value.ToUniversalTime() }
    if ($Value -is [datetime]) { if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return $null }; return [datetimeoffset]::new($Value.ToUniversalTime()) }
    $text = T $Value
    if ($text -notmatch '(Z|[+-]\d{2}:\d{2})$') { return $null }
    $parsed = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.ToUniversalTime() }
    return $null
}

# --- Sources --------------------------------------------------------------------------------
$qualification = [Collections.Generic.List[object]]::new()
$notReady = @('Not collected', 'Schema mismatch', 'Stale', 'Freshness unknown', 'Partial coverage')
function Import-Narrow([string]$Path, [string[]]$Columns) {
    # Wide inventories (the AD user export has about 190 columns) are streamed and narrowed to the
    # columns this generator reads; an absent optional column reads as empty.
    # Property names are case-insensitive: keep one spelling of each column.
    $unique = [Collections.Generic.List[string]]::new(); $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($c in $Columns) { if ($seen.Add($c)) { $unique.Add($c) } }
    @(Import-Csv -LiteralPath $Path | Select-Object -Property $unique.ToArray())
}
function Get-Source([string]$Name, [string]$File, [string]$Producer, [string[]]$Columns, [string[]]$RefreshColumns = @(), [string[]]$Optional = @()) {
    $path = Join-Path $last $File
    $state = [ordered]@{Name = $Name; File = $File; Producer = $Producer; Rows = @(); Status = 'Observed'; CollectedUtc = $null; RefreshDate = $null; Reason = ''}
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $state.Status = 'Not collected'; $state.Reason = 'File not captured'; return [pscustomobject]$state }
    $first = @(Import-Csv -LiteralPath $path | Select-Object -First 1)
    if ($first.Count -eq 0) { $state.Status = 'Not collected'; $state.Reason = 'No rows'; return [pscustomobject]$state }
    $present = @($first[0].PSObject.Properties.Name)
    $missing = @($Columns | Where-Object { $present -notcontains $_ })
    if ($missing.Count) { $state.Status = 'Schema mismatch'; $state.Reason = 'Missing columns: ' + ($missing -join ', '); return [pscustomobject]$state }
    $rows = Import-Narrow $path (@($Columns) + @($RefreshColumns) + @($Optional) | Select-Object -Unique)
    $state.Rows = $rows
    # Business freshness comes from the producer receipt, never from the file time alone.
    $stem = 'SmartInventory_' + [IO.Path]::GetFileNameWithoutExtension($Producer)
    $receiptPath = Join-Path $last ($stem + '.current.json.txt')
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { $state.Status = 'Freshness unknown'; $state.Reason = 'Producer receipt not captured'; return [pscustomobject]$state }
    $proof = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
    $records = @(if ($proof.PSObject.Properties['Files']) { @($proof.Files) | Where-Object { $_.PSObject.Properties['File'] -and $_.File -ceq $File } })
    $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    if ((Get-Prop $proof 'Status') -cne 'Completed' -or $records.Count -ne 1 -or (Get-Prop $records[0] 'Status') -cne 'Success' -or [string](Get-Prop $records[0] 'SHA256') -ine $hash) {
        $state.Status = 'Freshness unknown'; $state.Reason = 'Receipt does not qualify this capture'; return [pscustomobject]$state
    }
    $completed = Read-ReceiptUtc $(if ($null -ne (Get-Prop $records[0] 'CompletedAtUtc')) { Get-Prop $records[0] 'CompletedAtUtc' } else { Get-Prop $proof 'CompletedAtUtc' })
    if ($null -eq $completed) { $state.Status = 'Freshness unknown'; $state.Reason = 'Receipt completion time has no timezone'; return [pscustomobject]$state }
    $state.CollectedUtc = $completed
    if ((Get-Prop $proof 'IsPartialInventory') -ne $false -or (Get-Prop $records[0] 'IsPartialInventory') -ne $false) { $state.Status = 'Partial coverage'; $state.Reason = 'Partial inventory'; return [pscustomobject]$state }
    if (($R - $completed).TotalHours -gt $SourceMaxAgeHours) { $state.Status = 'Stale'; $state.Reason = "Collected more than $SourceMaxAgeHours hours before the reference"; return [pscustomobject]$state }
    if ($RefreshColumns.Count) {
        $dates = @($rows | ForEach-Object { $Fast::F($_, $RefreshColumns) } | Where-Object { $_ -match '^\d{4}-\d{2}-\d{2}$' } | Sort-Object -Unique)
        if (-not $dates.Count) { $state.Status = 'Freshness unknown'; $state.Reason = 'No Report Refresh Date'; return [pscustomobject]$state }
        $state.RefreshDate = $dates[-1]
        $refresh = [datetimeoffset]::ParseExact($dates[-1], 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal)
        if (($R.UtcDateTime.Date - $refresh.UtcDateTime.Date).TotalDays -gt $ReportMaxAgeDays) { $state.Status = 'Stale'; $state.Reason = "Report refreshed more than $ReportMaxAgeDays days before the reference" }
    }
    return [pscustomobject]$state
}
$sources = [ordered]@{
    AdUsers     = Get-Source 'AD users' 'AD_Users_AllDomains.csv' 'SmartM365-ActiveDirectory-Inventory.ps1' @('ObjectGUID', 'ObjectSID', 'DomainName', 'Enabled', 'AccountType', 'LastLogonDateIso8601', 'WhenCreatedIso8601') @() @('ImmutableId_AD', 'SamAccountName', 'UserPrincipalName', 'DisplayName', 'HasMailbox', 'Country', 'CountryName', 'PasswordLastSetDateIso8601', 'IsPasswordNeverExpires', 'OnPremMailboxSizeMB', 'AccountTypeSource', 'WorkforcePersona')
    AdComputers = Get-Source 'AD computers' 'AD_Computers_AllDomains.csv' 'SmartM365-ActiveDirectory-Inventory.ps1' @('ObjectGUID', 'DomainName', 'Enabled', 'OperatingSystem', 'LastLogonDateIso8601', 'WhenCreatedIso8601', 'EntraMatchCountByObjectGuid', 'ExistsInEntraByObjectGuid', 'EntraApproximateLastSignInDateTime') @() @('Name', 'SamAccountName', 'OperatingSystemShortName', 'pwdLastSetDateIso8601', 'IsInDisabledObjectsOU')
    EntraUsers  = Get-Source 'Entra users' 'M365_Users_Active.csv' 'SmartM365-ActiveUsers-Inventory.ps1' @('Object Id', 'User principal name', 'OnPremisesSecurityIdentifier', 'OnPremisesImmutableId', 'AccountEnabled', 'LastSignInDateTimeIso8601', 'LastNonInteractiveSignInDateTimeIso8601', 'LastSuccessfulSignInDateTimeIso8601')
    EntraDevices = Get-Source 'Entra devices' 'M365_Entra_Devices.csv' 'SmartM365-EntraDevices-Inventory.ps1' @('DeviceId')
    Intune      = Get-Source 'Intune devices' 'Intune_ManagedDevices_All.csv' 'SmartM365-Devices-Inventory.ps1' @('AzureADDeviceId', 'LastSyncDateTime') @() @('ManagedDeviceId')
    Exo         = Get-Source 'Exchange Online mailboxes' 'Exchange_EXO_Mailboxes_AllDomains.csv' 'SmartM365-EXO-Mailboxes-Inventory.ps1' @('MailboxGuid', 'RecipientTypeDetails', 'UserPrincipalName', 'PrimarySmtpAddress', 'ExternalDirectoryObjectId', 'OnPremisesImmutableId', 'LitigationHoldEnabled', 'RetentionHoldEnabled', 'LastUserActionTimeIso8601', 'LastLogonTimeIso8601') @() @('ImmutableId', 'WhenMailboxCreatedIso8601')
    OnPrem      = Get-Source 'On-premises mailboxes' 'Exchange_OnPrem_Mailboxes_AllDomains.csv' 'SmartM365-Exchange-Local-Mailboxes-Inventory.ps1' @('ExchangeGuid', 'ObjectGUID', 'RecipientType', 'PrimarySMTPaddress', 'LitigationHoldEnabled', 'RetentionHoldEnabled', 'InPlaceHolds', 'LastLogonTimeIso8601') @() @('DomainName', 'PrimarySmtpAddress', 'WhenMailboxCreatedIso8601')
    Activity    = Get-Source 'Microsoft 365 user activity' 'M365_Users_Activity.csv' 'SmartM365-M365UserActivity-Inventory.ps1' @('UserPrincipalName', 'LastActivityDate') @('Report Refresh Date', 'ReportRefreshDate') @('IsDeleted', 'Is Deleted')
    MailboxUsage = Get-Source 'Mailbox usage' 'M365_Mailbox_Usage.csv' 'SmartM365-M365UserActivity-Inventory.ps1' @('User Principal Name', 'Last Activity Date') @('Report Refresh Date', 'ReportRefreshDate') @('Is Deleted', 'IsDeleted', 'Storage Used (Byte)')
}
function Test-Ready([string]$Name) { $sources[$Name].Status -eq 'Observed' }
$excluded = @{Servers = 0; SharedOrResourceMailboxes = 0; UnkeyableUsers = 0; UnkeyableComputers = 0; UnkeyableMailboxes = 0}

# Prepared outputs of the same batch (User Inventory from Workforce, All Devices from Estate).
$inventoryRows = @(if (Test-Path -LiteralPath $UserInventoryPath -PathType Leaf) { Import-Narrow $UserInventoryPath @('User Source ID', 'Country', 'User Type', 'Workforce Persona', 'Console Last Activity UTC') })
$deviceRows = @(if (Test-Path -LiteralPath $AllDevicesPath -PathType Leaf) { Import-Narrow $AllDevicesPath @('AD Object GUID', 'Country', 'Country Basis') })

# --- Indexes ----------------------------------------------------------------------------------
$entraBySid = @{}; $entraByImmutable = @{}; $entraByUpn = @{}; $entraById = @{}
foreach ($e in $sources.EntraUsers.Rows) {
    $Fast::Add($entraBySid, $Fast::KF($e, 'OnPremisesSecurityIdentifier'), $e)
    $Fast::Add($entraByImmutable, $Fast::F($e, 'OnPremisesImmutableId'), $e)      # base64: case-sensitive
    $Fast::Add($entraByUpn, $Fast::KF($e, 'User principal name'), $e)
    $Fast::Add($entraById, $Fast::KF($e, 'Object Id'), $e)
}
$adUserByGuid = @{}; $adUserBySid = @{}; $adUserByImmutable = @{}
foreach ($u in $sources.AdUsers.Rows) {
    $Fast::Add($adUserByGuid, $Fast::KF($u, 'ObjectGUID'), $u)
    $Fast::Add($adUserBySid, $Fast::KF($u, 'ObjectSID'), $u)
    $Fast::Add($adUserByImmutable, $Fast::F($u, 'ImmutableId_AD'), $u)
}
$inventoryById = @{}; foreach ($i in $inventoryRows) { $Fast::Add($inventoryById, $Fast::KF($i, 'User Source ID'), $i) }
$deviceByAdGuid = @{}; foreach ($d in $deviceRows) { $Fast::Add($deviceByAdGuid, $Fast::KF($d, 'AD Object GUID'), $d) }
$entraDeviceById = @{}; foreach ($d in $sources.EntraDevices.Rows) { $Fast::Add($entraDeviceById, $Fast::KF($d, 'DeviceId'), $d) }
$intuneByEntraId = @{}; foreach ($d in $sources.Intune.Rows) { $Fast::Add($intuneByEntraId, $Fast::KF($d, 'AzureADDeviceId'), $d) }
$activityByUpn = @{}; foreach ($a in $sources.Activity.Rows) { if (-not ($Fast::BF($a, 'IsDeleted', 'Is Deleted'))) { $Fast::Add($activityByUpn, $Fast::KF($a, 'UserPrincipalName'), $a) } }
$usageByUpn = @{}; foreach ($a in $sources.MailboxUsage.Rows) { if (-not ($Fast::BF($a, 'Is Deleted', 'IsDeleted'))) { $Fast::Add($usageByUpn, $Fast::KF($a, 'User Principal Name'), $a) } }
$userMailboxTypes = @('usermailbox')
$exoUser = [Collections.Generic.List[object]]::new(); $exoByOwnerId = @{}; $exoByImmutable = @{}
foreach ($m in $sources.Exo.Rows) {
    if (($Fast::KF($m, 'RecipientTypeDetails')) -notin $userMailboxTypes) { $excluded.SharedOrResourceMailboxes++; continue }
    $exoUser.Add($m)
    $Fast::Add($exoByOwnerId, $Fast::KF($m, 'ExternalDirectoryObjectId'), $m)
    $Fast::Add($exoByImmutable, $Fast::F($m, 'OnPremisesImmutableId', 'ImmutableId'), $m)
}
$onPremUser = [Collections.Generic.List[object]]::new(); $onPremByOwnerGuid = @{}
foreach ($m in $sources.OnPrem.Rows) {
    if (($Fast::KF($m, 'RecipientType')) -notin $userMailboxTypes) { $excluded.SharedOrResourceMailboxes++; continue }
    $onPremUser.Add($m)
    $Fast::Add($onPremByOwnerGuid, $Fast::KF($m, 'ObjectGUID'), $m)
}

# --- Signals and decisions -------------------------------------------------------------------
function Get-Bucket([AllowNull()]$Latest) {
    if ($null -eq $Latest) { return @('No observed activity', 4) }
    if ($Latest.Utc -gt $oneYear) { return @('< 1 year', 1) }
    if ($Latest.Utc -gt $twoYears) { return @('1-2 years', 2) }
    return @('2+ years', 3)
}
function Read-Created([string]$Raw) { $s = $Fast::NewSignal('Created', $Raw, 'Observed', 'Qualified', $R); if ($s.Quality -eq 'Qualified') { $s.Utc } else { $null } }
function Resolve-Decision {
    # Priority: hold, non-human, imperative Review, Active, decisive Review, Candidate.
    param($Latest, [object[]]$Signals, [string]$HoldState, [bool]$NonHuman, [string[]]$Imperative, [AllowNull()]$Created, [string[]]$Secondary)
    $bucket = Get-Bucket $Latest
    $reasons = [Collections.Generic.List[string]]::new(); foreach ($r in $Secondary) { if ($r) { $reasons.Add($r) } }
    $decisive = [Collections.Generic.List[string]]::new()
    foreach ($s in $Signals) { if ($s.Quality -in @('Future date', 'Unqualified timezone')) { $decisive.Add("$($s.Name): $($s.Quality)") } }
    $adLatest = $null -ne $Latest -and $Latest.Name -in @('AD Last Logon', 'Owner AD Last Logon')
    if ($adLatest -and $Latest.Utc -le $oneYear -and $Latest.Utc -ge $latencyFloor) { $decisive.Add('Near threshold (AD replication latency)') }
    if ($adLatest -and $Latest.Utc -le $twoYears -and $Latest.Utc -ge $twoYears.AddDays(-$AdLatencyDays)) { $reasons.Add('Near 2-year threshold (AD replication latency)') }
    if ($null -eq $Latest) {
        if ($null -eq $Created) { $decisive.Add('No observed activity and creation date unknown') }
        elseif ($Created -gt $oneYear) { $decisive.Add('No observed activity and created less than one year ago') }
    }
    $state = ''; $main = ''
    if ($HoldState -eq 'Hold') { $state = 'Excluded - hold'; $main = 'Mailbox hold' }
    elseif ($NonHuman) { $state = 'Excluded - non-human'; $main = 'Non-human account population' }
    elseif ($Imperative.Count) { $state = 'Review'; $main = $Imperative[0]; for ($i = 1; $i -lt $Imperative.Count; $i++) { $reasons.Add($Imperative[$i]) } }
    elseif ($bucket[1] -eq 1) { $state = 'Active'; $main = "Activity within one year: $($Latest.Name)" }
    elseif ($decisive.Count) { $state = 'Review'; $main = $decisive[0]; for ($i = 1; $i -lt $decisive.Count; $i++) { $reasons.Add($decisive[$i]) } }
    else { $state = 'Candidate'; $main = @{1 = ''; 2 = 'No activity for over one year'; 3 = 'No activity for over two years'; 4 = 'No observed activity on any qualified source'}[$bucket[1]] }
    # Every other finding stays visible; recent activity is shown when another rule took precedence.
    if ($state -in @('Excluded - hold', 'Excluded - non-human', 'Active')) { foreach ($r in $Imperative) { $reasons.Add($r) }; foreach ($r in $decisive) { $reasons.Add($r) } }
    if ($state -ne 'Active' -and $bucket[1] -eq 1) { $reasons.Add("Recent activity: $($Latest.Name)") }
    $stateSort = @{'Active' = 1; 'Candidate' = 2; 'Review' = 3; 'Excluded - hold' = 4; 'Excluded - non-human' = 5}[$state]
    [pscustomobject]@{Bucket = $bucket[0]; BucketSort = $bucket[1]; State = $state; StateSort = $stateSort; Reason = $main; Secondary = (Join-Unique $reasons)}
}
function Get-Population([string]$AccountType) {
    $type = if ($AccountType) { $AccountType } else { 'Unclassified Account' }
    Get-SmartM365AccountPopulation -PopulationRules $populationRules -AccountType $type
}
function Get-Persona([AllowNull()]$Inventory, [AllowNull()]$AdUser) {
    # User Inventory first (strong Entra match); otherwise the persona the AD inventory computed with the same workbook.
    if ($null -ne $Inventory -and ($Fast::F($Inventory, 'Workforce Persona'))) { return [pscustomobject]@{ Persona = $Fast::F($Inventory, 'Workforce Persona'); Basis = 'User Inventory' } }
    if ($null -ne $AdUser -and ($Fast::F($AdUser, 'WorkforcePersona'))) { return [pscustomobject]@{ Persona = $Fast::F($AdUser, 'WorkforcePersona'); Basis = 'AD attributes' } }
    [pscustomobject]@{ Persona = 'Unknown'; Basis = 'Unknown' }
}
function Get-Classification([AllowNull()]$Entra) {
    # User Inventory classification is reused only after a strong, unique Entra match.
    $result = [ordered]@{Status = 'Not matched'; Row = $null}
    if ($null -eq $Entra) { return [pscustomobject]$result }
    $hit = $Fast::One($inventoryById, $Fast::KF($Entra, 'Object Id'))
    if ($hit.Count -eq 1) { $result.Status = 'Matched'; $result.Row = $hit.Row } elseif ($hit.Count -gt 1) { $result.Status = 'Ambiguous' }
    [pscustomobject]$result
}
function Resolve-Entra($AdUser) {
    # SID first, then ImmutableId; both must point to one object. UPN alone is a weak hint.
    $result = [ordered]@{Status = ''; Method = ''; Row = $null}
    if (-not (Test-Ready 'EntraUsers')) { $result.Status = 'Unknown'; $result.Method = 'Entra source not qualified'; return [pscustomobject]$result }
    $bySid = $Fast::One($entraBySid, $Fast::KF($AdUser, 'ObjectSID'))
    $byImmutable = $Fast::One($entraByImmutable, $Fast::F($AdUser, 'ImmutableId_AD'))
    if ($bySid.Count -gt 1 -or $byImmutable.Count -gt 1 -or ($bySid.Count -eq 1 -and $byImmutable.Count -eq 1 -and -not [object]::ReferenceEquals($bySid.Row, $byImmutable.Row))) {
        $result.Status = 'Ambiguous'; $result.Method = 'Multiple Entra candidates'; return [pscustomobject]$result
    }
    if ($bySid.Count -eq 1) { $result.Status = 'Matched'; $result.Method = 'SID'; $result.Row = $bySid.Row; return [pscustomobject]$result }
    if ($byImmutable.Count -eq 1) { $result.Status = 'Matched'; $result.Method = 'ImmutableId'; $result.Row = $byImmutable.Row; return [pscustomobject]$result }
    $byUpn = $Fast::One($entraByUpn, $Fast::KF($AdUser, 'UserPrincipalName'))
    if ($byUpn.Count -ge 1) { $result.Status = 'Weak'; $result.Method = 'UPN only'; return [pscustomobject]$result }
    $result.Status = 'No Entra object'; $result.Method = 'No SID, ImmutableId or UPN match'; return [pscustomobject]$result
}
function Get-ReportMatch([hashtable]$Index, [string]$SourceName, [AllowNull()]$EntraRow, [string]$EntraStatus) {
    # Graph usage reports are keyed by UPN: read only through the UPN of a strongly matched Entra object.
    if (-not (Test-Ready $SourceName)) { return @{Match = 'Qualified'; Row = $null} }   # source status carries the reason
    if ($EntraStatus -eq 'No Entra object') { return @{Match = 'Not applicable'; Row = $null} }
    if ($EntraStatus -eq 'Weak') { return @{Match = 'Weak match'; Row = $null} }
    if ($EntraStatus -ne 'Matched') { return @{Match = 'Ambiguous match'; Row = $null} }
    $hit = $Fast::One($Index, $Fast::KF($EntraRow, 'User principal name'))
    if ($hit.Count -eq 1) { return @{Match = 'Qualified'; Row = $hit.Row} }
    if ($hit.Count -gt 1) { return @{Match = 'Ambiguous match'; Row = $null} }
    return @{Match = 'No report row'; Row = $null}
}
function Get-HoldState([object[]]$Mailboxes, [bool]$Known) {
    if (-not $Known) { return 'Unknown' }
    if (-not $Mailboxes.Count) { return 'Not applicable' }
    foreach ($m in $Mailboxes) { if (($Fast::BF($m, 'LitigationHoldEnabled')) -or ($Fast::BF($m, 'RetentionHoldEnabled')) -or ($Fast::F($m, 'InPlaceHolds'))) { return 'Hold' } }
    return 'No mailbox-level hold'
}
function Join-Flags([object[]]$Mailboxes, [string]$Column) {
    if (-not $Mailboxes.Count) { return '' }
    foreach ($m in $Mailboxes) { if ($Fast::BF($m, $Column)) { return 'True' } }
    return 'False'
}
function Join-Unique([Collections.Generic.List[string]]$Items) {
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    (@(foreach ($i in $Items) { if ($seen.Add($i)) { $i } }) -join '; ')
}
function Get-Present([object[]]$Items) { , @(foreach ($i in $Items) { if ($null -ne $i) { $i } }) }

# --- Users ------------------------------------------------------------------------------------
$userRows = [Collections.Generic.List[object]]::new()
$adUsersReady = Test-Ready 'AdUsers'
$exchangeReady = (Test-Ready 'Exo') -and (Test-Ready 'OnPrem')
foreach ($u in $sources.AdUsers.Rows) {
    $guid = $Fast::KF($u, 'ObjectGUID')
    if (-not $guid -or $adUserByGuid[$guid].Count -ne 1) { $excluded.UnkeyableUsers++; continue }
    $entra = Resolve-Entra $u
    $class = Get-Classification $entra.Row
    $inv = $class.Row; $ad = $u
    $persona = Get-Persona $inv $u
    $accountType = $Fast::F($u, 'AccountType'); $population = Get-Population $accountType
    $imperative = [Collections.Generic.List[string]]::new(); $secondary = [Collections.Generic.List[string]]::new()
    switch ($entra.Status) { 'Ambiguous' { $imperative.Add('Ambiguous Entra match') } 'Weak' { $imperative.Add('Weak Entra match (UPN only)') } 'Unknown' { $imperative.Add('Entra source not qualified') } }
    if ($population -eq 'Review required') { $secondary.Add('Account population requires classification review') }
    # Linked user mailboxes: EXO by Entra object or ImmutableId, on-premises by AD object.
    $exoLinked = @(); $onPremLinked = @()
    if ($null -ne $entra.Row) { $exoLinked += $Fast::List($exoByOwnerId, $Fast::KF($entra.Row, 'Object Id')) }
    $immutable = $Fast::F($u, 'ImmutableId_AD')
    if ($immutable -and $exoByImmutable.ContainsKey($immutable)) { foreach ($m in $exoByImmutable[$immutable]) { if ($exoLinked -notcontains $m) { $exoLinked += $m } } }
    $onPremLinked = @($Fast::List($onPremByOwnerGuid, $guid))
    $hasMailboxAd = $Fast::BF($u, 'HasMailbox')
    $mailboxKnown = $true; $hasMailbox = ''
    if ($exoLinked.Count -or $onPremLinked.Count) { $hasMailbox = 'Yes' }
    elseif (-not $hasMailboxAd -and $exchangeReady -and $adUsersReady) { $hasMailbox = 'No' }
    else { $hasMailbox = 'Unknown'; $mailboxKnown = $false; $imperative.Add('Mailbox presence unknown') }
    $holdState = Get-HoldState (@($exoLinked) + @($onPremLinked)) $mailboxKnown
    if ($holdState -eq 'No mailbox-level hold' -and $exoLinked.Count) { $secondary.Add('Organization retention policies not assessed') }
    $hosting = if ($exoLinked.Count -and $onPremLinked.Count) { 'Hybrid (both observed)' } elseif ($exoLinked.Count) { 'Exchange Online' } elseif ($onPremLinked.Count) { 'On-premises' } elseif ($hasMailbox -eq 'No') { 'None' } else { 'Unknown' }
    $exoMatch = if ($hasMailbox -eq 'Unknown') { 'Ambiguous match' } elseif ($exoLinked.Count -eq 1) { 'Qualified' } elseif ($exoLinked.Count -gt 1) { 'Ambiguous match' } else { 'Not applicable' }
    $onPremMatch = if ($hasMailbox -eq 'Unknown') { 'Ambiguous match' } elseif ($onPremLinked.Count -eq 1) { 'Qualified' } elseif ($onPremLinked.Count -gt 1) { 'Ambiguous match' } else { 'Not applicable' }
    if ($exoMatch -eq 'Ambiguous match' -and $hasMailbox -ne 'Unknown') { $imperative.Add('Several Exchange Online mailboxes linked') }
    if ($onPremMatch -eq 'Ambiguous match' -and $hasMailbox -ne 'Unknown') { $imperative.Add('Several on-premises mailboxes linked') }
    $exo = if ($exoLinked.Count -eq 1) { $exoLinked[0] } else { $null }
    $onPrem = if ($onPremLinked.Count -eq 1) { $onPremLinked[0] } else { $null }
    $activity = Get-ReportMatch $activityByUpn 'Activity' $entra.Row $entra.Status
    $usage = if ($exoMatch -eq 'Qualified') { Get-ReportMatch $usageByUpn 'MailboxUsage' $entra.Row $entra.Status } else { @{Match = 'Not applicable'; Row = $null} }
    $entraMatch = switch ($entra.Status) { 'Matched' { 'Qualified' } 'No Entra object' { 'Not applicable' } 'Weak' { 'Weak match' } default { 'Ambiguous match' } }
    $consoleMatch = if ($class.Status -eq 'Matched') { 'Qualified' } else { 'No report row' }
    $signals = @(
        $Fast::NewSignal('AD Last Logon', $Fast::F($u, 'LastLogonDateIso8601'), $sources.AdUsers.Status, 'Qualified', $R)
        $Fast::NewSignal('Entra Last Sign-In', $Fast::F($entra.Row, 'LastSignInDateTimeIso8601'), $sources.EntraUsers.Status, $entraMatch, $R)
        $Fast::NewSignal('Entra Last Non-Interactive Sign-In', $Fast::F($entra.Row, 'LastNonInteractiveSignInDateTimeIso8601'), $sources.EntraUsers.Status, $entraMatch, $R)
        $Fast::NewSignal('Entra Last Successful Sign-In', $Fast::F($entra.Row, 'LastSuccessfulSignInDateTimeIso8601'), $sources.EntraUsers.Status, $entraMatch, $R)
        $Fast::NewSignal('EXO Last User Action', $Fast::F($exo, 'LastUserActionTimeIso8601'), $sources.Exo.Status, $exoMatch, $R)
        $Fast::NewSignal('EXO Last Logon', $Fast::F($exo, 'LastLogonTimeIso8601'), $sources.Exo.Status, $exoMatch, $R)
        $Fast::NewSignal('On-Prem Last Logon', $Fast::F($onPrem, 'LastLogonTimeIso8601'), $sources.OnPrem.Status, $onPremMatch, $R)
        $Fast::NewSignal('Mailbox Usage Last Activity', $Fast::F($usage.Row, 'Last Activity Date'), $sources.MailboxUsage.Status, $usage.Match, $R)
        $Fast::NewSignal('M365 Last Activity', $Fast::F($activity.Row, 'LastActivityDate'), $sources.Activity.Status, $activity.Match, $R)
        $Fast::NewSignal('Console Last Activity', $Fast::F($inv, 'Console Last Activity UTC'), 'Observed', $consoleMatch, $R)
    )
    foreach ($s in $signals) { if ($s.Quality -in $notReady) { $imperative.Add("$($s.Name): $($s.Quality)") } }
    $latest = $Fast::Latest($signals)
    $created = Read-Created ($Fast::F($u, 'WhenCreatedIso8601'))
    $decision = Resolve-Decision -Latest $latest -Signals $signals -HoldState $holdState -NonHuman ($population -eq 'Non-human') -Imperative @($imperative) -Created $created -Secondary @($secondary)
    $enabled = $Fast::BF($u, 'Enabled')
    $action = switch ($decision.State) { 'Candidate' { if ($enabled) { 'Disable account' } else { 'Delete account' } } 'Review' { 'Review evidence' } default { 'Keep' } }
    $country = if ($null -ne $inv) { $Fast::F($inv, 'Country') } else { $Fast::F($u, 'CountryName', 'Country') }
    $row = [ordered]@{
        'User Cleanup Key' = New-Key 'cleanup-ad-user' $guid; 'AD Object GUID' = $guid; 'Entra Object ID' = ($Fast::F($entra.Row, 'Object Id'))
        'AD Domain' = $(if ($Fast::F($u, 'DomainName')) { $Fast::F($u, 'DomainName') } else { 'Unknown' }); 'SamAccountName' = ($Fast::F($u, 'SamAccountName')); 'User Principal Name' = ($Fast::F($u, 'UserPrincipalName')); 'Display Name' = ($Fast::F($u, 'DisplayName'))
        'Enabled' = $(if ($enabled) { 'True' } else { 'False' }); 'When Created' = ($Fast::Utc($created)); 'Password Last Set' = ($Fast::Utc((Read-Created ($Fast::F($u, 'PasswordLastSetDateIso8601'))))); 'Password Never Expires' = $(if ($Fast::BF($u, 'IsPasswordNeverExpires')) { 'True' } else { 'False' })
        'Identity Match Status' = $entra.Status; 'Identity Match Method' = $entra.Method; 'Classification Match Status' = $class.Status
        'Country' = $(if ($country) { $country } else { 'Unknown' }); 'Country Basis' = $(if ($null -ne $inv) { 'User Inventory Evidence' } elseif ($country) { 'AD attribute' } else { 'Unknown' })
        'Account Type' = $(if ($accountType) { $accountType } else { 'Unclassified Account' }); 'Account Population' = $population; 'Account Population Source' = 'AD AccountType'; 'Account Classification Rule Version' = $classificationRuleVersion
        'User Type' = $(if ($null -ne $inv -and ($Fast::F($inv, 'User Type'))) { $Fast::F($inv, 'User Type') } else { 'Unknown' }); 'Workforce Persona' = $persona.Persona; 'Persona Basis' = $persona.Basis; 'Account Type Source' = $(if ($null -ne $ad -and ($Fast::F($ad, 'AccountTypeSource'))) { $Fast::F($ad, 'AccountTypeSource') } elseif ($null -ne $ad) { 'SWI rules' } else { 'Unknown' })
        'Has Mailbox' = $hasMailbox; 'Hosting Location' = $hosting; 'Litigation Hold' = (Join-Flags (@($exoLinked) + @($onPremLinked)) 'LitigationHoldEnabled'); 'Retention Hold' = (Join-Flags (@($exoLinked) + @($onPremLinked)) 'RetentionHoldEnabled')
        'In-Place Holds' = $(if ($onPremLinked.Count) { $inPlace = 'False'; foreach ($x in $onPremLinked) { if ($Fast::F($x, 'InPlaceHolds')) { $inPlace = 'True' } }; $inPlace } elseif ($exoLinked.Count) { 'Not collected' } else { '' }); 'Hold State' = $holdState
    }
    foreach ($s in $signals) { $row[$s.Name] = $Fast::Utc($s.Utc); $row["$($s.Name) Quality"] = $s.Quality }
    $row['Entra Match Status'] = $entraMatch; $row['M365 Report Match Status'] = $activity.Match; $row['Mailbox Usage Match Status'] = $usage.Match
    $row['Last Activity Date'] = $Fast::Utc($(if ($null -ne $latest) { $latest.Utc } else { $null })); $row['Last Activity Source'] = $(if ($null -ne $latest) { $latest.Name } else { 'None' })
    $row['Days Since Last Activity'] = $(if ($null -ne $latest) { [string][math]::Floor(($R - $latest.Utc).TotalDays) } else { '' })
    $row['Inactivity Bucket'] = $decision.Bucket; $row['Bucket Sort'] = [string]$decision.BucketSort; $row['Cleanup State'] = $decision.State; $row['State Sort'] = [string]$decision.StateSort
    $row['Decision Reason'] = $decision.Reason; $row['Secondary Reasons'] = $decision.Secondary; $row['Recommended Action'] = $action
    $row['AD Source Status'] = $sources.AdUsers.Status; $row['AD Collected UTC'] = $Fast::Utc($sources.AdUsers.CollectedUtc)
    $row['Entra Source Status'] = $sources.EntraUsers.Status; $row['Entra Collected UTC'] = $Fast::Utc($sources.EntraUsers.CollectedUtc)
    $row['EXO Source Status'] = $sources.Exo.Status; $row['EXO Collected UTC'] = $Fast::Utc($sources.Exo.CollectedUtc)
    $row['On-Prem Source Status'] = $sources.OnPrem.Status; $row['On-Prem Collected UTC'] = $Fast::Utc($sources.OnPrem.CollectedUtc)
    $row['M365 Activity Source Status'] = $sources.Activity.Status; $row['M365 Activity Report Date'] = [string]$sources.Activity.RefreshDate
    $row['Reference UTC'] = $referenceText; $row['Evidence Date'] = $evidenceDate; $row['Evidence Status'] = $(if ($adUsersReady) { 'Observed' } else { $sources.AdUsers.Status })
    $userRows.Add([pscustomobject]$row)
}

# --- Computers --------------------------------------------------------------------------------
$computerRows = [Collections.Generic.List[object]]::new()
$adComputerByGuid = @{}; foreach ($c in $sources.AdComputers.Rows) { $Fast::Add($adComputerByGuid, $Fast::KF($c, 'ObjectGUID'), $c) }
foreach ($c in $sources.AdComputers.Rows) {
    $guid = $Fast::KF($c, 'ObjectGUID')
    if (-not $guid -or $adComputerByGuid[$guid].Count -ne 1) { $excluded.UnkeyableComputers++; continue }
    $os = $Fast::F($c, 'OperatingSystem', 'OperatingSystemShortName')
    if ($os -match '(?i)server') { $excluded.Servers++; continue }
    $imperative = [Collections.Generic.List[string]]::new()
    $scope = if ($os) { 'In scope' } else { 'Undetermined' }
    if ($scope -eq 'Undetermined') { $imperative.Add('Operating system unknown: workstation scope undetermined') }
    # Hybrid-joined workstation: Entra DeviceId = AD ObjectGUID; Intune AzureADDeviceId = Entra DeviceId.
    $entraHit = $Fast::One($entraDeviceById, $guid)
    $adCount = $Fast::F($c, 'EntraMatchCountByObjectGuid'); $adExists = $Fast::BF($c, 'ExistsInEntraByObjectGuid')
    $entraStatus = ''
    if (-not (Test-Ready 'EntraDevices')) { $entraStatus = 'Unknown' }
    elseif ($entraHit.Count -gt 1 -or ($adCount -and $adCount -notin @('0', '1'))) { $entraStatus = 'Ambiguous' }
    elseif ($entraHit.Count -eq 1 -and $adCount -eq '1') { $entraStatus = 'Matched' }
    elseif ($entraHit.Count -eq 0 -and -not $adExists -and $adCount -eq '0') { $entraStatus = 'No Entra object' }
    else { $entraStatus = 'Inconsistent' }
    if ($entraStatus -in @('Ambiguous', 'Inconsistent', 'Unknown')) { $imperative.Add("Entra device match: $entraStatus") }
    $entraMatch = switch ($entraStatus) { 'Matched' { 'Qualified' } 'No Entra object' { 'Not applicable' } default { 'Ambiguous match' } }
    $intuneStatus = ''; $intuneRow = $null
    if ($entraStatus -eq 'No Entra object') { $intuneStatus = 'Not applicable' }
    elseif ($entraStatus -ne 'Matched') { $intuneStatus = 'Unknown' }
    elseif (-not (Test-Ready 'Intune')) { $intuneStatus = 'Unknown' }
    else {
        $intuneHit = $Fast::One($intuneByEntraId, $guid)
        if ($intuneHit.Count -eq 1) { $intuneStatus = 'Matched'; $intuneRow = $intuneHit.Row } elseif ($intuneHit.Count -gt 1) { $intuneStatus = 'Ambiguous' } else { $intuneStatus = 'Not enrolled' }
    }
    if ($intuneStatus -eq 'Ambiguous') { $imperative.Add('Intune device match: Ambiguous') }
    $intuneMatch = switch ($intuneStatus) { 'Matched' { 'Qualified' } 'Not applicable' { 'Not applicable' } 'Not enrolled' { 'Not applicable' } 'Unknown' { 'Ambiguous match' } default { 'Ambiguous match' } }
    $signals = @(
        $Fast::NewSignal('AD Last Logon', $Fast::F($c, 'LastLogonDateIso8601'), $sources.AdComputers.Status, 'Qualified', $R)
        $Fast::NewSignal('Entra Last Sign-In', $Fast::F($c, 'EntraApproximateLastSignInDateTime'), $sources.EntraDevices.Status, $entraMatch, $R)
        $Fast::NewSignal('Intune Last Sync', $Fast::F($intuneRow, 'LastSyncDateTime'), $sources.Intune.Status, $intuneMatch, $R)
    )
    foreach ($s in $signals) { if ($s.Quality -in $notReady) { $imperative.Add("$($s.Name): $($s.Quality)") } }
    $latest = $Fast::Latest($signals)
    $created = Read-Created ($Fast::F($c, 'WhenCreatedIso8601'))
    $decision = Resolve-Decision -Latest $latest -Signals $signals -HoldState 'Not applicable' -NonHuman $false -Imperative @($imperative) -Created $created -Secondary @()
    $enabled = $Fast::BF($c, 'Enabled')
    $action = switch ($decision.State) { 'Candidate' { if ($enabled) { 'Disable computer' } else { 'Delete computer object' } } 'Review' { 'Review evidence' } default { 'Keep' } }
    $deviceHit = $Fast::One($deviceByAdGuid, $guid)
    $row = [ordered]@{
        'Computer Cleanup Key' = New-Key 'cleanup-ad-computer' $guid; 'AD Object GUID' = $guid; 'Entra Device ID' = $(if ($entraStatus -eq 'Matched') { $guid } else { '' }); 'Intune Device ID' = ($Fast::F($intuneRow, 'ManagedDeviceId'))
        'AD Domain' = $(if ($Fast::F($c, 'DomainName')) { $Fast::F($c, 'DomainName') } else { 'Unknown' }); 'Computer Name' = (($Fast::F($c, 'Name', 'SamAccountName')) -replace '\$$', ''); 'Enabled' = $(if ($enabled) { 'True' } else { 'False' })
        'Operating System' = $(if ($os) { $os } else { 'Unknown' }); 'Scope State' = $scope; 'In Disabled OU' = $(if ($Fast::BF($c, 'IsInDisabledObjectsOU')) { 'True' } else { 'False' })
        'When Created' = ($Fast::Utc($created)); 'Password Last Set' = ($Fast::Utc((Read-Created ($Fast::F($c, 'pwdLastSetDateIso8601')))))
        'Country' = $(if ($deviceHit.Count -eq 1 -and ($Fast::F($deviceHit.Row, 'Country'))) { $Fast::F($deviceHit.Row, 'Country') } else { 'Unknown' }); 'Country Basis' = $(if ($deviceHit.Count -eq 1 -and ($Fast::F($deviceHit.Row, 'Country Basis'))) { $Fast::F($deviceHit.Row, 'Country Basis') } elseif ($deviceHit.Count -gt 1) { 'Ambiguous device match' } else { 'Not matched' })
        'Account Type' = 'Not applicable'; 'Account Population' = 'Not applicable'; 'Workforce Persona' = 'Not applicable'
        'Entra Match Status' = $entraStatus; 'Intune Match Status' = $intuneStatus
    }
    foreach ($s in $signals) { $row[$s.Name] = $Fast::Utc($s.Utc); $row["$($s.Name) Quality"] = $s.Quality }
    $row['Last Activity Date'] = $Fast::Utc($(if ($null -ne $latest) { $latest.Utc } else { $null })); $row['Last Activity Source'] = $(if ($null -ne $latest) { $latest.Name } else { 'None' })
    $row['Days Since Last Activity'] = $(if ($null -ne $latest) { [string][math]::Floor(($R - $latest.Utc).TotalDays) } else { '' })
    $row['Inactivity Bucket'] = $decision.Bucket; $row['Bucket Sort'] = [string]$decision.BucketSort; $row['Cleanup State'] = $decision.State; $row['State Sort'] = [string]$decision.StateSort
    $row['Decision Reason'] = $decision.Reason; $row['Secondary Reasons'] = $decision.Secondary; $row['Recommended Action'] = $action
    $row['AD Source Status'] = $sources.AdComputers.Status; $row['AD Collected UTC'] = $Fast::Utc($sources.AdComputers.CollectedUtc)
    $row['Entra Source Status'] = $sources.EntraDevices.Status; $row['Entra Collected UTC'] = $Fast::Utc($sources.EntraDevices.CollectedUtc)
    $row['Intune Source Status'] = $sources.Intune.Status; $row['Intune Collected UTC'] = $Fast::Utc($sources.Intune.CollectedUtc)
    $row['Reference UTC'] = $referenceText; $row['Evidence Date'] = $evidenceDate; $row['Evidence Status'] = $(if (Test-Ready 'AdComputers') { 'Observed' } else { $sources.AdComputers.Status })
    $computerRows.Add([pscustomobject]$row)
}

# --- Mailboxes --------------------------------------------------------------------------------
$mailboxRows = [Collections.Generic.List[object]]::new()
$sides = [Collections.Generic.List[object]]::new()
foreach ($m in $exoUser) { $sides.Add([pscustomobject]@{Side = 'exo'; Guid = ($Fast::KF($m, 'MailboxGuid')); Row = $m}) }
foreach ($m in $onPremUser) { $sides.Add([pscustomobject]@{Side = 'onprem'; Guid = ($Fast::KF($m, 'ExchangeGuid')); Row = $m}) }
function Resolve-Owner($Side) {
    # EXO: AD user by ImmutableId, else through the Entra object and its SID; on-premises: AD ObjectGUID.
    $result = [ordered]@{Status = 'Not found'; Ad = $null; Entra = $null; EntraStatus = 'Unknown'}
    if ($Side.Side -eq 'exo') {
        $entraHit = $Fast::One($entraById, $Fast::KF($Side.Row, 'ExternalDirectoryObjectId'))
        if (-not (Test-Ready 'EntraUsers')) { $result.EntraStatus = 'Unknown' }
        elseif ($entraHit.Count -eq 1) { $result.Entra = $entraHit.Row; $result.EntraStatus = 'Matched' }
        elseif ($entraHit.Count -gt 1) { $result.EntraStatus = 'Ambiguous' } else { $result.EntraStatus = 'Not found' }
        $adHit = $Fast::One($adUserByImmutable, $Fast::F($Side.Row, 'OnPremisesImmutableId', 'ImmutableId'))
        if ($adHit.Count -eq 0 -and $null -ne $result.Entra) { $adHit = $Fast::One($adUserBySid, $Fast::KF($result.Entra, 'OnPremisesSecurityIdentifier')) }
        if ($adHit.Count -gt 1 -or $entraHit.Count -gt 1) { $result.Status = 'Ambiguous'; return [pscustomobject]$result }
        if ($adHit.Count -eq 1) { $result.Status = 'Matched'; $result.Ad = $adHit.Row; return [pscustomobject]$result }
        if ($null -ne $result.Entra -and -not ($Fast::F($result.Entra, 'OnPremisesImmutableId')) -and -not ($Fast::F($result.Entra, 'OnPremisesSecurityIdentifier'))) { $result.Status = 'Matched (cloud-only)'; return [pscustomobject]$result }
        return [pscustomobject]$result
    }
    $adHit = $Fast::One($adUserByGuid, $Fast::KF($Side.Row, 'ObjectGUID'))
    if ($adHit.Count -gt 1) { $result.Status = 'Ambiguous'; return [pscustomobject]$result }
    if ($adHit.Count -eq 1) {
        $result.Status = 'Matched'; $result.Ad = $adHit.Row
        $entra = Resolve-Entra $adHit.Row; $result.EntraStatus = $entra.Status; if ($entra.Status -eq 'Matched') { $result.Entra = $entra.Row }
    }
    return [pscustomobject]$result
}
$groups = @{}
foreach ($s in $sides) { if (-not $s.Guid) { $excluded.UnkeyableMailboxes++; continue }; $Fast::Add($groups, $s.Guid, $s) }
foreach ($guid in @($groups.Keys | Sort-Object)) {
    $members = @($groups[$guid])
    $exoSides = @(foreach ($x in $members) { if ($x.Side -eq 'exo') { $x } }); $onPremSides = @(foreach ($x in $members) { if ($x.Side -eq 'onprem') { $x } })
    if ($exoSides.Count -gt 1 -or $onPremSides.Count -gt 1) { $excluded.UnkeyableMailboxes += $members.Count; continue }
    $units = @()
    if ($exoSides.Count -eq 1 -and $onPremSides.Count -eq 1) {
        $exoOwner = Resolve-Owner $exoSides[0]; $onPremOwner = Resolve-Owner $onPremSides[0]
        $coherent = $exoOwner.Status -eq 'Matched' -and $onPremOwner.Status -eq 'Matched' -and [object]::ReferenceEquals($exoOwner.Ad, $onPremOwner.Ad)
        if ($coherent) { $units += [pscustomobject]@{Key = New-Key 'cleanup-mailbox' $guid; Exo = $exoSides[0]; OnPrem = $onPremSides[0]; Owner = $exoOwner; Conflict = $false} }
        else {
            $units += [pscustomobject]@{Key = New-Key 'cleanup-mailbox-exo' $guid; Exo = $exoSides[0]; OnPrem = $null; Owner = $exoOwner; Conflict = $true}
            $units += [pscustomobject]@{Key = New-Key 'cleanup-mailbox-onprem' $guid; Exo = $null; OnPrem = $onPremSides[0]; Owner = $onPremOwner; Conflict = $true}
        }
    } elseif ($exoSides.Count -eq 1) { $units += [pscustomobject]@{Key = New-Key 'cleanup-mailbox' $guid; Exo = $exoSides[0]; OnPrem = $null; Owner = (Resolve-Owner $exoSides[0]); Conflict = $false} }
    else { $units += [pscustomobject]@{Key = New-Key 'cleanup-mailbox' $guid; Exo = $null; OnPrem = $onPremSides[0]; Owner = (Resolve-Owner $onPremSides[0]); Conflict = $false} }
    foreach ($unit in $units) {
        $exo = if ($null -ne $unit.Exo) { $unit.Exo.Row } else { $null }; $onPrem = if ($null -ne $unit.OnPrem) { $unit.OnPrem.Row } else { $null }
        $owner = $unit.Owner; $ad = $owner.Ad; $entraRow = $owner.Entra
        $imperative = [Collections.Generic.List[string]]::new(); $secondary = [Collections.Generic.List[string]]::new()
        if ($unit.Conflict) { $imperative.Add('Conflicting hybrid identity') }
        if ($null -ne $exo -and $null -ne $onPrem) { $secondary.Add('Hybrid duplicate') }
        switch ($owner.Status) { 'Not found' { $imperative.Add('Owner not found') } 'Ambiguous' { $imperative.Add('Owner match ambiguous') } }
        $class = Get-Classification $entraRow; $inv = $class.Row
        $persona = Get-Persona $inv $ad
        $accountType = if ($null -ne $ad) { $Fast::F($ad, 'AccountType') } else { '' }
        $population = if ($null -ne $ad) { Get-Population $accountType } else { 'Review required' }
        if ($population -eq 'Review required') { $secondary.Add('Account population requires classification review') }
        $present = Get-Present @($exo, $onPrem)
        $holdState = Get-HoldState $present $true
        if ($holdState -eq 'No mailbox-level hold' -and $null -ne $exo) { $secondary.Add('Organization retention policies not assessed') }
        $ownerAdMatch = if ($null -ne $ad) { 'Qualified' } elseif ($owner.Status -eq 'Matched (cloud-only)') { 'Not applicable' } else { 'Ambiguous match' }
        $ownerEntraMatch = switch ($owner.EntraStatus) { 'Matched' { 'Qualified' } 'No Entra object' { 'Not applicable' } 'Weak' { 'Weak match' } default { 'Ambiguous match' } }
        if ($owner.EntraStatus -in @('Ambiguous', 'Weak', 'Not found')) { $imperative.Add("Owner Entra match: $($owner.EntraStatus)") }
        $mailboxUpn = $Fast::KF($exo, 'UserPrincipalName')
        $usageMatch = 'Not applicable'; $usageRow = $null
        if ($null -ne $exo -and (Test-Ready 'MailboxUsage')) { $hit = $Fast::One($usageByUpn, $mailboxUpn); if ($hit.Count -eq 1) { $usageMatch = 'Qualified'; $usageRow = $hit.Row } elseif ($hit.Count -gt 1) { $usageMatch = 'Ambiguous match' } else { $usageMatch = 'No report row' } }
        elseif ($null -ne $exo) { $usageMatch = 'Qualified' }
        $activity = Get-ReportMatch $activityByUpn 'Activity' $entraRow $owner.EntraStatus
        $ownerSignIn = @(foreach ($column in @('LastSignInDateTimeIso8601', 'LastNonInteractiveSignInDateTimeIso8601', 'LastSuccessfulSignInDateTimeIso8601')) { $Fast::NewSignal('Owner Entra Last Sign-In', $Fast::F($entraRow, $column), $sources.EntraUsers.Status, $ownerEntraMatch, $R) })
        $ownerSignInLatest = $Fast::Latest($ownerSignIn)
        $ownerSignInSignal = if ($null -ne $ownerSignInLatest) { $ownerSignInLatest } else { $ownerSignIn[0] }
        $signals = @(
            $Fast::NewSignal('EXO Last User Action', $Fast::F($exo, 'LastUserActionTimeIso8601'), $sources.Exo.Status, $(if ($null -ne $exo) { 'Qualified' } else { 'Not applicable' }), $R)
            $Fast::NewSignal('EXO Last Logon', $Fast::F($exo, 'LastLogonTimeIso8601'), $sources.Exo.Status, $(if ($null -ne $exo) { 'Qualified' } else { 'Not applicable' }), $R)
            $Fast::NewSignal('On-Prem Last Logon', $Fast::F($onPrem, 'LastLogonTimeIso8601'), $sources.OnPrem.Status, $(if ($null -ne $onPrem) { 'Qualified' } else { 'Not applicable' }), $R)
            $Fast::NewSignal('Mailbox Usage Last Activity', $Fast::F($usageRow, 'Last Activity Date'), $sources.MailboxUsage.Status, $usageMatch, $R)
            $Fast::NewSignal('Owner AD Last Logon', $Fast::F($ad, 'LastLogonDateIso8601'), $sources.AdUsers.Status, $ownerAdMatch, $R)
            $ownerSignInSignal
            $Fast::NewSignal('Owner M365 Last Activity', $Fast::F($activity.Row, 'LastActivityDate'), $sources.Activity.Status, $activity.Match, $R)
            $Fast::NewSignal('Owner Console Last Activity', $Fast::F($inv, 'Console Last Activity UTC'), 'Observed', $(if ($class.Status -eq 'Matched') { 'Qualified' } else { 'No report row' }), $R)
        )
        foreach ($s in $signals) { if ($s.Quality -in $notReady) { $imperative.Add("$($s.Name): $($s.Quality)") } }
        $latest = $Fast::Latest($signals)
        $created = Read-Created ($Fast::F($exo, 'WhenMailboxCreatedIso8601')); if ($null -eq $created) { $created = Read-Created ($Fast::F($onPrem, 'WhenMailboxCreatedIso8601')) }
        $decision = Resolve-Decision -Latest $latest -Signals $signals -HoldState $holdState -NonHuman ($population -eq 'Non-human') -Imperative @($imperative) -Created $created -Secondary @($secondary)
        $ownerState = if ($null -ne $ad) { if ($Fast::BF($ad, 'Enabled')) { 'Enabled' } else { 'Disabled' } } elseif ($null -ne $entraRow) { if ($Fast::BF($entraRow, 'AccountEnabled')) { 'Enabled' } else { 'Disabled' } } elseif ($owner.Status -eq 'Ambiguous') { 'Ambiguous' } else { 'Not found' }
        $action = switch ($decision.State) { 'Candidate' { if ($ownerState -eq 'Disabled') { 'Remove mailbox after retention check' } else { 'Review owner' } } 'Review' { if ($owner.Status -in @('Not found', 'Ambiguous')) { 'Review owner' } else { 'Review evidence' } } default { 'Keep' } }
        $sizeGb = ''
        if ($null -ne $usageRow -and ($Fast::F($usageRow, 'Storage Used (Byte)')) -match '^\d+$') { $sizeGb = ([double]($Fast::F($usageRow, 'Storage Used (Byte)')) / 1GB).ToString('0.###', [Globalization.CultureInfo]::InvariantCulture) }
        elseif ($null -ne $onPrem -and $null -ne $ad -and ($Fast::F($ad, 'OnPremMailboxSizeMB')) -match '^\d+([.,]\d+)?$') { $sizeGb = ([double](($Fast::F($ad, 'OnPremMailboxSizeMB')) -replace ',', '.') / 1024).ToString('0.###', [Globalization.CultureInfo]::InvariantCulture) }
        $country = if ($null -ne $inv) { $Fast::F($inv, 'Country') } elseif ($null -ne $ad) { $Fast::F($ad, 'CountryName', 'Country') } else { '' }
        $row = [ordered]@{
            'Mailbox Cleanup Key' = $unit.Key; 'Mailbox Identity Basis' = 'Mailbox GUID'; 'Mailbox GUID' = $guid
            'Primary SMTP Address' = $(if ($null -ne $exo) { $Fast::F($exo, 'PrimarySmtpAddress') } else { $Fast::F($onPrem, 'PrimarySMTPaddress', 'PrimarySmtpAddress') })
            'Hosting Location' = $(if ($null -ne $exo -and $null -ne $onPrem) { 'Hybrid (both observed)' } elseif ($null -ne $exo) { 'Exchange Online' } else { 'On-premises' })
            'Recipient Type' = $(if ($null -ne $exo) { $Fast::F($exo, 'RecipientTypeDetails') } else { $Fast::F($onPrem, 'RecipientType') }); 'Scope State' = 'In scope'; 'When Mailbox Created' = ($Fast::Utc($created))
            'Owner AD Object GUID' = ($Fast::KF($ad, 'ObjectGUID')); 'Owner Entra Object ID' = ($Fast::F($entraRow, 'Object Id')); 'Owner Match Status' = $owner.Status; 'Owner Account State' = $ownerState
            'AD Domain' = $(if ($Fast::F($ad, 'DomainName')) { $Fast::F($ad, 'DomainName') } elseif ($Fast::F($onPrem, 'DomainName')) { $Fast::F($onPrem, 'DomainName') } else { 'Unknown' })
            'Country' = $(if ($country) { $country } else { 'Unknown' }); 'Country Basis' = $(if ($null -ne $inv) { 'User Inventory Evidence' } elseif ($country) { 'AD attribute' } else { 'Unknown' })
            'Account Type' = $(if ($accountType) { $accountType } elseif ($null -ne $ad) { 'Unclassified Account' } else { 'Unknown' }); 'Account Population' = $population; 'Account Population Source' = $(if ($null -ne $ad) { 'AD AccountType' } else { 'Owner not in AD' }); 'Account Classification Rule Version' = $classificationRuleVersion
            'User Type' = $(if ($null -ne $inv -and ($Fast::F($inv, 'User Type'))) { $Fast::F($inv, 'User Type') } else { 'Unknown' }); 'Workforce Persona' = $persona.Persona; 'Persona Basis' = $persona.Basis; 'Account Type Source' = $(if ($null -ne $ad -and ($Fast::F($ad, 'AccountTypeSource'))) { $Fast::F($ad, 'AccountTypeSource') } elseif ($null -ne $ad) { 'SWI rules' } else { 'Unknown' }); 'Classification Match Status' = $class.Status
            'Litigation Hold' = (Join-Flags $present 'LitigationHoldEnabled'); 'Retention Hold' = (Join-Flags $present 'RetentionHoldEnabled')
            'In-Place Holds' = $(if ($null -ne $onPrem) { if ($Fast::F($onPrem, 'InPlaceHolds')) { 'True' } else { 'False' } } else { 'Not collected' }); 'Hold State' = $holdState; 'Mailbox Size GB' = $sizeGb
        }
        foreach ($s in $signals) { $row[$s.Name] = $Fast::Utc($s.Utc); $row["$($s.Name) Quality"] = $s.Quality }
        $row['Mailbox Usage Match Status'] = $usageMatch; $row['Owner M365 Report Match Status'] = $activity.Match
        $row['Last Activity Date'] = $Fast::Utc($(if ($null -ne $latest) { $latest.Utc } else { $null })); $row['Last Activity Source'] = $(if ($null -ne $latest) { $latest.Name } else { 'None' })
        $row['Days Since Last Activity'] = $(if ($null -ne $latest) { [string][math]::Floor(($R - $latest.Utc).TotalDays) } else { '' })
        $row['Inactivity Bucket'] = $decision.Bucket; $row['Bucket Sort'] = [string]$decision.BucketSort; $row['Cleanup State'] = $decision.State; $row['State Sort'] = [string]$decision.StateSort
        $row['Decision Reason'] = $decision.Reason; $row['Secondary Reasons'] = $decision.Secondary; $row['Recommended Action'] = $action
        $row['EXO Source Status'] = $sources.Exo.Status; $row['EXO Collected UTC'] = $Fast::Utc($sources.Exo.CollectedUtc)
        $row['On-Prem Source Status'] = $sources.OnPrem.Status; $row['On-Prem Collected UTC'] = $Fast::Utc($sources.OnPrem.CollectedUtc)
        $row['Mailbox Usage Source Status'] = $sources.MailboxUsage.Status; $row['Mailbox Usage Report Date'] = [string]$sources.MailboxUsage.RefreshDate
        $row['Reference UTC'] = $referenceText; $row['Evidence Date'] = $evidenceDate; $row['Evidence Status'] = 'Observed'
        $mailboxRows.Add([pscustomobject]$row)
    }
}

# --- Outputs ----------------------------------------------------------------------------------
function Write-Evidence([string]$Path, [object[]]$Rows, [string[]]$Header) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    # The Power BI readers select columns strictly: an empty table still publishes its header.
    if ($Rows.Count) { $Rows | Select-Object $Header | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8 }
    else { [IO.File]::WriteAllText($Path, (($Header | ForEach-Object { '"' + $_ + '"' }) -join ',') + [Environment]::NewLine, [Text.UTF8Encoding]::new($false)) }
}
$signalColumns = { param([string[]]$Names) foreach ($n in $Names) { $n; "$n Quality" } }
$userHeader = @('User Cleanup Key', 'AD Object GUID', 'Entra Object ID', 'AD Domain', 'SamAccountName', 'User Principal Name', 'Display Name', 'Enabled', 'When Created', 'Password Last Set', 'Password Never Expires',
    'Identity Match Status', 'Identity Match Method', 'Classification Match Status', 'Country', 'Country Basis', 'Account Type', 'Account Population', 'Account Population Source', 'Account Classification Rule Version', 'User Type', 'Workforce Persona', 'Persona Basis', 'Account Type Source',
    'Has Mailbox', 'Hosting Location', 'Litigation Hold', 'Retention Hold', 'In-Place Holds', 'Hold State') +
    @(& $signalColumns @('AD Last Logon', 'Entra Last Sign-In', 'Entra Last Non-Interactive Sign-In', 'Entra Last Successful Sign-In', 'EXO Last User Action', 'EXO Last Logon', 'On-Prem Last Logon', 'Mailbox Usage Last Activity', 'M365 Last Activity', 'Console Last Activity')) +
    @('Entra Match Status', 'M365 Report Match Status', 'Mailbox Usage Match Status', 'Last Activity Date', 'Last Activity Source', 'Days Since Last Activity', 'Inactivity Bucket', 'Bucket Sort', 'Cleanup State', 'State Sort', 'Decision Reason', 'Secondary Reasons', 'Recommended Action',
    'AD Source Status', 'AD Collected UTC', 'Entra Source Status', 'Entra Collected UTC', 'EXO Source Status', 'EXO Collected UTC', 'On-Prem Source Status', 'On-Prem Collected UTC', 'M365 Activity Source Status', 'M365 Activity Report Date', 'Reference UTC', 'Evidence Date', 'Evidence Status')
$computerHeader = @('Computer Cleanup Key', 'AD Object GUID', 'Entra Device ID', 'Intune Device ID', 'AD Domain', 'Computer Name', 'Enabled', 'Operating System', 'Scope State', 'In Disabled OU', 'When Created', 'Password Last Set', 'Country', 'Country Basis',
    'Account Type', 'Account Population', 'Workforce Persona', 'Entra Match Status', 'Intune Match Status') +
    @(& $signalColumns @('AD Last Logon', 'Entra Last Sign-In', 'Intune Last Sync')) +
    @('Last Activity Date', 'Last Activity Source', 'Days Since Last Activity', 'Inactivity Bucket', 'Bucket Sort', 'Cleanup State', 'State Sort', 'Decision Reason', 'Secondary Reasons', 'Recommended Action',
    'AD Source Status', 'AD Collected UTC', 'Entra Source Status', 'Entra Collected UTC', 'Intune Source Status', 'Intune Collected UTC', 'Reference UTC', 'Evidence Date', 'Evidence Status')
$mailboxHeader = @('Mailbox Cleanup Key', 'Mailbox Identity Basis', 'Mailbox GUID', 'Primary SMTP Address', 'Hosting Location', 'Recipient Type', 'Scope State', 'When Mailbox Created', 'Owner AD Object GUID', 'Owner Entra Object ID', 'Owner Match Status', 'Owner Account State',
    'AD Domain', 'Country', 'Country Basis', 'Account Type', 'Account Population', 'Account Population Source', 'Account Classification Rule Version', 'User Type', 'Workforce Persona', 'Persona Basis', 'Account Type Source', 'Classification Match Status',
    'Litigation Hold', 'Retention Hold', 'In-Place Holds', 'Hold State', 'Mailbox Size GB') +
    @(& $signalColumns @('EXO Last User Action', 'EXO Last Logon', 'On-Prem Last Logon', 'Mailbox Usage Last Activity', 'Owner AD Last Logon', 'Owner Entra Last Sign-In', 'Owner M365 Last Activity', 'Owner Console Last Activity')) +
    @('Mailbox Usage Match Status', 'Owner M365 Report Match Status', 'Last Activity Date', 'Last Activity Source', 'Days Since Last Activity', 'Inactivity Bucket', 'Bucket Sort', 'Cleanup State', 'State Sort', 'Decision Reason', 'Secondary Reasons', 'Recommended Action',
    'EXO Source Status', 'EXO Collected UTC', 'On-Prem Source Status', 'On-Prem Collected UTC', 'Mailbox Usage Source Status', 'Mailbox Usage Report Date', 'Reference UTC', 'Evidence Date', 'Evidence Status')
Write-Evidence $UserOutputPath @($userRows) $userHeader
Write-Evidence $ComputerOutputPath @($computerRows) $computerHeader
Write-Evidence $MailboxOutputPath @($mailboxRows) $mailboxHeader

$counts = @{AdUsers = @{Excluded = 0; Unkeyable = $excluded.UnkeyableUsers}; AdComputers = @{Excluded = $excluded.Servers; Unkeyable = $excluded.UnkeyableComputers}; Exo = @{Excluded = 0; Unkeyable = 0}; OnPrem = @{Excluded = 0; Unkeyable = 0}}
foreach ($name in $sources.Keys) {
    $s = $sources[$name]
    $qualification.Add([pscustomobject][ordered]@{
        'Source' = $s.Name; 'File' = $s.File; 'Producer' = $s.Producer; 'Status' = $s.Status; 'Reason' = $s.Reason; 'Rows' = [string]@($s.Rows).Count
        'Collected UTC' = ($Fast::Utc($s.CollectedUtc)); 'Report Refresh Date' = [string]$s.RefreshDate
        'Out Of Scope Rows' = [string]$(if ($counts.ContainsKey($name)) { $counts[$name].Excluded } else { 0 }); 'Unkeyable Rows' = [string]$(if ($counts.ContainsKey($name)) { $counts[$name].Unkeyable } else { 0 })
        'Reference UTC' = $referenceText; 'Evidence Date' = $evidenceDate
    })
}
$qualification.Add([pscustomobject][ordered]@{'Source' = 'User mailboxes (EXO and on-premises)'; 'File' = ''; 'Producer' = ''; 'Status' = $(if ($exchangeReady) { 'Observed' } else { 'Not ready' }); 'Reason' = 'Shared, room, equipment and discovery mailboxes are out of scope'; 'Rows' = [string]$mailboxRows.Count; 'Collected UTC' = ''; 'Report Refresh Date' = ''; 'Out Of Scope Rows' = [string]$excluded.SharedOrResourceMailboxes; 'Unkeyable Rows' = [string]$excluded.UnkeyableMailboxes; 'Reference UTC' = $referenceText; 'Evidence Date' = $evidenceDate})
Write-Evidence $QualificationOutputPath @($qualification) @('Source', 'File', 'Producer', 'Status', 'Reason', 'Rows', 'Collected UTC', 'Report Refresh Date', 'Out Of Scope Rows', 'Unkeyable Rows', 'Reference UTC', 'Evidence Date')

if ($excluded.UnkeyableMailboxes) { Write-Warning "Directory cleanup: $($excluded.UnkeyableMailboxes) mailbox rows without a unique mailbox GUID were not published." }
if ($excluded.UnkeyableUsers -or $excluded.UnkeyableComputers) { Write-Warning "Directory cleanup: $($excluded.UnkeyableUsers) user and $($excluded.UnkeyableComputers) computer rows without a unique AD ObjectGUID were not published." }
foreach ($s in $sources.Values) { if ($s.Status -ne 'Observed') { Write-Warning "Directory cleanup source $($s.Name): $($s.Status) ($($s.Reason)); affected objects are Review." } }
$summary = foreach ($pair in @(@('users', $userRows), @('computers', $computerRows), @('mailboxes', $mailboxRows))) {
    $byState = @($pair[1] | Group-Object 'Cleanup State' | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
    "$($pair[0]) $($pair[1].Count) ($byState)"
}
Write-Host ('Directory cleanup evidence: ' + ($summary -join '; ') + "; reference $referenceText.")

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAwn2hUaRv9SBOl
# GzDFujc1B7hZzE4Rsxu8IEUceRYPCKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP4RQ9C51xpKee6QheZwLJYQSGvxs3g/k+mI28yohhtqMA0GCSqG
# SIb3DQEBAQUABIIBgISUiTv2EyzZjwnRJp+6SZ8P++xSkn8xJ5NH4NUVDUVEuFL8
# lqe/ZZgSfWRSkhAOIqa73Q3Ap27/MkvFPxa1aCR/aCQX+m0orkR3pHFKXu3kA1FB
# J78tZMjvVzx6oiNtRFTJgd9ikuGwsC4WwTzwhG303P38ggqbK98tHy0RUfm7/AZY
# oLyGwAvltFimXhKLCi7YE0PHTnfwxy8i6EQT/HztJq3+ODGCTNW8zz0r2PktGYQe
# rimn89P+mCkxKEQPs9HbZC6pJiWn/n5f2S4g2hT/vJtcpU+5IHqZt+K7KyFp0L7q
# Hw8qkgz8OKX0Xn/+naj8ouAn1uYj54HUrUlE5u/VGQ+WSU7gzvQ3jrI/8quuHGRH
# M62I0yn+FlfOtyDEDzagpMEHoPIeJzoec6mrzpwLM8MQSDf84yZJaJi1WbX3XLLY
# lY3HfFoObF5WSyRtchFNI6PDxysT1qGslItYCa3WSgGnWjABEdDGn3KwPECvRjU7
# iPgoEsb+cSzU00JXmqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkwMjE1
# MjVaMC8GCSqGSIb3DQEJBDEiBCB9wvVH6b6vWEhqKrK/KTc0yQYpTZsePqLPzoHu
# SN2D0jANBgkqhkiG9w0BAQEFAASCAgCeSb50PUSThdaVMjzzEyHX3SCnR5yv4VUJ
# zFh6yBDJ/rX6kqMQ2+PIAnMIr1qNwXfCawzfj0i/LjPLlq1n8WlFb4GFdCNAq5UA
# Nk7U/SCY74ASztCAOXzD1EbS81cizjOSEmYSTj4s3fJo1ievdAudchfBKrLo0Y1l
# ZsPqHC2qAWttmQDW93NXhdf7kcxEN030wSEbs2/l2bDnbIJDdN0wP/bq63ls8+FP
# p2sEEfKJo2AV+AsD0eHUByAM66wHGX+soHeSYNBBQ5GHVlirtb58v89CNTyjquOt
# URaRp6IjvYIeAPkmWi+Tfo3Or5o3UqifNlk9i+o5l7lavdeX5g2qltCDBpiqQNES
# y/g4zmiDLRz48l+r7ss4iX0tXKftNrD8yJUQ+6iGNKvYRVjt+LBEoO3D+SZ5cWPB
# aaFWwigV3NBcjSj0J7QOvCMxX+KumPl/6kAVR1josoAFrJWqv1XyIbhHq05HD8rg
# oCFQm+Y8pydm4cbpZYQm7CLBcw148SQTzo0S4N4csmqh76bdehAi9b//Y8JoDVbG
# aBZg+WUcTjGZgAP685Ay0cpM0Z7k12RAPxqwRDS/5Tfik0gX2+vriJcmsldHDWaz
# sHwvSu7NpvbTiQWSC1h4j9QUm+Yb0jlsAMdKxmV9mo6dABiBpikPPOTWXOOZqani
# ytjge/Iovg==
# SIG # End signature block
