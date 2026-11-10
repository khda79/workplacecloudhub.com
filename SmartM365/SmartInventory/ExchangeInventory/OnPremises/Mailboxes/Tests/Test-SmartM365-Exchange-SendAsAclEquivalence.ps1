<#
.SYNOPSIS
Read-only qualification: compares the current Exchange Get-ADPermission SendAs
collection with a direct Active Directory DACL read, mailbox by mailbox.

.DESCRIPTION
For every native on-premises mailbox (or a filtered sample), this script computes
the SendAs value twice:
  - Old method: the exact Get-ADPermission filter used by
    SmartM365-Exchange-Local-Mailboxes-Inventory.ps1 (IncludeADPermission).
  - New method: nTSecurityDescriptor read with Get-ADObject, explicit ACEs only,
    extended rights whose name contains 'Send', SID translated locally, same
    string filters as the old method.
It writes per-mailbox details, a per-domain summary, the extended-right map and a
log, then prints a verdict. Nothing is written to Exchange or Active Directory.

Run on an Exchange 2016 inventory server (Windows PowerShell 5.1), with the same
service account as the orchestrator job.

.VERSION
1.0.0

.PARAMETER DomainName
Optional list of domain FQDNs to test (default: all domains of the forest).

.PARAMETER MaxMailboxesPerDomain
Optional sample size per domain (0 = all mailboxes).

.PARAMETER OutputDirectory
Output folder. Default: ProgramData\SmartM365\Qualification\SendAsAcl\<timestamp>.

.PARAMETER SelfTestOnly
Runs the offline filter tests only (no Exchange, no Active Directory access).
#>
[CmdletBinding()]
param(
    [string[]]$DomainName = @(),
    [int]$MaxMailboxesPerDomain = 0,
    [string]$OutputDirectory = '',
    [switch]$SelfTestOnly
)

# No StrictMode: the old-method block must behave exactly like production
# (for example .Count on a single unique SendAs string).
$ErrorActionPreference = 'Stop'
$script:QualificationVersion = '1.0.0'
$script:LogFile = $null

function Write-QualificationLog {
    param([Parameter(Mandatory = $true)][string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 }
}

function Get-SmartM365SendAsFromSecurityDescriptor {
    <#
    Candidate replacement for Get-ADPermission SendAs collection.
    Keeps the old semantics: explicit (non-inherited) ACEs whose extended right
    name matches '*Send*', Allow and Deny alike (the old filter does not test Deny),
    'NT AUTHORITY\SELF' and unresolved S-1-5-21 SIDs excluded, case-sensitive unique.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.DirectoryServices.ActiveDirectorySecurity]$SecurityDescriptor,
        [Parameter(Mandatory = $true)][hashtable]$SendRightGuids,
        [Parameter(Mandatory = $true)][scriptblock]$ResolveSid
    )

    $all = $SecurityDescriptor.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])
    $extendedRight = [System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight
    $users = @()
    $explicitCount = 0
    $matchedCount = 0
    $denyMatchedCount = 0
    $allExtendedRightsCount = 0
    foreach ($rule in $all) {
        if ($rule.IsInherited) { continue }
        $explicitCount++
        if (($rule.ActiveDirectoryRights -band $extendedRight) -ne $extendedRight) { continue }
        if ($rule.ObjectType -eq [guid]::Empty) { $allExtendedRightsCount++; continue }
        if (-not $SendRightGuids.ContainsKey($rule.ObjectType.ToString('D'))) { continue }
        $matchedCount++
        if ($rule.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Deny) { $denyMatchedCount++ }
        $name = [string](& $ResolveSid $rule.IdentityReference.Value)
        if ($name -notlike 'NT AUTHORITY\SELF' -and $name -notmatch '^S-1-5-21-') { $users += $name }
    }
    $unique = @($users | Select-Object -Unique)
    [pscustomobject]@{
        Users                   = $unique
        TotalAceCount           = @($all).Count
        ExplicitAceCount        = $explicitCount
        MatchedAceCount         = $matchedCount
        DenyMatchedAceCount     = $denyMatchedCount
        AllExtendedRightsAceCount = $allExtendedRightsCount
    }
}

function Invoke-SendAsSelfTest {
    $sendAs = 'ab721a54-1e2f-11d0-9819-00aa0040529b'
    $receiveAs = 'ab721a56-1e2f-11d0-9819-00aa0040529b'
    $rights = @{ $sendAs = 'Send-As' }
    $names = @{
        'S-1-5-10'                   = 'NT AUTHORITY\SELF'
        'S-1-5-21-1-2-3-1001'        = 'CONTOSO\alice'
        'S-1-5-21-1-2-3-1002'        = 'CONTOSO\bob'
        'S-1-5-21-1-2-3-1003'        = 'CONTOSO\carol'
        'S-1-5-21-1-2-3-1004'        = 'CONTOSO\dave'
        'S-1-5-21-1-2-3-1005'        = 'CONTOSO\erin'
    }
    $resolver = { param($sid) if ($names.ContainsKey($sid)) { $names[$sid] } else { $sid } }.GetNewClosure()
    $sddl = 'D:' +
        "(OA;;CR;$sendAs;;S-1-5-10)" +                 # SELF: excluded
        "(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)" +      # alice: kept
        "(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)" +      # alice duplicate: unique
        "(OA;ID;CR;$sendAs;;S-1-5-21-1-2-3-1002)" +    # bob inherited: excluded
        "(OD;;CR;$sendAs;;S-1-5-21-1-2-3-1003)" +      # carol deny: kept (old semantics)
        "(OA;;CR;$receiveAs;;S-1-5-21-1-2-3-1004)" +   # dave Receive-As: excluded
        "(OA;;RP;$sendAs;;S-1-5-21-1-2-3-1005)" +      # erin read property: excluded
        "(OA;;CR;$sendAs;;S-1-5-21-9-9-9-9999)" +      # orphan SID: excluded
        '(A;;CR;;;S-1-5-21-1-2-3-1005)' +              # all extended rights: counted, excluded
        '(A;;GA;;;BA)'
    $sd = New-Object System.DirectoryServices.ActiveDirectorySecurity
    $sd.SetSecurityDescriptorSddlForm($sddl)
    $r = Get-SmartM365SendAsFromSecurityDescriptor -SecurityDescriptor $sd -SendRightGuids $rights -ResolveSid $resolver
    $failures = @()
    $actual = ($r.Users -join ';')
    if ($actual -ne 'CONTOSO\alice;CONTOSO\carol') { $failures += "Users='$actual'" }
    if ($r.MatchedAceCount -ne 5) { $failures += "MatchedAceCount=$($r.MatchedAceCount)" }
    if ($r.DenyMatchedAceCount -ne 1) { $failures += "DenyMatchedAceCount=$($r.DenyMatchedAceCount)" }
    if ($r.AllExtendedRightsAceCount -ne 1) { $failures += "AllExtendedRightsAceCount=$($r.AllExtendedRightsAceCount)" }
    if ($r.TotalAceCount -ne 10) { $failures += "TotalAceCount=$($r.TotalAceCount)" }
    if ($r.ExplicitAceCount -ne 9) { $failures += "ExplicitAceCount=$($r.ExplicitAceCount)" }
    $empty = New-Object System.DirectoryServices.ActiveDirectorySecurity
    $empty.SetSecurityDescriptorSddlForm('D:(A;;GA;;;BA)')
    $e = Get-SmartM365SendAsFromSecurityDescriptor -SecurityDescriptor $empty -SendRightGuids $rights -ResolveSid $resolver
    if (@($e.Users).Count -ne 0 -or ($e.Users -join ';') -ne '') { $failures += 'Empty DACL returned users' }
    if ($failures.Count -gt 0) { throw ('Self-test failed: ' + ($failures -join ', ')) }
    Write-Host 'Self-test passed (10 ACE cases + empty DACL).'
}

function Get-DomainFqdnFromDn {
    param([string]$DistinguishedName)
    $parts = @($DistinguishedName -split ',' | Where-Object { $_ -like 'DC=*' } | ForEach-Object { $_.Substring(3) })
    return ($parts -join '.')
}

Invoke-SendAsSelfTest
if ($SelfTestOnly) { return }

if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Run this qualification in Windows PowerShell 5.1, like the orchestrator job.' }

if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $env:ProgramData ('SmartM365\Qualification\SendAsAcl\' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$script:LogFile = Join-Path $OutputDirectory 'SendAsAclEquivalence.log'
$detailsPath = Join-Path $OutputDirectory 'SendAsAclEquivalence-Details.csv'
$summaryPath = Join-Path $OutputDirectory 'SendAsAclEquivalence-Summary.csv'
$rightsPath = Join-Path $OutputDirectory 'SendAsAclEquivalence-RightsMap.csv'

Write-QualificationLog "SendAs ACL equivalence qualification v$($script:QualificationVersion) started on $env:COMPUTERNAME as $env:USERDOMAIN\$env:USERNAME. Read-only."
Write-QualificationLog "Output: $OutputDirectory"

Import-Module ActiveDirectory -ErrorAction Stop
$sessionHelper = Join-Path $PSScriptRoot '..\..\..\..\..\Modules\SmartM365.Core\SmartM365-ExchangeOnPremSession.ps1'
. (Resolve-Path $sessionHelper).ProviderPath

$configNc = (Get-ADRootDSE).configurationNamingContext
$rightsRows = @(Get-ADObject -SearchBase ("CN=Extended-Rights,$configNc") -LDAPFilter '(objectClass=controlAccessRight)' -Properties rightsGuid, displayName)
$sendRightGuids = @{}
foreach ($row in $rightsRows) {
    $rg = [guid]::Empty
    if ([guid]::TryParse([string]$row.rightsGuid, [ref]$rg) -and ([string]$row.Name -like '*Send*')) {
        $sendRightGuids[$rg.ToString('D')] = [string]$row.Name
    }
}
$rightsRows | Where-Object { [string]$_.Name -like '*Send*' -or [string]$_.displayName -like '*Send*' } |
    Select-Object Name, displayName, rightsGuid |
    Export-Csv -Path $rightsPath -NoTypeInformation -Encoding UTF8
Write-QualificationLog ("Extended rights matched by name '*Send*': {0} ({1})" -f $sendRightGuids.Count, (($sendRightGuids.Values | Sort-Object) -join ', '))
if ($sendRightGuids.Count -eq 0) { throw 'No extended right named like *Send* found; stopping.' }

$sidCache = @{}
$resolveSid = {
    param($sid)
    if ($sidCache.ContainsKey($sid)) { return $sidCache[$sid] }
    $name = $sid
    try { $name = ([System.Security.Principal.SecurityIdentifier]$sid).Translate([System.Security.Principal.NTAccount]).Value } catch { $name = $sid }
    $sidCache[$sid] = $name
    return $name
}

$context = $null
$details = New-Object 'System.Collections.Generic.List[object]'
try {
    $context = Start-SmartM365ExchangeOnPremSession -RequiredCommands @('Get-Mailbox', 'Get-ADPermission', 'Set-ADServerSettings')
    Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop | Out-Null
    Write-QualificationLog "Exchange remote session ready: $($context.Endpoint)"

    $mailboxes = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue)
    Write-QualificationLog "Native mailboxes acquired: $($mailboxes.Count)"
    $selected = New-Object 'System.Collections.Generic.List[object]'
    $groups = $mailboxes | Group-Object { Get-DomainFqdnFromDn ([string]$_.DistinguishedName) } | Sort-Object Name
    foreach ($group in $groups) {
        if ($DomainName.Count -gt 0 -and $DomainName -notcontains $group.Name) { continue }
        $take = @($group.Group | Sort-Object { [string]$_.DistinguishedName })
        if ($MaxMailboxesPerDomain -gt 0 -and $take.Count -gt $MaxMailboxesPerDomain) { $take = @($take | Select-Object -First $MaxMailboxesPerDomain) }
        foreach ($m in $take) { $selected.Add($m) }
        Write-QualificationLog ("Domain {0}: {1} mailboxes selected of {2}" -f $group.Name, $take.Count, $group.Count)
    }

    $total = $selected.Count
    $i = 0
    $lastProgress = Get-Date
    foreach ($Mbx in $selected) {
        $i++
        $dn = [string]$Mbx.DistinguishedName
        $domain = Get-DomainFqdnFromDn $dn

        # Old method: literal copy of the production IncludeADPermission block.
        $oldValue = ''; $oldCount = ''; $oldError = ''; $oldAceCount = ''; $oldUnresolved = ''
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            $SendAsUsersList = @()
            $SendAsUserCount = 0
            $rawPermissions = @(Get-ADPermission -Identity $Mbx.DistinguishedName)
            $sendAsPermissions = $rawPermissions | Where-Object {
                ($_.ExtendedRights -like "*Send*") -and
                ($_.IsInherited -eq $false) -and
                ($_.User -notlike "NT AUTHORITY\SELF") -and
                ($_.User -notmatch '^S-1-5-21-')
            }
            if ($sendAsPermissions) {
                foreach ($perm in $sendAsPermissions) { $SendAsUsersList += $perm.User }
                $SendAsUsersList = $SendAsUsersList | Select-Object -Unique
                $SendAsUserCount = $SendAsUsersList.Count
            }
            $oldValue = ($SendAsUsersList -join ";")
            $oldCount = [string]$SendAsUserCount
            $oldAceCount = [string]$rawPermissions.Count
            $oldUnresolved = [string]@($rawPermissions | Where-Object { [string]$_.User -match '^S-1-5-21-' }).Count
        }
        catch { $oldError = $_.Exception.Message; $oldValue = 'Error retrieving SendAs'; $oldCount = 'Error' }
        $sw.Stop(); $oldMs = $sw.ElapsedMilliseconds

        # New method: direct DACL read.
        $newValue = ''; $newCount = ''; $newError = ''; $acl = $null
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            $adObject = Get-ADObject -Identity $dn -Server $domain -Properties nTSecurityDescriptor -ErrorAction Stop
            $acl = Get-SmartM365SendAsFromSecurityDescriptor -SecurityDescriptor $adObject.nTSecurityDescriptor -SendRightGuids $sendRightGuids -ResolveSid $resolveSid
            $newValue = ($acl.Users -join ';')
            $newCount = [string]@($acl.Users).Count
        }
        catch { $newError = $_.Exception.Message; $newValue = 'Error retrieving SendAs'; $newCount = 'Error' }
        $sw.Stop(); $newMs = $sw.ElapsedMilliseconds

        $oldSet = @($oldValue -split ';' | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique) -join ';'
        $newSet = @($newValue -split ';' | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique) -join ';'
        $details.Add([pscustomobject]@{
            DomainName = $domain; Name = [string]$Mbx.Name; DistinguishedName = $dn; ObjectGuid = [string]$Mbx.Guid
            HasDatabase = [bool]([string]$Mbx.Database)
            OldSendAs = $oldValue; OldSendAsCount = $oldCount; NewSendAs = $newValue; NewSendAsCount = $newCount
            ExactEqual = ($oldValue -ceq $newValue -and $oldCount -eq $newCount -and -not $oldError -and -not $newError)
            SetEqualIgnoreCase = ($oldSet -eq $newSet -and -not $oldError -and -not $newError)
            OldMs = $oldMs; NewMs = $newMs; OldError = $oldError; NewError = $newError
            OldAceCount = $oldAceCount; OldUnresolvedSidAceCount = $oldUnresolved
            NewTotalAceCount = $(if ($acl) { $acl.TotalAceCount } else { '' })
            NewExplicitAceCount = $(if ($acl) { $acl.ExplicitAceCount } else { '' })
            NewMatchedAceCount = $(if ($acl) { $acl.MatchedAceCount } else { '' })
            NewDenyMatchedAceCount = $(if ($acl) { $acl.DenyMatchedAceCount } else { '' })
            NewAllExtendedRightsAceCount = $(if ($acl) { $acl.AllExtendedRightsAceCount } else { '' })
        })

        if (($i % 50) -eq 0 -or ((Get-Date) - $lastProgress).TotalMinutes -ge 5 -or $i -eq $total) {
            $diffs = @($details | Where-Object { -not $_.ExactEqual }).Count
            Write-QualificationLog ("Progress {0}/{1} ({2}); differences so far={3}" -f $i, $total, $domain, $diffs)
            $lastProgress = Get-Date
        }
    }
}
finally {
    if ($details.Count -gt 0) { $details | Export-Csv -Path $detailsPath -NoTypeInformation -Encoding UTF8 }
    if ($null -ne $context) { Stop-SmartM365ExchangeOnPremSession -State $context }
}

$summary = foreach ($g in ($details | Group-Object DomainName | Sort-Object Name)) {
    $rows = @($g.Group)
    [pscustomobject]@{
        DomainName = $g.Name; Mailboxes = $rows.Count
        ExactEqual = @($rows | Where-Object { $_.ExactEqual }).Count
        SetEqualIgnoreCase = @($rows | Where-Object { $_.SetEqualIgnoreCase }).Count
        OldErrors = @($rows | Where-Object { $_.OldError }).Count
        NewErrors = @($rows | Where-Object { $_.NewError }).Count
        OldSeconds = [math]::Round((($rows | Measure-Object OldMs -Sum).Sum) / 1000, 1)
        NewSeconds = [math]::Round((($rows | Measure-Object NewMs -Sum).Sum) / 1000, 1)
        OldAvgAce = [math]::Round((($rows | Where-Object { $_.OldAceCount -ne '' } | ForEach-Object { [int]$_.OldAceCount } | Measure-Object -Average).Average), 1)
        OldAvgUnresolvedSidAce = [math]::Round((($rows | Where-Object { $_.OldUnresolvedSidAceCount -ne '' } | ForEach-Object { [int]$_.OldUnresolvedSidAceCount } | Measure-Object -Average).Average), 1)
        DenyMatchedMailboxes = @($rows | Where-Object { $_.NewDenyMatchedAceCount -ne '' -and [int]$_.NewDenyMatchedAceCount -gt 0 }).Count
    }
}
$summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
$summary | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

$all = $details.Count
$exact = @($details | Where-Object { $_.ExactEqual }).Count
$verdict = $(if ($all -gt 0 -and $exact -eq $all) { 'EQUIVALENT' } else { 'DIFFERENCES' })
Write-QualificationLog ("Verdict: {0}; exact={1}/{2}; old={3:N0}s; new={4:N0}s; details={5}" -f $verdict, $exact, $all, ((($details | Measure-Object OldMs -Sum).Sum) / 1000), ((($details | Measure-Object NewMs -Sum).Sum) / 1000), $detailsPath)
