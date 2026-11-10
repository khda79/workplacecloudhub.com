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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCovUObeXEQwgqD
# kUvxVDOLQS4Xg7Zkx4F2yySydNvjoqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICQpEGs4FVIM1pIT0JvHXa3gZQEt6WEtcRw5SuL3CoKcMA0GCSqG
# SIb3DQEBAQUABIIBgGHzJWwoi8fvHm04BmBDNvsb1Lv6pBFQMWqZh8t0hrZz8wPH
# 9NEc550mOQKu/bxQTAPw1X/lxD0oQeCU9mqfsj7RjP/Mbl6JQQDowagb+bjWrBh2
# 83sVtsbckYNfO6+32G7HhhPLwgJfUHC2cSALTutot5OP1yJvFzdFdMpLZWgKoZP7
# epNZpp80DAjTiP3fQrtv4a40gYyw0DBryOZN/9gK4ONVKeHjyV3beNqyGp0hWGwG
# k8XA6srthEx+/vFWyiHfk4RDbNnsxQS7YuZKwyfqLJMbFfARw2jbATct5ZHjsOOo
# I9WakKAIH3oAvBFXTZy/QwYumGAZdz2Wv22tBS++xCZ6gF2VxWRGazEccgqFd1Ha
# XjxYRCvBLfP2+jckbtcbSNf3wx/n1mb8oVLA5Tx3CLQChucl93/xpAs3Z06rCsZ8
# jtrfC2BW5ObcXrEllcb7RvcUqT8/SXEWiQ3WLaunXFGHKD3E8XjpLtI6UctOQDoK
# JcC49JaiY0hP2Q8bE6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTAxMjU5
# NTNaMC8GCSqGSIb3DQEJBDEiBCATj3NPjyPWQrVQZX0ddN8/ca8P9KaRwODn6rCT
# 6tehOTANBgkqhkiG9w0BAQEFAASCAgBPKOWXyQDX+SJMVOyFdu0FtzhkF77kpCF6
# zE4b9vfThFoLz9d1jT6AXfkUmruA0iDy4qvsxOuVbzadUEctdxu43xyYyZw+RX3V
# V1iAcWRl2n879HEJXvC+nFDEQzrV3HDST/G7w8qb3rIcWR63Xhms3aHiMLZgQ7Yt
# OuXPpDPX4eN0MoHh8+6C99EQnhTWKZwy8nhGDVGalrV8aawXarxGHzyAfBog6W0+
# h64TXIWs512WXrg8mkgR3W9HNwQ3TzR6Ar0J62YfBXgiQ9Xg3AEY7yWOZlZIaFa6
# C87owthECfzXn8SXFwPFSnFfCjTnvarTZrvHp+JW2fkWcSbzXxajtitQArMahLkm
# RdJ1x3mTP+om5oXCB+VrU16EYJfPTAUI9+40PWBYq3OVw5DF5XkCAIgt+Ex4L7Og
# 11/hEViC+hfrKImE43BGxjXUNI9VjcueumKurHtE6t90gGXQ4yARkLqxFGKfuqoC
# 9kP0gXjdCoYPo37eogv2ojuXmUvo7u0C5D6iIwemPmtS5sX0858iO1+S/Qjkmk/K
# rOXWG+ZWEUuudxOSPCmCbC9eLE9aH9KyLpdA+EtNYwmYYuoPkQ+PAjQhk2GuidIj
# L2NNJ7ExEMQXjASszCPiZTO/tB5UMTE38aC2sVU8ijqa/S0mw5z6a7Z+BWmx7p8+
# sIIXS7mbIA==
# SIG # End signature block
