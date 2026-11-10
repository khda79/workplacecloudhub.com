#Requires -Version 5.1
<# .SYNOPSIS Offline tests for the local mailbox SendAs DACL reader and its Get-ADPermission fallback. .VERSION 1.0.0 #>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$checks = 0
function Check([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message }; $script:checks++ }
function Wire($Value) { [Management.Automation.PSSerializer]::Deserialize([Management.Automation.PSSerializer]::Serialize($Value, 10)) }

$collector = Join-Path $smartRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($collector, [ref]$tokens, [ref]$errors)
Check ($errors.Count -eq 0) 'Collector does not parse'
foreach ($name in @('Get-SmartM365SendAsFromSecurityDescriptor', 'Initialize-SmartM365SendAsAclContext', 'Resolve-SmartM365SendAsSid', 'Get-SmartM365MailboxSendAsUsers')) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    Check ($null -ne $node) ('SendAs dependency missing: ' + $name)
    . ([scriptblock]::Create($node.Extent.Text))
}
$blockNode = $ast.Find({ param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '^If \(\$IncludeADPermission\)' }, $true)
Check ($null -ne $blockNode) 'IncludeADPermission block missing'
Check ($blockNode.Extent.Text -match 'Get-SmartM365MailboxSendAsUsers') 'IncludeADPermission block does not use the DACL reader'

$sendAs = 'ab721a54-1e2f-11d0-9819-00aa0040529b'
$receiveAs = 'ab721a56-1e2f-11d0-9819-00aa0040529b'
$sendTo = 'ab721a55-1e2f-11d0-9819-00aa0040529b'
$names = @{
    'S-1-5-10' = 'NT AUTHORITY\SELF'
    'S-1-5-21-1-2-3-1001' = 'CONTOSO\alice'
    'S-1-5-21-1-2-3-1002' = 'CONTOSO\bob'
    'S-1-5-21-1-2-3-1003' = 'CONTOSO\carol'
    'S-1-5-21-1-2-3-1004' = 'CONTOSO\dave'
    'S-1-5-21-1-2-3-1005' = 'CONTOSO\erin'
}
$resolver = { param($sid) if ($names.ContainsKey($sid)) { $names[$sid] } else { $sid } }.GetNewClosure()
function New-Sd([string]$Sddl) { $sd = New-Object System.DirectoryServices.ActiveDirectorySecurity; $sd.SetSecurityDescriptorSddlForm($Sddl); return $sd }

# 1. Pure DACL filter.
$sddl = 'D:' +
    "(OA;;CR;$sendAs;;S-1-5-10)" +
    "(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)" +
    "(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)" +
    "(OA;ID;CR;$sendAs;;S-1-5-21-1-2-3-1002)" +
    "(OD;;CR;$sendAs;;S-1-5-21-1-2-3-1003)" +
    "(OA;;CR;$receiveAs;;S-1-5-21-1-2-3-1004)" +
    "(OA;;RP;$sendAs;;S-1-5-21-1-2-3-1005)" +
    "(OA;;CR;$sendAs;;S-1-5-21-9-9-9-9999)" +
    "(OA;;CR;$sendTo;;S-1-5-21-1-2-3-1005)" +
    '(A;;CR;;;S-1-5-21-1-2-3-1004)' +
    '(A;;GA;;;BA)'
$rights = @{ $sendAs = 'Send-As'; $sendTo = 'Send-To' }
$r = Get-SmartM365SendAsFromSecurityDescriptor -SecurityDescriptor (New-Sd $sddl) -SendRightGuids $rights -ResolveSid $resolver
Check (($r.Users -join ';') -ceq 'CONTOSO\alice;CONTOSO\carol;CONTOSO\erin') ("Filter users: " + ($r.Users -join ';'))
Check ($r.MatchedAceCount -eq 6 -and $r.DenyMatchedAceCount -eq 1 -and $r.AllExtendedRightsAceCount -eq 1) 'Filter counters'
Check ($r.TotalAceCount -eq 11 -and $r.ExplicitAceCount -eq 10) 'ACE counters'
$e = Get-SmartM365SendAsFromSecurityDescriptor -SecurityDescriptor (New-Sd 'D:(A;;GA;;;BA)') -SendRightGuids $rights -ResolveSid $resolver
Check (@($e.Users).Count -eq 0) 'Empty DACL returned users'

# 2. AD reader with mocked ActiveDirectory cmdlets.
$script:adCalls = New-Object 'System.Collections.Generic.List[string]'
$script:failAd = $false
$script:objectSddl = $sddl
function WriteLog { param([string]$Message, [string]$Level = 'INFO') $script:logLines += ,("[$Level] $Message") }
function Get-ADRootDSE { [CmdletBinding()] param() [pscustomobject]@{ configurationNamingContext = 'CN=Configuration,DC=example,DC=invalid' } }
function Get-ADObject {
    [CmdletBinding()]
    param($Identity, $Server, $SearchBase, $LDAPFilter, $Properties)
    if ($SearchBase) {
        $script:adCalls.Add('rights')
        return @(
            [pscustomobject]@{ Name = 'Send-As'; rightsGuid = $sendAs },
            [pscustomobject]@{ Name = 'Send-To'; rightsGuid = $sendTo },
            [pscustomobject]@{ Name = 'Receive-As'; rightsGuid = $receiveAs }
        )
    }
    $script:adCalls.Add("object|$Server|$Identity")
    if ($script:failAd) { throw 'Synthetic DC unavailable' }
    [pscustomobject]@{ DistinguishedName = $Identity; nTSecurityDescriptor = (New-Sd $script:objectSddl) }
}
$script:logLines = @()
$script:SendAsRightGuids = $null
Initialize-SmartM365SendAsAclContext
Check (@($script:logLines | Where-Object { $_ -match 'SendAs ACL reader ready: 2 extended rights' }).Count -eq 1) 'Reader readiness not logged'
# Synthetic SIDs cannot be translated by LSA: pre-seed the reader cache.
foreach ($k in $names.Keys) { $script:SendAsSidCache[$k] = $names[$k] }
$script:SendAsSidCache['S-1-5-21-9-9-9-9999'] = 'S-1-5-21-9-9-9-9999'
$dn ='CN=Shared One,OU=Users,DC=fr,DC=example,DC=invalid'
$users = @(Get-SmartM365MailboxSendAsUsers -DistinguishedName $dn)
Check (($users -join ';') -ceq 'CONTOSO\alice;CONTOSO\carol;CONTOSO\erin') 'Reader users'
Check ($script:adCalls[0] -eq 'rights' -and $script:adCalls[1] -eq "object|fr.example.invalid|$dn") ('AD calls: ' + ($script:adCalls -join ' / '))
Check ($script:SendAsRightGuids.Count -eq 2) 'Rights map should keep Send-As and Send-To only'
$null = Get-SmartM365MailboxSendAsUsers -DistinguishedName $dn
Check (@($script:adCalls | Where-Object { $_ -eq 'rights' }).Count -eq 1) 'Rights map loaded more than once'
# A cached SID that the reader resolved must be reused.
$script:SendAsSidCache = @{ 'S-1-5-21-1-2-3-1001' = 'CACHED\alice' }
$script:objectSddl = "D:(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)"
Check ((@(Get-SmartM365MailboxSendAsUsers -DistinguishedName $dn) -join ';') -eq 'CACHED\alice') 'SID cache ignored'
$threw = $false
try { $null = Get-SmartM365MailboxSendAsUsers -DistinguishedName 'CN=NoDomain' } catch { $threw = $true }
Check $threw 'DN without domain component must fail'

# 3. Actual IncludeADPermission block: DACL path, fallback path, error path.
function Write-LogMailboxesProcessing { param([string]$Message) $script:logLines += ,$Message }
function Write-Host { param([Parameter(ValueFromRemainingArguments = $true)]$Object) }
$script:adPermissionMode = 'Rows'
function Get-ADPermission {
    [CmdletBinding()]
    param($Identity)
    if ($script:adPermissionMode -eq 'Throw') { throw 'Synthetic Get-ADPermission failure' }
    Wire @(
        [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; ExtendedRights = @('Send-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\alice'; ExtendedRights = @('Send-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\alice'; ExtendedRights = @('Send-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\bob'; ExtendedRights = @('Send-As'); IsInherited = $true },
        [pscustomobject]@{ User = 'S-1-5-21-9-9-9-9999'; ExtendedRights = @('Send-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\dave'; ExtendedRights = @('Receive-As'); IsInherited = $false },
        [pscustomobject]@{ User = 'CONTOSO\carol'; ExtendedRights = @('Send-As'); IsInherited = $false }
    )
}
function Invoke-Block {
    $script:userObj = [pscustomobject]@{ Name = 'Shared One' }
    $userObj = $script:userObj
    $Mbx = Wire ([pscustomobject]@{ DistinguishedName = $dn; DisplayName = 'Shared One' })
    $IncludeADPermission = $true
    . ([scriptblock]::Create($blockNode.Extent.Text))
    return $userObj
}
$script:SendAsSidCache = @{}
foreach ($k in $names.Keys) { $script:SendAsSidCache[$k] = $names[$k] }
$script:SendAsSidCache['S-1-5-21-9-9-9-9999'] = 'S-1-5-21-9-9-9-9999'
$script:SendAsAclReadCount = 0; $script:SendAsAclFallbackCount = 0
$script:failAd = $false

$script:objectSddl = $sddl
$o = Invoke-Block
Check ($o.SendAs -ceq 'CONTOSO\alice;CONTOSO\carol;CONTOSO\erin' -and $o.SendAsCount -eq 3) ("DACL path: $($o.SendAs) / $($o.SendAsCount)")
$script:objectSddl = "D:(OA;;CR;$sendAs;;S-1-5-21-1-2-3-1001)(A;;GA;;;BA)"
$o = Invoke-Block
Check ($o.SendAs -ceq 'CONTOSO\alice' -and $o.SendAsCount -eq 1) 'Single SendAs count'
$script:objectSddl = 'D:(A;;GA;;;BA)'
$o = Invoke-Block
Check ($o.SendAs -ceq '' -and $o.SendAsCount -eq 0) 'Empty SendAs'
Check ($script:SendAsAclReadCount -eq 3 -and $script:SendAsAclFallbackCount -eq 0) 'ACL read counters'

$script:failAd = $true
$script:logLines = @()
$o = Invoke-Block
Check ($o.SendAs -ceq 'CONTOSO\alice;CONTOSO\carol' -and $o.SendAsCount -eq 2) ("Fallback path: $($o.SendAs) / $($o.SendAsCount)")
Check ($script:SendAsAclFallbackCount -eq 1) 'Fallback counter'
Check (@($script:logLines | Where-Object { $_ -match 'Falling back to Get-ADPermission' }).Count -ge 1) 'Fallback not logged'

$script:adPermissionMode = 'Throw'
$o = Invoke-Block
Check ($o.SendAs -eq 'Error retrieving SendAs' -and $o.SendAsCount -eq 'Error') 'Double failure must publish Error'

Write-Output ('PASS: {0} SendAs ACL checks; {1}.' -f $checks, $PSVersionTable.PSVersion)
