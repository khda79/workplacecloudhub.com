# Version 1.1.0
<# Synthetic reader -> consolidation -> business rules; no collector or data reads. #>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1')
$checks=0
function Check([bool]$Ok,[string]$Message) { if(-not $Ok){throw $Message}; $script:checks++ }
$now=[datetimeoffset]'2026-10-22T12:00:00Z'
function Evidence([string]$Iso) { Read-IdentityDate ([pscustomobject]@{Date='legacy';DateIso8601=$Iso}) Date -Now $now }
$unknown=Read-IdentityDate ([pscustomobject]@{Date='2026-10-21T00:00:00Z'}) Date -Now $now
Check ($unknown.Status -eq 'Unqualified legacy value' -and $null -eq $unknown.Utc) 'Old inferred Z was trusted'
foreach($raw in @('2026-10-21T00:00:00Z','21/10/2026 00:00:00')) {
    Check ((ConvertTo-IdentityNativeIso $raw) -eq '') 'String/deserialized value was promoted to native proof'
}
$typed=[datetimeoffset]'2026-10-21T02:00:00+02:00'
Check ((Evidence (ConvertTo-IdentityNativeIso $typed)).Utc -eq [datetimeoffset]'2026-10-21T00:00:00Z') 'Offset not preserved'
$unspecified=[datetime]::SpecifyKind($typed.DateTime,[DateTimeKind]::Unspecified)
Check ((Evidence (ConvertTo-IdentityNativeIso $unspecified)).Status -eq 'Timezone not established') 'Unspecified acquired a timezone'
$serialized=$typed | ConvertTo-Json -Compress | ConvertFrom-Json -DateKind String
Check ((ConvertTo-IdentityNativeIso $serialized) -eq '') 'Serialization without Kind proof was trusted'
foreach($case in @(@('bad','Invalid'),@('2026-10-21T00:00:00','Timezone not established'),@('1601-01-01T00:00:00Z','Sentinel'),@('2027-01-01T00:00:00Z','Future'))) {
    Check ((Evidence $case[0]).Status -eq $case[1]) "Unexpected qualification: $($case[1])"
}
$missing=Read-IdentityDate $null Date -Now $now
Check ($missing.Status -eq 'Absent') 'Missing evidence not distinct'
$filetime=([datetime]'2026-10-21T00:00:00Z').ToFileTimeUtc()
$ft=Read-IdentityDate ([pscustomobject]@{Date='not-parseable';Ticks=$filetime}) Date Ticks -Now $now
Check ($ft.Status -eq 'Qualified' -and $ft.Proof -eq 'AD FILETIME') 'FILETIME proof lost'
foreach($days in @(30,45,90,180)) {
    $recent=Evidence ($now.AddDays(-$days).ToString('o'))
    $old=Evidence ($now.AddDays(-$days-1).ToString('o'))
    foreach($bad in @($unknown,$missing,(Evidence 'bad'),(Evidence '2027-01-01T00:00:00Z'),(Evidence '1601-01-01T00:00:00Z'))) {
        $allUnknown=Merge-IdentityActivity @($bad,$bad) $days -Now $now
        Check ($null -eq $allUnknown.Recent -and $null -eq $allUnknown.Old) 'Unknown became inactive/active'
        $mixedOld=Merge-IdentityActivity @($old,$bad) $days -Now $now
        Check ($null -eq $mixedOld.Old -and $mixedOld.Coverage -eq 'Partial') 'Partial old evidence became global inactivity'
        $mixedRecent=Merge-IdentityActivity @($recent,$bad) $days -Now $now
        Check ($mixedRecent.Recent -eq $true -and $mixedRecent.Old -eq $false -and $mixedRecent.Coverage -eq 'Partial') 'Positive recent observation lost'
        Check ((Format-IdentityActivityCategory $mixedOld $days) -eq 'Unknown / unqualified activity') 'Unknown category says never/inactive'
    }
    $completeOld=Merge-IdentityActivity @($old,$old) $days -Now $now
    Check ($completeOld.Old -eq $true -and $completeOld.Coverage -eq 'Complete') 'Complete old evidence not retained'
}
$crossDay=Evidence '2026-10-22T00:30:00+02:00'
Check ($crossDay.Utc.ToString('yyyy-MM-dd') -eq '2026-10-21') 'Offset crossing UTC day lost'

# Inspect actual consumer rules: never execute entry points or their external actions.
$root=Split-Path $PSScriptRoot -Parent
$paths=@('SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-UsersEnrichment.ps1',
    'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Enrichment.ps1',
    'SmartInventory/ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1',
    'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1')
foreach($path in $paths) {
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root $path),[ref]$tokens,[ref]$errors)
    Check ($errors.Count -eq 0) "Parser: $path"
    if($path -like '*UsersEnrichment*') {
        $assignments=$ast.FindAll({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and
            $n.Left.Extent.Text -match "^\`$out\['(?:ResolvedLastLogonStatus|HasNoResolvedLastLogon|ResolvedLastLogonDateTime|IsEnabledWithNoResolvedLastLogon)'\]$"},$true)
        Check ($assignments.Count -eq 4) 'User activity flags not covered'
        foreach($adActivity in @($unknown,$missing,(Evidence 'bad'),(Evidence '2026-10-21T00:00:00Z'))) {
            $out=[ordered]@{}; $isEnabled=$true
            foreach($assignment in $assignments){. ([scriptblock]::Create($assignment.Extent.Text))}
            if($adActivity.Status -in @('Invalid','Unqualified legacy value')) {
                Check ($out.HasNoResolvedLastLogon -eq '' -and $out.IsEnabledWithNoResolvedLastLogon -eq '') 'Actual user flags turn unknown into never-connected'
            }
            Check ($out.ResolvedLastLogonStatus -eq $adActivity.Status) 'Actual user status lost'
        }
        Check ($ast.Extent.Text.Contains('$combinedActivity = @($adActivity, $entraSuccessful)')) 'Attempted sign-in substituted for successful activity'
    }

}
Write-Output "PASS: $checks identity activity checks; synthetic only."
