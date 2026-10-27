# Version 1.1.0
<# Synthetic scope and temporal-selection consumer regressions. No operational entry point. #>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1')
$checks=0
function Check([bool]$Ok,[string]$Message){if(-not $Ok){throw $Message};$script:checks++}
function LoadAst([string]$Relative){$t=$null;$e=$null;$a=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot "../$Relative"),[ref]$t,[ref]$e);Check ($e.Count -eq 0) "Parser: $Relative";return $a}
$users=LoadAst 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-UsersEnrichment.ps1'
$pc=LoadAst 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Enrichment.ps1'
$now=[datetimeoffset]'2026-10-22T12:00:00Z'
$old=[pscustomobject]@{Status='Qualified';Utc=[datetimeoffset]'2025-01-01T00:00:00Z';Proof='Synthetic'}
$missing=Read-IdentityDate $null LastLogonTime
$na=[pscustomobject]@{Status='Not applicable';Utc=$null;Proof='Positive synthetic scope proof'}
Check ((Merge-IdentityActivity @($old,$na) 180 -Now $now).Old -eq $true) 'Excluded source dilutes applicable scope'
Check ($null -eq (Merge-IdentityActivity @($old,$missing) 180 -Now $now).Old) 'Missing expected source excluded'
Check ($null -eq (Merge-IdentityActivity @($na) 180 -Now $now).Old) 'No applicable source implies inactivity'
$guid='11111111-1111-1111-1111-111111111111'
$ad=[pscustomobject]@{ObjectGUID=$guid;LastLogonDateIso8601='2025-01-01T00:00:00Z'}
$remote=[pscustomobject]@{ObjectGUID=$guid;RecipientTypeDetails='RemoteUserMailbox'}
$cloud=[pscustomobject]@{LastLogonTimeIso8601='2025-01-01T00:00:00Z'}
$recent=[pscustomobject]@{LastLogonTimeIso8601='2026-10-21T00:00:00Z'}
foreach($consumer in @($users)) {
    $variable=if($consumer -eq $users){'$mailboxActivitySet'}else{'$activityEvidence'}
    $assignment=$consumer.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq $variable},$true)
    Check ($null -ne $assignment) 'Actual consumer list not found'
    foreach($scenario in @('remote-old','expected-missing','collection-error','conflict','duplicate','recent-partial')) {
        $user=$ad;$u=$ad;$localMailbox=$null;$lm=$null;$exoMailbox=$cloud;$em=$cloud;$exoStat=$null;$st=$null
        $remoteRows=@($remote)
        switch($scenario){
            'expected-missing' {$remoteRows=@()}
            'collection-error' {$remoteRows=@();$localMailbox=[pscustomobject]@{LastLogonTime='Error'};$lm=$localMailbox}
            'conflict' {$localMailbox=$cloud;$lm=$cloud}
            'duplicate' {$remoteRows=@($remote,$remote)}
            'recent-partial' {$remoteRows=@();$exoMailbox=$recent;$em=$recent}
        }
        $remoteScopeByGuid=New-IdentityRemoteScopeIndex $remoteRows
        . ([scriptblock]::Create($assignment.Extent.Text))
        $entries=if($consumer -eq $users){$mailboxActivitySet}else{$activityEvidence}
        $merged=Merge-IdentityActivity $entries 180 -Now $now
        if($scenario -eq 'remote-old') {Check ($merged.Old -eq $true -and $merged.ExcludedCount -eq 1 -and $merged.Exclusions) 'Native remote scope did not remove only local logon'}
        elseif($scenario -eq 'recent-partial'){Check ($merged.Recent -eq $true -and $merged.Coverage -eq 'Partial') 'Positive recent cloud activity lost'}
        else {Check ($null -eq $merged.Old -and $merged.ExcludedCount -eq 0) "Unknown/ambiguous scope excluded: $scenario"}
    }
}
# Execute the actual device-list assignment, with missing native maps preserved as unknown.
$deviceAssignment=$pc.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$deviceActivity' -and $n.Right.Extent.Text.Contains('Get-IdentityDeviceActivitySet')},$true)
$row=$ad;$nativeEntra=$null;$nativeManaged=$null
. ([scriptblock]::Create($deviceAssignment.Extent.Text))
Check ($null -eq (Merge-IdentityActivity $deviceActivity 90 -Now $now).Old) 'Missing device maps became not applicable'
Check (@($deviceActivity | Where-Object Status -eq 'Not applicable').Count -eq 0) 'Absent enrollment is not exclusion proof'

# Load actual accumulation, eligibility, and output projection functions only.
foreach($name in @('LoadCsv','AddLatest','Eligibility','V','SetWu')){
    $node=$pc.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    . ([scriptblock]::Create($node.Extent.Text))
}
function Get-SourceCandidatePaths($Name){ @() }
function WriteLog { param($Message) }
$projectionScratch=Join-Path ([IO.Path]::GetTempPath()) ('identity-wu-projection-'+[guid]::NewGuid().ToString('N')+'.csv')
function Wu($State,$Policy,$Iso='',$Raw='2026-10-01T09:00:00Z') {
    [pscustomobject]@{PolicyId=$Policy;UpgradeEligibility=$State;AggregateState=$State;ExportDateTime=$Raw;ExportDateTimeIso8601=$Iso}
}
$older=Wu 'Capable' 'p1' '' '2026-10-01T09:00:00Z'
$newer=Wu 'Not Capable' 'p1' '' '2026-10-21T09:00:00Z'
$qualifiedOld=Wu 'Capable' 'p1' '2026-10-01T09:00:00Z'
$qualifiedNew=Wu 'Not Capable' 'p1' '2026-10-21T09:00:00Z'
$cases=@(
    @{Rows=@($older,$newer);Ambiguous=$true},
    @{Rows=@($qualifiedOld,$newer);Ambiguous=$true},
    @{Rows=@($qualifiedOld,$qualifiedNew);Ambiguous=$false},
    @{Rows=@($qualifiedOld,(Wu 'Not Capable' 'p1' '2026-10-01T09:00:00Z'));Ambiguous=$true},
    @{Rows=@($older,$older);Ambiguous=$false},
    @{Rows=@($older,(Wu 'Capable' 'p2'));Ambiguous=$true},
    @{Rows=@($older);Ambiguous=$false},
    @{Rows=@($qualifiedOld,$qualifiedNew,$qualifiedNew);Ambiguous=$false},
    @{Rows=@((Wu 'Capable' 'p1' '2099-10-21T09:00:00Z'));Ambiguous=$true}
)
$eligibilityAssignments=$pc.FindAll({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -in @('$upgradeRaw','$upgradeLabel','$w11EligibilityValue')},$true)
$guard=$pc.Find({param($n)$n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text.Contains('$wuAnySelection.Status')},$true)
$sourceAssignment=$pc.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$out[''Windows11UpgradeEligibilitySource'']'},$true)
try {
foreach($case in $cases){
    $representationResults=@()
    foreach($representation in @('PSCustomObject','ProjectedHashtable','OrderedDictionary')) {
    $results=@()
    foreach($reverse in @($false,$true)){
        $map=@{};$rows=@($case.Rows);if($reverse){[array]::Reverse($rows)}
        if($representation -eq 'ProjectedHashtable') {
            $rows | Export-Csv -LiteralPath $projectionScratch -NoTypeInformation
            $rows=@(LoadCsv 'synthetic-only.csv' @('PolicyId','UpgradeEligibility','AggregateState','ExportDateTime','ExportDateTimeIso8601') $projectionScratch)
            Check ($rows[0] -is [hashtable]) 'Real LoadCsv projection not exercised'
        } elseif($representation -eq 'OrderedDictionary') {
            $rows=@($rows | ForEach-Object {
                $record=[ordered]@{};foreach($property in $_.PSObject.Properties){$record[$property.Name]=$property.Value};$record
            })
        }
        foreach($item in $rows){AddLatest $map 'device' $item}
        $wuAnySelection=Resolve-IdentityWindowsUpdateRows $map.device
        $wuAny=$wuAnySelection.Row;$intune=$null;$isWin11=$false
        Check (($null -eq $wuAny) -eq $case.Ambiguous) 'Unexpected WU ambiguity'
        foreach($assignment in $eligibilityAssignments){. ([scriptblock]::Create($assignment.Extent.Text))}
        $migrationPhase='PHASE 1';$migrationReason='Synthetic independent inputs';$migrationAction='Synthetic action'
        . ([scriptblock]::Create($guard.Extent.Text))
        $out=[ordered]@{};. ([scriptblock]::Create($sourceAssignment.Extent.Text))
        SetWu $out $wuAny 'SyntheticPolicy'
        if($case.Ambiguous){Check ($w11EligibilityValue -eq 'Unknown' -and $migrationPhase -eq 'UNDETERMINED' -and $migrationAction -eq '6 - To qualify') 'Ambiguous WU still drives a recommendation'}
        $results+=([ordered]@{Eligibility=$w11EligibilityValue;Source=$out.Windows11UpgradeEligibilitySource;Phase=$migrationPhase;Action=$migrationAction;State=$out.WindowsUpdateSyntheticPolicyAggregateState;Status=$wuAnySelection.Status} | ConvertTo-Json -Compress)
    }
    Check ($results[0] -ceq $results[1]) 'Permutation changed a business decision'
    $representationResults+=$results[0]
    }
    Check (@($representationResults | Sort-Object -Unique).Count -eq 1) 'Row representation changed the decision'
}
} finally { Remove-Item -LiteralPath $projectionScratch -ErrorAction SilentlyContinue }
# The device-wide map must not collapse different configured policy grains.
$wuAnyByDevice=@{};$wuAnchorByDevice=@{};$wu24ByDevice=@{};$wu25ByDevice=@{}
$deviceKey='device';$anchorKey='p1';$policy24Key='p2';$policy25Key='p3'
$policyLoop=$pc.Find({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Variable.Extent.Text -eq '$row' -and $n.Extent.Text.Contains('AddLatest $wuAnyByDevice')},$true)
$policyGuards=$policyLoop.Body.Statements | Where-Object { $_ -is [Management.Automation.Language.IfStatementAst] }
foreach($row in @($older,(Wu 'Not Capable' 'p2'))){
    $policyKey=$row.PolicyId;AddLatest $wuAnyByDevice $deviceKey $row
    foreach($statement in $policyGuards){. ([scriptblock]::Create($statement.Extent.Text))}
}
Check ($null -eq (Resolve-IdentityWindowsUpdateRows $wuAnyByDevice.device).Row) 'Different policies collapsed at device grain'
Check ((Resolve-IdentityWindowsUpdateRows $wuAnchorByDevice.device).Row.UpgradeEligibility -eq 'Capable') 'Anchor policy grain lost'
Check ((Resolve-IdentityWindowsUpdateRows $wu24ByDevice.device).Row.UpgradeEligibility -eq 'Not Capable') '24H2 policy grain lost'
Write-Output "PASS: $checks scope/Windows Update checks; actual consumer lists and AST rules; synthetic only."
