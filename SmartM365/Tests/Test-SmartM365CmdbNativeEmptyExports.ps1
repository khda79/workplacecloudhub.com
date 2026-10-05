<#
.SYNOPSIS
Offline native CSV projection, empty-export and fail-preserving regression checks.
.DESCRIPTION
Loads selected AST functions only. Acquisition and publication are mocked.
All actual writes use one synthetic temporary directory; no collector is run.
.VERSION
1.0.6
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$sourceRoot = Split-Path $PSScriptRoot -Parent
$results = [System.Collections.Generic.List[object]]::new()
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('cmdb-native-empty-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
function Assert-True { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message } }
function Assert-Throws {
    param([scriptblock]$Body,[string]$Pattern)
    $caught = $null
    try { & $Body | Out-Null } catch { $caught = $_ }
    if ($null -eq $caught -or $caught.Exception.Message -notlike $Pattern) { throw "Expected failure: $Pattern" }
}
function Test-Case {
    param([string]$Name,[scriptblock]$Body)
    try { & $Body; $results.Add([pscustomobject]@{ Name=$Name; Passed=$true; Error='' }) }
    catch { $results.Add([pscustomobject]@{ Name=$Name; Passed=$false; Error=($_.Exception.Message + "`n" + $_.ScriptStackTrace) }) }
}
function Read-TestAst {
    param([string]$RelativePath)
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $sourceRoot $RelativePath),[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
    $ast
}
function Import-TestFunction {
    param($Ast,[string]$Name)
    $node=$Ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name },$true)
    if ($null -eq $node) { throw "Missing function: $Name" }
    Set-Item -Path "Function:script:$Name" -Value ([scriptblock]::Create($node.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
}
function Get-TestSelectColumns {
    param($Command)
    foreach ($element in $Command.CommandElements | Select-Object -Skip 1) {
        $items=if ($element -is [Management.Automation.Language.ArrayLiteralAst]) { @($element.Elements) } else { @($element) }
        foreach ($item in $items) {
            if ($item -is [Management.Automation.Language.StringConstantExpressionAst]) { $item.Value }
            elseif ($item -is [Management.Automation.Language.HashtableAst]) {
                foreach ($pair in $item.KeyValuePairs) {
                    if ($pair.Item1.Value -eq 'Name') { & ([scriptblock]::Create($pair.Item2.Extent.Text)) }
                }
            }
        }
    }
}
function WriteLog { param([string]$Message,[string]$Level) }
function Get-SmartM365CoreContextValue { param([string]$Name,$DefaultValue) $DefaultValue }
function Register-SmartM365GeneratedCsv { param([string]$Path) }
function Add-SmartM365AdGeneratedCsvPath { param([string]$Path) }
function Invoke-SmartM365AdCsvReadWithRetry { param([string]$Path,[scriptblock]$ReadAction) & $ReadAction }
function Remove-SmartM365AdFileWithRetry { param([string]$Path) if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force } }
function Publish-SmartM365ExchangeLocalMailboxCsv { param($SourcePath,$LatestFileName,$HistoryLabel) [pscustomobject]@{ Path=$SourcePath } }
function Get-RemoteMailbox { [CmdletBinding()]param($OnPremisesOrganizationalUnit,$ResultSize,[switch]$ReadFromDomainController) if ($ReadFromDomainController) { throw 'Regressing domain-controller switch was used.' }; if ($script:failQuery) { throw 'Synthetic query failure.' } }
function ConvertFrom-SmartM365ExchangeRemoteMailboxWarnings { param($Warnings) $script:mockWarnings }
function Add-SmartM365LocalMailboxIssue { param($Category,$Operation,$MailboxIdentity,$Message,$SuggestedAction,$ObjectGuid,$NativeRecordRetained) }
$global:SmartM365TenantKey='synthetic-test'; $global:SmartM365OrganizationKey='synthetic'
$global:SmartM365EnvironmentKey='test'; $global:SmartM365TenantId='00000000-0000-0000-0000-000000000001'
$global:SmartM365RequireCsvValidationRules=$true
$adAst=Read-TestAst 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1'
foreach ($name in @('Get-SmartM365AdNativeColumns','Get-SmartM365AdCsvColumns','Complete-SmartM365AdDomainCsvSchema','Combine-CsvFiles')) { Import-TestFunction $adAst $name }
$exchangeAst=Read-TestAst 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
foreach ($name in @('Get-SmartM365RemoteMailboxColumns','Get-SmartM365LocalMailboxColumns','Export-SmartM365EmptyLocalMailboxPopulation','Export-CsvAtomic','Resolve-SmartM365MailboxWarningNativeGuid','Assert-SmartM365MailboxNativePopulation','Find-SmartM365MailboxSmtpConflict','Invoke-SmartM365RemoteMailboxPopulationQuery','Invoke-SmartM365ExchangeRemoteMailboxInventory')) { Import-TestFunction $exchangeAst $name }
$autopilotAst=Read-TestAst 'SmartInventory/M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1'
Import-TestFunction $autopilotAst 'Get-InventoryColumns'
$TargetDomains=@(); $OnlyADPermission=$false; $IncludeRemoteMailboxDelegation=$false
$script:mockWarnings=@(); $script:failQuery=$false
$script:RemoteMailboxForestPopulation=$null
function Reset-LocalEmptyProof {
    $script:LocalMailboxAcquisitionMode='Forest'
    $script:LocalMailboxForestPopulation=@{ 'DC=synthetic,DC=invalid'=@() }
    $script:LocalMailboxAcquisitions=New-Object 'Collections.Generic.List[object]'
    $script:LocalMailboxAcquisitions.Add([pscustomobject]@{
        Scope=''; ResultSize='Unlimited'; ReadFromDomainController=$false; AcquiredThisRun=$true
        QueryCompleted=$true; ProjectionCompleted=$true; Rows=2; ObservedRows=2
    })
    $script:LocalMailboxPopulationCoverageComplete=$true
    $script:LocalMailboxIssues=New-Object 'Collections.Generic.List[object]'
}
try {
    Test-Case 'Autopilot empty schema matches its actual native projection' {
        $table=$autopilotAst.Find({ param($n) $n -is [Management.Automation.Language.HashtableAst] -and $n.KeyValuePairs.Count -gt 0 -and $n.KeyValuePairs[0].Item1.Value -eq 'Autopilot ID' },$true)
        $actual=@($table.KeyValuePairs | ForEach-Object { $_.Item1.Value })
        Assert-True ((@(Get-InventoryColumns) -join '|') -ceq ($actual -join '|')) 'Autopilot empty schema drifted.'
    }
    Test-Case 'Entra user empty schema matches its actual native projection' {
        $ast=Read-TestAst 'SmartInventory/M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1'
        $assignment=$ast.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$obj' },$true)
        $table=$assignment.Find({ param($n) $n -is [Management.Automation.Language.HashtableAst] },$true)
        $actual=@($table.KeyValuePairs | ForEach-Object { $_.Item1.Value })
        $command=$ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Export-SmartM365Csv' },$true)
        $elements=$command.CommandElements
        $columnsNode=$null
        for ($index=0; $index -lt $elements.Count; $index++) {
            if ($elements[$index] -is [Management.Automation.Language.CommandParameterAst] -and $elements[$index].ParameterName -eq 'Columns') { $columnsNode=$elements[$index+1]; break }
        }
        Assert-True (@($columnsNode.FindAll({ param($n) $n -is [Management.Automation.Language.VariableExpressionAst] },$true)).Count -eq 0) 'Empty schema is no longer a literal.'
        $script:usersEmptyColumns=@(& ([scriptblock]::Create($columnsNode.Extent.Text)))
        Assert-True (($script:usersEmptyColumns -join '|') -ceq ($actual -join '|')) 'Entra user empty schema drifted.'
    }
    foreach ($producer in @('SmartInventory/M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1','SmartInventory/M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1')) {
        Test-Case "$(Split-Path $producer -Leaf) retains MAXITEMS filename isolation" {
            $ast=Read-TestAst $producer
            $assignment=$ast.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$runBaseFileName' },$true)
            Assert-True ($assignment.Right.Extent.Text -match '^Add-SmartM365MaxItemsSuffixToBaseName ') 'Limited output can overwrite the canonical name.'
            $command=$ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Export-SmartM365Csv' },$true)
            Assert-True ($command.Extent.Text -match '-BaseFileName \$runBaseFileName\b') 'Publisher does not use the isolated basename.'
        }
    }
    foreach ($kind in @('Users','Groups','DirectoryObjects','OUs','Contacts')) {
        Test-Case "AD $kind empty schema matches the actual nonempty projector" {
            $expected=@(Get-SmartM365AdNativeColumns $kind) -join '|'
            $matched=$false
            foreach ($command in $adAst.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Select-Object' },$true)) {
                if ((@(Get-TestSelectColumns $command) -join '|') -ceq $expected) { $matched=$true; break }
            }
            Assert-True $matched 'No actual nonempty projector matches the empty schema.'
        }
    }
    Test-Case 'AD computer empty schema matches native fixed and configured-group columns' {
        $assignment=$adAst.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$computerRow' },$true)
        $table=$assignment.Find({ param($n) $n -is [Management.Automation.Language.HashtableAst] },$true)
        $actual=@($table.KeyValuePairs | ForEach-Object { $_.Item1.Value })
        $actual+=@(1..10 | ForEach-Object { 'IsMemberOfConfiguredGroup{0:D2}' -f $_ })
        $actual+='MatchedConfiguredGroups'
        Assert-True ((@(Get-SmartM365AdNativeColumns Computers) -join '|') -ceq ($actual -join '|')) 'Native configured-group schema drifted.'
    }
    Test-Case 'AD workstation projection retains and counts missing DNS without widening OS scope' {
        foreach ($name in @('Get-ADStringValue','Get-DomainNameShort','Get-NormalizedDomainAndSam','Convert-GuidToImmutableId')) { Import-TestFunction $adAst $name }
        $query=$adAst.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-ADComputer' -and $n.Extent.Text -match '-Filter \$computerFilter\b' },$true)
        Assert-True ($null -ne $query) 'Native computer acquisition missing.'
        $pipeline=$query.Parent
        Assert-True (($pipeline.PipelineElements.GetCommandName() -join '|') -ceq 'Get-ADComputer|ForEach-Object|Add-SmartM365TenantKey|Export-Csv') 'Unexpected acquisition topology or a row-dropping filter.'
        $filterAssignment=$adAst.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$computerFilter' },$true)
        $computerFilter=& ([scriptblock]::Create($filterAssignment.Right.Extent.Text))
        $currentDomainName='synthetic.example'; $CurrentObjectType='Computers'; $domainSid=''
        $DomainFriendlyNames=[pscustomobject]@{}
        $GroupNameByDNCache=@{}; $GroupParentsByDNCache=@{}; $GroupNameBySIDCache=@{}
        $ResolveNestedComputerGroups=$false; $ConfiguredComputerGroupNames=@(1..10 | ForEach-Object { '' })
        $computerCount=0; $computerWithoutDnsHostNameCount=0
        $dnsValues=@($null,'','   ','test-3.synthetic.example','server.synthetic.example','linux.synthetic.example')
        $operatingSystems=@('','Windows 10','Windows 7','Windows 11','Windows Server 2022','Linux')
        $syntheticComputers=@(for ($index=0; $index -lt $dnsValues.Count; $index++) {
            $record=[ordered]@{}
            foreach ($column in @(Get-SmartM365AdNativeColumns Computers)) { $record[$column]=$null }
            foreach ($column in @('LastLogonTimestamp','pwdLastSet','MemberOf','primaryGroupID','ObjectSID')) { $record[$column]=$null }
            $record.Name='test-'+$index; $record.SamAccountName=$record.Name+'$'
            $record.DNSHostName=$dnsValues[$index]; $record.OperatingSystem=$operatingSystems[$index]
            $record.ObjectGUID=[guid]('00000000-0000-0000-0000-{0:D12}' -f ($index+1))
            $record.ObjectSID=[pscustomobject]@{ Value=('S-1-5-21-1-2-3-{0}' -f (1000+$index)) }
            [pscustomobject]$record
        })
        function Get-ADComputer {
            [CmdletBinding()]param([scriptblock]$Filter,[string]$Server,[string[]]$Properties)
            $expected='(OperatingSystem -like "*Windows*" -and OperatingSystem -notlike "*Server*") -or (OperatingSystem -notlike "*")'
            Assert-True (($Filter.ToString().Trim() -replace '\s+',' ') -ceq $expected) 'Workstation OS scope changed.'
            # Translate AD attribute syntax and its absent-attribute LDAP presence test.
            $mockText=$Filter.ToString().Replace('OperatingSystem','$_.OperatingSystem')
            $mockText=$mockText.Replace('($_.OperatingSystem -notlike "*")','([string]::IsNullOrEmpty($_.OperatingSystem))')
            $mockFilter=[scriptblock]::Create($mockText)
            $syntheticComputers | Where-Object -FilterScript $mockFilter
        }
        function Get-ComputerGroupNames {
            param($Computer,$Server,$DomainSid,$ResolveNestedGroups,$GroupNameByDNCache,$GroupParentsByDNCache,$GroupNameBySIDCache)
            @()
        }
        # Execute the actual query/projector only; do not run a collector or publish tenant data.
        $projection=(@($pipeline.PipelineElements | Select-Object -First 2 | ForEach-Object { $_.Extent.Text }) -join ' | ')
        $observed=& ([scriptblock]::Create('$rows=@('+ $projection +'); [pscustomobject]@{ Rows=@($rows); Total=$computerCount; WithoutDns=$computerWithoutDnsHostNameCount }'))
        Assert-True ($observed.Rows.Count -eq 4 -and $observed.Total -eq 4 -and $observed.WithoutDns -eq 3) ("Missing DNS rows were dropped or counted incorrectly: Rows={0}; Total={1}; WithoutDns={2}." -f $observed.Rows.Count,$observed.Total,$observed.WithoutDns)
        for ($index=0; $index -lt 4; $index++) {
            $row=$observed.Rows[$index]; $native=$syntheticComputers[$index]
            Assert-True ($row.Name -ceq $native.Name -and $row.ObjectGUID -eq $native.ObjectGUID -and $row.SID -ceq $native.ObjectSID.Value) 'Native identity changed.'
            Assert-True ($row.DNSHostName -ceq $native.DNSHostName) 'DNS evidence was fabricated or normalized.'
            Assert-True ($row.ImmutableId_AD -ceq [Convert]::ToBase64String($native.ObjectGUID.ToByteArray())) 'GUID-based identity was lost.'
            Assert-True ((@($row.PSObject.Properties.Name) -join '|') -ceq (@(Get-SmartM365AdNativeColumns Computers) -join '|')) 'Native schema changed.'
        }
    }
    Test-Case 'Remote empty schema matches the actual native record projector' {
        $converter=$exchangeAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertTo-SmartM365ExchangeRemoteMailboxRecord' },$true)
        $assignment=$converter.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$record' },$true)
        $table=$assignment.Find({ param($n) $n -is [Management.Automation.Language.HashtableAst] },$true)
        $actual=@($table.KeyValuePairs | ForEach-Object { $_.Item1.Value })
        $actual+=@('ExchangeGuid','ImmutableId','ObjectGuid','CollectedAtUtc','NativeIdentityStatus')
        Assert-True ((@(Get-SmartM365RemoteMailboxColumns) -join '|') -ceq ($actual -join '|')) 'Remote native schema drifted.'
    }
    foreach ($moduleRelative in @('Modules/SmartM365.Core/SmartM365.Core.psm1','Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
        $coreAst=Read-TestAst $moduleRelative
        foreach ($name in @('Get-SmartM365CsvValidationBaseName','Get-SmartM365CsvValidationRule','Assert-SmartM365CsvDataCompleteness','Add-SmartM365CsvValidationRule','Initialize-SmartM365DefaultCsvValidationRules','Add-SmartM365TenantKeyToCsvData','Write-SmartM365CsvAtomically')) { Import-TestFunction $coreAst $name }
        if ($moduleRelative -notlike '*WindowsPowerShell5*') { Import-TestFunction $coreAst 'Write-SmartM365PreparedCsvAtomically' }
        $global:SmartM365CsvValidationRules=@{}
        Initialize-SmartM365DefaultCsvValidationRules
        $engineName=if ($moduleRelative -like '*WindowsPowerShell5*') { 'PS5 module' } else { 'Core module' }
        $localPath=Join-Path $testRoot ($engineName.Replace(' ','-')+'/Exchange_OnPrem_Mailboxes_AllDomains.csv')
        $domainPath=Join-Path (Split-Path $localPath -Parent) 'Exchange_OnPrem_Mailboxes_synthetic.invalid.csv'
        foreach ($rowCount in @(0,1,2)) {
            Test-Case "$engineName actual domain export retains $rowCount native rows" {
                # Match the collector's non-strict runtime, including PS5 singleton behavior.
                & {
                    Set-StrictMode -Off
                    Reset-LocalEmptyProof
                    $script:LocalMailboxPopulationCoverageComplete=$false
                    $domainName='synthetic.invalid'; $distinguishedName='DC=synthetic,DC=invalid'
                    $pathsForMailboxProcessing=@($distinguishedName)
                    $perDomainCsvFullPath=Join-Path (Split-Path $localPath -Parent) ("Exchange_OnPrem_Mailboxes_cardinality{0}.csv" -f $rowCount)
                    Add-SmartM365CsvValidationRule -Rules $global:SmartM365CsvValidationRules -BaseFileName $perDomainCsvFullPath -CriticalFields @('ObjectGUID') -RequiredColumns (Get-SmartM365LocalMailboxColumns)
                    $script:domainFixtureRows=@(for ($index=1; $index -le $rowCount; $index++) {
                        $values=[ordered]@{}
                        foreach ($column in @(Get-SmartM365LocalMailboxColumns)) { $values[$column]='' }
                        $values.ObjectGUID=([guid]("00000000-0000-0000-0000-{0:D12}" -f $index)).ToString('D')
                        [pscustomobject]$values
                    })
                    $script:LocalMailboxForestPopulation[$distinguishedName]=@($script:domainFixtureRows)
                    function MailboxesProcessing { param($IncludedLDAPPaths) $script:domainFixtureRows }
                    $domainFunction=$exchangeAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Process-SpecificDomain' },$true)
                    $acquire=$domainFunction.Body.ProcessBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -eq '$domainDataFromProcessing' -and $_.Right.Extent.Text -like '*MailboxesProcessing -*' }
                    $publish=$domainFunction.Body.ProcessBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.IfStatementAst] -and $_.Extent.Text.Contains('Export-CsvAtomic -InputObject $domainDataFromProcessing') }
                    Assert-True (@($acquire).Count -eq 1 -and @($publish).Count -eq 1) 'Actual domain acquisition/export branch is ambiguous.'
                    & ([scriptblock]::Create(($acquire.Extent.Text,$publish.Extent.Text -join "`n")))
                    $written=@(Import-Csv -LiteralPath $perDomainCsvFullPath)
                    Assert-True ($written.Count -eq $rowCount) 'Domain export lost or fabricated native rows.'
                    for ($index=0; $index -lt $rowCount; $index++) {
                        Assert-True ($written[$index].ObjectGUID -eq $script:domainFixtureRows[$index].ObjectGUID) 'Domain export changed native identity.'
                    }
                    Assert-True (-not $script:LocalMailboxPopulationCoverageComplete) 'Domain export invented global coverage.'
                }
            }
        }
        Test-Case "$engineName empty global forest preserves old rows and shared validation" {
            New-Item -ItemType Directory -Path (Split-Path $localPath -Parent) -Force | Out-Null
            [pscustomobject]@{ ObjectGUID='00000000-0000-0000-0000-000000000001'; DomainName='synthetic.invalid' } | Export-Csv -LiteralPath $localPath -NoTypeInformation
            Reset-LocalEmptyProof
            $defaultRule=$global:SmartM365CsvValidationRules['Exchange_OnPrem_Mailboxes_AllDomains']
            Assert-True (-not $defaultRule.AllowEmptyDataset) 'Fixture unexpectedly weakened the default rule.'
            $before=(Get-FileHash -LiteralPath $localPath).Hash
            $script:LocalMailboxAcquisitions[0].Rows=0; $script:LocalMailboxAcquisitions[0].ObservedRows=0
            Assert-Throws { Export-SmartM365EmptyLocalMailboxPopulation -Path $localPath } '*Unconfirmed empty local mailbox forest*'
            Assert-True ((Get-FileHash -LiteralPath $localPath).Hash -eq $before) 'Empty global forest overwrote the last CSV.'
            Assert-True (-not $global:SmartM365CsvValidationRules.ContainsKey($localPath) -and [object]::ReferenceEquals($defaultRule,$global:SmartM365CsvValidationRules['Exchange_OnPrem_Mailboxes_AllDomains'])) 'Shared validation changed.'
            Assert-Throws { Write-SmartM365CsvAtomically -Data @() -Path $localPath -Columns (Get-SmartM365LocalMailboxColumns) } '*empty*'
        }
        Test-Case "$engineName empty native domain writes without claiming completed global projection" {
            Reset-LocalEmptyProof
            $script:LocalMailboxAcquisitions[0].Rows=2; $script:LocalMailboxAcquisitions[0].ObservedRows=2
            $script:LocalMailboxAcquisitions[0].ProjectionCompleted=$false; $script:LocalMailboxPopulationCoverageComplete=$false
            $domainPath=Join-Path (Split-Path $localPath -Parent) 'Exchange_OnPrem_Mailboxes_synthetic.invalid.csv'
            $defaultRule=$global:SmartM365CsvValidationRules['Exchange_OnPrem_Mailboxes_AllDomains']
            Export-SmartM365EmptyLocalMailboxPopulation -Path $domainPath -DomainScope 'DC=synthetic,DC=invalid'
            Assert-True (@(Import-Csv -LiteralPath $domainPath).Count -eq 0 -and @(Get-SmartM365AdCsvColumns $domainPath).Count -eq (4+@(Get-SmartM365LocalMailboxColumns).Count)) 'Domain empty export is incomplete.'
            Assert-True (-not $script:LocalMailboxAcquisitions[0].ProjectionCompleted -and -not $script:LocalMailboxPopulationCoverageComplete) 'Domain write fabricated global coverage.'
            Assert-True (-not $global:SmartM365CsvValidationRules.ContainsKey($domainPath) -and [object]::ReferenceEquals($defaultRule,$global:SmartM365CsvValidationRules['Exchange_OnPrem_Mailboxes_AllDomains'])) 'Domain write weakened shared validation.'
        }
        foreach ($defect in @('DomainMode','NoPopulation','NoQuery','FailedQuery','LimitedQuery','ScopedQuery','StaleQuery','RegressingSwitch','ZeroForest','Nonempty','UnknownWarning')) {
            Test-Case "$engineName empty local export preserves last CSV when proof is invalid / $defect" {
                Reset-LocalEmptyProof
                $before=(Get-FileHash -LiteralPath $domainPath).Hash
                switch ($defect) {
                    'DomainMode' { $script:LocalMailboxAcquisitionMode='Domain' }
                    'NoPopulation' { $script:LocalMailboxForestPopulation=$null }
                    'NoQuery' { $script:LocalMailboxAcquisitions.Clear() }
                    'FailedQuery' { $script:LocalMailboxAcquisitions[0].QueryCompleted=$false }
                    'LimitedQuery' { $script:LocalMailboxAcquisitions[0].ResultSize='1' }
                    'ScopedQuery' { $script:LocalMailboxAcquisitions[0].Scope='synthetic' }
                    'StaleQuery' { $script:LocalMailboxAcquisitions[0].AcquiredThisRun=$false }
                    'RegressingSwitch' { $script:LocalMailboxAcquisitions[0].ReadFromDomainController=$true }
                    'ZeroForest' { $script:LocalMailboxAcquisitions[0].ObservedRows=0 }
                    'Nonempty' { $script:LocalMailboxForestPopulation['DC=synthetic,DC=invalid']=@([pscustomobject]@{ Guid=[guid]::NewGuid() }) }
                    'UnknownWarning' { $script:LocalMailboxIssues.Add([pscustomobject]@{ BlocksCmdbQualification=$true }) }
                }
                Assert-Throws { Export-SmartM365EmptyLocalMailboxPopulation -Path $domainPath -DomainScope 'DC=synthetic,DC=invalid' } '*'
                Assert-True ((Get-FileHash -LiteralPath $domainPath).Hash -eq $before -and -not $global:SmartM365CsvValidationRules.ContainsKey($domainPath)) 'Invalid evidence changed the last CSV or validation.'
            }
        }
        Test-Case "$engineName unknown or nonempty domain cannot be published as zero" {
            Reset-LocalEmptyProof
            $before=(Get-FileHash -LiteralPath $localPath).Hash
            Assert-Throws { Export-SmartM365EmptyLocalMailboxPopulation -Path $localPath -DomainScope 'DC=absent,DC=invalid' } '*not proven empty*'
            $script:LocalMailboxForestPopulation['DC=synthetic,DC=invalid']=@([pscustomobject]@{ Guid=[guid]::NewGuid() })
            Assert-Throws { Export-SmartM365EmptyLocalMailboxPopulation -Path $localPath -DomainScope 'DC=synthetic,DC=invalid' } '*not proven empty*'
            Assert-True ((Get-FileHash -LiteralPath $localPath).Hash -eq $before) 'Domain zero proof replaced real rows.'
        }
        Test-Case "$engineName failed empty local write restores an existing exact-path rule" {
            Reset-LocalEmptyProof
            $invalidPath=Join-Path $testRoot ($engineName.Replace(' ','-')+'-locked.csv')
            [pscustomobject]@{ ObjectGUID='00000000-0000-0000-0000-000000000001' } | Export-Csv -LiteralPath $invalidPath -NoTypeInformation
            $previous=@{ AllowEmptyDataset=$false }
            $global:SmartM365CsvValidationRules[$invalidPath]=$previous
            $lock=[IO.File]::Open($invalidPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
            try { Assert-Throws { Export-SmartM365EmptyLocalMailboxPopulation -Path $invalidPath -DomainScope 'DC=synthetic,DC=invalid' } '*' }
            finally { $lock.Dispose() }
            Assert-True ([object]::ReferenceEquals($previous,$global:SmartM365CsvValidationRules[$invalidPath])) 'Write failure leaked an empty-export rule.'
            $global:SmartM365CsvValidationRules.Remove($invalidPath)
        }
        if ($engineName -eq 'Core module') {
            $schemas=@{
                Intune_Autopilot_Devices=@(Get-InventoryColumns)
                M365_Users_Active=$script:usersEmptyColumns
                M365_Entra_VerifiedDomains=@('Id','IsVerified','IsDefault','IsInitial','AuthenticationType','SupportedServices','AvailabilityStatus')
            }
            foreach ($base in $schemas.Keys) {
                Test-Case "$base actual empty write keeps its complete native schema" {
                    $path=Join-Path $testRoot ($base+'.csv')
                    $columns=@($schemas[$base])
                    Write-SmartM365CsvAtomically -Data @() -Path $path -Columns $columns
                    $expected=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId')+$columns
                    Assert-True ((@(Get-SmartM365AdCsvColumns $path) -join '|') -ceq ($expected -join '|')) 'CSV writer dropped native headers.'
                    Assert-True (@(Import-Csv -LiteralPath $path).Count -eq 0) 'CSV writer fabricated data.'
                    $before=(Get-FileHash -LiteralPath $path).Hash
                    Assert-Throws { Write-SmartM365CsvAtomically -Data @() -Path $path -Columns @('TenantKey') } '*missing required column*'
                    Assert-True ((Get-FileHash -LiteralPath $path).Hash -eq $before) 'Invalid schema overwrote a qualified empty CSV.'
                }
            }
        }
        foreach ($kind in @('Users','Computers','Groups','DirectoryObjects','OUs','Contacts')) {
            Test-Case "$engineName AD $kind empty export retains every header" {
                $domainFolder=Join-Path $testRoot ($engineName.Replace(' ','-')+'-'+$kind)
                New-Item -ItemType Directory -Path $domainFolder -Force | Out-Null
                $nativePath=Join-Path $domainFolder ('AD_'+$kind+'_synthetic.csv')
                @() | Export-Csv -LiteralPath $nativePath -NoTypeInformation -Encoding UTF8
                Complete-SmartM365AdDomainCsvSchema -Path $nativePath -Kind $kind
                $headers=@(Get-SmartM365AdCsvColumns $nativePath)
                $expected=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId')+@(Get-SmartM365AdNativeColumns $kind)
                Assert-True (($headers -join '|') -ceq ($expected -join '|')) 'Native empty schema drifted.'
                Assert-True (@(Import-Csv -LiteralPath $nativePath).Count -eq 0) 'Fabricated a data row.'
                $combined=Join-Path $testRoot ('AD_'+$kind+'_AllDomains.csv')
                Combine-CsvFiles -SourceFolder $domainFolder -Filter '*.csv' -DestinationFile $combined
                Assert-True ((@(Get-SmartM365AdCsvColumns $combined) -join '|') -ceq ($expected -join '|')) 'Combined empty schema drifted.'
                Assert-True (@(Import-Csv -LiteralPath $combined).Count -eq 0) 'Combined empty export fabricated rows.'
            }
        }
        foreach ($kind in @('Users','Computers')) {
            Test-Case "$engineName AD $kind enrichment writes all columns for zero rows" {
                $helper=if ($kind -eq 'Users') { 'SmartM365-ActiveDirectory-UsersEnrichment.ps1' } else { 'SmartM365-ActiveDirectory-Enrichment.ps1' }
                $helperAst=Read-TestAst ('SmartInventory/ActiveDirectoryInventory/'+$helper)
                $function=$helperAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq ('Invoke-SmartM365Ad'+$kind+'EnrichedCsv') },$true)
                if ($kind -eq 'Users') {
                    $assignment=$function.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$calculatedColumns' },$true)
                    $calculatedColumns=@(& ([scriptblock]::Create($assignment.Right.Extent.Text)))
                } else {
                    $columnAst=Read-TestAst 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-EnrichedColumns.ps1'
                    $calculatedColumns=@(& ([scriptblock]::Create($columnAst.EndBlock.Extent.Text)))
                    $rename=$function.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$rawCalculatedColumnRenames' },$true)
                    $rawCalculatedColumnRenames=& ([scriptblock]::Create($rename.Right.Extent.Text))
                }
                $native=Join-Path $testRoot ($engineName.Replace(' ','-')+'-'+$kind+'/AD_'+$kind+'_synthetic.csv')
                $CombinedComputersCsv=$native; $CombinedUsersCsv=$native
                $OutputFolder=Join-Path $testRoot ('enriched-'+$kind)
                $enrichedRows=[System.Collections.Generic.List[object]]::new()
                # Execute only the final schema/write statements, not enrichment acquisition or private workbook reads.
                $tail=[System.Collections.Generic.List[string]]::new(); $capture=$false
                foreach ($statement in $function.Body.EndBlock.Statements) {
                    if ($statement -is [Management.Automation.Language.AssignmentStatementAst] -and $statement.Left.Extent.Text -eq '$enrichedCsv') { $capture=$true }
                    if (-not $capture) { continue }
                    $tail.Add($statement.Extent.Text)
                    if ($statement -is [Management.Automation.Language.PipelineAst] -and $statement.PipelineElements[0] -is [Management.Automation.Language.CommandAst] -and $statement.PipelineElements[0].GetCommandName() -eq 'Write-SmartM365CsvAtomically') { break }
                }
                Assert-True ($tail.Count -eq 4) 'Unexpected enrichment write topology.'
                & ([scriptblock]::Create($tail -join "`n"))
                $output=Join-Path $OutputFolder ('AD_'+$kind+'_AllDomains.csv')
                $actual=@(Get-SmartM365AdCsvColumns $output)
                $expected=@(Get-SmartM365AdCsvColumns $native)
                $expected+=@($calculatedColumns | Where-Object { $expected -inotcontains $_ })
                Assert-True (($actual -join '|') -ceq ($expected -join '|')) 'Enriched zero schema differs from its native + calculated schema.'
                Assert-True (@(Import-Csv -LiteralPath $output).Count -eq 0) 'Enrichment fabricated a row.'
            }
        }
        Test-Case "$engineName inconsistent empty domain schemas preserve last output" {
            $folder=Join-Path $testRoot 'inconsistent'
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
            foreach ($name in @('first','second')) {
                $path=Join-Path $folder ($name+'.csv')
                Add-SmartM365CsvValidationRule -Rules $global:SmartM365CsvValidationRules -BaseFileName $name -CriticalFields @('DomainName') -AllowEmptyDataset
                $columns=@('DomainName','ObjectGUID')
                if ($name -eq 'second') { $columns+= 'UnexpectedColumn' }
                Write-SmartM365CsvAtomically -Data @() -Path $path -Columns $columns
            }
            $last=Join-Path $testRoot 'AD_Users_AllDomains.csv'
            $before=(Get-FileHash -LiteralPath $last).Hash
            Assert-Throws { Combine-CsvFiles -SourceFolder $folder -Filter '*.csv' -DestinationFile $last } '*schemas differ*'
            Assert-True ((Get-FileHash -LiteralPath $last).Hash -eq $before) 'Last valid output was changed.'
        }
        # Remote mailbox rules are owned by the Windows PowerShell 5 compatibility module.
        if ($engineName -eq 'PS5 module') {
            $remoteFolder=Join-Path $testRoot 'remote'
            Test-Case 'Remote zero preserves the last valid native CSV' {
                New-Item -ItemType Directory -Path $remoteFolder -Force | Out-Null
                $last=Join-Path $remoteFolder 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv'
                [pscustomobject]@{ ObjectGuid='00000000-0000-0000-0000-000000000001' } | Export-Csv -LiteralPath $last -NoTypeInformation
                $before=(Get-FileHash -LiteralPath $last).Hash
                Assert-Throws { Invoke-SmartM365ExchangeRemoteMailboxInventory -RemoteOutputPath $remoteFolder } '*Unconfirmed empty remote mailbox population*'
                Assert-True ((Get-FileHash -LiteralPath $last).Hash -eq $before) 'Remote zero overwrote the last CSV.'
            }
            Test-Case 'Failed remote query does not overwrite the last valid export' {
                $last=Join-Path $remoteFolder 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv'
                $before=(Get-FileHash -LiteralPath $last).Hash
                $script:failQuery=$true
                try { Assert-Throws { Invoke-SmartM365ExchangeRemoteMailboxInventory -RemoteOutputPath $remoteFolder -IncludedLDAPPaths @('synthetic-scope') } '*Synthetic query failure*' }
                finally { $script:failQuery=$false }
                Assert-True ((Get-FileHash -LiteralPath $last).Hash -eq $before) 'Query failure overwrote the CSV.'
            }
            Test-Case 'Remote acquisition warnings cannot qualify an empty export' {
                $last=Join-Path $remoteFolder 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv'
                $before=(Get-FileHash -LiteralPath $last).Hash
                $script:mockWarnings=@([pscustomobject]@{ Issue='Synthetic'; ObjectPath=''; Warning='Unavailable evidence'; SuggestedAction='Review' })
                try { Assert-Throws { Invoke-SmartM365ExchangeRemoteMailboxInventory -RemoteOutputPath $remoteFolder } '*Unconfirmed empty remote mailbox population*' }
                finally { $script:mockWarnings=@() }
                Assert-True ((Get-FileHash -LiteralPath $last).Hash -eq $before) 'Warnings overwrote the CSV.'
            }
        }
    }
    $results | Format-Table -AutoSize
    if (@($results | Where-Object { -not $_.Passed }).Count) {
        $results | Where-Object { -not $_.Passed } | Format-List
        throw 'Native empty-export regression failed.'
    }
    Write-Host ("PASS: {0} native empty-export checks; PowerShell {1}." -f $results.Count,$PSVersionTable.PSVersion)
}
finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $tempParent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if ($resolved.StartsWith($tempParent,[StringComparison]::OrdinalIgnoreCase) -and ([IO.Path]::GetFileName($resolved) -like 'cmdb-native-empty-*')) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCryc/jWeQjMSZo
# fNJmPkfa3Kg9R6p7lnZ7LlRzudkxH6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJDvbh3NnqwIczEcwcr2H+FCusBgYnX5YV7WzUJtqWneMA0GCSqG
# SIb3DQEBAQUABIIBgEjyV4daBKzwox9vRlxj+PdHPI2EItj23WgP8WrTr+LG+lqx
# 17OFEWoYPEY+cl6vMlIXzlEzjcvLHEZ2nup02NzZqZoSvhvoa6o/cvhH3XuhJzzr
# nO+NBglIUpPneMM+PCBiafMzOX/vzwSrYIkEXvXvpRue04tPm4MAezuhQjxrgdiL
# hQijWlhM3CsJ9N3tubjQmivSqlwAfAgxmZn2kknvwOLpO/4mZcHhN8cNas4kK7bi
# mzJzAi3ibfTLHInOq98acfKHftf+1FeXVECXjFx6RxzcNkHPQMHl34AsecJz+9On
# mPat61YiQS53a3cmToV7+xaL8OUxDfWXdKGzjRabESjXa4348IdKk9hC0Drp+5sR
# fcgVus+NFoPLMLGirBJYFyEfdZuT0i/uWimJZVJGBc6IXdOsfGlRLLQkrp1BSen/
# xdvsfCia7ixV9l0s1ExVgoj49Mn1JSLmx8iU7ohaG3LkvDxEl0OY/1lfvCqMkz4k
# BjxNWPusd+G9YKQv1KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDUxMzAz
# MzdaMC8GCSqGSIb3DQEJBDEiBCCDO01gVud0YSXG+ILhD2ak7EK0WjZhrkY5Yb/V
# 7iLvWzANBgkqhkiG9w0BAQEFAASCAgA9ahKy83YSDZriibxjcct+HUtyDW6HuQbR
# 4oZyTkrst7vXDAaKEBqYKX9Ks76eI8SATmLwpIIBGJCIxNG1bH2UHju+oiJPIF8u
# 7g3HIL4IY0hbcX/qCNI5YaWrtrockT+mQOhseI987vbyhqwsDXo/0D0X0My3HxxT
# 3B7L1YrHc44Vrc1D4uWAnGkmp6FSSNQAI6s6u8D8ZthE9Jtsp57XEX0M2wv/jHNt
# uyBn1vUX7jatJv2XsRHoZllZJb6pmhPhQ1TKPGGqts9o7zWvbNLTiIUCrihOLVJn
# JI4WX5wQqPJhblCIoLIVc+Qp/qjHtSn23NsmATuedIRydn6hppMva1zJYK4cFCp6
# EjprUP4TSx2EXDgtA1UUUd3aFj9H+ng4KLhloJql67RMKv93smKIe4PH19w4fxBT
# RWzdD6/GbUpR/y5eQUeqhgkNE4n4dpfPRmxyw8EJps9JfIaGbDitfRFPWahnya8P
# jkAeYGcE7TBxluXOXCRdX+WdPQHqJlZfoKpgJwQHqHuqOhERjKUjmkzUU/2suhT4
# tcum1K0UQG23BGOJx19hYBcsd72l1xJPNeIy9Zq0seTsEeuPyzqNIe7x8t/u64Wz
# cicxxb0C22cdVW91wL3CV2j8K65d720OYBR8DrgRRvwZ0Mt1Ww+RdK/Xem+30yoM
# Nvr8DPJ3Ww==
# SIG # End signature block
