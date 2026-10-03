<#
.SYNOPSIS
Offline native CSV projection, empty-export and fail-preserving regression checks.
.DESCRIPTION
Loads selected AST functions only. Acquisition and publication are mocked.
All actual writes use one synthetic temporary directory; no collector is run.
.VERSION
1.0.2
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
function Get-RemoteMailbox { [CmdletBinding()]param($OnPremisesOrganizationalUnit,$ResultSize) if ($script:failQuery) { throw 'Synthetic query failure.' } }
function ConvertFrom-SmartM365ExchangeRemoteMailboxWarnings { param($Warnings) $script:mockWarnings }
function Add-SmartM365LocalMailboxIssue { param($Category,$Operation,$MailboxIdentity,$Message,$SuggestedAction,$ObjectGuid,$NativeRecordRetained) }
$global:SmartM365TenantKey='synthetic-test'; $global:SmartM365OrganizationKey='synthetic'
$global:SmartM365EnvironmentKey='test'; $global:SmartM365TenantId='00000000-0000-0000-0000-000000000001'
$global:SmartM365RequireCsvValidationRules=$true
$adAst=Read-TestAst 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1'
foreach ($name in @('Get-SmartM365AdNativeColumns','Get-SmartM365AdCsvColumns','Complete-SmartM365AdDomainCsvSchema','Combine-CsvFiles')) { Import-TestFunction $adAst $name }
$exchangeAst=Read-TestAst 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
foreach ($name in @('Get-SmartM365RemoteMailboxColumns','Export-CsvAtomic','Resolve-SmartM365MailboxWarningNativeGuid','Assert-SmartM365MailboxNativePopulation','Find-SmartM365MailboxSmtpConflicts','Invoke-SmartM365ExchangeRemoteMailboxInventory')) { Import-TestFunction $exchangeAst $name }
$autopilotAst=Read-TestAst 'SmartInventory/M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1'
Import-TestFunction $autopilotAst 'Get-InventoryColumns'
$TargetDomains=@(); $OnlyADPermission=$false; $IncludeRemoteMailboxDelegation=$false
$script:mockWarnings=@(); $script:failQuery=$false
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
        $global:SmartM365CsvValidationRules=@{}
        Initialize-SmartM365DefaultCsvValidationRules
        $engineName=if ($moduleRelative -like '*WindowsPowerShell5*') { 'PS5 module' } else { 'Core module' }
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
            Test-Case 'Remote successful zero writes the full native schema' {
                $result=Invoke-SmartM365ExchangeRemoteMailboxInventory -RemoteOutputPath $remoteFolder
                $expected=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId')+@(Get-SmartM365RemoteMailboxColumns)
                Assert-True ($result.RecordCount -eq 0) 'False population.'
                Assert-True ((@(Get-SmartM365AdCsvColumns $result.CombinedCsv) -join '|') -ceq ($expected -join '|')) 'Remote schema missing.'
                Assert-True (@(Import-Csv -LiteralPath $result.CombinedCsv).Count -eq 0) 'Remote data fabricated.'
            }
            Test-Case 'Failed remote query does not overwrite a successful empty export' {
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
                try { Assert-Throws { Invoke-SmartM365ExchangeRemoteMailboxInventory -RemoteOutputPath $remoteFolder } '*warnings prevent qualification*' }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBB2Txo3qc4RE/r
# BWGnRadR2YNW68K934Fuh2R/KZ6BFqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPaAfNbO4qi2Tf/+DpM4dG+nq8mJDAHP2r9c4Ddq9z6bMA0GCSqG
# SIb3DQEBAQUABIIBgAsyNZUuiRyxNTffMD70vEksLHnERxhXwlDFPR6hkcRxm1+4
# NgRyaZIDCVT13aDvYIRZMudIEA8hBSoawLLiCCOONiI3NJbDojVrPZLsVc5V9RAE
# XZZEzQokpBApnWG4swBc+6hzykcff8SibMitM6j/gksTJPOiCi5GDlQx523BmtKV
# BLZ1cNwcUYsHSVxKo5K2YZAyZtMPfBXbjVoyH//lRgYZM/IDjzx5AHo7qtIn9sa0
# z6TI12VxVo2yukWr5dye7uR8N83Kt5Q2mb1pN5wBOH7nlZCzolLfWVblTb0Zm0KP
# Qk6bpBNYye49UUoGIy2Pt4utDDg7fU0uH109HQWc9KNUaIcVnNpkhoxLXeSvXm6u
# E5edXCW47WFSK3mbGWj7qitaKeR1RB0IF2sfhu/w6wDEcUP81+iof1jbcDTINKkW
# 7C1c5MkdY4ctqQjdLjlCAAi6gcVexfcRz8xI0cNDpXhrmdE7xTggXUTiJ0p16Icr
# n8oBV8FkG5F1PkiFD6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxODIy
# NTJaMC8GCSqGSIb3DQEJBDEiBCA3+sDtzTeByYWrFC7RLSuCWfcRY4zV2uWnORIh
# 04GFMDANBgkqhkiG9w0BAQEFAASCAgCEY1FFTuH0NXePgK30qOwc3ehwC4meqt0v
# OWVDoR1gPlYS5wkqvJl97TBTxJDIOazyI7moj9UaabjAmHIv3CboaGRO0+2zh2xf
# wLG0K4sX2T2gWtM0plW045pSxZiG4ET4mh1eyvR8Df8gmqOvDHE+RpI0vfuYTwSW
# CtVFcXP0yREey2wCWGhLNZrl5wVemsKga4OhJrg6F674qFikZzGVnmnDylj+toGk
# j43GzeAuBNwQ9H0nKssjAYOrtDyT/ernLxQMzL/CcXKzshXz/VM5MwbV732K3iJA
# gAkq5V7RESQ3b/cayHjzkEGPIz6qtyN7wm0zfhEmCsViYbjI8qZ6v4DxH4hvdfDn
# +wc8W/IfKZcNpz3l+oSukVNEqOz+0GNvA5jk0d2mH5tNmEjhr4jOGpPDL1FS42AP
# 0DGZonou/1mEBounDEfON8giUz7mG/5VSow0xIfmbzCzRkyn1O+4gtfnJFWAFfrK
# 260tGaqse2Xqu3XZHVzoq60rE0VEjhKYbV2tJ3GejKWdMgd+q6u0vsZy/AN73gL4
# 4d0C5rNY50GyyFpotle3DFTO0QQdT1oXkqB/smYfoqXElpaGJgUV7CKinj3WcL4h
# FxQBAlflWx6ypz4zFuhPpukcRvcKE7TvBJRCZvKlkAY6woR2B4jMg67x2dSWlbb7
# CT17Bt1JJQ==
# SIG # End signature block
