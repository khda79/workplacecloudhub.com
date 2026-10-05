<#
.SYNOPSIS
Synthetic regression tests for the complete SmartInventory Microsoft Graph collector audit.
.VERSION
1.0.20
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$GitRevision,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$repoRoot = Split-Path $SourceRoot -Parent
$results = New-Object 'System.Collections.Generic.List[object]'
$sourceCache = @{}
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-Graph-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)

function Assert-Offline { param([bool]$Condition,[string]$Message) if(-not$Condition){throw $Message} }
function Test-OfflineCase {
    param([string]$Name,[scriptblock]$Body)
    try { & $Body; $results.Add([pscustomobject]@{Name=$Name;Passed=$true;Error=''}) }
    catch { $results.Add([pscustomobject]@{Name=$Name;Passed=$false;Error=$_.Exception.Message}) }
}
function Get-OfflineSourceText {
    param([string]$RelativeToSourceRoot)
    $cacheKey = '{0}|{1}' -f $GitRevision,$RelativeToSourceRoot
    if($sourceCache.ContainsKey($cacheKey)){return [string]$sourceCache[$cacheKey]}
    if([string]::IsNullOrWhiteSpace($GitRevision)){
        $text=[IO.File]::ReadAllText((Join-Path $SourceRoot $RelativeToSourceRoot))
        $sourceCache[$cacheKey]=$text
        return $text
    }
    $gitPath=('SmartM365/'+($RelativeToSourceRoot -replace '\\','/'))
    $lines=@(& git -c ("safe.directory={0}" -f $repoRoot) -C $repoRoot show ("{0}:{1}" -f $GitRevision,$gitPath) 2>$null)
    if($LASTEXITCODE -ne 0){throw "Unable to read $gitPath from $GitRevision."}
    $text=($lines -join "`n")
    $sourceCache[$cacheKey]=$text
    return $text
}
function Import-OfflineFunctions {
    param([string]$RelativeToSourceRoot,[string[]]$Names)
    $text=Get-OfflineSourceText $RelativeToSourceRoot
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseInput($text,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw "Source parse failed: $RelativeToSourceRoot"}
    $definitions=foreach($name in $Names){
        $node=$ast.Find({param($candidate) $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq $name},$true)
        if($null-eq$node){throw "Function not found: $name"}
        $node.Extent.Text
    }
    New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
}

$paths=[ordered]@{
    Discovered='SmartInventory/M365Inventory/IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1'
    BackupMailboxes='SmartInventory/ExchangeInventory/BackupProtection/SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1'
    DeviceSystem='SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Device-System-Inventory.ps1'
    Bios='SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-BIOS-Inventory.ps1'
    Compliance='SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1'
    Upgrade='SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-UpgradeEligibility.ps1'
    Rbac='SmartInventory/M365Inventory/IntuneInventory/RBAC/SmartM365-Intune-RBAC-GroupMembers.ps1'
    Remediations='SmartInventory/M365Inventory/IntuneInventory/SmartM365-Intune-ExportRemediationScripts.ps1'
    WindowsUpdate='SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1'
    Autopatch='SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'
    SharePoint='SmartInventory/M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1'
    Teams='SmartInventory/M365Inventory/Teams/SmartM365-Teams-Inventory.ps1'
    TeamsPhone='SmartInventory/M365Inventory/Teams/SmartM365-TeamsPhonePstnUsage-Inventory.ps1'
}

try {
    Test-OfflineCase 'Backup mailbox pagination rejects incomplete page-limited inventories' {
        $m=Import-OfflineFunctions $paths.BackupMailboxes @('Invoke-SmartM365GraphCollectionRequest')
        try {
            $all=& $m {
                function WriteLog { param([string]$Message,[string]$Level) }
                function Invoke-MgGraphRequest {
                    [CmdletBinding()]
                    param([string]$Method,[string]$Uri,[string]$OutputType)
                    if($Uri -eq 'p1'){return '{"value":[{"id":"one"}],"@odata.nextLink":"p2"}'}
                    return '{"value":[{"id":"two"}]}'
                }
                Invoke-SmartM365GraphCollectionRequest -Uri 'p1' -MaxPages 2
            }
            Assert-Offline (@($all).Count -eq 2) 'Complete two-page result was not returned.'
            $rejected=$false
            try {
                & $m {
                    function WriteLog { param([string]$Message,[string]$Level) }
                    function Invoke-MgGraphRequest {
                        [CmdletBinding()]
                        param([string]$Method,[string]$Uri,[string]$OutputType)
                        return '{"value":[{"id":"one"}],"@odata.nextLink":"p2"}'
                    }
                    Invoke-SmartM365GraphCollectionRequest -Uri 'p1' -MaxPages 1
                } | Out-Null
            } catch {
                $rejected=$_.Exception.Message -match 'Refusing to publish a partial inventory'
            }
            Assert-Offline $rejected 'A page-limited partial result was accepted.'
        } finally {Remove-Module $m -Force}
    }

    Test-OfflineCase 'All manual Graph pagers reject malformed pages and cycles in source' {
        foreach($entry in $paths.GetEnumerator()){
            $text=Get-OfflineSourceText $entry.Value
            Assert-Offline ($text -match 'repeated @odata\.nextLink') "$($entry.Key) has no nextLink cycle guard."
            Assert-Offline ($text -match 'without a value property|value property is missing') "$($entry.Key) has no malformed collection-page guard."
        }
    }

    Test-OfflineCase 'Teams Phone pager follows all pages' {
        $m=Import-OfflineFunctions $paths.TeamsPhone @('Get-SmartM365GraphCollection')
        try {
            $result=&$m { Get-SmartM365GraphCollection -InitialUri 'p1' -Operation fixture -RequestInvoker {param($u) if($u-eq'p1'){[pscustomobject]@{value=@([pscustomobject]@{id=1});'@odata.nextLink'='p2'}}else{[pscustomobject]@{value=@([pscustomobject]@{id=2})}}} }
            Assert-Offline (@($result.Items).Count -eq 2) 'Teams Phone did not return both pages.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Teams Phone pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.TeamsPhone @('Get-SmartM365GraphCollection')
        try{$caught=$false;try{&$m {Get-SmartM365GraphCollection -InitialUri p1 -Operation fixture -RequestInvoker {[pscustomobject]@{unexpected=1}}}|Out-Null}catch{$caught=$_.Exception.Message-match'value'};Assert-Offline $caught 'Malformed Teams Phone page was accepted.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Teams Phone pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.TeamsPhone @('Get-SmartM365GraphCollection')
        try{&$m {$script:n=0};$caught=$false;try{&$m {Get-SmartM365GraphCollection -InitialUri p1 -Operation fixture -RequestInvoker {$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Teams Phone cycle was not rejected.'}finally{Remove-Module $m -Force}
    }

    Test-OfflineCase 'Discovered Apps pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.Discovered @('Get-ShortGraphErrorMessage','Get-GraphRetryDelaySeconds','Invoke-GraphPagedRequest')
        try{&$m {$script:GraphMaxRetryAttempts=1;$script:GraphRetryDefaultSeconds=1;$script:GraphRetryMaxSeconds=1;$script:Stat_GraphCalls=0;$script:Stat_ThrottleRetries=0;function script:Invoke-MgGraphRequest{param($Method,$Uri,$OutputType,$StatusCodeVariable,$ErrorAction,[switch]$SkipHttpErrorCheck)Set-Variable -Name $StatusCodeVariable -Value 200 -Scope 1;[pscustomobject]@{unexpected=1}}};$caught=$false;try{&$m {Invoke-GraphPagedRequest -InitialUri p1 -MaxRetries 1}|Out-Null}catch{$caught=$_.Exception.Message-match'value'};Assert-Offline $caught 'Malformed Discovered Apps page was accepted.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Discovered Apps pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.Discovered @('Get-ShortGraphErrorMessage','Get-GraphRetryDelaySeconds','Invoke-GraphPagedRequest')
        try{&$m {$script:n=0;$script:GraphMaxRetryAttempts=1;$script:GraphRetryDefaultSeconds=1;$script:GraphRetryMaxSeconds=1;$script:Stat_GraphCalls=0;$script:Stat_ThrottleRetries=0;function script:Invoke-MgGraphRequest{param($Method,$Uri,$OutputType,$StatusCodeVariable,$ErrorAction,[switch]$SkipHttpErrorCheck)$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};Set-Variable -Name $StatusCodeVariable -Value 200 -Scope 1;[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}};$caught=$false;try{&$m {Invoke-GraphPagedRequest -InitialUri p1 -MaxRetries 1}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Discovered Apps cycle was not rejected.'}finally{Remove-Module $m -Force}
    }

    Test-OfflineCase 'Remediations pager follows all pages' {
        $m=Import-OfflineFunctions $paths.Remediations @('Get-ObjectValue','Test-ObjectProperty','Invoke-GraphGetAllPages')
        try{&$m {$script:n=0;function script:Invoke-GraphGet{param($Uri)$script:n++;if($script:n-eq1){[pscustomobject]@{value=@([pscustomobject]@{id=1});'@odata.nextLink'='p2'}}else{[pscustomobject]@{value=@([pscustomobject]@{id=2})}}}};$items=@(&$m {Invoke-GraphGetAllPages p1});Assert-Offline ($items.Count-eq2) 'Remediations did not return both pages.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Remediations pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.Remediations @('Get-ObjectValue','Test-ObjectProperty','Invoke-GraphGetAllPages')
        try{&$m {function script:Invoke-GraphGet{[pscustomobject]@{unexpected=1}}};$caught=$false;try{&$m {Invoke-GraphGetAllPages p1}|Out-Null}catch{$caught=$_.Exception.Message-match'value'};Assert-Offline $caught 'Malformed remediation page was accepted.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Remediations pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.Remediations @('Get-ObjectValue','Test-ObjectProperty','Invoke-GraphGetAllPages')
        try{&$m {$script:n=0;function script:Invoke-GraphGet{$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}};$caught=$false;try{&$m {Invoke-GraphGetAllPages p1}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Remediation cycle was not rejected.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Remediations pager accepts Hashtable Graph pages' {
        $m=Import-OfflineFunctions $paths.Remediations @('Get-ObjectValue','Test-ObjectProperty','Invoke-GraphGetAllPages')
        try{
            &$m {$script:n=0;function script:Invoke-GraphGet{param($Uri)$script:n++;if($script:n-eq1){@{value=@([pscustomobject]@{id='synthetic-remediation-1'});'@odata.nextLink'='p2'}}else{@{value=@([pscustomobject]@{id='synthetic-remediation-2'})}}}}
            $items=@(&$m {Invoke-GraphGetAllPages p1})
            Assert-Offline ($items.Count-eq2) 'Hashtable Graph pages were rejected or truncated by Remediations.'
        }finally{Remove-Module $m -Force}
    }

    Test-OfflineCase 'SharePoint pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.SharePoint @('Get-SpoPropertyValue','ConvertTo-SpoText','Get-SpoGraphPagedValues')
        try{$caught=$false;try{&$m {Get-SpoGraphPagedValues -Uri p1 -RequestInvoker {[pscustomobject]@{unexpected=1}}}|Out-Null}catch{$caught=$_.Exception.Message-match'value'};Assert-Offline $caught 'Malformed SharePoint page was accepted.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'SharePoint pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.SharePoint @('Get-SpoPropertyValue','ConvertTo-SpoText','Get-SpoGraphPagedValues')
        try{&$m {$script:n=0};$caught=$false;try{&$m {Get-SpoGraphPagedValues -Uri p1 -RequestInvoker {$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'SharePoint cycle was not rejected.'}finally{Remove-Module $m -Force}
    }

    Test-OfflineCase 'Compliance pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.Compliance @('Write-ComplianceInfo','Test-ComplianceGraphProperty','Get-ComplianceGraphPropertyValue','Get-ComplianceGraphPageShape','Invoke-GraphPagedCollection')
        try{&$m {function script:Invoke-WithRetry{param($Operation,$Script)&$Script};function script:Invoke-MgGraphRequest{[pscustomobject]@{unexpected=1}}};$caught=$false;try{&$m {Invoke-GraphPagedCollection -Uri p1}|Out-Null}catch{$caught=$_.Exception.Message-match'value' -and $_.Exception.Message-match'Type=' -and $_.Exception.Message-match'Keys='};Assert-Offline $caught 'Malformed compliance page was accepted or lacked safe shape diagnostics.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.Compliance @('Write-ComplianceInfo','Test-ComplianceGraphProperty','Get-ComplianceGraphPropertyValue','Get-ComplianceGraphPageShape','Invoke-GraphPagedCollection')
        try{&$m {$script:n=0;function script:WriteLog{param($Message,$Level)};function script:Invoke-WithRetry{param($Operation,$Script)&$Script};function script:Invoke-MgGraphRequest{$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}};$caught=$false;try{&$m {Invoke-GraphPagedCollection -Uri p1}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Compliance cycle was not rejected.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance pager accepts multi-page dictionary responses' {
        $m=Import-OfflineFunctions $paths.Compliance @('Write-ComplianceInfo','Test-ComplianceGraphProperty','Get-ComplianceGraphPropertyValue','Get-ComplianceGraphPageShape','Invoke-GraphPagedCollection')
        try{&$m {$script:n=0;function script:WriteLog{param($Message,$Level)};function script:Invoke-WithRetry{param($Operation,$Script)&$Script};function script:Invoke-MgGraphRequest{$script:n++;if($script:n-eq1){return @{value=@(@{id='one'});'@odata.nextLink'='p2'}};return @{value=@(@{id='two'})}}};$items=@(&$m {Invoke-GraphPagedCollection -Uri p1});Assert-Offline ($items.Count-eq2) 'Compliance dictionary pages were not fully collected.';Assert-Offline ($items[0]['id']-eq'one' -and $items[1]['id']-eq'two') 'Compliance dictionary page order or values changed.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance pager retains PSCustomObject compatibility' {
        $m=Import-OfflineFunctions $paths.Compliance @('Write-ComplianceInfo','Test-ComplianceGraphProperty','Get-ComplianceGraphPropertyValue','Get-ComplianceGraphPageShape','Invoke-GraphPagedCollection')
        try{&$m {function script:WriteLog{param($Message,$Level)};function script:Invoke-WithRetry{param($Operation,$Script)&$Script};function script:Invoke-MgGraphRequest{[pscustomobject]@{value=@([pscustomobject]@{id='object'})}}};$items=@(&$m {Invoke-GraphPagedCollection -Uri p1});Assert-Offline ($items.Count-eq1 -and $items[0].id-eq'object') 'Compliance PSCustomObject page compatibility regressed.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance Entra enrichment indexes a paged snapshot and preserves trust types' {
        $m=Import-OfflineFunctions $paths.Compliance @('Get-ComplianceGraphPropertyValue','Get-EntraDeviceIndex','Resolve-DirInfoFromEntraDevice')
        try {
            &$m {
                function script:Write-ComplianceInfo { param($Message) }
                function script:Invoke-GraphPagedCollection {
                    param($Uri,$Operation)
                    if($Uri -notmatch '/devices\?\$select=id,deviceId,trustType') { throw 'Unexpected Entra query.' }
                    @(@{id='object-h';deviceId='device-h';trustType='ServerAd'},[pscustomobject]@{id='object-r';deviceId='device-r';trustType='Workplace'},@{id='object-a';deviceId='device-a';trustType='AzureAd'})
                }
            }
            $actual=&$m {
                $index=Get-EntraDeviceIndex
                [pscustomobject]@{
                    Count=$index.Count
                    Hybrid=Resolve-DirInfoFromEntraDevice -EntraDevice $index['DEVICE-H'] -FallbackUpn 'user@example.test'
                    Registered=Resolve-DirInfoFromEntraDevice -EntraDevice $index['device-r']
                    AADOnly=Resolve-DirInfoFromEntraDevice -EntraDevice $index['device-a']
                    Missing=Resolve-DirInfoFromEntraDevice -EntraDevice $null -FallbackUpn 'user@example.test'
                }
            }
            Assert-Offline ($actual.Count-eq3 -and $actual.Hybrid.EntraObjectId-eq'object-h' -and $actual.Hybrid.DirectorySource-eq'Hybrid') 'Compliance Entra ObjectId or Hybrid classification was lost.'
            Assert-Offline ($actual.Registered.DirectorySource-eq'Registered' -and $actual.AADOnly.DirectorySource-eq'AADOnly') 'Compliance Entra trust types were not preserved.'
            Assert-Offline ($actual.Missing.DirectorySource-eq'NotFound' -and -not $actual.Missing.EntraObjectId -and -not $actual.Missing.AD_OU) 'Compliance unmatched device was misclassified or given an invented OU.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance Entra enrichment rejects empty and ambiguous snapshots' {
        $m=Import-OfflineFunctions $paths.Compliance @('Get-ComplianceGraphPropertyValue','Get-EntraDeviceIndex')
        try {
            &$m {function script:Write-ComplianceInfo { param($Message) }; function script:Invoke-GraphPagedCollection { @() }}
            $emptyRejected=$false;try{&$m {Get-EntraDeviceIndex}|Out-Null}catch{$emptyRejected=$_.Exception.Message -match 'no devices'}
            &$m {function script:Invoke-GraphPagedCollection { @(@{id='one';deviceId='same'},@{id='two';deviceId='same'}) }}
            $duplicateRejected=$false;try{&$m {Get-EntraDeviceIndex}|Out-Null}catch{$duplicateRejected=$_.Exception.Message -match 'Multiple Entra objects'}
            Assert-Offline ($emptyRejected -and $duplicateRejected) 'Compliance accepted an empty or ambiguous Entra snapshot.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance directory enrichment is default and no longer queries per device' {
        $text=Get-OfflineSourceText $paths.Compliance
        $template=Get-OfflineSourceText 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.local.json.template'
        Assert-Offline ($text -match '\[bool\]\$EnableDirectoryEnrichment\s*=\s*\$true' -and $template -match '"EnableDirectoryEnrichment"\s*:\s*true') 'Compliance Entra enrichment is not enabled by default.'
        Assert-Offline ($text -notmatch 'Resolve-DirInfoFromGraph|onPremisesDistinguishedName|onPremisesDomainName') 'Compliance still uses the per-device or unsupported directory lookup.'
    }
    Test-OfflineCase 'Compliance collection guards and fatal summary use canonical helpers' {
        $text=Get-OfflineSourceText $paths.Compliance
        Assert-Offline ($text -notmatch "PSObject\.Properties\['value'\]") 'Compliance still contains a PSCustomObject-only collection guard.'
        Assert-Offline ((($text|Select-String -Pattern "Test-ComplianceGraphProperty -InputObject .* -Name 'value'" -AllMatches).Matches.Count)-ge7) 'Compliance does not route every collection response through the canonical property helper.'
        Assert-Offline ($text.Contains('$global:SmartM365ErrorCount = [Math]::Max(1, [int]$global:SmartM365ErrorCount)')) 'Compliance fatal errors do not increment the execution error count.'
        Assert-Offline ($text.Contains("Complete-SmartM365ExecutionContext -Status 'Failed' -ErrorRecord `$script:ComplianceFatalError -FailureStage 'ComplianceInventory'")) 'Compliance fatal error details and failure stage are not passed to the failed execution summary.'
        Assert-Offline ($text.Contains("Complete-SmartM365ExecutionContext -Status 'Auto'")) 'Compliance successful execution does not use the automatic success summary.'
        Assert-Offline ($text -notmatch "Complete-SmartM365ExecutionContext -Status 'Auto'[^\r\n]*-FailureStage") 'Compliance successful execution still emits a failure stage.'
    }
    Test-OfflineCase 'Compliance legacy per-device detail collection remains disabled' {
        $text=Get-OfflineSourceText $paths.Compliance
        $template=Get-OfflineSourceText 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.local.json.template'
        Assert-Offline ($text -match '\[bool\]\$IncludePolicyStates\s*=\s*\$false') 'Compliance policy-state detail is not disabled by default.'
        Assert-Offline ($template -match '"IncludePolicyStates"\s*:\s*false') 'Compliance local template still enables policy-state detail by default.'
        Assert-Offline ($text -match '\$script:IncludePolicyStatesEffective\s*=\s*\$false') 'The retired per-device policy-state workflow can still be enabled.'
        Assert-Offline ($text -match 'PolicyStateAutoDisableDeviceThreshold' -and $text -match 'IncludePolicyStatesExplicit') 'Compliance large-tenant automatic safeguard is missing.'
        Assert-Offline ($text -match 'Write-ComplianceInfo -Message \("Detailed compliance policy-state collection was automatically disabled') 'Compliance automatic large-tenant safeguard still produces a warning status.'
        Assert-Offline ($text -match 'PolicyStateMaxRuntimeMinutes' -and $text -match 'Assert-PolicyStateRuntimeAvailable') 'Compliance runtime circuit breaker is missing.'
        Assert-Offline ($text -match 'Managed Windows devices selected for compliance summary:[^\r\n]+' -and $text -notmatch 'Write-Host \("Managed Windows devices selected') 'The selected-device milestone is still emitted without the timestamped logger.'
    }
    Test-OfflineCase 'Compliance canonical detail export is default and uses one launcher' {
        $text=Get-OfflineSourceText $paths.Compliance
        $launcher=Get-OfflineSourceText 'SmartInventory/Launchers/Cloud/Start-SmartM365-Devices-Compliance-Inventory.cmd'
        $template=Get-OfflineSourceText 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.local.json.template'
        $orchestrator=Get-OfflineSourceText 'SmartInventory/Orchestrator/Orchestrator-Jobs.json.template'
        Assert-Offline ($text -match '\[switch\]\$CollectPolicyDetails') 'Compliance collector does not expose the dedicated detail switch.'
        Assert-Offline ($text -match '\[switch\]\$SummaryOnly') 'Compliance collector does not expose the explicit summary-only escape hatch.'
        Assert-Offline ($text -match '\$script:PolicyExportEnabled\s*=\s*-not\s+\$SummaryOnly\.IsPresent') 'Canonical compliance detail export is not enabled by default.'
        Assert-Offline ($launcher -match '-File\s+"%SCRIPT_DIR%SmartM365-Devices-Compliance-Inventory\.ps1"\s+-Tenant\s+prod\s+-Connect\s+%\*') 'Standard compliance launcher does not run the default complete collection.'
        Assert-Offline (-not (Test-Path -LiteralPath (Join-Path $SourceRoot 'SmartInventory/Launchers/Cloud/Start-SmartM365-Devices-Compliance-Detailed-Inventory.cmd'))) 'The redundant detailed compliance launcher still exists.'
        Assert-Offline ($text -match "reportName\s*=\s*'DevicePolicySettingsComplianceReportV3'") 'Canonical compliance export does not use the qualified setting-level report.'
        Assert-Offline ($text -match '-CanonicalPath\s+\$policyMainCsv' -and $text -match '-TimestampedPath\s+\$policyTsCsv' -and $text -match '-LatestPath\s+\$policyLastCsv') 'Canonical compliance export is not wired to the official DATA-ALL and DATA-LAST paths.'
        Assert-Offline ($text -match '-SummaryRows\s+\$rows\.ToArray\(\)') 'Canonical compliance export does not safely convert the generic summary list before parameter binding.'
        Assert-Offline ($text -match 'Export-SmartM365Csv -Data \$exportOutput' -and $text -match 'last valid detailed DATA-LAST and SharePoint files were preserved') 'Canonical compliance export does not use the guarded publication path.'
        Assert-Offline ($text -match 'ExceptionType=\{0\}; Message=\{1\}; ScriptStackTrace=\{2\}') 'Canonical compliance export failure logging does not retain exception type and stack trace.'
        Assert-Offline ($template -match '"PolicyExportTimeoutMinutes"\s*:\s*30') 'Canonical compliance export timeout is missing from the local template.'
        Assert-Offline ($orchestrator -match '"Name"\s*:\s*"Intune-Devices-Compliance-Inventory"[\s\S]*?"Arguments"\s*:\s*""[\s\S]*?"EstimatedDurationMinutes"\s*:\s*15') 'Orchestrator compliance job is not aligned with the default complete collection.'
        Assert-Offline ($text -match "lastReportedDateTime\s*=\s*''") 'Canonical compliance export no longer preserves the empty legacy timestamp contract.'
    }
    Test-OfflineCase 'Compliance export derives deterministic policy states' {
        $m=Import-OfflineFunctions $paths.Compliance @('Get-ComplianceExportPolicyState')
        try {
            $errorState=&$m {Get-ComplianceExportPolicyState -Aggregate ([pscustomobject]@{HasError=$true;HasNonCompliant=$true;HasUnknown=$false;HasCompliant=$true;HasNotApplicable=$false})}
            $nonCompliantState=&$m {Get-ComplianceExportPolicyState -Aggregate ([pscustomobject]@{HasError=$false;HasNonCompliant=$true;HasUnknown=$false;HasCompliant=$true;HasNotApplicable=$false})}
            $unknownState=&$m {Get-ComplianceExportPolicyState -Aggregate ([pscustomobject]@{HasError=$false;HasNonCompliant=$false;HasUnknown=$true;HasCompliant=$true;HasNotApplicable=$false})}
            $compliantState=&$m {Get-ComplianceExportPolicyState -Aggregate ([pscustomobject]@{HasError=$false;HasNonCompliant=$false;HasUnknown=$false;HasCompliant=$true;HasNotApplicable=$true})}
            $notApplicableState=&$m {Get-ComplianceExportPolicyState -Aggregate ([pscustomobject]@{HasError=$false;HasNonCompliant=$false;HasUnknown=$false;HasCompliant=$false;HasNotApplicable=$true})}
            Assert-Offline ($errorState-eq'error' -and $nonCompliantState-eq'nonCompliant' -and $unknownState-eq'unknown' -and $compliantState-eq'compliant' -and $notApplicableState-eq'notApplicable') 'Compliance export state precedence changed.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance export streams, aggregates, and publishes only validated device-policy rows' {
        $fixtureCsv=Join-Path $testRoot 'compliance-export-fixture.csv'
        @'
DeviceId,PolicyId,SettingId,PolicyVersion,UserId,SettingInstanceId,SettingName,SettingNm,SettingStatus,StateDetails,ErrorCode,ErrorType,SettingValue
d1,p1,s1,8,u1,i1,secureBootEnabled,,Compliant,,,,
d1,p1,s2,8,u1,i2,bitLockerEnabled,,Not compliant,,,,
d2,p1,s1,8,u2,i1,secureBootEnabled,,Error,,,,
d2,p1,s2,8,u2,i2,bitLockerEnabled,,Compliant,,,,
d3,p1,s1,8,u3,i1,secureBootEnabled,,Compliant,,,,
d1,p2,s9,1,u1,i9,firewallEnabled,,Not compliant,,,,
'@ | Set-Content -LiteralPath $fixtureCsv -Encoding UTF8
        $m=Import-OfflineFunctions $paths.Compliance @('Get-SafeProperty','Map-SettingCategory','Get-ComplianceExportPolicyState','Invoke-CompliancePolicyExport')
        try {
            &$m {
                param($FixtureCsv)
                $script:fixtureCsv=$FixtureCsv
                $script:PolicyExportTimeoutMinutes=5
                $script:SettingRuleMap=@(
                    @{Pattern='secureboot(enabled)?';Category='SecureBoot'},
                    @{Pattern='bitlocker|encrypt';Category='BitLocker'}
                )
                $script:capturedRows=@()
                $script:capturedColumns=@()
                $script:capturedCanonicalPath=''
                $script:capturedPublication=$null
                $script:publishCalls=0
                function script:Write-ComplianceInfo { param($Message) }
                function script:WriteLog { param($Message,$Level) }
                function script:Invoke-WithRetry { param($Operation,$Script) &$Script }
                function script:Invoke-GraphPagedCollection {
                    [pscustomobject]@{id='p1';displayName='Synthetic Windows compliance';version=8;'@odata.type'='#microsoft.graph.windows10CompliancePolicy'}
                }
                function script:Get-PolicyConfiguredCategories {
                    $set=[System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    [void]$set.Add('SecureBoot');[void]$set.Add('BitLocker');return $set
                }
                function script:Invoke-MgGraphRequest {
                    param($Method,$Uri,$Body,$ContentType,$ErrorAction)
                    if($Method-eq'POST'){return [pscustomobject]@{id='synthetic-job'}}
                    return [pscustomobject]@{status='completed';url='https://example.invalid/synthetic.zip'}
                }
                function script:Invoke-WebRequest { param($Uri,$OutFile,$ErrorAction) [IO.File]::WriteAllBytes($OutFile,[byte[]]@()) }
                function script:Expand-Archive { param($LiteralPath,$DestinationPath,[switch]$Force) Copy-Item -LiteralPath $script:fixtureCsv -Destination (Join-Path $DestinationPath 'fixture.csv') -Force }
                function script:Write-SmartM365CsvAtomically {
                    param($Data,$Path,$Columns)
                    $script:capturedRows=@($Data);$script:capturedColumns=@($Columns);$script:capturedCanonicalPath=$Path
                }
                function script:Export-SmartM365Csv {
                    param($Data,$TimestampedPath,$LatestPath,$Columns)
                    $script:publishCalls++
                    $script:capturedPublication=[pscustomobject]@{Data=@($Data);TimestampedPath=$TimestampedPath;LatestPath=$LatestPath;Columns=@($Columns)}
                    return [pscustomobject]@{TimestampedPath=$TimestampedPath;LatestPath=$LatestPath}
                }
            } $fixtureCsv
            $devices=@([pscustomobject]@{id='d1'},[pscustomobject]@{id='d2'})
            $summary=New-Object System.Collections.Generic.List[object]
            $summary.Add([pscustomobject]@{DeviceName='DEVICE-1';AzureADDeviceId='aad-1';EntraObjectId='entra-1';AD_Domain='example.test';AD_OU='OU=One';DirectorySource='Hybrid'})
            $summary.Add([pscustomobject]@{DeviceName='DEVICE-2';AzureADDeviceId='aad-2';EntraObjectId='entra-2';AD_Domain='example.test';AD_OU='OU=Two';DirectorySource='Hybrid'})
            $canonical=Join-Path $testRoot 'canonical.csv'
            $timestamped=Join-Path $testRoot 'timestamped.csv'
            $latest=Join-Path $testRoot 'latest.csv'
            $exportResult=&$m {param($d,$s,$c,$t,$l) Invoke-CompliancePolicyExport -Devices $d -SummaryRows $s -CanonicalPath $c -TimestampedPath $t -LatestPath $l} $devices $summary.ToArray() $canonical $timestamped $latest
            $captured=&$m {[pscustomobject]@{Rows=@($script:capturedRows);Columns=@($script:capturedColumns);CanonicalPath=$script:capturedCanonicalPath;Publication=$script:capturedPublication;PublishCalls=$script:publishCalls}}
            $first=@($captured.Rows|Where-Object DeviceName -eq 'DEVICE-1')[0]
            $second=@($captured.Rows|Where-Object DeviceName -eq 'DEVICE-2')[0]
            Assert-Offline ($exportResult.RowCount-eq2 -and $exportResult.SelectedSettingRows-eq4 -and $exportResult.TotalSettingRows-eq6) 'Compliance export row filtering or counters changed.'
            Assert-Offline ($first.state-eq'nonCompliant' -and $first.settingCount-eq2 -and $first.nonCompliantSettingCount-eq1 -and $first.SecureBoot-eq'Pass' -and $first.'BitLocker/Encryption'-eq'Fail') ("Compliance export did not aggregate noncompliant settings and categories: {0}" -f ($first|ConvertTo-Json -Compress))
            Assert-Offline ($second.state-eq'error' -and $second.SecureBoot-eq'' -and $second.'BitLocker/Encryption'-eq'' -and $second.lastReportedDateTime-eq'') 'Compliance export did not preserve indeterminate category and timestamp semantics.'
            Assert-Offline ($captured.Columns -contains 'lastReportedDateTime' -and $captured.Columns -contains 'BitLocker/Encryption') 'Compliance export output contract lost required columns.'
            Assert-Offline ($captured.CanonicalPath-eq$canonical -and $captured.PublishCalls-eq1 -and $captured.Publication.TimestampedPath-eq$timestamped -and $captured.Publication.LatestPath-eq$latest) 'Validated compliance export was not published to all canonical paths.'

            &$m {
                $script:publishCalls=0
                function script:Write-SmartM365CsvAtomically { param($Data,$Path,$Columns) throw 'synthetic canonical write failure' }
                function script:Export-SmartM365Csv { param($Data,$TimestampedPath,$LatestPath,$Columns) $script:publishCalls++ }
            }
            $caught=$false
            try { &$m {param($d,$s,$c,$t,$l) Invoke-CompliancePolicyExport -Devices $d -SummaryRows $s -CanonicalPath $c -TimestampedPath $t -LatestPath $l} $devices $summary.ToArray() $canonical $timestamped $latest | Out-Null } catch { $caught=$_.Exception.Message-match'synthetic canonical write failure' }
            $publishCallsAfterFailure=&$m {$script:publishCalls}
            Assert-Offline ($caught -and $publishCallsAfterFailure-eq0) 'A failed canonical write still attempted to replace DATA-LAST or upload to SharePoint.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance batch progress reports rate ETA and completion' {
        $m=Import-OfflineFunctions $paths.Compliance @('Get-GraphBatchRetryDelaySeconds','Assert-PolicyStateRuntimeAvailable','Invoke-GraphBatchWithSubRequestRetry')
        try {
            &$m {
                $script:GraphMaxRetryAttempts=2
                $script:GraphRetryMaxSeconds=1
                $script:PolicyStateDeadlineUtc=[datetime]::UtcNow.AddMinutes(5)
                $script:PolicyStateMaxRuntimeMinutes=5
                $script:PolicyStateBatchProgressInterval=1
                $script:PolicyStateCollectionDisabled=$false
                $script:PolicyDetailCollectionComplete=$true
                $script:PolicyStateCircuitBreakerLogged=$false
                $script:PolicyStateBatchRetryCount=0
                $script:PolicyStateBatchThrottleCount=0
                $script:messages=[System.Collections.Generic.List[string]]::new()
                function script:Write-ComplianceInfo { param($Message) [void]$script:messages.Add([string]$Message) }
                function script:Write-ComplianceWarning { param($Message) [void]$script:messages.Add([string]$Message) }
                function script:Invoke-WithRetry { param($Operation,$Script) &$Script }
                function script:Invoke-MgGraphRequest {
                    param($Method,$Uri,$Body,$ContentType,$ErrorAction)
                    $payload=$Body|ConvertFrom-Json
                    [pscustomobject]@{responses=@($payload.requests|ForEach-Object{[pscustomobject]@{id=[string]$_.id;status=200;body=[pscustomobject]@{value=@()};headers=@{}}})}
                }
            }
            $requests=@(1..45|ForEach-Object{[pscustomobject]@{id=[string]$_;method='GET';url='/synthetic'}})
            $map=&$m {param($r) Invoke-GraphBatchWithSubRequestRetry -Requests $r -Operation 'Get Intune compliance policy states batch'} $requests
            $state=&$m {[pscustomobject]@{Messages=@($script:messages);Retries=$script:PolicyStateBatchRetryCount;Throttles=$script:PolicyStateBatchThrottleCount}}
            $joined=$state.Messages -join "`n"
            Assert-Offline ($map.Count-eq45) 'Compliance batch fixture did not complete all sub-requests.'
            Assert-Offline ($joined -match 'batch 1/3' -and $joined -match 'batch 3/3' -and $joined -match 'rate ' -and $joined -match 'ETA ') 'Compliance batch progress is missing batch/rate/ETA details.'
            Assert-Offline ($state.Retries-eq0 -and $state.Throttles-eq0) 'Compliance batch counters changed on successful responses.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance policy-state circuit breaker preserves summary processing' {
        $m=Import-OfflineFunctions $paths.Compliance @('Assert-PolicyStateRuntimeAvailable')
        try {
            &$m {
                $script:PolicyStateDeadlineUtc=[datetime]::UtcNow.AddSeconds(-1)
                $script:PolicyStateMaxRuntimeMinutes=1
                $script:PolicyStateCollectionDisabled=$false
                $script:PolicyDetailCollectionComplete=$true
                $script:PolicyStateCircuitBreakerLogged=$false
                $script:warnings=0
                function script:Write-ComplianceWarning { param($Message) $script:warnings++ }
            }
            $caught=$false
            try { &$m {Assert-PolicyStateRuntimeAvailable -Operation fixture} } catch { $caught=[bool]$_.Exception.Data['SmartM365PolicyStateCircuitBreaker'] }
            $state=&$m {[pscustomobject]@{Disabled=$script:PolicyStateCollectionDisabled;Complete=$script:PolicyDetailCollectionComplete;Logged=$script:PolicyStateCircuitBreakerLogged;Warnings=$script:warnings}}
            Assert-Offline ($caught -and $state.Disabled -and -not $state.Complete -and $state.Logged -and $state.Warnings-eq1) 'Compliance circuit breaker did not stop details while preserving the summary path.'
        } finally {Remove-Module $m -Force}
    }

    Test-OfflineCase 'Autopatch pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.Autopatch @('Invoke-GraphGetAll')
        try{&$m {function script:Invoke-AutopatchGraphRequest{[pscustomobject]@{unexpected=1}}};$caught=$false;try{&$m {Invoke-GraphGetAll p1}|Out-Null}catch{$caught=$_.Exception.Message-match'value'};Assert-Offline $caught 'Malformed Autopatch page was accepted.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Autopatch pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.Autopatch @('Invoke-GraphGetAll')
        try{&$m {$script:n=0;function script:Invoke-AutopatchGraphRequest{$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}};$caught=$false;try{&$m {Invoke-GraphGetAll p1}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Autopatch cycle was not rejected.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Autopatch excludes only no-alert code zero and retains nonzero alerts' {
        $m=Import-OfflineFunctions $paths.Autopatch @('Test-AutopatchAlertMessage','Convert-FeatureRowsToAlertDetails','Convert-QualityRowsToAlertDetails','Convert-QualityErrorRowsToAlertDetails','Group-AlertSummary')
        try {
            $observed=&$m {
                $rows=@(
                    [pscustomobject]@{LatestAlertMessage='0';DeviceId='one';PolicyId='policy';DeviceName='PC1'},
                    [pscustomobject]@{LatestAlertMessage=' 0 ';DeviceId='two';PolicyId='policy';DeviceName='PC2'},
                    [pscustomobject]@{LatestAlertMessage='42';DeviceId='three';PolicyId='policy';DeviceName='PC3'}
                )
                $feature=@(Convert-FeatureRowsToAlertDetails -Rows $rows -PolicyMap @{})
                $quality=@(Convert-QualityRowsToAlertDetails -Rows $rows -PolicyMap @{})
                $errors=@(Convert-QualityErrorRowsToAlertDetails -Rows @([pscustomobject]@{AlertMessage='0';PolicyId='policy'},[pscustomobject]@{AlertMessage='nonzero';PolicyId='policy'}) -PolicyMap @{})
                $summary=@(Group-AlertSummary -Details $feature)
                [pscustomobject]@{Feature=$feature;Quality=$quality;Errors=$errors;Summary=$summary;Unknown=(Test-AutopatchAlertMessage '987654')}
            }
            Assert-Offline ($observed.Feature.Count -eq 1 -and $observed.Feature[0].AlertName -eq '42') 'Feature report retained no-alert zero or lost a nonzero code.'
            Assert-Offline ($observed.Quality.Count -eq 1 -and $observed.Quality[0].AlertName -eq '42') 'Quality report retained no-alert zero or lost a nonzero code.'
            Assert-Offline ($observed.Errors.Count -eq 1 -and $observed.Errors[0].AlertName -eq 'nonzero') 'Quality error report mishandled zero or a real alert.'
            Assert-Offline ($observed.Summary.Count -eq 1 -and $observed.Summary[0].Impact -eq 1 -and $observed.Unknown) 'Autopatch summary or unknown nonzero alert was lost.'
            $source=Get-OfflineSourceText $paths.Autopatch
            Assert-Offline ($source -match 'Quality update coverage: Graph policy profiles=') 'Quality policy/report row coverage is not logged.'
        }
        finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Autopatch reports recovered transient retry attempts' {
        $m=Import-OfflineFunctions $paths.Autopatch @('Get-AutopatchGraphStatusCode','Get-AutopatchGraphRetryDelaySeconds','Invoke-AutopatchGraphRequest')
        try {
            $observed=&$m {
                $script:GraphTransientRetryCount=0
                $script:requestCount=0
                function script:Invoke-MgGraphRequest {
                    [CmdletBinding()]
                    param($Method,$Uri,$OutputType,$Body,$ContentType)
                    $script:requestCount++
                    if($script:requestCount -eq 1){throw 'synthetic HTTP 429'}
                    [pscustomobject]@{id='recovered'}
                }
                function script:Write-Log { param($Message,$Level) }
                function script:Start-Sleep { param($Seconds) }
                $response=Invoke-AutopatchGraphRequest -Method GET -Uri 'synthetic://report' -MaxAttempts 2
                [pscustomobject]@{Id=$response.id;Requests=$script:requestCount;Retries=$script:GraphTransientRetryCount}
            }
            Assert-Offline ($observed.Id -eq 'recovered' -and $observed.Requests -eq 2 -and $observed.Retries -eq 1) 'Autopatch did not count a recovered transient attempt.'
            $text=Get-OfflineSourceText $paths.Autopatch
            Assert-Offline ($text.Contains('$script:CompletionStatus = ''CompletedWithWarnings''')) 'Autopatch still reports Success after logged warnings.'
        } finally {Remove-Module $m -Force}
    }

    Test-OfflineCase 'Windows Update pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.WindowsUpdate @('Invoke-GraphGetAllPages')
        try{&$m {function script:Invoke-GraphRestMethodWithRetry{[pscustomobject]@{unexpected=1}}};$caught=$false;try{&$m {Invoke-GraphGetAllPages -Uri p1 -Headers @{}}|Out-Null}catch{$caught=$_.Exception.Message-match'value'};Assert-Offline $caught 'Malformed Windows Update page was accepted.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Windows Update pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.WindowsUpdate @('Invoke-GraphGetAllPages')
        try{&$m {$script:n=0;function script:Invoke-GraphRestMethodWithRetry{$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}};$caught=$false;try{&$m {Invoke-GraphGetAllPages -Uri p1 -Headers @{}}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Windows Update cycle was not rejected.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Windows Update OS distribution deduplicates and classifies devices' {
        $m=Import-OfflineFunctions $paths.WindowsUpdate @('Get-WinUpdateRowValue','Normalize-DeviceName','Get-WinUpdateOsBuild','Get-WinUpdateOsCoverageSummary')
        try {
            $rows=@(
                [pscustomobject]@{DeviceId='device-1';DeviceName='PC-01';OSVersion='10.0.26100.1'},
                [pscustomobject]@{DeviceId='device-1';DeviceName='PC-01';OSVersion='10.0.19045.1'},
                [pscustomobject]@{DeviceId='device-2';DeviceName='PC-02';OSVersion='10.0.22631.1'},
                [pscustomobject]@{DeviceId='';DeviceName='PC-03';OSVersion='10.0.19045.1'},
                [pscustomobject]@{DeviceId='';DeviceName='PC-04';OSVersion=''},
                [pscustomobject]@{DeviceId='device-5';DeviceName='PC-05';OSVersion='6.1.7601.0'}
            )
            $coverage=&$m {param($r)Get-WinUpdateOsCoverageSummary -Rows $r -MinimumBuild 26100} $rows
            Assert-Offline ($coverage.TotalDevices-eq5) 'The OS distribution did not deduplicate devices.'
            Assert-Offline ($coverage.Windows11-eq2 -and $coverage.Windows10-eq1 -and $coverage.UnknownOrOther-eq2) 'Windows 11, Windows 10, and unknown/other classification is incorrect.'
            Assert-Offline ($coverage.Windows11Pct-eq40 -and $coverage.Windows10Pct-eq20 -and $coverage.UnknownOrOtherPct-eq40) 'OS distribution percentages are incorrect.'
            Assert-Offline ($coverage.Covered-eq1 -and $coverage.UnknownOsVersion-eq1) 'Existing OS coverage metrics changed unexpectedly.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Windows Update fleet accepts tenant-scoped Graph rows before CSV publication' {
        $m=Import-OfflineFunctions $paths.WindowsUpdate @('Get-WinUpdateOsBuild','New-WinUpdateActivationIndex','Get-WinUpdateActivationState','Get-WinUpdateIntuneFleetCountrySummary','Get-WinUpdateAdWithoutIntuneSummary')
        try {
            $users=@([pscustomobject]@{TenantKey='org-prod';'Object Id'='user-1';CountryOrRegion='France'})
            $devices=@(
                [pscustomobject]@{TenantKey='org-prod';OS='Windows';'Device ID'='device-1';'Azure AD Device ID'='id-1';'OS version'='10.0.26100.1';UserId='user-1'},
                [pscustomobject]@{TenantKey='org-prod';OS='Windows';'Device ID'='device-2';'Azure AD Device ID'='id-2';'OS version'='10.0.19045.1';UserId='user-1'},
                [pscustomobject]@{TenantKey='org-prod';OS='Windows';'Device ID'='device-3';'Azure AD Device ID'='id-3';'OS version'='';UserId=''},
                [pscustomobject]@{TenantKey='org-prod';OS='Windows';'Device ID'='device-4';'Azure AD Device ID'='id-4';'OS version'='10.0.19045.1';UserId=''},
                [pscustomobject]@{TenantKey='org-prod';OS='Windows';'Device ID'='device-5';'Azure AD Device ID'='id-5';'OS version'='10.0.19045.1';UserId=''}
            )
            $entra=@(
                [pscustomobject]@{TenantKey='org-prod';DeviceId='id-1';AccountEnabled='True'},
                [pscustomobject]@{TenantKey='org-prod';DeviceId='id-2';AccountEnabled='False'},
                [pscustomobject]@{TenantKey='org-prod';DeviceId='id-3';AccountEnabled='False'},
                [pscustomobject]@{TenantKey='org-prod';DeviceId='id-4';AccountEnabled='False'},
                [pscustomobject]@{TenantKey='org-prod';DeviceId='id-7';AccountEnabled='True'},
                [pscustomobject]@{TenantKey='other-prod';DeviceId='id-5';AccountEnabled='True'}
            )
            $ad=@(
                [pscustomobject]@{TenantKey='org-prod';ObjectGUID='id-1';Enabled='False';OperatingSystemShortName='Windows 11';IsActiveInLast45Days='False'},
                [pscustomobject]@{TenantKey='org-prod';ObjectGUID='id-2';Enabled='True';OperatingSystemShortName='Windows 10';IsActiveInLast45Days='True'},
                [pscustomobject]@{TenantKey='org-prod';ObjectGUID='id-3';Enabled='False';OperatingSystemShortName='Windows 10';IsActiveInLast45Days='False'},
                [pscustomobject]@{TenantKey='org-prod';ObjectGUID='id-6';Enabled='True';OperatingSystemShortName='Windows 11';IsActiveInLast45Days='True'},
                [pscustomobject]@{TenantKey='org-prod';ObjectGUID='id-7';Enabled='False';OperatingSystemShortName='Windows 10';IsActiveInLast45Days='False'}
            )
            $index=&$m {param($e,$a)New-WinUpdateActivationIndex -EntraDeviceRows $e -AdComputerRows $a -TenantKey 'org-prod'} $entra $ad
            $graphRows=@(
                [pscustomobject]@{PolicyId='reference';DeviceId='device-1'},
                [pscustomobject]@{PolicyId='other';DeviceId='device-1'},
                [pscustomobject]@{PolicyId='other';DeviceId='device-2'},
                [pscustomobject]@{PolicyId='reference';DeviceId='report-only'}
            )
            $runtime=&$m {param($d,$u,$p,$i)Get-WinUpdateIntuneFleetCountrySummary -IntuneDeviceRows $d -ActiveUserRows $u -PolicyRows $p -ActivationIndex $i -ReferencePolicyId 'reference' -TenantKey 'org-prod'} $devices $users $graphRows $index
            Assert-Offline ($runtime.Total.Total-eq2 -and $runtime.Total.Windows11-eq1 -and $runtime.Total.Windows10-eq1 -and $runtime.Total.ReferencePolicy-eq1 -and $runtime.Total.OtherPolicyOnly-eq1 -and $runtime.Total.NeitherPolicy-eq0 -and $runtime.ReportOnlyReferenceIds-eq1) 'Untagged Graph policy rows did not reconcile with enabled Intune devices.'
            Assert-Offline ($runtime.AllIntuneWindows-eq5 -and $runtime.ExcludedDisabledBoth-eq1 -and $runtime.ExcludedUnqualified-eq2) 'Disabled or unqualified Intune devices were included.'
            $publishedRows=@($graphRows|ForEach-Object{[pscustomobject]@{TenantKey='org-prod';PolicyId=$_.PolicyId;DeviceId=$_.DeviceId}})+@([pscustomobject]@{TenantKey='other-prod';PolicyId='reference';DeviceId='device-3'})
            $published=&$m {param($d,$u,$p,$i)Get-WinUpdateIntuneFleetCountrySummary -IntuneDeviceRows $d -ActiveUserRows $u -PolicyRows $p -ActivationIndex $i -ReferencePolicyId 'reference' -TenantKey 'org-prod'} $devices $users $publishedRows $index
            Assert-Offline ($published.Total.ReferencePolicy-eq1 -and $published.Total.OtherPolicyOnly-eq1 -and $published.Total.NeitherPolicy-eq0) 'Tagged policy rows were not tenant-isolated.'
            $adSummary=&$m {param($a,$d,$i)Get-WinUpdateAdWithoutIntuneSummary -AdComputerRows $a -IntuneDeviceRows $d -ActivationIndex $i -TenantKey 'org-prod'} $ad $devices $index
            Assert-Offline ($adSummary.UnmatchedTotal-eq2 -and $adSummary.All.Total-eq2 -and $adSummary.EnabledActive45.Total-eq1) 'AD activation did not use the Entra-or-AD rule.'
            $wrongTenantRejected=$false
            try { &$m {param($d,$u,$p,$i)Get-WinUpdateIntuneFleetCountrySummary -IntuneDeviceRows $d -ActiveUserRows $u -PolicyRows $p -ActivationIndex $i -ReferencePolicyId 'reference' -TenantKey 'prod'} $devices $users $graphRows $index | Out-Null } catch {$wrongTenantRejected=$true}
            Assert-Offline $wrongTenantRejected 'The profile selector was accepted as a CSV tenant key.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Windows Update Entra OS fallback uses the Intune ID bridge' {
        $m=Import-OfflineFunctions $paths.WindowsUpdate @('Get-WinUpdateOsBuild','Add-WinUpdateEntraOsVersionFallback')
        try {
            $rows=@(
                [pscustomobject]@{DeviceId='managed-1';DeviceName='PC-01';OSVersion='';ReadinessMatch='NotMatched'},
                [pscustomobject]@{DeviceId='managed-2';DeviceName='PC-02';OSVersion='10.0.19045.1';ReadinessMatch='Matched'},
                [pscustomobject]@{DeviceId='managed-3';DeviceName='PC-03';OSVersion='';ReadinessMatch='NotMatched'}
            )
            $intune=@(
                [pscustomobject]@{'Device ID'='managed-1';'Azure AD Device ID'='entra-1'},
                [pscustomobject]@{'Device ID'='managed-2';'Azure AD Device ID'='entra-2'}
            )
            $entra=@(
                [pscustomobject]@{DeviceId='entra-1';OperatingSystemVersion='10.0.22631.1'},
                [pscustomobject]@{DeviceId='entra-1';OperatingSystemVersion='10.0.26100.2'},
                [pscustomobject]@{DeviceId='entra-2';OperatingSystemVersion='10.0.26100.3'}
            )
            $result=&$m {param($r,$i,$e)Add-WinUpdateEntraOsVersionFallback -Rows $r -IntuneDeviceRows $i -EntraDeviceRows $e} $rows $intune $entra
            Assert-Offline ($result.Rows[0].OSVersion-eq'10.0.26100.2' -and $result.Rows[0].OSVersionSource-eq'Entra') 'The Entra fallback did not use the Intune ID bridge or the highest available build.'
            Assert-Offline ($result.Rows[1].OSVersion-eq'10.0.19045.1' -and $result.Rows[1].OSVersionSource-eq'Readiness') 'An existing OS version was overwritten by the fallback.'
            Assert-Offline ([string]::IsNullOrWhiteSpace($result.Rows[2].OSVersion) -and $result.Rows[2].OSVersionSource-eq'Unavailable') 'An unmatched device was not retained as unavailable.'
            Assert-Offline ($result.ExistingVersionRows-eq1 -and $result.EntraVersionRows-eq1 -and $result.UnavailableRows-eq1) 'The Entra fallback counters are incorrect.'
        } finally {Remove-Module $m -Force}
    }
    Test-OfflineCase 'Windows Update Entra fallback converts the generic row list safely' {
        $text=Get-OfflineSourceText $paths.WindowsUpdate
        Assert-Offline ($text.Contains('-Rows $enrichedRows.ToArray()')) 'The Entra fallback does not convert its generic row list with ToArray().'
        Assert-Offline (-not $text.Contains('-Rows @($enrichedRows)')) 'The Entra fallback still uses the failing generic-list array subexpression.'
        $rows=New-Object System.Collections.Generic.List[object]
        $rows.Add([pscustomobject]@{DeviceId='managed-1';OSVersion='';ReadinessMatch='NotMatched'})|Out-Null
        $converted=$rows.ToArray()
        Assert-Offline ($converted -is [object[]] -and $converted.Count-eq1) 'Generic row-list conversion did not produce object[].'
    }
    Test-OfflineCase 'Windows Update status age uses status dates and reports coverage' {
        $text=Get-OfflineSourceText $paths.WindowsUpdate
        Assert-Offline ($text.Contains("'LastUpdateStatusTime'") -and $text.Contains("'LastUpdateStatusDateTime'")) 'Supported last-status dates are missing.'
        Assert-Offline (-not $text.Contains("'PolicyLastModifiedTime'")) 'Policy modification time must not masquerade as a device status date.'
        Assert-Offline ($text.Contains('DaysSinceLastStatus coverage:') -and $text.Contains('LastStatusAgeCoverage=')) 'Status-age coverage is not visible in run logs.'
        Assert-Offline ($text.Contains("Metric='Last-status date source'")) 'The report does not identify the status-age source.'
    }
    Test-OfflineCase 'Windows Update email places the OS distribution before fleet coverage' {
        $text=Get-OfflineSourceText $paths.WindowsUpdate
        Assert-Offline ($text.Contains('Windows version distribution')) 'The Windows version distribution title is missing.'
        Assert-Offline ($text -match '(?s)\$severityInputsHtml\s+\$windowsVersionDistributionSection\s+\$intuneFleetSection\s+\$adComparisonSection\s+\$fleetCoverageSection') 'The fleet and AD tables are not placed between OS distribution and coverage.'
        Assert-Offline ($text.Contains('#2563eb') -and $text.Contains('#f59e0b') -and $text.Contains('#94a3b8')) 'The approved Windows distribution colors are missing.'
        Assert-Offline ($text.Contains('OS VERSION UNKNOWN') -and -not $text.Contains('UNKNOWN / OTHER')) 'The OS-version-unavailable label is not explicit.'
    }
    Test-OfflineCase 'Windows Update report health uses proportional severity thresholds' {
        $m=Import-OfflineFunctions $paths.WindowsUpdate @('Get-WinUpdateReportHealth')
        try {
            $currentRun=&$m {Get-WinUpdateReportHealth -TotalUniqueDevices 20701 -ActionRequiredCount 132 -Priority0Count 70 -CoverageAvailable $true -KnownOsCoveragePct 93.55 -FleetDevices 19547 -OsVersionUnknownCount 2415}
            Assert-Offline ($currentRun.Status-eq'WARNING') 'The validated current-run inputs should be WARNING.'
            Assert-Offline ($currentRun.Priority0Threshold-eq208) 'The proportional P0 threshold should be 208 for 20,701 devices.'
            Assert-Offline ($currentRun.ActionRequiredRatePct-eq0.64 -and $currentRun.OsVersionUnknownRatePct-eq12.35) 'The severity rates were not calculated as expected.'

            $p0Critical=&$m {Get-WinUpdateReportHealth -TotalUniqueDevices 20701 -ActionRequiredCount 208 -Priority0Count 208 -CoverageAvailable $true -KnownOsCoveragePct 99 -FleetDevices 19547 -OsVersionUnknownCount 0}
            Assert-Offline ($p0Critical.Status-eq'CRITICAL') 'The proportional P0 threshold did not trigger CRITICAL.'

            $actionRateCritical=&$m {Get-WinUpdateReportHealth -TotalUniqueDevices 10000 -ActionRequiredCount 100 -Priority0Count 0 -CoverageAvailable $true -KnownOsCoveragePct 99 -FleetDevices 10000 -OsVersionUnknownCount 0}
            Assert-Offline ($actionRateCritical.Status-eq'CRITICAL') 'A 1 percent action-required rate did not trigger CRITICAL.'

            $coverageCritical=&$m {Get-WinUpdateReportHealth -TotalUniqueDevices 10000 -ActionRequiredCount 0 -Priority0Count 0 -CoverageAvailable $true -KnownOsCoveragePct 89.99 -FleetDevices 10000 -OsVersionUnknownCount 0}
            Assert-Offline ($coverageCritical.Status-eq'CRITICAL') 'Known-OS coverage below 90 percent did not trigger CRITICAL.'

            $unknownCritical=&$m {Get-WinUpdateReportHealth -TotalUniqueDevices 10000 -ActionRequiredCount 0 -Priority0Count 0 -CoverageAvailable $true -KnownOsCoveragePct 99 -FleetDevices 10000 -OsVersionUnknownCount 1500}
            Assert-Offline ($unknownCritical.Status-eq'CRITICAL') 'An OS-unknown rate of 15 percent did not trigger CRITICAL.'

            $healthy=&$m {Get-WinUpdateReportHealth -TotalUniqueDevices 10000 -ActionRequiredCount 0 -Priority0Count 0 -CoverageAvailable $true -KnownOsCoveragePct 95 -FleetDevices 10000 -OsVersionUnknownCount 499}
            Assert-Offline ($healthy.Status-eq'OK') 'Healthy boundary inputs should be OK.'

            $unavailable=&$m {Get-WinUpdateReportHealth -TotalUniqueDevices 10000 -ActionRequiredCount 0 -Priority0Count 0 -CoverageAvailable $false -KnownOsCoveragePct 0 -FleetDevices 0 -OsVersionUnknownCount 0}
            Assert-Offline ($unavailable.Status-eq'WARNING') 'Unavailable OS coverage should be WARNING.'
        } finally {Remove-Module $m -Force}
    }

    Test-OfflineCase 'RBAC pager follows all pages and rejects cycles' {
        $m=Import-OfflineFunctions $paths.Rbac @('Test-RbacGraphProperty','Get-RbacGraphPropertyValue','Get-RbacGraphPageShape','Get-RbacGraphCollection')
        try{$items=@(&$m {Get-RbacGraphCollection p1 -RequestInvoker {param($u)if($u-eq'p1'){[pscustomobject]@{value=@([pscustomobject]@{id=1});'@odata.nextLink'='p2'}}else{[pscustomobject]@{value=@([pscustomobject]@{id=2})}}}});Assert-Offline ($items.Count-eq2) 'RBAC did not return both pages.';$caught=$false;try{&$m {Get-RbacGraphCollection p1 -RequestInvoker {[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'RBAC cycle was not rejected.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'RBAC pager accepts Hashtable Graph pages' {
        $m=Import-OfflineFunctions $paths.Rbac @('Test-RbacGraphProperty','Get-RbacGraphPropertyValue','Get-RbacGraphPageShape','Get-RbacGraphCollection')
        try{$items=@(&$m {Get-RbacGraphCollection p1 -RequestInvoker {param($u)if($u-eq'p1'){@{value=@([pscustomobject]@{id='hash-1'});'@odata.nextLink'='p2'}}else{@{value=@([pscustomobject]@{id='hash-2'})}}}});Assert-Offline ($items.Count-eq2 -and $items[0].id-eq'hash-1' -and $items[1].id-eq'hash-2') 'RBAC rejected or truncated Hashtable Graph pages.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'RBAC malformed page reports response shape without values' {
        $m=Import-OfflineFunctions $paths.Rbac @('Test-RbacGraphProperty','Get-RbacGraphPropertyValue','Get-RbacGraphPageShape','Get-RbacGraphCollection')
        try{$caught=$false;try{&$m {Get-RbacGraphCollection p1 -RequestInvoker {@{unexpected='secret-value'}}}|Out-Null}catch{$caught=$_.Exception.Message-match'Hashtable' -and $_.Exception.Message-match'unexpected' -and $_.Exception.Message-notmatch'secret-value'};Assert-Offline $caught 'RBAC malformed-page diagnostic omitted its safe response shape or exposed a value.'}finally{Remove-Module $m -Force}
    }

    Test-OfflineCase 'Licensing retries transient HTTP 500 and honors Retry-After' {
        $m=Import-OfflineFunctions 'SmartInventory/M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1' @('Invoke-GraphWithRetry')
        try{&$m {$script:n=0;$script:slept=0;function script:Start-Sleep{param($Seconds)$script:slept=$Seconds}};$value=&$m {Invoke-GraphWithRetry -MaxRetries 3 -ScriptBlock {$script:n++;if($script:n-eq1){$e=[Exception]::new('synthetic 500');$e.Data['StatusCode']=500;$e.Data['Retry-After']=7;throw $e};'ok'}};$state=&$m {[pscustomobject]@{Calls=$script:n;Slept=$script:slept}};Assert-Offline ($value-eq'ok'-and$state.Calls-eq2-and$state.Slept-eq7) 'Licensing retry did not honor the transient status and Retry-After delay.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Remediations retries throttling and honors Retry-After' {
        $m=Import-OfflineFunctions $paths.Remediations @('Get-IntuneRemediationRetryDelaySeconds','Invoke-GraphGet')
        try{&$m {$script:n=0;$script:slept=0;function script:Start-Sleep{param($Seconds)$script:slept=$Seconds};function script:Write-Warning{param($Message)};function script:Invoke-MgGraphRequest{$script:n++;if($script:n-eq1){$e=[Exception]::new('synthetic 429');$e.Data['StatusCode']=429;$e.Data['Retry-After']=6;throw $e};[pscustomobject]@{value=@()}}};$null=&$m {Invoke-GraphGet -Uri p1 -MaxAttempts 3};$state=&$m {[pscustomobject]@{Calls=$script:n;Slept=$script:slept}};Assert-Offline ($state.Calls-eq2-and$state.Slept-eq6) 'Remediations retry did not honor Retry-After.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Device System retries transient HTTP 500 and honors Retry-After' {
        $m=Import-OfflineFunctions $paths.DeviceSystem @('Invoke-GraphSafe')
        try{&$m {$script:n=0;$script:slept=0;function script:Start-Sleep{param($Seconds)$script:slept=$Seconds};function script:WriteLog{param($Message,$Level)};function script:Invoke-MgGraphRequest{$script:n++;if($script:n-eq1){$e=[Exception]::new('synthetic 500');$e.Data['StatusCode']=500;$e.Data['Retry-After']=5;throw $e};[pscustomobject]@{value=@()}}};$null=&$m {Invoke-GraphSafe -Uri p1 -MaxRetries 2};$state=&$m {[pscustomobject]@{Calls=$script:n;Slept=$script:slept}};Assert-Offline ($state.Calls-eq2-and$state.Slept-eq5) 'Device System retry did not honor Retry-After.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Device System accepts Hashtable run-state pages and uses bounded pacing defaults' {
        $m=Import-OfflineFunctions $paths.DeviceSystem @('Test-GraphResponseProperty','Get-GraphResponsePropertyValue','Get-PlatformScriptRunStates')
        try{
            &$m {
                $script:RunStatePageSize=500;$script:RunStatePagePauseMilliseconds=0;$script:n=0
                function script:WriteLog{param($Message,$Level)}
                function script:Invoke-GraphSafe{param($Uri,$MaxRetries,$BaseDelaySeconds)$script:n++;if($script:n-eq1){@{value=@([pscustomobject]@{managedDevice=[pscustomobject]@{id='synthetic-device-1'};resultMessage='SecureBoot:True';lastStateUpdateDateTime='2026-01-01T00:00:00Z'});'@odata.nextLink'='p2'}}else{@{value=@([pscustomobject]@{managedDevice=[pscustomobject]@{id='synthetic-device-2'};resultMessage='SecureBoot:False';lastStateUpdateDateTime='2026-01-01T00:00:00Z'})}}}
            }
            $result=&$m {Get-PlatformScriptRunStates -ScriptId 'synthetic-script'}
            Assert-Offline ($result.Count-eq2) 'Device System rejected or truncated Hashtable run-state pages.'
            $source=Get-OfflineSourceText $paths.DeviceSystem
            Assert-Offline ($source -match '\$ManagedDevicePageSize\s*=\s*500' -and $source -match '\$ManagedDevicePagePauseMilliseconds\s*=\s*500' -and $source -match '\$MaxRetries\s*=\s*8') 'Device System bounded throttle settings are missing.'
        }finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'BIOS accepts Hashtable and PSCustomObject Graph collection pages' {
        $m=Import-OfflineFunctions $paths.Bios @('Test-BiosGraphResponseProperty','Get-BiosGraphResponsePropertyValue')
        try{
            $hashPage=@{value=@([pscustomobject]@{id='synthetic-bios-1'});'@odata.nextLink'='p2'}
            $objectPage=[pscustomobject]@{value=@([pscustomobject]@{id='synthetic-bios-2'});'@odata.nextLink'='p3'}
            $hashValue=&$m {param($p) Get-BiosGraphResponsePropertyValue -Response $p -Name 'value'} $hashPage
            $objectValue=&$m {param($p) Get-BiosGraphResponsePropertyValue -Response $p -Name 'value'} $objectPage
            $hashNext=&$m {param($p) Get-BiosGraphResponsePropertyValue -Response $p -Name '@odata.nextLink'} $hashPage
            Assert-Offline ((&$m {param($p) Test-BiosGraphResponseProperty -Response $p -Name 'value'} $hashPage) -and $hashValue[0].id-eq'synthetic-bios-1' -and $hashNext-eq'p2') 'BIOS rejected a Hashtable Graph collection page.'
            Assert-Offline ((&$m {param($p) Test-BiosGraphResponseProperty -Response $p -Name 'value'} $objectPage) -and $objectValue[0].id-eq'synthetic-bios-2') 'BIOS rejected a PSCustomObject Graph collection page.'
            $source=Get-OfflineSourceText $paths.Bios
            Assert-Offline ($source -match 'Test-BiosGraphResponseProperty\s+-Response\s+\$page\s+-Name\s+''value''' -and $source -match 'Get-BiosGraphResponsePropertyValue\s+-Response\s+\$page\s+-Name\s+''@odata\.nextLink''') 'BIOS managed-device pagination does not use the dictionary-safe accessors.'
        }finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'SharePoint retries transient HTTP 409 and honors Retry-After' {
        $m=Import-OfflineFunctions $paths.SharePoint @('Invoke-SpoWithRetry')
        try{&$m {$script:n=0;$script:slept=0;function script:Start-Sleep{param($Seconds)$script:slept=$Seconds};function script:Write-SpoLog{param($Message,$Level)}};$value=&$m {Invoke-SpoWithRetry -MaxAttempts 3 -ScriptBlock {$script:n++;if($script:n-eq1){$e=[Exception]::new('synthetic 409');$e.Data['StatusCode']=409;$e.Data['Retry-After']=4;throw $e};'ok'}};$state=&$m {[pscustomobject]@{Calls=$script:n;Slept=$script:slept}};Assert-Offline ($value-eq'ok'-and$state.Calls-eq2-and$state.Slept-eq4) 'SharePoint retry did not honor Retry-After.'}finally{Remove-Module $m -Force}
    }

    Test-OfflineCase 'Discovered Apps atomic copy preserves prior destination on replacement failure' {
        $m=Import-OfflineFunctions $paths.Discovered @('Copy-DiscoveredAppsFileAtomically')
        $source=Join-Path $testRoot 'source.csv';$destination=Join-Path $testRoot 'latest.csv'
        [IO.File]::WriteAllText($source,'new');[IO.File]::WriteAllText($destination,'last-valid')
        try{&$m {function script:Move-Item{throw 'synthetic replace failure'}};$caught=$false;try{&$m {param($s,$d)Copy-DiscoveredAppsFileAtomically $s $d} $source $destination}catch{$caught=$true};Assert-Offline $caught 'Synthetic replacement failure was not raised.';Assert-Offline ([IO.File]::ReadAllText($destination)-eq'last-valid') 'Prior latest CSV was modified by a failed replacement.'}finally{Remove-Module $m -Force}
    }

    Test-OfflineCase 'Compliance canonical export validation precedes publication' {
        $text=Get-OfflineSourceText $paths.Compliance
        $gate=$text.IndexOf("if (`$aggregates.Count -eq 0)",[StringComparison]::Ordinal)
        $materialize=$text.IndexOf('$exportOutput = @($exportRows',[StringComparison]::Ordinal)
        $canonicalWrite=$text.IndexOf('Write-SmartM365CsvAtomically -Data $exportOutput',[StringComparison]::Ordinal)
        $publish=$text.IndexOf('Export-SmartM365Csv -Data $exportOutput',[StringComparison]::Ordinal)
        Assert-Offline ($gate-ge0 -and $materialize-gt$gate -and $canonicalWrite-gt$materialize -and $publish-gt$canonicalWrite) 'Compliance export can publish before the complete validated result is materialized.'
        Assert-Offline ($text -match 'last valid detailed DATA-LAST and SharePoint files were preserved') 'Compliance export failure does not explicitly preserve the last valid published files.'
    }

    Test-OfflineCase 'Upgrade Eligibility defers profile identity to the canonical core context' {
        $text=Get-OfflineSourceText $paths.Upgrade
        Assert-Offline ($text -notmatch 'TenantKey\s*=\s*\[string\]\$Tenant') 'Upgrade Eligibility still stamps the profile selector into TenantKey.'
        Assert-Offline (($text|Select-String -Pattern 'Export-SmartM365Csv -Data' -AllMatches).Matches.Count -ge 4) 'Upgrade Eligibility no longer publishes through the identity-aware exporter.'

        $m=Import-OfflineFunctions 'Modules/SmartM365.Core/SmartM365.Core.psm1' @('Get-SmartM365CoreContextValue','Add-SmartM365TenantKeyToCsvData')
        try {
            &$m {
                $script:SmartM365CoreTenantKey='synthetic-canonical'
                $script:SmartM365CoreOrganizationKey='synthetic-org'
                $script:SmartM365CoreEnvironmentKey='test'
                $script:SmartM365CoreTenantId='00000000-0000-0000-0000-000000000001'
            }
            $rows=@(
                [pscustomobject]@{ReportType='HardwareReadinessSummary';RunId='summary-run'},
                [pscustomobject]@{DeviceName='DEVICE-001';GraphId='graph-001';RunId='detail-run'}
            )
            $result=&$m {param($r)Add-SmartM365TenantKeyToCsvData -Data $r} $rows
            Assert-Offline (@($result.Data).Count -eq 2) 'Canonical identity injection changed the Upgrade Eligibility row count.'
            Assert-Offline (@($result.Data|Where-Object {$_.TenantKey -cne 'synthetic-canonical'}).Count -eq 0) 'Canonical TenantKey was not injected into every Upgrade Eligibility row.'
            Assert-Offline (($result.Data[0].PSObject.Properties.Name|Select-Object -First 4)-join',' -ceq 'TenantKey,OrganizationKey,EnvironmentKey,TenantId') 'Upgrade Eligibility identity column order changed.'
        } finally {Remove-Module $m -Force}
    }

    Test-OfflineCase 'Upgrade Eligibility accepts dictionary and object Graph pages' {
        $m=Import-OfflineFunctions $paths.Upgrade @('Test-UpgradeEligibilityGraphProperty','Get-UpgradeEligibilityGraphProperty','Get-UpgradeEligibilityGraphCollection')
        try {
            $result=&$m {
                function script:WriteLogSmartM365 { param($Message,$Level) }
                $orderedPage=[ordered]@{
                    value=@([pscustomobject]@{id='ordered-1'})
                    '@odata.nextLink'='https://example.invalid/page2'
                }
                $dictionaryPage=[Collections.Generic.Dictionary[string,object]]::new()
                $dictionaryPage['value']=@([pscustomobject]@{id='dictionary-2'})
                $dictionaryPage['@odata.nextLink']='https://example.invalid/page3'
                $objectPage=[pscustomobject]@{value=@([pscustomobject]@{id='object-3'})}
                $pages=@{
                    'https://example.invalid/page1'=$orderedPage
                    'https://example.invalid/page2'=$dictionaryPage
                    'https://example.invalid/page3'=$objectPage
                }
                @(Get-UpgradeEligibilityGraphCollection -Uri 'https://example.invalid/page1' -Invoker {param($uri)$pages[$uri]})
            }
            Assert-Offline (@($result).Count -eq 3) 'Upgrade Eligibility pagination did not retain all dictionary/object rows.'
            Assert-Offline ((@($result.id) -join ',') -ceq 'ordered-1,dictionary-2,object-3') 'Upgrade Eligibility pagination changed row order.'

            $missingValueRejected=&$m {
                function script:WriteLogSmartM365 { param($Message,$Level) }
                try {
                    Get-UpgradeEligibilityGraphCollection -Uri 'https://example.invalid/malformed' -Invoker {param($uri)[ordered]@{'@odata.context'='synthetic'}} | Out-Null
                    return $false
                }
                catch { return $_.Exception.Message -like '*without a value property*' }
            }
            Assert-Offline $missingValueRejected 'Upgrade Eligibility accepted a malformed page without value.'
        } finally {Remove-Module $m -Force}
    }

    Test-OfflineCase 'Upgrade Eligibility replaces stale duplicate snapshots with an empty current CSV' {
        $m=Import-OfflineFunctions $paths.Upgrade @('Export-UpgradeEligibilityDuplicateAudit')
        try {
            $result=&$m {
                $script:exports=New-Object 'System.Collections.Generic.List[object]'
                function script:Export-SmartM365Csv {
                    param($Data,$TimestampedPath,$LatestPath,$Columns)
                    $script:exports.Add([pscustomobject]@{
                        Data=@($Data)
                        TimestampedPath=[string]$TimestampedPath
                        LatestPath=[string]$LatestPath
                        Columns=@($Columns)
                    }) | Out-Null
                }
                $withDuplicates=@(
                    [pscustomobject]@{DeviceName='DEVICE-01';NormalizedDeviceName='device-01';GraphId='graph-01';RunId='run-1'},
                    [pscustomobject]@{DeviceName='DEVICE-01.contoso.test';NormalizedDeviceName='device-01';GraphId='graph-02';RunId='run-1'}
                )
                $withoutDuplicates=@(
                    [pscustomobject]@{DeviceName='DEVICE-01';NormalizedDeviceName='device-01';GraphId='graph-01';RunId='run-2'},
                    [pscustomobject]@{DeviceName='DEVICE-02';NormalizedDeviceName='device-02';GraphId='graph-02';RunId='run-2'}
                )
                $first=Export-UpgradeEligibilityDuplicateAudit -Rows $withDuplicates -TimestampedPath 'history-1.csv' -LatestPath 'latest.csv'
                $second=Export-UpgradeEligibilityDuplicateAudit -Rows $withoutDuplicates -TimestampedPath 'history-2.csv' -LatestPath 'latest.csv'
                [pscustomobject]@{First=$first;Second=$second;Calls=$script:exports.ToArray()}
            }
            Assert-Offline ($result.Calls.Count -eq 2) 'The no-duplicate run did not publish a replacement snapshot.'
            Assert-Offline ($result.First.GroupCount -eq 1 -and $result.First.RowCount -eq 2) 'The duplicate run audit is incorrect.'
            Assert-Offline ($result.Second.GroupCount -eq 0 -and $result.Second.RowCount -eq 0) 'The no-duplicate run did not produce an empty audit.'
            Assert-Offline (@($result.Calls[1].Data).Count -eq 0) 'The replacement duplicate snapshot retained stale rows.'
            Assert-Offline ($result.Calls[1].LatestPath -ceq 'latest.csv') 'The no-duplicate run did not target the current snapshot.'
            Assert-Offline (@($result.Calls[1].Columns) -contains 'DuplicateKey' -and @($result.Calls[1].Columns) -contains 'RunId') 'The empty duplicate snapshot schema is incomplete.'
        } finally {Remove-Module $m -Force}
    }

    Test-OfflineCase 'Consumer contracts retain required Graph CSV sources' {
        $intelligenceSchema=Get-Content -LiteralPath (Join-Path (Split-Path $SourceRoot -Parent) 'SmartWorkplaceIntelligence/config/source-schema.json') -Raw|ConvertFrom-Json
        $intelligenceSelection=Get-Content -LiteralPath (Join-Path (Split-Path $SourceRoot -Parent) 'SmartWorkplaceIntelligence/config/source-selection.json') -Raw|ConvertFrom-Json
        $finops=Get-Content -LiteralPath (Join-Path (Split-Path $SourceRoot -Parent) 'SmartFinOps/Config/SmartFinOps-Workplace-SourceContracts.json') -Raw|ConvertFrom-Json
        $required=@('M365_Users_Active.csv','M365_Users_Activity.csv','M365_Licenses_Users.csv','M365_Entra_Devices.csv','Intune_Devices_Inventory.csv','Intune_Devices_Compliance.csv','Intune_Devices_UpgradeEligibility.csv','Intune_WindowsUpdate_Status.csv','M365_Teams_Teams.csv','M365_Teams_PhoneUserUsage.csv','M365_Copilot_UserUsage.csv')
        $serialized=(ConvertTo-Json @($intelligenceSchema,$intelligenceSelection,$finops) -Depth 30)
        foreach($file in $required){Assert-Offline ($serialized.Contains($file)) "Consumer contract missing $file."}
    }

    Test-OfflineCase 'SDK-backed list collectors request complete result sets' {
        $sdkPaths=@(
            'SmartInventory/M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1',
            'SmartInventory/M365Inventory/Domains/SmartM365-VerifiedDomains-Inventory.ps1',
            'SmartInventory/M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1',
            'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1',
            'SmartInventory/M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1',
            'SmartInventory/M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1'
        )
        foreach($path in $sdkPaths){Assert-Offline ((Get-OfflineSourceText $path)-match '(?m)Get-Mg\w+[^\r\n]*\s-All\b') "$path has no SDK -All collection path."}
    }
}
finally {
    if(Test-Path -LiteralPath $testRoot){Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue}
}

$summary=[pscustomobject]@{Engine=$PSVersionTable.PSVersion.ToString();GitRevision=$GitRevision;Passed=@($results|Where-Object Passed).Count;Failed=@($results|Where-Object{-not$_.Passed}).Count;Total=$results.Count;Results=$results}
if($ResultPath){$parent=Split-Path $ResultPath -Parent;if($parent -and -not(Test-Path $parent)){New-Item -ItemType Directory -Path $parent -Force|Out-Null};$summary|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $ResultPath -Encoding UTF8}
$summary|ConvertTo-Json -Depth 8
if($summary.Failed -gt 0){exit 1}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCAMc4iBapVCiXL
# WTlqZGLzjz0lwm9L0r2eZtLQPyzEBKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCAG6H1G+aANdcu/KsxMR3+
# 9wwFUCahxGgTXYrZRSC8dzANBgkqhkiG9w0BAQEFAASCAYALbZILFKpBWy4v3opW
# VgXz6QOdRgxFWMfp0X3g8pqJvfYI2nXiiDfLdnUrGxR3BZebjHchJ3kiqHvhJXwD
# D5odvdcZAiHxEwb3fTsrHQ8k0J/RzVPb0yPpHZR75/hrsm3XSk0OoK4iV1TqRls1
# L9QcHkTV/PA4nmF71eGW9IKNU8ZR4sI5k+FWf0JBtoO1jo5FvdMv0pzuZspXV/Sv
# zsU1A37qkzHkeZiiB3hM7ai3pmvtFlzpZHzV4/i7bN3GadUnSVBYFY0c1ewZ0USR
# LHFCAlQlItCYFb2tlX1xlnVyPrcxLlno6jPmjZg5bQAnzfQwleG6YFGlFP645Vok
# 5q5Ce4Hkv9Nmwdp4sNYNY7rTvdqQcLTaBWP+lQUmbMiDA2LBdHn8AibxZ8YzKJaS
# Um9Ype5ljl6vswT9m6jYsr+QXx7tjiIcbuulLTow9v8ZI1PlJv9L/PwA7LrrFG/4
# r1K/WjnKpFDupJgCIOTruzc+ulw/MTQ8fpY/+DRIUuhTOrw=
# SIG # End signature block
