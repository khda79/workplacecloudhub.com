<#
.SYNOPSIS
Synthetic regression tests for the complete SmartInventory Microsoft Graph collector audit.
.VERSION
1.0.4
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
        $m=Import-OfflineFunctions $paths.Compliance @('Test-ComplianceGraphProperty','Get-ComplianceGraphPropertyValue','Get-ComplianceGraphPageShape','Invoke-GraphPagedCollection')
        try{&$m {function script:Invoke-WithRetry{param($Operation,$Script)&$Script};function script:Invoke-MgGraphRequest{[pscustomobject]@{unexpected=1}}};$caught=$false;try{&$m {Invoke-GraphPagedCollection -Uri p1}|Out-Null}catch{$caught=$_.Exception.Message-match'value' -and $_.Exception.Message-match'Type=' -and $_.Exception.Message-match'Keys='};Assert-Offline $caught 'Malformed compliance page was accepted or lacked safe shape diagnostics.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.Compliance @('Test-ComplianceGraphProperty','Get-ComplianceGraphPropertyValue','Get-ComplianceGraphPageShape','Invoke-GraphPagedCollection')
        try{&$m {$script:n=0;function script:Invoke-WithRetry{param($Operation,$Script)&$Script};function script:Invoke-MgGraphRequest{$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}};$caught=$false;try{&$m {Invoke-GraphPagedCollection -Uri p1}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Compliance cycle was not rejected.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance pager accepts multi-page dictionary responses' {
        $m=Import-OfflineFunctions $paths.Compliance @('Test-ComplianceGraphProperty','Get-ComplianceGraphPropertyValue','Get-ComplianceGraphPageShape','Invoke-GraphPagedCollection')
        try{&$m {$script:n=0;function script:Invoke-WithRetry{param($Operation,$Script)&$Script};function script:Invoke-MgGraphRequest{$script:n++;if($script:n-eq1){return @{value=@(@{id='one'});'@odata.nextLink'='p2'}};return @{value=@(@{id='two'})}}};$items=@(&$m {Invoke-GraphPagedCollection -Uri p1});Assert-Offline ($items.Count-eq2) 'Compliance dictionary pages were not fully collected.';Assert-Offline ($items[0]['id']-eq'one' -and $items[1]['id']-eq'two') 'Compliance dictionary page order or values changed.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance pager retains PSCustomObject compatibility' {
        $m=Import-OfflineFunctions $paths.Compliance @('Test-ComplianceGraphProperty','Get-ComplianceGraphPropertyValue','Get-ComplianceGraphPageShape','Invoke-GraphPagedCollection')
        try{&$m {function script:Invoke-WithRetry{param($Operation,$Script)&$Script};function script:Invoke-MgGraphRequest{[pscustomobject]@{value=@([pscustomobject]@{id='object'})}}};$items=@(&$m {Invoke-GraphPagedCollection -Uri p1});Assert-Offline ($items.Count-eq1 -and $items[0].id-eq'object') 'Compliance PSCustomObject page compatibility regressed.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Compliance collection guards and fatal summary use canonical helpers' {
        $text=Get-OfflineSourceText $paths.Compliance
        Assert-Offline ($text -notmatch "PSObject\.Properties\['value'\]") 'Compliance still contains a PSCustomObject-only collection guard.'
        Assert-Offline ((($text|Select-String -Pattern "Test-ComplianceGraphProperty -InputObject .* -Name 'value'" -AllMatches).Matches.Count)-ge7) 'Compliance does not route every collection response through the canonical property helper.'
        Assert-Offline ($text.Contains('$global:SmartM365ErrorCount = [Math]::Max(1, [int]$global:SmartM365ErrorCount)')) 'Compliance fatal errors do not increment the execution error count.'
        Assert-Offline ($text.Contains('Complete-SmartM365ExecutionContext -Status $finalStatus -ErrorRecord $script:ComplianceFatalError')) 'Compliance fatal error details are not passed to the execution summary.'
    }

    Test-OfflineCase 'Autopatch pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.Autopatch @('Invoke-GraphGetAll')
        try{&$m {function script:Invoke-AutopatchGraphRequest{[pscustomobject]@{unexpected=1}}};$caught=$false;try{&$m {Invoke-GraphGetAll p1}|Out-Null}catch{$caught=$_.Exception.Message-match'value'};Assert-Offline $caught 'Malformed Autopatch page was accepted.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Autopatch pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.Autopatch @('Invoke-GraphGetAll')
        try{&$m {$script:n=0;function script:Invoke-AutopatchGraphRequest{$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}};$caught=$false;try{&$m {Invoke-GraphGetAll p1}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Autopatch cycle was not rejected.'}finally{Remove-Module $m -Force}
    }

    Test-OfflineCase 'Windows Update pager rejects malformed page' {
        $m=Import-OfflineFunctions $paths.WindowsUpdate @('Invoke-GraphGetAllPages')
        try{&$m {function script:Invoke-GraphRestMethodWithRetry{[pscustomobject]@{unexpected=1}}};$caught=$false;try{&$m {Invoke-GraphGetAllPages -Uri p1 -Headers @{}}|Out-Null}catch{$caught=$_.Exception.Message-match'value'};Assert-Offline $caught 'Malformed Windows Update page was accepted.'}finally{Remove-Module $m -Force}
    }
    Test-OfflineCase 'Windows Update pager rejects repeated nextLink' {
        $m=Import-OfflineFunctions $paths.WindowsUpdate @('Invoke-GraphGetAllPages')
        try{&$m {$script:n=0;function script:Invoke-GraphRestMethodWithRetry{$script:n++;if($script:n-ge4){throw 'synthetic safety stop'};[pscustomobject]@{value=@();'@odata.nextLink'='p1'}}};$caught=$false;try{&$m {Invoke-GraphGetAllPages -Uri p1 -Headers @{}}|Out-Null}catch{$caught=$_.Exception.Message-match'repeated'};Assert-Offline $caught 'Windows Update cycle was not rejected.'}finally{Remove-Module $m -Force}
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

    Test-OfflineCase 'Compliance incomplete policy detail gate precedes publication' {
        $text=Get-OfflineSourceText $paths.Compliance
        $gate=$text.IndexOf("elseif (-not `$script:PolicyDetailCollectionComplete)",[StringComparison]::Ordinal)
        $publish=$text.IndexOf('Export-SmartM365Csv -Data @($polOut)',[StringComparison]::Ordinal)
        Assert-Offline ($gate-ge0 -and $publish-gt$gate) 'Compliance policy detail can publish before the completeness gate.'
        Assert-Offline (($text|Select-String -Pattern '\$script:PolicyDetailCollectionComplete = \$false' -AllMatches).Matches.Count -ge 2) 'Compliance failures do not mark policy detail incomplete.'
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

    Test-OfflineCase 'Consumer contracts retain required Graph CSV sources' {
        $dashboardSchema=Get-Content -LiteralPath (Join-Path (Split-Path $SourceRoot -Parent) 'SmartWorkplaceDashboard/source-schema.json') -Raw|ConvertFrom-Json
        $dashboardSelection=Get-Content -LiteralPath (Join-Path (Split-Path $SourceRoot -Parent) 'SmartWorkplaceDashboard/source-selection.json') -Raw|ConvertFrom-Json
        $finops=Get-Content -LiteralPath (Join-Path (Split-Path $SourceRoot -Parent) 'SmartFinOps/Config/SmartFinOps-Workplace-SourceContracts.json') -Raw|ConvertFrom-Json
        $required=@('M365_Users_Active.csv','M365_Users_Activity.csv','M365_Licenses_Users.csv','M365_Entra_Devices.csv','Intune_Devices_Inventory.csv','Intune_Devices_Compliance.csv','Intune_Devices_UpgradeEligibility.csv','Intune_WindowsUpdate_Status.csv','M365_Teams_Teams.csv','M365_Teams_PhoneUserUsage.csv','M365_Copilot_UserUsage.csv')
        $serialized=(ConvertTo-Json @($dashboardSchema,$dashboardSelection,$finops) -Depth 30)
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDXKYKYq+LRas2x
# q/mmtdbgGdr3eEZ/x8RwcA0J/NiBhKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP3I/ogt9UIHMjl522V9A182Bj/SQ94BESnxRaM2tA1JMA0GCSqG
# SIb3DQEBAQUABIIBgJGvfqIOknV1gOnB+agQzatQZpsH2oVEKCScXpcbRvnsC7+3
# NFrf5qwzPllfV7yh9Kn88bAkrfpRE3FCjigcD3m3SEm6Mh62ZLH81YwNxuG6Gqzj
# 2p3AygYx9ApI0tpPY6k01irWGaK2KhE4q1FUjGfkwrW1fwAEp4cwPGr5MMzqh0Yj
# NBYHxwawLrLA4f0jJT2wOYa3I6lu/dmJRNOVE6g9mzXh6dlHT3ZPDxYiOws7MzY3
# FFSsCJmw4hvAWRRsJNUEQBd3NOjfOUvli6XINmjW1kwulfcmo/IFMXbM3HxzNecQ
# 4dstfLtuGF0NkTBWzQ4uUHxZSMIfkygFYtg+Ln46Lk8ty4Saha1fuuls1TxVmL/v
# n/4X3JtYh+EAO3HnLTZiLjITagy9t+dfrDOgg9Oq+J2hGMwtqKjAeDpLZw64RZ4T
# gWJ7WGFExlp1PSXpsnPKeYXKB1uKKarLt5HdRntOa9/eZv+XWA0C+0zZKVUiCgsK
# 2H0eTiSvcofVDVJv8aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjEyMDQ5
# MDNaMC8GCSqGSIb3DQEJBDEiBCBcukXfy6weD+9MljwPDo/GZ6NSKoDzM6QTHCaA
# Ttx/tjANBgkqhkiG9w0BAQEFAASCAgAHTPFT0WHOi7yX/183mFb9/X98VbCoQqv2
# pQdz/tGM93O84EDzKQoNf+azX4rnwx/udn/6bPd7oqcmTil0/kb3etadIYzhbfgp
# xVFleMpkpuIL1VmDeK5LXDDc9w6umJhOz6PuJv3UEnjZt3FAP7P1maovvTnmUF6N
# sBYsDlOKm8uxKkIpMfx/ntGC/wLn5k+PMuI+CcqJpPyRoYZreYW1o5KyC0ECIA2X
# n+9aZHgTUaeBOcHznnnczGIyDD5Abo5scJZYdjyZEAlfehyLxfS0MZK7zLWden8S
# OEiBM3tUIhi098+sN9+s6bYXlZhNKh2snRoasgALf+vwX1LoscJW1q3jE3U+QbHx
# VMe2QZuuq8Rqy12IsUKvC26xnX/vCseStz7Fc4WVu+lPn4V9+XHVo5ImzpRS6ZNu
# V+mUWL2AjNuOsFoRiXp7n3oDAY5ofL81ZY9XuNZshzCAg0hKyIWfzA25SmEuPG9l
# nnESs7CdXosOkidVBb1WqOYduj8A2JOMUW+HHYZBedV5sB5rreGwqxLX+5pyaoQB
# 6WOfVhAhcoWyydCQpYUu5HPm7jqrychTrjBgnfUTY4/0SfCSO6IqJzGrlTSetpv/
# G7vNnmmgllnQ0VH1xE0qpqglbJ8AK3fKBtDoUz21rejfQQ6oITqavlfsdZhv/+dS
# JS4r0Mdymg==
# SIG # End signature block
