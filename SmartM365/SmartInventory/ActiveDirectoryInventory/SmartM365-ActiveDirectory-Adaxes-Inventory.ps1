#Requires -Version 7.0
<#
.SYNOPSIS
Collects configurable Adaxes attributes of Active Directory users, keyed by ObjectGUID.
.DESCRIPTION
Reads users through the Adaxes service with the Adaxes PowerShell module, so Adaxes virtual
attributes (adm-*) are available. The attributes and their CSV column names come from the
Attributes list of the script configuration. AD_Adaxes_Users.csv is joined to the AD inventory
on ObjectGUID. Exception to the folder rule: reads go through Adaxes, not the ActiveDirectory module.
.VERSION
1.0.0
.REQUIREMENTS
Adaxes PowerShell module (Adaxes installer component) on the collection server, and read access
in Adaxes to the configured attributes for the running account.
#>
[CmdletBinding()]
param([string]$Tenant='test',[switch]$ValidateOnly,[switch]$SelfTest,[switch]$EnableConfiguredExternalActions,[string]$OutputPath,[string]$LatestCsvFolderPath)

# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$ScriptVersion='1.0.0'
$script:CompletionStatus='Success'; $script:CompletionError=$null; $script:Runtime=$null; $script:ServiceFailover=$false
$script:FixedColumns=@('RunId','CollectedAtUtc','CollectionStatus','IsPartialInventory','SourceScriptVersion','Domain','ObjectGUID')

# Attributes entries: { "Column": "<CSV column>", "Attribute": "<LDAP or adm-* name>" }.
function Get-AdaxesAttributeMap {
    param([AllowNull()]$Attributes)
    $map=[System.Collections.Generic.List[object]]::new()
    $columns=[System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($fixed in $script:FixedColumns){[void]$columns.Add($fixed)}
    [void]$columns.Add('TenantKey')
    foreach($entry in @($Attributes)){
        if($null -eq $entry){continue}
        $column=[string](Get-SmartM365EvidenceProperty $entry @('Column'))
        $attribute=[string](Get-SmartM365EvidenceProperty $entry @('Attribute'))
        if([string]::IsNullOrWhiteSpace($column) -or [string]::IsNullOrWhiteSpace($attribute)){throw 'Each Attributes entry needs a non-empty Column and Attribute.'}
        if($column -notmatch '^[A-Za-z][A-Za-z0-9_]*$'){throw "Attributes column name is not a plain identifier: $column"}
        if(-not $columns.Add($column)){throw "Attributes column name is reserved or duplicated: $column"}
        $map.Add([pscustomobject]@{Column=$column.Trim();Attribute=$attribute.Trim()})
    }
    if($map.Count -eq 0){throw 'The Attributes list of the configuration is empty; add at least one { Column, Attribute } entry.'}
    return @($map)
}

function ConvertTo-AdaxesCsvValue {
    param([AllowNull()]$Value)
    if($null -eq $Value){return ''}
    if($Value -is [byte[]]){return [Convert]::ToBase64String($Value)}
    # Adaxes returns UTC dates; an unspecified kind is treated as UTC, not as local time.
    if($Value -is [datetime]){
        $utc=if($Value.Kind -eq [DateTimeKind]::Local){$Value.ToUniversalTime()}else{[datetime]::SpecifyKind($Value,[DateTimeKind]::Utc)}
        return $utc.ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)
    }
    if($Value -is [string]){return $Value}
    if($Value -is [System.Collections.IEnumerable]){return ((@($Value) | ForEach-Object { ConvertTo-AdaxesCsvValue $_ }) -join ' | ')}
    return [string]$Value
}

function ConvertTo-AdaxesUserRow {
    param([Parameter(Mandatory)]$User,[Parameter(Mandatory)][string]$Domain,[Parameter(Mandatory)][object[]]$AttributeMap,[string]$RunId,[string]$CollectedAtUtc,[bool]$IsPartial)
    $row=[ordered]@{RunId=$RunId;CollectedAtUtc=$CollectedAtUtc;CollectionStatus='Observed';IsPartialInventory=$IsPartial;SourceScriptVersion=$ScriptVersion;Domain=$Domain;ObjectGUID=[string](Get-SmartM365EvidenceProperty $User @('ObjectGUID'))}
    foreach($item in $AttributeMap){$row[$item.Column]=ConvertTo-AdaxesCsvValue (Get-SmartM365EvidenceProperty $User @($item.Attribute))}
    return [pscustomobject]$row
}

# AdaxesService is one host name or a list tried in order: the next service is used when a read fails.
function Get-AdaxesServiceList {
    param([AllowNull()]$Configured)
    $services=[System.Collections.Generic.List[string]]::new()
    foreach($item in @($Configured)){
        $name=([string]$item).Trim()
        if($name -and -not ($services -contains $name)){$services.Add($name)}
    }
    if($services.Count -eq 0){throw 'AdaxesService is not configured: set one Adaxes service host name or a list of them in the script configuration.'}
    return @($services)
}

function Get-AdaxesDomainUsers {
    param([Parameter(Mandatory)][string[]]$Services,[Parameter(Mandatory)][string]$Domain,[Parameter(Mandatory)][string]$LdapFilter,[Parameter(Mandatory)][string[]]$Properties,[scriptblock]$Reader={param($s,$d,$f,$p) Get-AdmUser -AdaxesService $s -Server $d -LDAPFilter $f -Properties $p -ErrorAction Stop})
    $errors=[System.Collections.Generic.List[string]]::new()
    foreach($service in $Services){
        try{
            $users=@(& $Reader $service $Domain $LdapFilter $Properties)
            return [pscustomobject]@{Service=$service;Users=$users;Errors=@($errors)}
        }
        catch{
            if($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]){throw}
            $errors.Add(('{0}: {1}' -f $service,$_.Exception.Message))
        }
    }
    throw ($errors -join ' | ')
}

function Get-AdaxesDomainList {
    param([AllowNull()]$Configured)
    $domains=@(@($Configured) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { ([string]$_).Trim() } | Sort-Object -Unique)
    if($domains.Count -gt 0){return $domains}
    return @([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Domains | ForEach-Object { [string]$_.Name } | Sort-Object -Unique)
}

if($SelfTest){
    Import-Module (Join-Path $PSScriptRoot '..\Common\SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.11' -Force
    $map=Get-AdaxesAttributeMap @([pscustomobject]@{Column='AccountType';Attribute='adm-CustomAttributeText1'},[pscustomobject]@{Column='Generic';Attribute='adm-CustomAttributeBoolean1'},[pscustomobject]@{Column='LastSeen';Attribute='adm-CustomAttributeDate1'},[pscustomobject]@{Column='Tags';Attribute='adm-CustomAttributeTextMultiValue1'})
    $user=[pscustomobject]@{ObjectGUID=[guid]'2b26b47b-22a2-4fd1-81b4-6b1312936ffb';'adm-CustomAttributeText1'='Named';'adm-CustomAttributeBoolean1'=$false;'adm-CustomAttributeDate1'=[datetime]::new(2026,10,27,16,58,14,[DateTimeKind]::Utc);'adm-CustomAttributeTextMultiValue1'=@('A','B')}
    $row=ConvertTo-AdaxesUserRow $user 'contoso.local' $map 'run' 'now' $false
    if($row.ObjectGUID -ne '2b26b47b-22a2-4fd1-81b4-6b1312936ffb' -or $row.AccountType -ne 'Named' -or $row.Generic -ne 'False' -or $row.LastSeen -ne '2026-10-27T16:58:14Z' -or $row.Tags -ne 'A | B'){throw 'Adaxes row self-test failed.'}
    $empty=ConvertTo-AdaxesUserRow ([pscustomobject]@{ObjectGUID='x'}) 'contoso.local' $map 'run' 'now' $true
    if($empty.AccountType -ne '' -or $empty.IsPartialInventory -ne $true){throw 'Adaxes missing-attribute self-test failed.'}
    foreach($invalid in @(@(),@([pscustomobject]@{Column='ObjectGUID';Attribute='a'}),@([pscustomobject]@{Column='A';Attribute='a'},[pscustomobject]@{Column='a';Attribute='b'}),@([pscustomobject]@{Column='Bad Name';Attribute='a'}),@([pscustomobject]@{Column='A';Attribute=''}))){
        $rejected=$false; try{$null=Get-AdaxesAttributeMap $invalid}catch{$rejected=$true}
        if(-not $rejected){throw 'Adaxes attribute map validation self-test failed.'}
    }
    if(@(Get-AdaxesDomainList @(' b.local ','a.local','b.local')) -join ',' -ne 'a.local,b.local'){throw 'Adaxes domain list self-test failed.'}
    if((@(Get-AdaxesServiceList 'adx1') -join ',') -ne 'adx1' -or (@(Get-AdaxesServiceList @('adx1',' adx2 ','adx1','')) -join ',') -ne 'adx1,adx2'){throw 'Adaxes service list self-test failed.'}
    $rejected=$false; try{$null=Get-AdaxesServiceList @('',' ')}catch{$rejected=$true}; if(-not $rejected){throw 'Adaxes empty service list self-test failed.'}
    $reader={param($s,$d,$f,$p) if($s -eq 'down'){throw 'service unavailable'}; @([pscustomobject]@{ObjectGUID="$s/$d"})}
    $read=Get-AdaxesDomainUsers -Services @('down','up') -Domain 'contoso.local' -LdapFilter '(x)' -Properties @('a') -Reader $reader
    if($read.Service -ne 'up' -or $read.Users.Count -ne 1 -or $read.Errors.Count -ne 1 -or $read.Errors[0] -notlike 'down:*'){throw 'Adaxes service failover self-test failed.'}
    $rejected=$false; try{$null=Get-AdaxesDomainUsers -Services @('down') -Domain 'contoso.local' -LdapFilter '(x)' -Properties @('a') -Reader $reader}catch{$rejected=$_.Exception.Message -like 'down: service unavailable'}
    if(-not $rejected){throw 'Adaxes all-services-down self-test failed.'}
    Write-Output 'PASS: Adaxes users offline contract'; $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
}

try{
    $tenantContextPath=Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'Config\SmartM365-TenantContext.ps1'
    . $tenantContextPath
    $effectiveConfig=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $PSScriptRoot '..\Common\SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.11' -Force
    $script:Runtime=Initialize-SmartM365EvidenceRuntime -ScriptPath $PSCommandPath -EffectiveConfig $effectiveConfig -DefaultOutputRelativePath 'ActiveDirectory\Adaxes' -OutputPath $OutputPath -LatestCsvFolderPath $LatestCsvFolderPath -ValidateOnly:$ValidateOnly -EnableConfiguredExternalActions:$EnableConfiguredExternalActions
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $script:Runtime.LatestCsvFolderPath -ReadOnly:($ValidateOnly -or $SelfTest)
    $config=$script:Runtime.Config
    $services=Get-AdaxesServiceList (Get-SmartM365EvidenceProperty $config @('AdaxesService') @())
    $ldapFilter=[string](Get-SmartM365EvidenceProperty $config @('LdapFilter'))
    if([string]::IsNullOrWhiteSpace($ldapFilter)){throw 'LdapFilter is not configured.'}
    $attributeMap=Get-AdaxesAttributeMap (Get-SmartM365EvidenceProperty $config @('Attributes') @())
    $nonBlocking=@(@(Get-SmartM365EvidenceProperty $config @('NonBlockingDomainErrors') @()) | ForEach-Object { ([string]$_).Trim() })
    if(-not (Get-Module -ListAvailable -Name Adaxes)){throw 'Adaxes PowerShell module is not installed on this server. Install the Adaxes PowerShell Module component from the Adaxes installer.'}
    Import-Module Adaxes -ErrorAction Stop
    $domains=Get-AdaxesDomainList (Get-SmartM365EvidenceProperty $config @('Domains') @())
    if($domains.Count -eq 0){throw 'No domain to query: Domains is empty and the current forest returned no domain.'}
    Write-SmartM365EvidenceLog -Message ("Adaxes service(s) in order: {0}; {1} domain(s); {2} attribute(s): {3}." -f ($services -join ', '),$domains.Count,$attributeMap.Count,(($attributeMap | ForEach-Object { '{0}={1}' -f $_.Column,$_.Attribute }) -join ', '))

    $properties=@($attributeMap | ForEach-Object Attribute | Sort-Object -Unique)
    $usersByDomain=[ordered]@{}; $failed=[System.Collections.Generic.List[string]]::new()
    foreach($domain in $domains){
        try{
            $read=Get-AdaxesDomainUsers -Services $services -Domain $domain -LdapFilter $ldapFilter -Properties $properties
            foreach($skipped in $read.Errors){Write-SmartM365EvidenceLog -Message ("{0}: Adaxes service failed, next one tried: {1}" -f $domain,$skipped) -Level WARNING}
            if($read.Errors.Count -gt 0){$script:ServiceFailover=$true}
            $usersByDomain[$domain]=$read.Users
            Write-SmartM365EvidenceLog -Message ("{0}: {1} user(s) through {2}." -f $domain,$read.Users.Count,$read.Service) -Level SUCCESS
        }
        catch{
            if($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]){throw}
            # A partial domain read is never published: one unreadable object stops the whole domain.
            if($nonBlocking -notcontains $domain){throw ("Adaxes read failed for domain {0}: {1}" -f $domain,$_.Exception.Message)}
            $failed.Add($domain)
            Write-SmartM365EvidenceLog -Message ("{0}: skipped (NonBlockingDomainErrors): {1}" -f $domain,$_.Exception.Message) -Level WARNING
        }
    }
    if($usersByDomain.Count -eq 0){throw 'Adaxes read failed for every domain; previous export kept.'}
    $isPartial=$failed.Count -gt 0
    $rows=[System.Collections.Generic.List[object]]::new()
    foreach($domain in $usersByDomain.Keys){foreach($user in $usersByDomain[$domain]){$rows.Add((ConvertTo-AdaxesUserRow $user $domain $attributeMap $script:Runtime.RunId $script:Runtime.CollectedAtUtc $isPartial))}}
    if($ValidateOnly){ Write-SmartM365EvidenceLog -Message ("Adaxes read returned {0} user(s). No CSV published." -f $rows.Count) -Level SUCCESS; $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return }
    Export-SmartM365EvidenceDataset $script:Runtime 'AD_Adaxes_Users' @($rows) (@($script:FixedColumns) + @($attributeMap | ForEach-Object Column)) | Out-Null
    Write-SmartM365EvidenceLog -Message ("AD_Adaxes_Users: {0} user(s) from {1} domain(s); skipped domain(s): {2}." -f $rows.Count,$usersByDomain.Count,$(if($isPartial){$failed -join ', '}else{'none'})) -Level SUCCESS
    if($isPartial -or $script:ServiceFailover){$script:CompletionStatus='CompletedWithWarnings'}

    $smartM365RunGuard.ReachedTerminal=$true
}
catch{
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_
$script:CompletionStatus='Failed';$script:CompletionError=$_;throw}
finally{
    $smartM365RunGuard.InCleanup=$true
    if($script:Runtime){Complete-SmartM365EvidenceRuntime -Status $script:CompletionStatus -ErrorRecord $script:CompletionError -FailureStage 'AdaxesUsersInventory'}}
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}
