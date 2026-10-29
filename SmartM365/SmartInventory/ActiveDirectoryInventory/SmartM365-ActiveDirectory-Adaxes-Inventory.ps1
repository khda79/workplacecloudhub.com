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
1.0.1
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
$ScriptVersion='1.0.1'
$script:CompletionStatus='Success'; $script:CompletionError=$null; $script:Runtime=$null; $script:ReadWarnings=$false
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

# Get-AdmUser cannot build a user whose objectClass the running account cannot read, and that error ends the
# whole domain read on every service. Such an object is named in the error, excluded by DN and the domain is read again.
function Get-AdaxesUnreadableObjectDn {
    param([AllowNull()][string]$Message)
    $match=[regex]::Match([string]$Message,"attribute 'objectClass' of the directory object '(?<Dn>.+)' is unavailable")
    if($match.Success){return $match.Groups['Dn'].Value}
    return $null
}

function ConvertTo-AdaxesLdapFilterValue {
    param([Parameter(Mandatory)][string]$Value)
    $builder=[System.Text.StringBuilder]::new()
    foreach($char in $Value.ToCharArray()){
        switch($char){
            '\' {[void]$builder.Append('\5c')}
            '*' {[void]$builder.Append('\2a')}
            '(' {[void]$builder.Append('\28')}
            ')' {[void]$builder.Append('\29')}
            ([char]0) {[void]$builder.Append('\00')}
            default {[void]$builder.Append($char)}
        }
    }
    return $builder.ToString()
}

function Get-AdaxesDomainUsers {
    param([Parameter(Mandatory)][string[]]$Services,[Parameter(Mandatory)][string]$Domain,[Parameter(Mandatory)][string]$LdapFilter,[Parameter(Mandatory)][string[]]$Properties,[ValidateRange(0,1000)][int]$MaxUnreadableObjects=25,[scriptblock]$Reader={param($s,$d,$f,$p) Get-AdmUser -AdaxesService $s -Server $d -LDAPFilter $f -Properties $p -ErrorAction Stop})
    $serviceErrors=[System.Collections.Generic.List[string]]::new()
    $excluded=[System.Collections.Generic.List[string]]::new()
    while($true){
        $filter=if($excluded.Count -eq 0){$LdapFilter}else{'(&{0}{1})' -f $LdapFilter,(($excluded | ForEach-Object { '(!(distinguishedName={0}))' -f (ConvertTo-AdaxesLdapFilterValue $_) }) -join '')}
        $attemptErrors=[System.Collections.Generic.List[string]]::new(); $unreadable=$null
        foreach($service in $Services){
            try{
                $users=@(& $Reader $service $Domain $filter $Properties)
                return [pscustomobject]@{Service=$service;Users=$users;Errors=@($serviceErrors);Unreadable=@($excluded)}
            }
            catch{
                if($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]){throw}
                $unreadable=Get-AdaxesUnreadableObjectDn $_.Exception.Message
                if($unreadable){break}
                $attemptErrors.Add(('{0}: {1}' -f $service,$_.Exception.Message)); $serviceErrors.Add($attemptErrors[-1])
            }
        }
        if(-not $unreadable){throw ($attemptErrors -join ' | ')}
        if($excluded -contains $unreadable){throw "Excluding the unreadable object did not remove it from the search: $unreadable"}
        if($excluded.Count -ge $MaxUnreadableObjects){throw ("More than {0} unreadable object(s) (MaxUnreadableObjectsPerDomain); last: {1}" -f $MaxUnreadableObjects,$unreadable)}
        $excluded.Add($unreadable)
    }
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
    $denied='CN=svc\, test (old),OU=A*,DC=contoso,DC=local'
    $unreadableMessage="An instance of 'Softerra.Adaxes.PowerShellModule.Directory.ADUser' can't be created because the attribute 'objectClass' of the directory object '$denied' is unavailable."
    if((Get-AdaxesUnreadableObjectDn $unreadableMessage) -ne $denied -or $null -ne (Get-AdaxesUnreadableObjectDn 'other error')){throw 'Adaxes unreadable object parsing self-test failed.'}
    if((ConvertTo-AdaxesLdapFilterValue $denied) -ne 'CN=svc\5c, test \28old\29,OU=A\2a,DC=contoso,DC=local'){throw 'Adaxes LDAP filter escaping self-test failed.'}
    $script:SelfTestServices=[System.Collections.Generic.List[string]]::new()
    $reader={param($s,$d,$f,$p) $script:SelfTestServices.Add($s); if($f -notlike '*distinguishedName=*'){throw $unreadableMessage}; @([pscustomobject]@{ObjectGUID=$f})}
    $read=Get-AdaxesDomainUsers -Services @('adx1','adx2') -Domain 'contoso.local' -LdapFilter '(x)' -Properties @('a') -Reader $reader
    if($read.Service -ne 'adx1' -or (@($read.Unreadable) -join ';') -ne $denied -or $read.Errors.Count -ne 0 -or ($script:SelfTestServices -join ',') -ne 'adx1,adx1' -or $read.Users[0].ObjectGUID -ne '(&(x)(!(distinguishedName=CN=svc\5c, test \28old\29,OU=A\2a,DC=contoso,DC=local)))'){throw 'Adaxes unreadable object exclusion self-test failed.'}
    $always={param($s,$d,$f,$p) throw ($unreadableMessage -replace [regex]::Escape($denied),('CN=u{0},DC=c' -f [regex]::Matches($f,'distinguishedName').Count))}
    $rejected=$false; try{$null=Get-AdaxesDomainUsers -Services @('adx1') -Domain 'contoso.local' -LdapFilter '(x)' -Properties @('a') -MaxUnreadableObjects 2 -Reader $always}catch{$rejected=$_.Exception.Message -like 'More than 2 unreadable*'}
    if(-not $rejected){throw 'Adaxes unreadable object limit self-test failed.'}
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
    $maxUnreadable=[int](Get-SmartM365EvidenceProperty $config @('MaxUnreadableObjectsPerDomain') 25)
    if(-not (Get-Module -ListAvailable -Name Adaxes)){throw 'Adaxes PowerShell module is not installed on this server. Install the Adaxes PowerShell Module component from the Adaxes installer.'}
    Import-Module Adaxes -ErrorAction Stop
    $domains=Get-AdaxesDomainList (Get-SmartM365EvidenceProperty $config @('Domains') @())
    if($domains.Count -eq 0){throw 'No domain to query: Domains is empty and the current forest returned no domain.'}
    Write-SmartM365EvidenceLog -Message ("Adaxes service(s) in order: {0}; {1} domain(s); {2} attribute(s): {3}." -f ($services -join ', '),$domains.Count,$attributeMap.Count,(($attributeMap | ForEach-Object { '{0}={1}' -f $_.Column,$_.Attribute }) -join ', '))

    $properties=@($attributeMap | ForEach-Object Attribute | Sort-Object -Unique)
    $usersByDomain=[ordered]@{}; $failed=[System.Collections.Generic.List[string]]::new()
    foreach($domain in $domains){
        try{
            $read=Get-AdaxesDomainUsers -Services $services -Domain $domain -LdapFilter $ldapFilter -Properties $properties -MaxUnreadableObjects $maxUnreadable
            foreach($skipped in $read.Errors){Write-SmartM365EvidenceLog -Message ("{0}: Adaxes service failed, next one tried: {1}" -f $domain,$skipped) -Level WARNING}
            foreach($dn in $read.Unreadable){Write-SmartM365EvidenceLog -Message ("{0}: user skipped, its objectClass is not readable by the running account: {1}" -f $domain,$dn) -Level WARNING}
            if($read.Errors.Count -gt 0 -or $read.Unreadable.Count -gt 0){$script:ReadWarnings=$true}
            $usersByDomain[$domain]=$read.Users
            Write-SmartM365EvidenceLog -Message ("{0}: {1} user(s) through {2}." -f $domain,$read.Users.Count,$read.Service) -Level SUCCESS
        }
        catch{
            if($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]){throw}
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
    if($isPartial -or $script:ReadWarnings){$script:CompletionStatus='CompletedWithWarnings'}

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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBm1scr0Uh1F65F
# QNhYJAcdEt/U6hh2tghSb9pzs5JSJqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILJNo/w/4mdxT6vaae52imYvNloLgu4gSy6q7obSMeJfMA0GCSqG
# SIb3DQEBAQUABIIBgANkMBLu4ZoCMCUFB+GBa+XFfbc/QkZKXYpuVpqIzYC7Dm0V
# zS6OWDrbNbBsPbNvBOasSrNxSxLuev8xXlwVuIneKrn3O/OyDeg6ddi5VRe7q9Qn
# eG3ngfNSWRQOaCCM9yzDFUpZxgvbNpUWe+Nky7CnjsuCTUNuER7h5r/BnGnMsI7D
# lbDi5hMtfxWdoOhLNwjQCJp9dS8E4qAYxNDQBz7l2sYpgTxwibfpE40LTv5o1UN3
# qDkmeql85XXvk3+L3ePNKmI+wPQUugSN/W2fuat8sI3ZdoVd8oBLRe9jTwY8HYeb
# SNsacuoB4mY6GK3x4v+IKRN60JFrb6ljAL0Je6OjjJ+clYhEBpQHqBlXk3WzscnP
# we24f0ECeqzKZhFUng7nQlO89wT9HMgfAMPygSWJ5bDJMOmbq+ol+rrXMyTplLax
# Ncf08CHj+0Pd4T45yWD301vpMVPlfB4557/TGMFjhp1grHUFxhfs4RXamJOpflHl
# xt76JDfqA7AkdnN6IKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkwMDIx
# MjFaMC8GCSqGSIb3DQEJBDEiBCBtFJhbEscA3wRYQN26kogjmPZkhif85T7gS4Sy
# G6zNjDANBgkqhkiG9w0BAQEFAASCAgBNzlmHLBkn7tTY0Y1QChaiAMDOV92m1SG8
# H17XqSx9vuIXaaJooHi+7LuGrPr8fWjmm0bKQiL/IyoUio7p1TpZZMCa9b3/Occx
# e/eSvzF6DNw+ADciueqjKBTZXqXDbIGmtuKUC40RC3W4fB4LnIdFk681Z8qJE1CV
# fyAJfhrcRsjOANfxRymm61yAIQdarVIlqRAzNyvyHgRSBgEABUf6HsCBHkCrM4RJ
# Hu799IOv+JroLu00Ioi5v3m6ICCYLoQG0MxQLh+shEe2NGqeY3hSMAWvWxNM3zJh
# HdHcronfgoqp3n/zduH0e5IuxcBhwVyESx7pDM43UUdYhUaxfojiVk/ED9BMeP+7
# 6VN300DgY/yU40VxSv4ar0uNtjB3LgdfHcuq8hMFnutDqwKni/5SKrr5sJjmwQZQ
# 5OqieQRSOt/8Pf8G9AVoSzfGbT1aPcURN2USyyb1TFevQJwoWMBlbLrh4gurQyM+
# Q4vLV+4TgjNf6ogO0t/nb+9Dux2s+tMqOovGuHqcqxN6241876FXRCVOXqoug8Wf
# ld+r1mqP0wrhSBn0ZNTCljPgFt8Vtpw9525h3LJGi3KxZUr+S2LIapR4Hnhw+ax6
# ulm3qwKc01+wEz7SMIQXXEUaNPqDCVUtX6noARm5TjkzEuYgNBRb8pAdXIQGU9h/
# etFVoogKng==
# SIG # End signature block
