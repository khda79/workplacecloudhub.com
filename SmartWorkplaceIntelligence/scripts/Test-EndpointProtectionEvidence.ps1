#Requires -Version 7.0
[CmdletBinding()]
param()
# Synthetic fixtures only. No tenant initialization, collectors, production files or Power BI.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'EndpointProtectionEvidence.psm1') -Force
$script:checks=0
function Assert-EP([bool]$Condition,[string]$Message) {$script:checks++; if (-not $Condition) {throw $Message}}
function Assert-EPThrows([scriptblock]$Action,[string]$Message) {
    $threw=$false; try {& $Action | Out-Null} catch {$threw=$true}; Assert-EP $threw $Message
}
function New-Requirement([string]$Product='Microsoft Defender Antivirus',[string]$Role='Antivirus',[string]$Mode='Active') {
    @{Role=$Role;Product=$Product;ExpectedMode=$Mode;EvidenceProductName='';Reason=''}
}
function New-PolicyRaw { @{Enabled=$true;PolicyVersion='synthetic-v1';MaxObservationAgeHours=48;Requirements=@((New-Requirement));ScopeRules=@()} }
function Normalize($Raw) { ConvertTo-EndpointProtectionPolicy -Configuration @{TenantKey='synthetic-test';EndpointProtectionPolicy=$Raw;Secret='must-not-leak'} -TenantKey 'synthetic-test' }
function New-Defender([string]$Id='one') { [pscustomobject]@{TenantKey='synthetic-test';DeviceId=$Id;CollectionStatus='Observed';IsPartialInventory='False';LastReportedDateTime='2026-01-02T00:00:00Z';MalwareProtectionEnabled='True';RealTimeProtectionEnabled='True';SignatureUpdateOverdue='False'} }
function New-System { [pscustomobject]@{TenantKey='synthetic-test';'Device Source ID'='one';'Observation UTC'='2026-01-02T00:00:00Z';'Expected AV Name'='Example AV';'Expected AV State'='Detected';'AV Qualification'='Evaluable registration observation';'AV Comparison'='No contradiction observed';'Prepared Inventory Match'='Matched';'Inventory Match'='Matched'} }
$inventory=@([pscustomobject]@{'Device Source ID'='one';Country='Synthetic country'})
$asOf=[datetimeoffset]'2026-01-03T00:00:00Z'
function Evaluate($Policy=(Normalize (New-PolicyRaw)),[object[]]$Defender=@((New-Defender)),[object[]]$System=@(),[object[]]$Inventory=$inventory) {
    return ,@(New-EndpointProtectionEvidence -Policy $Policy -InventoryRows $Inventory -DefenderRows $Defender -SystemRows $System -AsOfUtc $asOf)
}
$default=ConvertTo-EndpointProtectionPolicy -Configuration @{} -TenantKey 'synthetic-test'
Assert-EP (-not $default.Enabled -and $default.Requirements.Count -eq 0) 'Missing configuration must be disabled and neutral.'
Assert-EP ((Evaluate -Policy $default).Count -eq 0) 'Disabled policy must not produce evaluations.'
Assert-EP (@(New-EndpointProtectionPolicyRows $default).Count -eq 0) 'Disabled policy must not produce policy rows.'
$policy=Normalize (New-PolicyRaw)
Assert-EP (($policy | ConvertTo-Json -Depth 10) -notmatch 'must-not-leak|Secret') 'Tenant secrets leaked into normalized policy.'
$good=Evaluate
Assert-EP ($good.Count -eq 1 -and $good[0].'Evaluation State' -eq 'Healthy') 'Qualified active Defender checks should pass.'
Assert-EP ($good[0].'Endpoint Protection Key' -eq 'synthetic-test|one|antivirus|microsoft defender antivirus') 'Technical tenant-aware identity failed.'
Assert-EP ($good[0].Country -eq 'Synthetic country') 'Inventory country context lost.'
Assert-EP ($good[0].Freshness -eq 'Current') 'Observation freshness failed.'
foreach ($field in @('MalwareProtectionEnabled','RealTimeProtectionEnabled','SignatureUpdateOverdue')) {
    $d=New-Defender; $d.$field=if ($field -eq 'SignatureUpdateOverdue') {'True'} else {'False'}
    Assert-EP ((Evaluate -Defender @($d))[0].'Evaluation State' -eq 'Gap') "Explicit failure ignored: $field"
    foreach ($missing in @('','Invalid','Unknown',$null)) {
        $d=New-Defender; $d.$field=$missing
        Assert-EP ((Evaluate -Defender @($d))[0].'Evaluation State' -eq 'Unknown') "Missing/invalid $field incorrectly evaluated."
    }
}
foreach ($time in @('','invalid','2026-01-02T00:00:00','2025-12-01T00:00:00Z','2026-01-04T00:00:00Z')) {
    $d=New-Defender; $d.LastReportedDateTime=$time
    Assert-EP ((Evaluate -Defender @($d))[0].'Evaluation State' -eq 'Unknown') "Unqualified time evaluated: $time"
}
$d=New-Defender; $d.LastReportedDateTime='2026-01-02T02:00:00+02:00'
Assert-EP ((Evaluate -Defender @($d))[0].'Evaluation State' -eq 'Healthy') 'Explicit offset normalization failed.'
$raw=New-PolicyRaw; $raw.MaxObservationAgeHours=$null
Assert-EP ((Evaluate -Policy (Normalize $raw))[0].'Evaluation State' -eq 'Unknown') 'Unconfigured age policy evaluated.'
foreach ($partial in @('True','','Unknown')) {
    $d=New-Defender; $d.IsPartialInventory=$partial
    Assert-EP ((Evaluate -Defender @($d))[0].'Evaluation State' -eq 'Unknown') 'Partial or unknown collection evaluated.'
}
$d=New-Defender; $d.CollectionStatus='Error'
Assert-EP ((Evaluate -Defender @($d))[0].'Evaluation State' -eq 'Unknown') 'Errored collection evaluated.'
Assert-EP ((Evaluate -Defender @())[0].'Evaluation State' -eq 'Unknown') 'Missing Defender rows evaluated.'
Assert-EP ((Evaluate -Defender @((New-Defender),(New-Defender)))[0].'Evaluation State' -eq 'Unknown') 'Duplicate observations evaluated.'
$raw=New-PolicyRaw; $raw.Requirements=@()
Assert-EP ((Evaluate -Policy (Normalize $raw))[0].'Evaluation Reason' -eq 'Policy requirements not configured') 'Empty requirements treated as healthy.'
$raw=New-PolicyRaw; $raw.Requirements=@((New-Requirement 'Example AV'))
$third=Normalize $raw; $system=New-System
$result=Evaluate -Policy $third -System @($system)
Assert-EP ($result[0].'Registration State' -eq 'Detected' -and $result[0].'Evaluation State' -eq 'Unknown') 'AV registration claimed product health.'
$system.'Expected AV State'='NotDetected'
$result=Evaluate -Policy $third -System @($system)
Assert-EP ($result[0].'Registration State' -eq 'NotDetected' -and $result[0].'Evaluation State' -eq 'Unknown') 'Registration absence claimed an EDR gap.'
foreach ($field in @('AV Qualification','AV Comparison','Prepared Inventory Match','Inventory Match','Observation UTC')) {
    $s=New-System; $s.$field=if ($field -eq 'AV Comparison') {'Conflict'} else {''}
    Assert-EP ((Evaluate -Policy $third -System @($s))[0].'Registration State' -eq 'Unknown') "Unqualified registration exposed as detected: $field"
}
foreach ($collectionState in @('Fresh','Stale','NotCollected','Error','Unknown','Pending','NotApplicable','Failed','ScriptError','InvalidOutput','InvalidTimestamp','FutureObservation','')) {
    $s=New-System
    $s | Add-Member 'Local Data Source' 'IntuneRemediation'
    $s | Add-Member 'Local Data Freshness' $collectionState
    $result=Evaluate -Policy $third -System @($s)
    Assert-EP ($result[0].'Evaluation State' -eq 'Unknown') 'Recurring registration claimed product health.'
    Assert-EP ($result[0].'Registration State' -eq $(if($collectionState -eq 'Fresh'){'Detected'}else{'Unknown'})) 'Unqualified recurring collection exposed a positive registration.'
    if ($collectionState -in @('Pending','NotApplicable','Failed','ScriptError','InvalidOutput','InvalidTimestamp','FutureObservation')) {
        Assert-EP ($result[0].Freshness -eq $collectionState) 'Endpoint protection lost the local collection qualifier.'
    }
}
$s=New-System; $s.'Expected AV Name'='Legacy display'
Assert-EP ((Evaluate -Policy $third -System @($s))[0].'Registration State' -eq 'Unknown') 'Implicit provider alias guessed.'
$raw.Requirements[0].EvidenceProductName='Legacy display'
Assert-EP ((Evaluate -Policy (Normalize $raw) -System @($s))[0].'Registration State' -eq 'Detected') 'Explicit provider alias failed.'
$raw=New-PolicyRaw; $raw.Requirements=@((New-Requirement 'Example AV' 'EDR'))
Assert-EP ((Evaluate -Policy (Normalize $raw) -System @((New-System)))[0].'Evaluation State' -eq 'Unknown') 'EDR health inferred from AV registration.'
$raw.Requirements=@((New-Requirement 'Microsoft Defender for Endpoint' 'EDR'))
Assert-EP ((Evaluate -Policy (Normalize $raw))[0].'Evaluation State' -eq 'Unknown') 'Defender Antivirus flags inferred MDE health.'
$raw.Requirements=@((New-Requirement 'Microsoft Defender Antivirus' 'Antivirus' 'Passive'))
Assert-EP ((Evaluate -Policy (Normalize $raw))[0].'Evaluation State' -eq 'Unknown') 'Defender passive mode inferred from active checks.'
$raw.Requirements=@((New-Requirement 'Example AV' 'Antivirus' 'NotRequired')); $raw.Requirements[0].Reason='Synthetic reviewed exception'
Assert-EP ((Evaluate -Policy (Normalize $raw))[0].'Evaluation State' -eq 'NotApplicable') 'Documented exception not respected.'
$raw=New-PolicyRaw; $raw.ScopeRules=@(@{ScopeId='third-party';DeviceSourceIds=@('ONE');Requirements=@((New-Requirement 'Example AV'))})
$mixed=Evaluate -Policy (Normalize $raw) -System @((New-System)) -Inventory @($inventory[0],[pscustomobject]@{'Device Source ID'='two'}) -Defender @((New-Defender 'one'),(New-Defender 'two'))
Assert-EP ($mixed.Count -eq 2 -and $mixed[0].'Expected Product' -eq 'Example AV' -and $mixed[1].'Evaluation State' -eq 'Healthy') 'Scope replacement or default mixed policy failed.'
$raw.ScopeRules+=@(@{ScopeId='overlap';DeviceSourceIds=@('one');Requirements=@((New-Requirement))})
Assert-EP ((Evaluate -Policy (Normalize $raw))[0].'Evaluation Reason' -eq 'Overlapping policy scopes') 'Overlapping scopes silently chose a winner.'
$raw=New-PolicyRaw; $raw.ScopeRules=@(@{ScopeId='empty-scope';DeviceSourceIds=@('one');Requirements=@()})
Assert-EP ((Evaluate -Policy (Normalize $raw))[0].'Evaluation State' -eq 'Unknown') 'Empty scope requirements inherited defaults or created an exemption.'
$raw.ScopeRules+=@(@{ScopeId='EMPTY-SCOPE';DeviceSourceIds=@('two');Requirements=@()})
Assert-EPThrows {Normalize $raw} 'Case-insensitive duplicate scope ID accepted.'
$raw=New-PolicyRaw; $raw.Requirements+=@((New-Requirement 'Example EDR' 'EDR'))
$dual=Evaluate -Policy (Normalize $raw)
Assert-EP ($dual.Count -eq 2 -and @($dual.'Endpoint Protection Key' | Sort-Object -Unique).Count -eq 2) 'AV and EDR identities collapsed into one device observation.'
foreach ($bad in @(@{Enabled='false';PolicyVersion='';MaxObservationAgeHours=$null;Requirements=@();ScopeRules=@()}, @{Enabled=$true;PolicyVersion='';MaxObservationAgeHours=48;Requirements=@();ScopeRules=@()})) {
    Assert-EPThrows {Normalize $bad} 'Malformed enable/version accepted.'
}
foreach ($age in @(-1,0,'48',[double]::PositiveInfinity,[double]::NaN,$true)) {
    $raw=New-PolicyRaw; $raw.MaxObservationAgeHours=$age; Assert-EPThrows {Normalize $raw} 'Invalid age policy accepted.'
}
foreach ($badRequirement in @((New-Requirement 'Example AV' 'EDR' 'Passive'),(New-Requirement 'Example AV' 'Antivirus' 'NotRequired'),(New-Requirement 'bad|key'),(New-Requirement 'Example AV' 'Other'))) {
    $raw=New-PolicyRaw; $raw.Requirements=@($badRequirement); Assert-EPThrows {Normalize $raw} 'Malformed requirement accepted.'
}
$raw=New-PolicyRaw; $raw.Requirements+=@((New-Requirement)); Assert-EPThrows {Normalize $raw} 'Duplicate requirement accepted.'
$raw=New-PolicyRaw; $raw.Requirements+=@((New-Requirement 'Example AV')); Assert-EPThrows {Normalize $raw} 'Multiple active antivirus requirements accepted.'
$raw=New-PolicyRaw; $raw.ScopeRules=@(@{ScopeId='Default';DeviceSourceIds=@('one');Requirements=@()}); Assert-EPThrows {Normalize $raw} 'Reserved scope ID accepted.'
$raw.ScopeRules=@(@{ScopeId='scope';DeviceSourceIds=@();Requirements=@()}); Assert-EPThrows {Normalize $raw} 'Empty scope IDs accepted.'
$raw.ScopeRules=@(@{ScopeId='scope';DeviceSourceIds=@('one','ONE');Requirements=@()}); Assert-EPThrows {Normalize $raw} 'Duplicate scope device IDs accepted.'
Assert-EPThrows {ConvertTo-EndpointProtectionPolicy -Configuration @{TenantKey='other-test'} -TenantKey 'synthetic-test'} 'Tenant policy mismatch accepted.'
$d=New-Defender; $d.TenantKey='other-test'; Assert-EPThrows {Evaluate -Defender @($d)} 'Cross-tenant observation accepted.'
$d.TenantKey=''; Assert-EPThrows {Evaluate -Defender @($d)} 'Tenantless raw observation accepted.'
Assert-EPThrows {Evaluate -Inventory @($inventory[0],$inventory[0])} 'Duplicate inventory IDs accepted.'
Assert-EPThrows {Evaluate -Inventory @([pscustomobject]@{'Device Source ID'='';'Device Name'='one'})} 'Device name fallback accepted.'
Assert-EPThrows {Evaluate -Inventory @([pscustomobject]@{TenantKey='other-test';'Device Source ID'='one'})} 'Cross-tenant prepared inventory accepted.'
$productRoot=Split-Path -Parent $PSScriptRoot
$repository=Split-Path -Parent $productRoot
$template=Get-Content -LiteralPath (Join-Path $repository 'SmartM365/Config/Tenants/tenant.local.json.template') -Raw | ConvertFrom-Json
$templatePolicy=ConvertTo-EndpointProtectionPolicy -Configuration $template -TenantKey $template.TenantKey
Assert-EP (-not $templatePolicy.Enabled -and $templatePolicy.Requirements.Count -eq 0 -and $templatePolicy.ScopeRules.Count -eq 0) 'Public template activates a customer-specific policy.'
$base=Get-Content -LiteralPath (Join-Path $productRoot 'config/prepared-evidence-contract.json.txt') -Raw | ConvertFrom-Json
$extension=Get-Content -LiteralPath (Join-Path $productRoot 'config/endpoint-protection-contract.json.txt') -Raw | ConvertFrom-Json
$baseSnapshot=$base | ConvertTo-Json -Depth 20 -Compress
$merged=Merge-EndpointProtectionContract $base $extension
Assert-EP ($merged.tables.Count -eq $base.tables.Count+2 -and ($base | ConvertTo-Json -Depth 20 -Compress) -eq $baseSnapshot) 'Opt-in contract changed base tables.'
Assert-EPThrows {Merge-EndpointProtectionContract $merged $extension} 'Duplicate contract extension accepted.'
Assert-EP ((($extension.tables[0].columns.name) -join '|') -eq ((Get-EndpointProtectionPolicyColumns) -join '|')) 'Policy contract/header mismatch.'
Assert-EP ((($extension.tables[1].columns.name) -join '|') -eq ((Get-EndpointProtectionEvidenceColumns) -join '|')) 'Evidence contract/header mismatch.'
Assert-EP (($good[0].PSObject.Properties.Name -join '|') -eq ((Get-EndpointProtectionEvidenceColumns) -join '|')) 'Evaluation object/header mismatch.'
Assert-EP (((@(New-EndpointProtectionPolicyRows $policy))[0].PSObject.Properties.Name -join '|') -eq ((Get-EndpointProtectionPolicyColumns) -join '|')) 'Policy object/header mismatch.'
# Isolated synthetic integration. Scratch retained for inspection; no recursive deletion.
$scratch=Join-Path ([IO.Path]::GetTempPath()) ('EndpointProtectionOffline-'+[guid]::NewGuid().ToString('N'))
$source=Join-Path $scratch 'source'; $stage=Join-Path $scratch 'stage'
New-Item -ItemType Directory -Path (Join-Path $source 'DATA-LAST'),$stage -Force | Out-Null
$policyPath=Join-Path $scratch 'policy.json.txt'; $inventoryPath=Join-Path $stage 'DeviceInventoryEvidence.csv'; $systemPath=Join-Path $stage 'DeviceSystemEvidence.csv'
$policy | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $policyPath -Encoding utf8
$inventory | Export-Csv -LiteralPath $inventoryPath -NoTypeInformation
New-Defender | Export-Csv -LiteralPath (Join-Path $source 'DATA-LAST/Intune_EndpointSecurity_DefenderAgents.csv') -NoTypeInformation
New-System | Export-Csv -LiteralPath $systemPath -NoTypeInformation
$inputHash=(Get-FileHash -LiteralPath $inventoryPath).Hash
& (Join-Path $PSScriptRoot 'New-EndpointProtectionEvidence.ps1') -DataRoot $source -InventoryEvidencePath $inventoryPath -SystemEvidencePath $systemPath -PolicyPath $policyPath -OutputRoot $stage -AsOfUtc $asOf
$csv=@(Import-Csv -LiteralPath (Join-Path $stage 'EndpointProtectionEvidence.csv'))
Assert-EP ($csv.Count -eq 1 -and $csv[0].'Evaluation State' -eq 'Healthy') 'Synthetic generator did not create qualified CSV.'
Assert-EP ((Get-FileHash -LiteralPath $inventoryPath).Hash -eq $inputHash) 'Synthetic source was modified.'
$validationJson=& (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root $stage -ContractPath (Join-Path $productRoot 'config/endpoint-protection-contract.json.txt')
Assert-EP ($LASTEXITCODE -eq 0 -and ($validationJson | ConvertFrom-Json).Passed) 'Generated CSVs failed the real prepared-evidence validator.'
# Exercise the existing publisher against these two synthetic tables only.
# This isolated temporary DATA-POWERBI is not the tenant's DATA-POWERBI.
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
$syntheticOutput=Join-Path $scratch 'DATA-POWERBI'
$syntheticBatch=Publish-PreparedEvidenceBatch -StagingRoot $stage -OutputRoot $syntheticOutput -TenantKey 'synthetic-test' -ContractPath (Join-Path $productRoot 'config/endpoint-protection-contract.json.txt') -Provenance @{Mode='SyntheticEndpointProtectionTest'} -SkipRetention
$manifest=Get-Content -LiteralPath (Join-Path $syntheticBatch.BatchPath 'batch.json.txt') -Raw | ConvertFrom-Json
Assert-EP ($manifest.TenantKey -eq 'synthetic-test' -and $manifest.Files.Count -eq 2) 'Synthetic endpoint contract publication failed.'
foreach ($file in $manifest.Files) {
    Assert-EP ((Get-FileHash -LiteralPath (Join-Path $syntheticBatch.BatchPath $file.File)).Hash -eq $file.SHA256) 'Synthetic manifest hash mismatch.'
    Assert-EP (@(Import-Csv -LiteralPath (Join-Path $syntheticBatch.BatchPath $file.File) | Where-Object TenantKey -ne 'synthetic-test').Count -eq 0) 'Synthetic batch lost tenant identity.'
}
Assert-EPThrows {& (Join-Path $PSScriptRoot 'New-EndpointProtectionEvidence.ps1') -DataRoot $source -InventoryEvidencePath $inventoryPath -SystemEvidencePath $systemPath -PolicyPath $policyPath -OutputRoot (Join-Path $source 'output')} 'Generator wrote below raw source.'
$beforeDisabled=(Get-FileHash -LiteralPath (Join-Path $stage 'EndpointProtectionEvidence.csv')).Hash
$policy.Enabled=$false; $policy | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $policyPath
Assert-EPThrows {& (Join-Path $PSScriptRoot 'New-EndpointProtectionEvidence.ps1') -DataRoot $source -InventoryEvidencePath $inventoryPath -SystemEvidencePath $systemPath -PolicyPath $policyPath -OutputRoot $stage} 'Generator accepted disabled policy.'
Assert-EP ((Get-FileHash -LiteralPath (Join-Path $stage 'EndpointProtectionEvidence.csv')).Hash -eq $beforeDisabled) 'Disabled generator changed staged evidence.'
$policy.Enabled=$true; $policy.MaxObservationAgeHours=48.5
$policy | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $policyPath
# Header-only inventory is a valid empty population, not a protected population.
[pscustomobject]@{'Device Source ID'='';Country=''} | ConvertTo-Csv -NoTypeInformation | Select-Object -First 1 | Set-Content -LiteralPath $inventoryPath
& (Join-Path $PSScriptRoot 'New-EndpointProtectionEvidence.ps1') -DataRoot $source -InventoryEvidencePath $inventoryPath -SystemEvidencePath $systemPath -PolicyPath $policyPath -OutputRoot $stage -AsOfUtc $asOf
Assert-EP (@(Import-Csv -LiteralPath (Join-Path $stage 'EndpointProtectionEvidence.csv')).Count -eq 0) 'Empty inventory invented a protected device.'
Assert-EP ((Get-Content -LiteralPath (Join-Path $stage 'EndpointProtectionEvidence.csv')).StartsWith('"TenantKey","Endpoint Protection Key"')) 'Empty CSV lost its schema.'
Assert-EP ((Import-Csv -LiteralPath (Join-Path $stage 'EndpointProtectionPolicy.csv')).'Max Observation Age Hours' -eq '48.5') 'Numeric policy serialization is not invariant.'
$raw=New-PolicyRaw; $raw.ScopeRules=@(@{ScopeId='not-in-inventory';DeviceSourceIds=@('absent');Requirements=@((New-Requirement))})
Normalize $raw | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $policyPath
$warnings=@()
& (Join-Path $PSScriptRoot 'New-EndpointProtectionEvidence.ps1') -DataRoot $source -InventoryEvidencePath $inventoryPath -SystemEvidencePath $systemPath -PolicyPath $policyPath -OutputRoot $stage -AsOfUtc $asOf -WarningVariable warnings -WarningAction SilentlyContinue
Assert-EP ($warnings.Count -eq 1 -and [string]$warnings[0] -match 'absent from this inventory') 'Unmatched scope IDs were silently ignored.'
$unmatchedPolicy=@(Import-Csv -LiteralPath (Join-Path $stage 'EndpointProtectionPolicy.csv') | Where-Object 'Scope ID' -eq 'not-in-inventory')
Assert-EP ($unmatchedPolicy.Count -eq 1 -and $unmatchedPolicy[0].'Device Source IDs' -eq '["absent"]') 'Unmatched scope IDs not preserved for review.'
$pipeline=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Raw
$build=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Invoke-PreparedEvidenceBuild.ps1') -Raw
Assert-EP ($pipeline -match 'if \(\$endpointPolicy.Enabled\)' -and $pipeline -match '@endpointContractArgs') 'Opt-in publication contract not wired.'
Assert-EP ($pipeline -match 'ConvertTo-EndpointProtectionPolicy -Configuration \$AdCountryConfiguration') 'Effective tenant configuration not wired.'
Assert-EP ($build -match 'if \(\$EndpointProtectionPolicyPath\)' -and $build -match "Add-Task 'EndpointProtection' 'Devices' 'New-EndpointProtectionEvidence.ps1'") 'Build generator not wired.'
Assert-EP ($pipeline -match "'EndpointProtectionPolicy.json.txt','PreparedEndpointContract.json.txt'") 'Run-local endpoint policy cleanup not wired.'
Write-Host "PASS: $script:checks endpoint protection checks. Synthetic scratch: $scratch"

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAk4gC+oWey0RWj
# Zxdqof1ZxaxH0QSsB6ceynHrGi/J76CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICdLwx5fj+6OD7GwSO8VNds2QKD9za16VXZo7/+SlCGrMA0GCSqG
# SIb3DQEBAQUABIIBgCTK+T5AArw80WInEK3DDL4E4WidsKnZEzA6Q9gwwnMbuZ3W
# 1NVDirzawNA3jdaCxFEtr+GESxmoj5nf3LhTV/JfDxq+0Bq7kin70OewiSOR4PwD
# PRUf5J7BK7Z5vR7jO0zOcgR311uwncg3Y0XR629FwKdVyXjH4XE3CNMOn7b/p67H
# q4Gc3SVuttZ6wAKHi6pTb66OzENca1z9E10eKc1GzpvbRjzjPb4zzeUheGKpLCjj
# ycA5CB7966YiqFqkwnIVMBdC+OQokQFKDMM2WwPwamGtJcXTfVMYL7A1NPFe32AH
# YSCS9V+DSGS36BhIXHF/c96ctQeXfNJ/A27hc0NXMxxBRBtRpnwGTVh/ZUZty7HZ
# 1CNhGGYgkdc6oic7bLmFuQ2v5j7paN0TJKn7RHiSTVWpJN9cxkSq4gJmj3BmqquB
# dpox155sx2yGirvBiBxanYs0Q3CQRgFTj78FIzG8GVfgdJfNhfq8JIa7ma51sUmL
# q27YawYREIG67N0vE6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMzE2
# MzNaMC8GCSqGSIb3DQEJBDEiBCCLUyEHf9j0gPPOlQKAiwHO9vTtBIdnh1wJnQfi
# 0GNPnzANBgkqhkiG9w0BAQEFAASCAgBOALWORJU8Z9GkuL3LQtYGqVKrRePWRBPc
# 9sEL7yGo04Wc2hS/y/IARNouEB9MCf9yUZh8mW0zLqq+pEPw9bf1PgC95X7VBObP
# ZVt8w0IXY52HtcG7FV5iii00jZ7CgwgmuypXRgrz7WPAFS8Q7gCkolLA5V3B58wN
# p7CblYrjfB8U7wx8cpOEtqzNbdTHvC/dSXrGC9nD2nNdBixHdofQKEeAoEe/hBcU
# cpMV9UvyEg12/U6i3iYUX6KrVRTQKn5KUEktH4tiAXsdHAiwYE4kFTS0SSjT/IGX
# SqgLddcuozlUp84WXufTQFqdCSdDg7BYwBD1/OQAXRgdlSCSbQdyTV9B63SYZmDc
# ueXIh59JrUOfVKoSTANFmBQwfXYL5ORD+vQHKj99ixEcMs8YZhCKFUIXav7hnHdX
# ZdsDAqIgbddVNS/ZHG2wJoa5dAuadE7aa9J7v/JQRBQ2py2pi1dPRj2YuUEgeMkt
# NbyzZBa16qT7rsHObu3ReEaXD8+dFTArjmAccHgBKOjgIuk4kW48dli+M1dG4+Km
# IMYM83lyoQJBw6ZkrIKoqBL51w/CnhhXpc6OFAl6H6RqXvc2psk0fS2Eo6yywm+v
# 94fsMcZupCqAM9+hogPHfzx4UguVa88x/qmMp891YSOPMAYeKMf3zQjvQYeb9MBQ
# NLJUWOKauw==
# SIG # End signature block
