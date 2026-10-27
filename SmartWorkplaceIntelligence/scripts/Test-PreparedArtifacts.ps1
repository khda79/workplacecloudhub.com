# Version 1.1.0
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'PreparedArtifacts.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('WorkplaceArtifactsTest-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null
$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
function Reject([scriptblock]$Action,[string]$Message){$failed=$false;try{& $Action}catch{$failed=$true};Check $failed $Message}
try {
    $source=Join-Path $root 'source';$dest=Join-Path $root 'batch'
    New-Item -ItemType Directory -Path (Join-Path $source 'DATA-LAST'),$dest | Out-Null
    $spec=Get-Content (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw|ConvertFrom-Json
    foreach($file in $spec.currentFiles){
        $path=Join-Path $source "DATA-LAST/$file"
        if($file -like '*.csv'){@{TenantKey='synthetic';SnapshotId=('a'*32)}|Export-Csv -LiteralPath $path -NoTypeInformation}
        else{@{TenantKey='synthetic';SnapshotId=('a'*32);SchemaVersion=2;Products=@(1,2,3,4);OtherProducts=@(1,2,3);AssignmentIssues=@{Available=$false;Rows=@()}}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $path}
    }
    foreach($history in $spec.history){
        $dir=Join-Path $source "$($history.root)/2026-W40";New-Item -ItemType Directory -Path $dir -Force|Out-Null
        @{TenantKey='synthetic';Id='one'}|Export-Csv -LiteralPath (Join-Path $dir $history.file) -NoTypeInformation
        @{Week='2026-W40';SnapshotCreatedAtUtc='2026-10-01T00:00:00Z'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'manifest.json')
    }
    $plan=@(Get-PreparedArtifactSources -DataRoot $source)
    Check ($plan.Count -eq ($spec.currentFiles.Count+4+@($spec.optionalCurrentFiles).Count)) 'Incomplete capture plan.'
    Check (@($plan | Where-Object { $_.ContainsKey('IsOptional') -and $_.IsOptional }).Count -eq @($spec.optionalCurrentFiles).Count) 'Optional artifact files are not optional capture items.'
    $entries=@(Copy-PreparedArtifacts -SnapshotRoot $source -Destination $dest -TenantKey 'synthetic')
    Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $entries
    Check ($entries.Count -eq ($spec.currentFiles.Count+5)) 'Incomplete artifact manifest.'
    foreach($entry in $entries|Where-Object SourceRelative -NE 'GeneratedFromVerifiedCapture'){
        Check ((Get-FileHash (Join-Path $source $entry.SourceRelative)).Hash -eq $entry.SHA256) 'Capture bytes changed.'
    }
    Check (-not @($entries | Where-Object File -Like 'Source-Exchange_OnPrem_Servers_*').Count) 'Absent optional infrastructure was fabricated.'
    $inventory=Join-Path $source 'DATA-LAST/Exchange_OnPrem_Servers_Inventory.csv'
    $compute=Join-Path $source 'DATA-LAST/Exchange_OnPrem_Servers_Compute.csv'
    '"TenantKey";"Name";"Description"', '"synthetic";"mail-one";"quoted; semicolon"' | Set-Content -LiteralPath $inventory
    '"TenantKey";"ExchangeServerName";"MemoryGB"', '"synthetic";"mail-one";"64"' | Set-Content -LiteralPath $compute
    $infraDest=Join-Path $root 'infrastructure';New-Item -ItemType Directory -Path $infraDest|Out-Null
    $infraEntries=@(Copy-PreparedArtifacts -SnapshotRoot $source -Destination $infraDest -TenantKey 'synthetic')
    Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $infraDest -Entries $infraEntries
    Check (@($infraEntries | Where-Object File -Like 'Source-Exchange_OnPrem_Servers_*').Count -eq 2) 'Semicolon infrastructure was not captured.'
    Check ($infraEntries.Count -eq $entries.Count+2) 'Optional infrastructure changed required membership.'
    $goodCompute=Get-Content -LiteralPath $compute -Raw
    foreach($bad in @(
        ('"TenantKey";"ExchangeServerName"'+"`n"+'"foreign";"mail-one"'),
        ('"TenantKey";"ExchangeServerName"'+"`n"+'"";"mail-one"'),
        ('"OtherKey";"ExchangeServerName"'+"`n"+'"synthetic";"mail-one"'),
        ('"TenantKey";"TenantKey"'+"`n"+'"synthetic";"synthetic"'),
        ('"TenantKey";"ExchangeServerName"'+"`n"+'"synthetic";"mail-one";"extra"')
    )){
        Set-Content -LiteralPath $compute -Value $bad
        $badDest=Join-Path $root ([guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $badDest|Out-Null
        Reject {Copy-PreparedArtifacts -SnapshotRoot $source -Destination $badDest -TenantKey 'synthetic'} 'Invalid optional infrastructure tenant/schema accepted.'
        Check (-not (Test-Path -LiteralPath (Join-Path $badDest 'Source-Exchange_OnPrem_Servers_Compute.csv'))) 'Rejected optional infrastructure was copied.'
    }
    Set-Content -LiteralPath $compute -Value $goodCompute
    $copied=Join-Path $infraDest 'Source-Exchange_OnPrem_Servers_Compute.csv'
    '"TenantKey";"ExchangeServerName"', '"foreign";"mail-one"' | Set-Content -LiteralPath $copied
    $copiedEntry=@($infraEntries | Where-Object File -CEQ 'Source-Exchange_OnPrem_Servers_Compute.csv')[0]
    $copiedEntry.Bytes=(Get-Item -LiteralPath $copied).Length
    $copiedEntry.SHA256=(Get-FileHash -LiteralPath $copied).Hash
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $infraDest -Entries $infraEntries} 'Foreign optional tenant accepted despite matching size/hash.'
    Check (-not (Test-PreparedArtifactName '../current.json.txt')) 'Traversal accepted.'
    Check (-not (Test-PreparedArtifactName 'Source-secret.json.txt')) 'Unknown artifact accepted.'
    Check (-not (Test-PreparedArtifactName 'History-Windows-2026-W54-manifest.json.txt')) 'Invalid ISO week accepted.'
    Reject {Copy-PreparedArtifacts -SnapshotRoot $source -Destination $dest -TenantKey 'synthetic'} 'Existing immutable artifacts overwritten.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries @($entries+$entries[0])} 'Duplicate entry accepted.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries @($entries|Where-Object File -NE 'PreparedHistoryIndex.json.txt')} 'Missing index accepted.'
    # Retired artifacts are accepted only for explicitly reviewed historical contract hashes.
    $legacyName='Source-Exchange_HybridIdentity_Issues.csv'
    $oldHash=$spec.legacyCurrentFiles[0].contractSHA256[-1]
    Check (-not (Test-PreparedArtifactName $legacyName)) 'Retired artifact accepted for a new batch.'
    Check (Test-PreparedArtifactName $legacyName -ContractSHA256 $oldHash) 'Reviewed historical artifact refused.'
    $legacyPath=Join-Path $dest $legacyName
    [pscustomobject]@{TenantKey='synthetic';ObjectGUID='contoso-account'} | Export-Csv -LiteralPath $legacyPath -NoTypeInformation
    $legacyEntry=[pscustomobject]@{File=$legacyName;Bytes=(Get-Item $legacyPath).Length;SHA256=(Get-FileHash $legacyPath).Hash}
    $legacyEntries=@($entries)+@($legacyEntry)
    Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $legacyEntries -ContractSHA256 $oldHash
    Check $true 'Historical artifacts accepted with their intact hashes.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $entries -ContractSHA256 $oldHash} 'Historical required artifact silently lost.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $legacyEntries} 'Stale legacy file admitted to a new batch.'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $legacyEntries -ContractSHA256 ('0'*64)} 'Unknown historical contract accepted.'
    Add-Content -LiteralPath $legacyPath 'tampered'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $legacyEntries -ContractSHA256 $oldHash} 'Tampered historical artifact accepted.'
    $changed=Join-Path $dest $entries[0].File;Add-Content -LiteralPath $changed 'tampered'
    Reject {Assert-PreparedArtifacts -TenantKey 'synthetic' -BatchRoot $dest -Entries $entries} 'Changed bytes accepted.'
    $dest2=Join-Path $root 'batch2';New-Item -ItemType Directory -Path $dest2|Out-Null
    Reject {Copy-PreparedArtifacts -SnapshotRoot $source -Destination $dest2 -TenantKey 'foreign'} 'Wrong tenant accepted.'
    Check (@(Get-ChildItem $dest2 -File).Count -eq 0) 'Identity failure wrote output.'
    'Prepared artifact offline checks passed: '+$checks
} finally {
    $resolved=[IO.Path]::GetFullPath($root)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -match '^WorkplaceArtifactsTest-[a-f0-9]{32}$'){
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
