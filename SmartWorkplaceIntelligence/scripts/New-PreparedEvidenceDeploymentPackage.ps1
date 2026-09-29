[CmdletBinding()]
param([Parameter(Mandatory)][string]$OutputRoot)
# Build-only utility: no tenant initialization, installation or publication.
$ErrorActionPreference='Stop'
$product=Split-Path $PSScriptRoot -Parent
$repo=Split-Path $product -Parent
$names=@(
    'Invoke-PreparedEvidenceBuild.ps1','PreparedEvidencePipeline.psm1',
    'Test-PreparedEvidence.ps1','Test-PreparedEvidencePipeline.ps1','Test-PreparedEvidenceRules.ps1',
    'Test-PreparedEvidenceDeploymentPackage.ps1',
    'New-WorkforceIdentityEvidence.ps1','New-DeviceInventoryEvidence.ps1',
    'New-DeviceLifecycleExperienceTrendEvidence.ps1','New-ApplicationInventoryEvidence.ps1',
    'New-CollaborationEvidence.ps1','New-ContentStorageEvidence.ps1','New-MailboxEvidence.ps1',
    'New-LicensingEvidence.ps1','New-SecurityEvidence.ps1','New-BackupResilienceEvidence.ps1',
    'New-DataTrustEvidence.ps1','New-ExecutiveTrendEvidence.ps1'
)
$paths=@($names | ForEach-Object { 'SmartWorkplaceIntelligence/scripts/'+$_ })+@(
    'SmartWorkplaceIntelligence/config/prepared-evidence-contract.json',
    'SmartWorkplaceIntelligence/config/prepared-source-contract.json',
    'SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1',
    'SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.local.json.template',
    'SmartM365/SmartInventory/PreparedEvidence/README.md',
    'SmartM365/SmartInventory/PreparedEvidence/DEPLOYMENT.md',
    'SmartM365/SmartInventory/Launchers/Cloud/Test-SmartM365-WorkplaceEvidence-Prepare.cmd',
    'SmartM365/SmartInventory/Launchers/Cloud/Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd'
)
$stamp=[datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8)
$package=Join-Path ([IO.Path]::GetFullPath($OutputRoot)) ('WorkplaceEvidence-Prepare-'+$stamp)
New-Item -ItemType Directory -Path $package | Out-Null
$files=foreach($relative in $paths) {
    $source=Join-Path $repo $relative
    if (-not(Test-Path -LiteralPath $source -PathType Leaf)) {throw "Missing package source: $relative"}
    $destination=Join-Path $package $relative
    New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination
    $hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    if ((Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash -ne $hash) {throw "Copy mismatch: $relative"}
    [ordered]@{Path=$relative;Bytes=(Get-Item -LiteralPath $destination).Length;SHA256=$hash}
}
$jobs=Get-Content -LiteralPath (Join-Path $repo 'SmartM365/SmartInventory/Orchestrator/Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
$job=@($jobs.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')
if ($job.Count -ne 1 -or $job[0].Enabled) {throw 'Expected exactly one disabled preparation job.'}
$jobPath=Join-Path $package 'WorkplaceEvidence-Prepare.job.json'
@{Jobs=$job} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jobPath -Encoding utf8
$files+=@{Path='WorkplaceEvidence-Prepare.job.json';Bytes=(Get-Item $jobPath).Length;SHA256=(Get-FileHash $jobPath).Hash}
$required=@(
    'SmartM365/Config/SmartM365-TenantContext.ps1',
    'SmartM365/Modules/SmartM365.Core/SmartM365.Core.psd1',
    'SmartM365/Modules/SmartM365.Core/SmartM365.Core.psm1',
    'SmartM365/SmartInventory/Config/AccountClassification.psd1'
)
$prerequisites=foreach($relative in $required) {
    @{Path=$relative;ReferenceSHA256=(Get-FileHash -LiteralPath (Join-Path $repo $relative)).Hash}
}
@{SchemaVersion=1;CreatedUtc=[datetime]::UtcNow.ToString('o');Files=@($files);ExistingPrerequisites=@($prerequisites);AutomaticInstall=$false} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $package 'package-manifest.json') -Encoding utf8
& (Join-Path $PSScriptRoot 'Test-PreparedEvidenceDeploymentPackage.ps1') -PackageRoot $package
$zip=$package+'.zip'
Compress-Archive -Path (Join-Path $package '*') -DestinationPath $zip
[pscustomobject]@{Package=$package;Archive=$zip;Files=@($files).Count;ArchiveSHA256=(Get-FileHash $zip).Hash}
