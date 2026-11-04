<#
.SYNOPSIS
    Test the PowerShell 7 to Windows PowerShell identity bridge with synthetic data.
.VERSION
    1.0.2
#>
#Requires -Version 7.4
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath $PSScriptRoot).ProviderPath
$scratch=Join-Path $root ('.identity-bridge-test-'+[guid]::NewGuid().ToString('N'))
if(-not $scratch.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe test path.'}
$oldPath=$env:PSModulePath
$bridge=$null
try {
    $module=Join-Path $scratch 'PnP.PowerShell'
    New-Item -ItemType Directory -Path $module -Force|Out-Null
    @'
@{RootModule='PnP.PowerShell.psm1';ModuleVersion='99.0.0';GUID='a481ac4f-9521-4708-a0b6-812b24fab22b';FunctionsToExport='*'}
'@ | Set-Content -LiteralPath (Join-Path $module 'PnP.PowerShell.psd1') -Encoding utf8
    @'
function Get-PnPList { [CmdletBinding()]param($Identity,$Includes,$Connection) [pscustomobject]@{Title='Documents';BaseType='DocumentLibrary';Id=[guid]'c285d922-b71a-4b41-b546-76a25c44a173';RootFolder=[pscustomobject]@{ServerRelativeUrl='/sites/source/Shared Documents'}} }
function Get-PnPFile { [CmdletBinding()]param($Url,[switch]$AsListItem,$Connection) @{Id=8;FileRef='/sites/source/Shared Documents/Folder/file.txt';ParentList=[pscustomobject]@{Id=[guid]'c285d922-b71a-4b41-b546-76a25c44a173'}} }
function Get-PnPProperty { [CmdletBinding()]param($ClientObject,$Property,$Connection) }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath (Join-Path $module 'PnP.PowerShell.psm1') -Encoding utf8
    $env:PSModulePath=$scratch+[IO.Path]::PathSeparator+$oldPath
    . (Join-Path $root '../Scripts/Diagnostics/SmartM365-SharePointMigration-SpoIdentityBridge.ps1')
    $bridge=Start-SmartM365SpoIdentityBridge -Connection ([pscustomobject]@{Url='https://tenant.sharepoint.com/sites/source'})
    $client=Join-Path $scratch 'client.ps1'
    @'
param($Helper,$PipeName)
$ErrorActionPreference='Stop'
. $Helper
$args=@{PipeName=$PipeName;SiteUrl='https://tenant.sharepoint.com/sites/source';ListName='Documents';SourceFilePath='Folder/file.txt';ExpectedItemId=8}
$result=Get-SmartM365BridgedSpoSourceIdentity @args
if($result.SourceItemId -ne 8){throw 'Valid identity was rejected.'}
$bad=$args.Clone();$bad.ExpectedItemId=9
try {Get-SmartM365BridgedSpoSourceIdentity @bad|Out-Null;throw 'Wrong ID accepted.'} catch {if($_.Exception.Message -notmatch 'verification failed'){throw}}
$bad=$args.Clone();$bad.SiteUrl='https://other.sharepoint.com/sites/source'
try {Get-SmartM365BridgedSpoSourceIdentity @bad|Out-Null;throw 'Wrong site accepted.'} catch {if($_.Exception.Message -notmatch 'verification failed'){throw}}
Write-Output 'PASS: bridge accepts exact identity, rejects incorrect ID and site, and reuses its connection.'
'@ | Set-Content -LiteralPath $client -Encoding utf8
    $helper=Join-Path $root '../Scripts/Diagnostics/SmartM365-SharePointMigration-SpoIdentityBridgeClient.ps1'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $client -Helper $helper -PipeName $bridge.PipeName
    if($LASTEXITCODE -ne 0){throw 'Windows PowerShell bridge client failed.'}
    $sharegate=Join-Path $scratch 'ShareGate'
    New-Item -ItemType Directory -Path $sharegate -Force|Out-Null
    @'
@{RootModule='ShareGate.psm1';ModuleVersion='99.0.0';GUID='b4be4596-bf51-4d18-83c0-d022d6995b19';FunctionsToExport='*'}
'@ | Set-Content -LiteralPath (Join-Path $sharegate 'ShareGate.psd1') -Encoding utf8
    @'
function Connect-Site { [CmdletBinding()]param([string]$Url,[switch]$Browser,[switch]$ModernAuth) [pscustomobject]@{Address=$Url} }
function Get-List { [CmdletBinding()]param($Site,[string]$Name) [pscustomobject]@{Title=$Name;Site=$Site} }
function Get-Folder { [CmdletBinding()]param($List,[string]$Path) [pscustomobject]@{Path=$Path} }
function Get-File { [CmdletBinding()]param($List,[string]$Path) if($Path -eq 'file.txt'){throw 'not found'};[pscustomobject]@{Address=($List.Site.Address+'/'+$List.Title+'/'+$Path)} }
function New-CopySettings { [CmdletBinding()]param([string]$OnContentItemExists) [pscustomobject]@{Mode=$OnContentItemExists} }
function Copy-Content { [CmdletBinding()]param($SourceList,$DestinationList,$SourceItemId,$DestinationFolder,$CopySettings,$TaskName) Set-Content -LiteralPath $env:SMARTM365_BRIDGE_TEST_COPY -Value 'called';throw 'COPY_REACHED' }
function Export-Report { [CmdletBinding()]param($CopyResult,[string]$Path) }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath (Join-Path $sharegate 'ShareGate.psm1') -Encoding utf8
    $integration=Join-Path $scratch 'integration.ps1'
    @'
param($Qualification,$Scratch,$PipeName)
$ErrorActionPreference='Stop'
$zone=@([TimeZoneInfo]::GetSystemTimeZones()|Where-Object {$hour=[TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow,$_).Hour;$hour -ge 10 -and $hour -le 14}|Select-Object -First 1)[0]
if(-not $zone){throw 'No safe test timezone.'}
foreach($id in @(8,9)){
 $project=Join-Path $Scratch ('project-'+$id);$analysis=Join-Path $project 'ShareGate\Diagnostics\Analysis'
 New-Item -ItemType Directory -Path $analysis -Force|Out-Null
 $row=[pscustomobject]@{SessionId='261103-1';RuleId='SG-ACCESS-SOURCE';State='To fix';AccessSide='Source';ObjectType='File';ItemKey=('file-'+$id);SourceItemId=$id;SourceUrl='https://tenant.sharepoint.com/sites/source';SourceList='Documents';DestinationUrl='https://tenant.sharepoint.com/sites/target';DestinationList='Documents';'Raw: Source path'='Folder/file.txt';'Raw: Destination path'='Folder/file.txt'}
 $csv=Join-Path $analysis 'ClassifiedRows.csv';$row|Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8
 $hash=(Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash
 $env:SMARTM365_BRIDGE_TEST_COPY=Join-Path $project 'copy-called.txt';$env:SMARTM365_BRIDGE_TEST_ID=[string]$id
 function global:Read-Host { param($Prompt) 'QUALIFY ONE ITEM '+$env:SMARTM365_BRIDGE_TEST_ID+' v1.0.7' }
 $errorText=''
 try { & $Qualification -ProjectRoot $project -AnalysisDirectory $analysis -SessionId '261103-1' -SourceItemId $id -FarmTimeZoneId $zone.Id -ExpectedAnalysisHash $hash -SpoIdentityPipe $PipeName -Run -ConfirmQualification | Out-Null }
 catch { $errorText=$_.Exception.Message }
 if($id -eq 8){if($errorText -notmatch 'COPY_REACHED' -or -not (Test-Path -LiteralPath $env:SMARTM365_BRIDGE_TEST_COPY)){throw ('Correct ID did not reach simulated copy: '+$errorText)}}
 else {if($errorText -notmatch 'verification failed' -or (Test-Path -LiteralPath $env:SMARTM365_BRIDGE_TEST_COPY)){throw ('Wrong ID was not blocked before simulated copy: '+$errorText)}}
}
Write-Output 'PASS: qualification reaches mocked copy only for the verified source ID.'
'@ | Set-Content -LiteralPath $integration -Encoding utf8
    $qualification=Join-Path $root '../Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGatePathQualification.ps1'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $integration -Qualification $qualification -Scratch $scratch -PipeName $bridge.PipeName
    if($LASTEXITCODE -ne 0){throw 'Simulated ShareGate identity integration failed.'}
}
finally {
    if($bridge){Stop-SmartM365SpoIdentityBridge -Bridge $bridge}
    $env:PSModulePath=$oldPath
    if(Test-Path -LiteralPath $scratch){Remove-Item -LiteralPath $scratch -Recurse -Force}
}
