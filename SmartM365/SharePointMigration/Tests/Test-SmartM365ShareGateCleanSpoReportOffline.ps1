<#
.SYNOPSIS
    Reject incomplete or failing ShareGate roll-ups for SPO source qualification.
.VERSION
    1.0.1
#>
#Requires -Version 5.1
$ErrorActionPreference='Stop'
$dir=Join-Path $PSScriptRoot '..\Scripts\Diagnostics'
. (Join-Path $dir 'SmartM365-SharePointMigration-DestinationPath.ps1')
. (Join-Path $dir 'SmartM365-SharePointMigration-ShareGateReportReader.ps1')
$aliases=Get-SmartM365ShareGateReportAliases -ConfigRoot (Join-Path $PSScriptRoot '..\Config')
function Make-Rows([string]$rollupStatus,[string]$rollupErrors,[string]$rollupWarnings,[string]$importStatus){
 return @(
  [pscustomobject]@{'Source ID'='8';Status='Success';Errors='';Warnings='';'Destination path'='Folder/file.txt';'Microsoft 365 Import: Status'=''},
  [pscustomobject]@{'Source ID'='';Status=$rollupStatus;Errors=$rollupErrors;Warnings=$rollupWarnings;'Destination path'='';'Microsoft 365 Import: Status'=$importStatus}
 )
}
Assert-SmartM365ShareGateCleanSpoReport -Rows (Make-Rows 'Success' '' '' 'Finished') -SourceItemId '8' -ExpectedDestinationPath 'Folder/file.txt' -Aliases $aliases
foreach($case in @(
 @{Status='Error';Errors='';Warnings='';Import='Finished'},
 @{Status='Started';Errors='';Warnings='';Import='Finished'},
 @{Status='Success';Errors='import failed';Warnings='';Import='Finished'},
 @{Status='Success';Errors='';Warnings='missing field';Import='Finished'},
 @{Status='Success';Errors='';Warnings='';Import='Started'},
 @{Status='Success';Errors='';Warnings='';Import=''}
)){
 $blocked=$false
 try{Assert-SmartM365ShareGateCleanSpoReport -Rows (Make-Rows $case.Status $case.Errors $case.Warnings $case.Import) -SourceItemId '8' -ExpectedDestinationPath 'Folder/file.txt' -Aliases $aliases}catch{$blocked=$true}
 if(-not $blocked){throw 'An erroneous, warning-bearing or incomplete roll-up was accepted.'}
}
$blocked=$false
try{Assert-SmartM365ShareGateCleanSpoReport -Rows (Make-Rows 'Success' '' '' 'Finished') -SourceItemId '8' -ExpectedDestinationPath 'Other/file.txt' -Aliases $aliases}catch{$blocked=$true}
if(-not $blocked){throw 'A different destination path was accepted.'}
$ambiguous=Make-Rows 'Success' '' '' 'Finished'
$ambiguous[1] | Add-Member -NotePropertyName 'Avertissements' -NotePropertyValue 'hidden warning'
$blocked=$false
try{Assert-SmartM365ShareGateCleanSpoReport -Rows $ambiguous -SourceItemId '8' -ExpectedDestinationPath 'Folder/file.txt' -Aliases $aliases}catch{$blocked=$true}
if(-not $blocked){throw 'A warning in an alternate alias was accepted.'}
Write-Output 'PASS: SPO qualification requires a clean item and completed, error-free import roll-up.'
