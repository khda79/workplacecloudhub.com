<#
.SYNOPSIS
    Offline fail-closed tests for SPO source file identity verification.
.VERSION
    1.0.1
#>
#Requires -Version 7.4
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../Scripts/Diagnostics/SmartM365-SharePointMigration-SpoSourceIdentity.ps1')
$script:listId=[guid]'c285d922-b71a-4b41-b546-76a25c44a173'
$script:itemId=8;$script:itemPath='/sites/source/Shared Documents/Folder/Example file.txt';$script:itemListId=$script:listId
$script:root='/sites/source/Shared Documents';$script:reads=0;$script:listTitle='Documents';$script:listType='DocumentLibrary'
function Get-PnPList { [CmdletBinding()]param($Identity,$Includes,$Connection) $script:reads++;[pscustomobject]@{Title=$script:listTitle;BaseType=$script:listType;Id=$script:listId;RootFolder=[pscustomobject]@{ServerRelativeUrl=$script:root}} }
function Get-PnPFile { [CmdletBinding()]param($Url,[switch]$AsListItem,$Connection) $script:reads++;@{Id=$script:itemId;FileRef=$script:itemPath;ParentList=[pscustomobject]@{Id=$script:itemListId}} }
function Get-PnPProperty { [CmdletBinding()]param($ClientObject,$Property,$Connection) }
$source=[pscustomobject]@{Url='https://tenant.sharepoint.com/sites/source/'}
$case=@{Connection=$source;SiteUrl='https://tenant.sharepoint.com/sites/source';ListName='Documents';SourceFilePath='Folder/Example%20file.txt';ExpectedItemId=8}
$ok=Get-SmartM365SpoSourceIdentity @case
if($ok.SourceItemId -ne 8 -or $ok.SourceFilePath -ne 'Folder/Example file.txt' -or $ok.LibraryId -ne [string]$script:listId -or $script:reads -ne 2){throw 'Valid source identity was not verified.'}
function Assert-Rejected($arguments,[string]$label){$blocked=$false;try{Get-SmartM365SpoSourceIdentity @arguments|Out-Null}catch{$blocked=$true};if(-not $blocked){throw "Unsafe $label was accepted."}}
$script:itemId=9;Assert-Rejected $case 'stale item ID';$script:itemId=8
$script:itemPath='/sites/source/Shared Documents/Folder/other.txt';Assert-Rejected $case 'different source path';$script:itemPath='/sites/source/Shared Documents/Folder/Example file.txt'
$script:itemListId=[guid]'30595633-9caa-4cff-b2c4-5db7269b13d4';Assert-Rejected $case 'different parent library';$script:itemListId=$script:listId
$script:listTitle='Other';Assert-Rejected $case 'wrong library title';$script:listTitle='Documents'
$script:listType='GenericList';Assert-Rejected $case 'non-document library';$script:listType='DocumentLibrary'
$script:root='/sites/other/Shared Documents';Assert-Rejected $case 'library outside source web';$script:root='/sites/source/Shared Documents'
foreach($path in @('../secret.txt','Folder/%2fsecret.txt','Folder/%252fsecret.txt','Folder/%zz.txt','/absolute.txt','Folder/a#b.txt')){$bad=$case.Clone();$bad.SourceFilePath=$path;Assert-Rejected $bad "path $path"}
$bad=$case.Clone();$bad.Connection=[pscustomobject]@{Url='https://other.sharepoint.com/sites/source'};Assert-Rejected $bad 'wrong tenant connection'
$bad=$case.Clone();$bad.SiteUrl='https://tenant.sharepoint.com/sites/source?x=1';Assert-Rejected $bad 'source query'
$bad=$case.Clone();$bad.SiteUrl='https://evil.example/sites/source';Assert-Rejected $bad 'non-SPO source'
Write-Host 'PASS: source ID, exact path, parent library, tenant and unsafe paths are checked offline.'
