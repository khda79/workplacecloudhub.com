<#
.SYNOPSIS
    Verify CSOM owning-context isolation with strict fictional client objects.
.VERSION
    1.0.0
#>
#Requires -Version 7.4
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2
$project=Split-Path $PSScriptRoot -Parent
$script:Checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:Checks++}
$tokens=$null;$errors=$null
$path=Join-Path $project 'Scripts/Inventory/SmartM365-SharePointTarget-PermissionInventory.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Target permission source does not parse.'}
foreach($name in @('Get-SPOClientObjectContext','Get-SPOPermissionListPage','Add-SPOInheritanceRead','Get-SPOPageInheritance','Initialize-SPORoleAssignmentBatch')){
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    if(-not $node){throw ('Missing function: '+$name)}
    $body=$node.Body.Extent.Text
    Set-Item ('Function:'+$name) ([scriptblock]::Create($body.Substring(1,$body.Length-2)))
}
foreach($name in @('Connect-PnPOnline','Get-PnPWeb','Get-PnPList','Get-PnPListItem','Get-PnPProperty','Invoke-WebRequest','Invoke-RestMethod','Start-Process')){
    Set-Item ('Function:'+$name) {throw 'A live transport was reached by an offline context test.'}
}
$script:ConnectionContextReads=0
function Get-PnPContext {param($Connection) $script:ConnectionContextReads++;$script:OtherContext}
function Invoke-SPORead {param($Label,[scriptblock]$Operation) & $Operation}
$script:Warnings=[Collections.Generic.List[string]]::new()
function Write-ConsoleWarning {param($Message) $script:Warnings.Add($Message)}
function Get-SPOExceptionDetails {param($Exception) $Exception.Message}
$script:SPOInheritanceLoad=$null
$script:SPOInheritanceItemType=$null
$script:SPOInheritanceSelectors=$null
Add-Type 'using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Xml;
namespace Microsoft.SharePoint.Client {
 public class CamlQuery { public string ViewXml {get;set;} public ListItemCollectionPosition ListItemCollectionPosition {get;set;} }
 public class ListItemCollectionPosition { public string PagingInfo {get;set;} }
}
public interface IOwnedFixture { StrictOwnerContext Context {get;} }
public class OwnedFixture : IOwnedFixture { public StrictOwnerContext Context {get;set;} }
public class OwnedRoleAssignment {
 public OwnedFixture Member {get;set;}
 public OwnedFixture RoleDefinitionBindings {get;set;}
}
public class OwnedRoleCollection : List<OwnedRoleAssignment>, IOwnedFixture { public StrictOwnerContext Context {get;set;} }
public class OwnedItem : OwnedFixture {
 public int Id {get;set;}
 public bool HasUniqueRoleAssignments {get;set;}
 public OwnedRoleCollection RoleAssignments {get;set;}
}
public class OwnedPage : List<OwnedItem>, IOwnedFixture {
 public StrictOwnerContext Context {get;set;}
 public Microsoft.SharePoint.Client.ListItemCollectionPosition ListItemCollectionPosition {get;set;}
}
public class OwnedList : OwnedFixture {
 public string Title {get;set;}
 public List<OwnedItem> Items=new List<OwnedItem>();
 public OwnedPage GetItems(Microsoft.SharePoint.Client.CamlQuery query) {
  var xml=new XmlDocument(); xml.LoadXml(query.ViewXml);
  int limit=int.Parse(xml.SelectSingleNode("//RowLimit").InnerText);
  int after=query.ListItemCollectionPosition==null?0:int.Parse(query.ListItemCollectionPosition.PagingInfo);
  var page=new OwnedPage{Context=Context};
  foreach(var item in Items){if(item.Id>after && page.Count<limit)page.Add(item);}
  if(page.Count>0 && page[page.Count-1].Id<Items[Items.Count-1].Id)
   page.ListItemCollectionPosition=new Microsoft.SharePoint.Client.ListItemCollectionPosition{PagingInfo=page[page.Count-1].Id.ToString()};
  return page;
 }
}
public class StrictOwnerContext {
 public int Queries; public int Loads;
 public void Load(object value){
  var owned=value as IOwnedFixture;
  if(owned==null || !object.ReferenceEquals(owned.Context,this))throw new Exception("The object is used in the context different from the one associated with the object.");
  Loads++;
 }
 public void Load<T>(T value,params Expression<Func<T,object>>[] selectors){Load((object)value);}
 public void ExecuteQuery(){Queries++;}
}'
$owner=[StrictOwnerContext]::new();$script:OtherContext=[StrictOwnerContext]::new()
$script:SPOPermissionConnection=[pscustomobject]@{Context=$script:OtherContext}
$list=[OwnedList]::new();$list.Context=$owner;$list.Title='DEMO Docs'
foreach($id in @(3,7,11,17,22)){
    $item=[OwnedItem]::new();$item.Context=$owner;$item.Id=$id;$item.HasUniqueRoleAssignments=($id%2 -eq 0)
    $item.RoleAssignments=[OwnedRoleCollection]::new();$item.RoleAssignments.Context=$owner
    $member=[OwnedFixture]::new();$member.Context=$owner
    $bindings=[OwnedFixture]::new();$bindings.Context=$owner
    $assignment=[OwnedRoleAssignment]::new();$assignment.Member=$member;$assignment.RoleDefinitionBindings=$bindings
    $item.RoleAssignments.Add($assignment);$list.Items.Add($item)
}
# First prove that this fixture rejects the exact context error seen in production.
$failed=$false
try{$script:OtherContext.Load($list.Items[0])}catch{$failed=$_.Exception.Message -match 'context different'}
Check $failed 'Fixture must reject cross-context loads.'
$seen=[Collections.Generic.List[int]]::new();$position=$null
$script:SPOPermissionConnection.Context=$owner
# Simulate a reused connection being redirected after the list was obtained.
$script:SPOPermissionConnection.Context=$script:OtherContext
do{
    $page=Get-SPOPermissionListPage $list $position 2
    foreach($item in $page.Items){$seen.Add($item.Id)}
    $position=$page.NextPagingInfo
}while($position)
Check (($seen -join ',') -eq '3,7,11,17,22' -and $owner.Queries -eq 3) 'Owning-context paging lost sparse IDs or the final singleton.'
Check ($script:OtherContext.Queries -eq 0 -and $script:ConnectionContextReads -eq 0) 'Paging used the mutable connection context.'
$inheritance=Get-SPOPageInheritance -PageItems $list.Items.ToArray() -Web ([pscustomobject]@{Url='https://demo.invalid'}) -List $list -ListUrl '/Docs'
Check ($inheritance.Count -eq 5 -and $inheritance[22] -and -not $inheritance[3]) 'Inheritance reads lost values or used another context.'
$loaded=Initialize-SPORoleAssignmentBatch -Items $list.Items.ToArray() -ListTitle 'DEMO Docs'
Check ($loaded.Count -eq 5 -and $owner.Queries -eq 6) 'Role preload must use two requests on the owning context.'
Check ($script:OtherContext.Queries -eq 0 -and $script:ConnectionContextReads -eq 0) 'Inheritance or role preload used the connection context.'
$failed=$false
try{Add-SPOInheritanceRead -Context $script:OtherContext -Item $list.Items[0]}catch{$failed=$true}
Check $failed 'Cross-context inheritance must fail before queuing a request.'
$failed=$false
try{Get-SPOClientObjectContext ([pscustomobject]@{Title='DEMO missing context'})}catch{$failed=$true}
Check $failed 'A missing owner context must never fall back to the connection context.'
$mixed=[OwnedItem]::new();$mixed.Context=$script:OtherContext;$mixed.Id=99
$loaded=Initialize-SPORoleAssignmentBatch -Items @($list.Items[0],$mixed) -ListTitle 'DEMO mixed'
Check ($loaded.Count -eq 0 -and $script:Warnings.Count -eq 1) 'Mixed-context batches must remain unpreloaded for individual fallback.'
$failed=$false
try{Get-SPOPageInheritance -PageItems @($mixed) -Web ([pscustomobject]@{Url='https://demo.invalid'}) -List $list -ListUrl '/Docs'}catch{$failed=$true}
Check $failed 'Contradictory list/item contexts must not silently establish inheritance.'
('PASS: {0} CSOM owning-context checks; fictional objects only.' -f $script:Checks)

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDKOjJ2QJQTjz26
# Q39jh+q9FwVRlhxuIYZ3xUlJswUOPqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCoQv3+yghBqTuJQjAB1BIs
# mNLcohG3UK+OXDng9VStOTANBgkqhkiG9w0BAQEFAASCAYBDZRhJ3NyBWrfRu9G2
# Tpqt/VTqZUV9ZPp1dlc5RQSUUyAErK5zJfOyIoOxW2qAIuNCljj08UpM5MMiZWeS
# hYPL4j6F9glk1ZacOgtpXkQpvlL75VFSZ6woHnY4EqOs4BJk81AaaVRyqRo8qbrH
# 8OvZAcjJ8f7XNcrwQf63zk5WkjkXJoT11k5Dtvrk7yeYwmDLPjDxpwzqFtYtUe/z
# t4M+VdAvK3ONBbOkZznqm3WtNv6iYDItGlfchLd9bLEIfQHoMw6dISjjHt9Z7nyb
# cTrrXwamcEWKeygu1UMWqsRjd1zweck7yM+Yv8bpOENKapoSVbZ2LMve6+GmP7ZT
# YzCMhVho3LJMDfcOBrc8STpRS+ynYwVqdoLEtaf6AQlJt01AlW4EaR2QULCBInZ8
# IVYrCcvCldX5K/3bV8ruammfD7IvlwcqgM1mmT42hxWtWh+YnK0GLhtKGyeOrQcG
# iQtD0/WNH1umI3Reibh25NopTKQUi1WQ7ZvGTGIhqZSA8LQ=
# SIG # End signature block
