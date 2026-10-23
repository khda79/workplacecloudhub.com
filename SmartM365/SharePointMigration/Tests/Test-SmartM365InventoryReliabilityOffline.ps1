<#
.SYNOPSIS
    Regression checks for inventory data fidelity, paged reads and workbook snapshots.
.VERSION
    1.0.6
#>
#Requires -Version 7.4
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
$root = Join-Path $PSScriptRoot '..'
function Import-TestFunction {
    param([string]$Path, [string[]]$Names)
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ($errors | ForEach-Object Message) }
    foreach ($name in $Names) {
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if (-not $node) { throw "Missing function: $name" }
        $body=$node.Body.Extent.Text
        Set-Item "function:global:$name" ([scriptblock]::Create($body.Substring(1,$body.Length-2)))
    }
}
function Assert-True { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message } }
function Write-ConsoleWarning { param([string]$Message) }
$script:RetryMessages=[Collections.Generic.List[string]]::new()
function Write-SPOReadRetry { param([string]$Message) $script:RetryMessages.Add($Message) }
function Write-DiagnosticPhase { param([string]$State,[string]$Message) }
$script:Delays=[Collections.Generic.List[int]]::new()
function Start-Sleep { param([int]$Seconds,[int]$Milliseconds) $script:Delays.Add($Seconds) }

foreach ($file in Get-ChildItem (Join-Path $root 'Scripts/Inventory') -Filter '*Inventory.ps1') {
    Import-TestFunction $file.FullName @('Format-InventoryDuration')
    foreach ($case in @(@{ Minutes=31; Expected='00:31:00' },@{ Minutes=59; Expected='00:59:00' },@{ Minutes=106; Expected='01:46:00' },@{ Minutes=1501; Expected='1d 01:01:00' })) {
        Assert-True ((Format-InventoryDuration ([timespan]::FromMinutes($case.Minutes))) -eq $case.Expected) "Rounded duration in $($file.Name)."
    }
}
$permission=Join-Path $root 'Scripts/Inventory/SmartM365-SharePointTarget-PermissionInventory.ps1'
Import-TestFunction $permission @('Connect-ToSPOWeb','Get-PnPGroupTitle','Get-AssociatedWebGroupNames','Invoke-SPORead','Get-SPOExceptionDetails','Get-SPORoleAssignmentIdentity','Get-RoleAssignmentRows','Get-PrincipalInfo','Get-PrincipalMembershipInfo','Get-EmptyPrincipalMembershipInfo','Get-SPOClientObjectContext','Add-SPOInheritanceRead','Get-SPOPageInheritance','Initialize-SPORoleAssignmentBatch','Write-SPOItemError','Get-SPOPermissionListPage','Export-SPOPermissionPage','Export-ItemPermissionInventory','Write-InventoryError')
$script:SPOPermissionConnection='fixture'
$script:SPOInheritanceLoad=$null; $script:SPOInheritanceItemType=$null
$script:AssociatedWebGroupCache=@{}
$script:PropertyCalls=0; $script:GroupFail=$false; $script:CollectionFault=$false; $script:CollectionAttempts=0
function Get-PnPProperty {
    [CmdletBinding()] param($ClientObject,[string[]]$Property,$Connection)
    $script:PropertyCalls++
    if ($Property.Count -gt 1) {
        Assert-True (($Property -join ',') -eq 'Member,RoleDefinitionBindings') 'Role properties were not loaded together.'
        $id = [int]$ClientObject.PrincipalId
        $script:RoleCalls[$id]++
        if ($id -eq 17) {
            switch ($script:RoleFault) {
                'Recover' { if ($script:RoleCalls[$id] -lt 3) { throw [System.Net.Http.HttpRequestException]::new('Error while copying content to a stream.', [IO.IOException]::new("Transport connection reset.`r`nFixture details for item 403.")) } }
                'Persistent' { throw [System.Net.Http.HttpRequestException]::new('Error while copying content to a stream.', [IO.IOException]::new('Transport connection reset.')) }
                'Denied' { throw 'Access denied.' }
                'WrappedDenied' { throw [System.Net.Http.HttpRequestException]::new('Error while copying content to a stream.', [UnauthorizedAccessException]::new('Access is denied.')) }
                'Missing' { throw 'Item does not exist.' }
                'Unknown' { throw 'Unexpected fixture failure.' }
            }
        }
        return
    }
    if ($Property -eq 'RoleAssignments' -and $script:CollectionFault -and $ClientObject.Id -eq 4) {
        $script:CollectionAttempts++
        if ($script:CollectionAttempts -lt 3) { throw 'HTTP 503 temporarily unavailable.' }
    }
    if ($Property -eq 'Title') { if ($script:GroupFail) { throw 'Group access denied.' }; return $ClientObject.Title }
    if ($Property -eq 'HasUniqueRoleAssignments') {
        if (($ClientObject.Id -eq 3 -and $script:Missing) -or $ClientObject.Id -eq $script:MissingId) { throw 'Item does not exist.' }
        return ($ClientObject.Id % 2 -eq 0)
    }
    return $ClientObject.($Property[0])
}
$temp=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ReliabilityTest-'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $temp)
$script:ErrorPath=Join-Path $temp 'errors.csv'; $script:ErrorCsvCreated=$false
try {
    $Tenant='fixture'; $TenantId=''; $ClientId='fixture'; $Thumbprint=''; $DeviceLogin=$false; $Interactive=$true; $ForceAuthentication=$false
    $script:TokenSummaryContexts=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $script:TokenReports=0; $script:Identity='user@example.test'
    function Write-ConsoleMessage { param([string]$Message) }
    function Write-SPOConnectionIdentity { param($Connection,$Url) $script:SPOConnectedAccount=$script:Identity }
    function Write-PnPTokenSummary { param($Connection) $script:TokenReports++ }
    function Connect-PnPOnline {
        [CmdletBinding()] param($Url,$ClientId,$Tenant,[switch]$Interactive,[switch]$ReturnConnection)
        Assert-True ([bool]$ReturnConnection) 'PnP connection was not returned for explicit command binding.'
        return 'fixture'
    }
    Connect-ToSPOWeb 'https://example.test/sites/a'
    Connect-ToSPOWeb 'https://example.test/sites/a/sub'
    Assert-True ($script:TokenReports -eq 1) 'Token diagnostic was repeated for one authentication context.'
    $script:Identity='another@example.test'
    Connect-ToSPOWeb 'https://example.test/sites/b'
    Assert-True ($script:TokenReports -eq 2) 'Token diagnostic was suppressed for a different account.'
    $web=[pscustomobject]@{ Url='https://example.test/sites/a'; Title='Fixture'; AssociatedMemberGroup=[pscustomobject]@{Title='Members'}; AssociatedOwnerGroup=[pscustomobject]@{Title='Owners'}; AssociatedVisitorGroup=$null }
    $groups=Get-AssociatedWebGroupNames $web
    Assert-True ($groups.AssociatedMemberGroup -eq 'Members' -and $groups.AssociatedOwnerGroup -eq 'Owners' -and $groups.AssociatedVisitorGroup -eq '') 'Associated group titles were lost.'
    $calls=$script:PropertyCalls; [void](Get-AssociatedWebGroupNames $web)
    Assert-True ($script:PropertyCalls -eq $calls) 'Associated groups were fetched repeatedly for one web.'
    $script:GroupFail=$true; $web.Url='https://example.test/sites/b'; [void](Get-AssociatedWebGroupNames $web)
    Assert-True ($script:ErrorCsvCreated -and @(Import-Csv $script:ErrorPath -Delimiter ';').Count -eq 2) 'Group failures did not mark the inventory incomplete.'
    $script:GroupFail=$false; $web.Url='https://example.test/sites/a'
    Remove-Item $script:ErrorPath; $script:ErrorCsvCreated=$false
    $script:Attempts=0
    $result=@(Invoke-SPORead -Label 'test' -Operation { $script:Attempts++; 'partial'; if ($script:Attempts -lt 3) { throw 'HttpClient.Timeout' }; 'complete' })
    Assert-True (($result -join ',') -eq 'partial,complete' -and $script:Attempts -eq 3 -and ($script:Delays -join ',') -eq '5,15') 'Read retries emitted partial responses or exceeded their bound.'
    $script:Attempts=0; $failed=$false
    try { Invoke-SPORead -Label 'denied' -Operation { $script:Attempts++; throw 'Access denied' } } catch { $failed=$true }
    Assert-True ($failed -and $script:Attempts -eq 1) 'Access denied was retried or ignored.'
    $script:Attempts=0; $failed=$false
    try { Invoke-SPORead -Label 'timeout' -Operation { $script:Attempts++; throw 'HttpClient.Timeout' } } catch { $failed=$true }
    Assert-True ($failed -and $script:Attempts -eq 3) 'Persistent timeout did not stop after three attempts.'

    $roles = @(17,18 | ForEach-Object {
        [pscustomobject]@{ PrincipalId=$_; Member=[pscustomobject]@{ Id=$_; Title="Principal $_"; LoginName="user$_@example.test"; PrincipalType='User' }; RoleDefinitionBindings=@([pscustomobject]@{Name='Read'},[pscustomobject]@{Name='Contribute'}) }
    })
    foreach ($case in @(
        @{Fault='None'; Calls=1; Rows=2; Errors=0},
        @{Fault='Recover'; Calls=3; Rows=2; Errors=0},
        @{Fault='Persistent'; Calls=3; Rows=1; Errors=1},
        @{Fault='Denied'; Calls=1; Rows=1; Errors=1},
        @{Fault='WrappedDenied'; Calls=1; Rows=1; Errors=1},
        @{Fault='Missing'; Calls=1; Rows=1; Errors=1},
        @{Fault='Unknown'; Calls=1; Rows=1; Errors=1}
    )) {
        if (Test-Path $script:ErrorPath) { Remove-Item $script:ErrorPath }
        $script:ErrorCsvCreated=$false; $script:Delays.Clear(); $script:RetryMessages.Clear()
        $script:RoleCalls=@{17=0;18=0}; $script:RoleFault=$case.Fault
        $rows=@(Get-RoleAssignmentRows -ObjectScope 'Item' -ObjectUrl 'https://example.test/sites/a/docs/folder' -ObjectTitle 'Fixture folder' -ItemId 501 -RoleAssignments $roles)
        Assert-True ($script:RoleCalls[17] -eq $case.Calls -and $script:RoleCalls[18] -eq 1) "Unexpected read count for $($case.Fault)."
        Assert-True ($rows.Count -eq $case.Rows -and @($rows.PrincipalId | Sort-Object -Unique).Count -eq $case.Rows) "Missing or duplicate role rows for $($case.Fault)."
        Assert-True (@($rows | Where-Object PermissionLevels -ne 'Read|Contribute').Count -eq 0) 'Grouped loading changed permission levels.'
        $expectedDelays=if ($case.Calls -eq 3) { '5,15' } else { '' }
        Assert-True (($script:Delays -join ',') -eq $expectedDelays) "Unexpected retry delays for $($case.Fault)."
        if ($case.Calls -eq 3) {
            Assert-True ($script:RetryMessages.Count -eq 2 -and $script:RetryMessages[0] -match 'role assignment #1.*docs/folder.*principal ID=17.*HttpRequestException.*IOException' -and $script:RetryMessages[0] -notmatch '[\r\n]') 'Retry diagnostics lost path, principal or inner exception context.'
        }
        Assert-True ($script:ErrorCsvCreated -eq [bool]$case.Errors) "Incomplete inventory marker changed for $($case.Fault)."
        if ($case.Errors) {
            $errors=@(Import-Csv $script:ErrorPath -Delimiter ';')
            Assert-True ($errors.Count -eq 1 -and $errors[0].Scope -eq 'Item RoleAssignment' -and $errors[0].ItemId -eq '501' -and $errors[0].ItemUrl -eq 'https://example.test/sites/a/docs/folder') 'Terminal error lost its object context.'
            Assert-True ($errors[0].Message -match 'principal ID=17.*user17@example.test' -and $errors[0].Message -notmatch '[\r\n]') 'Principal identity or single-line diagnostics were lost.'
            if ($case.Fault -eq 'Persistent') { Assert-True ($errors[0].Message -match 'HttpRequestException.*IOException.*Transport connection reset') 'Inner transport exception was lost.' }
        }
    }
    Remove-Item $script:ErrorPath; $script:ErrorCsvCreated=$false
    # Preloaded role assignments (batched path) must produce exactly the rows of the
    # per-assignment path, without any further per-assignment request.
    $script:RoleCalls=@{17=0;18=0}; $script:RoleFault='None'
    $readRows=@(Get-RoleAssignmentRows -ObjectScope 'Item' -ObjectUrl 'https://example.test/sites/a/docs/folder' -ObjectTitle 'Fixture folder' -ItemId 501 -RoleAssignments $roles)
    $script:RoleCalls=@{17=0;18=0}
    $preloadedRows=@(Get-RoleAssignmentRows -ObjectScope 'Item' -ObjectUrl 'https://example.test/sites/a/docs/folder' -ObjectTitle 'Fixture folder' -ItemId 501 -RoleAssignments $roles -Preloaded)
    Assert-True ($script:RoleCalls[17] -eq 0 -and $script:RoleCalls[18] -eq 0) 'Preloaded role assignments were read again one by one.'
    Assert-True (($preloadedRows | ConvertTo-Json -Depth 3) -ceq ($readRows | ConvertTo-Json -Depth 3)) 'Preloaded role assignments changed the exported permission rows.'

    Add-Type @'
using System;
using System.Linq.Expressions;
using System.Collections.Generic;
namespace Microsoft.SharePoint.Client {
 public class ListItemCollectionPosition { public string PagingInfo {get;set;} }
 public class CamlQuery {
  public string ViewXml {get;set;}
  public ListItemCollectionPosition ListItemCollectionPosition {get;set;}
 }
}
public class ReliabilityPage : List<ReliabilityItem> {
 public Microsoft.SharePoint.Client.ListItemCollectionPosition ListItemCollectionPosition {get;set;}
}
public class ReliabilityPageList {
 public List<ReliabilityItem> Items=new List<ReliabilityItem>();
 public ReliabilityPage GetItems(Microsoft.SharePoint.Client.CamlQuery query) {
  var xml=new System.Xml.XmlDocument(); xml.LoadXml(query.ViewXml);
  if(xml.SelectSingleNode("//Where")!=null && Items.Count>5000) throw new Exception("The attempted operation is prohibited because it exceeds the list view threshold.");
  if(xml.SelectSingleNode("/View").Attributes["Scope"].Value!="RecursiveAll") throw new Exception("Recursive scope lost");
  if(xml.SelectSingleNode("//OrderBy").Attributes["Override"].Value!="TRUE") throw new Exception("Default view order retained");
  int limit=int.Parse(xml.SelectSingleNode("//RowLimit").InnerText);
  if(limit>2000 || xml.SelectSingleNode("//RowLimit").Attributes["Paged"].Value!="TRUE") throw new Exception("Page limit lost");
  foreach(string field in new[]{"ID","FileRef","FileLeafRef","FSObjType","UniqueId"})
   if(xml.SelectSingleNode("//ViewFields/FieldRef[@Name='"+field+"']")==null) throw new Exception("Field lost: "+field);
  int after=query.ListItemCollectionPosition==null?0:int.Parse(query.ListItemCollectionPosition.PagingInfo.Split('=')[2]);
  var page=new ReliabilityPage();
  foreach(var item in Items) { if(item.Id<=after) continue; if(page.Count==limit) break; page.Add(item); }
  if(page.Count>0 && page[page.Count-1].Id<Items[Items.Count-1].Id)
   page.ListItemCollectionPosition=new Microsoft.SharePoint.Client.ListItemCollectionPosition { PagingInfo="Paged=TRUE&p_ID="+page[page.Count-1].Id };
  return page;
 }
}
public class ReliabilityPageContext {
 public int Queries;
 public void Load(object items) { if(!(items is ReliabilityPage)) throw new Exception("Unexpected collection"); }
 public void ExecuteQuery() { Queries++; }
}
public class ReliabilityItem {
 public int Id {get;set;}
 public bool HasUniqueRoleAssignments {get;set;}
 public Dictionary<string,object> FieldValues {get;set;}
 public object[] RoleAssignments {get;set;}
}
public class ReliabilityContext {
 public int Queries;
 public bool Missing;
 public int MissingId;
 public int FailQuery;
 public string FailMessage="HttpClient.Timeout";
 public int RequestLimitBytes=2097152;
 public List<int> BatchSizes=new List<int>();
 public List<int> ReadIds=new List<int>();
 public List<ReliabilityItem> Pending=new List<ReliabilityItem>();
 public void Load<T>(T item, params Expression<Func<T,object>>[] selectors) {
   if(selectors.Length!=1 || !selectors[0].ToString().Contains("HasUniqueRoleAssignments")) throw new Exception("Unexpected CSOM selector");
   Pending.Add((ReliabilityItem)(object)item);
 }
 public void ExecuteQuery() {
   Queries++;
   var items=Pending.ToArray(); Pending.Clear();
   BatchSizes.Add(items.Length);
   // Simulate a large CSOM object path per item and SPO's actual message limit.
   if(items.Length*2048>RequestLimitBytes) throw new Exception("The request message is too big. The server does not allow messages larger than 2097152 bytes.");
   if(Queries==FailQuery) throw new Exception(FailMessage);
   if(Missing || Array.Exists(items,item=>item.Id==MissingId)) throw new Exception("Item does not exist.");
   foreach(var item in items) { item.HasUniqueRoleAssignments=item.Id%2==0; ReadIds.Add(item.Id); }
 }
}
'@
    $script:Context=[ReliabilityContext]::new()
    $script:Missing=$false
    $script:MissingId=0; $script:LargePages=$false
    function Get-PnPContext { param($Connection) return $script:Context }
    # Transport fixture; owner-context isolation is covered by Test-SmartM365CsomContextOffline.ps1.
    function Get-SPOClientObjectContext { param($ClientObject) return $script:Context }
    function ConvertTo-AbsoluteSharePointUrl { param([string]$WebUrl,[string]$ServerRelativeUrl) return "https://example.test$ServerRelativeUrl" }
    function Get-SiteCollectionUrlFromWebUrl { param($WebUrl) return $WebUrl }
    function Write-ItemPermissionHeartbeat { }
    function Get-RoleAssignmentRows { param($ItemId) return [pscustomobject]@{ ItemId=$ItemId } }
    $script:Exported=[Collections.Generic.List[object]]::new()
    function Export-PermissionRows { param($Rows,$CsvPath) foreach($r in $Rows) { $script:Exported.Add($r) } }
    $script:Items=@(foreach($id in 1..4) {
        $item=[ReliabilityItem]::new(); $item.Id=$id; $item.RoleAssignments=@()
        $item.FieldValues=[Collections.Generic.Dictionary[string,object]]::new()
        $item.FieldValues['FileRef']="/sites/a/Docs/$id.txt"; $item.FieldValues['FileLeafRef']="$id.txt"
        $item.FieldValues['UniqueId']=[string]$id; $item.FieldValues['FSObjType']='0'; $item
    })
    function Get-PnPListItem { [CmdletBinding()] param($List,[int]$Id,$Connection) throw 'Item does not exist. Existence probe.' }
    $savedPageReader=${function:Get-SPOPermissionListPage}
    # Exercise the real page reader against a threshold-enforcing CSOM fixture.
    $fixtureList=[ReliabilityPageList]::new()
    foreach($id in 1..6001) { $item=[ReliabilityItem]::new(); $item.Id=$id*3; $fixtureList.Items.Add($item) }
    $script:Context=[ReliabilityPageContext]::new()
    $position=$null; $seen=[Collections.Generic.List[int]]::new()
    do {
        $page=Get-SPOPermissionListPage -List $fixtureList -PagingInfo $position -PageSize 8000
        foreach($item in $page.Items) { $seen.Add($item.Id) }
        $position=$page.NextPagingInfo
    } while($position)
    Assert-True ($seen.Count -eq 6001 -and ($seen -join ',') -eq ($fixtureList.Items.Id -join ',') -and $script:Context.Queries -eq 4) 'Threshold-safe continuation lost sparse IDs or the final singleton page.'
    $fixtureList.Items.Clear()
    $page=Get-SPOPermissionListPage -List $fixtureList -PageSize 2000
    Assert-True ($page.Items.Count -eq 0 -and -not $page.NextPagingInfo) 'An empty library did not finish as an empty page.'
    $script:Context=[ReliabilityContext]::new()
    function Get-SPOPermissionListPage {
        param($List,[string]$PagingInfo,[int]$PageSize)
        if (-not $PagingInfo) {
            $end = if ($script:LargePages) { 1999 } else { 1 }
            return [pscustomobject]@{ Items=$script:Items[0..$end]; NextPagingInfo='Paged=TRUE&p_ID=next' }
        }
        if ($script:LargePages) {
            Assert-True ($script:Exported.Count -eq 1000) 'Large first page was not exported before fetching the final page.'
            return [pscustomobject]@{ Items=@($script:Items[2000]); NextPagingInfo=$null }
        }
        Assert-True ($script:Exported.Count -eq 1) 'First page was not processed before fetching the second page.'
        return [pscustomobject]@{ Items=$script:Items[2..3]; NextPagingInfo=$null }
    }
    $list=[pscustomobject]@{ Title='Docs'; BaseTemplate=101; BaseType='DocumentLibrary'; RootFolder=[pscustomobject]@{ServerRelativeUrl='/sites/a/Docs'} }
    $PageSize=2
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 2
    Assert-True ($script:Context.Queries -eq 2 -and ($script:Exported.ItemId -join ',') -eq '2,4' -and -not $script:ErrorCsvCreated) 'Inheritance was not loaded once per page, or unique items were lost.'
    $script:Context=[ReliabilityContext]::new(); $script:Exported.Clear(); $script:Delays.Clear(); $script:CollectionFault=$true
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 2
    Assert-True ($script:CollectionAttempts -eq 3 -and ($script:Exported.ItemId -join ',') -eq '2,4' -and -not $script:ErrorCsvCreated -and ($script:Delays -join ',') -eq '5,15') 'Role-assignment collection retry lost item permissions or recorded a recoverable failure.'
    $script:CollectionFault=$false
    # A lost next-page response retries that continuation, including gaps in IDs.
    $script:Context=[ReliabilityContext]::new(); $script:Exported.Clear(); $script:Delays.Clear()
    $script:PageAttempts=0; $script:ResumeQueries=[Collections.Generic.List[string]]::new(); $script:TerminalPageFailure=$false
    $savedMockPageReader=${function:Get-SPOPermissionListPage}
    function Get-SPOPermissionListPage {
        param($List,[int]$PageSize,[string]$PagingInfo)
        $script:PageAttempts++; $script:ResumeQueries.Add($PagingInfo)
        if (-not $PagingInfo) {
            return [pscustomobject]@{ Items=$script:Items[0..1]; NextPagingInfo='Paged=TRUE&p_ID=7' }
        }
        if ($script:TerminalPageFailure -or $script:PageAttempts -eq 2) { throw 'Response ended prematurely.' }
        return [pscustomobject]@{ Items=$script:Items[2..3]; NextPagingInfo=$null }
    }
    $script:Items[1].Id=7; $script:Items[2].Id=13; $script:Items[3].Id=18
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 2
    Assert-True ($script:PageAttempts -eq 3 -and $script:ResumeQueries[1] -eq 'Paged=TRUE&p_ID=7' -and $script:ResumeQueries[2] -eq $script:ResumeQueries[1] -and ($script:Exported.ItemId -join ',') -eq '18' -and $script:SPOPermissionPageState.Processed -eq 4 -and -not $script:ErrorCsvCreated) 'Next-page retry duplicated or skipped completed pages.'
    $script:Context=[ReliabilityContext]::new(); $script:Exported.Clear(); $script:PageAttempts=0; $script:TerminalPageFailure=$true; $script:Delays.Clear()
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 2
    Assert-True ($script:PageAttempts -eq 4 -and $script:ErrorCsvCreated -and $script:SPOPermissionPageState.Processed -eq 2 -and ($script:Delays -join ',') -eq '5,15') 'Terminal page failure was published or retried without a bound.'
    Remove-Item $script:ErrorPath; $script:ErrorCsvCreated=$false
    foreach($i in 0..3) { $script:Items[$i].Id=$i+1 }
    Set-Item function:Get-SPOPermissionListPage $savedMockPageReader
    $script:Missing=$true; $script:Context.Missing=$true
    $inheritance=Get-SPOPageInheritance -PageItems $script:Items[2..3] -Web $web -List $list -ListUrl 'https://example.test/sites/a/Docs'
    $record=@(Import-Csv $script:ErrorPath -Delimiter ';')
    Assert-True ($inheritance.Count -eq 1 -and $inheritance[4] -and -not $inheritance.ContainsKey(3)) 'Failed inheritance reads were silently treated as inherited.'
    Assert-True ($record.Count -eq 1 -and $record[0].ItemId -eq '3' -and $record[0].ItemUrl -eq 'https://example.test/sites/a/Docs/3.txt' -and $record[0].Message -match 'Existence check failed') 'Missing item evidence lacks ID, path or existence check.'
    Remove-Item $script:ErrorPath; $script:ErrorCsvCreated=$false; $script:Missing=$false
    $script:Items=@(foreach($id in 1..2001) {
        $item=[ReliabilityItem]::new(); $item.Id=$id; $item.RoleAssignments=@()
        $item.FieldValues=[Collections.Generic.Dictionary[string,object]]::new()
        $item.FieldValues['FileRef']="/sites/a/Docs/$id.txt"; $item.FieldValues['FileLeafRef']="$id.txt"
        $item.FieldValues['UniqueId']=[string]$id; $item.FieldValues['FSObjType']='0'; $item
    })
    $script:Context=[ReliabilityContext]::new()
    foreach($item in $script:Items[0..1999]) { Add-SPOInheritanceRead -Context $script:Context -Item $item }
    $failed=$false
    try { $script:Context.ExecuteQuery() } catch { $failed=$_.Exception.Message -match '2097152' }
    Assert-True $failed 'Oversized-request fixture did not reproduce the server error.'
    $script:Context=[ReliabilityContext]::new(); $script:Exported.Clear(); $script:LargePages=$true; $PageSize=2000
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 200
    Assert-True ($script:Context.Queries -eq 21 -and ($script:Context.BatchSizes | Measure-Object -Maximum).Maximum -le 100) 'Inheritance requests exceeded the batch cap or lost the final singleton page.'
    Assert-True (($script:Context.ReadIds -join ',') -eq ((1..2001) -join ',') -and $script:SPOPermissionPageState.Processed -eq 2001) 'Large-page inheritance dropped, reordered or duplicated items.'
    Assert-True (($script:Exported.ItemId -join ',') -eq ((2..2000 | Where-Object { $_ % 2 -eq 0 }) -join ',') -and -not $script:ErrorCsvCreated) 'Large-page export lost unique permissions or marked a complete read as failed.'
    $script:Context=[ReliabilityContext]::new(); $script:Context.FailQuery=8; $script:Delays.Clear()
    $inheritance=Get-SPOPageInheritance -PageItems $script:Items -Web $web -List $list -ListUrl 'https://example.test/sites/a/Docs'
    Assert-True ($inheritance.Count -eq 2001 -and $script:Context.Queries -eq 22 -and ($script:Context.ReadIds -join ',') -eq ((1..2001) -join ',') -and ($script:Delays -join ',') -eq '5') 'A retried middle batch lost items, duplicated successful reads or retried the whole page.'
    $script:Context=[ReliabilityContext]::new(); $script:Context.MissingId=103; $script:MissingId=103
    $inheritance=Get-SPOPageInheritance -PageItems $script:Items -Web $web -List $list -ListUrl 'https://example.test/sites/a/Docs'
    $record=@(Import-Csv $script:ErrorPath -Delimiter ';')
    Assert-True ($inheritance.Count -eq 2000 -and -not $inheritance.ContainsKey(103) -and $inheritance.ContainsKey(2001) -and $script:Context.Queries -eq 21) 'A missing item stopped later batches or was treated as inherited.'
    Assert-True ($record.Count -eq 1 -and $record[0].ItemId -eq '103' -and $script:ErrorCsvCreated) 'The failed batch did not retain exact missing-item evidence.'
    Remove-Item $script:ErrorPath; $script:ErrorCsvCreated=$false; $script:MissingId=0
    $script:Context=[ReliabilityContext]::new(); $script:Context.FailQuery=2; $script:Context.FailMessage='Access denied'
    $script:Exported.Clear()
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 200
    $record=@(Import-Csv $script:ErrorPath -Delimiter ';')
    Assert-True ($script:Context.Queries -eq 2 -and $script:ErrorCsvCreated -and $script:Exported.Count -eq 0 -and $record[0].Message -match 'IDs 101-200.*Access denied') 'A failed middle batch was retried, silently accepted or lacked request context.'
    # Batched role-assignment reads: two requests per 20 unique items; a batch that fails
    # without a transient cause is left to the unchanged one-by-one path.
    Add-Type @'
public class BatchRoleContext {
 public int Queries; public int Loads; public int FailQuery;
 public void Load(object target) { Loads++; }
 public void ExecuteQuery() { Queries++; if (Queries == FailQuery) throw new System.Exception("Access denied"); }
}
'@
    $batchItems=@(foreach($id in 1..45) { [pscustomobject]@{ Id=$id; RoleAssignments=@($roles) } })
    $script:Context=[BatchRoleContext]::new()
    $loaded=Initialize-SPORoleAssignmentBatch -Items $batchItems -ListTitle 'Docs'
    Assert-True ($loaded.Count -eq 45 -and $script:Context.Queries -eq 6 -and $script:Context.Loads -eq (45 + 45 * 2 * 2)) 'Role assignments were not read in two requests per batch of 20 items.'
    $script:Context=[BatchRoleContext]::new(); $script:Context.FailQuery=3; $script:Delays.Clear()
    $loaded=Initialize-SPORoleAssignmentBatch -Items $batchItems -ListTitle 'Docs'
    Assert-True ($loaded.Count -eq 25 -and -not $loaded.Contains(21) -and -not $loaded.Contains(40) -and $loaded.Contains(41) -and $script:Delays.Count -eq 0) 'A failed batch was retried, preloaded, or stopped later batches.'
    $loaded=Initialize-SPORoleAssignmentBatch -Items @() -ListTitle 'Docs'
    Assert-True ($loaded.Count -eq 0) 'An empty page did not return an empty preload set.'

    $source=[IO.File]::ReadAllText($permission)
    Assert-True ($source -match 'if \(\$script:ErrorCsvCreated\)\s*\{\s*throw') 'Incomplete inventory publication guard was removed.'

    Import-TestFunction (Join-Path $root 'Scripts/Diagnostics/SmartM365-SharePointMigration-Diagnostics.ps1') @('Copy-StableShareGateWorkbook')
    $book=Join-Path $temp 'report.xlsx'; $copy=Join-Path $temp 'snapshot.xlsx'
    function New-TestWorkbook {
        param([string]$Path)
        $zip=[IO.Compression.ZipFile]::Open($Path,[IO.Compression.ZipArchiveMode]::Create)
        try { foreach($name in @('[Content_Types].xml','xl/workbook.xml')) { $stream=[IO.StreamWriter]::new($zip.CreateEntry($name).Open()); try { $stream.Write('<fixture/>') } finally { $stream.Dispose() } } } finally { $zip.Dispose() }
    }
    New-TestWorkbook $book
    [void](Copy-StableShareGateWorkbook $book $copy)
    Assert-True ((Get-FileHash $book).Hash -eq (Get-FileHash $copy).Hash) 'Snapshot bytes differ from original.'
    $writer=[IO.File]::Open($book,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $failed=$false
    try { [void](Copy-StableShareGateWorkbook $book $copy -Attempts 2) } catch { $failed=$true } finally { $writer.Dispose() }
    Assert-True $failed 'A locked workbook was analyzed.'
    [IO.File]::WriteAllText($book,'incomplete export')
    $failed=$false
    try { [void](Copy-StableShareGateWorkbook $book $copy -Attempts 2) } catch { $failed=$true }
    Assert-True $failed 'An incomplete XLSX package was accepted.'
    Import-TestFunction (Join-Path $root 'Scripts/Launchers/Generic/SmartM365-SharePointMigration-Launcher.ps1') @('Write-LauncherRunResult')
    Import-TestFunction (Join-Path $root 'Scripts/Launchers/Generic/SmartM365-SharePointMigration-TargetScanBatch.ps1') @('Add-BatchResult')
    function Write-BatchLine { param([string]$Message) }
    $RunResultPath=Join-Path $temp 'result.json.txt'; $RunResultId='test-run'; $MigrationName='SiteA'; $Action='ScanTargetFiles'
    $script:LauncherRunStartedAt=Get-Date; $script:LauncherStatus='SUCCESS'; $script:LauncherRunLogPath=Join-Path $temp 'run.log'
    $script:LauncherOutputCsvPath=Join-Path $temp 'inventory.csv'
    Set-Content $script:LauncherOutputCsvPath 'fixture'
    Write-LauncherRunResult
    $job=[pscustomobject]@{ ResultPath=$RunResultPath; RunId=$RunResultId; Migration=$MigrationName; Action=$Action; Started=Get-Date }
    $script:Results=[Collections.Generic.List[object]]::new()
    Add-BatchResult -Job $job -Status SUCCESS -ExitCode 0
    Assert-True ($script:Results[0].OutputLog -eq $script:LauncherRunLogPath -and $script:Results[0].OutputCsv -eq $script:LauncherOutputCsvPath) 'Actual interactive scan log/output paths were not preserved.'
    $job.RunId='another-run'
    Add-BatchResult -Job $job -Status SUCCESS -ExitCode 0
    Assert-True (-not $script:Results[1].RunLog -and -not $script:Results[1].OutputCsv) 'Mismatched run receipt was accepted.'
    Remove-Item $script:LauncherOutputCsvPath
    $script:LauncherStatus='FAILED'
    Write-LauncherRunResult
    $receipt=Get-Content $RunResultPath -Raw | ConvertFrom-Json
    Assert-True ($receipt.Status -eq 'FAILED' -and -not $receipt.OutputCsv -and $receipt.LogPath -eq $script:LauncherRunLogPath) 'Failed scan receipt advertises a nonexistent published inventory.'
    'PASS: durations, group fidelity/cache, token contexts, grouped role reads/retries/diagnostics, bounded reads, size-limited streamed inheritance, missing-item evidence, stable workbook snapshots and batch receipts.'
}
finally {
    Get-ChildItem -LiteralPath $temp -File | Remove-Item -Force
    Remove-Item -LiteralPath $temp
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDPrZSyDEYYfjvi
# bpyG1miU0uK/GapkjpREDVWeXVWmbqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDkvvvN1/drJT17Tj+xGfJ6
# +u+6isqUXtb6V+2foyetsDANBgkqhkiG9w0BAQEFAASCAYAl6XeeLtRouK/Cwqk9
# SSfeX20ou9SV8lkCYCnYCELBTHOinxfltJP+QZcAMO/sybF2Z457ojyIqnyHfWCw
# IyUWL5q50d0RZXXwXWgVsRD3+ZU8hReWMDzPHptuGTLDSKWEG3ZbPxY745AdY8uj
# x01NsBwfwAdGyNdwDYS4LFpyOPZ7wJZEJdwQxmzHnN9+hFqQ46hdmm/f0k7QzitV
# XrRE7By2N4g5RuD06PcMrRMg7vXhDzcAOa4aWeyyTdeBHI9P7drSqzJZlZCF/oO+
# 5ChveUcQry5STqlOmRAV0i9pvmmaloq3EPF8CO13Qbi50Lk4ZVG1tBNvJwTgHNOK
# pe6F9NV3IFA7VnxAxH8D0M3L0Ji8qcvj6xcQR+qjgacLUS313WLTP9HBBqTvC7HC
# 5+IIGE8WQgJb5cKE6YRUv8YFOvelhGcgrwEz1unR6ZJY0h/rwcipEWqvQwBLn+di
# x4k0oUpQZ51ab0D5+Bc9wshFpR+8Few74bYLIG2r/Z2jW0A=
# SIG # End signature block
