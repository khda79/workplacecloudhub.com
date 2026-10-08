<#
.SYNOPSIS
    Regression checks for inventory data fidelity, paged reads and workbook snapshots.
.VERSION
    1.0.3
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
Import-TestFunction $permission @('Connect-ToSPOWeb','Get-PnPGroupTitle','Get-AssociatedWebGroupNames','Invoke-SPORead','Get-SPOExceptionDetails','Get-SPORoleAssignmentIdentity','Get-RoleAssignmentRows','Get-PrincipalInfo','Get-PrincipalMembershipInfo','Get-EmptyPrincipalMembershipInfo','Add-SPOInheritanceRead','Get-SPOPageInheritance','Write-SPOItemError','Export-SPOPermissionPage','Export-ItemPermissionInventory','Write-InventoryError')
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

    Add-Type @'
using System;
using System.Linq.Expressions;
using System.Collections.Generic;
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
    function Get-PnPListItem {
        [CmdletBinding()] param($List,[int]$PageSize,[string[]]$Fields,[string]$Query,$Connection,[scriptblock]$ScriptBlock,[int]$Id)
        if ($Id) { throw 'Item does not exist. Existence probe.' }
        if ($script:LargePages) {
            & $ScriptBlock $script:Items[0..1999]
            Assert-True ($script:Exported.Count -eq 1000) 'Large first page was not exported before fetching the final page.'
            & $ScriptBlock @($script:Items[2000])
            return
        }
        & $ScriptBlock $script:Items[0..1]
        Assert-True ($script:Exported.Count -eq 1) 'First page was not processed before fetching the second page.'
        & $ScriptBlock $script:Items[2..3]
    }
    $list=[pscustomobject]@{ Title='Docs'; BaseTemplate=101; BaseType='DocumentLibrary'; RootFolder=[pscustomobject]@{ServerRelativeUrl='/sites/a/Docs'} }
    $PageSize=2
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 2
    Assert-True ($script:Context.Queries -eq 2 -and ($script:Exported.ItemId -join ',') -eq '2,4' -and -not $script:ErrorCsvCreated) 'Inheritance was not loaded once per page, or unique items were lost.'
    $script:Context=[ReliabilityContext]::new(); $script:Exported.Clear(); $script:Delays.Clear(); $script:CollectionFault=$true
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 2
    Assert-True ($script:CollectionAttempts -eq 3 -and ($script:Exported.ItemId -join ',') -eq '2,4' -and -not $script:ErrorCsvCreated -and ($script:Delays -join ',') -eq '5,15') 'Role-assignment collection retry lost item permissions or recorded a recoverable failure.'
    $script:CollectionFault=$false
    # A lost next-page response resumes after the completed page, including gaps in IDs.
    $script:Context=[ReliabilityContext]::new(); $script:Exported.Clear(); $script:Delays.Clear()
    $script:PageAttempts=0; $script:ResumeQueries=[Collections.Generic.List[string]]::new(); $script:TerminalPageFailure=$false
    $savedListItem=${function:Get-PnPListItem}
    function Get-PnPListItem {
        [CmdletBinding()] param($List,[int]$PageSize,[string]$Query,$Connection,[scriptblock]$ScriptBlock,[int]$Id,[string[]]$Fields)
        $script:PageAttempts++; $script:ResumeQueries.Add($Query)
        if ($Id) { throw 'Item does not exist.' }
        if ($Query -match "<Value Type='Counter'>0</Value>") {
            & $ScriptBlock $script:Items[0..1]
            throw 'Response ended prematurely.'
        }
        if ($script:TerminalPageFailure) { throw 'Response ended prematurely.' }
        & $ScriptBlock $script:Items[2..3]
    }
    $script:Items[1].Id=7; $script:Items[2].Id=13; $script:Items[3].Id=18
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 2
    Assert-True ($script:PageAttempts -eq 2 -and $script:ResumeQueries[1] -match "<Value Type='Counter'>7</Value>" -and ($script:Exported.ItemId -join ',') -eq '18' -and $script:SPOPermissionPageState.Processed -eq 4 -and -not $script:ErrorCsvCreated) 'Next-page retry duplicated or skipped completed pages.'
    $script:Context=[ReliabilityContext]::new(); $script:Exported.Clear(); $script:PageAttempts=0; $script:TerminalPageFailure=$true; $script:Delays.Clear()
    Export-ItemPermissionInventory -Web $web -List $list -CsvPath 'fixture.csv' -ProgressInterval 2
    Assert-True ($script:PageAttempts -eq 3 -and $script:ErrorCsvCreated -and $script:SPOPermissionPageState.Processed -eq 2 -and ($script:Delays -join ',') -eq '5,15') 'Terminal page failure was published or retried without a bound.'
    Remove-Item $script:ErrorPath; $script:ErrorCsvCreated=$false
    foreach($i in 0..3) { $script:Items[$i].Id=$i+1 }
    Set-Item function:Get-PnPListItem $savedListItem
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCKG4LOJKWufbrF
# 479ucyqZu5C0PT4Q72MQ/qKwP9ps8qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL8pqCOCJ8ePWU6M7By/7KKj16uwDm7OuceWr1AvULdoMA0GCSqG
# SIb3DQEBAQUABIIBgFc7L8XGmPem5Lml6/eLYVMN1ZMz4WWQHSdOb5ksNj8raP/B
# sJk/UIGRFOsJb/nOcm+dEtcSUdn2YTqvPcC5uU9+PB3tnAK6v9lW0LOMMf+FJpC5
# 16JB51kAIkf3QTVn1KXFq2/FGUKZnMrvkgBlmfJZWDuUD9oGMmlVr/4kaQoL9+zz
# y24k28R153zoDB8JlbTBMiCDA4jGniAXiyMBYD6DepLWSUXOiBVPJ90emcoM/slU
# AH64f2MmcxuSr86kp5MZCDjimRDnK+gsG0Qze2KNDHD8ykqTzMk6feq9jQhGNkQ8
# qCJnHRH1ya1ReHyrrImTiGTRBzzn2zMwQTFMktK7qfHjAZon/gfeEK/AsA4mCZ7r
# vMzuaSq1XmtrMeH91486FnhBY53SKP78mJhpmdg5wGIby16XBgvih0XaGZYB7fLX
# WxUykeELOQkPaekxujgXFJJj5sqiF1GET2qsYXbB5FHvKw75oJmH1OgbRBoimW0I
# Pq+GM+fXvn0RHVuXI6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgxNDU4
# MjVaMC8GCSqGSIb3DQEJBDEiBCDLXn2Fzi+DLC4Rxa6Yvs9XcWLoT6cFzsQpgQHB
# HNr5SzANBgkqhkiG9w0BAQEFAASCAgAgzEaa9qd5T4zOVz8c0QSjJ5yRPYhIF6DG
# FoIOQbde1NfMOraNiDVp7hkYX51Olor98mz8fHBxY2WZJxJe8rLE6lGOnB3JGhX7
# xN5IHISicKdnXETQSBquhTGkRVhk1HOonONUkyXJmjzFIC9PhypTWaSEAhMtmR95
# QbpwSL3IDXg2mlhrwRlp12buTi/3C7CHuj3O//nR6cs1aOkvcoc57UkKeYwBGfxR
# 9L4prdtCGxpTJm0y2WCKiLozOij5EYlO/HgbbA+zXY4psabejXGfz2vXrmOEjT5l
# 2Y3lJL/mqMbCJQLkk7qZPNrEhKMUirq7tzCRwwElYdeAmbZFRQ9uo4LoaJJJfe6D
# VQJi/GFo3S8MreaHz7Rd2ap1w8Jiacj5tNebFzASyyHOBc3jggzFPFzbzU1xPBcW
# yetfzah2iksV/gopzr3Uvog8x+DCnNRR2beAMlQSxDnZLRr6K6QW1E4RDgayh5Y5
# sm/ZTuVqYQgCCOj5J0XqhUUpf2ZmI95/M/CsCXrp/8keyms3tjbIopmuQ9cMB6iA
# wSCWPSPiFvoVwnAOUkd5G+0sdNogczweHtyitRxbK1mr8JKOxsU/CUh+S0Wyfj8P
# KBIw/mecd41HpCdeEzMUrp0oJ7Y6CYF6nzw6rUKO1Qyy2qravgVun+J+LRvbEg5h
# nkegI1w88A==
# SIG # End signature block
