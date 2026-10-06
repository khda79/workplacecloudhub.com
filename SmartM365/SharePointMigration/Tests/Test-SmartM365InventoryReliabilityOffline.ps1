<#
.SYNOPSIS
    Regression checks for inventory data fidelity, paged reads and workbook snapshots.
.VERSION
    1.0.1
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
function Write-SPOReadRetry { param([string]$Message) }
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
Import-TestFunction $permission @('Connect-ToSPOWeb','Get-PnPGroupTitle','Get-AssociatedWebGroupNames','Invoke-SPORead','Add-SPOInheritanceRead','Get-SPOPageInheritance','Write-SPOItemError','Export-SPOPermissionPage','Export-ItemPermissionInventory','Write-InventoryError')
$script:SPOPermissionConnection='fixture'
$script:SPOInheritanceLoad=$null; $script:SPOInheritanceItemType=$null
$script:AssociatedWebGroupCache=@{}
$script:PropertyCalls=0; $script:GroupFail=$false
function Get-PnPProperty {
    [CmdletBinding()] param($ClientObject,[string]$Property,$Connection)
    $script:PropertyCalls++
    if ($Property -eq 'Title') { if ($script:GroupFail) { throw 'Group access denied.' }; return $ClientObject.Title }
    if ($Property -eq 'HasUniqueRoleAssignments') {
        if (($ClientObject.Id -eq 3 -and $script:Missing) -or $ClientObject.Id -eq $script:MissingId) { throw 'Item does not exist.' }
        return ($ClientObject.Id % 2 -eq 0)
    }
    return $ClientObject.$Property
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
        [CmdletBinding()] param($List,[int]$PageSize,[string[]]$Fields,$Connection,[scriptblock]$ScriptBlock,[int]$Id)
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
    'PASS: durations, group fidelity/cache, token contexts, bounded reads, size-limited streamed inheritance, missing-item evidence, stable workbook snapshots and batch receipts.'
}
finally {
    Get-ChildItem -LiteralPath $temp -File | Remove-Item -Force
    Remove-Item -LiteralPath $temp
}


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDqMlwKvN9L6tQx
# 9NJ7yM9uXDuKY2GwlY8J14xxtgzUh6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCvRVNTHy9nHIqrAerh7Rqy
# aPp6sql5AvLk6AatfsoKmjANBgkqhkiG9w0BAQEFAASCAYCZlDxwbI89G+nGTg7H
# Lcnup49g7iD+VnqnJFaiFz54Ymw3nitSg8LlYMLGatvs4woDnPsPyXjXGZ2V5JbV
# S87BbbeEoejlO1rrV3gLe/1FvpkJPZRbS8sHU0lFc97dKBkBzjVjbRu8Luv/ny1v
# EQBu6KFj+Gxy2RQ81HIYTObhQUYbNe0U0oY254Hwj8oDCFu91JINk7cehXLKDGKu
# vx6PIn3ZbwCrvqzZctPtLso2+zBpTaRyiBEVQg8pqe0RRI53eg84tj0fkDGgOqIN
# m9AsBqWheNQaIOb9Z26vrdLlVjactGzdjf4+SJyGdOyZYFqUtcebaPGtaEUHQRE5
# 6IPS1TFG/F+aGsgMt02Ba8SfTKUYsGhgcDBhDgG8d+ACXbq2Z52G8OidL62JTqoV
# gha0gCpLTsgG4jBOhSWOru/fBDw1nRtuzY0nsPYhi5OLDVFUwpEHNjcITh4vrjSb
# 0bKc45qPr5ZPnNqgGKeewMSWxE9YsNILfZ1okfg8b+ezm9Q=
# SIG # End signature block
