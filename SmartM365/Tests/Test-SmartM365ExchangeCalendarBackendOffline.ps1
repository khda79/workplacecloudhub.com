#Requires -Version 5.1
<# .SYNOPSIS Synthetic calendar session, backend isolation and terminal publication tests. .VERSION 1.0.1 #>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$savedGlobals=@{}
foreach($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId',
    'SmartM365CsvValidationRules','SmartM365RequireCsvValidationRules','RetentionMaxCSV','csvFilePath1','csvFilePath2','csvFilePath3','csvGeneratedPaths')) {
    $variable=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $savedGlobals[$name]=@{Exists=($null -ne $variable);Value=$(if($variable){$variable.Value}else{$null})}
}
try {
$global:SmartM365CsvValidationRules=@{}
$global:SmartM365RequireCsvValidationRules=$true
$global:csvGeneratedPaths=$null
$global:SmartM365TenantId=''
$smartRoot = Split-Path $PSScriptRoot -Parent
$collector = Join-Path $smartRoot 'SmartInventory/ExchangeInventory/OnPremises/CalendarPermissions/SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
$source = Get-Content -LiteralPath $collector -Raw
$tokens=$null; $parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$parseErrors)
$checks=0
function Check([bool]$Condition,[string]$Message) { if(-not $Condition){throw $Message}; $script:checks++ }
Check ($parseErrors.Count -eq 0) 'Collector parser failed'
foreach($name in @('Invoke-Quiet','Initialize-SmartM365CalendarExchangeConnection','Get-SmartM365CalendarFailureCategory',
    'Get-SmartM365CalendarBackendName','New-SmartM365CalendarLookupException','Test-SmartM365CalendarBackendPreflight',
    'New-SmartM365CalendarErrorRecord','Get-CalendarFoldersSafe','Try-GetFolderPermission','Add-NoPermissionRow',
    'Add-CalendarResultRows','Assert-SmartM365CalendarQualification','New-CalendarProgressTracker','Get-CalendarProgressLine')) {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    Check ($null -ne $node) ('Missing production function: '+$name)
    . ([scriptblock]::Create($node.Extent.Text))
}
function WriteLog {param($Message,$Level) $script:messages.Add([string]$Message)}
function Test-HasCommand {param($Name) return $true}
function RoundTrip($Value) { [Management.Automation.PSSerializer]::Deserialize([Management.Automation.PSSerializer]::Serialize($Value)) }
function New-Mailbox([string]$Identity='synthetic-one',[string]$Database='SYNTHETIC-DB01') {
    RoundTrip ([pscustomobject]@{Identity=$Identity;PrimarySmtpAddress=($Identity+'@example.test');UserPrincipalName=($Identity+'@example.test');Guid='11111111-1111-1111-1111-111111111111';ServerName='SYNTHETIC-EX01';Database=$Database})
}
function Reset-Backend {
    $script:UnavailableCalendarBackends=@{}; $script:CalendarBackendPreflightResults=@{}
    $script:BackendPreflightTimeoutSeconds=17; $script:statisticsCalls=0; $script:probeCalls=0
    $script:messages=New-Object 'System.Collections.Generic.List[string]'
    $script:probeMode='Healthy'; $script:statisticsMode='Healthy'; $script:permissionMode='Explicit'
}
function Test-MAPIConnectivity {
    [CmdletBinding()] param($Database,$Server,$PerConnectionTimeout,$AllConnectionsTimeout)
    $script:probeCalls++
    Check ($PerConnectionTimeout -eq 17 -and $AllConnectionsTimeout -eq 17) 'Timeout parameters were lost'
    if($script:probeMode -eq 'ThrowDenied'){throw 'MAPI failed: MapiExceptionNoAccess; information store unavailable; 0x80070005'}
    if($script:probeMode -eq 'ThrowTimeout'){throw 'MAPI RPC connection timed out'}
    if($script:probeMode -eq 'Denied'){return RoundTrip ([pscustomobject]@{Result='Failure';Error='MapiExceptionNoAccess; information store unavailable; 0x80070005'})}
    if($script:probeMode -eq 'Unavailable'){return RoundTrip ([pscustomobject]@{Result='Failure';Error='Information Store is unavailable'})}
    if($script:probeMode -eq 'Unknown'){return RoundTrip ([pscustomobject]@{Result='Failure';Error='Unclassified synthetic probe failure'})}
    if($script:probeMode -eq 'Empty'){return}
    RoundTrip ([pscustomobject]@{Result='Success';Error=''})
}
function Get-MailboxFolderStatistics {
    [CmdletBinding()] param($Identity,$FolderScope)
    $script:statisticsCalls++
    Check ($FolderScope -eq 'Calendar') 'Wrong folder scope'
    if($script:statisticsMode -eq 'Denied'){throw 'MapiAccessDeniedException: access is denied'}
    if($script:statisticsMode -eq 'Unavailable'){throw 'Information Store is unavailable'}
    if($script:statisticsMode -eq 'Empty'){return}
    if($script:statisticsMode -eq 'TextCounts') {
        RoundTrip ([pscustomobject]@{FolderPath='/Small';FolderType='Calendar';ItemsInFolder='9'})
        RoundTrip ([pscustomobject]@{FolderPath='/Unknown';FolderType='Calendar';ItemsInFolder='unknown'})
        RoundTrip ([pscustomobject]@{FolderPath='/Large';FolderType='Calendar';ItemsInFolder='100'})
        return
    }
    RoundTrip ([pscustomobject]@{FolderPath='/Calendrier';FolderType='Calendar';ItemsInFolder=0})
}
function Get-MailboxFolderPermission {
    [CmdletBinding()] param($Identity)
    if($script:permissionMode -eq 'Denied'){throw 'MapiExceptionNoAccess: access is denied'}
    RoundTrip ([pscustomobject]@{User='Default';AccessRights=@('AvailabilityOnly')})
    RoundTrip ([pscustomobject]@{User='Anonymous';AccessRights=@('None')})
    if($script:permissionMode -eq 'Explicit'){RoundTrip ([pscustomobject]@{User='synthetic-reader';AccessRights=@('Reviewer')})}
}
foreach($message in @('MapiExceptionNoAccess','MapiAccessDeniedException','0x80070005','access is denied','permission denied','MAPI failed: information store unavailable; MapiExceptionNoAccess')) {
    Check ((Get-SmartM365CalendarFailureCategory $message) -eq 'AccessDenied') 'Access denied classified as backend outage'
}
foreach($mode in @('Denied','ThrowDenied','Unknown','Empty')) {
    Reset-Backend; $script:probeMode=$mode
    $mbx=New-Mailbox
    $folders=@(Get-CalendarFoldersSafe $mbx $true)
    $null=Get-CalendarFoldersSafe (New-Mailbox 'synthetic-two') $true
    Check ($folders.Count -eq 1 -and $folders[0].ItemsInFolder -eq 0) 'Remote empty folder was lost'
    Check ($script:probeCalls -eq 1 -and $script:statisticsCalls -eq 2) 'Inconclusive probe prevented mailbox reads or repeated per database'
    Check ($script:UnavailableCalendarBackends.Count -eq 0) 'Probe denial/unknown contaminated backend cache'
    Check ($script:messages.Count -gt 0) 'Inconclusive probe lost its diagnostic'
}
foreach($mode in @('Unavailable','ThrowTimeout')) {
    Reset-Backend; $script:probeMode=$mode; $caught=$null
    try {Get-CalendarFoldersSafe (New-Mailbox) $true|Out-Null}catch{$caught=$_}
    Check ($caught.Exception.Data['SmartM365Category'] -eq 'BackendUnavailable') 'Real outage classification lost'
    Check ($script:UnavailableCalendarBackends.Count -eq 1 -and $script:UnavailableCalendarBackends.ContainsKey('database:synthetic-db01')) 'Outage must be cached only for its database'
    $caught=$null
    try {Get-CalendarFoldersSafe (New-Mailbox 'synthetic-two') $true|Out-Null}catch{$caught=$_}
    Check ($caught.Exception.Data['SmartM365Category'] -eq 'BackendPreviouslyUnavailable' -and $script:probeCalls -eq 1) 'Same database not bounded'
    $script:probeMode='Healthy'
    Check (@(Get-CalendarFoldersSafe (New-Mailbox 'synthetic-three' 'SYNTHETIC-DB02') $true).Count -eq 1) 'Other database on same server was skipped'
}
Reset-Backend; $script:statisticsMode='Denied'
foreach($id in @('synthetic-one','synthetic-two')) {
    $caught=$null; try {Get-CalendarFoldersSafe (New-Mailbox $id) $true|Out-Null}catch{$caught=$_}
    Check ($caught.Exception.Data['SmartM365Category'] -eq 'AccessDenied') 'Target denial diagnostic lost'
}
Check ($script:statisticsCalls -eq 2 -and $script:UnavailableCalendarBackends.Count -eq 0) 'Target denial poisoned database'
Reset-Backend; $script:statisticsMode='Unavailable'
$caught=$null; try {Get-CalendarFoldersSafe (New-Mailbox) $true|Out-Null}catch{$caught=$_}
Check ($caught.Exception.Data['SmartM365Category'] -eq 'BackendUnavailable' -and $script:UnavailableCalendarBackends.Count -eq 1) 'Statistics outage not scoped to database'
$script:statisticsMode='Healthy'
Check (@(Get-CalendarFoldersSafe (New-Mailbox 'synthetic-two' 'SYNTHETIC-DB02') $true).Count -eq 1) 'Statistics outage skipped other database'

Reset-Backend; $script:statisticsMode='TextCounts'
$chosen=@(Get-CalendarFoldersSafe (New-Mailbox) $true)
Check ($chosen.Count -eq 1 -and $chosen[0].FolderPath -eq '/Large') 'Text counts were sorted lexically or unknown became zero'

# Exercise the production connection function; no Exchange or network calls.
function Start-SmartM365ExchangeOnPremSession {
    param([string[]]$RequiredCommands)
    $script:sessionStarts++
    Check ($RequiredCommands.Count -eq 5 -and $RequiredCommands -contains 'Get-MailboxFolderPermission') 'Wrong remote import allowlist'
    if($script:connectionFailure){throw 'Synthetic remote connection failure; no LocalSnapin fallback'}
    [pscustomobject]@{Endpoint='http://synthetic.example.test/PowerShell/'}
}
function Set-ADServerSettings {[CmdletBinding()]param([bool]$ViewEntireForest) Check $ViewEntireForest 'Forest scope missing'}
function Get-PSSnapin {[CmdletBinding()]param($Name,[switch]$Registered) if($Registered){[pscustomobject]@{Name=$Name}}}
function Add-PSSnapin {[CmdletBinding()]param($Name) $script:snapins++}
$script:sessionStarts=0; $script:snapins=0; $script:connectionFailure=$false; $script:ExchangeOnPremContext=$null
Initialize-SmartM365CalendarExchangeConnection Remote
Initialize-SmartM365CalendarExchangeConnection Remote
Check ($script:sessionStarts -eq 1 -and $script:snapins -eq 0) 'Remote mode must open one session and no snap-in'
$script:ExchangeOnPremContext=$null; $script:connectionFailure=$true; $caught=$null
try {Initialize-SmartM365CalendarExchangeConnection Remote}catch{$caught=$_}
Check ($null -ne $caught -and $script:snapins -eq 0) 'Remote failure silently fell back'
Initialize-SmartM365CalendarExchangeConnection LocalSnapin
Check ($script:snapins -eq 1 -and $script:sessionStarts -eq 2) 'LocalSnapin no longer uses existing loading path'
Check ($source -match 'RequireExchangeOnPrem:\(\$script:ExchangeConnectionMode -ne ''Remote''\)') 'Remote preflight still loads native snap-in'
Check ($source -match 'Stop-SmartM365ExchangeOnPremSession -State \$script:ExchangeOnPremContext') 'Owned session cleanup absent'
Check ($source -notmatch '\bInvoke-Command\b|\bForEach-Object\s+-Parallel\b|\bStart-Job\b') 'Remote scriptblock or concurrent Exchange execution introduced'

# Execute the actual outer finally, including when logging fails before cleanup.
$terminalTry=$ast.Find({param($n) $n -is [Management.Automation.Language.TryStatementAst] -and $null -ne $n.Finally -and $n.Finally.Extent.Text -match 'Stop-SmartM365ExchangeOnPremSession'},$true)
Check ($null -ne $terminalTry) 'Terminal session cleanup missing'
$cleanupText=$terminalTry.Finally.Extent.Text
$cleanupText=$cleanupText.Substring(1,$cleanupText.Length-2).Replace('if($smartM365RunGuard.ExitCode -eq 130){exit 130}','')
function Stop-SmartM365ExchangeOnPremSession {param($State) $script:sessionStops++}
function Complete-SmartM365InventoryRun {param($State) $script:runCompletions++}
$script:sessionStops=0; $script:runCompletions=0; $smartM365RunGuard=[pscustomobject]@{ExitCode=0}
$script:ExchangeOnPremContext=[pscustomobject]@{Timer=[Diagnostics.Stopwatch]::StartNew()}
. ([scriptblock]::Create($cleanupText))
Check ($script:sessionStops -eq 1 -and $script:runCompletions -eq 1) 'Terminal cleanup leaked owned resources'
function WriteLog {param($Message,$Level) throw 'Synthetic logging failure'}
try {. ([scriptblock]::Create($cleanupText))}catch{Check ($_.Exception.Message -match 'Synthetic logging failure') 'Unexpected cleanup exception'}
Check ($script:sessionStops -eq 2 -and $script:runCompletions -eq 2) 'Logging failure prevented cleanup'
function WriteLog {param($Message,$Level) $script:messages.Add([string]$Message)}
$script:ExchangeOnPremContext=$null
$caught=$null;try{Initialize-SmartM365CalendarExchangeConnection Invalid}catch{$caught=$_}
Check ($null -ne $caught) 'Invalid mode was accepted'

# Run the actual processing/export section, with temporary files and offline I/O stubs.
$processingStart=$source.IndexOf('# ------------------------- Processing Loop')
$processingEnd=$source.IndexOf('Write-Host "`n=== SUMMARY ==="',$processingStart)
Check ($processingStart -gt 0 -and $processingEnd -gt $processingStart) 'Processing section markers missing'
$processing=[scriptblock]::Create($source.Substring($processingStart,$processingEnd-$processingStart))
function Add-SmartM365MaxItemsSuffixToBaseName {param($BaseFileName) $BaseFileName}
function Get-ScriptLocalConfigValue {param($Config,$Name,$DefaultValue) if($Name -eq 'LatestCsvFolderPath'){$script:fixtureLatest}else{$DefaultValue}}
# Use real CSV validation/publication code; mock only external context, receipts and uploads.
$compatPath=Join-Path $smartRoot 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1'
$compatAst=[Management.Automation.Language.Parser]::ParseFile($compatPath,[ref]$tokens,[ref]$parseErrors)
Check ($parseErrors.Count -eq 0) 'Compatibility parser failed'
foreach($name in @('Write-SmartM365CsvAtomically','Copy-SmartM365FileAtomically','ExportAndCopyCsv',
    'Add-SmartM365TenantKeyToCsvData','Get-SmartM365CoreContextValue','Assert-SmartM365CsvDataCompleteness',
    'Get-SmartM365CsvValidationBaseName','Get-SmartM365CsvValidationRule','Initialize-SmartM365DefaultCsvValidationRules','Add-SmartM365CsvValidationRule')) {
    $node=$compatAst.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    Check ($null -ne $node) ('Missing CSV helper: '+$name)
    . ([scriptblock]::Create($node.Extent.Text))
}
$global:SmartM365TenantKey='test'; $global:SmartM365OrganizationKey='synthetic'; $global:SmartM365EnvironmentKey='test'
Initialize-SmartM365DefaultCsvValidationRules
function Limit-SmartM365RowsForMaxItems {param($Data) $Data}
function Test-SmartM365MaxItemsMode {return $false}
function Start-SmartM365SourcePublication {param($Path)}
function Register-SmartM365SourceCsv {param($Path)}
function Invoke-SmartM365SharePointCsvUpload {param($LocalFilePath)}
function Remove-SmartM365SharePointFile {param($LocalFilePath) $script:staleRemovals++}
function RemoveOldFiles {param($FolderPath,$FilePattern,$MaxFiles) $script:retentionCalls++}
function Publish-SmartM365CalendarWeeklyHistory {param($SourceCsvPaths,$Config,$FallbackRootPath) $script:historyCalls++}
$fixtureRoot=Join-Path ([IO.Path]::GetTempPath()) ('exchange-calendar-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixtureRoot
try {
    foreach($primary in @($true,$false)) {
        foreach($case in @('Success','StatisticsDenied','PermissionsDenied','NoFolders','KnownEmpty')) {
            Reset-Backend
            $runPath=Join-Path $fixtureRoot ([guid]::NewGuid().ToString('N')); $null=New-Item -ItemType Directory -Path $runPath
            $script:fixtureLatest=$runPath; $OutputPath=$runPath; $logTextFile=Join-Path $runPath 'diagnostic.log'
            $PrimaryOnly=$primary; $EmitNoPermRow=($case -ne 'KnownEmpty'); $ScriptLocalConfig=[pscustomobject]@{}
            $mailboxes=@((New-Mailbox)); $total=1; $processed=0; $qualifiedCalendars=0
            $results=New-Object 'System.Collections.Generic.List[object]'; $errors=New-Object 'System.Collections.Generic.List[object]'
            $collectionTimer=[Diagnostics.Stopwatch]::StartNew(); $script:historyCalls=0; $script:staleRemovals=0
            $script:retentionCalls=0; $global:RetentionMaxCSV=30
            if($case -eq 'StatisticsDenied'){$script:statisticsMode='Denied'}
            if($case -eq 'PermissionsDenied'){$script:permissionMode='Denied'}
            if($case -in @('NoFolders','KnownEmpty')){$script:permissionMode='DefaultOnly'}
            if($case -eq 'NoFolders'){$script:statisticsMode='Empty';$script:permissionMode='Denied'}
            $canonical=Join-Path $runPath 'Exchange_OnPrem_MailboxCalendarPermissions_AllDomains.csv'
            Set-Content -LiteralPath $canonical -Value 'previous-qualified-snapshot'
            $caught=$null
            try {. $processing}catch{$caught=$_}
            if($case -in @('Success','KnownEmpty')) {
                Check ($null -eq $caught -and $qualifiedCalendars -eq 1) ('Qualified run failed: '+$case+'; '+$caught)
                Check ($script:historyCalls -eq 1) 'Qualified run did not publish history'
                Check ((Get-Content -LiteralPath $canonical -Raw) -notmatch 'previous-qualified-snapshot') 'Qualified snapshot not replaced'
                if($case -eq 'KnownEmpty'){
                    Check (@(Import-Csv -LiteralPath $canonical).Count -eq 0) 'Known empty permissions became invented rows'
                    Check ((Get-Content -LiteralPath $canonical -Raw) -match '"TenantKey"') 'Empty tenant schema missing'
                    Check (-not $global:SmartM365CsvValidationRules['Exchange_OnPrem_MailboxCalendarPermissions_AllDomains'].AllowEmptyDataset) 'Empty rule override leaked'
                }
                else {Check (@(Import-Csv -LiteralPath $canonical)[0].AccessRights -eq 'Reviewer') ('Deserialized permission projection failed: primary='+$primary+'; csv='+(Get-Content -LiteralPath $canonical -Raw))}
            }
            else {
                Check ($null -ne $caught -and $caught.Exception.Message -match 'no qualified calendar') ('Unqualified run incorrectly succeeded: '+$case)
                Check ($script:historyCalls -eq 0 -and $script:staleRemovals -eq 0 -and $script:retentionCalls -eq 0) 'Failure continued to history or stale diagnostic cleanup'
                Check ((Get-Content -LiteralPath $canonical -Raw).Trim() -eq 'previous-qualified-snapshot') 'Failure replaced previous qualified canonical CSV'
                $diagnostic=Join-Path $runPath 'Exchange_OnPrem_MailboxCalendarPermissions_Errors.csv'
                Check (Test-Path -LiteralPath $diagnostic) 'Failure lost diagnostic CSV'
                $diagnosticRows=@(Import-Csv -LiteralPath $diagnostic)
                Check ($diagnosticRows.Count -gt 0) 'Failure diagnostic was empty'
                if($case -ne 'NoFolders'){Check ($diagnosticRows[0].Category -eq 'AccessDenied') 'Access denied lost from diagnostic CSV'}
            }
        }
    }
    Check (@($script:messages | Where-Object { $_ -like 'Calendar permissions progress: 1/1 mailboxes (100.0%);*ETA done.' }).Count -gt 0) 'Processing loop did not log the final progress line'
    Assert-SmartM365CalendarQualification -ExpectedMailboxes 0 -QualifiedCalendars 0
    Check $true 'Zero expected mailboxes must not fail qualification'
}
finally {
    # Delete only this test-created tree, never collector output.
    if([IO.Path]::GetFullPath($fixtureRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
Write-Host ("Calendar backend offline checks passed: {0}" -f $checks)
}
finally {
    foreach($name in $savedGlobals.Keys) {
        if($savedGlobals[$name].Exists){Set-Variable -Name $name -Value $savedGlobals[$name].Value -Scope Global}
        else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPsKpmWHRFEOb4
# rKy7UdwKEY9rifvyKJGU+kMIaUa7OaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIE8sX3V7qw50npUPaq2Z9GAEp37UfKKMWDbbOBQG7r+LMA0GCSqG
# SIb3DQEBAQUABIIBgGLKkcoqgkKbqV34HL1bBPg1/se5YR4W2Cjz3lHrejI/w1L+
# 6l6uY6fvy+DM1Ik4SOtOZZHwi/H6KDmhC15P9ScNmhuvpklTREeBDgkeqqAeCzlm
# 8WFLnD9xYGzfHdSEbMse/7kiPWWVy8YIb3ZYOCqoH0p41fSpF42hNMf+fHVc3yng
# Z72CvLAAUUFJmWC+Tp6f1jpaB9vMzIXg5XFjeviqYwl3LXxpOK2TC4yStWLCN7M1
# 4fhTS4JLTOZ1YgUAxkHKK5cuZ659ZUMPU/IU2HRSTT3dIIrnP/bBQIW7SuDfhAWH
# dfVen05qzSh6akmsIjmUZXKdwulyTPJGrZuia3H8Bj3YFFrrZ527kbVElb+z4BTe
# vhltpQ+yRDf3iRGErL4jwgRLuRWX+dDOenmQMxHjC1YF4VkKfhnx7yXZ5cC08il4
# Q6uOQr0H2sSG/0G5nTYAsPtZZexVAiuAWc6gxTTrUWl0Kce1HVWZ5Ma1yHwZpf6g
# tVYUZT/KzroS8cjFr6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzQ0
# MzBaMC8GCSqGSIb3DQEJBDEiBCBIoQFocGv8/UKxiMpM0WfsUmFMnXAMzB1jvDV4
# /EgRfTANBgkqhkiG9w0BAQEFAASCAgBzxExkiakWPL2La9vbxFK/1aBwPj+CE2lk
# 3sEnTh+aBJRw8QNCLYiE8kLh/MzbAaWb56fxuF3yM9faBnsXp07vbH43br79KgnW
# Xr8OHbeqZziert47FnpXetL2IRurgW58x5dzCq0XkEJktm/ZwIQOHpW655krmUJc
# ZFw7I9PBnuaCgwEwG+CsyKvaoduO0Rw3XL8YsiAC3NPRl99KCJ2dcNSLgc4mKLnw
# jvoKHl1rsZgORMz5U8QuR2Hkj0DzuwCk/LE7QTxHGBiOqoTpEvAl1QFu/6CchbPu
# 4lIBk6zAlhvhcCVXtL2vDDY/zrWpzEgPgYv+YYSdY0ShZFNcUam83Y5JF1vSXszA
# PHCiNPRb/HRbBecRsHYytDF3Q5QzrtvIUVEZVXQs9qC8/qirMFH8y8PyzLxUaoRe
# 9sgTYIOjuJHq8s5wqvsio4TaumKZhEL3bAVB/kELask9dMWkILuMmoKck1FyJK/A
# kWTX4XMNqGMIkjsOtIdihaN8RMwFJQ/pNfRDMkFz1sB2KWZEoOPRAsXZ8strfyvM
# CIcXQjnVzuSOLuEFOw2SI8UG9dl979sr0I2o9Kuuv6mJqASoxEYN9Q1yLJGyO5ge
# TJe2Af0Lo2I9rttP0ktnOfhLm4wb7v2/yb5eDtyK7z8cNJws857Yk0A31jlV+KuL
# hRZH7qcOiA==
# SIG # End signature block
