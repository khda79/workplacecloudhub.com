#Requires -Version 5.1
<#
.SYNOPSIS
Synthetic configuration lock, retry and diagnostic regression tests.
.VERSION
1.0.0
.NOTES
Only newly created temporary files are used. No configuration initialization,
collector, tenant, real share, authentication or cloud operation is executed.
#>
[CmdletBinding()]
param([string]$FixtureRoot = (Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonRetry-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../Config/SmartM365-TenantContext.ps1')
$module = Get-Module SmartM365.JsonTransport
$originalPolicy = & $module { (Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock }
$root = [IO.Path]::GetFullPath($FixtureRoot)
if (Test-Path -LiteralPath $root) { throw 'Use a new fixture directory; existing data is never accepted.' }
New-Item -ItemType Directory -Path $root | Out-Null
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Fixture([string]$Name) {
    $path = Join-Path $root ($Name + '.local.json.txt')
    [IO.File]::WriteAllText($path, '{"Setting":"private","Nested":{"Keep":"custom"},"Extra":42}')
    [IO.File]::WriteAllText(($path + '.template'), '{"Setting":"default","Nested":{"Keep":"default","Add":7},"Added":true}')
    return $path
}
function Failure([scriptblock]$Action) {
    try { & $Action | Out-Null } catch { return $_.Exception.Message }
    throw 'Expected synthetic failure did not occur.'
}
if (-not ('SmartM365.Tests.JsonRetryLease' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Threading;
namespace SmartM365.Tests {
    public sealed class JsonRetryLease : IDisposable {
        private FileStream stream;
        private Thread worker;
        public Exception Error;
        public static JsonRetryLease Hold(string path, int milliseconds, string update, bool exclusive) {
            var lease = new JsonRetryLease();
            lease.stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                exclusive ? FileShare.None : FileShare.ReadWrite);
            if (milliseconds > 0) {
                lease.worker = new Thread(delegate() {
                    try {
                        Thread.Sleep(milliseconds);
                        if (!String.IsNullOrEmpty(update)) File.WriteAllText(path, update);
                    } catch (Exception e) { lease.Error = e; }
                    finally { lease.stream.Dispose(); }
                });
                lease.worker.IsBackground = true;
                lease.worker.Start();
            }
            return lease;
        }
        public void Dispose() {
            if (worker != null) worker.Join();
            stream.Dispose();
        }
    }
}
'@
}
try {
    & $module { function script:Get-SmartM365JsonTransportPolicy { @{ Mode='Readers'; QualifiedUncRoots=@(); QualifiedLocalDataRoots=@() } } }
    $path = Fixture 'nominal'
    $config = Read-SmartM365JsonConfig $path -Required
    Check ($config.Setting -eq 'private' -and $config.Nested.Keep -eq 'custom' -and $config.Extra -eq 42) 'Private values changed.'
    Check ($config.Added -and $config.Nested.Add -eq 7) 'Missing keys were not added.'
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $null = Read-SmartM365JsonConfig $path -Required
    Check ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Already enriched configuration was rewritten.'
    Check ((Get-Command Read-SmartM365JsonConfig).Parameters.ContainsKey('ConfigurationWriteTimeoutSeconds')) 'Configuration wait budget is not configurable.'

    $path = Fixture 'released'
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 700, $null, $false)
    try {
        $records = @(Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 3 6>&1)
        $config = $records | Where-Object { $_ -is [Collections.IDictionary] }
        Check ($config.Added -and $config.Setting -eq 'private') 'Released lock did not permit enrichment.'
        $text = ($records | Where-Object { $_ -is [Management.Automation.InformationRecord] } | Out-String)
        Check ($text -match 'JSON destination locked' -and $text -match 'retry in' -and $text -match 'timeout=3s') 'Retry progress was not logged.'
        Check ($text -match 'JSON replacement succeeded after') 'Recovery was not logged.'
    } finally { $lease.Dispose() }
    Check ($null -eq $lease.Error) 'Synthetic release worker failed.'

    # Prove that configuration enrichment survives a lock beyond the former 10s budget.
    $path = Fixture 'default-budget'
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 11000, $null, $false)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try { $config = Read-SmartM365JsonConfig $path -Required 6>$null }
    finally { $lease.Dispose() }
    Check ($config.Added -and $null -eq $lease.Error) 'Default configuration budget did not survive an 11s lock.'
    Check ($watch.Elapsed.TotalSeconds -ge 10 -and $watch.Elapsed.TotalSeconds -lt 25) 'Default retry recovery exceeded its expected test window.'

    $path = Fixture 'timeout'
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 0, $null, $false)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try { $message = Failure { Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 1 6>$null } }
    finally { $lease.Dispose() }
    $watch.Stop()
    Check ($message -match 'Atomic replacement timed out' -and $message -like "*$path*" -and $message -match 'during template update') 'Timeout diagnosis lost its phase or path.'
    Check ($message -notmatch 'not valid JSON|Invalid JSON syntax') 'A lock was reported as invalid JSON.'
    Check ($watch.Elapsed.TotalSeconds -ge 0.9 -and $watch.Elapsed.TotalSeconds -lt 5) 'Replacement timeout was not bounded.'
    Check ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Timed-out write changed the configuration.'

    $path = Fixture 'zero-budget'
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 0, $null, $false)
    $watch.Restart()
    try { $message = Failure { Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 0 6>$null } }
    finally { $lease.Dispose() }
    Check ($message -match 'timed out') 'Zero retry budget did not fail explicitly.'
    Check ($watch.Elapsed.TotalSeconds -lt 1) 'Zero retry budget waited.'
    Check ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Zero-budget failure changed configuration.'

    $path = Fixture 'concurrent'
    $concurrent = '{"Setting":"changed-by-other-writer","Extra":99}'
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 600, $concurrent, $false)
    try { $message = Failure { Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 3 6>$null } }
    finally { $lease.Dispose() }
    Check ($null -eq $lease.Error) 'Synthetic concurrent writer failed.'
    Check ($message -match 'refusing a lost update') 'A change during the wait was not detected.'
    Check ([IO.File]::ReadAllText($path) -ceq $concurrent) 'A concurrent update was overwritten.'

    $path = Fixture 'exclusive'
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 500, $null, $true)
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes('{"Setting":"private","Added":true}')
        $null = Write-SmartM365JsonBytesAtomically -Path $path -Bytes $bytes -ExpectedSHA256 $hash -ReplacementTimeoutSeconds 3 -Validate { param($d) if ($d.Setting -ne 'private') { throw 'Wrong fixture.' } } 6>$null
    } finally { $lease.Dispose() }
    Check ((Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).Added) 'A transient exclusive lock during hash verification was not retried.'

    $path = Fixture 'invalid'
    [IO.File]::WriteAllText($path, '{')
    $watch.Restart()
    $message = Failure { Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 3 }
    Check ($message -match 'Invalid JSON syntax' -and $message -match 'during read') 'Invalid JSON lost its syntax diagnosis.'
    Check ($watch.Elapsed.TotalSeconds -lt 1 -and [IO.File]::ReadAllText($path) -eq '{') 'Invalid JSON was retried or overwritten.'

    $path = Fixture 'invalid-template'
    [IO.File]::WriteAllText(($path + '.template'), '{')
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $message = Failure { Read-SmartM365JsonConfig $path -Required }
    Check ($message -match 'template update' -and $message -notmatch 'not valid JSON|Invalid JSON syntax') 'A bad template was attributed to runtime JSON syntax.'
    Check ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'A bad template changed runtime configuration.'

    $path = Join-Path $root 'directory.json.txt'
    New-Item -ItemType Directory -Path $path | Out-Null
    $watch.Restart()
    $message = Failure { Write-SmartM365JsonBytesAtomically -Path $path -Bytes ([Text.Encoding]::UTF8.GetBytes('{}')) -ReplacementTimeoutSeconds 3 -Validate { param($d) if ($null -eq $d) { throw 'Missing fixture.' } } 6>$null }
    Check ($message -match 'Atomic replacement failed' -and $message -notmatch 'timed out') 'A permanent destination error was treated as a transient lock.'
    Check ($watch.Elapsed.TotalSeconds -lt 1 -and (Test-Path -LiteralPath $path -PathType Container)) 'Permanent destination failure was retried or changed its target.'

    foreach ($code in @(32, 33, 5, 19)) {
        $native = [IO.IOException]::new('Synthetic native error.', (-2147024896 + $code))
        $wrapped = [InvalidOperationException]::new('Synthetic wrapper.', $native)
        $actual = & $module { param($e) Get-SmartM365JsonLockErrorCode $e } $wrapped
        $expectedCode = if ($code -in @(32, 33)) { $code } else { 0 }
        Check ($actual -eq $expectedCode) 'Native error classification retried a permanent error or missed a lock.'
    }
    [pscustomobject]@{ Passed=$script:passed; FixtureRoot=$root; Evidence='Synthetic local files only; no preparation, private configuration, real share or cloud access.' }
} finally {
    & $module { param($body) Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $body } $originalPolicy
}
