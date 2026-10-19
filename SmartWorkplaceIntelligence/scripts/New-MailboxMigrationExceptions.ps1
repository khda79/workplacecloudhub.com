# Version 0.1.0: immutable planning snapshot and operator-confirmed retention identities.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$OutputRoot,
    [Parameter(Mandatory)][string]$TenantKey,
    [string]$NativeEvidenceAuditPath,
    [switch]$Enabled,
    [string]$BindingsPath,
    [string]$PythonCommand = 'python'
)
$ErrorActionPreference = 'Stop'
$python = Get-Command $PythonCommand -CommandType Application -ErrorAction Stop | Select-Object -First 1
$version = & $python.Source -B -c 'import sys; print(".".join(map(str,sys.version_info[:3])))'
if ($LASTEXITCODE -ne 0 -or [version]$version -lt [version]'3.10') { throw 'Python 3.10+ is required for mailbox migration exceptions.' }
$arguments = @('-B', (Join-Path $PSScriptRoot 'workplace_mailbox_exceptions.py'), '--data-root', $DataRoot, '--output-root', $OutputRoot, '--tenant', $TenantKey)
if ($Enabled) { $arguments += '--enabled' }
if ($NativeEvidenceAuditPath) { $arguments += @('--audit', $NativeEvidenceAuditPath) }
if ($BindingsPath) { $arguments += @('--bindings', $BindingsPath) }
& $python.Source @arguments
if ($LASTEXITCODE -ne 0) { throw "Mailbox migration exception preparation failed (exit $LASTEXITCODE)." }
