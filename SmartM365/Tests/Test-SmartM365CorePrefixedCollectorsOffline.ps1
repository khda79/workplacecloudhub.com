#Requires -Version 7.0
<#
.SYNOPSIS
Offline check of the collectors that import SmartM365.Core with -Prefix Core (Endpoint Analytics,
Autopatch): their Graph retry and export download paths run against the Core module imported exactly
as in production, without any Core function loaded by hand, and resolve every Core call.
.VERSION
1.0.0
.NOTES
The export download reads a loopback HTTP URL through the real Core helper. Graph calls are replaced by
synthetic errors. No tenant, Graph, storage or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Stand-ins accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Retry waits are recorded instead of slept.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
# Same import as the collectors: prefixed only. No unprefixed Core command may resolve in this process.
Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.99' -Prefix Core -Force -WarningAction SilentlyContinue
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
Check ($null -eq (Get-Command Get-SmartM365GraphErrorStatusCode -ErrorAction SilentlyContinue) -and $null -ne (Get-Command Get-CoreSmartM365GraphErrorStatusCode -ErrorAction SilentlyContinue)) 'The test process does not reproduce the prefixed Core import.'

function Import-CollectorFunction {
    param([Parameter(Mandatory)][string]$RelativePath, [Parameter(Mandatory)][string[]]$Names, [string]$Prelude)
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root $RelativePath), [ref]$null, [ref]$null)
    $text = foreach ($name in $Names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $node) { throw "$name is missing in $RelativePath." }
        $node.Extent.Text
    }
    New-Module -ScriptBlock ([scriptblock]::Create($Prelude + "`n" + ($text -join "`n")))
}
$standIns = @'
$script:Answers = [Collections.Generic.Queue[object]]::new()
$script:Calls = 0
function Start-Sleep { param([int]$Seconds) $null = $Seconds }
function Write-Log { param($Message, $Level) $null = $Message, $Level }
function Write-EALog { param($Message, $Level) $null = $Message, $Level }
function Next-Answer { $script:Calls++; $answer = $script:Answers.Dequeue(); if ($answer -is [Exception]) { throw $answer }; $answer }
function Set-Answers([object[]]$Items) { $script:Answers.Clear(); $script:Calls = 0; foreach ($item in $Items) { $script:Answers.Enqueue($item) } }
'@
$timeout = [System.Threading.Tasks.TaskCanceledException]::new('The request was canceled due to the configured HttpClient.Timeout of 100 seconds elapsing.')
$sdk500 = [Exception]::new('Response status code does not indicate success: InternalServerError (Internal Server Error).')

# A real export package served from a local file:// URL.
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-CorePrefixed-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
try {
    $csv = Join-Path $temporary 'report.csv'
    [IO.File]::WriteAllText($csv, "DeviceId,Name`r`nd1,Device 1`r`n")
    $zip = Join-Path $temporary 'report.zip'
    Compress-Archive -Path $csv -DestinationPath $zip
    # Minimal HTTP server on the loopback address: one response per download, then it stops.
    $zipBytes = [IO.File]::ReadAllBytes($zip)
    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $server = [powershell]::Create()
    [void]$server.AddScript({
        param($Listener, [byte[]]$Body, [int]$Count)
        for ($i = 0; $i -lt $Count; $i++) {
            $client = $Listener.AcceptTcpClient()
            try {
                $stream = $client.GetStream()
                $buffer = [byte[]]::new(8192)
                $null = $stream.Read($buffer, 0, $buffer.Length)
                $header = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 200 OK`r`nContent-Type: application/zip`r`nContent-Length: $($Body.Length)`r`nConnection: close`r`n`r`n")
                $stream.Write($header, 0, $header.Length); $stream.Write($Body, 0, $Body.Length); $stream.Flush()
            } finally { $client.Close() }
        }
    }).AddArgument($listener).AddArgument($zipBytes).AddArgument(2)
    $null = $server.BeginInvoke()
    $zipUrl = 'http://127.0.0.1:{0}/exports/report.zip?sv=synthetic&sig=synthetic' -f ([Net.IPEndPoint]$listener.LocalEndpoint).Port

    # ---- Endpoint Analytics: status, retry and download paths call Core with the Core prefix ----
    $ea = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1' -Names @('Get-EAStatusCode', 'Get-EARetryAfterSeconds', 'Invoke-EAGraphRequest', 'Invoke-EAWebDownload') -Prelude ($standIns + "`n`$MaxRetryCount = 3")
    try {
        $r = & $ea { param($a) Set-Answers $a; $v = Invoke-EAGraphRequest GET 'https://graph.microsoft.com/beta/x' -NoSleep -Invoker { Next-Answer }; [pscustomobject]@{ Value = $v; Calls = $script:Calls } } @($sdk500, 'ok')
        Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Endpoint Analytics did not classify a Graph SDK 500 through the prefixed Core.'
        $r = & $ea { param($a) Set-Answers $a; $v = Invoke-EAGraphRequest GET 'https://graph.microsoft.com/beta/x' -NoSleep -Invoker { Next-Answer }; [pscustomobject]@{ Value = $v; Calls = $script:Calls } } @($timeout, 'ok')
        Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Endpoint Analytics did not retry a timeout through the prefixed Core.'
        $target = Join-Path $temporary 'ea.zip'
        & $ea { param($u, $o) Invoke-EAWebDownload $u $o } $zipUrl $target
        Check ((Get-FileHash -LiteralPath $target).Hash -eq (Get-FileHash -LiteralPath $zip).Hash) 'Endpoint Analytics did not download the export through the prefixed Core.'
    } finally { Remove-Module $ea -Force }

    # ---- Autopatch: status, retry and the export import path ----
    $ap = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1' -Names @('Get-AutopatchGraphStatusCode', 'Get-AutopatchGraphRetryDelaySeconds', 'Invoke-AutopatchGraphRequest', 'Import-ExportedCsv') -Prelude ($standIns + @'

$script:GraphTransientRetryCount = 0
function Invoke-MgGraphRequest { [CmdletBinding()] param($Method, $Uri, $Body, $ContentType, $OutputType) Next-Answer }
function Ensure-Folder { param($Path) New-Item -ItemType Directory -Path $Path -Force | Out-Null }
function Start-ExportJob { param($ReportName, $Select, $Filter) [pscustomobject]@{ id = [guid]::NewGuid().ToString('N') } }
function Wait-ExportJob { param($JobId) [pscustomobject]@{ status = 'completed'; url = $script:ZipUrl } }
'@)
    try {
        $r = & $ap { param($a) Set-Answers $a; $v = Invoke-AutopatchGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/x'; [pscustomobject]@{ Value = $v; Calls = $script:Calls } } @($sdk500, 'ok')
        Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Autopatch did not classify a Graph SDK 500 through the prefixed Core.'
        $r = & $ap { param($a) Set-Answers $a; $v = Invoke-AutopatchGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/x'; [pscustomobject]@{ Value = $v; Calls = $script:Calls } } @($timeout, 'ok')
        Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Autopatch did not retry a timeout through the prefixed Core.'
        $rows = & $ap { param($u) $script:ZipUrl = $u; @(Import-ExportedCsv -ReportName 'SyntheticAutopatch') } $zipUrl
        Check (@($rows).Count -eq 1 -and $rows[0].DeviceId -eq 'd1') 'Autopatch did not import the export through the prefixed Core.'
    } finally { Remove-Module $ap -Force }
}
finally {
    if (Get-Variable -Name listener -ErrorAction Ignore) { $listener.Stop() }
    if (Get-Variable -Name server -ErrorAction Ignore) { $server.Dispose() }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks prefixed Core collector checks; PowerShell $($PSVersionTable.PSVersion)."
