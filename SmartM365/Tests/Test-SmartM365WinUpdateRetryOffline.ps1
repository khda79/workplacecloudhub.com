#Requires -Version 7.0
<#
.SYNOPSIS
Exercise Windows Update polling and ZIP retries with the real prefixed Core module offline.
.VERSION
1.0.0
.NOTES
Only extracted collector functions run. HTTP, waits and random samples are stand-ins;
no tenant, SDK connection, collector, synchronized data or real network access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Stand-ins accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='HTTP, waits and randomness are controlled inside the test modules.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.103' -Prefix Core -Force -PassThru -Global -WarningAction SilentlyContinue
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
function Get-SyntheticFailure([int]$Status, [string]$RetryAfter) {
    $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]$Status)
    if ($RetryAfter) { $null = $response.Headers.TryAddWithoutValidation('Retry-After', $RetryAfter) }
    [Microsoft.PowerShell.Commands.HttpResponseException]::new("HTTP $Status synthetic failure", $response)
}
function Get-SyntheticRecord($Exception) { [Management.Automation.ErrorRecord]::new($Exception, 'Synthetic', 'NotSpecified', $null) }
$sourcePath = Join-Path $root 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$null, [ref]$null)
$names = @('Get-GraphHttpStatusCode','Get-GraphRetryDelaySeconds','Test-GraphTransientError','Invoke-GraphRestMethodWithRetry','Get-IntuneExportJob','Download-ExportZip')
$definitions = foreach ($name in $names) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $node) { throw "Missing collector function: $name" }
    $node.Extent.Text
}
$collector = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-WinUpdateRetry-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
try {
    Check ($null -eq (Get-Command Get-SmartM365RetryDelay -ErrorAction SilentlyContinue)) 'An unprefixed Core import hides collector prefix errors.'
    Check ((Get-Content $sourcePath -Raw) -match "-MinimumVersion '1.0.103' -Prefix Core") 'Collector does not require the shared retry helper version.'
    # Guard the actual touched request path; these functions must remain REST-only.
    foreach ($text in @($definitions) + @(& $core { ${function:Invoke-SmartM365ExportFileDownload}.ToString(); ${function:Get-SmartM365RetryDelay}.ToString() })) {
        Check ($text -notmatch 'Set-MgRequestContext|Invoke-MgGraphRequest|EvidenceCollector') 'The tested request path now depends on the Graph SDK request context.'
    }
    & $collector {
        $script:Answers = [Collections.Generic.Queue[object]]::new()
        $script:Sleeps = [Collections.Generic.List[int]]::new()
        function script:Write-Log { param($Message, $Level, $Stage) }
        function script:Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds) }
        function script:Invoke-RestMethod {
            [CmdletBinding()] param($Method, $Uri, $Headers, $Body, $ContentType)
            $answer = $script:Answers.Dequeue()
            if ($answer -is [Exception]) { throw $answer }
            $answer
        }
    }
    $errorRecord = Get-SyntheticRecord (Get-SyntheticFailure 503 '2')
    foreach ($attempt in 1..8) {
        $delay = & $collector { param($e,$a) Get-GraphRetryDelaySeconds -ErrorRecord $e -Attempt $a } $errorRecord $attempt
        $floor = [math]::Min(300, 15 * [math]::Pow(2, $attempt - 1))
        Check ($delay -ge $floor -and $delay -le 300) "Polling floor or cap violated at attempt $attempt : $delay"
    }
    $delay = & $collector { param($e) Get-GraphRetryDelaySeconds -ErrorRecord $e -Attempt 1 } (Get-SyntheticRecord (Get-SyntheticFailure 503 '90'))
    Check ($delay -eq 90) 'Polling shortened Retry-After above its local floor.'
    $samples = @(1..40 | ForEach-Object { & $collector { param($e) Get-GraphRetryDelaySeconds -ErrorRecord $e -Attempt 1 } (Get-SyntheticRecord (Get-SyntheticFailure 503)) })
    Check (($samples | Measure-Object -Minimum).Minimum -ge 15 -and ($samples | Measure-Object -Maximum).Maximum -le 30 -and @($samples | Sort-Object -Unique).Count -gt 1) 'Polling lost bounded jitter.'
    $result = & $collector {
        param($failure)
        $script:Answers.Enqueue($failure); $script:Answers.Enqueue([pscustomobject]@{status='completed'})
        $value = Get-IntuneExportJob -JobId 'synthetic' -Headers @{} -MaxAttempts 5
        [pscustomobject]@{ Status=$value.status; Sleeps=$script:Sleeps.ToArray() }
    } (Get-SyntheticFailure 504 '40')
    Check ($result.Status -eq 'completed' -and $result.Sleeps.Count -eq 1 -and $result.Sleeps[0] -eq 40) 'Real status wrapper did not retry 504 with Retry-After through prefixed Core.'
    foreach ($status in @(401,403)) {
        $retry = & $collector { param($e) Test-GraphTransientError -ErrorRecord $e } (Get-SyntheticRecord (Get-SyntheticFailure $status))
        Check (-not $retry) "Permanent HTTP $status is retried."
    }
    # ZIP wrapper calls the real Core downloader; mock only its HTTP boundary.
    & $core {
        $script:Answers = [Collections.Generic.Queue[object]]::new()
        $script:Sleeps = [Collections.Generic.List[int]]::new()
        $script:Logs = [Collections.Generic.List[string]]::new()
        $script:Calls = 0; $script:Authorized = $false
        $script:Sample = 0.0
        function script:Get-Random { param($Minimum,$Maximum) $script:Sample }
        function script:Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds) }
        function script:WriteLog { param($Message,$Level) $script:Logs.Add($Message) }
        function script:Invoke-WebRequest {
            [CmdletBinding()] param($Method,$Uri,$OutFile,$TimeoutSec,$Headers)
            $script:Calls++; if ($Headers) { $script:Authorized = $true }
            [IO.File]::WriteAllText($OutFile, 'partial')
            $answer = $script:Answers.Dequeue()
            if ($answer -is [Exception]) { throw $answer }
            [IO.File]::WriteAllText($OutFile, 'zip')
        }
    }
    $target = Join-Path $temporary 'report.zip'
    & $core {
        param($withHeader,$withoutHeader,$capped)
        $script:Answers.Enqueue($withHeader); $script:Answers.Enqueue($withoutHeader); $script:Answers.Enqueue($capped); $script:Answers.Enqueue('ok')
    } (Get-SyntheticFailure 503 '37') (Get-SyntheticFailure 503) (Get-SyntheticFailure 503 '900')
    & $collector { param($out) Download-ExportZip -DownloadUrl 'https://storage.example.invalid/report.zip?sig=synthetic-secret' -DestinationZipPath $out } $target
    $state = & $core { [pscustomobject]@{Sleeps=$script:Sleeps.ToArray(); Calls=$script:Calls; Authorized=$script:Authorized; Logs=$script:Logs -join "`n"} }
    Check (($state.Sleeps -join ',') -eq '37,10,120') "ZIP ignored Retry-After, jitter lower bound or 120s cap: $($state.Sleeps -join ',')"
    Check ($state.Calls -eq 4 -and -not $state.Authorized -and [IO.File]::ReadAllText($target) -eq 'zip') 'ZIP wrapper failed to publish the recovered download without bearer.'
    Check ($state.Logs -match 'Retry-After' -and $state.Logs -match 'Backoff' -and $state.Logs -notmatch 'synthetic-secret') 'ZIP retry source missing or SAS query leaked.'
    & $core { $script:Sleeps.Clear(); $script:Answers.Enqueue([Exception]::new('HTTP 503')); $script:Answers.Enqueue('ok') }
    & $collector { param($out) Download-ExportZip -DownloadUrl 'https://storage.example.invalid/report.zip' -DestinationZipPath $out } $target
    Check ((& $core { $script:Sleeps[0] }) -eq 5) 'ZIP lower jitter bound did not reach 5s.'
    & $core { $script:Sleeps.Clear(); $script:Sample=1.0; $script:Answers.Enqueue([Exception]::new('HTTP 503')); $script:Answers.Enqueue('ok') }
    & $collector { param($out) Download-ExportZip -DownloadUrl 'https://storage.example.invalid/report.zip' -DestinationZipPath $out } $target
    Check ((& $core { $script:Sleeps[0] }) -eq 10) 'ZIP upper jitter bound did not reach 10s (lower bound is 5s).'
    & $core { $script:Sleeps.Clear(); $script:Calls=0; foreach ($i in 1..6) { $script:Answers.Enqueue([Exception]::new('HTTP 503')) } }
    $terminal = $null
    try { & $collector { param($out) Download-ExportZip -DownloadUrl 'https://storage.example.invalid/report.zip' -DestinationZipPath $out } $target } catch { $terminal = $_ }
    $state = & $core { [pscustomobject]@{Calls=$script:Calls; Sleeps=$script:Sleeps.ToArray()} }
    Check ($terminal -and $terminal.Exception.Message -match 'failed after 6 attempt' -and $state.Calls -eq 6 -and $state.Sleeps.Count -eq 5) 'ZIP retry attempts are not bounded or sleep after terminal failure.'
    Check (($state.Sleeps | Measure-Object -Maximum).Maximum -le 120 -and [IO.File]::ReadAllText($target) -eq 'zip') 'ZIP retry exceeded its cap or overwrote the previous successful file on failure.'
    Check (@(Get-ChildItem -LiteralPath $temporary -Filter '*.partial').Count -eq 0) 'Failed ZIP left partial files.'
    Write-Output "PASS: $script:checks Windows Update retry checks (offline)."
}
finally {
    & $core { Remove-Item Function:\Invoke-WebRequest,Function:\Start-Sleep,Function:\WriteLog,Function:\Get-Random -ErrorAction SilentlyContinue }
    Remove-Module $collector -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue
}
