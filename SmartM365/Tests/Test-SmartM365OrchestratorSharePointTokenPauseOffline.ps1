<#
.SYNOPSIS
Offline checks of the orchestrator SharePoint token pause (orchestrator 1.5.94): when the Graph token
cannot be obtained, the next uploads are skipped for SharePointTokenPauseMinutes instead of waiting
for the same token retries again, then resume.

.VERSION
1.0.0
#>
#requires -Version 7.2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$checks = 0
function Check([bool]$Condition, [string]$Message) { if (-not $Condition) { throw "FAIL: $Message" }; $script:checks++ }

$orchestratorPath = Join-Path $PSScriptRoot '..\SmartInventory\Orchestrator\SmartM365-Inventory-Orchestrator.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Resolve-Path $orchestratorPath), [ref]$tokens, [ref]$errors)
Check (@($errors).Count -eq 0) 'Orchestrator does not parse.'
$text = Get-Content -LiteralPath $orchestratorPath -Raw
Check ($text -match '\.VERSION\r?\n1\.5\.94\r?\n') 'Orchestrator .VERSION is not 1.5.94.'
Check ($text -match "\`$ScriptVersion = `"1\.5\.94`"") 'Orchestrator ScriptVersion is not 1.5.94.'
Check ($text -match '\$script:SharePointTokenPauseMinutes = 15') 'Token pause default is not 15 minutes.'

$node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-OrchestratorSharePointUpload' }, $true)
Check ($null -ne $node) 'Invoke-OrchestratorSharePointUpload not found.'

$stubs = @'
$script:Log = [Collections.Generic.List[string]]::new()
$script:ConnectCalls = 0
$script:UploadCalls = 0
$script:TokenAvailable = $false
$script:SharePointUploadedFileState = @{}
$script:SharePointTokenPausedUntilUtc = [datetime]::MinValue
$script:SharePointTokenPauseMinutes = 15
$script:SharePointTokenPausedSkips = 0
$script:Settings = [pscustomobject]@{
    SharedDataFolderPath = 'C:\none'; OrchestratorLogFolderPath = 'C:\none'
    SharePointSiteHostname = 'tenant.sharepoint.test'; SharePointSitePath = '/sites/x'
    SharePointLibraryDisplayName = 'Documents'; SharePointTargetFolderPath = 'X'
}
function Test-OrchestratorUploadCategoryEnabled { param($Category) $true }
function Test-OrchestratorSharePointUploadConfigured { $true }
function Get-SmartM365SharePointUploadCategory { param($Path, $DataAllRoot, $LogAllRoot) 'DataAll' }
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Log.Add("[$Level] $Message") }
function Connect-SmartM365GraphForSharePointUpload { param($AppId, $TenantId, $Thumbprint) $script:ConnectCalls++; return $script:TokenAvailable }
function Invoke-SmartM365SharePointCsvUpload {
    [CmdletBinding()] param($LocalFilePath, $Enabled, $SiteHostname, $SitePath, $LibraryDisplayName, $TargetFolderPath, [switch]$EnsureParentFolders)
    $script:UploadCalls++
    [pscustomobject]@{ WebUrl = 'https://tenant.sharepoint.test/' + [IO.Path]::GetFileName($LocalFilePath); SharePointPath = '' }
}
'@
$module = New-Module -ScriptBlock ([scriptblock]::Create($stubs + [Environment]::NewLine + $node.Extent.Text))

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-TokenPause-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    $files = foreach ($name in 'a.log', 'b.csv', 'c.json', 'd.json') {
        $path = Join-Path $root $name; Set-Content -LiteralPath $path -Value $name; $path
    }

    # 1. Token unavailable: one token attempt for the whole batch, no upload, then a pause.
    $results = & $module { param($Paths) foreach ($p in $Paths) { Invoke-OrchestratorSharePointUpload -LocalFilePath $p -Reason 'periodic orchestrator upload' } } $files
    $state = & $module { [pscustomobject]@{ Connect = $script:ConnectCalls; Upload = $script:UploadCalls; Skips = $script:SharePointTokenPausedSkips; Until = $script:SharePointTokenPausedUntilUtc; Log = @($script:Log) } }
    Check (@($results | Where-Object { $_ }).Count -eq 0) 'An upload reported success without a token.'
    Check ($state.Connect -eq 1) "Token requested $($state.Connect) times for one batch; expected 1."
    Check ($state.Upload -eq 0) 'Upload attempted without a token.'
    Check ($state.Skips -eq 3) "Expected 3 skipped uploads, got $($state.Skips)."
    Check ($state.Until -gt [datetime]::UtcNow.AddMinutes(14) -and $state.Until -le [datetime]::UtcNow.AddMinutes(15)) 'Pause is not about 15 minutes.'
    Check (@($state.Log | Where-Object { $_.StartsWith('[WARN] SharePoint uploads paused for 15 minutes') -and $_.Contains('a.log') }).Count -eq 1) 'Pause warning missing or not naming the first file.'

    # 2. Token back but pause not expired: still skipped, no token request.
    & $module { $script:TokenAvailable = $true }
    $null = & $module { param($p) Invoke-OrchestratorSharePointUpload -LocalFilePath $p -Reason 'job-runs CSV' -Force } $files[0]
    $state = & $module { [pscustomobject]@{ Connect = $script:ConnectCalls; Upload = $script:UploadCalls; Skips = $script:SharePointTokenPausedSkips } }
    Check ($state.Connect -eq 1 -and $state.Upload -eq 0 -and $state.Skips -eq 4) 'An upload ran during the pause.'

    # 3. Pause expired and token available: uploads resume and the skips are reported once.
    & $module { $script:SharePointTokenPausedUntilUtc = [datetime]::UtcNow.AddSeconds(-1) }
    $results = & $module { param($Paths) foreach ($p in $Paths) { Invoke-OrchestratorSharePointUpload -LocalFilePath $p -Reason 'periodic orchestrator upload' } } $files
    $state = & $module { [pscustomobject]@{ Connect = $script:ConnectCalls; Upload = $script:UploadCalls; Skips = $script:SharePointTokenPausedSkips; Log = @($script:Log) } }
    Check (@($results | Where-Object { $_ }).Count -eq 4) 'Uploads did not resume after the pause.'
    Check ($state.Upload -eq 4) "Expected 4 uploads after the pause, got $($state.Upload)."
    Check ($state.Skips -eq 0) 'Skip counter not reset after resume.'
    Check (@($state.Log | Where-Object { $_ -like '*SharePoint uploads resumed*4 upload(s) were skipped*' }).Count -eq 1) 'Resume message missing or wrong count.'

    # 4. Unchanged files are not re-uploaded and do not request a token.
    $before = $state.Connect
    $null = & $module { param($Paths) foreach ($p in $Paths) { Invoke-OrchestratorSharePointUpload -LocalFilePath $p -Reason 'periodic orchestrator upload' } } $files
    $state = & $module { [pscustomobject]@{ Connect = $script:ConnectCalls; Upload = $script:UploadCalls } }
    Check ($state.Connect -eq $before -and $state.Upload -eq 4) 'Unchanged files triggered a token request or an upload.'
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "Orchestrator SharePoint token pause offline: $checks checks passed."
