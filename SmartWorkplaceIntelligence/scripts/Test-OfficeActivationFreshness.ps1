# Version 1.0.0
[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-OfflineTests'))
# Synthetic local files only. No collectors, tenant calls, production data or uploads.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:checks = 0
function Check([bool]$Pass, [string]$Message) { if (-not $Pass) { throw $Message }; $script:checks++ }
function Import-ScriptFunction([string]$File, [string]$Name) {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $File), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors.Message -join '; ') }
    $node = $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true) | Where-Object Name -eq $Name
    if (@($node).Count -ne 1) { throw "Function lookup failed: $Name" }
    [scriptblock]::Create($node.Extent.Text)
}

$now = [datetime]::new(2026, 10, 22, 12, 0, 0, [DateTimeKind]::Utc)
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$raw = Join-Path $fixture 'raw'
$last = Join-Path $raw 'DATA-LAST'
New-Item -ItemType Directory -Path $last -Force | Out-Null
$activationPath = Join-Path $last 'M365_Apps_Activations.csv'
$trustPath = Join-Path $fixture 'DataTrustEvidence.csv'

function Write-ActivationFixture([datetime]$ReportDate, [datetime]$AcquiredUtc) {
    [pscustomobject][ordered]@{
        TenantKey = 'synthetic-test'; RunId = 'run-1'; 'Report Refresh Date' = $ReportDate.ToString('yyyy-MM-dd')
        'User Principal Name' = 'user1@example.test'; 'Product Type' = 'MICROSOFT 365 APPS FOR ENTERPRISE'
        'Last Activated Date' = $ReportDate.AddDays(-3).ToString('yyyy-MM-dd'); Windows = '1'; Mac = '0'
    } | Export-Csv -LiteralPath $activationPath -NoTypeInformation
    [IO.File]::SetLastWriteTimeUtc($activationPath, $AcquiredUtc)
}
function Get-ActivationTrust([int]$WarningAgeHours = 0) {
    $trustArgs = @{DataRoot = $raw; OutputPath = $trustPath; AsOfDate = $now.Date; AsOfUtc = $now}
    if ($WarningAgeHours) { $trustArgs.OfficeActivationWarningAgeHours = $WarningAgeHours }
    & (Join-Path $PSScriptRoot 'New-DataTrustEvidence.ps1') @trustArgs | Out-Null
    $rows = @(Import-Csv -LiteralPath $trustPath | Where-Object 'File Name' -eq 'M365_Apps_Activations.csv')
    Check ($rows.Count -eq 1) 'Data Trust must carry exactly one Office activation source row'
    $rows[0]
}

try {
    # Fresh: acquired two hours ago; Graph's usual two-day report lag must not warn.
    Write-ActivationFixture $now.AddDays(-2) $now.AddHours(-2)
    $fresh = Get-ActivationTrust
    Check ($fresh.'Evidence Domain' -eq 'Applications' -and $fresh.'Source Name' -eq 'Microsoft 365 Apps activations') 'Office activation source identity changed'
    Check ($fresh.'Freshness State' -eq 'Fresh' -and $fresh.'Decision Readiness' -eq 'Ready' -and $fresh.'Evidence Status' -eq 'Observed') 'Fresh acquisition was not reported Fresh'
    Check ($fresh.'Snapshot Date' -eq $now.AddDays(-2).ToString('yyyy-MM-dd')) 'Snapshot date must stay the business Report Refresh Date'
    Check ($fresh.'Snapshot Basis' -match '^Business snapshot: Report Refresh Date; last acquisition 2026-10-22T10:00:00Z, Aging after 72 hours$') "Acquisition basis not traced: $($fresh.'Snapshot Basis')"

    # Kept old: the producer left the previous export (written 100 hours ago) unchanged.
    Write-ActivationFixture $now.AddDays(-6) $now.AddHours(-100)
    $kept = Get-ActivationTrust
    Check ($kept.'Freshness State' -eq 'Aging') 'Kept export past the threshold was not reported Aging'
    Check ($kept.'Decision Readiness' -eq 'Ready' -and $kept.'Evidence Status' -eq 'Observed') 'Aging warning must never block the source'
    Check ($kept.'Recommended Action' -match 'source producer') 'Aging source must point at the producer'
    $tolerant = Get-ActivationTrust -WarningAgeHours 120
    Check ($tolerant.'Freshness State' -eq 'Fresh' -and $tolerant.'Snapshot Basis' -match 'Aging after 120 hours') 'Configured threshold was ignored'

    # A business date already Stale stays Stale and still requires a refresh.
    Write-ActivationFixture $now.AddDays(-40) $now.AddHours(-100)
    $stale = Get-ActivationTrust
    Check ($stale.'Freshness State' -eq 'Stale' -and $stale.'Decision Readiness' -eq 'Refresh required') 'Stale business date was softened by the acquisition rule'

    # Absent: Data Trust reports it missing; Workforce keeps requiring it.
    Remove-Item -LiteralPath $activationPath
    $missing = Get-ActivationTrust
    Check ($missing.'File State' -eq 'Missing' -and $missing.'Decision Readiness' -eq 'Blocked' -and $missing.'Freshness State' -eq 'Unknown') 'Missing Office activation export was not reported Blocked'

    . (Import-ScriptFunction 'New-WorkforceIdentityEvidence.ps1' 'Get-OfficeActivationFreshnessWarning')
    $reference = [datetimeoffset]$now
    Check ($null -eq (Get-OfficeActivationFreshnessWarning -AcquiredUtc $now.AddHours(-2) -LatestReportDate $now.AddDays(-2) -ReferenceUtc $reference -WarningAgeHours 72)) 'Fresh acquisition raised a warning'
    Check ($null -eq (Get-OfficeActivationFreshnessWarning -AcquiredUtc $now.AddHours(-72) -LatestReportDate $now.AddDays(-5) -ReferenceUtc $reference -WarningAgeHours 72)) 'Warning raised at the threshold instead of past it'
    $warning = Get-OfficeActivationFreshnessWarning -AcquiredUtc $now.AddHours(-100) -LatestReportDate $now.AddDays(-6).Date -ReferenceUtc $reference -WarningAgeHours 72
    Check ($warning -match '^Office activation report is aging: M365_Apps_Activations\.csv was last acquired 2026-10-18T08:00:00Z \(100 hours ago; warning after 72 hours\); latest Report Refresh Date=2026-10-16\.') "Kept export warning text changed: $warning"
    $unknown = Get-OfficeActivationFreshnessWarning -AcquiredUtc $now.AddHours(-100) -LatestReportDate $null -ReferenceUtc $reference -WarningAgeHours 72
    Check ($unknown -match 'latest Report Refresh Date=unknown\.') 'Missing report date was not reported unknown'

    # Wiring: Workforce warns through the relayed warning stream and still requires the file.
    $workforce = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'New-WorkforceIdentityEvidence.ps1') -Raw
    Check ($workforce.Contains("Get-RequiredCsv 'DATA-LAST\M365_Apps_Activations.csv'")) 'Workforce no longer requires the Office activation export'
    Check ($workforce.Contains('if ($officeActivationWarning) { Write-Warning $officeActivationWarning }')) 'Workforce does not emit the freshness warning'
    $build = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Invoke-PreparedEvidenceBuild.ps1') -Raw
    Check ($build.Contains("`$wf+=@('-OfficeActivationWarningAgeHours',[string]`$OfficeActivationWarningAgeHours)") -and $build.Contains("`$trustArgs+=@('-OfficeActivationWarningAgeHours',[string]`$OfficeActivationWarningAgeHours)")) 'Build does not pass the threshold to Workforce and Data Trust'
    $pipeline = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Raw
    Check ($pipeline.Contains("`$systemPolicyArgs+=@('-OfficeActivationWarningAgeHours',`$OfficeActivationWarningAgeHours)") -and $pipeline.Contains('$provenance.OfficeActivationFreshness=')) 'Pipeline does not pass or record the threshold'
    $prepareRoot = Join-Path $PSScriptRoot '../../SmartM365/SmartInventory/PreparedEvidence'
    $prepare = Get-Content -LiteralPath (Join-Path $prepareRoot 'SmartM365-WorkplaceEvidence-Prepare.ps1') -Raw
    Check ($prepare.Contains('OfficeActivationWarningAgeHours=$officeActivationWarning')) 'Prepare does not pass the threshold setting'
    # Template keys are written into the private configuration, whose fingerprint retention compares across runs.
    $template = Get-Content -LiteralPath (Join-Path $prepareRoot 'SmartM365-WorkplaceEvidence-Prepare.local.json.template') -Raw
    Check (-not $template.Contains('PreparedOfficeActivationWarningAgeHours')) 'The optional threshold must stay out of the template'
    $resolution = [regex]::Match($prepare, "(?ms)^\s*\`$officeActivationWarning=ConfigValue 'PreparedOfficeActivationWarningAgeHours'\r?\n.*?must be between 1 and 8760\.'\}\r?\n")
    Check $resolution.Success 'Prepare threshold resolution not found'
    $resolve = [scriptblock]::Create('param($Configured) function ConfigValue([string]$Name) { $Configured }' + "`n" + $resolution.Value + '$officeActivationWarning')
    Check ((& $resolve $null) -eq 72 -and (& $resolve '') -eq 72) 'Absent setting must default to 72 hours'
    Check ((& $resolve 120) -eq 120) 'Configured setting was ignored'
    $rejected = $false; try { & $resolve 0 | Out-Null } catch { $rejected = $true }
    Check $rejected 'Out-of-range setting was accepted'
    foreach ($worker in 'New-WorkforceIdentityEvidence.ps1', 'New-DataTrustEvidence.ps1', 'Invoke-PreparedEvidenceBuild.ps1') {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $worker), [ref]$tokens, [ref]$errors)
        $parameter = @($ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'OfficeActivationWarningAgeHours' })
        Check ($parameter.Count -eq 1 -and $parameter[0].DefaultValue.Extent.Text -eq '72') "Default threshold must be 72 hours: $worker"
    }
}
finally {
    Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host "PASS: $script:checks Office activation freshness checks. Synthetic local evidence only; no collectors or live writes."
