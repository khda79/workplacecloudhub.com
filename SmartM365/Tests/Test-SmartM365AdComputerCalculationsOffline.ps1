<# Synthetic replay of the actual AD computer enrichment loop. No collector or tenant connection. #>
[CmdletBinding()]
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'
$checks = 0
function Check([bool]$Ok, [string]$Message) {
    if (-not $Ok) { throw $Message }
    $script:checks++
}
$path = Join-Path $PSScriptRoot '../SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Enrichment.ps1'
if ($SourcePath) { $path = $SourcePath }
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
Check ($errors.Count -eq 0) 'Enrichment parser failed'
. (Join-Path $PSScriptRoot '../SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1')
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Common/SmartM365.AccountClassification.psd1') -MinimumVersion '1.1.0'
$entry = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-SmartM365AdComputersEnrichedCsv' }, $true)
foreach ($definition in $entry.Body.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
$loop = $entry.Body.Find({ param($n) $n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Extent.Text.StartsWith('foreach ($row in $sourceRows)') }, $true)
Check ($null -ne $loop) 'Actual enrichment loop missing'
function WriteLog { param($Message, $Level) }
function Get-SmartM365WorkplaceSite { param($SiteClassification, $SiteCode) return $null }
$siteClassification = @{}
$integrationFallback = @{ComputerCountryOuSuffix=''; Computers=@(); Default='INTEGRATED'}
$headquartersOuFragments = @()
$windows11OutOfScopeDomains = @()
$calculatedColumnsToAdd = @()
$rawCalculatedColumnRenames = [ordered]@{}
$progressInterval = 2000
$enrichmentTimer = [Diagnostics.Stopwatch]::StartNew()
$loopTimer = [Diagnostics.Stopwatch]::StartNew()
$maps = 'intuneByAad','entraByDeviceId','localByAad','wuAnyByDevice','wuAnchorByDevice','wu24ByDevice','wu25ByDevice','usersByUpn','licensedUsersByUpn','nativeEntraById','nativeManagedByAad','issuesByGuid','issueScoreByGuid','entraByNameAll','entraByDeviceIdAll','serverAdEntraByName','usersBySam','usersByDomainAndSam','mailboxByUpn','mailboxSizeByUpn','licenseGroupByUpn','hardwareCounts'
foreach ($name in $maps) { Set-Variable -Name $name -Value @{} }
$licenseRows = @()
$guid = '11111111-1111-1111-1111-111111111111'
$cases = @(
    @{Name='Win10 supported'; Os='Windows 10 Enterprise'; Ad='10.0 (19045)'; Intune=''; Build='19045'; Min='True'; Ltsc='False'},
    @{Name='Win10 below threshold'; Os='Windows 10'; Ad='10.0 (18363)'; Intune=''; Build='18363'; Min='False'; Ltsc='False'},
    @{Name='Win10 threshold'; Os='Windows 10'; Ad='10.0 (19041)'; Intune=''; Build='19041'; Min='True'; Ltsc='False'},
    @{Name='LTSC by name'; Os='Windows 10 Enterprise LTSC'; Ad=''; Intune=''; Build=''; Min='False'; Ltsc='True'},
    @{Name='LTSC by build'; Os='Windows 10'; Ad='10.0 (17763)'; Intune=''; Build='17763'; Min='False'; Ltsc='True'},
    @{Name='Win11'; Os='Windows 11'; Ad='10.0 (26100)'; Intune=''; Build='26100'; Min='True'; Ltsc='False'},
    @{Name='AD version absent'; Os='Windows 10'; Ad=$null; Intune='10.0.19045.1'; Build='19045'; Min='True'; Ltsc='False'},
    @{Name='Both versions absent'; Os='Windows 10'; Ad=$null; Intune=$null; Build=''; Min='False'; Ltsc='False'},
    @{Name='Malformed versions'; Os='Windows 10'; Ad='invalid'; Intune='invalid'; Build=''; Min='False'; Ltsc='False'},
    @{Name='All OS fields absent'; Os=$null; Ad=$null; Intune=$null; Build=''; Min='False'; Ltsc='False'},
    @{Name='Server'; Os='Windows Server 2019'; Ad='10.0 (17763)'; Intune=''; Build='17763'; Min='False'; Ltsc='False'}
)
$sourceRows = @($cases | ForEach-Object { [pscustomobject]@{Name=$_.Name; ObjectGUID=$guid; Enabled='True'; OperatingSystem=$_.Os; OperatingSystemVersion=$_.Ad} })
$sourceRowCount = $sourceRows.Count
# Replay consecutive rows in one scope, including missing data after populated data.
$enrichedRows = [Collections.Generic.List[object]]::new()
$iteration = $loop.Body.Extent.Text
$iteration = $iteration.Substring(1, $iteration.Length - 2)
for ($i = 0; $i -lt $cases.Count; $i++) {
    $row = $sourceRows[$i]
    $intuneByAad = @{}
    if ($cases[$i].Intune) { $intuneByAad[$guid] = @{OSVersion=$cases[$i].Intune} }
    . ([scriptblock]::Create($iteration))
    $result = $enrichedRows[$i]
    foreach ($pair in @(@('OperatingSystemBuildNumber','Build'), @('MeetsWindows11MinimumOsBuild','Min'), @('IsWindows10Ltsc','Ltsc'))) {
        Check ($result.($pair[0]) -ceq $cases[$i][$pair[1]]) ($cases[$i].Name + ': ' + $pair[0])
    }
    Check ($result.Windows11MigrationPhase -in @('UNDETERMINED','OUT OF SCOPE')) ($cases[$i].Name + ': absent activity must remain unqualified')
    Check ($result.EntityLabel -eq 'NOTFOUND') ($cases[$i].Name + ': missing site')
    Check ($result.ExistsInEntraByObjectGuid -eq 'False') ($cases[$i].Name + ': missing Entra')
    # Exercise the migration decisions with qualified recent activity as well.
    $phase = GetMigrationPhase 'CAPABLE' $osMinToUpdateW11 $isWindows10Ltsc $true (GetOsMajor $cases[$i].Os) 'INTEGRATED' $false '' $true ($null -ne $intune) ''
    $expectedPhase = if (-not (GetOsMajor $cases[$i].Os)) { 'OUT OF SCOPE' } elseif ($cases[$i].Ltsc -eq 'True' -or $cases[$i].Min -eq 'False') { 'PHASE 2' } else { 'PHASE 1' }
    Check ($phase -eq $expectedPhase) ($cases[$i].Name + ': migration phase with recent activity')
    Write-Output ('PASS: ' + $cases[$i].Name)
}
Check ($enrichedRows.Count -eq $cases.Count) 'Row count changed'
Write-Output ("PASS: {0} checks; {1} synthetic rows; actual enrichment loop; no collector execution." -f $checks, $cases.Count)
