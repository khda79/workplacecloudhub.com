#Requires -Version 5.1
<#
.SYNOPSIS
    Synthetic ImportExcel startup/repair tests. Never touches installed modules or tenant data.
.VERSION
    1.0.2
#>
[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ImportExcelTests-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$smartRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-ImportExcelHealth.ps1')
$script:Assertions = 0
function Assert-EP([bool]$Condition,[string]$Name) {
    if (-not $Condition) { throw "Assertion failed: $Name" }
    $script:Assertions++
}
function Assert-EPThrows([scriptblock]$Action,[string]$Pattern,[string]$Name) {
    $caught = $null
    try { & $Action | Out-Null } catch { $caught = $_ }
    Assert-EP ($null -ne $caught -and $caught.Exception.Message -match $Pattern) $Name
}
function New-EPFixture([string]$Name,[string]$Version='1.2.3',[string]$Guid='bac3b4bb-710b-4e65-9c32-e3f472c87549') {
    $base=Join-Path (Join-Path (Join-Path $TestRoot $Name) 'ImportExcel') $Version
    $null=New-Item -ItemType Directory -Path $base -Force
    "@{ RootModule='ImportExcel.psm1'; ModuleVersion='$Version'; GUID='$Guid'; FunctionsToExport=@('*'); CmdletsToExport=@(); AliasesToExport=@(); VariablesToExport=@() }" | Set-Content (Join-Path $base 'ImportExcel.psd1') -Encoding UTF8
    '. (Join-Path $PSScriptRoot ''Plot.ps1'')' | Set-Content (Join-Path $base 'ImportExcel.psm1') -Encoding UTF8
    "function Import-Excel {}`nfunction Export-Excel {}`nfunction Open-ExcelPackage {}`nfunction Close-ExcelPackage {}`nfunction Get-ExcelSheetInfo {}" | Set-Content (Join-Path $base 'Plot.ps1') -Encoding UTF8
    'Synthetic fixture. No EPPlus and no business data.' | Set-Content (Join-Path $base 'README.txt') -Encoding UTF8
    [pscustomobject]@{ Path=(Join-Path $base 'ImportExcel.psd1');ModuleBase=$base;Version=[version]$Version;Guid=[guid]$Guid }
}
function Reset-EPFixtureModule {
    Get-Module ImportExcel | Where-Object { $_.ModuleBase.StartsWith($TestRoot,[StringComparison]::OrdinalIgnoreCase) } | Remove-Module -Force
}
$script:RealCandidates = ${function:Get-SmartM365ImportExcelCandidates}
$script:RealProbe = ${function:Test-SmartM365ImportExcelLoad}
$script:RealRepair = ${function:Repair-SmartM365ImportExcelMissingFiles}
$script:Candidates = @()
$script:Saves = 0; $script:Installs = 0; $script:RepositoryTrusted = $true
$script:SavedVersion = ''; $script:RepositoryReference = $null
function Get-SmartM365ImportExcelCandidates {
    param([string[]]$ModuleRoots, [switch]$IncludeReferenceRoots)
    if ($ModuleRoots) { & $script:RealCandidates -ModuleRoots $ModuleRoots }
    else { $script:Candidates }
}
function Get-PSRepository { param($Name) [pscustomobject]@{InstallationPolicy=$(if($script:RepositoryTrusted){'Trusted'}else{'Untrusted'})} }
function Get-PackageProvider { param($Name,[switch]$ListAvailable) [pscustomobject]@{Name='NuGet'} }
function Save-Module {
    [CmdletBinding()]param($Name,$RequiredVersion,$Repository,$Path)
    Assert-EP ($Name -eq 'ImportExcel' -and $Repository -eq 'PSGallery') 'Repository name and package constrained'
    $script:Saves++; $script:SavedVersion=$RequiredVersion
    $destination=Join-Path (Join-Path $Path 'ImportExcel') $RequiredVersion
    $null=New-Item -ItemType Directory -Path $destination -Force
    Get-ChildItem -LiteralPath $script:RepositoryReference.ModuleBase | Copy-Item -Destination $destination -Recurse
}
function Install-Module {
    [CmdletBinding()]param($Name,$Scope,$Repository,[switch]$Force,[switch]$AllowClobber)
    Assert-EP ($Name -eq 'ImportExcel' -and $Scope -eq 'CurrentUser') 'Initial installation scope retained'
    $script:Installs++; $script:Candidates=@($script:RepositoryReference)
}
function Write-SmartM365ImportExcelHealthLog { param($Message,$Level) }
$null=New-Item -ItemType Directory -Path $TestRoot -Force
$policyBefore=Get-ExecutionPolicy
$modulePathBefore=$env:PSModulePath
try {
    $healthy=New-EPFixture 'healthy'
    Assert-EP ((Get-SmartM365ImportExcelExecutionPolicy) -eq [string]$policyBefore) 'Built-in security dependency preserves effective policy'
    Assert-EP ((Get-SmartM365ImportExcelExecutionPolicy) -eq [string]$policyBefore) 'Repeated security initialization is safe'
    & {
        # Synthetic host/type-data failure, before probing or package management.
        $beforeSaves=$script:Saves; $beforeInstalls=$script:Installs
        function Import-Module {
            [CmdletBinding()]param($Name,[switch]$Global)
            Assert-EP ($Name -eq (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Security/Microsoft.PowerShell.Security.psd1')) 'Security dependency pinned to current engine'
            throw 'Synthetic TypeData ObjectSecurity: member AuditToString is already present.'
        }
        function Get-SmartM365ImportExcelCandidates { throw 'Candidate discovery must not run after a host failure' }
        foreach ($options in @(@{InstallIfMissing=$true}, @{CheckOnly=$true}, @{Scope='AllUsers';InstallIfMissing=$true})) {
            $failure=$null
            try { Initialize-SmartM365ImportExcel @options -RepairRoot (Join-Path $TestRoot 'host-failure') } catch { $failure=$_ }
            Assert-EP ($null -ne $failure -and $failure.FullyQualifiedErrorId -like 'SmartM365.ImportExcel.HostEnvironment*') 'Host failure has a distinct terminating error'
            Assert-EP ($failure.Exception.Message -match 'AuditToString' -and $failure.Exception.Message -match 'No ImportExcel repair') 'Host failure retains actionable original cause'
        }
        Assert-EPThrows { & $script:RealProbe $healthy.Path } 'host environment check failed' 'Direct probe stops before child-process creation'
        Assert-EP ($script:Saves -eq $beforeSaves -and $script:Installs -eq $beforeInstalls) 'Host failure never reaches package management'
        Assert-EP (-not (Test-Path (Join-Path $TestRoot 'host-failure'))) 'Host failure creates no repair directory'
    }
    Assert-EP ((& $script:RealProbe $healthy.Path).Success) 'Real isolated same-edition import and command checks'
    $discovered=@(& $script:RealCandidates -ModuleRoots @((Join-Path $TestRoot 'healthy')))
    Assert-EP ($discovered.Count -eq 1 -and $discovered[0].Version -eq [version]'1.2.3') 'Data-only manifest discovery'
    $manifestMissing=New-EPFixture 'manifest-missing'
    Remove-Item (Join-Path $manifestMissing.ModuleBase 'ImportExcel.psd1')
    Assert-EPThrows { & $script:RealCandidates -ModuleRoots @((Join-Path $TestRoot 'manifest-missing')) } 'no manifest' 'Missing manifest never triggers destructive reinstallation'
    $script:Candidates=@($healthy)
    $ready=Initialize-SmartM365ImportExcel -SkipRepositoryRepair -RepairRoot (Join-Path $TestRoot 'backups')
    Assert-EP (-not $ready.Repaired -and $ready.Path -eq $healthy.Path -and $script:Saves -eq 0) 'Healthy import does not repair or download'
    Assert-EP (-not (Test-Path (Join-Path $TestRoot 'backups'))) 'Healthy import writes no repair directory'
    Reset-EPFixtureModule

    $broken=New-EPFixture 'broken'
    Remove-Item -LiteralPath (Join-Path $broken.ModuleBase 'Plot.ps1')
    Assert-EP (-not (& $script:RealProbe $broken.Path).Success) 'Missing Plot.ps1 reproduced in isolated process'
    $script:Candidates=@($broken,$healthy)
    Assert-EPThrows { Initialize-SmartM365ImportExcel -CheckOnly -RepairRoot (Join-Path $TestRoot 'check-only') } 'check-only load failed' 'Check-only mode never repairs a broken installed module'
    Assert-EP (-not (Test-Path (Join-Path $TestRoot 'check-only')) -and -not (Test-Path (Join-Path $broken.ModuleBase 'Plot.ps1'))) 'Read-only check leaves broken files and backup root untouched'
    $beforeHash=(Get-FileHash (Join-Path $broken.ModuleBase 'ImportExcel.psm1')).Hash
    $script:Candidates=@($broken,$healthy)
    $ready=Initialize-SmartM365ImportExcel -SkipRepositoryRepair -RepairRoot (Join-Path $TestRoot 'backups')
    Assert-EP ($ready.Repaired -and $ready.Version -eq [version]'1.2.3') 'Missing file repaired without version change'
    Assert-EP ((Get-FileHash (Join-Path $broken.ModuleBase 'ImportExcel.psm1')).Hash -eq $beforeHash) 'Existing root module unchanged'
    Assert-EP ((Get-FileHash (Join-Path $broken.ModuleBase 'Plot.ps1')).Hash -eq (Get-FileHash (Join-Path $healthy.ModuleBase 'Plot.ps1')).Hash) 'Restored hash matches healthy source'
    $receipt=Get-Content (Join-Path $ready.BackupPath 'repair.json.txt') -Raw | ConvertFrom-Json
    Assert-EP ($receipt.State -eq 'Verified' -and @($receipt.Missing).Count -eq 1) 'Verified backup receipt'
    Assert-EP (Test-Path (Join-Path $ready.BackupPath 'before/ImportExcel.psm1')) 'Existing files backed up'
    Assert-EP ((Get-Module ImportExcel).ModuleBase -eq $broken.ModuleBase) 'Parent pins repaired user installation'
    Reset-EPFixtureModule

    $rootMissing=New-EPFixture 'root-missing'
    Remove-Item (Join-Path $rootMissing.ModuleBase 'ImportExcel.psm1')
    Assert-EP (@(& $script:RealCandidates -ModuleRoots @((Join-Path $TestRoot 'root-missing'))).Count -eq 1) 'Missing root does not look like an absent installation'
    $script:Candidates=@($rootMissing,$healthy)
    Assert-EP (Initialize-SmartM365ImportExcel -SkipRepositoryRepair -RepairRoot (Join-Path $TestRoot 'backups')).Repaired 'Root-module missing-file repair'
    Reset-EPFixtureModule

    $different=New-EPFixture 'different'
    Remove-Item (Join-Path $different.ModuleBase 'Plot.ps1')
    'Existing customized root module' | Add-Content (Join-Path $different.ModuleBase 'ImportExcel.psm1')
    $customHash=(Get-FileHash (Join-Path $different.ModuleBase 'ImportExcel.psm1')).Hash
    $script:Candidates=@($different,$healthy)
    Assert-EPThrows { Initialize-SmartM365ImportExcel -SkipRepositoryRepair -RepairRoot (Join-Path $TestRoot 'backups') } 'existing file differs' 'Different existing files are not overwritten'
    Assert-EP ((Get-FileHash (Join-Path $different.ModuleBase 'ImportExcel.psm1')).Hash -eq $customHash -and -not (Test-Path (Join-Path $different.ModuleBase 'Plot.ps1'))) 'Refused repair leaves installation unchanged'

    $noReference=New-EPFixture 'no-reference'
    Remove-Item (Join-Path $noReference.ModuleBase 'Plot.ps1')
    $wrongVersion=New-EPFixture 'wrong-version' '1.2.2'
    $wrongGuid=New-EPFixture 'wrong-guid' '1.2.3' 'de151cd2-e486-4312-9c8b-85b14e52c7e5'
    $script:Candidates=@($noReference,$wrongVersion,$wrongGuid)
    Assert-EPThrows { Initialize-SmartM365ImportExcel -SkipRepositoryRepair -RepairRoot (Join-Path $TestRoot 'backups') } 'same-version local' 'Wrong versions/GUIDs are not repair sources'
    Assert-EP ($script:Saves -eq 0) 'Offline mode never downloads packages'
    $script:Candidates=@($noReference)
    $script:RepositoryReference=$healthy
    $ready=Initialize-SmartM365ImportExcel -RepairRoot (Join-Path $TestRoot 'backups')
    Assert-EP ($ready.Repaired -and $script:Saves -eq 1 -and $script:SavedVersion -eq '1.2.3') 'Repository fallback requests exact installed version'
    Reset-EPFixtureModule

    $untrusted=New-EPFixture 'untrusted'
    Remove-Item (Join-Path $untrusted.ModuleBase 'Plot.ps1')
    $script:Candidates=@($untrusted);$script:RepositoryTrusted=$false
    Assert-EPThrows { Initialize-SmartM365ImportExcel -RepairRoot (Join-Path $TestRoot 'backups') } 'not trusted' 'No automatic repository trust change'
    Assert-EP ($script:Saves -eq 1) 'Untrusted repository never downloaded'
    $script:RepositoryTrusted=$true

    $script:Candidates=@()
    Assert-EPThrows { Initialize-SmartM365ImportExcel -SkipRepositoryRepair } 'not installed' 'Optional/worker calls do not introduce initial installation'
    Assert-EP ($script:Installs -eq 0) 'Absent optional dependency not installed'
    $ready=Initialize-SmartM365ImportExcel -InstallIfMissing -SkipRepositoryRepair -RepairRoot (Join-Path $TestRoot 'backups')
    Assert-EP ($script:Installs -eq 1 -and $ready.Path -eq $healthy.Path) 'Existing initial-install policy retained'
    Reset-EPFixtureModule

    # Security failure classification: no repair, repository or permission changes.
    function Test-SmartM365ImportExcelLoad { param($ManifestPath) [pscustomobject]@{Success=$false;Error='Script is not digitally signed';ErrorId='PSSecurityException';Version=''} }
    $script:Candidates=@($healthy)
    Assert-EPThrows { Initialize-SmartM365ImportExcel -RepairRoot (Join-Path $TestRoot 'backups') } 'security/access failure' 'AllSigned failure is not bypassed or reinstalled'
    Assert-EP ($script:Saves -eq 1 -and $script:Installs -eq 1) 'Security errors do not touch package management'
    Set-Item Function:Test-SmartM365ImportExcelLoad -Value $script:RealProbe

    # Failed post-copy probe rolls back only files restored by this attempt.
    $rollback=New-EPFixture 'rollback'
    Remove-Item (Join-Path $rollback.ModuleBase 'Plot.ps1')
    $rollbackHash=(Get-FileHash (Join-Path $rollback.ModuleBase 'ImportExcel.psm1')).Hash
    function Test-SmartM365ImportExcelLoad { param($ManifestPath) [pscustomobject]@{Success=$false;Error='Synthetic post-copy load failure';ErrorId='Test';Version=''} }
    Assert-EPThrows { Repair-SmartM365ImportExcelMissingFiles -Target $rollback -Reference $healthy -RepairRoot (Join-Path $TestRoot 'backups') } 'repair failed' 'Post-copy failure is fatal'
    Assert-EP (-not (Test-Path (Join-Path $rollback.ModuleBase 'Plot.ps1')) -and (Get-FileHash (Join-Path $rollback.ModuleBase 'ImportExcel.psm1')).Hash -eq $rollbackHash) 'Rollback preserves every prior file'
    Set-Item Function:Test-SmartM365ImportExcelLoad -Value $script:RealProbe

    # Requested AllUsers root: never silently use the healthy CurrentUser installation.
    $script:Candidates=@($healthy)
    $allUsersRoot=Join-Path $TestRoot 'all-users/ImportExcel'
    $ready=Initialize-SmartM365ImportExcel -Scope AllUsers -RequiredModuleRoot $allUsersRoot -SkipRepositoryRepair -RepairRoot (Join-Path $TestRoot 'backups')
    Assert-EP ($ready.Path -eq (Join-Path $allUsersRoot '1.2.3/ImportExcel.psd1') -and $ready.Repaired) 'Verified AllUsers promotion retains version and target root'
    Reset-EPFixtureModule
    Assert-EPThrows { Initialize-SmartM365ImportExcel -Scope AllUsers -RequiredModuleRoot $allUsersRoot -ManifestPath $healthy.Path -SkipRepositoryRepair } 'outside the required installation scope' 'Explicit manifest cannot bypass AllUsers requirement'
    Assert-EPThrows { Initialize-SmartM365ImportExcel -ManifestPath (Join-Path $TestRoot 'unknown.psd1') -SkipRepositoryRepair } 'not discovered' 'Explicit path is never substituted'

    Assert-EP ((Get-ExecutionPolicy) -eq $policyBefore -and $env:PSModulePath -ceq $modulePathBefore) 'Execution policy and PSModulePath unchanged'
    $helper=[IO.File]::ReadAllText((Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-ImportExcelHealth.ps1'))
    Assert-EP ($helper -notmatch 'Set-ExecutionPolicy|Set-PSRepository|Start-Process.+RunAs|ExecutionPolicy Bypass|Remove-Item.+Recurse') 'No security weakening or recursive cleanup'
    Assert-EP ($helper -match 'WaitOne\(45000\)' -and $helper -match 'ReleaseMutex' -and $helper -match 'FileMode\]::CreateNew') 'Cross-process lock and exclusive file creation'
    Assert-EPThrows { Repair-SmartM365ImportExcelMissingFiles -Target $broken -Reference $healthy -RepairRoot $broken.ModuleBase } 'outside module' 'Backup destination cannot be inside repaired installation'

    # Two real processes racing on the same synthetic broken module: one repair.
    $concurrent=New-EPFixture 'concurrent'
    Remove-Item (Join-Path $concurrent.ModuleBase 'Plot.ps1')
    $parallelRoot=Join-Path $TestRoot 'parallel-backups'
    $targetInput=@{Path=$concurrent.Path;ModuleBase=$concurrent.ModuleBase;Version=[string]$concurrent.Version;Guid=[string]$concurrent.Guid}
    $referenceInput=@{Path=$healthy.Path;ModuleBase=$healthy.ModuleBase;Version=[string]$healthy.Version;Guid=[string]$healthy.Guid}
    $parallelInput=[ordered]@{Helper=(Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-ImportExcelHealth.ps1');Target=$targetInput;Reference=$referenceInput;RepairRoot=$parallelRoot}
    $inputBase64=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($parallelInput | ConvertTo-Json -Depth 5 -Compress)))
    $worker=@'
$ErrorActionPreference='Stop'
$fixture=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__INPUT__')) | ConvertFrom-Json
. $fixture.Helper
foreach($candidate in @($fixture.Target,$fixture.Reference)) { $candidate.Version=[version]$candidate.Version; $candidate.Guid=[guid]$candidate.Guid }
function Get-SmartM365ImportExcelCandidates { param($ModuleRoots,[switch]$IncludeReferenceRoots) @($fixture.Target,$fixture.Reference) }
function Write-SmartM365ImportExcelHealthLog { param($Message,$Level) }
function Save-Module { throw 'Network operation forbidden in parallel fixture.' }
function Install-Module { throw 'Installation forbidden in parallel fixture.' }
$result=Initialize-SmartM365ImportExcel -SkipRepositoryRepair -RepairRoot $fixture.RepairRoot
[Console]::WriteLine('PARALLEL_RESULT:' + ($result | Select-Object Repaired,BackupPath | ConvertTo-Json -Compress))
'@
    $worker=$worker.Replace('__INPUT__',$inputBase64)
    $engine=Join-Path $PSHOME $(if($PSVersionTable.PSVersion.Major -ge 6){'pwsh.exe'}else{'powershell.exe'})
    $processes=@()
    try {
        foreach($index in 1..2){
            $info=New-Object Diagnostics.ProcessStartInfo
            $info.FileName=$engine; $info.Arguments='-NoProfile -NonInteractive -ExecutionPolicy ' + [string](Get-ExecutionPolicy) + ' -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($worker))
            $info.UseShellExecute=$false;$info.CreateNoWindow=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
            $process=New-Object Diagnostics.Process;$process.StartInfo=$info
            $null=$process.Start()
            $processes+=[pscustomobject]@{Process=$process;Output=$process.StandardOutput.ReadToEndAsync();Errors=$process.StandardError.ReadToEndAsync()}
        }
        $results=@()
        foreach($entry in $processes){
            if(-not $entry.Process.WaitForExit(60000)){ $entry.Process.Kill();throw 'Synthetic concurrent worker timed out.' }
            $output=$entry.Output.GetAwaiter().GetResult();$errors=$entry.Errors.GetAwaiter().GetResult()
            Assert-EP ($entry.Process.ExitCode -eq 0) "Concurrent worker succeeded: $errors"
            $line=@($output -split '\r?\n' | Where-Object { $_.StartsWith('PARALLEL_RESULT:') })
            Assert-EP ($line.Count -eq 1) 'Unambiguous concurrent worker result'
            $results+=$line[0].Substring('PARALLEL_RESULT:'.Length) | ConvertFrom-Json
        }
        Assert-EP (@($results | Where-Object Repaired).Count -eq 1) 'Exactly one process repairs under the shared mutex'
        Assert-EP (@(Get-ChildItem $parallelRoot -Directory -Filter 'repair-*').Count -eq 1) 'Only one parallel repair backup'
        Assert-EP ((& $script:RealProbe $concurrent.Path).Success) 'Concurrent repaired module loads normally'
    } finally { foreach($entry in $processes){ if(-not $entry.Process.HasExited){$entry.Process.Kill()};$entry.Process.Dispose() } }

    # AST checks cover all direct production consumers and indirect entry points.
    $consumers=@(
        'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1',
        'SmartInventory/ExchangeInventory/Quarantine/SmartM365-EXO-QuarantineMessages-Report.ps1',
        'SmartInventory/ExchangeInventory/OnPremises/ProxyAddresses/SmartM365-Check-ProxyAddresses-Exchange.ps1',
        'SmartInventory/M365Inventory/IntuneInventory/RBAC/SmartM365-Intune-RBAC-GroupMembers.ps1',
        'SmartInventory/M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1',
        'SmartInventory/M365Inventory/Teams/SmartM365-Teams-Inventory.ps1',
        'SmartInventory/M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1',
        'SmartInventory/M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.ps1',
        'SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1',
        'SmartInventory/Common/SmartM365.WorkplaceClassification.psm1'
    )
    foreach($relative in $consumers){
        $path=Join-Path $smartRoot $relative
        # PS7 consumers can be UTF-8 without BOM; PS5 ParseFile assumes ANSI.
        # Decode explicitly for this static scan; never execute a consumer.
        $tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseInput([IO.File]::ReadAllText($path),[ref]$tokens,[ref]$errors)
        Assert-EP (@($errors).Count -eq 0) "Consumer parses: $relative"
        $text=$ast.Extent.Text
        Assert-EP ($text -match 'Initialize-SmartM365ImportExcel|RequiredModules.*ImportExcel') "Startup helper coverage: $relative"
        Assert-EP ($text -notmatch 'Import-Module\s+(?:-Name\s+)?ImportExcel\b') "No unchecked name-based import: $relative"
    }
    $prep=Get-Content (Join-Path $smartRoot $consumers[8]) -Raw
    Assert-EP ($prep -match 'if \(-not \(\$TransferOnly -or \$ConvertMetadata -or \$RepairLegacyHistory -or \$ValidateOnly\)\)') 'Non-building preparation modes skip module repair'
    Assert-EP ($prep.IndexOf('Initialize-SmartM365ImportExcel') -lt $prep.IndexOf('Initialize-PreparedMetadataNames')) 'Preparation checks before output metadata/capture'
    $rbac=Get-Content (Join-Path $smartRoot $consumers[3]) -Raw
    Assert-EP ($rbac -match 'if \(\$script:ImportExcelReady\)' -and $rbac.IndexOf('Initialize-SmartM365ImportExcel') -lt $rbac.IndexOf('Connect-SmartM365CloudSession @connectParams')) 'RBAC optional Excel verified before cloud connection'
    $teams=Get-Content (Join-Path $smartRoot $consumers[5]) -Raw
    Assert-EP ($teams.IndexOf("Check ImportExcel startup health") -lt $teams.IndexOf('if($MailOnly){')) 'Teams MailOnly and collection share startup check'
    $proxy=Get-Content (Join-Path $smartRoot $consumers[2]) -Raw
    Assert-EP ($proxy -match "WindowsPowerShell/Modules/ImportExcel" -and $proxy -match 'Scope AllUsers -RequiredModuleRoot') 'ProxyAddresses AllUsers requirement unchanged'
    $startupCheck=$proxy.IndexOf('$script:ImportExcelManifestPath = Ensure-ProxyAddressesImportExcelAllUsers')
    Assert-EP ($startupCheck -gt 0 -and $startupCheck -lt $proxy.IndexOf('Add-PSSnapin $snapinName')) 'ProxyAddresses verifies security and Excel before loading Exchange'
    $proxyMinimum=[regex]::Match($proxy,"SmartM365-WindowsPowerShell5\.psd1'\) -MinimumVersion '(?<V>[0-9.]+)'")
    Assert-EP ($proxyMinimum.Success -and [version]$proxyMinimum.Groups['V'].Value -ge [version]'1.0.58') 'ProxyAddresses requires the fixed compatibility module (1.0.58 or later)'
    foreach($moduleRelative in @('Modules/SmartM365.Core/SmartM365.Core.psm1','Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')){
        $text=Get-Content (Join-Path $smartRoot $moduleRelative) -Raw
        Assert-EP ($text -match 'SmartM365-ImportExcelHealth.ps1' -and $text -match 'if \(\$moduleName -eq ''ImportExcel''\)') "Shared preflight routing: $moduleRelative"
        & {
            $tokens=$null;$errors=$null
            $ast=[Management.Automation.Language.Parser]::ParseInput($text,[ref]$tokens,[ref]$errors)
            $definition=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-SmartM365Preflight'},$true)
            . ([scriptblock]::Create($definition.Extent.Text))
            $script:PreflightHealthCalls=@();$script:PreflightHealthFail=$false
            function WriteLog { param($Message,$Level) }
            function Initialize-SmartM365ImportExcel {
                param($Scope,$Repository,[switch]$InstallIfMissing,[switch]$SkipRepositoryRepair)
                $script:PreflightHealthCalls+=[pscustomobject]@{Scope=$Scope;Repository=$Repository;Install=[bool]$InstallIfMissing;Offline=[bool]$SkipRepositoryRepair}
                if($script:PreflightHealthFail){throw 'Synthetic startup dependency failure'}
            }
            Assert-EP (Invoke-SmartM365Preflight -RequiredModules ImportExcel -InstallModuleScope AllUsers -InstallModuleRepository 'AuthorizedFixtureRepository') 'Real preflight routes installed/import checks through helper'
            Assert-EP ($script:PreflightHealthCalls.Count -eq 1 -and $script:PreflightHealthCalls[0].Scope -eq 'AllUsers' -and $script:PreflightHealthCalls[0].Repository -eq 'AuthorizedFixtureRepository' -and $script:PreflightHealthCalls[0].Install) 'Preflight forwards approved scope/repository'
            $null=Invoke-SmartM365Preflight -RequiredModules ImportExcel -SkipMissingModuleInstall
            Assert-EP (-not $script:PreflightHealthCalls[1].Install -and $script:PreflightHealthCalls[1].Offline) 'Skip-missing-install mode blocks repository mutation'
            $script:PreflightHealthFail=$true
            Assert-EPThrows { Invoke-SmartM365Preflight -RequiredModules ImportExcel } 'preflight failed' 'Required dependency failure stops preflight'
        }
    }
    "PASS: $script:Assertions offline assertions; PowerShell $($PSVersionTable.PSVersion); fixtures=$TestRoot. No real modules repaired, collections, preparation or transfers."
} finally {
    Reset-EPFixtureModule
    Set-Item Function:Get-SmartM365ImportExcelCandidates -Value $script:RealCandidates
    Set-Item Function:Test-SmartM365ImportExcelLoad -Value $script:RealProbe
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC4HIorZS/mJa3P
# cVlwSAeYROW/4zmzItNgDBQViTSt8KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAxd00b967ybPEJ9jTOXMaPvVXhlO0DJ64xyw6CWg2uRMA0GCSqG
# SIb3DQEBAQUABIIBgJLoru//shOCbc0kpJSDVMEgBY0m8za5Ylht8Tgvook/gtTz
# tInCR1pr0ABa/v2oHWoqtMtrckEFZwGU1LRGKHMM53UIWcUYe1JS1FnoLF2Tf8N6
# cNsY6GLN01BMojOoIvmCidrPjtMsyV0vOihm+A8LItdSkdpA5B7PqJ5Z8JUtQXiE
# tc5aTPf0H9LBpFYAxOMbGzs2tR29u5slKsie7YClRfrWcNL5nSOgunPMiVuZWDQE
# XHBZjswJAtx6QAhMcrN3aw90qCVIqyjrYndGyhJ03lKnr61gG3WBIQDiArA98jB8
# 1d4/sYVgNX2sEUVaHnQERUjvsGXyquydzluxA3L8EQOfxc1mJVKelAAOA7s7yM/4
# LOiXEGMK6Z/Mf5DlAqHtkT9XrvMxZyeQzXOiGA5Tnz0ULyet4DZ5J54ZdMRrNV9x
# lGQDY6sOfRKZDimh+en2rQ6zoZ5ioUXN00tK4wA2wACyG2aImVCIPKKAXDCgSpMA
# YgSwJfUy76V0rA5ts6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxMDA4
# MTlaMC8GCSqGSIb3DQEJBDEiBCC7I2QVU26XjSr+I/yqqqAo1nrhR049NJfbxkrE
# LHapbjANBgkqhkiG9w0BAQEFAASCAgCjAFqFiUOnYme3ClpK/hYr2IigTr0huYrl
# Euyw3r61/WpkrOxGNyrIR7/lzpss3lHJv/5MLugv68uSGo3+assIiKVghYvZmTsq
# 378AeBYbt7CMHdL3bEWUyfdEM+1hRD7+xW0tQHuAseMyOOPQynB1vV/d17vYpr7j
# iU35iUpxu8sGZoV5DCc8kfYfYB9z8Q+EU5QpnReVr9Y5/LCUqtKGlIjvw8o9fzGD
# 2WYN/AhdHtHdHB9aVjmBFP7LnBNH+u5IWPfps8I8XlWFgYeX3ZDC6sGaukX40s4Z
# sMf8zgZ963Gxzjxxmnp1CrFc3U+iC6mtHOEfC8/5ToaDZsjuQcZ8BzryK59xZtXU
# t3Lzdg01czn8LO2uianXB6BHm7hBG4mJd1bn/+74UuHjl01YWe/+yj52/+SHKeZV
# 5yartDd+YmPkaK6Gzg/Zs5lcV8b3x+RvSfIPWJdZVGbMAWI5hY+fuYqeE3lb3Za6
# HJ7IKamvOWmabFtGvKxX5RhlOmTyVg0zDWPGg42lEPvP7Q+vvE4AglSJGfhsOyt2
# 3YlQnM/HiFLYKgerASHys13fBEcr2YIwWkHjAgio1NfdJuWBgL+EvihOVca7Emqf
# wzhLP4FdUIESQRb9NDYdGp7ekAmhz+/oD4yjvnHDzP5NsmObrz6mgFELhUglnm2d
# Xr6KJogT+Q==
# SIG # End signature block
