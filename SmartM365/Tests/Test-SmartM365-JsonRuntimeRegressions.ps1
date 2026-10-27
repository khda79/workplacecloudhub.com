# Version 1.1.0
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$module=Get-Module SmartM365.JsonTransport
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Runtime-' + [guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory $root
$script:passed=0
function Check($value,$message) { if(-not $value){throw $message}; $script:passed++ }
function Reject($action,$message) { $failed=$false; try { & $action | Out-Null } catch {$failed=$true}; Check $failed $message }
& $module { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@()} } }
foreach($case in @('valid','tenant','collector','incomplete','conflict','invalid')) {
    $history=Join-Path $root "$case/Tenants/fixture/DATA-ALL/Intune/Autopilot/Devices/WeeklyHistory"
    $week=Join-Path $history '2026-W28'; $null=New-Item -ItemType Directory $week -Force
    $oldRoot='C:\previous\SMART-M365\Tenants\fixture\DATA-ALL\Intune\Autopilot\Devices\WeeklyHistory'
    if($case -eq 'tenant'){$oldRoot=$oldRoot.Replace('\fixture\','\other\')}
    if($case -eq 'collector'){$oldRoot=$oldRoot.Replace('\Autopilot\','\BIOS\')}
    $doc=[ordered]@{Week='2026-W28';HistoryLabel='SmartM365 inventory';HistoryRootPath=$oldRoot;Files=@('Inventory.csv');UpdatedAt='2026-07-10T00:00:00Z'}
    $old=Join-Path $week 'manifest.json'
    [IO.File]::WriteAllText($old,($doc|ConvertTo-Json),[Text.UTF8Encoding]::new($true))
    $hash=(Get-FileHash $old).Hash
    if($case -ne 'incomplete'){[IO.File]::WriteAllText((Join-Path $week 'Inventory.csv'),"Id`r`n42`r`n")}
    if($case -eq 'conflict'){[IO.File]::WriteAllText("$old.txt",(($doc|ConvertTo-Json).Replace('2026-07-10','2026-07-11')))}
    if($case -eq 'invalid'){[IO.File]::WriteAllText("$old.txt",'{')}
    if($case -eq 'valid') {
        $null=Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel 'SmartM365 inventory'
        Check ((Get-FileHash "$old.txt").Hash -eq $hash -and -not (Test-Path $old)) 'Relocation changed bytes or retained legacy.'
        $null=Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel 'SmartM365 inventory'
        Check ((Get-FileHash "$old.txt").Hash -eq $hash) 'Resume rewrote relocated history.'
    } else {
        Reject { Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel 'SmartM365 inventory' } "Unsafe $case accepted."
        Check ((Get-FileHash $old).Hash -eq $hash) "Failure damaged legacy $case."
    }
}
# Simulate inherited restricted PSModulePath and disabled autoload in a separate PS5 process.
$child=Join-Path $root 'NativeRuntime.ps1'
$manifest=(Resolve-Path (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1')).Path.Replace("'","''")
$policy=(Resolve-Path (Join-Path $PSScriptRoot '../Config/SmartM365-JsonTransport.policy.psd1')).Path.Replace("'","''")
$body=@"
`$ErrorActionPreference='Stop'
`$env:PSModulePath='C:\synthetic-no-modules'
`$PSModuleAutoLoadingPreference='None'
Import-Module (`$PSHOME + '/Modules/Microsoft.PowerShell.Management/Microsoft.PowerShell.Management.psd1') -Global
Import-Module '$manifest' -Force
`$p=Import-PowerShellDataFile -LiteralPath '$policy'
if(`$p.Mode -notin @('Readers','JsonText')){throw 'Policy not loaded'}
if(-not (Get-FileHash -LiteralPath '$policy').Hash){throw 'Hash unavailable'}
"@
[IO.File]::WriteAllText($child,$body)
& "$env:SystemRoot/System32/WindowsPowerShell/v1.0/powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $child
Check ($LASTEXITCODE -eq 0) 'Native PS5 utility commands unavailable.'
[pscustomobject]@{Passed=$script:passed;FixtureRoot=$root;Evidence='Synthetic only; no tenant startup, collection or remote access'}
