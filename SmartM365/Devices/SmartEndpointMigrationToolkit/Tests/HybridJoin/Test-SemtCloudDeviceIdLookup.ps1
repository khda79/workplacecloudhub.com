<#
.SYNOPSIS
    Offline tests: scoped Intune and Entra exports find a device by its Entra device ID when its
    cloud name differs from the computer name (simulated Microsoft Graph module).
.VERSION
    1.0.0
#>
#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-deviceid-test-' + [guid]::NewGuid().ToString('N'))
$oldModulePath = $env:PSModulePath
$passed = 0
try {
    $moduleDir = Join-Path $testRoot 'Modules\Microsoft.Graph.Authentication'
    New-Item -ItemType Directory -Path $moduleDir -Force | Out-Null
    $moduleSource = @'
function Get-MgContext { @{ AuthType='Delegated'; TenantId='22222222-2222-2222-2222-222222222222'; Account='tech@example.test' } }
function Invoke-MgGraphRequest {
    param($Method,$Uri,$Headers)
    $decoded = [System.Uri]::UnescapeDataString([string]$Uri)
    if ($decoded -like '*managedDevices*') {
        if ($decoded -match "azureADDeviceId eq '33333333-3333-3333-3333-333333333333'") {
            return @{ value=@(@{ id='m1'; deviceName='CLOUD-OTHER-NAME'; azureADDeviceId='33333333-3333-3333-3333-333333333333'; lastSyncDateTime='2026-11-05T10:00:00Z' }) }
        }
        if ($decoded -match "deviceName eq 'PC-BYNAME'") {
            return @{ value=@(@{ id='m2'; deviceName='PC-BYNAME'; azureADDeviceId='44444444-4444-4444-4444-444444444444'; lastSyncDateTime='2026-11-05T10:00:00Z' }) }
        }
        return @{ value=@() }
    }
    if ($decoded -match "deviceId eq '33333333-3333-3333-3333-333333333333'") {
        return @{ value=@(@{ id='e1'; displayName='CLOUD-OTHER-NAME'; deviceId='33333333-3333-3333-3333-333333333333'; trustType='ServerAd'; alternativeSecurityIds=@(@{ type=2 }) }) }
    }
    return @{ value=@() }
}
Export-ModuleMember -Function Get-MgContext,Invoke-MgGraphRequest
'@
    Set-Content -LiteralPath (Join-Path $moduleDir 'Microsoft.Graph.Authentication.psm1') -Value $moduleSource -Encoding UTF8
    New-ModuleManifest -Path (Join-Path $moduleDir 'Microsoft.Graph.Authentication.psd1') -RootModule 'Microsoft.Graph.Authentication.psm1' -ModuleVersion '999.0.0' -FunctionsToExport @('Get-MgContext','Invoke-MgGraphRequest')
    $env:PSModulePath = (Join-Path $testRoot 'Modules') + ';' + $oldModulePath

    $computers = Join-Path $testRoot 'Computers.txt'
    Set-Content -LiteralPath $computers -Value @('PC-RENAMED.example.test','PC-BYNAME','PC-ABSENT') -Encoding ASCII
    $ids = Join-Path $testRoot 'DeviceIds.csv'
    Set-Content -LiteralPath $ids -Value @('ComputerName,DeviceId','PC-RENAMED.example.test,33333333-3333-3333-3333-333333333333','PC-ABSENT,55555555-5555-5555-5555-555555555555','PC-BYNAME,not-a-guid') -Encoding ASCII

    $intuneCsv = Join-Path $testRoot 'DevicesIntune.csv'
    & (Join-Path $root 'Scripts\SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1') -NoConnect -SkipModuleInstall -ComputerListPath $computers -DeviceIdListPath $ids -OutputPath $intuneCsv | Out-Null
    $intune = @(Import-Csv -LiteralPath $intuneCsv)
    $renamed = @($intune | Where-Object { $_.ComputerName -eq 'PC-RENAMED' })
    if ($renamed.Count -ne 1 -or $renamed[0].IntuneInventoryPresent -ne 'True' -or $renamed[0].AzureADDeviceId -ne '33333333-3333-3333-3333-333333333333' -or $renamed[0].LookupMatchedBy -ne 'DeviceId' -or $renamed[0].DeviceName -ne 'CLOUD-OTHER-NAME') { throw 'Intune export did not find the renamed device by its Entra device ID.' }
    $byName = @($intune | Where-Object { $_.ComputerName -eq 'PC-BYNAME' })
    if ($byName.Count -ne 1 -or $byName[0].LookupMatchedBy -ne 'Name') { throw 'Intune name match was not kept.' }
    $absent = @($intune | Where-Object { $_.ComputerName -eq 'PC-ABSENT' })
    if ($absent.Count -ne 1 -or $absent[0].IntuneInventoryPresent -ne 'False') { throw 'An absent device was reported present.' }
    Write-Host 'PASS: Intune scoped export finds a renamed device by Entra device ID'; $passed++

    $entraCsv = Join-Path $testRoot 'DevicesEntra.csv'
    & (Join-Path $root 'Scripts\SmartM365-EndpointMigration-Export-EntraDevicesCsv.ps1') -NoConnect -SkipModuleInstall -ComputerListPath $computers -DeviceIdListPath $ids -OutputPath $entraCsv | Out-Null
    $entra = @(Import-Csv -LiteralPath $entraCsv)
    $renamed = @($entra | Where-Object { $_.ComputerName -eq 'PC-RENAMED' })
    if ($renamed.Count -ne 1 -or $renamed[0].EntraInventoryPresent -ne 'True' -or $renamed[0].DeviceId -ne '33333333-3333-3333-3333-333333333333' -or $renamed[0].LookupMatchedBy -ne 'DeviceId' -or $renamed[0].DisplayName -ne 'CLOUD-OTHER-NAME') { throw 'Entra export did not find the renamed device by its device ID.' }
    if (@($entra | Where-Object { $_.ComputerName -eq 'PC-ABSENT' -and $_.EntraInventoryPresent -eq 'False' }).Count -ne 1) { throw 'An absent Entra device was reported present.' }
    Write-Host 'PASS: Entra scoped export finds a renamed device by device ID'; $passed++

    $hj = [IO.File]::ReadAllText((Join-Path $root 'Campaigns\HybridJoin\SmartEndpointMigration.Campaign.HybridJoin.psm1'))
    if (([regex]::Matches($hj, '-DeviceIdListPath \$deviceIdPath')).Count -ne 2 -or $hj -notmatch "Cycle\{0\}_CloudRefreshDeviceIds\.csv") { throw 'Post-cycle confirmation does not pass the local device IDs.' }
    Write-Host 'PASS: post-cycle Intune and Entra confirmation pass the local device IDs'; $passed++
}
finally {
    $env:PSModulePath = $oldModulePath
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host ("Summary: Passed={0}; Failed=0" -f $passed)
