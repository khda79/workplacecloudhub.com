<#
.SYNOPSIS
Runs or requests a SmartInventory job, or a manual launch, from the jobs catalog.

.DESCRIPTION
Single entry point that replaces the SmartInventory .cmd launchers; SmartInventory.cmd starts it
with Windows PowerShell 5.1. The catalog is
Orchestrator-Jobs.json.template (Jobs and Launches); see SmartM365.Inventory.Launch.psm1.

-Name <Job> [-Now] runs the entry now on this server with its runtime (pwsh 7 or Windows
PowerShell 5.1) and execution mode (Direct or LocalCache), outside the orchestrator.
-Name <Job>[,<Job>...] -Request submits the jobs through SmartM365-Inventory-Pipeline.ps1: the
resident orchestrators run them on the elected server with their dependencies and locks.
-List (or no argument) shows the catalog. -WhatIf shows the resolved command without running anything.
Any other argument is passed to the launched script (or to the pipeline with -Request).

Tenant: -Tenant, then SMARTM365_ORCHESTRATOR_TENANT, then DefaultTenant of
SmartM365.global.local.json. Execution mode: -ExecutionMode, then Execution.Mode of the job,
then OnPremExecutionMode of SmartM365.global.local.json (LocalCache by default).

The exit code of the launched script is returned. Preparation errors return 1.
-List and -WhatIf show the WorkplaceCloudHub banners; a launched script shows its own.

.EXAMPLE
.\Invoke-SmartM365InventoryJob.ps1 -List

.EXAMPLE
.\Invoke-SmartM365InventoryJob.ps1 Exchange2016-Local-Mailboxes-Inventory -WhatIf

.EXAMPLE
.\Invoke-SmartM365InventoryJob.ps1 AD-Inventory -Request -IncludeDependencies

.VERSION
1.1.0
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)][string[]]$Name = @(),
    [switch]$List,
    [switch]$Now,
    [switch]$Request,
    [string]$Tenant = '',
    [ValidateSet('', 'Auto', 'Direct', 'LocalCache')][string]$ExecutionMode = '',
    [string]$JobsManifestPath = '',
    [Parameter(ValueFromRemainingArguments)][string[]]$ExtraArguments = @()
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.1.0'

function Write-Host {
    param([Parameter(Position = 0)][AllowEmptyString()][string]$Object = '', [ConsoleColor]$ForegroundColor = 'Gray')
    foreach ($line in ($Object -split "\r?\n")) {
        Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line) -ForegroundColor $ForegroundColor
    }
}

function Resolve-InventoryJobsManifestPath {
    <# Runtime jobs manifest that overrides job Arguments: the central manifest, else the local one. #>
    param([Parameter(Mandatory)]$Catalog, [Parameter(Mandatory)][string]$TenantValue)
    $orchestratorFolder = Join-Path $Catalog.SmartInventoryRoot 'Orchestrator'
    $localConfigPath = Get-SmartM365JsonReadPath -Path (Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator.local.json') -Optional
    if (-not $localConfigPath) { $localConfigPath = Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator.local.json.template' }
    $localConfig = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($localConfigPath, [Text.Encoding]::UTF8))
    $centralValue = $localConfig.PSObject.Properties['CentralConfigurationEnabled']
    $centralEnabled = ($null -eq $centralValue) -or ([string]$centralValue.Value -in @('True', 'true', '1'))
    if (-not $centralEnabled) { return (Get-SmartM365JsonReadPath -Path (Join-Path $orchestratorFolder 'Orchestrator-Jobs.json') -Optional) }

    $dataFolder = '{{DataAllRootPath}}\Orchestrator'
    $dataFolderValue = $localConfig.PSObject.Properties['OrchestratorDataFolderPath']
    if ($null -ne $dataFolderValue -and -not [string]::IsNullOrWhiteSpace([string]$dataFolderValue.Value) -and [string]$dataFolderValue.Value -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { $dataFolder = [string]$dataFolderValue.Value }
    if ($dataFolder -match '\{\{') {
        $roots = Get-SmartM365InventoryOutputRoots -SmartM365Root $Catalog.SmartM365Root -Tenant $TenantValue
        foreach ($property in $roots.PSObject.Properties) { $dataFolder = $dataFolder.Replace('{{' + $property.Name + '}}', [string]$property.Value) }
    }
    if ($dataFolder -match '\{\{') { throw "OrchestratorDataFolderPath is not resolved: $dataFolder" }
    $sharedDataFolder = $dataFolder
    if ((Split-Path -Path $sharedDataFolder -Leaf) -eq $env:COMPUTERNAME) { $sharedDataFolder = Split-Path -Path $sharedDataFolder -Parent }
    return (Get-SmartM365JsonReadPath -Path (Join-Path $sharedDataFolder 'Config\Orchestrator-Jobs.json') -Optional)
}

function Wait-InventoryLauncherKey {
    # Same as the pause of the legacy .cmd launchers: only in a user console, never unattended.
    $inputRedirected = $true
    try { $inputRedirected = [Console]::IsInputRedirected } catch { $inputRedirected = $true }
    if ([Environment]::UserInteractive -and -not $inputRedirected -and -not $env:SMARTM365_ORCHESTRATOR_TENANT) {
        [void](Read-Host -Prompt 'Press Enter to close')
    }
}

$exitCode = 0
try {
    . (Join-Path $PSScriptRoot '..\..\Config\SmartM365-TenantContext.ps1')
    $launchModuleFile = Join-Path $PSScriptRoot 'SmartM365.Inventory.Launch.psm1'
    Import-Module -Name $launchModuleFile -Force -ErrorAction Stop

    $catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot (Split-Path -Path $PSScriptRoot -Parent)
    if ($List -or @($Name).Count -eq 0) {
        Write-SmartM365StartupBanner
        Write-Host ('SmartInventory catalog: {0}' -f $catalog.TemplatePath)
        foreach ($entry in $catalog.Entries) {
            $definition = $entry.Definition
            $runtime = if ($definition.PSObject.Properties['Runtime'] -and $definition.Runtime) { [string]$definition.Runtime } elseif ($definition.PSObject.Properties['PowerShellEdition'] -and $definition.PowerShellEdition) { [string]$definition.PowerShellEdition } else { 'PowerShell7' }
            $execution = if ($definition.PSObject.Properties['Execution'] -and $null -ne $definition.Execution) { 'Execution:' + [string]$definition.Execution.CacheName } else { 'Direct' }
            $description = if ($definition.PSObject.Properties['Description']) { [string]$definition.Description } else { [string]$definition.ScriptPath }
            Write-Host ('{0,-58} {1,-6} {2,-17} {3,-42} {4}' -f $entry.Name, $entry.Kind, $runtime, $execution, $description)
        }
        Write-SmartM365CompletionBanner -Status Success -ScriptName 'Invoke-SmartM365InventoryJob'
        exit 0
    }
    if ($Now -and $Request) { throw 'Choose either -Now or -Request.' }
    $action = if ($Request) { 'Request' } else { 'Now' }

    if (-not [string]::IsNullOrWhiteSpace($JobsManifestPath) -or $action -eq 'Now') {
        $overlayPath = $JobsManifestPath
        if ([string]::IsNullOrWhiteSpace($overlayPath)) {
            try {
                $tenantForOverlay = (Resolve-SmartM365InventoryTenant -Tenant $Tenant -SmartM365Root $catalog.SmartM365Root).Value
                $overlayPath = Resolve-InventoryJobsManifestPath -Catalog $catalog -TenantValue $tenantForOverlay
            }
            catch {
                Write-Host ('[WARN] Runtime jobs manifest not resolved; template Arguments are used. {0}' -f $_.Exception.Message) -ForegroundColor Yellow
                $overlayPath = ''
            }
        }
        if ($overlayPath) { $catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot $catalog.SmartInventoryRoot -OverlayPath $overlayPath }
    }

    $plan = Resolve-SmartM365InventoryLaunch -Catalog $catalog -Name $Name -Action $action -Tenant $Tenant -ExecutionMode $ExecutionMode -ExtraArguments $ExtraArguments
    if ($WhatIfPreference) {
        Write-SmartM365StartupBanner
        Write-Host (Format-SmartM365InventoryPlan -Plan $plan -Catalog $catalog)
        Write-Host ('Catalog: {0}; runtime manifest: {1}' -f $catalog.TemplatePath, $(if ($catalog.OverlayPath) { $catalog.OverlayPath } else { 'none' }))
        Write-SmartM365CompletionBanner -Status Success -ScriptName 'Invoke-SmartM365InventoryJob'
        exit 0
    }
    $exitCode = Invoke-SmartM365InventoryLaunch -Plan $plan -Catalog $catalog
    if ($exitCode -ne 0) { Write-Host ('{0} ended with exit code {1}.' -f $plan.Name, $exitCode) -ForegroundColor Yellow }
}
catch {
    Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red
    Wait-InventoryLauncherKey
    $exitCode = 1
}
exit $exitCode

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDSIyFlBGDCjoX+
# 3Pjb/SUdzerD0lTO5mIjEVhiExh//KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJDiNb97sRj8+8ltb5eo1h0L/3px7LJlk9dRSIun9vHWMA0GCSqG
# SIb3DQEBAQUABIIBgHY/C8cSBrwsaBPIWScemzLbFZn26m20MCcx0b0CC7TjvXdL
# XLR0De9Ul6LwSXooEMMJbUJSkb7eDFoUQOsAtjlMzeZD/Msotxpol2JSYljTK7zv
# XKunBWX1Ax1v1gxmFfsnO353RUPZW2YnKfPI58QxZVQrWuyIjif6Rw40SiGRKs2E
# YClr6RePd0ljAhU1dznxvGw2Lj5UZUZPXQoLkb7M1zlIlOu8lbl3aZBe7IygT7Pl
# Yc6LNJnK5F7lL/wqi7apo5GmLa3JcNVG03z2pcEXuHUngHHljswPJX/kfEFdZJa5
# ggkq3MJ3e1ErAYmkqonzqk2dRWPnv1wGt85gbNHkLL96LWD+KZS1hxgi3mUU3cys
# UOrU4Ne63RWDo2GZ2gAKXehPe9e2oPxgPuzIXENsI2kSTz1iQAUAzhqADRAnbUQB
# 0KGPsCPgRrc3ArBwY1l6zfpp7E0XYfOKTd1x5rCi7r/+DqG2mB4JhHR8EAgazCfu
# 1WXxFdaEElc+IFHsMaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYwOTU2
# MDlaMC8GCSqGSIb3DQEJBDEiBCDp6I7vCA0uDCogl+EghpJ9ftQhG4yx1XiIzPYL
# EM+fpzANBgkqhkiG9w0BAQEFAASCAgB4Kcods4sQI6ey952kdxLk4YNBUZM7gB4M
# u150SGQW5c49OyGp5jxPrzp3KC0d+81x0b/+cb8udLduGFtJ6curuoTiEHOnmLAd
# SAoHej8RFf7u3GxqTmz3y/WLpmmggrO8Ci06A0eIg0gsLSuzde2eagR+bp4c+ARo
# /TTklyo46culyVUkmJCWsK6HzQb98+rhMwuoUkcBoF7rAx+x3s9ERgnbmoisUJev
# lxKrCNdcfVD4BIk09+D4PPJpVvc+V3Ya+DO9B1CXLjtEEOOtmmMi3Rawy6BYuBpM
# FDqCig8lBjMALigvEetF/6F1XrHu/2BgdXAixSFKqhKmaHxOsccvulUbfmYKyZAP
# RrPczb1YlnKKbs0C8KE8b7lacfeYu9PZiw9lx5+RorgNzSzUOJSAoGgHnKGq/WCN
# Eb/pJ34/+6SvE2n/evHufQoYHvYvs/zFfMcl9fL6PMDR2tbQlZdOuTnxsNZKx+Mz
# Tm0HRoBb7EeFxB1UuGgPx5ZTn5ZUb/YvTmMb+LbUik03JX9OMFKOKcLn2jY0VC2R
# xYoSU+n1ufX8fqvnUD9a5FaCH24qXnGQLOlunGoCy+xhdEnZU5YDPt2+/qsTBD6X
# 7x5YoAbx3LS40Kn1Mlf0uCwb/JJ21EzjXziC1xS4+dh4fj19X4g1L4bsX93DUz04
# lF8iiTQdww==
# SIG # End signature block
