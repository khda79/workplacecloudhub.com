#Requires -Version 7.0
<#
.SYNOPSIS
Offline tests of the job script Version column (Operations > Jobs and Planning) and of the
Requests sub-tab under Operations. Synthetic scripts in a temporary folder; no job is launched.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$folder=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:Cases=0
function Assert-Case {param([bool]$Condition,[string]$Message) $script:Cases++; if(-not $Condition){throw $Message}}
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.GuiWorker.psm1') -Force
$insightsPath=Join-Path $folder 'SmartM365.Orchestrator.Insights.psm1'
$tokens=$null;$errors=$null
$insightsAst=[Management.Automation.Language.Parser]::ParseFile($insightsPath,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Insights parsing failed.'}
$versionFunction=$insightsAst.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-SmartM365OrchestratorScriptVersion'},$true)
if($null -eq $versionFunction){throw 'Get-SmartM365OrchestratorScriptVersion is missing from the Insights module.'}
. ([scriptblock]::Create($versionFunction.Extent.Text))
$guiPath=Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'
$guiText=[IO.File]::ReadAllText($guiPath)
$ast=[Management.Automation.Language.Parser]::ParseFile($guiPath,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'GUI parsing failed.'}
$definitions=@{}
foreach($name in @('ConvertFrom-OrchestratorGuiXaml','Get-OrchestratorGuiPropertyValue','Get-GuiJobScriptPaths','Get-GuiJobScriptVersion','Get-GuiRefreshWork','Select-GuiTab',
    'Refresh-PlanningView','Refresh-LaunchView','Receive-GuiRefresh','Refresh-AllViewsSync')){
    $definition=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
    if($null -eq $definition){throw "Missing GUI function: $name"}
    $definitions[$name]=$definition.Extent.Text
}
foreach($name in @('ConvertFrom-OrchestratorGuiXaml','Get-OrchestratorGuiPropertyValue','Get-GuiJobScriptPaths','Get-GuiJobScriptVersion','Get-GuiRefreshWork','Select-GuiTab')){ . ([scriptblock]::Create($definitions[$name])) }

$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GuiScriptVersion-Test-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $root
$lock=$null
try {
    $scripts=Join-Path $root 'Inventory'
    $null=New-Item -ItemType Directory -Path (Join-Path $scripts 'Sub')
    $utf8Bom=[Text.UTF8Encoding]::new($true)
    [IO.File]::WriteAllText((Join-Path $scripts 'Sub\Header.ps1'),"<#`r`n.SYNOPSIS`r`nSample`r`n.VERSION`r`n1.2.3`r`n#>`r`n`$ScriptVersion = '1.2.3'`r`n",$utf8Bom)
    [IO.File]::WriteAllText((Join-Path $scripts 'Inline.ps1'),"<#`n.VERSION 2.0.1-beta`n#>`n")
    [IO.File]::WriteAllText((Join-Path $scripts 'NoVersion.ps1'),"<#`n.SYNOPSIS`nNo version`n#>`n")
    [IO.File]::WriteAllText((Join-Path $scripts 'BadVersion.ps1'),"<#`n.VERSION`nnot a version; Remove-Item C:\`n#>`n")
    [IO.File]::WriteAllText((Join-Path $scripts 'Notes.txt'),"<#`n.VERSION`n9.9.9`n#>`n")
    [IO.File]::WriteAllText((Join-Path $scripts 'Late.ps1'),('#'+('x'*9000)+"`n<#`n.VERSION`n3.0.0`n#>`n"))
    [IO.File]::WriteAllText((Join-Path $scripts 'Locked.ps1'),"<#`n.VERSION`n4.0.0`n#>`n")
    [IO.File]::WriteAllText((Join-Path $scripts 'Shared.ps1'),"<#`n.VERSION`n5.0.0`n#>`n")
    # Never executed: a script whose body would write a marker file when run.
    $marker=Join-Path $root 'executed.txt'
    [IO.File]::WriteAllText((Join-Path $scripts 'Trap.ps1'),"<#`n.VERSION`n6.0.0`n#>`nSet-Content -LiteralPath '$marker' -Value run`n")

    # 1. Extraction rules.
    $map=Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $scripts -ScriptPath @('Sub\Header.ps1','Inline.ps1','NoVersion.ps1','BadVersion.ps1','Notes.txt','Late.ps1','Missing.ps1','Trap.ps1','',' ')
    Assert-Case ($map['Sub\Header.ps1'] -eq '1.2.3') 'Header .VERSION on the next line not read (UTF-8 BOM).'
    Assert-Case ($map['Inline.ps1'] -eq '2.0.1-beta') 'Inline .VERSION not read.'
    Assert-Case ($map['NoVersion.ps1'] -eq '?') 'Missing .VERSION is not shown as ?.'
    Assert-Case ($map['BadVersion.ps1'] -eq '?') 'Invalid .VERSION text reached the grid.'
    Assert-Case ($map['Notes.txt'] -eq '?') 'A non-.ps1 file was read.'
    Assert-Case ($map['Late.ps1'] -eq '?') 'The reader went past the 8 KB header window.'
    Assert-Case ($map['Missing.ps1'] -eq 'missing') 'Absent script is not shown as missing.'
    Assert-Case ($map['Trap.ps1'] -eq '6.0.0' -and -not (Test-Path -LiteralPath $marker)) 'The script was executed or not read.'
    Assert-Case ($map.Count -eq 8) 'Empty script paths produced entries.'
    Assert-Case ($map['sub\HEADER.ps1'] -eq '1.2.3') 'The map is not case-insensitive.'
    $empty=Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $scripts -ScriptPath @()
    Assert-Case ($empty -is [hashtable] -and $empty.Count -eq 0) 'Empty input did not return an empty map.'

    # 2. Files held open by another process.
    $writer=[IO.FileStream]::new((Join-Path $scripts 'Shared.ps1'),'Open','ReadWrite','Read')
    try { Assert-Case ((Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $scripts -ScriptPath 'Shared.ps1')['Shared.ps1'] -eq '5.0.0') 'A script open for writing (git pull in progress) could not be read.' }
    finally { $writer.Dispose() }
    $lock=[IO.FileStream]::new((Join-Path $scripts 'Locked.ps1'),'Open','Read','None')
    $lockedMap=Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $scripts -ScriptPath @('Locked.ps1','Inline.ps1')
    Assert-Case ($lockedMap['Locked.ps1'] -eq '?' -and $lockedMap['Inline.ps1'] -eq '2.0.1-beta') 'An exclusively locked script broke the other versions.'
    $lock.Dispose();$lock=$null

    # 3. Cache: an unchanged file is not re-read, a changed file is.
    $cache=[hashtable]::Synchronized(@{})
    $first=Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $scripts -ScriptPath 'Inline.ps1' -Cache $cache
    $inlineFull=[IO.Path]::GetFullPath((Join-Path $scripts 'Inline.ps1'))
    Assert-Case ($first['Inline.ps1'] -eq '2.0.1-beta' -and $cache.ContainsKey($inlineFull)) 'Cache entry not created.'
    $cache[$inlineFull].Version='7.7.7'
    Assert-Case ((Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $scripts -ScriptPath 'Inline.ps1' -Cache $cache)['Inline.ps1'] -eq '7.7.7') 'An unchanged script was re-read despite the cache.'
    [IO.File]::WriteAllText($inlineFull,"<#`n.VERSION 2.0.2`n#>`n")
    Assert-Case ((Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $scripts -ScriptPath 'Inline.ps1' -Cache $cache)['Inline.ps1'] -eq '2.0.2') 'A changed script was not re-read.'
    Assert-Case ((Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $scripts -ScriptPath 'Missing.ps1' -Cache $cache)['Missing.ps1'] -eq 'missing') 'Missing script with cache failed.'

    # 4. Real catalog: every job script of the shipped template declares a readable version.
    $inventoryRoot=[IO.Path]::GetFullPath((Join-Path $folder '..'))
    $template=Get-Content -LiteralPath (Join-Path $folder 'Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
    $paths=@(Get-GuiJobScriptPaths -JobsDocument $template)
    $clock=[Diagnostics.Stopwatch]::StartNew()
    $real=Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $inventoryRoot -ScriptPath $paths -Cache ([hashtable]::Synchronized(@{}))
    $coldMs=$clock.ElapsedMilliseconds
    $unreadable=@($real.Keys | Where-Object { $real[$_] -in @('?','missing') })
    Assert-Case ($paths.Count -gt 0 -and $unreadable.Count -eq 0) ('Catalog scripts without a readable version: '+($unreadable -join ', '))

    # 5. Background refresh worker: versions read off the UI thread, failure isolated.
    $mockModule=Join-Path $root 'SyntheticManagement.psm1'
    [IO.File]::WriteAllText($mockModule,($versionFunction.Extent.Text+"`n"+@'
function Get-SmartM365OrchestratorConfigurationSnapshot {param($SharedDataFolderPath) [pscustomobject]@{Jobs=[pscustomobject]@{Jobs=@([pscustomobject]@{Name='HeaderJob';ScriptPath='Sub\Header.ps1'},[pscustomobject]@{Name='NoPath'})};Cluster=[pscustomobject]@{ExpectedOrchestratorServers=@()}}}
function Get-SmartM365OrchestratorRecentRuns {param($SharedDataFolderPath,$Days) @()}
function Get-SmartM365OrchestratorServerStatus {param($SharedDataFolderPath,$ClusterDocument) @()}
function Get-SmartM365OrchestratorHistory {param($SharedDataFolderPath,$From,$To,$Server,$JobName,$Status) @()}
function Get-SmartM365OrchestratorConfigurationVersions {param($SharedDataFolderPath) @()}
function Read-SmartM365OrchestratorJson {param($Path) Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json}
function Get-SmartM365OrchestratorOperations {param($SharedDataFolderPath,$ClusterDocument,$MailFolderPath,$MailHours) [pscustomobject]@{Thread=[Threading.Thread]::CurrentThread.ManagedThreadId}}
function Get-SmartM365OrchestratorRecentPipelineRuns {param($SharedDataFolderPath,$Count) @()}
'@))
    $shared=Join-Path $root 'Shared'
    $null=New-Item -ItemType Directory -Path (Join-Path $shared 'Config')
    '{"ExpectedOrchestratorServers":[]}' | Set-Content -LiteralPath (Join-Path $shared 'Config/Orchestrator-Cluster.json')
    $templatePath=Join-Path $root 'jobs.template.json'
    '{"Jobs":[{"Name":"InlineJob","ScriptPath":"Inline.ps1"},{"Name":"MissingJob","ScriptPath":"Gone.ps1"}]}' | Set-Content -LiteralPath $templatePath
    function Invoke-FullWorker {
        param([string]$TemplatePath,[hashtable]$Cache)
        $context=[pscustomobject]@{Worker=$null;Progress=[hashtable]::Synchronized(@{Message=''});Input=$null;Clock=[Diagnostics.Stopwatch]::new()}
        $inputData=[pscustomobject]@{
            Mode='Full';Root=$shared;MailFolderPath='';ManagementModulePath=$mockModule;InsightsModulePath=$mockModule
            HistoryFrom=(Get-Date).AddDays(-1);HistoryTo=(Get-Date);HistoryServer='';HistoryJob='';HistoryStatus='';OnlyFailures=$false
            LogEventsHours=0;LogEventsCache=$null;IncludeDashboard=$false;ConfigurationProfile='';ConfigRootPath='';ConfigViewModulePath=''
            SmartInventoryRoot=$scripts;JobsTemplatePath=$TemplatePath;ScriptVersionCache=$Cache
        }
        Assert-Case ([bool](Start-SmartM365OrchestratorGuiWorker -Context $context -Work (Get-GuiRefreshWork) -InputObject $inputData)) 'Refresh worker did not start.'
        $deadline=(Get-Date).AddSeconds(30);$received=$null
        while($null -eq $received -and (Get-Date) -lt $deadline){Start-Sleep -Milliseconds 50;$received=Receive-SmartM365OrchestratorGuiWorker -Context $context}
        Assert-Case ($null -ne $received -and -not $received.Error) ('Refresh worker failed: '+$(if($received){$received.Error}else{'timeout'}))
        return @($received.Output)[0]
    }
    $workerCache=[hashtable]::Synchronized(@{})
    $result=Invoke-FullWorker -TemplatePath $templatePath -Cache $workerCache
    Assert-Case ($result.ScriptVersions -is [hashtable] -and -not $result.ScriptVersionsError) 'Worker did not return the version map.'
    Assert-Case ($result.ScriptVersions['Sub\Header.ps1'] -eq '1.2.3' -and $result.ScriptVersions['Inline.ps1'] -eq '2.0.2' -and $result.ScriptVersions['Gone.ps1'] -eq 'missing') 'Worker map misses published or catalog scripts.'
    Assert-Case ($result.Operations.Thread -ne [Threading.Thread]::CurrentThread.ManagedThreadId) 'The refresh work ran on the UI thread.'
    Assert-Case ($workerCache.Count -eq 2) 'Worker did not fill the shared cache.'
    [IO.File]::WriteAllText($templatePath,'{ not json')
    $broken=Invoke-FullWorker -TemplatePath $templatePath -Cache $workerCache
    Assert-Case ($broken.ScriptVersionsError -and $null -eq $broken.ScriptVersions -and $null -ne $broken.Operations) 'A version read failure was not isolated from the refresh.'

    # 6. GUI rows and static wiring.
    $script:ScriptVersions=$result.ScriptVersions
    Assert-Case ((Get-GuiJobScriptVersion -Job ([pscustomobject]@{ScriptPath=' sub\header.ps1 '})) -eq '1.2.3') 'Row lookup is not trimmed or case-insensitive.'
    Assert-Case ((Get-GuiJobScriptVersion -Job ([pscustomobject]@{ScriptPath='Unknown.ps1'})) -eq '' -and (Get-GuiJobScriptVersion -Job ([pscustomobject]@{Name='x'})) -eq '' -and (Get-GuiJobScriptVersion -Job $null) -eq '') 'Unknown script did not render empty.'
    Assert-Case ($definitions['Refresh-PlanningView'] -match 'Version = Get-GuiJobScriptVersion -Job \$job') 'Planning rows have no Version.'
    Assert-Case ($definitions['Refresh-LaunchView'] -match 'Version = Get-GuiJobScriptVersion -Job \$definition') 'Operations > Jobs rows have no Version.'
    $receive=$definitions['Receive-GuiRefresh']
    Assert-Case ($receive -notmatch 'Get-SmartM365OrchestratorScriptVersion' -and $receive.IndexOf('$script:ScriptVersions =') -gt 0 -and $receive.IndexOf('$script:ScriptVersions =') -lt $receive.IndexOf('Refresh-PlanningView')) 'Versions are read on the UI thread or applied after the rows.'
    Assert-Case ($definitions['Get-GuiRefreshWork'] -match "if \(\`$InputData.Mode -eq 'Full'\)[\s\S]*Get-SmartM365OrchestratorScriptVersion[\s\S]*catch \{ \`$scriptVersionsError") 'Worker version read is not limited to Full mode or not isolated.'
    Assert-Case ($definitions['Refresh-AllViewsSync'] -match 'Get-SmartM365OrchestratorScriptVersion[\s\S]*-Cache \$script:ScriptVersionCache') 'Synchronous refresh does not use the cache.'
    Assert-Case ([IO.File]::ReadAllText($insightsPath) -match "'Get-SmartM365OrchestratorScriptVersion'") 'Insights does not export the version reader.'

    # 7. XAML: Version columns and the Requests sub-tab.
    $assignment=$ast.Find({param($node)$node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml'},$true)
    $window=ConvertFrom-OrchestratorGuiXaml -Text $assignment.Right.Expression.Value
    foreach($gridName in @('LaunchJobsGrid','PlanningGrid')){
        $grid=$window.FindName($gridName)
        $headers=@($grid.Columns | ForEach-Object { [string]$_.Header })
        Assert-Case ($headers[0] -eq 'Job' -and $headers[1] -eq 'Version' -and $grid.Columns[1].Binding.Path.Path -eq 'Version') "$gridName has no Version column after Job."
    }
    $script:Controls=@{}
    foreach($name in @('MainTabs','OperationsTabs','OperationsRequestsTab','RequestsGrid','CancelRequestButton')){ $script:Controls[$name]=$window.FindName($name); if(-not $script:Controls[$name]){throw "Missing XAML control: $name"} }
    Assert-Case (@($script:Controls.MainTabs.Items | Where-Object { [string]$_.Header -eq 'Requests' }).Count -eq 0) 'Requests is still a main tab.'
    Assert-Case ($script:Controls.OperationsTabs.Items.Contains($script:Controls.OperationsRequestsTab) -and $script:Controls.RequestsGrid.Parent.Parent.Parent -eq $script:Controls.OperationsRequestsTab) 'Requests is not a sub-tab of Operations.'
    Select-GuiTab -Header 'Requests'
    Assert-Case ([string]$script:Controls.MainTabs.SelectedItem.Header -eq 'Operations' -and $script:Controls.OperationsTabs.SelectedItem -eq $script:Controls.OperationsRequestsTab) 'Navigation to Requests (dashboard action) failed.'
    Select-GuiTab -Header 'Planning'
    Assert-Case ([string]$script:Controls.MainTabs.SelectedItem.Header -eq 'Planning') 'Navigation to other tabs changed.'
    Assert-Case ($guiText -notmatch 'Requests tab' -and $guiText -match 'Operations > Requests') 'User messages still point to the old Requests tab.'
    foreach($controlList in @($ast.FindAll({param($node)$node -is [Management.Automation.Language.ArrayLiteralAst] -and $node.Extent.Text -match "'OperationsTabs'"},$true))){
        Assert-Case ($controlList.Extent.Text -match "'OperationsRequestsTab'") 'OperationsRequestsTab missing from a control list.'
    }
    "[{0}] PASS: {1} offline job script version and Requests sub-tab checks ({2} catalog scripts read in {3} ms, cold cache). No job launched." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$script:Cases,$paths.Count,$coldMs
}
finally{
    if($lock){$lock.Dispose()}
    $resolved=[IO.Path]::GetFullPath($root)
    if($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $resolved)){
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCzWTUVIuDISfz/
# sryp9eOr3seUAVN6yjm3ToKFDHlIYKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEnqTbRSodm34cSGA2RlT/3wRq7vpVaYi2y+wYqLory3MA0GCSqG
# SIb3DQEBAQUABIIBgAwoV4PDsigyIWsr/g57FAMd2cNC1uhk2o1MPbyoOHlA3RYY
# uyg1pFMxcp+TJMX4F719+C3NUH4cXixyjtR9vrEp80de9EO1lm8X7yLbzWouHhIy
# O1yDTeRqp62EzaVtp0lKlxTD2dlN4MHwKct2bRwB2g9tiRQ6+qKRb1CaMjSIkCXF
# g6yMd/fVepbuIxKjLs7gFXr3ItgQRxmRwVn1fFAA2Ism964r7gWrWpXsHK3uHNJF
# m2UWMqUxbJRe6Up/GShi5uet00+lMHimnUX8/Hmlu71YGUuKFEXqu4KyvIGvkcU7
# qB8SoKqgo/9P5YHxrUrJzfoOOlCZnuYndsUtHTbiC13NT7TmHbjakFNbl223Y2uJ
# jT6KrMErHSh3h7RN0WppUYOhy0t03BQFKlfNSl6YJtw0raLLUQkmXa3qxUUHZrzv
# 7bdKLMzGOPn/eyM5V4ptfzKe/e/MMH7SfQ8b244qqfu4fdbm/IdMU3qU76q2cFyF
# 3hE/UZuxRh2faL4LeqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMzUy
# MzhaMC8GCSqGSIb3DQEJBDEiBCBvX5SUjZX3iAZKcKcoUZH6Yh7z7iSSP779S27c
# 0hkjqTANBgkqhkiG9w0BAQEFAASCAgCM8egZZgt3e2CMqDtsVoPkwp68FG+xVctE
# 6EdnKvwD0vVKwyr5KNG+RvNTpb8+cXSxp24maxoaBpUy8tqEByRmyxLFzlA8pJDB
# 5gs4ThK+WbbKQ1sTzi8HdBsGh5egxUGKvNJkCa9V8AZFKQ0QT3O6DK9umcNwiyfJ
# UnSpL/9Pf+FzNwSRbvGKCi/QSn1QKWrRYZS21ObLTtPTR2Bq3dd+rTywQ34S6eIu
# JJd6BOEpRl6r6g4o8c/eRORT3GNRt1recLSf8sZTaGeAAoTyQ0eu31zF9upMKm5z
# atxuA2bigcRmaC+SLs7CbSJeSdd06xb32WmhblNnTG15ej1p6TeRa8JQlDctF+HD
# Jp48GkO22biTY57NkQGM2fLxmvv4Y4aXgri5Xw3hILvbwNgPucbAV2mTabe9/ENR
# 0zhJeTY60kMlJlptjEyI6gKGdJPWRYpcU/KWfPhO4gmWwQoTDC378IU7SDLcO/ey
# iLAV0kv+QtTAOXGnQiMdoGRU9jj3Yv3lJuPSzYMvIQ452sLgN3BkmioWh66qc9Aj
# wSSetSiVpR+HpYEUCHnmflSZMhUNfJVs8KTGF6ENn7jLwtuOyn1FA0zcvn9vt+8f
# i2aGQy8ccdw837U7TV1GpM+GbvROnYe2wDAdr6ZtvAXLr4eMoazEn52OPsW7kvVR
# 6Bix+GgXIw==
# SIG # End signature block
