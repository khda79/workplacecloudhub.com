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
