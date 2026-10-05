<#
.SYNOPSIS
    Batch launcher UI and command helpers for the migration dashboard.
.VERSION
    1.0.1
.DESCRIPTION
    Function-only library. Uses existing batch launchers and small batch summaries;
    never reads inventories or connects to a tenant to populate the dashboard.
#>

function Get-SmartM365BatchDefinition {
    param([ValidateSet('Source','Target','Comparison')][string]$Kind)
    $folder = @{ Source='source-scan-batches'; Target='target-scan-batches'; Comparison='comparison-batches' }[$Kind]
    $suffix = @{ Source='SourceScanBatch'; Target='TargetScanBatch'; Comparison='ComparisonBatch' }[$Kind]
    [pscustomobject]@{ Launcher="Start-SmartM365-SharePointMigration-$suffix.cmd"; LogFolder=$folder }
}

function Get-SmartM365BatchCommand {
    param(
        [ValidateSet('Source','Target','Comparison')][string]$Kind,
        [string]$Root,
        [ValidateSet('FilesOnly','PermissionsOnly','Both')][string]$Mode='Both',
        [ValidateSet('Interactive','Certificate')][string]$AuthMode='Interactive',
        [string[]]$Names=@(), [switch]$PlanOnly,
        [ValidatePattern('^(?:\d{8}-\d{6}-[a-f0-9]{8})?$')][string]$BatchId=''
    )
    if (-not $Root -or $Root -match '[%"!^\r\n]') { throw 'The toolkit path contains characters unsupported by the CMD launcher.' }
    foreach ($name in $Names) {
        if (-not $name -or $name -match '[,%"!^\r\n]' -or $name -ne $name.Trim()) { throw 'A migration name contains characters unsupported by the batch launcher.' }
    }
    $definition = Get-SmartM365BatchDefinition $Kind
    $launcher = Join-Path ([IO.Path]::GetFullPath($Root)) $definition.Launcher
    $modeParameter = if ($Kind -eq 'Comparison') { 'ComparisonMode' } else { 'InventoryMode' }
    $parts = @('"' + $launcher + '"', "-$modeParameter $Mode")
    if ($Kind -ne 'Source') { $parts += "-AuthMode $AuthMode" }
    if ($Kind -eq 'Target') { $parts += '-MaxParallel 2' }
    if ($Names.Count) { $parts += '-MigrationNames "' + ($Names -join ',') + '"' }
    if ($BatchId) { $parts += "-BatchId $BatchId" }
    if ($PlanOnly) { $parts += '-PlanOnly' }
    $command = $parts -join ' '
    [pscustomobject]@{ Launcher=$launcher; Command=$command; CmdArguments=('/d /v:off /s /c "' + $command + '"') }
}

function Read-SmartM365BatchTail {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        [void]$stream.Seek([Math]::Max(0, $stream.Length-8192), [IO.SeekOrigin]::Begin)
        $reader = [IO.StreamReader]::new($stream)
        try { $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

function Get-SmartM365BatchResult {
    param([string]$Root, [ValidateSet('Source','Target','Comparison')][string]$Kind, [string]$BatchId='')
    $parent = Join-Path $Root ('Migrations\logs\' + (Get-SmartM365BatchDefinition $Kind).LogFolder)
    $directory = $null
    if ($BatchId) { $path=Join-Path $parent $BatchId; if (Test-Path -LiteralPath $path -PathType Container) { $directory=Get-Item -LiteralPath $path } }
    elseif (Test-Path -LiteralPath $parent -PathType Container) {
        $directory = Get-ChildItem -LiteralPath $parent -Directory | Where-Object Name -Match '^\d{8}-\d{6}-[a-f0-9]{8}$' | Sort-Object Name -Descending | Select-Object -First 1
    }
    if (-not $directory) { return [pscustomobject]@{ Directory=$parent; Summary=''; State='No batch result yet.'; Success=0; Failed=0; Total=0; Date='' } }
    $summary = Join-Path $directory.FullName 'summary.csv'
    # Keep an array even before the summary exists or when it has just one row.
    $rows = @(if (Test-Path -LiteralPath $summary -PathType Leaf) { Import-Csv -LiteralPath $summary -ErrorAction Stop })
    $success = @($rows | Where-Object Status -eq 'Success').Count
    $failed = @($rows | Where-Object Status -ne 'Success').Count
    $tail = Read-SmartM365BatchTail (Join-Path $directory.FullName 'batch.log')
    $complete = $tail -match 'Batch finished:'
    # CSVs may be written during a comparison. A CSV alone is not completion evidence.
    $state = if (-not $complete) { 'Running or incomplete; review batch.log.' } elseif ($failed -or $tail -match 'Batch finished:.*(?:[1-9]\d* failed|interrupted=True)') { 'Completed with errors.' } else { 'Completed.' }
    $date = [datetime]::ParseExact($directory.Name.Substring(0,15), 'yyyyMMdd-HHmmss', [Globalization.CultureInfo]::InvariantCulture).ToString('yyyy-MM-dd HH:mm')
    [pscustomobject]@{ Directory=$directory.FullName; Summary=$(if (Test-Path -LiteralPath $summary) { $summary } else { '' }); State=$state; Success=$success; Failed=$failed; Total=$rows.Count; Date=$date }
}

function New-SmartM365BatchView {
    $cards = foreach ($kind in @('Source','Target','Comparison')) {
        $title = @{ Source='SOURCE SCANS'; Target='DESTINATION SCANS'; Comparison='COMPARISONS' }[$kind]
        $description = @{ Source='Files, then permissions. One scan at a time. Run on a source SharePoint farm server.'; Target='Files, then permissions. Up to two scans at once in either authentication mode.'; Comparison='Files, then permissions. One comparison at a time, using existing scan inventories.' }[$kind]
        $column = @{ Source=0; Target=1; Comparison=2 }[$kind]
        $extra = if ($kind -eq 'Source') {
@'
<TextBlock Text="Source toolkit path (use the shared path when copying the command)" TextWrapping="Wrap" Margin="0,14,0,5"/>
<TextBox x:Name="batchSourceRoot" MinHeight="30" Padding="6"/>
<TextBlock x:Name="batchSourcePrerequisites" Text="Requires: Windows PowerShell 5.1, SharePoint snap-in, elevated account with SharePoint Shell/database access, readable toolkit and Python 3. Check on the source server before running." TextWrapping="Wrap" Foreground="#866000" Margin="0,8,0,8"/>
<WrapPanel><Button x:Name="batchSourceCheck" Content="Check prerequisites" Style="{DynamicResource BtnGhost}" Margin="0,0,6,0"/><Button x:Name="batchSourceCopy" Content="Copy command" Style="{DynamicResource BtnGhost}"/></WrapPanel>
'@
        } else {
@"
<TextBlock Text="Authentication" Margin="0,14,0,5"/>
<ComboBox x:Name="batch${kind}Auth" Height="30" SelectedIndex="0"><ComboBoxItem Content="Interactive"/><ComboBoxItem Content="Certificate"/></ComboBox>
<TextBlock Text="$(if ($kind -eq 'Target') {'Interactive sign-in opens separate scan consoles. The limit of two applies to this batch.'} else {'Requires both source and destination scans. Permission comparisons may request Entra sign-in.'})" TextWrapping="Wrap" Foreground="#5F6B7A" Margin="0,10,0,0"/>
"@
        }
@"
<Border Grid.Column="$column" Style="{DynamicResource StepCard}" Margin="$(if ($column -lt 2) {'0,0,12,0'} else {'0'})" VerticalAlignment="Stretch">
 <Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
 <StackPanel>
  <TextBlock Text="$title" Style="{DynamicResource SectionLabel}"/>
  <TextBlock Text="$description" TextWrapping="Wrap" MinHeight="52"/>
  <TextBlock Text="Include" Margin="0,14,0,5"/>
  <ComboBox x:Name="batch${kind}Mode" Height="30" SelectedIndex="0"><ComboBoxItem Content="Both" Tag="Both"/><ComboBoxItem Content="Files" Tag="FilesOnly"/><ComboBoxItem Content="Permissions" Tag="PermissionsOnly"/></ComboBox>
  $extra
 </StackPanel>
 <StackPanel Grid.Row="2" Margin="0,22,0,0">
  <WrapPanel><Button x:Name="batch${kind}Preview" Content="Preview plan" Style="{DynamicResource BtnGhost}" Margin="0,0,6,6"/><Button x:Name="batch${kind}Run" Content="Run batch" Style="{DynamicResource Btn}" Margin="0,0,0,6"/></WrapPanel>
  <TextBlock x:Name="batch${kind}Status" Text="Ready to preview." TextWrapping="Wrap" Foreground="#0078D4" FontSize="13" Margin="0,5,0,12"/>
  <Border Background="#F4F8FC" CornerRadius="6" Padding="10"><StackPanel>
   <TextBlock Text="LATEST BATCH" Style="{DynamicResource SectionLabel}"/>
   <TextBlock x:Name="batch${kind}Latest" Text="No batch result yet." TextWrapping="Wrap"/>
   <TextBlock x:Name="batch${kind}Counts" Text="Successful actions: —   Failed actions: —" TextWrapping="Wrap" Margin="0,8,0,0" FontWeight="SemiBold"/>
  </StackPanel></Border>
  <WrapPanel Margin="0,12,0,0"><Button x:Name="batch${kind}Logs" Content="Open logs" Style="{DynamicResource BtnGhost}" Margin="0,0,6,6"/><Button x:Name="batch${kind}Summary" Content="Open summary.csv" Style="{DynamicResource BtnGhost}" Margin="0,0,0,6"/></WrapPanel>
 </StackPanel></Grid>
</Border>
"@
    }
    [xml]$markup = @"
<StackPanel xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TextBlock.FontSize="12" TextBlock.Foreground="#233346">
 <TextBlock Text="BATCH RUNS" Style="{DynamicResource SectionLabel}"/>
 <Border Style="{DynamicResource StepCard}"><StackPanel>
  <TextBlock Text="Migration scope" FontWeight="SemiBold" FontSize="14"/>
  <TextBlock Text="Independent of the migration selected in the header. Preview the plan before running a batch." TextWrapping="Wrap" Foreground="#5F6B7A" Margin="0,6,0,10"/>
  <CheckBox x:Name="batchAll" Content="All configured migrations" IsChecked="True"/>
  <ListBox x:Name="batchNames" SelectionMode="Multiple" Visibility="Collapsed" MaxHeight="130" Margin="0,10,0,0"/>
  <TextBlock x:Name="batchScope" Margin="0,8,0,0" Foreground="#0078D4"/>
 </StackPanel></Border>
 <Grid Margin="0,8,0,0"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>$($cards -join "`n")</Grid>
 <TextBlock Text="Logs and summary.csv contain one result per action. Closing the dashboard does not stop a running batch. Parallel limits apply per batch; avoid starting another destination batch on another machine at the same time." TextWrapping="Wrap" Foreground="#5F6B7A" Margin="0,14,0,0"/>
</StackPanel>
"@
    [Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($markup))
}

function Sync-SmartM365BatchMigrations {
    param([string[]]$Names)
    if (-not $script:BatchGui) { return }
    $list = $script:BatchGui.View.FindName('batchNames')
    $oldNames = @($list.Items | ForEach-Object { [string]$_ })
    $newNames = @($Names | Sort-Object -Unique)
    if (($oldNames -join "`n") -ne ($newNames -join "`n")) {
        $selected = @($list.SelectedItems | ForEach-Object { [string]$_ })
        $list.Items.Clear()
        foreach ($name in $newNames) { [void]$list.Items.Add($name); if ($name -in $selected) { [void]$list.SelectedItems.Add($name) } }
    }
    Update-SmartM365BatchControls
}

function Get-SmartM365BatchSelection {
    param([string]$Kind, [switch]$PlanOnly, [string]$BatchId='')
    $v=$script:BatchGui.View
    $all=$v.FindName('batchAll').IsChecked
    $names = @(if (-not $all) { $v.FindName('batchNames').SelectedItems | ForEach-Object { [string]$_ } })
    if (-not $all -and $names.Count -eq 0) { throw 'Select at least one migration, or choose All configured migrations.' }
    $root = if ($Kind -eq 'Source') { $v.FindName('batchSourceRoot').Text.Trim() } else { $script:BatchGui.Root }
    $auth = if ($Kind -eq 'Source') { 'Interactive' } else { [string]$v.FindName("batch${Kind}Auth").SelectedItem.Content }
    @{ Kind=$Kind; Root=$root; Mode=[string]$v.FindName("batch${Kind}Mode").SelectedItem.Tag; AuthMode=$auth; Names=$names; PlanOnly=[bool]$PlanOnly; BatchId=$BatchId }
}

function Update-SmartM365BatchControls {
    if (-not $script:BatchGui) { return }
    $v=$script:BatchGui.View
    $all=$v.FindName('batchAll').IsChecked
    $v.FindName('batchNames').Visibility=if($all){'Collapsed'}else{'Visible'}
    $count=if($all){$v.FindName('batchNames').Items.Count}else{$v.FindName('batchNames').SelectedItems.Count}
    $v.FindName('batchScope').Text="$count migration(s) in scope. Files and permissions run as separate actions."
    foreach ($kind in @('Source','Target','Comparison')) {
        $busy=$script:BatchGui.Active.ContainsKey($kind)
        $v.FindName("batch${kind}Preview").IsEnabled=($count -gt 0 -and -not $busy)
        $v.FindName("batch${kind}Run").IsEnabled=($count -gt 0 -and -not $busy -and ($kind -ne 'Source' -or $script:BatchGui.SourceReady))
        $v.FindName("batch${kind}Mode").IsEnabled=-not $busy
        if($kind -ne 'Source'){$v.FindName("batch${kind}Auth").IsEnabled=-not $busy}
    }
    $v.FindName('batchSourceCheck').IsEnabled=(-not $script:BatchGui.Probe -and -not $script:BatchGui.Active.ContainsKey('Source'))
    $v.FindName('batchSourceRoot').IsEnabled=-not $script:BatchGui.Active.ContainsKey('Source')
}

function Start-SmartM365BatchFromGui {
    param([string]$Kind, [switch]$PlanOnly)
    try {
        if($script:BatchGui.Active.ContainsKey($Kind)){throw 'This batch is already running from the dashboard.'}
        if($Kind -eq 'Source' -and -not $PlanOnly -and -not $script:BatchGui.SourceReady){throw 'Check source prerequisites on this server before running.'}
        $id='{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'),[guid]::NewGuid().ToString('N').Substring(0,8)
        $request=Get-SmartM365BatchSelection $Kind -PlanOnly:$PlanOnly -BatchId $id
        $invocation=Get-SmartM365BatchCommand @request
        if(-not (Test-Path -LiteralPath $invocation.Launcher -PathType Leaf)){throw "Launcher not found: $($invocation.Launcher)"}
        $folder=Join-Path $script:BatchGui.Root 'Migrations\logs\batch-gui-runs'
        [void](New-Item -ItemType Directory -Path $folder -Force)
        $path=Join-Path $folder "$id.json.txt"
        $request.Status='Starting'; $request.StartedUtc=[DateTimeOffset]::UtcNow.ToString('o'); $request.Computer=$env:COMPUTERNAME
        $request | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding utf8
        $worker=Join-Path $script:BatchGui.HelperRoot 'SmartM365-SharePointMigration-BatchGuiRun.ps1'
        $process=Start-Process -FilePath (Join-Path $PSHOME 'pwsh.exe') -ArgumentList ('-NoExit -NoProfile -File "{0}" -RequestPath "{1}"' -f $worker,$path) -WindowStyle Normal -PassThru
        $script:BatchGui.Active[$Kind]=@{ Path=$path; Process=$process; Id=$id; Root=$request.Root; PlanOnly=[bool]$PlanOnly; Started=Get-Date }
        $script:BatchGui.View.FindName("batch${Kind}Status").Text=if($PlanOnly){'Opening preview console…'}else{'Starting batch in a separate console…'}
        Update-SmartM365BatchControls
    } catch { $script:BatchGui.View.FindName("batch${Kind}Status").Text=$_.Exception.Message }
}

function Start-SmartM365SourceProbe {
    $v=$script:BatchGui.View
    $script:BatchGui.SourceReady=$false
    try {
        $selection=Get-SmartM365BatchSelection Source -PlanOnly
        $command=Get-SmartM365BatchCommand @selection
        if(-not (Test-Path -LiteralPath $command.Launcher -PathType Leaf)){throw 'Source launcher not found in the toolkit path.'}
        # Read-only SharePoint checks in Windows PowerShell 5.1 plus a log-write probe.
        $probeCode="`$root='" + $selection.Root.Replace("'","''") + "'`r`n" + @'
$ErrorActionPreference='Stop'
try {
 if($PSVersionTable.PSVersion.Major -ne 5){throw 'Windows PowerShell 5.1 is required.'}
 $env:PSModulePath=(Join-Path $PSHOME 'Modules')+[IO.Path]::PathSeparator+$env:PSModulePath
 $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
 $principal=New-Object Security.Principal.WindowsPrincipal($identity)
 if(-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Run the dashboard elevated on the source farm server.'}
 Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
 $null=Get-SPWebApplication -ErrorAction Stop
 foreach($relative in @('Scripts\Launchers\Generic\SmartM365-SharePointMigration-SourceScanBatch.ps1','Scripts\Inventory\SmartM365-SharePointSource-FileInventory.ps1','Scripts\Inventory\SmartM365-SharePointSource-PermissionInventory.ps1','Scripts\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1','Scripts\Compare\scan_evidence.py','Scripts\console_lifecycle.py')){
  $path=Join-Path $root $relative
  if(-not (Test-Path -LiteralPath $path -PathType Leaf)){throw ('Required toolkit dependency not found: '+$relative)}
  if($relative.EndsWith('.ps1') -and (Get-AuthenticodeSignature -LiteralPath $path).Status -ne 'Valid'){throw ('Toolkit signature is not trusted: '+$relative)}
 }
 $testFile=Join-Path $root ('Migrations\logs\source-probe-'+[guid]::NewGuid().ToString('N')+'.tmp')
 $null=New-Item -ItemType Directory -Path (Split-Path $testFile) -Force
 try{[IO.File]::WriteAllText($testFile,'write access probe')}finally{if(Test-Path -LiteralPath $testFile){Remove-Item -LiteralPath $testFile -Force}}
 $python=Join-Path $root 'Tools\Python\python.exe'
 if(-not (Test-Path -LiteralPath $python -PathType Leaf)){
  $available=@(Get-Command python,py -ErrorAction SilentlyContinue | Where-Object { $_.Source -notlike '\\*' })
  if(-not $available.Count){throw 'Python 3 is required (portable Tools\Python or a local installation).'}
  $command=$available[0]
  if($command.Name -eq 'py.exe' -or $command.Name -eq 'py'){& $command.Source -3 --version *> $null}else{& $command.Source --version *> $null}
  if($LASTEXITCODE -ne 0){throw 'The installed Python command cannot start.'}
 }
 Write-Output ('READY: SharePoint Shell/database access verified for ' + $identity.Name)
 exit 0
}catch{Write-Output ('NOT READY: ' + $_.Exception.Message);exit 1}
'@
        $base=Join-Path $env:TEMP ('SmartM365-SourceProbe-' + [guid]::NewGuid().ToString('N'))
        $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probeCode))
        $exe=Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $process=Start-Process $exe -ArgumentList "-NoProfile -NonInteractive -EncodedCommand $encoded" -WindowStyle Hidden -RedirectStandardOutput "$base.out" -RedirectStandardError "$base.err" -PassThru
        $script:BatchGui.Probe=@{Process=$process; Base=$base; Started=Get-Date; Root=$selection.Root}
        $v.FindName('batchSourcePrerequisites').Text='Checking local SharePoint farm prerequisites…'
    } catch { $v.FindName('batchSourcePrerequisites').Text='Not ready: ' + $_.Exception.Message }
    Update-SmartM365BatchControls
}

function Refresh-SmartM365BatchGui {
    param([switch]$Force)
    if(-not $script:BatchGui){return}
    $v=$script:BatchGui.View
    if($script:BatchGui.Probe){
        $p=$script:BatchGui.Probe
        if(-not $p.Process.HasExited -and ((Get-Date)-$p.Started).TotalSeconds -gt 30){$p.Process.Kill();$p.Process.WaitForExit()}
        if($p.Process.HasExited){
            $output=if(Test-Path -LiteralPath ($p.Base+'.out')){[IO.File]::ReadAllText($p.Base+'.out').Trim()}else{''}
            $ready=($p.Process.ExitCode -eq 0 -and $output -match '^READY:' -and $v.FindName('batchSourceRoot').Text.Trim() -eq $p.Root)
            $script:BatchGui.SourceReady=$ready
            $v.FindName('batchSourcePrerequisites').Text=if($v.FindName('batchSourceRoot').Text.Trim() -ne $p.Root){'Toolkit path changed during the check. Check prerequisites again.'}elseif($ready){$output + '. Signed dependencies and log write access checked. Local Python staging is validated again by the batch.'}elseif($output){$output}else{'Not ready: prerequisite check timed out or could not start. Run on the source farm server.'}
            foreach($file in @($p.Base+'.out',$p.Base+'.err')){if(Test-Path -LiteralPath $file){Remove-Item -LiteralPath $file -Force}}
            $p.Process.Dispose();$script:BatchGui.Probe=$null
        }
    }
    foreach($kind in @($script:BatchGui.Active.Keys)){
        $run=$script:BatchGui.Active[$kind]
        try{
            $receipt=Get-Content -LiteralPath $run.Path -Raw | ConvertFrom-Json
            $elapsed=[int]((Get-Date)-$run.Started).TotalSeconds
            $v.FindName("batch${kind}Status").Text="$($receipt.Status) · ${elapsed}s · see the batch console."
            if($receipt.Status -in @('Succeeded','Failed','Previewed')){
                $v.FindName("batch${kind}Status").Text="$($receipt.Status) · exit $($receipt.ExitCode)" + $(if($receipt.Error){' · '+$receipt.Error}else{''})
                $script:BatchGui.Active.Remove($kind);$Force=$true
            } elseif($run.Process.HasExited){
                $v.FindName("batch${kind}Status").Text='Console closed before completion was confirmed; review logs.'
                $script:BatchGui.Active.Remove($kind);$Force=$true
            }
        }catch{$v.FindName("batch${kind}Status").Text='Waiting for batch status…'}
    }
    if($Force -or ((Get-Date)-$script:BatchGui.LastRefresh).TotalSeconds -ge 5){
        $script:BatchGui.LastRefresh=Get-Date
        foreach($kind in @('Source','Target','Comparison')){
            try{
                $root=if($kind -eq 'Source'){$v.FindName('batchSourceRoot').Text.Trim()}else{$script:BatchGui.Root}
                $active=$script:BatchGui.Active[$kind]
                $identity=if($active -and -not $active.PlanOnly){$active.Id}else{''}
                $result=Get-SmartM365BatchResult $root $kind -BatchId $identity
                $script:BatchGui.Results[$kind]=$result
                $v.FindName("batch${kind}Latest").Text=($result.Date+' · '+$result.State).Trim(' ','·')
                $v.FindName("batch${kind}Counts").Text=if($result.Total){"Successful actions: $($result.Success)   Failed actions: $($result.Failed)"}else{'Successful actions: —   Failed actions: —'}
                $v.FindName("batch${kind}Logs").IsEnabled=Test-Path -LiteralPath $result.Directory -PathType Container
                $v.FindName("batch${kind}Summary").IsEnabled=[bool]$result.Summary
            }catch{$v.FindName("batch${kind}Latest").Text='Could not read batch results: '+$_.Exception.Message;$v.FindName("batch${kind}Summary").IsEnabled=$false}
        }
    }
    Update-SmartM365BatchControls
}

function Initialize-SmartM365BatchGui {
    param($Panel, [string]$Root, [string]$SourceRoot, [switch]$ValidateOnly)
    $view=New-SmartM365BatchView
    $Panel.Content=$view
    $script:BatchGui=@{View=$view;Root=$Root;HelperRoot=$PSScriptRoot;Active=@{};Results=@{};SourceReady=$false;Probe=$null;LastRefresh=[datetime]::MinValue}
    $view.FindName('batchSourceRoot').Text=$SourceRoot
    foreach($kind in @('Source','Target','Comparison')){
        foreach($suffix in @('Mode','Run','Preview','Logs','Summary','Status','Latest','Counts')){if(-not $view.FindName("batch${kind}${suffix}")){throw "Missing batch control: $kind$suffix"}}
        $view.FindName("batch${kind}Run").IsEnabled=$false
        $view.FindName("batch${kind}Summary").IsEnabled=$false
        $view.FindName("batch${kind}Logs").IsEnabled=$false
    }
    if($ValidateOnly){return $view}
    $view.FindName('batchAll').Add_Click({Update-SmartM365BatchControls})
    $view.FindName('batchNames').Add_SelectionChanged({Update-SmartM365BatchControls})
    $view.FindName('batchSourceRoot').Add_TextChanged({$script:BatchGui.SourceReady=$false;$script:BatchGui.View.FindName('batchSourcePrerequisites').Text='Toolkit path changed. Check prerequisites again on the source farm server.';Update-SmartM365BatchControls})
    $view.FindName('batchSourceCheck').Add_Click({Start-SmartM365SourceProbe})
    $view.FindName('batchSourceCopy').Add_Click({try{$selection=Get-SmartM365BatchSelection Source;$command=Get-SmartM365BatchCommand @selection;[Windows.Clipboard]::SetText($command.Command);$script:BatchGui.View.FindName('batchSourceStatus').Text='Command copied. Run it on the source farm server.'}catch{$script:BatchGui.View.FindName('batchSourceStatus').Text=$_.Exception.Message}})
    foreach($kind in @('Source','Target','Comparison')){
        foreach($suffix in @('Preview','Run','Logs','Summary')){$view.FindName("batch${kind}${suffix}").Tag=$kind}
        $view.FindName("batch${kind}Preview").Add_Click({param($sender,$eventArgs) Start-SmartM365BatchFromGui ([string]$sender.Tag) -PlanOnly})
        $view.FindName("batch${kind}Run").Add_Click({param($sender,$eventArgs) Start-SmartM365BatchFromGui ([string]$sender.Tag)})
        $view.FindName("batch${kind}Logs").Add_Click({param($sender,$eventArgs) try{Start-Process explorer.exe -ArgumentList ('"'+$script:BatchGui.Results[[string]$sender.Tag].Directory+'"')}catch{[void][Windows.MessageBox]::Show($_.Exception.Message,'Batch logs')}})
        $view.FindName("batch${kind}Summary").Add_Click({param($sender,$eventArgs) try{Invoke-Item -LiteralPath $script:BatchGui.Results[[string]$sender.Tag].Summary}catch{[void][Windows.MessageBox]::Show($_.Exception.Message,'Batch summary')}})
    }
    $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[timespan]::FromSeconds(1)
    $timer.Add_Tick({if($script:BatchGui.Panel.IsVisible -or $script:BatchGui.Active.Count -or $script:BatchGui.Probe){Refresh-SmartM365BatchGui}})
    $script:BatchGui.Panel=$Panel;$script:BatchGui.Timer=$timer;$timer.Start()
}

function Stop-SmartM365BatchGui {
    if(-not $script:BatchGui){return}
    if($script:BatchGui.Timer){$script:BatchGui.Timer.Stop()}
    if($script:BatchGui.Probe){$process=$script:BatchGui.Probe.Process;if(-not $process.HasExited){$process.Kill()};$process.Dispose()}
    # Batch consoles are independent. Never kill them when the GUI closes.
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCChyLGoB7IOl1Qn
# JjtpchggNtjtyztj2EWoQWXk3wlJnKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBgzsDTyHYVTySDA5M7JPyq
# F31dBYG4fbCASNdyvyXUfTANBgkqhkiG9w0BAQEFAASCAYBWy0cazpf50RPPT3k6
# TE3LLvyYFqmLxWAJPEa+bfGUf3DO1TW2DVQMD5YV5SARRKYCgtJ4T9IzoW+ERGiL
# yZgHneXama62a+fVwnNXsnk+vwbFe48WvD/JjM8VgQzoCNVsMQwJL7UDPbAZvyJ+
# 1FqZ7/7T+gy958qaw+XiVcDjWPaPWJ63noUqUOv3yLSUCscCJXNoUVzu7olfB1yl
# KpJj0M/11aZRHwW0pYQVc1uErymgJSCCPYAMyVX7tqgAUAURjSn5YhDXHLR6o055
# e9PZw3I90D5BFBFrwBwHVu5J4EBD4xB2eAR4lTHnLbviSLmo1AhDEQFP/ICbDsRC
# GcJv1ZROVP99dF0AqBRKmM6Bs4E7uBP1IIU+kc4v2AXlkxqg/Tr1bkuRQ6GI/MtC
# lyEUozk05CgOd+6kDa+ospzde3wy8QcHlrrWLbkq4axSG1ApioDxxXF9RUF1Farg
# 67a2ZQhac8eDhi2UpQIGoPHJ2xxPjvFWxXAYLlzN4q6S9E8=
# SIG # End signature block
