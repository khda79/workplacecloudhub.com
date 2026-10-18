#Requires -Version 5.1
<#
.SYNOPSIS
Offline demonstration host for the original SmartM365 WPF views.
.VERSION
1.0.0
.NOTES
The same file is copied locally into each supported application. It reads only
the shipped UI, demo profile and report/template sources. No operational startup
code or event handler is imported. Profiles own the fictional data and actions.
#>

function Get-SmartM365DemoAst {
    param([string]$Path)
    $tokens=$null; $parseErrors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    if($parseErrors.Count){throw ('Demo source parse failed: '+$Path)}
    return $ast
}

function Get-SmartM365DemoMarkup {
    param([string]$Path,[string]$Variable='xaml',[string]$Root='Window')
    $ast=Get-SmartM365DemoAst $Path
    $assignment=$ast.Find({param($node)
        $node -is [Management.Automation.Language.AssignmentStatementAst] -and
        $node.Left.Extent.Text -match ('\$'+[regex]::Escape($Variable)+'$') -and
        $node.Right.Extent.Text -match ('<'+$Root+'\b')
    },$true)
    if(-not $assignment){throw "Demo UI source not found: $Variable in $Path"}
    $expression=$assignment.Right.Expression
    if($expression -isnot [Management.Automation.Language.StringConstantExpressionAst] -and
       $expression -isnot [Management.Automation.Language.ExpandableStringExpressionAst]){throw 'Demo markup must be a literal here-string.'}
    $text=$expression.Value
    # These are the only substitutions in the Analyzer's original window markup.
    $text=$text.Replace('$script:AppName','Smart Endpoint Diagnostics Analyzer').Replace('$script:AppVersion','DEMO')
    if($text -match '\$\(|\$script:'){throw 'Unsupported dynamic expression in demo markup.'}
    return $text
}

function Get-SmartM365DemoFunctionText {
    param([string]$Path,[string[]]$Names)
    $ast=Get-SmartM365DemoAst $Path
    foreach($name in $Names){
        $definition=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
        if(-not $definition){throw "Demo source function not found: $name"}
        $definition.Extent.Text
    }
}

function New-SmartM365DemoView {
    param([string]$Markup)
    $reader=[Xml.XmlNodeReader]::new([xml]$Markup)
    return [Windows.Markup.XamlReader]::Load($reader)
}

function Get-SmartM365DemoControls {
    param($View,[string]$Markup)
    $controls=@{}
    foreach($match in [regex]::Matches($Markup,'(?:x:Name|Name)="(?<Name>[A-Za-z][A-Za-z0-9_]*)"')){
        $name=$match.Groups['Name'].Value
        $control=$View.FindName($name)
        if($control){$controls[$name]=$control}
    }
    return $controls
}

function Get-SmartM365DemoLabel {
    param([string]$Name)
    $text=$Name -replace '^(lbl|txt|cmb|btn|chk|Btn)','' -replace '(Text|Box|Label|Button|Combo|Check|Value|Grid)$',''
    return ($text -creplace '([a-z])([A-Z])','$1 $2')
}

function New-SmartM365DemoRow {
    param([string[]]$Columns,[int]$Index=0,[hashtable]$Values=@{})
    $row=[ordered]@{}
    foreach($column in $Columns){
        if($Values.ContainsKey($column)){$row[$column]=$Values[$column];continue}
        $number=$Index+1
        $row[$column]=switch -Regex ($column){
            '^(Computer|ComputerName|DeviceName|Device|HostName|ConnectionTarget)$' {'CTSO-LAP-{0:0000}' -f $number;break}
            '(Email|UserPrincipal|RequestedBy|LastUser|UserName)' {'user{0:00}@contoso.example' -f $number;break}
            '(^IP|AddressList)' {'192.0.2.'+(40+$number);break}
            '^(Name|DisplayName|Title)$' {'Contoso sample '+$number;break}
            '(Count|Rows|Total|Matched|Files|Folders|Size|Length|Capacity|Occurrences|Minutes|Penalty|Score|GB|Percent)' {10+$Index;break}
            '(Enabled|Available|Online|Reachable|Joined|Compliant|IsBlocking)$' {$true;break}
            '(Timestamp|Created|Modified|Updated|Date|Time|Seen|NotAfter)' {'2026-10-18 09:30:00';break}
            '(Status|State|Result|Decision)' {@('Healthy','Warning','Ready')[$Index%3];break}
            'Severity|Level' {@('Information','Warning','Error')[$Index%3];break}
            'Brush|Color' {'#E6F4FF';break}
            'Version' {'1.0.'+$number;break}
            'Path|File|Folder|Directory' {'Demo/Evidence/sample-'+$number+'.txt';break}
            'Detail|Message|Description|Recommendation|Meaning|Action' {'Fictional Contoso evidence for demonstration.';break}
            'Source|Provider|Publisher' {'Contoso demo';break}
            '^(Id|ID|ObjectId|DeviceId)$' {'00000000-0000-4000-8000-'+$number.ToString('000000000000');break}
            default {'Sample '+$number}
        }
    }
    foreach($key in $Values.Keys){$row[$key]=$Values[$key]}
    return [pscustomobject]$row
}

function Set-SmartM365DemoGrid {
    param([string]$Name,[object[]]$Rows,[string[]]$Columns=@())
    $s=$script:SmartM365Demo
    if(-not $s.Controls.ContainsKey($Name)){throw "Demo grid not found: $Name"}
    $grid=$s.Controls[$Name]
    # Template columns bind to row properties too. Retain those original cells
    # when a profile supplies a more specific fixture for the bound columns.
    $fields=@($grid.Columns|ForEach-Object {if($_ -is [Windows.Controls.DataGridBoundColumn] -and $_.Binding.Path){[string]$_.Binding.Path.Path}})
    if($s.Data.ContainsKey('GridBindings:'+$Name)){$fields+=$s.Data['GridBindings:'+$Name]}
    $index=0
    foreach($row in $Rows){
        $missing=@($fields|Where-Object {$_ -notmatch '\.' -and $_ -notin $row.PSObject.Properties.Name}|Select-Object -Unique)
        if($missing.Count){$defaults=New-SmartM365DemoRow -Columns $missing -Index $index;foreach($property in $defaults.PSObject.Properties){$row|Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value}}
        $index++
    }
    if($grid.Columns.Count -eq 0 -or $Columns.Count -and $s.Data.ContainsKey('DynamicGrid:'+$Name)){
        $grid.Columns.Clear()
        $s.Data['DynamicGrid:'+$Name]=$true
        $grid.AutoGenerateColumns=$false
        if(-not $Columns.Count -and $Rows.Count){$Columns=@($Rows[0].PSObject.Properties.Name)}
        foreach($field in $Columns){
            $column=[Windows.Controls.DataGridTextColumn]::new()
            $column.Header=Get-SmartM365DemoLabel $field
            $column.Binding=[Windows.Data.Binding]::new($field)
            $column.Width=[Windows.Controls.DataGridLength]::new(1,[Windows.Controls.DataGridLengthUnitType]::Star)
            [void]$grid.Columns.Add($column)
        }
    }
    $s.Rows[$Name]=@($Rows)
    $grid.ItemsSource=@($Rows)
    if($Rows.Count){$grid.SelectedIndex=0}
}

function Set-SmartM365DemoCombo {
    param([string]$Name,[object[]]$Values,[int]$SelectedIndex=0)
    $control=$script:SmartM365Demo.Controls[$Name]
    $control.ItemsSource=$null; $control.Items.Clear()
    $control.ItemsSource=@($Values)
    $control.SelectedIndex=$SelectedIndex
}

function Write-SmartM365DemoActivity {
    param([string]$Message)
    $s=$script:SmartM365Demo
    # A portable prefix prevents the signed-in Windows name leaking into captures.
    $display=$Message.Replace([IO.Path]::GetTempPath(),'%TEMP%\')
    $line='[2026-10-18 09:30:00] [DEMO] '+$display
    $s.Activity.Add($line)
    foreach($name in @('ActivityBox','ActivityText','OutputBox','LogTextBox','txtLogContent')){
        if($s.Controls.ContainsKey($name) -and $s.Controls[$name] -is [Windows.Controls.TextBox]){
            if($s.Controls[$name].Text -and -not $s.Controls[$name].Text.EndsWith("`n")){$s.Controls[$name].AppendText([Environment]::NewLine)}
            $s.Controls[$name].AppendText($line+[Environment]::NewLine)
            $s.Controls[$name].ScrollToEnd()
        }
    }
    foreach($name in @('StatusText','StatusTitle','lblConfigStatus')){if($s.Controls.ContainsKey($name)){$s.Controls[$name].Text='DEMO - '+$display}}
    if($s.Controls.ContainsKey('listActivity')){$s.Controls.listActivity.ItemsSource=@($s.Activity)}
}

function Get-SmartM365DemoOutputPath {
    param([string]$Name)
    $s=$script:SmartM365Demo
    if([IO.Path]::GetFileName($Name) -ne $Name){throw 'Demo export names must be file names.'}
    [void][IO.Directory]::CreateDirectory($s.OutputRoot)
    return Join-Path $s.OutputRoot $Name
}

function Show-SmartM365DemoFile {
    param([string]$Path)
    $s=$script:SmartM365Demo
    $full=[IO.Path]::GetFullPath($Path)
    if(-not $full.StartsWith($s.OutputRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Only this session demo files can be opened.'}
    if(-not (Test-Path -LiteralPath $full -PathType Leaf)){throw 'The demo output was not generated.'}
    $s.LastOutput=$full
    Write-SmartM365DemoActivity ('Output ready: '+$full)
    if(-not $s.ValidateOnly){Start-Process -FilePath $full | Out-Null}
}

function Save-SmartM365DemoHtml {
    param([string]$Name,[string]$Title,[string]$Body)
    $path=Get-SmartM365DemoOutputPath $Name
    $safeTitle=[Net.WebUtility]::HtmlEncode($Title)
    $html='<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>'+ $safeTitle +'</title><style>body{font:15px Segoe UI,Arial;background:#f4f7fb;color:#18324a;margin:28px}main{background:white;padding:28px;border-radius:12px;max-width:1400px}table{border-collapse:collapse;width:100%}td,th{padding:10px;text-align:left;border-bottom:1px solid #dce4ed}th{background:#e7f1fb}pre{white-space:pre-wrap}a{color:#0078d4}.demo{background:#fff3cd;padding:12px;border-radius:6px}h1{margin-top:12px}</style></head><body><main><div class="demo">DEMO - Sample data | Contoso | 2026-10-18 09:30</div><h1>'+ $safeTitle +'</h1>'+ $Body +'</main></body></html>'
    [IO.File]::WriteAllText($path,$html,[Text.UTF8Encoding]::new($false))
    return $path
}

function Show-SmartM365DemoDetail {
    param([string]$Title,[string]$Text)
    $s=$script:SmartM365Demo
    $markup='<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="900" Height="650" Background="#F4F7FB" WindowStartupLocation="CenterOwner"><Grid Margin="20"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><TextBlock x:Name="DemoDetailTitle" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,14"/><TextBox x:Name="DemoDetailBody" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Padding="14"/><Button x:Name="DemoDetailClose" Grid.Row="2" Content="Close" Width="100" HorizontalAlignment="Right" Margin="0,14,0,0" Padding="10,6"/></Grid></Window>'
    $window=New-SmartM365DemoView $markup
    $window.Title=$Title+' | DEMO - Sample data'
    $window.FindName('DemoDetailTitle').Text=$Title+' - DEMO'
    $window.FindName('DemoDetailBody').Text=$Text
    $close=$window.FindName('DemoDetailClose');$close.Tag=$window
    $close.Add_Click({param($button)$button.Tag.Close()})
    [void]$s.ChildWindows.Add($window)
    if(-not $s.ValidateOnly){$window.Owner=$s.Window;$window.Show()}
    return $window
}

function Set-SmartM365DemoMarking {
    param([string]$Path)
    $html=[IO.File]::ReadAllText($Path)
    $banner='<div style="padding:12px 24px;background:#fff3cd;color:#654900;font:14px Segoe UI">DEMO - Sample data | Contoso | Simulated results; no live assessment or action performed.</div>'
    $html=[regex]::Replace($html,'(?i)<body([^>]*)>',('<body$1>'+$banner),1)
    $html=$html.Replace('Live strict','DEMO - simulated sources')
    [IO.File]::WriteAllText($Path,$html,[Text.UTF8Encoding]::new($false))
}

function Invoke-SmartM365DemoAction {
    param([string]$Name)
    $s=$script:SmartM365Demo
    $s.LastAction=$Name
    $s.ActionCounts[$Name]=1+[int]$s.ActionCounts[$Name]
    if(Invoke-SmartM365DemoProfileAction -Name $Name){return}
    if($Name -match '(?<Stem>.+?)(Down|Up)(Button)?$'){
        $stem=$matches.Stem
        $target=@($stem+'Text',$stem+'Box',$stem)
        foreach($key in $target){if($s.Controls.ContainsKey($key) -and $s.Controls[$key] -is [Windows.Controls.TextBox]){
            $number=0
            if([int]::TryParse($s.Controls[$key].Text,[ref]$number)){
                $delta=if($Name -match 'Down'){-1}else{1}
                $s.Controls[$key].Text=[string][Math]::Max(0,$number+$delta)
            }
            break
        }}
    }elseif($Name -match 'Reset|Refresh'){
        Initialize-SmartM365DemoProfile
    }elseif($Name -match 'Export.*Csv|ExportReport|SaveAllCloud|SaveCloudScripts|SaveAllButton|SaveLocalButton'){
        $rows=if($s.Rows.Count){@($s.Rows.Values|Select-Object -First 1)[0]}else{@([pscustomobject]@{Mode='DEMO';Result='Simulated'})}
        $path=Get-SmartM365DemoOutputPath ($Name+'-DEMO.csv')
        $rows|Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
        Write-SmartM365DemoActivity ('Fictional CSV exported: '+$path)
        return
    }elseif($Name -match 'Bundle|Anon'){
        $path=Get-SmartM365DemoOutputPath ($Name+'-DEMO.zip')
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        Add-Type -AssemblyName System.IO.Compression
        $stream=[IO.File]::Open($path,[IO.FileMode]::Create)
        $archive=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
        try{$entry=$archive.CreateEntry('DEMO-Summary.txt');$writer=[IO.StreamWriter]::new($entry.Open());try{$writer.Write('DEMO - Sample data. CTSO-LAP-0142. No device or user evidence was collected.')}finally{$writer.Dispose()}}finally{$archive.Dispose();$stream.Dispose()}
        Write-SmartM365DemoActivity ('Synthetic support ZIP exported: '+$path)
        return
    }
    foreach($control in $s.Controls.Values){if($control -is [Windows.Controls.ProgressBar]){$control.IsIndeterminate=$false;$control.Value=$control.Maximum}}
    Write-SmartM365DemoActivity ($Name+' simulated successfully. Changes remain in demo memory.')
}

function Initialize-SmartM365DemoDefaults {
    $s=$script:SmartM365Demo
    foreach($name in @($s.Controls.Keys)){
        $control=$s.Controls[$name]
        if($control -is [Windows.Controls.TextBox] -or $control -is [Windows.Controls.ComboBox] -or $control -is [Windows.Controls.DatePicker]){$control.IsEnabled=$true}
        if($control -is [Windows.Controls.TextBlock] -and [string]::IsNullOrWhiteSpace($control.Text)){$control.Text=Get-SmartM365DemoLabel $name}
        if($control -is [Windows.Controls.Button] -or $control -is [Windows.Controls.Primitives.ToggleButton]){
            $control.IsEnabled=$true
            if(-not $control.Content){$control.Content=Get-SmartM365DemoLabel $name}
        }
        if($control -is [Windows.Controls.TextBox] -and -not $control.Text -and $name -notmatch '(Search|Filter|Credential|ApiKey)'){
            $control.Text=if($name -match '(Limit|Days|Minutes|Timeout|Delay|Throttle|Max|Lease|Tokens|Retry|Start|End|Ipg|Jitter|Prefix).*?(Text|Box)$'){'15'}elseif($name -match 'Computer'){'CTSO-LAP-0142'}elseif($name -match '(Path|Folder|Root|Source|Map)'){'Demo/'+(Get-SmartM365DemoLabel $name)}else{'DEMO - Fictional Contoso evidence. No collection or action was performed.'}
        }
        if($control -is [Windows.Controls.ComboBox] -and $control.Items.Count -eq 0){
            $control.ItemsSource=@('All','Ready','Warning','Error');$control.SelectedIndex=0
        }elseif($control -is [Windows.Controls.ComboBox] -and $control.SelectedIndex -lt 0 -and $control.Items.Count){$control.SelectedIndex=0}
        if($control -is [Windows.Controls.DataGrid]){
            $fields=@();$booleans=@()
            foreach($column in $control.Columns){
                if($column -is [Windows.Controls.DataGridBoundColumn] -and $column.Binding -and $column.Binding.Path){
                    $field=[string]$column.Binding.Path.Path
                    if($field -notmatch '\.'){$fields+=$field;if($column -is [Windows.Controls.DataGridCheckBoxColumn]){$booleans+=$field}}
                }
            }
            if(-not $fields.Count){$fields=@('Name','Status','Details')}
            $rows=@(for($i=0;$i -lt 6;$i++){$values=@{};foreach($key in $booleans){$values[$key]=$i%3 -ne 0};New-SmartM365DemoRow -Columns $fields -Index $i -Values $values})
            Set-SmartM365DemoGrid -Name $name -Rows $rows -Columns $fields
        }
    }
    foreach($name in @('BusyOverlay','ScanOverlay','panelDiagAnalysisLoading','panelCrossCheckLoading')){if($s.Controls.ContainsKey($name)){$s.Controls[$name].Visibility='Collapsed'}}
    foreach($name in @('chkAutoRefresh')){if($s.Controls.ContainsKey($name)){$s.Controls[$name].IsChecked=$false;$s.Controls[$name].IsEnabled=$false}}
    foreach($name in @('HeaderLogo','HeaderLogoImage','imgLogo','LogoImage')){
        if($s.Controls.ContainsKey($name)){
            $path=Join-Path $s.AssetRoot 'WorkplaceCloudHub-lockup-WPF.png'
            if(Test-Path -LiteralPath $path){$s.Controls[$name].Source=[Windows.Media.Imaging.BitmapFrame]::Create([uri]$path);$s.Controls[$name].Visibility='Visible';if($s.Controls.ContainsKey('HeaderLogoFallback')){$s.Controls.HeaderLogoFallback.Visibility='Collapsed'}}
            $s.Controls[$name].Cursor='Hand';$s.Controls[$name].ToolTip='Open WorkplaceCloudHub.com'
            $s.Controls[$name].Add_MouseLeftButtonUp({if(-not $script:SmartM365Demo.ValidateOnly){Start-Process 'https://workplacecloudhub.com'}})
        }
    }
}

function Start-SmartM365GuiDemo {
    [CmdletBinding()]
    param([string]$GuiScriptPath,[string]$ProfilePath,[ValidateSet('Overview','Issues','Completed')][string]$Scenario='Overview',[string]$View='',
        [ValidateRange(1000,3840)][int]$Width=1500,[ValidateRange(650,2160)][int]$Height=950,[switch]$ValidateOnly)
    $ErrorActionPreference='Stop'
    if([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA){throw 'The demo requires STA. Use its Demo CMD launcher.'}
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    $markup=Get-SmartM365DemoMarkup -Path $GuiScriptPath
    $markup=$markup.Replace('LIVE STRICT','DEMO - SIMULATED').Replace('Live strict','Demo - simulated')
    $window=New-SmartM365DemoView $markup
    $window.Title+=' | DEMO - Sample data'
    $window.WindowState='Normal';$window.Width=$Width;$window.Height=$Height
    $window.Topmost=$false;$window.UseLayoutRounding=$true
    $folder=Split-Path $GuiScriptPath -Parent
    $assetRoot=if((Split-Path $folder -Leaf) -eq 'Scripts'){Split-Path $folder -Parent}else{$folder}
    $script:SmartM365Demo=[pscustomobject]@{
        Mode='Demo';Scenario=$Scenario;Window=$window;Controls=(Get-SmartM365DemoControls $window $markup)
        GuiPath=$GuiScriptPath;AppRoot=$folder;AssetRoot=$assetRoot;ValidateOnly=[bool]$ValidateOnly
        OutputRoot=(Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Demo/'+[IO.Path]::GetFileNameWithoutExtension($GuiScriptPath)+'/'+[guid]::NewGuid().ToString('N')))
        Rows=@{};Data=@{};Reports=@{};Activity=[Collections.Generic.List[string]]::new();ActionCounts=@{};LastAction='';LastOutput=''
        ChildWindows=[Collections.Generic.List[object]]::new();ViewNames=[Collections.Generic.List[string]]::new();Tabs=[Collections.Generic.List[object]]::new()
    }
    $s=$script:SmartM365Demo
    foreach($node in ([xml]$markup).SelectNodes('//*[local-name()="DataGrid"]')){
        $name=$node.GetAttribute('Name','http://schemas.microsoft.com/winfx/2006/xaml')
        if(-not $name){$name=$node.GetAttribute('Name')}
        $s.Data['GridBindings:'+$name]=@([regex]::Matches($node.OuterXml,'\{Binding\s+(?:Path=)?(?<Field>[A-Za-z][A-Za-z0-9_]*)(?:[,\s}]|$)')|ForEach-Object {$_.Groups['Field'].Value}|Select-Object -Unique)
    }
    . (Join-Path (Split-Path $ProfilePath -Parent) 'SmartM365.GuiDemo.Reports.ps1')
    . $ProfilePath
    Initialize-SmartM365DemoDefaults
    Initialize-SmartM365DemoProfile
    $body=$window.Content;$window.Content=$null
    $dock=[Windows.Controls.DockPanel]::new()
    $bar=[Windows.Controls.StackPanel]::new();$bar.Orientation='Horizontal';$bar.Margin='16,8,16,0'
    [Windows.Controls.DockPanel]::SetDock($bar,'Top')
    $label=[Windows.Controls.TextBlock]::new();$label.Text='DEMO - Sample data | Scene';$label.VerticalAlignment='Center';$label.Margin='0,0,12,0'
    [void]$bar.Children.Add($label)
    $scene=[Windows.Controls.ComboBox]::new();$scene.Width=170;$scene.ItemsSource=@('Overview','Issues','Completed');$scene.SelectedItem=$Scenario
    $scene.Add_SelectionChanged({$script:SmartM365Demo.Scenario=[string]$script:SmartM365Demo.Controls.DemoScenario.SelectedItem;Initialize-SmartM365DemoProfile;Write-SmartM365DemoActivity ('Scene loaded: '+$script:SmartM365Demo.Scenario)})
    $s.Controls.DemoScenario=$scene;[void]$bar.Children.Add($scene)
    $export=[Windows.Controls.Button]::new();$export.Content='Export demo outputs';$export.Margin='12,0,0,0';$export.Padding='12,4'
    $export.Add_Click({[void](Export-SmartM365DemoProfileOutputs);Write-SmartM365DemoActivity ('All demo outputs are ready: '+$script:SmartM365Demo.OutputRoot)})
    $s.Controls.DemoExportOutputs=$export;[void]$bar.Children.Add($export)
    $open=[Windows.Controls.Button]::new();$open.Content='Open demo outputs';$open.Margin='8,0,0,0';$open.Padding='12,4'
    $open.Add_Click({$s=$script:SmartM365Demo;if(-not (Test-Path -LiteralPath $s.OutputRoot)){[void](Export-SmartM365DemoProfileOutputs)};if(-not $s.ValidateOnly){Start-Process -FilePath $s.OutputRoot|Out-Null}})
    $s.Controls.DemoOpenOutputs=$open;[void]$bar.Children.Add($open)
    [void]$dock.Children.Add($bar);[void]$dock.Children.Add($body);$window.Content=$dock
    foreach($name in @($s.Controls.Keys)){
        $control=$s.Controls[$name]
        if($name -like 'Demo*'){continue}
        if($control -is [Windows.Controls.Button] -or $control -is [Windows.Controls.Primitives.ToggleButton] -and $control -isnot [Windows.Controls.CheckBox] -and $control -isnot [Windows.Controls.RadioButton]){
            $control.Tag=$name
            $control.Add_Click({param($demoControl)try{Invoke-SmartM365DemoAction -Name ([string]$demoControl.Tag)}catch{Write-SmartM365DemoActivity ('Demo action failed: '+$_.Exception.Message)}})
        }
    }
    Connect-SmartM365DemoProfileEvents
    $pending=[Collections.Generic.Stack[object]]::new();$pending.Push($window)
    while($pending.Count){
        $element=$pending.Pop()
        if($element -is [Windows.Controls.TabControl]){[void]$s.Tabs.Add($element)}
        if($element -is [Windows.DependencyObject]){foreach($child in [Windows.LogicalTreeHelper]::GetChildren($element)){if($child -is [Windows.DependencyObject]){$pending.Push($child)}}}
    }
    foreach($control in $s.Tabs){foreach($tab in $control.Items){if($tab -is [Windows.Controls.TabItem]){[void]$s.ViewNames.Add([string]$tab.Header);if($View -and $tab.Header -eq $View){$control.SelectedItem=$tab}}}}
    if($View){foreach($name in @($s.Controls.Keys|Where-Object {$_ -like 'tab*'})){$control=$s.Controls[$name];if($control -is [Windows.Controls.Primitives.ToggleButton] -and ($name -eq $View -or [string]$control.Content -eq $View)){Invoke-SmartM365DemoAction $name}}}
    $icon=Join-Path $assetRoot 'WorkplaceCloudHub.ico'
    if(Test-Path -LiteralPath $icon){$window.Icon=[Windows.Media.Imaging.BitmapFrame]::Create([uri]$icon)}
    Write-SmartM365DemoActivity 'Fictional Contoso data loaded. All business actions are simulated; no operational session was initialized.'
    if($ValidateOnly){return $s}
    $splash=$null
    try{
        $splashPath=Join-Path $assetRoot 'SmartM365.GuiSplash.ps1'
        if(Test-Path -LiteralPath $splashPath){
            . $splashPath
            $splash=Start-SmartM365GuiSplash -ProductName ($window.Title -replace ' \|.*$','') -Subtitle 'DEMO - Sample data' -MinimumDurationMs 6000
            $script:SmartM365DemoSplash=$splash
            $window.Add_ContentRendered({Hide-SmartM365GuiSplash -Splash $script:SmartM365DemoSplash})
        }
        [void]$window.ShowDialog()
    }finally{
        foreach($child in $s.ChildWindows){if($child.IsVisible){$child.Close()}}
        if($splash){Close-SmartM365GuiSplash -Splash $splash}
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAib1EKLW/x++MK
# LtqVYn+o/VTKDOR+w1ygdfOxSHfKdaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFBcfYGc15q7JjUg4V1KJUbq4Wu5hhUPfr98SxpDuwsTMA0GCSqG
# SIb3DQEBAQUABIIBgKWxTE3sGS49cHWUktDSQcyBKxfRUna0lBNONvU78cgS0/Cx
# OaikGnTFOJC9n8ye/RpNmoo2DsbNy/rHrsx4+XUCb+LRKy9UTq50zqlTOJlpBk6b
# 1ocFGDxRLapzIB57FfkXioJrArT6iLhcqmhle8u7pR0TG2P5NciLeAa5+J0fFqDl
# wyXINHokAjDnJu4eLz6BZyRu3lKuo6aSTkg9JnDDjJTtnQxOMQv2ts1Vbv9ZR5do
# x2SyAOGWmmCWZytYOMFA/sMm9CULV5/OMXFrWSIaH3eSnwtErhM3AcaL/d/XkAuP
# /ATsCmx0mjQCZrQtWbOBYSfbbU301iVu8d03LOAai6VlNGWuz2qxsFwY2b4oTpaR
# pbBtPKK1PzWQCq7Ozxqpy270DQ+qlHrd0E62FQBu7QiMrWjixwDwGOfJpCn1iaU2
# I45llZHHeNxtZBhCjYs0vauuFraXc1GOWdwjgwChKTJ1H1MfXPK0GZWZcldFfn8H
# reJra82sUyWffz5RjqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMDQ2
# MzJaMC8GCSqGSIb3DQEJBDEiBCB89UDHnLvBNrwUL3L1McipJ88AQnUCO7CdlRO0
# wcqMtDANBgkqhkiG9w0BAQEFAASCAgCqQRhl2cQ8333tTveBdAlmLsGcAaQ5EGZI
# EP+Z64e9cvo3xmdd9w6k7x7RjkivzTh5lweiykdcHI51SHn/fdxfHia8IB41jeCD
# QY1shhZEmiX6NPUawewP8TM+EcJixC/8EheM5g5dnTioQHTf4AvCsPVHixZd/nOB
# KjdK+jn3LuUVWs7de76sHqhkiVNSJo9gw/18DXEvgSXFVYNlIhk/C1vJhcot/Cn4
# dJzOPaocOiUS442tAK7goqep7wNCjg/9w570JAnr4xTc+aFngAglxr/u49FpkKo8
# rCswAU00zTxtXYKmBVG+uKXPbnik1YBReTxGIdbc4/Em/wJOMXcieYPgicaXacsO
# YvwPGau/xMLAwRnb721lNbNdlmL4eewe9p1mJ1x03Qu+ZEeJJYXjxq8pxuETo3dA
# 9tpHAYC51yJsYvI1IE1e3RGRXkH4V+rGNMCvnWtnQjF9AXggHDsE2Lsi04zL1dM5
# 1nhqp3sgMGchLDESrVeOvog4szgdJUKHB5EtfBPBuPtCULK0dG8jFXbPv13JfyzD
# MFogIhGSQOPRE4tlKeSlHTcfMPzexucFBX2QsqUGpsDHIEU61B01jGMv/u0++pzG
# B3zwqxWQ6WyWmDDatuXz4NXWszxvfJF7wP/BHJ3wMhnh0ZExgDymko8NIDX5Ahpu
# xn74R+JQnA==
# SIG # End signature block
