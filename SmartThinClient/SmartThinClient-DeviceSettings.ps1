# Version: 1.1.4
# Native, current-user device controls for an Explorer-free workspace.
function Initialize-SmartThinClientDeviceSettings {
    if('SmartThinClient.DeviceSettings' -as [type]){return}
    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms,System.Drawing
    Add-Type -TypeDefinition @'
using System;using System.Collections.Generic;using System.Runtime.InteropServices;
[assembly:DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
namespace SmartThinClient {
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] internal struct DevMode {
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=32)]public string Name;
  public ushort Spec,Driver,Size,Extra;public uint Fields;public int X,Y;public uint Orientation,FixedOutput;
  public short Color,Duplex,YResolution,TTOption,Collate;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=32)]public string Form;
  public ushort LogPixels;public uint Bits,Width,Height,Flags,Frequency,ICMMethod,ICMIntent,MediaType,Dither,Reserved1,Reserved2,PanningWidth,PanningHeight;
 }
 public sealed class DisplayChoice {
  internal DevMode Native;
  public uint Width {get{return Native.Width;}}public uint Height {get{return Native.Height;}}public uint Frequency {get{return Native.Frequency;}}
  public override string ToString(){return Width+" x "+Height+" / "+Frequency+" Hz";}
 }
 public sealed class WifiChoice {public Guid InterfaceId;public string Profile;public override string ToString(){return Profile;}}
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] internal struct WlanInterface {public Guid Id;[MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)]public string Description;public uint State;}
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] internal struct WlanProfile {[MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)]public string Name;public uint Flags;}
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] internal struct WlanConnection {public uint Mode;[MarshalAs(UnmanagedType.LPWStr)]public string Profile;public IntPtr Ssid,Bssids;public uint BssType,Flags;}
 [ComImport,Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] internal interface IMMEnumerator {
  [PreserveSig]int EnumAudioEndpoints(int flow,uint mask,out IntPtr devices);[PreserveSig]int GetDefaultAudioEndpoint(int flow,int role,out IMMDevice device);[PreserveSig]int GetDevice([MarshalAs(UnmanagedType.LPWStr)]string id,out IMMDevice device);[PreserveSig]int RegisterEndpointNotificationCallback(IntPtr client);[PreserveSig]int UnregisterEndpointNotificationCallback(IntPtr client);
 }
 [ComImport,Guid("D666063F-1587-4E43-81F1-B948E807363F"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] internal interface IMMDevice {
  [PreserveSig]int Activate(ref Guid iid,uint context,IntPtr parameters,[MarshalAs(UnmanagedType.IUnknown)]out object instance);[PreserveSig]int OpenPropertyStore(uint access,out IntPtr store);[PreserveSig]int GetId([MarshalAs(UnmanagedType.LPWStr)]out string id);[PreserveSig]int GetState(out uint state);
 }
 [ComImport,Guid("5CDF2C82-841E-4546-9722-0CF74078229A"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] internal interface IEndpointVolume {
  [PreserveSig]int RegisterControlChangeNotify(IntPtr callback);[PreserveSig]int UnregisterControlChangeNotify(IntPtr callback);[PreserveSig]int GetChannelCount(out uint count);
  [PreserveSig]int SetMasterVolumeLevel(float value,ref Guid context);[PreserveSig]int SetMasterVolumeLevelScalar(float value,ref Guid context);[PreserveSig]int GetMasterVolumeLevel(out float value);[PreserveSig]int GetMasterVolumeLevelScalar(out float value);
  [PreserveSig]int SetChannelVolumeLevel(uint channel,float value,ref Guid context);[PreserveSig]int SetChannelVolumeLevelScalar(uint channel,float value,ref Guid context);[PreserveSig]int GetChannelVolumeLevel(uint channel,out float value);[PreserveSig]int GetChannelVolumeLevelScalar(uint channel,out float value);
  [PreserveSig]int SetMute([MarshalAs(UnmanagedType.Bool)]bool value,ref Guid context);[PreserveSig]int GetMute([MarshalAs(UnmanagedType.Bool)]out bool value);
 }
 public sealed class AudioState {public float Volume;public bool Muted;}
 public static class DeviceSettings {
  [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)]static extern bool EnumDisplaySettings(string device,int mode,ref DevMode data);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int ChangeDisplaySettingsEx(string device,ref DevMode data,IntPtr hwnd,uint flags,IntPtr parameters);
  [DllImport("wlanapi.dll")]static extern uint WlanOpenHandle(uint version,IntPtr reserved,out uint negotiated,out IntPtr handle);
  [DllImport("wlanapi.dll")]static extern uint WlanCloseHandle(IntPtr handle,IntPtr reserved);
  [DllImport("wlanapi.dll")]static extern uint WlanEnumInterfaces(IntPtr handle,IntPtr reserved,out IntPtr list);
  [DllImport("wlanapi.dll")]static extern uint WlanGetProfileList(IntPtr handle,ref Guid id,IntPtr reserved,out IntPtr list);
  [DllImport("wlanapi.dll")]static extern uint WlanConnect(IntPtr handle,ref Guid id,ref WlanConnection data,IntPtr reserved);
  [DllImport("wlanapi.dll")]static extern void WlanFreeMemory(IntPtr memory);
  [DllImport("bthprops.cpl",SetLastError=true)]static extern IntPtr BluetoothFindFirstRadio(ref uint size,out IntPtr radio);
  [DllImport("bthprops.cpl")]static extern bool BluetoothFindRadioClose(IntPtr find);
  [DllImport("kernel32.dll")]static extern bool CloseHandle(IntPtr handle);
  [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)]static extern bool OpenPrinter(string name,out IntPtr printer,IntPtr defaults);
  [DllImport("winspool.drv",SetLastError=true)]static extern bool ClosePrinter(IntPtr printer);
  [DllImport("winspool.drv",SetLastError=true)]static extern bool PrinterProperties(IntPtr owner,IntPtr printer);
  [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)]public static extern bool SetDefaultPrinter(string name);
  static void Check(int hr){if(hr<0)Marshal.ThrowExceptionForHR(hr);}
  static void Win32(uint error){if(error!=0)throw new System.ComponentModel.Win32Exception(unchecked((int)error));}
  public static DisplayChoice CurrentDisplay(){var mode=new DevMode();mode.Size=(ushort)Marshal.SizeOf(typeof(DevMode));if(!EnumDisplaySettings(null,-1,ref mode))throw new InvalidOperationException("Display information is unavailable.");return new DisplayChoice {Native=mode};}
  public static DisplayChoice[] DisplayModes(){var result=new List<DisplayChoice>();var seen=new HashSet<string>();var current=CurrentDisplay();for(int i=0;i<10000;i++){var mode=new DevMode();mode.Size=(ushort)Marshal.SizeOf(typeof(DevMode));if(!EnumDisplaySettings(null,i,ref mode))break;if(mode.Bits==current.Native.Bits&&mode.Width>0&&mode.Height>0){var item=new DisplayChoice{Native=mode};if(seen.Add(item.ToString()))result.Add(item);}}return result.ToArray();}
  public static void SetDisplay(DisplayChoice choice,bool persist){if(choice==null)throw new ArgumentNullException("choice");var mode=choice.Native;mode.Fields=0x00040000|0x00080000|0x00100000|0x00400000;int test=ChangeDisplaySettingsEx(null,ref mode,IntPtr.Zero,2,IntPtr.Zero);if(test!=0)throw new InvalidOperationException("This display mode is not available (Windows result "+test+").");int result=ChangeDisplaySettingsEx(null,ref mode,IntPtr.Zero,persist?1u:0u,IntPtr.Zero);if(result!=0)throw new InvalidOperationException("Display mode could not change (Windows result "+result+").");}
  public static WifiChoice[] WifiProfiles(){var result=new List<WifiChoice>();IntPtr client=IntPtr.Zero,interfaces=IntPtr.Zero;uint version;uint status=WlanOpenHandle(2,IntPtr.Zero,out version,out client);if(status!=0)return result.ToArray();try{Win32(WlanEnumInterfaces(client,IntPtr.Zero,out interfaces));int count=Marshal.ReadInt32(interfaces);for(int i=0;i<count;i++){var item=(WlanInterface)Marshal.PtrToStructure(IntPtr.Add(interfaces,8+i*Marshal.SizeOf(typeof(WlanInterface))),typeof(WlanInterface));IntPtr profiles=IntPtr.Zero;try{Win32(WlanGetProfileList(client,ref item.Id,IntPtr.Zero,out profiles));int n=Marshal.ReadInt32(profiles);for(int p=0;p<n;p++){var profile=(WlanProfile)Marshal.PtrToStructure(IntPtr.Add(profiles,8+p*Marshal.SizeOf(typeof(WlanProfile))),typeof(WlanProfile));result.Add(new WifiChoice{InterfaceId=item.Id,Profile=profile.Name});}}finally{if(profiles!=IntPtr.Zero)WlanFreeMemory(profiles);}}}finally{if(interfaces!=IntPtr.Zero)WlanFreeMemory(interfaces);WlanCloseHandle(client,IntPtr.Zero);}return result.ToArray();}
  public static void ConnectWifi(WifiChoice choice){if(choice==null)throw new ArgumentNullException("choice");IntPtr client;uint version;Win32(WlanOpenHandle(2,IntPtr.Zero,out version,out client));try{var connection=new WlanConnection{Mode=0,Profile=choice.Profile,BssType=3};var id=choice.InterfaceId;Win32(WlanConnect(client,ref id,ref connection,IntPtr.Zero));}finally{WlanCloseHandle(client,IntPtr.Zero);}}
  static AudioState Audio(float? volume,bool? mute){object enumerator=null,endpoint=null;IMMDevice device=null;try{enumerator=Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")));Check(((IMMEnumerator)enumerator).GetDefaultAudioEndpoint(0,1,out device));var iid=new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");Check(device.Activate(ref iid,23,IntPtr.Zero,out endpoint));var control=(IEndpointVolume)endpoint;var context=Guid.Empty;if(volume.HasValue){if(Single.IsNaN(volume.Value)||volume.Value<0||volume.Value>1)throw new ArgumentOutOfRangeException("volume");Check(control.SetMasterVolumeLevelScalar(volume.Value,ref context));}if(mute.HasValue)Check(control.SetMute(mute.Value,ref context));var result=new AudioState();Check(control.GetMasterVolumeLevelScalar(out result.Volume));Check(control.GetMute(out result.Muted));return result;}finally{if(endpoint!=null)Marshal.FinalReleaseComObject(endpoint);if(device!=null)Marshal.FinalReleaseComObject(device);if(enumerator!=null)Marshal.FinalReleaseComObject(enumerator);}}
  public static AudioState GetAudio(){return Audio(null,null);}public static void SetAudio(float volume,bool muted){Audio(volume,muted);}
  public static bool HasBluetoothRadio(){uint size=4;IntPtr radio;IntPtr find=BluetoothFindFirstRadio(ref size,out radio);try{return find!=IntPtr.Zero;}finally{if(radio!=IntPtr.Zero)CloseHandle(radio);if(find!=IntPtr.Zero)BluetoothFindRadioClose(find);}}
  public static void ShowPrinterProperties(string name,IntPtr owner){IntPtr printer;if(!OpenPrinter(name,out printer,IntPtr.Zero))throw new System.ComponentModel.Win32Exception();try{if(!PrinterProperties(owner,printer)){int error=Marshal.GetLastWin32Error();if(error!=0)throw new System.ComponentModel.Win32Exception(error);}}finally{ClosePrinter(printer);}}
 }
}
'@ -ErrorAction Stop
}

function Show-SmartThinClientDeviceSettings {
    param([ValidateSet('Network','Display','Sound','DateAndTime','Bluetooth','Printers')][string]$Page,$Owner,[switch]$PreviewMode)
    Initialize-SmartThinClientDeviceSettings
    [xml]$xaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Smart ThinClient Device Settings" Width="620" Height="550" MinWidth="520" MinHeight="440" WindowStartupLocation="CenterOwner" Background="#F5F8FB" FontFamily="Segoe UI">
 <Grid Margin="18"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,12"><StackPanel><TextBlock Name="Heading" FontSize="23" FontWeight="SemiBold" Foreground="#1F2937"/><TextBlock Text="Live device settings. Changes affect this device, including in Preview." TextWrapping="Wrap" Foreground="#5F6B7A" Margin="0,8,0,0"/></StackPanel></Border>
  <Border Grid.Row="1" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="16"><ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel Name="Body"/></ScrollViewer></Border>
  <Grid Grid.Row="2" Margin="0,12,0,0"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="100"/></Grid.ColumnDefinitions><TextBlock Name="Status" Foreground="#5F6B7A" TextWrapping="Wrap" Margin="0,0,12,0"/><Button Name="CloseButton" Grid.Column="1" Content="Close" Height="36" IsCancel="True" IsDefault="True"/></Grid>
 </Grid>
</Window>
'@
    $window=[Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml));[void]$window.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    if($Owner){$window.Owner=$Owner;$window.Icon=$Owner.Icon}
    $titles=@{Network='Network';Display='Display';Sound='Sound';DateAndTime='Date and time';Bluetooth='Bluetooth';Printers='Printers'}
    $window.FindName('Heading').Text=$titles[$Page];$window.Title='Smart ThinClient - '+$titles[$Page]
    $body=$window.FindName('Body');$status=$window.FindName('Status');$window.FindName('CloseButton').Add_Click({$window.Close()})
    $addText={param([string]$Text)$item=[Windows.Controls.TextBlock]::new();$item.Text=$Text;$item.TextWrapping='Wrap';$item.Foreground='#334155';$item.Margin='0,0,0,12';[void]$body.Children.Add($item);$item}
    $addButton={param([string]$Text)$item=[Windows.Controls.Button]::new();$item.Content=$Text;$item.Height=36;$item.Margin='0,4,0,8';$item.HorizontalAlignment='Left';$item.Padding='14,0';[void]$body.Children.Add($item);$item}
    $fail={param($ErrorRecord)$status.Text='Unable to change settings: '+$ErrorRecord.Exception.Message;$status.Foreground='#B42318'}
    $displayState=[pscustomobject]@{Pending=$false;Original=$null;Selected=$null;Remaining=0};$displayTimer=[Windows.Threading.DispatcherTimer]::new();$displayTimer.Interval=[TimeSpan]::FromSeconds(1)
    try {
        switch($Page){
            'Network' {
                $interfaces=@([Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()|Where-Object NetworkInterfaceType -ne Loopback)
                foreach($adapter in $interfaces){
                    $ip=$adapter.GetIPProperties();$addresses=@($ip.UnicastAddresses|Where-Object {$_.Address.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork}|ForEach-Object {$_.Address.IPAddressToString})
                    [void](& $addText ($adapter.Name+' | '+$adapter.OperationalStatus+[Environment]::NewLine+'IPv4: '+($addresses -join ', ')+[Environment]::NewLine+'DNS: '+(($ip.DnsAddresses|ForEach-Object {$_.IPAddressToString}) -join ', ')))
                }
                if(-not $interfaces.Count){[void](& $addText 'No network adapter was detected.')}
                [void](& $addText 'Saved Wi-Fi networks');$profiles=@([SmartThinClient.DeviceSettings]::WifiProfiles())
                if($profiles.Count){$wifi=[Windows.Controls.ComboBox]::new();$wifi.ItemsSource=$profiles;$wifi.SelectedIndex=0;$wifi.Margin='0,0,0,10';[void]$body.Children.Add($wifi);$connect=& $addButton 'Connect';$connect.Add_Click({try{[SmartThinClient.DeviceSettings]::ConnectWifi($wifi.SelectedItem);$status.Text='Connection requested. The current network may disconnect.'}catch{& $fail $_}})}else{[void](& $addText 'No saved Wi-Fi network is available. Ethernet details are shown above.')}
                $status.Text='IP/DNS changes and adding new Wi-Fi profiles require an administrator-managed network configuration.'
            }
            'Display' {
                $current=[SmartThinClient.DeviceSettings]::CurrentDisplay();[void](& $addText ('Current primary display resolution: '+$current))
                if([Windows.Forms.SystemInformation]::TerminalServerSession){[void](& $addText 'Resolution is controlled by your Remote Desktop connection. Resize or configure the RDP client to change it.')}else{
                    $modes=@([SmartThinClient.DeviceSettings]::DisplayModes());$mode=[Windows.Controls.ComboBox]::new();$mode.ItemsSource=$modes;$mode.SelectedIndex=0;foreach($choice in $modes){if($choice.ToString() -eq $current.ToString()){$mode.SelectedItem=$choice;break}};$mode.Margin='0,0,0,12';[void]$body.Children.Add($mode)
                    $test=& $addButton 'Test resolution';$keep=& $addButton 'Keep resolution';$keep.IsEnabled=$false
                    $test.Add_Click({try{$displayState.Original=[SmartThinClient.DeviceSettings]::CurrentDisplay();$displayState.Selected=$mode.SelectedItem;[SmartThinClient.DeviceSettings]::SetDisplay($displayState.Selected,$false);$displayState.Pending=$true;$displayState.Remaining=15;$status.Text='Keep this resolution? Reverting in 15 seconds.';$keep.IsEnabled=$true;$test.IsEnabled=$false;$displayTimer.Start()}catch{& $fail $_}})
                    $keep.Add_Click({try{[SmartThinClient.DeviceSettings]::SetDisplay($displayState.Selected,$true);$displayState.Pending=$false;$displayTimer.Stop();$keep.IsEnabled=$false;$test.IsEnabled=$true;$status.Text='Display mode saved.'}catch{& $fail $_}})
                    $displayTimer.Add_Tick({$displayState.Remaining--;if($displayState.Remaining -le 0){try{[SmartThinClient.DeviceSettings]::SetDisplay($displayState.Original,$false);$status.Text='Previous display mode restored.'}catch{& $fail $_};$displayState.Pending=$false;$displayTimer.Stop();$keep.IsEnabled=$false;$test.IsEnabled=$true}else{$status.Text='Keep this resolution? Reverting in '+$displayState.Remaining+' seconds.'}})
                }
            }
            'Sound' {
                $audio=[SmartThinClient.DeviceSettings]::GetAudio();[void](& $addText 'Default playback device volume');$volume=[Windows.Controls.Slider]::new();$volume.Minimum=0;$volume.Maximum=100;$volume.Value=[Math]::Round($audio.Volume*100);$volume.TickFrequency=1;$volume.IsSnapToTickEnabled=$true;$volume.Margin='0,12,0,16';[void]$body.Children.Add($volume)
                $mute=[Windows.Controls.CheckBox]::new();$mute.Content='Mute';$mute.IsChecked=$audio.Muted;$mute.Margin='0,0,0,12';[void]$body.Children.Add($mute)
                $save=& $addButton 'Apply volume';$save.Add_Click({try{[SmartThinClient.DeviceSettings]::SetAudio([single]($volume.Value/100),[bool]$mute.IsChecked);$status.Text='Playback volume updated.'}catch{& $fail $_}})
            }
            'DateAndTime' {
                $nowText=[DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss',[Globalization.CultureInfo]::InvariantCulture);[void](& $addText ('Current local time: '+$nowText))
                $zones=[Windows.Controls.ComboBox]::new();$zones.ItemsSource=@([TimeZoneInfo]::GetSystemTimeZones());$zones.DisplayMemberPath='DisplayName';$zones.SelectedValuePath='Id';$zones.SelectedValue=[TimeZoneInfo]::Local.Id;$zones.Margin='0,0,0,12';[void]$body.Children.Add($zones)
                $zoneApply=& $addButton 'Apply time zone';$zoneApply.Add_Click({try{Set-TimeZone -Id ([string]$zones.SelectedValue) -ErrorAction Stop;$status.Text='Time zone updated.'}catch{& $fail $_}})
                [void](& $addText 'Set date and time (administrator required)');$clock=[Windows.Controls.TextBox]::new();$clock.Text=$nowText;$clock.Margin='0,0,0,8';[void]$body.Children.Add($clock)
                $clockApply=& $addButton 'Apply date and time';$identity=[Security.Principal.WindowsIdentity]::GetCurrent();try{$clockApply.IsEnabled=([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)}finally{$identity.Dispose()}
                $clockApply.Add_Click({try{$date=[DateTime]::ParseExact($clock.Text,'yyyy-MM-dd HH:mm:ss',[Globalization.CultureInfo]::InvariantCulture);[void](Set-Date -Date $date -ErrorAction Stop);$status.Text='Local date and time updated.'}catch{& $fail $_}})
            }
            'Bluetooth' {
                $radio=[SmartThinClient.DeviceSettings]::HasBluetoothRadio();[void](& $addText $(if($radio){'A Bluetooth radio is available. Use the Windows pairing wizard to add a device.'}else{'No Bluetooth radio was detected on this device. Bluetooth pairing is unavailable in this session.'}))
                $pair=& $addButton 'Pair a device';$pair.IsEnabled=$radio;$pair.Add_Click({try{Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\DevicePairingWizard.exe') -ErrorAction Stop;$status.Text='Device pairing wizard requested.'}catch{& $fail $_}})
            }
            'Printers' {
                $printers=@([Drawing.Printing.PrinterSettings]::InstalledPrinters|ForEach-Object {[string]$_})
                if($printers.Count){$printer=[Windows.Controls.ComboBox]::new();$printer.ItemsSource=$printers;$printer.SelectedItem=([Drawing.Printing.PrinterSettings]::new()).PrinterName;if($printer.SelectedIndex -lt 0){$printer.SelectedIndex=0};$printer.Margin='0,0,0,12';[void]$body.Children.Add($printer)
                    $properties=& $addButton 'Printer properties';$properties.Add_Click({try{$handle=[Windows.Interop.WindowInteropHelper]::new($window).Handle;[SmartThinClient.DeviceSettings]::ShowPrinterProperties([string]$printer.SelectedItem,$handle)}catch{& $fail $_}})
                    $makeDefault=& $addButton 'Set as default printer';$makeDefault.Add_Click({try{if(-not[SmartThinClient.DeviceSettings]::SetDefaultPrinter([string]$printer.SelectedItem)){throw [ComponentModel.Win32Exception]::new()};$status.Text='Default printer updated.'}catch{& $fail $_}})
                }else{[void](& $addText 'No printer is installed for this user.')}
                $status.Text='Printer installation and drivers are managed by an administrator.'
            }
        }
    }catch{$status.Text='Device settings are unavailable: '+$_.Exception.Message;$status.Foreground='#B42318'}
    try {[void]$window.ShowDialog()}finally{$displayTimer.Stop();if($displayState.Pending){[SmartThinClient.DeviceSettings]::SetDisplay($displayState.Original,$false)}}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDn6gwnCCihCWwY
# n2NdhmLgfQ+bJg+t493eO2ThxlgqNqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAly1fLZ8PdEyNKLAUN6dUk0hEX1iBcD7LWBf7GkGJOkMA0GCSqG
# SIb3DQEBAQUABIIBgKEi5dP9IqvD1VW/7DoAoT9/20GoJDk5fjYNcUJw3L5Vcvw7
# 7lix3QZS6k9JnG5JEj/8RetMHiwxo9f5K79l301u7k+4pIPL4gP02RT7sUkyOrEw
# g3z4/VPqDfJVH1pPUY2lUw1g+amKmu6hG2YkWKshbNrPYg5V4N03jkGgzrwIh1N9
# 4NC6sU0kg7Vcw3zjh6CH5Z0r8YDhgnLrb9udIrzsWNpFxpGqsuezt9PwFCWzXa/4
# X+em5PsRlRKpjCRqgsLHpv4eBntFOUrJk4ax2HP1fKVc1P/V628mir9LT4FN42Rg
# EyoHTZ20/4rqKUU8oOR+Zu8PENBtt/sgEWiJo4AKYL+Z7tZusTam7bfU9Fdm6S6C
# JHCIb07+FPWHhDpz++r5qs1e7SNjuT5lBkcjk9qfuF4I/lnvWXbV38G5R6iQn6QX
# 9btjnuRoQmB/cMJ7E+qEfteNS+1naWXtTzDjqW1w0+FV5ulre+2ZA0VhEeo3GfpP
# Te51bBZzF52SHKDM46GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxNzE0
# MTdaMC8GCSqGSIb3DQEJBDEiBCCXJRPuMx8KcczWLqeA9NYbh+GHJPBmuV9ZorQv
# rRJMjDANBgkqhkiG9w0BAQEFAASCAgCSvK4O7wrx1fQ7KF1TsaK5a1n9bxnWTZPy
# K6rg6udATdkSXgrFeeTmNLEQBmVe7ROkk+bjpsQsg3RDQH6TCuzT83TP1dmsdxmE
# uMuLbjTcNoAnwXOw2troUa+Ask9aLEoXlCch9kcXeFzuV0Qbfc8b1r4uQFwudvVW
# f/YyDSDTEnRMbCCdUOPFXHGG6mG15pxmzMeE5O5oojh73488aWEq4kyfzrvA3mX6
# DVu2eSYC7cIKNZgnck7y51BAWwX1dyMdgeTAoRQLY2sJByCqnGMleuk2zRi7oZ3y
# 7EKDFBMl0mP92YuN+efW7lVRy1Au1CseBRJbdp8QRbZXC309FdGexYLgnXZiQs1T
# 7yclgJ7+ZXiBKK0A+VZLhJRItSfRXhPusbOVGsO0qB6VV1HmrkeUbaahIvmrpBhS
# rxlyjrG8+Oaind4qK6bJkEbvkTP15JRQT1y2jF/w/z/gwHU55H5H84f4eMs8B9h1
# 7WCRH/lOHTqjJT5X2EN9+/ggGn0F/DL5NAAse8BZXMM6DCWvcnLQYcyvJLj/rVxM
# 8dWcdKoJZ2tKfN9cuAjKCNAUtpL2sKn3y06GUvEK/k0W699LZpqF1xYVEhlMuwN3
# uDWgCK8dMCs3F+Vvo/Pgt0L8iAaM7GPLVKd/vulWPvmy6SAuqEoNHrzJU+akF/D1
# xe93uHXIMw==
# SIG # End signature block
