<#
.SYNOPSIS
    Offline SharePoint on-premises infrastructure inventory tests.
.VERSION
    1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$global:EnableTeamsNotifications=$false
$global:SmartM365GlobalConfig=[pscustomobject]@{ProfileKey='test';EnableTeamsNotifications=$false}
function global:Invoke-OfflineSuppressedTeamsNotification { [CmdletBinding()]param([Parameter(ValueFromRemainingArguments=$true)]$Remaining) }
Set-Alias -Name Send-SmartM365TeamsNotification -Value Invoke-OfflineSuppressedTeamsNotification -Scope Global -Option AllScope
$folder = Split-Path $PSScriptRoot -Parent
. (Join-Path $folder '..\..\..\Config\SmartM365-TenantContext.ps1')
$scriptPath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1'
$templatePath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.local.json.template'
$registryPath = Join-Path $folder '..\..\..\Modules\SmartM365.Core\SmartM365-SourceReceipts.json.txt'
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count) { throw "Infrastructure parser errors: $($errors.Count)" }
$versionMatch = [regex]::Match($ast.Extent.Text, '(?s)\.VERSION\s+([0-9.]+)')
if (-not $versionMatch.Success -or [version]$versionMatch.Groups[1].Value -lt [version]'1.0.12') { throw 'Infrastructure version was not updated for the combined mail policy.' }
if (@($ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Send-InfrastructureRunMail' }, $true)).Count -ne 0) { throw 'Infrastructure still sends a separate summary email.' }
if ($ast.Extent.Text -notmatch 'ForceMail is retained for command compatibility') { throw 'Legacy ForceMail behavior is not explained.' }
foreach ($column in @('DatabaseSizeBytes','DatabaseSizeStatus','NeedsUpgradeIncludeChildren','NeedsUpgradeStatus')) {
    if ($ast.Extent.Text -notmatch [regex]::Escape("'$column'")) { throw "Content database schema lacks $column" }
}
if ($ast.Extent.Text -notmatch "\[Alias\('ForceSendDailySummary'\)\]\[switch\]\`$ForceMail") { throw 'ForceMail alias is missing.' }
$firstUploadConfigRead = $ast.Extent.Text.IndexOf('$global:SharePointSiteHostname =', [StringComparison]::Ordinal)
$resolverDefinition = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Resolve-InventoryConfigTokens' }, $true))
if ($firstUploadConfigRead -lt 0 -or $resolverDefinition.Count -ne 1 -or $resolverDefinition[0].Extent.EndOffset -ge $firstUploadConfigRead) { throw 'Infrastructure reads upload configuration before the token resolver is defined.' }
$template = Get-Content $templatePath -Raw | ConvertFrom-Json
if ($template.EnableSharePointUpload -ne $true -or $template.EnableWeeklyHistory -ne $true -or $template.SendMailMode -ne 'Graph') { throw 'Infrastructure template policy is invalid.' }
$registry = Get-Content $registryPath -Raw | ConvertFrom-Json
$producer = @($registry.Producers | Where-Object Script -eq 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1')
if ($producer.Count -ne 1 -or $producer[0].Files.Count -ne 6) { throw 'Infrastructure source receipt registration is invalid.' }
if ($ast.Extent.Text -match '(?im)^\s*(Set-SPSite|Set-SPWeb|Set-SPContentDatabase|Add-SPShellAdmin|Remove-SPSite)\b') { throw 'A SharePoint write command was found.' }
if ($ast.Extent.Text -notmatch '-Rows\s+\$rows\[\$kind\]\.ToArray\(\)') { throw 'Infrastructure CSV buffers must be converted to arrays for Windows PowerShell 5.1.' }
foreach ($name in @('Get-InventoryConfigValue','Resolve-InventoryConfigTokens','Assert-InventoryPath','Get-ConfiguredWebApplications','Get-ObservedProperty','Get-SharePointServerCount','Get-WebApplicationZonePublicUrl','Get-FarmConfigurationDatabaseName','Get-ContentDatabaseExtendedFields','Get-InfrastructureFarmEdition','Get-QualifiedPreviousInfrastructureCounts','New-InfrastructureMailHtml','Send-InfrastructureRunMail','Write-RunCsv','Ensure-GraphAuthenticationModule','Invoke-DailySummaryMail','Send-InventorySummaryMail','Publish-QualifiedCsvUploads')) {
    $functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true))
    if ($functionAst.Count -ne 1) { throw "Missing function: $name" }
    . ([scriptblock]::Create($functionAst[0].Extent.Text))
}
$script:EffectiveConfig = [pscustomobject]@{
    SmartM365RootPath='C:\SmartM365'; WorkspaceRootPath='{{SmartM365RootPath}}'; ProfileKey='prod'
    DataAllRootPath='{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-ALL'
    OutputRoot='{{DataAllRootPath}}\SharePoint\OnPrem\Infrastructure'
    LatestCsvFolderPath='{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-LAST'
}
if ((Get-InventoryConfigValue 'OutputRoot') -ne 'C:\SmartM365\Data\Tenants\prod\DATA-ALL\SharePoint\OnPrem\Infrastructure') { throw 'Nested infrastructure OutputRoot tokens were not resolved.' }
if ((Get-InventoryConfigValue 'LatestCsvFolderPath') -ne 'C:\SmartM365\Data\Tenants\prod\DATA-LAST') { throw 'Nested infrastructure DATA-LAST tokens were not resolved.' }
$script:EffectiveConfig.OutputRoot = '{{MissingPath}}\SharePoint'
try { $null = Get-InventoryConfigValue 'OutputRoot'; throw 'Missing token was accepted.' }
catch { if ($_.Exception.Message -ne 'Unresolved configuration token in: OutputRoot') { throw } }
$root = 'C:\ProgramData\SmartM365\LauncherCache\SharePointInfrastructure\SmartM365'
try { $null = Assert-InventoryPath -Path (Join-Path $root 'Data\Tenants\prod\DATA-ALL') -Name 'OutputRoot'; throw 'Disposable infrastructure output path was accepted.' }
catch { if ($_.Exception.Message -notlike 'OutputRoot resolves inside the disposable launcher cache.*') { throw } }
if ((Assert-InventoryPath -Path 'C:\SmartM365\DATA' -Name 'OutputRoot') -ne 'C:\SmartM365\DATA') { throw 'Persistent infrastructure output path was rejected.' }
$script:EffectiveConfig = [pscustomobject]@{ IncludedWebApplicationUrls=@('https://content-a'); ExcludedWebApplicationUrls=@('https://content-b') }
function Get-SPWebApplication { @([pscustomobject]@{Url='https://content-a/'},[pscustomobject]@{Url='https://content-b/'}) }
$selected = @(Get-ConfiguredWebApplications)
if ($selected.Count -ne 1 -or $selected[0].Url -ne 'https://content-a/') { throw 'Web application filtering failed.' }
$script:EffectiveConfig = [pscustomobject]@{ IncludedWebApplicationUrls=@(); ExcludedWebApplicationUrls=@() }
if (@(Get-ConfiguredWebApplications).Count -ne 2) { throw 'Empty filters must retain both web applications.' }
$script:EffectiveConfig = [pscustomobject]@{ IncludedWebApplicationUrls=@('https://content-a','https://content-b'); ExcludedWebApplicationUrls=@() }
if (@(Get-ConfiguredWebApplications).Count -ne 2) { throw 'Multiple included URLs must be matched individually.' }
$coreManifest = Join-Path $folder '..\..\..\Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
Import-Module $coreManifest -Force -ErrorAction Stop
$script:BuildEditionMapPath = Join-Path $folder 'SharePoint-OnPrem-BuildEditions.psd1'
$global:SmartM365TenantKey = 'OFFLINE'
$global:SmartM365OrganizationKey = 'OFFLINE'
$global:SmartM365EnvironmentKey = 'TEST'
$global:SmartM365TenantId = '00000000-0000-0000-0000-000000000001'
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SPInfra-Test-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $path = Write-RunCsv -Name 'Empty.csv' -Rows @() -Columns @('TenantKey','FarmId','CollectionStatus') -Folder $tempRoot
    if (-not $global:csvGeneratedPaths.Contains($path)) { throw 'Infrastructure run CSV was not registered for execution summary.' }
    $text = [IO.File]::ReadAllText($path)
    if ($text -notmatch 'TenantKey.*;.*FarmId.*;.*CollectionStatus' -or @(Import-Csv $path -Delimiter ';').Count -ne 0) { throw 'Empty infrastructure CSV contract failed.' }
    $buffers = @{ Farms = (New-Object 'System.Collections.Generic.List[object]') }
    $buffers.Farms.Add([pscustomobject]@{ TenantKey='OFFLINE'; FarmId='FARM'; CollectionStatus='Collected' })
    $bufferedPath = Write-RunCsv -Name 'Buffered.csv' -Rows $buffers['Farms'].ToArray() -Columns @('TenantKey','FarmId','CollectionStatus') -Folder $tempRoot
    $bufferedRows = @(Import-Csv -LiteralPath $bufferedPath -Delimiter ';')
    if ($bufferedRows.Count -ne 1 -or $bufferedRows[0].FarmId -ne 'FARM') { throw 'Windows PowerShell 5.1 infrastructure buffer serialization failed.' }
    $latest = Join-Path $tempRoot 'DATA-LAST'
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest
    foreach ($name in $producer[0].Files) {
        Write-SmartM365CsvAtomically -Data @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM'}) -Path (Join-Path $latest $name) -Columns @('TenantKey','FarmId') -Delimiter ';' -NoTenantKey
    }
    $proofPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    $proof = Get-Content $proofPath -Raw | ConvertFrom-Json
    if ($proof.Status -ne 'Completed' -or $proof.Files.Count -ne 6) { throw 'Infrastructure receipt completion failed.' }
    $proofHash = (Get-FileHash $proofPath -Algorithm SHA256).Hash
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest
    $null = Complete-SmartM365SourceReceipt -Status Failed -ErrorCount 1
    if ((Get-FileHash $proofPath -Algorithm SHA256).Hash -ne $proofHash) { throw 'A failed run changed the preceding infrastructure receipt.' }
    $marker = Join-Path $tempRoot 'Infrastructure-DailySummary.sent'
    $script:sendCount = 0
    function WriteLog { param($Message, $Level) }
    $farmServers = @(
        [pscustomobject]@{Role='WebFrontEnd'},
        [pscustomobject]@{Role='ApplicationWithSearch'},
        [pscustomobject]@{Role='Invalid'},
        [pscustomobject]@{Role='Invalid'}
    )
    if ((Get-SharePointServerCount -Servers $farmServers) -ne 2) { throw 'Farm server count includes external entries.' }
    $zoneWebApplication = [pscustomobject]@{Id='WEB-1'}
    $zoneWebApplication | Add-Member -MemberType ScriptMethod -Name GetResponseUri -Value { param($zone) if ($zone -eq 'Default') { return [uri]'https://public.example.test/' }; return $null }
    if ((Get-WebApplicationZonePublicUrl -WebApplication $zoneWebApplication -Zone 'Default') -ne 'https://public.example.test/') { throw 'Zone public URL was not read from the response URI.' }
    if ((Get-WebApplicationZonePublicUrl -WebApplication $zoneWebApplication -Zone 'Intranet') -ne '') { throw 'A missing zone public URL was not left empty.' }
    $invalidZoneWebApplication = [pscustomobject]@{Id='WEB-2'}
    $invalidZoneWebApplication | Add-Member -MemberType ScriptMethod -Name GetResponseUri -Value { param($zone) return 'Microsoft.SharePoint.Administration.SPAlternateUrl' }
    if ((Get-WebApplicationZonePublicUrl -WebApplication $invalidZoneWebApplication -Zone 'Default') -ne '') { throw 'An object type name was accepted as a zone public URL.' }
    $script:databaseProbeCount = 0
    $script:mockDatabases = @([pscustomobject]@{Type='Content Database';Name='Content_A'},[pscustomobject]@{Type='Configuration Database';Name='Config_A'})
    function Get-SPDatabase { [CmdletBinding()] param() $script:databaseProbeCount++; $script:mockDatabases }
    $directFarm = [pscustomobject]@{ConfigurationDatabase=[pscustomobject]@{Name='Direct_Config'}}
    if ((Get-FarmConfigurationDatabaseName -Farm $directFarm) -ne 'Direct_Config' -or $script:databaseProbeCount -ne 0) { throw 'Direct configuration database name was not used.' }
    if ((Get-FarmConfigurationDatabaseName -Farm ([pscustomobject]@{})) -ne 'Config_A' -or $script:databaseProbeCount -ne 1) { throw 'Configuration database fallback failed.' }
    $script:mockDatabases = @([pscustomobject]@{Type='Content Database';Name='Content_A'})
    if ((Get-FarmConfigurationDatabaseName -Farm ([pscustomobject]@{})) -ne '') { throw 'A missing configuration database was accepted.' }
    $script:mockDatabases = @([pscustomobject]@{Type='Configuration Database';Name='Config_A'},[pscustomobject]@{Type='Configuration Database';Name='Config_B'})
    if ((Get-FarmConfigurationDatabaseName -Farm ([pscustomobject]@{})) -ne '') { throw 'Ambiguous configuration databases were accepted.' }
    Remove-Item Function:\Get-SPDatabase
    $sendAction = { $script:sendCount++ }
    if (-not (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction)) { throw 'First infrastructure email was skipped.' }
    if (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction) { throw 'Same-day infrastructure email was repeated.' }
    if (-not (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction -Force) -or $script:sendCount -ne 2) { throw 'Forced infrastructure email failed.' }
    $failedMarker = Join-Path $tempRoot 'Infrastructure-Failed.sent'
    try { $null = Invoke-DailySummaryMail -MarkerPath $failedMarker -SendAction { throw 'Simulated SMTP failure' }; throw 'Mail failure was ignored.' }
    catch { if ($_.Exception.Message -ne 'Simulated SMTP failure') { throw } }
    if (Test-Path -LiteralPath $failedMarker) { throw 'Failed infrastructure email wrote a marker.' }
    $script:mailParameters = $null
    $script:mailSendCount = 0
    function SendEmailHtmlReport { [CmdletBinding()] param($From,$To,$Subject,$BodyHtml,$MailPurpose,$SmtpServer,$SmtpPort,$SendMailMode,$Cc) $script:mailParameters = $PSBoundParameters; $script:mailSendCount++ }
    $script:EffectiveConfig = [pscustomobject]@{ From='sender@example.test'; To=''; ErrorMailTo='recipient@example.test'; SmtpServer='smtp.example.test'; SmtpPort=25; SendMailMode='SMTP'; Cc='' }
    Send-InventorySummaryMail -Subject 'Offline summary' -BodyHtml '<p>Summary</p>'
    if ($script:mailParameters.To -ne 'recipient@example.test' -or $script:mailParameters.From -ne 'sender@example.test' -or $script:mailParameters.SendMailMode -ne 'Graph' -or $script:mailParameters.ContainsKey('Attachments') -or $script:mailParameters.ContainsKey('SmtpServer')) { throw 'Infrastructure Graph mail routing or attachment policy failed.' }
    $script:graphAvailable = $false
    $script:graphInstallerAvailable = $true
    $script:graphInstallCount = 0
    $script:graphImportCount = 0
    function Get-Module { [CmdletBinding()] param([switch]$ListAvailable,[string]$Name) if ($script:graphAvailable -and $Name -eq 'Microsoft.Graph.Authentication') { [pscustomobject]@{Name=$Name;Version=[version]'2.0.0'} } }
    function Get-Command { [CmdletBinding()] param([string]$Name) if (($Name -eq 'Install-Module' -and $script:graphInstallerAvailable) -or ($Name -in @('Connect-MgGraph','Get-MgContext','Invoke-MgGraphRequest') -and $script:graphAvailable)) { [pscustomobject]@{Name=$Name} } }
    function Install-Module { [CmdletBinding()] param([string]$Name,[string]$Scope,[string]$Repository,[switch]$Force,[switch]$AllowClobber) if ($Name -ne 'Microsoft.Graph.Authentication' -or $Scope -ne 'CurrentUser' -or $Repository -ne 'PSGallery' -or -not $Force -or -not $AllowClobber) { throw 'Unexpected Graph installation parameters.' }; $script:graphInstallCount++; $script:graphAvailable = $true }
    function Import-Module { [CmdletBinding()] param([string]$Name) if ($Name -ne 'Microsoft.Graph.Authentication') { throw 'Unexpected Graph import.' }; $script:graphImportCount++ }
    $originalProtocol = [Net.ServicePointManager]::SecurityProtocol
    Ensure-GraphAuthenticationModule
    Ensure-GraphAuthenticationModule
    if ($script:graphInstallCount -ne 1 -or $script:graphImportCount -ne 2 -or [Net.ServicePointManager]::SecurityProtocol -ne $originalProtocol) { throw 'Infrastructure Graph bootstrap was not idempotent or changed the process TLS policy.' }
    $script:graphAvailable = $false
    $script:graphInstallerAvailable = $false
    try { Ensure-GraphAuthenticationModule; throw 'Missing Graph installer was accepted.' }
    catch { if ($_.Exception.Message -ne 'Install-Module is unavailable.') { throw } }
    Remove-Item Function:\Get-Module,Function:\Get-Command,Function:\Install-Module,Function:\Import-Module
    $script:uploadCalls = New-Object 'System.Collections.Generic.List[string]'
    function Invoke-SmartM365SharePointCsvUpload { param($LocalFilePath) $script:uploadCalls.Add($LocalFilePath); if ($LocalFilePath -eq 'failed.csv') { return }; return [pscustomobject]@{ LocalFilePath=$LocalFilePath } }
    $global:EnableSharePointUpload = $false
    if ((Publish-QualifiedCsvUploads -Paths @('run.csv','latest.csv')) -ne 0 -or $script:uploadCalls.Count -ne 0) { throw 'Disabled infrastructure upload was attempted.' }
    $global:EnableSharePointUpload = $true
    if ((Publish-QualifiedCsvUploads -Paths @('run.csv','latest.csv','run.csv','failed.csv')) -ne 1 -or $script:uploadCalls.Count -ne 3) { throw 'Infrastructure upload qualification failed.' }
    . (Join-Path $folder '..\..\..\Modules\SmartM365.Core\SmartM365-SharePointUploadPolicy.ps1')
    $outputBase=Join-Path $tempRoot 'policy-all';$latestRoot=Join-Path $tempRoot 'policy-last';$global:LogPath=Join-Path $tempRoot 'policy-log'
    $global:SmartM365CollectorUploadPolicies=@{DataAll=[pscustomobject]@{Enabled=$false};DataLast=[pscustomobject]@{Enabled=$true};LogAll=[pscustomobject]@{Enabled=$true}}
    $script:uploadCalls.Clear();$runPath=Join-Path $outputBase 'run.csv';$lastPath=Join-Path $latestRoot 'latest.csv'
    if ((Publish-QualifiedCsvUploads -Paths @($runPath,$lastPath)) -ne 0 -or $script:uploadCalls.Count -ne 1 -or $script:uploadCalls[0] -ne $lastPath) { throw 'Infrastructure upload category selection failed.' }
    Remove-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue

    $extended = Get-ContentDatabaseExtendedFields -Database ([pscustomobject]@{NeedsUpgradeIncludeChildren=$true;DiskSizeRequired=999999})
    if ($extended.DatabaseSizeBytes -ne '' -or $extended.DatabaseSizeStatus -ne 'Unavailable' -or $extended.NeedsUpgradeIncludeChildren -ne 'True' -or $extended.NeedsUpgradeStatus -ne 'Collected') { throw 'Read-only content database extended fields are incorrect.' }
    $missingExtended = Get-ContentDatabaseExtendedFields -Database ([pscustomobject]@{})
    if ($missingExtended.NeedsUpgradeIncludeChildren -ne '' -or $missingExtended.NeedsUpgradeStatus -ne 'PropertyUnavailable') { throw 'Missing upgrade property was treated as false.' }
    $falseExtended = Get-ContentDatabaseExtendedFields -Database ([pscustomobject]@{NeedsUpgradeIncludeChildren=$false})
    if ($falseExtended.NeedsUpgradeIncludeChildren -ne 'False' -or $falseExtended.NeedsUpgradeStatus -ne 'Collected') { throw 'Observed false upgrade state was lost.' }
    if ((Get-InfrastructureFarmEdition '16.0.10417.20198') -ne '2019' -or (Get-InfrastructureFarmEdition '16.0.5565.1001') -ne '2016' -or (Get-InfrastructureFarmEdition '16.0.99999.0') -ne 'Unknown') { throw 'Farm edition mapping failed.' }

    $reportRows = @{
        Farms=@([pscustomobject]@{BuildVersion='16.0.10417.20198'})
        Servers=@(
            [pscustomobject]@{ServerName='S1';Role='Application';Status='Offline'},
            [pscustomobject]@{ServerName='SQL1';Role='Invalid';Status='Offline'},
            [pscustomobject]@{ServerName='SMTP1';Role='Invalid';Status='Online'}
        )
        ServiceApplications=@([pscustomobject]@{Name='Search';Status='Stopped'})
        WebApplications=@([pscustomobject]@{WebApplicationId='W1';DefaultUrl='https://example.test';ApplicationPool='Pool';ContentDatabaseCount=1})
        WebApplicationZones=@([pscustomobject]@{WebApplicationId='W1';Zone='Default';AuthenticationMode='Forms'})
        ContentDatabases=@([pscustomobject]@{Name='DB1';DatabaseServer='SQL1';Status='Offline';IsReadOnly='True';NeedsUpgradeIncludeChildren='True';NeedsUpgradeStatus='Collected';DatabaseSizeBytes='';DatabaseSizeStatus='Unavailable'})
    }
    $reportHtml = New-InfrastructureMailHtml -Status Qualified -Rows $reportRows -UploadFailures 2 -PreviousCounts @{Servers=0;ServiceApplications=1;WebApplications=1;ContentDatabases=1} -TenantName 'offline' -Duration ([timespan]::FromSeconds(31))
    foreach ($expected in @('SMARTM365 SHAREPOINT ONPREM','QUALIFIED','Duration: 00:00:31','Servers','1 (+1)','2 external','Alerts','Server not online','Service application not started','Content database read-only','Content database needs upgrade','Authentication by zone','Default: Forms','2019','16.0.10417.20198','SharePoint_OnPrem_Servers.csv')) {
        if ($reportHtml -notmatch [regex]::Escape($expected)) { throw "Infrastructure mail lacks $expected" }
    }
    if ($reportHtml -match 'Farm build</div>|Files / rows</div>|>Invalid</td>') { throw 'Redundant card or raw invalid server role is still present.' }
    if ($reportHtml -notmatch '>Files<' -or $reportHtml -notmatch '(?s)SharePoint_OnPrem_Servers\.csv</td>\s*<td[^>]*>3</td>') { throw 'CSV file and raw row counts are absent from Files.' }
    if ($reportHtml -match '>SQL1</td>|>SMTP1</td>|>External</td>') { throw 'External entries must not appear in the Servers email table.' }
    if ($reportHtml -notmatch '<td nowrap="nowrap"[^>]*>https://example\.test</td>' -or $reportHtml -notmatch '<td nowrap="nowrap"[^>]*>Pool</td>') { throw 'Web application URL or pool may break within a word.' }
    if ($reportHtml -notmatch '(?s)>Farm edition</div>\s*<div[^>]*>2019</div>\s*<div[^>]*>16\.0\.10417\.20198</div>') { throw 'Farm build is not the small detail below the edition.' }
    if ($reportHtml -match 'Top 5 largest content databases') { throw 'Top-five size ranking was shown without measured sizes.' }
    $reportRows.ContentDatabases[0].DatabaseSizeBytes = '1024'
    $reportRows.ContentDatabases[0].DatabaseSizeStatus = 'Collected'
    $sizedHtml = New-InfrastructureMailHtml -Status Qualified -Rows $reportRows -TenantName 'offline'
    if ($sizedHtml -notmatch 'Top 5 largest content databases' -or $sizedHtml -notmatch '1024') { throw 'Measured database size was not ranked.' }
    $reportRows.ContentDatabases[0].DatabaseSizeBytes = ''
    $reportRows.ContentDatabases[0].DatabaseSizeStatus = 'Unavailable'
    $unknownHtml = New-InfrastructureMailHtml -Status Failed -Rows $null -TenantName 'offline'
    if ($unknownHtml -notmatch 'FAILED' -or $unknownHtml -notmatch 'n/a' -or $unknownHtml -match '>Files<') { throw 'Failed-run mail misrepresented unavailable CSV files.' }
    $reportRows.Servers = @()
    $partialHtml = New-InfrastructureMailHtml -Status Partial -Rows $reportRows -Failures @('Servers: Access is denied') -TenantName 'offline'
    if ($partialHtml -notmatch 'PARTIAL' -or $partialHtml -notmatch 'Access is denied' -or $partialHtml -notmatch 'n/a' -or $partialHtml -match '>Files<') { throw 'Partial-run mail misrepresented unavailable CSV files.' }
    $reportRows.Servers = @([pscustomobject]@{ServerName='S1';Role='Application';Status='Offline'})
    $serverBuffer = New-Object 'System.Collections.Generic.List[object]'
    $serverBuffer.Add([pscustomobject]@{ServerName='S1';Role='Application';Status='Online'})
    $serverBuffer.Add([pscustomobject]@{ServerName='S2';Role='WebFrontEnd';Status='Online'})
    $reportRows.Servers = $serverBuffer
    $bufferedHtml = New-InfrastructureMailHtml -Status Qualified -Rows $reportRows -TenantName 'offline'
    if ($bufferedHtml -notmatch '(?s)>Servers</div>\s*<div[^>]*>2</div>') { throw 'Generic-list run buffers were not counted correctly in email.' }
    $reportRows.Servers = @([pscustomobject]@{ServerName='S1';Role='Application';Status='Offline'})
    $outputBase = Join-Path $tempRoot 'mail-gates'
    $script:StartedAt = Get-Date
    $Tenant = 'offline'
    $ForceMail = $false
    $beforeMail = $script:mailSendCount
    Send-InfrastructureRunMail -Status Partial -Rows $reportRows
    Send-InfrastructureRunMail -Status Failed -Rows $reportRows
    Send-InfrastructureRunMail -Status Qualified -Rows $reportRows
    Send-InfrastructureRunMail -Status Qualified -Rows $reportRows
    if ($script:mailSendCount -ne ($beforeMail + 2)) { throw 'Qualified and issue daily mail counters were not independent.' }
    $ForceMail = $true
    Send-InfrastructureRunMail -Status Qualified -Rows $reportRows
    if ($script:mailSendCount -ne ($beforeMail + 3)) { throw 'ForceMail did not bypass the qualified daily counter.' }
    $ForceMail = $false
    function SendEmailHtmlReport { throw 'Simulated Graph mail failure' }
    $ForceMail = $true
    Send-InfrastructureRunMail -Status Failed -Rows $reportRows
    $ForceMail = $false
    if (Test-Path -LiteralPath (Join-Path $outputBase 'SmartM365-SharePoint-OnPrem-Infrastructure-IssueDailySummary.sent') -PathType Leaf) {
        # The prior successful issue mail owns this marker; failed forced sends must not change it.
        $markerDate = (Get-Content -LiteralPath (Join-Path $outputBase 'SmartM365-SharePoint-OnPrem-Infrastructure-IssueDailySummary.sent') -Raw).Trim()
        if ($markerDate -ne (Get-Date).ToString('yyyy-MM-dd')) { throw 'Issue mail marker changed unexpectedly.' }
    }

    $comparisonRoot = Join-Path $tempRoot 'comparison'
    New-Item -ItemType Directory -Path $comparisonRoot -Force | Out-Null
    $comparisonFiles = @()
    foreach ($kind in @('Farms','Servers','ServiceApplications','WebApplications','WebApplicationZones','ContentDatabases')) {
        $fileName = "SharePoint_OnPrem_$kind.csv"
        $filePath = Join-Path $comparisonRoot $fileName
        $comparisonData = if ($kind -eq 'Servers') {
            @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM';RunId='PRIOR';Role='WebFrontEnd'},[pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM';RunId='PRIOR';Role='Invalid'})
        } else { @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM';RunId='PRIOR'}) }
        $comparisonColumns = if ($kind -eq 'Servers') { @('TenantKey','FarmId','RunId','Role') } else { @('TenantKey','FarmId','RunId') }
        Write-SmartM365CsvAtomically -Data $comparisonData -Path $filePath -Columns $comparisonColumns -Delimiter ';' -NoTenantKey
        $comparisonFiles += [pscustomobject]@{File=$fileName;RunId='RECEIPT';Status='Success';Rows=@($comparisonData).Count;SHA256=(Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash}
    }
    $comparisonReceipt = [pscustomobject]@{Owner='SmartInventory-SourceReceipt';Status='Completed';RunId='RECEIPT';Files=$comparisonFiles}
    $comparisonReceipt | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $comparisonRoot 'SmartInventory_SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.current.json.txt')
    $previous = Get-QualifiedPreviousInfrastructureCounts -LatestRoot $comparisonRoot -FarmId 'FARM' -TenantKey 'OFFLINE'
    if ($null -eq $previous -or $previous.ContentDatabases -ne 1 -or $previous.Servers -ne 1) { throw 'Verified prior-run delta input or external server count was rejected.' }
    if ($null -ne (Get-QualifiedPreviousInfrastructureCounts -LatestRoot $comparisonRoot -FarmId 'OTHER' -TenantKey 'OFFLINE')) { throw 'Delta accepted a different farm.' }
    Add-Content -LiteralPath (Join-Path $comparisonRoot 'SharePoint_OnPrem_ContentDatabases.csv') -Value 'tampered'
    if ($null -ne (Get-QualifiedPreviousInfrastructureCounts -LatestRoot $comparisonRoot -FarmId 'FARM' -TenantKey 'OFFLINE')) { throw 'Delta accepted a tampered CSV.' }
    $databasePath = Join-Path $comparisonRoot 'SharePoint_OnPrem_ContentDatabases.csv'
    Write-SmartM365CsvAtomically -Data @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM';RunId='OTHER'}) -Path $databasePath -Columns @('TenantKey','FarmId','RunId') -Delimiter ';' -NoTenantKey
    $comparisonFiles[-1].SHA256 = (Get-FileHash -LiteralPath $databasePath -Algorithm SHA256).Hash
    $comparisonReceipt | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $comparisonRoot 'SmartInventory_SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.current.json.txt')
    if ($null -ne (Get-QualifiedPreviousInfrastructureCounts -LatestRoot $comparisonRoot -FarmId 'FARM' -TenantKey 'OFFLINE')) { throw 'Delta accepted inconsistent CSV run IDs.' }
} finally {
    if ([IO.Path]::GetFullPath($tempRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Output 'PASS: infrastructure parser, PS5 CSV, receipts, disabled operational mail, edition mapping, database status, and legacy HTML helpers.'

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDD9p+P7t5PiKj9
# WSM7bpMfqRldq9yHGK0k/T7HMCJNKKCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgP4AS7co0ZjLa
# ofgyQRNERWOdcW5xFDUsAhT5j/DdLKgwDQYJKoZIhvcNAQEBBQAEggGAIOSE7wNS
# lt2SH+YXQel8z1alDrAWN/23ATiFDcWoDBL5evlBirqlL04RQL0wsiZWQ02m7yls
# PWMCf9UQLzl+CrWzqOrQ8rLJqsax+uDirjbvw6oSVLN5/jARRNeLzCooIZ1VLs24
# hcpkM2MrbsWU4nzhhDhgT4MvDN5yvuDYWF1AcJsTPDRteebcE9LSJ3Ny3/5l7YZY
# Ivbx+YRZVivHwFSRJY8vCJEfpstKaX2lNTR5eucHKDYQ5NL/Xac/tE0GzezO/B/X
# wlr7c6iVc2oDV8etRlplxh9vb6Jdx19NvpRAlBUUJybmH7RKm9ON9J2JQA6LKN7f
# YEI+p22fezrM+qy2jQ3Sx7KMg5WVK9V2qN6yWdJKqi8ET9ICnxNoSMtElwxGy0or
# 7tKnUgj3MW9vqlO02CtpAghEv/SyWnYVHAdO4KCpMJSnCHp/C4AyzUsPLxtGpODj
# F3xDBXP/NpcFDwohhROvnuPmB+VPGdGarMZOyRFDAmx/PlZSYv9qXF17oYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTAxMTY1OTI5WjA/BgkqhkiG9w0BCQQxMgQwscLXpeOGf+QgrPcmwxFeqdZN
# fSe6o+bu1odmUbD7y8wpuO9BT6azYznzIdj84d8MMA0GCSqGSIb3DQEBAQUABIIC
# AEcvnmo55CsaGig7lpSDJ8vMeK8awrKzSy9irl69aQYPa39xhOjwN+GUvJuUYNqN
# 1OdpFEZ8/L9Q079XNgKYl0auNn4JSu7LY6Mx6miFplr14Q2X+yG3jIpHETEJKJcS
# ljgcHqnLReso6eZumW4exA/iOgvpPBnQanNFThexx2GgqwB5P9p4elbGikp1KLB9
# JduHUClvNlrnr1Y6BEqVPChEScQDtdRy4RylhH5JdXIbFYxksi1fdTHE9yt3hc4z
# j792qI71gpVlRxSNq4sFes7wbBV4WletlHUtpZzoMSkZzgZOz+gzcjJkn+y2BkDy
# yFmpnsurkcR7W6hKRoiSc9Bi4etX+78d3QE7Xl6dZzGVGZWGhAQ4yS7EryaXlqVv
# HzcM77QwC8ug3OrXYRCecX7+PDg5WVT5fTwU8idyO06PSefATCGaNXEXyxocMDmq
# VJuzCTmMbcvXeRuJ4+oITML/McV+mLDCuqn3Qgd5Qgt4u9TsLlN6Xykz9dXiXwm/
# bn+DhSW+e+wcfX/6AdN7+1cn0POU0rEBKyQZ8nDxfQUExor2+gyJi/FVS3kxvjdL
# V8OB5ZITDcoBc8gc3sfgqZ/W8svWg/nnQuVTwTS6hyvBcFLVgg8FIgHfr/U7Hbkw
# LyjdGhcs8MP2/dcrKHXhhrXInO0Ns5qBoT5C8G9Z0W56
# SIG # End signature block
