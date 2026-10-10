[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$folder = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1'
$templatePath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.local.json.template'
$registryPath = Join-Path $folder '..\..\..\Modules\SmartM365.Core\SmartM365-SourceReceipts.json.txt'
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count) { throw "Infrastructure parser errors: $($errors.Count)" }
$versionMatch = [regex]::Match($ast.Extent.Text, '(?s)\.VERSION\s+([0-9.]+)')
if (-not $versionMatch.Success -or [version]$versionMatch.Groups[1].Value -lt [version]'1.0.11') { throw 'Infrastructure version was not updated for zone URLs and server counts.' }
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
catch { if ($_.Exception.Message -ne 'Unresolved configuration value: OutputRoot') { throw } }
$root = 'C:\ProgramData\SmartM365\LauncherCache\SharePointInfrastructure\SmartM365'
try { $null = Assert-InventoryPath -Path (Join-Path $root 'Data\Tenants\prod\DATA-ALL') -Name 'OutputRoot'; throw 'Disposable infrastructure output path was accepted.' }
catch { if ($_.Exception.Message -notlike 'OutputRoot resolves inside the disposable launcher cache.*') { throw } }
if ((Assert-InventoryPath -Path 'C:\SmartM365\DATA' -Name 'OutputRoot') -ne 'C:\SmartM365\DATA') { throw 'Persistent infrastructure output path was rejected.' }
$script:EffectiveConfig = [pscustomobject]@{ IncludedWebApplicationUrls=@('https://content-a'); ExcludedWebApplicationUrls=@('https://content-b') }
function Get-SPWebApplication { @([pscustomobject]@{Url='https://content-a/'},[pscustomobject]@{Url='https://content-b/'}) }
$selected = @(Get-ConfiguredWebApplications)
if ($selected.Count -ne 1 -or $selected[0].Url -ne 'https://content-a/') { throw 'Web application filtering failed.' }
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
Write-Output 'PASS: infrastructure parser, PS5 CSV, receipts, daily mail gate/routing, edition mapping, database status, alerts, safe delta and HTML.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBBAyyPsyksXA2z
# i4LYLYsGZOZLRj4axHJMS0dduzLg1aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIArjxw3ajCoJDgadUe7TU71eQkIK3Wdi5IZVwK0zZ5P3MA0GCSqG
# SIb3DQEBAQUABIIBgIQW5ym4gslhGZq+ArHUuQP3vAEm7FmG3mFK97AAmyyyIqYa
# QRVwbFFV8Emu8odY07+jlRWt6C4HdG5OV/mdGVKfZLJfacSVA7REAnyiKXokpZYu
# lNV8eTrRObUCziNnQHnUI0f+Gl9ZLRpsQggZZTczdykYzGDPVJWAA9WehHdhTGqY
# 8D5OSQsoLNQoliIaMkGqj0Pu84vZpdYQ7Db6MZyYqhSvy8Hy2thnYSOzCvFuu3/X
# QQhMGicixlmIkCokreuCeJH9zUPmhfO3JgnLVMPJAR29D3Z8Ow0iYHSghn2IPb15
# Yi6KaO2n3Bb77+fvmLqxnZUVj8HasOjYw9kws/d8EFZxOqlJ+1mvW4HOt0kRtCZs
# dGkLZhuGYbGGKRPYpJmexCdM1hrlzb6bJeZ7GaDi02DT5/+Ai9/BNtSOkcAEEERU
# AOK0y2xvyEKNdpzTVYx73hS+d7oV+caaBe7w6l9VI/NVkuDRJWywoTtetDRi7DIv
# bjjxOecPG+HK/uNdTqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTAwNzI2
# MDhaMC8GCSqGSIb3DQEJBDEiBCBt8mUVRdmWIWifbaiUViQ4ULMhilzzW+ls909j
# 3/M0qTANBgkqhkiG9w0BAQEFAASCAgBBpbV0SP77LcVSHlCHZN44/ckfBuHd/Cem
# D22eM26XhgdIdbeoBELlFnrZ36fldRn44sRI+T3V7uSNEozX07Vohkx64vIwNIA6
# IayYYF9Bv7PSG09XnYWc673yXJxSfFFyUzAJjrdRfFDGXazEX/PQdPoqQbqm/unL
# doEg6vs5vrDjqJXlco25rX95PeU57yIvCLoi2bBq0zGjHcHCNP8A0fb0UtqkrqDA
# x/Rk6SGKjzPLLtmlBoY0oojcKU7skW5HvNNlgqHuj1SEJkNL4GGaJ4xmQajH+BIm
# wdOtt04S5de4sRVB3Ts3EIMxm+n88MAe/SdOgNawal0cfwWUmSQdEoubntGVroY2
# RYjw5WN4qRkwoggIoy7l84TmtMqf4i2wt5RMN767Qy43haoNnlZ9UEXej1CzU6tN
# +GVecghT0DnxyKOaHaB5eS6N4Wovg7uCJOI7/ViR7qRsNYNi3v7Yxe/fhbU7NmZP
# wRMKTc/gU7nRHH5cU7QULUweuMNbvvPiN+SHh+Wxlm7lf+tTgtJ4g8+4pXesScW8
# MUNbyFydWfGygS2B9vJUczAaCKg4GFH7VPqQNia/9qu4Pu2mh8reVtfvbLod/UBB
# VBC3X51GUNWEfrKaCYGTj5Y1My1TXlpMCIkf87nS2N52978ranW0jaZdr4FbXILQ
# Z1CaT/ujqw==
# SIG # End signature block
