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
if ($ast.Extent.Text -notmatch '(?s)\.VERSION\s+1\.0\.9') { throw 'Infrastructure version was not updated for the mail correction.' }
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
foreach ($name in @('Get-InventoryConfigValue','Resolve-InventoryConfigTokens','Assert-InventoryPath','Get-ConfiguredWebApplications','Get-ObservedProperty','Get-FarmConfigurationDatabaseName','Get-ContentDatabaseExtendedFields','Get-InfrastructureFarmEdition','Get-QualifiedPreviousInfrastructureCounts','New-InfrastructureMailHtml','Send-InfrastructureRunMail','Write-RunCsv','Ensure-GraphAuthenticationModule','Invoke-DailySummaryMail','Send-InventorySummaryMail','Publish-QualifiedCsvUploads')) {
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
            [pscustomobject]@{ServerName='SQL1';Role='Invalid';Status='Online'},
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
    if ($reportHtml.IndexOf('>S1</td>') -gt $reportHtml.IndexOf('>SQL1</td>') -or $reportHtml.IndexOf('>S1</td>') -gt $reportHtml.IndexOf('>SMTP1</td>') -or $reportHtml.IndexOf('>External</td>') -lt 0) { throw 'External servers were not shown last with the External role.' }
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBXbyz7i88zvAJp
# UTWlq1wV1Tiw980Uu0NVQE4aOgSwXqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC7uu/Sg13781kT+EGXvfvE
# 72Bcn4ViTUWoqYQej0y6djANBgkqhkiG9w0BAQEFAASCAYBYdlDWnRerZy4RuqHM
# 2uWCfJJDW2U0lOnDikIUyoUvTyTcEC09ZptGRot/4JrlyVK9B70d8QBXtMtZk8A5
# nDXGvJwfWHPHaJSo9Uznxn38/va2pRq4X5u02O0UJorzHppS4bur9pDMO9J1a9ql
# TivgGtYsCHpauUCrALYwXyzMSmhrIqoFVAX6AuAbpm5wqwuB3Sc+mMcRTXJxiA8h
# tIpotb8W15vGFmB992GqOKC/ZnD/VuPV61+dRsLTgtIfwgSpkALtJyrL803dQT/V
# 1f0jCWCMdEpz/OjY+qAVdfenu3cXiFpnZo+DUb58cj/ZqXFTW9/2i+X+avfOwtMo
# E1kRH8NWlioKXmxib2qTJO0wnPKr28Yq3NZ9psfXmEzKCTv7VJA+36l7RseJxzLZ
# cobafX+9JmBKWnNLqSaV13EU8Yq03VtkO4fK2SgsVrTiK1sfYPZ+42Giyk7fdzY+
# WgpJx4VcL8UD8RNvTkNor/ssnQYJ13dDYIIbXxJu+uj1Yjs=
# SIG # End signature block
