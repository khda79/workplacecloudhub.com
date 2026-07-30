<#
.SYNOPSIS
    Publishes a SmartM365 Windows 11 Upgrade Toolkit .intunewin package to Intune.
.DESCRIPTION
    Creates a Win32 LOB app in Intune with Microsoft Graph beta, uploads the encrypted package payload, commits the content version, and configures PowerShell detection for the generated language package.
.VERSION
    1.0.17
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$IntuneWinPath,
    [string]$PackageId,
    [string]$PackageVersion,
    [ValidatePattern('^[a-z]{2}-[A-Z]{2}$')][string]$Language = 'fr-FR',
    [string]$RequirementLanguage,
    [string]$DisplayName,
    [string]$Description,
    [string]$Publisher = 'WorkplaceCloudHub',
    [string]$Developer = 'WorkplaceCloudHub',
    [string]$Owner = 'SmartM365',
    [string]$InstallCommandLine = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1',
    [string]$UninstallCommandLine = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Uninstall',
    [ValidateSet('system','user')][string]$RunAsAccount = 'system',
    [ValidateSet('suppress','allow','basedOnReturnCode','force')][string]$DeviceRestartBehavior = 'suppress',
    [int]$InstallCommandLineTimeoutMinutes = 180,
    [ValidateRange(-1, 2147483647)][int]$MinimumFreeDiskSpaceInMB = -1,
    [int]$UploadBlockSizeMB = 16,
    [int]$AzureUploadMaxRetries = 5,
    [int]$PollSeconds = 10,
    [int]$PollTimeoutMinutes = 45,
    [switch]$DisableLanguageRequirementRule,
    [string]$ExistingAppId,
    [switch]$UpdateMetadataOnly,
    [switch]$UpdateDetectionRules,
    [switch]$ForceCreateNew,
    [string]$FinalizeExistingAppId,
    [string]$FinalizeContentVersionId = '1',
    [string]$GraphBaseUri = 'https://graph.microsoft.com/beta',
    [switch]$NoConnect
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message)
    Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Assert-GraphCommand {
    if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        throw 'Microsoft.Graph PowerShell is required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser'
    }
}

function Get-XmlValue {
    param([xml]$Xml, [string[]]$Names)

    foreach ($name in $Names) {
        $node = $Xml.SelectSingleNode("//*[local-name()='$name']")
        if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) { return [string]$node.InnerText }
        $attr = $Xml.SelectSingleNode("//@*[local-name()='$name']")
        if ($attr -and -not [string]::IsNullOrWhiteSpace($attr.Value)) { return [string]$attr.Value }
    }
    return $null
}

function Read-CompanionPackageMetadata {
    param([string]$Path)
    $result = [ordered]@{
        PackageId = ''
        PackageVersion = ''
        Language = ''
        DisplayName = ''
        SetupCacheFolder = ''
        PackageMode = ''
        RequiresExistingSetupCache = $false
    }
    $packageDir = Split-Path -Parent $Path
    $manifestCandidates = @(
        (Join-Path $packageDir 'PackageManifest.json'),
        (Join-Path (Split-Path -Parent $packageDir) 'Source\PackageManifest.json')
    )

    foreach ($candidate in $manifestCandidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            try {
                $manifest = Get-Content -LiteralPath $candidate -Raw | ConvertFrom-Json
                if ($manifest.PackageId) { $result.PackageId = [string]$manifest.PackageId }
                if ($manifest.PackageVersion) { $result.PackageVersion = [string]$manifest.PackageVersion }
                if ($manifest.Language) { $result.Language = [string]$manifest.Language }
                if ($manifest.DisplayName) { $result.DisplayName = [string]$manifest.DisplayName }
                if ($manifest.SetupCacheFolder) { $result.SetupCacheFolder = [string]$manifest.SetupCacheFolder }
                if ($manifest.PackageMode) { $result.PackageMode = [string]$manifest.PackageMode }
                if ($manifest.PSObject.Properties['RequiresExistingSetupCache']) { $result.RequiresExistingSetupCache = [bool]$manifest.RequiresExistingSetupCache }
                break
            }
            catch { }
        }
    }

    $detectPath = Join-Path $packageDir 'Detect.ps1'
    if (Test-Path -LiteralPath $detectPath -PathType Leaf) {
        try {
            $detect = Get-Content -LiteralPath $detectPath -Raw
            $idMatch = [regex]::Match($detect, "(?m)^\s*\`$packageId\s*=\s*'(?<Value>[^']+)'")
            $versionMatch = [regex]::Match($detect, "(?m)^\s*\`$packageVersion\s*=\s*'(?<Value>[^']+)'")
            if ($idMatch.Success) { $result.PackageId = $idMatch.Groups['Value'].Value }
            if ($versionMatch.Success) { $result.PackageVersion = $versionMatch.Groups['Value'].Value }
        }
        catch { }
    }

    [pscustomobject]$result
}

function Get-LanguageFromPackageId {
    param([string]$Value)
    $match = [regex]::Match($Value, 'Win11-(?<Language>[a-z]{2}-[A-Z]{2})$', 'IgnoreCase')
    if ($match.Success) { return $match.Groups['Language'].Value }
    return $null
}

function Read-IntuneWinMetadata {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $detectionEntry = $zip.Entries | Where-Object { $_.FullName -match '(^|/)Detection\.xml$' } | Select-Object -First 1
        if (-not $detectionEntry) { throw 'Detection.xml was not found inside the .intunewin package.' }

        $reader = New-Object System.IO.StreamReader($detectionEntry.Open())
        try { [xml]$detectionXml = $reader.ReadToEnd() }
        finally { $reader.Dispose() }

        $contentEntry = $zip.Entries |
            Where-Object { $_.FullName -match '(^|/)Contents/.*\.(intunewin|bin)$' } |
            Sort-Object Length -Descending |
            Select-Object -First 1
        if (-not $contentEntry) { throw 'Encrypted package content was not found inside the .intunewin package.' }

        $metadata = [ordered]@{
            SetupFile = Get-XmlValue -Xml $detectionXml -Names @('SetupFile','SetupFilePath','ApplicationName','Name')
            FileName = Get-XmlValue -Xml $detectionXml -Names @('FileName','EncryptedFileName')
            UnencryptedContentSize = Get-XmlValue -Xml $detectionXml -Names @('UnencryptedContentSize','Size')
            EncryptionKey = Get-XmlValue -Xml $detectionXml -Names @('EncryptionKey')
            MacKey = Get-XmlValue -Xml $detectionXml -Names @('MacKey')
            InitializationVector = Get-XmlValue -Xml $detectionXml -Names @('InitializationVector')
            Mac = Get-XmlValue -Xml $detectionXml -Names @('Mac')
            ProfileIdentifier = Get-XmlValue -Xml $detectionXml -Names @('ProfileIdentifier')
            FileDigest = Get-XmlValue -Xml $detectionXml -Names @('FileDigest')
            FileDigestAlgorithm = Get-XmlValue -Xml $detectionXml -Names @('FileDigestAlgorithm')
            ContentEntryName = $contentEntry.FullName
            SizeEncrypted = [int64]$contentEntry.Length
        }

        if ([string]::IsNullOrWhiteSpace($metadata.FileName)) { $metadata.FileName = Split-Path -Leaf $Path }
        if ([string]::IsNullOrWhiteSpace($metadata.SetupFile)) { $metadata.SetupFile = 'Install.ps1' }
        if ([string]::IsNullOrWhiteSpace($metadata.ProfileIdentifier)) { $metadata.ProfileIdentifier = 'ProfileVersion1' }
        if ([string]::IsNullOrWhiteSpace($metadata.FileDigestAlgorithm)) { $metadata.FileDigestAlgorithm = 'SHA256' }

        foreach ($required in @('UnencryptedContentSize','EncryptionKey','MacKey','InitializationVector','Mac','FileDigest')) {
            if ([string]::IsNullOrWhiteSpace([string]$metadata[$required])) { throw "Unable to read IntuneWin encryption metadata field: $required" }
        }

        [pscustomobject]$metadata
    }
    finally {
        $zip.Dispose()
    }
}

function Invoke-GraphJson {
    param(
        [ValidateSet('GET','POST','PATCH','DELETE')][string]$Method,
        [string]$Uri,
        [object]$Body
    )

    if ($PSBoundParameters.ContainsKey('Body')) {
        $json = $Body | ConvertTo-Json -Depth 30
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $json -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
    }

    Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType PSObject -ErrorAction Stop
}

function ConvertTo-ODataStringLiteral {
    param([AllowNull()][string]$Value)
    return ([string]$Value -replace "'", "''")
}

function Get-GraphCollectionItems {
    param([string]$Uri)

    $items = New-Object System.Collections.ArrayList
    $next = $Uri
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $page = Invoke-GraphJson -Method GET -Uri $next
        foreach ($item in @($page.value)) { [void]$items.Add($item) }
        $next = ''
        if ($page.PSObject.Properties['@odata.nextLink']) { $next = [string]$page.'@odata.nextLink' }
    }

    return @($items.ToArray())
}

function Get-IntuneWin32AppDetails {
    param(
        [Parameter(Mandatory = $true)][string]$GraphBaseUri,
        [Parameter(Mandatory = $true)][object]$App
    )

    if ($null -eq $App -or -not $App.PSObject.Properties['id'] -or [string]::IsNullOrWhiteSpace([string]$App.id)) { return $App }
    try { return Invoke-GraphJson -Method GET -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$($App.id)" }
    catch { return $App }
}

function Test-IntuneAppHasCommittedContent {
    param([object]$App)

    if ($null -eq $App) { return $false }
    if (-not $App.PSObject.Properties['committedContentVersion']) { return $false }
    $value = [string]$App.committedContentVersion
    return -not [string]::IsNullOrWhiteSpace($value)
}

function Format-IntuneAppCandidate {
    param([object]$App)

    $committed = ''
    if ($App.PSObject.Properties['committedContentVersion']) { $committed = [string]$App.committedContentVersion }
    if ([string]::IsNullOrWhiteSpace($committed)) { $committed = '<none>' }
    return "{0} ({1}; committedContentVersion={2})" -f $App.displayName,$App.id,$committed
}

function Resolve-ExistingIntuneWin32App {
    param(
        [string]$GraphBaseUri,
        [string]$AppId,
        [string]$DisplayName,
        [string]$PackageId
    )

    if (-not [string]::IsNullOrWhiteSpace($AppId)) {
        $app = Invoke-GraphJson -Method GET -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$AppId"
        if (-not (Test-IntuneAppHasCommittedContent -App $app)) {
            throw "Existing Intune app '$AppId' has no committedContentVersion. Delete the incomplete app, finalize its first content version if it was uploaded, or use -ForceCreateNew to create a separate app."
        }
        return $app
    }

    if ([string]::IsNullOrWhiteSpace($DisplayName)) { return $null }

    $displayNameLiteral = ConvertTo-ODataStringLiteral -Value $DisplayName
    $filter = [uri]::EscapeDataString("displayName eq '$displayNameLiteral'")
    $select = [uri]::EscapeDataString('id,displayName,notes')
    $matches = @(Get-GraphCollectionItems -Uri "$GraphBaseUri/deviceAppManagement/mobileApps?`$filter=$filter&`$select=$select")
    if ($matches.Count -eq 0) { return $null }
    $detailedMatches = @($matches | ForEach-Object { Get-IntuneWin32AppDetails -GraphBaseUri $GraphBaseUri -App $_ })

    $packageMatches = @($detailedMatches | Where-Object { [string]$_.notes -match [regex]::Escape("PackageId=$PackageId") })
    $candidatePool = if ($packageMatches.Count -gt 0) { $packageMatches } else { $detailedMatches }
    $committedCandidates = @($candidatePool | Where-Object { Test-IntuneAppHasCommittedContent -App $_ })

    if ($committedCandidates.Count -eq 1) { return $committedCandidates[0] }
    if ($committedCandidates.Count -gt 1) {
        $ids = ($committedCandidates | ForEach-Object { Format-IntuneAppCandidate -App $_ }) -join '; '
        throw "Multiple committed Intune apps match display name '$DisplayName'. Specify -ExistingAppId or use -ForceCreateNew. Matches: $ids"
    }

    $ids = ($candidatePool | ForEach-Object { Format-IntuneAppCandidate -App $_ }) -join '; '
    throw "Intune app '$DisplayName' exists but no matching app has a committed first content version. Delete the incomplete app or use -ForceCreateNew. Matches: $ids"
}

function Copy-OrderedHashtable {
    param([object]$InputObject)
    $copy = [ordered]@{}
    foreach ($key in $InputObject.Keys) { $copy[$key] = $InputObject[$key] }
    return $copy
}

function ConvertTo-Base64Utf8 {
    param([AllowNull()][string]$Value)
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$Value))
}

function Set-DetectScriptPackageMetadata {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptContent,
        [Parameter(Mandatory = $true)][string]$PackageId,
        [Parameter(Mandatory = $true)][string]$PackageVersion
    )

    $escapedPackageId = $PackageId -replace "'", "''"
    $escapedPackageVersion = $PackageVersion -replace "'", "''"
    $updated = [regex]::Replace($ScriptContent, "(?m)^(\s*\`$packageId\s*=\s*)'[^']*'", "`$1'$escapedPackageId'", 1)
    $updated = [regex]::Replace($updated, "(?m)^(\s*\`$packageVersion\s*=\s*)'[^']*'", "`$1'$escapedPackageVersion'", 1)
    return $updated
}
function New-EndpointRequirementRule {
    param(
        [Parameter(Mandatory = $true)][string]$Language,
        [string]$SetupCacheFolder,
        [switch]$RequireSetupCache
    )

    $cacheCheck = ''
    if ($RequireSetupCache) {
        $escapedCacheFolder = $SetupCacheFolder.Replace("'", "''")
        $cacheCheck = @"
    `$cachePath = Join-Path 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia' '$escapedCacheFolder'
    if (-not (Test-Path -LiteralPath (Join-Path `$cachePath 'setup.exe') -PathType Leaf)) { Write-Output 'MISSING_CACHE_SETUP_EXE'; exit 0 }
    if (-not (Test-Path -LiteralPath (Join-Path `$cachePath 'sources\install.wim') -PathType Leaf)) { Write-Output 'MISSING_CACHE_INSTALL_WIM'; exit 0 }
"@
    }

    $script = @"
try {
    `$locale = ''
    try {
        `$languageKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language'
        `$installLanguage = [string](Get-ItemProperty -LiteralPath `$languageKey -ErrorAction Stop).InstallLanguage
        if (-not [string]::IsNullOrWhiteSpace(`$installLanguage)) {
            `$lcid = [Convert]::ToInt32(`$installLanguage, 16)
            `$locale = ([System.Globalization.CultureInfo]::GetCultureInfo(`$lcid)).Name
        }
    }
    catch { `$locale = '' }
    if ([string]::IsNullOrWhiteSpace(`$locale)) {
        try { `$locale = [System.Globalization.CultureInfo]::InstalledUICulture.Name }
        catch { `$locale = '' }
    }
    if ([string]::IsNullOrWhiteSpace(`$locale)) { `$locale = 'UNKNOWN' }
    if (`$locale -ine '$Language') { Write-Output `$locale; exit 0 }
$cacheCheck    Write-Output 'OK'
    exit 0
}
catch {
    Write-Output 'UNKNOWN'
    exit 0
}
"@

    $label = "Windows setup language must match $Language"
    if ($RequireSetupCache) { $label = "$label and local setup cache must exist" }

    return [ordered]@{
        '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptRequirement'
        displayName = $label
        enforceSignatureCheck = $false
        runAs32Bit = $false
        runAsAccount = 'system'
        scriptContent = ConvertTo-Base64Utf8 -Value $script
        detectionType = 'string'
        operator = 'equal'
        detectionValue = 'OK'
    }
}

function Get-GraphObjectDiagnosticText {
    param([object]$Value)

    if ($null -eq $Value) { return '<null>' }
    $pairs = New-Object System.Collections.ArrayList
    foreach ($property in @($Value.PSObject.Properties | Sort-Object Name)) {
        if ($property.Name -match 'azureStorageUri|Uri|Sas|Secret|Token') { continue }
        $propertyValue = $property.Value
        if ($null -eq $propertyValue) { $propertyValue = '<null>' }
        elseif ($propertyValue -is [string]) { $propertyValue = [string]$propertyValue }
        elseif ($propertyValue -is [ValueType]) { $propertyValue = [string]$propertyValue }
        else { $propertyValue = ($propertyValue | ConvertTo-Json -Depth 6 -Compress) }
        [void]$pairs.Add(('{0}={1}' -f $property.Name,$propertyValue))
    }

    if ($pairs.Count -eq 0) { return '<no diagnostic properties>' }
    return ($pairs.ToArray() -join '; ')
}

function Wait-GraphContentFileState {
    param(
        [string]$Uri,
        [string[]]$SuccessStates,
        [string[]]$FailureStates,
        [int]$PollSeconds,
        [int]$TimeoutMinutes,
        [switch]$RequireAzureStorageUri
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $lastFile = $null
    do {
        $file = Invoke-GraphJson -Method GET -Uri $Uri
        $lastFile = $file
        $state = [string]$file.uploadState
        if ($FailureStates -contains $state) {
            $detail = Get-GraphObjectDiagnosticText -Value $file
            throw "Intune content file entered failure state: $state; Detail=$detail"
        }
        if (($SuccessStates -contains $state) -and (-not $RequireAzureStorageUri -or -not [string]::IsNullOrWhiteSpace([string]$file.azureStorageUri))) { return $file }
        Write-Step "Waiting for Intune content file state. Current=$state"
        Start-Sleep -Seconds $PollSeconds
    } while ((Get-Date) -lt $deadline)

    $lastDetail = Get-GraphObjectDiagnosticText -Value $lastFile
    throw "Timed out waiting for Intune content file state. Expected=$($SuccessStates -join ','); LastDetail=$lastDetail"
}

function Join-SasQuery {
    param([string]$Uri, [string]$Query)
    if ($Uri.Contains('?')) { return "$Uri&$Query" }
    return "$Uri`?$Query"
}

function Send-AzureHttpRequestWithRetry {
    param(
        [System.Net.Http.HttpClient]$Client,
        [scriptblock]$RequestFactory,
        [string]$Operation,
        [int]$MaxRetries
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $request = $null
        $response = $null
        try {
            $request = & $RequestFactory
            $response = $Client.SendAsync($request).GetAwaiter().GetResult()
            if ($response.IsSuccessStatusCode) { return }

            $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if ($attempt -ge $MaxRetries) { throw "$Operation failed. Status=$([int]$response.StatusCode); Body=$body" }
            Write-Step ("{0} attempt {1}/{2} failed. Status={3}. Retrying." -f $Operation,$attempt,$MaxRetries,[int]$response.StatusCode)
        }
        catch {
            if ($attempt -ge $MaxRetries) { throw ("{0} failed after {1} attempt(s): {2}" -f $Operation,$MaxRetries,$_.Exception.Message) }
            Write-Step ("{0} attempt {1}/{2} failed: {3}. Retrying." -f $Operation,$attempt,$MaxRetries,$_.Exception.Message)
        }
        finally {
            if ($response) { $response.Dispose() }
            if ($request) { $request.Dispose() }
        }

        Start-Sleep -Seconds ([Math]::Min(60, 5 * $attempt))
    }
}

function Get-SasExpiryUtc {
    param([string]$Uri)

    try {
        $query = ([Uri]$Uri).Query.TrimStart('?')
        foreach ($part in @($query -split '&')) {
            $nameValue = $part -split '=', 2
            if ($nameValue.Count -eq 2 -and $nameValue[0] -eq 'se') {
                $decoded = [System.Net.WebUtility]::UrlDecode($nameValue[1])
                return ([datetimeoffset]::Parse($decoded, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal)).UtcDateTime
            }
        }
    }
    catch { }

    return [datetime]::MinValue
}

function Request-IntuneAzureStorageUriRenewal {
    param(
        [Parameter(Mandatory = $true)][string]$FileUri,
        [int]$PollSeconds,
        [int]$TimeoutMinutes
    )

    Write-Step 'Requesting renewed Azure staging blob SAS from Intune.'
    Invoke-GraphJson -Method POST -Uri "$FileUri/renewUpload" -Body @{} | Out-Null
    $renewedFile = Wait-GraphContentFileState -Uri $FileUri -SuccessStates @('azureStorageUriRenewalSuccess') -FailureStates @('azureStorageUriRenewalFailed','azureStorageUriRequestFailed') -PollSeconds $PollSeconds -TimeoutMinutes $TimeoutMinutes -RequireAzureStorageUri
    $renewedUri = [string]$renewedFile.azureStorageUri
    if ([string]::IsNullOrWhiteSpace($renewedUri)) { throw 'Intune renewed upload state did not include azureStorageUri.' }
    return $renewedUri
}

function Send-AzureBlockBlobFromIntuneWin {
    param(
        [string]$Path,
        [string]$ContentEntryName,
        [string]$AzureStorageUri,
        [int]$BlockSizeMB,
        [int]$MaxRetries,
        [string]$FileUri,
        [int]$PollSeconds = 10,
        [int]$PollTimeoutMinutes = 30,
        [int]$RenewBeforeExpiryMinutes = 5
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.Net.Http

    $currentAzureStorageUri = $AzureStorageUri
    $sasExpiryUtc = Get-SasExpiryUtc -Uri $currentAzureStorageUri

    function Use-FreshAzureStorageUri {
        if ([string]::IsNullOrWhiteSpace($FileUri)) { return }

        $renewAtUtc = (Get-Date).ToUniversalTime().AddMinutes($RenewBeforeExpiryMinutes)
        if ($sasExpiryUtc -ne [datetime]::MinValue -and $sasExpiryUtc -le $renewAtUtc) {
            Write-Step ("Azure staging blob SAS expires at {0:u}; renewing before continuing upload." -f $sasExpiryUtc)
            $script:SmartM365RenewedAzureStorageUri = Request-IntuneAzureStorageUriRenewal -FileUri $FileUri -PollSeconds $PollSeconds -TimeoutMinutes $PollTimeoutMinutes
            $script:SmartM365RenewedAzureStorageUriExpiryUtc = Get-SasExpiryUtc -Uri $script:SmartM365RenewedAzureStorageUri
        }
    }

    $blockSize = [Math]::Max(4, $BlockSizeMB) * 1MB
    $blockIds = New-Object System.Collections.Generic.List[string]
    $client = New-Object System.Net.Http.HttpClient
    $client.Timeout = [TimeSpan]::FromMinutes(30)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $ContentEntryName } | Select-Object -First 1
        if (-not $entry) { throw "Encrypted package content entry not found: $ContentEntryName" }
        $stream = $entry.Open()
        try {
            $buffer = New-Object byte[] $blockSize
            $index = 0
            $sent = [int64]0
            do {
                $read = $stream.Read($buffer, 0, $buffer.Length)
                if ($read -le 0) { break }

                $payload = New-Object byte[] $read
                [Array]::Copy($buffer, $payload, $read)

                Use-FreshAzureStorageUri
                if (Test-Path -LiteralPath 'Variable:\script:SmartM365RenewedAzureStorageUri') {
                    $currentAzureStorageUri = $script:SmartM365RenewedAzureStorageUri
                    $sasExpiryUtc = $script:SmartM365RenewedAzureStorageUriExpiryUtc
                    Remove-Variable -Name SmartM365RenewedAzureStorageUri -Scope Script -ErrorAction SilentlyContinue
                    Remove-Variable -Name SmartM365RenewedAzureStorageUriExpiryUtc -Scope Script -ErrorAction SilentlyContinue
                    if ($sasExpiryUtc -ne [datetime]::MinValue) { Write-Step ("Renewed Azure staging blob SAS expires at {0:u}." -f $sasExpiryUtc) }
                }

                $blockId = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('block-{0:D8}' -f $index)))
                $encodedBlockId = [Uri]::EscapeDataString($blockId)
                $blockUri = Join-SasQuery -Uri $currentAzureStorageUri -Query "comp=block&blockid=$encodedBlockId"
                $currentIndex = $index

                try {
                    Send-AzureHttpRequestWithRetry -Client $client -MaxRetries $MaxRetries -Operation "Azure block upload index=$currentIndex" -RequestFactory {
                        $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, [Uri]$blockUri)
                        $request.Headers.Add('x-ms-version', '2020-10-02')
                        $request.Content = New-Object System.Net.Http.ByteArrayContent -ArgumentList (, $payload)
                        $request
                    }
                }
                catch {
                    $message = [string]$_.Exception.Message
                    if ([string]::IsNullOrWhiteSpace($FileUri) -or ($message -notmatch 'Status=403|AuthenticationFailed|Signature not valid in the specified time frame')) { throw }

                    Write-Step ("Azure block upload index=$currentIndex failed with expired/invalid SAS. Requesting renewal and retrying the same block.")
                    $currentAzureStorageUri = Request-IntuneAzureStorageUriRenewal -FileUri $FileUri -PollSeconds $PollSeconds -TimeoutMinutes $PollTimeoutMinutes
                    $sasExpiryUtc = Get-SasExpiryUtc -Uri $currentAzureStorageUri
                    if ($sasExpiryUtc -ne [datetime]::MinValue) { Write-Step ("Renewed Azure staging blob SAS expires at {0:u}." -f $sasExpiryUtc) }
                    $blockUri = Join-SasQuery -Uri $currentAzureStorageUri -Query "comp=block&blockid=$encodedBlockId"

                    Send-AzureHttpRequestWithRetry -Client $client -MaxRetries $MaxRetries -Operation "Azure block upload index=$currentIndex after SAS renewal" -RequestFactory {
                        $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, [Uri]$blockUri)
                        $request.Headers.Add('x-ms-version', '2020-10-02')
                        $request.Content = New-Object System.Net.Http.ByteArrayContent -ArgumentList (, $payload)
                        $request
                    }
                }

                $blockIds.Add($blockId) | Out-Null
                $sent += $read
                $index++
                if (($index % 40) -eq 0) { Write-Step ("Uploaded {0:N2} GB to Azure staging blob." -f ($sent / 1GB)) }
            } while ($true)
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $zip.Dispose()
        $client.Dispose()
    }

    if ($blockIds.Count -eq 0) { throw 'No Azure block was uploaded.' }

    $xmlBuilder = New-Object System.Text.StringBuilder
    [void]$xmlBuilder.Append('<?xml version="1.0" encoding="utf-8"?><BlockList>')
    foreach ($blockId in $blockIds) { [void]$xmlBuilder.AppendFormat('<Latest>{0}</Latest>', [System.Security.SecurityElement]::Escape($blockId)) }
    [void]$xmlBuilder.Append('</BlockList>')

    Use-FreshAzureStorageUri
    if (Test-Path -LiteralPath 'Variable:\script:SmartM365RenewedAzureStorageUri') {
        $currentAzureStorageUri = $script:SmartM365RenewedAzureStorageUri
        $sasExpiryUtc = $script:SmartM365RenewedAzureStorageUriExpiryUtc
        Remove-Variable -Name SmartM365RenewedAzureStorageUri -Scope Script -ErrorAction SilentlyContinue
        Remove-Variable -Name SmartM365RenewedAzureStorageUriExpiryUtc -Scope Script -ErrorAction SilentlyContinue
        if ($sasExpiryUtc -ne [datetime]::MinValue) { Write-Step ("Renewed Azure staging blob SAS expires at {0:u}." -f $sasExpiryUtc) }
    }

    $client2 = New-Object System.Net.Http.HttpClient
    $client2.Timeout = [TimeSpan]::FromMinutes(30)
    try {
        $commitUri = Join-SasQuery -Uri $currentAzureStorageUri -Query 'comp=blocklist'
        $blockList = $xmlBuilder.ToString()
        Send-AzureHttpRequestWithRetry -Client $client2 -MaxRetries $MaxRetries -Operation 'Azure block list commit' -RequestFactory {
            $request2 = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, [Uri]$commitUri)
            $request2.Headers.Add('x-ms-version', '2020-10-02')
            $request2.Content = New-Object System.Net.Http.StringContent($blockList, [Text.Encoding]::UTF8, 'application/xml')
            $request2
        }
    }
    finally {
        $client2.Dispose()
    }
}
$resolvedIntuneWinPath = (Resolve-Path -LiteralPath $IntuneWinPath).ProviderPath
if (-not (Test-Path -LiteralPath $resolvedIntuneWinPath -PathType Leaf)) { throw "IntuneWin package not found: $IntuneWinPath" }
if ([IO.Path]::GetExtension($resolvedIntuneWinPath) -ne '.intunewin') { throw "Expected a .intunewin file: $resolvedIntuneWinPath" }

$companionMetadata = Read-CompanionPackageMetadata -Path $resolvedIntuneWinPath
if ([string]::IsNullOrWhiteSpace($PackageId) -and $companionMetadata.PackageId) { $PackageId = [string]$companionMetadata.PackageId }
if ([string]::IsNullOrWhiteSpace($PackageId)) { $PackageId = [IO.Path]::GetFileNameWithoutExtension($resolvedIntuneWinPath) }
if ([string]::IsNullOrWhiteSpace($PackageVersion) -and $companionMetadata.PackageVersion) { $PackageVersion = [string]$companionMetadata.PackageVersion }
if ($Language -eq 'fr-FR' -and $companionMetadata.Language) { $Language = [string]$companionMetadata.Language }
$languageFromPackageId = Get-LanguageFromPackageId -Value $PackageId
if ($Language -eq 'fr-FR' -and $languageFromPackageId) { $Language = $languageFromPackageId }
if ([string]::IsNullOrWhiteSpace($PackageVersion)) { throw "PackageVersion is required for Intune detection. Provide -PackageVersion or place the generated Detect.ps1 next to the .intunewin package." }
$packageDir = Split-Path -Parent $resolvedIntuneWinPath
$detectScriptPath = Join-Path $packageDir 'Detect.ps1'
if (-not (Test-Path -LiteralPath $detectScriptPath -PathType Leaf)) { throw "Generated Detect.ps1 not found next to the .intunewin package: $detectScriptPath" }
$detectScriptContent = Get-Content -LiteralPath $detectScriptPath -Raw -ErrorAction Stop
if ([string]::IsNullOrWhiteSpace($detectScriptContent)) { throw "Generated Detect.ps1 is empty: $detectScriptPath" }
$detectScriptContent = Set-DetectScriptPackageMetadata -ScriptContent $detectScriptContent -PackageId $PackageId -PackageVersion $PackageVersion
if ([string]::IsNullOrWhiteSpace($RequirementLanguage)) { $RequirementLanguage = $Language }
$setupCacheFolder = [string]$companionMetadata.SetupCacheFolder
if ([string]::IsNullOrWhiteSpace($setupCacheFolder)) { $setupCacheFolder = "Win11-$Language" }
$packageMode = [string]$companionMetadata.PackageMode
$requiresExistingSetupCache = [bool]$companionMetadata.RequiresExistingSetupCache
if ([string]::IsNullOrWhiteSpace($packageMode) -and $PackageId -match '-WithCacheOnly$') {
    $packageMode = 'WithCacheOnly'
    $requiresExistingSetupCache = $true
}
if ([string]::IsNullOrWhiteSpace($packageMode)) { $packageMode = 'WithMedia' }
if ([string]::IsNullOrWhiteSpace($DisplayName) -and $companionMetadata.DisplayName) { $DisplayName = [string]$companionMetadata.DisplayName }
if ([string]::IsNullOrWhiteSpace($DisplayName) -and $PackageId -match '-WithCacheOnly$') { $DisplayName = "Windows11UpgradeToolkit-$Language-WithCacheOnly" }
if ([string]::IsNullOrWhiteSpace($DisplayName)) {
    $DisplayName = "Windows11UpgradeToolkit-$Language"
}
if ([string]::IsNullOrWhiteSpace($Description)) {
    if ($requiresExistingSetupCache) { $Description = "SmartM365 Windows 11 Upgrade Toolkit cache-only package for Windows setup language $Language. Requires an existing local setup media cache and starts the upgrade task asynchronously." }
    else { $Description = "SmartM365 Windows 11 Upgrade Toolkit package for Windows setup language $Language. Installs local media cache and starts the upgrade task asynchronously." }
}

$effectiveMinimumFreeDiskSpaceInMB = [int]$MinimumFreeDiskSpaceInMB
if ($effectiveMinimumFreeDiskSpaceInMB -lt 0) {
    if ($packageMode -eq 'WithMedia') { $effectiveMinimumFreeDiskSpaceInMB = 40960 }
    else { $effectiveMinimumFreeDiskSpaceInMB = 0 }
}

if ($UpdateMetadataOnly -and $ForceCreateNew) { throw '-UpdateMetadataOnly cannot be used with -ForceCreateNew because no content upload or app creation is performed.' }

Write-Step "Reading IntuneWin metadata: $resolvedIntuneWinPath"
$metadata = Read-IntuneWinMetadata -Path $resolvedIntuneWinPath
$encryptedSize = [int64]$metadata.SizeEncrypted
$plainSize = [int64]$metadata.UnencryptedContentSize

$appBody = [ordered]@{
    '@odata.type' = '#microsoft.graph.win32LobApp'
    displayName = $DisplayName
    description = $Description
    publisher = $Publisher
    developer = $Developer
    owner = $Owner
    notes = "PackageId=$PackageId; PackageVersion=$PackageVersion; Language=$Language; RequirementLanguage=$RequirementLanguage; PackageMode=$packageMode; RequiresExistingSetupCache=$requiresExistingSetupCache; MinimumFreeDiskSpaceInMB=$effectiveMinimumFreeDiskSpaceInMB"
    isFeatured = $false
    privacyInformationUrl = $null
    informationUrl = $null
    fileName = [string]$metadata.FileName
    setupFilePath = [string]$metadata.SetupFile
    installCommandLine = $InstallCommandLine
    uninstallCommandLine = $UninstallCommandLine
    installExperience = @{
        '@odata.type' = '#microsoft.graph.win32LobAppInstallExperience'
        runAsAccount = $RunAsAccount
        deviceRestartBehavior = $DeviceRestartBehavior
    }
    minimumSupportedOperatingSystem = @{
        '@odata.type' = '#microsoft.graph.windowsMinimumOperatingSystem'
        v10_1607 = $true
    }
    applicableArchitectures = 'x64'
    requirementRules = @(
        if (-not $DisableLanguageRequirementRule) { New-EndpointRequirementRule -Language $RequirementLanguage -SetupCacheFolder $setupCacheFolder -RequireSetupCache:$requiresExistingSetupCache }
    )
    detectionRules = @(
        @{
            '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptDetection'
            enforceSignatureCheck = $false
            runAs32Bit = $false
            scriptContent = ConvertTo-Base64Utf8 -Value $detectScriptContent
        }
    )
    returnCodes = @(
        @{ returnCode = 0; type = 'success' },
        @{ returnCode = 1707; type = 'success' },
        @{ returnCode = 3010; type = 'softReboot' },
        @{ returnCode = 1641; type = 'hardReboot' },
        @{ returnCode = 1618; type = 'retry' }
    )
    installCommandLineTimeoutInMinutes = $InstallCommandLineTimeoutMinutes
}
if ($effectiveMinimumFreeDiskSpaceInMB -gt 0) { $appBody['minimumFreeDiskSpaceInMB'] = $effectiveMinimumFreeDiskSpaceInMB }

if (-not [string]::IsNullOrWhiteSpace($FinalizeExistingAppId)) {
    if ([string]::IsNullOrWhiteSpace($FinalizeContentVersionId)) { throw 'FinalizeContentVersionId is required when FinalizeExistingAppId is used.' }
    Write-Step "Finalizing existing Intune app: AppId=$FinalizeExistingAppId; ContentVersion=$FinalizeContentVersionId; DisplayName=$DisplayName"
    if (-not $PSCmdlet.ShouldProcess($FinalizeExistingAppId, "Finalize existing Intune Win32 app content version $FinalizeContentVersionId")) {
        [pscustomobject]@{
            AppId = $FinalizeExistingAppId
            DisplayName = $DisplayName
            ContentVersionId = $FinalizeContentVersionId
            Action = 'FinalizeExistingApp'
        }
        return
    }

    Assert-GraphCommand
    if (-not $NoConnect) {
        Write-Step 'Connecting to Microsoft Graph.'
        Connect-MgGraph -Scopes @('DeviceManagementApps.ReadWrite.All') -NoWelcome | Out-Null
    }

    $finalizeBody = [ordered]@{
        '@odata.type' = '#microsoft.graph.win32LobApp'
        displayName = $DisplayName
        committedContentVersion = $FinalizeContentVersionId
    }
    Invoke-GraphJson -Method PATCH -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$FinalizeExistingAppId" -Body $finalizeBody | Out-Null
    Write-Step 'Existing Intune app content version committed.'
    [pscustomobject]@{
        AppId = $FinalizeExistingAppId
        DisplayName = $DisplayName
        ContentVersionId = $FinalizeContentVersionId
        Action = 'FinalizeExistingApp'
    }
    return
}

if ($UpdateMetadataOnly) {
    Write-Step "Updating app metadata only: $DisplayName"
    if (-not $PSCmdlet.ShouldProcess($DisplayName, 'Update Intune Win32 app metadata without uploading package content')) {
        [pscustomobject]@{
            DisplayName = $DisplayName
            PackageId = $PackageId
            PackageVersion = $PackageVersion
            Language = $Language
            RequirementLanguage = $RequirementLanguage
            LanguageRequirementRuleEnabled = (-not [bool]$DisableLanguageRequirementRule)
            PackageMode = $packageMode
            RequiresExistingSetupCache = $requiresExistingSetupCache
            MinimumFreeDiskSpaceInMB = $effectiveMinimumFreeDiskSpaceInMB
            IntuneWinPath = $resolvedIntuneWinPath
            DetectionScriptPath = $detectScriptPath
            ExistingAppId = $ExistingAppId
            Action = 'UpdateMetadataOnly'
        }
        return
    }

    Assert-GraphCommand
    if (-not $NoConnect) {
        Write-Step 'Connecting to Microsoft Graph.'
        Connect-MgGraph -Scopes @('DeviceManagementApps.ReadWrite.All') -NoWelcome | Out-Null
    }

    $app = Resolve-ExistingIntuneWin32App -GraphBaseUri $GraphBaseUri -AppId $ExistingAppId -DisplayName $DisplayName -PackageId $PackageId
    if (-not $app) { throw "No existing Intune app found for metadata-only update. DisplayName=$DisplayName; PackageId=$PackageId. Specify -ExistingAppId if needed." }
    $appId = [string]$app.id
    if ([string]::IsNullOrWhiteSpace($appId)) { throw 'Existing Intune app lookup did not return a mobile app id.' }

    $metadataPatchBody = Copy-OrderedHashtable -InputObject $appBody
    foreach ($metadataOnlyExcludedProperty in @('applicableArchitectures','fileName','setupFilePath','requirementRules','detectionRules','notes')) {
        if ($metadataPatchBody.Contains($metadataOnlyExcludedProperty)) { $metadataPatchBody.Remove($metadataOnlyExcludedProperty) }
    }
    if ($UpdateDetectionRules) {
        $metadataPatchBody['notes'] = $appBody['notes']
        $metadataPatchBody['detectionRules'] = $appBody['detectionRules']
    }
    if ($effectiveMinimumFreeDiskSpaceInMB -le 0) { $metadataPatchBody['minimumFreeDiskSpaceInMB'] = $null }

    Invoke-GraphJson -Method PATCH -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$appId" -Body $metadataPatchBody | Out-Null
    Write-Step "Existing Intune app metadata updated without content upload: $appId"
    [pscustomobject]@{
        AppId = $appId
        DisplayName = $DisplayName
        PackageId = $PackageId
        PackageVersion = $PackageVersion
        Language = $Language
        RequirementLanguage = $RequirementLanguage
        LanguageRequirementRuleEnabled = (-not [bool]$DisableLanguageRequirementRule)
        PackageMode = $packageMode
        RequiresExistingSetupCache = $requiresExistingSetupCache
        MinimumFreeDiskSpaceInMB = $effectiveMinimumFreeDiskSpaceInMB
        DetectionScriptPath = $detectScriptPath
        UpdatedExistingApp = $true
        UploadedContent = $false
        Action = 'UpdateMetadataOnly'
    }
    return
}

Write-Step "Publishing app: $DisplayName"
if (-not $PSCmdlet.ShouldProcess($DisplayName, 'Create or update Intune Win32 app and upload package content')) {
    [pscustomobject]@{
        DisplayName = $DisplayName
        PackageId = $PackageId
        PackageVersion = $PackageVersion
        Language = $Language
        RequirementLanguage = $RequirementLanguage
        LanguageRequirementRuleEnabled = (-not [bool]$DisableLanguageRequirementRule)
        PackageMode = $packageMode
        RequiresExistingSetupCache = $requiresExistingSetupCache
        MinimumFreeDiskSpaceInMB = $effectiveMinimumFreeDiskSpaceInMB
        IntuneWinPath = $resolvedIntuneWinPath
        FileName = $metadata.FileName
        SetupFile = $metadata.SetupFile
        SizeEncrypted = $encryptedSize
        UnencryptedContentSize = $plainSize
        DetectionScriptPath = $detectScriptPath
        UploadBlockSizeMB = $UploadBlockSizeMB
        AzureUploadMaxRetries = $AzureUploadMaxRetries
        ForceCreateNew = [bool]$ForceCreateNew
        ExistingAppId = $ExistingAppId
    }
    return
}

Assert-GraphCommand
if (-not $NoConnect) {
    Write-Step 'Connecting to Microsoft Graph.'
    Connect-MgGraph -Scopes @('DeviceManagementApps.ReadWrite.All') -NoWelcome | Out-Null
}

$app = $null
$appId = ''
$updatedExistingApp = $false
if (-not $ForceCreateNew) {
    $app = Resolve-ExistingIntuneWin32App -GraphBaseUri $GraphBaseUri -AppId $ExistingAppId -DisplayName $DisplayName -PackageId $PackageId
}

if ($app) {
    $appId = [string]$app.id
    if ([string]::IsNullOrWhiteSpace($appId)) { throw 'Existing Intune app lookup did not return a mobile app id.' }
    $updatedExistingApp = $true
    Write-Step "Using existing Intune app: $appId ($DisplayName)"
}
else {
    $app = Invoke-GraphJson -Method POST -Uri "$GraphBaseUri/deviceAppManagement/mobileApps" -Body $appBody
    $appId = [string]$app.id
    if ([string]::IsNullOrWhiteSpace($appId)) { throw 'Graph did not return a mobile app id.' }
    Write-Step "Created Intune app: $appId"
}

$contentVersion = Invoke-GraphJson -Method POST -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions" -Body @{}
$contentVersionId = [string]$contentVersion.id
if ([string]::IsNullOrWhiteSpace($contentVersionId)) { throw 'Graph did not return a content version id.' }
Write-Step "Created content version: $contentVersionId"

$fileBody = [ordered]@{
    '@odata.type' = '#microsoft.graph.mobileAppContentFile'
    name = [string]$metadata.FileName
    size = $plainSize
    sizeEncrypted = $encryptedSize
    manifest = $null
    isDependency = $false
}
$file = Invoke-GraphJson -Method POST -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions/$contentVersionId/files" -Body $fileBody
$fileId = [string]$file.id
if ([string]::IsNullOrWhiteSpace($fileId)) { throw 'Graph did not return a content file id.' }
$fileUri = "$GraphBaseUri/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions/$contentVersionId/files/$fileId"
Write-Step "Created content file: $fileId"

$file = Wait-GraphContentFileState -Uri $fileUri -SuccessStates @('azureStorageUriRequestSuccess','azureStorageUriRenewalSuccess') -FailureStates @('azureStorageUriRequestFailed','azureStorageUriRenewalFailed') -PollSeconds $PollSeconds -TimeoutMinutes $PollTimeoutMinutes -RequireAzureStorageUri
Write-Step "Uploading encrypted package payload to Azure staging blob. BlockSizeMB=$UploadBlockSizeMB; MaxRetries=$AzureUploadMaxRetries"
Send-AzureBlockBlobFromIntuneWin -Path $resolvedIntuneWinPath -ContentEntryName ([string]$metadata.ContentEntryName) -AzureStorageUri ([string]$file.azureStorageUri) -BlockSizeMB $UploadBlockSizeMB -MaxRetries $AzureUploadMaxRetries -FileUri $fileUri -PollSeconds $PollSeconds -PollTimeoutMinutes $PollTimeoutMinutes
Write-Step 'Azure upload completed.'

$commitBody = [ordered]@{
    fileEncryptionInfo = [ordered]@{
        '@odata.type' = '#microsoft.graph.fileEncryptionInfo'
        encryptionKey = [string]$metadata.EncryptionKey
        macKey = [string]$metadata.MacKey
        initializationVector = [string]$metadata.InitializationVector
        mac = [string]$metadata.Mac
        profileIdentifier = [string]$metadata.ProfileIdentifier
        fileDigest = [string]$metadata.FileDigest
        fileDigestAlgorithm = [string]$metadata.FileDigestAlgorithm
    }
}
Invoke-GraphJson -Method POST -Uri "$fileUri/commit" -Body $commitBody | Out-Null
Write-Step 'Commit requested for uploaded package.'

Wait-GraphContentFileState -Uri $fileUri -SuccessStates @('commitFileSuccess') -FailureStates @('commitFileFailed') -PollSeconds $PollSeconds -TimeoutMinutes $PollTimeoutMinutes | Out-Null
$finalPatchBody = Copy-OrderedHashtable -InputObject $appBody
if ($finalPatchBody.Contains('applicableArchitectures')) { $finalPatchBody.Remove('applicableArchitectures') }
$finalPatchBody['committedContentVersion'] = $contentVersionId
Invoke-GraphJson -Method PATCH -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$appId" -Body $finalPatchBody | Out-Null
Write-Step 'App metadata and content version committed.'

[pscustomobject]@{
    AppId = $appId
    DisplayName = $DisplayName
    PackageId = $PackageId
    PackageVersion = $PackageVersion
    Language = $Language
    RequirementLanguage = $RequirementLanguage
    LanguageRequirementRuleEnabled = (-not [bool]$DisableLanguageRequirementRule)
    PackageMode = $packageMode
    RequiresExistingSetupCache = $requiresExistingSetupCache
    MinimumFreeDiskSpaceInMB = $effectiveMinimumFreeDiskSpaceInMB
    ContentVersionId = $contentVersionId
    ContentFileId = $fileId
    DetectionScriptPath = $detectScriptPath
    UpdatedExistingApp = [bool]$updatedExistingApp
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAV/Zx7YwGiUueO
# PXoBpWOmSiWBUoO5MgCTNSUCd0wGEKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIBD9CXwtPh/eHbHC1GFXYHhHL03cQv6J40Z1KBJKelkRMA0GCSqG
# SIb3DQEBAQUABIIBgHzWXjvJ9aAietUmH7wI/fG2BlEYXDLJthCd0jBLkdQkvTRc
# JKOHTryqq3ylrDjjQ/7SgXMjb0wjTYeB74t/5Mjxdxjnollprvp1wLhztGKRcLJ+
# XYAtXmehsv4ZeFz5CsV2YcP9S6tXEqI7136BS+3a7ThIAN2v4Pn1RdNwH2hVrxox
# z7rmOcgbC+9y3+zaoxjqZ9htBAS98PWpqTqdtj8Ed6Slri6GiwV8tNRncJw60d4a
# 29pyNaJGlea0VwMjBX5iikqquFP0TxvP/UU0bfhgaR30gstapdD+Rj9V9gg5zRuJ
# 1EaIMuicrPsZdHZIMynJ2P1+iHXGbXph41CKpt2oeZ8NYsuHzrcQc2YRM3gXy+t+
# xgcosfMP0DfZWwvB3wDTXT7I+asu6wtE5vx2H/n7KmAN8ZQQNcOTTuz0YM6kPp0s
# +8enodW4dOIpAnMLSM11zpA7zXRynVd7tXMvRg4dsAJzDxFTdKzxd7J/hKT0WVPA
# deBZSDTLbSW3dxSf96GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MzAwNTM2
# MjhaMC8GCSqGSIb3DQEJBDEiBCCuYKLAtKj24eQNdq1yKQ9w5fbT4HZV8QoAWNmg
# NRdv1TANBgkqhkiG9w0BAQEFAASCAgCWUtxOsJ95hyStMVah6tLH1/azlulaMEYy
# PJKNeD8kJvcIjuJWIO2zQoIHkOO8/llQlrqcY7mSMHEgCxnDzfiU/TFqiC2h7DRf
# +b47A8fOb8yHMcglP/lw3sRWxx+f6NXvmG1wnzh9DElgIE2N3eltM6FFWrsDftn6
# RKo7wOZkZQ5HG+wHOHxJfKj4ynRrUKTiuwL/No2vhX6R22YqIUYxCm4A/MyB+nwz
# yRxC7fr2d4VbU6lRd82m0D2dYvvyIs9fnyPLtGQREZ24+6+gWcf7eRgqrWhobfDP
# S9Oi5UZ8SJAp9XwHSKPBR79w17/AtDXOxZVM2HnBwOhSce4kqOHxePE3bAWjAPiV
# P256zqt6pkCANwjKYNYxqsViRAuJaAo2bbh9gsFaHUUF22/6CbWotgOwFafvXH+t
# 3Ey6aN8MIkuRKjCQQJYRV2aJVxNC227oOxiVVsl9Cr4/QjSVlhTBYKkVooT/tByp
# tv8IA0xpSjwpbLJzHXR2ZdB2gAhLmj2F1RNi11iLdFFKc9IaoJc9Eiou67YO8AVp
# 8RuHyODF+yUAqTw1rl+v856y1xd2JarlBD1NKFdDygWJrSwYfVETtAfyGgq/8GoX
# YTzoes4ZHQkwkTaL9OACAFnvHrTwur4Hz/D6FwlIOgMw2a3OlhHdhH5adXdu2IuL
# 46yICj/MRQ==
# SIG # End signature block
