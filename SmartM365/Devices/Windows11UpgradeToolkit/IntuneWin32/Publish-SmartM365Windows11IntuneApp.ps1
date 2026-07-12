<#
.SYNOPSIS
    Publishes a SmartM365 Windows 11 Upgrade Toolkit .intunewin package to Intune.
.DESCRIPTION
    Creates a Win32 LOB app in Intune with Microsoft Graph beta, uploads the encrypted package payload, commits the content version, and configures PowerShell detection for the generated language package.
.VERSION
    1.0.16
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
    `$locale = (Get-WinSystemLocale).Name
    if ([string]::IsNullOrWhiteSpace(`$locale)) { `$locale = 'UNKNOWN' }
    if (`$locale -ne '$Language') { Write-Output `$locale; exit 0 }
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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDYTe0EUuastwQF
# GcURteTG3ay3jNhVYPuCo75Dq15+naCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCClnhnQZ8eqpHAFuyYp
# 1JV1O0l5lHbapkMLB+vKNlow+jANBgkqhkiG9w0BAQEFAASCAYBK842Rju/B1M03
# p8jHuYHjA3WwWRKu26QbMOHdVrg4IYhmcspthRhbGdoRubA9eULXPWku7fStvvE0
# GRXjZz5m8SNV3nF3e1zMzF70jqnARgu/S1/Dr6Tf2R6rJTUqxShaSQnDkqbNyu9A
# uCKYFweesyTlHiRgMSYM2nX5SewvAF3zz9QkvsTxNplWTC+zfyRfeL3rI5qsvEij
# sjTgihXh3mDWC/f6MDV8fabgD9EmqqpUK/aGGdXdNMAdstAvQ7xdHfH+cz7Vbkmm
# 8EWTZL5n0j7CXJ6AobY2mzuL15dn3WuTEqP29rulIp223nTSREnSTiKq9bJGVYxg
# xYjRMq7rFdui/N3MKg9K/Xh2kBQOfmFaQr02GfCWCi98a20t4yXSxgTCBE/y5qfF
# O/KLrQc2viVUo1U+7P3pofldSdluzmo8m1EtCuOsdZu3kS4P2PYzBF6OKlx+SF3i
# KZxUI6ALi5xbccZsab0aPNAH/7RtLZNqQF4qDu+5na/+v1QgXlk=
# SIG # End signature block
