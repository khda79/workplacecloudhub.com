#Requires -Version 7.2

<#
.SYNOPSIS
    Publishes SmartM365 Endpoint Diagnostics Analyzer to Intune with interactive authentication.

.DESCRIPTION
    Preview is the default and performs no Microsoft Graph connection or tenant
    change. With -Execute, the script connects interactively, creates a new
    Intune Win32 app, uploads and commits the .intunewin content, optionally
    creates or reuses a pilot security group, and optionally assigns the app as
    Available or Required to that group. Available is the default. Use -ResumeAppId to safely continue a publication
    that failed after its app content was committed.

    The script never replaces an existing assignment set. It creates the pilot
    assignment individually and preserves all other assignments.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [string]$IntuneWinPath,

    [string]$DetectionScriptPath = (Join-Path $PSScriptRoot 'SmartM365-EndpointDiagnosticsAnalyzer-Detection.ps1'),
    [string]$PackageVersion = '',
    [string]$AppDisplayName = '',
    [string]$Description = 'SmartM365 desktop analyzer for Windows and Microsoft Intune endpoint diagnostics.',
    [string]$Publisher = 'WorkplaceCloudHub',
    [string]$Developer = 'WorkplaceCloudHub',
    [string]$Owner = 'SmartM365',
    [string]$TenantId = '',
    [string]$PilotGroupDisplayName = 'GG-INTUNE-SmartM365-EndpointDiagnosticsAnalyzer-Pilot',
    [string]$PilotGroupMailNickname = 'GG-INTUNE-SmartM365-EndpointDiagnosticsAnalyzer-Pilot',
    [string]$PilotGroupDescription = 'Pilot devices for SmartM365 Endpoint Diagnostics Analyzer.',
    [string]$PilotGroupId = '',
    [string]$SupersedeAppId = '',
    [string]$ResumeAppId = '',
    [ValidateSet('update', 'replace')]
    [string]$SupersedenceType = 'update',
    [switch]$CreatePilotGroup,
    [switch]$AssignPilotGroup,
    [ValidateSet('available', 'required')]
    [string]$PilotAssignmentIntent = 'available',
    [switch]$AllowDuplicateDisplayName,
    [switch]$Execute,
    [ValidateRange(4, 100)]
    [int]$UploadBlockSizeMB = 16,
    [ValidateRange(1, 10)]
    [int]$AzureUploadMaxRetries = 5,
    [ValidateRange(2, 60)]
    [int]$PollSeconds = 5,
    [ValidateRange(5, 120)]
    [int]$PollTimeoutMinutes = 30,
    [string]$GraphBaseUri = 'https://graph.microsoft.com/beta',
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$requiredScopes = @(
    'DeviceManagementApps.ReadWrite.All'
    'Group.ReadWrite.All'
)

function Write-Step {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Information `
        -MessageData ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Message) `
        -InformationAction Continue
}

function Get-NormalizedThumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function ConvertTo-Base64Utf8 {
    param([AllowNull()][string]$Value)
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$Value))
}

function ConvertTo-ODataStringLiteral {
    param([AllowNull()][string]$Value)
    return ([string]$Value -replace "'", "''")
}

function Assert-GraphCommand {
    foreach ($commandName in @('Connect-MgGraph','Disconnect-MgGraph','Get-MgContext','Invoke-MgGraphRequest')) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            throw ("Microsoft.Graph.Authentication is required. Missing command: {0}. Install with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -f $commandName)
        }
    }
}

function Get-XmlValue {
    param(
        [Parameter(Mandatory = $true)][xml]$Xml,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    foreach ($name in $Names) {
        $node = $Xml.SelectSingleNode("//*[local-name()='$name']")
        if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
            return [string]$node.InnerText
        }
        $attribute = $Xml.SelectSingleNode("//@*[local-name()='$name']")
        if ($attribute -and -not [string]::IsNullOrWhiteSpace($attribute.Value)) {
            return [string]$attribute.Value
        }
    }
    return ''
}

function Read-IntuneWinPackage {
    param([Parameter(Mandatory = $true)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $detectionEntry = $archive.Entries |
            Where-Object { $_.FullName -match '(^|/)Detection\.xml$' } |
            Select-Object -First 1
        if (-not $detectionEntry) {
            throw 'Detection.xml was not found inside the .intunewin package.'
        }

        $reader = New-Object IO.StreamReader($detectionEntry.Open())
        try {
            [xml]$detectionXml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        $contentEntry = $archive.Entries |
            Where-Object { $_.FullName -match '(^|/)Contents/.*\.(intunewin|bin)$' } |
            Sort-Object Length -Descending |
            Select-Object -First 1
        if (-not $contentEntry) {
            throw 'Encrypted package content was not found inside the .intunewin package.'
        }

        $metadata = [ordered]@{
            SetupFile              = Get-XmlValue -Xml $detectionXml -Names @('SetupFile','SetupFilePath','ApplicationName','Name')
            FileName               = Get-XmlValue -Xml $detectionXml -Names @('FileName','EncryptedFileName')
            UnencryptedContentSize = Get-XmlValue -Xml $detectionXml -Names @('UnencryptedContentSize','Size')
            EncryptionKey          = Get-XmlValue -Xml $detectionXml -Names @('EncryptionKey')
            MacKey                 = Get-XmlValue -Xml $detectionXml -Names @('MacKey')
            InitializationVector   = Get-XmlValue -Xml $detectionXml -Names @('InitializationVector')
            Mac                    = Get-XmlValue -Xml $detectionXml -Names @('Mac')
            ProfileIdentifier      = Get-XmlValue -Xml $detectionXml -Names @('ProfileIdentifier')
            FileDigest             = Get-XmlValue -Xml $detectionXml -Names @('FileDigest')
            FileDigestAlgorithm    = Get-XmlValue -Xml $detectionXml -Names @('FileDigestAlgorithm')
            ContentEntryName       = [string]$contentEntry.FullName
            SizeEncrypted          = [int64]$contentEntry.Length
        }

        if ([string]::IsNullOrWhiteSpace($metadata.FileName)) {
            $metadata.FileName = [IO.Path]::GetFileName($Path)
        }
        if ([string]::IsNullOrWhiteSpace($metadata.ProfileIdentifier)) {
            $metadata.ProfileIdentifier = 'ProfileVersion1'
        }
        if ([string]::IsNullOrWhiteSpace($metadata.FileDigestAlgorithm)) {
            $metadata.FileDigestAlgorithm = 'SHA256'
        }
        foreach ($requiredName in @(
            'SetupFile'
            'UnencryptedContentSize'
            'EncryptionKey'
            'MacKey'
            'InitializationVector'
            'Mac'
            'FileDigest'
        )) {
            if ([string]::IsNullOrWhiteSpace([string]$metadata[$requiredName])) {
                throw "Unable to read required IntuneWin metadata field: $requiredName"
            }
        }

        return [pscustomobject]$metadata
    }
    finally {
        $archive.Dispose()
    }
}

function Invoke-GraphJson {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET','POST','PATCH','DELETE')]
        [string]$Method,
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [object]$Body
    )

    if ($PSBoundParameters.ContainsKey('Body')) {
        $json = $Body | ConvertTo-Json -Depth 30
        return Invoke-MgGraphRequest `
            -Method $Method `
            -Uri $Uri `
            -Body $json `
            -ContentType 'application/json' `
            -OutputType PSObject `
            -ErrorAction Stop
    }

    return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType PSObject -ErrorAction Stop
}

function Get-GraphCollection {
    param([Parameter(Mandatory = $true)][string]$Uri)

    $items = New-Object Collections.Generic.List[object]
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $page = Invoke-GraphJson -Method GET -Uri $nextUri
        foreach ($item in @($page.value)) {
            $items.Add($item)
        }
        $nextUri = ''
        if ($page.PSObject.Properties['@odata.nextLink']) {
            $nextUri = [string]$page.'@odata.nextLink'
        }
    }
    return @($items.ToArray())
}

function Copy-OrderedHashtable {
    param([Parameter(Mandatory = $true)][Collections.IDictionary]$InputObject)

    $copy = [ordered]@{}
    foreach ($key in $InputObject.Keys) {
        $copy[$key] = $InputObject[$key]
    }
    return $copy
}

function Get-GraphObjectDiagnosticText {
    param([object]$Value)

    if ($null -eq $Value) {
        return '<null>'
    }
    $parts = New-Object Collections.Generic.List[string]
    foreach ($property in @($Value.PSObject.Properties | Sort-Object Name)) {
        if ($property.Name -match 'azureStorageUri|Uri|Sas|Secret|Token') {
            continue
        }
        $propertyValue = $property.Value
        if ($null -eq $propertyValue) {
            $propertyValue = '<null>'
        }
        elseif ($propertyValue -isnot [string] -and $propertyValue -isnot [ValueType]) {
            $propertyValue = $propertyValue | ConvertTo-Json -Depth 5 -Compress
        }
        $parts.Add(('{0}={1}' -f $property.Name,$propertyValue))
    }
    return ($parts -join '; ')
}

function Wait-GraphContentFileState {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string[]]$SuccessStates,
        [Parameter(Mandatory = $true)][string[]]$FailureStates,
        [int]$PollIntervalSeconds,
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
            throw ("Intune content file entered failure state: {0}; Detail={1}" -f
                $state,(Get-GraphObjectDiagnosticText -Value $file))
        }
        if (($SuccessStates -contains $state) -and
            (-not $RequireAzureStorageUri -or -not [string]::IsNullOrWhiteSpace([string]$file.azureStorageUri))) {
            return $file
        }
        Write-Step "Waiting for Intune content state. Current=$state"
        Start-Sleep -Seconds $PollIntervalSeconds
    } while ((Get-Date) -lt $deadline)

    throw ("Timed out waiting for Intune content state. Expected={0}; LastDetail={1}" -f
        ($SuccessStates -join ','),(Get-GraphObjectDiagnosticText -Value $lastFile))
}

function Join-SasQuery {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$Query
    )
    if ($Uri.Contains('?')) {
        return "$Uri&$Query"
    }
    return "$Uri`?$Query"
}

function Send-AzureHttpRequestWithRetry {
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$Client,
        [Parameter(Mandatory = $true)][scriptblock]$RequestFactory,
        [Parameter(Mandatory = $true)][string]$Operation,
        [int]$MaxRetries
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $request = $null
        $response = $null
        try {
            $request = & $RequestFactory
            $response = $Client.SendAsync($request).GetAwaiter().GetResult()
            if ($response.IsSuccessStatusCode) {
                return
            }
            $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if ($attempt -ge $MaxRetries) {
                throw "$Operation failed. Status=$([int]$response.StatusCode); Body=$body"
            }
            Write-Step ("{0} attempt {1}/{2} failed. Status={3}. Retrying." -f
                $Operation,$attempt,$MaxRetries,[int]$response.StatusCode)
        }
        catch {
            if ($attempt -ge $MaxRetries) {
                throw ("{0} failed after {1} attempt(s): {2}" -f
                    $Operation,$MaxRetries,$_.Exception.Message)
            }
            Write-Step ("{0} attempt {1}/{2} failed: {3}. Retrying." -f
                $Operation,$attempt,$MaxRetries,$_.Exception.Message)
        }
        finally {
            if ($response) {
                $response.Dispose()
            }
            if ($request) {
                $request.Dispose()
            }
        }
        Start-Sleep -Seconds ([Math]::Min(60,5 * $attempt))
    }
}

function Send-AzureBlockBlobFromIntuneWin {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ContentEntryName,
        [Parameter(Mandatory = $true)][string]$AzureStorageUri,
        [int]$BlockSizeMB,
        [int]$MaxRetries
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.Net.Http

    $blockSize = [Math]::Max(4,$BlockSizeMB) * 1MB
    $blockIds = New-Object Collections.Generic.List[string]
    $client = New-Object Net.Http.HttpClient
    $client.Timeout = [TimeSpan]::FromMinutes(30)
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.Entries |
            Where-Object { $_.FullName -eq $ContentEntryName } |
            Select-Object -First 1
        if (-not $entry) {
            throw "Encrypted package content entry not found: $ContentEntryName"
        }

        $stream = $entry.Open()
        try {
            $buffer = New-Object byte[] $blockSize
            $index = 0
            do {
                $read = $stream.Read($buffer,0,$buffer.Length)
                if ($read -le 0) {
                    break
                }

                $payload = New-Object byte[] $read
                [Array]::Copy($buffer,$payload,$read)
                $blockId = [Convert]::ToBase64String(
                    [Text.Encoding]::ASCII.GetBytes(('block-{0:D8}' -f $index))
                )
                $blockUri = Join-SasQuery `
                    -Uri $AzureStorageUri `
                    -Query ("comp=block&blockid={0}" -f [Uri]::EscapeDataString($blockId))
                $currentIndex = $index
                Send-AzureHttpRequestWithRetry `
                    -Client $client `
                    -MaxRetries $MaxRetries `
                    -Operation "Azure block upload index=$currentIndex" `
                    -RequestFactory {
                        $request = [Net.Http.HttpRequestMessage]::new(
                            [Net.Http.HttpMethod]::Put,
                            [Uri]$blockUri
                        )
                        $request.Headers.Add('x-ms-version','2020-10-02')
                        $request.Content = New-Object Net.Http.ByteArrayContent -ArgumentList (,$payload)
                        $request
                    }

                $blockIds.Add($blockId)
                $index++
            } while ($true)
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $archive.Dispose()
        $client.Dispose()
    }

    if ($blockIds.Count -eq 0) {
        throw 'No Azure block was uploaded.'
    }

    $xmlBuilder = New-Object Text.StringBuilder
    [void]$xmlBuilder.Append('<?xml version="1.0" encoding="utf-8"?><BlockList>')
    foreach ($blockId in $blockIds) {
        [void]$xmlBuilder.AppendFormat(
            '<Latest>{0}</Latest>',
            [Security.SecurityElement]::Escape($blockId)
        )
    }
    [void]$xmlBuilder.Append('</BlockList>')

    $commitClient = New-Object Net.Http.HttpClient
    $commitClient.Timeout = [TimeSpan]::FromMinutes(30)
    try {
        $commitUri = Join-SasQuery -Uri $AzureStorageUri -Query 'comp=blocklist'
        $blockList = $xmlBuilder.ToString()
        Send-AzureHttpRequestWithRetry `
            -Client $commitClient `
            -MaxRetries $MaxRetries `
            -Operation 'Azure block list commit' `
            -RequestFactory {
                $request = [Net.Http.HttpRequestMessage]::new(
                    [Net.Http.HttpMethod]::Put,
                    [Uri]$commitUri
                )
                $request.Headers.Add('x-ms-version','2020-10-02')
                $request.Content = New-Object Net.Http.StringContent(
                    $blockList,
                    [Text.Encoding]::UTF8,
                    'application/xml'
                )
                $request
            }
    }
    finally {
        $commitClient.Dispose()
    }
}

function Get-ExistingAppByDisplayName {
    param([Parameter(Mandatory = $true)][string]$DisplayName)

    $literal = ConvertTo-ODataStringLiteral -Value $DisplayName
    $filter = [Uri]::EscapeDataString("displayName eq '$literal'")
    return @(Get-GraphCollection -Uri "$GraphBaseUri/deviceAppManagement/mobileApps?`$filter=$filter")
}

function Resolve-PilotGroup {
    param(
        [string]$GroupId,
        [string]$DisplayName,
        [string]$MailNickname,
        [string]$GroupDescription,
        [switch]$CreateWhenMissing
    )

    if (-not [string]::IsNullOrWhiteSpace($GroupId)) {
        return Invoke-GraphJson `
            -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/groups/${GroupId}?`$select=id,displayName,securityEnabled,mailEnabled"
    }

    $literal = ConvertTo-ODataStringLiteral -Value $DisplayName
    $filter = [Uri]::EscapeDataString("displayName eq '$literal'")
    $groups = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$filter&`$select=id,displayName,securityEnabled,mailEnabled")
    if ($groups.Count -gt 1) {
        throw "Multiple groups use the proposed display name. Specify -PilotGroupId."
    }
    if ($groups.Count -eq 1) {
        if (-not [bool]$groups[0].securityEnabled) {
            throw "Existing group is not security-enabled: $($groups[0].id)"
        }
        Write-Step "Reusing pilot group: $($groups[0].displayName) ($($groups[0].id))"
        return $groups[0]
    }
    if (-not $CreateWhenMissing) {
        return $null
    }

    $groupBody = [ordered]@{
        displayName     = $DisplayName
        description     = $GroupDescription
        mailEnabled     = $false
        mailNickname    = $MailNickname
        securityEnabled = $true
        groupTypes      = @()
    }
    $group = Invoke-GraphJson -Method POST -Uri 'https://graph.microsoft.com/v1.0/groups' -Body $groupBody
    Write-Step "Created pilot group: $($group.displayName) ($($group.id))"
    return $group
}

function Add-AppSupersedence {
    param(
        [Parameter(Mandatory = $true)][string]$NewAppId,
        [Parameter(Mandatory = $true)][string]$SupersededAppId,
        [Parameter(Mandatory = $true)][ValidateSet('update', 'replace')][string]$Type
    )

    $relationshipsUri = "$GraphBaseUri/deviceAppManagement/mobileAppRelationships"
    $existingRelationships = @(Get-GraphCollection -Uri $relationshipsUri)
    $currentRelationships = @($existingRelationships | Where-Object {
        [string]$_.sourceId -eq $NewAppId
    })
    $existing = @($currentRelationships | Where-Object {
        [string]$_.targetId -eq $SupersededAppId -and
        [string]$_.supersedenceType -eq $Type
    })
    if ($existing.Count -gt 0) {
        Write-Step "Supersedence already exists: $NewAppId -> $SupersededAppId ($Type)"
        return $existing[0]
    }

    $relationshipSpecs = New-Object Collections.Generic.List[object]
    foreach ($currentRelationship in $currentRelationships) {
        if ([string]$currentRelationship.targetId -eq $SupersededAppId) {
            continue
        }
        $odataType = [string]$currentRelationship.'@odata.type'
        if ([string]::IsNullOrWhiteSpace($odataType) -or
            [string]::IsNullOrWhiteSpace([string]$currentRelationship.targetId)) {
            throw 'An existing Intune app relationship could not be safely preserved.'
        }
        $preservedRelationship = [ordered]@{
            '@odata.type' = $odataType
            targetId      = [string]$currentRelationship.targetId
        }
        if ($currentRelationship.PSObject.Properties['supersedenceType'] -and
            -not [string]::IsNullOrWhiteSpace([string]$currentRelationship.supersedenceType)) {
            $preservedRelationship.supersedenceType = [string]$currentRelationship.supersedenceType
        }
        elseif ($currentRelationship.PSObject.Properties['dependencyType'] -and
            -not [string]::IsNullOrWhiteSpace([string]$currentRelationship.dependencyType)) {
            $preservedRelationship.dependencyType = [string]$currentRelationship.dependencyType
        }
        else {
            throw "Unsupported existing Intune app relationship type: $odataType"
        }
        $relationshipSpecs.Add($preservedRelationship)
    }

    $newRelationship = [ordered]@{
        '@odata.type'    = '#microsoft.graph.mobileAppSupersedence'
        targetId         = $SupersededAppId
        supersedenceType = $Type
    }
    $relationshipSpecs.Add($newRelationship)
    $updateBody = [ordered]@{
        relationships = @($relationshipSpecs.ToArray())
    }
    $updateUri = "$GraphBaseUri/deviceAppManagement/mobileApps/$NewAppId/updateRelationships"
    Invoke-GraphJson -Method POST -Uri $updateUri -Body $updateBody | Out-Null
    Write-Step "Created Intune supersedence: $NewAppId -> $SupersededAppId ($Type)"
    return [pscustomobject]$newRelationship
}

function Add-PilotAssignment {
    param(
        [Parameter(Mandatory = $true)][string]$AppId,
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][ValidateSet('available', 'required')][string]$Intent
    )

    $assignmentsUri = "$GraphBaseUri/deviceAppManagement/mobileApps/$AppId/assignments"
    $existingAssignments = @(Get-GraphCollection -Uri $assignmentsUri)
    $existingPilotAssignment = @($existingAssignments | Where-Object {
        $_.target -and [string]$_.target.groupId -eq $GroupId -and [string]$_.intent -eq $Intent
    })
    if ($existingPilotAssignment.Count -gt 0) {
        Write-Step ((Get-Culture).TextInfo.ToTitleCase($Intent) + ' pilot assignment already exists.')
        return
    }
    $assignmentBody = [ordered]@{
        '@odata.type' = '#microsoft.graph.mobileAppAssignment'
        intent        = $Intent
        target        = [ordered]@{
            '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
            groupId        = $GroupId
        }
        settings      = [ordered]@{
            '@odata.type'                = '#microsoft.graph.win32LobAppAssignmentSettings'
            notifications                = 'showAll'
            deliveryOptimizationPriority = 'notConfigured'
        }
    }
    Invoke-GraphJson -Method POST -Uri $assignmentsUri -Body $assignmentBody | Out-Null
    Write-Step ("Assigned app as {0} to pilot group: {1}" -f $Intent,$GroupId)
}
$resolvedIntuneWinPath = (Resolve-Path -LiteralPath $IntuneWinPath -ErrorAction Stop).ProviderPath
if ([IO.Path]::GetExtension($resolvedIntuneWinPath) -ne '.intunewin') {
    throw "Expected a .intunewin package: $resolvedIntuneWinPath"
}
$resolvedDetectionScriptPath = (Resolve-Path -LiteralPath $DetectionScriptPath -ErrorAction Stop).ProviderPath

$versionManifestCandidates = @(
    (Join-Path $PSScriptRoot 'SmartM365-EndpointDiagnosticsAnalyzer.version.json')
    (Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartM365-EndpointDiagnosticsAnalyzer.version.json')
)
$versionManifestPath = $versionManifestCandidates |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($PackageVersion)) {
    if ([string]::IsNullOrWhiteSpace($versionManifestPath)) {
        throw 'PackageVersion was not supplied and the local version manifest was not found.'
    }
    $versionManifest = Get-Content -LiteralPath $versionManifestPath -Raw | ConvertFrom-Json
    $PackageVersion = [string]$versionManifest.PackageVersion
}
if ([string]::IsNullOrWhiteSpace($PackageVersion)) {
    throw 'PackageVersion cannot be empty.'
}
if ([string]::IsNullOrWhiteSpace($AppDisplayName)) {
    $AppDisplayName = "SmartM365 Endpoint Diagnostics Analyzer $PackageVersion"
}
if ($AssignPilotGroup -and -not $CreatePilotGroup -and [string]::IsNullOrWhiteSpace($PilotGroupId)) {
    throw '-AssignPilotGroup requires -CreatePilotGroup or -PilotGroupId.'
}
if (-not [string]::IsNullOrWhiteSpace($SupersedeAppId)) {
    $supersededGuid = [guid]::Empty
    if (-not [guid]::TryParse($SupersedeAppId, [ref]$supersededGuid)) {
        throw "-SupersedeAppId must be a valid GUID: $SupersedeAppId"
    }
}
if (-not [string]::IsNullOrWhiteSpace($ResumeAppId)) {
    $resumeGuid = [guid]::Empty
    if (-not [guid]::TryParse($ResumeAppId, [ref]$resumeGuid)) {
        throw "-ResumeAppId must be a valid GUID: $ResumeAppId"
    }
    if ($ResumeAppId -eq $SupersedeAppId) {
        throw '-ResumeAppId and -SupersedeAppId must identify different Intune apps.'
    }
}

$detectionScriptContent = Get-Content -LiteralPath $resolvedDetectionScriptPath -Raw
if ([string]::IsNullOrWhiteSpace($detectionScriptContent)) {
    throw "Detection script is empty: $resolvedDetectionScriptPath"
}
$expectedVersionPattern = [regex]::Escape($PackageVersion)
if ($detectionScriptContent -notmatch "\`$ExpectedVersion\s*=\s*'$expectedVersionPattern'") {
    throw "Detection script does not contain the expected package version: $PackageVersion"
}
$detectionSignature = Get-AuthenticodeSignature -LiteralPath $resolvedDetectionScriptPath
$actualSigner = if ($detectionSignature.SignerCertificate) {
    Get-NormalizedThumbprint $detectionSignature.SignerCertificate.Thumbprint
}
else {
    ''
}
if ($actualSigner -ne (Get-NormalizedThumbprint $ExpectedSignerThumbprint) -or
    $detectionSignature.Status -in @('NotSigned','HashMismatch')) {
    throw ("Detection script signer validation failed: status={0}; signer={1}" -f
        $detectionSignature.Status,$actualSigner)
}

$metadata = Read-IntuneWinPackage -Path $resolvedIntuneWinPath
$packageHash = (Get-FileHash -LiteralPath $resolvedIntuneWinPath -Algorithm SHA256).Hash
$preview = [pscustomobject]@{
    Mode                       = if ($Execute) { 'Execute' } else { 'Preview' }
    AppDisplayName             = $AppDisplayName
    PackageVersion             = $PackageVersion
    IntuneWinPath              = $resolvedIntuneWinPath
    IntuneWinSha256            = $packageHash
    SetupFile                  = $metadata.SetupFile
    DetectionScriptPath        = $resolvedDetectionScriptPath
    DetectionSignerThumbprint  = $actualSigner
    RequiredDelegatedScopes    = $requiredScopes
    InteractiveAuthentication  = $true
    PilotGroupDisplayName      = $PilotGroupDisplayName
    PilotGroupId               = $PilotGroupId
    PilotGroupCreationRequested = [bool]$CreatePilotGroup
    PilotAssignmentRequested   = [bool]$AssignPilotGroup
    PilotAssignmentIntent      = if ($AssignPilotGroup) { $PilotAssignmentIntent } else { 'none' }
    SupersedeAppId             = $SupersedeAppId
    SupersedenceRequested      = -not [string]::IsNullOrWhiteSpace($SupersedeAppId)
    SupersedenceType           = if ([string]::IsNullOrWhiteSpace($SupersedeAppId)) { 'none' } else { $SupersedenceType }
    ResumeAppId                = $ResumeAppId
    ResumeRequested            = -not [string]::IsNullOrWhiteSpace($ResumeAppId)
    ChangesAttempted           = $false
}

if (-not $Execute) {
    return $preview
}
if (-not $PSCmdlet.ShouldProcess(
        $AppDisplayName,
        'Publish Intune Win32 app, supersede a prior version, and optionally create/assign the pilot group'
    )) {
    return $preview
}
$preview.ChangesAttempted = $true

Assert-GraphCommand
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
catch {
    Write-Verbose 'No existing Microsoft Graph session required disconnection.'
}

$connectParameters = @{
    Scopes      = $requiredScopes
    NoWelcome   = $true
    ErrorAction = 'Stop'
}
if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
    $connectParameters.TenantId = $TenantId
}
Write-Step 'Connecting interactively to Microsoft Graph.'
Connect-MgGraph @connectParameters | Out-Null
$graphContext = Get-MgContext
if (-not $graphContext) {
    throw 'Microsoft Graph did not return an authenticated context.'
}
Write-Step ("Connected to Microsoft Graph. TenantId={0}; Account={1}" -f
    $graphContext.TenantId,$graphContext.Account)

$supersededApp = $null
if (-not [string]::IsNullOrWhiteSpace($SupersedeAppId)) {
    $supersededApp = Invoke-GraphJson `
        -Method GET `
        -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/${SupersedeAppId}?`$select=id,displayName,publisher,notes"
    if (-not $supersededApp -or [string]$supersededApp.id -ne $SupersedeAppId) {
        throw "Superseded Intune app could not be validated: $SupersedeAppId"
    }
    $matchesProductName = [string]$supersededApp.displayName -like '*SmartM365 Endpoint Diagnostics Analyzer*'
    $matchesPreviewNotes = [string]$supersededApp.notes -like '*PackageVersion=*'
    if ([string]$supersededApp.publisher -ne 'WorkplaceCloudHub' -or
        (-not $matchesProductName -and -not $matchesPreviewNotes)) {
        throw ("Superseded app is not a WorkplaceCloudHub Endpoint Diagnostics Analyzer app. Id={0}; DisplayName={1}; Publisher={2}; Notes={3}" -f
            $supersededApp.id,$supersededApp.displayName,$supersededApp.publisher,$supersededApp.notes)
    }
    Write-Step ("Validated superseded Intune app: {0} ({1})" -f
        $supersededApp.displayName,$supersededApp.id)
}

$appId = ''
$contentVersionId = ''
$pilotGroup = $null
$supersedence = $null
try {
    if (-not [string]::IsNullOrWhiteSpace($ResumeAppId)) {
        $resumeApp = Invoke-GraphJson `
            -Method GET `
            -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/${ResumeAppId}"
        $expectedVersionNote = "PackageVersion=$PackageVersion"
        if (-not $resumeApp -or [string]$resumeApp.id -ne $ResumeAppId) {
            throw "Existing Intune app could not be validated for resume: $ResumeAppId"
        }
        if ([string]$resumeApp.displayName -ne $AppDisplayName -or
            [string]$resumeApp.publisher -ne $Publisher -or
            [string]$resumeApp.notes -notlike "*$expectedVersionNote*") {
            throw ("Existing Intune app does not match this publication. Id={0}; DisplayName={1}; Publisher={2}; Notes={3}" -f
                $resumeApp.id,$resumeApp.displayName,$resumeApp.publisher,$resumeApp.notes)
        }
        $appId = [string]$resumeApp.id
        $contentVersionId = [string]$resumeApp.committedContentVersion
        if ([string]::IsNullOrWhiteSpace($contentVersionId)) {
            throw "Existing Intune app has no committed content version and cannot be resumed: $ResumeAppId"
        }
        Write-Step ("Resuming existing Intune app: {0} ({1}); committed content version={2}" -f
            $resumeApp.displayName,$appId,$contentVersionId)
    }
    else {
    $existingApps = @(Get-ExistingAppByDisplayName -DisplayName $AppDisplayName)
    if ($existingApps.Count -gt 0 -and -not $AllowDuplicateDisplayName) {
        $existingText = $existingApps |
            ForEach-Object { '{0} ({1})' -f $_.displayName,$_.id }
        throw ("An Intune app already uses this display name. Choose another version/name or explicitly pass -AllowDuplicateDisplayName. Existing={0}" -f
            ($existingText -join ', '))
    }

    $appBody = [ordered]@{
        '@odata.type' = '#microsoft.graph.win32LobApp'
        displayName = $AppDisplayName
        description = $Description
        publisher = $Publisher
        developer = $Developer
        owner = $Owner
        notes = "PackageVersion=$PackageVersion; UpdateChannel=Intune; GalleryAutomaticUpdate=False"
        isFeatured = $false
        informationUrl = 'https://workplacecloudhub.com'
        privacyInformationUrl = 'https://workplacecloudhub.com/privacy/'
        fileName = [string]$metadata.FileName
        setupFilePath = [string]$metadata.SetupFile
        installCommandLine = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1 -InstallScope AllUsers -PackageSource Intune'
        uninstallCommandLine = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Uninstall.ps1 -InstallScope AllUsers -Confirm:$false'
        installExperience = [ordered]@{
            '@odata.type' = '#microsoft.graph.win32LobAppInstallExperience'
            runAsAccount = 'system'
            deviceRestartBehavior = 'suppress'
        }
        minimumSupportedOperatingSystem = [ordered]@{
            '@odata.type' = '#microsoft.graph.windowsMinimumOperatingSystem'
            v10_1607 = $true
        }
        applicableArchitectures = 'x64'
        requirementRules = @()
        detectionRules = @(
            [ordered]@{
                '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptDetection'
                enforceSignatureCheck = $false
                runAs32Bit = $false
                scriptContent = ConvertTo-Base64Utf8 -Value $detectionScriptContent
            }
        )
        returnCodes = @(
            @{ returnCode = 0; type = 'success' }
            @{ returnCode = 1707; type = 'success' }
            @{ returnCode = 3010; type = 'softReboot' }
            @{ returnCode = 1641; type = 'hardReboot' }
            @{ returnCode = 1618; type = 'retry' }
        )
        installCommandLineTimeoutInMinutes = 15
    }

    $app = Invoke-GraphJson `
        -Method POST `
        -Uri "$GraphBaseUri/deviceAppManagement/mobileApps" `
        -Body $appBody
    $appId = [string]$app.id
    if ([string]::IsNullOrWhiteSpace($appId)) {
        throw 'Microsoft Graph did not return a mobile app id.'
    }
    Write-Step "Created Intune app: $appId"

    $contentVersion = Invoke-GraphJson `
        -Method POST `
        -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions" `
        -Body @{}
    $contentVersionId = [string]$contentVersion.id
    if ([string]::IsNullOrWhiteSpace($contentVersionId)) {
        throw 'Microsoft Graph did not return a content version id.'
    }
    Write-Step "Created content version: $contentVersionId"

    $fileBody = [ordered]@{
        '@odata.type' = '#microsoft.graph.mobileAppContentFile'
        name = [string]$metadata.FileName
        size = [int64]$metadata.UnencryptedContentSize
        sizeEncrypted = [int64]$metadata.SizeEncrypted
        manifest = $null
        isDependency = $false
    }
    $file = Invoke-GraphJson `
        -Method POST `
        -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions/$contentVersionId/files" `
        -Body $fileBody
    $fileId = [string]$file.id
    if ([string]::IsNullOrWhiteSpace($fileId)) {
        throw 'Microsoft Graph did not return a content file id.'
    }
    $fileUri = "$GraphBaseUri/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions/$contentVersionId/files/$fileId"

    $file = Wait-GraphContentFileState `
        -Uri $fileUri `
        -SuccessStates @('azureStorageUriRequestSuccess','azureStorageUriRenewalSuccess') `
        -FailureStates @('azureStorageUriRequestFailed','azureStorageUriRenewalFailed') `
        -PollIntervalSeconds $PollSeconds `
        -TimeoutMinutes $PollTimeoutMinutes `
        -RequireAzureStorageUri

    Write-Step 'Uploading encrypted package content to the Intune staging blob.'
    Send-AzureBlockBlobFromIntuneWin `
        -Path $resolvedIntuneWinPath `
        -ContentEntryName ([string]$metadata.ContentEntryName) `
        -AzureStorageUri ([string]$file.azureStorageUri) `
        -BlockSizeMB $UploadBlockSizeMB `
        -MaxRetries $AzureUploadMaxRetries

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
    Wait-GraphContentFileState `
        -Uri $fileUri `
        -SuccessStates @('commitFileSuccess') `
        -FailureStates @('commitFileFailed') `
        -PollIntervalSeconds $PollSeconds `
        -TimeoutMinutes $PollTimeoutMinutes | Out-Null

    $finalAppBody = Copy-OrderedHashtable -InputObject $appBody
    $finalAppBody.Remove('applicableArchitectures')
    $finalAppBody['committedContentVersion'] = $contentVersionId
    Invoke-GraphJson `
        -Method PATCH `
        -Uri "$GraphBaseUri/deviceAppManagement/mobileApps/$appId" `
        -Body $finalAppBody | Out-Null
    Write-Step 'App metadata and content version committed.'
    }

    if (-not [string]::IsNullOrWhiteSpace($SupersedeAppId)) {
        $supersedence = Add-AppSupersedence `
            -NewAppId $appId `
            -SupersededAppId $SupersedeAppId `
            -Type $SupersedenceType
    }

    if ($CreatePilotGroup -or -not [string]::IsNullOrWhiteSpace($PilotGroupId)) {
        $pilotGroup = Resolve-PilotGroup `
            -GroupId $PilotGroupId `
            -DisplayName $PilotGroupDisplayName `
            -MailNickname $PilotGroupMailNickname `
            -GroupDescription $PilotGroupDescription `
            -CreateWhenMissing:$CreatePilotGroup
    }
    if ($AssignPilotGroup) {
        if (-not $pilotGroup -or [string]::IsNullOrWhiteSpace([string]$pilotGroup.id)) {
            throw 'Pilot group could not be resolved for assignment.'
        }
        Add-PilotAssignment -AppId $appId -GroupId ([string]$pilotGroup.id) -Intent $PilotAssignmentIntent
    }

    [pscustomobject]@{
        Result                  = 'PASS'
        AppId                   = $appId
        AppDisplayName          = $AppDisplayName
        PackageVersion          = $PackageVersion
        ContentVersionId        = $contentVersionId
        TenantId                = [string]$graphContext.TenantId
        PilotGroupId            = if ($pilotGroup) { [string]$pilotGroup.id } else { '' }
        PilotGroupDisplayName   = if ($pilotGroup) { [string]$pilotGroup.displayName } else { $PilotGroupDisplayName }
        PilotAssignmentCreated  = [bool]$AssignPilotGroup
        PilotAssignmentIntent   = if ($AssignPilotGroup) { $PilotAssignmentIntent } else { 'none' }
        SupersededAppId         = $SupersedeAppId
        SupersedenceCreated     = $null -ne $supersedence
        SupersedenceType        = if ($null -ne $supersedence) { $SupersedenceType } else { 'none' }
        GalleryAutomaticUpdate  = $false
    }
}
catch {
    if (-not [string]::IsNullOrWhiteSpace($appId)) {
        Write-Warning ("Publication failed after the Intune app was created. Review incomplete app id: {0}" -f $appId)
    }
    throw
}
finally {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-Verbose 'Microsoft Graph disconnect failed during cleanup.'
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBAFV+uv905ikDk
# BuLW5cekqdqNC6Tn5p3OyV38pPY+hKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIFgvdl6B5+Cll3lk0zw1ByWoip8YeOk46rdI77ke+NRMA0GCSqG
# SIb3DQEBAQUABIIBgEQVMSyRNyTmdCv9wWp/eJYIFOLaPslltf8DSIJFmSh9iNz+
# HqCXMETS2W+hhjYmgR5PdBKYmQVPaXCZO5oB3g5XiwhY7R+7yQp+Db1q6N0XyudG
# 0dhyWq6/xv57SWhSBKr2tqI5jyPowdjCuuXE2q+faT9brfamx5TOcza0fCZrFxtO
# q9LDWo7caboZ/L05oa93NACLn/1dzRN92Q8H1tl7SaRCHWYtYfEpEiLO95trnpmD
# 6oRKd6LvCBXL7RZJohr/7b7xkZ2TqdEOoNDM3465hahIcaZshrqh9MXtVbP+lyLc
# 7ctuq46XUMoR0O679jgH9a5wRBSBCv3ozNzmMtFTL/MH4ZfIQ1xZzR6lQ7YAl9Pc
# 3gMR8kWscWxeCinYSby/9gtMBp3SCPqL2SaC9D9diViho4bBCYd+6K0kcof3GYyL
# IEN5XLYZmk9w8mJ6aai3yof+HvdK3hBnWQJTjxyS6nBWOS5XQWSjtYpnZfQi4Xej
# v9BzURQbODUcFm+IN6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA4MTIxMDUy
# NThaMC8GCSqGSIb3DQEJBDEiBCANi/g/+Y8hxX5cqJjUZpD6Mld5tKpFuwHtYC2K
# dcpWYDANBgkqhkiG9w0BAQEFAASCAgCWo0eU2OHpAFBTonM8u6RsVzJmf497r3eJ
# 3hUTr9ZfWF1GYzZttal6FgrndMv72Ohu6Cmho/JfY1TJgv0t54ItG1hajo3sAqY1
# 3I3LU9FBURpPj82hVgcGu/zGKdUDxS+G0AxUkuOwv2mD7RAo9nJ+4wuNC08Wlrfe
# I9CFVDhlpE0KhEs5h/oWXupR8SJlFuQ3Dv5yx98jePrZ1tgrCzE1h4320v/d6qYe
# lqxuG6mm3nIgxMQFnp49fVeHq4KepKMYz/iohyDyge6/dgmLtvAkjA55CMcIDJ/r
# qcoWlSn318OVeW7Kp9CKN83cxwD0Jg1+j+SDUgHcqT7OdNCKgdtAw1eRYsOIkc9E
# /xsJ22GauPfEdSULOEA52dr6q4B6lbz+bGKL7BAmJgXda/Xy1OQnZVGeJ4QYGo7I
# 2137xmDYil8/rZvm10KRVGXTM05JWIYFxCUxCSo6Gjs9ZFN2lOtvv1gg7oXryc+d
# A4kHb2RNk1Kyazcf4radhNbivk7AwgVbD1/nfmjJ0OA1meNHE4DF3bYZxbQBbxLx
# SVoROMCKRfIP9egFhTXeBWNvGAObP8/pRB1RTxay6rpvJvOSp7xhGMpfuZ4nHe2b
# flRRLnoZ5GWKNcGr+JqFQony5SI1lW8D90aHzuFvcbc37+a1S+3qRA4XsTfdIqxB
# UyE+zppokw==
# SIG # End signature block
