Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.2' -Global -ErrorAction Stop

$script:DistributedModuleVersion = '1.1.10'

function Resolve-DistributedJsonPath {
    param([Parameter(Mandatory)][string]$Path)
    $names = Get-SmartM365JsonNames $Path
    $leaf = [IO.Path]::GetFileName($names.Legacy)
    $parent = [IO.Path]::GetFileName([IO.Path]::GetDirectoryName($names.Legacy))
    $existing = Get-SmartM365JsonReadPath -Path $Path -Optional
    $isLease = $false
    if ($existing) { $isLease = $null -ne (Read-SmartM365JsonDocument $Path).Document.PSObject.Properties['LeaseId'] }
    $owner = if ($isLease) { 'Orchestrator concurrency lease' } else { 'Orchestrator distributed state' }
    $compatibleOwners = if ($isLease) { @('Orchestrator distributed state') } else { @() }
    Resolve-SmartM365OwnedJsonPath -Path $Path -Owner $owner -CompatibleJournalOwners $compatibleOwners -Validate {
        param($document)
        if (-not $document.PSObject.Properties['JobName'] -or -not $document.PSObject.Properties['OwnerServer'] -or
            (-not $document.PSObject.Properties['ClaimId'] -and -not $document.PSObject.Properties['LeaseId'])) { throw 'Distributed state owner/schema mismatch.' }
        if ($document.PSObject.Properties['ClaimId']) {
            $expectedLeaf = ([datetime]$document.OccurrenceUtc).ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ') + '.json'
            if ($parent -ne (ConvertTo-SafeFileName $document.JobName) -or ($leaf -ne $expectedLeaf -and $leaf -ne ($expectedLeaf + '.stale.' + $document.ClaimId + '.json'))) { throw 'Distributed claim filename/identity mismatch.' }
        } else {
            $expectedLeaf = (ConvertTo-SafeFileName $document.ConcurrencyKey) + '.json'
            if ($leaf -ne $expectedLeaf -and $leaf -ne ($expectedLeaf + '.stale.' + $document.LeaseId + '.json')) { throw 'Distributed lease filename/identity mismatch.' }
        }
    }
}

function Convert-SmartM365OrchestratorDistributedHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ClaimsRootPath,[Parameter(Mandatory)][string]$LeasesRootPath,[scriptblock]$OnProgress)
    if ((Get-SmartM365JsonTransportPolicy).Mode -ne 'JsonText') { return }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $lastReport = 0.0
    $checked = 0; $skipped = 0; $resolved = 0
    if ($OnProgress) { & $OnProgress 'Distributed JSON history scan started.' | Out-Null }
    $entries = @{}
    $claimPaths = [Collections.Generic.List[string]]::new()
    $paths = @()
    if (Test-Path -LiteralPath $ClaimsRootPath) {
        foreach ($folder in @(Get-ChildItem -LiteralPath $ClaimsRootPath -Directory -ErrorAction Stop)) {
            if ($folder.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Distributed history folder is a reparse point.' }
            foreach ($file in @(Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction Stop)) {
                $entries[$file.FullName] = $file
                if ($file.Name -match '^\d{8}T\d{9}Z\.json(?:\.stale\.[0-9a-f]{32}\.json)?(?:\.txt)?$') { $claimPaths.Add($file.FullName) }
            }
            if ($OnProgress -and ($timer.Elapsed.TotalSeconds - $lastReport) -ge 5) {
                & $OnProgress ("Distributed JSON history listing: claims={0}; folder={1}." -f $claimPaths.Count,$folder.Name) | Out-Null
                $lastReport = $timer.Elapsed.TotalSeconds
            }
        }
    }
    foreach ($path in $claimPaths) {
        $checked++
        $completed = $false
        if ($path.EndsWith('.json.txt', [StringComparison]::OrdinalIgnoreCase)) {
            $legacy = $path.Substring(0, $path.Length - 4)
            $journal = $legacy + '.migration.log'
            # Only historical claims qualify. Live leases always use the locked resolver.
            # Readers still validate the current JSON whenever a claim is consumed.
            if (-not $entries.ContainsKey($legacy) -and -not ($entries[$path].Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                if (-not $entries.ContainsKey($journal)) { $completed = $true }
                elseif (-not ($entries[$journal].Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                    # One sequential receipt read replaces repeated JSON reads, hashes,
                    # policy loads and lock operations for an already completed migration.
                    $lastLine = $null
                    foreach ($line in [IO.File]::ReadLines($journal)) {
                        if (-not [string]::IsNullOrWhiteSpace($line)) { $lastLine = $line }
                    }
                    try {
                        $receipt = ConvertFrom-Json -InputObject $lastLine -ErrorAction Stop
                        $completed = $receipt.Phase -eq 'Completed' -and $receipt.Owner -eq 'Orchestrator distributed state'
                    } catch { $completed = $false }
                }
            }
        }
        if ($completed) { $skipped++ } else { $paths += $path }
        if ($OnProgress -and ($timer.Elapsed.TotalSeconds - $lastReport) -ge 5) {
            & $OnProgress ("Distributed JSON history scan: checked={0}/{1}; alreadyCurrent={2}; pending={3}." -f $checked,$claimPaths.Count,$skipped,$paths.Count) | Out-Null
            $lastReport = $timer.Elapsed.TotalSeconds
        }
    }
    if (Test-Path -LiteralPath $LeasesRootPath) {
        $paths += @(Get-ChildItem -LiteralPath $LeasesRootPath -File -ErrorAction Stop | Where-Object { $_.Name -match '\.json(?:\.txt)?$' } | ForEach-Object { $_.FullName })
    }
    foreach ($path in @($paths | ForEach-Object { (Get-SmartM365JsonNames $_).Legacy } | Sort-Object -Unique)) {
        $null = Resolve-DistributedJsonPath $path
        $resolved++
        if ($OnProgress -and ($timer.Elapsed.TotalSeconds - $lastReport) -ge 5) {
            & $OnProgress ("Distributed JSON history recovery: resolved={0}; alreadyCurrent={1}; current={2}." -f $resolved,$skipped,[IO.Path]::GetFileName($path)) | Out-Null
            $lastReport = $timer.Elapsed.TotalSeconds
        }
    }
    if ($OnProgress) { & $OnProgress ("Distributed JSON history scan complete: checked={0}; alreadyCurrent={1}; resolved={2}; seconds={3:N1}." -f $checked,$skipped,$resolved,$timer.Elapsed.TotalSeconds) | Out-Null }
}

function ConvertTo-SafeFileName {
    param([Parameter(Mandatory = $true)][string]$Value)

    $safe = $Value
    foreach ($character in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$character, '_')
    }
    return $safe
}

function Write-JsonAtomically {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Value,
        [int]$Depth = 12,
        [string]$ExpectedSHA256 = ''
    )

    $parent = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop)
    }

    $Path = Resolve-DistributedJsonPath $Path
    $json = $Value | ConvertTo-Json -Depth $Depth
    $null = Write-SmartM365JsonBytesAtomically -Path $Path -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($json)) -ExpectedSHA256 $ExpectedSHA256 -Validate {
        param($document) if ($document -isnot [pscustomobject]) { throw 'Distributed state must be an object.' }
    }
}

function Get-CertificateReadiness {
    param([string]$Thumbprint)

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
        return [pscustomobject]@{ Ready = $false; Detail = 'Certificate thumbprint is not configured.' }
    }

    $normalized = ($Thumbprint -replace '\s', '').ToUpperInvariant()
    foreach ($storePath in @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')) {
        $certificate = Get-ChildItem -LiteralPath $storePath -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $normalized -and $_.NotAfter -gt (Get-Date) } |
            Select-Object -First 1
        if ($null -ne $certificate) {
            return [pscustomobject]@{
                Ready = $true
                Detail = "Usable certificate found in $storePath; expires $($certificate.NotAfter.ToString('o'))."
            }
        }
    }

    return [pscustomobject]@{ Ready = $false; Detail = "Certificate $normalized was not found or is expired." }
}

function Invoke-IsolatedCapabilityProbe {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$EnginePath,
        [Parameter(Mandatory = $true)][string]$ScriptText,
        [hashtable]$Environment = @{},
        [int]$TimeoutSeconds = 90
    )

    if (-not (Test-Path -LiteralPath $EnginePath -PathType Leaf)) {
        return [pscustomobject]@{ Name = $Name; Ready = $false; Roles = @(); Detail = "PowerShell engine not found: $EnginePath" }
    }

    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptText))
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $EnginePath
    $startInfo.Arguments = "-NoLogo -NoProfile -NonInteractive -EncodedCommand $encoded"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($key in $Environment.Keys) {
        $startInfo.Environment[[string]$key] = [string]$Environment[$key]
    }

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw "Could not start $EnginePath."
        }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit([math]::Max(1, $TimeoutSeconds) * 1000)) {
            try { $process.Kill($true) } catch { Write-Verbose ("Probe process kill failed: {0}" -f $_.Exception.Message) }
            return [pscustomobject]@{ Name = $Name; Ready = $false; Roles = @(); Detail = "Read-only probe timed out after $TimeoutSeconds seconds." }
        }
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        $match = [regex]::Match($stdout, '(?m)^SMARTM365_CAPABILITY_JSON:(?<Json>.+)$')
        if (-not $match.Success) {
            $detail = (($stderr + "`n" + $stdout).Trim() -replace '[\r\n]+', ' ')
            if ($detail.Length -gt 600) { $detail = $detail.Substring(0, 600) }
            if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "Probe exited with code $($process.ExitCode) without a result." }
            return [pscustomobject]@{ Name = $Name; Ready = $false; Roles = @(); Detail = $detail }
        }
        $result = $match.Groups['Json'].Value | ConvertFrom-Json -ErrorAction Stop
        return [pscustomobject]@{
            Name = $Name
            Ready = [bool]$result.Ready
            Roles = @($result.Roles | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
            Detail = [string]$result.Detail
        }
    }
    catch {
        return [pscustomobject]@{ Name = $Name; Ready = $false; Roles = @(); Detail = $_.Exception.Message }
    }
    finally {
        $process.Dispose()
    }
}

function Get-CapabilityProbeResult {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [bool]$Ready,
        [string]$Detail,
        [string[]]$Roles = @()
    )

    return [pscustomobject]@{
        Name = $Name
        Ready = $Ready
        Roles = @($Roles)
        Detail = $Detail
    }
}

function Get-SmartM365OrchestratorServerCapability {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ServerName,
        [Parameter(Mandatory = $true)][string]$SharedDataFolderPath,
        [string]$TenantId,
        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$Organization,
        [ValidateSet('Static', 'ReadOnly')][string]$ProbeMode = 'ReadOnly',
        [int]$ProbeTimeoutSeconds = 90
    )

    $generatedAtUtc = [datetime]::UtcNow
    $certificate = Get-CertificateReadiness -Thumbprint $CertificateThumbprint
    $results = [System.Collections.Generic.List[object]]::new()

    $probeFolder = Join-Path -Path $SharedDataFolderPath -ChildPath 'Election\Probe'
    try {
        [void](New-Item -ItemType Directory -Path $probeFolder -Force)
        $probePath = Join-Path -Path $probeFolder -ChildPath ('{0}.{1}.{2}.tmp' -f (ConvertTo-SafeFileName $ServerName), $PID, [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probePath, $generatedAtUtc.ToString('o'), [System.Text.UTF8Encoding]::new($false))
        Remove-Item -LiteralPath $probePath -Force -ErrorAction Stop
        $results.Add((Get-CapabilityProbeResult -Name 'SharedRuntime' -Ready $true -Detail 'Shared election storage is writable.'))
    }
    catch {
        $results.Add((Get-CapabilityProbeResult -Name 'SharedRuntime' -Ready $false -Detail $_.Exception.Message))
    }

    $cloudConfigurationReady = $certificate.Ready -and
        -not [string]::IsNullOrWhiteSpace($TenantId) -and
        -not [string]::IsNullOrWhiteSpace($AppId)
    $pwshPath = (Get-Process -Id $PID).Path
    $windowsPowerShellPath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'

    if ($ProbeMode -eq 'Static') {
        $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'Graph' -Ready ($cloudConfigurationReady -and $null -ne $graphModule) -Detail ("Static check: module={0}; {1}" -f $(if ($graphModule) { $graphModule.Version } else { 'missing' }), $certificate.Detail)))
        $exoModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'EXO' -Ready ($cloudConfigurationReady -and $null -ne $exoModule -and -not [string]::IsNullOrWhiteSpace($Organization)) -Detail ("Static check: module={0}; organizationConfigured={1}; {2}" -f $(if ($exoModule) { $exoModule.Version } else { 'missing' }), (-not [string]::IsNullOrWhiteSpace($Organization)), $certificate.Detail)))
        $adModule = Get-Module -ListAvailable -Name ActiveDirectory | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'AD' -Ready ($null -ne $adModule) -Detail ("Static check: ActiveDirectory module {0}." -f $(if ($adModule) { 'available' } else { 'missing' }))))
        $exchangeSnapIn = & $windowsPowerShellPath -NoLogo -NoProfile -NonInteractive -Command "if (Get-PSSnapin -Registered -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue) { 'yes' } else { 'no' }" 2>$null
        $results.Add((Get-CapabilityProbeResult -Name 'ExchangeOnPrem' -Ready (($exchangeSnapIn | Select-Object -Last 1) -eq 'yes') -Detail 'Static check of the Exchange Management snap-in registration.'))
        $teamsModule = Get-Module -ListAvailable -Name MicrosoftTeams | Sort-Object Version -Descending | Select-Object -First 1
        $results.Add((Get-CapabilityProbeResult -Name 'TeamsPowerShell' -Ready ($cloudConfigurationReady -and $null -ne $teamsModule) -Detail ("Static check: module={0}; {1}" -f $(if ($teamsModule) { $teamsModule.Version } else { 'missing' }), $certificate.Detail)))
    }
    else {
        $environment = @{
            SMART_TENANT_ID = $TenantId
            SMART_APP_ID = $AppId
            SMART_CERT_THUMBPRINT = $CertificateThumbprint
            SMART_ORGANIZATION = $Organization
        }
        $graphScript = @'
$ErrorActionPreference = 'Stop'
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -TenantId $env:SMART_TENANT_ID -ClientId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -NoWelcome
    $roles = @((Get-MgContext).Scopes | Sort-Object -Unique)
    $result = @{ Ready = $true; Roles = $roles; Detail = "Read-only Graph authentication succeeded with $($roles.Count) application roles." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
finally { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'Graph' -EnginePath $pwshPath -ScriptText $graphScript -Environment $environment -TimeoutSeconds $ProbeTimeoutSeconds))

        $exoScript = @'
$ErrorActionPreference = 'Stop'
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -AppId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -Organization $env:SMART_ORGANIZATION -ShowBanner:$false
    $null = Get-AcceptedDomain -ResultSize 1 -ErrorAction Stop
    $result = @{ Ready = $true; Roles = @(); Detail = 'Read-only Exchange Online probe succeeded.' }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
finally { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'EXO' -EnginePath $pwshPath -ScriptText $exoScript -Environment $environment -TimeoutSeconds $ProbeTimeoutSeconds))

        $adScript = @'
$ErrorActionPreference = 'Stop'
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $rootDse = Get-ADRootDSE -ErrorAction Stop
    $result = @{ Ready = $true; Roles = @(); Detail = "Read-only AD probe succeeded for $($rootDse.defaultNamingContext)." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'AD' -EnginePath $windowsPowerShellPath -ScriptText $adScript -TimeoutSeconds $ProbeTimeoutSeconds))

        $exchangeScript = @'
$ErrorActionPreference = 'Stop'
try {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
    $server = Get-ExchangeServer -ErrorAction Stop | Select-Object -First 1
    $result = @{ Ready = $true; Roles = @(); Detail = "Read-only Exchange on-premises probe succeeded against $($server.Name)." }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'ExchangeOnPrem' -EnginePath $windowsPowerShellPath -ScriptText $exchangeScript -TimeoutSeconds $ProbeTimeoutSeconds))

        $teamsScript = @'
$ErrorActionPreference = 'Stop'
try {
    Import-Module MicrosoftTeams -MinimumVersion 4.7.1 -ErrorAction Stop
    Connect-MicrosoftTeams -TenantId $env:SMART_TENANT_ID -ApplicationId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT | Out-Null
    $null = Get-CsTenant -ErrorAction Stop
    $result = @{ Ready = $true; Roles = @(); Detail = 'Read-only Teams PowerShell probe succeeded.' }
}
catch { $result = @{ Ready = $false; Roles = @(); Detail = $_.Exception.Message } }
finally {
    try {
        if ($null -ne (Get-Command -Name Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue)) {
            Disconnect-MicrosoftTeams -ErrorAction Stop | Out-Null
        }
    }
    catch {}
}
"SMARTM365_CAPABILITY_JSON:$($result | ConvertTo-Json -Compress -Depth 5)"
'@
        $results.Add((Invoke-IsolatedCapabilityProbe -Name 'TeamsPowerShell' -EnginePath $pwshPath -ScriptText $teamsScript -Environment $environment -TimeoutSeconds $ProbeTimeoutSeconds))
    }

    $readyCapabilities = @($results | Where-Object Ready | ForEach-Object Name | Sort-Object -Unique)
    $graphResult = $results | Where-Object Name -eq 'Graph' | Select-Object -First 1
    return [pscustomobject]@{
        SchemaVersion = 1
        ModuleVersion = $script:DistributedModuleVersion
        ServerName = $ServerName.ToUpperInvariant()
        GeneratedAtUtc = $generatedAtUtc.ToString('o')
        ProbeMode = $ProbeMode
        ReadyCapabilities = $readyCapabilities
        GraphAppRoles = @($graphResult.Roles | Sort-Object -Unique)
        Results = @($results)
    }
}

function Test-SmartM365OrchestratorCapabilityMatch {
    param(
        [Parameter(Mandatory = $true)]$ServerCapabilities,
        [string[]]$RequiredCapabilities = @(),
        [string[]]$RequiredGraphAppRoles = @()
    )

    $ready = @($ServerCapabilities.ReadyCapabilities | ForEach-Object { [string]$_ })
    $roles = @($ServerCapabilities.GraphAppRoles | ForEach-Object { [string]$_ })
    $missingCapabilities = @($RequiredCapabilities | Where-Object { $_ -notin $ready } | Sort-Object -Unique)
    $missingRoles = @($RequiredGraphAppRoles | Where-Object { $_ -notin $roles } | Sort-Object -Unique)
    return [pscustomobject]@{
        Eligible = ($missingCapabilities.Count -eq 0 -and $missingRoles.Count -eq 0)
        MissingCapabilities = $missingCapabilities
        MissingGraphAppRoles = $missingRoles
    }
}

function Get-ScheduleRunsPerDay {
    param($Schedule)

    if ($null -eq $Schedule) { return 0.0 }
    $times = @($Schedule.Times).Count
    if ([string]$Schedule.Type -eq 'Weekly') {
        return [double]($times * @($Schedule.DaysOfWeek).Count) / 7.0
    }
    return [double]$times
}

function Get-DependencyComponent {
    param([Parameter(Mandatory = $true)][object[]]$Jobs)

    $jobsByName = @{}
    foreach ($job in $Jobs) { $jobsByName[[string]$job.Name] = $job }
    $adjacency = @{}
    foreach ($job in $Jobs) { $adjacency[[string]$job.Name] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
    foreach ($job in $Jobs) {
        foreach ($dependencyName in @($job.DependsOn)) {
            if ($jobsByName.ContainsKey([string]$dependencyName)) {
                [void]$adjacency[[string]$job.Name].Add([string]$dependencyName)
                [void]$adjacency[[string]$dependencyName].Add([string]$job.Name)
            }
        }
    }

    $visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $components = @()
    foreach ($name in @($jobsByName.Keys | Sort-Object)) {
        if ($visited.Contains($name)) { continue }
        $queue = [System.Collections.Generic.Queue[string]]::new()
        $queue.Enqueue($name)
        [void]$visited.Add($name)
        $members = @()
        while ($queue.Count -gt 0) {
            $current = $queue.Dequeue()
            $members += $current
            foreach ($neighbor in $adjacency[$current]) {
                if ($visited.Add($neighbor)) { $queue.Enqueue($neighbor) }
            }
        }
        $components += ,@($members | Sort-Object)
    }
    return $components
}

function Test-ServerJobPolicy {
    param(
        [Parameter(Mandatory = $true)][string]$ServerName,
        [Parameter(Mandatory = $true)][object[]]$Jobs,
        [hashtable]$ServerJobPolicies = @{}
    )

    if (-not $ServerJobPolicies.ContainsKey($ServerName)) { return $true }
    $policy = $ServerJobPolicies[$ServerName]
    $onlyJobsRequiring = @()
    if ($policy -is [System.Collections.IDictionary]) {
        if ($policy.Contains('OnlyJobsRequiring')) { $onlyJobsRequiring = @($policy['OnlyJobsRequiring']) }
    }
    elseif ($null -ne $policy -and $policy.PSObject.Properties['OnlyJobsRequiring']) {
        $onlyJobsRequiring = @($policy.OnlyJobsRequiring)
    }
    $onlyJobsRequiring = @($onlyJobsRequiring | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    if ($onlyJobsRequiring.Count -eq 0) { return $true }

    foreach ($job in $Jobs) {
        $jobCapabilities = @($job.RequiredCapabilities | ForEach-Object { [string]$_ })
        if (@($onlyJobsRequiring | Where-Object { $_ -notin $jobCapabilities }).Count -gt 0) { return $false }
    }
    return $true
}

function Get-SmartM365OrchestratorElectionPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Jobs,
        [Parameter(Mandatory = $true)][object[]]$ServerCapabilities,
        [hashtable]$ServerWeights = @{},
        [hashtable]$ServerJobPolicies = @{},
        [hashtable]$DurationMinutesByJob = @{},
        [AllowNull()]$PreviousPlan = $null,
        [switch]$PreservePreviousOwners,
        [datetime]$NowUtc = [datetime]::UtcNow
    )

    $electedJobs = @($Jobs | Where-Object { $_.Enabled -and [string]$_.AssignmentMode -eq 'Elected' })
    $capabilitiesByServer = @{}
    foreach ($serverCapability in $ServerCapabilities) {
        $capabilitiesByServer[[string]$serverCapability.ServerName] = $serverCapability
    }
    $loads = @{}
    foreach ($serverName in $capabilitiesByServer.Keys) { $loads[$serverName] = 0.0 }

    $previousOwnerByJob = @{}
    if ($PreservePreviousOwners -and $null -ne $PreviousPlan -and $PreviousPlan.PSObject.Properties['Assignments']) {
        foreach ($assignment in @($PreviousPlan.Assignments)) {
            if (-not $assignment.PSObject.Properties['JobName'] -or -not $assignment.PSObject.Properties['OwnerServer']) { continue }
            $jobName = [string]$assignment.JobName
            $ownerServer = ([string]$assignment.OwnerServer).ToUpperInvariant()
            if (-not [string]::IsNullOrWhiteSpace($jobName) -and -not [string]::IsNullOrWhiteSpace($ownerServer)) {
                $previousOwnerByJob[$jobName] = $ownerServer
            }
        }
    }

    $groups = @()
    # Dependencies constrain execution order through shared occurrence claims,
    # not placement: cloud and on-premises prerequisites can live on different hosts.
    foreach ($electedJob in $electedJobs) {
        $component = @([string]$electedJob.Name)
        $componentJobs = @($electedJobs | Where-Object { $_.Name -in $component })
        $requiredCapabilities = @($componentJobs | ForEach-Object { @($_.RequiredCapabilities) } | Sort-Object -Unique)
        $requiredRoles = @($componentJobs | ForEach-Object { @($_.RequiredGraphAppRoles) } | Sort-Object -Unique)
        $loadMinutes = 0.0
        foreach ($job in $componentJobs) {
            $duration = if ($DurationMinutesByJob.ContainsKey([string]$job.Name)) {
                [double]$DurationMinutesByJob[[string]$job.Name]
            }
            elseif ($null -ne $job.EstimatedDurationMinutes -and [double]$job.EstimatedDurationMinutes -gt 0) {
                [double]$job.EstimatedDurationMinutes
            }
            else { 5.0 }
            $loadMinutes += $duration * (Get-ScheduleRunsPerDay -Schedule $job.Schedule)
        }
        $candidates = @()
        foreach ($serverName in @($capabilitiesByServer.Keys | Sort-Object)) {
            $match = Test-SmartM365OrchestratorCapabilityMatch -ServerCapabilities $capabilitiesByServer[$serverName] -RequiredCapabilities $requiredCapabilities -RequiredGraphAppRoles $requiredRoles
            if ($match.Eligible -and (Test-ServerJobPolicy -ServerName $serverName -Jobs $componentJobs -ServerJobPolicies $ServerJobPolicies)) { $candidates += $serverName }
        }
        $previousOwners = @(
            $component |
                ForEach-Object { if ($previousOwnerByJob.ContainsKey([string]$_)) { $previousOwnerByJob[[string]$_] } } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )
        $preservedOwner = ''
        if ($PreservePreviousOwners -and $previousOwners.Count -eq 1 -and $previousOwners[0] -in $candidates) {
            $preservedOwner = [string]$previousOwners[0]
        }
        $groups += [pscustomobject]@{
            GroupKey = (@($component) -join '+')
            Jobs = @($component)
            RequiredCapabilities = $requiredCapabilities
            RequiredGraphAppRoles = $requiredRoles
            LoadMinutesPerDay = [math]::Round($loadMinutes, 4)
            Candidates = $candidates
            PreservedOwner = $preservedOwner
        }
    }

    $assignments = @()
    $unassigned = @()
    $orderedGroups = @($groups | Sort-Object @{ Expression = 'LoadMinutesPerDay'; Descending = $true }, GroupKey)
    foreach ($group in $orderedGroups) {
        if (@($group.Candidates).Count -eq 0) {
            $unassigned += [pscustomobject]@{ GroupKey = $group.GroupKey; Jobs = $group.Jobs; Reason = 'No live server satisfies every required capability, Graph application role and server job policy.' }
            continue
        }
        if ([string]::IsNullOrWhiteSpace([string]$group.PreservedOwner)) { continue }
        $owner = [string]$group.PreservedOwner
        $loads[$owner] = [double]$loads[$owner] + [double]$group.LoadMinutesPerDay
        foreach ($jobName in $group.Jobs) {
            $assignments += [pscustomobject]@{
                JobName = $jobName
                OwnerServer = $owner
                GroupKey = $group.GroupKey
                LoadMinutesPerDay = $group.LoadMinutesPerDay
            }
        }
    }

    foreach ($group in $orderedGroups) {
        if (@($group.Candidates).Count -eq 0 -or -not [string]::IsNullOrWhiteSpace([string]$group.PreservedOwner)) { continue }
        $owner = @($group.Candidates | Sort-Object @{
                Expression = {
                    $weight = if ($ServerWeights.ContainsKey([string]$_) -and [double]$ServerWeights[[string]$_] -gt 0) { [double]$ServerWeights[[string]$_] } else { 1.0 }
                    [double]$loads[[string]$_] / $weight
                }
            }, @{ Expression = { [string]$_ } })[0]
        $loads[$owner] = [double]$loads[$owner] + [double]$group.LoadMinutesPerDay
        foreach ($jobName in $group.Jobs) {
            $assignments += [pscustomobject]@{
                JobName = $jobName
                OwnerServer = $owner
                GroupKey = $group.GroupKey
                LoadMinutesPerDay = $group.LoadMinutesPerDay
            }
        }
    }

    return [pscustomobject]@{
        SchemaVersion = 2
        PlanId = [guid]::NewGuid().ToString('N')
        GeneratedAtUtc = $NowUtc.ToString('o')
        EligibleServers = @($capabilitiesByServer.Keys | Sort-Object)
        Assignments = @($assignments | Sort-Object JobName)
        UnassignedGroups = @($unassigned)
        ServerLoads = @($loads.Keys | Sort-Object | ForEach-Object {
                [pscustomobject]@{
                    ServerName = $_
                    Weight = if ($ServerWeights.ContainsKey($_)) { [double]$ServerWeights[$_] } else { 1.0 }
                    LoadMinutesPerDay = [math]::Round([double]$loads[$_], 4)
                }
            })
    }
}

function Test-SmartM365OrchestratorCanPreserveOwners {
    [CmdletBinding()]
    param(
        [AllowNull()]$PreviousPlan = $null,
        [Parameter(Mandatory = $true)][object[]]$ServerCapabilities
    )

    if ($null -eq $PreviousPlan -or
        -not $PreviousPlan.PSObject.Properties['SchemaVersion'] -or
        [int]$PreviousPlan.SchemaVersion -lt 2 -or
        -not $PreviousPlan.PSObject.Properties['EligibleServers']) {
        return $false
    }
    if ($PreviousPlan.PSObject.Properties['UnassignedGroups'] -and @($PreviousPlan.UnassignedGroups).Count -gt 0) {
        return $false
    }

    $previousServers = @(
        $PreviousPlan.EligibleServers |
            ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
    $currentServers = @(
        $ServerCapabilities |
            ForEach-Object { ([string]$_.ServerName).Trim().ToUpperInvariant() } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
    if ($previousServers.Count -ne $currentServers.Count) { return $false }
    return (($previousServers -join '|') -ceq ($currentServers -join '|'))
}

function Get-SmartM365OrchestratorOccurrenceClaim {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ClaimsRootPath,
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][datetime]$Occurrence
    )

    $jobFolder = Join-Path -Path $ClaimsRootPath -ChildPath (ConvertTo-SafeFileName $JobName)
    $occurrenceUtc = $Occurrence.ToUniversalTime()
    $claimPath = Join-Path -Path $jobFolder -ChildPath ($occurrenceUtc.ToString('yyyyMMddTHHmmssfffZ') + '.json'); $claimPath = Resolve-DistributedJsonPath $claimPath
    if (-not (Test-Path -LiteralPath $claimPath -PathType Leaf)) { return $null }
    $claim = (Read-SmartM365JsonDocument $claimPath).Document
    return [pscustomobject]@{
        ClaimPath = $claimPath
        Claim = $claim
    }
}

function Enter-SmartM365OrchestratorOccurrenceClaim {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ClaimsRootPath,
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][datetime]$Occurrence,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [Parameter(Mandatory = $true)][string]$PlanId,
        [int]$SafeMinutes = 60,
        [string]$HeartbeatRootPath = '',
        [int]$HeartbeatStaleMinutes = 5
    )

    $jobFolder = Join-Path -Path $ClaimsRootPath -ChildPath (ConvertTo-SafeFileName $JobName)
    [void](New-Item -ItemType Directory -Path $jobFolder -Force)
    $occurrenceUtc = $Occurrence.ToUniversalTime()
    $claimPath = Join-Path -Path $jobFolder -ChildPath ($occurrenceUtc.ToString('yyyyMMddTHHmmssfffZ') + '.json'); $claimPath = Resolve-DistributedJsonPath $claimPath
    $claim = [ordered]@{
        SchemaVersion = 1
        ClaimId = [guid]::NewGuid().ToString('N')
        JobName = $JobName
        OccurrenceUtc = $occurrenceUtc.ToString('o')
        OwnerServer = $OwnerServer.ToUpperInvariant()
        PlanId = $PlanId
        OrchestratorPid = $PID
        Status = 'Claimed'
        Attempt = 0
        CreatedAtUtc = [datetime]::UtcNow.ToString('o')
        UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
        SafeUntilUtc = [datetime]::UtcNow.AddMinutes([math]::Max(1, $SafeMinutes)).ToString('o')
    }
    $json = $claim | ConvertTo-Json -Depth 6

    try {
        $null = Write-SmartM365JsonBytesAtomically -Path $claimPath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($json)) -ExpectedSHA256 'ABSENT' -Validate { param($document) if (-not $document.PSObject.Properties['ClaimId']) { throw 'ClaimId missing.' } }
        return [pscustomobject]@{ Acquired = $true; Reused = $false; ClaimPath = $claimPath; Claim = [pscustomobject]$claim; Reason = '' }
    }
    catch [System.IO.IOException] {
        try {
            $existing = (Read-SmartM365JsonDocument $claimPath).Document
            $sameOwner = [string]$existing.OwnerServer -eq $OwnerServer.ToUpperInvariant()
            $sameOrchestratorProcess = $false
            try { $sameOrchestratorProcess = [int]$existing.OrchestratorPid -eq $PID } catch { $sameOrchestratorProcess = $false }
            $reusable = [string]$existing.Status -eq 'RetryScheduled' -or
                ([string]$existing.Status -eq 'Claimed' -and $sameOrchestratorProcess)
            if ($sameOwner -and $reusable) {
                return [pscustomobject]@{ Acquired = $true; Reused = $true; ClaimPath = $claimPath; Claim = $existing; Reason = 'Existing non-terminal claim belongs to this server.' }
            }

            $terminal = [string]$existing.Status -in @('Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted')
            $safeUntilUtc = [datetime]::MaxValue
            try { $safeUntilUtc = ([datetime]$existing.SafeUntilUtc).ToUniversalTime() } catch { $safeUntilUtc = [datetime]::MaxValue }
            if (-not $terminal -and [datetime]::UtcNow -gt $safeUntilUtc -and -not [string]::IsNullOrWhiteSpace($HeartbeatRootPath)) {
                $heartbeatFresh = $false
                try {
                    $heartbeatPath = Join-Path -Path (Join-Path -Path $HeartbeatRootPath -ChildPath ([string]$existing.OwnerServer)) -ChildPath 'Orchestrator-Heartbeat.json'
                    $heartbeat = if(Get-SmartM365JsonReadPath $heartbeatPath -Optional){(Read-SmartM365JsonDocument $heartbeatPath).Document}else{$null}
                    $heartbeatAgeMinutes = if($heartbeat){([datetime]::UtcNow - ([datetime]$heartbeat.Timestamp).ToUniversalTime()).TotalMinutes}else{[double]::PositiveInfinity}
                    $heartbeatFresh = $heartbeatAgeMinutes -le [math]::Max(1, $HeartbeatStaleMinutes)
                }
                catch { throw ('Cannot determine peer liveness; existing coordination state preserved: ' + $_.Exception.Message) }
                # A restarted orchestrator on the same server may replace its own
                # expired claim even though its new heartbeat is healthy.
                if ($sameOwner) { $heartbeatFresh = $false }

                if (-not $heartbeatFresh) {
                    $takeoverLockPath = (Get-SmartM365JsonNames $claimPath).Legacy + '.takeover.lock'
                    $takeoverStream = $null
                    try {
                        $takeoverStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                        $confirmed = (Read-SmartM365JsonDocument $claimPath).Document
                        $confirmedSafeUntilUtc = ([datetime]$confirmed.SafeUntilUtc).ToUniversalTime()
                        if ([string]$confirmed.ClaimId -eq [string]$existing.ClaimId -and
                            [string]$confirmed.Status -notin @('Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted') -and
                            [datetime]::UtcNow -gt $confirmedSafeUntilUtc) {
                            $archivePath = '{0}.stale.{1}.json' -f (Get-SmartM365JsonNames $claimPath).Legacy, [string]$confirmed.ClaimId
                            if ((Get-SmartM365JsonTransportPolicy).Mode -eq 'JsonText') { $archivePath += '.txt' }
                            Move-Item -LiteralPath $claimPath -Destination $archivePath -ErrorAction Stop
                        }
                    }
                    catch [System.IO.IOException] {
                        return [pscustomobject]@{ Acquired = $false; Reused = $false; ClaimPath = $claimPath; Claim = $existing; Reason = 'Another orchestrator is evaluating takeover of the expired claim.' }
                    }
                    finally {
                        if ($null -ne $takeoverStream) {
                            $takeoverStream.Dispose()
                            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                    if (-not (Test-Path -LiteralPath $claimPath)) {
                        return Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $ClaimsRootPath -JobName $JobName -Occurrence $Occurrence -OwnerServer $OwnerServer -PlanId $PlanId -SafeMinutes $SafeMinutes -HeartbeatRootPath $HeartbeatRootPath -HeartbeatStaleMinutes $HeartbeatStaleMinutes
                    }
                }
            }
            return [pscustomobject]@{ Acquired = $false; Reused = $false; ClaimPath = $claimPath; Claim = $existing; Reason = "Occurrence already claimed by $($existing.OwnerServer) with status $($existing.Status)." }
        }
        catch {
            return [pscustomobject]@{ Acquired = $false; Reused = $false; ClaimPath = $claimPath; Claim = $null; Reason = "Claim exists but cannot be read safely: $($_.Exception.Message)" }
        }
    }
}

function Set-SmartM365OrchestratorOccurrenceClaim {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory = $true)][string]$ClaimPath,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [Parameter(Mandatory = $true)][ValidateSet('Claimed', 'Running', 'RetryScheduled', 'Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted')][string]$Status,
        [int]$Attempt = 0,
        [int]$SafeMinutes = 60,
        [string]$Detail = ''
    )

    $ClaimPath = Resolve-DistributedJsonPath $ClaimPath
    $claimDocument = Read-SmartM365JsonDocument $ClaimPath
    $claim = $claimDocument.Document
    if ([string]$claim.OwnerServer -ne $OwnerServer) {
        throw "Claim owner mismatch for $ClaimPath. Expected $OwnerServer, found $($claim.OwnerServer)."
    }
    $claim.Status = $Status
    $claim.Attempt = $Attempt
    $claim.UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
    $claim.SafeUntilUtc = [datetime]::UtcNow.AddMinutes([math]::Max(1, $SafeMinutes)).ToString('o')
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        if ($claim.PSObject.Properties.Name -contains 'Detail') { $claim.Detail = $Detail }
        else { $claim | Add-Member -NotePropertyName Detail -NotePropertyValue $Detail }
    }
    if ($PSCmdlet.ShouldProcess($ClaimPath, "Set occurrence claim status to $Status")) {
        Write-JsonAtomically -Path $ClaimPath -Value $claim -ExpectedSHA256 $claimDocument.SHA256
    }
    return $claim
}

function Get-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasesRootPath,
        [Parameter(Mandatory = $true)][string]$ConcurrencyKey
    )

    $leasePath = Join-Path -Path $LeasesRootPath -ChildPath ((ConvertTo-SafeFileName $ConcurrencyKey) + '.json'); $leasePath = Resolve-DistributedJsonPath $leasePath
    if (-not (Test-Path -LiteralPath $leasePath -PathType Leaf)) { return $null }
    $lease = (Read-SmartM365JsonDocument $leasePath).Document
    return [pscustomobject]@{
        LeasePath = $leasePath
        Lease = $lease
    }
}

function Enter-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasesRootPath,
        [Parameter(Mandatory = $true)][string]$ConcurrencyKey,
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][datetime]$Occurrence,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [int]$SafeMinutes = 60,
        [string]$HeartbeatRootPath = '',
        [int]$HeartbeatStaleMinutes = 5
    )

    [void](New-Item -ItemType Directory -Path $LeasesRootPath -Force)
    $leasePath = Join-Path -Path $LeasesRootPath -ChildPath ((ConvertTo-SafeFileName $ConcurrencyKey) + '.json'); $leasePath = Resolve-DistributedJsonPath $leasePath
    $owner = $OwnerServer.ToUpperInvariant()
    $occurrenceUtc = $Occurrence.ToUniversalTime()
    $lease = [ordered]@{
        SchemaVersion = 1
        LeaseId = [guid]::NewGuid().ToString('N')
        ConcurrencyKey = $ConcurrencyKey
        JobName = $JobName
        OccurrenceUtc = $occurrenceUtc.ToString('o')
        OwnerServer = $owner
        OrchestratorPid = $PID
        CreatedAtUtc = [datetime]::UtcNow.ToString('o')
        UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
        SafeUntilUtc = [datetime]::UtcNow.AddMinutes([math]::Max(1, $SafeMinutes)).ToString('o')
    }
    $json = $lease | ConvertTo-Json -Depth 6

    try {
        $null = Write-SmartM365JsonBytesAtomically -Path $leasePath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($json)) -ExpectedSHA256 'ABSENT' -Validate { param($document) if (-not $document.PSObject.Properties['LeaseId']) { throw 'LeaseId missing.' } }
        return [pscustomobject]@{ Acquired = $true; Reused = $false; LeasePath = $leasePath; Lease = [pscustomobject]$lease; Reason = '' }
    }
    catch [System.IO.IOException] {
        try {
            $existing = (Read-SmartM365JsonDocument $leasePath).Document
            $sameOwner = [string]$existing.OwnerServer -eq $owner
            $sameProcess = $false
            try { $sameProcess = [int]$existing.OrchestratorPid -eq $PID } catch { $sameProcess = $false }
            $existingOccurrenceUtc = [datetime]::MinValue
            try { $existingOccurrenceUtc = ([datetime]$existing.OccurrenceUtc).ToUniversalTime() } catch { $existingOccurrenceUtc = [datetime]::MinValue }
            $sameOccurrence = [string]$existing.JobName -eq $JobName -and [math]::Abs(($existingOccurrenceUtc - $occurrenceUtc).TotalSeconds) -le 1
            if ($sameOwner -and $sameProcess -and $sameOccurrence) {
                return [pscustomobject]@{ Acquired = $true; Reused = $true; LeasePath = $leasePath; Lease = $existing; Reason = 'Existing concurrency lease belongs to this process and occurrence.' }
            }

            $safeUntilUtc = [datetime]::MaxValue
            try { $safeUntilUtc = ([datetime]$existing.SafeUntilUtc).ToUniversalTime() } catch { $safeUntilUtc = [datetime]::MaxValue }
            if ([datetime]::UtcNow -gt $safeUntilUtc) {
                $heartbeatFresh = $false
                if (-not [string]::IsNullOrWhiteSpace($HeartbeatRootPath)) {
                    try {
                        $heartbeatPath = Join-Path -Path (Join-Path -Path $HeartbeatRootPath -ChildPath ([string]$existing.OwnerServer)) -ChildPath 'Orchestrator-Heartbeat.json'
                        $heartbeat = if(Get-SmartM365JsonReadPath $heartbeatPath -Optional){(Read-SmartM365JsonDocument $heartbeatPath).Document}else{$null}
                        $heartbeatAgeMinutes = if($heartbeat){([datetime]::UtcNow - ([datetime]$heartbeat.Timestamp).ToUniversalTime()).TotalMinutes}else{[double]::PositiveInfinity}
                        $heartbeatFresh = $heartbeatAgeMinutes -le [math]::Max(1, $HeartbeatStaleMinutes)
                    }
                    catch { throw ('Cannot determine peer liveness; existing coordination state preserved: ' + $_.Exception.Message) }
                }
                if ($sameOwner) { $heartbeatFresh = $false }

                if (-not $heartbeatFresh) {
                    $takeoverLockPath = (Get-SmartM365JsonNames $leasePath).Legacy + '.takeover.lock'
                    $takeoverStream = $null
                    try {
                        if ((Test-Path -LiteralPath $takeoverLockPath -PathType Leaf) -and
                            ([datetime]::UtcNow - (Get-Item -LiteralPath $takeoverLockPath).LastWriteTimeUtc).TotalMinutes -gt 5) {
                            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
                        }
                        $takeoverStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                        $confirmed = (Read-SmartM365JsonDocument $leasePath).Document
                        $confirmedSafeUntilUtc = ([datetime]$confirmed.SafeUntilUtc).ToUniversalTime()
                        if ([string]$confirmed.LeaseId -eq [string]$existing.LeaseId -and [datetime]::UtcNow -gt $confirmedSafeUntilUtc) {
                            $archivePath = '{0}.stale.{1}.json' -f (Get-SmartM365JsonNames $leasePath).Legacy, [string]$confirmed.LeaseId
                            if ((Get-SmartM365JsonTransportPolicy).Mode -eq 'JsonText') { $archivePath += '.txt' }
                            Move-Item -LiteralPath $leasePath -Destination $archivePath -ErrorAction Stop
                        }
                    }
                    catch [System.IO.IOException] {
                        return [pscustomobject]@{ Acquired = $false; Reused = $false; LeasePath = $leasePath; Lease = $existing; Reason = 'Another orchestrator is evaluating takeover of the expired concurrency lease.' }
                    }
                    finally {
                        if ($null -ne $takeoverStream) {
                            $takeoverStream.Dispose()
                            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                    if (-not (Test-Path -LiteralPath $leasePath)) {
                        return Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $LeasesRootPath -ConcurrencyKey $ConcurrencyKey -JobName $JobName -Occurrence $Occurrence -OwnerServer $OwnerServer -SafeMinutes $SafeMinutes -HeartbeatRootPath $HeartbeatRootPath -HeartbeatStaleMinutes $HeartbeatStaleMinutes
                    }
                }
            }
            $reason = "ConcurrencyKey '$ConcurrencyKey' is held by job $($existing.JobName) on $($existing.OwnerServer) until $($existing.SafeUntilUtc)."
            return [pscustomobject]@{ Acquired = $false; Reused = $false; LeasePath = $leasePath; Lease = $existing; Reason = $reason }
        }
        catch {
            return [pscustomobject]@{ Acquired = $false; Reused = $false; LeasePath = $leasePath; Lease = $null; Reason = "Concurrency lease exists but cannot be read safely: $($_.Exception.Message)" }
        }
    }
}

function Set-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasePath,
        [Parameter(Mandatory = $true)][string]$LeaseId,
        [Parameter(Mandatory = $true)][string]$OwnerServer,
        [Parameter(Mandatory = $true)][datetime]$SafeUntilUtc
    )

    $LeasePath = Resolve-DistributedJsonPath $LeasePath
    if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $false }
    $takeoverLockPath = (Get-SmartM365JsonNames $LeasePath).Legacy + '.takeover.lock'
    $lockStream = $null
    try {
        for ($attempt = 1; $attempt -le 10 -and $null -eq $lockStream; $attempt++) {
            try {
                $lockStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            }
            catch [System.IO.IOException] {
                if ($attempt -eq 10) { throw }
                Start-Sleep -Milliseconds 100
            }
        }

        if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $false }
        $current = (Read-SmartM365JsonDocument $LeasePath).Document
        if ([string]$current.LeaseId -ne $LeaseId -or [string]$current.OwnerServer -ne $OwnerServer.ToUpperInvariant()) {
            return $false
        }

        $targetSafeUntilUtc = $SafeUntilUtc.ToUniversalTime()
        $currentSafeUntilUtc = [datetime]::MinValue
        $hasCurrentSafeUntil = [datetime]::TryParse(
            [string]$current.SafeUntilUtc,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::RoundtripKind,
            [ref]$currentSafeUntilUtc
        )
        $sameDeadline = $hasCurrentSafeUntil -and
            [math]::Abs(($currentSafeUntilUtc.ToUniversalTime() - $targetSafeUntilUtc).TotalSeconds) -le 1
        $sameProcess = $false
        try { $sameProcess = [int]$current.OrchestratorPid -eq $PID } catch { $sameProcess = $false }
        if ($sameDeadline -and $sameProcess) { return $true }

        $current.OrchestratorPid = $PID
        $current.UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
        $current.SafeUntilUtc = $targetSafeUntilUtc.ToString('o')
        Write-JsonAtomically -Path $LeasePath -Value $current
        return $true
    }
    finally {
        if ($null -ne $lockStream) {
            $lockStream.Dispose()
            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Exit-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasePath,
        [Parameter(Mandatory = $true)][string]$LeaseId,
        [Parameter(Mandatory = $true)][string]$OwnerServer
    )

    $LeasePath = Resolve-DistributedJsonPath $LeasePath
    if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $true }
    $takeoverLockPath = (Get-SmartM365JsonNames $LeasePath).Legacy + '.takeover.lock'
    $lockStream = $null
    try {
        for ($attempt = 1; $attempt -le 10 -and $null -eq $lockStream; $attempt++) {
            try {
                $lockStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            }
            catch [System.IO.IOException] {
                if ($attempt -eq 10) { throw }
                Start-Sleep -Milliseconds 100
            }
        }
        if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $true }
        $current = (Read-SmartM365JsonDocument $LeasePath).Document
        if ([string]$current.LeaseId -ne $LeaseId -or [string]$current.OwnerServer -ne $OwnerServer.ToUpperInvariant()) {
            return $false
        }
        $consumedLease = Read-SmartM365JsonDocument $LeasePath
        if ($consumedLease.Document.LeaseId -ne $LeaseId -or $consumedLease.Document.OwnerServer -ne $OwnerServer.ToUpperInvariant()) { return $false }
        Complete-SmartM365JsonConsumption -Path $LeasePath -Owner 'Orchestrator concurrency lease' -ExpectedSHA256 $consumedLease.SHA256
        return $true
    }
    finally {
        if ($null -ne $lockStream) {
            $lockStream.Dispose()
            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-SmartM365OrchestratorFreshProducerReceipt {
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)]$Job,
        [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][datetime]$Now,
        [Parameter(Mandatory)][double]$MaxAgeHours
    )

    if ($MaxAgeHours -le 0 -or [string]::IsNullOrWhiteSpace($TenantKey)) { return $null }
    $scriptName = [IO.Path]::GetFileName([string]$Job.ScriptPath)
    if ($scriptName -notmatch '^SmartM365-[A-Za-z0-9-]+\.ps1$') { return $null }
    # One script can back several jobs with different arguments; its receipt cannot identify those variants.
    if ([IO.Path]::GetFileNameWithoutExtension($scriptName) -cne ('SmartM365-' + [string]$Job.Name)) { return $null }
    if ([IO.Path]::GetFileName($SharedDataFolderPath.TrimEnd([char[]]'\/')) -cne 'Orchestrator') { return $null }
    $dataAll = Split-Path -Path $SharedDataFolderPath -Parent
    if ([IO.Path]::GetFileName($dataAll) -cne 'DATA-ALL') { return $null }
    $latest = Join-Path (Split-Path -Path $dataAll -Parent) 'DATA-LAST'
    $receiptPath = Join-Path $latest ('SmartInventory_{0}.current.json' -f [IO.Path]::GetFileNameWithoutExtension($scriptName))
    try {
        $readPath = Get-SmartM365JsonReadPath -Path $receiptPath -Optional
        if (-not $readPath) { return $null }
        $receipt = (Read-SmartM365JsonDocument -Path $readPath).Document
        foreach ($field in @('Owner','Status','TenantKey','Producer','RunId','StartedAtUtc','CompletedAtUtc','Errors','IsPartialInventory','Files')) {
            if (-not $receipt.PSObject.Properties[$field]) { return $null }
        }
        if ([string]$receipt.Owner -cne 'SmartInventory-CmdbSourceReceipt' -or
            [string]$receipt.Status -cne 'Completed' -or
            [string]$receipt.TenantKey -cne $TenantKey -or
            [string]$receipt.Producer -cne $scriptName -or
            [string]::IsNullOrWhiteSpace([string]$receipt.RunId) -or
            [int]$receipt.Errors -ne 0 -or
            $null -eq $receipt.IsPartialInventory -or [bool]$receipt.IsPartialInventory) { return $null }

        $started = [datetimeoffset]::MinValue
        $completed = [datetimeoffset]::MinValue
        $styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
        if (-not [datetimeoffset]::TryParse([string]$receipt.StartedAtUtc, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$started) -or
            -not [datetimeoffset]::TryParse([string]$receipt.CompletedAtUtc, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$completed)) { return $null }
        $nowUtc = $Now.ToUniversalTime()
        if ($started -gt $completed -or $completed.UtcDateTime -gt $nowUtc.AddMinutes(5) -or
            $completed.UtcDateTime -lt $nowUtc.AddHours(-1 * $MaxAgeHours)) { return $null }

        $files = @($receipt.Files)
        if ($files.Count -eq 0) { return $null }
        foreach ($file in $files) {
            foreach ($field in @('File','RunId','Producer','Status','Errors','IsPartialInventory','SHA256')) {
                if (-not $file.PSObject.Properties[$field]) { return $null }
            }
            $fileName = [string]$file.File
            if ([string]::IsNullOrWhiteSpace($fileName) -or $fileName -cne [IO.Path]::GetFileName($fileName) -or
                [string]$file.RunId -cne [string]$receipt.RunId -or
                [string]$file.Producer -cne $scriptName -or
                [string]$file.Status -cne 'Success' -or [int]$file.Errors -ne 0 -or
                $null -eq $file.IsPartialInventory -or [bool]$file.IsPartialInventory -or
                [string]$file.SHA256 -notmatch '^[A-Fa-f0-9]{64}$') { return $null }
            $csvPath = Join-Path $latest $fileName
            $csv = Get-Item -LiteralPath $csvPath -ErrorAction SilentlyContinue
            if ($null -eq $csv -or $csv.PSIsContainer -or $csv.LinkType -in @('SymbolicLink', 'Junction')) { return $null }
        }
        return $completed.UtcDateTime
    }
    catch { return $null }
}

Export-ModuleMember -Function @(
    'Get-SmartM365OrchestratorServerCapability',
    'Test-SmartM365OrchestratorCapabilityMatch',
    'Test-SmartM365OrchestratorCanPreserveOwners',
    'Get-SmartM365OrchestratorElectionPlan',
    'Get-SmartM365OrchestratorOccurrenceClaim',
    'Enter-SmartM365OrchestratorOccurrenceClaim',
    'Set-SmartM365OrchestratorOccurrenceClaim',
    'Get-SmartM365OrchestratorConcurrencyLease',
    'Enter-SmartM365OrchestratorConcurrencyLease',
    'Set-SmartM365OrchestratorConcurrencyLease',
    'Exit-SmartM365OrchestratorConcurrencyLease',
    'Convert-SmartM365OrchestratorDistributedHistory',
    'Get-SmartM365OrchestratorFreshProducerReceipt'
)

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCElrs50scqPLFS
# pooEjqA9Z8RxnMOm76AGANVykWDW/qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDpxzGLzmbvpjvXxzSbWoux
# JOSD/k6hnUTjxIEkpHr+czANBgkqhkiG9w0BAQEFAASCAYCohfZbLHZo3WqMJiYL
# 3f+6Sl3ZEoRIfv4f7zW9xooWItaOOI1cCCSMdQH91GwU9BqhjP4IpqzAg3UleRLQ
# 58shLTrKYj0l5JrExlp80MvRWWDmYO/UcAwLgz4VL8rDUxr6PYAvCv6Cf7i6niWQ
# zUGNqdm+gwS0ALPFi+ODd2nQVAYZUxzTFoxbIL1kdEuwdiVHDYs6T+HNkqP0etUE
# XovM8r4JV9EYNTKbZowXj2APaAFu7eb0gcjwbKqruKdrY7F/sxU/KD5ly7MfG4rT
# 3AM8BpflF2T8IbhK6ZP+vjVEK0EheB39a6Zl+4BfmNZ6ZCBPD8Jh1gZyQqJFHcVv
# vuG9wJaD2LM/Ag+pBIAFHHPw/iwR8bsCAPv5i8mzcv3hhQ7fZnE6bDRistJuot3Z
# GdWz24BnZgFA9oBhMTrcT8dk0/H/Q5gAm2X/bAu9XK28lv9nvJ4CBi2Ykgkq/bh+
# bOJzLG3MmSYX7jagGV1TsfFOSCWzItCXA9sKLADZFlDYBd4=
# SIG # End signature block
