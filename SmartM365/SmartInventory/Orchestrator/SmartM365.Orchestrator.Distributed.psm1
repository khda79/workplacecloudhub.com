Set-StrictMode -Version Latest

$script:DistributedModuleVersion = '1.1.2'

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
        [int]$Depth = 12
    )

    $parent = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $parent -Force)
    }

    $temporaryPath = '{0}.{1}.{2}.tmp' -f $Path, $PID, [guid]::NewGuid().ToString('N')
    try {
        $json = $Value | ConvertTo-Json -Depth $Depth
        [System.IO.File]::WriteAllText($temporaryPath, $json, [System.Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
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
finally { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null }
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
    foreach ($component in @(Get-DependencyComponent -Jobs $electedJobs)) {
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
    $claimPath = Join-Path -Path $jobFolder -ChildPath ($occurrenceUtc.ToString('yyyyMMddTHHmmssfffZ') + '.json')
    if (-not (Test-Path -LiteralPath $claimPath -PathType Leaf)) { return $null }
    $claim = Get-Content -LiteralPath $claimPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
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
    $claimPath = Join-Path -Path $jobFolder -ChildPath ($occurrenceUtc.ToString('yyyyMMddTHHmmssfffZ') + '.json')
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
        $stream = [System.IO.File]::Open($claimPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
        try {
            $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($json)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Flush($true)
        }
        finally { $stream.Dispose() }
        return [pscustomobject]@{ Acquired = $true; Reused = $false; ClaimPath = $claimPath; Claim = [pscustomobject]$claim; Reason = '' }
    }
    catch [System.IO.IOException] {
        try {
            $existing = Get-Content -LiteralPath $claimPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $sameOwner = [string]$existing.OwnerServer -eq $OwnerServer.ToUpperInvariant()
            $sameOrchestratorProcess = $false
            try { $sameOrchestratorProcess = [int]$existing.OrchestratorPid -eq $PID } catch { $sameOrchestratorProcess = $false }
            $reusable = [string]$existing.Status -eq 'RetryScheduled' -or
                ([string]$existing.Status -eq 'Claimed' -and $sameOrchestratorProcess)
            if ($sameOwner -and $reusable) {
                return [pscustomobject]@{ Acquired = $true; Reused = $true; ClaimPath = $claimPath; Claim = $existing; Reason = 'Existing non-terminal claim belongs to this server.' }
            }

            $terminal = [string]$existing.Status -in @('Success', 'Failed', 'TimedOut', 'Interrupted')
            $safeUntilUtc = [datetime]::MaxValue
            try { $safeUntilUtc = ([datetime]$existing.SafeUntilUtc).ToUniversalTime() } catch { $safeUntilUtc = [datetime]::MaxValue }
            if (-not $terminal -and [datetime]::UtcNow -gt $safeUntilUtc -and -not [string]::IsNullOrWhiteSpace($HeartbeatRootPath)) {
                $heartbeatFresh = $false
                try {
                    $heartbeatPath = Join-Path -Path (Join-Path -Path $HeartbeatRootPath -ChildPath ([string]$existing.OwnerServer)) -ChildPath 'Orchestrator-Heartbeat.json'
                    $heartbeat = Get-Content -LiteralPath $heartbeatPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    $heartbeatAgeMinutes = ([datetime]::UtcNow - ([datetime]$heartbeat.Timestamp).ToUniversalTime()).TotalMinutes
                    $heartbeatFresh = $heartbeatAgeMinutes -le [math]::Max(1, $HeartbeatStaleMinutes)
                }
                catch { $heartbeatFresh = $false }
                # A restarted orchestrator on the same server may replace its own
                # expired claim even though its new heartbeat is healthy.
                if ($sameOwner) { $heartbeatFresh = $false }

                if (-not $heartbeatFresh) {
                    $takeoverLockPath = $claimPath + '.takeover.lock'
                    $takeoverStream = $null
                    try {
                        $takeoverStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                        $confirmed = Get-Content -LiteralPath $claimPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                        $confirmedSafeUntilUtc = ([datetime]$confirmed.SafeUntilUtc).ToUniversalTime()
                        if ([string]$confirmed.ClaimId -eq [string]$existing.ClaimId -and
                            [string]$confirmed.Status -notin @('Success', 'Failed', 'TimedOut', 'Interrupted') -and
                            [datetime]::UtcNow -gt $confirmedSafeUntilUtc) {
                            $archivePath = '{0}.stale.{1}.json' -f $claimPath, [string]$confirmed.ClaimId
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
        [Parameter(Mandatory = $true)][ValidateSet('Claimed', 'Running', 'RetryScheduled', 'Success', 'Failed', 'TimedOut', 'Interrupted')][string]$Status,
        [int]$Attempt = 0,
        [int]$SafeMinutes = 60,
        [string]$Detail = ''
    )

    $claim = Get-Content -LiteralPath $ClaimPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
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
        Write-JsonAtomically -Path $ClaimPath -Value $claim
    }
    return $claim
}

function Get-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasesRootPath,
        [Parameter(Mandatory = $true)][string]$ConcurrencyKey
    )

    $leasePath = Join-Path -Path $LeasesRootPath -ChildPath ((ConvertTo-SafeFileName $ConcurrencyKey) + '.json')
    if (-not (Test-Path -LiteralPath $leasePath -PathType Leaf)) { return $null }
    $lease = Get-Content -LiteralPath $leasePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
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
    $leasePath = Join-Path -Path $LeasesRootPath -ChildPath ((ConvertTo-SafeFileName $ConcurrencyKey) + '.json')
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
        $stream = [System.IO.File]::Open($leasePath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
        try {
            $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($json)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Flush($true)
        }
        finally { $stream.Dispose() }
        return [pscustomobject]@{ Acquired = $true; Reused = $false; LeasePath = $leasePath; Lease = [pscustomobject]$lease; Reason = '' }
    }
    catch [System.IO.IOException] {
        try {
            $existing = Get-Content -LiteralPath $leasePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
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
                        $heartbeat = Get-Content -LiteralPath $heartbeatPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                        $heartbeatAgeMinutes = ([datetime]::UtcNow - ([datetime]$heartbeat.Timestamp).ToUniversalTime()).TotalMinutes
                        $heartbeatFresh = $heartbeatAgeMinutes -le [math]::Max(1, $HeartbeatStaleMinutes)
                    }
                    catch { $heartbeatFresh = $false }
                }
                if ($sameOwner) { $heartbeatFresh = $false }

                if (-not $heartbeatFresh) {
                    $takeoverLockPath = $leasePath + '.takeover.lock'
                    $takeoverStream = $null
                    try {
                        if ((Test-Path -LiteralPath $takeoverLockPath -PathType Leaf) -and
                            ([datetime]::UtcNow - (Get-Item -LiteralPath $takeoverLockPath).LastWriteTimeUtc).TotalMinutes -gt 5) {
                            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
                        }
                        $takeoverStream = [System.IO.File]::Open($takeoverLockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                        $confirmed = Get-Content -LiteralPath $leasePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                        $confirmedSafeUntilUtc = ([datetime]$confirmed.SafeUntilUtc).ToUniversalTime()
                        if ([string]$confirmed.LeaseId -eq [string]$existing.LeaseId -and [datetime]::UtcNow -gt $confirmedSafeUntilUtc) {
                            $archivePath = '{0}.stale.{1}.json' -f $leasePath, [string]$confirmed.LeaseId
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

function Exit-SmartM365OrchestratorConcurrencyLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LeasePath,
        [Parameter(Mandatory = $true)][string]$LeaseId,
        [Parameter(Mandatory = $true)][string]$OwnerServer
    )

    if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $true }
    $takeoverLockPath = $LeasePath + '.takeover.lock'
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
        $current = Get-Content -LiteralPath $LeasePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ([string]$current.LeaseId -ne $LeaseId -or [string]$current.OwnerServer -ne $OwnerServer.ToUpperInvariant()) {
            return $false
        }
        Remove-Item -LiteralPath $LeasePath -Force -ErrorAction Stop
        return $true
    }
    finally {
        if ($null -ne $lockStream) {
            $lockStream.Dispose()
            Remove-Item -LiteralPath $takeoverLockPath -Force -ErrorAction SilentlyContinue
        }
    }
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
    'Exit-SmartM365OrchestratorConcurrencyLease'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA41+zveZuiCRk2
# D14Pe3GHUs/VX/vLUtyjY9dimPC/tKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIG1Z5wT9sinBpedqTVznPJNaHFZfRE/o6zDJG3V6brzuMA0GCSqG
# SIb3DQEBAQUABIIBgHuFhv+s/th3wd+lR8VdXcDm/qeJpxjGoYiDd10TQQh4cxwQ
# 3YKXOP9IkT7mq3qiFhdppQT18h9JF4Hf/hrykLenrt1kgDDaXA/oTrorXPh4Ds+p
# iW/gMral7cfNK3A57lbxIb1yitczoFb8m/TpOv9s0B8Jkw8lgk+ayRWkkLutB+yP
# oNjDCa9+eh+4Ovv3XOP3tSkWAav7V89Kj/5Kss6pqBKu7p/L5a5snuLEfS0t5Z4/
# WRZvRnDBm3T5pPowedEgRbm1n+Jic6aOn2Vg8jYF+8wgce0yrZEWiLV0V+1fVbj9
# jyHkYR+1tzAFxQdNjKl7G//56LBTigqrTqvk4PuTXYOVYv32Lzr4vUyDWy7CV+XS
# 1eFcMt7eVtzwl49y2d3SSmvPKbhlN9PPowT+HEpWxSwb/U144vJPTsPqdMG9Eb82
# cWao2dK6/z9Lhu38OBs/V6lkwpcsgtPbBQ0ko4q0SeB4FKe+gko3NDx70ZqVPtTs
# 9aO60PpEQvnxM2MiGaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjEwODM4
# MDVaMC8GCSqGSIb3DQEJBDEiBCBmXqJ7+cVw9TC8oTiLl4UnMghnQx0uC9uHrNHD
# BwB6OTANBgkqhkiG9w0BAQEFAASCAgAokNUftX9T10icG1GYlpb5u2ZPQj6RUTBK
# JXbq5T1+w2QZOqJJUiS8t8PvbFg1Oa7EteIk15J6avdqEswo0ccNBXsP3Qx7twOk
# 3rHECNPClDu6ycug/8ifSTmo9UDr+JqwJL9CyTUPTUryp5LMemZa/qSb8krCUrDr
# h2d+I9nMTzjEQ5ax+JElAZGi0OVQqJJvCKKm7xeKLofApmMUIWPemQBeAkXOwbcN
# dzgzNM6cNbFoy2dEa7pMZ8siCPoXVl+/xgok4ikrGsoSSmGUnUnVPtfKO5wC7WeN
# BErMIeqo6kCeLVmL3Xk7qoO1/SW3i8qBP2IYmYrQH2of9V7/TrJ3UWJmeBElG2Ka
# yYjWAvvt2osI+D9BHSXF6U4dksTzVE8sjTtIE8qBheQMzgWps7M5RTJFpucjK2Iy
# 25OElQJ1wNh6JyWrU2obGaY5byVRRZl9T7Ih03MeXiVJJfGEnv8DrWp/qMHjvBMQ
# mXtd4MJkzKKFXgt5AwdEYyBBi0TkAI4rA0iCNqe4HUoMzWzu0pRc4uqvBTPdayU/
# GdPU4jFCKwO8WMbYyyvA8WTF3SM4GFPebtrEasBXpapLTSu7462KLw+HRx9ZuMQ6
# gM2KOmnlbP8oGLPABA3cj/ElzwS3lQfiT2OYWqwcYi8Zml3Pl3XUG4dBK/hGGaYa
# rHq6Jj1lwQ==
# SIG # End signature block
