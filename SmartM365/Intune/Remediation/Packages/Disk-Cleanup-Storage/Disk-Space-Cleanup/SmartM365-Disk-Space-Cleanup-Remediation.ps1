<#
.SYNOPSIS
    Frees system drive space with controlled cleanup levels without forcing a reboot.

.VERSION
    1.8
#>
# Name: SmartM365-Disk-Space-Cleanup-Remediation.ps1
# Version: 1.8
# Description: Frees system drive space with controlled cleanup levels without forcing a reboot.

$ErrorActionPreference = "Stop"

$ScriptName = "SmartM365-Disk-Space-Cleanup-Remediation"
$Version = "1.8"
$Scenario = "Disk-Space-Cleanup"
$MinimumTargetFreeSpaceGB = 50
$Windows10Only = $true
$CleanupLevel = "Moderate"
$WindowsOldMinimumAgeDays = 14
$WindowsSetupResidueMinimumAgeDays = 14
$SmartM365RuntimeCleanupMinimumAgeDays = 14
$CustomDirectoryMinimumAgeDays = 0
$BrowserCacheMinimumAgeDays = 0
$LogRetentionCount = 10
$IncludeInstallerPatchCache = $false
$SystemDrive = $env:SystemDrive
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$RunTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = Join-Path -Path $LogRoot -ChildPath ("{0}-Remediation_{1}.log" -f $Scenario, $RunTimestamp)

$CustomCleanupDirectories = @(
    @{ Label = "BMCClientPatchDownload"; Path = "C:\Program Files\BMC Software\Client Management\Client\data\FileStore\downstream\PatchDownload"; RequiredLevel = "Moderate"; DeleteRoot = $false }
    # Example:
    # @{ Label = "CustomCache"; Path = "C:\ProgramData\Vendor\Cache"; RequiredLevel = "Moderate"; DeleteRoot = $false }
)

$script:ActionsPerformed = New-Object System.Collections.Generic.List[string]
$script:SkippedActions = New-Object System.Collections.Generic.List[string]
$script:StoppedUpdateServices = New-Object System.Collections.Generic.List[string]

function ConvertTo-SingleLineValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return "Unknown"
    }

    $text = [string]$Value
    $text = $text -replace '[\r\n\t]+', ' '
    $text = $text -replace '\s{2,}', ' '
    return $text.Trim()
}

function Write-IntuneResult {
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Values)

    $parts = New-Object System.Collections.Generic.List[string]

    foreach ($key in $Values.Keys) {
        $parts.Add(("{0}={1}" -f $key, (ConvertTo-SingleLineValue -Value $Values[$key])))
    }

    Write-Output ($parts -join " ")
}

function Convert-NumberToInvariantText {
    param([Parameter(Mandatory = $true)][double]$Value)
    return $Value.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Join-CompactList {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Items,
        [int]$MaximumItems = 8
    )

    if ($Items.Count -eq 0) {
        return "None"
    }

    $selected = @($Items | Select-Object -First $MaximumItems)

    if ($Items.Count -gt $MaximumItems) {
        $selected += ("More:{0}" -f ($Items.Count - $MaximumItems))
    }

    return ($selected -join ",")
}

function Initialize-LogFile {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null

    if ($LogRetentionCount -gt 0) {
        Get-ChildItem -LiteralPath $LogRoot -Filter "$Scenario-Remediation_*.log" -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -Skip $LogRetentionCount |
            ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                }
                catch {
                    $null = $_
                }
            }
    }
}

function Write-SmartM365Log {
    param([Parameter(Mandatory = $true)][string]$Message)

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, (ConvertTo-SingleLineValue -Value $Message)
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

function Get-SystemDriveFreeSpaceGB {
    try {
        if ([string]::IsNullOrWhiteSpace($SystemDrive)) {
            return $null
        }

        $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$SystemDrive'" -ErrorAction Stop

        if ($null -eq $drive -or $drive.DriveType -ne 3) {
            return $null
        }

        return [math]::Round(($drive.FreeSpace / 1GB), 2)
    }
    catch {
        return $null
    }
}

function Test-TargetFreeSpaceReached {
    $freeSpaceGB = Get-SystemDriveFreeSpaceGB

    if ($null -eq $freeSpaceGB) {
        return $false
    }

    return ($freeSpaceGB -ge $MinimumTargetFreeSpaceGB)
}

function Get-CleanupLevelValue {
    param([Parameter(Mandatory = $true)][string]$Level)

    switch ($Level) {
        "Safe" { return 1 }
        "Moderate" { return 2 }
        "Aggressive" { return 3 }
        default { return 2 }
    }
}

function Test-CleanupLevelAllowed {
    param([Parameter(Mandatory = $true)][string]$RequiredLevel)
    return ((Get-CleanupLevelValue -Level $CleanupLevel) -ge (Get-CleanupLevelValue -Level $RequiredLevel))
}

function Test-SetupOrUpdateProcessActive {
    foreach ($processName in @("setup", "setuphost", "setupprep", "dism", "tiworker", "trustedinstaller", "usoclient")) {
        if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
            return $true
        }
    }

    return $false
}

function Invoke-UpdateServiceStop {
    foreach ($serviceName in @("wuauserv", "bits", "dosvc")) {
        try {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

            if ($null -ne $service -and $service.Status -eq "Running") {
                Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                $script:StoppedUpdateServices.Add($serviceName)
                Write-SmartM365Log "ServiceStopRequested=$serviceName"
            }
        }
        catch {
            Write-SmartM365Log "ServiceStopFailed=$serviceName Message=$($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 2
}

function Invoke-UpdateServiceStart {
    foreach ($serviceName in @($script:StoppedUpdateServices | Select-Object -Unique)) {
        try {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

            if ($null -ne $service -and $service.Status -ne "Running") {
                Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                Write-SmartM365Log "ServiceStartRequested=$serviceName"
            }
        }
        catch {
            Write-SmartM365Log "ServiceStartFailed=$serviceName Message=$($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 2
}

function Clear-FolderContent {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-SmartM365Log "CleanupSkipped=$Label Reason=EmptyPath"
        return
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-SmartM365Log "CleanupSkipped=$Label Reason=PathNotFound"
        return
    }

    try {
        Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-SmartM365Log "CleanupItemSkipped=$Label Message=$($_.Exception.Message)"
            }
        }

        Write-SmartM365Log "CleanupCompleted=$Label Path=$Path"
    }
    catch {
        Write-SmartM365Log "CleanupPartial=$Label Message=$($_.Exception.Message)"
    }
}

function Clear-OldFolderContent {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][int]$MinimumAgeDays,
        [bool]$DeleteRoot = $false
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-SmartM365Log "CleanupSkipped=$Label Reason=EmptyPath"
        return
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-SmartM365Log "CleanupSkipped=$Label Reason=PathNotFound"
        return
    }

    $cutoff = (Get-Date).AddDays(-1 * [math]::Abs($MinimumAgeDays))

    if ($DeleteRoot) {
        try {
            $rootItem = Get-Item -LiteralPath $Path -ErrorAction Stop

            if ($rootItem.LastWriteTime -le $cutoff) {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
                Write-SmartM365Log "CleanupCompleted=$Label Mode=DeleteRoot MinimumAgeDays=$MinimumAgeDays Path=$Path"
                return
            }

            Write-SmartM365Log "CleanupSkipped=$Label Reason=RootTooRecent MinimumAgeDays=$MinimumAgeDays Path=$Path"
            return
        }
        catch {
            Write-SmartM365Log "CleanupPartial=$Label Mode=DeleteRoot Message=$($_.Exception.Message)"
            return
        }
    }

    $removedFiles = 0
    $removedDirectories = 0

    try {
        Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -le $cutoff } |
            ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                    $removedFiles++
                }
                catch {
                    Write-SmartM365Log "CleanupItemSkipped=$Label Message=$($_.Exception.Message)"
                }
            }

        Get-ChildItem -LiteralPath $Path -Recurse -Force -Directory -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            ForEach-Object {
                try {
                    $hasChildren = @(Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue).Count -gt 0

                    if (-not $hasChildren -and $_.LastWriteTime -le $cutoff) {
                        Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                        $removedDirectories++
                    }
                }
                catch {
                    Write-SmartM365Log "CleanupDirectorySkipped=$Label Message=$($_.Exception.Message)"
                }
            }

        Write-SmartM365Log "CleanupCompleted=$Label Mode=OldContent MinimumAgeDays=$MinimumAgeDays RemovedFiles=$removedFiles RemovedDirectories=$removedDirectories Path=$Path"
    }
    catch {
        Write-SmartM365Log "CleanupPartial=$Label Message=$($_.Exception.Message)"
    }
}

function Clear-UserTempContent {
    $profileCount = 0

    try {
        $usersRoot = Join-Path -Path $SystemDrive -ChildPath "Users"

        if (-not (Test-Path -LiteralPath $usersRoot)) {
            Write-SmartM365Log "CleanupSkipped=UserTemp Reason=UsersRootNotFound"
            return
        }

        $profileFolders = Get-ChildItem -LiteralPath $usersRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin @("Default", "Default User", "Public", "All Users") }

        foreach ($profileFolder in $profileFolders) {
            $userTemp = Join-Path -Path $profileFolder.FullName -ChildPath "AppData\Local\Temp"

            if (Test-Path -LiteralPath $userTemp) {
                $profileCount++
                Get-ChildItem -LiteralPath $userTemp -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    try {
                        Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-SmartM365Log "CleanupItemSkipped=UserTemp Message=$($_.Exception.Message)"
                    }
                }
            }
        }

        Write-SmartM365Log "CleanupCompleted=UserTemp ProfileTempFolders=$profileCount"
    }
    catch {
        Write-SmartM365Log "CleanupPartial=UserTemp Message=$($_.Exception.Message)"
    }
}

function Get-BrowserCacheDefinition {
    param([Parameter(Mandatory = $true)][string]$Browser)

    switch ($Browser) {
        "Edge" {
            return [pscustomobject]@{
                ProfileRootRelativePath = "AppData\Local\Microsoft\Edge\User Data"
                CacheRelativePaths = @("Cache", "Code Cache", "GPUCache", "Service Worker\CacheStorage", "DawnCache", "ShaderCache", "GrShaderCache", "Media Cache")
            }
        }
        "Chrome" {
            return [pscustomobject]@{
                ProfileRootRelativePath = "AppData\Local\Google\Chrome\User Data"
                CacheRelativePaths = @("Cache", "Code Cache", "GPUCache", "Service Worker\CacheStorage", "DawnCache", "ShaderCache", "GrShaderCache", "Media Cache")
            }
        }
        "Firefox" {
            return [pscustomobject]@{
                ProfileRootRelativePath = "AppData\Local\Mozilla\Firefox\Profiles"
                CacheRelativePaths = @("cache2", "startupCache", "shader-cache")
            }
        }
        default {
            return $null
        }
    }
}

function Clear-BrowserUserCacheContent {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][string]$Browser,
        [Parameter(Mandatory = $true)][string]$ProcessName
    )

    $userProfileCount = 0
    $browserProfileCount = 0
    $cacheFolderCount = 0
    $removedFiles = 0
    $removedDirectories = 0
    $skippedItems = 0
    $browserDefinition = Get-BrowserCacheDefinition -Browser $Browser

    if ($null -eq $browserDefinition) {
        $script:SkippedActions.Add("${Label}:UnsupportedBrowser")
        Write-SmartM365Log "CleanupSkipped=$Label Reason=UnsupportedBrowser Browser=$Browser"
        return
    }

    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
        $script:SkippedActions.Add("${Label}:ProcessActive")
        Write-SmartM365Log "CleanupSkipped=$Label Reason=BrowserProcessActive ProcessName=$ProcessName"
        return
    }

    try {
        $usersRoot = Join-Path -Path $SystemDrive -ChildPath "Users"

        if (-not (Test-Path -LiteralPath $usersRoot)) {
            Write-SmartM365Log "CleanupSkipped=$Label Reason=UsersRootNotFound"
            return
        }

        $cutoff = (Get-Date).AddDays(-1 * [math]::Abs($BrowserCacheMinimumAgeDays))
        $windowsProfiles = Get-ChildItem -LiteralPath $usersRoot -Force -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin @("Default", "Default User", "Public", "All Users") }

        foreach ($windowsProfile in $windowsProfiles) {
            $userProfileCount++
            $browserProfileRoot = Join-Path -Path $windowsProfile.FullName -ChildPath $browserDefinition.ProfileRootRelativePath

            if (-not (Test-Path -LiteralPath $browserProfileRoot)) {
                continue
            }

            $browserProfiles = Get-ChildItem -LiteralPath $browserProfileRoot -Force -Directory -ErrorAction SilentlyContinue

            foreach ($browserProfile in $browserProfiles) {
                $browserProfileHasCache = $false

                foreach ($cacheRelativePath in $browserDefinition.CacheRelativePaths) {
                    $cachePath = Join-Path -Path $browserProfile.FullName -ChildPath $cacheRelativePath

                    if (-not (Test-Path -LiteralPath $cachePath)) {
                        continue
                    }

                    $cacheFolderCount++
                    $browserProfileHasCache = $true

                    Get-ChildItem -LiteralPath $cachePath -Recurse -Force -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -le $cutoff } |
                        ForEach-Object {
                            try {
                                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                                $removedFiles++
                            }
                            catch {
                                $skippedItems++
                            }
                        }

                    Get-ChildItem -LiteralPath $cachePath -Recurse -Force -Directory -ErrorAction SilentlyContinue |
                        Sort-Object FullName -Descending |
                        ForEach-Object {
                            try {
                                $hasChildren = @(Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue).Count -gt 0

                                if (-not $hasChildren -and $_.LastWriteTime -le $cutoff) {
                                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                                    $removedDirectories++
                                }
                            }
                            catch {
                                $skippedItems++
                            }
                        }
                }

                if ($browserProfileHasCache) {
                    $browserProfileCount++
                }
            }
        }

        Write-SmartM365Log "CleanupCompleted=$Label Browser=$Browser UserProfilesScanned=$userProfileCount BrowserProfilesWithCache=$browserProfileCount CacheFolders=$cacheFolderCount RemovedFiles=$removedFiles RemovedDirectories=$removedDirectories SkippedItems=$skippedItems MinimumAgeDays=$BrowserCacheMinimumAgeDays"
    }
    catch {
        Write-SmartM365Log "CleanupPartial=$Label Browser=$Browser ErrorType=$($_.Exception.GetType().Name)"
    }
}

function Clear-RecycleBinSafe {
    try {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue | Out-Null
        Write-SmartM365Log "RecycleBinCleanupCompleted"
    }
    catch {
        Write-SmartM365Log "RecycleBinCleanupSkipped Message=$($_.Exception.Message)"
    }
}

function Invoke-DismComponentCleanup {
    try {
        $process = Start-Process -FilePath "dism.exe" -ArgumentList "/Online", "/Cleanup-Image", "/StartComponentCleanup" -Wait -PassThru -WindowStyle Hidden
        Write-SmartM365Log "DismComponentCleanupExitCode=$($process.ExitCode)"
    }
    catch {
        Write-SmartM365Log "DismComponentCleanupFailed Message=$($_.Exception.Message)"
    }
}

function Test-DirectoryAgeAllowed {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][int]$MinimumAgeDays
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-SmartM365Log "CleanupSkipped=$Label Reason=PathNotFound"
        return $false
    }

    try {
        $item = Get-Item -LiteralPath $Path -ErrorAction Stop
        $ageDays = ((Get-Date) - $item.LastWriteTime).TotalDays
        Write-SmartM365Log ("{0}AgeDays={1}" -f $Label, ([math]::Round($ageDays, 1)))
        return ($ageDays -ge $MinimumAgeDays)
    }
    catch {
        Write-SmartM365Log "DirectoryAgeCheckFailed=$Label Message=$($_.Exception.Message)"
        return $false
    }
}

function Clear-SmartM365OldRuntimeContent {
    $runtimeRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation"
    $runtimeTargets = @(
        @{ Label = "SmartM365Logs"; Path = (Join-Path -Path $runtimeRoot -ChildPath "Logs") },
        @{ Label = "SmartM365Temp"; Path = (Join-Path -Path $runtimeRoot -ChildPath "Temp") },
        @{ Label = "SmartM365Cache"; Path = (Join-Path -Path $runtimeRoot -ChildPath "Cache") }
    )

    foreach ($target in $runtimeTargets) {
        Clear-OldFolderContent -Path $target.Path -Label $target.Label -MinimumAgeDays $SmartM365RuntimeCleanupMinimumAgeDays
    }
}

function Invoke-CustomCleanupDirectoryList {
    foreach ($customDirectory in $CustomCleanupDirectories) {
        if (-not $customDirectory.ContainsKey("Path") -or [string]::IsNullOrWhiteSpace([string]$customDirectory.Path)) {
            continue
        }

        $label = "CustomDirectory"
        $requiredLevel = "Moderate"
        $minimumAgeDays = $CustomDirectoryMinimumAgeDays
        $deleteRoot = $false

        if ($customDirectory.ContainsKey("Label") -and -not [string]::IsNullOrWhiteSpace([string]$customDirectory.Label)) {
            $label = [string]$customDirectory.Label
        }

        if ($customDirectory.ContainsKey("RequiredLevel") -and -not [string]::IsNullOrWhiteSpace([string]$customDirectory.RequiredLevel)) {
            $requiredLevel = [string]$customDirectory.RequiredLevel
        }

        if ($customDirectory.ContainsKey("MinimumAgeDays")) {
            $minimumAgeDays = [int]$customDirectory.MinimumAgeDays
        }

        if ($customDirectory.ContainsKey("DeleteRoot")) {
            $deleteRoot = [bool]$customDirectory.DeleteRoot
        }

        Invoke-CleanupStep -Label $label -RequiredLevel $requiredLevel -Action {
            Clear-OldFolderContent -Path ([string]$customDirectory.Path) -Label $label -MinimumAgeDays $minimumAgeDays -DeleteRoot $deleteRoot
        }
    }
}

function Invoke-CleanupStep {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][string]$RequiredLevel,
        [Parameter(Mandatory = $true)][scriptblock]$Action
    )

    if (-not (Test-CleanupLevelAllowed -RequiredLevel $RequiredLevel)) {
        $script:SkippedActions.Add("${Label}:Level")
        Write-SmartM365Log "CleanupSkipped=$Label Reason=CleanupLevel RequiredLevel=$RequiredLevel CurrentLevel=$CleanupLevel"
        return
    }

    if (Test-TargetFreeSpaceReached) {
        $script:SkippedActions.Add("${Label}:TargetReached")
        Write-SmartM365Log "CleanupSkipped=$Label Reason=TargetFreeSpaceAlreadyReached"
        return
    }

    $beforeGB = Get-SystemDriveFreeSpaceGB
    $skippedBefore = $script:SkippedActions.Count
    Write-SmartM365Log "CleanupStarted=$Label FreeSpaceGBBefore=$beforeGB"

    try {
        & $Action
    }
    catch {
        Write-SmartM365Log "CleanupFailed=$Label Message=$($_.Exception.Message)"
    }

    $afterGB = Get-SystemDriveFreeSpaceGB
    $gainGB = 0

    if ($null -ne $beforeGB -and $null -ne $afterGB) {
        $gainGB = [math]::Round(($afterGB - $beforeGB), 2)
    }

    Write-SmartM365Log "CleanupFinished=$Label FreeSpaceGBAfter=$afterGB GainGB=$gainGB"

    if ($script:SkippedActions.Count -eq $skippedBefore) {
        $script:ActionsPerformed.Add(("{0}:{1}GB" -f $Label, (Convert-NumberToInvariantText -Value $gainGB)))
    }
}

try {
    Initialize-LogFile

    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

    if ($Windows10Only -and $operatingSystem.Caption -notmatch "Windows 10") {
        Write-IntuneResult -Values ([ordered]@{ Status = "NotApplicable"; Reason = "NotWindows10"; Script = $ScriptName; Version = $Version })
        exit 0
    }

    Write-SmartM365Log "RemediationStarted Script=$ScriptName Version=$Version"
    Write-SmartM365Log "SystemDrive=$SystemDrive"
    Write-SmartM365Log "MinimumTargetFreeSpaceGB=$MinimumTargetFreeSpaceGB"
    Write-SmartM365Log "CleanupLevel=$CleanupLevel"
    Write-SmartM365Log "WindowsOldMinimumAgeDays=$WindowsOldMinimumAgeDays"
    Write-SmartM365Log "WindowsSetupResidueMinimumAgeDays=$WindowsSetupResidueMinimumAgeDays"
    Write-SmartM365Log "SmartM365RuntimeCleanupMinimumAgeDays=$SmartM365RuntimeCleanupMinimumAgeDays"
    Write-SmartM365Log "BrowserCacheMinimumAgeDays=$BrowserCacheMinimumAgeDays"

    $beforeFreeGB = Get-SystemDriveFreeSpaceGB
    Write-SmartM365Log "FreeSpaceGBBefore=$beforeFreeGB"

    if ($null -eq $beforeFreeGB) {
        Write-SmartM365Log "Status=Failed Reason=FreeSpaceUnknownBeforeCleanup"
        Write-IntuneResult -Values ([ordered]@{ Status = "Failed"; Reason = "FreeSpaceUnknown"; RequiredFreeSpaceGB = $MinimumTargetFreeSpaceGB; LogPath = $LogPath; Script = $ScriptName; Version = $Version })
        exit 1
    }

    if ($beforeFreeGB -ge $MinimumTargetFreeSpaceGB) {
        Write-SmartM365Log "Status=AlreadyReady"
        Write-IntuneResult -Values ([ordered]@{ Status = "AlreadyReady"; FreeSpaceGBBefore = $beforeFreeGB; RequiredFreeSpaceGB = $MinimumTargetFreeSpaceGB; LogPath = $LogPath; Script = $ScriptName; Version = $Version })
        exit 0
    }

    $setupOrUpdateActive = Test-SetupOrUpdateProcessActive
    Write-SmartM365Log "SetupOrUpdateProcessActive=$setupOrUpdateActive"

    Invoke-CleanupStep -Label "WindowsTemp" -RequiredLevel "Safe" -Action {
        Clear-FolderContent -Path (Join-Path -Path $env:WINDIR -ChildPath "Temp") -Label "WindowsTemp"
    }

    Invoke-CleanupStep -Label "WindowsUpdateDownloadCache" -RequiredLevel "Safe" -Action {
        if ($setupOrUpdateActive) {
            $script:SkippedActions.Add("WindowsUpdateDownloadCache:UpdateActive")
            Write-SmartM365Log "CleanupSkipped=WindowsUpdateDownloadCache Reason=SetupOrUpdateProcessActive"
            return
        }

        Invoke-UpdateServiceStop
        Clear-FolderContent -Path (Join-Path -Path $env:WINDIR -ChildPath "SoftwareDistribution\Download") -Label "WindowsUpdateDownloadCache"
    }

    Invoke-CleanupStep -Label "DeliveryOptimizationCache" -RequiredLevel "Safe" -Action {
        if ($setupOrUpdateActive) {
            $script:SkippedActions.Add("DeliveryOptimizationCache:UpdateActive")
            Write-SmartM365Log "CleanupSkipped=DeliveryOptimizationCache Reason=SetupOrUpdateProcessActive"
            return
        }

        Invoke-UpdateServiceStop
        Clear-FolderContent -Path (Join-Path -Path $env:WINDIR -ChildPath "ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache") -Label "DeliveryOptimizationCache"
    }

    Invoke-CleanupStep -Label "UserTemp" -RequiredLevel "Moderate" -Action { Clear-UserTempContent }
    Invoke-CleanupStep -Label "EdgeUserCache" -RequiredLevel "Moderate" -Action { Clear-BrowserUserCacheContent -Label "EdgeUserCache" -Browser "Edge" -ProcessName "msedge" }
    Invoke-CleanupStep -Label "ChromeUserCache" -RequiredLevel "Moderate" -Action { Clear-BrowserUserCacheContent -Label "ChromeUserCache" -Browser "Chrome" -ProcessName "chrome" }
    Invoke-CleanupStep -Label "FirefoxUserCache" -RequiredLevel "Moderate" -Action { Clear-BrowserUserCacheContent -Label "FirefoxUserCache" -Browser "Firefox" -ProcessName "firefox" }
    Invoke-CleanupStep -Label "RecycleBin" -RequiredLevel "Moderate" -Action { Clear-RecycleBinSafe }

    Invoke-CleanupStep -Label "WindowsOld" -RequiredLevel "Moderate" -Action {
        $windowsOldPath = Join-Path -Path $SystemDrive -ChildPath "Windows.old"

        if ($setupOrUpdateActive) {
            $script:SkippedActions.Add("WindowsOld:UpdateActive")
            Write-SmartM365Log "CleanupSkipped=WindowsOld Reason=SetupOrUpdateProcessActive"
            return
        }

        if (-not (Test-DirectoryAgeAllowed -Path $windowsOldPath -Label "WindowsOld" -MinimumAgeDays $WindowsOldMinimumAgeDays)) {
            $script:SkippedActions.Add("WindowsOld:Age")
            Write-SmartM365Log "CleanupSkipped=WindowsOld Reason=MinimumAgeNotReached MinimumAgeDays=$WindowsOldMinimumAgeDays"
            return
        }

        Clear-FolderContent -Path $windowsOldPath -Label "WindowsOld"
    }

    Invoke-CleanupStep -Label "WindowsSetupBT" -RequiredLevel "Moderate" -Action {
        $windowsSetupBtPath = Join-Path -Path $SystemDrive -ChildPath '$WINDOWS.~BT'

        if ($setupOrUpdateActive) {
            $script:SkippedActions.Add("WindowsSetupBT:UpdateActive")
            Write-SmartM365Log "CleanupSkipped=WindowsSetupBT Reason=SetupOrUpdateProcessActive"
            return
        }

        if (-not (Test-DirectoryAgeAllowed -Path $windowsSetupBtPath -Label "WindowsSetupBT" -MinimumAgeDays $WindowsSetupResidueMinimumAgeDays)) {
            $script:SkippedActions.Add("WindowsSetupBT:Age")
            Write-SmartM365Log "CleanupSkipped=WindowsSetupBT Reason=MinimumAgeNotReached MinimumAgeDays=$WindowsSetupResidueMinimumAgeDays"
            return
        }

        Clear-FolderContent -Path $windowsSetupBtPath -Label "WindowsSetupBT"
    }

    Invoke-CleanupStep -Label "WindowsSetupWS" -RequiredLevel "Moderate" -Action {
        $windowsSetupWsPath = Join-Path -Path $SystemDrive -ChildPath '$WINDOWS.~WS'

        if ($setupOrUpdateActive) {
            $script:SkippedActions.Add("WindowsSetupWS:UpdateActive")
            Write-SmartM365Log "CleanupSkipped=WindowsSetupWS Reason=SetupOrUpdateProcessActive"
            return
        }

        if (-not (Test-DirectoryAgeAllowed -Path $windowsSetupWsPath -Label "WindowsSetupWS" -MinimumAgeDays $WindowsSetupResidueMinimumAgeDays)) {
            $script:SkippedActions.Add("WindowsSetupWS:Age")
            Write-SmartM365Log "CleanupSkipped=WindowsSetupWS Reason=MinimumAgeNotReached MinimumAgeDays=$WindowsSetupResidueMinimumAgeDays"
            return
        }

        Clear-FolderContent -Path $windowsSetupWsPath -Label "WindowsSetupWS"
    }

    Invoke-CleanupStep -Label "SmartM365RuntimeOldContent" -RequiredLevel "Moderate" -Action {
        Clear-SmartM365OldRuntimeContent
    }

    Invoke-CustomCleanupDirectoryList
    Invoke-CleanupStep -Label "DismComponentCleanup" -RequiredLevel "Moderate" -Action { Invoke-DismComponentCleanup }

    Invoke-CleanupStep -Label "InstallerPatchCache" -RequiredLevel "Aggressive" -Action {
        if (-not $IncludeInstallerPatchCache) {
            $script:SkippedActions.Add("InstallerPatchCache:Disabled")
            Write-SmartM365Log "CleanupSkipped=InstallerPatchCache Reason=SensitiveTargetDisabled"
            return
        }

        Clear-FolderContent -Path (Join-Path -Path $env:WINDIR -ChildPath 'Installer\$PatchCache$') -Label "InstallerPatchCache"
    }

    if ($script:StoppedUpdateServices.Count -gt 0) {
        Invoke-UpdateServiceStart
    }

    $afterFreeGB = Get-SystemDriveFreeSpaceGB
    Write-SmartM365Log "FreeSpaceGBAfter=$afterFreeGB"

    if ($null -eq $afterFreeGB) {
        Write-SmartM365Log "Status=CompletedButFreeSpaceUnknown"
        Write-IntuneResult -Values ([ordered]@{ Status = "CompletedButFreeSpaceUnknown"; FreeSpaceGBBefore = $beforeFreeGB; RequiredFreeSpaceGB = $MinimumTargetFreeSpaceGB; Actions = (Join-CompactList -Items $script:ActionsPerformed); Skipped = (Join-CompactList -Items $script:SkippedActions); LogPath = $LogPath; Script = $ScriptName; Version = $Version })
        exit 1
    }

    if ($afterFreeGB -lt $MinimumTargetFreeSpaceGB) {
        Write-SmartM365Log "Status=CompletedButStillBelowThreshold"
        Write-IntuneResult -Values ([ordered]@{ Status = "CompletedButStillBelowThreshold"; FreeSpaceGBBefore = $beforeFreeGB; FreeSpaceGBAfter = $afterFreeGB; RequiredFreeSpaceGB = $MinimumTargetFreeSpaceGB; Actions = (Join-CompactList -Items $script:ActionsPerformed); Skipped = (Join-CompactList -Items $script:SkippedActions); LogPath = $LogPath; Script = $ScriptName; Version = $Version })
        exit 1
    }

    Write-SmartM365Log "Status=Completed"
    Write-IntuneResult -Values ([ordered]@{ Status = "Completed"; FreeSpaceGBBefore = $beforeFreeGB; FreeSpaceGBAfter = $afterFreeGB; RequiredFreeSpaceGB = $MinimumTargetFreeSpaceGB; Actions = (Join-CompactList -Items $script:ActionsPerformed); Skipped = (Join-CompactList -Items $script:SkippedActions); LogPath = $LogPath; Script = $ScriptName; Version = $Version })
    exit 0
}
catch {
    try {
        if ($script:StoppedUpdateServices.Count -gt 0) {
            Invoke-UpdateServiceStart
        }

        Write-SmartM365Log "Status=Error Message=$($_.Exception.Message)"
    }
    catch {
        $null = $_
    }

    Write-IntuneResult -Values ([ordered]@{ Status = "Error"; Message = $_.Exception.Message; LogPath = $LogPath; Script = $ScriptName; Version = $Version })
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCHsNplxZtJggm6
# +cWjxBHRYRkr5NocWryQm2FqyelkUqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPKvMiQrwSj7f0HwQyron34RI2rRSub8+DwdAwmq5WklMA0GCSqG
# SIb3DQEBAQUABIIBgCTGmA+80j1b3qiVAaEucMys/RZ+zbCVYHkzQcAYhl56t20O
# mkUl4RsWFUzWBoY2jbDAbdX75kY+Oq0MiJzy7JgfqHFSawclyX1xo8IfnuTrtNDV
# t06QBoe6DMeRV6msxo5lMOaO7A312pAa3nn4koEDvWZoeQdBCOmzaYf7D98pS5sY
# 6ZiMtLAreDsbLZs2+EmB7HG5dCpc6zBlbfPnNKlw6WLHdDCpM1Dbm10BHApOuTbP
# W4mHxL+cZ0lX/Qlcgw8fmx4fR2o/CM8rTRPbLuhxjlU04ObfMIIF0Cn0R/BOwz7k
# pNj6mY83wVotmgURPelxGuyyQ0IpxmCnl77JiBjn7B21j5q2QmqwWr5SQvJFY5El
# FLXFCWzs18lQGzaK59+tpED9WmZ/il/vmQckkQmMEAm59AGZXrlgDbE3w01VtLIR
# wHKu91dDHlbM5lZObOxDx5OMMFTG7eVBgahPpbayA11if3o/CSmgEzNat7EJLfxY
# lIMi5q8q5AOB7keoxaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# NDhaMC8GCSqGSIb3DQEJBDEiBCCjl7ZPphaYNKPKTtv4lreLEvYQwBD9SL9kuFGT
# AHk4+TANBgkqhkiG9w0BAQEFAASCAgBxriUmOYe1JsSnkCK2eQHUXaoeULtGsCC7
# ZdubDCX4rd80m13x5uxvkSX+caLHqFKw7MtJ2oNS1l906mBEypG7UurohD+vQUkQ
# VfrIXyDDSDUV38xA6p8uO6UwsdUvxGHcjqdRUPfHVfY4hfaIo/gaVLuPHw8zx729
# vrhpofFpkWXnB7mCJtuHEt953/7T9iQaAKubnneaOlZdt83FsE3K4DrASzUXd4SP
# Uik54CokUIib/I3njxpL9r0+8Ih5sS8l3wau/OnQlecvwJ4EU7X74rTgLGAfAE33
# L93/OHclRfC/trFnG1svUkOVYYftrzyM1pBLheFvDJXCP3LjLLuRnkLQD5yLOk3S
# 2tcDbDYWb2r8e2k2+Bz/hvuzIYFj+SBSfz7JtePjeJoShwmXWBQRNnFhKayecFpm
# rf5WKfdyyZPTWDXcw4/xZWPe1w0DFfziNkQQWTblsLuMBiTuIZpLZ4wpbCF7Lah2
# NjOL6t7BlvLSAnyPpAs1wNVMXFSF/PCMt/+hVe113Uj+b+HJsvQh3uJULWrz7KHa
# 3rJ98sJ8WfCSfGEcP+37ptmQ/s6AdeVQwA49j5rMo8IwcONDP5+GJwWHlCll/FDy
# fvAi1rMOQYwvU8UOBIbp8y+fYfvkSi1qnky7d90WWCvvvh1JrVJgeKKBT3m2wmhO
# 2PUGjaxwuw==
# SIG # End signature block
