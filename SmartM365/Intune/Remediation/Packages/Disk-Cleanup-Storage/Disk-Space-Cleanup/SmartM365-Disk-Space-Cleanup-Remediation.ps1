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
