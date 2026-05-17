# Name: SmartM365-DO-Issues-Remediation.ps1
# Version: 1.0
# Remediation: Repair WU Content Engine

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$RenameTargets = @(
    @{ Path = 'C:\Windows\SoftwareDistribution'; NewNamePrefix = 'SoftwareDistribution.old' },
    @{ Path = 'C:\Windows\System32\catroot2'; NewNamePrefix = 'catroot2.old' }
)

$RemoveTargets = @(
    'C:\ProgramData\Microsoft\Windows\DeliveryOptimization'
)

function Write-RemediationLog {
    param([string]$Message)
    Write-Output ("DOIssuesRemediation {0}" -f $Message)
}

function Test-AllowedPath {
    param([string]$Path)

    $allowedPaths = @(
        'C:\Windows\SoftwareDistribution',
        'C:\Windows\System32\catroot2',
        'C:\ProgramData\Microsoft\Windows\DeliveryOptimization'
    )

    $normalizedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    return $allowedPaths | Where-Object {
        [System.IO.Path]::GetFullPath($_).TrimEnd('\') -ieq $normalizedPath
    }
}

function Invoke-ServiceStopSafe {
    param([string]$Name)

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-RemediationLog "Service=$Name Status=NotFound"
        return
    }

    if ($DryRun) {
        Write-RemediationLog "Service=$Name Action=Stop Status=DryRun"
        return
    }

    Write-RemediationLog "Service=$Name Action=Stop"
    Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
}

function Invoke-ServiceStartSafe {
    param([string]$Name)

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-RemediationLog "Service=$Name Status=NotFound"
        return
    }

    if ($DryRun) {
        Write-RemediationLog "Service=$Name Action=Start Status=DryRun"
        return
    }

    Write-RemediationLog "Service=$Name Action=Start"
    Start-Service -Name $Name -ErrorAction SilentlyContinue
}

function Rename-CacheFolderSafe {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$Path,
        [string]$NewNamePrefix
    )

    if (-not (Test-AllowedPath -Path $Path)) {
        throw "Refusing rename outside approved target list: $Path"
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-RemediationLog "Path=$Path Status=NotFound"
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
    $newName = "{0}.{1}" -f $NewNamePrefix, $timestamp

    if ($DryRun) {
        Write-RemediationLog "Path=$Path NewName=$newName Status=DryRun"
        return
    }

    if ($PSCmdlet.ShouldProcess($Path, "Rename to $newName")) {
        Write-RemediationLog "Path=$Path NewName=$newName Status=Renaming"
        Rename-Item -LiteralPath $Path -NewName $newName -ErrorAction SilentlyContinue
        Write-RemediationLog "Path=$Path NewName=$newName Status=Renamed"
    }
}

function Remove-FolderSafe {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$Path)

    if (-not (Test-AllowedPath -Path $Path)) {
        throw "Refusing removal outside approved target list: $Path"
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-RemediationLog "Path=$Path Status=NotFound"
        return
    }

    if ($DryRun) {
        Write-RemediationLog "Path=$Path Status=DryRun"
        return
    }

    if ($PSCmdlet.ShouldProcess($Path, 'Remove folder')) {
        Write-RemediationLog "Path=$Path Status=Removing"
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
        Write-RemediationLog "Path=$Path Status=Removed"
    }
}

function Invoke-UsoClientSafe {
    $usoClientPath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\UsoClient.exe'

    if (-not (Test-Path -LiteralPath $usoClientPath -PathType Leaf)) {
        Write-RemediationLog "UsoClient=StartScan Status=NotFound"
        return
    }

    if ($DryRun) {
        Write-RemediationLog "UsoClient=StartScan Status=DryRun"
        return
    }

    Start-Process -FilePath $usoClientPath -ArgumentList 'StartScan' -WindowStyle Hidden -ErrorAction SilentlyContinue
    Write-RemediationLog "UsoClient=StartScan Status=Triggered"
}

try {
    Write-RemediationLog "Status=Started DryRun=$($DryRun.IsPresent)"

    foreach ($serviceName in @('usosvc', 'wuauserv', 'bits', 'dosvc', 'cryptsvc')) {
        Invoke-ServiceStopSafe -Name $serviceName
    }

    foreach ($target in $RenameTargets) {
        Rename-CacheFolderSafe -Path $target.Path -NewNamePrefix $target.NewNamePrefix
    }

    foreach ($targetPath in $RemoveTargets) {
        Remove-FolderSafe -Path $targetPath
    }

    foreach ($serviceName in @('cryptsvc', 'bits', 'dosvc', 'wuauserv', 'usosvc')) {
        Invoke-ServiceStartSafe -Name $serviceName
    }

    Start-Sleep -Seconds 10
    Invoke-UsoClientSafe

    Write-RemediationLog 'Status=Completed'
    exit 0
}
catch {
    Write-RemediationLog "Status=Error Message=$($_.Exception.Message)"
    exit 1
}
