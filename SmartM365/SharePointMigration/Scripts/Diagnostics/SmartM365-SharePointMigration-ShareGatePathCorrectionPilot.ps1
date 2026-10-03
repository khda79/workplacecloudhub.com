<#
.SYNOPSIS
    Three-file ShareGate path correction pilot after a reviewed one-file qualification.
.VERSION
    1.0.0
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [Parameter(Mandatory)][string]$AnalysisDirectory,
    [Parameter(Mandatory)][string]$QualificationDirectory,
    [Parameter(Mandatory)][string]$SessionId,
    [Parameter(Mandatory)][string]$SourceItemIdsCsv,
    [ValidateNotNullOrEmpty()][string]$FarmTimeZoneId = 'W. Europe Standard Time',
    [ValidateRange(1,60)][int]$EstimatedMinutesPerItem = 2,
    [ValidateRange(0,240)][int]$MaintenanceMarginMinutes = 60,
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedAnalysisHash = '',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedQualificationHash = '',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedPlanHash = '',
    [switch]$DryRun,
    [switch]$Run,
    [switch]$ConfirmPathCorrection
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$version = '1.0.0'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {
    if ($DryRun -and $Run) { throw 'Choose either -DryRun or -Run.' }
    if ($ConfirmPathCorrection -and -not $Run) { throw '-ConfirmPathCorrection applies only with -Run.' }
    if ($Run -and (-not $ConfirmPathCorrection -or -not $ExpectedAnalysisHash -or
        -not $ExpectedQualificationHash -or -not $ExpectedPlanHash)) {
        throw 'A real run requires -Run, -ConfirmPathCorrection, and all three reviewed SHA256 hashes.'
    }
    if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
        throw 'Use Windows PowerShell 5.1 (powershell.exe) for ShareGate.'
    }
    . (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-FarmMaintenance.ps1')
    . (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-DestinationPath.ps1')
    . (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-ShareGateReportReader.ps1')
    $aliases = Get-SmartM365ShareGateReportAliases -ConfigRoot (Join-Path $PSScriptRoot '..\..\Config')
    $farmZone = [TimeZoneInfo]::FindSystemTimeZoneById($FarmTimeZoneId)
    $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase 'path correction preparation'

    $parts = @($SourceItemIdsCsv.Split(',') | ForEach-Object { $_.Trim() })
    if ($parts.Count -ne 3 -or @($parts | Sort-Object -Unique).Count -ne 3) {
        throw 'SourceItemIdsCsv must contain exactly three distinct positive IDs.'
    }
    $ids = @($parts | ForEach-Object {
        $parsed = 0
        if (-not [int]::TryParse($_, [ref]$parsed) -or $parsed -le 0) {
            throw 'SourceItemIdsCsv must contain exactly three distinct positive IDs.'
        }
        $parsed
    })
    $project = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
    $analysis = (Resolve-Path -LiteralPath $AnalysisDirectory -ErrorAction Stop).ProviderPath
    $qualification = (Resolve-Path -LiteralPath $QualificationDirectory -ErrorAction Stop).ProviderPath
    $diagnosticsRoot = Join-Path $project 'ShareGate\Diagnostics'
    foreach ($directory in @($analysis,$qualification)) {
        if (-not $directory.StartsWith($diagnosticsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'AnalysisDirectory and QualificationDirectory must be under this project ShareGate\Diagnostics folder.'
        }
    }
    if ((Split-Path -Leaf $qualification) -notlike 'PathQualification-*' -or $analysis -eq $qualification) {
        throw 'QualificationDirectory must be a separate PathQualification folder.'
    }
    $classifiedPath = Join-Path $analysis 'ClassifiedRows.csv'
    $qualificationResultPath = Join-Path $qualification 'PathQualification-Result.json.txt'
    $qualificationReportPath = Join-Path $qualification 'ShareGate-Report.csv'
    $qualificationLogPath = Join-Path $qualification 'PathQualification.log'
    foreach ($path in @($classifiedPath,$qualificationResultPath,$qualificationReportPath,$qualificationLogPath)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required evidence is missing: $path" }
    }
    $analysisHash = (Get-FileHash -LiteralPath $classifiedPath -Algorithm SHA256).Hash
    if ($ExpectedAnalysisHash -and $analysisHash -ne $ExpectedAnalysisHash.ToUpperInvariant()) {
        throw 'ClassifiedRows.csv hash differs from the reviewed analysis.'
    }

    function Get-PathCorrectionHash {
        param([string[]]$Paths)
        $componentHashes = @($Paths | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash })
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($componentHashes -join ':'))))).Replace('-', '')
        }
        finally { $sha.Dispose() }
    }
    $qualificationFiles = @($qualificationResultPath,$qualificationReportPath,$qualificationLogPath)
    $qualificationHash = Get-PathCorrectionHash -Paths $qualificationFiles
    if ($ExpectedQualificationHash -and $qualificationHash -ne $ExpectedQualificationHash.ToUpperInvariant()) {
        throw 'Qualification evidence hash differs from the reviewed result, report, or log.'
    }
    $qualified = Get-Content -LiteralPath $qualificationResultPath -Raw -ErrorAction Stop | ConvertFrom-Json
    if ($qualified.Qualification -ne 'Passed' -or $qualified.Error -or $qualified.ShareGateResult -ne 'Success' -or
        $qualified.SessionId -ne $SessionId -or $qualified.AnalysisSHA256 -ne $analysisHash -or
        $qualified.RootBefore -ne 'Absent' -or $qualified.RootAfter -ne 'Absent' -or
        -not $qualified.SourceRead -or -not $qualified.DestinationBefore -or
        $qualified.DestinationBefore -ne $qualified.DestinationAfter -or
        [string]$qualified.SourceItemId -in $parts -or $qualified.CopySessionId -notmatch '^\d{6}-\d+$') {
        throw 'Qualification result does not prove a different item on the same reviewed analysis.'
    }
    $qualifiedDestination = ConvertTo-SmartM365ShareGateRelativePath -Path ([string]$qualified.DestinationFilePath) -Side QualifiedDestination
    $sourceUri = [uri]$qualified.SourceUrl
    $destinationUri = [uri]$qualified.DestinationUrl
    if ($sourceUri.Scheme -ne 'https' -or $destinationUri.Scheme -ne 'https' -or
        $sourceUri.Host -eq $destinationUri.Host -or
        $destinationUri.Host -notmatch '(?i)\.sharepoint\.(com|cn|de)$') {
        throw 'Qualification endpoints must be distinct HTTPS source and SharePoint Online hosts.'
    }
    if ((ConvertTo-SmartM365ShareGateRelativePath -Path ([string]$qualified.ExportedDestinationPath) -Side QualifiedExport) -ne $qualifiedDestination) {
        throw 'Qualification export path differs from its expected destination path.'
    }
    $qualificationRows = @(Import-Csv -LiteralPath $qualificationReportPath -Encoding UTF8)
    if (-not $qualificationRows.Count -or -not (Test-SmartM365ShareGateReportField -Row $qualificationRows[0] -Field 'SourceItemId' -Aliases $aliases)) {
        throw 'Qualification report lacks a source item ID column.'
    }
    $qualificationItemRows = @($qualificationRows | Where-Object {
        (Get-SmartM365ShareGateReportValue -Row $_ -Field 'SourceItemId' -Aliases $aliases) -eq [string]$qualified.SourceItemId
    })
    if ($qualificationItemRows.Count -ne 1 -or
        (Get-SmartM365ShareGateReportValue -Row $qualificationItemRows[0] -Field 'Status' -Aliases $aliases) -ne 'Success' -or
        (Get-SmartM365ShareGateReportValue -Row $qualificationItemRows[0] -Field 'DestinationPath' -Aliases $aliases) -ne $qualifiedDestination -or
        (Get-SmartM365ShareGateReportValue -Row $qualificationItemRows[0] -Field 'Errors' -Aliases $aliases) -or
        (Get-SmartM365ShareGateReportValue -Row $qualificationItemRows[0] -Field 'Warnings' -Aliases $aliases) -or
        [string]$qualificationItemRows[0].'Session ID' -ne [string]$qualified.CopySessionId) {
        throw 'Qualification report does not confirm one successful item at the reviewed destination path.'
    }
    $qualificationLog = Get-Content -LiteralPath $qualificationLogPath -Raw -ErrorAction Stop
    $versionMatch = [regex]::Match($qualificationLog, 'ShareGate module version=(?<version>\d+(?:\.\d+){2,3});')
    if (-not $versionMatch.Success -or $qualificationLog -notmatch 'Qualification=Passed; ShareGate=Success;') {
        throw 'Qualification log lacks a passed result or ShareGate version.'
    }
    $shareGateVersion = $versionMatch.Groups['version'].Value

    $classified = @(Import-Csv -LiteralPath $classifiedPath -Encoding UTF8 | Where-Object {
        $_.SessionId -eq $SessionId -and $_.RuleId -eq 'SG-ACCESS-SOURCE' -and
        $_.State -eq 'To fix' -and $_.AccessSide -eq 'Source' -and $_.ObjectType -eq 'File'
    })
    $plan = @(
        foreach ($id in $ids) {
            $matches = @($classified | Where-Object { $_.SourceItemId -eq [string]$id })
            if (@($matches | ForEach-Object ItemKey | Sort-Object -Unique).Count -ne 1) {
                throw "Source ID $id must identify exactly one eligible source 401 file in this session."
            }
            $first = $matches[0]
            if (@($matches | Where-Object {
                $_.SourceUrl -ne $first.SourceUrl -or $_.SourceList -ne $first.SourceList -or
                $_.DestinationUrl -ne $first.DestinationUrl -or $_.DestinationList -ne $first.DestinationList -or
                $_.'Raw: Source path' -ne $first.'Raw: Source path' -or
                $_.'Raw: Destination path' -ne $first.'Raw: Destination path'
            }).Count) { throw "Rows for source ID $id disagree on item routing." }
            $route = Resolve-SmartM365ShareGateDestinationPath -Row $first
            if (-not $route.DestinationFolder -or $first.SourceUrl -ne $qualified.SourceUrl -or
                $first.SourceList -ne $qualified.SourceList -or $first.DestinationUrl -ne $qualified.DestinationUrl -or
                $first.DestinationList -ne $qualified.DestinationList) {
                throw "Source ID $id is outside the qualified site, list, or subfolder scope."
            }
            [pscustomobject]@{
                ItemKey=$first.ItemKey; SourceItemId=$id; SourceUrl=$first.SourceUrl; SourceList=$first.SourceList;
                DestinationUrl=$first.DestinationUrl; DestinationList=$first.DestinationList;
                SourceFilePath=$route.SourceFilePath; DestinationFilePath=$route.DestinationFilePath;
                DestinationFolder=$route.DestinationFolder; FileName=$route.FileName
            }
        }
    )
    if ($plan.Count -ne 3 -or @($plan | ForEach-Object ItemKey | Sort-Object -Unique).Count -ne 3) {
        throw 'Path correction plan must contain exactly three distinct items.'
    }
    $canonical = @($analysisHash,$qualificationHash) + @($plan | ForEach-Object {
        @($_.ItemKey,$_.SourceItemId,$_.SourceUrl,$_.SourceList,$_.SourceFilePath,
          $_.DestinationUrl,$_.DestinationList,$_.DestinationFilePath,$_.DestinationFolder) -join "`t"
    })
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $planHash = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($canonical -join "`n"))))).Replace('-', '') }
    finally { $sha.Dispose() }
    if ($ExpectedPlanHash -and $planHash -ne $ExpectedPlanHash.ToUpperInvariant()) { throw 'Path correction plan hash differs from the reviewed plan.' }

    function Assert-PathCorrectionSchedule {
        param([int]$Remaining, [string]$Phase)
        $state = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase $Phase
        $window = $state.FarmTime.Date.AddHours(23).AddMinutes(45)
        if ($state.FarmTime -ge $window) { $window = $window.AddDays(1) }
        $estimatedEnd = $state.FarmTime.AddMinutes(($Remaining * $EstimatedMinutesPerItem) + $MaintenanceMarginMinutes)
        if ($estimatedEnd -ge $window) {
            throw ('Refusing {0}: estimated finish plus margin reaches farm maintenance at {1:yyyy-MM-dd HH:mm:ss}.' -f $Phase,$window)
        }
        return $estimatedEnd
    }
    $estimatedEnd = Assert-PathCorrectionSchedule -Remaining 3 -Phase 'three-item path correction preparation'
    Write-Output ('{0} Mode={1}; Session={2}; itemCount=3; writes={3}; AnalysisSHA256={4}; QualificationSHA256={5}; PlanSHA256={6}' -f
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$(if ($Run) { 'Run' } else { 'DryRun' }),$SessionId,
        $(if ($Run) { 'at most three SPO destination item copies after confirmation' } else { 'none' }),
        $analysisHash,$qualificationHash,$planHash)
    Write-Output ('{0} Qualified item={1}; ShareGate={2}; estimatedMinutes={3}; marginMinutes={4}; estimatedEndFarmTime={5:yyyy-MM-dd HH:mm:ss}' -f
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$qualified.SourceItemId,$shareGateVersion,
        (3 * $EstimatedMinutesPerItem),$MaintenanceMarginMinutes,$estimatedEnd)
    foreach ($item in $plan) {
        Write-Output ('{0} ID={1}; source={2} | {3} | {4}; destination={5} | {6} | {7}; folder={8}' -f
            (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$item.SourceItemId,$item.SourceUrl,$item.SourceList,$item.SourceFilePath,
            $item.DestinationUrl,$item.DestinationList,$item.DestinationFilePath,$item.DestinationFolder)
    }
    $phrase = 'CORRECT 3 ITEMS ' + $SessionId + ' v' + $version
    Write-Output ('{0} Required confirmation for a future Run: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$phrase)
    if (-not $Run) { return }

    $entered = Read-Host ('Type exactly "{0}" to authorize three SPO destination item copies' -f $phrase)
    if ($entered -cne $phrase) { throw 'Path correction confirmation was not entered exactly. No copy was started.' }
    $null = Assert-PathCorrectionSchedule -Remaining 3 -Phase 'three-item path correction startup'
    if ((Get-PathCorrectionHash -Paths $qualificationFiles) -ne $qualificationHash -or
        (Get-FileHash -LiteralPath $classifiedPath -Algorithm SHA256).Hash -ne $analysisHash) {
        throw 'Reviewed evidence changed after confirmation. No copy was started.'
    }

    $module = @(Get-Module -ListAvailable -Name ShareGate | Sort-Object Version -Descending | Select-Object -First 1)
    if ($module.Count -ne 1) { throw 'ShareGate module is not installed in Windows PowerShell 5.1.' }
    Import-Module -Name $module[0].Path -ErrorAction Stop
    if ([string]$module[0].Version -ne $shareGateVersion) { throw 'Installed ShareGate version differs from the qualified version.' }
    foreach ($name in @('Connect-Site','Get-List','Get-Folder','Get-File','Copy-Content','New-CopySettings','Export-Report')) {
        if (-not (Get-Command -Name $name -Module ShareGate -ErrorAction SilentlyContinue)) { throw "Required ShareGate cmdlet is missing: $name" }
    }
    $copyCommand = Get-Command -Name Copy-Content -Module ShareGate
    $requiredParameters = @('SourceList','DestinationList','SourceItemId','DestinationFolder','CopySettings','TaskName')
    if (-not @($copyCommand.ParameterSets | Where-Object {
        $names = @($_.Parameters | ForEach-Object Name)
        @($requiredParameters | Where-Object { $_ -notin $names }).Count -eq 0
    }).Count) { throw 'Installed Copy-Content lacks the required item-scoped DestinationFolder parameter set.' }
    $settings = ShareGate\New-CopySettings -OnContentItemExists Overwrite -ErrorAction Stop
    if (-not $settings) { throw 'New-CopySettings returned no settings.' }

    $lockPath = Join-Path $diagnosticsRoot ('PathCorrection-' + $SessionId + '.lock')
    try { $runLock = [IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None) }
    catch { throw "Another path correction run may be active for session $SessionId. Exclusive lock: $lockPath" }
    try {
        $runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'),[guid]::NewGuid().ToString('N')
        $output = Join-Path $diagnosticsRoot ('PathCorrection-' + $runId)
        $reportDir = Join-Path $output 'Reports'
        New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
        $logPath = Join-Path $output 'PathCorrection.log'
        $resultsPath = Join-Path $output 'PathCorrection-Results.csv'
        $summaryPath = Join-Path $output 'PathCorrection-Summary.json.txt'
        $results = [System.Collections.Generic.List[object]]::new()
        $runStatus = 'Running'
        foreach ($item in $plan) {
            $results.Add([pscustomobject]@{ SourceItemId=$item.SourceItemId; ItemKey=$item.ItemKey;
                SourceFilePath=$item.SourceFilePath; DestinationFilePath=$item.DestinationFilePath;
                DestinationFolder=$item.DestinationFolder; Status='NotAttempted'; CopySessionId='';
                DestinationItemUrl=''; RootBefore=''; RootAfter=''; ReportPath=''; Error='' })
        }
        function Write-CorrectionLog {
            param([string]$Message)
            $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Message
            Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
            Write-Output $line
        }
        function Save-CorrectionState {
            $temporary = Join-Path $output ('.' + [guid]::NewGuid().ToString('N') + '.tmp')
            try {
                $results.ToArray() | Export-Csv -LiteralPath $temporary -NoTypeInformation -Encoding UTF8
                Move-Item -LiteralPath $temporary -Destination $resultsPath -Force
            }
            finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
            $summary = [ordered]@{ ScriptVersion=$version; SessionId=$SessionId; RunStatus=$runStatus;
                Actor=($env:USERDOMAIN+'\'+$env:USERNAME);
                Machine=$env:COMPUTERNAME; AnalysisSHA256=$analysisHash; QualificationSHA256=$qualificationHash;
                PlanSHA256=$planHash; Planned=3; Success=@($results|Where-Object Status -EQ 'Success').Count;
                Skipped=@($results|Where-Object Status -EQ 'Skipped').Count;
                Error=@($results|Where-Object Status -EQ 'Error').Count;
                Unreported=@($results|Where-Object Status -EQ 'Unreported').Count;
                NotAttempted=@($results|Where-Object Status -EQ 'NotAttempted').Count;
                ResultsPath=$resultsPath; LogPath=$logPath }
            $temporary = Join-Path $output ('.' + [guid]::NewGuid().ToString('N') + '.tmp')
            try {
                ($summary | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $temporary -Encoding UTF8
                Move-Item -LiteralPath $temporary -Destination $summaryPath -Force
            }
            finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
        }
        function Get-ExactCorrectionList {
            param($Site,[string]$Name)
            $found = @(ShareGate\Get-List -Site $Site -Name $Name -ErrorAction Stop)
            $exact = @($found | Where-Object {
                ($_.PSObject.Properties['Title'] -and $_.Title -eq $Name) -or
                ($_.PSObject.Properties['Name'] -and $_.Name -eq $Name)
            })
            if ($exact.Count -ne 1) { throw "Get-List did not resolve exactly one list named '$Name'." }
            return $exact[0]
        }
        function Get-OneCorrectionFile {
            param($List,[string]$Path,[switch]$AllowMissing)
            try { $files = @(ShareGate\Get-File -List $List -Path $Path -ErrorAction Stop) }
            catch {
                if ($AllowMissing -and $_.Exception.Message -match '(?i)(notfound|not found|could not be found|cannot find|does not exist|doesn.t exist|\b404\b|introuvable|n.existe pas|n.a pas .t. trouv.)') { return $null }
                throw
            }
            if ($files.Count -gt 1 -or ($files.Count -eq 0 -and -not $AllowMissing)) { throw "Get-File did not resolve exactly one file at $Path." }
            if ($files.Count -eq 0) { return $null }
            return $files[0]
        }
        Save-CorrectionState
        Write-CorrectionLog ('Actor={0}\{1}; Machine={2}; ShareGate={3}; SourceAuth=CurrentWindowsIdentity; DestinationAuth=Browser; CopySettings=Overwrite; PlanSHA256={4}' -f
            $env:USERDOMAIN,$env:USERNAME,$env:COMPUTERNAME,$shareGateVersion,$planHash)
        $sourceSite = ShareGate\Connect-Site -Url $plan[0].SourceUrl -ErrorAction Stop
        $destinationSite = ShareGate\Connect-Site -Url $plan[0].DestinationUrl -Browser -ErrorAction Stop
        $sourceList = Get-ExactCorrectionList -Site $sourceSite -Name $plan[0].SourceList
        $destinationList = Get-ExactCorrectionList -Site $destinationSite -Name $plan[0].DestinationList
        for ($index = 0; $index -lt 3; $index++) {
            $item = $plan[$index]
            $result = $results[$index]
            $null = Assert-PathCorrectionSchedule -Remaining (3 - $index) -Phase ("path correction item $($index + 1)/3")
            if ((Get-PathCorrectionHash -Paths $qualificationFiles) -ne $qualificationHash -or
                (Get-FileHash -LiteralPath $classifiedPath -Algorithm SHA256).Hash -ne $analysisHash) {
                throw 'Reviewed evidence changed before an item copy.'
            }
            $reportPath = Join-Path $reportDir ('Item-{0:D2}.csv' -f ($index + 1))
            $temporaryReport = Join-Path $reportDir ('.Item-{0:D2}-{1}.csv' -f ($index + 1),[guid]::NewGuid().ToString('N'))
            $taskName = 'SmartM365 path correction ' + $SessionId + ' ' + $runId + ' ID=' + $item.SourceItemId
            $copyStarted = $false
            try {
                $null = Get-OneCorrectionFile -List $sourceList -Path $item.SourceFilePath
                $null = Get-OneCorrectionFile -List $destinationList -Path $item.DestinationFilePath
                Assert-SmartM365ShareGateDestinationFolder -DestinationList $destinationList -DestinationFolder $item.DestinationFolder
                $rootBefore = Get-OneCorrectionFile -List $destinationList -Path $item.FileName -AllowMissing
                $result.RootBefore = if ($rootBefore) { [string]$rootBefore.Address } else { 'Absent' }
                $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase ("item $($index + 1)/3 copy")
                Write-CorrectionLog ('Starting item {0}/3; ID={1}; folder={2}; task={3}' -f
                    ($index + 1),$item.SourceItemId,$item.DestinationFolder,$taskName)
                $copyStarted = $true
                $copyResult = ShareGate\Copy-Content -SourceList $sourceList -DestinationList $destinationList -SourceItemId @($item.SourceItemId) -DestinationFolder $item.DestinationFolder -CopySettings $settings -TaskName $taskName -ErrorAction Stop
                if (-not $copyResult -or @($copyResult).Count -ne 1) { throw 'Copy-Content did not return exactly one CopyResult.' }
                foreach ($name in @('SessionId','SessionID','CopySessionId','Id')) {
                    $property = $copyResult.PSObject.Properties[$name]
                    if ($property -and [string]$property.Value -match '^\d{6}-\d+$') { $result.CopySessionId = [string]$property.Value; break }
                }
                ShareGate\Export-Report -CopyResult $copyResult -Path $temporaryReport -ErrorAction Stop | Out-Null
                if (-not (Test-Path -LiteralPath $temporaryReport -PathType Leaf)) { throw 'Export-Report did not create a CSV.' }
                Move-Item -LiteralPath $temporaryReport -Destination $reportPath -Force
                $result.ReportPath = $reportPath
                $reportRows = @(Import-Csv -LiteralPath $reportPath -Encoding UTF8)
                $itemRows = @($reportRows | Where-Object {
                    (Get-SmartM365ShareGateReportValue -Row $_ -Field 'SourceItemId' -Aliases $aliases) -eq [string]$item.SourceItemId
                })
                if ($itemRows.Count -ne 1) { throw 'ShareGate report did not contain exactly one row for the selected item.' }
                $status = Get-SmartM365ShareGateReportValue -Row $itemRows[0] -Field 'Status' -Aliases $aliases
                if ($status -notin @('Success','Skipped')) { throw "ShareGate reported $status for the selected item." }
                if ((Get-SmartM365ShareGateReportValue -Row $itemRows[0] -Field 'Errors' -Aliases $aliases) -or
                    (Get-SmartM365ShareGateReportValue -Row $itemRows[0] -Field 'Warnings' -Aliases $aliases)) {
                    throw 'ShareGate reported an item error or warning.'
                }
                $importStates = @($reportRows | ForEach-Object { [string]$_.'Microsoft 365 Import: Status' } | Where-Object { $_ })
                if (-not @($importStates | Where-Object { $_ -eq 'Finished' }).Count -or
                    @($importStates | Where-Object { $_ -ne 'Finished' }).Count) {
                    throw 'ShareGate export does not confirm a finished Microsoft 365 import.'
                }
                $exportedPath = Get-SmartM365ShareGateReportValue -Row $itemRows[0] -Field 'DestinationPath' -Aliases $aliases
                if (-not $exportedPath -or
                    (ConvertTo-SmartM365ShareGateRelativePath -Path $exportedPath -Side ExportedDestination) -ne $item.DestinationFilePath) {
                    throw 'ShareGate export does not prove the expected destination subfolder path.'
                }
                $after = Get-OneCorrectionFile -List $destinationList -Path $item.DestinationFilePath
                $result.DestinationItemUrl = [string]$after.Address
                $rootAfter = Get-OneCorrectionFile -List $destinationList -Path $item.FileName -AllowMissing
                $result.RootAfter = if ($rootAfter) { [string]$rootAfter.Address } else { 'Absent' }
                if (-not $rootBefore -and $rootAfter) { throw 'A new file appeared at the destination library root.' }
                $result.Status = $status
                if ($status -eq 'Skipped') { throw 'ShareGate skipped the item; correction must be reviewed before any further copy.' }
                Write-CorrectionLog ('Completed ID={0}; Status={1}; CopySession={2}; Destination={3}; RootBefore={4}; RootAfter={5}; Report={6}' -f
                    $item.SourceItemId,$status,$result.CopySessionId,$result.DestinationItemUrl,$result.RootBefore,$result.RootAfter,$reportPath)
            }
            catch {
                if ($result.Status -eq 'NotAttempted') { $result.Status = if ($copyStarted) { 'Unreported' } else { 'Error' } }
                $result.Error = $_.Exception.Message
                $runStatus = 'Stopped'
                Save-CorrectionState
                Write-CorrectionLog ('Stopped after ID={0}: {1}; Report={2}' -f $item.SourceItemId,$result.Error,$result.ReportPath)
                throw
            }
            finally { if (Test-Path -LiteralPath $temporaryReport) { Remove-Item -LiteralPath $temporaryReport -Force } }
            Save-CorrectionState
        }
        $runStatus = 'Completed'
        Save-CorrectionState
        Write-CorrectionLog ('Completed all three items; Results={0}; Summary={1}' -f $resultsPath,$summaryPath)
    }
    catch {
        if ($runStatus -ne 'Stopped') {
            $runStatus = 'Stopped'
            Save-CorrectionState
            Write-CorrectionLog ('Stopped before or between item copies: {0}' -f $_.Exception.Message)
        }
        throw
    }
    finally {
        if ($runLock) { $runLock.Dispose() }
    }
}
catch {
    $script:ConsoleLifecycleFailure = $_
    throw
}
finally {
    Complete-SmartM365MigrationConsoleLifecycle -Context $script:ConsoleLifecycleContext -Failure $script:ConsoleLifecycleFailure -Status $script:ConsoleLifecycleStatus
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAGXvU+9ytBMbeO
# 2aZKpM2fmM+/CtbGVuORC9y+dzDw86CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB14QJw1a3vJiq2Jxtf8sG+4CnpGC4QWXJFkAsnSFSeZMA0GCSqG
# SIb3DQEBAQUABIIBgCJhPKq6RKikvP7qZLl4Y/ZXG8Ky8MYCa3/CL6E48YSiuDN+
# l3jb2I1Kd17cLNMZpQouU6wxTvvhl9qp2yxibQSg9n+uipywL4xxjaQ098qNA0tt
# 8ttFB9/fiXT1YxnUdzozkh1Nw9HXKw2oHzlG3ATmLwdY3flwlOWmhalQ9mWs8LDG
# oOSvQ9ynjIsn6ew1aVDMMEYngLhlz4NL7DwkG6NFn/C6uusYCNddTD3QMKMcXc4+
# uIydNSlxHW7QV8yUSZ76UBFb2yrMVEadmQSmSdZGch+7HbYoK/yv92vJyh6NO/VP
# PvupqBXtqkTG7OTG3u9o70ROLDbe/AgJqBvyJ/JS63VfPwbUyXGs3olLLykVk8ZS
# gZCMuUjITY+mdErFHVLV6BGYNqNAYm4rzhvl6k0TvxCauFBohVpCXkYe/M82iy09
# kY747XrcHl5cl7AjW2iQ8124sPg8tOtRipOiH91WAXauoMoyctHGcZDir9cRfHm4
# nftQfmrx0RWKThuMvKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNzQw
# MDlaMC8GCSqGSIb3DQEJBDEiBCAULT+AIlM5Th06hsjfn8d/wJrn1UsLlFc3IPNY
# aEVwnzANBgkqhkiG9w0BAQEFAASCAgAWHMmrjca7X3ONH4GZFbc5aND+BiFCMABq
# 9rl35lAYq0VpRLIRlkYDcW060ltfBQaNISZ1y9Yvhra0MnabCc4huMGjhBgdlKNS
# ecPmXPINDQKj6AawMDLES3+neY32YOZnlGUuksuWb8IwjPHH0Z2zoHD4mSTqBeJO
# K5MSpnzhdDgiPQSrVMFROeXMw3wVl8NSykNgLzKb2KOPcxz7OzvfZCG/CTY4VWJ7
# /Fu9nvdzm3+/Ba7dNcVBtJLddLju4z+sOMRHUZs1GyI0rjeU7ZJp1om5T5jM1iMT
# 4+7xk6/hOhpuAvvgJn7Rd8SxgYeuLW9YmcckzkNAbcYc7Xg7iQs8yRp+Bi4jJs1Y
# ZwEGUJAZGnRkCHrakZqwnLowHSgWDmlhprk27AE84Wziy9LW0r3Zj0mOGMR2DBk+
# zHg9GFDme6c1L0FMWq/eOR3uHZw1xE60/xcaWkrRBVG+reGYqTesOW0+e6l2bDbl
# ckaqE1IdsKKHoBFx9Xza+BOxBEm84F1iyW+RWkLnHIIlnYbSF562SA0Xg2o+cgGF
# fF0+rXnBrN5IcpAozGvAnLZmjRI2scBDow1z1StiqE1jz2h2/Tk/9vRQ8jkz4lY4
# AvoT2ziP046HFs+48VXM1mUxweFmEH+T9pjN4q7qp8Ll/W5JovMVbdQnsEyz3ivh
# 1xZyx0SbLg==
# SIG # End signature block
