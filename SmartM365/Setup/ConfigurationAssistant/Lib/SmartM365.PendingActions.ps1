# Reviewed create-only setup actions, version 0.8.0. Preparation is local and does not create resources.
function New-SiPendingCreation {
    param([Parameter(Mandatory)][ValidateSet('Global','Profile','LocalAccount','DomainAccount')][string]$Kind,[Parameter(Mandatory)]$Data,[Security.SecureString]$Password)
    switch ($Kind) {
        'Global' {
            $fresh=Get-SiGlobalCreationPlan -Root $Data.Root
            if ($fresh.TemplateHash -cne $Data.TemplateHash -or $fresh.Destination -cne $Data.Destination) {throw 'Global template or destination changed.'}
            $label='Create global configuration: '+$fresh.Destination;$owned=$null
        }
        'Profile' {
            $fresh=Get-SiTenantCreationPlan -Root $Data.Root -ProfileKey $Data.ProfileKey -OrganizationKey $Data.OrganizationKey -EnvironmentKey $Data.EnvironmentKey
            if ($fresh.TemplateHash -cne $Data.TemplateHash -or $fresh.Destination -cne $Data.Destination) {throw 'Tenant template or destination changed.'}
            $label='Create profile '+$fresh.ProfileKey+'; organization: '+$fresh.OrganizationKey+'; environment: '+$fresh.EnvironmentKey+'; data key: '+$fresh.TenantKey+'; SharePoint upload enabled.';$owned=$null
        }
        'LocalAccount' {
            if ($Data.Name -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_.-]{0,19}$' -or $Data.Name.EndsWith('.') -or $null -eq $Password -or -not $Password.Length) {throw 'Enter a valid local account and password.'}
            $label='Create local service account: '+$Data.Name;$owned=$Password.Copy()
        }
        'DomainAccount' {
            if ($Data.Name -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,14}$' -or [Uri]::CheckHostName([string]$Data.DomainController) -ne [UriHostNameType]::Dns -or -not ([string]$Data.DomainController).Contains('.') -or ([string]$Data.OrganizationalUnit).Length -lt 4 -or ([string]$Data.OrganizationalUnit).Length -gt 1024 -or ([string]$Data.OrganizationalUnit) -match '[\x00-\x1f]' -or $null -eq $Password -or -not $Password.Length) {throw 'Enter a valid domain account, controller, organizational unit and password.'}
            $label='Create domain service account '+$Data.Name+' on '+$Data.DomainController+' in '+$Data.OrganizationalUnit;$owned=$Password.Copy()
        }
    }
    if ($Kind -in @('Global','Profile') -and $null -eq $Data.PSObject.Properties['PendingEdits']) {$Data | Add-Member -NotePropertyName PendingEdits -NotePropertyValue ([Collections.Generic.List[object]]::new())}
    [pscustomobject]@{Kind=$Kind;Data=$Data;Secret=$owned;Label=$label;State='Pending'}
}
function Get-SiPendingCreationSnapshot {
    param([Parameter(Mandatory)]$Action,[Parameter(Mandatory)][string]$Tenant,[Parameter(Mandatory)][string]$Step,[Parameter(Mandatory)][string]$Field)
    if ($Action.Kind -notin @('Global','Profile') -or (($Action.Kind -eq 'Global') -ne ($Step -eq 'share')) -or $Field -notin @(Get-SiEditableFields $Step)) {throw 'The setting cannot be edited for this pending creation.'}
    $null=New-SiPendingCreation -Kind $Action.Kind -Data $Action.Data
    $model=if ($Action.Kind -eq 'Global') {$Action.Data.Review | ConvertFrom-Json} else {$Action.Data.Model}
    $property=$model.PSObject.Properties[$Field];if ($null -eq $property) {throw 'The template does not define this setting.'}
    $value=if ($property.Value -is [pscustomobject] -or $property.Value -is [array]) {ConvertTo-Json -InputObject $property.Value -Depth 20 -Compress} elseif ($null -eq $property.Value) {'null'} else {[string]$property.Value}
    [pscustomobject]@{Root=$Action.Data.Root;Tenant=$Tenant;Step=$Step;Field=$Field;Path=$Action.Data.Destination;Hash='ABSENT';Present=$true;Value=$value;Virtual=$true;Action=$Action}
}
function Set-SiPendingCreationSetting {
    param([Parameter(Mandatory)]$Snapshot,[AllowEmptyString()][string]$Value)
    $fresh=Get-SiPendingCreationSnapshot -Action $Snapshot.Action -Tenant $Snapshot.Tenant -Step $Snapshot.Step -Field $Snapshot.Field
    if ($fresh.Value -cne $Snapshot.Value -or $fresh.Path -cne $Snapshot.Path) {throw 'The pending template changed; reload the setting.'}
    $null=ConvertTo-SiEditValue -Snapshot $fresh -Value $Value
    $edits=$Snapshot.Action.Data.PendingEdits
    for ($i=$edits.Count-1;$i -ge 0;$i--) {if ($edits[$i].Step -eq $Snapshot.Step -and $edits[$i].Field -eq $Snapshot.Field) {$edits.RemoveAt($i)}}
    if ($Value -cne $Snapshot.Value) {$edits.Add([pscustomobject]@{Step=$Snapshot.Step;Field=$Snapshot.Field;Value=$Value})}
}
function Set-SiPendingServiceAccount {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$Tenant,[Parameter(Mandatory)][string]$Account,[Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Actions,[Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Edits)
    $pending=@($Actions | Where-Object {$_.Kind -eq 'Profile' -and $_.Data.ProfileKey -ceq $Tenant})
    if ($pending.Count -gt 1) {throw 'Multiple pending profiles match the selected account.'}
    if ($pending.Count -eq 1) {
        $snapshot=Get-SiPendingCreationSnapshot -Action $pending[0] -Tenant $Tenant -Step 'tenant' -Field 'ServiceAccount'
        Set-SiPendingCreationSetting -Snapshot $snapshot -Value $Account
        return
    }
    $snapshot=Get-SiEditSnapshot -Root $Root -Tenant $Tenant -Step 'tenant' -Field 'ServiceAccount'
    $null=ConvertTo-SiEditValue -Snapshot $snapshot -Value $Account
    if ($snapshot.Present -and $snapshot.Value -ieq $Account) {
        for ($i=$Edits.Count-1;$i -ge 0;$i--) {if ($Edits[$i].Snapshot.Path -ieq $snapshot.Path -and $Edits[$i].Snapshot.Field -eq 'ServiceAccount') {$Edits.RemoveAt($i)}}
        return
    }
    Add-SiPendingEdit -Queue $Edits -Edit (New-SiPendingEdit -Snapshot $snapshot -Value $Account)
}
function Add-SiPendingCreation {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue,[Parameter(Mandatory)]$Action)
    if (@($Queue | Where-Object {($_.Kind -eq $Action.Kind -and $Action.Kind -ne 'FileEdit') -or ($_.Kind -in @('LocalAccount','DomainAccount') -and $Action.Kind -in @('LocalAccount','DomainAccount'))}).Count) {if ($null -ne $Action.Secret) {$Action.Secret.Dispose();$Action.Secret=$null};throw 'Discard the existing pending action of this type before replacing it.'}
    if ($Action.Kind -eq 'FileEdit') {foreach ($old in @($Queue | Where-Object {$_.Kind -eq 'FileEdit' -and $_.Data.Path -ieq $Action.Data.Path})) {
        if ($old.Data.Type -ne $Action.Data.Type -or $old.Data.Hash -cne $Action.Data.Hash -or $old.Data.Type -eq 'Reconciliation') {throw 'The pending changes for this file conflict or were reviewed against different versions.'}
        $oldField=$old.Data.Plan.Field
        $newField=$Action.Data.Plan.Field
        if ($oldField -ieq $newField) {[void]$Queue.Remove($old)}
    }}
    $Queue.Add($Action)
}
function New-SiPendingFileAction {
    param([Parameter(Mandatory)][ValidateSet('Orchestrator','Reconciliation')][string]$Type,[Parameter(Mandatory)]$Plan,[string]$Value='',[string]$Worker='',[string]$Root='',[string]$ScriptPath='')
    $snapshot=if ($Type -eq 'Reconciliation') {$null} else {$Plan}
    $path=if ($Type -eq 'Reconciliation') {$Plan.TargetPath} else {$snapshot.Path}
    $hash=if ($Type -eq 'Reconciliation') {$Plan.TargetHash} else {$snapshot.Hash}
    if (-not [IO.Path]::IsPathFullyQualified($path)) {throw 'A fully qualified configuration path is required.'}
    if ($Worker) {throw 'External script configuration workers are not supported.'}
    if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $hash) {throw 'Configuration changed; review it again.'}
    $description=switch ($Type) {
        'Orchestrator' {'Orchestrator '+$snapshot.Field+': '+$snapshot.Value+' -> '+$Value}
        'Reconciliation' {'Bootstrap additions: '+(($Plan.Additions | ForEach-Object {$_.Field+' = '+$_.Value}) -join ', ')}
    }
    $data=[pscustomobject]@{Type=$Type;Plan=$Plan;Value=$Value;Worker=$Worker;Root=$Root;ScriptPath=$ScriptPath;Path=$path;Hash=$hash}
    [pscustomobject]@{Kind='FileEdit';Data=$data;Secret=$null;Label=($description+'; file: '+$path);State='Pending'}
}
function New-SiPendingTestAction {
    param([Parameter(Mandatory)][ValidateSet('MailTest','ShareWrite','Console','PermissionTest')][string]$Kind,[Parameter(Mandatory)]$Data)
    $label=switch ($Kind) {
        'MailTest' {'Send a test message from '+$Data.From+' to '+$Data.To+' using '+$Data.Transport+'.'}
        'ShareWrite' {'Create and remove test marker '+$Data.Marker+' in '+$Data.Path+'; expected SID: '+$Data.Sid+'.'}
        'Console' {'Open orchestrator management console for '+$Data.Tenant+'; it may change shared configuration and logs. Script SHA-256: '+$Data.ScriptHash+'.'}
        'PermissionTest' {'Check Graph and Exchange permissions for '+$Data.Tenant+'; interactive authentication and consent may occur. Catalog: '+$Data.CatalogPath+'.'}
    }
    [pscustomobject]@{Kind=$Kind;Data=$Data;Secret=$null;Label=$label;State='Pending'}
}
function Invoke-SiPendingFileActions {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue)
    foreach ($action in @($Queue | Where-Object Kind -eq 'FileEdit')) {
        if ($action.State -ne 'Pending') {throw 'File action requires review.'}
        $d=$action.Data
        $null=New-SiPendingFileAction -Type $d.Type -Plan $d.Plan -Value $d.Value -Worker $d.Worker -Root $d.Root -ScriptPath $d.ScriptPath
    }
    foreach ($group in @($Queue | Where-Object Kind -eq 'FileEdit' | Group-Object {$_.Data.Path})) {
        $action=$group.Group[0];$d=$action.Data
        try {
            $result=switch ($d.Type) {
                'Orchestrator' {
                    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase);$values=[Collections.Generic.List[object]]::new()
                    foreach ($item in $group.Group) {$snap=$item.Data.Plan
                        if ($snap.Hash -cne $d.Hash -or -not $seen.Add($snap.Field)) {throw 'Orchestrator changes conflict.'}
                        $fresh=Get-SiOrchestratorSettingSnapshot -Path $snap.Path -Field $snap.Field -TemplatePath $snap.TemplatePath
                        foreach ($key in @('Hash','Kind','Value','Present','TemplateHash')) {if ($fresh.$key -cne $snap.$key) {throw 'Orchestrator configuration or template changed.'}}
                        $typed=ConvertTo-SiOrchestratorSettingValue -Snapshot $fresh -Value $item.Data.Value
                        $values.Add([pscustomobject]@{Field=$fresh.Field;Value=$typed})
                    }
                    $read=Read-SmartM365JsonDocument -Path $d.Path
                    if ($read.Coexistence -ne 'Single' -or $read.Path -cne $d.Path -or $read.SHA256 -cne $d.Hash) {throw 'Orchestrator configuration changed.'}
                    foreach ($item in $values) {if ($null -eq $read.Document.PSObject.Properties[$item.Field]) {Add-Member -InputObject $read.Document -NotePropertyName $item.Field -NotePropertyValue $item.Value} else {$read.Document.($item.Field)=$item.Value}}
                    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($read.Document | ConvertTo-Json -Depth 100)+[Environment]::NewLine)
                    Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $bytes -ExpectedSHA256 $d.Hash -Validate {param($doc) if ($doc -isnot [pscustomobject]) {throw 'Expected JSON object.'}} | Out-Null
                    [pscustomobject]@{State='Saved'}
                }
                'Reconciliation' {Invoke-SiBootstrapReconciliation -Plan $d.Plan}
            }
            if ($null -eq $result -or ($result.PSObject.Properties['State'] -and $result.State -ne 'Saved')) {throw 'The saved file could not be confirmed.'}
            foreach ($item in $group.Group) {[void]$Queue.Remove($item)}
        } catch {foreach ($item in $group.Group) {$item.State='Uncertain'};throw ('File action stopped; inspect '+$d.Path+' before retrying. '+$_.Exception.Message)}
    }
}
function Clear-SiPendingCreations {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue)
    foreach ($action in $Queue) {if ($null -ne $action.Secret) {$action.Secret.Dispose()}}
    $Queue.Clear()
}
function New-SiPendingTaskIntent {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,63}\z')][string]$Tenant,[Parameter(Mandatory)][string]$Executable,[Parameter(Mandatory)][string]$AccountLabel,[bool]$EnableAfterCreation=$false)
    if (-not [IO.Path]::IsPathFullyQualified($Root) -or -not [IO.Path]::IsPathFullyQualified($Executable) -or [string]::IsNullOrWhiteSpace($AccountLabel) -or $AccountLabel -match '[\r\n]') {throw 'Task inputs are invalid.'}
    $installer=Join-Path $Root 'SmartInventory/Orchestrator/Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'
    $orchestrator=Join-Path $Root 'SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
    foreach ($path in @($installer,$orchestrator,$Executable)) {if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {throw 'A required task installer or executable is missing.'}}
    $data=[pscustomobject]@{Root=$Root;Tenant=$Tenant;Executable=$Executable;InstallerHash=(Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash;AccountLabel=$AccountLabel;EnableAfterCreation=$EnableAfterCreation}
    $mode=if ($EnableAfterCreation) {'then enable it after verification'} else {'leave it disabled'}
    [pscustomobject]@{Kind='Task';Data=$data;Secret=$null;Label=('Create task for '+$AccountLabel+' using the existing installer; '+$mode+'; no replacement or immediate start.');State='Pending'}
}
function New-SiPendingPrerequisiteIntent {
    param([Parameter(Mandatory)]$Plan)
    if ($Plan.Scope -notin @('CurrentUser','AllUsers') -or $Plan.TrustRepository -isnot [bool] -or -not (Test-Path -LiteralPath $Plan.Path -PathType Leaf) -or (Get-FileHash -LiteralPath $Plan.Path -Algorithm SHA256).Hash -cne $Plan.SHA256) {throw 'The prerequisite installation plan is invalid or changed.'}
    [pscustomobject]@{Kind='Prerequisites';Data=$Plan;Secret=$null;Label=('Install missing modules for '+$Plan.Scope+'; trust PSGallery: '+$Plan.TrustRepository+'; installer SHA-256: '+$Plan.SHA256);State='Pending'}
}
function New-SiPendingBootstrapIntent {
    param([Parameter(Mandatory)]$Plan,[Parameter(Mandatory)][string]$Parent)
    if (-not [IO.Path]::IsPathFullyQualified($Parent) -or -not (Test-Path -LiteralPath $Parent -PathType Container)) {throw 'Select an existing absolute staging folder.'}
    $fresh=New-SiBootstrapStagePlan -Root $Plan.Root -BootstrapPlan $Plan.BootstrapPlan
    foreach ($file in $fresh.Files.Keys) {if ($fresh.Files[$file] -cne $Plan.Files[$file]) {throw 'A bootstrap dependency changed; review again.'}}
    [pscustomobject]@{Kind='Bootstrap';Data=[pscustomobject]@{Plan=$Plan;Parent=$Parent};Secret=$null;Label=('Run bootstrap for '+$Plan.BootstrapPlan.Directory+' / '+$Plan.BootstrapPlan.DisplayName+'; Teams and SharePoint: '+$Plan.BootstrapPlan.ConfigureTeams+'; isolated folder: '+$Parent+'. Creates or changes tenant resources; conflict-free generated settings are applied automatically.');State='Pending'}
}
function Start-SiPendingPrerequisiteAction {
    param([Parameter(Mandatory)]$Action,[scriptblock]$Start)
    if ($Action.Kind -ne 'Prerequisites' -or $Action.State -ne 'Pending') {throw 'Prerequisite action requires review.'}
    $null=New-SiPendingPrerequisiteIntent -Plan $Action.Data
    if ($null -eq $Start) {$Start={param($p) Start-SiPrerequisiteInstallation -Plan $p}}
    try {$run=& $Start $Action.Data;if ($null -eq $run) {throw 'No prerequisite process returned.'};$Action.State='Running';return $run} catch {$Action.State='Uncertain';throw 'Prerequisite launch could not be confirmed. Inspect modules before retrying.'}
}
function Start-SiPendingTaskAction {
    param([Parameter(Mandatory)]$Action,[Parameter(Mandatory)]$Identity,[scriptblock]$ListTasks,[scriptblock]$Start,[string[]]$AllowedExecutables=@())
    if ($Action.Kind -ne 'Task' -or $Action.State -ne 'Pending' -or $Identity.Managed -ne $false -or [string]$Identity.Sid -notmatch '^S-1-(?:\d+-)+\d+$') {throw 'The task action or standard service identity is unavailable.'}
    if ($Action.Data.AccountLabel.StartsWith('Pending account:',[StringComparison]::Ordinal)) {
        $name=$Action.Data.AccountLabel.Substring('Pending account:'.Length).Trim()
        if (-not $name -or ([string]$Identity.Account -notmatch ('(?i)(?:\\'+[regex]::Escape($name)+'$|^'+[regex]::Escape($name)+'@)'))) {throw 'The created service account does not match the reviewed task.'}
    } elseif ($Action.Data.AccountLabel -ine [string]$Identity.Account) {throw 'The selected service account changed; review the task again.'}
    $installer=Join-Path $Action.Data.Root 'SmartInventory/Orchestrator/Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'
    if ((Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash -cne $Action.Data.InstallerHash) {throw 'The task installer changed; review again.'}
    if ($null -eq $ListTasks) {$ListTasks={Get-ScheduledTask -ErrorAction Stop}}
    $plan=New-SiOrchestratorTaskPlan -Tenant $Action.Data.Tenant -SmartM365Root $Action.Data.Root -Executable $Action.Data.Executable -Account $Identity.Account -Sid $Identity.Sid -Tasks @(& $ListTasks) -AllowedExecutables $AllowedExecutables -Managed $false
    $plan | Add-Member -NotePropertyName InstallerHash -NotePropertyValue $Action.Data.InstallerHash
    if ($null -eq $Start) {$Start={param($p) Start-SiOrchestratorRegistration -Plan $p}}
    try {$run=& $Start $plan;if ($null -eq $run) {throw 'No task installer process returned.'};$Action.State='Running';return $run} catch {$Action.State='Uncertain';throw 'Task installer launch could not be confirmed. Inspect Task Scheduler before retrying.'}
}
function Get-SiPendingCreationReview {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue)
    if (-not $Queue.Count) {return 'No actions are pending.'}
    return ((@($Queue | ForEach-Object {
        $line=$_.Label+' ['+$_.State+']'
        if ($_.Kind -eq 'Global') {
            $model=$_.Data.Review | ConvertFrom-Json -ErrorAction Stop
            foreach ($edit in $_.Data.PendingEdits) {
                if ($edit.Step -ne 'share' -or $edit.Field -notin @(Get-SiEditableFields share) -or $null -eq $model.PSObject.Properties[$edit.Field]) {throw 'Invalid pending global setting.'}
                $model.($edit.Field)=ConvertTo-SiEditValue -Snapshot $edit -Value $edit.Value
            }
            foreach ($field in @('DefaultTenant','WorkspaceRootPath','OnPremExecutionMode','RetentionMaxCSV','RetentionMaxLogs')) {
                $property=$model.PSObject.Properties[$field]
                if ($null -ne $property -and $null -ne $property.Value) {$line+=[Environment]::NewLine+'  '+$field+': '+[string]$property.Value}
            }
            foreach ($edit in $_.Data.PendingEdits) {
                if ($edit.Field -notin @('DefaultTenant','WorkspaceRootPath','OnPremExecutionMode','RetentionMaxCSV','RetentionMaxLogs')) {
                    $value=if ($edit.Field -match '(?i)url|password|secret|token|webhook|credential|private.?key') {'(configured; masked)'} else {[string]$model.($edit.Field)}
                    $line+=[Environment]::NewLine+'  '+$edit.Field+': '+$value
                }
            }
            $line+=[Environment]::NewLine+'  Other global settings use the template defaults.'
        }
        if ($_.Kind -eq 'Profile') {
            foreach ($field in @('OrgDomain','SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath')) {
                $property=$_.Data.Model.PSObject.Properties[$field]
                $pending=@($_.Data.PendingEdits | Where-Object Field -eq $field | Select-Object -Last 1)
                if ($pending.Count) {$line+=[Environment]::NewLine+'  '+$field+': '+$pending[0].Value}
                elseif ($null -ne $property -and $property.Value -is [string]) {$line+=[Environment]::NewLine+'  '+$field+': '+$property.Value}
            }
        }
        if ($_.Kind -eq 'Profile') {foreach ($edit in $_.Data.PendingEdits) {if ($edit.Field -notin @('OrgDomain','SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath')) {$line+=[Environment]::NewLine+'  '+$edit.Field+': '+$edit.Value}}}
        $line
    })) -join [Environment]::NewLine)
}
function Invoke-SiPendingCreations {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue,[scriptblock]$Execute)
    if (@($Queue | Where-Object State -ne 'Pending').Count) {throw 'An earlier action has an uncertain result. Inspect it and discard the action before retrying.'}
    foreach ($action in $Queue) {
        if ($action.Kind -in @('Global','Profile')) {
            $check=New-SiPendingCreation -Kind $action.Kind -Data $action.Data
            if ($check.Label -cne $action.Label) {throw 'A creation plan changed; review again.'}
        }
    }
    if ($null -eq $Execute) {$Execute={param($action)
        switch ($action.Kind) {
            'Global' {New-SiGlobalConfiguration -Plan $action.Data -Edits @($action.Data.PendingEdits)}
            'Profile' {New-SiTenantProfile -Root $action.Data.Root -ProfileKey $action.Data.ProfileKey -OrganizationKey $action.Data.OrganizationKey -EnvironmentKey $action.Data.EnvironmentKey -ExpectedTemplateHash $action.Data.TemplateHash -Edits @($action.Data.PendingEdits)}
            'LocalAccount' {New-SiLocalServiceAccount -Name $action.Data.Name -Password $action.Secret}
            'DomainAccount' {New-SiDirectoryServiceAccount -Kind DomainUser -Name $action.Data.Name -DomainController $action.Data.DomainController -OrganizationalUnit $action.Data.OrganizationalUnit -Password $action.Secret}
        }
    }}
    $completed=[Collections.Generic.List[object]]::new()
    $ordered=@($Queue | Where-Object {$_.Kind -notin @('Task','Prerequisites','SigningCertificate','Bootstrap','FileEdit','MailTest','ShareWrite','Console','PermissionTest')} | Sort-Object @{Expression={switch ($_.Kind) {'Global' {0} 'Profile' {1} default {2}}}})
    foreach ($action in $ordered) {
        try {
            $result=& $Execute $action
            if ($null -eq $result) {throw 'No result was returned.'}
            $completed.Add([pscustomobject]@{Kind=$action.Kind;Result=$result})
            if ($null -ne $action.Secret) {$action.Secret.Dispose()};[void]$Queue.Remove($action)
        } catch {$action.State='Uncertain';if ($null -ne $action.Secret) {$action.Secret.Dispose();$action.Secret=$null};throw ('Apply stopped after '+$completed.Count+' creation action(s). Inspect the result before retrying: '+$_.Exception.Message)}
    }
    [pscustomobject]@{Completed=$completed.ToArray()}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqypm75TuiJUN+
# AJRapk9VJohe2d3lQo2k+02AUU9XEqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFklyiFhG1rfY10rJER3GJ7RwHKq/awoM/0NV2hq0ha1MA0GCSqG
# SIb3DQEBAQUABIIBgBBpDF/C0HSrEqUmmEz5Y9iXV4mOqTw372r2LD4u6UB6oVuT
# owpZVzwUmm/0tfgn+at4u7hr+TMSEHtdMkcQSH9euiEFZrcWy4+zJLda8yPpriYO
# JZrCnVhuZY1vD+hie9FQQiBfGlMDEmNiT//zSy0eYTNHZU7uXBhWNc6DRAqOq/VP
# 8JPfryUKTzXsVAF/xhemZOz52G1B8CwoefmdLbvCJRBGiF0vmKJ1Eld8cwtRPp7s
# F6vSAVqn9H0RMuW8gBoKe0RK3NlILNW07BzD2FtlP1fl0pVBZ+ExOkiIHHBbelIj
# yHAkv0I7Y6ikHcNIoHLEFyaR/P6NmeoMaD1G/sdJY9UgnPPEotVn+XBU3yNao4Wv
# Qs7CwkkHztcAyqDepv6f66IkO9HcxUP9A0r2uqbGz6SAkCVbIVhjUh7G3k+bxd4V
# /hO4vV4G9oixH4UfUI5Y/jB/5XNCD82oFR3NYMsvlJR9LjJ99MXj5xLHJdY6o8HL
# 9lQzE1RFNgTawze5r6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjdaMC8GCSqGSIb3DQEJBDEiBCBvqGE2jL09ailvJFjFUQCzz/xNUNwsusGwn5iV
# Vx5SxzANBgkqhkiG9w0BAQEFAASCAgAX0aLQVHuMlwqkuAT5I9qxilV/SMud9IU3
# gJdgPhMgl7cXzLzfpo74mmBLc2QPduyB8mxXY+FTJRGgOtEvpuR1LYp5z5PB2Cy8
# 2Z6LBB5EyAfbmQX7ftoyj3DLy6WWZQZxn8PZrJCbHgOl7pJM6QBeo5+TpOEt2Qeo
# IfqEjDxMHBZJj2RNGR+iskKZEQXMq96578JqY1ilhKMVj8w//zWFb5dcMepgCBDU
# ib4ROZFshR6bKs63Luj5/UWEQTHeL8qyKiLllvv/tj+FYs45wFTEE/ruaGSAnbVM
# agiPasrsSy90nBO8JBEHNtaEh2WSt91touVZeaOybuV137eZYQHNqvPNmHmJghnt
# Bd4WCqX7yLS+X3yD4N0TKqNlanNQ4D9uHIRkyCBrVZ3q9cyPMv3E6UL7oNy0tiRd
# bVopSp48R5y2ArMA3TK7pZJ5teJfEZ1QDusIgtNLjTFtzUTtwpUEL8W4ux7XE7xg
# bktNOhVVcLUKb9KZ3Qhi3Zaaa/VAA2lNlT9yO0eswvLxwYjyNIfcdcvLZU0gHBDR
# czP2SwtSUfbb/fBDRZql0lEu3pFQxjzqnZo2EwYwt/nncNP50wGJrsKAxDqDKBsy
# GAmz3cFc86wwnJQ0D6N5DjFkdTvfSZRbAgk315QIN1YzlWdRkiLEz6s1iBgL4O5s
# vjPnFntdow==
# SIG # End signature block
