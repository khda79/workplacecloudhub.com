<#
.SYNOPSIS
    Read-only source assessment helpers. Loaded by source inventories; no collection at import.
.VERSION
    1.0.1
#>

function Get-SPSAProperty {
    param($Object, [string]$Name, [switch]$Optional)
    if ($null -eq $Object -or -not $Object.PSObject.Properties[$Name]) {
        if ($Optional) { return $null }
        throw "Assessment property unavailable: $Name"
    }
    return $Object.$Name
}

function Merge-SPSAConfig {
    param($Default, $Current)
    foreach ($property in $Default.PSObject.Properties) {
        if (-not $Current.PSObject.Properties[$property.Name]) {
            $Current | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
        } elseif ($property.Value -is [pscustomobject] -and $Current.($property.Name) -is [pscustomobject]) {
            Merge-SPSAConfig $property.Value $Current.($property.Name)
        }
    }
}

function Write-SPSAJson {
    param([string]$Path, $Value)
    $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json -Depth 25), [Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $temporary -Destination $Path -Force -ErrorAction Stop
    } finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}

function New-SPSAAssessment {
    param([ValidateSet('File','Permission')][string]$Kind, [string]$RulesDirectory)
    if (-not $RulesDirectory) {
        foreach ($candidate in @((Join-Path $PSScriptRoot '..\..\Config'), (Join-Path $PSScriptRoot '..\Config'))) {
            if (Test-Path -LiteralPath (Join-Path $candidate 'source-assessment.rules.json.template')) { $RulesDirectory = $candidate; break }
        }
    }
    $template = Join-Path $RulesDirectory 'source-assessment.rules.json.template'
    $runtime = Join-Path $RulesDirectory 'source-assessment.rules.json.txt'
    $default = Get-Content -LiteralPath $template -Raw -Encoding UTF8 | ConvertFrom-Json
    $rules = if (Test-Path -LiteralPath $runtime) { Get-Content -LiteralPath $runtime -Raw -Encoding UTF8 | ConvertFrom-Json } else { $default }
    $before = $rules | ConvertTo-Json -Depth 25 -Compress
    Merge-SPSAConfig $default $rules
    if (-not (Test-Path -LiteralPath $runtime) -or $before -cne ($rules | ConvertTo-Json -Depth 25 -Compress)) { Write-SPSAJson $runtime $rules }
    foreach ($rule in $rules.FunctionalRules.PSObject.Properties) {
        if (($rule.Value.Effort -isnot [int] -and $rule.Value.Effort -isnot [long]) -or $rule.Value.Effort -lt 0 -or $rule.Value.Effort -gt 3) { throw "Invalid assessment effort rule: $($rule.Name)" }
    }
    [pscustomobject]@{
        Kind=$Kind; Rules=$rules; RulesSha256=(Get-FileHash -LiteralPath $runtime -Algorithm SHA256).Hash
        Items=@{}; Errors=[Collections.Generic.List[object]]::new(); Detectors=[Collections.Generic.List[object]]::new()
        SeenWebs=@{}; SeenSites=@{}; StandardMasks=$null; Scope=[ordered]@{}; StartedAtUtc=[datetime]::UtcNow.ToString('o')
    }
}

function Add-SPSAError {
    param($Context, [string]$Detector, [string]$Scope, [string]$Message, [bool]$AffectsCollection=$true)
    $Context.Errors.Add([pscustomobject]@{Detector=$Detector;Scope=$Scope;Message=$Message;AffectsCollection=$AffectsCollection})
}

function Invoke-SPSADetector {
    param($Context, [string]$Name, [string]$Scope, [scriptblock]$Read)
    try { & $Read; $status='Complete' }
    catch { Add-SPSAError $Context $Name $Scope $_.Exception.Message; $status='Partial' }
    $Context.Detectors.Add([pscustomobject]@{Detector=$Name;Scope=$Scope;Status=$status})
}

function Add-SPSAComponent {
    param($Context, [string]$Id, [string]$Category, [string]$Technology, [string]$Evidence)
    if (-not $Id) { throw 'Assessment component ID is required.' }
    $rule = $Context.Rules.FunctionalRules.PSObject.Properties[$Technology]
    $effort=$null; $reason='No matching classification rule'; $classification='Unqualified'
    if ($rule) { $effort=[int]$rule.Value.Effort; $reason=[string]$rule.Value.Reason; $classification='Estimated' }
    $key=($Category+':'+$Id).ToLowerInvariant()
    if (-not $Context.Items.ContainsKey($key)) {
        $Context.Items[$key]=[pscustomobject]@{Id=$Id;Category=$Category;Technology=$Technology;Effort=$effort;Classification=$classification;Rule=$Technology;Reason=$reason;Evidence=$Evidence}
    }
}

function Get-SPSAWorkflowSubscriptions {
    param($Web)
    # Server object model; never start, publish or update a workflow.
    if (-not ('Microsoft.SharePoint.WorkflowServices.WorkflowServicesManager' -as [type])) { Add-Type -AssemblyName Microsoft.SharePoint.WorkflowServices -ErrorAction Stop }
    $manager = New-Object Microsoft.SharePoint.WorkflowServices.WorkflowServicesManager($Web)
    $service = $manager.GetWorkflowSubscriptionService()
    if (-not $service) { throw 'Workflow subscription service is unavailable; absence cannot be established.' }
    return $service.EnumerateSubscriptions()
}

function Read-SPSAEventReceivers {
    param($Context, $Owner, [string]$Scope)
    foreach ($receiver in @(Get-SPSAProperty $Owner 'EventReceivers')) {
        $assembly=[string](Get-SPSAProperty $receiver 'Assembly')
        if ($assembly -and $assembly -notmatch '^(Microsoft\.SharePoint|Microsoft\.Office)\b') {
            Add-SPSAComponent $Context ($Scope+':'+$receiver.Id) 'Customization' 'ServerCode' ("Event receiver: {0}; assembly: {1}" -f $Scope,$assembly)
        }
    }
}

function Read-SPSAActions {
    param($Context, $Owner, [string]$Scope)
    foreach ($action in @(Get-SPSAProperty $Owner 'UserCustomActions')) {
        $technology='UnknownCustomAction'
        if ((Get-SPSAProperty $action 'ScriptSrc' -Optional) -or (Get-SPSAProperty $action 'ScriptBlock' -Optional)) { $technology='ClientScript' }
        $clientSideId=[string](Get-SPSAProperty $action 'ClientSideComponentId' -Optional)
        if ($clientSideId -and $clientSideId -ne [guid]::Empty.ToString()) { $technology='ClientSideComponent' }
        if (Get-SPSAProperty $action 'ControlAssembly' -Optional) { $technology='ServerCode' }
        Add-SPSAComponent $Context ($Scope+':'+$action.Id) 'Customization' $technology ("Custom action: {0}; ID: {1}; location: {2}" -f $Scope,$action.Id,$action.Location)
    }
}

function Read-SPSAFeatures {
    param($Context, $Owner, [string]$Scope)
    foreach ($feature in @(Get-SPSAProperty $Owner 'Features')) {
        $definition=Get-SPSAProperty $feature 'Definition'
        $solution=[string](Get-SPSAProperty $definition 'SolutionId')
        $receiver=[string](Get-SPSAProperty $definition 'ReceiverAssembly' -Optional)
        if ($receiver -and $receiver -notmatch '^(Microsoft\.SharePoint|Microsoft\.Office)\b') {
            Add-SPSAComponent $Context ($Scope+':receiver:'+ $feature.DefinitionId) 'Customization' 'ServerCode' "Custom feature receiver: $Scope; $receiver"
        }
        if ($solution -and $solution -ne [guid]::Empty.ToString()) {
            # A custom feature alone does not prove full-trust code or its remediation effort.
            Add-SPSAComponent $Context ($Scope+':'+$feature.DefinitionId) 'Customization' 'CustomFeature' ("Activated custom feature: {0}; solution: {1}; scope: {2}" -f $feature.DefinitionId,$solution,$Scope)
        }
    }
}

function Get-SPSAPageItems {
    param($List)
    $query=New-Object Microsoft.SharePoint.SPQuery
    $query.ViewAttributes='Scope="RecursiveAll"'
    $query.RowLimit=2000
    $query.ViewFields='<FieldRef Name="FileRef"/><FieldRef Name="FSObjType"/>'
    $query.Query='<Where><And><Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq><Or><Eq><FieldRef Name="File_x0020_Type"/><Value Type="Text">aspx</Value></Eq><Eq><FieldRef Name="File_x0020_Type"/><Value Type="Text">xsn</Value></Eq></Or></And></Where>'
    do {
        $items=$List.GetItems($query)
        foreach($item in $items){$item}
        $query.ListItemCollectionPosition=$items.ListItemCollectionPosition
    } while($query.ListItemCollectionPosition)
}

function Get-SPSARenderingProperties {
    param($Owner)
    [pscustomobject]@{
        JSLink=[string](Get-SPSAProperty $Owner 'JSLink' -Optional)
        XslLink=[string](Get-SPSAProperty $Owner 'XslLink' -Optional)
        Xsl=[string](Get-SPSAProperty $Owner 'Xsl' -Optional)
    }
}

function Read-SPSARenderingOverrides {
    param($Context, $Owner, [string]$Scope)
    $rendering=Get-SPSARenderingProperties $Owner
    # clienttemplates.js is the built-in renderer, not a custom dependency.
    $scripts=@($rendering.JSLink -split '\|' | Where-Object { $_.Trim() -and $_.Trim() -ine 'clienttemplates.js' })
    if($scripts.Count) {
        Add-SPSAComponent $Context ($Scope+':JSLink') 'Customization' 'ClientScript' ("Custom JSLink at {0}: {1}" -f $Scope,($scripts -join '|'))
    }
    if(-not [string]::IsNullOrWhiteSpace($rendering.XslLink) -or -not [string]::IsNullOrWhiteSpace($rendering.Xsl)) {
        Add-SPSAComponent $Context ($Scope+':Xsl') 'Customization' 'CustomXslt' ("Rendering override at {0}; XslLink={1}; inline XSL present={2}; compatibility needs qualification" -f $Scope,$rendering.XslLink,(-not [string]::IsNullOrWhiteSpace($rendering.Xsl)))
    }
}

function Read-SPSAPage {
    param($Context, $Web, $File)
    $url=[string](Get-SPSAProperty $File 'ServerRelativeUrl')
    if ($url -match '\.xsn$') { Add-SPSAComponent $Context $url 'Form' 'InfoPath' "InfoPath template: $url"; return }
    if ($url -notmatch '\.(aspx|master)$') { return }
    if ([string](Get-SPSAProperty $File 'CustomizedPageStatus') -eq 'Customized') {
        Add-SPSAComponent $Context $url 'Customization' 'CustomizedPage' "Customized page: $url; compatibility needs qualification"
    }
    if ($url -match '\.master$') { return }
    foreach ($part in @(Get-SPSAPageWebParts $Web $url)) {
        $type=$part.TypeName
        $assembly=$part.AssemblyName
        $technology='UnknownWebPart'
        if ($Context.Rules.WebPartTypes.PSObject.Properties[$type]) { $technology=[string]$Context.Rules.WebPartTypes.$type }
        elseif ($assembly -notmatch '^(Microsoft\.SharePoint|Microsoft\.Office)\b') { $technology='ServerCode' }
        Add-SPSAComponent $Context ($url+':'+$part.Id) 'Customization' $technology ("Web Part: {0}; type: {1}; assembly: {2}" -f $url,$type,$assembly)
        Read-SPSARenderingOverrides $Context $part ($url+':'+$part.Id)
    }
}

function Get-SPSAPageWebParts {
    param($Web, [string]$Url)
    $manager=$null
    try {
        $manager=$Web.GetLimitedWebPartManager($Url, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
        foreach ($part in $manager.WebParts) {
            $rendering=Get-SPSARenderingProperties $part
            [pscustomobject]@{Id=[string]$part.ID;TypeName=$part.GetType().FullName;AssemblyName=$part.GetType().Assembly.GetName().Name;JSLink=$rendering.JSLink;XslLink=$rendering.XslLink;Xsl=$rendering.Xsl}
        }
    } finally { if ($manager -is [IDisposable]) { $manager.Dispose() } }
}

function Read-SPSAList {
    param($Context, $Web, $List)
    $scope=[string](Get-SPSAProperty (Get-SPSAProperty $List 'RootFolder') 'ServerRelativeUrl')
    Invoke-SPSADetector $Context 'ListWorkflows2010' $scope {
        foreach ($workflow in @(Get-SPSAProperty $List 'WorkflowAssociations')) {
            Add-SPSAComponent $Context ($scope+':'+$workflow.Id) 'Workflow' 'Workflow2010' "List workflow association: $scope; $($workflow.Id)"
        }
    }
    Invoke-SPSADetector $Context 'Forms' $scope {
        if ([int](Get-SPSAProperty $List 'BaseTemplate') -eq 115) { Add-SPSAComponent $Context $scope 'Form' 'InfoPath' "XML form library: $scope" }
        foreach ($contentType in @(Get-SPSAProperty $List 'ContentTypes')) {
            $template=[string](Get-SPSAProperty $contentType 'DocumentTemplateUrl')
            if ($template -match '\.xsn(?:\?|$)') { Add-SPSAComponent $Context $template 'Form' 'InfoPath' "Content type InfoPath template: $template" }
            foreach ($property in @('NewFormUrl','EditFormUrl','DisplayFormUrl')) {
                $form=[string](Get-SPSAProperty $contentType $property)
                if ($form) { Add-SPSAComponent $Context ($scope+':'+$contentType.Id+':'+$property) 'Form' 'CustomForm' "Explicit content type form override: $scope; $property=$form" }
            }
            Read-SPSAEventReceivers $Context $contentType ($scope+':'+$contentType.Id)
        }
    }
    Invoke-SPSADetector $Context 'ListCustomizations' $scope {
        Read-SPSAActions $Context $List $scope
        Read-SPSAEventReceivers $Context $List $scope
        foreach($field in @(Get-SPSAProperty $List 'Fields')) {
            $type=[string](Get-SPSAProperty $field 'TypeAsString')
            if($type -notin @($Context.Rules.StandardFieldTypes)) { Add-SPSAComponent $Context ($scope+':'+$field.Id) 'Customization' 'CustomField' "Field type $type at $scope" }
            Read-SPSARenderingOverrides $Context $field ($scope+':'+$field.Id)
        }
    }
    Invoke-SPSADetector $Context 'PagesAndTemplates' $scope {
        # Includes hidden libraries and technical forms folders; never sample or read document binaries.
        if ([string](Get-SPSAProperty $List 'BaseType') -eq 'DocumentLibrary') {
            foreach($item in @(Get-SPSAPageItems $List)){ Read-SPSAPage $Context $Web $item.File }
        }
        foreach($form in @(Get-SPSAProperty $List 'Forms')) { Read-SPSAPage $Context $Web ($Web.GetFile($form.ServerRelativeUrl)) }
    }
}

function Add-SPSAWeb {
    param($Context, $Web)
    if (-not $Context) { return }
    $url=[string]$Web.Url
    if($Context.SeenWebs.ContainsKey($url)){return};$Context.SeenWebs[$url]=$true
    if($Context.Kind -eq 'Permission') { return }
    Invoke-SPSADetector $Context 'WebWorkflows2010' $url {
        foreach($workflow in @(Get-SPSAProperty $Web 'WorkflowAssociations')) { Add-SPSAComponent $Context ($url+':'+$workflow.Id) 'Workflow' 'Workflow2010' "Web workflow association: $url; $($workflow.Id)" }
    }
    Invoke-SPSADetector $Context 'Workflows2013' $url {
        foreach($workflow in @(Get-SPSAWorkflowSubscriptions $Web)) { Add-SPSAComponent $Context ($url+':'+$workflow.Id) 'Workflow' 'Workflow2013' "Workflow subscription: $url; $($workflow.Id)" }
    }
    Invoke-SPSADetector $Context 'WebCustomizations' $url {
        Read-SPSAActions $Context $Web $url
        Read-SPSAEventReceivers $Context $Web $url
        Read-SPSAFeatures $Context $Web $url
        foreach($property in @('MasterUrl','CustomMasterUrl')) {
            $master=[string](Get-SPSAProperty $Web $property -Optional)
            if($master) {
                Read-SPSAPage $Context $Web ($Web.GetFile($master))
                if([IO.Path]::GetFileName($master) -notin @('seattle.master','oslo.master','v4.master','default.master','minimal.master')) {
                    Add-SPSAComponent $Context $master 'Customization' 'CustomMasterPage' "Explicit master page reference: $url; $master"
                }
            }
        }
    }
    Invoke-SPSADetector $Context 'SiteAssessment' $url {
        $site=Get-SPSAProperty $Web 'Site'
        $siteUrl=[string](Get-SPSAProperty $site 'Url')
        if(-not $Context.SeenSites.ContainsKey($siteUrl)) {
            $Context.SeenSites[$siteUrl]=$true
            Invoke-SPSADetector $Context 'SiteEventReceivers' $siteUrl { Read-SPSAEventReceivers $Context $site $siteUrl }
            Invoke-SPSADetector $Context 'SiteCustomizations' $siteUrl {
                Read-SPSAFeatures $Context $site $siteUrl
                Read-SPSAActions $Context $site $siteUrl
            }
        }
    }
    Invoke-SPSADetector $Context 'ListsIncludingHiddenAndNonDocument' $url {
        $listIndex=0
        foreach($list in @(Get-SPSAProperty $Web 'Lists')) {
            $listIndex++
            Invoke-SPSADetector $Context 'ListAssessment' ($url+':list-index:'+ $listIndex) { Read-SPSAList $Context $Web $list }
        }
    }
}

function Get-SPSAStandardRoleMasks {
    param($Context)
    $masks=@{}
    foreach($p in $Context.Rules.StandardRoleMasks.PSObject.Properties){$masks[$p.Name]=[string]$p.Value}
    # Profiles mirror documented SharePoint Server default rights. Resolve enum bits
    # using the installed server assembly, never using a possibly modified web role.
    foreach($profile in $Context.Rules.StandardRolePermissions.PSObject.Properties) {
        if($masks.ContainsKey($profile.Name)){continue}
        $permissions=@($profile.Value) -join ', '
        $masks[$profile.Name]=([uint64][Microsoft.SharePoint.SPBasePermissions]([Enum]::Parse([Microsoft.SharePoint.SPBasePermissions],$permissions))).ToString([Globalization.CultureInfo]::InvariantCulture)
    }
    if(-not $masks.Count){throw 'No trusted standard-role baseline is configured.'}
    return $masks
}

function Add-SPSARoleBindings {
    param($Context, $RoleAssignments, [string]$Scope, [string]$WebUrl='')
    if(-not $Context){return}
    try {
        foreach($assignment in $RoleAssignments){
            foreach($role in $assignment.RoleDefinitionBindings){
                $type=[string](Get-SPSAProperty $role 'Type')
                $mask=([uint64](Get-SPSAProperty $role 'BasePermissions')).ToString([Globalization.CultureInfo]::InvariantCulture)
                $roleId=[string](Get-SPSAProperty $role 'Id')
                $key=$WebUrl+':'+$roleId+':'+$type+':'+$mask
                if($Context.Items.ContainsKey($key)){$Context.Items[$key].Occurrences++;continue}
                $builtIn=$false;$classification='Detected'
                if($type -ne 'None' -and $type -ne '0'){
                    if($null -eq $Context.StandardMasks){
                        try{$Context.StandardMasks=Get-SPSAStandardRoleMasks $Context}
                        catch{$Context.StandardMasks=@{};Add-SPSAError $Context 'StandardRoleBaseline' $Scope $_.Exception.Message $false}
                    }
                    if($Context.StandardMasks.ContainsKey($type)){$builtIn=($Context.StandardMasks[$type] -eq $mask)}
                    else{$classification='Unqualified';Add-SPSAError $Context 'RoleClassification' $Scope "No trusted baseline for role type $type" $false}
                }
                if(-not $Context.Items.ContainsKey($key)){$Context.Items[$key]=[pscustomobject]@{Id=$key;WebUrl=$WebUrl;RoleId=$roleId;Name=[string](Get-SPSAProperty $role 'Name');Occurrences=1;IsBuiltIn=$builtIn;IsUsed=$true;EffectivePermissions=$mask;RoleType=$type;Classification=$classification;Rule='Documented standard role type and effective rights';Evidence="Used role definition at $Scope; role ID $roleId"}}
            }
        }
    } catch { Add-SPSAError $Context 'UsedPermissionDefinitions' $Scope $_.Exception.Message }
}

function Complete-SPSAAssessment {
    param($Context, [string]$CsvPath)
    if(-not $Context){return}
    # Sidecar publication is separate from the primary CSV; never delete or mutate the CSV here.
    try{
        $items=@($Context.Items.Values|Sort-Object Id)
        $unqualified=@($items|Where-Object Classification -eq 'Unqualified').Count
        $collectionComplete=(@($Context.Errors|Where-Object AffectsCollection).Count -eq 0 -and $Context.SeenWebs.Count -gt 0)
        $qualified=($unqualified -eq 0 -and @($Context.Errors|Where-Object {-not $_.AffectsCollection}).Count -eq 0)
        $section=[ordered]@{
            Complete=($collectionComplete -and $qualified);CollectionComplete=$collectionComplete;QualificationComplete=$qualified
            ReviewedAtUtc=[datetime]::UtcNow.ToString('o');Evidence='Automatic source inventory assessment; see Items, Detectors and Errors'
            Items=$items;Errors=@($Context.Errors.ToArray());Detectors=@($Context.Detectors.ToArray());Scope=$Context.Scope
            Mode='Automatic';Estimated=($Context.Kind -eq 'File');UnqualifiedCount=$unqualified
        }
        $key=if($Context.Kind -eq 'File'){'FunctionalComponents'}else{'CustomPermissions'}
        $hashKey=if($Context.Kind -eq 'File'){'FileInventorySha256'}else{'PermissionInventorySha256'}
        $result=[ordered]@{SchemaVersion=1;ProducerVersion='1.0.1';TargetPlatform=[string]$Context.Rules.TargetPlatform;GeneratedAtUtc=[datetime]::UtcNow.ToString('o');RulesSha256=$Context.RulesSha256;Inputs=@{} }
        $result.Inputs[$hashKey]=(Get-FileHash -LiteralPath $CsvPath -Algorithm SHA256).Hash
        $result[$key]=$section
        $sidecar=$CsvPath+'.assessment.json.txt'
        $detailPath=Join-Path (Split-Path $CsvPath -Parent) ('AssessmentDetails-'+[IO.Path]::GetFileNameWithoutExtension($CsvPath)+'.csv')
        $result['DetailsFile']=[IO.Path]::GetFileName($detailPath)
        if($items.Count){$items|Export-Csv -LiteralPath $detailPath -Delimiter ';' -NoTypeInformation -Encoding UTF8}
        else{[IO.File]::WriteAllText($detailPath,'"Id";"Evidence"'+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))}
        $result['DetailsSha256']=(Get-FileHash -LiteralPath $detailPath -Algorithm SHA256).Hash
        Write-SPSAJson $sidecar $result
        Write-Host ('{0} Automatic source assessment: {1}; collection complete={2}; qualification complete={3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$sidecar,$collectionComplete,$qualified)
    }catch{Write-Warning ('Automatic source assessment was not published: '+$_.Exception.Message)}
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAMYbUgmar6xiHl
# hUY0x77NwMrWBkLsPov2H7Kf6V4Bw6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDr2zTWfkZ3FNUutQJ5ptc9
# EiSeuXLv9AVcj+5n1CLBaTANBgkqhkiG9w0BAQEFAASCAYAyDQUhebd0sDQNt72k
# 8DP98Mp1Z6z2BrGTJ9xQeXyXuFFqoDEbTY2KDWn8yriO2DHfbqteVxTd5noFiEtr
# Kx8QLHWC1AT9AtwEw9LQ3fKyA0brKZdPfNqTEUWTEMP9JZj7h3iJT23DXg1+QssN
# Bbue7swDCYjnY4BePgSVk9+8vrMiXTUYpkN6s23nV4EWNckEB7kvJ8NcRtI0gozl
# uVuMSVVAz2nltNZeradiNRqLmwpKpEWvGgWSysmAv+tdb4wzE93lfecJuXiOrwBV
# ejJtg8OzFYXBshdWYBy6exZhziSNFcDxNp0Qm1q1bpRkHhgJ8RmWAsKa70syzxuU
# x5BGpL5W223W5dbt+Z01QAx0fXa8KR1O7jvzOT03DVderS5E6827Y71yDqkkvD2+
# c/XjxNekjS8pmw8yvEb44LKvSE6Qjr4VyUmWPaLubgV0E00XcIbJ7HxBUXAjJCHl
# DqtqN8uZB2ydtdUjLVl5VcFXXy2hSrdIWKxe6/sZ9Bqk/QQ=
# SIG # End signature block
