#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
.SYNOPSIS
    Pester 5 tests for SmartM365.IntunePackager.Graph.psm1. Graph and Azure Storage are mocked;
    no network request is sent.

.VERSION
    1.0.6
#>

BeforeAll {
    $script:AppRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Core.psm1') -Force
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Graph.psm1') -Force
    . (Join-Path $PSScriptRoot 'SmartM365-IntunePackager-TestHelpers.ps1')
    Set-IpLogSink -Sink $null -Console $false
    Set-IpLogPath -Path ''
    Set-IpGraphSettings -BaseUri 'https://graph.microsoft.com/beta' -MaxRetries 3 -PollSeconds 1 -PollTimeoutMinutes 1
    $script:ModuleName = 'SmartM365.IntunePackager.Graph'
    $script:AppId = 'aaaaaaaa-0000-4000-8000-000000000001'

    function New-HttpException {
        param([int]$StatusCode, [int]$RetryAfterSeconds = 0)
        $exception = [Exception]::new("Response status code does not indicate success: $StatusCode.")
        $headers = [pscustomobject]@{ RetryAfter = $null }
        if ($RetryAfterSeconds -gt 0) { $headers.RetryAfter = [pscustomobject]@{ Delta = [TimeSpan]::FromSeconds($RetryAfterSeconds); Date = $null } }
        $exception | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = $StatusCode; Headers = $headers })
        return $exception
    }
}

Describe 'Scope and connection' {
    It 'requests only DeviceManagementApps.ReadWrite.All' {
        Get-IpGraphScopes | Should -Be @('DeviceManagementApps.ReadWrite.All')
    }

    It 'disconnects before connecting with delegated authentication' {
        $script:Calls = [Collections.Generic.List[string]]::new()
        Mock -ModuleName $script:ModuleName Assert-IpGraphModule { }
        Mock -ModuleName $script:ModuleName Disconnect-MgGraph { $script:Calls.Add('Disconnect') }
        Mock -ModuleName $script:ModuleName Connect-MgGraph { $script:Calls.Add('Connect') }
        Mock -ModuleName $script:ModuleName Get-MgContext { [pscustomobject]@{ TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; Account = 'admin@contoso.example'; ClientId = ''; AuthType = 'Delegated'; Scopes = @('DeviceManagementApps.ReadWrite.All') } }
        $profile = [pscustomobject]@{ Name = 'test'; TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; AuthMode = 'Delegated'; ClientId = ''; CertificateThumbprint = '' }
        $result = Connect-IpGraph -TenantProfile $profile
        $script:Calls | Should -Be @('Disconnect', 'Connect')
        $result.Identity | Should -Be 'admin@contoso.example'
        Should -Invoke -ModuleName $script:ModuleName Connect-MgGraph -Times 1 -ParameterFilter { ($Scopes -join ',') -eq 'DeviceManagementApps.ReadWrite.All' -and $TenantId -eq 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e' -and -not $CertificateThumbprint }
    }

    It 'uses the certificate from the tenant profile for application authentication' {
        Mock -ModuleName $script:ModuleName Assert-IpGraphModule { }
        Mock -ModuleName $script:ModuleName Disconnect-MgGraph { }
        Mock -ModuleName $script:ModuleName Connect-MgGraph { }
        Mock -ModuleName $script:ModuleName Get-MgContext { [pscustomobject]@{ TenantId = '6e1f3a9c-2b8d-4c7e-a5f0-9d3b7c1e8a24'; Account = ''; ClientId = '3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c'; AuthType = 'AppOnly'; Scopes = @('DeviceManagementApps.ReadWrite.All') } }
        $profile = [pscustomobject]@{ Name = 'cert'; TenantId = '6e1f3a9c-2b8d-4c7e-a5f0-9d3b7c1e8a24'; AuthMode = 'Certificate'; ClientId = '3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c'; CertificateThumbprint = 'AB12CD34EF567890AB12CD34EF567890AB12CD34' }
        (Connect-IpGraph -TenantProfile $profile).Identity | Should -Be 'application 3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c'
        Should -Invoke -ModuleName $script:ModuleName Connect-MgGraph -Times 1 -ParameterFilter { $ClientId -eq '3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c' -and $CertificateThumbprint -eq 'AB12CD34EF567890AB12CD34EF567890AB12CD34' -and -not $Scopes }
    }

    It 'resolves the effective sign-in mode' {
        Resolve-IpSignInMode -TenantProfile ([pscustomobject]@{ AuthMode = 'Delegated'; DelegatedSignIn = 'Auto'; ClientId = '' }) | Should -Be 'Wam'
        Resolve-IpSignInMode -TenantProfile ([pscustomobject]@{ AuthMode = 'Delegated'; DelegatedSignIn = 'Auto'; ClientId = '3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c' }) | Should -Be 'Browser'
        Resolve-IpSignInMode -TenantProfile ([pscustomobject]@{ AuthMode = 'Delegated'; DelegatedSignIn = 'Wam'; ClientId = '' }) -UseDeviceCode | Should -Be 'DeviceCode'
        Resolve-IpSignInMode -TenantProfile ([pscustomobject]@{ AuthMode = 'Certificate'; ClientId = 'x' }) | Should -Be 'Certificate'
        Resolve-IpSignInMode -TenantProfile ([pscustomobject]@{ AuthMode = 'Delegated' }) | Should -Be 'Wam'
    }

    It 'disables WAM for a browser sign-in and restores the option on disconnect' {
        $script:Options = [Collections.Generic.List[bool]]::new()
        Mock -ModuleName $script:ModuleName Assert-IpGraphModule { }
        Mock -ModuleName $script:ModuleName Disconnect-MgGraph { }
        Mock -ModuleName $script:ModuleName Connect-MgGraph { }
        Mock -ModuleName $script:ModuleName Get-MgGraphOption { [pscustomobject]@{ DisableLoginByWAM = $false } }
        Mock -ModuleName $script:ModuleName Set-MgGraphOption { $script:Options.Add($DisableLoginByWAM) }
        Mock -ModuleName $script:ModuleName Get-MgContext { [pscustomobject]@{ TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; Account = 'admin@contoso.example'; ClientId = '3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c'; AuthType = 'Delegated'; Scopes = @('DeviceManagementApps.ReadWrite.All') } }
        $profile = [pscustomobject]@{ Name = 'prod'; TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; AuthMode = 'Delegated'; DelegatedSignIn = 'Browser'; ClientId = '3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c'; CertificateThumbprint = '' }
        Connect-IpGraph -TenantProfile $profile | Out-Null
        Disconnect-IpGraph
        $script:Options | Should -Be @($true, $false)
        Should -Invoke -ModuleName $script:ModuleName Connect-MgGraph -Times 1 -ParameterFilter { $ClientId -eq '3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c' -and -not $UseDeviceCode }
    }

    It 'keeps the SDK option unchanged when WAM is already disabled' {
        Mock -ModuleName $script:ModuleName Assert-IpGraphModule { }
        Mock -ModuleName $script:ModuleName Disconnect-MgGraph { }
        Mock -ModuleName $script:ModuleName Connect-MgGraph { }
        Mock -ModuleName $script:ModuleName Get-MgGraphOption { [pscustomobject]@{ DisableLoginByWAM = $true } }
        Mock -ModuleName $script:ModuleName Set-MgGraphOption { }
        Mock -ModuleName $script:ModuleName Get-MgContext { [pscustomobject]@{ TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; Account = 'a'; ClientId = ''; AuthType = 'Delegated'; Scopes = @() } }
        $profile = [pscustomobject]@{ Name = 'prod'; TenantId = ''; AuthMode = 'Delegated'; DelegatedSignIn = 'Browser'; ClientId = '3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c'; CertificateThumbprint = '' }
        Connect-IpGraph -TenantProfile $profile | Out-Null
        Disconnect-IpGraph
        Should -Invoke -ModuleName $script:ModuleName Set-MgGraphOption -Times 0
    }

    It 'accepts a domain TenantId without a strict GUID comparison' {
        Mock -ModuleName $script:ModuleName Assert-IpGraphModule { }
        Mock -ModuleName $script:ModuleName Disconnect-MgGraph { }
        Mock -ModuleName $script:ModuleName Connect-MgGraph { }
        Mock -ModuleName $script:ModuleName Get-MgContext { [pscustomobject]@{ TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; Account = 'a'; ClientId = ''; AuthType = 'Delegated'; Scopes = @() } }
        $profile = [pscustomobject]@{ Name = 'prod'; TenantId = 'northwind-it.onmicrosoft.com'; AuthMode = 'Delegated'; ClientId = ''; CertificateThumbprint = '' }
        (Connect-IpGraph -TenantProfile $profile).TenantId | Should -Be 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'
        Should -Invoke -ModuleName $script:ModuleName Connect-MgGraph -Times 1 -ParameterFilter { $TenantId -eq 'northwind-it.onmicrosoft.com' }
    }

    It 'reports whether the WAM broker runtime is loaded' {
        Test-IpGraphBrokerLoaded | Should -BeOfType [bool]
        (Get-IpGraphRuntimeInfo).PowerShell | Should -Be $PSVersionTable.PSVersion.ToString()
    }

    It 'tells where to find the WAM sign-in window' {
        Mock -ModuleName $script:ModuleName Assert-IpGraphModule { }
        Mock -ModuleName $script:ModuleName Disconnect-MgGraph { }
        Mock -ModuleName $script:ModuleName Connect-MgGraph { }
        Mock -ModuleName $script:ModuleName Get-MgContext { [pscustomobject]@{ TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; Account = 'a'; ClientId = ''; AuthType = 'Delegated'; Scopes = @('DeviceManagementApps.Read.All') } }
        $script:Logged = [Collections.Generic.List[string]]::new()
        Set-IpLogSink -Sink { param($line, $level) $script:Logged.Add($line) } -Console $false
        try { Connect-IpGraph -TenantProfile ([pscustomobject]@{ Name = 'prod'; TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; AuthMode = 'Delegated'; ClientId = ''; CertificateThumbprint = '' }) -Scopes @('DeviceManagementApps.Read.All') | Out-Null }
        finally { Set-IpLogSink -Sink $null -Console $false }
        ($script:Logged -join "`n") | Should -Match "If you don't see the sign-in window, check the taskbar or press Alt\+Tab\."
        ($script:Logged -join "`n") | Should -Not -Match 'Granted permissions do not include'
        Should -Invoke -ModuleName $script:ModuleName Connect-MgGraph -Times 1 -ParameterFilter { ($Scopes -join ',') -eq 'DeviceManagementApps.Read.All' }
    }
    It 'reports a canceled sign-in without the ErrorActionPreference wrapper' {
        Mock -ModuleName $script:ModuleName Assert-IpGraphModule { }
        Mock -ModuleName $script:ModuleName Disconnect-MgGraph { }
        Mock -ModuleName $script:ModuleName Connect-MgGraph { throw [InvalidOperationException]::new('InteractiveBrowserCredential authentication failed: User canceled authentication.') }
        $profile = [pscustomobject]@{ Name = 'test'; TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; AuthMode = 'Delegated'; ClientId = ''; CertificateThumbprint = '' }
        { Connect-IpGraph -TenantProfile $profile } | Should -Throw "``[Connect``] Sign-in to tenant profile 'test' failed or was canceled: InteractiveBrowserCredential authentication failed: User canceled authentication."
    }

    It 'refuses a session connected to another tenant' {
        Mock -ModuleName $script:ModuleName Assert-IpGraphModule { }
        Mock -ModuleName $script:ModuleName Disconnect-MgGraph { }
        Mock -ModuleName $script:ModuleName Connect-MgGraph { }
        Mock -ModuleName $script:ModuleName Get-MgContext { [pscustomobject]@{ TenantId = '6e1f3a9c-2b8d-4c7e-a5f0-9d3b7c1e8a24'; Account = 'a'; ClientId = ''; AuthType = 'Delegated'; Scopes = @() } }
        $profile = [pscustomobject]@{ Name = 'test'; TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; AuthMode = 'Delegated'; ClientId = ''; CertificateThumbprint = '' }
        { Connect-IpGraph -TenantProfile $profile } | Should -Throw '*does not match the tenant profile*'
    }
}

Describe 'Graph error handling and throttling' {
    It 'reads the status code and Retry-After from the response' {
        $info = Get-IpHttpErrorInfo -ErrorRecord (New-HttpException -StatusCode 429 -RetryAfterSeconds 7)
        $info.StatusCode | Should -Be 429
        $info.RetryAfterSeconds | Should -Be 7
    }

    It 'extracts the Graph error from the raw SDK response text' {
        $raw = @(
            'GET https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$filter=x'
            'HTTP/1.1 400 Bad Request'
            'Transfer-Encoding: chunked'
            'request-id: 1b2c3d4e-5f60-4a7b-8c9d-0e1f2a3b4c5d'
            'Content-Type: application/json'
            ''
            '{"error":{"code":"BadRequest","message":"Parsing OData Select and Expand failed: Could not find a property named ''displayVersion'' on type ''microsoft.graph.mobileApp''.","innerError":{"date":"2026-10-18T03:26:42","request-id":"1b2c3d4e-5f60-4a7b-8c9d-0e1f2a3b4c5d"}}}'
        ) -join "`r`n"
        $info = Get-IpHttpErrorInfo -ErrorRecord ([Exception]::new($raw))
        $info.StatusCode | Should -Be 400
        $info.RequestId | Should -Be '1b2c3d4e-5f60-4a7b-8c9d-0e1f2a3b4c5d'
        $info.Message | Should -Be "Parsing OData Select and Expand failed: Could not find a property named 'displayVersion' on type 'microsoft.graph.mobileApp'."
        $script:RawError = $raw
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest { throw ([Exception]::new($script:RawError)) }
        try { Invoke-IpGraphRequest -Method GET -Uri 'deviceAppManagement/mobileApps' -Step 'FindApp'; throw 'No error was raised.' }
        catch {
            $_.Exception.Message | Should -Match '^\[FindApp\] GET /beta/deviceAppManagement/mobileApps failed \(HTTP 400\): Parsing OData Select'
            $_.Exception.Message | Should -Not -Match 'Transfer-Encoding|\r|\n'
        }
    }

    It 'falls back to the status name in the message' {
        (Get-IpHttpErrorInfo -ErrorRecord ([Exception]::new('Response status code does not indicate success: ServiceUnavailable (Service Unavailable).'))).StatusCode | Should -Be 503
    }

    It 'retries throttled requests and honors Retry-After' {
        $script:Attempts = 0
        Mock -ModuleName $script:ModuleName Start-IpSleep { }
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest {
            $script:Attempts++
            if ($script:Attempts -lt 3) { throw (New-HttpException -StatusCode 429 -RetryAfterSeconds 7) }
            [pscustomobject]@{ id = 'ok' }
        }
        (Invoke-IpGraphRequest -Method GET -Uri 'deviceAppManagement/mobileApps' -Step 'FindApp').id | Should -Be 'ok'
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 3 -Exactly
        Should -Invoke -ModuleName $script:ModuleName Start-IpSleep -Times 2 -Exactly -ParameterFilter { $Seconds -eq 7 }
    }

    It 'uses exponential backoff without Retry-After' {
        Get-IpRetryDelaySeconds -Attempt 1 -RetryAfterSeconds $null | Should -Be 2
        Get-IpRetryDelaySeconds -Attempt 3 -RetryAfterSeconds $null | Should -Be 8
        Get-IpRetryDelaySeconds -Attempt 10 -RetryAfterSeconds $null | Should -Be 60
    }

    It 'stops after the retry budget' {
        Mock -ModuleName $script:ModuleName Start-IpSleep { }
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest { throw (New-HttpException -StatusCode 503) }
        { Invoke-IpGraphRequest -Method GET -Uri 'deviceAppManagement/mobileApps' -Step 'FindApp' } | Should -Throw '*HTTP 503*'
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 4 -Exactly
    }

    It 'does not retry client errors and names the failed step' {
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest { throw (New-HttpException -StatusCode 400) }
        { Invoke-IpGraphRequest -Method POST -Uri 'deviceAppManagement/mobileApps' -Body @{} -Step 'CreateApp' } | Should -Throw '`[CreateApp`] POST /beta/deviceAppManagement/mobileApps failed (HTTP 400)*'
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 1 -Exactly
    }
}

Describe 'Win32 app search' {
    It 'selects only base-type properties with the Win32 type filter' {
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest {
            [pscustomobject]@{ value = @(
                    [pscustomobject]@{ id = 'b1'; displayName = 'Contoso Sample App'; publisher = 'Contoso Ltd'; lastModifiedDateTime = '2026-10-17T16:42:10Z' }
                    [pscustomobject]@{ id = 'b2'; displayName = 'CONTOSO Field Tools' }
                    [pscustomobject]@{ id = 'b4'; displayName = 'Fabrikam Viewer'; publisher = 'Fabrikam' }
                ) }
        }
        $apps = @(Find-IpWin32App -DisplayName 'contoso' -Contains)
        @($apps.id) | Should -Be @('b1', 'b2')
        $apps[1].publisher | Should -Be ''
        $apps[0].PSObject.Properties.Name | Should -Be @('id', 'displayName', 'publisher', 'lastModifiedDateTime')
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 1 -Exactly -ParameterFilter {
            $decoded = [Uri]::UnescapeDataString($Uri)
            $decoded -match "\`$filter=isof\('microsoft\.graph\.win32LobApp'\)" -and $decoded -match '\$select=id,displayName,publisher,lastModifiedDateTime$' -and $decoded -notmatch 'displayVersion|committedContentVersion|largeIcon'
        }
    }

    It 'matches an exact display name locally' {
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest { [pscustomobject]@{ value = @([pscustomobject]@{ id = 'c1'; displayName = "Contoso O'Sample" }, [pscustomobject]@{ id = 'c2'; displayName = "Contoso O'Sample 2" }) } }
        @((Find-IpWin32App -DisplayName "contoso o'sample").id) | Should -Be @('c1')
    }

    It 'follows @odata.nextLink until the last page' {
        $script:NextLink = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')&`$select=id,displayName,publisher,lastModifiedDateTime&`$skiptoken=page2"
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest {
            if ($Uri -match 'skiptoken=page2') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'p2'; displayName = 'Contoso Two' }) } }
            [pscustomobject]@{ value = @([pscustomobject]@{ id = 'p1'; displayName = 'Contoso One' }); '@odata.nextLink' = $script:NextLink }
        }
        @((Find-IpWin32App).id) | Should -Be @('p1', 'p2')
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 2 -Exactly
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 1 -Exactly -ParameterFilter { $Uri -ceq $script:NextLink }
    }

    It 'stops when Graph repeats the same next link' {
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest { [pscustomobject]@{ value = @(); '@odata.nextLink' = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$skiptoken=loop' } }
        { Find-IpWin32App } | Should -Throw '*same @odata.nextLink twice*'
    }

    It 'reads the Win32 details of one app by id' {
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest {
            [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = $script:AppId; displayName = 'Contoso Sample App'; displayVersion = '2.4.0'; committedContentVersion = '2'; setupFilePath = 'ContosoSampleApp.msi'; publishingState = 'published'; isAssigned = $false }
        }
        $details = Get-IpWin32AppDetails -AppId $script:AppId
        $details.displayVersion | Should -Be '2.4.0'
        $details.committedContentVersion | Should -Be '2'
        $details.isAssigned | Should -Be 'False'
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'GET' -and $Uri -match "/mobileApps/$($script:AppId)$" }
    }

    It 'refuses details for an app that is not a Win32 app' {
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest { [pscustomobject]@{ '@odata.type' = '#microsoft.graph.webApp'; id = $script:AppId } }
        { Get-IpWin32AppDetails -AppId $script:AppId } | Should -Throw '*is not a Win32 app*'
    }
}

Describe 'Azure Storage helpers' {
    It 'parses the SAS expiry' {
        Get-IpSasExpiryUtc -Uri 'https://contoso.blob.core.windows.net/c/f?sv=2020&se=2026-10-18T10%3A30%3A00Z&sig=x' | Should -Be ([datetime]::new(2026, 10, 18, 10, 30, 0, [DateTimeKind]::Utc))
        Get-IpSasExpiryUtc -Uri 'https://contoso.blob.core.windows.net/c/f?sig=x' | Should -Be ([datetime]::MinValue)
    }

    It 'decides when the SAS must be renewed' {
        $now = [datetime]::new(2026, 10, 18, 10, 0, 0, [DateTimeKind]::Utc)
        Test-IpSasRenewalNeeded -ExpiryUtc $now.AddMinutes(4) -NowUtc $now -RenewBeforeMinutes 5 | Should -BeTrue
        Test-IpSasRenewalNeeded -ExpiryUtc $now.AddMinutes(30) -NowUtc $now -RenewBeforeMinutes 5 | Should -BeFalse
        Test-IpSasRenewalNeeded -ExpiryUtc ([datetime]::MinValue) -NowUtc $now | Should -BeFalse
    }

    It 'builds block ids and the block list' {
        New-IpBlockId -Index 0 | Should -Be ([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes('block-00000000')))
        New-IpBlockListXml -BlockIds @('a', 'b') | Should -Be '<?xml version="1.0" encoding="utf-8"?><BlockList><Latest>a</Latest><Latest>b</Latest></BlockList>'
        Join-IpSasQuery -Uri 'https://x/y?sig=1' -Query 'comp=block' | Should -Be 'https://x/y?sig=1&comp=block'
    }

    It 'retries transient Azure failures and reports 403 for SAS renewal' {
        $client = [Net.Http.HttpClient]::new()
        try {
            $script:AzureResponses = [Collections.Generic.Queue[int]]::new([int[]]@(503, 201))
            Mock -ModuleName $script:ModuleName Start-IpSleep { }
            Mock -ModuleName $script:ModuleName Invoke-IpAzureBlobRequest { [pscustomobject]@{ StatusCode = $script:AzureResponses.Dequeue(); Body = '' } }
            Send-IpAzureBlobWithRetry -Client $client -Uri 'https://x/y' -Bytes ([byte[]](1, 2)) -Operation 'block 0' -MaxRetries 3 | Should -Be 'Ok'
            Mock -ModuleName $script:ModuleName Invoke-IpAzureBlobRequest { [pscustomobject]@{ StatusCode = 403; Body = 'AuthenticationFailed' } }
            Send-IpAzureBlobWithRetry -Client $client -Uri 'https://x/y' -Bytes ([byte[]](1, 2)) -Operation 'block 0' -MaxRetries 3 | Should -Be 'AuthFailed'
            Mock -ModuleName $script:ModuleName Invoke-IpAzureBlobRequest { [pscustomobject]@{ StatusCode = 400; Body = 'InvalidBlockList' } }
            { Send-IpAzureBlobWithRetry -Client $client -Uri 'https://x/y' -Bytes ([byte[]](1, 2)) -Operation 'block 0' -MaxRetries 3 } | Should -Throw '*HTTP 400*'
        }
        finally { $client.Dispose() }
    }

    It 'uploads every block, renews the SAS after a 403 and commits the block list' {
        $path = New-TestIntuneWin -Path (Join-Path $TestDrive 'upload\Sample.intunewin') -XmlText (Get-TestDetectionXml) -PayloadSize ([int](2.5 * 1MB))
        $script:Uris = [Collections.Generic.List[string]]::new()
        $script:Rejected = $false
        $script:Progress = [Collections.Generic.List[int64]]::new()
        Mock -ModuleName $script:ModuleName Request-IpSasRenewal { 'https://contoso.blob.core.windows.net/c/f?sig=renewed' }
        Mock -ModuleName $script:ModuleName Invoke-IpAzureBlobRequest {
            $script:Uris.Add($Uri)
            if (-not $script:Rejected) { $script:Rejected = $true; return [pscustomobject]@{ StatusCode = 403; Body = 'AuthenticationFailed' } }
            [pscustomobject]@{ StatusCode = 201; Body = '' }
        }
        $blocks = Send-IpIntuneWinContent -IntuneWinPath $path -ContentEntryName 'IntuneWinPackage/Contents/IntunePackage.intunewin' -AzureStorageUri 'https://contoso.blob.core.windows.net/c/f?sig=first' -FileUri 'files/1' -BlockSizeMB 1 -ProgressCallback { param($sent, $total) $script:Progress.Add($sent) }
        $blocks | Should -Be 3
        Should -Invoke -ModuleName $script:ModuleName Request-IpSasRenewal -Times 1 -Exactly
        $script:Uris.Count | Should -Be 5
        $script:Uris[0] | Should -Match 'sig=first&comp=block'
        $script:Uris[1] | Should -Match 'sig=renewed&comp=block'
        $script:Uris[4] | Should -Match 'sig=renewed&comp=blocklist$'
        $script:Progress[-1] | Should -Be ([int64](2.5 * 1MB))
    }

    It 'forces one SAS renewal at the requested block (validation option)' {
        $path = New-TestIntuneWin -Path (Join-Path $TestDrive 'force\Sample.intunewin') -XmlText (Get-TestDetectionXml) -PayloadSize ([int](3.5 * 1MB))
        $script:Uris = [Collections.Generic.List[string]]::new()
        Mock -ModuleName $script:ModuleName Request-IpSasRenewal { 'https://contoso.blob.core.windows.net/c/f?sig=forced' }
        Mock -ModuleName $script:ModuleName Invoke-IpAzureBlobRequest { $script:Uris.Add($Uri); [pscustomobject]@{ StatusCode = 201; Body = '' } }
        $script:Logged = [Collections.Generic.List[string]]::new()
        Set-IpLogSink -Sink { param($line, $level) $script:Logged.Add($line) } -Console $false
        try {
            Send-IpIntuneWinContent -IntuneWinPath $path -ContentEntryName 'IntuneWinPackage/Contents/IntunePackage.intunewin' -AzureStorageUri 'https://contoso.blob.core.windows.net/c/f?sig=first' -FileUri 'files/1' -BlockSizeMB 1 -ForceSasRenewalAtBlock 3 | Should -Be 4
        }
        finally { Set-IpLogSink -Sink $null -Console $false }
        Should -Invoke -ModuleName $script:ModuleName Request-IpSasRenewal -Times 1 -Exactly
        $script:Uris[1] | Should -Match 'sig=first'
        $script:Uris[2] | Should -Match 'sig=forced'
        $script:Uris[-1] | Should -Match 'sig=forced&comp=blocklist$'
        ($script:Logged -join "`n") | Should -Match 'forcing a SAS renewal after block 2/4, before block 3/4'
    }

    It 'logs the renewed SAS URI before the upload resumes' {
        $script:Logged = [Collections.Generic.List[string]]::new()
        Mock -ModuleName $script:ModuleName Start-IpSleep { }
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest {
            if ($Method -eq 'POST') { return $null }
            [pscustomobject]@{ uploadState = 'azureStorageUriRenewalSuccess'; azureStorageUri = 'https://contoso.blob.core.windows.net/c/f?sig=renewed' }
        }
        Set-IpLogSink -Sink { param($line, $level) $script:Logged.Add($line) } -Console $false
        try { Request-IpSasRenewal -FileUri 'deviceAppManagement/mobileApps/x/files/1' | Should -Be 'https://contoso.blob.core.windows.net/c/f?sig=renewed' }
        finally { Set-IpLogSink -Sink $null -Console $false }
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Uri -match '/renewUpload$' }
        ($script:Logged -join "`n") | Should -Match '\[SasRenewal\] Renewed SAS URI received \(state azureStorageUriRenewalSuccess\)\. Resuming the upload\.'
    }

    It 'stops the upload when canceled' {
        $path = New-TestIntuneWin -Path (Join-Path $TestDrive 'cancel\Sample.intunewin') -XmlText (Get-TestDetectionXml) -PayloadSize 2048
        Mock -ModuleName $script:ModuleName Invoke-IpAzureBlobRequest { [pscustomobject]@{ StatusCode = 201; Body = '' } }
        { Send-IpIntuneWinContent -IntuneWinPath $path -ContentEntryName 'IntuneWinPackage/Contents/IntunePackage.intunewin' -AzureStorageUri 'https://x/y?sig=1' -FileUri 'files/1' -BlockSizeMB 1 -CancelCheck { $true } } | Should -Throw '*canceled*'
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpAzureBlobRequest -Times 0
    }
}

Describe 'Publication' {
    BeforeAll {
        $script:PackagePath = New-TestIntuneWin -Path (Join-Path $TestDrive 'publish\ContosoSetup.intunewin') -XmlText (Get-TestDetectionXml -WithMsi) -PayloadSize 4096

        function Initialize-GraphFake {
            param([string]$ExistingType = '#microsoft.graph.win32LobApp', [int]$CommitFailures = 0, [string]$FailOn = '')
            $script:Graph = @{ Requests = [Collections.Generic.List[object]]::new(); Committed = $false; CommitFailures = $CommitFailures; ExistingType = $ExistingType; FailOn = $FailOn }
            Mock -ModuleName $script:ModuleName Start-IpSleep { }
            Mock -ModuleName $script:ModuleName Get-IpGraphContext { [pscustomobject]@{ TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'; Account = 'admin@contoso.example' } }
            Mock -ModuleName $script:ModuleName Send-IpIntuneWinContent { 1 }
            Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest {
                $graph = $script:Graph
                $graph.Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body })
                if ($graph.FailOn -and $Uri -match $graph.FailOn) { throw (New-HttpException -StatusCode 500) }
                switch -Regex ("$Method $Uri") {
                    '^POST .*/mobileApps$' { return [pscustomobject]@{ id = $script:AppId; '@odata.type' = '#microsoft.graph.win32LobApp' } }
                    '^GET .*/mobileApps/[0-9a-f-]+$' { return [pscustomobject]@{ id = $script:AppId; displayName = 'Existing'; '@odata.type' = $graph.ExistingType; committedContentVersion = '1' } }
                    '^POST .*/contentVersions$' { return [pscustomobject]@{ id = '2' } }
                    '^POST .*/files$' { return [pscustomobject]@{ id = 'file-1' } }
                    '^POST .*/commit$' { $graph.Committed = $true; return $null }
                    '^GET .*/files/file-1$' {
                        if (-not $graph.Committed) { return [pscustomobject]@{ uploadState = 'azureStorageUriRequestSuccess'; azureStorageUri = 'https://contoso.blob.core.windows.net/c/f?sig=1' } }
                        if ($graph.CommitFailures -gt 0) { $graph.CommitFailures--; $graph.Committed = $false; return [pscustomobject]@{ uploadState = 'commitFileFailed'; azureStorageUri = $null } }
                        return [pscustomobject]@{ uploadState = 'commitFileSuccess'; azureStorageUri = $null }
                    }
                    '^PATCH .*/mobileApps/[0-9a-f-]+$' { return $null }
                    default { throw "Unexpected request: $Method $Uri" }
                }
            }
        }
    }

    It 'returns the planned requests in dry run without calling Graph' {
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest { throw 'Graph must not be called.' }
        Mock -ModuleName $script:ModuleName Get-IpGraphContext { throw 'Graph context must not be read.' }
        $folder = Join-Path $TestDrive 'dryrun'
        $result = Publish-IpWin32App -Definition (New-TestDefinition) -IntuneWinPath $script:PackagePath -DryRun -DryRunOutputFolder $folder
        $result.Mode | Should -Be 'DryRun'
        $result.AppId | Should -BeNullOrEmpty
        $result.Assigned | Should -BeFalse
        @($result.Plan.Step) | Should -Be @('CreateApp', 'ContentVersion', 'ContentFile', 'SasWait', 'Upload', 'Commit', 'CommitWait', 'Finalize', 'Verify')
        @(Get-ChildItem -LiteralPath $folder -Filter '*.json').Count | Should -Be 9
        ($result.Plan | Where-Object Step -EQ 'Commit').Body.fileEncryptionInfo.encryptionKey | Should -Be '<redacted>'
        ($result.Plan | Where-Object Step -EQ 'CreateApp').Body.allowedArchitectures | Should -Be 'x64,arm64'
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 0
    }

    It 'treats -WhatIf like a dry run' {
        Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest { throw 'Graph must not be called.' }
        (Publish-IpWin32App -Definition (New-TestDefinition) -IntuneWinPath $script:PackagePath -WhatIf).Mode | Should -Be 'DryRun'
        Should -Invoke -ModuleName $script:ModuleName Invoke-IpGraphRawRequest -Times 0
    }

    It 'plans a lookup instead of a creation in update mode' {
        $definition = New-TestDefinition
        $definition.Publish.Mode = 'UpdateExisting'
        $definition.Publish.ExistingAppId = $script:AppId
        $result = Publish-IpWin32App -Definition $definition -IntuneWinPath $script:PackagePath -DryRun
        $result.Plan[0].Step | Should -Be 'GetApp'
        $result.Plan[0].Uri | Should -Match $script:AppId
    }

    It 'creates, uploads, commits and finalizes a new app without assigning it' {
        Initialize-GraphFake
        $result = Publish-IpWin32App -Definition (New-TestDefinition) -IntuneWinPath $script:PackagePath -Confirm:$false
        $result.Mode | Should -Be 'Publish'
        $result.AppId | Should -Be $script:AppId
        $result.ContentVersionId | Should -Be '2'
        $result.Assigned | Should -BeFalse
        $result.PortalUrl | Should -Be ('https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/' + $script:AppId)
        @($script:Graph.Requests | Where-Object { $_.Uri -match 'assignments|groups' }).Count | Should -Be 0
        $patch = $script:Graph.Requests | Where-Object Method -EQ 'PATCH' | Select-Object -Last 1
        $patchBody = $patch.Body | ConvertFrom-Json
        $patchBody.committedContentVersion | Should -Be '2'
        $patchBody.PSObject.Properties.Name | Should -Not -Contain 'applicableArchitectures'
        Should -Invoke -ModuleName $script:ModuleName Send-IpIntuneWinContent -Times 1 -Exactly
    }

    It 'adds a content version to an existing app in update mode' {
        Initialize-GraphFake
        $definition = New-TestDefinition
        $definition.Publish.Mode = 'UpdateExisting'
        $definition.Publish.ExistingAppId = $script:AppId
        $result = Publish-IpWin32App -Definition $definition -IntuneWinPath $script:PackagePath -Confirm:$false
        $result.UpdatedExistingApp | Should -BeTrue
        @($script:Graph.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/mobileApps$' }).Count | Should -Be 0
    }

    It 'refuses to update an app that is not a Win32 app' {
        Initialize-GraphFake -ExistingType '#microsoft.graph.windowsMobileMSI'
        $definition = New-TestDefinition
        $definition.Publish.Mode = 'UpdateExisting'
        $definition.Publish.ExistingAppId = $script:AppId
        { Publish-IpWin32App -Definition $definition -IntuneWinPath $script:PackagePath -Confirm:$false } | Should -Throw '*step GetApp*is not a Win32 app*'
    }

    It 'retries a failed commit once' {
        Initialize-GraphFake -CommitFailures 1
        (Publish-IpWin32App -Definition (New-TestDefinition) -IntuneWinPath $script:PackagePath -Confirm:$false).AppId | Should -Be $script:AppId
        @($script:Graph.Requests | Where-Object { $_.Uri -match '/commit$' }).Count | Should -Be 2
    }

    It 'reports the failed step and the incomplete app id' {
        Initialize-GraphFake -FailOn '/contentVersions$'
        $script:Logged = [Collections.Generic.List[string]]::new()
        Set-IpLogSink -Sink { param($line, $level) $script:Logged.Add($line) } -Console $false
        try {
            { Publish-IpWin32App -Definition (New-TestDefinition) -IntuneWinPath $script:PackagePath -Confirm:$false } | Should -Throw '*Publication failed at step ContentVersion*HTTP 500*'
        }
        finally { Set-IpLogSink -Sink $null -Console $false }
        ($script:Logged -join "`n") | Should -Match ([regex]::Escape("The app $($script:AppId) was created but its content is not committed"))
    }

    It 'requires a Graph connection outside dry run' {
        Mock -ModuleName $script:ModuleName Get-IpGraphContext { $null }
        { Publish-IpWin32App -Definition (New-TestDefinition) -IntuneWinPath $script:PackagePath -Confirm:$false } | Should -Throw '*Not connected to Microsoft Graph*'
    }

    It 'logs the portal link and the unassigned status on success' {
        Initialize-GraphFake
        $script:Logged = [Collections.Generic.List[string]]::new()
        Set-IpLogSink -Sink { param($line, $level) $script:Logged.Add($line) } -Console $false
        try { Publish-IpWin32App -Definition (New-TestDefinition) -IntuneWinPath $script:PackagePath -Confirm:$false | Out-Null }
        finally { Set-IpLogSink -Sink $null -Console $false }
        $text = $script:Logged -join "`n"
        $text | Should -Match 'Intune admin center: https://intune\.microsoft\.com/'
        $text | Should -Match 'The app is not assigned\. Create assignments in the Intune admin center\.'
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCzLLjaPx2q4oXu
# gkE1prH7FTwLpuEu8qEqo35RhuGtb6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPHWwRy+oUFlETs2l9c59R02VI5zAdMK5QjYAUTG6wYdMA0GCSqG
# SIb3DQEBAQUABIIBgE+GxAeqKscF3IMYPSr4ENnkiQNadEbjewmv48/KD7utX1/f
# 7yCx9Nr4fkdYFgm4AaC2S/XmO6socjp8k2VxK8FiyyNrI47vMbN2R0A6wjpFR1SB
# 1sn11xh3hdi8dy8Cz0bbzEXAqv7aPwuKHDTpYHKyxQMEvD9Tf8rYTe3wkQhhsygz
# veBKY8A69sua5NOuDoAivpoDsaMTKkMwAz06XCukp0h60L0fVPVyfaAhNOFyjfNi
# BDMY6FqCcUEYMq91PXsSsaSBLCZjghsFyfsWBfsgePFsFiD6V4daPTaOr/2jG8dw
# 9H6S94psiJctwlbsvThhBQbogrCIs0yx3wFrK1Po/17c14emOA17at7AY/byf0Xf
# wbdHswYvAr3Ck0+sQU+ZHMesFYuOYKbspuEtaGmp9KsJIcgzRNZ2IOa5JrG5UYbn
# s4FKWVeTrXgLmwaGC8eBkgjIireHUgAJWkeqITtvFbcDTyB3oDAeridO1wZ216nL
# 9baiOv52afMZ+JRIr6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzdaMC8GCSqGSIb3DQEJBDEiBCAgYyKGuEMKv7fiBOe0OkfFWJ47gFHvw1zxi+jK
# Q2UQVDANBgkqhkiG9w0BAQEFAASCAgAeO2AOGEoe4lTJWlbfPgbf81GYVsPBNX7z
# 72FH7nkpEfBNUZmtcTFlUGKDD3+F+GytVsR+uqWPM+pM6beAzL2Bu8jsEM8lTl3r
# qnMEj1YxMM4jYSi6/5BeQWnLhciuCJBfH4evgNA8dUFroD8JsN5OfEQUJRDSKsxU
# QQxZJ9LpXh3MiCG67ehdVqsoDKeKJ0XoSB0uDrEHhe2gBTecyyF33slsb/RrxqM0
# mNPdvfcK+uI11YSHAB71KJAy37t5CgdxW9h38diMlwxvULJ7ZV+2WAQv3fLgwVoj
# DxIuAavG2B3MXTK4FUzkU0wE17d8mzlSVSkq3gBCuVX3qWGrmyhgutDVRDyGbA0z
# 17AK7EYvHY+/QILO+OE5qPZnNStp+S8HY6MsNBS14UVHLooVTQmb/+BPxCTSFX8l
# kpQfknoEb+c9yz700GATPbieVPAi52IqKekcP9QRN7i6jNb3BoHC9tFmOaZkaIvT
# QX0MQWuezLOJmK44hTgnuz/8THhx4OQHewvYtbX4EkH678Liqd8A57NiqOhmtf4V
# VEd6LMJGR+Favk+w1WpMpLQqMXBMVavrrOZ70oIBjc8ewsSJjUCAIIrMiwCki77A
# jRi6HC4jhEWrpzyUDObMpgUZyTLhgzOGmCTONXrITrtl07uUDP5NZv1X/Vymsyyi
# 4KmNvM8nDA==
# SIG # End signature block
