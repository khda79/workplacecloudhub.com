<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 26.06.06.1238

<#
.SYNOPSIS

.\FreeBusyChecker.ps1

.DESCRIPTION

This script can be used to validate the Availability configuration of the following Exchange Server Versions:

- Exchange Server
- Exchange Online

Required Permissions:

    - Organization Management
    - Domain Admin

Please make sure that the account used is a member of the Local Administrator group. This should be fulfilled on Exchange Servers by being a member of the Organization Management group. However, if the group membership was adjusted, or in case the script is executed on a non-Exchange system like a management Server, you need to add your account to the Local Administrator group.

How To Run:

This script must be run as Administrator in Exchange Management Shell on an Exchange Server. You can provide no parameters, and the script will just run against Exchange On-Premises and Exchange Online to query for OAuth and DAuth configuration settings. It will compare existing values with standard values and provide details of what may not be correct.
Please take note that though this script may output that a specific setting is not a standard setting, it does not mean that your configurations are incorrect. For example, DNS may be configured with specific mappings that this script cannot evaluate.

To collect information for Exchange Online a connection to Exchange Online must be established before running the script using Connection Prefix "EO".

Example:

PS C:\scripts\FreeBusyChecker> Connect-ExchangeOnline -Prefix EO

.PARAMETER Auth
Allows you to choose the authentication type to validate.
.PARAMETER Org
Allows you to choose the organization type to validate.
.PARAMETER OnPremUser
Specifies the Exchange On Premise User that will be used to test Free Busy Settings.
.PARAMETER OnlineUser
Specifies the Exchange Online User that will be used to test Free Busy Settings.
.PARAMETER OnPremDomain
Specifies the domain for on-premises Organization.
.PARAMETER OnPremEWSUrl
Specifies the EWS (Exchange Web Services) URL for on-premises Exchange Server.
.PARAMETER OnPremLocalDomain
Specifies the local AD domain for the on-premises Organization.
.PARAMETER Help
Show help for this script.

.EXAMPLE
.\FreeBusyChecker.ps1
This cmdlet will run the Free Busy Checker script and Check Availability OAuth and DAuth Configurations both for Exchange On-Premises and Exchange Online.
.EXAMPLE
.\FreeBusyChecker.ps1 -Auth OAuth
This cmdlet will run the Free Busy Checker Script against OAuth Availability Configurations.
.EXAMPLE
.\FreeBusyChecker.ps1 -Auth DAuth
This cmdlet will run the Free Busy Checker Script against DAuth Availability Configurations.
.EXAMPLE
.\FreeBusyChecker.ps1 -Org ExchangeOnline
This cmdlet will run the Free Busy Checker Script for Exchange Online Availability Configurations.
.EXAMPLE
.\FreeBusyChecker.ps1 -Org ExchangeOnPremise
This cmdlet will run the Free Busy Checker Script for Exchange On-Premises OAuth or DAuth Availability Configurations.
.EXAMPLE
.\FreeBusyChecker.ps1 -Org All
This cmdlet will run the Free Busy Checker Script for Exchange On-Premises and Exchange Online OAuth or DAuth Availability Configurations.
.EXAMPLE
.\FreeBusyChecker.ps1 -Org ExchangeOnPremise -Auth OAuth
This cmdlet will run the Free Busy Checker Script for Exchange On-Premises Availability OAuth Configurations
#>

# Exchange On-Premises
#>
#region Properties and Parameters

#Requires -Module ExchangeOnlineManagement
#Requires -Module ActiveDirectory

[CmdletBinding(DefaultParameterSetName = "FreeBusyInfo_OP", SupportsShouldProcess)]

param(
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [ValidateSet('DAuth', 'OAuth', 'All', '')]
    [string[]]$Auth,
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [ValidateSet('ExchangeOnPremise', 'ExchangeOnline')]
    [string[]]$Org,
    [Parameter(Mandatory = $true, ParameterSetName = "Help")]
    [switch]$Help,
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [string]$OnPremisesUser,
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [string]$OnlineUser,
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [string]$OnPremDomain,
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [string]$OnPremEWSUrl,
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [string]$OnPremLocalDomain,
    [Parameter(Mandatory = $true, ParameterSetName = "ScriptUpdateOnly", HelpMessage = "Update only script.")]
    [switch]$ScriptUpdateOnly,
    [Parameter(Mandatory = $false, ParameterSetName = "SkipVersionCheck", HelpMessage = "Skip version check.")]
    [switch]$SkipVersionCheck
)
begin {
function OrgRelCheck($Script:OrgRelParameter) {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-OrganizationRelationship  | Where{($_.DomainNames -like $Script:ExchangeOnlineDomain )} | Select Identity,DomainNames,FreeBusy*,TarGet*,Enabled, ArchiveAccessEnabled"
    PrintDynamicWidthLine
    $Script:OrgRelParameter
    $countOrgRelIssues = (0)
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Get-OrganizationRelationship"
    PrintDynamicWidthLine
    $settingsList = New-Object System.Collections.ArrayList
    function AddSettingToList($list, $name, $value, $color) {
        $list.Add([PSCustomObject]@{
                Name  = $name
                Value = $value
                Color = $color
            }) | Out-Null
    }
    # Domain Names
    if ($Script:OrgRelParameter.DomainNames -like $Script:ExchangeOnlineDomain) {
        AddSettingToList -list $settingsList -name "Domain Names" -value "Domain Names include the $Script:ExchangeOnlineDomain Domain" -color "green"
    } else {
        AddSettingToList -list $settingsList -name "Domain Names" -value "Domain Names do Not Include the $Script:ExchangeOnlineDomain Domain" -color "red"
    }
    # FreeBusyAccessEnabled
    if ($Script:OrgRelParameter.FreeBusyAccessEnabled -like "True") {
        AddSettingToList -list $settingsList -name "FreeBusyAccessEnabled" -value "FreeBusyAccessEnabled is set to True" -color "green"
    } else {
        AddSettingToList -list $settingsList -name "FreeBusyAccessEnabled" -value "FreeBusyAccessEnabled is set to False" -color "red"
        $countOrgRelIssues++
    }
    # TarGetOwAUrl
    $standardValues = @("http://outlook.com/owa/$($Script:ExchangeOnlineDomain).", "https://outlook.office.com/mail.")
    if ([string]::IsNullOrWhiteSpace($Script:OrgRelParameter.TarGetOwAUrl)) {
        AddSettingToList -list $settingsList -name "TarGetOwAUrl" -value "TarGetOwAUrl Is Blank. Can also be configured to be $($standardValues[0]) or $($standardValues[1])" -color "green"
    } elseif ($Script:OrgRelParameter.TarGetOwAUrl -in $standardValues) {
        AddSettingToList -list $settingsList -name "TarGetOwAUrl" -value "TarGetOwAUrl Is $($Script:OrgRelParameter.TarGetOwAUrl). This is a possible standard value." -color "green"
    } else {
        $countOrgRelIssues++
    }
    # TarGetSharingEpr
    if ([string]::IsNullOrWhitespace($Script:OrgRelParameter.TarGetSharingEpr) -or $Script:OrgRelParameter.TarGetSharingEpr -eq "https://outlook.office365.com/EWS/Exchange.asmx") {
        AddSettingToList -list $settingsList -name "TarGetSharingEpr" -value "TarGetSharingEpr Is ideally blank. If set, should be Office 365 EWS endpoint. Example: https://outlook.office365.com/EWS/Exchange.asmx" -color "green"
    } else {
        AddSettingToList -list $settingsList -name "TarGetSharingEpr" -value "TarGetSharingEpr Should be blank or https://outlook.office365.com/EWS/Exchange.asmx. If set, should be Office 365 EWS endpoint." -color "red"
        $countOrgRelIssues++
    }
    # FreeBusyAccessScope
    if ([string]::IsNullOrWhitespace($Script:OrgRelParameter.FreeBusyAccessScope)) {
        AddSettingToList -list $settingsList -name "FreeBusyAccessScope" -value "FreeBusyAccessScope Is blank, this is the standard Value." -color "green"
    } else {
        AddSettingToList -list $settingsList -name "FreeBusyAccessScope" -value "FreeBusyAccessScope Should be Blank, that is the standard Value." -color "red"
        $countOrgRelIssues++
    }
    # TarGetAutoDiscoverEpr
    $Script:OrgRelTarGetAutoDiscoverEpr = $Script:OrgRelParameter.TarGetAutoDiscoverEpr
    if ([string]::IsNullOrWhitespace($Script:OrgRelTarGetAutoDiscoverEpr)) {
        $Script:OrgRelTarGetAutoDiscoverEpr = "Blank"
    }
    if ($Script:OrgRelParameter.TarGetAutoDiscoverEpr -like "https://AutoDiscover-s.outlook.com/AutoDiscover/AutoDiscover.svc/WSSecurity") {
        AddSettingToList -list $settingsList -name "TarGetAutoDiscoverEpr" -value "TarGetAutoDiscoverEpr Is correct" -color "green"
    } else {
        AddSettingToList -list $settingsList -name "TarGetAutoDiscoverEpr" -value "TarGetAutoDiscoverEpr Is not correct. Should be https://AutoDiscover-s.outlook.com/AutoDiscover/AutoDiscover.svc/WSSecurity" -color "red"
        $countOrgRelIssues++
    }
    # Enabled
    if ($Script:OrgRelParameter.enabled -like "True") {
        AddSettingToList -list $settingsList -name "Enabled" -value "Enabled is set to True" -color "green"
    } else {
        AddSettingToList -list $settingsList -name "Enabled" -value "Enabled is set to False. This may be intentional if Hybrid Free Busy lookups are done with OAuth and Intra Organization Connector." -color "yellow"
        $countOrgRelIssues++
    }
    # Display the settings list
    if ($countOrgRelIssues -eq 0) {
        Write-Host -ForegroundColor Green " Configurations Seem Correct"
    } else {
        Write-Host -ForegroundColor Red "  Configurations may not be Correct"
    }
    foreach ($setting in $settingsList) {
        Write-Host -ForegroundColor White " $($setting.Name):"
        Write-Host -ForegroundColor $setting.Color " $($setting.Value)"
    }
    $Script:OrgRelDomainNames = ""
    foreach ($domain in $Script:OrgRelParameter.DomainNames.Domain) {
        if ($Script:OrgRelDomainNames -ne "") {
            $Script:OrgRelDomainNames += "; "
        }
        $Script:OrgRelDomainNames += $domain
    }
    orgRelHtml
    Write-Host -ForegroundColor Yellow "`n  Reference: https://learn.microsoft.com/en-us/exchange/create-an-organization-relationship-exchange-2013-help"
}
function FedInfoCheck {
    Write-Host -ForegroundColor Green " Get-FederationInformation -DomainName $Script:ExchangeOnlineDomain  -BypassAdditionalDomainValidation | fl"
    PrintDynamicWidthLine
    $FedInfo = Get-federationInformation -DomainName $Script:ExchangeOnlineDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue | Select-Object *
    if (!$FedInfo) {
        $FedInfo = Get-federationInformation -DomainName $Script:ExchangeOnlineDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue | Select-Object *
    }
    $FedInfo
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Federation Information"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White   "  Domain Names: "
    if ($FedInfo.DomainNames -like "*$Script:ExchangeOnlineDomain*") {
        Write-Host -ForegroundColor Green "   Domain Names include the Exchange Online Domain "$Script:ExchangeOnlineDomain
        $Script:tdDomainNamesColor = "green"
        $Script:tdDomainNamesFL = "Domain Names include the Exchange Online Domain $Script:ExchangeOnlineDomain"
    } else {
        Write-Host -ForegroundColor Red "   Domain Names seem not to include the Exchange Online Domain "$Script:ExchangeOnlineDomain
        Write-Host  "   Domain Names: "$FedInfo.DomainNames
        $Script:tdDomainNamesColor = "Red"
        $Script:tdDomainNamesFL = "Domain Names seem not to include the Exchange Online Domain: $Script:ExchangeOnlineDomain"
    }
    Write-Host  -ForegroundColor White  "  TokenIssuerUris: "
    if ($FedInfo.TokenIssuerUris -like "*urn:federation:MicrosoftOnline*") {
        Write-Host -ForegroundColor Green "  "  $FedInfo.TokenIssuerUris
        $Script:tdTokenIssuerUrisColor = "green"
        $Script:tdTokenIssuerUrisFL = $FedInfo.TokenIssuerUris
    } else {
        Write-Host "   " $FedInfo.TokenIssuerUris
        Write-Host  -ForegroundColor Red "   TokenIssuerUris should be urn:federation:MicrosoftOnline"
        $Script:tdTokenIssuerUrisColor = "red"
        $Script:tdTokenIssuerUrisFL = "   TokenIssuerUris should be urn:federation:MicrosoftOnline"
    }
    Write-Host -ForegroundColor White   "  TarGetApplicationUri:"
    if ($FedInfo.TarGetApplicationUri -like "Outlook.com") {
        Write-Host -ForegroundColor Green "  "$FedInfo.TarGetApplicationUri
        $Script:tdTarGetApplicationUriColor = "green"
        $Script:tdTarGetApplicationUriFL = $FedInfo.TarGetApplicationUri
    } else {
        Write-Host -ForegroundColor Red "   "$FedInfo.TarGetApplicationUri
        Write-Host -ForegroundColor Red   "   TarGetApplicationUri should be Outlook.com"
        $Script:tdTarGetApplicationUriColor = "red"
        $Script:tdTarGetApplicationUriFL = "   TarGetApplicationUri should be Outlook.com"
    }
    Write-Host -ForegroundColor White   "  TarGetAutoDiscoverEpr:"
    if ($FedInfo.TarGetAutoDiscoverEpr -like "https://AutoDiscover-s.outlook.com/AutoDiscover/AutoDiscover.svc/WSSecurity") {
        Write-Host -ForegroundColor Green "   "$FedInfo.TarGetAutoDiscoverEpr
        $Script:tdTarGetAutoDiscoverEprColor = "green"
        $Script:tdTarGetAutoDiscoverEprFL = $FedInfo.TarGetAutoDiscoverEpr
    } else {
        Write-Host -ForegroundColor Red "   "$FedInfo.TarGetAutoDiscoverEpr
        Write-Host -ForegroundColor Red   " TarGetAutoDiscoverEpr should be https://AutoDiscover-s.outlook.com/AutoDiscover/AutoDiscover.svc/WSSecurity"
        $Script:tdTarGetAutoDiscoverEprColor = "red"
        $Script:tdTarGetAutoDiscoverEprFL = "   TarGetAutoDiscoverEpr should be https://AutoDiscover-s.outlook.com/AutoDiscover/AutoDiscover.svc/WSSecurity"
    }
    Write-Host -ForegroundColor White "  Federation Information TarGetApplicationUri vs Organization Relationship TarGetApplicationUri "
    if ($FedInfo.TarGetApplicationUri -like "Outlook.com") {
        if ($Script:OrgRel.TarGetApplicationUri -like $FedInfo.TarGetApplicationUri) {
            Write-Host -ForegroundColor Green "   => Federation Information TarGetApplicationUri matches the Organization Relationship TarGetApplicationUri "
            Write-Host  "       Organization Relationship TarGetApplicationUri:"  $Script:OrgRel.TarGetApplicationUri
            Write-Host  "       Federation Information TarGetApplicationUri:   "  $FedInfo.TarGetApplicationUri
            $Script:tdFederationInformationTAColor = "green"
            $Script:tdFederationInformationTA_FL = " => Federation Information TarGetApplicationUri matches the Organization Relationship TarGetApplicationUri"
        } else {
            Write-Host -ForegroundColor Red "   => Federation Information TarGetApplicationUri should be Outlook.com and match the Organization Relationship TarGetApplicationUri "
            Write-Host  "       Organization Relationship TarGetApplicationUri:"  $Script:OrgRel.TarGetApplicationUri
            Write-Host  "       Federation Information TarGetApplicationUri:   "  $FedInfo.TarGetApplicationUri
            $Script:tdFederationInformationTAColor = "red"
            $Script:tdFederationInformationTA_FL = " => Federation Information TarGetApplicationUri should be Outlook.com and match the Organization Relationship TarGetApplicationUri"
        }
    }
    Write-Host -ForegroundColor White  "  Federation Information TarGetAutoDiscoverEpr vs Organization Relationship TarGetAutoDiscoverEpr "
    if ($Script:OrgRel.TarGetAutoDiscoverEpr -like $FedInfo.TarGetAutoDiscoverEpr) {
        Write-Host -ForegroundColor Green "   => Federation Information TarGetAutoDiscoverEpr matches the Organization Relationship TarGetAutoDiscoverEpr "
        Write-Host  "       Organization Relationship TarGetAutoDiscoverEpr:"  $Script:OrgRel.TarGetAutoDiscoverEpr
        Write-Host  "       Federation Information TarGetAutoDiscoverEpr:   "  $FedInfo.TarGetAutoDiscoverEpr
        $Script:tdTarGetAutoDiscoverEprVSColor = "green"
        $Script:tdTarGetAutoDiscoverEprVS_FL = "=> Federation Information TarGetAutoDiscoverEpr matches the Organization Relationship TarGetAutoDiscoverEpr"
    } else {
        Write-Host -ForegroundColor Red "   => Federation Information TarGetAutoDiscoverEpr should match the Organization Relationship TarGetAutoDiscoverEpr"
        Write-Host  "       Organization Relationship TarGetAutoDiscoverEpr:"  $Script:OrgRel.TarGetAutoDiscoverEpr
        Write-Host  "       Federation Information TarGetAutoDiscoverEpr:   "  $FedInfo.TarGetAutoDiscoverEpr
        $Script:tdTarGetAutoDiscoverEprVSColor = "red"
        $Script:tdTarGetAutoDiscoverEprVS_FL = "=> Federation Information TarGetAutoDiscoverEpr should match the Organization Relationship TarGetAutoDiscoverEpr"
    }
    Write-Host -ForegroundColor Yellow "`n  Reference: https://learn.microsoft.com/en-us/exchange/configure-a-federation-trust-exchange-2013-help#what-do-you-need-to-know-before-you-begin"
    PrintDynamicWidthLine
    $FedInfoDomainNames = ""
    $FedInfoDomainNames = ""
    foreach ($domain in $FedInfo.DomainNames.Domain) {
        if ($FedInfoDomainNames -ne "") {
            $FedInfoDomainNames += "; "
        }
        $FedInfoDomainNames += $domain
    }
    $Script:aux = $FedInfo.DomainNames
    $Script:FedInfoTokenIssuerUris = $FedInfo.TokenIssuerUris
    $Script:FedInfoTarGetAutoDiscoverEpr = $FedInfo.TarGetAutoDiscoverEpr
    $Script:FedInfoTarGetApplicationUri = $FedInfo.TarGetApplicationUri
    FedInfoHtml
}
function FedTrustCheck {
    $Script:FedTrust = $null
    Write-Host -ForegroundColor Green " Get-FederationTrust | fl ApplicationUri,TokenIssuerUri,OrgCertificate,TokenIssuerCertificate,
    TokenIssuerPrevCertificate, TokenIssuerMetadataEpr,TokenIssuerEpr"
    PrintDynamicWidthLine
    $Script:FedTrust = Get-FederationTrust | Select-Object ApplicationUri, TokenIssuerUri, OrgCertificate, TokenIssuerCertificate, TokenIssuerPrevCertificate, TokenIssuerMetadataEpr, TokenIssuerEpr
    $FedTrust
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Federation Trust"
    PrintDynamicWidthLine
    $CurrentTime = Get-Date
    Write-Host -ForegroundColor White " Federation Trust Application Uri:"
    if ($FedTrust.ApplicationUri -like "FYDIBOHF25SPDLT.$Script:ExchangeOnPremDomain") {
        Write-Host -ForegroundColor Green " " $FedTrust.ApplicationUri
        $Script:tdFedTrustApplicationUriColor = "green"
        $Script:tdFedTrustApplicationUriFL = $FedTrust.ApplicationUri
    } else {
        Write-Host -ForegroundColor Red "  Federation Trust Application Uri Should be "$FedTrust.ApplicationUri
        $Script:tdFedTrustApplicationUriColor = "red"
        $Script:tdFedTrustApplicationUriFL = "  Federation Trust Application Uri Should be $FedTrust.ApplicationUri"
    }
    Write-Host -ForegroundColor White " TokenIssuerUri:"
    if ($FedTrust.TokenIssuerUri.AbsoluteUri -like "urn:federation:MicrosoftOnline") {
        Write-Host -ForegroundColor Green " "$FedTrust.TokenIssuerUri.AbsoluteUri
        $Script:tdFedTrustTokenIssuerUriColor = "green"
        $Script:tdFedTrustTokenIssuerUriFL = $FedTrust.TokenIssuerUri.AbsoluteUri
    } else {
        Write-Host -ForegroundColor Red " Federation Trust TokenIssuerUri should be urn:federation:MicrosoftOnline"
        $Script:tdFedTrustTokenIssuerUriColor = "red"
        $Script:tdFedTrustTokenIssuerFL = " Federation Trust TokenIssuerUri is currently $FedTrust.TokenIssuerUri.AbsoluteUri but should be urn:federation:MicrosoftOnline"
    }
    Write-Host -ForegroundColor White " Federation Trust Certificate Expiry:"
    if ($FedTrust.OrgCertificate.NotAfter.Date -gt $CurrentTime) {
        Write-Host -ForegroundColor Green "  Not Expired"
        Write-Host  "   - Expires on " $FedTrust.OrgCertificate.NotAfter.DateTime
        $Script:tdFedTrustOrgCertificateNotAfterDateColor = "green"
        $Script:tdFedTrustOrgCertificateNotAfterDateFL = $FedTrust.OrgCertificate.NotAfter.DateTime
    } else {
        Write-Host -ForegroundColor Red " Federation Trust Certificate is Expired on " $FedTrust.OrgCertificate.NotAfter.DateTime
        $Script:tdFedTrustOrgCertificateNotAfterDateColor = "red"
        $Script:tdFedTrustOrgCertificateNotAfterDateFL = $FedTrust.OrgCertificate.NotAfter.DateTime
    }
    Write-Host -ForegroundColor White " `Federation Trust Token Issuer Certificate Expiry:"
    if ($FedTrust.TokenIssuerCertificate.NotAfter.DateTime -gt $CurrentTime) {
        Write-Host -ForegroundColor Green "  Not Expired"
        Write-Host  "   - Expires on " $FedTrust.TokenIssuerCertificate.NotAfter.DateTime
        $Script:tdFedTrustTokenIssuerCertificateNotAfterDateTimeColor = "green"
        $Script:tdFedTrustTokenIssuerCertificateNotAfterDateTimeFL = $FedTrust.TokenIssuerCertificate.NotAfter.DateTime
    } else {
        Write-Host -ForegroundColor Red "  Federation Trust TokenIssuerCertificate Expired on " $FedTrust.TokenIssuerCertificate.NotAfter.DateTime
        $Script:tdFedTrustTokenIssuerCertificateNotAfterDateTimeColor = "red"
        $Script:tdFedTrustTokenIssuerCertificateNotAfterDateTimeFL = $FedTrust.TokenIssuerCertificate.NotAfter.DateTime
    }
    $FedTrustTokenIssuerMetadataEpr = "https://nexus.microsoftonline-p.com/FederationMetadata/2006-12/FederationMetadata.xml"
    Write-Host -ForegroundColor White " `Token Issuer Metadata EPR:"
    if ($FedTrust.TokenIssuerMetadataEpr.AbsoluteUri -like $FedTrustTokenIssuerMetadataEpr) {
        Write-Host -ForegroundColor Green "  Token Issuer Metadata EPR is " $FedTrust.TokenIssuerMetadataEpr.AbsoluteUri
        #test if it can be reached
        $Script:tdFedTrustTokenIssuerMetadataEprAbsoluteUriColor = "green"
        $Script:tdFedTrustTokenIssuerMetadataEprAbsoluteUriFL = $FedTrust.TokenIssuerMetadataEpr.AbsoluteUri
    } else {
        Write-Host -ForegroundColor Red " Token Issuer Metadata EPR is Not " $FedTrust.TokenIssuerMetadataEpr.AbsoluteUri
        $Script:tdFedTrustTokenIssuerMetadataEprAbsoluteUriColor = "red"
        $Script:tdFedTrustTokenIssuerMetadataEprAbsoluteUriFL = $FedTrust.TokenIssuerMetadataEpr.AbsoluteUri
    }
    $FedTrustTokenIssuerEpr = "https://login.microsoftonline.com/extSTS.srf"
    Write-Host -ForegroundColor White " Token Issuer EPR:"
    if ($FedTrust.TokenIssuerEpr.AbsoluteUri -like $FedTrustTokenIssuerEpr) {
        Write-Host -ForegroundColor Green "  Token Issuer EPR is:" $FedTrust.TokenIssuerEpr.AbsoluteUri
        #test if it can be reached
        $Script:tdFedTrustTokenIssuerEprAbsoluteUriColor = "green"
        $Script:tdFedTrustTokenIssuerEprAbsoluteUriFL = $FedTrust.TokenIssuerEpr.AbsoluteUri
    } else {
        Write-Host -ForegroundColor Red "  Token Issuer EPR is Not:" $FedTrust.TokenIssuerEpr.AbsoluteUri
        $Script:tdFedTrustTokenIssuerEprAbsoluteUriColor = "red"
        $Script:tdFedTrustTokenIssuerEprAbsoluteUriFL = $FedTrust.TokenIssuerEpr.AbsoluteUri
    }
    $Script:FedInfoTokenIssuerUris = $FedInfo.TokenIssuerUris
    $Script:FedInfoTarGetApplicationUri = $FedInfo.TarGetApplicationUri
    $script:FedInfoTarGetAutoDiscoverEpr = $FedInfo.TarGetAutoDiscoverEpr
    fedTrustHtml
    Write-Host -ForegroundColor Yellow "`n  Reference: https://learn.microsoft.com/en-us/exchange/configure-a-federation-trust-exchange-2013-help"
}
function AvailabilityAddressSpaceCheck {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-AvailabilityAddressSpace $Script:ExchangeOnlineDomain | fl ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
    PrintDynamicWidthLine
    $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $Script:ExchangeOnlineDomain -ErrorAction SilentlyContinue | Select-Object ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    if (!$AvailabilityAddressSpace) {
        $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $Script:ExchangeOnlineDomain -ErrorAction SilentlyContinue | Select-Object ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    }
    $AvailabilityAddressSpace
    $Script:tdAvailabilityAddressSpaceName = $AvailabilityAddressSpace.Name
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - On-Prem Availability Address Space Check"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " ForestName: "
    if ($AvailabilityAddressSpace.ForestName -like $Script:ExchangeOnlineDomain) {
        Write-Host -ForegroundColor Green " " $AvailabilityAddressSpace.ForestName
        $Script:tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $Script:tdAvailabilityAddressSpaceForestColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  ForestName appears not to be correct."
        Write-Host -ForegroundColor White " Should contain the " $Script:ExchangeOnlineDomain
        $Script:tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $Script:tdAvailabilityAddressSpaceForestColor = "red"
    }
    Write-Host -ForegroundColor White " UserName: "
    if ($AvailabilityAddressSpace.UserName -like "") {
        Write-Host -ForegroundColor Green "  Blank"
        $Script:tdAvailabilityAddressSpaceUserName = " Blank"
        $Script:tdAvailabilityAddressSpaceUserNameColor = "green"
    } else {
        Write-Host -ForegroundColor Red " UserName is NOT correct. "
        Write-Host -ForegroundColor White "  Normally it should be blank"
        $Script:tdAvailabilityAddressSpaceUserName = $AvailabilityAddressSpace.UserName
        $Script:tdAvailabilityAddressSpaceUserNameColor = "red"
    }
    Write-Host -ForegroundColor White " UseServiceAccount: "
    if ($AvailabilityAddressSpace.UseServiceAccount -like "True") {
        Write-Host -ForegroundColor Green "  True"
        $Script:tdAvailabilityAddressSpaceUseServiceAccount = $AvailabilityAddressSpace.UseServiceAccount
        $Script:tdAvailabilityAddressSpaceUseServiceAccountColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  UseServiceAccount appears not to be correct."
        Write-Host -ForegroundColor White "  Should be True"
        $Script:tdAvailabilityAddressSpaceUseServiceAccount = $AvailabilityAddressSpace.UseServiceAccount
        $Script:tdAvailabilityAddressSpaceUseServiceAccountColor = "red"
    }
    Write-Host -ForegroundColor White " AccessMethod:"
    if ($AvailabilityAddressSpace.AccessMethod -like "InternalProxy") {
        Write-Host -ForegroundColor Green "  InternalProxy"
        $Script:tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $Script:tdAvailabilityAddressSpaceAccessMethodColor = "green"
    } else {
        Write-Host -ForegroundColor Red " AccessMethod appears not to be correct."
        Write-Host -ForegroundColor White " Should be InternalProxy"
        $Script:tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $Script:tdAvailabilityAddressSpaceAccessMethodColor = "red"
    }
    Write-Host -ForegroundColor White " ProxyUrl: "
    $Script:tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
    if ([String]::Equals($Script:tdAvailabilityAddressSpaceProxyUrl, $Script:ExchangeOnPremEWS, [StringComparison]::OrdinalIgnoreCase)) {
        Write-Host -ForegroundColor Green " "$AvailabilityAddressSpace.ProxyUrl
        #$Script:tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $Script:tdAvailabilityAddressSpaceProxyUrlColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  ProxyUrl appears not to be correct."
        Write-Host -ForegroundColor White "  Should be $Script:ExchangeOnPremEWS[0] and not $Script:tdAvailabilityAddressSpaceProxyUrl"
        #$Script:tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $Script:tdAvailabilityAddressSpaceProxyUrlColor = "red"
    }
    Write-Host -ForegroundColor Yellow "`n  Reference: https://learn.microsoft.com/en-us/powershell/module/exchange/Get-availabilityaddressspace?view=exchange-ps"
    AvailabilityAddressSpaceHtml
}
function AutoDVirtualDCheck {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-AutoDiscoverVirtualDirectory -Server $($server) | Select Identity,Name,ExchangeVersion,*authentication*"
    PrintDynamicWidthLine
    FetchAutoDiscoverInformation
    $Script:AutoDiscoveryVirtualDirectory
    #$AutoDFL = $Script:AutoDiscoveryVirtualDirectory | Format-List
    $script:html += ""
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - On-Prem Get-AutoDiscoverVirtualDirectory"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White "  WSSecurityAuthentication:"
    if ($Script:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication -eq "True") {
        foreach ( $ser in $Script:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)"
            $Script:AutoD_VD_Identity = $ser.Identity
            $Script:AutoD_VD_Name = $ser.Name
            $Script:AutoD_VD_InternalAuthenticationMethods = $ser.InternalAuthenticationMethods
            $Script:AutoD_VD_ExternalAuthenticationMethods = $ser.ExternalAuthenticationMethods
            $Script:AutoD_VD_WSAuthentication = $ser.WSSecurityAuthentication
            $Script:AutoD_VD_WSAuthenticationColor = "green"
            $Script:AutoD_VD_WindowsAuthentication = $ser.WindowsAuthentication
            if ($Script:AutoD_VD_WindowsAuthentication -eq "True") {
                $Script:AutoD_VD_WindowsAuthenticationColor = "green"
            } else {
                $Script:AutoD_VD_WindowsAuthenticationColor = "red"
            }
            $Script:AutoD_VD_InternalNblBypassUrl = $ser.InternalNblBypassUrl
            $Script:AutoD_VD_InternalUrl = $ser.InternalUrl
            $Script:AutoD_VD_ExternalUrl = $ser.ExternalUrl
            autoDVDHtmlOK
            $Script:serWSSecurityAuthenticationColor = "Green"
        }
    } else {
        Write-Host -ForegroundColor Red " WSSecurityAuthentication is NOT correct."
        foreach ( $ser in $Script:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)"
            $Script:serWSSecurityAuthenticationColor = "Red"
            Write-Host " $($ser.Identity) "
            $Script:AutoD_VD_Identity = $ser.Identity
            $Script:AutoD_VD_Name = $ser.Name
            $Script:AutoD_VD_InternalAuthenticationMethods = $ser.InternalAuthenticationMethods
            $Script:AutoD_VD_ExternalAuthenticationMethods = $ser.ExternalAuthenticationMethods
            $Script:AutoD_VD_WSAuthentication = $ser.WSSecurityAuthentication
            $Script:AutoD_VD_WSAuthenticationColor = "green"
            $Script:AutoD_VD_WindowsAuthentication = $ser.WindowsAuthentication
            if ($Script:AutoD_VD_WindowsAuthentication -eq "True") {
                $Script:AutoD_VD_WindowsAuthenticationColor = "green"
            } else {
                $Script:AutoD_VD_WindowsAuthenticationColor = "red"
            }
            $Script:AutoD_VD_InternalNblBypassUrl = $ser.InternalNblBypassUrl
            $Script:AutoD_VD_InternalUrl = $ser.InternalUrl
            $Script:AutoD_VD_ExternalUrl = $ser.ExternalUrl
            autoDVDHtmlNotOK
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)"
            $Script:serWSSecurityAuthenticationColor = "Red"
        }
        Write-Host -ForegroundColor White "  Should be True "
    }
    Write-Host -ForegroundColor White "`n  WindowsAuthentication:"
    if ($Script:AutoDiscoveryVirtualDirectory.WindowsAuthentication -eq "True") {
        foreach ( $ser in $Script:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
    } else {
        Write-Host -ForegroundColor Red " WindowsAuthentication is NOT correct."
        foreach ( $ser in $Script:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
        Write-Host -ForegroundColor White "  Should be True "
    }
    Write-Host -ForegroundColor Yellow "`n  Reference: https://learn.microsoft.com/en-us/powershell/module/exchange/Get-AutoDiscovervirtualdirectory?view=exchange-ps"
}
function EWSVirtualDirectoryCheck {
    Write-Host -ForegroundColor Green " Get-WebServicesVirtualDirectory -Server $($server)| Select Identity,Name,ExchangeVersion,*Authentication*,*url"
    PrintDynamicWidthLine
    $Script:WebServicesVirtualDirectory
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Get-WebServicesVirtualDirectory"
    PrintDynamicWidthLine
    EWSVirtualDHeaderHtml
    Write-Host -ForegroundColor White "  WSSecurityAuthentication:"
    if ($Script:WebServicesVirtualDirectory.WSSecurityAuthentication -like "True") {
        foreach ( $EWS in $Script:WebServicesVirtualDirectory) {
            Write-Host " $($EWS.Identity)"
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) "
            $Script:EwsVDIdentity = $EWS.Identity
            $Script:EwsVDName = $EWS.Name
            $Script:EwsVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $Script:EwsVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $Script:EwsVD_WSAuthentication = $EWS.WSSecurityAuthentication
            $Script:EwsVD_WSAuthenticationColor = "green"
            $Script:EwsVDWindowsAuthentication = $EWS.WindowsAuthentication
            if ($Script:EwsVDWindowsAuthentication -eq "True") {
                $Script:EwsVDWindowsAuthenticationColor = "green"
            } else {
                $Script:EWS_DWindowsAuthenticationColor = "red"
            }
            $Script:EwsVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
            $Script:EwsVDInternalUrl = $EWS.InternalUrl
            $Script:EwsVDExternalUrl = $EWS.ExternalUrl
            EwsVDHtmlOK
        }
    } else {
        Write-Host -ForegroundColor Red " WSSecurityAuthentication should be True."
        foreach ( $EWS in $Script:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication) "
            $Script:EwsVDIdentity = $EWS.Identity
            $Script:EwsVDName = $EWS.Name
            $Script:EwsVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $Script:EwsVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $Script:EwsVD_WSAuthentication = $EWS.WSSecurityAuthentication
            $Script:EwsVD_WSAuthenticationColor = "green"
            $Script:EwsVDWindowsAuthentication = $EWS.WindowsAuthentication
            if ($Script:EwsVDWindowsAuthentication -eq "True") {
                $Script:EwsVDWindowsAuthenticationColor = "green"
            } else {
                $Script:EWS_DWindowsAuthenticationColor = "red"
            }
            $Script:EwsVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
            $Script:EwsVDInternalUrl = $EWS.InternalUrl
            $Script:EwsVDExternalUrl = $EWS.ExternalUrl
        }
        Write-Host -ForegroundColor White "  Should be True"
    }
    Write-Host -ForegroundColor White "`n  WindowsAuthentication:"
    if ($Script:WebServicesVirtualDirectory.WindowsAuthentication -like "True") {
        foreach ( $EWS in $Script:WebServicesVirtualDirectory) {
            Write-Host " $($EWS.Identity)"
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($EWS.WindowsAuthentication) "
        }
    } else {
        Write-Host -ForegroundColor Red " WindowsAuthentication should be True."
        foreach ( $EWS in $Script:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication) "
        }
        Write-Host -ForegroundColor White "  Should be True"
    }
    $html | Out-File -FilePath $Script:htmlFile
}
function TestOrgRel {
    PrintDynamicWidthLine
    #$Script:TestFail = 0
    $Script:OrgRelIdentity = $Script:OrgRel.Identity
    $Script:OrgRelTarGetApplicationUri = $Script:OrgRel.TarGetApplicationUri
    if ( $Script:OrgRelTarGetApplicationUri -like "Outlook.com" -or $Script:OrgRelTarGetApplicationUri -like "outlook.com") {
        Write-Host -ForegroundColor Green "Test-OrganizationRelationship -Identity $Script:OrgRelIdentity  -UserIdentity $Script:UserOnPrem"
        #need to grab errors and provide alerts in error case
        PrintDynamicWidthLine
        $TestOrgRel = Test-OrganizationRelationship -Identity "$($Script:OrgRelIdentity)"  -UserIdentity $Script:UserOnPrem -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        #$TestOrgRel
        if ($TestOrgRel[16] -like "No Significant Issues to Report") {
            Write-Host -ForegroundColor Green "`n No Significant Issues to Report"
            TestOrgRelHtmlOK
        } else {
            TestOrgRelHtmlNotOK
            Write-Host -ForegroundColor Red "`n Test Organization Relationship Completed with errors"
        }
        $TestOrgRel[0]
        $TestOrgRel[1]
        $i = 0
        while ($i -lt $TestOrgRel.Length) {
            $element = $TestOrgRel[$i]
            #if ($element.Contains("RESULT: Success.")) {
            if ($element -like "*RESULT: Success.*") {
                $TestOrgRelStep = $TestOrgRel[$i - 1]
                $TestOrgRelStep
                Write-Host -ForegroundColor Green "$element"
                if (![string]::IsNullOrWhitespace($TestOrgRelStep)) {
                    $Script:html += "
                    <div></b> <span style='color:black'> <b> $TestOrgRelStep :</b></span> <span style='color:green'>$element</span></div>"
                }
            }

            else {
                if ($element -like "*RESULT: Error*") {
                    $TestOrgRelStep = $TestOrgRel[$i - 1]
                    $TestOrgRelStep
                    Write-Host -ForegroundColor Red "$element"
                    if (![string]::IsNullOrWhitespace($TestOrgRelStep)) {
                        $Script:html += "
                        <div></b> <span style='color:black'> <b> $TestOrgRelStep : </b></span> <span style='color:red'>$element</span></div>"
                    }
                }
            }
            $i++
        }
    } else {
        Write-Host -ForegroundColor Green " Test-OrganizationRelationship -Identity $Script:OrgRelIdentity  -UserIdentity $Script:UserOnPrem"
        #need to grab errors and provide alerts in error case
        PrintDynamicWidthLine
        Write-Host -ForegroundColor Red "`n Test-OrganizationRelationship can't be run if the Organization Relationship TarGet Application uri is not correct. Organization Relationship TarGet Application Uri should be Outlook.com"
        TestOrgRelHtmlNoUri
    }
    Write-Host -ForegroundColor Yellow "`n  Reference: https://techcommunity.microsoft.com/t5/exchange-team-blog/how-to-address-federation-trust-issues-in-hybrid-configuration/ba-p/1144285"
    PrintDynamicWidthLine
    $Script:html += "</td>
    </tr>"
    $html | Out-File -FilePath $Script:htmlFile
}
function TestFedTrust {
    PrintDynamicWidthLine
    $TestFedTrustFail = 0
    $a = Test-FederationTrust -UserIdentity $Script:UserOnPrem -verbose -ErrorAction SilentlyContinue #fails the first time on multiple occasions so we have a ghost FedTrustCheck
    if ($a) {
        #Using to suppress waring for unused variable
    }
    Write-Host -ForegroundColor Green  " Test-FederationTrust -UserIdentity $Script:UserOnPrem -verbose"
    PrintDynamicWidthLine
    $TestFedTrust = Test-FederationTrust -UserIdentity $Script:UserOnPrem -verbose -ErrorAction SilentlyContinue
    $TestFedTrust
    $Script:html += "<tr>
    <th ColSpan='2' style='color:white;'><b>Summary - On Premise Test-FederationTrust</b></th>
    </tr>
    <tr>
    <td><b> Test-FederationTrust -UserIdentity $Script:UserOnPrem</b></td>
    <td>"
    $i = 0
    while ($i -lt $TestFedTrust.type.Count) {
        $test = $TestFedTrust.type[$i]
        $testType = $TestFedTrust.Type[$i]
        $testMessage = $TestFedTrust.Message[$i]
        $TestFedTrustID = $($TestFedTrust.ID[$i])
        if ($test -eq "Error") {
            $Script:html += "

            <div> <span style='color:red'><b>$testType :</b></span> - <div> <b>$TestFedTrustID </b> - $testMessage  </div>
            "
            $TestFedTrustFail++
        }
        if ($test -eq "Success") {

            $Script:html += "

            <div> <span style='color:green'><b>$testType :</b> </span> - <b>$TestFedTrustID </b> - $testMessage</div>"
        }
        $i++
    }

    if ($TestFedTrustFail -eq 0) {
        Write-Host -ForegroundColor Green " Federation Trust Successfully tested"
        $Script:html += "
        <p></p>
        <div class=�green�> <span style='color:green'> Federation Trust Successfully tested </span></div>"
    } else {
        Write-Host -ForegroundColor Red " Federation Trust test with Errors"
        $Script:html += "
        <p></p>
        <div class=�red�> <span style='color:red'> Federation Trust tested with Errors </span></div>"
    }
    $TestFederationTrustCertificate = Test-FederationTrustCertificate -ErrorAction SilentlyContinue
    if ($TestFederationTrustCertificate) {
        PrintDynamicWidthLine
        Write-Host -ForegroundColor Green " Test-FederationTrustCertificate"
        PrintDynamicWidthLine
        $TestFederationTrustCertificate
        $Script:html += "<tr>
                <th ColSpan='2' style='color:white;'><b>Summary - Test-FederationTrustCertificate</b></th>
                </tr>
                <tr>
                <td><b> Test-FederationTrustCertificate</b></td>
                <td>"
        $j = 0
        while ($j -lt $TestFederationTrustCertificate.Count) {
            $TestFederationTrustCertificateJ = "<div>" + $TestFederationTrustCertificate.site[$j] + "</div><div>" + $TestFederationTrustCertificate.state[$j] + "</div><div>" + $TestFederationTrustCertificate.Thumbprint[$j] + "</div>"
            $Script:html += "
                $TestFederationTrustCertificateJ
                "
            $j++
        }
        $Script:html += "</td>"
    }
    $html | Out-File -FilePath $Script:htmlFile
}
function IntraOrgConCheck {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-IntraOrganizationConnector | Select Name,TarGetAddressDomains,DiscoveryEndpoint,Enabled"
    PrintDynamicWidthLine
    $IOC =$Script:IntraOrgCon | Format-List
    $IOC
    $Script:tdIntraOrgTarGetAddressDomain =$Script:IntraOrgCon.TarGetAddressDomains
    $Script:tdDiscoveryEndpoint =$Script:IntraOrgCon.DiscoveryEndpoint
    $Script:tdEnabled =$Script:IntraOrgCon.Enabled

    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Get-IntraOrganizationConnector"
    PrintDynamicWidthLine
    $IntraOrgTarGetAddressDomain =$Script:IntraOrgCon.TarGetAddressDomains.Domain
    $IntraOrgTarGetAddressDomain = $IntraOrgTarGetAddressDomain.ToLower()
    Write-Host -ForegroundColor White " TarGet Address Domains: "
    if ($Script:IntraOrgCon.TarGetAddressDomains -like "*$Script:ExchangeOnlineDomain*" -or $Script:IntraOrgCon.TarGetAddressDomains -like "*$Script:ExchangeOnlineAltDomain*" ) {
        Write-Host -ForegroundColor Green " "$Script:IntraOrgCon.TarGetAddressDomains
        $Script:tdIntraOrgTarGetAddressDomainColor = "green"
    } else {
        Write-Host -ForegroundColor Red " TarGet Address Domains appears not to be correct."
        Write-Host -ForegroundColor White " Should contain the $Script:ExchangeOnlineDomain domain or the $Script:ExchangeOnlineAltDomain domain."
        $Script:tdIntraOrgTarGetAddressDomainColor = "red"
    }
    Write-Host -ForegroundColor White " DiscoveryEndpoint: "
    if ($Script:IntraOrgCon.DiscoveryEndpoint -like "https://AutoDiscover-s.outlook.com/AutoDiscover/AutoDiscover.svc") {
        Write-Host -ForegroundColor Green "  https://AutoDiscover-s.outlook.com/AutoDiscover/AutoDiscover.svc"
        $Script:tdDiscoveryEndpointColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  The DiscoveryEndpoint appears not to be correct. "
        Write-Host -ForegroundColor White "  It should represent the address of EXO AutoDiscover endpoint."
        Write-Host  "  Examples: https://AutoDiscover-s.outlook.com/AutoDiscover/AutoDiscover.svc; https://outlook.office365.com/AutoDiscover/AutoDiscover.svc "
        $Script:tdDiscoveryEndpointColor = "red"
    }
    Write-Host -ForegroundColor White " Enabled: "
    if ($Script:IntraOrgCon.Enabled -like "True") {
        Write-Host -ForegroundColor Green "  True "
        $Script:tdEnabledColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  On-Prem Intra Organization Connector is not Enabled"
        Write-Host -ForegroundColor White "  In order to use OAuth it Should be True."
        Write-Host "  If it is set to False, the Organization Relationship (DAuth) , if enabled, is used for the Hybrid Availability Sharing"
        $Script:tdEnabledColor = "red"
    }
    Write-Host -ForegroundColor Yellow "https://techcommunity.microsoft.com/t5/exchange-team-blog/demystifying-hybrid-free-busy-what-are-the-moving-parts/ba-p/607704"
    IntraOrgConCheckHtml
}
function AuthServerCheck {
    #PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-AuthServer | Select Name, Realm, IssuerIdentifier, TokenIssuingEndpoint, AuthMetadataUrl, Enabled"
    PrintDynamicWidthLine
    $AuthServer = Get-AuthServer |
        Where-Object { $_.Name -like "EvoSts*" } |
        Sort-Object Name |
        Select-Object -First 1 Name, Realm, IssuerIdentifier, TokenIssuingEndpoint, AuthMetadataUrl, Enabled
    $AuthServer
    $Script:tDAuthServerIssuerIdentifier = $AuthServer.IssuerIdentifier
    $Script:tDAuthServerTokenIssuingEndpoint = $AuthServer.TokenIssuingEndpoint
    $Script:tDAuthServerAuthMetadataUrl = $AuthServer.AuthMetadataUrl
    $Script:tDAuthServerEnabled = $AuthServer.Enabled
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Auth Server"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " IssuerIdentifier: "
    if ($AuthServer.IssuerIdentifier -like "https://sts.windows.net/$($AuthServer.Realm)/" ) {
        Write-Host -ForegroundColor Green " " $AuthServer.IssuerIdentifier
        $Script:tDAuthServerIssuerIdentifierColor = "green"
    } else {
        Write-Host -ForegroundColor Red " IssuerIdentifier appears not to be correct."
        Write-Host -ForegroundColor White " Should be https://sts.windows.net/<Cloud Tenant ID>/"
        $Script:tDAuthServerIssuerIdentifierColor = "red"
    }
    Write-Host -ForegroundColor White " TokenIssuingEndpoint: "
    if ($AuthServer.TokenIssuingEndpoint -like "https://login.windows.net/common/oauth2/token*" ) {
        Write-Host -ForegroundColor Green " " $AuthServer.TokenIssuingEndpoint
        $Script:tDAuthServerTokenIssuingEndpointColor = "green"
    } else {
        Write-Host -ForegroundColor Red " TokenIssuingEndpoint appears not to be correct."
        Write-Host -ForegroundColor White " Should be  https://login.windows.net/common/oauth2/token"
        $Script:tDAuthServerTokenIssuingEndpointColor = "red"
    }
    Write-Host -ForegroundColor White " AuthMetadataUrl: "
    if ($AuthServer.AuthMetadataUrl -like "https://login.windows.net/*/federationmetadata/2007-06/federationmetadata.xml" ) {
        Write-Host -ForegroundColor Green " " $AuthServer.AuthMetadataUrl
        $Script:tDAuthServerAuthMetadataUrlColor = "green"
    } else {
        Write-Host -ForegroundColor Red " AuthMetadataUrl appears not to be correct."
        Write-Host -ForegroundColor White " Should be  https://login.windows.net/<Initial Tenant Domain>/FederationMetadata/2007-06/FederationMetadata.xml"
        $Script:tDAuthServerAuthMetadataUrlColor = "red"
    }
    Write-Host -ForegroundColor White " Enabled: "
    if ($AuthServer.Enabled -like "True" ) {
        Write-Host -ForegroundColor Green " " $AuthServer.Enabled
        $Script:tDAuthServerEnabledColor = "green"
    } else {
        Write-Host -ForegroundColor Red " Enabled: False "
        Write-Host -ForegroundColor White " Should be True"
        $Script:tDAuthServerEnabledColor = "red"
    }
    AuthServerCheckHtml
}
function PartnerApplicationCheck {
    #PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-PartnerApplication |  ?{`$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000'
    -and `$_.Realm -eq ''} | Select Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer,
    AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name"
    PrintDynamicWidthLine
    $PartnerApplication = Get-PartnerApplication | Where-Object { $_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000' -and $_.Realm -eq '' } | Select-Object Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer, AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name
    $PartnerApplication
    $Script:tdPartnerApplicationEnabled = $PartnerApplication.Enabled
    $Script:tdPartnerApplicationApplicationIdentifier = $PartnerApplication.ApplicationIdentifier
    $Script:tdPartnerApplicationCertificateStrings = $PartnerApplication.CertificateStrings
    $Script:tdPartnerApplicationAuthMetadataUrl = $PartnerApplication.AuthMetadataUrl
    $Script:tdPartnerApplicationRealm = $PartnerApplication.Realm
    $Script:tdPartnerApplicationUseAuthServer = $PartnerApplication.UseAuthServer
    $Script:tdPartnerApplicationAcceptSecurityIdentifierInformation = $PartnerApplication.AcceptSecurityIdentifierInformation
    $Script:tdPartnerApplicationLinkedAccount = $PartnerApplication.LinkedAccount
    $Script:tdPartnerApplicationIssuerIdentifier = $PartnerApplication.IssuerIdentifier
    $Script:tdPartnerApplicationAppOnlyPermissions = $PartnerApplication.AppOnlyPermissions
    $Script:tdPartnerApplicationActAsPermissions = $PartnerApplication.ActAsPermissions
    $Script:tdPartnerApplicationName = $PartnerApplication.Name
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Partner Application"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " Enabled: "
    if ($PartnerApplication.Enabled -like "True" ) {
        Write-Host -ForegroundColor Green " " $PartnerApplication.Enabled
        $Script:tdPartnerApplicationEnabledColor = "green"
    } else {
        Write-Host -ForegroundColor Red " Enabled: False "
        Write-Host -ForegroundColor White " Should be True"
        $Script:tdPartnerApplicationEnabledColor = "red"
    }
    Write-Host -ForegroundColor White " ApplicationIdentifier: "
    if ($PartnerApplication.ApplicationIdentifier -like "00000002-0000-0ff1-ce00-000000000000" ) {
        Write-Host -ForegroundColor Green " " $PartnerApplication.ApplicationIdentifier
        $Script:tdPartnerApplicationApplicationIdentifierColor = "green"
    } else {
        Write-Host -ForegroundColor Red " ApplicationIdentifier does not appear to be correct"
        Write-Host -ForegroundColor White " Should be 00000002-0000-0ff1-ce00-000000000000"
        $Script:tdPartnerApplicationApplicationIdentifierColor = "red"
    }
    Write-Host -ForegroundColor White " AuthMetadataUrl: "
    if ([string]::IsNullOrWhitespace( $PartnerApplication.AuthMetadataUrl)) {
        Write-Host -ForegroundColor Green "  Blank"
        $Script:tdPartnerApplicationAuthMetadataUrlColor = "green"
        $Script:tdPartnerApplicationAuthMetadataUrl = "Blank"
    } else {
        Write-Host -ForegroundColor Red " AuthMetadataUrl does not seem to be correct"
        Write-Host -ForegroundColor White " Should be Blank"
        $Script:tdPartnerApplicationAuthMetadataUrlColor = "red"
        $Script:tdPartnerApplicationAuthMetadataUrl = " Should be Blank"
    }
    Write-Host -ForegroundColor White " Realm: "
    if ([string]::IsNullOrWhitespace( $PartnerApplication.Realm)) {
        Write-Host -ForegroundColor Green "  Blank"
        $Script:tdPartnerApplicationRealmColor = "green"
        $Script:tdPartnerApplicationRealm = "Blank"
    } else {
        Write-Host -ForegroundColor Red "  Realm does not seem to be correct"
        Write-Host -ForegroundColor White " Should be Blank"
        $Script:tdPartnerApplicationRealmColor = "Red"
        $Script:tdPartnerApplicationRealm = "Should be Blank"
    }
    Write-Host -ForegroundColor White " LinkedAccount: "
    if ($PartnerApplication.LinkedAccount -like "$Script:ExchangeOnPremDomain/Users/Exchange Online-ApplicationAccount" -or $PartnerApplication.LinkedAccount -like "$Script:ExchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  ) {
        Write-Host -ForegroundColor Green " " $PartnerApplication.LinkedAccount
        $Script:tdPartnerApplicationLinkedAccountColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  LinkedAccount value does not appear to be correct"
        Write-Host -ForegroundColor White "  Should be $Script:ExchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"
        Write-Host "  If you value is empty, set it to correspond to the Exchange Online-ApplicationAccount which is located at the root of Users container in AD. After you make the change, reboot the Servers."
        Write-Host "  Example: contoso.com/Users/Exchange Online-ApplicationAccount"
        $Script:tdPartnerApplicationLinkedAccountColor = "red"
        $Script:tdPartnerApplicationLinkedAccount
    }
    PartnerApplicationCheckHtml
}
function ApplicationAccountCheck {
    #PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-user '$Script:ExchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount' | Select Name, RecipientType, RecipientTypeDetails, UserAccountControl"
    PrintDynamicWidthLine
    $ApplicationAccount = Get-user "$Script:ExchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount" | Select-Object Name, RecipientType, RecipientTypeDetails, UserAccountControl
    $ApplicationAccount
    $Script:tdApplicationAccountRecipientType = $ApplicationAccount.RecipientType
    $Script:tdApplicationAccountRecipientTypeDetails = $ApplicationAccount.RecipientTypeDetails
    $Script:tdApplicationAccountUserAccountControl = $ApplicationAccount.UserAccountControl
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Application Account"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " RecipientType: "
    if ($ApplicationAccount.RecipientType -like "User" ) {
        Write-Host -ForegroundColor Green " " $ApplicationAccount.RecipientType
        $Script:tdApplicationAccountRecipientTypeColor = "green"
    } else {
        Write-Host -ForegroundColor Red " RecipientType value is $ApplicationAccount.RecipientType "
        Write-Host -ForegroundColor White " Should be User"
        $Script:tdApplicationAccountRecipientTypeColor = "red"
    }
    Write-Host -ForegroundColor White " RecipientTypeDetails: "
    if ($ApplicationAccount.RecipientTypeDetails -like "LinkedUser" ) {
        Write-Host -ForegroundColor Green " " $ApplicationAccount.RecipientTypeDetails
        $Script:tdApplicationAccountRecipientTypeDetailsColor = "green"
    } else {
        Write-Host -ForegroundColor Red " RecipientTypeDetails value is $ApplicationAccount.RecipientTypeDetails"
        Write-Host -ForegroundColor White " Should be LinkedUser"
        $Script:tdApplicationAccountRecipientTypeDetailsColor = "red"
    }
    Write-Host -ForegroundColor White " UserAccountControl: "
    if ($ApplicationAccount.UserAccountControl -like "AccountDisabled, PasswordNotRequired, NormalAccount" ) {
        Write-Host -ForegroundColor Green " " $ApplicationAccount.UserAccountControl
        $Script:tdApplicationAccountUserAccountControlColor = "green"
    } else {
        Write-Host -ForegroundColor Red " UserAccountControl value does not seem correct"
        Write-Host -ForegroundColor White " Should be AccountDisabled, PasswordNotRequired, NormalAccount"
        $Script:tdApplicationAccountUserAccountControlColor = "red"
    }
    ApplicationAccountCheckHtml
}
function ManagementRoleAssignmentCheck {
    Write-Host -ForegroundColor Green " Get-ManagementRoleAssignment -RoleAssignee Exchange Online-ApplicationAccount | Select Name,Role -AutoSize"
    PrintDynamicWidthLine
    $ManagementRoleAssignment = Get-ManagementRoleAssignment -RoleAssignee "Exchange Online-ApplicationAccount" | Select-Object Name, Role
    $M = $ManagementRoleAssignment | Out-String
    $M
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Management Role Assignment for the Exchange Online-ApplicationAccount"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " Role: "
    if ($ManagementRoleAssignment.Role -like "*UserApplication*" ) {
        Write-Host -ForegroundColor Green "  UserApplication Role Assigned"
        $Script:tdManagementRoleAssignmentUserApplication = " UserApplication Role Assigned"
        $Script:tdManagementRoleAssignmentUserApplicationColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  UserApplication Role not present for the Exchange Online-ApplicationAccount"
        $Script:tdManagementRoleAssignmentUserApplication = " UserApplication Role not present"
        $Script:tdManagementRoleAssignmentUserApplicationColor = "red"
    }
    if ($ManagementRoleAssignment.Role -like "*ArchiveApplication*" ) {
        Write-Host -ForegroundColor Green "  ArchiveApplication Role Assigned"
        $Script:tdManagementRoleAssignmentArchiveApplication = " ArchiveApplication Role Assigned"
        $Script:tdManagementRoleAssignmentArchiveApplicationColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  ArchiveApplication Role not present for the Exchange Online-ApplicationAccount"
        $Script:tdManagementRoleAssignmentArchiveApplication = " ArchiveApplication Role not Assigned"
        $Script:tdManagementRoleAssignmentArchiveApplicationColor = "red"
    }
    if ($ManagementRoleAssignment.Role -like "*LegalHoldApplication*" ) {
        Write-Host -ForegroundColor Green "  LegalHoldApplication Role Assigned"
        $Script:tdManagementRoleAssignmentLegalHoldApplication = " LegalHoldApplication Role Assigned"
        $Script:tdManagementRoleAssignmentLegalHoldApplicationColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  LegalHoldApplication Role not present for the Exchange Online-ApplicationAccount"
        $Script:tdManagementRoleAssignmentLegalHoldApplication = " LegalHoldApplication Role Assigned"
        $Script:tdManagementRoleAssignmentLegalHoldApplicationColor = "green"
    }
    if ($ManagementRoleAssignment.Role -like "*Mailbox Search*" ) {
        Write-Host -ForegroundColor Green "  Mailbox Search Role Assigned"
        $Script:tdManagementRoleAssignmentMailboxSearch = " Mailbox Search Role Assigned"
        $Script:tdManagementRoleAssignmentMailboxSearchColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  Mailbox Search Role not present for the Exchange Online-ApplicationAccount"
        $Script:tdManagementRoleAssignmentMailboxSearch = " Mailbox Search Role Not Assigned"
        $Script:tdManagementRoleAssignmentMailboxSearchColor = "red"
    }
    if ($ManagementRoleAssignment.Role -like "*TeamMailboxLifecycleApplication*" ) {
        Write-Host -ForegroundColor Green "  TeamMailboxLifecycleApplication Role Assigned"
        $Script:tdManagementRoleAssignmentTeamMailboxLifecycleApplication = " TeamMailboxLifecycleApplication Role Assigned"
        $Script:tdManagementRoleAssignmentTeamMailboxLifecycleApplicationColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  TeamMailboxLifecycleApplication Role not present for the Exchange Online-ApplicationAccount"
        $Script:tdManagementRoleAssignmentTeamMailboxLifecycleApplication = " TeamMailboxLifecycleApplication Role Not Assigned"
        $Script:tdManagementRoleAssignmentTeamMailboxLifecycleApplicationColor = "red"
    }
    if ($ManagementRoleAssignment.Role -like "*MailboxSearchApplication*" ) {
        Write-Host -ForegroundColor Green "  MailboxSearchApplication Role Assigned"
        $Script:tdManagementRoleMailboxSearchApplication = " MailboxSearchApplication Role Assigned"
        $Script:tdManagementRoleMailboxSearchApplicationColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  MailboxSearchApplication Role not present for the Exchange Online-ApplicationAccount"
        $Script:tdManagementRoleMailboxSearchApplication = " MailboxSearchApplication Role Not Assigned"
        $Script:tdManagementRoleMailboxSearchApplicationColor = "red"
    }
    if ($ManagementRoleAssignment.Role -like "*MeetingGraphApplication*" ) {
        Write-Host -ForegroundColor Green "  MeetingGraphApplication Role Assigned"
        $Script:tdManagementRoleMeetingGraphApplication = " MeetingGraphApplication Role Assigned"
        $Script:tdManagementRoleMeetingGraphApplicationColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  MeetingGraphApplication Role not present for the Exchange Online-ApplicationAccount"
        $Script:tdManagementRoleMeetingGraphApplication = " MeetingGraphApplication Role Not Assigned"
        $Script:tdManagementRoleMeetingGraphApplicationColor = "red"
    }
    $Script:tdManagementRoleMeetingGraphApplication = " MailboxSearchApplication Role Assigned"
    $Script:tdManagementRoleMeetingGraphApplicationColor = "green"
    ManagementRoleAssignmentCheckHtml
}
function AuthConfigCheck {
    Write-Host -ForegroundColor Green " Get-AuthConfig | Select *Thumbprint, ServiceName, Realm, Name"
    PrintDynamicWidthLine
    $AuthConfig = Get-AuthConfig | Select-Object *Thumbprint, ServiceName, Realm, Name
    $AC = $AuthConfig | Format-List
    $AC
    $Script:tDAuthConfigName = $AuthConfig.Name
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Auth Config"
    PrintDynamicWidthLine
    if (![string]::IsNullOrWhitespace($AuthConfig.CurrentCertificateThumbprint)) {
        Write-Host " Thumbprint: "$AuthConfig.CurrentCertificateThumbprint
        Write-Host -ForegroundColor Green " Certificate is Assigned"
        $Script:tDAuthConfigCurrentCertificateThumbprint = $AuthConfig.CurrentCertificateThumbprint
        $Script:tDAuthConfigCurrentCertificateThumbprintColor = "green"
    } else {
        Write-Host " Thumbprint: "$AuthConfig.CurrentCertificateThumbprint
        Write-Host -ForegroundColor Red " No valid certificate Assigned "
        $Script:tDAuthConfigCurrentCertificateThumbprintColor = "red"
        $Script:tDAuthConfigCurrentCertificateThumbprint = "$AuthConfig.CurrentCertificateThumbprint - No valid certificate Assigned "
    }
    if ($AuthConfig.ServiceName -like "00000002-0000-0ff1-ce00-000000000000" ) {
        Write-Host " ServiceName: "$AuthConfig.ServiceName
        Write-Host -ForegroundColor Green " Service Name Seems correct"
        $Script:tDAuthConfigServiceNameColor = "green"
        $Script:tDAuthConfigServiceName = $AuthConfig.ServiceName
    } else {
        Write-Host " ServiceName: "$AuthConfig.ServiceName
        Write-Host -ForegroundColor Red " Service Name does not Seems correct. Should be 00000002-0000-0ff1-ce00-000000000000"
        $Script:tDAuthConfigServiceNameColor = "red"
        $Script:tDAuthConfigServiceName = "$AuthConfig.ServiceName  Should be 00000002-0000-0ff1-ce00-000000000000"
    }
    if ([string]::IsNullOrWhitespace($AuthConfig.Realm)) {
        Write-Host " Realm: "
        Write-Host -ForegroundColor Green " Realm is Blank"
        $Script:tDAuthConfigRealmColor = "green"
        $Script:tDAuthConfigRealm = " Realm is Blank"
    } else {
        Write-Host " Realm: "$AuthConfig.Realm
        Write-Host -ForegroundColor Red " Realm should be Blank"
        $Script:tDAuthConfigRealmColor = "red"
        $Script:tDAuthConfigRealm = "$Script:tDAuthConfig.Realm - Realm should be Blank"
    }
    AuthConfigCheckHtml
}
function CurrentCertificateThumbprintCheck {
    $thumb = Get-AuthConfig | Select-Object CurrentCertificateThumbprint
    $thumbprint = $thumb.CurrentCertificateThumbprint
    #PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-ExchangeCertificate -Thumbprint $thumbprint | Select FriendlyName, Issuer, Services, NotAfter, Status, HasPrivateKey, Subject, Thumb*"
    PrintDynamicWidthLine
    $CurrentCertificate = Get-ExchangeCertificate $thumb.CurrentCertificateThumbprint | Select-Object  FriendlyName, Issuer, Services, NotAfter, Status, HasPrivateKey, Subject, Thumb*
    $CC = $CurrentCertificate | Format-List
    $CC

    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Microsoft Exchange Server Auth Certificate"
    PrintDynamicWidthLine
    if ($CurrentCertificate.Issuer -like "CN=Microsoft Exchange Server Auth Certificate" ) {
        Write-Host " Issuer: " $CurrentCertificate.Issuer
        Write-Host -ForegroundColor Green " Issuer is CN=Microsoft Exchange Server Auth Certificate"
        $Script:tdCurrentCertificateIssuer = "   $($CurrentCertificate.Issuer) - Issuer is CN=Microsoft Exchange Server Auth Certificate"
        $Script:tdCurrentCertificateIssuerColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  Issuer is not CN=Microsoft Exchange Server Auth Certificate"
        $Script:tdCurrentCertificateIssuer = "   $($CurrentCertificate.Issuer) - Issuer is Not CN=Microsoft Exchange Server Auth Certificate"
        $Script:tdCurrentCertificateIssuerColor = "red"
    }
    if ($CurrentCertificate.Services -like "SMTP" ) {
        Write-Host " Services: " $CurrentCertificate.Services
        Write-Host -ForegroundColor Green "  Certificate enabled for SMTP"
        $Script:tdCurrentCertificateServices = "  $($tdCurrentCertificate.Services) - Certificate enabled for SMTP"
        $Script:tdCurrentCertificateServicesColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  Certificate Not enabled for SMTP"
        $Script:tdCurrentCertificateServices = "  $($tdCurrentCertificate.Services) - Certificate Not enabled for SMTP"
        $Script:tdCurrentCertificateServicesColor = "red"
    }
    if ($CurrentCertificate.Status -like "Valid" ) {
        Write-Host " Status: " $CurrentCertificate.Status
        Write-Host -ForegroundColor Green "  Certificate is valid"
        $Script:tdCurrentCertificateStatus = "  Certificate is valid"
        $Script:tdCurrentCertificateStatusColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  Certificate is not Valid"
        $Script:tdCurrentCertificateStatus = "  Certificate is Not Valid"
        $Script:tdCurrentCertificateStatusColor = "red"
    }
    if ($CurrentCertificate.Subject -like "CN=Microsoft Exchange Server Auth Certificate" ) {
        Write-Host " Subject: " $CurrentCertificate.Subject
        Write-Host -ForegroundColor Green "  Subject is CN=Microsoft Exchange Server Auth Certificate"
        $Script:tdCurrentCertificateSubject = "  Subject is CN=Microsoft Exchange Server Auth Certificate"
        $Script:tdCurrentCertificateSubjectColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  Subject is not CN=Microsoft Exchange Server Auth Certificate"
        $Script:tdCurrentCertificateSubject = "  $($CurrentCertificate.Subject) - Subject should be CN=Microsoft Exchange Server Auth Certificate"
        $Script:tdCurrentCertificateSubjectColor = "red"
    }
    Write-Host -ForegroundColor White "`n Checking Exchange Auth Certificate Distribution `n"
    $CheckAuthCertDistribution = foreach ($name in (Get-ExchangeServer).name) { Get-ExchangeCertificate -Thumbprint (Get-AuthConfig).CurrentCertificateThumbprint -Server $name -ErrorAction SilentlyContinue | Select-Object Identity, thumbprint, Services, subject }
    foreach ($Server in $CheckAuthCertDistribution) {
        $ServerName = ($Server -split "\.")[0]
        Write-Host -ForegroundColor White  "  Server: " $ServerName
        #Write-Host  "   Thumbprint: " $Thumbprint
        if ($Server.Thumbprint -like $thumbprint) {
            Write-Host  "   Thumbprint: "$Server.Thumbprint
            Write-Host  "   Subject: "$Server.Subject
            $ServerIdentity = $Server.Identity
            $Script:tdCheckAuthCertDistribution = "   <div>Certificate with Thumbprint: $($Server.Thumbprint) Subject: $($Server.Subject) is present in Server $ServerIdentity</div>"
            $Script:tdCheckAuthCertDistributionColor = "green"
        }
        if ($Server.Thumbprint -ne $thumbprint) {
            Write-Host -ForegroundColor Red "  Auth Certificate seems Not to be present in $ServerName"
            $Script:tdCheckAuthCertDistribution = "   Auth Certificate seems Not to be present in $ServerName"
            $Script:tdCheckAuthCertDistributionColor = "Red"
        }
    }
    CurrentCertificateThumbprintCheckHtml
}
function OAuthConnectivityCheck {
    Write-Host -ForegroundColor Green " Test-OAuthConnectivity -Service EWS -TarGetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $Script:UserOnPrem"
    PrintDynamicWidthLine
    $Script:OAuthConnectivity = Test-OAuthConnectivity -Service EWS -TarGetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $Script:UserOnPrem
    if ($Script:OAuthConnectivity.ResultType -ne 'Success') {
        $Script:OAuthConnectivity
    }
    if ($Script:OAuthConnectivity.Detail.FullId -like '*(401) Unauthorized*') {
        Write-Host -ForegroundColor Red "Error: The remote Server returned an error: (401) Unauthorized"
        if ($Script:OAuthConnectivity.Detail.FullId -like '*The user specified by the user-context in the token does not exist*') {
            Write-Host -ForegroundColor Yellow "The user specified by the user-context in the token does not exist"
            Write-Host "Please run Test-OAuthConnectivity with a different Exchange On Premises Mailbox"
        }
    }
    Write-Host -ForegroundColor Green " Summary - Test OAuth Connectivity"
    PrintDynamicWidthLine
    if ($Script:OAuthConnectivity.ResultType -like "Success") {
        Write-Host -ForegroundColor Green "$($Script:OAuthConnectivity.ResultType). OAuth Test was completed successfully "
        $Script:OAuthConnectivityResultType = " OAuth Test was completed successfully "
        $Script:OAuthConnectivityResultTypeColor = "green"
    } else {
        Write-Host -ForegroundColor Red " $Script:OAuthConnectivity.ResultType - OAuth Test was completed with Error. "
        Write-Host -ForegroundColor White " Please rerun Test-OAuthConnectivity -Service EWS -TarGetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        $Script:OAuthConnectivityResultType = " <div>OAuth Test was completed with Error.</div><div>Please rerun Test-OAuthConnectivity -Service EWS -TarGetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox <On Premises Mailbox> | fl to confirm the test failure</div>"
        $Script:OAuthConnectivityResultTypeColor = "red"
    }
    Write-Host -ForegroundColor Green " Reference: "
    Write-Host -ForegroundColor White " Configure OAuth authentication between Exchange and Exchange Online organizations"
    Write-Host -ForegroundColor Yellow " https://technet.microsoft.com/en-us/library/dn594521(v=exchg.150).aspx"
    OAuthConnectivityCheckHtml
}
function AutoDVirtualDCheckOauth {
    Write-Host -ForegroundColor Green " Get-AutoDiscoverVirtualDirectory -Server $($server) | Select Identity, Name,ExchangeVersion,*authentication*"
    PrintDynamicWidthLine
    FetchAutoDiscoverInformation
    $AD = $AutoDiscoveryVirtualDirectoryOAuth | Format-List
    $AD
    AutoDVirtualDCheckOauthHtmlHead
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Get-AutoDiscoverVirtualDirectory"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White "  InternalAuthenticationMethods"
    if ($AutoDiscoveryVirtualDirectoryOAuth.InternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  InternalAuthenticationMethods Include OAuth Authentication Method "
            $Script:AutoD_VD_Identity = $EWS.Identity
            $Script:AutoD_VD_Name = $EWS.Name
            $Script:AutoD_VD_InternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $Script:AutoD_VD_ExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $Script:AutoD_VD_WSAuthentication = $EWS.WSSecurityAuthentication
            $Script:AutoD_VD_WSAuthenticationColor = "green"
            $Script:AutoD_VD_WindowsAuthentication = $EWS.WindowsAuthentication
            $Script:AutoD_VD_OAuthAuthentication = $EWS.OAuthAuthentication
            if ($Script:AutoD_VD_WindowsAuthentication -eq "True") {
                $Script:AutoD_VD_WindowsAuthenticationColor = "green"
            } else {
                $Script:AutoD_VD_WindowsAuthenticationColor = "red"
            }
            if ($Script:AutoD_VD_OAuthAuthentication -eq "True") {
                $Script:AutoD_VD_OAuthAuthenticationColor = "green"
            } else {
                $Script:AutoD_VD_OAuthAuthenticationColor = "red"
            }
            $Script:AutoD_VD_InternalNblBypassUrl = $EWS.InternalNblBypassUrl
            $Script:AutoD_VD_InternalUrl = $EWS.InternalUrl
            $Script:AutoD_VD_ExternalUrl = $EWS.ExternalUrl
            AutoDVirtualDCheckOauthHtmlOk
        }
    } else {
        Write-Host -ForegroundColor Red "  InternalAuthenticationMethods seems not to include OAuth Authentication Method."
        $Script:AutoD_VD_Identity = $EWS.Identity
        $Script:AutoD_VD_Name = $EWS.Name
        $Script:AutoD_VD_InternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
        $Script:AutoD_VD_ExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
        $Script:AutoD_VD_WSAuthentication = $EWS.WSSecurityAuthentication
        $Script:AutoD_VD_WSAuthenticationColor = "green"
        $Script:AutoD_VD_OAuthAuthentication = $EWS.OAuthAuthentication
        $Script:AutoD_VD_WindowsAuthentication = $EWS.WindowsAuthentication
        if ($Script:AutoD_VD_WindowsAuthentication -eq "True") {
            $Script:AutoD_VD_WindowsAuthenticationColor = "green"
        } else {
            $Script:AutoD_VD_WindowsAuthenticationColor = "red"
        }
        $Script:AutoD_VD_InternalNblBypassUrl = $EWS.InternalNblBypassUrl
        $Script:AutoD_VD_InternalUrl = $EWS.InternalUrl
        $Script:AutoD_VD_ExternalUrl = $EWS.ExternalUrl
    }
    Write-Host -ForegroundColor White "`n  ExternalAuthenticationMethods"
    if ($AutoDiscoveryVirtualDirectoryOAuth.ExternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  ExternalAuthenticationMethods Include OAuth Authentication Method "
        }
    } else {
        Write-Host -ForegroundColor Red "  ExternalAuthenticationMethods seems not to include OAuth Authentication Method."
    }
    Write-Host -ForegroundColor White "`n  WSSecurityAuthentication:"
    if ($AutoDiscoveryVirtualDirectoryOAuth.WSSecurityAuthentication -like "True") {
        foreach ( $AdVd in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($AdVd.Identity) "
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($AdVd.WSSecurityAuthentication)"
        }
    } else {
        Write-Host -ForegroundColor Red "  WSSecurityAuthentication settings are NOT correct."
        foreach ( $AdVd in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($AdVd.Identity) "
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($AdVd.WSSecurityAuthentication).  WSSecurityAuthentication setting should be True."
        }
        Write-Host -ForegroundColor White "  Should be True "
    }
    Write-Host -ForegroundColor White "`n  WindowsAuthentication:"
    if ($AutoDiscoveryVirtualDirectoryOAuth.WindowsAuthentication -eq "True") {
        foreach ( $ser in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
    } else {
        Write-Host -ForegroundColor Red " WindowsAuthentication is NOT correct."
        foreach ( $ser in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
        Write-Host -ForegroundColor White "  Should be True "
    }
}
function EWSVirtualDirectoryCheckOAuth {
    Write-Host -ForegroundColor Green " Get-WebServicesVirtualDirectory  -Server $($server)| Select Identity,Name,ExchangeVersion,*Authentication*,*url"
    PrintDynamicWidthLine
    FetchEWSInformation
    $W = $WebServicesVirtualDirectoryOAuth | Format-List
    $W
    EWSVirtualDirectoryCheckOAuthHtmlHead
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - On-Prem Get-WebServicesVirtualDirectory"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White "  InternalAuthenticationMethods"
    if ($WebServicesVirtualDirectoryOAuth.InternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $WebServicesVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  InternalAuthenticationMethods Include OAuth Authentication Method "
            $Script:EwsVDIdentity = $EWS.Identity
            $Script:EwsVDName = $EWS.Name
            $Script:EwsVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $Script:EwsVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $Script:EwsVD_WSAuthentication = $EWS.WSSecurityAuthentication
            $Script:EwsVD_WSAuthenticationColor = "green"
            $Script:EwsVDWindowsAuthentication = $EWS.WindowsAuthentication
            $Script:EwsVDOAuthAuthentication = $EWS.OAuthAuthentication
            if ($Script:EwsVDWindowsAuthentication -eq "True") {
                $Script:EwsVDWindowsAuthenticationColor = "green"
            } else {
                $Script:EWS_VDWindowsAuthenticationColor = "red"
            }
            if ($EwsVDOAuthAuthentication -eq "True") {
                $Script:EwsVDW_OAuthAuthenticationColor = "green"
            } else {
                $Script:EWS_DOAuthAuthenticationColor = "red"
            }
            $Script:EwsVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
            $Script:EwsVDInternalUrl = $EWS.InternalUrl
            $Script:EwsVDExternalUrl = $EWS.ExternalUrl
            EWSVirtualDirectoryCheckOAuthHtmlOk
        }
    } else {
        Write-Host -ForegroundColor Red "  InternalAuthenticationMethods seems not to include OAuth Authentication Method."
        $Script:EwsVDIdentity = $EWS.Identity
        $Script:EwsVDName = $EWS.Name
        $Script:EwsVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
        $Script:EwsVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
        $Script:EwsVD_WSAuthentication = $EWS.WSSecurityAuthentication
        $Script:EwsVD_WSAuthenticationColor = "green"
        $Script:EwsVDWindowsAuthentication = $EWS.WindowsAuthentication
        $Script:EwsVDOAuthAuthentication = $EWS.OAuthAuthentication
        if ($Script:EwsVDWindowsAuthentication -eq "True") {
            $Script:EwsVDWindowsAuthenticationColor = "green"
        } else {
            $Script:EWS_DWindowsAuthenticationColor = "red"
        }
        if ($EwsVDOAuthAuthentication -eq "True") {
            $Script:EwsVDW_OAuthAuthenticationColor = "green"
        } else {
            $Script:EWS_DOAuthAuthenticationColor = "red"
        }
        $Script:EwsVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
        $Script:EwsVDInternalUrl = $EWS.InternalUrl
        $Script:EwsVDExternalUrl = $EWS.ExternalUrl
        EWSVirtualDirectoryCheckOAuthHtmlNotOk
    }
    Write-Host -ForegroundColor White "`n  ExternalAuthenticationMethods"
    if ($WebServicesVirtualDirectoryOAuth.ExternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $WebServicesVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  ExternalAuthenticationMethods Include OAuth Authentication Method "
        }
    } else {
        Write-Host -ForegroundColor Red "  ExternalAuthenticationMethods seems not to include OAuth Authentication Method."
    }
    Write-Host -ForegroundColor White "`n  WSSecurityAuthentication:"
    if ($WebServicesVirtualDirectoryOAuth.WSSecurityAuthentication -like "True") {
        foreach ( $EWS in $WebServicesVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) "
        }
    } else {
        Write-Host -ForegroundColor Red "  WSSecurityAuthentication is NOT correct."
        foreach ( $EWS in $WebServicesVirtualDirectoryOauth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication)"
        }
        Write-Host -ForegroundColor White "  Should be True"
    }
    #PrintDynamicWidthLine
    Write-Host -ForegroundColor White "`n  WindowsAuthentication:"
    if ($WebServicesVirtualDirectoryOauth.WindowsAuthentication -eq "True") {
        foreach ( $ser in $WebServicesVirtualDirectoryOauth) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
    } else {
        Write-Host -ForegroundColor Red " WindowsAuthentication is NOT correct."
        foreach ( $ser in $WebServicesVirtualDirectoryOauth) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
        Write-Host -ForegroundColor White "  Should be True "
    }
}

function ExoOrgRelCheck () {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-EOOrganizationRelationship  | Where{($_.DomainNames -like $Script:ExchangeOnPremDomain )} | Select Identity,DomainNames,FreeBusy*,TarGet*,Enabled"
    PrintDynamicWidthLine
    $ExoOrgRel
    PrintDynamicWidthLine
    Write-Host  -ForegroundColor Green " Summary - Organization Relationship"
    PrintDynamicWidthLine
    Write-Host  " Domain Names:"
    if ($exoOrgRel.DomainNames -like $Script:ExchangeOnPremDomain) {
        Write-Host -ForegroundColor Green "  Domain Names Include the $Script:ExchangeOnPremDomain Domain"
        $Script:tdExoOrgRelDomainNames = $exoOrgRel.DomainNames
        $Script:tdExoOrgRelDomainNamesColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  Domain Names do Not Include the $Script:ExchangeOnPremDomain Domain"
        $exoOrgRel.DomainNames
        $Script:tdExoOrgRelDomainNames = "$($exoOrgRel.DomainNames) - Domain Names do Not Include the $Script:ExchangeOnPremDomain Domain"
        $Script:tdExoOrgRelDomainNamesColor = "green"
    }
    #FreeBusyAccessEnabled
    Write-Host  " FreeBusyAccessEnabled:"
    if ($exoOrgRel.FreeBusyAccessEnabled -like "True" ) {
        Write-Host -ForegroundColor Green "  FreeBusyAccessEnabled is set to True"
        $Script:tdExoOrgRelFreeBusyAccessEnabled = "$($exoOrgRel.FreeBusyAccessEnabled)"
        $Script:tdExoOrgRelFreeBusyAccessEnabledColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  FreeBusyAccessEnabled : False"
        $Script:tdExoOrgRelFreeBusyAccessEnabled = "$($exoOrgRel.FreeBusyAccessEnabled). Free busy access is not enabled for the organization Relationship"
        $Script:tdExoOrgRelFreeBusyAccessEnabledColor = "Red"
    }
    #FreeBusyAccessLevel
    Write-Host  " FreeBusyAccessLevel:"
    if ($exoOrgRel.FreeBusyAccessLevel -like "AvailabilityOnly" ) {
        Write-Host -ForegroundColor Green "  FreeBusyAccessLevel is set to AvailabilityOnly"
        $Script:tdExoOrgRelFreeBusyAccessLevel = "$($exoOrgRel.FreeBusyAccessLevel)"
        $Script:tdExoOrgRelFreeBusyAccessLevelColor = "green"
    }
    if ($exoOrgRel.FreeBusyAccessLevel -like "LimitedDetails" ) {
        Write-Host -ForegroundColor Green "  FreeBusyAccessLevel is set to LimitedDetails"
        $Script:tdExoOrgRelFreeBusyAccessLevel = "$($exoOrgRel.FreeBusyAccessLevel)"
        $Script:tdExoOrgRelFreeBusyAccessLevelColor = "green"
    }
    if ($exoOrgRel.FreeBusyAccessLevel -ne "AvailabilityOnly" -and $exoOrgRel.FreeBusyAccessLevel -ne "LimitedDetails") {
        Write-Host -ForegroundColor Red "  FreeBusyAccessEnabled : False"
        $Script:tdExoOrgRelFreeBusyAccessLevel = "$($exoOrgRel.FreeBusyAccessLevel)"
        $Script:tdExoOrgRelFreeBusyAccessLevelColor = "red"
    }
    #TarGetApplicationUri
    Write-Host  " TarGetApplicationUri:"
    $a = "FYDIBOHF25SPDLT." + $Script:ExchangeOnPremDomain
    $HybridAgentTargetSharingEpr = "http://outlook.office.com/"
    $HATargetAutodiscoverEpr = "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/"
    if ($exoOrgRel.TarGetSharingEpr -like "*resource.mailboxMigration.his.MSAppProxy.net/EWS/Exchange.asmx") {
        if ($exoOrgRel.TarGetApplicationUri -like $HybridAgentTargetSharingEpr) {
            Write-Host -ForegroundColor Green "  TarGetApplicationUri is $($exoOrgRel.TarGetSharingEpr) . This is correct when Hybrid Agent is in use"
            $Script:tdExoOrgRelTarGetApplicationUri = "  TarGetApplicationUri is $($exoOrgRel.TarGetSharingEpr) . This is correct when Hybrid Agent is in use"
            $Script:tdExoOrgRelTarGetApplicationUriColor = "green"
        } else {
            Write-Host -ForegroundColor Red "  TarGetApplicationUri should be  $HybridAgentTargetSharingEpr when Hybrid Agent is used"
            $Script:tdExoOrgRelTarGetApplicationUri = "  TarGetApplicationUri should be $HybridAgentTargetSharingEpr when Hybrid Agent is used. Please Check if Exchange On Premise Federation is correctly configured."
            $Script:tdExoOrgRelTarGetApplicationUriColor = "red"
        }
    } else {
        if ($exoOrgRel.TarGetApplicationUri -like $FedTrust.ApplicationUri) {
            Write-Host -ForegroundColor Green "  TarGetApplicationUri is" $FedTrust.ApplicationUri.OriginalString
            $Script:tdExoOrgRelTarGetApplicationUri = "  TarGetApplicationUri is $($FedTrust.ApplicationUri.OriginalString)"
            $Script:tdExoOrgRelTarGetApplicationUriColor = "green"
        } else {
            Write-Host -ForegroundColor Red "  TarGetApplicationUri should be " $a
            $Script:tdExoOrgRelTarGetApplicationUri = "  TarGetApplicationUri should be $a. Please Check if Exchange On Premise Federation is correctly configured."
            $Script:tdExoOrgRelTarGetApplicationUriColor = "red"
        }
    }
    #TarGetSharingEpr
    Write-Host  " TarGetSharingEpr:"
    if ($exoOrgRel.TarGetSharingEpr -like "*resource.mailboxMigration.his.MsAppProxy.net/EWS/Exchange.asmx") {
        Write-Host -ForegroundColor Green "  TarGetSharingEpr is points to resource.mailboxMigration.his.MsAppProxy.net/EWS/Exchange.asmx. This means Hybrid Agent is in use."
        $Script:tdExoOrgRelTarGetSharingEpr = "TarGetSharingEpr is points to resource.mailboxMigration.his.MsAppProxy.net/EWS/Exchange.asmx. This means Hybrid Agent is in use."
        $Script:tdExoOrgRelTarGetSharingEprColor = "green"
    } else {
        if ([string]::IsNullOrWhitespace($exoOrgRel.TarGetSharingEpr)) {
            Write-Host -ForegroundColor Green "  TarGetSharingEpr is blank. This is the standard Value."
            $Script:tdExoOrgRelTarGetSharingEpr = "TarGetSharingEpr is blank. This is the standard Value."
            $Script:tdExoOrgRelTarGetSharingEprColor = "green"
        } else {
            Write-Host -ForegroundColor Red "  TarGetSharingEpr should be blank. If it is set, it should be the On-Premises Exchange Servers EWS ExternalUrl endpoint."
            $Script:tdExoOrgRelTarGetSharingEpr = "  TarGetSharingEpr should be blank. If it is set, it should be the On-Premises Exchange Servers EWS ExternalUrl endpoint."
            $Script:tdExoOrgRelTarGetSharingEprColor = "red"
        }
    }
    Write-Host  " TarGetAutoDiscoverEpr:"
    if ($exoOrgRel.TarGetSharingEpr -like "*resource.mailboxMigration.his.MSAppProxy.net/EWS/Exchange.asmx") {

        if ($exoOrgRel.TarGetAutoDiscoverEpr -like $HATargetAutodiscoverEpr) {
            Write-Host -ForegroundColor Green "  TarGetAutoDiscoverEpr is $($exoOrgRel.TarGetAutoDiscoverEpr) . This is correct when Hybrid Agent is in use"

            $Script:tdExoOrgRelTarGetAutoDiscoverEpr = "TarGetAutoDiscoverEpr is $($exoOrgRel.TarGetAutoDiscoverEpr) . This is correct when Hybrid Agent is in use"
            $Script:tdExoOrgRelTarGetAutoDiscoverEprColor = "green"
        } else {
            Write-Host -ForegroundColor Red "  TarGetAutoDiscoverEpr is not $HATargetAutodiscoverEpr . This is the correct  value when Hybrid Agent is in use."
            $Script:tdExoOrgRelTarGetAutoDiscoverEpr = "  TarGetAutoDiscoverEpr is not $HATargetAutodiscoverEpr. This is the correct  value when Hybrid Agent is in use."
            $Script:tdExoOrgRelTarGetAutoDiscoverEprColor = "red"
        }
    }

    else {

        if ($exoOrgRel.TarGetAutoDiscoverEpr -like $FedInfoEOP.TarGetAutoDiscoverEpr) {
            Write-Host -ForegroundColor Green "  TarGetAutoDiscoverEpr is" $exoOrgRel.TarGetAutoDiscoverEpr
            $Script:tdExoOrgRelTarGetAutoDiscoverEpr = $exoOrgRel.TarGetAutoDiscoverEpr
            $Script:tdExoOrgRelTarGetAutoDiscoverEprColor = "green"
        } else {
            Write-Host -ForegroundColor Red "  TarGetAutoDiscoverEpr is not" $FedInfoEOP.TarGetAutoDiscoverEpr
            $Script:tdExoOrgRelTarGetAutoDiscoverEpr = "  TarGetAutoDiscoverEpr is not $($FedInfoEOP.TarGetAutoDiscoverEpr)"
            $Script:tdExoOrgRelTarGetAutoDiscoverEprColor = "red"
        }
    }
    #Enabled
    Write-Host  " Enabled:"
    if ($exoOrgRel.enabled -like "True" ) {
        Write-Host -ForegroundColor Green "  Enabled is set to True"
        $Script:tdExoOrgRelEnabled = "  True"
        $Script:tdExoOrgRelEnabledColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  Enabled is set to False."
        $Script:tdExoOrgRelEnabled = "  False"
        $Script:tdExoOrgRelEnabledColor = "red"
    }
    ExoOrgRelCheckHtml
}
function ExoFedOrgIdCheck {
    Write-Host -ForegroundColor Green " Get-FederatedOrganizationIdentifier | select AccountNameSpace,Domains,Enabled"
    PrintDynamicWidthLine
    $ExoFedOrgId = Get-EOFederatedOrganizationIdentifier | Select-Object AccountNameSpace, Domains, Enabled
    $eFedOrgID = $ExoFedOrgId | Format-List
    $eFedOrgID
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Online Federated Organization Identifier"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " Domains: "
    if ($ExoFedOrgId.Domains -like "*$Script:ExchangeOnlineDomain*") {
        Write-Host -ForegroundColor Green " " $ExoFedOrgId.Domains
        $Script:tdExoFedOrgIdDomains = $ExoFedOrgId.Domains
        $Script:tdExoFedOrgIdDomainsColor = "green"
    } else {
        Write-Host -ForegroundColor Red " Domains are NOT correct."
        Write-Host -ForegroundColor White " Should contain the $ExchangeOnlineMDomain"
        $Script:tdExoFedOrgIdDomains = "$($ExoFedOrgId.Domains) . Domains Should contain the $ExchangeOnlineMDomain"
        $Script:tdExoFedOrgIdDomainsColor = "red"
    }
    Write-Host -ForegroundColor White " Enabled: "
    if ($ExoFedOrgId.Enabled -like "True") {
        Write-Host -ForegroundColor Green "  True "
        $Script:tdExoFedOrgIdEnabled = $ExoFedOrgId.Enabled
        $Script:tdExoFedOrgIdEnabledColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  Enabled is NOT correct."
        Write-Host -ForegroundColor White " Should be True"
        $Script:tdExoFedOrgIdEnabled = $ExoFedOrgId.Enabled
        $Script:tdExoFedOrgIdEnabledColor = "green"
    }
    ExoFedOrgIdCheckHtml
}
function SharingPolicyCheck {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-SharingPolicy | select Domains,Enabled,Name,Identity"
    PrintDynamicWidthLine
    $Script:SPOnline = Get-EOSharingPolicy | Select-Object  Domains, Enabled, Name, Identity
    $SPOnline | Format-List
    $domain1 = (($SPOnline.domains[0] -split ":") -split " ")
    $domain2 = (($SPOnline.domains[1] -split ":") -split " ")
    $Script:SPOnpremDomain1 = $Script:SPOnprem.Domains.Domain[0]
    $Script:SPOnpremAction1 = $Script:SPOnprem.Domains.Actions[0]
    $Script:SPOnpremDomain2 = $Script:SPOnprem.Domains.Domain[1]
    $Script:SPOnpremAction2 = $Script:SPOnprem.Domains.Actions[1]
    $SPOnlineDomain1 = $domain1[0]
    $SPOnlineAction1 = $domain1[1]
    $SPOnlineDomain2 = $domain2[0]
    $SPOnlineAction2 = $domain2[1]
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Sharing Policy"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " Exchange On Premises Sharing domains:`n"
    Write-Host -ForegroundColor White "  Domain:"
    Write-Host "   " $Script:SPOnpremDomain1
    Write-Host -ForegroundColor White "  Action:"
    Write-Host "   " $Script:SPOnpremAction1
    Write-Host -ForegroundColor White "  Domain:"
    Write-Host "   " $Script:SPOnpremDomain2
    Write-Host -ForegroundColor White "  Action:"
    Write-Host "   " $Script:SPOnpremAction2
    Write-Host -ForegroundColor White "`n  Exchange Online Sharing Domains: `n"
    Write-Host -ForegroundColor White "  Domain:"
    Write-Host "   " $SPOnlineDomain1
    Write-Host -ForegroundColor White "  Action:"
    Write-Host "   " $SPOnlineAction1
    Write-Host -ForegroundColor White "  Domain:"
    Write-Host "   " $SPOnlineDomain2
    Write-Host -ForegroundColor White "  Action:"
    Write-Host "   " $SPOnlineAction2

    if ($Script:SPOnpremDomain1 -eq $SPOnlineDomain1 -and $Script:SPOnpremAction1 -eq $SPOnlineAction1) {
        if ($Script:SPOnpremDomain2 -eq $SPOnlineDomain2 -and $Script:SPOnpremAction2 -eq $SPOnlineAction2) {
            Write-Host -ForegroundColor Green "`n  Exchange Online Sharing Policy Domains match Exchange On Premise Sharing Policy Domains"
            $Script:tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy matches Exchange On Premise Sharing Policy Domain"
            $Script:tdSharpingPolicyCheckColor = "green"
        }

        else {
            Write-Host -ForegroundColor Red "`n   Sharing Domains appear not to be correct."
            Write-Host -ForegroundColor White "   Exchange Online Sharing Policy Domains appear not to match Exchange On Premise Sharing Policy Domains"
            $Script:tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy Domains not match Exchange On Premise Sharing Policy Domains"
            $Script:tdSharpingPolicyCheckColor = "red"
        }
    } elseif ($Script:SPOnpremDomain1 -eq $SPOnlineDomain2 -and $Script:SPOnpremAction1 -eq $SPOnlineAction2) {
        if ($Script:SPOnpremDomain2 -eq $SPOnlineDomain1 -and $Script:SPOnpremAction2 -eq $SPOnlineAction1) {
            Write-Host -ForegroundColor Green "`n  Exchange Online Sharing Policy Domains match Exchange On Premise Sharing Policy Domains"
            $Script:tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy matches Exchange On Premise Sharing Policy Domain"
            $Script:tdSharpingPolicyCheckColor = "green"
        }

        else {
            Write-Host -ForegroundColor Red "`n   Sharing Domains appear not to be correct."
            Write-Host -ForegroundColor White "   Exchange Online Sharing Policy Domains appear not to match Exchange On Premise Sharing Policy Domains"
            $Script:tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy Domains not match Exchange On Premise Sharing Policy Domains"
            $Script:tdSharpingPolicyCheckColor = "red"
        }
    } else {
        Write-Host -ForegroundColor Red "`n   Sharing Domains appear not to be correct."
        Write-Host -ForegroundColor White "   Exchange Online Sharing Policy Domains appear not to match Exchange On Premise Sharing Policy Domains"
        $Script:tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy Domains not match Exchange On Premise Sharing Policy Domains"
        $Script:tdSharpingPolicyCheckColor = "red"
    }
    PrintDynamicWidthLine
    SharingPolicyCheckHtml
}
function ExoTestOrgRelCheck {
    $exoIdentity = $ExoOrgRel.Identity
    $exoOrgRelTarGetApplicationUri = $exoOrgRel.TarGetApplicationUri
    $exoOrgRelTarGetOWAUrl = $ExoOrgRel.TarGetOwAUrl
    Write-Host -ForegroundColor Green " Test-OrganizationRelationship -Identity $exoIdentity -UserIdentity $Script:UserOnline"
    PrintDynamicWidthLine
    if ((![string]::IsNullOrWhitespace($exoOrgRelTarGetApplicationUri)) -and (![string]::IsNullOrWhitespace($exoOrgRelTarGetOWAUrl))) {
        $ExoTestOrgRel = Test-EOOrganizationRelationship -Identity $exoIdentity -UserIdentity $Script:UserOnline -WarningAction SilentlyContinue
        $i = 2
        while ($i -lt $ExoTestOrgRel.Length) {
            $element = $ExoTestOrgRel[$i]
            $aux = "0"
            if ($element -like "*RESULT:*" -and $aux -like "0") {
                $el = $element.TrimStart()
                if ($element -like "*Success.*") {
                    Write-Host -ForegroundColor Green "  $el"
                    $aux = "1"
                } elseif ($element -like "*Error*" -or $element -like "*Unable*") {
                    Write-Host -ForegroundColor Red "  $el"
                    $aux = "1"
                }
            } elseif ($aux -like "0" ) {
                if ($element -like "*STEP*" -or $element -like "*Complete*") {
                    Write-Host -ForegroundColor White "  $element"
                    $aux = "1"
                } else {
                    $ID = $element.ID
                    $Status = $element.Status
                    $Description = $element.Description
                    if (![string]::IsNullOrWhitespace($ID)) {
                        Write-Host -ForegroundColor White "`n  ID         : $ID"
                        if ($Status -like "*Success*") {
                            Write-Host -ForegroundColor White "  Status     : $Status"
                        }
                        if ($status -like "*error*") {
                            Write-Host -ForegroundColor White "  Status     : $Status"
                        }
                        Write-Host -ForegroundColor White "  Description: $Description"
                        Write-Host -ForegroundColor yellow "  Note: Test-Organization Relationship fails on Step 3 with error MismatchedFederation if Hybrid Agent is in use"
                    }
                    #$element
                    $aux = "1"
                }
            }
            $i++
        }
    }

    elseif ((([string]::IsNullOrWhitespace($exoOrgRelTarGetApplicationUri)) -and ([string]::IsNullOrWhitespace($exoOrgRelTarGetOWAUrl)))) {
        Write-Host -ForegroundColor Red "  Error: Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TarGetApplicationUri and TarGetOwAUrl are not set"
    } elseif ((([string]::IsNullOrWhitespace($exoOrgRelTarGetApplicationUri)) )) {
        Write-Host -ForegroundColor Red "  Error: Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TarGetApplicationUri is not set"
    } elseif ((([string]::IsNullOrWhitespace($exoOrgRelTarGetApplicationUri)) )) {
        Write-Host -ForegroundColor Red "  Error: Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TarGetOwAUrl is not set"
    }
    ExoTestOrgRelCheckHtml
}


function EXOIntraOrgConCheck {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Get-IntraOrganizationConnector | Select TarGetAddressDomains,DiscoveryEndpoint,Enabled"
    PrintDynamicWidthLine
    $Script:ExoIntraOrgCon = Get-EOIntraOrganizationConnector | Select-Object TarGetAddressDomains, DiscoveryEndpoint, Enabled
    $IOC = $Script:ExoIntraOrgCon | Format-List
    $IOC
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Online Intra Organization Connector"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " TarGet Address Domains: "
    if ($Script:ExoIntraOrgCon.TarGetAddressDomains -like "*$Script:ExchangeOnPremDomain*") {
        Write-Host -ForegroundColor Green " " $Script:ExoIntraOrgCon.TarGetAddressDomains
        $Script:tdExoIntraOrgConTarGetAddressDomains = $Script:ExoIntraOrgCon.TarGetAddressDomains
        $Script:tdExoIntraOrgConTarGetAddressDomainsColor = "green"
    } else {
        Write-Host -ForegroundColor Red " TarGet Address Domains is NOT correct."
        Write-Host -ForegroundColor White " Should contain the $Script:ExchangeOnPremDomain"
        $Script:tdExoIntraOrgConTarGetAddressDomains = " $($Script:ExoIntraOrgCon.TarGetAddressDomains) . Should contain the $Script:ExchangeOnPremDomain"
        $Script:tdExoIntraOrgConTarGetAddressDomainsColor = "red"
    }
    Write-Host -ForegroundColor White " DiscoveryEndpoint: "
    if ($Script:ExoIntraOrgCon.DiscoveryEndpoint -like $Script:EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint) {
        Write-Host -ForegroundColor Green $Script:ExoIntraOrgCon.DiscoveryEndpoint
        $Script:tdExoIntraOrgConDiscoveryEndpoints = $Script:ExoIntraOrgCon.DiscoveryEndpoint
        $Script:tdExoIntraOrgConDiscoveryEndpointsColor = "green"
    } else {
        if ($Script:ExoIntraOrgCon.DiscoveryEndpoint -like "*resource.mailboxMigration.his.MSAppProxy.net*") {
            Write-Host -ForegroundColor Green " " $Script:ExoIntraOrgCon.DiscoveryEndpoint
            Write-Host -ForegroundColor Yellow " Discovery Endpoint includes resource.mailboxMigration.his.MSAppProxy.net. Hybrid configuration is implemented using Hybrid Agent "
            $Script:tdExoIntraOrgConDiscoveryEndpoints = $Script:ExoIntraOrgCon.DiscoveryEndpoint
            $Script:tdExoIntraOrgConDiscoveryEndpointsColor = "green"
        } else {
            Write-Host -ForegroundColor Red " DiscoveryEndpoint is NOT correct. "
            Write-Host -ForegroundColor White "  Should be " $Script:EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint
            $Script:tdExoIntraOrgConDiscoveryEndpoints = "$($Script:ExoIntraOrgCon.DiscoveryEndpoint) . Should be $($Script:EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint)"
            $Script:tdExoIntraOrgConDiscoveryEndpointsColor = "red"
        }
    }
    Write-Host -ForegroundColor White " Enabled: "
    if ($Script:ExoIntraOrgCon.Enabled -like "True") {
        Write-Host -ForegroundColor Green "  True "
        $Script:tdExoIntraOrgConEnabled = "True"
        $Script:tdExoIntraOrgConEnabledColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  False."
        Write-Host -ForegroundColor White " Should be True"
        $Script:tdExoIntraOrgConEnabled = "False . Should be True"
        $Script:tdExoIntraOrgConEnabledColor = "red"
    }
    EXOIntraOrgConCheckHtml
}
function EXOIntraOrgConfigCheck {
    Write-Host -ForegroundColor Green " Get-IntraOrganizationConfiguration | Select OnPremiseTarGetAddresses"
    PrintDynamicWidthLine
    $Script:ExoIntraOrgConfig = Get-EOOnPremisesOrganization | Select-Object OrganizationGuid | Get-EOIntraOrganizationConfiguration | Select-Object * | Where-Object { $_.OnPremiseTarGetAddresses -like "*$Script:ExchangeOnPremDomain*" }
    $IOConfig = $Script:ExoIntraOrgConfig | Format-List
    $IOConfig
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Exchange Online Intra Organization Configuration"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " OnPremiseTarGetAddresses: "
    if ($Script:ExoIntraOrgConfig.OnPremiseTarGetAddresses -like "*$Script:ExchangeOnPremDomain*") {
        Write-Host -ForegroundColor Green " " $Script:ExoIntraOrgConfig.OnPremiseTarGetAddresses
        $Script:tdExoIntraOrgConfigOnPremiseTarGetAddresses = $Script:ExoIntraOrgConfig.OnPremiseTarGetAddresses
        $Script:tdExoIntraOrgConfigOnPremiseTarGetAddressesColor = "green"
    } else {
        Write-Host -ForegroundColor Red " OnPremise TarGet Addresses are NOT correct."
        Write-Host -ForegroundColor White " Should contain the $Script:ExchangeOnPremDomain"
        $Script:tdExoIntraOrgConfigOnPremiseTarGetAddresses = $Script:ExoIntraOrgConfig.OnPremiseTarGetAddresses
        $Script:tdExoIntraOrgConfigOnPremiseTarGetAddressesColor = "red"
    }
    EXOIntraOrgConfigCheckHtml
}
function EXOAuthServerCheck {
    Write-Host -ForegroundColor Green " Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | select name,IssuerIdentifier,enabled"
    PrintDynamicWidthLine
    $exoAuthServer = Get-EOAuthServer -Identity 00000001-0000-0000-c000-000000000000 | Select-Object name, IssuerIdentifier, enabled
    $AuthServer = $exoAuthServer | Format-List
    $AuthServer
    $Script:tdExoAuthServerName = $exoAuthServer.Name
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - Exchange Online Authorization Server"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " IssuerIdentifier: "
    if ($exoAuthServer.IssuerIdentifier -like "00000001-0000-0000-c000-000000000000") {
        Write-Host -ForegroundColor Green " " $exoAuthServer.IssuerIdentifier
        $Script:tdExoAuthServerIssuerIdentifier = $exoAuthServer.IssuerIdentifier
        $Script:tdExoAuthServerIssuerIdentifierColor = "green"
        if ($exoAuthServer.Enabled -like "True") {
            Write-Host -ForegroundColor Green "  True "
            $Script:tdExoAuthServerEnabled = $exoAuthServer.Enabled
            $Script:tdExoAuthServerEnabledColor = "green"
        } else {
            Write-Host -ForegroundColor Red "  Enabled is NOT correct."
            Write-Host -ForegroundColor White " Should be True"
            $Script:tdExoAuthServerEnabled = "$($exoAuthServer.Enabled) . Should be True"
            $Script:tdExoAuthServerEnabledColor = "red"
        }
    } else {
        Write-Host -ForegroundColor Red " Authorization Server object is NOT correct."
        Write-Host -ForegroundColor White " Enabled: "
        $Script:tdExoAuthServerIssuerIdentifier = "$($exoAuthServer.IssuerIdentifier) - Authorization Server object should be 00000001-0000-0000-c000-000000000000"
        $Script:tdExoAuthServerIssuerIdentifierColor = "red"

        if ($exoAuthServer.Enabled -like "True") {
            Write-Host -ForegroundColor Green "  True "
            $Script:tdExoAuthServerEnabled = $exoAuthServer.Enabled
            $Script:tdExoAuthServerEnabledColor = "green"
        } else {
            Write-Host -ForegroundColor Red "  Enabled is NOT correct."
            Write-Host -ForegroundColor White " Should be True"
            $Script:tdExoAuthServerEnabled = "$($exoAuthServer.Enabled) . Should be True"
            $Script:tdExoAuthServerEnabledColor = "red"
        }
    }
    EXOAuthServerCheckHtml
}
function ExoTestOAuthCheck {
    Write-Host -ForegroundColor Green " Test-OAuthConnectivity -Service EWS -TarGetUri $Script:ExchangeOnPremEWS -Mailbox $Script:UserOnline "
    PrintDynamicWidthLine
    $ExoTestOAuth = Test-EOOAuthConnectivity -Service EWS -TarGetUri $Script:ExchangeOnPremEWS -Mailbox $Script:UserOnline -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if ($ExoTestOAuth.ResultType.Value -like 'Success' ) {
        $Script:tdOAuthConnectivityResultType = "$($ExoTestOAuth.ResultType.Value) - OAuth Test was completed successfully"
        $Script:tdOAuthConnectivityResultTypeColor = "green"
    } else {
        $ExoTestOAuth | Format-List
        $Script:tdOAuthConnectivityResultType = "$($ExoTestOAuth.ResultType) - OAuth Test was completed with Error. Please rerun Test-OAuthConnectivity -Service EWS -TarGetUri <EWS target URI> -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        $Script:tdOAuthConnectivityResultTypeColor = "red"
    }

    if ($ExoTestOAuth.Detail.FullId -like '*(401) Unauthorized*') {
        Write-Host -ForegroundColor Red "The remote Server returned an error: (401) Unauthorized"
    }

    if ($ExoTestOAuth.Detail.FullId -like '*The user specified by the user-context in the token does not exist*') {
        Write-Host -ForegroundColor Yellow "The user specified by the user-context in the token does not exist"
        Write-Host "Please run Test-OAuthConnectivity with a different Exchange Online Mailbox"
    }

    if ($ExoTestOAuth.Detail.FullId -like '*error_category="invalid_token"*') {
        Write-Host -ForegroundColor Yellow "This token profile 'S2SAppActAs' is not applicable for the current protocol"
    }
    Write-Host -ForegroundColor Green " Summary - Test-OAuthConnectivity"
    PrintDynamicWidthLine
    if ($ExoTestOAuth.ResultType.value -like "Success") {
        Write-Host -ForegroundColor Green " OAuth Test was completed successfully "
        $Script:tdOAuthConnectivityResultType = "  OAuth Test was completed successfully"
        $Script:tdOAuthConnectivityResultTypeColor = "green"
    } else {
        Write-Host -ForegroundColor Red " OAuth Test was completed with Error. "
        Write-Host -ForegroundColor White " Please rerun Test-OAuthConnectivity -Service EWS -TarGetUri <EWS tarGet URI> -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        $Script:tdOAuthConnectivityResultType = "$($ExoTestOAuth.ResultType) - OAuth Test was completed with Error. Please rerun Test-OAuthConnectivity -Service EWS -TarGetUri <EWS tarGet URI> -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        $Script:tdOAuthConnectivityResultTypeColor = "red"
    }
    Write-Host -ForegroundColor Green "`n References: "
    Write-Host -ForegroundColor White " Configure OAuth authentication between Exchange and Exchange Online organizations"
    Write-Host -ForegroundColor Yellow " https://technet.microsoft.com/en-us/library/dn594521(v=exchg.150).aspx"
    ExoTestOAuthCheckHtml
}
#On Prem DAuth
function showParametersHtml() {
    $script:html = "<!DOCTYPE html>
<!DOCTYPE html>
<html>
<head>
<title>Hybrid Free Busy Configuration Checker</title>
<style>
body {
  font-family: Arial;
  background-color: white;
}
table, th {
max-width: 95%;
margin-left: 2%;
margin-right: 2%;
  border: 1px solid black;
  border-collapse: collapse;
  padding: 5px;
  font-family: Courier;
  background-color: white;
  table-layout: fixed;
   font-family: Arial;
}
td {
  border: 1px solid black;
  border-collapse: collapse;
  padding: 5px;
  font-family: Arial;
  background-color: white;
  width: 50%;
  max-width: 50%;
  word-wrap: break-word;
}
th {
  background-color: blue;
  text-align: left;
   font-family: Arial;
}
.green { color: green; }
.red { color: red; }
.yellow { color: yellow; }
.white { color: white; }
.black { color: black; }
.orange { color: orange; }
.Black {
  font-weight: 500;
}
 p {
  font-weight : 548;
}
h1 {
  color: #00a2ed;
  padding-left: 2%;
}
h2 {
  color: #00a2ed;
  padding-left: 2%;
}
h3 {
  color: #00a2ed;
  padding-left: 2%;
}
ul {
  padding-left: 8%;
}

.microsoft {
background-color: #f25022;
box-shadow:
    28px 0 0 0 #7fba00,
    0 28px 0 0 #00a4ef,
    28px 28px 0 0 #ffb900;
height: 25px;
width: 25px;
margin-top: 1%;
margin-right: 1%;
}
</style>
</head>
<body>
        <div class='Black' style='display: -webkit-box;margin-left: 2%;'>
        <div class='microsoft'></div>
        <h1 style='padding-left: 2%;'>Microsoft CSS - Exchange Hybrid Free Busy Configuration Checker</h1></div>


     <div class='Black' style = 'padding-left: 0%;'>
          <h2><b>Parameters:</b></h2>
          <ul>
            <li>
              <p>Log File Path:</p>
              <span style='color:green; font-weight:500; padding-left:2%;'>$LogFile</span>
            </li>
            <li>
              <p>Office 365 Domain:</p>
              <span style='color:green; font-weight:500; padding-left:2%'>$Script:ExchangeOnlineDomain</span>
            </li>
            <li>
              <p>AD root Domain:</p>
              <span style='color:green; font-weight:500; padding-left:2%'>$Script:ExchangeOnPremLocalDomain</span>
            </li>
            <li>
              <p>Exchange On Premises Domain:</p>
              <span style='color:green; font-weight:500; padding-left:2%'>$Script:ExchangeOnPremDomain</span>
            </li>
            <li>
              <p>Exchange On Premises External EWS url:</p>
              <span style='color:green; font-weight:500; padding-left:2%'>$Script:ExchangeOnPremEWS</span>
            </li>
            <li>
              <p>On Premises Hybrid Mailbox:</p>
              <span style='color:green; font-weight:500; padding-left:2%'>$Script:UserOnPrem</span>
            </li>
            <li>
              <p>Exchange Online Mailbox:</p>
              <span style='color:green; font-weight:500; padding-left:2%'>$Script:UserOnline</span>
            </li>
          </ul>
        </div>
        <div class='Black'  style = 'padding-left: 0%;'><h2>Configuration:</h2></div>
        <p style='margin-left:2%;'>TLS 1.2 should be Enabled in order for Hybrid Free Busy to work. To confirm TLS Settings please Run the HealthChecker Script</p>
          <ul>
            <li><a href='https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/'>Microsoft Exchange Health Checker Script</a></li>
          </ul>
        <h3>Useful Links</h3>
          <ul>
            <li><a href='https://techcommunity.microsoft.com/t5/exchange-team-blog/demystifying-hybrid-free-busy-finding-errors-and-troubleshooting/ba-p/607727'>Demystifying Hybrid Free Busy: Finding Errors and Troubleshooting</a></li>
            <li><a href='https://support.microsoft.com/en-us/topic/how-to-troubleshoot-free-busy-issues-in-a-hybrid-deployment-of-on-premises-exchange-Server-and-exchange-online-in-office-365-ae03e199-b439-a84f-8db6-11bc0d7fbdf0'>How to Troubleshoot Free Busy Issues in a Hybrid Deployment of On-Premises Exchange Server and Exchange Online in Office 365</a></li>
            <li><a href='https://techcommunity.microsoft.com/t5/exchange-team-blog/the-hybrid-mesh/ba-p/605910'>The Hybrid Mesh</a></li>
            <li><a href='https://techcommunity.microsoft.com/t5/exchange-team-blog/how-to-address-federation-trust-issues-in-hybrid-configuration/ba-p/1144285'>How to Address Federation Trust Issues in Hybrid Configuration</a></li>
            <li><a href='https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?redirectSourcePath=%252farticle%252fOffice-365-URLs-and-IP-address-ranges-8548a211-3fe7-47cb-abb1-355ea5aa88a2&view=o365-worldwide'>Office 365 URLs and IP Address Ranges</a></li>
            <li><a href='https://techcommunity.microsoft.com/legacyfs/online/media/2019/01/FB_Errors.FixesV6.pdf'>Free Busy Errors and Fixes</a></li>
          </ul>
          "
    $html | Out-File -FilePath $Script:htmlFile
}
function orgRelHtml() {
    $script:html += @"
 <div class='Black'><p></p></div>
 <div class='Black'><h2><b>`n Exchange On Premise Free Busy Configuration: `n</b></h2></div>
 <div class='Black'><p></p></div>
<table style='width:100%'>
<tr>
<th ColSpan='2' style='text-align:center; color:white;'><b>Exchange On Premise DAuth Configuration</b></th>
</tr>
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-OrganizationRelationship</th>
</tr>
<tr>
<td><b>Get-OrganizationRelationship</b></td>
<td>
"@
    foreach ($setting in $settingsList) {
        $color = $setting.Color
        if ($color -eq 'yellow') {
            $color = 'red'
        }
        $script:html += @"
    <div> <b>$($setting.Name): </b> <span style='color: $($color)'>$($setting.Value)</span></div>
"@
    }
    $script:html += @"
</td>
</tr>
"@
    $script:html | Out-File -FilePath $Script:htmlFile
}
function FedInfoHtml() {
    $script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-FederationInformation</th>
</tr>
<tr>
<td><b>Get-FederationInformation -Domain $Script:ExchangeOnPremDomain</b></td>
<td>
  <div> <b>Domain Names: </b> <span style='color:$tdDomainNamesColor'>$tdDomainNamesFL</span></div>
  <div> <b>TokenIssuerUris: </b> <span style='color:$tdTokenIssuerUrisColor'>$tdTokenIssuerUrisFL</span></div>
  <div> <b>TarGetApplicationUri: </b> <span style='color:$tdTarGetApplicationUriColor'>$tdTarGetApplicationUriFL</span></div>
  <div> <b>TarGetAutoDiscoverEpr: </b> <span style='color:$tdTarGetAutoDiscoverEprColor'>$tdTarGetAutoDiscoverEprFL</span></div>
  <div> <b>TarGetApplicationUri - Federation Information vs Organization Relationship: </b> <span style='color:$tdTarGetAutoDiscoverEprVSColor'>$tdFederationInformationTA_FL</span></div>
  <div> <b>TarGetAutoDiscoverEpr - Federation Information vs Organization Relationship:</b> <span style='color:$tdTarGetAutoDiscoverEprVSColor'>$tdTarGetAutoDiscoverEprVS_FL</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function fedTrustHtml() {
    $script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Test-FederationTrust</th>
</tr>
<tr>
<td><b>Get-FederationTrust | select ApplicationUri, TokenIssuerUri, OrgCertificate, TokenIssuerCertificate, TokenIssuerPrevCertificate, TokenIssuerMetadataEpr, TokenIssuerEpr</b></td>
<td>
<div> <b>Application Uri: </b> <span style='color:$tdFedTrustApplicationUriColor'>$tdFedTrustApplicationUriFL</span></div>
<div> <b>TokenIssuerUris: </b> <span style='color:$tdFedTrustTokenIssuerUriColor'>$tdFedTrustTokenIssuerUriFL</span></div>
<div> <b>Certificate Expiry: </b> <span style='color:$tdFedTrustOrgCertificateNotAfterDateColor'>$tdFedTrustOrgCertificateNotAfterDateFL</span></div>
<div> <b>Token Issuer Certificate Expiry: </b> <span style='color:$tdFedTrustTokenIssuerCertificateNotAfterDateTimeColor'>$tdFedTrustTokenIssuerCertificateNotAfterDateTimeFL</span></div>
<div> <b>Token Issuer Metadata EPR:</b> <span style='color:$tdFedTrustTokenIssuerMetadataEprAbsoluteUriColor'>$tdFedTrustTokenIssuerMetadataEprAbsoluteUriFL</span></div>
<div> <b>Token Issuer EPR: </b> <span style='color:$tdFedTrustTokenIssuerEprAbsoluteUriColor'>$tdFedTrustTokenIssuerEprAbsoluteUriFL</span></div>

</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function AvailabilityAddressSpaceHtml() {
    $script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - On-Premise Get-AvailabilityAddressSpace</th>
</tr>
<tr>
<td><b> Get-AvailabilityAddressSpace $Script:ExchangeOnlineDomain | fl ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name</b></td>
<td>
<div> <b>Forest Name: </b> $Script:tdAvailabilityAddressSpaceForestName</div>
<div> <b>Name: </b>$Script:tdAvailabilityAddressSpaceName</div>
<div> <b>UserName: </b> <span style='color:$Script:tdAvailabilityAddressSpaceUserNameColor'>$Script:tdAvailabilityAddressSpaceUserName</span></div>
<div> <b>Access Method: </b> <span style='color:$Script:tdAvailabilityAddressSpaceAccessMethodColor'>$Script:tdAvailabilityAddressSpaceAccessMethod</span></div>
<div> <b>ProxyUrl: </b> <span style='color:$Script:tdAvailabilityAddressSpaceProxyUrlColor'>$Script:tdAvailabilityAddressSpaceProxyUrl</span></div>
</td>
</tr>"
    $html | Out-File -FilePath $Script:htmlFile
}
function autoDVDHtmlOK() {
    $script:html +=
    "<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-AutoDiscoverVirtualDirectory</th>
</tr>
<tr>
<td><b>Get-AutoDiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*</b></td>
<td>
<div><b>============================</b></div>
<div><b>Identity:</b> $Script:AutoD_VD_Identity</div>
<div><b>Name:</b> $Script:AutoD_VD_Name </div>
<div><b>InternalAuthenticationMethods:</b> $Script:AutoD_VD_InternalAuthenticationMethods </div>
<div><b>ExternalAuthenticationMethods:</b> $Script:AutoD_VD_ExternalAuthenticationMethods </div>
<div><b>WSAuthentication:</b> <span style='color:green'>$Script:AutoD_VD_WSAuthentication</span></div>
<div><b>WindowsAuthentication:</b> <span style='color:green'>$Script:AutoD_VD_WindowsAuthentication</span></div>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function autoDVDHtmlNotOK() {
    $script:html +=
    "<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-AutoDiscoverVirtualDirectory</th>
</tr>
<tr>
<td><b>Get-AutoDiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*</b></td>
<td>
<div><b>============================</b></div>
<div><b>Identity:</b> $Script:AutoD_VD_Identity</div>
<div><b>Name:</b> $Script:AutoD_VD_Name </div>
<div><b>InternalAuthenticationMethods:</b> $Script:AutoD_VD_InternalAuthenticationMethods </div>
<div><b>ExternalAuthenticationMethods:</b> $Script:AutoD_VD_ExternalAuthenticationMethods </div>
<div><b>WSAuthentication:</b> <span style='color:green'>$Script:AutoD_VD_WSAuthentication</span></div>
<div><b>WindowsAuthentication:</b> <span style='color:green'>$Script:AutoD_VD_WindowsAuthentication</span></div>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function EWSVirtualDHeaderHtml() {
    $script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-WebServicesVirtualDirectory</th>
</tr>
<tr>
<td><b> Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url</b></td>
<td >"
}
function EwsVDHtmlOK() {
    $script:html +=
    " <div><b>============================</b></div>
<div><b>Identity:</b>$Script:EwsVDIdentity</div>
<div><b>Name:</b>$Script:EwsVDName </div>
<div><b>InternalAuthenticationMethods:</b>$Script:EwsVDInternalAuthenticationMethods </div>
<div><b>ExternalAuthenticationMethods:</b>$Script:EwsVDExternalAuthenticationMethods </div>
<div><b>WSAuthentication:</b> <span style='color:green'>$EwsVD_WSAuthentication</span></div>
<div><b>WindowsAuthentication:</b> <span style='color:$Script:EwsVDWindowsAuthenticationColor'>$Script:EwsVDWindowsAuthentication</span></div>
<div><b>InternalUrl:</b>$Script:EwsVDInternalUrl </div>
<div><b>ExternalUrl:</b>$Script:EwsVDExternalUrl </div>
</td>
</tr>  "
    $html | Out-File -FilePath $Script:htmlFile
}
function EwsVDHtmlNotOK() {
    $script:html +=
    " <div><b>============================</b></div>
<div><b>Identity:</b>$Script:EwsVDIdentity</div>
<div><b>Name:</b>$Script:EwsVDName </div>
<div><b>InternalAuthenticationMethods:</b>$Script:EwsVDInternalAuthenticationMethods </div>
<div><b>ExternalAuthenticationMethods:</b>$Script:EwsVDExternalAuthenticationMethods </div>
<div><b>WSAuthentication:</b> <span style='color:red'>$EwsVD_WSAuthentication</span></div>
<div><b>WindowsAuthentication:</b> <span style='color:$Script:EwsVDWindowsAuthenticationColor'>$Script:EwsVDWindowsAuthentication</span></div>
<div><b>InternalUrl:</b>$Script:EwsVDInternalUrl </div>
<div><b>ExternalUrl:</b>$Script:EwsVDExternalUrl </div>
</td>
</tr>  "
    $html | Out-File -FilePath $Script:htmlFile
}
function TestOrgRelHtmlOK() {
    $Script:html += "<tr>
<th ColSpan='2' style='color:white;'><b>Summary - Test-OrganizationRelationship</b></th>
</tr>
<tr>
<td><b>Test-OrganizationRelationship -Identity $Script:OrgRelIdentity  -UserIdentity $Script:UserOnPrem</b></td>
<td>
<div class='green'> <b>No Significant Issues to Report</b><div>"
}
function TestOrgRelHtmlNotOK() {
    $Script:html += "<tr>
<th ColSpan='2' style='color:white;'><b>Summary - Test-OrganizationRelationship</b></th>
</tr>
<tr>
<td><b>Test-OrganizationRelationship -Identity $Script:OrgRelIdentity  -UserIdentity $Script:UserOnPrem</b></td>
<td>
<div class='red'> <b>Test Organization Relationship Completed with errors</b><div>"
}
function TestOrgRelHtmlNoUri() {
    $Script:html += "
<tr>
<th ColSpan='2' style='color:white;'><b>Summary - Test-OrganizationRelationship</b></th>
</tr>
<tr>
<td><b>Test-OrganizationRelationship</b></td>
<td>
<div class='red'> <b> Test-OrganizationRelationship can't be run if the Organization Relationship Target Application uri is not correct. Organization Relationship Target Application Uri should be Outlook.com</b><div>"
}
#On Prem OAuth
function IntraOrgConCheckHtml() {
    # Build HTML table row
    if ($Auth -like "OAuth") {
        $Script:html += "
  <div class='Black'><p></p></div>
  <div class='Black'><h2><b>`n Exchange On Premise Free Busy Configuration: `n</b></h2></div>
  <div class='Black'><p></p></div>"
    }
    $Script:html += "
<table style='width:100%'>
<tr>
<th ColSpan='2' style='text-align:center; color:white;'>Exchange On Premise OAuth Configuration</th>
</tr>
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-IntraOrganizationConnector</th>
</tr>
<tr>
<td><b>Get-IntraOrganizationConnector:</b></td>
<td>
  <div><b>TarGet Address Domains:</b><span style='color: $Script:tdIntraOrgTarGetAddressDomainColor'>$($Script:tdIntraOrgTarGetAddressDomain)</span></div>
  <div><b>Discovery Endpoint:</b><span style='color: $Script:tdDiscoveryEndpointColor;'>$($Script:tdDiscoveryEndpoint)</span></div>
  <div><b>Enabled:</b><span style='color: $Script:tdEnabledColor;'>$($Script:tdEnabled)</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function AuthServerCheckHtml() {
    $Script:html += "
<tr>
  <th ColSpan='2' style='color:white;'>Summary - Get-AuthServer</th>
</tr>
<tr>
  <td><b> Get-AuthServer | Select Name,IssuerIdentifier,TokenIssuingEndpoint,AuthMetadataUrl,Enabled</b></td>
  <td>
    <div><b>IssuerIdentifier:</b><span style='color: $Script:tDAuthServerIssuerIdentifierColor'>$($Script:tDAuthServerIssuerIdentifier)</span></div>
    <div><b>TokenIssuingEndpoint:</b><span style='color: $Script:tDAuthServerTokenIssuingEndpointColor;'>$($Script:tDAuthServerTokenIssuingEndpoint)</span></div>
    <div><b>AuthMetadataUrl:</b><span style='color: $Script:tDAuthServerAuthMetadataUrlColor;'>$($Script:tDAuthServerAuthMetadataUrl)</span></div>
    <div><b>Enabled:</b><span style='color: $Script:tDAuthServerEnabledColor;'>$($Script:tDAuthServerEnabled)</span></div>
  </td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function PartnerApplicationCheckHtml() {
    $Script:html += "
<tr>
  <th ColSpan='2' style='color:white;'>Summary - Get-PartnerApplication</th>
</tr>
<tr>
  <td><b> Get-PartnerApplication |  ?{`$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000'
-and `$_.Realm -eq ''} | Select Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer,
AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name</b></td>
  <td>
    <div><b>Enabled:</b><span style='color:$Script:tdPartnerApplicationEnabledColor'>$($tdPartnerApplicationEnabled)</span></div>
    <div><b>ApplicationIdentifier:</b><span style='color:$Script:tdPartnerApplicationApplicationIdentifierColor;'>$($tdPartnerApplicationApplicationIdentifier)</span></div>
    <div><b>CertificateStrings:</b><span style='color:$Script:tdPartnerApplicationCertificateStringsColor;'>$($tdPartnerApplicationCertificateStrings)</span></div>
    <div><b>AuthMetadataUrl:</b><span style='color:$Script:tdPartnerApplicationAuthMetadataUrlColor;'>$($tdPartnerApplicationAuthMetadataUrl)</span></div>
    <div><b>Realm:</b><span style='color:$Script:tdPartnerApplicationRealmColor'>$($tdPartnerApplicationRealm)</span></div>
    <div><b>LinkedAccount:</b><span style='color:$Script:tdPartnerApplicationLinkedAccountColor;'>$($tdPartnerApplicationLinkedAccount)</span></div>
    <div><b>IssuerIdentifier:</b><span style='color:$Script:tdPartnerApplicationEnabledColor'>$($tdPartnerApplicationEnabled)</span></div>
    <div><b>AppOnlyPermissions:</b><span style='color:$Script:tdPartnerApplicationApplicationIdentifierColor;'>$($tdPartnerApplicationApplicationIdentifier)</span></div>
    <div><b>ActAsPermissions:</b><span style='color:$Script:tdPartnerApplicationCertificateStringsColor;'>$($tdPartnerApplicationCertificateStrings)</span></div>
    <div><b>Name:</b><span style='color:$Script:tdPartnerApplicationAuthMetadataUrlColor;'>$($tdPartnerApplicationAuthMetadataUrl)</span></div>
  </td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function ApplicationAccountCheckHtml() {
    $Script:html += "
  <tr>
  <th ColSpan='2' style='color:white;'>Summary - Get-User ApplicationAccount</th>
</tr>
<tr>
  <td><b>  Get-user '$Script:ExchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount' | Select Name, RecipientType, RecipientTypeDetails, UserAccountControl':</b></td>
  <td>
    <div><b>RecipientType:</b><span style='color: $Script:tdApplicationAccountRecipientTypeColor'>$($Script:tdApplicationAccountRecipientType)</span></div>
    <div><b>RecipientTypeDetails:</b><span style='color: $Script:tdApplicationAccountRecipientTypeDetailsColor;'>$($Script:tdApplicationAccountRecipientTypeDetails)</span></div>
    <div><b>UserAccountControl:</b><span style='color: $Script:tdApplicationAccountUserAccountControlColor;'>$($Script:tdApplicationAccountUserAccountControl)</span></div>

  </td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function ManagementRoleAssignmentCheckHtml() {
    $Script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-ManagementRoleAssignment</th>
</tr>
<tr>
<td><b>  Get-ManagementRoleAssignment -RoleAssignee Exchange Online-ApplicationAccount | Select Name,Role</b></td>
<td>
  <div><b>UserApplication Role:</b><span style='color: $Script:tdManagementRoleAssignmentUserApplicationColor'>$($Script:tdManagementRoleAssignmentUserApplication)</span></div>
  <div><b>ArchiveApplication Role:</b><span style='color: $Script:tdManagementRoleAssignmentArchiveApplicationColor;'>$($Script:tdManagementRoleAssignmentArchiveApplication)</span></div>
  <div><b>LegalHoldApplication Role:</b><span style='color: $Script:tdManagementRoleAssignmentLegalHoldApplicationColor;'>$($Script:tdManagementRoleAssignmentLegalHoldApplication)</span></div>
  <div><b>Mailbox Search Role:</b><span style='color: $Script:tdManagementRoleAssignmentMailboxSearchColor'>$($Script:tdManagementRoleAssignmentMailboxSearch)</span></div>
  <div><b>TeamMailboxLifecycleApplication Role:</b><span style='color: $Script:tdManagementRoleAssignmentTeamMailboxLifecycleApplicationColor;'>$($Script:tdManagementRoleAssignmentTeamMailboxLifecycleApplication)</span></div>
  <div><b>MailboxSearchApplication Role:</b><span style='color: $Script:tdManagementRoleMailboxSearchApplicationColor;'>$($Script:tdManagementRoleMailboxSearchApplication)</span></div>
  <div><b>MeetingGraphApplication Role:</b><span style='color: $Script:tdManagementRoleMeetingGraphApplicationColor;'>$($Script:tdManagementRoleMeetingGraphApplication)</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function AuthConfigCheckHtml() {
    $Script:html += "
  <tr>
  <th ColSpan='2' style='color:white;'>Summary - Get-AuthConfig</th>
</tr>
<tr>
  <td><b>  Get-AuthConfig | Select-Object *Thumbprint, ServiceName, Realm, Name</b></td>
  <td>
    <div><b>Name:</b><span >$($Script:tDAuthConfigName)</span></div>
    <div><b>Thumbprint:</b><span style='color: $Script:tDAuthConfigCurrentCertificateThumbprintColor'>$($Script:tDAuthConfigCurrentCertificateThumbprint)</span></div>
    <div><b>ServiceName:</b><span style='color:$Script:tDAuthConfigServiceNameColor;'>$( $Script:tDAuthConfigServiceName)</span></div>
    <div><b>Realm:</b><span style='color: $Script:tDAuthConfigRealmColor;'>$($Script:tDAuthConfigRealm)</span></div>
  </td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function CurrentCertificateThumbprintCheckHtml() {
    $Script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-ExchangeCertificate AuthCertificate</th>
</tr>
<tr>
<td><b>  Get-ExchangeCertificate $thumb.CurrentCertificateThumbprint | Select-Object *</b></td>
<td>
  <div><b>Issuer:</b><span style='color:$Script:tdCurrentCertificateIssuerColor'>$($tdCurrentCertificateIssuer)</span></div>
  <div><b>Services:</b><span style='color:$Script:tdCurrentCertificateServicesColor'>$($tdCurrentCertificateServices)</span></div>
  <div><b>Status:</b><span style='color:$tdCurrentCertificateStatusColor;'>$($Script:tdCurrentCertificateStatus)</span></div>
  <div><b>Subject:</b><span style='color:$Script:tdCurrentCertificateSubjectColor;'>$($tdCurrentCertificateSubject)</span></div>
  <div><b>Distribution:</b><span style='color:$Script:tdCheckAuthCertDistributionColor;'>$($tdCheckAuthCertDistribution)</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function OAuthConnectivityCheckHtml() {
    $Script:html += "
  <tr>
  <th ColSpan='2' style='color:white;'>Summary - Test-OAuthConnectivity</th>
</tr>
<tr>
  <td><b>  Test-OAuthConnectivity -Service EWS -TarGetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $Script:UserOnPrem | fl</b></td>
  <td>
    <div><b>Result:</b><span style='color: $Script:OAuthConnectivityResultTypeColor'> $Script:OAuthConnectivityResultType</span></div>
  </td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function AvailabilityAddressSpaceCheckOAuth() {
    Write-Host -ForegroundColor Green " Get-AvailabilityAddressSpace $Script:ExchangeOnlineDomain | Select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
    PrintDynamicWidthLine
    $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $Script:ExchangeOnlineDomain | Select-Object ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    $AAS = $AvailabilityAddressSpace | Format-List
    $AAS
    if ($Auth -contains "OAuth") {
    }
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " Summary - On-Prem Availability Address Space"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor White " ForestName: "
    if ($AvailabilityAddressSpace.ForestName -like $Script:ExchangeOnlineDomain) {
        Write-Host -ForegroundColor Green " "$AvailabilityAddressSpace.ForestName
        $Script:tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $Script:tdAvailabilityAddressSpaceForestNameColor = "green"
    } else {
        Write-Host -ForegroundColor Red " ForestName is NOT correct. "
        Write-Host -ForegroundColor White " Should be $Script:ExchangeOnlineDomain "
        $Script:tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $Script:tdAvailabilityAddressSpaceForestNameColor = "red"
    }
    Write-Host -ForegroundColor White " UserName: "
    if ($AvailabilityAddressSpace.UserName -like "") {
        Write-Host -ForegroundColor Green "  Blank "
        $Script:tdAvailabilityAddressSpaceUserName = "  Blank. This is the correct value. "
        $Script:tdAvailabilityAddressSpaceUserNameColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  UserName is NOT correct. "
        Write-Host -ForegroundColor White "  Should be blank "
        $Script:tdAvailabilityAddressSpaceUserName = "  Blank. This is the correct value. "
        $Script:tdAvailabilityAddressSpaceUserNameColor = "red"
    }
    Write-Host -ForegroundColor White " UseServiceAccount: "
    if ($AvailabilityAddressSpace.UseServiceAccount -like "True") {
        Write-Host -ForegroundColor Green "  True "
        $Script:tdAvailabilityAddressSpaceUseServiceAccount = $AvailabilityAddressSpace.UseServiceAccount
        $Script:tdAvailabilityAddressSpaceUseServiceAccountColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  UseServiceAccount is NOT correct."
        Write-Host -ForegroundColor White "  Should be True "
        $Script:tdAvailabilityAddressSpaceUseServiceAccount = "$($tAvailabilityAddressSpace.UseServiceAccount). Should be True"
        $Script:tdAvailabilityAddressSpaceUseServiceAccountColor = "red"
    }
    Write-Host -ForegroundColor White " AccessMethod: "
    if ($AvailabilityAddressSpace.AccessMethod -like "InternalProxy") {
        Write-Host -ForegroundColor Green "  InternalProxy "
        $Script:tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $Script:tdAvailabilityAddressSpaceAccessMethodColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  AccessMethod is NOT correct. "
        Write-Host -ForegroundColor White "  Should be InternalProxy "
        $Script:tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $Script:tdAvailabilityAddressSpaceAccessMethodColor = "red"
    }
    Write-Host -ForegroundColor White " ProxyUrl: "
    if ($AvailabilityAddressSpace.ProxyUrl -like $Script:ExchangeOnPremEWS) {
        Write-Host -ForegroundColor Green " "$AvailabilityAddressSpace.ProxyUrl
        $Script:tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $Script:tdAvailabilityAddressSpaceProxyUrlColor = "green"
    } else {
        Write-Host -ForegroundColor Red "  ProxyUrl is NOT correct. "
        Write-Host -ForegroundColor White "  Should be $Script:ExchangeOnPremEWS"
        $Script:tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $Script:tdAvailabilityAddressSpaceProxyUrlColor = "red"
    }
    AvailabilityAddressSpaceCheckOAuthHtml
}
function  AvailabilityAddressSpaceCheckOAuthHtml() {
    $Script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-AvailabilityAddressSpace</th>
</tr>
<tr>
<td><b>  Get-AvailabilityAddressSpace $Script:ExchangeOnlineDomain | Select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name</b></td>
<td>
  <div><b>AddressSpaceForestName:</b><span style='color: $Script:tdAvailabilityAddressSpaceForestNameColor'>$($Script:tdAvailabilityAddressSpaceForestName)</span></div>
  <div><b>AddressSpaceUserName:</b><span style='color: $Script:tdAvailabilityAddressSpaceUserNameColor'>$($Script:tdAvailabilityAddressSpaceUserName)</span></div>
  <div><b>UseServiceAccount:</b><span style='color:$Script:tdAvailabilityAddressSpaceUseServiceAccountColor;'>$( $Script:tdAvailabilityAddressSpaceUseServiceAccount)</span></div>
  <div><b>AccessMethod:</b><span style='color: $Script:tdAvailabilityAddressSpaceAccessMethodColor;'>$($Script:tdAvailabilityAddressSpaceAccessMethod)</span></div>
  <div><b>ProxyUrl:</b><span style='color: $Script:tdAvailabilityAddressSpaceProxyUrlColor;'>$($Script:tdAvailabilityAddressSpaceProxyUrl)</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function AutoDVirtualDCheckOauthHtmlHead() {
    $script:html += "<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-AutoDiscoverVirtualDirectory</th>
</tr>
<tr>
<td><b>Get-AutoDiscoverVirtualDirectory:</b></td>
<td>"
    $html | Out-File -FilePath $Script:htmlFile
}
function  AutoDVirtualDCheckOauthHtmlOk() {
    $script:html +=
    " <div><b>============================</b></div>
        <div><b>Identity:</b> $Script:AutoD_VD_Identity</div>
        <div><b>Name:</b> $Script:AutoD_VD_Name </div>
        <div><b>InternalAuthenticationMethods:</b> $Script:AutoD_VD_InternalAuthenticationMethods </div>
        <div><b>ExternalAuthenticationMethods:</b> $Script:AutoD_VD_ExternalAuthenticationMethods </div>
        <div><b>WSAuthentication:</b> <span style='color:green'>$Script:AutoD_VD_WSAuthentication</span></div>
        <div><b>WindowsAuthentication:</b> <span style='color:green'>$Script:AutoD_VD_WindowsAuthentication</span></div>
        <div><b>OAuthAuthentication:</b> <span style='color:$Script:AutoD_VD_OAuthAuthenticationColor'>$Script:AutoD_VD_OAuthAuthentication</span></div>
        "
    $html | Out-File -FilePath $Script:htmlFile
}
function  AutoDVirtualDCheckOauthHtmlNotOk() {
    $script:html +=
    " <div><b>============================</b></div>
        <div><b>Identity:</b> $Script:AutoD_VD_Identity</div>
        <div><b>Name:</b> $Script:AutoD_VD_Name </div>
        <div><b>InternalAuthenticationMethods:</b> $Script:AutoD_VD_InternalAuthenticationMethods </div>
        <div><b>ExternalAuthenticationMethods:</b> $Script:AutoD_VD_ExternalAuthenticationMethods </div>
        <div><b>WSAuthentication:</b> <span style='color:red'>$Script:AutoD_VD_WSAuthentication</span></div>
        <div><b>WindowsAuthentication:</b> <span style='color:$Script:AutoD_VD_WindowsAuthenticationColor'>$Script:AutoD_VD_WindowsAuthentication</span></div>
        <div><b>OAuthAuthentication:</b> <span style='color:$Script:AutoD_VD_OAuthAuthenticationColor'>$Script:AutoD_VD_OAuthAuthentication</span></div>
        "
    $html | Out-File -FilePath $Script:htmlFile
}
function EWSVirtualDirectoryCheckOAuthHtmlHead() {
    $script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-WebServicesVirtualDirectory</th>
</tr>
<tr>
<td><b>Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url</b></td>
<td >"
    $html | Out-File -FilePath $Script:htmlFile
}
function  EWSVirtualDirectoryCheckOAuthHtmlOk() {
    $script:html +=
    " <div><b>============================</b></div>
        <div><b>Identity:</b>$Script:EwsVDIdentity</div>
        <div><b>Name:</b>$Script:EwsVDName </div>
        <div><b>InternalAuthenticationMethods:</b>$Script:EwsVDInternalAuthenticationMethods </div>
        <div><b>ExternalAuthenticationMethods:</b>$Script:EwsVDExternalAuthenticationMethods </div>
        <div><b>WSAuthentication:</b> <span style='color:green'>$EwsVD_WSAuthentication</span></div>
        <div><b>WindowsAuthentication:</b> <span style='color:$Script:EwsVDWindowsAuthenticationColor'>$Script:EwsVDWindowsAuthentication</span></div>
        <div><b>OAuthAuthentication:</b> <span style='color:$EwsVDW_OAuthAuthenticationColor'>$EwsVDOAuthAuthentication</span></div>
        <div><b>InternalUrl:</b>$Script:EwsVDInternalUrl </div>
        <div><b>ExternalUrl:</b>$Script:EwsVDExternalUrl </div>  "
    $html | Out-File -FilePath $Script:htmlFile
}
function  EWSVirtualDirectoryCheckOAuthHtmlNotOk() {
    $script:html +=
    " <div><b>============================</b></div>
        <div><b>Identity:</b>$Script:EwsVDIdentity</div>
        <div><b>Name:</b>$Script:EwsVDName </div>
        <div><b>InternalAuthenticationMethods:</b>$Script:EwsVDInternalAuthenticationMethods </div>
        <div><b>ExternalAuthenticationMethods:</b>$Script:EwsVDExternalAuthenticationMethods </div>
        <div><b>WSAuthentication:</b> <span style='color:red'>$EwsVD_WSAuthentication</span></div>
        <div><b>WindowsAuthentication:</b> <span style='color:$Script:EwsVDWindowsAuthenticationColor'>$Script:EwsVDWindowsAuthentication</span></div>
        <div><b>OAuthAuthentication:</b> <span style='color:$EwsVDW_OAuthAuthenticationColor'>$EwsVDOAuthAuthentication</span></div>
        <div><b>InternalUrl:</b>$Script:EwsVDInternalUrl </div>
        <div><b>ExternalUrl:</b>$Script:EwsVDExternalUrl </div>  "
    $html | Out-File -FilePath $Script:htmlFile
}
#Exo HTML DAuth output
function ExoOrgRelCheckHtml() {
    $script:html += "
<div class='Black'><p></p></div>
<div class='Black'><p></p></div>
<tr>
  <th ColSpan='2' style='text-align:center; color:white;'>Exchange Online DAuth Configuration</th>
</tr>
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-OrganizationRelationship</th>
</tr>
<tr>
<td><b>  Get-OrganizationRelationship  | Where{($_.DomainNames -like $Script:ExchangeOnPremDomain )} | Select Identity,DomainNames,FreeBusy*,TarGet*,Enabled</b></td>
<td>
  <div><b>Domain Names:</b><span >$($Script:tdExoOrgRelDomainNamesData)</span></div>
  <div><b>FreeBusyAccessEnabled:</b><span style='color:$Script:tdExoOrgRelFreeBusyAccessEnabledColor'>$($Script:tdExoOrgRelFreeBusyAccessEnabled)</span></div>
  <div><b>FreeBusyAccessLevel::</b><span style='color:$Script:tdExoOrgRelFreeBusyAccessLevelColor;'>$( $Script:tdExoOrgRelFreeBusyAccessLevel)</span></div>
  <div><b>TarGetApplicationUri:</b><span style='color: $Script:tdExoOrgRelTarGetApplicationUriColor;'>$($Script:tdExoOrgRelTarGetApplicationUri)</span></div>
  <div><b>TarGetOwAUrl:</b><span >$($tdExoOrgRelTarGetOwAUrl)</span></div>
  <div><b>TarGetSharingEpr:</b><span style='color: $Script:tdExoOrgRelTarGetSharingEprColor'>$($Script:tdExoOrgRelTarGetSharingEpr)</span></div>
  <div><b>TarGetAutoDiscoverEpr:</b><span style='color:$tdExoOrgRelFreeBusyAccessScopeColor;'>$($Script:tdExoOrgRelFreeBusyAccessScope)</span></div>
  <div><b>Enabled:</b><span style='color: $Script:tdExoOrgRelEnabledColor;'>$($Script:tdExoOrgRelEnabled)</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function ExoFedOrgIdCheckHtml() {
    $script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-FederatedOrganizationIdentifier</th>
</tr>
<tr>
<td><b>  Get-FederatedOrganizationIdentifier | select AccountNameSpace,Domains,Enabled</b></td>
<td>
<div><b>Domains:</b><span style='color: $Script:tdExoFedOrgIdDomainsColor;'>$($Script:tdExoFedOrgIdDomains)</span></div>
<div><b>Enabled:</b><span style='color: $Script:tdExoFedOrgIdEnabledColor;'>$($Script:tdExoFedOrgIdEnabled)</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function SharingPolicyCheckHtml() {
    $script:html += "
<tr>
  <th ColSpan='2' style='color:white;'>Summary - Get-SharingPolicy</th>
</tr>
<tr>
  <td><b>  Get-SharingPolicy | select Domains,Enabled,Name,Identity</b></td>
  <td>
    <div><b>Exchange On Premises Sharing domains:</b></div>
    <div><b>Domain:</b>$($Script:SPOnprem.Domains.Domain[0])</div>
    <div><b>Action:</b>$($Script:SPOnprem.Domains.Actions[0])</div>
    <div><b>Domain:</b>$($Script:SPOnprem.Domains.Domain[1])</div>
    <div><b>Action:</b>$($Script:SPOnprem.Domains.Actions[1])</div>
    <div><p></p></div>
    <div><b>Exchange Online Sharing domains:</b></div>
    <div><b>Domain:</b>$($domain1[0])</div>
    <div><b>Action:</b>$( $domain1[1])</div>
    <div><b>Domain:</b>$($domain2[0])</div>
    <div><b>Action:</b>$( $domain2[1])</div>
    <div><p></p></div>
    <div><b>Sharing Policy - Exchange Online vs Exchange On Premise:</b></div>
    <div><span style='color: $Script:tdSharpingPolicyCheckColor;'>$($Script:tdSharpingPolicyCheck)</span></div>
  </td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function ExoTestOrgRelCheckHtml() {
    $exoIdentity = $ExoOrgRel.Identity
    $exoOrgRelTarGetApplicationUri = $exoOrgRel.TarGetApplicationUri
    $exoOrgRelTarGetOWAUrl = $ExoOrgRel.TarGetOwAUrl
    $script:html += "
  <tr>
      <th ColSpan='2' style='color:white;'>Summary - Test-OrganizationRelationship</th>
  </tr>
  <tr>
      <td><b>  Test-OrganizationRelationship -Identity $exoIdentity -UserIdentity $Script:UserOnline</b></td>
  <td>"
    if ((![string]::IsNullOrWhitespace($exoOrgRelTarGetApplicationUri)) -and (![string]::IsNullOrWhitespace($exoOrgRelTarGetOWAUrl))) {
        $i = 2
        while ($i -lt $ExoTestOrgRel.Length) {
            $element = $ExoTestOrgRel[$i]
            $aux = "0"
            if ($element -like "*RESULT:*" -and $aux -like "0") {
                $el = $element.TrimStart()
                if ($element -like "*Success.*") {
                    $Script:html += "
              <div> <b> $ExoTestOrgRelStep </b> <span style='color:green'> $el</span>"
                    $aux = "1"
                } elseif ($element -like "*Error*" -or $element -like "*Unable*") {
                    $Script:html += "
              <div> <b> $ExoTestOrgRelStep </b> <span style='color:red'> $el</span>"
                    $aux = "1"
                }
            } elseif ($aux -like "0" ) {
                if ($element -like "*STEP*" -or $element -like "*Complete*") {
                    $Script:html += "
              <p></p>
              <div> <b> $ExoTestOrgRelStep </b> <span style='color:black'> $element</span></div>"
                    $aux = "1"
                } else {
                    $ID = $element.ID
                    $Status = $element.Status
                    $Description = $element.Description
                    if (![string]::IsNullOrWhitespace($ID)) {
                        $Script:html += "<div> <b>ID: </b> <span style='color:black'> $ID</span></div>"
                        if ($Status -like "*Success*") {
                            $Script:html += "<div> <b>Status:</b> <span style='color:green'> $Status</span></div>"
                        }

                        if ($status -like "*error*") {
                            $Script:html += "<div> <b>Status:</b> <span style='color:red'> $Status</span></div>"
                        }

                        $Script:html += "<div> <b>Description: </b> <span style='color:black'> $Description</span></div>
                  <div><span style='color:blue'>Note: Test-Organization Relationship fails on Step 3 with error MismatchedFederation if Hybrid Agent is in use</span></div>"
                    }
                    #$element
                    $aux = "1"
                }
            }
            $i++
        }
    }

    elseif ((([string]::IsNullOrWhitespace($exoOrgRelTarGetApplicationUri)) -and ([string]::IsNullOrWhitespace($exoOrgRelTarGetOWAUrl)))) {
        $Script:html += "
  <div> <span style='color:red'> Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TarGetApplicationUri and TarGetOwAUrl are not set</span>"
    } elseif ((([string]::IsNullOrWhitespace($exoOrgRelTarGetApplicationUri)) )) {
        $Script:html += "
  <div> <span style='color:red'> Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TarGetApplicationUri is not set</span>"
    } elseif ((([string]::IsNullOrWhitespace($exoOrgRelTarGetApplicationUri)) )) {
        $Script:html += "
  <div> <span style='color:red'> Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TarGetApplicationUri is not set</span>"
    }
    $Script:html += "</td>
</tr>"
    $html | Out-File -FilePath $Script:htmlFile
}
#Exo Oauth Functions
function EXOIntraOrgConCheckHtml() {
    $script:html += "
<tr>
<th ColSpan='2' style='text-align:center; color:white;'><b>Exchange Online OAuth Configuration</b></th>
</tr>
<tr>
<th ColSpan='2' style=' color:white;'><b>Summary - Get-IntraOrganizationConnector</b></th>
</tr>
<tr>
<td><b>  Get-IntraOrganizationConnector | Select-Object TarGetAddressDomains, DiscoveryEndpoint, Enabled</b></td>
<td>
  <div><b>TarGet Address Domains:</b><span style='color:$Script:tdExoIntraOrgConTarGetAddressDomainsColor;'>' $($tdExoIntraOrgConTarGetAddressDomains)'</span></div>
  <div><b>DiscoveryEndpoint:</b><span style='color: $Script:tdExoIntraOrgConDiscoveryEndpointsColor;'>' $($Script:tdExoIntraOrgConDiscoveryEndpoints)'</span></div>
  <div><b>Enabled:</b><span style='color:$Script:tdExoIntraOrgConEnabledColor;'> $($Script:tdExoIntraOrgConEnabled)</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function EXOIntraOrgConfigCheckHtml() {
    $script:html += "
<tr>
  <th ColSpan='2' style=color:white;'><b>Summary - Get-IntraOrganizationConfiguration</b></th>
</tr>
<tr>
  <td><b>  Get-IntraOrganizationConfiguration | Select OnPremiseTarGetAddresses</b></td>
  <td>
    <div><b>OnPremiseTarGetAddresses:</b><span style='color: $Script:tdExoIntraOrgConfigOnPremiseTarGetAddressesColor;'>$($Script:tdExoIntraOrgConfigOnPremiseTarGetAddresses)</span></div>
  </td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function EXOAuthServerCheckHtml() {
    $script:html += "
<tr>
<th ColSpan='2' style='color:white;'>Summary - Get-AuthServer</th>
</tr>
<tr>
<td><b>  Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | select name,IssuerIdentifier,enabled</b></td>
<td>
<div><b>Name:</b><span style='color: $Script:tdExoAuthServerNameColor;'>$($Script:tdExoAuthServerName)</span></div>
<div><b>IssuerIdentifier:</b><span style='color: $Script:tdExoAuthServerIssuerIdentifierColor;'>$($Script:tdExoAuthServerIssuerIdentifier)</span></div>
<div><b>Enabled:</b><span style='color: $Script:tdExoAuthServerEnabledColor;'>$($Script:tdExoAuthServerEnabled)</span></div>
</td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function ExoTestOAuthCheckHtml() {
    $script:html += "
<tr>
  <th ColSpan='2' style='color:white;'><b>Summary - Test-OAuthConnectivity</b></th>
</tr>
<tr>
  <td><b>  Test-OAuthConnectivity -Service EWS -TarGetUri $($Script:ExchangeOnPremEWS) -Mailbox $Script:UserOnline </b></td>
  <td>
    <div><b>Result:</b><span style='color:$Script:tdOAuthConnectivityResultTypeColor;'>$($tdOAuthConnectivityResultType)</span></div>
  </td>
</tr>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function lookupMethodDAuthHtml() {
    $Script:html += "
                    <div  style = 'padding-left: 0%;'>
                      <h3>Intra Organization Connector Enabled: <b>True</b></h3>
                      <p> <span style='color: green; font-wight: 550; padding-left:2%; '>Checking DAuth only as -Auth DAuth option was selected</span></p>
                    </div>
                    <div  style = 'padding-left: 0%;'>
                      <ul>
                        <li>
                          This script can be run using the <b>-Auth OAuth</b> parameter to Check for DAuth configurations only.
                          <br />
                          <span style='padding-left: 2%;'>
                            <br />
                            <b>Example:</b> ./FreeBusyChecker.ps1 -Auth OAuth
                          </span>
                        </li>
                        <br />
                        <li>
                          This script can be run using the <b>-Auth All</b> parameter to Check for both OAuth and DAuth configurations.
                          <br />
                          <span style='padding-left: 2%;'>
                            <br />
                            <b>Example:</b> ./FreeBusyChecker.ps1 -Auth All
                          <span style='padding:2%;'>
                        </li>
                      </ul>
                    </div>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function lookupMethodOauthHtml() {
    $Script:html += "
  <div  style = 'padding-left: 0%;'>
    <h3>Intra Organization Connector Enabled: <b>True</b></h3>
    <p> <span style='color: green; font-wight: 550; padding-left:2%; '>Checking OAuth only as Free Busy Lookup is done using OAuth when the Intra Organization Connector is Enabled</span></p>
  </div>
  <div  style = 'padding-left: 0%;'>
    <ul>
      <li>
        This script can be run using the <b>-Auth DAuth</b> parameter to Check for DAuth configurations only.
        <br />
        <span style='padding-left: 2%;'>
          <br />
          <b>Example:</b> ./FreeBusyChecker.ps1 -Auth DAuth
        </span>
      </li>
      <br />
      <li>
        This script can be run using the <b>-Auth All</b> parameter to Check for both OAuth and DAuth configurations.
        <br />
        <span style='padding-left: 2%;'>
          <br />
          <b>Example:</b> ./FreeBusyChecker.ps1 -Auth All
        <span style='padding:2%;'>
      </li>
    </ul>
  </div>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function lookupMethodCheckAllHtml() {
    $Script:Auth = ""
    Write-Host -ForegroundColor White "    -> Free Busy Lookup is done using OAuth when the Intra Organization Connector is Enabled"
    Write-Host -ForegroundColor White "    -> Checking both OAuth and DAuth as -Auth All option was selected"
    $Script:html += "
              <div  style = 'padding-left: 0%;'>
                <h3>Intra Organization Connector Enabled: <b>True</b></h3>
                <p> <span style='color: green; font-wight: 550; padding-left:2%; '>Checking both OAuth and DAuth as -Auth All option was selected</span></p>
              </div>
              <div  style = 'padding-left: 0%;'>
                <ul>
                  <li>
                    This script can be run using the <b>-Auth DAuth</b> parameter to Check for DAuth configurations only.
                    <br />
                    <span style='padding-left: 2%;'>
                      <br />
                      <b>Example:</b> ./FreeBusyChecker.ps1 -Auth DAuth
                    </span>
                  </li>
                  <br />
                  <li>
                    This script can be run using the <b>-Auth All</b> parameter to Check for both OAuth and DAuth configurations.
                    <br />
                    <span style='padding-left: 2%;'>
                      <br />
                      <b>Example:</b> ./FreeBusyChecker.ps1 -Auth All
                    <span style='padding:2%;'>
                  </li>
                </ul>
              </div>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function lookupMethodDAuthOauthDisabledHtml() {
    $Script:html += "
                    <div  style = 'padding-left: 0%;'>
                      <h3>Intra Organization Connector Enabled: <b>False</b></h3>
                      <p> <span style='color: green; font-wight: 550; padding-left:2%; '>Checking DAuth as OAuth is not Enabled</span></p>
                    </div>
                    <div  style = 'padding-left: 0%;'>
                      <ul>
                        <li>
                          This script can be run using the <b>-Auth OAuth</b> parameter to Check for OAuth configurations only.
                          <br />
                          <span style='padding-left: 2%;'>
                            <br />
                            <b>Example:</b> ./FreeBusyChecker.ps1 -Auth OAuth
                          </span>
                        </li>
                        <br />
                        <li>
                          This script can be run using the <b>-Auth All</b> parameter to Check for both OAuth and DAuth configurations.
                          <br />
                          <span style='padding-left: 2%;'>
                            <br />
                            <b>Example:</b> ./FreeBusyChecker.ps1 -Auth All
                          <span style='padding:2%;'>
                        </li>
                      </ul>
                    </div>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function lookupMethodOauthOauthDisabledHtml() {
    $Script:html += "
                  <div  style = 'padding-left: 0%;'>
                    <h3>Intra Organization Connector Enabled: <b>False</b></h3>
                    <p> <span style='color: green; font-wight: 550; padding-left:2%; '>Checking OAuth as -Auth OAuth parameter was selected</span></p>
                  </div>
                  <div  style = 'padding-left: 0%;'>
                    <ul>
                      <li>
                        This script can be run using the <b>-Auth All</b> parameter to Check for both OAuth and DAuth configurations.
                        <br />
                        <span style='padding-left: 2%;'>
                          <br />
                          <b>Example:</b> ./FreeBusyChecker.ps1 -Auth All
                        <span style='padding:2%;'>
                      </li>
                    </ul>
                  </div>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function lookupMethodAllOauthDisabledHtml() {
    $Script:html += "
              <div  style = 'padding-left: 0%;'>
                <h3>Intra Organization Connector Enabled: <b>False</b></h3>
                <p> <span style='color: green; font-wight: 550; padding-left:2%; '>Checking both for OAuth and DAuth as -Auth All parameter was selected</span></p>
              </div>
              <div  style = 'padding-left: 0%;'>
                <ul>
                  <li>
                    This script can be run using the <b>-Auth OAuth</b> parameter to Check for OAuth only.
                    <br />
                    <span style='padding-left: 2%;'>
                      <br />
                      <b>Example:</b> ./FreeBusyChecker.ps1 -Auth OAuth
                    <span style='padding:2%;'>
                  </li>
                </ul>
              </div>
"
    $html | Out-File -FilePath $Script:htmlFile
}
function exoHeaderHtml() {
    $Script:html += "
 </table>
  <div class='Black'><p></p></div>
  <div class='Black'><p></p></div>
         <div class='Black'><h2><b>`n Exchange Online Free Busy Configuration: `n</b></h2></div>
         <div class='Black'><p></p></div>
         <div class='Black'><p></p></div>
 <table style='width:100%; margin-top:30px;'>
"
}
function PrintDynamicWidthLine {
    $screenWidth = $host.UI.RawUI.WindowSize.Width
    if ($screenWidth -gt 180) {
        $length = [math]::floor($screenWidth / 1.65)
    } else {
        $length = [math]::floor($screenWidth - 5)
    }
    $line = "=" * $length
    Write-Host $line
}
function ShowHelp() {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Yellow "`n  Valid Input Option Parameters!"
    Write-Host -ForegroundColor White "`n  Parameter: Auth"
    Write-Host -ForegroundColor White "   Options  : DAuth; OAUth"
    Write-Host  "    DAuth             : DAuth Authentication"
    Write-Host  "    OAuth             : OAuth Authentication"
    Write-Host  "    Default Value     : No switch input means the script will collect both DAuth and OAuth Availability configuration detail"
    Write-Host -ForegroundColor White "`n  Parameter: Org"
    Write-Host -ForegroundColor White "   Options  : ExchangeOnPremise; Exchange Online"
    Write-Host  "    ExchangeOnPremise : Use ExchangeOnPremise parameter to collect Availability information in the Exchange On Premise Tenant"
    Write-Host  "    ExchangeOnline    : Use Exchange Online parameter to collect Availability information in the Exchange Online Tenant"
    Write-Host  "    Default Value     : No switch input means the script will collect both Exchange On Premise and Exchange Online Availability configuration detail"
    Write-Host -ForegroundColor White "`n  Parameter: Pause"
    Write-Host  "                 : Use the Pause parameter to use this script pausing after each test done."
    Write-Host -ForegroundColor White "`n  Parameter: Help"
    Write-Host  "                 : Use the Help parameter to use display valid parameter Options. `n`n"
}

function loadingParameters() {
    $Script:LogFile = "$PSScriptRoot\FreeBusyChecker.txt"
    $BuildVersion = "26.06.06.1238"
    $LogFileName = [System.IO.Path]::GetFileNameWithoutExtension($LogFile) + "_" + $Script:startingDate + ([System.IO.Path]::GetExtension($Script:LogFile))
    Start-Transcript -Path $LogFileName -Append
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green " `n  Free Busy Configuration Information Checker `n "
    Write-Host -ForegroundColor White $BuildVersion
    Write-Host -ForegroundColor Green "  Loading Parameters..... `n "
}
function ShowParameters() {
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Green "  Loading modules for AD, Exchange"
    PrintDynamicWidthLine
    Write-Host   "  Color Scheme"
    PrintDynamicWidthLine
    Write-Host -ForegroundColor Red "  Look out for Red!"
    Write-Host -ForegroundColor Yellow "  Yellow - Example information or Links"
    Write-Host -ForegroundColor Green "  Green - In Summary Sections it means OK. Anywhere else it's just a visual aid."
    PrintDynamicWidthLine
    Write-Host   "  Parameters:"
    PrintDynamicWidthLine
    Write-Host  -ForegroundColor White " Log File Path:"
    Write-Host -ForegroundColor Green "  $Script:LogFile"
    Write-Host  -ForegroundColor White " Office 365 Domain:"
    Write-Host -ForegroundColor Green "  $Script:ExchangeOnlineDomain"
    Write-Host  -ForegroundColor White " AD root Domain"
    Write-Host -ForegroundColor Green "  $Script:ExchangeOnPremLocalDomain"
    Write-Host -ForegroundColor White " Exchange On Premises Domain:  "
    Write-Host -ForegroundColor Green "  $Script:ExchangeOnPremDomain"
    Write-Host -ForegroundColor White " Exchange On Premises External EWS url:"
    Write-Host -ForegroundColor Green "  $Script:ExchangeOnPremEWS"
    Write-Host -ForegroundColor White " On Premises Hybrid Mailbox:"
    Write-Host -ForegroundColor Green "  $Script:UserOnPrem"
    Write-Host -ForegroundColor White " Exchange Online Mailbox:"
    Write-Host -ForegroundColor Green "  $Script:UserOnline"
    showParametersHtml
}
function hostOutputIntraOrgConEnabled($Auth) {
    PrintDynamicWidthLine
    if ($Auth -like "DAuth") {
        Write-Host -ForegroundColor yellow "  Warning: Intra Organization Connector Enabled True -> Running for DAuth only as -Auth DAuth option was selected`n  "
        if ($Script:OrgRel.enabled -like "True") {
            Write-Host -ForegroundColor White "           Organization Relationship Enabled True `n  "
        }
        Write-Host -ForegroundColor White "      This script can be Run using the -Auth All parameter to Check for both OAuth and DAuth configuration. `n `n         Example: ./FreeBusyChecker.ps1 -Auth All"
        lookupMethodDAuthHtml
    }
    if ($Auth -like "") {
        $Auth = "OAuth"
        Write-Host -ForegroundColor White "    -> Free Busy Lookup is done using OAuth when the Intra Organization Connector is Enabled"
        Write-Host -ForegroundColor White "    -> Running for OAuth only as OAuth takes precedence over DAuth;"
        Write-Host -ForegroundColor White "`n         This script can be Run using the -Auth All parameter to Check for both OAuth and DAuth configuration. `n `n         Example: ./FreeBusyChecker.ps1 -Auth All"
        Write-Host -ForegroundColor White "`n         This script can be Run using the -Auth DAuth parameter to Check for DAuth configuration only. `n `n         Example: ./FreeBusyChecker.ps1 -Auth DAuth"
        lookupMethodOAuthHtml
    }
    if ($Auth -like "All") {
        lookupMethodCheckAllHtml
    }
    return $Auth
}
function hostOutputIntraOrgConNotEnabled() {
    if ($Auth -like "" -or $Auth -like "DAuth") {
        PrintDynamicWidthLine
        Write-Host -ForegroundColor yellow "  Warning: Intra Organization Connector Enabled False -> Running for DAuth only as OAuth is not enabled"
        Write-Host -ForegroundColor White "`n       This script can be Run using the '-Auth OAuth' parameter to Check for OAuth configurations only. `n"
        Write-Host -ForegroundColor White "             Example: ./FreeBusyChecker.ps1 -Auth OAuth"
        Write-Host -ForegroundColor White "`n       This script can be Run using the '-Auth All' parameter to Check for both OAuth and DAuth configuration. `n"
        Write-Host -ForegroundColor White "             Example: ./FreeBusyChecker.ps1 -Auth All"
        lookupMethodDAuthOauthDisabledHtml
    }
    if ($Auth -like "OAuth") {
        PrintDynamicWidthLine
        Write-Host -ForegroundColor yellow "  Warning: Intra Organization Connector Enabled False -> Running for OAuth only as -Auth OAuth parameter was selected"
        Write-Host -ForegroundColor White "`n       This script can be Run using the '-Auth All' parameter to Check for both OAuth and DAuth configuration. `n"
        Write-Host -ForegroundColor White "             Example: ./FreeBusyChecker.ps1 -Auth All"
        lookupMethodOauthOauthDisabledHtml
    }
    if ($Auth -like "All") {
        PrintDynamicWidthLine
        Write-Host -ForegroundColor yellow "  Warning: Intra Organization Connector Enabled False -> Running both for OAuth and DAuth as -Auth All parameter was selected"
        Write-Host -ForegroundColor White "`n       This script can be Run using the '-Auth OAuth' parameter to Check for OAuth configuration only. `n"
        Write-Host -ForegroundColor White "             Example: ./FreeBusyChecker.ps1 -Auth OAuth"
        lookupMethodAllOauthDisabledHtml
    }
}
function Test-ExchangeOnlineConnection {
    Write-Host -ForegroundColor Green " Checking Exchange Online Configuration"
    Write-Host " Testing Connection to Exchange Online with EO Prefix."
    try {
        $CheckExoMailbox = Get-EOMailbox $Script:UserOnline -ErrorAction Stop
        if ($null -ne $CheckExoMailbox) {
            return $true
        } else {
            return $false
        }
    } catch {
        return $false
    }
}
function FetchAutoDiscoverInformation {
    if (-not $Script:AutoDiscoveryVirtualDirectory -or -not $Script:AutoDiscoveryVirtualDirectoryOAuth) {
        $Script:AutoDiscoveryVirtualDirectory = Get-AutoDiscoverVirtualDirectory -Server $Script:Server | Select-Object Identity, Name, ExchangeVersion, *authentication* -ErrorAction SilentlyContinue
        $Script:AutoDiscoveryVirtualDirectoryOAuth = $Script:AutoDiscoveryVirtualDirectory
    }
}
function FetchEWSInformation {
    if (-not $Script:WebServicesVirtualDirectory -or -not $Script:WebServicesVirtualDirectoryOAuth) {
        $Script:WebServicesVirtualDirectory = Get-WebServicesVirtualDirectory -Server $Script:Server | Select-Object Identity, Name, ExchangeVersion, *Authentication*, *url -ErrorAction SilentlyContinue
        $Script:WebServicesVirtualDirectoryOAuth = $Script:WebServicesVirtualDirectory

        if ([string]::IsNullOrWhiteSpace($Script:ExchangeOnPremEWS)) {
            $externalUrl = $Script:WebServicesVirtualDirectory |
                Where-Object { $null -ne $_.ExternalUrl } |
                Select-Object -First 1 -ExpandProperty ExternalUrl

            if ($null -ne $externalUrl) {
                $Script:ExchangeOnPremEWS = $externalUrl.ToString()
            }
        }
    }
}
function CheckIfExchangeServer {
    param (
        [string]$Server
    )
    $exchangeServer = Get-ExchangeServer $Server -ErrorAction SilentlyContinue
    if (!$exchangeServer) {
        Write-Host "$Server is not an Exchange Server. This script should be run in Exchange Server Management Shell"
        exit
    }
}
function CheckParameters {
    $MissingParameters = @()
    if ([string]::IsNullOrWhiteSpace($Script:ExchangeOnlineDomain)) {
        $MissingParameters += "Exchange Online Domain. Example: contoso.mail.onmicrosoft.com"
    }
    if ([string]::IsNullOrWhiteSpace($Script:ExchangeOnPremLocalDomain)) {
        $MissingParameters += "Exchange On Premises Local Domain.  Example: . 'C:\scripts\FreeBusyChecker\FreeBusyChecker.ps1' -OnPremisesUser John@Contoso.com"
    }
    if ([string]::IsNullOrWhiteSpace($Script:ExchangeOnPremDomain)) {
        $MissingParameters += "Exchange On Premises Domain.  Example: -OnPremLocalDomain Contoso.local"
    }
    if ([string]::IsNullOrWhiteSpace($Script:ExchangeOnPremEWS)) {
        $MissingParameters += "Exchange On Premises EWS Virtual Directory External URL.  Example:  'C:\FreeBusyChecker.ps1' -OnPremEWSUrl https://mail.contoso.com/EWS/Exchange.asmx"
    }
    if ([string]::IsNullOrWhiteSpace($Script:UserOnPrem)) {
        $MissingParameters += "On Premises User Mailbox.  Example: 'C:\FreeBusyChecker.ps1' -OnPremisesUser John@Contoso.com"
    }
    if ([string]::IsNullOrWhiteSpace($Script:UserOnline)) {
        $MissingParameters += "Exchange Online Mailbox.  Example: 'C:\FreeBusyChecker.ps1' -OnlineUser John@Contoso.onmicrosoft.com"
    }

    if ($MissingParameters.Count -gt 0) {
        foreach ($param in $MissingParameters) {
            Write-Host -ForegroundColor Red "Please provide a value for $param."
        }
        exit 1
    }
    Write-Host -ForegroundColor Cyan "`n All parameters are valid."
    return
}


function Invoke-CatchActionError {
    [CmdletBinding()]
    param(
        [ScriptBlock]$CatchActionFunction
    )

    if ($null -ne $CatchActionFunction) {
        & $CatchActionFunction
    }
}

function Invoke-CatchActionErrorLoop {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [int]$CurrentErrors,
        [Parameter(Mandatory = $false, Position = 1)]
        [ScriptBlock]$CatchActionFunction
    )
    process {
        if ($null -ne $CatchActionFunction -and
            $Error.Count -ne $CurrentErrors) {
            $i = 0
            while ($i -lt ($Error.Count - $currentErrors)) {
                & $CatchActionFunction $Error[$i]
                $i++
            }
        }
    }
}

# Confirm that either Remote Shell or EMS is loaded from an Edge Server, Exchange Server, or a Tools box.
# It does this by also initializing the session and running Get-EventLogLevel. (Server Management RBAC right)
# All script that require Confirm-ExchangeShell should be at least using Server Management RBAC right for the user running the script.
function Confirm-ExchangeShell {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [bool]$LoadExchangeShell = $true,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed: LoadExchangeShell: $LoadExchangeShell"
        $currentErrors = $Error.Count
        $edgeTransportKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'
        $setupKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
        $remoteShell = (-not(Test-Path $setupKey))
        $toolsServer = (Test-Path $setupKey) -and
        (-not(Test-Path $edgeTransportKey)) -and
        ($null -eq (Get-ItemProperty -Path $setupKey -Name "Services" -ErrorAction SilentlyContinue))
        Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

        function IsExchangeManagementSession {
            [OutputType("System.Boolean")]
            param(
                [ScriptBlock]$CatchActionFunction
            )

            $getEventLogLevelCallSuccessful = $false
            $isExchangeManagementShell = $false

            try {
                $currentErrors = $Error.Count
                $attempts = 0
                do {
                    $eventLogLevel = Get-EventLogLevel -ErrorAction Stop | Select-Object -First 1
                    $attempts++
                    if ($attempts -ge 5) {
                        throw "Failed to run Get-EventLogLevel too many times."
                    }
                } while ($null -eq $eventLogLevel)
                $getEventLogLevelCallSuccessful = $true
                foreach ($e in $eventLogLevel) {
                    Write-Verbose "Type is: $($e.GetType().Name) BaseType is: $($e.GetType().BaseType)"
                    if (($e.GetType().Name -eq "EventCategoryObject") -or
                        (($e.GetType().Name -eq "PSObject") -and
                        ($null -ne $e.SerializationData))) {
                        $isExchangeManagementShell = $true
                    }
                }
                Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
            } catch {
                Write-Verbose "Failed to run Get-EventLogLevel"
                Invoke-CatchActionError $CatchActionFunction
            }

            return [PSCustomObject]@{
                CallWasSuccessful = $getEventLogLevelCallSuccessful
                IsManagementShell = $isExchangeManagementShell
            }
        }
    }
    process {
        $isEMS = IsExchangeManagementSession $CatchActionFunction
        if ($isEMS.CallWasSuccessful) {
            Write-Verbose "Exchange PowerShell Module already loaded."
        } else {
            if (-not ($LoadExchangeShell)) { return }

            #Test 32 bit process, as we can't see the registry if that is the case.
            if (-not ([System.Environment]::Is64BitProcess)) {
                Write-Warning "Open a 64 bit PowerShell process to continue"
                return
            }

            if (Test-Path "$setupKey") {
                Write-Verbose "We are on Exchange 2013 or newer"

                try {
                    $currentErrors = $Error.Count
                    if (Test-Path $edgeTransportKey) {
                        Write-Verbose "We are on Exchange Edge Transport Server"
                        [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exShell.psc1" -ErrorAction Stop

                        foreach ($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn) {
                            Write-Verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                            Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                        }

                        Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop
                    } else {
                        Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                        Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                    }
                    Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

                    Write-Verbose "Imported Module. Trying Get-EventLogLevel Again"
                    $isEMS = IsExchangeManagementSession $CatchActionFunction
                    if (($isEMS.CallWasSuccessful) -and
                        ($isEMS.IsManagementShell)) {
                        Write-Verbose "Successfully loaded Exchange Management Shell"
                    } else {
                        Write-Warning "Something went wrong while loading the Exchange Management Shell"
                    }
                } catch {
                    Write-Warning "Failed to Load Exchange PowerShell Module..."
                    Invoke-CatchActionError $CatchActionFunction
                }
            } else {
                Write-Verbose "Not on an Exchange or Tools server"
            }
        }
    }
    end {

        $returnObject = [PSCustomObject]@{
            ShellLoaded = $isEMS.CallWasSuccessful
            Major       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMajor" -ErrorAction SilentlyContinue).MsiProductMajor)
            Minor       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMinor" -ErrorAction SilentlyContinue).MsiProductMinor)
            Build       = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMajor" -ErrorAction SilentlyContinue).MsiBuildMajor)
            Revision    = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMinor" -ErrorAction SilentlyContinue).MsiBuildMinor)
            EdgeServer  = $isEMS.CallWasSuccessful -and (Test-Path $setupKey) -and (Test-Path $edgeTransportKey)
            ToolsOnly   = $isEMS.CallWasSuccessful -and $toolsServer
            RemoteShell = $isEMS.CallWasSuccessful -and $remoteShell
            EMS         = $isEMS.IsManagementShell
        }

        return $returnObject
    }
}

<#
    This set of code is designed to handle updating your script as this code is basically the same everywhere, making this a common file to avoid duplication.
    Just need to dot load the file to your script and have the correct parameters, then this code does the work for you.
    These are the parameters that you should have within your script.
    This needs to be done within the main part of the script, not inside a function to work correctly.

    [Parameter(Mandatory = $false, ParameterSetName = "ScriptUpdateOnly")]
    [switch]$ScriptUpdateOnly,

    [switch]$SkipVersionCheck
#>





function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TargetUri
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Verbose "Proxy server configuration detected"
            Write-Verbose $proxyObject.OriginalString
            return $true
        } else {
            Write-Verbose "No proxy server configuration detected"
            return $false
        }
    } catch {
        Write-Verbose "Unable to check for proxy server configuration"
        return $false
    }
}

function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    [string]$errorInformation = [System.Environment]::NewLine + [System.Environment]::NewLine +
    "----------------Error Information----------------" + [System.Environment]::NewLine

    if ($null -ne $CurrentError.OriginInfo) {
        $errorInformation += "Error Origin Info: $($CurrentError.OriginInfo.ToString())$([System.Environment]::NewLine)"
    }

    $errorInformation += "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())$([System.Environment]::NewLine)"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        $errorInformation += "Inner Exception: $($CurrentError.Exception.StackTrace)$([System.Environment]::NewLine)"
    } elseif ($null -ne $CurrentError.Exception) {
        $errorInformation += "Inner Exception: $($CurrentError.Exception)$([System.Environment]::NewLine)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        $errorInformation += "Position Message: $($CurrentError.InvocationInfo.PositionMessage)$([System.Environment]::NewLine)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        $errorInformation += "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)$([System.Environment]::NewLine)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        $errorInformation += "Script Stack: $($CurrentError.ScriptStackTrace)$([System.Environment]::NewLine)"
    }

    $errorInformation += "-------------------------------------------------$([System.Environment]::NewLine)$([System.Environment]::NewLine)"

    & $Cmdlet $errorInformation
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")]
        [string]
        $Uri,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]
        $UseBasicParsing,

        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")]
        [hashtable]
        $ParametersObject,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [string]
        $OutFile
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    try {
        Invoke-WebRequest @params
    } catch {
        Write-VerboseErrorInformation
    }
}

<#
    Determines if the script has an update available.
#>
function Get-ScriptUpdateAvailable {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $BuildVersion = "26.06.06.1238"

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $result = [PSCustomObject]@{
        ScriptName     = $scriptName
        CurrentVersion = $BuildVersion
        LatestVersion  = ""
        UpdateFound    = $false
        Error          = $null
    }

    if ((Get-AuthenticodeSignature -FilePath $scriptFullName).Status -eq "NotSigned") {
        Write-Warning "This script appears to be an unsigned test build. Skipping version check."
    } else {
        try {
            $versionData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequestWithProxyDetection -Uri $VersionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
            $latestVersion = ($versionData | Where-Object { $_.File -eq $scriptName }).Version
            $result.LatestVersion = $latestVersion
            if ($null -ne $latestVersion) {
                $result.UpdateFound = ($latestVersion -ne $BuildVersion)
            } else {
                Write-Warning ("Unable to check for a script update as no script with the same name was found." +
                    "`r`nThis can happen if the script has been renamed. Please check manually if there is a newer version of the script.")
            }

            Write-Verbose "Current version: $($result.CurrentVersion) Latest version: $($result.LatestVersion) Update found: $($result.UpdateFound)"
        } catch {
            Write-Verbose "Unable to check for updates: $($_.Exception)"
            $result.Error = $_
        }
    }

    return $result
}


function Confirm-Signature {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $File
    )

    $IsValid = $false
    $MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
    $MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

    try {
        $sig = Get-AuthenticodeSignature -FilePath $File

        if ($sig.Status -ne 'Valid') {
            Write-Warning "Signature is not trusted by machine as Valid, status: $($sig.Status)."
            throw
        }

        $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

        if (-not $chain.Build($sig.SignerCertificate)) {
            Write-Warning "Signer certificate doesn't chain correctly."
            throw
        }

        if ($chain.ChainElements.Count -le 1) {
            Write-Warning "Certificate Chain shorter than expected."
            throw
        }

        $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

        if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
            Write-Warning "Top-level certificate in chain is not a root certificate."
            throw
        }

        if ($rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2010 -and $rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2011) {
            Write-Warning "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)."
            throw
        }

        Write-Host "File signed by $($sig.SignerCertificate.Subject)"

        $IsValid = $true
    } catch {
        $IsValid = $false
    }

    $IsValid
}

<#
.SYNOPSIS
    Overwrites the current running script file with the latest version from the repository.
.NOTES
    This function always overwrites the current file with the latest file, which might be
    the same. Get-ScriptUpdateAvailable should be called first to determine if an update is
    needed.

    In many situations, updates are expected to fail, because the server running the script
    does not have internet access. This function writes out failures as warnings, because we
    expect that Get-ScriptUpdateAvailable was already called and it successfully reached out
    to the internet.
#>
function Invoke-ScriptUpdate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([boolean])]
    param ()

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $oldName = [IO.Path]::GetFileNameWithoutExtension($scriptName) + ".old"
    $oldFullName = (Join-Path $scriptPath $oldName)
    $tempFullName = (Join-Path ((Get-Item $env:TEMP).FullName) $scriptName)

    if ($PSCmdlet.ShouldProcess("$scriptName", "Update script to latest version")) {
        try {
            Invoke-WebRequestWithProxyDetection -Uri "https://github.com/microsoft/CSS-Exchange/releases/latest/download/$scriptName" -OutFile $tempFullName
        } catch {
            Write-Warning "AutoUpdate: Failed to download update: $($_.Exception.Message)"
            return $false
        }

        try {
            if (Confirm-Signature -File $tempFullName) {
                Write-Host "AutoUpdate: Signature validated."
                if (Test-Path $oldFullName) {
                    Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                }
                Move-Item $scriptFullName $oldFullName
                Move-Item $tempFullName $scriptFullName
                Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                Write-Host "AutoUpdate: Succeeded."
                return $true
            } else {
                Write-Warning "AutoUpdate: Signature could not be verified: $tempFullName."
                Write-Warning "AutoUpdate: Update was not applied."
            }
        } catch {
            Write-Warning "AutoUpdate: Failed to apply update: $($_.Exception.Message)"
        }
    }

    return $false
}

<#
    Determines if the script has an update available. Use the optional
    -AutoUpdate switch to make it update itself. Pass -Confirm:$false
    to update without prompting the user. Pass -Verbose for additional
    diagnostic output.

    Returns $true if an update was downloaded, $false otherwise. The
    result will always be $false if the -AutoUpdate switch is not used.
#>
function Test-ScriptVersion {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Need to pass through ShouldProcess settings to Invoke-ScriptUpdate')]
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $false)]
        [switch]
        $AutoUpdate,
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $updateInfo = Get-ScriptUpdateAvailable $VersionsUrl
    if ($updateInfo.UpdateFound) {
        if ($AutoUpdate) {
            return Invoke-ScriptUpdate
        } else {
            Write-Warning "$($updateInfo.ScriptName) $BuildVersion is outdated. Please download the latest, version $($updateInfo.LatestVersion)."
        }
    }

    return $false
}

$BuildVersion = "26.06.06.1238"
Write-Host ("$($script:MyInvocation.MyCommand.Name) script version $($BuildVersion)") -ForegroundColor Green

$scriptVersionParams = @{
    AutoUpdate = $true
    Confirm    = $false
}

# This needs to be set prior to injecting this file to other scripts.
if (-not ([string]::IsNullOrEmpty($versionsUrl))) {
    $scriptVersionParams.Add("VersionsUrl", $versionsUrl)
}

if ($ScriptUpdateOnly) {
    switch (Test-ScriptVersion @scriptVersionParams) {
        ($true) { Write-Host ("Script was successfully updated") -ForegroundColor Green }
        ($false) { Write-Host ("No update of the script performed") -ForegroundColor Yellow }
        default { Write-Host ("Unable to perform ScriptUpdateOnly operation") -ForegroundColor Red }
    }
    exit
}

if ((-not($SkipVersionCheck)) -and
    (Test-ScriptVersion @scriptVersionParams)) {
    Write-Host ("Script was updated. Please re-run the command") -ForegroundColor Yellow
    exit
}
} end {
    $Script:countOrgRelIssues = (0)
    $Script:WebServicesVirtualDirectory = $null
    $Script:Server = hostname
    $Script:startingDate = (Get-Date -Format yyyyMMdd_HHmmss)
    $Script:htmlFile = "$PSScriptRoot\FBCheckerOutput_$($Script:startingDate).html"

    loadingParameters
    #Parameter input

    if (-not $OnlineUser) {
        $Script:UserOnline = Get-RemoteMailbox -ResultSize 1 -WarningAction SilentlyContinue
        $Script:UserOnline = $Script:UserOnline.RemoteRoutingAddress.SmtpAddress
    } else {
        $Script:UserOnline = Get-RemoteMailbox $OnlineUser -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        $Script:UserOnline = $Script:UserOnline.RemoteRoutingAddress.SmtpAddress
    }

    $Script:ExchangeOnlineDomain = ($Script:UserOnline -split "@")[1]

    if ($Script:ExchangeOnlineDomain -like "*.mail.onmicrosoft.com") {
        $Script:ExchangeOnlineAltDomain = (($Script:ExchangeOnlineDomain.Split(".")))[0] + ".onmicrosoft.com"
    } else {
        $Script:ExchangeOnlineAltDomain = (($Script:ExchangeOnlineDomain.Split(".")))[0] + ".mail.onmicrosoft.com"
    }
    $Script:temp = "*" + $Script:ExchangeOnlineDomain
    $Script:UserOnPrem = ""
    if (-not  $OnPremisesUser) {
        $Script:UserOnPrem = Get-mailbox -ResultSize 2 -WarningAction SilentlyContinue -Filter 'EmailAddresses -like $temp -and HiddenFromAddressListsEnabled -eq $false' -ErrorAction SilentlyContinue
        if ($Script:UserOnPrem) {
            $Script:UserOnPrem = $Script:UserOnPrem[1].PrimarySmtpAddress.Address
        }
    } else {
        $Script:UserOnPrem = Get-mailbox $OnPremisesUser -WarningAction SilentlyContinue -Filter 'EmailAddresses -like $temp -and HiddenFromAddressListsEnabled -eq $false' -ErrorAction SilentlyContinue
        $Script:UserOnPrem = $Script:UserOnPrem.PrimarySmtpAddress.Address
    }
    $Script:ExchangeOnPremDomain = ($Script:UserOnPrem -split "@")[1]

    if (-not $OnPremEWSUrl) {
        FetchEWSInformation
    } else {
        FetchEWSInformation
        $Script:ExchangeOnPremEWS = ($OnPremEWSUrl)
    }

    if (-not $OnPremDomain) {
        $ADDomain = Get-ADDomain
        $Script:ExchangeOnPremLocalDomain = $ADDomain.forest
    } else {
        $Script:ExchangeOnPremLocalDomain = $OnPremDomain
    }

    $Script:ExchangeOnPremLocalDomain = $ADDomain.forest
    if ([string]::IsNullOrWhitespace($ADDomain)) {
        $Script:ExchangeOnPremLocalDomain = $Script:ExchangeOnPremDomain
    }

    if ($Script:ExchangeOnPremDomain) {
        $Script:FedInfoEOP = Get-federationInformation -DomainName $Script:ExchangeOnPremDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select-Object *
    }
    #endregion

    if ($Help) {
        PrintDynamicWidthLine
        ShowHelp
        PrintDynamicWidthLine
        exit
    }
    #region Show Parameters
    $Script:IntraOrgCon = Get-IntraOrganizationConnector -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Where-Object { $_.TargetAddressDomains -contains $Script:ExchangeOnlineDomain } | Select-Object Name, TarGetAddressDomains, DiscoveryEndpoint, Enabled
    ShowParameters
    CheckParameters
    if ($Script:IntraOrgCon.enabled -eq $true) {
        $Auth = hostOutputIntraOrgConEnabled($Auth)
    }
    if ($Script:IntraOrgCon.enabled -eq $false) {
        hostOutputIntraOrgConNotEnabled
    }
    # Free busy Lookup methods
    PrintDynamicWidthLine
    $Script:OrgRel = Get-OrganizationRelationship | Where-Object { ($_.DomainNames -like $Script:ExchangeOnlineDomain) }  -WarningAction SilentlyContinue -ErrorAction SilentlyContinue  | Select-Object Enabled, Identity, DomainNames, FreeBusy*, TarGet*
    $Script:EDiscoveryEndpoint = Get-IntraOrganizationConfiguration -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object OnPremiseDiscoveryEndpoint
    $Script:SPDomainsOnprem = Get-SharingPolicy -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Format-List Domains
    $Script:SPOnprem = Get-SharingPolicy  -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object *

    if ($Org -contains 'ExchangeOnPremise' -or -not $Org) {
        #region DAuth Checks
        if ($Auth -like "DAuth" -or -not $Auth -or $Auth -like "All") {
            Write-Host "  Testing DAuth Configuration"
            OrgRelCheck -OrgRelParameter $Script:OrgRel
            PrintDynamicWidthLine
            FedInfoCheck
            FedTrustCheck
            AutoDVirtualDCheck
            PrintDynamicWidthLine
            EWSVirtualDirectoryCheck
            AvailabilityAddressSpaceCheck
            TestFedTrust
            TestOrgRel
        }
        #endregion
        #region OAuth Check
        if ($Auth -like "OAuth" -or -not $Auth -or $Auth -like "All") {
            Write-Host "  Testing OAuth Configuration"
            IntraOrgConCheck
            PrintDynamicWidthLine
            AuthServerCheck
            PrintDynamicWidthLine
            PartnerApplicationCheck
            PrintDynamicWidthLine
            ApplicationAccountCheck
            PrintDynamicWidthLine
            ManagementRoleAssignmentCheck
            PrintDynamicWidthLine
            AuthConfigCheck
            PrintDynamicWidthLine
            CurrentCertificateThumbprintCheck
            PrintDynamicWidthLine
            AutoDVirtualDCheckOAuth
            PrintDynamicWidthLine
            EWSVirtualDirectoryCheckOAuth
            PrintDynamicWidthLine
            AvailabilityAddressSpaceCheckOAuth
            PrintDynamicWidthLine
            OAuthConnectivityCheck
            PrintDynamicWidthLine
        }
        #endregion
    }
    # EXO Part
    if ($Org -contains 'ExchangeOnline' -or -not $Org) {
        #region ConnectExo
        $Exo = Test-ExchangeOnlineConnection
        if (-not ($Exo)) {
            Write-Host -ForegroundColor Red "`n Please connect to Exchange Online Using the EXO V3 module using EO as connection Prefix to collect Exchange OnLine Free Busy configuration Information."
            Write-Host -ForegroundColor Cyan "`n`n   Example: PS C:\Connect-ExchangeOnline -Prefix EO"
            Write-Host -ForegroundColor Yellow "`n   More Info at:https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps"
            exit
        }
        Write-Host " Connected to Exchange Online."
        $Script:ExoOrgRel = Get-EOOrganizationRelationship | Where-Object { ($_.DomainNames -like $Script:ExchangeOnPremDomain ) } | Select-Object Enabled, Identity, DomainNames, FreeBusy*, TarGet*
        $Script:ExoIntraOrgCon = Get-EOIntraOrganizationConnector | Select-Object Name, TarGetAddressDomains, DiscoveryEndpoint, Enabled
        $Script:tarGetAddressPr1 = ("https://AutoDiscover." + $Script:ExchangeOnPremDomain + "/AutoDiscover/AutoDiscover.svc/WSSecurity")
        $Script:tarGetAddressPr2 = ("https://" + $Script:ExchangeOnPremDomain + "/AutoDiscover/AutoDiscover.svc/WSSecurity")
        exoHeaderHtml

        #endregion

        #region ExoDAuthCheck
        if ($Auth -like "DAuth" -or -not $Auth -or $Auth -like "All") {
            PrintDynamicWidthLine
            Write-Host $TestingExoDAuthConfiguration
            ExoOrgRelCheck
            PrintDynamicWidthLine
            ExoFedOrgIdCheck
            PrintDynamicWidthLine
            ExoTestOrgRelCheck
            SharingPolicyCheck
        }
        #endregion

        #region ExoOauthCheck
        if ($Auth -like "OAuth" -or -not $Auth -or $Auth -like "All") {
            Write-Host $TestingExoOAuthConfiguration
            ExoIntraOrgConCheck
            PrintDynamicWidthLine
            EXOIntraOrgConfigCheck
            PrintDynamicWidthLine
            EXOAuthServerCheck
            PrintDynamicWidthLine
            ExoTestOAuthCheck
            PrintDynamicWidthLine
        }
        #endregion

        Write-Host -ForegroundColor Green $ThatIsAllForTheExchangeOnlineSide

        PrintDynamicWidthLine
    }

    Stop-Transcript
}

# SIG # Begin signature block
# MIInQQYJKoZIhvcNAQcCoIInMjCCJy4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgATzSOcWLXqj7
# 2ZemTp9ZYPb0pbWZNrzHGjnNl+VNUKCCDLowggX1MIID3aADAgECAhMzAAACHU0Z
# yE7XD1dIAAAAAAIdMA0GCSqGSIb3DQEBCwUAMFcxCzAJBgNVBAYTAlVTMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBD
# b2RlIFNpZ25pbmcgUENBIDIwMjQwHhcNMjYwNDE2MTg1OTQzWhcNMjcwNDE1MTg1
# OTQzWjB0MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYD
# VQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQDQvewXxx9gZZFC6Ys1WBay8BJ8kGA4JQnH5CMafqOASlTpK9H8
# o5ZXTXt0caVQTNMUPt445wXYD+dFtaKWTwDn1I52oUSrC9vJin1Gsqt+zyKJL5Dg
# 3eQXbQNR61DmMy20GLTIO3SFed9Rfi/ophgCLGFLDR3r0KvHjwMb/jYWS0celV/4
# Lz27LfAekm8v9E5IXaeiXbAUYZKK090n4CVl3JBtbN+9DtI9SNu/yjvozW52/u7R
# X/Ttpa/KDlpuokZ+Zcbvmtd9ur9gFLvZzh41o9MsE/clQtdaFWGvuo6Jua/ntpgk
# ey3E5/vBFe+MJPG6phdnuo6r57ZudCudiI1bAgMBAAGjggGbMIIBlzAOBgNVHQ8B
# Af8EBAMCB4AwHwYDVR0lBBgwFgYKKwYBBAGCN0wIAQYIKwYBBQUHAwMwHQYDVR0O
# BBYEFH6QuMwqcPG0hQlQ6c5jCtTTLrVeMEUGA1UdEQQ+MDykOjA4MR4wHAYDVQQL
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xFjAUBgNVBAUTDTIzMDAxMis1MDc1NTkw
# HwYDVR0jBBgwFoAUf1k/VCHarU/vBeXmo9ctBpQSCDEwYAYDVR0fBFkwVzBVoFOg
# UYZPaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwQ29kZSUyMFNpZ25pbmclMjBQQ0ElMjAyMDI0LmNybDBtBggrBgEFBQcBAQRh
# MF8wXQYIKwYBBQUHMAKGUWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljcm9zb2Z0JTIwQ29kZSUyMFNpZ25pbmclMjBQQ0ElMjAyMDI0LmNy
# dDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBKTbYOjzwTG/DXGaz9
# s6+fQeaTtDcFmMY+5UyVFCyj7Pv+5i37qfX8lSL/tBIfYQfWsMuBQlfZurJD6r4H
# VJ2CeH+1fgiq8dcHdVKoZ3Sa2qXoX3cq9iS8cVb06B7+5/XJ7I0OxHH9fDsvJ3T3
# w5V/ZtAIFmLrl+P0CtG+92uzRsn0nTbdFjOkLMLWPLAU3THohKRlSEMgFJpPkm5n
# 5UAZ35xX6FWCrDLsSKb555bTifwa8mJBwdlof0bmfYidH+dxZ1FdDxvLnNl9zeKs
# A4kejaaIqqIPguhwAti5Ql7BlTNoJNwxCvBmqW2MQLnCkYN/VVUsR3V2x/rcTNzo
# Bf/Z/SpROvdaA2ZOOd1uioXJt3tdLQ7vHpqpib0KfWr/FWXW10q38VxfCnRQBqzb
# SuztR7nEMuzX7Ck+B/XaPDXd1qh72+QYyB0Z2VzWmO9zsnb9Uq/dwu8LGeQqnyu6
# 7SDGACvnXii2fb9+US492VTnXSnFKyqwgzUyFMtZK1/sHYTv6bG4TtQUygQxTN+Z
# V+aJIlKO2MqZ7bKrAnOzS9m6NgoTdWOq11bTOZwKlIEV/EhV9SWkDmdpR/hPPT2v
# 6TEj4F8PT/zHjRezIU5c/DGlt/VhY/pK0XkJtEyMmmS1BMtjU/rqBZVMIm3dnxQs
# /TBByr+Cf8Z1r7aifQVQ+WSqzjCCBr0wggSloAMCAQICEzMAAAA5O7Y3Gb8GHWcA
# AAAAADkwDQYJKoZIhvcNAQEMBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRl
# IEF1dGhvcml0eSAyMDExMB4XDTI0MDgwODIwNTQxOFoXDTM2MDMyMjIyMTMwNFow
# VzELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEo
# MCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAyNDCCAiIwDQYJ
# KoZIhvcNAQEBBQADggIPADCCAgoCggIBANgBnB7jOMeqlRYHNa265v4IY9fH8TKh
# emHfPINe1gpLaV3dhg324WwH06LcHbpnsBukCDNitryo0dtS/EW6I/yEL/bLSY8h
# KpbfQuWusBPr9qazYcDxCW/qnjb5JsI1s8bNOg3bVATvQVL4tcf03aTycsz8QeCd
# M0l/yHRObJ9QqazM1r6VPEOJ7LL+uEEb73w6QCuhs89a1uv1zerOYMnsneRRwCbp
# yW11IcggU0cRKDDq1pjVJzIbIF6+oiXXbReOsgeI8zu1FyQfK0fVkaya8SmVHQ/t
# Of23mZ4W9k0Ri22QW9p3UgSC5OUDktKxxcCmGL6tXLfOGSWHIIV4YrTJTT6PNty5
# REojHJuZHArkF9VnHTERWoTjAzfI3kP+5b4alUdhgAZ7ttOu1bVnXfHaqPYl2rPs
# 20ji03LOVWsh/radgE17es5hL+t6lV0eVHrVhsssROWJuz2MXMCt7iw7lFPG9LXK
# Gjsmonn2gotGdHIuEg5JnJMJVmixd5LRlkmgYRZKzhxSCwyoGIq0PhaA7Y+VPct5
# pCHkijcIIDm0nlkK+0KyepolcqGm0T/GYQRMhHJlGOOmVQop36wUVUYklUy++vDW
# eEgEo4s7hxN6mIbf2MSIQ/iIfMZgJxC69oukMUXCrOC3SkE/xIkgpfl22MM1itkZ
# 35nNXkMolU1lAgMBAAGjggFOMIIBSjAOBgNVHQ8BAf8EBAMCAYYwEAYJKwYBBAGC
# NxUBBAMCAQAwHQYDVR0OBBYEFH9ZP1Qh2q1P7wXl5qPXLQaUEggxMBkGCSsGAQQB
# gjcUAgQMHgoAUwB1AGIAQwBBMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# ci06AjGQQ7kUBU7h6qfHMdEjiTQwWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0MjAx
# MV8yMDExXzAzXzIyLmNybDBeBggrBgEFBQcBAQRSMFAwTgYIKwYBBQUHMAKGQmh0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0MjAx
# MV8yMDExXzAzXzIyLmNydDANBgkqhkiG9w0BAQwFAAOCAgEAFJQfOChP7onn6fLI
# MKrSlN1WYKwDFgAddymOUO3FrM8d7B/W/iQ6DxXsDn7D5W4wMwYeLystcEqfkjz4
# NURRgazyMu5yRzQh4LqjA4tStTcJh1opExo7nn5PuPBYnbu0+THSuVHTe0VTTPVh
# ily/piFrDo3axQ9P4C+Ol5yet+2gTfekICS5xS+cYfSIvgn0JksVBVMYVI5QFu/q
# hnLhsEFEUzG8fvv0hjgkO+lkpV9ty6GkN4vdnd7ya6Q6aR9y34aiM1qmxaxBi6OU
# nyNl6fkuun/diTFnYDLTppOkr/mg5WSfCiDVMNCxtj4wPKC5OmHm1DQIt/MNokbb
# H3UGsFP1QbzsLocuSqLCvH09Io3fDPTmscR9Y75G4qX7RTX8AdBPo0I6OEojf39z
# uFZt0qOHm65YWQE69cZM2ueE1MB05dNNgHK9gTE7zKvK/fg8B2qjW88MT/WF5V5u
# vZGtqa9FSL2RazArA+rDPuf6JGYz4HpgMZHB4S6szWSKYBv0VisCzfxgeU+dquXW
# 9bd0auYlOB58DPcOYKdc3Se94g+xL4pcEhbB54JOgAkwYTu/9dLeH2pDqeJZAABV
# DWRQCaXfO5LgyKwKCLYXpigrZYCjUSBcr+Ve8PFWMhVTQl0v4q8J/AUmQN5W4n10
# 1cY2L4A7GTQG1h32HHAvfQESWP0xghndMIIZ2QIBATBuMFcxCzAJBgNVBAYTAlVT
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jv
# c29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMjQCEzMAAAIdTRnITtcPV0gAAAAAAh0w
# DQYJYIZIAWUDBAIBBQCggZAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJ
# KoZIhvcNAQkEMSIEIF2w+8bFRM5vP+OrCz/crHnxjpVKKUtHFATRxwdO1J9OMEIG
# CisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEBBQAEggEAm/rVfs6LZ6flTYGk
# Ag329pySlKaaRL+z69RNTym0dQvLLTVHgxLIOKlyxZzhD5Gtj2xjvm1Sm++GKoZU
# 0ySFFq0GhPXDT3GpbWjrAL83coGAt69RXXmTMhCajq7qH5hmQJORD9vocrBs05Tz
# RASOO/h9lEDD8tVt38yW7aDvFm5WC25Sf9z5D63M32lOq0Lo0z3SNBjEzg1QJjTo
# lXsWJvp4GwAjvEfnP7BTJT3S21fv2sKjdJz23pdwWnAjl8QQ2mnCiRYHfGMvPajR
# uJ4AxhPBIs6s3q3m7MXXiidvBzH3dSIcHHS4RABZ4e00sZ6bQWcy6FTJSN6GbHS7
# NWs88aGCF60wghepBgorBgEEAYI3AwMBMYIXmTCCF5UGCSqGSIb3DQEHAqCCF4Yw
# gheCAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFaBgsqhkiG9w0BCRABBKCCAUkEggFF
# MIIBQQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCA6HZ45WXsI2AgI
# GJhRvH/JBDCpX1OnsCyrW6pd1SiLYgIGaq6yEk7BGBMyMDI2MTAxOTE4MDQyMy4z
# MTdaMASAAgH0oIHZpIHWMIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExp
# bWl0ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo0MzFBLTA1RTAtRDk0NzEl
# MCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCCEfswggcoMIIF
# EKADAgECAhMzAAACHUvAkoc4hX45AAEAAAIdMA0GCSqGSIb3DQEBCwUAMHwxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDgxNDE4NDgzM1oXDTI2MTEx
# MzE4NDgzM1owgdMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# LTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEn
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjQzMUEtMDVFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEF
# AAOCAg8AMIICCgKCAgEAorSgaAA8oOl4ph574zw29egUN8DDepRHLX8FM1zHNJmX
# G6KrSqUKwzcKafopuYdPTETTCvb9aJfESuAU0iGNUFI/D6R0kvdfpe2oPX+E3sbT
# QvGi4JPH5qdIYUaJ45V/4bqe8eNvbWzpC+ZKjH193DeiI1XAI918JoQmBhlEXo/T
# on1721luZJgincsf5LjMY3jX84WyXUSX3dsS7h/7xVI+w1yjg7pa+0y3o/me2Tsv
# 6UJUdSTQap5ORGSfCnclnP1z3IiiWIWr3Vo7aIPWsgJzq3m5GxpxUHCQk8qzUhk5
# 0y/uB+LGE3WIK2C77iy9iFsSfSLUnyMEzGRDW9mXHT4PH7Ozz6CHqQEiNvwcHqlv
# lCh1pHQh1NXQSAqOoVBs5mi6easf6yxWTfe5DrR79503r8pU6VqC2Y9XMRU4wH9Q
# bYXYsIUZ33Jmndy22W1LBDAbxBPQHCBlncGDU3BgdhVUVLe80mggFO98FdkWho67
# w4kPdCTRkvdvkY8PrQYE/nQjHXCa0g7LcMttZb6ejMHfQ+tUWXv6+nZ4Ynkr2Oka
# xclFCw4RIYNMWD26AWbQj/WEdzga18fKtw66L5gzXPza6jFBfPJeKE3H8QAuwpir
# mH4ms+5nUjNNQOmNgqJn0U1+3Yn7ClswD79YN0r3fdbYBMDApBZJpNlK7q7HXRsC
# AwEAAaOCAUkwggFFMB0GA1UdDgQWBBSEWfBxNEamZtXm8gl92Yq80jfxXTAfBgNV
# HSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5o
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBU
# aW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwG
# CCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRz
# L01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNV
# HRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIH
# gDANBgkqhkiG9w0BAQsFAAOCAgEAkdweB4yxvLspLKq0D+miyD4Q0EcxVFpNZuJx
# iR54gWRkeTDDuymNeB03JhlsBpbwSYJ5uZSgDBCvwHED2VL8lJpFlOprJzxsXWC2
# NTfA+O+PO5Fk5jw6LHh6jeBADDEdQAx3Hqi7Zm0JwvQ93z5f6dtxkm29WqOcHYXR
# XfAQwy1hSrLXyfeblqR66jpP/9n0fCkWU4ggsUjQpQ2Ngj1DV09J4Y3y7p9Nd81+
# Xs6qYo++7RKm8qiB/5NDeigOLjlAeFgiEXIRUJW+mJyqpQw+OORlaqcFjR8Hu0G+
# /7bMdek68YX+kPpDBk7Ue+I/xgiYJ1xcDRBn/vczLtN72+RIlD4UgXYLuBSCk//p
# DEPX5z39Cr+rkc6E4Y28FPk4BhloAyvp628P4xfElQY8TcxraUbZShypocE6ny95
# D1K1BkltZmrHVKCxmglnuOlM15NKIrXFlXCzdqpCtIwQ417wNAVF/QDPvzzbumPd
# Ti6fb0tLbScYobV6zvbBsMsKEME4Tj1b9oIXC8dybJq4nbboEXYpRwi1QAbpSNrn
# +PxGW9uf1q63FnMJu4gm3Oh63njW/iVf723quzyHrSijWMgY0HiRiHQi0Jyu0h8M
# dhRUp7mxbmLQckPiOFwAlIaUN/k725y/aLWpkRU6fqmLlEOyH5WpyLd23AYy9r8v
# +Qoba6swggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqGSIb3
# DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIw
# MAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAx
# MDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# 5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4XYDP/
# XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTzxXb1
# hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7uhp7
# M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlwaQ3K
# Ni1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedGbsoy
# 1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGaRnXNxF80
# 3RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03dJQc
# NIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG4tG9ahha
# YQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5UPkL
# iWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReTwDwV
# 2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZMBIG
# CSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8RjUp
# zxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAEVTBT
# MFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAwwCgYI
# KwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGG
# MA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186a
# GMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3Br
# aS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsG
# AQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZIhvcN
# AQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9MTO1
# OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2LpypglYA
# A7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OOPcbz
# aN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8DJ6L
# GYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA0j3m
# Sj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1RtnWN0
# SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjcZxko
# JLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq77EFm
# PWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQKBqJC482
# 2rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328y+l7
# vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIDVjCC
# Aj4CAQEwggEBoYHZpIHWMIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExp
# bWl0ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo0MzFBLTA1RTAtRDk0NzEl
# MCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsO
# AwIaAxUAuoO+BKbfXzqyfi9GLEdWHkCLeT+ggYMwgYCkfjB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQsFAAIFAO6AujUwIhgPMjAyNjEw
# MTkxNTQ4MDVaGA8yMDI2MTAyMDE1NDgwNVowdDA6BgorBgEEAYRZCgQBMSwwKjAK
# AgUA7oC6NQIBADAHAgEAAgIKXzAHAgEAAgISkTAKAgUA7oILtQIBADA2BgorBgEE
# AYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYag
# MA0GCSqGSIb3DQEBCwUAA4IBAQAnj3gDcVbRZkFKnP25msZYDR/VjDKuxnvQjofr
# vp5rccty1/rZ2pNs3NOjuKPJI0IyoMC3D8pYPJD81bNMMQpAyJbjxDrSbKYL5juL
# y3sa1zqliG1qShUBZHjdUgS75FrGtbo6onJEWmw9UUwiRR3Fzx2tWuuZT4LcWwXu
# AvurnUp3MWu/6eDacoaJL2N1yadw0tegr7BNWIlAYbhU1sdPhE4hfEwdOg69EhxT
# vnSsMH1GW9RmMjFmz7JYS+6YltiDKxEdk8yNxCNhVDzkejkIopH489yOCIXz1yos
# 1ipXI/lAlQ6G2hj85ZabRxQJrMDrjjzdPabXAT8ln7Yu6uohMYIEDTCCBAkCAQEw
# gZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAIdS8CShziFfjkA
# AQAAAh0wDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0B
# CRABBDAvBgkqhkiG9w0BCQQxIgQgT8KNWWNHQZH59xxegkdQ+/HQYjFJA0Ls/jFo
# 2aCpjw4wgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCCxtpXMXEiLJzrqM77e
# p4rTNwrMOj6gpWN9hZvpj5QFUTCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwAhMzAAACHUvAkoc4hX45AAEAAAIdMCIEIL3VJ+LUes0XJrMHy/JF
# 7jn9sjqGbM3GTzubr1PEMmggMA0GCSqGSIb3DQEBCwUABIICABXsANP6rhMTQ6kN
# RxL8EaLwfZaYDQybl//DrhPKyOGXh+4lStDoSM6g2bro+dIyHzIl/GLABbd3BdcB
# 9fDxdOHAZPX5eiC1Ej9lHbW1Rbd50KYRy3OiI3Zkg7CSUI1D1vGi0ZOkIl8Hskk8
# OMibItkun7Sa8Xuye9tSa2lpQQ1M/dEN3Y0VYrshW/3P4RJgHshPuafWt24xNUBu
# jpMooeALlXcfqEoBWg8sem8Ou0kk6g7bAgjUJEGskHN+PisyaNq9RfIWzPToke/+
# foRHAVYA8mErXTmMUC1NFV4v7IDcH3T2WdIBaTcZH5yuH/S9CzHxoHpF8f8G/njZ
# aJP8i/N+/hnozbxruiIcjth/XylLaiTo+JRjL+7nOUOIY8JFnxP1niL1nnPU9HWS
# iX9i+ysxGo7zCcOmYbAnwUliz1T6zaMzXnUBQemv1kLySkiZfy1UnEBrLNaL00g1
# O+VltoysLjMHBgqzGwHIm5z8rNb3xjKLeG9lAmYsWRk6yCet+8BrptHQFJUq59Wb
# BBuz2uYkbe94pbtvHyRsojzm9s+M0LJuyqCxOuT1S34MzREjRtzhkkXa8lkL3SZ6
# ytpd9SUHLTnKv1bsTd05O02J5mMD4odiUULZz6MtZH4vQznpJSyOx1R3NSCVUyrv
# Fw6ycjjMEWd72w51ME4/uUkf/WuT
# SIG # End signature block
