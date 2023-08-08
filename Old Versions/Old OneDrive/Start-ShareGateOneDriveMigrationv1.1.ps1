<#
.SYNOPSIS
A script for managing OneDrive migrations and pre-stage operations using Sharegate PowerShell module.

.DESCRIPTION
This script is designed to assist with OneDrive migrations and pre-stage operations using the Sharegate PowerShell module. It provides several options, including migrating data, running incremental syncs, pre-staging OneDrive, and adding secondary admins to OneDrive sites. The script supports using email addresses or URLs for the source and destination sites.

.PARAMETER RequestOneDrive
Indicates whether to pre-stage OneDrive. Use this switch if you want to request OneDrive site creation for a user.

.PARAMETER AddSecondaryAdmin
Indicates whether to add a secondary admin to the destination OneDrive site.

.PARAMETER Migrate
Indicates whether to perform an Incremental Sync migration.

.PARAMETER Test
Indicates whether to run an test sync of OneDrive. Use this switch for test migration. Will always run Incremental.

.PARAMETER SourceAdminUrl
Specifies the Source Admin Site URL.

.PARAMETER DestinationAdminUrl
Specifies the Destination Admin Site URL.

.PARAMETER SourceUPN
Specifies the source email address or UPN.

.PARAMETER SrcSiteURL
Specifies the source site URL.

.PARAMETER DestinationUPN
Specifies the destination email address or UPN.

.PARAMETER DstSiteUrl
Specifies the destination site URL.

.EXAMPLE
Start-ShareGateOneDriveMigration -DestinationAdminUrl $Global:DestinationAdminUrl -RequestOneDrive -DestinationUPN "user@destination.com"

This example will request OneDrive site creation for the specified user in the destination tenant.

.EXAMPLE
Start-ShareGateOneDriveMigration -DestinationAdminUrl $Global:DestinationAdminUrl -Migrate -CopyOperation Incremental -SourceAdminUrl "https://source-admin-url" -SourceUPN "user@source.com" -DestinationUPN "user@destination.com"

This example will perform an incremental migration of OneDrive content from the source site to the destination site using the provided email addresses and URLs.


.EXAMPLE
Start-ShareGateOneDriveMigration -DestinationAdminUrl $Global:DestinationAdminUrl-Migrate -CopyOperation Incremental -Test -SourceAdminUrl "https://source-admin-url"  -SourceUPN "user@source.com" -DestinationUPN "user@destination.com"

This example will perform a test incremental migration of OneDrive content from the source site to the destination site using the provided email addresseses.


.EXAMPLE
Start-ShareGateOneDriveMigration -DestinationAdminUrl $Global:DestinationAdminUrl -Migrate -CopyOperation Incremental -Test -SourceAdminUrl "https://source-admin-url"  -SrcSiteURL "https://source-OneDrive-url" -DstSiteUrl "https://destination-OneDrive-url"

This example will perform a test incremental migration of OneDrive content from the source site to the destination site using the provided URLs.


.EXAMPLE
Start-ShareGateOneDriveMigration -DestinationAdminUrl $Global:DestinationAdminUrl -AddSecondaryAdmin -DestinationUPN "user@destination.com"

This example will add a secondary admin to the specified user's OneDrive site in the destination tenant.
#>

# Global Settings
#Set Up Module, Variables, Credentials, and Connect to SharePoint Sites
Import-Module Sharegate
$AllOneDriveErrors = @{}

# Destination URL for the new tenant
$Global:DestinationAdminUrl = Read-Host "What is the Destination Admin Site URL?"
#"https://oakview-admin.sharepoint.com/"
# Source URL for the old tenant
$Global:SourceAdminUrl = Read-Host "What is the Source Admin Site URL?"
#"https://spectraxp-admin.sharepoint.com/"

# Connect to source and destination tenants using ShareGate PowerShell module
$Global:SourceTenant = Connect-Site -Url $SourceAdminUrl -Browser
$Global:DestinationTenant = Connect-Site -Url $DestinationAdminUrl -Browser

# Prompt user for source and destination credentials
# Consider using a more secure method for storing and retrieving credentials
Write-Host "Enter Source Credentials" -ForegroundColor Yellow
$global:SourceCredentials = Get-Credential
Write-Host "Enter Destination Credentials" -ForegroundColor Yellow
$global:DestinationCredentials = Get-Credential

# Function to get OneDrive URL for a user from the source or destination tenant
function Get-OneDriveUrlForUser {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UPN,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Source', 'Destination')]
        [string]$TenantLocation
    )

    # Get the tenant details for the specified tenant
    if ($TenantLocation -eq "Source") {
        $TenantShareGateDetails = $Global:SourceTenant
        $TenantAdminUrl = $Global:SourceAdminUrl
        $TenantCredentials = $Global:SourceCredentials
    } elseif ($TenantLocation -eq "Destination") {
        $TenantShareGateDetails = $Global:DestinationTenant
        $TenantAdminUrl = $Global:DestinationAdminUrl
        $TenantCredentials = $Global:DestinationCredentials
    }

    try {
        # Look up OneDrive URL using sharegate method
        $OneDriveUrlCheck = Get-OneDriveUrl -Tenant $TenantShareGateDetails -Email $UPN -ErrorAction Stop
        $OneDriveUrlCheck = $OneDriveUrlCheck.TrimEnd('/')
        Write-Host "$($TenantLocation) $($UPN): Found - " -ForegroundColor Green -NoNewline
        Write-Host "OneDrive $($OneDriveUrlCheck)" -ForegroundColor Cyan
        return $OneDriveUrlCheck
    } catch {
        try {
            # Look up OneDrive URL using the sharepoint online method
            Connect-SPOService -Url $TenantAdminUrl -Credential $TenantCredentials
            $OneDriveUrlCheck = Get-SPOSite -Filter "Owner -eq '$UPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ErrorAction SilentlyContinue
            $OneDriveUrlCheck = $OneDriveUrlCheck.TrimEnd('/')
            Write-Host "$($TenantLocation) $($UPN): Found - " -ForegroundColor Green -NoNewline
            Write-Host "OneDrive $($OneDriveUrlCheck)" -ForegroundColor Cyan
            return $OneDriveUrlCheck
        } catch {
            Write-Host "$($TenantLocation) OneDrive URL not found - $($UPN)" -ForegroundColor Yellow
            $AllOneDriveErrors[$UPN] = [PSCustomObject]@{
                Commandlet = "Get-OneDriveUrlForUser"
                FailureActivity = "UnableToFindOneDrive"
                Tenant = $TenantLocation
                User = $UPN
                Error = ($_.Exception)
            }

            #Request OneDrive
            if ($global:RequestOneDrive) {
                Request-SPOPersonalSite -UserEmails $DestinationUPN -ErrorAction Stop
                Write-Host "Destination OneDrive Site Requested for $($DestinationUPN)" -ForegroundColor Green
                $RequestedSite += $User
            }
            continue
        }
    }
    
}

# Function to Connect OneDrive URL for a user; if not able to connect, add the admin as a site admin
function Connect-OneDriveForUser {
    param (
        [Parameter(Mandatory=$true)]
        [string]$OneDriveURL,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Source', 'Destination')]
        [string]$TenantLocation
    )

    # Get the tenant details for the specified tenant
    if ($TenantLocation -eq "Source") {
        $TenantShareGateDetails = $Global:SourceTenant
        $TenantAdminUrl = $Global:SourceAdminUrl
        $TenantCredentials = $Global:SourceCredentials
    } elseif ($TenantLocation -eq "Destination") {
        $TenantShareGateDetails = $Global:DestinationTenant
        $TenantAdminUrl = $Global:DestinationAdminUrl
        $TenantCredentials = $Global:DestinationCredentials
    }
    
    #Attempt Connection to OneDrive Site
    try {
        $SiteDetails = Connect-Site -Url $OneDriveURL -UseCredentialsFrom $TenantShareGateDetails -ErrorAction Stop
        # Get Source OneDrive Documents
        $DocumentsLibrary = Get-List -Site $SiteDetails -Name "Documents" -ErrorAction Stop
        Return $DocumentsLibrary
    }
    catch {    
        try {
            $AdminCheck = Get-SPOUser -Site $OneDriveURL.ToString() -ErrorAction Stop
        }
        catch {
            #If failed, switch to TenantLocation Tenant
            Write-Host ".. switching to $($TenantLocation) Tenant... " -foregroundcolor Yellow -NoNewline
            Connect-SPOService $TenantAdminUrl -ModernAuth
        }
        $rootSiteURL = Get-SPOSite -limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
        $rootSiteURL = $rootSiteURL.url -replace '/sites.*', ''

        #add Destination Tenant Admin as Site Admin
        try {
            $result = Set-SPOUser -Site $OneDriveURL -LoginName $global:TenantCredentials.username.ToString() -IsSiteCollectionAdmin $true -ErrorAction Stop
            Write-Host "$($global:TenantCredentials.username.ToString()) Added as Site Admin." -ForegroundColor Green
        }
        catch {
            Write-Host "Unable to Add Admin for $($DestinationUPN)" -ForegroundColor Red
            $FailedToAddAdminToOneDrive += $User
            return $error[0]
        }
        
    }
}

function Start-ShareGateOneDriveMigration {
    [CmdletBinding(DefaultParameterSetName='Prestage')]
    Param(
        [Parameter(Mandatory=$false, HelpMessage="RequestOneDrive?")]
        [switch] $RequestOneDrive,
    
        [Parameter(ParameterSetName='Prestage', Mandatory=$false, HelpMessage="AddSecondaryAdmin?")]
        [switch] $AddSecondaryAdmin,
    
        [Parameter(ParameterSetName='Migrate', Mandatory=$false, HelpMessage="Migrate?")]
        [switch] $Migrate,
    
        [Parameter(ParameterSetName='Migrate', Mandatory=$false, HelpMessage="Run Test Incremental Sync of OneDrive?")]
        [switch] $Test,
    
        [Parameter(ParameterSetName='Migrate',Mandatory=$True, HelpMessage="What is the Source Admin Site URL")]
        [string] $SourceAdminUrl,
    
        [Parameter(Mandatory=$True, HelpMessage="What is the Destination Admin Site URL")]
        [string] $DestinationAdminUrl,

        [Parameter(ParameterSetName='Migrate', Mandatory=$false, HelpMessage="What is the source email address or UPN?")]
        [string] $SourceUPN,
    
        [Parameter(ParameterSetName='Migrate', Mandatory=$false, HelpMessage="What is the source site URL?")]
        [string] $SrcSiteUrl,
    
        [Parameter(Mandatory=$false, HelpMessage="What is the destination email address or UPN?")]
        [string] $DestinationUPN,
    
        [Parameter(Mandatory=$false, HelpMessage="What is the destination site URL?")]
        [string] $DstSiteUrl,

        [Parameter(ParameterSetName='Migrate',Mandatory = $true)]
        [ValidateSet('Incremental', 'Overwrite', 'Skip')]
        [string]$CopyOperation

    )
    $global:RequestOneDrive = $RequestOneDrive

    # Migration Copy Settings
    if ($CopyOperation -eq "Incremental") {
        $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
    }
    elseif ($CopyOperation -eq "Overwrite") {
        $copysettings = New-CopySettings -OnContentItemExists OverWrite
    }
    elseif ($CopyOperation -eq "Skip") {
        $copysettings = New-CopySettings -OnContentItemExists Skip
    }
    #New Heading
    Write-Host "$($DestinationUPN)" -ForegroundColor Black -BackgroundColor Cyan

    #Check Destination OneDrive URL
    if ($DstSiteUrl) {}
    else {
        $DstSiteUrl = Get-OneDriveUrlForUser -TenantLocation Destination -UPN $DestinationUPN -ErrorAction Stop
    }

    # Run OneDrive Migration
    if ($migrate) {
        #Write-Host "Migrate from $($SourceUPN))" -ForegroundColor Black -BackgroundColor Cyan
        # Connect to OneDrive Sites in Source
        if ($Global:SourceTenant -ne $null) {
            $SourceTenant = $Global:SourceTenant
        } else {
            $SourceTenant = Connect-Site -Url $SourceAdminUrl -Browser
        }
        ### Source - Start Region ###
        #Gather Source OneDrive URL
        if ($srcSiteUrl) {}
        else {
            try {
                $srcSiteUrl = Get-OneDriveUrlForUser -TenantLocation Source -UPN $SourceUPN -ErrorAction Stop
                $SourceDocumentLibrary = Connect-OneDriveForUser -OneDriveURL $srcSiteUrl -TenantLocation "Source"
                Write-Host "Connected to Source OneDrive .. " -ForegroundColor Green -NoNewline
            }
            catch {
                Write-Host "Unable to Connect to Source OneDrive. Skipping" -ForegroundColor Red
                Write-Host ""
                #Write-Host "Source $($SourceUPN): Not Connected" -ForegroundColor Red
                $AllOneDriveErrors[$SourceUPN] = [PSCustomObject]@{
                    Commandlet = $_.CategoryInfo.Activity
                    FailureActivity = "UnableToConnectOneDrive"
                    Tenant = $TenantShareGateDetails.Site
                    User = $SourceUPN
                    Error = ($_.Exception)
                }
                return
            }
        }
        ### Source - End Region ###
        
        ### Destination - Start Region ###

        # Pull Destination OneDrive and Migrate Only if Source is Found  
        # Connect to OneDrive Sites in Destination
        if ($Global:DestinationTenant -ne $null) {
            $DestinationTenant = $Global:DestinationTenant
        } else {
            $DestinationTenant = Connect-Site -Url $DestinationAdminUrl -Browser
        }                
        #Write-Host "Connected to Destination Tenant $($DestinationTenant.Site.toString())... " -NoNewline -BackgroundColor DarkGreen
        # Connect to OneDrive Sites in Destination
        $DestinationDocumentLibrary = Connect-OneDriveForUser -OneDriveURL $DstSiteUrl -TenantLocation "Destination"
        Write-Host "Connected to Destination OneDrive" -ForegroundColor Green
        ### Destination - End Region ###
            
        # Create a new object to store the results
        $OneDriveResults = New-Object PSObject
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceAdminTenantURL" -Value $sourceTenant.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationAdminTenantURL" -Value $destinationTenant.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceName" -Value $SourceDocumentLibrary.Title
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceSite" -Value $SourceDocumentLibrary.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationName" -Value $DestinationDocumentLibrary.Title
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationURL" -Value $DestinationDocumentLibrary.Address

        # Copy OneDrive Files from Source to Destination
        if ($DestinationDocumentLibrary) {
            if ($Test) {
                $TaskName = "Test OneDrive Migration $($SourceUPN) to $($DestinationDocumentLibrary.Title)"

                # Progress Bar Current 2
                Write-Progress -Id 2 -Activity "$($TaskName)"                
                # Test Move with Incremental using Insane Mode
                $Result = Copy-Content -SourceList $SourceDocumentLibrary -DestinationList $DestinationDocumentLibrary -InsaneMode -CopySettings $CopySettings -TaskName $TaskName -WarningAction SilentlyContinue -WhatIf
                $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value "Test-$CopyOperation"
            } else {
                $TaskName = "$($CopyOperation) OneDrive Migration $($SourceUPN) to $($DestinationUPN)"
                # Progress Bar Current 2
                Write-Progress -Id 2 -Activity "$($TaskName)"
                
                $Result = Copy-Content -SourceList $SourceDocumentLibrary -DestinationList $DestinationDocumentLibrary -InsaneMode -CopySettings $CopySettings -TaskName $TaskName -WarningAction SilentlyContinue
                $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value $CopyOperation
            }
        }
        # Add Job Result of Job 
        If ($Result) {
            Write-Progress -Id 2 -Activity "$($TaskName)" -Completed
            Write-Host "Migration Completed - $($TaskName)" -ForegroundColor Green
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result.Result
        } else {
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Result" -Value "Failed"
        }
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "ItemsCopied" -Value $Result.ItemsCopied
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Successes" -Value $Result.Successes
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Errors" -Value $Result.Errors
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Warnings" -Value $Result.Warnings
        $OneDriveResults | Format-List
    }
    
    #Check if OneDrive Admin Added
    if ($AddSecondaryAdmin) {
        try {
            if ($SPOPermUsers = (Get-SPOUser -Site $DstSiteUrl.ToString() -ErrorAction Stop).LoginName) {
            Write-Host "Already Site Admin for $($DstSiteUrl.ToString())" -ForegroundColor Yellow
            $AlreadySiteAdmin += $User
            }
        } catch {
            Connect-OneDriveForUser -OneDriveURL $DstSiteUrl -TenantLocation "Destination"
        }
    }
}

Start-ShareGateOneDriveMigration