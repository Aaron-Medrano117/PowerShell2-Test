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

.PARAMETER SourceEmailAddress
Specifies the source email address or UPN.

.PARAMETER SrcSiteURL
Specifies the source site URL.

.PARAMETER DestinationEmailAddress
Specifies the destination email address or UPN.

.PARAMETER DstSiteUrl
Specifies the destination site URL.

.EXAMPLE
.\OneDriveMigration.ps1 -RequestOneDrive -DestinationAdminUrl "https://destination-admin-url" -DestinationEmailAddress "user@destination.com"

This example will request OneDrive site creation for the specified user in the destination tenant.

.EXAMPLE
.\OneDriveMigration.ps1 -Migrate -SourceAdminUrl "https://source-admin-url" -DestinationAdminUrl "https://destination-admin-url" -SourceEmailAddress "user@source.com" -DestinationEmailAddress "user@destination.com"

This example will perform an incremental migration of OneDrive content from the source site to the destination site using the provided email addresses and URLs.


.EXAMPLE
.\OneDriveMigration.ps1 -Migrate -Test -SourceAdminUrl "https://source-admin-url" -DestinationAdminUrl "https://destination-admin-url" -SourceEmailAddress "user@source.com" -DestinationEmailAddress "user@destination.com"

This example will perform a test incremental migration of OneDrive content from the source site to the destination site using the provided email addresseses.


.EXAMPLE
.\OneDriveMigration.ps1 -Migrate -Test -SourceAdminUrl "https://source-admin-url" -DestinationAdminUrl "https://destination-admin-url" -SrcSiteURL "https://source-OneDrive-url" -DstSiteUrl "https://destination-OneDrive-url"

This example will perform a test incremental migration of OneDrive content from the source site to the destination site using the provided URLs.


.EXAMPLE
.\OneDriveMigration.ps1 -AddSecondaryAdmin -DestinationAdminUrl "https://destination-admin-url" -DestinationEmailAddress "user@destination.com"

This example will add a secondary admin to the specified user's OneDrive site in the destination tenant.
#>

# Global Settings

# Destination URL for the new tenant
$Global:DestinationAdminUrl = "https://oakview-admin.sharepoint.com/"
# Source URL for the old tenant
$Global:SourceAdminUrl = "https://spectraxp-admin.sharepoint.com/"

# Connect to source and destination tenants using ShareGate PowerShell module
$Global:SourceTenant = Connect-Site -Url $SourceAdminUrl -Browser
$Global:DestinationTenant = Connect-Site -Url $DestinationAdminUrl -Browser

# Prompt user for source and destination credentials
# Consider using a more secure method for storing and retrieving credentials
$global:SourceCredentials = Get-Credential
$global:DestinationCredentials = Get-Credential


function Start-SingleOneDriveMigrationShareGate {
    [CmdletBinding(DefaultParameterSetName='PrestageEmail')]
    Param(
        [Parameter(ParameterSetName='Prestage', Mandatory=$false, HelpMessage="RequestOneDrive?")]
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
        [string] $SourceEmailAddress,
    
        [Parameter(ParameterSetName='Migrate', Mandatory=$false, HelpMessage="What is the source site URL?")]
        [string] $SrcSiteUrl,
    
        [Parameter(Mandatory=$false, HelpMessage="What is the destination email address or UPN?")]
        [string] $DestinationEmailAddress,
    
        [Parameter(Mandatory=$false, HelpMessage="What is the destination site URL?")]
        [string] $DstSiteUrl
    )

    #Set Up Module, Variables, Credentials, and Connect to SharePoint Sites
    Import-Module Sharegate
    $AllOneDriveErrors = @()

    # Migration Copy Settings
    $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate

    if ($migrate) {
        # Connect to OneDrive Sites in Source
        if ($Global:SourceTenant -ne $null) {
            $SourceTenant = $Global:SourceTenant
        } else {
            $SourceTenant = Connect-Site -Url $SourceAdminUrl -Browser
        }
        Write-Host "Connected to Source Tenant $($SourceTenant.Site.toString()).." -nonewline -BackgroundColor DarkGreen

        try {   
            if ($srcSiteUrl) {}
            else {
                Write-Host "Checking for Source OneDrive Url" -nonewline
                Write-Host " $($SourceEmailAddress)" -nonewline -ForegroundColor Cyan
                Write-Host "  with ShareGate... " -nonewline
                # Get the Source OneDrive URL with ShareGate
                $srcSiteUrl = Get-OneDriveUrl -Tenant $global:sourceTenant -Email $SourceEmailAddress.tostring() -ErrorAction stop
                $srcSiteUrl = $srcSiteUrl.TrimEnd('/')
            }
        } catch {
            try {
                Write-Host "Checking for Source OneDrive Url with SharePoint... " -nonewline
                # Get the Source OneDrive URL with SharePoint
                Connect-SPOService -Url $SourceAdminUrl -Credential $SourceCredentials
                $OneDriveSrcUrlCheck = (Get-SPOSite -Filter "Owner -eq '$SourceEmailAddress' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ErrorAction Stop).Url
            } catch {
                Write-Host "Unable to find Source OneDrive Site for $SourceEmailAddress" -ForegroundColor Red
                $CurrentError = New-Object PSObject
                $CurrentError | Add-Member -Type NoteProperty -Name "Commandlet" -Value $_.CategoryInfo.Activity
                $CurrentError | Add-Member -Type NoteProperty -Name "FailureActivity" -Value "UnableToFindSourceOneDrive" -Force
                $CurrentError | Add-Member -Type NoteProperty -Name "Tenant" -Value $SourceTenant.Site -Force
                $CurrentError | Add-Member -Type NoteProperty -Name "User" -Value $SourceEmailAddress
                $CurrentError | Add-Member -Type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllOneDriveErrors += $CurrentError
            }
            return $AllOneDriveErrors
        }

        # Connect to OneDrive Sites in Source
        if ($SrcSiteUrl) {
        } elseif ($OneDriveSrcUrlCheck) {
            $SrcSiteUrl = $OneDriveSrcUrlCheck
            Write-host "Found $($SrcSiteUrl) " -nonewline -ForegroundColor Green
        }
        #Pre-Check Site Admin
        Write-Host "Attemtpting Connection to Source OneDrive... " -ForegroundColor Yellow -NoNewline
        try {
            $SrcSite = Connect-Site -Url $SrcSiteUrl -UseCredentialsFrom $global:SourceTenant -ErrorAction Stop
            
        }
        catch {
            try {
                $AdminCheck = Get-SPOUser -Site $SrcSiteUrl.ToString() -ErrorAction Stop
            }
            catch {
                #If failed, switch to Destination Tenant
                Write-Host ".. switching to Source Tenant... " -foregroundcolor Yellow -NoNewline
                Connect-SPOService $SourceAdminUrl -ModernAuth
                $rootSiteURL = Get-SPOSite -limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
                $rootURL = $rootSiteURL.url -replace '/sites.*', ''
                Write-Host "Connected to: $($rootURL)" -foregroundcolor Green -NoNewline

                #add Destination Tenant Admin as Site Admin
                Set-SPOUser -Site $SrcSiteUrl.ToString() -LoginName $global:SourceCredentials.username.ToString() -IsSiteCollectionAdmin $true -ErrorAction Stop
                Write-Host "$($global:SourceCredentials.username.ToString()) Added as Site Admin." -ForegroundColor Cyan

                # Wait for the SPOUser to be retrieved before proceeding
                while (!(Get-SPOUser -Site $SrcSiteUrl.ToString() -ErrorAction SilentlyContinue)) {
                    Write-Host " ." -NoNewline -ForegroundColor Yellow
                    Start-Sleep -Seconds 3
                }
                continue
            }
        }
        try {
            # Attempt connect to Source OneDrive
            $SrcSite = Connect-Site -Url $SrcSiteUrl -UseCredentialsFrom $global:SourceTenant -ErrorAction Stop

            # Get Source OneDrive Documents
            $SrcList = Get-List -Site $SrcSite -Name "Documents" -ErrorAction Stop
            
        }
        catch {
            return $error[0]
        }

        ##Check if Destination OneDrive is Enabled
        if ($srcList) {
            #Connected To Source OneDrive
            Write-Host "Connected" -ForegroundColor Green
            try {
                # Connect to OneDrive Sites in Destination
                if ($Global:DestinationTenant -ne $null) {
                    $DestinationTenant = $Global:DestinationTenant
                } else {
                    $DestinationTenant = Connect-Site -Url $DestinationAdminUrl -Browser
                }                
                Write-Host "Connected to Destination Tenant $($DestinationTenant.Site.toString())... " -NoNewline -BackgroundColor DarkGreen
                
                #Gather OneDrive URL
                if ($DstSiteUrl) {
                    Write-Host "Using $($DstSiteUrl) " -nonewline
                }
                else {
                    Write-Host "Checking for Destination OneDrive Url" -nonewline
                    Write-Host " $($DestinationEmailAddress) " -nonewline -ForegroundColor Cyan
                    Write-Host " with ShareGate... " -nonewline
                    # Checking for Destination OneDrive Url with ShareGate
                    if ($DstSiteUrl = Get-OneDriveUrl -Tenant $DestinationTenant -Email $DestinationEmailAddress.tostring() -ErrorAction SilentlyContinue) {
                        $DstSiteUrl = $DstSiteUrl.TrimEnd('/')
                        Write-host "Found $($DstSiteUrl)... " -nonewline -ForegroundColor Green
                    } else {
                        Write-Host "Checking Site Url for $($DestinationEmailAddress) .. " -ForegroundColor Cyan -NoNewline
                        Connect-SPOService -Url $DestinationAdminUrl -Credential $DestinationCredentials
                    
                        if ($OneDriveDstUrlCheck = (Get-SPOSite -Filter "Owner -eq '$DestinationEmailAddress' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ErrorAction Stop).URL) {
                            Write-Host "Site exists for $($OneDriveDstUrlCheck)" -ForegroundColor Magenta
                        } else {
                            Write-Host "No OneDrive Site Provisioned for $($DestinationEmailAddress) .." -ForegroundColor Red -NoNewline
                            Request-SPOPersonalSite -UserEmails $DestinationEmailAddress -ErrorAction Stop
                            Write-Host "OneDrive Site Requested for $($DestinationEmailAddress)" -ForegroundColor Green
                            return
                        }
                    }
                }
            } Catch {
                #If Destination Site does not exist
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "UnableToFindDestinationOneDrive" -Force
                $currenterror | Add-Member -type NoteProperty -Name "Tenant" -Value $sourceTenant.Site -Force
                $currenterror | Add-Member -type NoteProperty -Name "User" -Value $DestinationEmailAddress
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllOneDriveErrors += $currenterror
                return $currenterror
            }

            # Connect to OneDrive Sites in Destination
            if ($DstSiteUrl) {
            } elseif ($OneDriveDstUrlCheck) {
                $DstSiteUrl = $OneDriveDstUrlCheck
            }
            #Pre-Check Site Admin
            Write-Host "Attemtpting Connection to Destination OneDrive... " -ForegroundColor Yellow -NoNewline
            try {
                $AdminCheck = Get-SPOUser -Site $DstSiteUrl.ToString() -ErrorAction Stop
            }
            catch {
                #If failed, switch to Destination Tenant
                Write-Host ".. switching to Destination Tenant... " -foregroundcolor Yellow -NoNewline
                Connect-SPOService $DestinationAdminUrl -Credential $global:DestinationCredentials
                $rootSiteURL = Get-SPOSite -limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
                $rootURL = $rootSiteURL.url -replace '/sites.*', ''
                Write-Host "Connected to: $($rootURL)" -foregroundcolor Green -NoNewline

                #add Destination Tenant Admin as Site Admin
                Set-SPOUser -Site $DstSiteUrl.ToString() -LoginName $global:DestinationCredentials.username.ToString() -IsSiteCollectionAdmin $true -ErrorAction Stop
                Write-Host "$($global:DestinationCredentials.username.ToString()) Added as Site Admin." -ForegroundColor Cyan

                # Wait for the SPOUser to be retrieved before proceeding
                while (!(Get-SPOUser -Site $DstSiteUrl.ToString() -ErrorAction SilentlyContinue)) {
                    Write-Host " ." -NoNewline -ForegroundColor Yellow
                    Start-Sleep -Seconds 3
                }
                continue
            }

            # Get OneDrive Documents
            #Write-Host "Attemtpting Connection to Documents... " -nonewline -ForegroundColor Yellow
            try {
                # Get OneDrive Documents
                $DstSite = Connect-Site -Url $DstSiteUrl -UseCredentialsFrom $DestinationTenant -ErrorAction Stop              
                $DstList = Get-List -Site $DstSite -Name "Documents" -ErrorAction Stop
                Write-Host "Connected" -ForegroundColor Green
                
            } catch {
                return $error[0]
            }
            # Create a new object to store the results
            $OneDriveResults = New-Object PSObject
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceAdminTenantURL" -Value $sourceTenant.Address
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationAdminTenantURL" -Value $destinationTenant.Address
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceName" -Value $srcSite.Title
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceSite" -Value $srcSite.Address
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationName" -Value $dstSite.Title
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationURL" -Value $dstSite.Address

            # Copy OneDrive Files from Source to Destination
            if ($DstList) {
                if ($Test) {
                    $TaskName = "Test OneDrive Migration for $($SrcSite.Title) to $($DstSite.Title)"
    
                    # Progress Bar Current 2
                    Write-Progress -Id 2 -Activity "$($TaskName)"                
                    # Test Move with Incremental using Insane Mode
                    $Result = Copy-Content -SourceList $SrcList -DestinationList $DstList -InsaneMode -CopySettings $CopySettings -TaskName $TaskName -WarningAction SilentlyContinue -WhatIf
                    $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value "Test"
                } else {
                    $TaskName = "Incremental OneDrive Migration for $($SrcSite.Title) to $($DstSite.Title)"
                    # Progress Bar Current 2
                    Write-Progress -Id 2 -Activity "$($TaskName)"
                    
                    $Result = Copy-Content -SourceList $SrcList -DestinationList $DstList -InsaneMode -CopySettings $CopySettings -TaskName $TaskName -WarningAction SilentlyContinue
                    $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value "Incremental"
                }
            }
            # Add Job Result of Job 
            If ($Result) {
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
    }
    # Run Prestage OneDrive
    if ($RequestOneDrive) {
        # Connect to Destination SharePoint Site
        try {
            if ($DstSiteUrl = Get-OneDriveUrl -Tenant $DestinationTenant -Email $DestinationEmailAddress -ErrorAction Stop) {
                Write-Host "OneDrive $($DstSiteUrl) already Exists for $($DestinationEmailAddress)" -ForegroundColor Cyan
                $AlreadyExists += $DstSiteUrl
            } else {
                Request-SPOPersonalSite -UserEmails $DestinationEmailAddress -ErrorAction Stop
                Write-Host "OneDrive Site Requested for $($DestinationEmailAddress)" -ForegroundColor Green
                $RequestedSite += $User
            }
        } catch {
            # If Site does not exist
            if ($OneDriveDstUrlCheck = Get-SPOSite -Filter "Owner -eq '$DestinationEmailAddress' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ErrorAction SilentlyContinue) {
                $AlreadyExists += $OneDriveDstUrlCheck
            } else {
                Request-SPOPersonalSite -UserEmails $DestinationEmailAddress -ErrorAction Stop
                Write-Host "OneDrive Site Requested for $($DestinationEmailAddress)" -ForegroundColor Green
                $RequestedSite += $User
            }
        }
    }
    #Check if OneDrive Admin Added
    if ($AddSecondaryAdmin) {
        try {
            $DstSiteUrl = Get-OneDriveUrl -Tenant $DestinationTenant -Email $DestinationEmailAddress -ErrorAction Stop
            

            if ($SPOPermUsers = (Get-SPOUser -Site $DstSiteUrl.ToString() -ErrorAction Stop).LoginName) {
                Write-Host "Already Site Admin for $($DstSiteUrl.ToString())" -ForegroundColor Yellow
                $AlreadySiteAdmin += $User
            } else {
                try {
                    $AdminRequest = Set-SPOUser -Site $DstSiteUrl.ToString() -LoginName $DestinationServiceAccount.ToString() -IsSiteCollectionAdmin $true -ErrorAction Stop
                    $SiteAdminAdded += $User
                    Write-Host "Site Admin added for $($OneDriveDstUrlCheck.Url)" -ForegroundColor Green
                } catch {
                    Write-Host "Unable to Add Admin for $($DestinationEmailAddress)" -ForegroundColor Red
                    $FailedToAddAdminToOneDrive += $User
                }
            }
        } catch {
            try {
                Write-Host "Checking Site Url for $($DestinationEmailAddress) .. " -ForegroundColor Cyan -NoNewline
                $OneDriveDstUrlCheck = Get-SPOSite -Filter "Owner -eq '$DestinationEmailAddress' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ErrorAction Stop
                $AdminRequest = Set-SPOUser -Site $OneDriveDstUrlCheck.URL -LoginName $DestinationServiceAccount.ToString() -IsSiteCollectionAdmin $true -ErrorAction Stop
                Write-Host "Site Admin added for $($OneDriveDstUrlCheck.Url)" -ForegroundColor Magenta
            } catch {
                $NoOneDriveProvisioned += $DestinationEmailAddress
                Write-Host "No OneDrive Site Provisioned for $($DestinationEmailAddress)" -ForegroundColor Red
            }
        }
    }
}