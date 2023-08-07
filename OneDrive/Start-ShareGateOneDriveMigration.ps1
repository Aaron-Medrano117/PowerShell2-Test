<#
.SYNOPSIS
A script for managing OneDrive migrations and pre-stage operations using Sharegate PowerShell module.

.DESCRIPTION
This script is designed to assist with OneDrive migrations and pre-stage operations using the Sharegate PowerShell module. It provides several options, including migrating data, running incremental syncs, pre-staging OneDrive, and adding secondary admins to OneDrive sites. The script supports using email addresses or URLs for the source and destination sites.

.PARAMETER Operation
Specifies which action to perform
Accepts: RequestOneDrive, AddSecondaryAdmin, Migrate, Test. 
Accepts multiple values
Use RequestOneDrive to request OneDrive site creation for a destination user
Use AddSecondaryAdmin to grant the migration account admin access to the destination OneDrive site
Use Migrate to perform an data move migration
Use Test to perform a test migration.

.PARAMETER CopyOperation
Specifies the type of migration to perform.
Accepts Incremental, Overwrite, Skip
Default is Incremental
Incremental sync will copy new and updated files from the source to the destination
Overwrite will copy all files from the source to the destination, overwriting any existing files
Skip will skip the migration of files from the source to the destination


.PARAMETER SourceAdminUrl
Specifies the Source Admin Site URL.

.PARAMETER DestinationAdminUrl
Specifies the Destination Admin Site URL.

.PARAMETER SourceAccountInfo
Specifies the source email address or UPN.

.PARAMETER DestinationAccountInfo
Specifies the destination email address or UPN.


.EXAMPLE
Start-ShareGateOneDriveMigration -Operation RequestOneDrive -DestinationAccountInfo "user@destination.com"

This example will request OneDrive site creation for the specified user in the destination tenant.

.EXAMPLE
Start-ShareGateOneDriveMigration -Operation Migrate -CopyOperation Incremental -SourceAccountInfo "user@source.com" -DestinationAccountInfo "user@destination.com"

This example will perform an incremental migration of OneDrive content from the source site to the destination site using the provided email addresses or UPNs.


.EXAMPLE
Start-ShareGateOneDriveMigration -Test -Operation Migrate -CopyOperation Incremental -SourceAccountInfo "user@source.com" -DestinationAccountInfo "user@destination.com"

This example will perform a TEST incremental migration of OneDrive content from the source site to the destination site using the provided email addresseses.


.EXAMPLE
Start-ShareGateOneDriveMigration -Operation Migrate -CopyOperation Incremental -SourceAccountInfo "https://source-OneDrive-url" -DestinationAccountInfo "https://destination-OneDrive-url"

This example will perform a incremental migration of OneDrive content from the source site to the destination site using the provided URLs.


.EXAMPLE
Start-ShareGateOneDriveMigration -Operation AddSecondaryAdmin -DestinationAccountInfo "user@destination.com"

This example will add a secondary admin to the specified user's OneDrive site in the destination tenant.

.EXAMPLE
Start-ShareGateOneDriveMigration -Operation Migrate -CopyOperation Overwrite -SourceAccountInfo "https://source-OneDrive-url" -DestinationAccountInfo "https://destination-OneDrive-url"

This example will perform an overwrite migration of OneDrive content from the source site to the destination site using the provided URLs. 


#>


# Global Settings
#Set Up Module, Variables, Credentials, and Connect to SharePoint Sites
Import-Module Sharegate

# Verify if ShareGate Source and Destination Tenant have been connected
if ($null -eq $Global:SourceTenant) {
    Write-Host "Enter Source Credentials" -ForegroundColor Yellow
    $global:SourceCredentials = Get-Credential
    $Global:SourceAdminUrl = Read-Host "What is the Source Admin Site URL?"
    $Global:SourceTenant = Connect-Site -Url $Global:SourceAdminUrl -Browser
    
}
else {
    $Global:SourceAdminUrl = $Global:SourceTenant.Site
}
if ($null -eq $Global:DestinationTenant) {
    Write-Host "Enter Destination Credentials" -ForegroundColor Yellow
    $global:DestinationCredentials = Get-Credential
    $Global:DestinationAdminUrl = Read-Host "What is the Destination Admin Site URL?"
    $Global:DestinationTenant = Connect-Site -Url $Global:DestinationAdminUrl -Browser
}
else {
    $Global:DestinationAdminUrl = $Global:DestinationTenant.Site
}


function Start-ShareGateOneDriveMigration {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('RequestOneDrive', 'AddSecondaryAdmin', 'Migrate', 'Test')]
        [string[]]$Operation,

        [Parameter(Mandatory=$false, HelpMessage="What is the source email address, UPN, or OneDriveURL?")]
        [string]$SourceAccountInfo,

        [Parameter(Mandatory=$true, HelpMessage="What is the destination email address, UPN, or OneDriveURL?")]
        [string]$DestinationAccountInfo,

        [Parameter(Mandatory = $false, HelpMessage="What is the Copy Operation?")]
        [ValidateSet('Incremental', 'Overwrite', 'Skip')]
        [string]$CopyOperation,

        [Parameter(Mandatory = $false, HelpMessage="What is the Destination Folder Name to migrate data to?")]
        [ValidateSet('Custom', 'Migrate-OneDrive')]
        [string]$DestinationFolder

    )
    #Global Variables
    $AllOneDriveErrors = @{}
    # Verify if ShareGate Source and Destination Tenant have been connected
    if ($null -eq $Global:SourceTenant) {
        $Global:SourceAdminUrl = Read-Host "What is the Source Admin Site URL?"
        $Global:SourceTenant = Connect-Site -Url $Global:SourceAdminUrl -Browser
    }
    else {
        $Global:SourceAdminUrl = $Global:SourceTenant.Site
    }
    if ($null -eq $Global:DestinationTenant) {
        $Global:DestinationAdminUrl = Read-Host "What is the Destination Admin Site URL?"
        $Global:DestinationTenant = Connect-Site -Url $Global:DestinationAdminUrl -Browser
    }
    else {
        $Global:DestinationAdminUrl = $Global:DestinationTenant.Site
    }

    ## Functions
    # Function to get OneDrive URL for a user from the source or destination tenant
    function Get-OneDriveUrlForUser {
        param (
            [Parameter(Mandatory=$true)]
            [string]$UPN,
            [Parameter(Mandatory=$true)]
            [ValidateSet('Source', 'Destination')]
            [string]$TenantLocation,
            [Parameter(Mandatory=$false)]
            [switch]$RequestOneDrive
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
            Write-Error $_.Exception.Message
            $AllOneDriveErrors[$UPN] = [PSCustomObject]@{
                Commandlet = "Get-OneDriveUrlForUser"
                FailureActivity = "UnableToFindOneDrive"
                Tenant = $TenantLocation
                User = $UPN
                Error = ($_.Exception.Message)
            }
            try {
                # Look up OneDrive URL using the sharepoint online method
                Connect-SPOService -Url $TenantAdminUrl.tostring() -Credential $TenantCredentials
                $OneDriveUrlCheck = Get-SPOSite -Filter "Owner -eq '$UPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ErrorAction SilentlyContinue
                $OneDriveUrlCheck = $OneDriveUrlCheck.TrimEnd('/')
                Write-Host "$($TenantLocation) $($UPN): Found - " -ForegroundColor Green -NoNewline
                Write-Host "OneDrive $($OneDriveUrlCheck)" -ForegroundColor Cyan
                return $OneDriveUrlCheck
            } catch {
                Write-Host "$($TenantLocation) OneDrive URL not found - $($UPN)" -ForegroundColor Yellow
                #Request OneDrive
                if ($RequestOneDrive) {
                    Connect-SPOService -Url $TenantAdminUrl.tostring() -Credential $TenantCredentials
                    Request-SPOPersonalSite -UserEmails $DestinationAccountInfo -ErrorAction Stop
                    Write-Host "Destination OneDrive Site Requested for $($DestinationAccountInfo)" -ForegroundColor Green
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
            #$TenantShareGateDetails
            $SiteDetails = Connect-Site -Url $OneDriveURL -UseCredentialsFrom $TenantShareGateDetails -ErrorAction Stop
            # Get Source OneDrive Documents
            $DocumentsLibrary = Get-List -Site $SiteDetails -Name "Documents" -ErrorAction Stop
            #Write-Host "Already Site Admin for $($DocumentsLibrary.Source)" -ForegroundColor Yellow
            Return $DocumentsLibrary
        }
        catch {    
            try {
                $AdminCheck = Get-SPOUser -Site $OneDriveURL.ToString() -ErrorAction Stop
            }
            catch {
                #If failed, switch to TenantLocation Tenant
                Write-Host ".. switching to $($TenantLocation) Tenant... " -foregroundcolor Yellow -NoNewline
                Connect-SPOService $TenantAdminUrl -ModernAuth $true #-Credential $TenantCredentials -ErrorAction Stop
            }

            #add Destination Tenant Admin as Site Admin
            try {
                $result = Set-SPOUser -Site $OneDriveURL -LoginName $global:TenantCredentials.username.ToString() -IsSiteCollectionAdmin $true -ErrorAction Stop
                Write-Host "$($global:TenantCredentials.username.ToString()) Added as Site Admin." -ForegroundColor Green
            }
            catch {
                Write-Host "Unable to Add Admin for $($DestinationAccountInfo)" -ForegroundColor Red
                $FailedToAddAdminToOneDrive += $User
                return $error[0]
            }
            
        }
    }

    # Function to Migrate Data
    function Move-OneDrive {
        param (
            [Parameter(Mandatory=$true)]
            [string] $srcSiteUrl,
            
            [Parameter(Mandatory=$true)]
            [string] $DstSiteUrl,

            [Parameter(Mandatory=$False)]
            [string]$DestinationFolderName,

            [Parameter(Mandatory=$true)]
            [string] $CopyOperation
        )

        # Migration Copy Settings
        switch ($CopyOperation) {
            "Incremental" {
            $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
            }
            "Overwrite" {
                $copysettings = New-CopySettings -OnContentItemExists OverWrite
            }
            "Skip" {
                $copysettings = New-CopySettings -OnContentItemExists Skip
            }
            Default {
                $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
            }
        }

        ### Source - Begin Region ###
        try {
            $SourceDocumentLibrary = Connect-OneDriveForUser -OneDriveURL $srcSiteUrl -TenantLocation "Source" -ErrorAction Stop
            Write-Host "Connected to Source OneDrive .. " -ForegroundColor Green -NoNewline
        }
        catch {
            Write-Host "Unable to Connect to Source OneDrive. Skipping" -ForegroundColor Red
            Write-Host ""
            #Write-Host "Source $($SourceAccountInfo): Not Connected" -ForegroundColor Red
            $AllOneDriveErrors[$SourceAccountInfo] = [PSCustomObject]@{
                Commandlet = $_.CategoryInfo.Activity
                FailureActivity = "UnableToConnectOneDrive"
                Tenant = $TenantShareGateDetails.Site
                User = $SourceAccountInfo
                Error = ($_.Exception)
            }
            return
        }
        ### Source - End Region ###
        
        ### Destination - Start Region ###
        # Connect to OneDrive Sites in Destination
        $DestinationDocumentLibrary = Connect-OneDriveForUser -OneDriveURL $DstSiteUrl -TenantLocation "Destination"
        Write-Host "Connected to Destination OneDrive" -ForegroundColor Green
        ### Destination - End Region ###
            
        # Create a new object to store the results
        $OneDriveResults = New-Object PSObject
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceAdminTenantURL" -Value $Global:SourceTenant.Site
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationAdminTenantURL" -Value $Global:DestinationTenant.Site
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceAccountInfo" -Value $SourceAccountInfo
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceName" -Value $SourceDocumentLibrary.Title
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceSite" -Value $SourceDocumentLibrary.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationAccountInfo" -Value $DestinationAccountInfo
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationName" -Value $DestinationDocumentLibrary.Title
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationURL" -Value $DestinationDocumentLibrary.Address

        # Copy OneDrive Files from Source to Destination
        if ($DestinationDocumentLibrary) {
            switch ($Operationx) {
                "Test" {
                    # Test Copy with Insane Mode
                    $TaskName = "Test OneDrive Migration $($SourceAccountInfo) to $($DestinationAccountInfo)"
                    # Progress Bar Current 2
                    Write-Progress -Id 2 -Activity "$($TaskName)"  
                    if ($DestinationFolderName) {
                        $Result = Copy-Content -SourceList $SourceDocumentLibrary -DestinationList $DestinationDocumentLibrary -DestinationFolder $DestinationFolderName -InsaneMode -CopySettings $CopySettings -TaskName $TaskName -WarningAction SilentlyContinue -WhatIf
                    }
                    else {
                        $Result = Copy-Content -SourceList $SourceDocumentLibrary -DestinationList $DestinationDocumentLibrary -InsaneMode -CopySettings $CopySettings -TaskName $TaskName -WarningAction SilentlyContinue -WhatIf               
                    }
                    $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value "Test-$CopyOperation"
                 }
                Default {
                    # Copy with Insane Mode
                    $TaskName = "$($CopyOperation) OneDrive Migration $($SourceAccountInfo) to $($DestinationAccountInfo)"
                    # Progress Bar Current 2
                    Write-Progress -Id 2 -Activity "$($TaskName)"
                    
                    if ($DestinationFolderName) {
                        $Result = Copy-Content -SourceList $SourceDocumentLibrary -DestinationList $DestinationDocumentLibrary -DestinationFolder $DestinationFolderName -InsaneMode -CopySettings $CopySettings -TaskName $TaskName -WarningAction SilentlyContinue
                    }
                    else {
                        $Result = Copy-Content -SourceList $SourceDocumentLibrary -DestinationList $DestinationDocumentLibrary -InsaneMode -CopySettings $CopySettings -TaskName $TaskName -WarningAction SilentlyContinue
                    }
                    $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value $CopyOperation
                }
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
    
    # Main Function

    #New Heading
    Write-Host "$($DestinationAccountInfo)" -ForegroundColor Black -BackgroundColor Cyan

    #Check Destination OneDrive URL
    if ($DestinationAccountInfo -like "https://*") {
        $DstSiteUrl = $DestinationAccountInfo
        $DestinationAccountInfo = ($DestinationAccountInfo -split '/personal/')[1]
    }
    else {
        $DstSiteUrl = Get-OneDriveUrlForUser -TenantLocation Destination -UPN $DestinationAccountInfo -RequestOneDrive -ErrorAction Stop
    }
    
    switch ($Operation) {
        "RequestOneDrive" {
            #$global:RequestOneDrive = $true # Review if still needed or how to deal with string for Operation above
            #Get-OneDriveUrlForUser -TenantLocation Destination -UPN $DestinationAccountInfo -RequestOneDrive
        }
        "AddSecondaryAdmin" {

            #Check if connected to Correct Tenant
            try {
                Get-OneDriveUrlForUser -TenantLocation Destination -UPN $DestinationAccountInfo -RequestOneDrive

                $AdminCheck = Get-SPOUser -Site $OneDriveURL.ToString() -ErrorAction Stop
            }
            catch {
                #If failed, switch to TenantLocation Tenant
                Write-Host ".. switching to $($TenantLocation) Tenant... " -foregroundcolor Yellow -NoNewline
                Connect-SPOService $TenantAdminUrl -ModernAuth $true #-Credential $TenantCredentials -ErrorAction Stop
            }

            #add Destination Tenant Admin as Site Admin
            try {
                $global:RequestOneDrive = $null
                $result = Set-SPOUser -Site $OneDriveURL -LoginName $global:TenantCredentials.username.ToString() -IsSiteCollectionAdmin $true -ErrorAction Stop
                Write-Host "$($global:TenantCredentials.username.ToString()) Added as Site Admin." -ForegroundColor Green
            }
            catch {
                Write-Host "Unable to Add Admin for $($DestinationAccountInfo)" -ForegroundColor Red
                $FailedToAddAdminToOneDrive += $User
                return $error[0]
            }

            try {
                Connect-OneDriveForUser -OneDriveURL $DstSiteUrl -TenantLocation "Destination" -ErrorAction Stop
                #Write-Host "Migration Account is Already Site Admin for $DstSiteUrl" -ForegroundColor Yellow

            }
            catch {
                Write-Host $_.Exception.Message -ForegroundColor Red
            }
            
        }
        "Migrate" {
            $global:RequestOneDrive = $null
            #Check Source OneDrive URL
            if ($SourceAccountInfo -like "https://*") {
                $srcSiteUrl = $SourceAccountInfo
                $SourceAccountInfo = ($SourceAccountInfo -split '/personal/')[1]
            }
            else {
                $srcSiteUrl = Get-OneDriveUrlForUser -TenantLocation Source -UPN $SourceAccountInfo -ErrorAction Stop
            }

            #Migrate OneDrive Data from Source to Destination
            switch ($DestinationFolder) {
                "Custom" { 
                    $DestinationFolderName = Read-Host "What is the Destination Folder Name?"
                    Move-OneDrive -DstSiteUrl $DstSiteUrl -srcSiteUrl $srcSiteUrl -CopyOperation $CopyOperation -DestinationFolderName $DestinationFolderName
                 }
                "Migrated-OneDrive" {
                    $DestinationFolderName = "Migrated-OneDrive"
                    Move-OneDrive -DstSiteUrl $DstSiteUrl -srcSiteUrl $srcSiteUrl -CopyOperation $CopyOperation -DestinationFolderName $DestinationFolderName
                }
                Default {
                    Move-OneDrive -DstSiteUrl $DstSiteUrl -srcSiteUrl $srcSiteUrl -CopyOperation $CopyOperation
                }
            }
        }
        Default {}
    }
}

Start-ShareGateOneDriveMigration