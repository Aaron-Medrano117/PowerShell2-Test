function Match-AllMailUsers {
    param (
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$true)] [array] $ImportCSV,
        [Parameter(Mandatory=$false)] [string] $NewDomain
    )
    $ImportedUsers = Import-Csv $ImportCSV
    $AllUsers = @()
    
    #ProgressBar
    $progressref = ($ImportedUsers).count
    $progresscounter = 0

    foreach ($mailbox in $ImportedUsers)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($mailbox.DisplayName)"
        
        Write-Host "Checking for $($mailbox.displayName) in Tenant ..." -fore Cyan -NoNewline
        $newAddressSplit = $mailbox.PrimarySmtpAddress -split "@"
        $newMailboxAddress = $newAddressSplit[0] + "@" + $NewDomain
        if ($mailboxcheck = Get-Mailbox $mailbox.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced)
        {
            Write-Host "found mailbox  " -ForegroundColor Green -nonewline
        }
        elseif ($mailboxcheck = Get-Mailbox $newMailboxAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database)
        {
            Write-Host "found mailbox**  " -ForegroundColor Yellow -nonewline
        }
        elseif ($mailboxcheck = Get-Mailbox $mailbox.displayName -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database)
        {
           Write-Host "found mailbox*  " -ForegroundColor Yellow -nonewline
        }
        elseif ($recipientcheck = Get-Recipient $mailbox.PrimarySmtpAddress -ea silentlycontinue)
        {
            $mailboxcheck = Get-Mailbox $recipientcheck.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database 
            Write-Host "found recipient  " -ForegroundColor Yellow -nonewline
        }
        else
        {
            Write-Host "not found" -ForegroundColor red -NoNewline
            $msoluserscheck = @()
            $MBXStats = @()
        }
        if ($mailboxcheck)
        {
            $msoluserscheck = get-msoluser -UserPrincipalName $mailboxcheck.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation
            $MBXStats = Get-MailboxStatistics $mailboxcheck.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
            $mailbox | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $True
            $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluserscheck.UserPrincipalName
            $mailbox | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluserscheck.IsLicensed
            $mailbox | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluserscheck.Licenses.AccountSkuID -join ",")
            $mailbox | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $mailboxcheck.IsDirSynced
            $mailbox | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $msoluserscheck.PreferredDataLocation
            $mailbox | add-member -type noteproperty -name "Database_Destination" -Value $mailboxcheck.Database
            $mailbox | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $msoluserscheck.BlockCredential
            $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $mailboxcheck.PrimarySmtpAddress
            $mailbox | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipientcheck.RecipientTypeDetails   
            $mailbox | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize
            $mailbox | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount

            #get OneDrive Site details
            $SPOSite = $null
            $EmailAddressUpdate1 = $msoluserscheck.UserPrincipalName.Replace("@","_")
            $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
            $ODSite = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2

            try 
            {
                $SPOSITE = Get-SPOSITE -IncludePersonalSite $true -filter "url -like $ODSite" -ErrorAction SilentlyContinue
                Write-Host "Gathering OneDrive Details ..." -ForegroundColor Cyan -NoNewline
                
                $mailbox | Add-Member -type NoteProperty -Name "OneDriveURL_Destination" -Value $SPOSITE.url
                $mailbox | Add-Member -type NoteProperty -Name "Owner_Destination" -Value $SPOSITE.Owner
                $mailbox | Add-Member -type NoteProperty -Name "StorageUsageCurrent_Destination" -Value $SPOSITE.StorageUsageCurrent
                $mailbox | Add-Member -type NoteProperty -Name "Status_Destination" -Value $SPOSITE.Status
                $mailbox | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability_Destination" -Value $SPOSITE.SiteDefinedSharingCapability
                $mailbox | Add-Member -type NoteProperty -Name "LimitedAccessFileType_Destination" -Value $FDUser.LimitedAccessFileType
                
                Write-Host "done" -ForegroundColor Green
            }
            catch 
            {
                Write-Host "OneDrive Not Enabled for User" -ForegroundColor Yellow
            }
        }
        else 
        {
            $mailbox | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False
            $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "IsLicensed_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "Licenses_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "Database_Destination" -Value $mailboxcheck.Database
            $mailbox | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null  
            $mailbox | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $null
            $mailbox | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "OneDriveURL_Destination"  -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "Owner_Destination"  -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "StorageUsageCurrent_Destination" -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "Status_Destination" -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability_Destination" -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "LimitedAccessFileType_Destination" -Value $null
        }
        Write-host " .. done" -foregroundcolor green
        $AllUsers += $mailbox
    }
    $allUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath
}

Match-AllMailUsers -ImportCSV "C:\Users\fred5646\Rackspace Inc\MPS-TS-Dermavant - General\BOX_user_details.csv" -NewDomain dermavant.com -OutputCSVFilePath "C:\Users\fred5646\Rackspace Inc\MPS-TS-Dermavant - General\MatchedUsers_Dermavant.csv"


#Box Migration
<#
Requirements:
SharePoint Administrator or Global Admin
Administrative Account for Box
Authorize ShareGate in Box - https://support-desktop.sharegate.com/hc/en-us/articles/115000335474
OneDrives Are Provisioned
Assign yourself as Site Collection Admin on OneDrives
Destination Folder must already exist
Uses Explorer for Browser authentications. Require Cookies and Java Enabled

HelpFul Articles:
https://support-desktop.sharegate.com/hc/en-us/articles/115000321633-Walkthrough-Import-from-Box-com-to-OneDrive-for-Business-with-PowerShell
https://support-desktop.sharegate.com/hc/en-us/articles/115000322693-Box-Box-com-Migration-Tips
https://support-desktop.sharegate.com/hc/en-us/articles/360000726183-Impersonate-Owner-for-Box-com-Migration
https://support-desktop.sharegate.com/hc/en-us/articles/115000343993-Import-from-Box-com
https://support-desktop.sharegate.com/hc/en-us/articles/360038509671
https://support-desktop.sharegate.com/hc/en-us/articles/115008000008-This-document-is-currently-locked-by-another-inactive-user-and-cannot-be-retrieved
https://support-desktop.sharegate.com/hc/en-us/articles/115000598707-New-Property-Template


#>


## Import Sharegate Module
Import-Module Sharegate
#Set up Admin Access to Box
$box = Connect-Box -Email "aaron.medrano@rackspace.com" -Admin
#Grab All Box Users
$users = Get-BoxUsers -Box $box
#Connect to SharePoint Online
$credentials = Get-Credential
$tenant = Connect-Site -Url https://dermavant-admin.sharepoint.com/ -Credential $credentials

$myusername = "myusername"
$mypassword = ConvertTo-SecureString "mypassword" -AsPlainText -Force


#$tenant = Connect-Site -Url https://mytenant-admin.sharepoint.com -Credentials $credentials
$tenant = Connect-Site -Url https://dermavant-admin.sharepoint.com/ -Username $myusername -Password $mypassword

$tenant = Connect-Site -Url https://dermavant-admin.sharepoint.com/ -Credential $credentials

#FolderName to Move Box Documents to
$NewOneDriveFolderName = "From-Box"
foreach ($user in $DermavantBoxAccounts) {
    #Clear Variables
    Clear-Variable dstSite
    Clear-Variable dstList
    Clear-Variable dstSiteUrl

    If ($user.Status -eq 'active') {        
        If ($user.BoxAccountType -eq "Shared")
        {
            #connect to OneDrive Url
            $dstSite = Connect-Site -Url $user.SharePoint_URL -credential $credentials

            #Get Document Library List
            $dstList = Get-List -Site $dstSite -name $user.DocumentName
            $SharedDocumentSourcePath = $user.PrimarySmtpAddress

            if ($user.DocumentFolderName)
            {
                $DestinationFolder = $user.DocumentFolderName
                #Create Destination Folder - only for the shared documents
                Import-BoxDocument -Box $box -UserEmail $user.Owner -SourceFilePath $SharedDocumentSourcePath -DestinationList $dstList -DestinationFolder $DestinationFolder -whatif
            }
            else
            {
                #import ALL Documents from user to Destination List
                Import-BoxDocument -Box $box -UserEmail $user.Owner -SourceFilePath $SharedDocumentSourcePath -DestinationList $dstList  -whatif
            }
        }
        elseif ($user.BoxAccountType -eq "Personal"})
        {
            #Grab OneDrive URL for User
            $dstSiteUrl = Get-OneDriveUrl -Tenant $tenant -Email $user.PrimarySmtpAddress

            #connect to OneDrive Url
            $dstSite = Connect-Site -Url $dstSiteUrl -credential $credentials

            #Get Document Library List
            $dstList = Get-List -Site $dstSite -name "Documents"

            #import ALL Documents from user to Destination List
            Import-BoxDocument -Box $box -DestinationList $dstList -UserEmail $user.PrimarySmtpAddress -DestinationFolder $NewOneDriveFolderName -whatif
        }
        #remove site admin after importing
        #Remove-SiteCollectionAdministrator -Site $dstSite
    }
}

#perform an incremental sync
$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
Import-BoxDocument -Box $box -DestinationList $dstList -UserEmail $user.Email -DestinationFolder $NewOneDriveFolderName -CopySettings $copysettings -whatif

#Specify source and destination folders to import
Import-BoxDocument -Box $box -SourceFolder "MyFolder/DocumentsToImport" -UserEmail $user.Email -DestinationList $dstList -DestinationFolder "MyFolder/MySubFolder"

# Test moving Non-Clinical documents
$Chad_User = $users | ?{$_.name -like "Chad*"}
$dstSiteUrl_Non_Clinical = Connect-Site -Url "https://dermavant.sharepoint.com/sites/Non-Clinical/" -Browser
$dstList2 = Get-List -Site $dstSiteUrl_Non_Clinical -Name "NC Documents"
Import-BoxDocument -Box $box -UserEmail $Chad_User.email -SourceFolder "Department Workspaces/Non-Clinical" -DestinationList $dstList2 -WhatIf

<# Migrate using Property Template
 https://support-desktop.sharegate.com/hc/en-us/articles/115000598707-New-Property-Template
Example
Create Property Template to Copy Content
$propertyTemplate = New-PropertyTemplate -AuthorsAndTimestamps -VersionHistory -Permissions -WebParts -NoLinkCorrection -VersionLimit 5 -From "2012-01-01" -To "2012-12-31"
$srcSite = Connect-Site -Url "http://myfarm1/sites/mysourcesite"
$srcList = Get-List -Site $srcSite -Name "mysrclist"
$dstSite = Connect-Site -Url "http://myfarm1/sites/mydestinationsite"
$dstList = Get-List -Site $dstSite -Name "mydstlist"
Copy-Content -SourceList $srcList -DestinationList $dstList -Template $propertyTemplate

Note: When you define a Property template in your script, you have to list everything you need to preserve. In the example above, you are keeping authors and timestamps, version history, permissions, and web parts.
#>
$host.ui.RawUI.WindowTitle = “Brandon Hokenstad”
$user = "chad.berryhill@dermavant.com"
$SharedDocumentSourcePath = "My Documents - All Users/My Documents - Brandon Hokenstad"
#$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
$tenant = Connect-Site -Url https://dermavant-admin.sharepoint.com/ -Credential $credentials
$dstSite = Connect-Site -Url https://dermavant.sharepoint.com/sites/MedicalAffairs -Credential $credentials
$dstList = Get-List -Site $dstSite -name Documents

Import-BoxDocument -Box $box -UserEmail $user -SourceFilePath $SharedDocumentSourcePath -DestinationList $dstList -CopySettings $copysettings
Import-BoxDocument -Box $box -UserEmail $user -DestinationList $dstList -CopySettings $copysettings


## Migrate batches

$NewOneDriveFolderName = "From-Box"
foreach ($user in $DermavantBoxAccounts[20..39]) {
    #Clear Variables
    Clear-Variable dstSite
    Clear-Variable dstList
    Clear-Variable dstSiteUrl

    #Grab OneDrive URL for User
    $dstSiteUrl = Get-OneDriveUrl -Tenant $tenant -Email $user.PrimarySmtpAddress

    #connect to OneDrive Url
    $dstSite = Connect-Site -Url $dstSiteUrl -credential $credentials

    #Get Document Library List
    $dstList = Get-List -Site $dstSite -name "Documents"

    #import ALL Documents from user to Destination List
    Write-Host "Creating BOX Migration for $($user.DisplayName) .. " -foregroundcolor cyan -nonewline
    Import-BoxDocument -Box $box -DestinationList $dstList -UserEmail $user.PrimarySmtpAddress -DestinationFolder $NewOneDriveFolderName
    Write-Host "Done" -foregroundcolor green

#remove site admin after importing
#Remove-SiteCollectionAdministrator -Site $dstSite
}


$host.ui.RawUI.WindowTitle = "BOX patricia.barth"
$user = "patricia.barth@dermavant.com"
## Import Sharegate Module
Import-Module Sharegate
#Set up Admin Access to Box
$box = Connect-Box -Email "aaron.medrano@rackspace.com" -Admin
#Connect to SharePoint Online
$credentials = Get-Credential
$tenant = Connect-Site -Url https://dermavant-admin.sharepoint.com/ -Credential $credentials
$NewOneDriveFolderName = "From-Box"
#Grab OneDrive URL for User
$dstSiteUrl = Get-OneDriveUrl -Tenant $tenant -Email $user
#connect to OneDrive Url
$dstSite = Connect-Site -Url $dstSiteUrl -credential $credentials
#Get Document Library List
$dstList = Get-List -Site $dstSite -name "Documents"

Import-BoxDocument -Box $box -DestinationList $dstList -UserEmail $user -DestinationFolder $NewOneDriveFolderName -InsaneMode


###

$host.ui.RawUI.WindowTitle = “Chad Berryhill”
$user = "chad.berryhill@dermavant.com"
$oneDriveUser = "chad.berryhill@dermavant.com"
$SharedDocumentSourcePath = "My Documents - Chad Berryhill"
$tenant = Connect-Site -Url https://dermavant-admin.sharepoint.com/ -Credential $credentials
$NewOneDriveFolderName = "From-Box"
#Grab OneDrive URL for User
$dstSiteUrl = Get-OneDriveUrl -Tenant $tenant -Email $oneDriveUser
#connect to OneDrive Url
$dstSite = Connect-Site -Url $dstSiteUrl -credential $credentials
#Get Document Library List
$dstList = Get-List -Site $dstSite -name "Documents"

Import-BoxDocument -Box $box -DestinationList $dstList -UserEmail $user -DestinationFolder $NewOneDriveFolderName -InsaneMode
