<#
    Testing code using HashTables for the output. Creates a hashtable for each user and then adds it to the $allMailboxStats hashtable.
    After all users are processed, the $allMailboxStats hashtable is converted to a CSV file.
    This is a test to see if this is faster than using a PSCustomObject.
    #This also uses the functions in the MyModule.psm1 file.

    Uses PSObject to create a custom object array
    Includes SharePoint Details
    Includes Archive Mailbox Details
    Includes SendAs Permissions
    Includes Inactive Mailboxes
    Includes Group Mailboxes
    Includes Shared Mailboxes
    Does not include Discovery Search Mailboxes
    includes a progress bar
    Includes a timer
    Includes a try/catch to handle errors
    includes a do/while loop to handle errors
    Includes a function to handle the progress bar
    Includes a function to check if the user is connected to SharePoint Online
    Includes a function to check if the SharePoint Online module is installed

    Against Spectra Tenant took 8:18:23 to complete (using ExchangeOnlineManagement module V2 commands)
    -- 4143 mailboxes

    version 1.3.1
    Author: Aaron Medrano
    #removed wait time for OneDrive Checks #added Import-RequiredModules function

    version 1.3.2
    Author: Andrew Cronic
    Added Public Folders, Stats, Perms and Group Mailboxes + Group Mailbox Associated SharePoint site data.  Completed in 2:23 against 3492 items.  
    Still some improvements needed and I have some ideas to process each different data set independently to avoid unnecessary error handling / checks that are needed when all object types are processed in the same loop.

    version 1.3.3
    Author: Aaron Medrano
    Removed Test-SPOSite function
    Updated Connect-Office365RequiredServices function
    Removed parameters for Connect-Office365RequiredServices function to keep it generalized
#>



#Progress Helper
function Write-ProgressHelper {

    param (
        [int]$ProgressCounter,
        [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [int]$TotalCount
    )

    $secondsElapsed = (Get-Date) – $global:start
    $progresspercentcomplete = [math]::Round((($progresscounter / $TotalCount)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$TotalCount+"]"

    $progressParameters = @{
        Activity = $Activity
        Status = "$progressStatus $($secondsElapsed.ToString('hh\:mm\:ss'))"
        PercentComplete = $progresspercentcomplete
    }

 
    # if we have an estimate for the time remaining, add it to the Write-Progress parameters
    if ($secondsRemaining) {
        $progressParameters.SecondsRemaining = $secondsRemaining
    }

    if ($ID) {
        $progressParameters.ID = $ID
    }

    if ($CurrentOperation) {
        $progressParameters.CurrentOperation = $CurrentOperation
    }

    # Write the progress bar
    Write-Progress @progressParameters

    # estimate the time remaining
    $global:secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount – $progresscounter)

}

#Verify all required modules are installed
function Connect-Office365RequiredServices {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the UPN of the Account to authenticate')] [string]$userPrincipalName,
        [Parameter(Mandatory=$True,HelpMessage='Provide the SharePoint Online Admin URL. Name is usually formatted "https://<yourtenant>-admin.sharepoint.com"')] [string]$SPOAdminURL
	)
    #Clear Previous Variables
    Clear-Variable -Name exOnlineConnected,EXOModuleFound,MSOnlineModuleFound,MSOnlineConnected,SPOConnected,SPOModuleFound -ErrorAction SilentlyContinue

    # Check if required modules are installed
    ## ExchangeOnlineManagement Connection, module Import and installation Check
    Write-Host "Checking for Existing Connections and Required Modules" -foregroundcolor Cyan
    try {
        $EXOOrgCheck = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "Already Connected to Exchange Online: $($EXOOrgCheck.Name)" -foregroundcolor Green
        $exOnlineConnected = $true
    }
    Catch {
        Write-Host "ExchangeOnlineManagement module..." -NoNewline
        if ((Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue) -ne $null) {
            if (Get-Module -Name "ExchangeOnlineManagement") {
                Write-Host "Already Imported" -ForegroundColor Green
                $EXOModuleFound = $true
            }
            else {
                Write-Host "was not loaded. Importing module..." -NoNewline -ForegroundColor Yellow
                try {
                    Import-Module ExchangeOnlineManagement
                    Write-Host Completed -foregroundcolor Green
                    $EXOModuleFound = $true
                }
                catch {
                    Write-Error "Error importing ExchangeOnlineManagement module: $($_.Exception.Message)"
                    $EXOModuleFound = $False
                }
            } 
        }
        else {
            Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-the-exchange-online-powershell-module"
            $EXOModuleFound = $False
        }
    }
    ## MSOnline Connection, module Import and installation Check
    Try {
        $MSOCompanyCheck = Get-MsolCompanyInformation -ErrorAction Stop
        Write-Host "Already Connected to MSOnline: $($MSOCompanyCheck.DisplayName)" -foregroundcolor Green
        $msOnlineConnected = $true
    }
    Catch {
        Write-Host "MSOnline module..." -NoNewline
        if ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null) {
            if ((Get-Module -Name "MSOnline") -ne $null) {
                Write-Host "Already Imported" -ForegroundColor Green
                $MSOnlineModuleFound = $true
            }
            else {
                Write-Host "was not loaded. Importing module..." -NoNewline -ForegroundColor Yellow
                try {
                    Import-Module MSOnline
                    Write-Host Completed -foregroundcolor Green
                    $MSOnlineModuleFound = $true
                }
                catch {
                    Write-Error "Error importing MSOnline module: $($_.Exception.Message)"
                    $MSOnlineModuleFound = $False
                }
            } 
        }
        else {
            Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator"
            $MSOnlineModuleFound = $false
        }
    }

    ## Microsoft.Online.SharePoint Connection, module Import and installation Check
    Try {
        $rootSiteURL = Get-SPOSite -limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
        $rootURL = $rootSiteURL.url -replace '/sites.*', ''
        Write-Host "Already Connected to SharePoint Online: $($rootURL)" -foregroundcolor Green
        $spoConnected = $True
    }
    Catch {
        Write-Host "Microsoft.Online.SharePoint.PowerShell module.." -NoNewline
        if ((Get-InstalledModule -Name "Microsoft.Online.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -ne $null) {
            if ((Get-Module -Name "Microsoft.Online.SharePoint.PowerShell") -ne $null) {
                Write-Host "Already Imported" -ForegroundColor Green
                $SPOModuleFound = $true
            }
            else {
                Write-Host "was not loaded. Importing module..." -NoNewline -ForegroundColor Yellow
                try {
                    Import-Module Microsoft.Online.SharePoint.PowerShell
                    Write-Host Completed -foregroundcolor Green
                    $SPOModuleFound = $true
                }
                catch {
                    Write-Error "Error importing Microsoft.Online.SharePoint.PowerShell module: $($_.Exception.Message)"
                    $SPOModuleFound = $False
                }
            } 
        }
        else {
            Write-Error  "SharePoint Online module was not loaded. Run Install-Module Microsoft.Online.SharePoint.PowerShell as an Administrator"
            $SPOModuleFound = $false
        }
    }

    # Connect to Services - Skip if already connected. Connection check above
    Write-Host ""
    Write-Host "Establish Connections to Required Services" -foregroundcolor Cyan
    ## Connect to MsOnline
    if ($MSOnlineModuleFound -eq $true) {
        Write-Host "Connecting to MSOnline..." -NoNewline
        try {
            Connect-MsolService -ErrorAction Stop
            Write-Host "Completed" -foregroundcolor Green
            $msOnlineConnected = $true
        }
        catch {
            Write-Error "Error connecting to MSOnline: $($_.Exception.Message)"
        }
    }

    # Connect to ExchangeOnline
    if ($EXOModuleFound -eq $true) {
        Write-Host "Connecting to ExchangeOnline..." -NoNewline
        try {
            Connect-ExchangeOnline -UserPrincipalName $userPrincipalName -ErrorAction Stop *> Out-Null
            Write-Host "Completed" -foregroundcolor Green
            $exOnlineConnected = $true
        }
        catch {
            Write-Error "Error connecting to ExchangeOnline: $($_.Exception.Message)"
        }
    }

    # Connect to SharePoint Online
    if ($SPOModuleFound -eq $true) {
        Write-Host "Connecting to SharePoint Online..." -NoNewline
        try {
            Connect-SPOService -Url $SPOAdminURL -ErrorAction Stop
            Write-Host "Completed" -foregroundcolor Green
            $spoConnected = $true
        }
        catch {
            Write-Error "Error connecting to SharePoint Online: $($_.Exception.Message)"
        } 
    }

    #Full Check of Connected 
    if ($spoConnected -and $exOnlineConnected -and $msOnlineConnected) {
        Write-Host "Connected to required services." -ForegroundColor Green
    }
    else {
        Write-Error "Could not connect to required services. Please check the errors above."
    }
    Write-Host ""

    #Update Title Bar
    if ($exOnlineConnected) {
        $host.ui.RawUI.WindowTitle = $EXOOrgCheck.Name.tostring()
    }
    else {
        $EXOOrgCheck = Get-OrganizationConfig
        $host.ui.RawUI.WindowTitle = $EXOOrgCheck.Name.tostring()
    }
}

#used to scope number of mailboxes pulled for TESTING.  Set to 'Unlimited' for a full run or 5 for limited run
$resultSize = 'unlimited'

$global:start = Get-Date

#ProgressBar
$progresscounter = 1
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = ($allMailboxes).count
$ProgressPreference = "Continue"

#Hash Table to hold final report data
$allMailboxStats = [Ordered]@{}
$tenant = "Spectra"

#Array to store all errors encountered
$allErrors = @()

#Connect to all required O365 services for running this script
Connect-Office365RequiredServices

# Gather Mailboxes - Include InActive Mailboxes
Write-Host "Getting all mailboxes and inactive mailboxes..." -ForegroundColor Green
$allMailboxes = Get-EXOMailbox -ResultSize $resultSize -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -IncludeInactiveMailbox -PropertySets All -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "All mailboxes count: $($allMailboxes.count)"

# Gather Group Mailboxes
Write-Host "Getting all group mailboxes..." -ForegroundColor Green
$allGroupMailboxes += Get-Mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors
$allMailboxes += $allGroupMailboxes
Write-Host "All group mailboxes count: $($allGroupMailboxes.count)"
Write-Host ""

#Get Office 365 Group / Group Mailbox data with SharePoint URL data
Write-Host "Getting all unified groups (including soft deleted)..." -ForegroundColor Green
$allUnifiedGroups = Get-UnifiedGroup -resultSize $resultSize -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "All unified groups count: $($allUnifiedGroups.count)"
Write-Host ""

#Get Public Folder Data
Write-Host "Getting public folders..." -ForegroundColor Green
$PublicFolders = get-publicfolder -recurse -resultSize $resultSize -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "All public folders count: $($PublicFolders.count)"

Write-Host "Getting public folder statistics..." -ForegroundColor Green
$PublicFolderStatistics = get-publicfolder -recurse -resultSize $resultSize | get-publicfolderstatistics -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "All public folder statistics count: $($PublicFolderStatistics.count)"

Write-Host "Getting public folder client permissions..." -ForegroundColor Green
$PublicFolderPermissions = get-publicfolder -recurse -resultSize $resultSize | get-publicfolderclientpermission -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "All public folder permissions count: $($PublicFolderPermissions.count)"

###########################################################################################################################################

#Public Folder Data to Hash Table

###########################################################################################################################################

# Public Folders
#*****************

$PublicFoldersHash = @{}

foreach($pf in $PublicFolders) {

    $key = $pf.EntryId
    $value = $pf

    $PublicFoldersHash.add($key, $value)

}

# Public Folder Statistics
#**************************

$PublicFolderStatsHash = @{}

foreach($publicFolderStat in $PublicFolderStatistics) {

    $key = $PublicFolderStat.EntryId
    $value = $PublicFolderStat

    $PublicFolderStatsHash.add($key, $value)

}

# Public Folder Permissions
#***************************
$start = Get-Date
$PublicFolderPermissionsHash = @{}

Write-Host "Processing Public Folder Permissions..."

foreach($publicFolderPermission in $PublicFolderPermissions) {

    Write-ProgressHelper -Activity "Gathering Public Folder Permissions for $($publicFolderPermission.Identity)" -ProgressCounter ($progresscounter++) -TotalCount $PublicFolderPermissions.Count -ID 1 -StartTime $start
    

    $key = $publicFolderPermission.Identity

    if($PublicFolderPermissionsHash.ContainsKey($key))
    {   
        $user = $publicFolderPermission.User
        $primarySMTP = $publicFolderPermission.User.RecipientPrincipal.PrimarySmtpAddress
        $AccessRights = $publicFolderPermission.AccessRights

        $currentValue = $PublicFolderPermissionsHash[$key]

        if($primarySMTP)
        {
            $value = $currentValue += ",($($User) - $($primarySMTP) - $($AccessRights))"
        }
        else
        {
            $value = $currentValue += ",($($User) - $($AccessRights))"
        }

        $PublicFolderPermissionsHash[$key] = $value
    }

    else
    {
        $user = $publicFolderPermission.User
        $primarySMTP = $publicFolderPermission.User.RecipientPrincipal.PrimarySmtpAddress
        $AccessRights = $publicFolderPermission.AccessRights

        if($primarySMTP)
        {
            $value = "($($User) - $($primarySMTP) - $($AccessRights))"
        }
        else
        {
            $value = "($($User) - $($AccessRights))"
        }

        $PublicFolderPermissionsHash.add($key, $value)
    }

}


###########################################################################################################################################

#Unified Group Data to Hash Table

###########################################################################################################################################

Write-Host ""
Write-Host "Adding Unified Group data to Hash..." -ForegroundColor Green

$unifiedGroupHash = @{}


foreach ($group in $allUnifiedGroups) {
    
    $key = $group.ExchangeGuid
    $value = $group

    $unifiedGroupHash.Add($key, $value)

}
Write-Host "Unified group hash count: $($unifiedGroupHash.count)"

###########################################################################################################################################

#Group Mailbox Data to Hash Table

###########################################################################################################################################

Write-Host ""
Write-Host "Adding Group Mailbox Data to Hash..." -ForegroundColor Green
$groupMailboxHash = @{}

foreach ($groupMailbox in $allGroupMailboxes) {
    
    $key = $groupMailbox.ExchangeGuid
    $value = $groupMailbox

    $groupMailboxHash.Add($key, $value)

}

Write-Host "Group Mailbox Hash Count: $($groupMailboxHash.count)"


###########################################################################################################################################

#Get all OneDrive Personal Sites data and add to hash table with UserPrincipalName as HASH KEY and OneDrive data as HASH VALUE

###########################################################################################################################################

$OneDriveDataHash = @{}

$SharePointSiteHash = @{}

try
{
    Write-Host ""
    Write-Host "Getting all OneDrive site data..." -ForegroundColor Green
    $OneDriveSite = Get-SPOSite -Filter "URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop -ErrorVariable +allErrors

    #Get all SharePoint sites for associating with Office 365 Groups / GroupMailboxes
    Write-Host "Getting all SharePoint site data..." -ForegroundColor Green
    $SharePointSite = Get-SPOSite -IncludePersonalSite $false -Limit all -ErrorAction Stop -ErrorVariable +allErrors

}

catch
{
    Write-Host "Error getting all OneDrive sites data or SharePoint sites data."
}

#OneDrive data to OneDrive Hash
#************************************************************************************
foreach ($site in $OneDriveSite) {

    #Add to hash table
    $key = $site.Owner
    $value = $site

    if($OneDriveDataHash.ContainsKey($key))
    {
        Write-Host "Owner $($key) already in hash with current value" $OneDriveDataHash[$key].URL -ForegroundColor Yellow
        Write-Host "New url found: $($value.URL)" -ForegroundColor DarkMagenta
        #if key already exists in hash table, get current value and add new data
        $currentValue = $OneDriveDataHash[$key]
        $newValue = $currentValue
        $newValue.URL += ", $($value.URL)"
        $OneDriveDataHash.Remove($key)
        $OneDriveDataHash.add($key, $newValue)

        
    }
    else
    {
        #if key doesn't exist in hash table, add URL discovered for this key
        $OneDriveDataHash.Add($key, $value)
    }
}

#SharePoint data to SharePoint Hash
#************************************************************************************
foreach ($site in $SharePointSite) {

    #Add to hash table
    $key = $site.URL
    $value = $site

    #add to hash table
    $SharePointSiteHash.Add($key, $value)
  
}

Write-Host "OneDrive sites founds: $($OneDriveDataHash.Count)"
Write-Host "SharePoint sites founds: $($SharePointSiteHash.Count)"

###########################################################################################################################################

#MSOL User Data to Hash Table

###########################################################################################################################################

$msolUserHash = @{}

try
{

    Write-Host ""
    Write-Host "Getting all MSOL User data..." -ForegroundColor Green
    $msolUsers = Get-MsolUser -All -ErrorAction Stop -ErrorVariable +allErrors
    #Add deleted user information for inactive mailbox msol user data
    $msolUsers += Get-MsolUser -All -ErrorAction Stop -ErrorVariable +allErrors -ReturnDeletedUsers

}
 
catch
{
    Write-Host "Error getting all OneDrive sites data."
}

foreach ($user in $msolUsers) {

    #Add to hash table
    $key = $user.UserPrincipalName
    $value = $user

    $msolUserHash.Add($key, $value)
}

Write-Host "MSOL users founds: $($msolUserHash.Count)"

###########################################################################################################################################

#Mailbox Statistics to Hash Table

###########################################################################################################################################

$mailboxStatsHash = @{}

Write-Host ""
Write-Host "Getting primary mailbox stats..." -ForegroundColor Green


$primaryMailboxStats = get-exomailbox -ResultSize $resultSize -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors | Get-EXOMailboxStatistics -Properties LastLogonTime -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
Write-Host "Getting group mailbox stats..." -ForegroundColor Green
$primaryMailboxStats += get-mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction Continue -ErrorVariable +allErrors | Get-MailboxStatistics -ErrorAction Continue -ErrorVariable +allErrors

#process $primaryMailboxStats to add data to hash table with Key as MailboxGuid and Value as Array of Statistics Data
Write-Host "Adding archive stats to hash table..."
$primaryMailboxStats | foreach {
    
    $key = $_.MailboxGuid
    $value = $_
    $mailboxStatsHash.add($key, $value)

}

Write-Host "Individual Mailbox stats founds: $($mailboxStatsHash.Count)"
Write-Host ""

###########################################################################################################################################

#Archive Mailbox Stats to Hash Table

###########################################################################################################################################

$archiveMailboxStatsHash = @{}

Write-Host "Getting archive mailbox stats..." -ForegroundColor Green
$archiveMailboxStats = get-exomailbox -ResultSize $resultSize -Properties ArchiveStatus -IncludeInactiveMailbox | ? {$_.ArchiveStatus -ne "None"}| Get-EXOMailboxStatistics -Archive -Properties LastLogonTime -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
Write-Host "Getting group mailbox archive stats..." -ForegroundColor Green
$archiveMailboxStats += get-mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox | ? {$_.ArchiveStatus -ne "None"}| Get-MailboxStatistics -Archive -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients


Write-Host "Adding archive stats to hash table..."
#process $archiveMailboxStats to add data to hash table with Key as MailboxGuid and Value as Array of Statistics Data

$archiveMailboxStats | foreach {
    
    $key = $_.MailboxGuid
    $value = $_
    $archiveMailboxStatsHash.add($key, $value)

}


Write-Host "Archive Mailbox stats founds: $($archiveMailboxStatsHash.Count)"
Write-Host ""
Write-Host "Consolidating report data for each user..." -ForegroundColor Green


###########################################################################################################################################

#Consolidate Reporting for each user

###########################################################################################################################################

#ProgressBar
$start = Get-Date
$progresscounter = 1
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = ($allMailboxes).count
$ProgressPreference = "Continue"


foreach ($user in $allMailboxes) {
    #progress bar
    Write-ProgressHelper -Activity "Gathering Mailbox Details for $($user.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -ID 1 -StartTime $start
    
    #Pull MailboxStats and UserDetails
    #*******************************************************************************************************************

    if($user.ExchangeGuid)
    {
        $mbxStats = $mailboxStatsHash[$user.ExchangeGuid]
    }

    if($user.UserPrincipalName)
    {
        $msoluser = $msolUserHash[$user.UserPrincipalName]
        $oneDriveData = $OneDriveDataHash[$user.UserPrincipalName]
    }

    #If $user represents a GroupMailbox object set $msolUser data to pull from GroupMailboxHash which contains group mailbox data
    if($user.RecipientTypeDetails -eq "GroupMailbox")
    {
        $msoluser = $groupMailboxHash[$user.ExchangeGuid]
        $unifiedGroupData = $unifiedGroupHash[$user.ExchangeGuid]
        $groupMailboxData = $groupMailboxHash[$user.ExchangeGuid]

        if($unifiedGroupData)
        {
            $sharePointSiteData = $SharePointSiteHash[($unifiedGroupData.SharePointSiteUrl)]
        }
    }

    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    #get mailbox size in GB if TotalItemSize exists - null values break the hash table creation
    
    if($mbxStats.TotalItemSize)
    {
        $MBXSizeGB = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
    }
    else
    {
        $MBXSizeGB = 0
    }

    # Create User Hash Table
    #*******************************************************************************************************************

    if($user.RecipientTypeDetails -notcontains "Group")
    {
        $currentuser = [ordered]@{
            "DisplayName" = $msoluser.DisplayName
            "UserPrincipalName" = $msoluser.userprincipalname
            "IsLicensed" = $msoluser.IsLicensed
            "Licenses" = ($msoluser.Licenses.AccountSkuID -join ";")
            "License-DisabledArray" = ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";")
            "BlockCredential" = $msoluser.BlockCredential
            "IsInactiveMailbox" = $user.IsInactiveMailbox
            "Department" = $msoluser.Department
            "RecipientTypeDetails" = $user.RecipientTypeDetails
            "PrimarySmtpAddress" = $user.PrimarySmtpAddress
            "Alias" = $user.alias
            "WhenCreated" = $user.WhenCreated
            "LastLogonTime" = $mbxStats.LastLogonTime
            "EmailAddresses" = ($EmailAddresses -join ";")
            "LegacyExchangeDN" = ("x500:" + $user.legacyexchangedn)
            "HiddenFromAddressListsEnabled" = $user.HiddenFromAddressListsEnabled
            "DeliverToMailboxAndForward" = $user.DeliverToMailboxAndForward
            "ForwardingAddress" = $user.ForwardingAddress
            "ForwardingSmtpAddress" = $user.ForwardingSmtpAddress
            "MBXSize" = $MBXStats.TotalItemSize
            "MBXSize_GB" = $MBXSizeGB
            "MBXItemCount" = $MBXStats.ItemCount
            "LitigationHoldEnabled" = $user.LitigationHoldEnabled
            "LitigationHoldDuration" = $user.LitigationHoldDuration
            "InPlaceHolds" = $user.InPlaceHolds -join ";"
            "ArchiveStatus" = $user.ArchiveStatus
            "RetentionPolicy" = $user.RetentionPolicy
            #Public Folder Fields
            "PF_Identity" = $null
            "PF_MailEnabled" = $null
            "PF_ParentPath" = $null
            "PF_HasSubfolders" = $null
            "PF_ContentMailbox" = $null
            "LastModified" = $null
            "PF_ItemCount" = $null
            "PF_TotalItemSize" = $null
            "PF_MailboxOwnerID" = $null
            "PF_Permissions" = $null

        }
    }

    else
    {
        $currentuser = [ordered]@{
            "DisplayName" = $msoluser.DisplayName
            #User Fields set to Null
            "UserPrincipalName" = $null
            "IsLicensed" = $null
            "Licenses" = $null
            "License-DisabledArray" = $null
            "BlockCredential" = $null
            "IsInactiveMailbox" = $user.IsInactiveMailbox
            "Department" = $null
            #End User fields
            "RecipientTypeDetails" = $user.RecipientTypeDetails
            "PrimarySmtpAddress" = $user.PrimarySmtpAddress
            "Alias" = $user.alias
            "WhenCreated" = $user.WhenCreated
            "LastLogonTime" = $mbxStats.LastLogonTime
            "EmailAddresses" = ($EmailAddresses -join ";")
            "LegacyExchangeDN" = ("x500:" + $user.legacyexchangedn)
            "HiddenFromAddressListsEnabled" = $user.HiddenFromAddressListsEnabled
            "DeliverToMailboxAndForward" = $user.DeliverToMailboxAndForward
            "ForwardingAddress" = $user.ForwardingAddress
            "ForwardingSmtpAddress" = $user.ForwardingSmtpAddress
            "MBXSize" = $MBXStats.TotalItemSize
            "MBXSize_GB" = $MBXSizeGB
            "MBXItemCount" = $MBXStats.ItemCount
            "LitigationHoldEnabled" = $user.LitigationHoldEnabled
            "LitigationHoldDuration" = $user.LitigationHoldDuration
            "InPlaceHolds" = $user.InPlaceHolds -join ";"
            "ArchiveStatus" = $user.ArchiveStatus
            "RetentionPolicy" = $user.RetentionPolicy
            #Public Folder Fields
            "PF_Identity" = $null
            "PF_MailEnabled" = $null
            "PF_ParentPath" = $null
            "PF_HasSubfolders" = $null
            "PF_ContentMailbox" = $null
            "LastModified" = $null
            "PF_ItemCount" = $null
            "PF_TotalItemSize" = $null
            "PF_MailboxOwnerID" = $null
            "PF_Permissions" = $null
        }
    }

    #Pull Send on Behalf
    #*******************************************************************************************************************
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm -ErrorVariable +allErrors).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser["GrantSendOnBehalfTo"] = ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    #*******************************************************************************************************************

    #Exclude Inactive Mailbox to avoid errors
    if($user.IsInactiveMailbox -eq $false)
    {

        if ($mbxPermissions = Get-MailboxPermission $user.DistinguishedName -ErrorVariable +allErrors -ErrorAction SilentlyContinue | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
            $currentuser["FullAccessPerms"] = ($mbxPermissions.User -join ";")
        }
        else {$currentuser["FullAccessPerms"] = ($null)}
        # Mailbox Send As Check
        if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.DistinguishedName -ErrorVariable +allErrors -ErrorAction SilentlyContinue  | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
            $currentuser["SendAsPerms"] = ($sendAsPermsCheck.trustee -join ";")
        }
        else {$currentuser["SendAsPerms"] = ($null)}

    }

    # Archive Mailbox Check
    #*******************************************************************************************************************
    if ($user.ArchiveStatus -ne "None" -and $user.ArchiveStatus -ne $null)
    {
        $archiveStats = $archiveMailboxStatsHash[$user.ArchiveGuid]

        $currentuser["ArchiveSize"] = $ArchiveStats.TotalItemSize.Value

        if($ArchiveStats.TotalItemSize)
        {
            $currentuser["ArchiveSize-GB"] = [math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
        }
        else
        {
            $currentuser["ArchiveSize-GB"] = $null 
        }

        $currentuser["ArchiveItemCount"] = $ArchiveStats.ItemCount
    }

    else
    {
        $currentuser["ArchiveSize"] = $null
        $currentuser["ArchiveSize-GB"] = $null
        $currentuser["ArchiveItemCount"] = $null
    }

    #Get OneDrive URL if this user is an owner of a OneDrive site - If user is not inactive
    #*******************************************************************************************************************
    #Errors in if statement if userprincipalname is null

    if($OneDriveData)# -and ($user.IsInactiveMailbox -eq $false))
    {
        $currentuser["OneDriveURL_$($tenant)"] = $OneDriveData.URL
        $currentuser["OneDriveStorage-GB_$($tenant)"] = [math]::Round($OneDriveData.StorageUsageCurrent / 1024, 3)
        $currentuser["OneDriveLastContentModifiedDate"] = $OneDriveData.LastContentModifiedDate
    }

    #Group Mailbox Associated SharePoint Site mapping
    if($sharePointSiteData)
    {
        
        $currentuser["OneDriveURL_$($tenant)"] = $sharePointSiteData.URL
        $currentuser["OneDriveStorage-GB_$($tenant)"] = [math]::Round($sharePointSiteData.StorageUsageCurrent / 1024, 3)
        $currentuser["OneDriveLastContentModifiedDate"] = $sharePointSiteData.LastContentModifiedDate
    }

    if(!($oneDriveData) -and !($sharePointSiteData))
    {

        $currentuser["OneDriveURL_$($tenant)"] = $null
        $currentuser["OneDriveStorage-GB_$($tenant)"] = $null
        $currentuser["OneDriveLastContentModifiedDate"] = $null
    }

    #Combine all the data into one hash table
    #*******************************************************************************************************************
    $allMailboxStats[$User.PrimarySMTPAddress] = $currentuser
}


###########################################################################################################################################

#Add Public Folder Data to Reporting

###########################################################################################################################################

#Seperate for each for Public Folder data to avoid unnecessary error handling

#reset progress counter for next operation
$start = Get-Date
$progresscounter = 1
Write-Host "Processing public folders..." -ForegroundColor Green

foreach($pf in $PublicFolders) {

    Write-ProgressHelper -Activity "Gathering Public Folder Details for $($pf.Name)" -ProgressCounter ($progresscounter++) -TotalCount $PublicFolders.Count -ID 1 -StartTime $start
    
    $entryId = $pf.EntryId

    $currentuser = [ordered]@{
            "DisplayName" = $pf.Name
            "PF_Identity" = $pf.Identity
            "PF_MailEnabled" = $pf.MailEnabled
            "PF_ParentPath" = $pf.ParentPath
            "PF_HasSubfolders" = $pf.HasSubfolders
            "PF_ContentMailbox" = $pf.ContentMailboxName
            "WhenCreated" = $PublicFolderStatsHash[$entryId].CreationTime
            "LastModified" = $PublicFolderStatsHash[$entryId].LastModifiedTime
            "PF_ItemCount" = $PublicFolderStatsHash[$entryId].ItemCount
            "PF_TotalItemSize" = $PublicFolderStatsHash[$entryId].TotalItemSize
            "PF_MailboxOwnerID" = $PublicFolderStatsHash[$entryId].MailboxOwnerID
            "PF_Permissions" = $PublicFolderPermissionsHash[$pf.Identity]
        }

        $allMailboxStats[$entryId] = $currentuser
}


#Convert Hash Table to Custom Object Array for Export
#*******************************************************************************************************************

Write-Host "Converting Hash to Array for Export..."

$ExportAllMailboxStatsArray = @()
foreach ($key in $allMailboxStats.Keys) {
    $attributes = $allMailboxStats[$key]
    $customObject = New-Object -TypeName PSObject

    foreach ($attribute in $attributes.Keys) {
        $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)__$($tenant)" -Value $attributes[$attribute]
    }

    $ExportAllMailboxStatsArray += $customObject
}

$ExportAllMailboxStatsArray | Export-Csv $pwd\$($tenant)-AllMailboxStatsv4.csv -NoTypeInformation

Write-Host "Total number of report results: $($ExportAllMailboxStatsArray.count)"
Write-Host "Number Errors: $($allErrors.count)"

if($allErrors)
{
    $allErrors | Export-Csv $pwd\$($tenant)-errorReport.csv -NoTypeInformation
}

Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan