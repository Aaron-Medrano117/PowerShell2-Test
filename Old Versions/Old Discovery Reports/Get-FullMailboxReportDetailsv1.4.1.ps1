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

    v1.3.4
    Author: Aaron Medrano
    Updated MailboxStats to include Deleted Recipients
    Updated Connect-Office365RequiredServices function to include MGGraph connection
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
    $global:secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)

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

    ## MGGraph Connection, module Import and installation Check
    Try {
        $MSOCompanyCheck = Get-MgOrganization -ErrorAction Stop
        Write-Host "Already Connected to MGGraph: $($MSOCompanyCheck.DisplayName)" -foregroundcolor Green
        $msOnlineConnected = $true
    }
    Catch {
        Write-Host "MGGraph module..." -NoNewline
        if ((Get-InstalledModule -Name "Microsoft.Graph.Authentication" -ErrorAction SilentlyContinue) -ne $null) {
            if ((Get-Module -Name "MGGraph") -ne $null) {
                Write-Host "Already Imported" -ForegroundColor Green
                $MgGraphModuleFound = $true
            }
            else {
                Write-Host "was not loaded. Importing module..." -NoNewline -ForegroundColor Yellow
                try {
                    Import-Module MGGraph
                    Write-Host Completed -foregroundcolor Green
                    $MgGraphModuleFound = $true
                }
                catch {
                    Write-Error "Error importing MGGraph module: $($_.Exception.Message)"
                    $MgGraphModuleFound = $False
                }
            } 
        }
        else {
            Write-Error  "MGGraph module was not loaded. Run Install-Module MSOnline as an Administrator"
            $MgGraphModuleFound = $false
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

    ## Connect to MgGraph
    if ($MgGraphModuleFound -eq $true) {
        Write-Host "Connecting to MgGraph..." -NoNewline
        try {
            $RequiredScopes = @(
                "Directory.Read.All"
                "Directory.ReadWrite.All"
                "Organization.Read.All"
                "Organization.ReadWrite.All"
                "AuditLog.Read.All"
                "Directory.AccessAsUser.All"
                "EAS.AccessAsUser.All"
                "EWS.AccessAsUser.All"
                "IdentityRiskyUser.Read.All"
                "IdentityRiskyUser.ReadWrite.All"
                "IdentityUserFlow.Read.All"
                "IdentityUserFlow.ReadWrite.All"
                "IMAP.AccessAsUser.All"
                "POP.AccessAsUser.All"
                "TeamsAppInstallation.ReadForUser"
                "TeamsAppInstallation.ReadWriteForUser"
                "TeamsAppInstallation.ReadWriteSelfForUser"
                "TeamsTab.ReadWriteForUser"
                "TeamsTab.ReadWriteSelfForUser"
                "User.EnableDisableAccount.All"
                "User.Export.All"
                "User.Invite.All"
                "User.ManageIdentities.All"
                "User.Read.All"
                "User.ReadBasic.All"
                "User.ReadWrite.All"
                "UserActivity.ReadWrite.CreatedByApp"
                "UserAuthenticationMethod.Read.All"
                "UserAuthenticationMethod.ReadWrite"
                "UserAuthenticationMethod.ReadWrite.All"
                "User-LifeCycleInfo.Read.All"
                "User-LifeCycleInfo.ReadWrite.All"
                "CrossTenantUserProfileSharing.Read.All"
                "CrossTenantUserProfileSharing.ReadWrite.All"
                "IdentityRiskyUser.Read.All"
                "IdentityRiskyUser.ReadWrite.All"
                "IdentityUserFlow.Read.All"
                "IdentityUserFlow.ReadWrite.All"
                "User.EnableDisableAccount.All"
                "User.Export.All"
                "User.Invite.All"
                "User.ManageIdentities.All"
                "User.Read.All"
                "User.ReadWrite.All"
                "UserAuthenticationMethod.Read.All"
                "UserAuthenticationMethod.ReadWrite.All"
                "User-LifeCycleInfo.Read.All"
                "User-LifeCycleInfo.ReadWrite.All"
            )
            $result = Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
            Write-Host "Completed" -foregroundcolor Green
            $msOnlineConnected = $true
        }
        catch {
            Write-Error "Error connecting to MgGraph: $($_.Exception.Message)"
        }
    }

    # Connect to ExchangeOnline
    if ($EXOModuleFound -eq $true) {
        Write-Host "Connecting to ExchangeOnline..." -NoNewline
        try {
            $result = Connect-ExchangeOnline -UserPrincipalName $userPrincipalName -ErrorAction Stop *> Out-Null
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

#Connect to all required O365 services for running this script
Connect-Office365RequiredServices

#used to scope number of mailboxes pulled for TESTING.  Set to 'Unlimited' for a full run or 5 for limited run
$resultSize = 'unlimited'

$global:initialStart = Get-Date

#ProgressBar
$progresscounter = 1
[nullable[double]]$global:secondsRemaining = $null
$ProgressPreference = "Continue"

#Hash Table to hold final report data
$allMailboxStats = [Ordered]@{}
$tenantStatsHash= [Ordered]@{}
$tenant = "Spectra"

#Array to store all errors encountered
$allErrors = @()

#Array to hold tenant stats
$tenantStats = @()

# Gather Mailboxes - Include InActive Mailboxes
#Write-Progress -Activity "Getting all group mailboxes" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
$start = Get-Date

Write-Host "Getting all mailboxes and inactive mailboxes..." -ForegroundColor Cyan -nonewline
$allMailboxes = Get-EXOMailbox -ResultSize $resultSize -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -IncludeInactiveMailbox -PropertySets addresslist, archive, delivery, minimum -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

# Gather Group Mailboxes
$start = Get-Date
Write-Host "Getting all group mailboxes..." -ForegroundColor Cyan -nonewline
$allGroupMailboxes += Get-Mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

#Get Office 365 Group / Group Mailbox data with SharePoint URL data
#Write-Progress -Activity "Getting all unified groups (including soft deleted)" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
$start = Get-Date
Write-Host "Getting all unified groups (including soft deleted)..." -ForegroundColor Cyan -nonewline
$allUnifiedGroups = Get-UnifiedGroup -resultSize $resultSize -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

#Get Public Folder Data, Statistics, and Permissions
#Write-Progress -Activity "Getting public folders, Stats and Perms" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
$start = Get-Date
Write-Host "Getting public folders, Stats and Perms ..." -ForegroundColor Cyan -nonewline
$PublicFolders = get-publicfolder -recurse -resultSize $resultSize -ErrorAction SilentlyContinue -ErrorVariable +allErrors
$PublicFolderStatistics = get-publicfolder -recurse -resultSize $resultSize | get-publicfolderstatistics -ErrorAction SilentlyContinue -ErrorVariable +allErrors
$PublicFolderPermissions = get-publicfolder -recurse -resultSize $resultSize | get-publicfolderclientpermission -ErrorAction SilentlyContinue -ErrorVariable +allErrors
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green


###########################################################################################################################################

#Public Folder Data to Hash Table

###########################################################################################################################################


# Public Folders
#*****************

$PublicFoldersHash = @{}

foreach($pf in $PublicFolders) {

    $key = $pf.EntryId
    $value = $pf

    $PublicFoldersHash[$key] = $value

}

# Public Folder Statistics
#**************************

$PublicFolderStatsHash = @{}

foreach($publicFolderStat in $PublicFolderStatistics) {

    $key = $PublicFolderStat.EntryId
    $value = $PublicFolderStat

    $PublicFolderStatsHash[$key] = $value

}

# Public Folder Permissions
#***************************

#Progress Bar Parameters Reset
$start = Get-Date
$progresscounter = 1
$totalCount = $PublicFolderPermissions.count


$PublicFolderPermissionsHash = @{}

Write-Host "Processing Public Folder Permissions..." -ForegroundColor Cyan -nonewline

foreach($publicFolderPermission in $PublicFolderPermissions) {

    Write-ProgressHelper -Activity "Gathering Public Folder Permissions for $($publicFolderPermission.Identity)" -ProgressCounter ($progresscounter++) -TotalCount $totalCount -ID 2 -StartTime $start
    

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

        $PublicFolderPermissionsHash[$key] = $value
    }

}
Write-Host "Completed" -ForegroundColor Green

###########################################################################################################################################

# License SKUs and Service Plan IDs to HASH

###########################################################################################################################################

#Write-Progress -Activity "Adding License SKUs and Service Plan IDs to Hash" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))

Write-Host "Adding License SKUs and Service Plan IDs to Hash..." -ForegroundColor Cyan -nonewline
$start = Get-Date
# Get License SKUs
$skus = Get-MgSubscribedSku -ErrorAction Continue -ErrorVariable +allErrors

#Build a hashtable for looking up license names from license sku
$licenseSkuHash = @{}

#store Service Plan IDs and corresponding ServicePlan data for each ID
$servicePlanHash = @{}

#Add License SKUs to Hash Table and Service Plans under each SKU to another Hash Table
foreach ($sku in $skus) {
    $licenseSkuHash[$sku.SkuId] = $sku

    foreach($servicePlan in $sku.ServicePlans)
    {
        $key = $servicePlan.ServicePlanId

        #If plan id has NOT already been added to hash
        if(!($servicePlanHash.ContainsKey($key)))
        {
            
            $value = $servicePlan.ServicePlanName
            $servicePlanHash[$key] = $value
        }

        else
        {
            #Move to next object
            Continue
        }
    }
}

Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green


###########################################################################################################################################

#Unified Group Data to Hash Table

###########################################################################################################################################

Write-Host "Adding Unified Group data to Hash..." -ForegroundColor Cyan -nonewline
$start = Get-Date
$unifiedGroupHash = @{}
foreach ($group in $allUnifiedGroups) {
    
    $key = $group.ExchangeGuid
    $value = $group

    $unifiedGroupHash[$key] = $value

}
#Write-Host "Unified group hash count: $($unifiedGroupHash.count)"
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

###########################################################################################################################################

#Group Mailbox Data to Hash Table

###########################################################################################################################################

Write-Host "Adding Group Mailbox Data to Hash..." -ForegroundColor Cyan -nonewline
$start = Get-Date
$groupMailboxHash = @{}

foreach ($groupMailbox in $allGroupMailboxes) {
    
    $key = $groupMailbox.ExchangeGuid
    $value = $groupMailbox

    $groupMailboxHash[$key] = $value

}

#Write-Host "Group Mailbox Hash Count: $($groupMailboxHash.count)"
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green


###########################################################################################################################################

#Get all OneDrive Personal Sites data and add to hash table with UserPrincipalName as HASH KEY and OneDrive data as HASH VALUE

###########################################################################################################################################
$start = Get-Date

$OneDriveDataHash = @{}
$SharePointSiteHash = @{}

try {
    #Write-Host ""
    Write-Host "Getting all OneDrive site data..." -ForegroundColor Cyan -nonewline
    $OneDriveSite = Get-SPOSite -Filter "URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop -ErrorVariable +allErrors

    #Get all SharePoint sites for associating with Office 365 Groups / GroupMailboxes
    Write-Host "Getting all SharePoint site data..." -ForegroundColor Cyan -nonewline
    $SharePointSite = Get-SPOSite -IncludePersonalSite $false -Limit all -ErrorAction Stop -ErrorVariable +allErrors

}
catch {
    Write-Host "Error getting all OneDrive sites data or SharePoint sites data." -ForegroundColor Red
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
        $OneDriveDataHash[$key] = $value
    }
}

#SharePoint data to SharePoint Hash
#************************************************************************************
foreach ($site in $SharePointSite) {

    #Add to hash table
    $key = $site.URL
    $value = $site

    #add to hash table
    $SharePointSiteHash[$key] = $value
  
}
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

###########################################################################################################################################

#Microsoft Graph (mg) User Data to Hash Table

###########################################################################################################################################
$start = Get-Date
$mgUserHash = @{}

try {
    #Write-Host ""
    Write-Host "Getting all Microsoft Graph User data..." -ForegroundColor Cyan -nonewline

    #MS GRAPH update
    $mgUsers = Get-MGUser -All -ErrorAction Stop -ErrorVariable +allErrors

}
catch {
    Write-Host "Error getting all Microsoft Graph sites data." -ForegroundColor Red
}

foreach ($user in $mgUsers) {
    #Add to hash table
    $key = $user.UserPrincipalName
    $value = $user

    $mgUserHash[$key] = $value
}
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

###########################################################################################################################################

#Primary Mailbox Statistics to Hash Table

###########################################################################################################################################
$start = Get-Date
$mailboxStatsHash = @{}

#Write-Host ""
Write-Host "Getting primary mailbox stats..." -ForegroundColor Cyan -nonewline
Write-Progress -Activity "Adding primary mailbox stats" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))

$primaryMailboxStats = get-exomailbox -ResultSize $resultSize -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors | Get-EXOMailboxStatistics -Properties LastLogonTime -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
Write-Host "Getting group primary mailbox stats..." -ForegroundColor Green
$primaryMailboxStats += get-mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction Continue -ErrorVariable +allErrors | Get-MailboxStatistics -ErrorAction Continue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients

#process $primaryMailboxStats to add data to hash table with Key as MailboxGuid and Value as Array of Statistics Data
$primaryMailboxStats | foreach {
    $key = $_.MailboxGuid
    $value = $_
    $mailboxStatsHash[$key] = $value
}

Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

###########################################################################################################################################

#Archive Mailbox Stats to Hash Table

###########################################################################################################################################
$start = Get-Date
$archiveMailboxStatsHash = @{}

Write-Host "Getting archive mailbox stats..." -ForegroundColor Cyan -nonewline
$archiveMailboxStats = get-exomailbox -ResultSize $resultSize -Properties ArchiveStatus -IncludeInactiveMailbox | ? {$_.ArchiveStatus -ne "None"}| Get-EXOMailboxStatistics -Archive -Properties LastLogonTime -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
Write-Host "Getting group archive mailbox archive stats..." -ForegroundColor Cyan -nonewline
$archiveMailboxStats += get-mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox | ? {$_.ArchiveStatus -ne "None"}| Get-MailboxStatistics -Archive -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients


Write-Host
Write-Host "Adding all archive mailbox stats to hash table..." -ForegroundColor Green
#process $archiveMailboxStats to add data to hash table with Key as MailboxGuid and Value as Array of Statistics Data

$archiveMailboxStats | foreach {
    
    $key = $_.MailboxGuid
    $value = $_

    #errors if key is null
    if($key)
    {
        $archiveMailboxStatsHash[$key] = $value
    }

}
Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

Write-Host ""
Write-Host "Consolidating report data for each user / object..." -ForegroundColor Yellow
Write-Host ""


###########################################################################################################################################

#Consolidate Reporting for each user

###########################################################################################################################################

#ProgressBar
$start = Get-Date
$progresscounter = 1
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = ($allMailboxes).count
$ProgressPreference = "Continue"

Write-Host "Processing all non-group mailboxes..." -ForegroundColor Cyan -nonewline

#Process All Mailboxes - excluding group mailboxes which are processed in a seperate foreach
#*********************************************************************************************

foreach ($user in $allMailboxes) {
    #progress bar
    Write-ProgressHelper -Activity "Gathering Mailbox Details for $($user.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -ID 1 -StartTime $start

    #Pull MailboxStats and UserDetails
    #*******************************************************************************************************************
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    if($user.ExchangeGuid)
    {
        $mbxStats = $mailboxStatsHash[$user.ExchangeGuid]
    }
    if($user.UserPrincipalName -and $user.RecipientTypeDetails -ne "GroupMailbox")
    {
        $mgUser = $mgUserHash[$user.UserPrincipalName]
        $oneDriveData = $OneDriveDataHash[$user.UserPrincipalName]
    }

    #get mailbox size in GB if TotalItemSize exists - null values break the hash table creation 
    if($mbxStats.TotalItemSize)
    {
        $MBXSizeGB = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
    }
    else
    {
        $MBXSizeGB = 0
    }


    #Pull Licenses and Disabled Service Plans
    #*******************************************************************************************************************
    #Get licenses for this user
    $licenses = ($mgUser.AssignedLicenses)

    #Array for adding each license name to the allLicenses for this user
    $allLicenses = @()

    #Array for adding all disabled service names for this user
    $allDisabledPlans = @()

    #Process each license to get friendly names and disabled service plans for each license
    foreach($license in $licenses)
    {
        $licenseName = $licenseSkuHash[$license.SkuId].SkuPartNumber
        $allLicenses += $licenseName

        if($license.DisabledPlans)
        {
            foreach($disabledPlan in $license.DisabledPlans)
            {
                $disabledPlanName = $servicePlanHash[$disabledPlan]
                $allDisabledPlans += $disabledPlanName
            }
        }

    }

    # Create User Hash Table
    #*******************************************************************************************************************

    $currentuser = [ordered]@{
        "DisplayName" = $mgUser.DisplayName
        "UserPrincipalName" = $mgUser.userprincipalname
        "IsLicensed" = ($mgUser.AssignedLicenses.count -gt 0)
        "Licenses" = ($allLicenses -join ",")
        "License-DisabledArray" = ($allDisabledPlans -join ",")
        "AccountEnabled" = $mgUser.AccountEnabled
        "IsInactiveMailbox" = $user.IsInactiveMailbox
        "WhenSoftDeleted" = $user.WhenSoftDeleted
        "Department" = $mgUser.Department
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


    #Pull Send on Behalf
    #*******************************************************************************************************************
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Recipient $perm -ErrorVariable +allErrors).DisplayName
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

    #Get OneDrive URL if this user is an owner of a OneDrive site
    #*******************************************************************************************************************


    if($OneDriveData)# -and ($user.IsInactiveMailbox -eq $false))
    {
        $currentuser["OneDriveURL"] = $OneDriveData.URL
        $currentuser["OneDriveStorage-GB"] = [math]::Round($OneDriveData.StorageUsageCurrent / 1024, 3)
        $currentuser["OneDriveLastContentModifiedDate"] = $OneDriveData.LastContentModifiedDate
    }


    else
    {

        $currentuser["OneDriveURL"] = $null
        $currentuser["OneDriveStorage-GB"] = $null
        $currentuser["OneDriveLastContentModifiedDate"] = $null
    }

    #Combine all the data into one hash table
    #*******************************************************************************************************************

    $allMailboxStats[$User.PrimarySMTPAddress] = $currentuser

}

Write-Host "Completed" -ForegroundColor Green
###########################################################################################################################################

# Add Group Mailbox Data to Reporting

###########################################################################################################################################

#Progress Bar Parameters Reset
$start = Get-Date
$progresscounter = 1
$totalCount = $allGroupMailboxes.count

#Write-Host ""
Write-Host "Processing all group mailboxes..." -ForegroundColor Cyan -nonewline


foreach ($mailbox in $allGroupMailboxes) {
    #progress bar
    Write-ProgressHelper -Activity "Gathering Group Mailbox Details for $($mailbox.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -ID 1 -StartTime $start

        
    #Pull MailboxStats and UserDetails
    #*******************************************************************************************************************

    if($mailbox.ExchangeGuid)
    {
        $mbxStats = $mailboxStatsHash[$mailbox.ExchangeGuid]
    }

    #If $mailbox represents a GroupMailbox object set $mgUser data to pull from GroupMailboxHash which contains group mailbox data
    if($mailbox.RecipientTypeDetails -eq "GroupMailbox")
    {
        $mgUser = $groupMailboxHash[$mailbox.ExchangeGuid]
        $unifiedGroupData = $unifiedGroupHash[$mailbox.ExchangeGuid]

        if($unifiedGroupData.SharePointSiteUrl)
        {
            $sharePointSiteData = $SharePointSiteHash[($unifiedGroupData.SharePointSiteUrl)]
        }
    }

    $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses

    #get mailbox size in GB if TotalItemSize exists - null values break the hash table creation
    
    if($mbxStats.TotalItemSize)
    {
        $MBXSizeGB = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
    }
    else
    {
        $MBXSizeGB = 0
    }


    # Create Hash Table to add to Report Dataset
    #*******************************************************************************************************************
        $currentuser = [ordered]@{
            "DisplayName" = $mgUser.DisplayName
            "UserPrincipalName" = $mgUser.userprincipalname
            "IsLicensed" = ($mgUser.AssignedLicenses.count -gt 0)
            "Licenses" = ($allLicenses -join ",")
            "License-DisabledArray" = ($allDisabledPlans -join ",")
            "AccountEnabled" = $mgUser.AccountEnabled
            "IsInactiveMailbox" = $mailbox.IsInactiveMailbox
            "WhenSoftDeleted" = $mailbox.WhenSoftDeleted
            "Department" = $mgUser.Department
            "RecipientTypeDetails" = $mailbox.RecipientTypeDetails
            "PrimarySmtpAddress" = $mailbox.PrimarySmtpAddress
            "Alias" = $mailbox.alias
            "WhenCreated" = $mailbox.WhenCreated
            "LastLogonTime" = $mbxStats.LastLogonTime
            "EmailAddresses" = ($EmailAddresses -join ";")
            "LegacyExchangeDN" = ("x500:" + $mailbox.legacyexchangedn)
            "HiddenFromAddressListsEnabled" = $mailbox.HiddenFromAddressListsEnabled
            "DeliverToMailboxAndForward" = $mailbox.DeliverToMailboxAndForward
            "ForwardingAddress" = $mailbox.ForwardingAddress
            "ForwardingSmtpAddress" = $mailbox.ForwardingSmtpAddress
            "MBXSize" = $MBXStats.TotalItemSize
            "MBXSize_GB" = $MBXSizeGB
            "MBXItemCount" = $MBXStats.ItemCount
            "LitigationHoldEnabled" = $mailbox.LitigationHoldEnabled
            "LitigationHoldDuration" = $mailbox.LitigationHoldDuration
            "InPlaceHolds" = $mailbox.InPlaceHolds -join ";"
            "ArchiveStatus" = $mailbox.ArchiveStatus
            "RetentionPolicy" = $mailbox.RetentionPolicy
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
    

    #Pull Send on Behalf
    #*******************************************************************************************************************
    $grantSendOnBehalf = $mailbox.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Recipient $perm -ErrorVariable +allErrors).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser["GrantSendOnBehalfTo"] = ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    #*******************************************************************************************************************

    #Exclude Inactive Mailbox to avoid errors
    if($mailbox.IsInactiveMailbox -eq $false)
    {

        if ($mbxPermissions = Get-MailboxPermission $mailbox.DistinguishedName -ErrorVariable +allErrors -ErrorAction SilentlyContinue | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
            $currentuser["FullAccessPerms"] = ($mbxPermissions.User -join ";")
        }
        else {$currentuser["FullAccessPerms"] = ($null)}
        # Mailbox Send As Check
        if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $mailbox.DistinguishedName -ErrorVariable +allErrors -ErrorAction SilentlyContinue  | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
            $currentuser["SendAsPerms"] = ($sendAsPermsCheck.trustee -join ";")
        }
        else {$currentuser["SendAsPerms"] = ($null)}

    }

    # Archive Mailbox Check
    #*******************************************************************************************************************
    if ($mailbox.ArchiveStatus -ne "None" -and $mailbox.ArchiveStatus -ne $null)
    {
        $archiveStats = $archiveMailboxStatsHash[$mailbox.ArchiveGuid]

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

    #Group Mailbox Associated SharePoint Site mapping
    if($sharePointSiteData)
    {
        
        $currentuser["OneDriveURL"] = $sharePointSiteData.URL
        $currentuser["OneDriveStorage-GB"] = [math]::Round($sharePointSiteData.StorageUsageCurrent / 1024, 3)
        $currentuser["OneDriveLastContentModifiedDate"] = $sharePointSiteData.LastContentModifiedDate
    }

    else
    {

        $currentuser["OneDriveURL"] = $null
        $currentuser["OneDriveStorage-GB"] = $null
        $currentuser["OneDriveLastContentModifiedDate"] = $null
    }

    #Combine all the data into one hash table
    #*******************************************************************************************************************

    $allMailboxStats[$mailbox.PrimarySMTPAddress] = $currentuser

}

Write-Host "Completed" -ForegroundColor Green
###########################################################################################################################################

#Add Public Folder Data to Reporting

###########################################################################################################################################

#Seperate for each for Public Folder data to avoid unnecessary error handling

#reset progress counter for next operation
$start = Get-Date
$progresscounter = 1
$totalCount = $PublicFolders.Count


#Write-Host ""
Write-Host "Processing all public folders..." -ForegroundColor Cyan -nonewline

foreach($pf in $PublicFolders) {

    Write-ProgressHelper -Activity "Gathering Public Folder Details for $($pf.Name)" -ProgressCounter ($progresscounter++) -TotalCount $totalCount -ID 1 -StartTime $start
    
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
Write-Host "Completed" -ForegroundColor Green

#Convert Hash Table to Custom Object Array for Export
#*******************************************************************************************************************

Write-Host "Converting Hash to Array for Export..." -ForegroundColor Cyan -nonewline

$ExportAllMailboxStatsArray = @()
foreach ($key in $allMailboxStats.Keys) {
    $attributes = $allMailboxStats[$key]
    $customObject = New-Object -TypeName PSObject

    foreach ($attribute in $attributes.Keys) {
        $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)__$($tenant)" -Value $attributes[$attribute]
    }

    $ExportAllMailboxStatsArray += $customObject
}
Write-Host "Completed" -ForegroundColor Green

$ExportAllMailboxStatsArray | Export-Csv $pwd\$($tenant)-AllMailboxStats.csv -NoTypeInformation

Write-Host "Total number of report results: $($ExportAllMailboxStatsArray.count) | " -NoNewline
Write-Host "Full Report exported to $pwd\$($tenant)-AllMailboxStats.csv"


#Filter out blank lines added to errors
$allErrors = $allErrors | ? {$_.ErrorRecord}

if($allErrors)
{
    $allErrors | Export-Csv $pwd\$($tenant)-errorReport.csv -NoTypeInformation
}

Write-Host "Number Errors: $($allErrors.count) | " -NoNewline -ForegroundColor Red
Write-Host "Errors exported to $pwd\$($tenant)-errorReport.csv"
Write-Host ""
Write-Host "Completed in"((Get-Date) -$global:initialStart).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan
Write-Host ""

#Final Output of Recipient Counts
$tenantStatsHash["AllMailboxes"] = $allMailboxes.count
$tenantStatsHash["allGroupMailboxes"] = $allGroupMailboxes.count
$tenantStatsHash["allUnifiedGroups"] = $allUnifiedGroups.count
$tenantStatsHash["PublicFolders"] = $PublicFolders.count
$tenantStatsHash["PublicFolderPermissions"] = $PublicFolderPermissions.count
$tenantStatsHash["licenseSkus"] = $licenseSkuHash.count
$tenantStatsHash["servicePlanIDs"] = $servicePlanHash.count
$tenantStatsHash["OneDriveSites"] = $OneDriveDataHash.count
$tenantStatsHash["SharePointSites"] = $SharePointSiteHash.count
$tenantStatsHash["mgUsers"] = $mgUserHash.count
$tenantStatsHash["mailboxStats"] = $mailboxStatsHash.count
$tenantStatsHash["archiveMailboxStats"] = $archiveMailboxStatsHash.count


Write-Host "Recipient Count Table" -ForegroundColor White -BackgroundColor Green
$tenantStatsHash