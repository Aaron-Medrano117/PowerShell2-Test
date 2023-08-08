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

    v1.4.2
    Author: Aaron Medrano
    Updated HashTable updates from .Add() to [Key] = $value to reduce errors
    Updated Get-Mailbox in Permissions check to Get-Recipient to reduce errors
    Updated TenantStats from Array to HashTable and placed at end of script

    v1.4.3
    Author: Aaron Medrano
    Updated Required Scopes
    Updated Connect-Office365RequiredServices function to include MGGraph connection - proper variables used
    Updated Get-ExportPath function to properly check path and location
    Added SignInActivity for Users Array and Hashtable but commented out; takes too long
#>

###########################################################################################################################################

#Intial Variables and Functions

###########################################################################################################################################

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
        if ($null -ne (Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue)) {
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
        if ($null -ne (Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue)) {
            if ($null -ne (Get-Module -Name "MSOnline")) {
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
        $MGraphCompanyCheck = Get-MgOrganization -ErrorAction Stop
        if ($MGraphCompanyCheck.DisplayName -eq $MSOCompanyCheck.DisplayName){
            Write-Host "Already Connected to MGGraph: $($MGraphCompanyCheck.DisplayName)" -foregroundcolor Green
            $MGraphConnected = $true
        }
        else {
            Write-Host "Connected to Incorrect Tenant. Connected to MGGraph: $($MGraphCompanyCheck.DisplayName). Disconnecting." -foregroundcolor Yellow
            $result = Disconnect-MgGraph
            $MGraphConnected = $false
        }
    }
    Catch {
        Write-Host "MGGraph module..." -NoNewline
        if ($null -ne (Get-InstalledModule -Name Microsoft.Graph.* -ErrorAction SilentlyContinue)) {
            if ($null -ne (Get-Module -Name Microsoft.Graph.*)) {
                Write-Host "Already Imported" -ForegroundColor Green
                $MGraphModuleFound = $true
            }
            else {
                Write-Host "was not loaded. Importing module..." -NoNewline -ForegroundColor Yellow
                try {
                    Import-Module Microsoft.Graph
                    Write-Host Completed -foregroundcolor Green
                    $MGraphModuleFound = $true
                }
                catch {
                    Write-Error "Error importing MGGraph module: $($_.Exception.Message)"
                    $MGraphModuleFound = $False
                }
            } 
        }
        else {
            Write-Error  "MGGraph module was not loaded. Run Install-Module Microsoft.Graph as an Administrator"
            $MGraphModuleFound = $false
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
        if ($null -ne (Get-InstalledModule -Name "Microsoft.Online.SharePoint.PowerShell" -ErrorAction SilentlyContinue)) {
            if ($null -ne (Get-Module -Name "Microsoft.Online.SharePoint.PowerShell")) {
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
    if (($MGraphModuleFound -eq $true) -or ($MGraphConnected -eq $false)) {
        Write-Host "Connecting to MgGraph..." -NoNewline
            try {
                $RequiredScopes = @(
                    "Directory.ReadWrite.All"
                    "Organization.ReadWrite.All"
                    "AuditLog.Read.All"
                    "Directory.AccessAsUser.All"
                    "IdentityRiskyUser.ReadWrite.All"
                    "IdentityUserFlow.ReadWrite.All"
                    "EAS.AccessAsUser.All"
                    "EWS.AccessAsUser.All"
                    "TeamsAppInstallation.ReadWriteForUser"
                    "TeamsAppInstallation.ReadWriteSelfForUser"
                    "TeamsTab.ReadWriteForUser"
                    "TeamsTab.ReadWriteSelfForUser"
                    "User.EnableDisableAccount.All"
                    "User.Export.All"
                    "User.Invite.All"
                    "User.ManageIdentities.All"
                    "User.ReadWrite.All"
                    "UserActivity.ReadWrite.CreatedByApp"
                    "UserAuthenticationMethod.ReadWrite.All"
                    "User-LifeCycleInfo.ReadWrite.All"
                    "CrossTenantUserProfileSharing.ReadWrite.All"
                    "Device.Read.All"
                    "AuthenticationContext.ReadWrite.All"
                    "Policy.ReadWrite.AuthenticationMethod"
                    "Domain.ReadWrite.All"
                    "Group.ReadWrite.All"
                    "GroupMember.ReadWrite.All"
                    "IdentityRiskyUser.ReadWrite.All"
                    "LicenseAssignment.ReadWrite.All"
                    "SharePointTenantSettings.ReadWrite.All"
                    "Synchronization.ReadWrite.All"
                    "Team.ReadBasic.All"
                )
                $result = Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
                Select-MgProfile -Name "beta"
                Write-Host "Completed" -foregroundcolor Green
                $MGraphConnected = $true
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
function Install-ImportExcel {
    # Check if ImportExcel module is installed
    if (!(Get-Module -ListAvailable -Name ImportExcel)) {
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not install ImportExcel module. Defaulting to CSV output only."
            return $false
        }
    }

    # Import ImportExcel module
    try {
        Import-Module ImportExcel -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not import ImportExcel module. Defaulting to CSV output only."
        return $false
    }

    return $true
}

#Function to get Export Path
function Get-ExportPath {
    # Ask user for Export location
    Write-Host "Gather Export Path and/or File Name" -foregroundcolor Cyan
    $userInput = Read-Host -Prompt "Enter the file path (with .xlsx or .csv extension) or folder path to save the file"

    # If user input is empty, default to Desktop
    if ([string]::IsNullOrEmpty($userInput)) {
        $userInput = [Environment]::GetFolderPath("Desktop")
    }

    # File path processing
    $folderPath = ""
    $fileName = ""

    if ((Test-Path $userInput) -and (Get-Item -Path $userInput -ErrorAction SilentlyContinue).PSIsContainer) {
        $folderPath = $userInput
    } else {
        $folderPath = Split-Path -Path $userInput -Parent
        $fileName = Split-Path -Path $userInput -Leaf
    }

    # If folderPath is empty or invalid, default to current script location
    if ([string]::IsNullOrEmpty($folderPath) -or !(Test-Path $folderPath)) {
        $folderPath = $PSScriptRoot
    }

    # Check file extension and set default if none
    $extension = [IO.Path]::GetExtension($fileName)

    if ([string]::IsNullOrEmpty($extension)) {
        $extension = ".xlsx"
        $fileName = "$global:tenant-AllMailboxStats)" + $extension
    }

    # Full path
    $fullPath = Join-Path -Path $folderPath -ChildPath $fileName
    Write-Host ""
    return $fullPath, $extension
    
}

# Function to export data to file
function Export-DataToPath($fullPath, $extension, $data) {
    # If Excel export is not possible, default to CSV
    if (!$excelModuleInstalled -and $extension -eq ".xlsx") {
        $fullPath = [IO.Path]::ChangeExtension($fullPath, ".csv")
        $extension = ".csv"
    }

    # Check if folder path is valid
    $folderPath = Split-Path -Path $fullPath -Parent
    if ((Test-Path $folderPath) -and (Get-Item -Path $folderPath -ErrorAction SilentlyContinue).PSIsContainer) {
        if ($extension -eq ".csv") {
            # Export to CSV
            $data | Export-Csv -Path $fullPath -NoTypeInformation
        }
        elseif ($extension -eq ".xlsx") {
            # Export to Excel
            try {
                $data | Export-Excel -Path $fullPath -AutoSize
            }
            catch {
                # If Excel export fails, export to CSV instead
                $fullPath = [IO.Path]::ChangeExtension($fullPath, ".csv")
                $data | Export-Csv -Path $fullPath -NoTypeInformation
            }
        }
        else {
            Write-Host "Invalid file format. Only .csv or .xlsx is supported"
            return
        }
        Write-Host "Total number of report results: $($ExportAllMailboxStatsArray.count) | " -NoNewline
        Write-Host "Full Report exported to $fullPath" -ForegroundColor Cyan
    } else {
        Write-Host "Invalid folder path"
    }
}

#Connect to all required O365 services for running this script
Connect-Office365RequiredServices

#Tenant Name for Export
$global:tenant = "OVG"

#Get Export Path
$excelModuleInstalled = Install-ImportExcel
$ExportDetails = Get-ExportPath

#used to scope number of mailboxes pulled for TESTING.  Set to 'Unlimited' for a full run or 5 for limited run
$resultSize = 'unlimited'

#Global Start Time for Script
$global:initialStart = Get-Date
    ###ProgressBar
    $progresscounter = 1
    [nullable[double]]$global:secondsRemaining = $null
    $ProgressPreference = "Continue"

#Hash Table to hold final report data
$allMailboxStats = [Ordered]@{}
$allPublicFolderStats = @()
$tenantStatsHash = [Ordered]@{}

#Array to store all errors encountered
$allErrors = @()

###########################################################################################################################################

#Gather all Mailboxes, Group Mailboxes, Unified Groups, and Public Folders

###########################################################################################################################################

function Get-AllExchangeMailboxDetails {
    # Gather Mailboxes - Include InActive Mailboxes
    Write-Host "Gathering Exchange Online Objects and data" -ForegroundColor Black -BackgroundColor Yellow
    Write-Progress -Activity "Getting all mailboxes" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $start = Get-Date
    Write-Host "Getting all mailboxes and inactive mailboxes..." -ForegroundColor Cyan -nonewline
    #all Mailboxes (EXO Command)
    $allMailboxes = Get-EXOMailbox -ResultSize $resultSize -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -IncludeInactiveMailbox -PropertySets All -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    #all Group Mailboxes added to allMailboxes variable
    $allMailboxes += Get-Mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green


    #Add User Mailboxes to Hash
    Write-Progress -Activity "Adding User Mailbox Data to Hash" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $userMailboxHash = @{}
    $allUserMailboxes = $allMailboxes | ?{$_.RecipientTypeDetails -eq "UserMailbox"}
    foreach ($user in $allUserMailboxes) {
        $key = $user.ExchangeGuid.ToString()
        $value = $user
        $userMailboxHash[$key] = $value
    }

    #Add User Mailboxes to Hash
    Write-Progress -Activity "Adding User Mailbox Data to Hash" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $inactiveMailboxHash = @{}
    $allinActiveMailboxes = $allMailboxes | ?{$_.IsInactiveMailbox -eq $true}
    foreach ($inactiveMBX in $allinActiveMailboxes) {
        $key = $inactiveMBX.ExchangeGuid.ToString()
        $value = $inactiveMBX
        $inactiveMailboxHash[$key] = $value
    }

    #Add User Mailboxes to Hash
    Write-Progress -Activity "Adding Inactive User Data to Hash" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $nonuserMailboxHash = @{}
    $allNonUserMailboxes = $allMailboxes | ?{$_.RecipientTypeDetails -ne "UserMailbox" -or $_.RecipientTypeDetails -ne "GroupMailbox"}
    foreach ($nonUser in $allNonUserMailboxes) {
        $key = $nonUser.ExchangeGuid.ToString()
        $value = $nonUser
        $nonuserMailboxHash[$key] = $value
    }

    #Add Group Mailboxes to Hash
    Write-Progress -Activity "Adding Group Mailbox Data to Hash" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $groupMailboxHash = @{}
    $allGroupMailboxes = $allMailboxes | ?{$_.RecipientTypeDetails -eq "GroupMailbox"}
    foreach ($groupMailbox in $allGroupMailboxes) {
        $key = $groupMailbox.ExchangeGuid.ToString()
        $value = $groupMailbox
        $groupMailboxHash[$key] = $value
    }

    #Mailbox Statistics to Hash Table
    ###########################################################################################################################################
    ## Primary Mailbox Stats
    $start = Get-Date
    $mailboxStatsHash = @{}

    #Write-Host ""
    Write-Host "Getting primary mailbox stats..." -ForegroundColor Cyan -nonewline
    Write-Progress -Activity "Adding All primary mailbox (including Groups) stats" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    
    $start = Get-Date
    $primaryMailboxStats = get-exomailbox -ResultSize $resultSize -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors | Get-EXOMailboxStatistics -Properties LastLogonTime -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
    Write-Host "Getting group primary mailbox stats..." -ForegroundColor Cyan -nonewline
    $primaryMailboxStats += get-mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction Continue -ErrorVariable +allErrors | Get-MailboxStatistics -ErrorAction Continue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

    #process $primaryMailboxStats to add data to hash table with Key as MailboxGuid and Value as Array of Statistics Data
    Write-Host "Adding all mailbox stats to hash table..." -ForegroundColor Green -nonewline
    $primaryMailboxStats | ForEach-Object {
        $key = $_.MailboxGuid.ToString()
        $value = $_
        $mailboxStatsHash[$key] = $value
    }

    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

    ## Archive Mailbox Stats to Hash Table
    $start = Get-Date
    $archiveMailboxStatsHash = @{}

    Write-Host "Getting archive mailbox stats..." -ForegroundColor Cyan -nonewline
    Write-Progress -Activity "Getting All archive mailbox stats" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $archiveMailboxStats = get-exomailbox -ResultSize $resultSize -Properties ArchiveStatus -IncludeInactiveMailbox | Where-Object {$_.ArchiveStatus -ne "None"}| Get-EXOMailboxStatistics -Archive -Properties LastLogonTime -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
    Write-Host "Getting group archive mailbox archive stats..." -ForegroundColor Cyan -nonewline
    $archiveMailboxStats += get-mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox | Where-Object {$_.ArchiveStatus -ne "None"}| Get-MailboxStatistics -Archive -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients

    Write-Host "Adding to hash table..." -ForegroundColor Green -nonewline
    #process $archiveMailboxStats to add data to hash table with Key as MailboxGuid and Value as Array of Statistics Data
    $archiveMailboxStats | ForEach-Object {
        #errors if key is null
        if($key = $_.MailboxGuid.ToString()) {
            $archiveMailboxStatsHash[$key] = $value
            $value = $_
        }
    }
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

    $tenantStatsHash["UserMailboxes"] = $userMailboxHash
    $tenantStatsHash["GroupMailboxes"] = $groupMailboxHash
    $tenantStatsHash["NonUserMailboxes"] = $nonuserMailboxHash
    $tenantStatsHash["InActiveMailboxes"] = $inactiveMailboxHash
    $tenantStatsHash["AllMailboxes"] = $userMailboxHash + $groupMailboxHash + $nonuserMailboxHash + $inactiveMailboxHash
    $tenantStatsHash["PrimaryMailboxStats"] = $mailboxStatsHash
    $tenantStatsHash["ArchiveMailboxStats"] = $archiveMailboxStatsHash
}

<# Not Needed
function Get-AllExchangeGroupMailboxes {
    Write-Progress -Activity "Getting all group mailboxes" -Completed
    # Gather Group Mailboxes
    $start = Get-Date
    Write-Host "Getting all group mailboxes..." -ForegroundColor Cyan -nonewline
    Write-Progress -Activity "Getting all group mailboxes" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $allGroupMailboxes = Get-Mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
    return  $allGroupMailboxes
}
#>

function Get-AllUnifiedGroups {
    #Get Office 365 Group / Group Mailbox data with SharePoint URL data
    $start = Get-Date
    Write-Host "Getting all unified groups (including soft deleted)..." -ForegroundColor Cyan -nonewline
    Write-Progress -Activity "Getting unified groups" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $allUnifiedGroups = Get-UnifiedGroup -resultSize $resultSize -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

    Write-Progress -Activity "Adding Unified Group data to Hash" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $unifiedGroupHash = @{}
    foreach ($group in $allUnifiedGroups) {
        $key = $group.ExchangeGuid.ToString()
        $value = $group
        $unifiedGroupHash[$key] = $value
    }

    $tenantStatsHash["all_UnifiedGroups"] = $unifiedGroupHash
}

Get-AllExchangeMailboxDetails
Get-AllUnifiedGroups

###########################################################################################################################################

#Public Folder Data; Statistics; Permissions Convert to Hash Tables

###########################################################################################################################################

function Get-AllPublicFolderDetails {
    #Get Public Folder Data, Statistics, and Permissions
    #Write-Progress -Activity "Getting public folders, Stats and Perms" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $start = Get-Date
    Write-Host "Getting public folders, Stats and Perms ..." -ForegroundColor Cyan -nonewline
    Write-Progress -Activity "Getting all public folder details" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))
    $allPublicFolders = get-publicfolder -recurse -resultSize $resultSize -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    $PublicFoldersHash = @{}

    foreach($pf in $allPublicFolders) {
        $key = $pf.EntryId
        $value = $pf
        $PublicFoldersHash[$key] = $value
    }
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
    

    # Public Folder Statistics
    #**************************
    $PublicFolderStatistics = $allPublicFolders | get-publicfolderstatistics -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    $PublicFolderStatsHash = @{}
    foreach($publicFolderStat in $PublicFolderStatistics) {

        $key = $PublicFolderStat.EntryId
        $value = $PublicFolderStat

        $PublicFolderStatsHash[$key] = $value
    }

    # Public Folder Permissions
    #***************************
    $PublicFolderPermissions = $allPublicFolders | get-publicfolderclientpermission -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    #Progress Bar Parameters Reset
    $start = Get-Date
    $progresscounter = 1
    $totalCount = $PublicFolderPermissions.count

    $PublicFolderPermissionsHash = @{}
    Write-Host "Processing Public Folder Permissions..." -ForegroundColor Cyan -nonewline
    foreach($publicFolderPermission in $PublicFolderPermissions) {
        Write-ProgressHelper -Activity "Gathering Public Folder Permissions for $($publicFolderPermission.Identity)" -ProgressCounter ($progresscounter++) -TotalCount $totalCount -StartTime $start
        
        $key = $publicFolderPermission.Identity

        if($PublicFolderPermissionsHash.ContainsKey($key)) {   
            $user = $publicFolderPermission.User
            $primarySMTP = $publicFolderPermission.User.RecipientPrincipal.PrimarySmtpAddress
            $AccessRights = $publicFolderPermission.AccessRights

            $currentValue = $PublicFolderPermissionsHash[$key]

            if($primarySMTP) {
                $value = $currentValue += ",($($User) - $($primarySMTP) - $($AccessRights))"
            }
            else {
                $value = $currentValue += ",($($User) - $($AccessRights))"
            }

            $PublicFolderPermissionsHash[$key] = $value
        }

        else {
            $user = $publicFolderPermission.User
            $primarySMTP = $publicFolderPermission.User.RecipientPrincipal.PrimarySmtpAddress
            $AccessRights = $publicFolderPermission.AccessRights

            if($primarySMTP) {
                $value = "($($User) - $($primarySMTP) - $($AccessRights))"
            }
            else {
                $value = "($($User) - $($AccessRights))"
            }

            $PublicFolderPermissionsHash[$key] = $value
        }
    }
    Write-Progress -Activity "Gathering Public Folder Permissions for $($publicFolderPermission.Identity)" -Completed
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
    
    #Output
    $tenantStatsHash["all_PublicFolderDetails"] = $PublicFoldersHash
    $tenantStatsHash["all_PublicFolderStatistics"] = $PublicFolderStatsHash
    $tenantStatsHash["all_PublicFolderPermissions"] = $PublicFolderPermissionsHash

}

Get-AllPublicFolderDetails


###########################################################################################################################################

# License SKUs and Service Plan IDs to HASH

###########################################################################################################################################

function Get-AllLicenseSKUs {
    Write-Progress -Activity "Adding License SKUs and Service Plan IDs to Hash" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))

    # Get License SKUs
    $skus = Get-MgSubscribedSku -ErrorAction Continue -ErrorVariable +allErrors

    #Build a hashtable for looking up license names from license sku
    $licenseSkuHash = @{}

    #store Service Plan IDs and corresponding ServicePlan data for each ID
    $servicePlanHash = @{}

    #Add License SKUs to Hash Table and Service Plans under each SKU to another Hash Table
    foreach ($sku in $skus) {
        $licenseSkuHash[$sku.SkuId] = $sku

        foreach($servicePlan in $sku.ServicePlans) {
            $key = $servicePlan.ServicePlanId

            #If plan id has NOT already been added to hash
            if(!($servicePlanHash.ContainsKey($key))) {
                
                $value = $servicePlan.ServicePlanName
                $servicePlanHash[$key] = $value
            }

            else {
                #Move to next object
                Continue
            }
        }
    }
    $tenantStatsHash["all_LicenseSKUs"] = $licenseSkuHash
    $tenantStatsHash["all_servicePlans"] = $servicePlanHash
}
Get-AllLicenseSKUs



###########################################################################################################################################

#Get all OneDrive Personal Sites data and add to hash table with UserPrincipalName as HASH KEY and OneDrive data as HASH VALUE

###########################################################################################################################################
function Get-AllOneDriveDetails {
    $start = Get-Date

    $OneDriveDataHash = @{}

    Write-Host "Getting all OneDrive site data..." -ForegroundColor Cyan -nonewline
    $OneDriveSite = Get-SPOSite -Filter "URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop -ErrorVariable +allErrors

    #OneDrive data to OneDrive Hash
    #************************************************************************************
    foreach ($site in $OneDriveSite) {
        #Add to hash table
        $key = $site.Owner
        $value = $site

        if($OneDriveDataHash.ContainsKey($key)) {
            Write-Host "Owner $($key) already in hash with current value" $OneDriveDataHash[$key].URL -ForegroundColor Yellow
            Write-Host "New url found: $($value.URL)" -ForegroundColor DarkMagenta
            #if key already exists in hash table, get current value and add new data
            $currentValue = $OneDriveDataHash[$key]
            $newValue = $currentValue
            $newValue.URL += ", $($value.URL)"
            $OneDriveDataHash.Remove($key)
            $OneDriveDataHash.add($key, $newValue) 
        }
        else {
            #if key doesn't exist in hash table, add URL discovered for this key
            $OneDriveDataHash[$key] = $value
        }
    }
    
    $tenantStatsHash["all_OneDrives"] = $OneDriveDataHash
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
}
function Get-AllSharePointSiteDetails {
    $start = Get-Date
    $SharePointSiteHash = @{}

    #Get all SharePoint sites for associating with Office 365 Groups / GroupMailboxes
    Write-Host "Getting all SharePoint site data..." -ForegroundColor Cyan -nonewline
    $SharePointSite = Get-SPOSite -IncludePersonalSite $false -Limit all -ErrorAction Stop -ErrorVariable +allErrors

    #SharePoint data to SharePoint Hash
    #************************************************************************************
    foreach ($site in $SharePointSite) {
        #Add to hash table
        $key = $site.URL
        $value = $site
        $SharePointSiteHash[$key] = $value
    }
    $tenantStatsHash["all_SharePointSites"] = $SharePointSiteHash

    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
}

Get-AllOneDriveDetails
Get-AllSharePointSiteDetails

###########################################################################################################################################

#Microsoft Graph (mg) User Data to Hash Table

###########################################################################################################################################
function Get-allMGUserDetails {
    $start = Get-Date
    $mgUserHash = @{}

    try {
        #Write-Host ""
        Write-Host "Getting all Microsoft Graph User data..." -ForegroundColor Cyan -nonewline
        Write-Progress -Activity "Getting all Microsoft Graph User Data" -Status (((Get-Date) –$global:initialStart).ToString('hh\:mm\:ss'))

        #MS GRAPH update
        $mgUsers = Get-MGUser -All -ErrorAction Stop -ErrorVariable +allErrors
    }
    catch {
        Write-Host "Error getting all Microsoft Graph sites data." -ForegroundColor Red
    }

    foreach ($user in $mgUsers) {
        <#add sign in activity - takes too long to run
        $signInactivity = (Get-MGUSer -UserId $user.ID -Property SignInActivity).SignInActivity
        $user | Add-Member -MemberType NoteProperty -Name "LastSignInDateTime" -Value $signInactivity.LastSignInDateTime -Force
        $user | Add-Member -MemberType NoteProperty -Name "LastSignInRequestId" -Value $signInactivity.LastSignInRequestId -Force
        $user | Add-Member -MemberType NoteProperty -Name "LastNonInteractiveSignInDateTime" -Value $signInactivity.LastNonInteractiveSignInDateTime -Force
        $user | Add-Member -MemberType NoteProperty -Name "LastNonInteractiveSignInRequestId" -Value $signInactivity.LastNonInteractiveSignInRequestId -Force
        #>
        #Add to hash table
        $key = $user.UserPrincipalName
        $value = $user
        $mgUserHash[$key] = $value
    }

    $tenantStatsHash["all_MG-Users"] = $mgUserHash
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
    
}

Get-allMGUserDetails


###########################################################################################################################################

#Consolidate Reporting for each user

###########################################################################################################################################
Write-Host ""
Write-Host "Consolidating report data for each user / object..." -ForegroundColor Black -BackgroundColor Yellow
Write-Host ""

#ProgressBar
$start = Get-Date
$progresscounter = 1
[nullable[double]]$global:secondsRemaining = $null
$TotalCount =  $tenantStatsHash["UserMailboxes"].keys.count
$ProgressPreference = "Continue"

Write-Host "Processing all mailboxes (non-group)..." -ForegroundColor Cyan -nonewline

#Process All Mailboxes - excluding group mailboxes which are processed in a seperate foreach
#*********************************************************************************************

foreach ($mailbox in $allMailboxes) {
    #progress bar
    Write-ProgressHelper -Activity "Gathering Mailbox Details for $($mailbox.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -StartTime $start

    #Pull MailboxStats and UserDetails
    #*******************************************************************************************************************
    $EmailAddresses = $mailbox | Select-Object -ExpandProperty EmailAddresses

    if($mailbox.ExchangeGuid) {
        $mbxStats = $mailboxStatsHash[$mailbox.ExchangeGuid]
    }
    if($mailbox.UserPrincipalName -and $mailbox.RecipientTypeDetails -ne "GroupMailbox") {
        $mgUser = $mgUserHash[$mailbox.UserPrincipalName]
        $oneDriveData = $OneDriveDataHash[$mailbox.UserPrincipalName]
    }

    #get mailbox size in GB if TotalItemSize exists - null values break the hash table creation 
    if($mbxStats.TotalItemSize) {
        $MBXSizeGB = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
    }
    else {
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
    foreach($license in $licenses) {
        $licenseName = $licenseSkuHash[$license.SkuId].SkuPartNumber
        $allLicenses += $licenseName

        if($license.DisabledPlans) {
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
        #User Information
        "DisplayName" = $mgUser.DisplayName
        "UserPrincipalName" = $mgUser.userprincipalname
        "Department" = $mgUser.Department
        "IsLicensed" = ($mgUser.AssignedLicenses.count -gt 0)
        "Licenses" = ($allLicenses -join ",")
        "License-DisabledArray" = ($allDisabledPlans -join ",")
        "AccountEnabled" = $mgUser.AccountEnabled
        "IsInactiveMailbox" = $mailbox.IsInactiveMailbox
        "WhenSoftDeleted" = $mailbox.WhenSoftDeleted
        <#Login Activity
        "LastSignInDateTime" = $mgUser.LastSignInDateTime
        "LastSignInRequestId" = $mgUser.LastSignInRequestId
        "LastNonInteractiveSignInDateTime" = $mgUser.LastNonInteractiveSignInDateTime
        "LastNonInteractiveSignInRequestId" = $mgUser.LastNonInteractiveSignInRequestId
        #>
        "WhenCreated" = $mailbox.WhenCreated
        "LastLogonTime" = $mbxStats.LastLogonTime
        #mailbox information
        "RecipientTypeDetails" = $mailbox.RecipientTypeDetails
        "PrimarySmtpAddress" = $mailbox.PrimarySmtpAddress
        "HiddenFromAddressListsEnabled" = $mailbox.HiddenFromAddressListsEnabled
        "MBXSize" = $MBXStats.TotalItemSize
        "MBXSize_GB" = $MBXSizeGB
        "MBXItemCount" = $MBXStats.ItemCount
        "Alias" = $mailbox.alias
        "EmailAddresses" = ($EmailAddresses -join ";")
        "DeliverToMailboxAndForward" = $mailbox.DeliverToMailboxAndForward
        "ForwardingAddress" = $mailbox.ForwardingAddress
        "ForwardingSmtpAddress" = $mailbox.ForwardingSmtpAddress
        "LitigationHoldEnabled" = $mailbox.LitigationHoldEnabled
        "LitigationHoldDuration" = $mailbox.LitigationHoldDuration
        "InPlaceHolds" = $mailbox.InPlaceHolds -join ";"
        "RetentionPolicy" = $mailbox.RetentionPolicy
    
        <#Public Folder Fields
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
        #>
    }


    #Pull Send on Behalf
    #*******************************************************************************************************************
    $grantSendOnBehalf = $mailbox.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        if ($mailboxCheck = (Get-EXORecipient $perm -IncludeSoftDeletedRecipients -ErrorAction SilentlyContinue -ErrorVariable +allErrors).DisplayName) {
            $grantSendOnBehalfPerms += $mailboxCheck
        }
        else {
            $grantSendOnBehalfPerms += $perm
        }
    }
    $currentuser["GrantSendOnBehalfTo"] = ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    #*******************************************************************************************************************

    #Exclude Inactive Mailbox and group mailboxes to avoid errors
    if($mailbox.IsInactiveMailbox -eq $false -and $mailbox.UserPrincipalName -and $mailbox.RecipientTypeDetails -ne "GroupMailbox") {
        if ($mbxPermissions = Get-MailboxPermission $mailbox.DistinguishedName -ErrorVariable +allErrors -ErrorAction SilentlyContinue | Where-Object{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
            $currentuser["FullAccessPerms"] = ($mbxPermissions.User -join ";")
        }
        else {$currentuser["FullAccessPerms"] = ($null)}
        # Mailbox Send As Check
        if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $mailbox.DistinguishedName -ErrorVariable +allErrors -ErrorAction SilentlyContinue  | Where-Object{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
            $currentuser["SendAsPerms"] = ($sendAsPermsCheck.trustee -join ";")
        }
        else {$currentuser["SendAsPerms"] = ($null)}
    }

    # Archive Mailbox Check
    #*******************************************************************************************************************
    $currentuser["ArchiveStatus"] = $mailbox.ArchiveStatus
    if ($mailbox.ArchiveStatus -ne "None" -and $null -ne $mailbox.ArchiveStatus) {
        $archiveStats = $archiveMailboxStatsHash[$mailbox.ArchiveGuid]
        $currentuser["ArchiveSize"] = $ArchiveStats.TotalItemSize.Value

        if($ArchiveStats.TotalItemSize) {
            $currentuser["ArchiveSize-GB"] = [math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
        }
        else {
            $currentuser["ArchiveSize-GB"] = $null 
        }

        $currentuser["ArchiveItemCount"] = $ArchiveStats.ItemCount
    }

    else {
        $currentuser["ArchiveSize"] = $null
        $currentuser["ArchiveSize-GB"] = $null
        $currentuser["ArchiveItemCount"] = $null
    }

    #Get OneDrive URL if this user is an owner of a OneDrive site
    #*******************************************************************************************************************


    if($OneDriveData) {
        $currentuser["OneDriveURL"] = $OneDriveData.URL
        $currentuser["OneDriveStorage-GB"] = [math]::Round($OneDriveData.StorageUsageCurrent / 1024, 3)
        $currentuser["OneDriveLastContentModifiedDate"] = $OneDriveData.LastContentModifiedDate
        $currentuser["SharePointURL"] = $null
        $currentuser["SharePointStorage-GB"] = $null
        $currentuser["SharePointLastContentModifiedDate"] = $null
    }
    else {

        $currentuser["OneDriveURL"] = $null
        $currentuser["OneDriveStorage-GB"] = $null
        $currentuser["OneDriveLastContentModifiedDate"] = $null
        $currentuser["SharePointURL"] = $null
        $currentuser["SharePointStorage-GB"] = $null
        $currentuser["SharePointLastContentModifiedDate"] = $null
    }

    #Combine all the data into one hash table
    #*******************************************************************************************************************

    $allMailboxStats[$mailbox.PrimarySMTPAddress] = $currentuser

}

Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
###########################################################################################################################################

# Add Group Mailbox Data to Reporting

###########################################################################################################################################

#Progress Bar Parameters Reset
$start = Get-Date
$progresscounter = 1
$totalCount = $allGroupMailboxes.count

Write-Host "Processing all group mailboxes..." -ForegroundColor Cyan -nonewline

if($allGroupMailboxes) {
    foreach ($mailbox in $allGroupMailboxes) {
        #progress bar
        Write-ProgressHelper -Activity "Gathering Group Mailbox Details for $($mailbox.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -StartTime $start
            
        #Pull MailboxStats and UserDetails
        #*******************************************************************************************************************

        if($mailbox.ExchangeGuid) {
            $mbxStats = $mailboxStatsHash[$mailbox.ExchangeGuid]
        }

        #If $mailbox represents a GroupMailbox object set $mgUser data to pull from GroupMailboxHash which contains group mailbox data
        if($mailbox.RecipientTypeDetails -eq "GroupMailbox") {
            $mgUser = $groupMailboxHash[$mailbox.ExchangeGuid]
            $unifiedGroupData = $unifiedGroupHash[$mailbox.ExchangeGuid]

            if($unifiedGroupData.SharePointSiteUrl) {
                $sharePointSiteData = $SharePointSiteHash[($unifiedGroupData.SharePointSiteUrl)]
            }
        }
        $EmailAddresses = $mailbox | Select-Object -ExpandProperty EmailAddresses

        #get mailbox size in GB if TotalItemSize exists - null values break the hash table creation
        
        if($mbxStats.TotalItemSize) {
            $MBXSizeGB = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
        } else {
            $MBXSizeGB = 0
        }

        # Create Hash Table to add to Report Dataset
        #*******************************************************************************************************************
            $currentuser = [ordered]@{
                #User Information
                "DisplayName" = $mgUser.DisplayName
                "UserPrincipalName" = $mgUser.userprincipalname
                "Department" = $mgUser.Department
                "IsLicensed" = ($mgUser.AssignedLicenses.count -gt 0)
                "Licenses" = ($allLicenses -join ",")
                "License-DisabledArray" = ($allDisabledPlans -join ",")
                "AccountEnabled" = $mgUser.AccountEnabled
                "IsInactiveMailbox" = $mailbox.IsInactiveMailbox
                "WhenSoftDeleted" = $mailbox.WhenSoftDeleted
                <#Login Activity
                "LastSignInDateTime" = $mgUser.LastSignInDateTime
                "LastSignInRequestId" = $mgUser.LastSignInRequestId
                "LastNonInteractiveSignInDateTime" = $mgUser.LastNonInteractiveSignInDateTime
                "LastNonInteractiveSignInRequestId" = $mgUser.LastNonInteractiveSignInRequestId
                #>
                "WhenCreated" = $mailbox.WhenCreated
                "LastLogonTime" = $mbxStats.LastLogonTime
                #mailbox information
                "RecipientTypeDetails" = $mailbox.RecipientTypeDetails
                "PrimarySmtpAddress" = $mailbox.PrimarySmtpAddress
                "HiddenFromAddressListsEnabled" = $mailbox.HiddenFromAddressListsEnabled
                "MBXSize" = $MBXStats.TotalItemSize
                "MBXSize_GB" = $MBXSizeGB
                "MBXItemCount" = $MBXStats.ItemCount
                "Alias" = $mailbox.alias
                "EmailAddresses" = ($EmailAddresses -join ";")
                "DeliverToMailboxAndForward" = $mailbox.DeliverToMailboxAndForward
                "ForwardingAddress" = $mailbox.ForwardingAddress
                "ForwardingSmtpAddress" = $mailbox.ForwardingSmtpAddress
                "LitigationHoldEnabled" = $mailbox.LitigationHoldEnabled
                "LitigationHoldDuration" = $mailbox.LitigationHoldDuration
                "InPlaceHolds" = $mailbox.InPlaceHolds -join ";"
                "ArchiveStatus" = $mailbox.ArchiveStatus
                "RetentionPolicy" = $mailbox.RetentionPolicy
            
                <#Public Folder Fields
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
                #>
            }
        
        #Pull Send on Behalf
        #*******************************************************************************************************************
        $grantSendOnBehalf = $mailbox.GrantSendOnBehalfTo
        $grantSendOnBehalfPerms = @()
        foreach ($perm in $grantSendOnBehalf) {
            if ($mailboxCheck = (Get-EXORecipient $perm -IncludeSoftDeletedRecipients -ErrorAction SilentlyContinue -ErrorVariable +allErrors).DisplayName) {
                $grantSendOnBehalfPerms += $mailboxCheck
            }
            else {
                $grantSendOnBehalfPerms += $perm
            }
        }
        $currentuser["GrantSendOnBehalfTo"] = ($grantSendOnBehalfPerms -join ";")

        # Mailbox Full Access Check
        #*******************************************************************************************************************

        #Exclude Group Mailbox to avoid errors
        if($mailbox.RecipientTypeDetails -eq "GroupMailbox") {
            $currentuser["FullAccessPerms"] = ($null)
            $currentuser["SendAsPerms"] = ($null)
        }
        else {
            if ($mbxPermissions = Get-MailboxPermission $mailbox.DistinguishedName -ErrorVariable +allErrors -ErrorAction SilentlyContinue | Where-Object{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
                $currentuser["FullAccessPerms"] = ($mbxPermissions.User -join ";")
            }
            else {$currentuser["FullAccessPerms"] = ($null)}
            # Mailbox Send As Check
            if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $mailbox.DistinguishedName -ErrorVariable +allErrors -ErrorAction SilentlyContinue  | Where-Object{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
                $currentuser["SendAsPerms"] = ($sendAsPermsCheck.trustee -join ";")
            }
            else {$currentuser["SendAsPerms"] = ($null)}
        }

        # Archive Mailbox Check
        #*******************************************************************************************************************
        if ($mailbox.ArchiveStatus -ne "None" -and $null -ne $mailbox.ArchiveStatus) {
            $archiveStats = $archiveMailboxStatsHash[$mailbox.ArchiveGuid]

            $currentuser["ArchiveSize"] = $ArchiveStats.TotalItemSize.Value

            if($ArchiveStats.TotalItemSize) {
                $currentuser["ArchiveSize-GB"] = [math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
            }
            else {
                $currentuser["ArchiveSize-GB"] = $null 
            }

            $currentuser["ArchiveItemCount"] = $ArchiveStats.ItemCount
        }

        else {
            $currentuser["ArchiveSize"] = $null
            $currentuser["ArchiveSize-GB"] = $null
            $currentuser["ArchiveItemCount"] = $null
        }

        #Get SharePoint URL of Group - If user is not inactive
        #*******************************************************************************************************************
        #Errors in if statement if userprincipalname is null

        #Group Mailbox Associated SharePoint Site mapping
        if($sharePointSiteData) {    
            $currentuser["OneDriveURL"] = $null
            $currentuser["OneDriveStorage-GB"] = $null
            $currentuser["OneDriveLastContentModifiedDate"] = $null 
            $currentuser["SharePointURL"] = $sharePointSiteData.URL
            $currentuser["SharePointStorage-GB"] = [math]::Round($sharePointSiteData.StorageUsageCurrent / 1024, 3)
            $currentuser["SharePointLastContentModifiedDate"] = $sharePointSiteData.LastContentModifiedDate
        }

        else {
            $currentuser["OneDriveURL"] = $null
            $currentuser["OneDriveStorage-GB"] = $null
            $currentuser["OneDriveLastContentModifiedDate"] = $null 
            $currentuser["SharePointURL"] = $null
            $currentuser["SharePointStorage-GB"] = $null
            $currentuser["SharePointLastContentModifiedDate"] = $null
        }

        #Combine all the data into one hash table
        #*******************************************************************************************************************
        $allMailboxStats[$mailbox.PrimarySMTPAddress] = $currentuser

    }
    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
}
else {
    Write-Host "No Group Mailboxes found" -ForegroundColor Yellow
}


###########################################################################################################################################

#Add Public Folder Data to Reporting

###########################################################################################################################################

#Seperate for each for Public Folder data to avoid unnecessary error handling

#reset progress counter for next operation
$start = Get-Date
$progresscounter = 1
$totalCount = $PublicFolders.Count

Write-Host "Processing all public folders..." -ForegroundColor Cyan -nonewline
if ($PublicFolders.count -gt 2) {
    foreach($pf in $PublicFolders) {
        Write-ProgressHelper -Activity "Gathering Public Folder Details for $($pf.Name)" -ProgressCounter ($progresscounter++) -TotalCount $totalCount -StartTime $start
    
        $currentFolder = @{
                "DisplayName" = $pf.Name
                "RecipientTypeDetails" = "PublicFolder"
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
        $allPublicFolderStats += $currentFolder
    }

    $ExportAllPublicFoldersStatsArray = @()
    foreach ($key in $PublicFolders.Keys) {
        $attributes = $PublicFolders[$key]
        $customObject = New-Object -TypeName PSObject

        foreach ($attribute in $attributes.Keys) {
            $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)__$($tenant)" -Value $attributes[$attribute]
        }

        $ExportAllPublicFoldersStatsArray += $customObject
    }

    $ExportAllPublicFoldersStatsArray | Export-Excel -Path $ExportDetails[0] -WorksheetName "PublicFolderStats"

    Write-Host "Completed in "(((Get-Date) – $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
}
else {
    Write-Host "No Public Folders found" -ForegroundColor Yellow
}

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

# Output Location
Export-DataToPath -fullPath $ExportDetails[0] -extension $ExportDetails[1] -data $ExportAllMailboxStatsArray


#Filter out blank lines added to errors
$allErrors = $allErrors | Where-Object {$_.ErrorRecord}

if($allErrors) {
    $allErrors | Export-Csv $pwd\$($tenant)-errorReport.csv -NoTypeInformation
}

Write-Host "Number Errors: $($allErrors.count) | " -NoNewline -ForegroundColor Red
Write-Host "Errors exported to $pwd\$($tenant)-errorReport.csv"
Write-Host ""
Write-Host "Completed in"((Get-Date) -$global:initialStart).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan
Write-Host ""

Write-Host "Recipient Count Table" -ForegroundColor Black -BackgroundColor Green

#Final Output of Recipient Counts
$tenantStatsHash["AllMailboxes"] = $allMailboxes
$tenantStatsHash["allGroupMailboxes"] = $allGroupMailboxes
$tenantStatsHash["allUnifiedGroups"] = $allUnifiedGroups
$tenantStatsHash["PublicFolders"] = $PublicFolders
$tenantStatsHash["PublicFolderPermissions"] = $PublicFolderPermissions
$tenantStatsHash["licenseSkus"] = $licenseSkuHash
$tenantStatsHash["servicePlanIDs"] = $servicePlanHash
$tenantStatsHash["OneDriveSites"] = $OneDriveDataHash
$tenantStatsHash["SharePointSites"] = $SharePointSiteHash
$tenantStatsHash["mgUsers"] = $mgUserHash
$tenantStatsHash["mailboxStats"] = $mailboxStatsHash
$tenantStatsHash["archiveMailboxStats"] = $archiveMailboxStatsHash

$TenantStatsOutput = @()

foreach ($key in $tenantStatsHash.Keys) {
    $count = $tenantStatsHash[$key].Count
    # Create a custom object for the current key-value pair
    $object = New-Object -TypeName PSCustomObject -Property @{
        "Key" = $key
        "Count" = $count
    }
    # Add the custom object to the array
    $TenantStatsOutput += $object
}
$TenantStatsOutput