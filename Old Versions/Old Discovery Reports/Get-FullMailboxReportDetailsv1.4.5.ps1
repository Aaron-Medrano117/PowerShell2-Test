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

    v1.4.5
    Author: Aaron Medrano
    Updated portions of scripts into functions
    Use TenantStatsHash for reference in script


    v1.4.6
    Author: Aaron Medrano
    combined portions of scripts into functions
    updated export to export TenantStatsHash to CSV and combine into one Excel
    Added Domains and Admins to TenantStatsHash
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

    $secondsElapsed = (Get-Date) - $global:Initialstart
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

$AllDiscoveryErrors = @()
function Handle-ErrorHelper {
    param(
        [Parameter(Mandatory=$true)]
        [string]$errorVariable,
        [Parameter(Mandatory=$true)]
        [string]$errorMessage
    )

    if($errorVariable) {
        Write-Host $errorMessage -ForegroundColor Red
        foreach ($errorCheck in $errorVariable) {
            if ($errorCheck.Exception.Message -match "'[^']*/[^']*'") {
                $recipient = $matches[0].Trim("'")
            } else {
                $recipient = $null
            }
            
            $CurrentError = New-Object PSObject
            $CurrentError | Add-Member -Type NoteProperty -Name "ErrorMessage" -Value $errorMessage -Force
            $CurrentError | Add-Member -Type NoteProperty -Name "Commandlet" -Value $errorCheck.CategoryInfo.Activity
            $CurrentError | Add-Member -Type NoteProperty -Name "Reason" -Value $errorCheck.CategoryInfo.Reason -Force
            $CurrentError | Add-Member -Type NoteProperty -Name "Recipient" -Value $recipient
            $CurrentError | Add-Member -Type NoteProperty -Name "TargetObject" -Value $errorCheck.TargetObject -Force
            $AllDiscoveryErrors += $CurrentError
        }
    }
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
            $MSOCompanyCheck = Get-MsolCompanyInformation -ErrorAction Stop
            Write-Host "Completed. Connected to MSOnline: $($MSOCompanyCheck.DisplayName)" -foregroundcolor Green
            $msOnlineConnected = $true
        }
        catch {
            Write-Error "Error connecting to MSOnline: $($_.Exception.Message)"
        }
    }

    ## Connect to MgGraph
    if (($MGraphModuleFound -eq $true) -or ($MGraphConnected -eq $false)) {
        Write-Host "Connecting to MgGraph..." -NoNewline
        if ($null -ne $MGraphCompanyCheck.DisplayName -and $MGraphCompanyCheck.DisplayName -eq $MSOCompanyCheck.DisplayName){
            #Select-MgProfile -Name "beta"
            Write-Host "Already Connected to MGGraph: $($MGraphCompanyCheck.DisplayName)" -foregroundcolor Green
            $MGraphConnected = $true
        }
        else {
            if ($MGraphConnected -eq $true) {
                Write-Host "Connected to Incorrect Tenant. Connected to MGGraph: $($MGraphCompanyCheck.DisplayName). Disconnecting." -foregroundcolor Yellow
                $result = Disconnect-MgGraph
            }
            else {
                Write-Host "Not Connected to MGGraph. Connecting..." -NoNewline -ForegroundColor Yellow
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
                    #Select-MgProfile -Name "beta"
                    $MGraphCompanyCheck = Get-MgOrganization -ErrorAction Stop
                    
                    Write-Host "Completed. Connected to MGGraph: $($MGraphCompanyCheck.DisplayName)" -foregroundcolor Green
                    $MGraphConnected = $true
                    $global:tenant = $($MGraphCompanyCheck.DisplayName)
                }
                catch {
                    Write-Error "Error connecting to MgGraph: $($_.Exception.Message)"
                }
            }
            
        }

                  
    }

    # Connect to ExchangeOnline
    if ($EXOModuleFound -eq $true) {
        Write-Host "Connecting to ExchangeOnline..." -NoNewline
        try {
            $result = Connect-ExchangeOnline -UserPrincipalName $userPrincipalName -ErrorAction Stop *> Out-Null
            $EXOOrgCheck = Get-OrganizationConfig -ErrorAction Stop
            Write-Host "Completed. Connected to Exchange Online: $($EXOOrgCheck.Name)" -foregroundcolor Green
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
            $rootSiteURL = Get-SPOSite -limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
            $rootURL = $rootSiteURL.url -replace '/sites.*', ''
            Write-Host "Completed. Connected to SharePoint Online: $($rootURL)" -foregroundcolor Green
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
                if ($format -eq "xlsx") {
                    $dataToExport.GetEnumerator() | Export-Excel -Path (Join-Path -Path $folderPath -ChildPath "$fileName.xlsx") -WorksheetName $fileName -AutoSize -ClearSheet
                }
                elseif ($format -eq "csv") {
                    $dataToExport.GetEnumerator() | ConvertTo-Csv -NoTypeInformation | Set-Content (Join-Path -Path $folderPath -ChildPath "$fileName.csv")
                }
                
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
        return
    }
}

#Connect to all required O365 services for running this script
Connect-Office365RequiredServices

#Tenant Name for Export
#$global:tenant = Read-Host -Prompt "What is the tenant name?"

#Get Export Path
$excelModuleInstalled = Install-ImportExcel
$ExportDetails = Get-ExportPath

#used to scope number of mailboxes pulled for TESTING.  Set to 'Unlimited' for a full run or 5 for limited run
$resultSize = 'unlimited'

#Global Start Time for Script

$global:InitialStart = Get-Date
    ###ProgressBar
    $progresscounter = 1
    [nullable[double]]$global:secondsRemaining = $null
    $ProgressPreference = "Continue"

#Hash Table to hold final report data
$tenantStatsHash = [Ordered]@{}

#Array to store all errors encountered
$allErrors = @()

###########################################################################################################################################

#Gather all Mailboxes, Group Mailboxes, Unified Groups, and Public Folders
function Get-AllExchangeMailboxDetails {
    # Gather Mailboxes - Include InActive Mailboxes
    Write-Host "Gathering Exchange Online Objects and data" -ForegroundColor Black -BackgroundColor Yellow
    Write-Progress -Activity "Getting all mailboxes" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $start = Get-Date
    Write-Host "Getting all mailboxes and inactive mailboxes..." -ForegroundColor Cyan -nonewline
    #all Mailboxes (EXO Command)
    $allMailboxes = Get-EXOMailbox -ResultSize $resultSize -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -IncludeInactiveMailbox -PropertySets All -ErrorAction SilentlyContinue -ErrorVariable +allErrors 
    #all Group Mailboxes added to allMailboxes variable
    $allMailboxes += Get-Mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    $allMailboxes = $allMailboxes | select DisplayName, Office, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress, WhenMailboxCreated, UsageLocation, IsInactiveMailbox, WasInactiveMailbox, WhenSoftDeleted, InPlaceHolds, AccountDisabled, IsDirSynced, HiddenFromAddressListsEnabled, Alias, EmailAddresses, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers, AcceptMessagesOnlyFromSendersOrMembers, RejectMessagesFrom, RejectMessagesFromDLMembers, RejectMessagesFromSendersOrMembers, RequireSenderAuthenticationEnabled, WindowsEmailAddress, DistinguishedName, Identity, WhenChanged, WhenCreated, ExchangeObjectId, Guid, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress, LitigationHoldEnabled, RetentionHoldEnabled, DelayHoldApplied, RetentionPolicy, ExchangeGuid,  IsResource, IsShared, ResourceType, RoomMailboxAccountEnabled,  WindowsLiveID, MicrosoftOnlineServicesID, EffectivePublicFolderMailbox, MailboxPlan, ArchiveStatus, ArchiveState, ArchiveName, ArchiveGuid, AutoExpandingArchiveEnabled, DisabledArchiveGuid, PersistedCapabilities, CustomAttribute*, ExtensionCustomAttribute*
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green


    #Add User Mailboxes to Hash
    Write-Progress -Activity "Adding User Mailbox Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $allMailboxHash = @{}
    foreach ($user in $allMailboxes) {
        $key = $user.ExchangeGuid.ToString()
        $value = $user
        $allMailboxHash[$key] = $value
    }

    #Add User Mailboxes to Hash
    Write-Progress -Activity "Adding User Mailbox Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $userMailboxHash = @{}
    $allUserMailboxes = $allMailboxes | ?{$_.RecipientTypeDetails -eq "UserMailbox"}
    foreach ($user in $allUserMailboxes) {
        $key = $user.ExchangeGuid.ToString()
        $value = $user
        $userMailboxHash[$key] = $value
    }

    #Add User Mailboxes to Hash
    Write-Progress -Activity "Adding User Mailbox Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $inactiveMailboxHash = @{}
    $allinActiveMailboxes = $allMailboxes | ?{$_.IsInactiveMailbox -eq $true}
    foreach ($inactiveMBX in $allinActiveMailboxes) {
        $key = $inactiveMBX.ExchangeGuid.ToString()
        $value = $inactiveMBX
        $inactiveMailboxHash[$key] = $value
    }

    #Add User Mailboxes to Hash
    Write-Progress -Activity "Adding Inactive User Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $nonuserMailboxHash = @{}
    $allNonUserMailboxes = $allMailboxes | ?{$_.RecipientTypeDetails -ne "UserMailbox" -or $_.RecipientTypeDetails -ne "GroupMailbox"}
    foreach ($nonUser in $allNonUserMailboxes) {
        $key = $nonUser.ExchangeGuid.ToString()
        $value = $nonUser
        $nonuserMailboxHash[$key] = $value
    }

    #Add Group Mailboxes to Hash
    Write-Progress -Activity "Adding Group Mailbox Data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
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
    Write-Progress -Activity "Adding All primary mailbox (including Groups) stats" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    
    $start = Get-Date
    $primaryMailboxStats = get-exomailbox -ResultSize $resultSize -IncludeInactiveMailbox -ErrorAction SilentlyContinue -ErrorVariable +allErrors | Get-EXOMailboxStatistics -Properties LastLogonTime -ErrorAction SilentlyContinue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
    Write-Host "Getting group primary mailbox stats..." -ForegroundColor Cyan -nonewline
    $primaryMailboxStats += get-mailbox -ResultSize $resultSize -GroupMailbox -IncludeInactiveMailbox -ErrorAction Continue -ErrorVariable +allErrors | Get-MailboxStatistics -ErrorAction Continue -ErrorVariable +allErrors -IncludeSoftDeletedRecipients
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

    #process $primaryMailboxStats to add data to hash table with Key as MailboxGuid and Value as Array of Statistics Data
    Write-Host "Adding all mailbox stats to hash table..." -ForegroundColor Green -nonewline
    $primaryMailboxStats | ForEach-Object {
        $key = $_.MailboxGuid.ToString()
        $value = $_
        $mailboxStatsHash[$key] = $value
    }

    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

    ## Archive Mailbox Stats to Hash Table
    $start = Get-Date
    $archiveMailboxStatsHash = @{}

    Write-Host "Getting archive mailbox stats..." -ForegroundColor Cyan -nonewline
    Write-Progress -Activity "Getting All archive mailbox stats" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
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
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

    $tenantStatsHash["AllMailboxes"] = $allMailboxHash
    $tenantStatsHash["UserMailboxes"] = $userMailboxHash
    $tenantStatsHash["GroupMailboxes"] = $groupMailboxHash
    $tenantStatsHash["NonUserMailboxes"] = $nonuserMailboxHash
    $tenantStatsHash["InActiveMailboxes"] = $inactiveMailboxHash
    $tenantStatsHash["PrimaryMailboxStats"] = $mailboxStatsHash
    $tenantStatsHash["ArchiveMailboxStats"] = $archiveMailboxStatsHash
}

function Get-AllUnifiedGroups {
    #Get Office 365 Group / Group Mailbox data with SharePoint URL data
    $start = Get-Date
    Write-Host "Getting all unified groups (including soft deleted)..." -ForegroundColor Cyan -nonewline
    Write-Progress -Activity "Getting unified groups" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $allUnifiedGroups = Get-UnifiedGroup -resultSize $resultSize -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue -ErrorVariable +allErrors
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green

    Write-Progress -Activity "Adding Unified Group data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $unifiedGroupHash = @{}
    foreach ($group in $allUnifiedGroups) {
        $key = $group.ExchangeGuid.ToString()
        $value = $group
        $unifiedGroupHash[$key] = $value
    }

    $tenantStatsHash["UnifiedGroups"] = $unifiedGroupHash
}

#Public Folder Data; Statistics; Permissions Convert to Hash Tables
function Get-AllPublicFolderDetails {
    #Get Public Folder Data, Statistics, and Permissions
    $start = Get-Date
    Write-Host "Getting public folders, Stats and Perms ..." -ForegroundColor Cyan -nonewline
    Write-Progress -Activity "Getting all public folder details" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    $allPublicFolders = get-publicfolder -recurse -resultSize $resultSize -ErrorAction SilentlyContinue -ErrorVariable +allErrors | Select Identity, Name, MailEnabled, MailRecipientGuid, ParentPath, ContentMailboxName, EntryId, FolderSize, HasSubfolders, FolderClass, FolderPath, ExtendedFolderFlags
    $PublicFoldersHash = @{}

    foreach($pf in $allPublicFolders) {
        $key = $pf.EntryId
        $value = $pf
        $PublicFoldersHash[$key] = $value
    }
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
    

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
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
    
    #Output
    $tenantStatsHash["PublicFolderDetails"] = $PublicFoldersHash
    $tenantStatsHash["PublicFolderStats"] = $PublicFolderStatsHash
    $tenantStatsHash["PublicFolderPerms"] = $PublicFolderPermissionsHash

}

###########################################################################################################################################

#Microsoft Graph (mg) Data to Hash Table
function Get-allMGUserDetails {
    $start = Get-Date
    $mgUserHash = @{}

    try {
        #Write-Host ""
        Write-Host "Getting all Microsoft Graph User data..." -ForegroundColor Cyan -nonewline
        Write-Progress -Activity "Getting all Microsoft Graph User Data" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))

        #MS GRAPH update
        $mgUsers = Get-MGUser -All -Property SignInActivity -ErrorAction Stop -ErrorVariable +allErrors | select DisplayName, UserPrincipalName, UserType, Id, AccountEnabled, CreatedDateTime, Mail, JobTitle, Department, CompanyName, OfficeLocation, City, State, Country, OnPremisesSyncEnabled, OnPremisesDistinguishedName, OnPremisesLastSyncDateTime, UsageLocation, SignInActivity
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

    $tenantStatsHash["MG-Users"] = $mgUserHash
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
    
}
# License SKUs and Service Plan IDs to HASH
function Get-AllLicenseSKUs {
    Write-Progress -Activity "Adding License SKUs and Service Plan IDs to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))

    # Get License SKUs
    $skus = Get-MgSubscribedSku -ErrorAction Continue -ErrorVariable +allErrors | Select AppliesTo, ConsumedUnits, PrepaidUnits, ServicePlans, SkuId, SkuPartNumber

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
    $tenantStatsHash["LicenseSKUs"] = $licenseSkuHash
    $tenantStatsHash["ServicePlans"] = $servicePlanHash
}
#Gather all Office 365 Admins
function Get-AllOffice365Admins {
    param ()
    $start = Get-Date
    $tenantStatsHash["Office365-Admins"] = @{}

    Write-Host "Gathering All Admins ..." -ForegroundColor Cyan -nonewline
    $azureRoles = Get-MgDirectoryRole | Select DisplayName, ID, Description
    $progresscounter = 1
    $totalCount = $azureRoles.count
    $ProgressPreference = "Continue"

    foreach ($role in $azureRoles) {
        $currentRole = @()
        Write-ProgressHelper -ID 1 -Activity "Gathering $($role.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -StartTime $start 
        
        $userList = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id
        $progresscounter2 = 0
        $totalCount2 = $userList.count
        
        $ProgressPreference = "Continue"
        foreach ($user in $userList) {
            $progresscounter2 += 1
            $progresspercentcomplete = [math]::Round((($progresscounter2 / $totalCount2)*100),2)
            $progressStatus = "["+$progresscounter2+" / "+$totalCount2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Role Details: $($user.additionalproperties["displayName"])"

            $currentuser = [PSCustomObject]@{
                "Role" = $role.DisplayName
                "RoleDescription" = $role.Description
                "DisplayName" = $user.additionalproperties["displayName"]
                "UserPrincipalName" = $user.additionalproperties["userPrincipalName"]
                "userType" = $user.additionalproperties["userType"]
                "homepage" = $user.additionalproperties["homepage"]
            }
        
            $currentRole += $currentuser
        }
        $tenantStatsHash["Office365-Admins"][$role.DisplayName] = $currentRole
    }
    Write-Progress -Activity "Gathering $($role.DisplayName)" -Completed
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
    return $tenantStatsHash["Office365-Admins"]

    <# Convert the hashtable to custom objects and export to CSV
    $tenantStatsHash["Office365-Admins"].Values | ForEach-Object {
        $_.Values | ForEach-Object {
            $_ | Export-Csv -Path 'C:\Users\AMedranoA\Desktop\6-21-Admins.csv' -Append -NoTypeInformation
        }
    }
    #>

}
# Get all Office 365 Domains
function Get-AllOffice365Domains {
    param ()
    $start = Get-Date
    # Get all the domains
    $domains = Get-MgDomain

    $progresscounter = 1
    $totalCount = $domains.count
    $ProgressPreference = "Continue"

    # Prepare the results array
    $currentDomain = @()
    $tenantStatsHash["Domains"] = @{}

    foreach ($domain in $domains) {
        Write-ProgressHelper -ID 1 -Activity "Gathering $($domain.Id)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -StartTime $start 
        # Get the DNS records
        $aRecords = Resolve-DnsName -Name $domain.Id -Server 8.8.8.8 -Type A -ErrorAction SilentlyContinue
        $mxRecords = Resolve-DnsName -Name $domain.Id -Server 8.8.8.8 -Type MX -ErrorAction SilentlyContinue
        $NSRecords = Resolve-DnsName -Name $domain.Id -Server 8.8.8.8 -Type NS -ErrorAction SilentlyContinue

        # Add to the results array
        $currentDomain = New-Object PSObject -Property ([ordered]@{
            Domain = $domain.Id
            Verified = $domain.IsVerified
            NSRecords = if ($NSRecords) { ($NSRecords.NameHost -join ","| Out-String).Trim() } else { $null }
            ARecords = if ($aRecords) { ($aRecords.IPAddress -join "," | Out-String).Trim() } else { $null }
            MXRecords = if ($mxRecords) { ($mxRecords.NameExchange -join "," | Out-String).Trim()} else { $null }
            Office365MailExchanger = if (($mxRecords.NameExchange | Out-String).Trim() -like "*protection.outlook.com") {$true} else { $False }
        })
        $tenantStatsHash["Domains"][$domain.Id] = $currentDomain
    }

    Write-Progress -Activity "Gathering $($domain.Id)" -Completed

    # Display the results
    #return $results

    return $tenantStatsHash["Domains"]
}
###########################################################################################################################################

#Get all OneDrive Personal Sites data and add to hash table with UserPrincipalName as HASH KEY and OneDrive data as HASH VALUE
function Get-AllOneDriveDetails {
    $start = Get-Date

    $OneDriveDataHash = @{}

    Write-Host "Getting all OneDrive site data..." -ForegroundColor Cyan -nonewline
    $OneDriveSite = Get-SPOSite -Filter "URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop -ErrorVariable +allErrors | Select LastContentModifiedDate, StorageUsageCurrent, LockIssue, LockState, Url, Owner, StorageQuota, Title, IsTeamsConnected, TeamsChannelType 

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
    
    $tenantStatsHash["OneDrives"] = $OneDriveDataHash
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
}
function Get-AllSharePointSiteDetails {
    $start = Get-Date
    $SharePointSiteHash = @{}

    #Get all SharePoint sites for associating with Office 365 Groups / GroupMailboxes
    Write-Host "Getting all SharePoint site data..." -ForegroundColor Cyan -nonewline
    $SharePointSite = Get-SPOSite -IncludePersonalSite $false -Limit all -ErrorAction Stop -ErrorVariable +allErrors | Select LastContentModifiedDate, StorageUsageCurrent, LockIssue, LockState, Url, Owner, StorageQuota, Title, IsTeamsConnected, TeamsChannelType 

    #SharePoint data to SharePoint Hash
    #************************************************************************************
    foreach ($site in $SharePointSite) {
        #Add to hash table
        $key = $site.URL
        $value = $site
        $SharePointSiteHash[$key] = $value
    }
    $tenantStatsHash["SharePointSites"] = $SharePointSiteHash

    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
}

###########################################################################################################################################

#Consolidate Reporting for each user
function Combine-AllMailboxStats {
    param(
        [Parameter(Mandatory=$true)]
        [Hashtable]$tenantStatsHash
    )

    #Progress Bar Parameters Reset
    $allMailboxStats = [Ordered]@{}
    
    $progresscounter = 1
    $totalCount = $tenantStatsHash["AllMailboxes"].count
    $ProgressPreference = "Continue"
    Write-Host "Processing all mailboxes..." -ForegroundColor Cyan -nonewline

    foreach ($mailbox in $tenantStatsHash["AllMailboxes"].values) {
        #progress bar
        Write-ProgressHelper -Activity "Gathering $($mailbox.RecipientTypeDetails) Mailbox Details for $($mailbox.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -StartTime $start
        
        #null values
        $oneDriveData = $null
        $sharePointSiteData = $null
        $mgUser = $null
        $licenses = $null
        $allLicenses = @()
        $allDisabledPlans = @()
        $mbxStats = $null
        $MBXSizeGB = $null
        $ArchiveStats = $null

        #Pull MailboxStats, MGUserDetails, Licensing, and Disabled Service Plans
        #*******************************************************************************************************************
        $EmailAddresses = $mailbox | Select-Object -ExpandProperty EmailAddresses
        $onmicrosoftAlias = $EmailAddresses | Where-Object { $_.Address -like "*.onmicrosoft.com" } | Select-Object -First 1


        if($mailbox.ExchangeGuid) {
            $mbxStats = $tenantStatsHash["PrimaryMailboxStats"][$mailbox.ExchangeGuid.ToString()]
            if($mbxStats.TotalItemSize) {
                $MBXSizeGB = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
            } else {
                $MBXSizeGB = 0
            }
        }

        #If $mailbox represents a GroupMailbox object set $mgUser data to pull from GroupMailboxHash which contains group mailbox data
        if($mailbox.UserPrincipalName -and $mailbox.RecipientTypeDetails -ne "GroupMailbox") {
            $mgUser = $tenantStatsHash["all_MG-Users"][$mailbox.UserPrincipalName.ToString()]
            if($oneDriveData =  $tenantStatsHash["all_OneDrives"][$mailbox.UserPrincipalName]) {}
            else {$oneDriveData = $null}

            #Get licenses for this user
            $licenses = ($mgUser.AssignedLicenses)

            #Array for adding each license name to the allLicenses for this user
            $allLicenses = @()

            #Array for adding all disabled service names for this user
            $allDisabledPlans = @()

            #Process each license to get friendly names and disabled service plans for each license
            foreach($license in $licenses) {
                $licenseName = $tenantStatsHash["all_LicenseSKUs"][$license.SkuId].SkuPartNumber
                $allLicenses += $licenseName
                try {
                    foreach($disabledPlan in $license.DisabledPlans) {
                        #Write-Output $disabledPlan
                        $disabledPlanName = $tenantStatsHash["all_servicePlans"][$disabledPlan.toString()]
                        $allDisabledPlans += $disabledPlanName
                    }
                } catch {
                    $allDisabledPlans = $null
                }
            }
        }
        elseif($mailbox.RecipientTypeDetails -eq "GroupMailbox") {  
            $licenses = $null
            $allLicenses = @()
            $allDisabledPlans = @()

            $mgUser = $tenantStatsHash["GroupMailboxes"][$mailbox.ExchangeGuid.ToString()]
            $unifiedGroupData = $tenantStatsHash["all_UnifiedGroups"][$mailbox.ExchangeGuid.ToString()]

            if($unifiedGroupData.SharePointSiteUrl) {
                $sharePointSiteData = $tenantStatsHash["all_SharePointSites"][($unifiedGroupData.SharePointSiteUrl)]
            }
            else {$sharePointSiteData = $null}
        }
        else {
            $unifiedGroupData = $null
            $sharePointSiteData = $null
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
                "OnMicrosoftAlias" = $onmicrosoftAlias
                "EmailAddresses" = ($EmailAddresses -join ";")
                "DeliverToMailboxAndForward" = $mailbox.DeliverToMailboxAndForward
                "ForwardingAddress" = $mailbox.ForwardingAddress
                "ForwardingSmtpAddress" = $mailbox.ForwardingSmtpAddress
                "LitigationHoldEnabled" = $mailbox.LitigationHoldEnabled
                "LitigationHoldDuration" = $mailbox.LitigationHoldDuration
                "InPlaceHolds" = $mailbox.InPlaceHolds -join ";"
                "ArchiveStatus" = $mailbox.ArchiveStatus
                "RetentionPolicy" = $mailbox.RetentionPolicy
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
            $archiveStats = $tenantStatsHash["ArchiveMailboxStats"][$mailbox.ArchiveGuid.ToString()]
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

        #Get SharePoint/OneDrive Data
        if($OneDriveData) {
            $currentuser["OneDriveURL"] = $OneDriveData.URL
            $currentuser["OneDriveStorage"] = $OneDriveData.StorageUsageCurrent
            $currentuser["OneDriveStorage-GB"] = [math]::Round($OneDriveData.StorageUsageCurrent / 1024, 3)
            $currentuser["OneDriveLastContentModifiedDate"] = $OneDriveData.LastContentModifiedDate
            $currentuser["SharePointURL"] = $null
            $currentuser["SharePointStorage-GB"] = $null
            $currentuser["SharePointLastContentModifiedDate"] = $null
        }
        #Group Mailbox Associated SharePoint Site mapping
        elseif($sharePointSiteData) {    
            $currentuser["OneDriveURL"] = $null
            $currentuser["OneDriveStorage-GB"] = $null
            $currentuser["OneDriveLastContentModifiedDate"] = $null 
            $currentuser["SharePointURL"] = $sharePointSiteData.URL
            $currentuser["SharePointStorage"] = $sharePointSiteData.StorageUsageCurrent
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
        $tenantStatsHash["CombinedMailboxStats"] = $allMailboxStats
    return $tenantStatsHash
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
}
#Convert Hash Table to Custom Object Array for Export
function Convert-HashToArray {
    param (
        [Parameter(Mandatory=$true)]
        [Hashtable]$tenantStatsHashToConvert,
        [Parameter(Mandatory=$true)]
        [String]$tenant,
        [Parameter(Mandatory=$true)]
        [String]$table,
        [Parameter(Mandatory=$true)]
        [DateTime]$startTime
        
    )
    Write-Host "Converting Hash $($table) to Array for Export..." -ForegroundColor Cyan -nonewline
    $ExportArray = @()
    $totalCount = $tenantStatsHashToConvert.count    
    $progresscounter = 1
    $ProgressPreference = "Continue"

    foreach ($nestedKey in $tenantStatsHashToConvert.Keys) {
        Write-ProgressHelper -Activity "Converting $($nestedKey)" -ProgressCounter ($progresscounter++) -TotalCount $totalCount -StartTime $start
        #Write-Host $($nestedKey)
        $attributes = $tenantStatsHashToConvert[$nestedKey]
        if ($attributes -is [hashtable] -or $attributes -is [System.Collections.Specialized.OrderedDictionary]) {
            #Write-Host "$($nestedKey) is a hashtable or an ordered dictionary."
            Write-Host "." -foregroundcolor cyan -nonewline
            $customObject = New-Object -TypeName PSObject
            foreach ($attribute in $attributes.keys) {
                $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)_$($tenant)" -Value $attributes[$attribute]
            }
            $ExportArray += $customObject
        } 
        elseif ($attributes -is [array] -or $attributes -is [PSCustomObject] ) {
            Write-Host "." -foregroundcolor Yellow -nonewline
            #Write-Host "$($nestedKey) is an array or custom object."
            $ExportArray += $attributes
        }
        else {
            Write-Host "." -foregroundcolor red -nonewline
            #Write-Host "$($nestedKey) is of type: $($attributes.GetType().FullName)"
        }
    }
    Write-Host "Completed" -ForegroundColor Green
    Return $ExportArray
}
function Export-TenantStatsHashToExcel {
    param (
        [Parameter(Mandatory=$True)] [Hashtable]$tenantStatsHash,
        [Parameter(Mandatory=$True)] [Array]$ExportDetails
	)

    # Combine all CSV files into a single Excel file
    # Export each hashtable to a separate CSV file
    foreach ($table in $tenantStatsHash.Keys){
        $ExportStatsArray = Convert-HashToArray -table $table -tenantStatsHash $tenantStatsHash[$table] -tenant $global:tenant -StartTime $start
        # Use the table name to create a unique temporary file
        $tempPath = Join-Path -Path $env:TEMP -ChildPath ("{0}_{1}.csv" -f $table, $global:tenant)
        $ExportStatsArray | Export-Csv -Path $tempPath -NoTypeInformation
    }

    Write-Progress -Activity "Combine All CSV Files into One Excel" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
    Write-Host "Combine All CSV Files into One Excel" -ForegroundColor Black -BackgroundColor Yellow
    $ExportedCSVFiles = Get-ChildItem -Path $env:TEMP -Filter "*$global:tenant*.csv"
    $totalCount = $ExportedCSVFiles.Count
    $progresscounter = 1
    $start = Get-Date
    foreach ($file in $ExportedCSVFiles) {
        $worksheetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $worksheetName = $worksheetName.Replace("$global:tenant-","")
        Write-ProgressHelper -Activity "Adding Worksheet $($worksheetName) to $($ExportDetails[0])" -ProgressCounter ($progresscounter++) -TotalCount $totalCount -StartTime $start
        Import-Csv -Path $file.FullName | Export-Excel -Path $ExportDetails[0] -WorksheetName $worksheetName -ClearSheet
    }

    <# Delete the temporary CSV files
    Get-ChildItem -Path $env:TEMP -Filter "*$global:tenant*.csv" | Remove-Item
    #>
    
}
###########################################################################################################################################
### Main ###

Get-AllExchangeMailboxDetails
Get-allMGUserDetails
Get-AllUnifiedGroups
Get-AllPublicFolderDetails
Get-AllLicenseSKUs
Get-AllOneDriveDetails
Get-AllSharePointSiteDetails

Write-Host ""
Write-Host "Consolidating report data for each user / object..." -ForegroundColor Black -BackgroundColor Yellow
Write-Host ""

$start = Get-Date
#Combine Reports
$tenantStatsHash = Combine-AllMailboxStats -tenantStatsHash $tenantStatsHash

###########################################################################################################################################
### Export Reports ###

#Export Reports: Exports each individual hashtable to own CSV file and then combines into Excel file
Export-TenantStatsHashToExcel -tenantStatsHash $tenantStatsHash -ExportDetails $ExportDetails
Get-ErrorReportDetails

Write-Host "Completed in"((Get-Date) -$global:initialStart).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan
Write-Host ""
Write-Host "Recipient Count Table" -ForegroundColor Black -BackgroundColor Green

#Final Output of Recipient Counts
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

###########################################################################################################################################
