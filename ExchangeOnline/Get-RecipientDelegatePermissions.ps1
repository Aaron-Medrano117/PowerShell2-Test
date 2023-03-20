<# .SYNOPSIS
    This script can be used to pull delegatge permissions based on Calendar Permissions granted per mailbox. 
    Pulls unique values and exports lists of users sharing calendars and those not sharing calendars.

    Full List of Permissions are exported as "DelegatePermissions.csv". By Default exports to desktop.

    .PARAMETER OutputCSVFilePath
    Output File Path for Report. Can specify exactly where to save file and what to name it.
    .PARAMETER OutputCSVFolderPath
    Output Folder Path for Report. Designate where to save file as 'DelegatePermissions.csv"
    .PARAMETER OnPremises
    Switch used to specify if running On-Premises Exchange. Should support versions Exchange 2010 through Exchange 2016
    .PARAMETER Office365
    Switch used to specify if running against Office 365's Exchange Online.
    .PARAMETER CalendarPermissions
    Switch used to request Calendar Permissions. Exports all calendar folders of mailbox and their permissions
    .PARAMETER SendAs
    Switch used to request Send As Permissions for the mailboxes
    .PARAMETER FullAccess
    Switch used to request Full Access Permissions for the mailboxes
    .PARAMETER SendOnBehalf
    Switch used to request Send On Behalf Permissions for the mailboxes

   .EXAMPLE
   Pulls all available permissions for each mailbox in Office 365. Exports to default location of desktop.
   .\Get-RecipientDelegatePermissions.ps1 -Office365 -CalendarPermissions -FullAccess -SendAs -SendOnBehalf
   .EXAMPLE
   Pulls only Calendar Permissions for each mailbox in Office 365. Exports to user's documents folder.
   .\Get-RecipientDelegatePermissions.ps1 -Office365 -CalendarPermissions -OutputCSVFolderPath C:\user\documents
   .EXAMPLE
   Pulls all available permissions for each mailbox in On-Premises Exchange. Exports to default location of desktop.
   .\Get-RecipientDelegatePermissions.ps1 -OnPremises -CalendarPermissions -FullAccess -SendAs
   .EXAMPLE
   Pulls Full Access and Send As permissions for each mailbox in On-Premises Exchange. Exports to user's documents folder.
   .\Get-RecipientDelegatePermissions.ps1 -OnPremises -FullAccess -SendaAs -OutputCSVFolderPath C:\user\documents
#>

param (
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FILEPATH')] [string] $OutputCSVFilePath,
    [Parameter(Mandatory=$false,HelpMessage="Run against OnPremises Exchange?")] [switch]$OnPremises,
    [Parameter(Mandatory=$false,HelpMessage="Run against Office365 Exchange Online?")] [switch]$Office365,
    [Parameter(Mandatory=$false,HelpMessage="All Group Members Individually Exported?")] [switch]$AllGroupMembers,
    [Parameter(Mandatory=$false,HelpMessage="Run Calendar Permissions Report?")] [switch]$CalendarPermissions,
    [Parameter(Mandatory=$false,HelpMessage="Run Full Access Permissions Report?")] [switch]$FullAccess,
    [Parameter(Mandatory=$false,HelpMessage="Run Send OnBehalf Permissions Report?")] [switch]$SendOnBehalf,
    [Parameter(,Mandatory=$false,HelpMessage="Run Send As Permissions Report?")] [switch]$SendAs
)
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
    #$secondsElapsed = (Get-Date) – $global:start
    $global:secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($global:progressref – $progresscounter)
}

function ConvertTo-EmailAddressesLoop {
    param (
        [Parameter(Mandatory=$true,HelpMessage='InputArray to Convert EmailAddresses')] [array] $InputArray
    )
    $OutPutArray = @()
    $recipientCheck = @()
    foreach ($recipientObject in $InputArray) {
        #Check Perm User Mail Enabled; If OnPremises and If Office365
        if ($global:OnPremises) {
            if ($recipientCheck = Get-Recipient $recipientObject.DistinguishedName.ToString() -ea silentlycontinue) {
                $tempUser = $recipientCheck.PrimarySMTPAddress.ToString()
            }
            else {
                $tempUser = $recipientObject.Name.ToString()
            }
        }
        elseif ($global:Office365) {
            if ($recipientCheck = Get-EXORecipient $recipientObject.ToString() -ea silentlycontinue) {
                $tempUser = $recipientCheck.PrimarySMTPAddress.ToString()
            }
            else {
                $tempUser = $recipientObject.ToString()
            }
        }
        $OutPutArray += $tempUser
    }
    $OutPutArray
}

$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null

#Gather All Recipient
Write-Host "Gathering All Recipients .." -foregroundcolor cyan -nonewline
$allRecipients = Get-ExoRecipient -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*DiscoverySearchMailbox*"} | sort PrimarySmtpAddress
Write-Host "done" -foregroundcolor green

#ProgressBar
$progressref = ($allRecipients).count
$progresscounter = 0

#Build Array
$CollectPermissionsList = @()
$calendarPerms = @()
$objPermissions = @()
$SendAsPerms = @()

foreach ($obj in $allRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details for $($obj.DisplayName)"

    #Write-ProgressHelper -ProgressCounter ($progresscounter++) -Activity "Gathering $($role.DisplayName)" -ID 1 -TotalCount ($roles).count
    $primarySMTPAddress = $obj.PrimarySMTPAddress.tostring()
    $identity = $obj.Identity.ToString()
    $recipientTypeDetails = $obj.RecipientTypeDetails
    Write-Host "Checking Perms for $($obj.DisplayName)" -ForegroundColor Cyan -NoNewline
    Write-Host ".." -ForegroundColor Yellow -NoNewline

    if ($SendAs) {
        #Gather Send As Perms; If On-Premises, If Office 365
        if ($OnPremises) {
            [array]$SendAsPerms = Get-User $identity -EA SilentlyContinue | Get-ADPermission | Where {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "*S-1-*"}         
        }
        elseif ($Office365) {
            [array]$SendAsPerms = Get-RecipientPermission $identity -EA SilentlyContinue | Where {$_.Trustee.ToString() -ne "NT Authority\Self" -and $_.Trustee.ToString() -notlike "*S-1-*"}
        }
        #If Permissions Found
        if ($SendAsPerms) {
            Write-Host "SendAs.." -NoNewline -foregroundcolor DarkCyan
            #Output Perms
            foreach ($perm in $SendAsPerms) {
                $accessRights = $perm.AccessRights -join ","
                #Check Perm User Mail Enabled; If OnPremises and If Office365
                if ($OnPremises) {
                    if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.User.ToString()
                        $permRecipientDetails = $null
                    }
                }
                elseif ($Office365) {
                    if ($recipientCheck = Get-Recipient $perm.Trustee.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.Trustee.ToString()
                        $permRecipientDetails = $null
                    }
                }
                #If Recipient Is Group, Output Group Members Individually
                if ($AllGroupMembers) {
                    if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                        $GroupMembers = @()
                        if ($GroupMembers = Get-DistributionGroupMember $permUser) {
                            foreach ($member in $GroupMembers){
                            $currentPerm = new-object PSObject
                            $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                            $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                            $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                            $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "FullAccess"
                            $currentPerm | add-member -type noteproperty -name "PermUser" -Value $member.PrimarySMTPAddress.ToString() -Force
                            $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $member.RecipientTypeDetails -Force                   
                            $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                            $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
                            $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                            $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                            $CollectPermissionsList += $currentPerm
                            Write-Host "." -ForegroundColor DarkYellow -NoNewline
                        }
                        }
                    }
                }
                #All Else, Output Perm users details
                else {
                    $currentPerm = new-object PSObject
                    $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                    $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                    $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                    $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "SendAs"
                    $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                    $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                    
                    $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                    $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $null -Force
                    $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                    $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                    $CollectPermissionsList += $currentPerm
                    Write-Host "." -ForegroundColor Yellow -NoNewline
                }
            }
        }
    }
    if ($FullAccess) {
        if ($recipientTypeDetails -like "*Mailbox") {
            #Gather Full Access Perms; If On-Premises, If Office 365
            if ($OnPremises) {
                $objPermissions = Get-MailboxPermission $primarySMTPAddress -EA SilentlyContinue | Where {$_.User.ToString() -notlike "NT Authority*" -and $_.User.ToString() -notlike "*S-1-*" -and $_.AccessRights -like "FullAccess" -and $_.User.ToString() -notlike "*Exchange Domain Servers*" -and $_.User.ToString() -notlike "*Exchange Servers*" -and $_.User.ToString() -notlike "*Domain Admins*"}
            }
            elseif ($Office365) {
                $objPermissions = Get-EXOMailboxPermission $primarySMTPAddress -EA SilentlyContinue | Where {$_.User.ToString() -notlike "NT Authority*" -and $_.User.ToString() -notlike "*S-1-*"}
                }
            #If Permissions Found
            if ($objPermissions) {
                Write-Host "FullAccess.." -NoNewline -foregroundcolor DarkCyan
                #Output Perms
                foreach ($perm in $objPermissions) {
                    $accessRights = $perm.AccessRights -join ","
                    #Check Perm User Mail Enabled
                    if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.User.ToString()
                        $permRecipientDetails = $null
                    }
                    #If Recipient Is Group, Output Group Members Individually
                    if ($AllGroupMembers) {
                        if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                            $GroupMembers = @()
                            if ($GroupMembers = Get-DistributionGroupMember $permUser) {
                                foreach ($member in $GroupMembers){
                                $currentPerm = new-object PSObject
                                $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                                $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                                $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                                $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "FullAccess"
                                $currentPerm | add-member -type noteproperty -name "PermUser" -Value $member.PrimarySMTPAddress.ToString() -Force
                                $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $member.RecipientTypeDetails -Force                   
                                $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                                $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
                                $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                                $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                                $CollectPermissionsList += $currentPerm
                                Write-Host "." -ForegroundColor DarkYellow -NoNewline
                            }
                            }
                        }
                    } 
                    #All Else, Output Perm Details
                    else {
                        $currentPerm = new-object PSObject              
                        $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                        $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                        $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                        $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "FullAccess"
                        $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                        $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force
                        $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                        $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $null -Force
                        $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                        $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                        $CollectPermissionsList += $currentPerm
                        Write-Host "." -ForegroundColor Yellow -NoNewline 
                    }  
                }
            }
        }
    }
    if ($CalendarPermissions) {
        if ($recipientTypeDetails -like "*Mailbox") {
            #Gather Calendar; If On-Premises, If Office 365
            if ($OnPremises) {[array]$calendars = Get-MailboxFolderStatistics $identity | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}}
            elseif ($Office365) {[array]$calendars = Get-EXOMailboxFolderStatistics $identity | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}}

            #Check Calendars
            foreach ($calendar in $calendars) {
                $folderPath = $calendar.FolderPath.Replace('/','\')
                $id = "$primarySMTPAddress" + ":$folderPath"
                #Gather Calendar Permissions; If On-Premises, If Office 365
                if ($OnPremises) {[array]$calendarPerms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.User.ADRecipient -and $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" -and $_.User.ToString() -notlike "*S-1-*" -and $_.User.ADRecipient.PrimarySmtpAddress.ToString() -ne $primarySMTPAddress} }
                elseif ($Office365) {[array]$calendarPerms = Get-EXOMailboxFolderPermission $id -EA SilentlyContinue | Where {$_.user.Usertype.value -ne "Default" -and $_.user.usertype.value -ne "Anonymous" -and $_.user.usertype.value -notlike "*S-1-*"} }
                
                #output Per Calendar
                if ($calendarPerms) {
                    Write-Host "CalendarPerm.." -NoNewline -foregroundcolor DarkCyan
                    Write-Host $folderPath -ForegroundColor Green -NoNewline
                    #Output Perms
                    foreach ($perm in $calendarPerms) {
                        $accessRights = $perm.AccessRights -join ","
                        $SharingPermissionFlags = $perm.SharingPermissionFlags -join ","
                        $recipientChecks = @()
                        #Check Perm User Mail Enabled
                        if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                            $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                            $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                        }
                        else {
                            $permUser = $perm.User.ToString()
                            $permRecipientDetails = $null
                        }
                        #If Recipient Is Group, Output Group Members Individually   
                        if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                            $GroupMembers = @()
                            if ($GroupMembers = Get-DistributionGroupMember $permUser) {
                                foreach ($member in $GroupMembers){
                                $currentPerm = new-object PSObject
                                $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                                $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                                $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                                $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "FullAccess"
                                $currentPerm | add-member -type noteproperty -name "PermUser" -Value $member.PrimarySMTPAddress.ToString() -Force
                                $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $member.RecipientTypeDetails -Force                   
                                $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                                $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
                                $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $id
                                $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags
                                $CollectPermissionsList += $currentPerm
                                Write-Host "." -ForegroundColor DarkYellow -NoNewline
                            }
                            }
                        }
                        #If Multiple Recipients, Output Recipients Individually   
                        elseif ($recipientChecks = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                            foreach ($recipient in $recipientChecks) {
                                $currentPerm = new-object PSObject              
                                $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                                $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                                $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                                $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "Calendar"
                                $currentPerm | add-member -type noteproperty -name "PermUser" -Value $recipient.PrimarySMTPAddress.ToString() -Force
                                $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                                $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                                $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $null -Force
                                $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $id
                                $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags
                                $CollectPermissionsList += $currentPerm
                                Write-Host "." -ForegroundColor Yellow -NoNewline
                            }
                        }
                        #All Else, Output Perm users details
                        else {
                            $currentPerm = new-object PSObject              
                            $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                            $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                            $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                            $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "Calendar"
                            $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                            $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force
                            $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                            $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $null -Force
                            $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $id
                            $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags
                            $CollectPermissionsList += $currentPerm
                            Write-Host "." -ForegroundColor Yellow -NoNewline
                        }
                    }
                }
            }
        }   
    }
    #Send on Behalf Check
    if ($SendOnBehalf) {
        #If Recipient Detail is Mailbox
        if ($recipientTypeDetails -like "*Mailbox") {
            #Gather SendOnBehalf; If On-Premises, If Office 365
            if ($OnPremises) {[array]$SendOnBehalfToPerms = (Get-Mailbox $primarySMTPAddress -ea SilentlyContinue).GrantSendOnBehalfTo}
            elseif ($Office365) {[array]$SendOnBehalfToPerms =  (Get-EXOMailbox $primarySMTPAddress -ea SilentlyContinue).GrantSendOnBehalfTo}

            #Check SendOnBehalf
            if ($SendOnBehalfToPerms) {
                Write-Host "SendOnBehalfTo.." -NoNewline -foregroundcolor DarkCyan
                #Output Perms
                foreach ($perm in $SendOnBehalfToPerms) {
                    #Check Perm User Mail Enabled
                    if ($OnPremises) {$permObject = $perm.DistinguishedName.ToString()}
                    elseif ($Office365) {$permObject = $perm.ToString()}
                    if ($recipientCheck = Get-ExoRecipient $permObject -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $permObject
                        $permRecipientDetails = $null
                    }
                    #If Recipient Is Group, Output Group Members Individually
                    if ($AllGroupMembers) {
                        if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                            $GroupMembers = @()
                            if ($GroupMembers = Get-DistributionGroupMember $permUser) {
                                foreach ($member in $GroupMembers){
                                $currentPerm = new-object PSObject
                                $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                                $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                                $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                                $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "FullAccess"
                                $currentPerm | add-member -type noteproperty -name "PermUser" -Value $member.PrimarySMTPAddress.ToString() -Force
                                $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $member.RecipientTypeDetails -Force                   
                                $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                                $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
                                $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                                $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                                $CollectPermissionsList += $currentPerm
                                Write-Host "." -ForegroundColor DarkYellow -NoNewline
                            }
                            }
                        }
                    }
                    #All Else, Output Perm users details
                    else {
                        $currentPerm = new-object PSObject
                        $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                        $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                        $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                        $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "SendOnBehalfTo"
                        $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                        $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force     
                        $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $null
                        $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $null -Force
                        $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                        $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                        $CollectPermissionsList += $currentPerm
                        Write-Host "." -ForegroundColor Yellow -NoNewline
                    }
                }
            }
        } 
    }
    Write-Host "done" -ForegroundColor Green
}
Write-host ""

if ($OutputCSVFolderPath) {
    $CollectPermissionsList | Export-Csv "$OutputCSVFolderPath\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
    Write-host "Exported file 'DelegatePermissions.csv' List to $OutputCSVFolderPath" -ForegroundColor Cyan
}
elseif ($OutputCSVFilePath) {
    $CollectPermissionsList | Export-Csv $OutputCSVFilePath -NoTypeInformation -Encoding UTF8
    Write-host "Exported Permissions List to $OutputCSVFilePath" -ForegroundColor Cyan
}
else {
    try {
        $CollectPermissionsList | Export-Csv "$HOME\Desktop\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
        Write-host "Exported 'DelegatePermissions.csv' List to $HOME\Desktop" -ForegroundColor Cyan
    }
    catch {
        Write-Warning -Message "$($_.Exception)"
        Write-host ""
        $OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
        $CollectPermissionsList | Export-Csv "$OutputCSVFolderPath\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
    }
}  
