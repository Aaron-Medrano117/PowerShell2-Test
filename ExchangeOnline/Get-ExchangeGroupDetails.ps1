<# .SYNOPSIS
    This script can be used to pull delegatge permissions based on Calendar Permissions granted per mailbox.
    Pulls unique values and exports lists of users sharing calendars and those not sharing calendars
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
   Pulls all available mail groups in Office 365 including SendAs. Provides first 10 members and owners per group. Exports to default location of desktop.
   .\Get-ExchangeGroupDetails.ps1 -Office365 -SummaryMembers
   .EXAMPLE
   Pulls only Calendar Permissions for each mailbox in Office 365. Exports to user's documents folder.
   .\Get-ExchangeGroupDetails.ps1 -Office365 -SummaryMembers -OutputCSVFolderPath C:\user\documents
   .EXAMPLE
   Pulls all available permissions for each mailbox in On-Premises Exchange. Exports to Excel file path and specified worksheet name
   .\Get-ExchangeGroupDetails.ps1 -OnPremises -FullMembers -OutputExcelFilePath C:\user\documents\allmailgroupsdetails.xlsx -WorkSheetName "MailGroupsDetails"
    .EXAMPLE
   Pulls Full Access and Send As permissions for each mailbox in On-Premises Exchange. Exports to user's documents folder.
   .\Get-ExchangeGroupDetails.ps1 -OnPremises $SummaryMembers -OutputCSVFolderPath C:\user\documents
#>

param (
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file in CSV format? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file  in CSV format? Please provide full FILEPATH')] [string] $OutputCSVFilePath,
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file  in EXCEL (xlsx) format? Please provide full FILEPATH')] [string] $OutputExcelFilePath,
    [Parameter(Mandatory=$false,HelpMessage='What is the Excel Sheet name?')] [string] $WorkSheetName,
    [Parameter(Mandatory=$false,HelpMessage="Run against OnPremises Exchange?")] [switch]$OnPremises,
    [Parameter(Mandatory=$false,HelpMessage="Run against Office365 Exchange Online?")] [switch]$Office365
)
$global:OnPremises = $OnPremises
$global:Office365 = $Office365

$allMailGroups = Get-Recipient -RecipientTypeDetails group -ResultSize unlimited -EA SilentlyContinue
$allMailGroups += Get-Recipient -RecipientTypeDetails MailNonUniversalGroup -ResultSize unlimited -EA SilentlyContinue
$allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize unlimited -EA SilentlyContinue
$allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize unlimited -EA SilentlyContinue  
$allMailGroups += Get-Recipient -RecipientTypeDetails DynamicDistributionGroup -ResultSize unlimited -EA SilentlyContinue
$allMailGroupDetails = @()

#Convert Names to EmailAddresses loop
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

#Progress Bar Helper
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
 
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
foreach ($object in $allMailGroups | sort DisplayName) {
    if ($ProgressPreference = "SilentlyContinue") {
        $ProgressPreference = "Continue"
    }
    $identity = $object.identity.tostring()
    $PrimarySMTPAddress = $object.PrimarySMTPAddress.ToString()
    Write-ProgressHelper -Activity "Gathering Group Details for $($PrimarySMTPAddress)" -ProgressCounter ($progresscounter++) -TotalCount ($allMailGroups).count

    #Get Groups Details
    Write-Host "$($object.RecipientTypeDetails) - $($PrimarySMTPAddress) .." -ForegroundColor Cyan -NoNewline
    #Create Output Array
    $currentobject = new-object PSObject
    $currentobject | add-member -type noteproperty -name "DisplayName" -Value $object.DisplayName -Force
    $currentobject | add-member -type noteproperty -name "Identity" -Value $identity -Force
    $currentobject | add-member -type noteproperty -name "Name" -Value $object.Name -Force
    $currentobject | add-member -type noteproperty -name "Alias" -Value $object.alias -Force
    $currentobject | add-member -type noteproperty -name "Notes" -Value $object.Notes -Force
    $currentobject | add-member -type noteproperty -name "IsDirSynced" -Value $object.IsDirSynced -Force
    $currentobject | add-member -type NoteProperty -name "HiddenFromAddressListsEnabled" -Value $object.HiddenFromAddressListsEnabled -Force
    $currentobject | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $PrimarySMTPAddress -Force
    $currentobject | add-member -type noteproperty -name "RecipientTypeDetails" -Value $object.RecipientTypeDetails -Force
    
    ##Clear Group Details from Previous session
    $groupDetails = @()
    $ManagedByRecipients = @()
    $groupOwners = @()
    $groupMembersCheck = @()
    $GroupMembersAddresses = @()
    $groupMembers = @()
    $GroupFilter = @()
    $sendAsRecipients = @()
    $AcceptMessagesOnlyFromRecipients = @()
    $AcceptMessagesOnlyFromDLMembersRecipients = @()
    $AcceptMessagesOnlyFromSendersOrMembersRecipients = @()
    $RejectMessagesOnlyFromRecipients = @()
    $RejectMessagesOnlyFromDLMembersRecipients = @()
    $RejectMessagesOnlyFromSendersOrMembersRecipients = @()
    $ModeratedByRecipients = @()
    ## Get Email Addresses
    $EmailAddresses = $object | select -expandProperty EmailAddresses
    if ($SendAsCheck) {
        ##Gather Send As Perms; If On-Premises, If Office 365
        if ($OnPremises) {
            [array]$SendAsPerms = Get-ADPermission $identity -EA SilentlyContinue | Where {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "*S-1-*"}
        }
        elseif ($Office365) {
            [array]$SendAsPerms = Get-EXORecipientPermission $identity -EA SilentlyContinue | Where {$_.Trustee.ToString() -ne "NT Authority\Self" -and $_.Trustee.ToString() -notlike "*S-1-*"}
        }
        ###If Permissions Found
        if ($SendAsPerms) {
            $SendAsRecipients = @()
            #Output Perms
            $SendAsRecipients = ConvertTo-EmailAddressesLoop -InputArray $SendAsPerms
        }
    }
    ##Check If Dynamic Distribution Group
    if ($object.RecipientTypeDetails -eq "DynamicDistributionGroup") {
        $groupDetails = Get-DynamicDistributionGroup $PrimarySMTPAddress -ErrorAction SilentlyContinue
        #Recipient Filter Conversion
        $splitFilter = $groupDetails.RecipientFilter.ToString() -split '(\s+(?:-and|-not)\s+)'
        #$splitFilter = $filterString -split '(\s+(?:-and|-not)\s+)'
        $GroupFilter = @()
        foreach ($filter in $splitFilter) {
            if ($filter.StartsWith("(-not") -or $filter.StartsWith("-and")) {
                #Write-Host "." -ForegroundColor Red
                continue
            }
            elseif ($filter.Contains("-or")) {
                $optionResults = @()
                $filterDetails = $filter.Trim().Replace("(","").Replace(")","") -split "-or"
                $property = $filterDetails.Trim().Split(" ")[0]
                $options = @()
                foreach ($filterDetail in $filterDetails) {
                    $options += $filterDetail.Trim().Split("'")[1]
                    #Write-Host "." -ForegroundColor yellow
                }
                $optionResults += "$property is one of: $($options -join ', ')"
                $GroupFilter += $optionResults
            }
            elseif ($filter.Contains("eq")) {
                $property = ($filter.Trim().Split(" ")[0]).Replace("(","")
                $value = $filter.Trim().Split("'")[1]
                $GroupFilter += "$property equals $value"
                #Write-Host "." -ForegroundColor green
            }
        }
        #Group Members
        if ($OnPremises) {
            ###need to look into this one further
            $groupMembersCheck = $groupDetails.RecipientFilter
        }
        elseif ($Office365) {
            $groupMembersCheck = Get-DynamicDistributionGroupMember $PrimarySMTPAddress -ErrorAction SilentlyContinue -ResultSize unlimited -warningaction silentlycontinue
        }         
    }
    #Check if Group is DistributionGroup type
    elseif ($object.RecipientTypeDetails -eq "MailUniversalDistributionGroup" -or $object.RecipientTypeDetails -eq "MailUniversalSecurityGroup" -or $object.RecipientTypeDetails -eq "MailNonUniversalGroup") {
        $groupDetails = Get-DistributionGroup $PrimarySMTPAddress -ErrorAction SilentlyContinue
        $groupMembersCheck = Get-DistributionGroupMember $PrimarySMTPAddress -ResultSize unlimited -ErrorAction SilentlyContinue
    }
    #Check if Group Mailbox
    elseif ($object.RecipientTypeDetails -eq "GroupMailbox") {
        $groupDetails = Get-UnifiedGroup $PrimarySMTPAddress -ErrorAction SilentlyContinue
        $groupMembersCheck = Get-UnifiedGroupLinks -Identity $object -LinkType Member -ResultSize unlimited -ErrorAction SilentlyContinue
    }
    #Check Group Owners Size and Get Owners Addresses
    if ($object.ManagedBy.count -ge 1) {
        #Output Perms
        $ManagedByRecipients = ConvertTo-EmailAddressesLoop -InputArray $object.ManagedBy
        if ($ManagedByRecipients.count -gt 500) {
            $groupOwners= $ManagedByRecipients[0..499]
        }
        else {
            $groupOwners = $ManagedByRecipients
        }
        Write-Host "$($ManagedByRecipients.count) Owners found .." -ForegroundColor Yellow -NoNewline
    }
    #Check Group Members Size and Get Group Addresses
    if ($groupMembersCheck.count -ge 1) {
        if ($object.RecipientTypeDetails -eq "DynamicDistributionGroup" -or $object.RecipientTypeDetails -eq "GroupMailbox") {
            if ($groupMembersCheck.count -gt 20) {
                $groupMembers = $groupMembersCheck[0..20]
            }
            else {
                $groupMembers = $groupMembersCheck
            }
        }
        if ($groupMembersCheck.count -gt 500) {
            $groupMembers = $groupMembersCheck[0..499]
            }
        else {
            $groupMembers = $groupMembersCheck
        }
        $GroupMembersAddresses = ConvertTo-EmailAddressesLoop -InputArray $groupMembers
        Write-Host "$($groupMembersCheck.count) Members found .." -ForegroundColor Yellow -NoNewline
    }
    ## Gather Restricted Senders, Moderators, and Rejects
    ### AcceptMessagesOnlyFROM
    if ($groupDetails.AcceptMessagesOnlyFrom) {
        $AcceptMessagesOnlyFromRecipients = ConvertTo-EmailAddressesLoop -InputArray $groupDetails.AcceptMessagesOnlyFrom
    }
    ### AcceptMessagesOnlyFromDLMembers
    if ($groupDetails.AcceptMessagesOnlyFromDLMembers) {
        $AcceptMessagesOnlyFromDLMembersRecipients = ConvertTo-EmailAddressesLoop -InputArray $groupDetails.AcceptMessagesOnlyFromDLMembers
    }
    ### AcceptMessagesOnlyFromSendersOrMembers
    if ($groupDetails.AcceptMessagesOnlyFromSendersOrMembers) {
        $AcceptMessagesOnlyFromSendersOrMembersRecipients = ConvertTo-EmailAddressesLoop -InputArray $groupDetails.AcceptMessagesOnlyFromSendersOrMembers
    }
    ### RejectMessagesOnlyFROM
    if ($groupDetails.RejectMessagesOnlyFrom) {
        $RejectMessagesOnlyFromRecipients = ConvertTo-EmailAddressesLoop -InputArray $groupDetails.RejectMessagesOnlyFrom
    }
    ### RejectMessagesOnlyFromDLMembers
    if ($groupDetails.RejectMessagesOnlyFromDLMembers) {
        $RejectMessagesOnlyFromDLMembersRecipients = ConvertTo-EmailAddressesLoop -InputArray $groupDetails.RejectMessagesOnlyFromDLMembers
    }

    ### RejectMessagesOnlyFromSendersOrMembers
    if ($groupDetails.RejectMessagesOnlyFromSendersOrMembers) {
        $RejectMessagesOnlyFromSendersOrMembersRecipients = ConvertTo-EmailAddressesLoop -InputArray $groupDetails.AcceptMessagesOnlyFromSendersOrMembers
    }
    ### ModeratedBy
    if ($groupDetails.ModeratedBy) {
        $ModeratedByRecipients = ConvertTo-EmailAddressesLoop -InputArray $groupDetails.ModeratedBy    
    }
    
    #Output Group Details
    $currentobject | add-member -type NoteProperty -name "ResourceProvisioningOptions" -Value ($groupDetails.ResourceProvisioningOptions -join ",") -force
    $currentobject | add-member -type NoteProperty -name "IsMailboxConfigured" -Value $groupDetails.IsMailboxConfigured -force
    $currentobject | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ",") -Force
    $currentobject | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("X500:" + $groupdetails.LegacyExchangeDN) -Force
    $currentobject | add-member -type noteproperty -name "OwnersCount" -Value ($ManagedByRecipients | measure-object).count -Force
    $currentobject | add-member -type noteproperty -name "MembersCount" -Value ($groupMembersCheck | measure-object).count -Force
    $currentobject | add-member -type noteproperty -name "Owners" -Value ($groupOwners -join ",") -Force
    $currentobject | add-member -type noteproperty -name "Members" -Value ($GroupMembersAddresses -join ",") -Force
    $currentobject | add-member -type NoteProperty -name "DynamicGroup_RecipientFilter" -Value ($GroupFilter -join ";") -force
    $currentobject | add-member -type NoteProperty -name "HiddenGroupMembershipEnabled" -Value ($groupDetails.HiddenGroupMembershipEnabled -join ",") -force
    $currentobject | add-member -type NoteProperty -name "ModeratedBy" -Value ($ModeratedByRecipients -join ",") -force
    $currentobject | add-member -type NoteProperty -name "SendAs" -Value ($SendAsRecipients -join ",") -Force
    $currentobject | add-member -type NoteProperty -name "GrantSendOnBehalfTo" -Value ($groupDetails.GrantSendOnBehalfTo -join ",") -force
    $currentobject | add-member -type NoteProperty -name "RequireSenderAuthenticationEnabled" -Value ($groupDetails.RequireSenderAuthenticationEnabled -join ",") -force
    $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFrom" -Value ($AcceptMessagesOnlyFromRecipients -join ",") -force
    $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromDLMembers" -Value ($AcceptMessagesOnlyFromDLMembersRecipients -join ",") -force
    $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromSendersOrMembers" -Value ($AcceptMessagesOnlyFromSendersOrMembersRecipients -join ",") -force
    $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFrom" -Value ($RejectMessagesOnlyFromRecipients-join ",") -force
    $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromDLMembers" -Value ($RejectMessagesOnlyFromDLMembersRecipients -join ",") -force
    $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromSendersOrMembers" -Value ($RejectMessagesOnlyFromSendersOrMembersRecipients -join ",") -force
    $currentobject | add-member -type NoteProperty -name "AccessType" -Value $groupDetails.AccessType -force
    $currentobject | add-member -type NoteProperty -name "AllowAddGuests" -Value $groupDetails.AllowAddGuests -force
    $currentobject | add-member -type NoteProperty -name "SharePointSiteUrl" -Value $groupDetails.SharePointSiteUrl -force
    $allMailGroupDetails += $currentobject
    Write-Host "Done" -ForegroundColor Green
}
#Export
if ($OutputCSVFolderPath) {
    $allMailGroupDetails | Export-Csv "$OutputCSVFolderPath\AllMailGroupDetails.csv" -NoTypeInformation -Encoding UTF8
    Write-host "Exported 'AllMailGroupDetails.csv' List to $OutputCSVFolderPath" -ForegroundColor Cyan
}
elseif ($OutputCSVFilePath) {
    $allMailGroupDetails | Export-Csv $OutputCSVFilePath -NoTypeInformation -Encoding UTF8
    Write-host "Exported All Mail Group Details to $OutputCSVFilePath" -ForegroundColor Cyan
}
elseif ($OutputExcelFilePath) {
    if ($WorkSheetName) {
        $allMailGroupDetails | Export-Excel -WorksheetName $WorkSheetName -Path $OutputExcelFilePath
        Write-host "Exported All Mail Group Details to $OutputExcelFilePath and $WorkSheetName worksheet" -ForegroundColor Cyan
    }
    else {
        $allMailGroupDetails | Export-Excel -Path $OutputExcelFilePath
        Write-host "Exported All Mail Group Details to $OutputExcelFilePath" -ForegroundColor Cyan
    }
}
else {
    try {
        $allMailGroupDetails | Export-Csv "$HOME\Desktop\AllMailGroupDetails.csv" -NoTypeInformation -Encoding UTF8
        Write-host "Exported 'AllMailGroupDetails.csv' List to $HOME\Desktop" -ForegroundColor Cyan
    }
    catch {
        Write-Warning -Message "$($_.Exception)"
        Write-host ""
        $OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
        $allMailGroupDetails | Export-Csv "$OutputCSVFolderPath\AllMailGroupDetails.csv" -NoTypeInformation -Encoding UTF8
    }
}