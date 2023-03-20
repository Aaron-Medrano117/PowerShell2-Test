#Get DistributionGroups

function Import-ExchangeAndMsOnlineModule() {
    if (((Get-Module -Name "ExchangeOnlineManagement") -ne $null) -or ((Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue) -ne $null)) {
        return;
    }
    else {
        Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 2 can be found at https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"

    }
    if (((Get-Module -Name "MSOnline") -ne $null) -or ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null)) {
        return;
    }
    else {
        Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator"
    }

    $EXOmoduleLocation -eq "$env:ProgramFiles\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.5\netFramework\ExchangeOnlineManagement.psm1"
            if (Test-Path $EXOmoduleLocation)  {
                Import-Module -Name ExchangeOnlineManagement
                return
            }

    $MsOnlinemoduleLocation -eq "$env:ProgramFiles\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.5\netFramework\ExchangeOnlineManagement.psm1"
            if (Test-Path $MsOnlinemoduleLocation)  {
                Import-Module -Name MSOnline
                return
            }
}


##Get DistributionGroups
function Get-GroupDetails {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath
    )
    $allMailGroups = Get-Recipient -RecipientTypeDetails group -ResultSize unlimited -EA SilentlyContinue
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailNonUniversalGroup -ResultSize unlimited -EA SilentlyContinue
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize unlimited -EA SilentlyContinue
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize unlimited -EA SilentlyContinue   
    $allMailGroups += Get-Recipient -RecipientTypeDetails DynamicDistributionGroup -ResultSize unlimited -EA SilentlyContinue
    $allGroupDetails = @()

    #ProgressBarA
    $progressref = ($allMailGroups).count
    $progresscounter = 0

    #Gather Details
    foreach ($object in $allMailGroups) {
        #Variables
        $Identity = $object.Identity
        $PrimarySMTPAddress = $object.primarysmtpaddress.tostring()

        #ProgressBarB
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Group Details for $($object.DisplayName)"
        #Get Groups Details
        $groupMembers = @()
        $EmailAddresses = $object | select -expandProperty EmailAddresses

        #Gather Send As Perms; If On-Premises, If Office 365
        if ($OnPremises) {
            [array]$SendAsPerms = Get-ADPermission $identity -EA SilentlyContinue  | Where {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "*S-1-*"}         
        }
        elseif ($Office365) {
            [array]$SendAsPerms = Get-RecipientPermission $identity -EA SilentlyContinue | Where {$_.Trustee.ToString() -ne "NT Authority\Self" -and $_.Trustee.ToString() -notlike "*S-1-*"}
        }
        #If Permissions Found
        if ($SendAsPerms) {
            $SendAsRecipients = @()
            #Output Perms
            foreach ($perm in $SendAsPerms) {
                #Check Perm User Mail Enabled; If OnPremises and If Office365
                if ($OnPremises) {
                    if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                    }
                    else {
                        $permUser = $perm.User.ToString()
                    }
                }
                elseif ($Office365) {
                    if ($recipientCheck = Get-Recipient $perm.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                    }
                    else {
                        $permUser = $perm.ToString()
                    }
                }
                $SendAsRecipients += $permUser
            }
        }

        #Gather ManagedBy; If On-Premises, If Office 365
        #If Permissions Found
        if ($Object.ManagedBy) {
            $ManagedByRecipients = @()
            #Output Perms
            foreach ($perm in $Object.ManagedBy) {
                #Check Perm User Mail Enabled; If OnPremises and If Office365
                if ($OnPremises) {
                    if ($recipientCheck = Get-Recipient $perm.DistinguishedName.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                    }
                    else {
                        $permUser = $perm.User.ToString()
                    }
                }
                elseif ($Office365) {
                    if ($recipientCheck = Get-Recipient $perm.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                    }
                    else {
                        $permUser = $perm.ToString()
                    }
                }
                $ManagedByRecipients += $permUser
            }
        }
        
        #Create Output Array
        $currentobject = new-object PSObject
        $currentobject | add-member -type noteproperty -name "DisplayName" -Value $object.DisplayName -Force
        $currentobject | add-member -type noteproperty -name "Name" -Value $object.Name -Force
        $currentobject | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $PrimarySMTPAddress -Force
        $currentobject | add-member -type noteproperty -name "IsDirSynced" -Value $object.IsDirSynced -Force
        $currentobject | add-member -type noteproperty -name "RecipientTypeDetails" -Value $object.RecipientTypeDetails
        $currentobject | add-member -type noteproperty -name "Alias" -Value $object.alias -Force
        $currentobject | add-member -type noteproperty -name "Notes" -Value $object.Notes -Force
        $currentobject | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ",")
        $currentobject | add-member -type noteproperty -name "ManagedBy" -Value ($ManagedByRecipients -join ",")
        $currentobject | add-member -type NoteProperty -name "HiddenFromAddressListsEnabled" -Value $object.HiddenFromAddressListsEnabled -force
        $currentobject | add-member -type NoteProperty -name "SendAs" -Value ($SendAsRecipients -join ",") -force

        #Pull DynamicDistributionGroup Details
        if ($object.RecipientTypeDetails -eq "DynamicDistributionGroup") {
            $groupDetails = Get-DynamicDistributionGroup $PrimarySMTPAddress
            $groupOwners = ($dynamicGroup.ManagedBy | Get-Mailbox -ErrorAction SilentlyContinue).PrimarySMTPAddress.ToString()
            $groupMembers = Get-DynamicDistributionGroupMember $dynamicGroup.PrimarySMTPAddress -ErrorAction SilentlyContinue -ResultSize unlimited            
        }
        #Pull Mail Group Details
        elseif ($object.RecipientTypeDetails -eq "MailUniversalDistributionGroup" -or $object.RecipientTypeDetails -eq "MailUniversalSecurityGroup") {
            $groupDetails = Get-DistributionGroup $PrimarySMTPAddress -ErrorAction SilentlyContinue
            $groupMembersCheck = Get-DistributionGroupMember $PrimarySMTPAddress -ResultSize unlimited -ErrorAction SilentlyContinue
            if ($groupMembersCheck.count -gt 100) {
                $groupMembers = $groupMembers[0..100]
            }
            else {
                $groupMembers = $groupMembersCheck
            }
            $groupOwners = ($distributionGroup.ManagedBy | Get-Mailbox -ErrorAction SilentlyContinue).PrimarySMTPAddress.ToString()

        }
        #Gather Group Mailbox
        elseif ($object.RecipientTypeDetails -eq "GroupMailbox") {
            $groupDetails = Get-UnifiedGroup $PrimarySMTPAddress
            $groupMembers = Get-UnifiedGroupLinks -Identity $PrimarySMTPAddress -LinkType Member -ResultSize unlimited
            $groupOwners = Get-UnifiedGroupLinks -Identity $PrimarySMTPAddress -LinkType Owner -ResultSize unlimited
        }

            $currentobject | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("X500:" + $groupDetails.LegacyExchangeDN)
            $currentobject | add-member -type noteproperty -name "GroupOwners" -Value ($groupOwners.PrimarySmtpAddress -join ",") -Force
            $currentobject | add-member -type noteproperty -name "MembersCount" -Value ($groupMembers | measure-object).count-Force
            $currentobject | add-member -type noteproperty -name "Members" -Value ($groupMembers.PrimarySmtpAddress -join ",") -Force
            $currentobject | add-member -type NoteProperty -name "ModeratedBy" -Value ($groupDetails.ModeratedBy -join ",") -force 
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFrom" -Value ($groupDetails.AcceptMessagesOnlyFrom -join ",") -force
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromDLMembers" -Value ($groupDetails.AcceptMessagesOnlyFromDLMembers -join ",") -force
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromSendersOrMembers" -Value ($groupDetails.AcceptMessagesOnlyFromSendersOrMembers -join ",") -force
            $currentobject | add-member -type NoteProperty -name "GrantSendOnBehalfTo" -Value ($groupDetails.GrantSendOnBehalfTo -join ",") -force
            $currentobject | add-member -type NoteProperty -name "RequireSenderAuthenticationEnabled" -Value ($groupDetails.RequireSenderAuthenticationEnabled -join ",") -force
            $currentobject | add-member -type NoteProperty -name "HiddenGroupMembershipEnabled" -Value ($groupDetails.HiddenGroupMembershipEnabled -join ",") -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFrom" -Value ($groupDetails.RejectMessagesOnlyFrom -join ",") -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromDLMembers" -Value ($groupDetails.RejectMessagesOnlyFromDLMembers -join ",") -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromSendersOrMembers" -Value ($groupDetails.RejectMessagesOnlyFromSendersOrMembers -join ",") -force
            $currentobject | add-member -type NoteProperty -name "AccessType" -Value $groupDetails.AccessType -force
            $currentobject | add-member -type NoteProperty -name "AllowAddGuests" -Value $groupDetails.AllowAddGuests -force
            $currentobject | add-member -type NoteProperty -name "IsMailboxConfigured" -Value $groupDetails.IsMailboxConfigured -force
            $currentobject | add-member -type NoteProperty -name "ResourceProvisioningOptions" -Value ($groupDetails.ResourceProvisioningOptions -join ",") -force

            $allGroupDetails += $currentobject

    }
    #Export
    $allGroupDetails | Export-Excel $OutputCSVFilePath
}


# COMBINED Set Attributes including ManagedBy, Approved Senders, BypassModerationFromSendersOrMembers, GrantSendOnBehalf, SendAs to DistributionGroups
$matchedDistributionGroups = Import-CSV
$matchedMailboxes = Import-CSV

$AllGroupErrors = @()
$progressref = ($matchedDistributionGroups).count
$progresscounter = 0
foreach ($user in $matchedDistributionGroups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding SendAs Perms to $($user.DisplayName_Destination)"
    Write-Host "Updating Attribute for $($user.DisplayName_Destination).. " -ForegroundColor Cyan -NoNewline

    Set-DistributionGroup -identity $user.PrimarySmtpAddress_Destination -MemberJoinRestriction $user.MemberJoinRestriction -warningaction silentlycontinue
    
    #Stamp AcceptMessagesOnlyFromSendersOrMembers on Group
    if ($user.AcceptMessagesOnlyFromSendersOrMembers) {
        Write-Host "Adding AcceptMessagesOnlyFromSendersOrMembers .. " -ForegroundColor DarkCyan -NoNewline
        #Gather Users
        $membersSendArray = $user.AcceptMessagesOnlyFromSendersOrMembers -split ","

        #Progress Bar 2
        $progressref2 = ($membersSendArray).count
        $progresscounter2 = 0
        foreach ($member in $membersSendArray) {
            #Member Check
            $memberCheck = @()
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}

            #Progress Bar 2a
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Setting AcceptMessagesOnlyFromSendersOrMembers to $($memberCheck.PrimarySmtpAddress_Destination)"

            #Set Attribute     
            try {
                Set-DistributionGroup -identity $user.PrimarySmtpAddress_Destination -AcceptMessagesOnlyFromSendersOrMembers @{add=$memberCheck.PrimarySmtpAddress_Destination} -ea Stop -warningaction silentlycontinue
                Write-Host ". " -ForegroundColor Green -NoNewline
            }
            catch {
                try {
                    $recipientCheck = Get-Recipient $member -EA SilentlyContinue
                    Set-DistributionGroup -identity $user.PrimarySmtpAddress_Destination -AcceptMessagesOnlyFromSendersOrMembers @{add=$recipientCheck.PrimarySmtpAddress} -ea Stop -warningaction silentlycontinue

                    Write-Host ". " -ForegroundColor yellow -NoNewline
                }
                catch {
                    Write-Host ". " -ForegroundColor red -NoNewline

                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "AcceptMessagesOnlyFromSendersOrMembers" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllGroupErrors += $currenterror           
                    continue
                }
            }
        }
    }

    #Stamp BypassModerationFromSendersOrMembers  on Group
    if ($user.BypassModerationFromSendersOrMembers) {
        Write-Host "Adding BypassModerationFromSendersOrMembers .. " -ForegroundColor DarkCyan -NoNewline
        #Gather Users
        $membersByPassArray = $user.BypassModerationFromSendersOrMembers -split ","

        #Progress Bar 2
        $progressref2 = ($membersByPassArray).count
        $progresscounter2 = 0
        foreach ($member in $membersByPassArray) {
            #Member Check
            $memberCheck = @()
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}

            #Progress Bar 2a
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Setting BypassModerationFromSendersOrMembers to $($memberCheck.PrimarySmtpAddress_Destination)"

            #Set Attribute      
            try {
                Set-DistributionGroup -identity $user.PrimarySmtpAddress_Destination -BypassModerationFromSendersOrMembers @{add=$memberCheck.PrimarySmtpAddress_Destination} -ea Stop -warningaction silentlycontinue
                Write-Host ". " -ForegroundColor Green -NoNewline
            }
            catch {
                try {
                    $recipientCheck = Get-Recipient $member -EA SilentlyContinue
                    Set-DistributionGroup -identity $user.PrimarySmtpAddress_Destination -BypassModerationFromSendersOrMembers @{add=$recipientCheck.PrimarySmtpAddress} -ea Stop -warningaction silentlycontinue

                    Write-Host ". " -ForegroundColor yellow -NoNewline
                }
                catch {
                    Write-Host ". " -ForegroundColor red -NoNewline

                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "BypassModerationFromSendersOrMembers" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllGroupErrors += $currenterror           
                    continue
                }
            }
        }
    }

    #Stamp ModeratedBy on Group
    if ($user.ModeratedBy) {
        Write-Host "Adding ModeratedBy .. " -ForegroundColor DarkCyan -NoNewline
        #Gather Users
        $membersModeratedByArray = $user.ModeratedBy -split ","

        #Progress Bar 2
        $progressref2 = ($membersModeratedByArray).count
        $progresscounter2 = 0
        foreach ($member in $membersModeratedByArray) {
            #Member Check
            $memberCheck = @()
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}

            #Progress Bar 2a
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Setting ModeratedBy to $($memberCheck.PrimarySmtpAddress_Destination)"

            #Set Attribute      
            try {
                Set-DistributionGroup -identity $user.PrimarySmtpAddress_Destination -ModeratedBy @{add=$memberCheck.PrimarySmtpAddress_Destination} -ea Stop -warningaction silentlycontinue
                Write-Host ". " -ForegroundColor Green -NoNewline
            }
            catch {
                Write-Host ". " -ForegroundColor red -NoNewline

                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "ModeratedBy" -Force
                $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
                $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllGroupErrors += $currenterror           
                continue
            }
        }
    }

    #Stamp SendOnBehalf on Group
    if ($user.GrantSendOnBehalfTo) {
        Write-Host "Adding SendOnBehalfTo .. " -ForegroundColor DarkCyan -NoNewline
        #Gather Perm Users
        $membersGrantSendOnBehalfArray = $user.GrantSendOnBehalfTo -split ","

        #Progress Bar 2
        $progressref2 = ($membersGrantSendOnBehalfArray).count
        $progresscounter2 = 0
        foreach ($member in $membersGrantSendOnBehalfArray) {
            #Member Check
            $memberCheck = @()
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}

            #Progress Bar 2a
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Grant SendOnBehalf To $($memberCheck.PrimarySmtpAddress_Destination)"

            #Set Attribute     
            try {
                $permResult = Set-DistributionGroup -Identity $user.PrimarySmtpAddress_Destination -GrantSendOnBehalfTo @{add=$memberCheck.PrimarySmtpAddress_Destination} -confirm:$false -ea Stop
                Write-Host ". " -ForegroundColor Green -NoNewline
            }
            catch {
                Write-Host ". " -ForegroundColor red -NoNewline

                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SendOnBehalfTo" -Force
                $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
                $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
    }

    #Stamp SendAs Perms on Group
    if ($user.SendAsPerms) {
        Write-Host "Adding SendAsPerms .. " -ForegroundColor DarkCyan -NoNewline
        #Gather Perm Users
        $membersSendAsArray = $user.SendAsPerms -split ","

        #Progress Bar 2
        $progressref2 = ($membersSendAsArray).count
        $progresscounter2 = 0
        foreach ($member in $membersSendAsArray) {
            #Member Check
            $memberCheck = @()
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}

            #Progress Bar 2a
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting SendAs to $($memberCheck.PrimarySmtpAddress_Destination)"

            #Add Perms to Group      
            try {
                $permResult = Add-RecipientPermission $user.PrimarySmtpAddress_Destination -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -confirm:$false -ea Stop
                Write-Host ". " -ForegroundColor Green -NoNewline
            }
            catch {
                Write-Host ". " -ForegroundColor red -NoNewline

                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SendAsPerms" -Force
                $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
                $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
    }
    
    Write-Host " done " -ForegroundColor Green
}

## Add Members to DistributionGroups for Abaco
$matchedDistributionGroups = import-csv
$matchedRecipients = import-csv

$failures = @()
$progressref = ($matchedDistributionGroups).count
$progresscounter = 0
foreach ($group in $matchedDistributionGroups) {
    $destinationGroupEmail = $group.PrimarySMTPAddress_Destination
    $destinationDisplayName = $group.DisplayName_Destination
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Group Members for Distribution Group $($destinationDisplayName)"

    Write-Host "Checking for $($destinationDisplayName)... " -NoNewline -ForegroundColor Cyan
    if ($DLCheck = Get-DistributionGroup $destinationGroupEmail -ea silentlycontinue) {
        $GroupMembers = $group.Members -split ","
        $progressref2 = ($GroupMembers).count
        $progresscounter2 = 0
        Write-Host "Adding $($GroupMembers.count) Members .. " -NoNewline
        $dlMembersCheck = (Get-DistributionGroupMember $destinationGroupEmail).primarysmtpaddress -join ","
        $dlMemberArray = $dlMembersCheck -split ","
        foreach ($member in $GroupMembers) {
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding Member $($trimMember)"
            #Check for User in Matched list
            $trimMember = $member.trim()
            if ($matchedUser = $matchedRecipients | Where-Object {$_.PrimarySmtpAddress_Source -eq $trimMember}) {  
                $matchedMemberDestinationAddress = $matchedUser.PrimarySMTPAddress_Destination
                #Check if member in Group Members
                if ($dlMemberArray -contains $matchedMemberDestinationAddress) {
                    Write-Host "." -ForegroundColor Yellow -NoNewline
                }
                else {
                    ## Check if member Exists
                    if ($recipientCheck = Get-Recipient $matchedMemberDestinationAddress -ea silentlycontinue -ResultSize 1) {
                        try {
                            #Add DL Members
                            Add-DistributionGroupMember $destinationGroupEmail -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                            Write-Host "." -ForegroundColor Green -NoNewline
                        }
                        catch {
                            Write-Host "." -ForegroundColor Red -NoNewline
                            
                            #Build Error Array
                            $currenterror = new-object PSObject

                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                            $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                            $currenterror | add-member -type noteproperty -name "DistributionList" -Value $destinationGroupEmail
                            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                            $failures += $currenterror
                        }
                    }
                }             
            }
            else {
                Write-Host "No Matched Recipient $($trimMember) found.. " -ForegroundColor red -NoNewline
                $notFoundPermUser += $trimMember
            }
        }
    }
    else {
        Write-Host "Group Is not Enabled for Exchange." -ForegroundColor Yellow -NoNewline
    }
    Write-Host "Done"
}

# Stamp Send As Perms - Groups
$matchedDistributionGroups = Import-Csv
$AllErrorsPerms = @()
$progressref = $matchedDistributionGroups.count
$progresscounter = 0
foreach ($object in $matchedDistributionGroups) {
    #Set Variables
    $sourceEmail = $object.PrimarySmtpAddress_Source
    $destinationEmail = $object.PrimarySMTPAddress_Destination
    $destinationDisplayName = $object.DisplayName_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Stamping Perms on Groups $($destinationDisplayName)"

    #Stamp Send As Perms for UserMailbox
    if ($object.SendAsPerms) {
        $SendAsPerms = $object.SendAs -split ","
        $SendAsPermsUsers = $SendAsPerms | ?{$_ -notlike "*NAMPR16A*" -and $_ -ne "noreply@abaco.com" -and $_ -ne "smtp@abaco.com"}
        #Only Run for Legitimate Users
        if ($SendAsPermsUsers) {
            Write-Host "Stamping Perms on $($destinationDisplayName).. " -ForegroundColor Cyan -NoNewline
            #Progress Bar 2
            $progressref2 = ($SendAsPermsUsers).count
            $progresscounter2 = 0
            foreach ($member in $SendAsPermsUsers) {
                #Member Check
                $memberCheck = @()
                $memberCheck = $matchedRecipients | ? {$_.PrimarySmtpAddress_Source -eq $member}

                #Progress Bar 2a
                $progresscounter2 += 1
                $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Send As to $($memberCheck.PrimarySmtpAddress_Destination)"

                #Add Perms to Mailbox      
                try {
                    $permResult = Add-RecipientPermission -identity $destinationEmail -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -ea Stop
                    Write-Host "." -ForegroundColor Green -NoNewline
                }
                catch {
                    Write-Host "." -ForegroundColor red -NoNewline

                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantSendAs" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "User" -Value $sourceEmail -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationEmail -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrorsPerms += $currenterror           
                    continue
                }
            }
            Write-Host " Completed " -ForegroundColor Green
        }      
    }
}

