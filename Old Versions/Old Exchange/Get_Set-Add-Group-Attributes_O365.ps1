#This Requires the usage of "Get-ExchangeGroupDetails.ps1" to gather the group details and add members. Must match up groups as well

function Write-ProgressHelper {
    param (
        [int]$ProgressCounter,
        [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [int]$TotalCount,
        [datetime]$StartTime
    )
    #$ProgressPreference = "Continue"  
    if ($ProgressPreference = "SilentlyContinue") {
        $ProgressPreference = "Continue"
    }

    $secondsElapsed = (Get-Date) - $StartTime
    $secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)
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
    #$secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)

}

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

function Get-ExchangeGroupDetails {
    param (
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file in CSV format? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file  in CSV format? Please provide full FILEPATH')] [string] $OutputCSVFilePath,
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file  in EXCEL (xlsx) format? Please provide full FILEPATH')] [string] $OutputExcelFilePath,
    [Parameter(Mandatory=$false,HelpMessage='What is the Excel Sheet name?')] [string] $WorkSheetName,
    [Parameter(Mandatory=$false,HelpMessage="Run against OnPremises Exchange?")] [switch]$OnPremises,
    [Parameter(Mandatory=$false,HelpMessage="Run against Office365 Exchange Online?")] [switch]$Office365
    )

    $allMailGroups = Get-Recipient -RecipientTypeDetails group -ResultSize unlimited -EA SilentlyContinue
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailNonUniversalGroup -ResultSize unlimited -EA SilentlyContinue
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize unlimited -EA SilentlyContinue
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize unlimited -EA SilentlyContinue  
    $allMailGroups += Get-Recipient -RecipientTypeDetails DynamicDistributionGroup -ResultSize unlimited -EA SilentlyContinue
    $allMailGroupDetails = @()

    #ProgressBar
    $progresscounter = 1
    $notmatchedUsers = @()
    $start = Get-Date
    $totalCount = $distributionGroupMembers.count

    foreach ($object in $allMailGroups | sort DisplayName) {
        
        $identity = $object.identity.tostring()
        $PrimarySMTPAddress = $object.PrimarySMTPAddress.ToString()
        Write-ProgressHelper -Activity "Gathering Group Details for $($PrimarySMTPAddress)" -ProgressCounter ($progresscounter++) -TotalCount ($allMailGroups).count -StartTime $global:start

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
    $allMailGroupDetails | Export-Csv "$HOME\Desktop\AllMailGroupDetails.csv" -NoTypeInformation -Encoding UTF8
    Write-host "Exported 'AllMailGroupDetails.csv' List to $HOME\Desktop" -ForegroundColor Cyan

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
    Write-Host "Completed in "(((Get-Date) - $start).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
}


function Add-DistributionGroupMembers {
    param (
        [Parameter(Mandatory=$True)] [string]$DestinationTenant,
        [Parameter(Mandatory=$True)] [string]$SourceTenant,
        [Parameter(Mandatory=$True)] [Hash]$mailboxMapping,
        [Parameter(Mandatory=$True)] [Array]$matchedGroupsDetails,
        [Parameter(Mandatory=$false)] [switch]$test,
        [Parameter(Mandatory=$false)] [Array]$distributionGroupMembers
    )
    $AllGroupErrors = @()
    $progresscounter = 1
    $notmatchedUsers = @()
    $addedMembers = 0
    $alreadyMembers = 0
    $failedAddMembers = 0
    $start = Get-Date
    $totalCount = $matchedGroupsDetails.count

    foreach ($object in $matchedGroupsDetails) {
        Write-ProgressHelper -ID 1 -Activity "Adding Members: $($object."DisplayName$($DestinationTenant)")" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -StartTime $start
        if ($object.RecipientTypeDetails -eq "MailUniversalDistributionGroup" -or $object.RecipientTypeDetails -eq "MailUniversalSecurityGroup" -or $object.RecipientTypeDetails -eq "MailNonUniversalGroup") {
            Write-Host "Adding Members to $($object."DisplayName$($DestinationTenant)") ... " -ForegroundColor Cyan -NoNewline
            $groupMembers = $object."Members" -split ","
            $destinationGroup = $object."PrimarySMTPAddress$($DestinationTenant)"

            $totalCount2 = $groupMembers.count
            $progresscounter2 = 1
            $start2 = Get-Date
            foreach ($member in $groupMembers) {
                Write-ProgressHelper -ID 2 -Activity "Adding Member to $($member)" -ProgressCounter ($progresscounter2++) -TotalCount $TotalCount2 -StartTime $start2
                $memberCheck = @()

                if ($memberCheck = $matchedMailboxes | ? {$_."UserPrincipalName$($SourceTenant)" -eq $member}) {
                    #Member Check
                    $memberMatchedAddress = ($memberCheck."UserPrincipalName$($DestinationTenant)").toString()
        
                    #Add Member to Distribution Group        
                    try {
                        #Write-Host "Adding Member $($memberMatchedAddress) to $($destinationGroup) ... " -ForegroundColor Cyan -NoNewline
                        if ($test) {
                            Add-DistributionGroupMember -Identity $destinationGroup -Member $memberMatchedAddress -ea Stop -whatif
                        }
                        else {
                            Add-DistributionGroupMember -Identity $destinationGroup -Member $memberMatchedAddress -ea Stop
                        }
                        Write-Host ". " -ForegroundColor Green -NoNewline
                        $addedMembers++
                    } catch {
                        if ($_.Exception.Message -like "*is already a member of the group*") {
                            Write-Host "." -ForegroundColor Yellow -NoNewline
                            $alreadyMembers++
                        }
                        else {
                            Write-Host "." -ForegroundColor Red -NoNewline
                            $failedAddMembers++

                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationGroup -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberMatchedAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllGroupErrors += $currenterror
                        continue
                        }
                        
                    }
                }
                else {
                    Write-Host "." -ForegroundColor Red -NoNewline
                    $notmatchedUsers += $object."Member-PrimarySMTPAddress"
                }

            }
            Write-Host "done" -ForegroundColor Green 
        }
        else {
            Write-Host "Skipping $($object."DisplayName$($DestinationTenant)"): $($object.RecipientTypeDetails)" -ForegroundColor Yellow
        }        
    }

    Write-Output ""
    Write-Host "Added $($addedMembers) Members" -ForegroundColor Green
    Write-Host "Skipped $($alreadyMembers) Members" -ForegroundColor Yellow
    Write-Host "Failed to Add $($failedAddMembers) Members" -ForegroundColor Red
        
}

function Add-DistributionGroupAttributes {
    param (
        [Parameter(Mandatory=$True)] [Array]$matchedMembers,
        [Parameter(Mandatory=$True)] [Array]$matchedGroupsDetails,
        [Parameter(Mandatory=$True)] [string]$DestinationTenant,
        [Parameter(Mandatory=$True)] [string]$SourceTenant,
        [Parameter(Mandatory=$false)] [switch]$test,
        [Parameter(Mandatory=$false)] [Array]$distributionGroupMembers
    )

    $progresscounter = 1
    $AllGroupErrors = @()
    $notmatchedUsers = @()
    $notFoundGroups = @()
    $addedOwners = 0
    $failedAddOwners = 0
    $addModerators = 0
    $failedAddModerators = 0
    $addGrantSendOnBehalfTo = 0
    $failedAddGrantSendOnBehalfTo = 0
    $addSendAs = 0
    $failedAddSendAs = 0
    $addAcceptMessagesOnlyFrom = 0
    $failedAddAcceptMessagesOnlyFrom = 0
    $addRejectMessagesOnlyFrom = 0
    $failedAddRejectMessagesOnlyFrom = 0
    $RemoveMigrationOwner = 0
    $failedRemoveMigrationOwner = 0

    $start = Get-Date
    $totalCount = $matchedGroupsDetails.count

    foreach ($object in $matchedGroupsDetails) {
        Write-ProgressHelper -ID 1 -Activity "Updating Group: $($object."DisplayName_$($DestinationTenant)")" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -StartTime $start
        if ($object.RecipientTypeDetails -eq "MailUniversalDistributionGroup" -or $object.RecipientTypeDetails -eq "MailUniversalSecurityGroup" -or $object.RecipientTypeDetails -eq "MailNonUniversalGroup") {
            Write-Host "Updating $($object.RecipientTypeDetails): $($object."DisplayName_$($DestinationTenant)") ... " -ForegroundColor Cyan -NoNewline
            $destinationGroup = $object."PrimarySMTPAddress_$($DestinationTenant)"
            if ($recipientCheck = Get-ExoRecipient $destinationGroup -ea silentlycontinue) {
                ####Set Require Sender Authentication
                Set-DistributionGroup -Identity $destinationGroup -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($object.RequireSenderAuthenticationEnabled)) -ErrorAction SilentlyContinue -WarningAction silentlycontinue

                ###Add Owners
                if ($object.Owners) {
                    Write-Host "Add Owners.. " -ForegroundColor Yellow -NoNewline
                    $groupOwners = $object."Owners" -split ","
                    $totalCount2 = $groupOwners.count
                    $progresscounter2 = 1
                    $start2 = Get-Date
                    foreach ($member in $groupOwners) {
                        Write-ProgressHelper -ID 2 -Activity "Adding Owner: $($member)" -ProgressCounter ($progresscounter2++) -TotalCount $TotalCount2 -StartTime $start2
                        $memberCheck = @()

                        if ($memberCheck = $matchedMailboxes | ? {$_."UserPrincipalName$($SourceTenant)" -eq $member}) {
                            #Owner Check
                            $memberMatchedAddress = ($memberCheck."UserPrincipalName$($DestinationTenant)").toString()
                
                            #Add Owner to Distribution Group        
                            try {
                                if ($test) {
                                    Set-DistributionGroup -whatif -Identity $destinationGroup -ManagedBy @{add=$memberMatchedAddress} -ea Stop 
                                }
                                else {
                                    Set-DistributionGroup -Identity $destinationGroup -ManagedBy @{add=$memberMatchedAddress} -ea Stop
                                }
                                Write-Host "." -ForegroundColor Green -NoNewline
                                $addedOwners++
                            } catch {
                                Write-Host "." -ForegroundColor Red -NoNewline
                                $failedAddOwners++

                                $currenterror = new-object PSObject
                                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddOwner" -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationGroup -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberMatchedAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                                $AllGroupErrors += $currenterror
                                $currentError
                                continue
                            }
                        }
                        else {
                            Write-Host "." -ForegroundColor DarkYellow -NoNewline
                            $notmatchedUsers += $object."Member-PrimarySMTPAddress"
                        }

                    }
                }

                #Add ModeratedBy
                if ($object.ModeratedBy) {
                    Write-Host "Add ModeratedBy.. " -ForegroundColor Yellow -NoNewline
                    $groupModeratedBy = $object."ModeratedBy" -split ","
                    $totalCount2 = $groupModeratedBy.count
                    $progresscounter2 = 1
                    $start2 = Get-Date
                    foreach ($member in $groupModeratedBy) {
                        Write-ProgressHelper -ID 2 -Activity "Adding Moderator: $($member)" -ProgressCounter ($progresscounter2++) -TotalCount $TotalCount2 -StartTime $start2
                        $memberCheck = @()

                        if ($memberCheck = $matchedMailboxes | ? {$_."UserPrincipalName$($SourceTenant)" -eq $member}) {
                            #Member Check
                            $memberMatchedAddress = ($memberCheck."UserPrincipalName$($DestinationTenant)").toString()
                
                            #Add Member to Distribution Group        
                            try {
                                if ($test) {
                                    Set-DistributionGroup -whatif -Identity $destinationGroup -ModeratedBy @{add=$memberMatchedAddress} -ea Stop -WarningAction SilentlyContinue
                                }
                                else {
                                    Set-DistributionGroup -Identity $destinationGroup -ModeratedBy @{add=$memberMatchedAddress} -ea Stop -WarningAction SilentlyContinue
                                }
                                Write-Host ". " -ForegroundColor Green -NoNewline
                                $addModerators++
                            } catch {
                                Write-Host "." -ForegroundColor Red -NoNewline
                                $failedAddModerators++

                                $currenterror = new-object PSObject
                                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddModeratedBy" -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationGroup -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberMatchedAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                                $AllGroupErrors += $currenterror           
                                continue
                            }
                        }
                        else {
                            if (-not $memberCheck) {
                                Write-Host "." -ForegroundColor DarkYellow -NoNewline
                                $notmatchedUsers += $object."Member-PrimarySMTPAddress"
                            }
                        }

                    }
                }

                #Add SendOnBehalfto
                if ($object.GrantSendOnBehalfTo) {
                    Write-Host "Add GrantSendOnBehalfTo.. " -ForegroundColor Yellow -NoNewline
                    $groupGrantSendOnBehalfTo = $object."GrantSendOnBehalfTo" -split ","
                    $totalCount2 = $groupGrantSendOnBehalfTo.count
                    $progresscounter2 = 1
                    $start2 = Get-Date
                    foreach ($member in $groupGrantSendOnBehalfTo) {
                        Write-ProgressHelper -ID 2 -Activity "Adding SendOnBehalfTo: $($member)" -ProgressCounter ($progresscounter2++) -TotalCount $TotalCount2 -StartTime $start2
                        $memberCheck = @()

                        if ($memberCheck = $matchedMailboxes | ? {$_."UserPrincipalName$($SourceTenant)" -eq $member}) {
                            #Member Check
                            $memberMatchedAddress = ($memberCheck."UserPrincipalName$($DestinationTenant)").toString()
                
                            #Add Member to Distribution Group        
                            try {
                                if ($test) {
                                    Set-DistributionGroup -whatif -Identity $destinationGroup -GrantSendOnBehalfTo @{add=$memberMatchedAddress} -ea Stop -WarningAction SilentlyContinue
                                }
                                else {
                                    Set-DistributionGroup -Identity $destinationGroup -GrantSendOnBehalfTo @{add=$memberMatchedAddress} -ea Stop -WarningAction SilentlyContinue
                                }
                                Write-Host ". " -ForegroundColor Green -NoNewline
                                $addGrantSendOnBehalfTo++
                            } catch {
                                Write-Host "." -ForegroundColor Red -NoNewline
                                $failedAddGrantSendOnBehalfTo++

                                $currenterror = new-object PSObject
                                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantSendOnBehalfTo" -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationGroup -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberMatchedAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                                $AllGroupErrors += $currenterror           
                                continue
                            }
                        }
                        else {
                            if (-not $memberCheck) {
                                Write-Host "." -ForegroundColor DarkYellow -NoNewline
                                $notmatchedUsers += $object."Member-PrimarySMTPAddress"
                            }
                        }

                    }
                }

                #Add SendAs
                if ($object.SendAs) {
                    Write-Host "Add SendAs.. " -ForegroundColor Yellow -NoNewline
                    $groupSendAs = $object."SendAs" -split ","
                    $totalCount2 = $groupSendAs.count
                    $progresscounter2 = 1
                    $start2 = Get-Date
                    foreach ($member in $groupSendAs) {
                        Write-ProgressHelper -ID 2 -Activity "Adding SendAs: $($member)" -ProgressCounter ($progresscounter2++) -TotalCount $TotalCount2 -StartTime $start2
                        $memberCheck = @()

                        if ($memberCheck = $matchedMailboxes | ? {$_."UserPrincipalName$($SourceTenant)" -eq $member}) {
                            #Member Check
                            $memberMatchedAddress = ($memberCheck."UserPrincipalName$($DestinationTenant)").toString()
                
                            #Add Member to Distribution Group        
                            try {
                                if ($test) {
                                    Add-RecipientPermission -whatif -Identity $destinationGroup -Trustee $memberMatchedAddress -AccessRights SendAs -ea Stop -WarningAction SilentlyContinue
                                }
                                else {
                                    Add-RecipientPermission -Identity $destinationGroup -Trustee $memberMatchedAddress -AccessRights SendAs -ea Stop -WarningAction SilentlyContinue
                                }
                                Write-Host ". " -ForegroundColor Green -NoNewline
                                $addSendAs++
                            } catch {
                                Write-Host "." -ForegroundColor Red -NoNewline
                                $failedAddSendAs++

                                $currenterror = new-object PSObject
                                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantSendAs" -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationGroup -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberMatchedAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                                $AllGroupErrors += $currenterror           
                                continue
                            }
                        }
                        else {
                            if (-not $memberCheck) {
                                Write-Host "." -ForegroundColor DarkYellow -NoNewline
                                $notmatchedUsers += $object."Member-PrimarySMTPAddress"
                            }
                        }

                    }
                }

                #Add AcceptMessagesOnlyFrom
                if ($object.AcceptMessagesOnlyFrom) {
                    Write-Host "Add AcceptMessagesOnlyFrom.. " -ForegroundColor Yellow -NoNewline
                    $groupAcceptMessagesOnlyFrom = $object."AcceptMessagesOnlyFrom" -split ","
                    $totalCount2 = $groupAcceptMessagesOnlyFrom.count
                    $progresscounter2 = 1
                    $start2 = Get-Date
                    foreach ($member in $groupAcceptMessagesOnlyFrom) {
                        Write-ProgressHelper -ID 2 -Activity "Adding AcceptMessagesOnlyFrom: $($member)" -ProgressCounter ($progresscounter2++) -TotalCount $TotalCount2 -StartTime $start2
                        $memberCheck = @()

                        if ($memberCheck = $matchedMailboxes | ? {$_."UserPrincipalName$($SourceTenant)" -eq $member}) {
                            #Member Check
                            $memberMatchedAddress = ($memberCheck."UserPrincipalName$($DestinationTenant)").toString()
                
                            #Add Member to Distribution Group        
                            try {
                                if ($test) {
                                    Set-DistributionGroup -whatif -Identity $destinationGroup -AcceptMessagesOnlyFrom @{add=$memberMatchedAddress} -ea Stop -WarningAction SilentlyContinue
                                }
                                else {
                                    Set-DistributionGroup -Identity $destinationGroup -AcceptMessagesOnlyFrom @{add=$memberMatchedAddress} -ea Stop -WarningAction SilentlyContinue
                                }
                                Write-Host ". " -ForegroundColor Green -NoNewline
                                $addAcceptMessagesOnlyFrom++
                            } catch {
                                Write-Host "." -ForegroundColor Red -NoNewline
                                $failedAddAcceptMessagesOnlyFrom++

                                $currenterror = new-object PSObject
                                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddAcceptMessagesOnlyFrom" -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationGroup -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberMatchedAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                                $AllGroupErrors += $currenterror           
                                continue
                            }
                        }
                        else {
                            if (-not $memberCheck) {
                                Write-Host "." -ForegroundColor DarkYellow -NoNewline
                                $notmatchedUsers += $object."Member-PrimarySMTPAddress"
                            }
                        }

                    }
                }

                #Add RejectMessagesOnlyFrom
                if ($object.RejectMessagesOnlyFrom) {
                    Write-Host "Add RejectMessagesOnlyFrom.. " -ForegroundColor Yellow -NoNewline
                    $RejectMessagesOnlyFrom = $object."RejectMessagesOnlyFrom" -split ","
                    $totalCount2 = $RejectMessagesOnlyFrom.count
                    $progresscounter2 = 1
                    $start2 = Get-Date
                    foreach ($member in $RejectMessagesOnlyFrom) {
                        Write-ProgressHelper -ID 2 -Activity "Adding RejectMessagesOnlyFrom: $($member)" -ProgressCounter ($progresscounter2++) -TotalCount $TotalCount2 -StartTime $start2
                        $memberCheck = @()

                        if ($memberCheck = $matchedMailboxes | ? {$_."UserPrincipalName$($SourceTenant)" -eq $member}) {
                            #Member Check
                            $memberMatchedAddress = ($memberCheck."UserPrincipalName$($DestinationTenant)").toString()
                
                            #Add Member to Distribution Group        
                            try {
                                if ($test) {
                                    Set-DistributionGroup -whatif -Identity $destinationGroup -RejectMessagesOnlyFrom @{add=$memberMatchedAddress} -ea Stop -WarningAction SilentlyContinue
                                }
                                else {
                                    Set-DistributionGroup -Identity $destinationGroup -RejectMessagesOnlyFrom @{add=$memberMatchedAddress} -ea Stop -WarningAction SilentlyContinue
                                }
                                Write-Host ". " -ForegroundColor Green -NoNewline
                                $addRejectMessagesOnlyFrom++
                            } catch {
                                Write-Host "." -ForegroundColor Red -NoNewline
                                $failedAddRejectMessagesOnlyFrom++

                                $currenterror = new-object PSObject
                                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddRejectMessagesOnlyFrom" -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationGroup -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberMatchedAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                                $AllGroupErrors += $currenterror           
                                continue
                            }
                        }
                        else {
                            if (-not $memberCheck) {
                                Write-Host "." -ForegroundColor DarkYellow -NoNewline
                                $notmatchedUsers += $object."Member-PrimarySMTPAddress"
                            }
                        }

                    }
                }
                
                ####Remove Migration Users
                
                Write-Host "Remove Migration Account Owner.. " -ForegroundColor DarkYellow -NoNewline
                try {
                    if ($test) {
                        Set-DistributionGroup -whatif -Identity $destinationGroup -ManagedBy @{remove=$spectramig} -ea Stop -WarningAction silentlycontinue
                    }
                    else {
                        Set-DistributionGroup -Identity $destinationGroup -ManagedBy @{remove=$spectramig} -ea Stop -WarningAction silentlycontinue
                    }
                    Write-Host ". " -ForegroundColor Green -NoNewline
                    $RemoveMigrationOwner++
                } catch {
                    Write-Host "." -ForegroundColor Yellow -NoNewline
                    $failedRemoveMigrationOwner++

                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationGroup -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberMatchedAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllGroupErrors += $currenterror           
                    continue
                }
                Write-Host "done" -ForegroundColor Green
            }
            else {
                $notFoundGroups += $object
               Write-Host "Skipping $($object."DisplayName"): $($object.RecipientTypeDetails)" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Skipping $($object."DisplayName"): $($object.RecipientTypeDetails)" -ForegroundColor Yellow
        }        
    }

    Write-Output ""
    Write-Host "Unable to Find $($notFoundGroups.count) Groups" -ForegroundColor Yellow
    Write-Host "Unable to Find $($notmatchedUsers.count) Users" -ForegroundColor Yellow
    Write-Output ""
    Write-Host "Added $($addedOwners) Owners" -ForegroundColor Green
    Write-Host "Failed to Add $($failedAddOwners) Owners" -ForegroundColor Red
    Write-Host "Added $($addModerators) Moderators" -ForegroundColor Green
    Write-Host "Failed to Add $($failedAddModerators) Moderators" -ForegroundColor Red
    Write-Host "Added $($addGrantSendOnBehalfTo) GrantSendOnBehalfTo" -ForegroundColor Green
    Write-Host "Failed to Add $($failedAddGrantSendOnBehalfTo) GrantSendOnBehalfTo" -ForegroundColor Red
    Write-Host "Added $($addSendAs) SendAs" -ForegroundColor Green
    Write-Host "Failed to Add $($failedAddSendAs) SendAs" -ForegroundColor Red
    Write-Host "Added $($addAcceptMessagesOnlyFrom) AcceptMessagesOnlyFrom" -ForegroundColor Green
    Write-Host "Failed to Add $($failedAddAcceptMessagesOnlyFrom) AcceptMessagesOnlyFrom" -ForegroundColor Red
    Write-Host "Added $($addRejectMessagesOnlyFrom) RejectMessagesOnlyFrom" -ForegroundColor Green
    Write-Host "Failed to Add $($failedAddRejectMessagesOnlyFrom) RejectMessagesOnlyFrom" -ForegroundColor Red
    Write-Host "Removed $($RemoveMigrationOwner) Migration Owner" -ForegroundColor Green
    Write-Host "Failed to Remove $($failedRemoveMigrationOwner) Migration Owner" -ForegroundColor Red

      
}



$distributionGroupMembers = Import-CSV -Path #Path to file
$matchedGroupsDetails = Import-Excel -WorksheetName "AllMailGroupDetails" -Path #Path to file
$matchedMailboxes = Import-Excel -WorksheetName "MatchedMailboxes" -Path #Path to file

Get-ExchangeGroupDetails
Add-DistributionGroupMembers -SourceTenant Spectra -DestinationTenant OVG -matchedMembers $matchedMailboxes -matchedGroupsDetails $matchedGroupsDetails
Add-DistributionGroupAttributes -SourceTenant Spectra -DestinationTenant OVG  -matchedMembers $matchedMailboxes -matchedGroupsDetails $matchedGroupsDetails
```
