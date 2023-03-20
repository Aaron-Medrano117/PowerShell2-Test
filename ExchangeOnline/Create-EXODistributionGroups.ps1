## Create EXO DistributionGroups
$groupDetails = import-csv
$AllErrors = @()
$progressref = ($groupDetails).count
$progresscounter = 0
foreach ($group in $groupDetails) {
   $destinationEmail = $group.PrimarySMTPAddress
   $destinationDisplayName = $group.DisplayName
   $progresscounter += 1
   $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
   $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
   Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Distribution Group $($destinationDisplayName)"

    Write-Host "Checking for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    if (!($DLCheck = Get-DistributionGroup $destinationEmail -ea silentlycontinue)) {
        try {
            New-DistributionGroup -DisplayName $destinationDisplayName -name $destinationDisplayName -PrimarySMTPAddress $destinationEmail -ErrorAction Stop
            Write-Host "New Group Created" -ForegroundColor Green
            $createdADGroups += $group
        }
        catch {
            Write-Host ". " -ForegroundColor red -NoNewline
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToCreateGroup" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Object" -Value $destinationDisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $group.RecipientTypeDetails -Force
            $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $destinationEmail -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $AllErrors += $currenterror           
            continue
        }
    }
}

## Set Attributes DistributionGroups - Add X500, HiddenFromAddressList, RequireSenderAuthentication, Add Email Aliases, Add Approved and Rejected Senders
$groupDetails = import-csv
$progressref = ($groupDetails).count
$progresscounter = 0
foreach ($group in $groupDetails) {
    $destinationEmail = $group.PrimarySMTPAddress
    $destinationDisplayName = $group.DisplayName
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Set Distribution Group Settings for $($destinationDisplayName)"

    Write-Host "Checking for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    if ($group.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
        if ($DLCheck = Get-DistributionGroup $destinationEmail -ea silentlycontinue) {
            #HiddenFromAddressList & RequireSenderAuth - Boolean
            [boolean]$HiddenFromAddressListsEnabled = [boolean]::Parse($group.HiddenFromAddressListsEnabled)
            [boolean]$RequireSenderAuthenticationEnabled = [boolean]::Parse($group.RequireSenderAuthenticationEnabled)         
            Set-DistributionGroup $adGroupCheck.DistinguishedName -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -MemberJoinRestriction $group.MemberJoinRestriction -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -warningaction silentlycontinue   
            
            #add Email Addresses - X500
            $X500 = $group.LegacyExchangeDN
            Set-DistributionGroup -EmailAddresses @{add=$X500} -Identity $destinationEmail -warningaction silentlycontinue
            #add Email Addresses - All
            $emailAddresses = $group.EmailAddresses -split ","
            foreach ($address in $emailAddresses) {
                if (!($null -eq $address)) {
                    try {
                        Set-DistributionGroup  -EmailAddresses @{add=$address} -Identity $destinationEmail -warningaction silentlycontinue
                        Write-Host ". " -ForegroundColor Cyan -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdateGroup-Addresses" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Object" -Value $destinationDisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $group.RecipientTypeDetails -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $destinationEmail -Force
                        $currenterror | Add-Member -type NoteProperty -Name "SecondaryObject" -Value $address -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
            }

            #AcceptMessagesOnlyFrom
            $AcceptMessagesOnlyFrom = $group.AcceptMessagesOnlyFrom -split ","
            foreach ($user in $AcceptMessagesOnlyFrom) {
                if (!($null -eq $user)) {
                    try {
                        Set-DistributionGroup -AcceptMessagesOnlyFrom @{add=$user} -Identity $destinationEmail -warningaction silentlycontinue
                        Write-Host ". " -ForegroundColor Cyan -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdateGroup-AcceptMessagesOnlyFrom" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Object" -Value $destinationDisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $group.RecipientTypeDetails -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $destinationEmail -Force
                        $currenterror | Add-Member -type NoteProperty -Name "SecondaryObject" -Value $user -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
            }

            #RejectMessagesOnlyFrom
            $emailAddresses = $group.RejectMessagesOnlyFrom -split ","
            foreach ($user in $emailAddresses) {
                if (!($null -eq $user)) {
                    try {
                        Set-DistributionGroup -RejectMessagesOnlyFrom @{add=$user} -Identity $destinationEmail -warningaction silentlycontinue
                        Write-Host ". " -ForegroundColor Cyan -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdateGroupAddresses" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Object" -Value $destinationDisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $group.RecipientTypeDetails -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $destinationEmail -Force
                        $currenterror | Add-Member -type NoteProperty -Name "SecondaryObject" -Value $user -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
            }
        }
    }
}

## Add Members to DistributionGroups
$groupDetails = import-csv
$createdADGroups = @()
$progressref = ($groupDetails).count
$progresscounter = 0
foreach ($group in $groupDetails) {
    $destinationEmail = $group.PrimarySMTPAddress
    $destinationDisplayName = $group.DisplayName
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Group Members for Distribution Group $($destinationDisplayName)"

    Write-Host "Checking for $($destinationDisplayName)... " -NoNewline -ForegroundColor Cyan
    if ($DLCheck = Get-DistributionGroup $destinationEmail -ea silentlycontinue) {
        $GroupMembers = $group.Members -split ","
        $progressref2 = ($GroupMembers).count
        $progresscounter2 = 0
        Write-Host "Adding $($GroupMembers.count) Members.. " -NoNewline
        foreach ($member in $GroupMembers) {
            if (!($null -eq $member)) {
                # Match the Perm user
                $trimMember = $member.trim()
                $progresscounter2 += 1
                $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding member $($trimMember)"

                if ($recipientCheck = Get-Recipient $trimMember -ea silentlycontinue) {
                    try {
                        #Add DL Members
                        Add-DistributionGroupMember $destinationDisplayName -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                        Write-Host "." -ForegroundColor Green -NoNewline
                    }
                    catch {
                        #Build Error Array
                        Write-Host ". " -ForegroundColor red -NoNewline
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddGroupMember" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Object" -Value $destinationDisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $group.RecipientTypeDetails -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $destinationEmail -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
            }
        }
    }
    else {
        Write-Host "Group Is not Enabled for Exchange. " -ForegroundColor Yellow
    }
    Write-Host "Done"
}

#Create New Office365 Groups
$AllErrorsGroups = @()
$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($object in $O365Groups) {
    #Set Variables
    $Addresssplit = $object.PrimarySMTPAddress_Source -split "@"
    $DestinationPrimarySMTPAddress ="abaco." + $addressSplit[0] + "@ametek.com"
    $destinationDisplayName = "Abaco-" + $object.DisplayName_Source

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Office 365 Group $($DestinationPrimarySMTPAddress)"

    $newGroup = @()
	Write-Host "$($DestinationPrimarySMTPAddress) ... " -ForegroundColor Cyan -NoNewline
	
    #Check for Group, Create if it does not exist
	if (!($groupCheck = Get-UnifiedGroup $DestinationPrimarySMTPAddress -EA SilentlyContinue)) {
		Write-Host "creating ... " -ForegroundColor Yellow -NoNewline

        #Create Group
		try {
            $newGroup = New-UnifiedGroup -Name $destinationDisplayName -DisplayName $destinationDisplayName -PrimarySmtpAddress $DestinationPrimarySMTPAddress -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($object.RequireSenderAuthenticationEnabled)) -Confirm:$false
        }
        catch {
            $currenterror = new-object PSObject

            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $destinationDisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToCreate" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			Write-Host "failed to update" -ForegroundColor Red
            $AllErrorsGroups += $currenterror
			continue
        }

        #Set Access Type and notes
        try {
			$newGroup | Set-UnifiedGroup -AccessType $object.AccessType -Confirm:$false
            $newGroup | Set-UnifiedGroup -Notes $object.Notes -Confirm:$false -ErrorAction Stop
		}
		catch {
            $currenterror = new-object PSObject

            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $destinationDisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdate" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			Write-Host "failed to update" -ForegroundColor Red
            $AllErrorsGroups += $currenterror
			continue
		}
		Write-Host "done" -ForegroundColor Green
	}
}

# Add Members and Owners to Groups - Abaco
$allMatchedGroups = Import-Csv
$matchedMailboxes = Import-Csv
$allMatchedGroups = $allMatchedGroups | ?{$_.ExistsInDestination -eq $true}

$progressref = ($allMatchedGroups).count
$progresscounter = 0
$AllErrors_Groups = @()
foreach ($group in $allMatchedGroups) {
    #Set Variables
    $DestinationPrimarySMTPAddress = $group.PrimarySmtpAddress_Destination
    $destinationDisplayName = $group.DisplayName_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Members and Owners to $($destinationDisplayName)"
    Write-Host "Updating Group $($destinationDisplayName).. " -ForegroundColor Cyan -NoNewline

    if ($recipientCheck = Get-Recipient $DestinationPrimarySMTPAddress) {
            #Add Office365 Members
        if ($recipientCheck.RecipientTypeDetails -eq "GroupMailbox") {
            if ($group.Members) {
                $membersArray = $group.Members -split ","
                $membersArray = $membersArray | ?{$_ -ne $sourceMigrationAccount}
                Write-Host "Adding $($membersArray.count) Members .. " -ForegroundColor Cyan -NoNewline
                $membersCheck =  Get-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Member
                $ownersCheck = Get-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Owner
                
                #add Migration Account as Member for migration
                Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Member -Links $destinationMigrationAccount -ea silentlycontinue

                #Progress Bar 2
                $progressref2 = ($membersArray).count
                $progresscounter2 = 0
                foreach ($member in $membersArray) {
                    #Member Check
                    $memberAddress = @()
                    if ($member -like "*$migratingDomain") {
                        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}
                        $memberAddress = $memberCheck.PrimarySmtpAddress_Destination
                    }
                    else {
                        $memberAddress = $member
                    }
        
                    #Progress Bar 2a
                    $progresscounter2 += 1
                    $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                    $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                    Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding Member $($memberAddress)"
        
                    #Add Member to Office365 Group        
                    try {
                        if ($membersCheck | ?{$_.PrimarySMTPAddress -like $memberAddress}) {
                            Write-Host "." -ForegroundColor Yellow -NoNewline
                        }
                        else {
                            Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Member -Links $memberAddress -ea Stop
                            Write-Host "." -ForegroundColor Green -NoNewline
                        }
                        
                    }
                    catch {
                        Write-Host "." -ForegroundColor red -NoNewline
        
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $destinationDisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "GroupDestination_PrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors_Groups += $currenterror           
                        continue
                    }
                }
            } 
        
            if ($group.GroupOwners) {
                $membersArray = $group.GroupOwners -split ","
                $membersArray = $membersArray | ?{$_ -ne $sourceMigrationAccount}
                Write-Host "Adding $($membersArray.count) Owners .. " -ForegroundColor Cyan -NoNewline
                #add Migration Account as Owner for migration
                Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Owner -Links $destinationMigrationAccount -ea silentlycontinue
        
                #Progress Bar 2
                $progressref2 = ($membersArray).count
                $progresscounter2 = 0
                foreach ($member in $membersArray) {
                    #Member Check
                    $memberCheck = @()
                    $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}
        
                    #Progress Bar 2a
                    $progresscounter2 += 1
                    $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                    $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                    Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding Owner $($memberCheck.PrimarySmtpAddress_Destination)"
        
                    #Add Member to Office365 Group        
                    try {
                        if ($ownersCheck | ?{$_.PrimarySMTPAddress -like $memberCheck.PrimarySmtpAddress_Destination}) {
                            Write-Host "." -ForegroundColor Yellow -NoNewline
                        }
                        else {
                            Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Owner -Links $memberCheck.PrimarySmtpAddress_Destination -ea Stop
                            Write-Host "." -ForegroundColor Green -NoNewline
                        }
                        
                    }
                    catch {
                        Write-Host "." -ForegroundColor red -NoNewline
        
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddOwner" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $destinationDisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Group_Destination_PrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Owner_Source_PrimarySMTPAddress" -Value $member -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Owner_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors_Groups += $currenterror           
                        continue
                    }
                    
                }
            }
        }
        elseif ($recipientCheck.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
            if ($group.Members) {
                $membersArray = $group.Members -split ","
                $membersArray = $membersArray | ?{$_ -ne $MigrationAccount}
                Write-Host "Adding $($membersArray.count) Members .. " -ForegroundColor Cyan -NoNewline
                #Progress Bar 2
                $progressref2 = ($membersArray).count
                $progresscounter2 = 0
                $membersCheck = Get-DistributionGroupMember $DestinationPrimarySMTPAddress
                foreach ($member in $membersArray) {
                    #Member Check
                    $memberAddress = @()
                    if ($member -like "*$migratingDomain") {
                        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}
                        $memberAddress = $memberCheck.PrimarySmtpAddress_Destination
                    }
                    else {
                        $memberAddress = $member
                    }
        
                    #Progress Bar 2a
                    $progresscounter2 += 1
                    $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                    $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                    Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding Member $($memberAddress)"
        
                    #Add Member to Office365 Group        
                    try {
                        if ($membersCheck | ?{$_.PrimarySMTPAddress -like $memberAddress}) {
                            Write-Host "." -ForegroundColor Yellow -NoNewline
                        }
                        else {
                            Add-DistributionGroupMember -Identity $DestinationPrimarySMTPAddress -Member $memberAddress -ea Stop
                            Write-Host "." -ForegroundColor Green -NoNewline 
                        }
                        
                    }
                    catch {
                        Write-Host "." -ForegroundColor red -NoNewline
        
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $destinationDisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "GroupDestination_PrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors_Groups += $currenterror           
                        continue
                    }
                }
            } 
        
            if ($group.GroupOwners) {
                $membersArray = $group.GroupOwners -split ","
                $membersArray = $membersArray | ?{$_ -ne $MigrationAccount}
                Write-Host "Adding $($membersArray.count) Owners .. " -ForegroundColor Cyan -NoNewline
                #Progress Bar 2
                $progressref2 = ($membersArray).count
                $progresscounter2 = 0
                foreach ($member in $membersArray) {
                    #Member Check
                    $memberCheck = @()
                    $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}
        
                    #Progress Bar 2a
                    $progresscounter2 += 1
                    $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                    $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                    Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding Owner $($memberCheck.PrimarySmtpAddress_Destination)"
        
                    #Add Member to Office365 Group        
                    try {
                        if ($ownersCheck | ?{$_.PrimarySMTPAddress -like $memberCheck.PrimarySmtpAddress_Destination}) {
                            Write-Host "." -ForegroundColor Yellow -NoNewline
                        }
                        else {
                            $owner = $memberCheck.PrimarySmtpAddress_Destination
                            Set-DistributionGroup -Identity $DestinationPrimarySMTPAddress -ManagedBy @{add=$owner} -ea Stop
                            Write-Host "." -ForegroundColor Green -NoNewline
                        }
                    }
                    catch {
                        Write-Host "." -ForegroundColor red -NoNewline
        
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddOwner" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $destinationDisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Group_Destination_PrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Owner_Source_PrimarySMTPAddress" -Value $member -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Owner_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors_Groups += $currenterror           
                        continue
                    }
                    
                }
            }
        }
    }
    Write-Host " done " -ForegroundColor Green
}

# Stamp Perms to Full Access and SendAs Groups for non user mailboxes
$matchedMailboxes = Import-Csv
$nonUserMailboxes = $matchedMailboxes | ?{$_.RecipientTypeDetails_Destination -ne "UserMailbox" -and $_.Migrate -ne "No"}
$progressref = ($nonUserMailboxes).count
$progresscounter = 0
foreach ($mailbox in $nonUserMailboxes) {
    #Set Variables
    $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
    $addressSplit = $destinationEmail -split "@"
    $FullAccessResourceName = $addressSplit[0] + "_FullAccess@" + $addressSplit[1]
    $SendAsResourceName = $addressSplit[0] + "_SendAs@" + $addressSplit[1]

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Stamping Perm Groups to Non-UserMailbox $($destinationEmail)"
 
    #Stamp Perms
    Write-Host "Updating Perms for $($destinationEmail) .. " -ForegroundColor cyan -NoNewline

    #Get Current Mailbox Permissions
    $fullAccessPerms = Get-MailboxPermission $destinationEmail | ?{$_.user -notlike "*nt authority*"}

    #Remove Current Full Access Perms
    foreach ($perm in $fullAccessPerms) {
        Remove-MailboxPermission -Identity $destinationEmail -User $perm.User -AccessRights FullAccess -Confirm:$false
        Write-Host ". " -ForegroundColor Yellow -NoNewline
    }
    #Remove Current Send-As Perms
    $SendAsPerms = Get-RecipientPermission $destinationEmail -AccessRights SendAs | ?{$_.Trustee -notlike "*nt authority*"}
    foreach ($perm in $SendAsPerms) {
        Remove-RecipientPermission -Identity $destinationEmail -Trustee $perm.Trustee -AccessRights SendAs -Confirm:$false
        Write-Host ". " -ForegroundColor Yellow -NoNewline
    }

    #Add Full Access Permission
    $permResult = Add-MailboxPermission -AccessRights FullAccess -Identity $destinationEmail -User $FullAccessResourceName -Automapping $false -Confirm:$false
    $permResult = Add-RecipientPermission -AccessRights SendAs -Identity $destinationEmail -Trustee $SendAsResourceName -Confirm:$false
    Write-Host "Succeeded " -ForegroundColor Green
}

## Update Permissions Group Membership for non user mailboxes
$matchedMailboxes = Import-Csv
$nonUserMailboxes = $matchedMailboxes | ?{$_.RecipientTypeDetails_Destination -ne "UserMailbox" -and $_.Migrate -ne "No"}

$notFoundPermUser = @()
$failures = @()
$progressref = ($nonUserMailboxes).count
$progresscounter = 0
foreach ($mailbox in $nonUserMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Mailbox Perms for $($mailbox.PrimarySMTPAddress_Destination)"
    Write-Host "Updating Mailbox Perms for $($mailbox.PrimarySMTPAddress_Destination).. " -NoNewline -ForegroundColor Cyan
    $addressSplit = $mailbox.PrimarySMTPAddress_Destination -split "@"
    $FullAccessResourceEmailAddress = $addressSplit[0] + "_FullAccess@ametek.com"
    $SendAsResourceEmailAddress= $addressSplit[0] + "_SendAs@ametek.com"

    # Add Full Access Permission Users
    if ($FullAccessUsersSplit = $mailbox.FullAccessPerms_Source -split ",") {   
        $FullAccessUsers = $FullAccessUsersSplit | ?{$_ -notlike "*NAMPR16*"}
        $progressref2 = ($FullAccessUsers).count
        $progresscounter2 = 0
        Write-Host "Setting up $($FullAccessUsers.count) Users with Full Access.. " -NoNewline
        foreach ($perm in $FullAccessUsers) {
            # Match the Perm user
            $trimPermUser = $perm.trim()
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Full Access Perms to $($trimPermUser)"

            if ($matchedUser = $matchedMailboxes | Where-Object {$_.PrimarySmtpAddress_Source -eq $trimPermUser}) {  
                $destinationEmail = $matchedUser.PrimarySMTPAddress_Destination
                ## Check if Perm User Exists
                if ($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue) {
                    try {
                        #Add DL Members
                        Add-DistributionGroupMember $FullAccessResourceEmailAddress -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                        Write-Host "." -ForegroundColor Green -NoNewline
                    }
                    catch {
                        Write-Host "." -ForegroundColor Yellow -NoNewline
                        
                        #Build Error Array
                        $currenterror = new-object PSObject

                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                        $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                        $currenterror | add-member -type noteproperty -name "DistributionList" -Value $dFullAccessResourceEmailAddress
                        $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                        $failures += $currenterror
                    }
                } 
            }
            else {
                Write-Host " Not Matched Recipient $($trimPermUser).. " -ForegroundColor red -NoNewline
                $notFoundPermUser += $trimPermUser
            }
        }
    }
    # Add Send-As Permission Users
    if ($SendAsUsersSplit = $mailbox.SendAsPerms_Source -split ",") {      
        $SendAsUsers = $SendAsUsersSplit | ?{$_ -notlike "*NAMPR16*"}
        $progressref3 = ($SendAsUsers).count
        $progresscounter3 = 0
        Write-Host " Setting up $($SendAsUsers.count) Users with Send-As.. " -NoNewline  
        foreach ($perm in $SendAsUsers) {
            # Match the Perm user
            $trimPermUser = $perm.trim()
            $progresscounter3 += 1
            $progresspercentcomplete3 = [math]::Round((($progresscounter3 / $progressref3)*100),2)
            $progressStatus3 = "["+$progresscounter3+" / "+$progressref3+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete3 -Status $progressStatus3 -Activity "Granting SendAs Perms to $($trimPermUser)"

            if ($matchedUser = $matchedMailboxes | Where-Object {$_.PrimarySmtpAddress_Source -eq $trimPermUser}) {
                $destinationEmail = $matchedUser.PrimarySMTPAddress_Destination
                ## Check if Perm User Exists
                if ($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue)  {
                    try {
                        #Add DL Members
                        Add-DistributionGroupMember $SendAsResourceEmailAddress -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                        Write-Host "." -ForegroundColor Green -NoNewline
                    }
                    catch {
                        Write-Host "." -ForegroundColor Yellow -NoNewline
                        
                        #Build Error Array
                        $currenterror = new-object PSObject

                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                        $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                        $currenterror | add-member -type noteproperty -name "DistributionList" -Value $SendAsResourceEmailAddress
                        $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                        $failures += $currenterror
                    }
                } 
            }
            else {
                Write-Host " Not Matched Recipient $($trimPermUser).. " -ForegroundColor red -NoNewline
                $notFoundPermUser += $trimPermUser
            }
        }
    }
    Write-Host " done" -ForegroundColor Green
}

# Stamp Perms to Full Access and SendAs Groups for UserMailbox
$allmatchedMailboxes = Import-Csv
$userMailboxes = $allmatchedMailboxes | ?{$_.RecipientTypeDetails_Destination -eq "UserMailbox" -and $_.Migrate -ne "No"}
$AllErrorsPerms = @()
$progressref = $userMailboxes.count
$progresscounter = 0
foreach ($mailbox in $userMailboxes) {
    #Set Variables
    $sourceEmail = $mailbox.PrimarySmtpAddress_Source
    $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
    $destinationDisplayName = $mailbox.DisplayName_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Stamping Perms on UserMailbox $($destinationDisplayName)"

    Write-Host "Stamping Perms on UserMailbox $($destinationDisplayName).. " -ForegroundColor Cyan -NoNewline
    
    #Stamp Full Access Perms for UserMailbox
    if ($mailbox.FullAccessPerms_Source) {
        $fullAccessPerms = $mailbox.FullAccessPerms_Source -split ","
        $fullAccessPermUsers = $fullAccessPerms | ?{$_ -notlike "*NAMPR16A*" -and $_ -ne "noreply@abaco.com" -and $_ -ne "smtp@abaco.com"}
        #Only Run for Legitimate Users
        if ($fullAccessPermUsers) {
            Write-Host "FullAccess.. " -ForegroundColor DarkYellow -NoNewline
 
            #Progress Bar 2
            $progressref2 = ($fullAccessPermUsers).count
            $progresscounter2 = 0
            foreach ($member in $fullAccessPermUsers) {
                #Member Check
                $memberCheck = @()
                $memberCheck = $allmatchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}

                #Progress Bar 2a
                $progresscounter2 += 1
                $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Full Access to $($memberCheck.PrimarySmtpAddress_Destination)"

                #Add Perms to Mailbox      
                try {
                    $permResult = Add-MailboxPermission -identity $destinationEmail -AccessRights FullAccess -User $memberCheck.PrimarySmtpAddress_Destination -Automapping $false -ea Stop -warningaction silentlycontinue
                    Write-Host "." -ForegroundColor Green -NoNewline
                }
                catch {
                    Write-Host "." -ForegroundColor red -NoNewline

                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantFullAccess" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "User" -Value $sourceEmail -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $destinationEmail -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrorsPerms += $currenterror           
                    continue
                }
            }
        }  
    }
    if ($mailbox.SendAsPerms_Source) {
        $SendAsPerms = $mailbox.SendAsPerms_Source -split ","
        $SendAsPermsUsers = $SendAsPerms | ?{$_ -notlike "*NAMPR16A*" -and $_ -ne "noreply@abaco.com" -and $_ -ne "smtp@abaco.com"}
        #Only Run for Legitimate Users
        if ($SendAsPermsUsers) {
            Write-Host "SendAs.. " -ForegroundColor DarkYellow -NoNewline
            #Progress Bar 2
            $progressref2 = ($SendAsPermsUsers).count
            $progresscounter2 = 0
            foreach ($member in $SendAsPermsUsers) {
                #Member Check
                $memberCheck = @()
                $member = $member.trim()
                $memberCheck = $allmatchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}

                #Progress Bar 2a
                $progresscounter2 += 1
                $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Send As to $($memberCheck.PrimarySmtpAddress_Destination)"

                #Add Perms to Mailbox      
                try {
                    $permResult = Add-RecipientPermission -identity $destinationEmail -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -confirm:$false -ea Stop
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
        }      
    }
    
    Write-Host " Completed " -ForegroundColor Green
}