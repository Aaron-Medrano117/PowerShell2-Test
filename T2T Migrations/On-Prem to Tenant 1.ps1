#Cenlar
$PilotBatch1 = Import-Excel
$progressref = $PilotBatch1.count
$progresscounter = 0
foreach ($user in $PilotBatch1) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Enable Remote Archive $($user.DisplayName)"
    Set-mailbox -Identity $user.Mail -RetentionPolicy "Default Archive and Retention Policy" -whatif
    Enable-Mailbox -Identity $user.Mail -RemoteArchive -ArchiveDomain cenlarfsb.mail.onmicrosoft.com -whatif
}

# Gather All Recipient Objects
$AllRecipientDetails = @()
$allRecipients = Get-Recipient -ResultSize Unlimited
$progressref = $allRecipients.count
$progresscounter = 0
foreach ($recipient in $allRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Recipient Details for $($recipient.PrimarySMTPAddress)"

    $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $recipient.DisplayName
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ";") -force

    $AllRecipientDetails += $currentuser
}
$AllRecipientDetails | Export-Excel "C:\Users\aaron.medrano\Desktop\post migration\EHN-AllRecipients.xlsx"

#Migration Report; Pull MailboxStats - Part 1

$PilotBatch1 = Import-CSV
$PilotBatch1Stats = @()
$progressref = ($PilotBatch1).count
$progresscounter = 0
foreach ($user in $PilotBatch1) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    
    #Pull MailboxStats and UserDetails
    $mbxStats = Get-MailboxStatistics $user.PrimarySMTPAddress

    #Create User Array
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName -force
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.PrimarySmtpAddress -force
    $currentuser | add-member -type noteproperty -name "Alias" -Value $user.alias -force
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ";") -force
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $user.legacyexchangedn) -force
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $user.HiddenFromAddressListsEnabled -force
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $user.DeliverToMailboxAndForward -force
    $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $user.ForwardingAddress -force
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $user.ForwardingSmtpAddress -force
    $currentuser | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize -force
    $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount -force

    #Pull Send on Behalf
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.primarysmtpaddress | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value ($mbxPermissions.user -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value $null -force
    }
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.PrimarySMTPAddress | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value $null  -force
    }
    # Archive Mailbox Check
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus -force
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveName" -Value $user.ArchiveName -force
    if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {           
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $ArchiveStats.TotalItemSize.Value -force 
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $ArchiveStats.ItemCount -force
    }
    else {
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $null -force 
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $null -force
    }
    $PilotBatch1Stats += $currentuser
}

$PilotBatch1Stats | Export-CSV

#Migration Report; Add MigrationStats - Part 2
$PilotBatch1Stats = Import-CSV
$progressref = ($PilotBatch1Stats).count
$progresscounter = 0
foreach ($user in $PilotBatch1Stats) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Migration Details for $($user.DisplayName)"
    
    #Pull MailboxStats and UserDetails
    $moveStats = Get-MoveRequestStatistics $user.PrimarySMTPAddress

    #Create User Array
    $user | add-member -type noteproperty -name "StatusDetail" -Value $moveStats.StatusDetail -force
    $user | add-member -type noteproperty -name "PercentComplete" -Value $moveStats.PercentComplete -force
    $user | add-member -type noteproperty -name "TotalMailboxSize" -Value $moveStats.TotalMailboxSize -force
    $user | add-member -type noteproperty -name "TotalMailboxItemCount" -Value $moveStats.TotalMailboxItemCount -force
    $user | add-member -type noteproperty -name "ItemsTransferred" -Value $moveStats.ItemsTransferred -force
    $user | Add-Member -type NoteProperty -Name "BytesTransferred" -Value $moveStats.BytesTransferred -force
    $user | add-member -type noteproperty -name "SourceServer" -Value $moveStats.SourceServer -force
    $user | add-member -type noteproperty -name "RemoteHostName" -Value $moveStats.RemoteHostName -force
    $user | add-member -type noteproperty -name "StartTimeStamp" -Value $moveStats.StartTimeStamp -force
    $user | add-member -type noteproperty -name "InitialSeedingCompletedTimeStamp" -Value $moveStats.InitialSeedingCompletedTimeStamp -force
    $user | add-member -type noteproperty -name "CompletionTimeStamp" -Value $moveStats.CompletionTimeStamp -force
    $user | add-member -type noteproperty -name "OverallDuration" -Value $moveStats.OverallDuration -force
    $user | add-member -type noteproperty -name "TotalInProgressDuration" -Value $moveStats.TotalInProgressDuration -force
}

$PilotBatch1Stats | Export-CSV

# Pull Full Access Permissions for Pilot Group
$batch = Import-Excel
$allRecipientPermissions = Import-Excel
$progressref = ($batch).count
$progresscounter = 0
$PermAccounts = @()
foreach ($user in $batch) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering FullAccess Relations for $($user.DisplayName)"
    
    if ([array]$PermAccounts += $allRecipientPermissions |  ?{$_.PermUser -eq $user.Mail}){
        Write-Host "." -ForegroundColor Green -NoNewline
    }
    else {
        Write-Host "." -ForegroundColor Yellow -NoNewline
    }  
}

$allMailGroupMembers = @()
#ProgressBarA
$progressref = ($allMailGroups).count
$progresscounter = 0
foreach ($group in $allMailGroups) {
    #ProgressBarB
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Group Details for $($object.DisplayName)"

    #Pull DynamicDistributionGroup Details
    if ($object.RecipientTypeDetails -eq "DynamicDistributionGroup") {
        $groupMembers = Get-DynamicDistributionGroupMember $dynamicGroup.PrimarySMTPAddress -ErrorAction SilentlyContinue -ResultSize unlimited            
    }
    #Pull Mail Group Details
    elseif ($object.RecipientTypeDetails -eq "MailUniversalDistributionGroup" -or $object.RecipientTypeDetails -eq "MailUniversalSecurityGroup") {
        $groupMembers = Get-DistributionGroupMember $PrimarySMTPAddress -ResultSize unlimited -ErrorAction SilentlyContinue
    }
    #Gather Group Mailbox
    elseif ($object.RecipientTypeDetails -eq "GroupMailbox") {
        $groupMembers = Get-UnifiedGroupLinks -Identity $PrimarySMTPAddress -LinkType Member -ResultSize unlimited
    }

    Write-Host "Gathering Group Details for $($group.DisplayName) .." -NoNewline -ForegroundColor Cyan
    
        Write-Host "Found $($groupMembers.count) members .." -ForegroundColor Yellow -NoNewline
    
        $progressref2 = ($groupMembers).count
        $progresscounter2 = 0
        foreach ($member in $groupMembers) {
            #ProgressBarB
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -Id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "$($member.DisplayName)"

            $currentgroup = New-Object psobject
            $currentgroup | Add-Member -Type noteproperty -Name "Group-DisplayName" -Value $group.DisplayName
            $currentgroup | Add-Member -Type noteproperty -Name "Group-PrimarySMTPAddress" -Value $group.PrimarySMTPAddress
            $currentgroup | Add-Member -Type noteproperty -Name "Group-RecipientTypeDetails" -Value $group.RecipientTypeDetails
            $currentgroup | Add-Member -Type noteproperty -Name "Member-DisplayName" -Value $member.DisplayName
            $currentgroup | Add-Member -Type noteproperty -Name "Member-RecipientType" -Value $member.RecipientTypeDetails
            $currentgroup | Add-Member -Type noteproperty -Name "Member-PrimarySMTPAddress" -Value $member.PrimarySMTPAddress
                        
            $allMailGroupMembers += $currentgroup
        }
        Write-Host "done" -ForegroundColor Green
}
$allMailGroupMembers 