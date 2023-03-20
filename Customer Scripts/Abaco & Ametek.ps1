#Abaco & Ametek

#Gather Mailbox Stats
$sourceMailboxes = Get-Mailbox -ResultSize Unlimited -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'"
$sourceMailboxStats = @()
$progressref = ($sourceMailboxes).count
$progresscounter = 0
foreach ($user in $sourceMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    
    #Pull MailboxStats and UserDetails
    $mbxStats = Get-MailboxStatistics $user.PrimarySMTPAddress
    $msoluser = Get-MsolUser -UserPrincipalName $user.UserPrincipalName
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    #Create User Array
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $msoluser.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName_Source" -Value $msoluser.userprincipalname
    $currentuser | add-member -type noteproperty -name "IsLicensed_Source" -Value $msoluser.IsLicensed
    $currentuser | add-member -type noteproperty -name "Licenses_Source" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
    $currentuser | add-member -type noteproperty -name "License-DisabledArray_Source" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
    $currentuser | add-member -type noteproperty -name "Department_Source" -Value $msoluser.Department    
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $user.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress_Source" -Value $user.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "Alias_Source" -Value $user.alias
    $currentuser | add-member -type noteproperty -name "CustomAttribute7_Source" -Value $user.CustomAttribute7
    $currentuser | add-member -type noteproperty -name "WhenCreated_Source" -Value $user.WhenCreated
    $currentuser | add-member -type noteproperty -name "LastLogonTime_Source" -Value $mbxStats.LastLogonTime
    $currentuser | add-member -type noteproperty -name "EmailAddresses_Source" -Value ($EmailAddresses -join ";")
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN_Source" -Value ("x500:" + $user.legacyexchangedn)
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Source" -Value $user.HiddenFromAddressListsEnabled
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward_Source" -Value $user.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ForwardingAddress_Source" -Value $user.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress_Source" -Value $user.ForwardingSmtpAddress
    $currentuser | Add-Member -type NoteProperty -Name "MBXSize_Source" -Value $MBXStats.TotalItemSize
    $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount_Source" -Value $MBXStats.ItemCount

    #Pull Send on Behalf
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo_Source" -Value ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.primarysmtpaddress | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms_Source" -Value ($mbxPermissions.user -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms_Source" -Value $null
    }
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.PrimarySMTPAddress | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "SendAsPerms_Source" -Value ($sendAsPermsCheck.trustee -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "SendAsPerms_Source" -Value $null
    }
    # Archive Mailbox Check
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus_Source" -Value $user.ArchiveStatus
    if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {           
        $currentuser | add-member -type noteproperty -name "ArchiveSize_Source" -Value $ArchiveStats.TotalItemSize.Value
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount_Source" -Value $ArchiveStats.ItemCount
    }
    else {
        $currentuser | add-member -type noteproperty -name "ArchiveSize_Source" -Value $null
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount_Source" -Value $null
    }
    $sourceMailboxStats += $currentuser
}

# Match Mailboxes to existing MatchedMailbox spreadsheet
$sourceMailboxStats = Import-Csv
$allMatchedUsers = Import-Csv
$progressref = ($sourceMailboxStats).count
$progresscounter = 0

foreach ($user in $sourceMailboxStats) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName_Source)"
    
    #match against CSV
    if ($matchedCSVUser = $allMatchedUsers | ?{$_.UserPrincipalName_Source -eq $user.UserPrincipalName_Source}) {
        $user | add-member -type noteproperty -name "Migrate" -Value $matchedCSVUser.Migrate -force
        $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $matchedCSVUser.OneDriveUrl_Source -force
        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)_Source" -Value $matchedCSVUser."OneDriveCurrentStorage(MB)_Source" -force
        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)_Source" -Value $matchedCSVUser."OneDriveCurrentStorage(GB)_Source" -force
        $user | add-member -type noteproperty -name "OU" -Value $matchedCSVUser.OU
    }
    else {
        $user | add-member -type noteproperty -name "Migrate" -Value $null -force
        $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $null -force
        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)_Source" -Value $null -force
        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)_Source" -Value $null -force
        $user | add-member -type noteproperty -name "OU" -Value $null
    }
    
}

# Match Recipients to existing MatchedMailbox spreadsheet
$sourceAllRecipients = Import-Excel 
$allMatchedUsers = Import-Csv
$progressref = ($sourceAllRecipients).count
$progresscounter = 0

foreach ($user in $sourceAllRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.PrimarySMTPAddress)"
    
    #match against CSV
    if ($matchedCSVUser = $allMatchedUsers | ?{$_.PrimarySmtpAddress_Source -eq $user.PrimarySMTPAddress}) {
        $user | add-member -type noteproperty -name "Matched" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedCSVUser.DisplayName_Destination -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedCSVUser.RecipientTypeDetails_Destination -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedCSVUser.PrimarySmtpAddress_Destination -force
    }
    else {
        $user | add-member -type noteproperty -name "Matched" -Value $False -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -force
    } 
}

# Match Recipients to existing MatchedGroups spreadsheet
$sourceAllRecipients = Import-Excel 
$allMatchedGroups = Import-Csv
$progressref = ($sourceAllRecipients | ?{$_.Matched -eq $false}).count
$progresscounter = 0

foreach ($user in ($sourceAllRecipients | ?{$_.Matched -eq $false})) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.PrimarySMTPAddress)"
    
    #match against CSV
    if ($matchedCSVUser = $allMatchedGroups | ?{$_.PrimarySmtpAddress_Source -eq $user.PrimarySMTPAddress}) {
        $user | add-member -type noteproperty -name "Matched" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedCSVUser.DisplayName_Destination -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedCSVUser.RecipientTypeDetails_Destination -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedCSVUser.PrimarySmtpAddress_Destination -force
    }
    else {
        $user | add-member -type noteproperty -name "Matched" -Value $False -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -force
    } 
}

# Match Recipients to Ametek Tenant
$sourceAllRecipients = Import-Excel 
$progressref = ($sourceAllRecipients | ?{$_.Matched -eq $false}).count
$progresscounter = 0

foreach ($user in ($sourceAllRecipients | ?{$_.Matched -eq $false})) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.PrimarySMTPAddress)"
    
    $DisplayName = $user.DisplayName
    $newDisplayName = "Abaco-" + $user.DisplayName
    $newDisplayName2 = "Abaco - " + $user.DisplayName
    $newDisplayName3 = "Abaco " + $user.DisplayName
    $newDisplayName4 = $user.DisplayName -replace ("Abaco ","Abaco - ")
    $newDisplayName5 = $user.DisplayName -replace ("Abaco ","Abaco-")
    #match against CSV
    if ($matchedUser = Get-Recipient $newDisplayName -ErrorAction SilentlyContinue) {
        $user | add-member -type noteproperty -name "Matched" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedUser.DisplayName -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedUser.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedUser.PrimarySmtpAddress -force
    }
    elseif ($matchedUser = Get-Recipient $newDisplayName2 -ErrorAction SilentlyContinue)
    {
        $user | add-member -type noteproperty -name "Matched" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedUser.DisplayName -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedUser.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedUser.PrimarySmtpAddress -force
    }
    elseif ($matchedUser = Get-Recipient $newDisplayName3 -ErrorAction SilentlyContinue)
    {
        $user | add-member -type noteproperty -name "Matched" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedUser.DisplayName -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedUser.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedUser.PrimarySmtpAddress -force
    }
    elseif ($matchedUser = Get-Recipient $newDisplayName4 -ErrorAction SilentlyContinue)
    {
        $user | add-member -type noteproperty -name "Matched" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedUser.DisplayName -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedUser.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedUser.PrimarySmtpAddress -force
    }
    elseif ($matchedUser = Get-Recipient $newDisplayName5 -ErrorAction SilentlyContinue)
    {
        $user | add-member -type noteproperty -name "Matched" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedUser.DisplayName -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedUser.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedUser.PrimarySmtpAddress -force
    }
    elseif ($matchedUser = Get-Recipient $DisplayName -ErrorAction SilentlyContinue)
    {
        $user | add-member -type noteproperty -name "Matched" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedUser.DisplayName -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedUser.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedUser.PrimarySmtpAddress -force
    }
    else
    {
        $user | add-member -type noteproperty -name "Matched" -Value $False -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -force
    } 

}

#remove SIP, SPO, x500, onmicrosoft addresses, and non migrating domains  - matched recipients
$progressref = $matchedRecipients.count
$progresscounter = 0
foreach ($user in $matchedRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Alternate Address $($user.PrimarySMTPAddress)"

    $alternateAddresses = $user.EmailAddresses -split ","
    $newAlternateAddresses = $alternateAddresses | ?{$_ -notlike "*spo:*" -and $_ -notlike "*sip:*" -and $_ -notlike "*x500*" -and $_ -notlike "*@abaco1.mail.onmicrosoft.com" -and $_ -notlike "*@abaco1.onmicrosoft.com" -and $_ -ne $user.PrimarySMTPAddress}
    $user | add-member -type noteproperty -name "Filtered_Addresses" -Value ($newAlternateAddresses -join ";") -force
}


#Match Recipients to All Objects
$sourceAllRecipients = Import-Excel

$progressref = ($sourceAllRecipients | ?{$_.Matched -eq $false}).count
$progresscounter = 0
foreach ($recipient in ($sourceAllRecipients | ?{$_.Matched -eq $false})) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Recipient Details for $($recipient.PrimarySMTPAddress)"
    $matchedObject = @()
    $newEmailAddress = ($recipient.PrimarySMTPAddress -split "@")[0] + "@ametek.com"
    if ($matchedObject = Get-Recipient $recipient.DisplayName -ErrorAction SilentlyContinue) {
        $recipient | add-member -type noteproperty -name "Matched" -Value $True -force
        $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedObject.DisplayName -force
        $recipient | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedObject.RecipientTypeDetails -force
        $recipient | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $matchedObject.PrimarySMTPAddress -force
    }
    if ($matchedObject = Get-Recipient $newEmailAddress -ErrorAction SilentlyContinue) {
        $recipient | add-member -type noteproperty -name "Matched" -Value $True -force
        $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedObject.DisplayName -force
        $recipient | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedObject.RecipientTypeDetails -force
        $recipient | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $matchedObject.PrimarySMTPAddress -force
    }
    else {
        $recipient | add-member -type noteproperty -name "Matched" -Value $false -force
        $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $recipient | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $recipient | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $null -force
    }
}

# Match Mailboxes and add to same spreadsheet. Check based on NEWUPN, DisplayName, NEWPRIMARYSMTP
$sourceMailboxes = Import-Csv
$newDomain = Read-Host "What is the new domain migrating to?"
$progressref = ($sourceMailboxes).count
$progresscounter = 0
foreach ($user in $sourceMailboxes) {
    # Set Variables
    $DisplayNameSource = $user.DisplayName_Source
    $PrimarySMTPAddress = $user.PrimarySMTPAddress_Source
    $addressSplit = $user.UserPrincipalName_Source -split "@"
    $NewUPN = $addressSplit[0] + "@" + $newDomain
    $SMTPaddressSplit = $PrimarySMTPAddress -split "@"
    $newPrimarySMTPAddress = $SMTPaddressSplit[0] + "@" + $newDomain
    $abacoSharedAddress = "abaco." + $newPrimarySMTPAddress

    #ProgressBar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($DisplayNameSource)"
    
 
    #Old Address Mail User Check
    if ($mailUserCheck = Get-MailUser $PrimarySMTPAddress -ErrorAction SilentlyContinue) {
        $msoluser = Get-MsolUser -UserPrincipalName $mailUserCheck.userPrincipalName -ErrorAction SilentlyContinue
        Write-Host "$($msoluser.UserPrincipalName) MailUser found with migrating addres" -ForegroundColor Yellow   

        #Pull Mailbox Stats
        $recipient = Get-Recipient $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "MailUserMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $recipient.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "WhenCreate_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
        }
        else  {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }
    }
    #NEW UPN Check
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $NewUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with newUPN" -ForegroundColor Green
    
        #Pull Mailbox Stats
        $recipient = Get-Recipient $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "NewUPNMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $recipient.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "WhenCreate_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
        }
        else  {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }

    }  
    #NEW PrimarySMTPAddress Check
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $newPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow     
        
        #Pull Mailbox Stats
        $recipient = Get-Recipient $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "NEWSMTPAddressMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $recipient.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "WhenCreate_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
        }
        else  {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }

    }
    #UPN Check based on DisplayName
    elseif ($mailbox = Get-Mailbox $abacoSharedAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($mailbox.UserPrincipalName) User found with AbacoAddress" -ForegroundColor Yellow
        
        #Pull Mailbox Stats
        $msoluser = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        $recipient = Get-Recipient $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "DisplayNameMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $recipient.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "WhenCreate_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
        }
        else  {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }
    }
    #UPN Check based on DisplayName
    elseif ($msoluser = Get-Msoluser -SearchString $DisplayNameSource -ErrorAction SilentlyContinue) {
        if ($msoluser.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName)"
            $mailbox = Get-Mailbox $user.DisplayName -ErrorAction SilentlyContinue
            $msoluser = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        }
        Write-Host "$($msoluser.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
        
        #Pull Mailbox Stats
        $recipient = Get-Recipient $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "DisplayNameMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $recipient.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "WhenCreate_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
        }
        else  {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }
    }
    #not Found User
    else {
        Write-Host "  Unable to find user for $($DisplayNameSource)" -ForegroundColor Red
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $null -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "WhenCreate_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
    }    
}

## Assign Licenses to users and set forward
$matchedUsers = import-csv
$matchedUsers = $matchedUsers | ? {$_.ExistsInDestination -eq "True"}
$matchedUsers = $matchedUsers | ? {$_.Migrate -eq "Yes"}
$matchedUsers = $matchedUsers | ? {$_.RecipientTypeDetails -eq "UserMailbox"}
$progressref = ($matchedUsers).count
$progresscounter = 0
$failedtoUpdateUsers = @()

foreach ($user in $matchedUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating License Details for $($user.DisplayName_Destination)"
    
    #Clear Variables
    $disabledArray = @()
    $LicenseOptions = @()
    $ametekLicense = $null

    Write-Host "Updating $($user.DisplayName_Destination) .." -foregroundcolor cyan -nonewline
    if ($user.Licenses_Source -like "*E5*") {
        $disabledArray = "POWER_VIRTUAL_AGENTS_O365_P3","CDS_O365_P3","MICROSOFTBOOKINGS","RECORDS_MANAGEMENT","CUSTOMER_KEY","INFORMATION_BARRIERS","KAIZALA_STANDALONE","PREMIUM_ENCRYPTION","Deskless","SWAY","MCOEV","LOCKBOX_ENTERPRISE"
        $ametekLicense = (Get-MsolAccountSku | ?{$_.AccountSkuID -like "*E5*"}).AccountSkuID.Trim()
        Write-Host "Adding E5 license .." -foregroundcolor cyan -nonewline
    }
    elseif ($user.Licenses_Source -like "*E3*") {
        $disabledArray = "POWER_VIRTUAL_AGENTS_O365_P2","CDS_O365_P2","MICROSOFTBOOKINGS","KAIZALA_O365_P3","Deskless","SWAY"
        $ametekLicense = ((Get-MsolAccountSku | ?{$_.AccountSkuID -like "*E3*"}).AccountSkuID).Trim()
        Write-Host "Adding E3 license .. " -foregroundcolor cyan -nonewline
    }

    #Set Usage Location
    Set-Msoluser -UserPrincipalName $user.UserPrincipalName_Destination -UsageLocation US
    #Add License with Disabled Array
    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $ametekLicense -DisabledPlans $disabledArray
    try {
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_Destination -AddLicenses $ametekLicense -LicenseOptions $LicenseOptions -ErrorAction Stop
    }
    catch {
        Write-Host ". " -foregroundcolor Red -nonewline
        Write-Warning -Message "$($_.Exception)"
        $failedtoUpdateUsers += $user
    }

    #Add Additional Licenses
    $otherLicenseArraySplit = $user.Licenses -split ","
    if ($otherLicenseArray = $otherLicenseArraySplit | ?{$_ -ne "Abaco1:SPE_E5" -and $_ -ne "Abaco1:SPE_E3"}) {
        Write-Host "Additional Licenses .. " -foregroundcolor Cyan -nonewline
        foreach ($userlicense in $otherLicenseArray) {
            Write-Host ". " -foregroundcolor yellow -nonewline
            $licenseSplit = $userlicense -split ":"
            $licenseSplitName = $licenseSplit[1]
            $updatedLicense = ((Get-MsolAccountSku | ?{$_.AccountSkuID -like "*$licenseSplitName*"}).AccountSkuID).Trim() 
            try {
                Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_Destination -AddLicenses $updatedLicense -ErrorAction Stop
            }
            catch {
                Write-Host ". " -foregroundcolor Red -nonewline
            }
        }
    }  
    #Set Mailbox Forward
    Write-Host "Set Mailbox Forward.. " -foregroundcolor Cyan -nonewline
    $count = 0
    $success = $null
    do {
        if ($mailboxCheck = Get-Mailbox $user.PrimarySMTPAddress_Destination -ErrorAction SilentlyContinue) {
            Set-Mailbox -identity $user.PrimarySMTPAddress_Destination -ForwardingSmtpAddress $user.PrimarySMTPAddress -DeliverToMailboxAndForward $true
            $success = $true
            continue
        }
        else {
            Write-host "Next attempt in 30 seconds .. " -foregroundcolor yellow -nonewline
            Start-sleep -Seconds 30
            $count++
        }
    } until ($mailboxCheck -or $success -or $count -eq 4)
    Write-Host "done " -foregroundcolor green 
}

## match Groups from Source to Destination CONDENSED
$sourceGroups = Import-Csv 
$progressref = ($sourceGroups).count
$progresscounter = 0
foreach ($group in $sourceGroups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Matched Groups for $($group.DisplayName)"

    $groupCheck = @()
    $addressSplits = $group.PrimarySMTPAddress -split "@"
    $AbacoMatch = "abaco." + $group.PrimarySMTPAddress
    $AbacoMatch2 = "abaco-" + $group.PrimarySMTPAddress
    $newPrimarySMTPAddress = $addressSplits[0] + "@abaco.com"
    $AbacodisplayName = "Abaco - " + $group.DisplayName
    $AbacodisplayName2 = "Abaco-" + $group.DisplayName

    if ($groupCheck = Get-Recipient $AbacoMatch -ea silentlycontinue) {
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value "AbacoEmailMatch" -Force
    }
    elseif ($groupCheck = Get-Recipient $AbacodisplayName -ea silentlycontinue) {
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value "AbacoDisplayNameMatch" -Force
    }
    elseif ($groupCheck = Get-Recipient $newPrimarySMTPAddress -ea silentlycontinue) {
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value "NewAddressMatch" -Force
    } 
    elseif ($groupCheck = Get-Recipient $AbacoMatch2 -ea silentlycontinue) {
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value "AbacoEmailMatch" -Force
    }
    elseif ($groupCheck = Get-Recipient $AbacodisplayName2 -ea silentlycontinue) {
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value "AbacoDisplayNameMatch" -Force
    }
    elseif ($groupCheck = Get-Recipient $group.DisplayName -ea silentlycontinue) {
        if ($groupCheck.count -gt 1) {
            $group | add-member -type noteproperty -name "ExistsInDestination" -Value "Duplicate" -Force
        }
        else {
            $group | add-member -type noteproperty -name "ExistsInDestination" -Value "DisplayNameMatch" -Force
        }
    }

    if ($groupCheck){
        $EmailAddresses = $groupCheck | select -ExpandProperty EmailAddresses

        $group | add-member -type noteproperty -name "DisplayName_Destination" -Value $groupCheck.DisplayName -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $groupCheck.RecipientTypeDetails -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $groupCheck.PrimarySmtpAddress -Force
        $group | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -Force
    }
    else {
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value $false -Force
        $group | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -Force
        $group | add-member -type noteproperty -name "EmailAddresses_Destination" -Value $null -Force
    }
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

    Set-DistributionGroup -identity $user.PrimarySmtpAddress_Destination -MemberJoinRestriction $user.MemberJoinRestriction -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($user.RequireSenderAuthenticationEnabled)) -SendModerationNotifications $user.SendModerationNotifications -ModerationEnabled ([System.Convert]::ToBoolean($user.ModerationEnabled)) -ReportToOriginatorEnabled ([System.Convert]::ToBoolean($user.ReportToOriginatorEnabled)) -warningaction silentlycontinue
    
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

#Set User to disable exhange license
$matchedMailboxes = $matchedMailboxes | ? {$_.RecipientTypeDetails -eq "UserMailbox" -and $_.Migrate -ne "No"}
$matchedMailboxes = $matchedMailboxes | ? {$_.ExistsInDestination -eq "MailUserMatch" -and $_.Migrate -ne "No"}
$progressref = ($matchedMailboxes).count
$progresscounter = 0
foreach ($user in $matchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating License Details for $($user.DisplayName_Destination)"
    
    #Clear Variables
    $disabledArray = @()
    $LicenseOptions = @()
    $failedtoUpdateUsers = @()
    $ametekLicense = $null

    Write-Host "Updating $($user.DisplayName_Destination) .." -foregroundcolor cyan -nonewline
    if ($user.Licenses -like "*E5*") {
        $disabledArray = "POWER_VIRTUAL_AGENTS_O365_P3","CDS_O365_P3","MICROSOFTBOOKINGS","RECORDS_MANAGEMENT","INSIDER_RISK_MANAGEMENT","CUSTOMER_KEY","COMMUNICATIONS_COMPLIANCE","INFORMATION_BARRIERS","KAIZALA_STANDALONE","PREMIUM_ENCRYPTION","Deskless","SWAY","MCOEV","LOCKBOX_ENTERPRISE","Exchange_S_Enterprise"
        $ametekLicense = (Get-MsolAccountSku | ?{$_.AccountSkuID -like "*E5*"}).AccountSkuID.Trim()
        Write-Host "Adding E5 license .." -foregroundcolor cyan -nonewline
    }
    elseif ($user.Licenses -like "*E3*") {
        $disabledArray = "POWER_VIRTUAL_AGENTS_O365_P2","CDS_O365_P2","MICROSOFTBOOKINGS","KAIZALA_O365_P3","Deskless","SWAY","Exchange_S_Enterprise"
        $ametekLicense = ((Get-MsolAccountSku | ?{$_.AccountSkuID -like "*E3*"}).AccountSkuID).Trim()
        Write-Host "Adding E3 license .. " -foregroundcolor cyan -nonewline
    }
    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $ametekLicense -DisabledPlans $disabledArray
        #Update License Disabled Array
        try {
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_Destination -Removelicense $ametekLicense  -ErrorAction Stop
        }
        catch {
            Write-Host ". " -foregroundcolor Red -nonewline
            Write-Warning -Message "$($_.Exception)"
            $failedtoUpdateUsers += $user
        }

        Write-Host "done " -foregroundcolor green -nonewline
}

#Add license, disabled array AND disable exchange license
$matchedMailboxes = $matchedMailboxes | ? {$_.RecipientTypeDetails -eq "UserMailbox" -and $_.Migrate -ne "No"}
$matchedMailboxes = $matchedMailboxes | ? {$_.ExistsInDestination -eq "MailUserMatch" -and $_.Migrate -ne "No"}
$progressref = ($matchedMailboxes).count
$progresscounter = 0
foreach ($user in $matchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating License Details for $($user.DisplayName_Destination)"
    
    #Clear Variables
    $disabledArray = @()
    $LicenseOptions = @()
    $failedtoUpdateUsers = @()
    $ametekLicense = $null

    Write-Host "Updating $($user.DisplayName_Destination) .." -foregroundcolor cyan -nonewline
    if ($user.Licenses -like "*E5*") {
        $disabledArray = "POWER_VIRTUAL_AGENTS_O365_P3","CDS_O365_P3","MICROSOFTBOOKINGS","RECORDS_MANAGEMENT","INSIDER_RISK_MANAGEMENT","CUSTOMER_KEY","COMMUNICATIONS_COMPLIANCE","INFORMATION_BARRIERS","KAIZALA_STANDALONE","PREMIUM_ENCRYPTION","Deskless","SWAY","MCOEV","LOCKBOX_ENTERPRISE","Exchange_S_Enterprise"
        $ametekLicense = (Get-MsolAccountSku | ?{$_.AccountSkuID -like "*E5*"}).AccountSkuID.Trim()
        Write-Host "Adding E5 license .." -foregroundcolor cyan -nonewline
    }
    elseif ($user.Licenses -like "*E3*") {
        $disabledArray = "POWER_VIRTUAL_AGENTS_O365_P2","CDS_O365_P2","MICROSOFTBOOKINGS","KAIZALA_O365_P3","Deskless","SWAY","Exchange_S_Enterprise"
        $ametekLicense = ((Get-MsolAccountSku | ?{$_.AccountSkuID -like "*E3*"}).AccountSkuID).Trim()
        Write-Host "Adding E3 license .. " -foregroundcolor cyan -nonewline
    }
    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $ametekLicense -DisabledPlans $disabledArray
        #Update License Disabled Array
        try {
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_Destination -AddLicenses $ametekLicense -LicenseOptions $LicenseOptions -ErrorAction Stop
        }
        catch {
            Write-Host ". " -foregroundcolor Red -nonewline
        }

        Write-Host "done " -foregroundcolor green -nonewline
}

## Assign Licenses to users - enable Exchange
$matchedUsers = import-csv
$matchedUsers = $matchedUsers | ? {$_.Licenses_Source}
$progressref = ($matchedUsers).count
$progresscounter = 0
$failedtoUpdateUsers = @()

foreach ($user in $matchedUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating License Details for $($user.DisplayName_Destination)"
    
    #Clear Variables
    $disabledArray = @()
    $LicenseOptions = @()
    $ametekLicense = $null

    Write-Host "Updating $($user.DisplayName_Destination) .." -foregroundcolor cyan -nonewline
    if ($user.Licenses_Source -like "*E5*") {
        $disabledArray = "POWER_VIRTUAL_AGENTS_O365_P3","CDS_O365_P3","MICROSOFTBOOKINGS","RECORDS_MANAGEMENT","CUSTOMER_KEY","INFORMATION_BARRIERS","KAIZALA_STANDALONE","PREMIUM_ENCRYPTION","Deskless","SWAY","MCOEV","LOCKBOX_ENTERPRISE"
        $ametekLicense = (Get-MsolAccountSku | ?{$_.AccountSkuID -like "*E5*"}).AccountSkuID.Trim()
        Write-Host "Adding E5 license .." -foregroundcolor cyan -nonewline
    }
    elseif ($user.Licenses_Source -like "*E3*") {
        $disabledArray = "POWER_VIRTUAL_AGENTS_O365_P2","CDS_O365_P2","MICROSOFTBOOKINGS","KAIZALA_O365_P3","Deskless","SWAY"
        $ametekLicense = ((Get-MsolAccountSku | ?{$_.AccountSkuID -like "*E3*"}).AccountSkuID).Trim()
        Write-Host "Adding E3 license .. " -foregroundcolor cyan -nonewline
    }

    #Set Usage Location
    Set-Msoluser -UserPrincipalName $user.UserPrincipalName_Destination -UsageLocation US
    #Add License with Disabled Array
    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $ametekLicense -DisabledPlans $disabledArray
    if ((Get-MsolUser -UserPrincipalName $user.UserPrincipalName_Destination).licenses.accountskuid -contains $ametekLicense) {
        try {
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_Destination -LicenseOptions $LicenseOptions -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Host "updated " -foregroundcolor Green -nonewline
        }
        catch {
            Write-Host ". " -foregroundcolor Red -nonewline
            Write-Warning -Message "$($_.Exception)" -none
            $failedtoUpdateUsers += $user
        }
    }
    else {
        try {
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_Destination -AddLicenses $ametekLicense -LicenseOptions $LicenseOptions -ErrorAction Stop
            Write-Host "added " -foregroundcolor Green -nonewline
        }
        catch {
            Write-Host ". " -foregroundcolor Red -nonewline
            Write-Warning -Message "$($_.Exception)"
            $failedtoUpdateUsers += $user
        }
    }
    

    #Add Additional Licenses
    $otherLicenseArraySplit = $user.Licenses_Source -split ","
    if ($otherLicenseArray = $otherLicenseArraySplit | ?{$_ -ne "Abaco1:SPE_E5" -and $_ -ne "Abaco1:SPE_E3"}) {
        Write-Host "Additional Licenses .. " -foregroundcolor Cyan -nonewline
        foreach ($userlicense in $otherLicenseArray) {
            $licenseSplit = $userlicense -split ":"
            $licenseSplitName = $licenseSplit[1]
            $updatedLicense = ((Get-MsolAccountSku | ?{$_.AccountSkuID -like "*$licenseSplitName*"}).AccountSkuID).Trim()
            if (!((Get-MsolUser -UserPrincipalName $user.UserPrincipalName_Destination).licenses.accountskuid -contains $updatedLicense)) {
                try {
                    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_Destination -AddLicenses $updatedLicense -ErrorAction Stop
                    Write-Host ". " -foregroundcolor yellow -nonewline
                }
                catch {
                    Write-Host ". " -foregroundcolor Red -nonewline
                }
            }           
        }
    }
    Write-Host "done " -foregroundcolor Green
}

## SEt Forward
$matchedmailboxes = import-csv
$matchedmailboxes = $matchedmailboxes | ? {$_.ExistsInDestination -eq "True" -and $_.Migrate -ne "No"}
$progressref = ($matchedmailboxes).count
$progresscounter = 0
$failedtoUpdateUsers = @()

foreach ($user in $matchedmailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating License Details for $($user.DisplayName_Destination)"
    
    #Clear Variables
    $disabledArray = @()
    $LicenseOptions = @()
    $ametekLicense = $null
    $mailboxCheck = @()

    Write-Host "Updating $($user.PrimarySMTPAddress_Destination) .. to $($user.PrimarySMTPAddress) " -foregroundcolor cyan -nonewline
    #Set Mailbox Forward
    Write-Host "Set Mailbox Forward.. " -foregroundcolor Cyan -nonewline
    $count = 0
    $success = $null
    do {
        if ($mailboxCheck = Get-Mailbox $user.PrimarySMTPAddress_Destination) {
            Set-Mailbox -identity $user.PrimarySMTPAddress_Destination -ForwardingSmtpAddress $user.PrimarySMTPAddress -DeliverToMailboxAndForward $true
            $success = $true
            continue
        }
        else {
            Write-host "Next attempt in 30 seconds .. " -foregroundcolor yellow -nonewline
            Start-sleep -Seconds 30
            $count++
        }
    } until ($mailboxCheck -or $success -or $count -eq 4)
    Write-Host "done " -foregroundcolor green
}

#set mailboxes Retention Policy
$progressref = ($matchedUsers).count
$progresscounter = 0
$failedtoUpdateUsers = @()
foreach ($user in $matchedUsers) {  
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Setting Retention Hold for $($user.DisplayName_Destination)"

    Write-Host "$($user.DisplayName_Destination) .. " -foregroundcolor cyan -nonewline
    try {
        Set-Mailbox -RetentionHoldEnabled $true -Identity $user.PrimarySmtpAddress_Destination -ErrorAction Stop
        Write-Host  "done " -foregroundcolor Green
    }
    catch {
        Write-Host  "done " -foregroundcolor Red
        Write-Warning -Message "$($_.Exception)"
        $failedtoUpdateUsers += $user
    }   
}

#Get SharePoint Source Details
$matchedMailboxes = Import-Csv
$progressref = ($matchedMailboxes).count
$progresscounter = 0
foreach ($user in $matchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering SharePoint Details for $($user.DisplayName_Source)"
    $UPN = $null
    #Gather Stats and Output
    Write-Host "Checking OneDrive for $($user.DisplayName_Source) .. " -foregroundcolor Cyan -NoNewline
    if ($UPN = $user.UserPrincipalName_Source) {
        # Gather SharePoint Details
        $count = 0
        $success = $null
        do{
            try{
                $OneDriveSite = Get-SPOSite -Filter "Owner -eq '$UPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop
                if ($OneDriveSite) {
                    $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $OneDriveSite.url -force
                    $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)_Source" -Value $OneDriveSite.StorageUsageCurrent -Force
                    $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)_Source" -Value ($OneDriveSite.StorageUsageCurrent/1024) -Force
                    $user | add-member -type noteproperty -name "OneDriveLastContentModified_Source" -Value $OneDriveSite.LastContentModifiedDate -force
                    Write-Host " Found" -foregroundcolor green
                    $success = $true
                }
                else {
                        $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $null -force
                        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)_Source" -Value $null -Force
                        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)_Source" -Value $null -Force
                        $user | add-member -type noteproperty -name "OneDriveLastContentModified_Source" -Value $null -force
                        Write-Host " Not Found" -foregroundcolor red
                        $failed = $true
                }
            }
            catch{
                Write-host "Next attempt in 30 seconds" -foregroundcolor yellow -nonewline
                Start-sleep -Seconds 30
                $count++
            }
        }
        until($count -eq 5 -or $success -or $failed)
        if(!($success -or $failed)) {
            $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $null -force
            $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)_Source" -Value $null -Force
            $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)_Source" -Value $null -Force
            $user | add-member -type noteproperty -name "OneDriveLastContentModified_Source" -Value $null -force
            Write-Host " Not Found" -foregroundcolor red
        }
    }
    else {
        $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $null -force
        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)_Source" -Value $null -Force
        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)_Source" -Value $null -Force
        $user | add-member -type noteproperty -name "OneDriveLastContentModified_Source" -Value $null -force
    }

}

##### Resource and Shared Mailbox START REGION #####
function New-AmetekRemoteMailboxes {
    param (
        [Parameter(Mandatory=$True,HelpMessage="Which Type to Create?")] [string] $MailboxType,
        [Parameter(Mandatory=$True,HelpMessage="What is the CSV File Path")] [string] $ImportCSV
    )
    $ImportCSVUsers = Import-CSV $ImportCSV

    $resources = $ImportCSVUsers | ?{$_.RecipientTypeDetails_Source -eq "RoomMailbox" -and $_.Migrate -ne "No"}
    $sharedMailboxes = $ImportCSVUsers | ?{$_.RecipientTypeDetails_Source -eq "SharedMailbox" -and $_.Migrate -ne "No"}
    $EnabledOjbect = @()
    $createdObject = @()
    $notCreatedObject = @()
    $alreadyExists = @()
    $password = ConvertTo-SecureString "@metek!2022$" -AsPlainText -Force

    if ($MailboxType -eq "Resource") {
        $progressref = ($resources).count
        $progresscounter = 0     
        foreach ($mailbox in $resources) {
            #Set Variables
            $sourceEmail = $mailbox.PrimarySmtpAddress_Source
            $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
            $destinationDisplayName = $mailbox.DisplayName_Destination
            

            #Progress Bar - Resources
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Resource $($destinationDisplayName)"

            Write-Host "Checking for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan

            if ($ADUserCheck = Get-ADUser -Filter {UserPrincipalName -eq $destinationEmail} -ea silentlycontinue) {
                Write-Host "AD User already exists. Checking for Remote Room Mailbox .. " -ForegroundColor Yellow -NoNewline
                if (!($mailboxCheck = Get-RemoteMailbox $destinationEmail -ea silentlycontinue))
                {
                    $remoteMailboxCreate = Enable-RemoteMailbox $ADUserCheck.DistinguishedName -Room -RemoteRoutingAddress $sourceEmail
                    Write-Host "Enabled RemoteMailbox Successfully." -ForegroundColor Green
                    $EnabledOjbect += $mailbox
                }
                else {
                    Write-Host "Already Exists." -ForegroundColor Green
                    $alreadyExists += $mailbox
                }
            }
            else {
                try {
                    $OUCheck = Get-OrganizationalUnit $mailbox.OU -ErrorAction stop
                    $nonUserMailboxOU = "OU=Resources,"+ $OUCheck.DistinguishedName
                    $remoteMailboxCreate = New-RemoteMailbox -Shared -OnPremisesOrganizationalUnit $nonUserMailboxOU -DisplayName $destinationDisplayName -UserPrincipalName $destinationEmail -RemoteRoutingAddress $sourceEmail -PrimarySmtpAddress $destinationEmail -Name $destinationDisplayName -Password $password -ErrorAction Stop
                    Write-Host "New User Created" -ForegroundColor Green
                    $createdObject += $mailbox
                }
                catch {
                    Write-Host "No OU $($nonUserMailboxOU) Found or specified in Import File" -ForegroundColor Red
                    $notCreatedObject += $mailbox
                }
                $nonUserMailboxOU = $null
            }
        }
    }
    if ($MailboxType -eq "Shared") {
        $progressref = ($sharedMailboxes).count
        $progresscounter = 0    
        foreach ($mailbox in $sharedMailboxes) {
            #Set Variables
            $sourceEmail = $mailbox.PrimarySmtpAddress_Source
            $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
            $destinationDisplayName = $mailbox.DisplayName_Destination

            #Progress Bar - Shared Mailboxes
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Shared Mailbox $($destinationDisplayName)"

            Write-Host "Checking for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
            if ($ADUserCheck = Get-ADUser -Filter {UserPrincipalName -eq $destinationEmail} -ea silentlycontinue) {
                Write-Host "AD User already exists. Checking for Shared Mailbox " -ForegroundColor Yellow -NoNewline
                if (!($mailboxCheck = Get-RemoteMailbox $destinationEmail -ea silentlycontinue)) {
                    $remoteMailboxCreate = Enable-RemoteMailbox $ADUserCheck.DistinguishedName -Shared -RemoteRoutingAddress $sourceEmail
                    Write-Host "Enabled Remote Mailbox Successfully." -ForegroundColor Green
                    $EnabledOjbect += $mailbox
                }
                else {
                    Write-Host "Already Exists." -ForegroundColor Green
                    $alreadyExists += $mailbox
                }
            }
            else {
                try {
                    $OUCheck = Get-OrganizationalUnit $mailbox.OU -ErrorAction stop
                    $nonUserMailboxOU = "OU=Shared Mailboxes,"+ $OUCheck.DistinguishedName
                    $remoteMailboxCreate = New-RemoteMailbox -Shared -OnPremisesOrganizationalUnit $nonUserMailboxOU -DisplayName $destinationDisplayName -UserPrincipalName $destinationEmail -RemoteRoutingAddress $sourceEmail -PrimarySmtpAddress $destinationEmail -Name $destinationDisplayName -Password $password -ErrorAction Stop
                    Write-Host "New User Created" -ForegroundColor Green
                    $createdObject += $mailbox
                }
                catch {
                    Write-Host "No OU $($nonUserMailboxOU) Found or specified in Import File" -ForegroundColor Red
                    $notCreatedObject += $mailbox
                }
                $nonUserMailboxOU = $null
                
            }
        }
    }

    Write-Host ""
    Write-Host "Results!" -foregroundcolor Cyan

    #Output
    if ($MailboxType -eq "Resource") {
        Write-Host $EnabledOjbect.count "Enabled Resources" -ForegroundColor Green
        Write-Host $createdObject.count "Created Resources" -ForegroundColor Cyan
        Write-Host $alreadyExists.count "Already Existing Resources" -ForegroundColor Gray
        Write-Host $notCreatedObject.count "Failed to Create Resources" -ForegroundColor Red
        $notCreatedObject | out-gridview
    }
    if ($MailboxType -eq "Shared") {
        Write-Host $EnabledOjbect.count "Enabled Shared Mailbox" -ForegroundColor Green
        Write-Host $createdObject.count "Created Shared Mailbox" -ForegroundColor Cyan
        Write-Host $alreadyExists.count "Already Existing Resources" -ForegroundColor Gray
        Write-Host $notCreatedObject.count "Failed to Create Shared Mailbox" -ForegroundColor Red
        $notCreatedObject | out-gridview
    }
}

## Set Employee ID 
$nonUserMailboxes = $matchedMailboxes | ?{$_.RecipientTypeDetails_Destination -ne "UserMailbox" -and $_.Migrate -ne "No"}
$progressref = ($nonUserMailboxes).count
$progresscounter = 0    

foreach ($mailbox in $nonUserMailboxes) {
    $sourceEmail = $mailbox.PrimarySmtpAddress_Source
    $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
    $destinationDisplayName = $mailbox.DisplayName_Destination
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating EmployeeID $($destinationDisplayName)"

    $count = 0
    $success = $null
    Write-host "Setting Employee ID to Resource $($destinationDisplayName)" -foregroundcolor cyan -nonewline
    try{
        $mailboxCheck = Get-RemoteMailbox $destinationEmail -ErrorAction Stop
        if ($mailboxCheck) {
            Set-ADUser -Identity $mailboxCheck.DistinguishedName -EmployeeID Resource
            $success = $true
            Write-host "done" -foregroundcolor green
        }
    }
    catch{
        Write-host "Next attempt in 10 seconds" -foregroundcolor yellow -nonewline
        Start-sleep -Seconds 10
        $count++
    }
    do{
        try{
            $mailboxCheck = Get-RemoteMailbox $destinationEmail -ErrorAction Stop
            if ($mailboxCheck) {
                Set-ADUser -Identity $mailboxCheck.DistinguishedName -EmployeeID Resource
                $success = $true
                Write-host "done" -foregroundcolor green
            }
        }
        catch{
            Write-host "Next attempt in 10 seconds" -foregroundcolor yellow -nonewline
            Start-sleep -Seconds 10
            $count++
        }
        
}
until($count -eq 6 -or $success -or $failed) #>
}


#Create Permission Mail Enabled Security Groups for non user mailboxes
$matchedMailboxes = Import-Csv
$nonUserMailboxes = $matchedMailboxes | ?{$_.RecipientTypeDetails_Destination -ne "UserMailbox" -and $_.Migrate -ne "No"}
$progressref = ($nonUserMailboxes).count
$progresscounter = 0
$failures = @()
foreach ($mailbox in $nonUserMailboxes) {
    #Set Variables
    $sourceEmail = $mailbox.PrimarySmtpAddress_Source
    $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
    $destinationDisplayName = $mailbox.DisplayName_Destination
    $OUCheck = Get-OrganizationalUnit $mailbox.OU
    $addressSplit = $destinationEmail -split "@"
    $FullAccessResourceName = $destinationDisplayName + "_FullAccess"
    $SendAsResourceName = $destinationDisplayName + "_SendAs"
    $FullAccessResourceEmailAddress = $addressSplit[0] + "_FullAccess@ametek.com"
    $SendAsResourceEmailAddress= $addressSplit[0] + "_SendAs@ametek.com"

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Create Perm Groups for Non-UserMailbox $($destinationDisplayName)"
 
    #Create Perm Groups for Resources
    if ($mailbox.RecipientTypeDetails_Destination -eq 'RoomMailbox') {
        Write-Host "Creating Resource Permissions Groups $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
        $resourceOU = "OU=Resources,"+ $OUCheck.DistinguishedName

        #add full access group
        try {
            #Add DL Members
            $result = New-DistributionGroup -Type security -Name $FullAccessResourceName -DisplayName $FullAccessResourceName -PrimarySMTPAddress $FullAccessResourceEmailAddress -OrganizationalUnit $resourceOU -ErrorAction Stop
            Write-Host "done " -ForegroundColor Green
        }
        catch {
            Write-Host "Full Access Group Failed. Possibly Already Exists " -ForegroundColor Yellow
            
            #Build Error Array
            $currenterror = new-object PSObject

            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
            $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
            $currenterror | add-member -type noteproperty -name "DistributionList" -Value $dFullAccessResourceEmailAddress
            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
            $failures += $currenterror
        }
        #add Send As access group
        try {
            #Add DL Members
            $result = New-DistributionGroup -Type security -Name $SendAsResourceName -DisplayName $SendAsResourceName -PrimarySMTPAddress $SendAsResourceEmailAddress -OrganizationalUnit $resourceOU -ErrorAction Stop
            Write-Host "done " -ForegroundColor Green
        }
        catch {
            Write-Host "Send As Group Failed. Possibly Already Exists " -ForegroundColor Yellow
            
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

    #Create Perm Groups for Shared Mailboxes
    if ($mailbox.RecipientTypeDetails_Destination -eq 'SharedMailbox') {
        Write-Host "Creating Shared Mailbox Permissions Groups $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
        $sharedMailboxOU = "OU=Shared Mailboxes,"+ $OUCheck.DistinguishedName
        
        #add full access group
        try {
            #Add DL Members
            $result = New-DistributionGroup -Type security -Name $FullAccessResourceName -DisplayName $FullAccessResourceName -PrimarySMTPAddress $FullAccessResourceEmailAddress -OrganizationalUnit $sharedMailboxOU -ErrorAction Stop
            Write-Host "done " -ForegroundColor Green -NoNewline
        }
        catch {
            Write-Host "Full Access Group Failed. Possibly Already Exists " -ForegroundColor Yellow
            
            #Build Error Array
            $currenterror = new-object PSObject

            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
            $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
            $currenterror | add-member -type noteproperty -name "DistributionList" -Value $dFullAccessResourceEmailAddress
            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
            $failures += $currenterror
        }
        #add Send As access group
        try {
            #Add DL Members
            $result = New-DistributionGroup -Type security -Name $SendAsResourceName -DisplayName $SendAsResourceName -PrimarySMTPAddress $SendAsResourceEmailAddress -OrganizationalUnit $sharedMailboxOU -ErrorAction Stop
            Write-Host "done " -ForegroundColor Green -NoNewline
        }
        catch {
            Write-Host "Send As Group Failed. Possibly Already Exists " -ForegroundColor Yellow
            
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
$failures | Out-GridView

## Create DistributionGroups for Abaco
$Groups = import-csv
$createdDistributionGroups = @()
$createdADGroups = @()
$progressref = ($groups).count
$progresscounter = 0
foreach ($group in $groups) {
   $sourceEmail = $group.PrimarySmtpAddress_Source
   $destinationEmail = $group.PrimarySMTPAddress_Destination
   $destinationDisplayName = $group.DisplayName_Destination
   $progresscounter += 1
   $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
   $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
   Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Distribution Group $($destinationDisplayName)"

    Write-Host "Checking for in AD for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    if ($adGroupCheck = Get-ADGroup -Filter {Mail -eq $destinationEmail} -ea silentlycontinue) {
        Write-Host "AD User already exists. Checking for Distribution List " -ForegroundColor Yellow
        if (!($DLCheck = Get-DistributionGroup $destinationEmail -ea silentlycontinue)) {
            Enable-DistributionGroup $adGroupCheck.DistinguishedName -PrimarySMTPAddress $destinationEmail -DisplayName $displayName -alias $group.alias
            Write-Host "Created Successfully." -ForegroundColor Green
            $createdDistributionGroups += $group
        }
        else {
            Write-Host "Already Exists." -ForegroundColor Yellow
        }
    }
    else {
        $OUCheck = Get-OrganizationalUnit $group.OU
        $distributionListOU = "OU=Distribution Lists,"+ $OUCheck.DistinguishedName
        $createdADGroups += $group
        if ($distributionListOU) {
            try {
                New-DistributionGroup -OrganizationalUnit $distributionListOU -DisplayName $destinationDisplayName -name $destinationDisplayName -PrimarySMTPAddress $destinationEmail -ManagedBy $group.ManagedBy -ErrorAction Stop
                Write-Host "New Group Created" -ForegroundColor Green
            }
            catch {
                New-DistributionGroup -OrganizationalUnit $distributionListOU -DisplayName $destinationDisplayName -name $destinationDisplayName -PrimarySMTPAddress $destinationEmail
            }
        }
        else {
            Write-Host "No OU found for Group" -ForegroundColor red
        }
    }  
}

## Set Attributes DistributionGroups for Abaco
$Groups = import-csv
$progressref = ($groups).count
$progresscounter = 0
$updatedDistributionGroups = @()
foreach ($group in $groups) {
   $sourceEmail = $group.PrimarySmtpAddress_Source
   $destinationEmail = $group.PrimarySMTPAddress_Destination
   $destinationDisplayName = $group.DisplayName_Destination
   $progresscounter += 1
   $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
   $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
   Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Distribution Group $($destinationDisplayName)"

    Write-Host "Checking for in AD for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    if ($adGroupCheck = Get-ADGroup -Filter {Mail -eq $destinationEmail} -ea silentlycontinue) {
        if ($DLCheck = Get-DistributionGroup $destinationEmail -ea silentlycontinue) {
            [boolean]$HiddenFromAddressListsEnabled = [boolean]::Parse($group.HiddenFromAddressListsEnabled)
            [boolean]$RequireSenderAuthenticationEnabled = [boolean]::Parse($group.RequireSenderAuthenticationEnabled)
            Set-DistributionGroup $adGroupCheck.DistinguishedName -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -MemberJoinRestriction $group.MemberJoinRestriction -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -warningaction silentlycontinue
            Write-Host "Updated Successfully." -ForegroundColor Green
            
            #addEmail Addresses
            $emailAddresses = $group.EmailAddresses -split ","
            foreach ($address in $emailAddresses) {
                if ($address -like "*x500*") {
                    $sourceAddress = $address
                }
                else {
                    $smtpAddress = $address -Split ":"
                    $sourceAddress = $smtpAddress[1]
                }
                Set-DistributionGroup  -EmailAddresses @{add=$sourceAddress} -Identity $destinationEmail -warningaction silentlycontinue
            }
            $updatedDistributionGroups += $group
        }
        else {
            Write-Host "Already Exists." -ForegroundColor Yellow
        }
    }
    else {
        $OUCheck = Get-OrganizationalUnit $group.OU
        $distributionListOU = "OU=Distribution Lists,"+ $OUCheck.DistinguishedName
        $createdADGroups += $group
        if ($distributionListOU) {
            try {
                New-DistributionGroup -OrganizationalUnit $distributionListOU -DisplayName $destinationDisplayName -name $destinationDisplayName -PrimarySMTPAddress $destinationEmail -ManagedBy $group.ManagedBy -ErrorAction Stop
                Write-Host "New Group Created" -ForegroundColor Green
            }
            catch {
                New-DistributionGroup -OrganizationalUnit $distributionListOU -DisplayName $destinationDisplayName -name $destinationDisplayName -PrimarySMTPAddress $destinationEmail
            }
        }
        else {
            Write-Host "No OU found for Group" -ForegroundColor red
        }
    }  
}

## Add Members to DistributionGroups for Abaco
$Groups = import-csv
$createdADGroups = @()
$progressref = ($groups).count
$progresscounter = 0
foreach ($group in $groups) {
    $sourceEmail = $group.PrimarySmtpAddress_Source
    $destinationEmail = $group.PrimarySMTPAddress_Destination
    $destinationDisplayName = $group.DisplayName_Destination
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Group Members for Distribution Group $($destinationDisplayName)"

    Write-Host "Checking for $($destinationDisplayName) in AD ... " -NoNewline -ForegroundColor Cyan
    if ($adGroupCheck = Get-ADGroup -Filter {Mail -eq $destinationEmail} -ea silentlycontinue) {
        if ($DLCheck = Get-DistributionGroup $destinationEmail -ea silentlycontinue) {
            $GroupMembers = $group.Members -split ","
            $progressref2 = ($GroupMembers).count
            $progresscounter2 = 0
            Write-Host "Adding $($GroupMembers.count) Members .. " -NoNewline
            foreach ($perm in $GroupMembers) {
                # Match the Perm user
                $trimPermUser = $perm.trim()
                $progresscounter2 += 1
                $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Full Access Perms to $($trimPermUser)"

                if ($matchedUser = $matchedUsers | Where-Object {$_.PrimarySmtpAddress_Source -eq $trimPermUser}) {  
                    $destinationEmail = $matchedUser.PrimarySMTPAddress_Destination
                    ## Check if Perm User Exists
                    if ($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue) {
                        try {
                            #Add DL Members
                            Add-DistributionGroupMember $destinationDisplayName -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                            Write-Host "." -ForegroundColor Green -NoNewline
                        }
                        catch {
                            Write-Host "." -ForegroundColor Red -NoNewline
                            
                            #Build Error Array
                            $currenterror = new-object PSObject

                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                            $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                            $currenterror | add-member -type noteproperty -name "DistributionList" -Value $destinationEmail
                            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                            $failures += $currenterror
                        }
                    } 
                }
                elseif ($matchedUser = $groups | Where-Object {$_.PrimarySmtpAddress_Source -eq $trimPermUser}) {
                    $destinationEmail = $matchedUser.PrimarySMTPAddress_Destination
                    ## Check if Perm User Exists
                    if ($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue) {
                        try {
                            #Add DL Members
                            Add-DistributionGroupMember $destinationDisplayName -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                            Write-Host "." -ForegroundColor Green -NoNewline
                        }
                        catch {
                            Write-Host "." -ForegroundColor Red -NoNewline
                            
                            #Build Error Array
                            $currenterror = new-object PSObject

                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                            $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                            $currenterror | add-member -type noteproperty -name "DistributionList" -Value $destinationEmail
                            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                            $failures += $currenterror
                        }
                    } 
                }
                else {
                    Write-Host "Not Matched Recipient $($trimPermUser) .. " -ForegroundColor red -NoNewline
                    $notFoundPermUser += $trimPermUser
                }
            }
        }
        else {
            Write-Host "Group Is not Enabled for Exchange." -ForegroundColor Yellow
        }
    }
    #Create Group if it does not exist
    else {
        $OUCheck = Get-OrganizationalUnit $group.OU
        $distributionListOU = "OU=Distribution Lists,"+ $OUCheck.DistinguishedName
        $createdADGroups += $group
        if ($distributionListOU) {
            try {
                New-DistributionGroup -OrganizationalUnit $distributionListOU -DisplayName $destinationDisplayName -name $destinationDisplayName -PrimarySMTPAddress $destinationEmail -ManagedBy $group.ManagedBy -ErrorAction Stop
                Write-Host "New Group Created" -ForegroundColor Green
            }
            catch {
                New-DistributionGroup -OrganizationalUnit $distributionListOU -DisplayName $destinationDisplayName -name $destinationDisplayName -PrimarySMTPAddress $destinationEmail
            }
        }
        else {
            Write-Host "No OU found for Group" -ForegroundColor red
        }
    }
    Write-Host "Done"
}

#Create New Office365 Groups for abaco to ametek
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
$sourceMigrationAccount = "SA-Arraya01@abaco1.onmicrosoft.com"
$destinationMigrationAccount = "ArrayaMigration@AMETEKInc.onmicrosoft.com"
$migratingDomain = "@abaco.com"

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

#Update Visible in GAL Permission Mail Enabled Security Groups for non user mailboxes
$matchedMailboxes = Import-Csv
$nonUserMailboxes = $matchedMailboxes | ?{$_.RecipientTypeDetails_Destination -ne "UserMailbox" -and $_.Migrate -ne "No"}
$progressref = ($nonUserMailboxes).count
$progresscounter = 0
foreach ($mailbox in $nonUserMailboxes) {
    #Set Variables
    $sourceEmail = $mailbox.PrimarySmtpAddress_Source
    $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
    $destinationDisplayName = $mailbox.DisplayName_Destination
    $addressSplit = $destinationEmail -split "@"
    $FullAccessResourceName = $addressSplit[0] + "_FullAccess@" + $addressSplit[1]
    $SendAsResourceName = $addressSplit[0] + "_SendAs@" + $addressSplit[1]

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Update Visible In GAL For Perm Groups of $($destinationDisplayName)"
 
    #Update Visible in GAL Permission Mail Enabled Security Groups for resource mailboxes
    if ($mailbox.RecipientTypeDetails_Destination -eq 'RoomMailbox') {
        Write-Host "Update Resource Visible in GAL Permissions Groups $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
        Set-DistributionGroup -Identity $FullAccessResourceName -HiddenFromAddressListsEnabled $True -WarningAction silentlycontinue
        Set-DistributionGroup -Identity $SendAsResourceName -HiddenFromAddressListsEnabled $True -WarningAction silentlycontinue
        
        Write-Host "done" -ForegroundColor Green
    }

    #Update Visible in GAL Permission Mail Enabled Security Groups for shared mailboxes
    if ($mailbox.RecipientTypeDetails_Destination -eq 'SharedMailbox') {
        Write-Host "Update Shared Mailbox Visible in GAL Permissions Groups $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan        
        Set-DistributionGroup -Identity $FullAccessResourceName -HiddenFromAddressListsEnabled $True -WarningAction silentlycontinue
        Set-DistributionGroup -Identity $SendAsResourceName -HiddenFromAddressListsEnabled $True -WarningAction silentlycontinue
    
        Write-Host "done" -ForegroundColor Green
    }
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

#Match Mailboxes to Ametek (recheck)
$allMatchedMailboxes = Import-Excel "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco-Ametek-MatchingReport.xlsx" -WorksheetName "AllMatchedMailboxes"

$progressref = $allMatchedMailboxes.count
$progresscounter = 0
foreach ($user in $allMatchedMailboxes) {
    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Stamping Perms on UserMailbox $($user.PrimarySmtpAddress_Destination)"
    if ($mailboxCheck = Get-Mailbox $user.PrimarySmtpAddress_Destination -ea SilentlyContinue) {
        $user | add-member -type noteproperty -name "DisplayName_Check" -Value $mailboxCheck.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Check" -Value $mailboxCheck.UserPrincipalName -Force
        $user | add-member -type noteproperty -name "PrimarySMTPAddress_Check" -Value $mailboxCheck.PrimarySMTPAddress -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Check" -Value $mailboxCheck.RecipientTypeDetails -force
    }
    else {
        $user | add-member -type noteproperty -name "DisplayName_Check" -Value $null -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Check" -Value $null -Force
        $user | add-member -type noteproperty -name "PrimarySMTPAddress_Check" -Value $null -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Check" -Value $null -force
    }
}

#Match Recipients to Ametek (recheck)
$allMatchedGroups = Import-Excel -WorksheetName "Office365 Groups" -Path "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco-Ametek-MatchingReport.xlsx"

$progressref = $allMatchedGroups.count
$progresscounter = 0
foreach ($user in $allMatchedGroups) {
    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Recipient $($user.Mock_PrimarySMTPAddress)"
    if ($user.Mock_PrimarySMTPAddress -ne "") {
        if ($recipientCheck = Get-Recipient $user.Mock_PrimarySMTPAddress -ea SilentlyContinue) {

            $user | add-member -type noteproperty -name "DisplayName_Check" -Value $recipientCheck.DisplayName -force
            $user | add-member -type noteproperty -name "PrimarySMTPAddress_Check" -Value $recipientCheck.PrimarySMTPAddress -Force
            $user | add-member -type noteproperty -name "RecipientTypeDetails_Check" -Value $recipientCheck.RecipientTypeDetails -force
        }
        else {
            $user | add-member -type noteproperty -name "DisplayName_Check" -Value $null -force
            $user | add-member -type noteproperty -name "PrimarySMTPAddress_Check" -Value $null -Force
            $user | add-member -type noteproperty -name "RecipientTypeDetails_Check" -Value $null -force
        }
    }
    
}

#Match Current Employees to Matched Mailboxes
$allMatchedMailboxes = Import-Excel "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco-Ametek-MatchingReport.xlsx" -WorksheetName "AllMatchedMailboxes"

$UserDetails  = Import-Excel 
$progressref = $allMatchedMailboxes.count
$progresscounter = 0
foreach ($user in $allMatchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Stamping Perms on UserMailbox $($user.PrimarySmtpAddress_Source)"

    if ($matchedUser = $UserDetails | ? {$_."Email Address" -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "Current_Employee" -Value $true -force
    }
    else {
        $user | add-member -type noteproperty -name "Current_Employee" -Value $false -force
    }
}

#match old match stats
$allMatchedMailboxes = Import-Excel -WorksheetName "AllMatchedMailboxes" -Path "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco-Ametek-MatchingReport.xlsx"
$oldMatchedMailbox  = Import-Excel -WorksheetName "MatchedMailboxes_old" -path "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco-Ametek-MatchingReport.xlsx"
$progressref = $allMatchedMailboxes.count
$progresscounter = 0
foreach ($user in $allMatchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Old Details UserMailbox $($user.PrimarySmtpAddress_Source)"

    if ($matchedUser = $oldMatchedMailbox | ? {$_.PrimarySmtpAddress_Source -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $matchedUser.OneDriveUrl_Source -force
        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)_Source" -Value $matchedUser."OneDriveCurrentStorage(MB)_Source" -force
        $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)_Source" -Value $matchedUser."OneDriveCurrentStorage(GB)_Source" -force
    }
}

#Update old User's Details
$updatedMatchedMailboxDetails = Import-Excel -WorksheetName "Master List2" -path 'c:\Users\amedrano\Arraya Solutions\Thomas Jefferson - Einstein to Jefferson Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$updatedMatchedMailboxDetails = import-csv
$oldMatchedMailboxDetails  = Import-Excel -path 'c:\Users\amedrano\Arraya Solutions\Thomas Jefferson - Einstein to Jefferson Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$progressref = $updatedMatchedMailboxDetails.count
$progresscounter = 0
foreach ($user in $updatedMatchedMailboxDetails) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details UserMailbox $($user.PrimarySmtpAddress_Source)"

    if ($matchedUser = $oldMatchedMailboxDetails | ? {$_.PrimarySmtpAddress_Source -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "Migrate" -Value $matchedUser.Migrate -force
        $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $matchedUser.OneDriveUrl_Source -force
        $user | add-member -type noteproperty -name "OneDriveLastContentModified_Source" -Value $matchedUser.OneDriveLastContentModified_Source -force
        $user | add-member -type noteproperty -name "OU" -Value $matchedUser.OU -force
    }
    else {
        $user | add-member -type noteproperty -name "Migrate" -Value $null -force
        $user | add-member -type noteproperty -name "OneDriveUrl_Source" -Value $null -force
        $user | add-member -type noteproperty -name "OneDriveLastContentModified_Source" -Value $null -force
        $user | add-member -type noteproperty -name "OU" -Value $null -force
    }
}


$allMatchedMailboxes = Export-Excel "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco-Ametek-MatchingReport.xlsx" -WorksheetName "AllMatchedMailboxes" -Show

#Start Batch Cutover
Connect-AzureAD
Connect-MsolService
Connect-ExchangeOnline

#Examples
# Update Towcester Pilot Users - Remove Forward in Destination Tenant (Ametek)
Start-BatchCutoverUpdates -ImportCSV "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco_PilotUsers.csv" -Location Towcester -Tenant Destination -RemoveForward

# Update Towcester Pilot Users - Set Forward in Source Tenant (Abaco)
Start-BatchCutoverUpdates -ImportCSV "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco_PilotUsers.csv" -Location Towcester -Tenant Source -SetForward

# Update Towcester Pilot Users - Block Email Access and Sign Users Out in Source Tenant (Abaco)
Start-BatchCutoverUpdates -ImportCSV "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco_PilotUsers.csv" -Location Towcester -Tenant Source -BlockMailAccess -ForceLogOut

# Update Towcester Pilot Users - Block Email Access, Sign Users Out, Set Forward, Block OneDrive Access in Source Tenant (Abaco)
Start-BatchCutoverUpdates -ImportCSV "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco_PilotUsers.csv" -Location Towcester -Tenant Source -SetForward -BlockMailAccess -ForceLogOut -BlockOneDriveAccess

function Start-BatchCutoverUpdates {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the CSV File Path of OneDrive Users')] [array] $ImportCSV,
        [Parameter(Mandatory=$True,HelpMessage="Which Tenant to Update (Source or Destination?")] [string] $Tenant,
        [Parameter(ParameterSetName='Source',Mandatory=$false,HelpMessage="Set Forward?")] [switch] $SetForward,
        [Parameter(ParameterSetName='Source',Mandatory=$false,HelpMessage="Block Access to Mailbox?")] [switch] $BlockMailAccess,
        [Parameter(ParameterSetName='Source',Mandatory=$false,HelpMessage="Block Access to OneDrive?")] [switch] $BlockOneDriveAccess,
        [Parameter(ParameterSetName='Source',Mandatory=$false,HelpMessage="Disable Teams?")] [switch] $DisableTeams,
        [Parameter(Mandatory=$false,HelpMessage="Which Location? Austin,Huntsville, or Towcester")] [string] $Location,
        [Parameter(ParameterSetName='Destination',Mandatory=$false,HelpMessage="Remove Forward in Destination Account?")] [switch] $RemoveForward,
        [Parameter(Mandatory=$false,HelpMessage="Do you want to Test?")] [switch] $Test,
        [Parameter(ParameterSetName='Source',Mandatory=$false,HelpMessage="Force Log Out User?")] [switch] $ForceLogOut
    )
    #Create User Array
    $WaveGroupImport = Import-csv $ImportCSV
    <#if ($Location) {
        if ($Location -eq "Austin") {
            $WaveGroup = $WaveGroupImport | ?{$_.Location -eq "Austin"}
        }
        elseif ($Location -eq "Huntsville") {
            $WaveGroup = $WaveGroupImport | ?{$_.Location -eq "Huntsville"}
        }
        elseif ($Location -eq "Towcester") {
            $WaveGroup = $WaveGroupImport | ?{$_.Location -eq "Towcester"}
        }
        else {
            Write-Error "No Location Specified. Please re-run script and specify Austin,Huntsville, or Towcester"
            return
        }
    #>

    # Gather User Details
    $progressref = ($WaveGroupImport).count
    $progresscounter = 0
    $allErrors = @()
        
    foreach ($user in $WaveGroupImport) {
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($user.DisplayName_Source)"
        $SourceUPN = $user.UserPrincipalName_Source
        $SourcePrimarySMTPAddress = $user.PrimarySMTPAddress_Source
        $DesinationPrimarySMTPAddress = $user.PrimarySmtpAddress_Destination

        if ($Test) {
            if ($Tenant -eq "Destination") {
                # Jefferson
                Write-Host "Cutting Over User $($DesinationPrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
                ## Hide Einstein Contact In Jefferson
                if ($RemoveForward) {
                    try {
                        Set-Mailbox $DesinationPrimarySMTPAddress -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false -whatif
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
    
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "RemoveDestinationForward" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
            }
            elseif ($Tenant -eq "Source") {
                # Einstein
                Write-Host "Cutting Over User $($SourceUPN) .. " -foregroundcolor Cyan -nonewline
                ## Set Mailbox to Forward from Source to Destination Mailbox and Leave a Copy
                if ($SetForward) {
                    Write-Host "Set Forward on Mailbox ..  " -foregroundcolor Magenta -nonewline
                    Try{ 
                        Set-Mailbox $SourcePrimarySMTPAddress -DeliverToMailboxAndForward $true -ForwardingSmtpAddress $DesinationPrimarySMTPAddress -ErrorAction Stop -whatif
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    Catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SetForward" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }     
                ## Block OWA, Outlook Access
                if ($BlockMailAccess) {
                    Write-Host "Block Access To Mailbox ..  " -foregroundcolor DarkGray -nonewline
                    Try{ 
                        Set-CASMailbox $SourcePrimarySMTPAddress -MacOutlookEnabled $false -MAPIEnabled $false -OutlookMobileEnabled $false -OWAEnabled $false -ActiveSyncEnabled $false -OWAforDevices $false -IMAPEnabled $false -POPEnabled $false -ErrorAction Stop -whatif
                        if ((Get-Recipient $SourcePrimarySMTPAddress).RecipientTypeDetails -eq "UserMailbox") {
                            Set-Mailbox -Identity $SourcePrimarySMTPAddress -AccountDisabled:$True -whatif
                        }
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    Catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "BlockEmailAccess" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
                # Block OneDrive Access
                if ($BlockOneDriveAccess) {
                    $UPN = $SourceUPN
                    $SPOUPN = $UPN.replace("@abaco.com","_abaco_com")
                    $SPOSITE = Get-SPOSITE -IncludePersonalSite $true -filter "Url -like '-my.sharepoint.com/personal/$SPOUPN'" -ErrorAction SilentlyContinue
    
                    Write-Host "Blocking OneDrive Access .. " -foregroundcolor DarkCyan -nonewline
    
                    ### Remove User as Site Admin
                    try {
                        Write-Host ". " -foregroundcolor Green -nonewline
                        $adminRequest = Set-SPOUser -Site $SPOSITE.url -LoginName $UPN -IsSiteCollectionAdmin $false -ErrorAction Stop -whatif
                        #Remove-SPOUser -Site $SPOSITE.url -LoginName $UPN -confirm:$false
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
    
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "RemoveOneDriveAccess" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
                #Disable Teams
                if ($DisableTeams) {
                    Write-Host "Disabling Teams .. " -foregroundcolor Magenta -nonewline
                    $msoluser = Get-MsolUser -UserPrincipalName $SourceUPN 
                    $DisabledArray = @()
                    $allLicenses = ($msoluser).Licenses
                    $SKUID = ($msoluser).licenses.AccountSkuId | ?{$_ -like "*SPE_*"}
                    #$SKUID = "EHN:ENTERPRISEPACK"
                
                    ## Get Disabled Array
                    for($i = 0; $i -lt $AllLicenses.Count; $i++) {
                        $serviceStatus =  $AllLicenses[$i].ServiceStatus
                        foreach($service in $serviceStatus)
                        {
                            if($service.ProvisioningStatus -eq "Disabled") {
                                $disabledArray += ($service.ServicePlan).ServiceName
                            }
                        }
                    }
                
                    #Update users with Office E3 licenses to Microsoft E3 licenses with DisabledArray above.
                    #add Teams to DisabledArray
                    $disabledArray += "Teams1"
                    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $SKUID -DisabledPlans $disabledArray
                    
                    ### Update License and Disable Teams
                    try {
                        Set-MsolUserLicense -UserPrincipalName $SourceUPN -LicenseOptions $LicenseOptions -ErrorAction Stop -whatif
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "DisableTeams" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
                ## Block OWA, Outlook Access
                if ($ForceLogOut) {
                    Write-Host "Booting User's Current Sessions ..  " -foregroundcolor DarkGray -nonewline
                    Try{ 
                        Get-AzureADUser -SearchString $SourceUPN | Revoke-AzureADUserAllRefreshToken -whatif
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    Catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "BootUser" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
                Write-Host "Done" -foregroundcolor Green 
            }
        }
        else {
            if ($Tenant -eq "Destination") {
                # Jefferson
                Write-Host "Cutting Over User $($DesinationPrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
                ## Hide Einstein Contact In Jefferson
                if ($RemoveForward) {
                    try {
                        Set-Mailbox $DesinationPrimarySMTPAddress -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
    
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "RemoveDestinationForward" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
            }
            elseif ($Tenant -eq "Source") {
                # Einstein
                Write-Host "Cutting Over User $($SourceUPN) .. " -foregroundcolor Cyan -nonewline
                ## Set Mailbox to Forward from Source to Destination Mailbox and Leave a Copy
                if ($SetForward) {
                    Write-Host "Set Forward on Mailbox ..  " -foregroundcolor Magenta -nonewline
                    Try{ 
                        Set-Mailbox $SourcePrimarySMTPAddress -DeliverToMailboxAndForward $true -ForwardingSmtpAddress $DesinationPrimarySMTPAddress -ErrorAction Stop
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    Catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SetForward" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }     
                ## Block OWA, Outlook Access
                if ($BlockMailAccess) {
                    Write-Host "Block Access To Mailbox ..  " -foregroundcolor DarkGray -nonewline
                    Try{ 
                        Set-CASMailbox $SourcePrimarySMTPAddress -MacOutlookEnabled $false -MAPIEnabled $false -OutlookMobileEnabled $false -OWAEnabled $false -ActiveSyncEnabled $false -OWAforDevices $false -IMAPEnabled $false -POPEnabled $false -ErrorAction Stop
                        Set-Mailbox -Identity $SourcePrimarySMTPAddress -AccountDisabled:$True
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    Catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "BlockEmailAccess" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
                # Block OneDrive Access
                if ($BlockOneDriveAccess) {
                    $UPN = $SourceUPN
                    $SPOUPN = $UPN.replace("@abaco.com","_abaco_com")
                    $SPOSITE = Get-SPOSITE -IncludePersonalSite $true -filter "Url -like '-my.sharepoint.com/personal/$SPOUPN'" -ErrorAction SilentlyContinue
    
                    Write-Host "Blocking OneDrive Access .. " -foregroundcolor DarkCyan -nonewline
    
                    ### Remove User as Site Admin
                    try {
                        Write-Host ". " -foregroundcolor Green -nonewline
                        $adminRequest = Set-SPOUser -Site $SPOSITE.url -LoginName $UPN -IsSiteCollectionAdmin $false -ErrorAction Stop
                        #Remove-SPOUser -Site $SPOSITE.url -LoginName $UPN -confirm:$false
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
    
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "RemoveOneDriveAccess" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
                #Disable Teams
                if ($DisableTeams) {
                    Write-Host "Disabling Teams .. " -foregroundcolor Magenta -nonewline
                    $msoluser = Get-MsolUser -UserPrincipalName $SourceUPN
                    $DisabledArray = @()
                    $allLicenses = ($msoluser).Licenses
                    $SKUID = ($msoluser).licenses.AccountSkuId | ?{$_ -like "*SPE_*"}
                    #$SKUID = "EHN:ENTERPRISEPACK"
                
                    ## Get Disabled Array
                    for($i = 0; $i -lt $AllLicenses.Count; $i++) {
                        $serviceStatus =  $AllLicenses[$i].ServiceStatus
                        foreach($service in $serviceStatus)
                        {
                            if($service.ProvisioningStatus -eq "Disabled") {
                                $disabledArray += ($service.ServicePlan).ServiceName
                            }
                        }
                    }
                
                    #Update users with Office E3 licenses to Microsoft E3 licenses with DisabledArray above.
                    #add Teams to DisabledArray
                    $disabledArray += "Teams1"
                    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $SKUID -DisabledPlans $disabledArray
                    
                    ### Update License and Disable Teams
                    try {
                        Set-MsolUserLicense -UserPrincipalName $SourceUPN -LicenseOptions $LicenseOptions -ErrorAction Stop
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "DisableTeams" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
                ## Block OWA, Outlook Access
                if ($ForceLogOut) {
                    Write-Host "Booting User's Current Sessions ..  " -foregroundcolor DarkGray -nonewline
                    Try{ 
                        Get-AzureADUser -SearchString $SourceUPN | Revoke-AzureADUserAllRefreshToken
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    Catch {
                        Write-Host ". " -ForegroundColor red -NoNewline
                
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "BootUser" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $user.DisplayName_Source -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourceUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $SourcePrimarySMTPAddress -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $allErrors += $currenterror           
                        continue
                    }
                }
            }
        }
        Write-Host "Done" -foregroundcolor Green 
    }
}

#add abaco address prefix
$notfound = @()
$progressref = ($matchedMailboxes).count
$progresscounter = 0
foreach ($mailbox in $matchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating User $($mailbox.PrimarySmtpAddress_Source)"
    $oldaddressSplit = $mailbox.PrimarySmtpAddress_Source -split "@"
    $newMailboxAddress = $oldaddressSplit[0] + "@ametek.com"

    if ($mbxCheck = Get-RemoteMailbox $mailbox.PrimarySMTPAddress_Destination -ErrorAction SilentlyContinue) {
        Set-RemoteMailbox $mbxCheck.DistinguishedName -EmailAddresses @{add=$newMailboxAddress}
        Write-Host "Mailbox Updated for $($mailbox.PrimarySMTPAddress_Destination)" -foregroundcolor Green
    }
    elseif ($mailUserCheck = Get-Mailuser $mailbox.PrimarySMTPAddress_Destination -ErrorAction SilentlyContinue) {
        Set-MailUser $mailUserCheck.DistinguishedName -EmailAddresses @{add=$newMailboxAddress}
        Write-Host "MailUser Updated for $($mailbox.PrimarySMTPAddress_Destination)" -foregroundcolor Green
    }
    else {
        Write-Host "no Mailbox found for $($mailbox.PrimarySMTPAddress_Destination)" -foregroundcolor red
        $notfound += $mailbox
    }
}


# Remove Migration and Aaron's Account as Owner DistributionGroups - Ametek
$allDistributionGroups = Get-DistributionGroup -ResultSize unlimited
$progressref = ($allDistributionGroups).count
$progresscounter = 0
foreach ($group in $allDistributionGroups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Remove Migration Service Account to $($group.DisplayName)"

    $managedBy = $group.ManagedBy
    if ($managedBy -contains "Aaron Medrano") {
        try{
            Set-DistributionGroup -Identity $group.PrimarySmtpAddress -ManagedBy @{add="ArrayaMigration@AMETEKInc.onmicrosoft.com"}
            Set-DistributionGroup -Identity $group.PrimarySmtpAddress -ManagedBy @{remove="Aaron Medrano"} -ea Stop
            Write-Host "Removed Aaron Account" -foregroundcolor green
        }
        Catch {
            Write-Error "Unable To Remove Aaron Account"
        }
    }
    else {
        Write-Host "Migration Account Only Manager" -foregroundcolor Yellow
    }
}

#Gather OneDrive Details
$SourceOneDriveDetails = import-csv "C:\Users\amedrano\Downloads\AbacoOneDriveUsageAccountDetail4_18_2022 2_13_28 PM.csv"
$DestinationOneDriveDetails = import-csv "C:\Users\amedrano\Downloads\AmetekOneDriveUsageAccountDetail4_18_2022 2_14_11 PM.csv"
$allmatchedMailboxes = import-excel "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Abaco-Ametek-MatchingReport.xlsx"
$progressref = $allmatchedMailboxes.count
$progresscounter = 0
foreach ($mailbox in $allmatchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details $($mailbox.UserPrincipalName_Source)"
    #gatherOneDriveDetails - Abaco
    if ($matchedSource = $SourceOneDriveDetails | ?{$_."Owner Principal Name" -eq $mailbox.UserPrincipalName_Source}) {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Source" -Value $matchedSource."Site URL" -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Source" -Value $matchedSource."Last Activity Date" -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Source" -Value $matchedSource."File Count" -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Source" -Value $matchedSource."Storage Used (Byte)" -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Source" -Value $null -force
    }
    #gatherOneDriveDetails - Ametek
    if ($matchedDestination = $DestinationOneDriveDetails | ?{$_."Owner Principal Name" -eq $mailbox.UserPrincipalName_Destination}) {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Destination" -Value $matchedDestination."Site URL" -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Destination" -Value $matchedDestination."Last Activity Date" -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Destination" -Value $matchedDestination."File Count" -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Destination" -Value $matchedDestination."Storage Used (Byte)" -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageusage_Destination" -Value $null -force
    } 
}
$allmatchedMailboxes | export-csv -NoTypeInformation -encoding utf8 "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\OneDriveAbacoMatched-Mailboxes.csv"

#Get Retention Hold Enabled old User's Details
$matchedMailboxes  = Import-Excel -path 'c:\Users\amedrano\Arraya Solutions\Thomas Jefferson - Einstein to Jefferson Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$progressref = $matchedMailboxes.count
$progresscounter = 0
foreach ($user in $matchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details UserMailbox $($user.PrimarySmtpAddress_Destination)"

    if ($matchedMailbox = Get-Mailbox $user.PrimarySmtpAddress_Destination -ErrorAction SilentlyContinue) {
        $user | add-member -type noteproperty -name "RetentionHoldEnabled" -Value $matchedMailbox.RetentionHoldEnabled -force
    }
    else {
        $user | add-member -type noteproperty -name "RetentionHoldEnabled" -Value $null -force
    }
}


# Update PrimarySMTPAddress for Office365 Groups
$progressref = $Office365Groups.count
$progresscounter = 0
foreach ($group in $Office365Groups) {
    $oldAbacoAddress = $group.PrimarySMTPAddress_Source
    $ametekAddress = $group.PrimarySmtpAddress_Destination
    $newAmetekAddress = ($ametekAddress -split "@")[0] + "@ametek.com"
    $abacoPrefixAddress = ($oldAbacoAddress -split "@")[0] + "@ametek.com"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Office365 Group $($group.ametekAddress)"

    if ($group.RecipientTypeDetails_Source -eq "GroupMailbox") {
        if ($groupCheck = Get-UnifiedGroup $ametekAddress -ea silentlycontinue) {
            Write-Host "Updating $($groupcheck.primarysmtpaddress) to $($newAmetekAddress) and add alias $($abacoPrefixAddress)"
            Set-UnifiedGroup $groupcheck.primarysmtpaddress -PrimarySmtpAddress $newAmetekAddress
            Set-UnifiedGroup $groupcheck.primarysmtpaddress -EmailAddresses @{add=$abacoPrefixAddress}
        }
        else {
            Write-Host "no Group found"
            $notfoundgroup += $group
        }
    }    
}


$abacoMailUsers = Get-MailUser -ResultSize unlimited -Filter "ExternalEmailAddress -like '*@abaco.com'"
$progressref = $abacoMailUsers.count
$progresscounter = 0
foreach ($user in $abacoMailUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Enable Remote Mailbox for $($user.name)"
    Enable-RemoteMailbox $user.DistinguishedName
}

