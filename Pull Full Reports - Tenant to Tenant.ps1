# Pull Full Reports Tenant to Tenant

# Gather Mailbox Stats - Include SharePoint
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
    $UPN = $msoluser.UserPrincipalName
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    #Create User Output Array
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $msoluser.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $msoluser.userprincipalname
    $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msoluser.IsLicensed
    $currentuser | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ",") -force
    $currentuser | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ",") -force
    $currentuser | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential
    $currentuser | add-member -type noteproperty -name "Department" -Value $msoluser.Department    
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $user.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "Alias" -Value $user.alias
    $currentuser | add-member -type noteproperty -name "CustomAttribute7" -Value $user.CustomAttribute7
    $currentuser | add-member -type noteproperty -name "WhenCreated" -Value $user.WhenCreated
    $currentuser | add-member -type noteproperty -name "LastLogonTime" -Value $mbxStats.LastLogonTime
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ",")
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $user.legacyexchangedn)
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $user.HiddenFromAddressListsEnabled
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $user.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $user.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $user.ForwardingSmtpAddress
    $currentuser | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize
    $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount

    #Pull Send on Behalf
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($grantSendOnBehalfPerms -join ",")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.primarysmtpaddress | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value ($mbxPermissions.user -join ",") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value $null
    }
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.PrimarySMTPAddress | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ",") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value $null
    }
    # Archive Mailbox Check
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus
    if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {           
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $ArchiveStats.TotalItemSize.Value
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $ArchiveStats.ItemCount
    }
    else {
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $null
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $null
    }
    # Gather SharePoint Details
    $count = 0
    $success = $null

    do{
        try{
            $OneDriveSite = Get-SPOSite -Filter "Owner -eq '$UPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop
            if ($OneDriveSite) {
                $currentuser | add-member -type noteproperty -name "SourceOneDriveURL" -Value $OneDriveSite.url -force
                Write-Host ". " -foregroundcolor green -nonewline
                $success = $true
            }
           else {
                $currentuser | add-member -type noteproperty -name "SourceOneDriveURL" -Value $null -force
                Write-Host ". " -foregroundcolor red -nonewline
                $failed = $true
           }
        }
        catch{
            Write-host "Next attempt in 30 seconds" -foregroundcolor yellow -nonewline
            Start-sleep -Seconds 30
            $count++
        # Put the start-sleep in the catch statemtnt so we
        # don't sleep if the condition is true and waste time
        }
    }
    until($count -eq 5 -or $success -or $failed)

    if(!($success -or $failed)) {
        $currentuser | add-member -type noteproperty -name "SourceOneDriveURL" -Value $null -force
        Write-Host ". " -foregroundcolor red -nonewline
    }

    $sourceMailboxStats += $currentuser
}

#CONDENSED Match Mailboxes and add to same spreadsheet. Check based on CustomAttribute7 and DisplayName - Includes SharePoint
$matchedMailboxes = Import-Csv
$progressref = ($matchedMailboxes).count
$progresscounter = 0
$newDomain = Read-Host "What is the new Domain?"

foreach ($user in $matchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    $NewUPN = $user.CampusKey + "@" + $newDomain
    $addressSplit = $user.PrimarySmtpAddress -split "@"
    $ehnPrimarySMTPAddress = $addressSplit[0] + "-ehn@jefferson.edu"
    $newPrimarySMTPAddress = $addressSplit[0] + "@" + $newDomain

    #MailUserCheck Check
    if ($mailUser = Get-MailUser $user.PrimarySMTPAddress) {
        $msoluser = Get-Msoluser -UserPrincipalName $mailUser.UserPrincipalName -ErrorAction SilentlyContinue
        Write-Host "$($msoluser.UserPrincipalName) User found with MailUser" -ForegroundColor Green
        $MatchType = "MailUser"
    }
    #NEW PrimarySMTPAddress Check (EHN added)
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $ehnPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with EHN added" -ForegroundColor Yellow
        $MatchType = "CompanyNameAdded"
    }
    #NEW PrimarySMTPAddress Check
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $newPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
        $MatchType = "NewPrimarySMTPAddress"
    }
    #UPN Check based on DisplayName
    elseif ($msoluser = Get-Msoluser -SearchString $user.DisplayName -ErrorAction SilentlyContinue) {
        if ($msoluser.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName)"
            $mailbox = Get-Mailbox $user.DisplayName -ErrorAction SilentlyContinue
            $msoluser = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        }
        Write-Host "$($msoluser.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
        $MatchType = "DisplayName"
    }

    #Gather Stats and Output
    if ($msoluser) {
        $UPN = $msoluser.UserPrincipalName
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $true -force
        $user | add-member -type noteproperty -name "MatchType" -Value $MatchType -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ",") -force
        $user | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ",") -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $mailbox.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ",") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastUserAccessTime_Destination" -Value $mbxStats.LastUserAccessTime -force
        # Archive Check    
        if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
        }
        else  {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }
        # Gather SharePoint Details
        $count = 0
        $success = $null
        do{
            try{
                $OneDriveSite = Get-SPOSite -Filter "Owner -eq '$UPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop
                if ($OneDriveSite) {
                    $user | add-member -type noteproperty -name "OneDriveUrl_Destination" -Value $OneDriveSite.url -force
                    Write-Host ". " -foregroundcolor green -nonewline
                    $success = $true
                }
                else {
                        $user | add-member -type noteproperty -name "OneDriveUrl_Destination" -Value $null -force
                        Write-Host ". " -foregroundcolor red -nonewline
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
            $user | add-member -type noteproperty -name "OneDriveUrl_Destination" -Value $null -force
            Write-Host ". " -foregroundcolor red -nonewline
        }
    }
    
    else {
        Write-Host "  Unable to find user for $($user.DisplayName)" -ForegroundColor Red
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
        $user | add-member -type noteproperty -name "MatchType" -Value $null -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "Licenses" -Value $null -force
        $user | add-member -type noteproperty -name "License-DisabledArray" -Value $null -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $null -force
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
        $user | add-member -type noteproperty -name "LastUserAccessTime_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $null-force
        $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "OneDriveUrl_Destination" -Value $null -force
    }
}

#UpdateSourceMailbox with previously matched details
