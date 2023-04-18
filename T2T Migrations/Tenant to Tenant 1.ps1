#OldCompany & NewCompany

#Gather Mailbox Stats
$OutputCSVFolderPath = Read-Host "What is the folder path to store the file?"
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
    $mbxStats = Get-MailboxStatistics $user.DistinguishedName
    $msoluser = Get-MsolUser -UserPrincipalName $user.UserPrincipalName
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    #Create User Array
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $msoluser.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName_Source" -Value $msoluser.userprincipalname
    $currentuser | add-member -type noteproperty -name "IsLicensed_Source" -Value $msoluser.IsLicensed
    $currentuser | add-member -type noteproperty -name "Licenses_Source" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
    $currentuser | add-member -type noteproperty -name "License-DisabledArray_Source" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
    $currentuser | add-member -type noteproperty -name "BlockCredential_Source" -Value $msoluser.BlockCredential
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
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.DistinguishedName | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value ($mbxPermissions.user -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value $null
    }
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.DistinguishedName | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value $null
    }
    # Archive Mailbox Check
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus
    if ($ArchiveStats = Get-MailboxStatistics $user.DistinguishedName -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {           
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $ArchiveStats.TotalItemSize.Value
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $ArchiveStats.ItemCount
    }
    else {
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $null
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $null
    }
    $sourceMailboxStats += $currentuser
}
$sourceMailboxStats | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$OutputCSVFolderPath\SourceMailboxes.csv"

#Gather Campus Key from EHN AD Export
$sourceMailboxStats = Import-Csv
$ehnEnabledUsers = Import-Csv
$progressref = ($sourceMailboxStats).count
$progresscounter = 0

foreach ($user in $sourceMailboxStats) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Check Campus Key for $($user.DisplayName_Source)"

    #Check for Campus Key
    if ($userCheck = $ehnEnabledUsers | ?{$_.Logon -eq $user.UserPrincipalName_Source}) {    
        $user | add-member -type noteproperty -name "CampusKey" -Value $userCheck.CampusKey -force
    }
    else {
        $user | add-member -type noteproperty -name "CampusKey" -Value $null -force
    }
}

#Match Mailboxes and add to same spreadsheet. Check based on Campus Key, CustomAttribute7 and DisplayName
$matchedMailboxes = Import-Csv
$progressref = ($matchedMailboxes).count
$progresscounter = 0

foreach ($user in $matchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName_Source)"
    $NewUPN = $user.CustomAttribute7_Source + "@example.org"
    $CampusKeyUPN = $user.CampusKey + "@example.org"
    $addressSplit = $user.PrimarySmtpAddress_Source -split "@"
    $ehnPrimarySMTPAddress = $addressSplit[0] + "-old@example.org"
    $newPrimarySMTPAddress = $addressSplit[0] + "@example.org"
    $PrimarySMTPAddress = $user.primarysmtpaddress_Source

    # Campus Key Match
    if ($msoluser = Get-Msoluser -UserPrincipalName $CampusKeyUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with CampusKey" -ForegroundColor Green
        $mailbox = Get-EXOMailbox -PropertySets archive,addresslist,delivery $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-EXOMailboxStatistics -PropertySets All $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
    
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "CampusKeyMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $mailbox.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize.Value -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $mailbox.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime.ToShortDateString() -force
        $user | add-member -type noteproperty -name "LastUserActionTime_Destination" -Value $mbxStats.LastUserActionTime.ToShortDateString() -force
        
        if ($ArchiveStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
            }
            else  {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
            }
        }
    #New UPN Match - CustomAttribute7
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $NewUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with CampusKey" -ForegroundColor Green
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
    
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "CampusKeyMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $mailbox.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastUserAccessTime_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
            }
            else  {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
            }
    }
    #NEW PrimarySMTPAddress Check (EHN added)
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $ehnPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with EHN added" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "EHNMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed   -force
        $user | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $mailbox.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "LastUserAccessTime_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
            }
        else {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }
    }
    #NEW PrimarySMTPAddress Check
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $newPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "NEWSMTPAddressCheck" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed   -force
        $user | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $mailbox.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "LastUserAccessTime_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
            }
        else {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }
    }
    #UPN Check based on DisplayName
    elseif ($msoluser = Get-Msoluser -SearchString $user.DisplayName -ErrorAction SilentlyContinue) {
        if ($msoluser.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName_Source)"
            $mailbox = Get-Mailbox $user.DisplayName -ErrorAction SilentlyContinue
            $msoluser = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        }
        Write-Host "$($msoluser.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "DisplayNameMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $mailbox.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "LastUserAccessTime_Destination" -Value $mbxStats.LastUserAccessTime -force
        
        if ($ArchiveStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
            }
        else {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
        }
    }
    
    else {
        Write-Host "  Unable to find user for $($user.DisplayName_Source)" -ForegroundColor Red
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
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
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
    }
}

#Match Mailboxes and add to same spreadsheet. Check based on Campus Key, CustomAttribute7 and DisplayName - CONDENSED
$matchedMailboxes = Import-Excel -WorksheetName "Master List2" -Path 
$progressref = ($matchedMailboxes).count
$progresscounter = 0

foreach ($user in $matchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName_Source)"
    $NewUPN = $user.CustomAttribute7_Source + "@example.org"
    $CampusKeyUPN = $user.CampusKey + "@example.org"
    $addressSplit = $user.PrimarySmtpAddress_Source -split "@"
    $ehnPrimarySMTPAddress = $addressSplit[0] + "-old@example.org"
    $newPrimarySMTPAddress = $addressSplit[0] + "@example.org"

    # Campus Key Match
    if ($msoluser = Get-Msoluser -UserPrincipalName $CampusKeyUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with CampusKey" -ForegroundColor Green
    }
    #New UPN Match - CustomAttribute7
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $NewUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with CampusKey" -ForegroundColor Green
    }

    #NEW PrimarySMTPAddress Check (EHN added)
    elseif ($mailboxCheck = Get-Mailbox $ehnPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        $msoluser = Get-Msoluser -UserPrincipalName $mailboxCheck.UserPrincipalName -ErrorAction SilentlyContinue
        Write-Host "$($msoluser.UserPrincipalName) User found with EHN added" -ForegroundColor Yellow
    }
    #NEW PrimarySMTPAddress Check
    elseif ($mailboxCheck = Get-Mailbox $newPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        $msoluser = Get-Msoluser -UserPrincipalName $mailboxCheck.UserPrincipalName -ErrorAction SilentlyContinue
        Write-Host "$($msoluser.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
    }
    #UPN Check based on DisplayName
    elseif ($msoluser = Get-Msoluser -SearchString $user.DisplayName -ErrorAction SilentlyContinue) {
        if ($msoluser.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName_Source)"
            $mailbox = Get-Mailbox $user.DisplayName -ErrorAction SilentlyContinue
            $msoluser = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        }
        Write-Host "$($msoluser.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
    }

    if ($msoluser) {
        #Gather Stats
        $mailbox = Get-EXOMailbox -PropertySets archive,addresslist,delivery,Minimum $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-EXOMailboxStatistics -PropertySets All $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses

        #Output Stats
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "CampusKeyMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $mailbox.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize.Value -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $mailbox.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime.ToShortDateString() -force
        $user | add-member -type noteproperty -name "LastUserActionTime_Destination" -Value $mbxStats.LastUserActionTime.ToShortDateString() -force
        
        if ($ArchiveStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
            }
            else  {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
            }
    }
    else {
        Write-Host "  Unable to find user for $($user.DisplayName_Source)" -ForegroundColor Red

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
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
        $user | add-member -type noteproperty -name "LastUserActionTime_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
    }
}

#Check for Conflicts in own tenants
$ConflictingUsers = @()

#ProgressBar
$progressref = ($mailboxes).count
$progresscounter = 0

foreach ($mailbox in $mailboxes) {
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Mailbox Conflicting Name for $($mailbox.DisplayName)"
    
    $mailboxcheck = Get-Mailbox $mailbox.displayName -ea silentlycontinue
    if ($mailboxCheck.count -gt 1)
    {
        Write-Host "Multiple mailboxes found for $($mailbox.DisplayName) *" -ForegroundColor Yellow -NoNewline

        $msoluserscheck = get-msoluser -UserPrincipalName $mailbox.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation
        $MBXStats = Get-MailboxStatistics $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount

        $currentuser = new-object PSObject

        $currentuser | add-member -type noteproperty -name "DisplayName" -Value $mailbox.DisplayName
        $currentuser | add-member -type noteproperty -name "Name" -Value $mailbox.Name
        $currentuser | add-member -type noteproperty -name "Alias" -Value $mailbox.Alias
        $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $msoluserscheck.UserPrincipalName
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $mailbox.PrimarySmtpAddress
        $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msoluserscheck.IsLicensed
        $currentuser | add-member -type noteproperty -name "Licenses" -Value ($msoluserscheck.Licenses.AccountSkuID -join ";")
        $currentuser | add-member -type noteproperty -name "IsDirSynced" -Value $mailbox.IsDirSynced
        $currentuser | add-member -type noteproperty -name "PreferredDataLocation" -Value $msoluserscheck.PreferredDataLocation
        $currentuser | add-member -type noteproperty -name "Database" -Value $mailbox.Database
        $currentuser | add-member -type noteproperty -name "BlockSigninStatus" -Value $msoluserscheck.BlockCredential
        $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mailbox.RecipientTypeDetails   
        $currentuser | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize
        $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount
        Write-host " .. done" -foregroundcolor green

        $ConflictingUsers += $currentuser
    }
}
$ConflictingUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath

#Check For Conflicting Users Across Tenants Based on DisplayName
$OldCompanyMailboxes = import-csv

$matchedConflictingUsers = @()
#ProgressBar
$progressref = ($OldCompanyMailboxes).count
$progresscounter = 0
foreach ($mailbox in $OldCompanyMailboxes) {
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Mailbox Conflicting Name for $($mailbox.DisplayName)"
    
    if ($mailboxConflictcheck = Get-Mailbox $mailbox.DisplayName -ea silentlycontinue)
    {
        Write-Host "Mailbox found for $($mailbox.DisplayName) " -ForegroundColor Yellow -NoNewline
        foreach ($conflictMailbox in $mailboxConflictcheck) {
        $msoluserscheck = get-msoluser -UserPrincipalName $conflictMailbox.UserPrincipalName -ea silentlycontinue | select IsLicensed, licenses, BlockCredential, UserPrincipalName,Department,WhenCreated
        $MBXStats = Get-MailboxStatistics $conflictMailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount

        $currentuser = new-object PSObject
        $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $mailbox.DisplayName
        $currentuser | add-member -type noteproperty -name "Name_Source" -Value $mailbox.Name
        $currentuser | add-member -type noteproperty -name "UserPrincipalName_Name_Source" -Value $mailbox.UserPrincipalName
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $mailbox.PrimarySmtpAddress
        $currentuser | add-member -type noteproperty -name "EmailAddresses_Source" -Value $mailbox.EmailAddresses
        $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $mailbox.RecipientTypeDetails

        $currentuser | add-member -type noteproperty -name "DisplayName_destination" -Value $conflictMailbox.DisplayName
        $currentuser | add-member -type noteproperty -name "Department_destination" -Value $msoluserscheck.Department
        $currentuser | add-member -type noteproperty -name "WhenCreated_destination" -Value $msoluserscheck.WhenCreated
        $currentuser | add-member -type noteproperty -name "Name_destination" -Value $conflictMailbox.Name
        $currentuser | add-member -type noteproperty -name "Alias_destination" -Value $conflictMailbox.Alias
        $currentuser | add-member -type noteproperty -name "UserPrincipalName_destination" -Value $msoluserscheck.UserPrincipalName
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress_destination" -Value $conflictMailbox.PrimarySmtpAddress
        $currentuser | add-member -type noteproperty -name "IsLicensed_destination" -Value $msoluserscheck.IsLicensed
        $currentuser | add-member -type noteproperty -name "Licenses_destination" -Value ($msoluserscheck.Licenses.AccountSkuID -join ";")
        $currentuser | add-member -type noteproperty -name "IsDirSynced_destination" -Value $conflictMailbox.IsDirSynced
        $currentuser | add-member -type noteproperty -name "BlockSigninStatus_destination" -Value $msoluserscheck.BlockCredential
        $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $conflictMailbox.RecipientTypeDetails   
        $currentuser | Add-Member -type NoteProperty -Name "MBXSize_destination" -Value $MBXStats.TotalItemSize
        $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount_destination" -Value $MBXStats.ItemCount
        $matchedConflictingUsers += $currentuser
        
        }
        Write-host " .. done" -foregroundcolor green
    }
    
}
$matchedConflictingUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath


#Check For Conflicting Users Across Tenants Based on PrimarySMTPAddress prefix
$OldCompanyMailboxes = import-csv
$NewCompanyMailboxes = Import-Csv

$matchedConflictingUsers = @()
#ProgressBar
$progressref = ($OldCompanyMailboxes).count
$progresscounter = 0

foreach ($mailbox in $OldCompanyMailboxes){
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Mailbox Conflicting Name for $($mailbox.DisplayName)"
    $newPrimarySMTPAddressSplit = $mailbox.PrimarySmtpAddress -split "@"
    $prefixSMTPAddress = $newPrimarySMTPAddressSplit[0]

    if ($matchedUsers = $NewCompanyMailboxes | Where-Object {$_.PrimarySmtpAddress -like "*$prefixSMTPAddress*"})
    {
        Write-Host "Matched User found for $($mailbox.DisplayName) " -ForegroundColor Yellow -NoNewline

        foreach ($user in $matchedUsers) {
            #$conflictMailbox = Get-Mailbox $user.PrimarySmtpAddress -ea silentlycontinue
            
            #$msoluserscheck = get-msoluser -UserPrincipalName $conflictMailbox.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation
            #$MBXStats = Get-MailboxStatistics $conflictMailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
            
            $currentuser = new-object PSObject
            $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $mailbox.DisplayName
            $currentuser | add-member -type noteproperty -name "Name_Source" -Value $mailbox.Name
            $currentuser | add-member -type noteproperty -name "UserPrincipalName_Source" -Value $mailbox.UserPrincipalName
            $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $mailbox.PrimarySmtpAddress
            $currentuser | add-member -type noteproperty -name "EmailAddresses_Source" -Value $mailbox.EmailAddresses
            $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $mailbox.RecipientTypeDetails

            $currentuser | add-member -type noteproperty -name "DisplayName_destination" -Value $user.DisplayName
            $currentuser | add-member -type noteproperty -name "Name_destination" -Value $user.Name
            $currentuser | add-member -type noteproperty -name "UserPrincipalName_destination" -Value $user.UserPrincipalName
            $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress_destination" -Value $user.PrimarySmtpAddress
            $currentuser | add-member -type noteproperty -name "EmailAddresses_destination" -Value $user.EmailAddresses
            $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $user.RecipientTypeDetails

            $matchedConflictingUsers += $currentuser
        }
        Write-host " .. done" -foregroundcolor green
    }
}
$matchedConflictingUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath

#Check For Conflicting Users Across Tenants Based on PrimarySMTPAddress prefix 2 Exact Match
$OldCompanyMailboxes = import-csv
$NewCompanyMailboxes = Import-Csv

$matchedConflictingUsers = @()
#ProgressBar
$progressref = ($OldCompanyMailboxes).count
$progresscounter = 0

foreach ($mailbox in $OldCompanyMailboxes){
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Mailbox Conflicting Name for $($mailbox.DisplayName)"
    $OldCompanyPrimarySMTPAddressSplit = $mailbox.PrimarySmtpAddress -split "@"
    $OldCompanyprefixSMTPAddress = $OldCompanyPrimarySMTPAddressSplit[0]

    if ($matchedUsers = $NewCompanyMailboxes | Where-Object {$_.PrimarySmtpAddress -like "*$OldCompanyprefixSMTPAddress*"})
    {
        Write-Host "Matched User found for $($mailbox.DisplayName) " -ForegroundColor Yellow -NoNewline

        foreach ($user in $matchedUsers) {
            $NewCompanyPrimarySMTPAddressSplit = $user.PrimarySmtpAddress -split "@"
            $NewCompanyprefixSMTPAddress = $NewCompanyPrimarySMTPAddressSplit[0]
            if ($OldCompanyprefixSMTPAddress -eq $NewCompanyprefixSMTPAddress) {
                $currentuser = new-object PSObject
                $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $mailbox.DisplayName
                $currentuser | add-member -type noteproperty -name "Name_Source" -Value $mailbox.Name
                $currentuser | add-member -type noteproperty -name "UserPrincipalName_Source" -Value $mailbox.UserPrincipalName
                $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $mailbox.PrimarySmtpAddress
                $currentuser | add-member -type noteproperty -name "EmailAddresses_Source" -Value $mailbox.EmailAddresses
                $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $mailbox.RecipientTypeDetails
    
                $currentuser | add-member -type noteproperty -name "DisplayName_destination" -Value $user.DisplayName
                $currentuser | add-member -type noteproperty -name "Name_destination" -Value $user.Name
                $currentuser | add-member -type noteproperty -name "UserPrincipalName_destination" -Value $user.UserPrincipalName
                $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress_destination" -Value $user.PrimarySmtpAddress
                $currentuser | add-member -type noteproperty -name "EmailAddresses_destination" -Value $user.EmailAddresses
                $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $user.RecipientTypeDetails
                $matchedConflictingUsers += $currentuser 
            }
        }
        Write-host " .. done" -foregroundcolor green
    }
}
$matchedConflictingUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath

#Get DistributionGroupa
function Get-GroupDetails {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath
    )
    $allMailGroups = Get-Recipient -RecipientTypeDetails group -ResultSize unlimited
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailNonUniversalGroup -ResultSize unlimited
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize unlimited
    $allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize unlimited    
    $allMailGroups += Get-Recipient -RecipientTypeDetails DynamicDistributionGroup -ResultSize unlimited
    $allGroupDetails = @()

    #ProgressBarA
    $progressref = ($allMailGroups).count
    $progresscounter = 0

    foreach ($object in $allMailGroups)
    {
        #ProgressBarB
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Group Details for $($object.DisplayName)"
        #Get Groups Details
        $members = @()
        $EmailAddresses = $object | select -expandProperty EmailAddresses
        
        #Create Output Array
        $currentobject = new-object PSObject
        $currentobject | add-member -type noteproperty -name "DisplayName" -Value $object.DisplayName -Force
        $currentobject | add-member -type noteproperty -name "Name" -Value $object.Name -Force
        $currentobject | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $object.PrimarySMTPAddress -Force
        $currentobject | add-member -type noteproperty -name "IsDirSynced" -Value $object.IsDirSynced -Force
        $currentobject | add-member -type noteproperty -name "RecipientTypeDetails" -Value $object.RecipientTypeDetails
        $currentobject | add-member -type noteproperty -name "Alias" -Value $object.alias -Force
        $currentobject | add-member -type noteproperty -name "Notes" -Value $object.Notes -Force
        $currentobject | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ";")
        $currentobject | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("X500:" + $object.ServerLegacyDN)
        $currentobject | add-member -type noteproperty -name "ManagedBy" -Value ($object.ManagedBy -join ";")
        $currentobject | add-member -type NoteProperty -name "HiddenFromAddressListsEnabled" -Value $object.HiddenFromAddressListsEnabled -force

        #Pull DynamicDistributionGroup Details
        if ($object.RecipientTypeDetails -eq "DynamicDistributionGroup") {
            $dynamicGroup = Get-DynamicDistributionGroup $object.PrimarySMTPAddress
            $members = (Get-DynamicDistributionGroupMember $object.PrimarySMTPAddress -ResultSize unlimited).count
            
            $currentobject | add-member -type noteproperty -name "MembersCount" -Value ($members.tostring()) -Force
            $currentobject | add-member -type noteproperty -name "Members" -Value "A lot or a little" -Force
            $currentobject | add-member -type NoteProperty -name "ModeratedBy" -Value ($dynamicGroup.ModeratedBy -join ";") -force 
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFrom" -Value $dynamicGroup.AcceptMessagesOnlyFrom -force
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromDLMembers" -Value $dynamicGroup.AcceptMessagesOnlyFromDLMembers -force
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromSendersOrMembers" -Value $dynamicGroup.AcceptMessagesOnlyFromSendersOrMembers -force
            $currentobject | add-member -type NoteProperty -name "GrantSendOnBehalfTo" -Value $dynamicGroup.GrantSendOnBehalfTo -force
            $currentobject | add-member -type NoteProperty -name "RequireSenderAuthenticationEnabled" -Value $dynamicGroup.RequireSenderAuthenticationEnabled -force
            $currentobject | add-member -type NoteProperty -name "HiddenGroupMembershipEnabled" -Value $null -force         
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFrom" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromDLMembers" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromSendersOrMembers" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "AccessType" -Value $null -force    
            $currentobject | add-member -type NoteProperty -name "AllowAddGuests" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "IsMailboxConfigured" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "ResourceProvisioningOptions" -Value $null -force
            
        }
        elseif ($object.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
            $distributionGroup = Get-DistributionGroup $object.PrimarySMTPAddress
            $members = Get-DistributionGroupMember $object.PrimarySMTPAddress -ResultSize unlimited
            $currentobject | add-member -type noteproperty -name "MembersCount" -Value ($members.count) -Force
            $currentobject | add-member -type noteproperty -name "Members" -Value ($members[0..999].PrimarySMTPAddress -join ";") -Force
            $currentobject | add-member -type NoteProperty -name "DynamicRecipientFilter" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "ModeratedBy" -Value ($distributionGroup.ModeratedBy -join ";") -force 
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFrom" -Value $distributionGroup.AcceptMessagesOnlyFrom -force
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromDLMembers" -Value $distributionGroup.AcceptMessagesOnlyFromDLMembers -force
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromSendersOrMembers" -Value $distributionGroup.AcceptMessagesOnlyFromSendersOrMembers -force
            $currentobject | add-member -type NoteProperty -name "GrantSendOnBehalfTo" -Value $distributionGroup.GrantSendOnBehalfTo -force
            $currentobject | add-member -type NoteProperty -name "RequireSenderAuthenticationEnabled" -Value $distributionGroup.RequireSenderAuthenticationEnabled -force
            $currentobject | add-member -type NoteProperty -name "HiddenGroupMembershipEnabled" -Value $distributionGroup.HiddenGroupMembershipEnabled -force         
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFrom" -Value $groupmailboxdetails.RejectMessagesOnlyFrom -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromDLMembers" -Value $groupmailboxdetails.RejectMessagesOnlyFromDLMembers -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromSendersOrMembers" -Value $groupmailboxdetails.RejectMessagesOnlyFromSendersOrMembers -force
            $currentobject | add-member -type NoteProperty -name "AccessType" -Value $null -force    
            $currentobject | add-member -type NoteProperty -name "AllowAddGuests" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "IsMailboxConfigured" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "ResourceProvisioningOptions" -Value $null -force
        }
        elseif ($object.RecipientTypeDetails -eq "GroupMailbox") {
            $groupmailboxdetails = Get-UnifiedGroup $object.PrimarySMTPAddress
            $members = Get-UnifiedGroupLinks -Identity $object.PrimarySMTPAddress -LinkType Member -ResultSize unlimited
            $currentobject | add-member -type noteproperty -name "MembersCount" -Value ($groupmailboxdetails.GroupMemberCount) -Force
            $currentobject | add-member -type noteproperty -name "Members" -Value ($members.PrimarySMTPAddress -join ";") -Force
            $currentobject | add-member -type NoteProperty -name "DynamicRecipientFilter" -Value $null -force
            $currentobject | add-member -type NoteProperty -name "ModeratedBy" -Value ($groupmailboxdetails.ModeratedBy -join ";") -force 
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFrom" -Value $groupmailboxdetails.AcceptMessagesOnlyFrom -force
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromDLMembers" -Value $groupmailboxdetails.AcceptMessagesOnlyFromDLMembers -force
            $currentobject | add-member -type NoteProperty -name "AcceptMessagesOnlyFromSendersOrMembers" -Value $groupmailboxdetails.AcceptMessagesOnlyFromSendersOrMembers -force
            $currentobject | add-member -type NoteProperty -name "GrantSendOnBehalfTo" -Value $groupmailboxdetails.GrantSendOnBehalfTo -force
            $currentobject | add-member -type NoteProperty -name "RequireSenderAuthenticationEnabled" -Value $groupmailboxdetails.RequireSenderAuthenticationEnabled -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFrom" -Value $groupmailboxdetails.RejectMessagesOnlyFrom -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromDLMembers" -Value $groupmailboxdetails.RejectMessagesOnlyFromDLMembers -force
            $currentobject | add-member -type NoteProperty -name "RejectMessagesOnlyFromSendersOrMembers" -Value $groupmailboxdetails.RejectMessagesOnlyFromSendersOrMembers -force
            $currentobject | add-member -type NoteProperty -name "AccessType" -Value $groupmailboxdetails.AccessType -force
            $currentobject | add-member -type NoteProperty -name "AllowAddGuests" -Value $groupmailboxdetails.AllowAddGuests -force
            $currentobject | add-member -type NoteProperty -name "HiddenGroupMembershipEnabled" -Value $groupmailboxdetails.HiddenGroupMembershipEnabled -force
            $currentobject | add-member -type NoteProperty -name "IsMailboxConfigured" -Value $groupmailboxdetails.IsMailboxConfigured -force
            $currentobject | add-member -type NoteProperty -name "ResourceProvisioningOptions" -Value ($groupmailboxdetails.ResourceProvisioningOptions -join ";") -force
        }
        $allGroupDetails += $currentobject
    }
    #Export
    $allGroupDetails | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}

$OldCompanyGroups = import-csv
#ProgressBar
$progressref = ($OldCompanyGroups).count
$progresscounter = 0

foreach ($group in $OldCompanyGroups){
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Members for Group $($group.DisplayName)"
    Write-Host "Checking group $($group.DisplayName) " -NoNewline
    $members = @()
    if ($matchedGroup.RecipientTypeDetails -eq "DynamicDistributionGroup") {
        Write-Host "dynamic group" -ForegroundColor yellow -NoNewline
        $dynamicGroupFilter = Get-DynamicDistributionGroup $group.PrimarySMTPAddress | select RecipientFilter
        $members = Get-DynamicDistributionGroupMember $group.PrimarySMTPAddress -ResultSize unlimited
        $currentgroup | add-member -type noteproperty -name "MembersCount" -Value ($members.count) -Force
        $currentgroup | add-member -type noteproperty -name "Members" -Value "A lot or a little" -Force
        $currentgroup | add-member -type NoteProperty -name "DynamicRecipientFilter" -Value $dynamicGroupFilter -force
        
    }
    else {
        $members = Get-DistributionGroupMember $group.PrimarySMTPAddress -ResultSize unlimited
        $currentgroup | add-member -type noteproperty -name "MembersCount" -Value ($members.count) -Force
        $currentgroup | add-member -type noteproperty -name "Members" -Value ($members.PrimarySMTPAddress -join ";") -Force
        $currentgroup | add-member -type NoteProperty -name "DynamicRecipientFilter" -Value $null -force
    }
    write-host ".. done"
}

#Check For Conflicting Groups Across Tenants Based on DisplayName
$OldCompanyGroups = import-csv
$matchedConflictingGroups = @()
#ProgressBar
$progressref = ($OldCompanyGroups).count
$progresscounter = 0

foreach ($group in $OldCompanyGroups){
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "DisplayName Conflict Checking for Group $($group.DisplayName)"

    if ($matchedGroups = Get-Recipient $group.DisplayName -ea silentlycontinue)
    {
        Write-Host "Matched Group found for $($group.DisplayName) " -ForegroundColor Yellow -NoNewline

        foreach ($matchedGroup in $matchedGroups) {

            $currentgroup = new-object PSObject
            $currentgroup | add-member -type noteproperty -name "DisplayName_Source" -Value $group.DisplayName
            $currentgroup | add-member -type noteproperty -name "Name_Source" -Value $group.Name
            $currentgroup | add-member -type noteproperty -name "UserPrincipalName_Source" -Value $group.UserPrincipalName
            $currentgroup | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $group.PrimarySmtpAddress
            $currentgroup | add-member -type noteproperty -name "EmailAddresses_Source" -Value $group.EmailAddresses
            $currentgroup | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $group.RecipientTypeDetails
            $currentgroup | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Source" -Value $group.HiddenFromAddressListsEnabled
            $currentgroup | add-member -type noteproperty -name "WhenCreated_Source" -Value $group.WhenCreated
            $currentgroup | add-member -type noteproperty -name "MembersCount_Source" -Value $group.MembersCount
            $currentgroup | add-member -type noteproperty -name "Members_Source" -Value $group.Members
            $currentgroup | add-member -type noteproperty -name "DynamicRecipientFilter_Source" -Value $group.DynamicRecipientFilter

            $currentgroup | add-member -type noteproperty -name "DisplayName_destination" -Value $matchedGroup.DisplayName
            $currentgroup | add-member -type noteproperty -name "Name_destination" -Value $matchedGroup.Name
            $currentgroup | add-member -type noteproperty -name "UserPrincipalName_destination" -Value $matchedGroup.UserPrincipalName
            $currentgroup | add-member -type noteproperty -name "PrimarySMTPAddress_destination" -Value $matchedGroup.PrimarySmtpAddress
            $currentgroup | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($matchedGroup.EmailAddresses -join ";")
            $currentgroup | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $matchedGroup.RecipientTypeDetails
            $currentgroup | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_destination" -Value $matchedGroup.HiddenFromAddressListsEnabled
            $currentgroup | add-member -type noteproperty -name "WhenCreated_destination" -Value $matchedGroup.WhenCreated

            if ($matchedGroup.RecipientTypeDetails -eq "DynamicDistributionGroup") {
                $dynamicGroupFilter = Get-DynamicDistributionGroup $matchedGroup.PrimarySMTPAddress | select RecipientFilter
                $members = Get-DynamicDistributionGroupMember $matchedGroup.PrimarySMTPAddress -ResultSize unlimited
                $currentgroup | add-member -type noteproperty -name "MembersCount_destination" -Value ($members.count) -Force
                $currentgroup | add-member -type noteproperty -name "Members_destination" -Value ($members.PrimarySMTPAddress -join ";") -Force
                $currentgroup | add-member -type NoteProperty -name "DynamicRecipientFilter_destination" -Value $dynamicGroupFilter -force
                
            }
            else {
                $members = Get-DistributionGroupMember $matchedGroup.PrimarySMTPAddress -ResultSize unlimited
                $currentgroup | add-member -type noteproperty -name "MembersCount_destination" -Value ($members.count) -Force
                $currentgroup | add-member -type noteproperty -name "Members_destination" -Value ($members.PrimarySMTPAddress -join ";") -Force
                $currentgroup | add-member -type NoteProperty -name "DynamicRecipientFilter_destination" -Value $null -force
            }       
            $matchedConflictingGroups += $currentgroup
        }
        Write-host " .. done" -foregroundcolor green
    }
}
$matchedConflictingGroups | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath

## match Groups to NewCompany from OldCompany

$OldCompanyGroups = Import-Csv 
$progressref = ($OldCompanyGroups).count
$progresscounter = 0
foreach ($group in $OldCompanyGroups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($group.DisplayName)"

    $addressSplits = $group.PrimarySMTPAddress
    $groupCheck = @()
    $ehnMatch = $addressSplits[0] + "-old@example.org"
    $newPrimarySMTPAddress = $addressSplits[0] + "@example.org"
    $ehndisplayName = $group.DisplayName + "- EHN"
    if ($groupCheck = Get-Recipient $ehnMatch -ea silentlycontinue) {
        $EmailAddresses = $groupCheck | select -ExpandProperty EmailAddresses

        $group | add-member -type noteproperty -name "ExistsInDestination" -Value "EHNMatched" -Force
        $group | add-member -type noteproperty -name "DisplayName_destination" -Value $groupCheck.DisplayName -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $groupCheck.RecipientTypeDetails -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $groupCheck.PrimarySmtpAddress -Force
        $group | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $group | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $groupCheck.legacyexchangedn) -Force
    }
    elseif ($groupCheck = Get-Recipient $ehndisplayName -ea silentlycontinue) {
        $EmailAddresses = $groupCheck | select -ExpandProperty EmailAddresses

        $group | add-member -type noteproperty -name "ExistsInDestination" -Value $true -Force
        $group | add-member -type noteproperty -name "DisplayName_destination" -Value $groupCheck.DisplayName -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $groupCheck.RecipientTypeDetails -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $groupCheck.PrimarySmtpAddress -Force
        $group | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $group | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $groupCheck.legacyexchangedn) -Force
    }
    elseif ($groupCheck = Get-Recipient $newPrimarySMTPAddress -ea silentlycontinue) {
        $EmailAddresses = $groupCheck | select -ExpandProperty EmailAddresses

        $group | add-member -type noteproperty -name "ExistsInDestination" -Value "AddressMatch" -Force
        $group | add-member -type noteproperty -name "DisplayName_destination" -Value $groupCheck.DisplayName -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $groupCheck.RecipientTypeDetails -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $groupCheck.PrimarySmtpAddress -Force
        $group | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $group | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $groupCheck.legacyexchangedn) -Force
    }
    elseif ($groupCheck = Get-Recipient $group.DisplayName -ea silentlycontinue) {
        $EmailAddresses = $groupCheck | select -ExpandProperty EmailAddresses

        $group | add-member -type noteproperty -name "ExistsInDestination" -Value $true -Force
        $group | add-member -type noteproperty -name "DisplayName_destination" -Value $groupCheck.DisplayName -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $groupCheck.RecipientTypeDetails -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $groupCheck.PrimarySmtpAddress -Force
        $group | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $group | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $groupCheck.legacyexchangedn) -Force
    }
    else {
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value $false -Force
        $group | add-member -type noteproperty -name "DisplayName_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "EmailAddresses_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value $null -Force
    }
}

#Update GrantSendOnBehalf Perms to Mailboxes
$progressref = ($OldCompanyTJMBXStats | ? {$_.GrantSendOnBehalfTo}).count
$progresscounter = 0
foreach ($user in $OldCompanyTJMBXStats | ? {$_.GrantSendOnBehalfTo}) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    $grantSendOnBehalf = $grantSendOnBehalf = (get-mailbox $user.primarysmtpaddress).GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $user | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($grantSendOnBehalfPerms -join ";") -Force
}

# Grant Full Access Perms to Mailboxes
$allmatchedMailboxes = Import-Csv
$sharedMailboxes = Import-CSV 
$fullAccessPermUsers = $matchedMailboxes | ? {$_.FullAccessPerms -and $_.ExistsInDestination -eq "CampusKeyMatch"}
$sendAsPermUsers =  $matchedMailboxes | ? {$_.SendAsPerms -and $_.ExistsInDestination -eq "CampusKeyMatch"}
$AllErrors = @()

$progressref = ($fullAccessPermUsers).count
$progresscounter = 0
foreach ($user in $fullAccessPermUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Full Access Perms to $($user.DisplayName_Destination)"
    Write-Host "Grant Full Access Perms for $($user.DisplayName_Destination).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Perm Users
    $membersArray = $user.FullAccessPerms -split ","

    #Progress Bar 2
    $progressref2 = ($membersArray).count
    $progresscounter2 = 0
    foreach ($member in $membersArray) {
        #Member Check
        $memberCheck = @()
        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

        #Progress Bar 2a
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Full Access to $($memberCheck.PrimarySmtpAddress_Destination)"

        #Add Perms to Mailbox      
        try {
			$permResult = Add-MailboxPermission -identity $user.PrimarySmtpAddress_Destination -AccessRights FullAccess -User $memberCheck.PrimarySmtpAddress_Destination -ea Stop
            Write-Host ". " -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host ". " -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantFullAccess" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $group.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror           
			continue
		}
    }
    Write-Host " done " -ForegroundColor Green
}

# Grant SendAs Perms to Mailboxes
$matchedMailboxes = Import-Csv
$fullAccessPermUsers = $matchedMailboxes | ? {$_.FullAccessPerms -and $_.ExistsInDestination -ne $false}
$sendAsPermUsers =  $matchedMailboxes | ? {$_.SendAsPerms -and $_.ExistsInDestination -ne $false}
$AllErrors = @()

$progressref = ($sendAsPermUsers).count
$progresscounter = 0
foreach ($user in $sendAsPermUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding SendAs Perms to $($user.DisplayName_Destination)"
    Write-Host "Grant SendAs Perms for $($user.DisplayName_Destination).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Perm Users
    $membersArray = $user.SendAsPerms -split ","

    #Progress Bar 2
    $progressref2 = ($membersArray).count
    $progresscounter2 = 0
    foreach ($member in $membersArray) {
        #Member Check
        $memberCheck = @()
        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

        #Progress Bar 2a
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting SendAs to $($memberCheck.PrimarySmtpAddress_Destination)"

        #Add Perms to Mailbox      
        try {
			$permResult = Add-RecipientPermission $user.PrimarySmtpAddress_Destination -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -confirm:$false -ea Stop
            Write-Host ". " -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host ". " -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantSendAs" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror           
			continue
		}
    }
    Write-Host " done " -ForegroundColor Green
}

# Grant SendAs Perms to DistributionGroups
$matchedDistributionGroups = Import-Csv
$sendAsPermUsers =  $matchedDistributionGroups | ? {$_.SendAsPerms -and $_.ExistsInDestination -ne $false}
$AllErrors = @()

$progressref = ($sendAsPermUsers).count
$progresscounter = 0
foreach ($user in $sendAsPermUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding SendAs Perms to $($user.DisplayName_Destination)"
    Write-Host "Grant SendAs Perms for $($user.DisplayName_Destination).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Perm Users
    $membersArray = $user.SendAsPerms -split ","

    #Progress Bar 2
    $progressref2 = ($membersArray).count
    $progresscounter2 = 0
    foreach ($member in $membersArray) {
        #Member Check
        $memberCheck = @()
        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

        #Progress Bar 2a
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting SendAs to $($memberCheck.PrimarySmtpAddress_Destination)"

        #Add Perms to Mailbox      
        try {
			$permResult = Add-RecipientPermission $user.PrimarySmtpAddress_Destination -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -confirm:$false -ea Stop
            Write-Host ". " -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host ". " -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantSendAs" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror           
			continue
		}
    }
    Write-Host " done " -ForegroundColor Green
}

# Grant Send On Behalf Perms to DistributionGroups
$matchedDistributionGroups = Import-Csv
$PermUsers =  $matchedDistributionGroups | ? {$_.GrantSendOnBehalfTo -and $_.ExistsInDestination -ne $false}
$AllGroupErrors = @()

$progressref = ($PermUsers).count
$progresscounter = 0
foreach ($user in $PermUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Granting Send On Behalf Perms to $($user.DisplayName_Destination)"
    Write-Host "Grant Send On Behalf Perms for $($user.DisplayName_Destination).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Perm Users
    $membersArray = $user.GrantSendOnBehalfTo -split ","

    #Progress Bar 2
    $progressref2 = ($membersArray).count
    $progresscounter2 = 0
    foreach ($member in $membersArray) {
        #Member Check
        $memberCheck = @()
        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

        #Progress Bar 2a
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Perms to $($memberCheck.PrimarySmtpAddress_Destination)"

        #Add Perms      
        try {
			$permResult = Set-DistributionGroup -Identity $user.PrimarySmtpAddress_Destination -GrantSendOnBehalfTo @{add=$memberCheck.PrimarySmtpAddress_Destination} -confirm:$false -ea Stop
            Write-Host ". " -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host ". " -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantSendAs" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror           
			continue
		}
    }
    Write-Host " done " -ForegroundColor Green
}

# Set Attributes including ManagedBy, Approved Senders, BypassModerationFromSendersOrMembers to DistributionGroups
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
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

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

    #Stamp BypassModerationFromSendersOrMembers Messages on Group
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
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

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

    #Stamp ModeratedBy Messages on Group
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
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

            #Progress Bar 2a
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Setting ModeratedBy to $($memberCheck.PrimarySmtpAddress_Destination)"

            #Set Attribute   
            try {
                Set-DistributionGroup -identity $user.PrimarySmtpAddress_Destination -ModeratedBy @{add=$memberCheck.PrimarySmtpAddress} -ea Stop -warningaction silentlycontinue
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

    Write-Host " done " -ForegroundColor Green
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
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

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
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

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
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

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
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

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
            $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

            #Progress Bar 2a
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting SendAs to $($memberCheck.PrimarySmtpAddress_Destination)"

            #Add Perms to Mailbox      
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

#Match Users to Teams and SharePoint Usage
$OldCompanyTJMBXStats = Import-Csv
$teamsUserUsage = Import-Csv
$sharePointUserUsage = Import-Csv
$progressref = ($OldCompanyTJMBXStats).count
$progresscounter = 0

foreach ($user in $OldCompanyTJMBXStats) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Teams/SharePoint Usage for $($user.DisplayName)"

    $90DaysAgo = (Get-Date).AddMonths(-3)
    #Teams Activity Check
    if ($teamsUsageCheck = $teamsUserUsage | ?{$_."User Principal Name" -eq $user.UserPrincipalName}) {
        if ($oldDate = Get-Date ($teamsUsageCheck."Last Activity Date") -ErrorAction SilentlyContinue) {
            $user | add-member -type noteproperty -name "TeamsLastActivityDate" -Value $oldDate -Force
            $user | add-member -type noteproperty -name "TeamsActivitySum" -Value $teamsUsageCheck.ActivitySum -Force
        }
        else {
            $user | add-member -type noteproperty -name "TeamsLastActivityDate" -Value $null -Force
            $user | add-member -type noteproperty -name "TeamsActivitySum" -Value $teamsUsageCheck.ActivitySum -Force
        }
    }
    else {
        $user | add-member -type noteproperty -name "TeamsLastActivityDate" -Value $null -Force
        $user | add-member -type noteproperty -name "TeamsActivitySum" -Value $null -Force
    }

    #SharePoint Activity Check 
    if ($sharePointUserUsageCheck = $sharePointUserUsage | ?{$_."User Principal Name" -eq $user.UserPrincipalName}){
        if ($oldDate = Get-Date ($sharePointUserUsageCheck."Last Activity Date") -ErrorAction SilentlyContinue)
        {
            if ($oldDate -gt $90DaysAgo) {
                $user | add-member -type noteproperty -name "SharePointUsage" -Value $true -Force
                $user | add-member -type noteproperty -name "SharePointLastActivityDate" -Value $oldDate -Force
            }
            else {
                $user | add-member -type noteproperty -name "SharePointUsage" -Value $false -Force
                $user | add-member -type noteproperty -name "SharePointLastActivityDate" -Value $null -Force
            }
        }
        else {
            $user | add-member -type noteproperty -name "SharePointUsage" -Value $true -Force
            $user | add-member -type noteproperty -name "SharePointLastActivityDate" -Value $null -Force
        }
    }
    else {
        $user | add-member -type noteproperty -name "SharePointUsage" -Value "NotFound" -Force
        $user | add-member -type noteproperty -name "SharePointLastActivityDate" -Value $null -Force
    }
}

## Update MailboxData in OldCompany with FullAccess and SendAs
$OldCompanyTJMBXStats = Import-Csv 
$progressref = $OldCompanyTJMBXStats.count
$progresscounter = 0
foreach ($user in $OldCompanyTJMBXStats) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Perms for $($user.DisplayName)"

    #Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.primarysmtpaddress | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $user | add-member -type noteproperty -name "FullAccessPerms" -Value ($mbxPermissions.user -join ";") -Force
    }
    else {
        $user | add-member -type noteproperty -name "FullAccessPerms" -Value $null
    }
    # Mailbox Send As Check
	if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.PrimarySMTPAddress | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $user | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ";") -Force
    }
    else {
        $user | add-member -type noteproperty -name "SendAsPerms" -Value $null
    }
}

## Update MailboxData in OldCompany with Forwarding
$OldCompanyTJMBXStats = Import-Csv 
$progressref = $OldCompanyTJMBXStats.count
$progresscounter = 0
foreach ($user in $OldCompanyTJMBXStats) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Forwarding for $($user.DisplayName_Source)"

    #Mailbox Full Access Check
    if ($mbxCheck = Get-Mailbox $user.primarysmtpaddress_Source | Select ForwardingAddress, ForwardingSmtpAddress) {
        $user | add-member -type noteproperty -name "ForwardingAddress_Source" -Value $mbxCheck.ForwardingAddress -Force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Source" -Value $mbxCheck.ForwardingSmtpAddress -Force
    }
    else {
        $user | add-member -type noteproperty -name "ForwardingAddress_Source" -Value $null
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Source" -Value $null
    }
}

## match Non-UserMailboxes to NewCompany from OldCompany
$OldCompanyNonUserMailboxes = Import-Csv 
$progressref = ($OldCompanyNonUserMailboxes).count
$progresscounter = 0
foreach ($user in $OldCompanyNonUserMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    $addressSplits = $user.PrimarySMTPAddress
    $newPrimarySMTPAddress = $addressSplits[0] + "-old@example.org"
    $newDisplayName = $user.DisplayName + " - EHN"
    $newAlias = $user.alias +"-old"

    $mailboxCheck = @()
    if ($mailboxCheck = Get-Recipient $newPrimarySMTPAddress -ea silentlycontinue) {
        $EmailAddresses = $mailboxCheck | select -ExpandProperty EmailAddresses

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $true -Force
        $user | add-member -type noteproperty -name "DisplayName_destination" -Value $mailboxCheck.DisplayName -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $mailboxCheck.RecipientTypeDetails -Force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $mailboxCheck.PrimarySmtpAddress -Force
        $user | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $user | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $mailboxCheck.legacyexchangedn) -Force
    }
    elseif ($mailboxCheck = Get-Recipient $newDisplayName -ea silentlycontinue) {
        $EmailAddresses = $mailboxCheck | select -ExpandProperty EmailAddresses

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $true -Force
        $user | add-member -type noteproperty -name "DisplayName_destination" -Value $mailboxCheck.DisplayName -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $mailboxCheck.RecipientTypeDetails -Force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $mailboxCheck.PrimarySmtpAddress -Force
        $user | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $user | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $mailboxCheck.legacyexchangedn) -Force
    }
    elseif ($mailboxCheck = Get-Recipient $addressSplits[0] -ea silentlycontinue) {
        $EmailAddresses = $mailboxCheck | select -ExpandProperty EmailAddresses

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $true -Force
        $user | add-member -type noteproperty -name "DisplayName_destination" -Value $mailboxCheck.DisplayName -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $mailboxCheck.RecipientTypeDetails -Force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $mailboxCheck.PrimarySmtpAddress -Force
        $user | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $user | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $mailboxCheck.legacyexchangedn) -Force
    }
    elseif ($mailboxCheck = Get-Recipient $user.DisplayName -ea silentlycontinue) {
        $EmailAddresses = $mailboxCheck | select -ExpandProperty EmailAddresses

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $true -Force
        $user | add-member -type noteproperty -name "DisplayName_destination" -Value $mailboxCheck.DisplayName -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $mailboxCheck.RecipientTypeDetails -Force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $mailboxCheck.PrimarySmtpAddress -Force
        $user | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $user | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $mailboxCheck.legacyexchangedn) -Force
    }
    else {
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $false -Force
        $user | add-member -type noteproperty -name "DisplayName_destination" -Value $null -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $null -Force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $null -Force
        $user | add-member -type noteproperty -name "EmailAddresses_destination" -Value $null -Force
        $user | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value $null -Force
    }
}

#Gather SharePoint Details
$progressref = ($OldCompanyMailboxes).count
$progresscounter = 0
foreach ($user in $OldCompanyMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    try 
        {
            $UPN = $user.UserPrincipalName_destination
            $SPOSITE = Get-SPOSITE -IncludePersonalSite $true -filter "Owner -eq $UPN -and URL -like '*-my.sharepoint*'" -ErrorAction SilentlyContinue
            Write-Host "$($user.DisplayName) ..." -ForegroundColor Cyan -NoNewline
            
            $user | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $SPOSITE.url -Force
            $user | Add-Member -type NoteProperty -Name "Owner" -Value $SPOSITE.Owner -Force
            $user | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $SPOSITE.StorageUsageCurrent -Force
            $user | Add-Member -type NoteProperty -Name "Status" -Value $SPOSITE.Status -Force
            $user | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $SPOSITE.SiteDefinedSharingCapability -Force
            $user | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $FDUser.LimitedAccessFileType -Force
            
            Write-Host "done" -ForegroundColor Green
        }
        catch 
        {
            Write-Host "OneDrive Not Enabled for User" -ForegroundColor Yellow
            $user | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $null -Force
            $user | Add-Member -type NoteProperty -Name "Owner" -Value $null -Force
            $user | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $null -Force
            $user | Add-Member -type NoteProperty -Name "Status" -Value $null -Force
            $user | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $null -Force
            $user | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $null -Force
        }
}

#Create New Dynamic Distribution Groups
$AllErrors = @()
$currentError = @()
foreach ($dl in $O365Groups) {
    $newGroup = @()
	Write-Host "$($dl.Destination_PrimarySmtpAddress) ... " -ForegroundColor Cyan -NoNewline
	
	if (!($newGroup = Get-UnifiedGroup $dl.Destination_PrimarySmtpAddress -EA SilentlyContinue)) {
		Write-Host "creating ... " -ForegroundColor Yellow -NoNewline
		try {
			$newGroup = New-UnifiedGroup -Name $dl.Name -Alias $dl.Alias -DisplayName $dl.DisplayName -PrimarySmtpAddress $dl.Destination_PrimarySmtpAddress -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($dl.RequireSenderAuthenticationEnabled)) -ErrorAction Stop -Confirm:$false
		}
		catch {
            $currenterror = new-object PSObject

            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $dl.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $dl.Destination_PrimarySmtpAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToCreate" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $AllErrors += $currentuser
			Write-Host "fail to create" -ForegroundColor Red
			continue
		}
        try {
			$newGroup | Set-UnifiedGroup -HiddenFromAddressListsEnabled $true -Confirm:$false
            $newGroup | Set-UnifiedGroup -AccessType $dl.AccessType
            $newGroup | Set-UnifiedGroup -Notes $dl.Notes -Confirm:$false -ErrorAction Stop
		}
		catch {
            $currenterror = new-object PSObject

            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $dl.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $dl.Destination_PrimarySmtpAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdate" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			Write-Host "failed to update" -ForegroundColor Red
            $AllErrors += $currentuser
			continue
		}
		Write-Host "done" -ForegroundColor Green
	} 
    else {
		Write-Host "exists" -ForegroundColor DarkGreen
        try {
			$newGroup | Set-UnifiedGroup -HiddenFromAddressListsEnabled $true -Confirm:$false
            $newGroup | Set-UnifiedGroup -AccessType $dl.AccessType
            $newGroup | Set-UnifiedGroup -Notes $dl.Notes -Confirm:$false -ErrorAction Stop
		}
		catch {
            $currenterror = new-object PSObject

            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $dl.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $dl.Destination_PrimarySmtpAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdate" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			Write-Host "failed to update" -ForegroundColor Red
            $AllErrors += $currentuser
			continue
		}
	}
}

#Create New Office365 Groups
$AllErrorsGroups = @()
$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($dl in $O365Groups) {

    #Set Variables
    $Addresssplit = $dl.PrimarySMTPAddress_Source -split "@"
    $DestinationPrimarySMTPAddress = $addressSplit[0] + "@example.org"
    $destinationDisplayName = $dl.DisplayName_Source + " - EHN"

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Office 365 Group $($DestinationPrimarySMTPAddress)"

    $newGroup = @()
	Write-Host "$($DestinationPrimarySMTPAddress) ... " -ForegroundColor Cyan -NoNewline
	
	if (!($groupCheck = Get-UnifiedGroup $DestinationPrimarySMTPAddress -EA SilentlyContinue)) {
		Write-Host "creating ... " -ForegroundColor Yellow -NoNewline

		try {
            $newGroup = New-UnifiedGroup -Name $destinationDisplayName -DisplayName $destinationDisplayName -PrimarySmtpAddress $DestinationPrimarySMTPAddress -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($dl.RequireSenderAuthenticationEnabled)) -Confirm:$false
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

        try {
			$newGroup | Set-UnifiedGroup -AccessType $dl.AccessType -Confirm:$false
            $newGroup | Set-UnifiedGroup -Notes $dl.Notes -Confirm:$false -ErrorAction Stop
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
    else {
		Write-Host "exists" -ForegroundColor DarkGreen -NoNewline
        try {
			$groupCheck | Set-UnifiedGroup -AccessType $dl.AccessType -Confirm:$false
            $groupCheck | Set-UnifiedGroup -Notes $dl.Notes -Confirm:$false -ErrorAction Stop
            Write-Host "done" -ForegroundColor Green
		}
		catch {
            $currenterror = new-object PSObject

            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $groupCheck.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdate" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			Write-Host ".. failed to update" -ForegroundColor Red
            $AllErrorsGroups += $currenterror
			continue
		}
        
	}
}

## match Groups to NewCompany from OldCompany
$O365Groups = Import-Csv 
$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($group in $O365Groups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($group.DisplayName)"

    $groupCheck = @()
    if ($groupCheck = Get-UnifiedGroup $group.Destination_PrimarySMTPAddress -ea silentlycontinue) {
        $groupOwners = Get-UnifiedGroupLinks $groupCheck.PrimarySMTPAddress -LinkType Owner -ea silentlycontinue
        $groupMembers = Get-UnifiedGroupLinks $groupCheck.PrimarySMTPAddress -LinkType Member -ea silentlycontinue
        $EmailAddresses = $groupCheck | select -ExpandProperty EmailAddresses

        $group | add-member -type noteproperty -name "ExistsInDestination" -Value $true -Force
        $group | add-member -type noteproperty -name "DisplayName_destination" -Value $groupCheck.DisplayName -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $groupCheck.RecipientTypeDetails -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $groupCheck.PrimarySmtpAddress -Force
        $group | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $group | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $groupCheck.legacyexchangedn) -Force
        $group | add-member -type noteproperty -name "GroupOwners_destination" -Value ($groupOwners -join ";") -Force
        $group | add-member -type noteproperty -name "GroupMembers_destination" -Value ($groupMembers -join ";") -Force
    }
    elseif ($groupCheck = Get-Recipient $group.DisplayName -ea silentlycontinue) {
        $groupOwners = Get-UnifiedGroupLinks $groupCheck.PrimarySMTPAddress -LinkType Owner -ea silentlycontinue
        $groupMembers = Get-UnifiedGroupLinks $groupCheck.PrimarySMTPAddress -LinkType Member -ea silentlycontinue
        $EmailAddresses = $groupCheck | select -ExpandProperty EmailAddresses
        
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value $true -Force
        $group | add-member -type noteproperty -name "DisplayName_destination" -Value $groupCheck.DisplayName -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $groupCheck.RecipientTypeDetails -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $groupCheck.PrimarySmtpAddress -Force
        $group | add-member -type noteproperty -name "EmailAddresses_destination" -Value ($EmailAddresses -join ";") -Force
        $group | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value ("x500:" + $groupCheck.legacyexchangedn) -Force
        $group | add-member -type noteproperty -name "GroupOwners_destination" -Value ($groupOwners -join ";") -Force
        $group | add-member -type noteproperty -name "GroupMembers_destination" -Value ($groupMembers -join ";") -Force
    }
    else {
        $group | add-member -type noteproperty -name "ExistsInDestination" -Value $false -Force
        $group | add-member -type noteproperty -name "DisplayName_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "RecipientTypeDetails_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "PrimarySmtpAddress_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "EmailAddresses_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "LegacyExchangeDN_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "GroupOwners_destination" -Value $null -Force
        $group | add-member -type noteproperty -name "GroupMembers_destination" -Value $null -Force
    }
}

## Update GroupsData in OldCompany with Owners and Members
$O365Groups = Import-Csv 
$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($group in $O365Groups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($group.DisplayName)"

    $groupCheck = @()
    if ($groupCheck = Get-UnifiedGroup $group.PrimarySMTPAddress -ea silentlycontinue) {
        $groupOwners = Get-UnifiedGroupLinks $group.PrimarySMTPAddress -LinkType Owner -ea silentlycontinue
        $groupMembers = Get-UnifiedGroupLinks $group.PrimarySMTPAddress -LinkType Member -ea silentlycontinue
        $EmailAddresses = $groupCheck | select -ExpandProperty EmailAddresses

        $group | add-member -type noteproperty -name "GroupOwners" -Value ($groupOwners -join ";") -Force
        $group | add-member -type noteproperty -name "GroupMembers" -Value ($groupMembers -join ";") -Force
    }
}

## Update GroupsData in OldCompany with Owners
$O365Groups = Import-Csv 
$progressref = ($O365Groups | ?{$_.GroupOwners}).count
$progresscounter = 0
foreach ($group in ($O365Groups | ?{$_.GroupOwners})) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Owner Details for $($group.DisplayName)"

    $Owners = $group.GroupOwners -split ","
    $groupOwners = @()
    foreach ($owner in $Owners) {
        $mailboxCheck = Get-Mailbox $owner
        $groupOwners += $mailboxCheck.primarysmtpaddress
    }
    $group | add-member -type noteproperty -name "GroupOwners" -Value ($groupOwners -join ";") -Force
}

# Add Migration Account as Owner - OldCompany
$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($group in $O365Groups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Add Migration Service Account to $($group.DisplayName)"
    Add-UnifiedGroupLinks -Identity $group.name -LinkType Member -Links "Migration_serviceaccount1@ehn.onmicrosoft.com"
    Add-UnifiedGroupLinks -Identity $group.name -LinkType Owner -Links "Migration_serviceaccount1@ehn.onmicrosoft.com" 
}

# Add Migration Account as Owner - NewCompany
$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($group in $O365Groups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Add Migration Service Account to $($group.DisplayName)"
    Add-UnifiedGroupLinks -Identity $group.name -LinkType Member -Links "MigrationSvc@tjuv.onmicrosoft.com"
    Add-UnifiedGroupLinks -Identity $group.name -LinkType Owner -Links "MigrationSvc@tjuv.onmicrosoft.com" 
}

# Add Migration Account as Owner - NewCompany - REmaining EHN accounts
$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($group in $O365Groups) {
    #Set Variables
    $Addresssplit = $dl.PrimarySMTPAddress_Source -split "@"
    $DestinationPrimarySMTPAddress = $addressSplit[0] + "@example.org"
    $destinationDisplayName = $dl.DisplayName_Source + " - EHN"

    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Add Migration Service Account to $($destinationDisplayName)"
    Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Member -Links "MigrationSvc@tjuv.onmicrosoft.com"
    Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Owner -Links "MigrationSvc@tjuv.onmicrosoft.com" 
}

# Add Migration Account as Owner - OldCompany - From MigWiz Report
$progressref = ($addOwnersToTeams).count
$progresscounter = 0
foreach ($group in $addOwnersToTeams) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Add Migration Service Account to $($group.SourceEmailAddress)"
    Add-UnifiedGroupLinks -Identity $group.SourceEmailAddress -LinkType Member -Links "Migration_serviceaccount1@ehn.onmicrosoft.com"
    Add-UnifiedGroupLinks -Identity $group.SourceEmailAddress -LinkType Owner -Links "Migration_serviceaccount1@ehn.onmicrosoft.com" 
}
# Add Migration Account as Owner - NewCompany - From MigWiz Report
$progressref = ($addOwnersToTeams).count
$progresscounter = 0
foreach ($group in $addOwnersToTeams) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Add Migration Service Account to $($group.DestinationEmailAddress)"
    Add-UnifiedGroupLinks -Identity $group.DestinationEmailAddress -LinkType Member -Links "MigrationSvc@tjuv.onmicrosoft.com"
    Add-UnifiedGroupLinks -Identity $group.DestinationEmailAddress -LinkType Owner -Links "MigrationSvc@tjuv.onmicrosoft.com" 
}

# Add Members to Office365 Groups
$allO365Groups = Import-Csv
$matchedMailboxes = Import-Csv
$AllErrors_Members = @()

$progressref = ($allO365Groups).count
$progresscounter = 0
foreach ($group in $allO365Groups) {
    #Set Variables
    $DestinationPrimarySMTPAddress = $group.PrimarySMTPAddress_Destination
    $destinationDisplayName = $group.DisplayName_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Members to $($destinationDisplayName)"

    Write-Host "Updating Group Members for $($destinationDisplayName).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Members
    $membersArray = $group.Members -split ","

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
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding Member $($memberCheck.PrimarySmtpAddress_Destination)"

        #Add Member to Office365 Group        
        try {
			Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Member -Links $memberCheck.PrimarySmtpAddress_Destination -ea Stop
            Write-Host "." -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host "." -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $destinationDisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors_Members += $currenterror           
			continue
		}
    }
    Write-Host " done " -ForegroundColor Green
}

# Add Owners to Office365 Groups
$allO365Groups = Import-Csv
$matchedMailboxes = Import-Csv
$AllErrors = @()

$progressref = ($allO365Groups | ?{$_.GroupOwners}).count
$progresscounter = 0
foreach ($group in ($allO365Groups | ?{$_.GroupOwners})) {
    #Set Variables
    $DisplayName = $group.DisplayName_Source
    $PrimarySMTPAddressDestination = $group.PrimarySmtpAddress_Destination
    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Owners to $($DisplayName)"

    Write-Host "Updating Group Owners for $($DisplayName).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Members
    $membersArray = $group.GroupOwners -split ","

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
			Add-UnifiedGroupLinks -Identity $PrimarySMTPAddressDestination -LinkType Owner -Links $memberCheck.PrimarySmtpAddress_Destination -ea Stop
            Write-Host "." -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host "." -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddOwner" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "Group_Destination_PrimarySMTPAddress" -Value $PrimarySMTPAddressDestination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Owner_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "Owner_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror           
			continue
		}
        
    }
    Write-Host " done " -ForegroundColor Green
}

# Add Members and Owners to Office365 Groups - EHN
$O365Groups = Import-Csv
$matchedMailboxes = Import-Csv
$AllErrors_Groups = @()

$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($group in $O365Groups) {
    #Set Variables
    $Addresssplit = $group.PrimarySMTPAddress_Source -split "@"
    $DestinationPrimarySMTPAddress = $addressSplit[0] + "@example.org"
    $destinationDisplayName = $group.DisplayName_Source + " - EHN"

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Members and Owners to $($destinationDisplayName)"
    Write-Host "Updating Group $($destinationDisplayName).. " -ForegroundColor Cyan -NoNewline

    if ($group.Members_Source) {
        $membersArray = $group.Members_Source -split ","
        Write-Host "Adding $($membersArray.count) Members .. " -ForegroundColor Cyan -NoNewline
        #Progress Bar 2
        $progressref2 = ($membersArray).count
        $progresscounter2 = 0
        foreach ($member in $membersArray) {
            #Member Check
            $memberAddress = @()
            if ($member -like "*@og-example.org*") {
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
                Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Member -Links $memberAddress -ea Stop
                Write-Host "." -ForegroundColor Green -NoNewline
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

    if ($group.GroupOwners_Source) {
        $membersArray = $group.GroupOwners_Source -split ","
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
                Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Owner -Links $memberCheck.PrimarySmtpAddress_Destination -ea Stop
                Write-Host "." -ForegroundColor Green -NoNewline
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

    Write-Host " done " -ForegroundColor Green
}


# Add Members to Exchange Distribution Groups
$matchedDistributionGroups = Import-Csv
$matchedMailboxes = Import-Csv
$AllErrors = @()
$progressref = ($matchedDistributionGroups).count
$progresscounter = 0
foreach ($group in $matchedDistributionGroups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Members to $($group.DisplayName_Destination)"

    Write-Host "Updating Group Members for $($group.DisplayName_Destination).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Members
    $membersArray = $group.Members -split ","

    #Progress Bar 2
    $progressref2 = ($membersArray).count
    $progresscounter2 = 0
    foreach ($member in $membersArray) {
        #Member Check
        $memberCheck = @()
        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

        #Progress Bar 2a
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding Member $($memberCheck.PrimarySmtpAddress_Destination)"

        #Add Member to Distribution Groups     
        try {
			Add-DistributionGroupMember -Identity $group.PrimarySmtpAddress_Destination -Member $memberCheck.PrimarySmtpAddress_Destination -ea Stop
            Write-Host ". " -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host ". " -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $group.DisplayName_destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $group.PrimarySmtpAddress_destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror           
			continue
		}
        
    }
    Write-Host " done " -ForegroundColor Green
}

# Add Owners to Exchange Distribution Groups
$allO365Groups = Import-Csv
$matchedMailboxes = Import-Csv
$AllErrors = @()
$progressref = ($allO365Groups).count
$progresscounter = 0
foreach ($group in $allO365Groups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Owners to $($group.DisplayName_destination)"

    Write-Host "Updating Group Owners for $($group.DisplayName_destination).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Members
    $membersArray = $group.ManagedBy -split ","

    #Progress Bar 2
    $progressref2 = ($membersArray).count
    $progresscounter2 = 0
    foreach ($member in $membersArray) {
        #Member Check
        $memberCheck = @()
        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

        #Progress Bar 2a
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Adding Owner $($memberCheck.PrimarySmtpAddress_Destination)"

        #Add Owner to Distribution Groups      
        try {
			Set-DistributionGroup -Identity $group.PrimarySmtpAddress_destination -ManagedBy @{add=$memberCheck.PrimarySmtpAddress_Destination} -ea Stop
            Write-Host ". " -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host ". " -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddOwner" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $group.DisplayName_destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $group.PrimarySmtpAddress_destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Owner_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "Owner_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror           
			continue
		}
        
    }
    Write-Host " done " -ForegroundColor Green
}

# Remove Migration Account as Owner DistributionGroups - NewCompany
$progressref = ($allO365Groups).count
$progresscounter = 0
foreach ($group in $allO365Groups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Remove Migration Service Account to $($group.DisplayName_Source)"

    if ($group.RecipientTypeDetails_Source -eq "MailUniversalDistributionGroup" -and $group.ResourceProvisioningOptions_Source -notlike "*eam") {
        $managedBy = (Get-DistributionGroup -Identity $group.PrimarySmtpAddress_destination).ManagedBy
        if ($managedBy.count -gt 1) {
            try{
                Set-DistributionGroup -Identity $group.PrimarySmtpAddress_destination -ManagedBy @{remove="MigrationSvc"} -ea Stop
                Write-Host "Removed Migration Account" -foregroundcolor green
            }
            Catch {
                Write-Error "Unable To Remove Migration Account"
            }
        }
        else {
            Write-Host "Migration Account Only Manager" -foregroundcolor Yellow
        }
    }
}

# Create New Shared Mailboxes
$progressref = ($newSharedMailboxes).count
$progresscounter = 0
$AllErrors = @()
foreach ($user in $newSharedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating New Shared Mailbox $($user.DisplayName)"
    $addressSplit = $user.PrimarySmtpAddress -split "@"
    $newPrimarySMTPAddress = $addressSplit[0] + "-old@example.org"
    $newDisplayName = $user.DisplayName + " - EHN"
    $newAlias = $user.alias +"-old"
    Write-Host "Creating Shared Mailbox $($newPrimarySMTPAddress) " -ForegroundColor Cyan -NoNewline
    try {
        $newMailbox = New-Mailbox -Shared -name $newDisplayName -DisplayName $newDisplayName -alias $newAlias -PrimarySmtpAddress $newPrimarySMTPAddress -ea Stop
        Write-Host ".. created" -ForegroundColor Green
    }
    catch {
        $currenterror = new-object PSObject
        
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
        $currenterror | Add-Member -type NoteProperty -Name "SharedMailbox" -Value $user.DisplayName -Force
        $currenterror | Add-Member -type NoteProperty -Name "Source_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress -Force
        $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $newPrimarySMTPAddress -Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        Write-Host ".. failed to create" -ForegroundColor Red
        $AllErrors += $currenterror
        continue
    }
}

# Update Mail Contacts for Shared Mailboxes (fixes conflicts)
$progressref = ($newSharedMailboxes).count
$progresscounter = 0
$AllErrors = @()
foreach ($user in $newSharedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Existing Contact $($user.DisplayName)"
    $newName = $user.DisplayName + " - EHN"
    $oldName = $user.DisplayName
    Write-Host "Updating Existing Contact $($user.DisplayName) " -ForegroundColor Cyan -NoNewline
    try {
        Set-MailContact -Identity $oldName -Name $newName -ea Stop
        Write-Host ".. updated" -ForegroundColor Green
    }
    catch {
        $currenterror = new-object PSObject
        
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
        $currenterror | Add-Member -type NoteProperty -Name "SharedMailbox" -Value $user.DisplayName -Force
        $currenterror | Add-Member -type NoteProperty -Name "Source_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress -Force
        #$currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $newPrimarySMTPAddress -Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        Write-Host ".. failed to update" -ForegroundColor Red
        $AllErrors += $currenterror
        continue
    }
}

#Recheck Wave
$WaveGroup = import-csv
$notfoundUsers = @()
$foundRecipient = @()
$progressref = ($WaveGroup).count
$progresscounter = 0
foreach ($object in $WaveGroup) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($object.UserPrincipalName)"

    if ($msolUserCheck = Get-MsolUser -userprincipalname $object.UserPrincipalName -ea silentlycontinue){
        $mailboxCheck = Get-Mailbox $msolUserCheck.userPrincipalName -EA SilentlyContinue
        $object | add-member -type noteproperty -name "FoundInOldCompany" -Value $true -force
        $object | add-member -type noteproperty -name "FoundDisplayName" -Value $msolUserCheck.DisplayName -force
        $object | add-member -type noteproperty -name "FoundUserPrincipalName" -Value $msolUserCheck.UserPrincipalName -force
        $object | add-member -type noteproperty -name "IsLicensed" -Value $msolUserCheck.IsLicensed -force
        $object | add-member -type noteproperty -name "Department" -Value $msolUserCheck.Department -force
        $object | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mailboxCheck.RecipientTypeDetails -force
        $object | add-member -type noteproperty -name "FoundSMTPAddress" -Value $mailboxCheck.primarysmtpaddress -force
        $object | add-member -type noteproperty -name "CustomAttribute7" -Value $mailboxCheck.CustomAttribute7 -force       
    }   
    elseif ($mailboxCheck = Get-Mailbox $object.PrimarySMTPAddress -ea silentlycontinue) {
        $msolUserCheck   = Get-MsolUser -userprincipalname $mailboxCheck.UserPrincipalName -EA SilentlyContinue
        $object | add-member -type noteproperty -name "FoundInOldCompany" -Value $mailboxCheck.RecipientTypeDetails -force
        $object | add-member -type noteproperty -name "FoundDisplayName" -Value $msolUserCheck.DisplayName -force
        $object | add-member -type noteproperty -name "FoundUserPrincipalName" -Value $msolUserCheck.UserPrincipalName -force
        $object | add-member -type noteproperty -name "IsLicensed" -Value $msolUserCheck.IsLicensed -force
        $object | add-member -type noteproperty -name "Department" -Value $msolUserCheck.Department -force
        $object | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mailboxCheck.RecipientTypeDetails -force
        $object | add-member -type noteproperty -name "FoundSMTPAddress" -Value $mailboxCheck.primarysmtpaddress -force
        $object | add-member -type noteproperty -name "CustomAttribute7" -Value $mailboxCheck.CustomAttribute7 -force   
    }
    elseif ($recipientCheck = Get-Recipient $object.PrimarySMTPAddress -ea silentlycontinue) {
        $object | add-member -type noteproperty -name "FoundInOldCompany" -Value $recipientCheck.RecipientTypeDetails -force
        $object | add-member -type noteproperty -name "FoundDisplayName" -Value $recipientCheck.DisplayName -force
        $object | add-member -type noteproperty -name "FoundUserPrincipalName" -Value $null -force
        $object | add-member -type noteproperty -name "IsLicensed" -Value $null -force
        $object | add-member -type noteproperty -name "Department" -Value $null -force
        $object | add-member -type noteproperty -name "RecipientTypeDetails" -Value $null -force
        $object | add-member -type noteproperty -name "FoundSMTPAddress" -Value $null -force
        $object | add-member -type noteproperty -name "CustomAttribute7" -Value $null -force 
        $foundRecipient += $object
    }
    else {
        $object | add-member -type noteproperty -name "FoundInOldCompany" -Value $false -force
        $object | add-member -type noteproperty -name "FoundDisplayName" -Value $null -force
        $object | add-member -type noteproperty -name "FoundUserPrincipalName" -Value $null -force
        $object | add-member -type noteproperty -name "IsLicensed" -Value $null -force
        $object | add-member -type noteproperty -name "Department" -Value $null -force
        $object | add-member -type noteproperty -name "RecipientTypeDetails" -Value $null -force
        $object | add-member -type noteproperty -name "FoundSMTPAddress" -Value $null -force
        $object | add-member -type noteproperty -name "CustomAttribute7" -Value $null -force 
    
        $notfoundUsers += $object
    }
}

#Match Wave Mailboxes and add to same spreadsheet. Check based on CustomAttribute7 and DisplayName
$WaveGroup = Import-Csv
$progressref = ($WaveGroup).count
$progresscounter = 0

foreach ($user in $WaveGroup) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    $NewUPN = $user.CustomAttribute7 + "@example.org"
    $addressSplit = $user.PrimarySmtpAddress -split "@"
    $ehnPrimarySMTPAddress = $addressSplit[0] + "-old@example.org"

    #NEW UPN Check
    if ($msoluser = Get-Msoluser -UserPrincipalName $NewUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with CustomAttribute7" -ForegroundColor Green
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
    
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "CampusKeyMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $mailbox.ArchiveStatus -force
        $user | Add-Member -Type NoteProperty -Name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
    }
    #NEW PrimarySMTPAddress Check (EHN added)
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $ehnPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with EHN added" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "EHNMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed   -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $mailbox.ArchiveStatus -force
        $user | Add-Member -Type NoteProperty -Name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
    }
    #NEW PrimarySMTPAddress Check
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $newPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "NEWSMTPAddressCheck" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed   -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $mailbox.ArchiveStatus -force
        $user | Add-Member -Type NoteProperty -Name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
    }
    #UPN Check based on DisplayName
    elseif ($msoluser = Get-Msoluser -SearchString $user.DisplayName -ErrorAction SilentlyContinue) {
        if ($msoluser.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName)"
            $mailbox = Get-Mailbox $user.DisplayName -ErrorAction SilentlyContinue
            $msoluser = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        }
        Write-Host "$($msoluser.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $mbxStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "DisplayNameMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluser.IsLicensed   -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $mailbox.ArchiveStatus -force
        $user | Add-Member -Type NoteProperty -Name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
    }
    else {
        Write-Host "  Unable to find user for $($user.DisplayName)" -ForegroundColor Red
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "LastLogonTime_Destination" -Value $null -force
    }
}

#Enable Archive Mailbox
$matchedMailboxes = Import-Csv
$wave1 = Import-Csv
$enableArchiveUsers = $wave1 | ?{$_.ExistsInDestination -eq "CampusKeyMatch"}
$progressref = ($wave2).count
$progresscounter = 0

foreach ($user in $wave2) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Enabling Archive Mailbox for $($user.DisplayName_Destination)"
    try {
        Enable-Mailbox $user.PrimarySMTPAddress_Destination -Archive -ea stop
    }
    catch {
        Write-Error "$($_.Exception)"
    }
}

# Set Shared Mailbox to RoomMailbox
$progressref = ($nonUserMailboxes).count
$progresscounter = 0
foreach ($mbx in $nonUserMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Mailbox Type for $($mbx.DisplayName_Destination)"    
    If ($mbx.RecipientTypeDetails -eq "RoomMailbox"){
        Write-Host "Converted $($mbx.DisplayName_Destination) to RoomMailbox " -foregroundcolor Green
        Set-Mailbox $mbx.PrimarySmtpAddress_Destination -Type Room
    }
    elseif ($mbx.RecipientTypeDetails -eq "EquipmentMailbox"){
        Write-Host "Converted $($mbx.DisplayName_Destination) to EquipmentMailbox " -foregroundcolor Green
        Set-Mailbox $mbx.PrimarySmtpAddress_Destination -Type Equipment
    }
    else {
        Write-Host "Skipping $($mbx.DisplayName_Destination) " -foregroundcolor yellow
    }
}

# Create Mail Contacts On Premises for Cloud Only objects
$cloudONlyObjectsString = Read-Host "What is the CSV File to Import"
$cloudONlyObjects = Import-CSV $cloudONlyObjectsString
$progressref = ($cloudONlyObjects).count
$progresscounter = 0
foreach ($object in $cloudONlyObjects) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Mail Contact for $($object.DisplayName_Destination)"    
    Write-Host "Creating Mail Contact $($object.DisplayName_Destination) .. " -foregroundcolor Cyan -nonewline
    $newContact = New-MailContact -DisplayName $object.DisplayName_Destination -Name $object.DisplayName_Destination -ExternalEmailAddress $object.PrimarySmtpAddress_Destination
    Write-Host "done .. " -foregroundcolor Green
}

# Update Mail Contacts On Premises for Cloud Only objects
$cloudONlyObjectsString = Read-Host "What is the CSV File to Import"
$cloudONlyObjects = Import-CSV $cloudONlyObjectsString
$progressref = ($cloudONlyObjects).count
$progresscounter = 0
foreach ($object in $cloudONlyObjects) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Mail Contact for $($object.DisplayName_Destination) CustomAttribute2 with NewCompany-edu.mail.protection.outlook.com"    
    Write-Host "Updated Mail Contact $($object.DisplayName_Destination) .. " -foregroundcolor Cyan -nonewline
    $objectCheck = @()
    try {
        if ($objectCheck = Get-MailContact $object.PrimarySmtpAddress_Destination -ea stop) {
            Set-MailContact -Identity $objectCheck.PrimarySmtpAddress.tostring() -CustomAttribute2 "NewCompany-edu.mail.protection.outlook.com"
        }
        else {
            Write-Error "Unable to find contact $($ $object.PrimarySmtpAddress_Destination)"
        }
    }
    catch {
        Write-Error $_.Exception
    }
    Write-Host "done" -foregroundcolor Green
}

# Stamp Email Addresses on Mail Contacts for Cloud Only Objects
$progressref = ($cloudONlyObjects).count
$progresscounter = 0
foreach ($object in $cloudONlyObjects) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Mail Contact for $($object.DisplayName_Destination)"    

    Set-MailContact -Identity $object.PrimarySmtpAddress_Destination -CustomAttribute2 "NewCompany-edu.mail.protection.outlook.com"

    $emailAddressArray = $object.EmailAddresses -split ","
    foreach ($address in $emailAddressArray) {
        Set-MailContact -Identity $object.PrimarySmtpAddress_Destination -EmailAddresses @{add=$address}
        
    }
}

#Get DistributionGroup Members
$1kMembersGroups = import-csv
$1kGroupMembersAll = @()

#ProgressBarA
$progressref = ($1kMembersGroups).count
$progresscounter = 0
foreach ($object in $1kMembersGroups) {
    #ProgressBarB
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1  -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Group Members for $($object.DisplayName)"
    
    $members = Get-DistributionGroupMember $object.PrimarySMTPAddress -ResultSize unlimited

    $progressref2 = ($members).count
    $progresscounter2 = 0
    foreach ($member in $members) {
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2  -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Member Details for $($member)"
        
        #Create Output Array
        $memberDetails = Get-Recipient $member.PrimarySMTPAddress | select PrimarySMTPAddress, RecipientTypeDetails
        $currentobject = new-object PSObject
        $currentobject | add-member -type noteproperty -name "DisplayName" -Value $object.DisplayName -Force
        $currentobject | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $object.PrimarySMTPAddress -Force
        $currentobject | add-member -type noteproperty -name "Member" -Value $memberDetails.PrimarySMTPAddress -Force
        $currentobject | add-member -type noteproperty -name "MemberType" -Value $memberDetails.RecipientTypeDetails -Force

        $currentobject | Export-Csv -NoTypeInformation -Encoding utf8 C:\Users\amedrano\Desktop\1kGroupsFullMembers.csv -Append
    }
    
    $1kGroupMembersAll += $currentobject
}

#new Shared Mailbox Perm stamp
# Grant SendAs Perms to Mailboxes
$OldCompanyNonUserMailboxes = import-csv
$matchedMailboxes = Import-Csv
$fullAccessPermUsers = $OldCompanyNonUserMailboxes | ? {$_.FullAccessPerms -and $_.ExistsInDestination -ne $false}
$sendAsPermUsers =  $OldCompanyNonUserMailboxes | ? {$_.SendAsPerms -and $_.ExistsInDestination -ne $false}
$AllErrors = @()

$progressref = ($fullAccessPermUsers).count
$progresscounter = 0
foreach ($user in $fullAccessPermUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding SendAs Perms to $($user.DisplayName_Destination)"
    Write-Host "Grant SendAs Perms for $($user.DisplayName_Destination).. " -ForegroundColor Cyan -NoNewline
    
    #Gather Perm Users
    $membersArray = $user.SendAsPerms -split ","

    #Progress Bar 2
    $progressref2 = ($membersArray).count
    $progresscounter2 = 0
    foreach ($member in $membersArray) {
        #Member Check
        $memberCheck = @()
        $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $member}

        #Progress Bar 2a
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting SendAs to $($memberCheck.PrimarySmtpAddress_Destination)"

        #Add Perms to Mailbox      
        try {
			$permResult = Add-RecipientPermission $user.PrimarySmtpAddress_Destination -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -confirm:$false -ea Stop
            Write-Host ". " -ForegroundColor Green -NoNewline
		}
		catch {
            Write-Host ". " -ForegroundColor red -NoNewline

            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantFullAccess" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $user.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Source_PrimarySMTPAddress" -Value $member -Force
            $currenterror | Add-Member -type NoteProperty -Name "PermUser_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror           
			continue
		}
    }
    Write-Host " done " -ForegroundColor Green
}

# Add Members to Exchange Distribution Groups Individual over 1K Users
$matchedMailboxes = Import-Csv
$1kGroups = Import-Csv
$AllGroupErrors = @()
$groupCheck = "UrbanCoreEEMI@example.org"
$progressref = ($1kGroups | ?{$_.PrimarySMTPAddress_Destination -eq $groupCheck -and $_.Member -like "*og-example.org"}).count
$progresscounter = 0

Write-Host "Updating Group Members for $($groupCheck).. " -ForegroundColor Cyan -NoNewline
foreach ($object in ($1kGroups | ?{$_.PrimarySMTPAddress_Destination -eq $groupCheck -and $_.Member -like "*og-example.org"})) {
    #Member Check
    $memberCheck = @()
    $memberCheck = $matchedMailboxes | ? {$_.PrimarySmtpAddress -eq $object.member}

    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Member $($object.Member)"

    #Add Member to Distribution Group        
    try {
        Add-DistributionGroupMember -Identity $groupCheck -Member $memberCheck.PrimarySmtpAddress_Destination -ea Stop
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    catch {
        Write-Host ". " -ForegroundColor red -NoNewline

        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
        $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $object.PrimarySMTPAddress -Force
        $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $groupGroupCheck -Force
        $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $object.Member -Force
        $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllGroupErrors += $currenterror           
        continue
    }
}

# Update Members for Wave 1 Group

## Add Members to Exchange Distribution Groups
$WaveGroup = Import-Csv
$progressref = ($WaveGroup).count
$progresscounter = 0
$waveDLGroup = "Wave5migration@og-example.org"
foreach ($user in $WaveGroup) {
    $PrimarySMTPAddress =  $user.PrimarySmtpAddress_Source 
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Member $($PrimarySMTPAddress) to $($waveDLGroup)"

    #Add Member to Distribution Groups
    Add-DistributionGroupMember -Identity $waveDLGroup -Member $PrimarySMTPAddress -EA silentlycontinue
}

## Remove members to Wave 1
$progressref = ($groupmembers2).count
$progresscounter = 0
foreach ($group in $groupmembers2) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Removing Member  $($group.PrimarySMTPAddress)"

    #Add Member to Distribution Groups
    Remove-DistributionGroupMember -Identity "Wave1migration@og-example.org" -Member $group.PrimarySMTPAddress -confirm:$false
}


#Add SharePoint Details for Wave- Source
$progressref = ($WaveGroup).count
$progresscounter = 0
Write-Host "Updating SharePoint Details " -foregroundcolor Cyan -nonewline
foreach ($object in $WaveGroup) {
    $UPN = $object.userPrincipalName
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($UPN)"

    $count = 0
    $success = $null

    do{
        try{
            $OneDriveSite = Get-SPOSite -Filter "Owner -eq '$UPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop
            if ($OneDriveSite) {
                $object | add-member -type noteproperty -name "SourceOneDriveURL" -Value $OneDriveSite.url -force
                Write-Host ". " -foregroundcolor green -nonewline
                $success = $true
            }
           else {
                $object | add-member -type noteproperty -name "SourceOneDriveURL" -Value $null -force
                Write-Host ". " -foregroundcolor red -nonewline
                $failed = $true
           }
        }
        catch{
            Write-host "Next attempt in 5 seconds" -foregroundcolor yellow -nonewline
            Start-sleep -Seconds 5
            $count++
        # Put the start-sleep in the catch statemtnt so we
        # don't sleep if the condition is true and waste time
        }
        
    }
    until($count -eq 5 -or $success -or $failed)

    if(!($success -or $failed)) {
        $object | add-member -type noteproperty -name "SourceOneDriveURL" -Value $null -force
        Write-Host ". " -foregroundcolor red -nonewline
    }
}
Write-Host "done" -foregroundcolor green

#Add SharePoint Details - Destination
#Recheck Wave 2
$progressref = ($wave2).count
$progresscounter = 0
Write-Host "Updating SharePoint Details " -foregroundcolor Cyan -nonewline
foreach ($object in $wave2) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($object.UserPrincipalName_Destination)"

    $DestinationUPN = $object.UserPrincipalName_Destination
    $count = 0
    $success = $null

    if ($object.DestinationOneDriveURL) {
        Write-Host ". " -foregroundcolor yellow -nonewline
    }
    else {
        do{
            try{
                $OneDriveSite = Get-SPOSite -Filter "Owner -eq '$DestinationUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop
                if ($OneDriveSite) {
                    $object | add-member -type noteproperty -name "DestinationOneDriveURL" -Value $OneDriveSite.url -force
                    Write-Host ". " -foregroundcolor green -nonewline
                    $success = $true
                }
               else {
                    $object | add-member -type noteproperty -name "DestinationOneDriveURL" -Value $null -force
                    Write-Host ". " -foregroundcolor red -nonewline
                    $failed = $true
               }
            }
            catch{
                Write-host "Next attempt in 5 seconds" -foregroundcolor yellow -nonewline
                Start-sleep -Seconds 5
                $count++
            # Put the start-sleep in the catch statemtnt so we
            # don't sleep if the condition is true and waste time
            }
            
        }
        until($count -eq 5 -or $success -or $failed)
    
        if(!($success -or $failed)) {
            $object | add-member -type noteproperty -name "DestinationOneDriveURL" -Value $null -force
            Write-Host ". " -foregroundcolor red -nonewline
        }
    }  
}
Write-Host "done" -foregroundcolor green


#Wave Cutover Updates
Install-Module -Name ExchangeOnlineManagement
Install-Module MSOnline
Install-Module AzureAD

AzureADPreview\Connect-AzureAD
Connect-MsolService
Connect-ExchangeOnline
Connect-SPOService -Url $AdminURL

function Start-BatchCutoverUpdates {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the CSV File Path of OneDrive Users')] [array] $ImportCSV,
        [Parameter(Mandatory=$True,HelpMessage="Which Tenant to Update (Source or Destination?")] [string] $Tenant,
        [Parameter(Mandatory=$false,HelpMessage="Which Wave Is This?")] [string] $Group,
        [Parameter(Mandatory=$false,HelpMessage="Set Forward?")] [switch] $SetForward,
        [Parameter(Mandatory=$false,HelpMessage="Block Access to Mailbox?")] [switch] $BlockMailAccess,
        [Parameter(Mandatory=$false,HelpMessage="Block Access to OneDrive?")] [switch] $BlockOneDriveAccess,
        [Parameter(Mandatory=$false,HelpMessage="Disable Teams?")] [switch] $DisableTeams,
        [Parameter(Mandatory=$false,HelpMessage="Do you want to Test?")] [switch] $Test,
        [Parameter(Mandatory=$false,HelpMessage="Hide OldCompany Contacts in NewCompany?")] [switch] $HideOldCompanyContact,
        [Parameter(Mandatory=$false,HelpMessage="Force Log Out User?")] [switch] $ForceLogOut
    )
    #Create User Array
    $WaveGroup = Import-csv $ImportCSV

    # Gather User Details
    $progressref = ($WaveGroup).count
    $progresscounter = 0
    $allErrors = @()
        
    foreach ($user in $WaveGroup) {
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($user.DisplayName_Source)"
        $SourceUPN = $user.UserPrincipalName_Source
        $SourcePrimarySMTPAddress = $user.PrimarySMTPAddress_Source
        $DesinationPrimarySMTPAddress = $user.PrimarySmtpAddress_Destination

        if ($test) {
            if ($Tenant -eq "Destination") {
                # NewCompany
                Write-Host "Cutting Over User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
                ## Hide OldCompany Contact In NewCompany
                if ($HideOldCompanyContact) {
                    try {
                        Set-MailContact $SourcePrimarySMTPAddress -HiddenFromAddressListsEnabled $true
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
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
            }
            if ($Tenant -eq "Source") {
                # OldCompany
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
                        Set-CASMailbox $SourcePrimarySMTPAddress -MacOutlookEnabled $false -MAPIEnabled $false -OutlookMobileEnabled $false -OWAEnabled $false -ActiveSyncEnabled $false -OWAforDevices $false -ErrorAction Stop
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
                    $SPOUPN = $UPN.replace("@og-example.org","_OldCompany_edu")
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
                    #$SKUID = ($msoluser).licenses.AccountSkuId
                    $SKUID = "EHN:ENTERPRISEPACK"
                
                    ## Get Disabled Array
                    for($i = 0; $i -lt $AllLicenses.Count; $i++) {
                        $serviceStatus =  $AllLicenses[$i].ServiceStatus
                        foreach($service in $serviceStatus)
                        {
                            if($service.ProvisioningStatus -eq "Disabled")
                            {
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
                Write-Host "Done" -foregroundcolor Green 
            } 
        }
        else {
            if ($Tenant -eq "Destination") {
                # NewCompany
                Write-Host "Cutting Over User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
                ## Hide OldCompany Contact In NewCompany
                if ($HideOldCompanyContact) {
                    try {
                        Set-MailContact $SourcePrimarySMTPAddress -HiddenFromAddressListsEnabled $true
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
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
            }
            if ($Tenant -eq "Source") {
                # OldCompany
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
                        Set-CASMailbox $SourcePrimarySMTPAddress -MacOutlookEnabled $false -MAPIEnabled $false -OutlookMobileEnabled $false -OWAEnabled $false -ActiveSyncEnabled $false -OWAforDevices $false -ErrorAction Stop
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
                    $SPOUPN = $UPN.replace("@og-example.org","_OldCompany_edu")
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
                    #$SKUID = ($msoluser).licenses.AccountSkuId
                    $SKUID = "EHN:ENTERPRISEPACK"
                
                    ## Get Disabled Array
                    for($i = 0; $i -lt $AllLicenses.Count; $i++) {
                        $serviceStatus =  $AllLicenses[$i].ServiceStatus
                        foreach($service in $serviceStatus)
                        {
                            if($service.ProvisioningStatus -eq "Disabled")
                            {
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
                Write-Host "Done" -foregroundcolor Green 
            }
        }
    }
}

#Block OldCompany Guest Users
AzureADPreview\Connect-AzureAD
$GuestUsers = Get-AzureADUser -All $true | Where-Object {$_.UserType -eq 'Guest'}
$progressref = ($GuestUsers).count
$progresscounter = 0
foreach ($guest in $GuestUsers){
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Disable Guest User $($GuestUsers.UserPrincipalName)"
    Set-AzureADUser -ObjectID $guest.UserPrincipalName -AccountEnabled $false
}

#Post Migration Checks
## Initial
AzureADPreview\Connect-AzureAD
Connect-MsolService
Connect-ExchangeOnline

function Get-PostMigrationFullDetails {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the CSV File Path of Wave Users')] [array] $ImportCSV,
        [Parameter(Mandatory=$True,HelpMessage="Which Tenant to Pull Details from?")] [string] $Tenant,
        [Parameter(Mandatory=$True,HelpMessage="Which Wave Is This?")] [string] $Group,
        [Parameter(Mandatory=$false)][string] $ExportFolderPath
    )

    #Create User Array
    $WaveGroup = Import-csv $ImportCSV
    $postMigrationDetails = @()
    Set-Variable PrimarySMTPAddress, UPN

    # Gather User Details
    $progressref = ($WaveGroup).count
    $progresscounter = 0
    foreach ($user in $WaveGroup) {
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($user.DisplayName_Source)"

        #Clear Previous Variables
        $UPN = $null
        $PrimarySMTPAddress = $null
        $MSOLUser = @()
        $azureAuditLogsDetails = @()
        $azureSignInStatus = @()
        $mailboxDetails = @()
        $mailboxStats = @()
        $archiveMailboxStats = @()

        #Set Variables
        $shorterdate = $((Get-Date).ToShortDateString().Replace("/","-"))
        $shortTime = (get-date).ToShortTimeString()

        if ($Tenant -eq "Source") {
            $UPN = $user.UserPrincipalName_Source
            $PrimarySMTPAddress = $user.PrimarySMTPAddress_Source
            #Write-Host "Gathering Source Details .."
        }
        elseif ($Tenant -eq "Destination") {
            $UPN = $user.UserPrincipalName_Destination
            $PrimarySMTPAddress = $user.PrimarySMTPAddress_Destination
            #Write-Host "Gathering Destination Details .."
        }        
        #MSOLUserDetails
        $MSOLUser = Get-MsolUser -UserPrincipalName $UPN

        #Gather Azure Audit Logs
        $azureAuditLogsDetails = Get-AzureADAuditSignInLogs -Filter "UserPrincipalName eq '$UPN'" -Top 1 | Select CreatedDateTime, AppDisplayName, ClientAppUsed, Status
        if ($azureAuditLogsDetails.Status.ErrorCode -eq "0") {
            $azureSignInStatus = "Success"
        }
        elseif ($azureAuditLogsDetails -eq $null) {
            $azureSignInStatus = "NoResults"
        }
        else {
            $azureSignInStatus = $azureAuditLogsDetails.Status.FailureReason
        }

        #Gather Mailbox Details
        $mailboxDetails = Get-Mailbox $PrimarySMTPAddress -ea SilentlyContinue | Select PrimarySMTPAddress, HiddenFromAddressListsEnabled, DeliverToMailboxAndForward, ForwardingSmtpAddress, CustomAttribute7
        $mailboxStats = Get-MailboxStatistics $PrimarySMTPAddress -ea SilentlyContinue | Select LastLogonTime, TotalItemSize, TotalDeletedItemSize, LastInteractionTime
        $archiveMailboxStats = Get-MailboxStatistics $PrimarySMTPAddress -Archive -ea SilentlyContinue | Select TotalItemSize, TotalDeletedItemSize

        #Output

        if ($Tenant -eq "Source") {
            $currentuser = new-object PSObject
            $currentuser | add-member -type noteproperty -name "Date" -Value $shorterdate -Force
            $currentuser | add-member -type noteproperty -name "Time" -Value $shortTime -Force
            $currentuser | Add-Member -type NoteProperty -Name "Wave" -Value $Group -Force
            $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $MSOLUser.DisplayName -Force
            $currentuser | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $user.UserPrincipalName_Source -Force
            $currentuser | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Source" -Value $user.PrimarySMTPAddress_Source -Force
            $currentuser | Add-Member -type NoteProperty -Name "CustomAttribute7" -Value $user.CustomAttribute7 -Force
            $currentuser | Add-Member -type NoteProperty -Name "UserPrincipalName_Destination" -Value $user.UserPrincipalName_Destination -Force
            $currentuser | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $user.PrimarySMTPAddress_Destination -Force
            $currentuser | Add-Member -type NoteProperty -Name "MailboxLastLoginTime_Source" -Value $mailboxStats.LastLogonTime -Force
            $currentuser | Add-Member -type NoteProperty -Name "MailboxLastInteractionTime_Source" -Value $mailboxStats.LastInteractionTime -Force
            $currentuser | Add-Member -type NoteProperty -Name "AzureLastLoginTime_Source" -Value $azureAuditLogsDetails.CreatedDateTime -Force
            $currentuser | Add-Member -type NoteProperty -Name "AzureAppDisplayname_Source" -Value $azureAuditLogsDetails.AppDisplayName -Force
            $currentuser | Add-Member -type NoteProperty -Name "AzureClientAppUsed_Source" -Value $azureAuditLogsDetails.ClientAppUsed -Force
            $currentuser | Add-Member -type NoteProperty -Name "AzureLoginStatus_Source" -Value $azureSignInStatus -Force
            $currentuser | Add-Member -type NoteProperty -Name "IsLicensed_Source" -Value $MSOLUser.IsLicensed -Force
            $currentuser | add-member -type noteproperty -name "Licenses_Source" -Value ($MSOLUser.Licenses.AccountSkuID -join ";") -force
            $currentuser | add-member -type noteproperty -name "License-DisabledArray_Source" -Value ($MSOLUser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
            $currentuser | Add-Member -type NoteProperty -Name "HiddenFromAddressListsEnabled_Source" -Value $mailboxDetails.HiddenFromAddressListsEnabled -Force
            $currentuser | Add-Member -type NoteProperty -Name "DeliverToMailboxAndForward_Source" -Value $mailboxDetails.DeliverToMailboxAndForward -Force
            $currentuser | Add-Member -type NoteProperty -Name "ForwardingSmtpAddress_Source" -Value $mailboxDetails.ForwardingSmtpAddress -Force
            $currentuser | Add-Member -type NoteProperty -Name "CustomAttribute7_Source" -Value $mailboxDetails.CustomAttribute7 -Force
            $currentuser | Add-Member -type NoteProperty -Name "MailboxTotalItemSize_Source" -Value $mailboxStats.TotalItemSize -Force
            $currentuser | Add-Member -type NoteProperty -Name "MailboxTotalDeletedItemSize_Source" -Value $mailboxStats.TotalDeletedItemSize -Force
            $currentuser | Add-Member -type NoteProperty -Name "ArchiveTotalItemSize_Source" -Value $archiveMailboxStats.TotalItemSize -Force
            $currentuser | Add-Member -type NoteProperty -Name "ArchiveTotalDeletedItemSize_Source" -Value $archiveMailboxStats.TotalDeletedItemSize -Force
            $postMigrationDetails += $currentuser
        }
        elseif ($Tenant -eq "Destination") {
            $user | add-member -type noteproperty -name "Date" -Value $shorterdate -Force
            $user | add-member -type noteproperty -name "Time" -Value $shortTime -Force
            $user | Add-Member -type NoteProperty -Name "Wave" -Value $Group -Force
            $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $MSOLUser.DisplayName -Force
            $user | Add-Member -type NoteProperty -Name "UserPrincipalName_Destination" -Value $MSOLUser.UserPrincipalName -Force
            $user | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $mailboxDetails.PrimarySMTPAddress -Force
            $user | Add-Member -type NoteProperty -Name "MailboxLastLoginTime_Destination" -Value $mailboxStats.LastLogonTime -Force
            $user | Add-Member -type NoteProperty -Name "MailboxLastInteractionTime_Destination" -Value $mailboxStats.LastInteractionTime -Force
            $user | Add-Member -type NoteProperty -Name "AzureLastLoginTime_Destination" -Value $azureAuditLogsDetails.CreatedDateTime -Force
            $user | Add-Member -type NoteProperty -Name "AzureAppDisplayname_Destination" -Value $azureAuditLogsDetails.AppDisplayName -Force
            $user | Add-Member -type NoteProperty -Name "AzureClientAppUsed_Destination" -Value $azureAuditLogsDetails.ClientAppUsed -Force
            $user | Add-Member -type NoteProperty -Name "AzureLoginStatus_Destination" -Value $azureSignInStatus -Force
            $user | Add-Member -type NoteProperty -Name "IsLicensed_Destination" -Value $MSOLUser.IsLicensed -Force
            $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($MSOLUser.Licenses.AccountSkuID -join ";") -force
            $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($MSOLUser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
            $user | Add-Member -type NoteProperty -Name "HiddenFromAddressListsEnabled_Destination" -Value $mailboxDetails.HiddenFromAddressListsEnabled -Force
            $user | Add-Member -type NoteProperty -Name "DeliverToMailboxAndForward_Destination" -Value $mailboxDetails.DeliverToMailboxAndForward -Force
            $user | Add-Member -type NoteProperty -Name "ForwardingSmtpAddress_Destination" -Value $mailboxDetails.ForwardingSmtpAddress -Force
            $user | Add-Member -type NoteProperty -Name "MailboxTotalItemSize_Destination" -Value $mailboxStats.TotalItemSize -Force
            $user | Add-Member -type NoteProperty -Name "MailboxTotalDeletedItemSize_Destination" -Value $mailboxStats.TotalDeletedItemSize -Force
            $user | Add-Member -type NoteProperty -Name "ArchiveTotalItemSize_Destination" -Value $archiveMailboxStats.TotalItemSize -Force
            $user | Add-Member -type NoteProperty -Name "ArchiveTotalDeletedItemSize_Destination" -Value $archiveMailboxStats.TotalDeletedItemSize -Force

            $postMigrationDetails += $user
        }   
    }
    if ($ExportFolderPath) {
		$postMigrationDetails | Export-Csv "$ExportFolderPath\PostMigrationDetails-$Group.csv" -NoTypeInformation -Encoding UTF8
		Write-host "Exported Post Migration Details to $ExportFolderPath\PostMigrationDetails-$Group.csv" -ForegroundColor Cyan
	}
	else {
		try {
			$postMigrationDetails | Export-Csv "$HOME\Desktop\PostMigrationDetails-$Group.csv" -NoTypeInformation -Encoding UTF8
			Write-host "Exported Post Migration Details to $HOME\Desktop\PostMigrationDetails-$Group.csv" -ForegroundColor Cyan
		}
		catch {
			Write-Warning -Message "$($_.Exception)"
			Write-host ""
			$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
			$postMigrationDetails | Export-Csv "$OutputCSVFolderPath\PostMigrationDetails-$Group.csv.csv" -NoTypeInformation -Encoding UTF8
		}
	}
}

# Get Activity Post Migration Details
##Update from existing post migration details
function Get-ActivityReportDetails {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the CSV File Path of Wave Users')] [array] $ImportCSV,
        [Parameter(Mandatory=$True,HelpMessage="Which Tenant to Pull Details from?")] [string] $Tenant,
        [Parameter(Mandatory=$True,HelpMessage="Which Wave Is This?")] [string] $Group,
        [Parameter(Mandatory=$False,HelpMessage="Gather Audit Logs?")] [Switch] $GatherAuditLogs,
        [Parameter(Mandatory=$False,HelpMessage="Check Forwards?")] [Switch] $ForwardCheck,
        [Parameter(Mandatory=$False,HelpMessage="Gather All Mailbox Stats?")] [Switch] $AllMailboxStats,
        [Parameter(Mandatory=$False,HelpMessage="Check Archive?")] [Switch] $ArchiveCheck,
        [Parameter(Mandatory=$False,HelpMessage="Append to Existing File?")] [Switch] $Append,
        [Parameter(Mandatory=$false)][string] $ExportFolderPath
    )

    #Create User Array
    $WaveGroup = Import-csv $ImportCSV
    $postMigrationDetails = @()
    #Set Variables
    $shorterdate = $((Get-Date).ToShortDateString().Replace("/","-"))
    $shortTime = (get-date).ToShortTimeString()

    # Gather User Details
    $progressref = ($WaveGroup).count
    $progresscounter = 0
    foreach ($user in $WaveGroup) {
        #Clear Previous Variables
        $UPN = $null
        $PrimarySMTPAddress = $null
        $azureAuditLogsDetails = @()
        $azureSignInStatus = @()
        $mbxStats = @()
        $mailbox = @()
        $ArchiveStats = @()

        if ($Tenant -eq "Source") {
            $DisplayName = $user.DisplayName_Source
            $UPN = $user.UserPrincipalName_Source
            $PrimarySMTPAddress = $user.PrimarySmtpAddress_Source
        }
        elseif ($Tenant -eq "Destination") {
            $DisplayName = $user.DisplayName_Destination
            $UPN = $user.UserPrincipalName_Destination
            $PrimarySMTPAddress = $user.PrimarySmtpAddress_Destination
        }
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($DisplayName) - $($upn)"

        if ($Append) {
            $outputUserDetail = $user
        }
        else {
            ##Output Stats - Initial
            $outputUserDetail = new-object PSObject
            $outputUserDetail | add-member -type noteproperty -name "Date" -Value $shorterdate -Force
            $outputUserDetail | add-member -type noteproperty -name "Time" -Value $shortTime -Force
            $outputUserDetail | add-member -type noteproperty -name "DisplayName_Source" -Value $user.DisplayName_Source -force
            $outputUserDetail | add-member -type noteproperty -name "UserPrincipalName_Source" -Value $user.UserPrincipalName_Source -force
            $outputUserDetail | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $user.RecipientTypeDetails_Source -force
            $outputUserDetail | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $User.PrimarySMTPAddress_Source -force
            $outputUserDetail | add-member -type noteproperty -name "DisplayName_Destination" -Value $user.DisplayName_Destination -force
            $outputUserDetail | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $user.UserPrincipalName_Destination -force
            $outputUserDetail | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $user.RecipientTypeDetails_Destination -force
            $outputUserDetail | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $User.PrimarySMTPAddress_Destination -force
        }
        #Gather Azure Audit Logs
        if ($GatherAuditLogs) {
            $azureAuditLogsDetails = Get-AzureADAuditSignInLogs -Filter "UserPrincipalName eq '$UPN'" -Top 1 | Select CreatedDateTime, AppDisplayName, ClientAppUsed, Status
            if ($azureAuditLogsDetails.Status.ErrorCode -eq "0") {
                $azureSignInStatus = "Success"
            }
            elseif ($azureAuditLogsDetails -eq $null) {
                $azureSignInStatus = "NoResults"
            }
            else {
                $azureSignInStatus = $azureAuditLogsDetails.Status.FailureReason
            }

            $outputUserDetail | Add-Member -type NoteProperty -Name "MailboxLastActivity_$($Tenant)" -Value $mailboxStats.LastInteractionTime -Force
            $outputUserDetail | Add-Member -type NoteProperty -Name "AzureLastLoginTime_$($Tenant)" -Value $azureAuditLogsDetails.CreatedDateTime -Force
            $outputUserDetail | Add-Member -type NoteProperty -Name "AzureAppDisplayname_$($Tenant)" -Value $azureAuditLogsDetails.AppDisplayName -Force
            $outputUserDetail | Add-Member -type NoteProperty -Name "AzureClientAppUsed_$($Tenant)" -Value $azureAuditLogsDetails.ClientAppUsed -Force
            $outputUserDetail | Add-Member -type NoteProperty -Name "AzureLoginStatus_$($Tenant)" -Value $azureSignInStatus -Force
        }
        #Gather Forward Details
        if ($ForwardCheck) {
            #Gather Mailbox Details
            $mailbox = Get-EXOMailbox -PropertySets delivery -identity $PrimarySMTPAddress -ErrorAction SilentlyContinue
            $mbxStats = Get-EXOMailboxStatistics -PropertySets All -UserPrincipalName $UPN -ErrorAction SilentlyContinue

            $outputUserDetail | add-member -type noteproperty -name "DeliverToMailboxAndForward_$($Tenant)" -Value $mailbox.DeliverToMailboxAndForward -force
            $outputUserDetail | add-member -type noteproperty -name "ForwardingAddress_$($Tenant)" -Value $mailbox.ForwardingAddress -force
            $outputUserDetail | add-member -type noteproperty -name "ForwardingSmtpAddress_$($Tenant)" -Value $mailbox.ForwardingSmtpAddress -force
        }
        #Gather All Mailbox Stats
        if ($AllMailboxStats) {
            $mailbox = Get-EXOMailbox -PropertySets archive,addresslist,delivery,Minimum -identity $PrimarySMTPAddress -ErrorAction SilentlyContinue
            $mbxStats = Get-EXOMailboxStatistics -PropertySets All -UserPrincipalName $UPN -ErrorAction SilentlyContinue

            $outputUserDetail | add-member -type noteproperty -name "FoundPrimarySMTPAddress_$($Tenant)" -Value $mailbox.PrimarySMTPAddress -force
            $outputUserDetail | add-member -type noteproperty -name "FoundRecipientTypeDetails_$($Tenant)" -Value $mailbox.RecipientTypeDetails -force
            $outputUserDetail | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_$($Tenant)" -Value $mailbox.HiddenFromAddressListsEnabled -force
            $outputUserDetail | add-member -type noteproperty -name "DeliverToMailboxAndForward_$($Tenant)" -Value $mailbox.DeliverToMailboxAndForward -force
            $outputUserDetail | add-member -type noteproperty -name "ForwardingAddress_$($Tenant)" -Value $mailbox.ForwardingAddress -force
            $outputUserDetail | add-member -type noteproperty -name "ForwardingSmtpAddress_$($Tenant)" -Value $mailbox.ForwardingSmtpAddress -force
            $outputUserDetail | Add-Member -type NoteProperty -Name "MBXSize_$($Tenant)" -Value $MBXStats.TotalItemSize.Value -force
            $outputUserDetail | Add-Member -Type NoteProperty -name "MBXItemCount_$($Tenant)" -Value $MBXStats.ItemCount -force
            $outputUserDetail | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $mailbox.ArchiveStatus -force 
            $outputUserDetail | add-member -type noteproperty -name "LastLogonTime_$($Tenant)" -Value $mbxStats.LastLogonTime.ToShortDateString() -force -ErrorAction silentlycontinue
            $outputUserDetail | add-member -type noteproperty -name "LastUserActionTime_$($Tenant)" -Value $mbxStats.LastUserActionTime.ToShortDateString() -force -ErrorAction silentlycontinue
        }
        #Gather ArchiveStats
        if ($ArchiveCheck) {
            if ($ArchiveStats = Get-MailboxStatistics -identity $UPN -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                $outputUserDetail | add-member -type noteproperty -name "ArchiveSize_$($Tenant)" -Value $ArchiveStats.TotalItemSize.Value -force
                $outputUserDetail | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $ArchiveStats.ItemCount -force
            }
            else  {
                $outputUserDetail | add-member -type noteproperty -name "ArchiveSize_$($Tenant)" -Value $null -force
                $outputUserDetail | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
            }
        }
        $postMigrationDetails += $outputUserDetail
    }
    if ($ExportFolderPath) {
		$postMigrationDetails | Export-Csv "$ExportFolderPath\PostMigrationDetails-$Group.csv" -NoTypeInformation -Encoding UTF8
		Write-host "Exported Post Migration Details to $ExportFolderPath\PostMigrationDetails-$Group.csv" -ForegroundColor Cyan
	}
	else {
		try {
			$postMigrationDetails | Export-Csv "$HOME\Desktop\PostMigrationDetails-$Group.csv" -NoTypeInformation -Encoding UTF8
			Write-host "Exported Post Migration Details to $HOME\Desktop\PostMigrationDetails-$Group.csv" -ForegroundColor Cyan
		}
		catch {
			Write-Warning -Message "$($_.Exception)"
			Write-host ""
			$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
			$postMigrationDetails | Export-Csv "$OutputCSVFolderPath\PostMigrationDetails-$Group.csv" -NoTypeInformation -Encoding UTF8
		}
	}
}


## Get Post Migration Details - Just Mailbox Stats
$AllMailboxes = Import-EXCEL -WorkSheetName "PostMigrationDetails-ALL-513" -Path

$progressref = ($AllMailboxes).count
$progresscounter = 0
$Tenant = "Destination2"
foreach ($user in $AllMailboxes) {
    #Clear Variables
    $DisplayName = @()
    $UPN = @()
    $PrimarySMTPAddress = @()

    #Set Variables
    $DisplayName = $user.DisplayName_Destination
    $UPN = $user.UserPrincipalName_Destination
    $PrimarySMTPAddress = $user.PrimarySMTPAddress_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($DisplayName) - $($upn)"

    #Add Size Details
    if ($UPN) {
        #Get Mailbox Stats
        $mbxStats = Get-EXOMailboxStatistics -PropertySets All -UserPrincipalName $UPN -ErrorAction SilentlyContinue
        #Add Size Details
        $user | Add-Member -type NoteProperty -Name "MBXSize_$($Tenant)" -Value $MBXStats.TotalItemSize.Value -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_$($Tenant)" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $mailbox.ArchiveStatus -force 

        if ($ArchiveStats = Get-MailboxStatistics -identity $UPN -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
            $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)" -Value $ArchiveStats.TotalItemSize.Value -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $ArchiveStats.ItemCount -force
        }
        else  {
            $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
        }
    }
    else {
        $user | Add-Member -type NoteProperty -Name "MBXSize_$($Tenant)" -Value $null -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $null -force 
        $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
    }

    #Small Break
    start-sleep -Milliseconds 250
    
}


#rollback

function Start-RollbackBatchCutover {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the CSV File Path of OneDrive Users')] [array] $ImportCSV,
        [Parameter(Mandatory=$True,HelpMessage="Which Tenant to Update (Source or Destination?")] [string] $Tenant,
        [Parameter(Mandatory=$false,HelpMessage="Which Wave Is This?")] [string] $Group,
        [Parameter(Mandatory=$false,HelpMessage="Set Forward?")] [switch] $RemoveForward,
        [Parameter(Mandatory=$false,HelpMessage="Restore Access to Mailbox?")] [switch] $RestoreMailAccess,
        [Parameter(Mandatory=$false,HelpMessage="Block Access to OneDrive?")] [switch] $RestoreOneDriveAccess,
        [Parameter(Mandatory=$false,HelpMessage="Disable Teams?")] [switch] $EnableTeams,
        [Parameter(Mandatory=$false,HelpMessage="Do you want to Test?")] [switch] $Test,
        [Parameter(Mandatory=$false,HelpMessage="Hide OldCompany Contacts in NewCompany?")] [switch] $unhideOldCompanyContact,
        [Parameter(Mandatory=$false,HelpMessage="Force Log Out User?")] [switch] $ForceLogOut
    )
    #Create User Array
    $WaveGroup = Import-csv $ImportCSV

    # Gather User Details
    $progressref = ($WaveGroup).count
    $progresscounter = 0
    $allErrors = @()
        
    foreach ($user in $WaveGroup) {
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($user.DisplayName_Source)"
        $SourceUPN = $user.UserPrincipalName_Source
        $SourcePrimarySMTPAddress = $user.PrimarySMTPAddress_Source
        $DesinationPrimarySMTPAddress = $user.PrimarySmtpAddress_Destination

        if ($test) {
            if ($Tenant -eq "Destination") {
                # NewCompany
                Write-Host "Cutting Over User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
                ## Hide OldCompany Contact In NewCompany
                if ($HideOldCompanyContact) {
                    try {
                        Set-MailContact $SourcePrimarySMTPAddress -HiddenFromAddressListsEnabled $true
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
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
            }
            if ($Tenant -eq "Source") {
                # OldCompany
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
                        Set-CASMailbox $SourcePrimarySMTPAddress -MacOutlookEnabled $false -MAPIEnabled $false -OutlookMobileEnabled $false -OWAEnabled $false -ActiveSyncEnabled $false -OWAforDevices $false -ErrorAction Stop
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
                    $SPOUPN = $UPN.replace("@og-example.org","_OldCompany_edu")
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
                    #$SKUID = ($msoluser).licenses.AccountSkuId
                    $SKUID = "EHN:ENTERPRISEPACK"
                
                    ## Get Disabled Array
                    for($i = 0; $i -lt $AllLicenses.Count; $i++) {
                        $serviceStatus =  $AllLicenses[$i].ServiceStatus
                        foreach($service in $serviceStatus)
                        {
                            if($service.ProvisioningStatus -eq "Disabled")
                            {
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
                Write-Host "Done" -foregroundcolor Green 
            } 
        }
        else {
            if ($Tenant -eq "Destination") {
                # NewCompany
                Write-Host "Cutting Over User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
                ## Hide OldCompany Contact In NewCompany
                if ($HideOldCompanyContact) {
                    try {
                        Set-MailContact $SourcePrimarySMTPAddress -HiddenFromAddressListsEnabled $true
                        Write-Host ". " -ForegroundColor Green -NoNewline
                    }
                    catch {
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
            }
            if ($Tenant -eq "Source") {
                # OldCompany
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
                if ($RestoreMailAccess) {
                    Write-Host "Restore Access To Mailbox ..  " -foregroundcolor DarkGray -nonewline
                    Try{ 
                        #Set-CASMailbox $SourcePrimarySMTPAddress -MacOutlookEnabled $false -MAPIEnabled $false -OutlookMobileEnabled $false -OWAEnabled $false -ActiveSyncEnabled $false -OWAforDevices $false -ErrorAction Stop
                        Set-Mailbox -Identity $SourcePrimarySMTPAddress -AccountDisabled:$false
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
                    $SPOUPN = $UPN.replace("@og-example.org","_OldCompany_edu")
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
                    #$SKUID = ($msoluser).licenses.AccountSkuId
                    $SKUID = "EHN:ENTERPRISEPACK"
                
                    ## Get Disabled Array
                    for($i = 0; $i -lt $AllLicenses.Count; $i++) {
                        $serviceStatus =  $AllLicenses[$i].ServiceStatus
                        foreach($service in $serviceStatus)
                        {
                            if($service.ProvisioningStatus -eq "Disabled")
                            {
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
                Write-Host "Done" -foregroundcolor Green 
            }
        }
    }
}

$OldCompanyAddress = "RansomSa@og-example.org"
Set-CASMailbox $OldCompanyAddress -MacOutlookEnabled $True -MAPIEnabled $True -OutlookMobileEnabled $True -OWAEnabled $True -ActiveSyncEnabled $True
Set-Mailbox $OldCompanyAddress -DeliverToMailboxAndForward $false -ForwardingSmtpAddress $null
#enable Teams
$msoluser = Get-MsolUser -UserPrincipalName $OldCompanyAddress
$DisabledArray = @()
$allLicenses = ($msoluser).Licenses
#$SKUID = ($msoluser).licenses.AccountSkuId
$SKUID = "EHN:ENTERPRISEPACK"

## Get Disabled Array
for($i = 0; $i -lt $AllLicenses.Count; $i++) {
    $serviceStatus =  $AllLicenses[$i].ServiceStatus
    foreach($service in $serviceStatus)
    {
        if($service.ProvisioningStatus -eq "Disabled")
        {
            $disabledArray += ($service.ServicePlan).ServiceName
        }
    }
}
#add Teams to DisabledArray
$disabledArray2 = $disabledArray | ?{$_ -ne "Teams1"}
$LicenseOptions = New-MsolLicenseOptions -AccountSkuId $SKUID -DisabledPlans $disabledArray
Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -LicenseOptions $LicenseOptions

Set-MailContact $OldCompanyAddress -HiddenFromAddressListsEnabled $false



## Check If User has Access to Shared Mailbox and who else has access to Shared Mailbox impacted IF Migrated
$WaveGroup = Import-CSV
$fullAccessPerms = Import-CSV

# Gather User Details
$progressref = ($WaveGroup).count
$progresscounter = 0

$WaveSharedMailboxes = @()
foreach ($user in $WaveGroup) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($user.DisplayName)"

    #Grab list of Shared Mailboxes impacted by Wave users
    $SharedMailboxesWaveTemp = @()
    $SharedMailboxesWaveCombined = $Null
    $SharedMailboxesWaveTemp = $fullAccessPerms | ?{$_.User -eq $user.PrimarySMTPAddress}
    $SharedMailboxesWaveCombined = $SharedMailboxesWaveTemp.Identity -join ";"

    #Grab all PermUsers of SharedMailboxes
    $allsharedmailbox = @()
    $allsharedmailboxCombined = $Null
    foreach ($perm in $SharedMailboxesWaveTemp) {
        $allsharedmailbox += $fullAccessPerms | ?{$_.Identity -eq $perm.Identity}
        $allsharedmailboxCombined= $allsharedmailbox.User -join ";"
    }

    #Users Not In Wave
    $nonWaveUser = @()
    $nonWaveUsersCombined = $null
    foreach ($permuser in $allsharedmailbox) {
        if (!($nonWaveUsersCheck = $WaveGroup | ?{$_.PrimarySMTPAddress -eq $permuser.user})) {
            $nonWaveUser += $permuser.user
        }
        $nonWaveUsersCombined = $nonWaveUser -join ";"
    }
    
    #Output
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "Wave" -Value "Wave 2" -Force    
    $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $user.DisplayName -Force
    $currentuser | add-member -type noteproperty -name "WaveUser" -Value $user.PrimarySMTPAddress -Force
    $currentuser | add-member -type noteproperty -name "SharedMailbox" -Value $SharedMailboxesWaveCombined -Force
    $currentuser | add-member -type noteproperty -name "SharedMailboxCount" -Value ($SharedMailboxesWaveTemp | measure).count -Force
    $currentuser | add-member -type noteproperty -name "FullAccessUsers" -Value $allsharedmailboxCombined -Force
    $currentuser | add-member -type noteproperty -name "FullAccessUsersCount" -Value ($allsharedmailbox | measure).count -Force
    $currentuser | add-member -type noteproperty -name "UsersNotInCurrentWave" -Value $nonWaveUsersCombined -Force
    $currentuser | add-member -type noteproperty -name "UsersNotInCurrentWaveCount" -Value ($nonWaveUser | measure).count -Force
    $WaveSharedMailboxes += $currentuser
}


# SHORT Match Mailboxes and add to same spreadsheet. Check based on Campus Key and DisplayName
$WaveGroup = Import-Csv
$progressref = ($WaveGroup).count
$progresscounter = 0

foreach ($user in $WaveGroup) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.FoundDisplayName)"
    $NewUPN = $user.CampusKey + "@example.org"
    $addressSplit = $user.PrimarySmtpAddress -split "@"
    $ehnPrimarySMTPAddress = $addressSplit[0] + "-old@example.org"
    $newPrimarySMTPAddress = $user.PrimarySmtpAddress + "@example.org"

    #NEW UPN Check
    if ($msoluser = Get-Msoluser -UserPrincipalName $NewUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with CampusKey" -ForegroundColor Green
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
    
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "CampusKeyMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
    }
    #NEW PrimarySMTPAddress Check (EHN added)
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $ehnPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with EHN added" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "EHNMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force

    }
    #NEW PrimarySMTPAddress Check
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $newPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "NEWSMTPAddressCheck" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
    }
    #UPN Check based on DisplayName
    elseif ($msoluser = Get-Msoluser -SearchString $user.FoundDisplayName -ErrorAction SilentlyContinue) {
        if ($msoluser.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName)"
            $mailbox = Get-Mailbox $user.DisplayName -ErrorAction SilentlyContinue
            $msoluser = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        }
        Write-Host "$($msoluser.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
        $mailbox = Get-Mailbox $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
            
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value "DisplayNameMatch" -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluser.BlockCredential -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
    }
    
    else {
        Write-Host "  Unable to find user for $($user.DisplayName)" -ForegroundColor Red
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "BlockCredential" -Value $null -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -force
    }
}

# SHORT Match Mailboxes and Pull from IDentity Team data
$teamsUsers = Import-Csv
$userMailboxes = import-csv
$progressref = ($teamsUsers).count
$progresscounter = 0

foreach ($user in $teamsUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.UPN)"

    if ($userCheck = $waveUsers | ? {$_.UserPrincipalName -eq $user.UPN}) {
        $userCheck | Export-Csv -NoTypeInformation -Encoding UTF8 C:\Users\amedrano\Desktop\Batches\TeamsUsers.csv -Append
    }
}

#compare wave users to usermailboxes

$userMailboxes = import-csv
$waveUsers = import-csv
$progressref = ($userMailboxes).count
$progresscounter = 0
$NonWaveUsers = @()
foreach ($usermailbox in $userMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($usermailbox.UPN)"

    if (!($userCheck = $waveUsers | ? {$_.UserPrincipalName -eq $usermailbox.UserPrincipalName})) {
        $currentobject = new-object PSObject
        $currentobject | add-member -type noteproperty -name "DisplayName" -Value $usermailbox.DisplayName -Force
        $currentobject | add-member -type noteproperty -name "UserPrincipalName" -Value $usermailbox.UserPrincipalName -Force
        $currentobject | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $usermailbox.PrimarySMTPAddress -Force
        $currentobject | add-member -type noteproperty -name "RecipientTypeDetails" -Value $usermailbox.RecipientTypeDetails -Force
        $currentobject | add-member -type noteproperty -name "CampusKey" -Value $usermailbox.CustomAttribute7 -Force
        $currentobject | add-member -type noteproperty -name "WaveUser" -Value $false -Force
    }
    $NonWaveUsers += $currentobject
    $currentobject | Export-Csv -NoTypeInformation -Encoding UTF8 C:\Users\amedrano\Desktop\Batches\NonWaveUsers.csv -Append
}
#$NonWaveUsers | Export-Csv -NoTypeInformation -Encoding UTF8 C:\Users\amedrano\Desktop\Batches\NonWaveUsers.csv


# (one off) gather wave 3 user details - not forwarding
$allMatchedMailboxes = import-csv
$notForwardingusers = import-csv
$progressref = ($notForwardingusers).count
$progresscounter = 0
$MatchedMailboxesNoForward = @()
$notFoundMailboxes = @()
foreach ($usermailbox in $notForwardingusers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($usermailbox.PrimarySMTPAddress)"

    if ($userCheck = $allMatchedMailboxes | ? {$_.PrimarySMTPAddress_Source -eq $usermailbox.PrimarySMTPAddress}) {
        Write-Host "User $($userCheck.UserPrincipalName_Source) Found in All Mailboxes" -foregroundcolor green
        $MatchedMailboxesNoForward += $userCheck
    }
    else {
        Write-Host "User $($usermailbox.PrimarySMTPAddress) Not Found in All Mailboxes" -foregroundcolor Yellow
        $notFoundMailboxes  += $usermailbox
    }
}

#Recheck Teams Wave
$WaveGroup = import-csv
$progressref = ($WaveGroup).count
$progresscounter = 0
foreach ($object in $WaveGroup) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($object.UPN)"

    if ($msolUserCheck = Get-MsolUser -userprincipalname $object.UPN -ea silentlycontinue){
        $mailboxCheck = Get-Mailbox $msolUserCheck.userPrincipalName -EA SilentlyContinue
        $object | add-member -type noteproperty -name "FoundInOldCompany" -Value $true -force
        $object | add-member -type noteproperty -name "FoundDisplayName" -Value $msolUserCheck.DisplayName -force
        $object | add-member -type noteproperty -name "FoundUserPrincipalName" -Value $msolUserCheck.UserPrincipalName -force
        $object | add-member -type noteproperty -name "IsLicensed" -Value $msolUserCheck.IsLicensed -force
        $object | add-member -type noteproperty -name "Department" -Value $msolUserCheck.Department -force
        $object | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mailboxCheck.RecipientTypeDetails -force
        $object | add-member -type noteproperty -name "FoundSMTPAddress" -Value $mailboxCheck.primarysmtpaddress -force
        $object | add-member -type noteproperty -name "CustomAttribute7" -Value $mailboxCheck.CustomAttribute7 -force       
    } 
    else {
        $object | add-member -type noteproperty -name "FoundInOldCompany" -Value $false -force
        $object | add-member -type noteproperty -name "FoundDisplayName" -Value $null -force
        $object | add-member -type noteproperty -name "FoundUserPrincipalName" -Value $null -force
        $object | add-member -type noteproperty -name "IsLicensed" -Value $null -force
        $object | add-member -type noteproperty -name "Department" -Value $null -force
        $object | add-member -type noteproperty -name "RecipientTypeDetails" -Value $null -force
        $object | add-member -type noteproperty -name "FoundSMTPAddress" -Value $null -force
        $object | add-member -type noteproperty -name "CustomAttribute7" -Value $null -force
    }

}


#Recheck Teams Wave - check for campus key in IDentity Teams list
$inactivemailboxes = import-csv
$progressref = ($inactivemailboxes).count
$progresscounter = 0
foreach ($object in $inactivemailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User $($object.PrimarySmtpAddress)"

    $MBXStats = Get-MailboxStatistics -IncludeSoftDeletedRecipients -Identity $object.PrimarySmtpAddress
    $object | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize -force
    $object | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount -force
}


#add source SharePoint details
$matchedMailboxes = import-csv
#ProgressBar
$progressref = ($matchedMailboxes).count
$progresscounter = 0

foreach ($user in $matchedMailboxes){
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering SharePoint Details for $($user.DisplayName_Source)"
    $UPN = $user.UserPrincipalName_Source

    $SPOSITE = Get-SPOSITE -IncludePersonalSite $true -filter "Owner -eq $UPN -and URL -like '*-my.sharepoint*'" -ErrorAction SilentlyContinue
    $user | add-member -type noteproperty -name "OneDriveURL_Source" -Value $SPOSITE.URL -Force
    $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)" -Value $SPOSITE.StorageUsageCurrent -Force
    $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)" -Value ($SPOSITE.StorageUsageCurrent/1024) -Force
}

#add Destination SharePoint URLs
$matchedMailboxes = import-csv
#ProgressBar
$progressref = ($matchedMailboxes).count
$progresscounter = 0

foreach ($user in $matchedMailboxes){
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering SharePoint Details for $($user.DisplayName_Source)"
    $UPN = $user.UserPrincipalName_Source

    $SPOSITE = Get-SPOSITE -IncludePersonalSite $true -filter "Owner -eq $UPN -and URL -like '*-my.sharepoint*'" -ErrorAction SilentlyContinue
    $user | add-member -type noteproperty -name "OneDriveURL_Destination" -Value $SPOSITE.URL -Force
    $user | add-member -type noteproperty -name "OneDriveCurrentStorage(MB)_Destination" -Value $SPOSITE.StorageUsageCurrent -Force
    $user | add-member -type noteproperty -name "OneDriveCurrentStorage(GB)" -Value ($SPOSITE.StorageUsageCurrent/1024) -Force
}

#Check for Mailboxes in Migration Job Wave
$Wave5UserMailboxes = Import-Excel -WorksheetName "Wave 5 Users - 4368" -path "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\All-MatchedMailboxes_3-25_2022.xlsx"
$NonUserMailboxes = Import-Excel -WorksheetName "Wave 5 Non Users - 658" -path "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\All-MatchedMailboxes_3-25_2022.xlsx"
$GroupMailboxes = Import-CSV "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\MatchedGroups_EHN-TJU.csv"
$GroupMailboxes = $GroupMailboxes | ?{$_.IsMailboxConfigured_Source -eq $true}

$MigrationJobDetails  = Import-CSV "C:\Users\amedrano\Desktop\MatchedItems\All-MatchedMailboxes_TMP 3-25.csv"
$foundMigJob = @()
$notfoundMigJob = @()

$progressref = $NonUserMailboxes.count
$progresscounter = 0
foreach ($user in $NonUserMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details UserMailbox $($user.PrimarySmtpAddress)"

    if ($matchedUser = $MigrationJobDetails | ? {$_.SourceEmailAddress -eq $user.PrimarySmtpAddress}) {
        $foundMigJob += $matchedUser
    }
    else {
        $notfoundMigJob += $user
    }
}

#Update Wave User's Details
$updatedMatchedMailboxDetails = Import-Excel -WorksheetName "Master List2" -path 'c:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$updatedMatchedMailboxDetails = import-csv
$oldMatchedMailboxDetails  = Import-Excel -path 'c:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$progressref = $updatedMatchedMailboxDetails.count
$progresscounter = 0
foreach ($user in $updatedMatchedMailboxDetails) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details UserMailbox $($user.PrimarySmtpAddress_Source)"

    if ($matchedUser = $oldMatchedMailboxDetails | ? {$_.PrimarySmtpAddress_Source -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "CustomAttribute7_Source" -Value $matchedUser.CustomAttribute7_Source -force
        $user | add-member -type noteproperty -name "CampusKey" -Value $matchedUser.CampusKey -force
        $user | add-member -type noteproperty -name "ServiceAccounts" -Value $matchedUser.ServiceAccounts -force
        $user | add-member -type noteproperty -name "DisabledAccounts" -Value $matchedUser.DisabledAccounts -force
        $user | add-member -type noteproperty -name "Migrate" -Value $matchedUser.Migrate -force
        $user | add-member -type noteproperty -name "Wave1" -Value $matchedUser.Wave1 -force
        $user | add-member -type noteproperty -name "Wave2" -Value $matchedUser.Wave2 -force
        $user | add-member -type noteproperty -name "Wave3" -Value $matchedUser.Wave3 -force
        $user | add-member -type noteproperty -name "Wave5" -Value $matchedUser.Wave5 -force
        $user | add-member -type noteproperty -name "TeamsUser" -Value $matchedUser.TeamsUser -force
        $user | add-member -type noteproperty -name "MDM_ImpactedUsers" -Value $matchedUser.MDM_ImpactedUsers -force
        $user | add-member -type noteproperty -name "VIPsandAdmins" -Value $matchedUser.VIPsandAdmins -force
    }
    else {
        $user | add-member -type noteproperty -name "CustomAttribute7_Source" -Value $null -force
        $user | add-member -type noteproperty -name "CampusKey" -Value $null -force
        $user | add-member -type noteproperty -name "Migrate" -Value $null  -force
        $user | add-member -type noteproperty -name "Wave1" -Value $null  -force
        $user | add-member -type noteproperty -name "Wave2" -Value $null  -force
        $user | add-member -type noteproperty -name "Wave3" -Value $null  -force
        $user | add-member -type noteproperty -name "Wave5" -Value $null  -force
        $user | add-member -type noteproperty -name "TeamsUser" -Value $null  -force
        $user | add-member -type noteproperty -name "MDM_ImpactedUsers" -Value $null  -force
        $user | add-member -type noteproperty -name "VIPsandAdmins" -Value $null  -force
    }
}

#Update Wave User's Details
$updatedMatchedMailboxDetails = Import-Excel -WorksheetName "Master List2" -path 'c:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$updatedMatchedMailboxDetails = import-csv
$oldMatchedMailboxDetails  = Import-Excel -path 'c:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$progressref = $updatedMatchedMailboxDetails.count
$progresscounter = 0
foreach ($user in $updatedMatchedMailboxDetails) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details UserMailbox $($user.PrimarySmtpAddress_Source)"

    if ($matchedUser = $oldMatchedMailboxDetails | ? {$_.PrimarySmtpAddress_Source -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "CustomAttribute7_Source" -Value $matchedUser.CustomAttribute7_Source -force
    }
    else {
        $user | add-member -type noteproperty -name "CustomAttribute7_Source" -Value $null  -force
    }
}

#Update Post Migration Details - Include Wave Details
$postmigrationdetails = Import-Excel -WorksheetName "Master List2" -path 'c:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$updatedMatchedMailboxDetails = import-csv
$allmailboxdetails  = Import-Excel -path 'c:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\AllMatched_Mailboxes-3-21-2022.xlsx'
$progressref = $postmigrationdetails.count
$progresscounter = 0
foreach ($user in $postmigrationdetails) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details UserMailbox $($user.PrimarySmtpAddress_Source)"

    if ($matchedUser = $allmailboxdetails | ? {$_.UserPrincipalName_Source -eq $user.UserPrincipalName_Source}) {
        #$user | add-member -type noteproperty -name "CustomAttribute7_Source" -Value $matchedUser.CustomAttribute7_Source -force
        #$user | add-member -type noteproperty -name "CampusKey" -Value $matchedUser.CampusKey -force
        $user | add-member -type noteproperty -name "ServiceAccounts" -Value $matchedUser.ServiceAccounts -force
        $user | add-member -type noteproperty -name "DisabledAccounts" -Value $matchedUser.DisabledAccounts -force
        $user | add-member -type noteproperty -name "Migrate" -Value $matchedUser.Migrate -force
        $user | add-member -type noteproperty -name "Wave1" -Value $matchedUser.Wave1 -force
        $user | add-member -type noteproperty -name "Wave2" -Value $matchedUser.Wave2 -force
        $user | add-member -type noteproperty -name "Wave3" -Value $matchedUser.Wave3 -force
        $user | add-member -type noteproperty -name "Wave5" -Value $matchedUser.Wave5 -force
        $user | add-member -type noteproperty -name "TeamsUser" -Value $matchedUser.TeamsUser -force
        $user | add-member -type noteproperty -name "MDM_ImpactedUsers" -Value $matchedUser.MDM_ImpactedUsers -force
        $user | add-member -type noteproperty -name "VIPsandAdmins" -Value $matchedUser.VIPsandAdmins -force
    }
    else {
        #$user | add-member -type noteproperty -name "CustomAttribute7_Source" -Value $null -force
        #$user | add-member -type noteproperty -name "CampusKey" -Value $null -force
        $user | add-member -type noteproperty -name "Migrate" -Value $null  -force
        $user | add-member -type noteproperty -name "Wave1" -Value $null  -force
        $user | add-member -type noteproperty -name "Wave2" -Value $null  -force
        $user | add-member -type noteproperty -name "Wave3" -Value $null  -force
        $user | add-member -type noteproperty -name "Wave5" -Value $null  -force
        $user | add-member -type noteproperty -name "TeamsUser" -Value $null  -force
        $user | add-member -type noteproperty -name "MDM_ImpactedUsers" -Value $null  -force
        $user | add-member -type noteproperty -name "VIPsandAdmins" -Value $null  -force
    }
}

#Update SharePoint Site Details
$OldCompanySPOSites = Import-Excel -WorksheetName "EHN - SPO Sites" -path "C:\Users\amedrano\Desktop\SharePoint\EHN-TJU SharePoint Summary.xlsx"
$TJUSPOSites  = Import-Excel -WorksheetName "TJU - SPO Sites" -path "C:\Users\amedrano\Desktop\SharePoint\EHN-TJU SharePoint Summary.xlsx"
$progressref = $TJUSPOSites.count
$progresscounter = 0
foreach ($object in $TJUSPOSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details  $($object."Site URL")"

    if ($SiteDetails = Get-SPOSite -Identity $object."Site URL") {
        $object | add-member -type noteproperty -name "Title" -Value $SiteDetails.Title -force
    }
    else {
        $object | add-member -type noteproperty -name "Title" -Value $null -force
    }
}

$TJUSPOSites | Export-Excel -WorksheetName "TJU - SPO Sites" -path "C:\Users\amedrano\Desktop\SharePoint\EHN-TJU SharePoint Summary.xlsx"

## 
function New-NewCompanyResourceMailboxes {
    param (
        [Parameter(Mandatory=$True,HelpMessage="Which Type to Create?")] [string] $MailboxType,
        [Parameter(Mandatory=$True,HelpMessage="What is the CSV File Path")] [string] $ImportExcel,
        [Parameter(Mandatory=$True,HelpMessage="What is the CSV File Path")] [string] $WorkSheet
    )
    $ImportCSVUsers = Import-Excel -WorksheetName $WorkSheet -path $ImportExcel

    $resources = $ImportCSVUsers | ?{$_.RecipientTypeDetails_Source -eq "RoomMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
    $sharedMailboxes = $ImportCSVUsers | ?{$_.RecipientTypeDetails_Source -eq "SharedMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
    $EquipmentMailboxes = $ImportCSVUsers | ?{$_.RecipientTypeDetails_Source -eq "EquipmentMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
    $createdObject = @()
    $alreadyExists = @()
    $failedToCreate = @()

    if ($MailboxType -eq "Resource") {
        $progressref = ($resources).count
        $progresscounter = 0     
        foreach ($mailbox in $resources) {
            #Set Variables
            $addressSplit = $mailbox.PrimarySmtpAddress_Source -split "@"
            $destinationEmail = $addressSplit[0] + "@example.org"
            $destinationDisplayName = $mailbox.DisplayName_Source
            $EHNDisplayName = $mailbox.DisplayName_Source + " - EHN"
            $EHNAddress = $addressSplit[0] + "-old@example.org"
            
            #Progress Bar - Resources
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Resource $($destinationDisplayName)"

            Write-Host "Checking for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
            if (!($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue)) {
                try {
                    $remoteMailboxCreate = New-Mailbox -Room -DisplayName $destinationDisplayName -PrimarySmtpAddress $destinationEmail -Name $destinationDisplayName -ErrorAction Stop
                    Write-Host "Created Successfully." -ForegroundColor Green
                    $createdObject += $mailbox
                }
                catch {
                    try {
                        $remoteMailboxCreate = New-Mailbox -Room -DisplayName $EHNDisplayName -PrimarySmtpAddress $EHNAddress -Name $EHNDisplayName -ErrorAction Stop
                        Write-Host "Created EHN Address Successfully." -ForegroundColor Green
                        $createdObject += $mailbox
                    }
                    catch {
                        Write-Host ($_.Exception)
                        Write-Host "Failed to Create" -ForegroundColor Red
                        $failedToCreate += $mailbox
                    }
                }  
            }
            else {
                Write-Host "Already Exists" -ForegroundColor Yellow
                $alreadyExists += $mailbox
                try {
                    $remoteMailboxCreate = New-Mailbox -Room -DisplayName $EHNDisplayName -PrimarySmtpAddress $EHNAddress -Name $EHNDisplayName -ErrorAction Stop
                    Write-Host "Created EHN Address Successfully." -ForegroundColor Green
                    $createdObject += $mailbox
                }
                catch {
                    Write-Host ($_.Exception)
                    Write-Host "Failed to Create" -ForegroundColor Red
                    $failedToCreate += $mailbox
                }
            }
        }
    }
    if ($MailboxType -eq "Shared") {
        $progressref = ($sharedMailboxes).count
        $progresscounter = 0    
        foreach ($mailbox in $sharedMailboxes) {
            #Set Variables
            $addressSplit = $mailbox.PrimarySmtpAddress_Source -split "@"
            $destinationEmail = $addressSplit[0] + "@example.org"
            $destinationDisplayName = $mailbox.DisplayName_Source
            $EHNDisplayName = $mailbox.DisplayName_Source + " - EHN"
            $EHNAddress = $addressSplit[0] + "-old@example.org"
            
            #Progress Bar - Resources
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Resource $($destinationDisplayName)"

            Write-Host "Checking for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
            if (!($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue)) {
                try {
                    $remoteMailboxCreate = New-Mailbox -Shared -DisplayName $destinationDisplayName -PrimarySmtpAddress $destinationEmail -Name $destinationDisplayName -ErrorAction Stop
                    Write-Host "Created  Successfully." -ForegroundColor Green
                    $createdObject += $mailbox
                }
                catch {
                    try {
                        $remoteMailboxCreate = New-Mailbox -Shared -DisplayName $EHNDisplayName -PrimarySmtpAddress $EHNAddress -Name $EHNDisplayName -ErrorAction Stop
                        Write-Host "Created EHN Address Successfully." -ForegroundColor Green
                        $createdObject += $mailbox
                    }
                    catch {
                        Write-Host ($_.Exception)
                        Write-Host "Failed to Create" -ForegroundColor Red
                        $failedToCreate += $mailbox
                    }
                }  
            }
            else {
                Write-Host "Already Exists" -ForegroundColor Yellow
                $alreadyExists += $mailbox
                try {
                    $remoteMailboxCreate = New-Mailbox -Shared -DisplayName $EHNDisplayName -PrimarySmtpAddress $EHNAddress -Name $EHNDisplayName -ErrorAction Stop
                    Write-Host "Created EHN Address Successfully." -ForegroundColor Green
                    $createdObject += $mailbox
                }
                catch {
                    Write-Host ($_.Exception)
                    Write-Host "Failed to Create" -ForegroundColor Red
                    $failedToCreate += $mailbox
                }
            }
        }
    }
    if ($MailboxType -eq "Equipment") {
        $progressref = ($EquipmentMailboxes).count
        $progresscounter = 0    
        foreach ($mailbox in $EquipmentMailboxes) {
            #Set Variables
            $addressSplit = $mailbox.PrimarySmtpAddress_Source -split "@"
            $destinationEmail = $addressSplit[0] + "@example.org"
            $destinationDisplayName = $mailbox.DisplayName_Source
            $EHNDisplayName = $mailbox.DisplayName_Source + " - EHN"
            $EHNAddress = $addressSplit[0] + "-old@example.org"
            
            #Progress Bar - Resources
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Resource $($destinationDisplayName)"

            Write-Host "Checking for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
            if (!($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue)) {
                try {
                    $remoteMailboxCreate = New-Mailbox -Equipment -DisplayName $destinationDisplayName -PrimarySmtpAddress $destinationEmail -Name $destinationDisplayName -ErrorAction Stop
                    Write-Host "Created Successfully." -ForegroundColor Green
                    $createdObject += $mailbox
                }
                catch {
                    try {
                        $remoteMailboxCreate = New-Mailbox -Equipment -DisplayName $EHNDisplayName -PrimarySmtpAddress $EHNAddress -Name $EHNDisplayName -ErrorAction Stop
                        Write-Host "Created EHN Address Successfully." -ForegroundColor Green
                        $createdObject += $mailbox
                    }
                    catch {
                        Write-Host ($_.Exception)
                        Write-Host "Failed to Create" -ForegroundColor Red
                        $failedToCreate += $mailbox
                    }
                }  
            }
            else {
                Write-Host "Already Exists" -ForegroundColor Yellow
                $alreadyExists += $mailbox
                try {
                    $remoteMailboxCreate = New-Mailbox -Equipment -DisplayName $EHNDisplayName -PrimarySmtpAddress $EHNAddress -Name $EHNDisplayName -ErrorAction Stop
                    Write-Host "Created EHN Address Successfully." -ForegroundColor Green
                    $createdObject += $mailbox
                }
                catch {
                    Write-Host ($_.Exception)
                    Write-Host "Failed to Create" -ForegroundColor Red
                    $failedToCreate += $mailbox
                }
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
        Write-Host $failedToCreate.count "Failed to Create Resources" -ForegroundColor Red
        $failedToCreate | out-gridview
    }
    if ($MailboxType -eq "Shared") {
        Write-Host $EnabledOjbect.count "Enabled Shared Mailbox" -ForegroundColor Green
        Write-Host $createdObject.count "Created Shared Mailbox" -ForegroundColor Cyan
        Write-Host $alreadyExists.count "Already Existing Resources" -ForegroundColor Gray
        Write-Host $failedToCreate.count "Failed to Create Shared Mailbox" -ForegroundColor Red
        $failedToCreate | out-gridview
    }
    if ($MailboxType -eq "Equipment") {
        Write-Host $EnabledOjbect.count "Enabled Equipment Mailbox" -ForegroundColor Green
        Write-Host $createdObject.count "Created Equipment Mailbox" -ForegroundColor Cyan
        Write-Host $alreadyExists.count "Already Existing Resources" -ForegroundColor Gray
        Write-Host $failedToCreate.count "Failed to Create Equipment Mailbox" -ForegroundColor Red
        $failedToCreate | out-gridview
    }
}

function Remove-NewCompanyResourceMailboxes {
    param (
        [Parameter(Mandatory=$True,HelpMessage="Which Type to Create?")] [string] $MailboxType,
        [Parameter(Mandatory=$True,HelpMessage="What is the CSV File Path")] [string] $ImportExcel,
        [Parameter(Mandatory=$True,HelpMessage="What is the CSV File Path")] [string] $WorkSheet
    )
    $ImportCSVUsers = Import-Excel -WorksheetName $WorkSheet -path $ImportExcel

    $resources = $ImportCSVUsers | ?{$_.RecipientTypeDetails_Source -eq "RoomMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
    $sharedMailboxes = $ImportCSVUsers | ?{$_.RecipientTypeDetails_Source -eq "SharedMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
    $EquipmentMailboxes = $ImportCSVUsers | ?{$_.RecipientTypeDetails_Source -eq "EquipmentMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
    $removedObject = @()
    $failedToRemove = @()
    $NotFound = @()

    if ($MailboxType -eq "Shared") {
        $progressref = ($sharedMailboxes).count
        $progresscounter = 0     
        foreach ($mailbox in $sharedMailboxes) {
            #Set Variables
            $addressSplit = $mailbox.PrimarySmtpAddress_Source -split "@"
            $destinationEmail = $addressSplit[0] + "-old@example.org"
                        
            #Progress Bar - Resources
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Resource $($destinationEmail)"

            Write-Host "Checking for $($destinationEmail) ... " -NoNewline -ForegroundColor Cyan
            if (($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue)) {
                try {
                    Remove-Mailbox $recipientCheck.PrimarySMTPAddress -ErrorAction Stop -confirm:$false
                    Write-Host "Removed Successfully." -ForegroundColor Green
                    $removedObject += $mailbox
                }
                catch {
                    Write-Host ($_.Exception) -NoNewline
                    Write-Host "Failed to Remove" -ForegroundColor Red
                    $failedToRemove += $mailbox
                }  
            }
            else {
                Write-Host "NotFound Exists" -ForegroundColor Yellow
                $NotFound += $mailbox
            }
        }
    }

    Write-Host ""
    Write-Host "Results!" -foregroundcolor Cyan

    #Output
    Write-Host $removedObject.count "Removed Resources" -ForegroundColor Cyan
    Write-Host $failedToRemove.count "Failed to Remove Resources" -ForegroundColor Red
    $failedToRemove | out-gridview
}

$progressref = ($newSharedMailboxes).count
$progresscounter = 0    
foreach ($mailbox in $newSharedMailboxes) {
    #Set Variables
    $addressSplit = $mailbox.PrimarySmtpAddress -split "@"
    $destinationEmail = $addressSplit[0] + "@example.org"
    $destinationDisplayName = $mailbox.DisplayName
    $EHNDisplayName = $mailbox.DisplayName + " - EHN"
    $EHNAddress = $addressSplit[0] + "-old@example.org"
    
    #Progress Bar - Resources
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Resource $($destinationDisplayName)"

    Write-Host "Checking for $($destinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    if (!($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue)) {
        try {
            $remoteMailboxCreate = New-Mailbox -Shared -DisplayName $destinationDisplayName -PrimarySmtpAddress $destinationEmail -Name $destinationDisplayName -ErrorAction Stop
            Write-Host "Created  Successfully." -ForegroundColor Green
            $createdObject += $mailbox
        }
        catch {
            try {
                $remoteMailboxCreate = New-Mailbox -Shared -DisplayName $EHNDisplayName -PrimarySmtpAddress $EHNAddress -Name $EHNDisplayName -ErrorAction Stop
                Write-Host "Created EHN Address Successfully." -ForegroundColor Green
                $createdObject += $mailbox
            }
            catch {
                Write-Host ($_.Exception)
                Write-Host "Failed to Create" -ForegroundColor Red
                $failedToCreate += $mailbox
            }
        }  
    }
    else {
        Write-Host "Already Exists" -ForegroundColor Yellow
        $alreadyExists += $mailbox
        try {
            $remoteMailboxCreate = New-Mailbox -Shared -DisplayName $EHNDisplayName -PrimarySmtpAddress $EHNAddress -Name $EHNDisplayName -ErrorAction Stop
            Write-Host "Created EHN Address Successfully." -ForegroundColor Green
            $createdObject += $mailbox
        }
        catch {
            Write-Host ($_.Exception)
            Write-Host "Failed to Create" -ForegroundColor Red
            $failedToCreate += $mailbox
        }
    }
}

#Match Remaining Mailboxes and add to same spreadsheet. Check based on Campus Key, CustomAttribute7 and DisplayName - CONDENSED
$remainingMailboxes = Import-Excel -WorksheetName "Master List2" -Path C:\Users\amedrano\Desktop\MatchedItems\AllMatched_Mailboxes-3-21-2022.xlsx
$resources = $allmatchedMailboxes | ?{$_.RecipientTypeDetails_Source -eq "RoomMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
$sharedMailboxes = $allmatchedMailboxes | ?{$_.RecipientTypeDetails_Source -eq "SharedMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
$EquipmentMailboxes = $allmatchedMailboxes | ?{$_.RecipientTypeDetails_Source -eq "EquipmentMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}

$progressref = ($remainingMailboxes).count
$progresscounter = 0
foreach ($user in $remainingMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName_Source)"

    #Set Variables
    $addressSplit = $user.PrimarySmtpAddress_Source -split "@"
    $destinationEmail = $addressSplit[0] + "@example.org"
    $destinationDisplayName = $user.DisplayName_Source
    $EHNDisplayName = $user.DisplayName_Source + " - EHN"
    $EHNAddress = $addressSplit[0] + "-old@example.org"

    #NEW PrimarySMTPAddress Check (EHN added)
    if ($msoluser = Get-Msoluser -searchstring $EHNDisplayName -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with EHN added" -ForegroundColor Yellow
    }
    #NEW PrimarySMTPAddress Check
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $EHNAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
    }
    #UPN Check based on DisplayName
    elseif ($msoluser = Get-Msoluser -SearchString $user.DisplayName_Source -ErrorAction SilentlyContinue) {
        if ($msoluser.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName_Source)"
        }
        Write-Host "$($msoluser.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
    }

    if ($msoluser) {
        foreach ($user in $msoluser) {
            $mailbox = Get-Mailbox $user.DisplayName -ErrorAction SilentlyContinue
            $msoluserCheck = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
            $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses

            #Output Stats
            $user | add-member -type noteproperty -name "ExistsInDestination" -Value "True" -force
            $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluserCheck.DisplayName -force
            $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluserCheck.userprincipalname -force
            $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluserCheck.IsLicensed -force
            $user | add-member -type noteproperty -name "Licenses" -Value ($msoluserCheck.Licenses.AccountSkuID -join ";") -force
            $user | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluserCheck.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
            $user | add-member -type noteproperty -name "BlockCredential" -Value $msoluserCheck.BlockCredential -force
            $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
            $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
            $user | add-member -type noteproperty -name "Alias_Destination" -Value $mailbox.alias -force
            $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
            $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
            $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
            $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
            $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
            $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize.Value -force
            $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
            $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $mailbox.ArchiveStatus -force
            $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime.ToShortDateString() -force
            $user | add-member -type noteproperty -name "LastUserActionTime_Destination" -Value $mbxStats.LastUserActionTime.ToShortDateString() -force
            
            if ($ArchiveStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
                    $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
                    $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
            }
            else  {
                $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
            }
        }
    }
    else {
        Write-Host "  Unable to find user for $($user.DisplayName_Source)" -ForegroundColor Red

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
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
        $user | add-member -type noteproperty -name "LastUserActionTime_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
    }
}

#Check Recipients
$allmatchedMailboxes = Import-Excel -WorksheetName "Master List2" -Path C:\Users\amedrano\Desktop\MatchedItems\AllMatched_Mailboxes-3-21-2022.xlsx
$resources = $allmatchedMailboxes | ?{$_.RecipientTypeDetails_Source -eq "RoomMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
$sharedMailboxes = $allmatchedMailboxes | ?{$_.RecipientTypeDetails_Source -eq "SharedMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}
$EquipmentMailboxes = $allmatchedMailboxes | ?{$_.RecipientTypeDetails_Source -eq "EquipmentMailbox" -and $_.Migrate -ne "No" -and $_.ExistsInDestination -eq $false}

$foundRecipient = @()
$notfoundRecipient = @()
$progressref = ($newSharedMailboxes).count
$progresscounter = 0
foreach ($user in $newSharedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"

    #Set Variables
    $addressSplit = $user.PrimarySmtpAddress-split "@"
    $destinationEmail = $addressSplit[0] + "@example.org"
    $destinationDisplayName = $user.DisplayName
    $EHNDisplayName = $user.DisplayName + " - EHN"
    $EHNAddress = $addressSplit[0] + "-old@example.org"

    #NEW PrimarySMTPAddress Check (EHN added)
    if ($recipientCheck = Get-Recipient $EHNDisplayName -ErrorAction SilentlyContinue) {
        Write-Host "$($recipientCheck.DisplayName) User found with EHN added" -ForegroundColor Yellow
    }
    #NEW PrimarySMTPAddress Check
    elseif ($recipientCheck = Get-Recipient $EHNAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($recipientCheck.DisplayName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
    }
    else {
        Write-Host "$($destinationEmail) Not Found found" -ForegroundColor Red
    }
    if ($recipientCheck) {
        #Output Stats
        $currentuser = new-object PSObject
        $currentuser | add-member -type noteproperty -name "DisplayName_Destination" -Value $recipientCheck.DisplayName -force
        $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipientCheck.RecipientTypeDetails -force
        $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipientCheck.PrimarySmtpAddress -force
        $currentuser | add-member -type noteproperty -name "Name_Destination" -Value $recipientCheck.Name -force
        $foundRecipient += $currentuser
    }
    else {
        $notfoundRecipient += $user
    }
}

# update Mail contact
foreach ($recipient in $foundRecipient) {
    $tempName = $recipient.Name_Destination + " temp"
    Set-MailContact -identity $recipient.Name_Destination -Name $tempName
}

#Match Migration Failures to User Details
foreach ($user in $failedUsers){
    if ($matchedUser = $allmatchedMailboxes | ?{$_.PrimarySmtpAddress_Source -eq $user.SourceEmailAddress}) {
        $user | add-member -type noteproperty -name "DisplayName_Source" -Value $matchedUser.DisplayName_Source -force
        $user | add-member -type noteproperty -name "IsLicensed_Source" -Value $matchedUser.IsLicensed_Source -force
        $user | add-member -type noteproperty -name "Licenses_Source" -Value $matchedUser.Licenses_Source -force
    }
}

# Add Members and Owners to Groups (Distribution and Office365) - EHN
$O365Groups = Import-Csv
$matchedMailboxes = Import-Csv
$AllErrors_Groups = @()

$progressref = ($O365Groups).count
$progresscounter = 0
foreach ($group in $O365Groups) {
    #Set Variables
    $DestinationPrimarySMTPAddress = $group.PrimarySmtpAddress_NewCompany
    $destinationDisplayName = $group.DisplayName_NewCompany

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Members and Owners to $($destinationDisplayName)"
    Write-Host "Updating Group $($destinationDisplayName).. " -ForegroundColor Cyan -NoNewline

    if ($group.RecipientTypeDetails_Source -eq "GroupMailbox" -and $group.ResourceProvisioningOptions_Source -notlike "*eam") {
        Write-Host "Office 365 Group.. " -ForegroundColor DarkGreen -NoNewline
        if ($group.Members_Source) {
            $membersArray = $group.Members_Source -split ","
            $membersArray = $membersArray | ?{$_ -ne "Migration_serviceaccount1@ehn.onmicrosoft.com"}

            Write-Host "Adding $($membersArray.count) Members .. " -ForegroundColor Cyan -NoNewline
            #Progress Bar 2
            $progressref2 = ($membersArray).count
            $progresscounter2 = 0
            foreach ($member in $membersArray) {
                #Member Check
                $memberAddress = @()
                if ($member -like "*@og-example.org*") {
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
                    Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Member -Links $memberAddress -ea Stop
                    Write-Host "." -ForegroundColor Green -NoNewline
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
        if ($group.GroupOwners_Source) {
            $membersArray = $group.GroupOwners_Source -split ","
            $membersArray = $membersArray | ?{$_ -ne "Migration_serviceaccount1@ehn.onmicrosoft.com"}

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
                    Add-UnifiedGroupLinks -Identity $DestinationPrimarySMTPAddress -LinkType Owner -Links $memberCheck.PrimarySmtpAddress_Destination -ea Stop
                    Write-Host "." -ForegroundColor Green -NoNewline
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
    elseif ($group.RecipientTypeDetails_Source -eq "MailUniversalDistributionGroup" -and $group.ResourceProvisioningOptions_Source -notlike "*eam") {
        Write-Host "DistributionGroup.. " -ForegroundColor DarkCyan -NoNewline
        if ($group.Members_Source) {
            #Gather Members
            $membersArray = $group.Members -split ","
            $membersArray = $membersArray | ?{$_ -ne "Migration_serviceaccount1@ehn.onmicrosoft.com"}

            #Progress Bar 2
            $progressref2 = ($membersArray).count
            $progresscounter2 = 0
            foreach ($member in $membersArray) {
                #Member Check
                $memberAddress = @()
                if ($member -like "*@og-example.org*") {
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

                #Add Member to Distribution Groups     
                try {
                    Add-DistributionGroupMember -Identity $group.PrimarySmtpAddress_Destination -Member $memberAddress -ea Stop
                    Write-Host ". " -ForegroundColor Green -NoNewline
                }
                catch {
                    Write-Host ". " -ForegroundColor red -NoNewline

                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $group.DisplayName_destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $group.PrimarySmtpAddress_destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Member_Source_PrimarySMTPAddress" -Value $member -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Member_Destination_PrimarySMTPAddress" -Value $memberAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors_Groups += $currenterror           
                    continue
                }
                
            }
        }
        if ($group.GroupOwners_Source) {
            #Gather Members
            $membersArray = $group.GroupOwners_Source -split ","
            $membersArray = $membersArray | ?{$_ -ne "Migration_serviceaccount1@ehn.onmicrosoft.com"}

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

                #Add Owner to Distribution Groups      
                try {
                    Set-DistributionGroup -Identity $group.PrimarySmtpAddress_destination -ManagedBy @{add=$memberCheck.PrimarySmtpAddress_Destination} -ea Stop
                    Write-Host ". " -ForegroundColor Green -NoNewline
                }
                catch {
                    Write-Host ". " -ForegroundColor red -NoNewline

                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddOwner" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $group.DisplayName_destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $group.PrimarySmtpAddress_destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Owner_Source_PrimarySMTPAddress" -Value $member -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Owner_Destination_PrimarySMTPAddress" -Value $memberCheck.PrimarySmtpAddress_Destination -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors_Groups += $currenterror           
                    continue
                }
                
            }
        }
    }
    else {
        Write-Host "Skipping.. " -ForegroundColor Yellow -NoNewline
    }
    Write-Host " done " -ForegroundColor Green
}


#Get Distribution Groups
Connect-AzureAD
$groups=Get-AzureADGroup -All $true
$resultsarray =@()
$progressref = ($groups).count
$progresscounter = 0
ForEach ($group in $groups){
    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Members and Owners to $($group.DisplayName)"
    
    $members = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true
    $progressref2 = ($members).count
    $progresscounter2 = 0
    ForEach ($member in $members){
        #Progress Bar
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Updating Members and Owners to $($group.DisplayName)"
        $UserObject = new-object PSObject
        $UserObject | add-member  -membertype NoteProperty -name "Group Name" -Value $group.DisplayName
        $UserObject | add-member  -membertype NoteProperty -name "Member Name" -Value $member.DisplayName
        $UserObject | add-member  -membertype NoteProperty -name "ObjType" -Value $member.ObjectType
        $UserObject | add-member  -membertype NoteProperty -name "UserType" -Value $member.UserType
        $UserObject | add-member  -membertype NoteProperty -name "UserPrinicpalName" -Value $member.UserPrincipalName
        $resultsarray += $UserObject
    }
}

#Get Distribution Groups
Connect-AzureAD
$groups=Get-AzureADGroup -All $true
$resultsarray =@()
$progressref = ($groups).count
$progresscounter = 0
ForEach ($group in $groups){
    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Members and Owners to $($group.DisplayName)"
    
    $members = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true | ?{$_.Mail -like "*@og-example.org"}
    $progressref2 = ($members).count
    $progresscounter2 = 0
    ForEach ($member in $members){
        #Progress Bar
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Updating Members and Owners to $($group.DisplayName)"
        $UserObject = new-object PSObject
        $UserObject | add-member  -membertype NoteProperty -name "Group Name" -Value $group.DisplayName
        $UserObject | add-member  -membertype NoteProperty -name "Member Name" -Value $member.DisplayName
        $UserObject | add-member  -membertype NoteProperty -name "ObjType" -Value $member.ObjectType
        $UserObject | add-member  -membertype NoteProperty -name "UserType" -Value $member.UserType
        $UserObject | add-member  -membertype NoteProperty -name "UserPrinicpalName" -Value $member.UserPrincipalName
        $UserObject | add-member  -membertype NoteProperty -name "Mail" -Value $member.Mail
        $UserObject | Export-Csv -NoTypeInformation -Encoding utf8 -Path C:\Users\AbacoMigration\Desktop\NewCompany\NewCompanyGroupsAndMembers.csv -Append
    }
}

# Stamp Perms to FullAccess and SendAs Groups for UserMailbox
$allmatchedMailboxes = Import-Csv
$userMailboxes = $matchedMailboxes | ?{$_.RecipientTypeDetails_Destination -eq "UserMailbox" -and $_.Migrate -ne "No"}
$AllErrorsPerms = @()
$progressref = $allMatchedSharedMailboxes.count
$progresscounter = 0
foreach ($mailbox in $allMatchedSharedMailboxes) {
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
    
    #Stamp Full Access Perms for Mailbox
    if ($mailbox.FullAccessPerms_Source) {
        $fullAccessPerms = $mailbox.FullAccessPerms_Source -split ","
        $fullAccessPermUsers = $fullAccessPerms | ?{$_ -notlike "*NAMPR16A*"}
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
    #Stamp Full Access Perms for Mailbox
    if ($mailbox.SendAsPerms_Source) {
        $SendAsPerms = $mailbox.SendAsPerms_Source -split ","
        $SendAsPermsUsers = $SendAsPerms | ?{$_ -notlike "*NAMPR16A*"}
        #Only Run for Legitimate Users
        if ($SendAsPermsUsers) {
            Write-Host "SendAs.. " -ForegroundColor DarkYellow -NoNewline
            #Progress Bar 2
            $progressref2 = ($SendAsPermsUsers).count
            $progresscounter2 = 0
            foreach ($member in $SendAsPermsUsers) {
                #Member Check
                $memberCheck = @()
                $memberCheck = $allmatchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $member}

                #Progress Bar 2a
                $progresscounter2 += 1
                $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Send As to $($memberCheck.PrimarySmtpAddress_Destination)"

                #Add Perms to Mailbox      
                try {
                    $permResult = Add-RecipientPermission -identity $destinationEmail -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -confirm:$false -ea Stop -warningaction silentlycontinue
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
    #Stamp Send On Behalf
    if ($mailbox.GrantSendOnBehalfTo_Source) {
        $SendOnBehalfPerms = $mailbox.GrantSendOnBehalfTo_Source -split ","
        $SendOnBehalfPermsUsers = $SendOnBehalfPerms | ?{$_ -notlike "*NAMPR16A*"}
        #Only Run for Legitimate Users
        if ($SendOnBehalfPermsUsers) {
            Write-Host "SendOnBehalf.. " -ForegroundColor DarkYellow -NoNewline
            #Progress Bar 2
            $progressref2 = ($SendOnBehalfPermsUsers).count
            $progresscounter2 = 0
            foreach ($member in $SendOnBehalfPermsUsers) {
                #Member Check
                $memberCheck = @()
                $memberCheck = $allmatchedMailboxes | ? {$_.DisplayName_Source -eq $member}

                #Progress Bar 2a
                $progresscounter2 += 1
                $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Granting Send As to $($memberCheck.PrimarySmtpAddress_Destination)"

                #Add Perms to Mailbox      
                try {
                    $permResult = Add-RecipientPermission -identity $destinationEmail -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -confirm:$false -ea Stop -warningaction silentlycontinue
                    Write-Host "." -ForegroundColor Green -NoNewline
                }
                catch {
                    Write-Host "." -ForegroundColor red -NoNewline

                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailtoGrantSendOnBehalf" -Force
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

## unhide objects

$nonUserMailboxes = Import-Excel -WorksheetName "Wave5 Non Users2" -Path
$nonUserMailboxes2 = $nonUserMailboxes | ?{$_.HiddenFromAddressListsEnabled_Source -eq $false}

$AllErrors = @()
$progressref = $allMatchedSharedMailboxes.count
$progresscounter = 0
foreach ($mailbox in $allMatchedSharedMailboxes) {
    #Set Variables
    $sourceEmail = $mailbox.PrimarySmtpAddress_Source
    $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
    $destinationDisplayName = $mailbox.DisplayName_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating GAL attributes: $($destinationDisplayName)"
        
    try {
        Set-Mailbox $destinationEmail -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($mailbox.HiddenFromAddressListsEnabled_Source)) -ErrorAction Stop -warningaction silentlycontinue
    }
    catch {
        Write-Host ". " -ForegroundColor red -NoNewline

        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnabletoHideFromGAL" -Force
        $currenterror | Add-Member -type NoteProperty -Name "SourceEmail" -Value $sourceEmail -Force
        $currenterror | Add-Member -type NoteProperty -Name "destinationEmail" -Value $destinationEmail -Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllErrors += $currenterror           
        continue
    } 
}

#Get Mailbox Match Details
$Wave5UserMailboxes = Import-Excel -WorksheetName "Wave 5 Users - 4368" -path "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\All-MatchedMailboxes_3-25_2022.xlsx"
$NonUserMailboxes = Import-Excel -WorksheetName "Wave 5 Non Users - 658" -path "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\All-MatchedMailboxes_3-25_2022.xlsx"
$GroupMailboxes = Import-CSV "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\MatchedGroups_EHN-TJU.csv"

$MigrationJobDetails  = Import-CSV "C:\Users\amedrano\Desktop\MatchedItems\All-MatchedMailboxes_TMP 3-25.csv"
$sharedMailboxes = import-csv 
$foundMigJob = @()
$notfoundMigJob = @()

$progressref = $sharedRemainingMailboxes.count
$progresscounter = 0
foreach ($user in $sharedRemainingMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details UserMailbox $($user.PrimarySmtpAddress_Source)"

    if ($matchedUser = $allmatchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $user.PrimarySmtpAddress_Source}) {
        $foundMigJob += $matchedUser
    }
    else {
        $notfoundMigJob += $user
    }
}


# Stamp Perms to Full Access and SendAs - Shared Mailboxes
$allmatchedMailboxes = Import-Csv
$sharedMailboxes = Import-Csv "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\Batches\Wave5-NonUsers_Old.csv"

$AllErrorsPerms = @()
$progressref = $sharedMailboxes.count
$progresscounter = 0
foreach ($mailbox in $sharedMailboxes) {
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
        $fullAccessPermUsers = $fullAccessPerms | ?{$_ -notlike "*NAMPR16A*"}
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
        $SendAsPermsUsers = $SendAsPerms | ?{$_ -notlike "*NAMPR16A*"}
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
                    $permResult = Add-RecipientPermission -identity $destinationEmail -AccessRights SendAs -Trustee $memberCheck.PrimarySmtpAddress_Destination -confirm:$false -ea Stop -warningaction silentlycontinue
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

#update UPN
$AllErrorsPerms = @()
$progressref = $sharedMailboxes.count
$progresscounter = 0
foreach ($mailbox in $sharedMailboxes) {
    #Set Variables
    $sourceEmail = $mailbox.PrimarySmtpAddress_Source
    $destinationEmail = $mailbox.PrimarySMTPAddress_Destination
    $destinationDisplayName = $mailbox.DisplayName_Destination

    if ($mailboxCheck = Get-Mailbox $destinationEmail) {
        #Progress Bar
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating UPN from $($mailboxCheck.UserPrincipalName) to $($destinationEmail)"
        
        if ($mailboxCheck.UserPrincipalName -eq $destinationEmail) {
            Write-host "Updating UPN from $($mailboxCheck.UserPrincipalName) to $($destinationEmail)" -foregroundcolor green
            Set-MsolUserPrincipalName -NewUserPrincipalName $destinationEmail -UserPrincipalName $mailboxCheck.UserPrincipalName
        }
        else {
            Write-host "No Update to UPN. UPN already set to  Found for $($mailboxCheck.UserPrincipalName)" -foregroundcolor yellow
        }
    }
    else {
        Write-host "No Mailbox Found for $($destinationEmail)" -foregroundcolor red
    }

}

#Pull OldCompany VIP Delegate Perms
$OldCompanyVIPS = Import-Excel "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\VIP lists\OldCompany Updated list of VIP 10 4 21.xlsx"


#REGION Get list of calendar permissions


#Build Array
$calendarpermsList = @()
$perms = @()

#ProgressBar
$progressref = ($OldCompanyVIPS).count
$progresscounter = 0
foreach ($mbx in $OldCompanyVIPS) {
    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($mbx.DisplayName)"

    #Clear Variables
    $upn = $null
    $id = $null
    $inboxID = $null
    $sentItemsID = $null

    #Set Variables
	$upn = $mbx.PrimarySMTPAddress_Source
    Write-Host "Checking Mailbox $($upn) .."  -ForegroundColor Cyan -NoNewline

    #Gather Calendar Perms
    try {
        [array]$calendars = Get-MailboxFolderStatistics $upn -ea stop | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}
    }
    catch {
        Write-Host "Unable to find folders for mailbox. Mailbox possibly does not exist .." -foregroundcolor red -nonewline
    }
	
	foreach ($calendar in $calendars) 	{
		$folderPath = $calendar.FolderPath.Replace('/','\')
		$id = "$upn`:$folderPath"

		[array]$perms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.user.usertype.value -ne "Default" -and $_.user.usertype.value -ne "Anonymous" -and $_.user.usertype.value -notlike "*S-1-*"}
		if ($perms) {

            Write-Host "$($upn)`:" -ForegroundColor Cyan -NoNewline
            Write-Host $folderPath -ForegroundColor Green -NoNewline
            Write-Host " ..." -ForegroundColor darkCyan -NoNewline

            #ProgressBar
            $progressref2 = ($perms).count
            $progresscounter2 = 0
			foreach ($perm in $perms) {
                #Progress Bar
                $progresscounter2 += 1
                $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
                $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
                Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Gathering Perm Details for $($perm.user.DisplayName.tostring())"

				$accessRights = $perm.AccessRights -join ";"
				$SharingPermissionFlags = $perm.SharingPermissionFlags -join ";"
				
                $currentPerm = new-object PSObject				
				$currentPerm | add-member -type noteproperty -name "Mailbox" -Value $upn.ToString() -force
				$currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $id -force
				
				if ($recipientCheck = Get-Mailbox $perm.user.DisplayName.tostring() -ea silentlycontinue) {
                    $currentPerm | add-member -type noteproperty -name "PermUser_Name" -Value $perm.user.DisplayName.tostring() -Force
					$currentPerm | add-member -type noteproperty -name "PermUser_DisplayName" -Value $recipientCheck.DisplayName.ToString() -Force
					$currentPerm | add-member -type noteproperty -name "PermUser_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
				}
				else {
                    $currentPerm | add-member -type noteproperty -name "PermUser_Name" -Value $null -Force
					$currentPerm | add-member -type noteproperty -name "PermUser_DisplayName" -Value $null -Force
					$currentPerm | add-member -type noteproperty -name "PermUser_PrimarySMTPAddress" -Value $null -force
				}
				$currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights  -force
				$currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags  -force
				$calendarpermsList += $currentPerm
				Write-Host "." -ForegroundColor Yellow -NoNewline
			}
		}
	}

    #Add Inbox Permissions
    $inboxID = $upn + ":\Inbox"
    [array]$inboxperms = Get-MailboxFolderPermission $inboxID -EA SilentlyContinue | Where {$_.user.usertype.value -ne "Default" -and $_.user.usertype.value -ne "Anonymous" -and $_.user.usertype.value -notlike "*S-1-*"}
    if ($inboxperms) {
        Write-Host "$($inboxID)"  -ForegroundColor Cyan -NoNewline
		Write-Host " ..." -ForegroundColor darkCyan -NoNewline
        
        #ProgressBar
        $progressref2 = ($inboxperms).count
        $progresscounter2 = 0

        foreach ($perm in $inboxperms) 	{
            #Progress Bar
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Gathering Perm Details for $($perm.user.DisplayName.tostring())"

            $accessRights = $perm.AccessRights -join ";"
            $SharingPermissionFlags = $perm.SharingPermissionFlags -join ";"
            
            $currentPerm = new-object PSObject				
            $currentPerm | add-member -type noteproperty -name "Mailbox" -Value $upn.ToString() -force
            $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $inboxID -force
            
            if ($recipientCheck = Get-Mailbox $perm.user.DisplayName.tostring() -ea silentlycontinue) {
                $currentPerm | add-member -type noteproperty -name "PermUser_Name" -Value $perm.user.DisplayName.tostring() -Force
                $currentPerm | add-member -type noteproperty -name "PermUser_DisplayName" -Value $recipientCheck.DisplayName.ToString() -Force
                $currentPerm | add-member -type noteproperty -name "PermUser_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
            }
            else {
                $currentPerm | add-member -type noteproperty -name "PermUser_Name" -Value $null -Force
                $currentPerm | add-member -type noteproperty -name "PermUser_DisplayName" -Value $null -Force
                $currentPerm | add-member -type noteproperty -name "PermUser_PrimarySMTPAddress" -Value $null -force
            }
            $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights -force
            $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags -force
            $calendarpermsList += $currentPerm
            Write-Host "." -ForegroundColor Yellow -NoNewline
        }
    }
    #Add Sent Items Permissions
    $sentItemsID = $upn + ":\Sent Items"
    [array]$sentItemsPerms = Get-MailboxFolderPermission $inboxID -EA SilentlyContinue | Where {$_.user.usertype.value -ne "Default" -and $_.user.usertype.value -ne "Anonymous" -and $_.user.usertype.value -notlike "*S-1-*"}
    if ($sentItemsPerms) {
        Write-Host "$($sentItemsID)"  -ForegroundColor Cyan -NoNewline
		Write-Host " ..." -ForegroundColor darkCyan -NoNewline
        #ProgressBar
        $progressref2 = ($sentItemsPerms).count
        $progresscounter2 = 0

        foreach ($perm in $sentItemsPerms) 	{
            #Progress Bar
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Gathering Perm Details for $($perm.user.DisplayName.tostring())"

            $accessRights = $perm.AccessRights -join ";"
            $SharingPermissionFlags = $perm.SharingPermissionFlags -join ";"
           
            $currentPerm = new-object PSObject				
            $currentPerm | add-member -type noteproperty -name "Mailbox" -Value $upn.ToString() -force
            $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $sentItemsID -force
            
            if ($recipientCheck = Get-Mailbox $perm.user.DisplayName.tostring() -ea silentlycontinue) {
                $currentPerm | add-member -type noteproperty -name "PermUser_Name" -Value $perm.user.DisplayName.tostring() -Force
                $currentPerm | add-member -type noteproperty -name "PermUser_DisplayName" -Value $recipientCheck.DisplayName.ToString() -Force
                $currentPerm | add-member -type noteproperty -name "PermUser_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
            }
            else {
                $currentPerm | add-member -type noteproperty -name "PermUser_Name" -Value $null -Force
                $currentPerm | add-member -type noteproperty -name "PermUser_DisplayName" -Value $null -Force
                $currentPerm | add-member -type noteproperty -name "PermUser_PrimarySMTPAddress" -Value $null -force
            }
            $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights -force
            $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags -force
            $calendarpermsList += $currentPerm
            Write-Host "." -ForegroundColor Yellow -NoNewline
        }
    }
    Write-Host "done" -ForegroundColor green
}


#Grab Delegate Perms Combined

#Build Array
$calendarpermsList = @()
$perms = @()

#ProgressBar
$progressref = ($OldCompanyVIPS).count
$progresscounter = 0
foreach ($mbx in $OldCompanyVIPS) {
    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($mbx.DisplayName)"

    #Clear Variables
    $upn = $null
    $id = $null
    $inboxID = $null
    $sentItemsID = $null

    #Set Variables
	$upn = $mbx.PrimarySMTPAddress_Source
    Write-Host "Checking Mailbox $($upn) for unique folder perms.."  -ForegroundColor Cyan -NoNewline

    #Gather Calendar Perms
    try {
        [array]$folders = Get-MailboxFolderStatistics $upn -ea stop | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*" -or $_.FolderPath -eq "/inbox" -or $_.FolderPath -eq "/sent items"}
    }
    catch {
        Write-Host "Unable to find folders for mailbox. Mailbox possibly does not exist .." -foregroundcolor red -nonewline
    }
	#ProgressBar
    $progressref2 = ($folders).count
    $progresscounter2 = 0
	foreach ($folder in $folders) {
        #set Variables
        $folderPath = $folder.FolderPath.Replace('/','\')
		$id = "$upn`:$folderPath"

        #Progress Bar
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Gathering Perm Details for $($folderPath)"
		
		[array]$perms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.user.usertype.value -ne "Default" -and $_.user.usertype.value -ne "Anonymous" -and $_.user.usertype.value -notlike "*S-1-*"}
		if ($perms) {
            Write-Host $folderPath -ForegroundColor Yellow -NoNewline
            Write-Host ".." -foregroundcolor DarkCyan -nonewline

            #ProgressBar
            $progressref3 = ($perms).count
            $progresscounter3 = 0
			foreach ($perm in $perms) {
                #Progress Bar
                $progresscounter3 += 1
                $progresspercentcomplete3 = [math]::Round((($progresscounter3 / $progressref3)*100),2)
                $progressStatus3 = "["+$progresscounter3+" / "+$progressref3+"]"
                Write-progress -id 3 -PercentComplete $progresspercentcomplete3 -Status $progressStatus3 -Activity "Gathering Perm Details for $($perm.user.DisplayName.tostring())"

				$accessRights = $perm.AccessRights -join ";"
				$SharingPermissionFlags = $perm.SharingPermissionFlags -join ";"
				
                $currentPerm = new-object PSObject				
				$currentPerm | add-member -type noteproperty -name "Mailbox" -Value $upn.ToString() -force
				$currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $id -force
				
				if ($recipientCheck = Get-Mailbox $perm.user.DisplayName.tostring() -ea silentlycontinue) {
                    $currentPerm | add-member -type noteproperty -name "PermUser_Name" -Value $perm.user.DisplayName.tostring() -Force
					$currentPerm | add-member -type noteproperty -name "PermUser_DisplayName" -Value $recipientCheck.DisplayName.ToString() -Force
					$currentPerm | add-member -type noteproperty -name "PermUser_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
				}
				else {
                    $currentPerm | add-member -type noteproperty -name "PermUser_Name" -Value $null -Force
					$currentPerm | add-member -type noteproperty -name "PermUser_DisplayName" -Value $null -Force
					$currentPerm | add-member -type noteproperty -name "PermUser_PrimarySMTPAddress" -Value $null -force
				}
				$currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights  -force
				$currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags  -force
				$calendarpermsList += $currentPerm
			}
		}
	}
    Write-host " done" -foregroundcolor green
}

#Gather All Teams URLs
$msteams = Get-Team
$OldCompanyTeams = 
$NewCompanyTeams = 
#ProgressBar
$progressref = ($OldCompanyTeams).count
$progresscounter = 0
foreach ($team in $OldCompanyTeams) {
    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Teams Details for $($team.TeamName_Source)"

    if ($matchedTeam = $NewCompanyTeams | ?{$_."Team Name" -eq $team.TeamName_Source}){
        $team | add-member -type noteproperty -name "TeamName_Destination" -Value $matchedTeam."Team Name" -force
        $team | add-member -type noteproperty -name "SharePointSiteURL_Destination" -Value $matchedTeam."SharePoint Site URL" -force
        $team | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $matchedTeam."Primary SMTP Address" -force
        $team | add-member -type noteproperty -name "ManagedBy_Destination" -Value $matchedTeam."Managed By" -force
        $team | add-member -type noteproperty -name "Accesstype_Destination" -Value $matchedTeam."Access type" -force
        $team | add-member -type noteproperty -name "CreationTime_Destination" -Value $matchedTeam."Creation Time"  -force
    }
    else {
        $team | add-member -type noteproperty -name "TeamName_Destination" -Value $null -force
        $team | add-member -type noteproperty -name "SharePointSiteURL_Destination" -Value $null -force
        $team | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $null -force
        $team | add-member -type noteproperty -name "ManagedBy_Destination" -Value $null -force
        $team | add-member -type noteproperty -name "Accesstype_Destination" -Value $null -force
        $team | add-member -type noteproperty -name "CreationTime_Destination" -Value $null  -force
    } 
}

#Gather EXOMailbox Stats
$OutputCSVFolderPath = Read-Host "What is the folder path to store the file?"
$sourceMailboxes = Get-EXOMailbox -ResultSize Unlimited -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'"
$sourceMailboxStats = @()
$progressref = ($sourceMailboxes).count
$progresscounter = 0
foreach ($user in $sourceMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    
    #Pull MailboxStats and UserDetails
    $mbxStats = Get-EXOMailboxStatistics $user.PrimarySMTPAddress -PropertySets All 
    $msoluser = Get-MsolUser -UserPrincipalName $user.UserPrincipalName
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    #Create User Array
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $msoluser.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName_Source" -Value $msoluser.userprincipalname
    $currentuser | add-member -type noteproperty -name "IsLicensed_Source" -Value $msoluser.IsLicensed
    $currentuser | add-member -type noteproperty -name "Licenses_Source" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
    $currentuser | add-member -type noteproperty -name "License-DisabledArray_Source" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
    $currentuser | add-member -type noteproperty -name "BlockCredential_Source" -Value $msoluser.BlockCredential
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
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-EXOMailboxPermission $user.primarysmtpaddress | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value ($mbxPermissions.user -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value $null
    }
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-EXORecipientPermission -AccessRights SendAs -Identity $user.PrimarySMTPAddress | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value $null
    }
    # Archive Mailbox Check
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus
    if ($ArchiveStats = Get-EXOMailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {           
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $ArchiveStats.TotalItemSize.Value
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $ArchiveStats.ItemCount
    }
    else {
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $null
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $null
    }
    $sourceMailboxStats += $currentuser
}
$sourceMailboxStats | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$OutputCSVFolderPath\SourceMailboxes.csv"


#Gather SharedMailboxes
$OutputCSVFolderPath = Read-Host "What is the folder path to store the file?"
$sourceSharedMailboxes = Get-Mailbox -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'UserMailbox'"
$sourceSharedMailboxStats = @()
$progressref = ($sourceSharedMailboxes).count
$progresscounter = 0
foreach ($user in $sourceSharedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    
    #Pull MailboxStats and UserDetails
    $mbxStats = Get-MailboxStatistics $user.DistinguishedName
    $msoluser = Get-MsolUser -UserPrincipalName $user.UserPrincipalName
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    #Create User Array
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName_Source" -Value $msoluser.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName_Source" -Value $msoluser.userprincipalname
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $user.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress_Source" -Value $user.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "EmailAddresses_Source" -Value ($EmailAddresses -join ";")
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN_Source" -Value ("x500:" + $user.legacyexchangedn)
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Source" -Value $user.HiddenFromAddressListsEnabled
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward_Source" -Value $user.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ForwardingAddress_Source" -Value $user.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress_Source" -Value $user.ForwardingSmtpAddress

    #Pull Send on Behalf
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.DistinguishedName | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value ($mbxPermissions.user -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value $null
    }
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.DistinguishedName | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value $null
    }
    $sourceSharedMailboxStats += $currentuser
}
$sourceSharedMailboxStats | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$OutputCSVFolderPath\SourceSharedMailboxes.csv"

#Match Mailboxes and add to same spreadsheet. Check based on Campus Key, CustomAttribute7 and DisplayName - CONDENSED
$sourceMailboxes = import-csv 
$progressref = ($sourceMailboxes).count
$progresscounter = 0

foreach ($user in $sourceMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName_Source)"
    $NewUPN = $user.CustomAttribute7_Source + "@example.org"
    $CampusKeyUPN = $user.CampusKey + "@example.org"
    $addressSplit = $user.PrimarySmtpAddress_Source -split "@"
    $ehnPrimarySMTPAddress = $addressSplit[0] + "-old@example.org"
    $newPrimarySMTPAddress = $addressSplit[0] + "@example.org"

    # Campus Key Match
    if ($msoluser = Get-Msoluser -UserPrincipalName $CampusKeyUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with CampusKey" -ForegroundColor Green
    }
    #New UPN Match - CustomAttribute7
    elseif ($msoluser = Get-Msoluser -UserPrincipalName $NewUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($msoluser.UserPrincipalName) User found with CampusKey" -ForegroundColor Green
    }

    #NEW PrimarySMTPAddress Check (EHN added)
    elseif ($mailboxCheck = Get-Mailbox $ehnPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        $msoluser = Get-Msoluser -UserPrincipalName $mailboxCheck.UserPrincipalName -ErrorAction SilentlyContinue
        Write-Host "$($msoluser.UserPrincipalName) User found with EHN added" -ForegroundColor Yellow
    }
    #NEW PrimarySMTPAddress Check
    elseif ($mailboxCheck = Get-Mailbox $newPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        $msoluser = Get-Msoluser -UserPrincipalName $mailboxCheck.UserPrincipalName -ErrorAction SilentlyContinue
        Write-Host "$($msoluser.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
    }
    #UPN Check based on DisplayName
    elseif ($msoluser = Get-Msoluser -SearchString $user.DisplayName -ErrorAction SilentlyContinue) {
        if ($msoluser.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName_Source)"
            $mailbox = Get-Mailbox $user.DisplayName -ErrorAction SilentlyContinue
            $msoluser = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        }
        Write-Host "$($msoluser.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
    }

    if ($msoluser) {
        #Gather Stats
        $mailbox = Get-EXOMailbox -PropertySets archive,addresslist,delivery,Minimum $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses

        #Output Stats
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
    }
    else {
        Write-Host "  Unable to find user for $($user.DisplayName_Source)" -ForegroundColor Red

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $null -force
    }
}

#Match Mailboxes and add to same spreadsheet. Check based on Campus Key, CustomAttribute7 and DisplayName - CONDENSED
$sourceMailboxes = import-csv 
$progressref = ($sourceMailboxes).count
$progresscounter = 0
 
foreach ($user in $sourceMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName_Source)"
    $NewUPN = $user.CustomAttribute7_Source + "@example.org"
    #$CampusKeyUPN = $user.CampusKey + "@example.org"
    $addressSplit = $user.PrimarySmtpAddress_Source -split "@"
    $ehnPrimarySMTPAddress = $addressSplit[0] + "-old@example.org"
    $newPrimarySMTPAddress = $addressSplit[0] + "@example.org"
    #New UPN Match - CustomAttribute7
    if ($mailbox = Get-Mailbox $NewUPN -ErrorAction SilentlyContinue) {
        Write-Host "$($mailbox.UserPrincipalName) User found with CampusKey" -ForegroundColor Green
    }

    #NEW PrimarySMTPAddress Check (EHN added)
    elseif ($mailbox = Get-Mailbox $ehnPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($mailbox.UserPrincipalName) User found with EHN added" -ForegroundColor Yellow
    }
    #NEW PrimarySMTPAddress Check
    elseif ($mailbox = Get-Mailbox $newPrimarySMTPAddress -ErrorAction SilentlyContinue) {
        Write-Host "$($mailbox.UserPrincipalName) User found with NewPrimarySMTPAddress" -ForegroundColor Yellow
    }
    #UPN Check based on DisplayName
    elseif ($mailbox = Get-Mailbox -SearchString $user.DisplayName_Source -ErrorAction SilentlyContinue) {
        if ($mailbox.count -gt 1) {
            Write-Host "Multiple Users found. Checking for mailbox $($user.DisplayName_Source)"
        }
        Write-Host "$($mailbox.UserPrincipalName) User found with DisplayName" -ForegroundColor Yellow
    }

    if ($mailbox) {
        #Gather Stats
        $EmailAddresses = $mailbox | select -ExpandProperty EmailAddresses

        #Output Stats
        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $True -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $mailbox.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $mailbox.userprincipalname -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $mailbox.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
    }
    else {
        Write-Host "  Unable to find user for $($user.DisplayName_Source)" -ForegroundColor Red

        $user | add-member -type noteproperty -name "ExistsInDestination" -Value $False -force
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $null -force
    }
}

#Get SharePoint Group Perm Details
$matchedTeams = Import-excel "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\ShareGate\Teams\MatchedTeams.xlsx"
$progressref = $matchedTeams.count
$progresscounter = 0
$OldCompanyTeamsPerms = @()
foreach ($object in $matchedTeams) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details $($object.SharePointSiteURL_Destination)"

    $SPOSiteGroups = Get-SPOSiteGroup -Site $object.SharePointSiteURL_Destination
    foreach ($siteGroup in $SPOSiteGroups) {
        $roles = $siteGroup.Roles -join ','
        $newObject = New-Object PSObject
        $newObject | add-member -type noteproperty -name "TeamName_Destination" -Value $object.TeamName_Destination -force
        $newObject | add-member -type noteproperty -name "SharePointSiteURL_Destination" -Value $object.SharePointSiteURL_Destination -force
        $newObject | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $object.PrimarySMTPAddress_Destination -force
        $newObject | add-member -type noteproperty -name "ManagedBy_Destination" -Value $object.ManagedBy_Destination -force
        $newObject | add-member -type noteproperty -name "Accesstype_Destination" -Value $object.Accesstype_Destination -force
        $newObject | add-member -type noteproperty -name "CreationTime_Destination" -Value $object.CreationTime_Destination -force
        $newObject | add-member -type noteproperty -name "GroupTitle" -Value $siteGroup.Title -force
        $newObject | add-member -type noteproperty -name "GroupPerms" -Value $roles -force
        $OldCompanyTeamsPerms += $newObject
    }
}

#Gather One Drive Details for Matched Accounts
$NewCompanyOneDriveDetails = import-csv "C:\Users\amedrano\Downloads\NewCompanyOneDriveUsageAccountDetail4_18_2022 1_43_02 PM.csv"
$OldCompanyOneDriveDetails = import-csv "C:\Users\amedrano\Downloads\OldCompanyOneDriveUsageAccountDetail4_18_2022 1_40_59 PM.csv"
$allmatchedMailboxes = import-csv "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\AllMatched_Mailboxes.csv"
$progressref = $allmatchedMailboxes.count
$progresscounter = 0
foreach ($mailbox in $allmatchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details $($mailbox.UserPrincipalName_Source)"
    #gatherOneDriveDetails - OldCompany
    if ($matchedOldCompany = $OldCompanyOneDriveDetails | ?{$_."Owner Principal Name" -eq $mailbox.UserPrincipalName_Source}) {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Source" -Value $matchedOldCompany."Site URL" -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Source" -Value $matchedOldCompany."Last Activity Date" -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Source" -Value $matchedOldCompany."File Count" -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Source" -Value $matchedOldCompany."Storage Used (Byte)" -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Source" -Value $null -force
    }
    #gatherOneDriveDetails - NewCompany
    if ($matchedNewCompany = $NewCompanyOneDriveDetails | ?{$_."Owner Principal Name" -eq $mailbox.UserPrincipalName_Destination}) {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Destination" -Value $matchedNewCompany."Site URL" -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Destination" -Value $matchedNewCompany."Last Activity Date" -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Destination" -Value $matchedNewCompany."File Count" -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Destination" -Value $matchedNewCompany."Storage Used (Byte)" -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageusage_Destination" -Value $null -force
    } 
}

#Remove OldCompany Mail Contacts in NewCompany
$OldCompanyMailContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*example.org'" -ResultSize unlimited
$progressref = $OldCompanyMailContacts.count
$progresscounter = 0
foreach ($object in $OldCompanyMailContacts) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Removing Mail Contact $($object.PrimarySmtpAddress)"

    if (!($recipientCheck = Get-Recipient $object.PrimarySmtpAddress -ErrorAction SilentlyContinue)) {
        Remove-MailContact -confirm:$false $object.PrimarySmtpAddress
    }
}

#Remove OldCompany Mail Contacts in NewCompany
$OldCompanyMailContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*example.org'" -ResultSize unlimited
#$lastOldCompanyMailContacts = $OldCompanyMailContacts[-1..-10000]
$progressref = $OldCompanyMailContacts.count
$progresscounter = 0
foreach ($object in $OldCompanyMailContacts) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Removing Mail Contact $($object.PrimarySmtpAddress)"

    if (!($recipientCheck = Get-Recipient $object.PrimarySmtpAddress -ErrorAction SilentlyContinue)) {
        Remove-MailContact -confirm:$false $object.PrimarySmtpAddress
        Write-Host "." -ForegroundColor Green -NoNewline
    }
}

#Check Archiving Stats
foreach ($object in $allmatchedMailboxes) {
    $mailbox = 
    if ($ArchiveStats = Get-MailboxStatistics $mailbox.PrimarySMTPAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {
        $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $ArchiveStats.TotalItemSize.Value -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $ArchiveStats.ItemCount -force
    }
    else  {
        $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
    }
}

## Set SharePoint Group Perm
$matchedTeams = Import-excel "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\ShareGate\Teams\MatchedTeams.xlsx"

#Variables for Admin Center & Site Collection URL
$AdminCenterURL = "https://tjuv-admin.sharepoint.com"
#Connect to SharePoint Online
Connect-SPOService -url $AdminCenterURL -Credential (Get-Credential)

$MembersGroups = $matchedTeams | ?{$_.GroupTitle -like "*Members" -and $_.GroupPerms -like "*Edit*"}
$progressref = $MembersGroups.count
$progresscounter = 0
foreach ($object in $MembersGroups) {
    #Set Variables
    $GroupName = $object.GroupTitle
    $SiteURL = $object.SharePointSiteURL_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Setting MembersGroup Details $($SiteURL)"
       
    #Update Member Group to Add Edit Permissions
    Write-Host "Updating $($GroupName) Perms to Edit .." -ForegroundColor Cyan -nonewline
    $permResult = Set-SPOSiteGroup -Site $SiteURL -Identity $GroupName -PermissionLevelsToAdd "Edit"
    Write-Host "Completed" -ForegroundColor Green
}

#Gather One Drive Details for Matched Accounts
$migrationWizDetails = import-csv 
$allmatchedMailboxes = import-csv "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany - OldCompany to NewCompany Migration\Exchange Online\AllMatched_Mailboxes.csv"
$progressref = $allmatchedMailboxes.count
$progresscounter = 0
foreach ($mailbox in $allmatchedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details $($mailbox.UserPrincipalName_Source)"
    #gatherOneDriveDetails - OldCompany

    if ($matchedMigrationStats = $migrationWizDetails | ?{$_."Owner Principal Name" -eq $mailbox.UserPrincipalName_Source}) {
        if (condition) {
            <# Action to perform if the condition is true #>
        }
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Source" -Value $matchedOldCompany."Site URL" -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Source" -Value $matchedOldCompany."Last Activity Date" -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Source" -Value $matchedOldCompany."File Count" -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Source" -Value $matchedOldCompany."Storage Used (Byte)" -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Source" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Source" -Value $null -force
    }
    #gatherOneDriveDetails - NewCompany
    if ($matchedNewCompany = $NewCompanyOneDriveDetails | ?{$_."Owner Principal Name" -eq $mailbox.UserPrincipalName_Destination}) {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Destination" -Value $matchedNewCompany."Site URL" -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Destination" -Value $matchedNewCompany."Last Activity Date" -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Destination" -Value $matchedNewCompany."File Count" -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageUsage(Byte)_Destination" -Value $matchedNewCompany."Storage Used (Byte)" -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "OneDriveURL_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveLastActivityDate_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveFileCount_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "OneDriveStorageusage_Destination" -Value $null -force
    } 
}





#Match SharePoint Sites cross tenant
$EhnSites = import-csv
$TJUVSites = import-csv

$progressref = $EhnSites.count
$progresscounter = 0
foreach ($site in $EhnSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site $($Site.Site)"

    $matchedSite = @()
    #Check for EHN Labeled Site
    $EHNLabeledSite = $site.site +"-old"
    $EHNSite = $site.Site
    if ($matchedSite = $TJUVSites | ? {$_.Site -eq $EHNLabeledSite}) {
        Write-Host "."  -foregroundcolor Green -NoNewline
        $site | add-member -type noteproperty -name "SiteURL_Destination" -Value $matchedSite."SiteURL" -force
        $site | add-member -type noteproperty -name "Site_Destination" -Value $matchedSite.site -force
        $site | add-member -type noteproperty -name "RootWebTemplate_Destination" -Value $matchedSite."RootWebTemplate" -force
        $site | add-member -type noteproperty -name "Owner_Destination" -Value $matchedSite."OwnerPrincipalName" -force
        $site | add-member -type noteproperty -name "IsDeleted_Destination" -Value $matchedSite."IsDeleted" -force
        $site | add-member -type noteproperty -name "FileCount_Destination" -Value $matchedSite."FileCount" -force
        $site | add-member -type noteproperty -name "ActiveFileCount_Destination" -Value $matchedSite."ActiveFileCount" -force
        $site | add-member -type noteproperty -name "StorageUsed_Destination(Byte)" -Value $matchedSite."StorageUsed(Byte)" -force
        $site | add-member -type noteproperty -name "LastActivityDate_Destination" -Value $matchedSite."LastActivityDate" -force
        $site | add-member -type noteproperty -name "PageViewCount_Destination" -Value $matchedSite."PageViewCount" -force
        $site | add-member -type noteproperty -name "VisitedPageCount_Destination" -Value $matchedSite."VisitedPageCount" -force
    }
    #Check for Site Match Regular
    elseif ($matchedSite2 = $TJUVSites | ? {$_.Site -eq $EHNSite}) {
        Write-Host "."  -foregroundcolor Yellow -NoNewline
        $site | add-member -type noteproperty -name "SiteURL_Destination" -Value $matchedSite2."SiteURL" -force
        $site | add-member -type noteproperty -name "Site_Destination" -Value $matchedSite2.site -force
        $site | add-member -type noteproperty -name "RootWebTemplate_Destination" -Value $matchedSite2."RootWebTemplate" -force
        $site | add-member -type noteproperty -name "Owner_Destination" -Value $matchedSite2."OwnerPrincipalName" -force
        $site | add-member -type noteproperty -name "IsDeleted_Destination" -Value $matchedSite2."IsDeleted" -force
        $site | add-member -type noteproperty -name "FileCount_Destination" -Value $matchedSite2."FileCount" -force
        $site | add-member -type noteproperty -name "ActiveFileCount_Destination" -Value $matchedSite2."ActiveFileCount" -force
        $site | add-member -type noteproperty -name "StorageUsed_Destination(Byte)" -Value $matchedSite2."StorageUsed(Byte)" -force
        $site | add-member -type noteproperty -name "LastActivityDate_Destination" -Value $matchedSite2."LastActivityDate" -force
        $site | add-member -type noteproperty -name "PageViewCount_Destination" -Value $matchedSite2."PageViewCount" -force
        $site | add-member -type noteproperty -name "VisitedPageCount_Destination" -Value $matchedSite2."VisitedPageCount" -force
    }
    else {
        Write-Host "."  -foregroundcolor Red -NoNewline
        $site | add-member -type noteproperty -name "SiteURL_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "Site_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "RootWebTemplate_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "Owner_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "IsDeleted_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "FileCount_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "ActiveFileCount_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "StorageUsed_Destination(Byte)" -Value $null -force
        $site | add-member -type noteproperty -name "LastActivityDate_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "PageViewCount_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "VisitedPageCount_Destination" -Value $null -force
    }
}

#Match SharePoint Sites cross tenant - Site URL
$matchedSites = import-csv
$progressref = $matchedSites.count
$progresscounter = 0
foreach ($site in $matchedSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site $($Site.Title_Source)"

    $foundSPOSite = @()
    #Check for EHN Labeled Site
    $EHNSite = $site.SiteURL_Destination
    $EHNTeamSite = $site.MatchedTeamsSharePointSiteURL_Destination
    if ($EHNSite) {
        if ($foundSPOSite = Get-SPOSITE -Identity $EHNSite) {
            Write-Host "."  -foregroundcolor Green -NoNewline
            $site | add-member -type noteproperty -name "Title_DestinationFound" -Value $foundSPOSite.Title -force
            $site | add-member -type noteproperty -name "SiteUrl_DestinationFound" -Value $foundSPOSite.URL -force
            $site | add-member -type noteproperty -name "TeamTitle_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "TeamSiteURL_DestinationFound" -Value $null.URL -force
        }
        else {
            Write-Host "."  -foregroundcolor red -NoNewline
            $site | add-member -type noteproperty -name "Title_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "SiteUrl_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "TeamTitle_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "TeamSiteURL_DestinationFound" -Value $null.URL -force
        }
    }
    elseif ($EHNTeamSite) {
        if ($foundSPOSite = Get-SPOSITE -Identity $EHNSite) {
            Write-Host "."  -foregroundcolor yellow -NoNewline
            $site | add-member -type noteproperty -name "Title_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "SiteUrl_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "TeamTitle_DestinationFound" -Value $foundSPOSite.Title -force
            $site | add-member -type noteproperty -name "TeamSiteURL_DestinationFound" -Value $foundSPOSite.URL -force
        } 
        else {
            Write-Host "."  -foregroundcolor red -NoNewline
            $site | add-member -type noteproperty -name "Title_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "SiteUrl_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "TeamTitle_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "TeamSiteURL_DestinationFound" -Value $null -force
        }
    }
    else {
        Write-Host "."  -foregroundcolor red -NoNewline
        $site | add-member -type noteproperty -name "Title_DestinationFound" -Value $null -force
        $site | add-member -type noteproperty -name "SiteUrl_DestinationFound" -Value $null -force
        $site | add-member -type noteproperty -name "TeamTitle_DestinationFound" -Value $null -force
        $site | add-member -type noteproperty -name "TeamSiteURL_DestinationFound" -Value $null -force
    }
    start-sleep -Milliseconds 100
}

#Match SharePoint Sites cross tenant - Site URL
$matchedSites = import-csv
$progressref = $matchedSites.count
$progresscounter = 0
foreach ($site in $matchedSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site $($Site.Title_Source)"

    $foundSPOSite = @()
    #Check for EHN Labeled Site
    $EHNSite = $site.SiteURL_Destination
    $EHNTeamSite = $site.MatchedTeamsSharePointSiteURL_Destination
    if ($ehnSite -or $ehnTeamSite) {
        if ($foundSPOSite = Get-SPOSITE -Identity $EHNSite) {
            Write-Host "."  -foregroundcolor Green -NoNewline
            $site | add-member -type noteproperty -name "Title_DestinationFound" -Value $foundSPOSite.Title -force
            $site | add-member -type noteproperty -name "SiteUrl_DestinationFound" -Value $foundSPOSite.URL -force
            $site | add-member -type noteproperty -name "TeamTitle_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "TeamSiteURL_DestinationFound" -Value $null.URL -force
        }
        elseif ($foundSPOSite = Get-SPOSITE -Identity $EHNSite) {
            Write-Host "."  -foregroundcolor yellow -NoNewline
            $site | add-member -type noteproperty -name "Title_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "SiteUrl_DestinationFound" -Value $null -force
            $site | add-member -type noteproperty -name "TeamTitle_DestinationFound" -Value $foundSPOSite.Title -force
            $site | add-member -type noteproperty -name "TeamSiteURL_DestinationFound" -Value $foundSPOSite.URL -force
        }
    }
    else {
        Write-Host "."  -foregroundcolor red -NoNewline
        $site | add-member -type noteproperty -name "Title_DestinationFound" -Value $null -force
        $site | add-member -type noteproperty -name "SiteUrl_DestinationFound" -Value $null -force
        $site | add-member -type noteproperty -name "TeamTitle_DestinationFound" -Value $null -force
        $site | add-member -type noteproperty -name "TeamSiteURL_DestinationFound" -Value $null -force
    }
    start-sleep -Milliseconds 100
}


#Update Matched Sites Details from TJU
$MatchedSites = import-csv
$TJUVSites = import-csv

$progressref = $MatchedSites.count
$progresscounter = 0
foreach ($site in $MatchedSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site $($Site.SiteURL)"

    if ($site | ?{$_.SiteURL_Destination}) {
        $matchedSite = @()
        #Check for EHN Labeled Site
        $TJUVSite = $site.SiteURL_Destination
        if ($matchedSite = $TJUVSites | ? {$_.SiteURL -eq $TJUVSite}) {
            Write-Host "."  -foregroundcolor Green -NoNewline
            $site | add-member -type noteproperty -name "SiteURL_Destination" -Value $matchedSite."SiteURL" -force
            $site | add-member -type noteproperty -name "Site_Destination" -Value $matchedSite.site -force
            $site | add-member -type noteproperty -name "RootWebTemplate_Destination" -Value $matchedSite."RootWebTemplate" -force
            $site | add-member -type noteproperty -name "Owner_Destination" -Value $matchedSite."OwnerPrincipalName" -force
            $site | add-member -type noteproperty -name "IsDeleted_Destination" -Value $matchedSite."IsDeleted" -force
            $site | add-member -type noteproperty -name "FileCount_Destination" -Value $matchedSite."FileCount" -force
            $site | add-member -type noteproperty -name "ActiveFileCount_Destination" -Value $matchedSite."ActiveFileCount" -force
            $site | add-member -type noteproperty -name "StorageUsed_Destination(Byte)" -Value $matchedSite."StorageUsed(Byte)" -force
            $site | add-member -type noteproperty -name "LastActivityDate_Destination" -Value $matchedSite."LastActivityDate" -force
            $site | add-member -type noteproperty -name "PageViewCount_Destination" -Value $matchedSite."PageViewCount" -force
            $site | add-member -type noteproperty -name "VisitedPageCount_Destination" -Value $matchedSite."VisitedPageCount" -force
        }
        else {
            Write-Host "."  -foregroundcolor Red -NoNewline
            $site | add-member -type noteproperty -name "SiteURL_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "Site_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "RootWebTemplate_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "Owner_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "IsDeleted_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "FileCount_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "ActiveFileCount_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "StorageUsed_Destination(Byte)" -Value $null -force
            $site | add-member -type noteproperty -name "LastActivityDate_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "PageViewCount_Destination" -Value $null -force
            $site | add-member -type noteproperty -name "VisitedPageCount_Destination" -Value $null -force
        }
    }
}


#Add SharePoint Sites Title from Tenant - EHN
$matchedSites = import-csv

$progressref = $matchedSites.count
$progresscounter = 0
foreach ($site in $matchedSites) {
    $siteLookup = $site.SiteURL
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site to Group $($siteLookup)"

    #check EHN Site URL Title
    if ($matchedSite = Get-SPOSITE -Identity $siteLookup) {
        Write-Host "."  -foregroundcolor green -NoNewline
        $site | add-member -type noteproperty -name "Title_Source" -Value $matchedSite.Title -force        
    }
    else {
        Write-Host "."  -foregroundcolor Red -NoNewline
        $site | add-member -type noteproperty -name "Title_Source" -Value $null -force   
    }
    start-sleep -Milliseconds 600
}

#Add SharePoint Sites Title and Current Usage from Tenant - TJUV
$matchedSites = import-csv

$progressref = $matchedSites.count
$progresscounter = 0
foreach ($site in $matchedSites) {
    $siteLookup = @()
    $siteLookup = $site."Site URL Destination"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site to Group $($siteLookup)"

    if ($siteLookup) {
        if ($matchedSite = Get-SPOSITE -Identity $siteLookup) {
            Write-Host "."  -foregroundcolor green -NoNewline
            $site | add-member -type noteproperty -name "Title_Destination2" -Value $matchedSite.Title -force
            $site | add-member -type noteproperty -name "StorageUsageCurrent_Destination(MB)2" -Value $matchedSite.StorageUsageCurrent -force  
        }
        else {
            Write-Host "."  -foregroundcolor Red -NoNewline
            $site | add-member -type noteproperty -name "Title_Destination2" -Value $null -force
            $site | add-member -type noteproperty -name "StorageUsageCurrent_Destination(MB)2" -Value $null -force  
        }
        start-sleep -Milliseconds 600
    }
    
}

#Match SharePoint Sites to Matched Teams Report
$matchedSites = import-csv
$matchedTeams = import-excel 

$progressref = $matchedSites.count
$progresscounter = 0
foreach ($site in $matchedSites) {
    $siteLookup = $site.SiteURL
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site to Group $($siteLookup)"

    #check EHN Site URL Title
    if ($matcheditem = $matchedTeams | ?{$_.SharePointSiteURL_Source -eq $siteLookup}) {
        Write-Host "."  -foregroundcolor green -NoNewline
        $site | add-member -type noteproperty -name "IsTeam" -Value $true -force
        $site | add-member -type noteproperty -name "TeamName_Source" -Value $matcheditem.TeamName_Source -force
        $site | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $matcheditem.PrimarySMTPAddress_Source -force
        $site | add-member -type noteproperty -name "TeamName_Destination" -Value $matcheditem.TeamName_Destination -force
        $site | add-member -type noteproperty -name "SharePointSiteURL_Destination" -Value $matcheditem.SharePointSiteURL_Destination -force  
    }
    else {
        Write-Host "."  -foregroundcolor Red -NoNewline
        $site | add-member -type noteproperty -name "IsTeam" -Value $False -force
        $site | add-member -type noteproperty -name "TeamName_Source" -Value $null -force
        $site | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $null -force
        $site | add-member -type noteproperty -name "TeamName_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "SharePointSiteURL_Destination" -Value $null -force 
    }
}


#Add SharePoint Sites Title from Groups
$matchedSites = import-csv
$MatchedGroups = import-csv 

$progressref = $matchedSites.count
$progresscounter = 0
foreach ($site in $matchedSites) {
    $ehnSite = $site.SiteURL

    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site to Group $($ehnSite)"

    #check EHN Site URL Title
    if ($matchedGroup = $MatchedGroups | ? {$_."SiteURL_Source" -eq $ehnSite}) {
        Write-Host "."  -foregroundcolor green -NoNewline
        $site | add-member -type noteproperty -name "DisplayName_Source" -Value $matchedGroup."DisplayName_Source" -force
        $site | add-member -type noteproperty -name "RecipientTypeDetails" -Value $matchedGroup."RecipientTypeDetails_Source" -force
        $site | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $matchedGroup."PrimarySMTPAddress_Source" -force
        $site | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedGroup."DisplayName_Destination" -force
        $site | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $matchedGroup."PrimarySmtpAddress_Destination" -force
        $site | add-member -type noteproperty -name "SiteURL_Destination2" -Value $matchedGroup."SiteURL_Destination2" -force
    }
    else {
        Write-Host "."  -foregroundcolor Red -NoNewline
        $site | add-member -type noteproperty -name "DisplayName_Source" -Value $null -force
        $site | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $null -force
        $site | add-member -type noteproperty -name "PrimarySMTPAddress_Source" -Value $null -force
        $site | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "RecipientTypeDetails_Source" -Value $null -force
        $site | add-member -type noteproperty -name "SiteURL_Destination2" -Value $null -force
    }
}

#Match SharePoint Sites to Team Sites
$matchedSites = import-csv

$progressref = $matchedSites.count
$progresscounter = 0
foreach ($site in $matchedSites) {
    $site | add-member -type noteproperty -name "SiteURL_Destination" -Value $matchedSite."SiteURL" -force
    $site | add-member -type noteproperty -name "Site_Destination" -Value $matchedSite.site -force
    $site | add-member -type noteproperty -name "RootWebTemplate_Destination" -Value $matchedSite."RootWebTemplate" -force
    $site | add-member -type noteproperty -name "Owner_Destination" -Value $matchedSite."OwnerPrincipalName" -force
}

#TJUV SIte URL
$progressref = $MatchedGroups.count
$progresscounter = 0
foreach ($group in $MatchedGroups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site to Group $($group.PrimarySMTPAddress_Destination)"

    $unifiedGroupSPURL = Get-UnifiedGroup -IncludeAllProperties -Identity $group.PrimarySMTPAddress_Destination -ErrorAction SilentlyContinue | select -expandProperty SharePointSiteUrl
    $group | add-member -type noteproperty -name "SiteURL_Destination2" -Value $unifiedGroupSPURL -force
}

#Update Matched Sites Details from ShareGate Report Matched Sites
$MatchedSites = import-csv
$ShareGateTeamsSites = import-excel

$progressref = $MatchedSites.count
$progresscounter = 0
foreach ($site in $MatchedSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site $($Site.SiteURL)"

    $matchedSite = @()
    #Check for EHN Labeled Site
    $EHNSite = $site.SiteURL + "/"
    if ($matchedSite = $ShareGateTeamsSites | ? {$_."Source site address" -eq $EHNSite}) {
        Write-Host "."  -foregroundcolor Green -NoNewline
        $site | add-member -type noteproperty -name "SourceTeamName" -Value $matchedSite."Source team name" -force
        $site | add-member -type noteproperty -name "SourceSiteName" -Value $matchedSite."Source site name" -force
        $site | add-member -type noteproperty -name "SourceURL3" -Value $matchedSite."Source site address" -force
        $site | add-member -type noteproperty -name "DestinationTeamName" -Value $matchedSite."Destination team name" -force
        $site | add-member -type noteproperty -name "DestinationSiteName" -Value $matchedSite."Destination site name" -force
        $site | add-member -type noteproperty -name "DestinationURL3" -Value $matchedSite."Destination site address" -force
    }
    else {
        Write-Host "."  -foregroundcolor Red -NoNewline
        $site | add-member -type noteproperty -name "SourceTeamName" -Value $null -force
        $site | add-member -type noteproperty -name "SourceSiteName" -Value $null -force
        $site | add-member -type noteproperty -name "SourceURL3" -Value $null -force
        $site | add-member -type noteproperty -name "DestinationTeamName" -Value $null -force
        $site | add-member -type noteproperty -name "DestinationSiteName" -Value $null -force
        $site | add-member -type noteproperty -name "DestinationURL3" -Value $null -force
    }
}

#Update Matched Sites Details from Matched Group
$MatchedSites = import-csv
$MatchedGroups = import-excel

$progressref = $MatchedSites.count
$progresscounter = 0
foreach ($site in $MatchedSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site $($Site.SiteURL)"

    $matchedobject = @()
    #Check for EHN Labeled Site
    $EHNSite = $site.SiteURL
    if ($matchedobject = $MatchedGroups | ? {$_."SiteURL_Source" -eq $EHNSite}) {
        Write-Host "."  -foregroundcolor Green -NoNewline
        $site | add-member -type noteproperty -name "SiteURL_GroupSource" -Value $matchedobject.SiteURL_Source -force
        $site | add-member -type noteproperty -name "SiteURL_GroupDestination" -Value $matchedobject.SiteURL_Destination2 -force
        $site | add-member -type noteproperty -name "GroupPrimarySMTPAddress_Source" -Value $matchedobject.PrimarySMTPAddress_Source -force
        $site | add-member -type noteproperty -name "GroupPrimarySMTPAddress_Destination" -Value $matchedobject.PrimarySmtpAddress_Destination -force
    }
    else {
        Write-Host "."  -foregroundcolor Red -NoNewline
        $site | add-member -type noteproperty -name "SiteURL_GroupSource" -Value $null -force
        $site | add-member -type noteproperty -name "SiteURL_GroupDestination" -Value $null -force
        $site | add-member -type noteproperty -name "GroupPrimarySMTPAddress_Source" -Value $null -force
        $site | add-member -type noteproperty -name "GroupPrimarySMTPAddress_Destination" -Value $null -force
    }
}

#Update Matched Sites Details from TJUV Sites Report
$MatchedSites = import-csv
$TJUVSitesDetails = import-csv

$progressref = $MatchedSites.count
$progresscounter = 0
foreach ($site in $MatchedSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site $($Site.SiteURL)"

    $matchedobject = @()
    #Check for EHN Labeled Site
    $TJUVSite = $site.SiteURL_Destination
    if ($matchedobject = $TJUVSitesDetails | ? {$_.URl -eq $TJUVSite}) {
        Write-Host "."  -foregroundcolor Green -NoNewline
        $site | add-member -type noteproperty -name "SiteName_Destination" -Value $matchedobject.SiteName -force
        $site | add-member -type noteproperty -name "SiteURL_Destination2" -Value $matchedobject.URL -force
        $site | add-member -type noteproperty -name "IsTeams_Destination" -Value $matchedobject.Teams -force
        $site | add-member -type noteproperty -name "StorageUsed_Destination(GB)" -Value $matchedobject."Storage used (GB)" -force
        $site | add-member -type noteproperty -name "FileCount_Destination" -Value $matchedobject.Files -force
        $site | add-member -type noteproperty -name "LastActivity_Destination" -Value $matchedobject."Last activity (UTC)" -force
    }
    else {
        Write-Host "."  -foregroundcolor Red -NoNewline
        $site | add-member -type noteproperty -name "SiteName_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "SiteURL_Destination2" -Value $null -force
        $site | add-member -type noteproperty -name "IsTeams_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "StorageUsed_Destination(GB)" -Value $null -force
        $site | add-member -type noteproperty -name "FileCount_Destination" -Value $null -force
        $site | add-member -type noteproperty -name "LastActivity_Destination" -Value $null -force
    }
}

#Update Matched Sites Details from TJUV Activity Sites Report
$MatchedSites = import-csv
$TJUVSitesDetails = import-csv

$progressref = $MatchedSites.count
$progresscounter = 0
foreach ($site in $MatchedSites) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching SharePoint Site $($Site.SiteURL)"

    $matchedobject = @()
    #Check for EHN Labeled Site
    $TJUVSite = $site.SiteURL_Destination
    if ($matchedobject = $TJUVSitesDetails | ? {$_.SiteURl -eq $TJUVSite}) {
        Write-Host "."  -foregroundcolor Green -NoNewline
        $site | add-member -type noteproperty -name "StorageUsed_Destination2(Byte)" -Value $matchedobject."StorageUsed(Byte)" -force
        $site | add-member -type noteproperty -name "FileCount_Destination2" -Value $matchedobject.FileCount -force
        $site | add-member -type noteproperty -name "LastActivity_Destination2" -Value $matchedobject.LastActivityDate -force
    }
    else {
        Write-Host "."  -foregroundcolor Red -NoNewline
        $site | add-member -type noteproperty -name "StorageUsed_Destination2(Byte)" -Value $null -force
        $site | add-member -type noteproperty -name "FileCount_Destination2" -Value $null -force
        $site | add-member -type noteproperty -name "LastActivity_Destination2" -Value $null -force
    }
}

#Gather EXOMailbox Stats - Mailbox and Archive Only (EHN)
$postmigrationWave5 = Import-Excel -WorksheetName "PostMigration_Wave5" -Path "C:\Users\aaron.medrano\Desktop\post migration\PostMigrationDetails-ALL_5-6-2020.xlsx"
$progressref = ($postmigrationWave5).count
$progresscounter = 0
foreach ($user in $postmigrationWave5) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName_Source)"
    
    $PrimarySMTPAddress = @()
    $TotalItemSize = @()
    $TotalDeletedItemSize = @()
    $CombinedItemSize = @()
    $ArchiveTotalItemSize = @()
    $ArchiveTotalDeletedItemSize = @()
    $ArchiveCombinedItemSize = @()

    if ($PrimarySMTPAddress = $user.PrimarySmtpAddress_Source) {
        #Pull MailboxStats and UserDetails
        $mbxCheck = Get-EXOMailbox $PrimarySMTPAddress -PropertySets archive
        $mbxStats = Get-EXOMailboxStatistics $PrimarySMTPAddress

        $TotalItemSize = ([math]::Round(($MBXStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
        $TotalDeletedItemSize = ([math]::Round(($MBXStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
        $CombinedItemSize = $TotalItemSize + $TotalDeletedItemSize

        #Create User Array
        $user | Add-Member -type NoteProperty -Name "Size_Source2-MB" -Value $TotalItemSize -force
        $user | Add-Member -type NoteProperty -Name "DeletedSize_Source2" -Value $TotalDeletedItemSize -force
        $user | Add-Member -Type NoteProperty -name "ItemCount_Source2" -Value $MBXStats.ItemCount -force
        $user | Add-Member -type NoteProperty -Name "DeletedItemCount_Source2" -Value $MBXStats.DeletedItemCount -force
        $user | Add-Member -type NoteProperty -Name "CombinedSize_Source2-MB" -Value $CombinedItemSize -force
        $user | Add-Member -type NoteProperty -Name "CombinedItemCount_Source2" -Value ($MBXStats.DeletedItemCount + $MBXStats.ItemCount) -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Source2" -Value $mbxCheck.ArchiveStatus -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveName_Source2" -Value $mbxCheck.ArchiveName.tostring() -force

        # Archive Mailbox Check
        if ($ArchiveStats = Get-EXOMailboxStatistics $PrimarySMTPAddress -Archive -ErrorAction silentlycontinue) {    
            
            #Archive Counts
            $ArchiveTotalItemSize = ([math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
            $ArchiveTotalDeletedItemSize = ([math]::Round(($ArchiveStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
            $ArchiveCombinedItemSize = $ArchiveTotalItemSize + $ArchiveTotalDeletedItemSize

            $user | add-member -type noteproperty -name "ArchiveSize_Source2" -Value $ArchiveTotalItemSize -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedSize_Source2" -Value $ArchiveTotalDeletedItemSize -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Source2" -Value $ArchiveStats.ItemCount -force 
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_Source2" -Value $ArchiveStats.DeletedItemCount -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedSize_Source2-MB" -Value $ArchiveCombinedItemSize -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_Source2" -Value ($ArchiveStats.DeletedItemCount + $ArchiveStats.ItemCount) -force
        }
        else {
            $user | add-member -type noteproperty -name "ArchiveSize_Source2" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveDeletedSize_Source2" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Source" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_Source2" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveCombinedSize_Source2-MB" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_Source2" -Value $null -force
        }
    }
    else {
        Write-Host "." -foregroundcolor red -nonewline
        $user | Add-Member -type NoteProperty -Name "Size_Source2-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "DeletedTotalSize_Source2" -Value $null -force
        $user | Add-Member -Type NoteProperty -name "ItemCount_Source2" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "DeletedItemCount_Source2" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "CombinedSize_Source2-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "CombinedItemCount_Source2" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Source2" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveName_Source2" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveSize_Source2" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveDeletedSize_Source2" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Source2" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_Source2" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveCombinedSize_Source2-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_Source2" -Value $null -force
    }
}
$postmigrationWave5 | Export-Excel -WorksheetName "PostMigration_Wave5Details" -Path "C:\Users\aaron.medrano\Desktop\post migration\PostMigrationDetails-ALL_5-6-2020.xlsx"


#Gather EXOMailbox Stats - Mailbox and Archive Only (TJUV)
$postmigrationWave2 = Import-Excel -WorksheetName "PostMigration_Wave2" -Path "C:\Users\aaron.medrano\Desktop\post migration\PostMigrationDetails-ALL_5-6-2020.xlsx"
$progressref = ($postmigrationWave2).count
$progresscounter = 0
foreach ($user in $postmigrationWave2) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName_Source)"
    
    $PrimarySMTPAddressDestination = @()
    $TotalItemSize = @()
    $TotalDeletedItemSize = @()
    $CombinedItemSize = @()
    $ArchiveTotalItemSize = @()
    $ArchiveTotalDeletedItemSize = @()
    $ArchiveCombinedItemSize = @()

    if ($PrimarySMTPAddressDestination = $user.PrimarySMTPAddress_Destination) {
        #Pull MailboxStats and UserDetails
        $mbxCheck = Get-EXOMailbox $PrimarySMTPAddressDestination -PropertySets archive
        $mbxStats = Get-EXOMailboxStatistics $PrimarySMTPAddressDestination

        $TotalItemSize = ([math]::Round(($MBXStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
        $TotalDeletedItemSize = ([math]::Round(($MBXStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
        $CombinedItemSize = $TotalItemSize + $TotalDeletedItemSize

        #Create User Array
        $user | Add-Member -type NoteProperty -Name "Size_Destination2-MB" -Value $TotalItemSize -force
        $user | Add-Member -type NoteProperty -Name "DeletedSize_Destination2" -Value $TotalDeletedItemSize -force
        $user | Add-Member -Type NoteProperty -name "ItemCount_Destination2" -Value $MBXStats.ItemCount -force
        $user | Add-Member -type NoteProperty -Name "DeletedItemCount_Destination2" -Value $MBXStats.DeletedItemCount -force
        $user | Add-Member -type NoteProperty -Name "CombinedSize_Destination2-MB" -Value $CombinedItemSize
        $user | Add-Member -type NoteProperty -Name "CombinedItemCount_Destination2" -Value ($MBXStats.DeletedItemCount + $MBXStats.ItemCount) -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination2" -Value $mbxCheck.ArchiveStatus -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveName_Destination2" -Value $mbxCheck.ArchiveName.tostring()

        # Archive Mailbox Check
        if ($ArchiveStats = Get-EXOMailboxStatistics $PrimarySMTPAddressDestination -Archive -ErrorAction silentlycontinue) {    
            
            #Archive Counts
            $ArchiveTotalItemSize = ([math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
            $ArchiveTotalDeletedItemSize = ([math]::Round(($ArchiveStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
            $ArchiveCombinedItemSize = $ArchiveTotalItemSize + $ArchiveTotalDeletedItemSize

            $user | add-member -type noteproperty -name "ArchiveSize_Destination2" -Value $ArchiveTotalItemSize -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedSize_Destination2" -Value $ArchiveTotalDeletedItemSize -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination2" -Value $ArchiveStats.ItemCount -force 
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_Destination2" -Value $ArchiveStats.DeletedItemCount -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedSize_Destination2-MB" -Value $ArchiveCombinedItemSize
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_Destination2" -Value ($ArchiveStats.DeletedItemCount + $ArchiveStats.ItemCount) -force
        }
        else {
            $user | add-member -type noteproperty -name "ArchiveSize_Destination2" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveDeletedSize_Destination2" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_Destination2" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_Destination2" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveCombinedSize_Destination2-MB" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_Destination2" -Value $null -force
        }
    }
    else {
        Write-Host "." -foregroundcolor red -nonewline
        $user | Add-Member -type NoteProperty -Name "Size_Destination2-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "DeletedTotalSize_Destination2" -Value $null -force
        $user | Add-Member -Type NoteProperty -name "ItemCount_Destination2" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "DeletedItemCount_Destination2" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "CombinedSize_Destination2-MB" -Value $null
        $user | Add-Member -type NoteProperty -Name "CombinedItemCount_Destination2" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination2" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveName_Destination2" -Value $mbxCheck.ArchiveName.tostring()
        $user | add-member -type noteproperty -name "ArchiveSize_Destination2" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveDeletedSize_Destination2" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination2" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_Destination2" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveCombinedSize_Destination2-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_Destination2" -Value $null -force
    }
}
$postmigrationWave2 | Export-Excel -MoveToStart -WorksheetName "PostMigration_Wave3Details" -Path "C:\Users\aaron.medrano\Desktop\post migration\PostMigrationDetails-ALL_5-6-2020.xlsx"

#Add Mailbox Migration Details - Post Migration Waves
$postmigrationWave5 = Import-Excel -WorksheetName "PostMigration_Wave5Details" -Path "C:\Users\aaron.medrano\Desktop\post migration\PostMigrationDetails-ALL_5-6-2020.xlsx"
$mailboxMigrations = Import-Excel -WorksheetName "EHN-TJUVMBX" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"
$archiveMBXMigrations = Import-Excel -WorksheetName "EHN-TJUVMBXArchive" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"

$progressref = $postmigrationWave5.count
$progresscounter = 0
foreach ($user in $postmigrationWave5) {
    $migrationDetails = @()
    $archiveMigrationDetails = @()
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Migration Details $($user.PrimarySmtpAddress_Source)"

    $user | add-member -type noteproperty -name "SizePercentage" -Value ($user."Size_Destination-MB"/$user."Size_Source-MB") -force
    $user | add-member -type noteproperty -name "ItemCountPercentage" -Value ($user.ItemCount_Destination/$user.ItemCount_Source) -force
    $user | add-member -type noteproperty -name "ArchiveSizePercent" -Value ($user."ArchiveSize_Destination-MB"/$user."ArchiveSize_Source-MB") -force
    $user | add-member -type noteproperty -name "ArchiveItemCountPercent" -Value ($user.ArchiveItemCount_Destination/$user.ArchiveItemCount_Source) -force

    ##Mailbox Migration Details
    if ($migrationDetails = $mailboxMigrations | ? {$_.SourceEmailAddress -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "Migration-LastStatus" -Value $migrationDetails.LastStatus -force
        $user | add-member -type noteproperty -name "Migration-MigrationType" -Value $migrationDetails.MigrationType -force
        $user | add-member -type noteproperty -name "Migration-SuccessSizeTotal(MB)" -Value $migrationDetails."SuccessSizeTotal(MB)" -force
        $user | add-member -type noteproperty -name "Migration-FailureSizeTotal(MB)" -Value $migrationDetails."FailureSizeTotal(MB)" -force
        $user | add-member -type noteproperty -name "Migration-SuccessCountTotal" -Value $migrationDetails.SuccessCountTotal -force
        $user | add-member -type noteproperty -name "Migration-FailureCountTotal" -Value $migrationDetails.FailureCountTotal -force
        $user | add-member -type noteproperty -name "Migration-SizePercent" -Value ($migrationDetails."SuccessSizeTotal(MB)"/$user."Size_Source-MB") -force
        $user | add-member -type noteproperty -name "Migration-ItemCountPercent" -Value ($migrationDetails.SuccessCountTotal/$user.ItemCount_Source) -force
    }
    else {
        $user | add-member -type noteproperty -name "Migration-LastStatus" -Value $null -force
        $user | add-member -type noteproperty -name "Migration-MigrationType" -Value $null -force
        $user | add-member -type noteproperty -name "Migration-SuccessSizeTotal(MB)" -Value $null -force
        $user | add-member -type noteproperty -name "Migration-FailureSizeTotal(MB)" -Value $null -force
        $user | add-member -type noteproperty -name "Migration-SuccessCountTotal" -Value $null -force
        $user | add-member -type noteproperty -name "Migration-FailureCountTotal" -Value $null -force
        $user | add-member -type noteproperty -name "Migration-SizePercent" -Value $null -force
        $user | add-member -type noteproperty -name "Migration-ItemCountPercent" -Value $null -force
    }
    ##Mailbox Archive Migration Details
    if ($archiveMigrationDetails = $archiveMBXMigrations | ? {$_.SourceEmailAddress -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "ArchiveMigration-LastStatus" -Value $archiveMigrationDetails.LastStatus -force
        $user | add-member -type noteproperty -name "ArchiveMigration-MigrationType" -Value $archiveMigrationDetails.MigrationType -force
        $user | add-member -type noteproperty -name "ArchiveMigration-SuccessSizeTotal(MB)" -Value $archiveMigrationDetails."SuccessSizeTotal(MB)" -force
        $user | add-member -type noteproperty -name "ArchiveMigration-FailureSizeTotal(MB)" -Value $archiveMigrationDetails."FailureSizeTotal(MB)" -force
        $user | add-member -type noteproperty -name "ArchiveMigration-SuccessCountTotal" -Value $archiveMigrationDetails.SuccessCountTotal -force
        $user | add-member -type noteproperty -name "ArchiveMigration-FailureCountTotal" -Value $archiveMigrationDetails.FailureCountTotal -force
        $user | add-member -type noteproperty -name "ArchiveMigration-SizePercent" -Value ($archiveMigrationDetails."SuccessSizeTotal(MB)"/$user."ArchiveSize_Source-MB") -force
        $user | add-member -type noteproperty -name "ArchiveMigration-ItemCountPercent" -Value ($archiveMigrationDetails.SuccessCountTotal/$user.ArchiveItemCount_Source) -force
        
    }
    else {
        $user | add-member -type noteproperty -name "ArchiveMigration-LastStatus" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveMigration-MigrationType" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveMigration-SuccessSizeTotal(MB)" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveMigration-FailureSizeTotal(MB)" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveMigration-SuccessCountTotal" -Value $null -force
        $user | add-member -type noteproperty -name "ArchivMigration-FailureCountTotal" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveMigration-SizePercent" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveMigration-ItemCountPercent" -Value $null -force
    }
}
$postmigrationWave5 | Export-Excel -MoveToEnd -WorksheetName "PostMigration_Wave5DetailsB" -Path "C:\Users\aaron.medrano\Desktop\post migration\PostMigrationDetails-ALL_5-6-2020.xlsx"


## Remove Archive Mailbox Migrations
$ProjectKeywords = "Archive"
$customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName -ErrorAction stop
$allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | ?{$_.name -like "*$ProjectKeywords*"} | sort name
$allMigMailboxes = @()
$allMigMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $allProjects.id -RetrieveAll -ea stop

$sourceNoArchiveMigrationCheck = Import-Csv "C:\Users\aaron.medrano\Desktop\post migration\PostMigration Activity\SourceNoArchiveMigrationCheck.csv"

$progresscounter = 0
$progressref = $sourceNoArchiveMigrationCheck.count
foreach ($migration in $sourceNoArchiveMigrationCheck) {
    $SourceAddress = $migration.SourceEmailAddress
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Remove Archive Migration for $($SourceAddress)"
    
    $migrationMailboxDetails = $allMigMailboxes | ?{$_.ExportEmailAddress -eq $SourceAddress}
    foreach ($migration in $migrationMailboxDetails) {
        Remove-MW_Mailbox -Ticket $mwTicket -Id $migration.id -force
    }
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
$AllRecipientDetails | Export-Excel "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Domain Cutover\Abaco-AllRecipients.xlsx"

#Match Recipients to Mailboxes
$allOldCompanyRecipients = Import-Excel -WorksheetName "Matched-AllRecipients"

$progressref = $allOldCompanyRecipients.count
$progresscounter = 0
foreach ($recipient in $allOldCompanyRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Recipient Details for $($recipient.PrimarySMTPAddress)"
    $matchedObject = @()
    if (!($recipient.DisplayName_Destination)) {
        if ($matchedObject = $allMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $recipient.PrimarySMTPAddress}) {
            $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedObject.DisplayName_Destination -force
            $recipient | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedObject.RecipientTypeDetails_Destination -force
            $recipient | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $matchedObject.PrimarySMTPAddress_Destination -force
        }
        if ($matchedObject = $MatchedGroups | ? {$_.PrimarySmtpAddress_Source -eq $recipient.PrimarySMTPAddress}) {
            $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedObject.DisplayName_Destination -force
            $recipient | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedObject.RecipientTypeDetails_Destination -force
            $recipient | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $matchedObject.PrimarySMTPAddress_Destination -force
        }
        if ($matchedObject = $MatchedTeams | ? {$_.PrimarySmtpAddress_Source -eq $recipient.PrimarySMTPAddress}) {
            $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedObject.TeamName_Destination -force
            $recipient | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value "Team" -force
            $recipient | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $matchedObject.PrimarySMTPAddress_Destination -force
        }
    }
}

#Check Inactive Mailboxes already exist
$RemainingMBXs = Import-Excel
$progressref = $RemainingMBXs.count
$progresscounter = 0
foreach ($mbx in $RemainingMBXs) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Recipient Details for $($mbx.SourceEmailAddress)"
    $matchedObject = @()
    if ($mbx.Project -like "*Inactive Archive") {
        if ($matchedObject = Get-Mailbox -Archive $mbx.SourceEmailAddress -ErrorAction SilentlyContinue) {
            $mbx | add-member -type noteproperty -name "PrimaryMailbox Already Exists" -Value $true -force
        }
        else {
            $mbx | add-member -type noteproperty -name "PrimaryMailbox Already Exists" -Value $False -force
        }
    }
    if ($mbx.Project -like "*Non-User Archive") {
        if ($matchedObject = Get-Mailbox -Archive $mbx.SourceEmailAddress -ErrorAction SilentlyContinue) {
            $mbx | add-member -type noteproperty -name "PrimaryMailbox Already Exists" -Value $true -force
        }
        else {
            $mbx | add-member -type noteproperty -name "PrimaryMailbox Already Exists" -Value $False -force
        }
    }
}

#Get Inactive Mailboxes Stats compared to mailbox stats - remaining mailboxes
$RemainingMBXs = Import-Excel -WorksheetName "MBX Review2" -Path

$progressref = $RemainingMBXs.count
$progresscounter = 0
foreach ($mbx in $RemainingMBXs) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Recipient Details for $($mbx.SourceEmailAddress)"
    $matchedObject = @()
    if ($mbx."PrimaryMailbox Already Exists" -eq $true) {
            $ArchivembxStat = Get-Mailbox $mbx.SourceEmailAddress -Archive |  Get-EXOMailboxStatistics
            $InactiveArchivembxStat = Get-Mailbox $mbx.SourceEmailAddress -Archive -InactiveMailboxOnly |  Get-EXOMailboxStatistics
            
            $mbx | add-member -type noteproperty -name "InactiveMbxSize" -Value ([math]::Round(($InactiveArchivembxStat.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3)) -force
            $mbx | add-member -type noteproperty -name "InactiveMbxItemCount" -Value $InactiveArchivembxStat.ItemCount -force
            $mbx | add-member -type noteproperty -name "ExistingMbxSize" -Value ([math]::Round(($ArchivembxStat.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3)) -force
            $mbx | add-member -type noteproperty -name "ExistingMbxItemCount" -Value $ArchivembxStat.ItemCount -force
    }
    else {
        $mbx | add-member -type noteproperty -name "InactiveMbxSize" -Value $Null -force
        $mbx | add-member -type noteproperty -name "InactiveMbxItemCount" -Value $Null -force
        $mbx | add-member -type noteproperty -name "ExistingMbxSize" -Value $Null -force
        $mbx | add-member -type noteproperty -name "ExistingMbxItemCount" -Value $Null -force
    }
}


#Match CustomAttribute7 for Unmatched Recipients to TJU
$unmatchedRecipients = Import-Excel -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-AllRecipients.xlsx" -WorksheetName UnMatchedRecipients

$progressref = $unmatchedRecipients.count
$progresscounter = 0
foreach ($recipient in $unmatchedRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching Recipient Details $($recipient.PrimarySMTPAddress)"
    $msoluserCheck = @()
    $addressCheck = $recipient.CustomAttribute7 + "@example.org"
    if ($msoluserCheck = Get-MsolUser -SearchString $recipient.CustomAttribute7){     
        $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluserCheck.DisplayName -force
        $recipient | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluserCheck.UserPrincipalName -force
    }
    elseif ($msoluserCheck = Get-MsolUser -SearchString $recipient.DisplayName.tostring()) {
        $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluserCheck.DisplayName -force
        $recipient | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluserCheck.UserPrincipalName -force
    }
    elseif ($msoluserCheck = Get-Recipient $addressCheck) {
        $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $msoluserCheck.DisplayName -force
        $recipient | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluserCheck.UserPrincipalName -force
    }
    else {
        $recipient | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $recipient | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
    }
}

#Match mailboxes to PostMigrationDetails report
$progressref = $ehnMailboxes.count
$progresscounter = 0
foreach ($mailbox in $ehnMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching Recipient Details $($mailbox.PrimarySMTPAddress)"
    
    if ($objectMatch = $PostMigrationDetails | ?{$_.PrimarySmtpAddress_Source -eq $mailbox.PrimarySmtpAddress}) {
        $mailbox | add-member -type noteproperty -name "Wave" -Value $objectMatch.Wave -force
        $mailbox | add-member -type noteproperty -name "DisplayName_Destination" -Value $objectMatch.DisplayName_Destination -force
        $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $objectMatch.UserPrincipalName_Destination -force
        $mailbox | add-member -type noteproperty -name "Migration-LastStatus" -Value $objectMatch."Migration-LastStatus" -force
        $mailbox | add-member -type noteproperty -name "ArchiveMigration-LastStatus" -Value $objectMatch."ArchiveMigration-LastStatus" -force
        $mailbox | add-member -type noteproperty -name "Note" -Value $objectMatch.Note -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "Wave" -Value $null -force
        $mailbox | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
        $mailbox | add-member -type noteproperty -name "Migration-LastStatus" -Value $null -force
        $mailbox | add-member -type noteproperty -name "ArchiveMigration-LastStatus" -Value $null -force
        $mailbox | add-member -type noteproperty -name "Note" -Value $null -force
    }

}

#Add Mailbox Migration Details - Post Migration
$EHNMailboxes = Import-CSV 
$mailboxMigrations = Import-Excel -WorksheetName "EHN-TJUVMBX" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"
$archiveMBXMigrations = Import-Excel -WorksheetName "EHN-TJUVMBXArchive" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"
$InactiveMBXMigrations = Import-Excel -WorksheetName "EHN-TJUVInactive" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"


$progressref = $EHNMailboxes.count
$progresscounter = 0
foreach ($user in $EHNMailboxes) {
    $migrationDetails = @()
    $archiveMigrationDetails = @()
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Migration Details $($user.PrimarySmtpAddress_Source)"

    if ($migrationDetails = $InactiveMBXMigrations | ? {$_.SourceEmailAddress -eq $user.PrimarySmtpAddress_Source}) {
        fads
    }
    ##Mailbox Migration Details
    if ($migrationDetails = $mailboxMigrations | ? {$_.SourceEmailAddress -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "Migration-LastStatus" -Value $migrationDetails.LastStatus -force
        $user | add-member -type noteproperty -name "Migration-MigrationType" -Value $migrationDetails.MigrationType -force
    }
    else {
        $user | add-member -type noteproperty -name "Migration-LastStatus" -Value $null -force
        $user | add-member -type noteproperty -name "Migration-MigrationType" -Value $null -force
    }
    ##Mailbox Archive Migration Details
    if ($archiveMigrationDetails = $archiveMBXMigrations | ? {$_.SourceEmailAddress -eq $user.PrimarySmtpAddress_Source}) {
        $user | add-member -type noteproperty -name "ArchiveMigration-LastStatus" -Value $archiveMigrationDetails.LastStatus -force
        $user | add-member -type noteproperty -name "ArchiveMigration-MigrationType" -Value $archiveMigrationDetails.MigrationType -force}
    else {
        $user | add-member -type noteproperty -name "ArchiveMigration-LastStatus" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveMigration-MigrationType" -Value $null -force
    }
}
$EHNMailboxes | Export-Excel -MoveToEnd -WorksheetName "EHNMailboxes-6292022" -Path "C:\Users\aaron.medrano\Desktop\post migration\PostMigrationDetails-ALL_5-6-2020.xlsx"


#Add Mailbox Migration Details - Post Migration
$EHNMailboxes = Import-CSV 
$mailboxMigrations = Import-Excel -WorksheetName "EHN-TJUVMBX" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"
$archiveMBXMigrations = Import-Excel -WorksheetName "EHN-TJUVMBXArchive" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"
$InactiveMBXMigrations = Import-Excel -WorksheetName "EHN-TJUVInactive" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"


$progressref = $EHNMailboxes.count
$progresscounter = 0
foreach ($user in $EHNMailboxes) {
    $migrationDetails = @()
    $archiveMigrationDetails = @()
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Migration Details $($user.UserPrincipalName)"

    $inactiveMailboxCheck = @()
    if ($user.UserPrincipalName -like "*-inactive@*") {
        $inactiveMailboxCheck = $inactiveMailboxes | ?{$_.EHNAddress -eq $user.PrimarySmtpAddress}
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $inactiveMailboxCheck.TJUVDisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $inactiveMailboxCheck.TJUVAddress -force
    }
}



#Add Mailbox Migration Details - Post Migration 2
$EHNMailboxes = $mailboxMigrations = Import-Excel -WorksheetName "InactiveMBXs" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU Remaining Mailbox Migrations - 6012022.xlsx"
$mailboxMigrations = Import-Excel -WorksheetName "InactiveMBXs" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU Remaining Mailbox Migrations - 6012022.xlsx"
$ArchiveInactiveMBXMigrations = $mailboxMigrations | ?{$_.Project -like "*archive*"}
$InactiveMBXMigrations = $mailboxMigrations | ?{$_.Project -notlike "*archive*"}


$progressref = $EHNMailboxes.count
$progresscounter = 0
foreach ($user in $EHNMailboxes) {
    $migrationDetails = @()
    $archiveMigrationDetails = @()
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Migration Details $($user.SourceEmailAddress)"

    ##Mailbox Migration Details
    if ($migrationDetails = $InactiveMBXMigrations | ? {$_.SourceEmailAddress -eq $user.SourceEmailAddress}) {
        $user | add-member -type noteproperty -name "Project2" -Value $migrationDetails.Project -force
        $user | add-member -type noteproperty -name "Migration-LastStatus" -Value $migrationDetails.LastStatus -force
        $user | add-member -type noteproperty -name "Migration-MigrationType" -Value $migrationDetails.MigrationType -force
        $user | add-member -type noteproperty -name "SuccessSizeTotal(MB)" -Value $migrationDetails."SuccessSizeTotal(MB)" -force
    }
    else {
        $user | add-member -type noteproperty -name "Project2" -Value $migrationDetails.Project -force
        $user | add-member -type noteproperty -name "Migration-LastStatus" -Value $migrationDetails.LastStatus -force
        $user | add-member -type noteproperty -name "Migration-MigrationType" -Value $migrationDetails.MigrationType -force
        $user | add-member -type noteproperty -name "SuccessSizeTotal(MB)" -Value $migrationDetails."SuccessSizeTotal(MB)" -force
    }
    ##Mailbox Archive Migration Details
    if ($archiveMigrationDetails = $ArchiveInactiveMBXMigrations | ? {$_.SourceEmailAddress -eq $user.SourceEmailAddress}) {
        $user | add-member -type noteproperty -name "ArchiveProject2" -Value $archiveMigrationDetails.Project -force
        $user | add-member -type noteproperty -name "ArchiveMigration-LastStatus" -Value $archiveMigrationDetails.LastStatus -force
        $user | add-member -type noteproperty -name "ArchiveMigration-MigrationType" -Value $archiveMigrationDetails.MigrationType -force
        $user | add-member -type noteproperty -name "ArchiveSuccessSizeTotal(MB)" -Value $archiveMigrationDetails."SuccessSizeTotal(MB)" -force
    }
    else {
        $user | add-member -type noteproperty -name "ArchiveProject2" -Value $archiveMigrationDetails.Project -force
        $user | add-member -type noteproperty -name "ArchiveMigration-LastStatus" -Value $archiveMigrationDetails.LastStatus -force
        $user | add-member -type noteproperty -name "ArchiveMigration-MigrationType" -Value $archiveMigrationDetails.MigrationType -force
        $user | add-member -type noteproperty -name "ArchiveSuccessSizeTotal(MB)" -Value $archiveMigrationDetails."SuccessSizeTotal(MB)" -force
    }
}
$EHNMailboxes | Export-Excel -MoveToEnd -WorksheetName "EHNMailboxes-6292022" -Path "C:\Users\aaron.medrano\Desktop\post migration\PostMigrationDetails-ALL_5-6-2020.xlsx"



#remove SIP, SPO, x500, onmicrosoft addresses, and non migrating domains  - matched recipients
$progressref = $matchedRecipients.count
$progresscounter = 0
foreach ($user in $matchedRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Alternate Address $($user.PrimarySMTPAddress)"

    $alternateAddresses = $user.EmailAddresses -split ","
    $newAlternateAddresses = $alternateAddresses | ?{$_ -notlike "*spo:*" -and $_ -notlike "*sip:*" -and $_ -notlike "*x500*" -and $_ -notlike "*@ehn.mail.onmicrosoft.com" -and $_ -notlike "*@gw.og-example.org" -and $_ -notlike "*@exchange.og-example.org" -and $_ -notlike "*@ehn.onmicrosoft.com" -and $_ -ne $user.PrimarySMTPAddress}
    $user | add-member -type noteproperty -name "Filtered_Addresses" -Value ($newAlternateAddresses -join ";") -force
}

#remove SIP, SPO, x500, onmicrosoft addresses, and non migrating domains  - matched recipients
$progressref = $unmatchedRecipients.count
$progresscounter = 0
foreach ($user in $unmatchedRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Alternate Address $($user.PrimarySMTPAddress)"

    $alternateAddresses = $user.EmailAddresses -split ","
    $newAlternateAddresses = $alternateAddresses | ?{$_ -notlike "*spo:*" -and $_ -notlike "*sip:*" -and $_ -notlike "*x500*" -and $_ -notlike "*@ehn.mail.onmicrosoft.com" -and $_ -notlike "*@gw.og-example.org" -and $_ -notlike "*@exchange.og-example.org" -and $_ -notlike "*@ehn.onmicrosoft.com" -and $_ -ne $user.PrimarySMTPAddress}
    $user | add-member -type noteproperty -name "Aliases" -Value ($newAlternateAddresses -join ";") -force
}

#Match Recipients to Mailboxes
$allOldCompanyRecipients = Import-Excel -WorksheetName "Matched-AllRecipients"

$progressref = $allOldCompanyRecipients.count
$progresscounter = 0
foreach ($recipient in $allOldCompanyRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Recipient Details for $($recipient.PrimarySMTPAddress_Destination)"
    $matchedObject = @()

    if ($recipientCheck = Get-Recipient $recipient.PrimarySMTPAddress_Destination) {
        $recipient | add-member -type noteproperty -name "DisplayName_Destination2" -Value $recipientCheck.DisplayName -force
        $recipient | add-member -type noteproperty -name "RecipientTypeDetails_Destination2" -Value $recipientCheck.RecipientTypeDetails -force
        $recipient | add-member -type noteproperty -name "PrimarySMTPAddress_Destination2" -Value $recipientCheck.PrimarySMTPAddress -force
    }
    else {
        $recipient | add-member -type noteproperty -name "DisplayName_Destination2" -Value $null -force
        $recipient | add-member -type noteproperty -name "RecipientTypeDetails_Destination2" -Value $null -force
        $recipient | add-member -type noteproperty -name "PrimarySMTPAddress_Destination2" -Value $null -force
    }
}

#Gather EXOMailbox Stats - Mailbox and Archive Only (TJUV)
$InactiveMailboxes = Import-Excel -WorksheetName "MigWizDetails" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU Inactive Mailbox Migrations.xlsx"
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
$InactiveMailboxDetails = @()
foreach ($user in $InactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.SourceEmailAddress)"
    
    $PrimarySMTPAddressDestination = $user.SourceEmailAddress
    $TotalItemSize = @()
    $ArchiveTotalItemSize = @()

    $mbxCheck = Get-EXOMailbox $PrimarySMTPAddressDestination

    $TotalItemSize = ([math]::Round(($MBXStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))

    #Create User Array
    $tmp = new-object PSObject
    $tmp | Add-Member -type NoteProperty -Name "Project" -Value $user.Project -force
    $tmp | Add-Member -type NoteProperty -Name "SourceEmailAddress" -Value $user.SourceEmailAddress -force
    $tmp | Add-Member -type NoteProperty -Name "DestinationEmailAddress" -Value $user.DestinationEmailAddress -force
    $tmp | Add-Member -type NoteProperty -Name "SourceSize-MB" -Value $TotalItemSize -force
    $tmp | Add-Member -Type NoteProperty -name "SuccessSizeTotal(MB)" -Value $user."SuccessSizeTotal(MB)" -force
    $tmp | Add-Member -Type NoteProperty -Name "SourceArchiveStatus" -Value $mbxCheck.ArchiveStatus -force
    $tmp | Add-Member -Type NoteProperty -Name "SourceArchiveName" -Value $mbxCheck.ArchiveName.tostring()

    # Archive Mailbox Check
    if ($ArchiveStats = Get-EXOMailboxStatistics $PrimarySMTPAddressDestination -Archive -ErrorAction silentlycontinue) {    
        
        #Archive Counts
        $ArchiveTotalItemSize = ([math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))

        $tmp | add-member -type noteproperty -name "SourceArchiveSize" -Value $ArchiveTotalItemSize -force
    }
    else {
        $tmp | add-member -type noteproperty -name "SourceArchiveSize" -Value $null -force
    }
    $InactiveMailboxDetails += $tmp
}
$InactiveMailboxDetails | Export-Excel -MoveToStart -WorksheetName "InactiveMailboxes-Source" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU Inactive Mailbox Migrations.xlsx"



#Gather EXOMailbox Stats - Mailbox and Archive Only (TJUV)
$InactiveMailboxes = Import-Excel -WorksheetName "Inactive Details" -Path "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU Inactive Mailbox Migrations.xlsx"
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
foreach ($user in $InactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user."Source Primary SMTP")"

    $mbxCheck = Get-EXOMailbox $user."Source Primary SMTP" -IncludeInactiveMailbox -Properties CustomAttribute7,WhenSoftDeleted,WhenCreated

    $user | Add-Member -type NoteProperty -Name "Source Display Name" -Value $mbxCheck.DisplayName -force
    $user | Add-Member -type NoteProperty -Name "Source UPN" -Value $mbxCheck.UserPrincipalName -force
    $user | Add-Member -type NoteProperty -Name "Source Custom Attribute 7" -Value $mbxCheck.CustomAttribute7 -force
    $user | Add-Member -type NoteProperty -Name "Source WhenSoftDeleted" -Value $mbxCheck.WhenSoftDeleted -force
    $user | Add-Member -type NoteProperty -Name "Source WhenCreated" -Value $mbxCheck.WhenCreated -force
}
$InactiveMailboxes | Export-Excel -MoveToStart -WorksheetName "Inactive Details2" -Path "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxes\EHN-TJU Inactive Mailbox Migrations.xlsx"

#Gather EXOMailbox Stats - Mailbox and Archive Only (TJUV)
$InactiveMailboxes = Import-Excel -WorksheetName "Inactive Details2" -Path "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxes\EHN-TJU Inactive Mailbox Migrations.xlsx"
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
foreach ($user in $InactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user."Target Primary SMTP")"

    $mbxCheck = Get-EXOMailbox $user."Target Primary SMTP" -IncludeInactiveMailbox -Properties CustomAttribute7,WhenSoftDeleted,WhenCreated

    $user | Add-Member -type NoteProperty -Name "Target Display Name" -Value $mbxCheck.DisplayName -force
    $user | Add-Member -type NoteProperty -Name "Target UPN" -Value $mbxCheck.UserPrincipalName -force
}
$InactiveMailboxes | Export-Excel -MoveToStart -WorksheetName "Inactive Details2" -Path "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxes\EHN-TJU Inactive Mailbox Migrations.xlsx"


#SharePoint Site to SharePoint Usage (TJUV)
$sharePointSitesRedo = import-csv "C:\Users\aaron.medrano\Desktop\post migration\SharePoint\RedoSharePointSites.csv"
$sharePointStats = Import-Csv "C:\Users\aaron.medrano\Desktop\post migration\SharePoint\SharePointSiteUsageDetail7_29_2022 7_42_43 PM.csv"
$progressref = ($sharePointSitesRedo).count
$progresscounter = 0

foreach ($site in $sharePointSitesRedo) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Teams/SharePoint Usage for $($site."Site URL Destination")"

    #Activity Check
    if ($usageMatch = $sharePointStats | ?{$_."Site URL" -eq $site."Site URL Destination"}) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $site | add-member -type noteproperty -name "File Count - Destination2" -Value $usageMatch."File Count" -Force
        $site | add-member -type noteproperty -name "Storage Used (Byte) - Destination2" -Value $usageMatch."Storage Used (Byte)" -Force
    }
    else {
        Write-Host "." -ForegroundColor Red -NoNewline
        $site | add-member -type noteproperty -name "File Count - Destination2" -Value $null -Force
        $site | add-member -type noteproperty -name "Storage Used (Byte) - Destination2" -Value $null -Force
    }
}

$sharePointSitesRedo | Export-CSV -NoTypeInformation -Encoding UTF8 "C:\Users\aaron.medrano\Desktop\post migration\SharePoint\RedoSharePointSites.csv"

#Get OneDrive URLs
$remainingOneDriveUsers = Import-Csv "C:\Users\aaron.medrano\Desktop\post migration\PostMigration Activity\remigrate onedrive users.csv"
Connect-SPOService -Url $DestinationURL -Credential $DestinationCredentials

$progressref = ($remainingOneDriveUsers).count
$progresscounter = 0
foreach ($user in $remainingOneDriveUsers) {
    $destinationUPN = $user.UserPrincipalName_Destination
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking $($destinationUPN)"

    #Activity Check
    if (!($user.DestinationOneDriveURL))
    {
        if ($mailboxCheck = Get-SPOSite -Filter "Owner -eq '$destinationUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true) {
            Write-Host "." -ForegroundColor Green -NoNewline
            $user | add-member -type noteproperty -name "DestinationOneDriveURL" -Value $mailboxCheck.URL -Force
        }
        else {
            Write-Host "." -ForegroundColor Red -NoNewline
            $user | add-member -type noteproperty -name "DestinationOneDriveURL" -Value $null -Force
        }
    }   
}

$remainingOneDriveUsers | Export-CSV -NoTypeInformation -Encoding UTF8 "C:\Users\aaron.medrano\Desktop\post migration\PostMigration Activity\remigrate onedrive users.csv"


#Domain Cutover - Update UPN/PrimarySMTP Address
function Start-TenantDomainCutover {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$false,HelpMessage='AllCurrentUsers?')] [string] $AllCurrentUsers,
        [Parameter(Mandatory=$false,HelpMessage="Update UPN?")] [switch]$SwitchUPN,
        [Parameter(Mandatory=$false,HelpMessage="Update PrimarySMTPAddress")] [switch]$UpdatePrimarySMTP
    )
    Connect-MsolService
    Connect-ExchangeOnline
    #Gather OnMicrosoft Domain
    $onMicrosoftDomain = (Get-MsolDomain | ? {$_.name -like "*.onmicrosoft.com" -and $_.name -notlike "*.mail.onmicrosoft.com"}).Name
    #Gather All Users
    if ($AllCurrentUsers) {
        $allMSOLUsers = Get-MsolUser -All
    }
    elseif ($InputEXCELFilePath) {
        $allMSOLUsers = Import-CSV $InputCSVFilePath
    }
    elseif ($InputEXCELFilePath) {
        $allMSOLUsers = Import-Excel $InputCSVFilePath
    }
    
    #Progress Bar 1A
    $progressref = ($allMSOLUsers).count
    $progresscounter = 0
    $AllErrors = @()
    foreach ($msolUser in $allMsolUsers| sort userprincipalname) {
        #Progress Bar 1B
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Domain Cutover for $($msolUser.displayname) to $($onMicrosoftDomain)"
        #Variable - NewUPN
        $newUPN = ($msolUser.UserPrincipalName -split "@")[0] + "@" + $onMicrosoftDomain
        Write-Host "Update $($msolUser.displayname)" -ForegroundColor Cyan -NoNewline
        Write-Host ".." -ForegroundColor Yellow -NoNewline
        
        #Update UPN
        if ($SwitchUPN) {
            Write-Host "UPNUpdate.." -NoNewline -foregroundcolor DarkCyan
            if ($msolUser.UserPrincipalName -like "*$onMicrosoftDomain") {
                Write-Host "Skipping." -ForegroundColor Yellow -NoNewline
            }
            else {
                try {
                    #Set-MsolUserPrincipalName -UserPrincipalName $msolUser.UserPrincipalName -NewUserPrincipalName $newUPN
                    Write-host "$($newupn). " -foregroundcolor Green -NoNewline
                }
                catch {
                    Write-Error "$($newupn). " -foregroundcolor Red -NoNewline
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdateUPN" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $msolUser.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "NewUPN" -Value $newUPN -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrorsPerms += $currenterror           
                    continue
                }
            }
        }

        #Update PrimarySMTPAddress
        if ($UpdatePrimarySMTP) {
            Write-Host "PrimarySMTPAddressUpdate.." -NoNewline -foregroundcolor DarkCyan
            if (Get-Recipient $msolUser.UserPrincipalName -erroraction SilentlyContinue) {
                if ((Get-Recipient $msolUser.UserPrincipalName).primarysmtpaddress -like "*$onMicrosoftDomain") {
                    Write-Host "Skipping." -ForegroundColor Yellow -NoNewline
                }
                else {
                    try {
                        #Set-Mailbox -Identity $msolUser.UserPrincipalName -PrimarySMTPAddress $newUPN
                        Write-host "Updated." -foregroundcolor Green -NoNewline
                    }
                    catch {
                        Write-Error "Failed to Update." -foregroundcolor Red -NoNewline
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdatePrimarySMTPAddress" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $msolUser.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "NewUPN" -Value $newUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
            }
            else {
                Write-host "No Recipient Found." -foregroundcolor Red -NoNewline
            }
        }
        #Completed User
        Write-Host "Done" -ForegroundColor Green
    }
}

#Domain Cutover - Update UPN/PrimarySMTP Address
function Add-TenantDomainCutover {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$false,HelpMessage='AllCurrentUsers?')] [string] $AllCurrentUsers,
        [Parameter(Mandatory=$false,HelpMessage="Update UPN?")] [switch]$SwitchUPN,
        [Parameter(Mandatory=$false,HelpMessage="Update Aliases")] [switch]$AddAlias
    )
    Connect-MsolService
    Connect-ExchangeOnline


    #Gather All Users
    if ($InputCSVFilePath) {
        $allMatchedRecipients = Import-CSV $InputCSVFilePath
    }
    elseif ($InputEXCELFilePath) {
        $allMatchedRecipients = Import-Excel $InputCSVFilePath
    }
    
    #Progress Bar 1A
    $progressref = ($allMatchedRecipients).count
    $progresscounter = 0
    $AllErrors = @()
    $newDomain = "og-example.org"
    $notFoundUsers = @()
    foreach ($matchedRecipient in $allMatchedRecipients| sort userprincipalname) {
        #Progress Bar 1B
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Migrating Domain Aliases to $($matchedRecipient.DisplayName_Destination)"

        #Variable - NewUPN
        $newAddress = ($matchedRecipient.PrimarySMTPAddress -split "@")[0] + "@" + $newDomain
        Write-Host "Update $($matchedRecipient.displayname)" -ForegroundColor Cyan -NoNewline
        Write-Host ".." -ForegroundColor Yellow -NoNewline

        if (!($recipientCheck = Get-Recipient $matchedRecipient.PrimarySMTPAddress -ErrorAction SilentlyContinue)) {
            if ($matchedRecipient.RecipientTypeDetails_Destination -eq "GroupMailbox" -or $matchedRecipient.RecipientTypeDetails_Destination -eq "Team") {
                try {
                    Set-UnifiedGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                    if ($newAlias = $matchedRecipient.Aliases) {
                        Set-UnifiedGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    }
                    Write-host "$($newAddress). " -foregroundcolor Green -NoNewline
                }
                catch {
                    Write-Error "$($newAddress). " -foregroundcolor Red -NoNewline
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
                
            }
            elseif ($matchedRecipient.RecipientTypeDetails_Destination -eq "DynamicDistributionGroup") {
                try {
                    Set-DynamicDistributionGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                    if ($newAlias = $matchedRecipient.Aliases) {
                        Set-DynamicDistributionGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    }
                }
                catch {
                    Write-Error "$($newAddress). " -foregroundcolor Red -NoNewline
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($matchedRecipient.RecipientTypeDetails_Destination -eq "MailUniversalDistributionGroup") {
                try {
                    Set-DistributionGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                    if ($newAlias = $matchedRecipient.Aliases) {
                        Set-DistributionGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    }
                }
                catch {
                    Write-Error "$($newAddress). " -foregroundcolor Red -NoNewline
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($matchedRecipient.RecipientTypeDetails_Destination -like "*Mailbox") {
                try {
                    Set-Mailbox Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                    if ($newAlias = $matchedRecipient.Aliases) {
                        Set-Mailbox -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    }

                }
                catch {
                    Write-Error "$($newAddress). " -foregroundcolor Red -NoNewline
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            else {
                try {
                    Set-Mailbox Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                    if ($newAlias = $matchedRecipient.Aliases) {
                        Set-Mailbox -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    }
                }
                catch {
                    Write-Error "$($newAddress). " -foregroundcolor Red -NoNewline
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
        }        
        else {
            Write-Host "Recipient Not Found." -ForegroundColor Red -NoNewline
            $notFoundUsers += $matchedRecipient
        }

        Write-Host "Done" -ForegroundColor Green
    }
}


#Add Current DirSync Status
$OldCompanyUsers = Import-Csv "C:\Users\amedrano\Arraya Solutions\Thomas NewCompany External - OldCompany to NewCompany Migration\Domain Cutover\All_OldCompanyUsers_10_18_2022.csv"
Connect-AzureAD

$progressref = ($OldCompanyUsers).count
$progresscounter = 0
foreach ($user in $OldCompanyUsers) {
    $destinationUPN = $user.UserPrincipalName
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking $($destinationUPN)"
    
    $DirSyncCheck = Get-AzureADUser -ObjectId $destinationUPN | select -ExpandProperty DirSyncEnabled
    $user | add-member -type noteproperty -name "DirSyncStatus" -Value $DirSyncCheck -Force
}