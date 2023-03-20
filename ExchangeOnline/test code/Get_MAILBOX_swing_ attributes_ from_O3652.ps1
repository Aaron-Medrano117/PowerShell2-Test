#Gather Mailbox Stats

function Import-ExchangeMsOnlineAzureModule() {
    #Exchange Online Module
    if ((Get-Module -Name "ExchangeOnlineManagement") -ne $null) {
        return;
    }
    else {
        if ((Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue) -ne $null) {
            try {
                Write-Host "Connecting to ExchangeOnline ... " -nonewline -foregroundcolor cyan
                Connect-ExchangeOnline
                Write-Host "done" -foregroundcolor green
            }
            catch {
                Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
            }
        }
        else {
            try {
                Install-Module -Name ExchangeOnlineManagement 
            }
            catch {
                Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
            }
        }
    }
    #Microsoft Online Module
    if ((Get-Module -Name "MSOnline") -ne $null) {
        return;
    }
    else {
        if ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null) {
            try {
                Write-Host "Connecting to MSOnline ... " -nonewline -foregroundcolor cyan
                Connect-MsolService
                Write-Host "done" -foregroundcolor green
            }
            catch {
                Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
            }
        }
        else {
            try {
                Install-Module -Name MSOnline 
            }
            catch {
                Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
            }
        }
    }
    #Azure AD Module
    if ((Get-Module -Name "AzureAD") -ne $null) {
        return;
    }
    else {
        if ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -ne $null) {
            try {
                Write-Host "Connecting to AzureAD ... " -nonewline -foregroundcolor cyan
                Connect-AzureAD
                Write-Host "AzureAD" -foregroundcolor green
            }
            catch {
                Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
            }
        }
        else {
            try {
                Install-Module -Name AzureAD 
            }
            catch {
                Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
            }
        }
    }
    
}
Import-ExchangeMsOnlineAzureModule

#Connect To Exchange Online and MsOnline
Write-Host "Connecting to MsOnline ... " -nonewline -foregroundcolor cyan
Connect-MsolService



Connect-ExchangeOnline
Write-Host "done" -foregroundcolor green

#Gather Mailbox Stats
$sourceMailboxes = Get-ExoMailbox -ResultSize Unlimited -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -PropertySets archive, delivery, minimum
$sourceMailboxStats = @()
$progressref = ($sourceMailboxes).count
$progresscounter = 0
foreach ($user in $sourceMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    
    #Pull MailboxStats and UserDetails
    $mbxStats = Get-EXOMailboxStatistics $user.PrimarySMTPAddress
    $msoluser = Get-MsolUser -UserPrincipalName $user.UserPrincipalName | select IsLicensed, Licenses, LicenseAssignmentDetails, Department
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    #Create User Array
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName -force
    $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $user.userprincipalname -force
    $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msoluser.IsLicensed -force
    $currentuser | add-member -type noteproperty -name "Licenses" -Value ($msoluser.Licenses.AccountSkuID -join ",") -force
    $currentuser | add-member -type noteproperty -name "License-DisabledArray" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ",") -force
    $currentuser | add-member -type noteproperty -name "Department" -Value $msoluser.Department -force
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $user.RecipientTypeDetails -force
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.PrimarySmtpAddress -force
    $currentuser | add-member -type noteproperty -name "Alias" -Value $user.alias -force
    $currentuser | add-member -type noteproperty -name "WhenCreated" -Value $user.WhenCreated -force
    $currentuser | add-member -type noteproperty -name "LastLogonTime" -Value $mbxStats.LastLogonTime -force
    $currentuser | add-member -type noteproperty -name "LastInteractionTime" -Value $mbxStats.LastInteractionTime -force
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ",") -force
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $user.legacyexchangedn) -force
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $user.HiddenFromAddressListsEnabled -force
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $user.DeliverToMailboxAndForward -force
    $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $user.ForwardingAddress -force
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $user.ForwardingSmtpAddress -force
    $currentuser | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize -force
    $currentuser | Add-Member -type NoteProperty -Name "MBXSize-GB" -Value ([math]::Round(($MBXStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1GB,3))
    $currentuser | Add-Member -type NoteProperty -Name "TotalDeletedItemSize-GB" -Value ([math]::Round(($MBXStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1GB,3))
    $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount -force

    #Pull Send on Behalf
    if ($grantSendOnBehalf = $user.GrantSendOnBehalfTo) {
        $grantSendOnBehalfPerms = @()
        foreach ($perm in $grantSendOnBehalf) {
            $mailboxCheck = (Get-Mailbox $perm).DisplayName
            $grantSendOnBehalfPerms += $mailboxCheck
        }
        $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($grantSendOnBehalfPerms -join ",")
    }
    else {
        $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value $null
    }
    

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.primarysmtpaddress | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value ($mbxPermissions.user -join ",") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms" -Value $null -force
    }
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.PrimarySMTPAddress | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ",") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "SendAsPerms" -Value $null  -force
    }
    # Archive Mailbox Check
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus -force
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveName" -Value $user.ArchiveName.ToString() -force
    if ($user.ArchiveStatus -eq "Active") {
        $ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount     
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $ArchiveStats.TotalItemSize.Value -force 
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $ArchiveStats.ItemCount -force
    }
    else {
        $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $null -force 
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $null -force
    }
    $sourceMailboxStats += $currentuser
}
$sourceMailboxStats | Export-Excel -Path 'C:\Users\amedrano\Arraya Solutions\Spectra - Internal - 1777 - Spectra to OVG- Tenant-to-Tenant Migration\Customer Mailbox List\AllSpectraMailboxes_12-5-2022.xlsx'

# Match Mailboxes and add to same spreadsheet. Check based on NEWUPN, DisplayName, NEWPRIMARYSMTP
$sourceMailboxes = Import-Excel
$progressref = ($sourceMailboxes).count
$progresscounter = 0
$newDomain = Read-Host "What is the new domain migrating to?"

foreach ($user in $sourceMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    $addressSplit = $user.UserPrincipalName -split "@"
    $NewUPN = $addressSplit[0] + "@" + $newDomain
    $SMTPaddressSplit = $user.$PrimarySMTPAddress -split "@"
    $newPrimarySMTPAddress = $SMTPaddressSplit[0] + "@" + $newDomain
 
    #Old Address Mail User Check
    if ($mailUserCheck = Get-MailUser $user.PrimarySMTPAddress -ErrorAction SilentlyContinue) {
        $msoluser = Get-MsolUser -UserPrincipalName $mailUserCheck.userPrincipalName
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
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ",") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ",") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ",") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $recipient.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | add-member -type noteproperty -name "LastInteractionTime_Destination" -Value $mbxStats.LastInteractionTime -force
        $user | add-member -type noteproperty -name "WhenCreated_Destination" -Value $mbxStats.WhenCreated -force
        
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
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ",") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ",") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ",") -force
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
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ",") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ",") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ",") -force
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
    elseif ($msoluser = Get-Msoluser -SearchString $user.DisplayName -ErrorAction SilentlyContinue) {
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
        $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluser.Licenses.AccountSkuID -join ",") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_Destination" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ",") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_Destination" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_Destination" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_Destination" -Value ($EmailAddresses -join ",") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_Destination" -Value $recipient.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_Destination" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_Destination" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_Destination" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $mbxStats.LastLogonTime -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_Destination" -Value $user.ArchiveStatus -force
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
        Write-Host "  Unable to find user for $($user.DisplayName)" -ForegroundColor Red
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
        $user | add-member -type noteproperty -name "LastLogonTime_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveSize_Destination" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_Destination" -Value $null -force
    }    
}