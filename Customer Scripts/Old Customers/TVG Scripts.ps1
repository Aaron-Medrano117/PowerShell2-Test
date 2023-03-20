$TVGSharedMailboxes = @()
$MultipleAliases = @()

foreach ($user in $sharedtvgmailboxes)
{
        Write-Host "Gathering Mailbox Stats for $($user.DisplayName) .." -ForegroundColor Cyan -NoNewline
        $AliasCheck = Get-Recipient $user.Alias
        
        if ($AliasCheck.count -gt 1)
        {
            Write-Host "Multiple Aliases found for User" -ForegroundColor Red
            $MultipleAliases += $AliasCheck
        }

        else
        {
            $MBXStats = Get-MailboxStatistics $user.primarysmtpaddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
            $addresses = $user | select -ExpandProperty EmailAddresses
            $MSOLUser = Get-MsolUser -userprincipalname $user.userprincipalname
            $FanDuelGroupAddressSplit = $user.primarysmtpaddress -split "@"
            $FanDuelGroupAddress = $FanDuelGroupAddressSplit[0] + "@fanduelgroup.com"

            $currentuser = new-object PSObject
            
            $currentuser | add-member -type noteproperty -name "DisplayName" -Value $msoluser.DisplayName
            $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $msoluser.userprincipalname
            $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msoluser.IsLicensed
            $currentuser | add-member -type noteproperty -name "City" -Value $msoluser.City
            $currentuser | add-member -type noteproperty -name "Country" -Value $msoluser.Country
            $currentuser | add-member -type noteproperty -name "Department" -Value $msoluser.Department
            $currentuser | add-member -type noteproperty -name "Fax" -Value $msoluser.Fax
            $currentuser | add-member -type noteproperty -name "FirstName" -Value $msoluser.FirstName
            $currentuser | add-member -type noteproperty -name "LastName" -Value $msoluser.LastName
            $currentuser | add-member -type noteproperty -name "MobilePhone" -Value $msoluser.MobilePhone
            $currentuser | add-member -type noteproperty -name "Office" -Value $msoluser.Office
            $currentuser | add-member -type noteproperty -name "PhoneNumber" -Value $msoluser.PhoneNumber
            $currentuser | add-member -type noteproperty -name "PostalCode" -Value $msoluser.PostalCode
            $currentuser | add-member -type noteproperty -name "State" -Value $msoluser.State
            $currentuser | add-member -type noteproperty -name "StreetAddress" -Value $msoluser.StreetAddress
            $currentuser | add-member -type noteproperty -name "Title" -Value $msoluser.Title
            
            $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.PrimarySmtpAddress
            $currentuser | add-member -type noteproperty -name "WhenCreated" -Value $user.WhenCreated
            $currentuser | add-member -type noteproperty -name "FanDuelGroupAddress" -Value $FanDuelGroupAddress 
            $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($addresses -join ",")
            $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $user.legacyexchangedn)
            $currentuser | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value ($user.AcceptMessagesOnlyFrom -join ",")
            $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($user.GrantSendOnBehalfTo -join ",")
            $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $user.HiddenFromAddressListsEnabled
            $currentuser | add-member -type noteproperty -name "RejectMessagesFrom" -Value ($user.RejectMessagesFrom -join ",")
            $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $user.DeliverToMailboxAndForward
            $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $user.ForwardingAddress
            $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $user.ForwardingSmtpAddress
            $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $user.RecipientTypeDetails
            $currentuser | add-member -type noteproperty -name "Alias" -Value $user.alias
            $currentuser | add-member -type noteproperty -name "ExchangeGuid" -Value $user.ExchangeGuid
            $currentuser | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize
            $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount
            $currentuser | Add-Member -Type NoteProperty -Name "ArchiveGUID" -Value $user.ArchiveGuid
            $currentuser | add-member -type noteproperty -name "ArchiveState" -Value $user.ArchiveState
            $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus

            if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount)
            {
                Write-Host "Archive found ..." -ForegroundColor green -NoNewline
                
                $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $ArchiveStats.TotalItemSize.Value
                $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $ArchiveStats.ItemCount
            }

            else
            {
                Write-Host "no Archive found ..." -ForegroundColor Red -NoNewline
            }  
            
            Write-Host "done .." -ForegroundColor Green -NoNewline
            
        $TVGSharedMailboxes += $currentuser
        }
}


# Match Mailboxes

$sharedtvgmailboxes =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $TVGSharedMailboxes) {
    $DisplayName = $user.DisplayName
    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline
    if ($mailboxcheck = Get-Mailbox $DisplayName -ea silentlycontinue) {
        
        if ($mailboxcheck.count -gt 1)
        {
            Write-Host "Multiple Mailboxes found ..." -ForegroundColor Yellow -NoNewline
            foreach ($mbx in $mailboxcheck)
            {
                $foundusers += $mbx
                $currentuser = new-object PSObject

                #Get Mailbox Details
                $currentuser | add-member -type noteproperty -name "MatchingUser" -Value $user.DisplayName
                $currentuser | add-member -type noteproperty -name "MatchingPrimarySMTPAddress" -Value $user.PrimarySMTPAddress
                $currentuser | add-member -type noteproperty -name "ExistsOnO365" -Value $true
                $currentuser | add-member -type noteproperty -name "DisplayName" -Value $mbx.DisplayName
                $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $mbx.userprincipalname
                $currentuser | add-member -type noteproperty -name "RecipientType" -Value $mbx.RecipientTypeDetails
                $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $mbx.PrimarySMTPAddress
                


                # Get MSOL User Information
                $Msoluser = get-msoluser -userprincipalname $mbx.UserPrincipalName | select IsLicensed
                Write-Host "MSOLUser Details Gathered" -ForegroundColor Green
                $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $mbx.IsLicensed
            }
        }
        else
        {
            $foundusers += $user
            Write-Host "found ..." -ForegroundColor Green -NoNewline
            
            $currentuser = new-object PSObject

            #Get Mailbox Details
            $currentuser | add-member -type noteproperty -name "MatchingUser" -Value $user.DisplayName
            $currentuser | add-member -type noteproperty -name "MatchingPrimarySMTPAddress" -Value $user.PrimarySMTPAddress
            $currentuser | add-member -type noteproperty -name "ExistsOnO365" -Value $true
            $currentuser | add-member -type noteproperty -name "DisplayName" -Value $mailboxcheck.DisplayName
            $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $mailboxcheck.userprincipalname
            $currentuser | add-member -type noteproperty -name "RecipientType" -Value $mailboxcheck.RecipientTypeDetails
            $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $mailboxcheck.PrimarySMTPAddress


            # Get MSOL User Information
            $Msoluser = get-msoluser -userprincipalname $mailboxcheck.UserPrincipalName | select IsLicensed
            Write-Host "MSOLUser Details Gathered" -ForegroundColor Green
            $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $mailboxcheck.IsLicensed
        }
    }
    else
    {
        $notfoundusers += $user
        Write-Host "not found" -ForegroundColor red
        $currentuser = new-object PSObject

        #Get Mailbox Details
        $currentuser | add-member -type noteproperty -name "MatchingUser" -Value $user.DisplayName
        $currentuser | add-member -type noteproperty -name "MatchingPrimarySMTPAddress" -Value $user.PrimarySMTPAddress
        $currentuser | add-member -type noteproperty -name "ExistsOnO365" -Value $false
        $currentuser | add-member -type noteproperty -name "DisplayName" -Value $null
        $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $null
        $currentuser | add-member -type noteproperty -name "RecipientType" -Value $null
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $null
        
    }

    $sharedtvgmailboxes += $currentuser
}

#Create Share Mailbox
foreach ($sharedmailbox in $TVGSharedMailboxes | ? {$_.ExistsOnO365 -eq $false})
{
    #create the shared mailbox
    $FanDuelGroupAddress = $sharedmailbox.FanDuelGroupAddress.replace("@fanduelgroup.com","@fanduel.com")
    Write-host "Adding $($sharedmailbox.DisplayName) .." -foregroundcolor cyan -nonewline
    New-Mailbox -shared -PrimarySMTPAddress $FanDuelGroupAddress -Alias $sharedmailbox.alias -name $sharedmailbox.DisplayName -displayname $sharedmailbox.DisplayName
    Write-host "done " -foregroundcolor green
}

## Set PrimarySMTPAddress
foreach ($sharedmailbox in $TVGSharedMailboxes)
{
    Write-Host "Setting PrimarySMTPAddress for $($sharedmailbox.DisplayName) to $($sharedmailbox.MatchingPrimarySMTPAddress)" -foregroundcolor cyan -nonewline
    $mailboxcheck = Get-Mailbox $sharedmailbox.DisplayName | select -ExpandProperty PrimarySMTPAddress
    if (!($mailboxcheck -eq $sharedmailbox.MatchingPrimarySMTPAddres))
    Set-Mailbox $sharedmailbox.DisplayName -WindowsEmailAddress $sharedmailbox.MatchingPrimarySMTPAddress -whatif
    Write-Host "done" -foregroundcolor green
    else {
        Write-Host "don't need to update." -foregroundcolor yellow
    }
}

## Set UPN
foreach ($sharedmailbox in $sharedtvgmailboxes)
{
    Write-Host "Checking $($sharedmailbox.DisplayName) .. " -foregroundcolor cyan -nonewline
    if ($sharedmailbox.UserPrincipalName -like "*fanduelgroup.onmicrosoft.com")
    {   
        $UPNUpdate = $sharedmailbox.UserPrincipalName.replace("fanduelgroup.onmicrosoft.com","fanduel.com")
        Write-Host "Setting UPN to $($sharedmailbox.UPNUpdate)" -foregroundcolor cyan -nonewline
        Set-MsolUserPrincipalName -UserPrincipalName $sharedmailbox.UserPrincipalName -NewUserPrincipalName $UPNUpdate #-whatif
        Write-Host "done" -foregroundcolor green
    }
    else {
        Write-Host "don't need to update." -foregroundcolor yellow
    }
}

$permsList = @()
foreach ($mbx in $sharedtvgmailboxes)
{
	Write-Host "$($mbx.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	[array]$perms = Get-MailboxPermission $mbx.DisplayName | Where {!$_.IsInherited -and $_.AccessRights -like "*FullAccess*"}
	$perms = $perms | Where {$_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-*"}

	if ($perms)
	{
		foreach ($perm in $perms)
		{
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select Mailbox, UserWithFullAccess
			$tmp.Mailbox = $mbx.PrimarySmtpAddress.ToString()
			$tmp.UserWithFullAccess = $perm.User.ToString() | Get-Mailbox | select -ExpandProperty UserPrincipalName
			$permsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green
	}
	else
	{
		Write-Host " done" -ForegroundColor DarkCyan
	}
}

foreach ($mbx in $TVGSharedMailboxes)
{
    if ($mailboxcheck = get-mailbox $mbx.FDPrimarySMTPAddress -ea silentlycontinue)
    {
        Write-Host "Adding X500 to mailbox $($mbx.DisplayName)" -ForegroundColor Cyan -NoNewline
        $newX500 = $mbx.LegacyExchangeDN
        Set-Mailbox $mbx.FDPrimarySMTPAddress -EmailAddresses @{add=$newX500}
        Write-Host " .. done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
}

# Set Up Forwarding on FD tenant.

foreach ($mbx in $TVGSharedMailboxes | ? {$_.ForwardingSmtpAddress})
{
    if ($mailboxcheck = get-mailbox $mbx.DisplayName -ea silentlycontinue)
    {
        [boolean]$DeliverToMailboxAndForward = [boolean]::Parse($mbx.DeliverToMailboxAndForward)
        Write-Host "Updating Forwarding for mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        $mailboxcheck | Set-Mailbox -ForwardingSmtpAddress $mbx.ForwardingSmtpAddress -DeliverToMailboxAndForward:$DeliverToMailboxAndForward       
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
    #Read-Host " Pause to check"
}

# Hide from GAL

foreach ($mbx in $TVGSharedMailboxes | ? {$_.HiddenFromAddressListsEnabled -eq $true})
{
    if ($mailboxcheck = get-mailbox $mbx.DisplayName -ea silentlycontinue)
    {
        [boolean]$HiddenFromAddressListsEnabled = [boolean]::Parse($mbx.HiddenFromAddressListsEnabled)
        Write-Host "Updating HiddenInGAL for mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        $mailboxcheck | Set-Mailbox -HiddenFromAddressListsEnabled:$HiddenFromAddressListsEnabled
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
    #Read-Host " Pause to check"
}

# Set PIM data

foreach ($mbx in $TVGSharedMailboxes)
{
    Write-Host "Updating PIM Data for mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
    Set-MsolUser -UserPrincipalName $mbx.FDUPN -Department $mbx.Department -Office $mbx.Office -title $mbx.title
    Write-Host "done" -ForegroundColor Cyan
}

#Apply Full Access
foreach ($user in $TVGFullAccessPerms)
{
    Write-Host "Updating $($user.Mailbox). Granting $($user.UserWithFullAccess) FullAccess ..." #-NoNewline
    Add-MailboxPermission $user.Mailbox -User $user.UserWithFullAccess -AccessRights FullAccess
    #Write-Host " done" -ForegroundColor Green
    #Read-Host "pause to check"
}

#add Alternate Addresses
foreach ($mbx in $TVGSharedMailboxes)
{
    if ($mailboxcheck = get-mailbox $mbx.DisplayName -ea silentlycontinue)
    {
        Write-Host "Adding EmailAddresses to mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        [array]$EmailAddresses = $mbx.EmailAddresses -split ","
            foreach ($altAddress in ($EmailAddresses |  Where {($_-like "*@paddypowerbetfair.com" -or $_ -like "x500*")}))
            {
                if ($altAddress -like "*paddypowerbetfair.com")
                {
                    $Fanduelemailaddressupdate = $altAddress.replace("@paddypowerbetfair.com","@fanduel.com")
                    $tvgemailaddressupdate = $altAddress.replace("@paddypowerbetfair.com","@tvg.com")
                    $tvgnetworkemailaddressupdate = $altAddress.replace("@paddypowerbetfair.com","@tvgnetwork.com")
                    Write-Host "." -ForegroundColor DarkGreen -NoNewline
                    Set-Mailbox $mbx.DisplayName -EmailAddresses @{add=$Fanduelemailaddressupdate} -wa silentlycontinue
                    Set-Mailbox $mbx.DisplayName -EmailAddresses @{add=$tvgemailaddressupdate} -wa silentlycontinue
                    Set-Mailbox $mbx.DisplayName -EmailAddresses @{add=$tvgnetworkemailaddressupdate} -wa silentlycontinue
                }
                else
                {
                    Write-Host "." -ForegroundColor DarkGreen -NoNewline
                    Set-Mailbox $mbx.DisplayName -EmailAddresses @{add=$altAddress} -wa silentlycontinue 
                }
                $PrimarySMTPAddressTVG = $mbx.PrimarySMTPAddress
                Set-Mailbox $mbx.DisplayName -EmailAddresses @{add=$PrimarySMTPAddressTVG} -wa silentlycontinue                
            }
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
    #Read-Host "Pause to check"
}

$NameFullAccessPermsList = @()
#Updated PermsList
foreach ($perm in $permslist | ?{$_.UserWithFullAccess})
{
    #Get Mailbox Details
    $MailboxName = (get-mailbox $perm.Mailbox).displayname
    $UserWithFullAccessName = (get-mailbox $perm.UserWithFullAccess).displayname
    
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "Mailbox" -Value $MailboxName
    $currentuser | add-member -type noteproperty -name "UserWithFullAccess" -Value $UserWithFullAccessName
    $NameFullAccessPermsList += $currentuser
}

# Get Calendar Perms

$CalendarpermsList = @()
foreach ($mbx in $sharedtvgmailboxes)
{
	$upn = $mbx.UserPrincipalName
	[array]$calendars = $mbx | Get-MailboxFolderStatistics | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}
	
	foreach ($calendar in $calendars)
	{
		$folderPath = $calendar.FolderPath.Replace('/','\')
		$id = "$upn`:$folderPath"
		
		Write-Host "$($mbx.PrimarySmtpAddress)`:" -ForegroundColor Cyan -NoNewline
		Write-Host $folderPath -ForegroundColor Green -NoNewline
		Write-Host " ..." -ForegroundColor Cyan -NoNewline
		
		[array]$perms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.User -notlike "Default" -and $_.User -notlike "Anonymous" -and $_.User -notlike "*S-1*"}
		
		if ($perms)
		{
			foreach ($perm in $perms)
			{
                $currentcalendar = new-object PSObject
                
                $currentcalendar | add-member -type noteproperty -name "Mailbox" -Value $mbx.DisplayName
                $currentcalendar | add-member -type noteproperty -name "CalendarName" -Value $perm.FolderName
                $currentcalendar | add-member -type noteproperty -name "CalendarPath" -Value $id
                $currentcalendar | add-member -type noteproperty -name "User" -Value $perm.user
                $currentcalendar | add-member -type noteproperty -name "AccessRights" -Value ($perm.AccessRights -join ",")

                $CalendarpermsList += $currentcalendar

                Write-Host "." -ForegroundColor DarkCyan -NoNewline
			}
			
			Write-Host " done" -ForegroundColor Green
		}
		else
		{
			Write-Host "No Custom Perms found. done" -ForegroundColor Yellow
		}
	}
}

# Set DistributionGroups
foreach ($DL in $FDDLs)
{
   if ($DLCheck = Get-DistributionGroup $DL.Name)
   {
    Write-host "Updating DL $($DL.Name)" -foregroundcolor cyan
    $DLCheck | Set-DistributionGroup -RequireSenderAuthenticationEnabled $DL.RequireSenderAuthenticationEnabled -AcceptMessagesOnlyFromSendersOrMembers $DL.AcceptMessagesOnlyFromSendersOrMembers -AcceptMessagesOnlyFromDLMembers $DL.AcceptMessagesOnlyFromDLMembers -BypassModerationFromSendersOrMembers $DL.BypassModerationFromSendersOrMembers -ModeratedBy $DL.ModeratedBy -EmailAddresses @{add=$EmailAddresses}
   }
    else {
        $notfoundusers += $dl
        Write-host "No DL $($DL.Name) found" -foregroundcolor red
    }
}

$DL.Name
$DL.RequireSenderAuthenticationEnabled
$DL.AcceptMessagesOnlyFrom
$DL.AcceptMessagesOnlyFromDLMembers
$DL.AcceptMessagesOnlyFromSendersOrMembers
$DL.BypassModerationFromSendersOrMembers
$DL.ModeratedBy


### FD Groups Match

$AllGroups =@()
$foundgroups =@()
$notfoundgroups =@()

foreach ($group in $importgroups | sort Name) {
    Write-Host "Checking for $($group.DisplayName) on destination ..." -fore Cyan -NoNewline
    
    $currentgroup = new-object PSObject
                
    $currentgroup | add-member -type noteproperty -name "DisplayName" -Value $group.DisplayName
    $currentgroup | add-member -type noteproperty -name "Alias" -Value $group.Alias
    $currentgroup | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $group.PrimarySmtpAddress
    $currentgroup | add-member -type noteproperty -name "MemberDepartRestriction" -Value $group.MemberDepartRestriction
    $currentgroup | add-member -type noteproperty -name "RequireSenderAuthenticationEnabled" -Value $group.RequireSenderAuthenticationEnabled
    $currentgroup | add-member -type noteproperty -name "ManagedBy" -Value $group.ManagedBy
    $currentgroup | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value $group.AcceptMessagesOnlyFrom
    $currentgroup | add-member -type noteproperty -name "AcceptMessagesOnlyFromDLMembers" -Value $group.AcceptMessagesOnlyFromDLMembers
    $currentgroup | add-member -type noteproperty -name "AcceptMessagesOnlyFromSendersOrMembers" -Value $group.AcceptMessagesOnlyFromSendersOrMembers
    $currentgroup | add-member -type noteproperty -name "ModeratedBy" -Value $group.ModeratedBy
    $currentgroup | add-member -type noteproperty -name "BypassModerationFromSendersOrMembers" -Value $group.BypassModerationFromSendersOrMembers
    $currentgroup | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value $group.GrantSendOnBehalfTo
    $currentgroup | add-member -type noteproperty -name "ModerationEnabled" -Value $group.ModerationEnabled
    $currentgroup | add-member -type noteproperty -name "LegacyExchangeDN" -Value $group.LegacyExchangeDN
    $currentgroup | add-member -type noteproperty -name "EmailAddresses" -Value $group.EmailAddresses 
    
    if ($Group = Get-DistributionGroup $group.DisplayName -ea silentlycontinue)  {
        $GroupMembers = Get-DistributionGroupMember $group.DisplayName
        
        $foundgroups += $group.DisplayName
        Write-Host "found" -ForegroundColor Green
        $currentgroup | add-member -type noteproperty -name "ExistsOnO365" -Value $true
        $currentgroup | add-member -type noteproperty -name "MatchedDLName" -Value $group.ModerationEnabled
        $currentgroup | add-member -type noteproperty -name "MatchedDLPrimarySMTPAddress" -Value $group.LegacyExchangeDN
        $currentgroup | add-member -type noteproperty -name "DLMemberCount" -Value $GroupMembers.count
    }

    else
    {
        $notfoundgroups += $group
        $currentgroup | add-member -type noteproperty -name "ExistsOnO365" -Value $false
        $currentgroup | add-member -type noteproperty -name "MatchedDLName" -Value $group.ModerationEnabled
        $currentgroup | add-member -type noteproperty -name "MatchedDLPrimarySMTPAddress" -Value $group.LegacyExchangeDN
        $currentgroup | add-member -type noteproperty -name "DLMemberCount" -Value $GroupMembers.count

        Write-Host "not found" -ForegroundColor red
    }

    $AllGroups += $currentgroup
}

## Pull Exec Calendars
$CalendarpermsList = @()
foreach ($user in $mailboxes)
{
    $mbx = get-mailbox $user
	$upn = $mbx.UserPrincipalName
	[array]$calendars = $mbx | Get-MailboxFolderStatistics | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}
	
	foreach ($calendar in $calendars)
	{
		$folderPath = $calendar.FolderPath.Replace('/','\')
		$id = "$upn`:$folderPath"
		
		Write-Host "$($mbx.PrimarySmtpAddress)`:" -ForegroundColor Cyan -NoNewline
		Write-Host $folderPath -ForegroundColor Green -NoNewline
		Write-Host " ..." -ForegroundColor Cyan -NoNewline
		
		[array]$perms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.User -notlike "Default" -and $_.User -notlike "Anonymous" -and $_.User -notlike "*S-1*"}
		
		if ($perms)
		{
			foreach ($perm in $perms)
			{
                $currentcalendar = new-object PSObject
                
                $currentcalendar | add-member -type noteproperty -name "Mailbox" -Value $mbx.DisplayName
                $currentcalendar | add-member -type noteproperty -name "CalendarName" -Value $perm.FolderName
                $currentcalendar | add-member -type noteproperty -name "CalendarPath" -Value $id
                $currentcalendar | add-member -type noteproperty -name "User" -Value $perm.user
                $currentcalendar | add-member -type noteproperty -name "AccessRights" -Value ($perm.AccessRights -join ",")

                $CalendarpermsList += $currentcalendar

                Write-Host "." -ForegroundColor DarkCyan -NoNewline
			}
			
			Write-Host " done" -ForegroundColor Green
		}
		else
		{
			Write-Host "No Custom Perms found. done" -ForegroundColor Yellow
		}
	}
}

$CalendarpermsList | Export-csv "$($HOME)\$($csvFileName)" –notypeinformation –encoding utf8

<# apply calendar perms
.EXAMPLE
Pull all UserMailboxes to a single OU. This is has to be used with the OnPrem Switch. Using this will be needed in Hosted Exchange.
.\Get-Mailboxattributes.ps1 -OutputCSVFilePath "c:\temp" -OnPrem $true -hexOU contoso.com
#>

param(
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Output CSV File path")][string]$OutputCSVFilePath,
)

foreach ($CalendarPerm in $AllCalendarPerms)
{   
    $updateCalendarPath= $CalendarPerm.CalendarPath.Replace("@paddypowerbetfair.com","@fanduel.com")
    Write-Host "Updating $($updateCalendarPath). Adding Calendar perms for $($CalendarPerm.User) ..." -NoNewline -foregroundcolor cyan
    Add-MailboxFolderPermission -Identity $updateCalendarPath -User $CalendarPerm.User -AccessRights $CalendarPerm.AccessRights -ea silentlycontinue
    Write-Host "done" -ForegroundColor Green
}
$csvFileName = "AllLicensedMailboxes_$($date).csv"

| Export-csv "$($OutputCSVFilePath)\$($csvFileName)" –notypeinformation –encoding utf8


#Create Share Mailbox
foreach ($sharedmailbox in $foxbetsharedmailboxes)
{
    #create the shared mailbox
    $FanDuelGroupAddress = $sharedmailbox.EmailAddress.replace("@fox.bet","@fanduelgroup.onmicrosoft.com")
    $aliasSplit = $sharedmailbox.EmailAddress -split "@"
    $aliasString = $aliasSplit[0]
    Write-host "Adding $($sharedmailbox.DisplayName) .." -foregroundcolor cyan -nonewline
    New-Mailbox -shared -PrimarySMTPAddress $FanDuelGroupAddress -Alias $aliasString -name $sharedmailbox.DisplayName -displayname $sharedmailbox.DisplayName
    Write-host "done " -foregroundcolor green
}

foreach ($sharedmailbox in $foxbetsharedmailboxes)
{
    #create the shared mailbox
    Get-Mailbox $sharedmailbox.DisplayName | select DisplayName, alias, PrimarySMTPAddress
}

#Grant Full Access to SharedMailboxes

#Create Share Mailbox
foreach ($sharedmailbox in $foxbetsharedmailboxes)
{
    #create the shared mailbox
    $FanDuelGroupAddress = $sharedmailbox.EmailAddress.replace("@fox.bet","@fanduelgroup.onmicrosoft.com")
    $aliasSplit = $sharedmailbox.EmailAddress -split "@"
    $aliasString = $aliasSplit[0]
    [array]$perms = $sharedmailbox.Permissions -split ","
    
    foreach ($perm in $perms)
    {
        if ($mailboxcheck = Get-Mailbox $perm)
        {
            Write-host "Granting $($mailboxcheck.DisplayName) Full Access perms to $($sharedmailbox.DisplayName) .." -foregroundcolor cyan -nonewline
            Add-MailboxPermission $sharedmailbox.DisplayName -user $mailboxcheck.DisplayName -AccessRights FullAccess
            Write-host "done " -foregroundcolor green
        } 
    }       
}

## Add TVG Addresses
$allTVGAddresses = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\TVGALLAddresses.csv"

$MultipleAliases = @()
$updatedTVGUsers = @()
$notfoundusers = @()
$allTVGUsers = @()
foreach ($user in $allTVGAddresses)
{
    $currentuser = new-object psobject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName
    $currentuser | add-member -type noteproperty -name "TVGEmailAddress" -Value $user.TVGEmailAddress
    

    if ($mailboxcheck = get-mailbox $user.DisplayName -ea silentlycontinue)
    {
        $currentuser | add-member -type noteproperty -name "MailboxFound" -Value $true
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $mailboxcheck.PrimarySMTPAddress

        if ($Mailboxcheck.count -gt 1)
        {
            Write-Host "Multiple Mailboxes found for $($mailboxcheck.DisplayName). Skipping .." -foregroundcolor yellow
            $MultipleAliases += $mailboxcheck.PrimarySMTPAddress
            $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value "MultipleFound"
        }
        else {
            $TVGAddress = $user.TVGEmailAddress
            if (!($mailboxcheck | ?{$_.EmailAddresses -like "*$TVGAddress*"}))
            {
                Write-Host "Added TVG Address $($TVGAddress) to $($mailboxcheck.DisplayName)" -foregroundcolor green
                Set-Mailbox $user.DisplayName -EmailAddresses @{add=$TVGAddress}
                $updatedTVGUsers += $Mailboxcheck.PrimarySMTPAddress
                $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $mailboxcheck.PrimarySMTPAddress
            }
        }
    }
    else {
        Write-Host "No Mailbox Found for $($user.DisplayName)" -foregroundcolor red 
        $notfoundusers += $user
        $currentuser | add-member -type noteproperty -name "MailboxFound" -Value $false
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $null
    }
    $allTVGUsers += $currentuser
}

$foundusers =@()
$notfoundusers = @()
$TVGAddresses | ?{$_.MailboxFound -eq $False}| foreach {
    if (Get-Recipient $_.DisplayName -EA silentlycontinue)
    {
    Write-Host $_.DisplayName "found" -ForegroundColor green
    $foundusers += $_
    }
    else
    {
    Write-Host $_.DisplayName "notfound" -ForegroundColor red
    $notfoundusers += $_
    }
}

$updateFDmailboxes