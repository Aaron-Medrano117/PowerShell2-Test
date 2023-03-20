#FanDuel Scripts

#Get ALL MAILBOX DETAILS and ONEDRIVE DETAILS
function Get-ALLFDMAILBOXDETAILS {
    param ()
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Where {($_.PrimarySmtpAddress -like "*fanduel.com" -or $_.PrimarySMTPAddress -like "*@tvg*")} | sort PrimarySmtpAddress

    $FanDuelUsers = @()
    $MultipleAliases = @()
    $SitesNotFound = @()

    foreach ($user in $Mailboxes)
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

            #get OneDrive Site details
            
            $SPOSite = $null
            $ODSite = "https://betfairprod-my.sharepoint.com/personal/" + $user.alias.replace(".","_") + "_paddypowerbetfair_com"

            try {
                $SPOSITE = Get-SPOSITE $ODSite -ErrorAction SilentlyContinue
            }
            catch {
                Write-Host "OneDrive Not Enabled for User" -ForegroundColor Yellow
                $SitesNotFound += $FDUser
            }
            
            
            if ($SPOSITE)
            {
                Write-Host "Gathering OneDrive Details ..." -ForegroundColor Cyan -NoNewline
                
                $currentuser | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $ODSite
                $currentuser | Add-Member -type NoteProperty -Name "Owner" -Value $SPOSITE.Owner
                $currentuser | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $SPOSITE.StorageUsageCurrent
                $currentuser | Add-Member -type NoteProperty -Name "Status" -Value $SPOSITE.Status
                $currentuser | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $SPOSITE.SiteDefinedSharingCapability
                $currentuser | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $FDUser.LimitedAccessFileType
                
                Write-Host "done" -ForegroundColor Green
            }
        $FanDuelUsers += $currentuser
        }
    }

    $FanDuelUsers | Export-Csv -NoTypeInformation -Encoding utf8 "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetails.csv"

}
Get-ALLFDMAILBOXDETAILS

## FanDuel DLs

$FanDuelLists = @()
foreach ($dl in $FanDuelDLs)
{
    Write-Host "Gathering DL Details for $($dl.DisplayName) .." -ForegroundColor Cyan -NoNewline
    $DLMembers = Get-DistributionGroupMember $dl.alias
    Write-Host "done" -ForegroundColor Green
    
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $dl.DisplayName
    $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $dl.PrimarySMTPAddress
    $currentuser | add-member -type noteproperty -name "Alias" -Value $dl.alias
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value $dl.LegacyExchangeDN
    $currentuser | add-member -type noteproperty -name "DLMemberCount" -Value $DLMembers.Count
    $FanDuelLists += $currentuser
}

#Calendar Perms

$mailboxes = Get-Mailbox -ResultSize Unlimited | Where {($_.PrimarySmtpAddress -like "*fanduel.com" -or $_.PrimarySMTPAddress -like "*@tvg*")} | sort PrimarySmtpAddress

$permsList = @()
foreach ($mbx in $mailboxes[1001..2197])
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

                $permsList += $currentcalendar

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

# Send-As Perms
$recipients = Get-Recipient -ResultSize Unlimited | Where {($_.PrimarySmtpAddress -like "*fanduel.com" -or $_.PrimarySMTPAddress -like "*@tvg*")} | sort PrimarySmtpAddress

$sendAsPermsList = @()
foreach ($recipient in $recipients)
{
	Write-Host "$($recipient.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	$perms = $recipient | Get-RecipientPermission
	$perms = $perms | Where {$_.Trustee	-notlike "NT AUTHORITY*" -and $_.Trustee -notlike "S-1-*"}
	
	if ($perms)
	{
		foreach ($perm in $perms)
		{
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select Recipient, RecipientType, ObjectWithSendAs
			$tmp.Recipient = $recipient.PrimarySmtpAddress.ToString()
			$tmp.RecipientType = $recipient.RecipientTypeDetails
			$tmp.ObjectWithSendAs = $perm.Trustee.ToString() | Get-Recipient -EA SilentlyContinue | select -ExpandProperty PrimarySmtpAddress
			$sendAsPermsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green

	}
	else
	{
		Write-Host " No custom Perms found." -ForegroundColor Yellow
	}
}

$sendAsPermsList | Export-Csv -NoTypeInformation  "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\365 Source Exports\FanDuel_SendAsPerms.csv"


#### Multiple Alias check

$recipients = Get-Recipient -ResultSize unlimited | ? {$_.primarysmtpaddress -like "*fanduel.com" -or $_.PrimarySMTPAddress -like "*@tvg*"} | sort primarysmtpaddress

$MultipleAliases = @()

foreach ($user in $recipients)
{
    Write-Host "Gathering Mailbox Stats for $($user.DisplayName) .." -ForegroundColor Cyan -NoNewline
    $AliasCheck = Get-Recipient $user.Alias | select Name, Alias, primarysmtpaddress, RecipientType
    
    if ($AliasCheck.count -gt 1)
    {
        Write-Host "Multiple Aliases found for User" -ForegroundColor Red
        $MultipleAliases += $AliasCheck
    }
}

#### Full Access Check

$mailboxes = Get-Mailbox -ResultSize Unlimited | Where {($_.PrimarySmtpAddress -like "*fanduel.com" -or $_.PrimarySMTPAddress -like "*@tvg*")} | sort PrimarySmtpAddress

$permsList = @()
foreach ($mbx in $mailboxes)
{
	Write-Host "$($mbx.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	[array]$perms = $mbx | Get-MailboxPermission | Where {!$_.IsInherited -and $_.AccessRights -like "*FullAccess*"}
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

$permsList | Export-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\Perms_FullAccess.csv" -NoTypeInformation -Encoding UTF8


## Enable Archive for users
# have to be licensed first
foreach ($user in $importcsv | ? {$_.ArchiveGUID})
{
	Write-Host "Enabling Archive for $($user.DisplayName)"
	$fdgroupOnMicrosoft = $user.PrimarySMTPAddress.replace("@fanduel.com","@fanduelgroup.onmicrosoft.com")
    Enable-Mailbox -identity $fdgroupOnMicrosoft -Archive -ea silentlycontinue
}


### FD Groups Match

$AllGroups =@()
$foundgroups =@()
$notfoundgroups =@()

foreach ($group in $importgroups) {
    Write-Host "Checking for $($group.DisplayName) on tenant ..." -fore Cyan -NoNewline
    
    $tmp = "" | select DisplayName, PrimarySMTPAddress, Alias, SourceMemberCount, Matched, MatchedDestinationDLName, MatchedDestinationDLPrimarySMTPAddress, DestinationMemberCount
    $tmp.DisplayName = $group.DisplayName
    $tmp.PrimarySMTPAddress = $group.PrimarySMTPAddress
	$tmp.Alias = $group.Alias
	$tmp.SourceMemberCount = $group.count  
    
    if ($DL = Get-DistributionGroup $group.DisplayName -ea silentlycontinue)  {
        $GroupMembers = Get-DistributionGroupMember $group.DisplayName
        
        $foundgroups += $group.PrimarySMTPAddress
        Write-Host "found" -ForegroundColor Green
        $tmp.MatchedDestinationDLName = $DL.DisplayName
        $tmp.MatchedDestinationDLPrimarySMTPAddress = $DL.PrimarySMTPAddress
		$tmp.SourceMemberCount = $GroupMembers.count
		$tmp.DestinationMemberCount = $GroupMembers.count
        $tmp.Matched = $true
    }

    else
    {
        $notfoundgroups += $group
        Write-Host "not found" -ForegroundColor red
        $tmp.Matched = $False
    }

    $AllGroups += $tmp
}

##Fanduel Match MSOL 2! for FanDuel
$importcsv = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()
$MultipleUsers = @()

foreach ($user in $importcsv | sort DisplayName) {
    $DisplayName = $user.DisplayName
    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline

    $tmp = "" | select DisplayName, FDUPN, PrimarySmtpAddress, RecipientTypeDetails, IsInactiveMailbox, ExistsOnO365, PaddyPowerUserPrincipalName, IsLicensed, ExternalEmailAddress
    $tmp.DisplayName = $DisplayName
    $tmp.FDUPN = $user.FanDuelGroupAddress
    
    if ($MSOLUsers = Get-MsolUser -UserPrincipalName $user.FanDuelGroupAddress -ea silentlycontinue) {
		$tmp.ExistsOnO365 = $true

        if ($msolusers.count -gt 1)
        {
            Write-Host "Multiple Users Found. Skip" -ForegroundColor Red -NoNewline
            $tmp.RecipientTypeDetails = "MultipleUsersFound"
            $MultipleUsers += $msolusers
        }

        else
        {
            $foundusers += $user
            
            if ($mailbox = Get-Mailbox $msolusers.FanDuelGroupAddress -IncludeInactiveMailbox -ea silentlycontinue | ? {$_.PrimarySmtpAddress -Like "*@fanduel.com" -or $_.PrimarySmtpAddress -like "*@tvg*"})
            {
                $tmp.DisplayName = $mailbox.DisplayName
                $tmp.PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                $tmp.RecipientTypeDetails = $mailbox.RecipientTypeDetails
                $tmp.IsInactiveMailbox = $mailbox.IsInactiveMailbox
                Write-Host "found" -ForegroundColor Green -NoNewline                    
            }

            elseif ($recipient = Get-Recipient $DisplayName -ea silentlycontinue)
            {
                $tmp.DisplayName = $recipient.DisplayName
                $tmp.PrimarySmtpAddress = $recipient.PrimarySmtpAddress
                $tmp.RecipientTypeDetails = $recipient.RecipientType
                $tmp.ExternalEmailAddress = $recipient.ExternalEmailAddress
                Write-Host "found" -ForegroundColor Green -NoNewline
                    
            }
                
            $tmp.PaddyPowerUserPrincipalName = $msolusers.UserPrincipalName
            $tmp.IsLicensed = $msolusers.IsLicensed
        }
            Write-Host "..done" -ForegroundColor Green
    }      

    else
    {
        $notfoundusers += $user
        Write-Host "not found" -ForegroundColor red
        $tmp.DisplayName = $DisplayName
        $tmp.PrimarySmtpAddress = $user.PrimarySMTPAddress
    }

    $AllUsers += $tmp
}

##Fanduel Match Mailbox
$importcsv = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()
$MultipleUsers = @()

foreach ($user in $importcsv | sort DisplayName) {
    $DisplayName = $user.DisplayName
    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline

    $tmp = "" | select DisplayName, PrimarySmtpAddress, ExistsOnO365, RecipientType, UserPrincipalName, IsLicensed, IsInactiveMailbox
    $tmp.DisplayName = $user.DisplayName    

    if ($mailbox = Get-Mailbox $user.FanDuelGroupAddress -IncludeInactiveMailbox -ea silentlycontinue) {

		if ($mailbox.count -gt 1)
		{
			Write-Host "Multiple mailboxes found. Skipping." -foregroundcolor red
			$MultipleUsers += $User
		}
        
        $foundusers += $user
        Write-Host "found" -ForegroundColor Green
        $tmp.ExistsOnO365 = $true

        #Get Mailbox Details
        $tmp.PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        $tmp.RecipientType = $mailbox.RecipientTypeDetails
        $tmp.UserPrincipalName = $mailbox.UserPrincipalName
        $tmp.IsInactiveMailbox = $mailbox.IsInactiveMailbox

        # Get MSOL User Information
        $Msoluser = get-msoluser -userprincipalname $user.FanDuelGroupAddress | select IsLicensed

        $tmp.IsLicensed = $msoluser.IsLicensed
    }

    else
    {
        $notfoundusers += $user
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    $AllUsers += $tmp
}

##Fanduel Match Recipient
$FDMigMailboxes = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in ($FDMigMailboxes | ? {$_.IsLicensed -eq $true})) {
    $DisplayName = $user.DisplayName
    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline

    $tmp = "" | select DisplayName, PrimarySmtpAddress, ExistsOnO365, RecipientType, UserPrincipalName, IsLicensed
    $tmp.DisplayName = $DisplayName
    

    if ($mailboxcheck = Get-Mailbox $DisplayName -ea silentlycontinue) {
        
        if ($mailboxcheck.count -gt 1)
        {
            Write-Host "Multiple Mailboxes found ..." -ForegroundColor Yellow -NoNewline
            foreach ($mbx in $mailboxcheck)
            {
                $foundusers += $user

                #Get Mailbox Details
                $tmp.PrimarySmtpAddress = $mbx.PrimarySmtpAddress
                $tmp.RecipientType = $mbx.RecipientTypeDetails
                $tmp.UserPrincipalName = $mbx.UserPrincipalName

                # Get MSOL User Information
                $Msoluser = get-msoluser -userprincipalname $mbx.UserPrincipalName | select IsLicensed
                Write-Host "MSOLUser Details Gathered" -ForegroundColor Green
                $tmp.IsLicensed = $msoluser.IsLicensed
            }
        }
        else
        {
            $foundusers += $user
            Write-Host "found ..." -ForegroundColor Green -NoNewline
            $tmp.ExistsOnO365 = $true

            #Get Mailbox Details
            $tmp.PrimarySmtpAddress = $mailboxcheck.PrimarySmtpAddress
            $tmp.RecipientType = $mailboxcheck.RecipientTypeDetails
            $tmp.UserPrincipalName = $mailboxcheck.UserPrincipalName

            # Get MSOL User Information
            $Msoluser = get-msoluser -userprincipalname $mailboxcheck.UserPrincipalName | select IsLicensed
            Write-Host "MSOLUser Details Gathered" -ForegroundColor Green
            $tmp.IsLicensed = $msoluser.IsLicensed
        }
    }
    else
    {
        $notfoundusers += $user
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    $AllUsers += $tmp
}

##Fanduel Match USERs. Grab MSOL Properties. Checks for multiple matches
$importcsv = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $importcsv | sort DisplayName) {
    $DisplayName = $user.DisplayName
    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline

    $tmp = "" | select DisplayName, PrimarySmtpAddress, RecipientTypeDetails, IsInactiveMailbox, ExistsOnO365, UserPrincipalName, IsLicensed, ExternalEmailAddress

    if ($MSOLUsers = Get-MsolUser -searchstring $DisplayName -ea silentlycontinue) {

        $foundusers += $user
        

            foreach ($msoluser in $MSOLUsers)
            {
                if ($mailbox = Get-Mailbox $msoluser.userprincipalname -IncludeInactiveMailbox -ea silentlycontinue)
                {
                    $tmp.DisplayName = $mailbox.DisplayName
                    $tmp.PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    $tmp.RecipientTypeDetails = $mailbox.RecipientTypeDetails
                    $tmp.IsInactiveMailbox = $mailbox.IsInactiveMailbox
                    Write-Host "found" -ForegroundColor Green -NoNewline                    
                }

                elseif ($recipient = Get-Recipient $DisplayName -ea silentlycontinue | ? {$_.RecipientType -ne "UserMailbox"})
                {
                    $tmp.DisplayName = $recipient.DisplayName
                    $tmp.PrimarySmtpAddress = $recipient.PrimarySmtpAddress
                    $tmp.RecipientTypeDetails = $recipient.RecipientType
                    $tmp.ExternalEmailAddress = $recipient.ExternalEmailAddress
                    Write-Host "found" -ForegroundColor Green -NoNewline
                    
                }
                else
                {
                    Write-Host " .. Not a valid Exchange Recipient for $($msoluser.userprincipalname)" -ForegroundColor red -NoNewline
                }       
                
                $tmp.ExistsOnO365 = $true
                $tmp.UserPrincipalName = $msoluser.UserPrincipalName
                $tmp.IsLicensed = $MSOLUser.IsLicensed
            }
            Write-Host "..done" -ForegroundColor Cyan
        }
        

    else
    {
        $notfoundusers += $user
        Write-Host "not found" -ForegroundColor red
        $tmp.DisplayName = $user.DisplayName
        $tmp.PrimarySmtpAddress = $user.PrimarySmtpAddress
        $tmp.ExistsOnO365 = $False
    }

    $AllUsers += $tmp
}


### FanDuel Group Migration Addresses

$OnMicrosoft = @()
foreach ($user in $Mailboxes) {
    $EmailAddresses = $user.EmailAddresses -split ","
    
    foreach ($Address in $EmailAddresses | ? {$_ -like "*fanduelgroup.onmicrosoft.com"})
    {
        $tmp = "" | select DisplayName,PrimarySmtpAddress, Alias, OnMicrosoftAddress
        $tmp.DisplayName = $user.DisplayName
        $tmp.PrimarySmtpAddress = $user.PrimarySmtpAddress
        $tmp.Alias= $user.alias
        $tmp.OnMicrosoftAddress= $Address
        
        $OnMicrosoft += $tmp
    }
}


### Check DL Name
$dlcheck = @()
$AllGroups | ?{$_.Matched -eq $true -and $_.SourceMemberCount -eq $null} | foreach {
	$DL = Get-DistributionGroup $_.DisplayName
	$DLMembers = Get-DistributionGroupMember $_.primarysmtpaddress

	$tmp = "" | select DLNAME, DLMembersCount, DLMemberName
	$tmp.DLNAME = $DL.DisplayName
	$tmp.DLMembersCount = $DLMembers.count
	$tmp.DLMemberName = $DLMembers.name
	$dlcheck += $tmp
}
$dlcheck


# Check if softdeleted mailbox

$RemainingUsers =@()
$inactivemailbox =@()
$activemailbox =@()

foreach ($user in $failedimportcsv | sort upn) {
	$FDUser = $user.upn.replace("@paddypowerbetfair.com","@fanduel.com")
    Write-Host "Checking for $FDUser Soft Deleted Mailbox in Source Tenant ..." -fore Cyan -NoNewline

	$tmp = "" | select UPN, InActivePrimarySmtpAddress, IsInactiveMailbox, RecipientType, UserPrincipalName, IsLicensed
	$tmp.UPN = $user.upn  

    if ($mailbox = Get-Mailbox $FDUser -InactiveMailboxOnly -ea silentlycontinue) {
        
        $inactivemailbox += $user
		Write-Host "Inactive mailbox" -ForegroundColor red
		$tmp.IsInactiveMailbox = $true

        #Get Mailbox Details
        $tmp.InActivePrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        $tmp.RecipientType = $mailbox.RecipientTypeDetails
        $tmp.UserPrincipalName = $mailbox.UserPrincipalName
    }

    else
    {
        $activemailbox += $user
        Write-Host "MailUser" -ForegroundColor Green
        $tmp.IsInactiveMailbox = $False
    }

    $RemainingUsers += $tmp
}


#### check shared mailboxes have archive enabled.
foreach ($user in ($importcsv | ? {$_.RecipientTypeDetails -eq "SharedMailbox" -and $_.ArchiveGuid}))
{
	Write-Host "Creating User "$user.DisplayName" ..." -ForegroundColor Cyan -NoNewline
	$FDGroupAddress = $user.primarysmtpaddress.replace("@fanduel.com","@fanduelgroup.onmicrosoft.com")

	if (Get-Mailbox $FDGroupAddress -ea silentlycontinue) {
		Write-Host "Mailbox exists. Skipping" -ForegroundColor yellow
	}
   else
   {
	New-Mailbox -shared -name $user.DisplayName -primarysmtpaddress $FDGroupAddress -displayname $user.DisplayName -FirstName $user.FirstName -LastName $user.LastName -whatif
	Write-Host "done" -ForegroundColor Green
   }   
}

# Check if mailbox archive enabled

$ArchiveExists = @()
$EnabledArchive = @()

foreach ($user in $archivembxs)
{
	Write-Host "Checking User $user ..." -ForegroundColor Cyan -NoNewline
	
	if (Get-Mailbox $user -Archive -ea silentlycontinue) {
        Write-Host "Archive Mailbox exists. Skipping" -ForegroundColor yellow
        $ArchiveExists += $user
	}
   else
   {
        Enable-Mailbox $user -Archive
        Write-Host "Enabled Archive" -ForegroundColor Green
        $EnabledArchive += $user
   }   
}

# get list of all users and license status
$LicensedUsers = @()
foreach ($fduser in $mailboxes)
{
	$msoluser = Get-MsolUser -UserPrincipalName $fduser.UserPrincipalName | select DisplayName, UserPrincipalName, IsLicensed
	
		$currentuser = new-object PSObject
		
		$currentuser | add-member -type noteproperty -name "DisplayName" -Value $msoluser.DisplayName
		$currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $fduser.userprincipalname
		$currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msoluser.IsLicensed
		$currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $fduser.PrimarySMTPAddress

		$LicensedUsers += $currentuser
}

#ReplaceUPN
$msoluserscheck = get-msoluser -all | ?{$_.DisplayName -notlike '*ADM -*' -and $_.UserPrincipalName -eq "migrationwiz@fanduel.com" -and $_.Displayname -notlike "On-Premises*"}


foreach ($user in $msolusers | sort userprincipalname)
{
	$newUPN = $user.UserPrincipalName.replace("@fanduelgroup.onmicrosoft.com","@fanduelgroup.com")
	Write-Host "Updating UPN for $($user.displayname) to $newupn .." -ForegroundColor Cyan -NoNewline
	Set-MsolUserPrincipalName -UserPrincipalName $user.UserPrincipalName -NewUserPrincipalName $newUPN
	Write-Host "done" -ForegroundColor Green
}


####   Get Fanduel groups
$PaddyPowerGroups = Get-DistributionGroup -ResultSize Unlimited | ? {$_.PrimarySMTPAddress -notlike "*@fanduel*" -or $_.PrimarySMTPAddress -notlike "*@tvg*"}

$PaddyPowerGroupsMembers = @()
foreach ($group in $PaddyPowerGroups) {
    Write-Host "Gathering Group Details for $($group.DisplayName) .." -NoNewline -ForegroundColor Cyan
    $DLMembers = Get-DistributionGroupMember -Identity $group.primarysmtpaddress -resultsize unlimited | ? {$_.primarysmtpaddress -notlike "*betfair*" -and $_.PrimarySMTPAddress -notlike "*@ppb.com" -and $_.PrimarySMTPAddress -notlike "*@timeform.com"}
    Write-Host "Found $($DLMembers.count) members .." -ForegroundColor Yellow -NoNewline

    foreach ($member in $DLMembers) {
        $currentgroup = New-Object psobject
        $currentgroup | Add-Member -Type noteproperty -Name "DLGroupName" -Value $group.DisplayName
        $currentgroup | Add-Member -Type noteproperty -Name "DLGroupPrimarySMTPAddress" -Value $group.PrimarySMTPAddress
        $currentgroup | Add-Member -Type noteproperty -Name "Member" -Value $member.DisplayName
        $currentgroup | Add-Member -Type noteproperty -Name "RecipientType" -Value $member.RecipientType
        $currentgroup | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $member.PrimarySMTPAddress
                
        $PaddyPowerGroupsMembers += $currentgroup
        }
        Write-Host "done" -ForegroundColor Green
}

### Get Fanduel groups (less verbose)

$PaddyPowerGroups = Get-DistributionGroup -ResultSize Unlimited | ? {$_.PrimarySMTPAddress -notlike "*@fanduel*" -or $_.PrimarySMTPAddress -notlike "*@tvg*"}

$PaddyPowerGroupsMembers = @()
foreach ($group in $PaddyPowerGroups) {
    if ($DLMembers = Get-DistributionGroupMember -Identity $group.primarysmtpaddress -resultsize unlimited | ? {$_.primarysmtpaddress -notlike "*betfair*" -and $_.PrimarySMTPAddress -notlike "*@ppb.com" -and $_.PrimarySMTPAddress -notlike "*@timeform.com" -and $_.PrimarySMTPAddress -notlike "*@flutter.com" -and $_.PrimarySMTPAddress -notlike "*@blip.pt"})
    {
        Write-Host "Gathering Group Details for $($group.DisplayName) .." -NoNewline -ForegroundColor Cyan
    
        Write-Host "Found $($DLMembers.count) members .." -ForegroundColor Yellow -NoNewline
    
        foreach ($member in $DLMembers) {
        $currentgroup = New-Object psobject
        $currentgroup | Add-Member -Type noteproperty -Name "DLGroupName" -Value $group.DisplayName
        $currentgroup | Add-Member -Type noteproperty -Name "DLGroupPrimarySMTPAddress" -Value $group.PrimarySMTPAddress
        $currentgroup | Add-Member -Type noteproperty -Name "Member" -Value $member.DisplayName
        $currentgroup | Add-Member -Type noteproperty -Name "RecipientType" -Value $member.RecipientType
        $currentgroup | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $member.PrimarySMTPAddress
                    
        $PaddyPowerGroupsMembers += $currentgroup
        }
        Write-Host "done" -ForegroundColor Green
    }
}


### Get Fanduel groups (less verbose and looking for only domains)

$PaddyPowerGroups = Get-DistributionGroup -ResultSize Unlimited | ? {$_.PrimarySMTPAddress -notlike "*@fanduel*" -or $_.PrimarySMTPAddress -notlike "*@tvg*"}

$PaddyPowerGroupsMembers = @()
foreach ($group in $PaddyPowerGroups) {
    if ($DLMembers = Get-DistributionGroupMember -Identity $group.primarysmtpaddress -resultsize unlimited | ? {$_.primarysmtpaddress -like "*@fanduel.com" -or $_.PrimarySMTPAddress -like "*@tvg*"})
    {
        Write-Host "Gathering Group Details for $($group.DisplayName) .." -NoNewline -ForegroundColor Cyan
    
        Write-Host "Found $($DLMembers.count) members .." -ForegroundColor Yellow -NoNewline
    
        foreach ($member in $DLMembers) {
        $currentgroup = New-Object psobject
        $currentgroup | Add-Member -Type noteproperty -Name "DLGroupName" -Value $group.DisplayName
        $currentgroup | Add-Member -Type noteproperty -Name "DLGroupPrimarySMTPAddress" -Value $group.PrimarySMTPAddress
        $currentgroup | Add-Member -Type noteproperty -Name "Member" -Value $member.DisplayName
        $currentgroup | Add-Member -Type noteproperty -Name "RecipientType" -Value $member.RecipientType
        $currentgroup | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $member.PrimarySMTPAddress
                    
        $PaddyPowerGroupsMembers += $currentgroup
        }
        Write-Host "done" -ForegroundColor Green
    }
}

## Match Updated Export to Previous Export

function Compare-FanDuelUsers {
    param ()

    $OldUserList = Import-csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\FanDuelFullDetails_old.csv"
    $NewUserList = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\FanDuelFullDetails.csv"
    $recentlyModifiedUsers = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\RecentlyUpdatedUsers.csv"

    $removedUsers =@()
    $addedUsers = @()
    $updatedUsers = @()

foreach ($user in $OldUserList | sort DisplayName) {
    $userMatch = $NewUserList | ? {$_.PrimarySmtpAddress-eq $user.PrimarySmtpAddress}
    if (!$userMatch)
        {
            $removedUsers += $user
            Write-Host $user.DisplayName "Not found" -ForegroundColor Red
            
            $currentuser = new-object PSObject
            $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName
            $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $user.UserPrincipalName
            $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $user.PrimarySmtpAddress
            $currentuser | Add-Member -type NoteProperty -Name "Update" -Value "Removed"
            $currentuser | Add-Member -Type NoteProperty -Name "LastChecked" -Value (Get-Date)
            $updatedUsers += $currentuser
        }
}

foreach ($user in $NewUserList | sort DisplayName) {
    $userMatch = $OldUserList | ? {$_.PrimarySmtpAddress-eq $user.PrimarySmtpAddress}
    if (!$userMatch)
    {
        $addedUsers += $user
        Write-Host $user.DisplayName "Recently Added" -ForegroundColor Green

        $currentuser = new-object PSObject
        $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName
        $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $user.UserPrincipalName
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $user.PrimarySmtpAddress
        $currentuser | Add-Member -type NoteProperty -Name "Update" -Value "Added"
        $currentuser | add-member -type Noteproperty -name "LastChecked" -Value (Get-Date)

        $updatedUsers += $currentuser
    }
}

$recentlyModifiedUsers += $updatedUsers

$recentlyModifiedUsers | Export-Csv -NoTypeInformation -Encoding utf8 "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\RecentlyUpdatedUsers.csv"
}

Compare-FanDuelUsers


## Remove MSOLUSER

foreach ($user in $FDGROUPMAILBOXES)
{
    Remove-MsolUser -UserPrincipalName $user.userprincipalname -Force
}

## remove MSOLUSER from RecycleBIN

foreach ($user in $FDGROUPMAILBOXES)
{
    Remove-MsolUser -UserPrincipalName $user.userprincipalname -RemoveFromRecycleBin
}

$UpdateAddress =@()
$MissingAddress = @()
foreach ($mailbox in $archivembxs)
{
    if (!(Get-recipient $mailbox -ea silentlycontinue))
    {
        $SplitAddress = $mailbox.Split("@")
        if (!($UpdateAddress = Get-recipient $SplitAddress[0] -ea silentlycontinue))
        {
            Write-Host "No address found for $mailbox" -ForegroundColor Red
            $UpdateAddress += $mailbox
        }
        else
        {
            Write-Host "Address found for $mailbox" -ForegroundColor Green
            $MissingAddress += $mailbox
        }
    }
}


# Match Mailboxes

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $FDAllUsers | ? {$_.FDUserPrincipalName -eq "" -and $_.RecipientTypeDetails -eq "SharedMailbox"}) {
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
        
    }

    $AllUsers += $currentuser
}

## Set PrimarySMTPAddress

$remainingSharedMailboxes = $allUsers

foreach ($sharedmailbox in ${remainingSharedMailboxes})
{
    Write-Host "Setting PrimarySMTPAddress for $($sharedmailbox.DisplayName) to $($sharedmailbox.MatchingPrimarySMTPAddress)" -foregroundcolor cyan -nonewline
    Set-Mailbox $sharedmailbox.DisplayName -WindowsEmailAddress $sharedmailbox.MatchingPrimarySMTPAddress #-whatif
    Write-Host "done" -foregroundcolor green
}

foreach ($sharedmailbox in ${remainingSharedMailboxes})
{
    Write-Host "Setting UPN for $($sharedmailbox.DisplayName) to $($sharedmailbox.MatchingPrimarySMTPAddress)" -foregroundcolor cyan -nonewline
    Set-MsolUserPrincipalName -UserPrincipalName $sharedmailbox.UserPrincipalName -NewUserPrincipalName $sharedmailbox.MatchingPrimarySMTPAddress #-whatif
    Write-Host "done" -foregroundcolor green
}


## Combine full report and matched mailboxes

$FDAllUsers = $FanDuelUsers

$AllUsers = @()
$multiplemailboxes = @()
foreach ($mailbox in $FDAllUsers)
{
    $currentuser = new-object PSObject

    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $mailbox.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $mailbox.userprincipalname
    $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $mailbox.IsLicensed
    $currentuser | add-member -type noteproperty -name "City" -Value $mailbox.City
    $currentuser | add-member -type noteproperty -name "Country" -Value $mailbox.Country
    $currentuser | add-member -type noteproperty -name "Department" -Value $mailbox.Department
    $currentuser | add-member -type noteproperty -name "Fax" -Value $mailbox.Fax
    $currentuser | add-member -type noteproperty -name "FirstName" -Value $mailbox.FirstName
    $currentuser | add-member -type noteproperty -name "LastName" -Value $mailbox.LastName
    $currentuser | add-member -type noteproperty -name "MobilePhone" -Value $mailbox.MobilePhone
    $currentuser | add-member -type noteproperty -name "Office" -Value $mailbox.Office
    $currentuser | add-member -type noteproperty -name "PhoneNumber" -Value $mailbox.PhoneNumber
    $currentuser | add-member -type noteproperty -name "PostalCode" -Value $mailbox.PostalCode
    $currentuser | add-member -type noteproperty -name "State" -Value $mailbox.State
    $currentuser | add-member -type noteproperty -name "StreetAddress" -Value $mailbox.StreetAddress
    $currentuser | add-member -type noteproperty -name "Title" -Value $mailbox.Title    
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $mailbox.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "WhenCreated" -Value $mailbox.WhenCreated    
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value $mailbox.EmailAddresses
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value $mailbox.LegacyExchangeDN
    $currentuser | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value $mailbox.AcceptMessagesOnlyFrom
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value $mailbox.GrantSendOnBehalfTo
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $mailbox.HiddenFromAddressListsEnabled
    $currentuser | add-member -type noteproperty -name "RejectMessagesFrom" -Value $mailbox.RejectMessagesFrom
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $mailbox.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $mailbox.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $mailbox.ForwardingSmtpAddress
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mailbox.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "Alias" -Value $mailbox.alias
    $currentuser | add-member -type noteproperty -name "ExchangeGuid" -Value $mailbox.ExchangeGuid
    $currentuser | Add-Member -type NoteProperty -Name "MBXSize" -Value $mailbox.MBXSize
    $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $mailbox.MBXItemCount
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveGUID" -Value $mailbox.ArchiveGuid
    $currentuser | add-member -type noteproperty -name "ArchiveState" -Value $mailbox.ArchiveState
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $mailbox.ArchiveStatus
    $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $mailbox.ArchiveSize
    $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $mailbox.ArchiveItemCount
    $currentuser | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $mailbox.OneDriveURL
    $currentuser | Add-Member -type NoteProperty -Name "Owner" -Value $mailbox.Owner
    $currentuser | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $mailbox.StorageUsageCurrent
    $currentuser | Add-Member -type NoteProperty -Name "Status" -Value $mailbox.Status
    $currentuser | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $mailbox.SiteDefinedSharingCapability
    $currentuser | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $mailbox.LimitedAccessFileType

    $DisplayName = $mailbox.DisplayName
    $FDUPNUpdate = $mailbox.PrimarySmtpAddress.Replace("@fanduel.com","@fanduelgroup.com")
    

    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline
    
    if ($mailboxcheck = Get-Mailbox $mailbox.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName)
    {
        $MBXStats = Get-MailboxStatistics $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
        $msoluserscheck = get-msoluser -UserPrincipalName $mailboxcheck.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, BlockCredential, UserPrincipalName

        $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $true
        $currentuser | add-member -type noteproperty -name "FDPrimarySMTPAddress" -Value $mailboxcheck.PrimarySmtpAddress
        $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $mailboxcheck.RecipientTypeDetails
        $currentuser | Add-Member -type NoteProperty -Name "FDMBXSize" -Value $MBXStats.TotalItemSize
        $currentuser | Add-Member -Type NoteProperty -name "FDMBXItemCount" -Value $MBXStats.ItemCount
        $currentuser | add-member -type noteproperty -name "FDIsLicensed" -Value $msoluserscheck.IsLicensed
        $currentuser | add-member -type noteproperty -name "FDBlockSigninStatus" -Value $msoluserscheck.BlockCredential
        $currentuser | add-member -type noteproperty -name "FDUserPrincipalName" -Value $msoluserscheck.UserPrincipalName
        Write-Host "found mailbox." -ForegroundColor Green    
    }
        
    
    elseif ($recipientcheck = Get-Recipient $DisplayName)
    {
        if ($recipientcheck.count -gt 1)
        {
            $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value "MultipleRecipientsFound"
            $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $recipientcheck.RecipientTypeDetails
            $currentuser | add-member -type noteproperty -name "FDPrimarySMTPAddress" -Value $mailboxcheck.PrimarySmtpAddress
        }
        else {
            $mailboxcheck2 = Get-Mailbox $recipientcheck.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName
            $msoluserscheck2 = get-msoluser -UserPrincipalName $mailboxcheck2.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, BlockCredential, UserPrincipalName
            $MBXStats = Get-MailboxStatistics $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
         
            $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $true
            $currentuser | Add-Member -type NoteProperty -Name "FDMBXSize" -Value $MBXStats.TotalItemSize
            $currentuser | Add-Member -Type NoteProperty -name "FDMBXItemCount" -Value $MBXStats.ItemCount
            $currentuser | add-member -type noteproperty -name "FDIsLicensed" -Value $msoluserscheck2.IsLicensed
            $currentuser | add-member -type noteproperty -name "FDBlockSigninStatus" -Value $msoluserscheck2.BlockCredential
            $currentuser | add-member -type noteproperty -name "FDUserPrincipalName" -Value $msoluserscheck2.UserPrincipalName
        }
        Write-Host "found recipient." -ForegroundColor Yellow     
    }  
    else
    {
        Write-Host "not found" -ForegroundColor red
        $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False
        $currentuser | Add-Member -type NoteProperty -Name "FDMBXSize" -Value $null
        $currentuser | Add-Member -Type NoteProperty -name "FDMBXItemCount" -Value $null
    }     
    $AllUsers += $currentuser
    #Read-Host "Stopping to check user"
}
$allUsers | Export-Csv -NoTypeInformation -Encoding UTF8 "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

## Combine full report and matched mailboxes 2 **###

$FDAllUsers = $FanDuelUsers

$AllUsers = @()
foreach ($mailbox in $FDAllUsers)
{
    $currentuser = new-object PSObject

    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $mailbox.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $mailbox.userprincipalname
    $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $mailbox.IsLicensed
    $currentuser | add-member -type noteproperty -name "City" -Value $mailbox.City
    $currentuser | add-member -type noteproperty -name "Country" -Value $mailbox.Country
    $currentuser | add-member -type noteproperty -name "Department" -Value $mailbox.Department
    $currentuser | add-member -type noteproperty -name "Fax" -Value $mailbox.Fax
    $currentuser | add-member -type noteproperty -name "FirstName" -Value $mailbox.FirstName
    $currentuser | add-member -type noteproperty -name "LastName" -Value $mailbox.LastName
    $currentuser | add-member -type noteproperty -name "MobilePhone" -Value $mailbox.MobilePhone
    $currentuser | add-member -type noteproperty -name "Office" -Value $mailbox.Office
    $currentuser | add-member -type noteproperty -name "PhoneNumber" -Value $mailbox.PhoneNumber
    $currentuser | add-member -type noteproperty -name "PostalCode" -Value $mailbox.PostalCode
    $currentuser | add-member -type noteproperty -name "State" -Value $mailbox.State
    $currentuser | add-member -type noteproperty -name "StreetAddress" -Value $mailbox.StreetAddress
    $currentuser | add-member -type noteproperty -name "Title" -Value $mailbox.Title    
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $mailbox.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "WhenCreated" -Value $mailbox.WhenCreated    
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value $mailbox.EmailAddresses
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value $mailbox.LegacyExchangeDN
    $currentuser | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value $mailbox.AcceptMessagesOnlyFrom
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value $mailbox.GrantSendOnBehalfTo
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $mailbox.HiddenFromAddressListsEnabled
    $currentuser | add-member -type noteproperty -name "RejectMessagesFrom" -Value $mailbox.RejectMessagesFrom
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $mailbox.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $mailbox.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $mailbox.ForwardingSmtpAddress
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mailbox.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "Alias" -Value $mailbox.alias
    $currentuser | add-member -type noteproperty -name "ExchangeGuid" -Value $mailbox.ExchangeGuid
    $currentuser | Add-Member -type NoteProperty -Name "MBXSize" -Value $mailbox.MBXSize
    $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $mailbox.MBXItemCount
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveGUID" -Value $mailbox.ArchiveGuid
    $currentuser | add-member -type noteproperty -name "ArchiveState" -Value $mailbox.ArchiveState
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $mailbox.ArchiveStatus
    $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $mailbox.ArchiveSize
    $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $mailbox.ArchiveItemCount
    $currentuser | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $mailbox.OneDriveURL
    $currentuser | Add-Member -type NoteProperty -Name "Owner" -Value $mailbox.Owner
    $currentuser | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $mailbox.StorageUsageCurrent
    $currentuser | Add-Member -type NoteProperty -Name "Status" -Value $mailbox.Status
    $currentuser | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $mailbox.SiteDefinedSharingCapability
    $currentuser | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $mailbox.LimitedAccessFileType

    $DisplayName = $mailbox.DisplayName
    $FDUPNUpdate = $mailbox.PrimarySmtpAddress.Replace("@fanduel.com","@fanduelgroup.com")
    

    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline
    $mailboxcheck = Get-Mailbox $mailbox.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName

    if ($mailboxcheck)
    {
       Write-Host "found mailbox." -ForegroundColor Green -nonewline
    }
    elseif ($recipientcheck = Get-Recipient $mailbox.PrimarySmtpAddress -ea silentlycontinue)
    {
        $mailboxcheck = Get-Mailbox $recipientcheck.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName  
        Write-Host "found recipient." -ForegroundColor Yellow -nonewline
    }
    else
    {
        Write-Host "not found" -ForegroundColor red -NoNewline
        $msoluserscheck = @()
        $MBXStats = @()
    }
    if ($mailboxcheck)
    {
        $msoluserscheck = get-msoluser -UserPrincipalName $mailboxcheck.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, BlockCredential, UserPrincipalName
        $MBXStats = Get-MailboxStatistics $mailboxcheck.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
        $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $True
        $currentuser | add-member -type noteproperty -name "FDUserPrincipalName" -Value $msoluserscheck.UserPrincipalName
        $currentuser | add-member -type noteproperty -name "FDIsLicensed" -Value $msoluserscheck.IsLicensed
        $currentuser | add-member -type noteproperty -name "FDBlockSigninStatus" -Value $msoluserscheck.BlockCredential
        $currentuser | add-member -type noteproperty -name "FDPrimarySMTPAddress" -Value $mailboxcheck.PrimarySmtpAddress
        $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $recipientcheck.RecipientTypeDetails   
        $currentuser | Add-Member -type NoteProperty -Name "FDMBXSize" -Value $MBXStats.TotalItemSize
        $currentuser | Add-Member -Type NoteProperty -name "FDMBXItemCount" -Value $MBXStats.ItemCount
    }
    else 
    {
        $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False
        $currentuser | add-member -type noteproperty -name "FDUserPrincipalName" -Value $null
        $currentuser | add-member -type noteproperty -name "FDIsLicensed" -Value $null
        $currentuser | add-member -type noteproperty -name "FDBlockSigninStatus" -Value $null
        $currentuser | add-member -type noteproperty -name "FDPrimarySMTPAddress" -Value $null
        $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $null  
        $currentuser | Add-Member -type NoteProperty -Name "FDMBXSize" -Value $null
        $currentuser | Add-Member -Type NoteProperty -name "FDMBXItemCount" -Value $null
    }

    Write-host " .. done" -foregroundcolor green

    $AllUsers += $currentuser
    #Read-Host "Stopping to check user"
}
$allUsers | Export-Csv -NoTypeInformation "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"


### Shortened Script To Pull Folders for Content Search  #####

$archiveMailboxes = Get-Mailbox -Archive -resultsize Unlimited | ?{$_.primarysmtpaddress -like "*@fanduel.com" -or $_.primarysmtpaddress -like "*@tvg*"}

$folderQueries = @()

foreach ($archivembx in $archiveMailboxes)
{
   # List the folder Ids for the target mailbox
   $emailAddress = $archivembx.primarysmtpaddress

   Write-host "Gathering Folder Details for $($archivembx.primarysmtpaddress) ..." -NoNewline -ForegroundColor Cyan
  
   $folderStatisticsRecoverable = Get-MailboxFolderStatistics $emailAddress -Archive -folderscope recoverable
   foreach ($folderStatistic in $folderStatisticsRecoverable)
   {
      $folderIdentity = $folderStatistic.Identity;
      $folderId = $folderStatistic.FolderId;
      $foldersize = $folderStatistic.foldersize
      $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
      $nibbler= $encoding.GetBytes("0123456789ABCDEF");
      $folderIdBytes = [Convert]::FromBase64String($folderId);
      $indexIdBytes = New-Object byte[] 48;
      $indexIdIdx=0;
      $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
      $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";

      $folderStat = New-Object PSObject
      $folderStat | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $folderIdentity 
      $folderStat | Add-Member -MemberType NoteProperty -Name "FolderSize" -Value $foldersize
      $folderStat | Add-Member -MemberType NoteProperty -Name "FolderQuery" -Value $folderQuery
         
      $folderQueries += $folderStat
   }
   Write-host "done" -ForegroundColor Green
}

#$FDGROUPAzureUsers = $importCsv | foreach {Get-AzureADUser -SearchString $_} | select ObjectID, DisplayName, UserPrincipalName

$FDGROUPAzureUsers = $needToUpdate | foreach {Get-AzureADUser -SearchString $_} | select ObjectID, DisplayName, UserPrincipalName

##add members to Office 365 Licensing Group
$O365LicenseAzureGroup = Get-AzureADGroup -SearchString O365_E3_License
$addUsersToO365LicenseGroup = Import-Csv $filepath
 (CSV Needs DisplayName heading)

Write-host "Checking member status" -NoNewline -foregroundcolor cyan

foreach ($member in $addUsersToO365LicenseGroup) 
{
    if ($azureUserCheck = Get-AzureADUser -SearchString $member.DisplayName -ea silentlycontinue | select ObjectID, DisplayName, UserPrincipalName)
    {
        if (!(Get-AzureADGroupmember -ObjectId $O365LicenseAzureGroup.ObjectID -all $true -ea silentlycontinue | ? {$_.DisplayName -eq $azureUserCheck.DisplayName}))
        {
            Add-AzureADGroupMember -ObjectId $O365LicenseAzureGroup.ObjectID -RefObjectId $azureUserCheck.ObjectID
            Write-Host "Added member" $azureUserCheck.DisplayName -foregroundcolor green
        }
        else
        {
            Write-host ". " -NoNewline
        }
    }
    else
    {
        Write-Host "No user found for $($member.DisplayName)" -ForegroundColor Red
    }
}
# Remove users from group
$removelicenses = Get-Content "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\RemoveLicenses.txt"

$FoxBetAzureUsers = $removelicenses | foreach {Get-AzureADUser -SearchString $_} | select ObjectID, DisplayName, UserPrincipalName

Write-host "Checking member status" -NoNewline -foregroundcolor cyan


foreach ($member in $FoxBetAzureUsers) 
{
    if (Get-AzureADGroupmember -ObjectId ee233c80-a51e-4785-900b-ef2eb3a530e6 -all $true| ? {$_.DisplayName -eq $member.DisplayName} -ea silentlycontinue)
    {
        Remove-AzureADGroupMember -ObjectId ee233c80-a51e-4785-900b-ef2eb3a530e6 -MemberId $member.ObjectID
        Write-Host "Removed member" $member.DisplayName -foregroundcolor green -NoNewline
    }
    else
    {
        Write-host ". " -NoNewline
    }
}

## Remove Users from Office365 License Group
$O365LicenseAzureGroup = Get-AzureADGroup -SearchString O365_E3_License
$removeUsersToO365LicenseGroup = Import-Csv $filepath
 (CSV Needs DisplayName heading)

Write-host "Checking member status" -NoNewline -foregroundcolor cyan

foreach ($member in $removeUsersToO365LicenseGroup) 
{
    if ($azureUserCheck = Get-AzureADUser -SearchString $member.DisplayName -ea silentlycontinue | select ObjectID, DisplayName, UserPrincipalName)
    {
        if (Get-AzureADGroupmember -ObjectId $O365LicenseAzureGroup.ObjectID -all $true -ea silentlycontinue | ? {$_.DisplayName -eq $azureUserCheck.DisplayName})
        {
            Remove-AzureADGroupMember -ObjectId ee233c80-a51e-4785-900b-ef2eb3a530e6 -MemberId $member.ObjectID
            Write-Host "Removed member" $member.DisplayName -foregroundcolor green -NoNewline
        }
        else
        {
            Write-host ". " -NoNewline
        }
    }
    else
    {
        Write-Host "No user found for $($member.DisplayName)" -ForegroundColor Red
    }
}

#Remove specific email domain from Remote Mailboxes
$RemoveSMTPDomain1 = "*tvg.a-a-ron.org*"
$RemoveSMTPDomain2 = "*fanduel.a-a-ron.org*"
$domain = "a-a-ron.org"

$AllMailboxes = Get-Mailbox -OrganizationalUnit $domain | Where-Object {($_.EmailAddresses -like $RemoveSMTPDomain1 -or $_.EmailAddresses -like $RemoveSMTPDomain2)}
ForEach ($Mailbox in $AllMailboxes)
{
   $UPN = $Mailbox.UserPrincipalName
   $PrimarySMTPAddress = $Mailbox.PrimarySmtpAddress
   $EmailAddresses = ($Mailbox.EmailAddresses).SMTPAddress
   Write-Host "Disabling Email Address Policy for $UPN" -ForegroundColor Blue
   Set-Mailbox -Identity $UPN -EmailAddressPolicyEnabled $False -erroraction SilentlyContinue
   if($PrimarySMTPAddress -notlike $NewPrimarySMTP)
   {
        Write-Host "Setting $PrimarySMTPAddress to $UPN" -ForegroundColor Green
        Set-Mailbox -Identity $UPN -PrimarySMTPAddress $UPN 
   }
   ForEach ($Address in $EmailAddresses)
   {
        if($Address -like $RemoveSMTPDomain1 -or $Address -like $RemoveSMTPDomain2)
        {
            Write-Host "Removing $address from $UPN"-ForegroundColor DarkCyan
            Set-Mailbox $UPN -EmailAddresses @{remove="$Address"}
        }

   }
}


## check updated users exist

foreach ($user in $recentlyupdatedusers | ?{$_.Update -eq "Added"})
{
    $DisplayName = $user.DisplayName
    if($msolcheck = Get-MsolUser -searchstring $DisplayName)
    {
        $foundusers += $msolcheck
        Write-Host "Found user $($DisplayName)" -foregroundcolor green
    }
    else
    {
        $notfoundusers += $user
        Write-Host "Did not find user $($DisplayName)" -foregroundcolor red
    }
}

foreach ($user in $foxbetmsolusercheck)
{
    Write-Host "Updating User "$user.DisplayName" licenses .." -nonewline -foregroundcolor cyan
    #$licenseSkus = $user.licenses.accountskuid
    #Write-host "found "$licenseSkus.count". Removing ..." -nonewline -foregroundcolor yellow
    Set-msoluser -UserPrincipalName $user.UserPrincipalName -usagelocation US
    Start-Sleep -seconds 5

    foreach (${license} in ${licenses})
    {    
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $license -ea silentlycontinue
        Write-host "done" -foregroundcolor green -nonewline
    }
    Write-host ""
    #Read-Host "pause to check"
}

#ReplaceUPN (not during cutover)
foreach ($user in $foxbetmsolusercheck| sort userprincipalname)
{
	$newUPN = $user.UserPrincipalName.replace("@fanduelgroup.onmicrosoft.com","@fanduelgroup.com")
    Write-Host "Updating UPN for $($user.displayname) to $newupn .." -ForegroundColor Cyan -NoNewline
    $msolcheck = get-msoluser -userprincipalname $newUPN -ea silentlycontinue

    if ($msolcheck -ne $user.userprincipalname)
    {
        Set-MsolUserPrincipalName -UserPrincipalName $user.UserPrincipalName -NewUserPrincipalName $newUPN
	    Write-Host "done" -ForegroundColor Green
    }
    else
    {
        Write-host "don't need to update UPN." -foregroundcolor darkgreen
    }
}

# Remove users from group
Remove-AzureADGroupMember -ObjectId ee233c80-a51e-4785-900b-ef2eb3a530e6 -memberId (Get-AzureADUser -SearchString (Read-Host -Prompt "User")).objectid
Remove-AzureADGroupMember -ObjectId ee233c80-a51e-4785-900b-ef2eb3a530e6 -memberId (Get-AzureADUser -SearchString "tamas millian").objectid

#Add Users to group
Add-AzureADGroupMember -ObjectId ee233c80-a51e-4785-900b-ef2eb3a530e6 -RefObjectId (Get-AzureADUser -SearchString (Read-Host -Prompt "User")).objectid
Add-AzureADGroupMember -ObjectId ee233c80-a51e-4785-900b-ef2eb3a530e6 -RefObjectId (Get-AzureADUser -SearchString "tamas millian").objectid
# Add x500
$ALLFDMailboxes = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

foreach ($mbx in $ALLFDMailboxes)
{
    if ($mailboxcheck = get-mailbox $mbx.FDPrimarySMTPAddress -ea silentlycontinue)
    {
        Write-Host "Adding X500 to mailbox $($mbx.DisplayName)" -ForegroundColor Cyan -NoNewline
        $newX500 = $mbx.LegacyExchangeDN
        Set-Mailbox $mbx.FDPrimarySMTPAddress -EmailAddresses @{add=$newX500}
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
}

# Add Alternate EmailAddress to Mailboxes during cutover

$ALLFDMailboxes = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

foreach ($mbx in $ALLFDMailboxes | ?{$_.ExistsOnDestinationTenant -eq $true})
{
    if ($mailboxcheck = get-mailbox $mbx.DisplayName -ea silentlycontinue)
    {
        Write-Host "Adding EmailAddresses to mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        [array]$EmailAddresses = $mbx.EmailAddresses -split ","
            foreach ($altAddress in ($EmailAddresses |  Where {($_-like "*fanduel.com" -or $_ -like "*@tvg*" -or $_ -like "x500*")}))
            {
                Write-Host "." -ForegroundColor DarkGreen -NoNewline
                Set-Mailbox $mbx.DisplayName -EmailAddresses @{add=$altAddress} -wa silentlycontinue
            }
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
    #Read-Host "Pause to check"
}

##

$DGSettings = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

foreach ($DL in $DGSettings | ?{$_.ExistsOnDestinationTenant -eq $true})
{
    if ($dlcheck = Get-DistributionGroup $DL.PrimarySMTPAddress -ea silentlycontinue)
    {
        Write-Host "Adding EmailAddresses to DL $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        [array]$EmailAddresses = $DL.EmailAddresses -split ";"
            foreach ($altAddress in ($EmailAddresses |  Where {($_-like "*fanduel.com" -or $_ -like "*@tvg*" -or $_ -like "x500*")}))
            {
                Write-Host "." -ForegroundColor DarkGreen -NoNewline
                Set-DistributionGroup $DL.PrimarySMTPAddress -EmailAddresses @{add=$altAddress}
            }
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No DL found for $($DL.DisplayName)" -ForegroundColor Red
    }
    Read-Host "Pause to check"
}

# Update UPN and primary smtp address from FanDuelGroup.com to fanduel.com

$ALLFDMailboxes = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

$notfoundusers = @()
foreach ($mbx in $ALLFDMailboxes[0] | ?{$_.ExistsOnDestinationTenant -eq $true}) 
{
    if ($msolcheck = Get-msoluser -userprincipalname $mbx.FDUserPrincipalName)
    {
        Write-Host "Updating UPN for $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        $FDUPNUpdate = $mailbox.FDUserPrincipalName.Replace("@fanduelgroup.com","@fanduel.com")
        
        Set-MsolUserPrincipalName -userprincipalname $mbx.FDUserPrincipalName -NewUserPrincipalName $FDUPNUpdate
    }
        Write-Host "done" -ForegroundColor Cyan
    else 
    {
        Write-Host "No MSOLuser found for $($mbx.DisplayName)" -ForegroundColor Red
        $notfoundusers += $mbx
    }
}

# Update UPN from FanDuelGroup.com to fox.bet

$updateUPNUserList = Import-Csv (filepath)
#(CSV Needs DisplayName heading)
$notfoundusers = @()
foreach ($mbx in $updateUPNUserList) 
{
    if ($msolcheck = Get-msoluser -userprincipalname $mbx.UserPrincipalName)
    {
        Write-Host "Updating UPN for $($msolcheck.DisplayName) " -ForegroundColor Cyan -NoNewline
        $UPNUpdate = $msolcheck.UserPrincipalName.Replace("@fanduelgroup.com","@fox.bet")
        
        Set-MsolUserPrincipalName -userprincipalname $msolcheck.UserPrincipalName -NewUserPrincipalName $UPNUpdate
    }
        Write-Host "done" -ForegroundColor Cyan
    else 
    {
        Write-Host "No MSOLuser found for $($mbx.DisplayName)" -ForegroundColor Red
        $notfoundusers += $mbx
    }
}
# Update primary smtp address from FanDuelGroup.com to fox.bet
$updateMailboxUserList = Import-Csv (filepath)
(CSV Needs EmailAddress and heading)
$notfoundusers = @()
foreach ($mbx in $updateMailboxUserList)
{
    if ($mailboxcheck = get-mailbox $mbx.EmailAddress -ea silentlycontinue)
    {
        Write-Host "Updating PrimarySMTPAddress for $($mailboxcheck.DisplayName) " -ForegroundColor Cyan -NoNewline
        Set-Mailbox $mailboxcheck.DisplayName -WindowsEmailAddress  $mbx.EmailAddress
        Write-Host "done" -ForegroundColor Cyan
    }
    else 
    {
        Write-Host "No mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
        $notfoundusers += $mbx
    }
}

#####
foreach ($user in $remainingmsolusers | sort displayname)
{
    Write-Host "Updating UPN for $($user.DisplayName) " -ForegroundColor Cyan -NoNewline
    $FDUPNUpdate = $user.UserPrincipalName.Replace("@fanduelgroup.onmicrosoft.com","@fanduel.com")
    Set-MsolUserPrincipalName -userprincipalname $user.UserPrincipalName -NewUserPrincipalName $FDUPNUpdate
    Write-Host "done." -ForegroundColor Green
}

$notfoundusers = @()
foreach ($mbx in $ALLFDMailboxes | ?{$_.ExistsOnDestinationTenant -eq $true}) 
{
    if ($msolcheck = Get-msoluser -userprincipalname $mbx.FDUserPrincipalName)
    {
        Write-Host "Updating UPN for $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        $FDUPNUpdate = $mailbox.FDUserPrincipalName.Replace("@fanduelgroup.com","@fanduel.com")
        
        Set-MsolUserPrincipalName -userprincipalname $mbx.FDUserPrincipalName -NewUserPrincipalName $FDUPNUpdate
    }
        Write-Host "done" -ForegroundColor Cyan
    else 
    {
        Write-Host "No MSOLuser found for $($mbx.DisplayName)" -ForegroundColor Red
        $notfoundusers += $mbx
    }
    
}

$notfoundusers = @()
foreach ($mbx in $ALLFDMailboxes | ?{$_.ExistsOnDestinationTenant -eq $true})
{
    if ($mailboxcheck = get-mailbox $mbx.FDPrimarySMTPAddress -ea silentlycontinue)
    {
        Write-Host "Updating PrimarySMTPAddress for $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        Set-Mailbox $mbx.DisplayName -WindowsEmailAddress  $mbx.PrimarySmtpAddress
        Write-Host "done" -ForegroundColor Cyan
    }
    else 
    {
        Write-Host "No mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
        $notfoundusers += $mbx
    }
    Read-Host "Pause to check"
}

# Set Up Forwarding on PPB tenant.

$ALLFDMailboxes = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

foreach ($mbx in $ALLFDMailboxes)
{
    if ($mailboxcheck = get-mailbox $mbx.UserPrincipalName -ea silentlycontinue)
    {
        Write-Host "Updating Forwarding for mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        $mailboxcheck | Set-Mailbox -forwardingsmtpaddress $mbx.PrimarySMTPAddress
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
    #Read-Host " Pause to check"
}

# Set Up Forwarding on FD tenant.

$ALLFDMailboxes = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

foreach ($mbx in $ALLFDMailboxes | ? {$_.ForwardingAddress})
{
    if ($mailboxcheck = get-mailbox $mbx.DisplayName -ea silentlycontinue)
    {
        [boolean]$DeliverToMailboxAndForward = [boolean]::Parse($mbx.DeliverToMailboxAndForward)
        Write-Host "Updating Forwarding for mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        $mailboxcheck | Set-Mailbox -forwardingaddress $mbx.forwardingaddress -DeliverToMailboxAndForward:$DeliverToMailboxAndForward       
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
    #Read-Host " Pause to check"
}

# Hide from GAL

$ALLFDMailboxes = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

foreach ($mbx in $ALLFDMailboxes | ? {$_.HiddenFromAddressListsEnabled -eq $true})
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

### Apply Calendar Permissions

foreach ($CalendarPerm in $AllCalendarPerms)
{
    Write-Host "Updating $($CalendarPerm.CalendarPath). Adding Calendar perms for $($CalendarPerm.User) ..." -NoNewline
    Add-MailboxFolderPermission -Identity $CalendarPerm.CalendarPath -User $CalendarPerm.User -AccessRights $CalendarPerm.AccessRights -ea silentlycontinue
    Write-Host "done" -ForegroundColor Green
}

#Set GrandSendOnBehalf

$ALLFDMailboxes = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\FanDuelFullDetailsMatched.csv"

foreach ($mbx in $ALLFDMailboxes | ?{$_.GrantSendOnBehalfTo})
{
    $
    if ($mailboxcheck = get-mailbox $mbx.display -ea silentlycontinue)
    {
        $grantsendonbehalf = $mbx.GrantSendOnBehalfTo
        Write-Host "Updating GrantSendOnBehalfTo for mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        $mailboxcheck | Set-Mailbox  -GrantSendOnBehalfTo @(add=$grantsendonbehalf)
        Write-Host "done" -ForegroundColor Cyan
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
}

# Grant SendAs Perms

foreach ($sendasperm in $AllSendAsPerms | ? {$_.ObjectWithSendAs})
{
    $recipientaddressSplit = $sendasperm.Recipient -split "@"
    $objectwithSendasSplit = $sendasperm.ObjectWithSendAs -split "@"
    $objectwithSendAsDestination = Get-recipient $objectwithSendasSplit[0]

    if ($recipientcheck = Get-recipient $recipientaddressSplit[0] -ea silentlycontinue)
    {
        $grantsendonbehalf = $mbx.GrantSendOnBehalfTo
        Write-Host "Updating SendAs for recipient $($recipientcheck.displayName) " -ForegroundColor Cyan -NoNewline
        Add-RecipientPermission $recipientcheck.displayName -trustee $objectwithSendAsDestination.displayname -AccessRights SendAs -Confirm:$false
        Write-Host "done" -ForegroundColor Cyan
    }
    else {
        Write-host "No User found for $($sendasperm.Recipient)" -foregroundcolor red
    }
}

####
$AllMissingUsersDetails = @()
foreach ($mbx in $missingusers)
{
    if ($msoluserscheck = get-msoluser -searchstring $mbx -ea silentlycontinue | select DisplayName, IsLicensed, BlockCredential, UserPrincipalName)
    {
        if ($msoluserscheck.count -gt 1)
            {
                Write-Host "Multiple Users found skipping." -ForegroundColor Yellow
                $multiplemailboxes += $mailbox
                $currentuser | add-member -type noteproperty -name "ExistsOnDestinationTenant" -Value "Multiple Users found"
            }
        else
        {
            $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $true
            $currentuser | add-member -type noteproperty -name "FDIsLicensed" -Value $msoluserscheck.IsLicensed
            $currentuser | add-member -type noteproperty -name "FDBlockSigninStatus" -Value $msoluserscheck.BlockCredential
            $currentuser | add-member -type noteproperty -name "FDUserPrincipalName" -Value $msoluserscheck.UserPrincipalName
            

            if ($mailboxcheck = Get-Mailbox $msoluserscheck.UserPrincipalName -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails)
            {
                $MBXStats = Get-MailboxStatistics $mailboxcheck.primarysmtpaddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount

                $currentuser | add-member -type noteproperty -name "FDPrimarySMTPAddress" -Value $mailboxcheck.PrimarySmtpAddress
                $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $mailboxcheck.RecipientTypeDetails
                $currentuser | Add-Member -type NoteProperty -Name "FDMBXSize" -Value $MBXStats.TotalItemSize
                $currentuser | Add-Member -Type NoteProperty -name "FDMBXItemCount" -Value $MBXStats.ItemCount

                Write-Host "found." -ForegroundColor Green
            }
            
        }
    }
    elseif ($recipientcheck = Get-Recipient $mbx)
            {
                $msolcheck2 = get-msoluser -searchstring $recipientcheck.userprincipalname -ea silentlycontinue | select DisplayName, IsLicensed, BlockCredential, UserPrincipalName
                $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $true
                $currentuser | add-member -type noteproperty -name "FDPrimarySMTPAddress" -Value $recipientcheck.PrimarySmtpAddress
                $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $recipientcheck.RecipientTypeDetails                
            }

    else
    {
        Write-Host "not found" -ForegroundColor red
        $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False
        $currentuser | Add-Member -type NoteProperty -Name "FDMBXSize" -Value ""
        $currentuser | Add-Member -Type NoteProperty -name "FDMBXItemCount" -Value ""
    }     
    $AllMissingUsersDetails += $currentuser
}

##
$AllMissingUsersDetails = @()
foreach ($mbx in $missingusers)
{
    if ($msoluserscheck = get-msoluser -searchstring $mbx -ea silentlycontinue | select DisplayName, IsLicensed, BlockCredential, UserPrincipalName)
    {
        $currentuser = new-object PSObject
        $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $true

        $currentuser | add-member -type noteproperty -name "FDIsLicensed" -Value $msoluserscheck.IsLicensed
        $currentuser | add-member -type noteproperty -name "FDBlockSigninStatus" -Value $msoluserscheck.BlockCredential
        $currentuser | add-member -type noteproperty -name "FDUserPrincipalName" -Value $msoluserscheck.UserPrincipalName
        

        if ($mailboxcheck = Get-Mailbox $msoluserscheck.UserPrincipalName -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails)
        {
            $MBXStats = Get-MailboxStatistics $mailboxcheck.primarysmtpaddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount

            $currentuser | add-member -type noteproperty -name "FDPrimarySMTPAddress" -Value $mailboxcheck.PrimarySmtpAddress
            $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $mailboxcheck.RecipientTypeDetails
            $currentuser | Add-Member -type NoteProperty -Name "FDMBXSize" -Value $MBXStats.TotalItemSize
            $currentuser | Add-Member -Type NoteProperty -name "FDMBXItemCount" -Value $MBXStats.ItemCount

            Write-Host "found." -ForegroundColor Green
        }
    }
    $AllMissingUsersDetails += $currentuser
}


foreach ($confroom in $AllConferenceRooms | ?{$_.DisplayName})
{
    #create the Conference Room
    $alias = $confroom.EmailAddress -split "@"
    $FanDuelGroupAddress = $alias[0] + "@fanduelgroup.com"
    Write-host "Adding $($confroom.DisplayName) .." -foregroundcolor cyan -nonewline
    New-Mailbox -Room -PrimarySMTPAddress $FanDuelGroupAddress -Alias $alias[0] -Name $confroom.ZoomRoomName

    Start-Sleep -seconds 5

    #Set Mailbox Capacity
    Write-host "Setting Mailbox Capacity .. " -foregroundcolor cyan -nonewline
    Set-Mailbox $confroom.ZoomRoomName -ResourceCapacity $confroom.Capacity
    Write-host "done " -foregroundcolor green
}

#### Check Mailbox list to mig
$migmigrationmailboxcheck1Results = @()
foreach ($mbx in $migmigrationmailboxcheck1)
{
    if ($matcheduser = $allfanduelmatched | ?{$_.FDPrimarySMTPAddress -like "*$mbx*"})
    {
        $currentuser = new-object PSObject
        $currentuser | Add-Member -type NoteProperty -Name "MBXNameCheck" -Value $mbx
        $currentuser | Add-Member -type NoteProperty -Name "DisplayName" -Value $matcheduser.DisplayName
        $currentuser | Add-Member -type NoteProperty -Name "ExportEmailAddress" -Value $matcheduser.UserPrincipalName
        $currentuser | add-member -type noteproperty -name "ImportEmailAddress" -Value $matcheduser.FDPrimarySMTPAddress
        $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $matcheduser.FDRecipientTypeDetails
        $currentuser | add-member -type noteproperty -name "SRCMBXSize" -Value $matcheduser.MBXSize
        $currentuser | add-member -type noteproperty -name "SRCMBXItemCount" -Value $matcheduser.MBXItemCount

        $MBXStats = Get-MailboxStatistics $matcheduser.FDPrimarySMTPAddress | select TotalItemSize, ItemCount

        $currentuser | add-member -type noteproperty -name "FDMBXSize" -Value $MBXStats.TotalItemSize
        $currentuser | add-member -type noteproperty -name "FDMBXItemCount" -Value $MBXStats.ItemCount

        $migmigrationmailboxcheck1Results += $currentuser
    }
}


#### Check Mig Project to mailbox list
$notfoundMatchedUsers = @()
$migmigrationmailboxcheck2Results = @()
foreach ($mbx in $migmigrationmailboxcheck2)
{
    if ($matcheduser = $mailboxmigrationstats | ?{$_.DestinationEmailAddress -like "*$mbx*" -and $_.SourceType -eq "Mailbox"})
    {
        if ($MatchedUser2 = $allfanduelmatched | ?{$_.UserPrincipalName -eq $matcheduser.SourceEmailAddress})
        {
            $currentuser = new-object PSObject
            $currentuser | Add-Member -type NoteProperty -Name "MBXNameCheck" -Value $mbx
            $currentuser | Add-Member -type NoteProperty -Name "DisplayName" -Value $MatchedUser2.DisplayName
            $currentuser | Add-Member -type NoteProperty -Name "ExportEmailAddress" -Value $matcheduser.SourceEmailAddress
            $currentuser | add-member -type noteproperty -name "ImportEmailAddress" -Value $matcheduser.DestinationEmailAddress
            $currentuser | add-member -type noteproperty -name "Project" -Value $matcheduser.Project
            $currentuser | add-member -type noteproperty -name "FDPrimarySMTPAddress" -Value $MatchedUser2.FDPrimarySMTPAddress
            $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $MatchedUser2.IsLicensed
            $currentuser | add-member -type noteproperty -name "FDRecipientTypeDetails" -Value $MatchedUser2.FDRecipientTypeDetails
            $currentuser | add-member -type noteproperty -name "MBXSize" -Value $MatchedUser2.MBXSize
            $currentuser | add-member -type noteproperty -name "MBXItemCount" -Value $MatchedUser2.MBXItemCount

            $MBXStats = Get-MailboxStatistics $MatchedUser2.FDPrimarySMTPAddress | select TotalItemSize, ItemCount

            $currentuser | add-member -type noteproperty -name "FDMBXSize" -Value $MBXStats.TotalItemSize
            $currentuser | add-member -type noteproperty -name "FDMBXItemCount" -Value $MBXStats.ItemCount

            $migmigrationmailboxcheck2Results += $currentuser
        }
        else 
        {
            Write-Host "no user found for $($matcheduser.SourceEmailAddress)"
            $notfoundMatchedUsers += $matcheduser
        }
    }
    else 
    {
        Write-Host "no user found for $($mbx)"
        $notfoundMatchedUsers += $mbx
    }
}


### Update PrimarySMTP

foreach ($dl in $dls)
{
    $FDGEBMAILAddress = $dl.primarysmtpaddress -split "@"
    $PPBEmailAddress = $FDGEBMAILAddress[0] + "@paddypowerbetfair.com"
    Set-DistributionGroup $dl.name -primarysmtpaddress $PPBEmailAddress
}


#Create DistributionGroups

foreach ($dl in $allDLs | ? {$_.ExistsOnO365 -eq $False})
{
    #Creating DL
    Write-host "DL $($dl.DisplayName) .. " -foregroundcolor cyan -nonewline
    if (!(Get-DistributionGroup $dl.DisplayName -ea silentlycontinue))
    {
        [boolean]$RequireSenderAuthenticationEnabled = [boolean]::Parse($dl.RequireSenderAuthenticationEnabled)
        New-DistributionGroup -name $dl.DisplayName -alias $dl.alias -DisplayName $dl.displayname -PrimarySmtpAddress $dl.primarysmtpaddress -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled
        Start-Sleep -seconds 5
        
    }
    else
    {
        Write-host "DL already exists. Skipping .. " -foregroundcolor yellow
    }
    #Add X500
    Write-Host "Adding X500 .. " -ForegroundColor Cyan -NoNewline
    $newX500 = "x500:" +$dl.LegacyExchangeDN
    Set-DistributionGroup $dl.DisplayName -EmailAddresses @{add=$newX500}
    
    #Add DL EmailAddresses
    [array]$EmailAddresses = $dl.EmailAddresses -split ";"
    Write-Host "Adding AltEmailAddresses" -ForegroundColor DarkGreen -NoNewline
    foreach ($altAddress in ($EmailAddresses |  Where {($_-like "*@paddypowerbetfair.com" -or $_ -like "x500*")}))
    {
        if ($altAddress -like "*paddypowerbetfair.com")
        {
            $Fanduelemailaddressupdate = $altAddress.replace("@paddypowerbetfair.com","@fanduel.com")
            $tvgemailaddressupdate = $altAddress.replace("@paddypowerbetfair.com","@tvg.com")
            $tvgnetworkemailaddressupdate = $altAddress.replace("@paddypowerbetfair.com","@tvgnetwork.com")
            Write-Host "." -ForegroundColor DarkGreen -NoNewline
            Set-DistributionGroup $dl.DisplayName -EmailAddresses @{add=$Fanduelemailaddressupdate} -wa silentlycontinue
            Set-DistributionGroup $dl.DisplayName -EmailAddresses @{add=$tvgemailaddressupdate} -wa silentlycontinue
            Set-DistributionGroup $dl.DisplayName -EmailAddresses @{add=$tvgnetworkemailaddressupdate} -wa silentlycontinue
        }
        else
        {
            Write-Host "." -ForegroundColor DarkGreen -NoNewline
            Set-DistributionGroup $dl.DisplayName -EmailAddresses @{add=$altAddress} -wa silentlycontinue 
        }
    }
    Write-Host ". done" -ForegroundColor Cyan
    Read-Host "pause to check"     
}

#update primary smtp address

$notfoundusers = @()
foreach ($mbx in $AllUsers | ?{$_.ExistsOnDestinationTenant -eq $false}) 
{
    if ($msolcheck = Get-msoluser -searchstring $mbx.DisplayName)
    {
        Write-Host "Adding EmailAddress for $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        
        $EmailAddress = "smtp:" + $mbx.PrimarySMTPAddress
        Set-Mailbox $msolcheck.UserPrincipalName -EmailAddresses @{add=$EmailAddress}
    }
        Write-Host "done" -ForegroundColor Cyan
    else{
        Write-Host "No MSOLuser found for $($mbx.DisplayName)" -ForegroundColor Red
        $notfoundusers += $mbx
    }
}

# Check recently updated users added
$recentlyupdated = import-csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\RecentlyUpdatedUsers.csv"
$migrationwizmailboxes = import-csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\MailboxMigrationStatistics.csv"
$foundmigrations =@()
$notfoundmigrations = @()
foreach ($user in $recentlyupdated | ? {$_.Update -eq "Added"})
{
    $UPNSplit = $user.UserPrincipalName -split "@"
    $UPNPrefix = $UPNSplit[0]
    $migrationcheck = $migrationwizmailboxes | ? {$_.SourceEmailAddress -like "*$UPNPrefix*"}
    if ($migrationcheck)
    {
        Write-Host "$($user.DisplayName) migration project found" -foregroundcolor green
        $foundmigrations += $migrationcheck
    }
    else {
        Write-Host "$($user.DisplayName) migration project not found" -foregroundcolor red
        $notfoundmigrations += $user
    }
}
 # check in tenant
 $foundmailbox = @()
 $notfoundmailbox =@()
foreach ($missinguser in $notfoundmigrations)
{
    $mailboxcheck = Get-Mailbox $missinguser.DisplayName
    if ($mailboxcheck)
    {
        Write-Host "$($missinguser.DisplayName) mailbox found" -foregroundcolor green
        $foundmailbox += $mailboxcheck
    }
    else {
        $notfoundmailbox += $missinguser
        Write-Host "$($missinguser.DisplayName) not found" -foregroundcolor red
    }
}

#

$foundmigrations =@()
$notfoundmigrations = @()
foreach ($user in $allremainingmigmailboxes)
{
    $UPNSplit = $user.FDPrimarySmtpAddress -split "@"
    $UPNPrefix = $UPNSplit[0]
    $migrationcheck = $migrationwizmailboxes | ? {$_.DestinationEmailAddress -like "*$UPNPrefix*"}
    if ($migrationcheck)
    {
        Write-Host "$($user.FDName) migration project found" -foregroundcolor green
        $foundmigrations += $migrationcheck
    }
    else {
        Write-Host "$($user.FDName) migration project not found" -foregroundcolor red
        $notfoundmigrations += $user
    }
}

$foundmigrations =@()
$notfoundmigrations2 = @()
foreach ($user in $notfoundmigrations)
{
    $UPNSplit = $user.FDPrimarySmtpAddress -split "@"
    $UPNPrefix = $UPNSplit[0]
    $UPNPrefixSplit = $UPNPrefix -split ","
    $UPNPrefix2 = $UPNPrefixSplit[1]
    $migrationcheck = $migrationwizmailboxes | ? {$_.DestinationEmailAddress -like "*$UPNPrefix2*"}
    if ($migrationcheck)
    {
        Write-Host "$($user.FDName) migration project found" -foregroundcolor green
        $foundmigrations += $migrationcheck
    }
    else {
        Write-Host "$($user.FDName) migration project not found" -foregroundcolor red
        $notfoundmigrations2 += $user
    }
}
$AddedTVGUsers = @()
foreach ($mailbox in $allmailboxes)
{
    $UPNSplit = $mailbox.PrimarySMTPAddress -split "@"
    $TVGAddress = $UPNSplit[0] + "@tvg.com"
    if (!($mailbox | ?{$_.EmailAddresses -like "*$TVGAddress*"}))
    {
        $mailbox | Set-Mailbox -EmailAddresses @{add=$TVGAddress}
        Write-host "Updated $($mailbox.DisplayName) with TVG EmailAddress" -foregroundcolor green
        $AddedTVGUsers += $mailbox
    }
} 

$AddedTVGUsers2 = @()
Foreach ($user in $allmailboxes |? {$_.DisplayName -notlike "adm -*"}) {
    $TVGAddress2 = @()
    $mailboxcheck = Get-Mailbox $user.DisplayName | select emailaddresses
    $Msoluser = get-msoluser -userprincipalname $user.userprincipalname
    $TVGAddress2 = $user.DisplayName[0] + $Msoluser.LastName + "@tvg.com"
    if (!($mailboxcheck | ? {$_.EmailAddresses -like "*$TVGAddress2*"}))
        {
            Set-Mailbox $user.DisplayName -EmailAddresses @{add=$TVGAddress2}
            Write-host "Updated $($user.DisplayName) with TVG EmailAddress" -foregroundcolor green
            $AddedTVGUsers2 += $user
        }
    else {
        Write-Host "Already Updated $($user.DisplayName). No address to add" -foregroundcolor yellow
    }
}


foreach ($Mailbox in $AllFDEmailAddresses)
{
    if ($mailboxcheck = Get-Mailbox $Mailbox.DisplayName -ea silentlycontinue)
    {
        [array]$EmailAddresses = $Mailbox.EmailAddresses -split ";"
            foreach ($altAddress in ($EmailAddresses |  Where {($_-like "*@tvg*")}))
            {
                Write-Host "Adding $($altAddress) EmailAddresses to DL $($Mailbox.DisplayName) " -ForegroundColor Cyan -NoNewline
                Write-Host "." -ForegroundColor DarkGreen -NoNewline
                Set-Mailbox $mailboxcheck.PrimarySMTPAddress -EmailAddresses @{add=$altAddress}
            }
        Write-Host "done" -ForegroundColor Cyan
    }
    Read-Host "Pause to check"
}


foreach ($mailbox in $updateFDmailboxes)
{
    $FDGEBMAILAddress = $mailbox.primarysmtpaddress -split "@"
    $FDMAILAddress = $FDGEBMAILAddress[0] + "@fanduel.com"
    Set-Mailbox $mailbox.DisplayName -WindowsEmailAddress $FDMAILAddress
}


### Add FD mailboxes to group

$notfoundusers = @()
Write-host "Adding Members to DL TEMPFDMailboxes" -fore cyan -nonewline
foreach ($fdmailbox in $allFDMailboxes)
{
    $memberupdate = $fdmailbox.PrimarySMTPAddress.Replace("@fanduel.com","@paddypowerbetfair.com")
    if ($recipientcheck = Get-recipient $memberupdate -ea silentlycontinue)
    {
        Write-Host ". " -fore green -nonewline
        Add-DistributionGroupMember -Identity "TEMPFDMailboxes" -Member $recipientcheck.primarysmtpaddress 
    }
    else
    {
        Write-Host "no user found for $($fdmailbox.DisplayName)" -fore red -nonewline
        $notfoundusers += $fdmailbox
    }
}


## Get list of all OneDrive Mike has access to.

$MikePPBUser = "clarkem_vpn@paddypowerbetfair.com"

$ALLOneDriveSites = Get-SPOSite -IncludePersonalSite $true -limit all -Filter "URL -like -my.sharepoint.com/personal/" | select Owner, URL

$FoundSite = @()
Write-Host "Checking For Mike's OneDrive Access .. " -foregroundcolor cyan -nonewline
foreach ($OneDrive in $ALLOneDriveSites)
{
    Write-host ". " -foregroundcolor darkcyan -nonewline
    if (Get-SPOUser -Site $OneDrive.URL -LoginName $MikePPBUser -ea silentlycontinue)
    {
            Write-host "Found. " -foregroundcolor green -nonewline
            $FoundSite += $OneDrive
    }
}

# Remove Calendar Events
Remove-CalendarEvents -Identity (mailboxEmailAddress) -QueryWindowInDays 1024
#Remove Calendar Events user Organized
Remove-CalendarEvents -Identity (mailboxEmailAddress) -QueryWindowInDays 1024 -CancelOrganizedMeetings 