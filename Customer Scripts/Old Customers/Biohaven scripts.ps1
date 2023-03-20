## Biohaven Migration Project

#gather all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited 

$AllUsers = @()

foreach ($user in $Mailboxes)
{
    Write-Host "Gathering Mailbox Stats for $($user.DisplayName) .." -ForegroundColor Cyan -NoNewline
    
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
    Write-Host "done" -ForegroundColor Green
	
	$AllUsers += $currentuser
}

# Full Access Permissions

$mailboxes = Get-Mailbox -ResultSize Unlimited
cchomesmailboxes

$FullAccesspermsList = @()
foreach ($mbx in $mailboxes)
{
	Write-Host "$($mbx.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	[array]$perms = $mbx | Get-MailboxPermission | Where {!$_.IsInherited -and $_.AccessRights -like "*FullAccess*"}
	$perms = $perms | Where {$_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-*"}

	if ($perms)
	{
		foreach ($perm in $perms)
		{
            $recipientcheck = Get-recipient $perm.user
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select DisplayName, Mailbox, UserWithFullAccess, UserWithFullAccessAddress
			$tmp.DisplayName = $mbx.DisplayName.ToString()
            $tmp.Mailbox = $mbx.PrimarySmtpAddress.ToString()
			$tmp.UserWithFullAccess = $perm.User.ToString() | Get-Mailbox | select -ExpandProperty DisplayName
            $tmp.UserWithFullAccessAddress = $perm.User.ToString() | Get-Mailbox | select -ExpandProperty PrimarySMTPAddress
			$FullAccesspermsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green
	}
	else
	{
		Write-Host " done" -ForegroundColor DarkCyan
	}
}

# Send-As Perms
$recipients = Get-Recipient -ResultSize Unlimited

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
			$tmp = "" | select DisplayName, Recipient, RecipientType, ObjectWithSendAs, ObjectWithSendAsAddress
            $tmp.DisplayName = $recipient.DisplayName.ToString()
			$tmp.Recipient = $recipient.PrimarySmtpAddress.ToString()
			$tmp.RecipientType = $recipient.RecipientTypeDetails
			$tmp.ObjectWithSendAs = $perm.Trustee.ToString() | Get-Recipient -EA SilentlyContinue | select -ExpandProperty DisplayName
            $tmp.ObjectWithSendAsAddress = $perm.Trustee.ToString() | Get-Recipient | select -ExpandProperty PrimarySMTPAddress
			$sendAsPermsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green

	}
	else
	{
		Write-Host " done" -ForegroundColor DarkCyan
	}
}

##Match USERs. Grab MSOL Properties. Checks for multiple matches
$allMailboxes = Import-Csv

$allMSOLUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $allMailboxes | sort DisplayName) {
    $DisplayName = $user.DisplayName
    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline

    $tmp = "" | select DisplayName, CCTPrimarySmtpAddress, RecipientTypeDetails, IsInactiveMailbox, ExistsOnO365, UserPrincipalName, VVGPrimarySMTPAddress, IsLicensed, ExternalEmailAddress
    $tmp.DisplayName = $user.DisplayName
    $tmp.CCTPrimarySmtpAddress = $user.PrimarySmtpAddress

    if ($MSOLUsers = Get-MsolUser -searchstring $DisplayName -ea silentlycontinue) {

        $foundusers += $user
        
            foreach ($msoluser in $MSOLUsers)
            {
                if ($mailbox = Get-Mailbox $msoluser.userprincipalname -IncludeInactiveMailbox -ea silentlycontinue)
                {
                    $tmp.DisplayName = $mailbox.DisplayName
                    $tmp.VVGPrimarySMTPAddress = $mailbox.PrimarySmtpAddress
                    $tmp.RecipientTypeDetails = $mailbox.RecipientTypeDetails
                    $tmp.IsInactiveMailbox = $mailbox.IsInactiveMailbox
                    Write-Host "found" -ForegroundColor Green -NoNewline                    
                }

                elseif ($recipient = Get-Recipient $DisplayName -ea silentlycontinue)
                {
                    $tmp.DisplayName = $recipient.DisplayName
                    $tmp.VVGPrimarySMTPAddress = $recipient.PrimarySmtpAddress
                    $tmp.RecipientTypeDetails = $recipient.RecipientTypeDetails
                    $tmp.ExternalEmailAddress = $recipient.ExternalEmailAddress
                    Write-Host "found" -ForegroundColor Green -NoNewline
                }
                else
                {
                    Write-Host " .. Not a valid Exchange Recipient for $($msoluser.userprincipalname)" -ForegroundColor red -NoNewline
                }
            }
            $tmp.ExistsOnO365 = $true
            $tmp.UserPrincipalName = $msoluser.UserPrincipalName
            $tmp.IsLicensed = $MSOLUser.IsLicensed
            Write-Host "..done" -ForegroundColor Cyan
    }
    else
    {
        $notfoundusers += $user
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    $AllMSOLUsers += $tmp
}

### get 365 group size

$O365Groups = Get-UnifiedGroup -ResultSize Unlimited
 
$CustomResult=@() 
 
ForEach ($O365Group in $O365Groups){ 
If($O365Group.SharePointSiteUrl -ne $null) 
{ 
   $O365GroupSite=Get-SPOSite -Identity $O365Group.SharePointSiteUrl 
   $CustomResult += [PSCustomObject] @{ 
     GroupName =  $O365Group.DisplayName
     SiteUrl = $O365GroupSite.Url 
     StorageUsed_inMB = $O365GroupSite.StorageUsageCurrent
     Managedby = $O365Group.ManagedBy
     Description = $O365Group.Description
     Notes = $O365Group.Notes
     PrimarySMTPAddress = $O365Group.PrimarySMTPAddress
     AccessType = $O365Group.AccessType
  }
}} 
  
$CustomResult | select


##
<## Pull Delegate Calendar Perms
.EXAMPLE
Pull all UserMailboxes to a single OU. This is has to be used with the OnPrem Switch. Using this will be needed in Hosted Exchange.
.\Get-Mailboxattributes.ps1 -OutputCSVFilePath "c:\temp" -OnPrem $true -hexOU contoso.com
#>
param(
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Output CSV File path")][string]$OutputCSVFilePath,
)

$CalendarpermsList = @()
foreach ($user in $mailboxes)
{
    $mbx = get-mailbox $user.displayname
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
$csvFileName = "AllCalendarPermissions_$($date).csv"

$CalendarpermsList | Export-csv "Desired Path" –notypeinformation –encoding utf8


#ReplaceUPN
$msoluserscheck = get-msoluser -all

foreach ($user in $msolusers | sort userprincipalname)
{
	$newUPN = $user.UserPrincipalName.replace("@fanduelgroup.onmicrosoft.com","@fanduelgroup.com")
	Write-Host "Updating UPN for $($user.displayname) to $newupn .." -ForegroundColor Cyan -NoNewline
	Set-MsolUserPrincipalName -UserPrincipalName $user.UserPrincipalName -NewUserPrincipalName $newUPN
	Write-Host "done" -ForegroundColor Green
}

# Match and Combine Results ##
$AllUsers = @()
foreach ($mailbox in $AllKleoMailboxes)
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

    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline
    
    $mailboxcheck = Get-Mailbox $mailbox.DisplayName -ea silentlycontinue

    $MBXStats = Get-MailboxStatistics $mailboxcheck.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
    $msoluserscheck = get-msoluser -UserPrincipalName $mailboxcheck.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, BlockCredential, UserPrincipalName

    $currentuser | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $true
    $currentuser | add-member -type noteproperty -name "DestinationPrimarySMTPAddress" -Value $mailboxcheck.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "DestinationRecipientTypeDetails" -Value $mailboxcheck.RecipientTypeDetails
    $currentuser | Add-Member -type NoteProperty -Name "DestinationMBXSize" -Value $MBXStats.TotalItemSize
    $currentuser | Add-Member -Type NoteProperty -name "DestinationMBXItemCount" -Value $MBXStats.ItemCount
    $currentuser | add-member -type noteproperty -name "DestinationIsLicensed" -Value $msoluserscheck.IsLicensed
    $currentuser | add-member -type noteproperty -name "DestinationBlockSigninStatus" -Value $msoluserscheck.BlockCredential
    $currentuser | add-member -type noteproperty -name "DestinationUserPrincipalName" -Value $msoluserscheck.UserPrincipalName
    Write-Host "found mailbox." -ForegroundColor Green    
   
    $AllUsers += $currentuser
    #Read-Host "Stopping to check user"
}


#apply calendar perms
foreach ($CalendarPerm in $AllCalendarPerms)
{   
    $updateCalendarPath= $CalendarPerm.CalendarPath.Replace("@kleopharmaceuticals.com","@biohavenpharma.com")
    Write-Host "Updating $($updateCalendarPath). Adding Calendar perms for $($CalendarPerm.User) ..." -NoNewline -foregroundcolor cyan
    Add-MailboxFolderPermission -Identity $updateCalendarPath -User $CalendarPerm.User -AccessRights $CalendarPerm.AccessRights -ea silentlycontinue
    Write-Host "done" -ForegroundColor Green
}

#Set GrandSendOnBehalf

foreach ($mbx in $AllKleoMailboxes| ?{$_.GrantSendOnBehalfTo})
{
        $grantsendonbehalf = $mbx.GrantSendOnBehalfTo
        Write-Host "Updating GrantSendOnBehalfTo for mailbox $($mbx.DisplayName) " -ForegroundColor Cyan -NoNewline
        Set-Mailbox $mbx.DisplayName -GrantSendOnBehalfTo $grantsendonbehalf
        Write-Host "done" -ForegroundColor green
}

# Apply SendAs Perms

foreach ($sendasperm in $sendAsPermsList)
{
    $objectwithSendAsDestination = Get-recipient $sendasperm.ObjectWithSendAs

    if ($recipientcheck = Get-recipient $sendasperm.DisplayName -ea silentlycontinue)
    {
        $grantsendonbehalf = $mbx.GrantSendOnBehalfTo
        Write-Host "Updating SendAs for recipient $($recipientcheck.displayName) " -ForegroundColor Cyan -NoNewline
        Add-RecipientPermission $recipientcheck.displayName -trustee $objectwithSendAsDestination.displayname -AccessRights SendAs -Confirm:$false
        Write-Host "done" -ForegroundColor green
    }
    else {
        Write-host "No User found for $($sendasperm.DisplayName)" -foregroundcolor red
    }
}

# Add x500
$ALLMailboxes = Import-Csv 
foreach ($mbx in $ALLMailboxes)
{
    if ($mailboxcheck = get-mailbox $mbx.DisplayName -ea silentlycontinue)
    {
        Write-Host "Adding X500 to mailbox $($mbx.DisplayName)" -ForegroundColor Cyan -NoNewline
        $newX500 = $mbx.LegacyExchangeDN
        Set-Mailbox $mbx.DisplayName -EmailAddresses @{add=$newX500}
        Write-Host "done" -ForegroundColor green
    }
    else
    {
        Write-Host "No Mailbox found for $($mbx.DisplayName)" -ForegroundColor Red
    }
}

### Create 365 groups

$365Groups=Import-Csv "path to my csv file"
$GroupMembers = Import-csv

foreach ($Group in $365Groups)
    {
        $Name= $Group.GroupName
        $Mail= $Group.Email
        $Alias = $group."Group alias"
        $Owner= $Group.Owners
        $Notes = $group.Description
        $AccessType = $group."Group privacy"
        $PrimarySMTPAddress = $group."Group primary email".replace("@kleopharmaceuticals.com","@biohavenpharma.com")
        Write-Host "Creating Group $($name) .. " =-foregroundcolor cyan -nonewline
        New-UnifiedGroup -DisplayName $Name -alias $Alias -EmailAddresses $Mail -ManagedBy $Owner -Notes $Notes -AccessType $AccessType -Owners $group.OwnerName -PrimarySMTPAddress $PrimarySMTPAddress -ExoErrorAsWarning -Confirm:$false
        Write-Host "done" -foregroundcolor green
    }

    #Add Members to Office365 groups
foreach ($member in $GroupMembers)
    {
        $Name = $member."Group Name"
        Write-Host "Adding Member $($member.DisplayName) to group $($Name) .."
        Add-UnifiedGroupLinks $Name -LinkType Member -Links $member.DisplayName
        Write-Host "done" -foregroundcolor green
    }

# Shared Mailboxes
$sharedmailboxes

foreach ($sharedmailbox in $sharedmailboxes)
{
    get-msoluser -userprincipalname $sharedmailbox.userprincipalname | select DisplayName, UserPrincipalName, IsLicensed
}

###
#Get All Office 365 Groups
$O365Groups=Get-UnifiedGroup
ForEach ($Group in $UnifiedGroups)
{
    Write-Host "Group Name:" $Group.DisplayName -ForegroundColor Green
    Get-UnifiedGroupLinks –Identity $Group.Id –LinkType Members | Select DisplayName,PrimarySmtpAddress
 
    #Get Group Members and export to CSV
    Get-UnifiedGroupLinks –Identity $Group.Id –LinkType Members | Select-Object @{Name="Group Name";Expression={$Group.DisplayName}},`
         @{Name="User Name";Expression={$_.DisplayName}}, PrimarySmtpAddress | Export-CSV "C:\Users\fred5646\Rackspace Inc\Los Angeles Truck Center - CCTTS - General\Office365GroupMembers.csv" -NoTypeInformation -Append
}

### Create Groups
foreach ($Group in $365Groups)
    {
        $Name= $Group.DisplayName
        $Mail= $Group.Email
        $Alias = $group.Alias
        #$Owner= $Group.ManagedBy
        $Notes = $group.Notes
        $AccessType = $group.AccessType
        $PrimarySMTPAddress = $group.PrimarySmtpAddress.replace("@CentralCATrucks.com","@vvgtruck.com")
        Write-Host "Creating Group $($Name) .. " -foregroundcolor cyan -nonewline
        New-UnifiedGroup -DisplayName $Name -alias $Alias -Notes $Notes -AccessType $AccessType -PrimarySMTPAddress $PrimarySMTPAddress -ExoErrorAsWarning -Confirm:$false
        Write-Host "done" -foregroundcolor green
    }
 
    #Add Members to Office365 groups
foreach ($member in $GroupMembers)
    {
        $Name = $member."Group Name"
        Write-Host "Adding Member $($member.DisplayName) to group $($Name) .."
        Add-UnifiedGroupLinks $Name -LinkType Member -Links $member.DisplayName
        Write-Host "done" -foregroundcolor green
    }
    
    #Add Owners to Office365 Groups
foreach ($owner in $365Groups)
    {
        $Name = $owner.GroupName
        [array]$Owners = $owner.ManagedBy -split ","
        foreach ($owner in $Owners)
        {
            $recipientcheck = Get-MsolUser -searchstring $owner
            $OwnerDisplayName = $recipientcheck.DisplayName
            Write-Host "Adding Owner $($OwnerDisplayName) to group $($Name) .."
            Add-UnifiedGroupLinks $Name -LinkType Member -Links $OwnerDisplayName
            Add-UnifiedGroupLinks $Name -LinkType Owner -Links $OwnerDisplayName
            Write-Host "done" -foregroundcolor green 
        }
        
        
    }

    $allmsolusers
    foreach ($user in $allmsolusers | ?{$_.userprincipalname -like "*kleo*"})
    {
        Write-Host " Updating UPN for $($user.DisplayName) ..." -foregroundcolor cyan -nonewline 
        $UPNSplit = $user.UserPrincipalName -split "@"
        $newUPN = $UPNSplit[0] + "@appriver3651017509.onmicrosoft.com"
        Set-MsolUserPrincipalName -UserPrincipalName $user.userprincipalname -NewUserPrincipalName $newUPN
        Write-Host "done" -foregroundcolor green
    }

 
    foreach ($user in $allmsolusers2)
    {
        Write-Host " Updating PrimarySMTP for $($user.DisplayName) ..." -foregroundcolor cyan -nonewline 
        if ($mailboxcheck = get-mailbox $user.DisplayName -ea silentlycontinue)
        {
            Set-Mailbox $mailboxcheck.DisplayName -WindowsEmailAddress $user.PrimarySmtpAddress
            Write-Host "done" -foregroundcolor green
        }
        else {
            Write-host "no mailbox found. Skipping" -foregroundcolor yellow
        }
    }

    foreach ($user in $allmsolusers2)
    {
        Write-Host " Updating PrimarySMTP for $($user.DisplayName) ..." -foregroundcolor cyan -nonewline 
        if ($mailboxcheck = get-mailbox $user.DisplayName -ea silentlycontinue)
        {
            Set-Mailbox $mailboxcheck.DisplayName -WindowsEmailAddress $user.userprincipalname
            Write-Host "done" -foregroundcolor green
        }
        else {
            Write-host "no mailbox found. Skipping" -foregroundcolor yellow
        }
    }

    foreach ($user in $matchedusers)
    {
        Write-Host "Adding KleoPharma Email Address for $($user.DisplayName) ..." -foregroundcolor cyan -nonewline 
        if ($mailboxcheck = get-mailbox $user.DisplayName -ea silentlycontinue)
        {
            $KLEOPHARMAAddress = $user.PrimarySmtpAddress
            Set-Mailbox $mailboxcheck.DisplayName -EmailAddresses @{add=$KLEOPHARMAAddress}
            Write-Host "done" -foregroundcolor green
        }
        else {
            Write-host "no mailbox found. Skipping" -foregroundcolor yellow
        }
    }

###
    foreach ($user in $matchedusers)
    {
        $KLEODLName = $user.DisplayName + "- Kleo"
        $alias = $user.primarysmtpaddress -split "@"
        Write-Host "Creating KleoPharma DL for $($user.DisplayName) ..." -foregroundcolor cyan -nonewline 
        $PrimarySMTPAddress = $user.PrimarySMTPAddress
        New-DistributionGroup -DisplayName $KLEODLName -name $KLEODLName -PrimarySMTPAddress $PrimarySMTPAddress -alias $alias[0]
        Start-Sleep -seconds 5
        Set-DistributionGroup $KLEODLName -HiddenFromAddressListsEnabled:$True
        Write-Host "done" -foregroundcolor green
        Read-Host "pause to check"
    }
    foreach ($user in $matchedusers)
    {
        $KLEODLName = $user.DisplayName + " - Kleo"
        $alias = $user.primarysmtpaddress -split "@"
        Write-Host "Adding KleoPharma Member to $($user.DisplayName) ..." -foregroundcolor cyan -nonewline 
        Add-DistributionGroupMember $KLEODLName -member $user.DisplayName
    }

    # Update Export Email Addresses
foreach ($mailbox in $Mailboxes)
{
        $newExportAddress = $mailbox.ExportEmailAddress.replace("@kleopharmaceuticals.com","@‎appriver3651017509.onmicrosoft.com")
        $result = Set-MW_Mailbox -ticket $mwTicket -ConnectorId $connector.id -mailbox $mailbox -ExportEmailAddres $newExportAddress
}


foreach ($user in $mailboxes)
{
    $UPNSplit =  $user.PrimarySmtpAddress -split "@"
    $MAILAaddress = $UPNSplit[0] + "@appriver3651017509.mail.onmicrosoft.com"
    $MAILAaddress
    Set-Mailbox $user.PrimarySmtpAddress -EmailAddresses @{add=$MAILAaddress}
}