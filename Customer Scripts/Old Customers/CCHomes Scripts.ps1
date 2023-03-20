# CCHomes Scripts

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()
$MultipleUsers = @()

foreach ($user in $importcsv | sort DisplayName) {
    $DisplayName = $user.DisplayName
    Write-Host "Checking for $DisplayName in Tenant ..." -fore Cyan -NoNewline

    $currentuser = new-object psobject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName

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

##### Match RSE and HEX

$RSEfoundusers = @()
$RSEnotfoundusers = @()
foreach ($mbx in $RSEmailboxes)
{
    if ($msolusercheck = Get-MsolUser -userprincipalname $mbx.email)
    {
        $RSEfoundusers += $mbx
        $msolusercheck | select DisplayName, UserPrincipalName, IsLicensed
    }
    else 
    {
        $RSEnotfoundusers += $mbx
    }
}

$HEXfoundusers = @()
$HEXnotfoundusers = @()
$allUsers = @()
foreach ($mbx in $HEXmailboxes | sort displayname)
{
    $UPNSplit = $mbx.PrimarySMTPAddress -split "@"
    Write-Host "Checking for $($mbx.DisplayName) in Tenant ..." -fore Cyan -NoNewline
    $currentuser = new-object psobject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $mbx.DisplayName
    $currentuser | add-member -type noteproperty -name "HEXPrimarySMTPAddress" -Value $mbx.PrimarySMTPAddress

    if ($msolcheck = Get-MsolUser -UserPrincipalName $mbx.PrimarySMTPAddress -ea silentlycontinue)
    {
        $HEXfoundusers += $mbx
        Write-Host "found" -fore green
        $currentuser | add-member -type noteproperty -name "ExistsOnO365" -Value $true
        $currentuser | add-member -type noteproperty -name "365UserPrincipalName" -Value $msolcheck.UserPrincipalName
        $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msolcheck.IsLicensed

        if ($recipientcheck = Get-Recipient $msolcheck.userprincipalname -ea silentlycontinue)
        {
            $currentuser | add-member -type noteproperty -name "EXOPrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress
            $currentuser | add-member -type noteproperty -name "RecipientType" -Value $recipientcheck.RecipientTypeDetails
        }
        else {
            $currentuser | add-member -type noteproperty -name "EXOPrimarySMTPAddress" -Value "NoMailboxFound"
            $currentuser | add-member -type noteproperty -name "EXOPrimarySMTPAddress" -Value "NoRecipientFound"
        }
        
    }
    elseif ($msolcheck = Get-MsolUser -searchstring $UPNSplit[0] -ea silentlycontinue)
    {
        $HEXfoundusers += $mbx
        Write-Host "found" -fore green
        
        if ($msolcheck.count -gt 1)
        {
            $currentuser | add-member -type noteproperty -name "ExistsOnO365" -Value "multipleusersfound"
        }
        else
        {
            $currentuser | add-member -type noteproperty -name "ExistsOnO365" -Value $true
            $currentuser | add-member -type noteproperty -name "365UserPrincipalName" -Value $msolcheck.UserPrincipalName
            $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msolcheck.IsLicensed
            if ($recipientcheck = Get-Recipient $msolcheck.userprincipalname -ea silentlycontinue)
            {
                $currentuser | add-member -type noteproperty -name "EXOPrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress
                $currentuser | add-member -type noteproperty -name "RecipientType" -Value $recipientcheck.RecipientTypeDetails
            }
            else 
            {
                $currentuser | add-member -type noteproperty -name "EXOPrimarySMTPAddress" -Value "NoMailboxFound"
                $currentuser | add-member -type noteproperty -name "EXOPrimarySMTPAddress" -Value "NoRecipientFound"
            }
        }   
    }
    else 
    {
        $HEXnotfoundusers += $mbx
        Write-Host "notfound" -fore red
        $currentuser | add-member -type noteproperty -name "ExistsOnO365" -Value $False
        $currentuser | add-member -type noteproperty -name "365UserPrincipalName" -Value $null
        $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $null
        $currentuser | add-member -type noteproperty -name "EXOPrimarySMTPAddress" -Value $null
        $currentuser | add-member -type noteproperty -name "EXOPrimarySMTPAddress" -Value $null
    }
    $allUsers += $currentuser
}

## grab immutable id

foreach ($user in $cchomesusers)
{
    $guid = $user | select -ExpandProperty ObjectGUID
    $immutableID = [System.Convert]::ToBase64String(([GUID]($guid)).ToByteArray())
    $user.ImmutableID = $immutableID
}

# remove routing address
$cchomesmailboxes = get-mailbox -OrganizationalUnit cchomes.com |sort displayName

foreach ($mbx in $cchomesmailboxes)
{
    Write-Host "Checking mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $aliasarray = $mbx.EmailAddresses.ProxyAddressString
    Write-Host $aliasarray.count "aliases found . " -fore darkcyan -nonewline
    foreach ($alias in $aliasarray)
    {
        if ($alias -like "*@routing.cchomes.com")
        {
            Write-host "Removing address $($alias)"
            Set-Mailbox $mbx.alias -EmailAddresses @{remove=$alias}
        }
        else
        {
            Write-host "." -nonewline -fore Yellow
        } 
    }
    Write-Host "done." -fore green
}

# Full Access Permissions

$FullAccesspermsList = @()
foreach ($mbx in $cchomesmailboxes)
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

#Calendar Perms
$calendarPermsList = @()
foreach ($mbx in $cchomesmailboxes)
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

                $calendarPermsList += $currentcalendar

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

# Update Contacts
foreach ($contact in $hexcontacts) 
    {
    $EmailArray = $contact.EmailAddresses -split ","
    foreach ($alias in $emailarray | ? {$_ -like "*@cchomes.com"})
    {
        Set-MailContact $contact.DisplayName -EmailAddresses @{add=$alias}
    }
}

$allattributes = @()
foreach ($list in $dls)
    {
        write-host "Working on list" $list.primarysmtpaddress -fore cyan
        #remove any addresses we don't want and store the remaining ones in the $addresses variable
        $addresses = $list.emailaddresses | ?{$_ -notlike "SPO:*" -and $_ -notlike "sip:*" -and $_ -notlike "*onmicrosoft.com"}
        #create a custom object to store the attributes we're interested in
        $currentlist = "" | select DisplayName, FirstName, LastName, EmailAddresses, HiddenFromAddressListsEnabled, LegacyExchangeDN, PrimarySMTPAddress
        #populate that custom object
        $currentlist.DisplayName = $list.DisplayName
        $currentlist.FirstName = $list.FirstName
        $currentlist.LastName = $list.LastName
        $currentlist.EmailAddresses = ($addresses -join ",")
        $currentlist.HiddenFromAddressListsEnabled = $list.HiddenFromAddressListsEnabled
        $currentlist.LegacyExchangeDN = $list.LegacyExchangeDN
        $currentlist.PrimarySMTPAddress = $list.PrimarySMTPAddress
        #add the custom object to our array
        $allattributes += $currentlist
    }


    $forwads
    foreach ($user in $forwads)
    {
        if ($mailboxdetails = Get-Mailbox $user.email -ea silentlycontinue)
        {
            Write-Host "Update Forward for mailbox $($mailboxdetails.Displayname) .. " -ForegroundColor cyan -NoNewline
            $forwards = $user.forwardto -split ","
            foreach ($forward in $forwards)
            {
                Write-Host "Found $($forward.count) forwards. " -ForegroundColor cyan -NoNewline
                if ($RecipientCheck = get-recipient $forward -ea silentlycontinue)
                {
                    Write-Host "Adding forward $($forward) ... " -NoNewline
                    Set-Mailbox $mailboxdetails.primarysmtpaddress -forwardingaddress $recipientcheck.primarysmtpaddress
                }
                Write-Host "done" -ForegroundColor Green
            }
        }
    }
        
    Get-Mailbox $_.email | select DisplayName, primarysmtpaddress, forwardingsmtpaddress, forwardingaddress}