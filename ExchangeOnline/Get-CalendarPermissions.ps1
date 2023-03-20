
		  
		  #
$domain = # Customer domain
		  #

$mailboxes = Get-Mailbox -OrganizationalUnit $domain -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*.emailsrvr.com"} | sort PrimarySmtpAddress


#REGION Get list of calendar permissions

$uniqueDomains = $mailboxes | select -ExpandProperty PrimarySmtpAddress | select -ExpandProperty Domain | sort -Unique

$calendarpermsList = @()
foreach ($mbx in $mailboxes)
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
		
		[array]$perms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.User.ADRecipient -and $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" -and $_.User.ToString() -notlike "*S-1-*" -and $_.User.ADRecipient.PrimarySmtpAddress.ToString() -ne $mbx.PrimarySmtpAddress.ToString()}
		
		if ($perms)
		{
			foreach ($perm in $perms)
			{
				if ($uniqueDomains -contains $perm.User.ADRecipient.PrimarySmtpAddress.Domain)
				{
					$tmp = "" | select Mailbox, CalendarPath, User, AccessRights
					$tmp.Mailbox = $mbx.PrimarySmtpAddress.ToString()
					$tmp.CalendarPath = $calendar.FolderPath
					$tmp.User = $perm.User.ADRecipient.PrimarySmtpAddress.ToString()
					$tmp.AccessRights = $perm.AccessRights -join ","
					$permsList += $tmp
					Write-Host "." -ForegroundColor Yellow -NoNewline
				}
			}
			
			Write-Host " done" -ForegroundColor Green
		}
		else
		{
			Write-Host " done" -ForegroundColor DarkCyan
		}
	}
}

$calendarpermsList | Export-Csv "$HOME\Desktop\Perms_Calendar.csv" -NoTypeInformation -Encoding UTF8

#ENDREGION


#REGION Separate the calendar sharing users from the non-sharing users

$calendarSharingUsers = @()
$calendarSharingUsers += $calendarpermsList | select -ExpandProperty Mailbox
$calendarSharingUsers += $calendarpermsList | select -ExpandProperty User
$calendarSharingUsers = $calendarSharingUsers | sort -Unique | Where {$_}

$nonCalendarSharingUsers = @()
foreach ($mbx in $mailboxes)
{
	$upn = $mbx.UserPrincipalName
	$primarySmtpAddress = $mbx.PrimarySmtpAddress.ToString()
	
	Write-Host "$primarySmtpAddress ... " -ForegroundColor Cyan -NoNewline
	
	if ($calendarSharingUsers -notcontains $upn -and $calendarSharingUsers -notcontains $primarySmtpAddress)
	{
		Write-Host "done" -ForegroundColor Green
		$nonCalendarSharingUsers += $primarySmtpAddress
	}
	else
	{
		Write-Host "done" -ForegroundColor Yellow
	}
}

$calendarSharingUsers | Out-File $HOME\Desktop\RelatedUsers.txt
$nonCalendarSharingUsers | Out-File $HOME\Desktop\UnrelatedUsers.txt

#ENDREGION



## REGION Tenant To Tenant

$mailboxes = Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*DiscoverySearchMailbox*"} | sort PrimarySmtpAddress

#REGION Get list of calendar permissions

$calendarpermsList = @()
foreach ($mbx in $mailboxes)
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
		
		[array]$perms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.User.ADRecipient -and $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" -and $_.User.ToString() -notlike "*S-1-*" -and $_.User.ADRecipient.PrimarySmtpAddress.ToString() -ne $mbx.PrimarySmtpAddress.ToString()}
		
		if ($perms)
		{
			foreach ($perm in $perms)
			{
					$tmp = "" | select Mailbox, CalendarPath, User, AccessRights
					$tmp.Mailbox = $mbx.PrimarySmtpAddress.ToString()
					$tmp.CalendarPath = $calendar.FolderPath
					$tmp.User = $perm.User.ADRecipient.PrimarySmtpAddress.ToString()
					$tmp.AccessRights = $perm.AccessRights -join ","
					$calendarpermsList += $tmp
					Write-Host "." -ForegroundColor Yellow -NoNewline
			}
		}
		Write-Host " done" -ForegroundColor Green
	}
}

$calendarpermsList | Export-Csv "$HOME\Desktop\Perms_Calendar.csv" -NoTypeInformation -Encoding UTF8

#ENDREGION


#REGION Separate the calendar sharing users from the non-sharing users

$calendarSharingUsers = @()
$calendarSharingUsers += $calendarpermsList | select -ExpandProperty Mailbox
$calendarSharingUsers += $calendarpermsList | select -ExpandProperty User
$calendarSharingUsers = $calendarSharingUsers | sort -Unique | Where {$_}

$nonCalendarSharingUsers = @()
foreach ($mbx in $mailboxes)
{
	$upn = $mbx.UserPrincipalName
	$primarySmtpAddress = $mbx.PrimarySmtpAddress.ToString()
	
	Write-Host "$primarySmtpAddress ... " -ForegroundColor Cyan -NoNewline
	
	if ($calendarSharingUsers -notcontains $upn -and $calendarSharingUsers -notcontains $primarySmtpAddress)
	{
		Write-Host "done" -ForegroundColor Green
		$nonCalendarSharingUsers += $primarySmtpAddress
	}
	else
	{
		Write-Host "done" -ForegroundColor Yellow
	}
}

$calendarSharingUsers | Out-File $HOME\Desktop\RelatedUsers.txt
$nonCalendarSharingUsers | Out-File $HOME\Desktop\UnrelatedUsers.txt

#ENDREGION
