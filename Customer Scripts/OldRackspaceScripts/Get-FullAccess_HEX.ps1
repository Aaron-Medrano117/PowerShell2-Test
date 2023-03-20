
		  
		  #
$domain = # Customer domain name
		  #

$mailboxes = Get-Mailbox -OrganizationalUnit $domain -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*.emailsrvr.com"} | sort PrimarySmtpAddress


#REGION MEX05

$permsList = @()
foreach ($mbx in $mailboxes)
{
	Write-Host "$($mbx.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	[array]$perms = $mbx | Get-MailboxPermission | Where {!$_.IsInherited -and $_.AccessRights -like "*FullAccess*"}
	$perms = $perms | Where {$_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-*"}
	$perms = $perms | Where {$_.User -notlike "MEX0?\MEX0? *" -and $_.User -notlike "MEX0?\Migration*" -and $_.User -notlike "MEX0?\Managed Mail*"}
	
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

$permsList | Export-Csv "$HOME\Desktop\Perms_FullAccess.csv" -NoTypeInformation -Encoding UTF8

#ENDREGION


#REGION MEX06/8/9

$domainController = Get-ADDomainController -DomainName "acct.mlsrvr.com" -Discover
$serverFQDN = $domainController.Name + "." + $domainController.Domain

$permsList = @()
foreach ($mbx in $mailboxes)
{
	Write-Host "$($mbx.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	[array]$perms = $mbx | Get-MailboxPermission | Where {!$_.IsInherited -and $_.AccessRights -like "*FullAccess*"}
	$perms = $perms | Where {$_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-*"}
	$perms = $perms | Where {$_.User -notlike "MEX0?\MEX0? *" -and $_.User -notlike "MEX0?\Migration*" -and $_.User -notlike "MEX0?\Managed Mail*"}
	$perms = $perms | Where {$_.User.ToString() -ne $mbx.LinkedMasterAccount}
	
	if ($perms)
	{
		foreach ($perm in $perms)
		{
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select Mailbox, RecipientTypeDetails, UserWithFullAccess
			$tmp.Mailbox = $mbx.PrimarySmtpAddress.ToString()
			$tmp.RecipientTypeDetails = $mbx.RecipientTypeDetails
			$user = $perm.User.ToString().Replace("ACCT\","")
			$tmp.UserWithFullAccess = Get-ADUser -Filter {SamAccountName -eq $user} -Server $serverFQDN | select -ExpandProperty UserPrincipalName
			$permsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green
	}
	else
	{
		Write-Host " done" -ForegroundColor DarkCyan
	}
}

$permsList | Export-Csv "$HOME\Desktop\Perms_FullAccess.csv" -NoTypeInformation -Encoding UTF8

#ENDREGION


## Tenant to Tenant

$mailboxes = Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*DiscoverySearchMailbox*"}| sort PrimarySmtpAddress

$permsList = @()
foreach ($mbx in $mailboxes)
{
	[array]$perms = Get-MailboxPermission $mbx.PrimarySmtpAddress | Where {!$_.IsInherited -and $_.AccessRights -like "*FullAccess*"}
	$perms = $perms | Where {$_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-*"}

	if ($perms)
	{
		Write-Host "$($mbx.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
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
}

$permsList | Export-Csv -NoTypeInformation -Encoding UTF8 "$HOME\Desktop\Perms_FullAccess.csv"