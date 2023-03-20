
		  #
$domain = # Customer's domain name
		  #

$recipients = @()
$recipients += Get-DistributionGroup -OrganizationalUnit $domain -ResultSize Unlimited
$recipients += Get-Mailbox -OrganizationalUnit $domain -ResultSize Unlimited | Where {$_.PrimarySmtpAddress.ToString() -notlike "*.emailsrvr.com"}


#REGION MEX05

$permsList = @()
foreach ($recipient in $recipients | sort PrimarySmtpAddress)
{
	Write-Host "$($recipient.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	$perms = $recipient | Get-AdPermission | Where {!$_.IsInherited -and $_.ExtendedRights -like "*Send-As*"}
	$perms = $perms | Where {$_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-1-*" -and $_.User -notlike "*svc_besadmin*"}
	$perms = $perms | Where {$_.User -notlike "MEX05\MEX05 *" -and $_.User -notlike "MEX05\Managed *"}
	
	if ($perms)
	{
		foreach ($perm in $perms)
		{
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select Recipient, RecipientType, ObjectWithSendAs
			$tmp.Recipient = $recipient.PrimarySmtpAddress.ToString()
			$tmp.RecipientType = $recipient.RecipientTypeDetails
			$tmp.ObjectWithSendAs = $perm.User.ToString() | Get-Recipient -EA SilentlyContinue | select -ExpandProperty PrimarySmtpAddress
			$permsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green
	}
	else
	{
		Write-Host " done" -ForegroundColor DarkCyan
	}
}

$permsList | Export-Csv "$HOME\Desktop\Perms_SendAs.csv" -NoTypeInformation -Encoding UTF8

#ENDREGION


#REGION MEX06/8/9

$domainController = Get-ADDomainController -DomainName "acct.mlsrvr.com" -Discover
$serverFQDN = $domainController.Name + "." + $domainController.Domain

$permsList = @()
foreach ($recipient in $recipients | sort PrimarySmtpAddress)
{
	Write-Host "$($recipient.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	$perms = $recipient | Get-AdPermission | Where {!$_.IsInherited -and $_.ExtendedRights -like "*Send-As*"}
	$perms = $perms | Where {$_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-1-*" -and $_.User -notlike "*svc_besadmin*"}
	
	if ($recipient.RecipientTypeDetails -eq "LinkedMailbox")
	{
		$perms = $perms | Where {$_.User.ToString() -notlike $recipient.LinkedMasterAccount}
	}
	
	if ($perms)
	{
		foreach ($perm in $perms)
		{
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select Recipient, RecipientType, ObjectWithSendAs
			$tmp.Recipient = $recipient.PrimarySmtpAddress.ToString()
			$tmp.RecipientType = $recipient.RecipientTypeDetails
			$user = $perm.User.ToString().Replace("ACCT\","")
			$tmp.ObjectWithSendAs = Get-ADUser -Filter {SamAccountName -eq $user} -Server $serverFQDN | select -ExpandProperty UserPrincipalName
			$permsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green
	}
	else
	{
		Write-Host " done" -ForegroundColor DarkCyan
	}
}

$permsList | Export-Csv "$HOME\Desktop\Perms_SendAs.csv" -NoTypeInformation -Encoding UTF8

#ENDREGION


#REGION Tenant To Tenant

$recipients = Get-Recipient -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*DiscoverySearchMailbox*"} | sort PrimarySmtpAddress

$permsList = @()
foreach ($recipient in $recipients)
{
	$perms = $recipient | Get-RecipientPermission | Where {!$_.IsInherited -and $_.ExtendedRights -like "*Send-As*"}
	$perms = $perms | Where {$_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-1-*"}
	
	if ($recipient.RecipientTypeDetails -eq "UserMailbox")
	{
		$perms = $perms
	}
	
	if ($perms)
	{
		Write-Host "$($recipient.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline

		foreach ($perm in $perms)
		{
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select Recipient, RecipientType, ObjectWithSendAs
			$tmp.Recipient = $recipient.PrimarySmtpAddress
			$tmp.RecipientType = $recipient.RecipientTypeDetails
			$tmp.Trustee = $perm.Trustee
			$permsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green
	}
}

$permsList | Export-Csv "$HOME\Desktop\Perms_SendAs.csv" -NoTypeInformation -Encoding UTF8

#ENDREGION
