

$sendAsConfig = Import-Csv $HOME\Desktop\SendAs.csv
$uniqueRecipients = $sendAsConfig | select -ExpandProperty Recipient | sort | Get-Unique


$failedToStamp = @()
foreach ($uniqueRecipient in $uniqueRecipients)
{
	Write-Host "$uniqueRecipient ..." -ForegroundColor Cyan -NoNewline
	
	if ($recipient = Get-Recipient $uniqueRecipient -EA SilentlyContinue)
	{
		$objects = $sendAsConfig | Where {$_.Recipient -eq $uniqueRecipient} | select -ExpandProperty ObjectWithSendAs
		
		$errors = $false
		foreach ($object in $objects)
		{
			if ($assignee = Get-Recipient $object -EA SilentlyContinue)
			{
				try
				{
					Add-RecipientPermission $recipient.DistinguishedName -Trustee $assignee.ExchangeGuid.Guid -AccessRights SendAs -Confirm:$false -EA Stop -WA SilentlyContinue
					Write-Host "." -ForegroundColor Green -NoNewline
				}
				catch
				{
					$tmp = "" | select Recipient, ObjectWithSendAs
					$tmp.Recipient = $uniqueRecipient
					$tmp.ObjectWithSendAs = $object
					$failedToStamp += $tmp
					$errors = $true
					Write-Host "." -ForegroundColor Red -NoNewline
				}
			}
			else
			{
				Write-Host "." -ForegroundColor Yellow -NoNewline
			}
		}
		
		if (-not ($errors))
		{
			Write-Host " done" -ForegroundColor Green
		}
		else
		{
			Write-Host " done" -ForegroundColor Yellow
		}
	}
	else
	{
		Write-Host " not found" -ForegroundColor Red
	}
}

