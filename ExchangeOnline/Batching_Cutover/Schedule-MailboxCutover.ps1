#
# This function allows you to set the 'CompleteAfter' date for a single mailbox.
# The $DateTime parameter can also be set to 'now' if you want the cutover to occur immediately.
#

function Schedule-MailboxCutover
{
	[CmdletBinding()]
    param   (
                [string]
                $Mailbox,
				[string]
				$DateTime,
				[bool]
				$Confirm = $true
            )
	
	if (!$Mailbox)
	{
		Write-Host "Mailbox: " -ForegroundColor Cyan -NoNewline
		$Mailbox = (Read-Host).Trim()
	}
	
	if (!$DateTime)
	{
		Write-Host "Enter the date/time (ex: 1/1/2020 3:00 PM): " -ForegroundColor Cyan -NoNewline
		$DateTime = (Read-Host).Trim()
	}
	
	if (-not ($migUser = Get-MigrationUser $Mailbox.Trim()))
	{
		Write-Host "Unable to locate migration user for $mailbox." -ForegroundColor Red
		return
	}
	
	if ($DateTime -eq "now")
	{
		$cutoverDateTime = (Get-Date).ToUniversalTime()
	}
	elseif (-not ($cutoverDateTime = (Get-Date $DateTime).ToUniversalTime()))
	{		
		Write-Host "Unable to convert '$DateTime' to a valid format." -ForegroundColor Red
		return
	}
	else
	{
		$now = Get-Date
		if ($cutoverDateTime -lt $now)
		{
			Write-Host "The date/time you specified is in the past." -ForegroundColor Yellow
			Write-Host "Press 'Enter' if that's ok." -ForegroundColor Cyan
			$nothing = Read-Host
		}
		elseif ($cutoverDateTime -gt $now.AddDays(7))
		{
			Write-Host "The date/time you specified is more than a week in the future." -ForegroundColor Yellow
			Write-Host "Press 'Enter' if that's ok." -ForegroundColor Cyan
			$nothing = Read-Host
		}
	}
	
	if ($Confirm)
	{
		Write-Host
		Write-Host "Please confirm:" -ForegroundColor Green
		Write-Host "Mailbox: " -ForegroundColor Gray -NoNewline
		Write-Host $Mailbox -ForegroundColor Cyan
		Write-Host "Date/time: " -ForegroundColor Gray -NoNewline
		Write-Host $cutoverDateTime.ToLocalTime().ToString() -ForegroundColor Cyan
		Write-Host
		Write-Host "Press Enter if we're good." -ForegroundColor Yellow -NoNewline
		$nothing = Read-Host
	}
	else
	{
		Write-Host "$Mailbox ... " -ForegroundColor Cyan -NoNewline
	}
	
	$migUser | Set-MigrationUser -CompleteAfter $cutoverDateTime -Confirm:$false
	
	if (!$Confirm)
	{
		Write-Host "done" -ForegroundColor Green
	}
}