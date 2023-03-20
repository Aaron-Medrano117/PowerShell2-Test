
# The CSV must contain the following 3 columns:
# Mailbox
# DeliverToMailboxAndForward
# ForwardingAddress


$forwards = Import-Csv $HOME\Desktop\Forwards.csv

while ($true)
{
	foreach ($fwd in $forwards)
	{
		Write-Host "$($fwd.Mailbox) ... " -ForegroundColor Cyan -NoNewline
		if ($mbx = Get-Mailbox $fwd.Mailbox -ErrorAction SilentlyContinue)
		{
			if ((-not $mbx.ForwardingAddress) -and (-not $mbx.ForwardingSmtpAddress))
			{
				$mbx | Set-Mailbox -ForwardingAddress $fwd.ForwardingAddress -DeliverToMailboxAndForward $([System.Convert]::ToBoolean($fwd.DeliverToMailboxAndForward)) #-WhatIf
				Write-Host $fwd.ForwardingAddress -ForegroundColor Green
			}
			else
			{
				Write-Host "already configured" -ForegroundColor Yellow
			}
		}
		else
		{
			Write-Host "mailbox not found" -ForegroundColor DarkGray
		}
	}
	
	Start-Sleep -Seconds 15
	Write-Host
	Write-Host "============================================" -ForegroundColor Magenta
	Write-Host
}


