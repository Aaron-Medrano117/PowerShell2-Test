$mailboxes = Get-Mailbox -ResultSize Unlimited | sort PrimarySmtpAddress

$forwardList = @()
$forwardMailboxes = $mailboxes | Where {$_.ForwardingAddress -or $_.ForwardingSmtpAddress}
$progressref = ($forwardMailboxes).count
$progresscounter = 0

foreach ($mbx in $forwardMailboxes)
{
	$progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($mbx.DisplayName)"

	Write-Host $mbx.PrimarySmtpAddress -ForegroundColor Cyan
	$currentPerm = new-object PSObject
        
	$currentPerm | add-member -type noteproperty -name "Mailbox" -Value $mbx.PrimarySmtpAddress.ToString()
	$currentPerm | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $mbx.DeliverToMailboxAndForward
	$currentPerm | add-member -type noteproperty -name "ForwardingAddress" -Value  $mbx.ForwardingAddress | Get-Recipient | select -ExpandProperty PrimarySmtpAddress
	$currentPerm | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $mbx.ForwardingSmtpAddress
	$forwardList += $currentPerm
}
$forwardList | Export-Csv "$HOME\Desktop\Forwards.csv" -NoTypeInformation -Encoding UTF8
