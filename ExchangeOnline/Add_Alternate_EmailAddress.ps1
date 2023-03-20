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