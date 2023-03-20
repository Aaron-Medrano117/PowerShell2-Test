$domain =  Read-Host "What Domain?"

$mailboxes = get-mailbox -organizationalunit $domain -resultsize unlimited

foreach ($mailbox in $mailboxes) 
{
    $AlternateDomainAddresses = @()
    $RoutingDomainAddresses = @()
    Write-Host "Checking" $mailbox.DisplayName "..." -ForegroundColor Cyan -NoNewline
        if ($AlternateDomainAddresses = $mailbox |  ? {$_.EmailAddresses -like "*routing.*"} | select -ExpandProperty EmailAddresses) 
        {
            [array]$RoutingDomainAddresses = $AlternateDomainAddresses | ? {$_.addressstring -like "*routing."}
            Write-Host "Removing "$RoutingDomainAddresses.count" Routing Addresses found..." -ForegroundColor White -NoNewline
            #Write-Host ""$RoutingDomainAddresses.count" found ..."
                foreach ($routingAddress in $RoutingDomainAddresses)
                {
                    [string]$ProxyAddressString = $routingAddress.ProxyAddressString
                    #$ProxyAddressString
                    $mailbox | Set-Mailbox -EmailAddresses @{remove=$ProxyAddressString} -warningaction SilentlyContinue #-whatif
                }
            Write-Host "done" -ForegroundColor Green
        }
        else
        {
            Write-Host "No Routing Address found. Skipping" -ForegroundColor Yellow
        } 
}