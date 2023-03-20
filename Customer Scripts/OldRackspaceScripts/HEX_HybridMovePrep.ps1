#This script is used to prepare mailboxes on Hosted Exchange to Migrate over to Exchange Online
#Still in progress

#collect variables
$domain = Write-Host "What is the domain you wish to prepare for Exchange Online Hybrid Migration?" -ForegroundColor White
[string]$tenantname = Write-Host "What is the tenant domain on 365? (do not need .onmicrosoft)"

$domain_mailboxes = get-mailbox -organizationalunit $domain -resultsize unlimited

#Add .mail.onmicrosoft.com address to mailboxes
foreach ($mailbox in $domain_mailboxes) {
$MicrosoftMailAddress = "smtp:" + $mailbox.name + $tenant +".mail.onmicrosoft.com"
Write-Host "Adding $MicrosoftMailAddress to $mailbox" -ForegroundColor Blue

Get-mailbox $mailbox | set-mailbox -emailaddresses @{add="$MicrosoftMailAddress"}
Get-mailbox $mailbox | select DisplayName, alias, emailaddresses

}

#check and remove stored move requests
foreach ($mailbox in $domain_mailboxes) {
    if (get-moverequest -name $mailbox)
    {Write-Host "Found Move Request for $mailbox" -ForegroundColor Blue
    Get-moverequest $mailbox | select name, status, batchname
    
    Write-Host "Removing Move Request for $mailbox"
    Remove-moverequest -name $mailbox
    }
    else {Write-Host "No move found for $mailbox. Skipping ..." -ForegroundColor Green}
}