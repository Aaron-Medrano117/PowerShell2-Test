#Count Mobile devices for users

$mailboxes = get-mailbox -organizationalunit americancarcenter.com -resultsize unlimited
$allUsers = @()

foreach ($mbx in $mailboxes) 
{
    $ASDevices = Get-ActiveSyncDevice -mailbox $mbx.identity

    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $mbx.DisplayName
    $currentuser | add-member -type noteproperty -name "Email Address" -Value $mbx.PrimarySMTPAddress
    $currentuser | add-member -type noteproperty -name "MobileDevices" -Value $ASDevices.Count
    $allUsers += $currentuser
}