#Remove All Alias addresses from Recipients
$allRecipients = import-csv "C:\Users\jwilliams\OneDrive - Arraya Solutions\ametek\leavers.csv"

Connect-MsolService

#ProgressBar
$progressref = ($allRecipients).count
$progresscounter = 0
foreach ($recipient in $allRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($recipient.DisplayName)"

    #GeneralThings
    $oldPrimarySMTPAddress = $recipient.PrimarySmtpAddress
    $EmailAddressesSplit = $recipient.EmailAddresses -split ","
    $EmailAddresses = $EmailAddressesSplit | ? {$_ -like '*@alphasense.*' -and $_ -notlike "*onmicrosoft.com"}
    
    Write-Host "Setting PrimarySMTPAddress and Removing $($EmailAddresses.count) aliases from $($oldPrimarySMTPAddress)... " -foregroundcolor Cyan -NoNewline
    
    #Get NewPrimarySMTPAddress - check if .onmicrosoft domain is on current recipient
    if ($onMicrosoftEmailAddressCheck = Get-Recipient $oldPrimarySMTPAddress | select -ExpandProperty EmailAddresses | ?{$_ -like "*onmicrosoft.com"}) {
        $onMicrosoftAddressSplit  = $onMicrosoftEmailAddressCheck -split ":"
        $onMicrosoftAddress = $onMicrosoftAddressSplit[1]
        Write-Host "OnMicrosoft address found.. " -ForegroundColor DarkYellow  -NoNewline
    }
    # If no .onmicrosoft exists, create a new one from PrimarySMTPAddress
    else {
        $addressSplit = $oldPrimarySMTPAddress -split "@"
        $onMicrosoftAddress = $addressSplit[0] + "@alphasensecorp.onmicrosoft.com"
        Write-Host "Created OnMicrosoft address.. " -ForegroundColor DarkMagenta  -NoNewline
    }
    #Update UserPrincipalName and PrimarySMTPAddress
    Set-MsolUserPrincipalName -UserPrincipalName $oldPrimarySMTPAddress -NewUserPrincipalName $onMicrosoftAddress
    Set-Mailbox -Identity $oldPrimarySMTPAddress -WindowsEmailAddress $onMicrosoftAddress

    #Remove Aliases
    foreach ($alias in $EmailAddresses) {
        Set-Mailbox -Identity $onMicrosoftAddress -EmailAddresses @{remove=$alias}
        Write-Host ". " -ForegroundColor Green  -NoNewline
    }
    Write-Host "done" -foregroundcolor green
}

#
