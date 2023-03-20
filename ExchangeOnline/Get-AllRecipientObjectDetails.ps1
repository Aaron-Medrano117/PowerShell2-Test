# Gather All Recipient Objects
$AllRecipientDetails = @()
$allRecipients = Get-Recipient -ResultSize Unlimited
$progressref = $allRecipients.count
$progresscounter = 0
foreach ($recipient in $allRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Recipient Details for $($recipient.PrimarySMTPAddress)"

    $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $recipient.DisplayName
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ",") -force

    $AllRecipientDetails += $currentuser
}
$AllRecipientDetails | Export-Excel "C:\Users\aaron.medrano\Desktop\post migration\EHN-AllRecipients.xlsx"