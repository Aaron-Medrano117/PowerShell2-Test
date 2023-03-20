#ProgressBar
$progressref = ($sunCostUsers).count
$progresscounter = 0

foreach ($user in $sunCostUsers)   {
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($user.name)"
    $mbCheck = Get-Mailbox $user.Mail
    $MBXStats = Get-MailboxStatistics $user.Mail -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
    $addresses = $mbCheck | select -ExpandProperty EmailAddresses
    
    $user | add-member -type noteproperty -name "OrganizationalUnit" -Value $mbCheck.OrganizationalUnit
    $user | add-member -type noteproperty -name "DisplayName" -Value $mbCheck.DisplayName
    $user | add-member -type noteproperty -name "UserPrincipalName" -Value $mbCheck.userprincipalname        
    $user | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $mbCheck.PrimarySmtpAddress
    $user | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mbCheck.RecipientTypeDetails
    $user | add-member -type noteproperty -name "EmailAddresses" -Value ($addresses -join ',')
    $user | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $mbCheck.legacyexchangedn)
    $user | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value ($mbCheck.AcceptMessagesOnlyFrom -join ",")
    $user | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($mbCheck.GrantSendOnBehalfTo -join ",")
    $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $mbCheck.HiddenFromAddressListsEnabled
    $user | add-member -type noteproperty -name "RejectMessagesFrom" -Value ($mbCheck.RejectMessagesFrom -join ",")
    $user | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $mbCheck.DeliverToMailboxAndForward
    $user | add-member -type noteproperty -name "ForwardingAddress" -Value $mbCheck.ForwardingAddress
    $user | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $mbCheck.ForwardingSmtpAddress
    $user | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize
    $user | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount
    $user | Add-Member -Type NoteProperty -name "MBXDatabase" -Value $mbCheck.Database
    $user | Add-Member -Type NoteProperty -name "Identity" -Value $mbCheck.Identity
    $user | Add-Member -Type NoteProperty -Name "ArchiveGUID" -Value $mbCheck.ArchiveGuid
    $user | add-member -type noteproperty -name "ArchiveState" -Value $mbCheck.ArchiveState
    $user | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $mbCheck.ArchiveStatus

}

## Check if Users in Office 365
#ProgressBar
$progressref = ($suncoastUsers).count
$progresscounter = 0

foreach ($user in $suncoastUsers)   {
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking 365 for $($user.DisplayName)"
    if ($msolcheck = Get-MsolUser -SearchString $user.PrimarySmtpAddress) {
        $user | add-member -type noteproperty -name "InOffice365" -Value $True -Force
        $user | add-member -type noteproperty -name "IsLicensed" -Value $msolcheck.IsLicensed -Force
        $user | add-member -type noteproperty -name "Licenses" -Value ($msolcheck.licenses.accountskuid -join ",") -Force
    }
    else {
        $user | add-member -type noteproperty -name "InOffice365" -Value $False -Force
        $user | add-member -type noteproperty -name "IsLicensed" -Value $null -Force
        $user | add-member -type noteproperty -name "Licenses" -Value $null -Force
    }
}

$suncoastUsers | Export-Excel -Path "C:\Users\amedrano\Arraya Solutions\Suncoast Center, Inc - 1612 - Exchange Online Migration\Suncoast-Arraya Exchange Migration - External Share\Suncoast Users List.xlsx" -MoveToStart