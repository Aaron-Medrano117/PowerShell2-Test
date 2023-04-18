function Write-ProgressHelper {
	param (
	    [int]$ProgressCounter,
	    [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [int]$TotalCount
	)
    $secondsElapsed = (Get-Date) – $global:start
    $progresspercentcomplete = [math]::Round((($progresscounter / $TotalCount)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$TotalCount+"]"

    $progressParameters = @{
        Activity = $Activity
        Status = "$progressStatus $($secondsElapsed.ToString('hh\:mm\:ss'))"
        PercentComplete = $progresspercentcomplete
    }

    # if we have an estimate for the time remaining, add it to the Write-Progress parameters
    if ($secondsRemaining) {
        $progressParameters.SecondsRemaining = $secondsRemaining
    }
    if ($ID) {
        $progressParameters.ID = $ID
    }
    if ($CurrentOperation) {
        $progressParameters.CurrentOperation = $CurrentOperation
    }

    # Write the progress bar
    Write-Progress @progressParameters

    # estimate the time remaining
    $global:secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount – $progresscounter)
}

#Gather Mailbox Stats
$OutputCSVFolderPath = Read-Host "What is the folder path to store the file?"
$sourceMailboxes = Get-Mailbox -ResultSize Unlimited -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'"
$sourceMailboxStats = @()
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = $wavegroup.count

foreach ($user in $sourceMailboxes) {
    Write-ProgressHelper -Activity "Gathering Mailbox Details for $($user.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount

    #Pull MailboxStats and UserDetails
    $mbxStats = Get-MailboxStatistics $user.DistinguishedName
    $msoluser = Get-MsolUser -UserPrincipalName $user.UserPrincipalName
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    # Create User Hash Table
    $currentuser = @{
        "DisplayName" = $msoluser.DisplayName
        "UserPrincipalName" = $msoluser.userprincipalname
        "IsLicensed" = $msoluser.IsLicensed
        "Licenses" = ($msoluser.Licenses.AccountSkuID -join ";")
        "License-DisabledArray" = ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";")
        "BlockCredential" = $msoluser.BlockCredential
        "Department" = $msoluser.Department
        "RecipientTypeDetails" = $user.RecipientTypeDetails
        "PrimarySmtpAddress" = $user.PrimarySmtpAddress
        "Alias" = $user.alias
        "CustomAttribute7" = $user.CustomAttribute7
        "WhenCreated" = $user.WhenCreated
        "LastLogonTime" = $mbxStats.LastLogonTime
        "EmailAddresses" = ($EmailAddresses -join ";")
        "LegacyExchangeDN" = ("x500:" + $user.legacyexchangedn)
        "HiddenFromAddressListsEnabled" = $user.HiddenFromAddressListsEnabled
        "DeliverToMailboxAndForward" = $user.DeliverToMailboxAndForward
        "ForwardingAddress" = $user.ForwardingAddress
        "ForwardingSmtpAddress" = $user.ForwardingSmtpAddress
        "MBXSize" = $MBXStats.TotalItemSize
        "MBXSize_GB" = [math]::Round(($MBXStats.TotalItemSize.Value.ToBytes() / 1GB), 3)
        "MBXItemCount" = $MBXStats.ItemCount
        "LitigationHoldEnabled" = $user.LitigationHoldEnabled
        "LitigationHoldDuration" = $user.LitigationHoldDuration
        "InPlaceHolds" = $user.InPlaceHolds
    }

    #Pull Send on Behalf
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser["GrantSendOnBehalfTo"] = ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.DistinguishedName | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser["FullAccessPerms"] = ($mbxPermissions.User -join ";")
    }
    else {$currentuser["FullAccessPerms"] = ($null)}
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.DistinguishedName | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser["SendAsPerms"] = ($sendAsPermsCheck.trustee -join ";")
    }
    else {$currentuser["SendAsPerms"] = ($null)}
    # Archive Mailbox Check
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus
    if ($ArchiveStats = Get-MailboxStatistics $user.DistinguishedName -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {           
        $currentuser["ArchiveSize"] = $ArchiveStats.TotalItemSize.Value
        $currentuser["ArchiveSize-GB"] = $ArchiveStats.TotalItemSize.Value.ToBytes() / 1GB), 3)
        $currentuser["ArchiveItemCount"] = $ArchiveStats.ItemCount
    }
    else {
        $currentuser["ArchiveSize"] = $null
        $currentuser["ArchiveSize-GB"] = $null
        $currentuser["ArchiveItemCount"] = $null
    }
    $sourceMailboxStats += $currentuser
}
$sourceMailboxStats | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$OutputCSVFolderPath\SourceMailboxes.csv"