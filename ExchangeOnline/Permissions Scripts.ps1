## Get Mailbox Permissions

#Gather All Mailboxes
Write-Host "Gathering Mailboxes .." -foregroundcolor cyan -nonewline
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*DiscoverySearchMailbox*"} | sort PrimarySmtpAddress
Write-Host "done" -foregroundcolor green

#ProgressBar
$progressref = ($mailboxes).count
$progresscounter = 0

#Build Array
$allPermsList = @()
foreach ($mbx in $mailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Permissions for $($mbx.DisplayName)"
	
    $currentPerm = new-object PSObject        
    $currentPerm | add-member -type noteproperty -name "Mailbox" -Value $mbx.PrimarySmtpAddress.ToString()
    #Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $mbx.primarysmtpaddress | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentPerm | add-member -type noteproperty -name "FullAccessPerms" -Value ($mbxPermissions.user -join ",")
    }
    else {
        $currentPerm | add-member -type noteproperty -name "FullAccessPerms" -Value $null
    }
    # Mailbox Send As Check
	if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $mbx.PrimarySMTPAddress| ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentPerm | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ",")
    }
    else {
        $currentPerm | add-member -type noteproperty -name "SendAsPerms" -Value $null
    }
    $allPermsList += $currentPerm
}

## Get Distribution Permissions

#ProgressBar
$progressref = ($matchedDistributionGroups).count
$progresscounter = 0

#Build Array
$allPermsList = @()
foreach ($group in $matchedDistributionGroups)
{
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Distribution Group Permissions for $($group.DisplayName)"
	
    # Distribution Send As Check
	if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $group.PrimarySMTPAddress| ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $group | add-member -type noteproperty -name "SendAsPerms" -Value ($sendAsPermsCheck.trustee -join ",")
    }
    else {
        $group | add-member -type noteproperty -name "SendAsPerms" -Value $null
    }
}

## Stamp Send On Behalf Permissions
$matchedMailboxes = Import-Csv 
$sendOnBehalfMatchedUsers = $matchedMailboxes | ? {$_.GrantSendOnBehalfTo -and $_.ExistsInDestination -ne $false}
$failures = @()

#ProgressBar
$progressref = ($sendOnBehalfMatchedUsers).count
$progresscounter = 0
foreach ($mailbox in $sendOnBehalfMatchedUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Stamping Send On Behalf Permissions for $($mailbox.DisplayName_Destination)"
    Write-Host "Adding Send On Behalf Perms for $($mailbox.DisplayName_Destination)" -NoNewline -ForegroundColor Cyan
    $sendOnBehalfPerms = $mailbox.GrantSendOnBehalfTo -split ","
    foreach ($perm in $sendOnBehalfPerms) {
        try {
            $matchedUserArray = $matchedMailboxes | ?{$_.DisplayName -eq $perm}
            $matchedUPN = $matchedUserArray.UserPrincipalName_Destination
            Set-Mailbox -Identity $mailbox.PrimarySmtpAddress_Destination -GrantSendOnBehalfTo @{add=$matchedUPN} -ErrorAction Stop
            Write-Host " . " -ForegroundColor Green -NoNewline
        }
        catch {
            #Write-Warning -Message "$($_.Exception)"
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "TimeStamp" -Value ((Get-Date).ToString("MM-dd-yyyy"))
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantSendAsPerms" -Force
            $currenterror | add-member -type noteproperty -name "Source_User" -Value $mailbox.PrimarySmtpAddress
            $currenterror | add-member -type noteproperty -name "Destination_User" -Value $mailbox.PrimarySmtpAddress
            $currenterror | add-member -type noteproperty -name "Source_PermUser" -Value $perm
            $currenterror | add-member -type noteproperty -name "Destination_PermUser" -Value $matchedUPN
            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
            $currentError | Export-Csv -NoTypeInformation -encoding utf8 C:\Users\amedrano\Desktop\AllErrors.csv -Append
            $failures += $currenterror

            Write-Host " . " -ForegroundColor Red -NoNewline
        }
    }
    Write-Host "done" -ForegroundColor Green
}

## Find Not Found Users
foreach ($user in $failures){
    if ($mailboxCheck = Get-Mailbox $user.GrantSendonBehalf -ErrorAction SilentlyContinue) {
        Write-Host "Found user $($msolUserCheck.DisplayName)" -ForegroundColor Green
        $user | add-member -type noteproperty -name "Found" -Value $true -Force
        $user | add-member -type noteproperty -name "Found_DisplayName" -Value $mailboxCheck.DisplayName -Force
        $user | add-member -type noteproperty -name "Found_UPN" -Value $mailboxCheck.UserPrincipalName -Force
    }
    else {
        Write-Host "Did Not find user $($user.GrantSendonBehalf)" -ForegroundColor red
        $user | add-member -type noteproperty -name "Found" -Value $false -Force
        $user | add-member -type noteproperty -name "Found_DisplayName" -Value $null -Force
        $user | add-member -type noteproperty -name "Found_UPN" -Value $null -Force
    }
}

#Update Calendar Perm Users to DisplayName
#ProgressBar
$progressref = ($calendarpermsList).count
$progresscounter = 0
foreach ($calendar in $calendarpermsList) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    $calendarPath = $calendar.Mailbox + ":" + $calendar.CalendarPath
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Deatils for $($calendarPath)"
    if ($recipientCheck = Get-Mailbox $calendar.User -ea silentlycontinue) {
        $calendar | add-member -type noteproperty -name "User_DisplayName" -Value $recipientCheck.DisplayName -Force
        $calendar | add-member -type noteproperty -name "User_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress -Force
    }
    else {
        $calendar | add-member -type noteproperty -name "User_DisplayName" -Value $null -Force
        $calendar | add-member -type noteproperty -name "User_PrimarySMTPAddress" -Value $null -force
    }
}

#Stamp Calendar Perms
$matchedMailboxes = Import-Csv
$calendarPerms = import-csv
$AllErrorsPerms = @()

$progressref = ($calendarPerms | ? {$_.User_DisplayName}).count
$progresscounter = 0
foreach ($perm in ($calendarPerms | ? {$_.User_DisplayName})) {
    #Set Variables
    $calendarUser = @()
    $calendarPathSplit = $perm.CalendarPath -split ":"     
    $calendarUser = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $perm.User_PrimarySMTPAddress}
    $permUser = $matchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $perm.Mailbox}
    $calendarPath = $permUser.PrimarySmtpAddress_Destination + ":" + $calendarPathSplit[1]

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Stamping Perm $($perm.User_PrimarySMTPAddress) to $($calendarPath)"
    Write-Host "Grant $($perm.User_PrimarySMTPAddress) Calendar Perms for $($calendarPath).. " -ForegroundColor Cyan -NoNewline
    
    #Add Perms to Mailbox      
    try {
        $permResult = Add-MailboxFolderPermission -Identity $calendarPath -AccessRights $perm.AccessRights -User $calendarUser.PrimarySmtpAddress_Destination  -ea Stop
        Write-Host ". " -ForegroundColor Green
    }
    catch {
        Write-Host ". " -ForegroundColor red

        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "TimeStamp" -Value ((Get-Date).ToString("MM-dd-yyyy"))
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantCalendarPerms" -Force
        $currenterror | Add-Member -type NoteProperty -Name "Calendar" -Value $perm.CalendarPath -Force
        $currenterror | Add-Member -type NoteProperty -Name "Calendar_PrimarySMTPAddress_Destination" -Value $permUser.PrimarySmtpAddress_Destination -Force
        $currenterror | Add-Member -type NoteProperty -Name "PermUser_PrimarySMTPAddress_Source" -Value $perm.User_PrimarySMTPAddress -Force
        $currenterror | Add-Member -type NoteProperty -Name "PermUser_PrimarySMTPAddress_Destination" -Value $calendarUser.PrimarySmtpAddress_Destination -Force 
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllErrorsPerms += $currenterror
        $currentError | Export-Csv -NoTypeInformation -encoding utf8 "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\Source Exports\AllPermErrors.csv" -Append
        continue
    }
    Write-Host " done " -ForegroundColor Green
}

#Update Calendar Perm Users to DisplayName
#ProgressBar
$progressref = ($calendarpermsList).count
$progresscounter = 0
foreach ($calendar in $calendarpermsList) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    $calendarPath = $calendar.Mailbox + ":" + $calendar.CalendarPath
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Deatils for $($calendarPath)"
    if ($recipientCheck = Get-Mailbox $calendar.User -ea silentlycontinue) {
        $calendar | add-member -type noteproperty -name "User_DisplayName" -Value $recipientCheck.DisplayName -Force
        $calendar | add-member -type noteproperty -name "User_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress -Force
    }
    else {
        $calendar | add-member -type noteproperty -name "User_DisplayName" -Value $null -Force
        $calendar | add-member -type noteproperty -name "User_PrimarySMTPAddress" -Value $null -force
    }
}

#Stamp Delegate Perms
$allMatchedMailboxes = Import-Csv
$folderPerms = import-csv
$legitFolderPerms = $folderPerms | ?{$_.PermUser_DisplayName}
$AllErrorsPerms = @()

$progressref = $legitFolderPerms.count
$progresscounter = 0
foreach ($perm in $legitFolderPerms) {
    #Set Variables
    $calendarUser = @()
    $calendarPathSplit = $perm.CalendarPath -split ":"
    $calendarUser = $allMatchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $perm.Mailbox}  
    $permUser = $allMatchedMailboxes | ? {$_.PrimarySmtpAddress_Source -eq $perm.PermUser_PrimarySMTPAddress}
    $calendarPath = $permUser.PrimarySmtpAddress_Destination + ":" + $calendarPathSplit[1]
    $PermUserUPN = $permUser.UserPrincipalName_Destination
    $CalendarUserUPN = $calendarUser.UserPrincipalName_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Stamping Perm $($perm.PermUser_PrimarySMTPAddress) to $($calendarPath)"
    Write-Host "Grant $($PermUserUPN) Calendar Perms for $($calendarPath).. " -ForegroundColor Cyan -NoNewline
    
    #Add Perms to Mailbox      
    try {
        if ($perm.SharingPermissionFlags) {
            $permResult = Add-MailboxFolderPermission -Identity $calendarPath -AccessRights $perm.AccessRights -User $PermUserUPN -SharingPermissionFlags $perm.SharingPermissionFlags -ea Stop
        }
        else {
            $permResult = Add-MailboxFolderPermission -Identity $calendarPath -AccessRights $perm.AccessRights -User $PermUserUPN -ea Stop
        }
        
        Write-Host "." -ForegroundColor Green -nonewline
    }
    catch {
        Write-Host "." -ForegroundColor red

        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "TimeStamp" -Value ((Get-Date).ToString("MM-dd-yyyy"))
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToGrantCalendarPerms" -Force
        $currenterror | Add-Member -type NoteProperty -Name "Calendar" -Value $perm.CalendarPath -Force
        $currenterror | Add-Member -type NoteProperty -Name "Calendar_PrimarySMTPAddress_Destination" -Value $CalendarUserUPN -Force
        $currenterror | Add-Member -type NoteProperty -Name "PermUser_PrimarySMTPAddress_Source" -Value $permUser.User_PrimarySMTPAddress -Force
        $currenterror | Add-Member -type NoteProperty -Name "PermUser_PrimarySMTPAddress_Destination" -Value $permUser.PrimarySmtpAddress_Destination -Force 
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllErrorsPerms += $currenterror
        continue
    }
    Write-Host " done " -ForegroundColor Green
}