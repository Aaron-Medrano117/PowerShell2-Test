$allRecipients = Get-Recipient -ResultSize Unlimited 

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
    #$secondsElapsed = (Get-Date) – $global:start
    $global:secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($global:progressref – $progresscounter)
}

$SendAs = $true
$FullAccess = $true
$CalendarPermissions = $true
$SendOnBehalf = $true
$AllPerms = $true
$Office365 = $true

#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null

#Build Array
$CollectPermissionsList = @()
$failedUsers = @()

foreach ($obj in $allRecipients) {
    Write-ProgressHelper -ProgressCounter ($progresscounter++) -Activity "Gathering Details for $($obj.DisplayName)" -ID 1 -TotalCount ($allRecipients).count

    $primarySMTPAddress = $obj.PrimarySMTPAddress.tostring()
    $identity = $obj.Identity.ToString()
    $recipientTypeDetails = $obj.RecipientTypeDetails
    Write-Host "Checking Perms for $($obj.DisplayName)" -ForegroundColor Cyan -NoNewline
    Write-Host ".." -ForegroundColor Yellow -NoNewline
    if ($AllPerms) {
        $SendAs = $true
        $FullAccess = $true
        $CalendarPermissions = $true
        $SendOnBehalf = $true
    }
    #Clear Variables
    $SendAsPerms = @()
    $objPermissions = @()
    $calendarPerms = @()
    $SendOnBehalfToPerms = @()
    #SendAs Check
    if ($SendAs) {
        try {
            #Check Perms if OnPremises or Office365
            if ($OnPremises) {
                [array]$SendAsPerms = Get-User $obj.Identity -EA Stop | Get-ADPermission | Where {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "*S-1-*"}         
            }
            elseif ($Office365) {
                [array]$SendAsPerms = Get-RecipientPermission $obj.Identity -EA Stop | Where {$_.Trustee.ToString() -ne "NT Authority\Self" -and $_.Trustee.ToString() -notlike "*S-1-*"}
            }
        }
        catch {
            $failedUsers += $obj
            Write-Warning -Message "$($_.Exception.Message)"
        }        
        #If Permissions Found
        if ($SendAsPerms) {
            Write-Host "SendAs.." -NoNewline -foregroundcolor DarkCyan
            #Output Perms
            foreach ($perm in $SendAsPerms) {
                $accessRights = $perm.AccessRights -join ","
                #Check Perm User Mail Enabled; If OnPremises and If Office365
                if ($OnPremises) {
                    if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.User.ToString()
                        $permRecipientDetails = $null
                    }
                }
                elseif ($Office365) {
                    if ($recipientCheck = Get-Recipient $perm.Trustee.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.Trustee.ToString()
                        $permRecipientDetails = $null
                    }
                }
                #If Recipient Is Group, Output Group Members Individually
                if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                    $GroupPerms = $True
                }
                Else {
                    $GroupPerms = $False
                }
                $currentPerm = new-object PSObject
                $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "SendAs"
                $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                   
                $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $GroupPerms -Force
                $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                $CollectPermissionsList += $currentPerm

                Write-Host "." -ForegroundColor Yellow -NoNewline
            }
        }
    }
    #Full Access Check
    if ($FullAccess) {
        if ($recipientTypeDetails -like "*Mailbox" -and $recipientTypeDetails -ne "GroupMailbox") {
            #Check Perms if OnPremises or Office365
            try {
                if ($OnPremises) {
                    $objPermissions = Get-MailboxPermission $primarySMTPAddress -EA SilentlyContinue | Where {$_.User.ToString() -notlike "NT Authority*" -and $_.User.ToString() -notlike "*S-1-*" -and $_.AccessRights -like "FullAccess" -and $_.User.ToString() -notlike "*Exchange Domain Servers*" -and $_.User.ToString() -notlike "*Exchange Servers*" -and $_.User.ToString() -notlike "*Domain Admins*"}
                }
                elseif ($Office365) {
                    $objPermissions = Get-MailboxPermission $primarySMTPAddress -EA SilentlyContinue | Where {$_.User.ToString() -notlike "NT Authority*" -and $_.User.ToString() -notlike "*S-1-*"}
                    }
            }
            catch {
                $failedUsers += $obj
                Write-Warning -Message "$($_.Exception.Message)"
            } 
            #If Permissions Found
            if ($objPermissions) {
                Write-Host "FullAccess.." -NoNewline -foregroundcolor DarkCyan
                #Output Perms
                foreach ($perm in $objPermissions) {
                    $accessRights = $perm.AccessRights -join ","
                    #Check Perm User Mail Enabled
                    if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.User.ToString()
                        $permRecipientDetails = $null
                    }
                    #If Recipient Is Group, Output Group Members Individually   
                    if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                        $GroupPerms = $True
                    }
                    Else {
                        $GroupPerms = $False
                    }
   
                    $currentPerm = new-object PSObject
                    $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                    $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                    $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                    $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "FullAccess"
                    $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                    $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                   
                    $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                    $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $GroupPerms -Force
                    $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                    $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                    $CollectPermissionsList += $currentPerm
                    Write-Host "." -ForegroundColor DarkYellow -NoNewline
                }
            }
        }
    }
    #Calendar Permission Check
    if ($CalendarPermissions) {
        if ($recipientTypeDetails -like "*Mailbox" -and $recipientTypeDetails -ne "GroupMailbox") {
            try {
                [array]$calendars = Get-MailboxFolderStatistics $obj.identity | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}
            }
            catch {
                $failedUsers += $obj
                Write-Warning -Message "$($_.Exception.Message)"
            }
            #Gather Perms per Calendar
            foreach ($calendar in $calendars) {
                $folderPath = $calendar.FolderPath.Replace('/','\')
                $id = "$primarySMTPAddress" + ":$folderPath"
                #Gather Calendar Permissions; If On-Premises, If Office 365
                if ($OnPremises) {
                    [array]$calendarPerms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.User.ADRecipient -and $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" -and $_.User.ToString() -notlike "*S-1-*" -and $_.User.ADRecipient.PrimarySmtpAddress.ToString() -ne $primarySMTPAddress}
                }
                elseif ($Office365) {
                    [array]$calendarPerms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.user.Usertype.value -ne "Default" -and $_.user.usertype.value -ne "Anonymous" -and $_.user.usertype.value -notlike "*S-1-*"}
                }
                #output Per Calendar
                if ($calendarPerms) {
                    Write-Host "CalendarPerm.." -NoNewline -foregroundcolor DarkCyan
                    Write-Host $folderPath -ForegroundColor Green -NoNewline
                    #Output Perms
                    foreach ($perm in $calendarPerms) {
                        $accessRights = $perm.AccessRights -join ","
                        $SharingPermissionFlags = $perm.SharingPermissionFlags -join ","
                        $recipientCheck = @()
                        #Check Perm User Mail Enabled
                        if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                            $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                            $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                        }
                        else {
                            $permUser = $perm.User.ToString()
                            $permRecipientDetails = $null
                        }
                        #If Recipient Is Group, Output Group Members Individually   
                        if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                            $GroupPerms = $True
                        }
                        Else {
                            $GroupPerms = $False
                        }

                        $currentPerm = new-object PSObject
                        $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                        $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                        $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                        $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "Calendar"
                        $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                        $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                   
                        $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                        $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $GroupPerms -Force
                        $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $id
                        $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags
                        $CollectPermissionsList += $currentPerm
                        Write-Host "." -ForegroundColor DarkYellow -NoNewline
                    }
                }
            }
        }   
    }
    #Send on Behalf Check
    if ($SendOnBehalf) {
        #If Recipient Detail is Mailbox
        if ($recipientTypeDetails -like "*Mailbox" -and $recipientTypeDetails -ne "GroupMailbox") {
            try {
                $SendOnBehalfToPerms = (Get-Mailbox $primarySMTPAddress -ErrorAction Stop).GrantSendOnBehalfTo
            }
            catch {
                $failedUsers += $obj
                Write-Warning -Message "$($_.Exception.Message)"
            }
            if ($SendOnBehalfToPerms) {
                Write-Host "SendOnBehalfTo.." -NoNewline -foregroundcolor DarkCyan
                #Output Perms
                foreach ($perm in $SendOnBehalfToPerms) {
                    #Check Perm User Mail Enabled
                    if ($OnPremises) {
                        if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                            $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                            $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                        }
                        else {
                            $permUser = $perm.User.ToString()
                            $permRecipientDetails = $null
                        }
                    }
                    elseif ($Office365) {
                        if ($recipientCheck = Get-Recipient $perm.ToString() -ea silentlycontinue) {
                            $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                            $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                        }
                        else {
                            $permUser = $perm.ToString()
                            $permRecipientDetails = $null
                        }
                    }
                    #If Recipient Is Group, Output Group Members Individually 
                    if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                        $GroupPerms = $True
                    }
                    Else {
                        $GroupPerms = $False
                    }

                    $currentPerm = new-object PSObject
                    $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                    $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                    $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                    $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "SendOnBehalf"
                    $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                    $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                   
                    $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                    $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $GroupPerms -Force
                    $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                    $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                    $CollectPermissionsList += $currentPerm
                    Write-Host "." -ForegroundColor DarkYellow -NoNewline
                }
            }
        } 
    }
    Write-Host "done" -ForegroundColor Green
}

Write-host ""
Write-Host $failedUsers.count 'Recipients Generated Errors. Advise Re-running with Failed Users in $failedUsers variable' -ForegroundColor Cyan
$CollectPermissionsList | Export-Excel "$HOME\Desktop\AllSpectra-DelegatePermissions-012023.xlsx"

if ($OutputCSVFolderPath) {
    $CollectPermissionsList | Export-Csv "$OutputCSVFolderPath\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
    Write-host "Exported file 'DelegatePermissions.csv' List to $OutputCSVFolderPath" -ForegroundColor Cyan
}
elseif ($OutputCSVFilePath) {
    $CollectPermissionsList | Export-Csv $OutputCSVFilePath -NoTypeInformation -Encoding UTF8
    Write-host "Exported Permissions List to $OutputCSVFilePath" -ForegroundColor Cyan
}
else {
    try {
        $CollectPermissionsList | Export-Csv "$HOME\Desktop\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
    }
    catch {}
}

# Failed Users
$AllPerms = $true
$Office365 = $true

#Build Array
$failedUsers2 = @()
#ProgressBar
$progressref = ($failedUsers).count
$progresscounter = 0

foreach ($obj in $failedUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details for $($obj.DisplayName)"
    $primarySMTPAddress = $obj.PrimarySMTPAddress.tostring()
    $identity = $obj.Identity.ToString()
    $recipientTypeDetails = $obj.RecipientTypeDetails
    Write-Host "Checking Perms for $($obj.DisplayName)" -ForegroundColor Cyan -NoNewline
    Write-Host ".." -ForegroundColor Yellow -NoNewline
    if ($AllPerms) {
        $SendAs = $true
        $FullAccess = $true
        $CalendarPermissions = $true
        $SendOnBehalf = $true
    }
    #Clear Variables
    $SendAsPerms = @()
    $objPermissions = @()
    $calendarPerms = @()
    $SendOnBehalfToPerms = @()
    #SendAs Check
    if ($SendAs) {
        try {
            #Check Perms if OnPremises or Office365
            if ($OnPremises) {
                [array]$SendAsPerms = Get-User $obj.Identity -EA Stop | Get-ADPermission | Where {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "*S-1-*"}         
            }
            elseif ($Office365) {
                [array]$SendAsPerms = Get-RecipientPermission $obj.Identity -EA Stop | Where {$_.Trustee.ToString() -ne "NT Authority\Self" -and $_.Trustee.ToString() -notlike "*S-1-*"}
            }
        }
        catch {
            $failedUsers2 += $obj
            Write-Warning -Message "$($_.Exception.Message)"
        }        
        #If Permissions Found
        if ($SendAsPerms) {
            Write-Host "SendAs.." -NoNewline -foregroundcolor DarkCyan
            #Output Perms
            foreach ($perm in $SendAsPerms) {
                $accessRights = $perm.AccessRights -join ","
                #Check Perm User Mail Enabled; If OnPremises and If Office365
                if ($OnPremises) {
                    if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.User.ToString()
                        $permRecipientDetails = $null
                    }
                }
                elseif ($Office365) {
                    if ($recipientCheck = Get-Recipient $perm.Trustee.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.Trustee.ToString()
                        $permRecipientDetails = $null
                    }
                }
                #If Recipient Is Group, Output Group Members Individually
                if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                    $GroupPerms = $True
                }
                Else {
                    $GroupPerms = $False
                }
                $currentPerm = new-object PSObject
                $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "SendAs"
                $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                   
                $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $GroupPerms -Force
                $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                $CollectPermissionsList += $currentPerm

                Write-Host "." -ForegroundColor Yellow -NoNewline
            }
        }
    }
    #Full Access Check
    if ($FullAccess) {
        if ($recipientTypeDetails -like "*Mailbox" -and $recipientTypeDetails -ne "GroupMailbox") {
            #Check Perms if OnPremises or Office365
            try {
                if ($OnPremises) {
                    $objPermissions = Get-MailboxPermission $primarySMTPAddress -EA SilentlyContinue | Where {$_.User.ToString() -notlike "NT Authority*" -and $_.User.ToString() -notlike "*S-1-*" -and $_.AccessRights -like "FullAccess" -and $_.User.ToString() -notlike "*Exchange Domain Servers*" -and $_.User.ToString() -notlike "*Exchange Servers*" -and $_.User.ToString() -notlike "*Domain Admins*"}
                }
                elseif ($Office365) {
                    $objPermissions = Get-MailboxPermission $primarySMTPAddress -EA SilentlyContinue | Where {$_.User.ToString() -notlike "NT Authority*" -and $_.User.ToString() -notlike "*S-1-*"}
                    }
            }
            catch {
                $failedUsers2 += $obj
                Write-Warning -Message "$($_.Exception.Message)"
            } 
            #If Permissions Found
            if ($objPermissions) {
                Write-Host "FullAccess.." -NoNewline -foregroundcolor DarkCyan
                #Output Perms
                foreach ($perm in $objPermissions) {
                    $accessRights = $perm.AccessRights -join ","
                    #Check Perm User Mail Enabled
                    if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                        $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                        $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                    }
                    else {
                        $permUser = $perm.User.ToString()
                        $permRecipientDetails = $null
                    }
                    #If Recipient Is Group, Output Group Members Individually   
                    if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                        $GroupPerms = $True
                    }
                    Else {
                        $GroupPerms = $False
                    }
   
                    $currentPerm = new-object PSObject
                    $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                    $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                    $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                    $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "FullAccess"
                    $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                    $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                   
                    $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                    $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $GroupPerms -Force
                    $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                    $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                    $CollectPermissionsList += $currentPerm
                    Write-Host "." -ForegroundColor DarkYellow -NoNewline
                }
            }
        }
    }
    #Calendar Permission Check
    if ($CalendarPermissions) {
        if ($recipientTypeDetails -like "*Mailbox" -and $recipientTypeDetails -ne "GroupMailbox") {
            try {
                [array]$calendars = Get-MailboxFolderStatistics $obj.identity | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}
            }
            catch {
                $failedUsers2 += $obj
                Write-Warning -Message "$($_.Exception.Message)"
            }
            #Gather Perms per Calendar
            foreach ($calendar in $calendars) {
                $folderPath = $calendar.FolderPath.Replace('/','\')
                $id = "$primarySMTPAddress" + ":$folderPath"
                #Gather Calendar Permissions; If On-Premises, If Office 365
                if ($OnPremises) {
                    [array]$calendarPerms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.User.ADRecipient -and $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" -and $_.User.ToString() -notlike "*S-1-*" -and $_.User.ADRecipient.PrimarySmtpAddress.ToString() -ne $primarySMTPAddress}
                }
                elseif ($Office365) {
                    [array]$calendarPerms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.user.Usertype.value -ne "Default" -and $_.user.usertype.value -ne "Anonymous" -and $_.user.usertype.value -notlike "*S-1-*"}
                }
                #output Per Calendar
                if ($calendarPerms) {
                    Write-Host "CalendarPerm.." -NoNewline -foregroundcolor DarkCyan
                    Write-Host $folderPath -ForegroundColor Green -NoNewline
                    #Output Perms
                    foreach ($perm in $calendarPerms) {
                        $accessRights = $perm.AccessRights -join ","
                        $SharingPermissionFlags = $perm.SharingPermissionFlags -join ","
                        $recipientCheck = @()
                        #Check Perm User Mail Enabled
                        if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                            $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                            $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                        }
                        else {
                            $permUser = $perm.User.ToString()
                            $permRecipientDetails = $null
                        }
                        #If Recipient Is Group, Output Group Members Individually   
                        if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                            $GroupPerms = $True
                        }
                        Else {
                            $GroupPerms = $False
                        }

                        $currentPerm = new-object PSObject
                        $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                        $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                        $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                        $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "Calendar"
                        $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                        $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                   
                        $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                        $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $GroupPerms -Force
                        $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $id
                        $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags
                        $CollectPermissionsList += $currentPerm
                        Write-Host "." -ForegroundColor DarkYellow -NoNewline
                    }
                }
            }
        }   
    }
    #Send on Behalf Check
    if ($SendOnBehalf) {
        #If Recipient Detail is Mailbox
        if ($recipientTypeDetails -like "*Mailbox" -and $recipientTypeDetails -ne "GroupMailbox") {
            try {
                $SendOnBehalfToPerms = (Get-Mailbox $primarySMTPAddress -ErrorAction Stop).GrantSendOnBehalfTo
            }
            catch {
                $failedUsers2 += $obj
                Write-Warning -Message "$($_.Exception.Message)"
            }
            if ($SendOnBehalfToPerms) {
                Write-Host "SendOnBehalfTo.." -NoNewline -foregroundcolor DarkCyan
                #Output Perms
                foreach ($perm in $SendOnBehalfToPerms) {
                    #Check Perm User Mail Enabled
                    if ($OnPremises) {
                        if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                            $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                            $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                        }
                        else {
                            $permUser = $perm.User.ToString()
                            $permRecipientDetails = $null
                        }
                    }
                    elseif ($Office365) {
                        if ($recipientCheck = Get-Recipient $perm.ToString() -ea silentlycontinue) {
                            $permUser = $recipientCheck.PrimarySMTPAddress.ToString()
                            $permRecipientDetails = $recipientCheck.RecipientTypeDetails
                        }
                        else {
                            $permUser = $perm.ToString()
                            $permRecipientDetails = $null
                        }
                    }
                    #If Recipient Is Group, Output Group Members Individually 
                    if ($recipientCheck.RecipientTypeDetails -like "*Group*") {
                        $GroupPerms = $True
                    }
                    Else {
                        $GroupPerms = $False
                    }

                    $currentPerm = new-object PSObject
                    $currentPerm | add-member -type noteproperty -name "Identity" -Value $identity
                    $currentPerm | add-member -type noteproperty -name "MailObject" -Value $primarySMTPAddress
                    $currentPerm | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientTypeDetails
                    $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "SendOnBehalf"
                    $currentPerm | add-member -type noteproperty -name "PermUser" -Value $permUser -Force
                    $currentPerm | add-member -type noteproperty -name "PermRecipientTypeDetails" -Value $permRecipientDetails -Force                   
                    $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                    $currentPerm | add-member -type noteproperty -name "GroupPermission" -Value $GroupPerms -Force
                    $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                    $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                    $CollectPermissionsList += $currentPerm
                    Write-Host "." -ForegroundColor DarkYellow -NoNewline
                }
            }
        } 
    }
    Write-Host "done" -ForegroundColor Green
}

Write-host ""
Write-Host $failedUsers2.count 'Recipients Generated Errors. Advise Re-running with Failed Users in $failedUsers2 variable' -ForegroundColor Cyan