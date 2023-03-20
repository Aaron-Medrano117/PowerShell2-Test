# Create Inactive Mailboxes - EHN
$inactivemailboxes

$AllErrors = @()
$progressref = ($inactivemailboxes).count
$progresscounter = 0
foreach ($mailbox in $inactivemailboxes) {
    $newInactiveEHNAddress = ($mailbox.MicrosoftOnlineServicesID -split "@")[0] + "-inactive@EHN.MAIL.ONMICROSOFT.COM"
    $newInactiveDisplayName = $mailbox.DisplayName + "-Inactive"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Inactive Mailbox $($newInactiveEHNAddress)"
    
    Start-Sleep -Milliseconds 400
    if (!($recipientCheck = Get-Recipient $newInactiveEHNAddress -ea silentlycontinue)) {
        Write-Host "Creating New Mailbox $($newInactiveEHNAddress).." -ForegroundColor Cyan -NoNewline
        try {
            $newInactiveMailbox = New-Mailbox -InactiveMailbox $mailbox.ExchangeGuid -PrimarySmtpAddress $newInactiveEHNAddress -DisplayName $newInactiveDisplayName -Name $newInactiveDisplayName -MicrosoftOnlineServicesID $newInactiveEHNAddress -Password (ConvertTo-SecureString -String '#Einstein2022!' -AsPlainText -Force) -ErrorAction Stop
            do {
                Start-Sleep -Seconds 5
                write-host "." -ForegroundColor Yellow -NoNewline
            } until ($recipientCheck = Get-Recipient $newInactiveEHNAddress -ea silentlycontinue)
            Set-Mailbox $newInactiveEHNAddress -Type Shared
            Write-Host "Completed" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnableToCreateInactiveMailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $mailbox.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "UPN" -Value $mailbox.MicrosoftOnlineServicesID -Force
            $currenterror | Add-Member -type NoteProperty -Name "newInactiveEHNAddress" -Value $newInactiveEHNAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "ExchangeGuid" -Value $mailbox.ExchangeGuid -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception.Message) -Force
			$AllErrors += $currenterror 
        }
    }
}
$AllErrors | Export-csv -NoTypeInformation -Encoding utf8 "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxesErrors5.csv"


# Create Inactive Mailboxes - TJUV (Remaining)
$remainingMailboxes = $inactivemailboxes | ?{$_.TJUVAddress -eq $null}
$progressref = ($remainingMailboxes).count
$progresscounter = 0
$AllErrors = @()
foreach ($mailbox in $remainingMailboxes) {
    $OnMicrosoftMailDomain = "tjuv.mail.onmicrosoft.com"
    $newInactiveTJUVAddress = ($mailbox.MicrosoftOnlineServicesID -split "@")[0] + "-inactive@" + $OnMicrosoftMailDomain
    $newInactiveDisplayName = $mailbox.Name + "-inactive"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Inactive Mailbox $($newInactiveTJUVAddress)"
    
    Start-Sleep -Milliseconds 400
    if (!($recipientCheck = Get-Recipient $newInactiveTJUVAddress -ea silentlycontinue)) {
        Write-Host "Creating New Mailbox $($newInactiveTJUVAddress).." -ForegroundColor Cyan -NoNewline
        try {
            #Create Mailbox
            $newInactiveMailbox = New-Mailbox -Shared -PrimarySmtpAddress $newInactiveTJUVAddress -DisplayName $newInactiveDisplayName -Name $newInactiveDisplayName -ErrorAction Stop

            #Hide From GAL
            do {
                Start-Sleep -Seconds 5
                write-host "." -ForegroundColor Yellow -NoNewline
            } until ($recipientCheck = Get-Recipient $newInactiveTJUVAddress -ea silentlycontinue)
            Set-Mailbox $newInactiveTJUVAddress -HiddenFromAddressListsEnabled $true
            Write-Host "Completed" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnableToCreateInactiveMailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $mailbox.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "UPN" -Value $mailbox.MicrosoftOnlineServicesID -Force
            $currenterror | Add-Member -type NoteProperty -Name "newInactiveEHNAddress" -Value $newInactiveTJUVAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "ExchangeGuid" -Value $mailbox.ExchangeGuid -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror 
        }
    }
}
$AllErrors | Export-csv -NoTypeInformation -Encoding utf8 C:\Users\aaron.medrano\Desktop\InactiveMailboxesErrors3.csv


# Create Inactive Mailboxes - TJUV (Remaining) 2
$remainingMailboxes =
$progressref = ($remainingMailboxes).count
$progresscounter = 0
$AllErrors = @()
foreach ($mailbox in $remainingMailboxes) {
    $OnMicrosoftMailDomain = "tjuv.mail.onmicrosoft.com"
    $newInactiveTJUVAddress = ($mailbox.EHNAddress -split "@")[0] + "-inactive@" + $OnMicrosoftMailDomain
    $newInactiveDisplayName = $mailbox.EHNDisplayName + "-inactive"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Inactive Mailbox $($newInactiveTJUVAddress)"
    
    Start-Sleep -Milliseconds 400
    if (!($recipientCheck = Get-Recipient $newInactiveTJUVAddress -ea silentlycontinue)) {
        Write-Host "Creating New Mailbox $($newInactiveTJUVAddress).." -ForegroundColor Cyan -NoNewline
        try {
            #Create Mailbox
            $newInactiveMailbox = New-Mailbox -Shared -PrimarySmtpAddress $newInactiveTJUVAddress -DisplayName $newInactiveDisplayName -Name $newInactiveDisplayName -ErrorAction Stop

            #Hide From GAL
            do {
                Start-Sleep -Seconds 5
                write-host "." -ForegroundColor Yellow -NoNewline
            } until ($recipientCheck = Get-Recipient $newInactiveTJUVAddress -ea silentlycontinue)
            Set-Mailbox $newInactiveTJUVAddress -HiddenFromAddressListsEnabled $true
            Write-Host "Completed" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnableToCreateInactiveMailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $mailbox.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "UPN" -Value $mailbox.MicrosoftOnlineServicesID -Force
            $currenterror | Add-Member -type NoteProperty -Name "newInactiveEHNAddress" -Value $newInactiveTJUVAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "ExchangeGuid" -Value $mailbox.ExchangeGuid -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
			$AllErrors += $currenterror 
        }
    }
}
$AllErrors | Export-csv -NoTypeInformation -Encoding utf8 C:\Users\aaron.medrano\Desktop\InactiveMailboxesErrors3.csv

# Create Inactive Mailboxes - EHN (Remaining) 1
$failedInactiveMailboxes = Import-Csv "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxes\InactiveMailboxesErrors4.csv"

$AllErrors = @()
$nameConflictInactiveMailboxes = $failedInactiveMailboxes | ?{$_.ConflictType -eq "Name"}
$progressref = ($nameConflictInactiveMailboxes).count
$progresscounter = 0
foreach ($mailbox in $nameConflictInactiveMailboxes) {
    $newInactiveEHNAddress = $mailbox.newInactiveEHNAddress
    $newInactiveDisplayName = $mailbox.User + "-Inactive"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Inactive Mailbox $($newInactiveEHNAddress)"
    
    Start-Sleep -Milliseconds 400
    if (!($recipientCheck = Get-Recipient $newInactiveEHNAddress -ea silentlycontinue)) {
        Write-Host "Creating New Mailbox $($newInactiveEHNAddress).." -ForegroundColor Cyan -NoNewline
        try {
            $newInactiveMailbox = New-Mailbox -InactiveMailbox $mailbox.ExchangeGuid -PrimarySmtpAddress $newInactiveEHNAddress -DisplayName $newInactiveDisplayName -Name $newInactiveDisplayName -MicrosoftOnlineServicesID $newInactiveEHNAddress -Password (ConvertTo-SecureString -String '#Einstein2022!' -AsPlainText -Force) -ErrorAction Stop
            do {
                Start-Sleep -Seconds 5
                write-host "." -ForegroundColor Yellow -NoNewline
            } until ($recipientCheck = Get-Recipient $newInactiveEHNAddress -ea silentlycontinue)
            Set-Mailbox $newInactiveEHNAddress -Type Shared
            Write-Host "Completed" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnableToCreateInactiveMailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $mailbox.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "UPN" -Value $mailbox.MicrosoftOnlineServicesID -Force
            $currenterror | Add-Member -type NoteProperty -Name "newInactiveEHNAddress" -Value $newInactiveEHNAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "ExchangeGuid" -Value $mailbox.ExchangeGuid -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception.Message) -Force
			$AllErrors += $currenterror 
        }
    }
}
$AllErrors | Export-csv -NoTypeInformation -Encoding utf8 "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxesErrors5.csv"

# Create Remaining Inactive Mailboxes - EHN 2
$remaininginactivemailboxes

$AllErrors = @()
$progressref = ($remaininginactivemailboxes).count
$progresscounter = 0
foreach ($mailbox in $remaininginactivemailboxes) {
    $newInactiveEHNAddress = ($mailbox.MicrosoftOnlineServicesID -split "@")[0] + "-inactive@EHN.MAIL.ONMICROSOFT.COM"
    $newInactiveDisplayName = $mailbox.DisplayName + "-Inactive"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Inactive Mailbox $($newInactiveEHNAddress)"
    
    Start-Sleep -Milliseconds 400
    if (!($recipientCheck = Get-Recipient $newInactiveEHNAddress -ea silentlycontinue)) {
        Write-Host "Creating New Mailbox $($newInactiveEHNAddress).." -ForegroundColor Cyan -NoNewline
        try {
            $newInactiveMailbox = New-Mailbox -InactiveMailbox $mailbox.ExchangeGuid -PrimarySmtpAddress $newInactiveEHNAddress -DisplayName $newInactiveDisplayName -Name $newInactiveDisplayName -MicrosoftOnlineServicesID $newInactiveEHNAddress -Password (ConvertTo-SecureString -String '#Einstein2022!' -AsPlainText -Force) -ErrorAction Stop
            do {
                Start-Sleep -Seconds 5
                write-host "." -ForegroundColor Yellow -NoNewline
            } until ($recipientCheck = Get-Recipient $newInactiveEHNAddress -ea silentlycontinue)
            Set-Mailbox $newInactiveEHNAddress -Type Shared
            Write-Host "Completed" -ForegroundColor Green
            $mailbox | add-member -type noteproperty -name "EHN_DisplayName" -Value $recipientCheck.DisplayName -force
            $mailbox | add-member -type noteproperty -name "EHN_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress -force
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnableToCreateInactiveMailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $mailbox.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "UPN" -Value $mailbox.MicrosoftOnlineServicesID -Force
            $currenterror | Add-Member -type NoteProperty -Name "newInactiveEHNAddress" -Value $newInactiveEHNAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "ExchangeGuid" -Value $mailbox.ExchangeGuid -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception.Message) -Force
			$AllErrors += $currenterror 
        }
    }
    else {
        $mailbox | add-member -type noteproperty -name "EHN_DisplayName" -Value $recipientCheck.DisplayName -force
        $mailbox | add-member -type noteproperty -name "EHN_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress -force
 }
}
$AllErrors | Export-csv -NoTypeInformation -Encoding utf8 "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxesErrors6.csv"

#Create Mailbox and MailboxRestore Request for Address Matches - Remaining EHN Mailboxes
$AllErrors = @()

$proxyInactiveMailboxes = $failedInactiveMailboxes | ?{$_.ConflictType -like "*proxy"}
$progresscounter = 0
$progressref = ($proxyInactiveMailboxes).count
foreach ($mailbox in $proxyInactiveMailboxes) {
    $newInactiveEHNAddress = $mailbox.newInactiveEHNAddress
    $newInactiveDisplayName = $mailbox.User + "-Inactive"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Inactive Mailbox $($newInactiveEHNAddress) and MailboxRestoreRequest"
    
    Start-Sleep -Milliseconds 400
    if (!($recipientCheck = Get-Recipient $newInactiveEHNAddress -ea silentlycontinue)) {
        Write-Host "Creating New Mailbox $($newInactiveEHNAddress).." -ForegroundColor Cyan -NoNewline
        try {
            $newInactiveMailbox = New-Mailbox -Shared -PrimarySmtpAddress $newInactiveEHNAddress -DisplayName $newInactiveDisplayName -Name $newInactiveDisplayName -ErrorAction Stop
            do {
                Start-Sleep -Seconds 5
                write-host "." -ForegroundColor Yellow -NoNewline
            } until ($recipientCheck = Get-Recipient $newInactiveEHNAddress -ea silentlycontinue)
            ## Create New Mailbox Restore Request
            Write-Host "Creating Mailbox Restore Request.." -ForegroundColor Cyan -NoNewline
            $restoreJobResults = New-MailboxRestoreRequest -SourceMailbox $mailbox.ExchangeGuid -TargetMailbox $newInactiveEHNAddress -BatchName "InactiveMailboxes" -AllowLegacyDNMismatch -ErrorAction Stop
            Write-Host "Completed" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnableToCreateInactiveMailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $mailbox.User -Force
            $currenterror | Add-Member -type NoteProperty -Name "UPN" -Value $mailbox.UPN -Force
            $currenterror | Add-Member -type NoteProperty -Name "newInactiveEHNAddress" -Value $newInactiveEHNAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "ExchangeGuid" -Value $mailbox.ExchangeGuid -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception.Message) -Force
			$AllErrors += $currenterror 
        }
    }
}

#Create MailboxRestore Request for Address Matches - Remaining EHN Mailboxes
$AllErrors2 = @()
$progresscounter = 0
$progressref = ($failedInactiveMailboxes).count
foreach ($mailbox in $failedInactiveMailboxes) {
    $newInactiveEHNAddress = $mailbox.newInactiveEHNAddress
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Mailbox Restore Request $($newInactiveEHNAddress)"
    
    Start-Sleep -Milliseconds 400
    if (($mbxCheck = Get-Mailbox $newInactiveEHNAddress -ea silentlycontinue)) {
        $mailbox | add-member -type noteproperty -name "MailboxFound" -Value $true -force
        if (!($mbxRestoreRequestCheck = Get-MailboxRestoreRequest $newInactiveEHNAddress -ErrorAction silentlyContinue)) {
            $mailbox | add-member -type noteproperty -name "MBXRestoreJob" -Value $true -Force
            try {
                ## Create New Mailbox Restore Request
                Write-Host "Creating Mailbox Restore Request $($newInactiveEHNAddress).." -ForegroundColor Cyan -NoNewline
                $restoreJobResults = New-MailboxRestoreRequest -SourceMailbox $mailbox.ExchangeGuid -TargetMailbox $mbxCheck.ExchangeGuid.tostring() -BatchName "InactiveMailboxes" -AllowLegacyDNMismatch -ErrorAction Stop -warningaction silentlyContinue
                $mailbox | add-member -type noteproperty -name "MBXRestore_Status" -Value $restoreJobResults.Status -Force
                Write-Host "Completed" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed" -ForegroundColor Red
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnableToCreateInactiveMailbox" -Force
                $currenterror | Add-Member -type NoteProperty -Name "User" -Value $mailbox.User -Force
                $currenterror | Add-Member -type NoteProperty -Name "UPN" -Value $mailbox.UPN -Force
                $currenterror | Add-Member -type NoteProperty -Name "newInactiveEHNAddress" -Value $newInactiveEHNAddress -Force
                $currenterror | Add-Member -type NoteProperty -Name "ExchangeGuid" -Value $mailbox.ExchangeGuid -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception.Message) -Force
                $mailbox | add-member -type noteproperty -name "MBXRestore_Status" -Value $_.Exception.Message -Force
                $AllErrors2 += $currenterror 
            }
        }
        else {
            $mailbox | add-member -type noteproperty -name "MBXRestoreJob" -Value $true -Force
            $mailbox | add-member -type noteproperty -name "MBXRestore_Status" -Value $mbxRestoreRequestCheck.Status -Force
            Write-Host "Mailbox Restore found for $($mbxRestoreRequestCheck.TargetMailbox)" -foregroundcolor DarkCyan
        }
        
    }
    else {
        $mailbox | add-member -type noteproperty -name "MailboxFound" -Value $False -force
        $mailbox | add-member -type noteproperty -name "MBXRestoreJob" -Value $False -Force
        $mailbox | add-member -type noteproperty -name "MBXRestore_Status" -Value $null -Force

       Write-Host "No Mailbox found for $($newInactiveEHNAddress)" -ForegroundColor Red
    }
}

#Create MailboxRestore Request for Address Matches - Remaining EHN Mailboxes
$AllErrors2 = @()
$progresscounter = 0
$progressref = ($InactiveMailboxes).count
foreach ($mailbox in $InactiveMailboxes) {
    $newInactiveEHNAddress = $mailbox.newInactiveEHNAddress
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Mailbox Restore Request $($newInactiveEHNAddress)"
    
    Start-Sleep -Milliseconds 400
    if ($mailbox.RestoreMethod2 -eq "NotFound") {
        try {
            ## Create New Mailbox Restore Request
            $mailboxRestoreJobName = $mailbox.Name +"-Inactive2"
            Write-Host "Creating Mailbox Restore Request $($mailbox.EHNAddress3.ToString()).." -ForegroundColor Cyan -NoNewline
            $restoreJobResults = New-MailboxRestoreRequest -SourceMailbox $mailbox.ExchangeGuid -TargetMailbox $mailbox.EHNAddress3.ToString() -BatchName "InactiveMailboxes" -Name $mailboxRestoreJobName -AllowLegacyDNMismatch -ErrorAction Stop -warningaction silentlyContinue
            $mailbox | add-member -type noteproperty -name "MBXRestore_Status" -Value $restoreJobResults.Status -Force
            Write-Host "Completed" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnableToCreateInactiveMailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $mailbox.User -Force
            $currenterror | Add-Member -type NoteProperty -Name "UPN" -Value $mailbox.UPN -Force
            $currenterror | Add-Member -type NoteProperty -Name "newInactiveEHNAddress" -Value $newInactiveEHNAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "ExchangeGuid" -Value $mailbox.ExchangeGuid -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception.Message) -Force
            $mailbox | add-member -type noteproperty -name "MBXRestore_Status" -Value $_.Exception.Message -Force
            $AllErrors2 += $currenterror 
        }
    }
    
}        

# Get MailboxStats for Failed InActive Mailboxes
$failedInactiveMailboxes = Import-Csv "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxes\InactiveMailboxesErrors4.csv"
$progressref = ($failedInactiveMailboxes).count
$progresscounter = 0
foreach ($mailbox in $failedInactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking MailboxStats for Inactive Mailbox $($mailbox.UPN)"
    
    # Get MailboxStats of Deleted Object
    if ($inactiveMailboxStats = Get-MailboxStatistics $mailbox.ExchangeGuid -IncludeSoftDeletedRecipients | select TotalItemSize, ItemCount) {
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSize" -Value $inactiveMailboxStats.TotalItemSize -force
        $mailbox | add-member -type noteproperty -name "InActive_ItemCount" -Value $inactiveMailboxStats.ItemCount -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSize" -Value $null -force
        $mailbox | add-member -type noteproperty -name "InActive_ItemCount" -Value $null -force
    }
    
    #Get MailboxStats of Existing Object
    if ($conflictMailboxStats = Get-MailboxStatistics $mailbox.UPN -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount) {
        $mailbox | add-member -type noteproperty -name "Conflicting_TotalItemSize" -Value $conflictMailboxStats.TotalItemSize -force
        $mailbox | add-member -type noteproperty -name "Conflicting_ItemCount" -Value $conflictMailboxStats.ItemCount -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "Conflicting_TotalItemSize" -Value $null -force
        $mailbox | add-member -type noteproperty -name "Conflicting_ItemCount" -Value $null -force
    }
    #Get MailboxWhenDeleted
    if ($MailboxDetails = Get-Mailbox $mailbox.ExchangeGuid -IncludeInactiveMailbox | select Identity, WhenSoftDeleted, WasInactiveMailbox, IsInactiveMailbox) {
        $mailbox | add-member -type noteproperty -name "Identity_EHN" -Value $MailboxDetails.TotalItemSize -force
        $mailbox | add-member -type noteproperty -name "WhenSoftDeleted_EHN" -Value $MailboxDetails.WhenSoftDeleted -force
        $mailbox | add-member -type noteproperty -name "WasInactiveMailbox_EHN" -Value $MailboxDetails.WasInactiveMailbox -force
        $mailbox | add-member -type noteproperty -name "IsInactiveMailbox_EHN" -Value $MailboxDetails.IsInactiveMailbox -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "Identity_EHN" -Value $null -force
        $mailbox | add-member -type noteproperty -name "WhenSoftDeleted_EHN" -Value $null -force
        $mailbox | add-member -type noteproperty -name "WasInactiveMailbox_EHN" -Value $null -force
        $mailbox | add-member -type noteproperty -name "IsInactiveMailbox_EHN" -Value $null -force
    }
}

#check remaining inactive mailboxes in migration 1
$InactiveMailboxes = import-csv
$inactiveMailboxMigrations = import-csv
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
$noMigProject = @()
foreach ($mailbox in $InactiveMailboxes) {
    $inactiveMailbox = $mailbox.EHNAddress
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Inactive Mailbox project $($inactiveMailbox)"
    
    if (!($migmailboxCheck = $inactiveMailboxMigrations | ? {$_.SourceEmailAddress -eq $inactiveMailbox})) {
        $noMigProject += $mailbox
    }
}
$noMigProject | Export-Csv -NoTypeInformation -encoding utf8

#check if main inactive mailboxes in migration 2
$noMigProject = import-csv
$inactiveMailboxMigrations = import-csv
$progressref = ($noMigProject).count
$progresscounter = 0
$noMigProject2 = @()
foreach ($mailbox in $noMigProject) {
    $inactiveMailbox = $mailbox.TJUVAddress
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Inactive Mailbox project $($inactiveMailbox)"
    
    if (!($migmailboxCheck = $inactiveMailboxMigrations | ? {$_.DestinationEmailAddress -eq $inactiveMailbox})) {
        $noMigProject2 += $mailbox
    }
}
$noMigProject2 | Export-Csv -NoTypeInformation -encoding utf8

#check if main inactive mailboxes in migration 3
$noMigProject = import-csv
$inactiveMailboxMigrations = import-csv
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
$noMigProject3 = @()
foreach ($mailbox in $InactiveMailboxes) {
    $inactiveMailbox = $mailbox.TJUVAddress
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Inactive Mailbox project $($inactiveMailbox)"
    
    if (!($migmailboxCheck = $inactiveMailboxMigrations | ? {$_.DestinationEmailAddress -eq $inactiveMailbox})) {
        $noMigProject3 += $mailbox
    }
}

#check inactive mailbox in tenants - EHN
$inactivemailboxes = import-csv
$progressref = ($inactivemailboxes).count
$progresscounter = 0
foreach ($mailbox in $inactivemailboxes) {
    $inactiveDisplayName = $mailbox.Name + "-Inactive"
    $inactiveDisplayName2 = $mailbox.Name + "-Inactive2"
    $newInactiveEHNAddress = ($mailbox.MicrosoftOnlineServicesID -split "@")[0] + "-inactive@EHN.MAIL.ONMICROSOFT.COM"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Inactive Mailbox project $($inactiveDisplayName)"
    
    if ($mailboxCheck = Get-Mailbox $mailbox.ExchangeGuid -ErrorAction SilentlyContinue) {
        #Write-Host "." -ForegroundColor green -NoNewline
        $mailbox | add-member -type noteproperty -name "Found" -Value "ExchangeGuid" -force
        $mailbox | add-member -type noteproperty -name "EHNDisplayName3" -Value $mailboxCheck.DisplayName -force
        $mailbox | add-member -type noteproperty -name "EHNAddress3" -Value $mailboxCheck.primarysmtpaddress -force
        $mailbox | add-member -type noteproperty -name "RestoreMethod3" -Value "MailboxRecover" -force
    }
    
    elseif ($mailboxCheck = Get-Mailbox $newInactiveEHNAddress -ErrorAction SilentlyContinue) {
        #Write-Host "." -ForegroundColor yellow -NoNewline
        $mailbox | add-member -type noteproperty -name "Found" -Value "NewInactiveMailboxAddress" -force
        $mailbox | add-member -type noteproperty -name "EHNDisplayName3" -Value $mailboxCheck.DisplayName -force
        $mailbox | add-member -type noteproperty -name "EHNAddress3" -Value $mailboxCheck.primarysmtpaddress -force
        if ($mbxRestoreRequestCheck = Get-MailboxRestoreRequest -BatchName "InactiveMailboxes" -TargetMailbox $newInactiveEHNAddress) {
            $mailbox | add-member -type noteproperty -name "RestoreMethod3" -Value "MailboxRestoreRequest" -force
        }
        else {
            $mailbox | add-member -type noteproperty -name "RestoreMethod3" -Value "NotFound" -force
        }
    }
    
    elseif ($mailboxCheck = Get-Mailbox $inactiveDisplayName2 -ErrorAction SilentlyContinue) {
        #Write-Host "." -ForegroundColor DarkCyan -NoNewline
        $mailbox | add-member -type noteproperty -name "Found" -Value "NewInactiveMailboxName2" -force
        $mailbox | add-member -type noteproperty -name "EHNDisplayName3" -Value $mailboxCheck.DisplayName -force
        $mailbox | add-member -type noteproperty -name "EHNAddress3" -Value $mailboxCheck.primarysmtpaddress -force
        if ($mbxRestoreRequestCheck = Get-MailboxRestoreRequest -BatchName "InactiveMailboxes" -TargetMailbox $inactiveDisplayName2) {
            $mailbox | add-member -type noteproperty -name "RestoreMethod3" -Value "MailboxRestoreRequest_Name2" -force
        }
        else {
            $mailbox | add-member -type noteproperty -name "RestoreMethod3" -Value "NotFound" -force
        }
    }

    elseif ($mailboxCheck = Get-Mailbox $inactiveDisplayName -ErrorAction SilentlyContinue) {
        #Write-Host "." -ForegroundColor cyan -NoNewline
        $mailbox | add-member -type noteproperty -name "Found" -Value "NewInactiveMailboxName" -force
        $mailbox | add-member -type noteproperty -name "EHNDisplayName3" -Value $mailboxCheck.DisplayName -force
        $mailbox | add-member -type noteproperty -name "EHNAddress3" -Value $mailboxCheck.primarysmtpaddress -force
        if ($mbxRestoreRequestCheck = Get-MailboxRestoreRequest -BatchName "InactiveMailboxes" -TargetMailbox $inactiveDisplayName) {
            $mailbox | add-member -type noteproperty -name "RestoreMethod3" -Value "MailboxRestoreRequest_Name" -force
        }
        else {
            $mailbox | add-member -type noteproperty -name "RestoreMethod3" -Value "NotFound" -force
        }
    }
    else {
        #Write-Host "." -ForegroundColor red -NoNewline
        $mailbox | add-member -type noteproperty -name "Found" -Value $False -force
        $mailbox | add-member -type noteproperty -name "EHNDisplayName3" -Value $null -force
        $mailbox | add-member -type noteproperty -name "EHNAddress3" -Value $null -force
        $mailbox | add-member -type noteproperty -name "RestoreMethod3" -Value "NotFound" -force
    }
}

#check inactive mailbox in tenants - TJUV
$inactivemailboxes = import-csv
$progressref = ($inactivemailboxes).count
$progresscounter = 0
foreach ($mailbox in $inactivemailboxes) {
    $inactiveDisplayName = $mailbox.Name + "-Inactive"
    $inactiveTJUAddress = ($mailbox.EHNAddress2 -split "@")[0] + "@tjuv.mail.onmicrosoft.com"
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Inactive Mailbox project $($inactiveDisplayName)"
    
    if ($mailboxCheck = Get-Mailbox $inactiveDisplayName -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor green -NoNewline
        $mailbox | add-member -type noteproperty -name "TJUVDisplayName2" -Value $mailboxCheck.DisplayName -force
        $mailbox | add-member -type noteproperty -name "TJUVAddress2" -Value $mailboxCheck.primarysmtpaddress -force
    }
    elseif ($mailboxCheck = Get-Mailbox $inactiveTJUAddress -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor green -NoNewline
        $mailbox | add-member -type noteproperty -name "TJUVDisplayName2" -Value $mailboxCheck.DisplayName -force
        $mailbox | add-member -type noteproperty -name "TJUVAddress2" -Value $mailboxCheck.primarysmtpaddress -force
    }
    else {
        Write-Host "." -ForegroundColor red -NoNewline
        $mailbox | add-member -type noteproperty -name "TJUVDisplayName2" -Value $null -force
        $mailbox | add-member -type noteproperty -name "TJUVAddress2" -Value $null -force
    }
}

#check if nonuser archive mailboxes in migration
$nonUserArchiveMailboxes = import-csv
$nonUserArchiveMigrations = import-csv
$progressref = ($nonUserArchiveMailboxes).count
$progresscounter = 0
$noMigProject = @()
foreach ($mailbox in $nonUserArchiveMailboxes) {
    $archiveMailbox = $mailbox.PrimarySMTPAddress_Destination
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Archive Mailbox project $($archiveMailbox)"
    
    if (!($migmailboxCheck = $nonUserArchiveMigrations | ? {$_.DestinationEmailAddress -eq $archiveMailbox})) {
        $noMigProject += $mailbox
    }
}
$noMigProject | Export-Csv -NoTypeInformation -encoding utf8

# Get MailboxStats for InActive Mailboxes - EHN
$InactiveMailboxes = Import-Csv "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxes\InactiveMailboxes.csv"
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
foreach ($mailbox in $InactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking MailboxStats for Inactive Mailbox $($mailbox.MicrosoftOnlineServicesID)"
    $newInactiveEHNAddress = ($mailbox.MicrosoftOnlineServicesID -split "@")[0] + "-inactive@EHN.MAIL.ONMICROSOFT.COM"

    #Get MailboxWhenDeleted - Exchange GUID
    if ($MailboxDetails = Get-Mailbox $mailbox.ExchangeGuid -IncludeInactiveMailbox -ErrorAction SilentlyContinue | select DisplayName,PrimarySMTPAddress,Identity, WhenSoftDeleted, WasInactiveMailbox, IsInactiveMailbox,WhenCreated) {
        $mailbox | add-member -type noteproperty -name "DisplayName_EHN2" -Value $MailboxDetails.DisplayName -force
        $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_EHN2" -Value $MailboxDetails.PrimarySMTPAddress -force
        $mailbox | add-member -type noteproperty -name "Identity_EHN" -Value $MailboxDetails.Identity -force
    }
    #Get MailboxWhenDeleted - newInactiveEHNAddress
    elseif ($MailboxDetails = Get-Mailbox $newInactiveEHNAddress -IncludeInactiveMailbox -ErrorAction SilentlyContinue | select DisplayName,PrimarySMTPAddress,Identity, WhenSoftDeleted, WasInactiveMailbox, IsInactiveMailbox,WhenCreated) {
        $mailbox | add-member -type noteproperty -name "DisplayName_EHN2" -Value $MailboxDetails.DisplayName -force
        $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_EHN2" -Value $MailboxDetails.PrimarySMTPAddress -force
        $mailbox | add-member -type noteproperty -name "Identity_EHN" -Value $MailboxDetails.Identity -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "DisplayName_EHN2" -Value $null -force
        $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_EHN2" -Value $null -force
        $mailbox | add-member -type noteproperty -name "Identity_EHN" -Value $null -force
    }
}

# Get MailboxStats for InActive Mailboxes - EHN Full
$InactiveMailboxes = Import-Csv "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxes\InactiveMailboxes.csv"
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
foreach ($mailbox in $InactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking MailboxStats for Inactive Mailbox $($mailbox.EHNAddress)"

    #Get MailboxWhenDeleted - Exchange GUID
    if ($MailboxDetails = Get-Mailbox $mailbox.ExchangeGuid -ErrorAction SilentlyContinue | select DisplayName,PrimarySMTPAddress,Identity, WhenSoftDeleted, WasInactiveMailbox, IsInactiveMailbox,WhenCreated) {
        $inactiveMailboxStats = Get-MailboxStatistics $mailbox.ExchangeGuid -IncludeSoftDeletedRecipients | select TotalItemSize, ItemCount
        $mailbox | add-member -type noteproperty -name "WhenCreated_EHN" -Value $MailboxDetails.WhenCreated -force
        $mailbox | add-member -type noteproperty -name "WhenSoftDeleted_EHN" -Value $MailboxDetails.WhenSoftDeleted -force
        $mailbox | add-member -type noteproperty -name "WasInactiveMailbox_EHN" -Value $MailboxDetails.WasInactiveMailbox -force
        $mailbox | add-member -type noteproperty -name "IsInactiveMailbox_EHN" -Value $MailboxDetails.IsInactiveMailbox -force
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSize" -Value $inactiveMailboxStats.TotalItemSize -force
        $mailbox | add-member -type noteproperty -name "InActive_ItemCount" -Value $inactiveMailboxStats.ItemCount -force
    }
    #Get MailboxWhenDeleted - newInactiveEHNAddress
    elseif ($MailboxDetails = Get-Mailbox $mailbox.EHNAddress -ErrorAction SilentlyContinue | select DisplayName,PrimarySMTPAddress,Identity, WhenSoftDeleted, WasInactiveMailbox, IsInactiveMailbox,WhenCreated) {
        $inactiveMailboxStats = Get-MailboxStatistics $mailbox.MicrosoftOnlineServicesID -IncludeSoftDeletedRecipients | select TotalItemSize, ItemCount
        $mailbox | add-member -type noteproperty -name "WhenCreated_EHN" -Value $MailboxDetails.WhenCreated -force
        $mailbox | add-member -type noteproperty -name "WhenSoftDeleted_EHN" -Value $MailboxDetails.WhenSoftDeleted -force
        $mailbox | add-member -type noteproperty -name "WasInactiveMailbox_EHN" -Value $MailboxDetails.WasInactiveMailbox -force
        $mailbox | add-member -type noteproperty -name "IsInactiveMailbox_EHN" -Value $MailboxDetails.IsInactiveMailbox -force
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSize" -Value $inactiveMailboxStats.TotalItemSize -force
        $mailbox | add-member -type noteproperty -name "InActive_ItemCount" -Value $inactiveMailboxStats.ItemCount -force
    }
    else {
        $mailbox | add-member -type noteproperty -name "WhenCreated_EHN" -Value $MailboxDetails.WhenCreated -force
        $mailbox | add-member -type noteproperty -name "WhenSoftDeleted_EHN" -Value $null -force
        $mailbox | add-member -type noteproperty -name "WasInactiveMailbox_EHN" -Value $null -force
        $mailbox | add-member -type noteproperty -name "IsInactiveMailbox_EHN" -Value $null -force
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSize_EHN" -Value $null -force
        $mailbox | add-member -type noteproperty -name "InActive_ItemCount_EHN" -Value $null -force
    }
}

# Gather UnMigrated Account Post Migration Details
$PreviouslyUnMigratedAccounts = import-csv
$AllPostMigrationDetails = Import-Excel -WorksheetName "PostMigrationDetails-ALL-513" -Path
$PreviouslyUnMigratedDetails = @()
foreach ($object in $PreviouslyUnMigratedAccounts) {
    $SourceAddress = $object.PrimarySmtpAddressSource
    if ($matchedObject = $AllPostMigrationDetails | ?{$_."Primary Smtp Address Source" -eq $SourceAddress}) {
        $tmpObject = New-Object PSObject
        $tmpObject | add-member -Type noteproperty -Name "DisplayName_Source" -value $matchedObject."DisplayName_Source"
        $tmpObject | add-member -Type noteproperty -Name "UserPrincipalName_Source" -value $matchedObject."User Principal Name Source"
        $tmpObject | add-member -Type noteproperty -Name "CustomAttribute7" -value $matchedObject."Custom Attribute7 Source"
        $tmpObject | add-member -Type noteproperty -Name "PrimarySmtpAddress_Source" -value $matchedObject."Primary Smtp Address Source"
        $tmpObject | add-member -Type noteproperty -Name "ForwardingSMTPAddress_Source" -value $matchedObject."Forwarding Smtp Address source"
        $tmpObject | add-member -Type noteproperty -Name "MBXSize_Source" -value $matchedObject."MBX Size source"
        $tmpObject | add-member -Type noteproperty -Name "DisplayName_Destination" -value $matchedObject."Display Name Destination"
        $tmpObject | add-member -Type noteproperty -Name "UserPrincipalName_Destination" -value $matchedObject."UserPrincipalName_Destination"
        $tmpObject | add-member -Type noteproperty -Name "PrimarySMTPAddress_Destination" -value $matchedObject."Found Primary SMTP Address destination"
        $PreviouslyUnMigratedDetails += $tmpObject
    }
}

#Gather EXOMailbox Stats - Mailbox and Archive Only (Source). Sizes in MB
$WorksheetName = "EHN-TJUVInactive"
$ExcelFilePath = "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"
$inactiveMailboxes = Import-Excel -WorksheetName $WorksheetName -Path $ExcelFilePath
$Tenant = "Source"
$progressref = ($inactiveMailboxes).count
$progresscounter = 0
foreach ($user in $inactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.SourceEmailAddress)"
    
    $PrimarySMTPAddress = @()
    $TotalItemSize = @()
    $TotalDeletedItemSize = @()
    $CombinedItemSize = @()
    $ArchiveTotalItemSize = @()
    $ArchiveTotalDeletedItemSize = @()
    $ArchiveCombinedItemSize = @()

    $PrimarySMTPAddress = $user.SourceEmailAddress

    if ($mbxCheck = Get-EXOMailbox $PrimarySMTPAddress -PropertySets archive -ErrorAction SilentlyContinue) {
        #Pull MailboxStats and UserDetails
        $mbxStats = Get-EXOMailboxStatistics $PrimarySMTPAddress

        $TotalItemSize = ([math]::Round(($MBXStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
        $TotalDeletedItemSize = ([math]::Round(($MBXStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
        $CombinedItemSize = $TotalItemSize + $TotalDeletedItemSize

        #Create User Array
        $user | Add-Member -type NoteProperty -Name "Size_$($Tenant)-MB" -Value $TotalItemSize -force
        $user | Add-Member -type NoteProperty -Name "DeletedSize_$($Tenant)-MB" -Value $TotalDeletedItemSize -force
        $user | Add-Member -Type NoteProperty -name "ItemCount_$($Tenant)" -Value $MBXStats.ItemCount -force
        $user | Add-Member -type NoteProperty -Name "DeletedItemCount_$($Tenant)" -Value $MBXStats.DeletedItemCount -force
        $user | Add-Member -type NoteProperty -Name "CombinedSize_$($Tenant)-MB" -Value $CombinedItemSize -force
        $user | Add-Member -type NoteProperty -Name "CombinedItemCount_$($Tenant)" -Value ($MBXStats.DeletedItemCount + $MBXStats.ItemCount) -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $mbxCheck.ArchiveStatus -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveName_$($Tenant)" -Value $mbxCheck.ArchiveName.tostring() -force

        # Archive Mailbox Check
        if ($mbxCheck.ArchiveName) { 
            $ArchiveStats = Get-EXOMailboxStatistics $PrimarySMTPAddress -Archive -ErrorAction silentlycontinue   
            #Archive Counts
            $ArchiveTotalItemSize = ([math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
            $ArchiveTotalDeletedItemSize = ([math]::Round(($ArchiveStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
            $ArchiveCombinedItemSize = $ArchiveTotalItemSize + $ArchiveTotalDeletedItemSize

            $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-MB" -Value $ArchiveTotalItemSize -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedSize_$($Tenant)-MB" -Value $ArchiveTotalDeletedItemSize -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $ArchiveStats.ItemCount -force 
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $ArchiveStats.DeletedItemCount -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedSize_$($Tenant)-MB" -Value $ArchiveCombinedItemSize -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value ($ArchiveStats.DeletedItemCount + $ArchiveStats.ItemCount) -force
        }
        else {
            $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-MB" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveDeletedSize_$($Tenant)-MB" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveCombinedSize_$($Tenant)-MB" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value $null -force
        }
    }
    else {
        Write-Host "." -foregroundcolor red -nonewline
        $user | Add-Member -type NoteProperty -Name "Size_$($Tenant)-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "DeletedTotalSize_$($Tenant)-MB" -Value $null -force
        $user | Add-Member -Type NoteProperty -name "ItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "DeletedItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "CombinedSize_$($Tenant)-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "CombinedItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveName_$($Tenant)" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-MB" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveDeletedSize_$($Tenant)-MB" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveCombinedSize_$($Tenant)-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value $null -force
    }
}
$inactiveMailboxes | Export-Excel -WorksheetName $WorksheetName -Path $ExcelFilePath


#Gather EXOMailbox Stats - Mailbox and Archive Only (Destination). Sizes in MB
$WorksheetName = "EHN-TJUVInactive"
$ExcelFilePath = "C:\Users\aaron.medrano\Desktop\post migration\EHN-TJU-WaveMailboxMigrations.xlsx"
$inactiveMailboxes = Import-Excel -WorksheetName $WorksheetName -Path $ExcelFilePath
$Tenant = "Destination"
$progressref = ($inactiveMailboxes).count
$progresscounter = 0
foreach ($user in $inactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DestinationEmailAddress)"
    
    $PrimarySMTPAddress = @()
    $TotalItemSize = @()
    $TotalDeletedItemSize = @()
    $CombinedItemSize = @()
    $ArchiveTotalItemSize = @()
    $ArchiveTotalDeletedItemSize = @()
    $ArchiveCombinedItemSize = @()

    $PrimarySMTPAddress = $user.DestinationEmailAddress

    if ($mbxCheck = Get-EXOMailbox $PrimarySMTPAddress -PropertySets archive -ErrorAction SilentlyContinue) {
        #Pull MailboxStats and UserDetails
        $mbxStats = Get-EXOMailboxStatistics $PrimarySMTPAddress

        $TotalItemSize = ([math]::Round(($MBXStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
        $TotalDeletedItemSize = ([math]::Round(($MBXStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
        $CombinedItemSize = $TotalItemSize + $TotalDeletedItemSize

        #Create User Array
        $user | Add-Member -type NoteProperty -Name "Size_$($Tenant)-MB" -Value $TotalItemSize -force
        $user | Add-Member -type NoteProperty -Name "DeletedSize_$($Tenant)-MB" -Value $TotalDeletedItemSize -force
        $user | Add-Member -Type NoteProperty -name "ItemCount_$($Tenant)" -Value $MBXStats.ItemCount -force
        $user | Add-Member -type NoteProperty -Name "DeletedItemCount_$($Tenant)" -Value $MBXStats.DeletedItemCount -force
        $user | Add-Member -type NoteProperty -Name "CombinedSize_$($Tenant)-MB" -Value $CombinedItemSize -force
        $user | Add-Member -type NoteProperty -Name "CombinedItemCount_$($Tenant)" -Value ($MBXStats.DeletedItemCount + $MBXStats.ItemCount) -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $mbxCheck.ArchiveStatus -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveName_$($Tenant)" -Value $mbxCheck.ArchiveName.tostring() -force

        # Archive Mailbox Check
        if ($mbxCheck.ArchiveName) { 
            $ArchiveStats = Get-EXOMailboxStatistics $PrimarySMTPAddress -Archive -ErrorAction silentlycontinue 
            #Archive Counts
            $ArchiveTotalItemSize = ([math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
            $ArchiveTotalDeletedItemSize = ([math]::Round(($ArchiveStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/1MB,3))
            $ArchiveCombinedItemSize = $ArchiveTotalItemSize + $ArchiveTotalDeletedItemSize

            $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-MB" -Value $ArchiveTotalItemSize -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedSize_$($Tenant)-MB" -Value $ArchiveTotalDeletedItemSize -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $ArchiveStats.ItemCount -force 
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $ArchiveStats.DeletedItemCount -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedSize_$($Tenant)-MB" -Value $ArchiveCombinedItemSize -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value ($ArchiveStats.DeletedItemCount + $ArchiveStats.ItemCount) -force
        }
        else {
            $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-MB" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveDeletedSize_$($Tenant)-MB" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveCombinedSize_$($Tenant)-MB" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value $null -force
        }
    }
    else {
        Write-Host "." -foregroundcolor red -nonewline
        $user | Add-Member -type NoteProperty -Name "Size_$($Tenant)-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "DeletedTotalSize_$($Tenant)-MB" -Value $null -force
        $user | Add-Member -Type NoteProperty -name "ItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "DeletedItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "CombinedSize_$($Tenant)-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "CombinedItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $null -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveName_$($Tenant)" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-MB" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveDeletedSize_$($Tenant)-MB" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $null -force
        $user | add-member -type noteproperty -name "ArchiveCombinedSize_$($Tenant)-MB" -Value $null -force
        $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value $null -force
    }
}
$inactiveMailboxes | Export-Excel -WorksheetName $WorksheetName -Path $ExcelFilePath


#Gather EXOMailbox Stats - Mailbox and Archive Only (BOTH)
function Get-MigrationMailboxStatistics {
    param (
        [Parameter(ParameterSetName='ExcelImport',Position=1,Mandatory=$false,HelpMessage="Specify if using Excel Workbookt to Import list of Users?")] [switch] $ExcelImport,
        [Parameter(ParameterSetName='CSVImport',Position=2,Mandatory=$false,HelpMessage="Specify if using CSV file to Import list of Users??")] [string] $CSVImport,
        [Parameter(ParameterSetName='ExcelImport',Position=0,Mandatory=$false,HelpMessage="Migrate?")] [string] $WorksheetName,
        [Parameter(Mandatory=$True,HelpMessage="What Is the File Path to List of users")] [string] $ImportFilePath,
        [Parameter(Mandatory=$True,HelpMessage="Specify if pulling Source or Destination (target) Tenant Details?")] [string] $Tenant,
        [Parameter(Mandatory=$True,HelpMessage="What format do you wish to Export the Mailbox Sizes? Examples: MB for Megabyte and GB for Gigabyte.")][string] $SizeFormat
    )
    #Import List of Users
    try {
        if ($ExcelImport) {
            $migrationMailboxes = Import-Excel -WorksheetName $WorksheetName -Path $ImportFilePath
        }
        elseif ($CSVImport) {
            $migrationMailboxes = Import-CSV -Path $ImportFilePath
        }
        else {
            Write-Error "Missing REQUIRED Parameter - Import. Please run again and specify ExcelImport or CSVImport to Import list of Users"
            Return
        }
    }
    catch {
        Write-Error $_.Exception_Message
        Return
    }

    #Size Format
    if ($SizeFormat -eq "MB" -or $SizeFormat -eq "GB") {
    }
    else {
        Write-Error "Missing REQUIRED Paramater - SizeFormat. Unrecognized parameter input provided.  Please Run again. Examples: MB for Megabyte and GB for Gigabyte."
        Return
    }
    $formulaSizeFormat = "1" + $SizeFormat
    
    $progressref = ($migrationMailboxes).count
    $progresscounter = 0
    foreach ($user in $migrationMailboxes) {
        # Clear Previous Variable and Set Variables
        ## Primary Address
        $PrimarySMTPAddress = $null
        if ($Tenant -eq "Source") {
            $PrimarySMTPAddress = $user.SourceEmailAddress
        }
        elseif ($Tenant -eq "Destination") {
            $PrimarySMTPAddress = $user.DestinationEmailAddress
        }
        #Item Statistics
        $TotalItemSize = @()
        $TotalDeletedItemSize = @()
        $CombinedItemSize = @()
        $ArchiveTotalItemSize = @()
        $ArchiveTotalDeletedItemSize = @()
        $ArchiveCombinedItemSize = @()
        
        #Progress Bar
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($PrimarySMTPAddress)"

        if ($mbxCheck = Get-EXOMailbox $PrimarySMTPAddress -PropertySets archive -ErrorAction SilentlyContinue) {
            #Pull MailboxStats and UserDetails
            $mbxStats = Get-EXOMailboxStatistics $PrimarySMTPAddress

            $TotalItemSize = ([math]::Round(($MBXStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/$formulaSizeFormat,3))
            $TotalDeletedItemSize = ([math]::Round(($MBXStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/$formulaSizeFormat,3))
            $CombinedItemSize = $TotalItemSize + $TotalDeletedItemSize

            #Create User Array
            $user | Add-Member -type NoteProperty -Name "Size_$($Tenant)-$($SizeFormat)" -Value $TotalItemSize -force
            $user | Add-Member -type NoteProperty -Name "DeletedSize_$($Tenant)-$($SizeFormat)" -Value $TotalDeletedItemSize -force
            $user | Add-Member -Type NoteProperty -name "ItemCount_$($Tenant)" -Value $MBXStats.ItemCount -force
            $user | Add-Member -type NoteProperty -Name "DeletedItemCount_$($Tenant)" -Value $MBXStats.DeletedItemCount -force
            $user | Add-Member -type NoteProperty -Name "CombinedSize_$($Tenant)-$($SizeFormat)" -Value $CombinedItemSize -force
            $user | Add-Member -type NoteProperty -Name "CombinedItemCount_$($Tenant)" -Value ($MBXStats.DeletedItemCount + $MBXStats.ItemCount) -force
            $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $mbxCheck.ArchiveStatus -force
            $user | Add-Member -Type NoteProperty -Name "ArchiveName_$($Tenant)" -Value $mbxCheck.ArchiveName.tostring() -force

            # Archive Mailbox Check
            if ($mbxCheck.ArchiveName) {    
                $ArchiveStats = Get-EXOMailboxStatistics $PrimarySMTPAddress -Archive -ErrorAction silentlycontinue
                #Archive Counts
                $ArchiveTotalItemSize = ([math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/$formulaSizeFormat,3))
                $ArchiveTotalDeletedItemSize = ([math]::Round(($ArchiveStats.TotalDeletedItemSize.ToString() -replace “(.*\()|,| [a-z]*\)”, “”)/$formulaSizeFormat,3))
                $ArchiveCombinedItemSize = $ArchiveTotalItemSize + $ArchiveTotalDeletedItemSize

                $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-$($SizeFormat)" -Value $ArchiveTotalItemSize -force
                $user | Add-Member -type NoteProperty -Name "ArchiveDeletedSize_$($Tenant)-$($SizeFormat)" -Value $ArchiveTotalDeletedItemSize -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $ArchiveStats.ItemCount -force 
                $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $ArchiveStats.DeletedItemCount -force
                $user | Add-Member -type NoteProperty -Name "ArchiveCombinedSize_$($Tenant)-$($SizeFormat)" -Value $ArchiveCombinedItemSize -force
                $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value ($ArchiveStats.DeletedItemCount + $ArchiveStats.ItemCount) -force
            }
            else {
                $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-$($SizeFormat)" -Value $null -force
                $user | add-member -type noteproperty -name "ArchiveDeletedSize_$($Tenant)-$($SizeFormat)" -Value $null -force
                $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
                $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $null -force
                $user | add-member -type noteproperty -name "ArchiveCombinedSize_$($Tenant)-$($SizeFormat)" -Value $null -force
                $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value $null -force
            }
        }
        else {
            Write-Host "." -foregroundcolor red -nonewline
            $user | Add-Member -type NoteProperty -Name "Size_$($Tenant)-$($SizeFormat)" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "DeletedTotalSize_$($Tenant)-$($SizeFormat)" -Value $null -force
            $user | Add-Member -Type NoteProperty -name "ItemCount_$($Tenant)" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "DeletedItemCount_$($Tenant)" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "CombinedSize_$($Tenant)-$($SizeFormat)" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "CombinedItemCount_$($Tenant)" -Value $null -force
            $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($Tenant)" -Value $null -force
            $user | Add-Member -Type NoteProperty -Name "ArchiveName_$($Tenant)" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveSize_$($Tenant)-$($SizeFormat)" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveDeletedSize_$($Tenant)-$($SizeFormat)" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveItemCount_$($Tenant)" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveDeletedItemCount_$($Tenant)" -Value $null -force
            $user | add-member -type noteproperty -name "ArchiveCombinedSize_$($Tenant)-$($SizeFormat)" -Value $null -force
            $user | Add-Member -type NoteProperty -Name "ArchiveCombinedItemCount_$($Tenant)" -Value $null -force
        }
    }

    #Export Results
    if ($ExcelImport) {
        try {
            $migrationMailboxes | Export-Excel -WorksheetName $WorksheetName -Path $ImportFilePath
            Write-host "Exported Migration Mailbox Statistics to $ExcelFilePath" -ForegroundColor Cyan
        }
        catch {
            Write-Warning -Message "$($_.Exception)"
            Write-host ""
            $OutputExcelFilePath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full File Path with .xlsx file extension'
            $OutputExcelWorkSheetName = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide Worksheet Name'
            $migrationMailboxes | Export-Excel -WorksheetName $OutputExcelWorkSheetName -Path $OutputExcelFilePath
            Write-host "Exported Migration Mailbox Statistics to $OutputExcelFilePath" -ForegroundColor Cyan
        }
    }
    elseif ($CSVImport) {
        try {
            $migrationMailboxes | Export-Csv $ImportFilePath -NoTypeInformation -Encoding UTF8
            Write-host "Exported Migration Mailbox Statistics to $ImportFilePath" -ForegroundColor Cyan
        }
        catch {
            Write-Warning -Message "$($_.Exception)"
            Write-host ""
            $OutputCSVFilePath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path with .csv file extension'
            $migrationMailboxes | Export-Csv $OutputCSVFilePath -NoTypeInformation -Encoding UTF8
            Write-host "Exported Migration Mailbox Statistics to $OutputCSVFilePath" -ForegroundColor Cyan
        }
    }
}

#Gather EXOMailbox Stats - Mailbox and Archive Only (TJUV)
$InactiveMailboxes = Import-Excel -WorksheetName "Inactive Details2" -Path "C:\Users\aaron.medrano\Desktop\post migration\InactiveMailboxes\EHN-TJU Inactive Mailbox Migrations.xlsx"
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
foreach ($user in $InactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Removing Mailbox $($user.DisplayName)"
    if ($msolUserCheck = Get-MsolUser -UserPrincipalName $user.UserPrincipalName -ErrorAction SilentlyContinue) {
        try {
            Remove-MsolUser -UserPrincipalName $user.UserPrincipalName -Force -ErrorAction Stop
            Write-Host "." -NoNewline -ForegroundColor Green
        }
        catch {
            Start-Sleep -Seconds 2
            Remove-MsolUser -UserPrincipalName $user.UserPrincipalName -Force
        }
    }    
}
