Start-Transcript
Write-Host "Set-Forward version 1.3" -ForegroundColor Cyan
Write-Host "Script Details: Configure Forwarding from Source mailbox to Target mailbox based on CSV/Excel Input" -ForegroundColor Cyan
Write-Host "Version updated 3/20/2023" -ForegroundColor Cyan
Write-Host "Version author - Aaron Medrano" -ForegroundColor Cyan
Write-Host "Written by Aaron Medrano, Andrew Cronic, and John Williams - Arraya Solutions Modern Workplace Solutions Engineering Team" -ForegroundColor Cyan


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

$ImportFileLocation = Read-Host "What is the pathfile of the list of users? Supports CSV files"
#"C:\Office365\Users.csv"
#Cutover Users
if ($ImportFileLocation -like "*.csv") {
    Write-Host "CSV File Found"
    $allMailboxes = Import-CSV -Path $ImportFileLocation
}
elseif ($ImportFileLocation -like "*.xlxs") {
    Write-Host "XLXS File Found"
    $allMailboxes = Import-Excel -Path $ImportFileLocation
}
else {
    Write-Host "Unsupported version uploaded. Please try running again using CSV or an XLXS file type" -ForegroundColor Red
}


#Check if Headers are in Input - Source and Target Headings required
if ($allMailboxes.Source -or $allmailboxes.Target) {
    #ProgressBar
    $progresscounter = 1
    $global:start = Get-Date
    [nullable[double]]$global:secondsRemaining = $null

    $allErrors = @()
    $completedUsers = @()
    #Array to store mailbox data post-forwarding-changes / audit to confirm forwarding changes were successful
    $auditResults = @()

    foreach ($user in $allMailboxes) {
        $SourcePrimarySMTPAddress = $user.Source
        $DestinationPrimarySMTPAddress = $user.Target

        #progress bar
        Write-ProgressHelper -Activity "Updating Forwarding for $($SourcePrimarySMTPAddress)" -ProgressCounter ($progresscounter++) -TotalCount ($allMailboxes).count

        Write-Host "Cutting Over User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
        ## Set Mailbox to Forward from Source to Destination Mailbox
        Write-Host "Set Forward to $($DestinationPrimarySMTPAddress)  " -foregroundcolor Magenta -nonewline
        Try{       
            Set-Mailbox $SourcePrimarySMTPAddress -ForwardingSmtpAddress $DestinationPrimarySMTPAddress -ErrorAction Stop
            Write-Host "Completed" -ForegroundColor Green
            $completedUsers += $user
        }
        Catch {
            Write-Host "Failed" -ForegroundColor red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SetForward" -Force
            $currenterror | Add-Member -type NoteProperty -Name "SourcePrimarySMTPAddress" -Value $SourcePrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "DesinationPrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $allErrors += $currenterror
            continue
        }
    }
}
else {
    Write-Host "Missing required Headers. Please check that Excel or CSV file has the following headers" -ForegroundColor Red
    Write-Host ""Source" - This is the Source Email Address of the mailbox to configure the forward" -ForegroundColor Yellow
    Write-Host ""Target" - This is the Target Email Address to forward the source mailbox to" -ForegroundColor Yellow
}


#Audit Mailbox Forwarding previously set

if ($allMailboxes.Source -or $allmailboxes.Target) {
    #reset progresscounter
    $progresscounter = 1
    $global:start = Get-Date
    [nullable[double]]$global:secondsRemaining = $null
    Write-Host "Auditing forwarding settings set by previous operations..."

    foreach ($user in $allMailboxes) {
        $SourcePrimarySMTPAddress = $user.Source
        $DestinationPrimarySMTPAddress = $user.Target
        $ProgressPreference = "Continue" #added

        #progress bar
        Write-ProgressHelper -Activity "Getting Forwarding Settings for $($SourcePrimarySMTPAddress)" -ProgressCounter ($progresscounter++) -TotalCount ($allMailboxes).count
        Write-Host "Auditing forwarding settings for User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -NoNewline
        Try{       
            $mailboxData = Get-EXOMailbox $SourcePrimarySMTPAddress -Properties ForwardingSmtpAddress,DeliverToMailboxAndForward,ForwardingAddress -ErrorAction Stop #updated to EXO-Mailbox
            $mailboxData | Add-Member -type NoteProperty -Name "ForwardingAddressMismatch" -Value "" -Force

            if($mailboxData.ForwardingSmtpAddress -replace '^SMTP:' -ne $DestinationPrimarySMTPAddress)
            {
                $mailboxData.forwardingAddressMismatch = "TRUE"
            }
            else {
                $mailboxData.forwardingAddressMismatch = "FALSE"
            }

            Write-Host "Completed" -ForegroundColor Green #added
            $auditResults += $mailboxData | select UserPrincipalName, PrimarySmtpAddress, *Forward*
        }
        Catch {
            Write-Host "Failed" -ForegroundColor red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "Get-Mailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "SourcePrimarySMTPAddress" -Value $SourcePrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "DesinationPrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $allErrors += $currenterror
            continue
        }
    }

    #AUDIT END

    Write-Host "Displaying Audit Results in console:"
    Write-Host "ForwardingAddressMismatch property TRUE indicates the destination SMTP address was not set correctly" -ForegroundColor Yellow
    $AuditResults | ?{$_.ForwardingAddressMismatch -eq $true}

    Write-Host ""
    Write-Host "Exporting Audit Results .csv to $pwd"
    $auditResults | Export-Csv ".\ForwardingAuditReport-$(get-date -f yyyy-MM-dd-hh-mm).csv" -NoTypeInformation
    

    Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan
    Write-Host $completedUsers.count 'Users Forwarding Completed' -ForegroundColor Cyan
    Write-Host $allErrors.count 'Users Generated Errors. Check the $allErrors variable for more details' -ForegroundColor Red

    Write-Host ""
    Write-Host ""
    Write-Host ""
}
Stop-Transcript