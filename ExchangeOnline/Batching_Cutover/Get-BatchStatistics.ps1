#CSV File Check

function Get-BatchStatistics {
    param (
        [Parameter(Mandatory=$True)][array]$batch
    )
    #Create Arrays
    $fullMigStats = @()

    #ProgressBar1
    $progressref = ($batch).count
    $progresscounter = 0
        
    #Grab MigrationStats
    foreach ($user in $batch)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Pulling Mailbox Migration Details for $($user.mailbox)"

        #Get MigrationUser Stats
        if ($recipientCheck = get-recipient $user.mailbox -ea silentlycontinue)
        {
            if ($migcheck = Get-MigrationUser $recipientCheck.primarysmtpaddress -ea silentlycontinue | select identity,BatchId,LastSuccessfulSyncTime,HasUnapprovedSkippedItems,ErrorSummary) 
            {
                $moveStatistics = Get-MoveRequestStatistics $recipientCheck.primarysmtpaddress | select statusdetail,percentcomplete,CompletionTimestamp,CompleteAfter,totalmailboxsize,SkippedItemApprovalTimestamp
            }
            else
            {
                Write-Host "No migration user found for $($user.mailbox) ... " -ForegroundColor red -NoNewline
            }
        }
        else
        {
            Write-Host "No Recipient found for $($user.mailbox)" -ForegroundColor red
        }

        #Create Output Array
        $tmpMigStats = New-Object psobject
        $tmpMigStats | Add-Member -NotePropertyName "MigrationUser" -NotePropertyValue $migcheck.identity
        $tmpMigStats | Add-Member -NotePropertyName "Batch" -NotePropertyValue $migcheck.BatchId
        $tmpMigStats | Add-Member -NotePropertyName "StatusDetail" -NotePropertyValue $moveStatistics.statusdetail
        $tmpMigStats | Add-Member -NotePropertyName "LastSuccessfulSyncTime" -NotePropertyValue $migcheck.LastSuccessfulSyncTime
        $tmpMigStats | Add-Member -NotePropertyName "PercentComplete" -NotePropertyValue $moveStatistics.percentcomplete
        $tmpMigStats | Add-Member -NotePropertyName "CompletionTimestamp" -NotePropertyValue $moveStatistics.CompletionTimestamp
        $tmpMigStats | Add-Member -NotePropertyName "CompleteAfter" -NotePropertyValue $moveStatistics.CompleteAfter
        $tmpMigStats | Add-Member -NotePropertyName "TotalMailboxSize" -NotePropertyValue $moveStatistics.totalmailboxsize
        $tmpMigStats | Add-Member -NotePropertyName "HasUnapprovedSkippedItems" -NotePropertyValue $migcheck.HasUnapprovedSkippedItems
        $tmpMigStats | Add-Member -NotePropertyName "SkippedItemApprovalTimestamp" -NotePropertyValue $moveStatistics.SkippedItemApprovalTimestamp
        $tmpMigStats | Add-Member -NotePropertyName "DataConsistencyScore" -NotePropertyValue $moveStatistics.DataConsistencyScore
        $tmpMigStats | Add-Member -NotePropertyName "MoveMessageDetails" -NotePropertyValue $moveStatistics.Message
        $tmpMigStats | Add-Member -NotePropertyName "ErrorSummary" -NotePropertyValue $migcheck.ErrorSummary
        

        $fullMigStats += $tmpMigStats
    }

    #Output
    $fullMigStats
}

#TXT File check
function Get-BatchStatistics {
    param (
        [Parameter(Mandatory=$True)][array]$batch
    )
    #Create Arrays
    $fullMigStats = @()

    #ProgressBar1
    $progressref = ($batch).count
    $progresscounter = 0
        
    #Grab MigrationStats
    foreach ($user in $batch)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Pulling Mailbox Migration Details for $($user)"

        #Get MigrationUser Stats
        if ($recipientCheck = get-recipient $user -ea silentlycontinue)
        {
            if ($migcheck = Get-MigrationUser $recipientCheck.primarysmtpaddress -ea silentlycontinue | select identity,BatchId,LastSuccessfulSyncTime,HasUnapprovedSkippedItems,ErrorSummary) 
            {
                $moveStatistics = Get-MoveRequestStatistics $recipientCheck.primarysmtpaddress | select statusdetail,percentcomplete,CompletionTimestamp,CompleteAfter,totalmailboxsize,SkippedItemApprovalTimestamp
            }
            else
            {
                Write-Host "No migration user found for $($user) ... " -ForegroundColor red -NoNewline
            }
        }
        else
        {
            Write-Host "No Recipient found for $($user)" -ForegroundColor red
        }

        #Create Output Array
        $tmpMigStats = New-Object psobject
        $tmpMigStats | Add-Member -NotePropertyName "MigrationUser" -NotePropertyValue $migcheck.identity
        $tmpMigStats | Add-Member -NotePropertyName "Batch" -NotePropertyValue $migcheck.BatchId
        $tmpMigStats | Add-Member -NotePropertyName "StatusDetail" -NotePropertyValue $moveStatistics.statusdetail
        $tmpMigStats | Add-Member -NotePropertyName "LastSuccessfulSyncTime" -NotePropertyValue $migcheck.LastSuccessfulSyncTime
        $tmpMigStats | Add-Member -NotePropertyName "PercentComplete" -NotePropertyValue $moveStatistics.percentcomplete
        $tmpMigStats | Add-Member -NotePropertyName "CompletionTimestamp" -NotePropertyValue $moveStatistics.CompletionTimestamp
        $tmpMigStats | Add-Member -NotePropertyName "CompleteAfter" -NotePropertyValue $moveStatistics.CompleteAfter
        $tmpMigStats | Add-Member -NotePropertyName "TotalMailboxSize" -NotePropertyValue $moveStatistics.totalmailboxsize
        $tmpMigStats | Add-Member -NotePropertyName "HasUnapprovedSkippedItems" -NotePropertyValue $migcheck.HasUnapprovedSkippedItems
        $tmpMigStats | Add-Member -NotePropertyName "SkippedItemApprovalTimestamp" -NotePropertyValue $moveStatistics.SkippedItemApprovalTimestamp
        $tmpMigStats | Add-Member -NotePropertyName "DataConsistencyScore" -NotePropertyValue $moveStatistics.DataConsistencyScore
        $tmpMigStats | Add-Member -NotePropertyName "MoveMessageDetails" -NotePropertyValue $moveStatistics.Message
        $tmpMigStats | Add-Member -NotePropertyName "ErrorSummary" -NotePropertyValue $migcheck.ErrorSummary
        

        $fullMigStats += $tmpMigStats
    }

    #Output
    $fullMigStats
}