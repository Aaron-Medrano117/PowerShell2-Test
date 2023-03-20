function Set-BatchCutover {
    param (
        $CompleteAfter,
        [array]$batch,
        [switch]$BusinessStandard,
        [switch]$E3
    )
    if ($CompleteAfter -eq 1)
    {
        $cutovertime = "1"
    }
    else
    {
        $cutovertime = (get-date $CompleteAfter).ToUniversalTime()
    }

    #ProgressBar1
    $progressref = ($batch).count
    $progresscounter = 0

    foreach ($user in $batch)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Setting $($user.mailbox) cutover for $($completeafter)."

        if ($recipientCheck = get-Recipient $user.mailbox -ea silentlycontinue  | ? {$_.RecipientTypeDetails -eq "MailUser"})
        {
            if (Get-MigrationUser $recipientCheck.primarysmtpaddress -ea silentlycontinue)
            {
                Set-MoveRequest $User.mailbox -SkippedItemApprovalTime $cutovertime -ea silentlycontinue -wa silentlycontinue
                Write-Host "$($recipientCheck.primarysmtpaddress) cutover set for $($completeafter) ... " -ForegroundColor Cyan -NoNewline
                Set-MoveRequest $User.mailbox -completeafter $cutovertime -ea silentlycontinue -wa silentlycontinue
            }
            else
            {
                Write-Host "No migration user found for $($user.mailbox) ... " -ForegroundColor red -NoNewline
            }

            if ($BusinessStandard)
            {
                $msolusercheck = Get-MsolUser -SearchString $recipientCheck.name
                if (!($msolusercheck.licenses.accountskuid -contains "filethrutrial:O365_BUSINESS_PREMIUM"))
                {
                    Write-host "Updating to Business Standard license .. "
                    Set-MsolUserLicense -UserPrincipalName $msolusercheck.UserPrincipalName -AddLicenses "filethrutrial:O365_BUSINESS_PREMIUM"
                }
                else
                {
                    Write-host "Business Standard license already on user .. " -ForegroundColor Yellow
                }
            }
            if ($E3)
            {
                $msolusercheck = Get-MsolUser -SearchString $recipientCheck.name
                if (!($msolusercheck.licenses.accountskuid -contains "filethrutrial:ENTERPRISEPACK"))
                {
                    Write-host "Updating to E3 license .. " -NoNewline
                    Set-MsolUserLicense -UserPrincipalName $msolusercheck.UserPrincipalName -AddLicenses "filethrutrial:ENTERPRISEPACK"
                }
                else
                {
                    Write-host "E3 license already on user .. " -ForegroundColor Yellow
                }
            }
        }
        else
        {
            Write-Host "No Recipient Found for $($user.mailbox)" -ForegroundColor red
        }
        Write-Host "done" -ForegroundColor Green     
    }
}

## TXT File Import
function Set-BatchCutover {
    param (
        $CompleteAfter,
        [array]$batch,
        [switch]$BusinessStandard,
        [switch]$E3
    )
    if ($CompleteAfter -eq 1)
    {
        $cutovertime = "1"
    }
    else
    {
        $cutovertime = (get-date $CompleteAfter).ToUniversalTime()
    }

    #ProgressBar1
    $progressref = ($batch).count
    $progresscounter = 0

    foreach ($user in $batch)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Setting $($user) cutover for $($completeafter)."

        if ($recipientCheck = get-Recipient $user -ea silentlycontinue)
        {
            if (Get-MigrationUser $recipientCheck.primarysmtpaddress -ea silentlycontinue)
            {
                Set-MoveRequest $user -SkippedItemApprovalTime $cutovertime -ea silentlycontinue -wa silentlycontinue
                Write-Host "$($recipientCheck.primarysmtpaddress) cutover set for $($completeafter) ... " -ForegroundColor Cyan -NoNewline
                Set-MoveRequest $User -completeafter $cutovertime -ea silentlycontinue -wa silentlycontinue
            }
            else
            {
                Write-Host "No migration user found for $($user) ... " -ForegroundColor red -NoNewline
            }
        }
        else
        {
            Write-Host "No Recipient Found for $($user)" -ForegroundColor red
        }
        Write-Host "done" -ForegroundColor Green     
    }
}