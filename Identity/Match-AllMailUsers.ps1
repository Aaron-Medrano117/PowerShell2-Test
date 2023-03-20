function Match-AllMailUsers {
    param (
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$true)] [array] $ImportCSV,
        [Parameter(Mandatory=$false)] [string] $NewDomain
    )
    $ImportedUsers = Import-Csv $ImportCSV
    $AllUsers = @()
    
    #ProgressBar
    $progressref = ($ImportedUsers).count
    $progresscounter = 0

    foreach ($mailbox in $ImportedUsers)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($mailbox.DisplayName)"
        
        Write-Host "Checking for $($mailbox.displayName) in Tenant ..." -fore Cyan -NoNewline
        $newAddressSplit = $mailbox.PrimarySmtpAddress -split "@"
        $newMailboxAddress = $newAddressSplit[0] + "@" + $NewDomain
        if ($mailboxcheck = Get-Mailbox $mailbox.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced)
        {
            Write-Host "found mailbox  " -ForegroundColor Green -nonewline
        }
        elseif ($mailboxcheck = Get-Mailbox $mailbox.displayName -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database)
        {
           Write-Host "found mailbox*  " -ForegroundColor Yellow -nonewline
        }
        elseif ($mailboxcheck = Get-Mailbox $newMailboxAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database)
        {
            Write-Host "found mailbox**  " -ForegroundColor Yellow -nonewline
        }
        elseif ($recipientcheck = Get-Recipient $mailbox.PrimarySmtpAddress -ea silentlycontinue)
        {
            $mailboxcheck = Get-Mailbox $recipientcheck.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database 
            Write-Host "found recipient  " -ForegroundColor Yellow -nonewline
        }
        else
        {
            Write-Host "not found" -ForegroundColor red -NoNewline
            $msoluserscheck = @()
            $MBXStats = @()
        }
        if ($mailboxcheck)
        {
            $msoluserscheck = get-msoluser -UserPrincipalName $mailboxcheck.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation
            $MBXStats = Get-MailboxStatistics $mailboxcheck.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
            $mailbox | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $True
            $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluserscheck.UserPrincipalName
            $mailbox | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluserscheck.IsLicensed
            $mailbox | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluserscheck.Licenses.AccountSkuID -join ",")
            $mailbox | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $mailboxcheck.IsDirSynced
            $mailbox | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $msoluserscheck.PreferredDataLocation
            $mailbox | add-member -type noteproperty -name "Database_Destination" -Value $mailboxcheck.Database
            $mailbox | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $msoluserscheck.BlockCredential
            $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $mailboxcheck.PrimarySmtpAddress
            $mailbox | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailboxcheck.RecipientTypeDetails   
            $mailbox | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize
            $mailbox | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount

        }
        else 
        {
            $mailbox | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False
            $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "IsLicensed_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "Licenses_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "Database_Destination" -Value $mailboxcheck.Database
            $mailbox | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null  
            $mailbox | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $null
            $mailbox | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $null
        }
        Write-host " .. done" -foregroundcolor green
        $AllUsers += $mailbox
    }
    $allUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath
}
##Match MSOLUsers
function Match-AllMsolUsers {
    param (
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$true)] [array] $ImportCSV,
        [Parameter(Mandatory=$false)] [string] $NewDomain
    )
    $ImportedUsers = Import-Csv $ImportCSV
    $AllUsers = @()
    
    #ProgressBar
    $progressref = ($ImportedUsers).count
    $progresscounter = 0

    foreach ($mailbox in $ImportedUsers)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($mailbox.DisplayName)"
        
        Write-Host "Checking for $($mailbox.displayName) in Tenant ..." -fore Cyan -NoNewline
        $newAddressSplit = $mailbox.PrimarySmtpAddress -split "@"
        $newMailboxAddress = $newAddressSplit[0] + "@" + $NewDomain
        if ($msolusercheck = Get-MsolUser -searchstring $mailbox.DisplayName -ea silentlycontinue | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation)
        {
            Write-Host "found MSOLUser. " -ForegroundColor Green -nonewline
        }
        elseif ($msolusercheck = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName -ea silentlycontinue  | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation)
        {
            Write-Host "found MSOLUser*. " -ForegroundColor Green -nonewline
        }
        else
        {
            Write-Host "not found" -ForegroundColor red -NoNewline
            $msolusercheck = @()
            $MBXStats = @()
        }
        if ($msolusercheck)
        {
            $mailboxCheck = $mailboxcheck = Get-Mailbox $msolUserCheck.UserPrincipalName -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, IsDirSynced
            $MBXStats = Get-MailboxStatistics $mailboxcheck.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
            $mailbox | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $True -force
            $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msolusercheck.UserPrincipalName -force
            $mailbox | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msolusercheck.IsLicensed -force
            $mailbox | add-member -type noteproperty -name "Licenses_Destination" -Value ($msolusercheck.Licenses.AccountSkuID -join ",") -force
            $mailbox | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $mailboxcheck.IsDirSynced -force
            $mailbox | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $msolusercheck.PreferredDataLocation -force
            $mailbox | add-member -type noteproperty -name "Database_Destination" -Value $mailboxcheck.Database -force
            $mailbox | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $msolusercheck.BlockCredential -force
            $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $mailboxcheck.PrimarySmtpAddress -force
            $mailbox | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $mailboxcheck.RecipientTypeDetails -force
            $mailbox | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
            $mailbox | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force

        }
        else 
        {
            $mailbox | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False -force
            $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
            $mailbox | add-member -type noteproperty -name "IsLicensed_Destination" -Value $null -force
            $mailbox | add-member -type noteproperty -name "Licenses_Destination" -Value $null -force
            $mailbox | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $null -force
            $mailbox | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $null -force
            $mailbox | add-member -type noteproperty -name "Database_Destination" -Value $null -force
            $mailbox | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $null -force
            $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $null -force
            $mailbox | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force 
            $mailbox | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $null -force
            $mailbox | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $null -force
        }
        Write-host " .. done" -foregroundcolor green
        $AllUsers += $mailbox
    }
    $allUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath
}

Match-AllMailUsers -ImportCSV "C:\Users\fred5646\Rackspace Inc\MPS-TS-Dermavant - General\BOX_user_details.csv" -NewDomain dermavant.com -OutputCSVFilePath "C:\Users\fred5646\Rackspace Inc\MPS-TS-Dermavant - General\MatchedUsers_Dermavant.csv"

