# Match Mailboxes and add to same spreadsheet. Check based on UPN, DisplayName
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

$WorkSheetName = "Wave6_Details"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path "C:\Users\amedrano\Arraya Solutions\Spectra - External - Ext - 1777 Spectra to OVG T2T Migration\Project Files\1777-Spectra-OVG-Users Master List.xlsx"
$tenant = "OVG"
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
foreach ($user in $waveGroup) {
    #progress bar
    Write-ProgressHelper -Activity "Gathering Details for $($user.Name) in $($tenant) tenant" -ProgressCounter ($progresscounter++) -TotalCount ($waveGroup).count
    #Clear Variables
    $msoluser = @()
    $recipient = @()
    $mailbox = @()
    $mbxStats = @()
    $EmailAddresses = @()
    $EmailAddress = $null

    if ($null -ne $user.$tenant) {
        $EmailAddress = $user.$tenant.Trim()
    }
    else {
        $EmailAddress = $null
    }    
    #Address Check
    if ($null -ne $EmailAddress) {    
        if ($msoluser = Get-MsolUser -UserPrincipalName $EmailAddress -ErrorAction SilentlyContinue) {
            Write-Host "$($EmailAddress) found" -ForegroundColor Green
            $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $true -force
        }
        #Name Check
        elseif ($msoluser = Get-MsolUser -SearchString $user.FirstLast -ErrorAction SilentlyContinue) {
            Write-Host "$($user.FirstLast) found" -ForegroundColor Yellow
            $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $true -force  
        }
        #Name Check
        elseif ($msoluser = Get-MsolUser -SearchString $user.Name -ErrorAction SilentlyContinue) {
            Write-Host "$($user.Name) found" -ForegroundColor Yellow
            $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $true -force  
        }
        else {
            Write-Host " Unable to find user for $($user.Name)" -ForegroundColor Red
            $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $false -force
        }
    }
    #Name Check
    elseif ($msoluser = Get-MsolUser -SearchString $user.FirstLast -ErrorAction SilentlyContinue) {
        Write-Host "$($user.FirstLast) found" -ForegroundColor Yellow
        $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $true -force  
    }
    #Name Check
    elseif ($msoluser = Get-MsolUser -SearchString $user.Name -ErrorAction SilentlyContinue) {
        Write-Host "$($user.Name) found" -ForegroundColor Yellow
        $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $true -force  
    }
    else {
        Write-Host " Unable to find user for $($user.Name)" -ForegroundColor Red
        $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $false -force
    }
    if ($msoluser) {
        #Pull Mailbox Stats
        $recipient = Get-EXORecipient $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        if ($mailbox = Get-EXOMailbox $msoluser.UserPrincipalName -PropertySets addresslist, archive, delivery, minimum -ErrorAction SilentlyContinue) {
            $mbxStats = Get-EXOMailboxStatistics $mailbox.PrimarySMTPAddress -PropertySets All -ErrorAction SilentlyContinue
        }
        else {
            $mbxStats = $null
        }
        $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses
    }
        
        $user | add-member -type noteproperty -name "DisplayName_$($tenant)" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_$($tenant)" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_$($tenant)" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses_$($tenant)" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_$($tenant)" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_$($tenant)" -Value $recipient.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_$($tenant)" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_$($tenant)" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_$($tenant)" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_$($tenant)" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_$($tenant)" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_$($tenant)" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_$($tenant)" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_$($tenant)" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_$($tenant)" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($tenant)" -Value $mailbox.ArchiveStatus -force
}
$waveGroup | Export-Excel -WorksheetName "Wave6_Details2" -Path "C:\Users\amedrano\Arraya Solutions\Spectra - External - Ext - 1777 Spectra to OVG T2T Migration\Project Files\1777-Spectra-OVG-Users Master List.xlsx" -Show
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan


# RE-RUN Match Mailboxes and add to same spreadsheet. Check based on UPN, DisplayName
$pilotusers2 = Import-Excel -WorksheetName "Pilot Wave 2"
$HeadingFilter = "TempSolution_Note"
$Filter = "Re-Do"
$tenant = "OVG"
$progressref = ($wave3).count
$progresscounter = 0
foreach ($user in $wave3) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details for $($user.Name) in $($tenant) tenant"
    $EmailAddress = $null
    $EmailAddress = $user.$tenant

    #Clear Variables
    $msoluser = @()
    $recipient = @()
    $mailbox = @()
    $mbxStats = @()
    if ($user.$HeadingFilter -eq $Filter) {
        #Address Check
        if ($msoluser = Get-MsolUser -UserPrincipalName $EmailAddress -ErrorAction SilentlyContinue) {
            Write-Host "$($EmailAddress) found" -ForegroundColor Green
            $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $true -force
        }
        #Name Check
        elseif ($msoluser = Get-MsolUser -SearchString $user.Name -ErrorAction SilentlyContinue) {
            Write-Host "$($user.Name) found" -ForegroundColor Yellow
            $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $true -force  
        }
        else {
            Write-Host " Unable to find user for $($user.Name)" -ForegroundColor Red
            $user | add-member -type noteproperty -name "ExistsIn$($tenant)" -Value $false -force
            $msoluser = $null
        }
        if ($msoluser) {
            #Pull Mailbox Stats
            $recipient = Get-EXORecipient $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
            $mailbox = Get-EXOMailbox $msoluser.UserPrincipalName -PropertySets addresslist, archive, delivery, minimum -ErrorAction SilentlyContinue
            $mbxStats = Get-EXOMailboxStatistics $mailbox.PrimarySMTPAddress -PropertySets All -ErrorAction SilentlyContinue
            $EmailAddresses = $recipient | select -ExpandProperty EmailAddresses
        }
        $user | add-member -type noteproperty -name "DisplayName_$($tenant)" -Value $msoluser.DisplayName -force
        $user | add-member -type noteproperty -name "UserPrincipalName_$($tenant)" -Value $msoluser.userprincipalname -force
        $user | add-member -type noteproperty -name "IsLicensed_$($tenant)" -Value $msoluser.IsLicensed -force
        $user | add-member -type noteproperty -name "Licenses_$($tenant)" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
        $user | add-member -type noteproperty -name "License-DisabledArray_$($tenant)" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_$($tenant)" -Value $mailbox.RecipientTypeDetails -force
        $user | add-member -type noteproperty -name "PrimarySmtpAddress_$($tenant)" -Value $recipient.PrimarySmtpAddress -force
        $user | add-member -type noteproperty -name "Alias_$($tenant)" -Value $recipient.alias -force
        $user | add-member -type noteproperty -name "EmailAddresses_$($tenant)" -Value ($EmailAddresses -join ";") -force
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_$($tenant)" -Value $mailbox.HiddenFromAddressListsEnabled -force
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward_$($tenant)" -Value $mailbox.DeliverToMailboxAndForward -force
        $user | add-member -type noteproperty -name "ForwardingAddress_$($tenant)" -Value $mailbox.ForwardingAddress -force
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress_$($tenant)" -Value $mailbox.ForwardingSmtpAddress -force
        $user | Add-Member -type NoteProperty -Name "MBXSize_$($tenant)" -Value $MBXStats.TotalItemSize -force
        $user | Add-Member -Type NoteProperty -name "MBXItemCount_$($tenant)" -Value $MBXStats.ItemCount -force
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($tenant)" -Value $mailbox.ArchiveStatus -force
    }
}
$wave3 | Export-Excel -WorksheetName "Wave3_Details2" -Path "C:\Users\amedrano\Arraya Solutions\Spectra - External - Ext - 1777 Spectra to OVG T2T Migration\Project Files\1777-Spectra-OVG-Users Master List.xlsx"

#check for users in tenant
$foundUsers2 = @()
$notfoundUsers2 = @()
foreach ($user in $newspectraUsers) {
    $OVGUPN = $user.OakviewUPN
    #Check if user exists
    if ($msoluserCheck = Get-MsolUser -UserPrincipalName $OVGUPN -ErrorAction SilentlyContinue) {
        Write-Host "Found $($msolUserCheck.DisplayName) $($msolUserCheck.UserPrincipalName)" -ForegroundColor Green
        $foundUsers2 += $msoluserCheck
    }
    else {
        Write-Host "Not Found user $($OVGEmailAddress)" -ForegroundColor Red
        $notfoundUsers2 += $user
    }
}

#check for users in tenant - DisplayName
foreach ($user in $notfoundUsers2) {
    #Check if user exists
    if ($msoluserCheck = Get-MsolUser -SearchString $user.DisplayName -ErrorAction SilentlyContinue) {
        $foundUsers2 += $msoluserCheck
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "FoundUPN" -Value $msoluserCheck.UserPrincipalName -Force
    }
    else {
        Write-Host "." -ForegroundColor Red -NoNewline
        $user | add-member -type noteproperty -name "FoundUPN" -Value $False -Force
    }
}

# Update DisplayName and license users
$spectraUsers = Import-CSV

$progressref = ($spectraUsers).count
$progresscounter = 0
$matchedDisplayName = @()
$updatedUsers = @()
$notfound4 = @()
foreach ($user in $spectraUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    $oakviewUPN = $user.OakviewUPN
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Names for $($oakviewUPN)"

    Write-Host "Checking User $($oakviewUPN) .. " -NoNewline -ForegroundColor Cyan
    [String]$DisplayName = $user.DisplayName
    if ($msoluserCheck = Get-MsolUser -UserPrincipalName $oakviewUPN -ErrorAction SilentlyContinue) {
        if ($msoluserCheck.DisplayName -eq $user.DisplayName) {
            Write-Host "Skipping user. Name Matches" -ForegroundColor DarkYellow
            $matchedDisplayName += $user
        }
        else {
            Set-MsolUser -UserPrincipalName $oakviewUPN -DisplayName $DisplayName
            Write-Host "Updated DisplayName" -ForegroundColor Green
            $updatedUsers += $user
        }
    }
    else {
        Write-Host "User not found." -ForegroundColor red
        $notfound4 += $user
    }
}


#Set Password 
$waveremaining = Import-CSV

$progressref = ($waveremaining).count
$progresscounter = 0
foreach ($user in $waveremaining) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    $oakviewUPN = $user.OakviewUPN
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Names for $($oakviewUPN)"

    Write-Host "Checking User $($oakviewUPN) .. " -NoNewline -ForegroundColor Cyan
    if ($msoluserCheck = Get-MsolUser -UserPrincipalName $oakviewUPN -ErrorAction SilentlyContinue) {
        Write-Host "Updated Password" -ForegroundColor Green
        Set-MsolUserPassword -UserPrincipalName $oakviewUPN -NewPassword "O@kV!ew2022$"  -ForceChangePassword $true
    }
    else {
        Write-Host "User not found." -ForegroundColor red
    }
}

# Add Spectra Email
$allSpectraUsers = Import-CSV

$progressref = ($allSpectraUsers).count
$progresscounter = 0
foreach ($user in $allSpectraUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Names for $($user.DisplayName)"
    $spectraAddressSplit = $user.userPrincipalName -split "@"
    $spectraAddress = $spectraAddressSplit[0] + "@spectraxp.com"
    $bruleeCatering = $spectraAddressSplit[0] + "@brulee-catering.com"
    $niagaraFalls  = $spectraAddressSplit[0] + "@niagarafalls-cc.com"
    $bruleeEverday  = $spectraAddressSplit[0] + "@brulee-everyday.com"
    $DisplayName = $user.DisplayName
    if ($mailboxCheck = Get-Mailbox $spectraAddress -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraEmailAddress" -Value $mailboxCheck.PrimarySmtpAddress -Force
    }
    elseif ($mailboxCheck = Get-Mailbox $bruleeCatering -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraEmailAddress" -Value $mailboxCheck.PrimarySmtpAddress -Force
    }
    elseif ($mailboxCheck = Get-Mailbox $niagaraFalls -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraEmailAddress" -Value $mailboxCheck.PrimarySmtpAddress -Force
    }
    elseif ($mailboxCheck = Get-Mailbox $bruleeEverday -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraEmailAddress" -Value $mailboxCheck.PrimarySmtpAddress -Force
    }
    elseif ($mailboxCheck = Get-Mailbox $DisplayName -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraEmailAddress" -Value $mailboxCheck.PrimarySmtpAddress -Force
    }
    else {
        Write-Host "." -ForegroundColor red -NoNewline
        $user | add-member -type noteproperty -name "SpectraEmailAddress" -Value $null -Force
    }

    # MSOLUserCheck
    if ($msolUserCheck = Get-MsolUser -UserPrincipalName $mailboxCheck.UserPrincipalName -ErrorAction SilentlyContinue) {
        $user | add-member -type noteproperty -name "SpectraUPN" -Value $msolUserCheck.UserPrincipalName -Force
        $user | add-member -type noteproperty -name "SpectraIsLicensed" -Value $msolUserCheck.IsLicensed -Force
    }
    else {
        $user | add-member -type noteproperty -name "SpectraUPN" -Value $null -Force
        $user | add-member -type noteproperty -name "SpectraIsLicensed" -Value $null -Force
    }

}

# Add Spectra Email - nonSpectraAddressUser
$nonSpectraAddressUser = Import-CSV

$progressref = ($nonSpectraAddressUser).count
$progresscounter = 0
foreach ($user in $nonSpectraAddressUser) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Names for $($user.DisplayName)"
    $spectraAddressSplit = $user.userPrincipalName -split "@"
    $spectraAddress = $spectraAddressSplit[0] + "@spectraxp.com"
    $bruleeCatering = $spectraAddressSplit[0] + "@brulee-catering.com"
    $niagaraFalls  = $spectraAddressSplit[0] + "@niagarafalls-cc.com"
    $bruleeEverday  = $spectraAddressSplit[0] + "@brulee-everyday.com"
    $DisplayName = $user.DisplayName
    if ($msolUserCheck = Get-MsolUser -SearchString $spectraAddress -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraUPN" -Value $msolUserCheck.UserPrincipalName -Force
        $user | add-member -type noteproperty -name "SpectraIsLicensed" -Value $msolUserCheck.IsLicensed -Force
    }
    elseif ($msolUserCheck = Get-MsolUser -SearchString $bruleeCatering -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraUPN" -Value $msolUserCheck.UserPrincipalName -Force
        $user | add-member -type noteproperty -name "SpectraIsLicensed" -Value $msolUserCheck.IsLicensed -Force
    }
    elseif ($msolUserCheck = Get-MsolUser -SearchString $niagaraFalls -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraUPN" -Value $msolUserCheck.UserPrincipalName -Force
        $user | add-member -type noteproperty -name "SpectraIsLicensed" -Value $msolUserCheck.IsLicensed -Force
    }
    elseif ($msolUserCheck = Get-MsolUser -SearchString $bruleeEverday -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraUPN" -Value $msolUserCheck.UserPrincipalName -Force
        $user | add-member -type noteproperty -name "SpectraIsLicensed" -Value $msolUserCheck.IsLicensed -Force
    }
    elseif ($msolUserCheck = Get-MsolUser -SearchString $DisplayName -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SpectraUPN" -Value $msolUserCheck.UserPrincipalName -Force
        $user | add-member -type noteproperty -name "SpectraIsLicensed" -Value $msolUserCheck.IsLicensed -Force
    }
    else {
        Write-Host "." -ForegroundColor red -NoNewline
        $user | add-member -type noteproperty -name "SpectraUPN" -Value $null -Force
        $user | add-member -type noteproperty -name "SpectraIsLicensed" -Value $null -Force
    }

    #mailboxCheck
    if ($msolUserCheck) {
        $mailboxCheck = Get-Mailbox $spectraAddress -ErrorAction SilentlyContinue
        Write-Host "." -ForegroundColor DarkGreen -NoNewline
        $user | add-member -type noteproperty -name "SpectraEmailAddress" -Value $mailboxCheck.PrimarySMTPAddress -Force
    }
    else {
        Write-Host "." -ForegroundColor DarkRed -NoNewline
        $user | add-member -type noteproperty -name "SpectraEmailAddress" -Value $null -Force
    }
}

# Check For Exclude user
$allSpectraUsers = Import-CSV

$progressref = ($allSpectraUsers).count
$progresscounter = 0
foreach ($user in $allSpectraUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Names for $($user.DisplayName)"

    if ($excludeUsercheck = $spectraExclude | ? {$_."Spectra Email" -eq $user.SpectraEmailAddress}) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "ExcludeUser" -Value $True -Force
    }
    elseif ($excludeUsercheck = $spectraExclude | ? {$_.DisplayName -eq $user.name}) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "ExcludeUser" -Value $True -Force
    }
    else {
        Write-Host "." -ForegroundColor red -NoNewline
        $user | add-member -type noteproperty -name "ExcludeUser" -Value $False -Force
    }
    
}

#Set CompanyName to Spectra
# Specify if Wave 1
$allSpectraUsers = Import-CSV

$allSpectraUsers = $allSpectraUsers | ?{$_.ExcludeUser -eq $false}
$wave1 = $allSpectraUsers | ?{$_."Wave1 Comm" -eq "Yes"}
$wave2 = $allSpectraUsers | ?{$_."Wave2 Comm" -eq "Yes"}
$wave3 = $allSpectraUsers | ?{$_."Wave3 Comm" -eq "Yes"}

$progressref = ($allSpectraUsers).count
$progresscounter = 0
foreach ($user in $allSpectraUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Setting Company Name for $($user.DisplayName)"
    
    Get-AzureADUser -SearchString $user.OakviewUPN | Set-AzureADUser -CompanyName "Spectra"
}

# Add Members to Exchange Distribution Groups
$allSpectraUsers = Import-CSV

$allSpectraUsers = $allSpectraUsers | ?{$_.ExcludeUser -eq $false}
$wave1 = $allSpectraUsers | ?{$_."Wave1 Comm" -eq "Yes"}
$wave2 = $allSpectraUsers | ?{$_."Wave2 Comm" -eq "Yes"}
$wave3 = $allSpectraUsers | ?{$_."Wave3 Comm" -eq "Yes"}

$AllErrors = @()
$progressref = ($wave3).count
$progresscounter = 0
$waveGroup = "wave3@spectraxp.com"
foreach ($User in $wave3) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    $userAddress = $user.SpectraEmailAddress
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding member $($userAddress) to $($waveGroup)"
    
    #Add Member to Distribution Groups     
    try {
        Add-DistributionGroupMember -Identity $waveGroup -Member $userAddress -ea Stop
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    catch {
        Write-Host ". " -ForegroundColor red -NoNewline

        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
        $currenterror | Add-Member -type NoteProperty -Name "Group" -Value $waveGroup -Force
        $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $userAddress-Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllErrors += $currenterror           
        continue
    }
}

# $excludeUsers Members to Exchange Distribution Groups
$excludeUsers = Import-Csv
$AllErrors = @()
$progressref = ($excludeUsers).count
$progresscounter = 0
foreach ($group in $excludeUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Removing member $($group."Spectra Email")"
    
    #Add Member to Distribution Groups     
    try {
        Remove-DistributionGroupMember -Identity "wave1@spectraxp.com" -Member $group."Spectra Email" -Confirm:$false -ea Stop
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    catch {
        Write-Host ". " -ForegroundColor red -NoNewline

        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
        $currenterror | Add-Member -type NoteProperty -Name "Group" -Value "wave1@spectraxp.com" -Force
        $currenterror | Add-Member -type NoteProperty -Name "Destination_PrimarySMTPAddress" -Value $group.SpectraEmailAddress -Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllErrors += $currenterror           
        continue
    }
        
    Write-Host " done " -ForegroundColor Green
}

#Check If MFA Enabled
$allSpectraUsers = Import-excel
$AuthMethodsRegUsers = Import-Csv

$progressref = ($allSpectraUsers).count
$progresscounter = 0
foreach ($user in $allSpectraUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Names for $($user.DisplayName)"

    if ($Usercheck = $AuthMethodsRegUsers | ? {$_."User Name" -eq $user.OakviewUPN}) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SSPR Registered" -Value $Usercheck."SSPR Registered" -Force
        $user | add-member -type noteproperty -name "SSPR Enabled" -Value $Usercheck."SSPR Enabled" -Force
        $user | add-member -type noteproperty -name "MFA Registered" -Value $Usercheck."MFA Registered" -Force
        $user | add-member -type noteproperty -name "Methods Registered" -Value $Usercheck."Methods Registered" -Force

    }
    else {
        Write-Host "." -ForegroundColor red -NoNewline
        $user | add-member -type noteproperty -name "SSPR Registered" -Value $null -Force
        $user | add-member -type noteproperty -name "SSPR Enabled" -Value $null -Force
        $user | add-member -type noteproperty -name "MFA Registered" -Value $null -Force
        $user | add-member -type noteproperty -name "Methods Registered" -Value $null -Force
    }  
}

#Check If MFA Capable
$allSpectraUsers = Import-excel
$UserRegistrationReport = Import-Csv

$progressref = ($allSpectraUsers).count
$progresscounter = 0
foreach ($user in $allSpectraUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Names for $($user.DisplayName)"

    if ($Usercheck = $UserRegistrationReport | ? {$_.userPrincipalName -eq $user.OakviewUPN}) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "mfaCapable" -Value $Usercheck.mfaCapable -Force
    }
    else {
        Write-Host "." -ForegroundColor red -NoNewline
        $user | add-member -type noteproperty -name "mfaCapable" -Value $null -Force
    }  
}

# Get MailboxStats for InActive Mailboxes
$InactiveMailboxes = Import-Csv "C:\Users\aaron.medrano\Desktop\OakView\SpectraXP_InactiveMailboxes.csv"
$progressref = ($InactiveMailboxes).count
$progresscounter = 0
foreach ($mailbox in $InactiveMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking MailboxStats for Inactive Mailbox $($mailbox.PrimarySmtpAddress)"
    
    # Get MailboxStats of Deleted Object
    if ($inactiveMailboxStats = Get-MailboxStatistics $mailbox.ExchangeGuid -IncludeSoftDeletedRecipients | select TotalItemSize, ItemCount, @{Name="TotalItemSizeGB"; expression={[math]::Round((($_.TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),3)}}) {
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSize" -Value $inactiveMailboxStats.TotalItemSize -force
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSizeGB" -Value $inactiveMailboxStats.TotalItemSizeGB -force
        $mailbox | add-member -type noteproperty -name "InActive_ItemCount" -Value $inactiveMailboxStats.ItemCount -force
        
    }
    else {
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSize" -Value $null -force
        $mailbox | add-member -type noteproperty -name "InActive_TotalItemSizeGB" -Value $null -force
        $mailbox | add-member -type noteproperty -name "InActive_ItemCount" -Value $null -force
        
    }
}

# to GB
Get-MailboxStatistics ee193bfd-f80f-46e0-bda7-e3588f6787f0 -IncludeSoftDeletedRecipients | Select-Object TotalItemSize, @{Name="TotalItemSizeGB"; expression={[math]::Round((($_.TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),3)}}
# to MB
Get-MailboxStatistics ee193bfd-f80f-46e0-bda7-e3588f6787f0 -IncludeSoftDeletedRecipients | Select-Object TotalItemSize, @{Name="TotalItemSizeMB"; expression={[math]::Round((($_.TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1MB),3)}}

$licensedUsers = Get-MsolUser -All | ?{$_.islicensed -eq $true}

$failedtoAddMembers = @()
#ProgressBar
$progressref = ($licensedUsers).count
$progresscounter = 0
foreach ($user in $licensedUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding DL member $($user.DisplayName)"
	
    if ($mailboxCheck = Get-Mailbox $user.UserPrincipalName -ErrorAction SilentlyContinue) {
        $memberAdd = Add-DistributionGroupMember -identity All-LicensedUsers -Member $mailboxCheck.DistinguishedName
        Write-Host "." -NoNewline -ForegroundColor Green
    }
    else {
        Write-Host "." -NoNewline -ForegroundColor Red
        $failedtoAddMembers += $user
    }
}

#Gather Group Members. Ensure all are UserMailboxes
$AllGroupMembers = Get-DistributionGroupMember -Identity All-LicensedUsers -ResultSize Unlimited
$progressref = ($AllGroupMembers).count
$progresscounter = 0
$nonUesrMailboxes = @()
foreach ($member in $AllGroupMembers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking DL member $($member.DisplayName)"
	
    if ($member.RecipientTypeDetails -ne "UserMailbox") {
        $nonUesrMailboxes += $member
        Write-Host "." -NoNewline -ForegroundColor Red
    }
}


#Update license for pilot users - direct licensing
$progressref = ($spectraUsers).count
$progresscounter = 0
$updatedLicense = @()
$originalLicenseTenant = "spectraxp"
$newLicenseTenant = "oakview"

foreach ($user in $spectraUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    $oakviewUPN = $user."OVG UPN/Email Address"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Licensing for $($oakviewUPN)"

    Write-Host "Checking $($oakviewUPN).. " -NoNewline -ForegroundColor Cyan
    if ($msoluserCheck = Get-MsolUser -UserPrincipalName $oakviewUPN -ErrorAction SilentlyContinue) {
        $licensing = $user.Licenses -split ","
        Write-Host "Updating Licensing" -ForegroundColor DarkCyan -NoNewline
        foreach ($license in $licensing) {
            $licenseUpdate = $license.Replace("$originalLicenseTenant","$newLicenseTenant")
            if (Get-MsolAccountSku | ? {$_.accountskuID -eq $licenseUpdate}) {
                if ($msoluserCheck.Licenses.AccountSkuID -eq $licenseUpdate) {
                    Write-Host "." -ForegroundColor Yellow -NoNewline
                }
                else {
                    Write-Host "." -ForegroundColor Green -NoNewline
                    Set-MsolUserLicense -UserPrincipalName $oakviewUPN -AddLicenses $licenseUpdate
                    $updatedLicense += $user
                }
            }
        }
        Write-Host " done" -ForegroundColor Cyan
        
    }
    else {
        Write-Host "Not found." -ForegroundColor red
    }

    <#disabled license
    if($user.licenses.accountSKUID -match "EnterprisePACK") {
        $DisabledArray = @()
        $allLicenses = ($user).Licenses
        for($i = 0; $i -lt $AllLicenses.Count; $i++)
        {
            $serviceStatus =  $AllLicenses[$i].ServiceStatus
            foreach($service in $serviceStatus)
            {
                if($service.ProvisioningStatus -eq "Disabled")
                {
                    $disabledArray += ($service.ServicePlan).ServiceName
                }
            }
    
        }
        #Update users with Office E3 licenses to Microsoft E3 licenses with DisabledArray above.
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId libertylife:SPE_E3 -DisabledPlans $DisabledArray
        Write-host "Updating E3 license for $user"
        
        try {
        if (!Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses libertylife:SPE_E3 -RemoveLicenses libertylife:ENTERPRISEPACK  -LicenseOptions $LicenseOptions -verbose)
            Write-Host "Completed For $user"
    #>
}


#Remove Pilot Users
$licenseGroup = "OVG360 Exchange Plan 1 Licensed Users"
$progressref = ($spectraUsers).count
$progresscounter = 0
$allErrors = @()
foreach ($user in $spectraUsers) {
    $AZLicenseGroup = Get-AzureADGroup -SearchString $licenseGroup
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    $DesinationPrimarySMTPAddress = $user."OVG UPN/Email Address"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Removing $($user.DisplayName) from $($AZLicenseGroup.DisplayName)"
    $msoluserCheck = Get-MsolUser -SearchString $DesinationPrimarySMTPAddress
    Remove-AzureADGroupMember -ObjectId $AZLicenseGroup.ObjectID -MemberId $msoluserCheck.ObjectId
}

#Update users to license group
$WorkSheetName = "Wave5_Details"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path "C:\Users\amedrano\Arraya Solutions\Spectra - External - Ext - 1777 Spectra to OVG T2T Migration\Project Files\1777-Spectra-OVG-Users Master List.xlsx"
$tenant = "OVG"
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
$allErrors = @()
foreach ($user in $waveGroup) {
     #progress bar
     Write-ProgressHelper -Activity "Updating License for $($user.Name)" -ProgressCounter ($progresscounter++) -TotalCount ($waveGroup).count
    
    #Clear Variables
    $AZLicenseGroup = @()
    $azureUser = @()

    if ($user.UserPrincipalName_OVG) {
        #Get Azure License Group Details
        if ($user.Licenses_Spectra -like "*spectraxp:SPE_E3*") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MS365 E3 Licensed Users"
        }
        elseif ($user.Licenses_Spectra -like "*spectraxp:ENTERPRISEPACK*") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MS365 E3 Licensed Users"
        }
        elseif ($user.Licenses_Spectra -like "*spectraxp:SPB*") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MS Business Premium Licensed Users"
        }
        else {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 Exchange Plan 1 Licensed Users"
        }
    
        if ($azureUser = Get-AzureADUser -SearchString $user.UserPrincipalName_OVG) {
            Write-Host "Adding $($user.Name) to " -NoNewline
            Write-Host "$($AZLicenseGroup.DisplayName) .. " -NoNewline -ForegroundColor Cyan
            #Add user to license group
            try {
                Add-AzureADGroupMember -ObjectId $AZLicenseGroup.ObjectID -RefObjectId $azureUser.ObjectId -ErrorAction Stop
                Write-Host "Added. " -ForegroundColor Green -NoNewline
            }
            catch {
                if ($_.exception -like "*Message: One or more added object references already exist for the following modified properties: 'members'*") {
                    Write-Host "Already A Member " -ForegroundColor Yellow -NoNewline
                }
                else {
                    Write-Host ". " -ForegroundColor red -NoNewline
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "LicenseGroup" -Value $AZLicenseGroup.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $user.UserPrincipalName_OVG -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            #Remove from Exchange Standard Group
            if ($user.Licenses_Spectra -like "*spectraxp:SPE_E3*" -or $user.Licenses_Spectra -like "*spectraxp:ENTERPRISEPACK*" -or $user.Licenses_Spectra -like "*spectraxp:SPB*") {
                $EXOPlan1LicenseGroup = Get-AzureADGroup -SearchString "OVG360 Exchange Plan 1 Licensed Users"
                try {
                    Remove-AzureADGroupMember -ObjectId $EXOPlan1LicenseGroup.ObjectID -MemberId $azureUser.ObjectId -ErrorAction Stop
                    Write-Host "Removed EXO Plan 1. " -ForegroundColor Green -NoNewline
                }
                catch {
                if ($_.exception -like "*does not exist or one of its queried reference-property objects are not present*") {
                    Write-Host "Not A Member of EXO Plan 1 " -ForegroundColor Yellow -NoNewline
                }
                else {
                    Write-Host ". " -ForegroundColor Red
                
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToRemoveMember" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "LicenseGroup" -Value $AZLicenseGroup.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $user.UserPrincipalName_OVG -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }  
                }
            }
            Write-Host "done" -ForegroundColor Cyan
        }
    }

}

# Add PowerBI Standard or Flow - direct licensing
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null

$allErrors = @()
foreach ($user in $waveGroup) {
    #progress bar
    Write-ProgressHelper -Activity "Checking Licensing for $($user.UserPrincipalName_OVG)" -ProgressCounter ($progresscounter++) -TotalCount ($waveGroup).count

    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Licensing for $($user.UserPrincipalName_OVG)"
    $FlowLicense = Get-MsolAccountSku | ? {$_.accountskuID -like "*FLOW_FREE*"}
    $PowerBILicense = Get-MsolAccountSku | ? {$_.accountskuID -like "*POWER_BI_STANDARD*"}

    if ($user.UserPrincipalName_OVG) {
        if ($azureUser = Get-AzureADUser -SearchString $user.UserPrincipalName_OVG) {
            if ($user.Licenses_Spectra -like "*FLOW_FREE*") {
                Write-Host "Add FLOW_FREE to $($user.UserPrincipalName_OVG)" -ForegroundColor Green -NoNewline
                try {
                    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_OVG -AddLicenses $FlowLicense.AccountSkuId -ErrorAction Stop
                    Write-Host " done" -ForegroundColor Green
                }
                catch {
                    Write-Host ". " -ForegroundColor Red
                
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddFlow" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $user.UserPrincipalName_OVG -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            if ($user.Licenses_Spectra -like "*POWER_BI_STANDARD*") {
                Write-Host "Add POWER_BI_STANDARD to $($user.UserPrincipalName_OVG)" -ForegroundColor Green
                try {
                    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName_OVG -AddLicenses $PowerBILicense.AccountSkuId -ErrorAction Stop
                    Write-Host " done" -ForegroundColor Green
                }
                catch {
                    Write-Host ". " -ForegroundColor Red
                
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddFlow" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $user.UserPrincipalName_OVG -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
        }
    } 
}



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

#Temporary Cutover Users
$WorkSheetName = "Wave5_Details"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path "C:\Users\amedrano\Arraya Solutions\Spectra - External - Ext - 1777 Spectra to OVG T2T Migration\Project Files\1777-Spectra-OVG-Users Master List.xlsx"
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null

$allErrors = @()
foreach ($user in $waveGroup) {
    $DestinationPrimarySMTPAddress = $user.PrimarySmtpAddress_OVG 
    $SourcePrimarySMTPAddress = $user.PrimarySmtpAddress_Spectra
    #progress bar
    Write-ProgressHelper -Activity "Updating Forwarding for $($SourcePrimarySMTPAddress)" -ProgressCounter ($progresscounter++) -TotalCount ($waveGroup).count

    if ($user.ToMoveOrNotToMove -eq "Move") {
        if ($user.RecipientTypeDetails_Spectra -and $user.RecipientTypeDetails_OVG) {
            Write-Host "Cutting Over User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
            ## Set Mailbox to Forward from Source to Destination Mailbox
            Write-Host "Set Forward on Mailbox $($DestinationPrimarySMTPAddress)  " -foregroundcolor Magenta -nonewline
            Try{       
                Set-Mailbox $SourcePrimarySMTPAddress -ForwardingSmtpAddress $DestinationPrimarySMTPAddress -ErrorAction Stop -warningaction silentlycontinue
                Write-Host ". " -ForegroundColor Green
            }
            Catch {
                Write-Host ". " -ForegroundColor red
    
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SetForward" -Force
                $currenterror | Add-Member -type NoteProperty -Name "Name" -Value $user.Name -Force
                $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_Source" -Value $SourcePrimarySMTPAddress -Force
                $currenterror | Add-Member -type NoteProperty -Name "DesinationPrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $allErrors += $currenterror           
                continue
            }
        }
    }
}

#Remove Current Licensing -Brulee
$progressref = ($bruleeUsers).count
$progresscounter = 0
$allErrors = @()
foreach ($user in $bruleeUsers) {
    #Clear Variables
    $AZLicenseGroup = @()
    $azureUser = @()

    if ($user.UserPrincipalName_OVG) {
        #Get Azure License Group Details
        if ($user.Licenses_Spectra -like "*spectraxp:SPE_E3*") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MS365 E3 Licensed Users"
        }
        elseif ($user.Licenses_Spectra -like "*spectraxp:ENTERPRISEPACK*") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MS365 E3 Licensed Users"
        }
        elseif ($user.Licenses_Spectra -like "*spectraxp:SPB*") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MS Business Premium Licensed Users"
        }
    
        #progress bar
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"    
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating License for $($user.Name)"
    
        if ($azureUser = Get-AzureADUser -SearchString $user.UserPrincipalName_OVG) {
            Write-Host "Removing $($user.Name) from " -NoNewline
            Write-Host "$($AZLicenseGroup.DisplayName) .. " -NoNewline -ForegroundColor Cyan
            #Add user to license group
            try {
                Remove-AzureADGroupMember -ObjectId $AZLicenseGroup.ObjectID -MemberId $azureUser.ObjectId -ErrorAction Stop
                Write-Host "Removed. " -ForegroundColor Green -NoNewline
            }
            catch {
                if ($_.exception -like "*Message: One or more added object references already exist for the following modified properties: 'members'*") {
                    Write-Host "Already A Member " -ForegroundColor Yellow -NoNewline
                }
                else {
                    Write-Host ". " -ForegroundColor red -NoNewline
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToRemoveMember" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "LicenseGroup" -Value $AZLicenseGroup.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $user.UserPrincipalName_OVG -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            #Remove from Exchange Standard Group
            if ($user.Licenses_Spectra -like "*spectraxp:SPE_E3*" -or $user.Licenses_Spectra -like "*spectraxp:ENTERPRISEPACK*" -or $user.Licenses_Spectra -like "*spectraxp:SPB*") {
                $EXOPlan1LicenseGroup = Get-AzureADGroup -SearchString "OVG360 Exchange Plan 1 Licensed Users"
                try {
                    Add-AzureADGroupMember -ObjectId $EXOPlan1LicenseGroup.ObjectID -RefObjectId $azureUser.ObjectId -ErrorAction Stop
                    Write-Host "Added EXO Plan 1. " -ForegroundColor Green -NoNewline
                }
                catch {
                if ($_.exception -like "*does not exist or one of its queried reference-property objects are not present*") {
                    Write-Host "Already A Member of EXO Plan 1 " -ForegroundColor Yellow -NoNewline
                }
                else {
                    Write-Host ". " -ForegroundColor Red
                
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToAddMember" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "LicenseGroup" -Value $EXOPlan1LicenseGroup.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $user.UserPrincipalName_OVG -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }  
                }
            }
            Write-Host "done" -ForegroundColor Cyan
        }
    }

}

#Gather Sign in Logs
AzureADPreview\Connect-AzureAD
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

$WorkSheetName = "Master_MatchedUsers"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path "C:\Users\amedrano\Arraya Solutions\Spectra - External - Ext - 1777 Spectra to OVG T2T Migration\Project Files\1777-Spectra-OVG-Users Master List.xlsx"
$tenant = "OVG"

#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
foreach ($user in $waveGroup | sort "UserPrincipalName_$tenant") {
    if ($ProgressPreference = "SilentlyContinue") {
        $ProgressPreference = "Continue"
    }
    if ($user."ExistsIn$tenant" -eq $true) {
        $UPN = $user."UserPrincipalName_$tenant".Trim().ToLower()
        #progress bar
        Write-ProgressHelper -Activity "Gathering Azure Audit Logs $($UPN)" -ProgressCounter ($progresscounter++) -TotalCount ($waveGroup).count

        #Gather Azure Audit Logs
        try {
            $azureAuditLogsDetails = Get-AzureADAuditSignInLogs -Filter "UserPrincipalName eq '$UPN'" -Top 1 -ErrorAction Stop | Select CreatedDateTime, AppDisplayName, ClientAppUsed, Status
        }
        catch {
        if ($_.exception -match "Message:\s*(This request is throttled|Too Many Requests)") {
            $i = 0
            do {
                Write-Host "retry" -ForegroundColor Red -NoNewline
                $i += 1
                $statusError = "Attempt #$i - $($((Get-Date) - $global:start).ToString('hh\:mm\:ss'))"
                Write-Progress -Activity "Retrying Lookup"  -Status $statusError -Id 2
                Start-Sleep -Seconds 60
                $azureAuditLogsDetails = Get-AzureADAuditSignInLogs -Filter "UserPrincipalName eq '$UPN'" -Top 1 -ErrorAction Stop | Select CreatedDateTime, AppDisplayName, ClientAppUsed, Status
                Write-Progress -Activity "Retrying Lookup"  -Status $statusError -Id 2 -Completed
            } until (
                $i -eq 3 -or $azureAuditLogsDetails
            )
        }
        else {
            Write-Host $_.exception -NoNewline
        }
        }

        if ($azureAuditLogsDetails.Status.ErrorCode -eq "0") {
        $azureSignInStatus = "Success"
        Write-Host "." -ForegroundColor Green -NoNewline
        }
        elseif ($null -eq $azureAuditLogsDetails) {
        $azureSignInStatus = "NoResults"
        Write-Host "." -ForegroundColor DarkMagenta -NoNewline
        }
        else {
        $azureSignInStatus = $azureAuditLogsDetails.Status.FailureReason
        Write-Host "." -ForegroundColor Red -NoNewline
        }
        #Gather Results
        $user | Add-Member -type NoteProperty -Name "AzureLastLoginTime_$tenant" -Value $azureAuditLogsDetails.CreatedDateTime -Force
        $user | Add-Member -type NoteProperty -Name "AzureAppDisplayname_$tenant" -Value $azureAuditLogsDetails.AppDisplayName -Force
        $user | Add-Member -type NoteProperty -Name "AzureClientAppUsed_$tenant" -Value $azureAuditLogsDetails.ClientAppUsed -Force
        $user | Add-Member -type NoteProperty -Name "AzureLoginStatus_$tenant" -Value $azureSignInStatus -Force
    }
}
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan

## testing
$WorkSheetName = "Wave6_Details"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path "C:\Users\amedrano\Arraya Solutions\Spectra - External - Ext - 1777 Spectra to OVG T2T Migration\Project Files\1777-Spectra-OVG-Users Master List.xlsx"
$tenant = "OVG"

$global:progressref = ($waveGroup).count
$progresscounter = 0
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
foreach ($user in $waveGroup) {
    Write-ProgressHelper -Activity "Gathering Details for $($user.Name) in $($tenant) tenant" -ProgressCounter ($progresscounter++)
    $user.$tenant.Trim()
    Start-Sleep –Milliseconds (Get-Random –Minimum 700 –Maximum 2000)
}

function Write-ProgressHelper2 {
	param (
	    [int]$ProgressCounter,
	    [string]$Activity
	)
    $progresspercentcomplete = [math]::Round((($progresscounter / $global:progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$global:progressref+"]"
	Write-Progress -Activity $Activity -Status $progressStatus -PercentComplete $progresspercentcomplete
}