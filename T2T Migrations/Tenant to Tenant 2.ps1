## Start Script ##
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
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path ~\1777-OldCompany-OVG-Users Master List.xlsx"
$tenant = "OVG"
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = $wavegroup.count
foreach ($user in $waveGroup) {
    #progress bar
    Write-ProgressHelper -Activity "Gathering Details for $($user.Name) in $($tenant) tenant" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount
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
$waveGroup | Export-Excel -WorksheetName "Wave6_Details2" -Path ~\1777-OldCompany-OVG-Users Master List.xlsx" -Show
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan
## End Script ##

#Set CompanyName to OldCompany
# Specify if Wave 1
$allOldCompanyUsers = Import-CSV

$allOldCompanyUsers = $allOldCompanyUsers | ?{$_.ExcludeUser -eq $false}
$wave1 = $allOldCompanyUsers | ?{$_."Wave1 Comm" -eq "Yes"}
$wave2 = $allOldCompanyUsers | ?{$_."Wave2 Comm" -eq "Yes"}
$wave3 = $allOldCompanyUsers | ?{$_."Wave3 Comm" -eq "Yes"}

$progressref = ($allOldCompanyUsers).count
$progresscounter = 0
foreach ($user in $allOldCompanyUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Setting Company Name for $($user.DisplayName)"
    
    Get-AzureADUser -SearchString $user.NewCompanyUPN | Set-AzureADUser -CompanyName "OldCompany"
}

# Get MailboxStats for InActive Mailboxes
$InactiveMailboxes = Import-Csv "C:\Users\aaron.medrano\Desktop\NewCompany\OldCompanyXP_InactiveMailboxes.csv"
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


#Update users to license group

$WorkSheetName = "UserMailboxes2"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path "~\1777-OldCompany-OVG-AllMatched-Mailboxes.xlsx"
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
$AddedUsers = @()
$allErrors = @()
foreach ($user in $waveGroup) {
     #progress bar
     Write-ProgressHelper -Activity "Updating License for $($user.Name)" -ProgressCounter ($progresscounter++) -TotalCount ($waveGroup).count
    
    #Clear Variables
    $AZLicenseGroup = @()
    $azureUser = @()

    if ($user.UserPrincipalName_OVG) {
        $licenses = $user.Licenses_OldCompany.split(",")
        #Get Azure License Group Details
        if ($licenses -match "ATP_ENTERPRISE" -and $licenses -match "MCOMEETADV" -and $licenses -match "SPE_E3" -and $licenses -match "FLOW_FREE") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MicrosoftE3_ATP P1_Audio License Group"
        }
        elseif ($licenses -match "ENTERPRISEPACK" -and $licenses -match "ATP_ENTERPRISE" -and $licenses -match "AAD_PREMIUM" -and $licenses -match "MCOMEETADV") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 OfficeE3_AADP1_ATPP1_AudioConf_NonManagedDevice License"
        }
        elseif ($licenses -match "ENTERPRISEPACK" -and $licenses -match "ATP_ENTERPRISE" -and $licenses -match "AAD_PREMIUM") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 OfficeE3_ADP1_ATPP1 No_OldCompanyIT_Workstation"
        }
        elseif ($licenses -match "SPE_F1" -and $licenses -match "EXCHANGESTANDARD") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 Email Outlook Access Licensed Users"
        }
        elseif ($licenses -match "SPB" -and $licenses -match "MCOMEETADV" ) {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 Business Premium & Audio Conferencing License Group"
        }
        elseif ($licenses -match "ATP_ENTERPRISE" -and $licenses -match "SPE_E3") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 M365 E3 ATP P1 License Group"
        }
        elseif ($licenses -match "SPB") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 Corporate Users License Group"
        }
        elseif ($licenses -match "SPE_F1") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 Email Outlook Access Licensed Users"
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
                $AddedUsers += $user
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
            #"OVG360 MS365 E3 Licensed Users"
            #Remove from Exchange Standard Group
            $allLicensesGroups = @("OVG360 MS Business Premium Licensed Users","OVG360 Email Outlook Access Licensed Users","OVG360 Email Outlook Access Licensed Users","OVG360 Email Web Only License Group","OVG360 Exchange Plan 1 Licensed Users")
            #Remove from Groups
            foreach ($license in $allLicensesGroups | ?{$_ -ne $AZLicenseGroup.DisplayName}){
                $RemoveLicenseGroup = Get-AzureADGroup -SearchString $license
                try {
                    Remove-AzureADGroupMember -ObjectId $RemoveLicenseGroup.ObjectID -MemberId $azureUser.ObjectId -ErrorAction Stop
                    Write-Host "Removed $($RemoveLicenseGroup.DisplayName)" -ForegroundColor Green -NoNewline
                }
                catch {
                    if ($_.exception -like "*does not exist or one of its queried reference-property objects are not present*") {
                        Write-Host "." -ForegroundColor Yellow -NoNewline
                    }
                    else {
                        Write-Host ". " -ForegroundColor Red
                    
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity -Force
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToRemoveMember" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "LicenseGroup" -Value $RemoveLicenseGroup.DisplayName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $user.UserPrincipalName_OVG -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }  
                }
            }
                
        }
        Write-Host "done" -ForegroundColor Cyan
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
            if ($user.Licenses_OldCompany -like "*FLOW_FREE*") {
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
            if ($user.Licenses_OldCompany -like "*POWER_BI_STANDARD*") {
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
$WorkSheetName = "Wave6_Details"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path ~\1777-OldCompany-OVG-Users Master List.xlsx"
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null

$allErrors = @()
foreach ($user in $waveGroup) {
    $DestinationPrimarySMTPAddress = $user.PrimarySmtpAddress_OVG 
    $SourcePrimarySMTPAddress = $user.PrimarySmtpAddress_OldCompany
    #progress bar
    Write-ProgressHelper -Activity "Updating Forwarding for $($SourcePrimarySMTPAddress)" -ProgressCounter ($progresscounter++) -TotalCount ($waveGroup).count

    if ($user.ToMoveOrNotToMove -eq "Move") {
        if ($user.RecipientTypeDetails_OldCompany -and $user.RecipientTypeDetails_OVG) {
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
        if ($user.Licenses_OldCompany -like "*OldCompanyxp:SPE_E3*") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MS365 E3 Licensed Users"
        }
        elseif ($user.Licenses_OldCompany -like "*OldCompanyxp:ENTERPRISEPACK*") {
            $AZLicenseGroup = Get-AzureADGroup -SearchString "OVG360 MS365 E3 Licensed Users"
        }
        elseif ($user.Licenses_OldCompany -like "*OldCompanyxp:SPB*") {
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
            if ($user.Licenses_OldCompany -like "*OldCompanyxp:SPE_E3*" -or $user.Licenses_OldCompany -like "*OldCompanyxp:ENTERPRISEPACK*" -or $user.Licenses_OldCompany -like "*OldCompanyxp:SPB*") {
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
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path ~\1777-OldCompany-OVG-Users Master List.xlsx"
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

#Gather all Matched users in the Master List
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

$WorkSheetName = "UserMailboxes"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path ~\1777-OldCompany-OVG-MailboxList.xlsx"
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
$waveGroup | Export-Excel -WorksheetName "UserMailboxes" -Path ~\1777-OldCompany-OVG-MailboxList.xlsx" -Show
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan


#Check If MFA Enabled
## Export Registered Authentication Methods from https://portal.azure.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/UserRegistrationDetails
$AuthMethodsRegUsers = Import-Csv ~\exportUserRegistrationDetails_2023-3-22.csv"

$WorkSheetName = "UserMailboxes"
$waveGroup = Import-Excel -WorksheetName $WorkSheetName -Path ~\1777-OldCompany-OVG-MailboxList.xlsx"
$tenant = "OVG"
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
foreach ($user in $waveGroup) {
    Write-ProgressHelper -Activity "Gathering Details for $($user.Name) in $($tenant) tenant" -ProgressCounter ($progresscounter++) -TotalCount ($waveGroup).count
    $Usercheck = @()
    $EmailAddress = $null
    if ($null -ne $user.$tenant) {
        $EmailAddress = $user.$tenant.Trim()
    }
    else {
        $EmailAddress = $null
    }

    if ($Usercheck = $AuthMethodsRegUsers | ? {$_.userPrincipalName -eq $EmailAddress}) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $user | add-member -type noteproperty -name "SSPR Registered" -Value $Usercheck.ssprRegistered -Force
        $user | add-member -type noteproperty -name "SSPR Enabled" -Value $Usercheck.ssprEnabled -Force
        $user | add-member -type noteproperty -name "MFA Registered" -Value $Usercheck.mfaCapable -Force
        $user | add-member -type noteproperty -name "Methods Registered" -Value $Usercheck.methodsRegistered -Force
        $user | add-member -type noteproperty -name "defaultMfaMethod" -Value $Usercheck.defaultMfaMethod -Force

    }
    else {
        Write-Host "." -ForegroundColor red -NoNewline
        $user | add-member -type noteproperty -name "SSPR Registered" -Value $null -Force
        $user | add-member -type noteproperty -name "SSPR Enabled" -Value $null -Force
        $user | add-member -type noteproperty -name "MFA Registered" -Value $null -Force
        $user | add-member -type noteproperty -name "Methods Registered" -Value $null -Force
        $user | add-member -type noteproperty -name "defaultMfaMethod" -Value $null -Force
    }  
}

$waveGroup | Export-Excel -WorksheetName "OVG-MFACheck" -Path ~\1777-OldCompany-OVG-MailboxList.xlsx" -Show -ClearSheet
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan


# Get all users in the tenant with errors and include the error message for each user
$errorUsers = Get-MsolUser -HasErrorsOnly -all
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null

$AllErrorUsersDetails = @()
foreach ($user in $errorUsers) {
    Write-ProgressHelper -Activity "Gathering Error Details for $($user.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount ($errorUsers).count

    $currenterrorDetail = new-object PSObject
    if ($null -ne $user.Errors) {
         # Initialize an empty string for the error descriptions
         $errorDescriptions = ""

         # Loop through each error record
         foreach ($errorRecord in $user.Errors.ErrorDetail.ObjectErrors.ErrorRecord) {
             # Extract error description
             if ($errorRecord.ErrorDescription."#text") {
                 $errorDescription = $errorRecord.ErrorDescription."#text"
             }
             else {
                 $errorDescription = $errorRecord.ErrorDescription
             }
 
             # Add the error description to the error descriptions string
             if ($errorDescriptions -ne "") {
                 $errorDescriptions += "; "
             }
             $errorDescriptions += $errorDescription
         }

        $currenterrorDetail | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName -Force
        $currenterrorDetail | add-member -type noteproperty -name "UserPrincipalName" -Value $user.UserPrincipalName -Force
        $currenterrorDetail | add-member -type noteproperty -name "ErrorMessage" -Value $errorDescriptions -Force
    }
    else {
        $currenterrorDetail | add-member -type noteproperty -name "DisplayName" -Value $null -Force
        $currenterrorDetail | add-member -type noteproperty -name "UserPrincipalName" -Value $null -Force
        $currenterrorDetail | add-member -type noteproperty -name "ErrorMessage" -Value $null -Force
    }

    $AllErrorUsersDetails += $currenterrorDetail 
}

$AllErrorUsersDetails | Export-Excel -WorksheetName "OVG-ErrorCodes" -Path "C:\Users\amedrano\Arraya Solutions\OldCompany - External - Ext - 1777 OldCompany to OVG T2\UsersWithErrors.xlsx" -Show -ClearSheet
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan

 
# write a script to pull all the permissions of the D: drive and export to a csv file. Pull the first 3 levels of folders and subfolders. 
$global:start = Get-Date
$Folder = "D:\"
$Folder | Get-ChildItem -Recurse -Depth 3 | Get-Acl | Select-Object -Property Path, Access, IdentityReference | Export-Csv ~\1777-OldCompany-OVG-Permissions.csv" -NoTypeInformation
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan


#Match OneDrive Users to Mailboxes
$OVGOneDriveReport = Import-Csv ~\OVGOneDriveActivityUserDetail4_17_2023 5_48_48 PM.csv"
$OldCompanyOneDriveReport = Import-Csv ~\OldCompanyOneDriveActivityUserDetail4_17_2023 5_45_24 PM.csv"
$AllMatchedMailboxes = Import-Excel "~\1777-OldCompany-OVG-AllMatched-Mailboxes.xlsx" -WorksheetName "UserMailboxes"

# Initialize an empty array to store user hash tables
$FullOneDriveReportArray = @()

#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = ($AllMatchedMailboxes).count

foreach ($user in $AllMatchedMailboxes) {
    Write-ProgressHelper -Activity "Gathering OneDrive Details for $($user.DisplayName_OVG)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount
    $OVGUserDetails = $OVGOneDriveReport | ? {$_."User Principal Name" -eq $user.UserPrincipalName_OVG}
    $OldCompanyUserDetails = $OldCompanyOneDriveReport | ? {$_."User Principal Name" -eq $user.UserPrincipalName_OldCompany}

    # Create Hashtable for each user
    $FullOneDriveReport = @{
        DisplayName = $user.DisplayName_OVG
        UserPrincipalName_OVG = $user.UserPrincipalName_OVG
        IsDeleted_OVG = $OVGUserDetails."Is Deleted"
        LastActivityDate_OVG = $OVGUserDetails."Last Activity Date"
        ViewedOrEditedFileCount_OVG = $OVGUserDetails."Viewed Or Edited File Count"
        SyncedFileCount_OVG = $OVGUserDetails."Synced File Count"
        UserPrincipalName_OldCompany = $user.UserPrincipalName_OldCompany
        IsDeleted_OldCompany = $OldCompanyUserDetails."Is Deleted"
        LastActivityDate_OldCompany = $OldCompanyUserDetails."Last Activity Date"
        ViewedOrEditedFileCount_OldCompany = $OldCompanyUserDetails."Viewed Or Edited File Count"
        SyncedFileCount_OldCompany = $OldCompanyUserDetails."Synced File Count"
    }

    # Add the user's hash table to the array
    $FullOneDriveReportArray += $FullOneDriveReport
}
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan

#### HASH TABLE ONEDRIVE USERS
# Initialize an empty array to store user report data
#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = ($AllMatchedMailboxes).count
$FullOneDriveReportArray2 = @()

#Initialize Hash Tables to store OldCompany and OVG user data with UserPrincipalName as KEY and an array of corresponding user data as associated VALUE
$OVGUsersHash = @{}
$OldCompanyUsersHash = @{}
 
Write-Host ""
Write-Host "Processing imported data to add to hash table..."

#Process OVG imported data to populate $OVGUsersHash
$OVGOneDriveReport | foreach{
    #hash KEY
    $upn = $_.'User Principal Name'

    #values to add to array to be assigned as hash VALUE - will also include UPN in array
    $isDeleted = $_.'Is Deleted'
    $lastActivity = $_.'Last Activity Date'
    $viewedOrEditedFileCount = $_.'Viewed or Edited File Count'
    $syncedFileCount = $_.'Synced File Count'

    #initialize PSObject to be used as hash value
    $currentUser = new-object PSObject
    $currentUser | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $upn -Force
    $currentUser | Add-Member -type NoteProperty -Name "IsDeleted_OVG" -Value $isDeleted -Force
    $currentUser | Add-Member -type NoteProperty -Name "LastActivityDate_OVG" -Value $lastActivity -Force
    $currentUser | Add-Member -type NoteProperty -Name "ViewedOrEditedFileCount_OVG" -Value $viewedOrEditedFileCount -Force
    $currentUser | Add-Member -type NoteProperty -Name "SyncedFileCount_OVG" -Value $syncedFileCount -Force

    #Add to OVG Hash with UPN as HASH KEY and current user data object as HASH VALUE
    $OVGUsersHash.Add($upn, $currentUser)
}
Write-Host "Processing imported data to add to secondary hash table..."

#Process OldCompany imported data to fill associated hash table $OldCompanyUsersHash
$OldCompanyOneDriveReport | foreach{
    #hash KEY
    $upn = $_.'User Principal Name'

    #values to add to array to be assigned as hash VALUE - will also include UPN in array
    $isDeleted = $_.'Is Deleted'
    $lastActivity = $_.'Last Activity Date'
    $viewedOrEditedFileCount = $_.'Viewed or Edited File Count'
    $syncedFileCount = $_.'Synced File Count'

    #initialize PSObject to be used as hash value
    $currentUser = new-object PSObject
    $currentUser | Add-Member -type NoteProperty -Name "UserPrincipalName_OldCompany" -Value $upn -Force
    $currentUser | Add-Member -type NoteProperty -Name "IsDeleted_OldCompany" -Value $isDeleted -Force
    $currentUser | Add-Member -type NoteProperty -Name "LastActivityDate_OldCompany" -Value $lastActivity -Force
    $currentUser | Add-Member -type NoteProperty -Name "ViewedOrEditedFileCount_OldCompany" -Value $viewedOrEditedFileCount -Force
    $currentUser | Add-Member -type NoteProperty -Name "SyncedFileCount_OldCompany" -Value $syncedFileCount -Force

    #Add to OVG Hash with UPN as HASH KEY and current user data object as HASH VALUE
    $OldCompanyUsersHash.Add($upn, $currentUser)
}

foreach ($user in $AllMatchedMailboxes) {
    Write-ProgressHelper -Activity "Gathering OneDrive Details for $($user.DisplayName_OVG)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount
    #$OVGUserDetails = $OVGOneDriveReport | ? {$_."User Principal Name" -eq $user.UserPrincipalName_OVG}
    #$OldCompanyUserDetails = $OldCompanyOneDriveReport | ? {$_."User Principal Name" -eq $user.UserPrincipalName_OldCompany}

    if($user.UserPrincipalName_OVG)
    {
        $ovgUPN = $user.UserPrincipalName_OVG.toString()
    }

    if($user.UserPrincipalName_OldCompany)
    {
        $OldCompanyUPN = $user.UserPrincipalName_OldCompany.toString()
    }

    #obtain data set from hash table associated with OVG and OldCompany UPNs
    $OVGUserDetails = $OVGUsersHash[$ovgUPN]
    $OldCompanyUserDetails = $OldCompanyUsersHash[$OldCompanyUPN]

    $currentUser = new-object PSObject
    $currentUser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName_OVG -Force
    $currentUser | Add-Member -type NoteProperty -Name "UserPrincipalName_OVG" -Value $user.UserPrincipalName_OVG -Force
    $currentUser | Add-Member -type NoteProperty -Name "IsDeleted_OVG" -Value $OVGUserDetails.IsDeleted_OVG -Force
    $currentUser | Add-Member -type NoteProperty -Name "LastActivityDate_OVG" -Value $OVGUserDetails.LastActivityDate_OVG -Force
    $currentUser | Add-Member -type NoteProperty -Name "ViewedOrEditedFileCount_OVG" -Value $OVGUserDetails.ViewedOrEditedFileCount_OVG -Force
    $currentUser | Add-Member -type NoteProperty -Name "SyncedFileCount_OVG" -Value $OVGUserDetails.SyncedFileCount_OVG -Force
    $currentUser | Add-Member -type NoteProperty -Name "UserPrincipalName_OldCompany" -Value $user.UserPrincipalName_OldCompany -Force
    $currentUser | Add-Member -type NoteProperty -Name "IsDeleted_OldCompany" -Value $OldCompanyUserDetails.IsDeleted_OldCompany -Force
    $currentUser | Add-Member -type NoteProperty -Name "LastActivityDate_OldCompany" -Value $OldCompanyUserDetails.LastActivityDate_OldCompany -Force
    $currentUser | Add-Member -type NoteProperty -Name "ViewedOrEditedFileCount_OldCompany" -Value $OldCompanyUserDetails.ViewedOrEditedFileCount_OldCompany -Force
    $currentUser | Add-Member -type NoteProperty -Name "SyncedFileCount_OldCompany" -Value $OldCompanyUserDetails.SyncedFileCount_OldCompany -Force
 
    # Add the users to the array
    $FullOneDriveReportArray2 += $currentUser
}

Write-Host "Count of items in FullOneDriveReportArray: $($FullOneDriveReportArray2.Count)"
Write-Host""
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan