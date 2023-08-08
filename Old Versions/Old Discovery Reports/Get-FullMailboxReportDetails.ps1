<#
    Original Mailbox Report Details Script
    Uses PSObject to create a custom object array
    Includes SharePoint Details
    Includes Archive Mailbox Details
    Includes SendAs Permissions
    Includes Inactive Mailboxes
    Includes Group Mailboxes
    Includes Shared Mailboxes
    Does not include Discovery Search Mailboxes
    includes a progress bar
    Includes a timer
    Includes a try/catch to handle errors
    includes a do/while loop to handle errors
    Includes a function to handle the progress bar

    Against Spectra tenant took 5:47:12.  5 hours, 47 minutes, 12 seconds (using ExchangeOnlineManagement module V2 commands) 
    -- 4143 mailboxes
#>


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

function Test-SPOServiceConnection {
    param ()
    # Check if SharePoint Online module is installed
    if (Get-Command Connect-SPOService -ErrorAction SilentlyContinue) {
        Write-Host "SharePoint Online module is installed."
        try {
            # Check if connected to a SharePoint Online site
            $spoConnection = Get-SPOSite -Limit 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            if ($spoConnection -ne $null) {
                Write-Host "Connected to SharePoint Online site"
                return $true

            } else {
                Write-Host "Not connected to any SharePoint Online site."
            }
        } catch {
            Write-Host "Error checking SharePoint Online connection: $($_.Exception.Message)"
        }
    } else {
        Write-Host "SharePoint Online module is not installed."
    }
}

function Import-ExchangeAndMsOnlineModules {
    if (((Get-Module -Name "ExchangeOnlineManagement") -ne $null) -or ((Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue) -ne $null)) {
        return;
    }
    else {
        Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 2 can be found at https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"

    }
    if (((Get-Module -Name "MSOnline") -ne $null) -or ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null)) {
        return;
    }
    else {
        Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator"
    }

    $EXOmoduleLocation -eq "$env:ProgramFiles\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.5\netFramework\ExchangeOnlineManagement.psm1"
            if (Test-Path $EXOmoduleLocation)  {
                Import-Module -Name ExchangeOnlineManagement
                return
            }

    $MsOnlinemoduleLocation -eq "$env:ProgramFiles\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.5\netFramework\ExchangeOnlineManagement.psm1"
            if (Test-Path $MsOnlinemoduleLocation)  {
                Import-Module -Name MSOnline
                return
            }
}
$spoConnection = Test-SPOServiceConnection
Import-ExchangeAndMsOnlineModules

$global:start = Get-Date
# Gather Mailbox Stats - Include SharePoint
$sourceMailboxes = Get-Mailbox -ResultSize Unlimited -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -includeInactiveMailbox
$sourceMailboxes += Get-Mailbox -ResultSize Unlimited -GroupMailbox

$sourceMailboxStats = @()
$tenant = "Spectra"

#ProgressBar
$progresscounter = 1
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = ($sourceMailboxes).count
$ProgressPreference = "Continue"

foreach ($user in $sourceMailboxes) {
    #progress bar
    Write-ProgressHelper -Activity "Gathering Mailbox Details for $($user.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -ID 1
    
    #Pull MailboxStats and UserDetails
    $mbxStats = Get-MailboxStatistics $user.PrimarySMTPAddress
    $msoluser = Get-MsolUser -UserPrincipalName $user.UserPrincipalName
    $EmailAddresses = $user | select -ExpandProperty EmailAddresses

    #Create User Output Array
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName_$($tenant)" -Value $msoluser.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName_$($tenant)" -Value $msoluser.userprincipalname
    $currentuser | add-member -type noteproperty -name "IsLicensed_$($tenant)" -Value $msoluser.IsLicensed
    $currentuser | add-member -type noteproperty -name "Licenses_$($tenant)" -Value ($msoluser.Licenses.AccountSkuID -join ";") -force
    $currentuser | add-member -type noteproperty -name "License-DisabledArray_$($tenant)" -Value ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";") -force
    $currentuser | add-member -type noteproperty -name "BlockCredential_$($tenant)" -Value $msoluser.BlockCredential
    $currentuser | add-member -type noteproperty -name "Department_$($tenant)" -Value $msoluser.Department    
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails_$($tenant)" -Value $user.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress_$($tenant)" -Value $user.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "IsInactiveMailbox_$($tenant)" -Value $user.IsInactiveMailbox
    $currentuser | add-member -type noteproperty -name "Alias_$($tenant)" -Value $user.alias
    $currentuser | add-member -type noteproperty -name "WhenCreated_$($tenant)" -Value $user.WhenCreated
    $currentuser | add-member -type noteproperty -name "WhenSoftDeleted_$($tenant)" -Value $user.WhenSoftDeleted
    $currentuser | add-member -type noteproperty -name "LastLogonTime_$($tenant)" -Value $mbxStats.LastLogonTime
    $currentuser | add-member -type noteproperty -name "EmailAddresses_$($tenant)" -Value ($EmailAddresses -join ";")
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN_$($tenant)" -Value ("x500:" + $user.legacyexchangedn)
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled_$($tenant)" -Value $user.HiddenFromAddressListsEnabled
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward_$($tenant)" -Value $user.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ForwardingAddress_$($tenant)" -Value $user.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress_$($tenant)" -Value $user.ForwardingSmtpAddress
    $currentuser | Add-Member -type NoteProperty -Name "MBXSize_$($tenant)" -Value $MBXStats.TotalItemSize
    $currentuser | Add-Member -type NoteProperty -Name "TotalItemSize-MB_$($tenant)" -Value ([math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1MB,3)) -Force
    $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount_$($tenant)" -Value $MBXStats.ItemCount
    $currentuser | add-member -type noteproperty -name "PublicFolderMailbox_$($tenant)" -Value $user.EffectivePublicFolderMailbox -force

    #Pull Send on Behalf
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo_$($tenant)" -Value ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.primarysmtpaddress | ?{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms_$($tenant)" -Value ($mbxPermissions.user -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "FullAccessPerms_$($tenant)" -Value $null
    }
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.PrimarySMTPAddress | ?{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser | add-member -type noteproperty -name "SendAsPerms_$($tenant)" -Value ($sendAsPermsCheck.trustee -join ";") -Force
    }
    else {
        $currentuser | add-member -type noteproperty -name "SendAsPerms_$($tenant)" -Value $null
    }
    # Archive Mailbox & Retention Policy Check
    $currentuser | Add-Member -Type NoteProperty -Name "RetentionPolicy_$($tenant)" -Value $user.RetentionPolicy
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus_$($tenant)" -Value $user.ArchiveStatus
    if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount) {           
        $currentuser | add-member -type noteproperty -name "ArchiveSize_$($tenant)" -Value $ArchiveStats.TotalItemSize.Value
        $currentuser | Add-Member -type NoteProperty -Name "ArchiveSize_-MB_$($tenant)" -Value ([math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1MB,3)) -Force
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount_$($tenant)" -Value $ArchiveStats.ItemCount
    }
    else {
        $currentuser | add-member -type noteproperty -name "ArchiveSize_$($tenant)" -Value $null
        $currentuser | Add-Member -type NoteProperty -Name "ArchiveSize_-MB_$($tenant)" -Value $null -Force
        $currentuser | add-member -type noteproperty -name "ArchiveItemCount_$($tenant)" -Value $null
    }
    # Gather SharePoint Details
    if ($user.IsInactiveMailbox -eq $false) {
        $count = 0
        $success = $null

        do{
            try{
                $OneDriveSite = Get-SPOSite -Filter "Owner -eq '$($msoluser.UserPrincipalName)' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop | select Url, StorageUsageCurrent
                if ($OneDriveSite) {
                    $currentuser | add-member -type noteproperty -name "OneDriveURL_$($tenant)" -Value $OneDriveSite.url -force
                    $currentuser | add-member -type noteproperty -name "OneDriveStorage-MB_$($tenant)" -Value $OneDriveSite.StorageUsageCurrent -force
                    $success = $true
                }
            else {
                    $currentuser | add-member -type noteproperty -name "OneDriveURL_$($tenant)" -Value $null -force
                    $currentuser | add-member -type noteproperty -name "OneDriveStorage-MB_$($tenant)" -Value $null -force
                    $failed = $true
            }
            }
            catch{
                Write-host "Next attempt in 30 seconds" -foregroundcolor yellow -nonewline
                Start-sleep -Seconds 30
                $count++
            # Put the start-sleep in the catch statement
            # don't sleep if the condition is true and waste time
            }
        }
        until($count -eq 5 -or $success -or $failed)

        if(!($success -or $failed)) {
            $currentuser | add-member -type noteproperty -name "OneDriveURL_$($tenant)" -Value $null -force
            Write-Host ". " -foregroundcolor red -nonewline
        }
    }
    else {
        $currentuser | add-member -type noteproperty -name "OneDriveURL_$($tenant)" -Value $null -force
        $currentuser | add-member -type noteproperty -name "OneDriveStorage-MB_$($tenant)" -Value $null -force
    }

    #Combine all the data into one object
    $sourceMailboxStats += $currentuser
}
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan

$sourceMailboxStats | Export-Csv ~\Desktop\$($tenant)-AllSourceMailboxStatsv1.csv -NoTypeInformation

