<#
    Testing code using HashTables for the output. Creates a hashtable for each user and then adds it to the $allMailboxStats hashtable.
    After all users are processed, the $allMailboxStats hashtable is converted to a CSV file.
    This is a test to see if this is faster than using a PSCustomObject.
    #This also uses the functions in the MyModule.psm1 file.

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
    Includes a function to check if the user is connected to SharePoint Online
    Includes a function to check if the SharePoint Online module is installed

    Against Spectra Tenant took 8:18:23 to complete (using ExchangeOnlineManagement module V2 commands)
    -- 4143 mailboxes

    #removed wait time for OneDrive Checks #added Import-RequiredModules function

#>
$start = Get-Date
function Test-SPOServiceConnection {
    param ()
    # Check if SharePoint Online module is installed
    if (Get-Command Connect-SPOService -ErrorAction SilentlyContinue) {
        Write-Host "SharePoint Online module is installed."
        try {
            # Check if connected to a SharePoint Online site
            $spoConnection = Get-SPOSite -Limit 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            if ($null -ne $spoConnection) {
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

function Import-RequiredModules {
    #ExchangeOnlineManagement module V3 Installation Check
    if (($null -ne (Get-Module -Name "ExchangeOnlineManagement")) -or ($null -ne (Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue))) {
        return;
    }
    else {
        Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-the-exchange-online-powershell-module"

    }
    #MSOnline module Installation Check
    if (($null -ne (Get-Module -Name "MSOnline")) -or ($null -ne (Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue))) {
        return;
    }
    else {
        Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator"
    }

    #Microsoft.Online.SharePoint module Installation Check
    if (($null -ne (Get-Module -Name "Microsoft.Online.SharePoint.PowerShell")) -or ($null -ne (Get-InstalledModule -Name "Microsoft.Online.SharePoint.PowerShell" -ErrorAction SilentlyContinue))) {
        return;
    }
    else {
        Write-Error  "SharePoint Online module was not loaded. Run Install-Module Microsoft.Online.SharePoint.PowerShell as an Administrator"
    }

    #ExchangeOnlineManagement module V3 Import
    $EXOmoduleLocation -eq "$env:ProgramFiles\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.5\netFramework\ExchangeOnlineManagement.psm1"
    if (Test-Path $EXOmoduleLocation)  {
        Import-Module -Name ExchangeOnlineManagement
        return
    }
    #MSOnline module Import
    $MsOnlinemoduleLocation -eq "$env:ProgramFiles\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.5\netFramework\ExchangeOnlineManagement.psm1"
    if (Test-Path $MsOnlinemoduleLocation)  {
        Import-Module -Name MSOnline
        return
    }
    #Microsoft.Online.SharePoint module Import
    $SPOmoduleLocation = "$env:ProgramFiles\WindowsPowerShell\Modules\Microsoft.Online.SharePoint.PowerShell\16.0.21908.12000\Microsoft.Online.SharePoint.PowerShell.psm1"
    if (Test-Path $SPOmoduleLocation)  {
        Import-Module -Name Microsoft.Online.SharePoint.PowerShell
        return
    }
}

$spoConnection = Test-SPOServiceConnection
Import-RequiredModules

# Gather Mailbox Stats - Include InActive Mailboxes and Group Mailboxes
$allMailboxes = Get-Mailbox -ResultSize Unlimited -Filter "PrimarySMTPAddress -notlike '*DiscoverySearchMailbox*'" -includeInactiveMailbox
$allMailboxes += Get-Mailbox -ResultSize Unlimited -GroupMailbox

$allMailboxStats = [Ordered]@{}
$tenant = "Spectra"

#ProgressBar
$progresscounter = 1
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = ($allMailboxes).count
$ProgressPreference = "Continue"
foreach ($user in $allMailboxes) {
    #progress bar
    Write-ProgressHelper -Activity "Gathering Mailbox Details for $($user.DisplayName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -ID 1 -StartTime $start
    
    #Pull MailboxStats and UserDetails
    $mbxStats = Get-MailboxStatistics $user.DistinguishedName
    $msoluser = Get-MsolUser -UserPrincipalName $user.UserPrincipalName
    $EmailAddresses = $user | Select-Object -ExpandProperty EmailAddresses

    # Create User Hash Table
    $currentuser = [ordered]@{
        "DisplayName" = $msoluser.DisplayName
        "UserPrincipalName" = $msoluser.userprincipalname
        "IsLicensed" = $msoluser.IsLicensed
        "Licenses" = ($msoluser.Licenses.AccountSkuID -join ";")
        "License-DisabledArray" = ($msoluser.LicenseAssignmentDetails.Assignments.DisabledServicePlans -join ";")
        "BlockCredential" = $msoluser.BlockCredential
        "IsInactiveMailbox" = $user.IsInactiveMailbox
        "Department" = $msoluser.Department
        "RecipientTypeDetails" = $user.RecipientTypeDetails
        "PrimarySmtpAddress" = $user.PrimarySmtpAddress
        "Alias" = $user.alias
        "WhenCreated" = $user.WhenCreated
        "LastLogonTime" = $mbxStats.LastLogonTime
        "EmailAddresses" = ($EmailAddresses -join ";")
        "LegacyExchangeDN" = ("x500:" + $user.legacyexchangedn)
        "HiddenFromAddressListsEnabled" = $user.HiddenFromAddressListsEnabled
        "DeliverToMailboxAndForward" = $user.DeliverToMailboxAndForward
        "ForwardingAddress" = $user.ForwardingAddress
        "ForwardingSmtpAddress" = $user.ForwardingSmtpAddress
        "MBXSize" = $MBXStats.TotalItemSize
        "MBXSize_GB" = [math]::Round(($MBXStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
        "MBXItemCount" = $MBXStats.ItemCount
        "LitigationHoldEnabled" = $user.LitigationHoldEnabled
        "LitigationHoldDuration" = $user.LitigationHoldDuration
        "InPlaceHolds" = $user.InPlaceHolds -join ";"
        "ArchiveStatus" = $user.ArchiveStatus
        "RetentionPolicy" = $user.RetentionPolicy
    }

    #Pull Send on Behalf
    $grantSendOnBehalf = $user.GrantSendOnBehalfTo
    $grantSendOnBehalfPerms = @()
    foreach ($perm in $grantSendOnBehalf) {
        $mailboxCheck = (Get-Mailbox $perm).DisplayName
        $grantSendOnBehalfPerms += $mailboxCheck
    }
    $currentuser["GrantSendOnBehalfTo"] = ($grantSendOnBehalfPerms -join ";")

    # Mailbox Full Access Check
    if ($mbxPermissions = Get-MailboxPermission $user.DistinguishedName | Where-Object{$_.user -ne "NT AUTHORITY\SELF" -and $_.User -notlike "*NAMPR0*" -and $_.User -notlike "S-1-5-*"}) {
        $currentuser["FullAccessPerms"] = ($mbxPermissions.User -join ";")
    }
    else {$currentuser["FullAccessPerms"] = ($null)}
    # Mailbox Send As Check
    if ($sendAsPermsCheck = Get-RecipientPermission -AccessRights SendAs -Identity $user.DistinguishedName | Where-Object{$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -notlike "*NAMPR0*" -and $_.Trustee -notlike "S-1-5-*"}) {
        $currentuser["SendAsPerms"] = ($sendAsPermsCheck.trustee -join ";")
    }
    else {$currentuser["SendAsPerms"] = ($null)}
    # Archive Mailbox Check
    $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus
    if ($ArchiveStats = Get-MailboxStatistics $user.DistinguishedName -Archive -ErrorAction silentlycontinue | Select-Object TotalItemSize, ItemCount) {           
        $currentuser["ArchiveSize"] = $ArchiveStats.TotalItemSize.Value
        $currentuser["ArchiveSize-GB"] = [math]::Round(($ArchiveStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1GB,3)
    }
    else {
        $currentuser["ArchiveSize"] = $null
        $currentuser["ArchiveSize-GB"] = $null
        $currentuser["ArchiveItemCount"] = $null
    }

    # Gather SharePoint Details
    if ($spoConnection -and ($user.IsInactiveMailbox -eq $false)) {
        $OneDriveSite = $null
        $OneDriveSite = Get-SPOSite -Filter "Owner -eq '$($msoluser.UserPrincipalName)' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -Limit All -ErrorAction Stop | Select-Object Url, StorageUsageCurrent,LastContentModifiedDate
        if ($OneDriveSite) {
            $currentuser["OneDriveURL_$($tenant)"] = $OneDriveSite.Url
            $currentuser["OneDriveStorage-GB_$($tenant)"] = [math]::Round($OneDriveSite.StorageUsageCurrent / 1024, 3)
            $currentuser["OneDriveLastContentModifiedDate"] = $OneDriveSite.LastContentModifiedDate
        } else {
            $currentuser["OneDriveURL_$($tenant)"] = $null
            $currentuser["OneDriveStorage-GB_$($tenant)"] = $null
            $currentuser["OneDriveLastContentModifiedDate"] = $null
        }
    } else {
        $currentuser["OneDriveURL_$($tenant)"] = $null
        $currentuser["OneDriveStorage-GB_$($tenant)"] = $null
        $currentuser["OneDriveLastContentModifiedDate"] = $null
    }

    #Combine all the data into one hash table
    $allMailboxStats[$User.PrimarySMTPAddress] = $currentuser
}

#Convert Hash Table to Custom Object Array for Export
$ExportAllMailboxStatsArray = @()
foreach ($key in $allMailboxStats.Keys) {
    $attributes = $allMailboxStats[$key]
    $customObject = New-Object -TypeName PSObject

    foreach ($attribute in $attributes.Keys) {
        $customObject | Add-Member -MemberType NoteProperty -Name "$attribute__$tenant" -Value $attributes[$attribute]
    }

    $ExportAllMailboxStatsArray += $customObject
}
Write-Host (Get-ElapsedTime -StartTime $start) -ForegroundColor Cyan
Write-Host "Completed in"((Get-Date) - $start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan

$ExportAllMailboxStatsArray | Export-Csv ~\Desktop\$($tenant)-AllMailboxStatsv2.csv -NoTypeInformation
