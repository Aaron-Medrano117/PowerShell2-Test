<# .SYNOPSIS
    This script can be used to pull delegatge permissions based on Calendar Permissions granted per mailbox. 
    Pulls unique values and exports lists of users sharing calendars and those not sharing calendars.

    Full List of Permissions are exported as "DelegatePermissions.csv". By Default exports to desktop.

    .PARAMETER OutputCSVFilePath
    Output File Path for Report. Can specify exactly where to save file and what to name it.
    .PARAMETER OutputCSVFolderPath
    Output Folder Path for Report. Designate where to save file as 'DelegatePermissions.csv"
    .PARAMETER OnPremises
    Switch used to specify if running On-Premises Exchange. Should support versions Exchange 2010 through Exchange 2016
    .PARAMETER Office365
    Switch used to specify if running against Office 365's Exchange Online.
    .PARAMETER CalendarPermissions
    Switch used to request Calendar Permissions. Exports all calendar folders of mailbox and their permissions
    .PARAMETER SendAs
    Switch used to request Send As Permissions for the mailboxes
    .PARAMETER FullAccess
    Switch used to request Full Access Permissions for the mailboxes
    .PARAMETER SendOnBehalf
    Switch used to request Send On Behalf Permissions for the mailboxes

   .EXAMPLE
   Pulls all available permissions for each mailbox in Office 365. Exports to default location of desktop.
   .\Add-DelegatePermissions.ps1 -Office365 -CalendarPermissions -FullAccess -SendAs -SendOnBehalf
   .EXAMPLE
   Pulls only Calendar Permissions for each mailbox in Office 365. Exports to user's documents folder.
   .\Add-DelegatePermissions.ps1 -Office365 -CalendarPermissions -OutputCSVFolderPath C:\user\documents
   .EXAMPLE
   Pulls all available permissions for each mailbox in On-Premises Exchange. Exports to default location of desktop.
   .\Add-DelegatePermissions.ps1 -OnPremises -CalendarPermissions -FullAccess -SendAs
   .EXAMPLE
   Pulls Full Access and Send As permissions for each mailbox in On-Premises Exchange. Exports to user's documents folder.
   .\Add-DelegatePermissions.ps1 -OnPremises -FullAccess -SendaAs -OutputCSVFolderPath C:\user\documents
#>

param (
    [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
    [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
    [Parameter(Mandatory=$false,HelpMessage='Matched Mailbox List')] [string] $matchedMailboxes,
    [Parameter(Mandatory=$false,HelpMessage="Run against OnPremises Exchange?")] [switch]$OnPremises,
    [Parameter(Mandatory=$True)] [string]$DestinationTenant,
    [Parameter(Mandatory=$True)] [string]$SourceTenant,
    [Parameter(Mandatory=$false,HelpMessage="Run against Office365 Exchange Online?")] [switch]$Office365,
    [Parameter(Mandatory=$false,HelpMessage="Add Calendar Permissions?")] [switch]$CalendarPermissions,
    [Parameter(Mandatory=$false,HelpMessage="Add Mailbox Perms?")] [switch]$MailboxPerms,
    [Parameter(Mandatory=$false,HelpMessage="Add Full Access Permissions?")] [switch]$FullAccess,
    [Parameter(Mandatory=$false,HelpMessage="Add Send OnBehalf Permissions?")] [switch]$SendOnBehalf,
    [Parameter(,Mandatory=$false,HelpMessage="Add Send-As Permissions?")] [switch]$SendAs
)

function Write-ProgressHelper {
    param (
        [int]$ProgressCounter,
        [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [int]$TotalCount,
        [datetime]$StartTime
    )
    #$ProgressPreference = "Continue"  
    if ($ProgressPreference = "SilentlyContinue") {
        $ProgressPreference = "Continue"
    }

    $secondsElapsed = (Get-Date) - $StartTime
    $secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)
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
    #$secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)

}

#Gather Exported Permissions
if ($InputCSVFilePath) {
    $allDelegatePermissions = Import-CSV $InputCSVFilePath
}
elseif ($InputEXCELFilePath) {
    $allDelegatePermissions = Import-Excel $InputEXCELFilePath
}

#Gather Matched Mailboxes Permissions
if ($InputCSVFilePath) {
    $matchedMailboxes = Import-CSV $matchedMailboxes
    $activeMatchedMailboxes = $matchedMailboxes | ?{$_."IsInactiveMailbox_$($SourceTenant)" -eq $false}
}
elseif ($InputEXCELFilePath) {
    $matchedMailboxes = Import-Excel $matchedMailboxes
    $activeMatchedMailboxes = $matchedMailboxes | ?{$_."IsInactiveMailbox_$($SourceTenant)" -eq $false}
}

# Populate mailbox mapping
$MailboxMapping = @{}
foreach ($matchedMailbox in $activeMatchedMailboxes) {
    if ($matchedMailbox."PrimarySMTPAddress_$($DestinationTenant)") {
        $sourceAddress = $matchedMailbox."PrimarySMTPAddress_$($SourceTenant)"
        $destinationAddress = $matchedMailbox."PrimarySMTPAddress_$($DestinationTenant)"
        $MailboxMapping[$sourceAddress] = $destinationAddress
    }
}

function Get-GroupsDetailsHash {
    param (
        [Parameter(Mandatory=$True)] [string]$SourceTenant,
        [Parameter(Mandatory=$True)] [Array]$matchedGroupsDetails
    )
    # Populate mailbox mapping
    $DistributionGroupDetailsHash = @{}
    foreach ($group in $matchedGroupsDetails) {
        $sourceAddress = $matchedMailbox."PrimarySMTPAddress_$($SourceTenant)"
        $key = $group
        $DistributionGroupDetailsHash[$sourceAddress] = $key
    }
    return $DistributionGroupDetailsHash    
}
#REGION Add list of Delegate Permissions
#ProgressBar
$progresscounter = 1
$global:start = Get-Date

#Build Array
$AllPermissionErrors = @()
$completedPerms = @()
$notFoundUsers = @()
$permsAlreadyExist = @()

foreach ($obj in $allDelegatePermissions) {
    #Check MailObject and PermUser match in Mailbox Mapping
    if ($MailboxMapping.ContainsKey($obj.MailObject)) {
        $MatchedObjectDestinationAddress = $MailboxMapping[$obj.MailObject].ToLower()
    }
    if ($mailboxMapping.ContainsKey($obj.PermUser)) {
        $MatchedPermUserDestinationAddress = $MailboxMapping[$obj.PermUser].ToLower()
    }
    if ($MatchedObjectDestinationAddres -and $MatchedPermUserDestinationAddress) {
        Write-ProgressHelper -ProgressCounter ($progresscounter++) -Activity "Adding Delegate Permissions for $($MatchedObjectDestinationAddres): $($obj.PermissionType)" -ID 1 -TotalCount ($allDelegatePermissions).count -StartTime $global:start
        #Check Mail Perm Object
        if ($mailObjectCheck = Get-EXORecipient $MatchedObjectDestinationAddres -ErrorAction SilentlyContinue) {
            Write-Host "$MatchedObjectDestinationAddres.. " -ForegroundColor Cyan -NoNewline
           
            #Check for Perm Object
            if ($permuser = Get-ExoRecipient $MatchedPermUserDestinationAddress -ErrorAction SilentlyContinue) {
                #Calendar Permissions
                if ($CalendarPermissions) {
                    if ($obj.PermissionType -eq "Calendar") {
                        try {
                            $calendarPath = $obj.CalendarPath -split ":"
                            $newCalendarPath = $MatchedObjectDestinationAddress + ":" + $calendarPath[1]
                            Write-Host "CalendarPermissions.." -NoNewline -foregroundcolor DarkCyan
                            $permAdd = Add-MailboxFolderPermission -identity $newCalendarPath -User $permuser.PrimarySMTPAddress -AccessRights $obj.AccessRights -confirm:$false -ErrorAction Stop
                            Write-Host "." -ForegroundColor Green -NoNewline
                            $completedPerms += $obj
                        }
                        catch {
                            Write-Host "." -ForegroundColor Red -NoNewline
                            $currenterror = new-object PSObject
                            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                            $currenterror | Add-Member -type NoteProperty -Name "MailObject" -Value $obj.MailObject -Force
                            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "CalendarPermission" -Force
                            $currenterror | Add-Member -type NoteProperty -Name "PermUser" -Value $permuser.PrimarySMTPAddress -Force
                            $currenterror | Add-Member -type NoteProperty -Name "CalendarPath" -Value $obj.CalendarPath -Force
                            $currenterror | Add-Member -type NoteProperty -Name "AccessRights" -Value $obj.AccessRights -Force
                            $currenterror | Add-Member -type NoteProperty -Name "SharingPermissionFlags" -Value $obj.SharingPermissionFlags -Force
                            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                            $AllPermissionErrors += $currenterror          
                            continue
                        }
                    }
                }
                #Set Full Access
                if ($FullAccess) {
                    if ($obj.PermissionType -eq "FullAccess") {
                        if (!(Get-EXOMailboxPermission -Identity $mailObjectCheck.PrimarySmtpAddress -User $permuser.DisplayName -errorAction SilentlyContinue)) {
                            try {
                                Write-Host "FullAccess.." -NoNewline -foregroundcolor DarkCyan
                                $permAdd = Add-MailboxPermission -Identity $mailObjectCheck.PrimarySMTPAddress -AccessRights FullAccess -User $permuser.PrimarySMTPAddress -confirm:$false -ErrorAction Stop
                                Write-Host "." -ForegroundColor Green -NoNewline
                                $completedPerms += $obj
                            }
                            catch {
                                Write-Host "." -ForegroundColor Red -NoNewline
                                $currenterror = new-object PSObject
                                $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                                $currenterror | Add-Member -type NoteProperty -Name "MailObject" -Value $obj.MailObject -Force
                                $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "FullAccess" -Force
                                $currenterror | Add-Member -type NoteProperty -Name "PermUser" -Value $permuser.PrimarySMTPAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "CalendarPath" -Value $obj.CalendarPath -Force
                                $currenterror | Add-Member -type NoteProperty -Name "AccessRights" -Value $obj.AccessRights -Force
                                $currenterror | Add-Member -type NoteProperty -Name "SharingPermissionFlags" -Value $obj.SharingPermissionFlags -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                                $AllPermissionErrors += $currenterror          
                                continue
                            }
                        }
                        else {
                            Write-Host " Already Has Full Access. " -ForegroundColor Yellow -nonewline
                            $permsAlreadyExist += $obj
                        }
                    }
                }
                #Set Send As Permission
                if ($SendAs) {
                    if ($obj.PermissionType -eq "SendAs"){
                        if (!(Get-EXORecipientPermission -Trustee $permuser.DisplayName -Identity $mailObjectCheck.PrimarySmtpAddress -AccessRights SendAs -errorAction SilentlyContinue)) {
                            try {
                                Write-Host "SendAs.." -NoNewline -foregroundcolor DarkCyan
                                $permAdd = Add-RecipientPermission -AccessRights SendAs -Identity $mailObjectCheck.PrimarySMTPAddress -Trustee $permuser.PrimarySMTPAddress -ErrorAction Stop
                                Write-Host "." -ForegroundColor Green -NoNewline
                                $completedPerms += $obj
                            }
                            catch {
                                Write-Host "." -ForegroundColor Red -NoNewline
                                $currenterror = new-object PSObject
                                $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                                $currenterror | Add-Member -type NoteProperty -Name "MailObject" -Value $obj.MailObject -Force
                                $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SendAs" -Force
                                $currenterror | Add-Member -type NoteProperty -Name "PermUser" -Value $permuser.PrimarySMTPAddress -Force
                                $currenterror | Add-Member -type NoteProperty -Name "CalendarPath" -Value $obj.CalendarPath -Force
                                $currenterror | Add-Member -type NoteProperty -Name "AccessRights" -Value $obj.AccessRights -Force
                                $currenterror | Add-Member -type NoteProperty -Name "SharingPermissionFlags" -Value $obj.SharingPermissionFlags -Force
                                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                                $AllPermissionErrors += $currenterror          
                                continue
                            }
                    }
                        else {
                            Write-Host " Already Has Send As. " -ForegroundColor Yellow -nonewline
                            $permsAlreadyExist += $obj
                        }
                    }
                }
                
                #Send On BehalfTO
                if ($SendOnBehalf) {
                    if ($obj.PermissionType -eq "SendOnBehalfTo") {
                        try {
                            Write-Host "SendOnBehalfTo.." -NoNewline -foregroundcolor DarkCyan
                            $permAdd = Set-Mailbox -Identity $mailObjectCheck.PrimarySMTPAddress -GrantSendOnBehalfTo @{Add=$permuser.PrimarySMTPAddress} -ErrorAction Stop
                            Write-Host "." -ForegroundColor Green -NoNewline
                            $completedPerms += $obj
                        }
                        catch {
                            Write-Host "." -ForegroundColor Red -NoNewline
                            $currenterror = new-object PSObject
                            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
                            $currenterror | Add-Member -type NoteProperty -Name "MailObject" -Value $obj.MailObject -Force
                            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SendOnBehalfTo" -Force
                            $currenterror | Add-Member -type NoteProperty -Name "PermUser" -Value $permuser.PrimarySMTPAddress -Force
                            $currenterror | Add-Member -type NoteProperty -Name "CalendarPath" -Value $obj.CalendarPath -Force
                            $currenterror | Add-Member -type NoteProperty -Name "AccessRights" -Value $obj.AccessRights -Force
                            $currenterror | Add-Member -type NoteProperty -Name "SharingPermissionFlags" -Value $obj.SharingPermissionFlags -Force
                            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                            $AllPermissionErrors += $currenterror          
                            continue
                        }         
                    }
                }
                
                Write-Host " Done. " -ForegroundColor Green
            }
            else {
                Write-Host "$($permMatchedDestinationAddress) perm user not Found. Skipping" -ForegroundColor Red
                $NotFoundUsers += $obj.PermUser
            }
        }
        else {
            Write-Host "$($matchedDestinationAddress) perm object not Found. Skipping" -ForegroundColor Red
            $NotFoundUsers += $obj.MailObject
        }
    }
    else {
        break
    }
}
#ENDREGION

Write-host ""
Write-Host "Completed in "(((Get-Date) - $global:initialstart).ToString('hh\:mm\:ss'))"" -ForegroundColor Green
Write-host $AllPermissionErrors.count "Errors occured. Check the $AllPermissionErrors variable for list errors and details" -foregroundcolor red
Write-host $NotFoundUsers.count "Mail Objects were not found. Check the $NotFoundUsers variable for list of users" -foregroundcolor yellow
Write-host $permsAlreadyExist.count "Permissions Already Exists. Check the $permsAlreadyExist variable for list of permissions" -foregroundcolor yellow
Write-host $completedPerms.count "Permissions applied. Check the $completedPerms variable for list" -foregroundcolor Green

if ($OutputCSVFolderPath) {
    $AllPermissionErrors | Export-Csv "$OutputCSVFolderPath\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
    Write-host "Exported file 'DelegatePermissions.csv' List to $OutputCSVFolderPath" -ForegroundColor Cyan
}
elseif ($OutputCSVFilePath) {
    $AllPermissionErrors | Export-Csv $OutputCSVFilePath -NoTypeInformation -Encoding UTF8
    Write-host "Exported Permissions List to $OutputCSVFilePath" -ForegroundColor Cyan
}
else {
    try {
        $AllPermissionErrors | Export-Csv "$HOME\Desktop\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
        Write-host "Exported 'Error-DelegatePermissions.csv' List to $HOME\Desktop" -ForegroundColor Cyan
    }
    catch {
        Write-Warning -Message "$($_.Exception)"
        Write-host ""
        $OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
        $AllPermissionErrors | Export-Csv "$OutputCSVFolderPath\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
    }
}
