<# .SYNOPSIS
    This script can be used to pull delegatge permissions based on Calendar Permissions granted per mailbox. 
    Pulls unique values and exports lists of users sharing calendars and those not sharing calendars.

    Full List of Permissions are exported as "Perms_Calendar.csv". By Default exports to desktop.

   .PARAMETER OutputCSVFolderPath
   SubscriptionId - the subscription to use for key vault.

   .EXAMPLE
   Get-DelegatePermissions.ps1 -OutputCSVFolderPath C:\user\documents

   .EXAMPLE
   Get-DelegatePermissions.ps1
#>

param (
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full folderpath')] [string] $OutputCSVFolderPath,
	[switch]$OnPremises,
	[switch]$Office365
)
#Gather All Mailboxes
Write-Host "Gathering Mailboxes .." -foregroundcolor cyan -nonewline
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*DiscoverySearchMailbox*"}
#$mailboxes = Get-Mailbox -filter "PrimarySmtpAddress -notlike '*DiscoverySearchMailbox*'" -ResultSize Unlimited
Write-Host "done" -foregroundcolor green


#REGION Get list of calendar permissions

#ProgressBar 1 Initial
$progressref = ($mailboxes).count
$progresscounter = 0

#Build Array
$calendarpermsList = @()
$perms = @()

foreach ($mbx in $mailboxes)
{
    #ProgressBar 1
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($mbx.DisplayName)"

    #Calendar Array
	$primarySmtpAddress = $mbx.PrimarySMTPAddress
	[array]$calendars = Get-MailboxFolderStatistics $mbx.PrimarySMTPAddress | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}
	
    #Progress Bar 2 Initial
    $progressref2 = ($calendars).count
    $progresscounter2 = 0
	foreach ($calendar in $calendars)
	{
        #Folders
		$folderPath = $calendar.FolderPath.Replace('/','\')
		$id = "$primarySmtpAddress`:$folderPath"

        #Progress Bar 2
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Checking Calendar $($id)"
		
		try {
            [array]$perms = Get-MailboxFolderPermission $id -EA Stop | Where {$_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" -and $_.User.ToString() -notlike "*S-1-*"}
        }
        catch {
            $Output = "FAILED - $ID" | Out-file "$HOME\Desktop\Calendar_Errors.txt"-Append
            $Output = "$_.Exception.Message" | Out-file "$HOME\Desktop\Calendar_Errors.txt" -Append
        }
        #Progress Bar 3 Initial
        $progressref3 = ($perms).count
        $progresscounter3 = 0

        foreach ($perm in $perms)
        {
            #Progress Bar 3
            $progresscounter3 += 1
            $progresspercentcomplete3 = [math]::Round((($progresscounter3 / $progressref3)*100),2)
            $progressStatus3 = "["+$progresscounter3+" / "+$progressref3+"]"
            Write-progress -id 3 -PercentComplete $progresspercentcomplete3 -Status $progressStatus3 -Activity "Checking Calendar Perm for $($perm.User.DisplayName))"
            Write-Host "$($ID) .. Perm for $($perm.User.DisplayName)" -ForegroundColor Cyan

            $accessRights = $perm.AccessRights -join ","
            $currentPerm = new-object PSObject

            $currentPerm | add-member -type noteproperty -name "Mailbox" -Value $mbx.PrimarySmtpAddress.ToString()
            $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $calendar.FolderPath
            $currentPerm | add-member -type noteproperty -name "User" -Value $perm.User.DisplayName.ToString()
            $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
            
            #Write-Host "." -ForegroundColor Yellow -NoNewline
            $currentPerm | Export-Csv -NoTypeInformation -Encoding UTF8 "$HOME\Desktop\Perms_Calendar.csv" -Append
            
        }
	}
}
Write-host ""

if ($OutputCSVFolderPath) {
	$calendarpermsList | Export-Csv "$OutputCSVFolderPath\Perms_Calendar.csv" -NoTypeInformation -Encoding UTF8
    Write-host "Exported 'Perms_Calendar.csv' List to $OutputCSVFolderPath" -ForegroundColor Cyan
}
else {
	try {
		$calendarpermsList | Export-Csv "$HOME\Desktop\Perms_Calendar.csv" -NoTypeInformation -Encoding UTF8
		Write-host "Exported 'Perms_Calendar.csv' List to $HOME\Desktop" -ForegroundColor Cyan
	}
	catch {
		Write-Warning -Message "$($_.Exception)"
		Write-host ""
		$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
		$calendarpermsList | Export-Csv "$OutputCSVFolderPath\Perms_Calendar.csv" -NoTypeInformation -Encoding UTF8
	}
}
#ENDREGION


#REGION Separate the calendar sharing users from the non-sharing users
Write-Host ""
Write-Host "Creating Related and UnRelated Users based on Calendar Permissions"

$calendarSharingUsers = @()
$calendarSharingUsers += $calendarpermsList | select -ExpandProperty Mailbox
$calendarSharingUsers += $calendarpermsList | select -ExpandProperty User
$calendarSharingUsers = $calendarSharingUsers | sort -Unique | Where {$_}

$nonCalendarSharingUsers = @()

#ProgressBar
$progressref = ($mailboxes).count
$progresscounter = 0


foreach ($mbx in $mailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Reviewing Calendar Permissions for $($mbx.DisplayName)"

	$upn = $mbx.UserPrincipalName
	$primarySmtpAddress = $mbx.PrimarySmtpAddress.ToString()
	
	Write-Host "$primarySmtpAddress ... " -ForegroundColor Cyan -NoNewline
	
	if ($calendarSharingUsers -notcontains $upn -and $calendarSharingUsers -notcontains $primarySmtpAddress) {
		Write-Host "done" -ForegroundColor Green
		$nonCalendarSharingUsers += $primarySmtpAddress
	}
	else {
		Write-Host "done" -ForegroundColor Yellow
	}
}

#Output CSV Files
Write-Host ""
if ($OutputCSVFolderPath) {
	$calendarSharingUsers | Out-File $OutputCSVFolderPath\Desktop\RelatedUsers.txt
	$nonCalendarSharingUsers | Out-File $OutputCSVFolderPath\Desktop\UnrelatedUsers.txt
	Write-host "Exported 'Related' and 'UnRelatedUsers' lists to $OutputCSVFolderPath" -ForegroundColor Cyan
}
else {
	try {
		$calendarSharingUsers | Out-File $HOME\Desktop\RelatedUsers.txt
		$nonCalendarSharingUsers | Out-File $HOME\Desktop\UnrelatedUsers.txt
		Write-Host "Exported 'Related' and 'UnRelatedUsers' lists to $HOME\Desktop\" -ForegroundColor Cyan
	}
	catch {
		Write-Warning -Message "$($_.Exception)"
		Write-Host ""
		$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folderpath'
		$calendarSharingUsers | Out-File "$OutputCSVFolderPath\RelatedUsers.txt"
		$nonCalendarSharingUsers | Out-File "$OutputCSVFolderPath\UnrelatedUsers.txt"
		Write-host "Exported 'Related' and 'UnRelatedUsers' lists to $OutputCSVFolderPath" -ForegroundColor Cyan
	}
}
#ENDREGION