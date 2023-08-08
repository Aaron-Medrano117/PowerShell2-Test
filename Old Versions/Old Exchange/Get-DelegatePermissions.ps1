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
   .\Get-DelegatePermissions.ps1 -Office365 -CalendarPermissions -FullAccess -SendAs -SendOnBehalf
   .EXAMPLE
   Pulls only Calendar Permissions for each mailbox in Office 365. Exports to user's documents folder.
   .\Get-DelegatePermissions.ps1 -Office365 -CalendarPermissions -OutputCSVFolderPath C:\user\documents
   .EXAMPLE
   Pulls all available permissions for each mailbox in On-Premises Exchange. Exports to default location of desktop.
   .\Get-DelegatePermissions.ps1 -OnPremises -CalendarPermissions -FullAccess -SendAs
   .EXAMPLE
   Pulls Full Access and Send As permissions for each mailbox in On-Premises Exchange. Exports to user's documents folder.
   .\Get-DelegatePermissions.ps1 -OnPremises -FullAccess -SendaAs -OutputCSVFolderPath C:\user\documents
#>

param (
    [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
	[Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FILEPATH')] [string] $OutputCSVFilePath,
    [Parameter(Mandatory=$false,HelpMessage="Run against OnPremises Exchange?")] [switch]$OnPremises,
	[Parameter(Mandatory=$false,HelpMessage="Run against Office365 Exchange Online?")] [switch]$Office365,
    [Parameter(Mandatory=$false,HelpMessage="Run all Permissions Report?")] [switch]$allPermissions,
    [Parameter(Mandatory=$false,HelpMessage="Run Calendar Permissions Report?")] [switch]$CalendarPermissions,
    [Parameter(Mandatory=$false,HelpMessage="Run Full Access Permissions Report?")] [switch]$FullAccess,
	[Parameter(Mandatory=$false,HelpMessage="Run Send OnBehalf Permissions Report?")] [switch]$SendOnBehalf,
    [Parameter(,Mandatory=$false,HelpMessage="Run Send As Permissions Report?")] [switch]$SendAs
)
#Gather All Mailboxes
Write-Host "Gathering Mailboxes .." -foregroundcolor cyan -nonewline
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySmtpAddress -notlike "*DiscoverySearchMailbox*"} | sort PrimarySmtpAddress
Write-Host "done" -foregroundcolor green


#REGION Get list of calendar permissions
#ProgressBar
$progressref = ($mailboxes).count
$progresscounter = 0

#Build Array
$CollectPermissionsList = @()
$calendarPerms = @()
$mbxPermissions = @()
$SendAsPerms = @()

foreach ($mbx in $mailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($mbx.DisplayName)"
	$upn = $mbx.UserPrincipalName
	Write-Host "Checking Perms for $($mbx.PrimarySmtpAddress)" -ForegroundColor Cyan -NoNewline
    Write-Host ".." -ForegroundColor Yellow -NoNewline

    if ($allPermissions) {
        <# Action to perform if the condition is true #>
    }
	if ($CalendarPermissions) {
		[array]$calendars = $mbx | Get-MailboxFolderStatistics | Where {$_.FolderPath -eq "/Calendar" -or $_.FolderPath -like "/Calendar/*"}
	
		Write-Host "CalendarPerm.." -NoNewline -foregroundcolor DarkCyan
		foreach ($calendar in $calendars) {
			$folderPath = $calendar.FolderPath.Replace('/','\')
			$id = "$upn`:$folderPath"
			
			if ($OnPremises) {
				[array]$calendarPerms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.User.ADRecipient -and $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" -and $_.User.ToString() -notlike "*S-1-*" -and $_.User.ADRecipient.PrimarySmtpAddress.ToString() -ne $mbx.PrimarySmtpAddress.ToString()}
			}
			elseif ($Office365) {
				[array]$calendarPerms = Get-MailboxFolderPermission $id -EA SilentlyContinue | Where {$_.user.Usertype.value -ne "Default" -and $_.user.usertype.value -ne "Anonymous" -and $_.user.usertype.value -notlike "*S-1-*"}
			}
			if ($calendarPerms) {
				Write-Host $folderPath -ForegroundColor Green -NoNewline

                foreach ($perm in $calendarPerms) {
                    $accessRights = $perm.AccessRights -join ","
                    $SharingPermissionFlags = $perm.SharingPermissionFlags -join ","
                    $currentPerm = new-object PSObject				

                    $currentPerm | add-member -type noteproperty -name "Mailbox" -Value $mbx.PrimarySmtpAddress.ToString()
                    $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "Calendar"
                    
                    if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                        $currentPerm | add-member -type noteproperty -name "PermUser" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
                    }
                    else {
                        $currentPerm | add-member -type noteproperty -name "PermUser" -Value $null -force
                    }
                    $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                    $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $id
                    $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $SharingPermissionFlags
                    $CollectPermissionsList += $currentPerm
                    Write-Host "." -ForegroundColor Yellow -NoNewline
                }
			}
		}
	}
	if ($FullAccess) {
		if ($OnPremises) {
            $mbxPermissions = Get-MailboxPermission $mbx.PrimarySmtpAddress.ToString() -EA SilentlyContinue | Where {$_.User.ToString() -ne "NT Authority\Self" -and $_.User.ToString() -notlike "*S-1-*" -and $_.AccessRights -like "FullAccess"}
        }
        elseif ($Office365) {
            $mbxPermissions = Get-MailboxPermission $mbx.PrimarySmtpAddress.ToString() -EA SilentlyContinue | Where {$_.User.ToString() -ne "NT Authority\Self" -and $_.User.ToString() -notlike "*S-1-*"}
            }
		if ($mbxPermissions) {
            Write-Host "FullAccess.." -NoNewline -foregroundcolor DarkCyan
            foreach ($perm in $mbxPermissions) {
                if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
                    $accessRights = $perm.AccessRights -join ","
                    $currentPerm = new-object PSObject
                    $currentPerm | add-member -type noteproperty -name "Mailbox" -Value $mbx.PrimarySmtpAddress.ToString()
                    $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "FullAccess"
                    $currentPerm | add-member -type noteproperty -name "PermUser" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
                    $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
                    $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                    $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                    $CollectPermissionsList += $currentPerm
                    Write-Host "." -ForegroundColor Yellow -NoNewline
                }		
            }
		}
	}
	if ($SendAs) {
        Write-Host "SendAs.." -NoNewline -foregroundcolor DarkCyan
		if ($OnPremises) {
			[array]$SendAsPerms = Get-User $mbx.Identity  | Get-ADPermission -EA Stop | Where {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "*S-1-*"}         
        }
        elseif ($Office365) {
            [array]$SendAsPerms = Get-RecipientPermission $mbx.Identity -EA SilentlyContinue | Where {$_.Trustee.ToString() -ne "NT Authority\Self" -and $_.Trustee.ToString() -notlike "*S-1-*"}
        }
        foreach ($perm in $SendAsPerms) {
            $accessRights = $perm.AccessRights -join ","
            $currentPerm = new-object PSObject
            $currentPerm | add-member -type noteproperty -name "Mailbox" -Value $mbx.PrimarySmtpAddress.ToString()
            $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "SendAs"
            
			if ($OnPremises) {
				if ($recipientCheck = Get-Recipient $perm.User.ToString() -ea silentlycontinue) {
					$currentPerm | add-member -type noteproperty -name "PermUser" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
				}
				else {
					$currentPerm | add-member -type noteproperty -name "PermUser" -Value $null -force
				}
			}
			elseif ($Office365) {
				if ($recipientCheck = Get-Recipient $perm.Trustee.ToString() -ea silentlycontinue) {
					$currentPerm | add-member -type noteproperty -name "PermUser" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
				}
				else {
					$currentPerm | add-member -type noteproperty -name "PermUser" -Value $null -force
				}
			}
    
            $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $accessRights
            $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
            $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
            $CollectPermissionsList += $currentPerm
            Write-Host "." -ForegroundColor Yellow -NoNewline
        }
	}
	if ($SendOnBehalf) {
        Write-Host "SendOnBehalfTo.." -NoNewline -foregroundcolor DarkCyan
        if ($SendOnBehalfToPerms = $mbx.GrantSendOnBehalfTo) {
            foreach ($perm in $SendOnBehalfToPerms) {
                $currentPerm = new-object PSObject
                $currentPerm | add-member -type noteproperty -name "Mailbox" -Value $mbx.PrimarySmtpAddress.ToString()
                $currentPerm | add-member -type noteproperty -name "PermissionType" -Value "SendOnBehalfTo"
                #SendOnBehalfCheck       
                if ($recipientCheck = Get-Recipient $perm.DistinguishedName.ToString() -ea silentlycontinue) {
                    $currentPerm | add-member -type noteproperty -name "PermUser" -Value $recipientCheck.PrimarySMTPAddress.ToString() -Force
                }
                else {
                    $currentPerm | add-member -type noteproperty -name "PermUser" -Value $null -force
                }
                $currentPerm | add-member -type noteproperty -name "AccessRights" -Value $null
                $currentPerm | add-member -type noteproperty -name "CalendarPath" -Value $null
                $currentPerm | add-member -type noteproperty -name "SharingPermissionFlags" -Value $null
                $CollectPermissionsList += $currentPerm
                Write-Host "." -ForegroundColor Yellow -NoNewline
            }
        }   
	}
	Write-Host "done" -ForegroundColor Green
}
Write-host ""

if ($OutputCSVFolderPath) {
	$CollectPermissionsList | Export-Csv "$OutputCSVFolderPath\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
    Write-host "Exported file 'DelegatePermissions.csv' List to $OutputCSVFolderPath" -ForegroundColor Cyan
}
elseif ($OutputCSVFilePath) {
    $CollectPermissionsList | Export-Csv $OutputCSVFilePath -NoTypeInformation -Encoding UTF8
    Write-host "Exported Permissions List to $OutputCSVFilePath" -ForegroundColor Cyan
}
else {
	try {
		$CollectPermissionsList | Export-Csv "$HOME\Desktop\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
		Write-host "Exported 'DelegatePermissions.csv' List to $HOME\Desktop" -ForegroundColor Cyan
	}
	catch {
		Write-Warning -Message "$($_.Exception)"
		Write-host ""
		$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
		$CollectPermissionsList | Export-Csv "$OutputCSVFolderPath\DelegatePermissions.csv" -NoTypeInformation -Encoding UTF8
	}
}