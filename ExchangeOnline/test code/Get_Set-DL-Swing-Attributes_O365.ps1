

#####################################
# Get DLs swing attributes from HEX #
#####################################

function Get-DistributionGroupAttributes {
	param (
		[Parameter(Mandatory=$false)] [string] $customerDomain,
        [Parameter(Mandatory=$false)][switch] $HEX,
		[Parameter(Mandatory=$false)][switch] $O365,
		[Parameter(Mandatory=$false)][string] $ExportFilePath
	)

	if ($HEX)
	{
		$allDLs = Get-DistributionGroup -OrganizationalUnit $customerDomain -ResultSize Unlimited | sort PrimarySmtpAddress
	}
	elseif ($O365)
	{
		$allDLs = Get-DistributionGroup -ResultSize Unlimited | sort PrimarySmtpAddress
	}
	else {
		$allDLs = Get-DistributionGroup -ResultSize Unlimited | sort PrimarySmtpAddress
	}

	$dlsProperties = @()
	#ProgressBar
	$progressref = ($allDLs).count
	$progresscounter = 0
	foreach ($dl in $allDLs)
	{
		#ProgressBar2
		$progresscounter += 1
		$progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
		$progressStatus = "["+$progresscounter+" / "+$progressref+"]"
		Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering DistribtuionGroup Details for $($dl.DisplayName)"
		
		Write-Host "$($dl.PrimarySmtpAddress) ... " -ForegroundColor Cyan -NoNewline
		
		$group = New-Object -TypeName PSObject
		$group | Add-Member -MemberType NoteProperty -Name Identity -Value $dl.Identity
		$group | Add-Member -MemberType NoteProperty -Name Name -Value $dl.Name
		$group | Add-Member -MemberType NoteProperty -Name MemberJoinRestriction -Value $dl.MemberJoinRestriction
		$group | Add-Member -MemberType NoteProperty -Name ReportToManagerEnabled -Value $dl.ReportToManagerEnabled
		$group | Add-Member -MemberType NoteProperty -Name ReportToOriginatorEnabled -Value $dl.ReportToOriginatorEnabled
		$group | Add-Member -MemberType NoteProperty -Name SendOOFMessageToOriginatorEnabled -Value $dl.SendOOFMessageToOriginatorEnabled
		$group | Add-Member -MemberType NoteProperty -Name Alias -Value $dl.Alias
		$group | Add-Member -MemberType NoteProperty -Name DisplayName -Value $dl.DisplayName
		$group | Add-Member -MemberType NoteProperty -Name LegacyExchangeDN -Value $dl.LegacyExchangeDN
		$group | Add-Member -MemberType NoteProperty -Name HiddenFromAddressListsEnabled -Value $dl.HiddenFromAddressListsEnabled
		$group | Add-Member -MemberType NoteProperty -Name ModerationEnabled -Value $dl.ModerationEnabled
		$group | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $dl.PrimarySmtpAddress.ToString()
		$group | Add-Member -MemberType NoteProperty -Name RecipientType -Value $dl.RecipientType
		$group | Add-Member -MemberType NoteProperty -Name RecipientTypeDetails -Value $dl.RecipientTypeDetails
		$group | Add-Member -MemberType NoteProperty -Name RequireSenderAuthenticationEnabled -Value $dl.RequireSenderAuthenticationEnabled
		$group | Add-Member -MemberType NoteProperty -Name SendModerationNotifications -Value $dl.SendModerationNotifications
		$group | Add-Member -MemberType NoteProperty -Name WindowsEmailAddress -Value $dl.WindowsEmailAddress.ToString()
		$group | Add-Member -MemberType NoteProperty -Name Id -Value $dl.Id
		$group | Add-Member -MemberType NoteProperty -Name WhenCreated -Value $dl.WhenCreated.ToString()

		if ($HEX)
		{
			#On-prem Exchange
			#Get DLMembers
			$GroupMembers = Get-DistributionGroupMember $dl.identity -ResultSize unlimited
			$DLMembers = @()
			foreach ($member in $GroupMembers)
			{				
				$DLMembers += $member.primarysmtpaddress.tostring()
			}

			$group | Add-Member -MemberType NoteProperty -Name ManagedBy -Value ($dl.ManagedBy -join ",")
			$group | Add-Member -MemberType NoteProperty -Name AcceptMessagesOnlyFromSendersOrMembers -Value ($dl.AcceptMessagesOnlyFromSendersOrMembers -join ",")
			$group | Add-Member -MemberType NoteProperty -Name BypassModerationFromSendersOrMembers -Value ($dl.BypassModerationFromSendersOrMembers -join ",")
			$group | Add-Member -MemberType NoteProperty -Name EmailAddresses -Value (($dl.EmailAddresses -join ",") + (",X500:" + $dl.LegacyExchangeDN))
			$group | Add-Member -MemberType NoteProperty -Name GrantSendOnBehalfTo -Value ($dl.GrantSendOnBehalfTo -join ",")
			$group | Add-Member -MemberType NoteProperty -Name ModeratedBy -Value ($dl.ModeratedBy -join ",")
			$group | Add-Member -MemberType NoteProperty -Name RejectMessagesFromSendersOrMembers -Value ($dl.RejectMessagesFromSendersOrMembers -join ",")
			$group | Add-Member -MemberType NoteProperty -Name Members -Value ($DLMembers -join ",")
		}	

		elseif ($O365)
		{
			#Get SendAs Permissions
			if ([array]$sendAsPermsCheck = Get-RecipientPermission $dl.identity -ResultSize unlimited) {
				$group | Add-Member -MemberType NoteProperty -Name SendAs -Value ($sendAsPermsCheck.Trustee -join ",")
			}
			else {
				$group | Add-Member -MemberType NoteProperty -Name SendAs -Value $null
			}

			#Office 365 Groups
			$group | Add-Member -MemberType NoteProperty -Name ManagedBy -Value (($dl.ManagedBy | Get-Recipient -EA SilentlyContinue | select -ExpandProperty PrimarySmtpAddress) -join ",")
			$group | Add-Member -MemberType NoteProperty -Name AcceptMessagesOnlyFromSendersOrMembers -Value (($dl.AcceptMessagesOnlyFromSendersOrMembers | Get-Recipient -EA SilentlyContinue | select -ExpandProperty PrimarySmtpAddress) -join ",")
			$group | Add-Member -MemberType NoteProperty -Name BypassModerationFromSendersOrMembers -Value (($dl.BypassModerationFromSendersOrMembers | Get-Recipient -EA SilentlyContinue | select -ExpandProperty PrimarySmtpAddress) -join ",")
			$group | Add-Member -MemberType NoteProperty -Name EmailAddresses -Value (($dl.EmailAddresses -join ",") + (",X500:" + $dl.LegacyExchangeDN))
			$group | Add-Member -MemberType NoteProperty -Name GrantSendOnBehalfTo -Value (($dl.GrantSendOnBehalfTo | Get-Recipient -EA SilentlyContinue | select -ExpandProperty PrimarySmtpAddress) -join ",")
			$group | Add-Member -MemberType NoteProperty -Name ModeratedBy -Value (($dl.ModeratedBy | Get-Recipient -EA SilentlyContinue | select -ExpandProperty PrimarySmtpAddress) -join ",")
			$group | Add-Member -MemberType NoteProperty -Name RejectMessagesFromSendersOrMembers -Value (($dl.RejectMessagesFromSendersOrMembers | Get-Recipient -EA SilentlyContinue | select -ExpandProperty PrimarySmtpAddress) -join ",")
			$group | Add-Member -MemberType NoteProperty -Name MemberCount -Value ((Get-DistributionGroupMember $dl.primarysmtpaddress -ResultSize unlimited).count)
			$group | Add-Member -MemberType NoteProperty -Name Members -Value ((Get-DistributionGroupMember $dl.primarysmtpaddress | select -ExpandProperty PrimarySmtpAddress) -join ",")	

		}
		$dlsProperties += $group
		
		Write-Host "done" -ForegroundColor Green
	}
	if ($ExportFilePath) {
		$dlsProperties | Export-Csv "$ExportFilePath\DLProperties.csv" -NoTypeInformation -Encoding UTF8
		Write-host "Exported DL Property List to $ExportFilePath\DLProperties.csv" -ForegroundColor Cyan
	}
	else {
		try {
			$dlsProperties | Export-Csv "$HOME\Desktop\DLProperties.csv" -NoTypeInformation -Encoding UTF8
			Write-host "Exported DL Property List to $HOME\Desktop\DLProperties.csv" -ForegroundColor Cyan
		}
		catch {
			Write-Warning -Message "$($_.Exception)"
			Write-host ""
			$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
			$dlsProperties | Export-Csv "$OutputCSVFolderPath\DLProperties.csv" -NoTypeInformation -Encoding UTF8
		}
	}
}

#################################################################





#REGION O365

##############################################################################################

$dlsProperties = Import-Csv $HOME\Desktop\DLsProperties.csv | sort PrimarySmtpAddress

# Create any missing DLs
$failedToCreate = @()
$failedToUpdate = @()
foreach ($dl in $dlsProperties)
{
	Write-Host "$($dl.Destination_PrimarySmtpAddress) ... " -ForegroundColor Cyan -NoNewline
	
	if (-not (Get-DistributionGroup $dl.Destination_PrimarySmtpAddress -EA SilentlyContinue))
	{
		Write-Host "creating ... " -ForegroundColor Yellow -NoNewline
		try
		{
			$newDL = New-DistributionGroup -Name $dl.Name -Alias $dl.Alias -DisplayName $dl.DisplayName -PrimarySmtpAddress $dl.Destination_PrimarySmtpAddress -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($dl.RequireSenderAuthenticationEnabled)) -ManagedBy "MigrationSvc@tjuv.onmicrosoft.com"  -Confirm:$false 
		}
		catch
		{
			Write-Host "fail to create" -ForegroundColor Red
			$failedToCreate += $dl
			continue
		}
		
		try
		{
			$newDL | Set-DistributionGroup -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($dl.HiddenFromAddressListsEnabled)) -EmailAddresses $dl.EmailAddresses.Split(",") -Confirm:$false
		}
		catch
		{
			Write-Host "failed to update" -ForegroundColor Red
			$failedToUpdate += $dl
			continue
		}
		
		Write-Host "done" -ForegroundColor Green
	} else
	{
		Write-Host "exists" -ForegroundColor DarkGreen
		try
		{
			$newDL | Set-DistributionGroup -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($dl.HiddenFromAddressListsEnabled)) -MemberJoinRestriction $dl.MemberJoinRestriction -ReportToOriginatorEnabled ([System.Convert]::ToBoolean($dl.ReportToOriginatorEnabled)) -SendModerationNotifications $dl.SendModerationNotifications -Confirm:$false
			$newDL | Set-DistributionGroup -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($dl.HiddenFromAddressListsEnabled)) -EmailAddresses $dl.EmailAddresses.Split(",") -Confirm:$false
		}
		catch
		{
			Write-Host "failed to update" -ForegroundColor Red
			$failedToUpdate += $dl
			continue
		}
	}
}

# Temp Create DLs
$failedToCreate = @()
$failedToUpdate = @()
foreach ($dl in $dlsProperties)
{
	Write-Host "$($dl.Destination_PrimarySmtpAddress) ... " -ForegroundColor Cyan -NoNewline
	
	if (-not (Get-DistributionGroup $dl.Destination_PrimarySmtpAddress -EA SilentlyContinue))
	{
		Write-Host "creating ... " -ForegroundColor Yellow -NoNewline
		try
		{
			$newDL = New-DistributionGroup -Name $dl.Name -Alias $dl.Alias -DisplayName $dl.DisplayName -PrimarySmtpAddress $dl.Destination_PrimarySmtpAddress -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($dl.RequireSenderAuthenticationEnabled)) -ManagedBy "MigrationSvc@tjuv.onmicrosoft.com"  -Confirm:$false 
		}
		catch
		{
			Write-Host "fail to create" -ForegroundColor Red
			$failedToCreate += $dl
			continue
		}
		
		try
		{
			$newDL | Set-DistributionGroup -HiddenFromAddressListsEnabled $true -MemberJoinRestriction $dl.MemberJoinRestriction -ReportToOriginatorEnabled ([System.Convert]::ToBoolean($dl.ReportToOriginatorEnabled)) -SendModerationNotifications $dl.SendModerationNotifications -Confirm:$false
		}
		catch
		{
			Write-Host "failed to update" -ForegroundColor Red
			$failedToUpdate += $dl
			continue
		}
		
		Write-Host "done" -ForegroundColor Green
	} else
	{
		Write-Host "exists" -ForegroundColor DarkGreen
		try
		{
			$newDL | Set-DistributionGroup -HiddenFromAddressListsEnabled $true -MemberJoinRestriction $dl.MemberJoinRestriction -ReportToOriginatorEnabled ([System.Convert]::ToBoolean($dl.ReportToOriginatorEnabled)) -SendModerationNotifications $dl.SendModerationNotifications -Confirm:$false
		}
		catch
		{
			Write-Host "failed to update" -ForegroundColor Red
			$failedToUpdate += $dl
			continue
		}
	}
}

# Add group members
$groupNotFound = @()
$failedToAddMember = @()
$addedMembers = @()
$notfounduser = @()
foreach ($dl in $dlsProperties)
{
	Write-Host "$($dl.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	if (-not ($group = Get-DistributionGroup $dl.PrimarySmtpAddress -EA SilentlyContinue))
	{
		Write-Host " not found" -ForegroundColor Red
		$groupNotFound += $dl
		continue
	}
	
	[array]$currentMembersGuids = Get-DistributionGroupMember $group.Guid -EA SilentlyContinue | select -ExpandProperty Guid | select -ExpandProperty Guid
	[array]$newMembers = $dl.Members.Split(",")
	
	    #ProgressBar1
		$progressref = ($newMembers).count
		$progresscounter = 0
			
	foreach ($newMember in $newMembers)
	{
		#ProgressBar2
		$progresscounter += 1
		$progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
		$progressStatus = "["+$progresscounter+" / "+$progressref+"]"
		Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding $($newMember) for group $($dl.PrimarySmtpAddress)"

		if ($recipient = Get-Recipient $newMember -EA SilentlyContinue)
		{
			if ($currentMembersGuids -notcontains $recipient.Guid.Guid)
			{
				try
				{
					Add-DistributionGroupMember $group.Guid.Guid -Member $recipient.Guid.Guid -EA Stop #-WhatIf
					$tmp = "" | select DL, NewMember
					$tmp.DL = $dl.PrimarySmtpAddress
					$tmp.NewMember = $newMember
					$addedMembers += $tmp
					#Write-Host "." -ForegroundColor Green -NoNewline
				}
				catch
				{
					$tmp = "" | select DL, Member
					$tmp.DL = $dl.PrimarySmtpAddress
					$tmp.Member = $newMember
					$failedToAddMember += $tmp
					Write-Host "." -ForegroundColor Red -NoNewline
				}
			}
			else
			{
				#Write-Host "." -ForegroundColor DarkGray -NoNewline
			}
		}
		else
		{
			$tmp = "" | select DL, NewMember
			$tmp.DL = $dl.PrimarySmtpAddress
			$tmp.NewMember = $newMember
			$notfounduser += $tmp
			Write-Host "." -ForegroundColor Yellow -NoNewline
		}
	}
	
	Write-Host " done" -ForegroundColor Green
}

##############################################################################################

#ENDREGION




#REGION On-Prem

##############################################################################################

$dlsProperties = Import-Csv $HOME\Desktop\DLsProperties.csv

$foundGroups = @()
foreach ($dl in $dlsProperties)
{
	Write-Host "$($dl.PrimarySmtpAddress) ... " -ForegroundColor Cyan -NoNewline
	
	if (-not ($group = Get-Group $dl.PrimarySmtpAddress -ea SilentlyContinue))
	{
		if (-not ($group = Get-Group $dl.Name -ea SilentlyContinue))
		{
			$group = Get-Group $dl.DisplayName -ea SilentlyContinue
		}
	}
	
	if ($group)
	{
		Remove-ADGroup $group.DistinguishedName -Confirm:$false #-WhatIf
		$foundGroups += $group
		Write-Host "done" -ForegroundColor Green
	}
	else
	{
		Write-Host "not found" -ForegroundColor Yellow
	}
}


$failed = @()
$OU = "OU=E-Mail Groups,DC=ccmsi,DC=com"
foreach ($dl in $dlsProperties)
{
	Write-Host "$($dl.PrimarySmtpAddress) ... " -ForegroundColor Cyan -NoNewline
	
	$managedBy = $dl.ManagedBy.Split(",")
	
	if ($list = New-DistributionGroup -Name $dl.Name -Alias $dl.Alias -DisplayName $dl.DisplayName -PrimarySmtpAddress $dl.PrimarySmtpAddress -RequireSenderAuthenticationEnabled $([System.Convert]::ToBoolean($dl.RequireSenderAuthenticationEnabled)) -ManagedBy $managedBy -OrganizationalUnit $OU)
	{
		$emailAddresses = $dl.EmailAddresses.Split(",")
		Set-DistributionGroup $list.Guid.Guid -EmailAddresses $emailAddresses -HiddenFromAddressListsEnabled $([System.Convert]::ToBoolean($dl.HiddenFromAddressListsEnabled)) #-WhatIf
		
		Write-Host "done" -ForegroundColor Green
	}
	else
	{
		$failed += $dl
		Write-Host "failed" -ForegroundColor Red
	}
}


foreach ($dl in $dlsProperties)
{
	Write-Host "$($dl.PrimarySmtpAddress) ..." -ForegroundColor Cyan -NoNewline
	
	$list = Get-DistributionGroup $dl.PrimarySmtpAddress
	
	if ($members = $dl.Members.Split(",") | sort -Unique)
	{
		foreach ($member in $members)
		{
			Write-Host "." -ForegroundColor Magenta -NoNewline
			Add-DistributionGroupMember $list.Guid.Guid -Member $member
		}
	}
	
	Write-Host " done" -ForegroundColor Green
}

##############################################################################################

#ENDREGION


















