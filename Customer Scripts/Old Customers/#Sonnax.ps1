#Sonnax

$dlsProperties = Import-Csv $HOME\Desktop\DLsProperties.csv | sort PrimarySmtpAddress

# Create any missing DLs
$failedToCreate = @()
$failedToUpdate = @()
foreach ($dl in $dlsProperties)
{
    $tempEmailAddress = $dl.PrimarySmtpAddress.replace("@sonnax.com","_temp@sonnax.mail.onmicrosoft.com")
    $tempname = $dl.Name + "_temp"
	Write-Host "$($tempEmailAddress) ... " -ForegroundColor Cyan -NoNewline
    	
	if (-not (Get-DistributionGroup $tempEmailAddress -EA SilentlyContinue))
	{
        $managedBy = $dl.ManagedBy
		Write-Host "creating ... " -ForegroundColor Yellow
		try
		{
            
			$newDL = New-DistributionGroup -Name $tempname -Alias $dl.Alias -DisplayName $dl.DisplayName -PrimarySmtpAddress $tempEmailAddress -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($dl.RequireSenderAuthenticationEnabled)) -Confirm:$false #-WhatIf
		}
		catch
		{
			Write-Host "fail to create" -ForegroundColor Red
			$failedToCreate += $dl
			continue
		}
        try
        {
            $newDL | Set-DistributionGroup -HiddenFromAddressListsEnabled $true -Confirm:$false #-EmailAddresses $dl.EmailAddresses.Split(",") 
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
	}
}

# Add group members
$groupNotFound = @()
$failedToAddMember = @()
$addedMembers = @()
$notfounduser = @()
foreach ($dl in $dlsProperties)
{
    $tempEmailAddress = $dl.PrimarySmtpAddress.replace("@sonnax.com","_temp@sonnax.mail.onmicrosoft.com")
    $tempname = $dl.Name + "_temp"

	Write-Host "$($tempEmailAddress) ..." -ForegroundColor Cyan -NoNewline
	
	if (-not ($group = Get-DistributionGroup $tempEmailAddress -EA SilentlyContinue))
	{
		Write-Host " not found" -ForegroundColor Red
		$groupNotFound += $dl
		continue
	}
	
	[array]$currentMembersGuids = Get-DistributionGroupMember $group.Guid.tostring()-EA SilentlyContinue | select -ExpandProperty Guid | select -ExpandProperty Guid
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
		Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding $($newMember) for group $($tempEmailAddress)"

		if ($recipient = Get-Recipient $newMember -EA SilentlyContinue)
		{
			if ($currentMembersGuids -notcontains $recipient.Guid.Guid)
			{
				try
				{
					Add-DistributionGroupMember $group.Guid.tostring()-Member $recipient.Guid.tostring()-EA Stop #-WhatIf
					$tmp = "" | select DL, NewMember
					$tmp.DL = $tempEmailAddress
					$tmp.NewMember = $newMember
					$addedMembers += $tmp
					#Write-Host "." -ForegroundColor Green -NoNewline
				}
				catch
				{
					$tmp = "" | select DL, Member
					$tmp.DL = $tempEmailAddress
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
			$tmp.DL = $tempEmailAddress
			$tmp.NewMember = $newMember
			$notfounduser += $tmp
			Write-Host "." -ForegroundColor Yellow -NoNewline
		}
	}
	
	Write-Host " done" -ForegroundColor Green
}

# Update Groups Post
$failedToCreate = @()
$failedToUpdate = @()
foreach ($dl in $dlsProperties)
{
    $tempEmailAddress = $dl.PrimarySmtpAddress.replace("@sonnax.com","_temp@sonnax.mail.onmicrosoft.com")
    $tempname = $dl.Name + "_temp"
	Write-Host "$($tempEmailAddress) ... " -ForegroundColor Cyan -NoNewline
    	
	if (Get-DistributionGroup $dl.PrimarySmtpAddress -EA SilentlyContinue)
	{
        $managedBy = $dl.ManagedBy
		Write-Host "updating ... " -ForegroundColor Yellow
		try
		{

			Set-DistributionGroup $dl.DisplayName -PrimarySmtpAddress $dl.PrimarySmtpAddress -name $dl.name

            $addresses = $dl.EmailAddresses -split ","
			foreach ($address in $addresses)
			{
				Set-DistributionGroup $dl.DisplayName -EmailAddresses @{add=$address}
			}
        
            Set-DistributionGroup $dl.DisplayName -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($dl.HiddenFromAddressListsEnabled)) -RequireSenderAuthenticationEnabled ([System.Convert]::ToBoolean($dl.RequireSenderAuthenticationEnabled))
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
		Write-Host "doesn't exist" -ForegroundColor DarkGreen
		$failedToCreate += $dl
	}
}

# Add group members
$groupNotFound = @()
$failedToAddMember = @()
$addedMembers = @()
$notfounduser = @()
foreach ($dl in $dlsProperties)
{
    $tempEmailAddress = $dl.PrimarySmtpAddress
    $tempname = $dl.Name + "_temp"

	Write-Host "$($tempEmailAddress) ..." -ForegroundColor Cyan -NoNewline
	
	if (-not ($group = Get-DistributionGroup $tempEmailAddress -EA SilentlyContinue))
	{
		Write-Host " not found" -ForegroundColor Red
		$groupNotFound += $dl
		continue
	}
	
	[array]$currentMembersGuids = Get-DistributionGroupMember $group.Guid.tostring()-EA SilentlyContinue | select -ExpandProperty Guid | select -ExpandProperty Guid
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
		Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding $($newMember) to group $($tempEmailAddress)"

		if ($recipient = Get-Recipient $newMember -EA SilentlyContinue)
		{
			if ($currentMembersGuids -notcontains $recipient.Guid.Guid)
			{
				try
				{
					Add-DistributionGroupMember $group.Guid.tostring()-Member $recipient.Guid.tostring()-EA Stop #-WhatIf
					$tmp = "" | select DL, NewMember
					$tmp.DL = $tempEmailAddress
					$tmp.NewMember = $newMember
					$addedMembers += $tmp
					#Write-Host "." -ForegroundColor Green -NoNewline
				}
				catch
				{
					$tmp = "" | select DL, Member
					$tmp.DL = $tempEmailAddress
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
			$tmp.DL = $tempEmailAddress
			$tmp.NewMember = $newMember
			$notfounduser += $tmp
			Write-Host "." -ForegroundColor Yellow -NoNewline
		}
	}
	
	Write-Host " done" -ForegroundColor Green
}