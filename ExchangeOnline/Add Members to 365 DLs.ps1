### New Region ###
#Update Members to DLs

$DLProperties = Import-Csv $HOME\desktop\DLsProperties.csv
$DLMemberResults = @()

foreach ($dl in $DLProperties) {
	write-host "Adding Members to DL" $dl.displayName "..." -foregroundcolor cyan
	$Group = Get-DistributionGroup $dl.name
		try 
			{	
				$members = $dl.members.Split(",") | sort -Unique
				$DLGroupMembers = Get-DistributionGroupMember $Group.name

				#check if member exists in tenant
				foreach ($member in $members | ? {$_}) 
				{
					for ($a = 1; $a -le 100; $a++ ){
						Write-Progress -Activity "Adding "$dlmember.DisplayName" as member of "$Group.name"" -Status "$a% Complete:" -PercentComplete $a;
					}
					if ($DlMember = Get-recipient $member -ea silentlycontinue) 
					{
						#CheckDL Members
												
						if (!($DLGroupMembers | ? {$_.name -eq $dlmember.Name} -ErrorAction SilentlyContinue)) 
						{                    
													  
							Add-DistributionGroupMember $group.name -member $DlMember.primarysmtpaddress -erroraction Stop #-whatif
							
							# Write-Host "added" $dlmember.DisplayName "as member" -ForegroundColor Green
							$tmp = "" | Select Member, DistributionList, Result, ErrorMessage
							$tmp.Member = $member
							$tmp.DistributionList = $Group.name
							$tmp.Result = "Added Successfully"
							$tmp.ErrorMessage = ""
							$DLMemberResults += $tmp
						}
						else 
						{
							# Write-Host $DlMember.DisplayName "is already a member. Skipping" -ForegroundColor Green
							$tmp = "" | Select Member, DistributionList, Result, ErrorMessage
							$tmp.Member = $member
							$tmp.DistributionList = $Group.name
							$tmp.Result = "AlreadyAdded"
							$tmp.ErrorMessage = ""
							$DLMemberResults += $tmp
						}
					}
					else 
					{
						Write-host $member "could not be found nor added to" $dl.name -foregroundcolor red
						$tmp = "" | Select Member, DistributionList, Result, ErrorMessage
						$tmp.Member = $member
						$tmp.DistributionList = $Group.name
						$tmp.Result = "Could not be found nor added"
						$tmp.ErrorMessage = ""
						$DLMemberResults += $tmp
					}
				}
				Write-Host "DistributionList updated successfully"	-ForegroundColor Green
			}
		catch
		{
			#$failedDLs += $dl
			Write-Host "error updating Distribution List" -ForegroundColor red
			$tmp = "" | Select Member, DistributionList, Result, ErrorMessage
			#$tmp.Member = ""
			$tmp.DistributionList = $Group.name
			$tmp.Result = "Failed"
			$tmp.ErrorMessage = $Error[0].Exception.Message
			$DLMemberResults += $tmp
		}
	Write-Host ""
	}
	
### End Of Region ###



### New Region ###
#Update Members to DLs; 1 member per line

$DLProperties = Import-Csv $HOME\desktop\DLsProperties.csv
$DLMemberResults = @()

foreach ($dl in $DLProperties) {
Write-host "Adding member "$dl.Member" to DL" $dl.Group"..." -foregroundcolor cyan -NoNewline
$Group = Get-DistributionGroup $dl.Group
$DLGroupMembers = Get-DistributionGroupMember $Group.name

	#check if member exists in tenant
	try 
	{
		if ($DlMember = Get-recipient $dl.Member -ea silentlycontinue) 
		{
			#CheckDL Members
									
			if (!($DLGroupMembers | ? {$_.name -eq $dlmember.Name} -ErrorAction SilentlyContinue)) 
			{                    
											
				Add-DistributionGroupMember $group.name -member $DlMember.Name -erroraction Stop #-whatif
				
				Write-Host "added" $dlmember.DisplayName "as member" -ForegroundColor Green
				$tmp = "" | Select Member, DistributionList, Result, ErrorMessage
				$tmp.Member = $dlmember.Name
				$tmp.DistributionList = $Group.name
				$tmp.Result = "Added Successfully"
				$tmp.ErrorMessage = ""
				$DLMemberResults += $tmp
			}
			else 
			{
				Write-Host "is already a member. Skipping." -ForegroundColor dark cyan -NoNewline
				$tmp = "" | Select Member, DistributionList, Result, ErrorMessage
				$tmp.Member = $member
				$tmp.DistributionList = $Group.name
				$tmp.Result = "AlreadyAdded"
				$tmp.ErrorMessage = ""
				$DLMemberResults += $tmp
			}
		}
		else 
		{
			Write-host $member "could not be found nor added." $dl.name -foregroundcolor red
			$tmp = "" | Select Member, DistributionList, Result, ErrorMessage
			$tmp.Member = $member
			$tmp.DistributionList = $Group.name
			$tmp.Result = "Could not be found nor added"
			$tmp.ErrorMessage = ""
			$DLMemberResults += $tmp
		}
	}
	catch
	{
		#$failedDLs += $dl
		Write-Host "error updating Distribution List" -ForegroundColor red
		$tmp = "" | Select Member, DistributionList, Result, ErrorMessage
		$tmp.Member = $dl.Member
		$tmp.DistributionList = $dl.group
		$tmp.Result = "Failed"
		$tmp.ErrorMessage = $Error[0].Exception.Message
		$DLMemberResults += $tmp
	}
}
### End Of Region ###