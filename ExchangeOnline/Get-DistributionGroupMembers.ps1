
$allMailGroups = Get-Recipient -RecipientTypeDetails MailNonUniversalGroup -ResultSize unlimited
$allMailGroups += Get-Recipient -RecipientTypeDetails group -ResultSize unlimited
$allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize unlimited
$allMailGroups += Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize unlimited    
$allMailGroups += Get-Recipient -RecipientTypeDetails DynamicDistributionGroup -ResultSize unlimited

$OutputCSVFilePath = "C:\Users\ADamedrano\Desktop\AllGroupMembers.csv"
$allMailGroupMembers = @()

#ProgressBarA
$progressref = ($allMailGroups).count
$progresscounter = 0
foreach ($group in $allMailGroups) {
    #ProgressBarB
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Group Details for $($group.name)"

    #Clear Member Variable
    $groupMembers = @()

    #Pull DynamicDistributionGroup Details
    if ($group.RecipientTypeDetails -eq "DynamicDistributionGroup") {
        $groupMembers = Get-DynamicDistributionGroupMember $group.PrimarySMTPAddress.tostring()  -ErrorAction SilentlyContinue -ResultSize unlimited            
    }
    #Pull Mail Group Details
    elseif ($group.RecipientTypeDetails -eq "MailUniversalDistributionGroup" -or $group.RecipientTypeDetails -eq "MailUniversalSecurityGroup" -or $group.RecipientTypeDetails -eq "MailNonUniversalGroup") {
        $groupMembers = Get-DistributionGroupMember $group.PrimarySMTPAddress.tostring() -ResultSize unlimited -ErrorAction SilentlyContinue
    }
    #Gather Group Mailbox
    elseif ($group.RecipientTypeDetails -eq "GroupMailbox") {
        $groupMembers = Get-UnifiedGroupLinks -Identity $group.PrimarySMTPAddress.tostring()  -LinkType Member -ResultSize unlimited
    }

    Write-Host "Gathering Group Details for $($group.DisplayName) .." -NoNewline -ForegroundColor Cyan
    
    if ($groupMembers.count -gt 1) {
    Write-Host "Found $($groupMembers.count) members .." -ForegroundColor Yellow -NoNewline
    $progressref2 = ($groupMembers).count
    $progresscounter2 = 0
    foreach ($member in $groupMembers) {
            #ProgressBarB
            $progresscounter2 += 1
            $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
            $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
            Write-progress -Id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "$($member.name)"

            $currentgroup = New-Object psobject
            $currentgroup | Add-Member -Type noteproperty -Name "Group-DisplayName" -Value $group.DisplayName
            $currentgroup | Add-Member -Type noteproperty -Name "Group-PrimarySMTPAddress" -Value $group.PrimarySMTPAddress.tostring() 
            $currentgroup | Add-Member -Type noteproperty -Name "Group-RecipientTypeDetails" -Value $group.RecipientTypeDetails
            $currentgroup | Add-Member -Type noteproperty -Name "Member-DisplayName" -Value $member.DisplayName
            $currentgroup | Add-Member -Type noteproperty -Name "Member-PrimarySMTPAddress" -Value $member.PrimarySMTPAddress.tostring()
            $currentgroup | Add-Member -Type noteproperty -Name "Member-RecipientType" -Value $member.RecipientTypeDetails
                 
            $allMailGroupMembers += $currentgroup
        }
    }
    else {
     Write-Host "No members .." -ForegroundColor Red -NoNewline
    }
    Write-Host "done" -ForegroundColor Green
}
$allMailGroupMembers | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath


