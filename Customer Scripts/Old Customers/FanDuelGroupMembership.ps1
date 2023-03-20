# Check Groups with users with domain fanduel.com

$FanDuelLists = @()
foreach ($dl in $DLs)
{
    $FanDuelMembers = Get-DistributionGroupMember $dl.identity -resultsize unlimited | ? {$_.primarysmtpaddress -like "*fanduel.com"}

    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $dl.DisplayName
    $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $dl.PrimarySMTPAddress
    $currentuser | add-member -type noteproperty -name "ID" -Value $dl.ID
    $currentuser | add-member -type noteproperty -name "FanDuel_Members" -Value $FanDuelMembers.Count
    $currentuser | add-member -type noteproperty -name "Members" -Value $FanDuelMembers.PrimarySMTPAddress
    $FanDuelLists += $currentuser
}