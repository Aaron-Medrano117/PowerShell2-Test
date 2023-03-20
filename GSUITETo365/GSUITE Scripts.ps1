### Connect to GSUITE
Install-Module -Name PSGSuite -RequiredVersion 2.24.0

$ConfigName =  "GSuite"
$Preference = "Domain"
$P12KeyPath = # "C:\GSuite\psgsuite-284106-f422e66d3841.p12"
$AppEmail = # "psgsuite@xxxxxxxxxxxxxxxxxxxx.iam.gserviceaccount.com"
$AdminEmail = # "admin@example.info"
$Domain = # "example.info"
$ServiceAccountClientID = # "10745224254xxxxxxxxx"
 
Set-PSGSuiteConfig -ConfigName $ConfigName -P12KeyPath $P12KeyPath -AppEmail $AppEmail -AdminEmail $AdminEmail -Domain $Domain  -ServiceAccountClientID $ServiceAccountClientID

# Example
# Set-PSGSuiteConfig -ConfigName MyConfig -SetAsDefaultConfig -P12KeyPath "C:\Users\fred5646\Downloads\office365-migration-288301-9326959c65f7.p12" -AppEmail "gmail-onboarding@office365-migration-288301.iam.gserviceaccount.com" -AdminEmail "rackspace@proctoru.com" -Domain "proctoru.com" -Preference "Domain" -ServiceAccountClientID 100424271523516272781

#Export All users
$Mailboxes = Get-GSUserList | Where-Object {$_.IsMailboxSetup -eq $True} 

$AllGSUsers = @()
foreach ($User in $Mailboxes) { 

    $EXAttributes = $user.emails.Address
    [array]$UserAttributes = $User | select -ExpandProperty Name
    Write-Host "Getting Details for $($UserAttributes.FullName) ..." -NoNewline -foregroundcolor cyan

    $tmp = new-object PSObject

    $tmp | add-member -type noteproperty -name "UserPrincipalName" -Value $user.user
    $tmp | add-member -type noteproperty -name "LastLoginTime" -Value $User.LastLoginTime
    $tmp | add-member -type noteproperty -name "DisplayName" -Value $UserAttributes.FullName
    $tmp | add-member -type noteproperty -name "FirstName" -Value $UserAttributes.GivenName
    $tmp | add-member -type noteproperty -name "LastName" -Value $UserAttributes.FamilyName
    $tmp | add-member -type noteproperty -name "OrgUnitPath" -Value $user.OrgUnitPath
    $tmp | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.PrimaryEmail
    $tmp | add-member -type noteproperty -name "EmailAddresses" -Value ($EXAttributes -join ",")
    $tmp | add-member -type noteproperty -name "IncludeInGlobalAddressList" -Value $User.IncludeInGlobalAddressList
    $tmp | add-member -type noteproperty -name "IsSuspended" -Value $User.Suspended
    $tmp | add-member -type noteproperty -name "IsAdmin" -Value $User.IsAdmin
   
    $AllGSUsers += $tmp

    Write-Host "done" -foregroundcolor green

}

$AllGSUsers | Export-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\GSuiteUsers.csv" -Encoding utf8 -NoTypeInformation


#export GS Groups
$Groups = Get-GSGroup
$AllGSGroups = @()

Write-Host "Gathering $($Groups.Count) GSuite Groups and respective members ..." -ForegroundColor Cyan

foreach ($Group in $Groups) {
    #Write-Host "Gathering Details for $($Group.Name) ..." -ForegroundColor Cyan -NoNewline

    $GroupMembers = Get-GSGroupMember -Identity $Group.id

    $tmp = "" | Select Group, DirectMembersCount, PrimarySmtpAddress, Description, Members 

    #Write-Host "Adding $($Group.DirectMembersCount) members to group ..." -ForegroundColor Yellow -NoNewline

    $tmp.Group =$Group.Name
    $tmp.DirectMembersCount = $Group.DirectMembersCount
    $tmp.PrimarySmtpAddress = $Group.Email
    $tmp.Description = $Group.Description
    $tmp.Members = ($GroupMembers.Email -join ",")
    
    $AllGSGroups += $tmp

    #Write-Host "done" -ForegroundColor Green
}

Write-Host "done" -ForegroundColor Green


#export GS Group Members
$Groups = Get-GSGroup
$AllGroupsMembers = @()

foreach ($Group in $Groups) {
    Write-Host "Gathering Details for $($Group.Name) ..." -ForegroundColor Cyan -NoNewline

    $GroupMembers = Get-GSGroupMember -Identity $Group.id

    #Write-Host "Adding $($Group.DirectMembersCount) members to group ..." -ForegroundColor Yellow -NoNewline

    foreach ($member in $GroupMembers) {

        $tmp = "" | Select Group, DirectMembersCount, PrimarySmtpAddress, Description, Member, MemberRole, MemberType
        $tmp.Group = $Group.Name
        $tmp.DirectMembersCount = $Group.DirectMembersCount
        $tmp.PrimarySmtpAddress = $Group.Email
        $tmp.Member = $member.email 
        $tmp.MemberRole = $member.Role
        $tmp.MemberType = $member.type

        $AllGroupsMembers += $tmp
    }
    Write-Host "done" -ForegroundColor Green
}

# Gather All GSUITE Calendar Subscriptions
$allGSUsers = Get-GSUserList
$allGSCalendars = @()

foreach ($GSUser in $AllGSUsers)
{
    $GSCalendars = @()

    $GSCalendars = Get-GSCalendarSubscription -User $GSUser.user
    Write-Host "Found $($GSCalendars.count) GSCalendar Subscriptions for $($GSUser.user). Gathering details ..." -NoNewline -ForegroundColor Cyan

    foreach ($GSCalendar in $GSCalendars)
    {
        $GSCalendar = Get-GSCalendarSubscription -User $GSUser.user -CalendarId $GSCalendar.id

        $tmp = new-object PSObject

        $tmp | add-member -type noteproperty -name "User" -Value $GSUser.user
        $tmp | add-member -type noteproperty -name "CalendarName" -Value $GSCalendar.Summary
        
        if ($GSUser.user -eq $GSCalendar.Summary)
        {
            $tmp | add-member -type noteproperty -name "OwnCalendar" -Value $True
        }

        else
        {
            $tmp | add-member -type noteproperty -name "OwnCalendar" -Value $False
        }

        $tmp | add-member -type noteproperty -name "AccessRole" -Value $GSCalendar.AccessRole
        $tmp | add-member -type noteproperty -name "ID" -Value $GSCalendar.ID
        $tmp | add-member -type noteproperty -name "Kind" -Value $GSCalendar.Kind

        $allGSCalendars += $tmp
    }

    Write-Host "done" -ForegroundColor Green   
}

#### Export GS Groups with ALL Details

#export GS Groups
function Get-ALLGSGroups {
    param ([Parameter(Mandatory=$false)] [string] $OutputCSVFilePath)
    $Groups = Get-GSGroup
    $AllGSGroups = @()
    foreach ($Group in $Groups) {
        Write-Host "Gathering Details for $($Group.Name) ... " -ForegroundColor Cyan -NoNewline
    
        $GroupMembers = Get-GSGroupMember -Identity $Group.id
        $Owners = $GroupMembers |? {$_.Role -eq "Owner"}
        $GroupSettings = Get-GSGroupSettings $Group.id
        
        #closed or open group
        if ($GroupSettings.WhoCanJoin -eq "ALL_IN_DOMAIN_CAN_JOIN")
        {
            $Group_MemberJoinRestriction = "Open"
        }
        else
        {
            $Group_MemberJoinRestriction = "Closed"
        }
    
        # Create Output Array
        $currentgroup = new-object PSObject
    
        $currentgroup | add-member -type noteproperty -name "Group" -Value $Group.Name
        $currentgroup | add-member -type noteproperty -name "DirectMembersCount" -Value $Group.DirectMembersCount
        $currentgroup | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $Group.Email
        $currentgroup | add-member -type noteproperty -name "Description" -Value $Group.Description
        $currentgroup | add-member -type noteproperty -name "Members" -Value ($GroupMembers.Email -join ",")
        $currentgroup | add-member -type noteproperty -name "Owners" -Value ($Owners.Email -join ",")
    
        $currentgroup | add-member -type noteproperty -name "AllowExternalMembers" -Value $GroupSettings.AllowExternalMembers
        $currentgroup | add-member -type noteproperty -name "MemberJoinRestriction" -Value $Group_MemberJoinRestriction
        $currentgroup | add-member -type noteproperty -name "ShowInGroupDirectory" -Value $GroupSettings.ShowInGroupDirectory
        
        # Match Group in Office 365
    
        Write-Host "Checking user in Office 365 ... " -NoNewline -ForegroundColor Cyan
        $EmailSplit = $Group.Email -split "@"
        if ($EXOGroup = get-recipient $EmailSplit[0] -ea silentlycontinue)
        { 
            $currentgroup | add-member -type noteproperty -name "ExistsOnO365" -Value $true
            $currentgroup | add-member -type noteproperty -name "O365_DisplayName" -value $EXOGroup.DisplayName
            $currentgroup | add-member -type noteproperty -name "O365_PrimarySMTPAddress" -Value $EXOGroup.PrimarySMTPAddress
            $currentgroup | add-member -type noteproperty -name "RecipientType" -Value $EXOGroup.RecipientTypeDetails
            Write-Host "found" -ForegroundColor Green -NoNewline
        }
    
        elseif ($EXOGroupDisplay = Get-recipient $group.Name -ea silentlycontinue)
        {
            $currentgroup | add-member -type noteproperty -name "ExistsOnO365" -Value $true
            $currentgroup | add-member -type noteproperty -name "O365_DisplayName" -value $EXOGroupDisplay.DisplayName
            $currentgroup | add-member -type noteproperty -name "O365_PrimarySMTPAddress" -Value $EXOGroupDisplay.PrimarySMTPAddress
            $currentgroup | add-member -type noteproperty -name "RecipientType" -Value $EXOGroupDisplay.RecipientTypeDetails
            Write-Host "found*" -ForegroundColor Green -NoNewline
        }
    
        else
        {
            Write-Host "not found" -ForegroundColor red -NoNewline
            $currentgroup | add-member -type noteproperty -name "ExistsOnO365" -Value $False
            $currentgroup | add-member -type noteproperty -name "O365_DisplayName" -value ""
            $currentgroup | add-member -type noteproperty -name "O365_PrimarySMTPAddress" -Value ""
            $currentgroup | add-member -type noteproperty -name "RecipientType" -Value ""
        }
        $AllGSGroups += $currentgroup 
        Write-Host " .. done" -ForegroundColor Green
    }

    if ($OutputCSVFilePath)
    {
        $AllGSGroups | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
    }
    else
    {
        $OutputCSVFilePath = Read-Host "Where do you wish to export the file?"
        $csvFileName = "GSUITE_MatchedGroups.csv"
        Write-host "Exported Matched Groups Report" -foregroundcolor cyan
        $AllGSGroups | Export-Csv -NoTypeInformation -Encoding utf8 "$($OutputCSVFilePath)\$($csvFileName)"
    } 
}
