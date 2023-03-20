#AmericanCarCenter

$DLs = Get-DistributionGroup -organizationalunit americancarcenter.com | sort DisplayName

$alldlmembers = @()

foreach ($dl in $dls)
{
    $members = Get-DistributionGroupMember $dl.alias                            
    foreach ($member in $members)
        {
            $tmp = "" | Select DistributionList, Member, RecipientType
            $tmp.DistributionList = $dl.name
            $tmp.Member = $member.primarysmtpaddress
            $tmp.RecipientType = $member.recipienttype
        }

    $alldlmembers += $tmp
}


$AllDG = Get-DistributionGroup -organizationalunit americancarcenter.com | sort DisplayName

$i = 0
$output = @()
Foreach($dg in $allDg)
{
$Members = Get-DistributionGroupMember $Dg.alias -resultsize unlimited

if($members.count -eq 0)
{
$managers = $Dg | Select @{Name='DistributionGroupManagers';Expression={[string]::join(";", ($_.Managedby))}}
$manageremail = Get-Mailbox $managers.DistributionGroupManagers -ErrorAction SilentlyContinue -resultsize unlimited

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmptyGroup
$userObj | Add-Member NoteProperty -Name "Alias" -Value EmptyGroup
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value EmptyGroup
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value EmptyGroup
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmptyGroup
$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
$userObj | Add-Member NoteProperty -Name "Distribution Group Managers Primary SMTP address" -Value $manageremail.primarysmtpaddress
$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
$userObj | Add-Member NoteProperty -Name "Not Allowed from Internet" -Value $DG.RequireSenderAuthenticationEnabled

$output += $UserObj  

}
else
{
Foreach($Member in $members)
 {

$managers = $Dg | Select @{Name='DistributionGroupManagers';Expression={[string]::join(";", ($_.Managedby))}}
$manageremail = Get-Mailbox $managers.DistributionGroupManagers -ErrorAction SilentlyContinue -resultsize unlimited

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Member.Name
$userObj | Add-Member NoteProperty -Name "Alias" -Value $Member.Alias
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Member.RecipientType
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Member.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Member.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
$userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
$userObj | Add-Member NoteProperty -Name "Distribution Group Managers Primary SMTP address" -Value $manageremail.primarysmtpaddress
$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
$userObj | Add-Member NoteProperty -Name "Not Allowed from Internet" -Value $DG.RequireSenderAuthenticationEnabled

$output += $UserObj  

 }
}
# update counters and write progress
$i++
Write-Progress -activity "Scanning Groups . . ." -status "Scanned: $i of $($allDg.Count)" -percentComplete (($i / $allDg.Count)  * 100)
}


###

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()
$foundDisplayName = @()

foreach ($user in $importcsv)
{
    Write-Host "Checking user $($user.DisplayName) in Office 365 ..." -fore Cyan -NoNewline
    $tmp = "" | select DisplayName, DistinguishedName, EmailAddress, ObjectGUID, ADUserPrincipalName, ExistsOnO365, O365_DisplayName, O365_UPN 
    $tmp.DisplayName = $user.DisplayName
    $tmp.DistinguishedName = $user.DistinguishedName
    $tmp.EmailAddress = $user.EmailAddress
    $tmp.ObjectGUID = $user.ObjectGUID
    $tmp.ADUserPrincipalName = $user.UserPrincipalName

    if ($msoluser = get-msoluser -userprincipalname $user.EmailAddress -ea silentlycontinue)
    { 
        $foundusers += $user.upn
        Write-Host "found" -ForegroundColor Green
        $tmp.O365_UPN = $msoluser.userprincipalname
        $tmp.O365_DisplayName = $msoluser.DisplayName
        $tmp.ExistsOnO365 = $true
    }

    elseif ($msolusersmtp = Get-msoluser -searchstring $user.DisplayName -ea silentlycontinue)
    {
        $foundDisplayName += $user.upn
        $foundUsers += $user.upn
        Write-Host "found*" -ForegroundColor Green
        $tmp.O365_UPN = $msolusersmtp.userprincipalname
        $tmp.O365_DisplayName = $msolusersmtp.DisplayName
        $tmp.ExistsOnO365 = $true
    }

    else
    {
        $notfoundusers += $user.upn
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    $AllUsers += $tmp
}

## RSE Found User
$RSEImport = Import-Csv "C:\Users\fred5646\Rackspace Inc\American Car Center - General\CP RSE email-mailboxes-20200511-181813.csv"

    $found = @()
    $notfound = @()

    foreach ($user in $RSEImport) {
        if (get-msoluser -searchstring $user.name) {
            $found += $user
        }
        else
        {
            $notfound += $user
        }
    }

    #### Create MSOLUSers

    foreach ($user in $notfound) {
        Write-Host "Creating User "$user.Name" ..." -ForegroundColor Cyan -NoNewline
        if ($user.DisplayName) {
            New-MsolUser -userprincipalname $user.email -displayname $user.DisplayName -FirstName $user.FirstName -LastName $user.LastName -PhoneNumber $user.BusinessPhone -password 'Vi6$q3tr5FMT'
            Write-Host "done" -ForegroundColor Green
        }
       else
       {
           Write-Host "Required Field missing. No DisplayName found" -ForegroundColor red
       }
        
    }


    New-MsolUser -userprincipalname christian.mason@americancarcenter.com -displayname "Christian Mason" -FirstName "Christian" -LastName "Mason" -password 'Vi6$q3tr5FMT'

# Update License
foreach ($user in $allusers | ? {$_.ExistsOnO365 -eq $true})
{
Write-Host "Adding License to $($user.DisplayName)..." -ForegroundColor Cyan -NoNewline
Set-MsolUserLicense -userprincipalname $user.O365_UPN -addlicenses Americancar:EXCHANGESTANDARD
Write-Host "done" -ForegroundColor Green
}

#Set Webmail forward for users

#

foreach ($webmailuser in $webmailusers)
{
    Write-Host "Updating forward for $($webmailuser.DisplayName)..." -ForegroundColor Magenta -NoNewline
    Set-Mailbox $webmailuser.EmailAddress -forwardingsmtpaddress $webmailuser.forwardingaddress -delivertomailboxandforward $true
    Write-Host "done" -ForegroundColor Green
}

#Update DL Membership with Webmail Users

$dlmemberlist = Import-Csv $HOME\desktop\DLsProperties.csv
$DLMemberResults = @()

foreach ($dl in ($dlmemberslist | ? {$_.RecipientType -eq "MailContact"}))
{
	write-host "Adding Member $($dl.MemberSMTPaddress) to DL" $dl.DistributionGroup "..." -foregroundcolor cyan
    Add-DistributionGroupMember -Identity $dl.DLPrimarySMTPaddress -Member $dl.MemberSMTPaddress
}
	
### End Of Region ###

## Remove Routing address

foreach ($user in $routingusers)
{
    $routingaddress = "smtp:" + $user.name + "@routing.americancarcenter.com"
    Write-Host "Removing Routing Address $routingaddress for $($user.name) ..." -ForegroundColor Magenta -NoNewline
    set-mailbox $user.Alias -emailaddresses @{remove=$routingaddress}
    Write-Host ". done" -ForegroundColor Green
}
#add emailaddresses
foreach ($user in $importcsv) 
{
    Write-host "Updating Email Addresses for $($user.PrimarySMTPAddress) ..." -NoNewline -fore cyan

    $AddressArray = ($user.emailaddresses -split ",")

    foreach ($address in $AddressArray)
    {
        Write-Host "Adding Address $($address) ..." -fore cyan -NoNewline
        Set-RemoteMailbox $user.PrimarySMTPAddress -EmailAddresses @{add=$address} -WhatIf
    }
Write-Host "done" -fore green
}

# match onpremise users against mail attribute 2 for American Car Center

$importcsv = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $importcsv | sort DisplayName) { 
    Write-Host "Checking for $($user.DisplayName) on premise ..." -fore Cyan -NoNewline
    $UPNLookup = $user.EmailAddress
    $tmp = "" | select DisplayName, O365UPN, PrimarySMTPAddress, EmailAddresses, ExistsOnPrem, ImmutableID, ADUPN, DistinguishedName
    $tmp.DisplayName = $user.DisplayName
    $tmp.O365UPN = $user.UserPrincipalName   
    $tmp.PrimarySMTPAddress = $user.PrimarySMTPAddress
    $tmp.EmailAddresses = $user.EmailAddresses

    $PrimarySMTP = $user.PrimarySMTPAddress
    
    if ($ADUser = Get-ADUser -filter {mail -eq $PrimarySMTP}) {
        
        $foundusers += $user
        Write-Host "found" -ForegroundColor Green
        $tmp.DistinguishedName = $ADUser.DistinguishedName
        $tmp.ADUPN = $ADUser.UserPrincipalName
        $tmp.DisplayName = $ADUser.Name
        $tmp.ExistsOnPrem = $true

        #create immutableID
        $UserimmutableID = [System.Convert]::ToBase64String(([GUID]$ADUser.ObjectGUID).ToByteArray())
        $tmp.ImmutableID = $UserimmutableID
    }

    else
    {
        $notfoundusers += $user.UserPrincipalName
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnPrem = $False
    }

    $AllUsers += $tmp
}

#Enable Remote Mailboxes

$FailedEnabledUsers = @()
$CompletedMailEnabledUsers = @()
foreach ($user in $importcsv | ?{$_.ExistsOnPrem -eq $true}) {
    if (Get-RemoteMailbox $user.PrimarySMTPAddress) {
        Write-host "Updating Email Addresses for $($user.PrimarySMTPAddress)" -NoNewline -fore cyan
        $AddressArray = ($user.emailaddresses -split ",")

        foreach ($address in $AddressArray) {
            if (Get-RemoteMailbox $user.PrimarySMTPAddress | ? {$_.EmailAddresses -like "*$($address)*"}) {
                Write-Host "Address $($address) already exists..." -fore green -NoNewline
            }
            
            else{
                Write-Host "Adding Address $($address) ..." -fore Magenta -NoNewline
                Set-RemoteMailbox $user.PrimarySMTPAddress -EmailAddresses @{add=$address}
            }
        }
    }

    else{
        Write-host "Enabling Mailbx for $($user.PrimarySMTPAddress) ..." -NoNewline -fore cyan
        if ($EnableMailbox = Enable-RemoteMailbox $user.ADUPN -RemoteRoutingAddress $user.PrimarySMTPAddress -primarysmtpaddress $user.PrimarySMTPAddress -erroraction silentlycontinue)
        {
            $EnableMailbox
            $CompletedMailEnabledUsers += $user
            Write-Host "Enabled. " -fore green -NoNewline
        }
        else
        {
            $FailedEnabledUsers  += $user
            Write-Host "Failed to enable. " -fore red -NoNewline
        }
    }   
Write-Host "done" -fore green
}