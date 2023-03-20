#match onpremise users

$importcsv = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $importcsv) { 
    Write-Host "Checking user $($user.DisplayName) onPremise ..." -fore Cyan -NoNewline
    [string]$UPNLookup = $user.New_Username + "@showrig.net"
    $NewUPN = $user.New_Username + "@sgps.net"

    $tmp = "" | select ExistsOnPrem, NewUPN, CSVDisplayName, DesiredOU, ADUPN, Name, DistinguishedName
    $tmp.CSVDisplayName = $user.DisplayName
    $tmp.NewUPN = $NewUPN
    $tmp.DesiredOU = $user.OU

    if ($ADUser = Get-AdUser -filter {UserPrincipalName -eq $UPNLookup}) {
        
        $foundusers += $NewUPN
        Write-Host "found" -ForegroundColor Green
        $tmp.DistinguishedName = $ADUser.DistinguishedName
        $tmp.ADUPN = $ADUser.UserPrincipalName
        $tmp.Name = $ADUser.Name
        $tmp.ExistsOnPrem = $true
    }

    else
    {
        $notfoundusers += $user.upn
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnPrem = $False
    }

    $AllUsers += $tmp
}

Write-Host "For full list of users found, Use foundusers variable"
Write-Host ""
Write-host "For Full list of non-matched users, Use notfoundusers variable"

$AllUsers | Export-Csv $filepath -Encoding utf8 -NoTypeInformation

#match onpremise users & Create Immutable ID

$importcsv = Import-csv $filepath

$allUsers =@()

foreach ($user in $importcsv) { 
    Write-Host "Checking user $($user.DisplayName) onPremise ..." -fore Cyan -NoNewline
    [string]$UPNLookup = $user.ADUPN

    $tmp = "" | select ExistsOnPrem, Displayname, NewUPN, email_aliases, DesiredOU, ADUPN, DistinguishedName, ObjectGUID, ImmutableID
    $tmp.ExistsOnPrem = $user.ExistsOnPrem
    $tmp.NewUPN    = $user.NewUPN
    $tmp.email_aliases = $user.email_aliases
    $tmp.DesiredOU = $user.DesiredOU
    $tmp.ADUPN = $user.ADUPN

    $ADUser = Get-AdUser -filter {UserPrincipalName -eq $UPNLookup}
        
    $tmp.DistinguishedName = $ADUser.DistinguishedName
    $tmp.Displayname = $ADUser.name
    $tmp.ObjectGUID = $ADUser.ObjectGUID

    #create immutableID
    $UserimmutableID = [System.Convert]::ToBase64String($ADUser.ObjectGUID.tobytearray())
    $tmp.ImmutableID = $UserimmutableID
    Write-Host "done"

    $AllUsers += $tmp
}


$AllUsers | Export-Csv $filepath -Encoding utf8 -NoTypeInformation



#set immutable id

$importcsv = Import-csv $filepath

foreach ($user in ($importcsv | ? {$_.ExistsOnO365})) {
    $upn = $user.O365_UPN
    Write-Host "updating immutableid for user $upn"
    Set-MsolUser -userprincipalname $upn -immutableid $user.ImmutableID
}

###
$MatchedUsers = Import-Csv $filepath

foreach ($user in ($MatchedUsers | ? {$_.ExistsOnO365 -eq $true})) {
    $upn = $user.O365_UPN
    Write-Host "updating immutableid for user $upn"
    Set-MsolUser -userprincipalname $upn -immutableid $user.ImmutableID
}

foreach ($user in ($MatchedUsers | ? {$_.ExistsOnO365 -eq $true})) {
    $upn = $user.O365_UPN
    #Write-Host "gathering immutableid for user $upn"
    Get-MsolUser -userprincipalname $upn | select DisplayName, UserPrincipalName, ImmutableID
}


#### CCMSI ####

#match onpremise users

$importcsv = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $importcsv) { 
    Write-Host "Checking user $($user.HEX_DisplayName) onPremise ..." -fore Cyan -NoNewline
    $UPNLookup = $user.HEX_UPN

    $tmp = "" | select HEX_DisplayName,HEX_UPN, HEX_PrimarySmtpAddress, ExistsOnO365, O365_DisplayName, O365_UPN, ExistsOnPrem, ADUPN, Name, DistinguishedName
    $tmp.HEX_DisplayName = $user.HEX_DisplayName
    $tmp.HEX_UPN  = $user.HEX_UPN
    $tmp.HEX_PrimarySmtpAddress = $user.HEX_PrimarySmtpAddress
    $tmp.ExistsOnO365 = $user.ExistsOnO365
    $tmp.O365_DisplayName = $user.O365_DisplayName
    $tmp.O365_UPN = $user.O365_UPN

    if ($ADUser = Get-AdUser -filter {UserPrincipalName -eq $UPNLookup}) {
        
        $foundusers += $user.HEX_UPN
        Write-Host "found" -ForegroundColor Green
        $tmp.DistinguishedName = $ADUser.DistinguishedName
        $tmp.ADUPN = $ADUser.UserPrincipalName
        $tmp.Name = $ADUser.Name
        $tmp.ExistsOnPrem = $true
    }

    else
    {
        $notfoundusers += $user.HEX_UPN
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnPrem = $False
    }

    $AllUsers += $tmp
}

Write-Host "For full list of users found, Use foundusers variable"
Write-Host ""
Write-host "For Full list of non-matched users, Use notfoundusers variable"

$AllUsers | Export-Csv $filepath -Encoding utf8 -NoTypeInformation


$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $importcsv) {
    Write-Host "Checking user $($user.HEX_UPN) in Office 365 ..." -fore Cyan -NoNewline
    $tmp = "" | select HEX_DisplayName,HEX_UPN, HEX_PrimarySmtpAddress, ExistsOnO365, O365_DisplayName, O365_UPN, ExistsOnPrem, ADUPN, DistinguishedName
    $tmp.HEX_DisplayName = $user.HEX_DisplayName
    $tmp.HEX_UPN = $user.HEX_UPN
    $tmp.HEX_PrimarySmtpAddress = $user.HEX_PrimarySmtpAddress,
    $tmp.ExistsOnPrem = $user.ExistsOnPrem
    $tmp.ADUPN = $user.ADUPN
    $tmp.DistinguishedName = $user.DistinguishedName


    if ($msoluser = get-msoluser -userprincipalname $user.HEX_UPN -ea silentlycontinue)
    { 
        $foundusers += $user.HEX_UPN
        Write-Host "found" -ForegroundColor Green
        $tmp.O365_UPN = $msoluser.userprincipalname
        $tmp.O365_DisplayName = $msoluser.DisplayName
        $tmp.ExistsOnO365 = $true
    }

    elseif ($msolusersmtp = Get-msoluser -userprincipalname $user.HEX_PrimarySmtpAddress -ea silentlycontinue)
    {
        $foundDisplayName += $user.HEX_UPN
        $foundUsers += $user.HEX_UPN
        Write-Host "found*" -ForegroundColor Green
        $tmp.O365_UPN = $msolusersmtp.userprincipalname
        $tmp.O365_DisplayName = $msolusersmtp.DisplayName
        $tmp.ExistsOnO365 = $true
    }

    else
    {
        $notfoundusers += $user.HEX_UPN
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    $AllUsers += $tmp
}