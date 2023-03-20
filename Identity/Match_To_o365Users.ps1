#check msol user based on upn
#if not
# check based on smtp address
Write-Host "This script will check the users in the csv for a matching user based on the MSOL user's UPN or PrimarySMTP address." -ForegroundColor White

$filepath = Read-Host "What is the FilePath of the CSV File to Import"

$importcsv = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()
$foundDisplayName = @()

foreach ($user in ($importcsv | ? {$_.upn})) {
    Write-Host "Checking user $($user.upn) in Office 365 ..." -fore Cyan -NoNewline
    $tmp = "" | select HEX_DisplayName, HEX_UPN, HEX_PrimarySmtpAddress, ExistsOnO365, O365_DisplayName, O365_UPN 
    $tmp.HEX_DisplayName = $user.DisplayName
    $tmp.HEX_UPN = $user.upn
    $tmp.HEX_PrimarySmtpAddress = $user.PrimarySmtpAddress

    if ($msoluser = get-msoluser -userprincipalname $user.upn -ea silentlycontinue)
    { 
        $foundusers += $user.upn
        Write-Host "found" -ForegroundColor Green
        $tmp.O365_UPN = $msoluser.userprincipalname
        $tmp.O365_DisplayName = $msoluser.DisplayName
        $tmp.ExistsOnO365 = $true
    }

    elseif ($msolusersmtp = Get-msoluser -userprincipalname $user.primarysmtpaddress -ea silentlycontinue)
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

Write-Host "For full list of users found, Use foundusers variable"
Write-Host ""
Write-host "For Full list of non-matched users, Use notfoundusers variable"

$AllUsers | Export-Csv $filepath -Encoding utf8 -NoTypeInformation

### Check 

#check msol user based on upn
#if not
# check based on display name
Write-Host "This script will check the users in the csv for a matching user based on the MSOL user's UPN or PrimarySMTP address." -ForegroundColor White

$filepath = Read-Host "What is the FilePath of the CSV File to Import"

$importcsv = Import-csv $filepath

$allUsers =@()
$foundUsers =@()
$notFoundUsers =@()

foreach ($user in $importcsv) {
    Write-Host "Checking user $($user.Displayname) in Office 365 ..." -fore Cyan -NoNewline
    
    $tmp = "" | select ExistsOnPrem, Displayname, NewUPN, email_aliases, DesiredOU, ADUPN, DistinguishedName, ObjectGUID, ImmutableID, ExistsOnO365, O365_DisplayName, O365_UPN
    $tmp.ExistsOnPrem = $user.ExistsOnPrem
    $tmp.NewUPN    = $user.NewUPN
    $tmp.email_aliases = $user.email_aliases
    $tmp.DesiredOU = $user.DesiredOU
    $tmp.ADUPN = $user.ADUPN
    $tmp.DistinguishedName = $user.DistinguishedName
    $tmp.Displayname = $user.Displayname
    $tmp.ObjectGUID = $user.ObjectGUID
    $tmp.ImmutableID = $user.ImmutableID

    #create immutableID

    if ($msoluserUPN = get-msoluser -userprincipalname $user.NewUPN -ea silentlycontinue)
    { 
        $foundusers += $user.NewUPN
        Write-Host "found" -ForegroundColor Green
        $tmp.O365_UPN = $msoluserUPN.userprincipalname
        $tmp.O365_DisplayName = $msoluserUPN.DisplayName
        $tmp.ExistsOnO365 = $true
    }

    elseif ($msoluser = Get-msoluser -searchstring $user.DisplayName -ea silentlycontinue)
    {
        $foundUsers += $user.NewUPN
        Write-Host "found*" -ForegroundColor Green
        $tmp.O365_UPN = $msoluser.userprincipalname
        $tmp.O365_DisplayName = $msoluser.DisplayName
        $tmp.ExistsOnO365 = $true
    }

    else
    {
        $notfoundusers += $user.NewUPN
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    $AllUsers += $tmp
}

$AllUsers | Export-Csv $filepath -Encoding utf8 -NoTypeInformation


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

Write-Host "For full list of users found, Use foundusers variable"
Write-Host ""
Write-host "For Full list of non-matched users, Use notfoundusers variable"

$AllUsers | Export-Csv $filepath -Encoding utf8 -NoTypeInformation
