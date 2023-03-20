#match users to Office 365 users based on upn or primary smtp address
Write-Host "This script will check the users in the csv for a matching user based on the MSOL user's UPN or PrimarySMTP address." -ForegroundColor White

$filepath = Read-Host "What is the FilePath of the CSV File to Import"

$importcsv = Import-csv $filepath

$AllUsers =@()
$foundusers =@()
$notfoundusers =@()
$foundDisplayName = @()

foreach ($user in ($importcsv)) {
    Write-Host "Checking user $($user.UserPrincipalName) in Office 365 ..." -fore Cyan -NoNewline
    $UserAddress = $user.UserPrincipalName -split '@'
    $UserSuffix = $UserAddress[0]

    $tmp = "" | select DisplayName, PrimarySmtpAddress, UserPrincipalName, OrgUnitPath, FirstName, LastName, TargetDeliveryAddress, EmailAddresses, Matched365User, ExistsOnO365, Is365Mailbox
    $tmp.DisplayName = $user.DisplayName
    $tmp.PrimarySmtpAddress = $user.PrimarySmtpAddress
    $tmp.UserPrincipalName    = $user.UserPrincipalName
    $tmp.OrgUnitPath    = $user.OrgUnitPath
    $tmp.FirstName    = $user.FirstName
    $tmp.LastName    = $user.LastName
    $tmp.TargetDeliveryAddress = $UserSuffix + "@o365.proctoru.com"
    $tmp.EmailAddresses    = $user.EmailAddresses

    #Check for MSOL User
    if ($msoluserupn = get-msoluser -userprincipalname $user.UserPrincipalName -ea silentlycontinue)
    { 
        $foundusers += $msoluserupn.UserPrincipalName
        Write-Host "found" -ForegroundColor Green
        $tmp.ExistsOnO365 = $true
        $tmp.Matched365User = $msoluserupn.UserPrincipalName
    }

    elseif ($msoluserprimarysmtp = get-msoluser -userprincipalname $user.primarysmtpaddress -ea silentlycontinue)
    {
        $foundusers += $msoluserprimarysmtp.UserPrincipalName
        Write-Host "found*" -ForegroundColor Green
        $tmp.ExistsOnO365 = $true
        $tmp.Matched365User = $msoluserprimarysmtp.UserPrincipalName
    }

    elseif ($msoluserdisplay = get-msoluser -searchstring $($user.DisplayName) -ea silentlycontinue)
    {
        $foundDisplayName += $msoluserdisplay.UserPrincipalName
        $foundusers += $msoluserdisplay.UserPrincipalName
        Write-Host "found*" -ForegroundColor Green
        $tmp.ExistsOnO365 = $true
        $tmp.Matched365User = $msoluserdisplay.UserPrincipalName
    }

    else
    {
        $notfoundusers += $user.UserPrincipalName
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    ### Check For Mailbox
    if ($mailboxupn = get-mailbox $user.UserPrincipalName -ea silentlycontinue)
    {
        $tmp.Is365Mailbox = $true
        $tmp.Matched365User = $mailboxupn.name
    }

    elseif ($mailboxsmtp = get-mailbox $user.primarysmtpaddress -ea silentlycontinue)
    {
        $tmp.Is365Mailbox = $true
        $tmp.Matched365User = $mailboxsmtp.name
    }

    elseif ($mailboxdisplay = get-mailbox $($user.DisplayName) -ea silentlycontinue)
    {
        $tmp.Is365Mailbox = $true
        $tmp.Matched365User = $mailboxdisplay.name
    }

    else
    {
        $tmp.Is365Mailbox = $false
    }

    $AllUsers += $tmp
}

Write-Host "For full list of users found, Use foundusers variable"
Write-Host ""
Write-host "For Full list of non-matched users, Use notfoundusers variable"

$AllUsers | Export-Csv $filepath -Encoding utf8 -NoTypeInformation




