# Match GSUITE USERS

<#

Requirements

Subdomain routing to GSUITE
Subdomain routing to Office 365

New Users will need:
External Address Points to GSUITE address
USERID includes primary domain

During the migration, the target domain points to Office365 subdomain

#>

# Match users below:

$importCsv = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\GSuiteUsers.csv"

foreach ($MailUser in $importCsv)
{
    Write-Host "Checking user $($user.upn) in Office 365 ..." -fore Cyan -NoNewline
    $tmp = "" | select DisplayName, PrimarySmtpAddress, UserPrincipalName, OrgUnitPath, FirstName, LastName, TargetDeliveryAddress, EmailAddresses, Matched365User, ExistsOnO365, Is365Mailbox
    $tmp.DisplayName = $user.DisplayName
    $tmp.PrimarySmtpAddress = $user.PrimarySmtpAddress
    $tmp.UserPrincipalName    = $user.UserPrincipalName
    $tmp.OrgUnitPath    = $user.OrgUnitPath
    $tmp.FirstName    = $user.FirstName
    $tmp.LastName    = $user.LastName
    $tmp.EmailAddresses    = $user.EmailAddresses
    $AllUsers += $tmp


    if (get-msoluser -userprincipalname $user.UserPrincipalName -ea silentlycontinue)
    { 
        $foundusers += $user.upn
        Write-Host "found" -ForegroundColor Green
        $tmp.ExistsOnO365 = $true
    }

    elseif (get-msoluser -userprincipalname $user.primarySmtpAddress -ea silentlycontinue)
    {
        $foundusers += $user.upn
        Write-Host "found*" -ForegroundColor Green
        $tmp.ExistsOnO365 = $true
    }

    else
    {
        $notfoundusers += $user.upn
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }
}


# Match Users Script 2

# Match users below:
$importCsv = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\GSuiteUsers.csv"

$AllUsers = @()
$foundusers = @()
$notfoundusers = @()

foreach ($User in $importCsv)
{
    $SplitPrimarySMTP = $user.PrimarySmtpAddress -split '@'
    $365RoutingAlias = $SplitPrimarySMTP[0] + "@o365.proctoru.com"
    $ExternalEmailAddress = $SplitPrimarySMTP[0] + "@gsuite.proctoru.com"
    $YardstickUPN = $SplitPrimarySMTP[0] + "@getyardstick.com"
    
    Write-Host "Checking user $($user.UserPrincipalName) in Office 365 ..." -fore Cyan -NoNewline
    $tmp = "" | select DisplayName, PrimarySmtpAddress, UserPrincipalName, OrgUnitPath, FirstName, LastName, EmailAddresses, ExistsOnO365, Is365MailObject, Matched365User, Match365_DisplayName, RoutingAlias, ExternalEmailAddress, Desired365UPN 
    $tmp.DisplayName = $user.DisplayName
    $tmp.PrimarySmtpAddress = $user.PrimarySmtpAddress
    $tmp.UserPrincipalName    = $user.UserPrincipalName
    $tmp.OrgUnitPath    = $user.OrgUnitPath
    $tmp.FirstName    = $user.FirstName
    $tmp.LastName    = $user.LastName
    $tmp.EmailAddresses    = $user.EmailAddresses
    $tmp.ExternalEmailAddress = $ExternalEmailAddress
    $tmp.RoutingAlias = $365RoutingAlias

    # Check for Azure User #
    if ($MSOLUser = get-msoluser -userprincipalname $user.UserPrincipalName -ea silentlycontinue)
    { 
        $foundusers += $user.UserPrincipalName
        Write-Host "found" -ForegroundColor Green
        $tmp.ExistsOnO365 = $true
        $tmp.Matched365User = $MSOLUser.UserPrincipalName
        $tmp.Match365_DisplayName = $MSOLUser.DisplayName
    }

    elseif ($MSOLUser2 = get-msoluser -userprincipalname $user.primarySmtpAddress -ea silentlycontinue)
    {
        $foundusers += $user.UserPrincipalName
        Write-Host "found*" -ForegroundColor Green
        $tmp.ExistsOnO365 = $true
        $tmp.Matched365User = $MSOLUser2.UserPrincipalName
        $tmp.Match365_DisplayName = $MSOLUser2.DisplayName
    }

    elseif ($MSOLUser3 = Get-msoluser -searchstring $user.DisplayName -ea silentlycontinue)
    {
        $foundUsers += $user.UserPrincipalName
        Write-Host "found*" -ForegroundColor Green
        $tmp.Matched365User = $MSOLUser3.UserPrincipalName
        $tmp.Match365_DisplayName = $MSOLUser3.DisplayName
        $tmp.ExistsOnO365 = $true
    }

    else
    {
        $notfoundusers += $user.UserPrincipalName
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    # Check for Mail Object #

    if (Get-Recipient $user.Matched365User -ea silentlycontinue)
    {
        $tmp.Is365MailObject = $true
    }

    else
    {
        $tmp.Is365MailObject = $false
    }

    if ($user.OrgUnitPath -eq "/Yardstick")
    {
        $tmp.Desired365UPN = $YardstickUPN
    }

    $AllUsers += $tmp
}