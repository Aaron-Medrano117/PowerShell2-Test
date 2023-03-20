# ProctorU Create Mail Users

<#

Requirements

Subdomain routing to GSUITE
Subdomain routing to Office 365

New Users will need:
External Address Points to GSUITE address
USERID includes primary domain

During the migration, the target domain points to Office365 subdomain

#>

#Add Mail User

$createdusers = @()
$failedusers = @()
$alreadyexist = @()
foreach ($user in $Phase1Users) 
{
    $DisplayName = $user.FirstName + " " + $user.LastName
    $alias = $user.EmailAddress -split "@"
    $ExternalAddress = $alias[0] + "@gsuite.proctoru.com"
    $UserID = $alias[0] + "@meazurelearning.com"
    $PW = (ConvertTo-SecureString -String 'Boulder92Belief' -AsPlainText -Force)
    
    if ($MailRecipient = Get-Recipient $user.EmailAddress -erroraction silentlycontinue)
    {
        Write-Host "User $($DisplayName) already exists" -ForegroundColor Green
        $alreadyexist += $user
    }
    else
    {
        if ($NewMailUser = New-MailUser -Alias $alias[0] -ExternalEmailAddress $ExternalAddress -Name $DisplayName -MicrosoftOnlineServicesID $UserID -Password $PW)
        {
            $NewMailUser
            $createdusers += $user
        }
        else
        {
            $failedusers += $user
        }
    }   
}

#Add proctoru.com domain and reset at next logon.

-ResetPasswordOnNextLogon $true 

foreach ($user in $Phase1Users)
{
    $DisplayName = $user.FirstName + " " + $user.LastName
    $alias = $user.EmailAddress -split "@"
    $SMTPAliasAddress = "smtp:" + $alias[0] + "@proctoru.com"
    $UserID = $alias[0] + "@meazurelearning.com"
    
    Write-Host "Adding ProctorU domain for $($DisplayName) ..." -ForegroundColor Cyan -NoNewline
    Set-MailUser -identity $UserID -EmailAddresses @{add=$SMTPAliasAddress}
   
    Write-Host "Adding Free Team's License ..." -ForegroundColor Cyan -NoNewline
    Set-MsolUser -userprincipalname $UserID -usagelocation US
    Set-MsolUserLicense -userprincipalname $UserID -addlicenses getyardstick:TEAMS_COMMERCIAL_TRIAL
    Write-Host "done" -ForegroundColor Green
}

foreach ($user in $Phase1Users)
{
    $DisplayName = $user.FirstName + " " + $user.LastName
    $alias = $user.EmailAddress -split "@"
    $UserID = $alias[0] + "@meazurelearning.com"
       
    Write-Host "Adding Free Team's License ..." -ForegroundColor Cyan -NoNewline
    Set-MsolUser -userprincipalname $UserID -usagelocation US
    Set-MsolUserLicense -userprincipalname $UserID -addlicenses getyardstick:TEAMS_COMMERCIAL_TRIAL
    Write-Host "done" -ForegroundColor Green
}

foreach ($user in $Phase1Users)
{
    $DisplayName = $user.FirstName + " " + $user.LastName
    $alias = $user.EmailAddress -split "@"
    $SMTPAliasAddress = "smtp:" + $alias[0] + "@proctoru.com"
    $UserID = $alias[0] + "@meazurelearning.com"
    
    Write-Host "Remove ProctorU domain for $($DisplayName) ..." -ForegroundColor Cyan -NoNewline
    Set-MailUser -identity $UserID -EmailAddresses @{remove=$SMTPAliasAddress}
    Write-Host "done" -ForegroundColor Green
}