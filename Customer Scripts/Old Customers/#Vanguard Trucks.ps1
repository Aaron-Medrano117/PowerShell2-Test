#Vanguard Trucks

#Add .onmicrosoft.com address
$mailboxes = get-mailbox -OrganizationalUnit vanguardtrucks.com |sort displayName

foreach ($mbx in $mailboxes)
{
    Write-Host "Checking mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $OnMicrosoftAddressSplit = ($mbx.primarysmtpaddress) -split "@"
    $OnMicrosoftAddress = $OnMicrosoftAddressSplit[0] + "@vanguardtrucks.mail.onmicrosoft.com"
    $aliasarray = $mbx.EmailAddresses.ProxyAddressString
    if (!($aliasarray -like "*@vanguardtrucks.mail.onmicrosoft.com*"))
    {
        Write-host "Adding address $($OnMicrosoftAddress)"
        Set-Mailbox $mbx.alias -EmailAddresses @{add=$OnMicrosoftAddress}
        $updatedUsers += $mbx
    }
    else
    {
        Write-host "Microsoft domain vanguardtrucks.mail.onmicrosoft.com already on mailbox" -nonewline -fore Yellow
    } 
    Write-Host "done." -fore green
}

# remove routing address
$mailboxes = get-mailbox -OrganizationalUnit vanguardtrucks.com |sort displayName
$updatedUsers = @()

foreach ($mbx in $mailboxes)
{
    Write-Host "Checking mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $aliasarray = $mbx.EmailAddresses.ProxyAddressString
    Write-Host $aliasarray.count "aliases found . " -fore darkcyan -nonewline
    foreach ($alias in $aliasarray)
    {
        if ($alias -like "*@routing.*")
        {
            Write-host "Removing address $($alias)"
            Set-Mailbox $mbx.alias -EmailAddresses @{remove=$alias}
            $updatedUsers += $mbx
        }
        else
        {
            Write-host "." -nonewline -fore Yellow
        } 
    }
    Write-Host "done." -fore green
}

# Create RSE MSOL Users

foreach ($rsembx in $rsemailboxes)
{
    Write-Host "Creating RSE MSOLUser $($rsembx.name)"
    if (!($msolusercheck = get-msoluser -userprincipalname $rsembx.email -ea silentlycontinue))
    {
        New-MsolUser -UserPrincipalName $rsembx.email -FirstName $rsembx.FirstName -LastName $rsembx.LastName -DisplayName $rsembx.DisplayName -password (ConvertTo-SecureString -String 'Year96Temerity' -AsPlainText -Force) -Title $rsembx.Title -City $rsembx.City -State $rsembx.State -UsageLocation US
    }
    else
    {
        Write-Host "User already Exists." -ForegroundColor Yellow
    }
}
}

## Update Users with Specified licenses
foreach ($user in $ExchangeMailboxes)
{
    $msolUserCheck = get-msoluser -userprincipalname $user.Email -ea silentlycontinue
            
    if ($msolUserCheck)
    {
        $license = $($user.license)
        #Set-Msoluser -UserPrincipalName $msolUserCheck.UserPrincipalName -UsageLocation US -verbose
        Set-MsolUserLicense -UserPrincipalName $msolUserCheck.UserPrincipalName -AddLicenses $license -verbose
        Write-Host "Updated user $($msolUserCheck.DisplayName) license to $($user.license)" -ForegroundColor Green
    }
}

## Update RSE Users with Kiosk
foreach ($user in $unlicensedUsers)
{
    $license = "vanguardtrucks:EXCHANGEDESKLESS"
    #Set-Msoluser -UserPrincipalName $user.UserPrincipalName -UsageLocation US
    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $license
    Write-Host "Updated user $($user.DisplayName) license to $($license)" -ForegroundColor Green   
}

#Set RS mailboxes to forward to RS subdomain
foreach ($mbx in $rsemailboxes)
{
    $addressSplit = $mbx.Email -split "@"
    $RSForwardAddress = $addressSplit[0] + "@routing.vanguardtrucks.com"
    Write-Host "Setting $($mbx.Email) to forward to .. $($RSForwardAddress) " -fore cyan -nonewline
    Set-Mailbox $mbx.Email -ForwardingSMTPAddress $RSForwardAddress
    Write-Host "done." -fore green
}

# Create RSE MSOL Users

foreach ($rsembx in $rsemailboxes)
{
    Write-Host "Updating Password RSE MSOLUser $($rsembx.name)"
    if (!($msolusercheck = get-msoluser -userprincipalname $rsembx.email -ea silentlycontinue))
    {
        New-MsolUser -UserPrincipalName $rsembx.email -FirstName $rsembx.FirstName -LastName $rsembx.LastName -DisplayName $rsembx.DisplayName -password (ConvertTo-SecureString -String 'Year96Temerity' -AsPlainText -Force) -Title $rsembx.Title -City $rsembx.City -State $rsembx.State -UsageLocation US
    }
    else
    {
        Write-Host "User already Exists." -ForegroundColor Yellow
    }
}
}

#EnableMFA

foreach ($user in $hexfinalbatch)
{
    if ($MSOLUSER = Get-MsolUser -UserPrincipalName $user -erroraction silentlycontinue)
    {
        Write-Host "Enable MFA for user $($user)"
        $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
        $st.RelyingParty = "*"
        $st.State = "Enabled"
        $enableMFA = @($st)
         
        #Enable MFA
        Set-msoluser -UserPrincipalName $user -StrongAuthenticationRequirements $enableMFA
    }
}

