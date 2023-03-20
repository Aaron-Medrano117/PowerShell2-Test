#The Room Place

foreach ($user in $rseuserOverQuota[0])
{
    $msolUserCheck = get-msoluser -userprincipalname $user
    
    #Update Licenses
    $DisabledArray_Deskless = "EXCHANGE_S_DESKLESS"
    $DisabledArray_BP = "EXCHANGE_S_STANDARD"
        
    if ($msolUserCheck.licenses.AccountSkuId -eq "reseller-account:DESKLESSPACK")
    {
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $msolUserCheck.licenses.AccountSkuId -DisabledPlans $DisabledArray_Deskless -Verbose
        Set-MsolUserLicense -UserPrincipalName $msolUserCheck.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
        Write-Host "Updated user $($msolUserCheck.DisplayName) license to $($msolUserCheck.licenses.AccountSkuId) with disabled $($DisabledArray_Deskless)" -ForegroundColor Green
    }
    elseif ($msolUserCheck.licenses.AccountSkuId -eq "reseller-account:O365_BUSINESS_PREMIUM")
    {
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $msolUserCheck.licenses.AccountSkuId -DisabledPlans $DisabledArray_BP -Verbose
        #Set-MsolUserLicense -UserPrincipalName $msolUserCheck.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
        Write-Host "Updated user $($msolUserCheck.DisplayName) license to $($msolUserCheck.licenses.AccountSkuId) with disabled $($DisabledArray_BP)"  -ForegroundColor Green 
    }
    elseif ($msolUserCheck.licenses.AccountSkuId -eq "reseller-account:O365_BUSINESS_ESSENTIALS")
    {
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $msolUserCheck.licenses.AccountSkuId -DisabledPlans $DisabledArray_BP -Verbose
        Set-MsolUserLicense -UserPrincipalName $msolUserCheck.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
        Write-Host "Updated user $($msolUserCheck.DisplayName) license to $($msolUserCheck.licenses.AccountSkuId) with disabled $($DisabledArray_BP)"  -ForegroundColor Green 
    }
}



## Update Users over quota with Business Basic license
foreach ($user in $rseuserOverQuota)
{
    $msolUserCheck = get-msoluser -userprincipalname $user
            
    if ($msolUserCheck.licenses.AccountSkuId -eq "reseller-account:DESKLESSPACK")
    {
        Set-MsolUserLicense -UserPrincipalName $msolUserCheck.UserPrincipalName -RemoveLicenses "reseller-account:DESKLESSPACK" -AddLicenses "reseller-account:O365_BUSINESS_ESSENTIALS" -verbose
        Write-Host "Updated user $($msolUserCheck.DisplayName) license to reseller-account:O365_BUSINESS_ESSENTIALS" -ForegroundColor Green
    }
}

#
#Set RS mailboxes to forward to RS subdomain
foreach ($mbx in $rseuserOverQuota)
{
    $addressSplit = $mbx -split "@"
    $RSForwardAddress = $addressSplit[0] + "@rs.theroomplace.com"
    Write-Host "Setting $($mbx) to forward to .. $($RSForwardAddress) " -fore cyan -nonewline
    Set-Mailbox $mbx -ForwardingSMTPAddress $RSForwardAddress
    Write-Host "done." -fore green
}

#Remove routing address
$mailboxes = get-mailbox -OrganizationalUnit theroomplace.com |sort displayName
$updatedUsers = @()

foreach ($mbx in $mailboxes)
{
    Write-Host "Checking mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $aliasarray = $mbx.EmailAddresses.ProxyAddressString
    Write-Host $aliasarray.count "aliases found . " -fore darkcyan -nonewline
    foreach ($alias in $aliasarray)
    {
        if ($alias -like "*@rs.*")
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

#Add OnMicrosoftAddress address
$mailboxes = get-mailbox -OrganizationalUnit theroomplace.com |sort displayName
$updatedUsers = @()

foreach ($mbx in $mailboxes)
{
    Write-Host "Checking mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $OnMicrosoftAddressSplit = ($mbx.primarysmtpaddress) -split "@"
    $OnMicrosoftAddress = $OnMicrosoftAddressSplit[0] + "@trpacqinc.mail.onmicrosoft.com"
    $aliasarray = $mbx.EmailAddresses.ProxyAddressString
    if (!($aliasarray -contains "*@trpacqinc.mail.onmicrosoft.com"))
    {
        Write-host "Adding address $($OnMicrosoftAddress)"
        Set-Mailbox $mbx.alias -EmailAddresses @{add=$OnMicrosoftAddress}
        $updatedUsers += $mbx
    }
    else
    {
        Write-host "Microsoft domain trpacqinc.mail.onmicrosoft.com already on mailbox" -nonewline -fore Yellow
    } 
    Write-Host "done." -fore green
}

## Check if User exists
$foundusers = @()
$notfoundusers = @()
foreach ($mbx in $rsemailboxes)
{
    $displayName = $mbx.DisplayName
    $UPN = $mbx.Email
    Write-Host "Checking for $($mbx.DisplayName) .." -fore cyan -nonewline
    if ($msolUserCheck = Get-MsolUser -SearchString $displayName)
    {
        Write-Host "found" -fore green
        $foundusers += $mbx
    }
    elseif ($msolUserCheck = Get-MsolUser -SearchString $UPN)
    {
        Write-Host "found*" -fore green
        $foundusers += $mbx
    }
    else
    {
        Write-Host "not found" -fore red
        $notfoundusers += $mbx
    }   
}

#create rsemailbox
foreach ($rsembx in $notfoundusers)
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

#Remove forward to RS subdomain
foreach ($mbx in $forwardmailboxes)
{
    Write-Host "Removing forward on $($mbx.primarysmtpaddress) " -fore cyan -nonewline
    Set-Mailbox $mbx.primarysmtpaddress -ForwardingSMTPAddress $null
    Write-Host "done." -fore green
}

#Get users not logged in today and reset password to Welcome1

$mailboxes = Get-Mailbox -ResultSize Unlimited | Where{$_.DisplayName -notlike "Discovery Search Mailbox"}
$MailboxDetails =@()

$progressref = ($mailboxes).count
$progresscounter = 0
foreach ($mailbox in $mailboxes)
{
    $lastlogintime = (Get-MailboxStatistics $mailbox.primarysmtpaddress).LastLogonTime
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($mailbox.DisplayName)"

    $mailboxstats = New-Object -TypeName PSObject
    $mailboxstats | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName
    $mailboxstats | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName
    $mailboxstats | Add-Member -MemberType NoteProperty -Name LastLoginTime -Value $lastlogintime
    $MailboxDetails += $mailboxstats
}

$newUsers = $MailboxDetails | ?{$_.lastlogintime -notlike "*7/21/2021*"}


#ProgressBar
$progressref = ($resetusers2).count
$progresscounter = 0
foreach ($user in $resetusers2)
{
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Password for for $($user)"
    #$upn = $user.UserPrincipalName
    Set-MsolUserPassword -UserPrincipalName $user -ForceChangePassword:$true -NewPassword "Welcome1"
}


#Set AD User

#ProgressBar

$office365_mailboxes = Import-Csv C:\Users\raxadmin\Desktop\Office365_MailboxProperties.csv
$onpremusers = $office365_mailboxes | ?{$_.existsOnPrem -eq $true}

$progressref = ($onpremusers).count
$progresscounter = 0
foreach ($user in $onpremusers)
{
    $distinguishedname = $null
    $RemoteRoutingAddress = $null
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Enabling Remote Mailbox for $($user.DisplayName)"

    $distinguishedname = $user.DistinguishedName
    $RemoteRoutingAddress = $user.PrimarySmtpAddress.replace("@theroomplace.com","@trpacqinc.mail.onmicrosoft.com")
    
    if (!(Get-RemoteMailbox $user.name))
    {
        Write-host "Creating RemoteMailbox .." -foregroundcolor green -nonewline
        Enable-RemoteMailbox $distinguishedname -RemoteRoutingAddress $RemoteRoutingAddress -EmailAddressPolicyEnabled:$false
    }    
    Write-host "Updating RemoteMailbox .." -foregroundcolor yellow -nonewline
    $emailAddressarray = $user.EmailAddresses -split ","
    Write-Host "Found $($emailaddressarray.count) for $($user.DisplayName) " -nonewline -foregroundcolor cyan
    start-sleep -Milliseconds 60
    foreach ($alias in $emailAddressarray)
    {
        Set-RemoteMailbox $user.name -emailaddresses @{add=$alias} -warningaction silentlycontinue
        Write-Host ". "  -foregroundcolor yellow -nonewline
    }
    Set-RemoteMailbox $user.name -alias $user.alias -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($list.HiddenFromAddressListsEnabled)) -wa silentlycontinue
    Write-host "Updated AD Attributes for user" -foregroundcolor green
}


#Set AD User

#ProgressBar

$office365_mailboxes = Import-Csv C:\Users\raxadmin\Desktop\Office365_MailboxProperties.csv
$onpremusers = $office365_mailboxes | ?{$_.existsOnPrem -eq $true}

$progressref = ($onpremusers).count
$progresscounter = 0
foreach ($user in $onpremusers)
{
    $distinguishedname = $null
    $RemoteRoutingAddress = $null
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Enabling Remote Mailbox for $($user.DisplayName)"

    $distinguishedname = $user.DistinguishedName
    $RemoteRoutingAddress = $user.PrimarySmtpAddress.replace("@theroomplace.com","@trpacqinc.mail.onmicrosoft.com")
    
    if (!(Get-RemoteMailbox $user.name))
    {
        Write-host "Creating RemoteMailbox .." -foregroundcolor green -nonewline
        Enable-RemoteMailbox $distinguishedname -RemoteRoutingAddress $RemoteRoutingAddress
        $emailAddressarray = $user.EmailAddresses -split ","
        Write-Host "Found $($emailaddressarray.count) for $($user.DisplayName) " -nonewline -foregroundcolor cyan
        start-sleep -Milliseconds 60
        Set-RemoteMailbox $RemoteRoutingAddress -EmailAddressPolicyEnabled:$false
        foreach ($alias in $emailAddressarray)
        {
            Set-RemoteMailbox $RemoteRoutingAddress -emailaddresses @{add=$alias} -warningaction silentlycontinue
            Write-Host ". "  -foregroundcolor yellow -nonewline
        }
    }       
    Set-RemoteMailbox $user.PrimarySmtpAddress -alias $user.alias -primarysmtpaddress $user.PrimarySMTPAddress -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($list.HiddenFromAddressListsEnabled)) -wa silentlycontinue -EmailAddressPolicyEnabled:$false
}

#Set AD User

#ProgressBar

$office365_mailboxes = Import-Csv C:\Users\raxadmin\Desktop\Office365_MailboxProperties.csv
$onpremusers = $office365_mailboxes | ?{$_.existsOnPrem -eq $true}

$progressref = ($onpremusers).count
$progresscounter = 0
foreach ($user in $onpremusers)
{
    $distinguishedname = $null
    $RemoteRoutingAddress = $null
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Enabling Remote Mailbox for $($user.DisplayName)"
    $RemoteRoutingAddress = $user.PrimarySmtpAddress.replace("@theroomplace.com","@trpacqinc.mail.onmicrosoft.com")
 
    Set-RemoteMailbox $user.name -RemoteRoutingAddress $user.PrimarySmtpAddress

    Write-host "Updated AD Attributes for $($user.DisplayName)" -foregroundcolor green
}

#find missing users
$progressref = ($onpremusers).count
$progresscounter = 0
$missingUsers2 = @()
foreach ($user in $onpremusers)
{
    $distinguishedname = $null
    $RemoteRoutingAddress = $null
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Enabling Remote Mailbox for $($user.DisplayName)"
    
    if (!(Get-RemoteMailbox $user.PrimarySmtpAddress -ea silentlycontinue))
    {
       Write-host "No User found for $($user.DisplayName)" -foregroundcolor red
       $missingUsers2 += $user
    }
}


#Create missing users

$progressref = ($missingUsers2).count
$progresscounter = 0
foreach ($user in $missingUsers2)
{
    $distinguishedname = $null
    $RemoteRoutingAddress = $null
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Enabling Remote Mailbox for $($user.DisplayName)"

    $distinguishedname = $user.DistinguishedName
    $RemoteRoutingAddress = $user.PrimarySmtpAddress.replace("@theroomplace.com","@trpacqinc.mail.onmicrosoft.com")
    
    Write-host "Creating RemoteMailbox $($user.DisplayName) .." -foregroundcolor green -nonewline
    Enable-RemoteMailbox $distinguishedname -RemoteRoutingAddress $RemoteRoutingAddress -primarysmtpaddress $user.PrimarySmtpAddress
    $emailAddressarray = $user.EmailAddresses -split ","
    Write-Host "Found $($emailaddressarray.count) " -nonewline -foregroundcolor cyan
    start-sleep -Milliseconds 60
    Set-RemoteMailbox $RemoteRoutingAddress -EmailAddressPolicyEnabled:$false -name $user.name
  
    Set-RemoteMailbox $user.PrimarySmtpAddress -alias $user.alias -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($list.HiddenFromAddressListsEnabled)) -wa silentlycontinue -EmailAddressPolicyEnabled:$false
}



#Update Primary Remote Addresses

$progressref = ($onpremusers).count
$progresscounter = 0
foreach ($user in $onpremusers)
{
    $distinguishedname = $null
    $RemoteRoutingAddress = $null
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Update RemoteMailbox for $($user.DisplayName)"

    $distinguishedname = $user.DistinguishedName
    $RemoteRoutingAddress = $user.PrimarySmtpAddress.replace("@theroomplace.com","@trpacqinc.mail.onmicrosoft.com")
    Set-RemoteMailbox $user.displayname -EmailAddressPolicyEnabled:$false -name $user.name
    Set-RemoteMailbox $user.displayname -alias $user.alias -primarysmtpaddress $user.PrimarySmtpAddress -name $user.name -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($list.HiddenFromAddressListsEnabled)) -wa silentlycontinue -EmailAddressPolicyEnabled:$false

    Write-host "Creating RemoteMailbox $($user.DisplayName) .." -foregroundcolor green -nonewline
}

#Update Mail users
$office365_mailboxes = Import-Csv C:\Users\raxadmin\Desktop\Office365_MailboxProperties.csv
$onpremusers = $office365_mailboxes | ?{$_.existsOnPrem -eq $true}
$onpremusers = $onpremusers | ? {$_.RecipientTypeDetails -eq "UserMailbox"}

$progressref = ($onpremusers).count
$progresscounter = 0
$missingremotemailbox = @()
foreach ($user in $onpremusers)
{
    $distinguishedname = $null
    $RemoteRoutingAddress = $null
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Remote Mailbox for $($user.DisplayName)"

    $distinguishedname = $user.DistinguishedName
    $RemoteRoutingAddress = $user.PrimarySmtpAddress.replace("@theroomplace.com","@trpacqinc.mail.onmicrosoft.com")
    
    if (!(Get-RemoteMailbox $user.PrimarySmtpAddress))
    {
        Write-host "Creating RemoteMailbox  .." -foregroundcolor green -nonewline
        Enable-RemoteMailbox $distinguishedname -RemoteRoutingAddress $RemoteRoutingAddress -primarysmtpaddress $user.primarysmtpaddress
        $missingremotemailbox += $user
        
    }
    Else
    {
        $emailAddressarray = $user.EmailAddresses -split ","
        Write-Host "Updating Remote Mailbox $($user.DisplayName)" -nonewline -foregroundcolor cyan
        start-sleep -Milliseconds 60
        Set-RemoteMailbox $user.PrimarySmtpAddress -EmailAddressPolicyEnabled:$false -name $user.name
        foreach ($alias in $emailAddressarray)
        {
            Set-RemoteMailbox $RemoteRoutingAddress -emailaddresses @{add=$alias} -warningaction silentlycontinue
            Write-Host ". "  -foregroundcolor yellow -nonewline
        }
        Set-RemoteMailbox $user.PrimarySmtpAddress -alias $user.alias -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($list.HiddenFromAddressListsEnabled)) -wa silentlycontinue -EmailAddressPolicyEnabled:$false
    }      
}



#Remove alias address address
$updatedUsers = @()

foreach ($mbx in $remotemailboxes)
{
    Write-Host "Checking mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $aliasarray = $mbx.EmailAddresses.ProxyAddressString
    Write-Host $aliasarray.count "aliases found . " -fore darkcyan -nonewline
    foreach ($alias in $aliasarray)
    {
        if ($alias -like "*x400*")
        {
            Write-host "Removing address $($alias)"
            Set-RemoteMailbox $mbx.alias -EmailAddresses @{remove=$alias}
            $updatedUsers += $mbx
        }
        else
        {
            Write-host "." -nonewline -fore Yellow
        } 
    }
    Write-Host "done." -fore green
}


#stamp ImmutableID
$progressref = ($mailboxproperties).count
$progresscounter = 0
$failedupdate = @()
foreach ($mailbox in $mailboxproperties)
{
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating ImmutableID for $($mailbox.DisplayName)"

    
    try
    {
        Set-MsolUser -UserPrincipalName $mailbox.O365UPN -ImmutableID $mailbox.ImmutableID
        Write-Host "Updated ImmutableID to $($mailbox.ImmutableID) for $($mailbox.O365UPN)" -foregroundcolor green
    }
    catch
    {
       $failedupdate += $mailbox 
    } 
}

#The Room Place

foreach ($user in $msolusers)
{
    $msolUserCheck = get-msoluser -userprincipalname $user.userprincipalname
    
    #Update Licenses
        
    if ($msolUserCheck.licenses.AccountSkuId -eq "reseller-account:DESKLESSPACK")
    {
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $msolUserCheck.licenses.AccountSkuId
        Set-MsolUserLicense -UserPrincipalName $msolUserCheck.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
        Write-Host "Updated license $($msolUserCheck.licenses.AccountSkuId) for user $($msolUserCheck.DisplayName)" -ForegroundColor Green
    }
    elseif ($msolUserCheck.licenses.AccountSkuId -eq "reseller-account:O365_BUSINESS_PREMIUM")
    {
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $msolUserCheck.licenses.AccountSkuId
        #Set-MsolUserLicense -UserPrincipalName $msolUserCheck.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
        Write-Host "Updated license $($msolUserCheck.licenses.AccountSkuId) for user $($msolUserCheck.DisplayName)"  -ForegroundColor Green 
    }
    elseif ($msolUserCheck.licenses.AccountSkuId -eq "reseller-account:O365_BUSINESS_ESSENTIALS")
    {
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $msolUserCheck.licenses.AccountSkuId
        Set-MsolUserLicense -UserPrincipalName $msolUserCheck.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
        Write-Host "Updated license $($msolUserCheck.licenses.AccountSkuId) for user $($msolUserCheck.DisplayName)"  -ForegroundColor Green 
    }
}

$msoluser | ? {$_.}
$msolusers = Get-MsolUser -all