# This script is used to import and create new users into ActiveDirectory
# Requires a csv import with user information, the ou where it is located, and creates users without a password and in a disabled state
# Cervey Customer specific


$ImportCSV = Read-Host "What is the csv file to import and create new AD users?"
$Cerveyimport = Import-Csv $ImportCSV

foreach ($cerveyuser in $Cerveyimport) 
{
   if (Get-ADUser -Identity ($Cerveyuser.displayname[0..12] -join"").trim())
     {
     Write-Host "User already exists. Skipping" $cerveyuser.displayname -ForegroundColor Green
     }
       else
       {
     Write-Host "Creating AD User for "$cerveyuser.Displayname"" -ForegroundColor White
     set-ADUser -UserPrincipalName $cerveyuser.UserPrincipalName -Name ($Cerveyuser.displayname[0..12] -join"").trim() -Path $cerveyuser.OrganizationalUnit -City $cerveyuser.City -Country $cerveyuser.Country -Department $cerveyuser.department -DisplayName $cerveyuser.Displayname -Fax $cerveyuser.fax -GivenName $cerveyuser.FirstName -Surname $cerveyuser.LastName	-MobilePhone $cerveyuser.MobilePhone -Office $cerveyuser.Office -OfficePhone $cerveyuser.PhoneNumber -PostalCode $cerveyuser.postalcode	-State $cerveyuser.state -StreetAddress $cerveyuser.stateaddress -Title $cerveyuser.title -EmailAddress $cerveyuser.PrimarySMTPAddress -OtherAttributes @{'LegacyExchangeDN'=$cerveyuser.LegacyExchangeDN; 'msExchHideFromAddressLists'=$cerveyuser.HiddenFromAddressListsEnabled; 'msExchRecipientTypeDetails'=$cerveyuser.RecipientTypeDetails; 'proxyAddresses'=$cerveyuser.EmailAddresses} 
     
     #verify user was created
     $userverifiication =  Get-ADUser -Identity ($Cerveyuser.displayname[0..12] -join"").trim() -Server 1034073-AD1
     Start-Sleep -Seconds 5
         if (!$userverifiication)
         {
         Write-host '"$cerveyuser.Displayname" was not created. Recreate user.' -ForegroundColor Yellow
         }
            else
            {
            Write-Host '"$cerveyuser.Displayname" created successfully' -ForegroundColor Green
            }
    
  }
}

#rename AD Object due to character limit during AD object creation

foreach ($cerveyuser in $Cerveyimport) 
{
$upn = $cerveyuser.userprincipalname
Get-ADUser -Filter {userprincipalname -eq $upn} | Rename-ADObject -NewName $cerveyuser.displayname
}

$allusers = @()
foreach ($cerveyuser in $Cerveyimport) 
{
$upn = $cerveyuser.userprincipalname
if (Get-ADUser -Filter {userprincipalname -eq $upn})
 {}
 else
 {
 $allusers += $upn
 $cerveyuser
 }
}

foreach ($cerveyuser in $Cerveyimport)
{
set-ADUser -Server 1034073-AD1 -UserPrincipalName $cerveyuser.UserPrincipalName -City $cerveyuser.City -Country $cerveyuser.Country -Department $cerveyuser.department -DisplayName $cerveyuser.Displayname -Fax $cerveyuser.fax -GivenName $cerveyuser.FirstName -Surname $cerveyuser.LastName	-MobilePhone $cerveyuser.MobilePhone -Office $cerveyuser.Office -OfficePhone $cerveyuser.PhoneNumber -PostalCode $cerveyuser.postalcode	-State $cerveyuser.state -StreetAddress $cerveyuser.stateaddress -Title $cerveyuser.title -EmailAddress $cerveyuser.PrimarySMTPAddress -OtherAttributes @{'LegacyExchangeDN'=$cerveyuser.LegacyExchangeDN; 'msExchHideFromAddressLists'=$cerveyuser.HiddenFromAddressListsEnabled; 'msExchRecipientTypeDetails'=$cerveyuser.RecipientTypeDetails; 'proxyAddresses'=$cerveyuser.EmailAddresses} 
}


#Enable Remote Mailboxes without a Remote Address set
foreach ($cerveyuser in $Cerveyimport) 
{
$upn = $cerveyuser.userprincipalname
$remoteaddress = $cerveyuser.onmicrosoft
#if ($remoteaddress) {

(Get-User $upn | enable-remotemailbox -RemoteRoutingAddress $remoteaddress)
  #  }
 }


foreach ($cerveyuser in $Cerveyimport)
{
$upn = $cerveyuser.userprincipalname

$LegacyExchangeDN = @()
$LegacyExchangeDN = $cerveyuser.LegacyExchangeDN

if (get-remotemailbox $upn) {

set-RemoteMailbox -identity $cerveyuser.UserPrincipalName  -EmailAddresses @{add=$LegacyExchangeDN} -HiddenFromAddressListsEnabled $false
    }
}