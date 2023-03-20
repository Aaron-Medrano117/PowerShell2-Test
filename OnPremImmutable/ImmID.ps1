#If you want to dump your existing AD to text file for reference uncomment the next line 
ldifde -f C:\export.txt -r "(Userprincipalname=*)" -l "objectGuid, userPrincipalName"

do{ 
# Query the local AD and get all the users output to grid for selection 
$ADGuidUser = Get-ADUser -Filter * | Select Name,ObjectGUID | Sort-Object Name | Out-GridView -Title "Select Local AD User To Get Immutable ID for" -PassThru 
#Convert the GUID to the Immutable ID format 
$UserimmutableID = [System.Convert]::ToBase64String($ADGuidUser.ObjectGUID.tobytearray())

$OnlineUser = Get-MsolUser | Select UserPrincipalName,DisplayName,ProxyAddresses,ImmutableID | Sort-Object DisplayName | Out-GridView -Title "Select The Office 365 Online User To HardLink The AD User To" -PassThru

OnPrem AD ImmutableID 
Set-MSOLuser -UserPrincipalName $OnlineUser.UserPrincipalName -ImmutableID $UserimmutableID

#List Users and ImmutableId 
Get-MsolUser | Select DisplayName,ImmutableID | Sort-Object DisplayName | Out-GridView -Title "Office 365 User List With Immutableid Showing"