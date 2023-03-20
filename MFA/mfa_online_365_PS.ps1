#Gets PAC UPN for Connect-EXOPSSession
$PACCredentials = Read-Host -Prompt 'Insert PAC Credential Username'
#Assigns the tenant name to a variable
$Tenant = Read-Host -Prompt 'Insert entire .onmicrosoft.com tenant name'
#Imports the Microsoft Online Module to work with and Sharepoint Module just in case.
Import-Module MsOnline
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
#Connects to the cloud based on the credentials you used before
Connect-EXOPSSession -UserPrincipalName $PACCredentials -DelegatedOrganization $Tenant