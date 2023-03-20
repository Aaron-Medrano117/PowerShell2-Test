## For US ONLY
Write-Host "   **For US PACS Only**" -ForegroundColor yellow

#Gets PAC UPN for Connect-EXOPSSession
$PAC_Username = Read-Host -Prompt 'Insert PAC Credential Username. IE SSO portion'
$PAC_Credentials = @()
$PAC_Credentials = $PAC_Username + "@managed365.onmicrosoft.com"

#Delegated Tenant (Customer Tenant)
$DelegatedOrg = Read-Host "Enter in customer's domain or tenant name (including the .onmicrosoft.com)"

#Connect to Exchange Online
Connect-EXOPSSession -UserPrincipalName $PAC_Credentials -DelegatedOrganization $DelegatedOrg -AzureADAuthorizationEndpointUri https://login.microsoftonline.us/common