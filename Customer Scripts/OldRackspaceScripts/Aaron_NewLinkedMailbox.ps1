<#Get Variables#>
$leadtechuser = Read-host "Which Lead Tech Needs to be Created? Enter firstname.lastname. Example: 'aaron.medrano'"

# **************************************************
#     CAUTION: THIS MUST BE SET CORRECTLY!!!!!!!
          $global:Environment = "MEX09"
# **************************************************

#Create Linked MasterAccount
#Auth To MGMT AD
$ADMGMT = get-credential

Write-Host "Searching for Mail User $leadtechuser ..."

#Find Mail User
$ltUPN = (Get-AdUser $leadtechuser).userprincipalname
$ltmailuser = Get-Mailuser $ltUPN

#Get DisplayName
$ltDisplayName = (Get-MailUser $ltmailuser).displayname
#Create New Linked Mailbox DisplayName
$LTLinkedDisplayName = $($ltDisplayName + " - Linked")

#Get Username
$ltusername = (Get-AdUser $leadtechuser).name

Write-Host "Updating Old User $leadtechuser ..." -ForegroundColor magenta
<#Remove AdUser#>
Remove-AdUser $leadtechuser -confirm:$false

Write-Host "Creating New Linked Mailbox ..." -ForegroundColor Green
#Wait for replication:
Start-Sleep -s 15

<#Create New Linked Mailbox#>
$LTLinkedMasterAccount = $($ltusername +"@mgmt.mlsrvr.com")
$Lt_OU = $($global:Environment + ".mlsrvr.com/Engineering/Support/L1.5")

New-Mailbox $LtLinkedDisplayName -LinkedMasterAccount $LTLinkedMasterAccount -LinkedCredential $ADMGMT -LinkedDomainController mgmtad01.mgmt.mlsrvr.com -DisplayName $LtLinkedDisplayName -OrganizationalUnit $Lt_OU -alias $ltusername

$NewLinkedMailbox = Get-Mailbox $ltUPN

Write-Host "New Linked Master Account Created for '$NewLinkedMailbox" -ForegroundColor Cyan