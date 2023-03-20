### This Script is designed for Cutover Migrations utilizing the Office 365 Cutover Migration Option.
### It gathers all Exchange Objects of a domain in Cloud Office Exchange including Distribution Lists, Exchange Mail Contacts, Mailboxes, and Resources
###  Sets all definited Exchange Objects to Visible in GAL and Adds Migration User as Owner for Distribution Lists

## Version 1.2 by Aaron Medrano 11/19/2019
# Add Throttling Policy to Unrestricted
# Skip Folder Users and Folder Admins Groups from DL Visible in GAL update

#Intro
cls
Write-Host " This Script is designed for Cutover Migrations utilizing the Office 365 Cutover Migration Option." -ForegroundColor White
Write-Host " Gathers all Exchange Objects of a domain in Cloud Office Exchange including Distribution Lists, Exchange Mail Contacts, Mailboxes, and Resources" -ForegroundColor White
Write-Host " Will set all definited Exchange Objects to Visible in GAL and Adds Migration User as Owner for Distribution Lists
"  -ForegroundColor White

##Get Variables
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
$migrationuser = Read-Host "What is the email address of the migration user you wish to utilize? 
 If the address does not exist, the mailbox will be created" 

#Validate provided email address
 Write-Host " Checking email address provided ..." -ForegroundColor Yellow
$error.Clear()
$migrationemailaddress = $migrationuser -split '@'
$migrationdomain = $migrationemailaddress[1]
$migration_alias = $migrationemailaddress[0] + "." + $migrationdomain
 if ($Error.Count -ge 1) {Write-Host "Invalid Email Address provided. Ending script" -ForegroundColor Red
Exit }

#check If domain is valid

$domain_verification = get-organizationalunit $migrationdomain -ErrorAction SilentlyContinue
if ($domain_verification) {Write-host "  Confirmed domain exists" -ForegroundColor Green}
    else {Write-Host ""
          Write-Host "  Domain $migrationdomain not found. Recheck domain spelling" -ForegroundColor Red
            Exit}


# Create Migration User
Write-Host "
  ** Validating Migration Admin Account ..." -ForegroundColor White

## Confirm Email Address already Exists
$migrationsvcaccount = get-mailbox $migrationuser -ErrorAction SilentlyContinue
if ($migrationsvcaccount) { Write-host "Confirmed user exists. Continuing process ..." -ForegroundColor Green}
    else {Write-Host "User $migrationuser not found. Creating Migration Admin ... You will be Prompted to enter a password
        " -ForegroundColor Cyan
                 try {New-Mailbox -userprincipalname "$migrationuser" -organizationalunit $migrationdomain -Name "Migration_Admin" -Alias $migration_alias -ErrorAction Stop -Confirm}
                 catch {Write-Host ""
                 Write-Host "The Display Name 'Migration Admin' is already in use. Please check if a Migration Admin has already been created" -ErrorAction stop -ForegroundColor Red
                 Exit}
            Write-Host "Be sure to note down the password, you will need it later" -ForegroundColor Yellow
            }
                   

#Add User to Address Book Policy and set Custom Attribute 10 to domain
Write-Host "
  ** Add "$migrationuser" to Address Book Policy of $migrationdomain" -ForegroundColor White
  $addressbookPolicyCheck = $migrationsvcaccount | ? {$_.AddressListMembership -match $migrationdomain}
  if ($addressbookPolicyCheck) {
  Write-Host "Address Book is already updated. Skipping ...
  " -ForegroundColor Cyan}
  else {Set-mailbox $migrationuser -addressbookpolicy "$migrationdomain abp" -ErrorAction SilentlyContinue -CustomAttribute10 $migrationdomain
        Write-Host "Address Book Updated
            " -ForegroundColor Green}


#Set Throttling Policy to Unrestricted
Write-Host "
  ** Set Throttling Policy for '$migrationuser' to Unrestricted" -ForegroundColor White
  $throttlingPolicyCheck = $migrationsvcaccount | ? {$_.ThrottlingPolicy -match "unrestricted"}
  if ($throttlingPolicyCheck) {
  Write-Host "Throttling Policy is already updated. Skipping ...
  " -ForegroundColor Cyan}
  else {Set-mailbox $migrationuser -throttlingpolicy unrestricted -ErrorAction SilentlyContinue
        Write-Host "Throttling Policy Updated
            " -ForegroundColor Green}

#gather all objects on domain
Write-Host "Gathering Distribution Groups for $migrationdomain ..."
$distributiongroups = get-distributiongroup -organizationalunit $migrationdomain -resultsize unlimited | ? {$_.name -ne "folderusers" -or $_.name -ne "folderadmins"}
$distributiongroupsvisible = $distributiongroups | ? {$_.HiddenFromAddressListsEnabled -eq $false}
$distributiongroupshidden = $distributiongroups | ? {$_.HiddenFromAddressListsEnabled -eq $true}

Write-Host "Gathering Mail Contacts for $migrationdomain"
$mailcontacts = get-mailcontact -organizationalunit $migrationdomain -resultsize unlimited
$mailcontactsvisible = $mailcontacts | ? {$_.HiddenFromAddressListsEnabled -eq $false}
$mailcontactshidden = $mailcontacts | ? {$_.HiddenFromAddressListsEnabled -eq $true}

Write-Host "Gathering Mailboxes and Resources for $migrationdomain ..."
$domainmailboxesvisible = $mailboxlist | ? {$_.HiddenFromAddressListsEnabled -eq $false}
$domainmailboxeshidden = $mailboxlist | ? {$_.HiddenFromAddressListsEnabled -eq $true}
Write-Host ""


## Make All Mailboxes Visible In GAL
$mailboxlist = get-mailbox -organizationalunit $migrationdomain -resultsize unlimited
write-Host ""($mailboxlist).count" Mailboxes and Resources Found for domain $migrationdomain" -ForegroundColor White
Write-Host ""($domainmailboxeshidden).count" Hidden Mailboxes Found for domain $migrationdomain" -ForegroundColor Yellow
if ($domainmailboxeshidden.count -eq 0) {
        Write-Host "All Mailboxes "$domainmailboxeshidden.count" Mailboxes already Visible in GAL. Skipping ...
        " -ForegroundColor cyan }
           else {Write-Host "**Setting "$mailcontactshidden.count" Mailboxes Visible in GAL ..." -ForegroundColor green
Write-Host ""
## Set All Mailboxes Visible in GAL State
foreach ($mailbox in $domainmailboxeshidden) {
            if ($mailbox.HiddenFromAddressListsEnabled -eq 1) {
                 $setmailboxvisible = Set-mailbox $mailbox -HiddenFromAddressListsEnabled $false -ErrorAction SilentlyContinue
                 $setmailboxvisible
                  if ($setmailboxvisible) {}
            else {Write-Host "'$mailbox' not updated. Try running again manually" -ForegroundColor Yellow}
                }
            }
}

## Make All MailContacts Visible In GAL
Write-Host ""($mailcontacts).count" Mail Contacts Found for domain $migrationdomain" -ForegroundColor white
Write-Host ""($mailcontactshidden).count" Hidden Mail Contacts Found for domain $migrationdomain" -ForegroundColor Yellow
if ($mailcontactshidden.count -eq 0) {
        Write-Host "All Mail Contacts Already Visible in GAL. Skipping..." -ForegroundColor Cyan}
        else {Write-Host "**Setting "($mailcontactshidden).count" Distribution Groups Visible in GAL ..." -ForegroundColor green
Write-Host ""
## Set All Mail Contacts Visible in GAL State
foreach ($contact in $mailcontactshidden) {
            if ($contact.HiddenFromAddressListsEnabled -eq 1) {
                 Set-mailcontact $contact -HiddenFromAddressListsEnabled $false -ErrorAction SilentlyContinue
                 Write-Host "  '$contact' Visible in GAL" -ForegroundColor Green}
           }
   }
             

## Make All Distribution Groups Visible In GAL
Write-Host ""
Write-Host ""($distributiongroups).count" Distribution Groups Found for domain $migrationdomain" -ForegroundColor white
Write-Host ""($distributiongroupshidden).count" Hidden Distribution Groups Found for domain $migrationdomain" -ForegroundColor Yellow
if ($distributiongroupshidden.count -eq 0) {
        Write-Host "All Distribution Groups already Visible in GAL. skipping" -ForegroundColor Cyan}
        else {Write-Host "**Setting "($distributiongroupshidden).count" Distribution Groups Visible in GAL ..." -ForegroundColor green
Write-Host ""

## Set All Distribution Groups Visible in GAL State
foreach ($group in $distributiongroups) {
            if ($group.HiddenFromAddressListsEnabled -eq 1) {
            Set-DistributionGroup $group -HiddenFromAddressListsEnabled $false -ErrorAction SilentlyContinue -ErrorVariable $FailedVisibleGAL_DL  }
                 if ($FailedVisibleGAL_DL) {Write-Host "Group '$group' not updated. Try running again manually" -ForegroundColor Yellow
                 }
                 else {Write-Host "Group $group updated successfully" -ForegroundColor Cyan} 
              }

}
                 
        
#Grant Full Access Permissions

Write-host " 

  Setting Full Access Permissions for for all users on $migrationdomain domain" -ForegroundColor White
foreach ($mbx in $mailboxlist) {
    $addMailboxPerms = add-mailboxpermission -Identity $mbx -user $migrationuser -accessrights FullAccess -automapping $false -ErrorVariable $mailboxpermission_error -WarningAction SilentlyContinue
    for ($a = 1; $a -le 100; $a++ )
                        {
                          Write-Progress -Activity "Granting '$migrationuser' Full Access to mailbox '$mbx'" -Status "$a% Complete:" -PercentComplete $a;
                        }
    if ($mailboxpermission_error.count -eq 1) {Write-Host "There was a problem adding permissions to '$mbx'" -ForegroundColor Red }
        }
     If ($mailboxpermission_error.count -eq 0) {Write-Host "Successfully Added Full Access Permissions to all Mailboxes" -ForegroundColor Green}

##Setting Migration User as Owner to Groups
Write-Host ""
Write-Host "  Setting '$migrationuser' as Owner to Groups" -ForegroundColor White
#Check If User is Owner of List
$Migration_Identity = (get-mailbox $migrationuser).Identity
$Groups_MigrationOwner = Get-DistributionGroup -OrganizationalUnit $migrationdomain -ResultSize unlimited | ? {$_.managedby -match $Migration_Identity}
        
        foreach ($group in $Groups_MigrationOwner) {
           Set-DistributionGroup -Identity $group -ManagedBy @{add=$Migration_Identity} -BypassSecurityGroupManagerCheck -OutVariable $DLOwner_error -ErrorAction Stop
          
           $DLOwner_error_results += $DLOwner_error

            for ($k = 1; $k -le 100; $k++ ) {
                            Write-Progress -Activity "Add '$migrationuser' as Owner to Distribution Group '$group'" -Status "$k% Complete:" -PercentComplete $k;
                        }
                
            if ($DLOwner_error) {write-host "Unable to update '$group'"}
     }
    if ($DLOwner_error_results.count -eq 0) {Write-host "All Distribution Groups Updated Successfully" -ForegroundColor Green}
    else {Write-Host "Some Distributions Failed to update. Manually check which groups failed to update" -ForegroundColor Red}
         
        
Write-Host "

****  Prerequistes Completed For Cutover Migration to Office 365 for domain $migrationdomain   *****" -ForegroundColor Cyan
