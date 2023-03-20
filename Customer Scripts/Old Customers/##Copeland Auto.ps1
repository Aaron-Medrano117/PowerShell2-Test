##Copeland Auto

### Create RSE Mailbox Users
foreach ($user in $copelandautoUsers)
{
    $DisplayName = $user.DisplayName
    $NameSplit = $user.DisplayName -split " "
    Write-Host "Checking for $($newupn) .." -NoNewline
    if (get-msoluser -userprincipalname $user.UserPrincipalName -ea silentlycontinue)
    {
        Write-Host "$($DisplayName) user Exists. Skipping"
    }
    else
    {
        Write-Host "Creating MSOLUser for $($DisplayName) " -ForegroundColor Cyan
        try
        {
            New-MsolUser -UserPrincipalName $user.UserPrincipalName -Usagelocation "US" -FirstName $NameSplit[0] -LastName $NameSplit[0] -DisplayName $DisplayName -Password Yef75LwC
        }
        catch
        {
            New-MsolUser -UserPrincipalName $user.UserPrincipalName -Usagelocation "US" -FirstName $NameSplit[0] -DisplayName $DisplayName -Password Yef75LwC
        }
    }
}


# Add 365 Mail Domain address
$mailboxes = get-mailbox -OrganizationalUnit copelandchevrolet.com |sort displayName
foreach ($mbx in $mailboxes)
{
    Write-Host "Adding Alias Address to mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $aliasaddress = $mbx.name + "@countoncopeland.mail.onmicrosoft.com"
    Set-Mailbox $mbx.alias -EmailAddresses @{add=$aliasaddress}
    Write-Host "done." -fore green
}


# Add 365 Mail Domain address
$mailusers = Get-Mailuser
foreach ($mbx in $mailusers)
{
    Write-Host "Adding Alias Address to mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $aliasaddress = $mbx.name + "@countoncopeland.mail.onmicrosoft.com"
    Set-MailUser $mbx.alias -EmailAddresses @{add=$aliasaddress}
    Write-Host "done." -fore green
}

### Create Distribution Lists
foreach ($list in ($CopelandAutoRecipients | ?{$_.RecipientTypeDetails -like "*distribution*"}))
{
    if (!($recipientcheck = get-recipient $list.primarysmtpaddress -ea silentlycontinue ))
    {
        Write-Host "Creating Distribution Group $($list.DisplayName)" -ForegroundColor Cyan -NoNewline

        New-DistributionGroup -name $list.name -DisplayName $list.DisplayName -alias $list.name -primarysmtpaddress $list.PrimarySmtpAddress -RequireSenderAuthenticationEnabled $false
        <#
        $emailAddressarray = $list.Members -split ","
        Start-Sleep -Seconds 3
        Write-Host "Adding" $emailaddressarray.count "aliases" -NoNewline
        foreach ($alias in $emailAddressarray)
        {
            Write-Host ". " -nonewline -foregroundcolor darkcyan
            Set-DistributionGroup $list.PrimarySmtpAddress -emailaddresses @{add=$alias} -ea silentlycontinue -wa silentlycontinue
        }
        #>
        # Set DL Hidden from Address List
        #Set-DistributionGroup $list.PrimarySmtpAddress -HiddenFromAddressListsEnabled ([System.Convert]::ToBoolean($list.HiddenFromAddressListsEnabled)) -ea silentlycontinue
        Write-Host "done" -ForegroundColor Green
    }
    else
    {
        Write-Host "Recipient exists for $($recipientcheck.DisplayName) as $($recipientcheck.recipienttypedetails). Skipping." -ForegroundColor Yellow
    }
}

### Add members simple
foreach ($recipient in ($CopelandAutoRecipients | ?{$_.RecipientTypeDetails -like "*distribution*" -and $_.RecipientType -ne "HostedExchange"}))
{
    Write-Host "Updating Members for $($recipient.DisplayName)"
    #Check if Member exists
    $members = $recipient.members -split ","
    foreach ($member in $members)
    {
        if ($recipientcheck = Get-EXORecipient $member -ea silentlycontinue)
        {
            Write-Host "Member $($member) found." -ForegroundColor Cyan -NoNewline
            Add-DistributionGroupMember  $recipient.primarysmtpaddress -member $member
            Write-Host "member added" -ForegroundColor Green        
        }
        else
        {
            Write-Host "Error adding member $($member). Not Found." -ForegroundColor Red 
        }
    } 
}

## Add domains
foreach ($mailbox in $mailboxes | sort displayName)
{
    $copelandtoyota =  $mailbox.alias + "@copelandtoyota.com"
    $countoncopeland =  $mailbox.alias + "@countoncopeland.com"
    $copelandchevrolet =  $mailbox.alias + "@copelandchevrolet.com"

    if ($mailboxcheck = get-mailbox $mailbox.identity)
    {
        Write-Host "Updating Email Address for $($mailbox.DisplayName) ..." -NoNewline -ForegroundColor Cyan
        Set-Mailbox $mailboxcheck.primarysmtpaddress -EmailAddresses @{add=$copelandtoyota,$countoncopeland,$copelandchevrolet}
        Write-Host "Done" -ForegroundColor Green
    }
}


## Set forwarding to RS Routing Domains

foreach ($user in $RSUsers)
{
    $routingAddresssplit = $user -split "@"
    $newRoutingAddress = $routingAddresssplit[0] + "@routing." + $routingAddresssplit[1]
    Write-Host "Setting forward of $($user) to $($newRoutingAddress)"
    Set-Mailbox $user -forwardingsmtpaddress $newRoutingAddress -DeliverToMailboxAndForward $false
}

#Create Groups simple
foreach ($list in $chevyDLs)
{
    if (!($recipientcheck = get-recipient $list.primarysmtpaddress -ea silentlycontinue ))
    {
        Write-Host "Creating Distribution Group $($list.DisplayName)" -ForegroundColor Cyan -NoNewline
        New-DistributionGroup -name $list.name -DisplayName $list.DisplayName -alias $list.name -primarysmtpaddress $list.PrimarySmtpAddress -RequireSenderAuthenticationEnabled $false
        Write-Host "done" -ForegroundColor Green
    }
    else
    {
        Write-Host "Recipient exists for $($recipientcheck.DisplayName) as $($recipientcheck.recipienttypedetails). Skipping." -ForegroundColor Yellow
    }
}

### Add members simple per domain
$failedToAddMember = @()
foreach ($recipient in $copelandDLs)
{
    Write-Host "Updating Members for $($recipient.DisplayName)"
    #Check if Member exists
    $members = $recipient.members -split ","
    foreach ($member in $members)
    {
        if ($recipientcheck = Get-Recipient $member -ea silentlycontinue)
        {
            Write-Host "Member $($member) found." -ForegroundColor Cyan -NoNewline
            Add-DistributionGroupMember  $recipient.primarysmtpaddress -member $member
            Write-Host "member added" -ForegroundColor Green        
        }
        else
        {
            Write-Host "Error adding member $($member). Not Found." -ForegroundColor Red
            $failedToAddMember += $member
        }
    }
}

#Update Manager for DLS
$rackspace_Dls = Get-DistributionGroup -ResultSize unlimited | ?{$_.managedby -eq "rackspace"}
foreach ($dl in ($rackspace_Dls | sort DisplayName))
{
    Write-Host "Updating DL Owner for $($dl.DisplayName).  " -fore cyan -NoNewline
    $dlmembers = @()
    $dlmembers = get-DistributionGroupMember $dl.name
    Write-Host "Adding "$dlmembers.count" owners to DL .. " -NoNewline
    foreach ($member in $dlmembers)
    {
        $newOwner = $member.name
        Set-DistributionGroup $dl.name -ManagedBy @{add=$newOwner}
    }
    
    #Remove Rackspace Owner
    Set-DistributionGroup $dl.name -ManagedBy @{remove="rackspace"}
    Write-Host "done" -fore green
}

foreach ($contact in $toyotacontacts)
{
    $NameSplit = $contact.ExternalEmailAddress -split "@" 

    New-MailContact -DisplayName $contact.DisplayName -ExternalEmailAddress $contact.ExternalEmailAddress -name $NameSplit[0]
}

foreach ($contact in $toyotacontacts)
{
    $NameSplit = $contact.ExternalEmailAddress -split "@" 
    $copelandchevrolet = $NameSplit[0] + "@copelandchevrolet.com"
    $copelandtoyota = $NameSplit[0] + "@copelandtoyota.com"

    Set-MailContact $contact.ExternalEmailAddress -emailaddresses @{Add=$copelandchevrolet,$copelandtoyota}
}