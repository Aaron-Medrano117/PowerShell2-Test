#REGION Get MAILBOX swing attributes from O365

##########################################
# Get MAILBOX swing attributes from O365 #
##########################################

$allMailboxes = Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySmtpAddress.ToString() -notlike "DiscoverySearchMailbox*"} | sort PrimarySmtpAddress

$mailboxProperties = @()
foreach ($mbx in $allMailboxes)
{
    Write-Host "$($mbx.PrimarySmtpAddress.ToString()) ... " -ForegroundColor Cyan -NoNewline
    
    $tmp = New-Object -TypeName PSObject
    $tmp | Add-Member -MemberType NoteProperty -Name OnPremUPN -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name O365UPN -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name ImmutableID -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name Name -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name DisplayName -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name ExchangeGuid -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name RecipientTypeDetails -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name EmailAddresses -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name LegacyExchangeDN -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name AcceptMessagesOnlyFrom -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name AcceptMessagesOnlyFromDLMembers -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name AcceptMessagesOnlyFromSendersOrMembers -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name Alias -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name IsShared -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name GrantSendOnBehalfTo -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name HiddenFromAddressListsEnabled -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name RejectMessagesFrom -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name RejectMessagesFromDLMembers -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name RejectMessagesFromSendersOrMembers -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name RequireSenderAuthenticationEnabled -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name FirstName -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name LastName -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name City -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name Company -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name Country -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name Department -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name Fax -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name MobilePhone -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name Office -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name PhoneNumber -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name POBox -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name PostalCode -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name State -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name StreetAddress -Value ""
    $tmp | Add-Member -MemberType NoteProperty -Name Title -Value ""
    
    $tmp.Name = $mbx.Name
    $tmp.O365UPN = $mbx.UserPrincipalName
    $tmp.DisplayName = $mbx.DisplayName
    $tmp.ExchangeGuid = $mbx.ExchangeGuid
    $tmp.RecipientTypeDetails = $mbx.RecipientTypeDetails
    $tmp.PrimarySmtpAddress = $mbx.PrimarySmtpAddress.ToString()
    $tmp.EmailAddresses = ($mbx.EmailAddresses | Where {$_ -notlike "*@routing.*" -and $_ -notlike "SIP:*" -and $_ -notlike "SPO:*"}) -join ","
    $tmp.EmailAddresses += ",X500:" + $mbx.LegacyExchangeDN
    $tmp.LegacyExchangeDN = $mbx.LegacyExchangeDN
    $tmp.AcceptMessagesOnlyFrom = $mbx.AcceptMessagesOnlyFrom -join ","
    $tmp.AcceptMessagesOnlyFromDLMembers = $mbx.AcceptMessagesOnlyFromDLMembers -join ","
    $tmp.AcceptMessagesOnlyFromSendersOrMembers = $mbx.AcceptMessagesOnlyFromSendersOrMembers -join ","
    $tmp.Alias = $mbx.Alias
    $tmp.IsShared = $mbx.IsShared
    $tmp.GrantSendOnBehalfTo = $mbx.GrantSendOnBehalfTo -join ","
    $tmp.HiddenFromAddressListsEnabled = $mbx.HiddenFromAddressListsEnabled
    $tmp.RejectMessagesFrom = $mbx.RejectMessagesFrom -join ","
    $tmp.RejectMessagesFromDLMembers = $mbx.RejectMessagesFromDLMembers -join ","
    $tmp.RejectMessagesFromSendersOrMembers = $mbx.RejectMessagesFromSendersOrMembers -join ","
    $tmp.RequireSenderAuthenticationEnabled = $mbx.RequireSenderAuthenticationEnabled

    if (-not ($msolUser = Get-MsolUser -UserPrincipalName $mbx.UserPrincipalName -EA SilentlyContinue))
    {
        continue
    }
    
    $tmp.FirstName = $msolUser.FirstName
    $tmp.LastName = $msolUser.LastName
    $tmp.City = $msolUser.City
    $tmp.Country = $msolUser.Country
    $tmp.Department = $msolUser.Department
    $tmp.Fax = $msolUser.Fax
    $tmp.MobilePhone = $msolUser.MobilePhone
    $tmp.Office = $msolUser.Office
    $tmp.PhoneNumber = $msolUser.PhoneNumber
    $tmp.PostalCode = $msolUser.PostalCode
    $tmp.State = $msolUser.State
    $tmp.StreetAddress = $msolUser.StreetAddress
    $tmp.Title = $msolUser.Title
    
    $mailboxProperties += $tmp
    
    Write-Host "done" -ForegroundColor Green
}

$mailboxProperties | Export-Csv $HOME\Desktop\MailboxProperties.csv -NoTypeInformation -Encoding UTF8


##################################################################

#ENDREGION

#REGION Populate the OnPremUPN and ImmutableID columns

##################################################
# Populate the OnPremUPN and ImmutableID columns #
##################################################

$mailboxProperties = Import-Csv $HOME\Desktop\MailboxProperties.csv | sort O365UPN # Path to CSV

#ProgressBar1
$progressref = ($mailboxProperties).count
$progresscounter = 0

foreach ($mbx in $mailboxProperties)
{
    $upn = $mbx.O365UPN
    $DisplayName = $mbx.DisplayName
    $EmailAddress = $mbx.PrimarySmtpAddress

    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for ADUser $($DisplayName)"

    #if ($adUser = Get-ADUser -Filter {Mail -eq $upn} -Properties ObjectGUID)
    if ($adUser = Get-ADUser -Filter {UserPrincipalName -eq $upn} -Properties ObjectGUID, mail)
    {
        Write-Host $upn -ForegroundColor Green
        $mbx | Add-Member -MemberType NoteProperty -Name "ExistsOnPrem" -Value $true -force
        $mbx | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $adUser.DistinguishedName -force
        $mbx | Add-Member -MemberType NoteProperty -Name "OnPremUPN" -Value $adUser.UserPrincipalName -force
        $objGUID = $adUser | select -ExpandProperty ObjectGUID | select -ExpandProperty Guid
        $ImmutableID = [System.Convert]::ToBase64String(([GUID]($objGUID)).ToByteArray())
        $mbx | Add-Member -MemberType NoteProperty -Name "Mail" -Value $adUser.mail -force
        $mbx | Add-Member -MemberType NoteProperty -Name "ImmutableID" -Value $ImmutableID -force
    }
    elseif ($adUser = Get-ADUser -filter {Mail -eq $emailAddress} -Properties ObjectGUID, mail)
    {
        Write-Host $upn -ForegroundColor Cyan
        $mbx | Add-Member -MemberType NoteProperty -Name "ExistsOnPrem" -Value $true -force
        $mbx | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $adUser.DistinguishedName -force
        $mbx | Add-Member -MemberType NoteProperty -Name "OnPremUPN" -Value $adUser.UserPrincipalName -force
        $objGUID = $adUser | select -ExpandProperty ObjectGUID | select -ExpandProperty Guid
        $ImmutableID = [System.Convert]::ToBase64String(([GUID]($objGUID)).ToByteArray())
        $mbx | Add-Member -MemberType NoteProperty -Name "Mail" -Value $adUser.mail -force
        $mbx | Add-Member -MemberType NoteProperty -Name "ImmutableID" -Value $ImmutableID -force
    }
    elseif ($adUser = Get-ADUser -filter {Name -eq $DisplayName} -Properties ObjectGUID, mail)
    {
        Write-Host $DisplayName -ForegroundColor Yellow
        $mbx | Add-Member -MemberType NoteProperty -Name "ExistsOnPrem" -Value $true -force
        $mbx | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $adUser.DistinguishedName -force
        $mbx | Add-Member -MemberType NoteProperty -Name "OnPremUPN" -Value $adUser.UserPrincipalName -force
        $objGUID = $adUser | select -ExpandProperty ObjectGUID | select -ExpandProperty Guid
        $ImmutableID = [System.Convert]::ToBase64String(([GUID]($objGUID)).ToByteArray())
        $mbx | Add-Member -MemberType NoteProperty -Name "Mail" -Value $adUser.mail -force
        $mbx | Add-Member -MemberType NoteProperty -Name "ImmutableID" -Value $ImmutableID -force
    }
    else
    {
        Write-Host $upn -ForegroundColor Red
        $mbx | Add-Member -MemberType NoteProperty -Name "ExistsOnPrem" -Value $false -force
        $mbx | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $Null -force
        $mbx | Add-Member -MemberType NoteProperty -Name "OnPremUPN" -Value $Null -force
        $mbx | Add-Member -MemberType NoteProperty -Name "Mail" -Value $adUser.mail -force
        $mbx | Add-Member -MemberType NoteProperty -Name "ImmutableID" -Value $Null -force
    }          
}

$mailboxProperties | Export-Csv $HOME\Desktop\MailboxProperties.csv -NoTypeInformation -Encoding UTF8

##################################################################
#ENDREGION

##################################################################
#STARTREGION

# CREATE REMOTE MAILBOXES FROM MATCHED ONPREM USERS

$office365_mailboxes = Import-Csv C:\Users\raxadmin\Desktop\MailboxProperties.csv
$onpremusers = $office365_mailboxes | ?{$_.existsOnPrem -eq $true}
$onpremusers = $onpremusers | ? {$_.RecipientTypeDetails -eq "UserMailbox"}

#INPUT Variables
$domain = "@theroomplace.com"
$microsoftdomain = "@trpacqinc.mail.onmicrosoft.com"

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
    $RemoteRoutingAddress = $user.PrimarySmtpAddress.replace($domain,$microsoftdomain)
    
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
##################################################################
#ENDREGION



##################################################################
#STARTREGION
## STAMP IMMUTABLE ID - HARD MATCH


$mailboxProperties = Import-Csv $HOME\Desktop\MailboxProperties.csv | sort O365UPN # Path to CSV

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

    Set-MsolUser -UserPrincipalName $mailbox.O365UPN -ImmutableID $mailbox.ImmutableID
    Write-Host "Updated ImmutableID to $($mailbox.ImmutableID) for $($mailbox.O365UPN)" -foregroundcolor green

}