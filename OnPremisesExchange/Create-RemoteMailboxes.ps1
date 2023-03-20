#Create Remote Mailboxes for OnPremises Users FROM Matched Office365 Users

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