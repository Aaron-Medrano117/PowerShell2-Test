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
                $tmp | Add-Member -MemberType NoteProperty -Name ExternalEmailAddress -Value ""
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
                $tmp.ExternalEmailAddress = $mbx.name + "@bolandservices.mail.onmicrosoft.com"
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

$mailboxProperties | Export-Csv $HOME\MailboxProperties.csv -NoTypeInformation -Encoding UTF8