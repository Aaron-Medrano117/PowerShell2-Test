<#
.SYNOPSIS
This script is designed to find mailbox attributes for mailboxes in Exchange Online.

.DESCRIPTION

Copyright (c) Rackspace US, Inc.

All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

This script was developed to help migrate management of Exchange Objects from one forest to the other but maintain user authentication in the source forest.

MANDATORY REQUIREMENT: 

Office 365 : MSOnline and Exchange Online Powershell Module
On Premise : Exchange Management Shell

Authors: 
Fred Bean - fred.bean@rackspace.com
Chad Matlock - chad.matlock@rackspace.com
=========================================

.PARAMETER OutputCSVFilePath
Provides location of output csv file to store results.
.PARAMETER MsolUserPrincipalName
Provide for a single user to run this against.
.PARAMETER OnPrem
Provide to use OnPrem Exchange Commands.
.PARAMETER hexOU
Used with OnPrem to scope to a single OU.

.EXAMPLE
Pull all required attributes for licensed mailboxes in the tenant.
.\Get-Mailboxattributes.ps1 -OutputCSVFilePath "c:\temp"

.EXAMPLE
Pull all required attributes for a single licensed mailbox in the tenant.
.\Get-Mailboxattributes.ps1 -OutputCSVFilePath "c:\temp" -MsolUserPrincipalName "john.doe@contoso.com"

.EXAMPLE
Pull all On Premise Exchange User Mailboxes.
.\Get-Mailboxattributes.ps1 -OutputCSVFilePath "c:\temp" -OnPrem"

.EXAMPLE
Pull all UserMailboxes to a single OU. This is has to be used with the OnPrem Switch. Using this will be needed in Hosted Exchange.
.\Get-Mailboxattributes.ps1 -OutputCSVFilePath "c:\temp" -OnPrem -hexOU contoso.com
#>

param(
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Output CSV File path")][string]$OutputCSVFilePath,
    [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Use this attribute to pull a single user")][string]$MsolUserPrincipalName,
    [parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Used to switch to use On Premise Exchange")][switch]$OnPrem,
    [parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Used to provide OU within Hosted Exchange")][string]$hexOU
)

$allattributes = @()
$date = Get-Date -Format yyyyMMdd


if($MsolUserPrincipalName)
{
    $users = Get-MsolUser -UserPrincipalName $MsolUserPrincipalName
    $csvFileName = "$($MsolUserPrincipalName)_$($date).csv"
    $msolDomain = Get-MsolDomain | Where-Object{$_.Name -like "*onmicrosoft.com" -and $_.Name -notlike "*mail.onmicrosoft.com"}
}
elseif($OnPrem)
{
    if($hexOU)
    {
        $users = Get-User -ResultSize unlimited -OrganizationalUnit $hexOU | Where-Object{$_.RecipientType -eq "UserMailbox"}
    }
    else
    {
        $users = Get-User -ResultSize unlimited | Where-Object{$_.RecipientType -eq "UserMailbox"}
    }
    $csvFileName = "AllOnPremMailboxes_$($date).csv"
}
else
{
    $users = Get-MsolUser -All | Where-Object{$_.islicensed -eq "True"}
    $csvFileName = "AllLicensedMailboxes_$($date).csv"
    $msolDomain = Get-MsolDomain | Where-Object{$_.Name -like "*onmicrosoft.com" -and $_.Name -notlike "*mail.onmicrosoft.com"}
}

foreach ($user in $users)
{
    Write-Host "$($user.identity)"
    $mailbox = get-mailbox $user.identity
    $addresses = $mailbox.emailaddresses | Where-Object{$_ -notlike "SPO:*" -and $_ -notlike "sip:*" -and $_ -notlike "*onmicrosoft.com" -and $_ -notlike "*x400*" -and $_ -notlike "*@transconex.com" -and $_ -notlike "*@lindenmotorfreight.com" -and $_ -notlike "*@trxdepr.com"}

    if($onprem)
    {
        $onMicrosoft = "N/A"
    }
    else
    {
        $onmicrosoft = $mailbox.emailaddresses | Where-Object{$_ -like "*$($msolDomain.Name)"}
    }

    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName
    $currentuser | add-member -type noteproperty -name "FirstName" -Value $user.FirstName
    $currentuser | add-member -type noteproperty -name "LastName" -Value $user.LastName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $user.userprincipalname
    $currentuser | add-member -type noteproperty -name "Identity" -Value $user.Identity
    $currentuser | add-member -type noteproperty -name "Department" -Value $user.Department
    $currentuser | add-member -type noteproperty -name "Office" -Value $user.Office
    $currentuser | add-member -type noteproperty -name "PhoneNumber" -Value $user.PhoneNumber
    $currentuser | add-member -type noteproperty -name "MobilePhone" -Value $user.MobilePhone
    $currentuser | add-member -type noteproperty -name "Fax" -Value $user.Fax
    $currentuser | add-member -type noteproperty -name "PostalCode" -Value $user.PostalCode
    $currentuser | add-member -type noteproperty -name "State" -Value $user.State
    $currentuser | add-member -type noteproperty -name "StreetAddress" -Value $user.StreetAddress
    $currentuser | add-member -type noteproperty -name "City" -Value $user.City
    $currentuser | add-member -type noteproperty -name "Country" -Value $user.Country
    $currentuser | add-member -type noteproperty -name "Title" -Value $user.Title

    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $mailbox.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($addresses -join ",")
    $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $mailbox.legacyexchangedn)
    $currentuser | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value ($mailbox.AcceptMessagesOnlyFrom -join ",")
    $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($mailbox.GrantSendOnBehalfTo -join ",")
    $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $mailbox.HiddenFromAddressListsEnabled
    $currentuser | add-member -type noteproperty -name "RejectMessagesFrom" -Value ($mailbox.RejectMessagesFrom -join ",")
    $currentuser | add-member -type noteproperty -name "OnMicrosoft" -Value $onmicrosoft
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $mailbox.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $mailbox.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $mailbox.ForwardingSmtpAddress
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mailbox.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "ExchangeGuid" -Value $mailbox.ExchangeGuid

    $allattributes += $currentuser
}

$allattributes | Export-csv "$($OutputCSVFilePath)\$($csvFileName)" –notypeinformation –encoding utf8
