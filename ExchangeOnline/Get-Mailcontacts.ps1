function Get-MailContacts
{

    <#
  
        .SYNOPSIS
            Gather the needed mailcontact and msolcontact attributes in and Office 365 tenant. 
  
        .DESCRIPTION
            Uses MSOL get-msolcontact and Exchange Online get-mailcontact cmdlets to gather information about contacts in an Office 365 tenant.
  
        .OUTPUTS 
            Populates an array with powershell objects containing interesting contact attributes and exports the array to csv.
  
        .EXAMPLE
            Get-MailContacts
  
    #> 

#get all mail contacts in the tenant
$contacts = Get-MailContact -resultsize unlimited
#create an empty array to store our custom contact objects
$allattributes = @()
#loop through all of the mail contacts
foreach ($contact in $contacts)
    {
        write-host "Working on contact" $contact.externalemailaddress -fore cyan
        #remove any addresses we don't want and store the remaining ones in the $addresses variable
        $addresses = $contact.emailaddresses | ?{$_ -notlike "SPO:*" -and $_ -notlike "sip:*" -and $_ -notlike "*onmicrosoft.com"}
        #get the related msolobject for the contact
        $msolcontact = Get-MSOLContact -searchstring $contact.DisplayName
        #create a custom object to store the attributes we're interested in
        $currentcontact = "" | select DisplayName, FirstName, LastName, EmailAddresses, HiddenFromAddressListsEnabled, LegacyExchangeDN, ExternalEmailAddress
        #populate that custom object
        $currentcontact.DisplayName = $contact.DisplayName
        $currentcontact.FirstName = $msolcontact.FirstName
        $currentcontact.LastName = $msolcontact.LastName
        $currentcontact.EmailAddresses = ($addresses -join ",")
        $currentcontact.HiddenFromAddressListsEnabled = $contact.HiddenFromAddressListsEnabled
        $currentcontact.LegacyExchangeDN = $contact.LegacyExchangeDN
        $currentcontact.ExternalEmailAddress = $contact.ExternalEmailAddress
        #add the custom object to our array
        $allattributes += $currentcontact
    }
#export the array that contains all of our contacts
$allattributes | export-csv <file path and name> -notypeinformation -encoding utf8
}
