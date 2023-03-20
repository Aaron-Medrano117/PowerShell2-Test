#Magnetrol Scripts
#Get ALL MAILBOX DETAILS and ONEDRIVE DETAILS
function Get-ALLMAILBOXDETAILS {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$False)] [Switch] $OneDriveCheck,
        [Parameter(Mandatory=$False)] [string] $OneDriveURL
        )
    $mailboxes = Get-Mailbox -ResultSize Unlimited | sort PrimarySmtpAddress

    $AllUsers = @()
    $SitesNotFound = @()
    #ProgressBar
    $progressref = ($Mailboxes).count
    $progresscounter = 0

    foreach ($user in $Mailboxes)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($user.DisplayName)"

        Write-Host "$($user.DisplayName) .." -ForegroundColor Cyan -NoNewline

        $MBXStats = Get-MailboxStatistics $user.primarysmtpaddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
        $addresses = $user | select -ExpandProperty EmailAddresses
        $MSOLUser = Get-MsolUser -userprincipalname $user.userprincipalname

        $currentuser = new-object PSObject
        
        $currentuser | add-member -type noteproperty -name "DisplayName" -Value $msoluser.DisplayName
        $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $msoluser.userprincipalname
        $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msoluser.IsLicensed
        $currentuser | add-member -type noteproperty -name "City" -Value $msoluser.City
        $currentuser | add-member -type noteproperty -name "Country" -Value $msoluser.Country
        $currentuser | add-member -type noteproperty -name "Department" -Value $msoluser.Department
        $currentuser | add-member -type noteproperty -name "Fax" -Value $msoluser.Fax
        $currentuser | add-member -type noteproperty -name "FirstName" -Value $msoluser.FirstName
        $currentuser | add-member -type noteproperty -name "LastName" -Value $msoluser.LastName
        $currentuser | add-member -type noteproperty -name "MobilePhone" -Value $msoluser.MobilePhone
        $currentuser | add-member -type noteproperty -name "Office" -Value $msoluser.Office
        $currentuser | add-member -type noteproperty -name "PhoneNumber" -Value $msoluser.PhoneNumber
        $currentuser | add-member -type noteproperty -name "PostalCode" -Value $msoluser.PostalCode
        $currentuser | add-member -type noteproperty -name "State" -Value $msoluser.State
        $currentuser | add-member -type noteproperty -name "StreetAddress" -Value $msoluser.StreetAddress
        $currentuser | add-member -type noteproperty -name "Title" -Value $msoluser.Title
        
        $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.PrimarySmtpAddress
        $currentuser | add-member -type noteproperty -name "WhenCreated" -Value $user.WhenCreated
        $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($addresses -join ",")
        $currentuser | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $user.legacyexchangedn)
        $currentuser | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value ($user.AcceptMessagesOnlyFrom -join ",")
        $currentuser | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($user.GrantSendOnBehalfTo -join ",")
        $currentuser | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $user.HiddenFromAddressListsEnabled
        $currentuser | add-member -type noteproperty -name "RejectMessagesFrom" -Value ($user.RejectMessagesFrom -join ",")
        $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $user.DeliverToMailboxAndForward
        $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $user.ForwardingAddress
        $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $user.ForwardingSmtpAddress
        $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $user.RecipientTypeDetails
        $currentuser | add-member -type noteproperty -name "Alias" -Value $user.alias
        $currentuser | add-member -type noteproperty -name "ExchangeGuid" -Value $user.ExchangeGuid
        $currentuser | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize
        $currentuser | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount
        $currentuser | Add-Member -Type NoteProperty -Name "ArchiveGUID" -Value $user.ArchiveGuid
        $currentuser | add-member -type noteproperty -name "ArchiveState" -Value $user.ArchiveState
        $currentuser | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $user.ArchiveStatus

        if ($ArchiveStats = Get-MailboxStatistics $user.primarysmtpaddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount)
        {
            Write-Host "Archive found ..." -ForegroundColor green -NoNewline
            
            $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $ArchiveStats.TotalItemSize.Value
            $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $ArchiveStats.ItemCount
        }
        else
        {
            Write-Host "No Archive found ..." -ForegroundColor Red -NoNewline
            $currentuser | add-member -type noteproperty -name "ArchiveSize" -Value $null
            $currentuser | add-member -type noteproperty -name "ArchiveItemCount" -Value $null
        }
        if ($OneDriveCheck) 
        {
            try {
                #Get OneDrive Site details      
                $SPOSite = $null
                $EmailAddressUpdate1 = $MSOLUser.UserPrincipalName.Replace("@","_")
                $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
                $ODSite = $OneDriveURL + $EmailAddressUpdate2
                $SPOSITE = Get-SPOSITE $ODSite -ErrorAction SilentlyContinue
            }
            catch {
                Write-Host "OneDrive Not Enabled for User ..." -ForegroundColor Yellow -NoNewline
                $SitesNotFound += $FDUser
            }
            if ($SPOSITE)
            {
                Write-Host "Gathering OneDrive Details ..." -ForegroundColor Cyan -NoNewline
                
                $currentuser | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $ODSite
                $currentuser | Add-Member -type NoteProperty -Name "Owner" -Value $SPOSITE.Owner
                $currentuser | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $SPOSITE.StorageUsageCurrent
                $currentuser | Add-Member -type NoteProperty -Name "Status" -Value $SPOSITE.Status
                $currentuser | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $SPOSITE.SiteDefinedSharingCapability
                $currentuser | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $FDUser.LimitedAccessFileType           
            }
            else 
            {
                $currentuser | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $null
                $currentuser | Add-Member -type NoteProperty -Name "Owner" -Value $null
                $currentuser | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $null
                $currentuser | Add-Member -type NoteProperty -Name "Status" -Value $null
                $currentuser | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $null
                $currentuser | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $null
            }
        }  
        Write-Host "done" -ForegroundColor Green
        $AllUsers += $currentuser
    }
    #Export
    $AllUsers | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}

# Get Source Mailbox Details based on Customer Provided CSV
function Get-AllMailUsers {
    param (
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$true)] [array] $ImportCSV,
        [Parameter(Mandatory=$False)] [Switch] $OneDriveCheck,
        [Parameter(Mandatory=$False)] [string] $OneDriveURL
    )
    $ImportedUsers = Import-Csv $ImportCSV
    $AllUsers = @()
    $SitesNotFound = @()
    #ProgressBar
    $progressref = ($ImportedUsers).count
    $progresscounter = 0

    foreach ($user in $ImportedUsers)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($user.DisplayName_OG)"

        Write-Host "$($user.DisplayName_OG) .." -ForegroundColor Cyan -NoNewline
        $SourceEmail = $user.SourceEmail.trim()

        $recipientCheck = Get-Recipient $SourceEmail -ErrorAction SilentlyContinue
        if ($MBX = Get-Mailbox $recipientCheck.PrimarySmtpAddress -ErrorAction SilentlyContinue) {
            $MBXStats = Get-MailboxStatistics $MBX.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
            $addresses = $MBX.EmailAddresses
            $MSOLUser = Get-MsolUser -userprincipalname $MBX.UserPrincipalName
        }
        else {
            $MBXStats = $null
            $addresses = $recipientCheck.EmailAddresses
            $MSOLUser = Get-MsolUser -SearchString $recipientCheck.DisplayName
        }
        
        
        $user | add-member -type noteproperty -name "DisplayName" -Value $msoluser.DisplayName
        $user | add-member -type noteproperty -name "UserPrincipalName" -Value $msoluser.userprincipalname
        $user | add-member -type noteproperty -name "IsLicensed" -Value $msoluser.IsLicensed
        $user | add-member -type noteproperty -name "City" -Value $msoluser.City
        $user | add-member -type noteproperty -name "Country" -Value $msoluser.Country
        $user | add-member -type noteproperty -name "Department" -Value $msoluser.Department
        $user | add-member -type noteproperty -name "Fax" -Value $msoluser.Fax
        $user | add-member -type noteproperty -name "FirstName" -Value $msoluser.FirstName
        $user | add-member -type noteproperty -name "LastName" -Value $msoluser.LastName
        $user | add-member -type noteproperty -name "MobilePhone" -Value $msoluser.MobilePhone
        $user | add-member -type noteproperty -name "Office" -Value $msoluser.Office
        $user | add-member -type noteproperty -name "PhoneNumber" -Value $msoluser.PhoneNumber
        $user | add-member -type noteproperty -name "PostalCode" -Value $msoluser.PostalCode
        $user | add-member -type noteproperty -name "State" -Value $msoluser.State
        $user | add-member -type noteproperty -name "StreetAddress" -Value $msoluser.StreetAddress
        $user | add-member -type noteproperty -name "Title" -Value $msoluser.Title
        
        $user | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $MBX.PrimarySmtpAddress
        $user | add-member -type noteproperty -name "WhenCreated" -Value $MBX.WhenCreated
        $user | add-member -type noteproperty -name "EmailAddresses" -Value ($addresses -join ",")
        $user | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $MBX.legacyexchangedn)
        $user | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value ($MBX.AcceptMessagesOnlyFrom -join ",")
        $user | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($MBX.GrantSendOnBehalfTo -join ",")
        $user | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $MBX.HiddenFromAddressListsEnabled
        $user | add-member -type noteproperty -name "RejectMessagesFrom" -Value ($MBX.RejectMessagesFrom -join ",")
        $user | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $MBX.DeliverToMailboxAndForward
        $user | add-member -type noteproperty -name "ForwardingAddress" -Value $MBX.ForwardingAddress
        $user | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $MBX.ForwardingSmtpAddress
        $user | add-member -type noteproperty -name "RecipientTypeDetails" -Value $recipientCheck.RecipientTypeDetails
        $user | add-member -type noteproperty -name "Alias" -Value $MBX.alias
        $user | add-member -type noteproperty -name "ExchangeGuid" -Value $MBX.ExchangeGuid
        $user | Add-Member -type NoteProperty -Name "MBXSize" -Value $MBXStats.TotalItemSize
        $user | Add-Member -Type NoteProperty -name "MBXItemCount" -Value $MBXStats.ItemCount
        $user | Add-Member -Type NoteProperty -Name "ArchiveGUID" -Value $MBX.ArchiveGuid
        $user | add-member -type noteproperty -name "ArchiveState" -Value $MBX.ArchiveState
        $user | Add-Member -Type NoteProperty -Name "ArchiveStatus" -Value $MBX.ArchiveStatus

        if ($ArchiveStats = Get-MailboxStatistics $sourceEmailAddress -Archive -ErrorAction silentlycontinue | select TotalItemSize, ItemCount)
        {
            Write-Host "Archive found ..." -ForegroundColor green -NoNewline
            
            $user | add-member -type noteproperty -name "ArchiveSize" -Value $ArchiveStats.TotalItemSize.Value
            $user | add-member -type noteproperty -name "ArchiveItemCount" -Value $ArchiveStats.ItemCount
        }
        else
        {
            Write-Host "No Archive found ..." -ForegroundColor Red -NoNewline
            $user | add-member -type noteproperty -name "ArchiveSize" -Value $null
            $user | add-member -type noteproperty -name "ArchiveItemCount" -Value $null
        }
        if ($OneDriveCheck) 
        {
            try {
                #Get OneDrive Site details      
                $SPOSite = $null
                $EmailAddressUpdate1 = $MSOLUser.UserPrincipalName.Replace("@","_")
                $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
                $ODSite = $OneDriveURL + $EmailAddressUpdate2
                $SPOSITE = Get-SPOSITE $ODSite -ErrorAction SilentlyContinue
            }
            catch {
                Write-Host "OneDrive Not Enabled for User ..." -ForegroundColor Yellow -NoNewline
                $SitesNotFound += $FDUser
            }
            if ($SPOSITE)
            {
                Write-Host "Gathering OneDrive Details ..." -ForegroundColor Cyan -NoNewline
                
                $user | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $ODSite
                $user | Add-Member -type NoteProperty -Name "Owner" -Value $SPOSITE.Owner
                $user | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $SPOSITE.StorageUsageCurrent
                $user | Add-Member -type NoteProperty -Name "Status" -Value $SPOSITE.Status
                $user | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $SPOSITE.SiteDefinedSharingCapability
                $user | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $FDUser.LimitedAccessFileType           
            }
            else 
            {
                $user | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $null
                $user | Add-Member -type NoteProperty -Name "Owner" -Value $null
                $user | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $null
                $user | Add-Member -type NoteProperty -Name "Status" -Value $null
                $user | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $null
                $user | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $null
            }
        }  
        Write-Host "done" -ForegroundColor Green
        $AllUsers += $user
    }
    #Export
    $AllUsers | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}

# Get Source Distribution Details based on Customer Provided CSV
function Get-AllMailGroups {
    param (
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$true)] [array] $ImportCSV
    )
    $ImportedGroups = Import-Csv $ImportCSV
    $AllGroups = @()
    $failures = @()
    #ProgressBar
    $progressref = ($ImportedGroups).count
    $progresscounter = 0

    foreach ($group in $ImportedGroups)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Distribution Group Stats for $($user.SourceDisplayName)"

        Write-Host "$($group.SourceDisplayName) .." -ForegroundColor Cyan -NoNewline
        try {
            $sourceEmailAddress = $group.SourceEmail.trim()
            $DistributionGroup = Get-DistributionGroup $sourceEmailAddress -ErrorAction Stop
            $addresses = $DistributionGroup.EmailAddresses
            $DistributionGroupMembers = (Get-DistributionGroupMember $sourceEmailAddress -ErrorAction Stop).PrimarySMTPAddress -join ","
            $addresses = $DistributionGroup.EmailAddresses
            
            $group | add-member -type noteproperty -name "DisplayName" -Value $DistributionGroup.DisplayName       
            $group | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $DistributionGroup.PrimarySmtpAddress
            $group | add-member -type noteproperty -name "WhenCreated" -Value $DistributionGroup.WhenCreated
            $group | add-member -type noteproperty -name "EmailAddresses" -Value ($addresses -join ",")
            $group | add-member -type noteproperty -name "LegacyExchangeDN" -Value ("x500:" + $DistributionGroup.legacyexchangedn)
            $group | add-member -type noteproperty -name "AcceptMessagesOnlyFrom" -Value ($DistributionGroup.AcceptMessagesOnlyFrom -join ",")
            $group | add-member -type noteproperty -name "GrantSendOnBehalfTo" -Value ($DistributionGroup.GrantSendOnBehalfTo -join ",")
            $group | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $DistributionGroup.HiddenFromAddressListsEnabled
            $group | add-member -type noteproperty -name "RejectMessagesFrom" -Value ($DistributionGroup.RejectMessagesFrom -join ",")
            $group | add-member -type noteproperty -name "RecipientTypeDetails" -Value $DistributionGroup.RecipientTypeDetails
            $group | add-member -type noteproperty -name "Alias" -Value $DistributionGroup.alias -Force
            $group | add-member -type noteproperty -name "ExchangeGuid" -Value $DistributionGroup.ExchangeGuid
            $group | add-member -type noteproperty -name "Members" -Value $DistributionGroupMembers
            Write-Host "done" -ForegroundColor Green
            $AllGroups += $group
        }
        catch
        {
            Write-Warning -Message "$($_.Exception)"
            $currenterror = new-object PSObject

            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
            $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
            $currenterror | add-member -type noteproperty -name "DistributionList" -Value $sourceEmailAddress
            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
            $failures += $currenterror
        }
        
    }
    #Export
    $AllGroups | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
    $failures | Export-Csv -NoTypeInformation -Encoding utf8 "C:\Users\amedrano\OneDrive - Arraya Solutions\Office365 Customers\Magnetrol Groups\Groups_Failures.csv"
}

## License Update
$customerMatchedUsers = Import-Csv "C:\Users\amedrano\OneDrive - Arraya Solutions\Office365 Customers\Magnetrol Groups\Customer_Provided Magnetrol_Mailboxes.csv"
$licenseUpdateUsers = $customerMatchedUsers | ?{$_.Licenses_Destination -like '*AMETEKInc:SPE_E3*'}
foreach($user in $licenseUpdateUsers)
{
    $msolUserCheck = Get-MsolUser -UserPrincipalName $user.UserPrincipalName_Destination
    if($msolUserCheck.licenses.accountSKUID -match "AMETEKInc:SPE_E3")
    {
        $DisabledArray = @()
        $allLicenses = ($msolUserCheck).Licenses
        for($i = 0; $i -lt $AllLicenses.Count; $i++)
        {
            $serviceStatus =  $AllLicenses[$i].ServiceStatus
            foreach($service in $serviceStatus)
            {
                if($service.ProvisioningStatus -eq "Disabled")
                {
                    $disabledArray += ($service.ServicePlan).ServiceName
                }
            }
        }
        $updatedDisableArray = $disabledArray |? {$_ -ne "EXCHANGE_S_ENTERPRISE"}
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId "AMETEKInc:SPE_E3" -DisabledPlans $updatedDisableArray -Verbose
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
        Write-Host "Enabled Exchange Service for $($user.UserPrincipalName_Destination)" -ForegroundColor Green
    }
}

#Set Forwarding back to Magnetrol
$customerMatchedUsers = Import-Csv "C:\Users\amedrano\OneDrive - Arraya Solutions\Office365 Customers\Magnetrol Groups\Customer_Provided Magnetrol_Mailboxes.csv"
$licenseUpdateUsers = $customerMatchedUsers | ?{$_.Licenses_Destination -like '*AMETEKInc:SPE_E3*'}

$updatedUsers = @()
$notUpdatedUsers = @()
foreach ($user in $licenseUpdateUsers)
{
    $magnetrolAddress = $user."Source Email Uniform Case"
    if ($mailboxcheck = Get-Mailbox $user.PrimarySMTPAddress_Destination -ea silentlycontinue)
    {
        Set-Mailbox $mailboxcheck.PrimarySmtpAddress -ForwardingSmtpAddress $magnetrolAddress
        Write-Host "Set Forwarding from $($user.PrimarySMTPAddress_Destination) to $($magnetrolAddress)" -ForegroundColor Green
        $updatedUsers += $user
        
    }
    else 
    {
        Write-Host "No Mailbox found for $($user.PrimarySMTPAddress_Destination)"
        $notUpdatedUsers += $user
    }
}
## Update Again
$updatedUsers2 = @()
$notUpdatedUsers2 = @()
foreach ($user in $notUpdatedUsers)
{
    $magnetrolAddress = $user."Source Email Uniform Case"
    if ($mailboxcheck = Get-Mailbox $user.PrimarySMTPAddress -ea silentlycontinue)
    {
        Set-Mailbox $mailboxcheck.PrimarySmtpAddress -ForwardingSmtpAddress $magnetrolAddress
        Write-Host "Set Forwarding from $($user.PrimarySMTPAddress) to $($magnetrolAddress)" -ForegroundColor Green
        $updatedUsers2 += $user
        
    }
    else 
    {
        Write-Host "No Mailbox found for $($user.PrimarySMTPAddress)"
        $notUpdatedUsers2 += $user
    }
}
#


##### Resource and Shared Mailbox START REGION #####
## Create Resources for Magnetrol

$createdResources = @()
$createdADusers = @()
foreach ($room in $resources)
{
    Write-Host "Checking for Resource $($room.DestinationDisplayName) ... " -NoNewline -ForegroundColor Cyan

    $destinationEmail = $room.DestinationEmail
    if ($ADUserCheck = Get-ADUser -Filter {UserPrincipalName -eq $destinationEmail} -ea silentlycontinue)
    {
        Write-Host "AD User already exists. Checking for Remote Room Mailbox " -ForegroundColor Yellow -NoNewline
        if (!($roomMailboxCheck = Get-RemoteMailbox $destinationEmail -ea silentlycontinue))
        {
            $addressSplit = $room.DestinationEmail -split "@"
            $remoteRoutingAddress = $addressSplit[0] + "@ametekinc.mail.onmicrosoft.com"
            Enable-RemoteMailbox $ADUserCheck.DistinguishedName -Room -RemoteRoutingAddress $remoteRoutingAddress
            Write-Host "Created Successfully." -ForegroundColor Green
            $createdResources += $room
        }
        else {
            Write-Host "Already Exists." -ForegroundColor Yellow
        }
    }
    else 
    {
        $OUCheck = Get-OrganizationalUnit $room.OU
        $resourceOU = "OU=Resources,"+ $OUCheck.DistinguishedName
        $mail = $room.DestinationEmail
        $createdADusers += $room
        if ($resourceOU) {
            #New-ADUser -path $resourceOU -EmployeeID Resource -DisplayName $room.DestinationDisplayName -name $room.DestinationDisplayName -OtherAttributes @{'mail'=$mail}
            
            Write-Host "New User Created" -ForegroundColor Green
        }
        else {
            Write-Host "No OU found for Resource" -ForegroundColor red
        } 
    }  
}

## Create SharedMailboxes for Magnetrol

$createdSharedMailboxes = @()
$createdADusers = @()
foreach ($mailbox in $sharedMailboxes)
{
    Write-Host "Checking for AD User $($mailbox.DestinationDisplayName) ... " -NoNewline -ForegroundColor Cyan

    $destinationEmail = $mailbox.DestinationEmail
    if ($ADUserCheck = Get-ADUser -Filter {UserPrincipalName -eq $destinationEmail} -ea silentlycontinue)
    {
        Write-Host "AD User already exists. Checking for Shared Mailbox " -ForegroundColor Yellow -NoNewline
        if (!($roomMailboxCheck = Get-RemoteMailbox $destinationEmail -ea silentlycontinue))
        {
            $addressSplit = $mailbox.DestinationEmail -split "@"
            $remoteRoutingAddress = $addressSplit[0] + "@ametekinc.mail.onmicrosoft.com"
            Enable-RemoteMailbox $ADUserCheck.DistinguishedName -Shared -RemoteRoutingAddress $remoteRoutingAddress
            Write-Host "Created Successfully." -ForegroundColor Green
            $createdSharedMailboxes += $mailbox
        }
        else {
            Write-Host "Already Exists." -ForegroundColor Yellow
        }
    }
    else 
    {
        $OUCheck = Get-OrganizationalUnit $mailbox.OU
        $sharedMailboxOU = "OU=Shared Mailboxes,"+ $OUCheck.DistinguishedName
        $mail = $mailbox.DestinationEmail
        $createdADusers += $mailbox
        if ($resourceOU) {
            #New-ADUser -path $sharedMailboxOU -EmployeeID Resource -DisplayName $mailbox.DestinationDisplayName -name $mailbox.DestinationDisplayName -OtherAttributes @{'mail'=$mail}
            
            Write-Host "New User Created" -ForegroundColor Green
        }
        else {
            Write-Host "No OU found for SharedMailbox" -ForegroundColor red
        } 
    }  
}

## Create DistributionGroups for Magnetrol

$createdDistributionGroups = @()
$createdADGroups = @()
foreach ($group in $groups)
{
    $destinationEmail = $group.'DestinationEmailAddress '.Trim()
    $displayName = $group.DestinationDGName.Trim()

    Write-Host "Checking for in AD for $($displayName) ... " -NoNewline -ForegroundColor Cyan
    if ($adGroupCheck = Get-ADGroup -Filter {Mail -eq $destinationEmail} -ea silentlycontinue)
    {
        Write-Host "AD User already exists. Checking for Distribution List " -ForegroundColor Yellow
        if (!($DLCheck = Get-DistributionGroup $destinationEmail -ea silentlycontinue))
        {
            Enable-DistributionGroup $adGroupCheck.DistinguishedName -PrimarySMTPAddress $destinationEmail -DisplayName $displayName -alias $group.alias
            Write-Host "Created Successfully." -ForegroundColor Green
            $createdDistributionGroups += $group
        }
        else {
            Write-Host "Already Exists." -ForegroundColor Yellow
        }
    }
    else 
    {
        $OUCheck = Get-OrganizationalUnit $group.OU
        $distributionListOU = "OU=Distribution Lists,"+ $OUCheck.DistinguishedName
        $createdADGroups += $group
        if ($distributionListOU) {
            New-DistributionGroup -OrganizationalUnit $distributionListOU -DisplayName $displayName -name $displayName -PrimarySMTPAddress $destinationEmail -ManagedBy $group.ManagedBy
            Write-Host "New Group Created" -ForegroundColor Green
        }
        else {
            Write-Host "No OU found for Group" -ForegroundColor red
        } 
    }  
}

#Create Permission Mail Enabled Security Groups for resources mailboxes
foreach ($mailbox in $resources)
{
    Write-Host "Creating Resource Permissions Groups $($mailbox.DestinationDisplayName) ... " -NoNewline -ForegroundColor Cyan

    $OUCheck = Get-OrganizationalUnit $mailbox.OU
    $addressSplit = $mailbox.DestinationEmail -split "@"
    $sharedMailboxOU = "OU=Rooms,"+ $OUCheck.DistinguishedName
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"
    $FullAccessResourceEmailAddress = $addressSplit[0] + "_FullAccess@ametek.com"
    $SendAsResourceEmailAddress= $addressSplit[0] + "_SendAs@ametek.com"
    
    New-DistributionGroup -Type security -Name $FullAccessResourceName -DisplayName $FullAccessResourceName -PrimarySMTPAddress $FullAccessResourceEmailAddress -OrganizationalUnit $sharedMailboxOU
    New-DistributionGroup -Type security -Name $SendAsResourceName -DisplayName $SendAsResourceName -PrimarySMTPAddress $SendAsResourceEmailAddress -OrganizationalUnit $sharedMailboxOU
    
    Write-Host "done" -ForegroundColor Green
}

#Create Permission Mail Enabled Security Groups for shared mailboxes - Updated
foreach ($mailbox in $sharedMailboxes)
{
    Write-Host "Creating Shared Mailbox Permissions Groups $($mailbox.DestinationDisplayName) ... " -NoNewline -ForegroundColor Cyan

    $OUCheck = Get-OrganizationalUnit $mailbox.OU
    $addressSplit = $mailbox.DestinationEmail -split "@"
    $sharedMailboxOU = "OU=Shared Mailboxes,"+ $OUCheck.DistinguishedName
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"
    $FullAccessResourceEmailAddress = $addressSplit[0] + "_FullAccess@ametek.com"
    $SendAsResourceEmailAddress= $addressSplit[0] + "_SendAs@ametek.com"
    
    New-DistributionGroup -Type security -Name $FullAccessResourceName -DisplayName $FullAccessResourceName -PrimarySMTPAddress $FullAccessResourceEmailAddress -OrganizationalUnit $sharedMailboxOU
    New-DistributionGroup -Type security -Name $SendAsResourceName -DisplayName $SendAsResourceName -PrimarySMTPAddress $SendAsResourceEmailAddress -OrganizationalUnit $sharedMailboxOU
    
    Write-Host "done" -ForegroundColor Green
}

#Update Permission Mail Enabled Security Groups for resources and shared mailboxes
foreach ($mailbox in $resources)
{
    Write-Host "Updating Resource Mailbox Permissions Groups $($mailbox.DestinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    $addressSplit = $mailbox.DestinationEmail -split "@"

    $OLDFullAccessName = $mailbox.ou + "-gg-" + $addressSplit[0] + "_FullAccess"
    $OLDSendAccessName = $mailbox.ou + "-gg-" + $addressSplit[0] + "_SendAs"
    $FullAccessResourceEmailAddress = $addressSplit[0] + "_FullAccess@ametek.com"
    $SendAsResourceEmailAddress = $addressSplit[0] + "_SendAs@ametek.com"
    $SAMaccountSendAs  = $mailbox.DestinationDisplayName + "_SendAs"
    $SAMaccountFullAccess = $mailbox.DestinationDisplayName + "_FullAccess"
    $recipientCheck = @()
    if ($recipientCheck = Get-recipient $OLDFullAccessName -ea silentlycontinue) {
        Write-Host "$($OLDFullAccessName) found and updating"
        Set-DistributionGroup -Identity $recipientCheck.primarysmtpaddress.tostring() -name $SAMaccountFullAccess -DisplayName $SAMaccountFullAccess -SamAccountName $SAMaccountFullAccess #-PrimarySMTPAddress $FullAccessResourceEmailAddress
        Write-Host "done" -ForegroundColor Green    
    }
    else {
        Write-Host "$($OLDFullAccessName) not found"
    }
    if ($recipientCheck = Get-recipient $OLDSendAccessName -ea silentlycontinue) {
        Write-Host "$($OLDSendAccessName) found and updating"
        Set-DistributionGroup -Identity $recipientCheck.primarysmtpaddress.tostring() -name $SAMaccountSendAs -DisplayName $SAMaccountSendAs -SamAccountName $SAMaccountSendAs  #-PrimarySMTPAddress $SendAsResourceEmailAddress
        Write-Host "done" -ForegroundColor Green 
    }
    else {
        Write-Host "$($OLDSendAccessName) not found"
    }
    
}

#Update Permission Mail Enabled Security Groups for resources and shared mailboxes
foreach ($mailbox in $sharedMailboxes)
{
    Write-Host "Updating Shared Mailbox Permissions Groups $($mailbox.DestinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    $addressSplit = $mailbox.DestinationEmail -split "@"

    $OLDFullAccessName = $mailbox.ou + "-gg-" + $addressSplit[0] + "_FullAccess"
    $OLDSendAccessName = $mailbox.ou + "-gg-" + $addressSplit[0] + "_SendAs"
    $FullAccessResourceEmailAddress = $addressSplit[0] + "_FullAccess@ametek.com"
    $SendAsResourceEmailAddress = $addressSplit[0] + "_SendAs@ametek.com"
    $SAMaccountSendAs  = $mailbox.DestinationDisplayName + "_SendAs"
    $SAMaccountFullAccess = $mailbox.DestinationDisplayName + "_FullAccess"
    $recipientCheck = @()
    if ($recipientCheck = Get-recipient $OLDFullAccessName -ea silentlycontinue) {
        Write-Host "$($OLDFullAccessName) found and updating"
        Set-DistributionGroup -Identity $recipientCheck.primarysmtpaddress.tostring() -PrimarySMTPAddress $FullAccessResourceEmailAddress -name $SAMaccountFullAccess -DisplayName $SAMaccountFullAccess -SamAccountName $SAMaccountFullAccess
        Write-Host "done" -ForegroundColor Green    
    }
    else {
        Write-Host "$($OLDFullAccessName) not found"
    }
    if ($recipientCheck = Get-recipient $OLDSendAccessName -ea silentlycontinue) {
        Write-Host "$($OLDSendAccessName) found and updating"
        Set-DistributionGroup -Identity $recipientCheck.primarysmtpaddress.tostring() -PrimarySMTPAddress $SendAsResourceEmailAddress -name $SAMaccountSendAs -DisplayName $SAMaccountSendAs -SamAccountName $SAMaccountSendAs     
        Write-Host "done" -ForegroundColor Green 
    }
    else {
        Write-Host "$($OLDSendAccessName) not found"
    }
    
}

#Update Visible in GAL Permission Mail Enabled Security Groups for resources
foreach ($mailbox in $resources)
{
    Write-Host "Update Resource Visible in GAL Permissions Groups $($mailbox.DestinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"
    
    Set-DistributionGroup -Identity $FullAccessResourceName -HiddenFromAddressListsEnabled $True
    Set-DistributionGroup -Identity $SendAsResourceName -HiddenFromAddressListsEnabled $True
    
    Write-Host "done" -ForegroundColor Green
}

#Update Visible in GAL Permission Mail Enabled Security Groups for shared mailboxes
foreach ($mailbox in $sharedMailboxes)
{
    Write-Host "Update Shared Mailbox Visible in GAL Permissions Groups $($mailbox.DestinationDisplayName) ... " -NoNewline -ForegroundColor Cyan
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"
    
    Set-DistributionGroup -Identity $FullAccessResourceName -HiddenFromAddressListsEnabled $True
    Set-DistributionGroup -Identity $SendAsResourceName -HiddenFromAddressListsEnabled $True

    Write-Host "done" -ForegroundColor Green
}

### Add Perms - Resources with Security Groups

foreach ($mailbox in $resources) 
{
    Write-Host "Updating Perms for $($mailbox.DestinationEmail) .. " -ForegroundColor cyan -NoNewline
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"
    #Remove Current Mailbox Permissions
    $fullAccessPerms = Get-MailboxPermission $mailbox.DestinationEmail | ?{$_.user -notlike "*nt authority*"}

    foreach ($perm in $fullAccessPerms) {
        Remove-MailboxPermission -Identity $mailbox.DestinationEmail -User $perm.User -AccessRights FullAccess -Confirm:$false
        Write-Host ". " -ForegroundColor Yellow -NoNewline
    }
    #Remove Current Send-As Perms
    $SendAsPerms = Get-RecipientPermission $sharedMailboxes[0].DestinationEmail -AccessRights SendAs | ?{$_.Trustee -notlike "*nt authority*"}
    foreach ($perm in $SendAsPerms) {
        Remove-RecipientPermission -Identity $mailbox.DestinationEmail -Trustee $perm.Trustee -AccessRights SendAs -Confirm:$false
        Write-Host ". " -ForegroundColor Yellow -NoNewline
    }

    #Add Full Access Permission
    Add-MailboxPermission -AccessRights FullAccess -Identity $mailbox.DestinationEmail -User $FullAccessResourceName -Confirm:$false
    Add-RecipientPermission -AccessRights SendAs -Identity $mailbox.DestinationEmail -Trustee $SendAsResourceName -Confirm:$false
    Write-Host "Succeeded " -ForegroundColor Green
}

## Update Permissions Group Membership
$sharedMailboxes = Import-Csv
$matchedUsers = Import-Csv

$notFoundPermUser = @()
$failures = @()
foreach ($mailbox in $sharedMailboxes | ?{$_.FullAccess}) 
{
    Write-Host "Updating Mailbox Perms for $($mailbox.DestinationEmail).. " -NoNewline -ForegroundColor Cyan
    $addressSplit = $mailbox.DestinationEmail -split "@"
    $FullAccessResourceEmailAddress = $addressSplit[0] + "_FullAccess@ametek.com"
    $SendAsResourceEmailAddress= $addressSplit[0] + "_SendAs@ametek.com"

    # Add Full Access Permission Users
    $FullAccessUsers = $mailbox.FullAccess -split ","
    Write-Host "Setting up $($FullAccessUsers.count) Users with Full Access .. " -NoNewline
    foreach ($perm in $FullAccessUsers) {
        # Match the Perm user
        $trimPermUser = $perm.trim()
        if ($matchedUser = $matchedUsers | Where-Object {$_.SourceEmail -eq $trimPermUser})
        {
            $destinationEmail = $matchedUser.DestinationEmail
            ## Check if Perm User Exists
            if ($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue)
            {
                try 
                {
                    #Add DL Members
                    Add-DistributionGroupMember $FullAccessResourceEmailAddress -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                    Write-Host ". " -ForegroundColor Green -NoNewline
                }
                catch
                {
                    Write-Host ". " -ForegroundColor Yellow -NoNewline
                    
                    #Build Error Array
                    $currenterror = new-object PSObject

                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                    $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                    $currenterror | add-member -type noteproperty -name "DistributionList" -Value $dFullAccessResourceEmailAddress
                    $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                    $failures += $currenterror
                }
            } 
        }
        else
        {
            $destinationEmail = $trimPermUser
            Write-Host "Not Matched Recipient $($destinationEmail) .. " -ForegroundColor red -NoNewline
            $notFoundPermUser += $trimPermUser
        }
    }

    # Add Send-As Permission Users
    $SendAsUsers = $mailbox.SendAs -split ","
    Write-Host "Setting up $($SendAsUsers.count) Users with Send-As .. " -NoNewline  
    foreach ($perm in $SendAsUsers) {
        # Match the Perm user
        $trimPermUser = $perm.trim()
        if ($matchedUser = $matchedUsers | Where-Object {$_.SourceEmail -eq $trimPermUser})
        {
            $destinationEmail = $matchedUser.DestinationEmail
            ## Check if Perm User Exists
            if ($recipientCheck = Get-Recipient $destinationEmail -ea silentlycontinue)
            {
                try 
                {
                    #Add DL Members
                    Add-DistributionGroupMember $SendAsResourceEmailAddress -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                    Write-Host ". " -ForegroundColor Green -NoNewline
                }
                catch
                {
                    Write-Host ". " -ForegroundColor Yellow -NoNewline
                    
                    #Build Error Array
                    $currenterror = new-object PSObject

                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                    $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                    $currenterror | add-member -type noteproperty -name "DistributionList" -Value $SendAsResourceEmailAddress
                    $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                    $failures += $currenterror
                }
            } 
        }
        else
        {
            $destinationEmail = $trimPermUser
            Write-Host "Not Matched Recipient $($destinationEmail) .. " -ForegroundColor red -NoNewline
            $notFoundPermUser += $trimPermUser
        }
    }
    Write-Host " done" -ForegroundColor Green
}

### Add Perms Shared Mailboxes with Security Groups

$failures = @()
foreach ($mailbox in $sharedMailboxes) 
{
    Write-Host "Updating Perms for $($mailbox.DestinationEmail) .. " -ForegroundColor cyan -NoNewline
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"
    #Remove Current Mailbox Permissions
    $fullAccessPerms = Get-MailboxPermission $mailbox.DestinationEmail | ?{$_.user -notlike "*nt authority*"}

    foreach ($perm in $fullAccessPerms) {
        Remove-MailboxPermission -Identity $mailbox.DestinationEmail -User $perm.User -AccessRights FullAccess -Confirm:$false
        Write-Host ". " -ForegroundColor Yellow -NoNewline
    }
    #Remove Current Send-As Perms
    $SendAsPerms = Get-RecipientPermission $sharedMailboxes[0].DestinationEmail -AccessRights SendAs | ?{$_.Trustee -notlike "*nt authority*"}
    foreach ($perm in $SendAsPerms) {
        Remove-RecipientPermission -Identity $mailbox.DestinationEmail -Trustee $perm.Trustee -AccessRights SendAs -Confirm:$false
        Write-Host ". " -ForegroundColor Yellow -NoNewline
    }

    #Add Full Access Permission
    Add-MailboxPermission -AccessRights FullAccess -Identity $mailbox.DestinationEmail -User $FullAccessResourceName -Confirm:$false
    Add-RecipientPermission -AccessRights SendAs -Identity $mailbox.DestinationEmail -Trustee $SendAsResourceName -Confirm:$false
    Write-Host "Succeeded " -ForegroundColor Green
}

# Add Perm Users to Perm Groups - Shared Mailboxes
foreach ($mailbox in $sharedMailboxes) {
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"
    Write-Host "Updating Full Access Perm Group for $($mailbox.DestinationEmail) .. " -ForegroundColor Cyan -NoNewline
    $FullAccessPermUsers = $mailbox.FullAccess -split ","
    $SendAsPermUsers = $mailbox.SendAs -split ","
    Write-Host "Adding $($FullAccessPermUsers.count) users with Full access to object .. " -ForegroundColor Yellow -NoNewline
    foreach ($fullAccessPerm in $FullAccessPermUsers) {
        Add-DistributionGroupMember -identity $FullAccessResourceName -Member $fullAccessPerm
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    Write-Host "Adding $($SendAsPermUsers.count) users with Full access to object .. " -ForegroundColor Yellow -NoNewline
    foreach ($sendAsPerm in $SendAsPermUsers) {
        Add-DistributionGroupMember -identity $SendAsResourceName -Member $sendAsPerm
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    Write-Host "done" -ForegroundColor Green
}

# Add Perm Users to Perm Groups - Resources
foreach ($mailbox in $resources) {
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"
    Write-Host "Updating Full Access Perm Group for $($mailbox.DestinationEmail) .. " -ForegroundColor Cyan -NoNewline
    $FullAccessPermUsers = $mailbox.FullAccess -split ","
    $SendAsPermUsers = $mailbox.SendAs -split ","
    Write-Host "Adding $($FullAccessPermUsers.count) users with Full access to object .. " -ForegroundColor Yellow -NoNewline
    foreach ($fullAccessPerm in $FullAccessPermUsers) {
        Add-DistributionGroupMember -identity $FullAccessResourceName -Member $fullAccessPerm
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    Write-Host "Adding $($SendAsPermUsers.count) users with Full access to object .. " -ForegroundColor Yellow -NoNewline
    foreach ($sendAsPerm in $SendAsPermUsers) {
        Add-DistributionGroupMember -identity $SendAsResourceName -Member $sendAsPerm
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    Write-Host "done" -ForegroundColor Green
}

# Set Full Access of Perm Group to Resource - Shared Mailboxes
foreach ($mailbox in $sharedMailboxes) {
    Write-Host "Updating Perms for $($mailbox.DestinationEmail) .. " -ForegroundColor cyan -NoNewline
    $FullAccessResourceName = $mailbox.DestinationDisplayName + "_FullAccess"
    $SendAsResourceName = $mailbox.DestinationDisplayName + "_SendAs"

    #Add Full Access Permission
    Add-MailboxPermission -AccessRights FullAccess -Identity $mailbox.DestinationEmail -User $FullAccessResourceName -Confirm:$false
    Add-RecipientPermission -AccessRights SendAs -Identity $mailbox.DestinationEmail -Trustee $SendAsResourceName -Confirm:$false
    Write-Host "Succeeded " -ForegroundColor Green
}

##### Resource and Shared Mailbox END REGION #####

#

### Allow External Senders
foreach ($group in ($groups | ?{$_.AcceptExternal -eq "yes"})) 
{
    $displayName = $group.DestinationDGName.Trim()
    Set-DistributionGroup $displayName -RequireSenderAuthenticationEnabled $false
    Write-Host "Updated External Senders Required for $($displayName)" -ForegroundColor Green
}

### Update Distribution Group Members
$matchedGroups = Import-Csv
$matchedUsers = Import-Csv
$DLProperties = Import-CSV

$notMatchedGroups = @()
$notFoundDistributionGroups = @()
$notFoundUsers = @()
$notMatchedMember = @()
$failedToAddMembers = @()
$failures = @()
foreach ($group in $DLProperties) 
{
    # Match the Group
    $sourceEmailAddress = $group.PrimarySmtpAddress
    if ($matchedGroup = $matchedGroups | Where-Object {$_.'EMAIL ADDRESS' -eq $sourceEmailAddress})
    {
        Write-Host "Matched Group $($group.Identity) .. " -ForegroundColor Green -NoNewline
        
        ## Check if Distribution Group Exists
        if ($dlCheck = Get-DistributionGroup $matchedGroup.'DestinationEmailAddress ' -ea silentlycontinue)
        {
            Write-Host "Distribution Group found .. " -ForegroundColor Green -NoNewline

            # Add Members
            $Members = $group.Members -split ","
            if ($Members)
            {
                Write-Host "Adding $($members.count) members .. " -NoNewline
                foreach ($member in $Members) 
                {
                    ## Match Members
                    if ($matchedUser = $matchedUsers | Where-Object {$_."Source Email Address" -eq $member})
                    {
                        ### Check if Users Exist
                        if ($recipientCheck = Get-Recipient $matchedUser.PrimarySmtpAddress -ea silentlycontinue)
                        {
                            try 
                            {
                                #Add DL Members
                                Add-DistributionGroupMember $dlcheck.PrimarySmtpAddress.ToString() -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                                Write-Host ". " -ForegroundColor Green -NoNewline
                            }
                            catch
                            {
                                Write-Host ". " -ForegroundColor Yellow -NoNewline
                                
                                #Build Error Array
                                $currenterror = new-object PSObject
        
                                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                                $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                                $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                                $currenterror | add-member -type noteproperty -name "DistributionList" -Value $dlcheck.PrimarySmtpAddress.ToString()
                                $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                                $failures += $currenterror
                            }
                        }
                        else 
                        {
                            Write-Host ". " -ForegroundColor Red -NoNewline
                            $notFoundUsers += $member    
                        }
                    }
                    else
                    {
                        Write-Host ". " -ForegroundColor Red -NoNewline
                        $notMatchedMember += $member
                    }
                }
                Write-Host "Completed" -ForegroundColor Green
            }
        }
        else {
            Write-Host "No Distribution Group Found" -ForegroundColor Red
            $notFoundDistributionGroups += $group
        }
    }
    else
    {
        Write-Host "Group Not Matched $($group.Identity)" -ForegroundColor Red

        #Build Error Array
        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value "Matching Group"
        $currenterror | add-member -type noteproperty -name "Reason" -Value "Unable to Match Group"
        $currenterror | add-member -type noteproperty -name "TargetName" -Value $group.Identity
        $currenterror | add-member -type noteproperty -name "DistributionList" -Value $group.PrimarySmtpAddress
        $currenterror | add-member -type noteproperty -name "Exception" -Value "$($group.DisplayName) Not Found. Unable to find $($group.PrimarySMTPAddress) in Customer Provided Spreadsheet"
        $failures += $currenterror
        $notMatchedGroups += $group
    }
}

### Update Distribution Group Members ATTEMPT 2 with Try statements
$matchedGroups = Import-Csv
$matchedUsers = Import-Csv
$DLProperties = Import-CSV

$notMatchedGroups = @()
$notFoundDistributionGroups = @()
$notFoundUsers = @()
$notMatchedMember = @()
$failedToAddMembers = @()
$failures = @()
foreach ($group in $DLProperties[0 .. 9]) 
{
    # Match the Group
    $sourceEmailAddress = $group.PrimarySmtpAddress
    if (
        $matchedGroup = $matchedGroups | Where-Object {$_.'EMAIL ADDRESS' -eq $sourceEmailAddress})
    {
        Write-Host "Matched Group $($group.Identity) .. " -ForegroundColor Green -NoNewline
        
        ## Check if Distribution Group Exists
        if ($dlCheck = Get-DistributionGroup $matchedGroup.'DestinationEmailAddress ' -ea silentlycontinue)
        {
            Write-Host "Distribution Group found .. " -ForegroundColor Green -NoNewline

            # Add Members
            $Members = $group.Members -split ","
            if ($Members)
            {
                Write-Host "Adding $($members.count) members .. " -NoNewline
                foreach ($member in $Members) 
                {
                    ## Match Members
                    if ($matchedUser = $matchedUsers | Where-Object {$_."Source Email Address" -eq $member})
                    {
                         ### Check if Users Exist
                        if ($recipientCheck = Get-Recipient $matchedUser.PrimarySmtpAddress -ea silentlycontinue)
                        {
                            try {
                                #Add DL Members
                                Add-DistributionGroupMember $dlcheck.PrimarySmtpAddress.ToString() -Member $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                                Write-Host ". " -ForegroundColor Green -NoNewline
                            }
                            catch
                            {
                                Write-Warning -Message "$($_.Exception)"
                                $failedToAddMembers += $_
                                $currenterror = new-object PSObject
        
                                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                                $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                                $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                                $currenterror | add-member -type noteproperty -name "DistributionList" -Value $dlcheck.PrimarySmtpAddress.ToString()
                                $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                                $failures += $currenterror
                            }
                        }
                        else 
                        {
                            Write-Host ". " -ForegroundColor Red -NoNewline
                            $notFoundUsers += $member    
                        }
                    }
                    else
                    {
                        Write-Host ". " -ForegroundColor Red -NoNewline
                        $notMatchedMember += $member
                    }
                }
                Write-Host "Completed" -ForegroundColor Green
            }
        }
        else {
            Write-Host "No Distribution Group Found" -ForegroundColor Red
            $notFoundDistributionGroups += $group
        }
    }
    else
    {
        Write-Host "Group Not Matched $($group.Identity)" -ForegroundColor Red
        $notMatchedGroups += $group
    }
}

### Add Full Access Perms ALL Mailboxes (Individual)

$matchedUsers = Import-Csv
$FullAccessPerms = Import-CSV

$notFoundUsers = @()
$notMatchedMember = @()
$failures = @()
foreach ($perm in $FullAccessPerms) 
{
    # Match the Group
    $sourceEmailAddress = $perm.Mailbox
    $sourcePermUser = $perm.UserWithFullAccess
    
    if ($matchedUser = $matchedUsers | Where-Object {$_.'SourceEmail' -eq $sourceEmailAddress})
    {
        Write-Host "Matched User $($sourceEmailAddress) .. " -ForegroundColor Green -NoNewline
        $destinationEmail = $matchedUser.DestinationEmail

        ## Check if Mailbox Exists
        if ($mailboxCheck = Get-Mailbox $destinationEmail -ea silentlycontinue)
        {
            Write-Host "First Mailbox found .. " -ForegroundColor Green -NoNewline

            ## Match Perm User
            if ($matchedPerm = $matchedUsers | Where-Object {$_.'SourceEmail' -eq $sourcePermUser})
            {
                $destinationPermUser = $matchedPerm.DestinationEmail
                ### Check if Users Exist
                if ($recipientCheck = Get-Recipient $destinationPermUser -ea silentlycontinue)
                {
                    try 
                    {
                        #Add Full Access Permission
                        Add-MailboxPermission -AccessRights FullAccess -Identity $mailboxCheck.PrimarySmtpAddress.ToString() -User $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                        Write-Host "Succeeded " -ForegroundColor Green
                    }
                    catch
                    {
                        Write-Host "Failed" -ForegroundColor Red
                        
                        #Build Error Array
                        $currenterror = new-object PSObject

                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                        $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                        $currenterror | add-member -type noteproperty -name "DistributionList" -Value $dlcheck.PrimarySmtpAddress.ToString()
                        $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                        $failures += $currenterror
                    }
                }
                else 
                {
                    Write-Host "Perm User $($destinationPermUser) Not Found " -ForegroundColor Red
                    $notFoundUsers += $destinationPermUser
                }
            }
            else  {
                Write-Host "User Not Matched $($sourcePermUser)" -ForegroundColor Red
                $notMatchedMember += $sourcePermUser
            }
        }
        else 
        {
            Write-Host "No User $($destinationEmail) Found" -ForegroundColor Red
            $notFoundUsers += $destinationEmail
        }
    }
    else  {
        Write-Host "User Not Matched $($sourceEmailAddress)" -ForegroundColor Red

        #Build Error Array
        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value "Matching Group"
        $currenterror | add-member -type noteproperty -name "Reason" -Value "Unable to Match Group"
        $currenterror | add-member -type noteproperty -name "TargetName" -Value $perm.Identity
        $currenterror | add-member -type noteproperty -name "DistributionList" -Value $perm.PrimarySmtpAddress
        $currenterror | add-member -type noteproperty -name "Exception" -Value "$($perm.DisplayName) Not Found. Unable to find $($perm.PrimarySMTPAddress) in Customer Provided Spreadsheet"
        $failures += $currenterror
        $notMatchedMember += $sourceEmailAddress
    }
}


### Add SendAs Access Perms - Shared Mailboxes (Individual)

$sharedMailboxes = Import-Csv

$notFoundUsers = @()
$notMatchedMember = @()
$failures = @()
foreach ($user in $sharedMailboxes) {
    # Match the User
    $sourceEmailAddress = $user.SourceEmail

    if ($matchedUser = $matchedUsers | Where-Object {$_.'SourceEmail' -eq $sourceEmailAddress}) {
        Write-Host "Matched User $($sourceEmailAddress) .. " -ForegroundColor Green -NoNewline
        $destinationEmail = $matchedUser.DestinationEmail

        ## Check if Mailbox Exists
        if ($mailboxCheck = Get-Mailbox $destinationEmail -ea silentlycontinue) {
            Write-Host "First Mailbox found .. " -ForegroundColor Green -NoNewline    
            
            $SendAsArray = $user.SendAs -split ","
            foreach ($perm in $SendAsArray) {
                ## Match Perm User
                if ($matchedPerm = $matchedUsers | Where-Object {$_.'SourceEmail' -eq $perm}) {
                    $destinationPermUser = $matchedPerm.DestinationEmail
                    ### Check if Users Exist
                    if ($recipientCheck = Get-Recipient $destinationPermUser -ea silentlycontinue) {
                        try {
                            #Add Full Access Permission
                            Add-RecipientPermission -AccessRights SendAs -Identity $mailboxCheck.PrimarySmtpAddress.ToString() -Trustee $recipientCheck.PrimarySMTPAddress.ToString() -ErrorAction Stop
                            Write-Host "Succeeded " -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Failed" -ForegroundColor Red
                            
                            #Build Error Array
                            $currenterror = new-object PSObject

                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | add-member -type noteproperty -name "Reason" -Value $_.CategoryInfo.Reason
                            $currenterror | add-member -type noteproperty -name "TargetName" -Value $_.CategoryInfo.TargetName
                            $currenterror | add-member -type noteproperty -name "DistributionList" -Value $dlcheck.PrimarySmtpAddress.ToString()
                            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                            $failures += $currenterror
                        }
                    }
                }
                else {
                    Write-Host "Perm User $($perm) Not Found " -ForegroundColor Red
                    $notFoundUsers += $perm
                }
            }
            else  {
                Write-Host "User Not Matched $($sourcePermUser)" -ForegroundColor Red
                $notMatchedMember += $sourcePermUser
            }
        }
        else {
            Write-Host "No User $($destinationEmail) Found" -ForegroundColor Red
            $notFoundUsers += $destinationEmail
        }
    }
    else  {
        Write-Host "User Not Matched $($sourceEmailAddress)" -ForegroundColor Red

        #Build Error Array
        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value "Matching Group"
        $currenterror | add-member -type noteproperty -name "Reason" -Value "Unable to Match Group"
        $currenterror | add-member -type noteproperty -name "TargetName" -Value $perm.Identity
        $currenterror | add-member -type noteproperty -name "DistributionList" -Value $perm.PrimarySmtpAddress
        $currenterror | add-member -type noteproperty -name "Exception" -Value "$($perm.DisplayName) Not Found. Unable to find $($perm.PrimarySMTPAddress) in Customer Provided Spreadsheet"
        $failures += $currenterror
        $notMatchedMember += $sourceEmailAddress
    }
}


## GET OFFICE365 Teams Details
function Get-AllOffice65GroupDetails {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$True)] [string] $InputCsvFile
        )

    $allOffice365Groups = @()
    $O365Groups = Import-Csv $InputCsvFile
    #ProgressBar
    $progressref = ($O365Groups).count
    $progresscounter = 0

    foreach ($group in $O365Groups)
    {
        $unifiedGroupSetings = Get-UnifiedGroup -Identity $group.GroupName

        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Office365 Group Stats for $($group.GroupName)"
        Write-Host "$($unifiedGroupSetings.DisplayName) .." -ForegroundColor Cyan -NoNewline

        $EmailAddresses = $unifiedGroupSetings.EmailAddresses -join ","
        $Members = Get-UnifiedGroupLinks -Identity $unifiedGroupSetings.Name -LinkType Member
        $MembersDisplayName = $Members.Name -join ","
        $MembersAddresses = $Members.PrimarySmtpAddress -join ","
        $Owners = Get-UnifiedGroupLinks -Identity $unifiedGroupSetings.Name -LinkType Owner
        $OwnersDisplayName = $Owners.Name -join ","
        $OwnersAddresses = $Owners.PrimarySmtpAddress -join ","
        $ManagedBy = $unifiedGroupSetings.ManagedBy -join ","
        
        $currentgroup = new-object PSObject
        $currentgroup | add-member -type noteproperty -name "DisplayName" -Value $unifiedGroupSetings.DisplayName -Force
        $currentgroup | add-member -type noteproperty -name "Name" -Value $unifiedGroupSetings.Name -Force
        $currentgroup | add-member -type noteproperty -name "IsMailboxConfigured" -Value $unifiedGroupSetings.IsMailboxConfigured -Force
        $currentgroup | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $unifiedGroupSetings.PrimarySmtpAddress -Force
        $currentgroup | add-member -type noteproperty -name "Alias" -Value $unifiedGroupSetings.Alias -Force
        $currentgroup | add-member -type noteproperty -name "RecipientTypeDetails" -Value $unifiedGroupSetings.RecipientTypeDetails -Force
        $currentgroup | add-member -type noteproperty -name "RequireSenderAuthenticationEnabled" -Value $unifiedGroupSetings.RequireSenderAuthenticationEnabled -Force
        $currentgroup | add-member -type noteproperty -name "GroupMemberCount" -Value $GroupMemberCount.GroupMemberCount -Force
        $currentgroup | add-member -type noteproperty -name "EmailAddresses" -Value $EmailAddresses -Force
        $currentgroup | add-member -type noteproperty -name "LegacyExchangeDN" -Value $unifiedGroupSetings.LegacyExchangeDN -Force
        $currentgroup | add-member -type noteproperty -name "MembersDisplayName" -Value $MembersDisplayName -Force
        $currentgroup | add-member -type noteproperty -name "MembersAddresses" -Value $MembersAddresses -Force
        $currentgroup | add-member -type noteproperty -name "ManagedBy" -Value $ManagedBy -Force
        $currentgroup | add-member -type noteproperty -name "OwnersDisplayName" -Value $OwnersDisplayName -Force
        $currentgroup | add-member -type noteproperty -name "OwnersAddresses" -Value $OwnersAddresses -Force
        $currentgroup | add-member -type noteproperty -name "WhenCreated" -Value $unifiedGroupSetings.WhenCreated -Force
        $currentgroup | add-member -type noteproperty -name "DestinationDisplayName" -Value $group.DestinationDisplayName -Force
        $currentgroup | add-member -type noteproperty -name "DestinationEmailAddress" -Value $group.DestinationEmailAddress -Force
        $currentgroup | add-member -type noteproperty -name "SourceSharePointSiteUrl" -Value $unifiedGroupSetings.SharePointSiteUrl
        $currentgroup | add-member -type noteproperty -name "SourceSharePointDocumentsUrl" -Value $unifiedGroupSetings.SharePointDocumentsUrl
        $currentgroup | add-member -type noteproperty -name "SourceSharePointNotebookUrl" -Value $unifiedGroupSetings.SharePointNotebookUrl


        $allOffice365Groups += $currentgroup
        Write-Host "done" -ForegroundColor Green
    }
    
    #Export
    $allOffice365Groups | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}


## Match OFFICE365 Teams Details Added
function Match-AllOffice65GroupDetails {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$True)] [string] $InputCsvFile
        )

    $allOffice365Groups = @()
    $O365Groups = Import-Csv $InputCsvFile
    #ProgressBar
    $progressref = ($O365Groups).count
    $progresscounter = 0

    foreach ($group in $O365Groups)
    {
        $unifiedGroupSetings = Get-UnifiedGroup -Identity $group.DestinationEmailAddress

        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Office365 Group Stats for $($group.DestinationEmailAddress)"
        Write-Host "$($unifiedGroupSetings.DisplayName) .." -ForegroundColor Cyan -NoNewline

        $EmailAddresses = $unifiedGroupSetings.EmailAddresses -join ","
        $Members = Get-UnifiedGroupLinks -Identity $unifiedGroupSetings.Name -LinkType Member
        $MembersDisplayName = $Members.Name -join ","
        $MembersAddresses = $Members.PrimarySmtpAddress -join ","
        $Owners = Get-UnifiedGroupLinks -Identity $unifiedGroupSetings.Name -LinkType Owner
        $OwnersDisplayName = $Owners.Name -join ","
        $OwnersAddresses = $Owners.PrimarySmtpAddress -join ","
        $ManagedBy = $unifiedGroupSetings.ManagedBy -join ","
        
        $group | add-member -type noteproperty -name "Destination_DisplayName" -Value $unifiedGroupSetings.DisplayName -Force
        $group | add-member -type noteproperty -name "DestinationName" -Value $unifiedGroupSetings.Name -Force
        $group | add-member -type noteproperty -name "DestinationIsMailboxConfigured" -Value $unifiedGroupSetings.IsMailboxConfigured -Force
        $group | add-member -type noteproperty -name "DestinationPrimarySmtpAddress" -Value $unifiedGroupSetings.PrimarySmtpAddress -Force
        $group | add-member -type noteproperty -name "DestinationAlias" -Value $unifiedGroupSetings.Alias -Force
        $group | add-member -type noteproperty -name "DestinationRecipientTypeDetails" -Value $unifiedGroupSetings.RecipientTypeDetails -Force
        $group | add-member -type noteproperty -name "DestinationRequireSenderAuthenticationEnabled" -Value $unifiedGroupSetings.RequireSenderAuthenticationEnabled -Force
        $group | add-member -type noteproperty -name "DestinationGroupMemberCount" -Value $GroupMemberCount.GroupMemberCount -Force
        $group | add-member -type noteproperty -name "DestinationEmailAddresses" -Value $EmailAddresses -Force
        $group | add-member -type noteproperty -name "DestinationLegacyExchangeDN" -Value $unifiedGroupSetings.LegacyExchangeDN -Force
        $group | add-member -type noteproperty -name "DestinationMembersDisplayName" -Value $MembersDisplayName -Force
        $group | add-member -type noteproperty -name "DestinationMembersAddresses" -Value $MembersAddresses -Force
        $group | add-member -type noteproperty -name "DestinationManagedBy" -Value $ManagedBy -Force
        $group | add-member -type noteproperty -name "DestinationOwnersDisplayName" -Value $OwnersDisplayName -Force
        $group | add-member -type noteproperty -name "DestinationOwnersAddresses" -Value $OwnersAddresses -Force
        $group | add-member -type noteproperty -name "DestinationWhenCreated" -Value $unifiedGroupSetings.WhenCreated -Force
        $group | add-member -type noteproperty -name "DestinationSharePointSiteUrl" -Value $unifiedGroupSetings.SharePointSiteUrl
        $group | add-member -type noteproperty -name "DestinationSharePointDocumentsUrl" -Value $unifiedGroupSetings.SharePointDocumentsUrl
        $group | add-member -type noteproperty -name "DestinationSharePointNotebookUrl" -Value $unifiedGroupSetings.SharePointNotebookUrl


        $allOffice365Groups += $group
        Write-Host "done" -ForegroundColor Green
    }
    
    #Export
    $allOffice365Groups | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}

## Fix borked Exchange users
    $Office365Users = @()
    #ProgressBar
    $progressref = ($magnetrolusers).count
    $progresscounter = 0

    foreach ($user in $magnetrolusers)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Office365 User License Stats for $($user.UserPrincipalName)"
        Write-Host "$($user.DisplayName) .. " -ForegroundColor Cyan -NoNewline

        if ($msolUserCheck = (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).licenses.servicestatus | ?{$_.ProvisioningStatus -eq "error"}) {
            $currentuser = new-object PSObject
            $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName -Force
            $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $user.UserPrincipalName -Force
    
            $Office365Users += $currentuser
            Write-Host "done" -ForegroundColor Green 
        }          
    }

## Create UnifiedGroups and Add Members

$Office365Groups = Import-Csv "C:\Users\amedrano\OneDrive - Arraya Solutions\Office365 Customers\Magnetrol Groups\Office365 Groups\o365teams_Export.csv"

foreach ($group in $Office365Groups) {
    New-UnifiedGroup -displayname $group.DestinationDisplayName -RequireSenderAuthenticationEnabled $True -name $group.DestinationDisplayName -PrimarySMTPAddress $group.DestinationEmailAddress
}

foreach ($group in $Office365Groups) {
    Write-Host "Adding Members and Owners for $($group.DestinationDisplayName) .. " -NoNewline -ForegroundColor Cyan
    $Members = $group.MembersDisplayName -split ","
    Write-Host "$($Members.count) Members found .. " -NoNewline -ForegroundColor Yellow
    foreach ($member in $Members) {
        Add-UnifiedGroupLinks -Identity $group.DestinationEmailAddress -LinkType Member -Links $member
    }
    Start-Sleep -Seconds 5
    $Owners = $group.OwnersDisplayName -split ","
    Write-Host "$($Owners.count) Owners found .. " -NoNewline -ForegroundColor Yellow
    foreach ($owner in $Owners) {
        Add-UnifiedGroupLinks -Identity $group.DestinationEmailAddress -LinkType Member -Links $owner
        Add-UnifiedGroupLinks -Identity $group.DestinationEmailAddress -LinkType Owner -Links $owner
    }
    Write-Host "Removing Aaron Medrano as Owner and Member .. " -NoNewline -ForegroundColor Yellow
    
    Remove-UnifiedGroupLinks -Identity $group.DestinationEmailAddress -LinkType Owner -Links "Aaron Medrano" -confirm:$false
    Remove-UnifiedGroupLinks -Identity $group.DestinationEmailAddress -LinkType Member -Links "Aaron Medrano" -confirm:$false
    Write-Host "done" -ForegroundColor Green
}
#add Email addresses
foreach ($group in $Office365Groups) {
    Write-Host "Adding EmailAddresses to $($group.DestinationDisplayName) .. " -NoNewline -ForegroundColor Cyan
    $EmailAddresses = $group.EmailAddresses -split ","
    Write-Host "$($EmailAddresses.count) EmailAddresses found .. " -NoNewline -ForegroundColor Yellow
    foreach ($alias in $EmailAddresses) {
        Set-UnifiedGroup -Identity $group.DestinationEmailAddress -EmailAddresses @{add=$alias} 
    }
    Write-Host "done" -ForegroundColor Green
}

# Match Magnetrolusers
function Match-AllMsolUsers {
    param (
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$true)] [array] $ImportCSV,
        [Parameter(Mandatory=$false)] [string] $NewDomain
    )
    $ImportedUsers = Import-Csv $ImportCSV
    $AllUsers = @()
    
    #ProgressBar
    $progressref = ($ImportedUsers).count
    $progresscounter = 0

    foreach ($user in $ImportedUsers)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($user.DisplayName)"
        
        Write-Host "Checking for $($user.displayName) in Tenant ..." -fore Cyan -NoNewline
        $newAddressSplit = $user.PrimarySmtpAddress -split "@"
        $newMailboxAddress = $newAddressSplit[0] + "@" + $NewDomain
        if ($msolusercheck = Get-MsolUser -UserPrincipalName $user.UserPrincipalName -ea silentlycontinue  | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation)
        {
            Write-Host "found MSOLUser*. " -ForegroundColor Green -nonewline
        }
        elseif ($msolusercheck = Get-MsolUser -UserPrincipalName $newMailboxAddress -ea silentlycontinue  | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation)
        {
            Write-Host "found MSOLUser*. " -ForegroundColor Green -nonewline
        }
        elseif ($msolusercheck = Get-MsolUser -searchstring $user.DisplayName -ea silentlycontinue | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation)
        {
            Write-Host "found MSOLUser. " -ForegroundColor Green -nonewline
        }
        else
        {
            Write-Host "not found" -ForegroundColor red -NoNewline
            $msolusercheck = @()
            $MBXStats = @()
        }
        if ($msolusercheck)
        {
            $recipientCheck = Get-Recipient $msolUserCheck.UserPrincipalName -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, IsDirSynced
            $mailboxCheck = $mailboxcheck = Get-Mailbox $msolUserCheck.UserPrincipalName -ea silentlycontinue | select IsDirSynced
            $MBXStats = Get-MailboxStatistics $mailboxcheck.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
            $user | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $True -force
            $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msolusercheck.UserPrincipalName -force
            $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msolusercheck.IsLicensed -force
            $user | add-member -type noteproperty -name "Licenses_Destination" -Value ($msolusercheck.Licenses.AccountSkuID -join ",") -force
            $user | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $mailboxcheck.IsDirSynced -force
            $user | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $msolusercheck.PreferredDataLocation -force
            $user | add-member -type noteproperty -name "Database_Destination" -Value $mailboxcheck.Database -force
            $user | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $msolusercheck.BlockCredential -force
            $user | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $recipientCheck.PrimarySmtpAddress -force
            $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipientCheck.RecipientTypeDetails -force
            $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize -force
            $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount -force

        }
        else 
        {
            $user | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False -force
            $user | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "IsLicensed_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "Licenses_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "Database_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $null -force
            $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -force 
            $user | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $null -force
            $user | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $null -force
        }
        Write-host " .. done" -foregroundcolor green
        $AllUsers += $user
    
    $allUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath
    }
}

#Create Shared Mailbox for Public Folder Contacts
$pfsharedMailboxes = import-csv
foreach ($mailbox in $pfsharedMailboxes) {
    if (!($recipientcheck = get-recipient $mailbox.PFEmailAddress -ea silentlycontinue)) {
        New-Mailbox -Shared -DisplayName $mailbox.PFDisplayName -PrimarySmtpAddress $mailbox.PFEmailAddress -name $mailbox.name
    }
    else {
        Write-Host "Recipient already exists for $($mailbox.PFEmailAddress)" -ForegroundColor Yellow
    }
}    

#Add Full Access Permissions

#Match DistributionGroups
function Match-AllDistributionGroups {
    param (
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$true)] [array] $ImportCSV,
        [Parameter(Mandatory=$false)] [string] $NewDomain
    )

    $ImportedGroups = Import-Csv $ImportCSV
    $AllGroups = @()
    #ProgressBar
    $progressref = ($ImportedGroups).count
    $progresscounter = 0

    foreach ($group in $ImportedGroups)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Distribution Group $($user.SourceDisplayName)"

        Write-Host "$($group.SourceDisplayName) .." -ForegroundColor Cyan -NoNewline
            $sourceEmailAddress = $group.SourceEmail.trim()
            if ($DistributionGroup = Get-DistributionGroup $sourceEmailAddress -ea silentlycontinue) {
            $DistributionGroupMembers = (Get-DistributionGroupMember $sourceEmailAddress -ErrorAction silentlycontinue).PrimarySMTPAddress -join ","
            $addresses = $DistributionGroup.EmailAddresses

            $group | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $True -force
            $group | add-member -type noteproperty -name "Destination_DisplayName" -Value $DistributionGroup.DisplayName       
            $group | add-member -type noteproperty -name "Destination_PrimarySmtpAddress" -Value $DistributionGroup.PrimarySmtpAddress
            $group | add-member -type noteproperty -name "Destination_EmailAddresses" -Value ($addresses -join ",")
            $group | add-member -type noteproperty -name "Destination_LegacyExchangeDN" -Value ("x500:" + $DistributionGroup.legacyexchangedn)
            $group | add-member -type noteproperty -name "Destination_AcceptMessagesOnlyFrom" -Value ($DistributionGroup.AcceptMessagesOnlyFrom -join ",")
            $group | add-member -type noteproperty -name "Destination_GrantSendOnBehalfTo" -Value ($DistributionGroup.GrantSendOnBehalfTo -join ",")
            $group | add-member -type noteproperty -name "Destination_HiddenFromAddressListsEnabled" -Value $DistributionGroup.HiddenFromAddressListsEnabled
            $group | add-member -type noteproperty -name "Destination_RejectMessagesFrom" -Value ($DistributionGroup.RejectMessagesFrom -join ",")
            $group | add-member -type noteproperty -name "Destination_RecipientTypeDetails" -Value $DistributionGroup.RecipientTypeDetails
            $group | add-member -type noteproperty -name "Destination_Alias" -Value $DistributionGroup.alias -Force
            $group | add-member -type noteproperty -name "Destination_ExchangeGuid" -Value $DistributionGroup.ExchangeGuid
            $group | add-member -type noteproperty -name "Destination_Members" -Value $DistributionGroupMembers
            Write-Host "done" -ForegroundColor Green
            $AllGroups += $group
        }
        else {
            $group | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False -force
            $group | add-member -type noteproperty -name "Destination_DisplayName" -Value $null     
            $group | add-member -type noteproperty -name "Destination_PrimarySmtpAddress" -Value $null
            $group | add-member -type noteproperty -name "Destination_EmailAddresses" -Value $null
            $group | add-member -type noteproperty -name "Destination_LegacyExchangeDN" -Value $null
            $group | add-member -type noteproperty -name "Destination_AcceptMessagesOnlyFrom" -Value $null
            $group | add-member -type noteproperty -name "Destination_GrantSendOnBehalfTo" -Value $null
            $group | add-member -type noteproperty -name "Destination_HiddenFromAddressListsEnabled" -Value $null
            $group | add-member -type noteproperty -name "Destination_RejectMessagesFrom" -Value $null
            $group | add-member -type noteproperty -name "Destination_RecipientTypeDetails" -Value $null
            $group | add-member -type noteproperty -name "Destination_Alias" -Value $null
            $group | add-member -type noteproperty -name "Destination_ExchangeGuid" -Value $null
            $group | add-member -type noteproperty -name "Destination_Members" -Value $null
            Write-Host "done" -ForegroundColor Green
            $AllGroups += $group
        }    
    $AllGroups | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath
    }
}

#Clear Access Permissions
$resources = @()
foreach ($user in $fullaccesspermsusers) {
$recipientcheck = get-recipient $user.Mailbox
if ($recipientcheck | ?{$_.recipienttypedetails -eq "UserMailbox"})
    {
        Write-Host "Storing Mailbox $($user.mailbox)"
        $resources += $user.mailbox
    }
}

foreach ($user in $removePermUsers) {
    $perms = Get-MailboxPermission $user | ?{$_.user -ne "NT AUTHORITY\SELF"}
    foreach ($perm in $perms) {
        Remove-MailboxPermission -Identity $user -User $perm.user -AccessRights $perm.AccessRights -Confirm:$false
        Write-Host "Removed $($perm.user) $($perm.AccessRights) permission from $($user)"
    }
}

#Convert Mail Users to RemoteMailboxes
$progressref = ($MagnetrolUsers).count
$progresscounter = 0

foreach ($user in $MagnetrolUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($user.DestinationEmail_OG)"

    Write-Host "Updating $($user.DestinationEmail_OG) to RemoteMailbox .. " -ForegroundColor Cyan -NoNewline
    
    if ($recipientcheck = Get-Recipient $user.DestinationEmail_OG | ?{$_.RecipientTypeDetails -eq "MailUser"}) {
        Enable-RemoteMailbox $recipientCheck.DistinguishedName
    }
    Write-Host "done " -ForegroundColor Green
}


### Cutover START REGION ####

# Force Users to be logged out:
$MagnetrolUsers = Import-Csv "~\Customer_Provided Magnetrol_Mailboxes-UPDATED.csv"

$progressref = ($MagnetrolUsers).count
$progresscounter = 0
foreach ($user in $MagnetrolUsers){
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    $UserPrincipalName = $user.SourceEmail
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Revoking Session for $($UserPrincipalName)"
        
    Write-Host "Force Logged Off user $($user.DisplayName)" -ForegroundColor Green
    Revoke-SPOUserSession -user $UserPrincipalName -Confirm:$false
    Get-AzureADUser -SearchString $UserPrincipalName | Revoke-AzureADUserAllRefreshToken
}

#Update UPN to OnMicrosoft
$msolUPNUsers = Get-MSOLUSER -all | ?{$_.UserPrincipalName -notlike "*.onmicrosoft.com"}
$progressref = ($msolUPNUsers).count

$progresscounter = 0
foreach ($user in $msolUPNUsers){
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    #Set Variables
    $UserPrincipalName = $user.UserPrincipalName
    $UPNSPlit = $user.UserPrincipalName -split "@"
    $NewUPN = $UPNSplit[0] + "@magnetrol.onmicrosoft.com"

    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating UPN for $($UserPrincipalName)"
    Set-MsolUserPrincipalName  -UserPrincipalName $UserPrincipalName -NewUserPrincipalName $newUPN
    Write-Host "UPN Updated for $($user.DisplayName)" -ForegroundColor Green
}



#Add Alias Address to On-PremUsers - TEST DIANA
$MagnetrolUsers = Import-Csv

foreach ($user in ($MagnetrolUsers | ? {$_.DIsplayName -eq "Diana Trinidad"})) {
    $EmailAddressesSplit = $user.EmailAddresses -split ","
    $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*smtp:*" -or $_ -like "*X500:*" -and $_ -notlike "*.onmicrosoft.com*"}
    Write-Host "Adding $($EmailAddresses.count) aliases to $($user.DestinationEmail_OG) .. " -ForegroundColor Cyan -NoNewline
    foreach ($alias in $EmailAddresses) {
        $simpleAlias = $alias.replace("smtp:","")
        Set-RemoteMailbox -Identity $user.DestinationEmail_OG -EmailAddresses @{add=$simpleAlias}
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    Write-Host "done " -ForegroundColor Green
}

#Add Alias Address to On-PremUsers (REMOTEMAILBOXES)
$MagnetrolUsers = Import-Csv

$progressref = ($MagnetrolUsers).count
$progresscounter = 0

foreach ($user in $MagnetrolUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($user.DestinationEmail_OG)"

    $EmailAddressesSplit = $user.EmailAddresses -split ","
    $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*smtp:*" -or $_ -like "*X500:*" -and $_ -notlike "*.onmicrosoft.com*"}
    Write-Host "Adding $($EmailAddresses.count) aliases to $($user.DestinationEmail_OG) .. " -ForegroundColor Cyan -NoNewline
    foreach ($alias in $EmailAddresses) {
        $simpleAlias = $alias.replace("smtp:","")
        Set-RemoteMailbox -Identity $user.DestinationEmail_OG -EmailAddresses @{add=$simpleAlias}
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    $X500 = $user.LegacyExchangeDN
    Set-RemoteMailbox -Identity $user.DestinationEmail_OG -EmailAddresses @{add=$X500}
    Write-Host "done " -ForegroundColor Green
}

#Add Alias Address to On-PremUsers (MAILUSERS and REMOTESHAREDMAILBOXES)
$MagnetrolUsers = Import-Csv

$progressref = ($MagnetrolUsers).count
$progresscounter = 0

foreach ($user in $MagnetrolUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($user.DestinationEmail_OG)"

    $EmailAddressesSplit = $user.EmailAddresses -split ","
    $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*smtp:*" -or $_ -like "*X500:*" -and $_ -notlike "*.onmicrosoft.com*"}
    Write-Host "Adding $($EmailAddresses.count) aliases to $($user.DestinationEmail_OG) .. " -ForegroundColor Cyan -NoNewline
    
    if ($recipientcheck = Get-Recipient $user.DestinationEmail_OG | ?{$_.RecipientTypeDetails -eq "MailUser"}) {
        foreach ($alias in $EmailAddresses) {
            $simpleAlias = $alias.replace("smtp:","")
            Set-MailUser -Identity $user.DestinationEmail_OG -EmailAddresses @{add=$simpleAlias}
            Write-Host ". " -ForegroundColor Green -NoNewline
        }
        $X500 = $user.LegacyExchangeDN
        Set-MailUser -Identity $user.DestinationEmail_OG -EmailAddresses @{add=$X500}
    }
    elseif ($recipientcheck = Get-Recipient $user.DestinationEmail_OG | ?{$_.RecipientTypeDetails -eq "RemoteSharedMailbox"}) {
        foreach ($alias in $EmailAddresses) {
            $simpleAlias = $alias.replace("smtp:","")
            Set-RemoteMailbox -Identity $user.DestinationEmail_OG -EmailAddresses @{add=$simpleAlias}
            Write-Host ". " -ForegroundColor Green -NoNewline
        }
        $X500 = $user.LegacyExchangeDN
        Set-RemoteMailbox -Identity $user.DestinationEmail_OG -EmailAddresses @{add=$X500}
    }
    Write-Host "done " -ForegroundColor Green
}

#Add Alias Address to On-PremGroups
$MagnetrolGroups = Import-Csv

$progressref = ($MagnetrolGroups).count
$progresscounter = 0
foreach ($group in $MagnetrolGroups) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($group.DestinationEmail_OG)"

    $EmailAddressesSplit = $group.EmailAddresses -split ","
    $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*smtp:*" -or $_ -like "*X500:*" -and $_ -notlike "*.onmicrosoft.com*"}
    Write-Host "Adding $($EmailAddresses.count) aliases to $($group.DestinationEmail_OG) .. " -ForegroundColor Cyan -NoNewline
    foreach ($alias in $EmailAddresses) {
        $simpleAlias = $alias.replace("smtp:","")
        Set-DistributionGroup -Identity $group.DestinationEmail_OG -EmailAddresses @{add=$simpleAlias}
        Write-Host ". " -ForegroundColor Green -NoNewline
    }
    $X500 = $group.LegacyExchangeDN
    Set-DistributionGroup -Identity $group.DestinationEmail_OG -EmailAddresses @{add=$X500}
    Write-Host "done " -ForegroundColor Green
}

#Remove All Alias addresses from Recipients
#Clear All domains from source tenant

$allRecipients = import-csv

#ProgressBar
$progressref = ($allRecipients).count
$progresscounter = 0
foreach ($recipient in $allRecipients) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($recipient.DisplayName)"

    $oldPrimarySMTPAddress = $recipient.PrimarySMTPAddress
    if ($recipient.RecipientTypeDetails -like "*Mailbox*") {
        $EmailAddressesSplit = $recipient.EmailAddresses -split ","
        $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*@magnetrol.*" -or $_ -like "*@orioninstruments.com" -or $_ -like "*@introtek.com" -or $_ -like "*@innovativesensing.com" -and $_ -notlike "*@magnetrol.onmicrosoft.com"}
        #Get NewPrimarySMTPAddress
        $recipientCheck = Get-Recipient $oldPrimarySMTPAddress | select -ExpandProperty EmailAddresses
        $newPrimarySMTPAddress = $recipientCheck | ?{$_ -like "*@magnetrol.onmicrosoft.com"}
        $newPrimarySMTPAddressSplit = $newPrimarySMTPAddress -split ":"
        Write-Host "Setting PrimarySMTPAddress and Removing $($EmailAddresses.count) aliases from $($oldPrimarySMTPAddress) ... " -foregroundcolor green -NoNewline
        #Update PrimarySMTPAddress
        Set-Mailbox -Identity $oldPrimarySMTPAddress -WindowsEmailAddress $newPrimarySMTPAddressSplit[1]
        #Remove Aliases
        foreach ($alias in $EmailAddresses) {
            Set-Mailbox -Identity $newPrimarySMTPAddressSplit[1] -EmailAddresses @{remove=$alias}
            Write-Host ". " -ForegroundColor Green  -NoNewline
        }
    }
    elseif ($recipient.RecipientTypeDetails -like "*DistributionGroup*") {
        $EmailAddressesSplit = $recipient.EmailAddresses -split ","
        $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*@magnetrol.*" -or $_ -like "*@orioninstruments.com" -or $_ -like "*@introtek.com" -or $_ -like "*@innovativesensing.com" -and $_ -notlike "*@magnetrol.onmicrosoft.com"}
        #Get NewPrimarySMTPAddress
        $recipientCheck = Get-Recipient $oldPrimarySMTPAddress | select -ExpandProperty EmailAddresses
        $newPrimarySMTPAddress = $recipientCheck | ?{$_ -like "*@magnetrol.onmicrosoft.com"}
        $newPrimarySMTPAddressSplit = $newPrimarySMTPAddress -split ":"
        Write-Host "Setting PrimarySMTPAddress and Removing $($EmailAddresses.count) aliases from $($oldPrimarySMTPAddress) ... " -foregroundcolor DarkCyan -NoNewline
        #Update PrimarySMTPAddress
        Set-DistributionGroup -Identity $oldPrimarySMTPAddress -PrimarySMTPAddress $newPrimarySMTPAddressSplit[1]
        #Remove Aliases
        Write-Host "Removing $($EmailAddresses.count) aliases from $($oldPrimarySMTPAddress) ... "
        foreach ($alias in $EmailAddresses) {
            Set-DistributionGroup -Identity $newPrimarySMTPAddressSplit[1] -EmailAddresses @{remove=$alias}
            Write-Host ". " -ForegroundColor Green  -NoNewline
        }
    }
    elseif ($recipient.RecipientTypeDetails -like "*MailContact*") {
        $EmailAddressesSplit = $recipient.EmailAddresses -split ","
        $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*@magnetrol.*" -or $_ -like "*@orioninstruments.com" -or $_ -like "*@introtek.com" -or $_ -like "*@innovativesensing.com" -and $_ -notlike "*@magnetrol.onmicrosoft.com"}
        #Get NewPrimarySMTPAddress
        $recipientCheck = Get-Recipient $oldPrimarySMTPAddress | select -ExpandProperty EmailAddresses
        $newPrimarySMTPAddress = $recipientCheck | ?{$_ -like "*@magnetrol.onmicrosoft.com"}
        $newPrimarySMTPAddressSplit = $newPrimarySMTPAddress -split ":"
        Write-Host "Setting PrimarySMTPAddress and Removing $($EmailAddresses.count) aliases from $($oldPrimarySMTPAddress) ... " -foregroundcolor yellow -NoNewline
        #Update PrimarySMTPAddress
        Set-MailContact -Identity $oldPrimarySMTPAddress -WindowsEmailAddress $newPrimarySMTPAddressSplit[1]
        #Remove Aliases
        foreach ($alias in $EmailAddresses) {
            Set-MailContact -Identity $newPrimarySMTPAddressSplit[1] -EmailAddresses @{remove=$alias}
            Write-Host ". " -ForegroundColor Green  -NoNewline
        }
    }
    elseif ($recipient.RecipientTypeDetails -like "*MailUser*") {
        $EmailAddressesSplit = $recipient.EmailAddresses -split ","
        $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*@magnetrol.*" -or $_ -like "*@orioninstruments.com" -or $_ -like "*@introtek.com" -or $_ -like "*@innovativesensing.com" -and $_ -notlike "*@magnetrol.onmicrosoft.com"}
        #Get NewPrimarySMTPAddress
        $recipientCheck = Get-Recipient $oldPrimarySMTPAddress | select -ExpandProperty EmailAddresses
        $newPrimarySMTPAddress = $recipientCheck | ?{$_ -like "*@magnetrol.onmicrosoft.com"}
        $newPrimarySMTPAddressSplit = $newPrimarySMTPAddress -split ":"
        Write-Host "Setting PrimarySMTPAddress and Removing $($EmailAddresses.count) aliases from $($oldPrimarySMTPAddress) ... " -foregroundcolor yellow -NoNewline
        #Update PrimarySMTPAddress
        Set-Mailuser -Identity $oldPrimarySMTPAddress -PrimarySMTPAddress $newPrimarySMTPAddressSplit[1]
        foreach ($alias in $EmailAddresses) {
            Set-MailUser -Identity $newPrimarySMTPAddressSplit[1] -EmailAddresses @{remove=$alias}
            Write-Host ". " -ForegroundColor Green  -NoNewline
        }
    }
    elseif ($recipient.RecipientTypeDetails -like "*PublicFolder*") {
        $EmailAddressesSplit = $recipient.EmailAddresses -split ","
        $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*@magnetrol.*" -or $_ -like "*@orioninstruments.com" -or $_ -like "*@introtek.com" -or $_ -like "*@innovativesensing.com" -and $_ -notlike "*@magnetrol.onmicrosoft.com"}
        #Get NewPrimarySMTPAddress
        $recipientCheck = Get-Recipient $oldPrimarySMTPAddress | select -ExpandProperty EmailAddresses
        $newPrimarySMTPAddress = $recipientCheck | ?{$_ -like "*@magnetrol.onmicrosoft.com"}
        $newPrimarySMTPAddressSplit = $newPrimarySMTPAddress -split ":"
        Write-Host "Setting PrimarySMTPAddress and Removing $($EmailAddresses.count) aliases from $($oldPrimarySMTPAddress) ... " -foregroundcolor yellow -NoNewline
        #Update PrimarySMTPAddress
        Set-MailPublicFolder -Identity $oldPrimarySMTPAddress -PrimarySMTPAddress $newPrimarySMTPAddressSplit[1]
        foreach ($alias in $EmailAddresses) {
            Set-MailPublicFolder -Identity $newPrimarySMTPAddressSplit[1] -EmailAddresses @{remove=$alias}
            Write-Host ". " -ForegroundColor Green  -NoNewline
        }
    }
    Write-Host "done" -foregroundcolor green
}

#Remove Forwarding on Ametek
$customerMatchedUsers = Import-Csv "C:\Users\amedrano\OneDrive - Arraya Solutions\Office365 Customers\Magnetrol Groups\Customer_Provided Magnetrol_Mailboxes.csv"

$progressref = ($customerMatchedUsers).count
$progresscounter = 0

$updatedUsers = @()
$notUpdatedUsers = @()
foreach ($user in $customerMatchedUsers)
{
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($user.Destination_UserPrincipalName_OG)"

    if ($mailboxcheck = Get-Mailbox $user.Destination_UserPrincipalName_OG -ea silentlycontinue)
    {
        Set-Mailbox $mailboxcheck.PrimarySmtpAddress -ForwardingSmtpAddress $null
        Write-Host "Removed Forwarding from $($user.Destination_UserPrincipalName_OG)" -ForegroundColor Green
        $updatedUsers += $user
    }
}

### Cutover END OF REGION

## Jim's Voice stuff

$UserList = Import-Csv "C:\Users\JBaker\OneDrive - Arraya Solutions\Projects\Ametek\Magnetrol\Userphonelist.csv"

$notfoundusers = @()
ForEach ($item in $UserList){
    $Identity = $($item.identity)
    $LineURI = $($item.lineuri)

    ## Check CS User Exists
        if ($csuserCheck = Get-CSUser -identity $Identity -ea silentlycontinue) {
            Write-Host "User $($Identity) Found and updating .. " -foregroundcolor cyan -nonewline
            Set-CsUser -Identity $Identity -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI $LineURI
            Write-Host "done" -foregroundcolor green
        }
        else {
            Write-Host "User $($Identity) Not Found" -foregroundcolor red
            $notfoundusers += $item
        }
}

## Shared PF Mailbox
$pfsharedMailboxes
foreach ($mailbox in $pfsharedMailboxes) {
    $PermUsers = $mailbox.PermUsers -split ","
    foreach ($user in $PermUsers) {
        Add-MailboxPermission -identity $mailbox.PFEmailAddress -User $user -AccessRights FullAccess -Automapping $false
    }
}

#CloudM Migration Report Check
foreach ($user in $cloudMReport) {
    $userID = $user.UserID.trim()
    $matchedUser = $allMatchedUsers | ? {$_.SourceEmail -like "*$userID*"}
    $user | add-member -type noteproperty -name "SourceEmail" -Value $matchedUser.SourceEmail -Force
    $user | add-member -type noteproperty -name "DestinationEmail" -Value $matchedUser.DestinationEmail -Force
}
$cloudMReport | export-csv -NoTypeInformation -Encoding utf8 

#Update Resource EmailAddress
foreach ($user in $sharedMailboxes) {
    $recipientCheck = Get-Recipient $user.DisplayName -ea silentlycontinue
    $user | add-member -type noteproperty -name "SourceOnMicrosoftEmail" -Value $recipientCheck.PrimarySMTPAddress -Force
}

# Add Email Addresses to SharedMailboxList
$allRecipients = import-csv "C:\Users\amedrano\OneDrive - Arraya Solutions\Office365 Customers\Magnetrol Groups\Exchange Online\OwnSourceExports\Magnetrol_AllRecipients.csv"
$sharedMailboxes = Import-Csv "C:\Users\amedrano\OneDrive - Arraya Solutions\Office365 Customers\Magnetrol Groups\Customer Source Exports\SharedMailboxes.csv"

foreach ($mailbox in $sharedMailboxes) {
    $SourceEmail = $mailbox.SourceEmail.trim()
    Write-Host "Gathering Details for $($SourceEmail)" -ForegroundColor Cyan -NoNewline
    if ($matchedUser = ($allRecipients | ?{$_.PrimarySmtpAddress -eq $SourceEmail})) {
        $mailbox | add-member -type noteproperty -name "EmailAddresses" -Value $matchedUser.EmailAddresses
        Write-Host "." -ForegroundColor Green -NoNewline
    }
    else {
        Write-Host "." -ForegroundColor red -NoNewline
    }
}
Write-Host "done"

#Add Alias Address to On-PremUsers (REMOTESHAREDMAILBOXES)
$MagnetrolUsers = Import-Csv

$progressref = ($sharedMailboxes).count
$progresscounter = 0

foreach ($user in $sharedMailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($user.DisplayName)"

    $EmailAddressesSplit = $user.EmailAddresses -split ","
    $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*smtp:*" -or $_ -like "*X500:*" -and $_ -notlike "*.onmicrosoft.com*"}
    Write-Host "Adding $($EmailAddresses.count) aliases to $($user.DestinationEmail) .. " -ForegroundColor Cyan -NoNewline
    
    if ($recipientcheck = Get-Recipient $user.DestinationEmail | ?{$_.RecipientTypeDetails -eq "RemoteSharedMailbox"}) {
        foreach ($alias in $EmailAddresses) {
            $simpleAlias = $alias.replace("smtp:","")
            Set-RemoteMailbox -Identity $user.DestinationEmail -EmailAddresses @{add=$simpleAlias}
            Write-Host ". " -ForegroundColor Green -NoNewline
        }
    }
    Write-Host "done " -ForegroundColor Green
}

#Get EmailAddresses
$groups = Import-Csv
$allRecipients = import-csv
foreach ($group in $groups) {
    $sourceEmailAddress = $group.SourceEmail
    if ($matchedgroup = ($allRecipients | ?{$_.PrimarySMTPAddress -eq $sourceEmailAddress})) {
        Write-Host "Checking Group $($group.SourceDisplayName)"
        $group | add-member -type noteproperty -name "EmailAddresses" -Value $matchedgroup.EmailAddresses -Force
    }
}

#Add Group EmailAddresses
$magnetrolGroups = Import-Csv

$progressref = ($magnetrolGroups).count
$progresscounter = 0

foreach ($group in $magnetrolGroups) {
    Write-Host "Updating Distribution Group $($group.DestinationName) .. " -ForegroundColor Cyan -NoNewline
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($group.DestinationName)"

    $EmailAddressesSplit = $group.EmailAddresses -split ","
    $EmailAddresses = $EmailAddressesSplit | ? {$_ -like "*smtp:*" -or $_ -like "*X500:*" -and $_ -notlike "*.onmicrosoft.com*"}
    Write-Host "Adding $($EmailAddresses.count) aliases to $($group.DestinationEmail) .. " -ForegroundColor Cyan -NoNewline
    
    if ($recipientcheck = Get-Recipient $group.DestinationEmail | ?{$_.RecipientTypeDetails -eq "RemoteSharedMailbox"}) {
        foreach ($alias in $EmailAddresses) {
            $simpleAlias = $alias.replace("smtp:","")
            Set-RemoteMailbox -Identity $user.DestinationEmail -EmailAddresses @{add=$simpleAlias}
            Write-Host ". " -ForegroundColor Green -NoNewline
        }
    }
    Write-Host "done" -ForegroundColor Green
}

#Match Users
$users = Import-Csv
foreach ($user in $users) {
    $displayName = $user.DisplayName
    if ($msolUserCheck = get-msoluser -SearchString $displayName -ErrorAction SilentlyContinue) {
        $user | add-member -type noteproperty -name "OnMicrosoftAddress" -Value $msolUserCheck.UserPrincipalName -Force
    }
    else {
        $user | add-member -type noteproperty -name "OnMicrosoftAddress" -Value $null -Force
    }
}
# Match Voice users 
$progressref = ($magnetrolUsers).count
$progresscounter = 0
foreach ($voice in $magnetrolUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($voice.DisplayName)"
    $displayName = $voice.DisplayName
    if ($voiceUserCheck = ($ametekVoiceUsers | ?{$_.DisplayName -eq $displayName})) {
        $voice | add-member -type noteproperty -name "Ametek_UPN" -Value $voiceUserCheck.Ametek_UserPrincipalName -Force
        $voice | add-member -type noteproperty -name "Ametek_VoiceRoutingPolicy" -Value $voiceUserCheck.AmetekOnlineVoiceRoutingPolicy -Force
    }
    else {
        $voice | add-member -type noteproperty -name "Ametek_UPN" -Value $null -Force
        $voice | add-member -type noteproperty -name "Ametek_VoiceRoutingPolicy" -Value $null -Force
    }
}

#Check Current Voice Policy
# Match Voice users 
$matchedVoiceUsers = $combinedVoiceUsers | ?{$_.Ametek_UPN}
$progressref = ($matchedVoiceUsers).count
$progresscounter = 0
foreach ($voiceUser in $matchedVoiceUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($voiceUser.DisplayName)"
    $displayName = $voice.DisplayName
    if ($voiceUserCheck = (Get-CsOnlineUser -Identity $voiceUser.Ametek_UPN -ea silentlycontinue)) {
        $voiceUser | add-member -type noteproperty -name "Recent_Ametek_VoiceRoutingPolicy" -Value $voiceUserCheck.OnlineVoiceRoutingPolicy -Force
    }
    else {
        $voiceUser | add-member -type noteproperty -name "Recent_Ametek_VoiceRoutingPolicy" -Value "Not Found" -Force
    }
}

#Get License

$progressref = ($matchedVoiceUsers).count
$progresscounter = 0
foreach ($voiceUser in $matchedVoiceUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($voiceUser.DisplayName)"
    $displayName = $voice.DisplayName
    if ($msolUserCheck = (Get-MsolUser -UserPrincipalName $voiceUser.Ametek_UPN -ea silentlycontinue)) {
        $voiceUser | add-member -type noteproperty -name "Licenses" -Value ($msolusercheck.Licenses.AccountSkuID -join ",")-Force
    }
    else {
        $voiceUser | add-member -type noteproperty -name "Licenses" -Value $null -Force
    }
}

#Grant VoiceRouting Policy
$usersToUpdate = $matchedVoiceUsers | ?{$_.MagnetrolOnlineVoiceRoutingPolicy -and $_.Recent_Ametek_VoiceRoutingPolicy -eq $null}
$progressref = ($usersToUpdate).count
$progresscounter = 0
foreach ($voiceUser in $usersToUpdate) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Doing things for $($voiceUser.DisplayName)"
    Grant-CsOnlineVoiceRoutingPolicy -Identity $voiceUser.Ametek_UPN -PolicyName $voiceUser.MagnetrolOnlineVoiceRoutingPolicy
}

#Update All Recipients with current OnMicrosoft address
$magnetrolAllRecipients = Import-Csv "C:\Users\amedrano\OneDrive - Arraya Solutions\Office365 Customers\Magnetrol Groups\Exchange Online\OwnSourceExports\Magnetrol_AllRecipients.csv"
foreach ($recipient in $magnetrolAllRecipients) {
    if ($recipientCheck = (Get-Recipient $recipient.PrimarySmtpAddress -ea silentlycontinue)) {
        $recipient | add-member -type noteproperty -name "Ametek_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress -Force
    }
    else {
        $recipient | add-member -type noteproperty -name "Ametek_PrimarySMTPAddress" -Value $null -Force
    }
}