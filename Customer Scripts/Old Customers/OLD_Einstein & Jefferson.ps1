# Create Dynamic Distribution Groups
## Resources
New-DynamicDistributionGroup -Name "SharedMailbox Dynamic Group - GALSYNC" -RecipientFilter "(RecipientTypeDetails -eq 'SharedMailbox')"

## Create Contact Dynamic Distribution Group
### Set MailContact - Suggest CustomAttribute5 to "GALSync"
$JeffersonMailContacts = Get-MailContact -filter {ExternalEmailAddress -like "*@jefferson.edu"} -ResultSize unlimited
#ProgressBar
$progressref = ($JeffersonMailContacts).count
$progresscounter = 0
 foreach ($contact in $JeffersonMailContacts) {
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Custom Attribute for $($contact.DisplayName)"
    Set-MailContact $contact.PrimarySmtpAddress -customAttribute5 "GALSync"
 }

$EinsteinMailContacts = Get-MailContact -filter {ExternalEmailAddress -like "*@einstein.edu"} -ResultSize unlimited
$progressref = ($EinsteinMailContacts).count
$progresscounter = 0
 foreach ($contact in $EinsteinMailContacts) {
    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Custom Attribute for $($contact.DisplayName)"
    Set-MailContact $contact.PrimarySmtpAddress -customAttribute5 "GALSync" -wa silentlycontinue
 }

## Create Dynamic Contact Groups
New-DynamicDistributionGroup -Name "Jefferson GALSync Group" -RecipientFilter "(CustomAttribute5 -eq 'GALSync')"

New-DynamicDistributionGroup -Name "Einstein GALSync Group" -RecipientFilter "(CustomAttribute5 -eq 'GALSync')"
 
# Get Resource Room Details
$resourceMailboxes = Get-Mailbox -ResultSize unlimited -Filter "RecipientTypeDetails -eq 'RoomMailbox' -or RecipientTypeDetails -eq 'EquipmentMailbox'"
$AllResourceDetails = @()
foreach ($room in $resourceMailboxes) {
    $CalendarProcessingDetails = Get-CalendarProcessing -identity $room.PrimarySmtpAddress
    $EmailAddresses = $room  | select -ExpandProperty EmailAddresses

    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $room.DisplayName
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $room.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $room.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "Name" -Value $room.Name
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ",")
    $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $room.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $room.ForwardingSmtpAddress
    $currentuser | add-member -type noteproperty -name "IsResource" -Value $room.IsResource
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $room.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ResourceCapacity" -Value $room.ResourceCapacity
    $currentuser | add-member -type noteproperty -name "ResourceType" -Value $room.ResourceType
    $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $room.UserPrincipalName
    $currentuser | add-member -type noteproperty -name "IsDirSynced" -Value $room.IsDirSynced
    $currentuser | add-member -type noteproperty -name "CustomAttribute9" -Value $room.CustomAttribute9
    $currentuser | add-member -type noteproperty -name "CustomAttribute10" -Value $room.CustomAttribute10
    $currentuser | add-member -type noteproperty -name "DistinguishedName" -Value $room.DistinguishedName
    $currentuser | add-member -type noteproperty -name "ExchangeObjectId" -Value $room.ExchangeObjectId
    $currentuser | add-member -type noteproperty -name "Guid" -Value $room.Guid

    #CalendarProcessing Details
    $currentuser | add-member -type noteproperty -name "AutomateProcessing" -Value $CalendarProcessingDetails.AutomateProcessing
    $currentuser | add-member -type noteproperty -name "AllowConflicts" -Value $CalendarProcessingDetails.AllowConflicts
    $currentuser | add-member -type noteproperty -name "AllowDistributionGroup" -Value $CalendarProcessingDetails.AllowDistributionGroup
    $currentuser | add-member -type noteproperty -name "AllowMultipleResources" -Value $CalendarProcessingDetails.AllowMultipleResources
    $currentuser | add-member -type noteproperty -name "BookingType" -Value $CalendarProcessingDetails.BookingType
    $currentuser | add-member -type noteproperty -name "AllowRecurringMeetings" -Value $CalendarProcessingDetails.AllowRecurringMeetings
    $currentuser | add-member -type noteproperty -name "EnforceCapacity" -Value $CalendarProcessingDetails.EnforceCapacity
    $currentuser | add-member -type noteproperty -name "ConflictPercentageAllowed" -Value $CalendarProcessingDetails.ConflictPercentageAllowed
    $currentuser | add-member -type noteproperty -name "MaximumConflictInstances" -Value $CalendarProcessingDetails.MaximumConflictInstances
    $currentuser | add-member -type noteproperty -name "ForwardRequestsToDelegates" -Value $CalendarProcessingDetails.ForwardRequestsToDelegates
    $currentuser | add-member -type noteproperty -name "DeleteAttachments" -Value $CalendarProcessingDetails.DeleteAttachments
    $currentuser | add-member -type noteproperty -name "DeleteComments" -Value $CalendarProcessingDetails.DeleteComments
    $currentuser | add-member -type noteproperty -name "RemovePrivateProperty" -Value $CalendarProcessingDetails.RemovePrivateProperty
    $currentuser | add-member -type noteproperty -name "DeleteSubject" -Value $CalendarProcessingDetails.DeleteSubject
    $currentuser | add-member -type noteproperty -name "OrganizerInfo" -Value $CalendarProcessingDetails.OrganizerInfo
    $currentuser | add-member -type noteproperty -name "TentativePendingApproval" -Value $CalendarProcessingDetails.TentativePendingApproval
    $currentuser | add-member -type noteproperty -name "ProcessExternalMeetingMessages" -Value $CalendarProcessingDetails.ProcessExternalMeetingMessages
    $currentuser | add-member -type noteproperty -name "AddNewRequestsTentatively" -Value $CalendarProcessingDetails.AddNewRequestsTentatively
    $currentuser | add-member -type noteproperty -name "RemoveOldMeetingMessages" -Value $CalendarProcessingDetails.RemoveOldMeetingMessages
    $currentuser | add-member -type noteproperty -name "AllRequestInPolicy" -Value $CalendarProcessingDetails.AllRequestInPolicy
    $currentuser | add-member -type noteproperty -name "AllBookInPolicy" -Value $CalendarProcessingDetails.AllBookInPolicy
    $currentuser | add-member -type noteproperty -name "AllRequestOutOfPolicy" -Value $CalendarProcessingDetails.AllRequestOutOfPolicy

    $AllResourceDetails += $currentuser
}

$AllResourceDetails = @()
$progressref = ($notFoundResources).count
$progresscounter = 0
foreach ($mailbox in $notFoundResources) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Resource Details for $($mailbox.DisplayName)"

    $CalendarProcessingDetails = Get-CalendarProcessing -identity $mailbox.PrimarySmtpAddress
    $room = get-mailbox $mailbox.PrimarySmtpAddress
    $EmailAddresses = $room  | select -ExpandProperty EmailAddresses

    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $room.DisplayName
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $room.PrimarySmtpAddress
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $room.RecipientTypeDetails
    $currentuser | add-member -type noteproperty -name "Name" -Value $room.Name
    $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($EmailAddresses -join ",")
    $currentuser | add-member -type noteproperty -name "ForwardingAddress" -Value $room.ForwardingAddress
    $currentuser | add-member -type noteproperty -name "ForwardingSmtpAddress" -Value $room.ForwardingSmtpAddress
    $currentuser | add-member -type noteproperty -name "IsResource" -Value $room.IsResource
    $currentuser | add-member -type noteproperty -name "DeliverToMailboxAndForward" -Value $room.DeliverToMailboxAndForward
    $currentuser | add-member -type noteproperty -name "ResourceCapacity" -Value $room.ResourceCapacity
    $currentuser | add-member -type noteproperty -name "ResourceType" -Value $room.ResourceType
    $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $room.UserPrincipalName
    $currentuser | add-member -type noteproperty -name "IsDirSynced" -Value $room.IsDirSynced
    $currentuser | add-member -type noteproperty -name "CustomAttribute9" -Value $room.CustomAttribute9
    $currentuser | add-member -type noteproperty -name "CustomAttribute10" -Value $room.CustomAttribute10
    $currentuser | add-member -type noteproperty -name "DistinguishedName" -Value $room.DistinguishedName
    $currentuser | add-member -type noteproperty -name "ExchangeObjectId" -Value $room.ExchangeObjectId
    $currentuser | add-member -type noteproperty -name "Guid" -Value $room.Guid

    #CalendarProcessing Details
    $currentuser | add-member -type noteproperty -name "AutomateProcessing" -Value $CalendarProcessingDetails.AutomateProcessing
    $currentuser | add-member -type noteproperty -name "AllowConflicts" -Value $CalendarProcessingDetails.AllowConflicts
    $currentuser | add-member -type noteproperty -name "AllowDistributionGroup" -Value $CalendarProcessingDetails.AllowDistributionGroup
    $currentuser | add-member -type noteproperty -name "AllowMultipleResources" -Value $CalendarProcessingDetails.AllowMultipleResources
    $currentuser | add-member -type noteproperty -name "BookingType" -Value $CalendarProcessingDetails.BookingType
    $currentuser | add-member -type noteproperty -name "AllowRecurringMeetings" -Value $CalendarProcessingDetails.AllowRecurringMeetings
    $currentuser | add-member -type noteproperty -name "EnforceCapacity" -Value $CalendarProcessingDetails.EnforceCapacity
    $currentuser | add-member -type noteproperty -name "ConflictPercentageAllowed" -Value $CalendarProcessingDetails.ConflictPercentageAllowed
    $currentuser | add-member -type noteproperty -name "MaximumConflictInstances" -Value $CalendarProcessingDetails.MaximumConflictInstances
    $currentuser | add-member -type noteproperty -name "ForwardRequestsToDelegates" -Value $CalendarProcessingDetails.ForwardRequestsToDelegates
    $currentuser | add-member -type noteproperty -name "DeleteAttachments" -Value $CalendarProcessingDetails.DeleteAttachments
    $currentuser | add-member -type noteproperty -name "DeleteComments" -Value $CalendarProcessingDetails.DeleteComments
    $currentuser | add-member -type noteproperty -name "RemovePrivateProperty" -Value $CalendarProcessingDetails.RemovePrivateProperty
    $currentuser | add-member -type noteproperty -name "DeleteSubject" -Value $CalendarProcessingDetails.DeleteSubject
    $currentuser | add-member -type noteproperty -name "OrganizerInfo" -Value $CalendarProcessingDetails.OrganizerInfo
    $currentuser | add-member -type noteproperty -name "TentativePendingApproval" -Value $CalendarProcessingDetails.TentativePendingApproval
    $currentuser | add-member -type noteproperty -name "ProcessExternalMeetingMessages" -Value $CalendarProcessingDetails.ProcessExternalMeetingMessages
    $currentuser | add-member -type noteproperty -name "AddNewRequestsTentatively" -Value $CalendarProcessingDetails.AddNewRequestsTentatively
    $currentuser | add-member -type noteproperty -name "RemoveOldMeetingMessages" -Value $CalendarProcessingDetails.RemoveOldMeetingMessages
    $currentuser | add-member -type noteproperty -name "AllRequestInPolicy" -Value $CalendarProcessingDetails.AllRequestInPolicy
    $currentuser | add-member -type noteproperty -name "AllBookInPolicy" -Value $CalendarProcessingDetails.AllBookInPolicy
    $currentuser | add-member -type noteproperty -name "AllRequestOutOfPolicy" -Value $CalendarProcessingDetails.AllRequestOutOfPolicy

    $AllResourceDetails += $currentuser
}

$AllResourceDetails | Export-Csv -NoTypeInformation -Encoding UTF8


$resourceMailboxes | foreach {
Write-Host "Updating CustomAttribute6 for Resource $($_.name) to Migrating-GalSync" -ForegroundColor green
Set-Mailbox $_.primarysmtpaddress -CustomAttribute6 "Migrating-GalSync"
}


$resourceMailboxes = Import-Csv 

$notFoundResources=@()
foreach ($resource in $resourceMailboxes) {
    Write-Host "Updating CustomAttribute6 for Resource $($resource.DisplayName) to Migrating-GalSync .. " -ForegroundColor cyan -NoNewlin
    $emailAddress = $resource.PrimarySMTPAddress
    if ($adUserCheck = Get-ADUser -filter {Mail -eq $emailAddress}) {
        Write-Host "Resource found .." -ForegroundColor green -NoNewline
        Set-Aduser -Identity $adUserCheck.DistinguishedName -add @{"ExtensionAttribute6"="Migrating-GALSync"}
    }
    else {
        Write-Host "Resource NOT found .." -ForegroundColor red -NoNewline
        $notFoundResources += $resource
    }
    Write-Host "done" -ForegroundColor Green
}

#Create Resource Contact

function Create-ResourceMailContacts {
    param (
        [Parameter(Mandatory=$True)] [string] $ImportCSV,
        [Parameter(Mandatory=$True)] [string] $FailureOutputCSV
    )

    $mailboxes = Import-Csv $ImportCSV
    #ProgressBar
    $progressref = ($mailboxes).count
    $progresscounter = 0
    $alreadyExists = @()
    $createdUsers = @()
    $failures = @()

    foreach ($user in $mailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Mailboxes for $($user.DisplayName)"

    if ($recipientCheck = Get-Recipient $user.PrimarySmtpAddress -ea silentlycontinue)
    {
        Write-Host "Contact Already Exists for $($user.PrimarySmtpAddress)" -ForegroundColor Cyan
        $alreadyExists += $user
    }
    else {
        try {
            New-MailContact -DisplayName $user.DisplayName -Name $user.DisplayName -ExternalEmailAddress $user.PrimarySmtpAddress -ErrorAction Stop
            Write-Host "Created Contact $($user.PrimarySmtpAddress)" -ForegroundColor Green
            $createdUsers += $user
        }
        catch
        {
            Write-Warning -Message "$($_.Exception)"
            $currenterror = new-object PSObject

            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | add-member -type noteproperty -name "User" -Value $user.PrimarySmtpAddress
            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
            $failures += $currenterror

            #Attempt 2 on creating Contact
            $newName = $user.DisplayName + " - Jefferson"
            try {
                New-MailContact -DisplayName $user.DisplayName -Name $newName -ExternalEmailAddress $user.PrimarySmtpAddress -ErrorAction Stop
                Write-Host "Attempt 2 Created Contact $($newName)" -ForegroundColor Yellow
            }
            catch {
                Write-Warning -Message "$($_.Exception)"
                $currenterror = new-object PSObject

                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | add-member -type noteproperty -name "User" -Value $user.PrimarySmtpAddress
                $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                $failures += $currenterror
            }  
        }
    }
    }
    Write-Host "$($mailusers.count) MailUsers found in list" -ForegroundColor Yellow
    Write-Host "$($notfoundUsers.count) Users Not Found in list" -ForegroundColor Yellow
    Write-Host "$($failures.count) Failures" -ForegroundColor Yellow

    $failures | Export-Csv -NoTypeInformation -Encoding utf8 $FailureOutputCSV
}

## Create VIP Users in Contacts

function Get-MAILBOXDETAILS {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$True)] [string] $ImportCSV
        )

    $mailboxes = Import-CsV $ImportCSV
    $AllUsers = @()

    #ProgressBar
    $progressref = ($Mailboxes).count
    $progresscounter = 0

    foreach ($user in $Mailboxes)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.userprincipalname)"

        Write-Host "$($user.userprincipalname) .." -ForegroundColor Cyan -NoNewline

        $mailbox = Get-Mailbox $user.userprincipalname | select PrimarySMTPAddress,name,alias,UserPrincipalName,CustomAttribute7
        $MSOLUser = Get-MsolUser -userprincipalname $mailbox.userprincipalname
        
        $currentuser = new-object PSObject
        
        $currentuser | add-member -type noteproperty -name "DisplayName" -Value $msoluser.DisplayName -Force
        $currentuser | add-member -type noteproperty -name "Name" -Value $mailbox.Name -Force
        $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $msoluser.userprincipalname -Force
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $mailbox.PrimarySMTPAddress -Force
        $currentuser | add-member -type noteproperty -name "CustomAttribute7" -Value $mailbox.CustomAttribute7 -Force
        $currentuser | add-member -type noteproperty -name "Alias" -Value $mailbox.alias -Force
        $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $msoluser.IsLicensed -Force
        $currentuser | add-member -type noteproperty -name "City" -Value $msoluser.City -Force
        $currentuser | add-member -type noteproperty -name "Country" -Value $msoluser.Country -Force
        $currentuser | add-member -type noteproperty -name "Department" -Value $msoluser.Department -Force
        $currentuser | add-member -type noteproperty -name "Fax" -Value $msoluser.Fax -Force
        $currentuser | add-member -type noteproperty -name "FirstName" -Value $msoluser.FirstName -Force
        $currentuser | add-member -type noteproperty -name "LastName" -Value $msoluser.LastName -Force
        $currentuser | add-member -type noteproperty -name "MobilePhone" -Value $msoluser.MobilePhone -Force
        $currentuser | add-member -type noteproperty -name "Office" -Value $msoluser.Office -Force
        $currentuser | add-member -type noteproperty -name "PhoneNumber" -Value $msoluser.PhoneNumber -Force
        $currentuser | add-member -type noteproperty -name "PostalCode" -Value $msoluser.PostalCode -Force
        $currentuser | add-member -type noteproperty -name "State" -Value $msoluser.State-Force 
        $currentuser | add-member -type noteproperty -name "StreetAddress" -Value $msoluser.StreetAddress -Force
        $currentuser | add-member -type noteproperty -name "Title" -Value $msoluser.Title -Force
        
        Write-Host "done" -ForegroundColor Green
        $AllUsers += $currentuser
    }
    #Export
    $AllUsers | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}

#Gather Batch1 Users
function Get-BatchMAILBOXDETAILS {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$True)] [string] $ImportCSV
        )

    $mailboxes = Import-CsV $ImportCSV
    $AllUsers = @()

    #ProgressBar
    $progressref = ($Mailboxes).count
    $progresscounter = 0

    foreach ($user in $Mailboxes)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.userprincipalname)"

        Write-Host "$($user.userprincipalname) .." -ForegroundColor Cyan -NoNewline

        $mailbox = Get-Mailbox $user.userprincipalname | select PrimarySMTPAddress
        $user | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $mailbox.PrimarySMTPAddress -Force

        Write-Host "done" -ForegroundColor Green
        $AllUsers += $user
    }
    #Export
    $AllUsers | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}

#Create MailContacts
function Create-MailContacts {
    param (
        [Parameter(Mandatory=$True)] [string] $ImportCSV,
        [Parameter(Mandatory=$True)] [string] $FailureOutputCSV
    )

    $mailboxes = Import-Csv $ImportCSV
    #ProgressBar
    $progressref = ($mailboxes).count
    $progresscounter = 0
    $alreadyExists = @()
    $createdUsers = @()
    $failures = @()

    foreach ($user in $mailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Mailboxes for $($user.DisplayName)"

    if ($recipientCheck = Get-Recipient $user.PrimarySMTPAddress -ea silentlycontinue)
    {
        Write-Host "Contact Already Exists for $($user.PrimarySMTPAddress)" -ForegroundColor Cyan
        $alreadyExists += $user
    }
    else {
        try {
            New-MailContact -DisplayName $user.DisplayName -Name $user.name -ExternalEmailAddress $user.PrimarySMTPAddress -ErrorAction Stop
            Write-Host "Created Contact $($user.PrimarySMTPAddress)" -ForegroundColor Green
            $createdUsers += $user
        }
        catch
        {
            Write-Warning -Message "$($_.Exception)"
            $currenterror = new-object PSObject

            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | add-member -type noteproperty -name "User" -Value $user.PrimarySMTPAddress
            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
            $failures += $currenterror

            #Attempt 2 on creating Contact
            $newName = $user.DisplayName + " - Jefferson"
            try {
                New-MailContact -DisplayName $user.DisplayName -Name $newName -ExternalEmailAddress $user.PrimarySMTPAddress -ErrorAction Stop
                Write-Host "Attempt 2 Created Contact $($newName)" -ForegroundColor Yellow
            }
            catch {
                Write-Warning -Message "$($_.Exception)"
                $currenterror = new-object PSObject

                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | add-member -type noteproperty -name "User" -Value $user.PrimarySMTPAddress
                $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                $failures += $currenterror
            }  
        }
    }
    }
    Write-Host "$($mailusers.count) MailUsers found in list" -ForegroundColor Yellow
    Write-Host "$($notfoundUsers.count) Users Not Found in list" -ForegroundColor Yellow
    Write-Host "$($failures.count) Failures" -ForegroundColor Yellow

    $failures | Export-Csv -NoTypeInformation -Encoding utf8 $FailureOutputCSV
}


function Update-MailContacts {
    param (
        [Parameter(Mandatory=$True)] [string] $ImportCSV,
        [Parameter(Mandatory=$True)] [string] $FailureCSVPath
    )
    $mailboxes = Import-Csv $ImportCSV
    $progressref = ($mailboxes).count
    $progresscounter = 0
    $updatedUsers = @()
    $notfoundUsers = @()
    $mailUsers = @()
    $failures = @()

    #Set Mail Contact Attributes
    foreach ($user in $mailboxes) {
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updated Mailboxes for $($user.DisplayName)"
        $EmailAddress = $user.UserPrincipalName.tostring()
        $PrimarySMTPAddress = $user.PrimarySMTPAddress.tostring()

        if ($recipientCheck = Get-Recipient $PrimarySMTPAddress -ResultSize unlimited -ea silentlycontinue)
        {
            #Update AzureAD GuestUser Attributes
            if ($recipientCheck.RecipientTypeDetails -eq "GuestMailUser") {
                try {
                    $mailUsers += $user
                    Write-Host "$($User.DisplayName) Found as MailUser" -foregroundcolor Yellow   
                    $azureADUser = Get-AzureADUser -Filter "Mail eq '$PrimarySMTPAddress'" -ErrorAction Stop | select ObjectID

                    #Set Guest User
                    Set-AzureADUser -ObjectId $azureADUser.ObjectId -Department $user.Department -PhysicalDeliveryOfficeName $user.Office -TelephoneNumber $user.PhoneNumber -JobTitle $user.Title -City $user.City -State $user.State -StreetAddress $user.StreetAddress
                    Set-MailUser $PrimarySMTPAddress -CustomAttribute5 "GALSync"
                }
                catch{
                    Write-Warning -Message "$($_.Exception)"
                    $currenterror = new-object PSObject

                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | add-member -type noteproperty -name "User" -Value $PrimarySMTPAddress
                    $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                    $failures += $currenterror
                } 
            }
            #Update MailContact Attributes
            else {
                try {
                    Set-Contact $PrimarySMTPAddress -Department $user.Department -Fax $user.fax -Office $user.Office -Phone $user.PhoneNumber -Title $user.Title -city $user.City -state $user.State -StreetAddress $user.StreetAddress -wa silentlycontinue -ea stop
                    Set-MailContact $PrimarySMTPAddress -CustomAttribute5 "GALSync" -ea stop -wa silentlycontinue
                    Set-MailContact $PrimarySMTPAddress -EmailAddresses @{add=$EmailAddress} -ea stop -wa silentlycontinue
                    Write-Host "Updated Contact $($PrimarySMTPAddress)" -ForegroundColor Green
                    $updatedUsers += $user
                }
                catch {
                    Write-Warning -Message "$($_.Exception)"
                    $currenterror = new-object PSObject

                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | add-member -type noteproperty -name "User" -Value $PrimarySMTPAddress
                    $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                    $failures += $currenterror
                } 
            }
        }
        elseif ($recipientCheck = Get-Recipient $EmailAddress -ResultSize unlimited -ea silentlycontinue) {
            #Update AzureAD GuestUser Attributes
            if ($recipientCheck.RecipientTypeDetails -eq "GuestMailUser") {
                try {
                    $mailUsers += $user
                    Write-Host "$($User.DisplayName) Found as MailUser**" -foregroundcolor Yellow   
                    $azureADUser = Get-AzureADUser -Filter "Mail eq '$EmailAddress'" -ErrorAction Stop | select ObjectID

                    #Set Guest User
                    Set-AzureADUser -ObjectId $azureADUser.ObjectId -Department $user.Department -PhysicalDeliveryOfficeName $user.Office -TelephoneNumber $user.PhoneNumber -JobTitle $user.Title -City $user.City -State $user.State -StreetAddress $user.StreetAddress
                    Set-MailUser $EmailAddress -CustomAttribute6 "GALSync"
                }
                catch{
                    Write-Warning -Message "$($_.Exception)"
                    $currenterror = new-object PSObject

                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | add-member -type noteproperty -name "User" -Value $user.UserPrincipalName
                    $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                    $failures += $currenterror
                } 
            }
            #Update MailContact Attributes
            else {
                try {
                    Set-Contact $EmailAddress -Department $user.Department -Fax $user.fax -Office $user.Office -Phone $user.PhoneNumber -Title $user.Title -city $user.City -state $user.State -StreetAddress $user.StreetAddress -wa silentlycontinue -ea stop
                    Set-MailContact $EmailAddress -CustomAttribute5 "GALSync" -ea stop -wa silentlycontinue
                    Set-MailContact $EmailAddress -EmailAddresses @{add=$PrimarySMTPAddress} -ea stop -wa silentlycontinue
                    Write-Host "Updated Contact $($EmailAddress)**" -ForegroundColor Yellow
                    $updatedUsers += $user
                }
                catch {
                    Write-Warning -Message "$($_.Exception)"
                    $currenterror = new-object PSObject

                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | add-member -type noteproperty -name "User" -Value $user.UserPrincipalName
                    $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                    $failures += $currenterror
                } 
            }
        }
        else {
            Write-Host "No recipient found for $($user.PrimarySMTPAddress)" -ForegroundColor Red
            $notfoundUsers += $user
        }
    }
    Write-Host ""
    Write-Host "$($mailusers.count) MailUsers found in list" -ForegroundColor Yellow
    Write-Host "$($notfoundUsers.count) Users Not Found in list" -ForegroundColor Yellow
    Write-Host "$($failures.count) Users Failed to Update" -ForegroundColor Red
    Write-Host "$($updatedUsers.count) MailContacts Updated" -ForegroundColor Cyan
    Write-Host ""

    $failures | Export-Csv -NoTypeInformation -Encoding utf8 $FailureOutputCSV
}

#Validate Address
function Get-ValidationOjbects {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$True)] [string] $ImportCSV
        )

    $mailboxes = Import-CsV $ImportCSV
    #ProgressBar
    $progressref = ($Mailboxes).count
    $progresscounter = 0

    foreach ($user in $Mailboxes)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Details for $($user.PrimarySMTPAddress) from Migrating Tenant"

        Write-Host "$($user.PrimarySMTPAddress) .." -ForegroundColor Cyan -NoNewline
        try {
            $recipientCheck = Get-Recipient $user.PrimarySMTPAddress -ErrorAction Stop
            $EmailAddresses = $recipientCheck | select -ExpandProperty EmailAddresses

            $user | add-member -type NoteProperty -name "MigratingTenant_DistinguishedName" -Value $recipientCheck.DistinguishedName -Force
            $user | add-member -type NoteProperty -name "MigratingTenant_Name" -Value $recipientCheck.Name -Force
            $user | add-member -type NoteProperty -name "MigratingTenant_RecipientTypeDetails" -Value $recipientCheck.RecipientTypeDetails -Force
            $user | add-member -type NoteProperty -name "MigratingTenant_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress -Force
            $user | add-member -type NoteProperty -name "MigratingTenant_EmailAddresses" -Value ($EmailAddresses -join ",") -Force
            $user | add-member -type noteproperty -name "Failure_Activity" -Value $null -Force
            $user | add-member -type noteproperty -name "Failure_Exception" -Value $null -Force
        }
        catch {
            Write-Warning -Message "$($_.Exception)"
            $user | add-member -type NoteProperty -name "MigratingTenant_DistinguishedName" -Value $null -Force
            $user | add-member -type NoteProperty -name "MigratingTenant_Name" -Value $null -Force
            $user | add-member -type NoteProperty -name "MigratingTenant_RecipientTypeDetails" -Value $null -Force
            $user | add-member -type NoteProperty -name "MigratingTenant_PrimarySMTPAddress" -Value $null -Force
            $user | add-member -type NoteProperty -name "MigratingTenant_EmailAddresses" -Value $null -Force
            $user | add-member -type noteproperty -name "Failure_Activity" -Value $_.CategoryInfo.Activity -Force
            $user | add-member -type noteproperty -name "Failure_Exception" -Value $_.Exception -Force
        }
        Write-Host "done" -ForegroundColor Green
    }
    #Export
    $Mailboxes | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}

#Check Resources
$progressref = ($jefferson_Resources).count
$progresscounter = 0
$notFoundResources = @()
$foundResources = @()

foreach ($resource in $jefferson_Resources) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for Resource $($user.DisplayName)"

    if ($recipientcheck = get-recipient $resource.PrimarySMTPAddress -ea silentlycontinue) {
        $foundResources += $recipientcheck
    }
    else {
        $notFoundResources += $resource
    }
}
Write-Host "$($foundResources.count) Resources found"
Write-Host "$($notFoundResources.count) Resources not Found"
$notFoundResources | ft

#Create Remaining users

$remainingJeffersonUsers = Import-Csv $ImportCSV
#ProgressBar
$progressref = ($remainingJeffersonUsers).count
$progresscounter = 0
$alreadyExists = @()
$createdUsers = @()
$failures = @()

foreach ($user in ($remainingJeffersonUsers | ? {$_.Name})) {
$progresscounter += 1
$progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
$progressStatus = "["+$progresscounter+" / "+$progressref+"]"
Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Mailboxes for $($user.DisplayName)"

if ($recipientCheck = Get-Recipient $user.PrimarySMTPAddress -ea silentlycontinue)
{
    Write-Host "Contact Already Exists for $($user.PrimarySMTPAddress)" -ForegroundColor Cyan
    $alreadyExists += $user
}
else {
    try {
        New-MailContact -DisplayName $user.DisplayName -Name $user.Name -ExternalEmailAddress $user.PrimarySMTPAddress -ErrorAction Stop
        Write-Host "Created Contact $($user.PrimarySMTPAddress)" -ForegroundColor Green
        $createdUsers += $user
    }
    catch
    {
        Write-Warning -Message "$($_.Exception)"
        $currenterror = new-object PSObject

        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | add-member -type noteproperty -name "User" -Value $user.UserPrincipalName
        $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
        $failures += $currenterror

        #Attempt 2 on creating Contact
        $newName = $user.Name + " - Jefferson"
        try {
            New-MailContact -DisplayName $user.DisplayName -Name $newName -ExternalEmailAddress $user.PrimarySMTPAddress -ErrorAction Stop
            Write-Host "Attempt 2 Created Contact $($newName)" -ForegroundColor Yellow
        }
        catch {
            Write-Warning -Message "$($_.Exception)"
            $currenterror = new-object PSObject

            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | add-member -type noteproperty -name "User" -Value $user.PrimarySMTPAddress
            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
            $failures += $currenterror
        }  
    }
}
}
Write-Host "$($mailusers.count) MailUsers found in list" -ForegroundColor Yellow
Write-Host "$($notfoundUsers.count) Users Not Found in list" -ForegroundColor Yellow
Write-Host "$($failures.count) Failures" -ForegroundColor Yellow

$failures | Export-Csv -NoTypeInformation -Encoding utf8 $FailureOutputCSV


## update Mail Contacts PrimarySMTPAddres for not found users
$notfoundUsers = @()
foreach ($user in $remainingJeffersonUsers) {
    $UPNAddress = $user.UserPrincipalName
    $PrimarySMTPAddress = $user.PrimarySMTPAddress
    if ($recipientCheck = get-recipient $PrimarySMTPAddress -ea silentlycontinue) {
        Write-Host "$($PrimarySMTPAddress) found ... " -ForegroundColor Cyan -NoNewline
        Set-MailContact $recipientCheck.PrimarySmtpAddress -EmailAddresses @{add=$UPNAddress}
    }
    else {
        Write-Host "No Recipient Found for $($PrimarySMTPAddress) .. " -ForegroundColor Yellow -NoNewline
        if ($recipientUPNCheck = get-recipient $UPNAddress) {
            Write-Host "$($UPNAddress) found ... " -NoNewline -ForegroundColor DarkMagenta
            Set-MailContact $recipientUPNCheck.PrimarySmtpAddress -EmailAddresses @{add=$PrimarySMTPAddress}
        }
        else {
            Write-Host "No Recipient Found for $($UPNAddress) .. " -ForegroundColor Red
            $notfoundUsers += $user
        }
    }
    Write-Host "done" -ForegroundColor Green
}

## update Mail Contacts PrimarySMTPAddres for not found users
foreach ($user in $remainingJeffersonUsers) {
    $PrimarySMTPAddress = $user.PrimarySMTPAddress
    $jeffersonName = $user.displayName + " - Jefferson"
    if (!($recipientcheck = Get-Recipient $jeffersonName -ea silentlycontinue)) {
        Write-Host "Creating $($PrimarySMTPAddress).  " -NoNewline
        New-MailContact -DisplayName $user.DisplayName -name $jeffersonName -ExternalEmailAddress $PrimarySMTPAddress
        Write-Host "done" -ForegroundColor Green
    }
}

#Update PIM Data for Contacts based on PrimarySMTPAddress
$progressref = ($remainingJeffersonUsers).count
$progresscounter = 0

#Set Mail Contact Attributes
foreach ($user in $remainingJeffersonUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updated Mailboxes for $($user.DisplayName)"
    $PrimarySMTPAddress = $user.PrimarySMTPAddress.tostring()

    Set-Contact $PrimarySMTPAddress -Department $user.Department -Fax $user.fax -Office $user.Office -Phone $user.PhoneNumber -Title $user.Title -city $user.City -state $user.State -StreetAddress $user.StreetAddress
    Set-Mailcontact $PrimarySMTPAddress -customAttribute5 "GALSync"
    Write-Host "Updated Contact $($PrimarySMTPAddress)" -ForegroundColor Green
}

#Compare Lists
$Jefferson_A5_Full = Import-Csv "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson - Einstein to Jefferson Migration\Exchange Online\TJ A5 - Full_2021-10-3.csv"
$Jefferson_AllUsers = Import-Csv "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson External - Einstein to Jefferson Migration\GALSync\Jefferson_Licensed_MailboxOutput.csv"
$jefferson_LicensedUsers = Import-Csv  "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson External - Einstein to Jefferson Migration\GALSync\Jefferson_Licensed_MailboxOutput.csv"

$progressref = ($Jefferson_AllUsers).count
$progresscounter = 0
$updatedUsers = @()
$notfoundUsers = @()
$foundUsers = @()
foreach ($user in $Jefferson_AllUsers) {
    $UPN = $user.userPrincipalName
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for for $($user.displayName)"

    if (($jefferson_LicensedUsers | ?{$_.UserPrincipalName -eq $UPN})) {
        $foundUsers += $user
    }
    else {
        Write-Host "User $($user.UserPrincipalName) not found" -ForegroundColor red
        $notfoundUsers += $user
    }
}
Write-Host "$($notfoundUsers) Users not found."

## Get Azure User Details
function Get-AzureUserDETAILS {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$True)] [string] $ImportCSV
        )

    $mailboxes = Import-CsV $ImportCSV
    $AllUsers = @()

    #ProgressBar
    $progressref = ($Mailboxes).count
    $progresscounter = 0

    foreach ($user in $Mailboxes)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.userprincipalname)"

        Write-Host "$($user.userprincipalname) .." -ForegroundColor Cyan -NoNewline

        $currentuser = new-object PSObject
        $emailAddresses = $user |select -ExpandProperty ProxyAddresses
        if ($user.AssignedLicenses) {
            $licensedState = $true
        }
        else {
            $licensedState = $false
        }
        
        $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName
        $currentuser | add-member -type noteproperty -name "FirstName" -Value $user.GivenName
        $currentuser | add-member -type noteproperty -name "LastName" -Value $user.Surname
        $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $user.userprincipalname
        $currentuser | add-member -type noteproperty -name "PrimarySMTPAddress" -Value $user.Mail
        $currentuser | add-member -type noteproperty -name "IsEnabled" -Value $user.AccountEnabled
        $currentuser | add-member -type noteproperty -name "IsLicensed" -Value $licensedState
        $currentuser | add-member -type noteproperty -name "ObjectType" -Value $user.ObjectType
        $currentuser | add-member -type noteproperty -name "ShowInAddressList" -Value $user.ShowInAddressList
        $currentuser | add-member -type noteproperty -name "CompanyName" -Value $user.CompanyName
        $currentuser | add-member -type noteproperty -name "JobTitle" -Value $user.JobTitle
        $currentuser | add-member -type noteproperty -name "City" -Value $user.City
        $currentuser | add-member -type noteproperty -name "Country" -Value $user.Country
        $currentuser | add-member -type noteproperty -name "Department" -Value $user.Department
        $currentuser | add-member -type noteproperty -name "Office" -Value $user.PhysicalDeliveryOfficeName
        $currentuser | add-member -type noteproperty -name "State" -Value $user.State
        $currentuser | add-member -type noteproperty -name "StreetAddress" -Value $user.StreetAddress
        $currentuser | add-member -type noteproperty -name "PostalCode" -Value $user.PostalCode
        $currentuser | add-member -type noteproperty -name "PhoneNumber" -Value $user.TelephoneNumber
        $currentuser | add-member -type noteproperty -name "MobilePhone" -Value $null
        $currentuser | add-member -type noteproperty -name "Fax" -Value $null
        $currentuser | add-member -type noteproperty -name "EmailAddresses" -Value ($emailAddresses -join ",")
        
        Write-Host "done" -ForegroundColor Green
        $AllUsers += $currentuser
    }
    #Export
    $AllUsers | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}

#Check VIPS
$progressref = ($einsteinVIPs).count
$progresscounter = 0
$notFoundResources = @()
$foundResources = @()

foreach ($resource in $einsteinVIPs) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for User $($user.DisplayName)"

    if ($recipientcheck = get-recipient $resource.PrimarySMTPAddress -ea silentlycontinue) {
        $foundResources += $recipientcheck
    }
    else {
        $notFoundResources += $resource
    }
}
Write-Host "$($foundResources.count) user found"
Write-Host "$($notFoundResources.count) user not Found"

## Set VIP Attributes in Azure
$einsteinVIPsInput = Read-Host "Provide Full Path of Einstein VIP CSV list (no Quotes needed)"
$einsteinVIPs = Import-Csv $einsteinVIPsInput
$progressref = ($einsteinVIPs).count
$progresscounter = 0

Connect-AzureAD

foreach ($user in $einsteinVIPs) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for User $($user.DisplayName)"

    if ($recipientcheck = Get-AzureADUser -SearchString $user.PrimarySMTPAddress -ea silentlycontinue) {
        Set-AzureADUser -ObjectId $recipientCheck.ObjectId -Department $user.Department -GivenName $user.FirstName -Surname $user.LastName -JobTitle $user.Title -StreetAddress $user.StreetAddress -PhysicalDeliveryOfficeName $user.Office -State $user.state -DisplayName $user.DisplayName -FacsimileTelephoneNumber $user.PhoneNumber
    }
}

$einsteinVIPs | foreach {
    $recipientCheck = Get-recipient $_.PrimarySMTPAddress
    if ($recipientCheck.RecipientTypeDetails -eq "MailContact") {
        set-mailcontact $_.PrimarySMTPAddress -HiddenFromAddressListsEnabled $false
    }
    elseif ($recipientCheck.RecipientTypeDetails -eq "GuestMailuser") {
        Set-MailUser $_.PrimarySmtpAddress -HiddenFromAddressListsEnabled $false
    }
}

#Gather Einstein Mailbox Stats in Jefferson Tenant
$einsteinTJGroup = Get-AzureADGroup -SearchString "TJ A5 Einstein Full"
$einsteinTJUsers = Get-AzureADGroupMember -ObjectId $einsteinTJGroup.ObjectId -All $true
$einsteinTJMBXStats = @()
$progressref = ($einsteinTJUsers).count
$progresscounter = 0
foreach ($user in $einsteinTJUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Details for $($user.DisplayName)"
    $mbxStats = Get-MailboxStatistics $user.Mail | select MailboxTypeDetail,LastInteractionTime,LastUserAccessTime,LastLogonTime,TotalItemSize,ItemCount
    $currentuser = new-object PSObject
    $currentuser | add-member -type noteproperty -name "DisplayName" -Value $user.DisplayName
    $currentuser | add-member -type noteproperty -name "UserPrincipalName" -Value $user.UserPrincipalName
    $currentuser | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.Mail
    $currentuser | add-member -type noteproperty -name "RecipientTypeDetails" -Value $mbxStats.MailboxTypeDetail
    $currentuser | add-member -type noteproperty -name "JobTitle" -Value $user.JobTitle
    $currentuser | add-member -type noteproperty -name "Department" -Value $user.Department
    $currentuser | add-member -type noteproperty -name "CompanyName" -Value $user.CompanyName
    $currentuser | add-member -type noteproperty -name "LastDirSyncTime" -Value $user.LastDirSyncTime
    $currentuser | add-member -type noteproperty -name "LastInteractionTime" -Value $mbxStats.LastInteractionTime
    $currentuser | add-member -type noteproperty -name "LastUserAccessTime" -Value $mbxStats.LastUserAccessTime
    $currentuser | add-member -type noteproperty -name "LastLogonTime" -Value $mbxStats.LastLogonTime
    $currentuser | add-member -type noteproperty -name "TotalItemSize" -Value $mbxStats.TotalItemSize
    $currentuser | add-member -type noteproperty -name "ItemCount" -Value $mbxStats.ItemCount
    $einsteinTJMBXStats += $currentuser
}

# recheck not found users
$progressref = ($notfoundUsers).count
$progresscounter = 0
foreach ($user in $notfoundUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking if User $($user.DisplayName) is in Matched Report"

    if (!($mailboxCheck = get-mailbox $user.DisplayName -ea silentlycontinue)) {
        $user | add-member -type noteproperty -name "ExistsinEinstein" -Value $false -Force
        $user | add-member -type noteproperty -name "UserPrincipalName_Einstein" -Value $null -Force
        $user | add-member -type noteproperty -name "CustomAttribute7" -Value $null -Force
    }
    else {
        $user | add-member -type noteproperty -name "ExistsinEinstein" -Value $true -Force
        $user | add-member -type noteproperty -name "UserPrincipalName_Einstein" -Value $mailboxCheck.UserPrincipalName -Force
        $user | add-member -type noteproperty -name "CustomAttribute7" -Value $mailboxCheck.customAttribute7 -Force
    }
}
# Filter out Domains in Sharing report

$ExternalDomains = @()
$progressref = ($SharedWith).count
$progresscounter = 0

foreach ($User in $SharedWith) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking User Array"

    $UserArray = $User."Shared With" -split ","

    $progressref1 = ($UserArray).count
    $progresscounter1 = 0

    foreach ($SharedUser in $UserArray) {
        $progresscounter1 += 1
        $progresspercentcomplete1 = [math]::Round((($progresscounter1/ $progressref1)*100),2)
        $progressStatus1 = "["+$progresscounter1+" / "+$progressref1+"]"
        Write-progress -Id 2 -PercentComplete $progresspercentcomplete1 -Status $progressStatus1 -Activity "Checking $($SharedUser)"

        if ($SharedUser -like "*@*") {
            $ExternalDomains += $SharedUser
        }
    }
}
# Match Einstein to Jefferson based on CustomAttribute7
$einsteinMailboxes = Import-Csv 
# Check if Users in "TJ A5 Einstein Full" are in matched csv list
$einsteinTJGroup = Get-AzureADGroup -SearchString "TJ A5 Einstein Full"
$einsteinTJUsers = Get-AzureADGroupMember -ObjectId $einsteinTJGroup.ObjectId -All $true
$matchedEinsteinUsers = Import-csv

$progressref = ($einsteinTJUsers).count
$progresscounter = 0
$notfoundUsers = @()
foreach ($user in $einsteinTJUsers) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking if User $($user.DisplayName) is in Matched Report"
    $UPN = $user.UserPrincipalName
    if (!($matchedEinsteinUsers | ? {$_.UserPrincipalName_Jefferson -eq $UPN})) {
        Write-Host "User $($user.DisplayName) not found" -foregroundcolor red
        $notfoundUsers += $user
    }
}