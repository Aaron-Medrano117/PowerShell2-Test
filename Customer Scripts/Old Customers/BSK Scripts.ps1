#Get All MailboxDetails
function Get-ALLMAILBOXDETAILS {
    param (
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath
    )
    $AllMailboxProperties = @()
    $mailboxes = Get-Mailbox -ResultSize Unlimited | sort PrimarySmtpAddress
    
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
        Write-Host "Gathering Mailbox Stats for $($user.DisplayName) .." -ForegroundColor Cyan -NoNewline

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
            Write-Host "no Archive found ..." -ForegroundColor Red -NoNewline
        }  

        #get OneDrive Site details
        
        $SPOSite = $null
        $EmailAddressUpdate1 = $MSOLUser.UserPrincipalName.Replace("@","_")
        $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
        $ODSite = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2

        try 
        {
            $SPOSITE = Get-SPOSITE -IncludePersonalSite $true -filter "url -like $ODSite" -ErrorAction SilentlyContinue
            Write-Host "Gathering OneDrive Details ..." -ForegroundColor Cyan -NoNewline
            
            $currentuser | Add-Member -type NoteProperty -Name "OneDriveURL" -Value $SPOSITE.url
            $currentuser | Add-Member -type NoteProperty -Name "Owner" -Value $SPOSITE.Owner
            $currentuser | Add-Member -type NoteProperty -Name "StorageUsageCurrent" -Value $SPOSITE.StorageUsageCurrent
            $currentuser | Add-Member -type NoteProperty -Name "Status" -Value $SPOSITE.Status
            $currentuser | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability" -Value $SPOSITE.SiteDefinedSharingCapability
            $currentuser | Add-Member -type NoteProperty -Name "LimitedAccessFileType" -Value $FDUser.LimitedAccessFileType
            
            Write-Host "done" -ForegroundColor Green
        }
        catch 
        {
            Write-Host "OneDrive Not Enabled for User" -ForegroundColor Yellow
        }

    $AllMailboxProperties += $currentuser
    }
    $AllMailboxProperties | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
}
Get-ALLMAILBOXDETAILS

## Move exisitng OneDrive Users from NAM to EUR region
foreach ($user in ($BSKUsers | ? {$_.OneDriveURL_Destination}))
{
    Write-Host "Moving OneDrive from NAM to EUR for $($user.DisplayName) .. " -ForegroundColor Cyan -NoNewline
    Start-SPOUserAndContentMove -UserPrincipalName $user.UserPrincipalName_Destination -DestinationDataLocation EUR
}

#Check Status OneDrive Users from NAM to EUR region
foreach ($user in ($BSKUsers | ? {$_.OneDriveURL_Destination}))
{
    Get-SPOUserAndContentMoveState -UserPrincipalName $user.UserPrincipalName_Destination | ft
}

#Check Recipients
$foundusers =@()
$notfoundUsers = @()

foreach ($user in $BSKUsers)
{
    Write-Host "Checking $($user)"
    if ($recipientCheck = Get-EXORecipient $user -ea silentlycontinue)
    {
        $foundusers += $recipientCheck
    }
    else
    {
        $notfoundUsers += $user
    }
}

foreach ($row in $table) {
    $results = Set-UserAndGroupMapping -MappingSettings $mappingSettings -Source $row.SourceValue -Destination $row.DestinationValue $row.sourcevalue
    }

$table | foreach {$result = Set-UserAndGroupMapping -MappingSettings $mappingSettings -Source $_.SourceValue -Destination $_.DestinationValue}


#match numbers

$matchednumbers2 = @()
$notmatchednumbers2 = @()
foreach ($number in ($bskTeleNumbers | ?{$_.TargetType}))
{
    foreach ($bskuser in $bskUserNumbers)
    {
            if ($number.id -eq $bskuser.Number)
            {
                Write-Host "Found Match for $($bskuser.DisplayName)" -ForegroundColor Green
                $matchednumbers2 += $bskuser
            }
            else
            {
                Write-Host "No Match found for $($number.id)" -ForegroundColor Red
            $notmatchednumbers2 += $number
            }
    }
}

### Move Sites to EUR
$updatedSPOSites = @()
$notFoundSPOSites = @()

foreach ($site in ($bsksi$temapping | ?{$_."Site Type" -ne "M365Group"}))
{
    $RORSite = $site.DestinationSite
    if ($SPOSiteCheck = Get-SPOSITE -Identity $RORSite)
    {
        Write-Host "Updating Destination Location for $($RORSite)"
        $result = Start-SPOSiteContentMove -SourceSiteUrl $RORSite -DestinationDataLocation EUR
        $result | ft
        $updatedSPOSites += $SPOSiteCheck
    }
    else
    {
        Write-Host "No Team found for $($groupname)" -ForegroundColor Red
        $notFoundSPOSites += $site
    }
}


#Update GroupSites
$updatedTeams = @()
$notFoundTeams = @()
#Connect to US Region
Connect-SPOService -URL https://raxglobal-admin.sharepoint.com/
foreach ($group in $BSKGroups)
{
    $groupname = $group."group name"
    #$groupID = $group."Group ID"
    if ($groupCheck = Get-UnifiedGroup $groupname -ea silentlycontinue)
    {
        Set-SPOUnifiedGroup -PreferredDataLocation EUR -GroupAlias $groupCheck.alias
        $result = Start-SPOUnifiedGroupMove -GroupAlias $groupCheck.alias -DestinationDataLocation EUR #-ValidationOnly
        $result | ft
        $updatedTeams += $groupCheck
    }
    else
    {
        Write-Host "No Team found for $($groupname)" -ForegroundColor Red
        $notFoundTeams += $group
    }
}


# remove I Teams
$removedTeams = @()
$notFoundTeams = @()

foreach ($group in $bskIGroups)
{
    $groupname = $group."group name"
    $groupID = $group."Group ID"
    if ($teamCheck = Get-Team -GroupId $groupID)
    {
        Write-Host "Remvoing Team for $($teamCheck.DisplayName)" -ForegroundColor Green
        Remove-Team -GroupId $groupID -ea silentlycontinue
        $removedTeams += $group
    }
    else
    {
        Write-Host "No Team found for $($groupname)" -ForegroundColor Red
        $notFoundTeams += $group
    }
}

# Set User's phone numbers
foreach ($BSKUser in ($BSKMatchedUsers |? {$_.ExistsOnDestinationTenant -eq $true}))
{
    Write-Host "Update $($BSKUser.DisplayName) .. " -ForegroundColor Cyan -NoNewline
    #Set-MsolUser -UserPrincipalName $BSKUser.UserPrincipalName_Destination -UsageLocation "DE"
    #Start-Sleep -seconds 10
    try
    {  
        Write-Host "Adding $($BSKUser.Number) .. " -ForegroundColor Cyan -NoNewline
        $emergencyLocation = Get-CsOnlineLisLocation -Description $BSKUser.EmergencyLocation
        Set-CsOnlineVoiceUser -Identity $BSKUser.UserPrincipalName_Destination -TelephoneNumber $BSKUser.Number -LocationID $emergencyLocation.locationid.guid
        Write-Host "done" -ForegroundColor green
    }
    catch
    {
        Write-Host "failed" -ForegroundColor red
    }
}


#match numbers
function Match-AllMailUsers {
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

    foreach ($mailbox in $ImportedUsers)
    {
        #ProgressBar2
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Mailbox Stats for $($mailbox.DisplayName)"
        
        Write-Host "Checking for $($mailbox.displayName) in Tenant ..." -fore Cyan -NoNewline
        $newAddressSplit = $mailbox.PrimarySmtpAddress -split "@"
        $newMailboxAddress = $newAddressSplit[0] + "@" + $NewDomain
        if ($mailboxcheck = Get-Mailbox $mailbox.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced)
        {
            Write-Host "found mailbox  " -ForegroundColor Green -nonewline
        }
        elseif ($mailboxcheck = Get-Mailbox $newMailboxAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database)
        {
            Write-Host "found mailbox**  " -ForegroundColor Yellow -nonewline
        }
        elseif ($mailboxcheck = Get-Mailbox $mailbox.displayName -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database)
        {
           Write-Host "found mailbox*  " -ForegroundColor Yellow -nonewline
        }
        elseif ($recipientcheck = Get-Recipient $mailbox.PrimarySmtpAddress -ea silentlycontinue)
        {
            $mailboxcheck = Get-Mailbox $recipientcheck.PrimarySmtpAddress -ea silentlycontinue | select PrimarySMTPAddress, RecipientTypeDetails, UserPrincipalName, IsDirSynced, Database 
            Write-Host "found recipient  " -ForegroundColor Yellow -nonewline
        }
        else
        {
            Write-Host "not found" -ForegroundColor red -NoNewline
            $msoluserscheck = @()
            $MBXStats = @()
        }
        if ($mailboxcheck)
        {
            $msoluserscheck = get-msoluser -UserPrincipalName $mailboxcheck.UserPrincipalName -ea silentlycontinue | select DisplayName, IsLicensed, licenses, BlockCredential, UserPrincipalName, PreferredDataLocation
            $MBXStats = Get-MailboxStatistics $mailboxcheck.PrimarySmtpAddress -ErrorAction SilentlyContinue | select TotalItemSize, ItemCount
            $mailbox | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $True
            $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $msoluserscheck.UserPrincipalName
            $mailbox | add-member -type noteproperty -name "IsLicensed_Destination" -Value $msoluserscheck.IsLicensed
            $mailbox | add-member -type noteproperty -name "Licenses_Destination" -Value ($msoluserscheck.Licenses.AccountSkuID -join ",")
            $mailbox | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $mailboxcheck.IsDirSynced
            $mailbox | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $msoluserscheck.PreferredDataLocation
            $mailbox | add-member -type noteproperty -name "Database_Destination" -Value $mailboxcheck.Database
            $mailbox | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $msoluserscheck.BlockCredential
            $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $mailboxcheck.PrimarySmtpAddress
            $mailbox | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $recipientcheck.RecipientTypeDetails   
            $mailbox | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $MBXStats.TotalItemSize
            $mailbox | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $MBXStats.ItemCount

        }
        else 
        {
            $mailbox | Add-Member -type NoteProperty -Name "ExistsOnDestinationTenant" -Value $False
            $mailbox | add-member -type noteproperty -name "UserPrincipalName_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "IsLicensed_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "Licenses_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "IsDirSynced_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "PreferredDataLocation_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "Database_Destination" -Value $mailboxcheck.Database
            $mailbox | add-member -type noteproperty -name "BlockSigninStatus_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $null
            $mailbox | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null  
            $mailbox | Add-Member -type NoteProperty -Name "MBXSize_Destination" -Value $null
            $mailbox | Add-Member -Type NoteProperty -name "MBXItemCount_Destination" -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "OneDriveURL_Destination"  -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "Owner_Destination"  -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "StorageUsageCurrent_Destination" -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "Status_Destination" -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "SiteDefinedSharingCapability_Destination" -Value $null
            $mailbox | Add-Member -type NoteProperty -Name "LimitedAccessFileType_Destination" -Value $null
        }

        Write-host " .. done" -foregroundcolor green

        $AllUsers += $mailbox
        #Read-Host "Stopping to check user"
    }
    $allUsers | Export-Csv -encoding UTF8 -NoTypeInformation $OutputCSVFilePath
}

## Remove us from Admin Groups
$M365Groups = Import-Csv
$failedGroups = @()
foreach ($group in $M365Groups) 
{
    try
    {
        Remove-UnifiedGroupLinks -Identity $group.Name -LinkType Members -Links amedrano@bright-skies.de -Confirm:$false
        Write-Host "Removed Aaron's account from $($group.DisplayName)" -ForegroundColor Cyan
    }
    catch
    {
       Write-Host "Unable to remove Aaron from $($group.DisplayName)"  -ForegroundColor red
       $failedGroups += $group 
    }
    
}


##

this is the script Richard Li sent me to turn all Sharepoint sites to read only
Get-SPOSite -Limit All -IncludePersonalSite $true | % {$URL = $_.URL; Write-host "Locking $URL" -f Green; Set-SPOSite -Identity $URL -LockState Readonly}