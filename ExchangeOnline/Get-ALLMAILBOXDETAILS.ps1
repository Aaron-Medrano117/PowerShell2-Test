function Get-ALLMAILBOXDETAILS {
    param (
        [Parameter(Mandatory=$True)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$False)] [Switch] $OneDriveCheck,
        [Parameter(Mandatory=$False)] [string] $OneDriveURL,
        [Parameter(Mandatory=$False)] [string] $importCSV
        )

    #Build Mailbox Array
    if ($importCSV) 
    {
        $mailboxes = Import-Csv $importCSV
    }
    else 
    {
        $mailboxes = Get-Mailbox -ResultSize Unlimited | ? {$_.primarysmtpaddress -notlike "*DiscoverySearchMailbox*"| sort PrimarySmtpAddress
    } 

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