#ProctorU Scripts

<#

Requirements

Subdomain routing to GSUITE
Subdomain routing to Office 365

New Users will need:
External Address Points to GSUITE address
USERID includes primary domain

During the migration, the target domain points to Office365 subdomain

#>

Set-PSGSuiteConfig -ConfigName MyConfig -SetAsDefaultConfig -P12KeyPath "C:\Users\fred5646\Downloads\office365-migration-288301-9326959c65f7.p12" -AppEmail "gmail-onboarding@office365-migration-288301.iam.gserviceaccount.com" -AdminEmail "rackspace@proctoru.com" -Domain "proctoru.com" -Preference "Domain" -ServiceAccountClientID 100424271523516272781

#Export All GSUITE users
$Mailboxes = Get-GSUserList | Where-Object {$_.IsMailboxSetup -eq $True} 

$AllGSUsers = @()
foreach ($User in $Mailboxes) { 

    $EXAttributes = $user.emails.Address
    [array]$UserAttributes = $User | select -ExpandProperty Name
    Write-Host "Getting Details for $($UserAttributes.FullName) ..." -NoNewline -foregroundcolor cyan

    $tmp = new-object PSObject

    $tmp | add-member -type noteproperty -name "UserPrincipalName" -Value $user.user
    $tmp | add-member -type noteproperty -name "LastLoginTime" -Value $User.LastLoginTime
    $tmp | add-member -type noteproperty -name "DisplayName" -Value $UserAttributes.FullName
    $tmp | add-member -type noteproperty -name "FirstName" -Value $UserAttributes.GivenName
    $tmp | add-member -type noteproperty -name "LastName" -Value $UserAttributes.FamilyName
    $tmp | add-member -type noteproperty -name "OrgUnitPath" -Value $user.OrgUnitPath
    $tmp | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.PrimaryEmail
    $tmp | add-member -type noteproperty -name "EmailAddresses" -Value ($EXAttributes -join ",")
    $tmp | add-member -type noteproperty -name "IncludeInGlobalAddressList" -Value $User.IncludeInGlobalAddressList
    $tmp | add-member -type noteproperty -name "IsSuspended" -Value $User.Suspended
    $tmp | add-member -type noteproperty -name "IsAdmin" -Value $User.IsAdmin
   
    $AllGSUsers += $tmp

    Write-Host "done" -foregroundcolor green

}

$AllGSUsers | Export-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\GSuiteUsers.csv" -Encoding utf8 -NoTypeInformation

###  Match Users from GSUITE to Tenant
function Get-MatchedProctorUsers {
    param ()
    $allMatchedUsers =@()
    $foundUsers =@()
    $notFoundUsers =@()
    $foundDisplayName = @()
    $foundUPNPrefix + @()

    foreach ($user in $AllGSUsers)
    {
        #clear array
        $recipientcheck = @()
        
        #Check Users and Stamp Previous GSuite Attributes
        Write-Host "Checking user $($user.DisplayName) in Office 365 ..." -fore Cyan -NoNewline
        $currentuser = New-Object psobject
        $currentuser | Add-Member -Type noteproperty -Name "DisplayName" -Value $user.DisplayName
        $currentuser | Add-Member -Type noteproperty -Name "UserPrincipalName" -Value $user.UserPrincipalName
        $currentuser | Add-Member -Type noteproperty -Name "PrimarySmtpAddress" -Value $user.PrimarySmtpAddress
        $currentuser | Add-Member -Type noteproperty -Name "EmailAddresses" -Value $user.EmailAddresses
        $currentuser | Add-Member -Type noteproperty -Name "IsSuspended" -Value $user.IsSuspended
        $currentuser | Add-Member -Type noteproperty -Name "IsAdmin" -Value $user.IsAdmin
        $currentuser | Add-Member -Type noteproperty -Name "OrgUnitPath" -Value $user.OrgUnitPath
        $currentuser | Add-Member -Type noteproperty -Name "LastLoginTime" -Value $user.LastLoginTime

        #Split UPN and specify MeazureLearning address
        $UPNSplit = $user.UserPrincipalName -split "@"
        $MeazureLearningUPN = $UPNSplit[0] + "@meazurelearning.com"

            #Search based on MeazureLearning UPN
            if ($msoluser = get-msoluser -userprincipalname $MeazureLearningUPN -ea silentlycontinue)
            {
                $foundusers += $user
                Write-Host "found" -ForegroundColor Green
                $recipientcheck = Get-Recipient $msoluser.displayname
                $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true
                $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluser.UserPrincipalName
                $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluser.DisplayName
                $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails
            }
            #Search based on DisplayName
            elseif ($msoluserName = Get-msoluser -searchstring $user.DisplayName -ea silentlycontinue)
            {
                $foundDisplayName += $user
                $foundUsers += $user
                Write-Host "found*" -ForegroundColor Green
                $recipientcheck = Get-Recipient $msoluserName.displayname
                $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true
                $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluserName.UserPrincipalName
                $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluserName.DisplayName
                $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails
            }
            #Search based on UPN Suffix
            elseif ($msoluser2 = Get-msoluser -searchstring $UPNSplit[0] -ea silentlycontinue)
            {
                $foundUPNPrefix += $user.upn
                $foundUsers += $user.upn
                Write-Host "found*" -ForegroundColor Green
                $recipientcheck = Get-Recipient $msoluser2.displayname
                $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true
                $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluser2.UserPrincipalName
                $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluser2.DisplayName
                $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails
            }
            else
            {
                $notfoundusers += $user
                Write-Host "not found" -ForegroundColor red
                $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $false
                $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value ""
                $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value ""
                $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value ""
            }

        $allMatchedUsers += $currentuser
    }
$allMatchedUsers | Export-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\ProctorU_GSUITE_MatchedUsers.csv" -Encoding utf8 -NoTypeInformation
}

## Matched Users per Organization (version 1.1)

$MatchedUsers =@()
$foundUsers =@()
$notFoundUsers =@()
$MultipleUsers = @()

foreach ($user in $InternalProctorUUsers)
{
    #clear array
    $recipientcheck = @()
    $msoluserName = @()
    $msoluser = @()
    
    #Check Users and Stamp Previous GSuite Attributes
    Write-Host "Checking user $($user.DisplayName) in Office 365 ..." -fore Cyan -NoNewline
    $currentuser = New-Object psobject
    $currentuser | Add-Member -Type noteproperty -Name "DisplayName" -Value $user.DisplayName
    $currentuser | Add-Member -Type noteproperty -Name "First Name" -Value $user."First Name"
    $currentuser | Add-Member -Type noteproperty -Name "Last Name" -Value $user."Last Name"
    $currentuser | Add-Member -Type noteproperty -Name "Email Address" -Value $user."Email Address"
    $currentuser | Add-Member -Type noteproperty -Name "Team/Dept" -Value $user."Team/Dept"
    $currentuser | Add-Member -Type noteproperty -Name "Org Unit Path" -Value $user."Org Unit Path"
    $currentuser | Add-Member -Type noteproperty -Name "Rollout Date / Time" -Value $user.'Rollout Date / Time'
    $currentuser | Add-Member -Type noteproperty -Name "Status" -Value $user.Status
    $currentuser | Add-Member -Type noteproperty -Name "Last Sign In" -Value $user."Last Sign In"
    $currentuser | Add-Member -Type noteproperty -Name "Email Usage" -Value $user."Email Usage"
    $currentuser | Add-Member -Type noteproperty -Name "Drive Usage" -Value $user."Drive Usage"
    $currentuser | Add-Member -Type noteproperty -Name "Total Storage" -Value $user."Total Storage"

    $UPNSplit = $user."Email Address" -split "@"

    #Search based on DisplayName
    if ($msoluser = get-msoluser -searchstring $user.DisplayName -all -ea silentlycontinue)
    {
        if ($msoluser.count -gt 1)
        {
            Write-Host "multiple users found. " -foregroundcolor yellow -nonewline
            $MultipleUsers += $user.DisplayName

            $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value "MultipleUsersFound"
            $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "IsLicensed" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "Licenses" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value ""
        }
        else {
            $foundusers += $user
            Write-Host "found. " -ForegroundColor Green -nonewline
            $recipientcheck = Get-Recipient $msoluser.UserPrincipalName -resultsize unlimited -ea silentlycontinue
            $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true
            $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluser.UserPrincipalName
            $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluser.DisplayName
            $currentuser | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msoluser.IsLicensed
            $currentuser | Add-Member -Type noteproperty -Name "Licenses" -Value ($msoluser.Licenses.accountskuid -join ",")
            $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails
            $currentuser | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress

            #OneDrive Details
            $EmailAddressUpdate1 = $msoluser.UserPrincipalName.Replace("@","_")
            $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
            $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2
            
            if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
            {
                Write-Host "Found OneDrive. " -fore green -nonewline
                $foundsite += $ODSite.URL
                $currentuser | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL
            }
            else {
                Write-Host "No OneDriveFound. " -fore red -nonewline
                $notfoundsite += $user.O365_UPN
                $currentuser | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound"
            }  
        }  
    }
    #Search based on UPN
    elseif ($msoluserName = Get-msoluser -searchstring $UPNSplit[0] -all -ea silentlycontinue)
    {
        if ($msoluserName.count -gt 1)
        {
            Write-Host "multiple users found. " -foregroundcolor yellow -nonewline
            $MultipleUsers += $user.DisplayName

            $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value "MultipleUsersFound"
            $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "IsLicensed" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "Licenses" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value ""
        }
        else {
            $foundUsers += $user
            Write-Host "found*. " -ForegroundColor Green -nonewline
            $recipientcheck = Get-Recipient $msoluserName.UserPrincipalName -resultsize unlimited -ea silentlycontinue
            $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true
            $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluserName.UserPrincipalName
            $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluserName.DisplayName
            $currentuser | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msoluserName.IsLicensed.accountskuid
            $currentuser | Add-Member -Type noteproperty -Name "Licenses" -Value ($msoluser.Licenses.accountskuid -join ",")
            $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails
            $currentuser | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress

            #OneDrive Details
            $EmailAddressUpdate1 = $msoluserName.UserPrincipalName.Replace("@","_")
            $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
            $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2
            
            if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
            {
                Write-Host "Found OneDrive. " -fore green -nonewline
                $foundsite += $ODSite.URL
                $currentuser | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL
            }
            else {
                Write-Host "No OneDriveFound. " -fore red -nonewline
                $notfoundsite += $user.O365_UPN
                $currentuser | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound"
            } 
        }
    }
    else
    {
        $notfoundusers += $user
        Write-Host "not found. " -ForegroundColor red -nonewline
        $currentuser | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $false
        $currentuser | Add-Member -Type noteproperty -Name "O365_UPN" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "IsLicensed" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "Licenses" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "RecipientType" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "OneDriveURL" -Value ""
    }
    
    Write-Host "done " -fore green        

    $MatchedUsers += $currentuser
}
$MatchedUsers | Export-Csv -Encoding utf8 -NoTypeInformation


## Matched Users per Organization (version 1.2)
# Works well to simply add values to existing table
function Get-MatchedBatchGSUITEUsers {
    param (
        [Parameter(Mandatory=$True)] $BatchCheck
        )
    $MatchedUsers =@()
    foreach ($user in $BatchCheck)
    {
        #clear array
        $recipientcheck = @()
        $msoluserName = @()
        $msoluser = @()
        $proctorucheck = @()
        
        #Check Users and Stamp Previous GSuite EmailAddresses
        Write-Host "Checking user $($user.DisplayName) in Office 365 ..." -fore Cyan -NoNewline
        $UPNSplit = $user."Email Address" -split "@"
        $proctorucheck = Get-recipient ($UPNSplit[0] + "@proctoru.com") -ea silentlycontinue

        if ($proctorucheck)
        {
            $user | Add-Member -Type noteproperty -Name "ConflictCheck" -Value $true -force
            $user | Add-Member -Type noteproperty -Name "ConflictRecipientType" -Value ($proctorucheck.RecipientTypeDetails -join ",") -force
        }
        else
        {
            $user | Add-Member -Type noteproperty -Name "ConflictCheck" -Value $false -force
            $user | Add-Member -Type noteproperty -Name "ConflictRecipientType" -Value "" -force
        }
        
        #GSuite Address Mapping Value
        [string]$meazurelearningAddress = $user."Email Address".Replace("@proctoru.com","@meazurelearning.com")
        $GSUITEAddressMapping = $user."Email Address" + "," + $meazurelearningAddress
        $user | Add-Member -Type noteproperty -Name "ConflictRecipientType" -Value $GSUITEAddressMapping -force

        #Search based on UPN
        if ($msoluserName = Get-msoluser -UserPrincipalName $meazurelearningAddress -ea silentlycontinue)
        {
            if ($msoluserName.count -gt 1)
            {
                Write-Host "multiple users found. " -foregroundcolor yellow -nonewline

                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value "MultipleUsersFound" -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value "" -force
            }
            else {
                Write-Host "found. " -ForegroundColor Green -nonewline
                $recipientcheck = Get-Recipient $msoluserName.UserPrincipalName -resultsize unlimited -ea silentlycontinue
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluserName.UserPrincipalName -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluserName.DisplayName -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msoluserName.IsLicensed -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value ($msoluserName.Licenses.accountskuid -join ",") -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress -force

                #OneDrive Details
                $EmailAddressUpdate1 = $msoluserName.UserPrincipalName.Replace("@","_")
                $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
                $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2
                
                if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
                {
                    Write-Host "Found OneDrive. " -fore green -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL -force
                }
                else {
                    Write-Host "No OneDriveFound. " -fore red -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound" -force
                } 
            }
        }
        #Search based on UPN
        elseif ($msoluserName = Get-msoluser -searchstring $UPNSplit[0] -all -ea silentlycontinue)
        {
            if ($msoluserName.count -gt 1)
            {
                Write-Host "multiple users found. " -foregroundcolor yellow -nonewline
                
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value "MultipleUsersFound" -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value "" -force
            }
            else {
                Write-Host "found*. " -ForegroundColor Green -nonewline
                $recipientcheck = Get-Recipient $msoluserName.UserPrincipalName -resultsize unlimited -ea silentlycontinue
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluserName.UserPrincipalName -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluserName.DisplayName -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msoluserName.IsLicensed -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value ($msoluserName.Licenses.accountskuid -join ",") -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress -force

                #OneDrive Details
                $EmailAddressUpdate1 = $msoluserName.UserPrincipalName.Replace("@","_")
                $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
                $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2
                
                if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
                {
                    Write-Host "Found OneDrive. " -fore green -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL -force
                }
                else {
                    Write-Host "No OneDriveFound. " -fore red -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound" -force
                } 
            }
        }
        #Search based on DisplayName
        elseif ($msoluser = get-msoluser -searchstring $user.DisplayName -all -ea silentlycontinue)
        {
            if ($msoluser.count -gt 1)
            {
                Write-Host "multiple users found. " -foregroundcolor yellow -nonewline

                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value "MultipleUsersFound" -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value "" -force
                
            }
            else {
                Write-Host "found. " -ForegroundColor Green -nonewline
                $recipientcheck = Get-Recipient $msoluser.UserPrincipalName -resultsize unlimited -ea silentlycontinue
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluser.UserPrincipalName -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluser.DisplayName -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msoluser.IsLicensed -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value ($msoluser.Licenses.accountskuid -join ",") -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress -force

                

                #OneDrive Details
                $EmailAddressUpdate1 = $msoluser.UserPrincipalName.Replace("@","_")
                $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
                $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2
                
                if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
                {
                    Write-Host "Found OneDrive. " -fore green -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL -force
                }
                else {
                    Write-Host "No OneDriveFound. " -fore red -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound" -force
                }
            }  
        }
        else
        {
            Write-Host "not found. " -ForegroundColor red -nonewline
            $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $false -force
            $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "Licenses" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "RecipientType" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "" -force
        }
        
        Write-Host "done " -fore green        

        $MatchedUsers += $user
    }
    $MatchedUsers | Export-Csv -Encoding utf8 -NoTypeInformation -Path (Read-Host "Filepath")
}

## Matched Users per Organization (version 1.2)
# Works well to simply add values to existing table
function Get-MatchedBatchGSUITEUsers2 {
    param (
        [Parameter(Mandatory=$True)] $BatchCheck
        )
    $MatchedUsers =@()

    foreach ($user in $BatchCheck)
    {
        #clear array
        $recipientcheck = @()
        $msoluserName = @()
        $msoluser = @()
        $proctorucheck = @()
        
        #Check Users and Stamp Previous GSuite EmailAddresses
        Write-Host "Checking user $($user.DisplayName) in Office 365 ..." -fore Cyan -NoNewline
        $UPNSplit = $user.UserPrincipalName -split "@"
        $proctorucheck = Get-recipient ($UPNSplit[0] + "@proctoru.com") -ea silentlycontinue

        if ($proctorucheck)
        {
            $user | Add-Member -Type noteproperty -Name "ConflictCheck" -Value $true -force
            $user | Add-Member -Type noteproperty -Name "ConflictRecipientType" -Value ($proctorucheck.RecipientTypeDetails -join ",") -force
        }
        else
        {
            $user | Add-Member -Type noteproperty -Name "ConflictCheck" -Value $false -force
            $user | Add-Member -Type noteproperty -Name "ConflictRecipientType" -Value "" -force
        }
        
        #GSuite Address Mapping Value
        [string]$meazurelearningAddress = $user.PrimarySmtpAddress.Replace("@proctoru.com","@meazurelearning.com")
        $GSUITEAddressMapping = $user.PrimarySmtpAddress + "," + $meazurelearningAddress
        $user | Add-Member -Type noteproperty -Name "ConflictRecipientType" -Value $GSUITEAddressMapping -force

        #Search based on UPN
        if ($msoluserName = Get-msoluser -UserPrincipalName $meazurelearningAddress -ea silentlycontinue)
        {
            if ($msoluserName.count -gt 1)
            {
                Write-Host "multiple users found. " -foregroundcolor yellow -nonewline
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value "MultipleUsersFound" -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value "" -force
            }
            else {
                Write-Host "found. " -ForegroundColor Green -nonewline
                $recipientcheck = Get-Recipient $msoluserName.UserPrincipalName -resultsize unlimited -ea silentlycontinue
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluserName.UserPrincipalName -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluserName.DisplayName -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msoluserName.IsLicensed -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value ($msoluserName.Licenses.accountskuid -join ",") -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress -force

                #OneDrive Details
                $EmailAddressUpdate1 = $msoluserName.UserPrincipalName.Replace("@","_")
                $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
                $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2
                
                if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
                {
                    Write-Host "Found OneDrive. " -fore green -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL -force
                }
                else {
                    Write-Host "No OneDriveFound. " -fore red -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound" -force
                } 
            }
        }
        #Search based on UPN
        elseif ($msoluserName = Get-msoluser -searchstring $UPNSplit[0] -all -ea silentlycontinue)
        {
            if ($msoluserName.count -gt 1)
            {
                Write-Host "multiple users found. " -foregroundcolor yellow -nonewline
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value "MultipleUsersFound" -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value "" -force
            }
            else {
                Write-Host "found*. " -ForegroundColor Green -nonewline
                $recipientcheck = Get-Recipient $msoluserName.UserPrincipalName -resultsize unlimited -ea silentlycontinue
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluserName.UserPrincipalName -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluserName.DisplayName -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msoluserName.IsLicensed -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value ($msoluserName.Licenses.accountskuid -join ",") -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress -force

                #OneDrive Details
                $EmailAddressUpdate1 = $msoluserName.UserPrincipalName.Replace("@","_")
                $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
                $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2
                
                if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
                {
                    Write-Host "Found OneDrive. " -fore green -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL -force
                }
                else {
                    Write-Host "No OneDriveFound. " -fore red -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound" -force
                } 
            }
        }
        #Search based on DisplayName
        elseif ($msoluser = get-msoluser -searchstring $user.DisplayName -all -ea silentlycontinue)
        {
            if ($msoluser.count -gt 1)
            {
                Write-Host "multiple users found. " -foregroundcolor yellow -nonewline
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value "MultipleUsersFound" -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value "" -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value "" -force
                
            }
            else {
                Write-Host "found. " -ForegroundColor Green -nonewline
                $recipientcheck = Get-Recipient $msoluser.UserPrincipalName -resultsize unlimited -ea silentlycontinue
                $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true -force
                $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value $msoluser.UserPrincipalName -force
                $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $msoluser.DisplayName -force
                $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msoluser.IsLicensed -force
                $user | Add-Member -Type noteproperty -Name "Licenses" -Value ($msoluser.Licenses.accountskuid -join ",") -force
                $user | Add-Member -Type noteproperty -Name "RecipientType" -Value $recipientcheck.RecipientTypeDetails -force
                $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value $recipientcheck.PrimarySMTPAddress -force

                

                #OneDrive Details
                $EmailAddressUpdate1 = ($msoluser.UserPrincipalName.Replace("@","_")).replace(".","_")
                #$EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
                $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate1
                
                if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
                {
                    Write-Host "Found OneDrive. " -fore green -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL -force
                }
                else {
                    Write-Host "No OneDriveFound. " -fore red -nonewline
                    $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound" -force
                }
            }  
        }
        else
        {
            Write-Host "not found. " -ForegroundColor red -nonewline
            $user | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $false -force
            $user | Add-Member -Type noteproperty -Name "O365_UPN" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "IsLicensed" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "Licenses" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "RecipientType" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "PrimarySMTPAddress" -Value "" -force
            $user | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "" -force
        }
        
        Write-Host "done " -fore green        

        $MatchedUsers += $user
    }
    $MatchedUsers | Export-Csv -Encoding utf8 -NoTypeInformation -Path (Read-Host "Filepath")
}

### Match Groups
$matchedGroups =@()
foreach ($group in $AllGSGroups | sort Group) {
    $recipientCheck = @()
    Write-Host "Checking user $($group.Group) in Office 365 ... " -fore Cyan -NoNewline
    $addressSplit = $group.PrimarySmtpAddress -split "@"

    if ($recipientCheck = Get-Recipient $addressSplit[0] -ea silentlycontinue)
    {
        Write-Host "found" -ForegroundColor Green
        $group | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true -force
        $group | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $recipientCheck.DisplayName -force
        $group | Add-Member -Type noteproperty -Name "O365_GroupType" -Value $recipientCheck.RecipientTypeDetails -force
        $group | Add-Member -Type noteproperty -Name "O365_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress -force
        $group | Add-Member -Type noteproperty -Name "RecipientTypeDetails" -Value $recipientCheck.RecipientTypeDetails -force
        $group | Add-Member -Type noteproperty -Name "O365_Note" -Value $recipientCheck.Note -force
    }

    elseif ($recipientCheck = Get-Recipient $group.group -ea silentlycontinue)
    {
        Write-Host "found" -ForegroundColor Green
        $group | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $true -force
        $group | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value $recipientCheck.DisplayName -force
        $group | Add-Member -Type noteproperty -Name "RecipientTypeDetails" -Value $recipientCheck.RecipientTypeDetails -force
        $group | Add-Member -Type noteproperty -Name "O365_PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress -force
        $group | Add-Member -Type noteproperty -Name "O365_Note" -Value $recipientCheck.Note -force
    }
    else
    {
        Write-Host "No Group Found" -ForegroundColor Red
        $group | Add-Member -Type noteproperty -Name "ExistsOnO365" -Value $false -force
        $group | Add-Member -Type noteproperty -Name "O365_DisplayName" -Value "" -force
        $group | Add-Member -Type noteproperty -Name "O365_GroupType" -Value "" -force
        $group | Add-Member -Type noteproperty -Name "O365_PrimarySMTPAddress" -Value "" -force
        $group | Add-Member -Type noteproperty -Name "RecipientTypeDetails" -Value "" -force
        $group | Add-Member -Type noteproperty -Name "O365_Note" -Value "" -force
    }
    $matchedGroups += $group
}

## Restrict Teams Creation

$GroupName = "sgIT-IT"  
$AllowGroupCreation = "False"  
  
Connect-AzureAD

$settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id
if(!$settingsObjectID)
{
    $template = Get-AzureADDirectorySettingTemplate | Where-object {$_.displayname -eq "group.unified"}
    $settingsCopy = $template.CreateDirectorySetting()
    New-AzureADDirectorySetting -DirectorySetting $settingsCopy
    $settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id
}

$settingsCopy = Get-AzureADDirectorySetting -Id $settingsObjectID
$settingsCopy["EnableGroupCreation"] = $AllowGroupCreation

if($GroupName)
{
  $settingsCopy["GroupCreationAllowedGroupId"] = (Get-AzureADGroup -SearchString $GroupName).objectid
}
 else {
$settingsCopy["GroupCreationAllowedGroupId"] = $GroupName
}
Set-AzureADDirectorySetting -Id $settingsObjectID -DirectorySetting $settingsCopy

(Get-AzureADDirectorySetting -Id $settingsObjectID).Values

## Check OneDrive URL created
$foundsite = @()
$notfoundsite = @()
foreach ($user in $ProctorUBPOUsers)
{
    Write-host "Checking status for $($user.DisplayName) .. " -foregroundcolor cyan -nonewline
    $UPNSplit = $user.O365_UPN -split "@"
    $URL = '-my.sharepoint.com/personal/' + $UPNSplit[0]
    Write-Host "checking for URL $($URL) .. " -nonewline -fore yellow
    if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
    {
        Write-Host "Found URL $($ODSite.URL)" -fore green
        $foundsite += $ODSite.URL
    }
    else {
        Write-Host "No OneDriveFound." -fore red
        $notfoundsite += $user.O365_UPN
    }
}


#add email addresses to matched users

foreach ($internaluser in $otherProctorUUsers)
{
    Write-Host "Checking User $($internaluser.DisplayName)" -fore darkcyan
    foreach ($gsuiteuser in $gsuiteuserdetails)
    {
        if ($gsuiteuser.DisplayName -eq $internaluser.DisplayName)
        {
            $internaluser | Add-Member -Type noteproperty -Name "GSuiteAddresses" -Value  $gsuiteuser.EmailAddresses -force
        }
        else
        {
            $internaluser | Add-Member -Type noteproperty -Name "GSuiteAddresses" -Value "" -force
        }
    }
}


foreach ($internaluser in $otherProctorUUsers)
{
    Write-Host "Checking User $($internaluser.DisplayName)" -fore darkcyan
    foreach ($gsuiteuser in $gsuiteuserdetails)
    {
        if ($gsuiteuser.PrimarySmtpAddress -eq $internaluser."Email Address")
        {
            $internaluser | Add-Member -Type noteproperty -Name "GSuiteAddresses" -Value  $gsuiteuser.EmailAddresses -force
        }
    }
}

#### Cutover Stuff

#Batch update
$ProctorUInternalUser310Morning = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\Batches\Internal ProctorU User List.csv" | ? {$_."Rollout Date / Time" -like "3/10 Morning*"}
$ProctorUInternalBCobb = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\Batches\Internal ProctorU User List.csv" | ? {$_.DisplayName -like "Benjamin Cobb"}
$CutoverBatch = $ProctorUInternalUser35Morning

### Disable Apps in License
    # check if license is already added and update. If no missing E3, then add license
    # Deskless = Staffhub
    # MCOSTANDARD = Skype for Business
    # KAIZALA_O365_P3 = Kaizala Pro
    # EXCHANGE_S_ENTERPRISE = Exchange Online
    # AccountSkuID = getyardstick:SPE_E3
$OE3Sku = Get-MsolAccountSku | ?{$_.accountskuid -eq "getyardstick:SPE_E3"}
$OE3Sku.AccountSkuId
foreach($user in $CutoverBatch)
{
    Write-Host "Checking user $($user.O365_DisplayName) .. " -ForegroundColor cyan -NoNewline
    if ($msoluser = Get-MsolUser -searchstring $user.O365_DisplayName)
    {
    <#    #Update UPN
        if ($msoluser.UserPrincipalName -eq $user.O365_UPN)
        {
            Write-Host "UPN does not need updating. Skipping ..." -ForegroundColor Yellow -NoNewline
        }
        else
        {
            Write-Host "UPN Updated .. " -ForegroundColor Green -NoNewline
            Set-MsolUserPrincipalName -UserPrincipalName $msoluser.UserPrincipalName -NewUserPrincipalName $user.O365_UPN
        }
#>
        #Update Licenses
        $DisabledArray = "KAIZALA_O365_P3","MCOSTANDARD","Deskless"
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $OE3Sku.AccountSkuId -DisabledPlans $DisabledArray -Verbose
        
        if ($msoluser.licenses.AccountSkuId -eq "getyardstick:SPE_E3")
        {
            Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
            Write-Host "Updated user $($user.O365_DisplayName) license to E3 with disabled apps" -ForegroundColor Green

            if ($msoluser.licenses.AccountSkuId -eq "getyardstick:TEAMS_COMMERCIAL_TRIAL")
            {
                Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -RemoveLicenses "getyardstick:TEAMS_COMMERCIAL_TRIAL"
            }
        }
        elseif ($msoluser.licenses.AccountSkuId -eq "getyardstick:TEAMS_COMMERCIAL_TRIAL")
        {
            Write-Host "Removed Teams Commercial. Updated user $($user.O365_DisplayName) license to E3 with disabled apps" -ForegroundColor Green
            Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -AddLicenses $OE3Sku.AccountSkuId -RemoveLicenses "getyardstick:TEAMS_COMMERCIAL_TRIAL"
            if ($msoluser.licenses.AccountSkuId -eq "getyardstick:SPE_E3")
            {
                Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
            }
        }
        elseif ($msoluser.licenses.AccountSkuId -eq "getyardstick:O365_BUSINESS_PREMIUM")
        {
            Write-Host "Removed Business Standard. Updated user $($user.O365_DisplayName) license to E3 with disabled apps" -ForegroundColor Green
            Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -AddLicenses $OE3Sku.AccountSkuId -RemoveLicenses "getyardstick:O365_BUSINESS_PREMIUM"
            if ($msoluser.licenses.AccountSkuId -eq "getyardstick:SPE_E3")
            {
                Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -LicenseOptions $LicenseOptions -verbose
            }
        }
        else
        {
            Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -AddLicenses $OE3Sku.AccountSkuId -LicenseOptions $LicenseOptions
            Write-Host "Added E3 license for $($user.O365_DisplayName)" -ForegroundColor Green
        }
    }
    else
    {
        Write-Host "No user $($user.O365_DisplayName) found" -ForegroundColor Red
    }
}

### Submit Full Migration
foreach ($migmailbox in $CutoverBatch)
{
    #Retrieve MailboxID
    $ExportSearchAddress = $migmailbox."Email Address"
    $MailboxId = $Mailboxes | ?{$_.ExportEmailAddress -eq $ExportSearchAddress}

    Write-Host "Checking item" $ExportSearchAddress "with ID:" $MailboxId.Id 
    $result = Add-MW_MailboxMigration -Ticket $mwTicket -MailboxId $MailboxId.Id -Type Full -ConnectorId $MailboxId.ConnectorId -UserId $mwTicket.UserId 
}

function Set-CutoverEmailAddresses {
    param (
        [Parameter(Mandatory=$True)] $CutoverBatch
        )
    # Set EmailAddresses After cutover
    foreach ($user in $CutoverBatch)
    {
        Write-Host "Checking user" $user.O365_DisplayName -foregroundcolor cyan
        $mailbox = get-mailbox $user.PrimarySMTPAddress
        $emailAddressSplit = $user."Email Address" -Split("@")
        $meazurelearningAddress = "smtp:" + $emailAddressSplit[0] + "@meazurelearning.com"
        $yardstickAddress = "smtp:" + $emailAddressSplit[0] + "@getyardstick.com"
        $proctoruaddress = "smtp:" + $emailAddressSplit[0] + "@proctoru.com"
    
        if ($mailbox | ?{$_.EmailAddresses -contains $meazurelearningaddress})
        {
            Write-host "MeazureLearning address "$meazurelearningAddress" found" -foregroundcolor green
        }
        else
        {
            Write-host "Meazurelearning address missing. Adding ..." -NoNewline -foregroundcolor yellow
            Set-Mailbox $user.PrimarySMTPAddress -EmailAddresses @{add=$meazurelearningAddress}
            Write-host "done" -foregroundcolor green
        }

        if ($mailbox | ?{$_.EmailAddresses -contains $yardstickAddress})
        {
            Write-host "GetYardStick address "$yardstickAddress" found" -foregroundcolor green
        }
        else
        {
            Write-host "GetYardStick address missing. Adding ..." -NoNewline -foregroundcolor yellow
            Set-Mailbox $user.PrimarySMTPAddress -EmailAddresses @{add=$yardstickAddress}
            Write-host "done" -foregroundcolor green
        }
        if ($mailbox | ?{$_.EmailAddresses -contains $proctoruaddress})
        {
            Write-host "ProctorU address "$proctoruaddress" found" -foregroundcolor green
        }
        else
        {
            Write-host "ProctorU address missing. Adding ..." -NoNewline -foregroundcolor yellow
            Set-Mailbox $user.PrimarySMTPAddress -EmailAddresses @{add=$proctoruaddress}
            Write-host "done" -foregroundcolor green
        }

        #Disable Forwarding
        Write-host "Disable Forwarding.. " -NoNewline -foregroundcolor yellow
        Set-Mailbox $user.PrimarySMTPAddress -DeliverToMailboxAndForward $false -ForwardingAddress $null -wa silentlycontinue
        Write-host "done" -foregroundcolor green

        #Show in GAL
        Write-host "Making Visible in GAL.. " -NoNewline -foregroundcolor yellow
        Set-Mailbox $user.PrimarySMTPAddress -HiddenFromAddressListsEnabled $false -wa silentlycontinue
        Write-host "done" -foregroundcolor green
    }

    ## Add alternate email address
    foreach ($user in $CutoverBatch | ?{$_.EmailAddresses})
    {
        Write-Host "Updating Email Address for $($user.O365_DisplayName) .." -ForegroundColor cyan -NoNewline
        $EmailArray = $user.EmailAddresses -split ","
        foreach ($address in $EmailArray | ? {$_ -notlike "*gsuite.proctoru.com"})
        {
            $mailbox = Get-Mailbox $user.O365_DisplayName | select -ExpandProperty EmailAddresses
            Write-Host "Checking for $($address) aliases" -ForegroundColor Cyan
            $emailAddressSplit = $address -Split("@")
            $meazurelearningAddress = "smtp:" + $emailAddressSplit[0] + "@meazurelearning.com"
            $yardstickAddress = "smtp:" + $emailAddressSplit[0] + "@getyardstick.com"
            $proctoruaddress = "smtp:" + $emailAddressSplit[0] + "@proctoru.com"

            if ($mailbox | ?{$_ -contains $meazurelearningaddress})
            {
                Write-host "MeazureLearning address "$meazurelearningAddress" found" -foregroundcolor green
            }
            else
            {
                Write-host "Meazurelearning address missing. Adding ..." -NoNewline -foregroundcolor yellow
                Set-Mailbox $user.PrimarySMTPAddress -EmailAddresses @{add=$meazurelearningAddress} -wa silentlycontinue
                Write-host "done" -foregroundcolor green
            }

        if ($mailbox | ?{$_ -contains $yardstickAddress})
        {
            Write-host "GetYardStick address "$yardstickAddress" found" -foregroundcolor green
        }
        else
        {
            Write-host "GetYardStick address missing. Adding ..." -NoNewline -foregroundcolor yellow
            Set-Mailbox $user.PrimarySMTPAddress -EmailAddresses @{add=$yardstickAddress} -wa silentlycontinue
            Write-host "done" -foregroundcolor green
        }
        if ($mailbox | ?{$_ -contains $proctoruaddress})
        {
            Write-host "ProctorU address "$proctoruaddress" found" -foregroundcolor green
        }
        else
        {
            Write-host "ProctorU address missing. Adding ..." -NoNewline -foregroundcolor yellow
            Set-Mailbox $user.PrimarySMTPAddress -EmailAddresses @{add=$proctoruaddress} -wa silentlycontinue
            Write-host "done" -foregroundcolor green
        }
        }
    }
}

function Set-UndoCutoverEmailAddresses {
    param (
        [Parameter(Mandatory=$True)] $CutoverBatch
        )
    # Set EmailAddresses After cutover
    foreach ($user in $CutoverBatch)
    {
        Write-Host "Checking user" $user.O365_DisplayName -foregroundcolor cyan
        $mailbox = get-mailbox $user.PrimarySMTPAddress
        $emailAddressSplit = $user."Email Address" -Split("@")
        $proctoruaddress = "smtp:" + $emailAddressSplit[0] + "@proctoru.com"
        $GsuiteForwardaddress = $emailAddressSplit[0] + "@proctoru.com"
    
        if ($mailbox | ?{$_.EmailAddresses -contains $proctoruaddress})
        {
            Write-host "ProctorU address "$proctoruaddress" found .. " -foregroundcolor green -NoNewline
            Set-Mailbox $user.PrimarySMTPAddress -EmailAddresses @{remove=$proctoruaddress}
            Write-host "done" -foregroundcolor green
        }
        else
        {
            Write-host "ProctorU address missing." -foregroundcolor yellow
            
        }

        #Disable Forwarding
        Write-host "Set Forwarding.. " -NoNewline -foregroundcolor yellow
        Set-Mailbox $user.PrimarySMTPAddress -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $GsuiteForwardaddress -wa silentlycontinue
        Write-host "done" -foregroundcolor green

        #Show in GAL
        Write-host "Hiding from GAL.. " -NoNewline -foregroundcolor yellow
        Set-Mailbox $user.PrimarySMTPAddress -HiddenFromAddressListsEnabled $true -wa silentlycontinue
        Write-host "done" -foregroundcolor green
    }

    ## Remove alternate email address
    foreach ($user in $CutoverBatch | ?{$_.EmailAddresses})
    {
        Write-Host "Updating Email Address for $($user.O365_DisplayName) .." -ForegroundColor cyan -NoNewline
        $EmailArray = $user.EmailAddresses -split ","
        foreach ($address in $EmailArray | ? {$_ -notlike "*gsuite.proctoru.com"})
        {
            $mailbox = Get-Mailbox $user.O365_DisplayName | select -ExpandProperty EmailAddresses
            Write-Host "Checking for $($address) aliases" -ForegroundColor Cyan
            $emailAddressSplit = $address -Split("@")
            $proctoruaddress = "smtp:" + $emailAddressSplit[0] + "@proctoru.com"

            if ($mailbox | ?{$_ -contains $proctoruaddress})
            {
                Write-host "ProctorU address "$proctoruaddress" found .." -foregroundcolor green -NoNewline
                Set-Mailbox $user.PrimarySMTPAddress -EmailAddresses @{remove=$proctoruaddress} -wa silentlycontinue
                Write-host "done" -foregroundcolor green
            }
            else
            {
                Write-host "ProctorU address missing." -foregroundcolor yellow
                
            }
        }
    }
}

### Check IF UPNS Match
$mismatchedUPN = @()
$matchedUPN = @()

foreach ($user in $MatchedUsers)
{
    if ($user.O365_UPN -eq $user.PrimarySMTPAddress)
    {
        $matchedUPN += $user.O365_DisplayName
    }
    else
    {
        $mismatchedUPN += $user.O365_DisplayName
    }
}

#Create BPO users
foreach ($user in $ProctorUBPOUsers | ? {$_.ExistsOnO365 -eq $false})
{
    Write-Host "Creating User $($user.DisplayName) .." -ForegroundColor Cyan -NoNewline
    if (Get-Msoluser -UserPrincipalName $user."Email Address" -ea silentlycontinue)
    {
        Write-Host "User Exists. Skipping." -ForegroundColor Yellow
    }
    else
    {
        New-MsolUser -UserPrincipalName $user.O365_UPN -DisplayName $user.DisplayName -FirstName $user."First Name" -LastName $user."Last Name" -Usagelocation US -Password "Procki2021!"
        Write-Host "User Created." -ForegroundColor Green
    }
}

# PreStage OneDrive
$ProctorUBPOUsers = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\Batches\BPO ProctorU User List.csv"
$CutoverBatch = $ProctorUBPOUsers
function Request-MultipleSPOPersonalSitesOLD {
    param (
        [Parameter(Mandatory=$True)] $CutoverBatch
    )
    $list = @()
    $i = 0
    $count = $ProctorUBPOUsers.count

    foreach ($u in $ProctorUBPOUsers) {
        $i++
        Write-Host "$i/$count"

        $upn = $u.O365_UPN
        $list += $upn

        if ($i -eq 199) {
            #We reached the limit
            Request-SPOPersonalSite -UserEmails $list -NoWait
            Start-Sleep -Seconds 10
            $count = ($ProctorUBPOUsers.count - $list.count)
            $list = @()
            $i = 0
        }
    }

    if ($i -gt 0) {
        Request-SPOPersonalSite -UserEmails $list -NoWait
    }
}
 # PreStage OneDrive 2 
function Request-MultipleSPOPersonalSites {
    param (
        [Parameter(Mandatory=$True)] $CutoverBatch
    )
    foreach($user in $CutoverBatch){
        if($oUser = Get-MsolUser -SearchString $user.O365_UPN)
        {
            Write-Host -ForegroundColor Gray $oUser.DisplayName "exists. Checking for OneDrive ..." -nonewline
            $odUser1 = $oUser.UserPrincipalName.Replace(".","_")
            $odUser = $odUser1.Replace("@","_")
            $URL = '-my.sharepoint.com/personal/' + $odUser
            Write-Host "checking for URL $($URL) .. " -nonewline -fore yellow
    
            if($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)" -ea silentlycontinue)
            {
            Write-Host -ForegroundColor Green $ODSite.URL
            }
            else
            {
            Write-Host -ForegroundColor Red $oUser.UserPrincipalName "does not have a drive provisioned. Provisioning OneDrive."
            Request-SPOPersonalSite -UserEmails $oUser.UserPrincipalName -NoWait
            }
        }
        else
        {
            Write-Host -ForegroundColor Red $user.DisplayName "does not exist"
            $user | Out-File "User_Does_Not_Exist.log" -Append
        }
    }
}

## Check for failed users
$failedMigrationUsersDetails = @()
foreach ($migmailbox in $failedmigmailboxes)
{
    $tmp = New-Object PSObject
    $tmp | Add-Member -MemberType NoteProperty -Name "User" -Value $migmailbox -force
    $tmp | Add-Member -MemberType NoteProperty -Name "MigrationStatus" -Value "Failed" -force

    if ($GSUser = Get-GSUser $migmailbox -ea silentlycontinue)
    {
        [array]$UserAttributes = $GSUser | select -ExpandProperty Name

        Write-Host "Getting Details for $($UserAttributes.FullName)" -foregroundcolor cyan
        
        $tmp | Add-Member -MemberType NoteProperty -Name "ExistsInGSuite" -Value $true -force
        $tmp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $UserAttributes.FullName -force
        $tmp | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $GSUser.User -force
        $tmp | Add-Member -MemberType NoteProperty -Name "IsMailboxSetup" -Value $GSUser.IsMailboxSetup -force
        $tmp | Add-Member -MemberType NoteProperty -Name "PrimaryEmail" -Value $GSUser.PrimaryEmail -force
    }
    else
    {
        $tmp | Add-Member -MemberType NoteProperty -Name "ExistsInGSuite" -Value $false -force
        $tmp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "" -force
        $tmp | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value "" -force
        $tmp | Add-Member -MemberType NoteProperty -Name "IsMailboxSetup" -Value "" -force
        $tmp | Add-Member -MemberType NoteProperty -Name "PrimaryEmail" -Value $GSUser.PrimaryEmail -force
    }
    $failedMigrationUsersDetails += $tmp
}

## Add E3 to BPO users
foreach ($user in $E3BPOUsers)
{
    $DisplayName = $user.DisplayName
    $DisplayName
    $msoluser = Get-MsolUser -SearchString $DisplayName
    $DisabledArray = "KAIZALA_O365_P3","MCOSTANDARD","Deskless"
    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId getyardstick:SPE_E3 -DisabledPlans $DisabledArray -Verbose
    Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -RemoveLicenses getyardstick:DESKLESSPACK -AddLicenses getyardstick:SPE_E3 -LicenseOptions $LicenseOptions
    Write-Host "Added E3 license for $($DisplayName)" -ForegroundColor Green
}


## Add E3 to BPO users
foreach ($user in $E3BPOUsers)
{
    $DisplayName = $user.DisplayName
    $DisplayName
    $msoluser = Get-MsolUser -SearchString $DisplayName
    $DisabledArray = "KAIZALA_O365_P3","MCOSTANDARD","Deskless"
    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId getyardstick:SPE_E3 -DisabledPlans $DisabledArray -Verbose
    Set-MsolUserLicense -UserPrincipalName $msoluser.UserPrincipalName -AddLicenses getyardstick:SPE_E3 -LicenseOptions $LicenseOptions
    Write-Host "Added E3 license for $($DisplayName)" -ForegroundColor Green
}

# check if failed users exist

## Check for failed users
$proctoruMigStats = Import-Csv "C:\Users\fred5646\Rackspace Inc\MPS-TS-ProctorU - General\Batches\proctoru_MigrationStatistics.csv"

$failedMigrationUsersDetails = @()
foreach ($migmailbox in $proctoruMigStats | ?{$_.Status -eq "Failed"})
{
    if ($GSUser = Get-GSUser $migmailbox.SourceEmailAddress -ea silentlycontinue)
    {
        [array]$UserAttributes = $GSUser | select -ExpandProperty Name

        Write-Host "Getting Details for $($UserAttributes.FullName)" -foregroundcolor cyan
        
        $migmailbox | Add-Member -MemberType NoteProperty -Name "ExistsInGSuite" -Value $true -force
        $migmailbox | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $UserAttributes.FullName -force
        $migmailbox | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $GSUser.User -force
        $migmailbox | Add-Member -MemberType NoteProperty -Name "IsMailboxSetup" -Value $GSUser.IsMailboxSetup -force
        $migmailbox | Add-Member -MemberType NoteProperty -Name "PrimaryEmail" -Value $GSUser.PrimaryEmail -force
    }
    else
    {
        $migmailbox | Add-Member -MemberType NoteProperty -Name "ExistsInGSuite" -Value $false -force
        $migmailbox | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "" -force
        $migmailbox | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value "" -force
        $migmailbox | Add-Member -MemberType NoteProperty -Name "IsMailboxSetup" -Value "" -force
        $migmailbox | Add-Member -MemberType NoteProperty -Name "PrimaryEmail" -Value "" -force
    }
    $failedMigrationUsersDetails += $migmailbox
}

###
## remove user from migration

foreach ($user in $BPOFailedUsers)
{
    $mwmbx = $allMigMailboxes | ? {$_.exportemailaddress -eq $user}
    $mbxproject = Get-MW_MailboxConnector -id $mwmbx.ConnectorId -ticket $mwticket
    Write-host "Removing user $($user) from $($mbxproject.name). " -foregroundcolor cyan -nonewline
    Remove-MW_Mailbox -Ticket $mwTicket -Id $mwmbx.id -force
    #Read-Host "pause to check"
}


### Match Groups v1.2
$AllGSGroups = Import-csv $filepath

$allGroups =@()
$foundGroups =@()
$notFoundGroups =@()

foreach ($group in $AllGSUITEGROUPS | sort Group) {
    Write-Host "Checking user $($group.Group) in Office 365 ..." -fore Cyan -NoNewline
    $currentgroup = new-object PSObject

    $MeazureLearningGroup = $group.PrimarySmtpAddress.replace("@proctoru.com","@meazurelearning.com")
    if ($EXOGroup = get-group $MeazureLearningGroup -ea silentlycontinue)
    { 
        $foundGroups += $MeazureLearningGroup
        
    
        $currentgroup | add-member -type noteproperty -name "Group" -Value $Group.Name
        $currentgroup | add-member -type noteproperty -name "DirectMembersCount" -Value $Group.DirectMembersCount
        $currentgroup | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $Group.Email
        $currentgroup | add-member -type noteproperty -name "Description" -Value $Group.Description
        $currentgroup | add-member -type noteproperty -name "Members" -Value ($GroupMembers.Email -join ",")
        $currentgroup | add-member -type noteproperty -name "Owners" -Value ($Owners.Email -join ",")
    
        $currentgroup | add-member -type noteproperty -name "AllowExternalMembers" -Value $GroupSettings.AllowExternalMembers
        $currentgroup | add-member -type noteproperty -name "MemberJoinRestriction" -Value $Group_MemberJoinRestriction
        $currentgroup | add-member -type noteproperty -name "ShowInGroupDirectory" -Value $GroupSettings.ShowInGroupDirectory

        Write-Host "found" -ForegroundColor Green
        
    }

    elseif ($EXOGroupDisplay = Get-group $group.DisplayName -ea silentlycontinue)
    {
        $notFoundGroups += $MeazureLearningGroup
        Write-Host "found*" -ForegroundColor Green
        $tmp.O365_UPN = $EXOGroupDisplay.userprincipalname
        $tmp.O365_DisplayName = $EXOGroupDisplay.DisplayName
        $tmp.RecipientType = $EXOGroupDisplay.RecipientTypeDetails
        $tmp.O365_PrimarySMTPAddress = $EXOGroupDisplay.WindowsEmailAddress
        $tmp.ExistsOnO365 = $true
    }

    else
    {
        $notfoundusers += $MeazureLearningGroup
        Write-Host "not found" -ForegroundColor red
        $tmp.ExistsOnO365 = $False
    }

    $allGroups += $tmp
}

## Apply Calendar Perms
foreach ($CalendarPerm in $sharedcalendarperms)
{
    if ($mbx = Get-EXOMailbox $CalendarPerm.User -ea silentlycontinue)
    {
        $UpdatedCalendarPath = $mbx.DisplayName + ":\calendar"
        if ($permMBX = Get-EXOMailbox $CalendarPerm.ID -ea silentlycontinue)
        {
            Write-Host "Updating $($UpdatedCalendarPath). Adding Calendar perms for $($CalendarPerm.User) ..." -NoNewline
            Add-MailboxFolderPermission -Identity $UpdatedCalendarPath -User $permMBX.primarysmtpaddress -AccessRights $CalendarPerm.AccessRole #-whatif
            Write-Host "done" -ForegroundColor Green
        }
        else
        {
            Write-Host "Can't set Permission. $($CalendarPerm.ID) Not Found." -ForegroundColor Red
        }
        
    }
    else
    {
        Write-Host "No Mailbox found for $($CalendarPerm.User)" -ForegroundColor Red
    }
}


#create new groups in Office365 from GSUITE

foreach ($GSGroup in $AllGSUITEGROUPS | ?{$_.Members})
{
    $AddressSplit = $GSGroup.PrimarySmtpAddress -split "@"
    $meazurelearningAddress = $AddressSplit[0] + "@meazurelearning.com"
    $Owners = $GSGroup.Owners -split ","
    $members = $GSGroup.Members -split ","
    [boolean]$BlockExternalSenders = [boolean]::Parse($GSGroup.RequireSenderAuthenticationEnabled)
    New-DistributionGroup -name $GSGroup.group -alias $AddressSplit[0] -PrimarySmtpAddress $meazurelearningAddress -RequireSenderAuthenticationEnabled $BlockExternalSenders -DisplayName $GSGroup.group -Members $Members -ManagedBy $Owners
}



foreach ($GSGroup in $AllGSUITEGROUPS | ?{$_.Members})
{
    if ($DLCheck = Get-DistributionGroup $GSGroup.group)
    {
        #Add EmailAddress
        $ProctorUAddress = $GSGroup.PrimarySmtpAddress
        Set-DistributionGroup $DLCheck.Identity -EmailAddresses @{add=$ProctorUAddress}
        
        
        #Add members
        $members = Get-DistributionGroupMember $dlcheck.identity
        $GSGroupMembers = $GSGroup.Members -split ","
        foreach ($memberverify in $GSGroupMembers)
        {
            $recipientcheck = get-recipient $memberverify -ea silentlycontinue
            if ($memberverify -notlike "*$($recipientcheck.name)")
            {
                Add-DistributionGroupMember $DLCheck.identity -member $memberverify      
            }
        }  
    }
}
    else
    {
        $AddressSplit = $GSGroup.PrimarySmtpAddress -split "@"
        $meazurelearningAddress = $AddressSplit[0] + "@meazurelearning.com"
        New-DistributionGroup -name $GSGroup.group -alias $AddressSplit[0] -PrimarySmtpAddress $meazurelearningAddress -RequireSenderAuthenticationEnabled $BlockExternalSenders -DisplayName $GSGroup.group
    }
}
####
foreach ($GSGroup in $AllGSUITEGROUPS | ?{$_.Members})
{
    if ($DLCheck = Get-DistributionGroup $GSGroup.group)
    {
        Write-Host "Updating DL $($DLCheck.name)" -ForegroundColor Cyan
        $ProctorUAddress = $GSGroup.PrimarySmtpAddress
        $Owners = $GSGroup.Owners -split ","
        [boolean]$BlockExternalSenders = [boolean]::Parse($GSGroup.RequireSenderAuthenticationEnabled)
        [boolean]$HiddenFromGAL = [boolean]::Parse($GSGroup.HiddenFromAddressList)
        Set-DistributionGroup $DLCheck.Identity -RequireSenderAuthenticationEnabled $BlockExternalSenders -HiddenFromAddressListsEnabled $HiddenFromGAL -ManagedBy $Owners
    }
}

## Add QA members
$members = $hoovermembers -split ","
foreach ($member in $members)
{
    if ($recipientcheck = get-recipient $member -ea silentlycontinue)
    {
        Write-Host "Adding Member $($recipientcheck.DisplayName)"
        Add-DistributionGroupMember "hoover@meazurelearning.com" -member $recipientcheck.identity
    }
}

## Re-add members
foreach ($GSGroup in $AllGSUITEGROUPS | ?{$_.Members})
{
    if ($DLCheck = Get-DistributionGroup $GSGroup.PrimarySmtpAddress)
    {
       Write-Host "Updating Members for $($DLCheck.name) .." -ForegroundColor cyan -NoNewline
        #Add members
        $GSGroupMembers = $GSGroup.Members -split ","
        foreach ($GSmember in $GSGroupMembers)
        {					                   	
                if ($recipientcheck = Get-Recipient $GSmember -ea silentlycontinue)
                {
                    Add-DistributionGroupMember $GSGroup.PrimarySmtpAddress -member $GSmember -ea silentlycontinue #-whatif
                    Write-Host "added" $GSmember "as member" -ForegroundColor darkGreen -NoNewline 
                }
                else
                {
                    Write-Host "No user found for" $GSmember -ForegroundColor yellow
                }		
        }
        Write-Host "done" -ForegroundColor Green
    }
}

## Check if Members exist

$missingmembers = @()
foreach ($group in $AllGSUITEGROUPS | sort Group) {
    $GSGroupMembers = $group.Members -split ","
    Write-Host $GSGroupMembers.count "Found for $($group.group)"
    foreach ($GSMember in $GSGroupMembers)
    {
        if (!($recipientcheck = Get-Recipient $gsmember -ea silentlycontinue))
        {
            Write-Host "Recipient not found for $($GSMember)" -ForegroundColor red
            $currentmember = new-object PSObject
            $currentmember | add-member -type noteproperty -name "GroupName" -Value $Group.Group
            $currentmember | add-member -type noteproperty -name "GroupPrimarySMTPAddress" -Value $Group.PrimarySmtpAddress
            $currentmember | add-member -type noteproperty -name "Member" -Value $GSMember
            $missingmembers += $currentmember
        }   
    }  
}
        
#OneDrive Details
$failedmigprojects


$foundMsolUser = @()
$notfoundMsolUser = @()
$foundsite = @()
$notfoundsite = @()
$allFailedMoverUsers = @()

foreach ($ODUser in $failedmigprojects) {
    Write-Host "Checking if $($ODUser.destination) exists in 365 .. " -ForegroundColor Cyan -NoNewline
    $currentuser = new-object PSObject
    $currentuser | Add-Member -Type noteproperty -Name "MoverIO_SourceUser" -Value $ODUser.Source
    $currentuser | Add-Member -Type noteproperty -Name "MoverIO_DestinationUser" -Value $ODUser.destination
    $currentuser | Add-Member -Type noteproperty -Name "MoverIOU_Status" -Value $ODUser."Last Status"

    #Get Batch
    $TagsSplit = $ODUser.Tags -split ","
    $currentuser | Add-Member -Type noteproperty -Name "Batch" -Value $TagsSplit[1]

    if ($msolUserCheck = Get-MsolUser -UserPrincipalName $ODUser.destination -ea silentlycontinue)
    {
        Write-Host "Found. " -ForegroundColor Green -NoNewline
        $foundMsolUser += $msolUserCheck | Select DisplayName, UserPrincipalName, IsLicensed

        #Add attributes to output
        $currentuser | Add-Member -Type noteproperty -Name "DisplayName" -Value $msolUserCheck.DisplayName
        $currentuser | Add-Member -Type noteproperty -Name "UserPrincipalName" -Value $msolUserCheck.UserPrincipalName
        $currentuser | Add-Member -Type noteproperty -Name "IsLicensed" -Value $msolUserCheck.IsLicensed

        #check OneDrive
        
        $EmailAddressUpdate1 = $msolUserCheck.UserPrincipalName.Replace("@","_")
        $EmailAddressUpdate2 = $EmailAddressUpdate1.Replace(".","_")
        $URL = '-my.sharepoint.com/personal/' + $EmailAddressUpdate2

        if ($ODSite = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like $($URL)")
        {
            Write-Host "Found OneDrive. " -fore green
            $foundsite += $ODSite.URL
            $currentuser | Add-Member -Type noteproperty -Name "OneDriveURL" -Value $ODSite.URL
        }
        else {
            Write-Host "No OneDriveFound. " -fore red
            $notfoundsite += $user.O365_UPN
            $currentuser | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound"
        }

        #Check in GSUITE
        if ($GSUserDetails = Get-GSUser $ODUser.Source -ea silentlycontinue)
        {
            $currentuser | Add-Member -Type noteproperty -Name "ExistsInGSUITE" -Value $True
            $currentuser | Add-Member -Type noteproperty -Name "OrgUnitPath" -Value $GSUserDetails.OrgUnitPath
            $currentuser | Add-Member -Type noteproperty -Name "LastLoginTime" -Value $GSUserDetails.LastLoginTime
        }
        else
        {
            $currentuser | Add-Member -Type noteproperty -Name "ExistsInGSUITE" -Value $False
            $currentuser | Add-Member -Type noteproperty -Name "OrgUnitPath" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "LastLoginTime" -Value ""
        }
    

    }
    else
    {
        Write-Host "Not Found" -ForegroundColor Red
        $notfoundMsolUser += $ODUser.destination
        $currentuser | Add-Member -Type noteproperty -Name "DisplayName" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "UserPrincipalName" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "IsLicensed" -Value ""
        $currentuser | Add-Member -Type noteproperty -Name "OneDriveURL" -Value "NoSiteFound"
        

        #Check in GSUITE
        if ($GSUserDetails = Get-GSUser $ODUser.Source -ea silentlycontinue)
        {
            $currentuser | Add-Member -Type noteproperty -Name "ExistsInGSUITE" -Value $True
            $currentuser | Add-Member -Type noteproperty -Name "OrgUnitPath" -Value $GSUserDetails.OrgUnitPath
            $currentuser | Add-Member -Type noteproperty -Name "LastLoginTime" -Value $GSUserDetails.LastLoginTime
        }
        else
        {
            $currentuser | Add-Member -Type noteproperty -Name "ExistsInGSUITE" -Value $False
            $currentuser | Add-Member -Type noteproperty -Name "OrgUnitPath" -Value ""
            $currentuser | Add-Member -Type noteproperty -Name "LastLoginTime" -Value ""
        }
    }
    $allFailedMoverUsers += $currentuser
}


#### Set Owners for Rackspace Groups
foreach ($group in $matchedGroups | ?{$_.ExistsOnO365 -eq $true})
{
    $owners = $group.Owners -split ","
    foreach ($user in $owners)
    {
        Write-Host "Add $($user) to group $($group.O365_DisplayName)"
        Set-DistributionGroup $group.O365_PrimarySMTPAddress -ManagedBy @{add=$user}
    }
}

## Update GSUITE User's License
Google Workspace Enterprise Plus

G Suite Basic