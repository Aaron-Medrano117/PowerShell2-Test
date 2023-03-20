##########

function Import-SharePointOnlineModule() {
    $moduleLocation = "C:\Program Files\WindowsPowerShell\Modules\Microsoft.Online.SharePoint.PowerShell"
    if (((Get-Module -Name "Microsoft.Online.SharePoint.PowerShell") -ne $null) -or ((Get-InstalledModule -Name "Microsoft.Online.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -ne $null))
    {
        Write-Host "Microsoft.Online.SharePoint.PowerShell Module Already Installed and Imported" -ForegroundColor Green
        return;
    }
    elseif (Test-Path $moduleLocation -ErrorAction SilentlyContinue) {
        Import-Module -Name $Microsoft.Online.SharePoint.PowerShell
        Write-Host "Imported Microsoft.Online.SharePoint.PowerShell module" -ForegroundColor Yellow
    }
    else {
        Try {
            Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -ErrorAction Stop
        }
        Catch {
            Write-Error  "Unable to install module for user. Re-run command Install-Module -name Microsoft.Online.SharePoint.PowerShell in Administrative PowerShell"
            return
        }
    }   
}

function Add-OneDriveAdmin{
    param (
    [Parameter(Mandatory=$True,HelpMessage="Enter the email address of the user who needs access to OneDrive?")] [string] $SiteCollAdmin,
    [Parameter(Mandatory=$True,HelpMessage="Enter the email address associated to the target OneDrive”)] [string] $OneDriveUser
)
    Write-Host "Adding $($SiteCollAdmin) as Admin to $($OneDriveUser)'s OneDrive .."

    #Lookup The OneDrive URL
    try {
        $OneDriveSite = Get-SPOSite -Filter "Owner -eq '$OneDriveUser' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop
    }
    catch {
        Write-Error "No One Drive found for $($OneDriveUser). Please ensure the user is licensed correctly and the OneDrive has been fully provisioned"
        return
    }
    #Add Admin to OneDrive
    #Add Site Collection Admin to the OneDrive
    Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SiteCollAdmin -IsSiteCollectionAdmin $True
    Write-Host "Site Collection Admin Added Successfully!" -ForegroundColor Green
    Write-Host "Displaying OneDrive Permissions to $($OneDriveSite.URL)”
    Get-SPOUser -Site $OneDriveSite.Url
}

Import-SharePointOnlineModule

#Connect to SPO Online
$AdminSiteURL = "https://lighthouselife-admin.sharepoint.com"
Write-host “Add User to OneDrive”
Write-host “Use this script to grant another user full access permission to an associate’s OneDrive”

#Connect to SharePoint Online
#Get Credentials to connect to SharePoint Admin Center
Write-host "Enter Admin Credentials to connect to SharePoint Online"
$Cred = Get-Credential
 
#Connect to SharePoint Online Admin Center
Connect-SPOService -Url $AdminSiteURL -credential $Cred

#Add Admin to OneDrive - Gather Details
$global:SiteCollAdmin = Read-Host "What is the Admin Username to Add?"
$global:OneDriveUser =  Read-Host "Enter the email address associated to the target OneDrive"
$global:OneDriveSite = Get-SPOSite -Filter "Owner -eq '$OneDriveUser' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -limit all -ErrorAction Stop

# Initial Add User as Admin to OneDrive
Add-OneDriveAdmin -SiteCollAdmin $global:SiteCollAdmin -OneDriveUser $global:OneDriveUser

#Grant Additional User to Same OneDrive
$GrantAdditionalUser = Read-Host "Would you like to grant an additional user full access user to the OneDrive $($OneDriveSite.URL)?”
if ($GrantAdditionalUser -eq "Yes") {
    #Add Admin to OneDrive - Gather Details
    Add-OneDriveAdmin -OneDriveUser $global:OneDriveUser
}
elseif ($GrantAdditionalUser -eq "No") {
    Write-Host "No Additional user added to $($OneDriveSite.URL)"
    continue
}
else {
    Write-Error "Invalid Option Specified. Please specify either Yes or No."
    $GrantAdditionalUser = Read-Host “Would you like to modify permissions for another OneDrive?”
    if ($GrantAdditionalUser -eq "Yes") {
        Add-OneDriveAdmin -OneDriveUser $global:OneDriveUser -confirm:$true
    }
    elseif ($GrantAdditionalUser -eq "No") {
        Write-Host "No Additional user added to $($global:OneDriveSite.URL) checked again"
        continue
    }
    else {
        continue
    }
}

#Run another permission update to another OneDrive
$UpdateAnotherOnedrive = Read-Host “Would you like to modify permissions for another OneDrive?”
if ($UpdateAnotherOnedrive -eq "Yes") {
    Add-OneDriveAdmin
}
elseif ($UpdateAnotherOnedrive -eq "No") {
    Write-Host "GoodBye"
    return
}
else {
    Write-Error "Invalid Option Specified. Please specify either Yes or No."
    $UpdateAnotherOnedrive = Read-Host “Would you like to modify permissions for another OneDrive?”
    if ($UpdateAnotherOnedrive -eq "Yes") {
        Add-OneDriveAdmin
    }
    elseif ($UpdateAnotherOnedrive -eq "No") {
        Write-Host "GoodBye"
        return
    }
    else {
        return
    }
}