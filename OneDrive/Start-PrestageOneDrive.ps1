## Prestage OneDrive Migration Job
function Start-PreStageOneDrive {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the CSV File Path of OneDrive Users')] [array] $ImportCSV,
        [Parameter(Mandatory=$True,HelpMessage="What is the Destination Admin Site URL")] [string] $DestinationURL,
        [Parameter(Mandatory=$false,HelpMessage="Enable OneDrive For Users?")][switch] $RequestOneDrive,
        [Parameter(Mandatory=$false,HelpMessage="Enable OneDrive For Users?")][switch] $AddSecondaryAdmin,
        [Parameter(Mandatory=$false,HelpMessage="Run for Only Licensed Users in CSV Import?")][switch] $LicensedUsersOnly,
        [Parameter(Mandatory=$True,HelpMessage="Domain TLD. IE com, edu")][string] $DomainTLD,
        [Parameter(Mandatory=$True)] 
        [System.Management.Automation.PSCredential] 
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()] $DestinationCredentials
    )
    #Set Up Module, Variables, Credentials, and Connect to SharePoint Sites
    Import-Module Sharegate
    Set-Variable dstSite, destinationUPN, destinationEmailAccount
    $DestinationServiceAccount = $DestinationCredentials.Username
    $destinationTenant = Connect-Site -Url $DestinationURL -Credential $DestinationCredentials
    Connect-SPOService -Url $DestinationURL -Credential $DestinationCredentials
    
    if ($LicensedUsersOnly) {
        $OneDriveUsers = (Import-Csv $ImportCSV) | ? {$_.IsLicensed_Jefferson -eq $true}
    }
    else {
        $OneDriveUsers = Import-Csv $ImportCSV
    }

    #Progress Bar Initial
    $progressref = ($OneDriveUsers).count
    $progresscounter = 0
    $AlreadyExists = @()
    $RequestedSite = @()
    $SiteAdminAdded =@()
    $AlreadySiteAdmin = @()
    $NoOneDriveProvisioned = @()
    $failedToAddAdminToOneDrive = @()
    
    foreach ($user in $OneDriveUsers) {
        #Clear Previous Variables
        Clear-Variable dstSite, destinationUPN, destinationEmailAccount
        $dstSiteUrl = @()

        #Progress Bar Current
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Submitting OneDrive Prestage for $($user.DisplayName_Source)"
        if ($destinationEmailAccount = $user.PrimarySmtpAddress_Destination) {
            $destinationUPN = $user.UserPrincipalName_Destination

            #Run Prestage OneDrive
            if ($RequestOneDrive) {
                #Connect to Destination SharePoint Site
                try {
                    if ($dstSiteUrl = Get-OneDriveUrl -Tenant $destinationTenant -Email $destinationEmailAccount -ErrorAction Stop) {
                        Write-Host "OneDrive $($dstSiteUrl) already Exists for $($destinationEmailAccount)" -foregroundcolor Cyan
                        $AlreadyExists += $dstSiteUrl
                    }
                    else {
                        Request-SPOPersonalSite -UserEmails $destinationEmailAccount -ErrorAction Stop
                        Write-Host "OneDrive Site Requested for $($destinationEmailAccount)" -foregroundcolor Green
                        $RequestedSite += $user
                    }
                }
                catch {
                    #If Site does not exist
                    if ($OneDriveDestinationURLCheck = Get-SPOSite -Filter "Owner -eq '$destinationUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ea SilentlyContinue) {
                        $AlreadyExists += $OneDriveDestinationURLCheck
                    }
                    else {
                        Request-SPOPersonalSite -UserEmails $destinationEmailAccount -ErrorAction Stop
                        Write-Host "OneDrive Site Requested for $($destinationEmailAccount)" -foregroundcolor Green
                        $RequestedSite += $user
                    }
                    
                }
            }
             #Check if OneDrive Admin Added
            if ($AddSecondaryAdmin) {
                try {
                    $dstSiteUrl = Get-OneDriveUrl -Tenant $destinationTenant -Email $destinationEmailAccount -ErrorAction stop
                    if ($DomainTLD -eq "com") {
                        $dstSiteUrlShort = $dstSiteUrl.replace("_com/","_com")
                    }
                    elseif ($DomainTLD -eq "edu"){
                        $dstSiteUrlShort = $dstSiteUrl.replace("_edu/","_edu")
                    }
                    if ($SPOPermUsers = (Get-SPOUser -site $dstSiteUrlShort.tostring() -ErrorAction Stop).LoginName) {
                        Write-Host "Already Site Admin for $($dstSiteUrlShort.tostring())" -ForegroundColor Yellow
                        $AlreadySiteAdmin += $user
                    }
                    else {
                        try {
                            $adminRequest = Set-SPOUser -Site $dstSiteUrlShort.tostring() -LoginName $DestinationServiceAccount.tostring() -IsSiteCollectionAdmin $true -ErrorAction Stop
                            $SiteAdminAdded += $user
                            Write-Host "Site Admin added for $($OneDriveDestinationURLCheck.Url)" -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Unable to Add Admin for $($destinationEmailAccount)" -ForegroundColor Red
                            $failedToAddAdminToOneDrive += $user
                        }
                    }
                }
                catch {
                    try {
                        Write-Host "Checking Site Url for $($destinationUPN) .. " -ForegroundColor Cyan -nonewline
                        $OneDriveDestinationURLCheck = Get-SPOSite -Filter "Owner -eq '$destinationUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ea stop
                        $adminRequest = Set-SPOUser -Site $OneDriveDestinationURLCheck.URL -LoginName $DestinationServiceAccount.tostring() -IsSiteCollectionAdmin $true -ErrorAction Stop
                        Write-Host "Site Admin added for $($OneDriveDestinationURLCheck.Url)" -ForegroundColor Magenta
                    }
                    catch {
                        $NoOneDriveProvisioned += $destinationEmailAccount
                        Write-Host "No OneDrive Site Provisioned for $($destinationEmailAccount)" -foregroundcolor Red
                    } 
                }
            }
        }
    }        
    #Results Output
    if ($RequestOneDrive) {
        Write-Host $RequestedSite.count "Requested Sites" -ForegroundColor Cyan
        $RequestedSite | out-gridview
        Write-Host $AlreadyExists.count "Sites Already Exist" -ForegroundColor Cyan
        $AlreadyExists | out-gridview

    }
    if ($AddSecondaryAdmin) {
        Write-Host $SiteAdminAdded.count "Site Admin Added" -ForegroundColor Cyan
        Write-Host $AlreadySiteAdmin.count "Already Contains $($DestinationServiceAccount) as Site Admin" -ForegroundColor Cyan
        Write-Host $NoOneDriveProvisioned.count "OneDrive Site Not Provisioned" -ForegroundColor Yellow
        Write-Host $failedToAddAdminToOneDrive.count "Failed to Add Admin to OneDrive" -ForegroundColor Magenta
    }
}