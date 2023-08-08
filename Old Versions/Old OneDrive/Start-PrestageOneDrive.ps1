## Pre-stage OneDrive
function Write-ProgressHelper {
	param (
	    [int]$ProgressCounter,
	    [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [int]$TotalCount
	)
    $secondsElapsed = (Get-Date) – $global:start
    $progresspercentcomplete = [math]::Round((($progresscounter / $TotalCount)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$TotalCount+"]"

    $progressParameters = @{
        Activity = $Activity
        Status = "$progressStatus $($secondsElapsed.ToString('hh\:mm\:ss'))"
        PercentComplete = $progresspercentcomplete
    }

    # if we have an estimate for the time remaining, add it to the Write-Progress parameters
    if ($secondsRemaining) {
        $progressParameters.SecondsRemaining = $secondsRemaining
    }
    if ($ID) {
        $progressParameters.ID = $ID
    }
    if ($CurrentOperation) {
        $progressParameters.CurrentOperation = $CurrentOperation
    }

    # Write the progress bar
    Write-Progress @progressParameters

    # estimate the time remaining
    $global:secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount – $progresscounter)
}

#Variables
$isExcel = $true
$isCSV = $false
$RequestOneDrive = $true
$AddSecondaryAdmin = $true
$tenant = "OVG"
$DestinationServiceAccount = "spectramig@oakviewgroup.com"
# Destination SharePoint Admin Site
$destinationSharePointAdminURL = "https://oakview-admin.sharepoint.com/"
# Destination Tenant
$destinationTenant = Connect-Site -Url $destinationSharePointAdminURL -Browser


#Destination SharePoint Online Module
Connect-SPOService -Url $destinationSharePointAdminURL -Credential $DestinationCredentials

#Output Hash Table
# Declare a hashtable to keep track of the users and their OneDrive admin status
$userAdminStatus = @{}
# Declare a hashtable to keep track of the users and their OneDrive status
$userStatus = @{}

#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
    
foreach ($user in $importedUsers) {
    #Clear Previous Variables
    Clear-Variable destinationUPN, destinationEmailAccount
    #progress bar
    Write-ProgressHelper -Activity "Reviewing OneDrive for $($user."$($tenant)")" -ProgressCounter ($progresscounter++) -TotalCount ($importedUsers).count
    
    if ($destinationEmailAccount = $user."$($tenant)") {
        $destinationUPN = $user."$($tenant)"
        #Run Prestage OneDrive
        if ($RequestOneDrive) {
            try {
                    # Check using the SharePoint Online Module
                    $oneDriveUrl = Get-SPOSite -Filter "Owner -eq '$destinationEmailAccount' -and URL -like '*-my.sharepoint.com*'" -IncludePersonalSite $true -ErrorAction Stop
                    if ($oneDriveUrl) {
                        Write-Host "OneDrive '$($oneDriveUrl.Url)' already exists for '$destinationEmailAccount'." -ForegroundColor Cyan
                        $userStatus[$destinationEmailAccount] = "Already Exists"
                    }
                    else {
                        # Request the OneDrive site
                        Request-SPOPersonalSite -UserEmails $destinationEmailAccount -ErrorAction Stop
                        Write-Host "OneDrive site requested for '$destinationEmailAccount'." -ForegroundColor Green
                        $userStatus[$destinationEmailAccount] = "Requested Site"
                    }
                }
                catch {
                    $errorMessage = $_.Exception.Message
                    Write-Host "An error occurred while checking OneDrive for '$destinationEmailAccount': $($_.Exception.Message)" -ForegroundColor Red
                    $userStatus[$destinationEmailAccount] = "Failed: $errorMessage"
                }
        }
        #Check if OneDrive Admin Added
        if ($AddSecondaryAdmin) {
                # Check if the OneDrive exists using the SharePoint Online Module
                try {
                    $OneDriveDestinationURLCheck = Get-SPOSite -Filter "Owner -eq '$destinationUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ErrorAction Stop
                    $adminRequest = Set-SPOUser -Site $OneDriveDestinationURLCheck.URL -LoginName $DestinationServiceAccount.tostring() -IsSiteCollectionAdmin $true -ErrorAction Stop
                    Write-Host "The service account has been added as a site admin for '$($OneDriveDestinationURLCheck.Url)'." -ForegroundColor Magenta
                    $userAdminStatus[$destinationEmailAccount] = "Site Admin Added"
                }
                catch {
                    # Check if the error message is due to "access is denied"
                    if ($_.Exception.Message -like "*access is denied*") {
                        Write-Host "Access is denied while checking OneDrive for '$destinationEmailAccount'." -ForegroundColor Red
                        $userAdminStatus[$destinationEmailAccount] = "Access Denied"
                    }
                    else {
                        $errorMessage = $_.Exception.Message
                        Write-Host "An error occurred while adding a site admin for '$destinationEmailAccount': $errorMessage" -ForegroundColor Red
                        $userAdminStatus[$destinationEmailAccount] = "Failed: $errorMessage"
                    }
                }
        }
}
}        
#Results Output

if ($RequestOneDrive) {
    # Print the list of users and their OneDrive status
    $userStatus.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Key): $($_.Value)" -ForegroundColor $(
            switch ($_.Value) {
                "Already Exists" { "Cyan" }
                "Requested Site" { "Green" }
                "Failed" { "Red" }
            }
        )
    }
}
if ($AddSecondaryAdmin) {
    # Print the list of users and their OneDrive admin status
    $userAdminStatus.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Key): $($_.Value)" -ForegroundColor $(
            switch ($_.Value) {
                "Already Site Admin" { "Yellow" }
                "Site Admin Added" { "Green" }
                "Access Denied" { "Red" }
                {$_ -like "Failed:*"} { "Red" }
            }
        )
    }
}
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan