#Set Lock State of OneDrive
# Connect to SharePoint Online
Connect-SPOService -Url "https://$tenantName-admin.sharepoint.com" -Credential $adminCredential

<#
Sets the lock state on a site. Valid values are: NoAccess, ReadOnly and Unlock. 
When the lock state of a site is ReadOnly, a message will appear on the site stating that the site is under maintenance and it is read-only.
When the lock state of a site is NoAccess, all traffic to the site will be blocked. 
    If parameter NoAccessRedirectUrl in the Set-SPOTenant cmdlet is set, traffic to sites that have a lock state NoAccess will be redirected to that URL. 
    If parameter NoAccessRedirectUrl is not set, a 403 error will be returned. It isn't possible to set the lock state on the root site collection.
#>
# Iterate through users and set OneDrive to read-only
$onedriveUsers = Import-Excel -WorksheetName "Block List" -Path "C:\Users\amedrano\Arraya Solutions\Spectra - External - Ext - 1777 Spectra to OVG T2T Migration\Project Files\OneDrive List.xlsx"

#Progress Helper

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
    $tenant = "Spectra"
    $siteState = "NoAccess"
    #ProgressBar
    $global:start = Get-Date
    $progresscounter = 1
    [nullable[double]]$global:secondsRemaining = $null
    $TotalCount = ($onedriveUsers).count
    $ProgressPreference = "Continue"

    foreach ($user in $onedriveUsers) {
        $userPrincipalName = $user."UserPrincipalName_$($tenant)"
        Write-ProgressHelper -Activity "Setting OneDrive State to $($siteState) : $($userPrincipalName)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -ID 1  
        if ($SPOSITE = Get-SPOSite -Filter "Owner -eq $($userPrincipalName) -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true) {
            if ($userPrincipalName -ne $null) {
                try {
                    # Set OneDrive to read-only
                    $SPOSITE | Set-SPOSite -Lockstate $siteState
                    Write-Host "Successfully set OneDrive to $($siteState) for " -foregroundcolor green -nonewline
                    Write-Host $($SPOSITE.Url) -foregroundcolor Cyan
                } catch {
                    Write-Host "Error setting OneDrive to $($siteState) for user $($SPOSITE.Url): $($_.Exception.Message)" -foregroundcolor red
                } 
            }
        }
        else {
            Write-Host "No OneDrive found for user "  -foregroundcolor red -nonewline
            Write-Host $($userPrincipalName) -foregroundcolor yellow
        }
    }
