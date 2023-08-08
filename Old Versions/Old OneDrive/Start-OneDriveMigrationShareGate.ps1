#Start-OneDriveMigrationShareGate
function Write-ProgressHelper {
    param (
        [int]$ProgressCounter,
        [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [int]$TotalCount,
        [datetime]$StartTime
    )
    #$ProgressPreference = "Continue"  
    if ($ProgressPreference = "SilentlyContinue") {
        $ProgressPreference = "Continue"
    }

    $secondsElapsed = (Get-Date) - $StartTime
    $secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)
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
    #$secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)

}

#Set Up Module, Variables, Credentials, and Connect to SharePoint Sites
Import-Module Sharegate
Import-Module Microsoft.Online.SharePoint.PowerShell
Set-Variable srcSite, dstSite, srcList, dstList, srcSiteUrl, dstSiteUrl

#Global Settings
## Tenant Names and Service Account - from the Excel File
$TenantSource = "Spectra"
$TenantDestination = "OVG"
$DestinationServiceAccount = "spectramig@oakviewgroup.com"

#SharePoint Admin Sites
$global:DestinationAdminUrl = "https://oakview-admin.sharepoint.com/"
$global:SourceAdminUrl = "https://spectraxp-admin.sharepoint.com/"

#Connect to SharePoint Admin Sites
$sourceTenant = Connect-Site -Url $global:SourceAdminUrl -Browser
$destinationTenant = Connect-Site -Url $global:DestinationAdminUrl -Browser

# Migration Copy Settings
$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate

# Credentials for SharePoint Admin Sites to Check for OneDrive
$sourceCredentials = Get-Credential
$DestinationCredentials = Get-Credential

#Import OneDrive List
$allMatchedUsers = Import-Excel "C:\Users\AMedranoA\Desktop\Spectra-AllMailboxStats-Matched.xlsx"
$OneDriveUsers = $allMatchedUsers[0..29]
$OneDriveUsers = $allMatchedUsers[30..59]
$OneDriveUsers = $allMatchedUsers[60..90]


$OneDriveUsers = $allMatchedUsers | ? {$_.OneDriveStatus_Spectra}
$OneDriveUsers = $OneDriveUsers | ? {$_.CompletionDate -eq "TBD"}

#Output Arrays
$AllOneDriveErrors = @()
$AllOneDriveResults = @()

#ProgressBar
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
$TotalCount = ($OneDriveUsers).count

foreach ($user in $OneDriveUsers) {
    #Clear Previous Variables
    Clear-Variable srcSite, dstSite, srcList, dstList, srcSiteUrl, dstSiteUrl

    #declare Variables
    $sourceUPN = $user.UPN_Falcons
    $destinationUPN = $user.UserPrincipalName
    #progress bar
    Write-ProgressHelper -Activity "Submitting OneDrive Migration for $($sourceUPN)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -ID 1 -StartTime $global:start
    
    #Connect to OneDrive Sites in Source
    write-host "$($sourceUPN): Source OneDrive..." -ForegroundColor Cyan -nonewline
    try {
        if ($null -ne $user."OneDriveURL_$($TenantSource)") {
            $srcSiteUrl = $user."OneDriveURL_$($TenantSource)"
        }
        else {
            Write-Host "Checking for Source OneDrive Url with ShareGate..." -nonewline
            $srcSiteUrl = Get-OneDriveUrl -Tenant $sourceTenant -Email $sourceUPN -ErrorAction stop
            $srcSiteUrl = $srcSiteUrl.TrimEnd('/')
        }
    } catch {
        try {
            #Write-Host "Checking for Source OneDrive Url with SharePoint"
            # Get the Source OneDrive URL with SharePoint
            Connect-SPOService -Url $global:SourceAdminUrl -credential $sourceCredentials
            $srcSiteUrl = (Get-SPOSite -Filter "Owner -eq '$sourceUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ea Stop).Url
        }
        catch {
            Write-Host "Unable to find Source OneDrive Site for $SourceEmailAddress" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "UnableToFindSourceOneDrive" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Tenant" -Value $sourceTenant.Site -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $sourceUPN
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $AllOneDriveErrors += $currenterror
        }
        Continue
    }

    #Get Source OneDrive Site
    try {
        #Get OneDrive Documents Library
        $srcSite = Connect-Site -Url $srcSiteUrl -UseCredentialsFrom $sourceTenant -ErrorAction Stop
        $srcList = Get-List -Site $srcSite -Name "Documents" -ErrorAction Stop
        Write-Host "Connected... " -ForegroundColor Green -nonewline
        
    }
    catch {
        Write-Host "Unable to Connect to Source OneDrive Site Provisioned for $($sourceUPN)" -foregroundcolor Red
        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "UnableToConnectToSourceOneDrive" -Force
        $currenterror | Add-Member -type NoteProperty -Name "Tenant" -Value $sourceTenant.Site -Force
        $currenterror | Add-Member -type NoteProperty -Name "User" -Value $sourceUPN
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllOneDriveErrors += $currenterror      
        Continue
    }

    #Connect to OneDrive Sites in Destination
    if ($srcList) {
        Write-Host "$($destinationUPN): Destination OneDrive..." -ForegroundColor Cyan -nonewline
        try {
            if ($null -ne $user."OneDriveURL_$($TenantDestination)") {
                $dstSiteUrl = $user."OneDriveURL_$($TenantDestination)"
            }
            else {
                Write-Host "Checking for Destination OneDrive Url with ShareGate... " -nonewline
                # Checking for Destination OneDrive Url with ShareGate
                if ($DstSiteUrl = Get-OneDriveUrl -Tenant $global:DestinationTenant -Email $destinationUPN -ErrorAction SilentlyContinue) {
                    $DstSiteUrl = $DstSiteUrl.TrimEnd('/')
                    Write-host "Found.. " -nonewline -ForegroundColor Green
                } else {
                    Write-Host "Checking SharePoint Site Url for $($destinationUPN) .. " -ForegroundColor Cyan -NoNewline
                    Connect-SPOService -Url $global:DestinationAdminUrl -Credential $DestinationCredentials -ErrorAction Stop
                    if ($OneDriveDstUrlCheck = (Get-SPOSite -Filter "Owner -eq '$destinationUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ErrorAction SilentlyContinue)) {
                        Write-Host "Site exists for $($OneDriveDstUrlCheck)" -ForegroundColor Magenta
                    } else {
                        Write-Host "No OneDrive Site Provisioned for $($destinationUPN) .." -ForegroundColor Red -NoNewline
                        Request-SPOPersonalSite -UserEmails $destinationUPN -ErrorAction Stop
                        Write-Host "OneDrive Site Requested for $($destinationUPN)" -ForegroundColor Green
                    }
                }
            }
        }
        catch {
            #If Destination Site does not exist
            Write-Host "Unable to Find Destination OneDrive Site Provisioned for $($destinationUPN)." -foregroundcolor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "UnableToFindDestinationOneDrive" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Tenant" -Value $destinationTenant.Site -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $destinationUPN
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $AllOneDriveErrors += $currenterror
            Continue
        }

        #Get Destination OneDrive Site
        try {
            $dstSite = Connect-Site -Url $dstSiteUrl -UseCredentialsFrom $destinationTenant -ErrorAction Stop
            #Get OneDrive Documents
            $dstList = Get-List -Site $dstSite -Name "Documents" -ErrorAction Stop
            Write-Host "Connected." -ForegroundColor Green
        }
        catch {                
            #If Migration Service Account is not enabled. Add account as Secondary Admin
            Set-SPOUser -Site $dstSiteUrl.tostring() -LoginName $DestinationServiceAccount.tostring() -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue
            Write-Host "$($DestinationServiceAccount) Added as Site Admin. " -ForegroundColor Green -nonewline 

            while (!(Get-SPOUser -Site $dstSiteUrl.tostring() -ErrorAction SilentlyContinue)) {
                    Write-Host " ." -NoNewline -foregroundcolor yellow
                    Start-Sleep -s 3
                }
            #Attempt again to connect to Destination OneDrive
            $dstSite = Connect-Site -Url $dstSiteUrl -UseCredentialsFrom $destinationTenant -ErrorAction Stop
            #Get OneDrive Documents
            $dstList = Get-List -Site $dstSite -Name "Documents" -ErrorAction Stop
            Write-Host "Destination OneDrive Connected... " -ForegroundColor Green
            Continue 
        }
    }
    
    #Connect to OneDrive Sites in Destination
    if ($dstList -and $srcList) {
        $OneDriveResults = New-Object PSObject
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceAdminTenantURL" -Value $sourceTenant.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationAdminTenantURL" -Value $destinationTenant.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceName" -Value $srcSite.Title
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceSite" -Value $srcSite.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationName" -Value $dstSite.Title
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationURL" -Value $dstSite.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value "Incremental"

        #Copy OneDrive Files from Source to Destination
        #Progress Bar Current 2
        $TaskName = "Incremental OneDrive Migration for $($srcSite.Title) to $($dstSite.Title)"
        Write-progress -id 2 -Activity "$($TaskName)"
        Write-Host $TaskName -ForegroundColor Magenta
        
        #Test Move with Incremental using Insane Mode
        if ($Test) {
            $Result = Copy-Content -SourceList $srcList -DestinationList $dstList -InsaneMode -CopySettings $copysettings -TaskName $TaskName -warningaction silentlycontinue -whatif
        }
        #Migrate Data with Incremental using Insane Mode
        else {
            $Result = Copy-Content -SourceList $srcList -DestinationList $dstList -InsaneMode -CopySettings $copysettings -TaskName $TaskName -warningaction silentlycontinue
        }

        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result.Result
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "ItemsCopied" -Value $Result.ItemsCopied
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Successes" -Value $Result.Successes
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Errors" -Value $Result.Errors
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Warnings" -Value $Result.Warnings
        $AllOneDriveResults += $OneDriveResults
    }
    else {
        Write-Host "Unable to find Source or Destination OneDrive Site - $($destinationUPN) " -ForegroundColor Red
    }
    Write-Host "Completed" -ForegroundColor Green
    $OneDriveResults | Export-Csv "$HOME\Desktop\OneDrive-MigrationResults.csv" -NoTypeInformation -Encoding UTF8 -Append
    $AllOneDriveErrors | Export-Csv "$HOME\Desktop\OneDrive-MigrationErrors.csv" -NoTypeInformation -Encoding UTF8 -Append
    
}
Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan
Write-Host $($AllOneDriveErrors.Count) "Errors Occurred" -ForegroundColor Red
