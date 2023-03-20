#Gather Admins with Migrating Domain(s) in 365
Connect-AzureAD
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
    #$secondsElapsed = (Get-Date) – $global:start
    $global:secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($global:progressref – $progresscounter)
}

$roles = Get-AzureADDirectoryRole
$MigratingDomain = Read-Host "What is the Migrating Domain"
$Admins = @()
#$global:progressref = ($roles).count
$progresscounter = 1
$global:start = Get-Date
[nullable[double]]$global:secondsRemaining = $null
foreach ($role in $roles) {
    Write-ProgressHelper -ProgressCounter ($progresscounter++) -Activity "Gathering $($role.DisplayName)" -ID 1 -TotalCount ($roles).count
    $allAdmins = @()
    if ($MigratingDomain) {
        $allAdmins = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId | ? {$_.UserPrincipalName -like "*$MigratingDomain"}
    }
    else {
        $allAdmins = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId
    }
    $global:progressref = ($roles).count
    foreach ($admin in $allAdmins) {
        Write-ProgressHelper -Activity "Details for $($admin.DisplayName)" -ProgressCounter ($progresscounter++) -ID 2 -TotalCount ($allAdmins).count
        $currentAdmin = new-object PSObject
        $currentAdmin | add-member -type noteproperty -name "Role" -Value $Role.DisplayName
        $currentAdmin | Add-Member -type NoteProperty -Name "Role_Description" -Value $Role.Description -Force
        $currentAdmin | Add-Member -type NoteProperty -Name "DisplayName" -Value $admin.DisplayName -Force
        $currentAdmin | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $admin.UserPrincipalName -Force
        $currentAdmin | Add-Member -type NoteProperty -Name "UserType" -Value $admin.UserType -Force
        $Admins += $currentAdmin
    }
}


#admin check - old school progress
$roles = Get-AzureADDirectoryRole
$MigratingDomain = Read-Host "What is the Migrating Domain"
$Admins = @()
$progressref = ($roles).count
$progresscounter = 0
foreach ($role in $roles) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering $($role.DisplayName)"

    $allAdmins = @()
    if ($MigratingDomain) {
        $allAdmins = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId | ? {$_.UserPrincipalName -like "*$MigratingDomain"}
    }
    else {
        $allAdmins = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId
    }
    $progressref2 = ($allAdmins).count
    $progresscounter2 = 0
    foreach ($admin in $allAdmins) {
        $progresscounter2 += 1
        $progresspercentcomplete2 = [math]::Round((($progresscounter2 / $progressref2)*100),2)
        $progressStatus2 = "["+$progresscounter2+" / "+$progressref2+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete2 -Status $progressStatus2 -Activity "Checking User $($admin.DisplayName)"
        $currentAdmin = new-object PSObject
        $currentAdmin | add-member -type noteproperty -name "Role" -Value $Role.DisplayName
        $currentAdmin | Add-Member -type NoteProperty -Name "Role_Description" -Value $Role.Description -Force
        $currentAdmin | Add-Member -type NoteProperty -Name "DisplayName" -Value $admin.DisplayName -Force
        $currentAdmin | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $admin.UserPrincipalName -Force
        $currentAdmin | Add-Member -type NoteProperty -Name "UserType" -Value $admin.UserType -Force
        $Admins += $currentAdmin
    }
}