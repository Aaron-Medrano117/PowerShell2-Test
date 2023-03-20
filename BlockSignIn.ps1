#Block users IN Azure AD
Import-Csv "C:\Testblock.csv" | ForEach-Object {
>>   Set-AzureADUser -ObjectId $($_.ObjectId) -AccountEnabled $false
>> }
 
#UnbLock users in Azure AD
Import-Csv "C:\Testblock.csv" | ForEach-Object {
>>   Set-AzureADUser -ObjectId $($_.ObjectId) -AccountEnabled $true
>> }

<## START - Block users IN Azure AD - 2 ##>
## Import Users
$BlockedUsers = Import-Csv "C:\Testblock.csv"

#Create Output Variables
$AllErrors = @()
$CompletedUsers = @()

#Progress Bar 1A
$progressref = ($BlockedUsers).count
$progresscounter = 0

foreach ($user in $BlockedUsers) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Blocking Access for $($user.DisplayName)"
    
    #Block Sign In Based on user's Object ID
    #Get Azure Object ID Details
    Write-Host "Checking for $($user.DisplayName)" -ForegroundColor Cyan -NoNewline
    if ($azureUser = Get-AzureADUser -SearchString $user.UserPrincipalName) {
        try {
            
            Set-AzureADUser -ObjectId $azureUser.ObjectId -AccountEnabled $false -ErrorAction Stop
            $CompletedUsers += $azureUser
            Write-Host "User Blocked" -ForegroundColor Green 
        }
        #If Failed to Block Sign in, Create Error Report
        catch {
            Write-Host "Failed to Block Access for $($azureUser.DisplayName)" -ForegroundColor Red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "BlockUser" -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $azureUser.DisplayName -Force
            $currenterror | Add-Member -type NoteProperty -Name "ObjectID" -Value $azureUser.ObjectId -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $AllErrors += $currenterror
            continue
        }
    }
    else {
        Write-Host "Unable to find $($user.DisplayName)" -ForegroundColor Red
        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FindUser" -Force
        $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.UserPrincipalName -Force
        $currenterror | Add-Member -type NoteProperty -Name "ObjectID" -Value $null -Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllErrors += $currenterror
        continue
    }
    
}
#Result Output
Write-Host $CompletedUsers.count "Users Blocked Successfully" -ForegroundColor Green
Write-Host $AllErrors.count "Users Failed to Block" -ForegroundColor Yellow

#Export Results to Desktop
$CompletedUsers | Export-Csv "$HOME\Desktop\BlockSignIn-CompletedUsers.csv" -NoTypeInformation -Encoding UTF8 -Append
$AllErrors | Export-Csv "$HOME\Desktop\BlockSignIn-FailedUsers.csv" -NoTypeInformation -Encoding UTF8 -Append

#Export Location Details
Write-Host "Exported Successful Users to Desktop: BlockSignIn-CompletedUsers.csv" -ForegroundColor Green
Write-Host "Exported Failed Users to Desktop: BlockSignIn-FailedUsers.csv" -ForegroundColor Yellow
<## END - Block users IN Azure AD - 2 ##>

<## START - UnBlock users in Azure AD - 2 ##>

## Import Users
$BlockedUsers = Import-Csv "C:\Testblock.csv"

#Create Output Variables
$AllErrors = @()
$CompletedUsers = @()

#Progress Bar 1A
$progressref = ($BlockedUsers).count
$progresscounter = 0

foreach ($user in $BlockedUsers) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Unblock Access for $($user.DisplayName)"
    
    #unBlock Sign In Based on user's Object ID
    try {
        Set-AzureADUser -ObjectId $user.ObjectId -AccountEnabled $true -ErrorAction Stop
        $CompletedUsers += $user
    }
    #If Failed to unBlock Sign in, Create Error Report
    catch {
        Write-Host "Failed to UnBlock Access for $($user.DisplayName)" -ForegroundColor Red
        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "UnBlockUser" -Force
        $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.DisplayName -Force
        $currenterror | Add-Member -type NoteProperty -Name "ObjectID" -Value $user.ObjectId -Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
        $AllErrors += $currenterror
        continue
    }
}
#Result Output
Write-Host $CompletedUsers.count "Users UnBlocked Successfully" -ForegroundColor Green
Write-Host $AllErrors.count "Users Failed to UnBlock" -ForegroundColor Yellow

#Export Results to Desktop
$CompletedUsers | Export-Csv "$HOME\Desktop\UnBlockSignIn-CompletedUsers.csv" -NoTypeInformation -Encoding UTF8 -Append
$AllErrors | Export-Csv "$HOME\Desktop\UnBlockSignIn-FailedUsers.csv" -NoTypeInformation -Encoding UTF8 -Append

#Export Location Details
Write-Host "Exported Successful Users to Desktop: UnBlockSignIn-CompletedUsers.csv" -ForegroundColor Green
Write-Host "Exported Failed Users to Desktop: UnBlockSignIn-FailedUsers.csv" -ForegroundColor Yellow

<## END - UnBlock users in Azure AD - 2 ##>