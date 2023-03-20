## Get Public Folder Details
$PFServer = "425EXCHCL2"
$PublicFolders = Get-PublicFolder -Server $PFServer -Recurse
$progressref = $PublicFolders.count
$progresscounter = 0
$PublicFolderDetails = @()
foreach ($pf in $PublicFolders) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Public Folder Permissions $($pf.Identity)"
    $PFStatistics = Get-PublicFolderStatistics $pf.Identity -Server $PFServer
    
    $currentObject = New-Object PsObject
    $currentObject | Add-Member -type NoteProperty -Name "Name" -Value $pf.Name
    $currentObject | Add-Member -type NoteProperty -Name "Identity" -Value $pf.Identity
    $currentObject | Add-Member -type NoteProperty -Name "ParentPath" -Value $pf.ParentPath
    $currentObject | Add-Member -type NoteProperty -Name "MailEnabled" -Value $pf.MailEnabled
    $currentObject | Add-Member -type NoteProperty -Name "HasRules" -Value $pf.HasRules
    $currentObject | Add-Member -type NoteProperty -Name "HasSubFolders" -Value $pf.HasSubFolders
    $currentObject | Add-Member -type NoteProperty -Name "CreationTime" -Value $PFStatistics.CreationTime
    $currentObject | Add-Member -type NoteProperty -Name "LastUserModificationTime" -Value $PFStatistics.LastUserModificationTime
    $currentObject | Add-Member -type NoteProperty -Name "LastUserAccessTime" -Value $PFStatistics.LastUserAccessTime
    $currentObject | Add-Member -type NoteProperty -Name "ItemCount" -Value $PFStatistics.ItemCount
    $currentObject | Add-Member -type NoteProperty -Name "TotalItemSize" -Value $PFStatistics.TotalItemSize -Force
    $currentObject | Add-Member -type NoteProperty -Name "TotalItemSize-MB" -Value ([math]::Round(($PFStatistics.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)","")/1MB,3)) -Force
    $currentObject | Add-Member -type NoteProperty -Name "EntryID" -Value $PFStatistics.EntryID
    $PublicFolderDetails += $currentObject
}
$PublicFolderDetails | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath C:\Users\adamedrano\Desktop\Cenlar-AllPublicFolderDetails.csv 

## Get Public Folder Permisions
$PFServer = "425EXCHCL2"
$PublicFolders = Get-PublicFolder -Server $PFServer -Recurse
$progressref = $PublicFolders.count
$progresscounter = 0
$PFPermDetails = @()
foreach ($pf in $PublicFolders) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Public Folder Permissions $($pf.Identity)"

    $PFPermissions = Get-PublicFolderClientPermission $pf.Identity -Server 425EXCHCL2
    foreach ($perm in $PFPermissions) {
        $AccessRights = $perm.AccessRights -join ";"
        $currentPerm = New-Object PsObject
        $currentPerm | Add-Member -type NoteProperty -Name "Identity" -Value $perm.Identity
        $currentPerm | Add-Member -type NoteProperty -Name "User" -Value $perm.User
        if ($recipientCheck = Get-Recipient $perm.User -ErrorAction SilentlyContinue) {
            $currentPerm | Add-Member -type NoteProperty -Name "UserDisplayName" -Value $recipientCheck.DisplayName
            $currentPerm | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipientCheck.PrimarySMTPAddress
        }
        else {
            $currentPerm | Add-Member -type NoteProperty -Name "UserDisplayName" -Value $null
            $currentPerm | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $null
        }
        $currentPerm | Add-Member -type NoteProperty -Name "AccessRights" -Value $AccessRights
        $PFPermDetails += $currentPerm
    }
}
$PFPermDetails | Export-Csv C:\Users\adamedrano\Desktop\Cenlar-AllPublicFolderPermissions.csv -NoTypeInformation