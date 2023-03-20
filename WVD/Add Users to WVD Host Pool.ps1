Add Users to WVD Host Pool
#CSV can just be a list of users. Can also utilize a txt file or anything else for the required varialbe $HostPoolUsers

$HostPoolUsers = Import-Csv $home\desktop\HostPoolUsers.csv

#Get the WVD Tenant Name
$AllTenantNames = Get-RdsTenant | Select-Object TenantName, FriendlyName
$AllTenantNames
#### Want to add in a part for people to choose which Tenant Name

[string]$TenantName = # $AllTenantNames[0].TenantName

#Get the WVD Host Pool Name

$AllHostPoolNames = Get-RdsHostPool | Select-Object TenantName, HostPoolName, FriendlyName, Persistent
$AllHostPoolNames
[string]$HostPoolName = $AllHostPoolNames[0].HostPoolName

# Get-APPPoolName
Get-RdsAppGroup

foreach ($upn in $HostPoolUsers) 
{
    $adduser = add-RdsAppGroupUser -TenantName $TenantName -HostPoolName $HostPoolName -AppGroupName "Desktop Application Group" -UserPrincipalName $upn
    if ($adduser) 
    {
        Write-Host "Added $upn" -ForegroundColor green
        $tmp = "" | select User, HostPoolName, TenantName, Result
        $tmp.User = $upn
        $tmp.HostPoolName = $HostPoolName
        $tmp.TenantName = $TenantName
        $tmp.Result = "Successful"
        $HostPoolUsersResults += $tmp

    }
    else 
    {
        Write-Host "Unable to add $upn to host pool $hostpoolname" 
        $tmp = "" | select User, HostPoolName, TenantName, Result
        $tmp.User = $upn
        $tmp.HostPoolName = $HostPoolName
        $tmp.TenantName = $TenantName
        $tmp.Result = "Failed"
        $HostPoolUsersResults += $tmp   
    }
}

$HostPoolUsersResults | Export-Csv $home\HostPoolUsersResults.csv -NoTypeInformation
    
 #### Example Section ####
 # add-RdsAppGroupUser -TenantName eu2-wvd -HostPoolName EU2-PDP-EXEC -AppGroupName "Desktop Application Group" -UserPrincipalName $upn
 # Write-Host "Added $upn" -ForegroundColor green

[arry]$users = "mnichols@deephavenmortgage.com","crupp@deephavenmortgage.com","nyoussef@deephavenmortgage.com","mwitt@deephavenmortgage.com","rmullane@deephavenmortgage.com","msutton@deephavenmortgage.com","mdavlin@deephavenmortgage.com","cwalton@deephavenmortgage.com","mwalker@deephavenmortgage.com","dpatel@deephavenmortgage.com","mlehnen@deephavenmortgage.com","jmccray@deephavenmortgage.com","mwarren@deephavenmortgage.com"

foreach ($upn in $users) 
{
    ## Example
    add-RdsAppGroupUser -TenantName eu2-wvd -HostPoolName EU2-PDP-EXEC -AppGroupName "Desktop Application Group" -UserPrincipalName $upn
    Write-Host "Added $upn" -ForegroundColor green
}

### Example Section ###