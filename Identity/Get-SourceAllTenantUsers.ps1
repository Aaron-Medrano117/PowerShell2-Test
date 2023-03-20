function Get-SourceAllTenantUsers
{
    
    <#
	.PARAMETER FolderPath
		Folder Location Path.

	.EXAMPLE .\Get-SourceTenantUsers -FolderPath "C:\"
    #>

    
    Param(
    [Parameter(Mandatory=$False)]
        [string]$FolderPath
    )

    if($FolderPath)
    {
        If(!(Test-Path $FolderPath))
        {
            md -Path $FolderPath 
        }
        $CSVpath = "$($FolderPath)\365UsersReport.csv"
    }
    else
    {
        $CSVpath = "~\Desktop\365UsersReport.csv"
    }
  
    $AllMsoluUsers = Get-MsolUser -All

    $licensedSharedMailboxProperties = @()
  
    foreach ($user in $AllMsoluUsers) 
    {
        Write-Host "$($user.displayname)" -ForegroundColor Yellow  
        $licenses = $user.Licenses
        $licenseArray = $licenses | foreach-Object {$_.AccountSkuId}
        $licenseString = $licenseArray -join ", "
        $ProxyAddresses = $user.ProxyAddresses
        $ProxyAddressString = $ProxyAddresses -join ", "
        $UPNPrefix = ($user.UserPrincipalName -split "@")[0]
        $DisabledArray = @()
        $allLicenses = ($user).Licenses
        for($i = 0; $i -lt $AllLicenses.Count; $i++)
        {
            $serviceStatus =  $AllLicenses[$i].ServiceStatus
            foreach($service in $serviceStatus)
            {
                if($service.ProvisioningStatus -eq "Disabled")
                {
                    $disabledArray += ($service.ServicePlan).ServiceName
                }
            }
    
        }
        $disabledString = $DisabledArray -join ", "
        $licensedSharedMailboxProperties = [pscustomobject][ordered]@{
            DisplayName       = $user.DisplayName
            FirstName         = $user.FirstName
            LastName          = $user.LastName
            UserPrincipalName = $user.UserPrincipalName
            UPNPrefix         = $UPNPrefix
            EmailAddresses    = $ProxyAddressString
            Title             = $user.Title
            IsLicensed        = $user.islicensed    
            Licenses          = $licenseString
            DisabledPlans     = $disabledString
            SigninBlocked     = $user.BlockCredential
            StreetAddress     = $user.StreetAddress
            City              = $user.City
            State             = $user.State
            Country           = $user.Country
            UseageLocation    = $user.UsageLocation
            Office            = $user.Office
            PostalCode        = $user.PostalCode   
        
        }
        $licensedSharedMailboxProperties | Export-CSV -Path $CSVpath -Append -NoTypeInformation   
    }
}
