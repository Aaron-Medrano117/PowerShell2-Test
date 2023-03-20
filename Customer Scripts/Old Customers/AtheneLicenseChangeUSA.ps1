$AllUsers = Get-MsolUser -all | Where-Object {$_.islicensed} |  ? {$_.licenses.accountSKUID -match "EnterprisePremium"}

foreach($user in $allusers)
{
    if($user.licenses.accountSKUID -match "EnterprisePACK")
    {
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

        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId libertylife:SPE_E3 -DisabledPlans $DisabledArray
        #Write-host "Updating E3 license for $user"
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses libertylife:SPE_E3 -RemoveLicenses libertylife:ENTERPRISEPACK  -LicenseOptions $LicenseOptions -verbose
    }
    elseif($user.licenses.accountSKUID -match "EnterprisePremium")
    {
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

        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId libertylife:SPE_E5 -DisabledPlans $DisabledArray
        #Write-host "Updating E5 license for $user"
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses libertylife:SPE_E5 -RemoveLicenses libertylife:ENTERPRISEPREMIUM  -LicenseOptions $LicenseOptions -verbose
    }
}