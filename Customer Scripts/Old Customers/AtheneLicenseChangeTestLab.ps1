$AllUsers = Get-MsolUser -all | Where-Object {$_.islicensed}

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

        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId onetestlab:SPE_E3 -DisabledPlans $DisabledArray
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses onetestlab:SPE_E3 -RemoveLicenses onetestlab:ENTERPRISEPACK  -LicenseOptions $LicenseOptions -verbose
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

        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId onetestlab:SPE_E5 -DisabledPlans $DisabledArray
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses onetestlab:SPE_E5 -RemoveLicenses onetestlab:ENTERPRISEPREMIUM  -LicenseOptions $LicenseOptions -verbose
    }
}