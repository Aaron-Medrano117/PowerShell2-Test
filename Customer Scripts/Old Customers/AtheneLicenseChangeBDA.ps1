$AllUsers = Get-MsolUser -userprincipal rmailhot@athene.bm | Where-Object {$_.islicensed}

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

        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId athenebda:SPE_E3 -DisabledPlans $DisabledArray -Verbose
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses athenebda:SPE_E3 -LicenseOptions $LicenseOptions -RemoveLicenses athenebda:ENTERPRISEPACK -verbose

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

        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId athenebda:SPE_E5 -DisabledPlans $DisabledArray
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses athenebda:SPE_E5 -RemoveLicenses athenebda:ENTERPRISEPREMIUM  -LicenseOptions $LicenseOptions -verbose
    }
}