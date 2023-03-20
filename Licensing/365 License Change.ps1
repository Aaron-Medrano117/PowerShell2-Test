#This script is designed to gather update the existing MSOL Users Licenses. This update specifically updates all users with Office E3 licenses to Microsoft E3 licenses AND updates existing users with Office E5 licenses to Microsoft E5 licenses WITH the currently disabled services

#Gather All Msol Users in Tenant with License matching EnterprisePack (Office E3)

$AllUsers = Get-MsolUser -all | Where-Object {$_.islicensed} |  ? {$_.licenses.accountSKUID -match "ENTERPRISEPACK"}

#Gather List of all Disabled Services for E3 users, build into an Array, and set create as a new disabledplan
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
        #Update users with Office E3 licenses to Microsoft E3 licenses with DisabledArray above.
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId libertylife:SPE_E3 -DisabledPlans $DisabledArray
        Write-host "Updating E3 license for $user"
        
        try {
        if (!Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses libertylife:SPE_E3 -RemoveLicenses libertylife:ENTERPRISEPACK  -LicenseOptions $LicenseOptions -verbose)
            Write-Host "Completed For $user"
    
    } 
    
    #Gather All Msol Users in Tenant with License matching EnterprisePack (Office E5) 
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
        #Update users with Office E5 licenses to Microsoft E5 licenses with DisabledArray above.
        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId libertylife:SPE_E5 -DisabledPlans $DisabledArray
        Write-host "Updating E3 license for $user"
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses libertylife:SPE_E5 -RemoveLicenses libertylife:ENTERPRISEPREMIUM  -LicenseOptions $LicenseOptions -verbose
         }
    }
}
