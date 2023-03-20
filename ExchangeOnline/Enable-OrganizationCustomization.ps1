    Enable-OrganizationCustomization
    $domainslist = Import-Csv "C:\Users\fred5646\Rackspace Inc\LA Truck Center - MPS-TS - General\vvg-safelist.csv"
    $domainslist | foreach {Set-HostedContentFilterPolicy -Identity default -AllowedSenderDomains @{add=$_.domain} -whatif}

    $domainslist = Get-Content "C:\Users\fred5646\Rackspace Inc\LA Truck Center - MPS-TS - General\vvg-safelist.csv"
    $domainslist | foreach {Set-HostedContentFilterPolicy -Identity default -AllowedSenderDomains @{add=$_} -whatif}

    $allowedsenders = Get-Content "C:\Users\fred5646\Rackspace Inc\LA Truck Center - MPS-TS - General\allowedsenders.txt"
    $allowedsenders | foreach {Set-HostedContentFilterPolicy -Identity default -AllowedSenders @{add=$_} -whatif}

    $allowedIPs = Get-Content "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\permit_ip_list.txt"
    $allowedIPs | foreach {Set-HostedConnectionFilterPolicy  -Identity default -IPAllowList @{add=$_} -whatif}

    $rackspaceIPs = Get-Content "C:\Users\fred5646\Rackspace Inc\MPS-TS-Fan Duel - General\External_Share\permit_ip_list.txt"
    $rackspaceIPs | foreach {Set-HostedConnectionFilterPolicy  -Identity default -IPAllowList @{remove=$_} -whatif}

    Set-HostedConnectionFilterPolicy  -Identity default -IPBlockList @{add=$blockedItem.Value}

foreach ($blockedItem in $allBlocklists)
{
    if ($blockedItem.BlocklisttType -eq "IP")
    {
        Set-HostedConnectionFilterPolicy  -Identity default -IPBlockList @{add=$blockedItem.Value}
    }
    if ($blockedItem.BlocklisttType -eq "hostname")
    {
        Set-HostedContentFilterPolicy -Identity default -BlockedSenderDomains @{add=$blockedItem.Value}
    }
    if ($blockedItem.BlocklisttType -eq "address")
    {
       Set-HostedContentFilterPolicy -Identity default -BlockedSenders @{add=$blockedItem.Value} 
    }
}

foreach ($allowedItem in $allSafelists)
{
    if ($allowedItem.SafelistType -eq "IP")
    {
        Write-Host "Adding $($allowedItem.Value) to Allow List"
        Set-HostedConnectionFilterPolicy  -Identity default -IPAllowList @{add=$allowedItem.Value}
    }
    if ($allowedItem.SafelistType -eq "hostname")
    {
        Write-Host "Adding $($allowedItem.Value) to Allow List"
        Set-HostedContentFilterPolicy -Identity default -AllowedSenderDomains @{add=$allowedItem.Value}
    }
    if ($allowedItem.SafelistType -eq "address")
    {
        Write-Host "Adding $($allowedItem.Value) to Allow List"
       Set-HostedContentFilterPolicy -Identity default -AllowedSenders @{add=$allowedItem.Value} 
    }
}

# Combined update spam settings
foreach ($spamSetting in $allSpamLists)
{
    if ($spamSetting.spamSetting -eq "Safelist")
    {
        if ($spamSetting.SpamSettingType -eq "IP")
        {
            Set-HostedConnectionFilterPolicy  -Identity default -IPAllowList @{add=$spamSetting.Value}
        }
        if ($spamSetting.SpamSettingType -eq "hostname")
        {
            Set-HostedContentFilterPolicy -Identity default -AllowedSenderDomains @{add=$spamSetting.Value}
        }
        if ($spamSetting.SpamSettingType -eq "address")
        {
        Set-HostedContentFilterPolicy -Identity default -AllowedSenders @{add=$spamSetting.Value} 
        }
    }
    elseif ($spamSetting.spamSetting -eq "Blacklist")
    {
       if ($spamSetting.SpamSettingType -eq "IP")
        {
            Set-HostedConnectionFilterPolicy  -Identity default -IPBlockList @{add=$spamSetting.Value}
        }
        if ($spamSetting.SpamSettingType -eq "hostname")
        {
            Set-HostedContentFilterPolicy -Identity default -BlockedSenderDomains @{add=$spamSetting.Value}
        }
        if ($spamSetting.SpamSettingType -eq "address")
        {
        Set-HostedContentFilterPolicy -Identity default -BlockedSenders @{add=$spamSetting.Value} 
        } 
    } 
}