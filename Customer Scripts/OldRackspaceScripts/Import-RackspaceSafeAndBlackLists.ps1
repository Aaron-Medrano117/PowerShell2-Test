<# Need to put in a count check. There is a limit of the number of entries for IP and for senders/domains

Safe sender limit: 1024 entries
Blocked sender limit per policy: 1024 entries

IP Allow or IP Block list limit - When configuring an IP Allow list or an IP Block list in the connection filter, you can specify a maximum of 1273 entries
where an entry is either a single IP address or a CIDR range of IP addresses from /24 to /32.
https://docs.microsoft.com/en-us/office365/servicedescriptions/exchange-online-protection-service-description/exchange-online-protection-limits

#>

#Add Rackspace Outobund IPs for migration
$rackspaceOutboundIPs = Get-Content $HOME\Desktop\RackspaceOutboundIPs.txt
Set-HostedConnectionFilterPolicy Default -IPAllowList @{add=$rackspaceOutboundIPs}

#blacklist Add


$blackListDomainsAndUsers = Get-Content $HOME\Desktop\blacklist.csv


# Blocked domains
$valid = @()
$invalid = @()
foreach ($domain in $blackListDomainsAndUsers | Where {$_ -like "@*"})
{
	$domainName = $domain.Replace("@","").Replace("%","")
	
	while ($domainName -like ".*")
	{
		$domainName = $domainName.Substring(1)
	}
	
	if ($domainName -notlike "*.*")
	{
		$domainName = ".$domainName"
		Write-Host $domainName -ForegroundColor White
	}
	elseif ($domainName -match '^([a-z0-9]+(-[a-z0-9]+)*\.)+[a-z]{2,}$')
	{
		$valid += $domainName
		Write-Host $domainName -ForegroundColor Cyan
	}
	else
	{
		$invalid += $domain
		Write-Host $domainName -ForegroundColor Red
	}
}

Set-HostedContentFilterPolicy Default -BlockedSenderDomains @{add=$valid}


# Blocked senders

$blockedSenders = $blackListDomainsAndUsers.Trim("%") | Where {$_.Split("@")[0]}

Set-HostedContentFilterPolicy Default -BlockedSenders @{add=$blockedSenders}



# Blocked IPs

$blackListIPs = Get-Content $HOME\Desktop\blacklist_IPs.csv

$blackList = @()
$blackListIPs | foreach {
	
	$ip = $_
	if ($ip -like "*.%")
	{
		$ip = $_.Replace(".%",".0/24")
	}
	
	$blackList += $ip
}

Set-HostedConnectionFilterPolicy Default -IPBlockList @{add=$blackList}



# Safe lists

$whiteListIPs = Get-Content $HOME\Desktop\safelistIPs.csv

$safeListIPs = @()
$whiteListIPs | foreach {
	
	$ip = $_
	if ($ip -like "*.%")
	{
		$ip = $_.Replace(".%",".0/24")
	}
	
	$safeListIPs += $ip
}

Set-HostedConnectionFilterPolicy Default -IPAllowList @{add=$safeListIPs}




$safeListDomainsAndSenders = Get-Content $HOME\Desktop\safelistSendersDomains.csv

$valid = @()
$invalid = @()
foreach ($domain in $safeListDomainsAndSenders | Where {$_ -like "@*"})
{
	$domainName = $domain.Replace("@","").Replace("%","")
	
	while ($domainName -like ".*")
	{
		$domainName = $domainName.Substring(1)
	}
	
	if ($domainName -notlike "*.*")
	{
		$domainName = ".$domainName"
		Write-Host $domainName -ForegroundColor White
	}
	elseif ($domainName -match '^([a-z0-9]+(-[a-z0-9]+)*\.)+[a-z]{2,}$')
	{
		$valid += $domainName
		Write-Host $domainName -ForegroundColor Cyan
	}
	else
	{
		$invalid += $domain
		Write-Host $domainName -ForegroundColor Red
	}
}

Set-HostedContentFilterPolicy Default -AllowedSenderDomains @{add=$valid}

## Allowed Senders
$allowedSenders = $safeListDomainsAndSenders.Trim("%") | Where {$_.Split("@")[0]}


Set-HostedContentFilterPolicy Default -AllowedSenders @{add=$allowedSenders}

######
Enable-OrganizationCustomization
$domainslist = get-content "C:\Users\fred5646\Rackspace Inc\LA Truck Center - MPS-TS - General\vvg-safelist.csv"
$domainslist | foreach {Set-HostedContentFilterPolicy -Identity default -AllowedSenderDomains @{add=$_} -whatif}

$allowedsenders = Get-Content "C:\Users\fred5646\Rackspace Inc\LA Truck Center - MPS-TS - General\allowedsenders.txt"
$allowedsenders | foreach {Set-HostedContentFilterPolicy -Identity default -AllowedSenders @{add=$_} -whatif}



$safeListDomainsAndSenders = Get-Content $HOME\Desktop\safelistSendersDomains.csv

$valid = @()
$invalid = @()
foreach ($domain in $safeListDomainsAndSenders | Where {$_ -like "@*"})
{
	$domainName = $domain.Replace("@","").Replace("%","")
	
	while ($domainName -like ".*")
	{
		$domainName = $domainName.Substring(1)
	}
	
	if ($domainName -notlike "*.*")
	{
		$domainName = ".$domainName"
		Write-Host $domainName -ForegroundColor White
	}
	elseif ($domainName -match '^([a-z0-9]+(-[a-z0-9]+)*\.)+[a-z]{2,}$')
	{
		$valid += $domainName
		Write-Host $domainName -ForegroundColor Cyan
	}
	else
	{
		$invalid += $domain
		Write-Host $domainName -ForegroundColor Red
	}
}