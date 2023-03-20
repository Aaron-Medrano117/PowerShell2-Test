#
# This function looks up the TTL values for a customer's MX and Autodiscover records
#

function Query-TTLs
{
	[CmdletBinding()]
    param   (
                [string]
                $Domain
            )
			
	if (!$Domain)
	{
		Write-Host "Enter domain name: " -ForegroundColor Cyan -NoNewline
		$Domain = (Read-Host).Trim()
	}

	if ($nameHost = Resolve-DnsName -Type NS -Name $Domain -QuickTimeout | select -First 1 -ExpandProperty NameHost)
	{
		try
		{
			$autodiscoverTTLseconds = Resolve-DnsName -Type CNAME -Server $nameHost -Name ("autodiscover." + "$Domain") -NoHostsFile -QuickTimeout -EA Stop | select -First 1 -ExpandProperty TTL
			$mxTTLseconds = Resolve-DnsName -Type MX -Server $nameHost -Name $Domain -NoHostsFile -QuickTimeout -EA Stop | select -First 1 -ExpandProperty TTL
		}
		catch
		{
			$noNameServer = $true
			$autodiscoverTTLseconds = Resolve-DnsName -Type CNAME -Name ("autodiscover." + "$Domain") -NoHostsFile -QuickTimeout | select -First 1 -ExpandProperty TTL
			$mxTTLseconds = Resolve-DnsName -Type MX -Name $Domain -NoHostsFile -QuickTimeout | select -First 1 -ExpandProperty TTL
		}
	}
	else
	{
		$noNameServer = $true
		$autodiscoverTTLseconds = Resolve-DnsName -Type CNAME -Name ("autodiscover." + "$Domain") -NoHostsFile -QuickTimeout | select -First 1 -ExpandProperty TTL
		$mxTTLseconds = Resolve-DnsName -Type MX -Name $Domain -NoHostsFile -QuickTimeout | select -First 1 -ExpandProperty TTL
	}
	
	$autodiscoverTTLmins = $autodiscoverTTLseconds / 60
	$mxTTLmins = $mxTTLseconds / 60
	
	Write-Host
	Write-Host "Domain: " -ForegroundColor Cyan -NoNewline
	Write-Host $Domain -ForegroundColor Yellow
	Write-Host "MX TTL is " -ForegroundColor Cyan -NoNewline
	Write-Host "$mxTTLmins minutes" -ForegroundColor Green
	Write-Host "Autodiscover TTL is " -ForegroundColor Cyan -NoNewline
	Write-Host "$autodiscoverTTLmins minutes" -ForegroundColor Green
	
	if ($noNameServer)
	{
		Write-Host "Could not use name server" -ForegroundColor DarkGray
	}
	
	Write-Host
}