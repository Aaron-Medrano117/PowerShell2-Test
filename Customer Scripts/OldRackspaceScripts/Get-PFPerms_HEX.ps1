
#
# This is for MEX06/8/9 only
#

$domain = #

$pfRootPath = "\" + $domain.ToCharArray()[0].ToString() + "\$domain"
$PFs = Get-PublicFolder $pfRootPath -Recurse


$permsList = @()
foreach ($pf in $PFs)
{
	Write-Host "$($pf.Identity) ..." -ForegroundColor Cyan -NoNewline
	
	$perms = $pf | Get-PublicFolderClientPermission | Where {$_.User.ToString() -ne "qbblrqzdd_pfread" -and $_.User -notlike "svc_*" -and $_.User -notlike "MEX0?*"}
	
	if ($perms)
	{
		foreach ($perm in $perms)
		{
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select Path, User, AccessRights
			$tmp.Path = $pf.Identity.ToString()
			$tmp.User = $perm.User.ToString()
			$tmp.AccessRights = $perm.AccessRights -join ","
			$permsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green
	}
	else
	{
		Write-Host " done" -ForegroundColor DarkCyan
	}
}

$permsList | Export-Csv "$HOME\Desktop\Perms_PF.csv" -NoTypeInformation -Encoding UTF8


#Region Tenant to Tenant

$PFs = Get-PublicFolder -Recurse -ResultSize unlimited


$pfPermsList = @()
foreach ($pf in $PFs)
{
	Write-Host "$($pf.Identity) ..." -ForegroundColor Cyan -NoNewline
	
	$perms = Get-PublicFolderClientPermission $pf.Identity
	
	if ($perms)
	{
		foreach ($perm in $perms)
		{
			Write-Host "." -ForegroundColor Yellow -NoNewline
			$tmp = "" | select Path, User, AccessRights
			$tmp.Path = $pf.Identity.ToString()
			$tmp.User = $perm.User.ToString()
			$tmp.AccessRights = $perm.AccessRights -join ","
			$pfPermsList += $tmp
		}
		
		Write-Host " done" -ForegroundColor Green
	}
	else
	{
		Write-Host " done" -ForegroundColor DarkCyan
	}
}

$pfPermsList | Export-Csv -NoTypeInformation -Encoding UTF8 "$HOME\Desktop\Perms_PF.csv"
