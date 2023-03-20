#PW change

$childersusers = Import-Csv "C:\Users\fred5646\OneDrive - Rackspace Inc\Desktop\Childers Oil\Smartermail Users.csv"

$failedusers = @()
$completedusers = @()
foreach ($msoluser in ($childersusers | ? {$_.username -ne "rfannin" -and $_.licensing}))
{
    $useprincipalname = $msoluser.username + "@doublekwik.com"
    $PW = $msoluser.CurrentPW
    Write-Host "Updating Password for $($useprincipalname) ..." -NoNewline -ForegroundColor Cyan
    if ($SetPW = Set-MsolUserPassword -UserPrincipalName $useprincipalname  -NewPassword $PW -ForceChangePassword $false)
    {
        $SetPW
        $completedusers += $msoluser
    }

    else
    {
        $failedusers += $msoluser
    } 
}


$tmppw = "Re5etP@55!"
$failedusers2 = @()
$completedusers2 = @()

foreach ($msoluser in $failedusers)
{
    $useprincipalname = $msoluser.username + "@doublekwik.com"
    Write-Host "Updating Password for $($useprincipalname) ..." -NoNewline -ForegroundColor Cyan
    if ($SetPW = Set-MsolUserPassword -UserPrincipalName $useprincipalname  -NewPassword $tmppw -ForceChangePassword $false)
    {
        $SetPW
        $completedusers2 += $msoluser
    }

    else
    {
        $failedusers2 += $msoluser
    } 
}
