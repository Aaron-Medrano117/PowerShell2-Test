foreach ($user in $bs_users)
{
    Write-Host "Adding ATP for $($user.DisplayName) ..." -ForegroundColor Cyan -NoNewline
    Set-MsolUserLicense -userprincipalname $user.userprincipalname -addlicenses accttwo:ATP_ENTERPRISE
    Write-Host "done" -ForegroundColor Green
}


$childers_users = Import-Csv 
foreach ($user in ($childers_users | ? {$_.licensing -eq "doublekwik:SPB"}))
{
    Write-Host "Adding $($user.licensing) for $($user.userprincipalname) ..." -ForegroundColor Cyan -NoNewline
    Set-MsolUserLicense -userprincipalname $user.userprincipalname -addlicenses doublekwik:SPB
    Write-Host "done" -ForegroundColor Green
}