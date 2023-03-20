#Carlos Rosario
### Create Missing Carlos Rosario users
foreach ($user in $missingusers)
{
    $DisplayName = $user.DisplayName
    Write-Host "Checking for $($DisplayName) .." -NoNewline
    if (get-msoluser -userprincipalname $user.PrimarySmtpAddress -ea silentlycontinue)
    {
        Write-Host "$($DisplayName) user Exists. Skipping"
    }
    else
    {
        Write-Host "Creating MSOLUser for $($DisplayName) with UPN $($user.PrimarySmtpAddress.ToString()) " -ForegroundColor Cyan
        New-MsolUser -UserPrincipalName $user.PrimarySmtpAddress.ToString() -Usagelocation "US" -FirstName $user.FirstName -LastName $user.LastName -DisplayName $DisplayName -Password zA5ZwjKGK7FZ
    }
}