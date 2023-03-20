### Create new ad user

$UsersCSV = Import-Csv C:\Users\fred5646\Downloads\Import_User_Template.csv

$createdusers = @()
$usersalreadyexist = @()

foreach ($user in $UsersCSV)
{
    $username = $user.username
    $displayname = $user."Display name"
    if (!($msoluser = Get-MsolUser -userprincipalname $username -ea silentlycontinue))
    {
        New-MsolUser -userprincipalname $username -displayname $displayname -PreferredDataLocation US
        Write-Host "Created New user $($username)" -ForegroundColor Cyan -NoNewline
        $createdusers += $user
    }
    else
    {
        Write-Host "User already exists for $($displayname)" -ForegroundColor Green
        $usersalreadyexist += $user
    }
}