$createdusers =@()
$notcreatedusers = @()

foreach ($user in ($importcsv | ? {$_.ExistsOnPrem -eq $false}))
{
    $DisplayName = $user.DisplayName
    $pw = $user.pw
    $NewUPN = $user.NewUPN

    if (!(get-aduser -Filter {UserPrincipalName -eq $NewUPN}))
    {
        Write-Host "Creating new user $DisplayName ..." -ForegroundColor Cyan -NoNewline
        New-ADUser -Name $DisplayName -AccountPassword (ConvertTo-SecureString $pw -AsPlainText -force) -Path $user.DesiredOU -UserPrincipalName $user.NewUPN -ErrorVariable $CreateFail -WhatIf
        $createdusers += $DisplayName
    }
    
    else
    {
        Write-Host "User already created"
        $notcreatedusers += $DisplayName        
    }
}


### Update UPN

$UpdateUsersCSV = Import-Csv #filepath


$updatedusers =@()
$notupdatedusers = @()

foreach ($user in ($UpdateUsersCSV | ? {$_.ADUPN -like "*showrig.net"}))
{
    $DisplayName = $user.DisplayName
    Write-Host "Updating UPN for user $DisplayName ..." -ForegroundColor Cyan -NoNewline

    $ADUPN = $user.ADUPN
    $DistinguishedName = $user.DistinguishedName
    $NewUPN = $user.NewUPN

    if (get-aduser -Filter {UserPrincipalName -eq $ADUPN})
    {
        if ($UPNUpdateUser = Set-ADUser -Identity $DistinguishedName -UserPrincipalName $NewUPN -ErrorAction SilentlyContinue) #-WhatIf)
        {
            $UPNUpdateUser
            $updatedusers += $NewUPN
            Write-Host "updated" -ForegroundColor green
        }

        else
        {
            Write-Host "Couldn't update user" -ForegroundColor red
            $notupdatedusers += $NewUPN
        }
        
    }
    
    else
    {
        Write-Host "User not found" -ForegroundColor Yellow       
    }
}