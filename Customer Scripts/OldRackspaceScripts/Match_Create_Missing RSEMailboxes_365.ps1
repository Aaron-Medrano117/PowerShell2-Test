$RSEImport = Import-Csv "C:\Users\fred5646\Rackspace Inc\American Car Center - General\CP RSE email-mailboxes-20200511-181813.csv"

    $found = @()
    $notfound = @()

    foreach ($user in $RSEImport) {
        if (get-msoluser -searchstring $user.name) {
            $found += $user
        }
        else
        {
            $notfound += $user
        }
    }

    #### Create MSOLUSers

    foreach ($user in $notfound) {
        Write-Host "Creating User "$user.Name" ..." -ForegroundColor Cyan -NoNewline
        if ($user.DisplayName) {
            New-MsolUser -userprincipalname $user.email -displayname $user.DisplayName -FirstName $user.FirstName -LastName $user.LastName -PhoneNumber $user.BusinessPhone -password 'Vi6$q3tr5FMT'
            Write-Host "done" -ForegroundColor Green
        }
       else
       {
           Write-Host "Required Field missing. No DisplayName found" -ForegroundColor red
       }
        
    }


    New-MsolUser -userprincipalname christian.mason@americancarcenter.com -displayname "Christian Mason" -FirstName "Christian" -LastName "Mason" -password 'Vi6$q3tr5FMT'