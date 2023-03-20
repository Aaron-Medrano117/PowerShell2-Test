[System.Console]::ForegroundColor = [System.ConsoleColor]::White

clear-host

Import-module activedirectory

write-host

write-host This Script will Get the ObjectGUID for a user and convert

write-host it to the Immutuable ID for use in Office 365

Write-Host

write-host Please choose one of the following:

write-host

write-host ‘1) Get ID for a Single User’

write-host ‘2) Get IDs for all Users’

write-host ‘3) Cancel’ -ForegroundColor Red

write-host

$option = Read-Host “Select an option [1-3]”

switch ($option)

{

‘1’{

write-verbose “Option 1 selected”

$GetUser = Read-Host -Prompt ‘Enter UserName’

$Users = get-aduser $GetUser  | select samaccountname,userprincipalname,objectguid,@{label=”ImmutableID”;expression={[System.Convert]::ToBase64String($_.objectguid.ToByteArray())}}

$Users

}

‘2’{

Write-host

Write-host Type the Path location to Export the results:   i.e. c:\source\ImmutableID.csv

$Path = Read-Host -Prompt ‘Enter Path’

$Users = get-aduser -filter * | select samaccountname,userprincipalname,objectguid,@{label=”ImmutableID”;expression={[System.Convert]::ToBase64String($_.objectguid.ToByteArray())}}

$users

$users | export-csv $Path

}

‘3’{

write-verbose “Option 3 selected”

break

}

}