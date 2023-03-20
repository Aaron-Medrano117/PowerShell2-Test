$importcsv = import-csv

$webmailusers = @()

foreach ($user in $importcsv | sort Displayname)
{
    $tmp = "" | select Displayname, DistinguishedName, EmailAddress, ObjectClass, ObjectGUID, UserPrincipalName, ImmutableID
    $tmp.Displayname    = $user.Displayname
    $tmp.DistinguishedName = $user.DistinguishedName
    $tmp.EmailAddress = $user.EmailAddress
    $tmp.ObjectClass = $user.ObjectClass
    $tmp.ObjectGUID = $user.ObjectGUID
    $tmp.UserPrincipalName = $user.UserPrincipalName

    #create immutableID
    $UserimmutableID = [System.Convert]::ToBase64String(([GUID]$user.ObjectGUID).ToByteArray())
    $tmp.ImmutableID = $UserimmutableID
    $webmailusers += $tmp
}

## Hosted Exchange
foreach ($user in $importcsv | sort Displayname)
{
    $tmp = "" | select Displayname, DistinguishedName, EmailAddress, ObjectClass, ObjectGUID, UserPrincipalName, ImmutableID
    $tmp.Displayname    = $user.Displayname
    $tmp.DistinguishedName = $user.DistinguishedName
    $tmp.EmailAddress = $user.EmailAddress
    $tmp.ObjectClass = $user.ObjectClass
    $tmp.ObjectGUID = $user.ObjectGUID
    $tmp.UserPrincipalName = $user.UserPrincipalName

    #create immutableID
    $UserimmutableID = [System.Convert]::ToBase64String(([GUID]$user.ObjectGUID).ToByteArray())
    $tmp.ImmutableID = $UserimmutableID
    $webmailusers += $tmp
}

