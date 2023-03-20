[CmdletBinding()]
param(
    [switch]$StampAccounts,
    [switch]$Convertonly
)

$users = import-csv .\AD_Users.csv

#$users = $importdata
$msolAccounts = @()
$msolAccountsnotfound = @()
function isGUID ($data) {
    try {
        $guid = [GUID]$data
        return 1
    } catch {
        #$notguid = 1
        return 0
    }
}
function isBase64 ($data) {
    try {
        $decodedII = [system.convert]::frombase64string($data)
        return 1
    } catch {
        return 0
    }
}
function get-convertedID
{
    param([string]$valuetoconvert)
    if ($valuetoconvert -eq $NULL) {
        DisplayHelp
        return
    }
    if (isGUID($valuetoconvert))
    {
        $guid = [GUID]$valuetoconvert
        $bytearray = $guid.tobytearray()
        $immutableID = [system.convert]::ToBase64String($bytearray)
        return ($immutableID)
    } elseif (isBase64($valuetoconvert)){
        $decodedII = [system.convert]::frombase64string($valuetoconvert)
        if (isGUID($decodedII)) {
            $decode = [GUID]$decodedii
            $decode
        } else {
            Write-Host "Value provided not in GUID or ImmutableID format."
            DisplayHelp
        }
    } else {
        Write-Host "Value provided not in GUID or ImmutableID format."
        DisplayHelp
    }
    #$immutableID = "ci2LdGtw+EKLJYL9hzOGDw=="
    #$decodedII = [system.convert]::frombase64string($immutableID)
    #$decode = [GUID]$decodedii
}
$msolAccounts = @()

if($Convertonly)
{
    Write-host "> Converting ids now"
    foreach($user in $users)
    {
        
        if($user.UserPrincipalName)
        {
            #Write-host "UserPrincipalName" $user.EmailAddress
            #Write-host "ObjectID" $user.objectguid
            $converted = get-convertedID -valuetoconvert $user.objectguid
            #Write-host "ImmutableID" $converted

            $msolAccount = New-Object System.Object
            $msolAccount | Add-Member -MemberType NoteProperty -Name HEXDisplayName -Value $user.DisplayName
            $msolAccount | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $user.UserPrincipalName
            $msolAccount | Add-Member -MemberType NoteProperty -Name ObjectGUID -Value $user.objectguid
            $msolAccount | Add-Member -MemberType NoteProperty -Name ImmutableID -Value $converted
            $msolAccount | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddress -Value $user.PrimarySMTPAddress
            $msolAccounts += $msolAccount
        }
    }

$msolAccounts | sort DisplayName |  export-csv .\converted_ImmutableID.csv -NoTypeInformation -Force -Confirm:$false
}

if($StampAccounts)
{
    Write-host "> Checking Users and Stamping"
    $msolAccounts = import-csv .\converted_ImmutableID.csv
    if($msolAccounts)
    {
        foreach($msolUser in $msolAccounts)
        {

            if((get-msoluser -SearchString $msolUser.userprincipalname))
            {
                Write-host -ForegroundColor green "Found"
                Write-host "> Found User: " $msolUser.userprincipalname
                Write-host "> Setting Immutable ID: " $msolUser.ImmutableID
                if($msoluser.userprincipalname)
                {
                    $myoutput = get-msoluser -SearchString $msoluser.userprincipalname -ErrorAction SilentlyContinue | Set-MsolUser -ImmutableId $msolUser.ImmutableID
                }
        
            }
            else 
            {
                Write-host -ForegroundColor red  "User Not Found"
                $msolUser.userprincipalname
                $msolAccountsnotfound += $msoluser

            }

        }
        
        $msolAccountsnotfound | export-csv .\Not_Stamped_Notfound.csv
    }
    else
    {
        Write-host "converted_ImmutableID.csv not found, please run -convertonly first"
    }

    
}
