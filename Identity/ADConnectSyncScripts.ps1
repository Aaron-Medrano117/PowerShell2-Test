$mailboxProperties = Import-Csv $HOME\Desktop\Duplicates.csv |  # Path to CSV


#Add Immutable ID to CSV

function Get-ImmutableID {

foreach ($mbx in $mailboxProperties)
 {
    $DisplayName = $mbx.DisplayName

    if ($adUser = Get-ADUser -Filter {DisplayName -like $DisplayName} -Properties ObjectGUID)
    {
    Write-Host $DisplayName -ForegroundColor Green
    $guid = $adUser | select -ExpandProperty ObjectGUID
    $immutableID = [System.Convert]::ToBase64String(([GUID]($guid)).ToByteArray())

    #add Immutable ID to $mailboxproperties array
    $mbx.ImmutableID = $immutableID
    }

 }

#Export updated CSV to desktop
$mailboxProperties | Export-Csv $HOME\Desktop\Duplicates.csv -NoTypeInformation -Encoding UTF8

}

# Remove User from Syncing

function Remove-UserSync {
foreach ($mbx in $mailboxProperties)
    {
    $DisplayName = $mbx.DisplayName

    if ($adUser = Get-ADUser -Filter {DisplayName -eq $DisplayName} -Properties AdminDescription)
    {

        Write-Host $DisplayName -ForegroundColor Green -NoNewline
        $adUser | set-aduser -add @{"admindescription"="User_"}

    }
    }
}

#Set User to Sync

foreach ($mbx in $mailboxProperties)
{
    $DisplayName = $mbx.DisplayName
    
    if ($adUser = Get-ADUser -Filter {DisplayName -eq $DisplayName} -Properties admindescription)
    {
    if ($aduser.admindescription -like "User_") 
    {
        Write-Host $DisplayName -ForegroundColor Green -NoNewline
        $adUser | set-aduser -remove @{"admindescription"="User_"}
        
        #Check if User is syncing or not.
                            {
        Write-host "... Set to Not Sync" -ForegroundColor Green
        }
        else
        {
        Write-host "... Set to Sync Still" -ForegroundColor Yellow
        }

    }
                              
}

#Remove Deleted Users Preventing ADConnect

[string]$hardmatchCSV = Read-Host "What is the file path of the csv file to import?"

$import_hardmatch = Import-Csv $hardmatchCSV


foreach ($user in $import_hardmatch)
{
#gather variables
$immutableID = $user.ImmutableID
$UPN = $user.OnMicrosoft_Remove
$Displayname = $user.DisplayName


$immutableID = $user.ImmutableID
$UPN = $user.OnMicrosoft_Remove

#remove Deleted User
if ($deleteduser = Get-MsolUser -UserPrincipalName $upn -ReturnDeletedUsers -erroraction SilentlyContinue) 
    {
    Write-Host ""
    Write-Host "Removing Deleted User $DisplayName" -ForegroundColor White -NoNewline
    $deleteduser | Remove-MsolUser -RemoveFromRecycleBin -force

    $confirmDelete = Get-MsolUser -UserPrincipalName $upn -ReturnDeletedUsers -ErrorAction SilentlyContinue

        if (!$confirmDelete)
        {
        write-host "... Completed" -ForegroundColor Green
        }
        else {
        Write-Host "... still exists" -ForegroundColor Red
        }
    }
}

#Set Immutable ID on 365

foreach ($user in $import_hardmatch)
{

$immutableID = $user.ImmutableID
$UPN = $user.OnMicrosoft_Remove
$Displayname = $user.DisplayName

        Set-MsolUser -UserPrincipalName $user.UPN_Keep -ImmutableId $immutableID
        Write-Host "Succesfully configured" $CheckImmutableID "for" $DisplayName -ForegroundColor Green
        
    Get-MsolUser -UserPrincipalName $user.UPN_Keep | ft DisplayName, UserPrincipalName, ImmutableId

}

# get Immutable ID for mismatched users 1
$UPN = @()
$immutableID = @()

$mismatchedUPN | foreach {
$UPN = $_.OnPremUPN
if ($AdUser = Get-ADUser -Filter {UserPrincipalName -eq $UPN} -ErrorAction SilentlyContinue)
    {
    Write-host $_.DisplayName -ForegroundColor White
        $GUID = $AdUser | select -ExpandProperty ObjectGUID
        $immutableID = [System.Convert]::ToBase64String(([GUID]($guid)).ToByteArray())

        if ($immutableID) {
            Write-Host "ImmutableID found for" $_.DisplayName -ForegroundColor Green
            $_.ImmutableID = $immutableID
            }
        else 
            {
            Write-Host "No Immutable ID found for" $_.DisplayName -ForegroundColor Red
            }
     }
else {
    Write-Host "Unable to find User for" $DisplayName -ForegroundColor Red
    }
}

#get immutable ID for mismatched remaining users 2
$UPN = @()
$immutableID = @() 
$mismatchedUPN2 = @()
$mismatched2 = 
$mailboxProperties | foreach {
if (Get-ADUser -filter {UserPrincipalName -eq $_.OnPremUPN} |?  ($_.RecipientTypeDetails -eq "LinkedMailbox")) {
$mismatchedUPN += $_
}
}


# Get immutable ID for mismatched remaining users 3
$mismatched3 = @() 
$mailboxProperties |?  {$_.RecipientTypeDetails -eq "LinkedMailbox" -and ($_.OnPremUPN -ne $_.HEXUPN)} | foreach {
$OnPremUPN =  $_.OnPremUPN
if ($Aduser = Get-ADUser -filter {UserPrincipalName -eq $OnPremUPN}) {

Write-host $_.DisplayName -ForegroundColor White
        $GUID = $AdUser | select -ExpandProperty ObjectGUID
        $immutableID = [System.Convert]::ToBase64String(([GUID]($guid)).ToByteArray())
$mismatched3 += $aduser
}
}


foreach ($mbx in $mismatchUPN) {
$OnPremUPN = $mbx.OnPremUPN
$HEXUPN = $mbx.HEXUPN

if ($ADUser1 = Get-ADUser -Filter {UserPrincipalName -eq $OnPremUPN})
    {
    Write-Host "Updating UPN for" $mbx.DisplayName "..." -ForegroundColor White -NoNewline
    $ADUser1 | Set-ADUser -UserPrincipalName $HEXUPN -ErrorAction SilentlyContinue
    Write-Host "done" -ForegroundColor Green
    }
else
    {
    Write-Host $mbx.DisplayName "not found" -ForegroundColor Red
    }

}


$OnPremUPN = @()
$HEXUPN = @()

foreach ($mbx in $mismatchUPN) {
$OnPremUPN = $mbx.OnPremUPN
$HEXUPN = $mbx.HEXUPN

#if ($ADUser1 = Get-ADUser -Filter {UserPrincipalName -eq $HexUPN})
#    {
#    Write-Host "UPN updated for" $mbx.DisplayName "..." -ForegroundColor White -NoNewline
#    $ADUser1 | Set-ADUser -UserPrincipalName $HEXUPN -ErrorAction SilentlyContinue
#    #Write-Host "done" -ForegroundColor Green
#    }

if ($Aduser = Get-ADUser -filter {UserPrincipalName -eq $HEXUPN}) {
        Write-host $_.DisplayName "..." -ForegroundColor White -NoNewline
        $GUID = $AdUser | select -ExpandProperty ObjectGUID
        $immutableID = [System.Convert]::ToBase64String(([GUID]($guid)).ToByteArray())
        $mbx.ImmutablID = $immutableID
        Write-Host "done" -ForegroundColor Green
    }
}
#else
#    {
#    Write-Host $mbx.DisplayName "not found" -ForegroundColor Red
#    }

}

# Get HEX UPN and convert to Immutable ID
#REGION MEX06/8/9

$domainController = Get-ADDomainController -DomainName "acct.mlsrvr.com" -Discover
$serverFQDN = $domainController.Name + "." + $domainController.Domain

$mailboxes = get-mailbox -organizationalunit apgpoly.com | select DisplayName, Alias, PrimarySMTPAddress, name, RecipientTypeDetails
foreach ($mailbox in $mailboxes)
{
    $HEXUPN = $mailbox.primarysmtpaddress.tostring()
    $Aduser = Get-ADUser -filter {UserPrincipalName -eq $HEXUPN} -Server $serverFQDN 
    Write-host "Getting Immutable ID for $($mailbox.DisplayName) ..." -ForegroundColor White -NoNewline
    $GUID = $AdUser | select -ExpandProperty ObjectGUID
    $immutableID = [System.Convert]::ToBase64String(([GUID]($guid)).ToByteArray())
    $mailbox | add-member -type noteproperty -name "UserPrincipalName" -Value $HEXUPN
    $mailbox | add-member -type noteproperty -name "ImmutableID" -Value $immutableID -force
}
