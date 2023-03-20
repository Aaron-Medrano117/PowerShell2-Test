$mailboxPropertiesInitial = Import-Csv C:\Users\raxadmin\Desktop\TheRoomPlace_AllLicensedMailboxes_20210824.csv
$mailboxProperties = $mailboxPropertiesInitial | ?{$_.ImmutableID -eq $null}
#ProgressBar1
$progressref = ($mailboxPropertiesInitial).count
$progresscounter = 0

foreach ($mbx in $mailboxPropertiesInitial)
{
    $upn = $mbx.UserPrincipalName
    $DisplayName = $mbx.DisplayName
    $EmailAddress = $mbx.PrimarySmtpAddress

    #ProgressBar2
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking for ADUser $($DisplayName)"

    #if ($adUser = Get-ADUser -Filter {Mail -eq $upn} -Properties ObjectGUID)
    if ($adUser = Get-ADUser -Filter {UserPrincipalName -eq $upn} -Properties ObjectGUID, mail)
    {
        Write-Host $upn -ForegroundColor Green
        $mbx | Add-Member -MemberType NoteProperty -Name "ExistsOnPrem" -Value $true -force
        $mbx | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $adUser.DistinguishedName -force
        $mbx | Add-Member -MemberType NoteProperty -Name "OnPremUPN" -Value $adUser.UserPrincipalName -force
        $objGUID = $adUser | select -ExpandProperty ObjectGUID | select -ExpandProperty Guid
        $ImmutableID = [System.Convert]::ToBase64String(([GUID]($objGUID)).ToByteArray())
        $mbx | Add-Member -MemberType NoteProperty -Name "Mail" -Value $adUser.mail -force
        $mbx | Add-Member -MemberType NoteProperty -Name "ImmutableID" -Value $ImmutableID -force
    }
    elseif ($adUser = Get-ADUser -filter {Mail -eq $emailAddress} -Properties ObjectGUID, mail)
    {
        Write-Host $upn -ForegroundColor Cyan
        $mbx | Add-Member -MemberType NoteProperty -Name "ExistsOnPrem" -Value $true -force
        $mbx | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $adUser.DistinguishedName -force
        $mbx | Add-Member -MemberType NoteProperty -Name "OnPremUPN" -Value $adUser.UserPrincipalName -force
        $objGUID = $adUser | select -ExpandProperty ObjectGUID | select -ExpandProperty Guid
        $ImmutableID = [System.Convert]::ToBase64String(([GUID]($objGUID)).ToByteArray())
        $mbx | Add-Member -MemberType NoteProperty -Name "Mail" -Value $adUser.mail -force
        $mbx | Add-Member -MemberType NoteProperty -Name "ImmutableID" -Value $ImmutableID -force
    }
    elseif ($adUser = Get-ADUser -filter {Name -eq $DisplayName} -Properties ObjectGUID, mail)
    {
        Write-Host $DisplayName -ForegroundColor Yellow
        $mbx | Add-Member -MemberType NoteProperty -Name "ExistsOnPrem" -Value $true -force
        $mbx | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $adUser.DistinguishedName -force
        $mbx | Add-Member -MemberType NoteProperty -Name "OnPremUPN" -Value $adUser.UserPrincipalName -force
        $objGUID = $adUser | select -ExpandProperty ObjectGUID | select -ExpandProperty Guid
        $ImmutableID = [System.Convert]::ToBase64String(([GUID]($objGUID)).ToByteArray())
        $mbx | Add-Member -MemberType NoteProperty -Name "Mail" -Value $adUser.mail -force
        $mbx | Add-Member -MemberType NoteProperty -Name "ImmutableID" -Value $ImmutableID -force
    }
    else
    {
        Write-Host $upn -ForegroundColor Red
        $mbx | Add-Member -MemberType NoteProperty -Name "ExistsOnPrem" -Value $false -force
        $mbx | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $Null -force
        $mbx | Add-Member -MemberType NoteProperty -Name "OnPremUPN" -Value $Null -force
        $mbx | Add-Member -MemberType NoteProperty -Name "Mail" -Value $adUser.mail -force
        $mbx | Add-Member -MemberType NoteProperty -Name "ImmutableID" -Value $Null -force
    }
}