$users = Import-Csv (folder path of the .csv file)
$users | ForEach-Object {
New-MailUser -Name $_.Name -ExternalEmailAddress $_.ExternalEmailAddress -MicrosoftOnlineServicesID $_.MicrosoftOnlineServicesID -Password (ConvertTo-SecureString -String '$_.Password' -AsPlainText -Force)}


####

$importcsv = Import-Csv (folder path of the .csv file)
foreach ($user in $importcsv) 
{
    $upn = $user.newupn
    if ($ADUser = get-aduser -Filter {UserPrincipalName -eq $upn })
    {
        Write-Host "Updating DisplayName for user $($upn) ..." -ForegroundColor cyan -NoNewline
        $ADUser | Set-ADUser -DisplayName $user.DisplayName

        if (!(get-remotemailbox $upn))
        {
            Write-Host "Enable Remote Mailbox" -NoNewline
            Enable-RemoteMailbox $upn -RemoteRoutingAddress $upn -whatif      
        }
        else
        {
            Write-Host "user already a remote mailbox" -NoNewline -ForegroundColor Yellow
            Write-Host "Set Email aliases ..." -NoNewline
            $emailAliasSplit = $user.email_aliases -split ","
                foreach ($alias in $emailAliasSplit)
                {
                    Set-remoteMailbox -identity $upn -emailaddresses @{add=$alias}
                }
        }
       
        
        write-host "... done" -ForegroundColor green
    }
    
    else
    {
        Write-Host "no user found for $($upn)" -ForegroundColor Red
    }
}


###

foreach ($user in $importcsv) 
{
    $upn = $user.newupn
    if ($ADUser = get-aduser -Filter {UserPrincipalName -eq $upn })
    {
        Write-Host "Updating DisplayName for user $($upn) ..." -ForegroundColor cyan -NoNewline
        $ADUser | Set-ADUser -DisplayName $user.DisplayName        
        write-host "... done" -ForegroundColor green
    }
    
    else
    {
        Write-Host "no user found for $($upn)" -ForegroundColor Red
    }
}

####

foreach ($user in $importcsv) 
{
    $upn = $user.newupn
    if ($ADUser = get-aduser -Filter {UserPrincipalName -eq $upn})
    {
        if (!(get-remotemailbox $upn -ea SilentlyContinue))
        {
            Write-Host "Enable Remote Mailbox" -NoNewline
            Enable-RemoteMailbox $upn -RemoteRoutingAddress $upn      
        }
        else
        {
            Write-Host "user already a remote mailbox" -NoNewline -ForegroundColor Yellow
            Write-Host "Set Email aliases ..." -NoNewline
            $emailAliasSplit = $user.email_aliases -split ","
                foreach ($alias in $emailAliasSplit)
                {
                    Set-remoteMailbox -identity $upn -emailaddresses @{add=$alias}
                }
        }
       
        
        write-host "... done" -ForegroundColor green
    }
    
    else
    {
        Write-Host "no user found for $($upn)" -ForegroundColor Red
    }
}


foreach ($user in $importcsv) 
{
    $upn = $user.newupn
    $ADUser = get-aduser -Filter {UserPrincipalName -eq $upn} | select distinguishedname
}