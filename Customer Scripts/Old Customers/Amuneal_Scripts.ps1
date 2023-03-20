# remove routing address
function Remove-RoutingDomain {
    param (
        HEXDomain,
        RoutingDomain
    )
    
}
$mailboxes = get-mailbox -OrganizationalUnit amuneal.com |sort displayName
$updatedUsers = @()

foreach ($mbx in $mailboxes)
{
    Write-Host "Checking mailbox $($mbx.DisplayName) .. " -fore cyan -nonewline
    $aliasarray = $mbx.EmailAddresses.ToStringArray()
    Write-Host $aliasarray.count "aliases found . " -fore darkcyan -nonewline
    foreach ($alias in $aliasarray)
    {
        if ($alias -like "*@routing.amuneal.com")
        {
            Write-host "Removing address $($alias)" -NoNewline
            Set-Mailbox $mbx.alias -EmailAddresses @{remove=$alias}
            $updatedUsers += $mbx
        }
        else
        {
            Write-host "." -nonewline -fore Yellow
        } 
    }
    Write-Host "done." -fore green
}

# Create RSE MSOL Users

foreach ($rsembx in $rsemailboxes)
{
    Write-Host "Creating RSE MSOLUser $($rsembx.name)"
    if (!($msolusercheck = get-msoluser -userprincipalname $rsembx.email -ea silentlycontinue))
    {
        New-MsolUser -UserPrincipalName $rsembx.email -FirstName $rsembx.FirstName -LastName $rsembx.LastName -DisplayName $rsembx.DisplayName -password (ConvertTo-SecureString -String 'Year96Temerity' -AsPlainText -Force) -Title $rsembx.Title -City $rsembx.City -State $rsembx.State -UsageLocation US
    }
    else
    {
        Write-Host "User already Exists." -ForegroundColor Yellow
    }
}