$RSEUserImport = Import-CSV "C:\Users\fred5646\Rackspace Inc\American Car Center - General\RSE Migration\Rackspace Email 365 License Status2.csv"

$updatedE1Users = @()
$skippedE1Users = @()

$StandardLicense = Get-MsolAccountSku | Where-Object {$_.AccountSkuID -like "*Standardpack"} | Select-Object -ExpandProperty AccountSkuID
foreach ($user in $RSEUserImport)
{
    if (Get-MsolUser -UserPrincipalName $user.UserPrincipalName | ? {($_.licenses).AccountSkuId -match $StandardLicense})
    {
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $StandardLicense
        Write-Host "Removing E1 license from user $($user.DisplayName)" -ForegroundColor Cyan
        $updatedE1Users += $user.UserPrincipalName
        }
    else
    {
        Write-Host "User $($user.DisplayName) not licensed. skipping"
        $skippedE1Users += $user.UserPrincipalName
    }
}