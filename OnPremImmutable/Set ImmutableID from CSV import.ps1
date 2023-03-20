[string]$hardmatchCSV = Read-Host "What is the file path of the csv file to import?"

$import_hardmatch = Import-Csv $hardmatchCSV

foreach ($user in $import_hardmatch)
{
Write-Host "Setting Immutable for user" $user.DisplayName -ForegroundColor White

Set-MsolUser -UserPrincipalName $user.UPN_Keep -ImmutableId $user.ImmutableID

Get-MsolUser -UserPrincipalName $user.UPN_Keep | ft DisplayName, UserPrincipalName, ImmutableId

}