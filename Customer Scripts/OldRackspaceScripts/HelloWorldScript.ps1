#Hello World Script
function Get-HelloWorld
{

$firstname = Read-Host "What is your first name?"
$lastname = Read-Host "What is your last name?"

Write-Host "Hello $firstname $lastname"
Write-Host "You are logged using folder path $env:OneDrive"
Write-Host "You are logged into $env:COMPUTERNAME"

}