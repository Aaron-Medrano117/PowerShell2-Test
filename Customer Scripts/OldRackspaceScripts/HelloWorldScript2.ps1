#Hello World Script
function Get-HelloWorld
{

param ($firstname, $lastname)

Write-Host "Hello $firstname $lastname"
Write-Host "You are logged using folder path $env:OneDrive"
Write-Host "You are logged into $env:COMPUTERNAME"

}