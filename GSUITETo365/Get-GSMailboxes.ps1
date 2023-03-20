<#>
.SYNOPSIS
This script is designed to find mail users within GSUITE for mailboxes in Exchange Online.

.DESCRIPTION

MANDATORY REQUIREMENT: 
In order to run this shell, you must complete the following:

1) Install PS Module
Install-Module -Name PSGSuite -RequiredVersion 2.24.0

2) Set PowerShell GSUITE Config file. Update below variables
$ConfigName =  "GSuite"
$Preference = "Domain"
$P12KeyPath = # "C:\GSuite\psgsuite-284106-f422e66d3841.p12"
$AppEmail = # "psgsuite@xxxxxxxxxxxxxxxxxxxx.iam.gserviceaccount.com"
$AdminEmail = # "admin@aventislab.info"
$Domain = # "aventislab.info"
$ServiceAccountClientID = # "10745224254xxxxxxxxx"
For More Details, check out: https://psgsuite.io/Initial%20Setup/
 
Set-PSGSuiteConfig -ConfigName $ConfigName -P12KeyPath $P12KeyPath -AppEmail $AppEmail -AdminEmail $AdminEmail -Domain $Domain  -ServiceAccountClientID $ServiceAccountClientID

.EXAMPLE
Pull all GSuite Users with Mailboxes and output to desired folder path.
.\Get-GSMailboxes.ps1 -OutputCSVFilePath "c:\temp"
#>

#Export All users

function Get-GSMailboxes {
    param (
        [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Output CSV File path")][string]$OutputCSVFilePath
    )

    $Mailboxes = Get-GSUserList | Where-Object {$_.IsMailboxSetup -eq $True} 

    $AllGSUsers = @()
    $csvFileName = "GSUITE_Users.csv"

    foreach ($User in $Mailboxes) { 

        $EXAttributes = $user.emails.Address
        [array]$UserAttributes = $User | select -ExpandProperty Name
        $username = $user.user -split '@'

        $tmp = new-object PSObject

        $tmp | add-member -type noteproperty -name "UserPrincipalName" -Value $user.user
        $tmp | add-member -type noteproperty -name "DisplayName" -Value $UserAttributes.FullName
        $tmp | add-member -type noteproperty -name "FirstName" -Value $UserAttributes.GivenName
        $tmp | add-member -type noteproperty -name "LastName" -Value $UserAttributes.FamilyName
        $tmp | add-member -type noteproperty -name "OrgUnitPath" -Value $user.OrgUnitPath
        $tmp | add-member -type noteproperty -name "PrimarySmtpAddress" -Value $user.PrimaryEmail
        $tmp | add-member -type noteproperty -name "EmailAddresses" -Value ($EXAttributes -join ",")
        $tmp | add-member -type noteproperty -name "HiddenFromAddressListsEnabled" -Value $User.IncludeInGlobalAddressList
        $tmp | add-member -type noteproperty -name "ExternalEmailAddress" -Value $username[0] + "@gs." + $username[1]
        $AllGSUsers += $tmp
    }

    $AllGSUsers | Export-csv "$($OutputCSVFilePath)\$($csvFileName)" –notypeinformation –encoding utf8
    
}
