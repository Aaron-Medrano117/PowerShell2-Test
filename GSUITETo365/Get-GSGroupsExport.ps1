<#>
.SYNOPSIS
This script is designed to find groups within GSUITE for Distribution Groups in Exchange Online.

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
Pull all GSuite Groups with Mailbox members and output to desired folder path.
.\Get-GSGroupsExport.ps1 -OutputCSVFilePath "c:\temp"
#>

#export GS Groups

function Get-GSGroupsExport {
    param (
        [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Output CSV File path")][string]$OutputCSVFilePath
    )

    $Groups = Get-GSGroup
    $AllGSGroups = @()

    $csvFileName = "GSUITE_Groups.csv"

    Write-Host "Gathering $($Groups.Count) GSuite Groups and respective members ..." -ForegroundColor Cyan -NoNewline

    foreach ($Group in $Groups) {
        $GroupMembers = Get-GSGroupMember -Identity $Group.id

        $tmp = "" | Select Group, DirectMembersCount, PrimarySmtpAddress, Description, Members 

        $tmp.Group =$Group.Name
        $tmp.DirectMembersCount = $Group.DirectMembersCount
        $tmp.PrimarySmtpAddress = $Group.Email
        $tmp.Description = $Group.Description
        $tmp.Members = ($GroupMembers.Email -join ",")
        
        $AllGSGroups += $tmp
    }
    Write-Host "done" -ForegroundColor Green
    $AllGSGroups | Export-csv "$($OutputCSVFilePath)\$($csvFileName)" –notypeinformation –encoding utf8
}



