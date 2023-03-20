<#>
.SYNOPSIS
This script is designed to pull the list of calendar subscriptions for each user to find all "shared calendars" in GSUITE.

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
.\Get-GSCalendarSubscriptions.ps1 -OutputCSVFilePath "c:\temp"
#>

# Gather All GSUITE Calendar Subscriptions

$allGSUsers = Get-GSUserList
$allGSCalendars = @()

foreach ($GSUser in $AllGSUsers)
{
    $GSCalendars = @()

    $GSCalendars = Get-GSCalendarSubscription -User $GSUser.user
    Write-Host "Found $($GSCalendars.count) GSCalendar Subscriptions for $($GSUser.user). Gathering details ..." -NoNewline -ForegroundColor Cyan

    foreach ($GSCalendar in $GSCalendars)
    {
        $GSCalendar = Get-GSCalendarSubscription -User $GSUser.user -CalendarId $GSCalendar.id

        $tmp = new-object PSObject

        $tmp | add-member -type noteproperty -name "User" -Value $GSUser.user
        $tmp | add-member -type noteproperty -name "Summary" -Value $GSCalendar.Summary
        $tmp | add-member -type noteproperty -name "AccessRole" -Value $GSCalendar.AccessRole
        $tmp | add-member -type noteproperty -name "ID" -Value $GSCalendar.ID
        $tmp | add-member -type noteproperty -name "Kind" -Value $GSCalendar.Kind

        $allGSCalendars += $tmp
    }

    Write-Host "done" -ForegroundColor Green   
}