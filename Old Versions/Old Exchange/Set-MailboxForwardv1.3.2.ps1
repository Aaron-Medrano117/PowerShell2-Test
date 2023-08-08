<#
.SYNOPSIS
Configure Forwarding from Source mailbox to Target mailbox based on CSV/Excel Input

.DESCRIPTION
For a tenant to tenant migration, this sets up the source mailbox to forward to the target mailbox. This script is designed to be run from the source tenant. Uses the ForwardingSMTPAddress attribute to set up forwarding.

.EXAMPLE
.\Set-MailboxForwardv1.3.2.ps1 -ImportFileLocation ".\ForwardingMailboxes.csv"

.EXAMPLE
.\Set-MailboxForwardv1.3.2.ps1 -ImportFileLocation ".\ForwardingMailboxes.xlsx"
.EXAMPLE
.\Set-MailboxForwardv1.3.2.ps1 -ImportFileLocation ".\ForwardingMailboxes.xlsx" -worksheetname "Sheet1"

.NOTES
Set-Forward version 1.3.3
Script Details: Configure Forwarding from Source mailbox to Target mailbox based on CSV/Excel Input
Version updated 3/21/2023
Version author - Aaron Medrano
Written by Aaron Medrano, Andrew Cronic, and John Williams - Arraya Solutions Modern Workplace Solutions Engineering Team
#>
param (
        [int]$ImportFileLocation,
        [string]$WorkSheetName
    )

Start-Transcript
#Check if the user is running the script in an Exchange Online PowerShell session
if (-not (Get-Module -Name ExchangeOnlineManagement)) {
    Write-Host "This script must be run in an Exchange Online PowerShell session. Please run the following command and try again:" -ForegroundColor Red
    Write-Host "Connect-ExchangeOnline" -ForegroundColor Red
    Write-Host "Must connect to Source tenant to configure forwarding" -ForegroundColor Red
    return
}

#output the tenant name
$onMicrosoftDomain = Get-MsolDomain | Where-Object {$_.IsDefault -eq $true}
Write-Host "You are connected to the following tenant: $($onMicrosoftDomain.Name)" -ForegroundColor Green

function Write-ProgressHelper {
    param (
        [int]$ProgressCounter,
        [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [int]$TotalCount
    )
    $secondsElapsed = (Get-Date) – $global:start
    $progresspercentcomplete = [math]::Round((($progresscounter / $TotalCount)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$TotalCount+"]"

    $progressParameters = @{
        Activity = $Activity
        Status = "$progressStatus $($secondsElapsed.ToString('hh\:mm\:ss'))"
        PercentComplete = $progresspercentcomplete
    }

    # if we have an estimate for the time remaining, add it to the Write-Progress parameters
    if ($secondsRemaining) {
        $progressParameters.SecondsRemaining = $secondsRemaining
    }
    if ($ID) {
        $progressParameters.ID = $ID
    }
    if ($CurrentOperation) {
        $progressParameters.CurrentOperation = $CurrentOperation
    }

    # Write the progress bar
    Write-Progress @progressParameters

    # estimate the time remaining
    $global:secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount – $progresscounter)
}

#Test Headers and Import Array of Objects
function Test-ImportedObjectHeaders {
    param (
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter(Mandatory)]
        [string[]]$RequiredHeaders,
        [string]$Worksheet
    )
    # Check file extension and import file
    if ($ImportFileLocation -match "\.csv$") {
        Write-Host "CSV File Found"
        $importedData = Import-CSV -Path $ImportFileLocation
    }
    elseif ($ImportFileLocation -match "\.xlsx$" -or $ImportFileLocation -match "\.xls$") {
        Write-Host "XLSX/XLS File Found"
        if ($Worksheet) {
            $importedData = Import-Excel -Path $ImportFileLocation -WorksheetName $Worksheet
        }
        else {
            $importedData = Import-Excel -Path $ImportFileLocation
        }
        
    }
    else {
        Write-Host "Unsupported version uploaded. Please try running again using CSV, XLSX, or an XLS file type" -ForegroundColor Red
        return
    }

    # Retrieve the headers of the imported CSV file
    $importedHeaders = $importedData[0].PSObject.Properties.Name

    # Compare the necessary headers array with the imported CSV headers array
    $missingHeaders = Compare-Object -ReferenceObject $requiredHeaders -DifferenceObject $importedHeaders -IncludeEqual | Where-Object {$_.SideIndicator -eq "<="}

    # Check if any necessary headers are missing and output an error message if necessary
    if ($missingHeaders) {
        Write-Host "Error: The following necessary headers are missing from the file: $($missingHeaders.InputObject -join ', ')"
    }
    else {
        $importedData
        Write-Host "Success: All necessary headers are present in the file."
    }
    Write-Host ""
}

# Import Variables
if ($ImportFileLocation) {
    $ImportFileLocation = $ImportFileLocation
}
else {
    $ImportFileLocation = Read-Host "What is the pathfile of the list of users? Supports CSV and XLSX/XLS files"
}
#Check if Headers are in Input - Source and Target Headings required
$requiredHeaders = @("Source", "Target")

if ($WorkSheetName) {
    $WorkSheetName = $WorkSheetName
}
else {
    $WorkSheetName = Read-Host "What is the worksheet name of the the excel file?"
}

# With WorkSheet Specified
if ($WorkSheetName) {
    $forwardingUsers = Test-ImportedObjectHeaders -FilePath $ImportFileLocation -RequiredHeaders $requiredHeaders -Worksheet $WorkSheetName
}
# Without WorkSheet Specified
else {
    $forwardingUsers = Test-ImportedObjectHeaders -FilePath $ImportFileLocation -RequiredHeaders $requiredHeaders
}


#Check if Headers are in Input - Source and Target Headings required
if ($forwardingUsers) {
    #ProgressBar
    $progresscounter = 1
    $global:start = Get-Date
    [nullable[double]]$global:secondsRemaining = $null

    ## Create variables for output
    # Array to store all completed users
    $completedUsers = @()
    # Array to store all errors
    $allErrors = @()
    # Array to store mailbox data post-forwarding-changes / audit to confirm forwarding changes were successful
    $auditResults = @()

    foreach ($user in $allMailboxes) {
        $SourcePrimarySMTPAddress = $user.Source
        $DestinationPrimarySMTPAddress = $user.Target

        #progress bar
        Write-ProgressHelper -Activity "Updating Forwarding for $($SourcePrimarySMTPAddress)" -ProgressCounter ($progresscounter++) -TotalCount ($forwardingUsers).count

        Write-Host "Cutting Over User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -nonewline
        ## Set Mailbox to Forward from Source to Destination Mailbox
        Write-Host "Set Forward to $($DestinationPrimarySMTPAddress)  " -foregroundcolor Magenta -nonewline
        Try{       
            Set-Mailbox $SourcePrimarySMTPAddress -ForwardingSmtpAddress $DestinationPrimarySMTPAddress -ErrorAction Stop
            Write-Host "Completed" -ForegroundColor Green
            $completedUsers += $user
        }
        Catch {
            Write-Host "Failed" -ForegroundColor red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "SetForward" -Force
            $currenterror | Add-Member -type NoteProperty -Name "SourcePrimarySMTPAddress" -Value $SourcePrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "DesinationPrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $allErrors += $currenterror
            continue
        }
    }
    Write-Host ""
}

#Audit Mailbox Forwarding previously set
if ($forwardingUsers) {
    #reset progresscounter
    $progresscounter = 1
    $global:start = Get-Date
    [nullable[double]]$global:secondsRemaining = $null
    Write-Host "Auditing forwarding settings set by previous operations..."

    foreach ($user in $forwardingUsers) {
        $SourcePrimarySMTPAddress = $user.Source
        $DestinationPrimarySMTPAddress = $user.Target
        $ProgressPreference = "Continue" #added

        #progress bar
        Write-ProgressHelper -Activity "Getting Forwarding Settings for $($SourcePrimarySMTPAddress)" -ProgressCounter ($progresscounter++) -TotalCount ($forwardingUsers).count
        Write-Host "Auditing forwarding settings for User $($SourcePrimarySMTPAddress) .. " -foregroundcolor Cyan -NoNewline
        Try{       
            $mailboxData = Get-EXOMailbox $SourcePrimarySMTPAddress -Properties ForwardingSmtpAddress,DeliverToMailboxAndForward,ForwardingAddress -ErrorAction Stop #updated to EXO-Mailbox
            $mailboxData | Add-Member -type NoteProperty -Name "ForwardingAddressMismatch" -Value "" -Force

            if($mailboxData.ForwardingSmtpAddress -replace '^SMTP:' -ne $DestinationPrimarySMTPAddress)
            {
                $mailboxData.forwardingAddressMismatch = "TRUE"
            }
            else {
                $mailboxData.forwardingAddressMismatch = "FALSE"
            }

            Write-Host "Completed" -ForegroundColor Green #added
            $auditResults += $mailboxData | select UserPrincipalName, PrimarySmtpAddress, *Forward*
        }
        Catch {
            Write-Host "Failed" -ForegroundColor red
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "Get-Mailbox" -Force
            $currenterror | Add-Member -type NoteProperty -Name "SourcePrimarySMTPAddress" -Value $SourcePrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "DesinationPrimarySMTPAddress" -Value $DestinationPrimarySMTPAddress -Force
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $allErrors += $currenterror
            continue
        }
    }

    #AUDIT END

    Write-Host "Displaying Audit Results in console:"
    Write-Host "ForwardingAddressMismatch property TRUE indicates the destination SMTP address was not set correctly" -ForegroundColor Yellow
    $AuditResults | ?{$_.ForwardingAddressMismatch -eq $true}

    Write-Host ""
    Write-Host "Exporting Audit Results .csv to $pwd"
    $auditResults | Export-Csv ".\ForwardingAuditReport-$(get-date -f yyyy-MM-dd-hh-mm).csv" -NoTypeInformation
    

    Write-Host "Completed in"((Get-Date) - $global:start).ToString('hh\:mm\:ss')"" -ForegroundColor Cyan
    write-host "$($completedUsers.count) / $($forwardingUsers.count) Users Forwarding Completed" -ForegroundColor Cyan
    Write-Host "$($allErrors.count) / $($forwardingUsers.count) Users Generated Errors. Check the $allErrors variable for more details" -ForegroundColor Red

    Write-Host ""
    Write-Host ""
    Write-Host ""
}
Stop-Transcript