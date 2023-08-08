param (
    [Parameter(Mandatory=$True)] [string] $CompanyName,
    [Parameter(Mandatory=$false)] [string] $ProjectName,
    [Parameter(Mandatory=$false)] [string] $ProjectKeyword,
    [Parameter(Mandatory=$True,HelpMessage="Choose between Full, Trial (Pre-Stage), or Verification Migration Types")] [string] $MigrationType,
    [Parameter(Mandatory=$True)] 
    [System.Management.Automation.PSCredential] 
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()] $Credentials,
    [Parameter(Mandatory=$false,HelpMessage="Specify the Source Email Address of the Migrating User")] [string] $SourceAddress,
    [Parameter(Mandatory=$false,HelpMessage="Specify the Destination Email Address of the Migrating User")] [string] $DestinationAddress,
    [Parameter(Mandatory=$false)] [array] $ImportUsers,
    [Parameter(Mandatory=$false)] [string] $OutputCSVFolderPath,
    [Parameter(Mandatory=$false)] [string] $DaysOlderThan,
    [Parameter(Mandatory=$false)] [switch] $ALL
)
    #Set Global Variables
    $global:CompanyName = $CompanyName
    $global:ProjectName = $ProjectName
    $global:SourceAddress = $SourceAddress
    $global:DestinationAddress =$DestinationAddress
    $global:ImportUserList = Import-CSV $ImportUsers
    $global:MigrationType = $MigrationType
    $shorterdate = $((Get-Date).ToShortDateString().Replace("/","-"))
    $Global:OutputCSVFilePath = $OutputCSVFolderPath + "\MigrationSubmission_$($shorterdate).csv"
    $global:DaysOlderThan = $DaysOlderThan
    $global:mwTicket = Get-MW_Ticket -Credentials $Credentials
    $global:btTicket = Get-BT_Ticket -Credentials $Credentials -ServiceType BitTitan
    $Global:ProjectKeyword = $ProjectKeyword

<#
SYNOPSIS: Submit Users for Pre-stage, Full, or Verification

Prerequiesites: You'll need to download the MigrationWiz PowerShell Module and Import the module
https://www.bittitan.com/doc/powershell.html#PagePowerShellintroductionmd-powershell-module-installation
Download from the above link
Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll'


EXAMPLE: Submit ALL Users for an Entire Customer to Full Migration with Stored Variable Credentials
Submit-MigWizMailboxMigration -CompanyName Company -All -MigrationType Full -Credentials $Credentials

EXAMPLE: Submit ALL Users for an Entire Customer to Pre-Stage Migration with Stored Variable Credentials
Submit-MigWizMailboxMigration -CompanyName Company -All -MigrationType Trial -Credentials $Credentials

EXAMPLE: Submit ALL Users for an Entire Customer to Verification Migration with Stored Variable Credentials
Submit-MigWizMailboxMigration -CompanyName Company -All -MigrationType Verification -Credentials $Credentials

EXAMPLE: Submit ALL Users for One Project to Full Migration
Submit-MigWizMailboxMigration -CompanyName Company -ProjectName MigWizProject -MigrationType Full -Credentials $Credentials

EXAMPLE: Submit ALL Users for One Project to Pre-Stage Migration
Submit-MigWizMailboxMigration -CompanyName Company -ProjectName MigWizProject -MigrationType Trial -Credentials $Credentials

EXAMPLE: Submit ALL Users for One Project to Verification Migration
Submit-MigWizMailboxMigration -CompanyName Company -ProjectName MigWizProject -MigrationType Verification -Credentials $Credentials

EXAMPLE: Submit ONE Users for in a Customer to Full Migration (with Source Email Address)
Submit-MigWizMailboxMigration -CompanyName Company -SourceAddress user@example.org -MigrationType Full -Credentials $Credentials

EXAMPLE: Submit ONE Users for in a Customer to Full Migration (with Destination Address)
Submit-MigWizMailboxMigration -CompanyName Company -DestinationAddress user@example.org -MigrationType Full -Credentials $Credentials

EXAMPLE: Import Users from CSV and Submit for Full Migration
Submit-MigWizMailboxMigration -CompanyName Company -ImportUsers $ImportCSV -MigrationType Full -Credentials $Credentials

EXAMPLE: Import Users from CSV and Submit for Pre-Stage Migration
Submit-MigWizMailboxMigration -CompanyName Company -ImportUsers $ImportCSV -MigrationType Trial -Credentials $Credentials

EXAMPLE: Import Users from CSV and Submit for Verification Migration
Submit-MigWizMailboxMigration -CompanyName Company -ImportUsers $ImportCSV -MigrationType Verification -Credentials $Credentials

EXAMPLE: Specify Username ONLY for Credentials - Submit ONE Users for in a Customer to Full Migration (with Source Email Address)
Submit-MigWizMailboxMigration -CompanyName Company -SourceAddress user@example.org -MigrationType Full -Credentials $Credentials

#>
function Submit-MigWizMailboxMigration {
    param ()
    #Import Module
    Import-MigrationWizModule  

    if ($Global:ProjectName)
    {
        Write-Host "Gathering Details for Project $($ProjectName)." -ForegroundColor Cyan -NoNewline
        $allProjects = Get-MW_MailboxConnector -Ticket $global:mwTicket -name $ProjectName
    }
    if ($Global:ProjectKeyword) {
        Write-Host "Gathering Details for Projects with keyword $($Global:ProjectKeyword)." -ForegroundColor Cyan -NoNewline
        $customer = Get-BT_Customer -Ticket $global:btTicket -CompanyName $CompanyName
        $allProjects = Get-MW_MailboxConnector -Ticket $global:mwTicket -OrganizationId $customer.OrganizationId | ?{$_.name -like "*$ProjectKeyword*"}
    }
    else
    {
        Write-Host "Gathering All Project Details for $($CompanyName). " -ForegroundColor Cyan -NoNewline
        $customer = Get-BT_Customer -Ticket $global:btTicket -CompanyName $CompanyName
        $allProjects = Get-MW_MailboxConnector -Ticket $global:mwTicket -OrganizationId $customer.OrganizationId | sort name
    }

    # Get Mailboxes across all projects
    Write-Host "Gathering Mailbox(es) Migration Details ..." -ForegroundColor Cyan
    $global:allMigMailboxes = @()
    $global:allMigMailboxes = Get-MW_Mailbox -Ticket $global:mwTicket -ConnectorId $allProjects.id -RetrieveAll

    #Submit Migration if the Source Address Supplied
    if ($Global:SourceAddress)
    {
        $MailboxIds = $global:allMigMailboxes | ?{$_.ExportEmailAddress -eq $SourceAddress}
        foreach ($MailboxId in $MailboxIds) {
            Submit-MailboxMigrations -MailboxDetails $MailboxId
        }
        Write-Host "Exported Results to $($Global:OutputCSVFilePath)"
    }
    #Submit Migration if the Destination Address Supplied
    elseif ($Global:DestinationAddress)
    {
        $MailboxIds = $global:allMigMailboxes | ?{$_.ImportEmailAddress -eq $DestinationAddress}
        foreach ($MailboxId in $MailboxIds) {
            Submit-MailboxMigrations -MailboxDetails $MailboxId
        }
        Write-Host "Exported Results to $($Global:OutputCSVFilePath)"
    }
    #Submit Migration if the Imported Users Supplied
    elseif ($global:ImportUserList) {
        $progressref = $global:ImportUserList.count
        $progresscounter = 0
        foreach ($miguser in $global:ImportUserList) {
            #Progress Bar
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            $MailboxIds = $global:allMigMailboxes | ?{$_.ExportEmailAddress -eq $miguser.SourceEmailAddress}
            Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Submtitting $($miguser.SourceEmailAddress) - $($global:MigrationType) Migration in $($Global:ProjectKeyword) Projects"
            
            #Submit Migration
            foreach ($MailboxId in $MailboxIds) {
                Submit-MailboxMigrations -MailboxDetails $MailboxId
            }
            
        }
        Write-Host ""
        Write-Host "Exported Results to $($Global:OutputCSVFilePath)"
    }
    #Submit Migration if No Specific Users Supplied
    elseif ($Global:ALL) {
        $progressref = $global:allMigMailboxes.count
        $progresscounter = 0
        foreach ($miguser in $global:allMigMailboxes) {
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            $MailboxIds = $global:allMigMailboxes | ?{$_.ExportEmailAddress -eq $miguser.SourceEmailAddress}
            Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Submtitting $($miguser.SourceEmailAddress) - $($global:MigrationType) Migration in Projects"
            
            foreach ($MailboxId in $MailboxIds) {
                Submit-MailboxMigrations -MailboxDetails $MailboxId
            }
        }
        Write-Host ""
        Write-Host "Exported Results to $($Global:OutputCSVFilePath)"
    }
}
function Import-MigrationWizModule() {
    if (((Get-Module -Name "BitTitanPowerShell") -ne $null) -or ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null))
    {
        return;
    }

    $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
    $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
    foreach ($moduleLocation in $moduleLocations)
    {
        if (Test-Path $moduleLocation)
        {
            Import-Module -Name $moduleLocation
            return
        }
    }
    Write-Error  "BitTitanPowerShell module was not loaded. Go to https://help.bittitan.com/hc/en-us/articles/115008108267-Install-the-BitTitan-SDK to download the SDK PowerShell"
}
function Export-MigrationFailureReport {
    param (
        [Parameter(Mandatory=$True)][array]$MailboxDetails
    )
    $ProjectName = (Get-MW_MailboxConnector -Ticket $global:mwTicket -Id $MailboxDetails.ConnectorId).Name
    $currentFailure = new-object PSObject

    $currentFailure | add-member -type noteproperty -name "MigrationUpdateDate" -Value $MailboxDetails.UpdateDate
    $currentFailure | add-member -type noteproperty -name "Mailbox" -Value $MailboxDetails.ExportEmailAddress.ToString()
    $currentFailure | add-member -type noteproperty -name "MigrationMailboxID" -Value $MailboxDetails.ID
    $currentFailure | add-member -type noteproperty -name "MigrationProject" -Value $ProjectName.ToString()
    $currentFailure | add-member -type noteproperty -name "Status" -Value "Failed"
    $currentFailure | add-member -type noteproperty -name "MigrationType" -Value $global:MigrationType
    $currentFailure | add-member -type noteproperty -name "MigrationItemTypes" -Value $result.ItemTypes
    $currentFailure | add-member -type noteproperty -name "MigrationLicenseSku" -Value $result.LicenseSku
    $currentFailure | add-member -type noteproperty -name "Exception_Message" -Value "$($_.Exception.Message)"

    $currentFailure | Export-Csv -NoTypeInformation -encoding utf8 -Path $Global:OutputCSVFilePath -Append
}

function Submit-MailboxMigrations {
    param (
        [Parameter(Mandatory=$True)][array]$MailboxDetails
    )   

    if ($global:MigrationType -eq "Trial") {
        try
        {
            $ProjectName = (Get-MW_MailboxConnector -Ticket $global:mwTicket -Id $MailboxDetails.ConnectorId).Name
            Write-Host "Submitting $($global:MigrationType) Migration for $($MailboxDetails.ExportEmailAddress) older than $($DaysOlderThan) in $($ProjectName)" -ForegroundColor Cyan -NoNewline
            $result = Add-MW_MailboxMigration -Ticket $global:mwTicket -MailboxId $MailboxDetails.id -Type Full -ItemTypes Mail -ItemEndDate ((Get-Date).AddDays(-$DaysOlderThan)) -ConnectorId $MailboxDetails.ConnectorId -UserId $global:mwTicket.UserId -ea stop
            
            $currentResult = new-object PSObject
            $currentResult | add-member -type noteproperty -name "MigrationUpdateDate" -Value $result.UpdateDate
            $currentResult | add-member -type noteproperty -name "Mailbox" -Value $MailboxDetails.ExportEmailAddress.ToString()
            $currentResult | add-member -type noteproperty -name "MigrationMailboxID" -Value $result.MailboxID
            $currentResult | add-member -type noteproperty -name "MigrationProject" -Value $ProjectName.ToString()
            $currentResult | add-member -type noteproperty -name "Status" -Value $result.Status
            $currentResult | add-member -type noteproperty -name "MigrationType" -Value $result.Type
            $currentResult | add-member -type noteproperty -name "MigrationItemTypes" -Value $result.ItemTypes
            $currentResult | add-member -type noteproperty -name "MigrationLicenseSku" -Value $result.LicenseSku
            $currentResult | add-member -type noteproperty -name "Exception_Message" -Value $null
            $currentResult | Export-Csv -NoTypeInformation -encoding utf8 -Path $Global:OutputCSVFilePath -Append
        }
        catch
        {
            Write-Error "Unable to Submit Migration Job."
            Export-MigrationFailureReport -MailboxDetails $MailboxDetails
        } 
    }
    else {
        try
        {
            $ProjectName = (Get-MW_MailboxConnector -Ticket $global:mwTicket -Id $MailboxDetails.ConnectorId).Name
            Write-Host "Submitting $($global:MigrationType) Migration for $($MailboxDetails.ExportEmailAddress) in $($ProjectName).." -NoNewline -ForegroundColor Cyan
            $result = Add-MW_MailboxMigration -Ticket $global:mwTicket -MailboxId $MailboxDetails.id -Type $global:MigrationType -ConnectorId $MailboxDetails.ConnectorId -UserId $global:mwTicket.UserId -ErrorAction Stop
    
            $currentResult = new-object PSObject
            $currentResult | add-member -type noteproperty -name "MigrationUpdateDate" -Value $result.UpdateDate
            $currentResult | add-member -type noteproperty -name "Mailbox" -Value $MailboxDetails.ExportEmailAddress.ToString()
            $currentResult | add-member -type noteproperty -name "MigrationMailboxID" -Value $result.MailboxID
            $currentResult | add-member -type noteproperty -name "MigrationProject" -Value $ProjectName.ToString()
            $currentResult | add-member -type noteproperty -name "Status" -Value $result.Status
            $currentResult | add-member -type noteproperty -name "MigrationType" -Value $result.Type
            $currentResult | add-member -type noteproperty -name "MigrationItemTypes" -Value $result.ItemTypes
            $currentResult | add-member -type noteproperty -name "MigrationLicenseSku" -Value $result.LicenseSku
            $currentResult | add-member -type noteproperty -name "Exception_Message" -Value $null
            $currentResult | Export-Csv -NoTypeInformation -Encoding utf8 -Path $Global:OutputCSVFilePath -Append

            Write-Host "Completed" -ForegroundColor Green
        }
        catch
        {
            Write-Error "Unable to Submit Migration Job."
            Export-MigrationFailureReport -MailboxDetails $MailboxDetails
        }
    }
}

Submit-MigWizMailboxMigration