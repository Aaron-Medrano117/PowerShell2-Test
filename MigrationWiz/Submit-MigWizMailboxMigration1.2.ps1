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
Submit-MigWizMailboxMigration -CompanyName Company -SourceAddress user@example.org -MigrationType Full -Credentials username@example.org

#>
param (
    [Parameter(Mandatory=$True)] [string] $CompanyName,
    [Parameter(Mandatory=$false)] [string] $ProjectName,
    [Parameter(Mandatory=$True,HelpMessage="Choose between Full, Trial (Pre-Stage), or Verification Migration Types")] [string] $MigrationType,
    [Parameter(Mandatory=$True)] 
    [System.Management.Automation.PSCredential] 
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()] $Credentials,
    [Parameter(Mandatory=$false,HelpMessage="Specify the Source Email Address of the Migrating User")] [string] $SourceAddress,
    [Parameter(Mandatory=$false,HelpMessage="Specify the Destination Email Address of the Migrating User")] [string] $DestinationAddress,
    [Parameter(Mandatory=$false)] [array] $ImportUsers,
    [Parameter(Mandatory=$false)] [switch] $ALL
    )
function Submit-MigWizMailboxMigration {  
    #Import Module
    Import-MigrationWizModule

    #$Credentials = Get-Credential
    $mwTicket = Get-MW_Ticket -Credentials $Credentials
    $btTicket = Get-BT_Ticket -Credentials $Credentials -ServiceType BitTitan

    if ($ProjectName)
    {
        Write-Host "Gathering Details for Project $($ProjectName) .." -ForegroundColor Cyan -NoNewline
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -name $ProjectName
    }
    else
    {
        Write-Host "Gathering All Project Details for $($CompanyName) .." -ForegroundColor Cyan -NoNewline
        $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
    }

    # Get Mailboxes across all projects
    Write-Host "Gathering All Mailboxes ..." -ForegroundColor Cyan
    $allMigMailboxes = @()
    $allMigMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $allProjects.id -RetrieveAll

    #Submit Migration if the Source Address Supplied
    if ($SourceAddress)
    {
        $MailboxId = $allMigMailboxes | ?{$_.ExportEmailAddress -eq $SourceAddress}
        if ($MigrationType -eq "Full")
        {
            ### Submit Full Migration
            Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
        }   
            ### Submit Pre-stage Migration
        elseif ($MigrationType -eq "Trial")
        {
            Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
        }
        elseif ($MigrationType -eq "Verification")
        {
            Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
        }
    }
    #Submit Migration if the Destination Address Supplied
    elseif ($DestinationAddress)
    {
        $MailboxId = $allMigMailboxes | ?{$_.ImportEmailAddress -eq $DestinationAddress}
        if ($MigrationType -eq "Full")
        {
            ### Submit Full Migration
            Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
        }   
            ### Submit Pre-stage Migration
        elseif ($MigrationType -eq "Trial")
        {
            Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
        }
        elseif ($MigrationType -eq "Verification")
        {
            Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
        }
    }
    #Submit Migration if the Imported Users Supplied
    elseif ($ImportUsers)
    {
        foreach ($miguser in $ImportUsers)
        {
            $MailboxId = $allMigMailboxes | ?{$_.ExportEmailAddress -eq $miguser.SourceEmailAddress}
            if ($MigrationType -eq "Full")
            {
                ### Submit Full Migration
                Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
            }   
            ### Submit Pre-stage Migration
            elseif ($MigrationType -eq "Trial")
            {
                Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
            }
            elseif ($MigrationType -eq "Verification")
            {
                Get-SubmissionResults -MailboxDetails $MailboxId -Type $MigrationType
            }
        } 
    }
    #Submit Migration if No Specific Users Supplied
    else
    {
            ### Submit Full Migration
        if ($MigrationType -eq "Full")
        {
            foreach ($migmailbox in $allMigMailboxes)
            {
                Get-SubmissionResults -MailboxDetails $migmailbox -Type $MigrationType
            }
        }   
            ### Submit Pre-stage Migration
        elseif ($MigrationType -eq "Trial")
        {
            foreach ($migmailbox in $allMigMailboxes)
            {
                Get-SubmissionResults -MailboxDetails $migmailbox -Type $MigrationType
            }
        }
        elseif ($MigrationType -eq "Verification")
            {
                foreach ($migmailbox in $allMigMailboxes)
                {
                    Get-SubmissionResults -MailboxDetails $migmailbox -Type $MigrationType
                }
            }
    }
}
function Import-MigrationWizModule()
{
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
    
    Write-Error "BitTitanPowerShell module was not loaded"
}
function Get-SubmissionResults {
    param (
        $Type,
        [array]$MailboxDetails
    )
    try
    {
        Write-Host "Submitting $($MigrationType) Migration for $($MailboxDetails.ExportEmailAddress) with ID:$($MailboxDetails.Id) .." -NoNewline
        $result = Add-MW_MailboxMigration -Ticket $mwTicket -MailboxId $MailboxDetails.id -Type $Type -ConnectorId $MailboxDetails.ConnectorId -UserId $mwTicket.UserId -ea silentlycontinue -ErrorVariable $migErrorVariable
        Write-Host " done." -ForegroundColor Green
    }
    catch
    {
        throw
        Write-Host "failed to submit job"
    }
}

Submit-MigWizMailboxMigration