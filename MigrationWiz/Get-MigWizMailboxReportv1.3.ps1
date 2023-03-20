<#
SYNOPSIS: Gather list of all Mailbox Projects for a Customer, or particular project

Prerequiesites: You'll need to download the MigrationWiz PowerShell Module and Import the module
https://www.bittitan.com/doc/powershell.html#PagePowerShellintroductionmd-powershell-module-installation
Download from the above link
Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll'

EXAMPLE: Gather all mailbox project stats for a customer
Get-MigWizMailboxReport -CompanyName ProctorU -Credentials $credential

Get-MigWizMailboxReport -CompanyName "FanDuel Group" -Credentials $credential

EXAMPLE: Gather all mailbox project stats for a single project
Get-MigWizMailboxReport - ProjectName "Fanduel 6 | T2T | AM" -Credentials $credential

EXAMPLE: Gather all mailbox project stats for a single project
Get-MigWizMailboxReport -PrimaryDomain example.org -Credentials $credential

EXAMPLE: Gather all mailbox project stats for a single project and specify Output location
Get-MigWizMailboxReport -CompanyName "FanDuel Group" -Credential $credentials -OutputCSVFilePath C:\Users\RSUSER\Desktop\MigrationWizProjectStats.csv 

#>
param (
    [Parameter(Mandatory=$false,HelpMessage="Specify CompanyName from MigrationWiz Customer")] [string] $CompanyName,
    [Parameter(Mandatory=$false,HelpMessage="Specify ProjectName from MigrationWiz Project")] [string] $ProjectName,
    [Parameter(Mandatory=$false,HelpMessage="Specify Project KeyWords")] [string] $ProjectKeywords,
    [Parameter(Mandatory=$false,HelpMessage="Specify PrimaryDomain from MigrationWiz Customer")] [string] $PrimaryDomain,
    [Parameter(Mandatory=$True)] 
    [System.Management.Automation.PSCredential] 
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()] $Credentials,
    [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath
)
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
    
    Write-Error  "BitTitanPowerShell module was not loaded. Go to https://help.bittitan.com/hc/en-us/articles/115008108267-Install-the-BitTitan-SDK to download the SDK PowerShell"
}
function Get-MigWizMailboxReport {
    param ()
    
    #Import Module
    Import-MigrationWizModule
    
    #Gather MigrationWizStatistics
    #Gather Tokens
    $mwTicket = Get-MW_Ticket -Credentials $Credentials
    $btTicket = Get-BT_Ticket -Credentials $Credentials -ServiceType BitTitan

    #Specify Single Project or All Projects for Customer
    try {
        if ($ProjectName) {
            $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -name $ProjectName -ErrorAction stop
            $customer = Get-BT_Customer -Ticket $btTicket -OrganizationId $allProjects.OrganizationId
        }
        elseif ($ProjectKeywords) {
            $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | ?{$_.name -like "*$ProjectKeywords*"} | sort name 
        }
        elseif ($PrimaryDomain) {
            $customer = Get-BT_Customer -Ticket $btTicket -PrimaryDomain $PrimaryDomain -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
        }
        elseif ($CompanyName) {
            $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
        }
        else {
            $CompanyName = Read-Host "What is the CompanyName for MigrationWiz?"
            $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId -ErrorAction stop | sort name
        }
    }
    catch {
        Write-Host "Failed finding MigrationWiz Project. Check Spelling." -ForegroundColor Red
    }
    
    try {
        # Get Mailboxes per connector
        $MailboxProjectStatistics = @()
        $allMigMailboxes = @()
        $allMigMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $allProjects.id -RetrieveAll -ea stop
        Write-host "Gathering all mailboxes for customer $($customer.CompanyName) for domain $($customer.PrimaryDomain). " -foregroundcolor cyan
        Write-host "$($allProjects.count) Projects found - $(($allProjects.name -join ",")). " -foregroundcolor cyan
        
        #Gathering Last Mailbox Project Status per Mailbox
        $progressref = ($allMigMailboxes).count
        $progresscounter = 0
        foreach ($mailbox in $allMigMailboxes | sort ExportEmailAddress) {
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Pulling Mailbox Migration Details for $($mailbox.ImportEmailAddress)"

            $mailboxmigrations = Get-MW_MailboxMigration -Ticket $mwticket -Mailboxid $mailbox.id -SortBy_CreateDate_Descending -PageSize 1
            $connector = Get-MW_MailboxConnector -Ticket $mwTicket -ID $mailbox.ConnectorId
            
            #Check Source Endpoint Name
            if ($connector.ExportConfiguration.ExchangeItemType -eq "Mailbox") {
                $SourceType = "Office 365 Mail"
            }
            else {
                if ($connector.ExportConfiguration.host -eq "imap.gmail.com")
                {
                    $SourceType = "GSUITE IMAP"
                }
                elseif ($connector.ExportConfiguration.host -eq "secure.emailsrvr.com")
                {
                    $SourceType = "Rackspace Secure IMAP"
                }
                elseif ($connector.ExportConfiguration.host -eq "imap.emailsrvr.com")
                {
                    $SourceType = "Rackspace IMAP"
                }
                else
                {
                    $SourceType = $connector.ExportConfiguration.host
                }               
            }

            #Check Destination Endpoint Name
            if ($connector.ImportConfiguration.ExchangeItemType -eq "Mailbox") {
                $DestinationType = "Office 365 Mail"
            }
            else {
                if ($connector.ImportConfiguration.host -eq "imap.gmail.com")
                {
                    $DestinationType = "GSUITE IMAP"
                }
                elseif ($connector.ImportConfiguration.host -eq "secure.emailsrvr.com")
                {
                    $DestinationType = "Rackspace Secure IMAP"
                }
                elseif ($connector.ImportConfiguration.host -eq "imap.emailsrvr.com")
                {
                    $DestinationType = "Rackspace IMAP"
                }
                else
                {
                    $DestinationType = "UnIdentified"
                }               
            }

            #Check MigrationStatus Size
            $ImportSuccessSizeTotals = (Get-MW_MailboxStat -MailboxId $mailbox.id -Ticket $mwTicket | select -ExpandProperty MigrationStatsInfos | ? {$_.TaskType -eq "Import"} | select -ExpandProperty migrationstats).SuccessSizeTotal

            $MailboxProjectStat = New-Object PSObject
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "CompanyName" -Value $customer.CompanyName
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "PrimaryDomain" -Value $customer.PrimaryDomain
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Project" -Value $connector.name
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceType" -Value $SourceType
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceEmailAddress" -Value $mailbox.ExportEmailAddress
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationType" -Value $DestinationType
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationEmailAddress" -Value $mailbox.ImportEmailAddress
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Status" -Value $mailboxmigrations.Status
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SuccessSizeTotal (MB)" -Value (($ImportSuccessSizeTotals | measure -Sum).sum/1000000)
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "TimeStamp" -Value $mailboxmigrations.CompleteDate
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "MigrationType" -Value $mailboxmigrations.Type
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ItemTypes" -Value $mailboxmigrations.ItemTypes
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureMessage" -Value $mailboxmigrations.FailureMessage

            $MailboxProjectStatistics += $MailboxProjectStat
        }
    }
    catch {
        Write-Host "Unable to Pull MigrationStats. Missing Requirements. Please Specify a PrimaryDomain, CompanyName, or a Project Name" -ForegroundColor red
    }
    if ($OutputCSVFilePath) {
		$MailboxProjectStatistics | Export-Csv $OutputCSVFilePath -NoTypeInformation -Encoding UTF8
		Write-host "Exported DL Property List to $OutputCSVFilePath"-ForegroundColor Cyan
	}
	else {
		try {
			$MailboxProjectStatistics | Export-Csv "$HOME\Desktop\MigrationWizReport-Mailboxes.csv" -NoTypeInformation -Encoding UTF8
			Write-host "Exported DL Property List to $HOME\Desktop\MigrationWizReport-Mailboxes.csv" -ForegroundColor Cyan
		}
		catch {
			Write-Warning -Message "$($_.Exception)"
			Write-host ""
			$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
			$MailboxProjectStatistics | Export-Csv "$OutputCSVFolderPath\MigrationWizReport-Mailboxes.csv" -NoTypeInformation -Encoding UTF8
		}
	}
}
Get-MigWizMailboxReport