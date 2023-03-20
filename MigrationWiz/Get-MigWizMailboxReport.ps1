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

EXAMPLE: Gather all mailbox project stats for a single project and specify Output location
Get-MigWizMailboxReport -CompanyName "FanDuel Group" -Credential $credentials -OutputCSVFilePath C:\Users\RSUSER\Desktop\MigrationWizProjectStats.csv 

#>
function Get-EndpointNames {
    param ($connector)
                #Check Source Endpoint Name
                if ($connector.ExportConfiguration.ExchangeItemType -eq "Mailbox")
                {
                    $SourceType = "Office 365 Mail"
                }
                else
                {
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
                if ($connector.ImportConfiguration.ExchangeItemType -eq "Mailbox")
                {
                    $DestinationType = "Office 365 Mail"
                }
                else
                {
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
                        $DestinationType = $connector.ImportConfiguration.host
                    }               
                }
}
function Get-MigrationStatistics {
    param ($mailbox,
    $connector,
    $mailboxmigrations)
    #Check Source Endpoint Name
    if ($connector.ExportConfiguration.ExchangeItemType -eq "Mailbox")
    {
        $SourceType = "Office 365 Mail"
    }
    else
    {
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
    if ($connector.ImportConfiguration.ExchangeItemType -eq "Mailbox")
    {
        $DestinationType = "Office 365 Mail"
    }
    else
    {
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
            $DestinationType = $connector.ImportConfiguration.host
        }               
    }

    #Check MigrationStatus Size
    $MailboxProjectStatistics = @()
    $ImportSuccessSizeTotals = Get-MW_MailboxStat -MailboxId $mailbox.id -Ticket $mwTicket | select -ExpandProperty MigrationStatsInfos | ? {$_.TaskType -eq "Import"} | select -ExpandProperty migrationstats
    $ExportSuccessSizeTotals = Get-MW_MailboxStat -MailboxId $mailbox.id -Ticket $mwTicket | select -ExpandProperty MigrationStatsInfos | ? {$_.TaskType -eq "Export"} | select -ExpandProperty migrationstats

    #Gathering Last Mailbox Project Status per Mailbox
    $MailboxProjectStat = New-Object PSObject
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Project" -Value $connector.name
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Type" -Value $mailboxmigrations.Type
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceType" -Value $SourceType
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceEmailAddress" -Value $mailbox.ExportEmailAddress
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationType" -Value $DestinationType
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationEmailAddress" -Value $mailbox.ImportEmailAddress
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Status" -Value $mailboxmigrations.Status
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ImportSuccessSize (GB)" -Value (($ImportSuccessSizeTotals.SuccessSize | measure -Sum).sum/1000000000)
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ImportSuccessSizeTotal (GB)" -Value (($ImportSuccessSizeTotals.SuccessSizeTotal | measure -Sum).sum/1000000000)
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ExportSuccessSize (GB)" -Value (($ExportSuccessSizeTotals.SuccessSize | measure -Sum).sum/1000000000)
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ExportSuccessSizeTotal (GB)" -Value (($ExportSuccessSizeTotals.SuccessSizeTotal | measure -Sum).sum/1000000000)
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "TimeStamp" -Value $mailboxmigrations.CompleteDate
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ItemTypes" -Value $mailboxmigrations.ItemTypes
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureMessage" -Value $mailboxmigrations.FailureMessage

    $MailboxProjectStatistics += $MailboxProjectStat

    if ($OutputCSVFilePath)
    {
        $MailboxProjectStatistics | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
    }
    else
    {
        $MailboxProjectStatistics
    }
}
function Get-MigWizMailboxReport {
    param (
        [Parameter(Mandatory=$false)] [string] $CompanyName,
        [Parameter(Mandatory=$false)] [string] $PrimaryDomain,
        [Parameter(Mandatory=$false)][string] $ProjectName,
        [Parameter(Mandatory=$false)][string] $SourceEmailAddress,
        [Parameter(Mandatory=$false)][string] $DestinationEmailAddress,
        [Parameter(Mandatory=$True)] 
        [System.Management.Automation.PSCredential] 
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()] $Credentials,
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath
    )
    #Global Variables
    $global:OutputCSVFilePath = $OutputCSVFilePath
    #Gather Tokens
    $mwTicket = Get-MW_Ticket -Credentials $Credentials
    $btTicket = Get-BT_Ticket -Credentials $Credentials -ServiceType BitTitan

    #Specify Single Project or All Projects for Customer
    if ($ProjectName)
    {
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -name $ProjectName
    }
    elseif ($PrimaryDomain)
    {
        $customer = Get-BT_Customer -Ticket $btTicket -PrimaryDomain $PrimaryDomain
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
    }
    elseif ($CompanyName)
    {
        $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
    }
    else
    {
        $customer = Get-BT_Customer -Ticket $btTicket -RetrieveAll
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
    }
    
    # Get Mailboxes per connector
    
    $allMigMailboxes = @()
    Write-host "Gathering all mailboxes for project $($allProjects.name). " -foregroundcolor cyan -NoNewline
    $allMigMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $allProjects.id -RetrieveAll

    if ($SourceEmailAddress)
    {
        $MailboxId = $allMigMailboxes | ?{$_.ExportEmailAddress -eq $SourceEmailAddress}
        $mailboxmigrationdetails = Get-MW_MailboxMigration -Ticket $mwticket -MailboxId $MailboxId.id -SortBy_CreateDate_Descending -PageSize 1
        $connectordetails = Get-MW_MailboxConnector -Ticket $mwTicket -ID $MailboxId.ConnectorId
        #Get MigrationStats
        Get-MigrationStatistics -mailbox $mailboxid -connector $connectordetails -mailboxmigrations $mailboxmigrationdetails
    }
    elseif ($DestinationEmailAddress)
    {
        $MailboxId = $allMigMailboxes | ?{$_.ImportEmailAddress -eq $DestinationEmailAddress}
        $mailboxmigrationdetails = Get-MW_MailboxMigration -Ticket $mwticket -MailboxId $MailboxId.id -SortBy_CreateDate_Descending -PageSize 1
        $connectordetails = Get-MW_MailboxConnector -Ticket $mwTicket -ID $MailboxId.ConnectorId
        #Get MigrationStats
        Get-MigrationStatistics -mailbox $mailboxid -connector $connectordetails -mailboxmigrations $mailboxmigrationdetails
    }
    else
    {
        Write-host "$($allMigMailboxes.count) users found. " -foregroundcolor cyan -NoNewline
        Write-host "Pulling Mailbox Migrations." -foregroundcolor cyan -NoNewline
        foreach ($mailbox in $allMigMailboxes | sort ExportEmailAddress)
        {
            $mailboxmigrationdetails = Get-MW_MailboxMigration -Ticket $mwticket -Mailboxid $mailbox.id -SortBy_CreateDate_Descending -PageSize 1
            $connectordetails = Get-MW_MailboxConnector -Ticket $mwTicket -ID $mailbox.ConnectorId
            #Get MigrationStats
            Get-MigrationStatistics -mailbox $mailbox -connector $connectordetails -mailboxmigrations $mailboxmigrationdetails
        }
    }
}