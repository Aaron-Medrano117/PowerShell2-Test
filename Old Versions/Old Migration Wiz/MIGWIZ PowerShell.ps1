#MIGWIZ PowerShell
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
    
    Write-Error "BitTitanPowerShell module was not loaded. Go to https://help.bittitan.com/hc/en-us/articles/115008108267-Install-the-BitTitan-SDK to download the SDK PowerShell"
}

Import-MigrationWizModule

Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll'

# Get both mw and bt ticket
$cred = Get-Credential
$mwTicket = Get-MW_Ticket -Credentials $cred
$btTicket = Get-BT_Ticket -Credentials $cred -ServiceType BitTitan

# Get customer, filtered by company name
$customer = Get-BT_Customer -Ticket $btTicket -CompanyName 'company-name-here'

##Get Connector details
$connector = Get-MW_MailboxConnector -Ticket $mwTicket -Name "project name"
#example
$connector = Get-MW_MailboxConnector -Ticket $mwTicket -Name "Fanduel 6 | T2T | AM"

#Grab all connectors under a customer
$allConnectors = Get-MW_MailboxConnector -Ticket $mwTicket -ExportConfigurationChecksum $connector.ExportConfigurationChecksum

#Grab all MigrationMailboxes in a Connector (this is primarily what you see on the full list of mailboxes page including Export and Import addresses)
$Mailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $connector.id -RetrieveAll

#Gather all the Migration Jobs for all the Mailboxes (ugly version and doesn't show exactly which mailbox it is tied to)
$mailboxmigrations = Get-MW_MailboxMigration -Ticket $mwticket -ConnectorId $Connector.Id -Mailboxid $Mailboxes.ID -RetrieveAll

## Get all Connectors for a company
$customer = Get-BT_Customer -Ticket $btTicket -CompanyName 'company-name-here'
Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId

#Update Export Email Address for one User
Set-MW_Mailbox -ticket $mwTicket -ConnectorId $connector.id -mailbox $mailbox -ExportEmailAddres $newExportAddress
#Update Export Email Address for one User (no output)
$result = Set-MW_Mailbox -ticket $mwTicket -ConnectorId $connector.id -mailbox $mailbox -ExportEmailAddres $newExportAddress

# Update Export Email Addresses
foreach ($mailbox in $Mailboxes)
{
    if ($mailbox.ExportEmailAddress -like "*@betfairprod.onmicrosoft.com")
    {
        $newExportAddress = $mailbox.ExportEmailAddress.replace("@betfairprod.onmicrosoft.com","@paddypowerbetfair.com")
        $result = Set-MW_Mailbox -ticket $mwTicket -ConnectorId $connector.id -mailbox $mailbox -ExportEmailAddres $newExportAddress
    }
    else
    {
        Write-host $mailbox.ExportEmailAddress "does not need to be updated"
    }
}


### Update mailbox export address for ALL connectors

$updatedUsers = @()
$notUpdatedUsers = @()

foreach ($connector in $allConnectors)
{
    $FDGMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $connector.id -RetrieveAll

    foreach ($mailbox in $allMailboxes)
    {
        if ($mailbox.ExportEmailAddress -like "*@betfairprod.onmicrosoft.com")
        {
            $newExportAddress = $mailbox.ExportEmailAddress.replace("@betfairprod.onmicrosoft.com","@paddypowerbetfair.com")
            $result = Set-MW_Mailbox -ticket $mwTicket -ConnectorId $connector.id -mailbox $mailbox -ExportEmailAddres $newExportAddress
            $updatedUsers += $result
            Write-host "Successfully updated" $result.ExportEmailAddress "in" $connector.name -foregroundcolor green
        }
        else
        {
            Write-host $mailbox.ExportEmailAddress "does not need to be updated" -foregroundcolor yellow
            $notUpdatedUsers += $mailbox
        }
    }
}

## Gather all mailboxes for across all projects

$allMigWizMailboxes = @()

foreach ($connector in $allConnectors)
{
    $Mailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $connector.id -RetrieveAll
    $allMigWizMailboxes += $allMailboxes
}


#### Check Mailbox Migration Status for a single Project

$connector = Get-MW_MailboxConnector -Ticket $mwTicket -Name "FanDuel 6 | T2T | AM"
$allMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $connector.id -RetrieveAll

$MailboxProjectStatistics = @()
foreach ($mailbox in $allMailboxes | sort ExportEmailAddress)
{
    $mailboxmigrations = Get-MW_MailboxMigration -Ticket $mwticket -ConnectorId $Connector.Id -Mailboxid $mailbox.id -RetrieveAll | sort CompleteDate
    Write-host "Found $($mailboxmigrations.count) migrations for $($mailbox.ExportEmailAddress). Gathering Details .." -foregroundcolor cyan -NoNewline

    $MailboxProjectStat = New-Object PSObject
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ExportEmailAddress" -Value $mailbox.ExportEmailAddress
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ImportEmailAddress" -Value $mailbox.ImportEmailAddress
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "StartDate" -Value $mailboxmigrations[-1].StartDate
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "CompleteDate" -Value $mailboxmigrations[-1].CompleteDate
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Status" -Value $mailboxmigrations[-1].Status
    $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureMessage" -Value $mailboxmigrations[-1].FailureMessage

    $MailboxProjectStatistics += $MailboxProjectStat
    Write-host "done" -foregroundcolor green
}

$MailboxProjectStatistics

### Gather Mailbox Details for all Projects

function Get-MigWizMailboxReport {
    param (
        [Parameter(Mandatory=$True)] [string] $CompanyName,
        [Parameter(Mandatory=$false)][string] $ProjectName
    )

    # Get Mailbox Connector(s)
    $MailboxProjectStatistics = @()
    $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName

    if ($ProjectName)
    {
        $AllConnectors = Get-MW_MailboxConnector -Ticket $mwTicket -name $ProjectName
    }
    else
    {
        $AllConnectors = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
    }

    # Get Mailboxes per connector
    foreach ($connector in $AllConnectors)
    {
        $allMailboxes = @()
        Write-host "Gathering all mailboxes for project $($connector.name). " -foregroundcolor cyan -NoNewline
        $allMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $connector.id -RetrieveAll
        $ALLMailboxes2 += $allMailboxes
        Write-host "$($ALLMailboxes2.count) users found. " -foregroundcolor cyan -NoNewline
        Write-host "Pulling Mailbox Migrations." -foregroundcolor cyan -NoNewline

        #Gathering Last Mailbox Project Status per Mailbox
        foreach ($mailbox in $ALLMailboxes2 | sort ExportEmailAddress)
        {
            Write-host "." -foregroundcolor gray -NoNewline
            $mailboxmigrations = Get-MW_MailboxMigration -Ticket $mwticket -ConnectorId $Connector.Id -Mailboxid $mailbox.id -RetrieveAll | sort CompleteDate

            $MailboxProjectStat = New-Object PSObject
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "CustomerName" -Value $customer.CompanyName
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "PrimaryDomain" -Value $customer.PrimaryDomain
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Project" -Value $connector.name
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceType" -Value $connector.ExportConfiguration.ExchangeItemType
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceEmailAddress" -Value $mailbox.ExportEmailAddress
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationType" -Value $connector.ImportConfiguration.ExchangeItemType
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationEmailAddress" -Value $mailbox.ImportEmailAddress
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "CompleteDate" -Value $mailboxmigrations[-1].CompleteDate
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Status" -Value $mailboxmigrations[-1].Status
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Type" -Value $mailboxmigrations[-1].Type
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ItemTypes" -Value $mailboxmigrations[-1].ItemTypes
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureMessage" -Value $mailboxmigrations[-1].FailureMessage

            $MailboxProjectStatistics += $MailboxProjectStat
        }
        Write-host "done" -foregroundcolor green
    }
    Write-host "Exported Migration Report" -foregroundcolor cyan
    $MailboxProjectStatistics | Export-Csv -NoTypeInformation -Encoding utf8 "$HOME\Desktop\\MailboxMigrationStatistics.csv"
}

#### ### Gather Mailbox Details for all Projects Attempt 2

function Get-MigWizMailboxReport2 {
    param (
        [Parameter(Mandatory=$True)] [string] $CompanyName,
        [Parameter(Mandatory=$false)] [string] $OutputCSVFilePath,
        [Parameter(Mandatory=$false)][string] $ProjectName
    )

    if ($ProjectName)
    {
        $AllConnectors = Get-MW_MailboxConnector -Ticket $mwTicket -name $ProjectName
    }
    else
    {
        $AllConnectors = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
    }

    <#retrieve ticket
    if ($mwticket.expirationdate -lt (Get-Date))
    {
        Write-host "Ticket expired Need to validate credential exists"
        if (!($credentials))
        {
            Write-host "Need New Credential"
            $credentials = Get-Credential          
        }
        else
        {
            
            Write-host "Credential Already Exists. New Ticket created. Expiration date $($credential.expirationdate)"
        }       
    }
    #>
    #$mwticket = Get-MW_Ticket -Credentials $credentials
    #$btTicket = Get-BT_Ticket -Credentials $credentials -ServiceType BitTitan

    # Get Mailbox Connector(s)
    $MailboxProjectStatistics = @()
    $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName
    $AllConnectors = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
    
    # Get Mailboxes per connector
    $allMigMailboxes = @()
    foreach ($connector in $AllConnectors)
    {
        Write-host "Gathering all mailboxes for project $($connector.name). " -foregroundcolor cyan -NoNewline
        $allMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $connector.id -RetrieveAll
        $allMigMailboxes += $allMailboxes
        Write-host "$($allMigMailboxes.count) users found. " -foregroundcolor cyan -NoNewline
        Write-host "Pulling Mailbox Migrations." -foregroundcolor cyan -NoNewline
        
        #Gathering Last Mailbox Project Status per Mailbox
        foreach ($mailbox in $allMigMailboxes | sort ExportEmailAddress)
        {
            Write-host "." -foregroundcolor gray -NoNewline
            $mailboxmigrations = Get-MW_MailboxMigration -Ticket $mwticket -Mailboxid $mailbox.id -SortBy_CreateDate_Descending -PageSize 1
            
            $MailboxProjectStat = New-Object PSObject
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "CustomerName" -Value $customer.CompanyName
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "PrimaryDomain" -Value $customer.PrimaryDomain
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Project" -Value $connector.name
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceType" -Value $connector.ExportConfiguration.ExchangeItemType
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationType" -Value $connector.ImportConfiguration.ExchangeItemType
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceEmailAddress" -Value $mailbox.ExportEmailAddress
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationEmailAddress" -Value $mailbox.ImportEmailAddress
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "CompleteDate" -Value $mailboxmigrations.CompleteDate
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Status" -Value $mailboxmigrations.Status
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ItemTypes" -Value $mailboxmigrations.ItemTypes
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureMessage" -Value $mailboxmigrations.FailureMessage

            $MailboxProjectStatistics += $MailboxProjectStat
        }
        Write-host "done" -foregroundcolor green
    }

    if ($OutputCSVFilePath)
    {
        $MailboxProjectStatistics | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFilePath
    }
    else
    {
        $OutputCSVFilePath = Read-Host "Where do you wish to export the file?"
        $csvFileName = $CompanyName + "_MigrationStatistics.csv"
        Write-host "Exported Migration Report" -foregroundcolor cyan
        $MailboxProjectStatistics | Export-Csv -NoTypeInformation -Encoding utf8 "$($OutputCSVFilePath)\$($csvFileName)"
    } 
}

# Update Export Email Addresses
foreach ($MWUser in $tmpmigupdates)
{
    $newExportAddress = $MWUser.old
    $mailbox = $FDMailboxes | ?{$_.ExportEmailAddress -eq $MWUser.export}
    $result = Set-MW_Mailbox -ticket $mwTicket -ConnectorId $connector.id -mailbox $mailbox -ExportEmailAddress $newExportAddress
    Write-host "updated Export address for $($MWUser.Export)" -foregroundcolor green
    #Read-Host
}

#####

$customer = Get-BT_Customer -Ticket $btTicket -CompanyName (Read-Host -Prompt "Customer")
$projects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId

# Retrieve items under each project
$projects | ForEach {
    Write-Host("`n Project: `"$($_.Name)`"")
    
    # Retrieve project items
    $projectItems = Get-MW_Mailbox -Ticket $mwTicket -ConnectorId $_.Id
    if ( -not $projectItems ) { 
        Write-Host("`t0 items")
    }
    else {
        # Retrieve the last migration submitted for each item
        $projectItems | ForEach {
            $projectItemMigration = Get-MW_MailboxMigration -Ticket $mwTicket -MailboxId $_.Id -SortBy_CreateDate_Descending -PageSize 1
            
            # Print result
            if ( -not $projectItemMigration ) { 
                Write-Host("`t $($_.ExportEmailAddress): No migrations")
            }
            else {
                Write-Host("`t $($_.ExportEmailAddress): Last migration: $($projectItemMigration.CreateDate), $($projectItemMigration.Status)")
            }            
        }
    }
}


## remove user from migration

$duplicatemigmailboxes
$fanduelgroupOnMicrosoftAddresses = $duplicatemigmailboxes | ?{$_.domain -eq "fanduelgroup.onmicrosoft.com"}

foreach ($duplicatemigmbx in $fanduelgroupOnMicrosoftAddresses | sort project)
{
    $connector = Get-MW_MailboxConnector -Ticket $mwTicket -name $duplicatemigmbx.project

    # Get Mailboxes per connector
    $allMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $connector.id -RetrieveAll

    $mwmbx = $allMailboxes | ? {$_.importemailaddress -eq $duplicatemigmbx.DestinationEmailAddress}

    Write-host "Removing user $($duplicatemigmbx.DestinationEmailAddress) from $($duplicatemigmbx.project). " -foregroundcolor cyan -nonewline
    Remove-MW_Mailbox -Ticket $mwTicket -Id $mwmbx.id -force
    Write-host "done." -foregroundcolor green
    #Read-Host "pause to check"
}

### Submit Full Migration
foreach ($migmailbox in $ProctorUInternalUser34Afternoon)
{
    #Retrieve MailboxID
    $ExportSearchAddress = $migmailbox."Email Address"
    $MailboxId = $Mailboxes | ?{$_.ExportEmailAddress -eq $ExportSearchAddress}

    Write-Host "Checking item" $ExportSearchAddress "with ID:" $MailboxId.Id 
    $result = Add-MW_MailboxMigration -Ticket $mwTicket -MailboxId $MailboxId.Id -Type Full -ConnectorId $MailboxId.ConnectorId -UserId $mwTicket.UserId 
}

### Check status in shell
$customer = Get-BT_Customer -Ticket $btTicket -CompanyName (Read-Host -Prompt "Customer")
$projects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId
$allMigMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $projects.id -RetrieveAll

$PUMailboxProjectStatistics = @()
$notfoundmigrations = @()
foreach ($mailbox in $ProctorUInternalUser34Afternoon | sort "Email Address")
{
    $ExportSearchAddress = $mailbox."Email Address"
    if ($MailboxId = $allMigMailboxes | ?{$_.ExportEmailAddress -eq $ExportSearchAddress})
    {
        Write-host "." -foregroundcolor gray -NoNewline
        $connectordetails = Get-MW_MailboxConnector -Id $MailboxId.ConnectorId -Ticket $mwTicket -ea silentlycontinue
        $mailboxmigrations = Get-MW_MailboxMigration -Ticket $mwticket -Mailboxid $MailboxId.id -SortBy_CreateDate_Descending -PageSize 1 -ea silentlycontinue
        
        $MailboxProjectStat = New-Object PSObject
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ProjectName" -Value $connectordetails.Name -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceEmailAddress" -Value $MailboxId.ExportEmailAddress -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationEmailAddress" -Value $MailboxId.ImportEmailAddress -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "CompleteDate" -Value $mailboxmigrations.CompleteDate -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Status" -Value $mailboxmigrations.Status -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ItemTypes" -Value $mailboxmigrations.ItemTypes -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureMessage" -Value $mailboxmigrations.FailureMessage -force
    }
    else
    {
        Write-Host "No user found for $($mailbox.DisplayName)"
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ProjectName" -Value "" -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceEmailAddress" -Value $ExportSearchAddress -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationEmailAddress" "" -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "CompleteDate" -Value "" -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Status" -Value "NotFound" -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ItemTypes" -Value "" -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureMessage" -Value "" -force
        $notfoundmigrations += $Mailbox
    }
    
    $PUMailboxProjectStatistics += $MailboxProjectStat
}
$PUMailboxProjectStatistics