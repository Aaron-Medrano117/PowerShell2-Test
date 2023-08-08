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
    [Parameter(Mandatory=$True,HelpMessage="Specify CompanyName from MigrationWiz Customer")] [string] $CompanyName,
    [Parameter(Mandatory=$false,HelpMessage="Specify ProjectName from MigrationWiz Project")] [string] $ProjectName,
    [Parameter(Mandatory=$false,HelpMessage="Specify Project KeyWords")] [string] $ProjectKeywords,
    [Parameter(Mandatory=$false,HelpMessage="Specify PrimaryDomain from MigrationWiz Customer")] [string] $PrimaryDomain,
    [Parameter(Mandatory=$True)] 
    [System.Management.Automation.PSCredential] 
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()] $Credentials,
    [Parameter(Mandatory=$True)] [string] $ExportFilePath,
    [Parameter(Mandatory=$True)] [string] $WorksheetName,
    [Parameter(Mandatory=$false)] [Switch] $OverrideWorksheet,
    [Parameter(Mandatory=$false)] [Switch] $AppendToWorkSheet
)
$script:creds = $Credentials
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
Function Connect-BitTitan {
    #[CmdletBinding()]

    if((Get-Module PackageManagement)) { 
        #Install Packages/Modules for Windows Credential Manager if required
        If(!(Get-PackageProvider -Name 'NuGet')){
            Install-PackageProvider -Name NuGet -Force
        }
        If((Get-Module PowerShellGet) -and !(Get-Module -ListAvailable -Name 'CredentialManager')){
            Install-Module CredentialManager -Force
            $useCredentialManager = $true
        } 
        else { 
            Import-Module CredentialManager
            $useCredentialManager = $true
        }

        if($useCredentialManager ) {
            # Authenticate
            $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'
        }
    }
    else{
        $useCredentialManager = $false
    }
    
    if(!$script:creds){
        $credentials2 = (Get-Credential -Message "Enter BitTitan credentials")
        if(!$credentials2) {
            $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
            Write-Host -ForegroundColor Red  $msg
        }

        if($useCredentialManager) {
            New-StoredCredential -Target 'https://migrationwiz.bittitan.com' -Persist 'LocalMachine' -Credentials $credentials | Out-Null
            
            $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' stored in Windows Credential Manager."
            Write-Host -ForegroundColor Green  $msg

            $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'

            $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
            Write-Host -ForegroundColor Green  $msg
        }
        else {
            $script:creds = $credentials2
        }
    }
    else{
        $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
    }
    try { 
        # Get a ticket and set it as default
        $script:btTicket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction Stop
        # Get a MW ticket
        $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction Stop 
    }
    catch {
        $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
        $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
        foreach ($moduleLocation in $moduleLocations) {
            if (Test-Path $moduleLocation) {
                Import-Module -Name $moduleLocation

                # Get a ticket and set it as default
                $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
                # Get a MW ticket
                $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 

                if(!$script:ticket -or !$script:mwTicket) {
                    $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
                    Write-Host -ForegroundColor Red  $msg
                    Exit
                }
                else {
                    $msg = "SUCCESS: Connected to BitTitan."
                    Write-Host -ForegroundColor Green  $msg
                }

                return
            }
        }

        $msg = "ACTION: Install BitTitan PowerShell SDK 'bittitanpowershellsetup.msi' downloaded from 'https://www.bittitan.com' and execute the script from there."
        Write-Host -ForegroundColor Yellow $msg
        Write-Host

        Start-Sleep 5

        $url = "https://www.bittitan.com/downloads/bittitanpowershellsetup.msi " 
        $result= Start-Process $url

        Exit
    }  

    if(!$script:btTicket -or !$script:mwTicket) {
        $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Exit
    }
    else {
        $msg = "SUCCESS: Connected to BitTitan."
        Write-Host -ForegroundColor Green  $msg
    }
}

function Write-ProgressHelper {
    param (
        [int]$ProgressCounter,
        [string]$Activity,
        [string]$ID,
        [string]$CurrentOperation,
        [int]$TotalCount,
        [datetime]$StartTime
    )
    #$ProgressPreference = "Continue"  
    if ($ProgressPreference = "SilentlyContinue") {
        $ProgressPreference = "Continue"
    }

    $secondsElapsed = (Get-Date) - $StartTime
    $secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)
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
    #$secondsRemaining = ($secondsElapsed.TotalSeconds / $progresscounter) * ($TotalCount - $progresscounter)

}

function Convert-HashToArray {
    param (
        [Parameter(Mandatory=$true)]
        [Hashtable]$tenantStatsHashToConvert,
        [Parameter(Mandatory=$false)]
        [String]$tenant,
        [Parameter(Mandatory=$false)]
        [String]$table,
        [Parameter(Mandatory=$true)]
        [DateTime]$startTime
        
    )
    Write-Host "Converting Hash to Array for Export..." -ForegroundColor Cyan -nonewline
    $ExportArray = @()
    $start = Get-Date
    $totalCount = $tenantStatsHashToConvert.count    
    $progresscounter = 1

    foreach ($nestedKey in $tenantStatsHashToConvert.Keys) {
        Write-ProgressHelper -Activity "Converting $($nestedKey)" -ProgressCounter ($progresscounter++) -TotalCount $totalCount -StartTime $start
        #Write-Host $($nestedKey)
        $attributes = $tenantStatsHashToConvert[$nestedKey]
        if ($attributes -is [hashtable] -or $attributes -is [System.Collections.Specialized.OrderedDictionary]) {
            #Write-Host "$($nestedKey) is a hashtable or an ordered dictionary."
            #Write-Host "." -foregroundcolor cyan -nonewline
            $customObject = New-Object -TypeName PSObject
            foreach ($attribute in $attributes.keys) {
                $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)" -Value $attributes[$attribute]
            }
            $ExportArray += $customObject
        } 
        elseif ($attributes -is [array] -or $attributes -is [PSCustomObject] ) {
            #Write-Host "." -foregroundcolor Yellow -nonewline
            #Write-Host "$($nestedKey) is an array or custom object."
            $ExportArray += $attributes
        }
        else {
            Write-Host "." -foregroundcolor red -nonewline
            #Write-Host "$($nestedKey) is of type: $($attributes.GetType().FullName)"
        }
    }
    Write-Host "Completed" -ForegroundColor Green
    Write-Progress -Activity "Converting $($nestedKey)" -Completed
    Return $ExportArray
}

function Get-MigWizMailboxReport {
    param (
        [Parameter(Mandatory=$True,HelpMessage="Specify CompanyName from MigrationWiz Customer")] [string] $CompanyName,
        [Parameter(Mandatory=$false,HelpMessage="Specify ProjectName from MigrationWiz Project")] [string] $ProjectName,
        [Parameter(Mandatory=$false,HelpMessage="Specify Project KeyWords")] [string] $ProjectKeywords,
        [Parameter(Mandatory=$false,HelpMessage="Specify PrimaryDomain from MigrationWiz Customer")] [string] $PrimaryDomain,
        [Parameter(Mandatory=$True)]
        [String]$ExportFilePath,
        [Parameter(Mandatory=$True)]
        [String]$WorksheetName
    )
    $initialStart = Get-Date

    #Import Module
    Import-MigrationWizModule
    
    #Connect to BitTitan - Gather Tokens
    Connect-BitTitan

    #Specify Single Project or All Projects for Customer

    #Gather all Customers and convert to Hash
    #$startTime = Get-Date
    $allCustomers = Get-BT_Customer -Ticket $script:btTicket -RetrieveAll
    $allCustomersHash = @{}
    #Write-Host "Completed all Customers in $((Get-Date) - $startTime)." -ForegroundColor Cyan
    foreach ($customer in $allCustomers) {
        $allCustomersHash[$customer.CompanyName.toString()] = $customer
    }

    #Gather All Mailbox Migration Jobs
    try {
        if ($ProjectName)
        {
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -name $ProjectName -ErrorAction stop
            $customer = Get-BT_Customer -Ticket $script:btTicket -OrganizationId $allProjects.OrganizationId -ErrorAction stop
            $AllProjectsHash = @{}
        }
        elseif ($ProjectKeywords)
        {
            $customer = $allCustomersHash[$CompanyName]
            #$customer = Get-BT_Customer -Ticket $script:btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -OrganizationId $customer.OrganizationId -ErrorAction stop | ?{$_.name -like "*$ProjectKeywords*"} | sort name
            $AllProjectsHash = @{}
        }
        elseif ($PrimaryDomain)
        {
            $customer = Get-BT_Customer -Ticket $script:btTicket -PrimaryDomain $PrimaryDomain -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -OrganizationId $customer.OrganizationId -ErrorAction stop | sort name
            $AllProjectsHash = @{}
        }
        elseif ($CompanyName)
        {
            $customer = $allCustomersHash[$CompanyName]
            #$customer = Get-BT_Customer -Ticket $script:btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -OrganizationId $customer.OrganizationId -ErrorAction stop | sort name
            $AllProjectsHash = @{}
        }
        else
        {
            $CompanyName = Read-Host "What is the CompanyName for MigrationWiz?"
            $customer = $allCustomersHash[$CompanyName]
            #$customer = Get-BT_Customer -Ticket $script:btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -OrganizationId $customer.OrganizationId -ErrorAction stop | sort name
            $AllProjectsHash = @{}
        }
    }
    catch {
        Write-Host "Failed finding MigrationWiz Project. Check Spelling of the project, company name, or primary domain.." -ForegroundColor Red
    }
    
    try {
        # Get Mailboxes per connector
        Write-host "Gathering All Project(s) and Details .." -foregroundcolor cyan -nonewline
        $allMigMailboxes = @()
        $allMigMailboxes = Get-MW_Mailbox -Ticket $script:mwticket -ConnectorId $allProjects.id -RetrieveAll -ea stop
        $allSpecifiedMailboxMigrations = Get-MW_MailboxMigration -Ticket $script:mwTicket -RetrieveAll -ConnectorId $allProjects.id -SortBy_CreateDate_Ascending
        $allSpecifiedMailboxMigJobHash = @{}
        Write-host "Completed" -foregroundcolor Green
        Write-host "$($allProjects.count) Projects found - $(($allProjects.name -join ",")). " -foregroundcolor cyan
        Write-host "Found $($allMigMailboxes.length.ToString()) mailboxes in Projects " -foregroundcolor cyan
    }
    catch {
        Write-Host "Unable to Pull MigrationStats. Missing Requirements. Please Specify a PrimaryDomain, CompanyName, or a Project Name" -ForegroundColor red
        Return
    }
    
    <# $startTime = Get-Date
    $allSpecifiedMailboxMigrations = Get-MW_MailboxMigration -Ticket $script:mwTicket -RetrieveAll -ConnectorId $allProjects.id -SortBy_CreateDate_Ascending
    $allSpecifiedMailboxMigrations = @{}
    Write-Host "Completed all MailboxMigrations in $((Get-Date) - $startTime)." -ForegroundColor Cyan
    
    $startTime = Get-Date
    $allConnectors = Get-MW_MailboxConnector -Ticket $script:mwTicket -RetrieveAll
    $AllMailboxConnectorsHash = @{}
    Write-Host "Completed all MailConnectors in $((Get-Date) - $startTime)." -ForegroundColor Cyan
    #>

    #$startTime = Get-Date
    $AllImportStats = Get-MW_MailboxStat -Ticket $script:mwTicket -RetrieveAll -SortBy_CreateDate_Ascending
    $allMailboxMigrationStatsHash = @{}
    #Write-Host "Completed all MailboxStats in $((Get-Date) - $startTime)." -ForegroundColor Cyan

    #Create Hash Tables for Mailbox Migration Jobs, Connectors, and Migration Stats
    foreach ($mailboxmigration in $allSpecifiedMailboxMigrations) {
        $allSpecifiedMailboxMigJobHash[$mailboxmigration.MailboxId.toString()] = $mailboxmigration
    }
    foreach ($project in $allProjects) {
        $AllProjectsHash[$project.id.toString()] = $project
    }
    foreach ($importstat in $AllImportStats) {
        $allMailboxMigrationStatsHash[$importstat.MailboxID.toString()] = $importstat
    }
    

    #Create Specified Mailbox Project Stats HashTables
    $MailboxProjectStatistics = [Ordered]@{}

    # Progress Details
    $progresscounter = 1
    $totalCount = ($allMigMailboxes).count

    #Create Mailbox Project Stats
    foreach ($mailbox in $allMigMailboxes | sort ExportEmailAddress)  {
        Write-ProgressHelper -Activity "Pulling Mailbox Migration Details for $($mailbox.ImportEmailAddress)" -ProgressCounter ($progresscounter++) -TotalCount $TotalCount -StartTime $initialStart

        # Gather Migration Details - Hash Table Query
        $mailboxmigrations = $allSpecifiedMailboxMigJobHash[$mailbox.Id.ToString()]
        $connector = $AllProjectsHash[$mailbox.ConnectorId.ToString()]
        $ImportTotals = $allMailboxMigrationStatsHash[$mailbox.Id.ToString()] | select -ExpandProperty MigrationStatsInfos | ? {$_.TaskType -eq "Import"} | select -ExpandProperty migrationstats
        $latestMailboxError = Get-MW_MailboxError -Ticket $script:mwTicket -MailboxId $mailbox.Id -SortBy_CreateDate_Descending -PageSize 1 -ErrorAction SilentlyContinue
        #skip Modern Auth Warning
        if ($latestMailboxError.Message -eq "Connecting to tenant using modern authentication") {
            $latestMailboxError.Message = $null
            $latestMailboxError.Severity = $null
        }
        else {
            $latestMailboxError.Message = $latestMailboxError.Message
            $latestMailboxError.Severity = $latestMailboxError.Severity
        }
        
        #Create Migration User Output
        $MailboxProjectStat = [ordered]@{
            "ProjectType" = "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)"
            "Project" = $connector.name
            "SourceEmailAddress" = $mailbox.ExportEmailAddress
            "DestinationEmailAddress" = $mailbox.ImportEmailAddress
            "UserMigrationBundleLicense" =  $mailbox.SyncSubscriptionStatus
            "MigrationType" = $mailboxmigrations.Type
            "LastCompletedTimeStamp" = $mailboxmigrations.CompleteDate
            "LastStatus" = $mailboxmigrations.Status
            "SuccessSizeTotal(GB)" = [math]::Round(($ImportTotals.SuccessSizeTotal | measure -Sum).sum/1000000000,3)
            "FailureSizeTotal(GB)" = [math]::Round(($ImportTotals.ErrorSizeTotal | measure -Sum).sum/1000000000,3)
            "SuccessCountTotal" = ($ImportTotals.SuccessCountTotal | measure -Sum).sum
            "FailureCountTotal" = ($ImportTotals.ErrorCountTotal | measure -Sum).sum
            "LatestMessage" = $latestMailboxError.Message
            "LatestMessageSeverity" = $latestMailboxError.Severity
            "FinalFailureMessage" = $mailboxmigrations.FailureMessage 
        }

        $MailboxProjectStatistics[$mailbox.ExportEmailAddress] = $MailboxProjectStat

        <# Gather Migration Details - Individual Mailbox Query
        $mailboxmigrations = Get-MW_MailboxMigration -Ticket $script:mwticket -Mailboxid $mailbox.id -SortBy_CreateDate_Descending -PageSize 1
        $connector = Get-MW_MailboxConnector -Ticket $script:mwTicket -ID $mailbox.ConnectorId
        $ImportTotals = Get-MW_MailboxStat -Ticket $script:mwTicket -MailboxId $mailbox.id  | select -ExpandProperty MigrationStatsInfos | ? {$_.TaskType -eq "Import"} | select -ExpandProperty migrationstats
        $latestMailboxError = Get-MW_MailboxError -Ticket $script:mwTicket -MailboxId $mailbox.Id -SortBy_CreateDate_Descending -PageSize 1 -ErrorAction SilentlyContinue       
        
        $MailboxProjectStat = New-Object PSObject

        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "ProjectType" -Value "$($connector.ProjectType)-$($connector.ExportType)-$($connector.ImportType)" -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "Project" -Value $connector.name -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SourceEmailAddress" -Value $mailbox.ExportEmailAddress -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "DestinationEmailAddress" -Value $mailbox.ImportEmailAddress -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "UserMigrationBundleLicense" -Value $mailbox.SyncSubscriptionStatus -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "LastCompletedTimeStamp" -Value $mailboxmigrations.CompleteDate -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "LastStatus" -Value $mailboxmigrations.Status -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SuccessSizeTotal(MB)" -Value ([math]::Round(($ImportTotals.SuccessSizeTotal | measure -Sum).sum/1000000,3)) -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureSizeTotal(MB)" -Value ([math]::Round(($ImportTotals.ErrorSizeTotal | measure -Sum).sum/1000000,3)) -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "SuccessCountTotal" -Value (($ImportTotals.SuccessCountTotal | measure -Sum).sum) -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FailureCountTotal" -Value (($ImportTotals.ErrorCountTotal | measure -Sum).sum) -force

        #Error Details
        if ($latestMailboxError = Get-MW_MailboxError -Ticket $script:mwTicket -MailboxId $mailbox.Id -SortBy_CreateDate_Descending -PageSize 1 -ErrorAction SilentlyContinue) {
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "LatestMessage" -Value $latestMailboxError.Message -force
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "LatestMessageSeverity" -Value $latestMailboxError.Severity -force
        }
        else {
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "LatestMessage" -Value $null -force
            $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "LatestMessageSeverity" -Value $null -force
        }
        
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "FinalFailureMessage" -Value $mailboxmigrations.FailureMessage -force
        $MailboxProjectStat | Add-Member -MemberType NoteProperty -Name "MigrationType" -Value $mailboxmigrations.Type -force

        $MailboxProjectStatistics += $MailboxProjectStat
        #>
    }
    $start = Get-Date
    $MailboxProjectStatisticsArray = Convert-HashToArray -tenantStatsHash $MailboxProjectStatistics -StartTime $start

    Write-Host "Completed gathering all MailboxStats in $((Get-Date) - $initialStart)." -ForegroundColor Cyan

    if ($ExportFilePath) {
        try {
            if ($OverrideWorksheet) {
                $MailboxProjectStatisticsArray | Export-Excel -path $ExportFilePath -WorksheetName $WorksheetName -ClearSheet
		        Write-host "Exported Migration Stats to $ExportFilePath" -ForegroundColor Cyan
            }
            elseif ($AppendToWorkSheet) {
                $MailboxProjectStatisticsArray | Export-Excel -path $ExportFilePath -WorksheetName $WorksheetName -Append
		        Write-host "Exported Migration Stats to $ExportFilePath" -ForegroundColor Cyan
            }
            else {
                $MailboxProjectStatisticsArray | Export-Excel -path $ExportFilePath -WorksheetName $WorksheetName
		        Write-host "Exported Migration Stats to $ExportFilePath" -ForegroundColor Cyan
            }
        }
        Catch {
            Write-Warning -Message "$($_.Exception)"
			Write-host ""
			$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
            $WorksheetName2 = Read-Host 'INPUT Required: What WorkSheet Name do you wish to Use?'
            $MailboxProjectStatisticsArray | Export-Excel "$OutputCSVFolderPath\MigrationWizReport-Mailboxes.csv" -WorksheetName $WorksheetName2
            Write-host "Exported Migration Stats to $OutputCSVFolderPath\MigrationWizReport-Mailboxes.csv" -ForegroundColor Cyan
        }
	}
	else {
		try {
			$MailboxProjectStatistics | Export-Excel "$HOME\Desktop\MigrationWizReport-Mailboxes.csv" -WorksheetName $WorksheetName
			Write-host "Exported Migration Stats to $HOME\Desktop\MigrationWizReport-Mailboxes.csv" -ForegroundColor Cyan
		}
		catch {
			Write-Warning -Message "$($_.Exception)"
			Write-host ""
			$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
            $WorksheetName2 = Read-Host 'INPUT Required: What WorkSheet Name do you wish to Use?'
            $MailboxProjectStatisticsArray | Export-Excel "$OutputCSVFolderPath\MigrationWizReport-Mailboxes.csv" -WorksheetName $WorksheetName2
            Write-host "Exported Migration Stats to $OutputCSVFolderPath\MigrationWizReport-Mailboxes.csv" -ForegroundColor Cyan
			
		}
	}
}

Get-MigWizMailboxReport -CompanyName $CompanyName -ExportFilePath $ExportFilePath -WorksheetName $WorksheetName