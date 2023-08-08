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
    [Parameter(Mandatory=$false)] [string] $ExportFilePath,
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

function Get-MigWizMailboxReport {
    param ()
    
    #Import Module
    Import-MigrationWizModule
    
    #Connect to BitTitan - Gather Tokens
    Connect-BitTitan

    #Specify Single Project or All Projects for Customer
    try
    {
        if ($ProjectName)
        {
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -name $ProjectName -ErrorAction stop
            $customer = Get-BT_Customer -Ticket $script:btTicket -OrganizationId $allProjects.OrganizationId
        }
        elseif ($ProjectKeywords)
        {
            $customer = Get-BT_Customer -Ticket $script:btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -OrganizationId $customer.OrganizationId | ?{$_.name -like "*$ProjectKeywords*"} | sort name
        }
        elseif ($PrimaryDomain)
        {
            $customer = Get-BT_Customer -Ticket $script:btTicket -PrimaryDomain $PrimaryDomain -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -OrganizationId $customer.OrganizationId | sort name
        }
        elseif ($CompanyName)
        {
            $customer = Get-BT_Customer -Ticket $script:btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -OrganizationId $customer.OrganizationId | sort name
        }
        else
        {
            $CompanyName = Read-Host "What is the CompanyName for MigrationWiz?"
            $customer = Get-BT_Customer -Ticket $script:btTicket -CompanyName $CompanyName -ErrorAction stop
            $allProjects = Get-MW_MailboxConnector -Ticket $script:mwTicket -OrganizationId $customer.OrganizationId -ErrorAction stop | sort name
        }
    }
    catch
    {
        Write-Host "Failed finding MigrationWiz Project. Check Spelling." -ForegroundColor Red
    }
    
    try {
        # Get Mailboxes per connector
        $MailboxProjectStatistics = @()
        $allMigMailboxes = @()
        $allMigMailboxes = Get-MW_Mailbox -Ticket $script:mwticket -ConnectorId $allProjects.id -RetrieveAll -ea stop
        Write-host "$($allProjects.count) Projects found - $(($allProjects.name -join ",")). " -foregroundcolor cyan
        Write-host "Found $($allMigMailboxes.length.ToString()) mailboxes in Projects " -foregroundcolor cyan
    }
    catch
    {
        Write-Host "Unable to Pull MigrationStats. Missing Requirements. Please Specify a PrimaryDomain, CompanyName, or a Project Name" -ForegroundColor red
        Return
    }

    #Gathering Last Mailbox Project Status per Mailbox
    $progressref = ($allMigMailboxes).count
    $progresscounter = 0
    foreach ($mailbox in $allMigMailboxes | sort ExportEmailAddress)
    {
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Pulling Mailbox Migration Details for $($mailbox.ImportEmailAddress)"

        #Get Migration, Connector, MigrationSize Details
        $mailboxmigrations = Get-MW_MailboxMigration -Ticket $script:mwticket -Mailboxid $mailbox.id -SortBy_CreateDate_Descending -PageSize 1
        $connector = Get-MW_MailboxConnector -Ticket $script:mwTicket -ID $mailbox.ConnectorId
        $ImportTotals = Get-MW_MailboxStat -Ticket $script:mwTicket -MailboxId $mailbox.id  | select -ExpandProperty MigrationStatsInfos | ? {$_.TaskType -eq "Import"} | select -ExpandProperty migrationstats
        
        #Create Migration User Output
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
    }

    if ($ExportFilePath) {
        try {
            if ($OverrideWorksheet) {
                $MailboxProjectStatistics | Export-Excel -path $ExportFilePath -WorksheetName $WorksheetName -ClearSheet
		        Write-host "Exported Migration Stats to $ExportFilePath" -ForegroundColor Cyan
            }
            elseif ($AppendToWorkSheet) {
                $MailboxProjectStatistics | Export-Excel -path $ExportFilePath -WorksheetName $WorksheetName -Append
		        Write-host "Exported Migration Stats to $ExportFilePath" -ForegroundColor Cyan
            }
            else {
                $MailboxProjectStatistics | Export-Excel -path $ExportFilePath -WorksheetName $WorksheetName
		        Write-host "Exported Migration Stats to $ExportFilePath" -ForegroundColor Cyan
            }
        }
        Catch {
            Write-Warning -Message "$($_.Exception)"
			Write-host ""
			$OutputCSVFolderPath = Read-Host 'INPUT Required: Where do you wish to save this file? Please provide full folder path'
            $WorksheetName2 = Read-Host 'INPUT Required: What WorkSheet Name do you wish to Use?'
            $MailboxProjectStatistics | Export-Excel "$OutputCSVFolderPath\MigrationWizReport-Mailboxes.csv" -WorksheetName $WorksheetName2
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
            $MailboxProjectStatistics | Export-Excel "$OutputCSVFolderPath\MigrationWizReport-Mailboxes.csv" -WorksheetName $WorksheetName2
            Write-host "Exported Migration Stats to $OutputCSVFolderPath\MigrationWizReport-Mailboxes.csv" -ForegroundColor Cyan
			
		}
	}
}
Get-MigWizMailboxReport