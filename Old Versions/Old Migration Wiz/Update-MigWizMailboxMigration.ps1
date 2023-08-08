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
function Update-MigWizMailboxMigration {
    param (
    [Parameter(Mandatory=$True,HelpMessage="Specify Company Name Projects are associated")] [string] $CompanyName,
    [Parameter(Mandatory=$false)] [string] $ProjectName,
    [Parameter(Mandatory=$True)] 
    [System.Management.Automation.PSCredential] 
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()] $Credentials,
    [Parameter(Mandatory=$True,HelpMessage="Specify Import CSV File Path")] [array] $ImportUsers,
    [Parameter(Mandatory=$True,HelpMessage="Specify If updating the Source or Destination Email Addresses")] [string] $UpdateType,
    [Parameter(Mandatory=$false,HelpMessage="Specify Project KeyWords")] [string] $ProjectKeywords
    )
    #Import Module
    Import-MigrationWizModule

    #Set Variables
    if ($UpdateType -eq "Destination") {
        continue
    }
    elseif ($UpdateType -eq "Source") {
        continue
    }
    else {
        Write-Error "Missing Required Option. Please specify if updating Source or Destination Email Addresses. Please choose one of the following: Source or Destination"
        exit
    }
    Write-Host "Running Updates for $($UpdateType) addresses in Migration Wiz Projects"

    #$Credentials = Get-Credential
    $mwTicket = Get-MW_Ticket -Credentials $Credentials
    $btTicket = Get-BT_Ticket -Credentials $Credentials -ServiceType BitTitan
    #Gather Projects
    if ($ProjectName)
    {
        Write-Host "Gathering Details for Project $($ProjectName) .." -ForegroundColor Cyan
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -name $ProjectName
    }
    elseif ($ProjectKeywords)
    {
        Write-Host "Gathering All $($ProjectKeywords) Project Details for $($CompanyName) .." -ForegroundColor Cyan
        $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | ?{$_.name -like "*$ProjectKeywords*"} | sort name
    }
    else
    {
        Write-Host "Gathering All Project Details for $($CompanyName) .." -ForegroundColor Cyan
        $customer = Get-BT_Customer -Ticket $btTicket -CompanyName $CompanyName
        $allProjects = Get-MW_MailboxConnector -Ticket $mwTicket -OrganizationId $customer.OrganizationId | sort name
    }
    
    ### Update mailbox Import address for ALL Projects
    $ImportCSVUsers = Import-csv $ImportUsers
    $updatedUsers = @()
    $notUpdatedUsers = @()
    $updatedUsers = @()
    $notUpdatedUsers = @()

    foreach ($connector in $allProjects) {
        Write-Host "Gathering All Mailboxes for $($connector.name)" -ForegroundColor Cyan
        $allMigMailboxes = @()
        $allMigMailboxes = Get-MW_Mailbox -Ticket $mwticket -ConnectorId $connector.id -RetrieveAll

        foreach ($user in $ImportCSVUsers) {
            #$sourceEmailAddress = $user.PrimarySMTPAddress_Source.tostring()
            #$destinationEmailAddress = $user.PrimarySmtpAddress_Destination.tostring()
            $sourceEmailAddress = $user.PrimarySmtpAddress_Spectra.tostring()
            $destinationEmailAddress = $user.PrimarySmtpAddress_OVG.tostring()

            #Update Destination Email Address
            if ($UpdateType -eq "Destination") {
                if (!($allMigMailboxes.importemailaddress -contains $destinationEmailAddress)) {
                    try {
                        $mailboxMigration = $allMigMailboxes | ? {$_.ExportEmailAddress -eq $sourceEmailAddress}
                        $result = Set-MW_Mailbox -ticket $mwTicket -ConnectorId $connector.id -mailbox $mailboxMigration -ImportEmailAddres $destinationEmailAddress
                        $updatedUsers += $result
                        Write-host "Successfully updated" $result.ImportEmailAddress "in $($connector.name)" -foregroundcolor green
                        }
                    catch {
                        Write-Warning -Message "$($_.Exception)"
                        Write-Warning -Message "Unable to Update $($mailboxMigration.ImportEmailAddress)" 
                    }
                }
                else {
                    Write-host $destinationEmailAddress "does not need to be updated" -foregroundcolor yellow
                    $notUpdatedUsers += $user
                }
            }
            #Update Source Email Address
            elseif ($UpdateType -eq "Source") {
                if (!($allMigMailboxes.ExportEmailAddress -contains $sourceEmailAddress)) {
                    try {
                        $mailboxMigration = $allMigMailboxes | ? {$_.ImportEmailAddress -eq $destinationEmailAddress}
                        $result = Set-MW_Mailbox -ticket $mwTicket -ConnectorId $connector.id -mailbox $mailboxMigration -ExportEmailAddres $destinationEmailAddress
                        $updatedUsers += $result
                        Write-host "Successfully updated" $result.ExportEmailAddress "in $($connector.name)" -foregroundcolor green
                        }
                    catch {
                        Write-Warning -Message "$($_.Exception)"
                        Write-Warning -Message "Unable to Update $($mailboxMigration.ExportEmailAddress)" 
                    }
                }
                else {
                    Write-host $sourceEmailAddress "does not need to be updated" -foregroundcolor yellow
                    $notUpdatedUsers += $user
                }
            }
        }
    }
}