Param (
    [String]
    $AdminUPN
)

function Initialize-ExchangeOnlinePowerShell
{
    <#

        .SYNOPSIS
            Used to initialize the Exchange Online PowerShell session.

        .DESCRIPTION
            Verifies that the pre-requisites are met and initializes the Exchange Online PowerShell session.

        .OUTPUTS
            None

        .EXAMPLE
            Initialize-ExchangeOnlinePowerShell

    #>

    [CmdletBinding()]
    param
    (
        # The UPN to use to connect to Exchange Online Powershell
        [String]
        $AdminUPN
    )

    [bool]$connectedToAzureAD = $false
    $activity = "Initialize Exchange Online PowerShell"
    
    if (-not (Get-ExchangeOnlineSession))
    {
        $exchangeOnlineSession = Import-ExchangeOnlineModernModule -AdminUPN $AdminUPN
    }

    if (Get-ExchangeOnlineSession)
    {
        Write-Host "Successfully connected Exchange Online PowerShell."
        $connectedToAzureAD = $true
    }
    else
    {
        Write-Host "Failed to connect Exchange Online PowerShell."
    }

    $connectedToAzureAD
}

function Get-ExchangeOnlineSession
{
    <#

        .SYNOPSIS
            Gets an open Exchange Online PowerShell session if one exists.

        .DESCRIPTION
            Gets an open Exchange Online PowerShell session if one exists.

        .OUTPUTS
            None

        .EXAMPLE
            Get-ExchangeOnlineSession

    #>

    $openSession = (Get-PSSession | Where-Object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -and ($_.State -eq 'Opened') })[0]

    if ($openSession -notlike $null)
    {
        Write-Host "Found an existing Exchange Online session in the open state."
        $openSession
    }
    else
    {
        Write-Host "Did not find an existing Exchange Online session in the open state."
    }
}

function Import-ExchangeOnlineModernModule
{
    <#

        .SYNOPSIS
            Used to initialize the Exchange Online PowerShell session.

        .DESCRIPTION
            Verifies that the pre-requisites are met and initializes the Exchange Online PowerShell session.

        .OUTPUTS
            None

        .EXAMPLE
            Import-ExchangeOnlineModernModule

    #>

    [CmdletBinding()]
    param
    (
        # The UPN to use to connect to Exchange Online Powershell
        [String]
        $AdminUPN
    )

    $modulePath = Get-ChildItem -Path $env:USERPROFILE\AppData\Local\Apps\2.0\ -Filter "Microsoft.Exchange.Management.ExoPowershellModule.manifest" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($modulePath -notlike $null)
    {
        Write-Host "Found the modern EXO PowerShell module path: $($modulePath.FullName)"

        try
        {
            $module =  Join-Path $modulePath.Directory.FullName "Microsoft.Exchange.Management.ExoPowershellModule.dll"
            Import-Module -FullyQualifiedName $module -Force
            $moduleScript =  Join-Path $modulePath.Directory.FullName "CreateExoPSSession.ps1"
            . $moduleScript
            Connect-EXOPSSession -UserPrincipalName $AdminUPN
            $exchangeOnlineSession = Get-ExchangeOnlineSession
            $exchangeOnlineSession
        }
        catch
        {
            Write-Host "Failed to connect to Exchange Online PowerShell.  $($_.Exception.Message)"
        }
    }
    else
    {
        Write-Host "Failed to find the Exchange Online PowerShell module.  $($_.Exception.Message)"
    }
}

Write-Host "Initializing Exchange PowerShell"
$location = Get-Location
Initialize-ExchangeOnlinePowerShell
Set-Location $location
Write-Host "Loading Batches"
$batches = Get-MigrationBatch
Write-Host "Select a Batch"

$x = 0
foreach ($batch in $batches)
{
    Write-Host "$x`: $($batch.Identity.Name)"
    $x++
}

$batch = $null
while($null -like $batch)
{
    [int]$batchNumber = Read-Host "Batch:"

    if ($batchNumber -ge 0 -and $batchNumber -le $batches.count)
    {
        $batch = $batches[$batchNumber]
    }
}

$dir = "$($batch.identity.Name)"
mkdir $dir -ErrorAction SilentlyContinue
Write-Host "Loading MigrationUsers for $($batch.identity.Name)"
$migrationUsers = Get-MigrationUser -batchid $batch.BatchGuid
    
foreach ($user in $migrationUsers)
{
    $targetFile = "./$dir/$($user.MailboxEmailAddress)`.log"
    if (Test-Path $targetFile)
    {
        Write-Host "Skipping $($user.MailboxEmailAddress), already exported"
    }
    else
    {
        Write-Host "Exporting $($user.MailboxEmailAddress)"
        $moveRequest = Get-MoveRequest $user.MailboxGuid
        ($moveRequest | Get-MoveRequestStatistics -IncludeReport | Select-Object -ExpandProperty Report) | Add-Content $targetFile -Encoding UTF8
    }
}
