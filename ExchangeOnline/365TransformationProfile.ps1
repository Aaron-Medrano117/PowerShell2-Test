function Connect-CustomerO365
{
                [CmdletBinding()]
    param   (
                [switch]
                $Existing,
                                                                [string]
                                                                $Username,
                                                                [string]
                                                                $Password,
                                                                [switch]
                                                                $IncludeSkype,
                                                                [switch]
                                                                $IncludeTeams
            )
                                                
                Get-PSSession | Remove-PSSession
                                                
                if ($Existing)
                {
                                $customer = Get-CustomerSelection
                                $creds = New-Object System.Management.Automation.PSCredential($customer.Username,($customer.Password | ConvertTo-SecureString -AsPlainText -Force))
                }
                else
                {
                                if ($Username -and $Password)
                                {
                                                $creds = New-Object System.Management.Automation.PSCredential($Username,($Password | ConvertTo-SecureString -AsPlainText -Force))
                                }
                                elseif ($Username)
                                {
                                                $creds = Get-Credential -Credential $Username
                                }
                                else
                                {
                                                $creds = Get-Credential
                                }
                }
                
                Write-Host
                Write-Host "Connecting to Office 365..." -ForegroundColor Yellow
                Write-Host
                
                $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $creds -Authentication "Basic" -AllowRedirection
                Import-PSSession $exchangeSession -AllowClobber -WarningAction SilentlyContinue
                
                Connect-MsolService -Credential $creds | Out-Null
                Connect-AzureAD -Credential $creds | Out-Null
                $ogName = Get-OrganizationConfig | select -ExpandProperty DisplayName
                $defaultDomain = Get-AcceptedDomain | Where {$_.Default} | select -ExpandProperty DomainName
                $host.UI.RawUI.WindowTitle = "$ogName ($defaultDomain)"
                
                if ($IncludeSkype)
                {
                                Import-Module SkypeOnlineConnector
                                $sfbSession = New-CsOnlineSession -Credential $creds
                                Import-PSSession $sfbSession
                }
                
                if ($IncludeTeams)
                {
                                Import-Module MicrosoftTeams
                                $teamsSession = Connect-MicrosoftTeams -Credential $creds
                                Import-PSSession $teamsSession
                }
}

function Get-BatchSyncErrors
{              
                [CmdletBinding()]
    param   (
                                                                [Parameter(Mandatory=$true)]
                [string]
                $BatchName
            )
                
                Verify-TenantConnection
                
                $failedSyncMailboxes = @()
                
                Get-MigrationUser -BatchId $BatchName | Where {$_.Status -eq "Failed"} | sort ErrorSummary, Identifier | foreach {
                
                                $tmp = "" | select Mailbox, Error
                                $tmp.Mailbox = $_.Identifier
                                $tmp.Error = $_.ErrorSummary
                                $failedSyncMailboxes += $tmp
                }
                
                return $failedSyncMailboxes
}

function Get-CustomerSelection
{
                $scriptDir = $PROFILE.Replace("\Microsoft.PowerShell_profile.ps1","")
                $existingCreds = Import-Csv "$scriptDir\CustomerCredentials.csv" | Where {$_.Customer -and $_.Username -and $_.Password -and $_.Enabled -eq 1} | sort Customer
                
                Clear-Host
                Write-Host
                Write-Host "Select a customer:" -ForegroundColor White
                $n = 1
                $existingCreds | foreach {
                                Write-Host $("{0:D2}" -f $n) -ForegroundColor Yellow -NoNewline
                                Write-Host " - " -ForegroundColor DarkGray -NoNewline
                                Write-Host $_.Customer -ForegroundColor Cyan
                                $n++
                }
                
                Write-Host
                Write-Host "Selection: " -ForegroundColor Green -NoNewline
                [int]$selection = (Read-Host).Trim()
                
                return $existingCreds[$selection - 1]
}

function Monitor-MailboxMoves
{
                [CmdletBinding()]
    param   (
                                                                [Parameter(Mandatory=$true)]
                [string]
                $BatchName,
                                                                [switch]
                                                                $Detailed,
                                                                [int]
                                                                $RefreshSeconds = 60
            )
                
                Verify-TenantConnection
                
                while ($true) {
                                [array]$moves = Get-MoveRequest -BatchName "MigrationService:$BatchName*" -ResultSize Unlimited -ErrorAction SilentlyContinue
                                $completedCount = @($moves | Where {$_.Status -like 'Completed*'}).Count
                                $inProgressCount = @($moves | Where {$_.Status -like 'InProgress'}).Count
                                $queuedCount = @($moves | Where {$_.Status -like 'Synced' -or $_.Status -like "Queued"}).Count
                                $failedCount = @($moves | Where {$_.Status -like 'Failed*'}).Count
                                Clear-Host
                                Write-Host
                                Write-Host "Batch: " -BackgroundColor DarkGray -ForegroundColor Black -NoNewline 
                                Write-Host "$BatchName " -BackgroundColor DarkGray -ForegroundColor White
                                Write-Host
                                Write-Host "Completed:`t $completedCount" -ForegroundColor Green
                                Write-Host "In Progress:`t $inProgressCount" -ForegroundColor Yellow
                                Write-Host "Synced:`t`t $queuedCount" -ForegroundColor Cyan
                                Write-Host "Failed:`t`t $failedCount" -ForegroundColor Red
                                Write-Host "Remaining:`t $($moves.Count - $completedCount) of $($moves.Count)" -ForegroundColor Gray
                                
                                if ($Detailed -or ($inProgressCount -le 3))
                                {
                                                Write-Host
                                                $moves | Where {$_.Status -like 'InProgress'} | Get-MoveRequestStatistics | sort PercentComplete | select -First 3 | ft
                                }
                                else
                                {
                                                Write-Host
                                }
                                
                                Write-Host "Last update:`t$(Get-Date -Format 'hh:mm:ss tt')" -ForegroundColor DarkGray
                                Write-Host
                                Start-Sleep -Seconds $RefreshSeconds
                }
}

function Schedule-MailboxCutover
{
                [CmdletBinding()]
    param   (
                [string]
                $Mailbox,
                                                                [string]
                                                                $DateTime,
                                                                [bool]
                                                                $Confirm = $true
            )
                
                Verify-TenantConnection
                
                if (!$Mailbox)
                {
                                Write-Host "Mailbox: " -ForegroundColor Cyan -NoNewline
                                $Mailbox = (Read-Host).Trim()
                }
                
                if (!$DateTime)
                {
                                Write-Host "Enter the date/time (ex: 1/1/2020 3:00 PM): " -ForegroundColor Cyan -NoNewline
                                $DateTime = (Read-Host).Trim()
                }
                
                if (-not ($migUser = Get-MigrationUser $Mailbox.Trim()))
                {
                                Write-Host "Unalbe to location migration user for $mailbox." -ForegroundColor Red
                                return
                }
                
                if (-not ($cutoverDateTime = (Get-Date $DateTime).ToUniversalTime()))
                {                              
                                Write-Host "Unable to convert '$DateTime' to a valid format." -ForegroundColor Red
                                return
                }
                else
                {
                                $now = Get-Date
                                if ($cutoverDateTime -lt $now)
                                {
                                                Write-Host "The date/time you specified is in the past." -ForegroundColor Yellow
                                                Write-Host "Press 'Enter' if that's ok." -ForegroundColor Cyan
                                                $nothing = Read-Host
                                }
                                elseif ($cutoverDateTime -gt $now.AddDays(7))
                                {
                                                Write-Host "The date/time you specified is more than a week in the future." -ForegroundColor Yellow
                                                Write-Host "Press 'Enter' if that's ok." -ForegroundColor Cyan
                                                $nothing = Read-Host
                                }
                }
                
                if ($Confirm)
                {
                                Write-Host
                                Write-Host "Please confirm:" -ForegroundColor Green
                                Write-Host "Mailbox: " -ForegroundColor Gray -NoNewline
                                Write-Host $Mailbox -ForegroundColor Cyan
                                Write-Host "Date/time: " -ForegroundColor Gray -NoNewline
                                Write-Host $cutoverDateTime.ToLocalTime().ToString() -ForegroundColor Cyan
                                Write-Host
                                Write-Host "Press Enter if we're good." -ForegroundColor Yellow -NoNewline
                                $nothing = Read-Host
                }
                
                $migUser | Set-MigrationUser -CompleteAfter $cutoverDateTime -Confirm:$false
}

function Verify-TenantConnection
{
                $openSessions = Get-PSSession | Where {$_.State -eq "Opened" -and $_.ConfigurationName -eq "Microsoft.Exchange"}
                
                if (!$openSessions)
                {
                                Connect-CustomerO365
                }
}
