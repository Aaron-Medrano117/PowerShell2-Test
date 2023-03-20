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
                $ExpectedTotalComplete,
                [int]
                $RefreshSeconds = 60
            )
   
    Write-Host
    Write-Host "Initializing. Standby..." -ForegroundColor Yellow
   
    while ($true) {
   
        $timer =  [System.Diagnostics.Stopwatch]::StartNew()
        [array]$migUsersMonitor = Get-MigrationUser -ResultSize Unlimited | Where {$_.BatchId -like "$BatchName*"}
        [array]$moves = Get-MoveRequest -BatchName "MigrationService:$BatchName*" -ResultSize Unlimited -ErrorAction SilentlyContinue
        $timer.Stop()
        $completedCount = @($moves | Where {$_.Status -like "Completed*"}).Count
        $syncedCount = @($moves | Where {$_.Status -eq 'Synced'}).Count
        $queuedCount = @($moves | Where {$_.Status -eq 'Queued'}).Count
       
        $inProgressCount = @($moves | Where {$_.Status -eq "InProgress"}).Count
        if ($inProgressCount -le 6 -or $Detailed)
        {
            $inProgressCount = 0
            foreach ($move in @($moves | Where {$_.Status -eq "InProgress"}))
            {
                $moveStats = $move | Get-MoveRequestStatistics
                if ($moveStats.CompleteAfter -gt (Get-Date) -and $moveStats.StatusDetail.Value -eq "Synced" -or $moveStats.StatusDetail.Value -eq "IncrementalSync")
                {
                    $syncedCount++
                }
                else
                {
                    $inProgressCount++
                }
            }
        }
       
        [array]$failedMoves = $moves | Where {$_.Status -like "Failed*"}
        [array]$failedMigUsers = $migUsersMonitor | Where {$_.Status -eq "Failed"}
        [array]$failedMoveGuids = $failedMoves.ExchangeGuid.Guid
       
        foreach ($failedMigUser in $failedMigUsers)
        {
            if ($failedMoveGuids -notcontains $failedMigUser.MailboxGuid.Guid -or $failedMigUser.MailboxGuid.Guid -eq "00000000-0000-0000-0000-000000000000")
            {
                $failedMoveGuids += $failedMigUser.MailboxGuid.Guid
            }
           
            $failedMigUsers = $failedMigUsers | Where {$_ -ne $failedMigUser}
        }
       
        Clear-Host
        Write-Host
        Write-Host "Batch: " -BackgroundColor DarkGray -ForegroundColor Black -NoNewline
        Write-Host "$BatchName " -BackgroundColor DarkGray -ForegroundColor White
        Write-Host       
        Write-Host "Completed:`t $completedCount" -ForegroundColor Green -NoNewline
       
        if ($ExpectedTotalComplete -and ($completedCount -lt $ExpectedTotalComplete))
        {
            Write-Host " of $ExpectedTotalComplete" -ForegroundColor DarkGreen -NoNewline
        }
        Write-Host
        Write-Host "In Progress:`t $inProgressCount" -ForegroundColor Yellow
        Write-Host "Synced:`t`t $syncedCount" -ForegroundColor Cyan
        Write-Host "Failed:`t`t $($failedMoveGuids.Count)" -ForegroundColor Red
        Write-Host "Queued:`t`t $queuedCount" -ForegroundColor DarkCyan
        Write-Host "Waiting:`t $($migUsersMonitor.Count - $moves.Count)" -ForegroundColor DarkCyan
        Write-Host "Remaining:`t $($migUsersMonitor.Count - $completedCount) of $($migUsersMonitor.Count)" -ForegroundColor Gray
       
        if ($Detailed -or ($inProgressCount -le 3))
        {
            Write-Host
            $moves | Where {$_.Status -eq "InProgress"} | select -ExpandProperty ExchangeGuid | select -ExpandProperty Guid | Get-MoveRequestStatistics | sort PercentComplete | select -First 3 | ft
        }
        else
        {
            Write-Host
        }
       
        Write-Host "Last update:`t$(Get-Date -Format 'hh:mm:ss tt')" -ForegroundColor DarkGray
        Write-Host
       
        $calculatedSleepTime = $RefreshSeconds - [math]::Round($timer.Elapsed.TotalSeconds)
        if ($calculatedSleepTime -lt 1)
        {
            $calculatedSleepTime = 0
        }
       
        Start-Sleep -Seconds $calculatedSleepTime
    }
}