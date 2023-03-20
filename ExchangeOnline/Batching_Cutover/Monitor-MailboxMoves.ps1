#
# This function takes a batch name and monitors the status of the migrations it contains and
# will refresh this info every 60 seconds by default.
# 
#
# $BatchName - By default it adds a wildcard character to the end of the batch name, so for example, if you had
#			   three batches you wanted to monitor called "Batch1", "Batch1.1", and "Batch1.2", you could
# 			   simply specify "Batch1" as the batch name and it will monitor all three of those batches.
#
# $Detailed - If you include this switch it will display a more detailed status of the three migrations that are
#			  farthest from being completed. However, the more migrations in an "InProgress" state, the longer it will take to return.
#			  This functionality also automatically kicks in when there are three or fewer migrations "InProgress" which doesn't take long to return.
#
# $RefreshSeconds - This is the number of seconds between refreshes. It's default value is 60.
#


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
	
	Write-Host
	Write-Host "Initializing. Standby..." -ForegroundColor Yellow
	
	while ($true) {
	
		[array]$migUsersMonitor = Get-MigrationUser | Where {$_.BatchId -like "$BatchName*"}
		[array]$moves = Get-MoveRequest -BatchName "MigrationService:$BatchName*" -ResultSize Unlimited -ErrorAction SilentlyContinue
		$completedCount = @($moves | Where {$_.Status -like "Completed*"}).Count
		$inProgressCount = @($moves | Where {$_.Status -eq "InProgress"}).Count
		$syncedCount = @($moves | Where {$_.Status -eq 'Synced'}).Count
		$queuedCount = @($moves | Where {$_.Status -eq 'Queued'}).Count
		
		$failedMoves = @($moves | Where {$_.Status -like "Failed*"})
		$failedMigUsers = $migUsersMonitor | Where {$_.Status -eq "Failed"}
		[array]$failedMoveGuids = $failedMoves.ExchangeGuid.Guid
		
		foreach ($failedMigUser in $failedMigUsers)
		{
			if ($failedMoveGuids -notcontains $failedMigUser.MailboxGuid.Guid)
			{
				$failedMoveGuids += $failedMigUser.MailboxGuid.Guid
			}
		}
		
		Clear-Host
		Write-Host
		Write-Host "Batch: " -BackgroundColor DarkGray -ForegroundColor Black -NoNewline 
		Write-Host "$BatchName " -BackgroundColor DarkGray -ForegroundColor White
		Write-Host
		Write-Host "Completed:`t $completedCount" -ForegroundColor Green
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
		
		Start-Sleep -Seconds $RefreshSeconds
	}
}
