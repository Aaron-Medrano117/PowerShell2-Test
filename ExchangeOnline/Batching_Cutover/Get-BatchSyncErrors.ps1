#
# This function takes a batch name and then pulls out a list of failed migrations.
# Some of those are failed migration users, some are failed move requests, and some are both.
# so the code will only generate a unique list of failed users, regardless of failure type.
#

function Get-BatchSyncErrors
{
	[CmdletBinding()]
    param   (
				[Parameter(Mandatory=$false)]
                [string]
                $BatchName
            )
			
	if (-not ($BatchName))
	{
		Write-Host
		Write-Host "Enter the batch name: " -ForegroundColor Cyan -NoNewline
		$BatchName = (Read-Host).Trim()
	}
	
	Write-Host "Locating batch '$BatchName' ... " -ForegroundColor Cyan -NoNewline
	if (-not ($migBatch = Get-MigrationBatch $BatchName -EA SilentlyContinue))
	{
		Write-Host "not found" -ForegroundColor Red
		return
	}
	Write-Host "found" -ForegroundColor Green
	
	Write-Host "Gathering a list of migration users ... " -ForegroundColor Cyan -NoNewline
	[array]$migUsers = Get-MigrationUser -BatchId $BatchName
	[array]$moves = Get-MoveRequest -BatchName "MigrationService:$BatchName" -ResultSize Unlimited
	Write-Host "done" -ForegroundColor Green
	Write-Host
	
	$failedMoves = @($moves | Where {$_.Status -like "Failed*"})
	$failedMigUsers = $migUsers | Where {$_.Status -eq "Failed"}
	[array]$failedMoveGuids = $failedMoves.ExchangeGuid.Guid
	
	foreach ($failedMigUser in $failedMigUsers)
	{
		if ($failedMoveGuids -notcontains $failedMigUser.MailboxGuid.Guid)
		{
			$failedMoveGuids += $failedMigUser.MailboxGuid.Guid
		}
	}
	
	if ($failedMoveGuids.Count -eq 0)
	{
		Write-Host "There are no failed migrations in batch '$BatchName'." -ForegroundColor Green
		Write-Host
		return
	}
	elseif ($failedMoveGuids.Count -ge 10)
	{
		Write-Host "Found $($failedMoveGuids.Count) failures. This will take a few moments to collect those errors. Standby..." -ForegroundColor Yellow
	}
	else
	{
		Write-Host "Collecting $($failedMoveGuids.Count) failures. Standby..." -ForegroundColor Yellow
	}
	
	$failedSyncMailboxes = @()
	foreach ($guid in $failedMoveGuids)
	{
		$tmp = "" | select Mailbox, Error
		
		if ($failedMigUser = $failedMigUsers | Where {$_.MailboxGuid.Guid -eq $guid})
		{
			$tmp.Mailbox = $failedMigUser.Identity
			$tmp.Error = $failedMigUser.ErrorSummary
			$failedSyncMailboxes += $tmp
		}
		elseif ($failedMove = $failedMoves | Where {$_.ExchangeGuid.Guid -eq $guid})
		{
			$migUser = Get-MigrationUser -MailboxGuid $failedMove.ExchangeGuid.Guid
			$failedMoveStats = $failedMove | Get-MoveRequestStatistics
			$tmp.Mailbox = $migUser.Identity
			$tmp.Error = $failedMoveStats.LastFailure.FailureType
			$failedSyncMailboxes += $tmp
		}
	}
	
	Write-Host
	return ($failedSyncMailboxes | sort Error, Mailbox)
}

