$batchName = "MigrationService:Batch1*"

while($true)
{
    $allMoves = Get-MoveRequest -BatchName "MigrationService:$batchName"
    

 


    $inProgressStatus = @()
    Write-Host "Getting In-Progress Mailbox Move Status'"

 

    foreach($inProgress in $($allMoves | ?{$_.Status -eq 'InProgress'}))
    {
        $inProgressStatus += ($inProgress | Get-MoveRequestStatistics | Select DisplayName,StatusDetail,BytesTransferredPerMinute,PercentComplete,TotalMailboxSize,Message)
    }

 

    $warnings = $($allMoves | ?{$_.Status -like '*warning'})
    $warningStatus = @()
    foreach($inProgress in $warnings)
    {
        $warningStatus += ($inProgress | Get-MoveRequestStatistics | Select DisplayName,StatusDetail,PercentComplete,TotalMailboxSize,Message)
        
    }

 

    $stalleds = $($allMoves | ?{$_.Status -like '*stalled*'})
    $stalledStatus = @()

 

    foreach($inProgress in $stalleds)
    {
        $stalledStatus += ($inProgress | Get-MoveRequestStatistics | Select DisplayName,StatusDetail,PercentComplete,TotalMailboxSize,Message)
            
    }
    clear
    Write-Host (Get-Date)
    Write-Host "Current Breakdown:"
    $allMoves | Group-Object Status | ft
    Write-Host "In Progress: "
    $inProgressStatus | FT

 

    if($warningStatus.Count -gt 0)
    {
        Write-Host "Completed With Warnings: "
        $warningStatus | FT
    }

 

    if($stalledStatus.Count -gt 0)
    {
        Write-Host "Stalled: "
        $stalledStatus | FT
    }

 

    start-sleep 30
}
