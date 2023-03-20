### Stop Sync of MRG.Corp ONPremUPN.

$mailboxProperties = Import-Csv $HOME\Desktop\MailboxProperties.csv

#$enabledMailUsers = @()
$failedMailUsers = @()
$UpdateUPNUsers = @()
$mailboxProperties | Where {$_.OnPremUPN -like "*mrg.corp"} | foreach {

	Write-Host "$($_.OnPremUPN)..." -ForegroundColor Cyan -NoNewline
	try
	{
        $UPN = $_.OnPremUPN
        $NewUPN = $_.HEXUPN
        if ($ADUser = Get-ADUser -filter {UserPrincipalName -eq $UPN}) {
		    $ADUser | Set-ADUser -UserPrincipalName $NewUPN -ErrorAction Stop
		    $UpdateUPNUsers += $_
		    Write-Host "done" -ForegroundColor Green
        else {
        Write-Host "already updated"
            }
        }	    
    }
	catch
	{
		$failedMailUsers += $_
		Write-Host "error" -ForegroundColor Red
	}
}