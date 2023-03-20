$tenantmailusers

$Results =()
foreach ($user in $tenantmailusers) {
    Write-Host "Updating ExternalAddress for" $user.DisplayName "..." -ForegroundColor Cyan -NoNewline
    $ExternalEmailAddress = $user.PrimarySmtpAddress
    try    {
        Set-MailUser $user.alias -ExternalEmailAddress $ExternalEmailAddress -ea stop
        Write-Host "Success" -ForegroundColor Green
        $tmp = "" | Select User, Results
        $tmp.User = $user.DisplayName
        $tmp.Results = "Successful"
        $Results += $tmp
    }
    catch 
    {
        # Exception is stored in the automatic variable Set-MailUser -External
        Write-Host "failed" -ForegroundColor red
        $tmp = "" | Select User, Results
        $tmp.User = $user.DisplayName
        $tmp.Results = "failed"
        $Results += $tmp
        
    }
}