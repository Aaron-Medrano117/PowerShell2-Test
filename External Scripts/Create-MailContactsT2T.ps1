## Create Mail Contacts script

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory=$True)] [string] $ImportCSV,
    [Parameter(Mandatory=$True)] [string] $FailureOutputCSV
)
function Create-MailContacts {
    param (
        [Parameter(Mandatory=$True)] [string] $ImportCSV,
        [Parameter(Mandatory=$True)] [string] $FailureOutputCSV
    )
    $mailboxes = Import-Csv $ImportCSV
    #ProgressBar
    $progressref = ($mailboxes).count
    $progresscounter = 0
    $alreadyExists = @()
    $createdUsers = @()
    $failures = @()

    foreach ($user in $mailboxes) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Creating Mailboxes for $($user.DisplayName)"

    if ($recipientCheck = Get-Recipient $user.PrimarySMTPAddress -ea silentlycontinue)
    {
        Write-Host "Contact Already Exists for $($user.PrimarySMTPAddress)" -ForegroundColor Cyan
        $alreadyExists += $user
    }
    else {
        try {
            New-MailContact -DisplayName $user.DisplayName -Name $user.name -ExternalEmailAddress $user.PrimarySMTPAddress -ErrorAction Stop
            Write-Host "Created Contact $($user.PrimarySMTPAddress)" -ForegroundColor Green
            $createdUsers += $user
        }
        catch
        {
            Write-Warning -Message "$($_.Exception)"
            $currenterror = new-object PSObject

            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
            $currenterror | add-member -type noteproperty -name "User" -Value $user.PrimarySMTPAddress
            $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
            $failures += $currenterror

            #Attempt 2 on creating Contact
            $newName = $user.DisplayName + " - Jefferson"
            try {
                New-MailContact -DisplayName $user.DisplayName -Name $newName -ExternalEmailAddress $user.PrimarySMTPAddress -ErrorAction Stop
                Write-Host "Attempt 2 Created Contact $($newName)" -ForegroundColor Yellow
            }
            catch {
                Write-Warning -Message "$($_.Exception)"
                $currenterror = new-object PSObject

                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | add-member -type noteproperty -name "User" -Value $user.PrimarySMTPAddress
                $currenterror | add-member -type noteproperty -name "Exception" -Value $_.Exception
                $failures += $currenterror
            }  
        }
    }
    }
    Write-Host "$($mailusers.count) MailUsers found in list" -ForegroundColor Yellow
    Write-Host "$($notfoundUsers.count) Users Not Found in list" -ForegroundColor Yellow
    Write-Host "$($failures.count) Failures" -ForegroundColor Yellow

    $alreadyExists | Select DisplayName, Name, PrimarySmtpAddress, RecipientTypeDetails | Out-GridView
    $failures | Export-Csv -NoTypeInformation -Encoding utf8 $FailureOutputCSV
}

Create-MailContacts -ImportCSV $ImportCSV -FailureOutputCSV $FailureOutputCSV