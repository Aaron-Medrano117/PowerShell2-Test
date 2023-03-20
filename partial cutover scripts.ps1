#Convert federated domains to managed
 
$federateddomains = Get-MsolDomain | ?{$_.authentication -eq "Federated"}
 
$federateddomains | foreach {Set-MsolDomainAuthentication -DomainName $_.name -Authentication Managed}
    # Get list of msolusers and proxyaddress of domains to remove
    $domainRemovalUsers += (Get-MsolUser -all | ? {$_.proxyaddresses -like "*@orioninstruments.com" -or $_.proxyaddresses -like "*@introtek.com" -or $_.proxyaddresses -like "*@innovativesensing.com"})


function Get-UnlicensedNonMailObjects {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$True,HelpMessage='Which Domain Migrating?')] [string] $DomainQuery,
        [Parameter(Mandatory=$false,HelpMessage='AllCurrentUsers?')] [switch] $AllCurrentUsers
    )
    #Import Modules
    function Import-ExchangeMsOnlineAzureModule() {
        #Exchange Online Module
        if ((Get-Module -Name "ExchangeOnlineManagement") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to ExchangeOnline ... " -nonewline -foregroundcolor cyan
                    Connect-ExchangeOnline
                    Write-Host "done" -foregroundcolor green
                }
                catch {
                    Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
                }
            }
            else {
                try {
                    Install-Module -Name ExchangeOnlineManagement 
                }
                catch {
                    Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
                }
            }
        }
        #Microsoft Online Module
        if ((Get-Module -Name "MSOnline") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to MSOnline ... " -nonewline -foregroundcolor cyan
                    Connect-MsolService
                    Write-Host "done" -foregroundcolor green
                }
                catch {
                    Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
            else {
                try {
                    Install-Module -Name MSOnline 
                }
                catch {
                    Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
        }
        #Azure AD Module
        if ((Get-Module -Name "AzureAD") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to AzureAD ... " -nonewline -foregroundcolor cyan
                    Connect-AzureAD
                    Write-Host "AzureAD" -foregroundcolor green
                }
                catch {
                    Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
            else {
                try {
                    Install-Module -Name AzureAD 
                }
                catch {
                    Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
        }
        
    }
    Import-ExchangeMsOnlineAzureModule
    # Get list of msolusers and proxyaddress of domains to remove
    $domainQueryString = "@" + $DomainQuery
    $domainRemovalUsers = Get-MsolUser -All -UnlicensedUsersOnly | ? {$_.proxyaddresses -like "*$domainQueryString"}

    #check if they are mail users
    #Progress Bar 1A
    $progressref = ($allMSOLUsers).count
    $progresscounter = 0
    $nonMailUsers = @()
    foreach ($user in $domainRemovalUsers) {
        #Progress Bar 1B
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking $($user.displayname)"
        if (!($recipientCheck = Get-Recipient $user.UserPrincipalName -ea silentlycontinue)) {
            Write-Host "$($user.UserPrincipalName) not found as Mail Recipient" -ForegroundColor Cyan
            $proxyAddresses = $user.ProxyAddresses
            $nonMailUsersCurrent = New-Object PSObject
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $user.DisplayName
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $user.UserPrincipalName
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "ProxyAddresses" -Value ($proxyAddresses -join ",")
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "WhenCreated" -Value $user.WhenCreated
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "LastDirSyncTime" -Value $user.LastDirSyncTime
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "IsLicensed" -Value $user.IsLicensed
            $nonMailUsers += $nonMailUsersCurrent
        }
    }
    $nonMailUsers | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFolderPath
        
}


function Get-UnlicensedNonMailObjects {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$True,HelpMessage='Which Domain Migrating?')] [string] $DomainQuery,
        [Parameter(Mandatory=$false,HelpMessage='AllCurrentUsers?')] [switch] $AllCurrentUsers,
        [Parameter(Mandatory=$false,HelpMessage='RemoveUnlicensedUsers?')] [switch] $RemoveUsers
    )
    #Import Modules
    function Import-ExchangeMsOnlineAzureModule() {
        #Exchange Online Module
        if ((Get-Module -Name "ExchangeOnlineManagement") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to ExchangeOnline ... " -nonewline -foregroundcolor cyan
                    Connect-ExchangeOnline
                    Write-Host "done" -foregroundcolor green
                }
                catch {
                    Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
                }
            }
            else {
                try {
                    Install-Module -Name ExchangeOnlineManagement 
                }
                catch {
                    Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
                }
            }
        }
        #Microsoft Online Module
        if ((Get-Module -Name "MSOnline") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to MSOnline ... " -nonewline -foregroundcolor cyan
                    Connect-MsolService
                    Write-Host "done" -foregroundcolor green
                }
                catch {
                    Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
            else {
                try {
                    Install-Module -Name MSOnline 
                }
                catch {
                    Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
        }
        #Azure AD Module
        if ((Get-Module -Name "AzureAD") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to AzureAD ... " -nonewline -foregroundcolor cyan
                    Connect-AzureAD
                    Write-Host "AzureAD" -foregroundcolor green
                }
                catch {
                    Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
            else {
                try {
                    Install-Module -Name AzureAD 
                }
                catch {
                    Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
        }
        
    }
    Import-ExchangeMsOnlineAzureModule
    # Get list of msolusers and proxyaddress of domains to remove
    $domainQueryString = "@" + $DomainQuery
    $domainRemovalUsers = Get-MsolUser -All -UnlicensedUsersOnly | ? {$_.proxyaddresses -like "*$domainQueryString"}

    #check if they are mail users
    #Progress Bar 1A
    $progressref = ($domainRemovalUsers).count
    $progresscounter = 0
    $nonMailUsers = @()
    foreach ($user in $domainRemovalUsers) {
        #Progress Bar 1B
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking $($user.displayname)"
        if (!($recipientCheck = Get-EXORecipient $user.UserPrincipalName -ea silentlycontinue)) {
            Write-Host "$($user.UserPrincipalName) not found as Mail Recipient" -ForegroundColor Cyan
            $proxyAddresses = $user.ProxyAddresses
            $nonMailUsersCurrent = New-Object PSObject
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $user.DisplayName
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $user.UserPrincipalName
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "ProxyAddresses" -Value ($proxyAddresses -join ",")
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "WhenCreated" -Value $user.WhenCreated
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "LastDirSyncTime" -Value $user.LastDirSyncTime
            $nonMailUsersCurrent | Add-Member -MemberType NoteProperty -Name "IsLicensed" -Value $user.IsLicensed
            $nonMailUsers += $nonMailUsersCurrent
            if ($RemoveUsers) {
                Remove-AzureADUser -ObjectId $user.userprincipalname
                Write-Host "$($user.UserPrincipalName) Removed" -ForegroundColor Green
            }
        }
    }
    $nonMailUsers | Export-Csv -NoTypeInformation -Encoding utf8 $OutputCSVFolderPath   
}

#Domain Cutover - Update UPN/PrimarySMTP Address (Source)
function Start-TenantDomainCutover {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$false,HelpMessage='AllCurrentUsers?')] [switch] $AllCurrentUsers,
        [Parameter(Mandatory=$false,HelpMessage="Update UPN?")] [switch]$SwitchUPN,
        [Parameter(Mandatory=$false,HelpMessage="Update PrimarySMTPAddress")] [switch]$PrimarySMTPAddress,
        [Parameter(Mandatory=$false,HelpMessage="How Many Batches/Groups?")] [String]$NumberBatches,
        [Parameter(Mandatory=$false,HelpMessage="Which Batch/Group to Run")] [String]$BatchNumber
    )
    function Import-ExchangeMsOnlineAzureModule() {
        #Exchange Online Module
        if ((Get-Module -Name "ExchangeOnlineManagement") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to ExchangeOnline ... " -nonewline -foregroundcolor cyan
                    Connect-ExchangeOnline
                    Write-Host "done" -foregroundcolor green
                }
                catch {
                    Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
                }
            }
            else {
                try {
                    Install-Module -Name ExchangeOnlineManagement 
                }
                catch {
                    Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
                }
            }
        }
        #Microsoft Online Module
        if ((Get-Module -Name "MSOnline") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to MSOnline ... " -nonewline -foregroundcolor cyan
                    Connect-MsolService
                    Write-Host "done" -foregroundcolor green
                }
                catch {
                    Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
            else {
                try {
                    Install-Module -Name MSOnline 
                }
                catch {
                    Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
        }
        #Azure AD Module
        if ((Get-Module -Name "AzureAD") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to AzureAD ... " -nonewline -foregroundcolor cyan
                    Connect-AzureAD
                    Write-Host "AzureAD" -foregroundcolor green
                }
                catch {
                    Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
            else {
                try {
                    Install-Module -Name AzureAD 
                }
                catch {
                    Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
        }
        
    }
    Import-ExchangeMsOnlineAzureModule
    #Gather OnMicrosoft Domain
    #$onMicrosoftDomain = (Get-MsolDomain | ? {$_.name -like "*.onmicrosoft.com" -and $_.name -notlike "*.mail.onmicrosoft.com"}).Name
    $allMSOLUsers = @()
    #Gather All Users
    if ($AllCurrentUsers) {
        $allMSOLUsers = Get-MsolUser -All
    }
    elseif ($InputCSVFilePath) {
        $allMSOLUsers = Import-CSV $InputCSVFilePath
    }
    elseif ($InputEXCELFilePath) {
        $allMSOLUsers = Import-Excel $InputCSVFilePath
    }
    elseif ($NumberBatches) {
        $alltmpMSOLUsers = Get-MsolUser -All
        $MsolUsersBatch = ($alltmpMSOLUsers.count)/($NumberBatches)
        if ($BatchNumber -eq 1) {
            $allMSOLUsers = $alltmpMSOLUsers[0..$MsolUsersBatch]     
        }
        elseif ($BatchNumber -eq 2) {
            $startingbatch = $MsolUsersBatch+1
            $allMSOLUsers = $alltmpMSOLUsers[$startingbatch..($MsolUsersBatch*2)]
        }
        elseif ($BatchNumber -gt 3) {
            $startingbatch = ($MsolUsersBatch*($BatchNumber-1))+1
            $allMSOLUsers = $alltmpMSOLUsers[$startingbatch..($MsolUsersBatch*2)]
        } 
        elseif ($BatchNumber -eq "Last")  {
            $allMSOLUsers = $alltmpMSOLUsers[0..(-1*$MsolUsersBatch)]
        }
    }
    
    Write-Host "$($allMSOLUsers.count) Users to Update"
    #Progress Bar 1A
    $progressref = ($allMSOLUsers).count
    $progresscounter = 0
    $AllErrors = @()
    foreach ($msolUser in $allMsolUsers) {
        #Progress Bar 1B
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Domain Cutover for $($msolUser.displayname) to $($onMicrosoftDomain)"
        #Variable - NewUPN
        $onMicrosoftDomain = "EHN.onmicrosoft.com"
        $newUPN = ($msolUser.UserPrincipalName -split "@")[0] + "@" + $onMicrosoftDomain
        Write-Host "Update $($msolUser.displayname)" -ForegroundColor Cyan -NoNewline
        Write-Host ".." -ForegroundColor Yellow -NoNewline
        
        #Update UPN
        if ($SwitchUPN) {
            Write-Host "UPNUpdate.." -NoNewline -foregroundcolor DarkCyan
            if ($msolUser.UserPrincipalName -like "*$onMicrosoftDomain") {
                Write-Host "Skipping." -ForegroundColor Yellow -NoNewline
            }
            elseif ($msolCheck = Get-MsolUser -UserPrincipalName $msolUser.UserPrincipalName -erroraction SilentlyContinue) {
                try {
                    $UPNUpdateStat = Set-MsolUserPrincipalName -UserPrincipalName $msolCheck.UserPrincipalName -NewUserPrincipalName $newUPN -ErrorAction Stop
                    Write-host "$($newupn). " -foregroundcolor Green -NoNewline
                }
                catch {
                    Write-Error "$($newupn). "
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdateUPN" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $msolUser.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "NewUPN" -Value $newUPN -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrorsPerms += $currenterror           
                    continue
                }
            }
            else {
               Write-Host "No UPN found. " -ForegroundColor Red -NoNewline
            }
        }

        #Update PrimarySMTPAddress
        if ($PrimarySMTPAddress) {
            Write-Host "PrimarySMTPAddressUpdate.." -NoNewline -foregroundcolor DarkCyan
            if (Get-EXORecipient $msolUser.UserPrincipalName -erroraction SilentlyContinue) {
                if ((Get-EXORecipient $msolUser.UserPrincipalName).primarysmtpaddress -like "*$onMicrosoftDomain") {
                    Write-Host "Skipping." -ForegroundColor Yellow -NoNewline
                }
                else {
                    try {
                        Set-Mailbox -Identity $msolUser.UserPrincipalName -WindowsEmailAddress $newUPN
                        Write-host "Updated." -foregroundcolor Green -NoNewline
                    }
                    catch {
                        Write-Error "Failed to Update."
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdatePrimarySMTPAddress" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $msolUser.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "NewUPN" -Value $newUPN -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
            }
            else {
                Write-host "No Recipient Found." -foregroundcolor Red -NoNewline
            }
        }
        #Completed User
        Write-Host "Done" -ForegroundColor Green
    }
    $AllErrors
}

#Domain Cutover - Update UPN/PrimarySMTP Address (Source)
function Start-TenantRecipientDomainCutover {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$false,HelpMessage='AllCurrentUsers?')] [switch] $AllCurrentUsers,
        [Parameter(Mandatory=$false,HelpMessage="Update UPN?")] [switch]$SwitchUPN,
        [Parameter(Mandatory=$false,HelpMessage="Update PrimarySMTPAddress")] [switch]$PrimarySMTPAddress,
        [Parameter(Mandatory=$false,HelpMessage="How Many Batches/Groups?")] [String]$NumberBatches,
        [Parameter(Mandatory=$false,HelpMessage="Which Batch/Group to Run")] [String]$BatchNumber
    )
    function Import-ExchangeMsOnlineAzureModule() {
        #Exchange Online Module
        if ((Get-Module -Name "ExchangeOnlineManagement") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to ExchangeOnline ... " -nonewline -foregroundcolor cyan
                    Connect-ExchangeOnline
                    Write-Host "done" -foregroundcolor green
                }
                catch {
                    Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
                }
            }
            else {
                try {
                    Install-Module -Name ExchangeOnlineManagement 
                }
                catch {
                    Write-Error  "ExchangeOnline module was not loaded. Run Install-Module ExchangeOnlineManagement as an Administrator. More details to install the EXO Version 3 can be found at https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0. More Details at https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"
                }
            }
        }
        #Microsoft Online Module
        if ((Get-Module -Name "MSOnline") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to MSOnline ... " -nonewline -foregroundcolor cyan
                    Connect-MsolService
                    Write-Host "done" -foregroundcolor green
                }
                catch {
                    Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
            else {
                try {
                    Install-Module -Name MSOnline 
                }
                catch {
                    Write-Error  "MSOnline module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
        }
        #Azure AD Module
        if ((Get-Module -Name "AzureAD") -ne $null) {
            return;
        }
        else {
            if ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -ne $null) {
                try {
                    Write-Host "Connecting to AzureAD ... " -nonewline -foregroundcolor cyan
                    Connect-AzureAD
                    Write-Host "AzureAD" -foregroundcolor green
                }
                catch {
                    Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
            else {
                try {
                    Install-Module -Name AzureAD 
                }
                catch {
                    Write-Error  "AzureAD module was not loaded. Run Install-Module MSOnline as an Administrator."
                }
            }
        }
        
    }
    Import-ExchangeMsOnlineAzureModule
    #Gather OnMicrosoft Domain
    $onMicrosoftDomain = (Get-MsolDomain | ? {$_.name -like "*.onmicrosoft.com" -and $_.name -notlike "*.mail.onmicrosoft.com"}).Name
    #Gather All Users
    if ($AllCurrentUsers) {
        $allRecipients = Get-EXORecipient -ResultSize Unlimited
    }
    elseif ($InputCSVFilePath) {
        $allRecipients = Import-CSV $InputCSVFilePath
    }
    elseif ($InputEXCELFilePath) {
        $allRecipients = Import-Excel $InputCSVFilePath
    }
    elseif ($NumberBatches) {
        $alltmpMSOLUsers = Get-MsolUser -All
        $MsolUsersBatch = ($alltmpMSOLUsers.count)/($NumberBatches)
        if ($BatchNumber -eq 1) {
            $allMSOLUsers = $alltmpMSOLUsers[0..$MsolUsersBatch]     
        }
        elseif ($BatchNumber -eq 2) {
            $startingbatch = $MsolUsersBatch+1
            $allMSOLUsers = $alltmpMSOLUsers[$startingbatch..($MsolUsersBatch*2)]
        }
        elseif ($BatchNumber -gt 3) {
            $startingbatch = ($MsolUsersBatch*($BatchNumber-1))+1
            $allMSOLUsers = $alltmpMSOLUsers[$startingbatch..($MsolUsersBatch*2)]
        } 
        elseif ($BatchNumber -eq "Last")  {
            $allMSOLUsers = $alltmpMSOLUsers[0..(-1*$MsolUsersBatch)]
        }
    }
    
    Write-Host "$($allRecipients.count) Users to Update"
    #Progress Bar 1A
    $onMicrosoftDomain
    $progressref = ($allRecipients).count
    $progresscounter = 0
    $AllErrors = @()
    foreach ($recipient in $allRecipients| sort userprincipalname) {
        #Progress Bar 1B
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Domain Cutover for $($recipient.displayname) to $($onMicrosoftDomain)"
        #Variable - NewUPN
        
        $newAddress = ($recipient.PrimarySMTPAddress -split "@")[0] + "@" + $onMicrosoftDomain
        Write-Host "Update $($recipient.displayname)" -ForegroundColor Cyan -NoNewline
        Write-Host ".." -ForegroundColor Yellow -NoNewline 
        #Update UPN
        if ($SwitchUPN) {
            Write-Host "UPNUpdate.." -NoNewline -foregroundcolor DarkCyan
            if ($msolUser.UserPrincipalName -like "*$onMicrosoftDomain") {
                Write-Host "Skipping." -ForegroundColor Yellow -NoNewline
            }
            elseif ($msolCheck = Get-MsolUser -UserPrincipalName $msolUser.UserPrincipalName -erroraction SilentlyContinue) {
                try {
                    $UPNUpdateStat = Set-MsolUserPrincipalName -UserPrincipalName $msolCheck.UserPrincipalName -NewUserPrincipalName $newUPN -ErrorAction Stop
                    Write-host "$($newupn). " -foregroundcolor Green -NoNewline
                }
                catch {
                    Write-Error "$($newupn). "
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdateUPN" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $msolUser.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "NewUPN" -Value $newUPN -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrorsPerms += $currenterror           
                    continue
                }
            }
            else {
               Write-Host "No UPN found. " -ForegroundColor Red -NoNewline
            }
        }

        #Update PrimarySMTPAddress
        if ($PrimarySMTPAddress) {
            if (!($recipient.PrimarySMTPAddress -like "*$onMicrosoftDomain")) {
                Write-Host "PrimarySMTPAddressUpdate.." -NoNewline -foregroundcolor DarkCyan
                if ($recipient.RecipientTypeDetails -eq "GroupMailbox" -or $recipient.RecipientTypeDetails -eq "Team") {
                    try {
                        Set-UnifiedGroup -Identity $recipient.PrimarySMTPAddress -WindowsEmailAddress $newAddress -ErrorAction Stop
                    }
                    catch {
                        Write-Error "$($newAddress). "
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                    
                }
                elseif ($recipient.RecipientTypeDetails -eq "DynamicDistributionGroup") {
                    try {
                        Set-DynamicDistributionGroup -Identity $recipient.PrimarySMTPAddress -WindowsEmailAddress $newAddress -ErrorAction Stop
                    }
                    catch {
                        Write-Error "$($newAddress). "
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
                elseif ($recipient.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                    try {
                        Set-DistributionGroup -Identity $recipient.PrimarySMTPAddress -WindowsEmailAddress $newAddress -ErrorAction Stop
                    }
                    catch {
                        Write-Error "$($newAddress). "
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
                elseif ($recipient.RecipientTypeDetails  -like "*Mailbox") {
                    try {
                        Set-Mailbox Identity $recipient.PrimarySMTPAddress -WindowsEmailAddress $newAddress -ErrorAction Stop    
                    }
                    catch {
                        Write-Error "$($newAddress). "
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
                else {
                    try {
                        Set-Mailbox Identity $recipient.PrimarySMTPAddress -WindowsEmailAddress $newAddress -ErrorAction Stop
                    }
                    catch {
                        Write-Error "$($newAddress). "
                        $currenterror = new-object PSObject
                        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                        $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                        $AllErrors += $currenterror           
                        continue
                    }
                }
            }
        }
        #Completed User
        Write-Host "Done" -ForegroundColor Green
    }
    $AllErrors
}

#Domain Cutover - Update UPN/PrimarySMTP Address (Destination)
function Add-TenantDomainCutover {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Where do you wish to save this file? Please provide full FOLDERPATH')] [string] $OutputCSVFolderPath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$false,HelpMessage='AllCurrentUsers?')] [string] $AllCurrentUsers,
        [Parameter(Mandatory=$false,HelpMessage="Update UPN?")] [switch]$SwitchUPN,
        [Parameter(Mandatory=$false,HelpMessage="Update Aliases")] [switch]$AddAlias,
        [Parameter(Mandatory=$false,HelpMessage="Skip UserMailboxes")] [switch]$NoUserMailboxes,
        [Parameter(Mandatory=$false,HelpMessage="Skip UserMailboxes")] [switch]$CloudMailboxes
    )

    #Gather All Users
    if ($InputCSVFilePath) {
        $allMatchedRecipients = Import-CSV $InputCSVFilePath
    }
    elseif ($InputEXCELFilePath) {
        $allMatchedRecipients = Import-Excel $InputEXCELFilePath
    }
    if ($NoUserMailboxes) {
        $allMatchedRecipients = $allMatchedRecipients | ?{$_.RecipientTypeDetails_Destination -ne "UserMailbox"}
    }
    if ($CloudMailboxes) {
        $allMatchedRecipients = $allMatchedRecipients | ?{$_.RecipientTypeDetails_Destination}
    }
    #Progress Bar 1A
    
    $progressref = ($allMatchedRecipients).count
    $progresscounter = 0
    $AllErrors = @()

    Write-Host "$($allMatchedRecipients.count) Users to Update" -ForegroundColor Magenta
    foreach ($matchedRecipient in $allMatchedRecipients) {
        #Progress Bar 1B
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -Id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Migrating Domain Aliases to $($matchedRecipient.DisplayName_Destination)"

        #Variable - NewUPN
        $newAddress = $matchedRecipient.PrimarySMTPAddress
        Write-Host "Update $($matchedRecipient.displayname)" -ForegroundColor Cyan -NoNewline
        Write-Host ".." -ForegroundColor Yellow -NoNewline

        if ($matchedRecipient.RecipientTypeDetails_Destination -eq "GroupMailbox" -or $matchedRecipient.RecipientTypeDetails_Destination -eq "Team") {
            try {
                Set-UnifiedGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                Write-Host "." -NoNewline -ForegroundColor Yellow
                if ($newAliases = $matchedRecipient.Aliases -split ",") {
                    foreach ($newAlias in $newAliases) {
                        Set-UnifiedGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                        Write-Host "." -NoNewline -ForegroundColor Blue
                    }
                }
                Write-host "$($newAddress). " -foregroundcolor Green -NoNewline
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        elseif ($matchedRecipient.RecipientTypeDetails_Destination -eq "DynamicDistributionGroup") {
            try {
                Set-DynamicDistributionGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                Write-Host "." -NoNewline -ForegroundColor Yellow
                if ($newAliases = $matchedRecipient.Aliases -split ",") {
                    foreach ($newAlias in $newAliases) {
                    Set-DynamicDistributionGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    Write-Host "." -NoNewline -ForegroundColor Blue
                    }
                }
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        elseif ($matchedRecipient.RecipientTypeDetails_Destination -eq "MailUniversalDistributionGroup") {
            try {
                Set-DistributionGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                Write-Host "." -NoNewline -ForegroundColor Yellow
                if ($newAliases = $matchedRecipient.Aliases -split ",") {
                    foreach ($newAlias in $newAliases) {
                    Set-DistributionGroup -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    Write-Host "." -NoNewline -ForegroundColor Blue
                    }
                }
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        <#elseif ($matchedRecipient.RecipientTypeDetails_Destination -eq "UserMailbox") {
            Write-Host "Skipping UserMailbox." -NoNewline -ForegroundColor DarkCyan
        }#>
        elseif ($matchedRecipient.RecipientTypeDetails_Destination -like "*Mailbox" -and $matchedRecipient.RecipientTypeDetails_Destination -ne "UserMailbox") {
            try {
                Set-Mailbox -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                Write-Host "." -NoNewline -ForegroundColor Yellow
                if ($newAliases = $matchedRecipient.Aliases -split ",") {
                    foreach ($newAlias in $newAliases) {
                    Set-Mailbox -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    Write-Host "." -NoNewline -ForegroundColor Blue
                    }
                }
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        elseif ($matchedRecipient.RecipientTypeDetails_Destination -eq "UserMailbox") {
            try {
                Set-Mailbox -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                Write-Host "." -NoNewline -ForegroundColor Yellow
                if ($newAliases = $matchedRecipient.Aliases -split ",") {
                    foreach ($newAlias in $newAliases) {
                    Set-Mailbox -Identity $matchedRecipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAlias}
                    Write-Host "." -NoNewline -ForegroundColor Blue
                    }
                }
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        elseif ($recipient.RecipientTypeDetails  -eq "MailUser") {
            try {
                Set-MailUser -Identity $recipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                Write-Host "." -NoNewline -ForegroundColor Yellow
                if ($newAliases = $matchedRecipient.Aliases -split ",") {
                    foreach ($newAlias in $newAliases) {
                    Set-MailUser -Identity $recipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress}
                    Write-Host "." -NoNewline -ForegroundColor Blue
                    }
                }
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        elseif ($recipient.RecipientTypeDetails  -eq "MailContact") {
            try {
                Set-MailContact -Identity $recipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress} -ErrorAction Stop
                Write-Host "." -NoNewline -ForegroundColor Yellow
                if ($newAliases = $matchedRecipient.Aliases -split ",") {
                    foreach ($newAlias in $newAliases) {
                    Set-MailContact -Identity $recipient.PrimarySMTPAddress_Destination -EmailAddresses @{add=$newAddress}
                    Write-Host "." -NoNewline -ForegroundColor Blue
                    }
                }
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "Aliases" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails_Destination" -Value $matchedRecipient -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress_Destination" -Value $matchedRecipient.UserPrincipalName -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        
        Write-Host "Done" -ForegroundColor DarkGreen
    }
}



##UPDATE PRIMARYSMTPADDRESS
$allEinsteinRecipients = Get-EXORecipient -ResultSize Unlimited | sort ?{$_.primarysmtpaddress -like "*@einstein.edu" -and $_.RecipientTypeDetails -ne "MailContact"}
#Progress Bar 1A
$onMicrosoftDomain = "ehn.onmicrosoft.com"
$progressref = ($allEinsteinRecipients).count
$progresscounter = 0
$AllErrors = @()
foreach ($recipient in $allEinsteinRecipients) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Domain Cutover for $($recipient.displayname) to $($onMicrosoftDomain)"
    #Variable - NewUPN
    
    $newAddress = ($recipient.PrimarySMTPAddress -split "@")[0] + "@" + $onMicrosoftDomain
    Write-Host "Update $($recipient.displayname) to $newAddress" -ForegroundColor Cyan -NoNewline
    Write-Host ".." -ForegroundColor Yellow -NoNewline
    #Update PrimarySMTPAddress
    if (!($recipient.PrimarySMTPAddress -like "*$onMicrosoftDomain")) {
        Write-Host "PrimarySMTPAddressUpdate.." -NoNewline -foregroundcolor DarkCyan
        if ($recipient.RecipientTypeDetails -eq "GroupMailbox" -or $recipient.RecipientTypeDetails -eq "Team") {
            try {
                Set-UnifiedGroup -Identity $recipient.PrimarySMTPAddress -PrimarySmtpAddress $newAddress -ErrorAction Stop
                Write-Host "Unified Group Done." -foregroundcolor Green
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
            
        }
        elseif ($recipient.RecipientTypeDetails -eq "DynamicDistributionGroup") {
            try {
                Set-DynamicDistributionGroup -Identity $recipient.PrimarySMTPAddress -PrimarySmtpAddress $newAddress -ErrorAction Stop
                Write-Host "Dynamic Group Done." -foregroundcolor Green
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        elseif ($recipient.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
            try {
                Set-DistributionGroup -Identity $recipient.PrimarySMTPAddress -PrimarySmtpAddress $newAddress -ErrorAction Stop
                Write-Host "Distribution Group Done." -foregroundcolor Green
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        elseif ($recipient.RecipientTypeDetails  -like "*Mailbox") {
            try {
                Set-Mailbox -Identity $recipient.PrimarySMTPAddress -WindowsEmailAddress $newAddress -ErrorAction Stop
                Write-Host "Mailbox Done." -foregroundcolor Green
            }
            catch {
                Write-Host "Failed. " -ForegroundColor Red
                Write-Host $_.Exception
                $currenterror = new-object PSObject
                $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                $AllErrors += $currenterror           
                continue
            }
        }
        
    }
}


##Remove Migrating Domain
function Remove-MigratingDomainSource {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$false,HelpMessage='MigratingDomain')] [string] $MigratingDomainString
    )
    ##Remove Migrating Domain
    $onMicrosoftDomain = (Get-MsolDomain | ? {$_.name -like "*.onmicrosoft.com" -and $_.name -notlike "*.mail.onmicrosoft.com"}).Name
    $MigratingDomain = "@" + $MigratingDomainString 
    $allMigratingRecipients = Get-EXORecipient -ResultSize Unlimited | ?{$_.EmailAddresses -like "*$MigratingDomain"} | sort DisplayName
    Write-Host "$($allMigratingRecipients.count) Users to Update"
    #Progress Bar 1A
    $onMicrosoftDomain
    $progressref = ($allMigratingRecipients).count
    $progresscounter = 0
    $AllErrors = @()
    foreach ($recipient in $allMigratingRecipients) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Domain Cutover for $($recipient.displayname) to $($onMicrosoftDomain)"
    #Variable - NewUPN
    if ($Aliases = $recipient.EmailAddresses -like "*$MigratingDomain") {
        Write-Host "Removing " -ForegroundColor Cyan -NoNewline
        Write-Host "$($aliases.count) " -ForegroundColor Yellow -NoNewline
        Write-Host "Aliases for $($recipient.displayname).. " -ForegroundColor Cyan -NoNewline
        foreach ($alias in $Aliases) {
            if ($recipient.RecipientTypeDetails -eq "GroupMailbox" -or $recipient.RecipientTypeDetails -eq "Team") {
                try {
                    Set-UnifiedGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }      
            }
            elseif ($recipient.RecipientTypeDetails -eq "DynamicDistributionGroup") {
                try {
                    Set-DynamicDistributionGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                try {
                    Set-DistributionGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails  -like "*Mailbox") {
                try {
                    Set-Mailbox -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails  -eq "MailUser") {
                try {
                    Set-MailUser -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails  -eq "MailContact") {
                try {
                    Set-MailContact -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
        }
        Write-Host "done" -ForegroundColor Green
    }

    }
}

##Remove Migrating Domain - Batches
$migratingDomain = "@gw.einstein.edu"
$allEinsteinRecipients = Get-EXORecipient -ResultSize Unlimited | ?{$_.EmailAddresses -like "*$migratingDomain"} | sort DisplayName
$LastEinsteinRecipients = $allEinsteinRecipients[-1..-2999]
$2ndLastBatch = $allEinsteinRecipients[-3000..-5001]
#Progress Bar 1A
$onMicrosoftDomain = "ehn.onmicrosoft.com"
$progressref = ($2ndLastBatch).count
$progresscounter = 0
$AllErrors = @()
foreach ($recipient in $2ndLastBatch) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Domain Cutover for $($recipient.displayname) to $($onMicrosoftDomain)"
    #Variable - NewUPN
    if ($Aliases = $recipient.EmailAddresses -like "*$migratingDomain") {
        Write-Host "Removing " -ForegroundColor Cyan -NoNewline
        Write-Host "$($aliases.count) " -ForegroundColor Yellow -NoNewline
        Write-Host "Aliases for $($recipient.displayname).. " -ForegroundColor Cyan -NoNewline
        foreach ($alias in $Aliases) {
            if ($recipient.RecipientTypeDetails -eq "GroupMailbox" -or $recipient.RecipientTypeDetails -eq "Team") {
                try {
                    Set-UnifiedGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }      
            }
            elseif ($recipient.RecipientTypeDetails -eq "DynamicDistributionGroup") {
                try {
                    Set-DynamicDistributionGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                try {
                    Set-DistributionGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails  -like "*Mailbox") {
                try {
                    Set-Mailbox -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails  -eq "MailUser") {
                try {
                    Set-MailUser -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                    Write-Host "." -NoNewline -foregroundcolor Yellow
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
        }
        Write-Host "done" -ForegroundColor Green
    }
    
}

##Remove Migrating Domain
function Remove-MigratingDomainSource {
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of CSV Import List')] [string] $InputCSVFilePath,
        [Parameter(Mandatory=$false,HelpMessage='Provide Full File Path of EXCEL Import List')] [string] $InputEXCELFilePath,
        [Parameter(Mandatory=$false,HelpMessage='MigratingDomain')] [string] $MigratingDomainString
    )
    ##Remove Migrating Domain
    $onMicrosoftDomain = (Get-MsolDomain | ? {$_.name -like "*.onmicrosoft.com" -and $_.name -notlike "*.mail.onmicrosoft.com"}).Name
    $MigratingDomain = "@" + $MigratingDomainString 
    $allMigratingRecipients = Get-EXORecipient -ResultSize Unlimited | ?{$_.EmailAddresses -like "*$MigratingDomain"} | sort DisplayName
    Write-Host "$($allMigratingRecipients.count) Users to Update"
    #Progress Bar 1A
    $onMicrosoftDomain
    $progressref = ($allMigratingRecipients).count
    $progresscounter = 0
    $AllErrors = @()
    foreach ($recipient in $allMigratingRecipients) {
        #Progress Bar 1B
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Domain Cutover for $($recipient.displayname) to $($onMicrosoftDomain)"
        #Update UPN
        if ($SwitchUPN) {
            Write-Host "UPNUpdate.." -NoNewline -foregroundcolor DarkCyan
            if ($msolUser.UserPrincipalName -like "*$onMicrosoftDomain") {
                Write-Host "Skipping." -ForegroundColor Yellow -NoNewline
            }
            elseif ($msolCheck = Get-MsolUser -UserPrincipalName $msolUser.UserPrincipalName -erroraction SilentlyContinue) {
                try {
                    $UPNUpdateStat = Set-MsolUserPrincipalName -UserPrincipalName $msolCheck.UserPrincipalName -NewUserPrincipalName $newUPN -ErrorAction Stop
                    Write-host "$($newupn). " -foregroundcolor Green -NoNewline
                }
                catch {
                    Write-Error "$($newupn). "
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "FailToUpdateUPN" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $msolUser.UserPrincipalName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "NewUPN" -Value $newUPN -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrorsPerms += $currenterror           
                    continue
                }
            }
            else {
            Write-Host "No UPN found. " -ForegroundColor Red -NoNewline
            }
        }

        #Update PrimarySMTPAddress
        if (!($recipient.PrimarySMTPAddress -like "*$onMicrosoftDomain")) {
            Write-Host "PrimarySMTPAddressUpdate.." -NoNewline -foregroundcolor DarkCyan
            if ($recipient.RecipientTypeDetails -eq "GroupMailbox" -or $recipient.RecipientTypeDetails -eq "Team") {
                try {
                    Set-UnifiedGroup -Identity $recipient.PrimarySMTPAddress -PrimarySmtpAddress $newAddress -ErrorAction Stop
                    Write-Host "Unified Group Done." -foregroundcolor Green
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
                
            }
            elseif ($recipient.RecipientTypeDetails -eq "DynamicDistributionGroup") {
                try {
                    Set-DynamicDistributionGroup -Identity $recipient.PrimarySMTPAddress -PrimarySmtpAddress $newAddress -ErrorAction Stop
                    Write-Host "Dynamic Group Done." -foregroundcolor Green
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                try {
                    Set-DistributionGroup -Identity $recipient.PrimarySMTPAddress -PrimarySmtpAddress $newAddress -ErrorAction Stop
                    Write-Host "Distribution Group Done." -foregroundcolor Green
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            elseif ($recipient.RecipientTypeDetails  -like "*Mailbox") {
                try {
                    Set-Mailbox -Identity $recipient.PrimarySMTPAddress -WindowsEmailAddress $newAddress -ErrorAction Stop
                    Write-Host "Mailbox Done." -foregroundcolor Green
                }
                catch {
                    Write-Host "Failed. " -ForegroundColor Red
                    Write-Host $_.Exception
                    $currenterror = new-object PSObject
                    $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                    $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "AddMigratingDomain" -Force
                    $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                    $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                    $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                    $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                    $AllErrors += $currenterror           
                    continue
                }
            }
            
        }
        #Remove Migrating Domain Aliases
        if ($RemoveAliases) {
            if ($Aliases = $recipient.EmailAddresses -like "*$MigratingDomain") {
                Write-Host "Removing " -ForegroundColor Cyan -NoNewline
                Write-Host "$($aliases.count) " -ForegroundColor Yellow -NoNewline
                Write-Host "Aliases for $($recipient.displayname).. " -ForegroundColor Cyan -NoNewline
                foreach ($alias in $Aliases) {
                    if ($recipient.RecipientTypeDetails -eq "GroupMailbox" -or $recipient.RecipientTypeDetails -eq "Team") {
                        try {
                            Set-UnifiedGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                            Write-Host "." -NoNewline -foregroundcolor Yellow
                        }
                        catch {
                            Write-Host "Failed. " -ForegroundColor Red
                            Write-Host $_.Exception
                            $currenterror = new-object PSObject
                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                            $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                            $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                            $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                            $AllErrors += $currenterror           
                            continue
                        }      
                    }
                    elseif ($recipient.RecipientTypeDetails -eq "DynamicDistributionGroup") {
                        try {
                            Set-DynamicDistributionGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                            Write-Host "." -NoNewline -foregroundcolor Yellow
                        }
                        catch {
                            Write-Host "Failed. " -ForegroundColor Red
                            Write-Host $_.Exception
                            $currenterror = new-object PSObject
                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                            $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                            $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                            $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                            $AllErrors += $currenterror           
                            continue
                        }
                    }
                    elseif ($recipient.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                        try {
                            Set-DistributionGroup -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                            Write-Host "." -NoNewline -foregroundcolor Yellow
                        }
                        catch {
                            Write-Host "Failed. " -ForegroundColor Red
                            Write-Host $_.Exception
                            $currenterror = new-object PSObject
                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                            $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                            $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                            $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                            $AllErrors += $currenterror           
                            continue
                        }
                    }
                    elseif ($recipient.RecipientTypeDetails  -like "*Mailbox") {
                        try {
                            Set-Mailbox -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                            Write-Host "." -NoNewline -foregroundcolor Yellow
                        }
                        catch {
                            Write-Host "Failed. " -ForegroundColor Red
                            Write-Host $_.Exception
                            $currenterror = new-object PSObject
                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                            $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                            $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                            $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                            $AllErrors += $currenterror           
                            continue
                        }
                    }
                    elseif ($recipient.RecipientTypeDetails  -eq "MailUser") {
                        try {
                            Set-MailUser -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                            Write-Host "." -NoNewline -foregroundcolor Yellow
                        }
                        catch {
                            Write-Host "Failed. " -ForegroundColor Red
                            Write-Host $_.Exception
                            $currenterror = new-object PSObject
                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                            $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                            $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                            $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                            $AllErrors += $currenterror           
                            continue
                        }
                    }
                    elseif ($recipient.RecipientTypeDetails  -eq "MailContact") {
                        try {
                            Set-MailContact -Identity $recipient.PrimarySMTPAddress -EmailAddresses @{Remove=$alias} -ErrorAction Stop
                            Write-Host "." -NoNewline -foregroundcolor Yellow
                        }
                        catch {
                            Write-Host "Failed. " -ForegroundColor Red
                            Write-Host $_.Exception
                            $currenterror = new-object PSObject
                            $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
                            $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "RemoveMigratingDomain" -Force
                            $currenterror | Add-Member -type NoteProperty -Name "DisplayName" -Value $recipient.DisplayName -Force
                            $currenterror | Add-Member -type NoteProperty -Name "RecipientTypeDetails" -Value $recipient.RecipientTypeDetails -Force
                            $currenterror | Add-Member -type NoteProperty -Name "PrimarySMTPAddress" -Value $recipient.PrimarySMTPAddress -Force
                            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
                            $AllErrors += $currenterror           
                            continue
                        }
                    }
                }
                Write-Host "done" -ForegroundColor Green
            }
        }      
    }
}
#Verify Matched Addresses
$EinsteinMailUsers = Import-CSV "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson External - Einstein to Jefferson Migration\Domain Cutover\NotFound-EinsteinUsers2.csv"

$notFoundUsers = @()
$foundUsers = @()
#Progress Bar 1A
$progressref = ($EinsteinMailUsers).count
$progresscounter = 0
Write-Host "$($EinsteinMailUsers.count) Users to Review" -ForegroundColor Magenta
foreach ($matchedRecipient in $EinsteinMailUsers) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Aliases to $($matchedRecipient.PrimarySMTPAddress)"
    if ($RecipientCheck = Get-EXORecipient $matchedRecipient.PrimarySMTPAddress -ErrorAction SilentlyContinue) {
        $FoundUsers += $matchedRecipient
        Write-Host "." -ForegroundColor Green -NoNewline
        $matchedRecipient | add-member -type noteproperty -name "AddressMigrated" -Value $True -Force
    }
    else {
        $notFoundUsers += $matchedRecipient
        Write-Host "." -ForegroundColor Red -NoNewline
        $matchedRecipient | add-member -type noteproperty -name "AddressMigrated" -Value $False -Force
    }
}

$notFoundUsers | Export-Csv -NoTypeInformation -Encoding utf8 "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson External - Einstein to Jefferson Migration\Domain Cutover\NotFound-EinsteinUsers2.csv"


#Verify Matched Addresses
$notFoundUsers = Import-CSV "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson External - Einstein to Jefferson Migration\Domain Cutover\NotFound-EinsteinUsers2.csv"

$notFoundUsers2 = @()
$foundUsers2 = @()
#Progress Bar 1A
$progressref = ($notFoundUsers).count
$progresscounter = 0
Write-Host "$($notFoundUsers.count) Users to Review" -ForegroundColor Magenta
foreach ($matchedRecipient in $notFoundUsers) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Aliases to $($matchedRecipient.PrimarySMTPAddress)"
    if ($RecipientCheck = Get-EXORecipient $matchedRecipient.PrimarySMTPAddress -ErrorAction SilentlyContinue) {
        $foundUsers2 += $matchedRecipient
        Write-Host "." -ForegroundColor Green -NoNewline
        $matchedRecipient | add-member -type noteproperty -name "AddressMigrated" -Value $True -Force
    }
    else {
        $notFoundUsers2 += $matchedRecipient
        Write-Host "." -ForegroundColor Red -NoNewline
        $matchedRecipient | add-member -type noteproperty -name "AddressMigrated" -Value $False -Force
    }
}

$notFoundUsers2 | Export-Csv -NoTypeInformation -Encoding utf8 "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson External - Einstein to Jefferson Migration\Domain Cutover\NotFound-EinsteinUsers2.csv"

## Add temporary Alias
## Import Users
$EinsteinMailUsers = Import-Csv "C:\NotFound-EinsteinUsers.csv"

#Progress Bar 1A
$progressref = ($EinsteinMailUsers).count
$progresscounter = 0
$aliasNumber = 0
foreach ($user in $EinsteinMailUsers) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Adding Aliases for $($user.DisplayName_Destination)"
    $aliasNumber += 1
    $tmpAlias = "temporaryAlias" + $aliasNumber + "@einstein.edu"
    Write-Host "Add $($tmpAlias) to $($user.PrimarySMTPAddress_Destination)"
    Set-AdUser -identity $user.PrimarySMTPAddress_Destination -ProxyAddresses @{Add=$tmpAlias}
}

## Add temporary Alias
## Import Users
$notFoundUsers = Import-Csv "C:\NotFound-EinsteinUsers.csv"
$allMatchedRecipients = Import-Excel

#Progress Bar 1A
$progressref = ($notFoundUsers).count
$progresscounter = 0
$aliasNumber = 0
foreach ($user in $notFoundUsers) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Matching User $($user.DisplayName)"
    if ($matchedObject = $allmatchedRecipients | ? {$_.PrimarySMTPAddress -eq $user.PrimarySMTPAddress}) {
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $matchedObject.DisplayName_Destination -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $matchedObject.RecipientTypeDetails_Destination -Force
        $user | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $matchedObject.PrimarySMTPAddress_Destination -Force
    }
    else {
        $user | add-member -type noteproperty -name "DisplayName_Destination" -Value $null -Force
        $user | add-member -type noteproperty -name "RecipientTypeDetails_Destination" -Value $null -Force
        $user | add-member -type noteproperty -name "PrimarySMTPAddress_Destination" -Value $null -Force
    }
}

#Verify Matched Addresses
$EinsteinMailUsers = Import-CSV "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson External - Einstein to Jefferson Migration\Domain Cutover\NotFound-EinsteinUsers2.csv"

#Progress Bar 1A
$progressref = ($EinsteinMailUsers).count
$progresscounter = 0
Write-Host "$($EinsteinMailUsers.count) Users to Review" -ForegroundColor Magenta
foreach ($matchedRecipient in $EinsteinMailUsers) {
    #Progress Bar 1B
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Checking Aliases to $($matchedRecipient.PrimarySMTPAddress)"
    if ($RecipientCheck = Get-EXORecipient $matchedRecipient.DisplayName -ErrorAction SilentlyContinue) {
        Write-Host "." -ForegroundColor Green -NoNewline
        $matchedRecipient | add-member -type noteproperty -name "MatchedTargetAddress" -Value $RecipientCheck.PrimarySMTPAddress -Force
    }
    else {
        Write-Host "." -ForegroundColor Red -NoNewline
        $matchedRecipient | add-member -type noteproperty -name "MatchedTargetAddress" -Value $False -Force
    }
}