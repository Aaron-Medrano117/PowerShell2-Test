<# Read Me #
Create and Start OneDrive Migration Job
This is a function, please copy the function into powershell and then run the function. 

## Example 1 Start ##

### Initial Variables ###
$sourceURL = "https://ehn-admin.sharepoint.com/"
$DestinationURL = "https://tjuv-admin.sharepoint.com/"
$SourceCredentials = Get-Credentials
#^Enter Migration Service Account
$DestinationCredentials = Get-Credential
#^Enter Migration Service Account
$filePath = "C:\Users\amedrano\Arraya Solutions\Jefferson_Matched-Mailboxes.csv"

### Function ###
Start-OneDriveMigrationShareGate -SourceURL $SourceURL -SourceCredentials $sourceCredentials -DestinationURL $DestinationURL -DestinationCredentials $DestinationCredentials -DomainTLD EDU -Incremental -ImportCSV $filepath

## Example 1 End ##

## Example 2 Start ##
### Initial Variables ###
$SourceURL = "https://abaco1-admin.sharepoint.com/"
$DestinationURL = "https://ametekinc-admin.sharepoint.com/"
$SourceCredentials = Get-Credentials
#^Enter Migration Service Account
$DestinationCredentials = Get-Credential
#^Enter Migration Service Account
$filePath = "C:\Users\amedrano\Arraya Solutions\Ametek - External - 1639 Abaco - Tenant to Tenant Migration\Exchange Docs\AbacoMatched-Mailboxes.csv"

### Function ###
Start-OneDriveMigrationShareGate -SourceURL $SourceURL -SourceCredentials $sourceCredentials -DestinationURL $DestinationURL -DestinationCredentials $DestinationCredentials -DomainTLD COM -Incremental -ImportCSV $filepath

#>
function Start-OneDriveMigrationShareGate {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the CSV File Path of OneDrive Users')] [array] $ImportCSV,
        [Parameter(Mandatory=$True,HelpMessage="What is the Source Admin Site URL")][string] $SourceURL,
        [Parameter(Mandatory=$True,HelpMessage="What is the Destination Admin Site URL")] [string] $DestinationURL,
        [Parameter(Mandatory=$false,HelpMessage="Domain TLD. IE com, edu")][string] $DomainTLD,
        [Parameter(Mandatory=$false,HelpMessage="Test OneDrive For Users?")][switch] $Test,
        [Parameter(Mandatory=$True)] 
        [System.Management.Automation.PSCredential] 
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()] $SourceCredentials,
        [Parameter(Mandatory=$True)] 
        [System.Management.Automation.PSCredential] 
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()] $DestinationCredentials,
        [Parameter(Mandatory=$false,HelpMessage="Run Incremental Sync of OneDrive?")][switch] $Incremental,
        [Parameter(Mandatory=$false,HelpMessage="Which Wave?")][String] $WaveGroup
    )
    #Set Up Module, Variables, Credentials, and Connect to SharePoint Sites
    Import-Module Sharegate
    Import-Module Microsoft.Online.SharePoint.PowerShell
    Set-Variable srcSite, dstSite, srcList, dstList, srcSiteUrl, dstSiteUrl, dstSiteUrlInitial, dstSiteUrl,destinationUPN
    $AllOneDriveErrors = @()
    $AllOneDriveResults = @()
    $sourceTenant = Connect-Site -Url $SourceURL -Credential $SourceCredentials
    $destinationTenant = Connect-Site -Url $DestinationURL -Credential $DestinationCredentials
    $OneDriveUsers = Import-Csv $ImportCSV
    $DestinationServiceAccount = $DestinationCredentials.Username
    

    #Progress Bar Initial
    $progressref = ($OneDriveUsers).count
    $progresscounter = 0

    foreach ($user in $OneDriveUsers) {
        #Clear Previous Variables
        Clear-Variable srcSite, dstSite, srcList, dstList, srcSiteUrl, dstSiteUrlInitial, dstSiteUrl, destinationUPN
        $srcSite = @()
        $dstSite = @()
        $srcList = @()
        $dstList = @()
        $dstSiteUrlInitial = @()
        $dstSiteUrl = @()

        #Progress Bar Current
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Submitting OneDrive Migration for $($user.UserPrincipalName_Source)"
        
        #Connect to OneDrive Sites in Source
        $sourceUPN = $user.UserPrincipalName_Source
        try {
            if ($srcSiteUrl = $user.SourceOneDriveURL) {
            }
            elseif ($srcSiteUrl = Get-OneDriveUrl -Tenant $sourceTenant -Email $user.PrimarySMTPAddress_Source -ErrorAction silentlycontinue) {
                if ($DomainTLD -eq "com") {
                    $srcSiteUrlShort = $srcSiteUrl.replace("_com/","_com")
                    $srcSiteUrl = $srcSiteUrlShort
                }
                elseif ($DomainTLD -eq "edu"){
                    $srcSiteUrlShort = $srcSiteUrl.replace("_edu/","_edu")
                    $srcSiteUrl = $srcSiteUrlShort
                }   
            }
            else {
                Connect-SPOService -Url $SourceURL -Credential $SourceCredentials
                elseif ($ODSourceURLCheck = Get-SPOSite -Filter "Owner -eq '$sourceUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true -ea SilentlyContinue) {
                $srcSiteUrl = $ODSourceURLCheck.url
                }
            }
            $srcSite = Connect-Site -Url $srcSiteUrl -Credential $SourceCredentials -ErrorAction Stop
            #Get OneDrive Documents
            $srcList = Get-List -Site $srcSite -Name "Documents" -ErrorAction Stop
        }
        catch {
            $currenterror = new-object PSObject
            $currenterror | add-member -type noteproperty -name "Commandlet" -Value $_.CategoryInfo.Activity
            $currenterror | Add-Member -type NoteProperty -Name "FailureActivity" -Value "UnableToFindSourceOneDrive" -Force
            $currenterror | Add-Member -type NoteProperty -Name "Tenant" -Value $sourceTenant.Site -Force
            $currenterror | Add-Member -type NoteProperty -Name "User" -Value $user.PrimarySMTPAddress_Source
            $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception) -Force
            $AllOneDriveErrors += $currenterror           
            continue
        }
        ##Check if Destination OneDrive is Enabled
        $destinationSMTPAddress = $user.PrimarySMTPAddress_Destination
        $destinationUPN = $user.UserPrincipalName_Destination
        if ($srcList) {
            if ($user.DestinationOneDriveURL) {
                $dstSiteUrl = $user.DestinationOneDriveURL
            }
            else {
                try {
                    $dstSiteUrlInitial = Get-OneDriveUrl -Tenant $destinationTenant -Email $destinationSMTPAddress -ErrorAction Stop
                    if ($DomainTLD -eq "com") {
                        $dstSiteUrlShort = $dstSiteUrlInitial.replace("_com/","_com")
                    }
                    elseif ($DomainTLD -eq "edu"){
                        $dstSiteUrlShort = $dstSiteUrlInitial.replace("_edu/","_edu")
                    }
                    $dstSiteUrl = $dstSiteUrlShort
                }
                catch {
                    Connect-SPOService -Url $DestinationURL -Credential $DestinationCredentials
                    $OneDriveDestinationURLCheck = Get-SPOSite -Filter "Owner -eq '$destinationUPN' -and URL -like '*-my.sharepoint*'" -IncludePersonalSite $true
                    $dstSiteUrl = $OneDriveDestinationURLCheck.Url
                }
            }
            #Connect to OneDrive Sites in Destination
            try {
                $dstSite = Connect-Site -Url $dstSiteUrl -Credential $DestinationCredentials -ErrorAction Stop
                #Get OneDrive Documents
                $dstList = Get-List -Site $dstSite -Name "Documents" -ErrorAction Stop
            }
            catch {                
                #If Migration Service Account is not enabled. Add account as Secondary Admin
                Set-SPOUser -Site $dstSiteUrl.tostring() -LoginName $DestinationServiceAccount.tostring() -IsSiteCollectionAdmin $true -ErrorAction Stop
                Write-Host "$($DestinationServiceAccount) Added as Site Admin. " -ForegroundColor Green -nonewline 

                while (!(Get-SPOUser -Site $dstSiteUrl.tostring() -ErrorAction SilentlyContinue)) {
                        Write-Host " ." -NoNewline -foregroundcolor yellow
                        Start-Sleep -s 3
                    }
                #Attempt again to connect to Destination OneDrive
                $dstSite = Connect-Site -Url $dstSiteUrl -Credential $DestinationCredentials -ErrorAction Stop
                #Get OneDrive Documents
                $dstList = Get-List -Site $dstSite -Name "Documents" -ErrorAction Stop
                
            }
        }    

        Write-Host $srcList.Address.AbsoluteUri $dstList.Address.AbsoluteUri

        #Progress Bar Current 2

        $OneDriveResults = New-Object PSObject
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceAdminTenantURL" -Value $sourceTenant.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationAdminTenantURL" -Value $destinationTenant.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceName" -Value $srcSite.Title
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SourceSite" -Value $srcSite.Address
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationName" -Value $dstSite.Title
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "DestinationURL" -Value $dstSite.Address

        #Copy OneDrive Files from Source to Destination
        if ($Incremental){
            #Progress Bar Current 2
            Write-progress -id 2 -Activity "Submitting Incremental OneDrive Migration"
            $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
            $TaskName = "Incremental OneDrive Migration for $($srcSite.Title) to $($dstSite.Title)"
            $TaskName

            #Test Move with Incremental using Insane Mode
            if ($Test) {
                $Result = Copy-Content -SourceList $srcList -DestinationList $dstList -InsaneMode -CopySettings $copysettings -TaskName $TaskName -warningaction silentlycontinue -whatif
            }
            #Migrate Data with Incremental using Insane Mode
            else {
                $Result = Copy-Content -SourceList $srcList -DestinationList $dstList -InsaneMode -CopySettings $copysettings -TaskName $TaskName -warningaction silentlycontinue
            }
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value "Incremental"
        }
        else {
            #Progress Bar Current 2
            Write-progress -id 2 -Activity "Submitting Initial OneDrive Migration"
            $TaskName = "Initial OneDrive Migration for $($srcSite.Title) to $($dstSite.Title)"

            #Test Move using Insane Mode
            if ($Test) {
                $Result = Copy-Content -SourceList $srcList -DestinationList $dstList -InsaneMode -TaskName $TaskName -warningaction silentlycontinue -whatif
            }
            #Migrate Data using Insane Mode
            else {
                $Result = Copy-Content -SourceList $srcList -DestinationList $dstList -InsaneMode -TaskName $TaskName -warningaction silentlycontinue
            }
            $OneDriveResults | Add-Member -MemberType NoteProperty -Name "SyncType" -Value "Initial"
        }

        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result.Result
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "ItemsCopied" -Value $Result.ItemsCopied
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Successes" -Value $Result.Successes
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Errors" -Value $Result.Errors
        $OneDriveResults | Add-Member -MemberType NoteProperty -Name "Warnings" -Value $Result.Warnings
        $AllOneDriveResults += $OneDriveResults
        $OneDriveResults | Export-Csv "$HOME\Desktop\OneDrive-MigrationResults.csv" -NoTypeInformation -Encoding UTF8 -Append
        $AllOneDriveErrors | Export-Csv "$HOME\Desktop\OneDrive-MigrationErrors.csv" -NoTypeInformation -Encoding UTF8 -Append
        
    }
}

