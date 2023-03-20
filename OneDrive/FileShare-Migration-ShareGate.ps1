Import-Module Sharegate
$csvFile = "C:\temp\onedrivemigration.csv"
$FileShareMigrations = Import-Csv $csvFile -Delimiter ","
$DestinationAddress = "https://brentwoodindustries-admin.sharepoint.com/"
$dstsiteConnection = Connect-Site -Url $DestinationAddress -Browser 

Set-Variable dstSite, dstList


#Individual User
Clear-Variable dstSite
Clear-Variable dstList
$SourceFolder = "K:\Shares\Groups"
$DestinationList = "https://brentwoodindustries.sharepoint.com/sites/Fileshares"

#Connect to Destination
$dstSite = Connect-Site -Url $DestinationList -UseCredentialsFrom $dstsiteConnection -ErrorAction Stop
#Get OneDrive Documents
$dstList = Get-List -Site $dstSite -Name "Documents" -ErrorAction Stop
#Create Empty Folder
Import-Document -SourceFilePath "C:\temp\FileShare" -DestinationList $dstList
#Import Documents
$TaskName = "Import File Share to $($dstList.Address.ToString())"

$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
Import-Document -SourceFolder $SourceFolder -DestinationList $dstList -DestinationFolder "FileShare" -TaskName $TaskName -copySettings $copySettings
Import-Document -SourceFolder $SourceFolder -DestinationList $dstList -TaskName $TaskName -copySettings $copySettings


#terminated users
$SourceFolder = "E:\SHARES\Users\twenrich"
$DestinationFolder = "terminated users/Tom Wenrich"
$DestinationList = "https://brentwoodindustries.sharepoint.com/sites/Fileshares"
$Office = "India"

#Connect to Destination
$dstSite = Connect-Site -Url $DestinationList -UseCredentialsFrom $dstsiteConnection -ErrorAction Stop
#Get OneDrive Documents
$dstList = Get-List -Site $dstSite -Name $Office -ErrorAction Stop
#Create Empty Folder
$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
#Import Documents
$TaskName = "Import File Share $($SourceFolder) to $($dstList.Address.ToString()) in $($DestinationFolder)"
Import-Document -SourceFolder $SourceFolder -DestinationList $dstList -DestinationFolder $DestinationFolder -TaskName $TaskName -copySettings $copySettings


# Users to OneDrive!

$Office = "India"
$progressref = ($FileShareMigrations | ? {$_.Office -eq $Office}).count
$progresscounter = 0
Set-Variable dstSite, dstList
foreach ($obj in $FileShareMigrations | ? {$_.Office -eq $Office}) {
    Clear-Variable dstSite
    Clear-Variable dstList
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Migratng OneDrive for $($obj.user) in $($Office)"
    if ($obj.DestinationURL -like "*sharepoint*") {
        $SourceFolder = $obj.DIRECTORY
        $DestinationList = $obj.DestinationURL
        $DestinationFolder = "FileShare"
        #$DestinationFolder = $obj.DestinationFolder

        #Connect to Destination
        $dstSite = Connect-Site -Url $DestinationList -UseCredentialsFrom $dstsiteConnection -ErrorAction Stop
        #Get OneDrive Documents
        $dstList = Get-List -Site $dstSite -Name "Documents" -ErrorAction Stop
        #Set To Incremental Sync
        $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate

        <#Create Empty Folder
        Write-Host "Creating Folder for $($obj.user) in $($obj.office)" -foregroundcolor cyan
        #Import-Document -SourceFilePath "C:\temp\FileShare" -DestinationList $dstList -copySettings $copySettings
        #>

        #Import Documents
        Write-Host "Importing OneDrive Files Folder for $($obj.user) in $($obj.office)" -foregroundcolor cyan
        $TaskName = "Import $($obj.user) FileShare to OneDrive $($dstList.Address.ToString())"
        Import-Document -SourceFolder $SourceFolder -DestinationList $dstList -DestinationFolder $DestinationFolder -TaskName $TaskName -copySettings $copySettings
        # Remove-SiteCollectionAdministrator -Site $dstSite
    }
    else {
        Write-Host "No Site found for $($obj.user) in $($obj.office)" -foregroundcolor red
    }
}


# General Group - Shared
$Office = "India"
$progressref = ($FileShareMigrations | ? {$_.Office -eq $Office}).count
$progresscounter = 0
Set-Variable dstSite, dstList
foreach ($obj in $FileShareMigrations | ? {$_.Office -eq $Office}) {
    Clear-Variable dstSite
    Clear-Variable dstList
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Migratng FileShare for $($obj.user) in $($Office)"
    if ($obj.DestinationURL -like "*sharepoint*") {
        $SourceFolder = $obj.DIRECTORY
        $DestinationList = $obj.DestinationURL
        $DestinationFolder = $obj.DestinationFolder

        #Connect to Destination
        $dstSite = Connect-Site -Url $DestinationList -UseCredentialsFrom $dstsiteConnection -ErrorAction Stop
        #Get OneDrive Documents
        $dstList = Get-List -Site $dstSite -Name $Office -ErrorAction Stop
        #Set To Incremental Sync
        $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate

        <#Create Empty Folder
        Write-Host "Creating Folder for $($obj.user) in $($obj.office)" -foregroundcolor cyan
        #Import-Document -SourceFilePath "C:\temp\FileShare" -DestinationList $dstList -copySettings $copySettings
        #>

        #Import Documents
        Write-Host "Importing FileShare Files for $($obj.user) in $($obj.office)" -foregroundcolor cyan
        $TaskName = "Import $($obj.user) FileShare to SharePoint $($dstList.Address.ToString())"
        Import-Document -SourceFolder $SourceFolder -DestinationList $dstList -DestinationFolder $DestinationFolder -TaskName $TaskName -copySettings $copySettings
        # Remove-SiteCollectionAdministrator -Site $dstSite
    }
    else {
        Write-Host "No Site found for $($obj.user) in $($obj.office)" -foregroundcolor red
    }
}

function Start-ShareGateFileShareMigration {
    param (
    [Parameter(Mandatory=$True,Position=1,HelpMessage="What is the Destination Admin SharePoint Site URL?")] [string] $DestinationSPOAdminSite,
    [Parameter(ParameterSetName='MultipleUse',Position=3,Mandatory=$false,HelpMessage="Import a CSV or Excel")] [switch]$MultipleUse,
    [Parameter(ParameterSetName='MultipleUse',Mandatory=$True,HelpMessage="Run against Office365 Exchange Online?")] [string]$InputCSVFilePath,
    [Parameter(ParameterSetName='MultipleUse',Mandatory=$True,HelpMessage="Run against Office365 Exchange Online?")] [string]$InputEXCELFilePath,
    [Parameter(ParameterSetName='MultipleUse',Mandatory=$True,HelpMessage="Run against Office365 Exchange Online?")] [string]$InputExcelWorkSheetName,
    [Parameter(ParameterSetName='SingleUse',Position=2,Mandatory=$false,HelpMessage="SingleUse")] [switch]$SingleUse,
    [Parameter(ParameterSetName='SingleUse',Mandatory=$True,HelpMessage="What is the Source Folder Path?")] [string]$SourceFolder,
    [Parameter(ParameterSetName='SingleUse',Mandatory=$True,HelpMessage="What is the Destination Site URL?")] [string]$DestinationSite,
    [Parameter(ParameterSetName='SingleUse',Mandatory=$false,HelpMessage="What is the Destination Folder Name?")] [string]$DestinationLibrary,
    [Parameter(ParameterSetName='SingleUse',Mandatory=$True,HelpMessage="What is the Destination Folder Name?")] [string]$DestinationFolder,
    [Parameter(Mandatory=$false,HelpMessage="Create Empty Folder?")] [switch]$CreateEmptyFolder,
    [Parameter(Mandatory=$true,HelpMessage="Empty Folder Name?")] [string]$FolderName
    )

    #Enable ShareGate Module
    Import-Module Sharegate

    #Global Variables
    [Global]$CreateEmptyFolder = $CreateEmptyFolder
    [Global]$FolderName = $FolderName
    
    #Connect Using Modern Auth
    [Global]$dstsiteConnection = Connect-Site -Url $DestinationSPOAdminSite -Browser

    #Import Document Function
    function Import-ShareGateFiles {
        param (
            [Parameter(Mandatory=$True,HelpMessage="What is the Source Folder Path?")] [string]$SourceFolderPath,
            [Parameter(Mandatory=$True,HelpMessage="What is the Destination Site URL?")] [string]$DestinationSiteAddress,
            [Parameter(Mandatory=$False,HelpMessage="What is the Destination Folder Name?")] [string]$DestinationLibraryName,
            [Parameter(Mandatory=$False,HelpMessage="What is the Destination Folder Name?")] [string]$DestinationFolderName
        )
            Clear-Variable dstSite
            Clear-Variable dstList
        #Connect to Destination
        $dstSite = Connect-Site -Url $DestinationSiteAddress -UseCredentialsFrom [global]$dstsiteConnection -ErrorAction Stop
        if ($DestinationLibraryName) {
            $dstList = Get-List -Site $dstSite -Name $DestinationLibraryName -ErrorAction Stop
            #Create Task
            $TaskName = "Import File Share $($SourceFolderPath) to $($dstList.Address.ToString()) in $($DestinationLibraryName)"
        }
        else {
            $dstList = Get-List -Site $dstSite -Name "Documents" -ErrorAction Stop
            $TaskName = "Import File Share $($SourceFolderPath) to $($dstList.Address.ToString())"
        }
        #Set To Incremental Sync
        $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
        Write-Host $TaskName -ForegroundColor Cyan
        #Test Folder Exists
        if ([Global]$CreateEmptyFolder) { 
            $FolderPath = "C:\temp" + [Global]$FolderName
            if (Test-Path $FolderPath) {
                Write-Host "$([Global]$FolderName) Exists"
                # Perform Delete file from folder operation
            }
            else {
                #PowerShell Create directory if not exists
                New-Item $FolderName -ItemType Directory
                Write-Host "Folder $([Global]$FolderName) Created successfully"
            }
            #Migrate Empty Folder
            Write-Host "Creating Destination Folder for $($SourceFolder) in $($dstList.Address.ToString())" -foregroundcolor cyan
            Import-Document -SourceFilePath $SourceFolder -DestinationList $dstList -copySettings $copySettings
        }
        #Import to Specific Folder
        if ($DestinationFolderName) {
            Import-Document -SourceFolder $SourceFolder -DestinationList $dstList -DestinationFolder $DestinationFolder -TaskName $TaskName -copySettings $copySettings
        }
        #Import to Document Library Directly
        else {
            Import-Document -SourceFolder $SourceFolder -DestinationList $dstList -TaskName $TaskName -copySettings $copySettings
        }
    }

    if ($MultipleUse) {
        if ($InputCSVFilePath) {
            $FileShareMigrations = Import-Csv $InputCSVFilePath
        }
        elseif ($InputEXCELFilePath) {
            if ($InputExcelWorkSheetName) {
                $FileShareMigrations = Import-Excel -Path $InputEXCELFilePath -WorksheetName $InputExcelWorkSheetName
            }
            else {
                $FileShareMigrations = Import-Excel -Path $InputEXCELFilePath
            }
        }
        #Progres Bar
        $progressref = ($FileShareMigrations).count
        $progresscounter = 0
        Set-Variable dstSite, dstList
        foreach ($obj in $FileShareMigrations) {
            #Progress Bar
            $progresscounter += 1
            $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
            $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
            Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Migratng FileShare for $($obj.SourcePath) in $($obj.DestinationURL)"
            if ($obj.Library) {
                if ($obj.DestinationFolder) {
                    Import-ShareGateFiles -SourceFolderPath $obj.SourcePath -DestinationSiteAddress $obj.DestinationURL -DestinationLibraryName $obj.Library -DestinationFolderName $obj.DestinationFolder
                }
                else {
                    Import-ShareGateFiles -SourceFolderPath $obj.SourcePath -DestinationSiteAddress $obj.DestinationURL -DestinationLibraryName $obj.Library
                }
            }
            else {
                Import-ShareGateFiles -SourceFolderPath $obj.SourcePath -DestinationSiteAddress $obj.DestinationURL
            }
        }
    }
    elseif ($SingleUse) {
        if ($DestinationLibrary) {
            if ($DestinationFolder) {
                Import-ShareGateFiles -SourceFolderPath $SourceFolder -DestinationSiteAddress $DestinationSite -DestinationLibraryName $DestinationLibrary -DestinationFolderName $DestinationFolder
            }
            else {
                Import-ShareGateFiles -SourceFolderPath $SourceFolder -DestinationSiteAddress $DestinationSite -DestinationLibraryName $DestinationLibrary
            }
        }
        else {
            Import-ShareGateFiles -SourceFolderPath $SourceFolder -DestinationSiteAddress $DestinationSite
        }
    }
    
}