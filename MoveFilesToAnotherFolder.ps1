# Move Files from One Folder to Another
$sourceFolderPath = Read-Host "What's the Folder path of the Folder to move from"
$destinationFolderPath = Read-Host "What's the Destination Folder path to move files to"
$NumberOfFiles = Read-Host "What's the Number of Files to move"

#Get the All the Child Items in folder
$folderItems = Get-ChildItem -Path $sourceFolderPath -Recurse 
$progressref = ($folderItems).count
$progresscounter = 0
 foreach ($item in $folderItems[0..$NumberOfFiles]) {
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Moving $($item.Name) to another folder"

    $sourceFilePath = $sourceFolderPath + "\" + $item.name
    
    Move-Item $sourceFilePath -Destination $destinationFolderPath
}

# Recursive Folders

## Move Files from One Folder to Another
$sourceFolderPath = Read-Host "What's the Top Level Folder path of the Folder to move from"

$NumberOfFiles = Read-Host "What's the Number of Files to move"

## Get the All the Child Items in folder
$Subfolders =  Get-ChildItem -path $sourceFolderPath -Directory -recurse

#$folderItems = Get-ChildItem -Path $sourceFolderPath -Recurse 
$progressref = ($Subfolders).count
$progresscounter = 0

foreach ($folder in $Subfolders) {
    $folderItems = Get-ChildItem -Path $folder.name -Recurse 
    $progressref = ($folderItems).count
    $progresscounter = 0

    $destinationFolderPath = $sourceFolderPath + "\" $folder.name + "\" + $folder.name + "A"

    foreach ($item in $folderItems[0..$NumberOfFiles]) {
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Moving $($item.Name) to another folder"

        $sourceFilePath = $sourceFolderPath + "\" + $item.name
        
        Move-Item $sourceFilePath -Destination $destinationFolderPath
    }
}

D:\Cloudmap\Aaron = SourceTopLevel Folder

if folder1 -gt 10000000000

D:\Cloudmap\Aaron\Folder1 = Old Folder
move X items to 
D:\Cloudmap\Aaron\Folder1A = New Folder



## Import CSV of Folders to move files from multiple Folders to Another

$importCSVPathName = Read-Host "What's the FolderPath of the CSV Import Job"

$importCSVFolders = Import-csv $importCSVPathName

#Get the All the Child Items in folder
$progressref1 = ($importCSVFolders).count
$progresscounter1 = 0

foreach ($folder in $importCSVFolders) {
    #Progress Bar 1 for Folders
    $progresscounter1 += 1
    $progresspercentcomplete1 = [math]::Round((($progresscounter1 / $progressref1)*100),2)
    $progressStatus1 = "["+$progresscounter1+" / "+$progressref1+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete1 -Status $progressStatus1 -Activity "Working on $($folder.SourceFolder) folder"
    
    #Create Variables
    $sourceFolderPath = $folder.SourceFolder
    $destinationFolderPath = $folder.DestinationFolder

    #Get Items per folder
    $folderItems = Get-ChildItem -Path $sourceFolderPath -Recurse

    #Folder Item Move Progress Bar
    $progressref = ($folderItems).count
    $progresscounter = 0
    $NumberOfItemsToMove = $folder.NumberOfItemsToMove

    foreach ($item in $folderItems[0..$NumberOfItemsToMove]) {
        #Folder Item Move Progress Bar
        $progresscounter += 1
        $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
        $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
        Write-progress -id 2 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Moving $($item.Name) to another $($destinationFolderPath)"
    
        #SourceFilePath
        $sourceFilePath = $sourceFolderPath + "\" + $item.name
        
        #Move Items from SourceFilePath to DestinationFolderPath
        Move-Item $sourceFilePath -Destination $destinationFolderPath
    }
}
 