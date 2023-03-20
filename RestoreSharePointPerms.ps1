### Set SharePoint Group Perm ###
$updatingTeams = Import-excel "C:\Users\amedrano\Arraya Solutions\Thomas Jefferson - Einstein to Jefferson Migration\ShareGate\Teams\MatchedTeams.xlsx"

#Variables for Admin Center & Site Collection URL
$AdminCenterURL = "https://tjuv-admin.sharepoint.com"
#Connect to SharePoint Online
Connect-SPOService -url $AdminCenterURL -Credential (Get-Credential)

$MembersGroups = $matchedTeams | ?{$_.GroupTitle -like "*Members" -and $_.GroupPerms -like "*Limited Access*"}
$progressref = $MembersGroups.count
$progresscounter = 0
foreach ($object in $MembersGroups) {
    #Set Variables
    $GroupName = $object.GroupTitle
    $SiteURL = $object.SharePointSiteURL_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Setting MembersGroup Details $($SiteURL)"
       
    #Update Member Group to Add Edit Permissions
    Write-Host "Updating $($GroupName) Perms to Edit .." -ForegroundColor Cyan -nonewline
    $permResult = Set-SPOSiteGroup -Site $SiteURL -Identity $GroupName -PermissionLevelsToAdd "Edit"
    Write-Host "Completed" -ForegroundColor Green
}

## Restore Inheritance ###
#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
  
#To call a non-generic Load Method
Function Invoke-LoadMethod() {
    Param(
            [Microsoft.SharePoint.Client.ClientObject]$Object = $(throw "Please provide a Client Object"), [string]$PropertyName
         )
   $Ctx = $Object.Context
   $Load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load")
   $Type = $Object.GetType()
   $ClientLoad = $Load.MakeGenericMethod($Type)
   
   $Parameter = [System.Linq.Expressions.Expression]::Parameter(($Type), $Type.Name)
   $Expression = [System.Linq.Expressions.Expression]::Lambda([System.Linq.Expressions.Expression]::Convert([System.Linq.Expressions.Expression]::PropertyOrField($Parameter,$PropertyName),[System.Object] ), $($Parameter))
   $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
   $ExpressionArray.SetValue($Expression, 0)
   $ClientLoad.Invoke($Ctx,@($Object,$ExpressionArray))
}
 
#Function to Delete Unique Permission from all lists of a Web
Function Reset-SPOListPermission([Microsoft.SharePoint.Client.Web]$Web) {
    Write-host -f Magenta "Searching Unique Permissions on the Site:"$web.Url   
        
    #Get All Lists of the web
    $Lists =  $Web.Lists
    $Ctx.Load($Lists)
    $Ctx.ExecuteQuery()
 
    #Exclude system lists
    $ExcludedLists = @("App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
     "Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Style Library",
     ,"Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Search Config List", "Solution Gallery", "Site Collection Images",
     "Suggested Content Browser Locations","TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Preservation Hold Library")
     
    #Iterate through each list
    ForEach($List in $Lists)
    {
        #Get the List
        $Ctx.Load($List)
        $Ctx.ExecuteQuery()
 
        If($ExcludedLists -NotContains $List.Title -and $List.Hidden -eq $false)
        {
            #Check if the given site is using unique permissions
            Invoke-LoadMethod -Object $List -PropertyName "HasUniqueRoleAssignments"
            $Ctx.ExecuteQuery()
  
            #Reset broken inheritance of the list
            If($List.HasUniqueRoleAssignments)
            {
                #delete unique permissions of the List
                $List.ResetRoleInheritance()
                $List.Update()
                $Ctx.ExecuteQuery()   
                Write-host -f Green "`tUnique Permissions Removed from the List: '$($List.Title)'"
            }
        }
    }
 
    #Process each subsite in the site
    $Subsites = $Web.Webs
    $Ctx.Load($Subsites)
    $Ctx.ExecuteQuery()       
    Foreach ($SubSite in $Subsites)
    {
        #Call the function Recursively
        Reset-SPOListPermission($Subsite)
    }
}

## Run for single site ##
#Get Credentials to connect
$Cred = Get-Credential

#Config Parameters
#Example: $SiteURL= "https://crescent.sharepoint.com/sites/marketing"
$SiteURL = Read-Host "What Site needs to be updated? Provide full URL"

Try {
    #Setup the context
    $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.UserName,$Cred.Password)
      
    #Get the Web
    $Web = $Ctx.Web
    $Ctx.Load($Web)
    $Ctx.ExecuteQuery()
     
    #Call the function to delete unique permission from all lists of a site collection
    Reset-SPOListPermission $Web
}
Catch {
    write-host -f Red "Error:" $_.Exception.Message
}

## Set SharePoint Group Perm - Run for Multiple Sites  ###
#Variables for Admin Center & Site Collection URL
$AdminCenterURL = "https://tjuv-admin.sharepoint.com"
#Connect to SharePoint Online
$Cred = Get-Credential
Connect-SPOService -url $AdminCenterURL -Credential $Cred

$updatingTeams = Import-excel ~\TeamsSitesPermsRequireUpdate.xlsx
$MembersGroups = $updatingTeams | ?{$_.GroupTitle -like "*Members" -and $_.GroupPerms -like "*Limited Access*"}
$progressref = $MembersGroups.count
$progresscounter = 0
$AllErrors = @()
foreach ($object in $MembersGroups) {
    #Set Variables
    $GroupName = $object.GroupTitle
    $SiteURL = $object.SharePointSiteURL_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Updating Members Permissions Group and Restore Inheritance for $($SiteURL)"
       
    #Add Edit Permissions to Members Permissions Group
    Write-Host "Updating $($GroupName) Perms to Edit .." -ForegroundColor Cyan -nonewline
    $permResult = Set-SPOSiteGroup -Site $SiteURL -Identity $GroupName -PermissionLevelsToAdd "Edit" -PermissionLevelsToRemove "Web-Only Limited Access"
    Write-Host ". " -ForegroundColor Green -NoNewline

    #Restore Inheritance
    Try {
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.UserName,$Cred.Password)
        Write-Host "Restored Inheritance .." -ForegroundColor Cyan -nonewline
        
        #Get the Web
        $Web = $Ctx.Web
        $Ctx.Load($Web)
        $Ctx.ExecuteQuery()
        
        #Call the function to delete unique permission from all lists of a site collection
        Reset-SPOListPermission $Web
        Write-Host "Completed" -ForegroundColor Green
    }
    Catch {
        Write-Host "Failed" -ForegroundColor Red
        $currenterror = new-object PSObject
        $currenterror | add-member -type noteproperty -name "Activity" -Value $_.CategoryInfo.Activity
        $currenterror | Add-Member -type NoteProperty -Name "FailureStatus" -Value "Unable to Restore Inheritance" -Force
        $currenterror | Add-Member -type NoteProperty -Name "SiteUrl" -Value $SiteURL -Force
        $currenterror | Add-Member -type NoteProperty -Name "GroupName" -Value $GroupName -Force
        $currenterror | Add-Member -type NoteProperty -Name "Error" -Value ($_.Exception.Message) -Force
        $AllErrors += $currenterror 
    }
}

#Read more: https://www.sharepointdiary.com/2016/01/sharepoint-online-delete-unique-permissions-using-powershell.html#ixzz7TTRZf2A8

# Get Permissions of Members

$progressref = $MembersGroups.count
$progresscounter = 0
foreach ($object in $MembersGroups) {
    #Set Variables
    $GroupName = $object.GroupTitle
    $SiteURL = $object.SharePointSiteURL_Destination

    #Progress Bar
    $progresscounter += 1
    $progresspercentcomplete = [math]::Round((($progresscounter / $progressref)*100),2)
    $progressStatus = "["+$progresscounter+" / "+$progressref+"]"
    Write-progress -id 1 -PercentComplete $progresspercentcomplete -Status $progressStatus -Activity "Gathering Group Perms for $($SiteURL)"
    
    $permResult = @()
    #Get  Permissions to Members Permissions Group
    $permResult = Get-SPOSiteGroup -Site $SiteURL -Group $GroupName

    $object | add-member -type NoteProperty -Name "GroupPerms_Updated" -Value ($permResult.Roles -join ",") -Force
}