#region Functions

function Get-MainMenu
{
    
    Get-DomainVerification
    $mainSelection = $null
    while($mainSelection -lt 1 -or $mainSelection -gt 3)
    {
        cls
        Write-Host @"
Hosted Exchange Support Tool

----------------------------------------------
----------------------------------------------

Please select which Set of reports you would 
like to run:

----------------------------------------------

1. Domain Reports
2. Mailbox Reports

3. Tier II Support Tools

"@
        [int]$mainSelection = Read-Host "Select either 1, 2, or 3"
        Switch($mainSelection)
        {
            1{Get-DomainMenu}
            2{Get-MailboxMenu}
            3{Get-AdvancedMenu}
        }
    }
}

function Get-AdvancedMenu
{
    Check-L2PlusGroupAccess
    $AdvSelection = $null
    while($AdvSelection -lt 1 -or $AdvSelection -gt 3)
    {
        cls
        Write-Host @"
Advanced Support Menu

----------------------------------------------
----------------------------------------------

Please select which Set of reports you would 
like to run:

----------------------------------------------

1. Message Tracking
2. Check MSMQ
3. Restart Exchange Tools
4. IIS Logs
5. Set Impersonation Rights

"@
        [int]$mainSelection = Read-Host "Select either 1, 2, 3, 4, or 5"
        Switch($mainSelection)
        {
            1{Get-ExchMessageTracking}
            2{Get-MSMQCurrentQueue}
            3{Reset-ExchangeTools}
            4{Get-IISLogApp}
            5{Set-ImpersonationRights}
        }
    }
}

function Get-DomainMenu
{
    $Script:DomainSelection = $null
    while($Script:DomainSelection -lt 1 -or $Script:DomainSelection -gt 10)
    {
        cls
        Write-Host @"
Domain Reports

----------------------------------------------
----------------------------------------------

Please select which Set of reports you would 
like to run:

----------------------------------------------

1 - Run All Domain Reports

2 - All Distribution Groups members
3 - Pull Legacy Dn/X500 Addresses
4 - Pull All GAL Entries
5 - Resource Mailbox Sizes
6 - All Mailbox Sizes
7 - Get all Domain Users Exchange Features
8 - All Permissions for the Domain
9 - Domain Automapping
10 - Exit and Restart

"@
        $Script:DomainSelection = Read-Host "Select either 1, 2, 3, 4, 5, 6, 7, 8 or 9"
        Switch($Script:DomainSelection)
        {
            1{Get-AllDomainReports}
            2{Get-AllDistributionGroupMembers}
            3{Get-AllX500Addresses}
            4{Get-DomainGALObjects}
            5{Get-ResourceMailboxSizes}
            6{Get-DomainMailboxSize}
            7{Get-CASMailboxFeatures}
            8{Get-DomainPermissions}
            9{Get-DomainAutoMapping}
            10{Get-MainMenu}
        }
    }
}

function Get-MailboxMenu
{
    Get-MailboxVerification
    $mailboxSelection = $null
    while($mailboxSelection -lt 1 -or $mailboxSelection -gt 7)
    {
        cls
        Write-Host @"
Mailbox Reports

----------------------------------------------
----------------------------------------------

Please select which Set of reports you would 
like to run:

----------------------------------------------

1 - Run All Mailbox Reports

2 - All Distribution Lists a User is a member of
3 - All Folders a User has Permissions to
4 - Users Mailbox Folder Sizes
5 - Users Inbox Rules
6 - User AutoMapping
7 - Exit and Restart

"@
        [int]$mailboxSelection = Read-Host "Select either 1, 2, 3, 4, 5, 6, or 7"
        Switch($mailboxSelection)
        {
            1{Get-AllMailboxReports}
            2{Get-MailboxDLMembership}
            3{Get-MailboxFolderPermissions}
            4{Get-MailboxFolderSizes}
            5{Get-MailboxInboxRules}
            6{Get-MailboxAutoMapping}
            7{Get-MainMenu}
        }
    }    
}

function Check-L2PlusGroupAccess
{
    $userGroups = Get-AllUserGroups
    if($userGroups | ? {$_.Value -like "*Managed Mail L2Plus*" -or $_.Value -like "*Organization Management*"})
    {
        Write-Host -ForegroundColor Green "L2 Access Verified Opening Advanced Menu"
        Start-Sleep -Seconds 3
    }
    else
    {
        Write-Host -ForegroundColor Red "You are logged into an account that does not have proper permissions for this menu"
        Write-Host -ForegroundColor Red "Restarting Script"
        Start-Sleep -Seconds 3
        Get-MainMenu
    }
}

Function Get-AllUserGroups 
{
    [cmdletbinding()]
    param()
    $Groups = [System.Security.Principal.WindowsIdentity]::GetCurrent().Groups
    foreach ($Group in $Groups) 
    {
      $GroupSID = $Group.Value
      $GroupName = New-Object System.Security.Principal.SecurityIdentifier($GroupSID)
      $GroupDisplayName = $GroupName.Translate([System.Security.Principal.NTAccount])
      $GroupDisplayName
    }
}

function Clean-ExportFolder
{
    $date = (get-date).AddDays(-7)
    $files = Get-ChildItem $Script:mydesk\exports | ?{$_.CreationTime -lt $date}
    if ($files)
    {
        Write-Host -ForegroundColor Yellow -NoNewline "Cleaning up $($file.count) files that are older than 7 days from Exports...."
        $files | Remove-Item -force | Out-Null
        Write-Host -ForegroundColor Green "Done"
        Start-Sleep -seconds 2
    }
}

function Get-ExchMessageTracking
{
    cls
    $sender = ""
    $recipient = ""
    $messageID = ""
    $startDate = ""
    $endDate = ""
    write-host "Running Message Tracking for $Script:Domain"
    $csv = $Script:us + "-MessageTracking.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv    
    }   
    Write-Host " "
    Write-Host -ForegroundColor Red "Note: Please do not run message tracking for more than a week"
    Write-Host " "
    $messageID = Read-Host "Message ID? (Blank if None)"
    if($messageID -eq "")
    {
        $sender = Read-Host "Who is the Sender? (Blank if None)"
        $recipient = Read-Host "Who is the Recipient? (Blank if None)"
    }
    Write-Host "Input date to export logging from (leave blank for today)." -ForegroundColor Cyan
    Write-Host "If you are wanting to export multiple days this will be the most recent day." -ForegroundColor Cyan
    $enddate = Read-Host "Date (MM/DD/YYYY)"
	if(!$enddate)
	{
	    $enddate = Get-Date -format g
	    Write-Host " "
	    Write-Host "No date specified...parsing against todays logs." -ForegroundColor Cyan
	}
    else 
    {
        $endDate = $endDate | Get-Date -Format g
        Write-Host " "
    }
    

    # establish date range array #
    Write-Host "Enter how many additional days back you would like to parse (Default: 0)" -ForegroundColor Cyan
    [int]$daystoparse = Read-Host "Number of days (max=7)"
    $daystoparse = $daystoparse + 1
    $startDate = (Get-Date $enddate).AddDays(-($daystoparse))
    $startDate = Get-Date $StartDate -Format g
    write-host "Running Message Tracking"
    $csv = $Script:Domain + "-MessageTracking.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv    
    }   
    Write-Host -foregroundcolor green -nonewline "Pulling Messagetracking logs...."
    try
    {
        if($sender -ne "" -and $messageID -eq "" -and $recipient -ne "")
        {
            Write-Host "Sender and Recipient"
            Get-TransportService | get-messagetrackinglog -resultsize unlimited -recipient $recipient -sender $sender -Start $startDate -End $endDate | select Timestamp,EventId,{$_.Sender},RecipientCount,{$_.Recipients},MessageSubject,ReturnPath,RelatedRecipientAddress,{$_.RecipientStatus},InternalMessageId,MessageId,Client-Ip,ClientHostname,Server-Ip,ServerHostname,SourceContext,ConnectorId,Source,TotalBytes,{$_.Reference},MessageInfo | sort-object Timestamp | export-csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
            Write-Host -foregroundcolor cyan "Mission Complete. Export is on your Desktop"
        }
        elseif($sender -eq "" -and $messageID -eq "" -and $recipient -ne "")
        {
            Write-Host "Recipient"
            Get-TransportService | get-messagetrackinglog -resultsize unlimited -recipient $recipient -Start $startDate -End $endDate -verbose | select Timestamp,EventId,Sender,RecipientCount,{$_.Recipients},MessageSubject,ReturnPath,RelatedRecipientAddress,{$_.RecipientStatus},InternalMessageId,MessageId,Client-Ip,ClientHostname,Server-Ip,ServerHostname,SourceContext,ConnectorId,Source,TotalBytes,{$_.Reference},MessageInfo | sort-object Timestamp | export-csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation            
            Write-Host -foregroundcolor cyan "Complete"
        }
        elseif($recipient -eq "" -and $messageID -eq "" -and $sender -ne "")
        {
            Write-Host "Sender Only"
            Get-TransportService | get-messagetrackinglog -resultsize unlimited -sender $sender -Start $startDate -End $endDate | select Timestamp,EventId,Sender,RecipientCount,{$_.Recipients},MessageSubject,ReturnPath,RelatedRecipientAddress,{$_.RecipientStatus},InternalMessageId,MessageId,Client-Ip,ClientHostname,Server-Ip,ServerHostname,SourceContext,ConnectorId,Source,TotalBytes,{$_.Reference},MessageInfo | sort-object Timestamp | export-csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
            Write-Host -foregroundcolor cyan "Complete"
        }
        elseif($messageID -ne "" -and $sender -eq "" -and $recipient -eq "")
        {
            Write-Host "MessageID"
            Get-TransportService | get-messagetrackinglog -resultsize unlimited -messageid $messageID -Start $endDate -End $startDate | select Timestamp,EventId,Sender,RecipientCount,{$_.Recipients},MessageSubject,ReturnPath,RelatedRecipientAddress,{$_.RecipientStatus},InternalMessageId,MessageId,Client-Ip,ClientHostname,Server-Ip,ServerHostname,SourceContext,ConnectorId,Source,TotalBytes,{$_.Reference},MessageInfo | sort-object Timestamp | export-csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
            Write-Host -foregroundcolor cyan "Complete"
        }

    }
    Catch
    {
        Write-Host " "
        Write-Host "There was a problem, please check input"
        Write-Host " "
        Write-Host " "
    }
}

Function Get-MGMTCredentials
{
    $PasswordFile = "C:\scripts\Utility_Server_Tool\pass.txt"
    $KeyFile = "C:\scripts\Utility_Server_Tool\AES.key"
    $User = "mgmt\svc_script"
    $key = Get-Content $KeyFile
    $Script:MGMTCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
}

function Get-LocalCredentials
{
    $localEnv = $Script:environment.Split(".")[0]
    $PasswordFile = "C:\scripts\Utility_Server_Tool\pass.txt"
    $KeyFile = "C:\scripts\Utility_Server_Tool\AES.key"
    $User = "$($localENV)\svc_script"
    $key = Get-Content $KeyFile
    $Script:LocalCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
}

function Reset-ExchangeTools
{
    Switch($Script:environment)
    {
        "mex05.mlsrvr.com"{$etServer = "ord2script01.mex05.mlsrvr.com"}
        "mex06.mlsrvr.com"{$etServer = "et01-ord1.mex06.mlsrvr.com"}
        "mex08.mlsrvr.com"{$etServer = "Script01-ORD1.mex08.mlsrvr.com"}
        "mex09.mlsrvr.com"{$etServer = "Script01-ORD1.mex09.mlsrvr.com"}
        default{Write-Host -ForegroundColor Red "Environment could not be determained. Closing Tool";Break}
    }
    Get-LocalCredentials
    Write-Host -ForegroundColor Cyan "Restarting Exchange Tools on $etServer"
    Invoke-Command -Credential $Script:LocalCred -ComputerName $etServer -ScriptBlock {Get-Service ExchangeToolsDirector | Restart-Service}
}

function Get-MSMQCurrentQueue
{
    Get-MGMTCredentials
    $NumberOfMessages = Read-Host "What Threshold do you want?"
    $srvList = "iad3-msmq01.mgmt.mlsrvr.com"
    $output = @()
    foreach ($srv in $srvList)
    {
        $counter = 0
        Write-Host("Checking : " + $srv)
        $Result = Invoke-Command -ComputerName IAD3-MSMQ01.mgmt.mlsrvr.com -Credential $Script:MGMTcred -ScriptBlock {Get-MsmqQueue -QueueType Private | ?{$_.path -like "*Domain*"}}
        if ($Result -ne $null)
	    {	
		    foreach ($a in $Result)
	        {
			    if ($a.MessageCount -gt $NumberOfMessages) 
			    {
		            $name = $a.path
                    $name = $name.split("\")[2]
                    $domain = $name.Split(".")[3]
                    $tld = $name.Split(".")[4]
                    $domainName = "$domain.$tld"
                    $output += "$domainName currently has $($a.MessageCount) messages in the queue"
			    }
		    }
        }
        if ($output)
        {
            $output
        }
        else
        {
            Write-Host -ForegroundColor Green "Queues are currently in good standing."
        }
    }
    Read-Host "Press Any Key to Continue"
}

function Send-EmailReport
{
    $answer2 = Read-Host ("Do you wish to email the CSV Exports to your Rackspace Email Address? Y/N")
    if ($answer2 -eq "y")
    {
        $i = 0
        do
        {
            $to = Read-Host ("Please Enter Your Rackspace Email Address:")
            $regex = "^[a-z]+\.[a-z]+@rackspace.*$"
            If ($to –notmatch $regex)
            {
                Write-Host ("Invalid Email Address $to")
            }
            else
            {
                            
                write-host $domain
                $env = (Get-WmiObject Win32_ComputerSystem).Domain
                $reportdate = Get-date -Format yyyy-MM-dd
                #Connection Details
                $smtpServer = “smtp.emailsrvr.com”
                $msg = new-object Net.Mail.MailMessage

                #Change port number for SSL to 587
                $smtp = New-Object Net.Mail.SmtpClient($SmtpServer, 25) 

                #Uncomment Next line for SSL  
                #$smtp.EnableSsl = $true

                #From Address
                $msg.From = "Reports@$($env)"
                #To Address, Copy the below line for multiple recipients
                $msg.To.Add($to)

                $msg.Subject = $Domain + "-" + $reportdate
                $msg.body = “

                Hello,

                Here is your Report for $domain.

                Have a great day!

                Your Rackspace Team

                "

                #your file location
                $files=Get-ChildItem “$mydesk\exports\$domain”

                Foreach($file in $files)
                {
                    Write-Host “Attaching File :- ” $file
                    $attachment = New-Object System.Net.Mail.Attachment –ArgumentList $mydesk\exports\$domain\$file
                    $msg.Attachments.Add($attachment)
                }
                $smtp.Send($msg)
                $attachment.Dispose();
                $msg.Dispose();
                $i++
                }
            }
            while($i -lt 1)
        }
}

function Get-AllMailboxReports
{
    Get-MailboxDLMembership
    Get-MailboxFolderPermissions
    Get-MailboxFolderSizes
    Get-MailboxInboxRules
    Get-MailboxAutoMapping
}

function Get-MailboxVerification
{
    $input = Read-Host ("Please Enter a Email Address")
    $Script:User = $input
    $Split = $Script:User -split "@"
    $Script:us = $($split[0])
    $domain2 = $($Split[1])
    if ($Script:Domain -ne $domain)
    {
        Write-Host -ForegroundColor Red "Mailbox is not a member of $script:Domain"
        Start-Sleep -seconds 1
        Get-MainMenu
    }
    else
    {
        Write-Host -ForegroundColor Green "Mailbox is a member of $script:Domain"
    }
}

function Get-MailboxInboxRules
{
    write-host "Running Report Inbox Rules for $User"
    $csv = $Script:us + "-Rules.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:Domain\$csv    
    }   
    Get-InboxRule -Mailbox $Script:User | select Name,@{Name=’Description’;Expression={[string]::join(";", ($_.Description))}} | Export-Csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
}

function Get-MailboxAutoMapping
{
    write-host "Running Report AutoMapping for $script:User"
    $csv = "$($Script:User)-Automapping.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:Domain\$csv    
    } 
    $mailbox = Get-Mailbox $script:User -ResultSize Unlimited
    $output = @()
    $distinguishedName = $mailbox.distinguishedName
    $parentmailbox = $mailbox.PrimarySMTPAddress
    $fullMailbox = Get-AdUser $distinguishedName -properties msExchDelegateListLink | Select msExchDelegateListLink
    if ($fullMailbox.msExchDelegateListLink -ne $null)
    {
        $automapping = $($Fullmailbox.msExchDelegateListLink).split(" ")
        foreach ($autom in $automapping)
        {
            $child = Get-Mailbox -identity $autom | Select PrimarySMTPAddress
            $MailboxAutoMap = New-Object -TypeName PSObject
            $MailboxAutoMap | Add-Member -MemberType NoteProperty -Name AutomappingTo -Value $child.PrimarySMTPAddress
            $MailboxAutoMap | Add-Member -MemberType NoteProperty -Name Mailbox -Value $parentMailbox
            $output += $MailboxAutoMap
        }
    }
    $output | Sort AutoMappingTo | Export-CSV $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
}

function Get-DomainAutoMapping
{
    write-host "Running Report Domain AutoMapping for $Script:Domain"
    $csv = "Domain-Automapping.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:Domain\$csv    
    } 
    $mailboxes = Get-Mailbox -OrganizationalUnit $Script:Domain -ResultSize Unlimited
    $output = @()
    foreach ($mailbox in $mailboxes)
    {
        $distinguishedName = $mailbox.distinguishedName
        $parentmailbox = $mailbox.PrimarySMTPAddress
        $fullMailbox = Get-AdUser $distinguishedName -properties msExchDelegateListLink | Select msExchDelegateListLink
        if ($fullMailbox.msExchDelegateListLink -ne $null)
        {
            $automapping = $($Fullmailbox.msExchDelegateListLink).split(" ")
            foreach ($autom in $automapping)
            {
                $child = Get-Mailbox -identity $autom | Select PrimarySMTPAddress
                $MailboxAutoMap = New-Object -TypeName PSObject
                $MailboxAutoMap | Add-Member -MemberType NoteProperty -Name AutoMappingGranted -Value $child.PrimarySMTPAddress
                $MailboxAutoMap | Add-Member -MemberType NoteProperty -Name Mailbox -Value $parentMailbox
                $output += $MailboxAutoMap
            }
        }
    }
    $output | Sort AutoMappingGranted | Export-CSV $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
}

function Get-MailboxFolderSizes
{
    write-host "Running Users Mailbox Folder Sizes"
    $csv = $Script:us + "-folders.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv    
    }   
    Get-MailboxFolderStatistics $Script:user | Select Name, Foldersize, ItemsinFolder | Export-Csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
}

function Get-MailboxFolderPermissions
{
    write-host "Running All Folders a User has Permissions to"
    $csv = $us + "-folderperms.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv    
    }   
    [string]$permitteduser = $Script:user
    $output = @()
    #Pulls all Users in Domain
    $domainMailbox = get-mailbox -organizationalunit $domain -resultsize Unlimited
    #Creates CSV with Titles
    #Creates loop for all users in a domain
    foreach ($mailbox in $domainMailbox)
    {
        #Switches Windows Email Address Object to String
        [string]$account = $mailbox.WindowsEmailAddress
        #Pulls all folders for each user
        $folders = Get-MailboxFolderStatistics -Identity $account
        #Creates Loop for Each Folder
        foreach($folder in $folders)
        {
            $folderPath = $folder.FolderPath
            $perm = get-mailboxfolderpermission ("$account`:$folderPath").ToString().Replace('/','\') -user $permitteduser -erroraction silentlycontinue
            #If User has perms to a folder Exports to the CSV
            if($perm)
            {
                $permission = New-Object -TypeName PSObject
                $permission | Add-Member -MemberType NoteProperty -Name Mailbox -Value $account
                $permission | Add-Member -MemberType NoteProperty -Name Folder -Value $folderPath
                $permission | Add-Member -MemberType NoteProperty -Name PermissionGranted -Value $perm.accessrights
                $permission | Add-Member -MemberType NoteProperty -Name User -Value $permitteduser
                $output += $permission
            }
        }
    }
    $output | export-csv $Script:mydesk\exports\$Script:Domain\$csv -NoTypeInformation
}

function Get-MailboxDLMembership
{
    write-host "Running All Distribution Lists a User is a member of "
    $csv = $Script:us + "-dl.csv"
    If (Test-Path $Script:mydesk\exports\$Script:Domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:Domain\$csv    
    }   
    $dist = get-distributiongroup -organizationalunit $Script:Domain
    $find = $dist | where { (Get-DistributionGroupMember $_ | foreach {$_.PrimarySmtpAddress}) -eq $Script:user } | select PrimarySmtpAddress,DisplayName
    if ($find -ne $null) 
    {
        $find | export-csv $Script:mydesk\exports\$Script:Domain\$csv -NoTypeInformation                      
    }
    else
    {
        Write-Host -ForegroundColor DarkRed "User is not on any DLs"                 
    }
}

function Get-DomainVerification
{
    $Script:Domain = Read-Host ("Please Enter the Domain you are working on")
    Write-Host -NoNewline "Verifying $($Script:Domain)....."
    $test = Get-organizationalunit $Script:Domain -ErrorAction SilentlyContinue
    if ($test -ne $null)
    {
        Write-Host -ForegroundColor green  "Verification Passed"
        Start-Sleep -seconds 2
        If (Test-Path $Script:myDesk\exports\$Script:Domain)
        {
            Remove-Item $Script:myDesk\exports\$Script:Domain -recurse -force   
        }
        New-Item -ItemType directory -Path $Script:myDesk\exports\$Script:Domain -erroraction silentlycontinue
    }
    else
      {
        Write-Host -ForegroundColor Red "Verification Failed"
        Write-Host "Please verify spelling of the domain"
        Write-Host "Restarting Script."
        Start-Sleep -Seconds 5
        Get-MainMenu
      }  
}

function Get-AllDomainReports
{
    Get-AllDistributionGroupMembers
    Get-AllX500Addresses
    Get-DomainGALObjects
    Get-ResourceMailboxSizes
    Get-DomainMailboxSize
    Get-CASMailboxFeatures
    Get-DomainPermissions
    Get-DomainAutoMapping
}

function Get-AllX500Addresses
{
    write-host "Running Pull Legacy Dn/X500 Addresses"
    $csv = "x500.csv"
    $csv1 = "x500dl.csv"
    $csv2 = "x500contacts.csv"
    If (Test-Path $mydesk\exports\$Script:Domain\$csv)
    {
        Remove-Item $mydesk\exports\$Script:Domain\$csv    
    }   
    If (Test-Path $mydesk\exports\$Script:Domain\$csv1)
    {
        Remove-Item $mydesk\exports\$Script:Domain\$csv1   
    }  
    If (Test-Path $mydesk\exports\$Script:Domain\$csv2)
    {
        Remove-Item $mydesk\exports\$Script:Domain\$csv2   
    }  
    get-mailbox -OrganizationalUnit $Script:Domain -resultsize unlimited | select pri*, legacy* | export-csv $mydesk\exports\$Script:Domain\$csv -NoTypeInformation
    get-distributiongroup -OrganizationalUnit $Script:Domain -resultsize unlimited | select pri*, legacy* | export-csv $mydesk\exports\$Script:Domain\$csv1 -NoTypeInformation
    get-mailcontact -OrganizationalUnit $Script:Domain -resultsize unlimited | select pri*, legacy* | export-csv $mydesk\exports\$Script:Domain\$csv2 -NoTypeInformation
}

function Get-AllDistributionGroupMembers
{
    write-host "Running All Distribution Groups members"
    $csv = "DL-Members.csv"
    $output = @()
    If (Test-Path $mydesk\exports\$Domain\$csv)
    {
        Remove-Item $mydesk\exports\$domain\$csv    
    }    
    $DLList = get-distributiongroup -OrganizationalUnit $Script:Domain -resultsize unlimited
    $DistributionGroups = New-Object -TypeName PSObject
    foreach ($DL in $DLList) 
    {
        # Get the member list
        $distributionGroup = $dl.WindowsEmailAddress
        $distributionGroupMembers = get-distributiongroupmember $dl -ResultSize unlimited
        foreach ($user in $distributionGroupMembers)
        {
            $member = New-Object -TypeName PSObject
            $member | Add-Member -MemberType NoteProperty -Name DistributionGroup -Value $dl.windowsEmailAddress
            $member | Add-Member -MemberType NoteProperty -Name MemberDisplayName -Value $user.DisplayName
            $member | Add-Member -MemberType NoteProperty -Name MemberSMTPAdress -Value $user.PrimarySMTPAddress
            $output += $member
        }

    }
    $output | Export-Csv $Script:mydesk\exports\$Script:Domain\$csv -NoTypeInformation 
}

function Get-DomainGALObjects
{
    write-host "Running All GAL Entries"
    $csv = "GAL.csv"
    $objects = @()
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv
    }
    $objects = get-user -organizationalunit $Script:domain -ResultSize Unlimited 
    $objects += Get-MailContact -organizationalunit $Script:domain -ResultSize Unlimited
    $objects | select FirstName, LastName, DisplayName, Initials, WindowsEmailAddress, company, Department, Manager, Title, Fax, AssistantName, Notes, Office, @{Name=’Mobile Phone’;Expression={[string]::join(";", ($_.MobilePhone))}}, @{Name=’Other Fax’;Expression={[string]::join(";", ($_.OtherFax))}}, @{Name=’Home Phone’;Expression={[string]::join(";", ($_.HomePhone))}}, @{Name=’Other Home Phone’;Expression={[string]::join(";", ($_.OtherHomePhone))}}, @{Name=’Other Telephone’;Expression={[string]::join(";", ($_.OtherTelephone))}}, @{Name=’Pager’;Expression={[string]::join(";", ($_.Pager))}}, @{Name=’Phone’;Expression={[string]::join(";", ($_.Phone))}}, StreetAddress, StateOrProvince, PostalCode, countryorregion,@{Name=’Post Office Box’;Expression={[string]::join(";", ($_.PostOfficeBox))}} | export-csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
}

function Get-ResourceMailboxSizes
{
    write-host "Running Resource Mailbox Sizes"
    $csv ="resource.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv 
    }
    $resource = Get-Mailbox -OrganizationalUnit $Script:domain -RecipientTypeDetails EquipmentMailbox -resultsize unlimited 
    if ($resource)
    {
       $resource | Get-MailboxStatistics | select displayname, totalitemsize | export-csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
    }
    else
    {
        Write-Host -ForegroundColor Red "No Resource Mailboxes Found."
    }
}

function Get-DomainMailboxSize
{
    write-host "Running All Mailbox Sizes"
    $csv = "users-sizes.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv
    }
    get-Mailbox -OrganizationalUnit $Script:domain -resultsize unlimited | Get-MailboxStatistics | select displayname,itemcount,totalitemsize | export-csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
}

function Get-CASMailboxFeatures
{
    write-host "Running all Domain Users Exchange Features"
    $csv = "users-features.csv"
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv
    }
    get-casmailbox -OrganizationalUnit $Script:domain -resultsize unlimited |  select PrimarySMTPAddress,ActivesyncEnabled,OWAEnabled,POPEnabled,IMAPEnabled,MAPIEnabled  | Export-Csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation  
}

function Get-DomainPermissions
{
    write-host "Running All Permissions for the Domain"
    $csv = "Domain-Permission.csv"
    $output = @()
    If (Test-Path $Script:mydesk\exports\$Script:domain\$csv)
    {
        Remove-Item $Script:mydesk\exports\$Script:domain\$csv
    }
    $Mailboxes = Get-Mailbox -organizationalunit $domain -RecipientType UserMailbox -ResultSize Unlimited 
    ForEach ($Mailbox in $Mailboxes)
    {
        $SendAs = Get-ADPermission $Mailbox.DistinguishedName | ? {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY\SELF" -and $_.User -notlike "MEX05\MEX05 Exchange Engineers" -and $_.User -notlike "MEX05\Migrations Team" -and $_.User -notlike "MEX05\Managed Mail L*"}
        foreach ($send in $sendas)
        {
            $permission = New-Object -TypeName PSObject
            $permission | Add-Member -MemberType NoteProperty -Name Mailbox -Value $mailbox.PrimarySMTPAddress
            $permission | Add-Member -MemberType NoteProperty -Name PermissionGranted -Value "SendAs"
            $permission | Add-Member -MemberType NoteProperty -Name User -Value $send.user
            $output += $permission
        }
        $FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and $_.User -notlike "NT AUTHORITY\SELF" -and $_.User -notlike "MEX05\MEX05 Exchange Engineers" -and $_.User -notlike "MEX05\Migrations Team" -and $_.User -notlike "MEX05\Managed Mail L*"}
        foreach ($user in $fullAccess)
        {
            $permission = New-Object -TypeName PSObject
            $permission | Add-Member -MemberType NoteProperty -Name Mailbox -Value $mailbox.PrimarySMTPAddress
            $permission | Add-Member -MemberType NoteProperty -Name PermissionGranted -Value "FullAccess"
            $permission | Add-Member -MemberType NoteProperty -Name User -Value $user.user
            $output += $permission
        }
    }
    $output | Export-Csv $Script:mydesk\exports\$Script:domain\$csv -NoTypeInformation
}

function Get-IISLogApp
{
    $iisQueryTarget = Read-Host "Please enter a domain, email address or a list of email addresses separated by commas"
    $iisQueryTarget = ($iisQueryTarget).Replace(" ", "")

    if($iisQueryTarget -like "*@*")
    {
        if($iisQueryTarget -like "*,*")
        {
            #Made this an ArrayList so I can remove invalid addresses
            [System.Collections.ArrayList]$mailboxList = $iisQueryTarget.Split(",")
            
            $mailboxList | ForEach-Object
            {
                Write-Host "Validating " + $_ + "... " -NoNewline
                $mailboxObj   = Get-Mailbox $_ -ErrorAction "SilentlyContinue"
                $iisQueryList = @()
                if(($mailboxObj | Measure-Object).Count -ne 1)
                {
                    Write-Host "Failed!" -ForegroundColor "Red"
                }
                else
                {
                    $iisQueryList += $mailboxObj
                    Write-Host "Success!" -ForegroundColor Green
                }
            }
        }
        else
        {
            $mailboxObj = Get-Mailbox $iisQueryTarget -ErrorAction "Stop"
            if (($mailboxObj | Measure-Object).Count -eq 1)
            {
                $iisQueryList = $mailboxObj 
            }
            else
            {
                Throw("Ambiguous Mailbox. Please verify that this address is not assigned to more than one object.")
            }
        }
    }
    else
    {
        Write-Host "Gathering mailboxes for domain" $iisQueryTarget "... " -NoNewline
        $iisQueryList = Get-Mailbox -ResultSize Unlimited -OrganizationalUnit $iisQueryTarget -ErrorAction "Stop"
        Write-Host "Success" -ForegroundColor Green
    }

    $iisQueryStartDate = Read-Host "Please enter a starting date (blank for today) mm/dd/yyyy"
    if (!$iisQueryStartDate)
    {
        $iisQueryStartDate = Get-Date
    }
    [int]$iisQueryDays = Read-Host "Number of days worth of logs needed (max 14)"
    $iisQueryResponse  = @() 
    $i = 1

    foreach ($mailbox in $iisQueryList)
    {
        Write-Progress -Activity ("Gathering IIS logs for " + $mailbox.PrimarySMTPAddress.ToString()) -Status ("Mailbox " + $i + " of " + ($iisQueryList | Measure-Object).Count) -PercentComplete (($i / ($iisQueryList | Measure-Object).Count) * 100) -Id 1
        $iisQueryResponse += Get-IISLogs -mailbox $mailbox.PrimarySMTPAddress -dateRangeStart $iisQueryStartDate -dateRangeDays $iisQueryDays
        $i++
    }
    if (($iisQueryList | Measure-Object).Count -gt 1)
    {
        $csvExportPath = $Script:mydesk+'\exports\'+$Script:Domain+'\'+$iisQueryList[0].PrimarySmtpAddress.ToString().Split('@')[1]+'-iis.csv'
    }
    else
    {
        $csvExportPath = $Script:mydesk+'\exports\'+$Script:Domain+'\'+$iisQueryList.PrimarySmtpAddress.ToString().Split('@')[1]+'-iis.csv'

    }
    $iisQueryResponse | Sort-Object -Property cs_username, timestamp | Select-Object * | Export-CSV $csvExportPath -NoTypeInformation
    #(Get-Content $csvExportPath) -notmatch '^\s*$' > $csvExportPath
}

function Get-IISLogs
{
    Param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true
    )]
    [string]
    $mailbox,
    $dateRangeStart = (Get-Date),
    [Parameter(
        Mandatory=$true
    )]
    [int]
    $dateRangeDays
    )
    BEGIN
    {
        Write-Verbose ("Starting" + $($MyInvocation).MyCommand)
        Write-Verbose ("Importing necessary modules")
        Import-Module \\ord2stts01.mex05.mlsrvr.com\C$\scripts\Utility_Server_Tool\modules\ElasticAPIModule2010.psm1 -ErrorAction "Stop"

        $iisLogResults = @()
    }
    PROCESS
    {
        $dateRange      = Get-DateRange -StartDate $dateRangeStart -NumberOfDays $dateRangeDays
        $mailboxObj     = Get-Mailbox $mailbox
        $i              = 1;
        foreach ($date in $dateRange)
        {
            $elasticIndex = "cas_mex05-" + $date.Date
            $iisLogQuery  = Get-ElasticData -ElasticUri "http://kb6.deepfield.io" -Index $elasticIndex -SearchKey "cs_username_na" -SearchTerm $mailboxObj.PrimarySMTPAddress.ToString()
            $iisLogQuery  += Get-ElasticData -ElasticUri "http://kb6.deepfield.io" -Index $elasticIndex -SearchKey "cs_username_na" -SearchTerm $mailboxObj.SamAccountName
            Write-Progress -Activity ("Searching date " + (Get-Date $date.Date -UFormat "%m/%d/%Y")) -Status ("Query " + $i + " of " + ($dateRange | Measure-Object).Count) -PercentComplete (($i / ($dateRange | Measure-Object).Count) * 100) -ParentId 1
            $i++

            if (($iisLogQuery | Measure-Object).Count -gt 0)
            {
                foreach($result in $iisLogQuery)
                {
                    $iisLogObj = New-Object -TypeName PSObject
                    $iisLogObj | Add-Member -MemberType NoteProperty -Name "timestamp" -Value $result.datetime
                    $iisLogObj | Add-Member -MemberType NoteProperty -Name "host" -Value $result.c_ip
                    $iisLogObj | Add-Member -MemberType NoteProperty -Name "cs_username" -Value $result.cs_username
                    $iisLogObj | Add-Member -MemberType NoteProperty -Name "cs_useragent" -Value $result.cs_useragent
                    $iisLogObj | Add-Member -MemberType NoteProperty -Name "cs_uri_stem" -Value $result.cs_uri_stem
                    $iisLogObj | Add-Member -MemberType NoteProperty -Name "sc_status" -Value $result.sc_status
                    $iisLogObj | Add-Member -MemberType NoteProperty -Name "sc_substatus" -Value $result.sc_substatus
                    $iisLogObj | Add-Member -MemberType NoteProperty -Name "environment" -Value $result.environment

                    $iisLogResults += $iisLogObj
                }
            }

            $iisLogQuery = $null
        }
    }
    END
    {
        Write-Verbose ("Removing Powershell modules")
        Remove-Module ElasticAPIModule -ErrorAction "SilentlyContinue"
        
        Write-Output ($iisLogResults)
        Write-Verbose ("Ending " + $($MyInvokation).MyCommand)
    }
}

#endregion


function Set-ImpersonationRights
{

<#Check to see if Exchange snapin is loaded.  If not, load it #>
If ( (Get-PSSnapin -Name *exchange* -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin *exchange*
}

CLS

<#Get variables from user#>
#Write-Host "What is the customer domain name that would like to configure application impersonation? " -ForegroundColor Yellow -NoNewline
#$strDomainA = Read-Host
Write-Host "What is the email address of the service account that will be doing the impersonating? " -ForegroundColor Yellow -NoNewline
$strSvc_Acct = Read-Host
$strdomain = $strSvc_Acct -split '@'
$strDomainA = $strDomain[1]
Write-Host ""

<#Set up RecipientRestrictionFilter#>
$filter = "CustomAttribute10 -eq '$Script:Domain'"

<#Set up ManagementScope and ManagementRoleAssignment names#>
$scopeName = "AppImpers-$Script:Domain"
$roleAssignmentName = "AppImpers-$Script:Domain"

<#Define the ManagmentScope#>
New-ManagementScope -Name $scopeName -RecipientRestrictionFilter  $filter -WarningAction SilentlyContinue | Out-Null

<#Define the ManagementRoleAssignment#>
New-ManagementRoleAssignment –Name $roleAssignmentName  -Role ApplicationImpersonation –User $strSvc_Acct -CustomRecipientWriteScope $scopeName -WarningAction SilentlyContinue | Out-Null

<#Output results to screen#>
Write-Host "New Management Scope" -ForegroundColor Cyan
Get-ManagementScope | Where {$_.name -like "*$Script:Domain*"} | select Name,RecipientFilter | fl

Write-Host	"New Management Role Assignment" -ForegroundColor Cyan
Get-ManagementRoleAssignment | Where {$_.name -like "*$Script:Domain*"} | select Name,CustomRecipientWriteScope,RoleAssignee | fl

Write-Host "`n`n"

Write-Host "*** There is no error handling in this script, so if the above output is blank, try running these commands individually and see if you get errors *** `n" -BackgroundColor Black -ForegroundColor Blue
Write-Host "New-ManagementScope -Name $scopeName -RecipientRestrictionFilter  $filter `n" -BackgroundColor Black -ForegroundColor White
Write-Host "New-ManagementRoleAssignment –Name $roleAssignmentName  -Role ApplicationImpersonation –User $strSvc_Acct -CustomRecipientWriteScope $scopeName `n" -BackgroundColor Black -ForegroundColor White

<#Grant Full Access Permissions without Automapping#>
$mbxs_domain = Get-mailbox -organizationalunit $Script:Domain
$sam_account = (Get-mailbox $strSvc_Acct).samaccountname

Write-Host	"Grant Full Access Permissions" -ForegroundColor Cyan
$mbxs_domain | add-mailboxpermission -user $sam_account -AccessRights FullAccess -Automapping $false
}

#region Script Variables
$Script:environment = (Get-WmiObject Win32_ComputerSystem).Domain
$Script:mydesk = [environment]::getfolderpath("desktop")
#endregion

#region Main
Add-PSSnapin *Exchange*
Clean-ExportFolder
Get-MainMenu
Send-EmailReport
Exit
#endregion