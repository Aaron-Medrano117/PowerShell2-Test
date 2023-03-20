#Reference Articlest to Update Teams Meeting Invites Post Migration with Graph API
# https://alexholmeset.blog/2018/10/10/getting-started-with-graph-api-and-powershell/
# https://alexholmeset.blog/2021/06/16/teams-meeting-tenant-to-tenant-migration/

#check if theres an attendee from the previous tenant.
$OldDomain = Read-Host "What is the Migrating Source/Old Domain?"
$NewDoamin = Read-Host "What is the Migrating Destination/New Domain?"

#Example
#$OldDomain = 'sourcedomainexample.com'
#$NewDoamin = 'destinationdomainexample.com'
 
#Cancelation message
$CancelationMessage = "We are moving over to a new system, so this meeting will be canceled. You will receive new invite from our new domain."
 
#From line 150, you find where to update the Client ID, Tenant ID and App secret.
 
function GetStringBetweenTwoStrings($text){
 
    #Regex pattern to compare two strings
    $pattern = "(?s)(?<=________________________________________________________________________________)(.*?)(?=________________________________________________________________________________)"
 
    #Perform the opperation
    $result = [regex]::Match($text,$pattern).value
 
    #Return result
    return $result
 
}
 
 
function Html-ToText {
    param([System.String] $html)
    
    # remove line breaks, replace with spaces
    $html = $html -replace "(`r|`n|`t)", " "
    # write-verbose "removed line breaks: `n`n$html`n"
    
    # remove invisible content
    @('head', 'style', 'script', 'object', 'embed', 'applet', 'noframes', 'noscript', 'noembed') | % {
     $html = $html -replace "<$_[^>]*?>.*?</$_>", ""
    }
    # write-verbose "removed invisible blocks: `n`n$html`n"
    
    # Condense extra whitespace
    $html = $html -replace "( )+", " "
    # write-verbose "condensed whitespace: `n`n$html`n"
    
    # Add line breaks
    @('div','p','blockquote','h[1-9]') | % { $html = $html -replace "</?$_[^>]*?>.*?</$_>", ("`n" + '$0' )} 
    # Add line breaks for self-closing tags
    @('div','p','blockquote','h[1-9]','br') | % { $html = $html -replace "<$_[^>]*?/>", ('$0' + "`n")} 
    # write-verbose "added line breaks: `n`n$html`n"
    
    #strip tags 
    $html = $html -replace "<[^>]*?>", ""
    # write-verbose "removed tags: `n`n$html`n"
      
    # replace common entities
    @( 
     @("&amp;bull;", " * "),
     @("&amp;lsaquo;", "<"),
     @("&amp;rsaquo;", ">"),
     @("&amp;(rsquo|lsquo);", "'"),
     @("&amp;(quot|ldquo|rdquo);", '"'),
     @("&amp;trade;", "(tm)"),
     @("&amp;frasl;", "/"),
     @("&amp;(quot|#34|#034|#x22);", '"'),
     @('&amp;(amp|#38|#038|#x26);', "&amp;"),
     @("&amp;(lt|#60|#060|#x3c);", "<"),
     @("&amp;(gt|#62|#062|#x3e);", ">"),
     @('&amp;(copy|#169);', "(c)"),
     @("&amp;(reg|#174);", "(r)"),
     @("&amp;nbsp;", " "),
     @("&amp;(.{2,6});", "")
    ) | % { $html = $html -replace $_[0], $_[1] }
    # write-verbose "replaced entities: `n`n$html`n"
    
    return $html
    
   }
 
function Get-MSGraphAppToken{
    <#  .SYNOPSIS
        Get an app based authentication token required for interacting with Microsoft Graph API
    .PARAMETER TenantID
        A tenant ID should be provided.
   
    .PARAMETER ClientID
        Application ID for an Azure AD application. Uses by default the Microsoft Intune PowerShell application ID.
   
    .PARAMETER ClientSecret
        Web application client secret.
          
    .EXAMPLE
        # Manually specify username and password to acquire an authentication token:
        Get-MSGraphAppToken -TenantID $TenantID -ClientID $ClientID -ClientSecert = $ClientSecret 
    .NOTES
        Author: Jan Ketil Skanke
        Contact: @JankeSkanke
        Created: 2020-15-03
        Updated: 2020-15-03
   
        Version history:
        1.0.0 - (2020-03-15) Function created      
    #>
[CmdletBinding()]
    param (
        [parameter(Mandatory = $true, HelpMessage = "Your Azure AD Directory ID should be provided")]
        [ValidateNotNullOrEmpty()]
        [string]$TenantID,
        [parameter(Mandatory = $true, HelpMessage = "Application ID for an Azure AD application")]
        [ValidateNotNullOrEmpty()]
        [string]$ClientID,
        [parameter(Mandatory = $true, HelpMessage = "Azure AD Application Client Secret.")]
        [ValidateNotNullOrEmpty()]
        [string]$ClientSecret
        )
Process {
    $ErrorActionPreference = "Stop"
         
    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
        }
      
    try {
        $MyTokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
        $MyToken =($MyTokenRequest.Content | ConvertFrom-Json).access_token
            If(!$MyToken){
                Write-Warning "Failed to get Graph API access token!"
                Exit 1
            }
        $MyHeader = @{"Authorization" = "Bearer $MyToken" }
       }
    catch [System.Exception] {
        Write-Warning "Failed to get Access Token, Error message: $($_.Exception.Message)"; break
    }
    return $MyHeader
    }
}

$OldTenantId = Read-Host "What is the Old Tenant ID"
$OldClientID = Read-Host "What is the Old Tenant's Client ID for the App"
$OldClientSecret = Read-Host "What is the Old Tenant's Client Secret for the App"
#$OldTenantId = 'xxxxxxxxxxxxx'
#$OldClientID = 'xxxxxxxxxxxxx'
#$OldClientSecret = "xxxxxxxxxxxxx"
$global:OldHeader = Get-MSGraphAppToken -TenantID $OldTenantId -ClientID $OldClientID -ClientSecret $OldClientSecret

$NewTenantId = Read-Host "What is the New Tenant ID"
$NewClientID = Read-Host "What is the New Tenant's Client ID for the App"
$NewClientSecret = Read-Host "What is the New Tenant's Client Secret for the App"

#$NewTenantId = 'xxxxxxxxxxxxx'
#$NewClientID = 'xxxxxxxxxxxxx'
#$NewClientSecret = "xxxxxxxxxxxxx"
$global:NewHeader = Get-MSGraphAppToken -TenantID $NewTenantId -ClientID $NewClientID -ClientSecret $NewClientSecret
 
#Gets all internal users in the old tenant.
 
$currentUri = "https://graph.microsoft.com/beta/users?`$filter=userType eq 'Member'"
 
$UsersOldTenant = while (-not [string]::IsNullOrEmpty($currentUri)) {
 
    # API Call
    Write-Host "`r`nQuerying $currentUri..." -ForegroundColor Yellow
    $apiCall = Invoke-WebRequest -Method "GET" -Uri $currentUri -ContentType "application/json" -Headers $global:OldHeader -ErrorAction Stop
     
    $nextLink = $null
    $currentUri = $null
 
    if ($apiCall.Content) {
 
        # Check if any data is left
        $nextLink = $apiCall.Content | ConvertFrom-Json | Select-Object '@odata.nextLink'
        $currentUri = $nextLink.'@odata.nextLink'
 
        $apiCall.Content | ConvertFrom-Json
 
    }
 
}
 
 
#Gets all internal users in the new tenant.
 
$currentUri = "https://graph.microsoft.com/beta/users?`$filter=userType eq 'Member'"
 
$UsersNewTenant = while (-not [string]::IsNullOrEmpty($currentUri)) {
 
    # API Call
    Write-Host "`r`nQuerying $currentUri..." -ForegroundColor Yellow
    $apiCall = Invoke-WebRequest -Method "GET" -Uri $currentUri -ContentType "application/json" -Headers $global:NewHeader -ErrorAction Stop
     
    $nextLink = $null
    $currentUri = $null
 
    if ($apiCall.Content) {
 
        # Check if any data is left
        $nextLink = $apiCall.Content | ConvertFrom-Json | Select-Object '@odata.nextLink'
        $currentUri = $nextLink.'@odata.nextLink'
 
        $apiCall.Content | ConvertFrom-Json
 
    }
 
}
 
 
foreach($UserNewTenant in $UsersNewTenant.value){
 
    $UserNewTenantUPN = $UserNewTenant.userprincipalname
     #Gets all events for the current user in the new tenant.
 
$currentUri = "https://graph.microsoft.com/beta/users/$UserNewTenantUPN/events"
 
$NewTenantTeamsMeetingsBulk = while (-not [string]::IsNullOrEmpty($currentUri)) {
 
    # API Call
    Write-Host "`r`nQuerying $currentUri..." -ForegroundColor Yellow
    $apiCall = Invoke-WebRequest -Method "GET" -Uri $currentUri -ContentType "application/json" -Headers $global:NewHeader -ErrorAction Stop
     
    $nextLink = $null
    $currentUri = $null
 
    if ($apiCall.Content) {
 
        # Check if any data is left
        $nextLink = $apiCall.Content | ConvertFrom-Json | Select-Object '@odata.nextLink'
        $currentUri = $nextLink.'@odata.nextLink'
 
        $apiCall.Content | ConvertFrom-Json
 
    }
 
}
$NewTenantTeamsMeetings =  $NewTenantTeamsMeetingsBulk.value | Where-Object{(get-date $($_.start).datetime -Format yyyy-MM-ddTHH:MM) -ge (get-date -Format yyyy-MM-ddTHH:MM)}
$NewTenantTeamsMeetingsSeriesPastStartDate = $NewTenantTeamsMeetingsBulk.value | Where-Object{(get-date $($_.start).datetime -Format yyyy-MM-ddTHH:MM) -lt (get-date -Format yyyy-MM-ddTHH:MM)} | Where-Object{$_.type -like "seriesMaster"}
 
    foreach($NewTenantTeamsMeeting in $NewTenantTeamsMeetings){
       
        $UserNewTenantUPNPrefix = $UserNewTenantUPN.Split('@')[0]
        
        If((($NewTenantTeamsMeeting.organizer).emailaddress).address.StartsWith($UserNewTenantUPNPrefix)){
            if($NewTenantTeamsMeeting.isOnlineMeeting -eq $true){
 
                $InviteText = Html-ToText -html ($NewTenantTeamsMeeting.body).content 
                $InviteTextToRemove = GetStringBetweenTwoStrings -text  $InviteText
                $InviteText = $InviteText.replace($InviteTextToRemove,'')
                $InviteText
 
            
 
            $inivteBody = @"
            {
                "subject": "$($NewTenantTeamsMeeting.subject)",
                "body": {
                  "contentType": "HTML",
                  "content": "$InviteText"
                },
                "start": $($NewTenantTeamsMeeting.start | ConvertTo-Json),
                "end": $($NewTenantTeamsMeeting.end | ConvertTo-Json),
                "recurrence":$($NewTenantTeamsMeeting.recurrence | ConvertTo-Json),
                "location":{
                    "displayName":"$(($NewTenantTeamsMeeting.location).displayname)"
                },
                "attendees": [
                  {
                    "emailAddress": $((($NewTenantTeamsMeeting.attendees).emailaddress | ConvertTo-Json).replace($OldDomain,$NewDoamin)),
                    "type": "required"
                  }
                ],
                "isOnlineMeeting": true,
                "onlineMeetingProvider": "teamsForBusiness"
              }
"@
             $inivteBody
 
            $NeweventURI = "https://graph.microsoft.com/v1.0/users/$UserNewTenantUPN/calendar/events/"
            $NeweventURI 
            Invoke-WebRequest -Method "POST" -Uri $NeweventURI -ContentType "application/json" -Headers $global:NewHeader -Body $inivteBody 
 
            }
             
             
        }
 
 
 
    }
 
    foreach($NewTenantTeamsMeetingSeriesPastStartDate in  $NewTenantTeamsMeetingsSeriesPastStartDate ){
 
    If((get-date (($NewTenantTeamsMeetingSeriesPastStartDate.recurrence).range).endDate -Format yyyy-MM-ddTHH:MM) -ge (Get-Date -Format yyyy-MM-ddTHH:MM)){
     
           
        $UserNewTenantUPNPrefix = $UserNewTenantUPN.Split('@')[0]
        
        If((($NewTenantTeamsMeetingSeriesPastStartDate.organizer).emailaddress).address.StartsWith($UserNewTenantUPNPrefix)){
            if($NewTenantTeamsMeetingSeriesPastStartDate.isOnlineMeeting -eq $true){
 
                $InviteText = Html-ToText -html ($NewTenantTeamsMeetingSeriesPastStartDate.body).content 
                $InviteTextToRemove = GetStringBetweenTwoStrings -text  $InviteText
                $InviteText = $InviteText.replace($InviteTextToRemove,'')
                $InviteText
 
            
 
            $inivteBody = @"
            {
                "subject": "$($NewTenantTeamsMeetingSeriesPastStartDate.subject)",
                "body": {
                  "contentType": "HTML",
                  "content": "$InviteText"
                },
                "start": $($NewTenantTeamsMeetingSeriesPastStartDate.start | ConvertTo-Json),
                "end": $($NewTenantTeamsMeetingSeriesPastStartDate.end | ConvertTo-Json),
                "recurrence":$($NewTenantTeamsMeetingSeriesPastStartDate.recurrence | ConvertTo-Json),
                "location":{
                    "displayName":"$(($NewTenantTeamsMeetingSeriesPastStartDate.location).displayname)"
                },
                "attendees": [
                  {
                    "emailAddress": $((($NewTenantTeamsMeetingSeriesPastStartDate.attendees).emailaddress | ConvertTo-Json).replace($OldDomain,$NewDoamin)),
                    "type": "required"
                  }
                ],
                "isOnlineMeeting": true,
                "onlineMeetingProvider": "teamsForBusiness"
              }
"@
             $inivteBody
 
            $NeweventURI = "https://graph.microsoft.com/v1.0/users/$UserNewTenantUPN/calendar/events/"
            $NeweventURI 
            Invoke-WebRequest -Method "POST" -Uri $NeweventURI -ContentType "application/json" -Headers $global:NewHeader -Body $inivteBody 
 
 
 
            }
             
             
        }
 
 
 
         
     
    }
     
    "there are this many meetings to delete"
    $($NewTenantTeamsMeetingsBulk.value).count
    foreach($NewTenantTeamsMeetingBulk in $NewTenantTeamsMeetingsBulk.value){
 
    $UserOldTenantupn = $UserNewTenantUPN.Replace($NewDoamin,$OldDomain)
    If((($NewTenantTeamsMeetingBulk.organizer).emailaddress).address.StartsWith($UserOldTenantupn)){
     
                $eventURI = "https://graph.microsoft.com/v1.0/users/$UserNewTenantUPN/events/$($NewTenantTeamsMeetingBulk.id)"
            "Deleted"
            $eventURI 
            $test = Invoke-WebRequest -Method "DELETE" -Uri $eventURI -ContentType "application/json" -Headers $global:NewHeader -ErrorAction Ignore
    
    }
    }
 
 
 
 
}
}
 
foreach($UserOldTenant in $UsersOldTenant.value){
 
    $UserOldTenantUPN = $UserOldTenant.userprincipalname
    $UserOldTenantUPNPrefix = $UserOldTenantUPN.Split('@')[0]
#Gets all events for the current user in the old tenant.
 
$currentUri = "https://graph.microsoft.com/beta/users/$UserOldTenantUPN/events"
 
$OldTenantMeetings = while (-not [string]::IsNullOrEmpty($currentUri)) {
 
    # API Call
    Write-Host "`r`nQuerying $currentUri..." -ForegroundColor Yellow
    $apiCall = Invoke-WebRequest -Method "GET" -Uri $currentUri -ContentType "application/json" -Headers $global:OldHeader -ErrorAction Stop
     
    $nextLink = $null
    $currentUri = $null
 
    if ($apiCall.Content) {
 
        # Check if any data is left
        $nextLink = $apiCall.Content | ConvertFrom-Json | Select-Object '@odata.nextLink'
        $currentUri = $nextLink.'@odata.nextLink'
 
        $apiCall.Content | ConvertFrom-Json
 
    }
 
}
 
     
            $OldTenantMeetingsNotSeries = $OldTenantMeetings.value | where-object{$_.type -notlike "seriesMaster"}
            "old tenant not meetings series"
            $OldTenantMeetingsNotSeries.count 
    foreach($OldTenantMeetingNotSeries in $OldTenantMeetingsNotSeries){
        If((Get-Date ($OldTenantMeetingNotSeries.end).datetime -Format yyyy-MM-dd) -ge (Get-Date -Format yyyy-MM-dd)){
        if((($OldTenantMeetingNotSeries.organizer).emailaddress).address.StartsWith($UserOldTenantUPNPrefix)){
        $OldTenantMeetingID 
        $OldTenantMeetingID = $OldTenantMeetingNotSeries.ID
 
        $CancelinivteBody = @"
        {
            "Comment": "$CancelationMessage"
          }
 
"@
 
$CancelinivteBody
 
        $CancelEventURI = "https://graph.microsoft.com/v1.0/users/$UserOldTenantUPN/events/$OldTenantMeetingID/cancel"
        $CancelEventURI
        Invoke-WebRequest -Method "POST" -Uri $CancelEventURI -ContentType "application/json" -Headers $global:OldHeader -Body $CancelinivteBody 
 
 
}
}
 
 
    }
             
            $OldTenantMeetingsSeries = $OldTenantMeetings.value | where-object{$_.type -like "seriesMaster"}
            "old tenant meetings series"
            $OldTenantMeetingsSeries.count
            foreach($OldTenantMeetingSeries in $OldTenantMeetingsSeries){
            If((get-date (($OldTenantMeetingSeries.recurrence).range).endDate -Format yyyy-MM-ddTHH:MM) -gt (Get-Date -Format yyyy-MM-ddTHH:MM)){
        if((($OldTenantMeetingSeries.organizer).emailaddress).address.StartsWith($UserOldTenantUPNPrefix)){
        $OldTenantMeetingID 
 
        $OldTenantMeetingID = $OldTenantMeetingSeries.ID
 
        $CancelinivteBody = @"
        {
            "Comment": "$CancelationMessage"
          }
 
"@
$CancelinivteBody
 
 
        $CancelEventURI = "https://graph.microsoft.com/v1.0/users/$UserOldTenantUPN/events/$OldTenantMeetingID/cancel"
        $CancelEventURI
        Invoke-WebRequest -Method "POST" -Uri $CancelEventURI -ContentType "application/json" -Headers $global:OldHeader -Body $CancelinivteBody
 
        }
}
 
 
    }
 
    }