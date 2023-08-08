$appName = "MigrationWiz-ModernAuth-EWS_Aaron"

function New-MgGraphMigrationWizModernAuthAppRegistration {
    param (
        [Parameter(Mandatory = $true)]
        [string]$appName,
        [Parameter(Mandatory = $false)]
        [string]$migTenant
    )

    # Get the Tenant Detail - Azure AD
    $tenantDetail = Get-MgOrganization
    $tenantID = $tenantDetail.Id

    # Define Source or Destination Tenant
    if ($migTenant -eq $null){
        $migTenant = Read-Host -Prompt "For the Migration Wiz Project, is this the Source or Destination tenant? (Source/Destination)"
    }
    if ($migTenant -eq "Source"){
        $AdvancedOptionsTenant = "Export"
    }elseif ($migTenant -eq "Destination"){
        $AdvancedOptionsTenant = "Import"
    }else{
        Write-Host "Invalid Entry. Please enter Source or Destination"
        $migTenant = Read-Host -Prompt "For the Migration Wiz Project, is this the Source or Destination tenant? (Source/Destination)"
    }


    # Create the app - MGGraph
    Write-Host "Creating the APP and Service Principal" -BackgroundColor Yellow -ForegroundColor Black
    $newRedirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $publicClientApplication = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphPublicClientApplication
    $publicClientApplication.RedirectUris = $newRedirectUri
    $homePageURL = "https://help.bittitan.com/hc/en-us/articles/360034124813-Authentication-Methods-for-Microsoft-365-All-Products-Migrations#enabling-modern-authentication-for-ews-between-migrationwiz-and-your-exchange-online-tenant-0-10"

    # Check if the app exists; If the app doesn't exist, create it
    if (!($app = Get-MgApplication -Filter "displayName eq '$appName'")) {
        $app = New-MgApplication -DisplayName $appName -SignInAudience "AzureADMultipleOrgs" -IsFallbackPublicClient -PublicClient $publicClientApplication -Web @{HomePageURL = $homePageURL} 
        $appClientID = $app.AppId
        Write-Output "Created $appName new application with AppId: $($app.AppId)"
    }
    else {
        # If app exists, add new redirect URI
        $redirectUris = $app.Web.RedirectUris
        if ($null -eq $redirectUris) { $redirectUris = @() }
        $redirectUris += $newRedirectUri
        Write-Output "Application already exists. AppId: $($app.AppId)"
    }

    # Add delay if necessary
    #Set-MgApplication -ApplicationId $app.AppId -AllowPublicClient $true
    Write-Output "Waiting for $appName to be created..."
    Start-Sleep -Seconds 15

    # Create a service principal for the app - MGGraph
    $servicePrincipal = New-MgServicePrincipal -AppId $app.AppId

    # Output the Service Principal Id
    Write-Host "Grant the Oauth Permissions - EWS for Office 365 Exchange Online" -BackgroundColor Yellow -ForegroundColor Black

    $servicePrincipalClientID = (Get-MgServicePrincipal -Filter "displayName eq '$($appName)'").ID
    Write-Output ""
    # Output the ClientId and ResourceId
    #Write-Output "App ClientId: $($appClientID)"
    Write-Output "ServicePrincipal-Client Id: $($servicePrincipalClientID)"
    #resource ID is specific to the resource (API) that you want to access. 
    #In this case, Office 365 Exchange Online is the resource. The ResourceID for Office 365 Exchange Online is d9e49bfe-e0b2-4070-a0f6-d919f6a31355

    $exchangeOnline = Get-MgServicePrincipal -Filter "DisplayName eq 'Office 365 Exchange Online'"
    $office365ExchangeOnlineResourceId = $exchangeOnline.Id
    #$office365ExchangeOnlineResourceId = "d9e49bfe-e0b2-4070-a0f6-d919f6a31355"
    Write-Output "Office 365 Exchange Online ResourceId: $($office365ExchangeOnlineResourceId)"

    ## create expiry time
        #not sure why an expiry is needed. Microsoft documentation doesn't include this step
        # https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.identity.signins/new-mgoauth2permissiongrant?view=graph-powershell-1.0

    # Get the current date
    $currentDate = Get-Date
    # Add 12 months to the current date
    $newDate = $currentDate.AddMonths(12)

    #Add Desired Permissions:
    $params = @{
        "ClientId" = $servicePrincipalClientID
        "ConsentType" = "AllPrincipals"
        "ResourceId" = $office365ExchangeOnlineResourceId
        "Scope" = "EWS.AccessAsUser.All"
        "expiryTime" = $newDate
    }
    # Grant the permissions
    $newOauthPerms = New-MgOauth2PermissionGrant -BodyParameter $params #| Format-List Id, ClientId, ConsentType, ResourceId, Scope

    #Confirm Permissions Granted
    Write-Output ""
    Write-Host "Verify Permissions Granted" -BackgroundColor Yellow -ForegroundColor Black
    Get-MgOauth2PermissionGrant -Filter "clientId eq '$servicePrincipalClientID' and consentType eq 'AllPrincipals'" | fl


    # Output the AppClientId and TenantId
    #Write-Output ""
    Write-Host "$($migTenant) Details Needed for MigrationWiz" -BackgroundColor Yellow -ForegroundColor Black
    Write-Output "$($tenantDetail.DisplayName) Application (Client) ID: ModernAuthClientId$($AdvancedOptionsTenant)=$($appClientID)"
    Write-Output "$($tenantDetail.DisplayName) Tenant Id: ModernAuthClientId$($AdvancedOptionsTenant)=$($tenantID)"
    Write-Output "$($newdate) is the expiry date for the permissions"

    Write-Output ""
    
}

#Grant

New-MgGraphMigrationWizModernAuthAppRegistration -appName $appName -migTenant Destination



function New-AzureADMigrationWizModernAuthAppRegistration {
    param (
        [Parameter(Mandatory = $true)]
        [string]$appName
    )
    # Import the required module
    Import-Module AzureAD

    # Sign in to Azure AD with Global Administrator credentials
    Connect-AzureAD

    # Get the Tenant Detail
    $tenantDetail = Get-AzureADTenantDetail

    # Define the new application
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"

    # Create the app - Azure AD
    $app = New-AzureADApplication -DisplayName $appName -ReplyUrls $redirectUri -PublicClient $true

    # Get the app - Azure AD
    $app = Get-AzureADApplication -Filter "displayName eq '$($appName)'"
    # Set the public client to true (enables public client flows) - Azure AD
    Set-AzureADApplication -ObjectId $app.ObjectId -PublicClient $true

    # Create a service principal for the app - Azure AD
    $servicePrincipal = New-AzureADServicePrincipal -AppId $newApp.AppId

    # Output the Service Principal Id
    Write-Output "ServicePrincipalId: $($servicePrincipal.Id)"

    #Get the App Client and Tenant Ids
    $appClientID = $newApp.AppId
    $tenantID = $tenantDetail.Id

}