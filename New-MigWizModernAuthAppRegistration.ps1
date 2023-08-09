<#
.synopsis
    Creates a new Modern Authentication Application for Migration Wiz
.DESCRIPTION
    Creates a new Modern Authentication Application for Migration Wiz
.EXAMPLE
    New-MgGraphMigrationWizModernAuthAppRegistration -appName MigrationWiz-ModernAuth-EWS -migTenant Destination
    Creates a new Modern Authentication Application for Migration Wizard for the Destination tenant
.EXAMPLE
    New-MgGraphMigrationWizModernAuthAppRegistration -appName MigrationWiz-ModernAuth-EWS -migTenant Source
    Creates a new Modern Authentication Application for Migration Wizard for the Source tenant
.NOTES


#>

function New-MigrationWizModernAuthAppRegistrationMGGraph {
    param (
        [Parameter(Mandatory = $true, HelpMessage ="Provide a name for the Migration Wizard Modern Authentication Application. Example: MigrationWiz-ModernAuth-EWS")]
        [string]$appName,
        [Parameter(Mandatory = $true, HelpMessage = "For Migration Wiz Project, is this the Source or Destination tenant? (Source/Destination)")]
        [ValidateSet("Source", "Destination")]
        [string]$migTenant
    )

    # Get the Tenant Detail - Azure AD
    $tenantDetail = Get-MgOrganization
    $tenantID = $tenantDetail.Id

    # Define Source or Destination Tenant
    #if ($migTenant -eq $null){
    #    $migTenant = Read-Host -Prompt "For the Migration Wiz Project, is this the Source or Destination tenant? (Source/Destination)"
    #}
    if ($migTenant -eq "Source"){
        $AdvancedOptionsTenant = "Export"
    }elseif ($migTenant -eq "Destination"){
        $AdvancedOptionsTenant = "Import"
    }
    #else{
    #   Write-Host "Invalid Entry. Please enter Source or Destination"
    #    $migTenant = Read-Host -Prompt "For the Migration Wiz Project, is this the Source or Destination tenant? (Source/Destination)"
    #}

    # Create the app - MGGraph
    Write-Host "Creating the Modern Authentication Application for Migration Wiz" -BackgroundColor Cyan -ForegroundColor Black
    $newRedirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $publicClientApplication = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphPublicClientApplication
    $publicClientApplication.RedirectUris = $newRedirectUri
    $homePageURL = "https://help.bittitan.com/hc/en-us/articles/360034124813-Authentication-Methods-for-Microsoft-365-All-Products-Migrations#enabling-modern-authentication-for-ews-between-migrationwiz-and-your-exchange-online-tenant-0-10"

    # Check if the app exists; If the app doesn't exist, create it
    if (!($app = Get-MgApplication -Filter "displayName eq '$appName'")) {
        $app = New-MgApplication -DisplayName $appName -SignInAudience "AzureADMultipleOrgs" -IsFallbackPublicClient -PublicClient $publicClientApplication -Web @{HomePageURL = $homePageURL} 
        $appClientID = $app.AppId
        Write-Host "Created $appName new application with AppId: $($app.AppId)" -ForegroundColor Green
        # Add delay if necessary
        #Set-MgApplication -ApplicationId $app.AppId -AllowPublicClient $true
        Write-Output "Waiting for $appName to be created..."
        Start-Sleep -Seconds 15

        $redirectUris = $app.Web.RedirectUris
        if ($null -eq $redirectUris) { $redirectUris = @() }
        $redirectUris += $newRedirectUri
    }
    else {
        # If app exists, add new redirect URI
        $appClientID = $app.AppId        
        Write-Warning "Application already exists. AppId: $($app.AppId)"

    }

    Write-Host "Checking Service Principal Has Been Created and Permissions Assigned" -BackgroundColor Cyan -ForegroundColor Black
    # Create a service principal for the app - MGGraph
    if ($servicePrincipalClientID = (Get-MgServicePrincipal -Filter "displayName eq '$($appName)'").ID) {
        Write-Warning "Service Principal already exists. Skipping creation."
        #Write-Output "ServicePrincipal-Client Id: $($servicePrincipalClientID)"        
    }
    else {
        $servicePrincipal = New-MgServicePrincipal -AppId $app.AppId
        $servicePrincipalClientID = (Get-MgServicePrincipal -Filter "displayName eq '$($appName)'").ID
        Write-Host "Created Service Principal for AppId: $($app.AppId)" -ForegroundColor Green
    }

    Write-Output ""
    # Output the ClientId and ResourceId
    #Write-Output "App ClientId: $($appClientID)"
    #Write-Output "ServicePrincipal-Client Id: $($servicePrincipalClientID)"
    #resource ID is specific to the resource (API) that you want to access. 
    #In this case, Office 365 Exchange Online is the resource. The ResourceID for Office 365 Exchange Online is d9e49bfe-e0b2-4070-a0f6-d919f6a31355

    $exchangeOnline = Get-MgServicePrincipal -Filter "DisplayName eq 'Office 365 Exchange Online'"
    $office365ExchangeOnlineResourceId = $exchangeOnline.Id
    #$office365ExchangeOnlineResourceId = "d9e49bfe-e0b2-4070-a0f6-d919f6a31355"
    #Write-Output "Office 365 Exchange Online ResourceId: $($office365ExchangeOnlineResourceId)"
    Write-Host "Grant the Oauth Permissions - EWS for Office 365 Exchange Online" -BackgroundColor Cyan -ForegroundColor Black

    # Output the Service Principal Id
    if ($OAUTHPerms = Get-MgOauth2PermissionGrant -Filter "clientId eq '$servicePrincipalClientID' and consentType eq 'AllPrincipals'") {
        Write-Warning "Oauth Permissions Already Granted. Skipping"

    }
    else {
        # Grant the permissions
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
        #$newOauthPerms = New-MgOauth2PermissionGrant -BodyParameter $params
        try {
            $OAUTHPerms = New-MgOauth2PermissionGrant -BodyParameter $params
        }
        catch {
            Write-Error $_.Exception.Message
        }
        
    }

    #Confirm Permissions Granted
    Write-Output ""
    #Write-Host "Verify Oauth Permissions Granted" -BackgroundColor Yellow -ForegroundColor Black
    #$OAUTHPerms | fl


    # Output the AppClientId and TenantId
    #Write-Output ""
    Write-Host "$($migTenant) Details Needed for MigrationWiz" -BackgroundColor Yellow -ForegroundColor Black
    Write-Output "$($tenantDetail.DisplayName) Application Name: $($appname)"
    Write-Output "$($tenantDetail.DisplayName) Application (Client) ID: ModernAuthClientId$($AdvancedOptionsTenant)=$($appClientID)"
    Write-Output "$($tenantDetail.DisplayName) Tenant Id: ModernAuthTenantId$($AdvancedOptionsTenant)=$($tenantID)"
    #Write-Output "$($newdate) is the expiry date for the permissions"

    Write-Output ""
    
}

New-MigrationWizModernAuthAppRegistrationMGGraph  #-appName $appName -migTenant Destination