[CmdletBinding()]
Param(    
    [Parameter(Mandatory = $false)]
    [switch]$EnableAllPolicies,
    [switch]$DefaultATPPolicyforO365,
    [switch]$SpamPolicy,
    [switch]$SafeslinkPolicy,
    [switch]$SafeAttachmentPolicy,
    [switch]$AntiPhishingPolicy,
    [switch]$MalwareBaselinePolicy,
    [switch]$OutBoundPolicy,
    [switch]$RollbackDeployment,
    [switch]$FullConfiguration
)

$Global:ErrorActionPreference = "SilentlyContinue"

#Capture all current accepted domains in tenant
$RecipientDomains = Get-AcceptedDomain | sort DomainName | select -ExpandProperty DomainName

function Update-DefaultATPPolicyforO365
{
    $atpPolicyCustomSettings=@{
        'EnableSafeLinksForClients' = $true;
        'EnableATPForSPOTeamsODB' = $false;
        'AllowClickThrough' = $false;
        'TrackClicks' = $true
    }
    
    Write-Host "Updating default ATP Policy for Office 365.." -ForegroundColor Yellow

    Set-AtpPolicyForO365 -EnableSafeLinksForO365Clients $true -EnableATPForSPOTeamsODB $false -AllowClickThrough $false -TrackClicks $true

    Write-Host "Update complete!" -ForegroundColor Green
}

function New-SpamBaselinePolicy
{
    $hostedContentFilterPolicyBaseline=@{
        'Name' = 'ATP - Anti-Spam Baseline Policy';
        'AdminDisplayName' = 'ATP - Anti-Spam Baseline Policy';
        'QuarantineRetentionPeriod' = 30;
        'MarkAsSpamBulkMail' = 'On';
        'HighConfidenceSpamAction' = 'Quarantine';
        'SpamAction' = 'MoveToJmf';
        'EnableEndUserSpamNotifications' = $true;
        'EndUserSpamNotificationFrequency' = 3;
        'BulkThreshold' = 6;
        'ZapEnabled' = $true;
        'BulkSpamAction' = 'MoveToJmf';
        'PhishSpamAction' = 'Quarantine';
        'SpamZapEnabled' = $true;
        'PhishZapEnabled' = $true;
        'HighConfidencePhishAction' = 'Quarantine';
    }

    $hostedContentFilterRuleBaseline=@{
        'Name' = 'ATP - Anti-Spam Baseline Rule';
        'HostedContentFilterPolicy' = 'ATP - Anti-Spam Baseline Policy';
        'RecipientDomainIs' = $RecipientDomains;
        'enabled' = $false
        'Priority' = 0
    }

    $policyCheck = Get-HostedContentFilterPolicy 'ATP - Anti-Spam Baseline Policy'
    $ruleCheck = Get-HostedContentFilterRule 'ATP - Anti-Spam Baseline Rule'
    
    Write-Host "Creating ATP Baseline Policy for Anti-Spam..." -ForegroundColor Yellow

        if ($policyCheck)
        {
            Write-Host "'ATP - Anti-Spam Baseline Policy' already exists and will not be created " -ForegroundColor Cyan
        }
        else
        {
            New-HostedContentFilterPolicy @hostedContentFilterPolicyBaseline | Out-Null
        }
        if ($ruleCheck)
        {
            Write-Host "'ATP - Anti-Spam Baseline Rule' already exists and will not be created" -ForegroundColor Cyan
        }
        else
        {   
            New-HostedContentFilterRule @hostedContentFilterRuleBaseline | Out-Null
        }

    Write-Host "Policy creation complete!" -ForegroundColor Green
}

function New-SafeslinkBaselinePolicy
{
    $safeLinksPolicyBaseline=@{
       'Name' = 'ATP - Safe Links Baseline Policy';
       'AdminDisplayName' = 'ATP - Safe Links Baseline Policy';
       'DoNotAllowClickThrough' =  $true;
       'DoNotTrackUserClicks' = $false;
       'DeliverMessageAfterScan' = $true;
       'EnableForInternalSender' = $true;
       'ScanUrls' = $true;
       'TrackClicks' = $true;
       'IsEnabled' = $false
    }

    $safeLinksPolicyStrict=@{
       'Name' = 'ATP - Safe Links Strict Policy';
       'AdminDisplayName' = 'ATP - Safe Links Strict Policy';
       'DoNotAllowClickThrough' =  $true;
       'DoNotTrackUserClicks' = $false;
       'DeliverMessageAfterScan' = $true;
       'EnableForInternalSender' = $true;
       'ScanUrls' = $true;
       'TrackClicks' = $true;
       'IsEnabled' = $false
    }

    $safeLinksRuleBaseline = @{
        'Name' = 'ATP - Safe Links Baseline Rule';
	    'SafeLinksPolicy' = 'ATP - Safe Links Baseline Policy';
	    'RecipientDomainIs' = $RecipientDomains;
	    'Enabled' = $false;
	    'Priority' = 0
}

    $safeLinksRuleStrict = @{
        'Name' = 'ATP - Safe Links Strict Rule';
	    'SafeLinksPolicy' = 'ATP - Safe Links Strict Policy';
	    'RecipientDomainIs' = $RecipientDomains;
	    'Enabled' = $false;
	    'Priority' = 1
    }
    
    Write-Host "Creating ATP Baseline Policy for SafeLinks..." -ForegroundColor Yellow
    
    $policyCheck = Get-SafeLinksPolicy 'ATP - Safe Links Baseline Policy'
    $ruleCheck = Get-SafeLinksRule 'ATP - Safe Links Baseline Rule'

        if ($policyCheck)
        {
            Write-Host "'ATP - Safe Links Baseline Policy' already exists and will not be created " -ForegroundColor Cyan
        }
        else
        {
            New-SafeLinksPolicy @safeLinksPolicyBaseline | Out-Null
        }
    
        if ($ruleCheck)
        {
            Write-Host "'ATP - Safe Links Baseline Rule' already exists and will not be created " -ForegroundColor Cyan
        }
        else
        {
            New-SafeLinksRule @safeLinksRuleBaseline | Out-Null
    }

    Write-Host "Policy creation complete!" -ForegroundColor Green

}

function New-SafeAttachmentBaselinePolicy
{
    $safeAttachmentPolicyBaseline=@{
       'Name' = 'ATP - Safe Attachments Baseline Policy';
       'AdminDisplayName' = 'ATP - Safe Attachments Baseline Policy';
       'Action' =  "Block";
       'ActionOnError' = $true;
       'Enable' = $false;
       'Redirect' = $false
    }

    $safeAttachRuleBaseline=@{
        'Name' = 'ATP - Safe Attachments Baseline Rule';
	    'SafeAttachmentPolicy' = 'ATP - Safe Attachments Baseline Policy';
	    'RecipientDomainIs' = $RecipientDomains;
	    'Enabled' = $false;
	    'Priority' = 0
    }


    Write-Host "Creating ATP Baseline Policy for Safe Attachments..." -ForegroundColor Yellow
    
    $policyCheck = Get-SafeAttachmentPolicy 'ATP - Safe Attachments Baseline Policy'
    $ruleCheck = Get-SafeAttachmentRule 'ATP - Safe Attachments Baseline Rule'
    
        if ($policyCheck)
        {
            Write-Host "'ATP - Safe Attachments Baseline Policy' already exists and will not be created " -ForegroundColor Cyan
        }
        else
        {
            New-SafeAttachmentPolicy @safeAttachmentPolicybaseline | Out-Null
        }
        if ($ruleCheck)
        {
            Write-Host "'ATP - Safe Attachments Baseline Rule' already exists and will not be created " -ForegroundColor Cyan
        }
        else
        {
            New-SafeAttachmentRule @safeAttachRuleBaseline | Out-Null
        }

    Write-Host "Policy creation complete!" -ForegroundColor Green

}

function New-AntiPhishingBaselinePolicy
{
    $phishPolicyBaseline=@{
       'Name' = 'ATP - Anti-Phishing Policy - Baseline';
       'AdminDisplayName' = 'ATP - Anti-Phishing Policy - Baseline';
       'AuthenticationFailAction' =  'MoveToJmf';
       'EnableAntispoofEnforcement' = $true;
       'Enabled' = $false;
       'EnableMailboxIntelligence' = $true;
       'EnableMailboxIntelligenceProtection' = $true;
       'MailboxIntelligenceProtectionAction' = 'Quarantine';
       'EnableOrganizationDomainsProtection' = $true;
       'EnableSimilarDomainsSafetyTips' = $true;
       'EnableSimilarUsersSafetyTips' = $true;
       'EnableTargetedDomainsProtection' = $false;
       'TargetedDomainsToProtect' = $RecipientDomains;
       'EnableTargetedUserProtection' = $false;
       'EnableUnauthenticatedSender' = $true;
       'EnableUnusualCharactersSafetyTips' = $true;
       'PhishThresholdLevel' = 2;
       'TargetedDomainProtectionAction' =  'Quarantine';
       'TargetedUserProtectionAction' =  'Quarantine';
       'ImpersonationProtectionState' = 'Manual'
    }

    $phishRuleBaseline = @{
        'Name' = 'ATP - Anti-Phishing Baseline Rule';
	    'AntiPhishPolicy' = "ATP - Anti-Phishing Policy - Baseline"; 
	    'RecipientDomainis' = $RecipientDomains;
	    'Enabled' = $false;
	    'Priority' = 0
    }

    Write-Host "Creating ATP Baseline Policy for Phishing..." -ForegroundColor Yellow
    
    $policyCheck = Get-AntiPhishPolicy 'ATP - Anti-Phishing Policy - Baseline'
    $ruleCheck = Get-AntiPhishRule 'ATP - Anti-Phishing Baseline Rule'
    
        if ($policyCheck)
        {
            Write-Host "'ATP - Anti-Phishing Policy - Baseline' already exists and will not be created " -ForegroundColor Cyan
        }
        else
        {
            New-AntiPhishPolicy @phishPolicyBaseline | Out-Null
        }

        if ($ruleCheck)
        {
            Write-Host "'ATP - Anti-Phishing Baseline Rule' already exists and will not be created " -ForegroundColor Cyan
        }
        else
        {
            New-AntiPhishRule @phishRuleBaseline | Out-Null
        }

    Write-Host "Policy creation complete!" -ForegroundColor Green
}

function New-MalwareBaselinePolicy
{
    $malwarePolicyBaseline = @{
        'name' = 'ATP - Malware Policy - Baseline';
        'AdminDisplayName' = 'ATP - Malware Policy - Baseline';
        'Action' = 'DeleteMessage';
        'EnableFileFilter' = $true;
        'ZapEnabled' = $true

    }

    $malwareRuleBaseline = @{
        'name' = 'ATP - Malware Baseline Rule';
        'MalwareFilterPolicy' = 'ATP - Malware Policy - Baseline'
        'enabled' = $false;
        'Priority' = 0;
        'RecipientDomainIs' = $RecipientDomains;
    }

    Write-Host "Creating ATP Baseline Policy for Malware..." -ForegroundColor Yellow

    $policyCheck = Get-AntiPhishPolicy 'ATP - Malware Policy - Baseline'
    $ruleCheck = Get-AntiPhishRule 'ATP - Malware Baseline Rule'
    
        if ($policyCheck)
        {
            Write-Host "'ATP - Malware Policy - Baseline' already exists and will not be created" -ForegroundColor Cyan
        }
        else
        {
            New-MalwareFilterPolicy @malwarePolicyBaseline | Out-Null
        }

        if ($ruleCheck)
        {
            Write-Host "'ATP - Malware Baseline Rule' already exists and will not be created" -ForegroundColor Cyan   
        }
        else
        {
            New-MalwareFilterRule @malwareRuleBaseline | Out-Null
        }

    Write-Host "Policy creation complete!" -ForegroundColor Green

}

function Remove-FullDeployment
{
    
    Write-Host "Rolling back deployment done by this script..."

    #set back original default settings for ATP Org policy
    Set-AtpPolicyForO365 -EnableSafeLinksForO365Clients $false -EnableATPForSPOTeamsODB $false -AllowClickThrough $false -TrackClicks $false

    #Remove Spam policy & rule
    Remove-HostedContentFilterPolicy "ATP - Anti-Spam Baseline Policy" -confirm:$false
    Remove-HostedContentFilterRule "ATP - Anti-Spam Baseline Rule" -confirm:$false

    #Remove SafeLinks policy & rule
    Remove-SafeLinksPolicy "ATP - Safe Links Baseline Policy" -confirm:$false
    Remove-SafeLinksRule "ATP - Safe Links Baseline Rule" -confirm:$false

    #Remove Safe Attachments policy & rule
    Remove-SafeAttachmentPolicy "ATP - Safe Attachments Baseline Policy" -confirm:$false
    Remove-SafeAttachmentRule "ATP - Safe Attachments Baseline Rule" -confirm:$false

    #Remove Phising policy & rule
    Remove-AntiPhishPolicy "ATP - Anti-Phishing Policy - Baseline" -confirm:$false
    Remove-AntiPhishRule "ATP - Anti-Phishing Baseline Rule" -confirm:$false

    #Remove Malware policy
    Remove-MalwareFilterPolicy "ATP - Malware Policy - Baseline" -confirm:$false
    Remove-MalwareFilterRule "ATP - Malware Baseline Rule" -confirm:$false

    #Remove Outbound Policy
    Remove-HostedOutboundSpamFilterPolicy 'ATP - Outbound Baseline Policy'
    Remove-HostedOutboundSpamFilterRule 'ATP - Outbound Baseline Rule' -confirm:$false

    Write-Host "Rollback Complete!"
}

function Enable-ATPCustomPolicies
{

    Write-Host "Enabling all custom policies..." -ForegroundColor Yellow
    
    Enable-HostedContentFilterRule "ATP - Anti-Spam Baseline Rule"
    Set-SafeLinksPolicy "ATP - Safe Links Baseline Policy" -isenabled $true
    Enable-SafeLinksRule "ATP - Safe Links Baseline Rule"
    Set-SafeAttachmentPolicy "ATP - Safe Attachments Baseline Policy" -enable $true
    Enable-SafeAttachmentRule "ATP - Safe Attachments Baseline Rule"
    Set-AntiPhishPolicy "ATP - Anti-Phishing Policy - Baseline" -enabled $true
    Enable-AntiPhishRule "ATP - Anti-Phishing Baseline Rule"
    Enable-MalwareFilterRule "ATP - Malware Baseline Rule"
    Enable-HostedOutboundSpamFilterRule "ATP - Outbound Baseline Rule"

    Write-Host "All policies are enabled" -ForegroundColor Green

}

function New-OutboundBaslinePolicy
{
    $outboundPolicyBaseline = @{
        'Name' = 'ATP - Outbound Baseline Policy';
        'AdminDisplayName' = 'ATP - Outbound Baseline Policy';
        'RecipientLimitExternalPerHour' = 500;
        'RecipientLimitInternalPerHour' = 1000;
        'RecipientLimitPerDay' = 1000;
        'ActionWhenThresholdReached' = 'BlockUser'
    }

    $outboundBaselineRule = @{
        'Name' = 'ATP - Outbound Baseline Rule';
        'Enabled' = $false;
        'SenderDomainIs' = $RecipientDomains;
        'HostedOutboundSpamFilterPolicy' = 'ATP - Outbound Baseline Policy';
        'Priority' = 0
    }
    
    Write-Host "Creating Outbound Policy..." -ForegroundColor Yellow

    $policyCheck = Get-AntiPhishPolicy 'ATP - Outbound Baseline Policy'
    $ruleCheck = Get-AntiPhishRule 'ATP - Outbound Baseline Rule'

    if ($policyCheck)
    {
        Write-Host "'ATP - Outbound Baseline Policy' already exists and will not be created" -ForegroundColor Cyan
    }
    else
    {
        New-HostedOutboundSpamFilterPolicy @outboundPolicyBaseline | Out-Null
    }

    If ($ruleCheck)
    {
        Write-Host "'ATP - Outbound Baseline Rule' already exists and will not be created" -ForegroundColor Cyan
    }
    else
    {
        New-HostedOutboundSpamFilterRule @outboundBaselineRule | Out-Null
    }

    Write-Host "Policy created!" -ForegroundColor Green

}

function Add-AllATPPolicies
{
    
    Write-Host "Deploying full ATP Baseline Policies..." -ForegroundColor Yellow

    Update-DefaultATPPolicyforO365 | Out-Null
    New-SpamBaselinePolicy | Out-Null
    New-SafeslinkBaselinePolicy | Out-Null
    New-SafeAttachmentBaselinePolicy | Out-Null
    New-AntiPhishingBaselinePolicy | Out-Null
    New-MalwareBaselinePolicy | Out-Null
    New-OutboundBaslinePolicy | Out-Null

    Write-Host "All policy creations complete!" -ForegroundColor Green
}


$selectionStatus = $PSBoundParameters.GetEnumerator() | %{$_ | ?{$_.Value -eq $true}}

If($selectionStatus)
{
    Switch ($PSBoundParameters.GetEnumerator().
        Where({$_.Value -eq $true}).Key)
        {
        
            'DefaultATPPolicyforO365' { Update-DefaultATPPolicyforO365 }
            'SpamPolicy' { New-SpamBaselinePolicy }
            'SafeslinkPolicy' { New-SafeslinkBaselinePolicy }
            'SafeAttachmentPolicy' { New-SafeAttachmentBaselinePolicy }
            'AntiPhishingPolicy' { New-AntiPhishingBaselinePolicy }
            'FullConfiguration' { Add-AllATPPolicies }
            'MalwareBaselinePolicy' { New-MalwareBaselinePolicy }
            'RollbackDeployment' { Remove-FullDeployment }
            'EnableAllPolicies' { Enable-ATPCustomPolicies }
            'OutBoundPolicy' { New-OutboundBaslinePolicy}

        }
}

else
{
    Write-Host "No value selected...please select a value!" -ForegroundColor Red
}

