function Get-EXOMailboxForwards
{
    <#

        .SYNOPSIS
            Discover Exchange Online Mailbox Forwarding Configuration including Inbox Rules

        .DESCRIPTION
            Ingests the Recipients list from Get-Recipient to display the Get-EXOMailbox Forwarding properties. Uses native Exchange cmdlets to discover InboxRules.

        .OUTPUTS
            Returns a custom object containing Exchange Online Mailbox Forwarding Configuration

        .EXAMPLE
            Get-Get-EXOMailboxForwards -Recipients $exchangeEnvironment["Recipients"]

    #>

    [CmdletBinding()]
    param (
        # An array of Recipients to run discovery against
        [array]
        $Recipients
    )
    $activity = "Exchange Online Mailbox Forwards"
    $discoveredMailboxForward = @()
    $ALLMailboxForward = @()
   
    Write-Log -Level "VERBOSE" -Activity $activity -Message "Gathering Exchange Online Recipient Forwarding Configuration." -WriteProgress
    
    foreach ($exoMailbox in $Recipients)
    {
        try
        {
            Write-Host "Getting recipient Forward object $($exoMailbox.Guid)." -ForegroundColor cyan -NoNewline
            $discoveredMailboxForward = "" | Select-Object ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxandForward, ForwardingRulesConfigured

            if (($null -notlike $exoMailbox.ForwardingAddress) -or ($null -notlike $exoMailbox.ForwardingSmtpAddress))
            {
                $discoveredMailboxForward.ForwardingAddress = $exoMailbox.ForwardingAddress
                $discoveredMailboxForward.ForwardingSmtpAddress = $exoMailbox.ForwardingSmtpAddress
                $discoveredMailboxForward.DeliverToMailboxandForward = $exoMailbox.DeliverToMailboxandForward
            }

            else
            {
                $discoveredMailboxForward.ForwardingAddress = "null"
                $discoveredMailboxForward.ForwardingSmtpAddress = "null"
                $discoveredMailboxForward.DeliverToMailboxandForward = "False"
            }
        }
        catch
        {
            #Write-Log -Level "WARNING" -Activity $activity -Message "Exchange Online Mailbox Forwarding result null for $($exoMailbox.Guid)."
            Write-Host "Exchange Online Mailbox Forwarding result null for $($exoMailbox.Guid)." -ForegroundColor cyan
        }
        
        try 
        {
            #Write-Log -Level "VERBOSE" -Activity $activity -Message "Getting recipient Forward Inbox Rules for object $($exoMailbox.Guid)." -WriteProgress
            Write-Host "Getting recipient Forward Inbox Rules for object $($exoMailbox.Guid)." -ForegroundColor cyan -NoNewline
            $exoMailboxRules = Get-InboxRule -mailbox $exoMailbox.Guid.ToString() | Where-Object {($NULL -ne $_.forwardAsAttachmentTo) –or ($NULL -ne $_.forwardTo) –or ($NULL -ne $_.redirectTo)} -ErrorAction stop
        }
        catch 
        {
            #Write-Log -Level "WARNING" -Activity $activity -Message "Failed to run Get-InboxRule against object $($exoMailbox.Guid). $($_.Exception.Message)"
            Write-Host "Failed to run Get-InboxRule against object $($exoMailbox.Guid)" -ForegroundColor Red
        }
        
        if ($null -notlike $exoMailboxRules)
        {
            Write-Host "InboxRules found for $($exoMailbox.Guid)." -ForegroundColor Green
            $discoveredMailboxForward.ForwardingRulesConfigured = "True"
        }
        else
        {
            #Write-Log -Level "WARNING" -Activity $activity -Message "Get-InboxRule result null for $($exoMailbox.Guid)."
            Write-Host "Get-InboxRule result null for $($exoMailbox.Guid)" -ForegroundColor Red
            $discoveredMailboxForward.ForwardingRulesConfigured = "False"
        }
        #$currentRecipient.Forwarding = $discoveredMailboxForward
        $ALLMailboxForward += $discoveredMailboxForward
    }
    
}  