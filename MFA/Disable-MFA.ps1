#Disable MFA
$pilotusers | foreach {Set-MsolUser -UserPrincipalName $_.identity -StrongAuthenticationRequirements @()}

<# Get the Current Strong Authentication Methods (SAM) and Strong Authentication Requirements (SAR) for users
The SAR has a value for "RememberDevicesNotIssuedBefore"
Example to store SAR as array
$SAR = Get-MsolUser -UserPrincipalName zmfa_test@ccmsi.com | select -ExpandProperty strongauthenticationrequirements

Example to clear
Set-msoluser -UserPrincipalName zmfa_test@ccmsi.com -StrongAuthenticationRequirements @()

Example to set value:
Set-msoluser -UserPrincipalName zmfa_test@ccmsi.com -StrongAuthenticationRequirements $SAR
ExtensionData                                    RelyingParty RememberDevicesNotIssuedBefore State
-------------                                    ------------ ------------------------------ -----
System.Runtime.Serialization.ExtensionDataObject *            11/4/2020 5:36:37 PM           Enabled



Example to get SAM
$SAM = Get-MsolUser -UserPrincipalName zmfa_test@ccmsi.com | select -ExpandProperty strongauthenticationmethods
PS C:\Users\fred5646> $STAM                                                                                             
ExtensionData                                    IsDefault MethodType
-------------                                    --------- ----------
System.Runtime.Serialization.ExtensionDataObject     False PhoneAppOTP
System.Runtime.Serialization.ExtensionDataObject      True PhoneAppNotification

#>


#EnableMFA

foreach ($user in $users)
{
    if ($MSOLUSER = Get-MsolUser -UserPrincipalName $user -erroraction silentlycontinue)
    {
        Write-Host "Enable MFA for user"
        $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
        $st.RelyingParty = "*"
        $st.State = "Enabled"
        $enableMFA = @($st)
         
        #Enable MFA
        Set-msoluser -UserPrincipalName $user -StrongAuthenticationRequirements $enableMFA
    }

}

#testing to get mfa disable and gather details

foreach ($user in $users)
{
    #Gather SAM and SAR for MFA
    $MSOLUSER = Get-MsolUser -UserPrincipalName $user
    $SAR = $MSOLUSER.StrongAuthenticationRequirements
    $SAM = $MSOLUSER.StrongAuthenticationMethods

    $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $st.UserPrincipalName = $user
    $st.RelyingParty = "*"
    $st.State = "Enabled"
    $st.RememberDevicesNotIssuedBefore = $SAR.RememberDevicesNotIssuedBefore
    $MFASetting += @($st)   
    
    #Disable MFA
    Set-MsolUser -UserPrincipalName $user -StrongAuthenticationRequirements @()

}