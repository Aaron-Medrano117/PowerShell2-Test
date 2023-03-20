$mailboxes =  get-mailbox -OrganizationalUnit americancarcenter.com -ResultSize unlimited

$NoMobileDeviceUsers = @()
$discoveredMobileDevices = @()

foreach ($mailbox in $mailboxes)
{
    $ActiveSyncDeviceStatistics = Get-ActiveSyncDeviceStatistics -Mailbox $mailbox
    if ($ActiveSyncDeviceStatistics.count -gt "0")
    {
        Write-Host "$($ActiveSyncDeviceStatistics.count) Devices found for $($mailbox.displayname)"

        foreach ($device in $ActiveSyncDeviceStatistics)
        {
            $mobiledevicestats =@()
            $mobiledevicestats = $device | Select-Object LastSuccessSync, DeviceModel, DeviceFriendlyName, DeviceType, FirstSyncTime, DeviceID, DeviceOS

            $tmp = "" | Select-Object DisplayName, EmailAddress, DeviceModel, DeviceFriendlyName, DeviceType, LastSuccessSync, FirstSyncTime, DeviceID, DeviceOS
            $tmp.DisplayName = $mailbox.DisplayName
            $tmp.EmailAddress = $mailbox.PrimarySMTPAddress
            $tmp.DeviceModel = $mobiledevicestats.DeviceModel
            $tmp.DeviceFriendlyName = $mobiledevicestats.DeviceFriendlyName
            $tmp.DeviceType = $mobiledevicestats.DeviceType
            $tmp.LastSuccessSync = $mobiledevicestats.LastSuccessSync
            $tmp.FirstSyncTime = $mobiledevicestats.FirstSyncTime
            $tmp.DeviceID = $mobiledevicestats.DeviceID
            $tmp.DeviceOS = $mobiledevicestats.DeviceOS
    
            $discoveredMobileDevices += $tmp
        } 
    }
    else
    {
        $NoMobileDeviceUsers += $mailbox
    }

    $discoveredMobileDevices | Export-Csv -Path C:\Users\aaron.medrano1\Desktop\AmericanCarCenter\americancarcenter_mobiledevices.csv -Encoding UTF8 -NoTypeInformation
}
