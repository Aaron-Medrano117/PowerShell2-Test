$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$DriverName = “imageRUNNER ADVANCE DX C5860i"
$DriverPath = “$PSScriptRoot\Driver”
$DriverInf = “$PSScriptRoot\Driver\Cnp60MA64.INF”
$portName = “172.20.10.20“

$checkPortExists = Get-Printerport -Name $portname -ErrorAction SilentlyContinue

if (-not $checkPortExists) {

Add-PrinterPort -name $portName -PrinterHostAddress “172.20.10.20“
}
cscript “C:\Windows\System32\Printing_Admin_Scripts\en-US\Prndrvr.vbs” -a -m “Canon Generic Plus PCL6 Suite590” -h $DriverPath -i $PSScriptRoot\Driver\Cnp60MA64.INF
$printDriverExists = Get-PrinterDriver -name $DriverName -ErrorAction SilentlyContinue

if ($printDriverExists)
{
Add-Printer -Name “Cannon C355i Suite590” -PortName $portName -DriverName $DriverName
}
else
{
Write-Warning “Printer Driver not installed”
}

## own printer installation powershell
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$DriverName = “imageRUNNER ADVANCE DX C5860i"
$DriverPath = “$PSScriptRoot\Driver”
$DriverInf = “$PSScriptRoot\Driver\Cnp60MA64.INF”
$portName = “172.20.10.20“
$PrinterName = "Canon Copier Suite 401"

$checkPortExists = Get-Printerport -Name $portname -ErrorAction SilentlyContinue

if (-not $checkPortExists) {

Add-PrinterPort -name $portName -PrinterHostAddress $portName
}
cscript “C:\Windows\System32\Printing_Admin_Scripts\en-US\Prndrvr.vbs” -a -m $DriverName -h $DriverPath -i $DriverInf
$printDriverExists = Get-PrinterDriver -name $DriverName -ErrorAction SilentlyContinue

if ($printDriverExists)
{
Add-Printer -Name $PrinterName -PortName $portName -DriverName $DriverName
}
else
{
Write-Warning “Printer Driver not installed”
}