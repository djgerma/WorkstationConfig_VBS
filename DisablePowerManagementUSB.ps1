$hubs = Get-WmiObject Win32_USBControllerDevice | Select-Object Name,DeviceID,Description
$powerMgmt = Get-WmiObject MSPower_DeviceEnable -Namespace root\wmi
foreach ($p in $powerMgmt)
{
 $IN = $p.InstanceName.ToUpper()
 foreach ($h in $hubs)
 {
  $PNPDI = $h.PNPDeviceID
                if ($IN -like "*$PNPDI*")
                {
                     $p.enable = $False
                     $p.psbase.put()
                }
 }
}
if ((Test-Path -Path "C:\TimeClock")) {
New-Item -Path 'C:\TimeClock\USBPowerManagement.dat' -ItemType File
} else {
exit
}