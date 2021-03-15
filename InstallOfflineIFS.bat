@echo off
Echo "Now Installing Offline IFS TimeClock!"

Timeout 2

dir C:\MOP_Install /b | findstr /I /M "TimeClockSetup"> C:\MOP_Install\OfflineIFSTemp.txt

set /p VAR2=<C:\MOP_Install\OfflineIFSTemp.txt

msiexec.exe /i "C:\MOP_Install\%VAR2%" /quiet /passive

Timeout 2

del "C:\MOP_Install\OfflineIFSTemp.txt"
del "C:\TimeClock\InstallNewIFS.txt"

Echo "Offline Time Clock Installed!"

Timeout 2



