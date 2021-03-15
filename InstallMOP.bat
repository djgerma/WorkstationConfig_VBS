@echo off

Echo "Now installing Merlin Operator Portal!"

Timeout 2

dir C:\MOP_Install\ /b | findstr /R /I /M "merlin" | findstr /e .exe> C:\MOP_Install\MerlinTemp.txt

Timeout 2

set /p VAR=<C:\MOP_Install\MerlinTemp.txt

"C:\MOP_Install\%VAR%" /extract C:\MOP_Install\

Timeout 2

dir C:\MOP_Install\ /b | findstr /R /I /M "merlin" | findstr /e .msi> C:\MOP_Install\MerlinTemp.txt

set /p VAR2=<C:\MOP_Install\MerlinTemp.txt

"C:\MOP_Install\%VAR2%" /passive

Timeout 2

del "C:\MOP_Install\Merlin*.*"

Echo "Merlin Operator Portal Installed!"

Timeout 2

::shutdown /r /f /t 60
