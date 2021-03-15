@echo off
Echo "Now Installing Symantec AntiVirus Software"

Timeout 2

dir C:\MOP_Install /b | findstr /I /M "setup"> C:\MOP_Install\InstallSAV.txt

set /p VAR2= < C:\MOP_Install\InstallSAV.txt

Timeout 2

"C:\MOP_Install\%VAR2%" /s

Timeout 2

del "C:\MOP_Install\InstallSAV.txt"

