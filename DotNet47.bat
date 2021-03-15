@echo off
Echo "Now Installing Dot Net 4.7.2!"

Timeout 2

dir C:\MOP_Install /b | findstr /I /M "NDP472"> C:\MOP_Install\DotNetTemp.txt

set /p VAR2=<C:\MOP_Install\DotNetTemp.txt

Timeout 2

"C:\MOP_Install\%VAR2%" /q /norestart

Timeout 2

del "C:\MOP_Install\DotNetTemp.txt"

Echo ".NET 4.7.2 Installed!"

Timeout 2