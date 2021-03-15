@echo off
Echo "Now Installing Java!"

Timeout 2

dir C:\MOP_Install /b | findstr /I /M "jre-"> C:\MOP_Install\JavaTemp.txt

set /p VAR2=<C:\MOP_Install\JavaTemp.txt

Timeout 2

"C:\MOP_Install\%VAR2%" INSTALL_SILENT=1 AUTO_UPDATE=0 REBOOT=0 SPONSORS=0 REMOVEOUTOFDATEJRES=1

Timeout 2

del "C:\MOP_Install\JavaTemp.txt"
del "C:\MOP_Install\jre-*.*"

Echo "Java Installed!"

Timeout 2

