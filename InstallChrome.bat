@echo off
Echo "Now Installing Google Chrome!"

Timeout 2

dir C:\MOP_Install /b | findstr /I /M "Google"> C:\MOP_Install\ChromeTemp.txt

set /p VAR2=<C:\MOP_Install\ChromeTemp.txt

msiexec.exe /i "C:\MOP_Install\%VAR2%" /quiet /passive

Timeout 2

del "C:\MOP_Install\ChromeTemp.txt"
del "C:\MOP_Install\GoogleChrome*.msi"

Echo "Google Chrome Installed!"

Timeout 2

Echo "Your MOP is now configured! Computer will reboot in 1 minute!"

Timeout 2

rem shutdown /r /f /t 60

