@echo off
Echo "Now Installing Dot Net 3.5!"

Timeout 2

wmic os get BuildNumber | findstr "1" > C:\MOP_Install\OsBuildNumber.ini

set /p VAR2=<C:\MOP_Install\OsBuildNumber.ini 

If %VAR2% == 7600	(Dism.exe /online /Enable-Feature /FeatureName:NetFx3)

If %VAR2% == 7601	(Dism.exe /online /Enable-Feature /FeatureName:NetFx3)

If %VAR2% == 17763 (Dism.exe /online /enable-feature /featurename:NetFX3 /source:\\Server\share\folder\Microsoft\Windows10Ent /LimitAccess)

If %VAR2% == 18362 (Dism.exe /online /enable-feature /featurename:NetFX3 /source:\\Server\share\folder\Win10Ver1909 /LimitAccess)

If %VAR2% == 18363 (Dism.exe /online /enable-feature /featurename:NetFX3 /source:\\Server\share\folder\Win10Ver1909 /LimitAccess)

If %VAR2% == 19041 (Dism.exe /online /enable-feature /featurename:NetFX3 /source:\\Server\share\folder\Win10Ver2004 /LimitAccess)

If %VAR2% == 19042 (Dism.exe /online /enable-feature /featurename:NetFX3 /source:\\Server\share\folder\Win10Ver2004 /LimitAccess)

Echo "Dot Net 3.5 Installed!"

Timeout 2