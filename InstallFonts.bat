@ECHO OFF
XCOPY "C:\TimeClock\3of9.ttf" "%systemroot%\fonts"
REG ADD "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" /v "3of9 (True Type)" /t REG_SZ /d "3of9.ttf" /f