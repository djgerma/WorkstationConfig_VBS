'##########################################Variable Declarations##################################################################################################
'SECTION 1
On Error Resume Next
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strUserProfile
Dim strAppRestart
Dim strDNC
Dim strInstallJava
Dim BtnCode
Dim i

strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
strAppRestart = "C:\TimeClock\IFSAppRestart.vbs"
strOSKRestart = "C:\TimeClock\OSKRestart.vbs"
strMOPRestart = "C:\TimeClock\MOPRestart.vbs"
strInstallMOP = "C:\TimeClock\InstallMOP.bat"
strCaffeine = "C:\TimeClock\caffeine.exe"
strMOP = """C:\Program Files (x86)\Memex Inc\Merlin Operator Portal (Ethernet)\merlin_operator_portal_ethernet\bin\merlin_operator_portal_ethernet.exe"""
strInstallJava = "C:\TimeClock\InstallJava.bat"
strInstallChrome = "C:\TimeClock\InstallChrome.bat"
strInstallDNet35 = "C:\TimeClock\DotNet35.bat"
strInstallDNet47 = "C:\TimeClock\DotNet47.bat"
'strPinIt = "C:\TimeClock\PinIt.exe"
strOfflineAppRestart = "C:\TimeClock\IFSOfflineAppRestart.vbs"
strInstallOfflineIFS = "C:\TimeClock\OfflineTClock.vbs"
strInstallFonts = "C:\TimeClock\InstallFonts.bat"
strUSBPowerMgmt = "C:\TimeClock\DisablePowerManagementUSB.ps1"
strInstallSAV = "C:\TimeClock\InstallSAV.bat"
strDefaultBrowser = "C:\TimeClock\ChangeDefaultBrowser.bat"
strWorkBenchRestart = "C:\TimeClock\IFSWorkBenchRestart.vbs"
strSoftwareReporterFix = "C:\TimeClock\software_reporter_fix.bat"
strMDT = "TS%"
strNow = Now()


strJava = "C:\MOP_Install\oracleJava.ini"
strDNet35 = "C:\MOP_Install\DNet35.ini"
strDNet47 = "C:\MOP_Install\DNet47.ini"
strUSBMgmt = "C:\TimeClock\USBPowerManagement.dat"

'It takes MDT application on average 10 to 15 seconds to start up. This is to make sure I give it time to actually start before checking if it is still working
objShell.Popup " Script will wait 30 seconds to make sure MDT Server is not working in the background! ", 3, "Warning", 64
WScript.Sleep 27000

'Section 2 Find and store local computername into variable for later usage
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
'Wscript.Echo strComputerName

'Actually check to see if MDT application is up and running
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select Name from Win32_Process WHERE Name LIKE '" & strMDT & "%'")

If colProcessList.count>0 then
    BtnCode= objShell.Popup ("MDT Application is running. You should wait until it finishes to click OK! MOP Setup Script will not continue until you click Ok on this message or 10 minutes passes from now!(" & strNow & ")", 600)
End If	


'Section 2.1
'Check if UAC is on. If it is, GPO did not apply. Reboot. Otherwise continue
If objShell.RegRead("HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA") > 0 Then
BtnCode= objShell.Popup ("UAC is turned ON. GPO did not take full effect. Computer will reboot now in 5 seconds", 5)
objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -f -t 0"
End If

	
'SECTION 3
'Disable USB Power Management so NFC and Barcode readers do not shut off
If Not objFSO.FileExists(strUSBMgmt) Then
	objShell.run("powershell -file C:\TimeClock\DisablePowerManagementUSB.ps1")
End If

'#######################################################################################################################################################################
'SECTION 4
'Always copy *.vbs and *.bat files. This is to make sure any updated scripts get pushed down every time. 
If objFSO.FolderExists("C:\TimeClock") Then
objFSO.CopyFile "\\Server\share\folder\TimeClock\*.vbs", "C:\TimeClock\"
objFSO.CopyFile "\\Server\share\folder\TimeClock\*.bat", "C:\TimeClock\"
objFSO.CopyFile "\\Server\share\folder\TimeClock\*.jpg", "C:\TimeClock\"
objFSO.CopyFile "\\Server\share\folder\TimeClock\*.txt", "C:\TimeClock\"
objFSO.CopyFile "\\Server\share\folder\TimeClock\*.ttf", "C:\TimeClock\"
objFSO.CopyFile "\\Server\share\folder\TimeClock\*.ps1", "C:\TimeClock\"
objFSO.CopyFile "\\Server\share\folder\TimeClock\caffeine.exe", "C:\TimeClock\"
'objFSO.CopyFile "\\Server\share\folder\TimeClock\PinIt.exe", "C:\TimeClock\"
'objFSO.CopyFile "\\Server\share\folder\TimeClock\AddOn\TClock9\*.*", "C:\TimeClock\"
'objFSO.DeleteFile "\\Server\share\folder\TimeClock\AddOn\TClock9\*.txt"
BtnCode= objShell.Popup ("C:\TimeClock\ Folder updated with new scripts", 3)
End If

'#######################################################################################################################################################################
'SECTION 5
'5.1. Check for TimeClock Folder and donotdelete.txt. If donotdelete.txt exists do not copy otherwise copy TimeClock folder contents
'Time clock Folder gets created automatically by GPO MOP LOCKDOWN>User Configuration>Preferences>Windows Settings>Folders
If objFSO.FolderExists("C:\TimeClock") Then
	If Not objFSO.FileExists("C:\TimeClock\DoNotDelete.txt") Then
		objFSO.CopyFile "\\Server\share\folder\TimeClock\*.*", "C:\TimeClock\"
		BtnCode= objShell.Popup ("TimeClock Setup files are getting copied to C:\TimeClock\", 3)
	End If
End If


'#######################################################################################################################################################################
'5.2. Check if Java is installed and what version. If not installed, install it
If Not objFSO.FileExists(strJava) Then

	Const HKEY_CURRENT_USER = &H80000001
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const REG_SZ = 1
	blnJavaInstalled = False
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colProducts = objWMIService.ExecQuery("SELECT Caption FROM Win32_Product")
	For Each objProduct in colProducts
		If Err.Number = 0 Then
		If InStr(UCase(objProduct.Name),"JAVA") Then
		'Wscript.Echo "Caption=" & objProduct.Caption
		blnJavaInstalled = True
		End If
		Else
			hDefKey = HKEY_LOCAL_MACHINE
			strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
			Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
			objReg.EnumKey hDefKey, strKeyPath, arrSubKeys
			For Each strSubkey In arrSubKeys
			strSubKeyPath = strKeyPath & "\" & strSubkey
			objReg.EnumValues hDefKey, strSubKeyPath, arrValueNames, arrTypes
			For i = LBound(arrValueNames) To UBound(arrValueNames)
			strValueName = arrValueNames(i)
			If arrTypes(i) = REG_SZ Then
			objReg.GetStringValue hDefKey, strSubKeyPath, strValueName, strValue
			If InStr(UCase(strValue),"JAVA") Then
				'Wscript.Echo strValueName & " = " & strValue
				blnJavaInstalled = True
				End If
			End If
			Next
			Next
		End If
	Next
If blnJavaInstalled <> True Then
	BtnCode= objShell.Popup ("Java was not found on your computer. It will be installed automatically!", 3)
	'This will copy Java installer to install directory
	objFSO.CopyFile "\\Server\share\folder\folder\JavaInstall\*.*", "C:\MOP_Install\"
	objShell.Run strInstallJava,1,True
	Set objFile = objFSO.CreateTextFile(strJava,True)
	BtnCode= objShell.Popup ("Java Installed!", 3)
Else
	Set objFile = objFSO.CreateTextFile(strJava,True)
End If
End If

'#######################################################################################################################################################################
'5.3 Check if DotNet is installed, if not installed, install it. 3.5 is required for IFS TimeClock and 4.6 + is required for MOP. 
If Not objFSO.FileExists(strDNet47) Then

Dim Act :Set Act = CreateObject("Wscript.Shell")
Dim Obj, Rg1, Rst

Dim Reg :Reg = Array("3.5 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5\Version")
On Error Resume Next 
    For Each Obj In Reg
     Rg1 = Split(Obj," - ")
     If IsNull(Act.RegRead(Rg1(1))) Then
      Rst = Rst &  "Missing Net Frame Work : " & Rg1(0) & vbCrLf 
     Else
      Rst = (Left(Act.RegRead(Rg1(1)),3)) & vbCrLf 
	 End If
    Next  
	
If (Left(Rst,3)) <> "3.5" Then
BtnCode= objShell.Popup (".Net Framework is not installed on your Computer. It will be installed Automatically!", 3)
'This will copy .Net4.x installer over to install directory
objFSO.CopyFile "\\Server\share\folder\folder\DotNet\4.7.2_Win10_Any\*.*", "C:\MOP_Install\"
objShell.Run strInstallDNet35,1,True
objShell.Run strInstallDNet47,1,True
Set objFile = objFSO.CreateTextFile(strDNet47,True)
BtnCode= objShell.Popup (".Net 3.5 and 4.7 installed on your Computer!", 3)
Else
Set objFile = objFSO.CreateTextFile(strDNet47,True)
End if
End if

'#######################################################################################################################################################################
'5.4. Check if MOP is installed, if not installed, install it.
If Not objFSO.FileExists("C:\Program Files (x86)\Memex Inc\Merlin Operator Portal (Ethernet)\merlin_operator_portal_ethernet\bin\merlin_operator_portal_ethernet.exe") Then
BtnCode= objShell.Popup ("Merlin Operator Portal Not Installed! It will be Installed Automatically!", 3)
'This will copy install files to install directory
objFSO.CopyFile "\\server\share\folder\software\Memex\MOP\Merlin Operator Portal\*.*", "C:\MOP_Install\"
WScript.Sleep(5000)
objShell.Run strInstallMOP,1,True
BtnCode= objShell.Popup ("Merlin Operator Portal Installed!", 3)
End If

'#######################################################################################################################################################################
'5.5. Check if Google Chrome is installed. If not installed, install it. 
If Not ((objFSO.FileExists("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") OR objFSO.FileExists("C:\Program Files\Google\Chrome\Application\chrome.exe"))) Then
BtnCode= objShell.Popup ("Google Chrome not Installed! It will be installed Automatically!", 3)
'This will copy install files to install directory
objFSO.CopyFile "\\server\share\folder\software\Memex\Chrome\*.*", "C:\MOP_Install\"
WScript.Sleep(5000)
objShell.Run strInstallChrome,1,True
BtnCode= objShell.Popup ("Google Chrome Installed!", 3)
End If

' Check if 3of9 Fonts is installed, if not, install it
If Not objFSO.FileExists("C:\Windows\Fonts\3of9.ttf") Then
	objShell.Run strInstallFonts,1,True
End If


'5.7 Check if Symantec AV is installed, if not install it and reboot...
If Not ObjFSO.FileExists("C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\DoScan.exe") Then
	If (UCase(Left(strComputerName,3))) = "XXX" Then
	objFSO.CopyFile "\\Server\share\folder\Software\Basic Software\08 Symantec\Current_Win64_Workstation\*.*", "C:\MOP_Install\"
	Else If (UCase(Left(strComputerName,3))) = "XXX" Then
	objFSO.CopyFile "\\Server\Share\Folder\applications\Symantec Endpoint Protection\*.*", "C:\MOP_Install\"
	Else
	objFSO.CopyFile "\\Server\share\folder\TimeClock\AddOn\Current_Win64_Workstation\*.*", "C:\MOP_Install\"
	End If 
	End If
	objShell.Run strInstallSAV,1,True
End If

'#######################################################################################################################################################################
'5.6. Check if Offline TimeClock is installed, if not install it if text file exists...
If ObjFSO.FileExists("C:\TimeClock\InstallNewIFS.txt") Then
	If Not ObjFSO.FileExists("C:\Program Files (x86)\IFS Applications\Time Clock\TimeClockExe.exe") Then
	objShell.Run strInstallOfflineIFS
	End If
End If

'#######################################################################################################################################################################
'5.8. MAP all neccessary shortcuts (MOP, Merlin DNC, File Explorer, IE with TCLOCK)
Set lnk = objShell.CreateShortcut(strUserProfile & "\Desktop\Merlin Operator Portal.LNK")
If Not objFSO.FileExists(lnk) Then
   lnk.TargetPath = "C:\Program Files (x86)\Memex Inc\Merlin Operator Portal (Ethernet)\merlin_operator_portal_ethernet\bin\merlin_operator_portal_ethernet.exe"
   lnk.Arguments = ""
   lnk.Description = "Merlin Operator Portal"
   'lnk.HotKey = ""
   lnk.IconLocation = "C:\Program Files (x86)\Memex Inc\Merlin Operator Portal (Ethernet)\merlin_operator_portal_ethernet\bin\merlin_operator_portal_ethernet.exe"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "C:\Program Files (x86)\Memex Inc\Merlin Operator Portal (Ethernet)\merlin_operator_portal_ethernet\bin\"
   lnk.Save
Set lnk = Nothing
End If

'Merlin DNC Shortcut
If (UCase(Left(strComputerName,3))) = "XXX" Then
	Set lnk = objShell.CreateShortcut(strUserProfile & "\Desktop\Merlin DNC.LNK")
	If Not objFSO.FileExists(lnk) Then
		lnk.TargetPath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
		lnk.Arguments = "http://some.url.local/path/path"
		lnk.Description = "Merlin DNC"
		'lnk.HotKey = ""
		lnk.IconLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
		lnk.WindowStyle = "1"
		lnk.WorkingDirectory = "C:\Program Files (x86)\Google\Chrome\Application"
		lnk.Save
		Set lnk = Nothing
		'Wscript.Echo (UCase(Left(strComputerName,3)))
	End If
Else If (UCase(Left(strComputerName,3))) = "XXX" Then
	Set lnk = objShell.CreateShortcut(strUserProfile & "\Desktop\Merlin DNC.LNK")
	If Not objFSO.FileExists(lnk) Then
		lnk.TargetPath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
		lnk.Arguments = "http://some.url.local/path/path"
		lnk.Description = "Merlin DNC"
		'lnk.HotKey = ""
		lnk.IconLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
		lnk.WindowStyle = "1"
		lnk.WorkingDirectory = "C:\Program Files (x86)\Google\Chrome\Application"
		lnk.Save
		Set lnk = Nothing
		'Wscript.Echo (UCase(Left(strComputerName,3)))
	End If
End If
End If

'File Explorer Shortcut
Set lnk = objShell.CreateShortcut(strUserProfile & "\Desktop\File Explorer.LNK")
If Not objFSO.FileExists(lnk) Then
   lnk.TargetPath = "C:\Windows\explorer.exe"
   lnk.Arguments = strUserProfile & "\Documents\"
   lnk.Description = "File Explorer"
   'lnk.HotKey = ""
   lnk.IconLocation = "C:\Windows\explorer.exe"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = strUserProfile & "\Documents\"
   lnk.Save
Set lnk = Nothing
End If

'Internet Explorer Shortcut
'Set lnk = objShell.CreateShortcut(strUserProfile & "\Desktop\Internet Explorer.LNK")
'If Not objFSO.FileExists(lnk) Then
'   lnk.TargetPath = "C:\Program Files\Internet Explorer\iexplore.exe"
'   lnk.Arguments = "http://some.url.local/path/path"
'   lnk.Description = "Time Clock"
   'lnk.HotKey = ""
'   lnk.IconLocation = "C:\Program Files\Internet Explorer\iexplore.exe"
'   lnk.WindowStyle = "1"
 '  lnk.WorkingDirectory = ""
'   lnk.Save
'Set lnk = Nothing
'End If

'IFS Offline TimeClock Shortcut
Set lnk = objShell.CreateShortcut(strUserProfile & "\Desktop\OfflineTimeClock.LNK")	'create shortcut
If Not objFSO.FileExists(lnk) Then

   If objFSO.FolderExists("C:\TimeClock_offline_files") Then 	'check if the offline folder is there before creating a shortcut
'-------------------------------------------------------------------------------------------------------
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")	'find computer's hostname

Set RegEx = CreateObject("vbscript.regexp") 	'extract number from hostname
RegEx.Pattern = "[^\d]"
RegEx.IgnoreCase = True 
RegEx.Global = True 
'numStr=RegEx.Replace(strComputerName, "") 
'msgbox numStr
numStrXXA = "1"
numStrXXB = "1"
numStrXXC = "1"
numStrXXD = "1"
numStrXXXE = "1"
numStrXXXF = "1"
numStrXXXGG = "1"
numStrXXH = "1"
numStrXXI = ""
numStrXXJ = "1"

If (UCase(Left(strComputerName,3))) = "XXA" Then
	strUserPass = (UCase(Left(strComputerName,3))) & numStrXXA 
	Else If (UCase(Left(strComputerName,3))) = "XXB" Then
		strUserPass = "XXB1"
		Else If (UCase(Left(strComputerName,3))) = "XXC" Then
			strUserPass = (UCase(Left(strComputerName,3))) & numStrXXC
			Else If (UCase(Left(strComputerName,3))) = "XXD" Then
				strUserPass = (UCase(Left(strComputerName,3))) & numStrXXD
				Else If (UCase(Left(strComputerName,3))) = "XXXE" Then
					strUserPass = (UCase(Left(strComputerName,3))) & numStrXXXE
					Else If (UCase(Left(strComputerName,3))) = "XXXF" Then
						strUserPass = (UCase(Left(strComputerName,3))) & numStrXXXF
						Else If (UCase(Left(strComputerName,3))) = "XXXGG" Then
							strUserPass = (UCase(Left(strComputerName,3))) & numStrXXXGG
							Else If (UCase(Left(strComputerName,3))) = "XXH" Then
								strUserPass = "XXH1"
								Else If (UCase(Left(strComputerName,3))) = "XXI" Then
									strUserPass = "XX"
									Else If (UCase(Left(strComputerName,3))) = "XXJ" Then
										strUserPass = (UCase(Left(strComputerName,3))) & numStrXXJ
									End If
								End If
							End If
						End If	
					End If	
				End If
			End If	
		End If
	End If
End If	

strUserID = "USER_ID=" & strUserPass		'concatenate user_id and first 3 characters of hostname and number
strPassID = "PASSWORD=" & strUserPass
'msgbox strUserID
'msgbox strPassID





strTargetPath1 = """CONNECTION_STRING=http://some.url.local/path/path"""	'setup target path strings
strTargetPath2 = chr(34) & strUserID & chr(34)	'enclose User_ID= string into double quotes
strTargetPath3 = chr(34) & strPassID & chr(34)
strTargetPath4 = """OFFLINE_FILEPATH=C:\TimeClock_offline_files"""
strTargetPath5 = """LANG_CODE=en-US"""
S = " "	'just a space

   lnk.TargetPath = "C:\Program Files (x86)\IFS Applications\Time Clock\TimeClockExe.exe"
   lnk.Arguments = S & strTargetPath1 & S & strTargetPath2 & S & strTargetPath3 & S & strTargetPath4 & S & strTargetPath5
   lnk.Description = "IFSTime Clock Test" 
  'lnk.HotKey = "F12"
   lnk.IconLocation = "C:\Program Files (x86)\IFS Applications\Time Clock\TimeClockExe.exe"
   lnk.WindowStyle = "1"
   lnk.WorkingDirectory = "C:\Program Files (x86)\IFS Applications\Time Clock"
   lnk.Save
Set lnk = Nothing
End If
End If

'5.8.1 Change Default Browser to IE

If objFSO.FileExists("C:\TimeClock\ChangeDefaultBrowser.bat") Then
		objShell.Run strDefaultBrowser
End If

'##############################DO YOU WANT WORKBENCH INSTALLED?############################################
If Not objFSO.FileExists(strWorkBenchRestart) Then

Const wshYes = 6
Const wshNo = 7
Const wshYesNoDialog = 4
Const wshQuestionMark = 32

intReturn = objShell.Popup("Do you want to install IFS Workbench?", _
    10, "IFS Workbench", wshYesNoDialog + wshQuestionMark)
If intReturn = wshYes Then
    objFSO.CopyFile "\\Server\share\folder\TimeClock\AddOn\IFSWorkBench\IFSWorkBenchRestart.vbs", "C:\TimeClock\"
	BtnCode= objShell.Popup ("IFS Workbench Installed", 3)
ElseIf intReturn = wshNo Then
    
End If
End If

'#######################################################################################################################################################################
'5.9. Check if TimeClock files have been copied, then execute IFS TimeClock and Caffeine Script to keep TClock Alive
If objFSO.FileExists("C:\TimeClock\DoNotDelete.txt") Then
		'objShell.Run strMOP
		'objShell.Run strPinIt
		'objShell.Run strAppRestart
		objShell.Run strMOPRestart
		If objFSO.FileExists("C:\Program Files (x86)\IFS Applications\Time Clock\TimeClockExe.exe") Then
		objShell.Run strOfflineAppRestart
		End If
		If objFSO.FileExists("C:\TimeClock\IFSWorkBenchRestart.vbs") Then
		objShell.Run strWorkBenchRestart
		End If
		objShell.Run strOSKRestart
		objShell.Run strCaffeine
		objShell.Run strSoftwareReporterFix
Else
If Not objFSO.FolderExists("C:\TimeClock") Then
	wscript.echo "Folder C:\TimeClock does not exist! Please Notify IT!"
End If
End If
'End if





