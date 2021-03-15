set Service = GetObject ("winmgmts:")
set WshShell = WScript.CreateObject("WScript.Shell")
strUserProfile = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")

'Name of the exe we want to watch
sEXEName = "TimeClockExe.exe"
'Path
sApplicationPath = strUserProfile & "\Desktop\OfflineTimeClock.LNK"

'Loop until the system is shutdown or user logs out
while true 
 bRunning = false
 for each Process in Service.InstancesOf ("Win32_Process")
  if Process.Name = sEXEName then
   bRunning=true
  End If
 next

'Is our app running?

if (not bRunning) then
 'No it is not, launch it
 WshShell.Run Chr(34) & sApplicationPath & Chr(34)
end if

'Sleep a while so we do not hog the cpu
WScript.Sleep(30000)

wend