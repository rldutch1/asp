' ----------------------------------------------------------------------------
' Script Author: Robert Holland
' Script Name: CreateShortcut.vbs
' Creation Date: Fri Jan 30 2015 12:00:00 GMT-0700 (US Mountain Standard Time)
' Last Modified:
' Copyright (c)2015
' Purpose: Create a shortcut on the user desktop.
' Source: http://superuser.com/questions/392061/how-to-make-a-shortcut-from-cmd
' Source: http://www.techrepublic.com/forums/questions/retrieve-local-user-profile-path-with-vbscript/
' ----------------------------------------------------------------------------

Set oWS = WScript.CreateObject("WScript.Shell")
Dim sCurrentProfile
Set oWshEnvironment = oWS.Environment("Process")
sCurrentProfile = oWshEnvironment("USERPROFILE")
'userProfilePath = oWS.ExpandEnvironmentStrings("%UserProfile%")
sLinkFile = "wget.LNK"
Set oLink = oWS.CreateShortcut(sLinkFile)
 '   oLink.TargetPath = "%userprofile%\wget.exe"
 	oLink.TargetPath = sCurrentProfile & "\wget.exe"
 '  oLink.TargetPath = "C:\Program Files\MyApp\MyProgram.EXE"
 '  oLink.Arguments = ""
 '  oLink.Arguments = "-x -y -z /k"
 '  oLink.Description = "MyProgram"
 '  oLink.HotKey = "ALT+CTRL+F"
 '  oLink.IconLocation = "C:\Program Files\MyApp\MyProgram.EXE, 2"
 '  oLink.WindowStyle = "1"
 '  oLink.WorkingDirectory = "C:\Program Files\MyApp"
oLink.Save

' "C:\Users\%username%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\LABNCMCv_3.3.exe"
' FSO.CopyFile "C:\Backup.Log.File.txt", sCurrentProfile & "\Desktop\" & "Backup.Log.File." & strFileDate1

' Testing:
' sCurrentProfile & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\LABNCMCv_3.3.exe.lnk"