' ----------------------------------------------------------------------------
' Script Author: Robert Holland
' Script Name: CreateShortcut.vbs
' Creation Date: Fri Jan 30 2015 12:00:00 GMT-0700 (US Mountain Standard Time)
' Last Modified:
' Copyright (c)2015
' Purpose: Create a shortcut in the user startup folder.
' Source: http://superuser.com/questions/392061/how-to-make-a-shortcut-from-cmd
' Source: http://www.techrepublic.com/forums/questions/retrieve-local-user-profile-path-with-vbscript/
' ----------------------------------------------------------------------------

Set oWS = WScript.CreateObject("WScript.Shell")
Dim uCurrentProfile, sFilename
Set oWshEnvironment = oWS.Environment("Process")
uCurrentProfile = oWshEnvironment("USERPROFILE")
uCurrentAppProfile = oWshEnvironment("APPDATA")
'userProfilePath = oWS.ExpandEnvironmentStrings("%UserProfile%")
sFilename = "filename.exe"
sLinkFile = uCurrentAppProfile & "\Microsoft\Windows\Start Menu\Programs\Startup\" & sFilename & ".LNK"
Set oLink = oWS.CreateShortcut(sLinkFile)
 	oLink.TargetPath = uCurrentProfile & "\" & sFilename
 '  oLink.Arguments = "-x -y -z /k"
 '  oLink.Description = "MyProgram"
 '  oLink.HotKey = "ALT+CTRL+F"
 '  oLink.IconLocation = "C:\Program Files\MyApp\MyProgram.EXE, 2"
 '  oLink.WindowStyle = "1"
 		oLink.WorkingDirectory = uCurrentProfile
oLink.Save
