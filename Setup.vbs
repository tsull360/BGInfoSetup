Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCWD = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(Wscript.ScriptName)))
strAllUsers = WshShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%")

Const OverwriteExisting = True

'Gather machine details.
'=====================================
Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
For Each System in SystemSet 
	strOSVer = System.Version
	'WScript.Echo strOSVer
	strOSArch = System.OSArchitecture
	'WScript.Echo strOSArch
	intOSType = System.ProductType
	'WScript.Echo intOSType
Next

'Determining Program Files location
'======================================
If strOSArch = "64-bit" Then
'Copy to Program Files (x86)
strCopyDest = "C:\Program Files (x86)\BGInfo"
If Not objFSO.FolderExists(strCopyDest) then
	objFSO.CreateFolder strCopyDest
	objFSO.CopyFile strCWD & "\*.vbs", strCopyDest, OverwriteExisting
else
	objFSO.CopyFile strCWD & "\*.vbs", strCopyDest, OverwriteExisting
End If

strCopyDestWin = "C:\Windows\SysWow64"

'WScript.Echo "Copying to program files x86"
'WScript.Echo strCopyDest
Else
'Copy to Program Files
strCopyDest = "C:\Program Files\BGInfo"
strCopyDestWin = "C:\Windows\System32"
objFSO.CopyFile strCWD & "\*.vbs", strCopyDest, OverwriteExisting
'WScript.Echo strCopyDest
'WScript.Echo "Copying to Program Files"
End If

'Copy files
'=========================================
if objFSO.FolderExists(strCopyDest) Then
Else
objFSO.CreateFolder(strCopyDest)
End If
if objFSO.FileExists(strCopyDest & "\BGInfo.exe") Then
objFSO.DeleteFile strCopyDest & "\BGInfo.exe"
objFSO.CopyFile strCWD & "\BGInfo.exe", strCopyDest & "\", OverwriteExisting
Else
objFSO.CopyFile strCWD & "\BGInfo.exe", strCopyDest & "\", OverwriteExisting
End If

'Determining Windows version
'=========================================
If InStr(strOSVer,"10.0") = 1 Then
'Windows 10/2016
'WScript.Echo "Found a 10 version"
If objFSO.FileExists(strCopyDest & "\default.bgi") Then
objFSO.DeleteFile strCopyDest & "\default.bgi"
objFSO.CopyFile strCWD & "\OS2k16.bgi", strCopyDest & "\default.bgi", OverwriteExisting
Else
objFSO.CopyFile strCWD & "\OS2k16.bgi", strCopyDest & "\default.bgi", OverwriteExisting
End If

ElseIf InStr(strOSVer,"6.3") = 1 Then
'Windows 8.1/2012 R2
'WScript.Echo "Found a 6.3 version"
If objFSO.FileExists(strCopyDest & "\default.bgi") Then
objFSO.DeleteFile strCopyDest & "\default.bgi"
objFSO.CopyFile strCWD & "\OS2k12.bgi", strCopyDest & "\default.bgi", OverwriteExisting
Else
objFSO.CopyFile strCWD & "\OS2k12.bgi", strCopyDest & "\default.bgi", OverwriteExisting
End If

ElseIf InStr(strOSVer,"6.2") = 1 Then
'Windows 8.0/2012
'WScript.Echo "Found a 6.2 version"
If objFSO.FileExists(strCopyDest & "\default.bgi") Then
objFSO.DeleteFile strCopyDest & "\default.bgi"
objFSO.CopyFile strCWD & "\OS2k12.bgi", strCopyDest & "\default.bgi"
Else
objFSO.CopyFile strCWD & "\OS2k12.bgi", strCopyDest & "\default.bgi"
End If

ElseIf InStr(strOSVer,"6.1") = 1 Then
'Windows 7/2008 R2
'WScript.Echo "Found a 6.1 version"
If objFSO.FileExists(strCopyDest & "\default.bgi") Then
objFSO.DeleteFile strCopyDest & "\default.bgi"
objFSO.CopyFile strCWD & "\OS2k8R2.bgi", strCopyDest & "\default.bgi"
Else
objFSO.CopyFile strCWD & "\OS2k8R2.bgi", strCopyDest & "\default.bgi"
End If

ElseIf InStr(strOSVer,"6.0") = 1 Then
'Windows Vista/2008
'WScript.Echo "Found a 6.0 version"
If objFSO.FileExists(strCopyDest & "\default.bgi") Then
objFSO.DeleteFile strCopyDest & "\default.bgi"
objFSO.CopyFile strCWD & "\OS2k8.bgi", strCopyDest & "\default.bgi"
Else
objFSO.CopyFile strCWD & "\OS2k8.bgi", strCopyDest & "\default.bgi"
End If

ElseIf InStr(strOSVer,"5.2") = 1 Then
'Windows XP/2003
'WScript.Echo "Found a 5.2 version"
If objFSO.FileExists(strCopyDest & "\default.bgi") Then
objFSO.DeleteFile strCopyDest & "\default.bgi"
objFSO.CopyFile strCWD & "\OS2k3.bgi", strCopyDest & "\default.bgi"
Else
objFSO.CopyFile strCWD & "\OS2k3.bgi", strCopyDest & "\default.bgi"
End If

ElseIf InStr(strOSVer,"5.0") = 1 Then
'Windows 2000
'WScript.Echo "Found a 5.0 version"

Else
'Couldn't determine OS Version
'WScript.Echo "Unable to determine OS version"

End If

'Create startup shortcut
'=========================================
set objShellLink = WshShell.CreateShortcut(strAllUsers & "\Microsoft\Windows\Start Menu\Programs\StartUp\BGInfo.lnk")
'Define where you would like the shortcut to point to. This can be a program, Web site, etc.
objShellLink.TargetPath = strCopyDest & "\Bginfo.exe"
objShellLink.WindowStyle = 1
'Define what icon the shortcut should use. Here we use the explorer.exe icon.
objShellLink.IconLocation = strCopyDest & "\BGInfo.exe"
objShellLink.Description = "BGInfo"
objShellLink.WorkingDirectory = allusersdesktop
objShellLink.Arguments = chr(34) & strCopyDest & "\default.bgi" & chr(34) & " /NOLICPROMPT /TIMER:0 /SILENT"
objShellLink.Save
Wscript.Sleep 5
WshShell.Run(chr(34) & strAllUsers & "\Microsoft\Windows\Start Menu\Programs\StartUp\BGInfo.lnk" & chr(34))