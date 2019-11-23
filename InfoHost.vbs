Const HKEY_LOCAL_MACHINE = &H80000002
 
strComputer = "."
strKeyPath = "Software\Microsoft\Virtual Machine\Guest\Parameters"
 
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
oReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
oReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "HostName", strHostName

If Len(strHostName) = 0 Then
strHostName = "Physical"
Else
strHostName = strHostName
End If

On Error Resume Next
Call WScript.Echo(strHostName)
Call Echo(strHostName)
On Error Goto 0

